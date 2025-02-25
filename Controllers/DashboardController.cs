using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;
using Dapper;
using CyrScanDashboard.Models;
using ExcelDataReader;
using System.Data;
using System.IO;
using System.Collections.Concurrent;
using System.Net;
using System.ComponentModel;
using System.Runtime.InteropServices;
using CyrScanDashboard.Services;

namespace CyrScanDashboard.Controllers;

[ApiController]
[Route("api/[controller]")]
public class DashboardController : ControllerBase
{
    private readonly string _connectionString = "Server=192.168.88.55,1433;Database=CyrScanDB;User Id=Serveur-CyrScan;Password=admin;";
    private readonly ExcelValidationService _excelValidationService;

    public DashboardController(ExcelValidationService excelValidationService)
    {
        _excelValidationService = excelValidationService;
    }
    [HttpGet("jobs")]
    public async Task<IActionResult> GetJobs()
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();
            var query = @"
                SELECT 
                    JobNumber,
                    COUNT(DISTINCT PartID) AS TotalParts,
                    SUM(Quantity) AS TotalQuantity, -- Quantité totale des tags scannés
                    SUM(ScannedQuantity) AS TotalScanned, -- Quantité effectivement scannée
                    (SELECT SUM(TotalQuantityJob) FROM ScannedTags s WHERE s.JobNumber = t.JobNumber) AS TotalQuantityJob, -- Nouvelle colonne
                    MAX(ScanDate) AS LastScanDate
                FROM ScannedTags t
                GROUP BY JobNumber
                ORDER BY MAX(ScanDate) DESC";
            var jobs = await connection.QueryAsync(query);
            return Ok(jobs);
        }
    }

    [HttpGet("jobs/{jobNumber}")]
    public async Task<IActionResult> GetJobDetails(int jobNumber)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();
            var query = @"
                SELECT 
                    s.PartID,
                    s.Quantity,
                    s.ScannedQuantity,
                    s.ScanDate,
                    q.TotalQuantityJob
                FROM ScannedTags s
                INNER JOIN (
                    SELECT JobNumber, SUM(TotalQuantityJob) AS TotalQuantityJob
                    FROM ScannedTags
                    WHERE JobNumber = @JobNumber
                    GROUP BY JobNumber
                ) q ON s.JobNumber = q.JobNumber
                WHERE s.JobNumber = @JobNumber
                ORDER BY s.ScanDate DESC";

            var details = await connection.QueryAsync<ScanRecord>(query, new { JobNumber = jobNumber });
            return Ok(details);
        }
    }

    [HttpGet("stats")]
    public async Task<IActionResult> GetDashboardStats()
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();
            var query = @"
                SELECT 
                    COUNT(DISTINCT JobNumber) AS TotalJobs,
                    COUNT(DISTINCT PartID) AS TotalUniqueParts,
                    SUM(TotalQuantityJob) AS TotalQuantityJob,
                    SUM(ScannedQuantity) AS TotalScannedItems,
                    (SELECT TOP 1 JobNumber 
                     FROM ScannedTags 
                     ORDER BY ScanDate DESC) AS LatestJob
                FROM (
                    SELECT JobNumber, PartID, SUM(Quantity) AS TotalQuantityJob, SUM(ScannedQuantity) AS ScannedQuantity
                    FROM ScannedTags
                    GROUP BY JobNumber, PartID
                ) AS JobQuantities";
            var stats = await connection.QuerySingleAsync(query);
            return Ok(stats);
        }
    }

    [HttpDelete("delete")]
    public async Task<IActionResult> DeleteScan([FromBody] ScanRequest request)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();

            var checkQuery = @"
        SELECT ScannedQuantity 
        FROM ScannedTags 
        WHERE JobNumber = @JobNumber AND PartID = @PartID";

            var existingQuantity = await connection.QuerySingleOrDefaultAsync<int?>(
                checkQuery,
                new { request.JobNumber, request.PartID }
            );

            if (!existingQuantity.HasValue)
            {
                return NotFound(new { message = "Tag non trouvé" });
            }

            if (existingQuantity.Value > 1)
            {
                // Reduce quantity by 1
                var updateQuery = @"
            UPDATE ScannedTags 
            SET ScannedQuantity = ScannedQuantity - 1, 
                ScanDate = @ScanDate 
            WHERE JobNumber = @JobNumber 
            AND PartID = @PartID";

                await connection.ExecuteAsync(updateQuery, new
                {
                    ScanDate = DateTime.Now,
                    request.JobNumber,
                    request.PartID
                });

                return Ok(new { message = "Quantité réduite de 1" });
            }
            else
            {
                // Delete the record if quantity is 1
                var deleteQuery = @"
            DELETE FROM ScannedTags
            WHERE JobNumber = @JobNumber AND PartID = @PartID";

                await connection.ExecuteAsync(deleteQuery, new { request.JobNumber, request.PartID });

                return Ok(new { message = "Tag supprimé avec succès" });
            }
        }
    }

    [HttpPost("scan")]
    public async Task<IActionResult> AddScan([FromBody] ScanRequest request)
    {
        // Validate the part against Excel data
        var validationResult = _excelValidationService.ValidatePart(
            request.JobNumber,
            request.PartID,
            request.Quantity
        );

        if (!validationResult.isValid)
        {
            // Return the validation error
            return BadRequest(new
            {
                success = false,
                message = validationResult.message,
                expectedQuantity = validationResult.expectedQuantity
            });
        }

        // If valid, proceed with database operations
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();

            // Check if entry exists
            var checkQuery = @"
        SELECT ScannedQuantity 
        FROM ScannedTags 
        WHERE JobNumber = @JobNumber AND PartID = @PartID";
            var existingQuantity = await connection.QuerySingleOrDefaultAsync<int?>(
                checkQuery,
                new { request.JobNumber, request.PartID }
            );

            if (existingQuantity.HasValue)
            {
                // Update existing entry
                var updateQuery = @"
            UPDATE ScannedTags 
            SET ScannedQuantity = @NewQuantity, 
                ScanDate = @ScanDate 
            WHERE JobNumber = @JobNumber 
            AND PartID = @PartID";
                await connection.ExecuteAsync(updateQuery, new
                {
                    NewQuantity = existingQuantity.Value + 1,
                    ScanDate = DateTime.Now,
                    request.JobNumber,
                    request.PartID
                });
            }
            else
            {
                // Insert new entry
                var insertQuery = @"
            INSERT INTO ScannedTags (
                JobNumber, PartID, Quantity, 
                ScannedQuantity, ScanDate, TotalQuantityJob
            ) VALUES (
                @JobNumber, @PartID, @Quantity, 
                1, @ScanDate, @TotalQuantityJob
            )";
                await connection.ExecuteAsync(insertQuery, new
                {
                    request.JobNumber,
                    request.PartID,
                    request.Quantity,
                    TotalQuantityJob = validationResult.totalQuantityJob,
                    ScanDate = DateTime.Now
                });
            }

            return Ok(new
            {
                success = true,
                message = "Scan saved successfully"
            });
        }
    }
}