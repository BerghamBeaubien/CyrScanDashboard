using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;
using Dapper;
using YourProject.Models;

namespace YourProject.Controllers;

[ApiController]
[Route("api/[controller]")]
public class DashboardController : ControllerBase
{
    private readonly string _connectionString = "Server=192.168.88.55,1433;Database=CyrScanDB;User Id=Serveur-CyrScan;Password=admin;";

    [HttpGet("jobs")]
    public async Task<IActionResult> GetJobs()
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();
            var query = @"
                SELECT 
                    JobNumber,
                    COUNT(DISTINCT PartID) as TotalParts,
                    SUM(Quantity) as TotalQuantity,
                    SUM(ScannedQuantity) as TotalScanned,
                    MAX(ScanDate) as LastScanDate
                FROM ScannedTags
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
                SELECT *
                FROM ScannedTags
                WHERE JobNumber = @JobNumber
                ORDER BY ScanDate DESC";
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
                    COUNT(DISTINCT JobNumber) as TotalJobs,
                    SUM(ScannedQuantity) as TotalScannedItems,
                    COUNT(DISTINCT PartID) as TotalUniqueParts,
                    (SELECT TOP 1 JobNumber 
                     FROM ScannedTags 
                     ORDER BY ScanDate DESC) as LatestJob
                FROM ScannedTags";
            var stats = await connection.QuerySingleAsync(query);
            return Ok(stats);
        }
    }
}