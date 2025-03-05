using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;
using Dapper;
using CyrScanDashboard.Models;
using System.Data;
using CyrScanDashboard.Services;
using Microsoft.IdentityModel.Tokens;

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
                COUNT(QRCode) AS TotalScanned,
                COUNT(DISTINCT PalletId) AS TotalPallets,
                MAX(TotalQuantityJob) AS TotalExpected,
                MAX(ScanDate) AS LastScanDate
            FROM ScannedTags
            GROUP BY JobNumber
            ORDER BY MAX(ScanDate) DESC";

            var jobs = await connection.QueryAsync<JobSummary>(query);
            return Ok(jobs);
        }
    }

    [HttpGet("jobs/{jobNumber}")]
    public async Task<IActionResult> GetJobDetails(string jobNumber)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();

            // R�cup�rer les informations des pi�ces scan�es
            var query = @"
            SELECT 
                s.PartID,
                COUNT(s.QRCode) AS ScannedCount,
                MAX(s.ScanDate) AS LastScanDate,
                STUFF((
                    SELECT DISTINCT ', ' + p.Name
                    FROM Pallets p
                    INNER JOIN ScannedTags st ON p.Id = st.PalletId
                    WHERE st.PartID = s.PartID AND st.JobNumber = @JobNumber
                    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS Pallets
            FROM ScannedTags s
            WHERE s.JobNumber = @JobNumber
            GROUP BY s.PartID
            ORDER BY MAX(s.ScanDate) DESC";

            var details = await connection.QueryAsync(query, new { JobNumber = jobNumber });

            // Utiliser le service pour r�cup�rer les informations Excel
            var excelParts = _excelValidationService.GetExcelJobDetails(jobNumber);
            var totalParts = excelParts.Count();

            // Combiner les r�sultats
            var result = new
            {
                DatabaseParts = details,
                TotalParts = totalParts,
                ExcelParts = excelParts,
                JobNumber = jobNumber
            };

            return Ok(result);
        }
    }

    [HttpGet("parts/{jobNumber}/{partId}")]
    public async Task<IActionResult> GetPartDetails(string jobNumber, string partId)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();
            var query = @"
                SELECT 
                    st.QRCode,
                    p.Name AS PalletName,
                    st.ScanDate
                FROM ScannedTags st
                JOIN Pallets p ON st.PalletId = p.Id
                WHERE st.JobNumber = @JobNumber AND st.PartID = @PartID
                ORDER BY st.ScanDate DESC";

            var details = await connection.QueryAsync(query, new { JobNumber = jobNumber, PartID = partId });
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
                COUNT(*) AS TotalScannedItems,
                (SELECT TOP 1 JobNumber 
                 FROM ScannedTags 
                 ORDER BY ScanDate DESC) AS LatestJob,
                (SELECT COUNT(DISTINCT Id) FROM Pallets) AS TotalPallets
            FROM ScannedTags";

            var stats = await connection.QuerySingleAsync(query);
            return Ok(stats);
        }
    }

    [HttpGet("pallets/{jobNumber}")]
    public async Task<IActionResult> GetPalletsByJob(string jobNumber)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();

            // Get all pallets for this job
            var palletsQuery = @"
            SELECT Id, JobNumber, Name, CreatedDate, SequenceNumber
            FROM Pallets
            WHERE JobNumber = @JobNumber
            ORDER BY SequenceNumber";

            var pallets = await connection.QueryAsync<Pallet>(palletsQuery, new { JobNumber = jobNumber });

            // For each pallet, count the scanned items
            var result = new List<Pallet>();
            foreach (var pallet in pallets)
            {
                var countQuery = @"
                SELECT COUNT(*) 
                FROM ScannedTags
                WHERE PalletId = @PalletId";

                var scannedCount = await connection.ExecuteScalarAsync<int>(countQuery, new { PalletId = pallet.Id });

                pallet.ScannedItems = scannedCount;
                result.Add(pallet);
            }

            return Ok(result);
        }
    }

    [HttpPost("pallets")]
    public async Task<IActionResult> CreatePallet([FromBody] CreatePalletRequest request)
    {
        try
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();

                // Check if JobNumber is valid
                if (string.IsNullOrEmpty(request.JobNumber))
                {
                    return BadRequest(new { message = "JobNumber invalide." });
                }

                // Get the next sequence number for this job
                var sequenceQuery = @"
                SELECT ISNULL(MAX(SequenceNumber), 0) + 1 
                FROM Pallets
                WHERE JobNumber = @JobNumber";

                var nextSequence = await connection.ExecuteScalarAsync<int>(sequenceQuery,
                    new { JobNumber = request.JobNumber });

                // Automatically create a name (PAL + sequence)
                string palletName = $"PAL{nextSequence}";

                // Execute the stored procedure
                var parameters = new DynamicParameters();
                parameters.Add("@JobNumber", request.JobNumber);
                parameters.Add("@Name", palletName);

                var newPallet = await connection.QuerySingleAsync<Pallet>(
                    "CreateNewPallet",
                    parameters,
                    commandType: CommandType.StoredProcedure);

                // New pallet has no items yet
                newPallet.ScannedItems = 0;

                return Ok(newPallet);
            }
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { message = $"Erreur: {ex.Message}" });
        }
    }

    [HttpPut("pallets/{id}")]
    public async Task<IActionResult> UpdatePalletName(int id, [FromBody] UpdatePalletRequest request)
    {
        try
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();

                var updateQuery = @"
                UPDATE Pallets
                SET Name = @Name
                WHERE Id = @Id";

                var rowsAffected = await connection.ExecuteAsync(updateQuery, new { Id = id, Name = request.Name });

                if (rowsAffected == 0)
                    return NotFound("Palette non trouv�e");

                return Ok(new { message = "Palette mise � jour avec succ�s" });
            }
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { message = $"Erreur: {ex.Message}" });
        }
    }

    [HttpDelete("pallets/{id}")]
    public async Task<IActionResult> DeletePallet(int id)
    {
        try
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();

                var deleteQuery = @"
                DELETE FROM Pallets
                WHERE Id = @Id";

                var rowsAffected = await connection.ExecuteAsync(deleteQuery, new { Id = id });

                if (rowsAffected == 0)
                    return NotFound("Palette non trouv�e");

                return Ok(new { message = "Palette supprim�e avec succ�s" });
            }
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { message = $"Erreur: {ex.Message}" });
        }
    }

    [HttpPost("scan")]
    public async Task<IActionResult> AddScan([FromBody] ScanRequest request)
    {
        // Validate the pallet ID
        if (request.PalletId <= 0)
        {
            return BadRequest(new
            {
                success = false,
                message = "Une palette doit �tre s�lectionn�e pour scanner des pi�ces"
            });
        }

        request.QRCode = FormatQRCode(request.QRCode);

        // Validate the part ID against Excel data
        var validationResult = _excelValidationService.ValidatePart(
            request.JobNumber,
            request.PartID,
            request.QRCode
        );

        if (!validationResult.isValid)
        {
            return BadRequest(new
            {
                success = false,
                message = validationResult.message
            });
        }

        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();

            // Check if QR code already scanned
            var checkQuery = "SELECT COUNT(*) FROM ScannedTags WHERE QRCode = @QRCode";
            var exists = await connection.ExecuteScalarAsync<int>(checkQuery, new { request.QRCode });

            if (exists > 0)
            {
                return BadRequest(new
                {
                    success = false,
                    message = "Ce QR code a d�j� �t� scann�"
                });
            }

            // Get TotalQuantityJob from Excel
            int totalQuantityJob = validationResult.totalQty;

            // Insert new scan record with TotalQuantityJob
            var insertQuery = @"
            INSERT INTO ScannedTags (
                JobNumber, PartID, QRCode, 
                ScanDate, PalletId, TotalQuantityJob
            ) VALUES (
                @JobNumber, @PartID, @QRCode, 
                @ScanDate, @PalletId, @TotalQuantityJob
            )";

            await connection.ExecuteAsync(insertQuery, new
            {
                request.JobNumber,
                request.PartID,
                request.QRCode,
                ScanDate = DateTime.Now,
                request.PalletId,
                TotalQuantityJob = totalQuantityJob
            });

            return Ok(new
            {
                success = true,
                message = "Scan saved successfully"
            });
        }
    }

    [HttpDelete("delete")]
    public async Task<IActionResult> DeleteScan([FromBody] DeleteScanRequest request)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();

            var deleteQuery = @"
            DELETE FROM ScannedTags
            WHERE JobNumber = @JobNumber 
            AND PartID = @PartID
            AND QRCode = @QRCode
            AND PalletId = @PalletId";

            var rowsAffected = await connection.ExecuteAsync(deleteQuery, new
            {
                request.JobNumber,
                request.PartID,
                request.QRCode,
                request.PalletId
            });

            if (rowsAffected == 0)
            {
                return NotFound(new { message = "Scan non trouv�" });
            }

            return Ok(new { message = "Scan supprim� avec succ�s" });
        }
    }

    [HttpGet("unscanned-pallets")]
    public async Task<IActionResult> GetUnscannedPallets()
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();
            var query = @"
        SELECT p.*
        FROM Pallets p
        LEFT JOIN ScannedTags s ON p.Id = s.PalletId
        WHERE s.PalletId IS NULL";

            var pallets = await connection.QueryAsync<Pallet>(query);
            return Ok(pallets);
        }
    }

    [HttpGet("jobs/{jobNumber}/complete")]
    public async Task<IActionResult> GetCompleteJobDetails(string jobNumber)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();

            // 1. Get basic job data from Excel
            var excelParts = _excelValidationService.GetExcelJobDetails(jobNumber);
            var totalParts = excelParts.Count();

            // 2. Get all scan details for all parts in one query
            var scanDetailsQuery = @"
            SELECT 
                st.PartID,
                st.QRCode,
                p.Name AS PalletName,
                st.ScanDate
            FROM ScannedTags st
            JOIN Pallets p ON st.PalletId = p.Id
            WHERE st.JobNumber = @JobNumber
            ORDER BY st.ScanDate DESC";

            var allScans = await connection.QueryAsync(scanDetailsQuery, new { JobNumber = jobNumber });

            // 3. Group scans by PartID for easier processing on frontend
            var scansByPart = allScans.GroupBy(s => s.PartID)
                .ToDictionary(
                    g => g.Key,
                    g => g.Select(s => new {
                        s.QRCode,
                        s.PalletName,
                        s.ScanDate
                    }).ToList()
                );

            // 4. Create summary data (similar to original job details endpoint)
            var partSummaries = scansByPart.Select(kvp => {
                var partId = kvp.Key;
                var scans = kvp.Value;

                // Count unique sequences for actual scanned count
                var uniqueSequences = new HashSet<string>();
                foreach (var scan in scans)
                {
                    var qrParts = scan.QRCode?.ToString().Split('-');
                    var sequence = qrParts != null && qrParts.Length > 0 ? qrParts[qrParts.Length - 1] : "";
                    if (!string.IsNullOrEmpty(sequence))
                    {
                        uniqueSequences.Add(sequence);
                    }
                }

                return new
                {
                    PartID = partId,
                    ScannedCount = uniqueSequences.Count,
                    LastScanDate = scans.Count > 0 ? scans.Max(s => s.ScanDate) : null,
                    Pallets = string.Join(", ", scans.Select(s => s.PalletName).Distinct())
                };
            }).ToList();

            // 5. Combine everything into a single response
            var result = new
            {
                DatabaseParts = partSummaries,
                ScanDetails = scansByPart,
                TotalParts = totalParts,
                ExcelParts = excelParts,
                JobNumber = jobNumber
            };

            return Ok(result);
        }
    }

    [HttpGet("palletContents/{palletId}")]
    public async Task<IActionResult> GetPalletContents(int palletId)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();

            // Get all scanned items for this pallet with only the requested fields
            var contentsQuery = @"
                SELECT 
                    PartId as partId,
                    QrCode as qrCode,
                    ScanDate as scanDate
                FROM ScannedTags
                WHERE PalletId = @PalletId
                ORDER BY ScanDate DESC";

            var palletContents = await connection.QueryAsync(contentsQuery, new { PalletId = palletId });

            return Ok(palletContents);
        }
    }

    [HttpGet("packaging/{palletId:int}")]
    public async Task<IActionResult> CreatePackaging(int palletId, string palLong, string palLarg)
    {
        try
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();

                // Get job number and pallet name
                var jobPalletInfo = await connection.QueryFirstOrDefaultAsync<JobPalletInfo>(
                    @"SELECT DISTINCT
                    st.JobNumber,
                    p.Name AS PalletName,
                    p.SequenceNumber,
                    p.Id
                FROM 
                    dbo.ScannedTags st
                JOIN 
                    dbo.Pallets p ON st.PalletId = p.Id
                WHERE 
                    st.PalletId = @PalletId",
                    new { PalletId = palletId });

                if (jobPalletInfo == null)
                {
                    return NotFound(new { message = "Aucune information trouv�e pour cette palette" });
                }

                // Get parts and their quantities
                var partsQuery = @"
            SELECT 
                PartId,
                COUNT(*) AS Quantity
            FROM ScannedTags
            WHERE PalletId = @PalletId
            GROUP BY PartId
            ORDER BY Quantity DESC";

                var partQuantities = await connection.QueryAsync<PartQuantity>(partsQuery, new { PalletId = palletId });

                // Get unique part names
                var partNames = partQuantities.Select(p => p.PartId).ToList();

                // Get part details from Excel
                var partDetails = _excelValidationService.GetPartPackagingDetails(jobPalletInfo.JobNumber, partNames);

                // Ensure order and quantities match
                var orderedPartDetails = partNames
                    .Select(partName => partDetails.FirstOrDefault(p => p.PartName == partName))
                    .Where(p => p != null)
                    .ToList();

                var quantities = partQuantities.Select(p => p.Quantity).ToArray();

                // Create packaging file
                string filePath = _excelValidationService.CreateEmballageFile(
                    jobPalletInfo.JobNumber,
                    jobPalletInfo.PalletName,
                    palLong,
                    palLarg,
                    orderedPartDetails,
                    quantities
                );

                return Ok(new
                {
                    message = "Fichier d'emballage cr�� avec succ�s",
                    filePath = filePath,
                    jobNumber = jobPalletInfo.JobNumber,
                    paletteName = jobPalletInfo.PalletName
                });
            }
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { message = $"Erreur: {ex.Message}" });
        }
    }

    //Methode pour Ajouter un 0 avant le numero de sequence si le numero est inferieur a 10
    private string FormatQRCode(string qrCode)
    {
        if (string.IsNullOrWhiteSpace(qrCode)) return qrCode;

        var parts = qrCode.Split('-');
        if (parts.Length > 1 && int.TryParse(parts[^1], out int lastNumber) && lastNumber is >= 1 and <= 9)
        {
            parts[^1] = $"0{lastNumber}"; // Prepend 0 if it's a single digit
        }

        return string.Join("-", parts);
    }

    private async Task<IEnumerable<(string PartId, int Quantity)>> GetPalletPartsGrouped(int palletId)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            var query = @"
            SELECT 
                PartId,
                COUNT(*) AS Quantity
            FROM ScannedTags
            WHERE PalletId = @PalletId
            GROUP BY PartId
            ORDER BY Quantity DESC";

            return await connection.QueryAsync<(string, int)>(
                query,
                new { PalletId = palletId }
            );
        }
    }
}