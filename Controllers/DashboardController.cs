using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Dapper;
using CyrScanDashboard.Models;
using System.Data;
using CyrScanDashboard.Services;

namespace CyrScanDashboard.Controllers;

[ApiController]
[Route("api/[controller]")]
public class DashboardController : ControllerBase
{
    private readonly string _connectionString = "Server=cyrscan-server,1433;Database=CyrScanDB;User Id=Serveur-CyrScan;Password=admin;TrustServerCertificate=True;";
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

            // Récupérer les informations des pièces scanées
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

            // Utiliser le service pour récupérer les informations Excel
            var excelParts = _excelValidationService.GetExcelJobDetails(jobNumber);
            var totalParts = excelParts.Count();

            // Combiner les résultats
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
                    return NotFound("Palette non trouvée");

                return Ok(new { message = "Palette mise à jour avec succès" });
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
                    return NotFound("Palette non trouvée");

                return Ok(new { message = "Palette supprimée avec succès" });
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
        // Vérifier si l'utilisateur est authentifié
        int userId = int.Parse(User.FindFirst("UserId")?.Value ?? "0");
        if (userId == 0)
        {
            return Unauthorized(new
            {
                success = false,
                message = "L'utilisateur doit être authentifié pour scanner des pièces"
            });
        }

        // Valider l'ID de la palette
        if (request.PalletId <= 0)
        {
            return BadRequest(new
            {
                success = false,
                message = "Une palette doit être sélectionnée pour scanner des pièces"
            });
        }

        request.QRCode = FormatQRCode(request.QRCode);

        // Valider le numéro de pièce par rapport aux données Excel
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

            var palletJobQuery = "SELECT JobNumber FROM Pallets WHERE Id = @PalletId";
            var palletJobNumber = await connection.QueryFirstOrDefaultAsync<string>(palletJobQuery, new { request.PalletId });

            if (palletJobNumber == null)
            {
                return BadRequest(new
                {
                    success = false,
                    message = "Palette introuvable"
                });
            }

            if (palletJobNumber != request.JobNumber)
            {
                return BadRequest(new
                {
                    success = false,
                    message = "Le numéro de job ne correspond pas à celui de la palette sélectionnée"
                });
            }

            // Vérifier si le code QR a déjà été scanné
            var checkQuery = "SELECT COUNT(*) FROM ScannedTags WHERE QRCode = @QRCode";
            var exists = await connection.ExecuteScalarAsync<int>(checkQuery, new { request.QRCode });

            if (exists > 0)
            {
                return BadRequest(new
                {
                    success = false,
                    message = "Ce Code QR a déjà été scanné"
                });
            }

            // Obtenir TotalQuantityJob depuis Excel
            int totalQuantityJob = validationResult.totalQty;

            // Insérer un nouvel enregistrement de scan avec TotalQuantityJob et l'ID de l'utilisateur
            var insertQuery = @"
        INSERT INTO ScannedTags (
            JobNumber, PartID, QRCode, 
            ScanDate, PalletId, TotalQuantityJob, ScannedByUserId
        ) VALUES (
            @JobNumber, @PartID, @QRCode, 
            @ScanDate, @PalletId, @TotalQuantityJob, @UserId
        )";

            await connection.ExecuteAsync(insertQuery, new
            {
                request.JobNumber,
                request.PartID,
                request.QRCode,
                ScanDate = DateTime.Now,
                request.PalletId,
                TotalQuantityJob = totalQuantityJob,
                UserId = userId
            });

            return Ok(new
            {
                success = true,
                message = "Scan Réussi"
            });
        }
    }

    [HttpDelete("delete")]
    public async Task<IActionResult> DeleteScan([FromBody] DeleteScanRequest request)
    {
        // Vérifier si l'utilisateur est authentifié
        int userId = int.Parse(User.FindFirst("UserId")?.Value ?? "0");
        if (userId == 0)
        {
            return Unauthorized(new
            {
                success = false,
                message = "L'utilisateur doit être authentifié pour supprimer des scans"
            });
        }

        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();

            // Commencer une transaction
            using (var transaction = connection.BeginTransaction())
            {
                try
                {
                    // Récupérer les détails du scan avant de le supprimer
                    var scanQuery = @"
                SELECT s.*, p.Name as PalletName
                FROM ScannedTags s
                LEFT JOIN Pallets p ON s.PalletId = p.Id
                WHERE s.QRCode = @QRCode AND s.PalletId = @PalletId";

                    var scan = await connection.QueryFirstOrDefaultAsync(scanQuery, new
                    {
                        request.QRCode,
                        request.PalletId
                    }, transaction);

                    if (scan == null)
                    {
                        return NotFound(new { message = "Scan non trouvé" });
                    }

                    // Insérer dans la table DeletedScans
                    var insertDeletedQuery = @"
                INSERT INTO DeletedScans (
                    JobNumber, PartID, QRCode, ScanDate, 
                    DeletedDate, PalletId, TotalQuantityJob, 
                    DeletedByUserId, PalletName
                ) VALUES (
                    @JobNumber, @PartID, @QRCode, @ScanDate,
                    GETDATE(), @PalletId, @TotalQuantityJob,
                    @DeletedByUserId, @PalletName
                )";

                    await connection.ExecuteAsync(insertDeletedQuery, new
                    {
                        scan.JobNumber,
                        scan.PartID,
                        scan.QRCode,
                        scan.ScanDate,
                        scan.PalletId,
                        scan.TotalQuantityJob,
                        DeletedByUserId = userId,
                        scan.PalletName
                    }, transaction);

                    // Supprimer le scan de la table ScannedTags
                    var deleteQuery = @"
                DELETE FROM ScannedTags
                WHERE QRCode = @QRCode AND PalletId = @PalletId";

                    await connection.ExecuteAsync(deleteQuery, new
                    {
                        request.QRCode,
                        request.PalletId
                    }, transaction);

                    // Nettoyer les anciens enregistrements si nécessaire (garder seulement les 1000 plus récents)
                    var cleanupQuery = @"
                WITH OldestRecords AS (
                    SELECT Id
                    FROM DeletedScans
                    ORDER BY DeletedDate DESC
                    OFFSET 1000 ROWS
                )
                DELETE FROM DeletedScans
                WHERE Id IN (SELECT Id FROM OldestRecords)";

                    await connection.ExecuteAsync(cleanupQuery, null, transaction);

                    // Valider la transaction
                    transaction.Commit();

                    return Ok(new { message = "Scan supprimé avec succès" });
                }
                catch (Exception ex)
                {
                    // Annuler la transaction en cas d'erreur
                    transaction.Rollback();
                    return StatusCode(500, new { message = $"Erreur lors de la suppression: {ex.Message}" });
                }
            }
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
                st.ScanDate,
                u.Username as ScannedByUser 
            FROM ScannedTags st
            JOIN Pallets p ON st.PalletId = p.Id
            LEFT JOIN Users u ON st.ScannedByUserId = u.Id
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
                        s.ScanDate,
                        s.ScannedByUser
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

    private async Task<IActionResult> CreatePackagingCommon(
    int palletId,
    string palLong,
    string palLarg,
    string palHaut,
    string notes,
    bool palFinal,
    IFormFile palletImage = null)
    {
        try
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();

                if (notes.Equals("-"))
                {
                    notes = "";
                }

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
                    return NotFound(new { message = "Aucune information trouvée pour cette palette" });
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
                    .Select(partName => partDetails.FirstOrDefault(p => p.PartName.Equals(partName, StringComparison.OrdinalIgnoreCase)))
                    .Where(p => p != null)
                    .ToList();

                var quantities = partQuantities.Select(p => p.Quantity).ToArray();

                // Create packaging file
                string filePath = _excelValidationService.CreateEmballageFile(
                    jobPalletInfo.JobNumber,
                    jobPalletInfo.PalletName,
                    palLong,
                    palLarg,
                    palHaut,
                    notes,
                    palFinal,
                    orderedPartDetails,
                    quantities,
                    palletImage
                );

                return Ok(new
                {
                    message = "Fichier d'emballage créé avec succès",
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

    [HttpGet("packaging/{palletId:int}")]
    public async Task<IActionResult> CreatePackaging(int palletId, string palLong, string palLarg, string palHaut, string Notes, bool palFinal)
    {
        return await CreatePackagingCommon(palletId, palLong, palLarg, palHaut, Notes, palFinal);
    }

    [HttpPost("packaging/{palletId:int}")]
    public async Task<IActionResult> CreatePackagingWithImage(int palletId, [FromForm] PackagingRequest request)
    {
        return await CreatePackagingCommon(
            palletId,
            request.PalLong,
            request.PalLarg,
            request.PalHaut,
            request.Notes,
            request.PalFinal,
            request.PalletImage);
    }
    [HttpGet("deleted-scans")]
    public async Task<IActionResult> GetDeletedScans([FromQuery] int page = 1, [FromQuery] int pageSize = 50, [FromQuery] string jobNumber = null)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            await connection.OpenAsync();

            // Construire la requête avec filtrage optionnel par jobNumber
            var whereClause = string.IsNullOrEmpty(jobNumber) ? "" : "WHERE d.JobNumber = @JobNumber";

            var query = $@"
        SELECT d.*, u.Username as DeletedByUsername
        FROM DeletedScans d
        LEFT JOIN Users u ON d.DeletedByUserId = u.Id
        {whereClause}
        ORDER BY d.DeletedDate DESC
        OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY";

            // Calculer l'offset pour la pagination
            var offset = (page - 1) * pageSize;

            // Exécuter la requête
            var deletedScans = await connection.QueryAsync(query, new
            {
                Offset = offset,
                PageSize = pageSize,
                JobNumber = jobNumber
            });

            // Obtenir le nombre total d'enregistrements pour la pagination
            var countQuery = $"SELECT COUNT(*) FROM DeletedScans {whereClause}";
            var totalCount = await connection.ExecuteScalarAsync<int>(countQuery, new { JobNumber = jobNumber });

            return Ok(new
            {
                data = deletedScans,
                totalCount,
                page,
                pageSize,
                totalPages = (int)Math.Ceiling((double)totalCount / pageSize)
            });
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