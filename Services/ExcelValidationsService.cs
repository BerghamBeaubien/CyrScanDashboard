// 1. First, create an Excel validation service
using ExcelDataReader;
using System.Collections.Concurrent;
using System.Data;

namespace CyrScanDashboard.Services
{
    public class ExcelValidationService
    {
        private readonly ConcurrentDictionary<int, (Dictionary<string, int> Parts, int TotalQuantityJob)> _jobCache =
            new ConcurrentDictionary<int, (Dictionary<string, int>, int)>();
        private readonly string _excelBasePath = @"P:\DESSINS\DIVERS FAB\MOUAD\Cyramp-Test-XL\";

        public (bool isValid, string message, int? expectedQuantity, int totalQuantityJob) ValidatePart(int jobNumber, string partId, int quantity)
        {
            try
            {
                // Get or load job data
                var (jobData, totalQuantityJob) = GetJobData(jobNumber);
                if (jobData == null)
                {
                    return (false, "Fichier Excel introuvable !", null, 0);
                }

                // Validate part existence and quantity
                if (jobData.TryGetValue(partId, out int expectedQuantity))
                {
                    if (quantity == expectedQuantity)
                    {
                        return (true, "Tag validé !", expectedQuantity, totalQuantityJob);
                    }
                    else
                    {
                        return (false, $"Erreur QTE : Attendu {expectedQuantity}, Scanné {quantity}", expectedQuantity, totalQuantityJob);
                    }
                }
                else
                {
                    return (false, "PartID introuvable dans Excel !", null, totalQuantityJob);
                }
            }
            catch (Exception ex)
            {
                return (false, $"Erreur de validation: {ex.Message}", null, 0);
            }
        }

        private (Dictionary<string, int> Parts, int TotalQuantityJob) GetJobData(int jobNumber)
        {
            // Return from cache if exists
            if (_jobCache.TryGetValue(jobNumber, out var cachedData))
            {
                return cachedData;
            }

            // Otherwise load from file
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            string searchPattern = $"{jobNumber}*.xlsm";
            string[] matchingFiles = Directory.GetFiles(_excelBasePath, searchPattern);

            if (matchingFiles.Length == 0)
            {
                return (null, 0);
            }

            string excelPath = matchingFiles[0];
            var jobData = new Dictionary<string, int>();
            int totalTags = 0;

            using (var fileStream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(fileStream))
            {
                var result = reader.AsDataSet();
                DataTable worksheet = result.Tables["CONTROLE"];

                if (worksheet == null)
                {
                    return (null, 0);
                }

                var totalTagsCell = worksheet.Rows[0][10]; // K1
                totalTags = Convert.ToInt32(totalTagsCell);

                for (int row = 1; row <= totalTags; row++)
                {
                    string excelPartId = worksheet.Rows[row][25]?.ToString()?.Trim(); // Z column
                    int qtyFromExcel = Convert.ToInt32(worksheet.Rows[row][26]?.ToString()); // AA column

                    if (!string.IsNullOrEmpty(excelPartId))
                    {
                        jobData[excelPartId] = qtyFromExcel;
                    }
                }
            }

            // Store in cache
            var cacheData = (jobData, totalTags);
            _jobCache[jobNumber] = cacheData;
            return cacheData;
        }
    }
}