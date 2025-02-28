using ExcelDataReader;
using System.Collections.Concurrent;
using System.Data;
using System.IO;

namespace CyrScanDashboard.Services
{
    public class ExcelValidationService
    {
        private readonly ConcurrentDictionary<string, HashSet<string>> _jobCache =
            new ConcurrentDictionary<string, HashSet<string>>();
        private readonly string _excelBasePath = @"P:\DESSINS\DIVERS FAB\MOUAD\Cyramp-Test-XL\";
        private readonly string _tempFolderPath = @"P:\DESSINS\DIVERS FAB\MOUAD\Cyramp-Test-XL\Temp (NE PAS SUPPRIMER)\";

        public ExcelValidationService()
        {
            // Ensure temp directory exists
            if (!Directory.Exists(_tempFolderPath))
            {
                Directory.CreateDirectory(_tempFolderPath);
            }
        }

        public (bool isValid, string message) ValidatePart(string jobNumber, string partId, string qrCode)
        {
            try
            {
                // Get or load job data
                var jobParts = GetJobParts(jobNumber);
                if (jobParts == null)
                {
                    return (false, "Fichier Excel introuvable !");
                }

                // Validate part existence - we only care that the partId exists in the Excel file
                if (jobParts.Contains(partId))
                {
                    return (true, "Tag validé !");
                }
                else
                {
                    return (false, "PartID introuvable dans Excel !");
                }
            }
            catch (Exception ex)
            {
                return (false, $"Erreur de validation: {ex.Message}");
            }
        }

        private HashSet<string> GetJobParts(string jobNumber)
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
                return null;
            }

            string excelPath = matchingFiles[0];
            var jobParts = new HashSet<string>();
            string tempFilePath = null;

            try
            {
                // Try to open the file directly first
                using (var fileStream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    return ProcessExcelFile(fileStream, jobNumber);
                }
            }
            catch (IOException) // File is likely locked
            {
                try
                {
                    // Create a temporary copy
                    string fileName = Path.GetFileName(excelPath);
                    tempFilePath = Path.Combine(_tempFolderPath, $"temp_{Guid.NewGuid()}_{fileName}");

                    // Copy with FileShare.ReadWrite to allow copying even when file is in use
                    File.Copy(excelPath, tempFilePath, true);

                    using (var fileStream = File.Open(tempFilePath, FileMode.Open, FileAccess.Read))
                    {
                        return ProcessExcelFile(fileStream, jobNumber);
                    }
                }
                catch (Exception ex)
                {
                    // Log the error if needed
                    return null;
                }
                finally
                {
                    // Clean up temp file if it exists
                    if (tempFilePath != null && File.Exists(tempFilePath))
                    {
                        try
                        {
                            File.Delete(tempFilePath);
                        }
                        catch
                        {
                            // Ignore cleanup errors
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Handle other exceptions
                return null;
            }
        }

        private HashSet<string> ProcessExcelFile(FileStream fileStream, string jobNumber)
        {
            var jobParts = new HashSet<string>();

            using (var reader = ExcelReaderFactory.CreateReader(fileStream))
            {
                var result = reader.AsDataSet();
                DataTable worksheet = result.Tables["CONTROLE"];

                if (worksheet == null)
                {
                    return null;
                }

                var totalTagsCell = worksheet.Rows[0][10]; // K1
                int k1Value = Convert.ToInt32(totalTagsCell);

                // If K1 is 0, determine the total by summing values in column D until we hit a 0
                int totalRows;
                if (k1Value == 0)
                {
                    totalRows = 1; // Start from row 2 (index 1)
                    while (totalRows < worksheet.Rows.Count)
                    {
                        var dValue = worksheet.Rows[totalRows][3]; // D column (index 3)
                        if (dValue == null || Convert.ToInt32(dValue) == 0)
                        {
                            break;
                        }
                        totalRows++;
                    }
                }
                else
                {
                    totalRows = k1Value;
                }

                for (int row = 1; row <= totalRows; row++)
                {
                    string excelPartId = worksheet.Rows[row][0]?.ToString()?.Trim(); // A column (index 0)

                    // If A column is empty, use C column
                    if (string.IsNullOrEmpty(excelPartId))
                    {
                        string cColumnPartId = worksheet.Rows[row][2]?.ToString()?.Trim(); // C column (index 2)
                        if (!string.IsNullOrEmpty(cColumnPartId))
                        {
                            // Just add the part ID once to the HashSet regardless of quantity
                            jobParts.Add(cColumnPartId);
                        }
                    }
                    else
                    {
                        // Original behavior for when A column has data
                        jobParts.Add(excelPartId);
                    }
                }
            }

            // Store in cache
            _jobCache[jobNumber] = jobParts;
            return jobParts;
        }

        public IEnumerable<object> GetExcelJobDetails(string jobNumber)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                string searchPattern = $"{jobNumber}*.xlsm";
                string[] matchingFiles = Directory.GetFiles(_excelBasePath, searchPattern);

                if (matchingFiles.Length == 0)
                {
                    return new List<object>();
                }

                string excelPath = matchingFiles[0];
                var jobParts = new List<object>();
                var partCounters = new Dictionary<string, int>();
                string tempFilePath = null;

                try
                {
                    // Try to open the file directly first
                    using (var fileStream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        return ProcessExcelJobDetails(fileStream);
                    }
                }
                catch (IOException) // File is likely locked
                {
                    try
                    {
                        // Create a temporary copy
                        string fileName = Path.GetFileName(excelPath);
                        tempFilePath = Path.Combine(_tempFolderPath, $"temp_{Guid.NewGuid()}_{fileName}");

                        // Copy with FileShare.ReadWrite to allow copying even when file is in use
                        File.Copy(excelPath, tempFilePath, true);

                        using (var fileStream = File.Open(tempFilePath, FileMode.Open, FileAccess.Read))
                        {
                            return ProcessExcelJobDetails(fileStream);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log the error if needed
                        return new List<object>();
                    }
                    finally
                    {
                        // Clean up temp file if it exists
                        if (tempFilePath != null && File.Exists(tempFilePath))
                        {
                            try
                            {
                                File.Delete(tempFilePath);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    return new List<object>();
                }
            }
            catch (Exception ex)
            {
                // Si une erreur se produit, retourner une liste vide
                return new List<object>();
            }
        }

        private IEnumerable<object> ProcessExcelJobDetails(FileStream fileStream)
        {
            var jobParts = new List<object>();
            var partCounters = new Dictionary<string, int>();

            using (var reader = ExcelReaderFactory.CreateReader(fileStream))
            {
                var result = reader.AsDataSet();
                DataTable worksheet = result.Tables["CONTROLE"];

                if (worksheet == null)
                {
                    return new List<object>();
                }

                var totalTagsCell = worksheet.Rows[0][10]; // K1
                int k1Value = Convert.ToInt32(totalTagsCell);

                // If K1 is 0, determine the total by summing values in column D until we hit a 0
                int totalRows;
                if (k1Value == 0)
                {
                    totalRows = 1; // Start from row 2 (index 1)
                    while (totalRows < worksheet.Rows.Count)
                    {
                        var dValue = worksheet.Rows[totalRows][3]; // D column (index 3)
                        if (dValue == null || Convert.ToInt32(dValue) == 0)
                        {
                            break;
                        }
                        totalRows++;
                    }
                }
                else
                {
                    totalRows = k1Value;
                }

                for (int row = 1; row <= totalRows; row++)
                {
                    string partId = worksheet.Rows[row][0]?.ToString()?.Trim(); // Colonne A

                    // If A column is empty, use C column and D column for quantity
                    if (string.IsNullOrEmpty(partId))
                    {
                        string cColumnPartId = worksheet.Rows[row][2]?.ToString()?.Trim(); // C column (index 2)
                        if (!string.IsNullOrEmpty(cColumnPartId))
                        {
                            // Get the count from D column
                            int dColumnCount = Convert.ToInt32(worksheet.Rows[row][3]); // D column (index 3)

                            // For each quantity in D column, create a separate part entry
                            for (int i = 0; i < dColumnCount; i++)
                            {
                                // Track sequence numbers for this part
                                if (!partCounters.ContainsKey(cColumnPartId))
                                {
                                    partCounters[cColumnPartId] = 1;
                                }
                                else
                                {
                                    partCounters[cColumnPartId]++;
                                }

                                var partInfo = new
                                {
                                    PartID = cColumnPartId,
                                    SeqNumber = partCounters[cColumnPartId]
                                };
                                jobParts.Add(partInfo);
                            }
                        }
                    }
                    else if (!string.IsNullOrEmpty(partId))
                    {
                        // Original behavior for when A column has data
                        if (!partCounters.ContainsKey(partId))
                        {
                            partCounters[partId] = 1;
                        }
                        else
                        {
                            partCounters[partId]++;
                        }

                        var partInfo = new
                        {
                            PartID = partId,
                            SeqNumber = partCounters[partId]
                        };
                        jobParts.Add(partInfo);
                    }
                }
            }

            return jobParts;
        }

        public int GetK1Value(string jobNumber)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                string searchPattern = $"{jobNumber}*.xlsm";
                string[] matchingFiles = Directory.GetFiles(_excelBasePath, searchPattern);

                if (matchingFiles.Length == 0)
                {
                    return 0;
                }

                string excelPath = matchingFiles[0];
                string tempFilePath = null;

                try
                {
                    // Try to open the file directly first
                    using (var fileStream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        return GetK1ValueFromStream(fileStream);
                    }
                }
                catch (IOException) // File is likely locked
                {
                    try
                    {
                        // Create a temporary copy
                        string fileName = Path.GetFileName(excelPath);
                        tempFilePath = Path.Combine(_tempFolderPath, $"temp_{Guid.NewGuid()}_{fileName}");

                        // Copy with FileShare.ReadWrite to allow copying even when file is in use
                        File.Copy(excelPath, tempFilePath, true);

                        using (var fileStream = File.Open(tempFilePath, FileMode.Open, FileAccess.Read))
                        {
                            return GetK1ValueFromStream(fileStream);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log the error if needed
                        return 0;
                    }
                    finally
                    {
                        // Clean up temp file if it exists
                        if (tempFilePath != null && File.Exists(tempFilePath))
                        {
                            try
                            {
                                File.Delete(tempFilePath);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        private int GetK1ValueFromStream(FileStream fileStream)
        {
            using (var reader = ExcelReaderFactory.CreateReader(fileStream))
            {
                var result = reader.AsDataSet();
                DataTable worksheet = result.Tables["CONTROLE"];

                if (worksheet == null)
                {
                    return 0;
                }

                var totalTagsCell = worksheet.Rows[0][10]; // K1
                return Convert.ToInt32(totalTagsCell);
            }
        }
    }
}