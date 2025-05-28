using ExcelDataReader;
using Microsoft.AspNetCore.Routing.Template;
using System.Collections.Concurrent;
using System.Data;
using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using System.Net.Mail;
using System.Net;
using System.Runtime.InteropServices;

namespace CyrScanDashboard.Services
{
    public class ExcelValidationService
    {
        private readonly ConcurrentDictionary<string, HashSet<string>> _jobCache =
            new ConcurrentDictionary<string, HashSet<string>>();
        private readonly string _excelBasePath = @"P:\CYRAMP\";
        private readonly string _tempFolderPath = @"P:\CYRAMP\TEMPORAIRE\";
        private const string OutputFolderPath = @"P:\Groupe AMP\Feuilles Emballage - CYRSCAN\";
        private const string TemplatePath = @"P:\DESSINS\DIVERS FAB\MOUAD\Template-Emballage.xlsm";

        public ExcelValidationService()
        {
            // Ensure temp directory exists
            if (!Directory.Exists(_tempFolderPath))
            {
                Directory.CreateDirectory(_tempFolderPath);
            }
        }

        public (bool isValid, string message, int totalQty) ValidatePart(string jobNumber, string partId, string qrCode)
        {
            try
            {
                // Get or load job data and total quantity
                var (jobParts, totalQty) = GetJobPartsWithTotal(jobNumber);
                if (jobParts == null)
                {
                    return (false, "Fichier Excel introuvable !", 0);
                }

                // Validate part existence - we only care that the partId exists in the Excel file
                if (jobParts.Contains(partId))
                {
                    return (true, "Tag validé !", totalQty);
                }
                else
                {
                    return (false, "PartID introuvable dans Excel !", totalQty);
                }
            }
            catch (Exception ex)
            {
                return (false, $"Erreur de validation: {ex.Message}", 0);
            }
        }

        private (HashSet<string>, int) GetJobPartsWithTotal(string jobNumber)
        {
            // Return from cache if exists
            if (_jobCache.TryGetValue(jobNumber, out var cachedData))
            {
                // We need to get the total from the Excel file since it's not cached
                int totalQty = CalculateTotalQuantity(jobNumber);
                return (cachedData, totalQty);
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
            string tempFilePath = null;

            try
            {
                // Try to open the file directly first
                using (var fileStream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    return ProcessExcelFileWithTotal(fileStream, jobNumber);
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
                        return ProcessExcelFileWithTotal(fileStream, jobNumber);
                    }
                }
                catch (Exception ex)
                {
                    // Log the error if needed
                    return (null, 0);
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
                return (null, 0);
            }
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

        private (HashSet<string>, int) ProcessExcelFileWithTotal(FileStream fileStream, string jobNumber)
        {
            var jobParts = new HashSet<string>();
            int totalQty = 0;

            using (var reader = ExcelReaderFactory.CreateReader(fileStream))
            {
                var result = reader.AsDataSet();
                DataTable worksheet = result.Tables["CONTROLE"];

                if (worksheet == null)
                {
                    return (null, 0);
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

                    // Calculate total quantity by summing D column values
                    totalQty = 0;
                    for (int row = 1; row <= totalRows; row++)
                    {
                        var dValue = worksheet.Rows[row][3]; // D column (index 3)
                        if (dValue != null)
                        {
                            totalQty += Convert.ToInt32(dValue);
                        }
                    }
                }
                else
                {
                    totalRows = k1Value;
                    totalQty = k1Value; // Set totalQty to K1 value if it's not 0
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
            return (jobParts, totalQty);
        }

        private int CalculateTotalQuantity(string jobNumber)
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
                    return CalculateTotalQuantityFromStream(fileStream);
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
                        return CalculateTotalQuantityFromStream(fileStream);
                    }
                }
                catch (Exception)
                {
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
            catch (Exception)
            {
                return 0;
            }
        }

        private int CalculateTotalQuantityFromStream(FileStream fileStream)
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
                int k1Value = Convert.ToInt32(totalTagsCell);

                // If K1 is 0, calculate total by summing D column
                if (k1Value == 0)
                {
                    int totalQty = 0;
                    int row = 1;

                    while (row < worksheet.Rows.Count)
                    {
                        var dValue = worksheet.Rows[row][3]; // D column (index 3)
                        if (dValue == null || Convert.ToInt32(dValue) == 0)
                        {
                            break;
                        }
                        totalQty += Convert.ToInt32(dValue);
                        row++;
                    }

                    return totalQty;
                }
                else
                {
                    return k1Value;
                }
            }
        }

        public ProjectData GetProjectData(string jobNumber)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                string searchPattern = $"{jobNumber}*.xlsm";
                string[] matchingFiles = Directory.GetFiles(_excelBasePath, searchPattern);

                if (matchingFiles.Length == 0)
                {
                    return null;
                }

                string excelPath = matchingFiles[0];
                string tempFilePath = null;

                try
                {
                    using (var fileStream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        return ProcessProjectData(fileStream);
                    }
                }
                catch (IOException)
                {
                    try
                    {
                        string fileName = Path.GetFileName(excelPath);
                        tempFilePath = Path.Combine(_tempFolderPath, $"temp_{Guid.NewGuid()}_{fileName}");

                        File.Copy(excelPath, tempFilePath, true);

                        using (var fileStream = File.Open(tempFilePath, FileMode.Open, FileAccess.Read))
                        {
                            return ProcessProjectData(fileStream);
                        }
                    }
                    finally
                    {
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
            catch
            {
                return null;
            }
        }

        private ProjectData ProcessProjectData(FileStream fileStream)
        {
            using (var reader = ExcelReaderFactory.CreateReader(fileStream))
            {
                var result = reader.AsDataSet();
                DataTable projetSheet = result.Tables["PROJET"];

                if (projetSheet == null)
                {
                    return null;
                }

                if (projetSheet.Rows[25][1]?.ToString()?.Trim() == "CODE")
                {
                    return new ProjectData
                    {
                        Client = projetSheet.Rows[0][1]?.ToString()?.Trim(),
                        Contact = projetSheet.Rows[2][1]?.ToString()?.Trim(),
                        Projet = projetSheet.Rows[10][1]?.ToString()?.Trim(),
                        Livraison = projetSheet.Rows[13][1]?.ToString()?.Trim(),
                        BonCommande = projetSheet.Rows[21][4]?.ToString()?.Trim(),
                        EcheanceB20 = projetSheet.Rows[19][1]?.ToString()?.Trim(),
                        EcheanceB21 = projetSheet.Rows[20][1]?.ToString()?.Trim(),
                        Contenu = projetSheet.Rows[26][4]?.ToString()?.Trim(),
                        BcClient = projetSheet.Rows[17][1]?.ToString()?.Trim(),
                        ChargeProjet = projetSheet.Rows[7][1]?.ToString()?.Trim(),
                        Fini = projetSheet.Rows[21][3]?.ToString()?.Trim()
                    };
                } else
                {
                    return new ProjectData
                    {
                        Client = projetSheet.Rows[0][2]?.ToString()?.Trim(),
                        Contact = projetSheet.Rows[5][2]?.ToString()?.Trim(),
                        Projet = projetSheet.Rows[10][2]?.ToString()?.Trim(),
                        Livraison = projetSheet.Rows[13][1]?.ToString()?.Trim(),
                        BonCommande = projetSheet.Rows[21][4]?.ToString()?.Trim(),
                        EcheanceB20 = projetSheet.Rows[19][1]?.ToString()?.Trim(),
                        EcheanceB21 = projetSheet.Rows[20][1]?.ToString()?.Trim(),
                        Contenu = projetSheet.Rows[26][4]?.ToString()?.Trim(),
                        BcClient = projetSheet.Rows[17][1]?.ToString()?.Trim(),
                        ChargeProjet = projetSheet.Rows[7][1]?.ToString()?.Trim(),
                        Fini = projetSheet.Rows[21][3]?.ToString()?.Trim()
                    };
                }
                
            }
        }

        public List<PartPackagingData> GetPartPackagingDetails(string jobNumber, List<string> partNames)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                string searchPattern = $"{jobNumber}*.xlsm";
                string[] matchingFiles = Directory.GetFiles(_excelBasePath, searchPattern);

                if (matchingFiles.Length == 0)
                {
                    return new List<PartPackagingData>();
                }

                string excelPath = matchingFiles[0];
                string tempFilePath = null;

                try
                {
                    using (var fileStream = File.Open(excelPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        return ProcessPartPackagingDetails(fileStream, partNames);
                    }
                }
                catch (IOException)
                {
                    try
                    {
                        string fileName = Path.GetFileName(excelPath);
                        tempFilePath = Path.Combine(_tempFolderPath, $"temp_{Guid.NewGuid()}_{fileName}");

                        File.Copy(excelPath, tempFilePath, true);

                        using (var fileStream = File.Open(tempFilePath, FileMode.Open, FileAccess.Read))
                        {
                            return ProcessPartPackagingDetails(fileStream, partNames);
                        }
                    }
                    finally
                    {
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
            catch
            {
                return new List<PartPackagingData>();
            }
        }

        private List<PartPackagingData> ProcessPartPackagingDetails(FileStream fileStream, List<string> partNames)
        {
            var partDetails = new List<PartPackagingData>();

            using (var reader = ExcelReaderFactory.CreateReader(fileStream))
            {
                var result = reader.AsDataSet();
                DataTable projetSheet = result.Tables["PROJET"];

                if (projetSheet == null)
                {
                    return partDetails;
                }

                foreach (string partName in partNames)
                {
                    for (int i = 0; i < projetSheet.Rows.Count; i++)
                    {
                        string excelPartName = projetSheet.Rows[i][3]?.ToString()?.Trim(); // Col D
                        if (excelPartName == partName)
                        {
                            string valHauteur = projetSheet.Rows[i][5]?.ToString()?.Trim(); // Col F
                            string valLargeur = projetSheet.Rows[i][6]?.ToString()?.Trim(); // Col G
                            string codeMateriel = projetSheet.Rows[i][0]?.ToString()?.Trim(); // Col A

                            // Look up codeMateriel in the M2:M19 range
                            string valFromMRange = null;
                            for (int j = 1; j <= 18; j++)
                            {
                                if (projetSheet.Rows[j][14]?.ToString()?.Trim() == codeMateriel)
                                {
                                    valFromMRange = projetSheet.Rows[j][17]?.ToString()?.Trim(); // Column R is index 17
                                    break;
                                }
                            }

                            partDetails.Add(new PartPackagingData
                            {
                                PartName = partName,
                                Hauteur = valHauteur,
                                Largeur = valLargeur,
                                CodeMateriel = codeMateriel,
                                MassePi2 = valFromMRange
                            });
                        }
                    }
                }
            }

            return partDetails;
        }

        public string CreateEmballageFile(string jobNumber, string paletteName, string palLong, string palLarg, string palHaut, string Notes,
    bool palFinal, List<PartPackagingData> partDetails, int[] quantities, IFormFile palletImage = null)
        {
            var projectData = GetProjectData(jobNumber);
            if (projectData == null)
            {
                throw new Exception("Impossible de trouver les données du projet");
            }

            string outputFileName = $"{jobNumber} {paletteName} EMBALLAGE.xlsm";
            if (palFinal)
            {
                outputFileName = $"{jobNumber} {paletteName} EMBALLAGE (FINALE).xlsm";
            }
            string outputPath = Path.Combine(OutputFolderPath, outputFileName);

            // Clean palette name by removing "PAL"
            paletteName = paletteName.Replace("PAL", "");

            using (var workbook = new XLWorkbook(TemplatePath))
            {
                var worksheet = workbook.Worksheet(1);
                worksheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
                int startRow = 11;
                double masseTotal = 0;
                int qteTotal = 0;

                for (int i = 0; i < partDetails.Count; i++)
                {
                    var detail = partDetails[i];
                    int qty = quantities[i];

                    worksheet.Cell(startRow, 1).Value = detail.PartName;
                    worksheet.Cell(startRow, 4).Value = detail.Hauteur;
                    worksheet.Cell(startRow, 6).Value = detail.Largeur;
                    worksheet.Cell(startRow, 3).Value = qty;
                    worksheet.Cell(startRow, 7).Value = CalculSurface(detail.Hauteur, detail.Largeur);

                    masseTotal += (Convert.ToDouble(detail.MassePi2) * qty);
                    qteTotal += qty;
                    startRow++;
                }

                // Set print area
                worksheet.PageSetup.PrintAreas.Clear();
                worksheet.PageSetup.PrintAreas.Add("A1:G" + (startRow - 1));

                // Add a thick border around the print area
                worksheet.Range($"A{startRow}:G{startRow + 2}").Style.Border.OutsideBorder = XLBorderStyleValues.None;
                worksheet.Range("A1:G" + (startRow-1)).Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                worksheet.Range("A1:G" + (startRow-1)).Style.Border.OutsideBorderColor = XLColor.Black;

                // Lock all cells, then unlock the 3 rows after the last data row
                worksheet.Protect();
                worksheet.Range("A1:G" + (startRow - 1)).Style.Protection.Locked = true;
                worksheet.Range($"A{startRow}:G{startRow + 2}").Style.Protection.Locked = false;
                worksheet.Range($"A{startRow}:G{startRow + 2}").Value = "";
                worksheet.Cell($"A{startRow}").Value = "Notes:";
                worksheet.Cell($"B{startRow}").Value = Notes;

                var echeanceValue = string.IsNullOrEmpty(projectData.EcheanceB21)
                    ? projectData.EcheanceB20
                    : projectData.EcheanceB21;

                worksheet.Cell("B1").Value = projectData.Client;
                worksheet.Cell("B2").Value = projectData.Contact;
                worksheet.Cell("B3").Value = projectData.Projet;
                worksheet.Cell("B4").Value = projectData.Livraison;
                worksheet.Cell("B5").Value = projectData.BonCommande;
                worksheet.Cell("B6").Value = Convert.ToDateTime(echeanceValue).ToString("MM/dd/yyyy");
                worksheet.Cell("B7").Value = qteTotal + " / " + projectData.Contenu;
                worksheet.Cell("B8").Value = masseTotal;
                worksheet.Cell("B9").Value = palFinal ? paletteName + " (FINALE)" : paletteName;
                worksheet.Cell("F1").FormulaA1 = $"=\"(\"&{jobNumber}&\")\"";
                worksheet.Cell("F9").Value = palLong + " X " + palLarg + " X " + palHaut;
                worksheet.Cell("G3").Value = projectData.BcClient;
                worksheet.Cell("G4").Value = projectData.ChargeProjet;
                worksheet.Cell("G8").Value = projectData.Fini;

                workbook.SaveAs(outputPath);

                // Save as PDF (Excel Interop API)
                string pdfFileName = Path.ChangeExtension(outputFileName, ".pdf");
                string pdfPath = Path.Combine(OutputFolderPath, pdfFileName);
                string imagePath = "";
                ExportToPdf(outputPath, pdfPath);

                if (palletImage != null && palletImage.Length > 0)
                {
                    string imageFileName = Path.ChangeExtension(outputFileName, ".jpg");
                    imagePath = Path.Combine(OutputFolderPath, imageFileName);
                    SavePalletImage(palletImage, imagePath);
                }

                SendNotificationEmail(jobNumber, pdfPath, projectData.ChargeProjet, palFinal, paletteName, imagePath);

                return pdfPath;
            }
        }

        private void ExportToPdf(string excelPath, string pdfPath)
        {
            dynamic excel = null;

            try
            {
                // Use late binding to avoid version-specific dependencies
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.ScreenUpdating = false;  // Added for performance
                excel.EnableEvents = false;    // Added for performance
                excel.AskToUpdateLinks = false; // Avoid prompts

                // Open workbook with optimizations
                dynamic workbook = excel.Workbooks.Open(excelPath,
                    UpdateLinks: false,    // Don't update links
                    ReadOnly: true         // Open read-only for better performance
                );

                try
                {
                    // Export as PDF with minimal options
                    workbook.ExportAsFixedFormat(0, pdfPath, 0, false, false);

                    // Close workbook immediately after export
                    workbook.Close(false);
                }
                finally
                {
                    // Release workbook immediately
                    if (workbook != null)
                    {
                        Marshal.FinalReleaseComObject(workbook);
                        workbook = null;
                    }
                }
            }
            finally
            {
                // Clean up Excel
                if (excel != null)
                {
                    excel.Quit();
                    Marshal.FinalReleaseComObject(excel);
                    excel = null;
                }

                // Immediate garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void SendNotificationEmail(string jobNumber, string pdfPath, string chargeProjet, bool palFinale, string paletteName, string imagePath)
        {
            string courrielMolly = "molly.drapeau@cyrell.qc.ca";
            string courrielExpedition = "expedition@cyrell.qc.ca";
            string chargeEmail = GetEmailForChargeProjet(chargeProjet);

            var toList = new List<string> { courrielMolly };
            if (!string.IsNullOrWhiteSpace(chargeEmail))
                toList.Add(chargeEmail);


            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                Credentials = new NetworkCredential("noreply.cyrscan@gmail.com", "lpye mmgp tggh hcca")
            };

            using (var message = new MailMessage())
            {
                message.From = new MailAddress("noreply.cyrscan@gmail.com", "Système d'emballage automatisé");
                message.To.Add(courrielExpedition);
                message.Subject = $"Palette #{paletteName} - Job #{jobNumber}";
                message.Body = $"Bonjour,\n\nVeuillez trouver ci-joint le document d'emballage pour la palette {paletteName} de la job #{jobNumber}.\n\nCordialement,\nSystème d'emballage automatisé. Ciao Bye Cimer";
                message.Attachments.Add(new Attachment(pdfPath));
                if (!string.IsNullOrWhiteSpace(imagePath) && File.Exists(imagePath))
                {
                    message.Attachments.Add(new Attachment(imagePath));
                }

                smtp.Send(message);
            }

            if (palFinale)
            {
                using (var message = new MailMessage())
                {
                    message.From = new MailAddress("noreply.cyrscan@gmail.com", "Système d'emballage automatisé");
                    foreach (var to in toList)
                        message.To.Add(to);

                    message.Subject = $"Palette finale - Job #{jobNumber}";
                    message.Body = $"Bonjour,\n\nVeuillez trouver ci-joint le document d'emballage pour la palette finale de la job #{jobNumber}.\n\nCordialement,\nSystème d'emballage automatisé.";
                    message.Attachments.Add(new Attachment(pdfPath));
                    if (!string.IsNullOrWhiteSpace(imagePath) && File.Exists(imagePath))
                    {
                        message.Attachments.Add(new Attachment(imagePath));
                    }

                    smtp.Send(message);
                }
            }
        }

        private string GetEmailForChargeProjet(string chargeProjet)
        {
            // Mapping of projet managers to their email addresses
            Dictionary<string, string> emailMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    { "Abdelalim", "abdelalim.elhajri@cyrell.qc.ca" },
                    { "Bhupinder", "bhupinder.mahey@cyrell.qc.ca" },
                    { "Cassie Giguère", "cassie.giguere@cyrell.qc.ca" },
                    { "Amélie Poirier-Borduas", "amelie.poirier-borduas@cyrell.qc.ca" },
                    { "David Théroux", "david.theroux@cyrell.qc.ca" },
                    { "Emeric", "emeric.buchen@cyrell.qc.ca" },
                    { "François Audy", "francois.audy@cyrell.qc.ca" },
                    { "Gabriel", "gabriel.borduas@cyrell.qc.ca" },
                    { "Hoàng T.", "vuhoang.tran@cyrell.qc.ca" },
                    { "J-F Urbain", "jean-francois_urbain@cyrell.qc.ca" },
                    { "Jose", "Jose.Lobato@cyrell.qc.ca" },
                    { "Meriem", "meriem.bouslimi@cyrell.qc.ca" },
                    { "Mouad Khalladi", "mouad.khalladi@cyrell.qc.ca" },
                    { "Pierre Boulanger", "pierre.boulanger@cyrell.qc.ca" }
                };

            if (emailMappings.TryGetValue(chargeProjet, out string email))
            {
                return email;
            }

            return null; // Return null if no specific email is found
        }

        private void SavePalletImage(IFormFile imageFile, string imagePath)
        {
            try
            {
                // Save the image
                using (var stream = new FileStream(imagePath, FileMode.Create))
                {
                    imageFile.CopyTo(stream);
                }
            }
            catch (Exception ex)
            {
                // Log the error but don't fail the entire operation
                // You might want to use your logging framework here
                Console.WriteLine($"Error saving image: {ex.Message}");
            }
        }

        private double CalculSurface(string valHauteur, string valLargeur)
        {
            try
            {
                double hauteur = Convert.ToDouble(valHauteur, CultureInfo.InvariantCulture);
                double largeur = Convert.ToDouble(valLargeur, CultureInfo.InvariantCulture);
                return (hauteur * largeur) / 144;
            }
            catch (FormatException)
            {
                return 0;
            }
        }
    }

    public class ProjectData
    {
        public string Client { get; set; }
        public string Contact { get; set; }
        public string Projet { get; set; }
        public string Livraison { get; set; }
        public string BonCommande { get; set; }
        public string EcheanceB20 { get; set; }
        public string EcheanceB21 { get; set; }
        public string Contenu { get; set; }
        public string BcClient { get; set; }
        public string ChargeProjet { get; set; }
        public string Fini { get; set; }
    }

    public class PartPackagingData
    {
        public string PartName { get; set; }
        public string Hauteur { get; set; }
        public string Largeur { get; set; }
        public string CodeMateriel { get; set; }
        public string MassePi2 { get; set; }
    }

    // Request model for the POST endpoint
    public class PackagingRequest
    {
        public string PalLong { get; set; }
        public string PalLarg { get; set; }
        public string PalHaut { get; set; }
        public string Notes { get; set; }
        public bool PalFinal { get; set; }
        public IFormFile PalletImage { get; set; }
    }

}