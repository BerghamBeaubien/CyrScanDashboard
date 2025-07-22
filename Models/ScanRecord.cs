namespace CyrScanDashboard.Models;

public class ScanRecord
{
    public int Id { get; set; }
    public string JobNumber { get; set; }
    public string PartID { get; set; }
    public string QRCode { get; set; }
    public DateTime ScanDate { get; set; }
}

public class ScanRequest
{
    public string JobNumber { get; set; }
    public string PartID { get; set; }
    public string QRCode { get; set; }
    public int PalletId { get; set; }
}

public class DeleteScanRequest
{
    public string QRCode { get; set; }
    public int PalletId { get; set; }
}

public class Pallet
{
    public int Id { get; set; }
    public string JobNumber { get; set; }
    public string Name { get; set; }
    public DateTime CreatedDate { get; set; }
    public int SequenceNumber { get; set; }
    public int ScannedItems { get; set; }
    public string hasPackagingBeenGenerated { get; set; }
    public string packagingImagePath { get; set; }
    public string packagingPdfPath { get; set; }
}

public class CreatePalletRequest
{
    public string JobNumber { get; set; }
}

public class UpdatePalletRequest
{
    public string Name { get; set; }
}

public class JobSummary
{
    public string JobNumber { get; set; }
    public int TotalParts { get; set; }
    public int TotalScanned { get; set; }
    public int TotalPallets { get; set; }
    public int TotalExpected { get; set; }
    public DateTime LastScanDate { get; set; }
}

public class JobPalletInfo
{
    public string JobNumber { get; set; }
    public string PalletName { get; set; }
    public int SequenceNumber { get; set; }
    public int Id { get; set; }
}

public class PartQuantity
{
    public string PartId { get; set; }
    public int Quantity { get; set; }
}