namespace YourProject.Models;

public class ScanRecord
{
    public int Id { get; set; }
    public int JobNumber { get; set; }
    public string PartID { get; set; }
    public int Quantity { get; set; }
    public int ScannedQuantity { get; set; }
    public DateTime ScanDate { get; set; }
}