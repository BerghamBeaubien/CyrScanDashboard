namespace CyrScanDashboard.Models
{
    public class Material
    {
        public int MatID { get; set; }
        public string Description { get; set; }
        public string Type { get; set; } // "Sheet Metal", "Extrusion", "Laine", "Other"
        public decimal? Thickness { get; set; }
        public string SheetSize { get; set; }
    }

    public class Location
    {
        public int LocID { get; set; }
        public string Description { get; set; }
        public int? MatID { get; set; }
        public int XPosition { get; set; }
        public int YPosition { get; set; }
        public decimal Quantity { get; set; } = 0;
    }

    public class LocationViewModel : Location
    {
        public string MaterialDescription { get; set; }
        public decimal ReservedQuantity { get; set; }
        public decimal AvailableQuantity => Quantity - ReservedQuantity;
    }

    public class Reservation
    {
        public int ReservationID { get; set; }
        public int LocID { get; set; }
        public decimal Quantity { get; set; }
        public string JobDescription { get; set; }
        public DateTime ReservationDate { get; set; }
        public string Status { get; set; } = "Pending"; // "Pending", "Confirmed", "Used", "Canceled"
    }

    public class ReservationViewModel : Reservation
    {
        public string LocationDescription { get; set; }
        public int? MatID { get; set; }
        public string MaterialDescription { get; set; }
    }

    public class ReservationStatusUpdate
    {
        public string Status { get; set; } // "Pending", "Confirmed", "Used", "Canceled"
    }
}