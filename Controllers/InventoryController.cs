using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Dapper;
using CyrScanDashboard.Models;
using System.Data;
using System.Threading.Tasks;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using Location = CyrScanDashboard.Models.Location;

namespace CyrScanDashboard.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class InventoryController : ControllerBase
    {
        private readonly string _connectionString = "Server=cyrscan-server,1433;Database=CyrScanDB;User Id=Serveur-CyrScan;Password=admin;TrustServerCertificate=True;";

        // GET: api/Inventory/Materials
        [HttpGet("Materials")]
        public async Task<IActionResult> GetMaterials()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                var query = @"
                SELECT 
                    MatID,
                    Description,
                    Type,
                    Thickness,
                    SheetSize
                FROM Materials
                ORDER BY Type, Description";

                var materials = await connection.QueryAsync<Material>(query);
                return Ok(materials);
            }
        }

        // GET: api/Inventory/Materials/{id}
        [HttpGet("Materials/{id}")]
        public async Task<IActionResult> GetMaterial(int id)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                var query = @"
                SELECT 
                    MatID,
                    Description,
                    Type,
                    Thickness,
                    SheetSize
                FROM Materials
                WHERE MatID = @Id";

                var material = await connection.QueryFirstOrDefaultAsync<Material>(query, new { Id = id });

                if (material == null)
                    return NotFound();

                return Ok(material);
            }
        }

        // POST: api/Inventory/Materials
        [HttpPost("Materials")]
        public async Task<IActionResult> AddMaterial(Material material)
        {
            // Validate the material data
            if (material.Type == "Sheet Metal" && (material.Thickness == null || string.IsNullOrEmpty(material.SheetSize)))
            {
                return BadRequest("Sheet Metal materials require Thickness and Sheet Size.");
            }

            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();

                material.MatID = await connection.ExecuteScalarAsync<int>("SELECT ISNULL(MAX(MatID), 0) + 1 FROM Materials");

                var query = @"
                INSERT INTO Materials (MatID,Description, Type, Thickness, SheetSize)
                VALUES (@MatID, @Description, @Type, @Thickness, @SheetSize);
                SELECT CAST(SCOPE_IDENTITY() as int)";

                var id = await connection.ExecuteScalarAsync<int>(query, material);
                material.MatID = id;

                return CreatedAtAction(nameof(GetMaterial), new { id = material.MatID }, material);
            }
        }

        // PUT: api/Inventory/Materials/{id}
        [HttpPut("Materials/{id}")]
        public async Task<IActionResult> UpdateMaterial(int id, Material material)
        {
            if (id != material.MatID)
                return BadRequest("ID mismatch");

            // Validate the material data
            if (material.Type == "Sheet Metal" && (material.Thickness == null || string.IsNullOrEmpty(material.SheetSize)))
            {
                return BadRequest("Sheet Metal materials require Thickness and Sheet Size.");
            }

            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();

                var query = @"
                UPDATE Materials
                SET Description = @Description,
                    Type = @Type,
                    Thickness = @Thickness,
                    SheetSize = @SheetSize
                WHERE MatID = @MatID";

                var affected = await connection.ExecuteAsync(query, material);

                if (affected == 0)
                    return NotFound();

                return NoContent();
            }
        }

        // DELETE: api/Inventory/Materials/{id}
        [HttpDelete("Materials/{id}")]
        public async Task<IActionResult> DeleteMaterial(int id)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();

                // Check if material is used in any location
                var checkQuery = "SELECT COUNT(*) FROM Locations WHERE MatID = @Id";
                var useCount = await connection.ExecuteScalarAsync<int>(checkQuery, new { Id = id });

                if (useCount > 0)
                    return BadRequest("Cannot delete material that is being used in locations.");

                var query = "DELETE FROM Materials WHERE MatID = @Id";
                var affected = await connection.ExecuteAsync(query, new { Id = id });

                if (affected == 0)
                    return NotFound();

                return NoContent();
            }
        }

        // GET: api/Inventory/Locations
        [HttpGet("Locations")]
        public async Task<IActionResult> GetLocations()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                var query = @"
                SELECT 
                    L.LocID,
                    L.Description,
                    L.MatID,
                    M.Description AS MaterialDescription,
                    L.XPosition,
                    L.YPosition,
                    L.Quantity,
                    (SELECT COALESCE(SUM(R.Quantity), 0) 
                     FROM Reservations R 
                     WHERE R.LocID = L.LocID AND R.Status IN ('Pending', 'Confirmed')) AS ReservedQuantity
                FROM Locations L
                LEFT JOIN Materials M ON L.MatID = M.MatID
                ORDER BY L.Description";

                var locations = await connection.QueryAsync<LocationViewModel>(query);
                return Ok(locations);
            }
        }

        // POST: api/Inventory/Locations
        [HttpPost("Locations")]
        public async Task<IActionResult> AddLocation(Location location)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();

                // Verify that the material exists
                if (location.MatID.HasValue)
                {
                    var materialCheck = "SELECT COUNT(*) FROM Materials WHERE MatID = @MatID";
                    var materialExists = await connection.ExecuteScalarAsync<int>(materialCheck, new { MatID = location.MatID });

                    if (materialExists == 0)
                        return BadRequest("Referenced material does not exist.");
                }

                var query = @"
                INSERT INTO Locations (Description, MatID, XPosition, YPosition, Quantity)
                VALUES (@Description, @MatID, @XPosition, @YPosition, @Quantity);
                SELECT CAST(SCOPE_IDENTITY() as int)";

                var id = await connection.ExecuteScalarAsync<int>(query, location);
                location.LocID = id;

                return CreatedAtAction(nameof(GetLocation), new { id = location.LocID }, location);
            }
        }

        // GET: api/Inventory/Locations/{id}
        [HttpGet("Locations/{id}")]
        public async Task<IActionResult> GetLocation(int id)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                var query = @"
                SELECT 
                    L.LocID,
                    L.Description,
                    L.MatID,
                    M.Description AS MaterialDescription,
                    L.XPosition,
                    L.YPosition,
                    L.Quantity,
                    (SELECT COALESCE(SUM(R.Quantity), 0) 
                     FROM Reservations R 
                     WHERE R.LocID = L.LocID AND R.Status IN ('Pending', 'Confirmed')) AS ReservedQuantity
                FROM Locations L
                LEFT JOIN Materials M ON L.MatID = M.MatID
                WHERE L.LocID = @Id";

                var location = await connection.QueryFirstOrDefaultAsync<LocationViewModel>(query, new { Id = id });

                if (location == null)
                    return NotFound();

                return Ok(location);
            }
        }

        // PUT: api/Inventory/Locations/{id}
        [HttpPut("Locations/{id}")]
        public async Task<IActionResult> UpdateLocation(int id, Location location)
        {
            if (id != location.LocID)
                return BadRequest("ID mismatch");

            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();

                // Verify that the material exists
                if (location.MatID.HasValue)
                {
                    var materialCheck = "SELECT COUNT(*) FROM Materials WHERE MatID = @MatID";
                    var materialExists = await connection.ExecuteScalarAsync<int>(materialCheck, new { MatID = location.MatID });

                    if (materialExists == 0)
                        return BadRequest("Referenced material does not exist.");
                }

                var query = @"
                UPDATE Locations
                SET Description = @Description,
                    MatID = @MatID,
                    XPosition = @XPosition,
                    YPosition = @YPosition,
                    Quantity = @Quantity
                WHERE LocID = @LocID";

                var affected = await connection.ExecuteAsync(query, location);

                if (affected == 0)
                    return NotFound();

                return NoContent();
            }
        }

        // POST: api/Inventory/Reservations
        [HttpPost("Reservations")]
        public async Task<IActionResult> CreateReservation(Reservation reservation)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();

                // Check if location exists
                var locationCheck = "SELECT COUNT(*) FROM Locations WHERE LocID = @LocID";
                var locationExists = await connection.ExecuteScalarAsync<int>(locationCheck, new { LocID = reservation.LocID });

                if (locationExists == 0)
                    return BadRequest("Location does not exist.");

                // Calculate available quantity
                var availableQuery = @"
                SELECT 
                    L.Quantity - COALESCE(SUM(R.Quantity), 0) AS AvailableQuantity
                FROM Locations L
                LEFT JOIN Reservations R ON L.LocID = R.LocID AND R.Status IN ('Pending', 'Confirmed')
                WHERE L.LocID = @LocID
                GROUP BY L.Quantity";

                var availableQty = await connection.QueryFirstOrDefaultAsync<decimal?>(availableQuery, new { LocID = reservation.LocID }) ?? 0;

                if (availableQty < reservation.Quantity)
                    return BadRequest($"Insufficient quantity available. Available: {availableQty}, Requested: {reservation.Quantity}");

                // Create reservation
                var query = @"
                INSERT INTO Reservations (LocID, Quantity, JobDescription, Status, ReservationDate)
                VALUES (@LocID, @Quantity, @JobDescription, @Status, GETDATE());
                SELECT CAST(SCOPE_IDENTITY() as int)";

                var id = await connection.ExecuteScalarAsync<int>(query, reservation);
                reservation.ReservationID = id;

                return CreatedAtAction(nameof(GetReservation), new { id = reservation.ReservationID }, reservation);
            }
        }

        // GET: api/Inventory/Reservations
        [HttpGet("Reservations")]
        public async Task<IActionResult> GetReservations()
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                var query = @"
                SELECT 
                    R.ReservationID,
                    R.LocID,
                    L.Description AS LocationDescription,
                    M.MatID,
                    M.Description AS MaterialDescription,
                    R.Quantity,
                    R.JobDescription,
                    R.ReservationDate,
                    R.Status
                FROM Reservations R
                JOIN Locations L ON R.LocID = L.LocID
                LEFT JOIN Materials M ON L.MatID = M.MatID
                ORDER BY R.ReservationDate DESC";

                var reservations = await connection.QueryAsync<ReservationViewModel>(query);
                return Ok(reservations);
            }
        }

        // GET: api/Inventory/Reservations/{id}
        [HttpGet("Reservations/{id}")]
        public async Task<IActionResult> GetReservation(int id)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();
                var query = @"
                SELECT 
                    R.ReservationID,
                    R.LocID,
                    L.Description AS LocationDescription,
                    M.MatID,
                    M.Description AS MaterialDescription,
                    R.Quantity,
                    R.JobDescription,
                    R.ReservationDate,
                    R.Status
                FROM Reservations R
                JOIN Locations L ON R.LocID = L.LocID
                LEFT JOIN Materials M ON L.MatID = M.MatID
                WHERE R.ReservationID = @Id";

                var reservation = await connection.QueryFirstOrDefaultAsync<ReservationViewModel>(query, new { Id = id });

                if (reservation == null)
                    return NotFound();

                return Ok(reservation);
            }
        }

        // PUT: api/Inventory/Reservations/{id}/Status
        [HttpPut("Reservations/{id}/Status")]
        public async Task<IActionResult> UpdateReservationStatus(int id, [FromBody] ReservationStatusUpdate statusUpdate)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                await connection.OpenAsync();

                var query = @"
                UPDATE Reservations
                SET Status = @Status
                WHERE ReservationID = @ReservationID";

                var affected = await connection.ExecuteAsync(query, new { ReservationID = id, Status = statusUpdate.Status });

                if (affected == 0)
                    return NotFound();

                return NoContent();
            }
        }
    }
}