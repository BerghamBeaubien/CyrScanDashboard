using CyrScanDashboard.Models;
using Dapper;
using Microsoft.Data.SqlClient;

namespace CyrScanDashboard.Services
{
    public class UserService : IUserService
    {
        private readonly string _connectionString;

        public UserService(IConfiguration configuration)
        {
            _connectionString = "Server=192.168.88.55,1433;Database=CyrScanDB;User Id=Serveur-CyrScan;Password=admin;TrustServerCertificate=True;";
        }

        public async Task<User> AuthenticateAsync(string username, string password)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            var user = await connection.QuerySingleOrDefaultAsync<User>(
                "SELECT * FROM Users WHERE Username = @Username",
                new { Username = username });

            if (user == null)
                return null;

            // Check password (using BCrypt)
            if (!BCrypt.Net.BCrypt.Verify(password, user.Password))
                return null;

            // Update last login date
            await connection.ExecuteAsync(
                "UPDATE Users SET LastLoginDate = @LastLoginDate WHERE Id = @Id",
                new { LastLoginDate = DateTime.UtcNow, Id = user.Id });

            return user;
        }

        public async Task<User> RegisterAsync(RegisterModel model, string role = "User")
        {
            // Validate if username already exists
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            var existingUser = await connection.QuerySingleOrDefaultAsync<User>(
                "SELECT * FROM Users WHERE Username = @Username",
                new { Username = model.Username });

            if (existingUser != null)
                throw new ApplicationException("Username is already taken");

            // Hash password
            string passwordHash = BCrypt.Net.BCrypt.HashPassword(model.Password);

            Console.WriteLine("Password Hashed = " + passwordHash + "User: " + model.Username );

            var newUser = new User
            {
                Username = model.Username,
                Password = passwordHash,
                Email = model.Email,
                Role = role,
                CreatedDate = DateTime.UtcNow
            };

            var id = await connection.QuerySingleAsync<int>(
                @"INSERT INTO Users (Username, Password, Email, Role, CreatedDate) 
                  VALUES (@Username, @Password, @Email, @Role, @CreatedDate);
                  SELECT CAST(SCOPE_IDENTITY() as int)",
                newUser);

            newUser.Id = id;
            return newUser;
        }

        public async Task<IEnumerable<User>> GetAllUsersAsync()
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            var users = await connection.QueryAsync<User>("SELECT * FROM Users");
            return users;
        }

        public async Task<User> GetUserByIdAsync(int id)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            var user = await connection.QuerySingleOrDefaultAsync<User>(
                "SELECT * FROM Users WHERE Id = @Id",
                new { Id = id });

            return user;
        }

        public async Task<bool> UpdateUserAsync(User user)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            // Check if password is being updated
            if (user.Password != null)
            {
                var result = await connection.ExecuteAsync(
                    @"UPDATE Users 
              SET Username = @Username, Email = @Email, Role = @Role, 
                  Password = @Password
              WHERE Id = @Id",
                    user);
                return result > 0;
            }
            else
            {
                // Update without changing the password
                var result = await connection.ExecuteAsync(
                    @"UPDATE Users 
              SET Username = @Username, Email = @Email, Role = @Role
              WHERE Id = @Id",
                    user);
                return result > 0;
            }
        }

        public async Task<bool> UpdateUserPasswordAsync(int userId, string newPassword)
        {
            string passwordHash = BCrypt.Net.BCrypt.HashPassword(newPassword);

            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            var result = await connection.ExecuteAsync(
                "UPDATE Users SET Password = @Password WHERE Id = @Id",
                new { Password = passwordHash, Id = userId });

            return result > 0;
        }

        public async Task<bool> DeleteUserAsync(int id)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            if (id == 1)
                throw new ApplicationException("Cannot delete the default admin user");

            var result = await connection.ExecuteAsync(
                "DELETE FROM Users WHERE Id = @Id",
                new { Id = id });

            return result > 0;
        }
    }
}