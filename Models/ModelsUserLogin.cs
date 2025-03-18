namespace CyrScanDashboard.Models
{
    public class User
    {
        public int Id { get; set; }
        public string Username { get; set; }
        public string Password { get; set; } // Will store hashed password
        public string Email { get; set; }
        public string Role { get; set; } // Admin or User
        public DateTime CreatedDate { get; set; }
        public DateTime? LastLoginDate { get; set; }
    }

    // UserUpdateModel class to handle updates
    public class UserUpdateModel
    {
        public int Id { get; set; }
        public string Username { get; set; }
        public string Email { get; set; }
        public string Password { get; set; } // Optional for updates
        public string Role { get; set; }
    }

    public class LoginModel
{
    public string Username { get; set; }
    public string Password { get; set; }
}

    public class RegisterModel
    {
        public string Username { get; set; }
        public string Password { get; set; }
        public string Email { get; set; }
        public string Role { get; set; } = "User"; // Default role
    }

    public class AuthResponse
    {
        public string Token { get; set; }
        public string Username { get; set; }
        public string Role { get; set; }
        public DateTime Expiration { get; set; }
    }
}