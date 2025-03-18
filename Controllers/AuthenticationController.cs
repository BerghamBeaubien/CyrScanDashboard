using CyrScanDashboard.Models;
using CyrScanDashboard.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;

namespace CyrScanDashboard.Controllers
{

    [Route("api/auth")]
    [ApiController]
    public class AuthController : ControllerBase
    {
        private readonly IUserService _userService;
        private readonly IConfiguration _configuration;

        public AuthController(IUserService userService, IConfiguration configuration)
        {
            _userService = userService;
            _configuration = configuration;
        }

        //// Temporary modification to AuthController
        //[HttpPost("register-initial-admin")]
        //[AllowAnonymous] // Allow this endpoint without authorization
        //public async Task<IActionResult> RegisterInitialAdmin([FromBody] RegisterModel model)
        //{
        //    // Check if any users exist
        //    var users = await _userService.GetAllUsersAsync();
        //    if (users.Any())
        //        return BadRequest("Initial admin already exists");

        //    try
        //    {
        //        var user = await _userService.RegisterAsync(model, "Admin");
        //        return Ok(new { message = "Initial admin registration successful", userId = user.Id });
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(new { message = ex.Message });
        //    }
        //}

        [HttpPost("login")]
        public async Task<IActionResult> Login([FromBody] LoginModel model)
        {
            if (!ModelState.IsValid)
            {
                Console.WriteLine("Login Attempted");
                return BadRequest(ModelState);
            }

            var user = await _userService.AuthenticateAsync(model.Username, model.Password);

            if (user == null)
                return Unauthorized(new { message = "Username or password is incorrect" });

            var token = GenerateJwtToken(user);

            return Ok(new AuthResponse
            {
                Token = token,
                Username = user.Username,
                Role = user.Role,
                Expiration = DateTime.UtcNow.AddDays(7)
            });
        }

        [HttpPost("register")]
        //[Authorize(Roles = "Admin")] // Only admins can register new users
        public async Task<IActionResult> Register([FromBody] RegisterModel model)
        {
            Console.WriteLine("Register Attempted");
            if (!ModelState.IsValid)
            {
                Console.WriteLine("ModelState Non Valid");
                return BadRequest(ModelState);
            }
            

            try
            {
                Console.WriteLine("Registering User");
                var user = await _userService.RegisterAsync(model, model.Role);
                return Ok(new { message = "Registration successful", userId = user.Id });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = ex.Message });
            }
        }

        private string GenerateJwtToken(User user)
        {
            var tokenHandler = new JwtSecurityTokenHandler();
            var key = Encoding.ASCII.GetBytes(_configuration["Jwt:Secret"]);

            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = new ClaimsIdentity(new[]
                {
                    new Claim(ClaimTypes.Name, user.Username),
                    new Claim(ClaimTypes.Role, user.Role),
                    new Claim("UserId", user.Id.ToString())
                }),
                Expires = DateTime.UtcNow.AddDays(7),
                SigningCredentials = new SigningCredentials(
                    new SymmetricSecurityKey(key),
                    SecurityAlgorithms.HmacSha256Signature)
            };

            var token = tokenHandler.CreateToken(tokenDescriptor);
            return tokenHandler.WriteToken(token);
        }
    }

    // User Management Controller (Admin Only)
    [Route("api/users")]
    [ApiController]
    public class UsersController : ControllerBase
    {
        private readonly IUserService _userService;

        public UsersController(IUserService userService)
        {
            _userService = userService;
        }

        [HttpGet]
        public async Task<IActionResult> GetAllUsers()
        {
            //Console.WriteLine("Tentative retrouver users");
            var users = await _userService.GetAllUsersAsync();
            //Console.WriteLine("Users retrouves");
            return Ok(users);
        }

        [HttpGet("{id}")]
        public async Task<IActionResult> GetUserById(int id)
        {
            var user = await _userService.GetUserByIdAsync(id);

            if (user == null)
                return NotFound();

            return Ok(user);
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateUser(int id, [FromBody] UserUpdateModel userUpdate)
        {
            if (id != userUpdate.Id)
                return BadRequest(new { message = "User ID mismatch" });

            var user = await _userService.GetUserByIdAsync(id);
            if (user == null)
                return NotFound();

            // Update the existing user with new values
            user.Username = userUpdate.Username;
            user.Email = userUpdate.Email;
            user.Role = userUpdate.Role;

            // Only update password if provided
            if (!string.IsNullOrEmpty(userUpdate.Password))
            {
                user.Password = BCrypt.Net.BCrypt.HashPassword(userUpdate.Password);
            }

            var result = await _userService.UpdateUserAsync(user);
            if (!result)
                return StatusCode(500, new { message = "Error updating user" });

            return Ok(new { message = "User updated successfully" });
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteUser(int id)
        {
            var user = await _userService.GetUserByIdAsync(id);

            if (user == null)
                return NotFound();

            var result = await _userService.DeleteUserAsync(id);

            if (!result)
                return StatusCode(500, new { message = "Error deleting user" });

            return Ok(new { message = "User deleted successfully" });
        }
    }
}