using CyrScanDashboard.Services;
var builder = WebApplication.CreateBuilder(args);

// Add this configuration to listen on all interfaces
builder.WebHost.ConfigureKestrel(serverOptions =>
{
    serverOptions.ListenAnyIP(5128, listenOptions =>
    {
        listenOptions.UseHttps("./certificate.pfx", "admin");
    });
});


// Your existing configuration
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll",
        builder =>
        {
            builder
                .AllowAnyOrigin()
                .AllowAnyMethod()
                .AllowAnyHeader();
        });
});

builder.Services.AddSingleton<ExcelValidationService>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseCors("AllowAll");
app.UseAuthorization();
app.MapControllers();
app.MapGet("/health", () => Results.Ok());
app.Run();