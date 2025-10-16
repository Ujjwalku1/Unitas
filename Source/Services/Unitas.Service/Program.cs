using Serilog;
using Unitas.Framework;

var builder = WebApplication.CreateBuilder(args);

// Set log file path inside wwwroot
var logFilePath = Path.Combine(builder.Environment.WebRootPath, "logs", "log-.txt");

Log.Logger = new LoggerConfiguration()
    .WriteTo.File(logFilePath, rollingInterval: RollingInterval.Day)
    .CreateLogger();

builder.Host.UseSerilog();


// Add services to the container.
 
builder.Services.AddControllers();

builder.Services.AddScoped<ExcelService>();
builder.Services.AddScoped<ExcelServiceManager>();


// Learn more about configuring OpenAPI at https://aka.ms/aspnet/openapi
builder.Services.AddOpenApi();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

app.UseAuthorization();

app.MapControllers();
app.UseCors(
       options => options
        .AllowAnyOrigin()
        .AllowAnyMethod()
        .AllowAnyHeader()
      );
app.Run();
