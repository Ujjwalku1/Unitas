using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Unitas.Framework;
using Unitas.Service.Model;
using static Org.BouncyCastle.Math.EC.ECCurve;

namespace Unitas.Service.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly ExcelService _excelService;
        private readonly IWebHostEnvironment _env;
        private readonly ILogger<ExcelController> _logger;
        private const string FileBaseName = "Unitas.xlsx";
        private const int MaxBackupFiles = 10;
      


        public ExcelController(ExcelService excelService, IWebHostEnvironment env, ILogger<ExcelController> logger)
        {
            _excelService = excelService;
            _env = env;
            _logger = logger;         
        } 

        [HttpPost("update-excel")]
        public IActionResult UpdateAndReadExcel(List<RequestModel> model)
        {

            string TemplateFolder = Path.Combine(_env.WebRootPath, "Upload", "Templates");
            const string TempFolderName = "execute";
            const int FileRetentionDays = 5;

            string sourceFilePath = Path.Combine(TemplateFolder, FileBaseName);
            string tempDirectory = Path.Combine(TemplateFolder, TempFolderName);

            try
            {
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                    _logger.LogInformation("Created temp directory: {Directory}", tempDirectory);
                }

                string copiedFilePath = Path.Combine(tempDirectory, $"Unitas_{Guid.NewGuid()}.xlsx");
                System.IO.File.Copy(sourceFilePath, copiedFilePath, overwrite: true);
                _logger.LogInformation("Copied file to: {CopiedFilePath}", copiedFilePath);

                if (model?.Any() == true)
                {
                    foreach (var request in model)
                    {
                       // ExcelNPOIExample.UpdateAndRecalculate(copiedFilePath, request.sheet, request.cell, request.value);
                         ExcelAsposeExample.UpdateAndRecalculate(copiedFilePath, request.sheet, request.cell, request.value);
                        // _excelService.UpdateCell(copiedFilePath, request.sheet, request.cell, request.value);
                        _logger.LogInformation("Updated cell {Cell} in sheet {Sheet} with value '{Value}'", request.cell, request.sheet, request.value);
                    }
                }

               // var excelService = new ExcelNpoiService(_config);
                //var result = _ExcelNpoiService.ReadSectionData(copiedFilePath);

              var result = _excelService.ReadSectionData(copiedFilePath);
                _logger.LogInformation("Successfully read section data from: {CopiedFilePath}", copiedFilePath);

                CleanOldFiles(tempDirectory, FileRetentionDays);

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occurred while processing Excel changes.");
                return StatusCode(500, $"Internal Server Error: {ex.Message}");
            }
        }


        //[HttpGet("get-records")]
        //public IActionResult GetRecords(string FileName)
        //{
        //    string filePath = Path.Combine(_env.WebRootPath, "Upload", "Templates", "Unitas Loan Sizer.xlsx");

        //    var result = _excelService.ReadSectionData(filePath); 

        //    return Ok(result);

        //} 

        [HttpPost]
        [Route("upload")]
        public async Task<IActionResult> Upload(IFormFile? file)
        {

            if (file == null || file.Length == 0)
            {
                _logger.LogWarning("Empty file upload attempt.");
                return BadRequest("No file provided.");
            }

            var uploadFolder = Path.Combine(_env.WebRootPath, "Upload", "Templates");
            Directory.CreateDirectory(uploadFolder);
            var unitasFilePath = Path.Combine(uploadFolder, FileBaseName);
            try
            {
                if (System.IO.File.Exists(unitasFilePath))
                {
                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    var backupFileName = $"Unitas_{timestamp}.xlsx";
                    var backupFilePath = Path.Combine(uploadFolder, backupFileName);

                    System.IO.File.Move(unitasFilePath, backupFilePath);
                    _logger.LogInformation("Existing Unitas.xlsx backed up as {BackupFile}", backupFileName);
                }
                await using (var stream = new FileStream(unitasFilePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                _logger.LogInformation("New file saved as Unitas.xlsx");
                CleanupOldBackups(uploadFolder);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "File upload failed.");
                return StatusCode(500, "Internal server error during file upload.");
            }

            return Ok("File uploaded successfully as Unitas.xlsx.");

        }

        private void CleanupOldBackups(string uploadFolder)
        {
            var backupFiles = Directory
                .GetFiles(uploadFolder, "Unitas_*.xlsx")
                .Select(path => new FileInfo(path))
                .OrderByDescending(f => f.CreationTimeUtc)
                .ToList();

            if (backupFiles.Count <= MaxBackupFiles)
                return;

            foreach (var fileToDelete in backupFiles.Skip(MaxBackupFiles))
            {
                try
                {
                    fileToDelete.Delete();
                    _logger.LogInformation("Old backup deleted: {File}", fileToDelete.Name);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to delete old backup: {File}", fileToDelete.FullName);
                }
            }
        }
        private void CleanOldFiles(string directoryPath, int olderThanDays)
        {
            try
            {
                var cutoffDate = DateTime.Now.AddDays(-olderThanDays);
                var files = Directory.GetFiles(directoryPath);

                foreach (var filePath in files)
                {
                    var fileInfo = new FileInfo(filePath);
                    if (fileInfo.CreationTime < cutoffDate)
                    {
                        fileInfo.Delete();
                        _logger.LogInformation("Deleted old file: {FileName}", fileInfo.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to clean up old files in directory: {DirectoryPath}", directoryPath);
            }
        }
    }
}
