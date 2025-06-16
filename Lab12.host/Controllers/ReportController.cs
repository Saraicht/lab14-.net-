using Lab12.Application.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace Lab12.host.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ReportController : ControllerBase
    {
        private readonly IExcelExportService _excelService;

        public ReportController(IExcelExportService excelService)
        {
            _excelService = excelService;
        }

        [HttpGet("create-excel")]
        public IActionResult CreateExcel()
        {
            _excelService.GenerateExcelToDownloads();
            return Ok("✅ Archivo Excel generado correctamente en la carpeta Descargas.");
        }

        [HttpGet("modify-excel")]
        public IActionResult ModifyExcel()
        {
            _excelService.ModifyExistingExcel();
            return Ok("✏️ Archivo Excel modificado correctamente (celda B2 actualizada).");
        }
        
        [HttpGet("create-table")]
        public IActionResult CreateTable()
        {
            _excelService.GenerateExcelTable();
            return Ok("📊 Archivo con tabla Excel generado correctamente.");
        }
        
        [HttpGet("styled-excel")]
        public IActionResult GenerateStyledExcel()
        {
            _excelService.GenerateStyledExcel();
            return Ok("🎨 Archivo Excel con estilos generado correctamente.");
        }

    }
}