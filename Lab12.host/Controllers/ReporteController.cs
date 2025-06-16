using Lab12.Application.Interfaces;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;

namespace Lab12.host.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ReporteController : ControllerBase
    {
        private readonly IReporteService _reporteService;

        public ReporteController(IReporteService reporteService)
        {
            _reporteService = reporteService;
        }

        [HttpGet("productos-excel")]
        public async Task<IActionResult> ExportarExcelProductos()
        {
            await _reporteService.ExportarProductosExcel();
            return Ok("✅ Reporte de productos generado correctamente.");
        }

        [HttpGet("productos-excel/categoria/{categoria}")]
        public async Task<IActionResult> ExportarPorCategoria(string categoria)
        {
            await _reporteService.ExportarProductosPorCategoria(categoria);
            return Ok($"✅ Reporte de productos por categoría '{categoria}' generado correctamente.");
        }
    }
}   