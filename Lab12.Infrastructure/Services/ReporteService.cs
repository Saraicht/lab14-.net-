using ClosedXML.Excel;
using Lab12.Application.Interfaces;
namespace Lab12.Infrastructure.Services
{
    public class ReporteService : IReporteService
    {
        private readonly IProductoRepository _productoRepository;

        public ReporteService(IProductoRepository productoRepository)
        {
            _productoRepository = productoRepository;
        }

        public async Task ExportarProductosExcel()
        {
            var productos = await _productoRepository.GetAllAsync();

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Productos");

            worksheet.Cell(1, 1).Value = "ID";
            worksheet.Cell(1, 2).Value = "Nombre";
            worksheet.Cell(1, 3).Value = "Descripción";
            worksheet.Cell(1, 4).Value = "Precio";
            worksheet.Cell(1, 5).Value = "Stock";
            worksheet.Cell(1, 6).Value = "Categoría";
            worksheet.Cell(1, 7).Value = "Fecha Ingreso";

            int row = 2;
            foreach (var p in productos)
            {
                worksheet.Cell(row, 1).Value = p.ProductoId;
                worksheet.Cell(row, 2).Value = p.Nombre;
                worksheet.Cell(row, 3).Value = p.Descripcion ?? "N/A";
                worksheet.Cell(row, 4).Value = p.Precio ?? 0;
                worksheet.Cell(row, 5).Value = p.Stock ?? 0;
                worksheet.Cell(row, 6).Value = p.Categoria ?? "N/A";
                worksheet.Cell(row, 7).Value = p.FechaIngreso?.ToString("yyyy-MM-dd") ?? "Sin fecha";
                row++;
            }

            worksheet.Columns().AdjustToContents();

            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads",
                $"reporte_productos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            workbook.SaveAs(path);
            Console.WriteLine($"✅ Reporte generado en: {path}");
        }

        public async Task ExportarProductosPorCategoria(string categoria)
        {
            var productos = await _productoRepository.GetAllAsync();
            var filtrados = productos.Where(p => p.Categoria == categoria).ToList();

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add($"Categoría - {categoria}");

            worksheet.Cell(1, 1).Value = "ID";
            worksheet.Cell(1, 2).Value = "Nombre";
            worksheet.Cell(1, 3).Value = "Descripción";
            worksheet.Cell(1, 4).Value = "Precio";
            worksheet.Cell(1, 5).Value = "Stock";
            worksheet.Cell(1, 6).Value = "Categoría";
            worksheet.Cell(1, 7).Value = "Fecha Ingreso";

            int row = 2;
            foreach (var p in filtrados)
            {
                worksheet.Cell(row, 1).Value = p.ProductoId;
                worksheet.Cell(row, 2).Value = p.Nombre;    
                worksheet.Cell(row, 3).Value = p.Descripcion ?? "N/A";
                worksheet.Cell(row, 4).Value = p.Precio ?? 0;
                worksheet.Cell(row, 5).Value = p.Stock ?? 0;
                worksheet.Cell(row, 6).Value = p.Categoria ?? "N/A";
                worksheet.Cell(row, 7).Value = p.FechaIngreso?.ToString("yyyy-MM-dd") ?? "Sin fecha";
                row++;
            }

            worksheet.Columns().AdjustToContents();

            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads",
                $"reporte_categoria_{categoria}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            workbook.SaveAs(path);
            Console.WriteLine($"✅ Reporte por categoría generado en: {path}");
        }
    }
}
