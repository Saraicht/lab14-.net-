using ClosedXML.Excel;
using Lab12.Application.Interfaces;
using System;
using System.IO;
using System.Linq;

namespace Lab12.Infrastructure.Services
{
    public class ExcelExportService : IExcelExportService
    {
        public void GenerateExcelToDownloads()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Hoja1");

            worksheet.Cell(1, 1).Value = "Nombre";
            worksheet.Cell(1, 2).Value = "Edad";
            worksheet.Cell(2, 1).Value = "Juan";
            worksheet.Cell(2, 2).Value = 28;

            var fileName = $"reporte_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            var downloadsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads",
                fileName
            );

            workbook.SaveAs(downloadsPath);
            Console.WriteLine($"📁 Archivo guardado en: {downloadsPath}");
        }

        public void ModifyExistingExcel()
        {
            var downloadsFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads"
            );

            var file = Directory.GetFiles(downloadsFolder, "reporte_*.xlsx")
                .OrderByDescending(File.GetCreationTime)
                .FirstOrDefault();

            if (file == null)
            {
                Console.WriteLine("⚠️ No se encontró ningún archivo para modificar.");
                return;
            }

            try
            {
                using var workbook = new XLWorkbook(file);
                var worksheet = workbook.Worksheet(1);
                worksheet.Cell(2, 2).Value = 30;
                workbook.Save();
                Console.WriteLine($"✏️ Archivo modificado correctamente: {file}");
            }
            catch (IOException ex)
            {
                Console.WriteLine($"❌ No se puede acceder al archivo: {ex.Message}");
                throw new IOException("El archivo está siendo usado por otro proceso. Por favor, ciérralo antes de modificarlo.");
            }
        }

        public void GenerateExcelTable()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Datos");

            worksheet.Cell(1, 1).Value = "ID";
            worksheet.Cell(1, 2).Value = "Nombre";
            worksheet.Cell(1, 3).Value = "Edad";
            worksheet.Cell(1, 4).Value = "Correo";
            worksheet.Cell(1, 5).Value = "Teléfono";
            worksheet.Cell(1, 6).Value = "Dirección";
            worksheet.Cell(1, 7).Value = "Estado Civil";
            worksheet.Cell(1, 8).Value = "Fecha Registro";

            var datos = new[]
            {
                new { Id = 1, Nombre = "Juan", Edad = 28, Correo = "juan@gmail.com", Telefono = "987654321", Direccion = "Av. Perú 123", Estado = "Soltero", Fecha = "2024-01-01" },
                new { Id = 2, Nombre = "María", Edad = 34, Correo = "maria@gmail.com", Telefono = "912345678", Direccion = "Jr. Lima 456", Estado = "Casada", Fecha = "2023-11-15" },
                new { Id = 3, Nombre = "Luis", Edad = 22, Correo = "luis@gmail.com", Telefono = "945612378", Direccion = "Calle 3 #110", Estado = "Soltero", Fecha = "2023-12-01" },
                new { Id = 4, Nombre = "Ana", Edad = 29, Correo = "ana@gmail.com", Telefono = "900123456", Direccion = "Av. Grau 200", Estado = "Casada", Fecha = "2022-06-12" },
                new { Id = 5, Nombre = "Carlos", Edad = 31, Correo = "carlos@gmail.com", Telefono = "988123456", Direccion = "MZ A LT 4", Estado = "Viudo", Fecha = "2022-08-20" },
                new { Id = 6, Nombre = "Lucía", Edad = 26, Correo = "lucia@gmail.com", Telefono = "977654321", Direccion = "Pasaje Sur 12", Estado = "Soltera", Fecha = "2021-05-10" },
                new { Id = 7, Nombre = "Pedro", Edad = 35, Correo = "pedro@gmail.com", Telefono = "922334455", Direccion = "Av. Mariscal 555", Estado = "Casado", Fecha = "2023-03-18" },
                new { Id = 8, Nombre = "Diana", Edad = 27, Correo = "diana@gmail.com", Telefono = "933112244", Direccion = "Jr. Los Olivos", Estado = "Soltera", Fecha = "2024-02-28" },
                new { Id = 9, Nombre = "Sofía", Edad = 33, Correo = "sofia@gmail.com", Telefono = "955667788", Direccion = "Calle San Juan", Estado = "Divorciada", Fecha = "2022-09-01" },
                new { Id = 10, Nombre = "Miguel", Edad = 30, Correo = "miguel@gmail.com", Telefono = "944556677", Direccion = "Av. Brasil 987", Estado = "Casado", Fecha = "2023-04-04" }
            };

            int row = 2;
            foreach (var d in datos)
            {
                worksheet.Cell(row, 1).Value = d.Id;
                worksheet.Cell(row, 2).Value = d.Nombre;
                worksheet.Cell(row, 3).Value = d.Edad;
                worksheet.Cell(row, 4).Value = d.Correo;
                worksheet.Cell(row, 5).Value = d.Telefono;
                worksheet.Cell(row, 6).Value = d.Direccion;
                worksheet.Cell(row, 7).Value = d.Estado;
                worksheet.Cell(row, 8).Value = d.Fecha;
                row++;
            }

            var range = worksheet.Range($"A1:H{row - 1}");
            range.CreateTable();

            var fileName = $"tabla_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            var downloadsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads",
                fileName
            );

            workbook.SaveAs(downloadsPath);
            Console.WriteLine($"📁 Tabla Excel guardada en: {downloadsPath}");
        }

        public void GenerateStyledExcel()
{
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Estilos");

            // Encabezados
            worksheet.Cell(1, 1).Value = "ID";
            worksheet.Cell(1, 2).Value = "Nombre";
            worksheet.Cell(1, 3).Value = "Edad";
            worksheet.Cell(1, 4).Value = "Correo";
            worksheet.Cell(1, 5).Value = "Teléfono";
            worksheet.Cell(1, 6).Value = "Dirección";
            worksheet.Cell(1, 7).Value = "Estado Civil";
            worksheet.Cell(1, 8).Value = "Fecha Registro";

            // Aplicar estilos a encabezados
            var headerRange = worksheet.Range("A1:H1");
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            // Datos
            var datos = new[]
            {
                new { Id = 1, Nombre = "Juan", Edad = 28, Correo = "juan@gmail.com", Telefono = "987654321", Direccion = "Av. Perú 123", Estado = "Soltero", Fecha = "2024-01-01" },
                new { Id = 2, Nombre = "María", Edad = 34, Correo = "maria@gmail.com", Telefono = "912345678", Direccion = "Jr. Lima 456", Estado = "Casada", Fecha = "2023-11-15" },
                new { Id = 3, Nombre = "Luis", Edad = 22, Correo = "luis@gmail.com", Telefono = "945612378", Direccion = "Calle 3 #110", Estado = "Soltero", Fecha = "2023-12-01" },
                new { Id = 4, Nombre = "Ana", Edad = 29, Correo = "ana@gmail.com", Telefono = "900123456", Direccion = "Av. Grau 200", Estado = "Casada", Fecha = "2022-06-12" },
                new { Id = 5, Nombre = "Carlos", Edad = 31, Correo = "carlos@gmail.com", Telefono = "988123456", Direccion = "MZ A LT 4", Estado = "Viudo", Fecha = "2022-08-20" },
                new { Id = 6, Nombre = "Lucía", Edad = 26, Correo = "lucia@gmail.com", Telefono = "977654321", Direccion = "Pasaje Sur 12", Estado = "Soltera", Fecha = "2021-05-10" },
                new { Id = 7, Nombre = "Pedro", Edad = 35, Correo = "pedro@gmail.com", Telefono = "922334455", Direccion = "Av. Mariscal 555", Estado = "Casado", Fecha = "2023-03-18" },
                new { Id = 8, Nombre = "Diana", Edad = 27, Correo = "diana@gmail.com", Telefono = "933112244", Direccion = "Jr. Los Olivos", Estado = "Soltera", Fecha = "2024-02-28" },
                new { Id = 9, Nombre = "Sofía", Edad = 33, Correo = "sofia@gmail.com", Telefono = "955667788", Direccion = "Calle San Juan", Estado = "Divorciada", Fecha = "2022-09-01" },
                new { Id = 10, Nombre = "Miguel", Edad = 30, Correo = "miguel@gmail.com", Telefono = "944556677", Direccion = "Av. Brasil 987", Estado = "Casado", Fecha = "2023-04-04" }
            };

            int row = 2;
            foreach (var d in datos)
            {
                worksheet.Cell(row, 1).Value = d.Id;
                worksheet.Cell(row, 2).Value = d.Nombre;
                worksheet.Cell(row, 3).Value = d.Edad;
                worksheet.Cell(row, 4).Value = d.Correo;
                worksheet.Cell(row, 5).Value = d.Telefono;
                worksheet.Cell(row, 6).Value = d.Direccion;
                worksheet.Cell(row, 7).Value = d.Estado;
                worksheet.Cell(row, 8).Value = d.Fecha;
                row++;
            }

            // Ajustar ancho de columnas automáticamente
            worksheet.Columns().AdjustToContents();

            // Guardar archivo
            var fileName = $"archivo_con_estilos_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            var downloadsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads",
                fileName
            );

            workbook.SaveAs(downloadsPath);
            Console.WriteLine($"🎨 Excel con estilos guardado en: {downloadsPath}");
        }

    }
}
