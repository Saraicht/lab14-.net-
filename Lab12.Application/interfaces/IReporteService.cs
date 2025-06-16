using System.Threading.Tasks;

namespace Lab12.Application.Interfaces
{
    public interface IReporteService
    {
        Task ExportarProductosExcel();
        Task ExportarProductosPorCategoria(string categoria);
    }
}