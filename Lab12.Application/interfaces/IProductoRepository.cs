using System.Collections.Generic;
using System.Threading.Tasks;
using ClassLibrary1.Entities;

namespace Lab12.Application.Interfaces
{
    public interface IProductoRepository
    {
        Task<List<Producto>> GetAllAsync();
    }
}