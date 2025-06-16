using Lab12.Application.Interfaces;
using Microsoft.EntityFrameworkCore;
using ClassLibrary1.Entities;

namespace Lab12.Infrastructure.Repositories
{
    public class ProductoRepository : IProductoRepository
    {
        private readonly Lab12dbContext _context;

        public ProductoRepository(Lab12dbContext context)
        {
            _context = context;
        }

        public async Task<List<Producto>> GetAllAsync()
        {
            return await _context.Productos.ToListAsync();
        }
    }
}