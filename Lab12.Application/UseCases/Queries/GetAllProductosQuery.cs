using Lab12.Application.Interfaces;
using Lab12.DTOs;
using MediatR;

namespace Lab12.UseCases.Queries
{
    public class GetAllProductosQuery : IRequest<List<ProductoDto>> { }

    public class GetAllProductosQueryHandler : IRequestHandler<GetAllProductosQuery, List<ProductoDto>>
    {
        private readonly IProductoRepository _repository;

        public GetAllProductosQueryHandler(IProductoRepository repository)
        {
            _repository = repository;
        }

        public async Task<List<ProductoDto>> Handle(GetAllProductosQuery request, CancellationToken cancellationToken)
        {
            var productos = await _repository.GetAllAsync();

            return productos.Select(p => new ProductoDto
            {
                ProductoId = p.ProductoId,
                Nombre = p.Nombre,
                Descripcion = p.Descripcion,
                Precio = p.Precio,
                Stock = p.Stock,
                Categoria = p.Categoria,
                FechaIngreso = p.FechaIngreso
            }).ToList();
        }
    }
}