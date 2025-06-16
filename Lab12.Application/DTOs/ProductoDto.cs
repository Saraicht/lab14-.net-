namespace Lab12.DTOs
{
    public class ProductoDto
    {
        public int ProductoId { get; set; }
        public string Nombre { get; set; } = null!;
        public string? Descripcion { get; set; }
        public decimal? Precio { get; set; }
        public int? Stock { get; set; }
        public string? Categoria { get; set; }
        public DateOnly? FechaIngreso { get; set; }
    }

}