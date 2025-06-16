using System;
using System.Collections.Generic;

namespace ClassLibrary1.Entities;

public partial class Producto
{
    public int ProductoId { get; set; }

    public string Nombre { get; set; } = null!;

    public string? Descripcion { get; set; }

    public decimal? Precio { get; set; }

    public int? Stock { get; set; }

    public string? Categoria { get; set; }

    public DateOnly? FechaIngreso { get; set; }
}
