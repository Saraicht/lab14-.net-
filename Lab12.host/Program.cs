using Lab12.Application.Interfaces;
using Lab12.Infrastructure.Repositories;
using Lab12.Infrastructure.Services;
using Microsoft.EntityFrameworkCore;
using ClassLibrary1.Entities;
using Lab12.UseCases.Queries;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddDbContext<Lab12dbContext>(options =>
    options.UseMySql(builder.Configuration.GetConnectionString("DefaultConnection"),
        ServerVersion.AutoDetect(builder.Configuration.GetConnectionString("DefaultConnection"))));

builder.Services.AddScoped<IReporteService, ReporteService>();
builder.Services.AddScoped<IProductoRepository, ProductoRepository>();
builder.Services.AddScoped<IReporteService, ReporteService>();
builder.Services.AddScoped<IExcelExportService, ExcelExportService>();

builder.Services.AddMediatR(cfg =>
    cfg.RegisterServicesFromAssembly(typeof(GetAllProductosQuery).Assembly));

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.MapControllers();
app.Run();