using System.IO;
using ClosedXML.Application.Interfaces;
using ClosedXML.Infrastructure.Services;

var builder = WebApplication.CreateBuilder(args);

// ==========================================
// 1. REGISTRO DE SERVICIOS (Dependency Injection)
// ==========================================
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddScoped<IExcelService, ExcelService>();


var app = builder.Build();

// ==========================================
// 2. CONFIGURACIÓN DEL PIPELINE HTTP
// ==========================================

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

// ==========================================
// 3. DEFINICIÓN DE ENDPOINTS
// ==========================================

app.MapGet("/api/reportes/parte2-ejemplo", (IExcelService excelService) =>
    {
        var fileBytes = excelService.CreateFirstExample();
        string fileName = "archivo_ejemplo.xlsx";
        string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        return Results.File(fileBytes, contentType, fileName);
    })
    .WithName("GetFirstExampleReport")
    .WithTags("Reportes (Laboratorio)");


app.MapGet("/api/reportes/parte2-guardar-local", (IExcelService excelService) =>
    {
        try
        {
            string folderPath = @"C:\Users\geanm\OneDrive\Desktop\6-ciclo\excel";
            string fileName = "archivo_creado_localmente.xlsx";
            string fullPath = Path.Combine(folderPath, fileName);
            Directory.CreateDirectory(folderPath);
            excelService.CreateFirstExampleLocal(fullPath);
            return Results.Ok($"¡Éxito! Archivo guardado en: {fullPath}");
        }
        catch (Exception ex)
        {
            return Results.Problem($"Error al guardar el archivo: {ex.Message}");
        }
    })
    .WithName("SaveFirstExampleReportLocally")
    .WithTags("Reportes (Laboratorio)");


app.Run();