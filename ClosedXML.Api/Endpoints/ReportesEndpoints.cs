using System.IO;
using ClosedXML.Application.Interfaces;
using Microsoft.AspNetCore.Builder; 
using Microsoft.AspNetCore.Http; 
using Microsoft.AspNetCore.Routing; 

namespace ClosedXML.Api.Endpoints
{
    public static class ReportesEndpoints
    {
        public static void MapReportesEndpoints(this IEndpointRouteBuilder app)
        {
            // Endpoint de descarga
            app.MapGet("/api/reportes/parte2-ejemplo", (IExcelService excelService) =>
            {
                var fileBytes = excelService.CreateFirstExample();
                string fileName = "archivo_ejemplo.xlsx";
                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                return Results.File(fileBytes, contentType, fileName);
            })
            .WithName("GetFirstExampleReport")
            .WithTags("Reportes (Laboratorio)");

            // Endpoint de guardado local
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

        }
    }
}