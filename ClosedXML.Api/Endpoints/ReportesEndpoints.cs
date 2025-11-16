using System.IO;
using ClosedXML.Application.Interfaces;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;
using Microsoft.AspNetCore.Mvc; 

namespace ClosedXML.Api.Endpoints
{
    public static class ReportesEndpoints
    {
        public static void MapReportesEndpoints(this IEndpointRouteBuilder app)
        {
            // 1. Endpoint de descarga (Parte 2)
            app.MapGet("/api/reportes/parte2-ejemplo", (IExcelService excelService) =>
            {
                var fileBytes = excelService.CreateFirstExample();
                string fileName = "archivo_ejemplo.xlsx";
                string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                return Results.File(fileBytes, contentType, fileName);
            })
            .WithName("GetFirstExampleReport")
            .WithTags("Reportes (Laboratorio)");

            // 2. Endpoint de guardado local (Parte 2)
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

            // 3. Endpoint de Modificación
            app.MapPut("/api/reportes/parte3-modificar", (IExcelService excelService, [FromQuery] int nuevaEdad = 30) =>
            {
                try
                {
                    string folderPath = @"C:\Users\geanm\OneDrive\Desktop\6-ciclo\excel";
                    string fileName = "archivo_creado_localmente.xlsx";
                    string fullPath = Path.Combine(folderPath, fileName);

                    if (!File.Exists(fullPath))
                    {
                        return Results.NotFound($"El archivo no existe. Ejecuta primero el endpoint de 'guardar-local'.");
                    }

                    excelService.ModifyExampleLocal(fullPath, nuevaEdad);

                    return Results.Ok($"¡Éxito! La edad de Juan ha sido actualizada a {nuevaEdad} en el archivo: {fileName}");
                }
                catch (Exception ex)
                {
                    return Results.Problem($"Error al modificar el archivo: {ex.Message}");
                }
            })
            .WithName("ModifyExampleReportLocally")
            .WithTags("Reportes (Laboratorio)");
            
                
            // 4. Endpoint de Tablas
            app.MapGet("/api/reportes/parte4-tabla", (IExcelService excelService) =>
                {
                    try
                    {
                        string folderPath = @"C:\Users\geanm\OneDrive\Desktop\6-ciclo\excel";
                        string fileName = "tabla_ejemplo.xlsx"; 
                        string fullPath = Path.Combine(folderPath, fileName);

                        Directory.CreateDirectory(folderPath);

                        // Llamamos al nuevo servicio
                        excelService.CreateExampleWithTable(fullPath);

                        return Results.Ok($"¡Éxito! Archivo con tabla guardado en: {fullPath}");
                    }
                    catch (Exception ex)
                    {
                        return Results.Problem($"Error al crear la tabla: {ex.Message}");
                    }
                })
                .WithName("CreateExampleWithTable")
                .WithTags("Reportes (Laboratorio)");
            
            // 5. Endopoint de estilos
            app.MapGet("/api/reportes/parte5-estilos", (IExcelService excelService) =>
                {
                    try
                    {
                        string folderPath = @"C:\Users\geanm\OneDrive\Desktop\6-ciclo\excel";
                        string fileName = "archivo_con_estilos.xlsx";
                        string fullPath = Path.Combine(folderPath, fileName);

                        Directory.CreateDirectory(folderPath);

                        excelService.CreateExampleWithStyles(fullPath);

                        return Results.Ok($"¡Éxito! Archivo con estilos guardado en: {fullPath}");
                    }
                    catch (Exception ex)
                    {
                        return Results.Problem($"Error al crear el archivo con estilos: {ex.Message}");
                    }
                })
                .WithName("CreateExampleWithStyles")
                .WithTags("Reportes (Laboratorio)");
        }
    }
}