using System.IO;
using ClosedXML.Application.Interfaces;
using ClosedXML.Excel;

namespace ClosedXML.Infrastructure.Services
{
    public class ExcelService : IExcelService
    {
        public byte[] CreateFirstExample()
        { 
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Hoja1");

                worksheet.Cell(1, 1).Value = "Nombre";
                worksheet.Cell(1, 2).Value = "Edad";
                worksheet.Cell(2, 1).Value = "Juan";
                worksheet.Cell(2, 2).Value = 28;
                
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }
        public void CreateFirstExampleLocal(string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Hoja1");
                worksheet.Cell(1, 1).Value = "Nombre";
                worksheet.Cell(1, 2).Value = "Edad";
                worksheet.Cell(2, 1).Value = "Juan";
                worksheet.Cell(2, 2).Value = 28;
                workbook.SaveAs(filePath);
            }
        }
        
        public void ModifyExampleLocal(string filePath, int newAge)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                worksheet.Cell(2, 2).Value = newAge;
                workbook.Save();
            }
        }
        
    }
}