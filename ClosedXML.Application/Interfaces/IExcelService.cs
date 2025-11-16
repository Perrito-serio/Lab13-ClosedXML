namespace ClosedXML.Application.Interfaces
{
    public interface IExcelService
    {
        byte[] CreateFirstExample();
         
        void CreateFirstExampleLocal(string filePath);
    }
}