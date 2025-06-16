namespace Lab12.Application.Interfaces
{
    public interface IExcelExportService
    {
        void GenerateExcelToDownloads();
        void ModifyExistingExcel(); // ← Agregado
        void GenerateExcelTable(); // ✅ nuevo
        
        void GenerateStyledExcel();



    }
}