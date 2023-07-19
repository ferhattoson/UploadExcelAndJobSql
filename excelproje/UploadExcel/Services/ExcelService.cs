using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UploadExcel.Models;
using UploadExcel.Models.ViewModels;
using UploadExcel.Services;

namespace UploadExcel.Services
{

    public class ExcelService
    {
        private readonly DBContext _dbContext;

        public ExcelService(DBContext dbContext)
        {
            _dbContext = dbContext;
        }

        public void UploadExcelFile(IFormFile excelFile)
        {
            if (excelFile != null && excelFile.Length > 0)
            {
                string uploadFolder = @"C:\Users\ferhat.toson\source\repos\UploadExcel\excelproje";
                string fileName = Path.GetFileName(excelFile.FileName);
                string filePath = Path.Combine(uploadFolder, fileName);

                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    excelFile.CopyTo(fileStream);
                }

               

       
            }
            else
            {
                
            }
        }

        public List<Employer> ProcessExcelData(IFormFile excelFile)
        {
            List<Employer> employers = new List<Employer>();

            using (var stream = excelFile.OpenReadStream())
            {
                IWorkbook myExcel;

                if (Path.GetExtension(excelFile.FileName) == ".xlsx")
                {
                    myExcel = new XSSFWorkbook(stream);
                }
                else
                {
                    myExcel = new HSSFWorkbook(stream);
                }

                ISheet excelSheet = myExcel.GetSheetAt(0);
                int rowCount = excelSheet.LastRowNum;

                for (int i = 1; i <= rowCount; i++)
                {
                    IRow row = excelSheet.GetRow(i);

                    employers.Add(new Employer
                    {
                        Name = row.GetCell(0).ToString(),
                        Surname = row.GetCell(1).ToString(),
                        Tel = row.GetCell(2).ToString(),
                        Mail = row.GetCell(3).ToString()
                    });
                }
            }

            return employers;
        }

        public void SaveEmployers(List<Employer> employers)
        {
            _dbContext.AddRange(employers);
            _dbContext.SaveChanges();
        }
    }

}
