using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using UploadExcel.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using UploadExcel.Models.ViewModels;
using Microsoft.EntityFrameworkCore;
using EFCore.BulkExtensions;

namespace UploadExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly DBContext _dbocontext;
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger, DBContext context)
        {
            _dbocontext = context;
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ShowData([FromForm] IFormFile ExcelFile)
        {
            Stream stream = ExcelFile.OpenReadStream();

            IWorkbook MyExcel = null;

            if (Path.GetExtension(ExcelFile.FileName) == ".xlsx")
            {
                MyExcel = new XSSFWorkbook(stream);
            }
            else
            {
                MyExcel = new HSSFWorkbook(stream);
            }

            ISheet ExcelSheet = MyExcel.GetSheetAt(0);

            int rowCount = ExcelSheet.LastRowNum;

            List<Employe> list = new List<Employe>();

            for (int i = 1; i <= rowCount; i++)
            {
                IRow row = ExcelSheet.GetRow(i);

                list.Add(new Employe
                {
                    Name = row.GetCell(0).ToString(),
                    Surname = row.GetCell(1).ToString(),
                    Tel = row.GetCell(2).ToString(),
                    Mail = row.GetCell(3).ToString()
                });
            }

            return StatusCode(StatusCodes.Status200OK, list);
        }

        //[HttpPost]
        //public IActionResult SendData([FromForm] IFormFile ExcelFile)
        //{
        //    Stream stream = ExcelFile.OpenReadStream();

        //    IWorkbook MyExcel = null;

        //    if (Path.GetExtension(ExcelFile.FileName) == ".xlsx")
        //    {
        //        MyExcel = new XSSFWorkbook(stream);
        //    }
        //    else
        //    {
        //        MyExcel = new HSSFWorkbook(stream);
        //    }

        //    ISheet ExcelSheet = MyExcel.GetSheetAt(0);

        //    int rowCount = ExcelSheet.LastRowNum;

        //    List<Employe> list = new List<Employe>();

        //    for (int i = 1; i <= rowCount; i++)
        //    {
        //        IRow row = ExcelSheet.GetRow(i);

        //        list.Add(new Employe
        //        {
        //            Name = row.GetCell(0).ToString(),
        //            Surname = row.GetCell(1).ToString(),
        //            Tel = row.GetCell(2).ToString(),
        //            Mail = row.GetCell(3).ToString()
        //        });
        //    }

        //    _dbocontext.BulkInsert(list);

        //    return StatusCode(StatusCodes.Status200OK, new { message = "ok" });
        //}
        [HttpPost]
        public IActionResult SendData([FromForm] IFormFile ExcelFile)
        {
            Stream stream = ExcelFile.OpenReadStream();

            IWorkbook MyExcel = null;

            if (Path.GetExtension(ExcelFile.FileName) == ".xlsx")
            {
                MyExcel = new XSSFWorkbook(stream);
            }
            else
            {
                MyExcel = new HSSFWorkbook(stream);
            }

            ISheet ExcelSheet = MyExcel.GetSheetAt(0);

            int rowCount = ExcelSheet.LastRowNum;

            List<Employe> list = new List<Employe>();

            for (int i = 1; i <= rowCount; i++)
            {
                IRow row = ExcelSheet.GetRow(i);

                list.Add(new Employe
                {
                    Name = row.GetCell(0).ToString(),
                    Surname = row.GetCell(1).ToString(),
                    Tel = row.GetCell(2).ToString(),
                    Mail = row.GetCell(3).ToString()
                });
            }

            _dbocontext.AddRange(list); // Verileri bağlam nesnesine ekleyin
            _dbocontext.SaveChanges(); // Değişiklikleri veritabanına kaydedin

            return StatusCode(StatusCodes.Status200OK, new { message = "ok" });
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}