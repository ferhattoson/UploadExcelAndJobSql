using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using UploadExcel.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using UploadExcel.Models.ViewModels;
using Microsoft.EntityFrameworkCore;
using EFCore.BulkExtensions;
using UploadExcel.Services;

namespace UploadExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly ExcelService _excelService;

        public HomeController(ILogger<HomeController> logger, ExcelService excelService)
        {
            _logger = logger;
            _excelService = excelService;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Upload(IFormFile excelFile)
        {
            _excelService.UploadExcelFile(excelFile);

            return Json(new { message = "Dosya başarıyla yüklendi." });
        }

        [HttpPost]
        public JsonResult StartSQLJob()
        {
            

            return Json(new { message = "SQL Job başarıyla başlatıldı." });
        }

        [HttpPost]
        public IActionResult ShowData([FromForm] IFormFile excelFile)
        {
            List<Employer> employers = _excelService.ProcessExcelData(excelFile);

            return StatusCode(StatusCodes.Status200OK, employers);
        }

        [HttpPost]
        public IActionResult SendData([FromForm] IFormFile excelFile)
        {
            List<Employer> employers = _excelService.ProcessExcelData(excelFile);
            _excelService.SaveEmployers(employers);

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