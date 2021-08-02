using System.Net.Http;
using System.Threading.Tasks;
using MH.Excel.Export;
using MH.Excel.Test.Models;
using Microsoft.AspNetCore.Mvc;

namespace MH.Excel.Test.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public async Task<IActionResult> Excel1()
        {
            var list = new TestClass().GetList();

            var excel = await ExportManager.ExportToXlsxAsync(list, "Test Simple Excel");

            return File(excel.FileContents, excel.ContentType, excel.FileDownloadName);
        }

        public async Task<IActionResult> Excel2()
        {
            var list = new TestClass().GetList();

            var excel = await ExportManager.ExportToXlsxAsync<TestClass, SubClassTest>(list, "Test 2 Level Excel");

            return File(excel.FileContents, excel.ContentType, excel.FileDownloadName);
        }

        public async Task<IActionResult> Excel3()
        {
            var list = new TestClass().GetList();

            var excel = await ExportManager.ExportToXlsxAsync<TestClass, SubClassTest, SubClassTest2>(list, "Test 2 Level Excel");

            return File(excel.FileContents, excel.ContentType, excel.FileDownloadName);
        }
    }
}