using System.Threading.Tasks;
using MH.Excel.Export;
using MH.Excel.Test.Models;
using Microsoft.AspNetCore.Mvc;

namespace MH.Excel.Test.Controllers
{
    public class HomeController : Controller
    {
        public async Task<IActionResult> Index()
        {
            var list = new TestClass().GetList();

            var excel = await ExportManager.ExportToXlsxAsync<TestClass, TestClass.SubClassTest>(list, "Test Excel Class");

            return File(excel.FileContents, excel.ContentType, excel.FileDownloadName);
        }
    }
}