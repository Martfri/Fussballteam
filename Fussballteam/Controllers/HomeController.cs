using Fussballteam.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Fussballteam.Services;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using System.Security.Claims;

namespace Fussballteam.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private string fileinfo;
        private ExcelPackage package;
        private List<PlayerModel> players = new List<PlayerModel>();
        private ExcelService Excel = new ExcelService();
        private String? UserName;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            fileinfo = "C:\\Users\\frick\\OneDrive\\Dokumente\\Fussballteam.xlsx";
            package = new ExcelPackage(path: this.fileinfo);
            
        }

        [HttpGet]
        public IActionResult Index(IFormCollection form)
        {
            UserName = User.FindFirstValue(ClaimTypes.Email);
            players = Excel.Reader(package, UserName);
            
            //var charts = Excel.Charts(package);
            //Console.WriteLine(charts);
            ////TempData["First"] = charts;
            //ViewBag.Message = charts.As.Picture;
            //ViewBag.Message = charts;          

             //Console.WriteLine(players.GroupBy(c => c.Age).Select(c => c.FirstOrDefault()).ToList());
            //Console.WriteLine(players.Select(c => c.Age).ToList());

            return View(players);
        }

        [HttpGet]
        public IActionResult Index1()
        {         
            return View();
        }

        
        public IActionResult check(string button)
        {
            UserName = User.FindFirstValue(ClaimTypes.Email);
            var salaries = Excel.AverageSalary(package, UserName);
            var age = Excel.AverageAge(package, UserName);

            if (button == "Get Average Salary")
            {
                TempData["ButtonValue"] = (salaries).ToString();
            }

            else if (button == "Get Average Age")
            {
                TempData["ButtonValue2"] = (age).ToString();
            }

            else {
                TempData["ButtonValue"] = "Error";
                TempData["ButtonValue2"] = "Error";
                    }

            return RedirectToAction("Index1");

        }


        [HttpGet]
        public IActionResult Index2()
        {
            UserName = User.FindFirstValue(ClaimTypes.Email);
            players = Excel.Reader(package, UserName);

            //TempData["1"] = players.GroupBy(c => c.Age).Select(c => c.FirstOrDefault()).ToList();
            //TempData["2"] = players.GroupBy(c => c.Age).Select(c => c.ToList().Count).ToArray();
            //ViewBag.Message = players.GroupBy(c => c.Age).Select(c => c.ToList().Count).ToArray();
            //Console.WriteLine(String.Join(" ", x));

            return View(players); 
        }
        
        public ActionResult Index3()
        {
            UserName = User.FindFirstValue(ClaimTypes.Email);
            players = Excel.Reader(package, UserName);
            return View(players);
        }

        public ActionResult Index4()
        {
            PlayerModel player = new PlayerModel();
            return View(player);
        }

        [HttpPost]
        public ActionResult CreatePlayer(string Name, string Position, string Age, string Salary)
        {            
            UserName = User.FindFirstValue(ClaimTypes.Email);
            int row = Excel.FindFirstEmptyRow(package, UserName);
            Excel.Writer(package, UserName, Name, Position, Age, Salary, row);
            
            return RedirectToAction("Index", "Home");
        }

        [HttpPost]
        public bool Delete(string name)
        {
            UserName = User.FindFirstValue(ClaimTypes.Email);
            try
            {
                int row = Excel.FindRowOfPlayer(package, UserName, name);
                Excel.DeletePlayer(package, UserName, row);
                return true;
            }
            catch (System.Exception)
            {
                return false;
            }
            
        }

        public ActionResult Update(string n)
        {
            UserName = User.FindFirstValue(ClaimTypes.Email);
            players = Excel.Reader(package, UserName);
            var x = players.First(x => x.Name == n);
            return View(x);
        }

        [HttpPost]
        public ActionResult UpdatePlayer(PlayerModel player)
        {
            int row = Excel.FindRowOfPlayer(package, UserName, player.Name);
            Excel.Writer(package, UserName, player.Name, player.Position, player.Age, player.Salary, row);
            return RedirectToAction("Index3", "Home");
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorView { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
