using Fussballteam.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Fussballteam.Services;
using OfficeOpenXml;
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
            fileinfo = ".\\Models\\Fussballteam.xlsx";
            package = new ExcelPackage(path: fileinfo);            
        }

        [HttpGet]
        public IActionResult Index(IFormCollection form)
        {
            UserName = User.FindFirstValue(ClaimTypes.Email);
            players = Excel.Reader(package, UserName);

            return View(players);
        }

        [HttpGet]
        public IActionResult Statistics()
        {         
            return View();
        }
        
        public IActionResult check(string button)
        {
            UserName = User.FindFirstValue(ClaimTypes.Email);
            var salaries = Excel.AverageSalary(package, UserName);
            var age = Excel.AverageAge(package, UserName);
            var highestSalary = Excel.highestSalary(package, UserName);
            var lowestSalary = Excel.lowestSalary(package, UserName);

            if (button == "Get Average Salary")
            {
                TempData["ButtonValue"] = (salaries).ToString();
            }

            else if (button == "Get Average Age")
            {
                TempData["ButtonValue2"] = (age).ToString();
            }

            else if (button == "Get Player with highest Salary")
            {
                TempData["ButtonValue3"] = highestSalary.ToString();
            }

            else if (button == "Get Player with lowest Salary")
            {
                TempData["ButtonValue4"] = lowestSalary.ToString();
            }

            else {
                TempData["ButtonValue"] = "Error";
                TempData["ButtonValue2"] = "Error";
                TempData["ButtonValue3"] = "Error";
                TempData["ButtonValue4"] = "Error";
            }

            return RedirectToAction("Statistics");
        }

        [HttpGet]
        public IActionResult Visualizations()
        {
            UserName = User.FindFirstValue(ClaimTypes.Email);
            players = Excel.Reader(package, UserName);
            DashboardModel dashboard = new DashboardModel();

            dashboard.Goalkeepers_count = players.Where(c => c.Position == "Goalkeeper").Count();
            dashboard.Centreback_count = players.Where(c => c.Position == "Centre-back").Count();
            dashboard.Leftback_count = players.Where(c => c.Position == "Left-back").Count();
            dashboard.Rightback_count = players.Where(c => c.Position == "Right-back").Count();
            dashboard.Defensive_Midfielder_count = players.Where(c => c.Position == "Defensive Midfielder").Count();
            dashboard.Central_Midfielder_count = players.Where(c => c.Position == "Central Midfielder").Count();
            dashboard.Playmaker_count = players.Where(c => c.Position == "Playmaker").Count();
            dashboard.Leftwinger_count = players.Where(c => c.Position == "Left-winger").Count();
            dashboard.Rightwinger_count = players.Where(c => c.Position == "Right-winger").Count();
            dashboard.Centre_forward_count = players.Where(c => c.Position == "Centre-forward").Count();

            dashboard.SalaryClass1_count = players.Where(c => Convert.ToInt32(c.Salary) >= 0 && Convert.ToInt32(c.Salary) <= 5000).Count();
            dashboard.SalaryClass2_count = players.Where(c => Convert.ToInt32(c.Salary) >= 5000 && Convert.ToInt32(c.Salary) <= 25000).Count();
            dashboard.SalaryClass3_count = players.Where(c => Convert.ToInt32(c.Salary) >= 25000 && Convert.ToInt32(c.Salary) <= 50000).Count();
            dashboard.SalaryClass4_count = players.Where(c => Convert.ToInt32(c.Salary) >= 50000 && Convert.ToInt32(c.Salary) <= 100000).Count();
            dashboard.SalaryClass5_count = players.Where(c => Convert.ToInt32(c.Salary) >= 100000).Count();

            return View(dashboard);           
        }
        
        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorView { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
