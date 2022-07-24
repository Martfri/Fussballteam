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
    public class PlayerController : Controller


    {
        private readonly ILogger<HomeController> _logger;
        private string fileinfo;
        private ExcelPackage package;
        private List<PlayerModel> players = new List<PlayerModel>();
        private ExcelService Excel = new ExcelService();
        private String? UserName;

        public PlayerController(ILogger<HomeController> logger)
        {
            _logger = logger;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            fileinfo = "C:\\Users\\frick\\OneDrive\\Dokumente\\Fussballteam.xlsx";
            package = new ExcelPackage(path: this.fileinfo);

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
    }
}
