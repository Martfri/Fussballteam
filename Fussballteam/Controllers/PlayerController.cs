using Fussballteam.Models;
using Microsoft.AspNetCore.Mvc;
using Fussballteam.Services;
using OfficeOpenXml;
using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;

namespace Fussballteam.Controllers
{
    public class PlayerController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private string fileinfo;
        private ExcelPackage package;
        private List<PlayerModel> players = new List<PlayerModel>();
        private ExcelService Excel = new ExcelService();
        private String? UserEmail;

        public PlayerController(ILogger<HomeController> logger)
        {
            _logger = logger;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            fileinfo = ".\\Models\\Fussballteam.xlsx";
            package = new ExcelPackage(path: this.fileinfo);
        }

        [Authorize]
        [HttpGet]
        public ActionResult DeletePlayer()
        {
            UserEmail = User.FindFirstValue(ClaimTypes.Email);
            players = Excel.Reader(package, UserEmail);
            return View(players);
        }

        [Authorize]
        [HttpGet]
        public ActionResult CreatePlayer()
        {
            PlayerModel player = new PlayerModel();
            return View(player);
        }

        [ValidateAntiForgeryToken]
        [Authorize]
        [HttpPost]
        public ActionResult CreatePlayer(string Name, string Position, string Age, string Salary)
        {
            UserEmail = User.FindFirstValue(ClaimTypes.Email);
            int row = Excel.FindFirstEmptyRow(package, UserEmail);

            if (row == -1)
            {
                throw new Exception("Player list is full");
            }
            else
            {
                Excel.Writer(package, UserEmail, Name, Position, Age, Salary, row);
            }

            return RedirectToAction(nameof(HomeController.Index), "Home");
        }

        [Authorize]
        [HttpPost]
        public bool Delete(string name)
        {
            UserEmail = User.FindFirstValue(ClaimTypes.Email);
            try
            {
                int row = Excel.FindRowOfPlayer(package, UserEmail, name);
                Excel.DeletePlayer(package, UserEmail, row);
                return true;
            }
            catch (System.Exception)
            {
                return false;
            }
        }

        [Authorize]
        [HttpGet]
        public ActionResult Update()
        {
            return View();
        }

        [ValidateAntiForgeryToken]
        [Authorize]
        [HttpPost]
        public ActionResult UpdatePlayer(PlayerModel player)
        {
            UserEmail = User.FindFirstValue(ClaimTypes.Email);
            players = Excel.Reader(package, UserEmail);
            int row = Excel.FindRowOfPlayer(package, UserEmail, player.Name);
            if (row == -1)
            {
                throw new Exception("Player does not exist");
            }
            else
            {
                Excel.Writer(package, UserEmail, player.Name, player.Position, player.Age, player.Salary, row);             
            }
            return RedirectToAction(nameof(HomeController.Index), "Home");
        }

        public class Foo
        {
            public int a { get; set; }
            public int b { get; set; }

        }

    
        [HttpPost]
        public string Testfunction([FromBody] Foo data)
        {
            var a = data.a;
            var b = data.b;
            Excel.Writer2(package, a, b);
            string res = Excel.Reader2(package);

            return res;
        }

    }
}

