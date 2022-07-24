﻿using Fussballteam.Models;
using Microsoft.AspNetCore.Mvc;
using Fussballteam.Services;
using OfficeOpenXml;
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

            fileinfo = ".\\Models\\Fussballteam.xlsx";
            package = new ExcelPackage(path: this.fileinfo);
        }

        [HttpGet]
        public ActionResult DeletePlayer()
        {
            UserName = User.FindFirstValue(ClaimTypes.Email);
            players = Excel.Reader(package, UserName);
            return View(players);
        }

        [HttpGet]
        public ActionResult CreatePlayer()
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

            return RedirectToAction(nameof(HomeController.Index), "Home");
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

        [HttpGet]
        public ActionResult Update()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UpdatePlayer(PlayerModel player)
        {
            UserName = User.FindFirstValue(ClaimTypes.Email);
            players = Excel.Reader(package, UserName);
            int row = Excel.FindRowOfPlayer(package, UserName, player.Name);
            Excel.Writer(package, UserName, player.Name, player.Position, player.Age, player.Salary, row);
            return RedirectToAction(nameof(HomeController.Index), "Home");
        }

        public class Foo
        {
            public int a { get; set; }
            public int b { get; set; }

        }

        [HttpPost]
        public string Testfunction([FromBody]Foo data)
        {
            var a = data.a;
            var b = data.b;        
            Excel.Writer2(package, a, b);
            string res = Excel.Reader2(package);

            return res;
        }



    }
}

