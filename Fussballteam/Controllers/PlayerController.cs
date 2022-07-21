using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Fussballteam.Models;
using Fussballteam.Services;

//namespace ReadExcel.Controllers
//{
//    public class PlayerController : Controller
//    {
//        [HttpGet]
//        public IActionResult Index()
//        {
//            return View(new List<PlayerModel>());
//        }

//        [HttpPost]
//        public IActionResult Index(IFormCollection form)
//        {
//            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
//            string fileinfo = "C:\\Users\\frick\\source\\repos\\Fussballteam\\Fussballteam";
//            ExcelPackage package = new ExcelPackage(path: fileinfo);
//            ExcelService Excel = new ExcelService();
//            List<PlayerModel> players = Excel.Reader();


//            return View(players);
//        }
//    }
//}



