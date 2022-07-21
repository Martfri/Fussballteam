using System;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Fussballteam.Models;

namespace Fussballteam.Services
{
    // TODO: Interface
    class ExcelService
    {
        //private string fileinfo;
        //private ExcelPackage package;
        private List<PlayerModel> players = new List<PlayerModel>();

        //public ExcelService()
        //{
        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        //    fileinfo = "C:\\Users\\frick\\OneDrive\\Dokumente\\Fussballteam.xlsx";
        //    package = new ExcelPackage(path: this.fileinfo);
        //}

        // TODO Rename
        public List<PlayerModel> Reader(ExcelPackage package, String UserName)
        {
            //ExcelWorksheet ws = this.package.Workbook.Worksheets["Arbeit"];
            ExcelWorksheet ws = package.Workbook.Worksheets[UserName];
            if (ws != null)
            {

                PlayerModel player;
                // TODO Konstanten!
                int row = 5;
                do
                {
                    player = new PlayerModel
                    {
                        Name = (ws.Cells[row, 1].Value ?? string.Empty).ToString().Trim(),
                        Position = (ws.Cells[row, 2].Value ?? string.Empty).ToString().Trim(),
                        Age = ((ws.Cells[row, 3].Value ?? string.Empty).ToString().Trim()),
                        Salary = ((ws.Cells[row, 5].Value ?? string.Empty).ToString().Trim()),
                    };

                    if (player.isValid())
                        players.Add(player);

                    row++;

                } while (player.isValid());

            }
            return players;

        }

        // TODO Rename
        public decimal AverageSalary(ExcelPackage package, String UserName)
        {
            //int avg = 0;
            //PropertyInfo[] properties = typeof(PlayerModel).GetType().GetProperties();

            //List<string> Salaries = Players.Select(o => o.Salary).ToList();

            //players.Average(w => Convert.ToDecimal(w.Salary));

            //return avg; 
            ExcelWorksheet ws = package.Workbook.Worksheets[UserName];
            if (ws != null)
            {
                var salary = (ws.Cells["F58"].Value);
                return Convert.ToDecimal(salary);
            }
            else
            {
                return 999999999;
            }

        }

        public decimal AverageAge(ExcelPackage package, String UserName)
        {

            ExcelWorksheet ws = package.Workbook.Worksheets[UserName];
            if (ws != null)
            {
                var age = (ws.Cells["C55"].Value);
                return Convert.ToDecimal(age);
            }
            else
            {
                return 999999999;
            }

        }

        //System.InvalidCastException: 'Unable to cast object of type '
        //    OfficeOpenXml.Drawing.Chart.ChartEx.ExcelHistogramChart' to type 'OfficeOpenXml.Drawing.ExcelPicture'.'

        //public String Charts(ExcelPackage package)
        //{
        //    ExcelWorksheet worksheet = package.Workbook.Worksheets["Charts"];
        //    List<ExcelDrawing> charts = new List<ExcelDrawing>();

        //    byte[] imageArray = worksheet.Drawings[0].As.Picture.Image.ImageBytes;

        //    string base64ImageRepresentation = Convert.ToBase64String(imageArray);

        //    var x = "ds";
        //    List<ExcelPicture> pic = new List<ExcelPicture>();
        //    pic.Append(x[0] as ExcelPicture);
        //    foreach (ExcelDrawing dw in x)
        //    {
        //        var z = dw as ExcelPicture;
        //        pic.Add(z);
        //    }
        //    return x;
        //}


    

        public int FindFirstEmptyRow(ExcelPackage package, String UserName)
        {
            //ExcelWorksheet ws = this.package.Workbook.Worksheets["Arbeit"];
            int row = -1;
            ExcelWorksheet ws = package.Workbook.Worksheets[UserName];
            if (ws != null)
            {
                row = 5;
                // TODO Konstanten!

                do
                {
                    {
                        row++;
                    };

                } while ((ws.Cells[row, 1].Value != null));
            }
            return row;
        }

        public void Writer(ExcelPackage package, String UserName, string name, string position, string age, string salary, int row)
        {            
            ExcelWorksheet ws = package.Workbook.Worksheets[UserName];
            if (ws != null)
            {                           
                ws.Cells[row, 1].Value = name;
                ws.Cells[row, 2].Value = position;
                ws.Cells[row, 3].Value = age;
                ws.Cells[row, 5].Value = salary;

                package.Save();
            }
        }

        public int FindRowOfPlayer(ExcelPackage package, String UserName, string name)
        {
            //ExcelWorksheet ws = this.package.Workbook.Worksheets["Arbeit"];
            int row = -1;
            ExcelWorksheet ws = package.Workbook.Worksheets[UserName];
            if (ws != null)
            {
                row = 5;
                // TODO Konstanten!
                do
                {
                    {
                        row++;

                    };

                } while ((ws.Cells[row, 1].Value ?? string.Empty).ToString().Trim() != name);

            }
            return row;
        }

        public bool DeletePlayer(ExcelPackage package, String UserName, int row)
        {

            ExcelWorksheet ws = package.Workbook.Worksheets[UserName];
            if (ws != null)
            {
                ws.Cells[row, 1].Value = null;
                ws.Cells[row, 2].Value = null;
                ws.Cells[row, 3].Value = null;
                ws.Cells[row, 5].Value = null;

                package.Save();

                return true;
            }
            return false;

        }
    }
}


    
