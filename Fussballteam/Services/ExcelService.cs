using OfficeOpenXml;
using Fussballteam.Models;

namespace Fussballteam.Services
{    
    class ExcelService
    {
        private List<PlayerModel> players = new List<PlayerModel>();

        public List<PlayerModel> Reader(ExcelPackage package, String UserName)
        {
            ExcelWorksheet ws = package.Workbook.Worksheets[UserName];
            if (ws != null)
            {
                PlayerModel player;
                for (int row = 5; row <= 53; row++)
                {
                    player = new PlayerModel
                    {
                        Name = (ws.Cells[row, 1].Value ?? string.Empty).ToString().Trim(),
                        Position = (ws.Cells[row, 2].Value ?? string.Empty).ToString().Trim(),
                        Age = ((ws.Cells[row, 3].Value ?? string.Empty).ToString().Trim()),
                        Salary = ((ws.Cells[row, 5].Value ?? string.Empty).ToString().Trim()),
                    };

                    if (player.isValid())
                    {
                        players.Add(player);
                    }
                }                
            }
            else
            {
                throw new Exception("Excel sheet needs to be named after User email");
            }
            return players;
        }

        public decimal AverageSalary(ExcelPackage package, String UserName)
        {
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

        public string highestSalary(ExcelPackage package, String UserName)
        {
            ExcelWorksheet ws = package.Workbook.Worksheets[UserName];
            if (ws != null)
            {
                var max = (ws.Cells["F68"].Value);
                return Convert.ToString(max);
            }
            else
            {
                return "Error";
            }
        }

        public string lowestSalary(ExcelPackage package, String UserName)
        {
            ExcelWorksheet ws = package.Workbook.Worksheets[UserName];
            if (ws != null)
            {
                var min = (ws.Cells["F65"].Value);
                return Convert.ToString(min);
            }
            else
            {
                return "Error";
            }
        }

        public int FindFirstEmptyRow(ExcelPackage package, String UserName)
        {
            ExcelWorksheet ws = package.Workbook.Worksheets[UserName];
            int row;
            for (row = 5; row <= 53; row++)
                {
                    if (ws.Cells[row, 1].Value != null)
                    {
                        continue;
                    }
                    else break;
                }
            
            if ((row == 53) && (ws.Cells[53, 1].Value != null))
            {
                return -1;
            }
            else
            {
                return row;
            }
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
            else
            {
                throw new Exception("Excel sheet needs to be named after User email");
            }
        }

        public int FindRowOfPlayer(ExcelPackage package, String UserName, string name)
        {
            int row;
            ExcelWorksheet ws = package.Workbook.Worksheets[UserName];

            for (row = 5; row <= 53; row++)
                {
                    if ((ws.Cells[row, 1].Value ?? string.Empty).ToString().Trim() != name)
                    {
                        continue;
                    }
                    else break;
                }
            
            if ((row == 53) && (ws.Cells[row, 1].Value != name))
            {
                return -1;
            }
            else
            {
                return row;
            }
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

        public void Writer2(ExcelPackage package, int i, int ii)
        {
            ExcelWorksheet ws = package.Workbook.Worksheets["Inputs"];            
            ws.Cells["C3"].Value = i;
            ws.Cells["C4"].Value = ii;
            package.Save();            
        }

        public string Reader2(ExcelPackage package)
        {
            ExcelWorksheet ws = package.Workbook.Worksheets["Outputs"];
            var ii = (ws.Cells["C3"].Value);
            string i = (ii).ToString();
            package.Save();

            return i;                          
        }
    }
}


    
