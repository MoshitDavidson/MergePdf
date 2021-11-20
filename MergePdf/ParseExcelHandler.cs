using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergePdf
{
    public class ParseExcelHandler
    {
        private string _inputFolder;
        public ParseExcelHandler(string inputFolder)
        {
            _inputFolder = inputFolder;       
        }
        public Dictionary<int, List<string>> ParseExcelFile(string fileName, int investorsCount)

        {
            Dictionary<int, List<string>> dic = new();
            Application app = new Application();
            Workbook wb = app.Workbooks.Open($"{_inputFolder}\\{fileName}");
            try
            {
                Worksheet ws = (Worksheet)wb.Sheets[1];

                for (int i = 1; i <= investorsCount; i++)
                {
                    var indexCell = (Microsoft.Office.Interop.Excel.Range)ws.Cells[i, 1];
                    int indexValue = int.Parse(indexCell.Value.ToString());
                    var nameCell = (Microsoft.Office.Interop.Excel.Range)ws.Cells[i, 2];
                    string nameValue = nameCell.Value.ToString().Trim();
                    if (dic.ContainsKey(indexValue))
                        dic[indexValue].Add(nameValue);

                    else
                        dic.Add(indexValue, new List<string>() { nameValue });

                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error to parse excel file {ex}");
            }

            finally
            {
                wb.Close();
                app.Quit();
            }
            return dic;
           
        }
    }
}
