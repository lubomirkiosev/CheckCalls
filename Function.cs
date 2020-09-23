namespace CheckCalls1._1
{
    using Syncfusion.Drawing;
    using Syncfusion.XlsIO;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows;

    public class Function
    {
        public static async Task Main(string[] args)
        {

            var line = await GetValueFromTxtDocunment(args[0]);

            int col = await GetColumn(args[1].ToString());

            FileStream inputStream = new FileStream(args[1], FileMode.Open, FileAccess.ReadWrite);

            using (ExcelEngine SaveExcelEngine = new ExcelEngine())
            {
                IWorkbook workbook = SaveExcelEngine.Excel.Workbooks.Open(inputStream);

                int firstRow = 6;

                var sheet = workbook.ActiveSheet;

                await Task.Run(() =>
                {
                    while (sheet.GetValueRowCol(firstRow, 1).ToString() != string.Empty)
                    {
                        var targetData = sheet.GetValueRowCol(firstRow, 1).ToString();

                        if (line.Contains(targetData) && sheet.GetValueRowCol(firstRow, col).ToString() == string.Empty)
                        {
                            sheet.Range[firstRow, col].CellStyle.Color = Color.White;
                            sheet.SetValueRowCol("T", firstRow, col);
                            workbook.Version = ExcelVersion.Excel2016;
                            workbook.SaveAs(inputStream);
                        }
                        else if (sheet.GetValueRowCol(firstRow, col).ToString() == string.Empty)
                        {
                            sheet.Range[firstRow, col].CellStyle.Color = Color.Yellow;
                            sheet.Range[firstRow, col].BorderAround(ExcelLineStyle.Thick, Color.Black);
                            workbook.Version = ExcelVersion.Excel2016;
                            workbook.SaveAs(inputStream);
                        }
                        firstRow++;
                    }
                });

                inputStream.Dispose();
            }
        }

        private static async Task<int> GetColumn(string filePath)
        {
            int col = 0;
            await Task.Run(() =>
            {
                if (DateTime.Now.Hour >= 0 && DateTime.Now.Hour < 3)
                {
                    if (filePath.Contains("Неделя") || filePath.Contains("Събота"))
                        col++;
                    col += 7;
                }
                else if (DateTime.Now.Hour >= 3 && DateTime.Now.Hour < 5)
                {
                    if (filePath.Contains("Неделя") || filePath.Contains("Събота"))
                        col++;
                    col += 8;
                }
                else if (DateTime.Now.Hour >= 5 && DateTime.Now.Hour < 7)
                {
                    if (filePath.Contains("Неделя") || filePath.Contains("Събота"))
                        col++;
                    col += 9;
                }
                else if (DateTime.Now.Hour >= 7 && DateTime.Now.Hour < 11)
                {
                    col = 4;
                }
                else if (DateTime.Now.Hour >= 11 && DateTime.Now.Hour < 17)
                {
                    col += 5;
                }
                else if (DateTime.Now.Hour >= 17 && DateTime.Now.Hour < 19)
                {
                    if (filePath.Contains("Неделя") || filePath.Contains("Събота"))
                        col++;
                    col += 5;
                }
                else
                {
                    if (filePath.Contains("Неделя") || filePath.Contains("Събота"))
                        col++;
                    col += 6;
                }
            });
            return col;
        }

        private static async Task<List<string>> GetValueFromTxtDocunment(string inputData)
        {
            var line = new List<string>();
            var replaceElements = new List<string>();

            await Task.Run(() =>
            {
                string[] separatingStrings = { " ", ",", ".", ":", ";", ", ", "\n", "\r" };

                line = inputData.Split(separatingStrings, StringSplitOptions.RemoveEmptyEntries).ToList();

                for (int i = 0; i < line.Count; i++)
                {
                    var currentValue = line[i];

                    if (currentValue.Contains("0") && currentValue.Length > 4)
                    {
                        var newElement = new StringBuilder(line[i]);

                        var lastIndexWithZero = currentValue.LastIndexOf('0');
                        newElement[lastIndexWithZero] = '-';
                        replaceElements.Add(newElement.ToString());
                    }
                }

            });
            line.AddRange(replaceElements);
            return line;
        }
    }
}