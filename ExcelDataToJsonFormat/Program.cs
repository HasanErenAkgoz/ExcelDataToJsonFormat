using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace ExcelDataToJsonFormat
{
    class Program
    {
        static void Main(string[] args)
        {
            String filePath = @"D:\demo\demo.xlsx";
            string dosya_yolu = @"D:\metinbelgesi.txt";

            try
            {
                IWorkbook workbook = null;
                FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                FileStream fs = new FileStream(dosya_yolu, FileMode.OpenOrCreate, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs);

                if (filePath.IndexOf(".xlsx") > 0)
                {
                    workbook = new XSSFWorkbook(fileStream);
                }
                else if (filePath.IndexOf(".xls") > 0)

                    workbook = new HSSFWorkbook(fileStream);
                Console.WriteLine(filePath.IndexOf(".xlsx"));
                Console.WriteLine(filePath.IndexOf(".xlsx"));
                ISheet sheet = workbook.GetSheetAt(0);
                if (sheet != null)
                {
                    int rowCount = sheet.LastRowNum;
                    for (int i = 0; i < rowCount; i++)
                    {
                        IRow curRow = sheet.GetRow(i);
                        var cellValueTitle = curRow.GetCell(0).StringCellValue.Trim();
                        var cellValueTr = curRow.GetCell(1).StringCellValue.Trim();
                        var cellValueEn = curRow.GetCell(2).StringCellValue.Trim();
                        var cellValueFr = curRow.GetCell(3).StringCellValue.Trim();
                        var cellValueIt = curRow.GetCell(4).StringCellValue.Trim();
                        var cellValueBg = curRow.GetCell(5).StringCellValue.Trim();
                        var cellValueRo = curRow.GetCell(6).StringCellValue.Trim();
                        var cellValueRu = curRow.GetCell(7).StringCellValue.Trim();
                        JsonClass jsonClass = new JsonClass()
                        {
                            label = cellValueTitle,
                            TrTR = cellValueTr,
                            EnUS = cellValueEn,
                            FrFR = cellValueFr,
                            ItIT = cellValueIt,
                            BgBG = cellValueBg,
                            RoRO = cellValueRo,
                            RuRU = cellValueRu,

                        };
                        string strResultJson = JsonConvert.SerializeObject(jsonClass);
                        File.WriteAllText(@"jsonClass.json", strResultJson);
                        sw.WriteLine(jsonClass.label + ": {\n"
                            + "\"tr-Tr\":" + "\"" + jsonClass.TrTR 
                            + "\",\n\"en-Us\":"+"\""+jsonClass.EnUS+ "\"," +
                            "\n\"fr-FR\":"+"\""+jsonClass.FrFR+ "\"," +
                            "\n\"it-IT\":" + "\"" + jsonClass.ItIT + "\"," +
                            "\n\"bg-BG\":" + "\"" + jsonClass.BgBG + "\",\n\"ro-RO\":" + "\"" + jsonClass.RoRO+ "\"," +
                            "\n\"ru-RU\":" + "\"" + jsonClass.RuRU + "\"\n},\n\n");
                            
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
    }
}

