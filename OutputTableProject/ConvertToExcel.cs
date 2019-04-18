using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.IO;

namespace OutputTableProject
{
    //Класс для получения ведомости в формате *.xlsx 
    public class ConvertToExcel
    {
        public struct PriceStruct
        {
            public string val;
        }
        //Метод для создания Excel файла с ведомостью
        public static void ConvertToXLS()
        {
            SqlCommand comand = new SqlCommand();
            decimal dec = 0.0M;
            string s = "";
            int id = comand.SelectInt("Id", "OutputTable", "Id != 0");
            List<string[]> arr = new List<string[]>();
                  
            for (int i = 1; i < id+1; i++) {
                var tempArr = new string[14];

                if (i != id)
                {
                    dec = comand.SelectInt("OrderNum", "OutputTable", "Id = " + i);
                    s = dec.ToString();
                    s = s.Replace(",", ".");
                    tempArr[0] = s;

                    s = comand.SelectStr("OtdelId", "OutputTable", "Id = " + i);
                    tempArr[1] = s;
                }

                s = comand.SelectStr("WorkName", "OutputTable", "Id = " + i);
                tempArr[2] = s;

                if (i != id)
                {
                    dec = comand.Select("Vol", "OutputTable", "Id = " + i);
                    s = dec.ToString();
                    s = s.Replace(",", ".");
                    tempArr[3] = s;

                    dec = comand.Select("Tiraj", "OutputTable", "Id = " + i);
                    s = dec.ToString();
                    s = s.Replace(",", ".");
                    tempArr[4] = s;
                }

                dec = comand.Select("CostOfDoneWork", "OutputTable", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr[5] = s;

                dec = comand.Select("PaperOfset65", "OutputTable", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr[6] = s;

                dec = comand.Select("PaperOfset80", "OutputTable", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr[7] = s;

                dec = comand.Select("PaperMag48", "OutputTable", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr[8] = s;

                dec = comand.Select("PaperMel200", "OutputTable", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr[9] = s;

                dec = comand.Select("PaperMel250", "OutputTable", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr[10] = s;

                dec = comand.Select("PaperMel115", "OutputTable", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr[11] = s;

                dec = comand.Select("PaperMelKart", "OutputTable", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr[12] = s;

                dec = comand.Select("ColorPaper", "OutputTable", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr[13] = s;

                arr.Add(tempArr);
            }

            var priceData = new List<PriceStruct>();       
            foreach (var item in arr)
            {       
                priceData.Add(new PriceStruct { val = item[0] });
                priceData.Add(new PriceStruct { val = item[1] });
                priceData.Add(new PriceStruct { val = item[2] });
                priceData.Add(new PriceStruct { val = item[3] });
                priceData.Add(new PriceStruct { val = item[4] });
                priceData.Add(new PriceStruct { val = item[5] });
                priceData.Add(new PriceStruct { val = item[6] });
                priceData.Add(new PriceStruct { val = item[7] });
                priceData.Add(new PriceStruct { val = item[8] });
                priceData.Add(new PriceStruct { val = item[9] });
                priceData.Add(new PriceStruct { val = item[10] });
                priceData.Add(new PriceStruct { val = item[11] });
                priceData.Add(new PriceStruct { val = item[12] });
                priceData.Add(new PriceStruct { val = item[13] });
            }

            var headData = new List<PriceStruct>();
                headData.Add(new PriceStruct { val = "№ Заказа" });
                headData.Add(new PriceStruct { val = "Отдел" });
                headData.Add(new PriceStruct { val = "           Наименование работ           " });
                headData.Add(new PriceStruct { val = "Объем(факт)" });
                headData.Add(new PriceStruct { val = "Тираж" });
                headData.Add(new PriceStruct { val = "Стоим. вып. раб." });
                headData.Add(new PriceStruct { val = "Офсет. 65гр." });
                headData.Add(new PriceStruct { val = "Офсет. 80гр." });
                headData.Add(new PriceStruct { val = "Газет. 48.8гр." });
                headData.Add(new PriceStruct { val = "Мел. 200гр." });
                headData.Add(new PriceStruct { val = "Мел. 250гр." });
                headData.Add(new PriceStruct { val = "Мел. 115гр." });
                headData.Add(new PriceStruct { val = "Мел. карт." });
                headData.Add(new PriceStruct { val = "Цвет. бумага" });

            //Создание Excel Документа и заполнение его данными
            using (var eP = new ExcelPackage())
            {
                
                eP.Workbook.Properties.Author = "GVC Soft";
                eP.Workbook.Properties.Title = "Ведомость о работе и расходах полиграфии за месяц";
                eP.Workbook.Properties.Company = "NSK GVC";

                var sheet = eP.Workbook.Worksheets.Add("Ведомость");

                sheet.Cells.Style.HorizontalAlignment = OfficeOpenXml
                .Style
                .ExcelHorizontalAlignment
                .Center;
                sheet.Cells.Style.VerticalAlignment = OfficeOpenXml
                .Style
                .ExcelVerticalAlignment
                .Center;

                var row = 13;
                var col = 1;

                //Шапка
                sheet.Cells[3, 11].Value = "\"Утверждаю\"";
                sheet.Cells[4, 11].Value = "Зам. начальника ГВЦ Нацстаткома";
                sheet.Cells[5, 11].Value = "______________________________";                
                sheet.Cells[6, 11].Value = "              \"           \"                     2019г.        ";
                sheet.Cells[6, 11].Style.Font.UnderLine = true;

                sheet.Cells[9, 7].Value = "Ведомость";
                sheet.Cells[9, 7].Style.Font.Bold = true;
                sheet.Cells[9, 7].Style.Font.Size = 16;

                sheet.Cells[10, 7].Value = "выполненных работ по ОПР ГВЦ Нацстаткома Кыргызской Республики";
                sheet.Cells[10, 7].Style.Font.Bold = true;
                sheet.Cells[10, 7].Style.Font.Size = 13;

                //Добавление заголовков
                foreach (var item in headData)
                {
                    sheet.Cells[row, col].Value = item.val;
                    sheet.Cells[row, col].Style.Font.Bold = true;
                    sheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].AutoFitColumns();
                    col++;
                }
                row++;
                col = 1;

                //Добавление данных
                int l = 0, k = 0 ;
                foreach (var item in priceData)
                {
                    k++;                                     
                    sheet.Cells[row, col].Value = item.val;
                    sheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    if (k < 3) { sheet.Cells[row, col].AutoFitColumns(); }
   
                    col++;
                    l++;
                    if (l % 14 == 0)
                    {
                        row++;
                        col = 1;
                    }               
                }
                
                //Сохраняем в файл
                var bin = eP.GetAsByteArray();
                File.WriteAllBytes(@"documents\Report.xlsx", bin);
            }

        }
    }
}
