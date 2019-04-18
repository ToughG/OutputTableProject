using OfficeOpenXml;
using OfficeOpenXml.Style;
using OutputTableProject;
using System.Collections.Generic;
using System.IO;

namespace OutputTableCroject
{
    //Класс для получения расчётных таблиц в формате *.xlsx
    public class GetCalcTables
    {      
        public struct PriceStruct
        {
            public string val;
        }
        //Метод для создания Excel файла с расчётными таблицами
        public static void ConvertCalcTables()
        {           
            SqlCommand comand = new SqlCommand();
            decimal dec = 0.0M;
            string s = "";
            int n;
            int id1 = comand.SelectInt("Id", "InputTable", "Id != 0");
            List<string[]> arr = new List<string[]>();

            List<string[]> arr1 = new List<string[]>();
            List<string[]> arr2 = new List<string[]>();
            List<string[]> arr3 = new List<string[]>();
            List<string[]> arr4 = new List<string[]>();
            List<string[]> arr5 = new List<string[]>();

            //-------------------------------Входная таблица
            var tempArr = new string[18];

            n = comand.SelectInt("OrderNum", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[0] = s;

            n = comand.SelectInt("OtdelId", "InputTable", "Id = " + id1);
            s = comand.SelectStr("Name", "Otdel", "Id = " + n);
            tempArr[1] = s;

            s = comand.SelectStr("WorkName", "InputTable", "Id = " + id1);
            tempArr[2] = s;

            s = comand.SelectStr("Format", "InputTable", "Id = " + id1);
            tempArr[3] = s;
       
            n = comand.SelectInt("Tiraj", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[4] = s;

            n = comand.SelectInt("ObFact", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[5] = s;

            n = comand.SelectInt("ColPageA4", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[6] = s;

            n = comand.SelectInt("ColPageA3", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[7] = s;

            n = comand.SelectInt("Paper65A3", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[8] = s;

            n = comand.SelectInt("Paper80A3", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[9] = s;

            n = comand.SelectInt("Paper120A3", "InputTable", "Id = " + id1);
            s = n.ToString();      
            tempArr[10] = s;

            n = comand.SelectInt("PaperMagA3", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[11] = s;

            n = comand.SelectInt("PaperMel200", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[12] = s;

            n = comand.SelectInt("PaperMel220", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[13] = s;

            n = comand.SelectInt("PaperMel115", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[14] = s;

            n = comand.SelectInt("PaperMelKart", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[15] = s;

            n = comand.SelectInt("ColPage", "InputTable", "Id = " + id1);
            s = n.ToString();
            tempArr[16] = s;

            s = comand.SelectStr("Date", "InputTable", "Id = " + id1);
            tempArr[17] = s;

            arr.Add(tempArr);

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
                priceData.Add(new PriceStruct { val = item[14] });
                priceData.Add(new PriceStruct { val = item[15] });
                priceData.Add(new PriceStruct { val = item[16] });
                priceData.Add(new PriceStruct { val = item[17] });
            }

            var headData = new List<PriceStruct>();
            headData.Add(new PriceStruct { val = "№ Заказа :" });
            headData.Add(new PriceStruct { val = "Отдел:" });
            headData.Add(new PriceStruct { val = "Наименование :" });
            headData.Add(new PriceStruct { val = "Формат :" });
            headData.Add(new PriceStruct { val = "Тираж :" });
            headData.Add(new PriceStruct { val = "Объём факт. :" });
            headData.Add(new PriceStruct { val = "Объём печ. лист. :" });
            headData.Add(new PriceStruct { val = "Цвет. стр. А4 :" });
            headData.Add(new PriceStruct { val = "Цвет. стр. А3 :" });
            headData.Add(new PriceStruct { val = "Листов А3 80гр. :" });
            headData.Add(new PriceStruct { val = "Листов А3 120гр. :" });
            headData.Add(new PriceStruct { val = "Листов газет. А3 :" });
            headData.Add(new PriceStruct { val = "Листов мел. 200г :" });
            headData.Add(new PriceStruct { val = "Листов мел. 220г :" });
            headData.Add(new PriceStruct { val = "Листов мел. 115г :" });
            headData.Add(new PriceStruct { val = "Листов мел. кар. :" });
            headData.Add(new PriceStruct { val = "Листов цвет. бум. :" });
            headData.Add(new PriceStruct { val = "Дата" });


            //--------------------------------------Печать на офсетных 
            for (int i = 1; i < 5; i++)
            {
                var tempArr1 = new string[5];

                s = comand.SelectStr("Name", "PrintOnOfset", "Id = " + i);
                tempArr1[0] = s;

                dec = comand.Select("Price", "PrintOnOfset", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr1[1] = s;

                dec = comand.Select("Amount", "PrintOnOfset", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr1[2] = s;

                dec = comand.Select("Sum", "PrintOnOfset", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr1[3] = s;

                arr1.Add(tempArr1);
            }

            var table1Data = new List<PriceStruct>();
            foreach (var item in arr1)
            {
                table1Data.Add(new PriceStruct { val = item[0] });
                table1Data.Add(new PriceStruct { val = item[1] });
                table1Data.Add(new PriceStruct { val = item[2] });
                table1Data.Add(new PriceStruct { val = item[3] });
            }

            var table1Head = new List<PriceStruct>();
            table1Head.Add(new PriceStruct { val = "Наименование" });
            table1Head.Add(new PriceStruct { val = "Цена(сом)" });
            table1Head.Add(new PriceStruct { val = "Количество" });
            table1Head.Add(new PriceStruct { val = "Сумма" });

            //------------------------------------Тиражирование на ризографе
            for (int i = 1; i < 3; i++)
            {
                var tempArr2 = new string[5];

                s = comand.SelectStr("Name", "TirajOnRizograph", "Id = " + i);
                tempArr2[0] = s;

                dec = comand.Select("Amount", "TirajOnRizograph", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr2[1] = s;

                dec = comand.Select("Price", "TirajOnRizograph", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr2[2] = s;

                dec = comand.Select("Cost", "TirajOnRizograph", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr2[3] = s;

                arr2.Add(tempArr2);
            }

            var table2Data = new List<PriceStruct>();
            foreach (var item in arr2)
            {
                table2Data.Add(new PriceStruct { val = item[0] });
                table2Data.Add(new PriceStruct { val = item[1] });
                table2Data.Add(new PriceStruct { val = item[2] });
                table2Data.Add(new PriceStruct { val = item[3] });
            }

            var table2Head = new List<PriceStruct>();
            table2Head.Add(new PriceStruct { val = "Наименование" });
            table2Head.Add(new PriceStruct { val = "Количество" });
            table2Head.Add(new PriceStruct { val = "Цена(сом)" });
            table2Head.Add(new PriceStruct { val = "Стоимость(сом)" });

            //-------------------------------Тиражирование на цветном принтере
            for (int i = 1; i < 5; i++)
            {
                var tempArr3 = new string[6];

                s = comand.SelectStr("Name", "TirajOnColPrint", "Id = " + i);
                tempArr3[0] = s;

                dec = comand.Select("Vol", "TirajOnColPrint", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr3[1] = s;

                dec = comand.Select("Tiraj", "TirajOnColPrint", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr3[2] = s;

                dec = comand.Select("Price", "TirajOnColPrint", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr3[3] = s;

                dec = comand.Select("Cost", "TirajOnColPrint", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr3[4] = s;

                arr3.Add(tempArr3);
            }

            var table3Data = new List<PriceStruct>();
            foreach (var item in arr3)
            {
                table3Data.Add(new PriceStruct { val = item[0] });
                table3Data.Add(new PriceStruct { val = item[1] });
                table3Data.Add(new PriceStruct { val = item[2] });
                table3Data.Add(new PriceStruct { val = item[3] });
                table3Data.Add(new PriceStruct { val = item[4] });
            }

            var table3Head = new List<PriceStruct>();
            table3Head.Add(new PriceStruct { val = "Наименование" });
            table3Head.Add(new PriceStruct { val = "Объём" });
            table3Head.Add(new PriceStruct { val = "Тираж" });
            table3Head.Add(new PriceStruct { val = "Цена(сом)" });
            table3Head.Add(new PriceStruct { val = "Стоимость(сом)" });

            //---------------------------Тиражирование на ксероксе
            for (int i = 1; i < 5; i++)
            {
                var tempArr4 = new string[6];

                s = comand.SelectStr("Name", "TirajOnKseroks", "Id = " + i);
                tempArr4[0] = s;

                if (i < 3)
                {
                    dec = comand.Select("Vol", "TirajOnKseroks", "Id = " + i);
                    s = dec.ToString();
                    s = s.Replace(",", ".");
                    tempArr4[1] = s;
                }

                dec = comand.Select("Tiraj", "TirajOnKseroks", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr4[2] = s;

                dec = comand.Select("Price", "TirajOnKseroks", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr4[3] = s;

                dec = comand.Select("Cost", "TirajOnKseroks", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr4[4] = s;

                arr4.Add(tempArr4);
            }

            var table4Data = new List<PriceStruct>();
            foreach (var item in arr4)
            {
                table4Data.Add(new PriceStruct { val = item[0] });
                table4Data.Add(new PriceStruct { val = item[1] });
                table4Data.Add(new PriceStruct { val = item[2] });
                table4Data.Add(new PriceStruct { val = item[3] });
                table4Data.Add(new PriceStruct { val = item[4] });
            }

            var table4Head = new List<PriceStruct>();
            table4Head.Add(new PriceStruct { val = "Наименование" });
            table4Head.Add(new PriceStruct { val = "Объём" });
            table4Head.Add(new PriceStruct { val = "Тираж" });
            table4Head.Add(new PriceStruct { val = "Цена(сом)" });
            table4Head.Add(new PriceStruct { val = "Стоимость(сом)" });

            //-----------------------------------Расход бумаги
            for (int i = 1; i < 10; i++)
            {
                var tempArr5 = new string[7];

                s = comand.SelectStr("Name", "PaperExpense", "Id = " + i);
                tempArr5[0] = s;

                dec = comand.Select("ToPrint", "PaperExpense", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr5[1] = s;
               
                dec = comand.Select("ToPrilad", "PaperExpense", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr5[2] = s;

                dec = comand.Select("AmountPages", "PaperExpense", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr5[3] = s;

                if (i < 5)
                {
                    dec = comand.Select("Sum", "PaperExpense", "Id = " + i);
                    s = dec.ToString();
                    s = s.Replace(",", ".");
                    tempArr5[4] = s;
                }

                dec = comand.Select("Price", "PaperExpense", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr5[5] = s;

                dec = comand.Select("Cost", "PaperExpense", "Id = " + i);
                s = dec.ToString();
                s = s.Replace(",", ".");
                tempArr5[6] = s;

                arr5.Add(tempArr5);
            }

            var table5Data = new List<PriceStruct>();
            foreach (var item in arr5)
            {
                table5Data.Add(new PriceStruct { val = item[0] });
                table5Data.Add(new PriceStruct { val = item[1] });
                table5Data.Add(new PriceStruct { val = item[2] });
                table5Data.Add(new PriceStruct { val = item[3] });
                table5Data.Add(new PriceStruct { val = item[4] });
                table5Data.Add(new PriceStruct { val = item[5] });
                table5Data.Add(new PriceStruct { val = item[6] });

            }

            var table5Head = new List<PriceStruct>();
            table5Head.Add(new PriceStruct { val = " Формат и сорт бумаги " });
            table5Head.Add(new PriceStruct { val = " На печать " });
            table5Head.Add(new PriceStruct { val = " На приладку " });
            table5Head.Add(new PriceStruct { val = " Всего (листов) " });
            table5Head.Add(new PriceStruct { val = "            Всего (кг)             " });
            table5Head.Add(new PriceStruct { val = " Цена (сом) " });
            table5Head.Add(new PriceStruct { val = " Стоимость бумаги(сом) " });

            //Создание Excel Документа и заполнение его данными
            using (var eC = new ExcelPackage())
            {

                eC.Workbook.Properties.Author = "GVC Soft";
                eC.Workbook.Properties.Title = "Расчётные таблицы по отчёту";
                eC.Workbook.Properties.Company = "NSK GVC";

                var sheet = eC.Workbook.Worksheets.Add("Расчётные таблицы");

                //sheet.Cells.Style.HorizontalAlignment = OfficeOpenXml
                //.Style
                //.ExcelHorizontalAlignment
                //.Center;
                //sheet.Cells.Style.VerticalAlignment = OfficeOpenXml
                //.Style
                //.ExcelVerticalAlignment
                //.Center;

                var row = 8;
                var col = 2;

                //Шапка

                sheet.Cells[5, 6].Value = "ГВЦ Нацстаткома Кыргызской Республики";
                sheet.Cells[5, 6].Style.Font.Bold = true;
                sheet.Cells[5, 6].Style.Font.Size = 16;

                sheet.Cells[6, 7].Value = "отдел полиграфических работ";
                sheet.Cells[6, 7].Style.Font.Bold = true;
                sheet.Cells[6, 7].Style.Font.Size = 13;

                //Добавление заголовков
                foreach (var item in headData)
                {
                    sheet.Cells[row, col].Value = item.val;                   
                    sheet.Cells[row, col].Style.Font.Bold = true;
                    sheet.Cells[row, col].AutoFitColumns();
                    row++;
                }               
                col++;

                //Добавление данных
                int l = 0;
                row = 8;
                foreach (var item in priceData)
                {
                    sheet.Cells[row, col].Value = item.val;
                    sheet.Cells[row, col].AutoFitColumns();
                    row++;
                    l++;
                }

                //Добавление таблиц                                  
                row = 27;
                col = 2;

                //Печать на офсетных машинах
                sheet.Cells[row, col].Value = "Печать текста на офсетных машинах";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.Font.Size = 13;
                row++;
                row++;

                foreach (var item in table1Head)
                {
                    sheet.Cells[row, col].Value = item.val;
                    sheet.Cells[row, col].Style.Font.Bold = true;
                    sheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    col++;
                }
                row++;
                col = 2;

                l = 0;
                foreach (var item in table1Data)
                {
                    sheet.Cells[row, col].Value = item.val;
                    sheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].AutoFitColumns();
                    col++;
                    l++;
                    if (l % 4 == 0)
                    {
                        row++;
                        col = 2;
                    }
                }

                //Тиражирование на ризографе
                row = 35;
                col = 2;
                sheet.Cells[row, col].Value = "Тиражирование на ризографе";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.Font.Size = 13;
                row++;
                row++;

                foreach (var item in table2Head)
                {
                    sheet.Cells[row, col].Value = item.val;
                    sheet.Cells[row, col].Style.Font.Bold = true;
                    sheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    col++;
                }
                row++;
                col = 2;

                l = 0;
                foreach (var item in table2Data)
                {
                    sheet.Cells[row, col].Value = item.val;
                    sheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    col++;
                    l++;
                    if (l % 4 == 0)
                    {
                        row++;
                        col = 2;
                    }
                }

                //Тиражирование на цветном принтере
                row = 41;
                col = 2;
                sheet.Cells[row, col].Value = "Тиражирование на цветном принтере";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.Font.Size = 13;
                row++;
                row++;

                foreach (var item in table3Head)
                {
                    sheet.Cells[row, col].Value = item.val;
                    sheet.Cells[row, col].Style.Font.Bold = true;
                    sheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    col++;
                }
                row++;
                col = 2;

                l = 0;
                foreach (var item in table3Data)
                {
                    sheet.Cells[row, col].Value = item.val;
                    sheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    col++;
                    l++;
                    if (l % 5 == 0)
                    {
                        row++;
                        col = 2;
                    }
                }

                //Тиражирование на ксероксе
                row = 49;
                col = 2;
                sheet.Cells[row, col].Value = "Тиражирование на ксероксе";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.Font.Size = 13;
                row++;
                row++;

                foreach (var item in table4Head)
                {
                    sheet.Cells[row, col].Value = item.val;
                    sheet.Cells[row, col].Style.Font.Bold = true;
                    sheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                 
                    col++;
                }
                row++;
                col = 2;

                l = 0;
                foreach (var item in table4Data)
                {
                    sheet.Cells[row, col].Value = item.val;
                    sheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    col++;
                    l++;
                    if (l % 5 == 0)
                    {
                        row++;
                        col = 2;
                    }
                }

                //Расход бумаги
                row = 57;
                col = 2;
                sheet.Cells[row, col].Value = "Расход бумаги";
                sheet.Cells[row, col].Style.Font.Bold = true;
                sheet.Cells[row, col].Style.Font.Size = 13;
                row++;
                row++;

                foreach (var item in table5Head)
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
                col = 2;

                l = 0;
                foreach (var item in table5Data)
                {
                    sheet.Cells[row, col].Value = item.val;
                    sheet.Cells[row, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[row, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    col++;
                    l++;
                    if (l % 7 == 0)
                    {
                        row++;
                        col = 2;
                    }
                }

                //Сохраняем в файл
                var bin = eC.GetAsByteArray();
                File.WriteAllBytes(@"documents\CalcTables.xlsx", bin);

            }

        }

    }
}
