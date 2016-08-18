using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Reflection;
using NPOI.HSSF.UserModel;

namespace MelavaWebsite.Common
{
    public class ExcelManager
    {
        public HSSFWorkbook GetWorkBookFor<T>(List<T> list)
        {
            //Create new Excel workbook
            var workbook = new HSSFWorkbook();
            List<string> excludeCol = new List<string>() { "CreatedBy", "CreatedOn", "ModifiedBy", "ModifiedOn", "CalledByUser", "CalledDate", "AuditEvent", "ExtensionData" };

            //Create new Excel sheet
            var sheet = workbook.CreateSheet();

            Type typ = typeof(T);

            PropertyInfo[] pi = typ.GetProperties();
            int i = 0;
            foreach (PropertyInfo p in pi)
            {
                if (excludeCol.Find(m => m.Contains(p.Name)) != null && excludeCol.Find(m => m.Contains(p.Name)).Length > 0)
                    continue;
                if ((string.Equals(p.Name.Substring(p.Name.Length - 2, 2), "id", StringComparison.OrdinalIgnoreCase)))
                    continue;

                sheet.SetColumnWidth(i, 30 * 256);
                i++;
            }

            //Create a header row
            var headerRow = sheet.CreateRow(0);
            i = 0;
            foreach (PropertyInfo p in pi)
            {
                if (excludeCol.Find(m => m.Contains(p.Name)) != null && excludeCol.Find(m => m.Contains(p.Name)).Length > 0)
                    continue;
                if ((string.Equals(p.Name.Substring(p.Name.Length - 2, 2), "id", StringComparison.OrdinalIgnoreCase)))
                    continue;

                //Set the column names in the header row
                headerRow.CreateCell(i).SetCellValue(p.Name);
                i++;
            }

            //(Optional) freeze the header row so it is not scrolled
            sheet.CreateFreezePane(0, 1);

            int rowNumber = 1;
            int columnNumber = 0;
            //Populate the sheet with values from the grid data
            foreach (object obj in list)
            {
                if (rowNumber == 65536)
                    break;
                //Create a new row
                var row = sheet.CreateRow(rowNumber++);
                columnNumber = 0;
                foreach (PropertyInfo p in pi)
                {
                    if (excludeCol.Find(m => m.Contains(p.Name)) != null && excludeCol.Find(m => m.Contains(p.Name)).Length > 0)
                        continue;
                    if ((string.Equals(p.Name.Substring(p.Name.Length - 2, 2), "id", StringComparison.OrdinalIgnoreCase)))
                        continue;

                    Type type = p.PropertyType;
                    if (p.GetValue(obj, null) == null)
                    {
                        row.CreateCell(columnNumber++).SetCellValue(string.Empty);
                        continue;
                    }

                    if (string.Equals(p.Name, "Status", StringComparison.OrdinalIgnoreCase))
                    {
                        if (p.PropertyType == typeof(System.String))
                        {
                            row.CreateCell(columnNumber++).SetCellValue((p.GetValue(obj, null)).ToString());
                            continue;
                        }

                        int status = Convert.ToInt16(p.GetValue(obj, null));
                        //Set values for the cells
                        if (typ.FullName.Contains("CustomerBillingContract"))
                        {
                            switch (status)
                            {
                                case 1:
                                    row.CreateCell(columnNumber++).SetCellValue("Renew");
                                    break;
                                case 2:
                                    row.CreateCell(columnNumber++).SetCellValue("Current");
                                    break;
                                case 3:
                                    row.CreateCell(columnNumber++).SetCellValue("Expired");
                                    break;
                                default:
                                    break;
                            }
                        }
                        else
                            row.CreateCell(columnNumber++).SetCellValue((status == 1 ? "Active" : "Disabled").ToString());

                    }
                    else
                    {
                        //Set values for the cells
                        row.CreateCell(columnNumber++).SetCellValue((p.GetValue(obj, null)).ToString());
                    }
                }
            }
            return workbook;
        }

        /// <summary>
        /// Method Name : GetWorkBookFor
        /// Description : Generate work book which helps to generate excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="lstT"></param>
        /// <param name="lstexcelColumConfiguration"></param>
        /// <returns></returns>
        public HSSFWorkbook GetWorkBookFor<T>(IEnumerable<T> lstT, List<ExcelColumConfiguration> lstexcelColumConfiguration)
        {
            //Create new Excel workbook
            var workbook = new HSSFWorkbook();

            List<T> list = lstT.ToList();
            int count = Convert.ToInt32(ConfigurationManager.AppSettings["MaximumRecordsPerSheet"]);
            List<List<T>> listArray = ChunkBy(list, count);
            IEnumerable<T> pageList;
            int sheetCount = 1;
            foreach (var item in listArray)
            {
                pageList = item;
                //Create new Excel sheet
                var sheet = workbook.CreateSheet("Sheet " + sheetCount);

                int rowNumber = SetExcelColumnWidthAndName(workbook, sheet, lstexcelColumConfiguration);

                //Add Data To Excel
                AddDataToExcel<T>(pageList, sheet, workbook, rowNumber, lstexcelColumConfiguration);

                sheetCount++;
            }
            //Write the workbook to a memory stream
            return workbook;
        }

        public static List<List<T>> ChunkBy<T>(List<T> source, int chunkSize)
        {
            return source
                .Select((x, i) => new { Index = i, Value = x })
                .GroupBy(x => x.Index / chunkSize)
                .Select(x => x.Select(v => v.Value).ToList())
                .ToList();
        }

        /// <summary>
        /// Method Name : SetExcelColumnWidthAndName
        /// Description : Set Column Width And Name to Excel
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="lstexcelColumConfiguration"></param>
        /// <returns></returns>
        public int SetExcelColumnWidthAndName(HSSFWorkbook wb, NPOI.SS.UserModel.ISheet sheet, List<ExcelColumConfiguration> lstexcelColumConfiguration)
        {
            //(Optional) set the width of the columns
            foreach (ExcelColumConfiguration excelColumConfiguration in lstexcelColumConfiguration)
            {
                sheet.SetColumnWidth(excelColumConfiguration.ColumnIndex, excelColumConfiguration.ColumnWidth);
            }
            //Create a header row
            var headerRow = sheet.CreateRow(0);
            int rowNumber = 1;

            //Create style for header row
            HSSFFont hFont = (HSSFFont)wb.CreateFont();
            hFont.FontHeightInPoints = 11;
            hFont.FontName = "Calibri";
            hFont.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

            HSSFCellStyle hStyle = (HSSFCellStyle)wb.CreateCellStyle();
            hStyle.SetFont(hFont);

            //(Optional) freeze the header row so it is not scrolled
            sheet.CreateFreezePane(0, 1);
            //Set the column names in the header row
            foreach (ExcelColumConfiguration excelColumConfiguration in lstexcelColumConfiguration)
            {
                NPOI.SS.UserModel.ICell headerCell = headerRow.CreateCell(excelColumConfiguration.ColumnIndex);//.SetCellValue(excelColumConfiguration.ColumnHeaderName);
                headerCell.SetCellValue(excelColumConfiguration.ColumnHeaderName);
                headerCell.CellStyle = hStyle;
            }
            return rowNumber;
        }

        /// <summary>
        /// Method Name : AddDataToExcel
        /// Description : Add dtata to excel file
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="lstT"></param>
        /// <param name="sheet"></param>
        /// <param name="rowNumber"></param>
        /// <param name="lstexcelColumConfiguration"></param>
        public void AddDataToExcel<T>(IEnumerable<T> lstT, NPOI.SS.UserModel.ISheet sheet, HSSFWorkbook wb,
            int rowNumber, List<ExcelColumConfiguration> lstexcelColumConfiguration)
        {
            //Create style for non header rows
            HSSFFont hFontNormal = (HSSFFont)wb.CreateFont();
            hFontNormal.FontHeightInPoints = 11;
            hFontNormal.FontName = "Calibri";
            hFontNormal.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Normal;

            HSSFCellStyle hStyleNormal = (HSSFCellStyle)wb.CreateCellStyle();
            hStyleNormal.SetFont(hFontNormal);

            //Populate the sheet with values from the grid data
            Type typ = typeof(T);
            PropertyInfo[] pi = typ.GetProperties();
            foreach (T columnData in lstT)
            {
                //Create a new row
                var row = sheet.CreateRow(rowNumber++);
                foreach (PropertyInfo p in pi)
                {
                    bool isExist = lstexcelColumConfiguration.Any(e => e.DBColumnName == p.Name);
                    if (isExist)
                    {
                        int columnIndex = (from config in lstexcelColumConfiguration
                                           where config.DBColumnName == p.Name
                                           select (config.ColumnIndex)).First();

                        string data = Convert.ToString(p.GetValue(columnData, null));

                        NPOI.SS.UserModel.ICell cell2 = row.CreateCell(columnIndex);
                        //Bug 16615:Release 9 | Exported file of Portfolio TID/ Mapping master screen doesn't have newly added fields.

                        if (string.Equals(p.Name, "Status", StringComparison.OrdinalIgnoreCase))
                        {
                            //cell2.SetCellValue(data == "1" ? "Active" : "InActive");
                            // Bug Id - 17856 : Excel data dose not match with grid data on multiple pages.
                            switch (data)
                            {
                                case "0":
                                    cell2.SetCellValue("Inactive");
                                    break;
                                case "1":
                                    cell2.SetCellValue("Active");
                                    break;
                                // Fix for Bug Id - 19250
                                //case "2":
                                //    cell2.SetCellValue("FutureInactive");
                                //    break;
                                case "3":
                                    cell2.SetCellValue("Current");
                                    break;
                                case "4":
                                    cell2.SetCellValue("Renewed");
                                    break;
                                case "5":
                                    cell2.SetCellValue("Expired");
                                    break;
                            }
                        }
                        else if (string.Equals(p.Name, "UsesAvailability", StringComparison.OrdinalIgnoreCase))
                        {


                            //cell2.SetCellValue(data == "True" ? "Yes" : "No");
                            switch (data)
                            {

                                case "True":
                                    cell2.SetCellValue("Yes");
                                    break;
                                case "False":
                                    cell2.SetCellValue("No");
                                    break;
                            }


                        }
                        else
                            cell2.SetCellValue(data);

                        //Apply Style to cells
                        cell2.CellStyle = hStyleNormal;

                    }
                }

            }
        }

    }

    public class ExcelColumConfiguration
    {
        public int ColumnIndex { get; set; }
        public int ColumnWidth { get; set; }
        public string ColumnHeaderName { get; set; }
        public string DBColumnName { get; set; }
    }
}
