using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace ExcelClassLibrary
{
    public class testchart
    {
        private XElement XmlData { get; set; }
        private XElement XmlWorkBook { get; set; }
        private WorksheetPart WorksheetPart { get; set; }

        private String WSheetName { get; set; }
        private string templatefilepath { get; set; }
        private string resultfilepath { get; set; }


        internal static XNamespace ns_s = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        internal static XNamespace ns_r = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        internal static XNamespace ns_c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");

        public testchart(string templatepath, string outputpath)
        {
            templatefilepath = templatepath;
            resultfilepath = outputpath;

        }


        public string AddSheetWithChart(string TemplateSheetName, List<List<object>> ChartData, string[] SeriesLabels, Dictionary<String, String> ReplacementDict)
        {
            //Assume error
            string result = "Error Creating Chart";
            WSheetName = TemplateSheetName;

            result = CopyFile(templatefilepath + "ExcelTemplate.xlsx", resultfilepath + "CloudReport1.xlsx");



            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(resultfilepath + "CloudReport1.xlsx", true))
                {

                    //create clone of template chart sheet as TemplateSheetName
                    CloneSheet(document, "TemplateChartSheet", WSheetName);
                }


                using (SpreadsheetDocument document = SpreadsheetDocument.Open(resultfilepath + "CloudReport1.xlsx", true))
                {

                    IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == WSheetName);

                    if (sheets.Count() == 0)
                    {
                        return result = "No worksheet found named as " + WSheetName;
                    }
                    else
                    {

                        WorkbookPart workbookPart = document.WorkbookPart;
                        int numOfRowsToAdd = ChartData.Count();

                        //currentRowNum = 3;
                        //Char currentColumn;

                        // inserting Chart data in excel sheet rows
                        //foreach (List<Object> rowitem in ChartData)
                        //{
                        //    //this is a line
                        //    currentColumn = 'B';

                        //    foreach (var obj in rowitem)
                        //    {
                        //inserted values are NOT of type string
                        //        UpdateValue(workbookPart, WSheetName, currentColumn + currentRowNum.ToString(), obj.ToString(), 3, false);
                        //        currentColumn++;
                        //    }

                        //    currentRowNum++;
                        //}

                        // update chart part to reflect the new data
                        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where((s) => s.Name == WSheetName).FirstOrDefault();
                        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                        ChartPart part = worksheetPart.DrawingsPart.ChartParts.First();

                        XElement XmlChart;
                        using (XmlReader xmlr = XmlReader.Create(part.GetStream()))
                        {
                            XmlChart = XElement.Load(xmlr);
                        }

                        XElement barChartData = XmlChart.Descendants(ns_c + "barChart").Single();

                        List<XElement> elements = new List<XElement>();
                        elements.AddRange(barChartData.Elements().Take(ChartData.Count()));
                        // these 2 values are hard coded for now 
                        Char startColChar = 'B';
                        int currentRowNum = 3;

                        // Char endColChar = 'B';
                        int colCount = 0;
                        int m = 0;
                        foreach (List<object> rowitem in ChartData)
                        {
                            Char endColChar = 'B';
                            colCount = rowitem.Count();

                            while (colCount > 1)
                            {
                                endColChar++;
                                colCount--;
                            }

                            XElement ser = new XElement(ns_c + "ser");
                            ser.Add(new XElement(ns_c + "idx", new XAttribute("val", m)));
                            ser.Add(new XElement(ns_c + "order", new XAttribute("val", m)));
                            ser.Add(new XElement(ns_c + "tx", new XElement(ns_c + "v", new XText(SeriesLabels[m].ToString()))));

                            m++;
                            XElement cat = new XElement(ns_c + "cat");
                            XElement strRef = new XElement(ns_c + "strRef");
                            cat.Add(strRef);
                            strRef.Add(new XElement(ns_c + "f", new XText("'" + WSheetName + "'!$" + "B" + "$2:$" + endColChar + "$2")));

                            //XElement strCache = new XElement(ns_c + "strCache");
                            //strRef.Add(strCache);
                            //strCache.Add(new XElement(ns_c + "ptCount", new XAttribute("val", rowitem.Count())));
                            //int j = 0;
                            //foreach (var obj in rowitem)
                            //{
                            //    strCache.Add(new XElement(ns_c + "pt",
                            //                               new XAttribute("idx", j),
                            //                               new XElement(ns_s + "v", new XText(obj.ToString()))));
                            //    j++;
                            //}   

                            XElement val = new XElement(ns_c + "val");
                            XElement numRef = new XElement(ns_c + "numRef");
                            val.Add(numRef);
                            numRef.Add(new XElement(ns_c + "f", new XText("'" + WSheetName + "'!$" + startColChar + "$" + currentRowNum + ":$" + endColChar + "$" + currentRowNum)));
                            XElement numCache = new XElement(ns_c + "numCache");
                            numRef.Add(numCache);
                            //numCache.Add(new XElement(ns_c + "formatCode", new XText(""$"#,##0_);\("$"#,##0\)");
                            numCache.Add(new XElement(ns_c + "ptCount", new XAttribute("val", rowitem.Count())));
                            int k = 0;
                            foreach (var obj in rowitem)
                            {
                                numCache.Add(new XElement(ns_c + "pt",
                                                           new XAttribute("idx", k),
                                                           new XElement(ns_s + "v", new XText(obj.ToString()))));
                                k++;
                            }
                            ser.Add(cat);
                            ser.Add(val);
                            elements.Add(ser);

                            currentRowNum++;
                        }

                        //Now we have all elements
                        barChartData.Elements().Remove();
                        barChartData.Add(elements);


                        using (Stream s = part.GetStream(FileMode.Create, FileAccess.Write))
                        {
                            using (XmlWriter xmlw = XmlWriter.Create(s))
                            {
                                XmlChart.WriteTo(xmlw);
                            }
                        }
                        result = "Chart updated";


                        ////save data in the sheet

                        //currentRowNum = 3;
                        //Char currentColumn;

                        ////   inserting Chart data in excel sheet rows
                        //foreach (List<Object> rowitem in ChartData)
                        //{
                        //    //this is a line
                        //    currentColumn = 'B';

                        //    foreach (var obj in rowitem)
                        //    {
                        //        //   inserted values are NOT of type string
                        //        UpdateValue(workbookPart, WSheetName, currentColumn + currentRowNum.ToString(), obj.ToString(), 3, false);
                        //        currentColumn++;
                        //    }

                        //    currentRowNum++;
                        //}


                        using (XmlReader xmlr = XmlReader.Create(worksheetPart.GetStream()))
                        {
                            XmlData = XElement.Load(xmlr);
                        }

                        XElement sheetData = XmlData.Descendants(ns_s + "sheetData").Single();
                        //now we have a series of row, first of all build the new data
                        List<XElement> wselements = new List<XElement>();
                        wselements.AddRange(sheetData.Elements().Take(ChartData.Count()));

                        //Int32 currentRowNum = numOfRowsToSkip + 1;
                        Char currentColumn;
                        currentRowNum = 1;
                        foreach (List<Object> rowitem in ChartData)
                        {
                            currentColumn = 'B';
                            XElement row = new XElement(ns_s + "row",
                             new XAttribute(ns_s + "r", currentRowNum),
                             new XAttribute(ns_s + "spans", "1:" + rowitem.Count));
                            int k = 0;
                            foreach (var obj in rowitem)
                            {
                                row.Add(new XElement(ns_s + "c",
                                                               new XAttribute(ns_s + "r", currentColumn + currentRowNum),
                                                               new XElement(ns_s + "v", new XText(obj.ToString()))));
                                currentColumn++;
                            }

                            wselements.Add(row);
                            currentRowNum++;
                        }


                        //Now we have all elements
                        sheetData.Elements().Remove();
                        sheetData.Add(wselements);
                        using (Stream s = worksheetPart.GetStream(FileMode.Create, FileAccess.Write))
                        {
                            using (XmlWriter xmlw = XmlWriter.Create(s))
                            {
                                XmlData.WriteTo(xmlw);
                            }
                        }



                    }

                    ////   Replace  [Tag Name] by “Tag Value” in a worksheet
                    foreach (KeyValuePair<string, string> item in ReplacementDict)
                    {
                        string tagname = item.Key;
                        string tagvalue = item.Value;

                        WorkbookPart workbookPart = document.WorkbookPart;
                        SharedStringTablePart sharedStringsPart = workbookPart.SharedStringTablePart;
                        IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Text> sharedStringTextElements = sharedStringsPart.SharedStringTable.Descendants<DocumentFormat.OpenXml.Spreadsheet.Text>();
                        DoReplace(sharedStringTextElements, tagname, tagvalue);

                        IEnumerable<WorksheetPart> worksheetParts = workbookPart.GetPartsOfType<WorksheetPart>();
                        foreach (var worksheet in worksheetParts)
                        {
                            var allTextElements = worksheet.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Text>();
                            DoReplace(allTextElements, tagname, tagvalue);
                        }
                    }


                }
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }


            return result;
        }


        #region "Helper methods"

        //Create copy of template sheet
        private static void CloneSheet(SpreadsheetDocument document, string templateWSheetName, string clonedWSheetName)
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where((s) => s.Name == templateWSheetName).FirstOrDefault();
            WorksheetPart sourceSheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

            //Take advantage of AddPart for deep cloning 
            SpreadsheetDocument tempSheet = SpreadsheetDocument.Create(new MemoryStream(), document.DocumentType);
            WorkbookPart tempWorkbookPart = tempSheet.AddWorkbookPart();
            WorksheetPart tempWorksheetPart = tempWorkbookPart.AddPart<WorksheetPart>(sourceSheetPart);

            //Add cloned sheet and all associated parts to workbook
            WorksheetPart clonedSheet = workbookPart.AddPart<WorksheetPart>(tempWorksheetPart);
            int numTableDefParts = sourceSheetPart.GetPartsCountOfType<TableDefinitionPart>();

            int tableId = numTableDefParts;

            //Clean up table definition parts (tables need unique ids)
            if (numTableDefParts != 0)
                FixupTableParts(clonedSheet, numTableDefParts, tableId);

            //There can only be one sheet that has focus 
            SheetViews views = clonedSheet.Worksheet.GetFirstChild<SheetViews>();
            if (views != null)
            {
                views.Remove();
                clonedSheet.Worksheet.Save();
            }

            //Add new sheet to main workbook part 
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            Sheet copiedSheet = new Sheet();
            copiedSheet.Name = clonedWSheetName;
            copiedSheet.Id = workbookPart.GetIdOfPart(clonedSheet);
            copiedSheet.SheetId = (uint)sheets.ChildElements.Count + 1;
            sheets.Append(copiedSheet);
            //Save Changes 
            workbookPart.Workbook.Save();

        }


        private static void FixupTableParts(WorksheetPart worksheetPart, int numTableDefParts, int tableId)
        {
            //Every table needs a unique id and name 
            foreach (TableDefinitionPart tableDefPart in worksheetPart.TableDefinitionParts)
            {
                tableId++;
                tableDefPart.Table.Id = (uint)tableId;
                tableDefPart.Table.DisplayName = "CopiedTable" + tableId;
                tableDefPart.Table.Name = "CopiedTable" + tableId;
                tableDefPart.Table.Save();
            }
        }

        //For String Tag Replacement
        private static void DoReplace(IEnumerable<DocumentFormat.OpenXml.Spreadsheet.Text> textElements, string txtkey, string txtreplacement)
        {
            foreach (var text in textElements)
            {
                if (text.Text.Contains(txtkey))
                    text.Text = text.Text.Replace(txtkey, txtreplacement);
            }
        }

        // Given a Worksheet and an address (like "AZ254"), either return a cell reference, or 
        // create the cell reference and return it.
        private Cell InsertCellInWorksheet(Worksheet ws, string addressName)
        {
            SheetData sheetData = ws.GetFirstChild<SheetData>();
            Cell cell = null;

            UInt32 rowNumber = GetRowIndex(addressName);
            Row row = GetRow(sheetData, rowNumber);

            // If the cell you need already exists, return it.
            // If there is not a cell with the specified column name, insert one.  
            Cell refCell = row.Elements<Cell>().
                Where(c => c.CellReference.Value == addressName).FirstOrDefault();
            if (refCell != null)
            {
                cell = refCell;
            }
            else
            {
                cell = CreateCell(row, addressName);
            }
            return cell;
        }

        private Cell CreateCell(Row row, String address)
        {
            Cell cellResult;
            Cell refCell = null;

            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference.Value, address, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            cellResult = new Cell();
            cellResult.CellReference = address;

            row.InsertBefore(cellResult, refCell);
            return cellResult;
        }

        private Row GetRow(SheetData wsData, UInt32 rowIndex)
        {
            var row = wsData.Elements<Row>().
            Where(r => r.RowIndex.Value == rowIndex).FirstOrDefault();
            if (row == null)
            {
                row = new Row();
                row.RowIndex = rowIndex;
                wsData.Append(row);
            }
            return row;
        }

        private UInt32 GetRowIndex(string address)
        {
            string rowPart;
            UInt32 l;
            UInt32 result = 0;

            for (int i = 0; i < address.Length; i++)
            {
                if (UInt32.TryParse(address.Substring(i, 1), out l))
                {
                    rowPart = address.Substring(i, address.Length - i);
                    if (UInt32.TryParse(rowPart, out l))
                    {
                        result = l;
                        break;
                    }
                }
            }
            return result;
        }

        public bool UpdateValue(WorkbookPart wbPart, string sheetName, string addressName, string value, UInt32Value styleIndex, bool isString)
        {
            // Assume failure.
            bool updated = false;

            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().Where((s) => s.Name == sheetName).FirstOrDefault();

            if (sheet != null)
            {
                Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(sheet.Id))).Worksheet;
                Cell cell = InsertCellInWorksheet(ws, addressName);

                if (isString)
                {
                    // Either retrieve the index of an existing string,
                    // or insert the string into the shared string table
                    // and get the index of the new item.
                    int stringIndex = InsertSharedStringItem(wbPart, value);

                    cell.CellValue = new CellValue(stringIndex.ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }
                else
                {
                    cell.CellValue = new CellValue(value);
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                }

                if (styleIndex > 0)
                    cell.StyleIndex = styleIndex;

                // Save the worksheet.
                ws.Save();
                updated = true;
            }

            return updated;
        }

        // Given the main workbook part, and a text value, insert the text into the shared
        // string table. Create the table if necessary. If the value already exists, return
        // its index. If it doesn't exist, insert it and return its new index.
        private int InsertSharedStringItem(WorkbookPart wbPart, string value)
        {
            int index = 0;
            bool found = false;
            var stringTablePart = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            // If the shared string table is missing, something's wrong.
            // Just return the index that you found in the cell.
            // Otherwise, look up the correct text in the table.
            if (stringTablePart == null)
            {
                // Create it.
                stringTablePart = wbPart.AddNewPart<SharedStringTablePart>();
            }

            var stringTable = stringTablePart.SharedStringTable;
            if (stringTable == null)
            {
                stringTable = new SharedStringTable();
            }

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in stringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == value)
                {
                    found = true;
                    break;
                }
                index += 1;
            }

            if (!found)
            {
                stringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(value)));
                stringTable.Save();
            }

            return index;
        }

        // Used to force a recalc of cells containing formulas. The
        // CellValue has a cached value of the evaluated formula. This
        // will prevent Excel from recalculating the cell even if 
        // calculation is set to automatic.
        private bool RemoveCellValue(WorkbookPart wbPart, string sheetName, string addressName)
        {
            bool returnValue = false;

            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().
                Where(s => s.Name == sheetName).FirstOrDefault();
            if (sheet != null)
            {
                Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(sheet.Id))).Worksheet;
                Cell cell = InsertCellInWorksheet(ws, addressName);

                // If there is a cell value, remove it to force a recalc
                // on this cell.
                if (cell.CellValue != null)
                {
                    cell.CellValue.Remove();
                }

                // Save the worksheet.
                ws.Save();
                returnValue = true;
            }

            return returnValue;
        }

        // USed to read the value of cell
        public string XLGetCellValue(WorkbookPart wbPart, string sheetName, string addressName)
        {
            string value = null;

            // Find the sheet with the supplied name, and then use that Sheet object
            // to retrieve a reference to the appropriate worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == sheetName).FirstOrDefault();

            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            // Retrieve a reference to the worksheet part, and then use its Worksheet property to get 
            // a reference to the cell whose address matches the address you've supplied:
            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().
              Where(c => c.CellReference == addressName).FirstOrDefault();

            // If the cell doesn't exist, return an empty string:
            if (theCell != null)
            {
                value = theCell.InnerText;

                // If the cell represents an integer number, you're done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and booleans
                // individually. For shared strings, the code looks up the corresponding
                // value in the shared string table. For booleans, the code converts 
                // the value into the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            // For shared strings, look up the value in the shared strings table.
                            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            // If the shared string table is missing, something's wrong.
                            // Just return the index that you found in the cell.
                            // Otherwise, look up the correct text in the table.
                            if (stringTable != null)
                            {
                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }

            return value;
        }

        public bool XLDeleteSheet(SpreadsheetDocument document, string sheetToDelete)
        {
            // Delete the specified sheet from within the specified workbook.
            // Return True if the sheet was found and deleted, False if it was not.
            // Note that this procedure might leave "orphaned" references, such as strings
            // in the shared strings table. You must take care when adding new strings, for example. 
            // The XLInsertStringIntoCell snippet handles this problem for you.


            WorkbookPart wbPart = document.WorkbookPart;

            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == sheetToDelete).FirstOrDefault();
            if (theSheet == null)
            {
                // The specified sheet doesn't exist.
                return false;
            }

            // Remove the sheet reference from the workbook.
            WorksheetPart worksheetPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            theSheet.Remove();

            // Delete the worksheet part.
            wbPart.DeletePart(worksheetPart);

            // Save the workbook.
            wbPart.Workbook.Save();

            return true;
        }


        // Given a workbook document, and a sheet name, return the WorksheetPart
        // corresponding to the supplied name. If the sheet doesn't exist,
        // the procedure throws an ArgumentException.

        public WorksheetPart XLGetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {

            WorkbookPart wbPart = document.WorkbookPart;

            // Find the sheet with the supplied name, and then use that Sheet object
            // to retrieve a reference to the appropriate worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == sheetName).FirstOrDefault();

            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            return (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
        }

        // Retrieve a List of all the sheets in a workbook.
        public List<Sheet> XLGetAllSheets(string fileName)
        {
            List<Sheet> allSheets = new List<Sheet>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                IEnumerable<Sheet> sheets = wbPart.Workbook.Descendants<Sheet>();
                allSheets = sheets.ToList();
            }
            return allSheets;
        }


        private string CopyFile(string source, string dest)
        {
            string result = "Copied file";
            try
            {
                // Overwrites existing files
                //To copy the Template Book as the result Book and work with the result from then on.
                File.Copy(source, dest, true);
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }
            return result;
        }
        # endregion
    }
}
