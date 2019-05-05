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
using System.Text.RegularExpressions;

namespace ExcelClassLibrary
{

    public class ExcelClass_old
    {
        private XElement XmlData { get; set; }
        private XElement XmlWorkBook { get; set; }
        private WorksheetPart WorksheetPart { get; set; }

        private string WSheetName { get; set; }
        private string templatefilepath { get; set; }
        private string resultfilepath { get; set; }


        internal static XNamespace ns_s = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        internal static XNamespace ns_r = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        internal static XNamespace ns_c = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/chart");

        public ExcelClass_old(string templatepath, string outputpath)
        {
            templatefilepath = templatepath;
            resultfilepath = outputpath;

        }
        

        public string AddSheetWithChart(string TemplateSheetName, List<List<object>> ChartData, string[] SeriesLabels, Dictionary<String, String> ReplacementDict)
        {
            //Assume error
            string result = "Error Creating Chart";
            WSheetName = TemplateSheetName;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(resultfilepath + "CloudReport.xlsx", true))
            {

                //create clone of template chart sheet as TemplateSheetName
                CloneSheet(document, "TemplateChartSheet", WSheetName);
            }

            try
            {

                using (SpreadsheetDocument document = SpreadsheetDocument.Open(resultfilepath + "CloudReport.xlsx", true))
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

                        Int32 currentRowNum = 4;
                        Char currentColumn;

                        //  inserting Chart data in excel sheet rows
                        foreach (List<Object> rowitem in ChartData)
                        {
                            //this is a line
                            currentColumn = 'B';

                            foreach (var obj in rowitem)
                            {
                                //inserted values are NOT of type string
                                UpdateValue(workbookPart, WSheetName, currentColumn + currentRowNum.ToString(), obj.ToString(), 10, false);
                                currentColumn++;
                            }

                            currentRowNum++;
                        }

                        // update chart part to reflect the new data
                        //Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where((s) => s.Name == WSheetName).FirstOrDefault();
                        //WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                        //ChartPart part = worksheetPart.DrawingsPart.ChartParts.First();

                        //XElement XmlChart;
                        //using (XmlReader xmlr = XmlReader.Create(part.GetStream()))
                        //{
                        //    XmlChart = XElement.Load(xmlr);
                        //}

                        //Char endColChar = 'B';                        
                        //List<Object> demorowitem = ChartData[1];
                        //int colCount = demorowitem.Count();

                        //while (colCount > 1)
                        //{
                        //    endColChar++;
                        //    colCount--;
                        //}

                        ////var catrange = from cat in XmlChart.Descendants(ns_c + "cat")
                        ////               from f in cat.Descendants(ns_c + "f")
                        ////               select f;
                        ////foreach (var catelement in catrange)
                        ////{
                        ////    String basevalue = catelement.Value.Split('$')[0];
                        ////    catelement.Value = basevalue + "$" + "B" + "$3:$" + endColChar + "$3";
                        ////}
                              
                        //var valrange = from val in XmlChart.Descendants(ns_c + "val")
                        //               from f in val.Descendants(ns_c + "f")
                        //               select f;

                        //Char startColChar = 'B';
                        //currentRowNum = 3;
                        //foreach (var valelement in valrange)
                        //{
                        //    String basevalue = valelement.Value.Split('$')[0];
                        //    valelement.Value = basevalue + "$" + startColChar + "$" + currentRowNum + ":$" + endColChar + "$" + currentRowNum;
                        //    startColChar++;
                        //}
                                              
                        //using (Stream s = part.GetStream(FileMode.Create, FileAccess.Write))
                        //{
                        //    using (XmlWriter xmlw = XmlWriter.Create(s))
                        //    {
                        //        XmlChart.WriteTo(xmlw);
                        //    }
                        //}
                        //result = "Chart updated";


                    }

                    //  Replace  [Tag Name] by “Tag Value” in a worksheet
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


        public string InitBookCreation(string TemplateFileName, string ResultFileName, string TemplateSheetsFileName)
        {
            //Assume Error
            string result = "Cannot Create Result File";

            result = CopyFile(templatefilepath + TemplateFileName, resultfilepath + ResultFileName);

            if (result == "Copied file")
            {
                try
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(resultfilepath + ResultFileName, true))
                    {
                        IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == TemplateSheetsFileName);
                        if (sheets.Count() == 0)
                        {
                            // The specified worksheet does not exist.
                            // Create a new worksheet with name as TemplateSheetsFileName
                            // Add a blank WorksheetPart.
                            WorksheetPart newWorksheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();
                            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                            Sheets wbsheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                            string relationshipId = document.WorkbookPart.GetIdOfPart(newWorksheetPart);

                            // Get a unique ID for the new worksheet.
                            uint sheetId = 1;
                            if (wbsheets.Elements<Sheet>().Count() > 0)
                            {
                                sheetId = wbsheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                            }

                            // Give the new worksheet a name.
                            string sheetName = TemplateSheetsFileName;

                            // Append the new worksheet and associate it with the workbook.
                            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                            wbsheets.Append(sheet);
                        }
                    }
                    result = "Result file ready";
                }
                catch (Exception ex)
                {
                    result = ex.Message;
                }

            }

            return result;

        }


        public string AddSheetWithTable(string TemplateSheetName, List<List<object>> TableData, Dictionary<String, String> ReplacementDict)
        {

            string result = "Error adding table data";
            WSheetName = TemplateSheetName;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(resultfilepath + "CloudReport.xlsx", true))
            {

                //create clone of template chart sheet as TemplateSheetName
                CloneSheet(document, "TemplateTableSheet", WSheetName);
            }

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(resultfilepath + "CloudReport.xlsx", true))
                {



                    IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == WSheetName);

                    if (sheets.Count() == 0)
                    {
                        return result = "No worksheet found named as " + WSheetName;
                    }
                    else
                    {
                        WorkbookPart workbookPart = document.WorkbookPart;
                       // int numOfRowsToAdd = TableData.Count();

                        Int32 currentRowNum = 3;
                        Char currentColumn;

                        //// inserting tabular data in excel sheet rows
                        //foreach (List<Object> rowitem in TableData)
                        //{

                        //    //this is a line
                        //    currentColumn = 'B';

                        //    foreach (var obj in rowitem)
                        //    {
                        //        //inserting values as string type
                        //       UpdateValue(workbookPart, WSheetName, currentColumn + currentRowNum.ToString(), obj.ToString(), 3, true);
                        //      //  XLInsertNumberIntoCell(resultfilepath + "CloudReport.xlsx", WSheetName, currentColumn + currentRowNum.ToString(), 4);
                        //        currentColumn++;
                        //    }

                        //    currentRowNum++;
                        //}
                        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where((s) => s.Name == WSheetName).FirstOrDefault();
                        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                       
                        using (XmlReader xmlr = XmlReader.Create(worksheetPart.GetStream()))
                        {
                            XmlData = XElement.Load(xmlr);
                        }

                       
                        XElement sheetData = XmlData.Descendants(ns_s + "sheetData").Single();

                        //now we have a series of row, first of all build the new data
                        List<XElement> wselements = new List<XElement>();
                        wselements.AddRange(sheetData.Elements().Take(TableData.Count()));

                        //Int32 currentRowNum = numOfRowsToSkip + 1;

                        currentRowNum = 3;
                        foreach (List<Object> rowitem in TableData)
                        {
                            
                            currentColumn = 'B';
                            XElement row = new XElement(ns_s + "row",
                             new XAttribute( "r", currentRowNum),
                             new XAttribute( "spans", "1:" + (rowitem.Count+1).ToString()));
                           
                            foreach (var obj in rowitem)
                            {
                                
                                row.Add(new XElement(ns_s + "c",
                                                               new XAttribute( "r", currentColumn + currentRowNum.ToString()),
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

                    //Replace  [Tag Name] by “Tag Value” in a worksheet
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

            result = "Table Data inserted";

            return result;
        }

        //Remove Template Sheets from Result spreadsheet.
        public string EndBookCreation(string ResultFileName)
        {
            string result = "Error ending book creation";

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(resultfilepath + ResultFileName, true))
                {

                    IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "TemplateChartSheet" || s.Name == "TemplateTableSheet");

                    if (sheets.Count() == 0)
                    {
                        return result = "No template chart sheet and data sheet found";
                    }
                    else
                    {
                        XLDeleteSheet(document, "TemplateChartSheet");
                        XLDeleteSheet(document, "TemplateTableSheet");
                        result = "Template Sheets Removed. End Book Creation succedded";
                    }

                    IEnumerable<Sheet> tempfilesheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "TemplateSheetsFile");

                    if (tempfilesheets.Count() == 0)
                    {
                        result = result + "sheet name 'TemplateSheetsFile' not present";
                    }
                    else
                    {
                        WorkbookPart workbookPart = document.WorkbookPart;
                        List<Sheet> allSheetslist = new List<Sheet>();
                        IEnumerable<Sheet> allsheets = workbookPart.Workbook.Descendants<Sheet>();
                        allSheetslist = allsheets.ToList();

                        Int32 currentRowNum = 1;

                        UpdateValue(workbookPart, "TemplateSheetsFile", "A" + currentRowNum.ToString(), "Sheet ID ", 3, true);                         
                        UpdateValue(workbookPart, "TemplateSheetsFile", "B" + currentRowNum.ToString(), "Sheet Name", 3, true);
                        currentRowNum++;
                        // inserting sheet details in TemplateSheetsFile sheet
                        foreach (Sheet sheet in allSheetslist)
                        {

                            //inserting values as string type
                            UpdateValue(workbookPart, "TemplateSheetsFile", "A" + currentRowNum.ToString(), sheet.SheetId, 3, false);
                            UpdateValue(workbookPart, "TemplateSheetsFile", "B" + currentRowNum.ToString(), sheet.Name.ToString(), 3, true);

                            currentRowNum++;
                        }
                        result = result + "\nTemplateSheetsFile updated";
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




        public static bool XLInsertNumberIntoCell(string fileName, string sheetName, string addressName, int value)
        {

            // Given a file, a sheet, and a cell, insert a specified value.
            // For example: InsertNumberIntoCell("C:\Test.xlsx", "Sheet3", "C3", 14)

            // Assume failure.
            bool returnValue = false;

            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetName).FirstOrDefault();
                if (theSheet != null)
                {
                    Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(theSheet.Id))).Worksheet;
                    Cell theCell = InsertCellInWorksheet2(ws, addressName);

                    // Set the value of cell A1.
                    theCell.CellValue = new CellValue(value.ToString());
                    theCell.DataType = new EnumValue<CellValues>(CellValues.Number);

                    // Save the worksheet.
                    ws.Save();
                    returnValue = true;
                }
            }

            return returnValue;
        }


        // Given a Worksheet and an address (like "AZ254"), either return a cell reference, or 
        // create the cell reference and return it.
        private static Cell InsertCellInWorksheet2(Worksheet ws, string addressName)
        {

            // Use regular expressions to get the row number and column name.
            // If the parameter wasn't well formed, this code
            // will fail:
            Regex rx = new Regex("^(?<col>\\D+)(?<row>\\d+)");
            Match m = rx.Match(addressName);
            uint rowNumber = uint.Parse(m.Result("$${row}"));
            string colName = m.Result("$${col}");

            SheetData sheetData = ws.GetFirstChild<SheetData>();
            string cellReference = (colName + rowNumber.ToString());
            Cell theCell = null;

            // If the worksheet does not contain a row with the specified row index, insert one.
            var theRow = sheetData.Elements<Row>().
              Where(r => r.RowIndex.Value == rowNumber).FirstOrDefault();
            if (theRow == null)
            {
                theRow = new Row();
                theRow.RowIndex = rowNumber;
                sheetData.Append(theRow);
            }

            // If the cell you need already exists, return it.
            // If there is not a cell with the specified column name, insert one.  
            Cell refCell = theRow.Elements<Cell>().
              Where(c => c.CellReference.Value == cellReference).FirstOrDefault();
            if (refCell != null)
            {
                theCell = refCell;
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                foreach (Cell cell in theRow.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                theCell = new Cell();
                theCell.CellReference = cellReference;

                theRow.InsertBefore(theCell, refCell);
            }
            return theCell;
        }


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
