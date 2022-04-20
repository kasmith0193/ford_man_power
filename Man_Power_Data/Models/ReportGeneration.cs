using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Man_Power_Data.Models
{


    class ReportGeneration
    {
        private IEnumerable<Section> SectionData;
        public void RerportGeneration()
        {

        }

        public void addSections(IEnumerable<Section> data)
        {
            SectionData = data;
        }

        public string CreateReport(string NewFileLocation, string TemplateFileLocation)
        {
            //make new file and open 
            File.Copy(TemplateFileLocation, NewFileLocation, true);
            SpreadsheetDocument newspreadsheetdoc = SpreadsheetDocument.Open(NewFileLocation, true);
            WorkbookPart workbookpart = newspreadsheetdoc.WorkbookPart;

            //open shared text table
            SharedStringTablePart sharedstringtablepart = workbookpart.SharedStringTablePart;



            foreach (Section sec in SectionData)
            {
                //open new file sections
                IEnumerable<Sheet> sheets = newspreadsheetdoc.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sec.Sheet);
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)newspreadsheetdoc.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();



                //open section excel data
                SpreadsheetDocument tempspreadsheetdoc = SpreadsheetDocument.Open(sec.FileSelection.FileLocation, false);
                IEnumerable<Sheet> tempsheets = tempspreadsheetdoc.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sec.Sheet);
                string temprelationshipId = tempsheets.First().Id.Value;
                WorksheetPart tempworksheetPart = (WorksheetPart)tempspreadsheetdoc.WorkbookPart.GetPartById(temprelationshipId);
                Worksheet tempworkSheet = tempworksheetPart.Worksheet;
                SheetData tempsheetData = tempworkSheet.GetFirstChild<SheetData>();
                WorkbookPart tempworkbookpart = tempspreadsheetdoc.WorkbookPart;
                SharedStringTablePart tempsharedstringtablepart = tempworkbookpart.GetPartsOfType<SharedStringTablePart>().First();
                //sheetData.AddChild(tempworkSheet.GetFirstChild<SheetData>());
                 foreach(Row r in tempsheetData.Descendants<Row>().Where(x => x.RowIndex >= sec.CopyFromRange.RowStart && x.RowIndex <= sec.CopyFromRange.RowEnd))
                {
                    Row row = sheetData.Elements<Row>().First(row => row.RowIndex == r.RowIndex);
                    foreach (Cell c in r.Descendants<Cell>().Where(x =>
                    {
                        string val = x.CellReference;
                        char[] trimval = { '1', '2', '3', '4', '5', '6', '7', '8', '9', '0' };
                        val = val.Trim(trimval);
                        int ci = 0;
                        val = val.ToUpper();
                        for (int ix = 0; ix < val.Length && val[ix] >= 'A'; ix++)
                        { ci = (ci * 26) + ((int)val[ix] - 64); }
                        if (ci >= sec.CopyFromRange.ColumnStart && ci <= sec.CopyFromRange.ColumnEnd)
                        { return true; }
                        else
                        { return false; }
                    }))
                    {
                        Cell newCell = new Cell();
                        newCell = (Cell)c.Clone();
                        Cell refCell = row.Elements<Cell>().First(cell => cell.CellReference == c.CellReference);
                        // check cell datatype
                        string value;
                        if (newCell != null)
                        {
                            value = newCell.InnerText;


                            //if null then number if anything else other type
                            if (newCell.DataType != null)
                            {
                                switch (newCell.DataType.Value)
                                {
                                    case CellValues.SharedString:
                                        char[] trimval = { '1', '2', '3', '4', '5', '6', '7', '8', '9', '0' };
                                        string hold = tempsharedstringtablepart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                                        int index = InsertSharedStringItem(tempsharedstringtablepart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText, sharedstringtablepart);
                                        
                                        refCell = InsertCellInWorksheet(newCell.CellReference.ToString().Trim(trimval), r.RowIndex , worksheetPart);
                                        refCell.CellValue = new CellValue(index.ToString());
                                        refCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
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

                                    case CellValues.Date:
                                        break;

                                    case CellValues.String:
                                        break;
                                }
                            }
                            else
                            {

                                CellValue cvalue = new CellValue(value);
                                if (newCell.CellFormula != null)
								{
                                    throw new Exception("Formula found in data.");
                                }
								else
								{
                                    refCell.CellValue = cvalue;
                                    
                                }
                            }
                        }
                    }
                }
                tempspreadsheetdoc.Close();
            }
            newspreadsheetdoc.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
            newspreadsheetdoc.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
            newspreadsheetdoc.Save();
            newspreadsheetdoc.Close();
            return NewFileLocation;
        }


        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }


    }

}
