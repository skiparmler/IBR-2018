using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using D = DocumentFormat.OpenXml.Drawing;
using System.Globalization;
using System.IO;
using System.Data;
using System.Collections;

namespace SKIIBR
{
    public class Functions
    {
        internal static void SetTableCellText(D.TableCell cell, string text, bool bold = false)
        {
            foreach(var p in cell.TextBody.Elements<D.Paragraph>())
            {
                foreach(var run in p.Elements<D.Run>())
                {
                    run.Text = new D.Text(string.Empty);
                }
            }
            var paragraph = cell.TextBody.Elements<D.Paragraph>().FirstOrDefault();
            if (paragraph != null)
            {                
                var run = paragraph.Elements<D.Run>().FirstOrDefault();
                if (run != null)
                {
                    run.Text = new D.Text(text);
                    if(bold)
                    {
                        run.GetFirstChild<D.RunProperties>().Bold = true;
                    }
                }
            }
        }
        internal static string GetTextFromSharedTable(SpreadsheetDocument extXls, int input)
        {
            return extXls.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements[input].InnerText.Trim();
        }

        internal static string GetTextFromSharedTable(SpreadsheetDocument extXls, Cell input)
        {
            if (input != null && input.DataType != null && input.DataType == CellValues.SharedString)
            {
                return extXls.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements[int.Parse(input.CellValue.InnerText)].InnerText.Trim();
            }
            else
            {
                return input?.CellValue?.Text ?? string.Empty;
            }
        }

        internal static DataSet GetSheetAsTables(string selectedPath, string fileName, string sheetName)
        {
            DataSet ibrData = new DataSet();

            using (FileStream fs = new FileStream(selectedPath + "\\Output\\" + fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument extXls = SpreadsheetDocument.Open(fs, true))
                {
                    var sheet = extXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(x => x.Name.ToString() == sheetName);
                    if (sheet != null)
                    {
                        WorksheetPart wsPart = (WorksheetPart)extXls.WorkbookPart.GetPartById(sheet.Id);
                        Worksheet ws = wsPart.Worksheet;
                        var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                        if (sheetData != null)
                        {
                            bool isHeader = true;
                            var rows = sheetData.Elements<Row>();
                            DataTable dataTable = new DataTable();
                            uint currentRowIndex = 0;
                            uint rowIndexDiff = 0;
                            for(int rowIndex = 0; rowIndex < rows.Count(); rowIndex++)
                            {
                                var currentRow = rows.ElementAt(rowIndex);
                                                                
                                if (isHeader)
                                {
                                    string tableName = GetTextFromSharedTable(extXls, currentRow.GetFirstChild<Cell>());
                                    if(ibrData.Tables.Contains(tableName))
                                    {
                                        tableName = "!" + tableName;
                                    }
                                    dataTable.TableName = tableName;
                                    isHeader = false;
                                    rowIndex++;
                                    foreach(var cell in rows.ElementAt(rowIndex).Elements<Cell>())
                                    {
                                        dataTable.Columns.Add(GetTextFromSharedTable(extXls, cell));
                                    }
                                }
                                else
                                {                                    
                                    var firstCellInRow = currentRow.GetFirstChild<Cell>();
                                    if (firstCellInRow != null && firstCellInRow.CellValue != null)
                                    {
                                        List<string> values = new List<string>();

                                        values.Add(GetTextFromSharedTable(extXls, firstCellInRow));
                                        var cells = currentRow.Elements<Cell>();
                                        for (int i = 1; i < cells.Count(); i++)
                                        {
                                            if (cells.ElementAt(i).CellValue != null)
                                            {
                                                values.Add(cells.ElementAt(i).CellValue.Text);
                                            }
                                        }

                                        var newRow = dataTable.NewRow();
                                        newRow.ItemArray = values.ToArray();
                                        dataTable.Rows.Add(newRow);

                                        var nextRow = rows.ElementAtOrDefault(rowIndex + 1);
                                        if(nextRow != null && nextRow.RowIndex - currentRow.RowIndex > 1)
                                        {
                                            ibrData.Tables.Add(dataTable);
                                            isHeader = true;
                                            dataTable = new DataTable();
                                        }
                                    }
                                    else
                                    {
                                        ibrData.Tables.Add(dataTable);
                                        isHeader = true;
                                        dataTable = new DataTable();
                                    }
                                }
                            }
                            if(dataTable.Rows.Count > 0)
                            {
                                ibrData.Tables.Add(dataTable);
                            }
                        }
                    }
                }
            }

            return ibrData;
        }

        internal static void UpdateManifestTable(Slide slide, string dataName, string selectedPath, int headerRowIndex, string identifier, List<string> actors, int nrOfColsFromStart)
        {
            List<string> aspects = new List<string>();

            var graphicsFrames = slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>();
            //GraphicFrame frameWithTable = graphicsFrames.FirstOrDefault(x => x.InnerXml.Contains(identifier));
            GraphicFrame frameWithTable = graphicsFrames.FirstOrDefault(x => x.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == identifier);
            var table = frameWithTable.GetFirstChild<D.Graphic>()?.GraphicData.GetFirstChild<D.Table>();
            if (table != null)
            {
                var tableRows = table.Elements<D.TableRow>();
                var rowHeights = tableRows.Select(x => x.Height);
                var height = rowHeights.Sum(x => long.Parse(x));
                var rowHeader = tableRows.ElementAt(headerRowIndex).CloneNode(true);
                var rowToUse = tableRows.ElementAt(headerRowIndex + 1).CloneNode(true);

                while (table.Elements<D.TableRow>().Count() > headerRowIndex)
                {
                    table.RemoveChild(table.Elements<D.TableRow>().Last());
                }
                if (headerRowIndex > 0)
                {
                    var firstRow = tableRows.First();
                    while (firstRow.Elements<D.TableCell>().Count() > 2)
                    {
                        firstRow.RemoveChild(firstRow.Elements<D.TableCell>().Last());
                    }
                }
                while (rowHeader.Elements<D.TableCell>().Count() > 2)
                {
                    rowHeader.RemoveChild(rowHeader.Elements<D.TableCell>().Last());
                }
                while (rowToUse.Elements<D.TableCell>().Count() > 2)
                {
                    rowToUse.RemoveChild(rowToUse.Elements<D.TableCell>().Last());
                }

                var headerCells = rowHeader.Elements<D.TableCell>();

                table.AppendChild(rowHeader);

                using (FileStream fs = new FileStream(selectedPath + "\\Output\\manifests.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (SpreadsheetDocument extXls = SpreadsheetDocument.Open(fs, false))
                    {
                        var sheet = extXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(x => x.Name == dataName);
                        if (sheet != null)
                        {
                            Worksheet ws = ((WorksheetPart)extXls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                            var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                            if (sheetData != null)
                            {
                                var rows = sheetData.Elements<Row>();
                                var headers = rows.First().Elements<Cell>();
                                var headerExtList = rowHeader.GetFirstChild<D.ExtensionList>();
                                int headerIndex = 1;
                                foreach (var cell in headers.Skip(1))
                                {
                                    string s = GetTextFromSharedTable(extXls, int.Parse(cell.InnerText));

                                    aspects.Add(s);

                                    if ((headerIndex + 1) > headerCells.Count())
                                    {
                                        var newHeaderCell = headerCells.ElementAt(1).CloneNode(true);
                                        SetTableCellText((D.TableCell)newHeaderCell, s);
                                        rowHeader.InsertBefore(newHeaderCell, headerExtList);

                                        if (headerRowIndex > 0)
                                        {
                                            var firstRowCells = tableRows.First();
                                            var firstRowExtList = firstRowCells.GetFirstChild<D.ExtensionList>();
                                            
                                            var cellToCopy = firstRowCells.Elements<D.TableCell>().ElementAt(1);
                                            var newCell = cellToCopy.CloneNode(true);
                                            tableRows.First().InsertBefore(newCell, firstRowExtList);
                                        }
                                    }
                                    else
                                    {
                                        var c = headerCells.ElementAt(headerIndex);
                                        SetTableCellText(c, s);
                                    }
                                    headerIndex++;
                                }

                                if (actors.Count == 0)
                                {
                                    foreach (var row in rows.Skip(1))
                                    {
                                        string actor = GetTextFromSharedTable(extXls, int.Parse(row.Elements<Cell>().ElementAt(1).InnerText));
                                        actors.Add(actor);
                                    }
                                }

                                foreach (var actor in actors)
                                {
                                    foreach (var row in rows)
                                    {
                                        if (GetTextFromSharedTable(extXls, row.Elements<Cell>().ElementAtOrDefault(1)) == actor)
                                        {
                                            var excelCells = row.Elements<Cell>();
                                            var newRow = (D.TableRow)rowToUse.CloneNode(true);
                                            var cells = newRow.Elements<D.TableCell>();
                                            var extList = newRow.GetFirstChild<D.ExtensionList>();
                                            int index = 1;
                                            //sätt företagsnamnet
                                            if (actor == "industry")
                                                SetTableCellText(cells.First(), "Branschen", true);
                                            else
                                                SetTableCellText(cells.First(), actor);
                                            
                                            

                                            foreach (var cell in excelCells.Skip(2))
                                            {
                                                string s = double.Parse(cell.InnerText.ToString(), CultureInfo.InvariantCulture).ToString("n1");
                                                if (s == "0,0")
                                                    s = "-";
                                                if ((index + 1) > cells.Count())
                                                {
                                                    var newCell = cells.ElementAt(1).CloneNode(true);
                                                    SetTableCellText((D.TableCell)newCell, s, actor == "industry");
                                                    newRow.InsertBefore((D.TableCell)newCell, extList);
                                                }
                                                else
                                                {
                                                    var c = cells.ElementAt(index);
                                                    SetTableCellText(c, s, actor == "industry");
                                                }
                                                index++;
                                            }
                                            table.AppendChild(newRow);
                                            break;
                                        }
                                    }

                                }
                            }
                        }
                    }
                }
                bool resizeCols = false;
                var colCount = table.GetFirstChild<D.TableRow>().Elements<D.TableCell>().Count();

                if (headerRowIndex > 0)
                {
                    tableRows.First().Elements<D.TableCell>().First().GridSpan.Value = colCount;
                }

                var tableGrid = table.GetFirstChild<D.TableGrid>();
                if (table.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() < colCount)
                {
                    while (table.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() < colCount)
                    {
                        D.GridColumn newCol = (D.GridColumn)table.GetFirstChild<D.TableGrid>().GetFirstChild<D.GridColumn>().CloneNode(true);
                        tableGrid.AppendChild(newCol);
                    }
                    resizeCols = true;
                }
                else if (table.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() > colCount)
                {
                    while (table.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() > colCount)
                    {
                        tableGrid.RemoveChild(tableGrid.Elements<D.GridColumn>().Last());
                        resizeCols = true;
                    }
                }

                if (true)
                {
                    long firstColWidth = (long)(2.5 * 360000);
                    var aCol = table.GetFirstChild<D.TableGrid>().GetFirstChild<D.GridColumn>();
                    long totalWidth = aCol.Width.Value * nrOfColsFromStart;
                    var nrOfCols = table.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count();
                    long widthPerCol = (totalWidth - firstColWidth) / (nrOfCols - 1);
                    table.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().First().Width.Value = firstColWidth; //c:a 2 cm
                    foreach (D.GridColumn col in table.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Skip(1))
                    {
                        col.Width.Value = widthPerCol;
                    }
                }

                var shapes = slide.CommonSlideData.ShapeTree.Descendants<Shape>();
                foreach (var shape in shapes)
                {
                    if (shape.InnerText.Contains("{Aspekt}: {Frågetext}"))
                    {
                        var aspectQuestions = GetAspectQuestions(selectedPath);
                        var firstParagraph = shape.TextBody.Elements<D.Paragraph>().FirstOrDefault(x => x.InnerText.Contains("{Aspekt}: {Frågetext}"));
                        var newParagraph = firstParagraph.CloneNode(true);

                        bool isFirst = true;
                        foreach (var s in aspects)
                        {
                            D.Paragraph paragraph = null;
                            if(isFirst)
                            {
                                paragraph = firstParagraph;
                                isFirst = false;
                            }
                            else
                            {
                                paragraph = (D.Paragraph)newParagraph.CloneNode(true);
                                //shape.TextBody.InsertAfter(paragraph, firstParagraph);
                                shape.TextBody.AppendChild(paragraph);
                            }
                            var runs = paragraph.Elements<D.Run>();
                            foreach (var run in runs)
                            {
                                if (run.InnerText.Contains("{Aspekt}"))
                                {
                                    run.Text = new D.Text(s + ": ");
                                }
                                else if (run.InnerText.Contains("{Frågetext}"))
                                {
                                    string text = aspectQuestions.ContainsKey(s) ? aspectQuestions[s] : s;
                                    
                                    run.Text = new D.Text(text);
                                }
                            }
                        }

                    }
                }
            }
        }

        internal static Dictionary<string, string> GetAspectQuestions(string selectedPath, int index)
        {
            Dictionary<string, string> aspectQuestions = new Dictionary<string, string>();

            string[] lines = File.ReadAllLines(selectedPath + "\\Input\\qtext.txt", Encoding.Default);

            foreach (var line in lines)
            {
                string[] aq = line.Split(new char[] { '\t' });
                if (!aspectQuestions.ContainsKey(aq.First()))
                {
                    aspectQuestions.Add(aq.First(), aq.ElementAt(index));
                }
            }

            return aspectQuestions;
        }

        internal static Dictionary<string, string> GetAspectQuestions(string selectedPath)
        {
            return GetAspectQuestions(selectedPath, 1);
        }

        internal static void UpdateResultatAspektChart(Slide slide, List<Tuple<string, double>> values, string displayName, string bransch, string segment, string year)
        {
            var chartPart = slide.SlidePart.ChartParts.FirstOrDefault();
            if (chartPart != null)
            {
                D.Charts.Chart chart = chartPart.ChartSpace.GetFirstChild<D.Charts.Chart>();
                var paragraphs = chart.Title.ChartText.RichText.Elements<D.Paragraph>();
                foreach (var paragraph in paragraphs)
                {
                    if (paragraph.InnerText.Contains("{"))
                    {
                        var runs = paragraph.Elements<D.Run>();
                        foreach (D.Run run in runs)
                        {
                            run.Text = new D.Text(run.Text.Text.Replace("{Resultatvariabel}", displayName).Replace("{BRANSCH}", bransch).Replace("{Segment}", segment).Replace("{ÅR}", year));
                        }
                    }
                }
                var plotArea = chart.GetFirstChild<D.Charts.PlotArea>();
                var barChart = plotArea.GetFirstChild<D.Charts.BarChart>();

                var numberReference = barChart.GetFirstChild<D.Charts.BarChartSeries>()?.GetFirstChild<D.Charts.Values>()?.GetFirstChild<D.Charts.NumberReference>();
                var numberRange = numberReference?.GetFirstChild<D.Charts.Formula>();
                var numberCache = numberReference?.GetFirstChild<D.Charts.NumberingCache>();
                numberCache.PointCount = new D.Charts.PointCount() { Val = (uint)values.Count };
                var numericPoints = numberCache?.Elements<D.Charts.NumericPoint>();

                var stringReference = barChart.GetFirstChild<D.Charts.BarChartSeries>()?.GetFirstChild<D.Charts.CategoryAxisData>()?.GetFirstChild<D.Charts.StringReference>();
                var textRange = stringReference?.GetFirstChild<D.Charts.Formula>();
                var textCache = stringReference?.GetFirstChild<D.Charts.StringCache>();
                textCache.PointCount = new D.Charts.PointCount() { Val = (uint)values.Count };
                var stringPoints = textCache?.Elements<D.Charts.StringPoint>();

                EmbeddedPackagePart embeddedExcel = chartPart.EmbeddedPackagePart;
                if (embeddedExcel != null)
                {
                    using (Stream str = embeddedExcel.GetStream())
                    {
                        using (SpreadsheetDocument xls = SpreadsheetDocument.Open(str, true))
                        {
                            string sheetName = "Liggande stapeldiagram";
                            var sheet = xls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name.HasValue && s.Name.Value == sheetName);
                            if (sheet != null)
                            {
                                Worksheet ws = ((WorksheetPart)xls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                                var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                                if (sheetData != null)
                                {
                                    int rowIndex = 3;

                                    //B3 första namn-cell, C3 första värde-cell
                                    foreach (var nameValue in values.OrderBy(x => x.Item2))
                                    {
                                        var currentRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
                                        if (currentRow == null)
                                        {
                                            currentRow = sheetData.InsertAt(new Row() { RowIndex = (uint)rowIndex }, sheetData.Elements<Row>().Count());
                                        }
                                        if (currentRow != null)
                                        {
                                            var nameCell = currentRow.Elements<Cell>().FirstOrDefault(x => x.CellReference.InnerText == "B" + rowIndex);
                                            if (nameCell == null)
                                            {
                                                nameCell = currentRow.InsertAt(new Cell() { CellReference = "B" + rowIndex }, 0);
                                            }
                                            if (nameCell != null)
                                            {
                                                string actor = nameValue.Item1;
                                                if (actor == "industry")
                                                    actor = "Branschen";
                                                nameCell.CellValue = new CellValue(actor);
                                                nameCell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);
                                            }
                                            var valueCell = currentRow.Elements<Cell>().FirstOrDefault(x => x.CellReference.InnerText == "C" + rowIndex);
                                            if (valueCell == null)
                                            {
                                                valueCell = currentRow.InsertAt(new Cell() { CellReference = "C" + rowIndex }, 1);
                                            }
                                            if (valueCell != null)
                                            {
                                                valueCell.CellValue = new CellValue(nameValue.Item2.ToString("n2").Replace(",", "."));
                                            }
                                        }
                                        rowIndex++;
                                    }

                                    textRange.Text = textRange.Text.Substring(0, textRange.Text.Length - 1) + (rowIndex - 1);
                                    numberRange.Text = numberRange.Text.Substring(0, numberRange.Text.Length - 1) + (rowIndex - 1);

                                    //uppdatera cachen
                                    int ncIndex = 0;
                                    textCache.RemoveAllChildren();
                                    numberCache.RemoveAllChildren();
                                    foreach (var v in values.OrderBy(x => x.Item2))
                                    {
                                        D.Charts.NumericPoint currentNP = numericPoints.ElementAtOrDefault(ncIndex);
                                        if (currentNP == null)
                                        {
                                            currentNP = new D.Charts.NumericPoint();
                                            currentNP.NumericValue = new D.Charts.NumericValue(v.Item2.ToString("n1").Replace(",", "."));
                                            currentNP.Index = (uint)ncIndex;
                                            numberCache.AppendChild(currentNP);
                                        }
                                        else
                                        {
                                            currentNP.NumericValue = new D.Charts.NumericValue(v.Item2.ToString("n1").Replace(",", "."));
                                        }
                                        string actor = v.Item1;
                                        if (actor == "industry")
                                            actor = "Branschen";
                                        D.Charts.StringPoint currentSP = stringPoints.ElementAtOrDefault(ncIndex);
                                        if (currentSP == null)
                                        {
                                            currentSP = new D.Charts.StringPoint();
                                            currentSP.InnerXml = $"<c:v xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">{System.Security.SecurityElement.Escape(actor)}</c:v>";
                                            currentSP.Index = (uint)ncIndex;
                                            textCache.AppendChild(currentSP);
                                        }
                                        else
                                        {
                                            currentSP.InnerXml = $"<c:v xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">{System.Security.SecurityElement.Escape(actor)}</c:v>";
                                        }
                                        ncIndex++;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        internal static void UpdateTableText(Slide slide, string oldText, string newText)
        {
            var gf = slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>();
            var tableGFs = gf.Where(x => x.Descendants<D.Table>().Count() > 0);
            foreach (var tableGF in tableGFs)
            {
                var tables = tableGF.Descendants<D.Table>();
                foreach (var table in tables)
                {
                    foreach (var row in table.Elements<D.TableRow>())
                    {
                        foreach (var cell in row.Elements<D.TableCell>())
                        {
                            if (cell.InnerText.Contains(oldText))
                            {
                                foreach (var run in cell.TextBody.GetFirstChild<D.Paragraph>().Elements<D.Run>())
                                {
                                    if (run.Text.Text.Contains(oldText))
                                    {
                                        run.Text = new D.Text(run.Text.Text.Replace(oldText, newText));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        internal static List<Tuple<string, List<double>>> GetHistoricValues(string dataName, string selectedPath, int currentYear, List<int> years)
        {
            List<Tuple<string, List<double>>> historicValues = new List<Tuple<string, List<double>>>();

            using (FileStream fs = new FileStream(selectedPath + "\\Output\\Historik.xlsx", FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument extXls = SpreadsheetDocument.Open(fs, true))
                {
                    var sheet = extXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(x => x.Name.ToString().ToLower() == dataName);
                    if (sheet != null)
                    {
                        WorksheetPart wsPart = (WorksheetPart)extXls.WorkbookPart.GetPartById(sheet.Id);
                        Worksheet ws = wsPart.Worksheet;
                        var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                        if (sheetData != null)
                        {
                            //gå igenom alla aktörer och se till att de finns med i historikfilen
                            var dataRows = sheetData.Elements<Row>().Where(x => x.RowIndex > 2);

                            var actors = GetActors(selectedPath);
                            foreach(var actor in actors)
                            {
                                var actorRow = dataRows.FirstOrDefault(x => GetTextFromSharedTable(extXls, x.GetFirstChild<Cell>()) == actor);
                                if(actorRow == null)
                                {
                                    var branschRow = dataRows.FirstOrDefault(x => GetTextFromSharedTable(extXls, x.GetFirstChild<Cell>()) == "Branschen");
                                    var newRow = (Row)dataRows.First().CloneNode(true);

                                    foreach (var cell in newRow.Elements<Cell>())
                                    {
                                        cell.CellReference.Value = cell.CellReference.Value.Replace(newRow.RowIndex.ToString(), branschRow.RowIndex.ToString());
                                        if(cell.CellReference.Value.StartsWith("B"))
                                        {
                                            cell.CellValue = new CellValue(actor);
                                            cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);
                                        }
                                        else
                                        {
                                            cell.CellValue = new CellValue("0.0");
                                        }
                                    }
                                    newRow.Elements<Cell>().Last().CellValue = new CellValue(GetValueFromLatentSummary(actor, dataName, selectedPath).ToString("n2").Replace(",", "."));

                                    newRow.RowIndex = branschRow.RowIndex;
                                    
                                    branschRow.RowIndex++;
                                    foreach (var cell in branschRow.Elements<Cell>())
                                    {
                                        cell.CellReference.Value = cell.CellReference.Value.Substring(0, 1) + branschRow.RowIndex.ToString();
                                    }
                                    sheetData.InsertBefore(newRow, branschRow);
                                }                                
                            }


                            bool addYear = false;
                            var historyRows = sheetData.Elements<Row>();
                            bool isData = false;

                            var headerRow = historyRows.FirstOrDefault(x => x.RowIndex == 2);
                            if (headerRow != null)
                            {
                                var lastCell = headerRow.Elements<Cell>().Last();
                                var lastHeader = GetTextFromSharedTable(extXls, lastCell);
                                if (!lastHeader.Contains((currentYear).ToString()))
                                {
                                    var newCell = (Cell)headerRow.Elements<Cell>().Last().CloneNode(true);
                                    newCell.CellValue = new CellValue(currentYear.ToString());
                                    newCell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);
                                    string reference = newCell.CellReference;
                                    char letter = reference[0];
                                    newCell.CellReference = newCell.CellReference.Value.Replace(letter.ToString(), (++letter).ToString());
                                    headerRow.AppendChild(newCell);
                                    addYear = true;
                                }
                            }
                            //Vi börjar på rad 3
                            foreach (var row in historyRows)
                            {
                                if (!isData)
                                {
                                    if(row.RowIndex == 2)
                                    {                                        
                                        foreach(var cell in row.Elements<Cell>())
                                        {
                                            string cellText = GetTextFromSharedTable(extXls, cell);
                                            var match = Regex.Match(cellText, "\\d{4}");
                                            if(match.Success)
                                            {
                                                years.Add(int.Parse(match.Value));
                                            }
                                        }
                                    }
                                    if (row.GetFirstChild<Cell>().CellReference.Value.Contains("3"))
                                    {
                                        isData = true;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }

                                if (addYear)
                                {
                                    string actorToFetch = GetTextFromSharedTable(extXls, row.GetFirstChild<Cell>());
                                    if (actorToFetch == "Branschen")
                                        actorToFetch = "industry";
                                    double newValue = GetValueFromLatentSummary(actorToFetch, dataName, selectedPath);
                                    var newCell = (Cell)row.Elements<Cell>().Last().CloneNode(true);
                                    newCell.CellValue = new CellValue(newValue.ToString("n2").Replace(",", "."));
                                    string reference = newCell.CellReference;
                                    char letter = reference[0];
                                    newCell.CellReference = newCell.CellReference.Value.Replace(letter.ToString(), (++letter).ToString());
                                    row.AppendChild(newCell);                                    
                                }


                                var excelCells = row.Elements<Cell>();

                                int start = excelCells.Count() - 5;
                                if (start < 1)
                                {
                                    start = 1;
                                }

                                string actorName = GetTextFromSharedTable(extXls, row.GetFirstChild<Cell>());
                                if (actorName != string.Empty)
                                {
                                    Tuple<string, List<double>> value = new Tuple<string, List<double>>(actorName, new List<double>());

                                    for (int i = start; i < excelCells.Count(); i++)
                                    {
                                        var excelCell = excelCells.ElementAt(i);
                                        if (excelCell.InnerText.Trim() != string.Empty)
                                        {
                                            value.Item2.Add(double.Parse(excelCell.InnerText, CultureInfo.InvariantCulture));
                                        }
                                        else
                                        {
                                            var headerCell = headerRow.Elements<Cell>().FirstOrDefault(x => x.CellReference.Value.Contains(excelCell.CellReference.Value.Substring(0, 1)));
                                            if(headerCell == null || headerCell.CellValue == null)
                                            {
                                                continue;
                                            }
                                            value.Item2.Add(0.0);
                                        }
                                    }

                                    historicValues.Add(value);
                                }

                            }
                            if (addYear)
                            {
                                if (!years.Contains(currentYear))
                                {
                                    years.Add(currentYear);
                                }
                                wsPart.Worksheet.Save();
                                extXls.Close();
                            }
                        }
                    }
                }
            }

            return historicValues;
        }

        internal static void UpdateHistoryTable(Slide slide, string dataName, List<Tuple<string, List<double>>> values, string selectedPath, bool visaFelmarginal, bool visaForklaringsniva, int headerRowIndex, string identifier, List<int> years)
        {
            bool addedHeaderCells = false;
            bool addedHeaderForklaringsNiva = false;
            bool addedHeaderFelmarginal = false;

            var kundaspektHistoryGF = slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>();
            
            //GraphicFrame kundaspektHistoryTable = kundaspektHistoryGF.FirstOrDefault(x => x.InnerXml.Contains(identifier));
            GraphicFrame kundaspektHistoryTable = kundaspektHistoryGF.FirstOrDefault(x => x.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == identifier);
            var kaTable = kundaspektHistoryTable.GetFirstChild<D.Graphic>()?.GraphicData.GetFirstChild<D.Table>();
            D.TableRow headerTableRow = (D.TableRow)kaTable.Elements<D.TableRow>().ElementAt(headerRowIndex).CloneNode(true);
            D.TableCell headerCell = (D.TableCell)headerTableRow.ElementAt(1).CloneNode(true);
            var headerExtList = headerTableRow.GetFirstChild<D.ExtensionList>();
            D.TableRow newTableRow = (D.TableRow)kaTable.Elements<D.TableRow>().ElementAt(headerRowIndex + 1).CloneNode(true);

            while (kaTable.Elements<D.TableRow>().Count() > headerRowIndex)
            {
                kaTable.RemoveChild(kaTable.Elements<D.TableRow>().Last());
            }

            while(headerTableRow.Elements<D.TableCell>().Count() > 1)
            {
                headerTableRow.RemoveChild(headerTableRow.Elements<D.TableCell>().Last());
            }

            kaTable.AppendChild(headerTableRow);

            if (kaTable != null)
            {
                foreach (var actorData in values)
                {
                    var rowToAdd = newTableRow.CloneNode(true);
                    var cellToClone = rowToAdd.Elements<D.TableCell>().ElementAt(1).CloneNode(true);
                    while(rowToAdd.Elements<D.TableCell>().Count() > 1)
                    {
                        rowToAdd.RemoveChild(rowToAdd.Elements<D.TableCell>().Last());
                    }                                
                    
                    var actorName = string.Empty;
                    int lastI = 0;
                    var extList = rowToAdd.GetFirstChild<D.ExtensionList>();

                    actorName = actorData.Item1;
                    SetTableCellText(rowToAdd.GetFirstChild<D.TableCell>(), actorName, actorName == "Branschen"); 

                    foreach(var value in actorData.Item2)
                    {
                        if(!addedHeaderCells)
                        {
                            headerTableRow.InsertBefore((D.TableCell)headerCell.CloneNode(true), headerExtList);
                        }
                        var currentTableCell = (D.TableCell)cellToClone.CloneNode(true);

                        string s = value.ToString("n1");
                        if (s == "0,0")
                            s = "-";

                        SetTableCellText(currentTableCell, s, actorName == "Branschen");
                        rowToAdd.InsertBefore(currentTableCell, extList);
                        lastI++;
                    }
                    addedHeaderCells = true;

                    if (actorName != "Branschen" && (visaFelmarginal || visaForklaringsniva) && File.Exists(selectedPath + "\\Output\\" + actorName + ".xlsx"))
                    {
                        //läs in Felmarginal och Förklaringsgrad också
                        using (FileStream actorFs = new FileStream(selectedPath + "\\Output\\" + actorName + ".xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            using (SpreadsheetDocument actorXls = SpreadsheetDocument.Open(actorFs, false))
                            {
                                if (visaFelmarginal)
                                {
                                    if (!addedHeaderFelmarginal)
                                    {
                                        var felmarginalHeader = (D.TableCell)headerCell.CloneNode(true);
                                        SetTableCellText(felmarginalHeader, "Felmarginal");
                                        headerTableRow.InsertBefore(felmarginalHeader, headerExtList);
                                        addedHeaderFelmarginal = true;
                                    }
                                    var latentActorSheet = actorXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(x => x.Name.ToString().ToLower() == "latent summary");
                                    if (latentActorSheet != null)
                                    {
                                        Worksheet latentWs = ((WorksheetPart)actorXls.WorkbookPart.GetPartById(latentActorSheet.Id)).Worksheet;
                                        var latentSheetData = latentWs.Elements<SheetData>().FirstOrDefault();
                                        if (latentSheetData != null)
                                        {
                                            foreach (var latentRow in latentSheetData.Elements<Row>())
                                            {
                                                string rowName = GetTextFromSharedTable(actorXls, int.Parse(latentRow.GetFirstChild<Cell>().InnerText));
                                                if (rowName.Contains(dataName))
                                                {
                                                    var tableCells = rowToAdd.Elements<D.TableCell>();
                                                    var felmarginal = double.Parse(latentRow.Elements<Cell>().ElementAt(13).InnerText, CultureInfo.InvariantCulture) * 1.96;
                                                    var currentTableCell = tableCells.ElementAtOrDefault(lastI + 1);
                                                    if (currentTableCell == null)
                                                    {
                                                        var newCell = tableCells.ElementAt(1).CloneNode(true);
                                                        currentTableCell = rowToAdd.InsertBefore((D.TableCell)newCell, extList);
                                                    }
                                                    SetTableCellText(currentTableCell, felmarginal.ToString("n2"));
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }

                                if (visaForklaringsniva)
                                {
                                    if (!addedHeaderForklaringsNiva)
                                    {
                                        var forklaringsNivaHeader = (D.TableCell)headerCell.CloneNode(true);
                                        SetTableCellText(forklaringsNivaHeader, "Förklaringsnivå");
                                        headerTableRow.InsertBefore(forklaringsNivaHeader, headerExtList);
                                        addedHeaderForklaringsNiva = true;
                                    }
                                    var innerSummarySheet = actorXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(x => x.Name.ToString().ToLower() == "inner_summary");
                                    if (innerSummarySheet != null)
                                    {
                                        Worksheet innerSummaryWS = ((WorksheetPart)actorXls.WorkbookPart.GetPartById(innerSummarySheet.Id)).Worksheet;
                                        var innerSummarySheetData = innerSummaryWS.Elements<SheetData>().FirstOrDefault();
                                        if (innerSummarySheetData != null)
                                        {
                                            foreach (var innerSummaryRow in innerSummarySheetData.Elements<Row>())
                                            {
                                                string rowName = GetTextFromSharedTable(actorXls, int.Parse(innerSummaryRow.GetFirstChild<Cell>().InnerText));
                                                if (rowName.Contains(dataName))
                                                {
                                                    var tableCells = rowToAdd.Elements<D.TableCell>();
                                                    var forklaringsgrad = (double.Parse(innerSummaryRow.Elements<Cell>().ElementAt(2).InnerText, CultureInfo.InvariantCulture)).ToString("p0");
                                                    var currentTableCell = tableCells.ElementAtOrDefault(lastI + 2);
                                                    if (currentTableCell == null)
                                                    {
                                                        var newCell = tableCells.ElementAt(1).CloneNode(true);
                                                        currentTableCell = rowToAdd.InsertBefore((D.TableCell)newCell, extList);
                                                    }

                                                    SetTableCellText(currentTableCell, forklaringsgrad);
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        //ta bort felmarginal och förklaringsgrad för branschen
                        var felmarginalCell = rowToAdd.Elements<D.TableCell>().ElementAtOrDefault(lastI + 1);
                        if (felmarginalCell == null)
                        {
                            var newCell = (D.TableCell)rowToAdd.Elements<D.TableCell>().ElementAt(1).CloneNode(true);
                            felmarginalCell = rowToAdd.InsertBefore(newCell, extList);
                            
                        }

                        SetTableCellText(felmarginalCell, " ");

                        var forklaringsgradCell = rowToAdd.Elements<D.TableCell>().ElementAtOrDefault(lastI + 2);
                        if (forklaringsgradCell == null)
                        {
                            var newCell = (D.TableCell)rowToAdd.Elements<D.TableCell>().ElementAt(1).CloneNode(true);
                            forklaringsgradCell = rowToAdd.InsertBefore(newCell, extList);                            
                        }
                        SetTableCellText(forklaringsgradCell, " ");
                    }
                    kaTable.AppendChild(rowToAdd);
                }

                var headerCells = headerTableRow.Elements<D.TableCell>();
                for (int i = 0; i < years.Count; i++)
                {
                    SetTableCellText(headerCells.ElementAt(i + 1), years.ElementAt(i).ToString());
                }

            }
            if (true)
            {
                int cellCount = headerTableRow.Elements<D.TableCell>().Count();

                var gridColumns = kaTable.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>();
                while(gridColumns.Count() > cellCount)
                {
                    kaTable.GetFirstChild<D.TableGrid>().RemoveChild(gridColumns.Last());
                }

                long firstColWidth = (long)(2.5 * 360000); //c:a 2,5 cm
                long endingCellsWidth = (long)(2 * 360000);
                var aCol = kaTable.GetFirstChild<D.TableGrid>().GetFirstChild<D.GridColumn>();
                long totalWidth = aCol.Width.Value * 6;
                var nrOfCols = kaTable.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count();
                long widthPerCol = (totalWidth - firstColWidth) / (nrOfCols - 1);
                kaTable.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().First().Width.Value = firstColWidth;
                var columns = kaTable.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Skip(1);
                foreach (D.GridColumn col in columns)
                {                    
                    col.Width.Value = widthPerCol;
                }
                if(visaFelmarginal && visaForklaringsniva)
                {
                    columns.Last().Width.Value = endingCellsWidth;
                    columns.ElementAt(columns.Count() - 2).Width.Value = endingCellsWidth;
                }
                else if(visaFelmarginal || visaForklaringsniva)
                {
                    columns.Last().Width.Value = endingCellsWidth;
                }
            }
        }

        internal static List<Tuple<string, double>> GetAspectForActors(string dataName, string selectedPath)
        {
            List<Tuple<string, double>> allValues = new List<Tuple<string, double>>();
            using (FileStream fs = new FileStream(selectedPath + "\\Output\\Latent_summary.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument extXls = SpreadsheetDocument.Open(fs, false))
                {
                    var sheet = extXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(x => x.Name == "Latent Summary");
                    if (sheet != null)
                    {
                        Worksheet ws = ((WorksheetPart)extXls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                        var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                        if (sheetData != null)
                        {
                            string colRef = "A";
                            var firstRow = sheetData.Elements<Row>().FirstOrDefault();
                            Cell headerCell = null;
                            if (firstRow != null)
                            {
                                headerCell = firstRow.Elements<Cell>().FirstOrDefault(x => GetTextFromSharedTable(extXls, int.Parse(x.InnerText)) == dataName);
                                if (headerCell != null)
                                {
                                    var cellReference = headerCell.CellReference;
                                    colRef = cellReference.InnerText.Substring(0, 1);
                                }
                            }

                            foreach (var row in sheetData.Elements<Row>().Skip(1))
                            {
                                string name = GetTextFromSharedTable(extXls, int.Parse(row.Elements<Cell>().FirstOrDefault(x => x.CellReference.Value.Contains("B"))?.InnerText));
                                double value = double.Parse(row.Elements<Cell>().FirstOrDefault(x => x.CellReference.Value.Contains(colRef))?.InnerText, CultureInfo.InvariantCulture);

                                allValues.Add(new Tuple<string, double>(name, value));
                            }
                        }
                    }
                }
            }

            return allValues;
        }

        internal static List<Tuple<string, double>> GetCoefs(string actor, string aspect, string selectedPath)
        {
            List<Tuple<string, double>> coefs = new List<Tuple<string, double>>();

            using (FileStream fs = new FileStream(selectedPath + "\\Output\\" + actor + ".xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument extXls = SpreadsheetDocument.Open(fs, false))
                {
                    var sheet = extXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(x => x.Name == "path_coefs");
                    if (sheet != null)
                    {
                        Worksheet ws = ((WorksheetPart)extXls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                        var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                        if (sheetData != null)
                        {
                            var rows = sheetData.Elements<Row>();
                            var header = rows.First();
                            var headerCells = header.Elements<Cell>();
                            var currentColumn = headerCells.FirstOrDefault(x => GetTextFromSharedTable(extXls, x).Contains($"path_coefs.{aspect}")).CellReference.Value.Substring(0, 1);
                            
                            foreach(var row in rows.Skip(1))
                            {
                                var cells = row.Elements<Cell>();
                                var name = GetTextFromSharedTable(extXls, cells.First());
                                var tempValue = cells.FirstOrDefault(x => x.CellReference.Value.StartsWith(currentColumn)).InnerText;
                                var value = double.Parse(tempValue, CultureInfo.InvariantCulture);

                                coefs.Add(new Tuple<string, double>(name, value));
                            }

                        }
                    }
                }
            }

            return coefs;
            
        }

        internal static List<Tuple<string, double>> GetAspectsByActor(string actor, string selectedPath)
        {
            List<Tuple<string, double>> allValues = new List<Tuple<string, double>>();
            using (FileStream fs = new FileStream(selectedPath + "\\Output\\Latent_summary.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument extXls = SpreadsheetDocument.Open(fs, false))
                {
                    var sheet = extXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(x => x.Name == "Latent Summary");
                    if (sheet != null)
                    {
                        Worksheet ws = ((WorksheetPart)extXls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                        var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                        if (sheetData != null)
                        {                            
                            var firstRow = sheetData.Elements<Row>().FirstOrDefault();
                            
                            foreach (var row in sheetData.Elements<Row>().Skip(1))
                            {
                                string name = GetTextFromSharedTable(extXls, int.Parse(row.Elements<Cell>().FirstOrDefault(x => x.CellReference.Value.Contains("B"))?.InnerText));
                                if(name.ToLower() == actor.ToLower())
                                {
                                    var cellsForActor = row.Elements<Cell>().Skip(2);
                                    foreach(var aspectCell in cellsForActor)
                                    {
                                        double value = double.Parse(aspectCell.InnerText, CultureInfo.InvariantCulture);
                                        string aspectName = GetTextFromSharedTable(extXls, firstRow.Elements<Cell>().FirstOrDefault(x => x.CellReference.Value.Contains(aspectCell.CellReference.Value.Substring(0, 1))));
                                        allValues.Add(new Tuple<string, double>(aspectName, value));
                                    }
                                    break;
                                }                                
                            }
                        }
                    }
                }
            }

            return allValues;
        }

        internal static void UpdateSpreadsTable(Slide slide, List<KeyValuePair<string, List<double>>> actorSpread, string displayName, string bransch, string segment, string year)
        {
            var spreadChartPart = slide.SlidePart.ChartParts.FirstOrDefault();
            if (spreadChartPart != null)
            {
                D.Charts.Chart chart = spreadChartPart.ChartSpace.GetFirstChild<D.Charts.Chart>();
                var paragraphs = chart.Title.ChartText.RichText.Elements<D.Paragraph>();
                foreach (var paragraph in paragraphs)
                {
                    if (paragraph.InnerText.Contains("{"))
                    {
                        var runs = paragraph.Elements<D.Run>();
                        foreach (D.Run run in runs)
                        {
                            run.Text = new D.Text(run.Text.Text.Replace("{Resultatvariabel}", displayName).Replace("{BRANSCH}", bransch).Replace("{Segment}", segment).Replace("{ÅR}", year));
                        }
                    }
                }
                var plotArea = chart.GetFirstChild<D.Charts.PlotArea>();
                var barChart = plotArea.GetFirstChild<D.Charts.BarChart>();

                int valueIndex = 0;
                foreach (var series in barChart.Elements<D.Charts.BarChartSeries>())
                {
                    var numberReference = series.GetFirstChild<D.Charts.Values>()?.GetFirstChild<D.Charts.NumberReference>();
                    var numberRange = numberReference?.GetFirstChild<D.Charts.Formula>();
                    var numberCache = numberReference?.GetFirstChild<D.Charts.NumberingCache>();
                    numberCache.PointCount = new D.Charts.PointCount() { Val = (uint)actorSpread.Count };
                    var numericPoints = numberCache?.Elements<D.Charts.NumericPoint>();

                    numberRange.Text = numberRange.Text.Substring(0, numberRange.Text.Length - 1) + (actorSpread.Count + 2);

                    numberCache.RemoveAllChildren();
                    int nCacheIndex = 0;
                    foreach (var v in actorSpread)
                    {
                        D.Charts.NumericPoint currentNP = numericPoints.ElementAtOrDefault(nCacheIndex);
                        if (currentNP == null)
                        {
                            currentNP = new D.Charts.NumericPoint();
                            currentNP.NumericValue = new D.Charts.NumericValue(v.Value.ElementAt(valueIndex).ToString("n2").Replace(",", "."));
                            currentNP.Index = (uint)nCacheIndex;
                            currentNP.FormatCode = "0%";
                            numberCache.AppendChild(currentNP);
                        }
                        else
                        {
                            currentNP.NumericValue = new D.Charts.NumericValue(v.Value.ElementAt(valueIndex).ToString("n2").Replace(",", "."));
                        }
                        nCacheIndex++;
                    }
                    valueIndex++;
                }

                var stringReference = barChart.GetFirstChild<D.Charts.BarChartSeries>()?.GetFirstChild<D.Charts.CategoryAxisData>()?.GetFirstChild<D.Charts.StringReference>();
                var textRange = stringReference?.GetFirstChild<D.Charts.Formula>();
                var textCache = stringReference?.GetFirstChild<D.Charts.StringCache>();
                textCache.PointCount = new D.Charts.PointCount() { Val = (uint)actorSpread.Count };
                var stringPoints = textCache?.Elements<D.Charts.StringPoint>();

                textRange.Text = textRange.Text.Substring(0, textRange.Text.Length - 1) + (actorSpread.Count + 2);

                textCache.RemoveAllChildren();
                int cacheIndex = 0;
                foreach (var v in actorSpread)
                {

                    D.Charts.StringPoint currentSP = stringPoints.ElementAtOrDefault(cacheIndex);
                    if (currentSP == null)
                    {
                        currentSP = new D.Charts.StringPoint();
                        currentSP.InnerXml = $"<c:v xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">{System.Security.SecurityElement.Escape(v.Key)}</c:v>";
                        currentSP.Index = (uint)cacheIndex;
                        textCache.AppendChild(currentSP);
                    }
                    else
                    {
                        currentSP.InnerXml = $"<c:v xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">{System.Security.SecurityElement.Escape(v.Key)}</c:v>";
                    }
                    cacheIndex++;
                }

                EmbeddedPackagePart embeddedExcel = spreadChartPart.EmbeddedPackagePart;
                if (embeddedExcel != null)
                {
                    using (Stream str = embeddedExcel.GetStream())
                    {
                        using (SpreadsheetDocument xls = SpreadsheetDocument.Open(str, true))
                        {
                            string sheetName = "nps_diagram";
                            var sheet = xls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name.HasValue && s.Name.Value == sheetName);
                            if (sheet != null)
                            {
                                Worksheet ws = ((WorksheetPart)xls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                                var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                                if (sheetData != null)
                                {
                                    int currentRowIndex = 3;

                                    string[] cols = new string[] { "B", "C", "D", "E" };

                                    foreach (var kvp in actorSpread)
                                    {
                                        var currentRow = sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex == currentRowIndex);
                                        if (currentRow == null)
                                        {
                                            currentRow = sheetData.InsertAt(new Row() { RowIndex = (uint)currentRowIndex }, sheetData.Elements<Row>().Count());
                                        }
                                        if (currentRow != null)
                                        {
                                            for (int cellIndex = 0; cellIndex < 4; cellIndex++)
                                            {
                                                var cell = currentRow.Elements<Cell>().ElementAtOrDefault(cellIndex);
                                                if (cell == null)
                                                {
                                                    cell = currentRow.AppendChild<Cell>(new Cell() { CellReference = cols[cellIndex] + currentRowIndex });
                                                }

                                                cell.CellValue = new CellValue(cellIndex == 0 ? kvp.Key : kvp.Value.ElementAt(cellIndex - 1).ToString("n2").Replace(",", "."));
                                                if (cellIndex == 0)
                                                {
                                                    cell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);
                                                }
                                                else
                                                {
                                                    cell.StyleIndex = 14U;
                                                }
                                            }
                                        }
                                        currentRowIndex++;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public static List<string> GetActors(string selectedPath)
        {
            List<string> companies = new List<string>();
            var companiesAsText = File.ReadAllLines(selectedPath + "\\Input\\q1names.txt", Encoding.Default);
            companies.AddRange(companiesAsText);

            return companies;
        }

        internal static void OrderTables(Slide slide, List<string> tableIdentifiers)
        {
            long start = (long)(2.19 * 360000);
            long offset = (long)(1.2 * 360000);
            long margin = (long)(0.5 * 360000);
            var gf = slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>();
            var tableGFs = gf.Where(x => x.Descendants<D.Table>().Count() > 0);
            foreach (var tableIdentifier in tableIdentifiers)
            {
                //GraphicFrame kundaspektHistoryTable = kundaspektHistoryGF.FirstOrDefault(x => x.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == identifier);
                var tableGF = tableGFs.FirstOrDefault(x => x.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == tableIdentifier);
                
                var table = tableGF.Descendants<D.Table>().FirstOrDefault();
                if(table != null)
                {                                          
                    tableGF.Transform.Offset.Y = start;
                    long totalHeight = table.Elements<D.TableRow>().Sum(x => long.Parse(x.Height));
                    start += totalHeight + offset + margin;                    
                }
                
            }

            var textBox = slide.CommonSlideData.ShapeTree.FirstOrDefault(x => x.InnerXml.Contains("textruta"));
            if (textBox != null)
            {
                var shapeProperty = textBox.GetFirstChild<ShapeProperties>();
                shapeProperty.Transform2D.Offset.Y = start;
            }

        }

        public static double GetValueFromLatentSummary(string actorToFetch, string dataName, string selectedPath)
        {
            double value = 0.0;

            using (FileStream fs = new FileStream(selectedPath + "\\Output\\Latent_summary.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument extXls = SpreadsheetDocument.Open(fs, false))
                {
                    var sheet = extXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>()
                        .FirstOrDefault(x => x.Name == "Latent Summary");
                    if (sheet != null)
                    {
                        Worksheet ws = ((WorksheetPart)extXls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                        var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                        if (sheetData != null)
                        {
                            string colRef = "A";
                            var firstRow = sheetData.Elements<Row>().FirstOrDefault();
                            Cell headerCell = null;
                            if (firstRow != null)
                            {
                                headerCell = firstRow.Elements<Cell>().FirstOrDefault(x => GetTextFromSharedTable(extXls, int.Parse(x.InnerText)) == dataName);
                                if (headerCell != null)
                                {
                                    var cellReference = headerCell.CellReference;
                                    colRef = cellReference.InnerText.Substring(0, 1);
                                }
                            }
                            foreach (var row in sheetData.Elements<Row>().Skip(1))
                            {
                                string name = GetTextFromSharedTable(extXls, int.Parse(row.Elements<Cell>().FirstOrDefault(x => x.CellReference.Value.Contains("B"))?.InnerText));
                                if (name == actorToFetch)
                                {
                                    value = double.Parse(row.Elements<Cell>().FirstOrDefault(x => x.CellReference.Value.Contains(colRef))?.InnerText, CultureInfo.InvariantCulture);
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            return value;
        }

        public static List<IPMA> GetIPMAData(string selectedPath, string actor)
        {
            var IPMAValues = new List<IPMA>();
            using (FileStream fs = new FileStream(selectedPath + "\\Output\\" + actor + ".xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument extXls = SpreadsheetDocument.Open(fs, false))
                {
                    var sheet = extXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>()
                        .FirstOrDefault(x => x.Name == "IPMA");
                    if (sheet != null)
                    {
                        Worksheet ws = ((WorksheetPart)extXls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                        var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                        if (sheetData != null)
                        {
                            var rows = sheetData.Elements<Row>();
                            foreach (var row in rows.Skip(1))
                            {
                                var cells = row.Elements<Cell>();
                                string aspectName = GetTextFromSharedTable(extXls, cells.First());

                                string performance = cells.ElementAtOrDefault(1)?.InnerText.Replace(".", ",") ?? "0,0";
                                string importance = cells.ElementAtOrDefault(2)?.InnerText.Replace(".", ",") ?? "0,0";

                                double performanceValue = 0.0;
                                double.TryParse(performance, out performanceValue);
                                double importanceValue = 0.0;
                                double.TryParse(importance, out importanceValue);

                                IPMA ipma = new IPMA();
                                ipma.Aspect = aspectName;
                                ipma.Performance = performanceValue;
                                ipma.Importance = importanceValue;

                                IPMAValues.Add(ipma);

                            }
                        }
                    }
                }
            }

            return IPMAValues;
        }

        internal static DataTable GetComplaintsData(string selectedPath)
        {
            //string dataFile = Directory.GetFiles(selectedPath + "\\Data\\").FirstOrDefault(x => x.EndsWith(".sav"));

            //RHandler rhandler = new RHandler(dataFile);
            //DataTable data = rhandler.GetComplaintsData();

            DataTable data = new DataTable();

            using (FileStream fs = new FileStream(selectedPath + "\\Output\\TabellerIBR.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument extXls = SpreadsheetDocument.Open(fs, false))
                {
                    var sheet = extXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>()
                        .FirstOrDefault(x => x.Name.ToString().ToLower() == "complaints");
                    if (sheet != null)
                    {
                        Worksheet ws = ((WorksheetPart)extXls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                        var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                        if (sheetData != null)
                        {
                            var rows = sheetData.Elements<Row>();
                            var headerRow = rows.ElementAt(1);
                            foreach (var cell in headerRow.Elements<Cell>())
                            {
                                data.Columns.Add(Functions.GetTextFromSharedTable(extXls, cell));
                            }

                            foreach (var row in rows.Skip(2))
                            {
                                DataRow dataRow = data.NewRow();
                                var cells = row.Elements<Cell>();

                                ArrayList alItems = new ArrayList();

                                alItems.Add(GetTextFromSharedTable(extXls, cells.ElementAt(0))); //aktörsnamn

                                alItems.Add(cells.ElementAt(1).CellValue.InnerText);

                                for (int i = 2; i < cells.Count(); i++)
                                {
                                    alItems.Add(double.Parse(cells.ElementAt(i).CellValue.InnerText, CultureInfo.InvariantCulture).ToString("p0"));
                                }

                                dataRow.ItemArray = alItems.ToArray();

                                data.Rows.Add(dataRow);
                            }
                        }
                    }

                }
            }

            return data;
        }        
    }
}
