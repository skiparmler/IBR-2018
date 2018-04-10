using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using System.Globalization;

namespace SKIIBR
{
    public partial class Form1 : Form
    {
        public string SelectedPath { get; set; }
        public string TemplatePath { get; set; }
        public Form1()
        {
            InitializeComponent();

#if DEBUG
            folderBrowser.SelectedPath = @"C:\Users\janste2\Documents\EPSI IBR\Bemanning";
            SelectedPath = folderBrowser.SelectedPath;
#endif
            dgvResultatVariabler.Rows.Insert(0, "epsi", "Kundnöjdhet");
            dgvResultatVariabler.Rows.Insert(1, "loyal", "Lojalitet");

            dgvKundDimensioner.Rows.Insert(0, "image", "Image");
            dgvKundDimensioner.Rows.Insert(1, "expect", "Förväntningar");
            dgvKundDimensioner.Rows.Insert(2, "prodq", "Produktkvalitet");
            dgvKundDimensioner.Rows.Insert(3, "servq", "Service");
            dgvKundDimensioner.Rows.Insert(4, "value", "Prisvärdhet");


            numYear.Value = DateTime.Now.Year;

            var segment = new List<string>()
            {
                "B2B",
                "B2C"
            };
            cbSegment.DataSource = segment;

            var branscher = new List<string>()
            {
                "SKI Bank",
                "SKI Mobiltelefoni",
                "SKI Digital-TV",
                "SKI Bredband",
                "SKI Sakförsäkring",
                "SKI Elhandel",
                "SKI Fjärrvärme",
                "SKI Elnät",
                "SKI Medarbetare",
                "SKI Tandvård",
                "SKI Samhälle",
                "SKI Fitness",
                "SKI Bolån",
                "SKI Privatlån",
                "SKI Livförsäkring",
                "SKI Sparande",
                "SKI Mäklare"
            };

            cbBranscher.DataSource = branscher;
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            var result = folderBrowser.ShowDialog();
            if (result == DialogResult.OK)
            {
                SelectedPath = folderBrowser.SelectedPath;
                lblPath.Text = SelectedPath;
            }
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {

            if (ValidateStructure())
            {
                btnExecute.Enabled = false;
                Go();
                btnExecute.Enabled = true;
            }
            else
            {
                MessageBox.Show("Felaktig mapp");
            }
        }

        private void Go()
        {
            Dictionary<string, string> aspectNames = new Dictionary<string, string>();
            foreach (DataGridViewRow dgvRow in dgvResultatVariabler.Rows) //varje rad är en aspekt
            {
                string dataName = dgvRow.Cells["dgvResultatVariablerNamn"].Value?.ToString().ToLower() ?? "epsi";

                string displayName = dgvRow.Cells["dgvResultatVariablerText"].Value?.ToString() ?? "Kundnöjdhet";

                if (!aspectNames.ContainsKey(dataName))
                {
                    aspectNames.Add(dataName, displayName);
                }
            }

            foreach (DataGridViewRow dgvRow in dgvKundDimensioner.Rows) //varje rad är en aspekt
            {
                string dataName = dgvRow.Cells["dataGridViewKundDimensionerNamn"].Value?.ToString().ToLower() ?? "image";

                string displayName = dgvRow.Cells["dataGridViewKundDimensionerText"].Value?.ToString() ?? "Image";

                if (!aspectNames.ContainsKey(dataName))
                {
                    aspectNames.Add(dataName, displayName);
                }
            }

            int maxPerSlide = 12;


            bool resultatAspekter = false;
            bool kunddimensioner = false;
            bool gapanalys = false;
            bool drivkraft = false;
            bool complaints = false;
            bool branschSpecifika = false;
            bool enableMoreBranschSpecifika = false;
            bool moreBranschSpecifika = false;
            bool segment = false;

            if (!Directory.Exists(SelectedPath + "\\Rapporter"))
            {
                Directory.CreateDirectory(SelectedPath + "\\Rapporter");
            }

            var outputFile = SelectedPath + "\\Rapporter\\" + cbBranscher.Text + ".pptx";
            File.Copy(TemplatePath, outputFile, true);

            using (PresentationDocument newDeck = PresentationDocument.Open(outputFile, true))
            {
                PresentationPart presPart = newDeck.PresentationPart;

                var items = presPart.Presentation.SlideIdList;
                //int loopIndex = 0;
                //foreach (SlideId item in items)
                for (int loopIndex = 0; true; loopIndex++)
                {
                    SlideId item = (SlideId)items.ElementAtOrDefault(loopIndex);
                    if (item == null)
                        break;
                    var part = presPart.GetPartById(item.RelationshipId);
                    var slide = (part as SlidePart).Slide;

                    var shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                    foreach (var shape in shapes)
                    {
                        SetText(shape, cbBranscher.Text, "{BRANSCH}");
                        SetText(shape, cbSegment.Text, "{Segment}");
                        SetText(shape, numYear.Value.ToString(), "{ÅR}");
                        SetText(shape, tbOrt.Text, "{Ort}");
                        SetText(shape, dateTimePicker.Text, "{Datum}");
                    }
                    #region Resultataspekter
                    if (resultatAspekter)
                    {
                        List<int> aspectSlidesIndices = new List<int>();

                        bool isFirst = true;

                        foreach (DataGridViewRow dgvRow in dgvResultatVariabler.Rows) //varje rad är en aspekt
                        {
                            string dataName = dgvRow.Cells["dgvResultatVariablerNamn"].Value?.ToString().ToLower() ?? "epsi";

                            string displayName = dgvRow.Cells["dgvResultatVariablerText"].Value?.ToString() ?? "Kundnöjdhet";

                            if (!isFirst)
                            {
                                if (dgvRow.Cells["dgvResultatVariablerText"].Value == null)
                                    break;

                                CreateNewSlides(aspectSlidesIndices, presPart, item, slide, items);

                                loopIndex++;
                                item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                if (item == null)
                                    break;
                                part = presPart.GetPartById(item.RelationshipId);
                                slide = (part as SlidePart).Slide;
                                shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                            }

                            foreach (var shape in shapes)
                            {
                                SetText(shape, displayName, "{Resultatvariabel}");
                                SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                SetText(shape, cbSegment.Text, "{Segment}");
                                SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                SetText(shape, tbOrt.Text, "{Ort}");
                                SetText(shape, dateTimePicker.Text, "{Datum}");
                            }

                            List<Tuple<string, double>> allValues = new List<Tuple<string, double>>();
                            List<Tuple<string, double>> values = new List<Tuple<string, double>>();

                            //hämta data från Latent_Summary

                            allValues = Functions.GetAspectForActors(dataName, SelectedPath);

                            bool firstLoop = true;
                            foreach (var allValue in allValues.OrderByDescending(x => x.Item2))
                            {
                                if (allValue.Item1 == "industry")
                                    continue;
                                values.Add(allValue);
                                if (values.Count == maxPerSlide)
                                {
                                    values.Add(allValues.FirstOrDefault(x => x.Item1 == "industry"));

                                    if (!firstLoop)
                                    {
                                        //Kopiera första sliden
                                        CreateNewSlides(new List<int>() { loopIndex }, presPart, item, slide, items);
                                        loopIndex++;
                                        item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                        if (item == null)
                                            break;
                                        part = presPart.GetPartById(item.RelationshipId);
                                        slide = (part as SlidePart).Slide;
                                        shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                                    }

                                    Functions.UpdateResultatAspektChart(slide, values, displayName, cbBranscher.Text, cbSegment.Text, numYear.Value.ToString());

                                    Functions.UpdateManifestTable(slide, dataName, SelectedPath, 0, "Tabell 9", values.Select(x => x.Item1).ToList(), 4);
                                    values.Clear();
                                    if (firstLoop)
                                    {
                                        aspectSlidesIndices.Add(loopIndex);
                                    }
                                    firstLoop = false;
                                }
                            }

                            if (values.Count > 0)
                            {
                                values.Add(allValues.FirstOrDefault(x => x.Item1 == "industry"));

                                if (!firstLoop)
                                {
                                    //Kopiera första sliden
                                    CreateNewSlides(new List<int>() { loopIndex }, presPart, item, slide, items);
                                    loopIndex++;
                                    item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                    if (item == null)
                                        break;
                                    part = presPart.GetPartById(item.RelationshipId);
                                    slide = (part as SlidePart).Slide;
                                    shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                                }
                                if (firstLoop)
                                {
                                    aspectSlidesIndices.Add(loopIndex);
                                }
                                Functions.UpdateResultatAspektChart(slide, values, displayName, cbBranscher.Text, cbSegment.Text, numYear.Value.ToString());

                                //tabellen
                                Functions.UpdateManifestTable(slide, dataName, SelectedPath, 0, "Tabell 9", values.Select(x => x.Item1).ToList(), 4);
                            }

                            //hoppa fram en slide - historiktabellen
                            loopIndex++;
                            item = (SlideId)items.ElementAtOrDefault(loopIndex);
                            if (item == null)
                                break;
                            part = presPart.GetPartById(item.RelationshipId);
                            slide = (part as SlidePart).Slide;

                            shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                            foreach (var shape in shapes)
                            {
                                SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                SetText(shape, cbSegment.Text, "{Segment}");
                                SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                SetText(shape, tbOrt.Text, "{Ort}");
                                SetText(shape, dateTimePicker.Text, "{Datum}");
                                SetText(shape, displayName, "{Resultatvariabel}");
                            }
                            List<int> years = new List<int>();
                            var historicValues = Functions.GetHistoricValues(dataName, SelectedPath, (int)numYear.Value, years);
                            firstLoop = true;
                            List<Tuple<string, List<double>>> valuesToUse = new List<Tuple<string, List<double>>>();
                            foreach (var actor in Functions.GetActors(SelectedPath))
                            {
                                var historicValue = historicValues.FirstOrDefault(x => x.Item1 == actor);
                                if (historicValue == null)
                                {
                                    historicValue = new Tuple<string, List<double>>(actor, new List<double>());
                                    foreach (var year in years.Skip(1))
                                    {
                                        historicValue.Item2.Add(0.0);
                                    }
                                    historicValue.Item2.Add(Functions.GetValueFromLatentSummary(actor, dataName, SelectedPath));
                                }
                                if (historicValue.Item1 == "Branschen")
                                    continue;

                                valuesToUse.Add(historicValue);
                                if (valuesToUse.Count == maxPerSlide)
                                {
                                    valuesToUse.Add(historicValues.FirstOrDefault(x => x.Item1 == "Branschen"));
                                    if (!firstLoop)
                                    {
                                        //Kopiera första sliden
                                        CreateNewSlides(new List<int>() { loopIndex }, presPart, item, slide, items);
                                        loopIndex++;
                                        item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                        if (item == null)
                                            break;
                                        part = presPart.GetPartById(item.RelationshipId);
                                        slide = (part as SlidePart).Slide;
                                        shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                                    }

                                    Functions.UpdateHistoryTable(slide, dataName, valuesToUse, SelectedPath, true, true, 0, "Tabell 9", years);
                                    valuesToUse.Clear();
                                    if (firstLoop)
                                    {
                                        aspectSlidesIndices.Add(loopIndex);
                                    }
                                    firstLoop = false;

                                }
                            }

                            if (valuesToUse.Count > 0)
                            {
                                valuesToUse.Add(historicValues.FirstOrDefault(x => x.Item1 == "Branschen"));

                                if (!firstLoop)
                                {
                                    //Kopiera första sliden
                                    CreateNewSlides(new List<int>() { loopIndex }, presPart, item, slide, items);
                                    loopIndex++;
                                    item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                    if (item == null)
                                        break;
                                    part = presPart.GetPartById(item.RelationshipId);
                                    slide = (part as SlidePart).Slide;
                                    shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                                }
                                if (firstLoop)
                                {
                                    aspectSlidesIndices.Add(loopIndex);
                                }
                                Functions.UpdateHistoryTable(slide, dataName, valuesToUse, SelectedPath, true, true, 0, "Tabell 9", years);
                            }

                            #region Spreads
                            //hoppa fram en slide - spridningsgrafen
                            loopIndex++;
                            item = (SlideId)items.ElementAtOrDefault(loopIndex);
                            if (item == null)
                                break;
                            part = presPart.GetPartById(item.RelationshipId);
                            slide = (part as SlidePart).Slide;

                            shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                            foreach (var shape in shapes)
                            {
                                SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                SetText(shape, cbSegment.Text, "{Segment}");
                                SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                SetText(shape, tbOrt.Text, "{Ort}");
                                SetText(shape, dateTimePicker.Text, "{Datum}");
                                SetText(shape, displayName, "{Resultatvariabel}");
                            }

                            //öppna spreadsdata
                            List<KeyValuePair<string, List<double>>> actorSpread = new List<KeyValuePair<string, List<double>>>();
                            using (FileStream fs = new FileStream(SelectedPath + "\\Output\\Spreadsdatan.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            {
                                using (SpreadsheetDocument extXls = SpreadsheetDocument.Open(fs, false))
                                {
                                    var sheet = extXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(x => x.Name.ToString().ToLower() == "sheet1");
                                    if (sheet != null)
                                    {
                                        Worksheet ws = ((WorksheetPart)extXls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                                        var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                                        if (sheetData != null)
                                        {
                                            var spreadRows = sheetData.Elements<Row>();
                                            //företagsnamnen finns på rad 9, kolumn 3 och framåt

                                            string firstCompany = string.Empty;
                                            var companyRow = spreadRows.FirstOrDefault(x => x.RowIndex == 9);
                                            var companyCells = companyRow.Elements<Cell>().Where(x => x.InnerText != string.Empty);
                                            foreach (var companyCell in companyCells)
                                            {
                                                string companyName = Functions.GetTextFromSharedTable(extXls, companyCell);
                                                if (firstCompany == string.Empty)
                                                {
                                                    firstCompany = companyName;
                                                }
                                                else if (companyName == firstCompany)
                                                {
                                                    break;
                                                }
                                                KeyValuePair<string, List<double>> actor = new KeyValuePair<string, List<double>>(companyName, new List<double>());
                                                actorSpread.Add(actor);
                                            }
                                            foreach (var row in spreadRows)
                                            {
                                                if (row.GetFirstChild<Cell>() == null)
                                                {
                                                    continue;
                                                }
                                                var firstCell = row.GetFirstChild<Cell>();
                                                if (!string.IsNullOrEmpty(firstCell.InnerText) && Functions.GetTextFromSharedTable(extXls, firstCell).ToLower().Contains(dataName))
                                                {
                                                    var badRow = row;
                                                    var neutralRow = spreadRows.FirstOrDefault(x => x.RowIndex == badRow.RowIndex + 1);
                                                    var goodRow = spreadRows.FirstOrDefault(x => x.RowIndex == badRow.RowIndex + 2);
                                                    //första datat finns i kolumn 3
                                                    for (int i = 0; i < actorSpread.Count; i++)
                                                    {
                                                        var bad = double.Parse(badRow.Elements<Cell>().ElementAt(2 + i).InnerText, CultureInfo.InvariantCulture);
                                                        var neutral = double.Parse(neutralRow.Elements<Cell>().ElementAt(2 + i).InnerText, CultureInfo.InvariantCulture);
                                                        var good = double.Parse(goodRow.Elements<Cell>().ElementAt(2 + i).InnerText, CultureInfo.InvariantCulture);

                                                        actorSpread.ElementAt(i).Value.AddRange(new List<double>() { bad, neutral, good });
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            //använd actorspread-resultatet
                            firstLoop = true;
                            List<KeyValuePair<string, List<double>>> spreadToUse = new List<KeyValuePair<string, List<double>>>();
                            foreach (var spreadData in actorSpread)
                            {
                                spreadToUse.Add(spreadData);
                                if (spreadToUse.Count == maxPerSlide)
                                {
                                    if (!firstLoop)
                                    {
                                        //Kopiera första sliden
                                        CreateNewSlides(new List<int>() { loopIndex }, presPart, item, slide, items);
                                        loopIndex++;
                                        item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                        if (item == null)
                                            break;
                                        part = presPart.GetPartById(item.RelationshipId);
                                        slide = (part as SlidePart).Slide;
                                        shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                                    }

                                    Functions.UpdateSpreadsTable(slide, spreadToUse, displayName, cbBranscher.Text, cbSegment.Text, numYear.Value.ToString());
                                    spreadToUse.Clear();
                                    if (firstLoop)
                                    {
                                        aspectSlidesIndices.Add(loopIndex);
                                    }
                                    firstLoop = false;
                                }
                            }

                            if (spreadToUse.Count > 0)
                            {
                                if (!firstLoop)
                                {
                                    //Kopiera första sliden
                                    CreateNewSlides(new List<int>() { loopIndex }, presPart, item, slide, items);
                                    loopIndex++;
                                    item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                    if (item == null)
                                        break;
                                    part = presPart.GetPartById(item.RelationshipId);
                                    slide = (part as SlidePart).Slide;
                                    shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                                }
                                if (firstLoop)
                                {
                                    aspectSlidesIndices.Add(loopIndex);
                                }
                                Functions.UpdateSpreadsTable(slide, spreadToUse, displayName, cbBranscher.Text, cbSegment.Text, numYear.Value.ToString());
                            }

                            //end resultaspekt
                            #endregion
                            isFirst = false;
                        }
                        resultatAspekter = false;
                    }
                    #endregion

                    #region Kunddimensioner a.k.a. drivande aspekter
                    if (kunddimensioner)
                    {
                        List<int> dimensionSlidesIndices = new List<int>();
                        dimensionSlidesIndices.Add(loopIndex);

                        bool isFirst = true;

                        foreach (DataGridViewRow dgvRow in dgvKundDimensioner.Rows) //varje rad är en dimension
                        {
                            string dataName = dgvRow.Cells["dataGridViewKundDimensionerNamn"].Value?.ToString().ToLower() ?? "image";

                            string displayName = dgvRow.Cells["dataGridViewKundDimensionerText"].Value?.ToString() ?? "Image";

                            if (!isFirst)
                            {
                                if (dgvRow.Cells["dataGridViewKundDimensionerText"].Value == null)
                                    break;

                                CreateNewSlides(dimensionSlidesIndices, presPart, item, slide, items);

                                loopIndex++;
                                item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                if (item == null)
                                    break;
                                part = presPart.GetPartById(item.RelationshipId);
                                slide = (part as SlidePart).Slide;
                                shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                            }

                            foreach (var shape in shapes)
                            {
                                SetText(shape, displayName, "{Kunddimension}");
                                SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                SetText(shape, cbSegment.Text, "{Segment}");
                                SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                SetText(shape, tbOrt.Text, "{Ort}");
                                SetText(shape, dateTimePicker.Text, "{Datum}");
                            }

                            List<int> years = new List<int>();
                            var historicValues = Functions.GetHistoricValues(dataName, SelectedPath, (int)numYear.Value, years);

                            List<Tuple<string, List<double>>> valuesToUse = new List<Tuple<string, List<double>>>();
                            bool firstLoop = true;
                            foreach (var actor in Functions.GetActors(SelectedPath))
                            {
                                var historicValue = historicValues.FirstOrDefault(x => x.Item1 == actor);
                                if (historicValue == null)
                                {
                                    historicValue = new Tuple<string, List<double>>(actor, new List<double>());
                                    foreach (var year in years.Skip(1))
                                    {
                                        historicValue.Item2.Add(0.0);
                                    }
                                    historicValue.Item2.Add(Functions.GetValueFromLatentSummary(actor, dataName, SelectedPath));
                                }

                                if (historicValue.Item1 == "Branschen")
                                    continue;
                                valuesToUse.Add(historicValue);
                                if (valuesToUse.Count == maxPerSlide)
                                {
                                    valuesToUse.Add(historicValues.FirstOrDefault(x => x.Item1 == "Branschen"));
                                    if (!firstLoop)
                                    {
                                        //Kopiera första sliden
                                        CreateNewSlides(new List<int>() { loopIndex }, presPart, item, slide, items);
                                        loopIndex++;
                                        item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                        if (item == null)
                                            break;
                                        part = presPart.GetPartById(item.RelationshipId);
                                        slide = (part as SlidePart).Slide;
                                        shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                                    }
                                    Functions.UpdateHistoryTable(slide, dataName, valuesToUse, SelectedPath, true, false, 1, "Tabell 9", years);

                                    Functions.UpdateTableText(slide, "{Kunddimension}", displayName);
                                    Functions.UpdateTableText(slide, "{ÅR}", ((int)numYear.Value).ToString());

                                    Functions.UpdateManifestTable(slide, dataName, SelectedPath, 1, "Tabell 6", valuesToUse.Select(x => x.Item1).ToList(), 6);

                                    Functions.OrderTables(slide, new List<string>() { "Tabell 9", "Tabell 6" });

                                    valuesToUse.Clear();
                                    firstLoop = false;
                                }
                            }

                            if (valuesToUse.Count > 0)
                            {
                                valuesToUse.Add(historicValues.FirstOrDefault(x => x.Item1 == "Branschen"));
                                if (!firstLoop)
                                {
                                    //Kopiera första sliden
                                    CreateNewSlides(new List<int>() { loopIndex }, presPart, item, slide, items);
                                    loopIndex++;
                                    item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                    if (item == null)
                                        break;
                                    part = presPart.GetPartById(item.RelationshipId);
                                    slide = (part as SlidePart).Slide;
                                    shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                                    foreach (var shape in shapes)
                                    {
                                        SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                        SetText(shape, cbSegment.Text, "{Segment}");
                                        SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                        SetText(shape, tbOrt.Text, "{Ort}");
                                        SetText(shape, dateTimePicker.Text, "{Datum}");
                                    }
                                }
                                Functions.UpdateHistoryTable(slide, dataName, valuesToUse, SelectedPath, true, false, 1, "Tabell 9", years);

                                Functions.UpdateTableText(slide, "{Kunddimension}", displayName);
                                Functions.UpdateTableText(slide, "{ÅR}", ((int)numYear.Value).ToString());

                                Functions.UpdateManifestTable(slide, dataName, SelectedPath, 1, "Tabell 6", valuesToUse.Select(x => x.Item1).ToList(), 6);

                                Functions.OrderTables(slide, new List<string>() { "Tabell 9", "Tabell 6" });
                            }


                            isFirst = false;
                        }
                        kunddimensioner = false;
                    }


                    #endregion

                    #region Gapanalys
                    if (gapanalys)
                    {
                        List<string> dataNames = new List<string>();
                        Dictionary<string, string> aspects = new Dictionary<string, string>();

                        foreach (DataGridViewRow dgvRow in dgvKundDimensioner.Rows)
                        {
                            if (dgvKundDimensioner.Rows.Count == 1 && dgvRow.Cells["dataGridViewKundDimensionerNamn"].Value == null)
                            {
                                dataNames.Add("image");
                                aspects.Add("image", "Image");
                            }
                            else
                            {
                                string dataName = dgvRow.Cells["dataGridViewKundDimensionerNamn"].Value?.ToString().ToLower();
                                if (dataName != null)
                                {
                                    dataNames.Add(dataName);

                                    string displayName = dgvRow.Cells["dataGridViewKundDimensionerText"].Value?.ToString();
                                    if (displayName != null)
                                    {
                                        if (!aspects.ContainsKey(dataName))
                                            aspects.Add(dataName, displayName);
                                    }
                                }
                            }
                        }

                        foreach (DataGridViewRow dgvRow in dgvResultatVariabler.Rows) //varje rad är en aspekt
                        {
                            if (dgvResultatVariabler.Rows.Count == 1 && dgvRow.Cells["dgvResultatVariablerNamn"] == null)
                            {
                                dataNames.Add("epsi");
                                aspects.Add("epsi", "Kundnöjdhet");
                            }
                            else
                            {
                                string dataName = dgvRow.Cells["dgvResultatVariablerNamn"].Value?.ToString().ToLower();
                                if (dataName != null)
                                {
                                    dataNames.Add(dataName);

                                    string displayName = dgvRow.Cells["dgvResultatVariablerText"].Value?.ToString();
                                    if (displayName != null)
                                    {
                                        if (!aspects.ContainsKey(dataName))
                                            aspects.Add(dataName, displayName);
                                    }
                                }
                            }
                        }

                        List<int> chartOrder = new List<int>() { 5, 3, 2, 1, 0, 6, 4 }; //TODO: rensa bort grafer som inte används
                        Dictionary<string, List<int>> aspectYears = new Dictionary<string, List<int>>();
                        Dictionary<string, Dictionary<string, List<double>>> actorAspects = new Dictionary<string, Dictionary<string, List<double>>>();

                        foreach (string dataName in dataNames)
                        {
                            List<int> years = new List<int>();
                            var historicData = Functions.GetHistoricValues(dataName, SelectedPath, (int)numYear.Value, years);
                            aspectYears.Add(dataName, years);
                            var branschData = historicData.FirstOrDefault(x => x.Item1 == "Branschen");

                            //foreach (var historicValue in historicData.Where(x => x.Item1 != "Branschen"))
                            //{
                            foreach (var actor in Functions.GetActors(SelectedPath))
                            {
                                var historicValue = historicData.FirstOrDefault(x => x.Item1 == actor);
                                if (historicValue == null)
                                {
                                    historicValue = new Tuple<string, List<double>>(actor, new List<double>());
                                    foreach (var year in years.Skip(1))
                                    {
                                        historicValue.Item2.Add(0.0);
                                    }
                                    historicValue.Item2.Add(Functions.GetValueFromLatentSummary(actor, dataName, SelectedPath));

                                }

                                if (!actorAspects.ContainsKey(historicValue.Item1))
                                {
                                    actorAspects.Add(historicValue.Item1, new Dictionary<string, List<double>>());
                                }
                                if (!actorAspects[historicValue.Item1].ContainsKey(dataName))
                                {
                                    actorAspects[historicValue.Item1].Add(dataName, new List<double>());
                                }
                                int index = 0;
                                foreach (var x in historicValue.Item2)
                                {
                                    double newValue = x;
                                    newValue -= branschData.Item2.ElementAt(index++);
                                    actorAspects[historicValue.Item1][dataName].Add(newValue);
                                }
                            }
                        }

                        var aspectQuestions = Functions.GetAspectQuestions(SelectedPath);

                        bool firstActor = true;

                        foreach (var actor in actorAspects.Keys)
                        {
                            if (!firstActor)
                            {
                                CreateNewSlides(new List<int>() { loopIndex - 2, loopIndex - 1, loopIndex }, presPart, item, slide, items);
                                loopIndex++;
                                item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                if (item == null)
                                    break;
                                part = presPart.GetPartById(item.RelationshipId);
                                slide = (part as SlidePart).Slide;
                            }

                            shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                            foreach (var shape in shapes)
                            {
                                SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                SetText(shape, cbSegment.Text, "{Segment}");
                                SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                SetText(shape, tbOrt.Text, "{Ort}");
                                SetText(shape, dateTimePicker.Text, "{Datum}");
                                SetText(shape, actor, "{Aktör}");
                            }


                            var chartParts = slide.SlidePart.ChartParts;
                            int chartIndex = 0;
                            foreach (var aspect in actorAspects[actor].Keys)
                            {
                                //gå igenom alla grafer

                                var chartPart = chartParts.ElementAtOrDefault(chartOrder.ElementAt(chartIndex));
                                D.Charts.Chart chart = chartPart.ChartSpace.GetFirstChild<D.Charts.Chart>();

                                chart.Title.ChartText.RichText.Elements<D.Paragraph>().First().Elements<D.Run>().First().Text = new D.Text(aspects[aspect]);
                                chart.Title.ChartText.RichText.Elements<D.Paragraph>().ElementAt(1).Elements<D.Run>().First().Text = new D.Text(actor);

                                var plotArea = chart.GetFirstChild<D.Charts.PlotArea>();
                                List<int> indicesToRemove = new List<int>();
                                for(int i = 0; i < actorAspects[actor][aspect].Count;i++)
                                {
                                    if (actorAspects[actor][aspect].ElementAt(i) < -20 || actorAspects[actor][aspect].ElementAt(i) > 20)
                                    {
                                        indicesToRemove.Add(i);
                                    }
                                }

                                if(indicesToRemove.Count > 0)
                                {
                                    indicesToRemove.Add(indicesToRemove.Last() + 1);
                                }
                                
                                var maxValue = Math.Ceiling(actorAspects[actor][aspect].Where(x => x < 20).Max());
                                var minValue = Math.Floor(actorAspects[actor][aspect].Where(x => x > -20).Min());

                                var valueToUse = Math.Max(maxValue, Math.Abs(minValue));

                                var maxAxis = (valueToUse + 2) < 5 ? 5 : valueToUse + 2;
                                var minAxis = -maxAxis;

                                var scaling = plotArea.Elements<D.Charts.ValueAxis>().FirstOrDefault()?.Elements<D.Charts.Scaling>().FirstOrDefault();
                                if (scaling != null)
                                {                                    
                                    D.Charts.MaxAxisValue maxAxisValue = new D.Charts.MaxAxisValue() { Val = maxAxis };
                                    D.Charts.MinAxisValue minAxisValue = new D.Charts.MinAxisValue() { Val = minAxis };

                                    scaling.AppendChild(maxAxisValue);
                                    scaling.AppendChild(minAxisValue);
                                }
                                var lineChart = plotArea.GetFirstChild<D.Charts.LineChart>();

                                foreach (var series in lineChart.Elements<D.Charts.LineChartSeries>())
                                {
                                    foreach (var indexToRemove in indicesToRemove)
                                    {
                                        D.Charts.DataPoint dpt = new D.Charts.DataPoint();
                                        dpt.Index = new D.Charts.Index();
                                        dpt.Index.Val = (uint)indexToRemove;
                                        D.Charts.ChartShapeProperties csPr = new D.Charts.ChartShapeProperties();
                                        D.Outline outline = new D.Outline();
                                        D.NoFill nofill = new D.NoFill();
                                        outline.AppendChild(nofill);
                                        csPr.AppendChild(outline);
                                        dpt.AppendChild(csPr);
                                        series.AppendChild(dpt);
                                    }

                                    var numberReference = series.GetFirstChild<D.Charts.Values>()?.GetFirstChild<D.Charts.NumberReference>();
                                    var numberRange = numberReference?.GetFirstChild<D.Charts.Formula>();
                                    var numberCache = numberReference?.GetFirstChild<D.Charts.NumberingCache>();
                                    int valueIndex = 0;

                                    numberCache.RemoveAllChildren();

                                    var labelReference = series.GetFirstChild<D.Charts.CategoryAxisData>()?.GetFirstChild<D.Charts.NumberReference>();
                                    var labelCache = labelReference?.GetFirstChild<D.Charts.NumberingCache>();

                                    labelCache.RemoveAllChildren();

                                    int ncIndex = 0;
                                    var years = aspectYears[aspect];
                                    foreach (var value in actorAspects[actor][aspect])
                                    {
                                        var currentNP = new D.Charts.NumericPoint();
                                        currentNP.NumericValue = new D.Charts.NumericValue(value.ToString("n1").Replace(",", "."));
                                        currentNP.Index = (uint)ncIndex;
                                        numberCache.AppendChild(currentNP);

                                        var currentSP = new D.Charts.NumericPoint();
                                        currentSP.NumericValue = new D.Charts.NumericValue(years.ElementAt(ncIndex).ToString());
                                        currentSP.Index = (uint)ncIndex;
                                        labelCache.AppendChild(currentSP);

                                        ncIndex++;
                                    }

                                    EmbeddedPackagePart embeddedExcel = chartPart.EmbeddedPackagePart;
                                    if (embeddedExcel != null)
                                    {
                                        using (Stream str = embeddedExcel.GetStream())
                                        {
                                            using (SpreadsheetDocument xls = SpreadsheetDocument.Open(str, true))
                                            {
                                                var sheet = xls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(x => x.Name == "Blad1");
                                                if (sheet != null)
                                                {
                                                    WorksheetPart wsPart = (WorksheetPart)xls.WorkbookPart.GetPartById(sheet.Id);
                                                    Worksheet ws = wsPart.Worksheet;
                                                    var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                                                    if (sheetData != null)
                                                    {
                                                        var reference = numberReference.Formula.Text.Split(new char[] { '!' }).Last().Replace("$", string.Empty).Split(new char[] { ':' });
                                                        string start = reference.First();
                                                        string stop = reference.Last();

                                                        int rowIndex = int.Parse(start.Substring(1, 1));
                                                        int lastRowIndex = rowIndex + (ncIndex - 1);//int.Parse(stop.Substring(1, 1));
                                                        string newFormula = numberReference.Formula.Text.Substring(0, numberReference.Formula.Text.LastIndexOf("$"));
                                                        numberReference.Formula.Text = newFormula + "$" + lastRowIndex.ToString();
                                                        valueIndex = 0;
                                                        string colRef = start.Substring(0, 1);

                                                        var rows = sheetData.Elements<Row>();
                                                        while (rowIndex <= lastRowIndex)
                                                        {
                                                            var currentRow = rows.FirstOrDefault(x => x.RowIndex == rowIndex);
                                                            if (currentRow == null || (actorAspects[actor][aspect].Count - 1) < valueIndex)
                                                            {
                                                                break;
                                                            }
                                                            var cells = currentRow.Elements<Cell>();
                                                            var currentCell = cells.FirstOrDefault(x => x.CellReference.Value == colRef + rowIndex.ToString());
                                                            currentCell.CellValue = new CellValue(actorAspects[actor][aspect].ElementAt(valueIndex).ToString("n1").Replace(",", "."));

                                                            string yearColRef = ((char)(colRef[0] - 1)).ToString();

                                                            var yearCell = cells.FirstOrDefault(x => x.CellReference.Value == yearColRef + rowIndex.ToString());
                                                            yearCell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);
                                                            yearCell.CellValue = new CellValue(years.ElementAt(valueIndex).ToString());

                                                            valueIndex++;
                                                            rowIndex++;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                chartIndex++;
                            }
                            firstActor = false;


                            //hoppa fram till frågespecifikationen
                            loopIndex++;
                            item = (SlideId)items.ElementAtOrDefault(loopIndex);
                            if (item == null)
                                break;
                            part = presPart.GetPartById(item.RelationshipId);
                            slide = (part as SlidePart).Slide;

                            shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                            foreach (var shape in shapes)
                            {
                                SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                SetText(shape, cbSegment.Text, "{Segment}");
                                SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                SetText(shape, tbOrt.Text, "{Ort}");
                                SetText(shape, dateTimePicker.Text, "{Datum}");
                                SetText(shape, actor, "{Aktör}");
                            }

                            List<Tuple<string, double>> diffPerQ = new List<Tuple<string, double>>();

                            var gapTableGF = slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>();
                            GraphicFrame gapContainer = gapTableGF.FirstOrDefault();// x => x.InnerXml.Contains("Gapanalys per modellfråga"));
                            var gapTable = gapContainer.GetFirstChild<D.Graphic>()?.GraphicData.GetFirstChild<D.Table>();
                            int startIndex = 3;
                            //foreach(var actor in GetActors())
                            //{

                            if (gapTable != null)
                            {

                                D.TableRow rowToUse = (D.TableRow)gapTable.Elements<D.TableRow>().ElementAt(startIndex).CloneNode(true);

                                Functions.SetTableCellText(gapTable.Elements<D.TableRow>().ElementAt(1).Elements<D.TableCell>().ElementAt(2), actor);

                                while (gapTable.Elements<D.TableRow>().Count() > (startIndex - 1))
                                {
                                    gapTable.RemoveChild(gapTable.Elements<D.TableRow>().Last());
                                }

                                using (FileStream fs = new FileStream(SelectedPath + "\\Output\\manifests.xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                {
                                    using (SpreadsheetDocument extXls = SpreadsheetDocument.Open(fs, false))
                                    {
                                        var sheets = extXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>();
                                        foreach (var sheet in sheets)
                                        {
                                            if (sheet.Name.ToString().Contains(" "))
                                            {
                                                continue;
                                            }
                                            Worksheet ws = ((WorksheetPart)extXls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                                            var sheetData = ws.Elements<SheetData>().FirstOrDefault();

                                            var subAspects = sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex == 1).Elements<Cell>();
                                            var dataRow = sheetData.Elements<Row>().FirstOrDefault(x => Functions.GetTextFromSharedTable(extXls, x.Elements<Cell>().ElementAtOrDefault(1)) == actor);
                                            if (dataRow == null)
                                            {
                                                continue;
                                            }
                                            var branschRow = sheetData.Elements<Row>().FirstOrDefault(x => Functions.GetTextFromSharedTable(extXls, x.Elements<Cell>().ElementAt(1)) == "industry");
                                            bool isData = false;
                                            int dataIndex = 2;
                                            foreach (var subAspect in subAspects)
                                            {
                                                if (!isData)
                                                {
                                                    if (subAspect.CellReference.Value == "B1")
                                                    {
                                                        isData = true;
                                                    }
                                                }
                                                else
                                                {
                                                    var newRow = rowToUse.CloneNode(true);
                                                    var aspectNameCell = newRow.Elements<D.TableCell>().ElementAt(0);
                                                    Functions.SetTableCellText(aspectNameCell, aspects.ContainsKey(sheet.Name) ? aspects[sheet.Name] : sheet.Name.ToString());
                                                    var qCell = newRow.Elements<D.TableCell>().ElementAt(1);
                                                    var aspectName = Functions.GetTextFromSharedTable(extXls, subAspect);
                                                    string aspectText = aspectQuestions.ContainsKey(aspectName) ? aspectQuestions[aspectName] : aspectName;
                                                    Functions.SetTableCellText(qCell, aspectText);

                                                    var actorValue = double.Parse(dataRow.Elements<Cell>().FirstOrDefault(x => x.CellReference == subAspect.CellReference.Value.Substring(0, 1) + dataRow.RowIndex.ToString()).InnerText, CultureInfo.InvariantCulture);
                                                    var branschValue = double.Parse(branschRow.Elements<Cell>().FirstOrDefault(x => x.CellReference == subAspect.CellReference.Value.Substring(0, 1) + branschRow.RowIndex.ToString()).InnerText, CultureInfo.InvariantCulture);
                                                    var difference = actorValue - branschValue;

                                                    var actorValueCell = newRow.Elements<D.TableCell>().ElementAt(2);
                                                    Functions.SetTableCellText(actorValueCell, actorValue.ToString("n1"));
                                                    var branschValueCell = newRow.Elements<D.TableCell>().ElementAt(3);
                                                    Functions.SetTableCellText(branschValueCell, branschValue.ToString("n1"));
                                                    var differenceValueCell = newRow.Elements<D.TableCell>().ElementAt(4);
                                                    Functions.SetTableCellText(differenceValueCell, difference.ToString("n1"));

                                                    diffPerQ.Add(new Tuple<string, double>(aspectName, difference));

                                                    gapTable.AppendChild(newRow);
                                                }
                                            }

                                        }
                                    }
                                }
                            }

                            //    break;
                            //}

                            loopIndex++;
                            item = (SlideId)items.ElementAtOrDefault(loopIndex);
                            if (item == null)
                                break;
                            part = presPart.GetPartById(item.RelationshipId);
                            slide = (part as SlidePart).Slide;

                            shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                            foreach (var shape in shapes)
                            {
                                SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                SetText(shape, cbSegment.Text, "{Segment}");
                                SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                SetText(shape, tbOrt.Text, "{Ort}");
                                SetText(shape, dateTimePicker.Text, "{Datum}");
                                SetText(shape, actor, "{Aktör}");
                            }

                            var diffQchartPart = slide.SlidePart.ChartParts.FirstOrDefault();
                            if (diffQchartPart != null)
                            {
                                D.Charts.Chart chart = diffQchartPart.ChartSpace.GetFirstChild<D.Charts.Chart>();
                                var paragraph = chart.Title.ChartText.RichText.Elements<D.Paragraph>().First();
                                var runs = paragraph.Elements<D.Run>();
                                runs.First().Text = new D.Text(actor);
                                while (runs.Count() > 1)
                                {
                                    paragraph.RemoveChild(runs.Last());
                                }

                                var plotArea = chart.GetFirstChild<D.Charts.PlotArea>();

                                var barChart = plotArea.GetFirstChild<D.Charts.BarChart>();

                                foreach (var series in barChart.Elements<D.Charts.BarChartSeries>())
                                {
                                    var numberReference = series.GetFirstChild<D.Charts.Values>()?.GetFirstChild<D.Charts.NumberReference>();
                                    var numberRange = numberReference?.GetFirstChild<D.Charts.Formula>();
                                    var numberCache = numberReference?.GetFirstChild<D.Charts.NumberingCache>();

                                    var stringReference = barChart.GetFirstChild<D.Charts.BarChartSeries>()?.GetFirstChild<D.Charts.CategoryAxisData>()?.GetFirstChild<D.Charts.StringReference>();
                                    var textRange = stringReference?.GetFirstChild<D.Charts.Formula>();
                                    var textCache = stringReference?.GetFirstChild<D.Charts.StringCache>();

                                    numberCache.RemoveAllChildren();
                                    textCache.RemoveAllChildren();

                                    int ncIndex = 0;
                                    diffPerQ.Reverse();
                                    foreach (var valuePair in diffPerQ)
                                    {
                                        var currentNP = new D.Charts.NumericPoint();
                                        currentNP.NumericValue = new D.Charts.NumericValue(valuePair.Item2.ToString("n1").Replace(",", "."));
                                        currentNP.Index = (uint)ncIndex;
                                        numberCache.AppendChild(currentNP);

                                        string aspectText = aspectQuestions.ContainsKey(valuePair.Item1) ? $"{valuePair.Item1}: {aspectQuestions[valuePair.Item1]}" : valuePair.Item1;
                                        var currentSP = new D.Charts.StringPoint();
                                        currentSP.InnerXml = $"<c:v xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">{System.Security.SecurityElement.Escape(aspectText)}</c:v>";
                                        currentSP.Index = (uint)ncIndex;
                                        textCache.AppendChild(currentSP);

                                        ncIndex++;
                                    }
                                }

                                EmbeddedPackagePart embeddedExcel = diffQchartPart.EmbeddedPackagePart;
                                if (embeddedExcel != null)
                                {
                                    using (Stream str = embeddedExcel.GetStream())
                                    {
                                        using (SpreadsheetDocument xls = SpreadsheetDocument.Open(str, true))
                                        {
                                            string sheetName = "Gapgraf";
                                            var sheet = xls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name.HasValue && s.Name.Value == sheetName);
                                            if (sheet != null)
                                            {
                                                Worksheet ws = ((WorksheetPart)xls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                                                var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                                                if (sheetData != null)
                                                {
                                                    int rowIndex = 0;
                                                    var rows = sheetData.Elements<Row>();

                                                    foreach (var valuePair in diffPerQ)
                                                    {
                                                        var currentRow = rows.ElementAtOrDefault(rowIndex);
                                                        if (currentRow == null)
                                                        {
                                                            currentRow = sheetData.InsertAt(new Row() { RowIndex = (uint)rowIndex }, sheetData.Elements<Row>().Count());
                                                            var labelCell = currentRow.InsertAt(new Cell() { CellReference = "F" + rowIndex }, 0);
                                                            labelCell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);
                                                            var valueCell = currentRow.InsertAt(new Cell() { CellReference = "G" + rowIndex }, 1);
                                                            valueCell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                                                        }
                                                        var cells = currentRow.Elements<Cell>();
                                                        if (rowIndex == 0)
                                                        {
                                                            rowIndex++;
                                                        }
                                                        string aspectText = aspectQuestions.ContainsKey(valuePair.Item1) ? aspectQuestions[valuePair.Item1] : valuePair.Item1;
                                                        cells.ElementAt(0).CellValue = new CellValue(aspectText);
                                                        cells.ElementAt(0).DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);

                                                        cells.ElementAt(1).CellValue = new CellValue(valuePair.Item2.ToString("n1").Replace(",", "."));

                                                        rowIndex++;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        gapanalys = false;
                    }
                    #endregion

                    #region Drivkraftsanalys
                    if (drivkraft)
                    {

                        bool firstActor = true;

                        List<string> requiredAspects = new List<string>() { "epsi", "loyal" };
                        List<string> requiredDimensions = new List<string>() { "image", "expect", "prodq", "servq", "value" };

                        List<string> selectedAspects = new List<string>();
                        List<string> selectedDimensions = new List<string>();

                        foreach (DataGridViewRow resultat in dgvResultatVariabler.Rows)
                        {
                            if (resultat.Cells["dgvResultatVariablerNamn"].Value != null)
                            {
                                selectedAspects.Add(resultat.Cells["dgvResultatVariablerNamn"].Value.ToString().ToLower());
                            }
                        }

                        foreach (DataGridViewRow resultat in dgvKundDimensioner.Rows)
                        {
                            if (resultat.Cells["dataGridViewKundDimensionerNamn"].Value != null)
                            {
                                selectedDimensions.Add(resultat.Cells["dataGridViewKundDimensionerNamn"].Value.ToString().ToLower().ToString());
                            }
                        }

                        bool useSkattadModell = requiredAspects.All(x => selectedAspects.Contains(x)) && requiredDimensions.All(x => selectedDimensions.Contains(x));

                        Dictionary<string, string> questionTexts = Functions.GetAspectQuestions(this.SelectedPath, 2);

                        foreach (var actor in Functions.GetActors(SelectedPath))
                        {

                            if (!firstActor)
                            {
                                CreateNewSlides(new List<int>() { loopIndex - 2, loopIndex - 1, loopIndex }, presPart, item, slide, items);
                                loopIndex++;
                                item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                if (item == null)
                                    break;
                                part = presPart.GetPartById(item.RelationshipId);
                                slide = (part as SlidePart).Slide;
                            }

                            shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                            foreach (var shape in shapes)
                            {
                                SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                SetText(shape, cbSegment.Text, "{Segment}");
                                SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                SetText(shape, tbOrt.Text, "{Ort}");
                                SetText(shape, dateTimePicker.Text, "{Datum}");
                                SetText(shape, actor, "{Aktör}");
                            }

                            #region Skattad strukturmodell
                            if (useSkattadModell || true)
                            {
                                var aspects = Functions.GetAspectsByActor(actor, SelectedPath);
                                var gs = slide.CommonSlideData.ShapeTree.Elements<GroupShape>().FirstOrDefault();
                                var groupedShapes = gs.Elements<Shape>();
                                var allAspects = new List<string>();
                                allAspects.AddRange(selectedAspects);
                                allAspects.AddRange(selectedDimensions);
                                foreach (var aspect in allAspects)
                                {
                                    var placeHolder = groupedShapes.FirstOrDefault(x => x.InnerText.ToLower() == $"{aspect.ToLower()}value");

                                    var value = aspects.FirstOrDefault(x => x.Item1 == aspect)?.Item2 ?? 0.0;

                                    placeHolder.TextBody.Elements<D.Paragraph>().FirstOrDefault().Elements<D.Run>().FirstOrDefault().Text = new D.Text(value.ToString("n1"));

                                    var coefs = Functions.GetCoefs(actor, aspect, SelectedPath);

                                    foreach (var coef in coefs.Where(x => x.Item2 != 0.0))
                                    {
                                        var placeHolderCoef = groupedShapes.FirstOrDefault(x => x.InnerText.ToLower() == $"{aspect.ToLower()}>{coef.Item1.ToLower()}");
                                        var paragraph = placeHolderCoef.TextBody.Elements<D.Paragraph>().FirstOrDefault();
                                        var runs = paragraph.Elements<D.Run>();
                                        runs.FirstOrDefault().Text = new D.Text(coef.Item2.ToString("n2"));

                                        while (runs.Count() > 1)
                                        {
                                            paragraph.RemoveChild(runs.Last());
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region IPMA

                            loopIndex++;
                            item = (SlideId)items.ElementAtOrDefault(loopIndex);
                            if (item == null)
                                break;
                            part = presPart.GetPartById(item.RelationshipId);
                            slide = (part as SlidePart).Slide;

                            shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                            foreach (var shape in shapes)
                            {
                                SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                SetText(shape, cbSegment.Text, "{Segment}");
                                SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                SetText(shape, tbOrt.Text, "{Ort}");
                                SetText(shape, dateTimePicker.Text, "{Datum}");
                                SetText(shape, actor, "{Aktör}");
                            }

                            var ipmaData = Functions.GetIPMAData(SelectedPath, actor);

                            var ipmaChartPart = slide.SlidePart.ChartParts.FirstOrDefault();

                            if (ipmaChartPart != null)
                            {
                                D.Charts.Chart chart = ipmaChartPart.ChartSpace.GetFirstChild<D.Charts.Chart>();

                                var plotArea = chart.GetFirstChild<D.Charts.PlotArea>();

                                var bubbleChart = plotArea.GetFirstChild<D.Charts.BubbleChart>();

                                foreach (var series in bubbleChart.Elements<D.Charts.BubbleChartSeries>())
                                {
                                    var numberXReference = series.GetFirstChild<D.Charts.XValues>()?.GetFirstChild<D.Charts.NumberReference>();
                                    var numberXRange = numberXReference?.GetFirstChild<D.Charts.Formula>();
                                    {
                                        var xFormula = numberXRange.Text;
                                        var startAndEnd = xFormula.Split(new char[] { ':' });
                                        var startParts = startAndEnd.First().Split(new char[] { '$' });
                                        int start = int.Parse(startParts.Last());
                                        int end = start + (ipmaData.Count - 1);

                                        xFormula = xFormula.Substring(0, xFormula.LastIndexOf("$")) + "$" + end.ToString();
                                        numberXRange.Text = xFormula;
                                    }
                                    var numberXCache = numberXReference?.GetFirstChild<D.Charts.NumberingCache>();

                                    var numberYReference = series.GetFirstChild<D.Charts.YValues>()?.GetFirstChild<D.Charts.NumberReference>();
                                    var numberYRange = numberYReference?.GetFirstChild<D.Charts.Formula>();
                                    {
                                        var yFormula = numberYRange.Text;
                                        var startAndEnd = yFormula.Split(new char[] { ':' });
                                        var startParts = startAndEnd.First().Split(new char[] { '$' });
                                        int start = int.Parse(startParts.Last());
                                        int end = start + (ipmaData.Count - 1);

                                        yFormula = yFormula.Substring(0, yFormula.LastIndexOf("$")) + "$" + end.ToString();
                                        numberXRange.Text = yFormula;
                                    }
                                    var numberYCache = numberYReference?.GetFirstChild<D.Charts.NumberingCache>();

                                    var bubbleSizeReference = series.GetFirstChild<D.Charts.BubbleSize>()?.GetFirstChild<D.Charts.NumberReference>();
                                    var bubbleSizeRange = bubbleSizeReference?.GetFirstChild<D.Charts.Formula>();
                                    {
                                        var bubbleFormula = bubbleSizeRange.Text;
                                        var startAndEnd = bubbleFormula.Split(new char[] { ':' });
                                        var startParts = startAndEnd.First().Split(new char[] { '$' });
                                        int start = int.Parse(startParts.Last());
                                        int end = start + (ipmaData.Count - 1);

                                        bubbleFormula = bubbleFormula.Substring(0, bubbleFormula.LastIndexOf("$")) + "$" + end.ToString();
                                        numberXRange.Text = bubbleFormula;
                                    }
                                    var bubbleSizeCache = bubbleSizeReference?.GetFirstChild<D.Charts.NumberingCache>();

                                    DocumentFormat.OpenXml.Office2013.Drawing.Chart.DataLabelsRangeChache dlrCache = new DocumentFormat.OpenXml.Office2013.Drawing.Chart.DataLabelsRangeChache();

                                    var extensionList = series.GetFirstChild<D.Charts.BubbleSerExtensionList>();
                                    if(extensionList != null)
                                    {
                                        var extension = extensionList.Elements<D.Charts.BubbleSerExtension>().ElementAtOrDefault(1);
                                        if(extension != null)
                                        {
                                            var dataLabelsRange = extension.GetFirstChild<DocumentFormat.OpenXml.Office2013.Drawing.Chart.DataLabelsRange>();
                                            if(dataLabelsRange != null)
                                            {
                                                var dataLabelsRangeCache = dataLabelsRange.GetFirstChild<DocumentFormat.OpenXml.Office2013.Drawing.Chart.DataLabelsRangeChache>();
                                                if(dataLabelsRangeCache != null)
                                                {
                                                    dlrCache = dataLabelsRangeCache;
                                                }
                                            }
                                        }
                                    }

                                    numberXCache.RemoveAllChildren();
                                    numberYCache.RemoveAllChildren();
                                    bubbleSizeCache.RemoveAllChildren();
                                    dlrCache.RemoveAllChildren();

                                    int ncIndex = 0;
                                    
                                    foreach (var ipma in ipmaData)
                                    {
                                        var currentXNP = new D.Charts.NumericPoint();
                                        currentXNP.NumericValue = new D.Charts.NumericValue(ipma.Importance.ToString("n1").Replace(",", "."));
                                        currentXNP.Index = (uint)ncIndex;
                                        numberXCache.AppendChild(currentXNP);

                                        var currentBubbleSize = new D.Charts.NumericPoint();
                                        currentBubbleSize.NumericValue = new D.Charts.NumericValue(ipma.Importance.ToString("n1").Replace(",", "."));
                                        currentBubbleSize.Index = (uint)ncIndex;
                                        bubbleSizeCache.AppendChild(currentBubbleSize);

                                        var currentYNP = new D.Charts.NumericPoint();
                                        currentYNP.NumericValue = new D.Charts.NumericValue(ipma.Performance.ToString("n2").Replace(",", "."));
                                        currentYNP.Index = (uint)ncIndex;
                                        numberYCache.AppendChild(currentYNP);

                                        string aspectText = aspectNames.ContainsKey(ipma.Aspect) ? aspectNames[ipma.Aspect] : ipma.Aspect;
                                        var currentSP = new D.Charts.StringPoint();
                                        currentSP.InnerXml = $"<c:v xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">{System.Security.SecurityElement.Escape(aspectText)}</c:v>";
                                        currentSP.Index = (uint)ncIndex;
                                        dlrCache.AppendChild(currentSP);

                                        ncIndex++;
                                    }
                                    
                                }

                                EmbeddedPackagePart embeddedExcel = ipmaChartPart.EmbeddedPackagePart;
                                if (embeddedExcel != null)
                                {
                                    using (Stream str = embeddedExcel.GetStream())
                                    {
                                        using (SpreadsheetDocument xls = SpreadsheetDocument.Open(str, true))
                                        {
                                            var sheet = xls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault();
                                            if (sheet != null)
                                            {
                                                Worksheet ws = ((WorksheetPart)xls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                                                var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                                                if (sheetData != null)
                                                {
                                                    int rowIndex = 3;
                                                    var rows = sheetData.Elements<Row>();

                                                    foreach (var ipma in ipmaData)
                                                    {
                                                        var currentRow = rows.FirstOrDefault(x => x.RowIndex == rowIndex);
                                                        if (currentRow == null)
                                                        {
                                                            currentRow = sheetData.InsertAt(new Row() { RowIndex = (uint)rowIndex }, sheetData.Elements<Row>().Count());
                                                            var labelCell = currentRow.InsertAt(new Cell() { CellReference = "B" + rowIndex }, 0);
                                                            labelCell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);
                                                            var valueCell = currentRow.InsertAt(new Cell() { CellReference = "C" + rowIndex }, 1);
                                                            valueCell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                                                        }
                                                        var cells = currentRow.Elements<Cell>();

                                                        string aspectText = aspectNames.ContainsKey(ipma.Aspect) ? aspectNames[ipma.Aspect] : ipma.Aspect;
                                                        cells.ElementAt(0).CellValue = new CellValue(aspectText);
                                                        cells.ElementAt(0).DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String);

                                                        cells.ElementAt(1).CellValue = new CellValue(ipma.Performance.ToString("n1").Replace(",", "."));
                                                        cells.ElementAt(2).CellValue = new CellValue(ipma.Importance.ToString("n2").Replace(",", "."));

                                                        rowIndex++;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion


                            loopIndex++;
                            item = (SlideId)items.ElementAtOrDefault(loopIndex);
                            if (item == null)
                                break;
                            part = presPart.GetPartById(item.RelationshipId);
                            slide = (part as SlidePart).Slide;

                            shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                            foreach (var shape in shapes)
                            {
                                SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                SetText(shape, cbSegment.Text, "{Segment}");
                                SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                SetText(shape, tbOrt.Text, "{Ort}");
                                SetText(shape, dateTimePicker.Text, "{Datum}");
                                SetText(shape, actor, "{Aktör}");
                            }

                            var chartParts = slide.SlidePart.ChartParts;

                            int chartIndex = chartParts.Count() - 1;

                            //vi enumererar diagrammen baklänges, så textrutorna måste också vara baklänges                            
                            List<string> textBoxes = new List<string>() { "textruta 28", "textruta 27", "textruta 33", "textruta 29", "textruta 20" };

                            int textBoxIndex = textBoxes.Count - 1;

                            foreach (DataGridViewRow dgvRow in dgvKundDimensioner.Rows)
                            {
                                if (dgvKundDimensioner.Rows.Count > 1 && dgvRow.Cells["dataGridViewKundDimensionerNamn"].Value == null)
                                {
                                    break;
                                }
                                string dataName = dgvRow.Cells["dataGridViewKundDimensionerNamn"].Value?.ToString().ToLower() ?? "image";

                                List<KeyValuePair<string, double>> qWeights = new List<KeyValuePair<string, double>>();

                                int qIndex = 1;
                                int aspectIndex = 2;
                                int weightIndex = 3;

                                if (File.Exists(SelectedPath + "\\Output\\" + actor + ".xlsx"))
                                {
                                    using (FileStream fs = new FileStream(SelectedPath + "\\Output\\" + actor + ".xlsx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                    {
                                        using (SpreadsheetDocument extXls = SpreadsheetDocument.Open(fs, false))
                                        {
                                            var sheet = extXls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(x => x.Name == "outer_model");
                                            if (sheet != null)
                                            {
                                                Worksheet ws = ((WorksheetPart)extXls.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
                                                var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                                                if (sheetData != null)
                                                {
                                                    foreach (var row in sheetData.Elements<Row>())
                                                    {
                                                        var cells = row.Elements<Cell>();
                                                        var aspectCell = cells.ElementAtOrDefault(aspectIndex);
                                                        if (aspectCell != null)
                                                        {
                                                            if (Functions.GetTextFromSharedTable(extXls, aspectCell) == dataName)
                                                            {
                                                                string q = string.Empty;
                                                                var qCell = cells.ElementAtOrDefault(qIndex);
                                                                if (qCell != null)
                                                                {
                                                                    q = Functions.GetTextFromSharedTable(extXls, qCell);
                                                                }

                                                                double weight = -1.0;
                                                                var weightCell = cells.ElementAtOrDefault(weightIndex);
                                                                if (weightCell != null)
                                                                {
                                                                    weight = double.Parse(weightCell.CellValue.Text, CultureInfo.InvariantCulture);
                                                                }

                                                                if (q != string.Empty && weight > -1.0)
                                                                {
                                                                    qWeights.Add(new KeyValuePair<string, double>(q, weight));
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    var textBoxShape = slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault(x => x.NonVisualShapeProperties.NonVisualDrawingProperties.Name == textBoxes[textBoxIndex]);
                                    var textBody = textBoxShape.Elements<TextBody>().FirstOrDefault();
                                    if (textBody != null)
                                    {
                                        var lastParagraph = textBody.Elements<D.Paragraph>().LastOrDefault();
                                        if (lastParagraph != null)
                                        {
                                            var run = lastParagraph.Elements<D.Run>().FirstOrDefault();
                                            if (run != null)
                                            {
                                                StringBuilder aspectTextBuilder = new StringBuilder();
                                                foreach(var qWeight in qWeights)
                                                {
                                                    aspectTextBuilder.AppendLine($"{qWeight.Key}: {questionTexts[qWeight.Key]}");
                                                }

                                                run.Text = new D.Text(aspectTextBuilder.ToString());
                                            }
                                        }
                                    }

                                    textBoxIndex--;

                                    var chartPart = chartParts.ElementAtOrDefault(chartIndex);
                                    D.Charts.Chart chart = chartPart.ChartSpace.GetFirstChild<D.Charts.Chart>();

                                    chart.Title.ChartText.RichText.Elements<D.Paragraph>().ElementAt(1).Elements<D.Run>().First().Text = new D.Text(actor);

                                    var plotArea = chart.GetFirstChild<D.Charts.PlotArea>();
                                    var radarChart = plotArea.GetFirstChild<D.Charts.RadarChart>();


                                    foreach (var series in radarChart.Elements<D.Charts.RadarChartSeries>())
                                    {
                                        var numberReference = series.GetFirstChild<D.Charts.Values>()?.GetFirstChild<D.Charts.NumberReference>();
                                        var numberRange = numberReference?.GetFirstChild<D.Charts.Formula>();
                                        var numberCache = numberReference?.GetFirstChild<D.Charts.NumberingCache>();
                                        var numericPoints = numberCache?.Elements<D.Charts.NumericPoint>();

                                        var stringReference = series.GetFirstChild<D.Charts.CategoryAxisData>()?.GetFirstChild<D.Charts.StringReference>();
                                        var textRange = stringReference?.GetFirstChild<D.Charts.Formula>();
                                        var textCache = stringReference?.GetFirstChild<D.Charts.StringCache>();
                                        var stringPoints = textCache?.Elements<D.Charts.StringPoint>();

                                        textCache.RemoveAllChildren();
                                        numberCache.RemoveAllChildren();

                                        int ncIndex = 0;

                                        foreach (var v in qWeights)
                                        {
                                            D.Charts.NumericPoint currentNP = numericPoints.ElementAtOrDefault(ncIndex);
                                            if (currentNP == null)
                                            {
                                                currentNP = new D.Charts.NumericPoint();
                                                currentNP.NumericValue = new D.Charts.NumericValue(v.Value.ToString("n2").Replace(",", "."));
                                                currentNP.Index = (uint)ncIndex;
                                                numberCache.AppendChild(currentNP);
                                            }
                                            else
                                            {
                                                currentNP.NumericValue = new D.Charts.NumericValue(v.Value.ToString("n2").Replace(",", "."));
                                            }

                                            D.Charts.StringPoint currentSP = stringPoints.ElementAtOrDefault(ncIndex);
                                            if (currentSP == null)
                                            {
                                                currentSP = new D.Charts.StringPoint();
                                                currentSP.InnerXml = $"<c:v xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">{System.Security.SecurityElement.Escape(v.Key)}</c:v>";
                                                currentSP.Index = (uint)ncIndex;
                                                textCache.AppendChild(currentSP);
                                            }
                                            else
                                            {
                                                currentSP.InnerXml = $"<c:v xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">{System.Security.SecurityElement.Escape(v.Key)}</c:v>";
                                            }
                                            ncIndex++;
                                        }

                                        int valueIndex = 0;

                                        EmbeddedPackagePart embeddedExcel = chartPart.EmbeddedPackagePart;
                                        if (embeddedExcel != null)
                                        {
                                            using (Stream str = embeddedExcel.GetStream())
                                            {
                                                using (SpreadsheetDocument xls = SpreadsheetDocument.Open(str, true))
                                                {
                                                    var sheet = xls.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(x => x.Name == "spindel");
                                                    if (sheet != null)
                                                    {
                                                        WorksheetPart wsPart = (WorksheetPart)xls.WorkbookPart.GetPartById(sheet.Id);
                                                        Worksheet ws = wsPart.Worksheet;
                                                        var sheetData = ws.Elements<SheetData>().FirstOrDefault();
                                                        if (sheetData != null)
                                                        {
                                                            var reference = numberReference.Formula.Text.Split(new char[] { '!' }).Last().Replace("$", string.Empty).Split(new char[] { ':' });
                                                            string start = reference.First();
                                                            string stop = reference.Last();

                                                            int rowIndex = int.Parse(start.Substring(1));
                                                            int lastRowIndex = rowIndex + qWeights.Count;
                                                            valueIndex = 0;
                                                            string colRef = start.Substring(0, 1);

                                                            numberReference.Formula.Text = numberReference.Formula.Text.Substring(0, numberReference.Formula.Text.LastIndexOf("$")) + (lastRowIndex - 1);

                                                            var rows = sheetData.Elements<Row>();
                                                            while (rowIndex < lastRowIndex)
                                                            {
                                                                var currentRow = rows.FirstOrDefault(x => x.RowIndex == rowIndex);
                                                                if (currentRow == null)
                                                                {
                                                                    break;
                                                                }
                                                                var cells = currentRow.Elements<Cell>();
                                                                var currentCell = cells.FirstOrDefault(x => x.CellReference.Value == colRef + rowIndex.ToString());
                                                                currentCell.CellValue = new CellValue(qWeights.ElementAt(valueIndex++).Value.ToString("n2").Replace(",", "."));
                                                                rowIndex++;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    chartIndex--;
                                }
                            }
                            firstActor = false;
                        }
                        drivkraft = false;
                    }
                    #endregion

                    #region Klagomål

                    if (complaints)
                    {
                        GraphicFrame complaintGraphicFrame = slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>().FirstOrDefault();
                        if (complaintGraphicFrame != null)
                        {
                            var complaintsTable = complaintGraphicFrame.GetFirstChild<D.Graphic>().GraphicData.GetFirstChild<D.Table>();

                            var rows = complaintsTable.Elements<D.TableRow>();
                            var headerRow = rows.ElementAt(1);
                            var rowToUse = rows.ElementAt(2).CloneNode(true);

                            while (rows.Count() > 2)
                            {
                                complaintsTable.RemoveChild(rows.Last());
                            }

                            DataTable data = Functions.GetComplaintsData(SelectedPath);

                            D.TableCell headerCellClone = (D.TableCell)rows.ElementAt(1).Elements<D.TableCell>().ElementAt(1).CloneNode(true);
                            //foreach(D.TableCell cell in rows.ElementAt(1).Elements<D.TableCell>().Skip(1))
                            int cellIdx = 0;
                            var headerCells = rows.ElementAt(1).Elements<D.TableCell>();

                            var headerExtList = headerRow.GetFirstChild<D.ExtensionList>();

                            foreach (DataColumn dataCol in data.Columns)
                            {
                                D.TableCell currentCell = new D.TableCell();
                                if (cellIdx < headerCells.Count())
                                {
                                    currentCell = headerCells.ElementAt(cellIdx);
                                }
                                else
                                {
                                    currentCell = (D.TableCell)headerCellClone.CloneNode(true);
                                    headerRow.InsertBefore(currentCell, headerExtList);
                                }
                                Functions.SetTableCellText(currentCell, dataCol.ColumnName);
                                cellIdx++;
                            }

                            foreach (DataRow dataRow in data.Rows)
                            {                                
                                
                                var newRow = rowToUse.CloneNode(true);
                                var extList = newRow.GetFirstChild<D.ExtensionList>();
                                var newCell = (D.TableCell)newRow.Elements<D.TableCell>().ElementAt(1).CloneNode(true);
                                var cells = newRow.Elements<D.TableCell>();
                                //Functions.SetTableCellText(cells.First(), dataRow[0].ToString());
                                for (int i = 0; i < dataRow.ItemArray.Length; i++)
                                {
                                    D.TableCell currentCell = new D.TableCell();
                                    if(i < cells.Count())
                                    {
                                        currentCell = cells.ElementAt(i);
                                    }
                                    else
                                    {
                                        currentCell = (D.TableCell)newCell.CloneNode(true);
                                        newRow.InsertBefore(currentCell, extList);
                                    }

                                    Functions.SetTableCellText(currentCell, dataRow[i].ToString());
                                }
                                complaintsTable.AppendChild(newRow);
                            }
                            int cellCount = headerRow.Elements<D.TableCell>().Count();
                            var tableGrid = complaintsTable.GetFirstChild<D.TableGrid>();
                            if (complaintsTable.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() < cellCount)
                            {
                                while (complaintsTable.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() < cellCount)
                                {
                                    D.GridColumn newCol = (D.GridColumn)complaintsTable.GetFirstChild<D.TableGrid>().GetFirstChild<D.GridColumn>().CloneNode(true);
                                    tableGrid.AppendChild(newCol);

                                    D.TableCell newHeaderCell = (D.TableCell)complaintsTable.GetFirstChild<D.TableRow>().ElementAt(1).CloneNode(true);
                                    var topHeaderExtList = complaintsTable.GetFirstChild<D.TableRow>().GetFirstChild<D.ExtensionList>();
                                    complaintsTable.GetFirstChild<D.TableRow>().InsertBefore(newHeaderCell, topHeaderExtList);
                                }
                            }
                            else if (complaintsTable.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() > cellCount)
                            {
                                while (complaintsTable.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() > cellCount)
                                {
                                    tableGrid.RemoveChild(tableGrid.Elements<D.GridColumn>().Last());
                                }
                            }

                            complaintsTable.GetFirstChild<D.TableRow>().GetFirstChild<D.TableCell>().GridSpan.Value = cellCount;
                            

                        }
                        complaints = false;
                    }

                    #endregion

                    #region Branschspecifika frågor
                    if (branschSpecifika)
                    {
                        bool isFirstSheet = true;
                        bool isAndelar = false;
                        bool isQuestionRow = false;
                        bool resizeGrid = true;
                        int rowIndex = 0;
                        
                        
                        var IBRData = Functions.GetSheetAsTables(SelectedPath, "TabellerIBR.xlsx", "IS");

                        foreach (DataTable dataTable in IBRData.Tables)
                        {
                            //List<D.Table> branschSpecifikTables = new List<D.Table>();
                            if (!isAndelar && !isFirstSheet)
                            {
                                CreateNewSlides(new List<int>() { loopIndex }, presPart, item, slide, items); //skapa ny slide
                                loopIndex++;
                                item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                if (item == null)
                                    break;
                                part = presPart.GetPartById(item.RelationshipId);
                                slide = (part as SlidePart).Slide;

                                shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                                foreach (var shape in shapes)
                                {
                                    SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                    SetText(shape, cbSegment.Text, "{Segment}");
                                    SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                    SetText(shape, tbOrt.Text, "{Ort}");
                                    SetText(shape, dateTimePicker.Text, "{Datum}");
                                }                                
                            }

                            D.Table andelar = null;
                            D.Table nöjdhet = null;

                            var branschSpecifikGraphicFrames = slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>();
                            if (branschSpecifikGraphicFrames != null)
                            {
                                foreach (var branschSpecifikGraphicFrame in branschSpecifikGraphicFrames)
                                {
                                    if(branschSpecifikGraphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == "Tabell 5")
                                    {
                                        andelar = branschSpecifikGraphicFrame.Descendants<D.Table>().FirstOrDefault();
                                    }
                                    else if(branschSpecifikGraphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == "Tabell 6")
                                    {
                                        nöjdhet = branschSpecifikGraphicFrame.Descendants<D.Table>().FirstOrDefault();
                                    }
                                }
                            }
                            //tabell 5 = andelar, tabell 6 = nöjdhet
                            /*var andelar = branschSpecifikTables.FirstOrDefault(x => x.InnerText.Contains("Andelar"));
                            var nöjdhet = branschSpecifikTables.FirstOrDefault(x => x.InnerText.Contains("Nöjdhet"));*/

                            resizeGrid = true;
                            isAndelar = !isAndelar;

                            var tableToUse = isAndelar ? andelar : nöjdhet;

                            string question = dataTable.TableName.TrimStart(new char[] { '!' });
                            Functions.SetTableCellText(tableToUse.Elements<D.TableRow>().FirstOrDefault().Elements<D.TableCell>().FirstOrDefault(), question + (isAndelar ? ": Andelar" : ": Nöjdhet"));

                            D.TableRow questionRow = tableToUse.Elements<D.TableRow>().ElementAt(1);
                            var cells = questionRow.Elements<D.TableCell>();
                            D.TableCell clonedCell = (D.TableCell)cells.ElementAt(1).CloneNode(true);
                            while (questionRow.Elements<D.TableCell>().Count() > 1)
                            {
                                questionRow.RemoveChild(questionRow.Elements<D.TableCell>().Last());
                            }
                            var extListQ = questionRow.GetFirstChild<D.ExtensionList>();

                            for (int colIdx = 2; colIdx < dataTable.Columns.Count; colIdx++)
                            {
                                DataColumn col = dataTable.Columns[colIdx];
                                D.TableCell newCell = (D.TableCell)clonedCell.CloneNode(true);

                                Functions.SetTableCellText(newCell, col.ColumnName);

                                questionRow.InsertBefore(newCell, extListQ);
                            }

                            var clonedRow = tableToUse.Elements<D.TableRow>().ElementAtOrDefault(2).CloneNode(true); //data-rad
                            while (tableToUse.Elements<D.TableRow>().Count() > 2)
                            {
                                tableToUse.RemoveChild(tableToUse.Elements<D.TableRow>().Last());
                            }

                            foreach (DataRow row in dataTable.Rows)
                            {
                                var gridColumns = tableToUse.TableGrid.Elements<D.GridColumn>();

                                var dataValueCells = row.ItemArray;
                                D.TableRow tableRow = (D.TableRow)clonedRow.CloneNode(true);
                                var extList = tableRow.GetFirstChild<D.ExtensionList>();
                                var dataCellToUse = tableRow.ElementAt(1).CloneNode(true);

                                var tableCells = tableRow.Elements<D.TableCell>();
                                Functions.SetTableCellText(tableCells.First(), dataValueCells[0].ToString());

                                while (tableRow.Elements<D.TableCell>().Count() > 1)
                                {
                                    tableRow.RemoveChild(tableRow.Elements<D.TableCell>().Last());
                                }
                                int cellCount = 1;
                                foreach (var dataCell in dataValueCells.Skip(2))
                                {
                                    D.TableCell newDataCell = (D.TableCell)dataCellToUse.CloneNode(true);

                                    string format = isAndelar ? "p0" : "n1";
                                    string formattedValue = "-";
                                    try
                                    {
                                        formattedValue = double.Parse(dataCell.ToString(), CultureInfo.InvariantCulture).ToString(format);
                                    }
                                    catch
                                    {
                                        ;
                                    }
                                    Functions.SetTableCellText(newDataCell, formattedValue);
                                    tableRow.InsertBefore(newDataCell, extList);
                                    cellCount++;
                                }

                                tableToUse.AppendChild(tableRow);

                                if (resizeGrid)
                                {
                                    var tableGrid = tableToUse.GetFirstChild<D.TableGrid>();
                                    if (tableToUse.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() < cellCount)
                                    {
                                        while (tableToUse.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() < cellCount)
                                        {
                                            D.GridColumn newCol = (D.GridColumn)tableToUse.GetFirstChild<D.TableGrid>().GetFirstChild<D.GridColumn>().CloneNode(true);
                                            tableGrid.AppendChild(newCol);

                                            D.TableCell newHeaderCell = (D.TableCell)tableToUse.GetFirstChild<D.TableRow>().ElementAt(1).CloneNode(true);
                                            var headerExtList = tableToUse.GetFirstChild<D.TableRow>().GetFirstChild<D.ExtensionList>();
                                            tableToUse.GetFirstChild<D.TableRow>().InsertBefore(newHeaderCell, headerExtList);
                                        }
                                    }
                                    else if (tableToUse.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() > cellCount)
                                    {
                                        while (tableToUse.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() > cellCount)
                                        {
                                            tableGrid.RemoveChild(tableGrid.Elements<D.GridColumn>().Last());
                                        }
                                    }

                                    tableToUse.GetFirstChild<D.TableRow>().GetFirstChild<D.TableCell>().GridSpan.Value = cellCount;
                                    resizeGrid = false;
                                }
                            }
                            isFirstSheet = false;
                        }

                        branschSpecifika = false;
                        moreBranschSpecifika = true;

                        CreateNewSlides(new List<int>() { loopIndex }, presPart, item, slide, items); //skapa ny slide
                        loopIndex++;
                        item = (SlideId)items.ElementAtOrDefault(loopIndex);
                        if (item == null)
                            break;
                        part = presPart.GetPartById(item.RelationshipId);
                        slide = (part as SlidePart).Slide;

                        shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                        foreach (var shape in shapes)
                        {
                            SetText(shape, cbBranscher.Text, "{BRANSCH}");
                            SetText(shape, cbSegment.Text, "{Segment}");
                            SetText(shape, numYear.Value.ToString(), "{ÅR}");
                            SetText(shape, tbOrt.Text, "{Ort}");
                            SetText(shape, dateTimePicker.Text, "{Datum}");
                        }
                    }

                    #endregion

                    #region Mer branschspecifika
                    if (moreBranschSpecifika)
                    {
                        bool addSlide = false;
                        bool isQuestionRow = false;
                        bool resizeGrid = true;
                        int rowIndex = 0;

                        List<D.Table> branschSpecifikTables = new List<D.Table>();

                        var branschSpecifikGraphicFrames = slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>();
                        if (branschSpecifikGraphicFrames != null)
                        {
                            foreach (var branschSpecifikGraphicFrame in branschSpecifikGraphicFrames)
                            {
                                if(branschSpecifikGraphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == "Tabell 6")
                                {
                                    slide.CommonSlideData.ShapeTree.RemoveChild(branschSpecifikGraphicFrame);
                                }
                                else
                                {
                                    branschSpecifikTables.Add(branschSpecifikGraphicFrame.Descendants<D.Table>().FirstOrDefault());
                                }
                                //var graphics = branschSpecifikGraphicFrame.Elements<D.Graphic>();
                                //foreach (var graphic in graphics)
                                //{
                                //    var tables = graphic.GraphicData.Elements<D.Table>();
                                //    foreach (var table in tables)
                                //    {
                                //        if (table.InnerText.Contains("Nöjdhet"))
                                //        {
                                //            slide.CommonSlideData.ShapeTree.RemoveChild(branschSpecifikGraphicFrame);
                                //        }
                                //        else
                                //        {
                                //            branschSpecifikTables.Add(table);
                                //        }
                                //    }
                                //}
                            }

                            var andelar = branschSpecifikTables.FirstOrDefault();
                            if (andelar != null)
                            {
                                var IBRData = Functions.GetSheetAsTables(SelectedPath, "TabellerIBR.xlsx", "IS index");
                                bool firstTable = true;
                                foreach (DataTable dataTable in IBRData.Tables)
                                {
                                    if (!firstTable)
                                    {
                                        resizeGrid = true;
                                        CreateNewSlides(new List<int>() { loopIndex }, presPart, item, slide, items); //skapa ny slide
                                        loopIndex++;
                                        item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                        if (item == null)
                                            break;
                                        part = presPart.GetPartById(item.RelationshipId);
                                        slide = (part as SlidePart).Slide;

                                        shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                                        foreach (var shape in shapes)
                                        {
                                            SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                            SetText(shape, cbSegment.Text, "{Segment}");
                                            SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                            SetText(shape, tbOrt.Text, "{Ort}");
                                            SetText(shape, dateTimePicker.Text, "{Datum}");
                                        }
                                    }

                                    branschSpecifikTables.Clear();

                                    branschSpecifikGraphicFrames = slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>();
                                    if (branschSpecifikGraphicFrames != null)
                                    {
                                        foreach (var branschSpecifikGraphicFrame in branschSpecifikGraphicFrames)
                                        {
                                            if (branschSpecifikGraphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == "Tabell 6")
                                            {
                                                slide.CommonSlideData.ShapeTree.RemoveChild(branschSpecifikGraphicFrame);
                                            }
                                            else
                                            {
                                                branschSpecifikTables.Add(branschSpecifikGraphicFrame.Descendants<D.Table>().FirstOrDefault());
                                            }


                                            //var graphics = branschSpecifikGraphicFrame.Elements<D.Graphic>();
                                            //foreach (var graphic in graphics)
                                            //{
                                            //    var tables = graphic.GraphicData.Elements<D.Table>();
                                            //    foreach (var table in tables)
                                            //    {
                                            //        if (table.InnerText.Contains("Nöjdhet"))
                                            //        {
                                            //            slide.CommonSlideData.ShapeTree.RemoveChild(branschSpecifikGraphicFrame);
                                            //        }
                                            //        else
                                            //        {
                                            //            branschSpecifikTables.Add(table);
                                            //        }
                                            //    }
                                            //}
                                        }
                                    }
                                    andelar = branschSpecifikTables.FirstOrDefault();
                                    //sätt rubrik
                                    string question = dataTable.TableName.TrimStart(new char[] {'!'});
                                    Functions.SetTableCellText(andelar.Elements<D.TableRow>().FirstOrDefault().Elements<D.TableCell>().FirstOrDefault(), question + " Andelar");

                                    //sätt kolumnnamn
                                    D.TableRow questionRow = andelar.Elements<D.TableRow>().ElementAt(1);
                                    var cells = questionRow.Elements<D.TableCell>();
                                    D.TableCell clonedCell = (D.TableCell)cells.ElementAt(1).CloneNode(true);
                                    while (questionRow.Elements<D.TableCell>().Count() > 1)
                                    {
                                        questionRow.RemoveChild(questionRow.Elements<D.TableCell>().Last());
                                    }
                                    var extListQ = questionRow.GetFirstChild<D.ExtensionList>();

                                    for (int colIdx = 2; colIdx < dataTable.Columns.Count; colIdx++)
                                    {
                                        DataColumn col = dataTable.Columns[colIdx];
                                        D.TableCell newCell = (D.TableCell)clonedCell.CloneNode(true);

                                        Functions.SetTableCellText(newCell, col.ColumnName);

                                        questionRow.InsertBefore(newCell, extListQ);
                                    }

                                    var clonedRow = andelar.Elements<D.TableRow>().ElementAtOrDefault(2).CloneNode(true); //data-rad
                                    while (andelar.Elements<D.TableRow>().Count() > 2)
                                    {
                                        andelar.RemoveChild(andelar.Elements<D.TableRow>().Last());
                                    }

                                    var tableToUse = andelar;
                                    foreach (DataRow row in dataTable.Rows)
                                    {
                                        var gridColumns = tableToUse.TableGrid.Elements<D.GridColumn>();

                                        var dataValueCells = row.ItemArray;
                                        D.TableRow tableRow = (D.TableRow)clonedRow.CloneNode(true);
                                        var extList = tableRow.GetFirstChild<D.ExtensionList>();
                                        var dataCellToUse = tableRow.ElementAt(1).CloneNode(true);

                                        var tableCells = tableRow.Elements<D.TableCell>();
                                        Functions.SetTableCellText(tableCells.First(), dataValueCells[0].ToString());

                                        while (tableRow.Elements<D.TableCell>().Count() > 1)
                                        {
                                            tableRow.RemoveChild(tableRow.Elements<D.TableCell>().Last());
                                        }
                                        int cellCount = 1;
                                        foreach (var dataCell in dataValueCells.Skip(2))
                                        {
                                            D.TableCell newDataCell = (D.TableCell)dataCellToUse.CloneNode(true);

                                            string format = "n1";
                                            string formattedValue = double.Parse(dataCell.ToString(), CultureInfo.InvariantCulture).ToString(format);
                                            Functions.SetTableCellText(newDataCell, formattedValue);
                                            tableRow.InsertBefore(newDataCell, extList);
                                            cellCount++;
                                        }

                                        tableToUse.AppendChild(tableRow);

                                        if (resizeGrid)
                                        {
                                            var tableGrid = tableToUse.GetFirstChild<D.TableGrid>();
                                            if (tableToUse.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() < cellCount)
                                            {
                                                while (tableToUse.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() < cellCount)
                                                {
                                                    D.GridColumn newCol = (D.GridColumn)tableToUse.GetFirstChild<D.TableGrid>().GetFirstChild<D.GridColumn>().CloneNode(true);
                                                    tableGrid.AppendChild(newCol);

                                                    D.TableCell newHeaderCell = (D.TableCell)tableToUse.GetFirstChild<D.TableRow>().ElementAt(1).CloneNode(true);
                                                    var headerExtList = tableToUse.GetFirstChild<D.TableRow>().GetFirstChild<D.ExtensionList>();
                                                    tableToUse.GetFirstChild<D.TableRow>().InsertBefore(newHeaderCell, headerExtList);
                                                }
                                            }
                                            else if (tableToUse.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() > cellCount)
                                            {
                                                while (tableToUse.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() > cellCount)
                                                {
                                                    tableGrid.RemoveChild(tableGrid.Elements<D.GridColumn>().Last());

                                                }
                                            }
                                            tableToUse.GetFirstChild<D.TableRow>().GetFirstChild<D.TableCell>().GridSpan.Value = cellCount;
                                            resizeGrid = false;
                                        }
                                    }
                                    firstTable = false;
                                }
                            }
                        }
                        branschSpecifika = false;
                        moreBranschSpecifika = false;
                    }
                    #endregion

                    #region Segment
                    if (segment || branschSpecifika)
                    {
                        bool isAndelar = false;
                        bool resizeGrid = true;

                        List<D.Table> branschSpecifikTables = new List<D.Table>();
                        D.Table andelar = new D.Table();
                        D.Table nöjdhet = new D.Table();

                        var branschSpecifikGraphicFrames = slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>();
                        if (branschSpecifikGraphicFrames != null)
                        {
                            foreach (var branschSpecifikGraphicFrame in branschSpecifikGraphicFrames)
                            {
                                if (branschSpecifikGraphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == "Tabell 5")
                                {
                                    andelar = branschSpecifikGraphicFrame.Descendants<D.Table>().FirstOrDefault();
                                }
                                else
                                {
                                    nöjdhet = branschSpecifikGraphicFrame.Descendants<D.Table>().FirstOrDefault();
                                }
                                //var graphics = branschSpecifikGraphicFrame.Elements<D.Graphic>();
                                //foreach (var graphic in graphics)
                                //{
                                //    var tables = graphic.GraphicData.Elements<D.Table>();
                                //    foreach (var table in tables)
                                //    {
                                //        branschSpecifikTables.Add(table);
                                //    }
                                //}
                            }

                            //var andelar = branschSpecifikTables.FirstOrDefault(x => !x.InnerText.Contains("Nöjdhet"));
                            //var nöjdhet = branschSpecifikTables.FirstOrDefault(x => x.InnerText.Contains("Nöjdhet"));
                            if (andelar != null && nöjdhet != null)
                            {
                                string sheetName = branschSpecifika ? "IS" : "Background";
                                DataSet IBRData = Functions.GetSheetAsTables(SelectedPath, "TabellerIBR.xlsx", sheetName);

                                bool firstTable = true;
                                foreach (DataTable dataTable in IBRData.Tables)
                                {
                                    resizeGrid = true;
                                    isAndelar = !isAndelar;
                                    if (!firstTable && isAndelar)
                                    {
                                        CreateNewSlides(new List<int>() { loopIndex }, presPart, item, slide, items); //skapa ny slide
                                        loopIndex++;
                                        item = (SlideId)items.ElementAtOrDefault(loopIndex);
                                        if (item == null)
                                            break;
                                        part = presPart.GetPartById(item.RelationshipId);
                                        slide = (part as SlidePart).Slide;

                                        shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                                        foreach (var shape in shapes)
                                        {
                                            SetText(shape, cbBranscher.Text, "{BRANSCH}");
                                            SetText(shape, cbSegment.Text, "{Segment}");
                                            SetText(shape, numYear.Value.ToString(), "{ÅR}");
                                            SetText(shape, tbOrt.Text, "{Ort}");
                                            SetText(shape, dateTimePicker.Text, "{Datum}");
                                        }
                                    }

                                    branschSpecifikTables.Clear();

                                    branschSpecifikGraphicFrames = slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>();
                                    if (branschSpecifikGraphicFrames != null)
                                    {
                                        foreach (var branschSpecifikGraphicFrame in branschSpecifikGraphicFrames)
                                        {
                                            if (branschSpecifikGraphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == "Tabell 5")
                                            {
                                                andelar = branschSpecifikGraphicFrame.Descendants<D.Table>().FirstOrDefault();
                                            }
                                            else
                                            {
                                                nöjdhet = branschSpecifikGraphicFrame.Descendants<D.Table>().FirstOrDefault();
                                            }
                                            //var graphics = branschSpecifikGraphicFrame.Elements<D.Graphic>();
                                            //foreach (var graphic in graphics)
                                            //{
                                            //    var tables = graphic.GraphicData.Elements<D.Table>();
                                            //    foreach (var table in tables)
                                            //    {
                                            //        branschSpecifikTables.Add(table);
                                            //    }
                                            //}
                                        }
                                    }
                                    //andelar = branschSpecifikTables.FirstOrDefault(x => !x.InnerText.Contains("Nöjdhet"));
                                    //nöjdhet = branschSpecifikTables.FirstOrDefault(x => x.InnerText.Contains("Nöjdhet"));

                                    var tableToUse = (isAndelar ? andelar : nöjdhet);

                                    //sätt rubrik
                                    string question = dataTable.TableName.Trim(new char[] { '!' });
                                    string textToAppend = (isAndelar ? "Andelar" : "Nöjdhet");
                                    Functions.SetTableCellText(tableToUse.Elements<D.TableRow>().FirstOrDefault().Elements<D.TableCell>().FirstOrDefault(), question + " " + textToAppend);

                                    //sätt kolumnnamn
                                    D.TableRow questionRow = tableToUse.Elements<D.TableRow>().ElementAt(1);
                                    var cells = questionRow.Elements<D.TableCell>();
                                    D.TableCell clonedCell = (D.TableCell)cells.ElementAt(1).CloneNode(true);
                                    while (questionRow.Elements<D.TableCell>().Count() > 1)
                                    {
                                        questionRow.RemoveChild(questionRow.Elements<D.TableCell>().Last());
                                    }
                                    var extListQ = questionRow.GetFirstChild<D.ExtensionList>();

                                    for (int colIdx = 2; colIdx < dataTable.Columns.Count; colIdx++)
                                    {
                                        DataColumn col = dataTable.Columns[colIdx];
                                        D.TableCell newCell = (D.TableCell)clonedCell.CloneNode(true);

                                        Functions.SetTableCellText(newCell, col.ColumnName);

                                        questionRow.InsertBefore(newCell, extListQ);
                                    }

                                    var clonedRow = tableToUse.Elements<D.TableRow>().ElementAtOrDefault(2).CloneNode(true); //data-rad
                                    while (tableToUse.Elements<D.TableRow>().Count() > 2)
                                    {
                                        tableToUse.RemoveChild(tableToUse.Elements<D.TableRow>().Last());
                                    }

                                    foreach (DataRow row in dataTable.Rows)
                                    {
                                        var gridColumns = tableToUse.TableGrid.Elements<D.GridColumn>();

                                        var dataValueCells = row.ItemArray;
                                        D.TableRow tableRow = (D.TableRow)clonedRow.CloneNode(true);
                                        var extList = tableRow.GetFirstChild<D.ExtensionList>();
                                        var dataCellToUse = tableRow.ElementAt(1).CloneNode(true);

                                        var tableCells = tableRow.Elements<D.TableCell>();
                                        Functions.SetTableCellText(tableCells.First(), dataValueCells[0].ToString());

                                        while (tableRow.Elements<D.TableCell>().Count() > 1)
                                        {
                                            tableRow.RemoveChild(tableRow.Elements<D.TableCell>().Last());
                                        }
                                        int cellCount = 1;
                                        foreach (var dataCell in dataValueCells.Skip(2))
                                        {
                                            D.TableCell newDataCell = (D.TableCell)dataCellToUse.CloneNode(true);

                                            string format = isAndelar ? "p0" : "n1";
                                            string valueToParse = string.IsNullOrEmpty(dataCell.ToString()) ? "0.0" : dataCell.ToString();
                                            string formattedValue = double.Parse(valueToParse, CultureInfo.InvariantCulture).ToString(format);
                                            Functions.SetTableCellText(newDataCell, formattedValue);
                                            tableRow.InsertBefore(newDataCell, extList);
                                            cellCount++;
                                        }

                                        tableToUse.AppendChild(tableRow);

                                        if (resizeGrid)
                                        {
                                            var tableGrid = tableToUse.GetFirstChild<D.TableGrid>();
                                            if (tableToUse.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() < cellCount)
                                            {
                                                while (tableToUse.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() < cellCount)
                                                {
                                                    D.GridColumn newCol = (D.GridColumn)tableToUse.GetFirstChild<D.TableGrid>().GetFirstChild<D.GridColumn>().CloneNode(true);
                                                    tableGrid.AppendChild(newCol);

                                                    D.TableCell newHeaderCell = (D.TableCell)tableToUse.GetFirstChild<D.TableRow>().ElementAt(1).CloneNode(true);
                                                    var headerExtList = tableToUse.GetFirstChild<D.TableRow>().GetFirstChild<D.ExtensionList>();
                                                    tableToUse.GetFirstChild<D.TableRow>().InsertBefore(newHeaderCell, headerExtList);
                                                }
                                            }
                                            else if (tableToUse.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() > cellCount)
                                            {
                                                while (tableToUse.GetFirstChild<D.TableGrid>().Elements<D.GridColumn>().Count() > cellCount)
                                                {
                                                    tableGrid.RemoveChild(tableGrid.Elements<D.GridColumn>().Last());

                                                }
                                            }
                                            tableToUse.GetFirstChild<D.TableRow>().GetFirstChild<D.TableCell>().GridSpan.Value = cellCount;
                                            resizeGrid = false;
                                        }
                                    }
                                    firstTable = false;
                                }
                            }
                        }
                        segment = false;
                        branschSpecifika = false;
                    }
                    #endregion

                    #region Marknadsandelar

                    

                    if (slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>().FirstOrDefault()?.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name == "Plassholder for innhold 3")
                    {
                        List<Tuple<string, double>> marknadsandelar = new List<Tuple<string, double>>();
                        var allLines = File.ReadAllLines(SelectedPath + "\\Input\\config.txt", Encoding.Default);
                        foreach (var line in allLines)
                        {
                            var values = line.Split(new string[] { "\t" }, StringSplitOptions.RemoveEmptyEntries);
                            double andelar = 0.0;

                            double.TryParse(values.Last(), out andelar);

                            marknadsandelar.Add(new Tuple<string, double>(values.ElementAt(1), andelar));
                        }

                        var maGF = slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>().FirstOrDefault();
                        var maTable = maGF.GetFirstChild<D.Graphic>()?.GraphicData.GetFirstChild<D.Table>();

                        var tableRows = maTable.Elements<D.TableRow>();
                        var cloneSource = tableRows.ElementAt(2).CloneNode(true);

                        while (tableRows.Count() > 2)
                        {
                            maTable.RemoveChild(tableRows.Last());
                        }

                        foreach (var marknadsandel in marknadsandelar)
                        {
                            var newRow = (D.TableRow)cloneSource.CloneNode(true);
                            var newCells = newRow.Elements<D.TableCell>();
                            Functions.SetTableCellText(newCells.First(), marknadsandel.Item1);
                            Functions.SetTableCellText(newCells.Last(), marknadsandel.Item2.ToString("p0"));

                            maTable.AppendChild(newRow);
                        }
                    }
                    #endregion

                    string slideDescription = slide.CommonSlideData.ShapeTree.Elements<Shape>().FirstOrDefault().NonVisualShapeProperties.NonVisualDrawingProperties.Description;
                    if (slideDescription != null)
                    {
                        slideDescription = slideDescription.Trim();
                    }
                    //Är vi på en slide som har följande text så vet vi att efter det ska resultataspekterna listas...
                    //if (slide.InnerText.Contains("Resultataspekter- "))
                    if (slideDescription == "Resultataspekter")
                    {
                        if (slide.InnerText.Contains("{Resultatvariabler}"))
                        {
                            bool firstName = true;
                            List<string> dNames = new List<string>();
                            foreach (DataGridViewRow dgvRow in dgvResultatVariabler.Rows) //varje rad är en aspekt
                            {
                                string displayName = string.Empty;
                                if (dgvRow.Cells["dgvResultatVariablerText"].Value == null)
                                {
                                    if (firstName)
                                    {
                                        displayName = "Kundnöjdhet";
                                    }
                                }
                                else
                                {
                                    displayName = dgvRow.Cells["dgvResultatVariablerText"].Value.ToString();
                                }
                                if (displayName != string.Empty)
                                {
                                    dNames.Add(displayName);
                                }
                                firstName = false;
                            }

                            string resultString = string.Empty;
                            for (int i = 0; i < dNames.Count; i++)
                            {
                                if (i == 0)
                                {
                                    resultString += dNames[i];
                                }
                                else if (i == dNames.Count - 1)
                                {
                                    resultString += " och " + dNames[i];
                                }
                                else
                                {
                                    resultString += ", " + dNames[i];
                                }
                            }

                            shapes = slide.CommonSlideData.ShapeTree.Descendants<P.Shape>();
                            foreach (var shape in shapes)
                            {
                                SetText(shape, resultString, "{Resultatvariabler}");
                            }

                        }
                        resultatAspekter = true;
                    }

                    //if (slide.InnerText.Contains("Drivande aspekter") && !slide.InnerText.Contains("Betyg på frågenivå") && loopIndex > 8)
                    if (slideDescription == "Drivande aspekter")                        
                    {
                        kunddimensioner = true;
                    }

                    //if (slide.InnerText.Contains("Gapanalys") && loopIndex > 8 && !slide.InnerText.Contains("Branschjusterad utveckling") && !slide.InnerText.ToLower().Contains("betyg och differens per frågenivå"))
                    if (slideDescription == "Gapanalys")
                    {
                        resultatAspekter = false;
                        kunddimensioner = false;
                        gapanalys = true;
                    }

                    //if (slide.InnerText.Contains("Drivkraftsanalys") && loopIndex > 8 && !slide.InnerText.Contains("spindeldiag"))
                    if (slideDescription == "Drivkraftsanalys")
                    {
                        drivkraft = true;
                    }

                    //if (slide.InnerText.Contains("Klagomål") && loopIndex > 8 && !slide.InnerText.ToLower().Contains("anledning att klaga"))
                    if (slideDescription == "Klagomål")
                    {
                        complaints = true;
                    }

                    //if (slide.InnerText.Contains("Branschspecifika frågor") && loopIndex > 8 && !slide.InnerText.ToLower().Contains("fokusfrågor") && !enableMoreBranschSpecifika)
                    if(slideDescription == "Branschspecifika frågor")
                    {
                        branschSpecifika = true;
                    }

                    //if (slide.InnerText.Contains("SEGMENT") && loopIndex > 8)
                    if(slideDescription == "SEGMENT")
                    {
                        segment = true;
                    }
                }

                var slidesToRemove = new List<int>();

                var itemsAgain = presPart.Presentation.SlideIdList;

                int slideIndex = 0;
                foreach (SlideId item in itemsAgain)
                {
                    if (item == null)
                        break;
                    var part = presPart.GetPartById(item.RelationshipId);
                    var slide = (part as SlidePart).Slide;

                    
                    slideIndex++;
                }

                foreach (var slide in slidesToRemove.OrderByDescending(x => x))
                {
                    SlideId slideId = (SlideId)items.ElementAt(slide);

                    DeleteSlide(newDeck, slideId);
                }
            }
            this.Focus();
            MessageBox.Show("Klar!");

        }

        public static void DeleteSlide(PresentationDocument presentationDocument, SlideId slideId)
        {
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }
            // Get the presentation from the presentation part.
            Presentation presentation = presentationDocument.PresentationPart.Presentation;
            // Remove the slide from the slide list.
            presentation.SlideIdList.RemoveChild(slideId);

            // Get the relationship ID of the slide.
            string slideRelId = slideId.RelationshipId;
            //// Remove references to the slide from all custom shows.
            if (presentation.CustomShowList != null)
            {
                // Iterate through the list of custom shows.
                foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
                {
                    if (customShow.SlideList != null)
                    {
                        SlideListEntry entry = customShow.SlideList.ChildElements.Where(s => ((SlideListEntry)s).Id == slideRelId).FirstOrDefault() as SlideListEntry;
                        if (entry != null)
                            customShow.SlideList.RemoveChild(entry);
                    }
                }
            }
            // Save the modified presentation.
            presentation.Save();

            // Remove the slide part.
            presentationDocument.PresentationPart.DeletePart(slideRelId);
        }

        private void CreateNewSlides(List<int> aspectSlideIndices, PresentationPart presPart, SlideId currentSlideId, Slide currentSlide, SlideIdList items)
        {
            SlideId insertionPoint = currentSlideId;
            foreach (var index in aspectSlideIndices)
            {
                var item = (SlideId)items.ElementAtOrDefault(index);
                if (item == null)
                    break;
                var part = (SlidePart)presPart.GetPartById(item.RelationshipId);

                SlidePart newSlidePart = presPart.AddNewPart<SlidePart>();
                newSlidePart.FeedData(part.GetStream(FileMode.Open));
                newSlidePart.AddPart(part.SlideLayoutPart);

                foreach (ImagePart ipart in part.ImageParts)
                {
                    ImagePart newipart = newSlidePart.AddImagePart(ipart.ContentType, part.GetIdOfPart(ipart));
                    newipart.FeedData(ipart.GetStream());
                }

                foreach (ChartPart cpart in part.ChartParts)
                {
                    if (cpart.EmbeddedPackagePart != null)
                    {
                        ChartPart newcpart = newSlidePart.AddNewPart<ChartPart>(part.GetIdOfPart(cpart));
                        newcpart.FeedData(cpart.GetStream());

                        newcpart.ChartSpace.Save();

                        EmbeddedPackagePart newepart = newcpart.AddEmbeddedPackagePart(cpart.EmbeddedPackagePart.ContentType);
                        newcpart.ChangeIdOfPart(newepart, cpart.GetIdOfPart(cpart.EmbeddedPackagePart));
                        newepart.FeedData(cpart.EmbeddedPackagePart.GetStream());
                    }
                }

                SlideIdList slideIdList = presPart.Presentation.SlideIdList;
                uint maxSlideId = 1;
                foreach (SlideId slideId in slideIdList.ChildElements)
                {
                    if (slideId.Id > maxSlideId)
                    {
                        maxSlideId = slideId.Id;
                    }
                }
                maxSlideId++;
                SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), insertionPoint);
                newSlideId.Id = maxSlideId;
                newSlideId.RelationshipId = presPart.GetIdOfPart(newSlidePart);
                insertionPoint = newSlideId;
            }
        }

        private static void SetText(P.Shape shape, string text, string placeholderText)
        {
            var paragraphs = shape.TextBody.ChildElements.OfType<D.Paragraph>();
            foreach (var paragraph in paragraphs)
            {
                if (paragraph.InnerText.Contains(placeholderText))
                {
                    var runs = paragraph.Elements<D.Run>();
                    foreach (var run in runs)
                    {
                        if (run.InnerText.Contains(placeholderText))
                        {
                            var texts = run.Elements<D.Text>();
                            foreach (var t in texts)
                            {
                                t.Text = t.Text.Replace(placeholderText, text);
                            }
                        }
                    }
                }
            }
        }

        private bool ValidateStructure()
        {
            bool hasData = System.IO.Directory.Exists(SelectedPath + "\\Data");
            bool hasTemplate = System.IO.File.Exists(SelectedPath + "\\Data\\Mall.pptx");
            if (hasTemplate)
            {
                TemplatePath = SelectedPath + "\\Data\\Mall.pptx";
            }
            bool hasInput = System.IO.Directory.Exists(SelectedPath + "\\Input");
            bool hasOutput = System.IO.Directory.Exists(SelectedPath + "\\Output");

            return hasData && hasInput && hasOutput && hasTemplate;
        }
    }


}
