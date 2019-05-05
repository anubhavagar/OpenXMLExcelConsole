using System;
using System.IO;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Web.Hosting;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using Xdr = DocumentFormat.OpenXml.Drawing;
using xVal = DocumentFormat.OpenXml.Validation;
using GanttChartAutomation;
using DocumentFormat.OpenXml.Spreadsheet;
//using DocumentFormat.OpenXml.InkML;

namespace Boeing.ReMICS.Presentation
{

    public class PresentationHelper
    {
        #region Declaration

        private uint incrementId = 1;
        private int incrementIt = 10000;
        private int _intPixel = 0;
        private Int64 _intcellHeight = 0;
        private DocumentFormat.OpenXml.Int64Value _intCellWidth;
        private DateTime _dtStatsDate;
        private int i = 0;
        private IList<Xdr.TableRow> rows = null;
        private Int64 _intHeigt = 0;
        private Shape _shptemptriangle = null;
        private Shape _shpNofillUpward = new Shape();
        private Shape _shpFillDownWard = new Shape();
        private Shape _shpFillDiamond = new Shape();
        private Shape _shpNofillDiamond = new Shape();
        private Int64 _intSymbolWidth = 0;
        private Shape _shpPageDetails = new Shape();

        private Shape _shpTrendBetter = new Shape();
        private Shape _shptrendPosition = new Shape();
        private ConnectionShape _cshpConnector = new ConnectionShape();
        private DateTime _dtstartDate;
        private DateTime _dtendDate;
        private Shape _cshpDateRectangle = new Shape();
        private ConnectionShape _cshpDateLine = new ConnectionShape();
        private Shape _shpTrendLowerd = new Shape();
        private Shape _shpTrendSame = new Shape();
        private Xdr.Table tableDate = new DocumentFormat.OpenXml.Drawing.Table();
        private Shape _shpContactDetails = new Shape();



        private Xdr.Table table = null; 


        private Shape temp = new Shape();
        private ConnectionShape _cshpDotLine = null;
        private Shape _cshpNonCriticl = new Shape();
        private Shape _cshpNonCriticlProgress = new Shape();
        private Shape _cshpPartialProgress = new Shape();
        private ConnectionShape _cshpLine = null;

        private ConnectionShape _cshpStatsLine = null;



        private Shape _shpHeading = null;
        private Shape _shpTitle = null;
        private Shape _shpHeadingIntegrated = null;
        private DocumentFormat.OpenXml.Presentation.Picture _shpStatsDetail = null;
        private int _intTaskhaving = -1;
        private Shape _shpFooterDetails = new Shape();

        #endregion



        private static byte[] pptTemplate;
        private static byte[] PPTTemplate
        {
            get
            {

                string folderPath = Environment.CurrentDirectory + @"\Template\";

                if (!Directory.Exists(folderPath))
                {
                    folderPath = System.Web.Hosting.HostingEnvironment.MapPath("\\Bin\\Template\\");
                }

                string templatePath = folderPath + @"\Presentationtemplate.pptx";

                if (pptTemplate == null)
                {
                    byte[] fileBytes = System.IO.File.ReadAllBytes(templatePath);
                    pptTemplate = fileBytes;
                }

                return pptTemplate;
            }
        }
        

        public string  BuildPresentation(string parameterList, string UserId)
        {
            GanttChartAutomation.GanttChartDataHelper gantt = new GanttChartAutomation.GanttChartDataHelper();
            GanttChartAutomation.GanttReportObject ganttReportData = gantt.GenerateGanttChartReport(parameterList, UserId);

            DateTime startDate, endDate;
            startDate = ganttReportData.StatusDate.AddMonths(-2);
            endDate = ganttReportData.StatusDate.AddMonths(3);
            ganttReportData.SetExplicitStartEndDate(startDate, endDate);
            _dtStatsDate = ganttReportData.StatusDate;

            IList<GanttChartAutomation.GanttReportObject> ganttChartPages = ganttReportData.PaginateDataWithWeightage(15);

            GanttChartAutomation.GanttReportObject testObject = ganttChartPages[0];


            byte[] fileBytes = PPTTemplate;
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(fileBytes, 0, fileBytes.Length);

                using (PresentationDocument presentationDocument = PresentationDocument.Open(stream, true))
                {
                    PresentationPart presentationPart = presentationDocument.PresentationPart;
                    SlidePart templateSlidePart = presentationPart.SlideParts.LastOrDefault();
                    SlidePart slidePart = presentationPart.SlideParts.FirstOrDefault();
                   
                   
                    SlidePart newSlide;
                    _dtstartDate = ganttChartPages[0].StartDate;
                    _dtendDate = ganttChartPages[0].EndDate;

                    for (int _intPageCount = 0; _intPageCount < ganttChartPages.Count; _intPageCount++)
                    {

                        bool lblnSlide = false;

                        GanttChartAutomation.GanttReportObject objganttReport = ganttChartPages[_intPageCount];

                        temp = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("PositionBox", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
                        _shpNofillDiamond = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("NoFillDiamond", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
                        _intCellWidth = temp.ShapeProperties.Transform2D.Extents.Cx;
                        getheight(temp.ShapeProperties.Transform2D.Extents.Cy, 15);
                        _intSymbolWidth = _shpNofillDiamond.ShapeProperties.Transform2D.Extents.Cx;
                        newSlide = CloneSlidePart(presentationPart, templateSlidePart);
                        InitializePosition(newSlide);


                        for (int i = 1; i <= objganttReport.RowCollection.Count; i++)
                        {
                            IList<Xdr.TableCell> cells = rows[i].ChildElements.OfType<Xdr.TableCell>().ToList();
                            _intTaskhaving += objganttReport.RowCollection[i - 1].Weightage;
                            SetCellText(cells[2], objganttReport.RowCollection[i - 1].TaskName.ToString(), objganttReport.RowCollection[i - 1].Weightage, rows[i]);
                            InsertTrend(objganttReport.RowCollection[i - 1].Trend, newSlide, i - 1, objganttReport.RowCollection[i - 1].Weightage);
                            SetCellText(cells[1], objganttReport.RowCollection[i - 1].CpValue == null ? "" : objganttReport.RowCollection[i - 1].CpValue.ToString(), objganttReport.RowCollection[i - 1].Weightage, rows[i]);

                            SetCellText(cells[3], objganttReport.RowCollection[i - 1].RaaValue == null ? "" : objganttReport.RowCollection[i - 1].RaaValue.ToString(), objganttReport.RowCollection[i - 1].Weightage, rows[i]);
                            SetCellText(cells[4], objganttReport.RowCollection[i - 1].BacValue == null ? "" : objganttReport.RowCollection[i - 1].BacValue.ToString(), objganttReport.RowCollection[i - 1].Weightage, rows[i]);
                            SetCellText(cells[5], objganttReport.RowCollection[i - 1].SvValue == null ? "" : objganttReport.RowCollection[i - 1].SvValue.ToString(), objganttReport.RowCollection[i - 1].Weightage, rows[i]);
                            SetShapeCe(cells[6], "", newSlide, objganttReport.RowCollection[i - 1].rowImageDetails, i - 1, objganttReport.RowCollection[i - 1].Weightage);
                            InsertPageDetasils(newSlide, objganttReport);
                            //InsertstatusLine(newSlide);
                        }
                        if (objganttReport.RowCollection.Count < 15)
                        {

                            //_cshpStatsLine.ShapeProperties.Transform2D.Extents.Cx = (_intTaskhaving + 1) * _intcellHeight;
                            //_cshpStatsLine.ShapeProperties.Transform2D.Offset.Y = 3848497;

                            for (int i = objganttReport.RowCollection.Count + 1; i <= 15; i++)
                            {
                                rows[i ].Remove();
                            }

                           
                        }
                        _intTaskhaving = -1;

                        RemoveShapePosition(newSlide, presentationPart);

                    }

                    DeleteTemplateSlide(presentationPart, templateSlidePart);


                    xVal.OpenXmlValidator validator = new xVal.OpenXmlValidator();
                    var errors = validator.Validate(presentationDocument).ToList();



                    if (errors.Count == 0)
                    {
                        presentationPart.Presentation.Save();
                    }
                    else
                    {
                        throw new ApplicationException("Error in document.");
                    }
                }

                byte[] presentationFile = stream.ToArray();
                string outputFolder = @"C:\\ReMICS\\OpenXMLPpt\\";
                if (!System.IO.Directory.Exists(outputFolder))
                {
                    System.IO.Directory.CreateDirectory(outputFolder);
                }

                string newPresentation = outputFolder + "OpenXML-PPT-" + Guid.NewGuid().ToString() + ".pptx";
                File.WriteAllBytes(newPresentation, presentationFile);

                return newPresentation;
            }
        }

        private void SetShapeCe(Xdr.TableCell cell, string strCellText, SlidePart templateSlidePart, GanttChartAutomation.RowImageDetails rowImageDetails, int pintrowNomb , int weightage)
        {
            //Declarations
            // Shape temp = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("squareBox", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();

            for (int _intConnectorType = 0; _intConnectorType < rowImageDetails.Connectors.Count; _intConnectorType++)
            {
                InsertConnector(templateSlidePart, rowImageDetails.Connectors[_intConnectorType], pintrowNomb, weightage);
            }

            for (int _intSymbolType = 0; _intSymbolType < rowImageDetails.Symbols.Count; _intSymbolType++)
            {
                InsertTriangle(templateSlidePart, rowImageDetails.Symbols[_intSymbolType].Type, pintrowNomb, rowImageDetails.Symbols[_intSymbolType].PositionDate, weightage);
            }
        }

        private void SetCellText(Xdr.TableCell cell, string strText, int weightage , Xdr.TableRow i)
        {
            // Add text dynamically.
            if (strText.Length > 150)
            {
                strText = strText.Substring(0, 149);
                i.Height = i.Height * weightage - 5000;
                _intTaskhaving++;
            }
            Xdr.Text text = cell.Descendants<Xdr.Text>().FirstOrDefault();
            if (text == default(Xdr.Text)) return;
            text.Text = strText;
            Xdr.Text newText = new Xdr.Text();
            newText.Text = strText;
            text = newText;
        }

        private void InsertTrend(SymbolType pstrTrendName, SlidePart pslidePart, int pintrowNomb , int weightage)
        {
            Shape newTrendShp = new Shape();

            if (pstrTrendName.Equals(SymbolType.TrendBetter))
            {
                newTrendShp = (Shape)_shpTrendBetter.Clone();
                newTrendShp.NonVisualShapeProperties.NonVisualDrawingProperties.Id = 500U + incrementId;
                pslidePart.Slide.CommonSlideData.ShapeTree.AppendChild(newTrendShp);
                newTrendShp.ShapeProperties.Transform2D.Offset.Y = _shptrendPosition.ShapeProperties.Transform2D.Offset.Y + _intcellHeight * (_intTaskhaving);
                newTrendShp.ShapeProperties.Transform2D.Offset.X = _shptrendPosition.ShapeProperties.Transform2D.Offset.X;
            }
            else if (pstrTrendName.Equals(SymbolType.TrendFlat))
            {
                newTrendShp = (Shape)_shpTrendSame.Clone();
                newTrendShp.NonVisualShapeProperties.NonVisualDrawingProperties.Id = 500U + incrementId;
                pslidePart.Slide.CommonSlideData.ShapeTree.AppendChild(newTrendShp);
                newTrendShp.ShapeProperties.Transform2D.Offset.Y = _shptrendPosition.ShapeProperties.Transform2D.Offset.Y + _intcellHeight * (_intTaskhaving);
                newTrendShp.ShapeProperties.Transform2D.Offset.X = _shptrendPosition.ShapeProperties.Transform2D.Offset.X;
            }
            else
            {
                newTrendShp = (Shape)_shpTrendLowerd.Clone();
                newTrendShp.NonVisualShapeProperties.NonVisualDrawingProperties.Id = 500U + incrementId;
                pslidePart.Slide.CommonSlideData.ShapeTree.AppendChild(newTrendShp);
                newTrendShp.ShapeProperties.Transform2D.Offset.Y = _shptrendPosition.ShapeProperties.Transform2D.Offset.Y + _intcellHeight * (_intTaskhaving);
                newTrendShp.ShapeProperties.Transform2D.Offset.X = _shptrendPosition.ShapeProperties.Transform2D.Offset.X;
            }

            incrementId++;
        }


        private void InsertTriangle(SlidePart ptempSlide, SymbolType pstrsymbolType, int pintrowNomb, DateTime pdtsymbolPosition , int weightage)
        {
            Shape clonedSymbol = new Shape();

            if (pstrsymbolType.Equals(SymbolType.UnfilledUpward))
            {
                clonedSymbol = (Shape)_shpNofillUpward.Clone();
                clonedSymbol.NonVisualShapeProperties.NonVisualDrawingProperties.Id = 1000U + incrementId;
                clonedSymbol.ShapeProperties.Transform2D.Offset.X = temp.ShapeProperties.Transform2D.Offset.X + getCoordinateForDate(pdtsymbolPosition, _dtstartDate, _dtendDate, _intCellWidth) - _intSymbolWidth / 2;
                clonedSymbol.ShapeProperties.Transform2D.Offset.Y = _shptrendPosition.ShapeProperties.Transform2D.Offset.Y + _intcellHeight * (_intTaskhaving); 
                ptempSlide.Slide.CommonSlideData.ShapeTree.AppendChild(clonedSymbol);
            }
            else if (pstrsymbolType.Equals(SymbolType.FilledDiamond))
            {
                clonedSymbol = (Shape)_shpFillDiamond.Clone();
                clonedSymbol.NonVisualShapeProperties.NonVisualDrawingProperties.Id = 1000U + incrementId;
                clonedSymbol.ShapeProperties.Transform2D.Offset.X = temp.ShapeProperties.Transform2D.Offset.X + getCoordinateForDate(pdtsymbolPosition, _dtstartDate, _dtendDate, _intCellWidth) - _intSymbolWidth / 2;
                clonedSymbol.ShapeProperties.Transform2D.Offset.Y = _shptrendPosition.ShapeProperties.Transform2D.Offset.Y + _intcellHeight * (_intTaskhaving);
                ptempSlide.Slide.CommonSlideData.ShapeTree.AppendChild(clonedSymbol);
            }
            else if (pstrsymbolType.Equals(SymbolType.DarkUpward))
            {
                clonedSymbol = (Shape)_shpFillDownWard.Clone();
                clonedSymbol.NonVisualShapeProperties.NonVisualDrawingProperties.Id = 1000U + incrementId;
                clonedSymbol.ShapeProperties.Transform2D.Offset.X = temp.ShapeProperties.Transform2D.Offset.X + getCoordinateForDate(pdtsymbolPosition, _dtstartDate, _dtendDate, _intCellWidth) - _intSymbolWidth / 2;
                clonedSymbol.ShapeProperties.Transform2D.Offset.Y = _shptrendPosition.ShapeProperties.Transform2D.Offset.Y + _intcellHeight * (_intTaskhaving);
                ptempSlide.Slide.CommonSlideData.ShapeTree.AppendChild(clonedSymbol);
            }
            else
            {
                clonedSymbol = (Shape)_shpNofillDiamond.Clone();
                clonedSymbol.NonVisualShapeProperties.NonVisualDrawingProperties.Id = 1000U + incrementId;
                clonedSymbol.ShapeProperties.Transform2D.Offset.X = temp.ShapeProperties.Transform2D.Offset.X + getCoordinateForDate(pdtsymbolPosition, _dtstartDate, _dtendDate, _intCellWidth) - _intSymbolWidth / 2;
                clonedSymbol.ShapeProperties.Transform2D.Offset.Y = _shptrendPosition.ShapeProperties.Transform2D.Offset.Y + _intcellHeight * (_intTaskhaving);
                ptempSlide.Slide.CommonSlideData.ShapeTree.AppendChild(clonedSymbol);
            }
            incrementId++;
        }

        void DeleteTemplateSlide(PresentationPart presentationPart, SlidePart slideTemplate)
        {
            //Get the list of slide ids
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;
            //Delete the template slide reference
            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.RelationshipId.Value.Equals("rId2"))
                    slideIdList.RemoveChild(slideId);
            }
            //Delete the template slide
            presentationPart.DeletePart(slideTemplate);
        }

        private void InsertConnector(SlidePart ptempSlide, ConnectorInfo pConnector, int pintrowNomb , int weightage)
        {
            Shape newShape = new Shape();
            ConnectionShape newConnectionShape = new ConnectionShape();
            Int64 lintstartPosition = 0;
            Int64 lintendPosition = 0;
            ConnectorType connectorType = pConnector.Type;

            lintstartPosition = getCoordinateForDate(pConnector.StartDate, _dtstartDate, _dtendDate, _intCellWidth);
            lintendPosition = getCoordinateForDate(pConnector.EndDate, _dtstartDate, _dtendDate, _intCellWidth);
            lintendPosition = lintendPosition - lintstartPosition;

            if (connectorType.Equals(ConnectorType.DotLineBlack))
            {
                newConnectionShape = (ConnectionShape)_cshpDotLine.Clone();
                newConnectionShape.NonVisualConnectionShapeProperties.NonVisualDrawingProperties.Id = 600U + incrementId;
                newConnectionShape.ShapeProperties.Transform2D.Extents.Cx = lintendPosition;
                newConnectionShape.ShapeProperties.Transform2D.Offset.X = temp.ShapeProperties.Transform2D.Offset.X + lintstartPosition;
                newConnectionShape.ShapeProperties.Transform2D.Offset.Y = _cshpConnector.ShapeProperties.Transform2D.Offset.Y + _intcellHeight * (_intTaskhaving);
                ptempSlide.Slide.CommonSlideData.ShapeTree.AppendChild(newConnectionShape);

            }
            else if (connectorType.Equals(ConnectorType.PartialProgress))
            {
                newShape = (Shape)_cshpPartialProgress.Clone();
                newShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id = 600U + incrementId;
                newShape.ShapeProperties.Transform2D.Extents.Cx = lintendPosition;
                newShape.ShapeProperties.Transform2D.Offset.X = temp.ShapeProperties.Transform2D.Offset.X + lintstartPosition;
                newShape.ShapeProperties.Transform2D.Offset.Y = temp.ShapeProperties.Transform2D.Offset.Y + _intcellHeight * (_intTaskhaving) -_intHeigt;
                ptempSlide.Slide.CommonSlideData.ShapeTree.AppendChild(newShape);
            }
            else if (connectorType.Equals(ConnectorType.NonCritical))
            {
                newShape = (Shape)_cshpNonCriticl.Clone();
                newShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id = 600U + incrementId;
                newShape.ShapeProperties.Transform2D.Extents.Cx = lintendPosition;
                newShape.ShapeProperties.Transform2D.Offset.X = temp.ShapeProperties.Transform2D.Offset.X + lintstartPosition;
                newShape.ShapeProperties.Transform2D.Offset.Y = _cshpConnector.ShapeProperties.Transform2D.Offset.Y + _intcellHeight * (_intTaskhaving) -_intHeigt;
                ptempSlide.Slide.CommonSlideData.ShapeTree.AppendChild(newShape);
            }
            else
            {
                newShape = (Shape)_cshpNonCriticlProgress.Clone();
                newShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id = 600U + incrementId;
                newShape.ShapeProperties.Transform2D.Extents.Cx = lintendPosition;
                newShape.ShapeProperties.Transform2D.Offset.X = temp.ShapeProperties.Transform2D.Offset.X + lintstartPosition;
                newShape.ShapeProperties.Transform2D.Offset.Y = _cshpConnector.ShapeProperties.Transform2D.Offset.Y + _intcellHeight * (_intTaskhaving) -_intHeigt;
                ptempSlide.Slide.CommonSlideData.ShapeTree.AppendChild(newShape);
            }
        }

        private void InsertstatusLine(SlidePart ptempSlide)
        {
            _cshpStatsLine.ShapeProperties.Transform2D.Offset.X = temp.ShapeProperties.Transform2D.Offset.X + getCoordinateForDate(_dtStatsDate, _dtstartDate, _dtendDate, _intCellWidth);
        }

        public Int64 getCoordinateForDate(DateTime dateValue, DateTime pdtstartDate, DateTime pdtendDate, DocumentFormat.OpenXml.Int64Value pintcellWidth)
        {
            TimeSpan firstDiff = dateValue - pdtstartDate;
            TimeSpan secondDiff = pdtendDate - pdtstartDate;

            return (firstDiff.Days * pintcellWidth / secondDiff.Days);
        }

        private void getheight(Int64 pintcellHeight, int pintrowCount)
        {
            _intcellHeight = pintcellHeight / pintrowCount;
        }

        private void Datespart(Xdr.Table tabletempory)
        {
            int lintgridColumn = 0;

            IList<Xdr.GridColumn> hhgf = tabletempory.TableGrid.Descendants<Xdr.GridColumn>().ToList();
            IList<Xdr.TableRow> tableDatedisplay = tabletempory.Descendants<Xdr.TableRow>().ToList();
            IList<Xdr.TableCell> tableDateCell = tabletempory.Descendants<Xdr.TableCell>().ToList();

            DateTime monthEnd;
            Int64 boxWidth = 0;
            Int64 oldBoxwidth = 0;
            Xdr.Text dateText = new Xdr.Text();
            for (DateTime tempDate = this._dtstartDate; tempDate < this._dtendDate; )
            {
                monthEnd = GetMonthEndDate(tempDate);

                monthEnd = (monthEnd < this._dtendDate) ? monthEnd : this._dtendDate;

                TimeSpan diff = monthEnd - tempDate;

                boxWidth = getCoordinateForDate(monthEnd, this._dtstartDate, this._dtendDate, _intCellWidth);

                if (diff.Days > 20)
                {
                    dateText = tableDateCell[lintgridColumn].Descendants<Xdr.Text>().FirstOrDefault();
                    dateText.Text = tempDate.ToString("MMM") + " " + tempDate.ToString("yy");
                }

                hhgf[lintgridColumn].Width = boxWidth - oldBoxwidth;
                lintgridColumn++;

                tempDate = monthEnd.AddDays(1);
                oldBoxwidth = boxWidth;
            }
        }

        public static DateTime GetMonthEndDate(DateTime endDate)
        {
            return new DateTime(endDate.Year, endDate.Month, DateTime.DaysInMonth(endDate.Year, endDate.Month));
        }

        private SlidePart CloneSlidePart(PresentationPart presentationPart, SlidePart slideTemplate)
        {
            SlidePart cloneSlidei = presentationPart.AddNewPart<SlidePart>("newSlide" + i);
            i++;
            cloneSlidei.FeedData(slideTemplate.GetStream(FileMode.Open));
            cloneSlidei.AddPart(slideTemplate.SlideLayoutPart);

            SlideIdList slideidlist = presentationPart.Presentation.SlideIdList;
            uint maxide = 0;
            SlideId beforeSlide = null;
            foreach (SlideId slideidw in slideidlist.ChildElements)
            {
                if (slideidw.Id > maxide)
                {
                    beforeSlide = slideidw;
                    maxide = slideidw.Id;
                }
            }
            maxide++;
            SlideId inside = slideidlist.InsertAfter(new SlideId(), beforeSlide);
            inside.Id = maxide;
            inside.RelationshipId = presentationPart.GetIdOfPart(cloneSlidei);
            return cloneSlidei;
        }

        private void InitializePosition(SlidePart templateSlidePart)
        {

            _shpTrendBetter = new Shape();
            _shpTrendBetter = new Shape();
            _shpTrendSame = new Shape();
            _shpNofillUpward = new Shape();
            _shpFillDownWard = new Shape();
            _shpHeading = new Shape();
            _shpTitle = new Shape();
            _shpNofillDiamond = new Shape();
            _shpStatsDetail = new DocumentFormat.OpenXml.Presentation.Picture();
            _shpHeadingIntegrated = new Shape();
            _cshpNonCriticl = new Shape();
            _cshpNonCriticlProgress = new Shape();
            _cshpPartialProgress = new Shape();
            _cshpDotLine = new ConnectionShape();
            _shptrendPosition = new Shape();
            _cshpConnector = new ConnectionShape();
            _shptrendPosition = new Shape();
            _shpFillDiamond = new Shape();
            _cshpLine = new ConnectionShape();
            _cshpStatsLine = new ConnectionShape();
            _shpPageDetails = new Shape();
            _shpFooterDetails = new Shape();

            Xdr.Table table = templateSlidePart.Slide.CommonSlideData.ShapeTree.
                        ChildElements.OfType<GraphicFrame>().FirstOrDefault().Graphic.GraphicData.
                        ChildElements.OfType<Xdr.Table>().FirstOrDefault();

            rows = table.ChildElements.OfType<Xdr.TableRow>().ToList();
            Xdr.TableRow table1 = new Xdr.TableRow();
          
            
            IList<Xdr.TableRow> rows1 = table.ChildElements.OfType<Xdr.TableRow>().ToList();
            temp = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("PositionBox", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            
            _shptemptriangle = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("arrowBox", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            tableDate = templateSlidePart.Slide.CommonSlideData.ShapeTree.
                       ChildElements.OfType<GraphicFrame>().LastOrDefault().Graphic.GraphicData.
                       ChildElements.OfType<Xdr.Table>().FirstOrDefault();

            _shpPageDetails = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("PageDetails", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _shpTrendBetter = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("TrendBetter", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _shpTrendLowerd = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("TrendWorsen", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _shpTrendSame = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("TrendFlat", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _shpFillDownWard = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("DarkUpward", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _shpNofillDiamond = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("NoFillDiamond", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _shpHeading = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("Heading", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            
            _shpHeadingIntegrated = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("headingintegrated", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _shpStatsDetail = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<DocumentFormat.OpenXml.Presentation.Picture>().Where(item => (item.NonVisualPictureProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualPictureProperties.NonVisualDrawingProperties.Description.Value.Equals("StatsDetail", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _shpTitle = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("Title", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _shpFillDiamond = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("FillDiamond", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _cshpNonCriticl = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("NonCritical", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _shpFooterDetails = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("FooterDetails", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _cshpNonCriticlProgress = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("NonCriticalProgress", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _shpContactDetails = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("ContactDetails", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _intHeigt = _cshpNonCriticl.ShapeProperties.Transform2D.Extents.Cy;
            _shpNofillUpward = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("NoFillUpward", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _cshpPartialProgress = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("PartialProgress", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _cshpDotLine = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<ConnectionShape>().Where(item => (item.NonVisualConnectionShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualConnectionShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("DotLineBlack", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _cshpLine = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<ConnectionShape>().Where(item => (item.NonVisualConnectionShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualConnectionShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("line", StringComparison.OrdinalIgnoreCase))).FirstOrDefault(); 
            _cshpConnector = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<ConnectionShape>().Where(item => (item.NonVisualConnectionShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualConnectionShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("Connector", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _cshpStatsLine = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<ConnectionShape>().Where(item => (item.NonVisualConnectionShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualConnectionShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("StatsLine", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();
            _shptrendPosition = templateSlidePart.Slide.CommonSlideData.ShapeTree.ChildElements.OfType<Shape>().Where(item => (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description != null) && (item.NonVisualShapeProperties.NonVisualDrawingProperties.Description.Value.Equals("trendPosition", StringComparison.OrdinalIgnoreCase))).FirstOrDefault();

            Datespart(tableDate);

        }

        private void RemoveShapePosition(SlidePart tempSlidepoe, PresentationPart presentatio)
        {
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_shptrendPosition);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_shpTrendBetter);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_shpFillDownWard);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_shpFillDiamond);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_shpNofillUpward);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_shpNofillDiamond);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_cshpNonCriticlProgress);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(temp);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_cshpDotLine);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_cshpPartialProgress);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_cshpNonCriticl);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_shpTrendLowerd);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_cshpConnector);
            tempSlidepoe.Slide.CommonSlideData.ShapeTree.RemoveChild(_shpTrendSame);
        }


        private void InsertPageDetasils(SlidePart tempSlidepart, GanttChartAutomation.GanttReportObject objganttReport)
        {
            Xdr.Paragraph pageDetails = _shpPageDetails.TextBody.Descendants<Xdr.Paragraph>().FirstOrDefault();
            Xdr.Run pagedetails = pageDetails.Descendants<Xdr.Run>().FirstOrDefault();
            Xdr.Paragraph ContactDetails = _shpContactDetails.TextBody.Descendants<Xdr.Paragraph>().FirstOrDefault();
            Xdr.Run contactdetails = ContactDetails.Descendants<Xdr.Run>().FirstOrDefault();

            //Xdr.Paragraph StatDetails = _shpStatsDetail.TextBody.Descendants<Xdr.Paragraph>().FirstOrDefault();
            //Xdr.Run statDetails = StatDetails.Descendants<Xdr.Run>().FirstOrDefault();
            Xdr.Paragraph HeadingIntegrated = _shpHeadingIntegrated.TextBody.Descendants<Xdr.Paragraph>().FirstOrDefault();
            Xdr.Run headingIntegrated = HeadingIntegrated.Descendants<Xdr.Run>().FirstOrDefault();

            Xdr.Paragraph TitleDetails = _shpTitle.TextBody.Descendants<Xdr.Paragraph>().FirstOrDefault();
            Xdr.Run titleDetails = TitleDetails.Descendants<Xdr.Run>().FirstOrDefault();
            Xdr.Paragraph HeadingDetails = _shpHeading.TextBody.Descendants<Xdr.Paragraph>().FirstOrDefault();
            Xdr.Run headingDetails = HeadingDetails.Descendants<Xdr.Run>().FirstOrDefault();


            Xdr.Paragraph FooterDetails = _shpFooterDetails.TextBody.Descendants<Xdr.Paragraph>().FirstOrDefault();
            Xdr.Run footerDetails = FooterDetails.Descendants<Xdr.Run>().FirstOrDefault();
            
            pagedetails.Text.Text = objganttReport.PageDetails;
            contactdetails.Text.Text = objganttReport.RIdetails;
            headingIntegrated.Text.Text = objganttReport.Title2;
            titleDetails.Text.Text = objganttReport.Title1;
            headingDetails.Text.Text = objganttReport.Title3;
            //statDetails.Text.Text = objganttReport.StatusDate.ToString();
            footerDetails.Text.Text = objganttReport.FooterText;
        }
    }
}
