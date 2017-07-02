using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Drawing;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;

namespace DynamicDocx
{
    public class DynamicDocx
    {
        //Create DocX
        public static void CreatDocX(string filepath, out WordprocessingDocument docx)
        {
            CU_DOCX.CreateDoc(filepath, out docx, false);
            //Register default style to the created document
            RegisterStyle(docx);
        }//Create DocX return obj ends

        //SaveDoc
        public static void SaveDoc(WordprocessingDocument wordDoc)
        {
            wordDoc.MainDocumentPart.Document.Save();
            wordDoc.Close();
        }//save docx Ends

        public void OpenDocX(string filepath, out WordprocessingDocument docx)
        {
            docx = WordprocessingDocument.Open(filepath, true);
        }//OpenDocX ends

        //Write text to docx with your style using docx object
        public static void WriteTextWithYourStyle(WordprocessingDocument docx, string text, string StyleName, JustificationValues textAlignment, string fontSize = "24", string fontColor = "black", bool isBold = false, bool isUnderLine = false, bool isItalic = false)
        {
            //Ernest Mushinge: Implicitly register style to docx from the subroutine then use it over when you want
            StyleRunProperties newStyleProperty = CU_DOCX.MakeStyleRunProps("Segoe UI", fontSize, fontColor, isBold, isUnderLine, isItalic);
            CU_DOCX.AddNewPStyle(docx, StyleName, StyleName, newStyleProperty);
            CU_DOCX.WriteParagraph(docx, CU_DOCX.MakeParagraph(StyleName, textAlignment, CU_DOCX.MakeRun(text, "newStyleName")));

        }//WriteTextWithStyle ends

        //Write image to docx with alignment using docx object, this subroutine takes in image bytes from the stream
        public static void WriteImage(WordprocessingDocument docx, string RegisteredStyleName, Byte[] ImageBytes, JustificationValues alignment)
        {
            string picID = CU_DOCX.InsertPic2DocLib(docx, new MemoryStream(ImageBytes), ImagePartType.Gif);
            Run TestMolImage = CU_DOCX.AddImageAsRun(picID);
            CU_DOCX.WriteParagraph(docx, CU_DOCX.MakeParagraph(RegisteredStyleName, alignment, TestMolImage));
        }//Write image to docx with alignment ends

        //serialize the image. Subroutine for images on disk to pass in a MemoryStream: Ernest used this subroutin for testing image on disk
        public static byte[] SerializeImage(string picPath)
        {
            MemoryStream m;
            byte[] imageBytes;
            using (Image image = Image.FromFile(picPath))
            {
                using (m = new MemoryStream())
                {
                    image.Save(m, image.RawFormat);
                    imageBytes = new byte[m.Length];
                    imageBytes = m.ToArray();
                }//end using
            }//end using
            return imageBytes;
        }//SerializeImage

        //Create a line to docx object
        public static void CreateLine(WordprocessingDocument doc)
        {
            TableCell tableCell = CU_DOCX.MakeTableCell(CU_DOCX.MakeParagraph());
            CU_DOCX.ToggleTableCellBorders(tableCell, bottom: true);
            Table tline_Alert = CU_DOCX.MakeTable(BorderValues.None, true, CU_DOCX.MakeTableRow(tableCell));
            CU_DOCX.WriteTable(doc, tline_Alert);
        }//CreatLine ends

        //Create Space/tab in docx object
        public static void CreateSpace(WordprocessingDocument doc)
        {
            CU_DOCX.WriteParagraph(doc, CU_DOCX.MakeParagraph("S00", JustificationValues.Left, CU_DOCX.MakeRun("")));
        }//Create Space ends

        //Ernest Mushinge; This subroutine will put data in the Row and Cells using a docX object
        public static void CreateTableAndInsertDataInRowsAndCellsDocx(WordprocessingDocument doc, string[] DataForRow, string[] DataForCells, string StyleName, string fontSize = "24", string fontColor = "black", String CellColor = "white", bool isBold = false, bool isUnderLine = false, bool isItalic = false)
        {
            //Register the style
            StyleRunProperties newStyleProperty = CU_DOCX.MakeStyleRunProps("Segoe UI", fontSize, fontColor, isBold, isUnderLine, isItalic);
            CU_DOCX.AddNewPStyle(doc, StyleName, StyleName, newStyleProperty);

            //To hold table cells for rows
            TableCell[] tbRow = new TableCell[DataForRow.Length];
            TableCell[] tbCell = new TableCell[DataForCells.Length];
            for (int k = 0; k < DataForRow.Length; k++)
            {
                if (DataForRow.Length == 0) return;
                tbRow[k] = CU_DOCX.MakeTableCell(CU_DOCX.MakeParagraph("S0B", JustificationValues.Left, CU_DOCX.MakeRun(DataForRow[k], StyleName)));
                /////////////////////////////////////////////////////
                //The code below is for shading/color the row of a table
                var tcp = new TableCellProperties(new TableCellWidth()
                {
                    Type = TableWidthUnitValues.Dxa,
                    Width = "2000",
                });
                // Add cell shading.
                var shading = new Shading()
                {
                    Color = "auto",
                    Fill = CellColor,
                    Val = ShadingPatternValues.Clear
                };
                tcp.Append(shading);
                tbRow[k].Append(tcp);
                ////////////////////////////////////////////////
                CU_DOCX.ToggleTableCellBorders(tbRow[k], 0, true, true, true, true);
            }//foreach ends

            for (int v = 0; v < DataForCells.Length; v++)
            {
                if (DataForCells.Length == 0) return;
                tbCell[v] = CU_DOCX.MakeTableCell(CU_DOCX.MakeParagraph(StyleName, JustificationValues.Left, CU_DOCX.MakeRun(DataForCells[v], StyleName)));

                CU_DOCX.ToggleTableCellBorders(tbCell[v], 0, true, true, true, true);
            }//for loop end

            TableRow TableRow = CU_DOCX.MakeTableRow(tbRow);
            //ThisRow is for cells data
            TableRow TableCells = CU_DOCX.MakeTableRow(tbCell);

            Table tableRow = CU_DOCX.MakeTable(BorderValues.None, true, TableRow);
            Table tableCell = CU_DOCX.MakeTable(BorderValues.None, true, TableCells);
            // Table table2 = CU_DOCX.MakeTable(BorderValues.None, true, TableCell2);
            CU_DOCX.WriteTable(doc, tableRow);
            CU_DOCX.WriteTable(doc, tableCell);

        }//InsertDataInRowsAndCells using docX object end

        //WriteImage; This subroutine will put data in the Row and Cells using a docX object
        public static void CreateTableWithImageDocx(WordprocessingDocument doc, string[] DataForRow, string[] DataForCells, string StyleName, JustificationValues position, bool DoYouWantBorders = true, string BorderColor = "228B22", string fontSize = "24", string fontColor = "black", bool isBold = false, bool isUnderLine = false, bool isItalic = false)
        {
            //Register the style
            StyleRunProperties newStyleProperty = CU_DOCX.MakeStyleRunProps("Segoe UI", fontSize, fontColor, isBold, isUnderLine, isItalic);
            CU_DOCX.AddNewPStyle(doc, StyleName, StyleName, newStyleProperty);

            // int TemproryLenght = DataForCells.Length;
            TableCell[] tbRow = new TableCell[DataForRow.Length];
            TableCell[] tbCell = new TableCell[DataForCells.Length];

            //to check if a string is an picuture Id
            Regex r = new Regex(@"\d+");
            //loop through array containng data data for Row
            for (int k = 0; k < DataForRow.Length; k++)
            {
                if (DataForRow.Length == 0) return;
                //if content to insert in table is a picture
                if (r.IsMatch(DataForRow[k]))
                {
                    tbRow[k] = CU_DOCX.MakeTableCell(CU_DOCX.MakeParagraph(StyleName, position, CU_DOCX.AddImageAsRun(DataForRow[k])));
                    if (DoYouWantBorders)
                        // CU_DOCX.ToggleTableCellBorders(tbRow[k], 0, true, true, true, true);
                        ToggleTableCellBordersWithColor(tbRow[k], 0, BorderColor, true, true, true, true);
                }
                else
                {//if the content is just text not pictures
                    tbRow[k] = CU_DOCX.MakeTableCell(CU_DOCX.MakeParagraph(StyleName, position, CU_DOCX.MakeRun(DataForRow[k])));
                    if (DoYouWantBorders)
                        // CU_DOCX.ToggleTableCellBorders(tbRow[k], 0, true, true, true, true);
                        ToggleTableCellBordersWithColor(tbRow[k], 0, BorderColor, true, true, true, true);
                }//else ends
            }//forloop ends for Row (content for row)

            //loop through array containng data data for cells
            for (int v = 0; v < DataForRow.Length; v++)
            {
                if (DataForCells.Length == 0) return;

                //if content is a picture
                if (r.IsMatch(DataForCells[v]))
                {
                    tbCell[v] = CU_DOCX.MakeTableCell(CU_DOCX.MakeParagraph(StyleName, position, CU_DOCX.AddImageAsRun(DataForCells[v])));
                    if (DoYouWantBorders)
                        //CU_DOCX.ToggleTableCellBorders(tbCell[v], 0, true, true, true, true);
                        ToggleTableCellBordersWithColor(tbCell[v], 0, BorderColor, true, true, true, true);
                }
                else
                {//if content is not a picture just text
                    tbCell[v] = CU_DOCX.MakeTableCell(CU_DOCX.MakeParagraph(StyleName, position, CU_DOCX.MakeRun(DataForCells[v])));
                    if (DoYouWantBorders)
                        // CU_DOCX.ToggleTableCellBorders(tbCell[v], 0, true, true, true, true);
                        ToggleTableCellBordersWithColor(tbCell[v], 0, BorderColor, true, true, true, true);
                }//else ends
            }//for loop end

            //This Table Row is for Header row
            TableRow TableRow = CU_DOCX.MakeTableRow(tbRow);
            //This Row is for cells data
            TableRow TableCells = CU_DOCX.MakeTableRow(tbCell);

            Table tableRow = MakeTable(BorderValues.None, true, "100", TableRow);
            Table tableCell = MakeTable(BorderValues.None, true, "100", TableCells);

            CU_DOCX.WriteTable(doc, tableRow);
            CU_DOCX.WriteTable(doc, tableCell);

        }//InsertDataInRowsAndCells using docX object end

        //This subroutine will write an InLine Sentence
        public static void InlineSentence(WordprocessingDocument docx, string inLineTitle, string inLineTitleChild, JustificationValues values)
        {
            Run r1 = CU_DOCX.MakeRun(inLineTitle, "C0B");
            Run r2 = CU_DOCX.MakeRun(inLineTitleChild, "S01");
            CU_DOCX.WriteParagraph(docx, CU_DOCX.MakeParagraph("S00", values, r1, r2));
        }//inLineParagraph ends

        //This subroutine will customize a table to put paddings inside the cells.For Exampl, image in the cells never used to have padding
        private static Table MakeTable(BorderValues bvV = BorderValues.None, bool FullWidth = false, string tableMargin = "10", params TableRow[] trs)
        {
            Table table = new Table();
            // Create a TableProperties object and specify its border information.
            TableProperties tblProp = new TableProperties(

                new TableBorders(

                    new TopBorder() { Val = new EnumValue<BorderValues>(bvV), Size = 24 },

                    // new TopBorder() { Color = "FFFF00" },
                    new BottomBorder() { Val = new EnumValue<BorderValues>(bvV), Size = 24 },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(bvV), Size = 24 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(bvV), Size = 24 },
                    new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(bvV), Size = 24 },
                    new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(bvV), Size = 24 }
                )

            );


            if (FullWidth) //sets preferred width to ~A4 width; default is TableWidth() { Type = TableWidthUnitValues.Auto, Width = "0" }
                tblProp.Append(new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = "9000" });

            //manage deafult table margins

            TableCellMarginDefault m = new TableCellMarginDefault();
            m.TableCellLeftMargin = new TableCellLeftMargin() { Width = 100, Type = TableWidthValues.Dxa };
            m.TableCellRightMargin = new TableCellRightMargin() { Width = 100, Type = TableWidthValues.Dxa };
            m.TopMargin = new TopMargin() { Width = tableMargin, Type = TableWidthUnitValues.Dxa };
            m.BottomMargin = new BottomMargin() { Width = tableMargin, Type = TableWidthUnitValues.Dxa };
            tblProp.TableCellMarginDefault = m;

            // Append the TableProperties object to the empty table.
            table.AppendChild<TableProperties>(tblProp);

            // Append table rows to the table.
            if (trs == null) return table;

            for (int i = 0; i < trs.Length; i++)
                table.Append(trs[i]);

            return table;
        }//Make table ends

        //This subroutine will allow changing border color of a table
        private static void ToggleTableCellBordersWithColor(TableCell tc, Int16 custWidth = 0, string BorderColor = "", bool left = false, bool right = false, bool top = false, bool bottom = false, TableVerticalAlignmentValues Vali = TableVerticalAlignmentValues.Top)
        {//adds/removes thin borders to the table cell, can also reset cellwidth in Dxa units: 1440 = 1 inch
            if (tc == null) return;

            TableCellProperties tcProps = new TableCellProperties();

            if (custWidth > 0)
                tcProps.Append(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = custWidth.ToString() });

            TableCellBorders tcBds = new TableCellBorders();

            if (left) tcBds.LeftBorder = new LeftBorder() { Val = BorderValues.BasicThinLines, Color = BorderColor };
            if (right) tcBds.RightBorder = new RightBorder() { Val = BorderValues.BasicThinLines, Color = BorderColor };
            if (top) tcBds.TopBorder = new TopBorder() { Val = BorderValues.BasicThinLines, Color = BorderColor };
            if (bottom) tcBds.BottomBorder = new BottomBorder() { Val = BorderValues.BasicThinLines, Color = BorderColor };

            TableCellVerticalAlignment tcVA = new TableCellVerticalAlignment() { Val = Vali };

            //overwrite
            tcProps.TableCellVerticalAlignment = tcVA;      //tcProps.Append(tcVA);
            tcProps.TableCellBorders = tcBds;               //tcProps.Append(tcBds);
            tc.TableCellProperties = tcProps;               //tc.Append(tcProps);            

        }//ToggleTableCell Ends

        //Ernest; This subroutine will automatically register Style to the document
        public static void RegisterStyle(WordprocessingDocument doc)
        {

            CU_DOCX.AddNewPStyle(doc, "S00", "S00", CU_DOCX.MakeStyleRunProps("Segoe UI", "24", ""));                     //default
            CU_DOCX.AddNewPStyle(doc, "S01", "S01", CU_DOCX.MakeStyleRunProps("Segoe UI", "24", "808080"));               //shaded
            CU_DOCX.AddNewPStyle(doc, "S02", "S02", CU_DOCX.MakeStyleRunProps("Segoe UI", "24", "FF0000", B: true));      //red
            CU_DOCX.AddNewPStyle(doc, "S03", "S03", CU_DOCX.MakeStyleRunProps("Segoe UI", "20", ""));                     //small;
            CU_DOCX.AddNewPStyle(doc, "S04", "S04", CU_DOCX.MakeStyleRunProps("Segoe UI", "20", "808080"));               //small and shaded;
            CU_DOCX.AddNewPStyle(doc, "S12", "S12", CU_DOCX.MakeStyleRunProps("Segoe UI", "24", "F6546A"));               //pale red
            CU_DOCX.AddNewPStyle(doc, "S13", "S13", CU_DOCX.MakeStyleRunProps("Segoe UI", "24", "0010F0"));               //blue

            CU_DOCX.AddNewPStyle(doc, "S0B", "S0B", CU_DOCX.MakeStyleRunProps("Segoe UI", "24", "", B: true));
            CU_DOCX.AddNewPStyle(doc, "S0U", "S0U", CU_DOCX.MakeStyleRunProps("Segoe UI", "24", "", U: true));
            CU_DOCX.AddNewPStyle(doc, "S0I", "S0I", CU_DOCX.MakeStyleRunProps("Segoe UI", "24", "", I: true));

            CU_DOCX.AddNewCStyle(doc, "C00", "C00", "S00", CU_DOCX.MakeStyleRunProps("Segoe UI", "24", ""));
            CU_DOCX.AddNewCStyle(doc, "C01", "C01", "S01", CU_DOCX.MakeStyleRunProps("Segoe UI", "24", "808080"));
            CU_DOCX.AddNewCStyle(doc, "C03", "C03", "S03", CU_DOCX.MakeStyleRunProps("Segoe UI", "20", ""));               //small               
            CU_DOCX.AddNewCStyle(doc, "C13", "C13", "S13", CU_DOCX.MakeStyleRunProps("Segoe UI", "24", "0010F0"));         //blue

            CU_DOCX.AddNewCStyle(doc, "C0B", "C0B", "S0B", CU_DOCX.MakeStyleRunProps(B: true));
            CU_DOCX.AddNewCStyle(doc, "C0U", "C0U", "S0U", CU_DOCX.MakeStyleRunProps(U: true));
            CU_DOCX.AddNewCStyle(doc, "C0I", "C0I", "S0I", CU_DOCX.MakeStyleRunProps(I: true));

        }//Register style ends 


    }//End of DynamicDocx
}//End of HTML_to_DOCX
