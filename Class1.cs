using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;
 

namespace ConsoleApp1
{
    class Class1
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table tbl = builder.StartTable();

            builder.InsertCell();
            builder.Write("region");

            tbl.PreferredWidth = PreferredWidth.FromPercent(100);
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(10);
            builder.CellFormat.Borders.LineWidth = 1.5;
            builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.SteelBlue;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;


            builder.InsertCell();
            builder.Write("country");
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(10);


            builder.InsertCell();
            builder.Write("symbol");
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(10);

            builder.EndRow();

            foreach(Cell cell in tbl.Rows[0].Cells)
            {
                Paragraph par = cell.FirstParagraph;
                par.Runs[0].Font.Bold = true;
                par.Runs[0].Font.Color = System.Drawing.Color.White;


            }
            Cell currentcell;
            currentcell = builder.InsertCell();
            builder.Write("oceania");
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            currentcell.CellFormat.VerticalMerge = CellMerge.First;
            builder.CellFormat.Borders.LineWidth = 1;
            builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.White;
            currentcell = builder.InsertCell();
            currentcell.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("india");
                







            doc.Save("G:\\fahad.docx");
        }
    }
}