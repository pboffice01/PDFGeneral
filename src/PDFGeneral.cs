using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;

public class PdfGeneral
{
    public class CellSpan{
  	public int CellIndex { get; set; }
  	public int SpanCount { get; set; }
    }
 
    public class CellBorder{
      public int CellIndex { get; set; }
      public int CellBord { get; set; }
    }
 
    public class CellText{
      public int CellIndex { get; set; }
      public BaseFont TextFont { get; set; }
      public string Text { get; set; }
      public int TextOrder { get; set; }
    }

    public byte[] ConvertToPDF(
        int RowCount,
        int ColumnCunt,
        float[] RowHeith,
        float[] ColumnWidth,
        List<CellSpan> RowSpan,
        List<CellSpan> ColumSpan,
        List<CellBorder> CellBorder,
        List<CellText> CellTests
    )
    {
        if (RowCount <= 0) throw new ArgumentOutOfRangeException(nameof(RowCount));
        if (ColumnCunt <= 0) throw new ArgumentOutOfRangeException(nameof(ColumnCunt));

        RowHeith ??= Array.Empty<float>();
        ColumnWidth ??= Array.Empty<float>();
        RowSpan ??= new List<CellSpan>();
        ColumSpan ??= new List<CellSpan>();
        CellBorder ??= new List<CellBorder>();
        CellTests ??= new List<CellText>();
 
        var rowSpanMap = RowSpan
            .GroupBy(x => x.CellIndex)
            .ToDictionary(g => g.Key, g => g.First().SpanCount);

        var colSpanMap = ColumSpan
            .GroupBy(x => x.CellIndex)
            .ToDictionary(g => g.Key, g => g.First().SpanCount);

        var borderMap = CellBorder
            .GroupBy(x => x.CellIndex)
            .ToDictionary(g => g.Key, g => g.First().CellBord);

        var textMap = CellTests
            .GroupBy(x => x.CellIndex)
            .ToDictionary(
                g => g.Key,
                g => g.OrderBy(x => x.TextOrder).ToList()
            );
 
        var covered = new bool[RowCount, ColumnCunt];

        using (var ms = new MemoryStream())
        {
            using (var doc = new Document(PageSize.A4, 36f, 36f, 36f, 36f))
            {
                PdfWriter.GetInstance(doc, ms);
                doc.Open();

                var table = new PdfPTable(ColumnCunt)
                {
                    WidthPercentage = 100f
                };

                 
                if (ColumnWidth.Length >= ColumnCunt)
                {
                    var widths = new float[ColumnCunt];
                    for (int c = 0; c < ColumnCunt; c++)
                    {
                        widths[c] = ColumnWidth[c];
                        if (widths[c] <= 0) widths[c] = 1f;
                    }
                    table.SetWidths(widths);
                }

                for (int r = 0; r < RowCount; r++)
                {
                    for (int c = 0; c < ColumnCunt; c++)
                    {
                        if (covered[r, c]) continue;

                        int cellIndex = r * ColumnCunt + c;

                        int rs = 1;
                        if (rowSpanMap.TryGetValue(cellIndex, out int rowSpanCount) && rowSpanCount > 0)
                            rs = rowSpanCount + 1;

                        int cs = 1;
                        if (colSpanMap.TryGetValue(cellIndex, out int colSpanCount) && colSpanCount > 0)
                            cs = colSpanCount + 1;

                         
                        if (r + rs > RowCount) rs = RowCount - r;
                        if (c + cs > ColumnCunt) cs = ColumnCunt - c;
 
                        for (int rr = r; rr < r + rs; rr++)
                        for (int cc = c; cc < c + cs; cc++)
                        {
                            if (rr == r && cc == c) continue;
                            covered[rr, cc] = true;
                        }
 
                        Phrase phrase = BuildPhrase(textMap, cellIndex);

                        var cell = new PdfPCell(phrase)
                        {
                            HorizontalAlignment = Element.ALIGN_LEFT,
                            VerticalAlignment = Element.ALIGN_MIDDLE,
                            Padding = 3f,
                            Colspan = cs,
                            Rowspan = rs
                        };
 
                        float minH = 0f;
                        for (int rr = r; rr < r + rs; rr++)
                        {
                            float h = (RowHeith.Length > rr) ? RowHeith[rr] : 0f;
                            if (h > 0) minH += h;
                        }
                        if (minH > 0) cell.MinimumHeight = minH;
 
                        if (borderMap.TryGetValue(cellIndex, out int bw))
                        {
                            cell.BorderWidth = Math.Max(0f, bw);
                        }
                        else
                        {
                            cell.BorderWidth = 0.5f;
                        }

                        table.AddCell(cell);
                    }
                }

                doc.Add(table);
                doc.Close();
            }

            return ms.ToArray();
        }
    }
    
    public byte[] GenerateFlowImage(byte[] inputPDF, float fillOpacity, float strokeOpacity, int degrees, float scalePercent, string flowImagePath)
        {
            #region WaterMask

            PdfGState pdfgstate = new PdfGState()
            {
                FillOpacity = fillOpacity,
                StrokeOpacity = strokeOpacity
            };

            MemoryStream ms = new MemoryStream();

            using (var reader = new PdfReader(inputPDF))
            {
                using (var stamper = new PdfStamper(reader, ms))
                {
                    int times = reader.NumberOfPages;
                    for (int i = 1; i <= times; i++)
                    {
                        Rectangle pagesize = reader.GetPageSize(i); 
                                                                    
                        Chunk ctitle = new Chunk(i.ToString().Trim() + " / " + times, FontFactory.GetFont("Futura", 12f, new BaseColor(0, 0, 0)));
                        Phrase ptitle = new Phrase(ctitle);
                        
                        string imageUrl = flowImagePath;
                        Image img = iTextSharp.text.Image.GetInstance(imageUrl);
                        img.ScalePercent(scalePercent);  
                        img.RotationDegrees = degrees; 
                        img.SetAbsolutePosition(pagesize.Width / 2 - 100, pagesize.Height / 2); 


                        PdfContentByte over = stamper.GetOverContent(i);
                        ColumnText.ShowTextAligned(over, Element.ALIGN_LEFT, ptitle, pagesize.Width / 2, 10, 0); 
                        over.SetGState(pdfgstate); 
                        over.AddImage(img); 

                    }
                    stamper.Close();
                }
            }

            #endregion


            return ms.ToArray();

        }


    private Phrase BuildPhrase(Dictionary<int, List<CellText>> textMap, int cellIndex)
    {
        var phrase = new Phrase();

        if (!textMap.TryGetValue(cellIndex, out var lines) || lines == null || lines.Count == 0)
            return phrase;

        // 指定楷體
        string kaiuPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
            "KAIU.TTF"
        );


        BaseFont defaultBaseFont = BaseFont.CreateFont(kaiuPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

        for (int i = 0; i < lines.Count; i++)
        {
            var t = lines[i];
            string s = t?.Text ?? string.Empty;
 
            BaseFont bf = t?.TextFont ?? defaultBaseFont;
 
            var font = new Font(bf, 10f);

            phrase.Add(new Chunk(s, font));

            if (i != lines.Count - 1)
            {
                phrase.Add(Chunk.NEWLINE);
            }
        }

        return phrase;
    }
}
