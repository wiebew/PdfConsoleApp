using Bogus;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.IO;

namespace PdfConsoleApp
{
    internal class BadgeGenerator
    {
        private const int PageHeight = 842;
        private const int PageWidth = 595;
        private const int CardWidth = 255;
        private const int CardHeight = 153;
        private const int XMargin = (PageWidth - 2 * CardWidth) / 2;
        private const int YMargin = (PageHeight - 5 * CardHeight) / 2;
        private const float BorderWidth = 0.2f;

        private readonly Document _document;
        private readonly PdfContentByte _contentByte;
        private readonly BaseFont _baseFont;
        private readonly Font _fontName;
        private readonly Font _fontOrganization;
        private readonly Font _fontBannerOrange;
        private readonly Font _fontBannerGray;
        private readonly Image _referenceImage;
        private readonly Chunk _orangeChunk;
        private readonly Chunk _grayChunk;

        public BadgeGenerator(Stream stream, string imagePath)
        {
            _baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, true);
            _fontName = new Font(_baseFont, 24.0f, Font.BOLD, new BaseColor(112, 48, 160));
            _fontOrganization = new Font(_baseFont, 14.0f, Font.ITALIC);
            _fontBannerOrange = new Font(_baseFont, 35.0f, Font.NORMAL, new BaseColor(220, 66, 0));
            _fontBannerGray = new Font(_baseFont, 35.0f, Font.NORMAL, new BaseColor(200, 200, 200));
            _referenceImage = Image.GetInstance("C:/Users/Wiebe/Pictures/badge_banner.jpg");
            _referenceImage.ScaleAbsolute(CardWidth, 51.0f);

            _document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.GetInstance(_document, stream);
            _document.Open();
            _contentByte = writer.DirectContent;
            _orangeChunk = new Chunk("RDW", _fontBannerOrange);
            _grayChunk = new Chunk(" Techday", _fontBannerGray);
        }

        private void AddPageMarkers()
        {
            _contentByte.SetLineWidth(BorderWidth);
            /* draw verticals */
            for (int i = 0; i < 3; i++)
            {
                _contentByte.MoveTo(XMargin + i * CardWidth, 0);
                _contentByte.LineTo(XMargin + i * CardWidth, PageHeight);
                _contentByte.Stroke();
            }

            /* draw horizontals */
            for (int i = 0; i < 6; i++)
            {
                _contentByte.MoveTo(0, YMargin + i * CardHeight);
                _contentByte.LineTo(PageWidth, YMargin + i * CardHeight);
                _contentByte.Stroke();
            }
        }

        private static BaseColor DeriveColorFromRgbString(string rgbValue)
        {
            if (rgbValue == null || !rgbValue[0].Equals('#') || (rgbValue.Length != 4 && rgbValue.Length != 7))
            {
                return null;
            }

            string rgb = rgbValue.Substring(1);
            int[] raw = new int[3];
            if (rgb.Length == 3)
            {
                for (int i = 0; i < 3; i++)
                {
                    int hex = Convert.ToByte(rgb.Substring(i, 1), 16);
                    raw[i] = hex * 16 + hex;
                }
            }
            else
            {
                for (int i = 0; i < 3; i++)
                {
                    raw[i] = Convert.ToByte(rgb.Substring(2 * i, 2), 16);
                }
            }
            return new BaseColor(raw[0], raw[1], raw[2]);
        }

        private void AddCard(float absoluteX, float absoluteY, string name, string organisation)
        {
            /* create a table to put the info in */
            PdfPTable table = new PdfPTable(1);
            table.SetTotalWidth(new float[] { CardWidth });
            table.LockedWidth = true;

            /* the cell for the name */
            PdfPCell cell = new PdfPCell(new Phrase(String.Format(name), _fontName));
            cell.FixedHeight = 51.0f;
            cell.BorderWidth = 0;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            table.AddCell(cell);

            /* the cell for the organization */
            cell = new PdfPCell(new Phrase(String.Format(organisation), _fontOrganization));
            cell.FixedHeight = 26.0f;
            cell.BorderWidth = 0;
            cell.PaddingBottom = 10.0f;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_TOP;
            table.AddCell(cell);

            /* the cell with the photo and event name */
            Chunk orangeChunk = new Chunk(_orangeChunk);
            Chunk grayChunk = new Chunk(_grayChunk);
            var confphrase = new Phrase(orangeChunk);
            confphrase.Add(grayChunk);
            cell = new PdfPCell(confphrase);
            cell.FixedHeight = 51.0f;
            cell.BorderWidth = 0;
            cell.PaddingBottom = 15.0f;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            table.AddCell(cell);

            /* the cell at the bottom to print the squares with session in*/
            cell = new PdfPCell();
            cell.FixedHeight = 25.0f;
            cell.BorderWidth = 0;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            table.AddCell(cell);

            /* write table at specified coordinates */
            table.WriteSelectedRows(0, -1, absoluteX, absoluteY + CardHeight, _contentByte);

            /* position image in the cell for them image and event name */
            Image image = Image.GetInstance(_referenceImage);
            image.SetAbsolutePosition(absoluteX, absoluteY + 25.0f);

            _document.Add(image);
        }

        public void FillPages()
        {
            Randomizer.Seed = new Random(8675309);
            Faker faker = new Faker("nl");
            AddPageMarkers();

            int colCount = 0;
            int rowCount = 0;

            for (int i = 0; i < 250; i++)
            {
                string name = faker.Lorem.Sentence(4);
                AddCard(XMargin + colCount * CardWidth, YMargin + rowCount * CardHeight, name, "RDW-ICT");

                colCount++;
                if (colCount > 1)
                {
                    rowCount++;
                    colCount = 0;
                }
                if (rowCount > 4)
                {
                    rowCount = 0;
                    _document.NewPage();
                    AddPageMarkers();
                }
            }
            _document.Close();
        }
    }

    internal class Program
    {
        private static void DrawSquare()
        {
        }

        private static void Main(string[] args)
        {
            Console.WriteLine("Creating PDF");

            string imagePath = "C:/Users/Wiebe/Pictures/badge_banner.jpg";
            FileStream stream = new FileStream("magweg.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            BadgeGenerator bg = new BadgeGenerator(stream, imagePath);
            bg.FillPages();
        }
    }
}