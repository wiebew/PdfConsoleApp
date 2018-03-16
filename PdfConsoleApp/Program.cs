using Bogus;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.IO;

namespace PdfConsoleApp
{
    internal class BadgeGenerator
    {
        public enum PersonType { user, organizer, speaker }

        private const int PageHeight = 842;
        private const int PageWidth = 595;
        private const int CardWidth = 253;
        private const int CardHeight = 153;
        private const int XMargin = (PageWidth - 2 * CardWidth) / 2;
        private const int YMargin = (PageHeight - 5 * CardHeight) / 2;
        private const float BorderWidth = 0.2f;
        private const int SquareSize = 20;

        private readonly Document _document;
        private readonly PdfContentByte _contentByte;
        private readonly BaseFont _baseFont;
        private readonly Font _fontOrganization;
        private readonly Font _fontBannerOrange;
        private readonly Font _fontBannerGray;
        private readonly Image _referenceImage;
        private readonly Chunk _orangeChunk;
        private readonly Chunk _grayChunk;
        private readonly Font _fontOrganizerName;
        private readonly Font _fontSpeakerName;
        private readonly Font _fontUserName;
        private readonly Font _fontSession;


        public BadgeGenerator(Stream stream, string imagePath)
        {
            _baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, true);

            _fontOrganizerName = new Font(_baseFont, 24.0f, Font.BOLD, new BaseColor(112, 48, 160));
            _fontSpeakerName = new Font(_baseFont, 24.0f, Font.BOLD, new BaseColor(0, 176, 240));
            _fontUserName = new Font(_baseFont, 24.0f, Font.BOLD, new BaseColor(0, 0, 0));
            _fontSession = new Font(_baseFont, 16.0f, Font.NORMAL);
            _fontOrganization = new Font(_baseFont, 16.0f, Font.ITALIC);
            _fontBannerOrange = new Font(_baseFont, 32.0f, Font.NORMAL, new BaseColor(236, 82, 0));
            _fontBannerGray = new Font(_baseFont, 32.0f, Font.NORMAL, new BaseColor(220, 220, 220));
            _referenceImage = Image.GetInstance("C:/Users/Wiebe/Pictures/badge_banner.jpg");
            _referenceImage.ScaleAbsolute(CardWidth, 47.0f);

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

        private Font SelecctFontForPerson(PersonType personType)
        {
            Font fontName = _fontUserName;

            if (personType == PersonType.organizer)
            {
                fontName = _fontOrganizerName;
            }
            if (personType == PersonType.speaker)
            {
                fontName = _fontSpeakerName;
            }
            return fontName;
        }

        private void AddOrganizatonCell(PdfPTable table, string organisation)
        {
            PdfPCell cell = new PdfPCell(new Phrase(String.Format(organisation), _fontOrganization)) { Colspan = 6 };
            cell.FixedHeight = 27.0f;
            cell.BorderWidth = 0;
            cell.PaddingBottom = 10.0f;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_TOP;
            table.AddCell(cell);
        }

        private void AddPersonCell(PdfPTable table, PersonType personType, string name)
        {
            Font fontName = SelecctFontForPerson(personType);

            PdfPCell cell = new PdfPCell(new Phrase(String.Format(name), fontName)) { Colspan = 6 };
            cell.FixedHeight = 51.0f;
            cell.BorderWidth = 0;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            table.AddCell(cell);
        }
        private void AddPhotoCell(PdfPTable table)
        {
            /* the cell with the photo and event name */
            Chunk orangeChunk = new Chunk(_orangeChunk);
            Chunk grayChunk = new Chunk(_grayChunk);
            var confphrase = new Phrase(orangeChunk);
            confphrase.Add(grayChunk);

            PdfPCell cell = new PdfPCell(confphrase) { Colspan = 6 };
            cell.FixedHeight = 47.0f;
            cell.BorderWidth = 0;
            cell.PaddingBottom = 10.0f;
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            table.AddCell(cell);
        }

        private void AddSessionsCells(PdfPTable table)
        {
            /* the cell at the bottom to print the squares with session in*/
            PdfPCell cell = new PdfPCell(new Phrase(String.Format("WW"), _fontSession));
            cell.FixedHeight = 21.0f;
            cell.BorderWidth = 1.1f;
            cell.BorderColor = new BaseColor(255, 255, 255);
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.BackgroundColor = new BaseColor(0,255,0);
            cell.Padding = 5.0f;
            table.AddCell(cell);
            table.AddCell(cell);
            table.AddCell(cell);
            table.AddCell(cell);
            table.AddCell(cell);
            table.AddCell(cell);
        }
        private void DrawSquare(float absoluteX, float absoluteY)
        {
            _contentByte.SetColorFill(new BaseColor(255, 0, 0));
            _contentByte.MoveTo(absoluteX, absoluteY);
            _contentByte.LineTo(absoluteX, absoluteY+SquareSize);
            _contentByte.LineTo(absoluteX + SquareSize, absoluteY + SquareSize);
            _contentByte.LineTo(absoluteX + SquareSize, absoluteY );
            _contentByte.Fill();
        }

        private void AddCard(float absoluteX, float absoluteY, string name, string organisation, PersonType personType)
        {

            /* create a table to put the info in */
            PdfPTable table = new PdfPTable(6);
            table.SetTotalWidth(new float[] { CardWidth / 6, CardWidth / 6, CardWidth / 6, CardWidth / 6, CardWidth / 6, CardWidth / 6});
            table.LockedWidth = true;

            AddPersonCell(table, personType, name);
            AddOrganizatonCell(table, organisation);
            AddPhotoCell(table);
            AddSessionsCells(table);

            /* write table at specified coordinates */
            table.WriteSelectedRows(0, -1, absoluteX, absoluteY + CardHeight, _contentByte);
     
            /* position image in the cell for them image and event name */
            Image image = Image.GetInstance(_referenceImage);
            image.SetAbsolutePosition(absoluteX, absoluteY + 22.0f);
            _document.Add(image);
        }

        public void FillPages()
        {
            Randomizer.Seed = new Random(8675309);
            Faker faker = new Faker("nl");

            int colCount = 0;
            int rowCount = 0;

            for (int i = 0; i < 250; i++)
            {
                string name = faker.Lorem.Sentence(4);
                AddCard(XMargin + colCount * CardWidth, YMargin + rowCount * CardHeight, name, "RDW-ICT", PersonType.user);

                colCount++; // alternate over columns
                if (colCount > 1)
                {
                    rowCount++; /* add a row when starting at column 0 */
                    colCount = 0;
                }
                if (rowCount > 4)
                {
                    // end of page reached, add new page, reset counters and draw grid for cutting the paper */
                    AddPageMarkers();
                    _document.NewPage();
                    rowCount = 0;
                }
            }
            AddPageMarkers();
            _document.Close();
        }
    }

    internal class Program
    {
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