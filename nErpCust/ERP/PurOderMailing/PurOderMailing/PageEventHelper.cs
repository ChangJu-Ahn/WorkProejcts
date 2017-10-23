using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OrderDraftMailSending
{
    class PageEventHelper : PdfPageEventHelper
    {
        PdfContentByte cb;
        PdfTemplate template;
        iTextSharp.text.Font fntPage;
        iTextSharp.text.Rectangle rectNumbering;

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            cb = writer.DirectContent;
            template = cb.CreateTemplate(50, 50);
            FontFactory.RegisterDirectory(@"C:\WINDOWS\Fonts");
            fntPage = FontFactory.GetFont("맑은 고딕", BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 8f, iTextSharp.text.Font.BOLD, BaseColor.BLACK);


            rectNumbering = new iTextSharp.text.Rectangle(document.PageSize.Width, document.PageSize.Height, 1f, document.PageSize.Height - 10f);

        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {

            //ColumnText colPageTxt = new ColumnText(writer.DirectContent);
            //Paragraph para = new Paragraph("", fntPage);
            //int pageN = writer.PageNumber;
            //String text = "Page " + pageN.ToString() + " of ";
            //float len = fntPage.BaseFont.GetWidthPoint(text, fntPage.Size);

            //para.Add(text);
            //colPageTxt.SetSimpleColumn(rectNumbering);
            //colPageTxt.AddElement(para);
            //colPageTxt.Go();


            base.OnEndPage(writer, document);

            int pageN = writer.PageNumber;
            String text = "Page " + pageN.ToString() + " of ";
            float len = fntPage.BaseFont.GetWidthPoint(text, fntPage.Size);

            iTextSharp.text.Rectangle pageSize = document.PageSize;

            cb.SetRGBColorFill(100, 100, 100);

            cb.BeginText();
            cb.SetFontAndSize(fntPage.BaseFont, fntPage.Size);
            cb.SetTextMatrix(document.Right/2, pageSize.GetBottom(document.BottomMargin));
            cb.ShowText(text);

            cb.EndText();

            cb.AddTemplate(template, document.Right/2 + len, pageSize.GetBottom(document.BottomMargin));
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);

            template.BeginText();
            template.SetFontAndSize(fntPage.BaseFont, fntPage.Size);
            template.SetTextMatrix(0, 0);
            template.ShowText("" + (writer.PageNumber));
            template.EndText();
        }

    }
}
