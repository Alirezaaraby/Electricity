package com.guilanzamineng;

import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.GrayColor;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfWriter;

public class HeaderFooterPageEvent extends PdfPageEventHelper {


    public void onEndPage(PdfWriter writer, Document document) {
        Font Copy_right_font = new Font(Font.FontFamily.HELVETICA,5, Font.ITALIC, GrayColor.GRAYBLACK);

        ColumnText.showTextAligned(writer.getDirectContent(), Element.ALIGN_CENTER, new Phrase("#Coded By Alirezaaraby",Copy_right_font), 306, 10, 0);
    }
}