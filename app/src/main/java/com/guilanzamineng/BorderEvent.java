package com.guilanzamineng;

import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfPTableEvent;

public class BorderEvent implements PdfPTableEvent {
    public void tableLayout(PdfPTable table, float[][] widths,
                            float[] heights, int headerRows, int rowStart,
                            PdfContentByte[] canvases) {
        float width[] = widths[0];
        float x1 = width[0];
        float x2 = width[width.length - 1];
        float y1 = heights[0];
        float y2 = heights[heights.length - 1];
        PdfContentByte cb = canvases[PdfPTable.LINECANVAS];
        cb.rectangle(x1, y1, x2 - x1, y2 - y1);
        cb.stroke();
        cb.resetRGBColorStroke();
    }
}