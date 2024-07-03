/**
 *
 */
package org.example.test.xDocReport;

import com.google.common.io.Files;
import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.pdf.*;
import fr.opensagres.xdocreport.converter.ConverterTypeTo;
import lombok.NoArgsConstructor;
import lombok.SneakyThrows;

import java.io.*;
import java.util.Map;

/**
 * @author YingHan 2021-12-28
 *
 */
@NoArgsConstructor
public class XDocReportUtil extends XDocReport {

    public XDocReportUtil(String template_fullpath, Map<String, Object> paramMap) {
        setTemplatePath(template_fullpath);
        setParamMap(paramMap);
    }

    @SneakyThrows
    @Override
    public ByteArrayOutputStream addWatermark(String inputFilePath, String outputFilePath, String watermark) {
        FileInputStream inputStream = new FileInputStream(inputFilePath);
        FileOutputStream outputStream = new FileOutputStream(outputFilePath);
        this.setWatermark(inputStream, outputStream, watermark);
        ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
        byteOut.write(Files.toByteArray(new File(outputFilePath)));
        inputStream.close();
        outputStream.close();
        byteOut.close();
        return byteOut;
    }

    public ByteArrayOutputStream download(String outputFileName, ConverterTypeTo typeto, boolean deleteTemp, String watermark) throws Exception {
        String rootPath = "temp/";
        ByteArrayOutputStream outputStream = build(rootPath, outputFileName, typeto, watermark);
        if (deleteTemp) {
            new File(rootPath + outputFileName).delete();
        }
        return outputStream;
    }

    /**
     * Add text water mark
     */
    private void setWatermark(InputStream inputStream, OutputStream outputStream, String watermark)
            throws Exception {

        int repeat = 3;
        int fontSize = 40;
        float opacity = 0.5f;

        Document document = new Document(PageSize.A4);
        //Read the existing PDF document
        PdfReader pdfReader = new PdfReader(inputStream);
        //Get the PdfStamper object
        PdfStamper pdfStamper = new PdfStamper(pdfReader, outputStream);

        //Get the PdfContentByte type by pdfStamper.
        for (int i = 1, pdfPageSize = pdfReader.getNumberOfPages() + 1; i < pdfPageSize; i++) {
            PdfContentByte pageContent = pdfStamper.getOverContent(i);
            pageContent.setGState(this.getPdfGState(opacity));
            pageContent.beginText();
            pageContent.setFontAndSize(this.getBaseFont(), fontSize);
            pageContent.setColorFill(new BaseColor(220, 220, 220));
            //            pageContent.showTextAligned(Element.ALIGN_CENTER, watermark, document.getPageSize().getWidth() / 2,
            //                    document.getPageSize().getHeight() / 2, 30);
            for (int x = 0; x <= repeat; x++) {
                for (int y = 0; y <= repeat; y++) {
                    pageContent.showTextAligned(Element.ALIGN_CENTER, watermark,
                            document.getPageSize().getWidth() / repeat * x, document.getPageSize().getHeight() / repeat * y,
                            45);
                }
            }
            pageContent.endText();
        }
        pdfStamper.close();
    }

    /**
     * Get BaseFont
     */
    private BaseFont getBaseFont() throws Exception {
        return BaseFont.createFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.EMBEDDED);
    }

    /**
     * Get PdfGState
     */
    private PdfGState getPdfGState(float opacity) {
        PdfGState graphicState = new PdfGState();
        graphicState.setFillOpacity(opacity);
        graphicState.setStrokeOpacity(1f);
        return graphicState;
    }
}
