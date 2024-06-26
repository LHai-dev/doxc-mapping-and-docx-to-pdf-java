package org.example;

import com.documents4j.api.DocumentType;
import com.documents4j.job.LocalConverter;
import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.pdf.BaseFont;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.xdocreport.converter.*;
import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.images.FileImageProvider;
import fr.opensagres.xdocreport.document.images.IImageProvider;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import fr.opensagres.xdocreport.template.formatter.FieldsMetadata;
import lombok.Data;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.awt.*;
import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

@Data
public class Main {
    private FieldsMetadata fieldsMetadata;
    private IXDocReport ixDocReport;
    private IContext iContext;

    public  Main(InputStream inputStream) {
        try {
            this.ixDocReport = XDocReportRegistry.getRegistry().loadReport(inputStream, TemplateEngineKind.Freemarker);
            this.iContext = ixDocReport.createContext();
            this.fieldsMetadata = ixDocReport.createFieldsMetadata();
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }
    public void assembleMap(Map<String, Object> map) {
        for (String key : map.keySet()) {
            Object value = map.get(key);
            if (value instanceof List list) {
                addList(key, list.get(0).getClass(), list);
            } else if (value instanceof String) {
                addText(key, value.toString());
            } else if (value instanceof File) {
                addImage(key, (File) value);
            } else {
                addObject(key, value.getClass(), value);
            }
        }
    }

    public void addText(String key, Object text) {
        this.iContext.put(key, text);
    }

    public void addObject(String key, Class<?> c, Object o) {
        try {
            this.fieldsMetadata.load(key, c);
            this.iContext.put(key, o);
        } catch (XDocReportException e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }
    public void addList(String key, Class<?> c, List list) {
        try {
            this.fieldsMetadata.load(key, c, true);
            this.iContext.put(key, list);
        } catch (XDocReportException e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }
    public void addImage(String key, File file) {
        this.fieldsMetadata.addFieldAsImage(key);
        IImageProvider img = new FileImageProvider(file);
        this.iContext.put(key, img);
    }

    public void addImage(String key, File file, float width, float height) {
        this.fieldsMetadata.addFieldAsImage(key);
        IImageProvider img = new FileImageProvider(file, true);
        img.setSize(width, height);
        this.iContext.put(key, img);
    }

    public void createWord(OutputStream outputStream) {
        try {
            this.ixDocReport.setCacheOriginalDocument(true);
            this.ixDocReport.setFieldsMetadata(this.fieldsMetadata);
            this.ixDocReport.process(this.iContext, outputStream);
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }
    /**
     * 生成Pdf
     */
    public void createPdf(OutputStream outputStream) {
        try {
            Options options = Options.getTo(ConverterTypeTo.PDF).via(ConverterTypeVia.XWPF);
            PdfOptions pdfOptions = getPdfOptions();

            options.subOptions(pdfOptions);
            IConverter converter = ConverterRegistry.getRegistry().getConverter(options);

            if (converter == null) {
                System.err.println("Converter not found in registry.");
                return; // Exit early if converter is null
            }

            // Assuming input is your Word template file
            FileInputStream is = new FileInputStream("src/main/resources/templates/ok_template.docx");
            try {
                converter.convert(is, outputStream, options);
            } finally {
                is.close();
            }

        } catch (Exception e) {
            throw new RuntimeException("Error generating PDF", e);
        } finally {
            try {
                outputStream.close(); // Close the output stream in finally block
            } catch (IOException e) {
                throw new RuntimeException("Error closing output stream", e);
            }
        }
    }
    private static PdfOptions getPdfOptions() {
        PdfOptions pdfOptions = PdfOptions.create();

        // Set custom font provider for PDF generation
        pdfOptions.fontProvider(new IFontProvider() {
            @Override
            public Font getFont(String familyName, String encoding, float size, int style, Color color) {
                try {
                    if (familyName.equalsIgnoreCase("KHMERMPTCMOUL")) {
                        BaseFont baseFont = BaseFont.createFont("C:\\Users\\Hai\\AppData\\Local\\Microsoft\\Windows\\Fonts\\KHMERMPTCMOUL.OTF", encoding, BaseFont.EMBEDDED);
                        return new Font(baseFont, size, style, color);
                    } else if (familyName.equalsIgnoreCase("KHMERMPTCMOUL")) {
                        BaseFont baseFont = BaseFont.createFont("C:\\Users\\Hai\\AppData\\Local\\Microsoft\\Windows\\Fonts\\KHMERMPTCMOUL.OTF", encoding, BaseFont.EMBEDDED);
                        return new Font(baseFont, size, style, color);
                    }
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
                return FontFactory.getFont(familyName, encoding, size, style, color);
            }
        });
        return pdfOptions;
    }

    /**
     * word(比较胡扯，只支持docx) 转 pdf
     */
    public static void wordConverterToPdf(InputStream source, OutputStream target) throws Exception {
        XWPFDocument doc = new XWPFDocument(source);
        PdfConverter.getInstance().convert(doc, target, null);
    }


    public static void main(String[] args) {
        try {
            File template = new File("src/main/resources/templates/Leave_Request_Template.docx");
            Main main = new Main(new FileInputStream(template));
            File wordOutputFile = new File("src/main/resources/templates/ok_template.docx");
            OutputStream outputStream1 = new FileOutputStream(wordOutputFile);
            main.createWord(outputStream1);

            // Generate PDF
            File pdfOutputFile = new File("src/main/resources/templates/ok_template.pdf"); // Fixed output file name
            OutputStream outputStream2 = new FileOutputStream(pdfOutputFile);
            main.createPdf(outputStream2);
        }catch (Exception e){
            throw new RuntimeException("Word template generation file error", e);
        }

    }

    private static Main getMain(File template) throws Exception {
        Main main = new Main(new FileInputStream(template));
        Map<String, Object> map = new HashMap<>();
        map.put("username", "Hai");
        map.put("phoneNumber", "012349929");
        map.put("organization", "DGC");
        map.put("department", "MPTC");
        map.put("office", "floor1");
        map.put("reason", "សូម");
        map.put("age", "20");
        map.put("position", "backend");
        main.assembleMap(map);
        main.addImage("img",new File("src/main/resources/images/lunlimhai.jpg"));

//        IConverter converter = LocalConverter.builder()
//                .baseFolder(new File("src/main/resources/"))
//                .workerPool(20, 25, 2, TimeUnit.SECONDS)
//                .processTimeout(3, TimeUnit.SECONDS)
//                .build();
//        File docxFile = new File("src/main/resources/templates/ok_template.docx");
//        File pdfFile = new File("src/main/resources/templates/ok_contemplations.pdf");
//        converter
//                .convert(docxFile).as(DocumentType.DOCX)
//                .to(pdfFile).as(DocumentType.PDF)
//                .schedule();
        return main;
    }


}