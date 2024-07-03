package org.example;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.xdocreport.converter.ConverterTypeTo;
import fr.opensagres.xdocreport.converter.ConverterTypeVia;
import fr.opensagres.xdocreport.converter.Options;
import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.images.FileImageProvider;
import fr.opensagres.xdocreport.document.images.IImageProvider;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import fr.opensagres.xdocreport.template.formatter.FieldsMetadata;
import lombok.Data;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
    /**
     * 组装map参数到合同模板中
     */
    public void assembleMap(Map<String, Object> map) {
        for (String key : map.keySet()) {
            Object value = map.get(key);
            if (value instanceof List) {
                List list = (List) value;
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

    /**
     * 普通文本
     */
    public void addText(String key, Object text) {
        this.iContext.put(key, text);
    }


    /**
     * 普通对象
     */
    public void addObject(String key, Class<?> c, Object o) {
        try {
            this.fieldsMetadata.load(key, c);
            this.iContext.put(key, o);
        } catch (XDocReportException e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }


    /**
     * List对象
     */
    public void addList(String key, Class<?> c, List list) {
        try {
            this.fieldsMetadata.load(key, c, true);
            this.iContext.put(key, list);
        } catch (XDocReportException e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }

    /**
     * 图片
     */
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


    /**
     * 生成Word
     */
    public void createWord(OutputStream outputStream) {
        try {
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
            this.ixDocReport.setFieldsMetadata(this.fieldsMetadata);
            this.ixDocReport.convert(this.iContext, options, outputStream);
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage(), e);
        }
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
            Map<String, Object> map = new HashMap<>();
            map.put("username", "Hai");
            map.put("phoneNumber", "012349929");
            map.put("organization", "DGC");
            map.put("department", "MPTC");
            map.put("office", "floor1");
            map.put("reason", "សូម");
            map.put("age", "20");
            map.put("position", "backend");
            main.addImage("img",new File("src/main/resources/images/lunlimhai.jpg"));
            main.assembleMap(map);

            File wordOutputFile = new File("src/main/resources/templates/ss.docx");
             OutputStream outputStream = new FileOutputStream(wordOutputFile);
            main.createWord(outputStream);

////            // Generate PDF
//            File pdfOutputFile = new File("src/main/resources/templates/ok_template.pdf"); // Fixed output file name
//            OutputStream outputStream2 = new FileOutputStream(pdfOutputFile);
//            wordConverterToPdf(source,outputStream2);
        }catch (Exception e){
            throw new RuntimeException("Word template generation file error", e);
        }

    }



}