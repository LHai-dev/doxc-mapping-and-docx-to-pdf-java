package org.example;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.util.IOUtils;

import java.io.*;

import fr.opensagres.xdocreport.converter.Options;
import fr.opensagres.xdocreport.converter.IConverter;
import fr.opensagres.xdocreport.converter.ConverterRegistry;
import fr.opensagres.xdocreport.converter.ConverterTypeTo;
import fr.opensagres.xdocreport.core.document.DocumentKind;


public class DocxToImages {
    static void replacePictureData(XWPFPictureData source, byte[] data) {
        try ( ByteArrayInputStream in = new ByteArrayInputStream(data);
              OutputStream out = source.getPackagePart().getOutputStream();
        ) {
            byte[] buffer = new byte[2048];
            int length;
            while ((length = in.read(buffer)) > 0) {
                out.write(buffer, 0, length);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public static void main(String[] args) throws Exception {
        String docPathTemplate = "C:\\Users\\Hai\\Downloads\\220ab115-6456-4f4d-92de-3d4026d3a205_bc4a3a60-8c22-4683-878f-886121c58e94.docx";
        String docPathAfterReplacing = "./WordResult.docx";
        String pdfPath = "./WordDocument.pdf";

        // replacde picture dara in headers
        try (XWPFDocument document = new XWPFDocument(new FileInputStream(docPathTemplate));
             FileOutputStream out = new FileOutputStream(docPathAfterReplacing);
        ) {
            for (XWPFHeader header : document.getHeaderList()) {
                for (XWPFPictureData pictureData : header.getAllPictures()) {
                    String fileName =  pictureData.getFileName();
                    System.out.println(fileName);
                    if (pictureData.getPictureType() == Document.PICTURE_TYPE_JPEG) {
                        FileInputStream is = new FileInputStream("D:\\(MPTC)\\untitled1111111111\\src\\main\\resources\\images\\lunlimhai.jpg");
                        byte[] bytes = IOUtils.toByteArray(is);
                        replacePictureData(pictureData, bytes);
                        System.out.println(fileName + " replaced by stackoverflowLogo.jpg");
                    }
                }
            }
            document.write(out);
        } catch (Exception ex) {
            ex.printStackTrace();
        }

        //convert to PDF
        try (InputStream in = new FileInputStream(new File(docPathAfterReplacing));
             OutputStream out = new FileOutputStream(new File(pdfPath));
        ) {
            Options options = Options.getFrom(DocumentKind.DOCX).to(ConverterTypeTo.PDF);
            IConverter converter = ConverterRegistry.getRegistry().getConverter(options);
            converter.convert(in, out, options);
        } catch (Exception ex) {
            ex.printStackTrace();
        }

    }
}
