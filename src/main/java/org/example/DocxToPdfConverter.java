package org.example;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.concurrent.Executors;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class DocxToPdfConverter {

    public static void main(String[] args) {
        // Example DOCX file paths
        String[] docxFiles = {"D:\\(MPTC)\\untitled1111111111\\src\\main\\resources\\templates\\Leave_Request_Template.docx", "D:\\(MPTC)\\untitled1111111111\\src\\main\\resources\\templates\\ok_template.docx"};

        try (var executor = Executors.newVirtualThreadPerTaskExecutor()) {
            for (String docxFile : docxFiles) {
                executor.submit(() -> convertDocxToPdf(docxFile));
            }
        } // AutoCloseable will shut down the executor here
    }

    private static void convertDocxToPdf(String docxFilePath) {
        try {
            // Load DOCX file
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(docxFilePath));

            // Define output PDF file path
            String pdfFilePath = docxFilePath.replace(".docx", ".pdf");
            
            // Create output stream for PDF
            try (OutputStream os = new FileOutputStream(new File(pdfFilePath))) {
                // Convert DOCX to PDF
                Docx4J.toPDF(wordMLPackage, os);
            }

            System.out.println("Converted " + docxFilePath + " to " + pdfFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
