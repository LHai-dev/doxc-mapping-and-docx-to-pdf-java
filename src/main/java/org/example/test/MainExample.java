package org.example.test;

import fr.opensagres.xdocreport.converter.ConverterTypeTo;
import org.example.test.xDocReport.XDocReportUtil;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

public class MainExample {

    public static void main(String[] args) throws Exception {
        // Sample input parameters
        String templatePath = "src/main/resources/templates/test.docx"; // Replace with your actual template path
        String outputFilePath = "src/main/resources/templates/"; // Output directory
        String outputFileName = "output.pdf"; // Output file name
        String watermarkText = "CONFIDENTIAL";

        // Prepare parameters for the report template
        Map<String, Object> paramMap = new HashMap<>();
        paramMap.put("language", "John Doe");
        paramMap.put("type", 30);
        paramMap.put("title", "2024-07-03");

        // Create XDocReportUtil instance
        XDocReportUtil reportUtil = new XDocReportUtil(templatePath, paramMap);

        // Generate PDF report
        ByteArrayOutputStream pdfOutputStream = reportUtil.download(outputFileName, ConverterTypeTo.PDF, true, watermarkText);

        // Save the generated PDF to disk or do further processing
        saveToFile(pdfOutputStream, outputFilePath + "/" + outputFileName);
    }

    // Helper method to save ByteArrayOutputStream to a file
    private static void saveToFile(ByteArrayOutputStream outputStream, String filePath) throws Exception {
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            outputStream.writeTo(fos);
        }
    }
}
