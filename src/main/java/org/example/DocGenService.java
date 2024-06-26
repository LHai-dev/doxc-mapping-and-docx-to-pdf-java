package org.example;

//import com.fasterxml.jackson.databind.ObjectMapper;
//import ee.sm.ti.teis.commongateway.docgen.dto.GeneratedDocumentDto;
//import ee.sm.ti.teis.docgen.config.QueueConfig.NotifyDocumentGenerated;
//import ee.sm.ti.teis.domain.docgen.GenerateDocument;
//import ee.sm.ti.teis.domain.docgen.GeneratedDocument;
//import ee.sm.ti.teis.domain.file.FileMetadata;
//import ee.sm.ti.teis.errors.CommonErrorCode;
//import ee.sm.ti.teis.exceptions.TeisBusinessException;
//import ee.sm.ti.teis.exceptions.TeisIllegalArgumentException;
//import ee.sm.ti.teis.exceptions.TeisResourceNotFoundException;
//import ee.sm.ti.teis.fileclient.FileStorageServiceClient;
//import ee.sm.ti.teis.files.FileMetaDataService;
//import ee.sm.ti.teis.servicerequest.RequestMetaDTO;
//import fr.opensagres.odfdom.converter.pdf.PdfOptions;

import com.fasterxml.jackson.databind.ObjectMapper;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.xdocreport.converter.ConverterTypeTo;
import fr.opensagres.xdocreport.converter.Options;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;

import java.io.*;
import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;

//import static ee.sm.ti.teis.exceptions.ServiceRequestExceptionHandler.handleError;
//import static ee.sm.ti.teis.types.enums.ObjectStatus.CURRENT;


@Slf4j
@RequiredArgsConstructor
public class DocGenService {
   public static PdfOptions pdfOptions;
    public DocGenService(PdfOptions pdfOptions) {
        DocGenService.pdfOptions = pdfOptions;
    }

    public void generateDocument(String templatePath, String jsonData, String outputFormatCode, String docxOutputPath, String pdfOutputPath) {
        Map<String, Object> model = mapJsonData(jsonData);
        Optional<DocumentFormat> outputFormatOpt = DocumentFormat.fromString(outputFormatCode);
        DocumentFormat outputFormat = outputFormatOpt.orElseGet(DocumentFormat::getDefaultFormat);

        if (outputFormat == DocumentFormat.DOCX || outputFormat == DocumentFormat.PDF) {
            if (outputFormat == DocumentFormat.DOCX) {
                byte[] docxContent = generateContent(templatePath, model, DocumentFormat.DOCX);
                saveToFile(docxContent, docxOutputPath);
            } else {
                byte[] pdfContent = generateContent(templatePath, model, DocumentFormat.PDF);
                saveToFile(pdfContent, pdfOutputPath);
            }
        } else {
            throw new IllegalArgumentException("Unsupported document format: " + outputFormatCode);
        }
    }

    private Map<String, Object> mapJsonData(String jsonData) {
        HashMap<String, Object> model;
        try {
            model = new ObjectMapper().readValue(jsonData, HashMap.class);
        } catch (IOException e) {
            throw new RuntimeException("Template processing failed. Invalid data.", e);
        }
        return model;
    }

    private byte[] generateContent(String templatePath, Map<String, Object> model, DocumentFormat outputFormat) {
        try (InputStream templateStream = new FileInputStream(templatePath);
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            IXDocReport report = XDocReportRegistry.getRegistry().loadReport(templateStream, TemplateEngineKind.Freemarker);
            IContext context = report.createContext();
            context.put("data", model);

            if (DocumentFormat.DOCX == outputFormat) {
                report.process(context, out);
            } else if (DocumentFormat.PDF == outputFormat) {
                Options options = Options.getTo(ConverterTypeTo.PDF).subOptions(pdfOptions);
                report.convert(context, options, out);
            }
            return out.toByteArray();
        } catch (Exception e) {
            throw new RuntimeException("Template processing failed. " + e.getMessage(), e);
        }
    }

    private void saveToFile(byte[] content, String filePath) {
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            fos.write(content);
        } catch (IOException e) {
            throw new RuntimeException("Failed to save file to path: " + filePath, e);
        }
    }

    public static void main(String[] args) {
        // Define your template path, JSON data, and output file paths
        String templatePath = "src/main/resources/templates/Leave_Request_Template.docx";
        String jsonData = """
                {
                  "username": "Hai",
                  "phoneNumber": "012349929",
                  "organization": "DGC",
                  "department": "MPC",
                  "office": "floor1",
                  "age": "20",
                  "position": "backend"
                }
                """;
        String outputFormatCode = "DOCX"; // Change to "PDF" for PDF output
        String docxOutputPath = "src/main/resources/templates/document.docx";
        String pdfOutputPath = "src/main/resources/templates/document.pdf";

        // Initialize PdfOptions if you have any specific options (otherwise, this can be empty)

        PdfOptions pdfOptions = DocGenService.pdfOptions;
        // Initialize the DocumentGenerator
        DocGenService documentGenerator = new DocGenService(pdfOptions);

        // Generate the document
        documentGenerator.generateDocument(templatePath, jsonData, outputFormatCode, docxOutputPath, pdfOutputPath);
    }

}
