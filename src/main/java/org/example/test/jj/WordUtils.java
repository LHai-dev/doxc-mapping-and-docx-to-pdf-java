package org.example.test.jj;

import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import fr.opensagres.xdocreport.template.formatter.FieldsMetadata;

import java.io.FileInputStream;
import java.io.IOException;

public class WordUtils {

    public static ExportData getExportData(String url) throws IOException, XDocReportException {
        final IXDocReport report = createReport(url);
        final IContext context = report.createContext();
        final FieldsMetadata fieldsMetadata = report.createFieldsMetadata();
        return new ExportData(report, context,fieldsMetadata);
    }

    private static IXDocReport createReport(String url) throws IOException, XDocReportException {
        FileInputStream fileInputStream = new FileInputStream(url);
        return XDocReportRegistry.getRegistry().loadReport(fileInputStream, TemplateEngineKind.Freemarker);
    }
}
