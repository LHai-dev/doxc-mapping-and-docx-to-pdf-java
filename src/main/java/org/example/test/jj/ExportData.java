package org.example.test.jj;

import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.images.FileImageProvider;
import fr.opensagres.xdocreport.document.images.IImageProvider;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.formatter.FieldsMetadata;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

public class ExportData {
    private FieldsMetadata fieldsMetadata;

    private IXDocReport report;
    private IContext context;

    public ExportData(IXDocReport report, IContext context,FieldsMetadata fieldsMetadata) {
        this.report = report;
        this.context = context;
        this.fieldsMetadata = report.createFieldsMetadata();

    }
    public void addImage(String key, File file) {
        this.fieldsMetadata.addFieldAsImage(key);
        IImageProvider img = new FileImageProvider(file);
        this.context.put(key, img);
    }
    public void process(OutputStream out) throws IOException, XDocReportException {
        report.process(context, out);
    }

    public void setData(String key, Object value) {
        FieldsMetadata fieldsMetadata = report.getFieldsMetadata() != null ? report.getFieldsMetadata() : new FieldsMetadata();
        fieldsMetadata.addFieldAsList(key);
        context.put(key, value);
    }
}
