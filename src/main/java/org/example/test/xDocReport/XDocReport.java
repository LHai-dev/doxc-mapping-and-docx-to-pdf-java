/**
 * 
 */
package org.example.test.xDocReport;

import com.google.common.io.Files;
import fr.opensagres.xdocreport.converter.ConverterTypeTo;
import fr.opensagres.xdocreport.converter.Options;
import fr.opensagres.xdocreport.core.document.DocumentKind;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import lombok.Setter;
import org.springframework.core.io.FileSystemResource;

import java.io.*;
import java.util.Map;

/**
 * @author YingHan on 2021.
 */
public abstract class XDocReport {

    @Setter
    protected String templatePath;

    @Setter
    protected Map<String, Object> paramMap;

    public ByteArrayOutputStream build(String outputFilePath, String outputFileName, ConverterTypeTo typeto, String watermark) throws Exception {
    	FileSystemResource resource = new FileSystemResource(templatePath);
        InputStream in = new FileInputStream(resource.getFile());
        //        InputStream in = classLoader.getResourceAsStream(isTemplate);

        // 1) Load ODT file and set Freemarker template engine and cache it to the registry
        IXDocReport report = XDocReportRegistry.getRegistry().loadReport(in, TemplateEngineKind.Freemarker);

        // 2) Create Java model context
        IContext context = report.createContext();
        context.putMap(paramMap);

        File outputFullPath = new File(outputFilePath + "/" + outputFileName);
        outputFullPath.getParentFile().mkdirs();
        OutputStream out = new FileOutputStream(outputFullPath);

        // 3) Generate report by merging Java model with the DOCX
        if (typeto == null) {
            report.process(context, out);
        } else {
            final Options options = Options.getTo(typeto).from(DocumentKind.DOCX);
            report.convert(context, options, out);
        }

        ByteArrayOutputStream byteOut = new ByteArrayOutputStream();
        byteOut.write(Files.toByteArray(outputFullPath));
        byteOut.close();
        in.close();
        out.close();
        if (watermark != null) {
            return addWatermark(outputFullPath.getAbsolutePath(), outputFullPath.getAbsolutePath(), watermark);
        } else {
            return byteOut;
        }
    }

    public abstract ByteArrayOutputStream addWatermark(String inputFilePath, String outputFilePath, String watermark);

}
