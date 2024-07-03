package org.example.test.jj;

import fr.opensagres.xdocreport.core.XDocReportException;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException, XDocReportException {
        final ExportData exportData = WordUtils.getExportData("src/main/resources/templates/Leave_Request_Template.docx");
        exportData.setData("name","លន់លៃមហែ");
        exportData.setData("username", "Hai");
        exportData.setData("phoneNumber", "012349929");
        exportData.setData("organization", "DGC");
        exportData.setData("department", "MPTC");
        exportData.setData("office", "floor1");
        exportData.setData("reason", "សូម");
        exportData.setData("age", "20");
        exportData.setData("position", "backend");
        exportData.addImage("img",new File("src/main/resources/images/lunlimhai.jpg"));
        exportData.process(new FileOutputStream("D:\\(MPTC)\\untitled1111111111\\src\\main\\resources\\templates\\out.docx"));

    }
}
