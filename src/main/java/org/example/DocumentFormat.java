package org.example;

import lombok.Getter;

import java.util.Optional;

import static java.util.Optional.empty;
import static java.util.Optional.of;

@Getter
public enum DocumentFormat {
    PDF("application/pdf", "pdf"),
    DOCX("application/vnd.openxmlformats-officedocument.wordprocessingml.document", "docx"),
    XHTML("application/xhtml+xml", "html");

    private final String contentType;
    private final String fileExtension;

    DocumentFormat(String contentType, String fileExtension) {
        this.contentType = contentType;
        this.fileExtension = fileExtension;
    }

    public static Optional<DocumentFormat> fromString(String code) {
        if (code == null) {
            return empty();
        }
        for (DocumentFormat df : DocumentFormat.values()) {
            if (df.name().equalsIgnoreCase(code.toUpperCase())) {
                return of(df);
            }
        }
        return empty();
    }

    public static DocumentFormat getDefaultFormat() {
        return PDF;
    }

}
