package com.db.dbcover.service;

import com.db.dbcover.config.ExcelTemplateProperties;
import com.db.dbcover.service.sheet.SheetBuilder;
import com.db.dbcover.service.sheet.SheetFormatter;
import com.db.dbcover.template.ExcelTemplateDefinition;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;
import lombok.RequiredArgsConstructor;

import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@Service
@RequiredArgsConstructor
public class ExcelGeneratorService {

    private static final int INITIAL_DATA_ROWS = 10000;

    private static final int HEADER_ROW = 0;

    private final ExcelTemplateProperties properties;

    public byte[] generateTemplate(String instrumentType) throws IOException {
        if (instrumentType == null || instrumentType.isBlank()) {
            throw new IllegalArgumentException("instrumentType must be provided");
        }
        ExcelTemplateDefinition templateDefinition = properties.resolvedInstrumentTemplates().get(instrumentType);
        if (templateDefinition == null) {
            throw new IllegalArgumentException("Unknown instrument type: " + instrumentType);
        }
        return generateTemplate(templateDefinition);
    }

    public byte[] generateTemplate(ExcelTemplateDefinition templateDefinition) throws IOException {
        if (templateDefinition == null) {
            throw new IllegalArgumentException("templateDefinition must not be null");
        }
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            DataFormat dataFormat = workbook.createDataFormat();
            SheetFormatter sheetFormatter = new SheetFormatter(workbook, dataFormat, HEADER_ROW, INITIAL_DATA_ROWS);
            SheetBuilder sheetBuilder = new SheetBuilder(workbook, sheetFormatter, HEADER_ROW);
            for (TemplateSheet sheetDefinition : templateDefinition.getSheets()) {
                sheetBuilder.buildSheet(sheetDefinition);
            }

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

}
