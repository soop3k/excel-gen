package com.db.dbcover.service;

import com.db.dbcover.config.ExcelTemplateProperties;
import com.db.dbcover.template.ExcelTemplateDefinition;
import com.db.dbcover.template.ExcelTemplateDefinition.Column;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Optional;

@Service
@RequiredArgsConstructor
public class ExcelGeneratorService {

    private static final int INITIAL_DATA_ROWS = 10000;

    private static final int HEADER_ROW = 0;
    private static final int INFO_ROW = 1;

    private static final short HEADER_FILL_COLOR = IndexedColors.BLUE_GREY.getIndex();
    private static final short OPTIONAL_HEADER_FONT_COLOR = IndexedColors.WHITE.getIndex();
    private static final short REQUIRED_HEADER_FONT_COLOR = IndexedColors.RED.getIndex();

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

            for (TemplateSheet sheetDefinition : templateDefinition.getSheets()) {
                createSheet(workbook, sheetDefinition, dataFormat);
            }

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private void createSheet(Workbook workbook,
                             TemplateSheet sheetDefinition,
                             DataFormat dataFormat) {
        Sheet sheet = initializeSheet(workbook, sheetDefinition);
        int columnCount = processColumns(sheet, sheetDefinition, dataFormat);
        finalizeSheet(sheet, columnCount);
    }

    private Sheet initializeSheet(Workbook workbook, TemplateSheet sheetDefinition) {
        Sheet sheet = workbook.createSheet(sheetDefinition.getName());

        sheet.createRow(HEADER_ROW); // header row
        sheet.createRow(INFO_ROW); // info row

        // protect sheet to enable headers and info cell locking
        sheet.protectSheet(sheetDefinition.getName());

        return sheet;
    }

    private int processColumns(Sheet sheet,
                               TemplateSheet sheetDefinition,
                               DataFormat dataFormat) {
        int columnIndex = 0;
        for (Column column : sheetDefinition.getColumns()) {
            processColumn(sheet, columnIndex, column, dataFormat);
            columnIndex++;
        }
        return columnIndex;
    }

    private void processColumn(Sheet sheet,
                               int columnIndex,
                               Column column,
                               DataFormat dataFormat) {
        Row headerRow = sheet.getRow(HEADER_ROW);
        Row infoRow = sheet.getRow(INFO_ROW);
        
        createHeaderCell(headerRow, columnIndex, column);
        createInfoCell(infoRow, columnIndex, column);
        applyColumnFormat(sheet, columnIndex, column, sheet.getWorkbook(), dataFormat);
        applyColumnValidation(sheet, sheet.getDataValidationHelper(), columnIndex, column);
        applyColumnTooltip(sheet.getWorkbook().getCreationHelper(), sheet.createDrawingPatriarch(), columnIndex, 
                          headerRow.getCell(columnIndex), column);
        sheet.autoSizeColumn(columnIndex);
    }

    private void createHeaderCell(Row headerRow, int columnIndex, Column column) {
        Cell headerCell = headerRow.createCell(columnIndex);
        headerCell.setCellValue(column.getHeader());
        
        Workbook workbook = headerRow.getSheet().getWorkbook();
        CellStyle headerStyle = createHeaderStyle(workbook, false);
        CellStyle requiredHeaderStyle = createHeaderStyle(workbook, true);
        
        headerCell.setCellStyle(column.isRequired() ? requiredHeaderStyle : headerStyle);
    }

    private void createInfoCell(Row infoRow, int columnIndex, Column column) {
        Cell infoCell = infoRow.createCell(columnIndex);
        infoCell.setCellValue(buildInfoCellValue(column));
        
        // Apply text wrapping to prevent long descriptions from affecting column width
        Workbook workbook = infoRow.getSheet().getWorkbook();
        CellStyle infoStyle = workbook.createCellStyle();
        infoStyle.setWrapText(true);
        infoStyle.setLocked(true);
        infoStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        infoCell.setCellStyle(infoStyle);
    }

    private void finalizeSheet(Sheet sheet, int columnCount) {
        if (columnCount > 0) {
            sheet.setAutoFilter(new CellRangeAddress(HEADER_ROW, INFO_ROW, 0, columnCount - 1));
        }
        sheet.createFreezePane(0, 2);
    }

    private void applyColumnFormat(Sheet sheet,
                                   int columnIndex,
                                   Column column,
                                   Workbook workbook,
                                   DataFormat dataFormat) {
        String normalizedFormat = column.resolvedFormat();
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(dataFormat.getFormat(normalizedFormat));
        sheet.setDefaultColumnStyle(columnIndex, style);
    }

    private void applyColumnValidation(Sheet sheet,
                                       DataValidationHelper helper,
                                       int columnIndex,
                                       Column column) {
        DataValidationConstraint constraint;

        switch (column.resolvedType()) {
            case LIST, BOOLEAN -> {
                List<String> values = column.resolvedAllowedValues();
                if (values.isEmpty()) {
                    return;
                }
                constraint = helper.createExplicitListConstraint(values.toArray(String[]::new));
            }
            case DATE ->
                    constraint = helper.createDateConstraint(
                            DataValidationConstraint.OperatorType.BETWEEN,
                            "DATE(1900,1,1)",
                            "DATE(9999,12,31)",
                            null
                    );
            case NUMBER ->
                    constraint = helper.createNumericConstraint(
                            DataValidationConstraint.ValidationType.DECIMAL,
                            DataValidationConstraint.OperatorType.BETWEEN,
                            "-1E307",
                            "1E307"
                    );
            default -> {
                return;
            }
        }

        CellRangeAddressList addressList = new CellRangeAddressList(2, 2 + INITIAL_DATA_ROWS, columnIndex, columnIndex);
        DataValidation validation = helper.createValidation(constraint, addressList);
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
    }

    private void applyColumnTooltip(CreationHelper creationHelper,
                                    org.apache.poi.ss.usermodel.Drawing<?> drawing,
                                    int columnIndex,
                                    Cell headerCell,
                                    Column column) {
        String tooltip = resolveColumnTooltip(column);

        ClientAnchor anchor = creationHelper.createClientAnchor();
        anchor.setCol1(columnIndex);
        anchor.setCol2(columnIndex + 2);
        anchor.setRow1(0);
        anchor.setRow2(2);

        Comment comment = drawing.createCellComment(anchor);
        comment.setString(creationHelper.createRichTextString(tooltip));
        headerCell.setCellComment(comment);
    }

    private String buildInfoCellValue(Column column) {
        return Optional.ofNullable(column.getDescription())
                .filter(value -> !value.isBlank())
                .orElse("");
    }

    private String resolveColumnTooltip(Column column) {
        return Optional.ofNullable(column.getTooltip())
                .filter(value -> !value.isBlank())
                .orElseGet(() -> buildInfoCellValue(column));
    }

    private CellStyle createHeaderStyle(Workbook workbook, boolean required) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(HEADER_FILL_COLOR);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(required ? REQUIRED_HEADER_FONT_COLOR : OPTIONAL_HEADER_FONT_COLOR);
        style.setFont(font);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setLocked(true);

        return style;
    }

}
