package com.db.dbcover.service;

import com.db.dbcover.config.ExcelTemplateProperties;
import com.db.dbcover.template.ExcelTemplateDefinition;
import com.db.dbcover.template.ExcelTemplateDefinition.Column;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;
import lombok.RequiredArgsConstructor;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
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

    private static final short HEADER_FILL_COLOR = IndexedColors.GREY_25_PERCENT.getIndex();
    private static final short OPTIONAL_HEADER_FONT_COLOR = IndexedColors.WHITE.getIndex();
    private static final short REQUIRED_HEADER_FONT_COLOR = IndexedColors.LIGHT_YELLOW.getIndex();

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


        applyColumnFormat(sheet, columnIndex, column, sheet.getWorkbook(), dataFormat);
        applyColumnValidation(sheet, sheet.getDataValidationHelper(), columnIndex, column);
        applyColumnTooltip(sheet, columnIndex, column);

        Row headerRow = sheet.getRow(HEADER_ROW);
        createHeaderCell(headerRow, columnIndex, column);
    }

    private void createHeaderCell(Row headerRow, int columnIndex, Column column) {
        Cell headerCell = headerRow.createCell(columnIndex);
        headerCell.setCellValue(column.getHeader());
        
        Workbook workbook = headerRow.getSheet().getWorkbook();
        CellStyle headerStyle = createHeaderStyle(workbook, false);
        CellStyle requiredHeaderStyle = createHeaderStyle(workbook, true);
        
        headerCell.setCellStyle(column.isRequired() ? requiredHeaderStyle : headerStyle);
    }


    private void finalizeSheet(Sheet sheet, int columnCount) {
        if (columnCount > 0) {
            sheet.setAutoFilter(new CellRangeAddress(
                    HEADER_ROW, HEADER_ROW,
                    0, columnCount - 1));
        }

        sheet.createFreezePane(HEADER_ROW, 1);
        autoSizeWithFilterPadding(sheet, 0, columnCount - 1);
    }

    private void applyColumnFormat(Sheet sheet,
                                   int columnIndex,
                                   Column column,
                                   Workbook workbook,
                                   DataFormat dataFormat) {
        String normalizedFormat = column.resolvedFormat();
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(dataFormat.getFormat(normalizedFormat));
        style.setLocked(false);
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

        int firstRow = HEADER_ROW + 1;
        CellRangeAddressList addressList = new CellRangeAddressList( firstRow, firstRow + INITIAL_DATA_ROWS, columnIndex, columnIndex);
        DataValidation validation = helper.createValidation(constraint, addressList);
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
    }

    private void applyColumnTooltip(Sheet sheet,
                                    int columnIndex,
                                    Column column) {
        String tooltip = resolveColumnTooltip(column);

        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        DataValidationConstraint dvConstraint = dvHelper.createCustomConstraint("ISNUMBER(" + columnIndex + ")");

        CellRangeAddressList addressList = new CellRangeAddressList(HEADER_ROW, INITIAL_DATA_ROWS, columnIndex, columnIndex);

        DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);
        validation.createPromptBox("", tooltip);
        validation.setShowPromptBox(true);

        sheet.addValidationData(validation);
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
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setLocked(true);

        return style;
    }

    private  void autoSizeWithFilterPadding(Sheet sheet, int firstCol, int lastCol) {
        final double FILTER_PADDING_PCT = 1.25;

        for (int col = firstCol; col <= lastCol; col++) {
            sheet.autoSizeColumn(col);
            int currentWidth = sheet.getColumnWidth(col);

            int newWidth = (int) (currentWidth * FILTER_PADDING_PCT);
            sheet.setColumnWidth(col, newWidth);
        }
    }

}
