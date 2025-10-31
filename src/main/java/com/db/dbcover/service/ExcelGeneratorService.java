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
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Service
@RequiredArgsConstructor
public class ExcelGeneratorService {

    private static final int SAMPLE_DATA_ROWS = 100;
    private static final short HEADER_FILL_COLOR = IndexedColors.GREY_25_PERCENT.getIndex();
    private static final short OPTIONAL_HEADER_FONT_COLOR = IndexedColors.BLACK.getIndex();
    private static final short REQUIRED_HEADER_FONT_COLOR = IndexedColors.RED.getIndex();
    private static final boolean HEADER_BOLD = true;

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
            CellStyle headerStyle = createHeaderStyle(workbook, false);
            CellStyle requiredHeaderStyle = createHeaderStyle(workbook, true);

            DataFormat dataFormat = workbook.createDataFormat();
            Map<String, CellStyle> formatStyles = new HashMap<>();

            for (TemplateSheet sheetDefinition : templateDefinition.getSheets()) {
                createSheet(workbook, sheetDefinition, headerStyle, requiredHeaderStyle, dataFormat, formatStyles);
            }

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private void createSheet(Workbook workbook,
                             TemplateSheet sheetDefinition,
                             CellStyle headerStyle,
                             CellStyle requiredHeaderStyle,
                             DataFormat dataFormat,
                             Map<String, CellStyle> formatStyles) {
        Sheet sheet = workbook.createSheet(sheetDefinition.getName());
        Row headerRow = sheet.createRow(0);
        Row infoRow = sheet.createRow(1);

        CreationHelper creationHelper = workbook.getCreationHelper();
        DataValidationHelper validationHelper = sheet.getDataValidationHelper();
        org.apache.poi.ss.usermodel.Drawing<?> drawing = sheet.createDrawingPatriarch();

        int columnIndex = 0;
        for (Column column : sheetDefinition.getColumns()) {
            Cell headerCell = headerRow.createCell(columnIndex);
            headerCell.setCellValue(column.getHeader());
            headerCell.setCellStyle(column.isRequired() ? requiredHeaderStyle : headerStyle);

            Cell infoCell = infoRow.createCell(columnIndex);
            infoCell.setCellValue(buildInfoCellValue(column));

            applyColumnFormat(sheet, columnIndex, column, workbook, dataFormat, formatStyles);
            applyColumnValidation(sheet, validationHelper, columnIndex, column);
            applyColumnTooltip(creationHelper, drawing, columnIndex, headerCell, column);

            sheet.setColumnWidth(columnIndex, 20 * 256);
            columnIndex++;
        }

        if (columnIndex > 0) {
            sheet.setAutoFilter(new CellRangeAddress(0, 1, 0, columnIndex - 1));
        }
        sheet.createFreezePane(0, 2);
    }

    private void applyColumnFormat(Sheet sheet,
                                   int columnIndex,
                                   Column column,
                                   Workbook workbook,
                                   DataFormat dataFormat,
                                   Map<String, CellStyle> formatStyles) {
        String normalizedFormat = column.resolvedFormat().trim();
        CellStyle style = formatStyles.computeIfAbsent(normalizedFormat, key -> {
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.setDataFormat(dataFormat.getFormat(key));
            return newStyle;
        });
        sheet.setDefaultColumnStyle(columnIndex, style);
    }

    private void applyColumnValidation(Sheet sheet,
                                       DataValidationHelper helper,
                                       int columnIndex,
                                       Column column) {
        DataValidationConstraint constraint;
        boolean suppressDropdown = true;

        switch (column.resolvedType()) {
            case LIST, BOOLEAN -> {
                List<String> values = column.resolvedAllowedValues();
                if (values.isEmpty()) {
                    return;
                }
                constraint = helper.createExplicitListConstraint(values.toArray(String[]::new));
                suppressDropdown = false;
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

        CellRangeAddressList addressList = new CellRangeAddressList(2, 2 + SAMPLE_DATA_ROWS, columnIndex, columnIndex);
        DataValidation validation = helper.createValidation(constraint, addressList);
        validation.setSuppressDropDownArrow(suppressDropdown);
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
        String description = Optional.ofNullable(column.getDescription())
                .filter(ExcelGeneratorService::hasText)
                .orElse(null);
        List<String> allowedValues = column.resolvedAllowedValues();

        return Stream.of(
                        "type: " + column.typeLabel(),
                        "required: " + (column.isRequired() ? "yes" : "no"),
                        description,
                        allowedValues.isEmpty() ? null : "allowed: " + allowedValues,
                        "format: " + column.resolvedFormat()
                )
                .filter(Objects::nonNull)
                .collect(Collectors.joining(" | "));
    }

    private String resolveColumnTooltip(Column column) {
        return Optional.ofNullable(column.getTooltip())
                .filter(ExcelGeneratorService::hasText)
                .orElseGet(() -> buildInfoCellValue(column));
    }

    private CellStyle createHeaderStyle(Workbook workbook, boolean required) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(HEADER_FILL_COLOR);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font font = workbook.createFont();
        font.setBold(HEADER_BOLD);
        font.setColor(required ? REQUIRED_HEADER_FONT_COLOR : OPTIONAL_HEADER_FONT_COLOR);
        style.setFont(font);
        return style;
    }

    private static boolean hasText(String value) {
        return value != null && !value.isBlank();
    }
}
