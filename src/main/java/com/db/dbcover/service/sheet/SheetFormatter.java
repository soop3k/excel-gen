package com.db.dbcover.service.sheet;

import com.db.dbcover.template.ExcelTemplateDefinition.Column;
import lombok.RequiredArgsConstructor;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.List;
import java.util.Optional;

@RequiredArgsConstructor
public class SheetFormatter {

    private final Workbook workbook;
    private final DataFormat dataFormat;
    private final int headerRowIndex;
    private final int initialDataRows;

    private CellStyle headerStyle;
    private CellStyle requiredHeaderStyle;

    public void applyColumnFormat(Sheet sheet, int columnIndex, Column column) {
        String normalizedFormat = column.resolvedFormat();
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(dataFormat.getFormat(normalizedFormat));
        style.setLocked(false);
        sheet.setDefaultColumnStyle(columnIndex, style);
    }

    public void applyColumnValidation(Sheet sheet, int columnIndex, Column column) {
        DataValidationConstraint constraint;

        DataValidationHelper helper = sheet.getDataValidationHelper();
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

        int firstRow = headerRowIndex + 1;
        CellRangeAddressList addressList = new CellRangeAddressList(
                firstRow,
                firstRow + initialDataRows,
                columnIndex,
                columnIndex
        );
        DataValidation validation = helper.createValidation(constraint, addressList);
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
    }

    public void applyColumnTooltip(Sheet sheet, int columnIndex, Column column) {
        String tooltip = resolveColumnTooltip(column);
        if (tooltip.isBlank()) {
            return;
        }

        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        DataValidationConstraint dvConstraint = dvHelper.createCustomConstraint("ISNUMBER(" + columnIndex + ")");

        CellRangeAddressList addressList = new CellRangeAddressList(
                headerRowIndex,
                headerRowIndex + initialDataRows,
                columnIndex,
                columnIndex
        );

        DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);
        validation.createPromptBox("", tooltip);
        validation.setShowPromptBox(true);

        sheet.addValidationData(validation);
    }

    public void applyHeaderStyle(Cell headerCell, boolean required) {
        CellStyle style = required ? getRequiredHeaderStyle() : getHeaderStyle();
        headerCell.setCellStyle(style);
    }

    public void finalizeSheet(Sheet sheet, int columnCount) {
        if (columnCount > 0) {
            sheet.setAutoFilter(new CellRangeAddress(
                    headerRowIndex, headerRowIndex,
                    0, columnCount - 1));
        }

        sheet.createFreezePane(headerRowIndex, 1);
        autoSizeWithFilterPadding(sheet, 0, columnCount - 1);
    }

    private CellStyle getHeaderStyle() {
        if (headerStyle == null) {
            headerStyle = createHeaderStyle(false);
        }
        return headerStyle;
    }

    private CellStyle getRequiredHeaderStyle() {
        if (requiredHeaderStyle == null) {
            requiredHeaderStyle = createHeaderStyle(true);
        }
        return requiredHeaderStyle;
    }

    private CellStyle createHeaderStyle(boolean required) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(required ? IndexedColors.LIGHT_YELLOW.getIndex() : IndexedColors.WHITE.getIndex());
        style.setFont(font);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setLocked(true);

        return style;
    }

    private void autoSizeWithFilterPadding(Sheet sheet, int firstCol, int lastCol) {
        final double FILTER_PADDING_PCT = 1.25;

        for (int col = firstCol; col <= lastCol; col++) {
            sheet.autoSizeColumn(col);
            int currentWidth = sheet.getColumnWidth(col);

            int newWidth = (int) (currentWidth * FILTER_PADDING_PCT);
            sheet.setColumnWidth(col, newWidth);
        }
    }

    private String resolveColumnTooltip(Column column) {
        return Optional.ofNullable(column.getTooltip())
                .filter(value -> !value.isBlank())
                .orElseGet(() -> buildInfoCellValue(column));
    }

    private String buildInfoCellValue(Column column) {
        return Optional.ofNullable(column.getDescription())
                .filter(value -> !value.isBlank())
                .orElse("");
    }
}
