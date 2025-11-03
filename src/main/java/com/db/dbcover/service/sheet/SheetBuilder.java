package com.db.dbcover.service.sheet;

import com.db.dbcover.template.ExcelTemplateDefinition.Column;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class SheetBuilder {

    private final Workbook workbook;
    private final SheetFormatter formatter;
    private final int headerRowIndex;

    public SheetBuilder(Workbook workbook, SheetFormatter formatter, int headerRowIndex) {
        this.workbook = workbook;
        this.formatter = formatter;
        this.headerRowIndex = headerRowIndex;
    }

    public void buildSheet(TemplateSheet sheetDefinition) {
        Sheet sheet = initializeSheet(sheetDefinition);
        int columnCount = populateColumns(sheet, sheetDefinition);
        formatter.finalizeSheet(sheet, columnCount);
    }

    private Sheet initializeSheet(TemplateSheet sheetDefinition) {
        Sheet sheet = workbook.createSheet(sheetDefinition.getName());
        sheet.createRow(headerRowIndex);
        sheet.protectSheet(sheetDefinition.getName());
        return sheet;
    }

    private int populateColumns(Sheet sheet, TemplateSheet sheetDefinition) {
        int columnIndex = 0;
        for (Column column : sheetDefinition.getColumns()) {
            processColumn(sheet, columnIndex, column);
            columnIndex++;
        }
        return columnIndex;
    }

    private void processColumn(Sheet sheet, int columnIndex, Column column) {
        formatter.applyColumnFormat(sheet, columnIndex, column);
        formatter.applyColumnValidation(sheet, columnIndex, column);
        formatter.applyColumnTooltip(sheet, columnIndex, column);

        Row headerRow = sheet.getRow(headerRowIndex);
        Cell headerCell = headerRow.createCell(columnIndex);
        headerCell.setCellValue(column.getHeader());
        formatter.applyHeaderStyle(headerCell, column.isRequired());
    }
}
