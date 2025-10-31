package com.db.dbcover.service;

import com.db.dbcover.config.ExcelTemplateProperties;
import com.db.dbcover.template.DefaultExcelTemplates;
import com.db.dbcover.template.ExcelTemplateDefinition;
import com.db.dbcover.template.ExcelTemplateDefinition.Column;
import com.db.dbcover.template.ExcelTemplateDefinition.ColumnType;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.PaneInformation;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.StringJoiner;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class ExcelGeneratorServiceTest {

    private ExcelTemplateProperties properties;
    private ExcelGeneratorService service;

    @BeforeEach
    void setUp() {
        properties = DefaultExcelTemplates.properties();
        service = new ExcelGeneratorService(properties);
    }

    @Test
    @DisplayName("generates configured sheets with metadata, freeze panes, and filters")
    void shouldGenerateAllConfiguredSheetsWithMetadata() throws IOException {
        ExcelTemplateDefinition definition = properties.resolvedInstrumentTemplates().get("MORTGAGE");
        byte[] workbookBytes = service.generateTemplate(definition);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(workbookBytes))) {
            for (TemplateSheet sheetDefinition : definition.getSheets()) {
                Sheet sheet = workbook.getSheet(sheetDefinition.getName());
                assertThat(sheet).as("sheet %s", sheetDefinition.getName()).isNotNull();

                Row headerRow = sheet.getRow(0);
                Row infoRow = sheet.getRow(1);
                assertThat(headerRow).isNotNull();
                assertThat(infoRow).isNotNull();

                List<Column> expectedColumns = sheetDefinition.getColumns();
                for (int columnIndex = 0; columnIndex < expectedColumns.size(); columnIndex++) {
                    Column column = expectedColumns.get(columnIndex);
                    Cell headerCell = headerRow.getCell(columnIndex);
                    Cell infoCell = infoRow.getCell(columnIndex);

                    assertThat(headerCell.getStringCellValue()).isEqualTo(column.getHeader());
                    assertThat(infoCell.getStringCellValue()).isEqualTo(buildInfoValue(column));

                    Comment comment = headerCell.getCellComment();
                    String commentText = comment != null ? comment.getString().getString() : null;
                    String expectedTooltip = column.getTooltip();
                    if (expectedTooltip == null || expectedTooltip.isBlank()) {
                        expectedTooltip = buildInfoValue(column);
                    }
                    assertThat(commentText).isEqualTo(expectedTooltip);

                    CellStyle columnStyle = sheet.getColumnStyle(columnIndex);
                    assertThat(columnStyle).as("style for %s", column.getHeader()).isNotNull();
                    assertThat(columnStyle.getDataFormatString()).isEqualTo(column.resolvedFormat());
                }

                PaneInformation paneInformation = sheet.getPaneInformation();
                assertThat(paneInformation).isNotNull();
                assertThat(paneInformation.isFreezePane()).isTrue();
                assertThat(paneInformation.getHorizontalSplitPosition()).isEqualTo((short) 2);
            }
        }
    }

    @Test
    @DisplayName("applies list and numeric validations based on column types")
    void shouldApplyValidations() throws IOException {
        ExcelTemplateDefinition definition = new ExcelTemplateDefinition();
        TemplateSheet sheet = TemplateSheet.builder()
                .name("VALIDATIONS")
                .columns(List.of(
                        column("FLAG", ColumnType.BOOLEAN, true, null, null, null, null),
                        column("CHOICE", ColumnType.LIST, false, null, null, null, List.of("A", "B")),
                        column("WHEN", ColumnType.DATE, false, null, null, null, null),
                        column("AMOUNT", ColumnType.NUMBER, false, null, null, null, null)
                ))
                .build();
        definition.setSheets(List.of(sheet));

        byte[] workbookBytes = service.generateTemplate(definition);
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(workbookBytes))) {
            Sheet workbookSheet = workbook.getSheet("VALIDATIONS");
            assertThat(workbookSheet).isNotNull();
            List<? extends DataValidation> validations = workbookSheet.getDataValidations();
            assertThat(validations).hasSize(4);

            Map<String, Integer> expectedTypes = Map.of(
                    "FLAG", DataValidationConstraint.ValidationType.LIST,
                    "CHOICE", DataValidationConstraint.ValidationType.LIST,
                    "WHEN", DataValidationConstraint.ValidationType.DATE,
                    "AMOUNT", DataValidationConstraint.ValidationType.DECIMAL
            );

            for (DataValidation validation : validations) {
                int columnIndex = validation.getRegions().getCellRangeAddresses()[0].getFirstColumn();
                String header = workbookSheet.getRow(0).getCell(columnIndex).getStringCellValue();
                assertThat(validation.getValidationConstraint().getValidationType()).isEqualTo(expectedTypes.get(header));
            }
        }
    }

    @Test
    @DisplayName("throws when instrument type is missing")
    void shouldRequireInstrumentType() {
        assertThatThrownBy(() -> service.generateTemplate(" "))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("instrumentType must be provided");
    }

    @Test
    @DisplayName("throws when instrument type is unknown")
    void shouldRejectUnknownInstrumentType() {
        assertThatThrownBy(() -> service.generateTemplate("UNKNOWN"))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("Unknown instrument type");
    }

    private static Column column(String header,
                                 ColumnType type,
                                 boolean required,
                                 String description,
                                 String format,
                                 String tooltip,
                                 List<String> allowedValues) {
        Column column = Column.builder()
                .header(header)
                .type(type)
                .required(required)
                .description(description)
                .format(format)
                .tooltip(tooltip)
                .build();
        column.setAllowedValues(allowedValues);
        return column;
    }

    private static String buildInfoValue(Column column) {
        StringJoiner joiner = new StringJoiner(" | ");
        joiner.add("type: " + column.typeLabel());
        joiner.add("required: " + (column.isRequired() ? "yes" : "no"));
        if (column.getDescription() != null && !column.getDescription().isBlank()) {
            joiner.add(column.getDescription());
        }
        List<String> allowedValues = column.resolvedAllowedValues();
        if (!allowedValues.isEmpty()) {
            joiner.add("allowed: " + allowedValues);
        }
        String format = column.resolvedFormat();
        if (format != null && !format.isBlank()) {
            joiner.add("format: " + format);
        }
        return joiner.toString();
    }
}
