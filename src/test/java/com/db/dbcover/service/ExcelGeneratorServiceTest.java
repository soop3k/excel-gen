package com.db.dbcover.service;

import com.db.dbcover.config.ExcelTemplateProperties;
import com.db.dbcover.template.DefaultExcelTemplates;
import com.db.dbcover.template.ExcelTemplateDefinition;
import com.db.dbcover.template.ExcelTemplateDefinition.Column;
import com.db.dbcover.template.ExcelTemplateDefinition.ColumnType;
import com.db.dbcover.template.ExcelTemplateDefinition.RequiredStatus;
import static com.db.dbcover.template.ExcelTemplateDefinition.RequiredStatus.REQUIRED;
import static com.db.dbcover.template.ExcelTemplateDefinition.RequiredStatus.NOT_REQUIRED;
import static com.db.dbcover.template.ExcelTemplateDefinition.ColumnType.BOOLEAN;
import static com.db.dbcover.template.ExcelTemplateDefinition.ColumnType.LIST;
import static com.db.dbcover.template.ExcelTemplateDefinition.ColumnType.DATE;
import static com.db.dbcover.template.ExcelTemplateDefinition.ColumnType.NUMBER;
import static com.db.dbcover.template.ExcelTemplateDefinition.ColumnType.TEXT;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;
import org.apache.poi.ss.usermodel.BorderFormatting;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class ExcelGeneratorServiceTest {

    private ExcelTemplateProperties properties;
    private ExcelGeneratorService service;

    @BeforeEach
    void setUp() {
        properties = DefaultExcelTemplates.properties();
        service = new ExcelGeneratorService(properties);
    }

    @Test
    void shouldGenerateAllConfiguredSheetsWithMetadata() throws IOException {
        ExcelTemplateDefinition definition = properties.resolvedInstrumentTemplates().get("MORTGAGE");
        byte[] workbookBytes = service.generateTemplate(definition);

        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(workbookBytes))) {
            org.apache.poi.ss.usermodel.DataFormat dataFormat = workbook.createDataFormat();
            for (TemplateSheet sheetDefinition : definition.getSheets()) {
                Sheet sheet = workbook.getSheet(sheetDefinition.getName());
                assertThat(sheet).as("sheet %s", sheetDefinition.getName()).isNotNull();

                Row headerRow = sheet.getRow(0);
                assertThat(headerRow).isNotNull();

                List<Column> expectedColumns = sheetDefinition.getColumns();
                for (int columnIndex = 0; columnIndex < expectedColumns.size(); columnIndex++) {
                    Column column = expectedColumns.get(columnIndex);
                    Cell headerCell = headerRow.getCell(columnIndex);

                    assertThat(headerCell.getStringCellValue()).isEqualTo(column.getHeader());

                    final int finalColumnIndex = columnIndex;
                    String tooltip = sheet.getDataValidations().stream().filter(
                            v -> v.getValidationConstraint().getFormula1().contains("ISNUMBER(" + finalColumnIndex +")"
                    )).findFirst().map(DataValidation::getPromptBoxText).orElse("");

                    String expectedTooltip = tooltip;
                    if (expectedTooltip.isBlank()) {
                        expectedTooltip = buildInfoValue(column);
                    }
                    assertThat(tooltip).isEqualTo(expectedTooltip);

                    CellStyle columnStyle = sheet.getColumnStyle(columnIndex);
                    assertThat(columnStyle).as("style for %s", column.getHeader()).isNotNull();
                    short expectedFormatIndex = dataFormat.getFormat(column.resolvedFormat());
                    assertThat(columnStyle.getDataFormat()).isEqualTo(expectedFormatIndex);
                }

                PaneInformation paneInformation = sheet.getPaneInformation();
                assertThat(paneInformation).isNotNull();
                assertThat(paneInformation.isFreezePane()).isTrue();
            }
        }
    }

    @Test
    void shouldApplyValidations() throws IOException {
        ExcelTemplateDefinition definition = new ExcelTemplateDefinition();
        TemplateSheet sheet = TemplateSheet.builder()
                .name("VALIDATIONS")
                .columns(List.of(
                        column("FLAG", BOOLEAN, REQUIRED, null, null, null, null),
                        column("CHOICE", LIST, NOT_REQUIRED, null, null, null, List.of("A", "B")),
                        column("WHEN", DATE, NOT_REQUIRED, null, null, null, null),
                        column("AMOUNT", NUMBER, NOT_REQUIRED, null, null, null, null)
                ))
                .build();
        definition.setSheets(List.of(sheet));

        byte[] workbookBytes = service.generateTemplate(definition);
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(workbookBytes))) {
            Sheet workbookSheet = workbook.getSheet("VALIDATIONS");
            assertThat(workbookSheet).isNotNull();
            List<? extends DataValidation> validations = workbookSheet.getDataValidations().stream().filter(
                    v -> !v.getValidationConstraint().getFormula1().contains("ISNUMBER")
            ).toList();
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
    void shouldHighlightEmptyRequiredCells() throws Exception {
        ExcelTemplateDefinition definition = new ExcelTemplateDefinition();
        TemplateSheet sheet = TemplateSheet.builder()
                .name("REQUIRED_FORMATTING")
                .columns(List.of(
                        column("MANDATORY", TEXT, REQUIRED, null, null, null, null),
                        column("OPTIONAL", TEXT, NOT_REQUIRED, null, null, null, null)
                ))
                .build();
        definition.setSheets(List.of(sheet));

        byte[] workbookBytes = service.generateTemplate(definition);
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(workbookBytes))) {
            Sheet workbookSheet = workbook.getSheet("REQUIRED_FORMATTING");
            assertThat(workbookSheet).isNotNull();

            SheetConditionalFormatting conditionalFormatting = workbookSheet.getSheetConditionalFormatting();
            assertThat(conditionalFormatting.getNumConditionalFormattings()).isEqualTo(1);

            ConditionalFormatting formatting = conditionalFormatting.getConditionalFormattingAt(0);
            assertThat(formatting.getNumberOfRules()).isEqualTo(1);

            CellRangeAddress range = formatting.getFormattingRanges()[0];
            assertThat(range.getFirstColumn()).isEqualTo(0);
            assertThat(range.getLastColumn()).isEqualTo(0);

            int firstRow = range.getFirstRow();
            assertThat(firstRow).isEqualTo(1);
            assertThat(range.getLastRow()).isEqualTo(firstRow + initialDataRows());

            ConditionalFormattingRule rule = formatting.getRule(0);
            assertThat(rule.getFormula1()).isEqualTo("LEN(TRIM($A2))=0");

            BorderFormatting borderFormatting = rule.getBorderFormatting();
            assertThat(borderFormatting).isNotNull();
            assertThat(borderFormatting.getBorderTop()).isEqualTo(BorderStyle.THIN);
            assertThat(borderFormatting.getBorderBottom()).isEqualTo(BorderStyle.THIN);
            assertThat(borderFormatting.getBorderLeft()).isEqualTo(BorderStyle.THIN);
            assertThat(borderFormatting.getBorderRight()).isEqualTo(BorderStyle.THIN);

            short redIndex = IndexedColors.RED.getIndex();
            assertThat(borderFormatting.getTopBorderColor()).isEqualTo(redIndex);
            assertThat(borderFormatting.getBottomBorderColor()).isEqualTo(redIndex);
            assertThat(borderFormatting.getLeftBorderColor()).isEqualTo(redIndex);
            assertThat(borderFormatting.getRightBorderColor()).isEqualTo(redIndex);
        }
    }

    @Test
    void shouldRequireInstrumentType() {
        assertThatThrownBy(() -> service.generateTemplate(" "))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("instrumentType must be provided");
    }

    @Test
    void shouldRejectUnknownInstrumentType() {
        assertThatThrownBy(() -> service.generateTemplate("UNKNOWN"))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("Unknown instrument type");
    }

    @Test
    void shouldDelegateValidationCreation() throws Exception {
        Method method = ExcelGeneratorService.class.getDeclaredMethod(
                "applyColumnValidation",
                Sheet.class,
                DataValidationHelper.class,
                int.class,
                Column.class
        );
        method.setAccessible(true);

        Sheet sheet = mock(Sheet.class);
        DataValidationHelper helper = mock(DataValidationHelper.class);
        DataValidationConstraint constraint = mock(DataValidationConstraint.class);
        DataValidation validation = mock(DataValidation.class);
        when(helper.createExplicitListConstraint(any(String[].class))).thenReturn(constraint);
        when(helper.createValidation(eq(constraint), any(CellRangeAddressList.class))).thenReturn(validation);

        Column listColumn = Column.builder()
                .header("CHOICE")
                .type(LIST)
                .build();
        listColumn.setAllowedValues(List.of("A", "B"));

        method.invoke(service, sheet, helper, 0, listColumn);

        verify(helper).createExplicitListConstraint(any(String[].class));
        verify(helper).createValidation(eq(constraint), any(CellRangeAddressList.class));
        verify(sheet).addValidationData(validation);
    }

    @Test
    void shouldSkipValidationForUnconstrainedTypes() throws Exception {
        Method method = ExcelGeneratorService.class.getDeclaredMethod(
                "applyColumnValidation",
                Sheet.class,
                DataValidationHelper.class,
                int.class,
                Column.class
        );
        method.setAccessible(true);

        Sheet sheet = mock(Sheet.class);
        DataValidationHelper helper = mock(DataValidationHelper.class);

        Column textColumn = Column.builder()
                .header("FREE_TEXT")
                .type(TEXT)
                .build();

        method.invoke(service, sheet, helper, 0, textColumn);

        verify(helper, never()).createValidation(any(), any());
        verify(sheet, never()).addValidationData(any());
    }

    private int initialDataRows() throws Exception {
        Field field = ExcelGeneratorService.class.getDeclaredField("INITIAL_DATA_ROWS");
        field.setAccessible(true);
        return field.getInt(null);
    }

    private static Column column(String header,
                                 ColumnType type,
                                 RequiredStatus requiredStatus,
                                 String description,
                                 String format,
                                 String tooltip,
                                 List<String> allowedValues) {
        Column column = Column.builder()
                .header(header)
                .type(type)
                .required(requiredStatus)
                .description(description)
                .format(format)
                .tooltip(tooltip)
                .build();
        column.setAllowedValues(allowedValues);
        return column;
    }

    private static String buildInfoValue(Column column) {
        String description = column.getDescription();
        return (description != null && !description.trim().isEmpty()) ? description : "";
    }
}
