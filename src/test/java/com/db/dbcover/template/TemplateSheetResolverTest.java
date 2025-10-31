package com.db.dbcover.template;

import com.db.dbcover.template.ExcelTemplateDefinition.Column;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSettings;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class TemplateSheetResolverTest {

    @Test
    @DisplayName("merges columns from multiple base sheets in declaration order")
    void shouldMergeColumnsFromBaseSheets() {
        TemplateSheet baseA = sheet("BASE_A", List.of(), List.of(
                column("A1"),
                column("A2")
        ));
        TemplateSheet baseB = sheet("BASE_B", List.of(), List.of(
                column("B1")
        ));
        TemplateSheet derived = sheet("DERIVED", List.of("BASE_A", "BASE_B"), List.of(
                column("C1")
        ));

        TemplateSettings combinedSettings = new TemplateSettings();
        combinedSettings.setSheets(List.of("DERIVED"));

        TemplateSheetResolver.ResolvedTemplates resolved = TemplateSheetResolver.resolve(
                List.of(baseA, baseB, derived),
                Map.of("COMBINED", combinedSettings)
        );

        TemplateSheet merged = resolved.sheetIndex().get("DERIVED");
        assertThat(merged.getColumns())
                .extracting(Column::getHeader)
                .containsExactly("A1", "A2", "B1", "C1");
    }

    @Test
    @DisplayName("combines sheets from multiple base templates without duplicates")
    void shouldMergeSheetsFromBaseTemplates() {
        TemplateSheet sheetA = sheet("SHEET_A", List.of(), List.of(column("A")));
        TemplateSheet sheetB = sheet("SHEET_B", List.of(), List.of(column("B")));
        TemplateSheet sheetC = sheet("SHEET_C", List.of(), List.of(column("C")));

        TemplateSettings baseOne = new TemplateSettings();
        baseOne.setSheets(List.of("SHEET_A"));

        TemplateSettings baseTwo = new TemplateSettings();
        baseTwo.setSheets(List.of("SHEET_B"));

        TemplateSettings combined = new TemplateSettings();
        combined.setBaseTemplates(List.of("BASE_ONE", "BASE_TWO"));
        combined.setSheets(List.of("SHEET_C", "SHEET_A"));

        Map<String, TemplateSettings> templates = new LinkedHashMap<>();
        templates.put("BASE_ONE", baseOne);
        templates.put("BASE_TWO", baseTwo);
        templates.put("COMBINED", combined);

        TemplateSheetResolver.ResolvedTemplates resolved = TemplateSheetResolver.resolve(
                List.of(sheetA, sheetB, sheetC),
                templates
        );

        ExcelTemplateDefinition definition = resolved.instrumentTemplates().get("COMBINED");
        assertThat(definition.getSheets())
                .extracting(TemplateSheet::getName)
                .containsExactly("SHEET_A", "SHEET_B", "SHEET_C");
    }

    private static TemplateSheet sheet(String name, List<String> baseSheets, List<Column> columns) {
        return TemplateSheet.builder()
                .name(name)
                .baseSheets(baseSheets)
                .columns(columns)
                .build();
    }

    private static Column column(String header) {
        return Column.builder()
                .header(header)
                .build();
    }
}
