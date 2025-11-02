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
