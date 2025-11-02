package com.db.dbcover.template;

import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;

import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

public final class TemplateSheetResolver {

    private TemplateSheetResolver() {
    }

    public static ResolvedTemplates resolve(List<TemplateSheet> sheets,
                                            Map<String, ExcelTemplateDefinition.TemplateSettings> instrumentTemplates) {
        if (instrumentTemplates == null || instrumentTemplates.isEmpty()) {
            throw new IllegalStateException("instrument-templates must not be empty");
        }

        Map<String, TemplateSheet> sheetIndex = resolveSheets(sheets);
        Map<String, ExcelTemplateDefinition> definitions = resolveInstrumentTemplates(sheetIndex, instrumentTemplates);
        return new ResolvedTemplates(sheetIndex, definitions);
    }

    private static Map<String, TemplateSheet> resolveSheets(List<TemplateSheet> sheets) {
        Map<String, TemplateSheet> source = new LinkedHashMap<>();
        if (sheets != null) {
            for (TemplateSheet sheet : sheets) {
                String name = sheet.getName();
                if (name == null || name.isBlank()) {
                    throw new IllegalArgumentException("template sheet name must not be blank");
                }
                source.put(name, sheet);
            }
        }

        Map<String, TemplateSheet> resolved = new LinkedHashMap<>();
        Set<String> stack = new LinkedHashSet<>();
        for (String name : source.keySet()) {
            resolveSheet(name, source, resolved, stack);
        }
        return resolved;
    }

    private static TemplateSheet resolveSheet(String name,
                                              Map<String, TemplateSheet> source,
                                              Map<String, TemplateSheet> resolved,
                                              Set<String> stack) {
        TemplateSheet cached = resolved.get(name);
        if (cached != null) {
            return cached;
        }

        TemplateSheet sheet = source.get(name);
        if (sheet == null) {
            throw new IllegalArgumentException("Unknown template sheet: " + name);
        }

        if (!stack.add(name)) {
            throw new IllegalArgumentException("Circular sheet reference: " + name);
        }

        LinkedHashMap<String, ExcelTemplateDefinition.Column> columns = new LinkedHashMap<>();
        for (String baseName : sheet.getBaseSheets()) {
            TemplateSheet base = resolveSheet(baseName, source, resolved, stack);
            for (ExcelTemplateDefinition.Column column : base.getColumns()) {
                columns.put(column.getHeader(), column);
            }
        }

        for (ExcelTemplateDefinition.Column column : sheet.getColumns()) {
            columns.put(column.getHeader(), column);
        }

        TemplateSheet merged = TemplateSheet.builder()
                .name(sheet.getName())
                .columns(columns.values().stream().toList())
                .build();
        resolved.put(name, merged);
        stack.remove(name);
        return merged;
    }

    private static Map<String, ExcelTemplateDefinition> resolveInstrumentTemplates(Map<String, TemplateSheet> sheetIndex,
                                                                                   Map<String, ExcelTemplateDefinition.TemplateSettings> templates) {
        Map<String, ExcelTemplateDefinition> definitions = new LinkedHashMap<>();
        templates.forEach((name, settings) ->
                definitions.put(name, ExcelTemplateDefinition.fromSettings(settings, sheetIndex)));
        return definitions;
    }

    public record ResolvedTemplates(Map<String, TemplateSheet> sheetIndex,
                                    Map<String, ExcelTemplateDefinition> instrumentTemplates) {
    }
}
