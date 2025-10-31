package com.db.dbcover.template;

import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

public final class TemplateSheetResolver {

    private TemplateSheetResolver() {
    }

    public static Map<String, TemplateSheet> resolveSheets(List<TemplateSheet> sheets) {
        Map<String, TemplateSheet> resolved = new LinkedHashMap<>();
        Map<String, TemplateSheet> source = new LinkedHashMap<>();
        if (sheets == null) {
            return resolved;
        }

        for (TemplateSheet sheet : sheets) {
            if (sheet.getName() == null || sheet.getName().isBlank()) {
                throw new IllegalArgumentException("template sheet name must not be blank");
            }
            source.put(sheet.getName(), sheet);
        }

        Set<String> visiting = new LinkedHashSet<>();
        for (String sheetName : source.keySet()) {
            resolveSheet(sheetName, source, resolved, visiting);
        }
        return resolved;
    }

    private static TemplateSheet resolveSheet(String sheetName,
                                              Map<String, TemplateSheet> source,
                                              Map<String, TemplateSheet> resolved,
                                              Set<String> visiting) {
        TemplateSheet cached = resolved.get(sheetName);
        if (cached != null) {
            return cached;
        }

        TemplateSheet sheet = source.get(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Unknown template sheet: " + sheetName);
        }

        if (!visiting.add(sheetName)) {
            throw new IllegalArgumentException("Circular sheet reference: " + sheetName);
        }

        Map<String, ExcelTemplateDefinition.Column> columns = new LinkedHashMap<>();
        List<String> baseSheets = sheet.getBaseSheets();
        if (baseSheets != null) {
            for (String baseName : baseSheets) {
                TemplateSheet baseSheet = resolveSheet(baseName, source, resolved, visiting);
                baseSheet.getColumns().forEach(column -> columns.put(column.getHeader(), column));
            }
        }

        List<ExcelTemplateDefinition.Column> currentColumns = sheet.getColumns();
        if (currentColumns != null) {
            currentColumns.forEach(column -> columns.put(column.getHeader(), column));
        }

        TemplateSheet resolvedSheet = TemplateSheet.builder()
                .name(sheet.getName())
                .columns(new ArrayList<>(columns.values()))
                .build();
        resolved.put(sheetName, resolvedSheet);
        visiting.remove(sheetName);
        return resolvedSheet;
    }

    public static Map<String, ExcelTemplateDefinition> resolveDefinitions(Map<String, TemplateSheet> sheetIndex,
                                                                          Map<String, ExcelTemplateDefinition.TemplateSettings> instrumentTemplates) {
        if (sheetIndex == null) {
            throw new IllegalArgumentException("sheetIndex must not be null");
        }
        if (instrumentTemplates == null || instrumentTemplates.isEmpty()) {
            throw new IllegalStateException("instrument-templates must not be empty");
        }

        Map<String, ExcelTemplateDefinition.TemplateSettings> mergedSettings = mergeInstrumentTemplates(instrumentTemplates);
        Map<String, ExcelTemplateDefinition> definitions = new LinkedHashMap<>();
        mergedSettings.forEach((name, settings) ->
                definitions.put(name, ExcelTemplateDefinition.fromSettings(settings, sheetIndex)));
        return Map.copyOf(definitions);
    }

    private static Map<String, ExcelTemplateDefinition.TemplateSettings> mergeInstrumentTemplates(Map<String, ExcelTemplateDefinition.TemplateSettings> source) {
        Map<String, ExcelTemplateDefinition.TemplateSettings> resolved = new LinkedHashMap<>();
        Map<String, Boolean> visiting = new LinkedHashMap<>();
        for (String templateName : source.keySet()) {
            resolveInstrumentTemplate(templateName, source, resolved, visiting);
        }
        return resolved;
    }

    private static ExcelTemplateDefinition.TemplateSettings resolveInstrumentTemplate(String templateName,
                                                                                      Map<String, ExcelTemplateDefinition.TemplateSettings> source,
                                                                                      Map<String, ExcelTemplateDefinition.TemplateSettings> resolved,
                                                                                      Map<String, Boolean> visiting) {
        ExcelTemplateDefinition.TemplateSettings cached = resolved.get(templateName);
        if (cached != null) {
            return cached;
        }

        ExcelTemplateDefinition.TemplateSettings template = source.get(templateName);
        if (template == null) {
            throw new IllegalArgumentException("Unknown instrument type: " + templateName);
        }

        if (Boolean.TRUE.equals(visiting.get(templateName))) {
            throw new IllegalArgumentException("Circular template reference: " + templateName);
        }

        visiting.put(templateName, true);

        ExcelTemplateDefinition.TemplateSettings merged = new ExcelTemplateDefinition.TemplateSettings();
        for (String baseName : Optional.ofNullable(template.getBaseTemplates()).orElse(List.of())) {
            ExcelTemplateDefinition.TemplateSettings base = resolveInstrumentTemplate(baseName, source, resolved, visiting);
            merged = mergeTemplateSettings(merged, base);
        }

        merged = mergeTemplateSettings(merged, template);

        visiting.put(templateName, false);
        resolved.put(templateName, merged);
        return merged;
    }

    private static ExcelTemplateDefinition.TemplateSettings mergeTemplateSettings(ExcelTemplateDefinition.TemplateSettings baseSettings,
                                                                                  ExcelTemplateDefinition.TemplateSettings overrideSettings) {
        LinkedHashSet<String> order = new LinkedHashSet<>();
        Optional.ofNullable(baseSettings.getSheets()).orElse(List.of()).forEach(order::add);
        Optional.ofNullable(overrideSettings.getSheets()).orElse(List.of()).forEach(sheetName -> {
            order.remove(sheetName);
            order.add(sheetName);
        });

        ExcelTemplateDefinition.TemplateSettings merged = new ExcelTemplateDefinition.TemplateSettings();
        merged.setSheets(new ArrayList<>(order));
        return merged;
    }
}
