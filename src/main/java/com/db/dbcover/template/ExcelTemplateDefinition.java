package com.db.dbcover.template;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Optional;

@Getter
public class ExcelTemplateDefinition {

    private List<TemplateSheet> sheets = new ArrayList<>();

    public List<TemplateSheet> getSheets() {
        return Collections.unmodifiableList(sheets);
    }

    public void setSheets(List<TemplateSheet> sheets) {
        this.sheets = sheets == null ? new ArrayList<>() : new ArrayList<>(sheets);
    }

    public static ExcelTemplateDefinition fromSettings(TemplateSettings settings,
                                                       java.util.Map<String, TemplateSheet> sheetIndex) {
        if (settings == null) {
            throw new IllegalArgumentException("settings must not be null");
        }
        if (sheetIndex == null) {
            throw new IllegalArgumentException("sheetIndex must not be null");
        }

        ExcelTemplateDefinition definition = new ExcelTemplateDefinition();
        List<TemplateSheet> resolvedSheets = new ArrayList<>();
        for (String sheetName : Optional.ofNullable(settings.getSheets()).orElse(List.of())) {
            TemplateSheet templateSheet = sheetIndex.get(sheetName);
            if (templateSheet == null) {
                throw new IllegalArgumentException("Unknown template sheet: " + sheetName);
            }
            resolvedSheets.add(templateSheet);
        }
        definition.setSheets(resolvedSheets);
        return definition;
    }

    @Getter
    @Setter
    @Builder
    public static class TemplateSheet {
        private String name;
        @Builder.Default
        private List<String> baseSheets = new ArrayList<>();
        @Builder.Default
        private List<Column> columns = new ArrayList<>();

        public List<String> getBaseSheets() {
            return Collections.unmodifiableList(baseSheets);
        }

        public List<Column> getColumns() {
            return Collections.unmodifiableList(columns);
        }

        public void setBaseSheets(List<String> baseSheets) {
            this.baseSheets = baseSheets == null ? new ArrayList<>() : new ArrayList<>(baseSheets);
            this.baseSheets.removeIf(name -> name == null || name.isBlank());
        }

        public void setColumns(List<Column> columns) {
            this.columns = columns == null ? new ArrayList<>() : new ArrayList<>(columns);
        }
    }

    @Getter
    @Setter
    @Builder
    public static class Column {
        private String header;
        private boolean required;
        private String description;
        private String format;
        private String tooltip;
        private ColumnType type;
        @Builder.Default
        private List<String> allowedValues = new ArrayList<>();

        public List<String> getAllowedValues() {
            return Collections.unmodifiableList(allowedValues);
        }

        public void setAllowedValues(List<String> allowedValues) {
            this.allowedValues = new ArrayList<>();
            if (allowedValues != null) {
                for (String value : allowedValues) {
                    if (value != null && !value.isBlank()) {
                        this.allowedValues.add(value);
                    }
                }
            }
        }

        public ColumnType resolvedType() {
            return Optional.ofNullable(type).orElse(ColumnType.TEXT);
        }

        public String resolvedFormat() {
            String explicit = this.format;
            if (explicit != null && !explicit.isBlank()) {
                return explicit;
            }
            return resolvedType().defaultFormat();
        }

        public List<String> resolvedAllowedValues() {
            ColumnType resolvedType = resolvedType();
            if (resolvedType == ColumnType.BOOLEAN && (allowedValues == null || allowedValues.isEmpty())) {
                return List.of("YES", "NO");
            }
            return allowedValues == null ? List.of() : List.copyOf(allowedValues);
        }

        public String typeLabel() {
            return resolvedType().label();
        }
    }

    public enum ColumnType {
        TEXT("TEXT", "@"),
        NUMBER("NUMBER", "#,##0.00############"),
        DATE("DATE", "dd/mm/yyyy"),
        LIST("LIST", "@"),
        BOOLEAN("BOOLEAN", "@");

        private final String label;
        private final String defaultFormat;

        ColumnType(String label, String defaultFormat) {
            this.label = label;
            this.defaultFormat = defaultFormat;
        }

        public String label() {
            return label;
        }

        public String defaultFormat() {
            return defaultFormat;
        }
    }

    @Getter
    @Setter
    public static class TemplateSettings {
        private List<String> sheets = new ArrayList<>();
        private List<String> baseTemplates = new ArrayList<>();

        public void setSheets(List<String> sheets) {
            this.sheets = sheets == null ? new ArrayList<>() : new ArrayList<>(sheets);
            this.sheets.removeIf(name -> name == null || name.isBlank());
        }

        public void setBaseTemplates(List<String> baseTemplates) {
            this.baseTemplates = baseTemplates == null ? new ArrayList<>() : new ArrayList<>(baseTemplates);
            this.baseTemplates.removeIf(name -> name == null || name.isBlank());
        }
    }
}
