package com.db.dbcover.template;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellType;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

import static com.db.dbcover.template.ExcelTemplateDefinition.RequiredStatus.REQUIRED;
import static com.db.dbcover.template.ExcelTemplateDefinition.RequiredStatus.NOT_REQUIRED;

@Getter
public class ExcelTemplateDefinition {

    private List<TemplateSheet> sheets = new ArrayList<>();

    public List<TemplateSheet> getSheets() {
        return Collections.unmodifiableList(sheets);
    }

    public void setSheets(List<TemplateSheet> sheets) {
        this.sheets = Optional.ofNullable(sheets).orElseGet(List::of);
    }

    public static ExcelTemplateDefinition fromSettings(TemplateSettings settings,
                                                       java.util.Map<String, TemplateSheet> sheetIndex) {
        ExcelTemplateDefinition definition = new ExcelTemplateDefinition();
        List<TemplateSheet> resolvedSheets = Optional.ofNullable(settings.getSheets())
                .orElseGet(List::of)
                .stream()
                .map(sheetName -> Optional.ofNullable(sheetIndex.get(sheetName))
                        .orElseThrow(() -> new IllegalArgumentException("Unknown template sheet: " + sheetName)))
                .toList();
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
            return Collections.unmodifiableList(baseSheets != null ? baseSheets : List.of());
        }

        public List<Column> getColumns() {
            return Collections.unmodifiableList(columns != null ? columns : List.of());
        }

        public void setBaseSheets(List<String> baseSheets) {
            this.baseSheets = Optional.ofNullable(baseSheets)
                    .orElseGet(List::of)
                    .stream()
                    .filter(ExcelTemplateDefinition::hasText)
                    .collect(Collectors.toCollection(ArrayList::new));
        }

        public void setColumns(List<Column> columns) {
            this.columns = Optional.ofNullable(columns).orElseGet(List::of);
        }
    }

    @Getter
    @Setter
    @Builder
    public static class Column {
        private String header;
        private RequiredStatus requiredStatus;
        private String description;
        private String format;
        private String tooltip;
        private ColumnType type;
        @Builder.Default
        private List<String> allowedValues = new ArrayList<>();

        public List<String> getAllowedValues() {
            return Collections.unmodifiableList(allowedValues != null ? allowedValues : new ArrayList<>());
        }

        public void setAllowedValues(List<String> allowedValues) {
            this.allowedValues = Optional.ofNullable(allowedValues)
                    .orElseGet(List::of)
                    .stream()
                    .filter(ExcelTemplateDefinition::hasText)
                    .collect(Collectors.toCollection(ArrayList::new));
        }

        public ColumnType resolvedType() {
            return Optional.ofNullable(type).orElse(ColumnType.TEXT);
        }

        public String resolvedFormat() {
            return Optional.ofNullable(format)
                    .map(String::trim)
                    .filter(ExcelTemplateDefinition::hasText)
                    .orElseGet(() -> resolvedType().defaultFormat());
        }

        public List<String> resolvedAllowedValues() {
            List<String> values = getAllowedValues();
            if (values.isEmpty() && resolvedType() == ColumnType.BOOLEAN) {
                return List.of("YES", "NO");
            }
            return values;
        }

        public boolean isRequired() {
            return requiredStatus != null && requiredStatus.isRequired();
        }

        public void setRequired(boolean required) {
            this.requiredStatus = required ? REQUIRED : NOT_REQUIRED;
        }
    }

    public enum RequiredStatus {
        REQUIRED, NOT_REQUIRED;
        
        public boolean isRequired() {
            return this == REQUIRED;
        }
    }

    public enum ColumnType {
        TEXT("TEXT", "@", CellType.STRING),
        NUMBER("NUMBER", "#,##0.00############", CellType.NUMERIC),
        DATE("DATE", "dd/mm/yyyy", CellType.NUMERIC),
        LIST("LIST", "@", CellType.STRING),
        BOOLEAN("BOOLEAN", "@", CellType.STRING);

        private final String label;
        private final String defaultFormat;
        private final CellType cellType;

        ColumnType(String label, String defaultFormat, CellType cellType) {
            this.label = label;
            this.defaultFormat = defaultFormat;
            this.cellType = cellType;
        }

        public String defaultFormat() {
            return defaultFormat;
        }

        public CellType cellType() {
            return cellType;
        }
    }

    @Getter
    @Setter
    public static class TemplateSettings {
        private List<String> sheets = new ArrayList<>();

        public void setSheets(List<String> sheets) {
            this.sheets = Optional.ofNullable(sheets)
                    .orElseGet(List::of)
                    .stream()
                    .filter(ExcelTemplateDefinition::hasText)
                    .collect(Collectors.toCollection(ArrayList::new));
        }
    }

    private static boolean hasText(String value) {
        return value != null && !value.isBlank();
    }
}
