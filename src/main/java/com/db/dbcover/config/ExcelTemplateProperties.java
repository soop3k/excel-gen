package com.db.dbcover.config;

import com.db.dbcover.template.ExcelTemplateDefinition;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSettings;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;
import com.db.dbcover.template.TemplateSheetResolver;
import lombok.Getter;
import org.springframework.boot.context.properties.ConfigurationProperties;

import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Getter
@ConfigurationProperties(prefix = "excel.template")
public class ExcelTemplateProperties {

    private final List<TemplateSheet> templateSheets = new ArrayList<>();
    private final Map<String, TemplateSettings> instrumentTemplates = new LinkedHashMap<>();
    private transient Map<String, TemplateSheet> resolvedTemplateSheets;
    private transient Map<String, ExcelTemplateDefinition> resolvedDefinitions;

    public void setTemplateSheets(List<TemplateSheet> templateSheets) {
        this.templateSheets.clear();
        if (templateSheets != null) {
            this.templateSheets.addAll(templateSheets);
        }
        this.resolvedTemplateSheets = null;
        this.resolvedDefinitions = null;
    }

    public void setInstrumentTemplates(Map<String, TemplateSettings> instrumentTemplates) {
        this.instrumentTemplates.clear();
        if (instrumentTemplates != null) {
            this.instrumentTemplates.putAll(instrumentTemplates);
        }
        this.resolvedDefinitions = null;
    }

    public static Builder builder() {
        return new Builder();
    }

    public Map<String, TemplateSheet> resolvedTemplateSheetIndex() {
        if (resolvedTemplateSheets == null) {
            resolvedTemplateSheets = TemplateSheetResolver.resolveSheets(getTemplateSheets());
        }
        return resolvedTemplateSheets;
    }

    public Map<String, ExcelTemplateDefinition> resolvedInstrumentTemplates() {
        if (resolvedDefinitions == null) {
            resolvedDefinitions = TemplateSheetResolver.resolveDefinitions(
                    resolvedTemplateSheetIndex(),
                    new LinkedHashMap<>(instrumentTemplates)
            );
        }
        return resolvedDefinitions;
    }

    public static final class Builder {
        private final List<TemplateSheet> templateSheets = new ArrayList<>();
        private final Map<String, TemplateSettings> instrumentTemplates = new LinkedHashMap<>();

        public Builder addTemplateSheet(TemplateSheet sheet) {
            if (sheet != null) {
                this.templateSheets.add(sheet);
            }
            return this;
        }

        public Builder templateSheets(Collection<TemplateSheet> sheets) {
            this.templateSheets.clear();
            if (sheets != null) {
                this.templateSheets.addAll(sheets);
            }
            return this;
        }

        public Builder putInstrumentTemplate(String name, TemplateSettings settings) {
            if (name != null && settings != null) {
                this.instrumentTemplates.put(name, settings);
            }
            return this;
        }

        public Builder instrumentTemplates(Map<String, TemplateSettings> templates) {
            this.instrumentTemplates.clear();
            if (templates != null) {
                this.instrumentTemplates.putAll(templates);
            }
            return this;
        }

        public ExcelTemplateProperties build() {
            ExcelTemplateProperties properties = new ExcelTemplateProperties();
            properties.templateSheets.addAll(this.templateSheets);
            properties.instrumentTemplates.putAll(this.instrumentTemplates);
            return properties;
        }
    }
}
