package com.db.dbcover.config;

import com.db.dbcover.template.ExcelTemplateDefinition;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSettings;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;
import com.db.dbcover.template.TemplateSheetResolver;
import com.db.dbcover.template.TemplateSheetResolver.ResolvedTemplates;
import org.springframework.boot.context.properties.ConfigurationProperties;

import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@ConfigurationProperties(prefix = "excel.template")
public class ExcelTemplateProperties {

    private final List<TemplateSheet> templateSheets;
    private final Map<String, TemplateSettings> instrumentTemplates;
    private transient ResolvedTemplates resolved;

    public ExcelTemplateProperties(List<TemplateSheet> templateSheets, Map<String, TemplateSettings> instrumentTemplates) {
        this.templateSheets = templateSheets != null ? templateSheets : List.of();
        this.instrumentTemplates = instrumentTemplates != null ? instrumentTemplates : Map.of();
    }

    public static Builder builder() {
        return new Builder();
    }

    public void initialize() {
        ensureResolved();
    }
    
    public Map<String, ExcelTemplateDefinition> resolvedInstrumentTemplates() {
        return ensureResolved().instrumentTemplates();
    }

    private ResolvedTemplates ensureResolved() {
        if (resolved == null) {
            resolved = TemplateSheetResolver.resolve(templateSheets, instrumentTemplates);
        }
        return resolved;
    }

    public static final class Builder {
        private final List<TemplateSheet> templateSheets = new ArrayList<>();
        private final Map<String, TemplateSettings> instrumentTemplates = new LinkedHashMap<>();

        public Builder templateSheets(Collection<TemplateSheet> sheets) {
            this.templateSheets.clear();
            if (sheets != null) {
                this.templateSheets.addAll(sheets);
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
            return new ExcelTemplateProperties(this.templateSheets, this.instrumentTemplates);
        }
    }
}
