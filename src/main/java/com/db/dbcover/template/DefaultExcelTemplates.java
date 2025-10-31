package com.db.dbcover.template;

import com.db.dbcover.config.ExcelTemplateProperties;
import com.db.dbcover.template.ExcelTemplateDefinition.Column;
import com.db.dbcover.template.ExcelTemplateDefinition.ColumnType;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSettings;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;
import lombok.AccessLevel;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

@NoArgsConstructor(access = AccessLevel.PRIVATE)
public final class DefaultExcelTemplates {

    public static ExcelTemplateProperties properties() {
        TemplateSheet instrumentDetails = sheet("INSTRUMENT_DETAILS",
                textColumn("INSTRUMENT_ID", true, "Unique instrument identifier"),
                textColumn("INSTRUMENT_NAME", true, "Instrument name"),
                listColumn("CURRENCY", true, List.of("PLN", "EUR", "USD"), "ISO 4217 currency code"),
                dateColumn("ISSUE_DATE", false, "Issue date", "dd/mm/yyyy", null)
        );

        TemplateSheet linkedDeals = sheet("LINKED_DEALS",
                textColumn("DEAL_ID", true, null),
                listColumn("DEAL_TYPE", true, List.of("PRIMARY", "SECONDARY", "TERTIARY"), "Deal type (e.g. PRIMARY, SECONDARY)"),
                dateColumn("DEAL_DATE", true, null, "dd.mm.yyyy", "Select the deal date in dd.mm.yyyy format"),
                numberColumn("NOTIONAL", true, "Notional amount", null, "Provide the notional amount in the deal currency")
        );

        TemplateSheet linkedAssets = sheet("LINKED_ASSETS",
                textColumn("ASSET_ID", true, null),
                textColumn("ASSET_CLASS", true, null),
                numberColumn("ASSET_VALUE", false, null)
        );

        TemplateSheet persistedIds = sheet("PERSISTED_IDS",
                textColumn("ENTITY_TYPE", true, null),
                textColumn("LEGACY_ID", true, null),
                textColumn("SOURCE_SYSTEM", false, null)
        );

        TemplateSheet linkedInstruments = sheet("LINKED_INSTRUMENTS",
                textColumn("MASTER_INSTRUMENT_ID", true, null),
                textColumn("RELATED_INSTRUMENT_ID", true, null),
                textColumn("RELATIONSHIP_TYPE", true, null)
        );

        TemplateSheet linkedParties = sheet("LINKED_PARTIES",
                textColumn("PARTY_ID", true, null),
                textColumn("PARTY_ROLE", true, "Role (e.g. ISSUER, GUARANTOR)"),
                textColumn("PARTY_NAME", true, null),
                textColumn("COUNTRY", false, null)
        );

        ExcelTemplateProperties properties = ExcelTemplateProperties.builder()
                .templateSheets(List.of(
                        instrumentDetails,
                        linkedDeals,
                        linkedAssets,
                        persistedIds,
                        linkedInstruments,
                        linkedParties
                ))
                .instrumentTemplates(Map.of(
                        "MORTGAGE", settings(List.of(
                                "INSTRUMENT_DETAILS",
                                "LINKED_DEALS",
                                "LINKED_ASSETS",
                                "PERSISTED_IDS",
                                "LINKED_INSTRUMENTS",
                                "LINKED_PARTIES"
                        ))
                ))
                .build();

        properties.resolvedTemplateSheetIndex();
        return properties;
    }

    private static TemplateSheet sheet(String name, Column... columns) {
        return TemplateSheet.builder()
                .name(name)
                .columns(List.of(columns))
                .build();
    }

    private static Column textColumn(String header, boolean required, String description) {
        return Column.builder()
                .header(header)
                .required(required)
                .description(description)
                .type(ColumnType.TEXT)
                .build();
    }

    private static Column numberColumn(String header, boolean required, String description) {
        return numberColumn(header, required, description, null, null);
    }

    private static Column numberColumn(String header,
                                       boolean required,
                                       String description,
                                       String format,
                                       String tooltip) {
        return Column.builder()
                .header(header)
                .required(required)
                .description(description)
                .format(format)
                .tooltip(tooltip)
                .type(ColumnType.NUMBER)
                .build();
    }

    private static Column dateColumn(String header,
                                     boolean required,
                                     String description,
                                     String format,
                                     String tooltip) {
        return Column.builder()
                .header(header)
                .required(required)
                .description(description)
                .format(format)
                .tooltip(tooltip)
                .type(ColumnType.DATE)
                .build();
    }

    private static Column listColumn(String header,
                                     boolean required,
                                     List<String> values,
                                     String description) {
        Column column = Column.builder()
                .header(header)
                .required(required)
                .description(description)
                .type(ColumnType.LIST)
                .build();
        column.setAllowedValues(values);
        return column;
    }

    private static TemplateSettings settings(List<String> sheets) {
        TemplateSettings settings = new TemplateSettings();
        settings.setSheets(sheets);
        return settings;
    }
}
