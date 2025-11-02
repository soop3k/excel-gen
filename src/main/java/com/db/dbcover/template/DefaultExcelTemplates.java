package com.db.dbcover.template;

import com.db.dbcover.config.ExcelTemplateProperties;
import com.db.dbcover.template.ExcelTemplateDefinition.Column;
import com.db.dbcover.template.ExcelTemplateDefinition.RequiredStatus;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSettings;
import com.db.dbcover.template.ExcelTemplateDefinition.TemplateSheet;
import lombok.AccessLevel;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

import static com.db.dbcover.template.ExcelTemplateDefinition.RequiredStatus.REQUIRED;
import static com.db.dbcover.template.ExcelTemplateDefinition.RequiredStatus.NOT_REQUIRED;
import static com.db.dbcover.template.ExcelTemplateDefinition.ColumnType.TEXT;
import static com.db.dbcover.template.ExcelTemplateDefinition.ColumnType.NUMBER;
import static com.db.dbcover.template.ExcelTemplateDefinition.ColumnType.DATE;
import static com.db.dbcover.template.ExcelTemplateDefinition.ColumnType.LIST;


@NoArgsConstructor(access = AccessLevel.PRIVATE)
public final class DefaultExcelTemplates {

    public static ExcelTemplateProperties properties() {
        TemplateSheet instrumentDetails = sheet("INSTRUMENT_DETAILS",
                textColumn("INSTRUMENT_ID", REQUIRED, "Unique instrument identifier"),
                textColumn("INSTRUMENT_NAME", REQUIRED, "Instrument name"),
                listColumn("CURRENCY", REQUIRED, List.of("PLN", "EUR", "USD"), "ISO 4217 currency code"),
                dateColumn("ISSUE_DATE", NOT_REQUIRED, "Issue date", "dd/mm/yyyy", null)
        );

        TemplateSheet linkedDeals = sheet("LINKED_DEALS",
                textColumn("DEAL_ID", REQUIRED),
                listColumn("DEAL_TYPE", REQUIRED, List.of("PRIMARY", "SECONDARY", "TERTIARY"), "Deal type (e.g. PRIMARY, SECONDARY)"),
                dateColumn("DEAL_DATE", REQUIRED, null, "dd.mm.yyyy", "Select the deal date in dd.mm.yyyy format"),
                numberColumn("NOTIONAL", REQUIRED, "Notional amount", null, "Provide the notional amount in the deal currency")
        );

        TemplateSheet linkedAssets = sheet("LINKED_ASSETS",
                textColumn("ASSET_ID", REQUIRED),
                textColumn("ASSET_CLASS", REQUIRED),
                numberColumn("ASSET_VALUE", NOT_REQUIRED)
        );

        TemplateSheet persistedIds = sheet("PERSISTED_IDS",
                textColumn("ENTITY_TYPE", REQUIRED),
                textColumn("LEGACY_ID", REQUIRED),
                textColumn("SOURCE_SYSTEM", NOT_REQUIRED)
        );

        TemplateSheet linkedInstruments = sheet("LINKED_INSTRUMENTS",
                textColumn("MASTER_INSTRUMENT_ID", REQUIRED),
                textColumn("RELATED_INSTRUMENT_ID", REQUIRED),
                textColumn("RELATIONSHIP_TYPE", REQUIRED)
        );

        TemplateSheet linkedParties = sheet("LINKED_PARTIES",
                textColumn("PARTY_ID", REQUIRED),
                textColumn("PARTY_ROLE", REQUIRED, "Role (e.g. ISSUER, GUARANTOR)"),
                textColumn("PARTY_NAME", REQUIRED),
                textColumn("COUNTRY", NOT_REQUIRED)
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

        properties.initialize();
        return properties;
    }

    private static TemplateSheet sheet(String name, Column... columns) {
        return TemplateSheet.builder()
                .name(name)
                .columns(List.of(columns))
                .build();
    }

    private static Column textColumn(String header, RequiredStatus requiredStatus) {
        return textColumn(header, requiredStatus, null);
    }

    private static Column textColumn(String header, RequiredStatus requiredStatus, String description) {
        return Column.builder()
                .header(header)
                .requiredStatus(requiredStatus)
                .description(description)
                .type(TEXT)
                .build();
    }

    private static Column numberColumn(String header, RequiredStatus requiredStatus) {
        return numberColumn(header, requiredStatus, null, null, null);
    }

    private static Column numberColumn(String header, RequiredStatus requiredStatus, String description) {
        return numberColumn(header, requiredStatus, description, null, null);
    }

    private static Column numberColumn(String header,
                                       RequiredStatus requiredStatus,
                                       String description,
                                       String format,
                                       String tooltip) {
        return Column.builder()
                .header(header)
                .requiredStatus(requiredStatus)
                .description(description)
                .format(format)
                .tooltip(tooltip)
                .type(NUMBER)
                .build();
    }

    private static Column dateColumn(String header, RequiredStatus requiredStatus) {
        return dateColumn(header, requiredStatus, null, null, null);
    }

    private static Column dateColumn(String header, RequiredStatus requiredStatus, String description) {
        return dateColumn(header, requiredStatus, description, null, null);
    }

    private static Column dateColumn(String header,
                                     RequiredStatus requiredStatus,
                                     String description,
                                     String format,
                                     String tooltip) {
        return Column.builder()
                .header(header)
                .requiredStatus(requiredStatus)
                .description(description)
                .format(format)
                .tooltip(tooltip)
                .type(DATE)
                .build();
    }

    private static Column listColumn(String header,
                                     RequiredStatus requiredStatus,
                                     List<String> values,
                                     String description) {
        Column column = Column.builder()
                .header(header)
                .requiredStatus(requiredStatus)
                .description(description)
                .type(LIST)
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
