# Excel Template Generator

A lightweight Spring Boot service that exposes an endpoint for downloading Excel workbooks generated from YAML or programmatic template definitions.

## Requirements

* Java 17
* Maven 3.9+

## Build and run

The project depends on the official Spring Boot and Apache POI artifacts. Use standard Maven commands to compile the application and execute the verification suite:

```bash
mvn clean verify
```

Start the application locally with:

```bash
mvn spring-boot:run
```

The `GET /excel/template` endpoint returns `template.xlsx`. Provide the `instrumentType` query parameter to pick an instrument-specific template (defined either in YAML or in code). Missing or unknown types trigger a `400 Bad Request` response.

## Template configuration

Template configuration lives under `excel.template` inside `src/main/resources/excel-templates.yml`, which is imported from `application.yml`:

* Declare reusable sheet blueprints inside `template-sheets`. Each entry lists the sheet name and its columns (headers, data types such as `TEXT`/`NUMBER`/`DATE`/`BOOLEAN`, required flags, descriptions, tooltips, optional Excel formats, and allowed values).
* Register every instrument—including the standard **MORTGAGE** template—inside `instrument-templates`. Each instrument simply references the sheet names that should appear in the generated workbook, pulling their definitions from `template-sheets`.
* Optionally inherit other instrument templates by listing them in `base-templates`. Sheets listed later in the hierarchy replace earlier ones with the same name.

Every generated sheet contains:

* Styled header cells (grey background, bold text, red font for required columns) and matching info cells frozen at the top.
* Built-in filters that span the header and info rows.
* Dropdown lists for columns backed by `allowed-values`, including automatic `YES`/`NO` lists for Boolean fields.
* Numeric and date data validations that restrict entry to valid numbers and Excel date values.
* Default formats applied to entire columns (`dd/mm/yyyy` for dates, `@` for text, `#,##0.00############` for numbers) with the option to override them via the `format` field.
* Classic header comments with either the configured tooltip text or the same metadata shown in the info row (including format hints).

Updating the YAML file and restarting the application is enough to regenerate the workbook with the new structure.

### Defining templates in code

Use the plain `ExcelTemplateDefinition` model to assemble templates programmatically—for example, in tests or integration flows:

```java
ExcelTemplateDefinition definition = new ExcelTemplateDefinition();

ExcelTemplateDefinition.TemplateSheet sheet = new ExcelTemplateDefinition.TemplateSheet();
sheet.setName("CUSTOM");

ExcelTemplateDefinition.Column id = new ExcelTemplateDefinition.Column();
id.setHeader("ID");
id.setType(ExcelTemplateDefinition.ColumnType.TEXT);
id.setRequired(true);
id.setDescription("Entity identifier");

ExcelTemplateDefinition.Column active = new ExcelTemplateDefinition.Column();
active.setHeader("ACTIVE");
active.setType(ExcelTemplateDefinition.ColumnType.BOOLEAN);
active.setTooltip("Set to YES or NO");

ExcelTemplateDefinition.Column settlement = new ExcelTemplateDefinition.Column();
settlement.setHeader("SETTLEMENT_DATE");
settlement.setType(ExcelTemplateDefinition.ColumnType.DATE);
settlement.setFormat("dd.mm.yyyy");
settlement.setDescription("Settlement date");

sheet.setColumns(List.of(id, active, settlement));
definition.setSheets(List.of(sheet));

byte[] bytes = excelGeneratorService.generateTemplate(definition);
```

The API mirrors the YAML structure, so you can mix both approaches and reuse the generator across modules. For a ready-made setup that matches the bundled YAML, call `DefaultExcelTemplates.properties()` to obtain `ExcelTemplateProperties` populated with the mortgage template.

### Multiple instrument types

Define alternative templates under `excel.template.instrument-templates`. During request handling the service merges any `base-templates` before resolving the sheet list, letting you reuse common definitions while still overriding the set of sheets for a specific instrument type.
