# 🧩 Unified Excel Sheet Management System

**(Multi-Sheet Support | Interface-Driven | Config-Based | Excel ↔ JSON ↔ Excel)**

---

## 🎯 Objective

To build a **configurable** and **extensible** Excel sheet system that can:

- ✅ Parse multi-sheet Excel files into structured JavaScript objects
- 🔁 Process data using per-sheet transformation/validation logic
- 📄 Generate Excel files **from scratch**, using only config (no pre-existing Excel)
- 🧠 Use **interface-driven** logic for reusable, plug-and-play sheet processors
- 🎨 Control **visibility, layout, and styling** using flags from a config (MDMS)

---

## 🧠 Core Architecture

```
                         ┌──────────────────────────────┐
                         │     SheetProcessorConfig     │
                         │      (multi-sheet aware)     │
                         └──────────────┬───────────────┘
                                        ▼
                         ┌────────────────────────────────────────┐
                         │ Process Flow (Excel → JSON → Output)   │
                         └──────────────┬─────────────────────────┘
                                        ▼
                    ┌──────────────────────────────────────┐
                    │   { sheetName: string[][] }          │ ◄─ Parsed Excel
                    └──────────────────────────────────────┘
                                        ▼
                        ┌──────────────────────────────┐
                        │  ProcessorClass.process()     │
                        └────────────┬─────────────────┘
                                     ▼
                    ┌──────────────────────────────────────────────┐
                    │ Output: { sheetName: string[][] }            │
                    └──────────────────────────────────────────────┘


(Separate Flow)


                         ┌────────────────────────────────────────┐
                         │ Generate Flow (Config → Excel → JSON) │
                         └──────────────┬─────────────────────────┘
                                        ▼
                          ┌─────────────────────────────┐
                          │ Start with config + context │
                          └──────────────┬──────────────┘
                                         ▼
                          ┌────────────────────────────┐
                          │ GeneratorClass.generate()  │
                          └──────────────┬─────────────┘
                                         ▼
                        ┌──────────────────────────────────────────────┐
                        │ Output: { sheetName: string[][] }            │
                        └──────────────────────────────────────────────┘
```

---

## 🔧 Top-Level Config

### `SheetProcessorConfigEntry`

```ts
export interface SheetProcessorConfigEntry {
  templateType: string;            // Unique identifier per template
  parsingIdentifier?: string;      // Optional fallback to detect input Excel
  sheets: SheetEntry[];            // One config per Excel sheet
}
```

---

## 📄 Per-Sheet Config

### `SheetEntry`

```ts
export interface SheetEntry {
  sheetName: string;               // Sheet tab name in Excel
  schemaName: string;              // JSON schema name to validate this sheet
  processingClass?: string;        // Processor used in Process Flow
  generationClass?: string;        // Generator used in Generate Flow
  sheetDataMapping?: SheetDataMapping[]; // Used only in Process Flow
}
```

---

## 🔁 Data Mapping (Processing Only)

```ts
export interface SheetDataMapping {
  inJsonPath: string;   // Input path in raw sheet data
  outJsonPath: string;  // Transformed output JSON path
}
```

---

## 🔁 Input / Output Format

For both flows, data exchanged is:

```ts
type SheetDataMap = Record<string, string[][]>; // sheetName -> 2D array
```

---

## 🔧 Interfaces

### 🛠 Sheet Processor (Used in Process Flow)

```ts
export interface SheetProcessorInterface {
  process(
    sheetData: SheetDataMap,
    context: Record<string, any>
  ): Promise<SheetDataMap>;
}
```

### 🧾 Sheet Generator (Used in Generate Flow)

```ts
export interface SheetGeneratorInterface {
  generate(
    context: Record<string, any>
  ): Promise<SheetDataMap>;
}
```

---

## 🔁 Process Flow Logic

```ts
export async function runExcelProcessing(
  templateType: string,
  inputData: SheetDataMap,
  contextData: Record<string, any>
): Promise<SheetDataMap> {
  const config = SheetProcessorConfig.find(cfg => cfg.templateType === templateType);
  if (!config) throw new Error(`No config found for ${templateType}`);

  const result: SheetDataMap = {};

  for (const sheet of config.sheets) {
    if (!sheet.processingClass) continue;
    const processor = getSheetProcessor(sheet.processingClass);
    const sheetInput = { [sheet.sheetName]: inputData[sheet.sheetName] || [] };
    const output = await processor.process(sheetInput, contextData);
    result[sheet.sheetName] = output[sheet.sheetName];
  }

  return result;
}
```

---

## 📄 Generate Flow Logic

```ts
export async function runExcelGeneration(
  templateType: string,
  contextData: Record<string, any>
): Promise<SheetDataMap> {
  const config = SheetProcessorConfig.find(cfg => cfg.templateType === templateType);
  if (!config) throw new Error(`No config found for ${templateType}`);

  const result: SheetDataMap = {};

  for (const sheet of config.sheets) {
    if (!sheet.generationClass) continue;
    const generator = getSheetGenerator(sheet.generationClass);
    const output = await generator.generate(contextData);
    result[sheet.sheetName] = output[sheet.sheetName];
  }

  return result;
}
```

---

## 🏭 Factory Loader

```ts
const processorRegistry: Record<string, new () => SheetProcessorInterface> = {
  EmployeeSheetProcessor,
  DepartmentSheetProcessor,
};

const generatorRegistry: Record<string, new () => SheetGeneratorInterface> = {
  EmployeeSheetGenerator,
  DepartmentSheetGenerator,
};

function getSheetProcessor(className: string): SheetProcessorInterface {
  const Cls = processorRegistry[className];
  if (!Cls) throw new Error(`Processor not found: ${className}`);
  return new Cls();
}

function getSheetGenerator(className: string): SheetGeneratorInterface {
  const Cls = generatorRegistry[className];
  if (!Cls) throw new Error(`Generator not found: ${className}`);
  return new Cls();
}
```

---

## 🧪 Sample Config

```ts
export const SheetProcessorConfig: SheetProcessorConfigEntry[] = [
  {
    templateType: "HRBulkUpload",
    parsingIdentifier: "hr-bulk",
    sheets: [
      {
        sheetName: "Employees",
        schemaName: "EmployeeSchema",
        processingClass: "EmployeeSheetProcessor",
        generationClass: "EmployeeSheetGenerator",
        sheetDataMapping: [
          { inJsonPath: "employees.name", outJsonPath: "name" },
          { inJsonPath: "employees.department", outJsonPath: "department" }
        ]
      },
      {
        sheetName: "Departments",
        schemaName: "DepartmentSchema",
        processingClass: "DepartmentSheetProcessor",
        generationClass: "DepartmentSheetGenerator",
        sheetDataMapping: [
          { inJsonPath: "departments.name", outJsonPath: "name" },
          { inJsonPath: "departments.manager", outJsonPath: "manager" }
        ]
      }
    ]
  }
];
```

---

## 🎨 MDMS Styling & Layout Flags

These flags will be added to the **existing MDMS admin-level column schema** to control **how columns appear or behave** in the Excel files.

```json
{
  "freezeInProcessedFile": true,         // ❄️ Freeze column in processed Excel
  "hideColumnInProcessedFile": true,     // 🙈 Hide column from processed Excel
  "columnColorInProcessedFile": "#FF5733", // 🎨 Column background color (processed only)
  "columnWidth": 100                     // 📏 Applies in both generated and processed files
}
```

### 🔍 Flag Usage Summary

| Flag                         | Used In        | Description                                                                 |
|------------------------------|----------------|-----------------------------------------------------------------------------|
| `freezeInProcessedFile`      | Processed Only | Keeps selected columns fixed while scrolling horizontally.                  |
| `hideColumnInProcessedFile`  | Processed Only | Excludes columns from being visible in the processed Excel.                |
| `columnColorInProcessedFile` | Processed Only | Adds background color for emphasis or visual clarity.                       |
| `columnWidth`                | Both Flows     | Controls column width in both generated and processed Excel files.         |

---

### 🧾 Example Column Definition in MDMS

```json
{
  "header": "Employee ID",
  "jsonPath": "employee.id",
  "freezeInProcessedFile": true,
  "hideColumnInProcessedFile": false,
  "columnColorInProcessedFile": "#CCE5FF",
  "columnWidth": 120
}
```

