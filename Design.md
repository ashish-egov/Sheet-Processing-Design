# 🧩 Unified Excel Sheet Management System  
### (Multi-Sheet Support | Interface-Driven | Config-Based | Excel ↔ JSON ↔ Excel)

---

## 🎯 Objective

Design a configurable and extensible Excel sheet management system with the following capabilities:

- 🧾 Parsing multi-sheet Excel files into structured JS objects
- 🔁 Processing data based on sheet-level configurations
- 📄 Generating Excel sheets purely from configurations, **without starting from Excel**
- 🧠 Using interface-driven transformation and validation logic for data processing
- 🧱 Supporting flags for controlling the structure, visibility, and styling of Excel sheets

---

## 🧠 Core Architecture

```
┌──────────────────────────────┐
│     SheetProcessorConfig     │
│      (multi-sheet aware)     │
└──────────────┬───────────────┘
               ▼
       ┌────────────────────────────────────────┐
       │ Process Flow (multi-sheet, Excel → JS) │
       └──────────────┬────────────────────────┘
                      ▼
          ┌──────────────────────────────┐
          │ { sheetName: string[][] }    │ ◄──── Excel parsed into object
          └──────────────┬──────────────┘
                         ▼
            ┌────────────────────────────┐
            │ ProcessorClass.process()  │
            └────────────┬──────────────┘
                         ▼
            ┌───────────────────────────────────────────────┐
            │ Output: { sheetName: string[][] } (processed) │
            └───────────────────────────────────────────────┘
                     
           (Separate Flow)
               
       ┌────────────────────────────────────────┐
       │ Generate Flow (config → Excel → JS) │
       └──────────────┬────────────────────────┘
                      ▼
            ┌────────────────────────────┐
            │ Start with config + context │
            └──────────────┬──────────────┘
                         ▼
            ┌────────────────────────────┐
            │ GeneratorClass.generate()  │
            └────────────┬──────────────┘
                         ▼
            ┌───────────────────────────────────────────────┐
            │ Output: { sheetName: string[][] } (generated) │
            └───────────────────────────────────────────────┘
```

---

## 🔧 Top-Level Config

```ts
export interface SheetProcessorConfigEntry {
  templateType: string;            // Unique identifier per use-case
  parsingIdentifier: string;       // Optional fallback for file detection
  sheets: SheetEntry[];            // One entry per sheet to process or generate
}
```

---

## 📄 Per Sheet Config

```ts
export interface SheetEntry {
  sheetName: string;               // Target sheet name
  schemaName: string;              // JSON schema to validate structure
  processingClass?: string;        // Class used during PROCESS flow
  generationClass?: string;        // Class used during GENERATE flow
  sheetDataMapping?: SheetDataMapping[]; // Used only in PROCESS flow
}
```

### ⚠️ Notes:
- `processingClass`: Required **only** for the **process flow**.
- `generationClass`: Required **only** for the **generate flow**.
- `sheetDataMapping`: Used **only in process flow** for data extraction/transformation. **Ignored during generation**.

---

## 🔁 Data Mapping (Processing-Only)

```ts
export interface SheetDataMapping {
  inJsonPath: string;
  outJsonPath: string;
}
```

### 📌 Purpose:
This mapping facilitates extraction or transformation of data from Excel sheets **during process flow only**.

---

## 🔁 I/O Structure

For both flows, the data exchanged is structured as an object mapping sheet names to 2D arrays of strings:

```ts
type SheetDataMap = Record<string, string[][]>;
```

---

## 🔧 Interfaces

### 🛠 Sheet Processor (Used in Process Flow)

```ts
export interface SheetProcessorInterface {
  process(
    sheetData: SheetDataMap,         // Input from parsed Excel
    context: Record<string, any>     // External/injected metadata
  ): Promise<SheetDataMap>;          // Output with same sheet names
}
```

### 🧾 Sheet Generator (Used in Generate Flow)

```ts
export interface SheetGeneratorInterface {
  generate(
    context: Record<string, any>     // Config + external context
  ): Promise<SheetDataMap>;          // Returns all generated sheets
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

## 🧾 Generate Flow Logic

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
          {
            inJsonPath: "employees.name",
            outJsonPath: "name"
          },
          {
            inJsonPath: "employees.department",
            outJsonPath: "department"
          }
        ]
      },
      {
        sheetName: "Departments",
        schemaName: "DepartmentSchema",
        processingClass: "DepartmentSheetProcessor",
        generationClass: "DepartmentSheetGenerator",
        sheetDataMapping: [
          {
            inJsonPath: "departments.name",
            outJsonPath: "name"
          },
          {
            inJsonPath: "departments.manager",
            outJsonPath: "manager"
          }
        ]
      }
    ]
  }
];
```

---

## 🧾 MDMS Column Flags (Used for Styling & Layout)

```json
{
  "freezeInProcessedFile": { "type": "boolean" },
  "hideColumnInProcessedFile": { "type": "boolean" },
  "columnColorInProcessedFile": {
    "type": "string",
    "pattern": "^#([A-Fa-f0-9]{6})$"
  },
  "columnWidth": {
    "type": "number",
    "minimum": 10,
    "maximum": 500
  }
}
```

### 🔍 Flag Behavior Summary

| Flag                          | Process Flow | Generate Flow | Description                                   |
|-------------------------------|--------------|----------------|-----------------------------------------------|
| `freezeInProcessedFile`       | ✅ Used       | ❌ Ignored      | Freeze columns in post-validation output only |
| `hideColumnInProcessedFile`   | ✅ Used       | ❌ Ignored      | Column visibility in processed Excel only     |
| `columnColorInProcessedFile`  | ✅ Used       | ❌ Ignored      | Styling for processed file readability        |
| `columnWidth`                 | ✅ Used       | ✅ Used         | Applies to both generated and processed files |

