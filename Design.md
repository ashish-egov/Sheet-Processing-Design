# 🧩 Unified Excel Sheet Management System  
**Multi-Sheet Support | Interface-Driven | Config-Based | Excel ↔ JSON ↔ Excel**

---

## 🎯 Objective

To build a configurable and extensible Excel sheet system that can:

✅ Parse multi-sheet Excel files into structured JavaScript objects  
🔁 Process data using per-sheet transformation/validation logic  
📄 Generate Excel files from scratch, using only config (no pre-existing Excel)  
🧠 Use interface-driven logic for reusable, plug-and-play sheet processors  
🎨 Control visibility, layout, and styling using flags from a config (MDMS)  

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
                    │ { sheetName: object[] }              │ ◄─ Parsed Excel
                    └──────────────────────────────────────┘
                                        ▼
                        ┌──────────────────────────────┐
                        │  ProcessorClass.process()     │
                        └────────────┬─────────────────┘
                                     ▼
                    ┌──────────────────────────────────────────────┐
                    │ Output: { sheetName: object[] }              │
                    └──────────────────────────────────────────────┘
```

---

### (Separate Flow)

```
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
                        │ Output: { sheetName: object[] }              │
                        └──────────────────────────────────────────────┘
```

---

## 🔧 Top-Level Config

```ts
export interface SheetProcessorConfigEntry {
  templateType: string;             // Unique identifier per template
  parsingIdentifier?: string;       // Optional fallback to detect input Excel
  sheets: SheetEntry[];             // Config per sheet/tab
}
```

---

## 📄 Per-Sheet Config

```ts
export interface SheetEntry {
  sheetName: string;                // Excel tab name
  schemaName: string;               // JSON schema name for validation
  processingClass?: string;         // For Process Flow
  generationClass?: string;         // For Generate Flow
  sheetDataMapping?: SheetDataMapping[]; // Used only in Process Flow
}
```

---

## 🔁 Data Mapping (Processing Only)

```ts
export interface SheetDataMapping {
  inJsonPath: string;    // Path in raw sheet object
  outJsonPath: string;   // Path in output object
}
```

---

## 🔁 Input / Output Format

```ts
type SheetDataMap = Record<string, object[]>; // sheetName -> array of row objects
```

---

## 🔧 Interfaces

### 🛠 Sheet Processor

```ts
export interface SheetProcessorInterface {
  process(
    sheetData: SheetDataMap,
    context: Record<string, any>
  ): Promise<SheetDataMap>;
}
```

### 🧾 Sheet Generator

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

These flags control how columns appear or behave in **processed Excel files** (some apply to generated files too).

```json
{
  "header": "Employee ID",
  "jsonPath": "employee.id",
  "freezeInProcessedFile": true,
  "hideColumnInProcessedFile": false,
  "color": "#CCE5FF",
  "columnWidth": 120
}
```

| Flag                    | Used In       | Description                                                        |
|-------------------------|---------------|--------------------------------------------------------------------|
| `freezeInProcessedFile` | Processed Only| Freezes the column for horizontal scrolling                        |
| `hideColumnInProcessedFile` | Processed Only| Excludes column from visible output in processed Excel             |
| `color`                 | Processed Only| Sets cell background color (for emphasis or style)                |
| `columnWidth`           | Both Flows    | Sets column width in both generated and processed Excel files      |
