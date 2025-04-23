## 🧩 Unified Excel Sheet Management System
### Multi-Sheet Support | Interface-Driven | Config-Based | Excel ↔ JSON ↔ Excel

### 🎯 Objective
Build a configurable system for processing and generating Excel sheets, supporting:
- Multi-sheet Excel parsing
- Data transformation/validation
- Excel generation from config
- Interface-driven logic for reusable sheet processors
- Flexible control over layout and styling via flags in config

---

### 🧠 Core Architecture Overview

1. **Top-Level Config**:
    - **SheetProcessorConfigEntry**: Describes how to process each template type (e.g., "HRBulkUpload").
    - **Per-Sheet Config**: For each sheet/tab in the template, provides processing and generation logic, data mapping, and flags.

2. **Process Flow** (Excel → JSON → Output):
    - Parse Excel → Map data → Apply transformations → Generate output.

3. **Generate Flow** (Config → Excel → JSON):
    - Use config to generate Excel from scratch → Map data → Output in structured format.

---

### 🔧 Key Config Structures

#### **Top-Level Config**
```typescript
export interface SheetProcessorConfigEntry {
  templateType: string;             // Template identifier
  sheets: SheetEntry[];             // List of sheets configuration
}
```

#### **Per-Sheet Config**
```typescript
export interface SheetEntry {
  sheetName: string;                // Excel tab name
  schemaName: string;               // JSON schema name for validation
  processingClass?: string;         // For Process Flow
  generationClass?: string;         // For Generate Flow
  sheetDataMapping?: SheetDataMapping[]; // Data mapping for Process Flow
}
```

#### **Data Mapping (For Processing)**
```typescript
export interface SheetDataMapping {
  inJsonPath: string;    // Path in raw sheet object
  outJsonPath: string;   // Path in output object
}
```

---

### 🔁 Input / Output Format
```typescript
type SheetDataMap = Record<string, object[]>; // sheetName -> array of row objects
```

---

### 🧠 Flow Logic: **Process Flow** (Excel → JSON → Output)
**Goal**: Convert Excel data into structured JSON objects, applying transformation/validation.

```typescript
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

#### **Flow Breakdown**:
1. **Fetch Config**: Get the configuration based on `templateType`.
2. **Loop through Sheets**: For each sheet in the config, apply the processing class.
3. **Processing Logic**: Use the `SheetProcessorInterface` to transform data.
4. **Return Processed Data**: Combine the results from each sheet into the final output.

---

### 🧠 Flow Logic: **Generate Flow** (Config → Excel → JSON)
**Goal**: Generate new Excel files from config-based data mappings.

```typescript
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

#### **Flow Breakdown**:
1. **Fetch Config**: Get the configuration based on `templateType`.
2. **Loop through Sheets**: For each sheet in the config, apply the generation class.
3. **Generation Logic**: Use the `SheetGeneratorInterface` to generate data for Excel.
4. **Return Generated Data**: Return the generated data, formatted per sheet.

---

### 🔧 Factory Loader
**Goal**: Dynamically load processors and generators based on config.

```typescript
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

### 🎨 MDMS Styling & Layout Flags
**Goal**: Control how columns appear in processed Excel files.

```typescript
{
  "header": "Employee ID",
  "jsonPath": "employee.id",
  "freezeInProcessedFile": true,    // Freeze the column for horizontal scrolling
  "hideColumnInProcessedFile": false, // Hide column in output
  "color": "#CCE5FF",               // Set background color
  "columnWidth": 120                // Set column width
}
```

**Flags**:
- `freezeInProcessedFile`: Freezes the column in the processed Excel file.
- `hideColumnInProcessedFile`: Excludes the column from the visible processed Excel output.
- `color`: Sets the background color for the cell.
- `columnWidth`: Sets the column width.

---

### 🧪 Sample Config Example
```typescript
export const SheetProcessorConfig: SheetProcessorConfigEntry[] = [
  {
    templateType: "HRBulkUpload",
    sheets: [
      {
        sheetName: "Employees",
        schemaName: "EmployeeSchema",
        processingClass: "EmployeeSheetProcessor",
        generationClass: "EmployeeSheetGenerator",
        sheetDataMapping: [
          { inJsonPath: "row.name", outJsonPath: "name" },
          { inJsonPath: "row.department", outJsonPath: "department" }
        ]
      },
      {
        sheetName: "Departments",
        schemaName: "DepartmentSchema",
        processingClass: "DepartmentSheetProcessor",
        generationClass: "DepartmentSheetGenerator",
        sheetDataMapping: [
          { inJsonPath: "row.name", outJsonPath: "name" },
          { inJsonPath: "row.manager", outJsonPath: "manager" }
        ]
      }
    ]
  }
];
```
