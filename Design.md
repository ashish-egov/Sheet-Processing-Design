# 🧩 Unified Excel Sheet Management System  
**Multi-Sheet | Config-Driven | Interface-Based | Excel ↔ JSON ↔ Excel**

---

## 🎯 Goal
Build a powerful and flexible system that can:

- ✅ Read Excel files with multiple sheets and convert them to structured JSON.
- 🔁 Process that data with custom logic for each sheet.
- 📄 Generate Excel files **from scratch** based on just config.
- 🔌 Use pluggable classes for logic (easy to extend).
- 🎨 Style and control sheet layout via config flags.

---

## 🧠 How It Works — Two Main Flows

### 1️⃣ Process Flow (Excel ➡ JSON ➡ Processed Output)

- You give an **input Excel file**.
- System uses **per-sheet processors** (via config).
- Each processor transforms or validates the sheet.
- Final result is an updated version of the Excel data in `string[][]` format.

```ts
Input: Excel → Convert to string[][] per sheet → Run logic → Output: Processed string[][]
```

---

### 2️⃣ Generate Flow (Config ➡ Excel ➡ JSON)

- You don’t give any Excel file.
- System looks up the config and runs **per-sheet generators**.
- Each generator creates a `string[][]` for that sheet based on rules or default values.
- Result is a fresh Excel built from scratch.

```ts
Input: Config + Context → Run Generator → Output: New Excel (as string[][])
```

---

## 📦 Top-Level Config

```ts
SheetProcessorConfigEntry {
  templateType: string,        // e.g. "HRBulkUpload"
  parsingIdentifier?: string,  // Optional ID to detect input file
  sheets: SheetEntry[]         // Sheet-wise config
}
```

---

## 📄 Per Sheet Config

```ts
SheetEntry {
  sheetName: string,               // Tab name in Excel
  schemaName: string,              // Validation schema
  processingClass?: string,        // For processing input Excel
  generationClass?: string,        // For generating new Excel
  sheetDataMapping?: SheetDataMapping[]  // Maps raw to output fields
}
```

---

## 🔄 Data Mapping (Used During Processing)

```ts
SheetDataMapping {
  inJsonPath: string,    // From raw sheet
  outJsonPath: string    // Where to place in output
}
```

---

## 🔁 Input & Output Format

Data always flows like this (in both flows):

```ts
type SheetDataMap = Record<string, string[][]>; // { sheetName → 2D array }
```

---

## 🧠 Interfaces to Implement

### 1. 🛠 Sheet Processor

```ts
interface SheetProcessorInterface {
  process(sheetData: SheetDataMap, context: any): Promise<SheetDataMap>;
}
```

### 2. 🧾 Sheet Generator

```ts
interface SheetGeneratorInterface {
  generate(context: any): Promise<SheetDataMap>;
}
```

---

## ▶️ Run Logic

### ✅ Run Process Flow

```ts
runExcelProcessing(templateType, inputData, context)
```

- Finds the config using `templateType`
- Runs the appropriate `processingClass` for each sheet
- Returns processed data

---

### 🆕 Run Generate Flow

```ts
runExcelGeneration(templateType, context)
```

- Finds the config using `templateType`
- Runs the appropriate `generationClass` for each sheet
- Returns freshly generated sheet data

---

## 🧰 Factory Loader (Behind the scenes)

These map class names (as strings) to actual code:

```ts
const processorRegistry = {
  EmployeeSheetProcessor,
  DepartmentSheetProcessor,
};

const generatorRegistry = {
  EmployeeSheetGenerator,
  DepartmentSheetGenerator,
};
```

When system sees `processingClass: "EmployeeSheetProcessor"`, it uses this registry to load the class and call its methods.

---

## 🧪 Sample Config

```ts
const SheetProcessorConfig = [
  {
    templateType: "HRBulkUpload",
    sheets: [
      {
        sheetName: "Employees",
        schemaName: "EmployeeSchema",
        processingClass: "EmployeeSheetProcessor",
        generationClass: "EmployeeSheetGenerator",
        sheetDataMapping: [
          { inJsonPath: "employees.name", outJsonPath: "name" },
          { inJsonPath: "employees.department", outJsonPath: "department" },
        ]
      },
      {
        sheetName: "Departments",
        schemaName: "DepartmentSchema",
        processingClass: "DepartmentSheetProcessor",
        generationClass: "DepartmentSheetGenerator",
        sheetDataMapping: [
          { inJsonPath: "departments.name", outJsonPath: "name" },
          { inJsonPath: "departments.manager", outJsonPath: "manager" },
        ]
      }
    ]
  }
];
```

---

## 🎨 Styling Flags (from MDMS)

You can add these flags to any column in your config:

| Flag                      | Where Used         | What It Does |
|---------------------------|--------------------|---------------------------|
| `freezeInProcessedFile`   | Processed Only     | Freezes column |
| `hideColumnInProcessedFile` | Processed Only  | Hides column |
| `columnColorInProcessedFile` | Processed Only | Sets background color |
| `columnWidth`             | Both Flows         | Sets column width |

### Example Column Definition:

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
