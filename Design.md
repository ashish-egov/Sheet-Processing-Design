# 🧩 Unified Excel Sheet Management System  
**Multi-Sheet | Config-Driven | Interface-Based | Excel ↔ JSON ↔ Excel**

---

## 🎯 What Is It?

A powerful system to:

- ✔️ Read Excel sheets and convert them into structured JSON
- ✔️ Apply data transformations and validations
- ✔️ Generate Excel templates based on config
- ✔️ Support multiple sheets in one Excel file
- ✔️ Dynamically control formatting/styling with metadata

---

## 📦 Input & Output Format

### ✅ `SheetDataMap`

```ts
type SheetDataMap = Record<string, object[]>;
```

- Each key is a **sheet name** (e.g., "Employees")
- The value is an **array of objects**
- These objects represent **rows in Excel**
- ✅ **Sometimes**, the **first object** is metadata (not a real row), when it contains:
  
```ts
{ areColumnHeaders: true, ...column formatting info }
```

---

### 🔍 Column Metadata Object (first row only if `areColumnHeaders: true`)

```ts
{
  areColumnHeaders: true,
  Name: {
    isLocked: true,
    orderNumber: -1,
    color: "#CCE5FF",
    width: 120,
    hidden: false
  },
  Department: {
    orderNumber: 0
  }
}
```

---

### 📤 Example Output SheetDataMap

```ts
{
  "Employees": [
    {
      areColumnHeaders: true,
      Name: { isLocked: true, orderNumber: -1, color: "#CCE5FF", width: 120 },
      Department: { orderNumber: 0 }
    },
    { Name: "Alice", Department: "HR" },
    { Name: "Bob", Department: "IT" }
  ]
}
```

- The first object is only used for formatting the final Excel.
- The actual data starts from the **second object**.

---

## 🧠 Architecture Overview

### 🔹 Top-Level Config

```ts
interface SheetProcessorConfigEntry {
  templateType: string;
  sheets: SheetEntry[];
}
```

### 🔸 Per-Sheet Config

```ts
interface SheetEntry {
  sheetName: string;
  schemaName: string;
  processingClass?: string;
  generationClass?: string;
  sheetDataMapping?: SheetDataMapping[];
}
```

### 🔹 Optional Data Mapping

```ts
interface SheetDataMapping {
  inJsonPath: string;
  outJsonPath: string;
}
```

---

## 🔁 Flow 1: Process Flow (Excel → JSON → Output)

```ts
async function runExcelProcessing(
  templateType: string,
  inputData: SheetDataMap,
  contextData: Record<string, any>
): Promise<SheetDataMap>
```

### ✅ How it works:

1. Find config using `templateType`
2. Loop through each sheet
3. Load the `processingClass` dynamically
4. Call its `.process()` method
5. Return processed `SheetDataMap`

---

## 🔁 Flow 2: Generate Flow (Config → Excel → JSON)

```ts
async function runExcelGeneration(
  templateType: string,
  contextData: Record<string, any>
): Promise<SheetDataMap>
```

### ✅ How it works:

1. Find config using `templateType`
2. Loop through each sheet
3. Load the `generationClass` dynamically
4. Call its `.generate()` method
5. Return generated `SheetDataMap`

---

## 🏭 Dynamic Class Registry

```ts
const processorRegistry = {
  EmployeeSheetProcessor,
  DepartmentSheetProcessor,
};

const generatorRegistry = {
  EmployeeSheetGenerator,
  DepartmentSheetGenerator,
};

function getSheetProcessor(className: string) {
  const Cls = processorRegistry[className];
  if (!Cls) throw new Error("Processor not found");
  return new Cls();
}

function getSheetGenerator(className: string) {
  const Cls = generatorRegistry[className];
  if (!Cls) throw new Error("Generator not found");
  return new Cls();
}
```

---

## 🧪 Sample Config

```ts
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
      }
    ]
  }
];
```

---

## 🧰 Helper (Extract Headers + Rows)

```ts
function extractHeaderAndRows(sheetData: object[]): {
  columnHeaderMeta?: object;
  rows: object[];
} {
  if (sheetData.length > 0 && (sheetData[0] as any).areColumnHeaders) {
    const [header, ...rows] = sheetData;
    return { columnHeaderMeta: header, rows };
  }
  return { rows: sheetData };
}
```
