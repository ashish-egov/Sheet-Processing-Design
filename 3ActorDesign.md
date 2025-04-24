## 🧩 Unified Excel Sheet Management System  
**Multi-Sheet | Config-Driven | Interface-Based | Excel ↔ JSON ↔ Excel**

---

### 🧠 1. TemplateManager  
**Role:** Central coordinator for both generate and process flows.

#### 🔁 Generate Flow:
- Input: `templateType` (e.g., `"HRBulkUpload"`)
- Loads MDMS config & schema:
  - Sheet names, columns, styling, locks, hidden flags, order, widths
- Creates a blank Excel with:
  - All sheets, placeholder headers, default styles
- Loads required `generationClass` instances
- Calls `.generate()` → gets `SheetDataMap`s
- Merges all sheet maps
- Fills Excel:
  - Column metadata (if provided)
  - Sample data rows
  - Applies styles, widths, locks
- Returns final Excel file

#### 🔁 Process Flow:
- Input: Uploaded Excel + `templateType`
- Loads config & schema
- Reads raw data + optional metadata from first row
- Loads required `processingClass` instances
- Calls `.process()` → gets validated `SheetDataMap`s
- Merges all sheet maps
- Updates original Excel (optional):
  - Transformed data
  - Styling or annotations
- Returns final Excel or processed JSON

---

### ⚙️ 2. Generator  
**Role:** Provides sample data + optional column metadata.

#### `.generate()` returns:
```ts
{
  "Employees": [
    {
      areColumnHeaders: true,
      Name: { color: "#CCE5FF", isLocked: true, width: 120, orderNumber: -2 },
      Department: { orderNumber: -1 }
    },
    { Name: "Alice", Department: "HR" },
    { Name: "Bob", Department: "IT" }
  ]
}
```
- First object = optional column metadata  
- Rest = sample rows

---

### 🔬 3. Processor  
**Role:** Validates + transforms uploaded data and may return column metadata.

#### `.process()` returns:
```ts
{
  "Employees": [
    {
      areColumnHeaders: true,
      name: { width: 100, color: "#D9EDF7" },
      department: { width: 90 }
    },
    { name: "Alice", department: "HR" },
    { name: "Bob", department: "IT" }
  ]
}
```
- First object = optional column metadata  
- Rest = cleaned, validated rows

---

### 📦 Data Structure: `SheetDataMap`
```ts
type SheetDataMap = Record<string, object[]>;
```
- **Key** = sheet name  
- **Value** = array of rows  
- **1st object** (optional): column metadata (e.g., width, color, lock, hidden)  
- Remaining objects: actual data rows

---

### 🧪 Sample Config
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

### 📊 Summary

| Actor           | Role                                      | Called During     | Output          |
|----------------|-------------------------------------------|-------------------|-----------------|
| TemplateManager| Coordinates everything                    | Generate & Process| Excel + JSON    |
| Generator       | Returns sample data + optional metadata   | Generate Flow     | `SheetDataMap`  |
| Processor       | Validates + transforms + optional metadata| Process Flow      | `SheetDataMap`  |

---
