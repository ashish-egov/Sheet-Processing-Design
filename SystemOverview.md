#### 🧩 Unified Excel Sheet Management System  
**Multi-Sheet | Config-Driven | Interface-Based | Excel ↔ JSON ↔ Excel**

---

#### 🔁 Two-Word Flows  

- **TemplateManager**
  - Generate Flow → `Create Template`
  - Process Flow → `Handle Upload`

- **Generator**
  - Input → `Template Type`
  - Output → `Sample Sheet`

- **Processor**
  - Input → `Uploaded File`
  - Output → `Validated Data`

---

#### 🎯 Actor Roles  

| Actor           | Role Description                              | Used In        | Output         |
|----------------|------------------------------------------------|----------------|----------------|
| TemplateManager | Central coordinator for both flows             | Generate, Process | Final Excel    |
| Generator       | Returns sample data + metadata                 | Generate       | SheetDataMap   |
| Processor       | Validates and transforms uploaded Excel        | Process        | SheetDataMap   |

---

#### 📦 Data Structure: `SheetDataMap`

```ts
type SheetDataMap = Record<string, object[]>;
```

- `Key` → Sheet name  
- `Value` → Array of row objects  
  - 1st object (optional): column metadata  
  - Remaining: actual row data  

---

#### 🧪 Sample Config

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

### 📄 `TemplateManager.md`

#### 🧠 TemplateManager  
**Role:** Central brain that orchestrates both flows.

---

#### 🔁 Generate Flow – `Create Template`

- Input → `templateType`
- Load → `MDMS config`, `Schema`
- Load → `generationClass`
- Call → `.generate()` → `SheetDataMap`
- Create → Blank Excel
- Fill → Headers, Styles, Locks
- Return → Final Excel File

---

#### 🔁 Process Flow – `Handle Upload`

- Input → `Excel + templateType`
- Load → `Config + Schema`
- Read → Sheet data
- Load → `processingClass`
- Call → `.process()` → `SheetDataMap`
- Merge → Processed Data
- Optional → Annotate original Excel
- Return → Final Excel File

---

### 📄 `Generator.md`

#### ⚙️ Generator  
**Role:** Supplies sample data and optional column metadata.

---

#### 🔁 Flow – `Sample Sheet`

- Input → `templateType`
- Output → `SheetDataMap`

```ts
.generate() returns:
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

- 1st Object → Column metadata (optional)
- Others → Sample data rows

---

### 📄 `Processor.md`

#### 🔬 Processor  
**Role:** Cleans, transforms, validates uploaded Excel data.

---

#### 🔁 Flow – `Validated Data`

- Input → `Excel file + templateType`
- Output → `SheetDataMap`

```ts
.process() returns:
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

- 1st Object → Column metadata (optional)
- Others → Cleaned data rows

---
