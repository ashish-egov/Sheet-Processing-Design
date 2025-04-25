#### 🧩 Unified Excel Sheet Management System  
**Multi-Sheet | Config-Driven | Interface-Based | Excel ↔ JSON ↔ Excel**

---

#### 🔁 Flows  

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
