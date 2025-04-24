## 🧩 Unified Excel Sheet Management System  
**Multi-Sheet | Config-Driven | Interface-Based | Excel ↔ JSON ↔ Excel**

---

## 🔁 WORD-TO-WORD FLOW

### 📥 **Process Flow (Excel → JSON → Output)**

1. User uploads an Excel file  
2. System receives a **templateType** (e.g., `"HRBulkUpload"`)  
3. It looks up the **config** for that templateType  
4. It identifies all **required sheets** to process  
5. For each sheet entry in the config:  
   - Finds the **processing class name**  
6. System **loads all required processing classes** dynamically  
7. For each processing class, it calls the `.process()` method  
   - It passes:
     - The data related to that sheet
     - Context info if available  
8. Each class returns a **SheetDataMap** like this:
   ```ts
   {
     "Employees": [ ...data rows... ],
     "Departments": [ ...data rows... ]
   }
   ```
9. System collects the SheetDataMap from all processing classes  
10. It merges them into one final **SheetDataMap**  
11. This SheetDataMap is returned to the caller

---

### 🧠 Column Header Metadata (Optional First Row)

- In each sheet inside SheetDataMap, the **first object may be** a metadata row  
- This object has: `{ areColumnHeaders: true, ...column styling info }`  
- It is **not a data row**  
- It's used to control:
  - Column **width**
  - **Locking** cells
  - **Color** formatting
  - **Order** of columns
  - Whether the column is **hidden**  
- If this metadata is not present, system assumes default formatting  
- This object is **ignored in validation/transformation**, and **only used during Excel generation**

---

### 📤 **Generate Flow (Config → JSON → Excel Template)**

1. User asks to generate an Excel template for a given `templateType`  
2. System looks up config for that templateType  
3. It finds all sheet entries and their **generation class names**  
4. It dynamically loads all required generation classes  
5. For each class, it calls the `.generate()` method  
6. Each class returns a **SheetDataMap**, like:
   ```ts
   {
     "Employees": [
       {
         areColumnHeaders: true,
         Name: { color: "#CCE5FF", isLocked: true, width: 120 },
         Department: { orderNumber: 0 }
       },
       { Name: "Alice", Department: "HR" },
       { Name: "Bob", Department: "IT" }
     ]
   }
   ```
7. System collects SheetDataMap from all generators  
8. Combines all of them into one final SheetDataMap  
9. This combined SheetDataMap is used to create an Excel file

---

## 📦 STRUCTURE

### ✅ `SheetDataMap`
```ts
type SheetDataMap = Record<string, object[]>;
```

- **Key** = Sheet name (e.g., `"Employees"`)  
- **Value** = Array of row objects  
- The **first object may be metadata** with `areColumnHeaders: true`, used for formatting the Excel file  
- All other objects are data rows

---

## 🧪 SAMPLE CONFIG

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

## 📤 SAMPLE OUTPUT (`SheetDataMap`)

```ts
{
  "Employees": [
    {
      areColumnHeaders: true,
      Name: { isLocked: true, color: "#CCE5FF", width: 120, orderNumber : -2 },
      Department: { orderNumber: -1 }
    },
    { Name: "Alice", Department: "HR" },
    { Name: "Bob", Department: "IT" }
  ]
}
```

- First object = Optional metadata for Excel formatting  
- Remaining = Actual row data  
- When generating Excel, metadata is used for:
  - Styling
  - Locking columns
  - Reordering
  - Hiding columns

---
