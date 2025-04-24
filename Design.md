## 🧩 Unified Excel Sheet Management System  
**Multi-Sheet | Config-Driven | Interface-Based | Excel ↔ JSON ↔ Excel**

---

## 🔁 WORD-TO-WORD FLOW

### 📥 **Process Flow (Excel → JSON → Output)**

1. User uploads an Excel file  
2. System receives a `templateType` (e.g., `"HRBulkUpload"`)  
3. It looks up the config for that `templateType`  
4. It identifies all required sheets to process  
5. For each sheet entry in the config:  
   - Finds the `processingClass` name  
6. System dynamically loads **all required processing classes**  
7. For each processing class, it calls the `.process()` method with:  
   - Sheet-specific input data  
   - Optional context info  
8. Each class returns a `SheetDataMap`, e.g.:
   ```ts
   {
     "Employees": [ ...rows... ],
     "Departments": [ ...rows... ]
   }
   ```
9. The system collects all SheetDataMaps from these classes  
10. It merges them into a final combined `SheetDataMap`  
11. This final output is returned to the caller

---

### 🧠 Column Header Metadata (Optional First Row)

- Inside each sheet's data, the **first object may** be a column metadata row  
- It will have:  
  ```ts
  { areColumnHeaders: true, ...column styling info }
  ```
- It is **not** a data row  
- It is used for **Excel styling**, **column ordering**, and **cell locking**
- If this object is **not** present, default formatting is used  

#### ✅ What It Does (If Present):
- Allows setting:
  - **Column width**
  - **Cell locking**
  - **Font/background color**
  - **Column order**
  - **Hidden columns**
- The system uses this during **Excel generation only**
- The object is **ignored** in data transformation or validation
- ⚠️ If the metadata includes a column not in the data rows:
  - It will be **added** as a new empty column in the Excel file  
- ⚠️ If the metadata includes a column that already exists:
  - It will **override formatting** (e.g., color, width, locked, etc.)

---

### 📤 **Generate Flow (Config → JSON → Excel Template)**

1. User requests to generate a template for a `templateType`  
2. System looks up the matching config  
3. Identifies all sheets and their `generationClass`  
4. Dynamically loads all necessary generator classes  
5. For each class, calls the `.generate()` method  
6. Each returns a `SheetDataMap` like:
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
7. System merges all SheetDataMaps  
8. Final result is converted into an Excel file and returned

---

## 📦 STRUCTURE

### ✅ `SheetDataMap`
```ts
type SheetDataMap = Record<string, object[]>;
```

- **Key** = Sheet name (e.g., `"Employees"`)  
- **Value** = Array of row objects  
- First object **may** be metadata (`areColumnHeaders: true`)  
- Remaining objects = actual data rows

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
      Name: { isLocked: true, color: "#CCE5FF", width: 120, orderNumber: -2 },
      Department: { orderNumber: -1 }
    },
    { Name: "Alice", Department: "HR" },
    { Name: "Bob", Department: "IT" }
  ]
}
```

- 🟦 First object = optional column metadata  
- ✅ Used only during Excel **generation**  
- 📝 Columns in the metadata:
  - Can **override formatting** of existing columns  
  - Can **add new columns** that weren’t in the data  


![image](https://github.com/user-attachments/assets/ccf5c18d-942d-4074-92b8-8b9c8bff7c78)
