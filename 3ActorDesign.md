# 🧩 Unified Excel Sheet Management System  
**Multi-Sheet | Config-Driven | Interface-Based | Excel ↔ JSON ↔ Excel**

We divide the system into **three actors**:

---

## 🔧 1. **TemplateManager**  
**Role:** The **brain** of the system. It coordinates all flows.

### 🔁 Word-by-Word Flow

#### For Generate Flow:
1. User provides a `templateType` (e.g., `"HRBulkUpload"`).
2. TemplateManager looks up the config for this templateType.
3. It reads MDMS Admin schema:
   - Finds sheet names, column names, formatting options, column locks, hidden flags, order numbers, etc.
4. Creates an **empty Excel file** (initial structure) based on MDMS schema:
   - Adds all defined sheets.
   - Adds placeholder headers.
   - Applies default styling, colors, and widths.
5. Loads and initializes all required `generationClass` instances for this template.
6. For each generator:
   - Calls `.generate()` to get a **SheetDataMap**.
7. Merges all returned SheetDataMaps into a single map.
8. Fills the previously created Excel file with this data:
   - Inserts column metadata (for first row override or addition).
   - Applies all styling rules from metadata.
   - Populates actual data rows.
9. Returns final Excel file to user.

#### For Process Flow:
1. User uploads an Excel file and provides a `templateType`.
2. TemplateManager looks up the config for that templateType.
3. Opens the uploaded Excel file and:
   - Extracts sheet names and raw data.
   - Reads metadata row (if present) for optional formatting info.
4. Loads and initializes all required `processingClass` instances for this template.
5. Passes each sheet’s raw data to the appropriate processor class:
6. Receives processed SheetDataMap from each processor.
7. Merges all maps into one.
8. Updates the original Excel file with:
   - Transformed rows
   - Optional column formatting
9. Returns final annotated Excel file to the user or downstream.

---

## 🧬 2. **Generator (Excel Template Generator)**  
**Role:** Generates default data + metadata for each sheet.  
It is specific to a sheet (e.g., `EmployeeSheetGenerator`).

### 🔁 Word-by-Word Flow
1. TemplateManager instantiates this class during **generate flow**.
2. `.generate()` is called with optional context.
3. The generator:
   - Constructs the first object with `areColumnHeaders: true`.
   - Adds metadata for each column:
     - Color
     - IsLocked
     - Width
     - OrderNumber
     - IsHidden (if applicable)
   - Appends optional sample data rows.
4. Returns a `SheetDataMap`:
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

5. This data is filled into the initial Excel template by the TemplateManager.

---

## ⚙️ 3. **Processor (Excel Data Processor)**  
**Role:** Validates + transforms data during Excel → JSON process.  
Each processor is sheet-specific (e.g., `EmployeeSheetProcessor`).

### 🔁 Word-by-Word Flow
1. TemplateManager instantiates this class during **process flow**.
2. `.process()` is called with:
   - Raw sheet data (object[])
   - Optional metadata (first row if present)
   - Optional context
3. Processor performs:
   - Schema-based or custom validation
   - Data transformation:
     - Trimming
     - Type casting
     - Mapping fields
     - Custom logic (e.g., value mapping, date parsing)
4. Returns a `SheetDataMap`, like:
```ts
{
  "Employees": [
    { name: "Alice", department: "HR" },
    { name: "Bob", department: "IT" }
  ]
}
```

5. TemplateManager collects and merges all SheetDataMaps.
6. It may re-annotate Excel file with validation results.

---

# 📦 Data Structure: `SheetDataMap`
```ts
type SheetDataMap = Record<string, object[]>;

// Example:
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

- First object = optional column metadata (used only for Excel generation)
- Following objects = actual data rows (used for both generate & process)

---

## 🧪 Sample Config (`SheetProcessorConfig`)
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

# 🔚 Summary Table

| Actor             | Responsibility                                          | Called During    | Output                         |
|------------------|----------------------------------------------------------|------------------|--------------------------------|
| **TemplateManager** | Coordinates everything, reads config, creates Excel    | Generate & Process | Final Excel + SheetDataMap     |
| **Generator**     | Returns sample SheetDataMap with formatting + data      | Generate Flow     | SheetDataMap for each sheet    |
| **Processor**     | Validates + transforms uploaded Excel sheet data        | Process Flow      | SheetDataMap for each sheet    |
