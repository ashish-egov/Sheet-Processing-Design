## 🔬 Processor  
**Role:** Transforms, validates, and optionally annotates uploaded Excel data.  
It acts as the **processing engine** in the **Process Flow**.

---

### 🔁 Flow – Validated Data

#### 📥 Input:
- Uploaded **Excel file**
- `templateType` (e.g., `"HRBulkUpload"`)

#### 📤 Output:
- A `SheetDataMap` containing:
  - Validated and transformed row data
  - Optional column metadata for styling or overrides

---

### 💡 `.process()` returns:

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

---

### 🧩 Column Metadata (Optional)

- The **first object** in each sheet’s array can include **column metadata**, marked by `areColumnHeaders: true`.
- Used to define:
  - `width` → Column width
  - `color` → Background color
  - Other properties like `isLocked`, `hidden`, etc.

#### 📌 Purpose:
- **Optional** → Not required in every processor
- **Override** → Admin schema styles or settings
- **Annotate** → Highlight errors, feedback with formatting
- **Add Columns** → Derived/calculated columns for enriched output

---

### 🛠️ Processing Capabilities

Processors can handle a variety of tasks:

- ✅ **Validation** → Check schema rules, required fields, data types, patterns  
- 🔄 **Transformation** → Format conversion, value mapping, string normalization  
- ➕ **Creation** → Add computed or derived fields  
- 📝 **Annotation** → Style cells, add tooltips or highlights for error feedback  

---

### 🔙 [Back to System Overview](https://github.com/ashish-egov/Sheet-Processing-Design/blob/main/Final%20Design/SystemOverview.md)
