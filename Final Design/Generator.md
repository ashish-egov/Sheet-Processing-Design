## ⚙️ Generator  
**Role:** Supplies sample data and optional column metadata for Excel sheet creation.  
Used exclusively during the **Generate Flow**.

---

### 🔁 Flow – Sample Sheet

#### 📥 Input:
- `templateType` (e.g., `"HRBulkUpload"`)

#### 📤 Output:
- A `SheetDataMap` object containing:
  - Sample rows
  - Optional column metadata for formatting and layout

#### 💡 `.generate()` returns:

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

---

### 🧩 Column Metadata (Optional)

- The **first object** in each sheet's array may contain **column metadata**, denoted by `areColumnHeaders: true`.
- This object can define:
  - `color` → Cell fill color
  - `width` → Column width
  - `isLocked` → Lock column for editing
  - `orderNumber` → Reorder columns visually
  - Other flags like `hidden`, etc.

#### 📌 Purpose:
- **Optional** → Not mandatory in every generator.
- **Override existing settings** → Can **override formatting** (e.g., widths, styles) already defined in admin schema.
- **Add new columns** → Can also introduce **additional columns** for temporary/sample purposes.

---

### 🔙 [Back to System Overview](https://github.com/ashish-egov/Sheet-Processing-Design/blob/main/Final%20Design/SystemOverview.md)
