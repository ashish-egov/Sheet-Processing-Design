## 🧠 TemplateManager  
**Role:** Central brain that orchestrates both `Generate` and `Process` flows.  
**Responsibility:** Coordinates config loading, class resolution, Excel creation/updation, and final file generation.

---

### 🔁 Generate Flow → **Create Template**

#### 📥 Input:
- `templateType` (e.g., `"HRBulkUpload"`)

#### ⚙️ Internal Steps:
1. Load MDMS Config (sheet structure, schema, styling, etc.)
2. Dynamically load the associated `generationClass`
3. Call `.generate()` → returns `SheetDataMap`
4. Create a blank Excel with all necessary sheets
5. Fill headers, column metadata, sample rows
6. Apply column widths, styling, locks, ordering

#### 📤 Output:
- Fully formatted **Excel file** with:
  - Sheet structure as per config
  - Column metadata (optional)
  - Sample data (if provided)

---

### 🔁 Process Flow → **Handle Upload**

#### 📥 Input:
- Uploaded **Excel file**
- `templateType` (to determine config/schema)

#### ⚙️ Internal Steps:
1. Load MDMS Config and validation schema
2. Parse and read Excel data into raw row objects
3. Dynamically load the associated `processingClass`
4. Call `.process()` → returns `SheetDataMap`
5. Merge processed data with Excel
6. Optionally annotate cells with errors, colors, comments, etc.

#### 📤 Output:
- Updated or annotated **Excel file** (same as input file but enriched)
- Optionally, cleaned **JSON data** if used in downstream processing

---

### ✅ Summary Table

| Flow           | Input                          | Output                      |
|----------------|--------------------------------|-----------------------------|
| **Generate**   | `templateType`                 | New, formatted Excel file   |
| **Process**    | Uploaded Excel + `templateType`| Annotated Excel file        |

---
