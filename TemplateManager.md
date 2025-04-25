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
