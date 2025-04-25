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
