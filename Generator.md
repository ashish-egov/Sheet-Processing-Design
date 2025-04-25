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
