# 🧩 Unified Excel Sheet Management System

## 🎯 Goal

A single-config driven, extensible system to:

- 📥 Read Excel data
- 🔁 Transform and validate it using pluggable logic
- 📤 Recreate Excel sheets using the processed data

> ✅ This system unifies the read/write flow using dynamic class resolution and a common interface.

---

## 🏗️ High-Level Architecture

```plaintext
             ┌────────────────────────┐
             │  SheetProcessorConfig  │
             └────────────┬───────────┘               
                          ▼                            
               ┌───────────────────────┐               
               │  Excel → JS Rows      │               
               └─────────┬─────────────┘               
                         ▼                             
       ┌──────────────────────────────────┐            
       │ Inject Context via sheetMapping  │◄────┐      
       └────────────────────┬─────────────┘     │       
                            ▼                   │       
               ┌───────────────────────┐        │       
               │   Get Class from      │        │       
               │   processingClass     │        │       
               └────────┬──────────────┘        │       
                        ▼                       │       
       ┌──────────────────────────────────┐     │       
       │  processor.process(rows, ctx)    │─────┘       
       └────────────────┬────────────────┘             
                        ▼
              ┌────────────────────┐
              │ 2D Array of Strings│
              └─────────┬──────────┘
                        ▼
              ┌────────────────────┐
              │   Write to Excel   │
              └────────────────────┘
```

---

## ⚙️ Config Design

```ts
export interface SheetProcessorConfigEntry {
  templateType: string;
  sheetName: string;
  schemaName: string;
  processingClass: string;
  validateUsingAjv: boolean;
  parsingIdentifier: string;
  sheetDataMapping: SheetDataMapping[];
}

export interface SheetDataMapping {
  inJsonPath: string;
  outJsonPath: string;
}

export const SheetProcessorConfig: SheetProcessorConfigEntry[] = [
  {
    templateType: "User",
    sheetName: "UserDetails",
    schemaName: "UserTemplate",
    processingClass: "UserSheetProcessor",
    validateUsingAjv: true,
    parsingIdentifier: "user-upload",
    sheetDataMapping: [
      { inJsonPath: "$.user.name", outJsonPath: "$.context.userNames" },
      { inJsonPath: "$.user.status", outJsonPath: "$.context.userStatus" }
    ]
  }
];
```

---

## 🧩 Common Interface

```ts
export interface SheetProcessorInterface {
  process(
    rows: Record<string, any>[],
    context: Record<string, any>
  ): Promise<string[][]>;
}
```

---

## 🛠️ Sample Processor

```ts
import { SheetProcessorInterface } from "./SheetProcessorInterface";

export class UserSheetProcessor implements SheetProcessorInterface {
  async process(rows: Record<string, any>[], context: Record<string, any>): Promise<string[][]> {
    return rows.map(row => {
      const fullName = `${row.firstName} ${row.lastName}`;
      const status = context?.userStatus?.[row.userId] || row.status;
      return [fullName, status];
    });
  }
}
```

---

## 🏭 Processor Factory

```ts
import { SheetProcessorInterface } from "./SheetProcessorInterface";
import { UserSheetProcessor } from "./UserSheetProcessor";
import { BoundarySheetProcessor } from "./BoundarySheetProcessor";

const processorRegistry: Record<string, new () => SheetProcessorInterface> = {
  UserSheetProcessor,
  BoundarySheetProcessor
};

export function getSheetProcessor(className: string): SheetProcessorInterface {
  const ProcessorClass = processorRegistry[className];
  if (!ProcessorClass) throw new Error(`Processor class not found: ${className}`);
  return new ProcessorClass();
}
```

---

## 🚀 Execution Flow

```ts
import { SheetProcessorConfig } from "./SheetProcessorConfig";
import { getSheetProcessor } from "./SheetProcessorFactory";

export async function runExcelProcessing(
  templateType: string,
  excelRows: Record<string, any>[],
  contextData: Record<string, any>
): Promise<string[][]> {
  const config = SheetProcessorConfig.find(cfg => cfg.templateType === templateType);
  if (!config) throw new Error(`No config found for ${templateType}`);

  const processor = getSheetProcessor(config.processingClass);
  const outputData = await processor.process(excelRows, contextData);

  return outputData;
}
```

---

## ✅ Optional Extensions

### AJV Validation Helpers

```ts
export function validateRowWithAjv(row: any, schema: any): boolean {
  return true; // Use AJV logic
}
```

---

### Abstract Base Class (Optional)

```ts
export abstract class BaseSheetProcessor implements SheetProcessorInterface {
  abstract process(
    rows: Record<string, any>[],
    context: Record<string, any>
  ): Promise<string[][]>;

  protected validateWithAjv(row: Record<string, any>, schema: any): boolean {
    return true;
  }
}
```

---

## 📁 Directory Structure

```
src/
├── processors/
│   ├── SheetProcessorInterface.ts
│   ├── UserSheetProcessor.ts
│   ├── BoundarySheetProcessor.ts
│   └── BaseSheetProcessor.ts
├── config/
│   └── SheetProcessorConfig.ts
├── factory/
│   └── SheetProcessorFactory.ts
├── service/
│   └── runExcelProcessing.ts
└── utils/
    └── validationHelpers.ts
```

---

## ✏️ Add a New Template

1. Create `MySheetProcessor.ts`
2. Export & register in `SheetProcessorFactory.ts`
3. Add config entry in `SheetProcessorConfig.ts`

> 🎉 That's it — plug-and-play!

---

## 🆕 MDMS Config Enhancements

### JSON Schema Updates

For each column under `stringProperties`, `numberProperties`, and `enumProperties`, extend the schema like this:

```json
{
  "freezeInProcessedFile": {
    "type": "boolean"
  },
  "hideColumnInProcessedFile": {
    "type": "boolean"
  },
  "columnColorInProcessedFile": {
    "type": "string",
    "pattern": "^#([A-Fa-f0-9]{6})$"
  },
  "columnWidth": {
    "type": "number",
    "minimum": 10,
    "maximum": 500
  }
}
```

---

### 🧠 Behavior

- **`freezeInProcessedFile`**
  - ➤ Keeps the column visible (pinned) while scrolling.
  - 🔒 Good for columns like IDs or names.

- **`hideColumnInProcessedFile`**
  - ➤ Hides this column in the final Excel output.
  - 🔐 Useful for internal-only fields.

- **`columnColorInProcessedFile`**
  - ➤ Applies a background fill to the column.
  - 🎨 Accepts hex format like `#FFDD00`, `#00BFFF`.

- **`columnWidth`**
  - ➤ Sets column width in Excel.
  - 📏 Typical values range from 20 to 50, max is 500.
