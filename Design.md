# 🧩 Unified Excel Sheet Management System

## 🎯 Goal

A single-config driven, extensible system to:

1.  **Read** Excel data
    
2.  **Transform and validate** it using pluggable logic
    
3.  **Recreate Excel sheets** using the processed data
    

This system replaces separate read/write configurations and transformation logic with a **unified flow**, powered by a common interface and dynamic class resolution.

----------

## 🏗️ High-Level Architecture

```
             ┌────────────────────────┐
             │  SheetProcessorConfig  │◄───────────────┐
             └────────────┬───────────┘                │
                          ▼                            │
               ┌───────────────────────┐               │
               │  Excel → JS Rows      │               │
               └─────────┬─────────────┘               │
                         ▼                             │
       ┌──────────────────────────────────┐            │
       │ Inject Context via sheetMapping  │◄────┐       │
       └────────────────────┬─────────────┘     │       │
                            ▼                   │       │
               ┌───────────────────────┐        │       │
               │   Get Class from      │        │       │
               │   processingClass     │        │       │
               └────────┬──────────────┘        │       │
                        ▼                       │       │
       ┌──────────────────────────────────┐     │       │
       │  processor.process(rows, ctx)    │─────┘       │
       └────────────────┬────────────────┘             ▼
                        ▼
              ┌────────────────────┐
              │ 2D Array of Strings│
              └─────────┬──────────┘
                        ▼
              ┌────────────────────┐
              │   Write to Excel   │
              └────────────────────┘

```

----------

## ⚙️ Config Design: `SheetProcessorConfig`

```ts
export interface SheetProcessorConfigEntry {
  templateType: string;
  sheetName: string;
  schemaName: string;
  processingClass: string; // Class implementing SheetProcessorInterface
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
  },
  // Add more template configs as needed
];

```

----------

## 🧩 Common Interface: `SheetProcessorInterface`

```ts
export interface SheetProcessorInterface {
  process(
    rows: Record<string, any>[],
    context: Record<string, any>
  ): Promise<string[][]>;
}

```

----------

## 🛠️ Sample Implementation: `UserSheetProcessor.ts`

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

----------

## 🏭 Processor Resolver: `SheetProcessorFactory.ts`

```ts
import { SheetProcessorInterface } from "./SheetProcessorInterface";
import { UserSheetProcessor } from "./UserSheetProcessor";
import { BoundarySheetProcessor } from "./BoundarySheetProcessor"; // Add other imports

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

----------

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

----------

## 🧪 Optional Extensions

### ✅ Validation Helpers

-   Add `validateRowWithAjv(row, schemaName)` in a utility file.
    
-   Call this inside `process()` if `config.validateUsingAjv` is true.
    

### 📊 Abstract Base Processor (Optional)

```ts
export abstract class BaseSheetProcessor implements SheetProcessorInterface {
  abstract process(
    rows: Record<string, any>[],
    context: Record<string, any>
  ): Promise<string[][]>;

  protected validateWithAjv(row: Record<string, any>, schema: any): boolean {
    // AJV validation logic
    return true;
  }
}

```

### 📁 Directory Structure Suggestion

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

----------

## 📘 How to Add a New TemplateType

1.  Implement a class like `MySheetProcessor.ts`
    
2.  Export it and add it to the factory registry
    
3.  Add a config entry in `SheetProcessorConfig`
    

That's it! The system will auto-handle the rest 🧠






---

### ✅ **What to Change:**

1. **Add `freezeInProcessedFile` (boolean)**  
   To each item in:
   - `stringProperties`
   - `numberProperties`
   - `enumProperties`

2. **Add `hideColumnInProcessedFile` (boolean)**  
   To each item in:
   - `stringProperties`
   - `numberProperties`
   - `enumProperties`

---

### 🔍 **What It Will Do:**

- **`freezeInProcessedFile`**  
  ➤ When set to `true`, this column will be **frozen (locked/pinned)** in the final processed Excel file.  
  This is useful for key columns like "ID", "Name", etc., so users can scroll while keeping them visible.

- **`hideColumnInProcessedFile`**  
  ➤ When set to `true`, this column will be **hidden** in the final processed Excel file.  
  This is useful for internal codes, metadata, or technical fields that end users shouldn't see or edit.

