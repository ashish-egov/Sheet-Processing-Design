# 🧩 Unified Excel Sheet Management System

## 🎯 Goal
A single-config driven, extensible system to:

- 📥 Read Excel data
- 🔁 Transform and validate it using pluggable logic
- 📤 Recreate Excel sheets using the processed data

This system replaces separate read/write configurations and transformation logic with a **unified flow**, powered by a common interface and dynamic class resolution.

---

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

## 🏭 Factory

```ts
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

### Validation Helpers

```ts
// utils/validationHelpers.ts
export function validateRowWithAjv(row: any, schema: any): boolean {
  // Use AJV to validate
  return true;
}
```

### Abstract Base Processor

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

## ✏️ How to Add a New TemplateType

1. Implement a new processor class (`MySheetProcessor.ts`)
2. Export it and register in `SheetProcessorFactory.ts`
3. Add config in `SheetProcessorConfig.ts`

✅ Done! It's now fully supported 🔄

---

## 🆕 Schema Enhancements

### What to Add:

In the MDMS config JSON schema under `enumProperties`, `numberProperties`, and `stringProperties`, add:

```json
{
  "freezeInProcessedFile": {
    "type": "boolean"
  },
  "hideColumnInProcessedFile": {
    "type": "boolean"
  }
}
```

### What It Will Do:

#### `freezeInProcessedFile`

- **True** ➝ Freeze this column in the final processed Excel.
- Use case: Key fields like `Name`, `ID` should remain visible while scrolling.

#### `hideColumnInProcessedFile`

- **True** ➝ Hide this column in the final processed Excel.
- Use case: Internal or technical fields that should not be visible/editable by users.

---
