# Context AI - Project Knowledge Memory

## Project Mission
SaaS for rapid solar estimation utilizing file ingestion and precision mathematical math runners. The application automates the ingestion of raw site telemetry (Field Study sheets) and warehouse material lists to instantly output flawless pricing models.

## Tech Stack
- Next.js 16 (App Router)
- React 19
- Tailwind CSS 4
- ExcelJS
- AI SDK (@ai-sdk/google, @ai-sdk/react)
- TypeScript

## Functional Workflows
File Upload -> UI Ingestion Context -> Agent Tool Mapping Matrix -> Math Engine Execution -> Blob Download Delivery

## Current Implementations & Changelog
- **Core Dashboard UI**: Fully implemented split-screen view in [page.tsx](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/page.tsx). Left panel features dropzones for Field Study and Material Price List Excel files and an AI Agent Chat Interface. Right panel showcases a live preview of calculated estimate summary metrics with approval states and download button.
- **AI Chat Endpoint**: Set up in [route.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/api/chat/route.ts) with `streamText` using Google Gemini, defining the `calculateEstimate` tool which returns a structured estimate including materials, labor, and margin.
- **Excel Export Endpoint**: Set up in [route.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/api/download-excel/route.ts) which delegates to `excelGenerator`.
- **Excel Sheet Generation Engine**: Set up in [excelGenerator.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/libs/excelGenerator.ts) utilizing `exceljs` to dynamically generate spreadsheet rows with totals for materials and labor.
- **Build Fixes (Imports & Typos)**: Fixed `z` import from `'zz'` to `'zod'` in [route.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/api/chat/route.ts), and corrected the import paths in [route.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/api/download-excel/route.ts) and [excelGenerator.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/libs/excelGenerator.ts) to resolve module loading failures during Next.js build.
- **AI Chat Tool Schema Typing Fix**: Modified [route.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/api/chat/route.ts) to use `inputSchema` with `zodSchema` instead of `parameters` to align with the Vercel AI SDK provider-utils typing constraints and resolve the TypeScript type checker error.
- **AI Chat Response Method Update**: Updated [route.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/api/chat/route.ts) to return `result.toUIMessageStreamResponse()` instead of `result.toDataStreamResponse()` to align with the Vercel AI SDK v6 API changes for streaming responses.
- **Excel Download NextResponse Body Fix**: Updated [route.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/api/download-excel/route.ts) to convert the `Buffer` returned by `excelGenerator` into a standard `Uint8Array` (`new Uint8Array(buffer)`) to satisfy the standard `BodyInit` type requirement of `NextResponse` and fix the TypeScript compilation error.
- **ExcelJS writeBuffer API Fix**: Modified [excelGenerator.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/libs/excelGenerator.ts) to call the correct ExcelJS method `workbook.xlsx.writeBuffer()` instead of `writeAsBuffer()`.
- **ExcelJS Promise Type Cast Fix**: Cast the return value of `workbook.xlsx.writeBuffer()` as `unknown as Promise<Buffer>` in [excelGenerator.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/libs/excelGenerator.ts) to resolve TypeScript type overlap mismatch between standard Node.js Buffer and ExcelJS Buffer.
- **Vercel AI SDK v6 React Hook Migration**: Migrated the `useChat` destructuring in [page.tsx](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/page.tsx) to match the v6 core specs. Removed non-existent `input`, `handleInputChange`, and `handleSubmit` return values, replacing them with a local React `useState` field, custom form submit handler, and standard `sendMessage({ text })` method trigger.
- **UIMessage Type Conformity Fix**: Updated the client message synchronization context in [page.tsx](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/page.tsx) to leverage the v6 `parts` text block format instead of the deprecated `content` string, resolving type mismatch errors.
- **AI Chat Custom Endpoint Route**: Setup the `useChat` initialization parameters in [page.tsx](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/page.tsx) to pass a `DefaultChatTransport` object detailing `/api/chat` as the target routing, since the custom `api` parameter was removed from the hook options.
- **High-Fidelity Slate & Emerald Dashboard Styling**: Completely overhauled the frontend styling in [page.tsx](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/page.tsx) to create a premium, gorgeous solar-themed SaaS interface, featuring slate layouts, glowing animated emerald accents, responsive prompt suggestions, drag-and-drop file indicators, and live itemized price calculation data grids.
- **ExcelJS API numFmt Property Update**: Modified [excelGenerator.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/libs/excelGenerator.ts) to apply the valid `.numFmt` property on sheet cell instances instead of `.numFormat` to comply with standard typing.
- **ExcelJS Literal Border Styling Casts**: Applied explicit `as const` type assertions on spreadsheet border style literal assignments in [excelGenerator.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/libs/excelGenerator.ts) to adhere strictly to ExcelJS `BorderStyle` union type expectations.
- **Excel Premium Branding & Layout Enhancements**: Upgraded the report generation template in [excelGenerator.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/libs/excelGenerator.ts) to render a top merged company typographic logo banner, zebra rows, double-underlined emerald green grand totals, and proper cell alignments.
- **AI SDK v6 Message Format Validation Fix**: Resolved a server-side Zod validation mismatch (`ZodError` checking roles/content) by importing and invoking `convertToModelMessages(messages)` in [route.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/api/chat/route.ts) to convert frontend `UIMessage[]` structures to backend `CoreMessage[]` structures before passing them to `streamText`.
- **Model Upgrade to Gemini 2.5 Flash**: Switched the LLM model identifier from `'gemini-1.5-pro'` to `'gemini-2.5-flash'` in [route.ts](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/api/chat/route.ts) to resolve API `NOT_FOUND` errors since the current API key environment explicitly supports the advanced `'models/gemini-2.5-flash'` family.
- **Offline Build Font Optimization**: Replaced network-dependent Google Fonts (`next/font/google` in [layout.tsx](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/layout.tsx)) with custom local UI font stacks (`Inter`, `Segoe UI`, Roboto) declared via CSS variables in [globals.css](file:///c:/Users/PARFAIT/Desktop/quick-estimator/src/app/globals.css) to eliminate build-time font fetching network requests and fix compilation failures in sandboxed/offline environments.
