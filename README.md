# VeriPay

## Payment Reconciliation & Verification System

VeriPay is a local-first payment reconciliation tool built to make a personal/client workflow faster and less repetitive. It imports Payment Records and GCash transaction statement workbooks, compares them by exact reference number, highlights records that need review, and creates annotated Excel outputs.

It is a private workflow tool, not a public payment processor, banking service, or SaaS product.

## Why I Built VeriPay

Reconciling payment records against transaction statements manually becomes time-consuming when there are many rows to check. VeriPay reduces that repetitive work, keeps verification runs organized, and leaves uncertain cases available for deliberate manual review.

The project is primarily a practical tool for my own workflow and work performed for a client, rather than a public commercial service.

## Features

- Payment Records XLSX import
- One or more GCash statement XLSX imports per verification run
- Exact reference-number reconciliation across the combined GCash transaction pool
- Missing-reference, reference-not-found, and duplicate-reference review states
- Safe handling for CASH, BANK, and unknown payment methods
- Manual review with an audit trail
- Matched Customer annotations for uniquely verified GCash transactions
- Verification history and workspace-scoped access
- Annotated Payment Records export
- One annotated GCash export for each uploaded GCash source workbook
- Local Supabase authentication and workspace isolation

## How Verification Works

For GCash payments, VeriPay uses the **exact transaction reference number** as the reconciliation key. It does not verify a payment from amount, customer name, dates, partial references, or fuzzy matching.

| Situation | Result |
| --- | --- |
| One exact GCash reference match across all uploaded sources | Verified |
| Missing Payment Record reference | Needs Review |
| Reference not found in any GCash source | Needs Review |
| Reference appears more than once, including across sources | Needs Review |
| CASH payment | Cash handling |
| BANK payment | Bank handling |

Leading zeros are significant. A reference with leading zeros does not match a different reference without them. Amount differences do not decide a match when one exact reference exists.

## Multiple GCash Sources

A verification run accepts one Payment Records workbook and one or more GCash statement workbooks. VeriPay evaluates reference uniqueness across the combined transaction pool while retaining the source workbook for each GCash transaction.

If the same reference appears in two GCash sources, VeriPay requires review rather than choosing an account arbitrarily. Each GCash source receives its own annotated export.

## Technology Stack

- Next.js, React, and TypeScript
- Tailwind CSS and shadcn/ui
- Supabase Auth and PostgreSQL
- ExcelJS for workbook parsing and annotated exports
- Zod for runtime validation
- Vitest for unit tests
- Playwright for browser tests
- Docker and the Supabase CLI for local services

## Local Setup

### Prerequisites

- Windows with Node.js 22.12 or newer
- pnpm 12 (the project uses Corepack)
- Docker Desktop
- Supabase CLI
- Git

Install dependencies from the repository root:

```powershell
corepack pnpm install
```

Local VeriPay needs these public environment variable names:

- `NEXT_PUBLIC_SUPABASE_URL`
- `NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY`

Use values from your **local** Supabase stack only. Keep local environment files untracked. Do not place service-role keys, passwords, or other credentials in the repository or README.

## Starting VeriPay

### Local launcher

If `VeriPay.bat` is present in your checkout:

1. Start Docker Desktop and wait for its engine to be ready.
2. Double-click `VeriPay.bat`.
3. The launcher starts local Supabase, starts VeriPay, and opens the browser after the login page is available.

Open VeriPay at `http://127.0.0.1:3000`.

### Manual start

1. Start the local Supabase stack:

   ```powershell
   supabase start
   ```

2. Set the two public environment variables above from the local `supabase status -o env` output. Do not use cloud values.

3. Start the application:

   ```powershell
   pnpm dev --hostname 127.0.0.1 --port 3000
   ```

## Stopping VeriPay

Press `Ctrl+C` in the development-server terminal to stop VeriPay.

To stop the local Supabase containers when they are no longer needed:

```powershell
supabase stop
```

`supabase db reset` is not a normal shutdown command. It recreates the local database and can remove local test data and local users.

## Using VeriPay

1. Start VeriPay and sign in.
2. Create or select the appropriate workspace.
3. Open **New Verification**.
4. Upload exactly one Payment Records XLSX workbook.
5. Upload one or more GCash statement XLSX workbooks.
6. Review each file preview and remove an accidental GCash source if needed.
7. Select **Verify Payments**.
8. Review Verified, Needs Review, Cash, and Bank outcomes.
9. Resolve manual-review cases where appropriate.
10. Download the annotated Payment Records workbook.
11. Download an annotated workbook for each GCash source.
12. Use History to revisit completed verification runs.

## Input Files

VeriPay supports `.xlsx` workbooks. Legacy `.xls` files are not supported.

- **Payment Records workbook:** customer, amount, method, reference, and other supported payment fields.
- **GCash statement workbook(s):** transaction rows containing the GCash reference and amount layout recognized by VeriPay.

Do not commit client payment files, source workbooks, generated outputs, or local environment files.

## Local-First Usage

VeriPay currently operates primarily as a local application. Docker and local Supabase provide the database and authentication environment needed for normal use. A deployment is not required for the current workflow.

## Data & Privacy

VeriPay processes payment reconciliation information and should be treated as a private workflow tool.

- Keep source workbooks and exports in a secure location.
- Do not commit customer or payment data.
- Do not commit `.env` files or credentials.
- Back up important source workbooks and verified outputs.
- Use local test data for development and automated tests.

## Development and Testing

Available project commands:

```powershell
pnpm typecheck       # TypeScript validation
pnpm lint            # ESLint with zero warnings
pnpm test            # Vitest unit tests
supabase test db     # Local database and RLS tests
pnpm test:e2e        # Playwright browser tests
pnpm build           # Production-style Next.js build
```

Browser and database tests require the local Supabase stack. Use local environment variables only.

## Project Status

VeriPay is actively developed as a personal workflow tool. The current setup is local-first; deployment may be added or changed later if operational needs require it.

## Disclaimer

VeriPay is an independent personal workflow tool. It is not affiliated with, endorsed by, or an official product of GCash or its operators.
