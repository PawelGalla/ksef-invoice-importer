# KSeF Invoice Importer

An Excel / VBA tool that turns purchase invoices downloaded from **KSeF** (Krajowy System e-Faktur, the Polish National e-Invoicing System) as XML files into a clean Excel list, **without duplicates**.

Drop the XML files in a folder, click **Import invoices**, and every new invoice lands on a list with its number, seller, amounts split by VAT rate, bank account, dates, payment status and settlement method. Corrections go to a separate sheet. Invoices that are already on the list are skipped.

## Features

- Reads both KSeF logical structures: **FA(2)** and **FA(3)**. It does not depend on the XML namespace, so it keeps working when the schema version changes.
- **No duplicates.** An invoice is identified by *seller NIP + invoice number*, checked across both the Invoices and Corrections sheets. Files imported earlier are skipped silently.
- **Append-only.** New invoices are added at the bottom. Existing rows, and any columns you add on the right (comments, approvals, etc.), are never touched.
- **Corrections** (`KOR`, `KOR_ZAL`, `KOR_ROZ`) go to a separate sheet, with the correction reason and type and a reference to the corrected invoice.
- **Payment status** (`paid` / `partial` / `pending`), plus the **amount outstanding**.
- **Settlement method** (`bank transfer` / `cash` / `offset` / `other`).
- **Bank account** without spaces, with the country code in a separate column.
- Net / VAT amounts **split by VAT rate**, plus net, VAT and gross totals.
- A **Log** sheet that records every processed file: added, duplicate, skipped or error.
- Self-healing workbook layout. Missing sheets are created automatically, so the `.bas` module also works on its own in an empty workbook.

## Quick start

1. Download [`KSeF_Invoice_Importer.xlsm`](KSeF_Invoice_Importer.xlsm).
2. Windows blocks macros in files downloaded from the internet. Right-click the file, choose **Properties**, tick **Unblock**, then click **OK**.
3. Open the file and click **Enable Content**.
4. On the **Start** sheet, set the folder that holds the KSeF XML files (default: `C:\xml_FAKTURY`).
5. Optionally enter your own NIP. Invoices where you are the seller will then be skipped.
6. Click **Import invoices**.

Try it with the fictitious invoices in [`samples/`](samples).

### Using only the VBA module

1. Open any workbook and press **Alt+F11**.
2. Choose **File → Import File…** and select [`src/KsefInvoiceImporter.bas`](src/KsefInvoiceImporter.bas).
3. Run the macro `ImportKsefInvoices` (Alt+F8).

The *Start*, *Invoices*, *Corrections* and *Log* sheets, including an import button, are created on the first run. Save the workbook as `.xlsm`.

## Output columns

**Invoices** sheet (columns A–AP):

| Group | Columns |
|---|---|
| Identification | Invoice number, KSeF number, Invoice type, Seller NIP, Seller name |
| Dates | Issue date, Sale / delivery date, Service period from, Service period to, Payment due date |
| Payment | Settlement method, Bank account country, Bank account number, Payment status, Payment date |
| Totals | Currency, Net total, VAT total, Gross total, Amount outstanding |
| By VAT rate | Net/VAT 23%, Net/VAT 8%, Net/VAT 5%, Net/VAT 4% (taxi), Net/VAT OSS, 0% domestic, 0% intra-EU supply, 0% export, Exempt, Not subject to VAT, Not subject – art. 100(1)(4), Reverse charge, Margin scheme |
| Technical | VAT in PLN (from invoice), XML created at, File name, Imported at |

**Corrections** sheet (columns A–AU) has the same columns. Five more are inserted after *Seller name*: Correction reason, Correction type, Corrected invoice number, Corrected invoice issue date, Corrected invoice KSeF number.

## Business rules

| Field | Rule |
|---|---|
| **Duplicate** | Same seller NIP (digits only) and same invoice number (case- and whitespace-insensitive). |
| **Payment status** | `paid` if the invoice is marked as paid (`Zaplacono = 1`), or if partial payments cover the gross total. `partial` if partial payments are listed (`ZnacznikZaplatyCzesciowej`, `ZaplataCzesciowa`). `pending` otherwise. |
| **Amount outstanding** | Gross total minus partial payments. Zero when paid. Negative for unpaid corrections, meaning money is due back to you. |
| **Payment date** | `DataZaplaty`, or the latest partial-payment date. |
| **Settlement method** | KSeF code 6 → `bank transfer`. Code 1 → `cash`. Codes 2–5 and 7 (card, voucher, cheque, credit, mobile) → `other`. For "other payment" (`PlatnoscInna`), the description decides: *kompensata / potrącenie / offset / netting* → `offset`, *przelew / transfer* → `bank transfer`, *gotówka / cash* → `cash`, anything else (PayU, BLIK, …) → `other`. **No payment method on the invoice → `bank transfer`.** |
| **Bank account** | Spaces, dashes and an optional `IBAN` prefix are removed. A leading two-letter prefix (e.g. `PL`, `DE`) goes to *Bank account country*. Without a prefix, or with no account at all, the country is `PL`. If the invoice lists several accounts, the first one is used. |
| **Payment due date** | The first `Termin`. A descriptive term (e.g. "14 dni od doręczenia") is shown as text. |
| **KSeF number** | Taken from the file name (`NIP-YYYYMMDD-XXXXXXXXXXXX-XX.xml`), because the invoice XML itself does not contain it. |

## Notes and limitations

- The payment status reflects the invoice **at import time**. Existing rows are never updated, so mark later payments manually, or add your own column for them.
- Amounts are shown in the invoice currency. No exchange-rate conversion is done. *VAT in PLN* is filled only when the invoice provides it (`P_14_xW`).
- Do not rename the sheets or reorder the built-in columns, because the macro relies on their positions. You can safely add your own columns to the right.
- Requires **Excel for Windows**. It uses `MSXML2.DOMDocument.6.0`, `Scripting.FileSystemObject`, `Scripting.Dictionary` and `VBScript.RegExp`, all built into Windows.

## Repository structure

```
KSeF_Invoice_Importer.xlsm      ready-to-use workbook (button + macro)
src/KsefInvoiceImporter.bas     VBA source code
samples/                        fictitious KSeF invoices for testing
```

## License

[MIT](LICENSE)
