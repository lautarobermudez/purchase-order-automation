# Purchase order → payment order automation

[![tests](https://github.com/lautarobermudez/purchase-order-automation/actions/workflows/ci.yml/badge.svg)](https://github.com/lautarobermudez/purchase-order-automation/actions/workflows/ci.yml)

Python scripts that turn purchase orders (spreadsheet or PDF), invoices and tax bills into **payment orders (OP)** for the accounts-payable team of a small Argentine group of companies, and that watch an inbox so most of them are generated without manual typing.

I built and run this at work. The company, supplier, bank and employee data in this public copy is replaced with placeholders (`ACME`, `UNIT_B`, `SUPPLIER_X`, `example.com`), and the company's official OP template is not included, so **it will not run as-is**. It is published to show the design and the code.

> **Built with AI assistance.** I wrote the requirements and the business rules, ran it in production and reviewed the output, using Claude Code to write most of the code. I can explain how each module works, but I would not call it hand-written.

## What it does
| Module | Purpose |
|---|---|
| `generar_ops.py` | Reads a purchase order (xlsx/PDF-derived JSON), fills the OP template (XML-level xlsx editing with the standard library), numbers the OPs, groups attachments and can email them. Dry-run by default; nothing is written without `--apply`. |
| `revisor_mail.py` | Reads the mailbox over IMAP in **read-only** mode, detects payment requests from authorized senders, runs a validation gate and generates the OP. It never deletes, moves or marks emails, and it does not email recipients until a human approves with `--enviar`. |
| `lector_facturas.py` | For invoices/bills that arrive **without** a purchase order: asks Claude (Claude Code in non-interactive mode, `Read` tool only) to extract the data as JSON, then applies plain-code checks. |

## Design decisions worth noting
- **Dry-run first.** Every command simulates unless `--apply` is passed.
- **Validation gate before generating anything:** supplier and amount present, bank details unless cash, known paying entity, breakdown adds up to the total, valid CUIT check digit, and no duplicate invoice (same CUIT + invoice number, or same CUIT + amount + date).
- **Treat invoice content as untrusted data.** The model gets only the `Read` tool (no shell, no network), its output is parsed as JSON, and everything passes through the gate. Invoices may contain text that tries to instruct the model; it is ignored.
- **Cost control:** a per-run cap, and an identical file (by hash) is read only once. It can be switched off in the config.
- **Human in the loop:** OPs wait for approval before being emailed; printing is idempotent (each OP is printed once).
- **Secrets in environment variables**, never in the repo (`revisor_config.json`, state files and `.env` are git-ignored; see `revisor_config.ejemplo.json`).

## Stack
Python 3 (standard library only: `imaplib`, `smtplib`, `zipfile`, `xml`, `subprocess`), Claude Code CLI for the invoice-reading step. It uses `fcntl`, so it runs on Linux/WSL.

## Tests
The validation rules are covered by 46 tests (standard library `unittest`, no dependencies), and CI runs them on every push:

```bash
python -m unittest discover -s tests -t . -v
```

## Known limitations
- Code and comments are in Spanish.
- Tests cover the validation rules (CUIT check digit, invoice-number parsing, the gate, duplicate detection, OP numbering), but not the IMAP, SMTP, printing or Claude-reading steps, which were checked with real documents and dry-runs.
- `generar_ops.py` and `revisor_mail.py` are long single files and should be split into modules.
- Tied to one OP template layout.
