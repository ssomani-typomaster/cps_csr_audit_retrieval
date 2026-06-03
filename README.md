# CPS Excel-Driven CSR Retrieval & Validation Tool

## Overview
This tool was created to automate a time-consuming weekly task for **CITI** however can used by any accounts. It eliminates manual downloading of the CSR by downloading both ECDSA and RSA in pem and text format. 

The script automatically executes structural decoding checks to validate parse results (`PASS`/`FAIL`), extracts Common Names (CN), maps Subject Alternative Names (SANs), and tracks granular account metrics into an aggregated execution summary.

---

## Features
* **Dual-Algorithm Support:** Concurrent retrieval and validation of paired ECDSA and RSA cryptographic keys.
* **Automated SAN Assessment:** Decodes raw PEM blocks on the fly to count and list every compiled Subject Alternative Name.
* **Smart Local Caching:** Caches API enrollment sets dynamically per `accountSwitchKey` to minimize transactional latency and limit API throttling.
* **Auto-Formatted Reporting:** Exports a cleanly styled, color-coded Excel summary spreadsheet with auto-adjusted tracking matrices.

---

## Technical Architecture & Workflow

1. **Input Stage:** The script seeks an input file named `csr_generator.xlsx` containing a sheet named `CSR Generator`.
2. **API Handshake:** Uses local EdgeGrid configurations to initialize signed sessions targeting Akamai endpoint matrices.
3. **Parsing Engine:** Utilizes cryptographic x509 bindings to map local PEM contexts.
4. **Isolating Outputs:** Generates a unique execution run environment matching the exact tracking signature `csr_run_YYYY-MM-DD_HH-MM-SS/`.

---

## Prerequisites

### 1. Local EdgeGrid Credentials
The tool requires an active Akamai API identity file situated at your home root: `~/.edgerc`. The file must contain a `[default]` profile token layer with functional read privileges for the CPS API:

```ini
[default]
client_secret = xxxx
host = xxxx.luna.akamaiapis.net
access_token = xxxx
client_token = xxxx
