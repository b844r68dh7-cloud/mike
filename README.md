# Mike — Extended Legal Practice Edition

![Mike](docs/assets/link-image.jpg)

[![Upstream PRs](https://img.shields.io/badge/Upstream_PRs-%23878--%23882-2563eb)](https://github.com/open-legal-products/mike/pulls)
[![OCR](https://img.shields.io/badge/Scanned_PDF_%26_Image_OCR-Mistral_%7C_OpenAI_%7C_Local_VLM-059669)](#1-multi-provider-ocr-for-scanned-pdfs--images-prs-882--881)
[![Hybrid RAG](https://img.shields.io/badge/Hybrid_RAG-Zero_pgvector_Dependency-7c3aed)](#2-extension-free-hybrid-semantic-search-rag-pr-880)
[![Office Add-in](https://img.shields.io/badge/Office_Add--in-Word_%2B_Outlook_Compose-d97706)](#3-microsoft-outlook-compose-task-pane-pr-878)

> **What is this fork?**
> This repository tracks upstream [`open-legal-products/mike`](https://github.com/open-legal-products/mike) and adds **five production-hardened extensions** built for real-world **M&A Due Diligence**, **Litigation**, and **International Arbitration** workflows—where bundles are full of scanned PDFs, concepts span thousands of pages across languages, and counsel work equally in **Microsoft Word** and **Microsoft Outlook**.
>
> **If these capabilities are useful to your practice, please ⭐ Star this repository** and share your use case on the upstream pull requests ([#882](https://github.com/open-legal-products/mike/pull/882), [#881](https://github.com/open-legal-products/mike/pull/881), [#880](https://github.com/open-legal-products/mike/pull/880), [#879](https://github.com/open-legal-products/mike/pull/879), [#878](https://github.com/open-legal-products/mike/pull/878))—the upstream maintainers actively look at community traction when deciding which extensions to merge into core.

---

## Why This Extended Edition Exists (Upstream vs. Extended Fork)

| Capability | Upstream `open-legal-products/mike` | This Extended Fork (`b844r68dh7-cloud/mike`) | Upstream PR |
| :--- | :--- | :--- | :--- |
| **Scanned PDFs & Court/Arbitral Exhibits** | Indexed as empty/textless pages; invisible to search and chat unless pre-OCR'd on the user's laptop | **Automatic server-side OCR** exclusively on textless pages during ingestion (`mistral-ocr` or any OpenAI-compatible Vision API: Cortecs, vLLM, Ollama, OpenAI, Azure) | [#882](https://github.com/open-legal-products/mike/pull/882) |
| **Raster Image Uploads (`PNG` / `JPEG`)** | Rejected by upload allowlist | **First-class `PNG`/`JPEG` document uploads** with magic-byte validation, automatic OCR transcription, and PDF rendition | [#881](https://github.com/open-legal-products/mike/pull/881) |
| **Document Retrieval across Large Bundles** | Literal substring match only (`find_documents`) | **Hybrid Semantic + Lexical RAG (`search_documents_semantic`)** using Reciprocal Rank Fusion (RRF) and native PostgreSQL `real[]` arrays (**zero `pgvector` requirement**) | [#880](https://github.com/open-legal-products/mike/pull/880) |
| **Office Add-in Host Support** | Microsoft Word only | **Unified Word + Outlook Compose add-in**: reads active email draft, quoted thread, and attachments into context, and inserts drafted replies at cursor or top | [#878](https://github.com/open-legal-products/mike/pull/878) |
| **Visual Diagrams & Illustrations** | Text/Office documents only | **Opt-in `generate_image` tool** (OpenAI, Azure, Automatic1111, or custom gateway) persisted as a versioned project document | [#879](https://github.com/open-legal-products/mike/pull/879) |

---

## Highlights of the Five Extensions

### 1. Multi-Provider OCR for Scanned PDFs & Images (PRs [#882](https://github.com/open-legal-products/mike/pull/882) & [#881](https://github.com/open-legal-products/mike/pull/881))
In M&A data rooms and dispute resolution productions, a large fraction of documents are **scanned paper PDFs** (signed contracts, notarized deeds, court exhibits, historical corporate books) or photographed pages (`PNG`/`JPEG`).
- **Why local pre-OCR fails in practice**: Asking lawyers or paralegals to run local OCR tools before uploading breaks the ingestion workflow and yields inconsistent quality depending on client hardware.
- **How our pipeline works**: When `OCR_ENABLED=true`, Mike inspects extracted PDF pages, preserves all born-digital native text untouched, and requests OCR **only for textless pages**. Both native Mistral OCR (`POST /v1/ocr`) and any OpenAI-compatible `/v1/chat/completions` vision model (including EU sovereign gateways like Cortecs, local vLLM/Ollama, OpenAI, OpenRouter, and Azure) are supported with strict credential host-isolation.

### 2. Extension-Free Hybrid Semantic Search / RAG (PR [#880](https://github.com/open-legal-products/mike/pull/880))
Literal keyword search misses clauses when opposing counsel or foreign-language exhibits phrase concepts differently (e.g., searching for *"waiver of set-off"* or *"clausola compromissoria"* across multilingual productions).
- **Zero `pgvector` dependency**: Embeddings are stored in standard PostgreSQL `real[]` float arrays and ranked with pure-SQL cosine distance fused with PostgreSQL `tsvector` lexical ranking via Reciprocal Rank Fusion (RRF). It runs out of the box on **any** PostgreSQL instance.
- **Strict tenancy & lifecycle guarantees**: Revision-bound chunks and status rows are automatically cleaned up on soft-delete, version replacement, or storage removal, and background maintenance sweeps automatically backfill and recover interrupted jobs.

### 3. Microsoft Outlook Compose Task Pane (PR [#878](https://github.com/open-legal-products/mike/pull/878))
The same `word-addin` React app now detects when it is hosted inside **Microsoft Outlook** (`Office.HostType.Outlook`) during message composition:
- Reads the active draft's subject, recipients, body/quoted thread, and file attachments (with attachment-list integrity checks and `<untrusted-content>` prompt fencing).
- Lets counsel draft or refine email responses using project documents and insert the result at the cursor or top of the message with one click.

### 4. Opt-in Visual Diagram & Image Generation (PR [#879](https://github.com/open-legal-products/mike/pull/879))
Adds a provider-agnostic `generate_image` tool for timelines, corporate structure visuals, or illustrative exhibits, supporting OpenAI, Azure OpenAI, Automatic1111, and custom gateways.

---

## Quick Configuration for Extended Features (`backend/.env`)

All five extensions are **opt-in and off by default** so existing deployments behave identically until configured:

```ini
# --- 1. Scanned PDF & Image OCR ---
OCR_ENABLED=true
OCR_PROVIDER=openai-compatible
OCR_BASE_URL=https://api.cortecs.ai/v1
OCR_API_KEY=your-api-key
OCR_MODEL=allenai/olmOCR-7B-0225-preview

# --- 2. Hybrid Semantic Search (RAG) ---
RAG_ENABLED=true
EMBEDDING_PROVIDER=openai-compatible
EMBEDDING_BASE_URL=https://api.cortecs.ai/v1
EMBEDDING_API_KEY=your-api-key
EMBEDDING_MODEL=qwen3-embedding-8b

# --- 3. Assistant Image Generation (optional) ---
IMAGE_GEN_PROVIDER=openai
```

---

![Mike](docs/assets/link-image.jpg)

Mike (MikeOSS) is an open-source legal AI platform for document review,
drafting, and legal research.

It combines a Next.js frontend, an Express backend, Supabase Auth/Postgres,
and Cloudflare R2-compatible object storage.

Website: [mikeoss.com](https://mikeoss.com)

![Mike assistant home screen](docs/assets/mike-home.png)

## Features

- Chat with legal documents and open matters
- Review documents and apply suggested edits
- Run reusable assistant and tabular-review workflows
- Organize projects, folders, and a document library
- Verify citations and research US case law with CourtListener
- Work from Microsoft Word with the task-pane add-in
- Run supported language models locally through Ollama

## Quick start

The included Docker Compose stack runs Mike, Supabase, RustFS object storage,
and local email capture without requiring managed infrastructure.

1. Copy the local environment templates:

   ```bash
   cp .env.example .env
   cp backend/.env.example backend/.env
   ```

2. In `backend/.env`, set `DOWNLOAD_SIGNING_SECRET` and
   `USER_API_KEYS_ENCRYPTION_SECRET` to separate values generated with:

   ```bash
   openssl rand -hex 32
   ```

3. Add an Anthropic, Gemini, or OpenAI API key to `backend/.env`, unless you
   plan to use Ollama exclusively.

4. Start the stack. Either run the published images (built once in CI, with
   their Sentry source maps and release) by pinning a version in `.env`:

   ```bash
   echo "MIKE_IMAGE_TAG=$(cat VERSION)" >> .env
   docker compose pull && docker compose up --no-build -d
   ```

   or build everything from this checkout:

   ```bash
   docker compose up --build
   ```

5. Open [http://localhost:3000](http://localhost:3000) and create an account.

To stop, run `npm run stack:down` rather than `docker compose stop`: stopped
containers are never removed on their own and keep images and volumes in use.
See "Stop and clean up" in `docs/local-development.md`.

The bundled credentials and infrastructure are intended for local development
only. See [Local development](docs/local-development.md) for service endpoints,
authentication behavior, Ollama setup, and first-run guidance.

### Telemetry

Mike sends scrubbed error reports (no document content, user identities, or
request data) to the project's Sentry by default. Development builds (`next dev`,
the backend run from source, the add-in's dev server) send nothing unless you
set a DSN yourself. Set `SENTRY_DISABLED=true` to
opt out, or set `SENTRY_DSN` and `FRONTEND_SENTRY_DSN` to use your own. Both
take effect at runtime for the published images as well, because browser
reports travel through the Next server. See the
[observability guide](docs/observability.md) for the exact policy, per-runtime
variables, and the [Sentry data audit](docs/sentry-data-audit.md).

## Repository

| Path | Purpose |
| --- | --- |
| `frontend/` | Next.js web application |
| `backend/` | Express API, document processing, and database access |
| `word-addin/` | Microsoft Word task-pane add-in |
| `backend/schema.sql` | Complete schema for fresh databases |
| `backend/migrations/` | Dated migrations for existing deployments |
| `docker-compose.yml` | Local application and infrastructure stack |
| `docs/` | Development, deployment, testing, and feature guides |

## Documentation

- [Documentation index](docs/README.md)
- [Local development](docs/local-development.md)
- [Manual and production deployment](docs/deployment.md)
- [Troubleshooting](docs/troubleshooting.md)
- [MCP connectors](docs/connectors.md)
- [Google Workspace: Drive, Gmail, and Calendar](docs/google-workspace.md)
- [CourtListener integration](docs/courtlistener.md)
- [Microsoft Word add-in](word-addin/README.md)
- [Tamper-evident exports](docs/tamper-evident-exports.md)
- [Safe local testing](docs/safe-local-testing.md)
- [Preset contract templates and publisher credits](docs/preset-templates.md)
- [End-to-end testing and CI](docs/e2e-ci.md)
- [Contributing](CONTRIBUTING.md)
- [Open-source credits](CREDITS.md)
- [Security policy](SECURITY.md)

## System workflows

Mike's system assistant and tabular-review workflows are maintained in the
[`Open-Legal-Products/mike-workflows`](https://github.com/Open-Legal-Products/mike-workflows)
repository. See [Contributing](CONTRIBUTING.md#mike-workflows) for how they are
packaged and synchronized with this application.

## License

Mike is available under the [GNU Affero General Public License v3.0](LICENSE).
Bundled [public contract templates](docs/preset-templates.md#maintaining-the-catalog)
retain their publishers' separate licenses and notices.
