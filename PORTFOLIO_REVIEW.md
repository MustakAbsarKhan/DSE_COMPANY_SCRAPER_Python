# Portfolio Review Brief

## Project Summary

This public repository demonstrates the architecture of a production-minded Python data pipeline for collecting DSE company fundamentals, normalizing inconsistent public HTML, and exporting an analysis-ready Excel workbook.

The production source is intentionally excluded to protect implementation details and reduce unauthorized copying or commercial reuse. For evaluation, reviewers can inspect this brief, the main README, the execution-flow diagram, interface skeletons, and selected redacted walkthrough materials.

## What This Demonstrates

- Async Python orchestration with bounded concurrency
- Resilient HTTP fetching with retry and adaptive throttle behavior
- Parser design for inconsistent HTML tables and labels
- Financial-field normalization and ordered workbook output
- Multi-sheet Excel reporting for business users
- Data-quality checks, outlier handling, and explainable screening labels
- Scheduled and manual automation design, documented without running in this public repo
- Clear module separation across client, parser, pipeline, and export layers

## Review-Friendly Files

| File or folder | What to look for |
| --- | --- |
| `main.py` | Async orchestration and pipeline flow preview |
| `core/client.py` | HTTP client, retry, and throttling interface preview |
| `core/parser.py` | Defensive parsing and normalization boundary preview |
| `export/excel.py` | Workbook sheet and processed-analysis schema preview |
| `pipelines/` | Sector, company-list, and company-profile stage boundaries |
| `docs/automation-architecture.yml` | Non-running production automation preview |
| `Execution Flow.jpg` | High-level system flow |

## Privacy Boundary

The project may be discussed, demonstrated, and reviewed for portfolio, hiring, or client evaluation. The private source, generated workbooks, and implementation details should not be redistributed, republished, used in commercial products, or presented as another person's work without written permission from the author.

## Suggested Reviewer Path

1. Read the README for scope and architecture.
2. Review the execution-flow diagram.
3. Inspect the pipeline modules to understand data movement.
4. Inspect parser and export modules for engineering depth.
5. Ask for a live walkthrough or sanitized output sample if business context is needed.
