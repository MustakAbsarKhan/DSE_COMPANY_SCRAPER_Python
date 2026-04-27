# Security Policy

## Supported Version

The maintained version is the latest code on the default branch.

## Reporting A Security Issue

Please report security concerns privately by email:

mustak.absar.khan@gmail.com

Do not open a public GitHub issue for sensitive security reports.

## Security Notes

- This project does not require API keys, passwords, tokens, or private credentials.
- Runtime logs and generated Excel exports should not be committed.
- If future features require secrets, store them in environment variables or GitHub Actions secrets.
- Do not hard-code credentials in source files, workflow files, notebooks, or documentation.
- Public source code can be copied technically; licensing and attribution protect permitted use, not physical access.

## Responsible Use

This scraper is intended to access public DSE pages respectfully. Keep request rates reasonable, honor website availability, and adjust throttling if the source website becomes slow or unstable.
