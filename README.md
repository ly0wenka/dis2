# dis2

This repository contains the current dissertation manuscript as a `.docx`.

## Download as PDF from GitHub Releases

When you **publish a GitHub Release**, GitHub Actions will:
- convert the latest `*.docx` to PDF (LibreOffice, headless)
- attach the generated PDF to the Release assets

### How to use
1. GitHub → **Releases** → **Draft a new release**
2. Create/select a tag (for example `97`) on `main`
3. Publish the release
4. Wait for the workflow **Build PDF for Release**
5. Download the `.pdf` from the release assets

### Re-run for an existing release (if needed)
GitHub → **Actions** → **Build PDF for Release** → **Run workflow** and set `tag` to the release tag.
