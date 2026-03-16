# Outlook Recipient Probe

This repo can now be deployed to GitHub Pages in addition to running locally as an ASP.NET static host.

## Local development

- Run the app locally and keep using the checked-in [`manifest.xml`](/f:/git/outlook-web-addin/OutlookRecipientProbe/manifest.xml), which points at `https://localhost:7123`.

## GitHub Pages deployment

- The GitHub Actions workflow at [`.github/workflows/deploy-pages.yml`](/f:/git/outlook-web-addin/OutlookRecipientProbe/.github/workflows/deploy-pages.yml) publishes the contents of `wwwroot` to GitHub Pages.
- During deployment, [`scripts/publish-pages.ps1`](/f:/git/outlook-web-addin/OutlookRecipientProbe/scripts/publish-pages.ps1) also generates a Pages-ready `manifest.xml` inside the published artifact.
- After Pages is enabled for the repository, the deployed manifest will be available at:
  `https://<owner>.github.io/<repo>/manifest.xml`

## Local Pages package test

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-pages.ps1 `
  -BaseUrl https://<owner>.github.io/<repo> `
  -OutputDir .\publish\pages
```
