# Publishing Nisar Office Studio free on the Microsoft Store

This guide puts the app in the Microsoft Store as a **free, open-source** app at **no cost to you**. Every step below is free.

| What | Cost | Why you need it |
|---|---|---|
| GitHub account and public repository | Free | Shares the source code (open source) and hosts the privacy policy page |
| GitHub Pages | Free | A public web address for the privacy policy |
| Microsoft Store individual developer account | Free (since September 2025) | Lets you publish in the Store |
| Store certification and signing | Free | Microsoft tests the app and signs it with Microsoft's own certificate |

**Why the Store solves the "certify" problem:** Windows blocks the plain installer (Smart App Control: "An Application Control policy has blocked this file") because it is not code-signed, and a signing certificate costs about $220–$250 a year. In the Store you never buy a certificate: after certification Microsoft signs the package with its own certificate, and Windows (including Smart App Control and SmartScreen) trusts apps installed from the Store ([Microsoft: code signing options](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/code-signing-options); [electron-builder AppX docs](https://www.electron.build/docs/appx): "The Windows Store signs the package with a Microsoft certificate during the submission process").

What only you can do: create the accounts (they need your own identity), and click Submit. Everything else is prepared in this folder.

---

## Step 1. Put the source code on GitHub (free, about 10 minutes)

1. Create a free account at **https://github.com/signup**. Remember the user name you choose; below it is called `nisarkhan4u-hash`.
2. Click **+** (top right) → **New repository**. Name: `nisar-office-studio`. Choose **Public**. Do **not** add a README, licence or .gitignore (the project already has them). Click **Create repository**.
3. Put your GitHub user name into the docs (this replaces every `nisarkhan4u-hash`):

   ```
   cd "C:\Nisar Word"
   npm run github-name -- nisarkhan4u-hash
   ```

4. Upload the code. The folder is already a Git repository with everything staged. In GitHub → **Settings → Emails**, copy your private `...@users.noreply.github.com` address (so your real e-mail is not published), then:

   ```
   cd "C:\Nisar Word"
   git config user.name "Nisar"
   git config user.email "<your ...@users.noreply.github.com address>"
   git add -A
   git commit -m "Nisar Office Studio 3.0.1"
   git branch -M main
   git remote add origin https://github.com/nisarkhan4u-hash/nisar-office-studio.git
   git push -u origin main
   ```

   The first push opens a browser window to sign in to GitHub. `node_modules`, `dist` and the phone version (`mobile/`) are excluded by `.gitignore`.

## Step 2. Publish the privacy policy (free, 2 minutes)

The Store asks for a privacy policy web address.

1. On GitHub open your repository → **Settings → Pages**.
2. Under **Build and deployment**: Source **Deploy from a branch**, Branch **main**, folder **/docs** → **Save**.
3. After a minute or two these pages are live:
   - `https://nisarkhan4u-hash.github.io/nisar-office-studio/` (a small home page with screenshots)
   - `https://nisarkhan4u-hash.github.io/nisar-office-studio/privacy-policy.html` (use this in the Store)

## Step 3. Register as a Microsoft Store developer (free, about 15 minutes)

Individual developer registration costs nothing (the old $19 fee is waived) and is open in nearly 200 markets ([Microsoft Learn](https://learn.microsoft.com/en-us/windows/apps/publish/whats-new-individual-developer); [Windows Developer Blog, 10 Sept 2025](https://blogs.windows.com/windowsdeveloper/2025/09/10/free-developer-registration-for-individual-developers-on-microsoft-store/)).

1. Go to **https://storedeveloper.microsoft.com**. Start here: going through Partner Center first can show the old paid sign-up.
2. Click **Get started for free** → **Individual developer**, and sign in with your Microsoft account (for example your Hotmail/Outlook account).
3. Verify your identity with a **government-issued ID and a selfie** (use your phone, good light and the original document).
4. Check the profile details and finish. You land in **Partner Center**.

## Step 4. Reserve the name and copy the identity (5 minutes)

1. Partner Center → **Apps and games** → **New product** → **MSIX or PWA app**.
2. Type **Nisar Office Studio** → **Check availability** → **Reserve product name**.
3. Open the product → **Product management** → **Product identity**. Copy the three values and run:

   ```
   cd "C:\Nisar Word"
   npm run store-identity -- "<Package/Identity/Name>" "<Package/Identity/Publisher>" "<Package/Properties/PublisherDisplayName>"
   ```

   Example: `npm run store-identity -- "12345Nisar.NisarOfficeStudio" "CN=1A2B3C4D-5E6F-7A8B-9C0D-1E2F3A4B5C6D" "Nisar"`. The values must match Partner Center exactly, or the upload is rejected.

## Step 5. Build the Store package (5 minutes)

```
cd "C:\Nisar Word"
npm run dist:store
```

The result is `C:\Nisar Word\dist\NisarOfficeStudio-3.0.1.appx` (about 225 MB). It is not signed, which is correct for the Store: Microsoft signs it. The build stops with a clear message if Step 4 has not been done yet.

It is built for 64-bit Windows 10 version 1809 (build 17763) and later, including Windows 11. The logos come from `build/appx/` (re-create them with `node scripts/make-store-assets.js`).

## Step 6. Submit (20 minutes, then wait for certification)

Partner Center → your app → **Start your submission**. Fill in each section; the texts are ready in [`store-listing.md`](store-listing.md).

1. **Pricing and availability:** Markets: all. Pricing: **Free**. Visibility: Public.
2. **Properties:**
   - Category: **Productivity**.
   - Privacy policy URL: the GitHub Pages address from Step 2.
   - Website and support contact: your GitHub repository and its Issues page.
   - Product declarations: the app does not collect personal information; it has no in-app purchases.
3. **Age ratings:** fill in the questionnaire as described in `store-listing.md`. The result should be the lowest rating.
4. **Packages:** drag in `dist\NisarOfficeStudio-3.0.1.appx` and wait for "Validated".
5. **Store listings → English (United States):** description, short description, search terms, release notes and screenshots from `docs/screenshots/`.
6. **Submission options → Notes for certification:** paste this:

   > Nisar Office Studio is a free, open-source (MIT) offline word processor and PDF editor built with Electron. It needs no account or internet connection. The package declares runFullTrust because it is a desktop (Win32) app packaged as MSIX; full trust is needed to open and save the documents the user picks and to run the bundled qpdf tool for PDF passwords. Source code: https://github.com/nisarkhan4u-hash/nisar-office-studio

   Every Electron/desktop app declares the restricted `runFullTrust` capability. Partner Center asks you to explain it, and the note above does that.
7. Click **Submit to the Store**. Certification usually takes **up to 3 business days**. You get an e-mail when the app is live, or a report saying what to fix.

After it is live:
- People install it from the Store, and it is signed and trusted by Smart App Control.
- `winget install "Nisar Office Studio"` also works, because winget searches the Microsoft Store too. You don't need to submit anything extra.
- Put the Store link in `README.md`.

## Step 7. Updates

1. Increase `"version"` in `package.json` (for example `3.0.1`).
2. Run `npm run dist:store`.
3. Partner Center → your app → **Update** (new submission) → upload the new `.appx` → Submit.
4. Store users get the update automatically.

Optionally, push a tag (`git tag v3.0.1` then `git push --tags`) and the GitHub workflow `.github/workflows/windows.yml` builds the installer and the Store package on GitHub's free Windows machines. It also attaches the installer to a GitHub Release.

---

## If certification fails: common reasons

| Report says | Fix |
|---|---|
| Package identity does not match | Run Step 4 again with the exact values, then rebuild. |
| Privacy policy missing or unreachable | Check that the GitHub Pages link opens in a private browser window. |
| Restricted capability runFullTrust | Make sure the certification note from Step 6 is there; reply in Partner Center if they ask again. |
| Listing mentions other companies' products | Remove "Microsoft Word", "Microsoft Office", "Adobe" or "Acrobat" from the listing text and screenshots. |
| App crashes or does not start on their test PC | Install the same `.appx` build on another PC to reproduce, fix, rebuild with a higher version and resubmit. |

---

## Optional, also free: a signed installer outside the Store

The Store is enough for most people. If you also want the `.exe` on GitHub Releases to be signed (so it is not blocked when downloaded directly), the **SignPath Foundation** gives free code-signing certificates to open-source projects ([signpath.org](https://signpath.org/), [conditions](https://signpath.org/terms)). The main conditions:
- An OSI-approved licence (MIT qualifies) with no proprietary components.
- An active project that is already released.
- The functionality described on the download page.
- Builds made automatically from the repository (the GitHub workflow above does this).
- A code-signing policy page.

The certificate names SignPath Foundation as the publisher, and they can revoke it if the rules are broken. Apply on their website after the project has been public for a while. Even a signed installer can be warned about until it builds reputation ([Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/5791210/application-signed-and-blocked-by-smart-app-contro)). The Store version does not have this problem.

Paid options (not needed): an individual code-signing certificate costs about $220–$250 a year ([SSL.com](https://www.ssl.com/guide/esigner-pricing-for-code-signing/)). Azure Artifact Signing costs $9.99 a month and is only for individuals in the USA and Canada ([quickstart](https://learn.microsoft.com/en-us/azure/artifact-signing/quickstart)).

## Files used for publishing

| File | Purpose |
|---|---|
| `package.json` → `build.appx` | Store identity, display name, supported Windows versions |
| `build/appx/*.png` | Store and Start-menu logos (`scripts/make-store-assets.js`) |
| `scripts/set-store-identity.js` | `npm run store-identity`: writes the Partner Center identity |
| `scripts/check-store-identity.js` | Stops `dist:store` while the identity is still a placeholder |
| `scripts/set-github-name.js` | `npm run github-name`: fills in your GitHub address |
| `scripts/make-notices.js` | `npm run notices`: regenerates `THIRD-PARTY-NOTICES.md` |
| `docs/privacy-policy.html`, `docs/index.html` | Pages for GitHub Pages |
| `docs/store-listing.md`, `docs/screenshots/` | Store listing text and images |
| `.github/workflows/windows.yml` | Free automatic builds on GitHub |
| `LICENSE`, `THIRD-PARTY-NOTICES.md` | MIT licence and the licences of included libraries |
