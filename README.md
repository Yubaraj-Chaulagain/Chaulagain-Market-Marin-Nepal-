# Home & Market Manager Nepal — Full Starter Package

This ZIP contains a browser-based offline-first starter app and a Google Apps Script backend that automatically creates its Google Spreadsheet and tabs.

## Included modules
- Customers and user records
- Rooms and room-rent bills
- Market product catalogue with image URL, price, stock and cart
- Orders and payment records (cash, bank QR, eSewa/Khalti test labels, Fonepay QR labels)
- Savings deposits and withdrawals
- Loans with per-loan Fixed Interest or Reducing Balance choice
- Loan repayments with principal/interest split and transaction history
- Gmail notification queue (Apps Script sends email when a notification row is synced)
- Local browser storage, offline sync queue, JSON backup/import
- GitHub Actions workflow to package source ZIP

## 1. Create the Google Sheets database automatically
1. Open `https://script.google.com/` and create a new Apps Script project.
2. Replace the editor contents with `apps-script/Code.gs`.
3. Select and run `setupProject()` once. Approve Google permissions.
4. In Apps Script, choose **Deploy → New deployment → Web app**.
5. Select an execution/access setting appropriate to your use. Copy the deployed `/exec` URL.
6. Open the app, sign in, go to **Settings / Sync**, paste the URL, save, then Sync.

The setup function creates a spreadsheet called **Home & Market Manager Nepal Database** with these tabs: Customers, Rooms, Room Rent, Products, Orders, Payments, Savings, Withdrawals, Loans, Loan Repayments, Transactions, Users, Notifications, Settings.

## 2. Login
Initial local demo login:
- Username: `admin`
- Password: `ChangeMe123!`

Change this immediately in the Users data before real use. This starter stores demo users locally and does not provide production-grade identity verification.

## 3. Offline and sync behavior
- Data is saved in the current browser's localStorage.
- Mutations are queued locally while offline or when the backend is unavailable.
- When online, the app attempts to send queued upserts to Apps Script.
- Use **Sync now** to fetch server data into the current browser.
- Download JSON backups regularly.

**Important:** localStorage is device/browser-specific. This package is not yet a secure multi-customer cloud application. The Apps Script endpoint must be protected with proper identity and authorization checks before exposing private customer/financial records publicly. Do not rely on a secret embedded in HTML; browser code is visible to users.

## 4. Loan interest
The app records the interest method on each loan:
- Fixed Interest: monthly interest uses the original principal.
- Reducing Balance: monthly interest uses remaining principal.

The repayment helper allocates the current monthly interest first, then reduces principal. This is a simple ledger calculation, not a full amortization engine with late fees, day-count conventions, or a finalized installment schedule. Verify all financial calculations before production use.

## 5. Payments
Cash and manual bank QR can be recorded. eSewa/Khalti options are only test/pending labels in this starter. Fonepay is not live integrated. Real eSewa/Khalti/Fonepay processing requires merchant onboarding, provider-issued credentials, backend signature verification, payment status lookup, callback handling, refunds/reconciliation, and production deployment. Never place merchant secrets in `index.html` or send them in chat.

## 6. Gmail notifications
A notification row synced to the Apps Script `notifications` tab invokes `MailApp.sendEmail`. The Apps Script owner must authorize MailApp. Gmail quotas and account permissions apply. The current UI queues a message; add business-specific templates and delivery-status handling before production.

## 7. GitHub Actions
Push this folder to a GitHub repository. The workflow `.github/workflows/package.yml` packages the project and uploads a downloadable ZIP artifact on each push or manual run.

## Production checklist
- Replace local demo login with Google Identity / a proper authenticated backend.
- Enforce per-role and per-customer authorization server-side on every API call.
- Validate all input and amounts server-side; add audit logs and backups.
- Add actual gateway integration only after merchant approval and sandbox tests.
- Add server-side stock locking, duplicate-payment protection and reconciliation.
- Review privacy, consent, retention, tax/invoice and Nepal legal requirements.
- Test restore, concurrent users, offline conflict resolution and all financial calculations.

## Android APK build with GitHub Actions

This project includes a native Android WebView wrapper in the `android/` folder. It loads the same offline-first `index.html` from the APK assets.

1. Upload/extract this project into a GitHub repository (keep the `.github/workflows/` folder).
2. In GitHub, open **Settings → Actions → General** and ensure the workflow has permission to create and approve pull requests / write repository contents, or at minimum allow **Read and write permissions** for `GITHUB_TOKEN`.
3. Push the project to the `main` or `master` branch, or open **Actions → Build APK and Publish GitHub Release → Run workflow**.
4. GitHub Actions builds the APK and automatically creates a versioned GitHub Release (for example `v1.0.12`).
5. Open the repository's **Releases** page. Download `Home-Market-Manager-Nepal.apk` directly from the release assets. No Artifacts download is needed.

If you publish a GitHub Release manually, the workflow also builds the APK and attaches it to that release. This workflow produces a debug-signed testing APK, not a Play Store production-signed build. A Play Store release requires your own signing keystore and protected GitHub Actions secrets.

The Android wrapper provides WebView, JavaScript, local browser storage, internet access, external HTTPS link handling, and Android back navigation. This does not make the backend or payment gateways production-secure; configure and test the Apps Script deployment and access controls before using real customer financial data.
