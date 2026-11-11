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


## Customer selection and record management update

- Customer dropdowns are populated from registered Customers for rooms, rent bills, orders, payments, savings, withdrawals, loans and repayments.
- Data tables now include Edit / Update and Delete actions.
- Product list also has an administration table with edit/delete controls.
- Changes are saved locally and queued to Google Apps Script using upsert/delete.
- Customer dropdowns will be empty until at least one customer is added.


## Customer selection, record management, and password reset update

- List views now include **Edit / Update** and **Delete** actions. Edits are queued for Google Sheets upsert; deletes are queued for the Apps Script delete endpoint.
- Room tenants, rent bills, orders, payments, savings, withdrawals, loans, and loan repayments use customer dropdowns where applicable. Create customers first.
- Login includes **Forgot password?**. The user enters their username and registered email, receives a 6-digit OTP by Gmail, then confirms the OTP and sets a new password.
- Password reset requires the updated `apps-script/Code.gs` to be redeployed as a Web App and MailApp permission to be authorized. After changing Code.gs, deploy a **new version** of the Web App and update the app's Web App URL if it changed.
- This is still a prototype: the current app stores passwords as plain text in the starter database and performs local demo login. Do not use it for real customer accounts or financial data until authentication is moved fully server-side, passwords are hashed, sessions/roles are enforced on every API call, and the backend is security-reviewed.


## Premium UI Theme
Applied a premium navy-blue and gold theme across the web app and Android WebView:
- Premium gradient header and dark navigation
- Soft glass-style cards and subtle shadows
- Gold primary buttons and polished form controls
- Refined tables, responsive mobile navigation, and background gradients
- Existing application logic and data fields are preserved


## Android launcher icon
The approved Home & Market Manager Nepal artwork (including the creator credit) is installed as the Android launcher icon.
- Legacy launcher icons: `android/app/src/main/res/mipmap-*`
- Adaptive launcher icon: `android/app/src/main/res/mipmap-anydpi-v26/`
- Manifest uses `@mipmap/ic_launcher` and `@mipmap/ic_launcher_round`.
- Original full-resolution artwork: `branding/Home-Market-Manager-Nepal-Icon-Original.png`


## Apps Script endpoint moved out of the app Settings screen
- The Settings screen no longer displays a Deployment URL input or Save URL button.
- The endpoint is configured in source code as `APPS_SCRIPT_WEB_APP_URL` in `index.html` (the same page is copied into the Android assets).
- Replace `PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_EXEC_URL_HERE` with the deployed Apps Script Web App `/exec` URL before building the APK.
- The full backend source remains in `apps-script/Code.gs`. Apps Script cannot execute merely by being placed inside an APK; Google requires it to be deployed as a Web App once. The spreadsheet/tabs can then be created by running `setupProject()` in Apps Script.
- Never put eSewa/Khalti/Fonepay merchant secrets in the HTML or APK.


## Automatic Google Sheets database setup
1. Open `apps-script/Code.gs` in Google Apps Script.
2. Deploy it as a Web App (execute as the deploying account; grant access as required).
3. Put the deployed `/exec` URL once in `APPS_SCRIPT_WEB_APP_URL` in `index.html`. The same file is copied into the Android WebView assets.
4. In the app, open Settings / Sync and tap **Sync now**.
5. The app sends the `setup` action. `setupProject()` automatically creates the spreadsheet (if not already created) and creates/checks all required tabs and header rows: customers, rooms, roomRent, products, orders, payments, savings, withdrawals, loans, loanRepayments, transactions, users, notifications, settings.
6. The spreadsheet URL is returned in the setup response. Later syncs reuse the same spreadsheet ID stored in Apps Script Script Properties; it does not create a new spreadsheet each time.

Important: Google Apps Script still must be deployed once and authorized by the owner. The app cannot deploy Apps Script or grant Google permissions by itself. Never put payment gateway secrets in the HTML.


## Admin self-service Apps Script creation and deployment (OAuth)
The Settings page includes an Admin-only **Create & Deploy Apps Script** action. It uses the official Apps Script REST API with a user-authorized OAuth access token, creates a script project, uploads `apps-script/Code.gs` plus a web-app manifest, creates a version, deploys it, saves the returned `/exec` URL in this app's local storage, then runs spreadsheet setup/sync.

**One-time Google Cloud prerequisites (cannot be skipped):**
1. Create/select a Google Cloud project and enable **Google Apps Script API**.
2. Configure the OAuth consent screen and add the Google account as a test user if the app is in Testing.
3. Create an OAuth Client ID of type **Web application**. Add the exact app origin to Authorized JavaScript origins. For local `file://` use, OAuth origins are not supported; run the web app from an allowed HTTPS origin. For Android WebView, Google OAuth may block embedded user agents; use a system-browser OAuth flow / native Google Sign-In for production.
4. Put the OAuth Web Client ID in the source constant `GOOGLE_OAUTH_CLIENT_ID` before building, or enter it in Admin Settings. Client ID is not a secret. Never put a Client Secret or service-account private key in the app.
5. `apps-script/Code.gs` must be served from the app's same-origin `/apps-script/Code.gs` path for this prototype auto-deployer. If hosted as a standalone Android asset, the fetch path may not be available; host the project files on HTTPS or bundle Code.gs into the deployment client.

Google account permission is requested at runtime. Apps Script deployment is owned by the Google account that authorizes it. The Web App manifest uses execute-as-deployer and access-anyone; review access/security before using real customer or financial data. This starter is not production-secure and does not replace backend authorization or financial compliance review.


Auto-deploy note: the Apps Script source is embedded in the app's JavaScript bundle so Android WebView does not need to fetch a separate Code.gs file at runtime.
