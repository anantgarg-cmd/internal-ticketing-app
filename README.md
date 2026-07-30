# Internal Ticketing App

This Google Apps Script web app is deployed for the Shadowfax Workspace domain. It runs as the deployment owner; regular users must never be asked to grant Spreadsheet, Drive, or external-request permissions.

## Deployment-owner OAuth authorization

GitHub Actions updates the existing DEV or Production deployment, but it cannot approve Google account permissions. After deploying either environment, its deployment owner must complete authorization once:

1. Open the corresponding Apps Script project while signed in as the account that owns or deployed that web app.
2. Select `authorizeApplication` from the function dropdown.
3. Click **Run**.
4. Click **Review permissions**.
5. Select the Shadowfax deployment-owner account.
6. Approve all required permissions.
7. Run `authorizeApplication` again.
8. Confirm it returns `authorized: true`.
9. Reopen that environment's `/exec` URL.

Repeat these steps separately for **DEV** (the Apps Script project/deployment configured for the `develop` branch) and **Production** (the project/deployment configured for `main`). Do not make either web app public; both deployments must remain domain-restricted and execute as the deploying user.

For a safe status check, run `getAuthorizationDiagnostic` manually in the editor. `REQUIRED` means the deployment owner still needs to complete the steps above; the diagnostic never returns the authorization URL.

### If Google's OAuth page is blank or broken

- Allow third-party cookies for `script.google.com`, `accounts.google.com`, and `googleusercontent.com`.
- Allow pop-ups and redirects for `script.google.com`.
- Retry in a clean Chrome profile signed in only with the intended Shadowfax deployment-owner account.

Do not change web-app access to public or ask regular users to authorize application services.
