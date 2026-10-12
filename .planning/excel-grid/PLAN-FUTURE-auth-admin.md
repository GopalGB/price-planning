# Future plan: authenticator sign-in and admin portal (NOT approved for build)

Source: Sir's brain dump, 2026-10-12. Build only after Sir approves the current release.

## Sign-up and sign-in

- Sign-up asks for first name, last name, work email and an authenticator code.
- The app shows a QR code (standard TOTP, RFC 6238, 6 digits, 30 s) that Microsoft Authenticator or
  Google Authenticator scans; the first code entered proves the scan worked before the account exists.
- Phase 1 (with the first rollout): password plus authenticator code.
- Phase 2 (later): authenticator code only, no password. Needs one-time recovery codes at sign-up
  and an admin "reset authenticator" action, because a lost phone would otherwise lock the user out.
- Users reach the app by its link on the company network; nothing is stored on their own PCs.

## Where account data lives

- On the office server, in the app's existing database (not on users' machines).
- Authenticator secrets encrypted at rest (key held by the server account, for example Windows DPAPI),
  never shown again after setup; passwords stay hashed as today.
- Login attempts rate limited and logged (security rules: 5 per minute).

## Admin portal

- Admins see all users, can enable or disable an account, reset an authenticator, and grant or
  remove admin rights.
- No category admins: any member of a category can give other people access to that category
  (as today); the portal only adds the global admin layer.

## Open questions for Sir (when this is picked up)

1. Should sign-up be open to anyone on the network, or need an admin approval or an invite?
2. Restrict sign-up to company email domains?
3. Who is the first admin?
