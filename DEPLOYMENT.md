# Vercel deployment

The app runs as a Next.js application on Vercel; MongoDB access is performed by Node.js route handlers. It is configured for Node 22, uncached API data, restricted sign-in, and in-memory browser downloads.

## Before deploying

1. Confirm the faculty advisor approves hosting the tracker and its data on personal Vercel and MongoDB Atlas accounts. Decide who will own and transfer those accounts after the current committee.
2. Push the project to a private GitHub repository. `.env.local` and other `.env*` files are ignored; never commit credentials.
3. Create a dedicated Atlas database user with read/write access only to `default_tracker`. Use a new, long random password. The URI is only used server-side.
4. Import the repository into the Vercel Hobby project. Add these variables to the **Production** environment:
   - `MONGODB_URI`
   - `MONGODB_DATABASE` (`default_tracker`)
   - `AUTH_GOOGLE_ID`
   - `AUTH_GOOGLE_SECRET`
   - `AUTH_SECRET` (generate a new random value)
   - `ALLOWED_EMAIL_DOMAIN` (`lbsim.ac.in`)
   - `INITIAL_ALLOWED_EMAILS` only if initial member seeding is needed
5. Deploy once to obtain the production URL. Configure Google OAuth with `https://<production-host>/api/auth/callback/google` as the redirect URI and the production origin as an authorized JavaScript origin. Set `AUTH_URL` in Vercel to the production URL and redeploy. Set `AUTH_TRUST_HOST=true` only if Auth.js reports `UntrustedHost`.
6. If available, select Vercel's Mumbai (`bom1`) function region near the Atlas cluster. Atlas may require an IP access-list entry for Vercel's changing egress addresses; use the narrowest option available and rely on the dedicated least-privilege database user. Do not reuse an Atlas administrator account.

Vercel Hobby is intended for non-commercial use and enforces function, payload, and bandwidth limits. This app has no scheduled-job dependency. Spreadsheet imports are limited to 4 MB per file and request, and are validated before processing. Excel exports are generated in the browser in memory; no serverless files are written. Review the current Vercel plan limits before launch.

## Verify after deployment

- Sign in with an active `@lbsim.ac.in` member on a phone and laptop. Confirm a personal Gmail account and an institutional address absent from the active member directory are refused.
- Check that student lookup, default logging, Excel roster/member imports, workbook export, DC Planner, After-DC Discussion, and Summary work.
- Confirm refreshing reflects changes and opening a protected `/api/...` data route without a session returns HTTP 401.
- Verify `MONGODB_URI` is set only in server-side Vercel environment settings, not in repository files or browser requests.
- Export the tracker workbook weekly during the season and keep an advisor-approved backup; the free Atlas tier does not include continuous backups.

Deploy changes before, not during, a DC meeting. Preview deployments have separate URLs and Google OAuth only works for configured redirect URLs; validate sign-in on production or locally.
