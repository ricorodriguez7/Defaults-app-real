# Vercel deployment

The app is a Next.js project. Its MongoDB-backed API routes run as Node.js functions; no separate backend server or scheduled job is required. Vercel functions are configured for Mumbai (`bom1`), and the upload routes have a 30-second limit.

## Before deploying

1. Confirm the faculty advisor approves hosting the placement tool on personal Vercel and MongoDB Atlas accounts. Vercel Hobby is intended for non-commercial, personal use; review the current [Hobby plan limits](https://vercel.com/docs/plans/hobby).
2. Push this project to a private GitHub repository owned by an account the placement cell can transfer or hand over.
3. In Atlas, create a dedicated database user with read/write access only to `default_tracker`. Use a long, unique password. If using Atlas Network Access with Vercel's dynamic function IPs, `0.0.0.0/0` may be required; use only the restricted database user and strong credentials.
4. Import the private repository into Vercel. Set these **Production** environment variables in the Vercel project settings:
   - `MONGODB_URI`: the Atlas URI for the restricted database user
   - `MONGODB_DATABASE`: `default_tracker`
   - `AUTH_GOOGLE_ID` and `AUTH_GOOGLE_SECRET`: the Google OAuth client credentials
   - `AUTH_SECRET`: a newly generated random secret
   - `AUTH_URL`: the deployed production URL, for example `https://your-app.vercel.app`
   - `ALLOW_DATA_RESET`: keep unset or set to `false` for real use. Set to `true` only in a testing deployment that needs the data-management reset screen; redeploy after changing it.
5. In Google Cloud Console, add `https://your-app.vercel.app/api/auth/callback/google` as an authorized redirect URI and the deployed origin as an authorized JavaScript origin. Publish the OAuth consent screen if members outside the test-user list need access.
6. Deploy, note the production URL, set `AUTH_URL` to that exact URL, and redeploy if it was not known before the first deployment.
7. In the app, add each member with the exact Google email address they use. Only active members in this directory can sign in.

Do not commit `.env.local` or put credentials in the repository. `.env.local` is gitignored. The example environment file contains placeholders only. Vercel environment variable changes take effect after redeployment.

## Operational checks

- Sign in as an active member on both a phone and a laptop. Verify that an unlisted or inactive account is refused.
- Log a default and confirm it appears in the app and Atlas.
- Upload student and member spreadsheets smaller than 4 MB, and verify exports download successfully. Large exports must be filtered or split.
- Check the DC planner, discussion workflow, and summary counts.
- Confirm an API request without a session returns `401`, and refresh after a change to confirm fresh data.
- Confirm `MONGODB_URI` is not present in the repository, page source, or browser requests.
- Export the workbook regularly and keep the backup with the faculty advisor; the free Atlas tier may not provide continuous backups.
- Defaults can be imported from the Defaults view using the sample Excel workbook. Imports are limited to 3,000 rows and under 4 MB; check Import history for eligible undo actions.
- For testing only, set `ALLOW_DATA_RESET=true` in the deployment environment to reveal Settings → Data management. The reset requires exporting a full workbook backup and typing `RESET`. Remove the variable or set it to `false` before real use.

Keep the Vercel, GitHub, Atlas, and Google Cloud accounts and their credentials in an approved password manager. Transfer the accounts and project ownership before the current committee graduates. For a bad release, promote the last working Vercel deployment. Test changes the evening before a DC, not on the day itself.
