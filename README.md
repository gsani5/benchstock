# Benchstock — Shared Lab Inventory

A real, deployable lab inventory website. Accounts and the shared database run on
**Supabase** (free tier); the site is hosted on **Netlify** (free tier). Everyone in
the lab signs in and works off one live inventory.

Features: inventory items with lot tracking, low-stock and expiration alerts,
printable Code 39 barcode labels, per-use stock deduction with a usage log, member
list, and CSV export. Changes sync live across everyone signed in.

---

## What you'll end up with

A URL like `https://your-lab-inventory.netlify.app` (or your own domain). Lab members
go there, register with their email + a password, and share the same inventory.

Total setup time: about 15 minutes. No coding required — just copy/paste.

---

## Step 1 — Create the Supabase project (the backend)

1. Go to **https://supabase.com**, sign up (your Outlook email is fine), create a new project.
2. Pick a name and a strong database password (save it somewhere). Choose the region
   closest to the lab. Wait ~2 min for it to provision.
3. In the left sidebar open **SQL Editor → New query**. Open the file
   `supabase/schema.sql` from this folder, paste the whole thing in, and click **Run**.
   You should see "Success". This creates the tables, access rules, and live sync.
4. (Optional) To preload the lab's reagents, run `supabase/seed.sql` the same way.

### Turn off email confirmation (recommended for an internal tool)
By default Supabase emails every new user a confirmation link before they can sign in.
For a lab tool that's usually friction you don't want:
- Go to **Authentication → Sign In / Providers → Email** and turn **Confirm email** off.
- Now anyone who registers can sign in immediately. (Leave it on if you'd rather verify
  each address — they'll just get a one-time confirmation email.)

### Get your two keys
- Go to **Project Settings → API**.
- Copy the **Project URL** and the **anon public** key. You'll paste these into Netlify
  in Step 3. (The anon key is safe to expose in a website — access is controlled by the
  rules in `schema.sql`.)

---

## Step 2 — Put the code on GitHub

Easiest path for Netlify is connecting a GitHub repo.
1. Create a free account at **https://github.com** if you don't have one.
2. Make a new repository (e.g. `benchstock`), then upload the **contents of this folder**
   to it (drag-and-drop works on github.com → "uploading an existing file").
   Do **not** upload `node_modules` or `dist` if present.

> Prefer no GitHub? See "Drag-and-drop deploy" at the bottom.

---

## Step 3 — Deploy on Netlify

1. Go to **https://netlify.com**, sign up, and choose **Add new site → Import an existing project**.
2. Connect GitHub and pick your `benchstock` repo.
3. Netlify reads `netlify.toml` automatically, so the build command (`npm run build`) and
   publish folder (`dist`) are already set. Don't change them.
4. Before the first deploy, open **Site configuration → Environment variables** and add:
   - `VITE_SUPABASE_URL` = your Project URL from Step 1
   - `VITE_SUPABASE_ANON_KEY` = your anon public key from Step 1
5. Click **Deploy**. In ~1 minute you get your live URL.

### Point Supabase at your URL (so password resets work)
In Supabase → **Authentication → URL Configuration**, set **Site URL** to your Netlify
URL (e.g. `https://your-lab-inventory.netlify.app`). This makes the "forgot password"
email link return people to your site.

---

## Step 4 — Share it

Send the Netlify URL to the lab. Each person clicks **Register**, enters their email,
name, and a password, and they're in. The first inventory they see is whatever you
seeded (or empty, ready to fill).

### Make someone an admin (optional)
Roles are cosmetic by default (shown on the Members tab). To mark someone admin, in
Supabase → **Table Editor → profiles**, edit their row and set `role` to `admin`.

---

## Run it on your computer first (optional)

To test locally before deploying:
1. Install Node.js (https://nodejs.org, LTS version).
2. In this folder, copy `.env.example` to `.env` and fill in your two Supabase keys.
3. Run:
   ```
   npm install
   npm run dev
   ```
   Open the printed `http://localhost:5173`.

---

## Custom domain (optional)

If the lab gets a domain (or has a subdomain available), Netlify → **Domain management**
walks you through pointing it at the site with free HTTPS.

---

## Drag-and-drop deploy (no GitHub)

1. Install Node.js, then in this folder run `npm install` and `npm run build`.
   This creates a `dist/` folder.
2. Go to **https://app.netlify.com/drop** and drag the `dist` folder in.
3. Add the two environment variables (Step 3.4) under the site's settings, then
   trigger a redeploy. Note: with this method you re-run `npm run build` and re-drag
   `dist` whenever the code changes.

---

## A note on security

This uses Supabase's real authentication (hashed passwords, secure sessions, password
reset) and database access rules, which is appropriate for an internal lab tool. The
access rule is "any signed-in member can read and edit the shared inventory" — fine for
a trusting lab. If you ever need stricter controls (e.g. only admins can delete, or
read-only members), those rules live in `supabase/schema.sql` and can be tightened.
