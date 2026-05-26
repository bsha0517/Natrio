# Natrio Organics — Distribution Suite

---

## 🗄 Step 1 — Create Supabase database

1. Go to **supabase.com** → Sign Up free
2. Click **New Project** → name it `natrio` → set a password → Create
3. Wait ~2 minutes, then go to **SQL Editor** in the left menu
4. Paste the SQL below and click **Run:**

```sql
-- Drop existing tables if re-running
drop table if exists inventory, customers, sales, expenses, deliveries,
  purchase_orders, audit_log, app_settings, quotes, users;

-- Create all tables (use text IDs — required for JS UIDs)
create table inventory       (id text primary key, data jsonb not null default '{}');
create table customers       (id text primary key, data jsonb not null default '{}');
create table sales           (id text primary key, data jsonb not null default '{}');
create table expenses        (id text primary key, data jsonb not null default '{}');
create table deliveries      (id text primary key, data jsonb not null default '{}');
create table purchase_orders (id text primary key, data jsonb not null default '{}');
create table audit_log       (id text primary key, data jsonb not null default '{}');
create table app_settings    (id text primary key, data jsonb not null default '{}');
create table quotes          (id text primary key, data jsonb not null default '{}');
create table users           (id text primary key, data jsonb not null default '{}');

-- Allow public read/write (app handles its own auth)
alter table inventory       enable row level security;
alter table customers       enable row level security;
alter table sales           enable row level security;
alter table expenses        enable row level security;
alter table deliveries      enable row level security;
alter table purchase_orders enable row level security;
alter table audit_log       enable row level security;
alter table app_settings    enable row level security;
alter table quotes          enable row level security;
alter table users           enable row level security;

create policy "allow all" on inventory       for all using (true) with check (true);
create policy "allow all" on customers       for all using (true) with check (true);
create policy "allow all" on sales           for all using (true) with check (true);
create policy "allow all" on expenses        for all using (true) with check (true);
create policy "allow all" on deliveries      for all using (true) with check (true);
create policy "allow all" on purchase_orders for all using (true) with check (true);
create policy "allow all" on audit_log       for all using (true) with check (true);
create policy "allow all" on app_settings    for all using (true) with check (true);
create policy "allow all" on quotes          for all using (true) with check (true);
create policy "allow all" on users           for all using (true) with check (true);
```

5. Go to **Settings → API** → copy your **Project URL** and **anon/public** key

---

## 🚀 Step 2 — Deploy to Vercel

1. Upload this folder to GitHub (drag & drop into a new repo)
2. Go to **vercel.com** → Add New Project → select your repo → Deploy
3. In Vercel → your project → **Settings → Environment Variables**, add:

| Name | Value |
|------|-------|
| `VITE_SUPABASE_URL` | Your Supabase Project URL |
| `VITE_SUPABASE_KEY` | Your Supabase anon/public key |

4. **Deployments → Redeploy** — done! 🎉

---

## 🔑 First Login

When the database is empty, the app creates one default admin account:
- **Username:** `admin`
- **Password:** `admin123`

**Change this immediately** after first login via the Users tab.

---

## 🤖 AI Assistant (optional)

In Vercel Environment Variables also add:
- `VITE_ANTHROPIC_API_KEY` = your key from console.anthropic.com

Or enter it directly in the AI Assistant tab inside the app.

---

## 💻 Run Locally

Create a `.env` file in the project root:
```
VITE_SUPABASE_URL=https://your-project.supabase.co
VITE_SUPABASE_KEY=your-anon-key
```

Then:
```bash
npm install
npm run dev
```
