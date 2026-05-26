# Natrio Organics — Distribution Suite

Full-featured wholesale management app with Supabase database backend.

---

## 🗄 Step 1 — Create your free Supabase database

1. Go to **supabase.com** → Sign Up
2. Click **New Project** → name it `natrio` → set a password → Create
3. Wait ~2 minutes for setup to complete
4. Go to **SQL Editor** → paste this and click **Run:**

```sql
create table if not exists inventory       (id bigint primary key, data jsonb not null);
create table if not exists customers       (id bigint primary key, data jsonb not null);
create table if not exists sales           (id bigint primary key, data jsonb not null);
create table if not exists expenses        (id bigint primary key, data jsonb not null);
create table if not exists deliveries      (id bigint primary key, data jsonb not null);
create table if not exists purchase_orders (id bigint primary key, data jsonb not null);
create table if not exists audit_log       (id bigint primary key, data jsonb not null);
create table if not exists app_settings    (id text    primary key, data jsonb not null);
create table if not exists quotes          (id text    primary key, data jsonb not null);
```

5. Go to **Settings → API** and copy your **Project URL** and **anon/public** key.

---

## 🚀 Step 2 — Deploy to Vercel

1. Upload this folder to GitHub (drag & drop into a new repo)
2. Go to **vercel.com** → Add New Project → select your repo → Deploy
3. In Vercel → your project → **Settings → Environment Variables**, add:

| Name | Value |
|------|-------|
| `VITE_SUPABASE_URL` | Your Supabase Project URL |
| `VITE_SUPABASE_KEY` | Your Supabase anon/public key |

4. Go to **Deployments → Redeploy** — done! 🎉

---

## 🤖 Step 3 — Enable AI Assistant (optional)

In Vercel Environment Variables also add:
- `VITE_ANTHROPIC_API_KEY` = your key from console.anthropic.com

Or enter it directly in the AI Assistant tab inside the app.

---

## 💻 Run Locally

```bash
npm install
npm run dev
```

Create a `.env` file:
```
VITE_SUPABASE_URL=https://your-project.supabase.co
VITE_SUPABASE_KEY=your-anon-key
```

---

## Login credentials (change these in the Users tab after first login)
- admin / admin123
- sales / sales123  
- viewer / viewer123
