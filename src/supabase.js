import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL || "";
const key = import.meta.env.VITE_SUPABASE_KEY || "";

export const supabase = url && key ? createClient(url, key) : null;
export const isConnected = () => !!supabase;

/* Test that the connection actually works by pinging app_settings */
export async function testConnection() {
  if (!supabase) return { ok: false, error: "No Supabase credentials set.\nAdd VITE_SUPABASE_URL and VITE_SUPABASE_KEY to Vercel environment variables, then redeploy." };
  const { error } = await supabase.from("app_settings").select("id").limit(1);
  if (error) return { ok: false, error: `Supabase error: ${error.message}\n\nCheck that:\n• Your Project URL and anon key are correct\n• You ran the SQL to create the tables\n• Your Supabase project is active` };
  return { ok: true };
}

export const db = {
  async load(table) {
    if (!supabase) return [];
    const { data, error } = await supabase.from(table).select("id, data").order("id", { ascending: true });
    if (error) throw new Error(`Load ${table}: ${error.message}`);
    return (data || []).map(r => r.data);
  },

  /* Batch upsert — throws on error */
  async upsertMany(table, items) {
    if (!supabase || !items.length) return;
    const { error } = await supabase.from(table).upsert(
      items.map(item => ({ id: String(item.id), data: item }))
    );
    if (error) throw new Error(`Upsert ${table}: ${error.message}`);
  },

  async deleteMany(table, ids) {
    if (!supabase || !ids.length) return;
    const { error } = await supabase.from(table).delete().in("id", ids.map(String));
    if (error) throw new Error(`Delete ${table}: ${error.message}`);
  },

  async upsertOne(table, item) {
    if (!supabase) return;
    const { error } = await supabase.from(table).upsert({ id: String(item.id), data: item });
    if (error) throw new Error(`UpsertOne ${table}: ${error.message}`);
  },

  async saveSetting(key, value) {
    if (!supabase) { localStorage.setItem("natrio_settings", JSON.stringify(value)); return; }
    const { error } = await supabase.from("app_settings").upsert({ id: key, data: value });
    if (error) throw new Error(`SaveSetting ${key}: ${error.message}`);
  },

  async loadSetting(key) {
    if (!supabase) return null;
    const { data, error } = await supabase.from("app_settings").select("data").eq("id", key).maybeSingle();
    if (error) return null;
    return data?.data || null;
  },
};
