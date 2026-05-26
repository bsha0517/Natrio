import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL || "";
const key = import.meta.env.VITE_SUPABASE_KEY || "";

export const supabase = url && key ? createClient(url, key) : null;
export const isConnected = () => !!supabase;

export const db = {
  /* Load all rows from a table, returning the data array */
  async load(table) {
    if (!supabase) return [];
    const { data, error } = await supabase
      .from(table)
      .select("id, data")
      .order("id", { ascending: true });
    if (error) { console.error(`[DB] Load ${table}:`, error.message); return []; }
    return (data || []).map(r => r.data);
  },

  /* Batch upsert (insert or update) many records */
  async upsertMany(table, items) {
    if (!supabase || !items.length) return;
    const { error } = await supabase
      .from(table)
      .upsert(items.map(item => ({ id: item.id, data: item })));
    if (error) console.error(`[DB] UpsertMany ${table}:`, error.message);
  },

  /* Delete records whose IDs are in the list */
  async deleteMany(table, ids) {
    if (!supabase || !ids.length) return;
    const { error } = await supabase.from(table).delete().in("id", ids);
    if (error) console.error(`[DB] DeleteMany ${table}:`, error.message);
  },

  /* Upsert a single record */
  async upsertOne(table, item) {
    if (!supabase) return;
    const { error } = await supabase
      .from(table)
      .upsert({ id: item.id, data: item });
    if (error) console.error(`[DB] UpsertOne ${table}:`, error.message);
  },

  /* Save a named setting (arbitrary JSON) */
  async saveSetting(key, value) {
    if (!supabase) return;
    const { error } = await supabase
      .from("app_settings")
      .upsert({ id: key, data: value });
    if (error) console.error(`[DB] SaveSetting ${key}:`, error.message);
  },

  /* Load a named setting */
  async loadSetting(key) {
    if (!supabase) return null;
    const { data, error } = await supabase
      .from("app_settings")
      .select("data")
      .eq("id", key)
      .maybeSingle();
    if (error) return null;
    return data?.data || null;
  },
};
