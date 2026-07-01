import { createClient } from '@supabase/supabase-js';

const SUPA_URL  = import.meta.env.VITE_SUPABASE_URL  as string;
const SUPA_KEY  = import.meta.env.VITE_SUPABASE_ANON_KEY as string;

export const supabase = SUPA_URL && SUPA_KEY
  ? createClient(SUPA_URL, SUPA_KEY)
  : null;

// Keys condivisi su cloud (tutto il resto rimane solo su IDB locale)
const CLOUD_KEYS = new Set([
  'ifb_logistics',
  'ifb_meatprices',
  'ifb_bevinfo',
  ...(["HK","CAN","MAC"] as const).flatMap(b => [
    `ifb_products_${b}`,
    `ifb_xrefs_${b}`,
    `ifb_airlist_${b}`,
    `ifb_sales_invoice_${b}`,
    `ifb_scattuali_${b}`,
    `ifb_prices_${b}`,
    `ifb_priceexceptions_${b}`,
    `ifb_fx`,
  ]),
]);

// ── Wrapper IDB ────────────────────────────────────────────────────────────────
let dbP: Promise<IDBDatabase> | null = null;
const openIDB = () => {
  if (!dbP) dbP = new Promise((res, rej) => {
    const r = indexedDB.open('ifb_store', 1);
    r.onupgradeneeded = e => (e.target as IDBOpenDBRequest).result.createObjectStore('store');
    r.onsuccess = e => res((e.target as IDBOpenDBRequest).result);
    r.onerror   = () => { dbP = null; rej(); };
  });
  return dbP;
};

export const IDB = {
  set: async (key: string, val: any) => {
    try {
      const db = await openIDB();
      await new Promise<void>((res, rej) => {
        const tx = db.transaction('store', 'readwrite');
        tx.objectStore('store').put(val, key);
        tx.oncomplete = () => res();
        tx.onerror = rej;
      });
      return true;
    } catch { return false; }
  },
  get: async (key: string, def: any = null) => {
    try {
      const db = await openIDB();
      return await new Promise(res => {
        const tx = db.transaction('store', 'readonly');
        const r  = tx.objectStore('store').get(key);
        r.onsuccess = () => res(r.result ?? def);
        r.onerror   = () => res(def);
      });
    } catch { return def; }
  },
  del: async (key: string) => {
    try {
      const db = await openIDB();
      await new Promise<void>(res => {
        const tx = db.transaction('store', 'readwrite');
        tx.objectStore('store').delete(key);
        tx.oncomplete = () => res();
        tx.onerror    = () => res();
      });
    } catch {}
  },
};

// ── CLOUD: IDB + Supabase sync ────────────────────────────────────────────────
export const CLOUD = {
  /** Legge da Supabase se disponibile e il key è condiviso; altrimenti da IDB */
  get: async (key: string, def: any = null) => {
    if (supabase && CLOUD_KEYS.has(key)) {
      try {
        const { data } = await supabase.from('app_data').select('data').eq('key', key).maybeSingle();
        if (data?.data !== undefined) {
          await IDB.set(key, data.data); // aggiorna cache locale
          return data.data;
        }
      } catch { /* offline → IDB fallback */ }
    }
    return IDB.get(key, def);
  },

  /** Scrive su IDB (immediato) + Supabase (async, solo admin) */
  set: async (key: string, val: any) => {
    await IDB.set(key, val);
    if (supabase && CLOUD_KEYS.has(key)) {
      try {
        await supabase.from('app_data')
          .upsert({ key, data: val, updated_at: new Date().toISOString() }, { onConflict: 'key' });
      } catch { /* ignore cloud errors */ }
    }
    return true;
  },

  del: IDB.del,
};

// ── Auth helpers ───────────────────────────────────────────────────────────────
export async function getSession() {
  if (!supabase) return null;
  const { data } = await supabase.auth.getSession();
  return data.session;
}

export async function getUserRole(email: string): Promise<'admin' | 'viewer' | null> {
  if (!supabase) return 'admin'; // no Supabase → local mode, full access
  try {
    const { data } = await supabase.from('user_roles').select('role').eq('email', email).maybeSingle();
    return (data?.role as 'admin' | 'viewer') ?? null;
  } catch { return null; }
}

export async function listUsers(): Promise<{email: string, role: string, created_at: string}[]> {
  if (!supabase) return [];
  const { data } = await supabase.from('user_roles').select('*').order('created_at');
  return data ?? [];
}

export async function inviteUser(email: string, role: 'viewer' | 'admin' = 'viewer') {
  if (!supabase) return;
  // Aggiunge il ruolo (la persona riceverà il magic link quando fa login)
  await supabase.from('user_roles').upsert({ email, role }, { onConflict: 'email' });
}

export async function removeUser(email: string) {
  if (!supabase) return;
  await supabase.from('user_roles').delete().eq('email', email);
}

export async function signInWithOtp(email: string) {
  if (!supabase) throw new Error('Supabase non configurato');
  const { error } = await supabase.auth.signInWithOtp({
    email,
    options: { emailRedirectTo: window.location.href },
  });
  if (error) throw error;
}

export async function signOut() {
  if (!supabase) return;
  await supabase.auth.signOut();
}
