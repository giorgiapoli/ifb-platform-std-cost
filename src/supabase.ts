import { createClient } from '@supabase/supabase-js';

const SUPA_URL  = import.meta.env.VITE_SUPABASE_URL  as string;
const SUPA_KEY  = import.meta.env.VITE_SUPABASE_ANON_KEY as string;

export const supabase = SUPA_URL && SUPA_KEY
  ? createClient(SUPA_URL, SUPA_KEY)
  : null;

// Accesso consentito solo al personale Inalca Food & Beverage
const COMPANY_DOMAIN = 'inalcafb.com';
export function isCompanyEmail(email: string): boolean {
  return email.trim().toLowerCase().endsWith(`@${COMPANY_DOMAIN}`);
}

// Nessun dato condiviso su Supabase — tutto rimane in IDB locale
// Supabase è usato solo per autenticazione (user_roles = solo ruoli Admin)
const CLOUD_KEYS = new Set<string>();

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
  /** Legge da Supabase se disponibile; preferisce il dato più recente tra Supabase e IDB */
  get: async (key: string, def: any = null) => {
    if (supabase && CLOUD_KEYS.has(key)) {
      try {
        const { data } = await supabase.from('app_data').select('data, updated_at').eq('key', key).maybeSingle();
        if (data?.data !== undefined) {
          const supaTs = data.updated_at ? new Date(data.updated_at).getTime() : 0;
          const localTs: number = (await IDB.get(`__ts__${key}`, 0)) as number;
          if (localTs > supaTs) {
            // IDB è più recente (write Supabase precedente fallito) → resync
            const localVal = await IDB.get(key, def);
            supabase.from('app_data')
              .upsert({ key, data: localVal, updated_at: new Date(localTs).toISOString() }, { onConflict: 'key' })
              .then().catch(() => {});
            return localVal;
          }
          await IDB.set(key, data.data);
          await IDB.set(`__ts__${key}`, supaTs);
          return data.data;
        }
      } catch { /* offline → IDB fallback */ }
    }
    return IDB.get(key, def);
  },

  /** Scrive su IDB (immediato, con timestamp) + Supabase con retry */
  set: async (key: string, val: any) => {
    const ts = Date.now();
    await IDB.set(key, val);
    await IDB.set(`__ts__${key}`, ts);
    if (supabase && CLOUD_KEYS.has(key)) {
      try {
        const { error } = await supabase.from('app_data')
          .upsert({ key, data: val, updated_at: new Date(ts).toISOString() }, { onConflict: 'key' });
        if (error) {
          await supabase.from('app_data')
            .upsert({ key, data: val, updated_at: new Date(ts).toISOString() }, { onConflict: 'key' });
        }
      } catch { /* offline — IDB ha già il dato */ }
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
  // Solo email aziendali @inalcafb.com possono accedere
  if (!isCompanyEmail(email)) return null;
  try {
    const { data } = await supabase.from('user_roles').select('role').eq('email', email).maybeSingle();
    // user_roles contiene solo gli Admin; tutti gli altri @inalcafb.com sono Viewer automaticamente
    return (data?.role as 'admin' | 'viewer') ?? 'viewer';
  } catch { return 'viewer'; }
}

export async function listUsers(): Promise<{email: string, role: string, created_at: string}[]> {
  if (!supabase) return [];
  const { data } = await supabase.from('user_roles').select('*').order('created_at');
  return data ?? [];
}

export async function inviteUser(email: string, role: 'viewer' | 'admin' = 'admin') {
  if (!supabase) return;
  if (!isCompanyEmail(email)) throw new Error(`Solo email @inalcafb.com consentite`);
  // user_roles gestisce solo il ruolo Admin; il Viewer è automatico per tutti @inalcafb.com
  await supabase.from('user_roles').upsert({ email, role }, { onConflict: 'email' });
}

export async function removeUser(email: string) {
  if (!supabase) return;
  await supabase.from('user_roles').delete().eq('email', email);
}

export async function signInWithOtp(email: string) {
  if (!supabase) throw new Error('Supabase non configurato');
  if (!isCompanyEmail(email)) throw new Error(`Accesso riservato al personale Inalca. Usa la tua email @inalcafb.com`);
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
