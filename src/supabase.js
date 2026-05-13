import { createClient } from '@supabase/supabase-js'

const url  = import.meta.env.VITE_SUPABASE_URL
const key  = import.meta.env.VITE_SUPABASE_ANON_KEY

export const supabase = (url && key) ? createClient(url, key) : null

// ── CRUD helpers ──────────────────────────────────────────────────

// Cargar todos los registros de una tabla
export async function dbLoad(table) {
  if (!supabase) return null
  const { data, error } = await supabase.from(table).select('*')
  if (error) { console.error(`Error cargando ${table}:`, error); return null }
  return data
}

// Upsert (insert o update) un registro
export async function dbUpsert(table, row) {
  if (!supabase) return false
  const { error } = await supabase.from(table).upsert(row)
  if (error) { console.error(`Error guardando en ${table}:`, error); return false }
  return true
}

// Eliminar un registro por id
export async function dbDelete(table, id) {
  if (!supabase) return false
  const field = table === 'clientes' ? 'empresa' : 'id'
  const { error } = await supabase.from(table).delete().eq(field, id)
  if (error) { console.error(`Error eliminando de ${table}:`, error); return false }
  return true
}

// Upsert múltiples registros
export async function dbUpsertMany(table, rows) {
  if (!supabase || !rows.length) return false
  const { error } = await supabase.from(table).upsert(rows)
  if (error) { console.error(`Error en upsert masivo ${table}:`, error); return false }
  return true
}

// Verificar si Supabase está configurado
export const isSupabaseConfigured = !!supabase
