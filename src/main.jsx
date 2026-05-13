import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import Auth from './Auth.jsx'
import { supabase } from './supabase.js'

const TABLE_MAP = {
  'sasoc:ventas:v1':         'ventas',
  'sasoc:clientes:v1':       'clientes',
  'sasoc:proveedores:v1':    'proveedores',
  'sasoc:facturas_prov:v1':  'facturas_prov',
  'sasoc:gastos:v1':         'gastos',
  'sasoc:pagos:v1':          'pagos',
  'sasoc:cf:cuentas:v1':     'cuentas',
  'sasoc:cf:movimientos:v1': 'movimientos',
  'sasoc:cotizaciones:v1':   'cotizaciones',
}

function toCamel(obj) {
  if (!obj || typeof obj !== 'object') return obj
  if (Array.isArray(obj)) return obj.map(toCamel)
  return Object.fromEntries(Object.entries(obj).map(([k,v]) => [k.replace(/_([a-z])/g,(_,c)=>c.toUpperCase()), toCamel(v)]))
}

function toSnake(obj) {
  if (!obj || typeof obj !== 'object') return obj
  if (Array.isArray(obj)) return obj.map(toSnake)
  return Object.fromEntries(Object.entries(obj).map(([k,v]) => [k.replace(/([A-Z])/g,'_$1').toLowerCase(), toSnake(v)]))
}

window.storage = {
  get: async (key) => {
    const table = TABLE_MAP[key]
    if (supabase && table) {
      const { data, error } = await supabase.from(table).select('*')
      if (!error && data) {
        const value = table === 'clientes'
          ? JSON.stringify(Object.fromEntries(data.map(r => [r.empresa, toCamel(r)])))
          : JSON.stringify(toCamel(data))
        localStorage.setItem(key, value)
        return { key, value }
      }
    }
    try { const val = localStorage.getItem(key); return val ? { key, value: val } : null } catch { return null }
  },
  set: async (key, value) => {
    try { localStorage.setItem(key, value) } catch {}
    const table = TABLE_MAP[key]
    if (supabase && table) {
      try {
        const parsed = JSON.parse(value)
        if (table === 'clientes') {
          const rows = Object.entries(parsed).map(([empresa, extra]) => toSnake({ empresa, ...extra }))
          if (rows.length) await supabase.from(table).upsert(rows)
        } else {
          const items = Array.isArray(parsed) ? parsed : []
          if (items.length) await supabase.from(table).upsert(items.map(toSnake))
        }
      } catch(e) { console.error('Supabase write error:', e) }
    }
    return { key, value }
  },
  delete: async (key) => { try { localStorage.removeItem(key) } catch {}; return { key, deleted: true } },
  list: async (prefix) => { try { return { keys: Object.keys(localStorage).filter(k => !prefix||k.startsWith(prefix)) } } catch { return { keys: [] } } }
}

function Root() {
  // Si Supabase está configurado, mostrar login
  // Si no, mostrar el CRM directamente (modo local)
  if (!supabase) {
    return <App user={null} onLogout={null} />
  }
  return (
    <Auth>
      {({ user, onLogout }) => <App user={user} onLogout={onLogout} />}
    </Auth>
  )
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode><Root /></React.StrictMode>
)
