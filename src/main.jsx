import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'

// ── Storage polyfill para Vercel (usa localStorage en vez de window.storage) ──
// En Claude Artifacts se usa window.storage. En Vercel usamos localStorage.
if (!window.storage) {
  window.storage = {
    get: async (key, shared) => {
      try {
        const val = localStorage.getItem(key)
        return val ? { key, value: val, shared: !!shared } : null
      } catch { return null }
    },
    set: async (key, value, shared) => {
      try {
        localStorage.setItem(key, value)
        return { key, value, shared: !!shared }
      } catch { return null }
    },
    delete: async (key) => {
      try {
        localStorage.removeItem(key)
        return { key, deleted: true }
      } catch { return null }
    },
    list: async (prefix) => {
      try {
        const keys = Object.keys(localStorage).filter(k => !prefix || k.startsWith(prefix))
        return { keys }
      } catch { return { keys: [] } }
    }
  }
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
)
