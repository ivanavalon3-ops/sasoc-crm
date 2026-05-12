# SASOC Merch — CRM Empresarial

CRM completo para empresa de merchandising. Módulos: Dashboard, Ventas, Clientes, Cotizaciones, Proveedores, Gastos/Pagos, Cash Flow y Estadísticas.

## Deploy en Vercel (5 minutos)

### Opción A — Desde GitHub (recomendado)

1. Subí esta carpeta a un repositorio de GitHub:
   - Entrá a [github.com/new](https://github.com/new)
   - Creá un repo llamado `sasoc-crm` (privado si querés)
   - Subí todos estos archivos

2. Conectá con Vercel:
   - Entrá a [vercel.com](https://vercel.com) → **New Project**
   - Seleccioná tu repo `sasoc-crm`
   - Vercel detecta Vite automáticamente
   - Hacé clic en **Deploy**
   - En 2 minutos tenés tu URL: `sasoc-crm.vercel.app`

### Opción B — Drag & Drop (más rápido)

1. Instalá las dependencias y generá el build:
   ```bash
   npm install
   npm run build
   ```
2. Entrá a [vercel.com](https://vercel.com)
3. Arrastrá la carpeta `dist/` a la pantalla de Vercel
4. ¡Listo! Te da la URL al instante.

## Desarrollo local

```bash
npm install
npm run dev
```

Abre http://localhost:5173

## Estructura

```
sasoc-crm/
├── src/
│   ├── main.jsx      # Entry point + storage polyfill
│   └── App.jsx       # CRM completo (todos los módulos)
├── public/
│   └── favicon.svg
├── index.html
├── vite.config.js
├── vercel.json
└── package.json
```

## Notas

- **Storage**: En Vercel los datos se guardan en `localStorage` del navegador. Los datos son por dispositivo/navegador.
- **Multi-usuario**: Para datos compartidos entre usuarios, el próximo paso es conectar Supabase (base de datos gratuita).
- **URL personalizada**: En Vercel podés agregar tu propio dominio gratis en Settings → Domains.
