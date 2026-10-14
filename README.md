# CALIBRACIÓN-DRONES · Calibración de boquillas

App web para registrar la calibración de boquillas de los drones del Departamento de Malezas y Madurantes (Ingenio Madre Tierra).
Los datos se guardan en la hoja de Google Sheets **BASE DE DATOS CALIBRACION DRONES**, pestaña **CALIBRACIONES**. Las fotos se guardan en la carpeta de Drive.

## Archivos
| Archivo | Para qué |
|---|---|
| `index.html` | La app (GitHub Pages) |
| `logo.png`, `icon-192.png`, `icon-512.png`, `manifest.json` | Logo e ícono para instalar en Android |
| `Code.gs` | Backend: se copia en **Extensiones › Apps Script** de la hoja (no se sube a Pages) |

## 1. Apps Script
1. Abra la hoja › **Extensiones › Apps Script**.
2. Borre lo que haya en `Código.gs` y pegue todo `Code.gs`. Guarde.
3. Elija la función **`autorizar`** › **Ejecutar** › acepte los permisos (Hojas y Drive).
4. **Implementar › Nueva implementación** › tipo **Aplicación web**
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier persona**
5. Copie la URL que termina en `/exec`.
6. Pruebe en el navegador: `URL/exec?action=diag`. Revise que en `dronesDetectados` salgan sus drones con velocidad y ancho.

> Cada vez que cambie `Code.gs`: **Implementar › Gestionar implementaciones › ✏️ › Versión: Nueva versión**, así se conserva la misma URL.

## 2. GitHub Pages
1. En `index.html`, cambie `PEGAR_AQUI_LA_URL_DEL_APPS_SCRIPT` por la URL `/exec`.
2. Suba los archivos al repositorio.
3. **Settings › Pages** › Branch `main` / root › Save.
4. Abra la URL de Pages desde el celular › menú ⋮ › **Agregar a pantalla principal**.

## Cálculos (igual que la hoja)
- Caudal total (L/min) = L/ha teórico × velocidad (m/s) × ancho (m) × 60 ÷ 10 000
- Caudal por boquilla = caudal ÷ 2 · MIN / MAX = caudal por boquilla ± % de variación máxima (pestaña CALCULO DE FLUJOS)
- Boquilla (L/min) = litros recolectados ÷ tiempo de calibración (min)
- Variación B1/B2 = (boquilla − caudal por boquilla) ÷ caudal por boquilla
- Descarga general = B1 + B2 · Variación general = (descarga − caudal) ÷ caudal
- Significancia = (B1 − B2) ÷ L/ha teórico
- Número de descarga = R1, R2, R3… según cuántas veces se calibró ese dron ese día
- Estado = DENTRO DE RANGO si ambas boquillas están entre MIN y MAX

La app agrega por sí sola las columnas **LABOR, ESTADO, REGISTRADO POR, MARCA DE TIEMPO** al final de CALIBRACIONES.

## Permisos (pestaña USUARIOS)
- Para entrar, el usuario necesita tener una **CONTRASEÑA** en la hoja.
- **BOTONES EXTRAS PERMITIDOS**: si dice `MEDIDOR DE FLUJOS` puede ver el medidor; si dice `DASHBOARD`, puede ver el dashboard.
- Si el **CARGO** es `ADMINISTRADOR`, puede ver todo.
