# Comparador Cotizaciones — Excel vs Supabase

Aplicación de escritorio (Windows y Mac) para comparar el desglose financiero del Excel **01-ESTADO DE CUENTA INTERNO** contra las cotizaciones en Supabase (proyecto NEO2).

**Repositorio:** [github.com/CarlosAHP/Pandas_Excel_comparacion](https://github.com/CarlosAHP/Pandas_Excel_comparacion)

---

## Obtener el proyecto desde GitHub

### Windows

1. Instala [Git para Windows](https://git-scm.com/download/win) (si no lo tienes).
2. Abre CMD o PowerShell y clona el repo:

```bat
cd %USERPROFILE%\Documents
git clone https://github.com/CarlosAHP/Pandas_Excel_comparacion.git
cd Pandas_Excel_comparacion
```

3. Sigue la [guía de instalación en Windows](#guía-de-instalación-en-windows) (Python, `.env`, `install.bat`, `run.bat`).

### Mac

```bash
cd ~/Documents
git clone https://github.com/CarlosAHP/Pandas_Excel_comparacion.git
cd Pandas_Excel_comparacion
./run.sh
```

> El archivo Excel **no** va en el repo. Coloca `01-ESTADO DE CUENTA INTERNO..xlsx` en la carpeta del proyecto o selecciónalo desde la app.

---

## Guía de instalación en Windows

Sigue estos pasos en orden. La instalación completa toma unos **10 minutos** la primera vez.

### Paso 1 — Instalar Python

1. Abre [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. Descarga **Python 3.12** (botón *Download Python 3.12.x*).
3. Ejecuta el instalador.
4. **Importante:** en la primera pantalla marca la casilla:
   - ✅ **Add python.exe to PATH**
5. Pulsa **Install Now** y espera a que termine.
6. Cierra y vuelve a abrir cualquier ventana de CMD si ya tenías una abierta.

**Comprobar que Python quedó bien:**

1. Presiona `Win + R`, escribe `cmd` y Enter.
2. Escribe:

```bat
python --version
```

Debe mostrar algo como `Python 3.12.x`.

3. Comprueba la interfaz gráfica:

```bat
python -c "import tkinter; print('OK')"
```

Si imprime `OK`, puedes continuar.

---

### Paso 2 — Copiar la carpeta del proyecto

Copia toda la carpeta `Exel_pandas_comparacion_PR` al PC, por ejemplo:

```
C:\Users\TuUsuario\Documentos\Exel_pandas_comparacion_PR
```

La carpeta debe contener al menos:

```
Exel_pandas_comparacion_PR/
├── install.bat          ← instalación (solo la primera vez)
├── run.bat              ← abrir la app (cada día)
├── requirements.txt
├── config/
├── src/
└── 01-ESTADO DE CUENTA INTERNO..xlsx   (opcional, se puede elegir otro)
```

Puedes recibir la carpeta por **USB**, **correo (zip)** o **Git**.

---

### Paso 3 — Configurar credenciales de Supabase

La app necesita conectarse a Supabase. Crea un archivo llamado **`.env`** dentro de la carpeta del proyecto.

**Opción A — Copiar del equipo de desarrollo**

Si alguien del equipo ya tiene la app funcionando, pide el archivo `.env` o `.env.local` y cópialo a la carpeta del proyecto.

**Opción B — Crearlo manualmente**

1. Abre el Bloc de notas.
2. Pega esto (con los valores reales):

```env
VITE_SUPABASE_URL=https://fvlrjcaltshccdbsseuk.supabase.co
SUPABASE_SERVICE_KEY=eyJ...tu_clave_aqui...
```

3. Guarda como:
   - **Nombre:** `.env`
   - **Tipo:** Todos los archivos (*.*)
   - **Ubicación:** la carpeta `Exel_pandas_comparacion_PR`

> También existe `.env.example` como plantilla. En Windows, si no ves la extensión, en el Explorador activa *Ver → Extensiones de nombre de archivo*.

**Variables necesarias:**

| Variable | Descripción |
|----------|-------------|
| `VITE_SUPABASE_URL` | URL del proyecto Supabase |
| `SUPABASE_SERVICE_KEY` | Service role key (no la anon key pública) |

---

### Paso 4 — Instalar la aplicación (solo la primera vez)

1. Abre la carpeta del proyecto en el Explorador de archivos.
2. Haz **doble clic** en **`install.bat`**.
3. Espera a que termine (crea el entorno virtual e instala librerías).
4. Al final debe decir **"Instalación completada"**. Pulsa una tecla para cerrar.

Si Windows muestra *"Windows protegió tu PC"*: clic en **Más información** → **Ejecutar de todas formas**.

---

### Paso 5 — Abrir la aplicación

Cada vez que quieras comparar:

1. Doble clic en **`run.bat`**
2. Se abre la ventana **"Comparador Cotizaciones — Excel vs Supabase (NEO2)"**

> `run.bat` también puede crear el entorno la primera vez si no ejecutaste `install.bat`, pero es mejor usar `install.bat` una vez para verificar que todo está correcto.

---

### Paso 6 — Usar la aplicación

1. **Archivo Excel:** por defecto carga `01-ESTADO DE CUENTA INTERNO..xlsx`. Puedes cambiarlo con *Seleccionar…*
2. Pulsa **Comparar** (puede tardar ~1 minuto mientras consulta Supabase).
3. Revisa los resultados agrupados por apartamento:
   - **Rojo** = concepto con diferencia entre Excel y Supabase
   - **Verde** = coincide
4. Usa el filtro *Solo unidades con diferencia* para ver solo lo que hay que revisar.
5. Pulsa **Exportar reporte** para generar un Excel con:
   - Instrucciones y mapeo de columnas Supabase
   - Hoja *Solo diferencias* con qué revisar en cada caso
   - Detalle por unidad

---

## Resumen rápido Windows

| Acción | Qué hacer |
|--------|-----------|
| Primera instalación | Python 3.12 → copiar carpeta → crear `.env` → `install.bat` |
| Uso diario | Doble clic en `run.bat` |
| Actualizar la app | Copiar la carpeta nueva (conserva tu `.env`) |
| Reinstalar dependencias | Ejecutar `install.bat` de nuevo |

---

## Solución de problemas (Windows)

### "Python no está en el PATH"

- Reinstala Python desde [python.org](https://www.python.org/downloads/) marcando **Add python.exe to PATH**.
- Cierra CMD y vuelve a abrir.

### "tkinter no está disponible"

- Reinstala Python desde python.org (no uses la versión de Microsoft Store si da problemas).
- En el instalador, elige **Customize** y asegúrate de que **tcl/tk and IDLE** esté marcado.

### Error de credenciales Supabase

- Verifica que `.env` existe en la misma carpeta que `run.bat`.
- Revisa que `VITE_SUPABASE_URL` y `SUPABASE_SERVICE_KEY` no tengan espacios extra.
- La clave debe ser la **service role key**, no la anon key.

### La ventana se cierra muy rápido

- Abre CMD, ve a la carpeta del proyecto (`cd C:\ruta\Exel_pandas_comparacion_PR`) y ejecuta `run.bat` para ver el mensaje de error.

### Antivirus bloquea `install.bat` o `run.bat`

- Añade la carpeta del proyecto a exclusiones del antivirus, o ejecuta como administrador solo si IT lo autoriza.

### "No se pudo crear el entorno virtual"

- Comprueba que tienes espacio en disco y permisos de escritura en la carpeta.
- No copies el proyecto a `C:\Program Files\` (mejor Documentos o Escritorio).

---

## Instalación en Mac

En macOS **no uses** el `python3` del sistema (Xcode): trae Tk 8.5 y la ventana queda en blanco.

```bash
brew install python@3.12 python-tk@3.12
cd Exel_pandas_comparacion_PR
python3.12 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Verifica Tk:

```bash
python -c "import tkinter as tk; print(tk.TkVersion)"   # 8.6+ (ideal 9.0)
```

**Uso en Mac:** `./run.sh` en terminal.

Si mezclaste Python 3.9 y 3.12, cierra terminales y usa solo `./run.sh`.

---

## Requisitos técnicos

- Python **3.11** o **3.12**
- Conexión a internet (consulta Supabase)
- Archivo `.env` o `.env.local` con credenciales

## Qué compara la herramienta

Por cada apartamento NEO2 con reserva activa en Supabase:

| Concepto | Excel | Supabase |
|----------|-------|----------|
| Reserva | Fila Reserva | `montoreserva_crmtratocotizacion` |
| Abono PVC | Promesa de compraventa | `montopromesa_crmtratocotizacion` |
| Enganche | Fila 33 (con reserva) | `engancheafraccionar + montoreserva` |
| Monto a financiar | Columna H/I | `montofinanciar_crmtratocotizacion` |
| Otros gastos | Gastos legales | `gastoscompranofinanc + otroscargosaddnofinanc` |
| Precio total | Precio total | `preciototalygastoscompranofin_crmtratocotizacion` |

Tolerancia de redondeo: **±0.01**

## Estructura del proyecto

```
config/field_mapping.yaml   # Mapeo Excel ↔ Supabase
src/
  main.py                   # Interfaz gráfica
  excel_loader.py           # Lectura del Excel
  supabase_client.py        # Consultas Supabase
  comparator.py             # Lógica de comparación
  report.py                 # Exportación del reporte
install.bat                 # Instalación Windows (primera vez)
run.bat                     # Abrir app Windows
run.sh                      # Abrir app Mac
.env.example                # Plantilla de credenciales
```

## Configuración avanzada

Edita `config/field_mapping.yaml` para ajustar etiquetas del Excel, campos Supabase o tolerancia.
