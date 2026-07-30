# 🤖 CONTEXTO RÁPIDO PARA IA - BIENVENIDA FACILITADORES CUN

> **Nota para cualquier Asistente IA (ChatGPT, Claude, Gemini, Antigravity)**:  
> Lee este archivo antes de realizar cualquier modificación. Contiene la arquitectura completa, las reglas de negocio y los puntos exactos de edición para trabajar de manera ultra rápida sin necesidad de inspeccionar el resto del proyecto.

---

## 📌 1. PROPÓSITO DEL PROYECTO
Procesar **exclusivamente los registros del reporte de Asistencia (CSV/Excel)** cargados en `Asistencia/`, aplicando una limpieza previa de caracteres especiales y codificación (estilo `procesar_reporte.ps1`), **enviando automáticamente a la pestaña `Duplicados` cualquier cédula duplicada o NO encontrada en Planta**, desplegando fórmulas nativas `=BUSCARV()` / `=VLOOKUP()` a lo largo de 1.500 filas en todas las columnas que consultan la pestaña **`PLANTA_BASE`**, y generando el **Tablero Web HTML Galáctico (con Logo CUN, desplegable/modal interactivo de Asistentes Válidos con Buscador en la Hero Card principal y popups por Escuela)**.

---

## 📁 2. MAPA DE ARCHIVOS Y ESTRUCTURA
```
Bienvenida Facilitadores/
├── Logo.png                             <-- Logo oficial CUN para la interfaz web
├── CONTEXTO_PROYECTO_IA.md              <-- (ESTE ARCHIVO) Contexto inmediato para la IA
├── ACTUALIZAR_TABLERO.bat                <-- Ejecutable de 1-Clic en Batch
├── Consolidado_Bienvenida_Facilitadores.xlsx <-- Excel generado con Fórmulas BUSCARV en 1.500 filas (Salida)
├── tablero.html                          <-- Tablero Web Galáctico principal (Salida)
├── tablero_respaldo.html                 <-- Copia autocontenida de respaldo
├── Asistencia/
│   ├── TCFRegistrodeAsistenciaFacilitadores*.csv <-- Reporte de asistencia de entrada
│   └── stap/procesar_reporte.ps1         <-- Script legacy de referencia
├── Planta/
│   └── PLANTA FACILITADORES PRINCIPAL 2026B.xlsx <-- Base de planta de entrada
└── Scripts/
    └── procesar_bienvenida.py            <-- 🌟 MOTOR PRINCIPAL EN PYTHON
```

---

## ⚙️ 3. REGLAS DE NEGOCIO OBLIGATORIAS

### A. Alcance Estricto del Consolidado
- **El consolidado principal y el tablero web contienen ÚNICAMENTE los registros válidos de asistencia**.
- La base completa de Planta se inserta con la `CEDULA` en Columna A dentro de la pestaña **`PLANTA_BASE`** para ser consultada dinámicamente mediante VLOOKUP.

### B. Estructura y Modales del Tablero HTML (Diseño Diapositiva 1)
- **Header**: Logo CUN original desplegado directamente en la cabecera.
- **Hero Card Principal Top (100% Ancho - `TOTAL REGISTROS ÚNICOS VÁLIDOS`)**:
  - Muestra la cifra total y barra animada del % sobre Planta (1.188 facilitadores).
  - **Interacción al Clic**: Al hacer clic en la tarjeta, abre una **Ventana Desplegable/Modal (`#modal-asistentes-completos`) con la lista completa de todos los asistentes válidos y un Buscador en tiempo real** para filtrar instantáneamente por Cédula, Nombre, Escuela, Regional o Modalidad.
- **Subgrid KPI**: Tarjeta `ASISTENCIA PRESENCIAL` (con modal general por regiones) y `ASISTENCIA VIRTUAL`.
- **Panel de Escuelas Dividido en 2 Columnas**:
  - **Columna Izquierda (`escuelas-panel`)**: Total por Escuela y barra de porcentaje.
  - **Columna Derecha (`escuelas-panel`)**: Conteo Virtual vs Presencial. Clic en cualquier escuela abre el modal emergente de desglose regional por sede (`#modal-escuela-detalle`).
- **Buscador Superior de Escuelas**: Filtra sincrónicamente ambas columnas en tiempo real.

### C. Despliegue de Fórmulas BUSCARV a lo largo de 1.500 Filas
- En las hojas `Consolidado_Unicos` y `Duplicados`, los campos recuperados de Planta contienen la fórmula **desplegada a lo largo de 1.500 filas**:
  - `INVITACIÓN` (Col A): `=IF(BN=" me","",IFERROR(VLOOKUP(BN, PLANTA_BASE!A:I, 2, FALSE), "NO ENCONTRADO EN PLANTA"))`
  - `CEDULA` (Col B): Número de Cédula de asistencia.
  - `NOMBRE` (Col C): Nombre limpio de asistencia o `=IF(BN=" me","",IFERROR(VLOOKUP(BN, PLANTA_BASE!A:I, 3, FALSE), ""))`
  - `Nombre Nivel 2` (Col D): `=IF(BN=" me","",IFERROR(VLOOKUP(BN, PLANTA_BASE!A:I, 5, FALSE), "NO ENCONTRADO EN PLANTA"))`
  - `Nombre Nivel 3` (Col E): `=IF(BN=" me","",IFERROR(VLOOKUP(BN, PLANTA_BASE!A:I, 6, FALSE), "NO ENCONTRADO EN PLANTA"))`
  - `Descripción Cargo` (Col F): `=IF(BN=" me","",IFERROR(VLOOKUP(BN, PLANTA_BASE!A:I, 8, FALSE), "DOCENTE / FACILITADOR"))`
  - `Nombre Centro Costo` (Col G): `=IF(BN=" me me","",IFERROR(VLOOKUP(BN, PLANTA_BASE!A:I, 9, FALSE), "GENERAL"))`

---

## 🛠️ 4. GUÍA DE MODIFICACIÓN RÁPIDA PARA LA IA

| Requerimiento del Usuario | Archivo a Editar | Qué Cambiar |
| :--- | :--- | :--- |
| **Modificar popups o estructura del HTML** | `Scripts/procesar_bienvenida.py` | Editar la plantilla `html_content` en Python y ejecutar `ACTUALIZAR_TABLERO.bat`. |
| **Añadir/Modificar fórmulas o columnas de Excel** | `Scripts/procesar_bienvenida.py` | Modificar la lógica de OpenPyXL y ejecutar el `.bat`. |
| **Ruta del ejecutable de Python** | `ACTUALIZAR_TABLERO.bat` | Cambiar la variable `PYTHON_EXE`. |

---

## 🚀 5. COMANDO DE EJECUCIÓN
```cmd
python Scripts/procesar_bienvenida.py
```
o ejecutando directamente `ACTUALIZAR_TABLERO.bat`.
