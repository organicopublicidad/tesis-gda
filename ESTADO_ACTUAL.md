# 📍 ESTADO ACTUAL — Proyecto GDA (LEER PRIMERO)

> **Propósito:** este archivo es el punto de arranque para CUALQUIER conversación nueva.
> Léelo completo antes de continuar. Contiene todas las decisiones cerradas.
> **Última actualización:** 2026-06-09

---

## 🎯 Identidad del proyecto

- **Candidato:** Daniel Alejandro Pajoy Bastos
- **Tesis:** Gobernanza Digital Asociativa (GDA) — café colombiano en territorios PDET
- **Director/tutor de apoyo:** asistente actuando como director pre-doctoral
- **Regla de eficiencia:** NO subir PDFs a Claude. Procesar local (pdfplumber/python-docx), convertir a MD/txt antes de leer. Python: `C:\Users\LENOVO\AppData\Local\Programs\Python\Python313\python.exe`

---

## ✅ DECISIONES CERRADAS (no reabrir sin causa)

### Pregunta de investigación (DEFINITIVA)
> *"¿Cómo un modelo sociotécnico de gobernanza digital asociativa habilita a las organizaciones caficultoras de pequeña escala en territorios PDET de Colombia para incrementar la proporción de valor capturado en origen (VCO) en la cadena global de valor del café?"*

EN: *"How does a socio-technical model of associative digital governance enable smallholder coffee organisations in Colombia's PDET territories to increase the share of value captured at origin (VCO) in the global coffee value chain?"*

### Título (Opción A, elegida)
> *"Gobernanza digital asociativa: un modelo sociotécnico para la captura de valor en origen en organizaciones caficultoras de territorios PDET de Colombia"*

### Palabras clave (5)
Gobernanza digital asociativa · Cadenas globales de valor · Captura de valor en origen · Acción colectiva · Soberanía de datos

### Constructo y método
- **Constructo:** GDA (ES) / ADG (EN) = subsistema social (gobernanza asociativa, Ostrom) + subsistema técnico (digital: trazabilidad, soberanía de datos), **optimizados conjuntamente** (teoría sociotécnica: Trist & Bamforth 1951 → Emery & Trist 1965 → Trist 1981 → Abbas & Michael 2026; vigencia Horváth & Kemser 2026). **GDA NO se cita** (es aporte original).
- **Variable dependiente:** VCO (proporción de valor capturado en origen). VCO₀ = línea base.
- **Hipótesis central:** la gobernanza explica más varianza del VCO que la tecnología (se prueba con Monte Carlo + índices de Sobol).
- **Moderador:** IPDET (Índice de Consolidación PDET) — clave para la validez externa del alcance "de Colombia" (validación en un caso, extrapolación nacional).
- **Método:** Design Science Research, mixto secuencial, 3 fases (Diagnóstica → Constructiva/Delphi → Evaluativa/Monte Carlo). Caso embebido + encuesta red FEMNCAFE.

### Decisiones de poda (IMPORTANTE)
- ❌ **Antifragilidad ELIMINADA** como objeto de investigación (no está en pregunta, ni texto, ni objetivos). Salieron Taleb (2012) y Teece et al. (1997).
- ❌ **DMDU eliminado** (camino A): la prueba de robustez se hace como "análisis de escenarios adversos (pruebas de estrés)" dentro del mismo Monte Carlo.
- ✅ **IPDET conservado** (imprescindible).
- ✅ Población: "organizaciones" caficultoras (no "asociaciones") — evita eco con "asociativa".
- ✅ Alcance: "PDET de Colombia" (nacional), no solo Cauca. Cauca/Popayán = caso de validación.

---

## 📄 ESTADO DE LAS SECCIONES (plantilla de 5 partes)

Las propuestas usan plantilla de 5 partes idéntica para las 22 universidades; solo cambia el párrafo de encaje al final de la Sección 3.

| Sección | Estado |
|---|---|
| 1. Título provisional y palabras clave | ✅ Listo |
| 2. Planteamiento del problema y justificación | ✅ Listo (~605 pal.) |
| 3. Estado del arte y encuadre teórico | ✅ Listo (núcleo ~615 pal. + encaje por universidad) |
| 4. Orientación metodológica | ✅ Listo (~450 pal.) |
| 5. Referencias | ✅ Listo (25 núcleo APA; UPM +2) |

**Generados:** 44 documentos `.docx` (22 universidades × ES/EN), uno por carpeta:
- `Propuesta_GDA_ES.docx` y `Research_Proposal_GDA_EN.docx`

---

## 🏛️ UNIVERSIDADES MAPEADAS (22)

**España (`06_Universidades/Aplicaciones_Activas/`):**
01_UPM_TAPAS ⭐APROBADA (Dra. Sonia Benito) · 02_Deusto · 03_Loyola · 04_Mondragon · 05_UAB · 06_UPV · 07_UV_Valencia · 08_UAM_Madrid · 09_UPC_Cataluna

**Internacionales (`06_Universidades/Aplicaciones_Internacionales/`):**
07_Bologna · 08_Firenze · 09_ZEF_Bonn · 10_Wageningen · 11_TU_Delft · 12_Utrecht · 13_Lund · 14_Helsinki_HELSUS · 15_Stockholm_Resilience · 16_Copenhagen_IFRO · 17_SLU_Suecia · 18_Aarhus_BSS · 19_NMBU_Noragric

Detalle de programas, becas y plazos: `06_Universidades/00_DOCUMENTO_MAESTRO_Doctorado.md`

---

## 🛠️ ARCHIVOS CLAVE

- **Generador (regenera las 44 propuestas):** `tesis-gda/generar_propuestas.py` — editar el núcleo una vez y reejecutar regenera todo.
- **Biblioteca de referencia (519 fuentes):** `C:\Users\LENOVO\Documents\Tesis de doctorado\Approved\` + índice `00_INDICE_MAESTRO.md`
- **Textos OCR / extraídos:** `...\Approved\Extracted\` (incl. Trist 1981 completo)
- **Propuesta original Minciencias (v0.1):** `Tesis de doctorado\Versiones de propuestas\Nuevo Modelo Institucional v 0.1.docx`

---

## ⏭️ PENDIENTES (próximos pasos)

1. **Objetivos formales** OG + OE1-OE4 (mapeados a las 3 fases; OE4 = evaluar efecto del modelo sobre VCO y su robustez — SIN antifragilidad). *No redactados aún.*
2. **Limpiar "antifragilidad"** de documentos de planeación: `Aplicaciones_Activas/00_INDICE.md` (línea ~24) y la tabla de énfasis de UPM (menciona Taleb).
3. **Propuesta formal completa por universidad** (siguiente fase — documento extenso, no la de 5 partes).
4. **UPM — Tabla de contenido** que pidió la Dra. Sonia Benito (entregable específico).
5. Decidir commit/push periódicos.

---

## 📌 Evaluación Minciencias (contexto)
La propuesta v0.1 enviada a la Conv. 975 "Becas para el Cambio" supera el umbral en criterios valorativos (45+20+20 = 85/85); el resultado final depende del Criterio 4 (enfoque diferencial, objetivo/documental). Rúbrica con niveles discretos (no escala continua).
