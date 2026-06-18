# 📍 ESTADO ACTUAL — Proyecto GDA (LEER PRIMERO)

> **Propósito:** este archivo es el punto de arranque para CUALQUIER conversación nueva.
> Léelo completo antes de continuar. Contiene todas las decisiones cerradas.
> **Última actualización:** 2026-06-17

---

## 🎯 Identidad del proyecto

- **Candidato:** Daniel Alejandro Pajoy Bastos · daniel@unicauca.edu.co · +57 317 538 5353
- **Tesis:** Gobernanza Digital Asociativa (GDA) — café colombiano en territorios PDET
- **Director/tutor de apoyo:** asistente actuando como director pre-doctoral
- **Regla de eficiencia:** NO subir PDFs a Claude. Procesar local (pdfplumber / python-docx / fitz). Python: `C:\Users\LENOVO\AppData\Local\Programs\Python\Python313\python.exe`
- **PDF desde Word:** PowerShell COM, `SaveAs` formato `17`.
- **Directorio tesis:** `C:\Users\LENOVO\Documents\tesis-gda`
- **Repositorio GitHub:** https://github.com/organicopublicidad/tesis-gda (rama master)

---

## ✅ DECISIONES CERRADAS DE TESIS (no reabrir sin causa)

### Pregunta de investigación (DEFINITIVA)
> *"¿Cómo un modelo sociotécnico de gobernanza digital asociativa habilita a las organizaciones caficultoras de pequeña escala en territorios PDET de Colombia para incrementar la proporción de valor capturado en origen (VCO) en la cadena global de valor del café?"*

EN: *"How does a socio-technical model of associative digital governance enable smallholder coffee organisations in Colombia's PDET territories to increase the share of value captured at origin (VCO) in the global coffee value chain?"*

### Título (Opción A, elegida)
> *"Gobernanza digital asociativa: un modelo sociotécnico para la captura de valor en origen en organizaciones caficultoras de territorios PDET de Colombia"*

### Palabras clave (7 en v8)
Gobernanza digital asociativa · Cadenas globales de valor · Captura de valor en origen · Acción colectiva · Soberanía de datos · Bioeconomía territorial · Desarrollo territorial

### Constructo y método
- **GDA / ADG (EN):** subsistema social (gobernanza asociativa, Ostrom) + subsistema técnico (trazabilidad, soberanía de datos), **optimizados conjuntamente**. Teoría sociotécnica: Trist & Bamforth 1951 → Emery & Trist 1965 → Trist 1981 → Abbas & Michael 2026 (vigencia Horváth & Kemser 2026). GDA NO se cita como término ajeno (es aporte original).
- **Variable dependiente:** VCO. VCO₀ = línea base.
- **Hipótesis central:** la dimensión de gobernanza (subsistema social) explica una mayor proporción de la varianza del VCO que la dimensión tecnológica (subsistema técnico); se evidencia que la optimización conjunta —no solo lo digital— es la condición habilitante. Contrastada con descomposición de varianza de Sobol.
- **Moderador:** IPDET (Índice de Consolidación PDET).
- **Método:** Design Science Research (Hevner et al., 2004), diseño mixto secuencial, 3 fases. fsQCA (censo FEMNCAFE ~20 orgs) + Monte Carlo-Sobol + escenarios adversos (pruebas de estrés).
- **Bioeconomía territorial:** palanca transversal del VCO (§1.4.4, §5.6, §6.4.7) — NO es eje ni variable dependiente. Conecta con PIIOM Minciencias.

### Decisiones de poda (CRÍTICO — no reintroducir)
- ❌ **Antifragilidad ELIMINADA como eje** — solo se menciona brevemente en §4.5.4 como contexto teórico (Taleb, 2012); la robustez se opera mediante escenarios adversos en Cap. 8. IAT-Café queda en Anexo 6 como instrumento paralelo.
- ❌ **DMDU eliminado** — la robustez va dentro del Monte Carlo como escenarios adversos.
- ✅ IPDET conservado · Población: "organizaciones" (no "asociaciones") · Alcance: PDET de Colombia; Cauca = caso de validación.

---

## 🏛️ ESTADO DE UNIVERSIDADES — 17 jun 2026

### INTERNACIONALES — ENVIADAS ✅

| Universidad | Programa | Estado | Hito próximo |
|---|---|---|---|
| **Firenze** 🇮🇹 | GeSoRAFA, curriculum EASRS (ciclo XLII) | ✅ ENVIADA + pagada (€30) | Entrevista remota **14 jul, 9:30 CEST** |
| **Bologna** 🇮🇹 | STAAA/DISTAL, Topic 4 (42.º ciclo) | ✅ ENVIADA + pagada (€10) · país corregido Colombia ✅ · cuenta: `daniel.pajoybastos@studio.unibo.it` | Resultados títulos **10 jul** · Oral **21 jul, 9:00 CEST** (MS Teams) |
| **Sapienza Roma** 🇮🇹 | PhD 42.º ciclo — convocatoria ordinaria D.R. 998/2026 | ✅ ENVIADA (17 jun 2026, antes de las 14:00 CET) | Pendiente resultado |

### INTERNACIONALES — PENDIENTES DE ACCIÓN

| Universidad | Deadline | Prioridad | Próximo paso |
|---|---|---|---|
| **ETH Zurich** 🇨🇭 | **26 jun** | Media — tema fijo suizo | Decidir sí/no |
| **ZEF-Bonn DAAD** 🇩🇪 | ~31 ago (RE-VERIFICAR) | Alta — beca completa para colombianos | Verificar fecha + preparar research proposal 3–5 págs |
| **Wageningen** 🇳🇱 | Rolling | Alta | Revisar vacantes semanalmente; identificar 2–3 supervisores |
| **CIRAD / Montpellier** 🇫🇷 | Rolling | Alta — el más afín (café/Sur Global) | Identificar supervisor y contactar |
| **Helsinki HELSUS** 🇫🇮 | 24 ago – 7 sep | Media | Preparar dossier |
| **Hohenheim FSC** 🇩🇪 | Verificar 2026/27 | Media | Verificar ventana |

Reserva: Stockholm Resilience · TU Delft · Utrecht · Lund · Copenhagen IFRO · SLU · Aarhus · NMBU.

### ESPAÑOLAS — INSCRITAS / PREINSCRITAS

| Universidad | Programa | Estado | Próximo hito |
|---|---|---|---|
| **UPM** 🇪🇸 ⭐ | TAPAS (agroambiental) | ✅ APROBADA — Dra. Sonia Benito (virtual) · ToC v8 + Propuesta v3 enviados 17 jun | Respuesta Sonia · Matrícula formal |
| **UCM** 🇪🇸 | ADE (1.º) + Innovación (2.º) | ✅ Preinscrita (DA028954) · ❌ Lejarriaga · ❌ Quirós · ❌ Merino (UAM) · ⏳ Bel Durán · ⏳ Heijs · **Entrevista M.I. Álvarez: 17 jun** | Resolución admisión 14 oct |
| **UAB** 🇪🇸 | Doctorat en Empresa | ✅ INSCRITO — tutor: Miguel Á. García Cestona · matrícula desde 17 jul | Confirmar director |
| **Deusto** 🇪🇸 | CETIS | 🔄 Preinscrito | Stand-by |
| **Loyola** 🇪🇸 | Desarrollo Inclusivo y Sostenible | 🔄 Preinscrito | Stand-by |
| **Mondragón** 🇪🇸 | Gestión Avanzada de Organizaciones | 🔄 Preinscrito | Stand-by |

### ESPAÑOLAS — PENDIENTES DE DECISIÓN

| Universidad | Programa | Deadline | Acción |
|---|---|---|---|
| **UV** 🇪🇸 | Economía Social — IUDESCOOP | **23 jun, 14:00** | Pagar tasa equivalencia 155,22€ + preinscripción Escola de Doctorat — **recordatorio programado 22 jun** |
| **UPV** 🇪🇸 | Agroalimentario | **15–26 jun** | Decidir sí/no (depende de UV) |

### ESPAÑOLAS — SUSPENDIDAS / DESCARTADAS

- **UAM** ❌ — descartada. ❌ Carlos Merino respondió negativamente (17 jun). Sin avalador activo.
- **UPC** ❌ — descartada.

---

## 📄 DOCUMENTOS ENVIADOS A SONIA BENITO (UPM) — 17 jun 2026

| Documento | Versión enviada | Archivo |
|---|---|---|
| Tabla de Contenido (ToC) | **v8** (correcciones de Daniel sobre v7) | `08_Capitulos/Tabla_de_Contenido_v8.docx` + PDF |
| Propuesta de investigación | **v3** (PDF generado desde v2 con correcciones de Daniel) | `01_UPM_TAPAS/Propuesta_GDA_UPM_ES_v3.pdf` |

### Cambios principales de Daniel en v8 vs v7 (generada por asistente)
- Títulos de sección: eliminadas todas las citas de autores (se dejan en el texto)
- Terminología unificada: "subsistema social y técnico" en lugar de variantes anteriores
- OG actualizado: "en coherencia con la Misión Bioeconomía y Territorio de las Políticas Orientadas por Misiones PIIOM"
- Anexo 6 renombrado: "Documentos de Política Pública" (neutro para UPM; oculta vínculo contractual Minciencias)
- Tabla cronograma: productos A6.1–A6.6 asignados por año

### Discrepancias menores pendientes de armonizar (NO urgentes)
1. **Número de versión:** la ToC v8 aún dice "Versión: 0.7" en la portada (no actualizado a 0.8)
2. **OE3:** difiere entre ToC v8 (parametrización/operacionalización) y Propuesta v3 (evaluación relativa); ambas coherentes pero no idénticas

---

## ⚠️ CORRECCIONES PENDIENTES PARA PRÓXIMA INTERVENCIÓN

> Aplicar cuando haya cambios que justifiquen generar nuevas versiones. NO generar versiones solo para esto.

### 1. Error de co-autores — Álvarez, Quirós et al. (2021) ← CRÍTICO
- **Referencia incorrecta en Propuesta v3:** `Álvarez, I., Quirós, C., Maldonado, G., & García, D. (2021)`
- **Referencia correcta (verificada en PDF `La_transicion_digital.pdf`):**
  > Álvarez, I., Quirós, C., Marín, R., Medina, L., & Biurrun, A. (2021). La transformación digital en Iberoamérica: una oportunidad para la inclusión en la era pos-COVID-19. En J. A. Sanahuja (Ed.), *La Agenda 2030 en Iberoamérica*. Fundación Carolina.
- Afecta: `generar_propuesta_upm_v2.py` (línea de referencia) y cualquier propuesta que la incluya.

### 2. WP08-09 = Álvarez, Marín & Maldonado (no "Álvarez & Marín")
- `WP_08-09.pdf` confirma: "Internal and external factors of competitiveness in the middle-income countries" — autores: **Isabel Álvarez, Raquel Marín, Georgina Maldonado** (WP 08/09, ICEI)
- Referencia correcta: `Álvarez, I., Marín, R., & Maldonado, G. (2008). Internal and external factors of competitiveness in the middle-income countries. WP 08/09. ICEI Working Papers.`

### 3. Archivo en biblioteca a renombrar
- `Approved/03_Capitulos_de_Libro/Alvarez-Quiros_2021_TransicionDigital_Iberoamerica.pdf`
  → renombrar a `Alvarez-Quiros-Marin_2021_TransicionDigital_Iberoamerica.pdf` (incluir Marín)

---

## 📄 PROPUESTAS — Archivos canónicos

**Propuesta EN (plantilla Bologna):**
`06_Universidades/Aplicaciones_Internacionales/07_Bologna/Research Proposal - STAAA_Daniel_Pajoy_v3.docx`

**Propuesta ES (base española):**
`06_Universidades/Aplicaciones_Activas/07_UV_Valencia/Proyecto_Investigacion_UV_ES.docx`

**CV:** `CV_Europass_Daniel_Pajoy_ES.docx` (el preferido para universidades españolas)

Propuestas customizadas generadas:
- **UPM** → `01_UPM_TAPAS/Propuesta_GDA_UPM_ES_v2.docx` · v3 (Daniel) enviada a Sonia 17 jun
- Deusto → `02_Deusto/Propuesta_GDA_Deusto_ES.docx`
- Loyola → `03_Loyola/Propuesta_GDA_Loyola_ES.docx`
- Mondragón → `04_Mondragon/Propuesta_GDA_Mondragon_ES.docx`
- UAB → `05_UAB/Propuesta_GDA_UAB_ES.docx`
- UPV → `06_UPV/Propuesta_GDA_UPV_ES.docx`
- UV → `07_UV_Valencia/Proyecto_Investigacion_UV_ES.docx`
- UCM (ADE) → `10_UCM_Complutense/Proyecto_Investigacion_UCM_ADE_ES.docx`
- UCM (Innovación) → `10_UCM_Complutense/Proyecto_Investigacion_UCM_Innovacion_ES.docx`

---

## 📝 TABLA DE CONTENIDO (ToC) — Historial de versiones

| Versión | Archivo | Estado | Novedades |
|---|---|---|---|
| v3 | `00_Tabla_de_Contenido_v3.md` | Referencia base | Estructura original 3 niveles |
| v5 | `Tabla_de_Contenido_v5.docx` | Superada | fsQCA integrado; numeración corregida; §5.6 renombrado |
| v6 | `Tabla_de_Contenido_v6.docx` | Correcciones Daniel | Hipótesis "optimización conjunta"; subtítulos fases; simplificaciones Cap. 8 |
| v7 | `Tabla_de_Contenido_v7.docx` | Generada por asistente | +§1.4.4 bioeconomía; +§4.5.4 fragilidad/robustez; +§5.6 bioeconomía (shift 5.6→5.7→5.8); +§6.4.7; +Anexo 6; refs Álvarez ×3 |
| **v8** | `Tabla_de_Contenido_v8.docx` + PDF | **ENVIADA A SONIA** | Correcciones Daniel sobre v7: citas fuera de títulos; PIIOM simplificado; Anexo 6 = "Documentos de Política Pública"; cronograma detallado |

**Script generador v7:** `08_Capitulos/generar_toc_v7.py`

---

## 🛠️ ARCHIVOS CLAVE

- **Este archivo:** `ESTADO_ACTUAL.md` — leer primero en cada conversación
- **Maestro de universidades + bitácora:** `06_Universidades/00_DOCUMENTO_MAESTRO_Doctorado.md`
- **Prompt para nueva conversación:** `PROMPT_NUEVA_CONVERSACION.md`
- **Biblioteca de referencia:** `C:\Users\LENOVO\Documents\Tesis de doctorado\Approved\` + `00_INDICE_MAESTRO.md`
  - `02_Libros/` — libros completos incl. Sanahuja 2021, Ostrom 1990, Taleb 2012
  - `03_Capitulos_de_Libro/` — capítulos incl. Álvarez-Quirós 2021 (TransicionDigital)
  - `06_Working_Papers/` — WPs Álvarez (×7), ICEI, NBER, SSRN, etc.
- **Credenciales personales:** `06_Universidades/00_Credenciales/` (diplomas traducidos — NO subir a repositorio público)
- **Scripts Python activos:**
  - `08_Capitulos/generar_toc_v7.py` — genera ToC v7 (base de v8)
  - `06_Universidades/Aplicaciones_Activas/01_UPM_TAPAS/generar_propuesta_upm_v2.py` — genera Propuesta UPM v2 (base de v3)

---

## ⏭️ PENDIENTES — Por urgencia

| Prioridad | Acción | Deadline |
|---|---|---|
| 🔴 | Decidir UV (tasa 155,22€ + preinscripción) — **recordatorio programado 22 jun** | 23 jun |
| 🔴 | Decidir ETH Zurich | 26 jun |
| 🔴 | Decidir UPV (depende de UV) | 26 jun |
| 🟠 | Preparar defensa oral Firenze — presentación + ensayo | 14 jul |
| 🟠 | Preparar defensa oral Bologna (MS Teams) | 21 jul |
| 🟠 | Esperar respuesta de Sonia Benito sobre ToC v8 + Propuesta v3 | — |
| 🟡 | UCM: resultado entrevista M.I. Álvarez (17 jun) → confirmar o descartar como directora | Pronto |
| 🟡 | UAB: confirmar director (García Cestona como tutor ya confirmado; director pendiente) | Pronto |
| 🟡 | ZEF-Bonn: verificar deadline exacto + preparar research proposal 3–5 págs | ~31 ago |
| 🟡 | CIRAD / Montpellier: identificar supervisor y contactar | Rolling |
| 🟡 | Wageningen: revisar vacantes semanalmente; identificar 2–3 supervisores | Rolling |
| 🟡 | Helsinki HELSUS: preparar dossier | 24 ago–7 sep |
| 🔵 | Correcciones bibliográficas pendientes (ver sección ⚠️) — al próximo cambio en propuestas | Próxima iteración |
| 🔵 | UV profesores (6 correos de aval): verificar emails y enviar | Pronto |

---

## 📌 Evaluación Minciencias (contexto)
- Propuesta v0.2 Conv. 975 "Becas para el Cambio" — supera umbral criterios valorativos (45+20+20 = 85/85).
- Resultado final depende del Criterio 4 (enfoque diferencial, objetivo/documental).
- **PIIOM Misión Bioeconomía y Territorio** (Resolución 1452/2024): integrada en ToC v7/v8 y Propuesta v3 como palanca transversal, no como eje. Anexo 6 contiene compromisos de productos.
- Productos comprometidos: IAT-Café (Tipo B), Arquitectura GDA (Tipo B), Protocolo MRV (Tipo B), Manual transferencia (Tipo C), Documento política pública (Tipo C), Protocolo taller co-creación (Tipo C).
