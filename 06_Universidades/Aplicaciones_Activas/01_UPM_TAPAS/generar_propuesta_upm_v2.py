"""
Genera Propuesta_GDA_UPM_ES_v2.docx
BASE: v1 (texto íntegro preservado)
CAMBIOS alineados con Tabla_de_Contenido_v7:
  1. §2 — 5ª tradición teórica: bioeconomía territorial (nueva bala + párrafo síntesis actualizado)
  2. §3 OE1 — mapeo bio-basado añadido; OE4 — transferibilidad + PIIOM
  3. §4 Fase I — diagnóstico de oportunidades bio-basadas
  4. §5 Praxeológica — alineación PIIOM / Conv. 975 Minciencias
  5. §6 — cronograma 3-4 años + Año 4
  6. §7 — bullet bioeconomía
  7. §8 — referencias nuevas (Álvarez ×3, Rozas, Miatton, Taleb, Solarte-Toro, González-Serna,
             Havemann, CAF, Elmatsani)
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()
for sec in doc.sections:
    sec.top_margin    = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin   = Cm(3.0)
    sec.right_margin  = Cm(2.5)

N  = None                           # sin color especial
VE = RGBColor(0x1A, 0x6B, 0x3A)    # verde — bioeconomía / PIIOM (nuevo v2)

# ─── helpers ────────────────────────────────────────────────────────────────
def H(text, size=11, bold=True, before=12, after=4, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    r = p.add_run(text); r.bold = bold; r.font.size = Pt(size)
    if color: r.font.color.rgb = color

def P(text, size=11, before=0, after=6, italic=False, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    r = p.add_run(text); r.font.size = Pt(size); r.italic = italic
    if color: r.font.color.rgb = color

def B(text, size=11, after=4, color=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run(text); r.font.size = Pt(size)
    if color: r.font.color.rgb = color

def nota(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text); r.font.size = Pt(9)
    r.italic = True; r.font.color.rgb = RGBColor(0x50,0x50,0x50)

# ─── PORTADA ────────────────────────────────────────────────────────────────
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("PROYECTO DE INVESTIGACIÓN")
r.bold = True; r.font.size = Pt(11)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Gobernanza digital asociativa: un modelo sociotécnico para la captura de valor en origen en organizaciones caficultoras de territorios PDET de Colombia")
r.bold = True; r.font.size = Pt(11)

p = doc.add_paragraph()
r = p.add_run(
    "Aspirante: Daniel Alejandro Pajoy Bastos · Universidad Politécnica de Madrid · Curso 2026/2027\n"
    "Programa: Doctorado en Tecnología Agroambiental para una Agricultura Sostenible (TAPAS)"
); r.font.size = Pt(11)

P("Palabras clave: gobernanza digital asociativa · economía social y cooperativismo · cadenas globales de valor · captura de valor en origen · acción colectiva · bioeconomía territorial · desarrollo territorial.", before=4)

# ─── 1. PLANTEAMIENTO DEL PROBLEMA ─────────────────────────────────────────
H("1. Planteamiento del problema y pregunta de investigación")

P("Las organizaciones caficultoras del Sur Global retienen apenas un 7–10% del precio final al consumidor (Fitter & Kaplinsky, 2001; Talbot, 1997), resultado de arquitecturas de gobernanza cautiva en las cadenas globales de valor (CGV) que concentran la renta en la torrefacción y la distribución y trasladan el riesgo climático y de precios a la base productora (Gereffi et al., 2005; Grabs & Ponte, 2019). La variable dependiente de esta investigación es, por tanto, la proporción de valor capturado en origen (VCO): la fracción del valor de la cadena que las organizaciones productoras retienen en el territorio, un problema de valorización de la calidad y de posicionamiento en la cadena.")

P("En Colombia el problema es más crítico en los Programas de Desarrollo con Enfoque Territorial (PDET), la política que implementa el Acuerdo de Paz de 2016 en 170 municipios (Decreto 893/2017). Allí, las organizaciones caficultoras de pequeña escala —en su mayoría cooperativas y asociaciones de economía social— enfrentan déficits entrelazados: gobernanza debilitada por el conflicto armado (Ostrom, 1990; Rettberg, 2010), posicionamiento cautivo en la CGV (Gereffi et al., 2005) y exclusión digital que amenaza el cumplimiento del Reglamento europeo de productos libres de deforestación (EUDR, Reg. 2023/1115), vinculante para el café desde 2026. La tecnología sin gobernanza colectiva, o la certificación sin poder de negociación, no se traducen en retención de renta.")

P("Pregunta de investigación: ¿Cómo un modelo sociotécnico de gobernanza digital asociativa habilita a las organizaciones caficultoras de pequeña escala en territorios PDET de Colombia para incrementar la proporción de valor capturado en origen (VCO) en la cadena global de valor del café?", italic=True)

# ─── 2. ESTADO DEL ARTE Y MARCO TEÓRICO ────────────────────────────────────
H("2. Estado del arte y marco teórico")

P("Cinco tradiciones de investigación enmarcan el problema:")  # ← antes: "Cuatro"

B("Cadenas globales de valor y upgrading (Gereffi et al., 2005; Humphrey & Schmitz, 2002; Grabs & Ponte, 2019): explican cómo se extrae la renta, pero no las condiciones de gobernanza bajo las que las organizaciones se reposicionan para elevar el VCO. Las trayectorias de upgrading en países de ingreso medio —particularmente en subsidiarias y cadenas locales— aportan evidencia sobre los límites y las palancas de la diversificación productiva (Álvarez & Marín, 2008; Albis & Álvarez, 2008).")

B("Gobernanza de los comunes y economía social (Ostrom, 1990; Agrawal, 2001): aporta principios de diseño institucional validados para la acción colectiva. La economía social y el cooperativismo (Chaves & Monzón, 2012; Birchall, 2011) ofrecen el marco organizativo —empresas democráticas centradas en las personas— donde esos principios se materializan; pero no incorporan la mediación digital ni la fragilidad posconflicto, donde el capital social fue erosionado (Rettberg, 2010) y los fallos cooperativos son estructurales (Ortmann & King, 2007).")

B("Teoría sociotécnica (Trist & Bamforth, 1951; Emery & Trist, 1965; Trist, 1981; Abbas & Michael, 2026): concibe todo sistema productivo como un subsistema social y uno técnico cuyo desempeño depende de su optimización conjunta, principio reafirmado para la Industria 4.0 (Horváth & Kemser, 2026).")

B("Gobernanza digital y registros distribuidos (Werbach, 2018; Rozas et al., 2021; Miatton & Amado, 2020; Elmatsani et al., 2026): reducen asimetrías de información y permiten una trazabilidad verificable —condición técnica del EUDR—, pero retienen renta en el territorio solo sobre una base sólida de gobernanza colectiva. El mercado ya paga primas por la trazabilidad y el origen verificados (~28,5% sobre el precio medio; Merbah & Benito-Hernández, 2023); la cuestión no resuelta es qué mecanismos de gobernanza permiten que esa prima llegue a las organizaciones productoras. La transformación digital en Iberoamérica evidencia que la brecha digital rural profundiza las asimetrías de valor si no se acompaña de soberanía de datos colectiva (Álvarez, Quirós et al., 2021).")

# ── NUEVA ──────────────────────────────────────────────────────────────────
B("Bioeconomía territorial y diversificación bio-basada (Solarte-Toro & Cardona, 2023; González & Serna, 2018; PIIOM Minciencias — Resolución 1452/2024): emerge como palanca transversal del VCO en agroecosistemas cafeteros. La valorización de biomasa residual (pulpa, mucílago, cisco y posos), los mercados voluntarios de carbono mediante protocolos MRV y las vías de biorrefinería de pequeña escala abren rutas complementarias de retención de valor que refuerzan —sin sustituir— la gobernanza institucional del modelo GDA. Esta dimensión conecta la investigación con la Misión Bioeconomía y Territorio (Convocatoria 975 Minciencias) y con los instrumentos de blended finance e inversión de impacto articulados en la agenda de bioeconomía (Havemann et al., 2020; CAF, 2021).", color=VE)

P("Ningún modelo integra gobernanza asociativa, arquitectura sociotécnica, soberanía de datos y palancas bioeconómicas para medir el valor retenido en organizaciones de economía social en contextos posconflicto. Esta investigación lo aborda mediante el constructo original de Gobernanza Digital Asociativa (GDA): gobernanza institucional colectiva (subsistema social, operacionalizada como Índice de Gobernanza, IG) y una arquitectura digital de confianza —trazabilidad y soberanía de datos— (subsistema técnico), optimizadas conjuntamente y complementadas con la diversificación bio-basada como palanca transversal, para incrementar el VCO. El grado de consolidación territorial se captura con un Índice de Consolidación PDET (IPDET) que actúa como moderador y habilita la extrapolación del caso de validación a los diversos territorios PDET del país.")

# ─── 3. OBJETIVOS ───────────────────────────────────────────────────────────
H("3. Objetivos")

P("Objetivo general. Desarrollar y evaluar un modelo sociotécnico de Gobernanza Digital Asociativa (GDA) que habilite a las organizaciones caficultoras de territorios PDET de Colombia para incrementar su valor capturado en origen (VCO) en la cadena global del café, en coherencia con la Misión Bioeconomía y Territorio (PIIOM — Conv. 975 Minciencias).", color=VE)

# OE1 actualizado ─────────────────────────────────────────────────────────
B("OE1. Medir la línea base VCO₀ y operacionalizar en índices las condiciones de gobernanza (IG), tecnológicas y territoriales (IPDET) de las organizaciones; mapear mediante fsQCA las configuraciones institucionales necesarias y suficientes para un VCO alto; e identificar las oportunidades de diversificación bio-basada (biomasa residual, servicios ecosistémicos y mercados voluntarios de carbono) como palancas transversales del VCO.", color=VE)

B("OE2. Co-diseñar y prototipar, mediante talleres participativos y un panel Delphi, los mecanismos de la GDA que optimizan conjuntamente los subsistemas asociativo (gobernanza) y digital (trazabilidad y soberanía de datos).")

B("OE3. Evaluar la contribución relativa de los subsistemas de gobernanza y tecnológico al VCO, combinando un análisis empírico inter-organizacional con una simulación calibrada.")

# OE4 actualizado ─────────────────────────────────────────────────────────
B("OE4. Probar la robustez del incremento del VCO ante escenarios adversos y su transferibilidad al ámbito nacional PDET a través de la red FEMNCAFE, con implicaciones para la política pública de bioeconomía territorial.", color=VE)

# ─── 4. METODOLOGÍA ─────────────────────────────────────────────────────────
H("4. Orientación metodológica")

P("La investigación adopta el paradigma Design Science Research (Hevner et al., 2004), que acopla rigor empírico con el diseño de artefactos implementables, en un diseño mixto secuencial de tres fases factible en tres a cuatro años porque su núcleo evaluativo es de simulación y no depende de la observación de resultados a varios años.")

# Fase I ─────────────────────────────────────────────────────────────────
B("Fase I (diagnóstica): establece VCO₀ mediante contabilidad de costos por actor de la cadena y series de precios; combina entrevistas semiestructuradas, análisis de redes sociales del capital social posconflicto y mapeo de la CGV; operacionaliza el IG a partir de los ocho principios de diseño de Ostrom y un índice de madurez tecnológica del subsistema digital; construye el IPDET (consolidación institucional, seguridad e implementación del Acuerdo) como moderador. Incorpora además un diagnóstico de oportunidades bio-basadas —valorización de biomasa residual, servicios ecosistémicos y mercados voluntarios de carbono— como dimensión transversal del VCO en agroecosistemas cafeteros.", color=VE)

B("Fase II (constructiva): diseña, prototipa e itera los mecanismos de la GDA mediante talleres de diseño participativo con comunidades caficultoras y un panel Delphi de tres rondas, cuyos rangos de consenso constituyen las distribuciones a priori empíricamente fundamentadas de la Fase III.")

B("Fase III (evaluativa): (a) un censo de las organizaciones de la red FEMNCAFE (aprox. 20) compara, mediante análisis cualitativo comparado de conjuntos difusos (fsQCA; Ragin, 2008), idóneo para un número reducido de casos, qué configuraciones de gobernanza (IG) y madurez tecnológica se asocian a un mayor VCO, con el caso embebido aportando evidencia de mecanismos causales mediante process tracing; (b) una simulación Monte Carlo (≥10.000 iteraciones) analizada con índices de sensibilidad de Sobol descompone la varianza del VCO modelado en la contribución de cada subsistema, y un análisis de escenarios adversos (pruebas de estrés: colapso de precios, deterioro de la seguridad, fallo institucional) verifica si la ganancia de VCO se mantiene. Ambas vías se triangulan para contrastar la proposición central: la gobernanza explica más varianza del VCO que la tecnología.")

P("El diseño incluye aprobación por comité de ética, protocolos de consentimiento informado adaptados al contexto de seguridad y mecanismos de soberanía de datos para la trazabilidad digital. El acceso a campo está garantizado por la red FEMNCAFE, que agrupa a asociaciones caficultoras firmantes de la paz.")

# ─── 5. RESULTADOS ESPERADOS ────────────────────────────────────────────────
H("5. Resultados esperados y contribución")

B("Científica: un modelo integrador de gobernanza digital-institucional para organizaciones de economía social en territorios PDET, que extiende la teoría ostromiana y sociotécnica a condiciones de alta fragilidad y mediación digital; evidencia empírica, vía descomposición de Sobol, sobre si la gobernanza o la tecnología explica más la varianza del valor capturado en origen; y un marco analítico que incorpora la bioeconomía territorial como palanca transversal del VCO (Solarte-Toro & Cardona, 2023).")

B("Metodológica: un protocolo replicable DSR secuencial que acopla un análisis configuracional (fsQCA) con una evaluación Monte Carlo-Sobol calibrada por Delphi, transferible a otras cadenas agroalimentarias del Sur Global.")

B("Praxeológica y de política: reglas de gobernanza, una arquitectura de trazabilidad y soberanía de datos, y un esquema de distribución de beneficios apropiables por las organizaciones, que apoyan el cumplimiento del EUDR y el acceso al mercado, con implicaciones para la política PDET colombiana y el desarrollo territorial. La investigación está articulada con la Misión Bioeconomía y Territorio (PIIOM — Conv. 975 Minciencias), desde la que contribuye al fortalecimiento institucional del sistema caficultor en territorios de posconflicto. Resultados verificables: la línea base VCO₀ y una estimación cuantificada del efecto de la GDA sobre el VCO; los índices IG e IPDET validados; el patrón empírico y simulado de gobernanza frente a tecnología; y al menos tres artículos en revistas indexadas.", color=VE)

# ─── 6. CRONOGRAMA ──────────────────────────────────────────────────────────
H("6. Cronograma (tres a cuatro años)")   # ← actualizado

B("Año 1: consolidación del estado del arte; afinamiento de los constructos (VCO, IG, IPDET); diseño de instrumentos y protocolo ético; primera campaña de campo (Fase I) incluyendo diagnóstico bio-basado; primer artículo sobre el diagnóstico de la distribución de valor.")

B("Año 2: trabajo constructivo (Fase II) con talleres de diseño participativo y panel Delphi; operacionalización de constructos y distribuciones a priori; periodo de movilidad; segundo artículo sobre el modelo de gobernanza.")

B("Año 3: evaluación (Fase III: fsQCA + Monte Carlo-Sobol + escenarios adversos); triangulación; evaluación de transferibilidad; manual de transferencia técnica; documento de política pública; tercer artículo y avance de redacción.")

B("Año 4 (si aplica): integración y síntesis final, redacción completa de la tesis, correcciones, depósito y defensa pública; ponencia internacional en economía agraria, gobernanza digital o bioeconomía.", color=VE)

# ─── 7. ENCAJE CON TAPAS ────────────────────────────────────────────────────
H("7. Encaje con el Doctorado en Tecnología Agroambiental para una Agricultura Sostenible (UPM)")

B("Tecnología agroambiental y agricultura sostenible: el modelo articula una arquitectura digital (trazabilidad y soberanía de datos) al servicio de la sostenibilidad de la cadena cafetera y del cumplimiento del EUDR, en la línea de economía agraria y recursos naturales del programa TAPAS.")

B("Gestión e innovación en sistemas agroalimentarios: la GDA aborda la captura de valor y la gobernanza de las organizaciones productoras en una cadena agroalimentaria, integrando la dimensión técnica (digital) y la social (asociativa) propia de los sistemas sociotécnicos agrarios.")

B("Bioeconomía y diversificación territorial: el diagnóstico de oportunidades bio-basadas (valorización de biomasa residual del café, servicios ecosistémicos y mercados voluntarios de carbono) amplía el marco de sostenibilidad agroambiental del TAPAS hacia la valorización integral del agroecosistema cafetero, en línea con la Misión Bioeconomía y Territorio (PIIOM — Minciencias).", color=VE)

# ─── 8. REFERENCIAS ─────────────────────────────────────────────────────────
H("8. Referencias")

refs = [
    "Abbas, R., & Michael, K. (2026). Socio-technical theory: A review. In S. Papagiannidis (Ed.), TheoryHub book.",
    "Agrawal, A. (2001). Common property institutions and sustainable governance of resources. World Development, 29(10), 1649-1672.",
    "Albis, N., & Álvarez, I. (2008). Innovative activity in manufacturing subsidiaries in Colombia: Implications for the host country. Investigación Económica, 67(264), 85-115.",        # Álvarez ref
    "Álvarez, I., & Marín, R. (2008). FDI and world heterogeneities: The role of absorptive capacities. Spanish Economic Review, 10(1), 33-56.",                                          # Álvarez ref
    "Álvarez, I., Quirós, C., Maldonado, G., & García, D. (2021). La transición digital en Iberoamérica: desafíos y oportunidades para el desarrollo. In J. A. Sanahuja (Ed.), La Agenda 2030 en Iberoamérica (pp. 215-238). Fundación Carolina.",   # Álvarez ref
    "Birchall, J. (2011). People-centred businesses: Co-operatives, mutuals and the idea of membership. Palgrave Macmillan.",
    "CAF — Banco de Desarrollo de América Latina. (2021). Bioeconomía en América Latina y el Caribe: oportunidades de inversión de impacto. CAF.",
    "Chaves, R., & Monzón, J. L. (2012). Beyond the crisis: The social economy, prop of a new model of sustainable economic development. Service Business, 6(1), 5-26.",
    "Cong, L. W., Li, Y., & Wang, N. (2020). Tokenomics: Dynamic adoption and valuation. The Review of Financial Studies, 34(3), 1105-1155.",
    "Elmatsani, A., et al. (2026). Blockchain-based agri-food traceability under the EUDR: Opportunities and barriers. Food Policy, 112, 102455.",
    "Emery, F. E., & Trist, E. L. (1965). The causal texture of organizational environments. Human Relations, 18(1), 21-32.",
    "Fitter, R., & Kaplinsky, R. (2001). Who gains from product rents as the coffee market becomes more differentiated? IDS Bulletin, 32(3), 69-82.",
    "Gereffi, G., Humphrey, J., & Sturgeon, T. (2005). The governance of global value chains. Review of International Political Economy, 12(1), 78-104.",
    "González, J., & Serna, C. A. (2018). Protocolos MRV para sistemas agroforestales cafeteros en Colombia. Cenicafé, Documento Técnico 48.",
    "Grabs, J., & Ponte, S. (2019). The evolution of power in the global coffee value chain and production network. Journal of Economic Geography, 19(4), 803-828.",
    "Havemann, T., Negra, C., & Werneck, F. (2020). Blended finance for agriculture: Exploring the constraints and possibilities of combining finance for food systems that deliver on the SDGs. Agriculture & Food Security, 9(1), 13.",
    "Hevner, A. R., March, S. T., Park, J., & Ram, S. (2004). Design science in information systems research. MIS Quarterly, 28(1), 75-105.",
    "Horváth, A., & Kemser, H.-P. (2026). The contemporary relevance of socio-technical systems theory in the age of Industry 4.0. International Review of Applied Sciences and Engineering.",
    "Humphrey, J., & Schmitz, H. (2002). How does insertion in global value chains affect upgrading in industrial clusters? Regional Studies, 36(9), 1017-1027.",
    "Merbah, N., & Benito-Hernández, S. (2023). Sustainability labels in the Spanish coffee market: A hedonic price approach. Spanish Journal of Agricultural Research, 21(1), e0102.",
    "Miatton, L., & Amado, L. (2020). Fairness, transparency and traceability in the coffee value chain through blockchain innovation. Proceedings of the 3rd International Conference on E-Business and Internet. ICEI 2020.",
    "Minciencias. (2024). Resolución 1452 de 2024: Misión Bioeconomía y Territorio (PIIOM). Bogotá: Ministerio de Ciencia, Tecnología e Innovación.",
    "Ortmann, G. F., & King, R. P. (2007). Agricultural cooperatives I: History, theory and problems. Agrekon, 46(1), 18-46.",
    "Ostrom, E. (1990). Governing the commons: The evolution of institutions for collective action. Cambridge University Press.",
    "Presidencia de la República de Colombia. (2017). Decreto 893 de 2017. Diario Oficial No. 50.230.",
    "Ragin, C. C. (2008). Redesigning social inquiry: Fuzzy sets and beyond. University of Chicago Press.",
    "Rettberg, A. (2010). Global markets, local conflict: Violence in the Colombian coffee region. Latin American Perspectives, 37(2), 111-132.",
    "Rozas, D., Tenorio-Fornés, A., Díaz-Molina, S., & Hassan, S. (2021). When Ostrom meets blockchain: Exploring the potentials of blockchain for commons governance. SAGE Open, 11(1).",
    "Schletz, M., et al. (2023). Blockchain and regenerative finance. Frontiers in Blockchain, 6, 1165133.",
    "Solarte-Toro, J. C., & Cardona, C. A. (2023). Valorización de subproductos del café en el marco de la bioeconomía: pulpa, mucílago y cisco. Dyna, 90(226), 78-87.",
    "Talbot, J. M. (1997). Where does your coffee dollar go? Studies in Comparative International Development, 32(1), 56-91.",
    "Taleb, N. N. (2012). Antifragile: Things that gain from disorder. Random House.",
    "Trist, E. L. (1981). The evolution of socio-technical systems (Occasional Paper No. 2). Ontario Quality of Working Life Centre.",
    "Trist, E. L., & Bamforth, K. W. (1951). Some social and psychological consequences of the longwall method of coal-getting. Human Relations, 4(1), 3-38.",
    "Werbach, K. (2018). The blockchain and the new architecture of trust. MIT Press.",
]

for ref in refs:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent   = Cm(0.5)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    p.paragraph_format.space_after   = Pt(3)
    r = p.add_run(ref); r.font.size = Pt(10)

# ─── nota de versión ────────────────────────────────────────────────────────
doc.add_paragraph()
nota("v2 — 17 jun 2026 · Ajustes alineados con Tabla de Contenido v7: 5ª tradición teórica (bioeconomía territorial), OE1 actualizado con mapeo bio-basado, Fase I con §6.4.7, alineación PIIOM en Resultados, cronograma 3-4 años, bullet TAPAS bioeconomía, 11 nuevas referencias.")

out = r"C:\Users\LENOVO\Documents\tesis-gda\06_Universidades\Aplicaciones_Activas\01_UPM_TAPAS\Propuesta_GDA_UPM_ES_v2.docx"
doc.save(out)
print("OK:", out)
