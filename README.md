# SIMER-ITM — Sistema Interactivo de Muestreo Estadístico para la Reacreditación Institucional

**Instituto Tecnológico Metropolitano (ITM) — Medellín, Colombia**  
**Proceso de Autoevaluación con Fines de Reacreditación Institucional de Alta Calidad (Acuerdo 01 de 2025 del CESU · CNA)**

---

## 👥 Destinatarios Ejecutivos y Autoría

- **Destinatarios Ejecutivos:**
  - **PhD. Hernán Salazar** — Vicerrector de Docencia (Líder General del Proceso de Reacreditación Institucional)
  - **Alexandra Montoya** — Directora de Gestión Académica
- **Autor e Investigador Principal:**
  - **Marvin Luis Jiménez Narváez** — Docente Ocasional, Departamento de Educación y Ciencias Básicas, Facultad de Ciencias Exactas y Aplicadas, ITM.

---

## ✨ Características Principales de `SIMER-ITM`

1. **Módulo de Parametrización Inicial Obligatoria (`P1` a `P12`):**
   - Panel interactivo inicial para configurar las **12 decisiones institucionales de depuración** sobre las anomalías detectadas en la auditoría forense de los marcos muestrales:
     - **Docentes (`P1`–`P4`):** Resolución de los 78 docentes duplicados por *cátedra extra* (`P1`), duplicados de mismo contrato según dominio `@correo.itm` vs `@itm` (`P2`), deduplicación exacta (`P3`) y tratamiento de los 239 registros en Cátedra sin correo electrónico (`P4`).
     - **Administrativos (`P5`–`P9`):** Tratamiento de las 19 plazas vacantes sin cédula (`BNLE`, `ANTIOQUIA 3`, `Vacante` — `P5`), duplicidad en vacante temporal de la cédula `71720089` (`VCACANTE - Robinson Sánchez Zambrano` — `P6`), recodificación de `DEPENDENCIA = "claudia patricia"` en la cédula `70382234` (`P7`), conductor en hoja `RECTORIA` sin correo con cédula `71611587` (`P8`), y selección entre **Censo ($N=211$)** o **Muestreo Estratificado** (`P9`).
     - **Empleadores (`P10`):** Selección de unidad estadística entre **Empresa única por NIT ($N = 303$)** o **Jefe / Contacto directo ($N = 531$)**.
     - **Graduados (`P11`):** Exclusión de graduados sin correo electrónico (**$N = 3,276$ con correo** vs. **$41,878$ sin correo**, sobre $45,154$ graduados únicos con `Último registro = 1`).
     - **Estudiantes (`P12` — Prioridad Institucional):** Selector entre **Forma 1** (`Facultad × Nivel Agrupado [Tecnológico, Universitario, Posgrado] × Modalidad`, $N = 29,904$) y **Forma 2** (`Estudiantes_Enriquecido.xlsx`, con $97.91\%$ de enlace histórico [$29,280 / 29,904$] y selección libre de variables categóricas y cuantitativas).

2. **Discretización Óptima de Variables Cuantitativas:**
   - Detección automática de variables continuas (`Promedio`, `Edad`, `Avance`, `Semestre`, `Horas Planeadas`, `Grado`) con soporte para:
     - **Regla Óptima de Dalenius-Hodges ($\text{acum }\sqrt{f}$)**.
     - **Cuantiles Empíricos (Terciles, Cuartiles, Quintiles)** y **Amplitud Igual / Sturges**.
     - **Puntos de Corte Personalizados** definidos por el usuario con histograma interactivo en vivo (Plotly.js).

3. **Motor Estadístico Exacto y Muestreo Bifásico ($n_{\text{inicial}} \to n_{\text{final}} - n_{\text{inicial}}$):**
   - Nivel de confianza de **$90\%$ a $99\%$ (paso de $1\%$)** y margen de error de **$1.0\%$ a $10.0\%$ (paso de $0.1\%$)**.
   - Muestreo Aleatorio Simple Sin Reemplazo (con FPC), Muestreo Aleatorio Estratificado con **Afijación Proporcional** y **Afijación Óptima de Neyman**, y **Censo**.
   - **Módulo de Muestra Preliminar ($n_{\text{inicial}} \ge 30$) vs. Muestra Definitiva ($n_{\text{final}}$):** Permite recalcular $n_{\text{final}, h}$ con las proporciones piloto $\hat{p}_{0h}$ y sortear únicamente los $n_{\text{restante}, h} = \max(0, n_{\text{final}, h} - n_{\text{inicial}, h})$ individuos faltantes sin repetir encuestados de la fase piloto.

4. **Protección Criptográfica de Datos Personales (Habeas Data — Ley 1581 de 2012):**
   - Todos los registros precargados están anonimizados mediante **SHA-256 con sal institucional**, sin exponer cédulas en texto plano, nombres propios, teléfonos ni correos personales.

5. **Ejecución Dual (`https://` en GitHub Pages y `file://` Local sin Servidor):**
   - Incluye tanto los archivos JSON estructurados en `data/*.json` como el paquete comprimido `data/embedded_data.js` para que funcione instantáneamente al abrir `index.html` con doble clic en cualquier computador o desde GitHub Pages.
