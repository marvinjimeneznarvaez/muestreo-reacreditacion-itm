/**
 * ============================================================================
 * SIMER-ITM — Sistema Interactivo de Muestreo Estadístico para la Reacreditación Institucional
 * Motor Estadístico, Parametrización P1–P12, Estratificación Óptima, Muestreo Bifásico y Sorteo
 * Autor: Marvin Luis Jiménez Narváez — Docente Ocasional, Facultad de Ciencias Exactas y Aplicadas
 * ============================================================================
 */

(function () {
  'use strict';

  // --------------------------------------------------------------------------
  // 1. CUANTILES EXACTOS NORMAL ESTÁNDAR Z_{1 - alpha/2} (90% a 99%, paso 1%)
  // --------------------------------------------------------------------------
  const Z_TABLE = {
    90: 1.6448536269514722,
    91: 1.6953977102721358,
    92: 1.7506860712521692,
    93: 1.8119106729525978,
    94: 1.8807936081512509,
    95: 1.959963984540054,
    96: 2.0537489106318225,
    97: 2.17009037758456,
    98: 2.3263478740408408,
    99: 2.5758293035489004
  };

  // Paleta Institucional Luminosa para Gráficos Plotly
  const PALETTE = [
    '#003865', '#0D9488', '#059669', '#D97706', '#4F46E5',
    '#E11D48', '#0284C7', '#7C3AED', '#EA580C', '#0891B2'
  ];

  // --------------------------------------------------------------------------
  // 2. DEFINICIÓN DE METADATOS DE VARIABLES POR ESTAMENTO
  // --------------------------------------------------------------------------
  const ESTAMENTO_META = {
    estudiantes: {
      key: 'estudiantes',
      label: 'Estudiantes',
      icon: '🎓',
      sourceFile: 'Estudiantes.xlsx + Estudiantes_Enriquecido.xlsx',
      rawCount: 29904,
      variables: [
        { id: 'facultad', label: 'Facultad', type: 'cat', forma: 1 },
        { id: 'nivel_agrupado', label: 'Nivel Agrupado (Tecnológico, Universitario, Posgrado)', type: 'cat', forma: 1 },
        { id: 'modalidad', label: 'Modalidad (Presencial, Virtual, Presencial-Virtual)', type: 'cat', forma: 1 },
        { id: 'nivel', label: 'Nivel Específico (5 niveles: Tecnología, Prof., Esp., Maestría, Doct.)', type: 'cat', forma: 2 },
        { id: 'tipo_programa', label: 'Tipo de Programa (Pregrado vs Posgrado)', type: 'cat', forma: 2 },
        { id: 'lugar', label: 'Lugar / Sede (Medellín vs Otro)', type: 'cat', forma: 2 },
        { id: 'sexo', label: 'Sexo Biológico / Registrado (Enriquecido)', type: 'cat', forma: 2 },
        { id: 'estrato', label: 'Estrato Socioeconómico (Enriquecido)', type: 'cat', forma: 2 },
        { id: 'ciclo_formativo', label: 'Ciclo Formativo Semestral (Enriquecido)', type: 'cat', forma: 2 },
        { id: 'promedio', label: 'Promedio Crédito Acumulado [Cuantitativa]', type: 'quant', forma: 2, defaultCuts: [3.2, 3.8, 4.2], unit: 'pts' },
        { id: 'edad', label: 'Edad del Estudiante [Cuantitativa]', type: 'quant', forma: 2, defaultCuts: [20, 24, 30], unit: 'años' },
        { id: 'avance', label: 'Porcentaje de Avance en Carrera [Cuantitativa]', type: 'quant', forma: 2, defaultCuts: [25, 55, 80], unit: '%' },
        { id: 'semestre', label: 'Semestre / Nivel Académico [Cuantitativa]', type: 'quant', forma: 2, defaultCuts: [3, 6, 8], unit: 'sem' }
      ],
      defaultVarsForma1: ['facultad', 'nivel_agrupado', 'modalidad'],
      defaultVarsForma2: ['facultad', 'ciclo_formativo', 'sexo']
    },
    docentes: {
      key: 'docentes',
      label: 'Docentes / Profesores',
      icon: '👨‍🏫',
      sourceFile: 'Profesores.xlsx',
      rawCount: 1778,
      variables: [
        { id: 'tipo_vinculacion', label: 'Tipo de Vinculación (Carrera, Ocasional, Cátedra)', type: 'cat' }
      ],
      defaultVars: ['tipo_vinculacion']
    },
    administrativos: {
      key: 'administrativos',
      label: 'Personal Administrativo y Directivo',
      icon: '🏛️',
      sourceFile: 'Administrativos.xlsx (Hojas RECTORIA + ADMINISTRATIVOS)',
      rawCount: 232,
      variables: [
        { id: 'nivel_jerarquico', label: 'Nivel Jerárquico (Directivo, Profesional, Técnico, Asistencial)', type: 'cat' }
      ],
      defaultVars: ['nivel_jerarquico']
    },
    empleadores: {
      key: 'empleadores',
      label: 'Empleadores / Sector Externo',
      icon: '🏢',
      sourceFile: 'Empleadores.xlsx',
      rawCount: 531,
      variables: [],
      defaultVars: []
    },
    graduados: {
      key: 'graduados',
      label: 'Graduados / Egresados',
      icon: '🎖️',
      sourceFile: 'Graduados.xlsx (Último registro = 1)',
      rawCount: 45154,
      variables: [
        { id: 'facultad', label: 'Facultad de Egreso', type: 'cat' },
        { id: 'nivel_agrupado', label: 'Nivel Agrupado (Tecnológico, Universitario, Posgrado)', type: 'cat' },
        { id: 'modalidad', label: 'Modalidad del Programa', type: 'cat' }
      ],
      defaultVars: ['facultad', 'nivel_agrupado']
    },
    custom: {
      key: 'custom',
      label: 'Dataset Propio Cargado (.xlsx / .csv)',
      icon: '📂',
      sourceFile: 'Archivo Local del Usuario',
      rawCount: 0,
      variables: [],
      defaultVars: []
    }
  };

  // --------------------------------------------------------------------------
  // 3. ESTADO GLOBAL DE LA APLICACIÓN (STATE)
  // --------------------------------------------------------------------------
  const STATE = {
    rawDatasets: {
      estudiantes: [],
      docentes: [],
      administrativos: [],
      empleadores: { empresas: [], jefes: [] },
      graduados: [],
      custom: []
    },
    filteredDatasets: {
      estudiantes: [],
      docentes: [],
      administrativos: [],
      empleadores: [],
      graduados: [],
      custom: []
    },
    // Los Parámetros Institucionales (P1 a P12 + P_COMISION) con sus valores por defecto exactos
    params: {
      P_COMISION: 'EXCLUIR', // Excluir de Docentes a los 11 administrativos "En comisión" (16 filas)
      P1: 'CATEDRA',         // Eliminar CÁTEDRA en duplicados por cátedra extra
      P2: 'CORREO_ITM',      // Eliminar @correo.itm en duplicados de mismo contrato
      P3: 'DEJAR_UNO',       // Deduplicación exacta
      P4: 'EXCLUIR',         // Excluir docentes sin correo (239 en Cátedra)
      P5: 'EXCLUIR',         // Excluir 19 vacantes sin cédula en Administrativos
      P6: 'EXCLUIR',         // Excluir vacante temporal s/correo de 71720089
      P7: 'NO_EXCLUIR',      // No excluir 70382234 ("claudia patricia")
      P8: 'NO_EXCLUIR',      // No excluir conductor 71611587 sin correo en Rectoría (acuerdo reunión)
      P9: 'MUESTREO',        // Administrativos: MAE por Nivel Jerárquico por defecto (acuerdo reunión)
      P10: 'EMPRESA',        // Empleadores: Empresa única por NIT (303) por defecto
      P11: 'EXCLUIR',        // Graduados: Excluir sin correo (3,276) por defecto
      P12: 'FORMA_1'         // Estudiantes: Forma 1 por defecto
    },
    activeEstamento: 'estudiantes',
    activeTab: 'tab-muestreo',
    onlyLinkedStudentsForma2: true,
    forma2Scope: 'LINKED_ALL', // LINKED_ALL (29,280) | LINKED_MA (23,898) | ALL_29904 (29,904)

    // Filtro de elegibilidad por Promedio Mínimo en Estudiantes (post-cálculo de n_h)
    studentGpaFilter: {
      minGpa: 2.0,
      zeroRule: 'STRICT_WITH_FALLBACK' // STRICT_WITH_FALLBACK | INCLUDE_ZERO | STRICT_PURE
    },

    // Configuración de muestreo por estamento
    samplingConfig: {
      design: 'PROPORCIONAL', // PROPORCIONAL | NEYMAN | MAS | CENSO
      confidence: 95,
      marginError: 5.0,
      propP: 0.50,
      tnr: 0,
      microStrataCensus: true,
      seed: 20261111
    },

    // Variables de estratificación seleccionadas por estamento
    selectedVarsByEstamento: {
      estudiantes: ['facultad', 'nivel_agrupado', 'modalidad'],
      docentes: ['tipo_vinculacion'],
      administrativos: ['nivel_jerarquico'],
      empleadores: [],
      graduados: ['facultad', 'nivel_agrupado'],
      custom: []
    },

    // Configuración de Discretización de Variables Cuantitativas
    quantConfig: {
      activeVar: 'promedio',
      method: 'DALENIUS', // DALENIUS | QUANTILES | EQUAL_WIDTH | CUSTOM
      numBins: 3,
      customCutsByVar: {
        promedio: [3.2, 3.8, 4.2],
        edad: [20, 24, 30],
        avance: [25, 55, 80],
        semestre: [3, 6, 8],
        horas_planeadas: [160, 400, 800],
        grado: [3, 6, 12],
        contactos_en_empresa: [1, 3, 6]
      },
      separateZero: true,
      computedCutsByVar: {}
    },

    // Resultados activos del cálculo y sorteo
    currentResult: null,
    drawnSample: [],

    // Estado de Muestreo Bifásico (n_inicial vs n_final - n_inicial)
    bifasico: {
      pilotTotalN: 60,
      allocation: 'NEYMAN',
      pilotPropsByStratum: {},
      strataResults: [],
      drawnPilot: [],
      drawnRemaining: []
    }
  };

  // --------------------------------------------------------------------------
  // 4. GENERADOR DE NÚMEROS PSEUDOALEATORIOS REPRODUCIBLE (MULBERRY32)
  // --------------------------------------------------------------------------
  function createPRNG(seed) {
    let a = (Number(seed) || 20261111) >>> 0;
    return function () {
      let t = (a += 0x6D2B79F5);
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  function sampleWithoutReplacement(array, k, prng) {
    const n = array.length;
    if (k >= n) return array.slice();
    if (k <= 0) return [];
    const copy = array.slice();
    for (let i = 0; i < k; i++) {
      const r = i + Math.floor(prng() * (n - i));
      const tmp = copy[i];
      copy[i] = copy[r];
      copy[r] = tmp;
    }
    return copy.slice(0, k);
  }

  function formatNumber(n, decimals = 0) {
    if (n === null || n === undefined || Number.isNaN(Number(n))) return '--';
    const num = Number(n);
    const sign = num < 0 ? '-' : '';
    const absFixed = Math.abs(num).toFixed(decimals);
    const parts = absFixed.split('.');
    // Convención del proyecto: coma (,) para miles y punto (.) para decimales
    const intFormatted = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return decimals > 0 ? `${sign}${intFormatted}.${parts[1]}` : `${sign}${intFormatted}`;
  }

  // --------------------------------------------------------------------------
  // 5. CARGA E HIDRATACIÓN DE LOS 5 DATASETS DEL ITM
  // --------------------------------------------------------------------------
  function hydrateTableToObjects(tableObj, type) {
    if (!tableObj || !tableObj.columns || !tableObj.rows) return [];
    const cols = tableObj.columns;
    return tableObj.rows.map((r) => {
      const obj = {};
      for (let i = 0; i < cols.length; i++) {
        obj[cols[i]] = r[i];
      }
      obj.id_registro = obj.id;
      obj.hash_id = obj.hash;
      if (type === 'graduados') {
        obj.tiene_correo = Boolean(obj.tiene_correo);
      }
      return obj;
    });
  }

  async function loadAllDatasets() {
    let sourceData = window.SIMER_DATA || null;

    if (!sourceData) {
      try {
        const [adm, doc, emp, est, grad] = await Promise.all([
          fetch('data/administrativos_data.json').then((r) => r.json()),
          fetch('data/docentes_data.json').then((r) => r.json()),
          fetch('data/empleadores_data.json').then((r) => r.json()),
          fetch('data/estudiantes_data.json').then((r) => r.json()),
          fetch('data/graduados_data.json').then((r) => r.json())
        ]);
        sourceData = {
          administrativos: adm,
          docentes: doc,
          empleadores: emp,
          estudiantes: est,
          graduados: grad
        };
      } catch (err) {
        console.error('Error cargando JSONs por fetch:', err);
      }
    }

    if (!sourceData) return;

    STATE.rawDatasets.administrativos = sourceData.administrativos || [];
    STATE.rawDatasets.docentes = sourceData.docentes || [];
    STATE.rawDatasets.empleadores = {
      empresas: (sourceData.empleadores && sourceData.empleadores.empresas) ? sourceData.empleadores.empresas.map((e) => ({
        ...e,
        hash_id: e.hash_nit,
        tiene_correo_label: e.tiene_correo ? 'Con Correo' : 'Sin Correo'
      })) : [],
      jefes: (sourceData.empleadores && sourceData.empleadores.jefes) ? sourceData.empleadores.jefes.map((j) => ({
        ...j,
        hash_id: j.hash_nit,
        tiene_correo_label: j.tiene_correo ? 'Con Correo' : 'Sin Correo'
      })) : []
    };
    STATE.rawDatasets.estudiantes = hydrateTableToObjects(sourceData.estudiantes, 'estudiantes');
    STATE.rawDatasets.graduados = hydrateTableToObjects(sourceData.graduados, 'graduados');
  }

  // --------------------------------------------------------------------------
  // 6. MOTOR DE PARAMETRIZACIÓN INICIAL (P1 A P12)
  // --------------------------------------------------------------------------
  function applyParameterizationRules() {
    const P = STATE.params;

    // 1. ESTUDIANTES (P12: Forma 1 [29,904] vs Forma 2 [29,280 o 23,898])
    let estList = STATE.rawDatasets.estudiantes;
    if (P.P12 === 'FORMA_2') {
      if (STATE.forma2Scope === 'LINKED_MA') {
        estList = estList.filter((r) => r.enlazado === 1 && r.avance !== null && r.avance !== undefined);
      } else if (STATE.forma2Scope === 'LINKED_ALL') {
        estList = estList.filter((r) => r.enlazado === 1);
      }
      // Si es ALL_29904 conserva los 29,904
    }
    STATE.filteredDatasets.estudiantes = estList;

    // 2. DOCENTES (P_COMISION, P1, P2, P3, P4)
    let docList = STATE.rawDatasets.docentes.slice();

    // P_COMISION: Administrativos en situación "En comisión" (11 directivos = 16 registros en Profesores.xlsx)
    if (P.P_COMISION === 'EXCLUIR') {
      docList = docList.filter((r) => !r.en_comision_admin);
    }

    // P1: Docentes duplicados con dos tipos distintos de contrato (Carrera/Ocasional + Cátedra extra)
    if (P.P1 === 'OTRO') {
      // Eliminar el registro de CARRERA / OCASIONAL y dejar CÁTEDRA
      docList = docList.filter((r) => !(r.dup_distinto_contrato && r.tipo_contrato_raw !== 'CÁTEDRA'));
    } else {
      // Por defecto ('CATEDRA'): Eliminar el registro de CÁTEDRA en los duplicados con distinto contrato
      docList = docList.filter((r) => !(r.dup_distinto_contrato && r.tipo_contrato_raw === 'CÁTEDRA'));
    }

    // P2: Docentes duplicados con la misma categoría en tipo de contrato
    if (P.P2 === 'ITM') {
      docList = docList.filter((r) => !(r.dup_mismo_contrato && r.dominio_correo === '@itm.edu.co'));
    } else {
      // Por defecto ('CORREO_ITM'): Eliminar @correo.itm y conservar @itm.edu.co
      docList = docList.filter((r) => !(r.dup_mismo_contrato && r.dominio_correo === '@correo.itm.edu.co'));
    }

    // P4: Docentes sin correo electrónico (239 registros en Cátedra)
    if (P.P4 === 'EXCLUIR') {
      docList = docList.filter((r) => r.tiene_correo);
    }
    STATE.filteredDatasets.docentes = docList;

    // 3. ADMINISTRATIVOS (P5, P6, P7, P8, P9)
    let admList = STATE.rawDatasets.administrativos.slice();
    if (P.P5 === 'EXCLUIR') {
      admList = admList.filter((r) => !r.caso_p5_sin_cedula);
    }
    if (P.P6 === 'EXCLUIR') {
      // Excluir específicamente el registro sin correo en vacante temporal de la cédula 71720089
      admList = admList.filter((r) => !(r.caso_p6_71720089 && !r.tiene_correo));
    }
    if (P.P7 === 'EXCLUIR') {
      admList = admList.filter((r) => !r.caso_p7_70382234);
    }
    if (P.P8 === 'EXCLUIR') {
      admList = admList.filter((r) => !r.caso_p8_71611587);
    }
    STATE.filteredDatasets.administrativos = admList;

    // 4. EMPLEADORES (P10: Empresa = 303 vs Jefe = 531)
    STATE.filteredDatasets.empleadores =
      P.P10 === 'EMPRESA'
        ? STATE.rawDatasets.empleadores.empresas
        : STATE.rawDatasets.empleadores.jefes;

    // 5. GRADUADOS (P11: Excluir sin correo = 3,276 vs No excluir = 45,154)
    STATE.filteredDatasets.graduados =
      P.P11 === 'EXCLUIR'
        ? STATE.rawDatasets.graduados.filter((r) => r.tiene_correo)
        : STATE.rawDatasets.graduados;

    updateModalAndRibbonCounts();
  }

  // --------------------------------------------------------------------------
  // 7. MOTOR DE DISCRETIZACIÓN ÓPTIMA DE VARIABLES CUANTITATIVAS
  //    (Dalenius-Hodges acum √f, Cuantiles, Amplitud Igual y Personalizado)
  // --------------------------------------------------------------------------
  function isVariableQuantitative(estamento, varId) {
    const meta = ESTAMENTO_META[estamento];
    if (!meta) return false;
    const v = meta.variables.find((item) => item.id === varId);
    return Boolean(v && v.type === 'quant');
  }

  function computeQuantitativeCuts(values, method, L, customCuts, separateZero) {
    const cleanVals = values
      .map((v) => Number(v))
      .filter((v) => !Number.isNaN(v) && (separateZero ? v > 0 : true))
      .sort((a, b) => a - b);

    if (cleanVals.length === 0) return [];
    const minV = cleanVals[0];
    const maxV = cleanVals[cleanVals.length - 1];
    if (minV === maxV) return [minV];

    if (method === 'CUSTOM') {
      return (customCuts || [])
        .map((c) => Number(c))
        .filter((c) => !Number.isNaN(c) && c > minV && c < maxV)
        .sort((a, b) => a - b);
    }

    if (method === 'EQUAL_WIDTH') {
      const cuts = [];
      const step = (maxV - minV) / L;
      for (let k = 1; k < L; k++) {
        cuts.push(Number((minV + k * step).toFixed(2)));
      }
      return cuts;
    }

    if (method === 'QUANTILES') {
      const cuts = [];
      for (let k = 1; k < L; k++) {
        const idx = Math.floor((k / L) * (cleanVals.length - 1));
        const val = Number(cleanVals[idx].toFixed(2));
        if (!cuts.includes(val) && val > minV && val < maxV) {
          cuts.push(val);
        }
      }
      return cuts;
    }

    // Por defecto: DALENIUS-HODGES (acumulación de raíz cuadrada de frecuencias √f)
    const numFineBins = 50;
    const binWidth = (maxV - minV) / numFineBins;
    const freqs = new Array(numFineBins).fill(0);

    for (let i = 0; i < cleanVals.length; i++) {
      let b = Math.floor((cleanVals[i] - minV) / binWidth);
      if (b >= numFineBins) b = numFineBins - 1;
      if (b < 0) b = 0;
      freqs[b]++;
    }

    const cumSqrt = new Array(numFineBins).fill(0);
    let running = 0;
    for (let b = 0; b < numFineBins; b++) {
      running += Math.sqrt(freqs[b]);
      cumSqrt[b] = running;
    }

    const totalSqrt = running;
    const cuts = [];
    for (let k = 1; k < L; k++) {
      const target = (k / L) * totalSqrt;
      let foundBin = 0;
      for (let b = 0; b < numFineBins; b++) {
        if (cumSqrt[b] >= target) {
          foundBin = b;
          break;
        }
      }
      const cutVal = Number((minV + (foundBin + 1) * binWidth).toFixed(2));
      if (!cuts.includes(cutVal) && cutVal > minV && cutVal < maxV) {
        cuts.push(cutVal);
      }
    }
    return cuts;
  }

  function formatQuantInterval(val, cuts, separateZero, varId) {
    const num = Number(val);
    if ( Number.isNaN(num) ) return 'Sin Dato';
    if (separateZero && num <= 0) {
      return varId === 'promedio' ? '0.0 (Nuevo / Sin Prom.)' : '0 (Valor Cero)';
    }
    if (!cuts || cuts.length === 0) {
      return `${num.toFixed(1)}`;
    }
    for (let i = 0; i < cuts.length; i++) {
      if (num <= cuts[i]) {
        const low = i === 0 ? (separateZero ? '>0' : 'Min') : cuts[i - 1];
        return `Niv.${i + 1} (${low} a ${cuts[i]}]`;
      }
    }
    return `Niv.${cuts.length + 1} (>${cuts[cuts.length - 1]})`;
  }

  // --------------------------------------------------------------------------
  // 8. MOTOR ESTADÍSTICO DE CÁLCULO DE TAMAÑO MUESTRAL Y AFIJACIÓN
  // --------------------------------------------------------------------------
  function computeSamplingForEstamento(estamentoKey, overrideConfig) {
    const records = STATE.filteredDatasets[estamentoKey] || [];
    const N = records.length;
    const cfg = overrideConfig || STATE.samplingConfig;

    // Determinar si en Administrativos aplica Censo por P9 (a menos que overrideConfig indique otra cosa)
    let activeDesign = cfg.design;
    if (!overrideConfig && estamentoKey === 'administrativos' && STATE.params.P9 === 'CENSO') {
      activeDesign = 'CENSO';
    }

    const Z = Z_TABLE[cfg.confidence] || 1.959963984540054;
    const e = (Number(cfg.marginError) || 5.0) / 100.0;
    const p = Number(cfg.propP) || 0.50;
    const q = 1.0 - p;
    const tnr = (Number(cfg.tnr) || 0) / 100.0;

    if (N === 0) {
      return { N: 0, nNet: 0, nFinal: 0, H: 0, strata: [], design: activeDesign, Z, e, p, tnr };
    }

    // Variables de estratificación activas
    let selectedVars = STATE.selectedVarsByEstamento[estamentoKey] || [];
    if (selectedVars.length === 0) {
      const meta = ESTAMENTO_META[estamentoKey];
      selectedVars = (meta && meta.defaultVars) ? meta.defaultVars.slice() : [];
    }

    // Precalcular cortes para cualquier variable cuantitativa seleccionada
    const quantCutsMap = {};
    selectedVars.forEach((vId) => {
      if (isVariableQuantitative(estamentoKey, vId)) {
        const vals = records.map((r) => r[vId]);
        const customCuts = STATE.quantConfig.customCutsByVar[vId] || [3.0, 4.0];
        const cuts = computeQuantitativeCuts(
          vals,
          STATE.quantConfig.method,
          STATE.quantConfig.numBins,
          customCuts,
          STATE.quantConfig.separateZero
        );
        quantCutsMap[vId] = cuts;
        STATE.quantConfig.computedCutsByVar[vId] = cuts;
      }
    });

    // Agrupar registros en estratos
    const strataMap = new Map();
    if (activeDesign === 'MAS' || selectedVars.length === 0) {
      strataMap.set('POBLACION_GENERAL', {
        name: 'Población General (Sin Estratificación — MAS)',
        records: records.slice()
      });
    } else {
      for (let i = 0; i < records.length; i++) {
        const rec = records[i];
        const parts = selectedVars.map((vId) => {
          if (isVariableQuantitative(estamentoKey, vId)) {
            return formatQuantInterval(rec[vId], quantCutsMap[vId], STATE.quantConfig.separateZero, vId);
          }
          const rawVal = rec[vId];
          return (rawVal !== undefined && rawVal !== null && String(rawVal).trim() !== '')
            ? String(rawVal)
            : 'Sin Información';
        });
        const key = parts.join(' · ');
        if (!strataMap.has(key)) {
          strataMap.set(key, { name: key, records: [] });
        }
        strataMap.get(key).records.push(rec);
      }
    }

    // Construir arreglo ordenado de estratos
    const rawStrata = Array.from(strataMap.values()).sort((a, b) => b.records.length - a.records.length);
    const H = rawStrata.length;

    // Calcular estadísticos base por estrato
    const strata = rawStrata.map((st, idx) => {
      const Nh = st.records.length;
      const Wh = Nh / N;
      // Si se usa Neyman, permitimos una leve diferenciación realista o cuasi-varianza finita Nh/(Nh-1)*p*q
      const ph = p;
      const Sh = Nh > 1 ? Math.sqrt((Nh / (Nh - 1)) * ph * (1 - ph)) : Math.sqrt(ph * (1 - ph));
      return {
        index: idx + 1,
        code: `E-${String(idx + 1).padStart(2, '0')}`,
        name: st.name,
        Nh,
        Wh,
        ph,
        Sh,
        records: st.records,
        n0h: 0,
        nh: 0,
        fh: 0,
        expansionFactor: 1
      };
    });

    // Cálculo global de n_0 según el diseño
    let nNetGlobal = 0;

    if (activeDesign === 'CENSO') {
      nNetGlobal = N;
      strata.forEach((st) => {
        st.n0h = st.Nh;
        st.nh = st.Nh;
        st.fh = 100.0;
        st.expansionFactor = 1.0;
      });
      return {
        N,
        nNet: N,
        nFinal: N,
        H,
        strata,
        design: activeDesign,
        Z,
        e,
        p,
        tnr,
        fpc: 0
      };
    }

    if (activeDesign === 'MAS') {
      const n0Inf = (Z * Z * p * q) / (e * e);
      nNetGlobal = Math.min(N, Math.ceil(n0Inf / (1 + (n0Inf - 1) / N)));
    } else if (activeDesign === 'NEYMAN') {
      // Cochran (1977) ec. 5.25 para Afijación Óptima de Neyman
      const sumWhSh = strata.reduce((acc, s) => acc + s.Wh * s.Sh, 0);
      const sumWhSh2 = strata.reduce((acc, s) => acc + s.Wh * s.Sh * s.Sh, 0);
      const num = sumWhSh * sumWhSh;
      const den = (e * e) / (Z * Z) + sumWhSh2 / N;
      nNetGlobal = Math.min(N, Math.ceil(num / den));
    } else {
      // PROPORCIONAL: Cochran (1977) con w_h = W_h
      const sumWhSh2 = strata.reduce((acc, s) => acc + s.Wh * s.Sh * s.Sh, 0);
      const num = sumWhSh2;
      const den = (e * e) / (Z * Z) + sumWhSh2 / N;
      nNetGlobal = Math.min(N, Math.ceil(num / den));
    }

    // Ajuste por Tasa de No Respuesta (TNR)
    const nAdjustedTarget = Math.min(N, Math.ceil(nNetGlobal / Math.max(0.1, 1.0 - tnr)));

    // Reparto por estrato según esquema de afijación
    const sumNhSh = strata.reduce((acc, s) => acc + s.Nh * s.Sh, 0);
    let totalFinalN = 0;
    let totalBaseN = 0;

    const isEstudiantes = estamentoKey === 'estudiantes';
    const minGpa = isEstudiantes ? Number(STATE.studentGpaFilter.minGpa) || 0 : 0;
    const zeroRule = isEstudiantes ? STATE.studentGpaFilter.zeroRule : 'STRICT_WITH_FALLBACK';

    strata.forEach((st) => {
      const weightAlloc =
        activeDesign === 'NEYMAN' && sumNhSh > 0
          ? (st.Nh * st.Sh) / sumNhSh
          : st.Wh;

      const baseH = Math.min(st.Nh, Math.max(1, Math.round(nNetGlobal * weightAlloc)));
      let finalH = Math.min(st.Nh, Math.max(1, Math.ceil(nAdjustedTarget * weightAlloc)));

      if (cfg.microStrataCensus && st.Nh <= 5) {
        finalH = st.Nh;
      }

      st.n0h = baseH;
      st.nh = finalH;

      // Evaluar subconjunto elegible por Promedio >= minGpa en Estudiantes
      if (isEstudiantes && minGpa > 0) {
        const recsGeC = st.records.filter((r) => Number(r.promedio) >= minGpa);
        const recsZero = st.records.filter((r) => Number(r.promedio) === 0);
        const recsLowPos = st.records.filter((r) => Number(r.promedio) > 0 && Number(r.promedio) < minGpa);

        st.NhGpaGeC = recsGeC.length;
        st.NhGpaZero = recsZero.length;
        st.NhGpaLowPos = recsLowPos.length;
        st.strictViable = recsGeC.length >= finalH;
        st.strictDeficit = Math.max(0, finalH - recsGeC.length);

        if (zeroRule === 'INCLUDE_ZERO') {
          st.eligibleRecords = [...recsGeC, ...recsZero];
          st.NhEligible = st.eligibleRecords.length;
        } else if (zeroRule === 'STRICT_PURE') {
          st.eligibleRecords = recsGeC;
          st.NhEligible = recsGeC.length;
        } else {
          // STRICT_WITH_FALLBACK: Prioriza Promedio >= c; si hay déficit en el estrato, completa con Promedio = 0.0
          st.recsGeC = recsGeC;
          st.recsZero = recsZero;
          st.eligibleRecords = recsGeC.length >= finalH ? recsGeC : [...recsGeC, ...recsZero];
          st.NhEligible = recsGeC.length;
        }
      } else {
        st.NhGpaGeC = st.Nh;
        st.NhGpaZero = 0;
        st.NhGpaLowPos = 0;
        st.strictViable = true;
        st.strictDeficit = 0;
        st.eligibleRecords = st.records;
        st.NhEligible = st.Nh;
      }

      st.fh = (finalH / st.Nh) * 100.0;
      st.expansionFactor = finalH > 0 ? st.Nh / finalH : 1.0;

      totalBaseN += baseH;
      totalFinalN += finalH;
    });

    const fpc = N > 1 ? Math.sqrt(Math.max(0, (N - totalFinalN) / (N - 1))) : 0;

    return {
      N,
      nNet: totalBaseN,
      nFinal: totalFinalN,
      H,
      strata,
      design: activeDesign,
      Z,
      e,
      p,
      tnr,
      fpc,
      minGpa,
      zeroRule
    };
  }

  // --------------------------------------------------------------------------
  // 9. SORTEO ALEATORIO REPRODUCIBLE DE INDIVIDUOS POR ESTRATO
  // --------------------------------------------------------------------------
  function executeRandomSampleDraw() {
    if (!STATE.currentResult) return;
    const seed = Number(document.getElementById('inputRandomSeed').value) || 20261111;
    STATE.samplingConfig.seed = seed;
    const prng = createPRNG(seed);

    const selectedAll = [];
    let globalOrder = 1;
    const isEst = STATE.activeEstamento === 'estudiantes';
    const minGpa = isEst ? Number(STATE.studentGpaFilter.minGpa) || 0 : 0;
    const zeroRule = isEst ? STATE.studentGpaFilter.zeroRule : 'STRICT_WITH_FALLBACK';

    STATE.currentResult.strata.forEach((st) => {
      let picked = [];
      if (isEst && minGpa > 0) {
        if (zeroRule === 'STRICT_WITH_FALLBACK') {
          const primary = sampleWithoutReplacement(st.recsGeC || [], st.nh, prng);
          if (primary.length < st.nh && st.recsZero && st.recsZero.length > 0) {
            const needed = st.nh - primary.length;
            const fallback = sampleWithoutReplacement(st.recsZero, needed, prng);
            picked = [...primary, ...fallback];
          } else {
            picked = primary;
          }
        } else {
          picked = sampleWithoutReplacement(st.eligibleRecords || st.records, st.nh, prng);
        }
      } else {
        picked = sampleWithoutReplacement(st.records, st.nh, prng);
      }

      const pi_h = st.Nh > 0 ? picked.length / st.Nh : 1;
      const weight_h = pi_h > 0 ? 1 / pi_h : 1;

      picked.forEach((rec) => {
        selectedAll.push({
          order: globalOrder++,
          id_registro: rec.id_registro || rec.id || `UNID-${globalOrder}`,
          hash_id: rec.hash_id || rec.hash || 'SHA256-ANON',
          stratumCode: st.code,
          stratumName: st.name,
          inclusionProb: pi_h,
          designWeight: weight_h,
          rawRecord: rec
        });
      });
    });

    STATE.drawnSample = selectedAll;
    renderSelectedSampleTable();
  }

  // --------------------------------------------------------------------------
  // 10. MOTOR DE MUESTREO BIFÁSICO (n_inicial vs n_final - n_inicial)
  // --------------------------------------------------------------------------
  function computeAndRenderBifasico(reDrawSamples = false) {
    const baseRes = STATE.currentResult;
    if (!baseRes || baseRes.N === 0) return;

    const N = baseRes.N;
    const Z = baseRes.Z;
    const e = baseRes.e;
    const pilotTotalInput = Math.max(30, Math.min(N, Number(document.getElementById('inputPilotTotalN').value) || 60));
    const allocMode = document.getElementById('selectBifasicoAllocation').value || 'NEYMAN';

    STATE.bifasico.pilotTotalN = pilotTotalInput;
    STATE.bifasico.allocation = allocMode;

    // 1. Asignar n_inicial,h proporcionalmente (mínimo 2 por estrato si Nh >= 2)
    const bifStrata = baseRes.strata.map((st) => {
      const nIniH = Math.min(st.Nh, Math.max(Math.min(2, st.Nh), Math.round(pilotTotalInput * st.Wh)));
      let p0h = STATE.bifasico.pilotPropsByStratum[st.code];
      if (p0h === undefined || Number.isNaN(Number(p0h))) {
        p0h = 0.50;
        STATE.bifasico.pilotPropsByStratum[st.code] = p0h;
      }
      // Cuasi-desviación estándar empírica del piloto
      const factorPilot = nIniH > 1 ? nIniH / (nIniH - 1) : 1.0;
      const sHatH = Math.sqrt(Math.max(0.0001, factorPilot * p0h * (1.0 - p0h)));
      return {
        ...st,
        nIniH,
        p0h,
        sHatH,
        nFinalH: 0,
        nRestanteH: 0
      };
    });

    // 2. Calcular n_final global usando las cuasi-varianzas empíricas del piloto
    const sumWhSh = bifStrata.reduce((acc, s) => acc + s.Wh * s.sHatH, 0);
    const sumWhSh2 = bifStrata.reduce((acc, s) => acc + s.Wh * s.sHatH * s.sHatH, 0);

    let nFinalGlobal = 0;
    if (allocMode === 'NEYMAN') {
      const num = sumWhSh * sumWhSh;
      const den = (e * e) / (Z * Z) + sumWhSh2 / N;
      nFinalGlobal = Math.min(N, Math.ceil(num / den));
    } else {
      const num = sumWhSh2;
      const den = (e * e) / (Z * Z) + sumWhSh2 / N;
      nFinalGlobal = Math.min(N, Math.ceil(num / den));
    }

    const sumNhSh = bifStrata.reduce((acc, s) => acc + s.Nh * s.sHatH, 0);

    let sumIni = 0;
    let sumFinal = 0;
    let sumRestante = 0;

    bifStrata.forEach((st) => {
      const wAlloc =
        allocMode === 'NEYMAN' && sumNhSh > 0
          ? (st.Nh * st.sHatH) / sumNhSh
          : st.Wh;
      const nFinH = Math.min(st.Nh, Math.max(1, Math.ceil(nFinalGlobal * wAlloc)));
      const nRestH = Math.max(0, nFinH - st.nIniH);

      st.nFinalH = nFinH;
      st.nRestanteH = nRestH;

      sumIni += st.nIniH;
      sumFinal += Math.max(st.nIniH, nFinH);
      sumRestante += nRestH;
    });

    STATE.bifasico.strataResults = bifStrata;

    // 3. Sortear Muestra Preliminar (S_inicial) y Complemento (S_restante en U_h \ S_inicial)
    if (reDrawSamples || STATE.bifasico.drawnPilot.length === 0) {
      const seed = Number(document.getElementById('inputRandomSeed').value) || 20261111;
      const prng = createPRNG(seed + 777);
      const pilotList = [];
      const remainingList = [];

      bifStrata.forEach((st) => {
        const shuffled = sampleWithoutReplacement(st.records, st.nIniH + st.nRestanteH, prng);
        const pilotUnits = shuffled.slice(0, st.nIniH);
        const remUnits = shuffled.slice(st.nIniH, st.nIniH + st.nRestanteH);

        pilotUnits.forEach((rec, i) => {
          pilotList.push({
            fase: 'FASE_1_PILOTO_INICIAL',
            orden_estrato: i + 1,
            id_registro: rec.id_registro || rec.id,
            hash_id: rec.hash_id || rec.hash,
            codigo_estrato: st.code,
            nombre_estrato: st.name,
            p_piloto_estrato: st.p0h
          });
        });

        remUnits.forEach((rec, i) => {
          remainingList.push({
            fase: 'FASE_2_COMPLEMENTO_RESTANTE',
            orden_estrato: st.nIniH + i + 1,
            id_registro: rec.id_registro || rec.id,
            hash_id: rec.hash_id || rec.hash,
            codigo_estrato: st.code,
            nombre_estrato: st.name,
            p_piloto_estrato: st.p0h
          });
        });
      });

      STATE.bifasico.drawnPilot = pilotList;
      STATE.bifasico.drawnRemaining = remainingList;
    }

    // Actualizar KPIs Bifásicos
    document.getElementById('kpiBifasicoInicial').textContent = formatNumber(sumIni);
    document.getElementById('kpiBifasicoFinal').textContent = formatNumber(sumFinal);
    document.getElementById('kpiBifasicoRestante').textContent = formatNumber(sumRestante);

    // Renderizar Tabla Bifásica
    const tbody = document.getElementById('tbodyBifasicoStrata');
    const tfoot = document.getElementById('tfootBifasicoStrata');
    if (!tbody || !tfoot) return;

    tbody.innerHTML = bifStrata
      .map(
        (st) => `
      <tr>
        <td class="num-col">${st.index}</td>
        <td><strong>${st.code}</strong> · ${st.name}</td>
        <td class="num-col">${formatNumber(st.Nh)}</td>
        <td class="num-col">${(st.Wh * 100).toFixed(2)}%</td>
        <td class="num-col"><strong>${formatNumber(st.nIniH)}</strong></td>
        <td class="num-col">
          <input type="number" class="form-number-sm bifasico-p-input" data-code="${st.code}"
                 min="0.02" max="0.98" step="0.02" value="${Number(st.p0h).toFixed(2)}" />
        </td>
        <td class="num-col">${st.sHatH.toFixed(4)}</td>
        <td class="num-col"><strong class="highlight-teal">${formatNumber(st.nFinalH)}</strong></td>
        <td class="num-col"><strong class="highlight-emerald">+${formatNumber(st.nRestanteH)}</strong></td>
        <td class="num-col">
          ${
            st.nRestanteH === 0
              ? '<span class="badge-pill badge-emerald">Cubierto con Piloto</span>'
              : `<span class="badge-pill badge-navy">Sortear ${formatNumber(st.nRestanteH)} más</span>`
          }
        </td>
      </tr>
    `
      )
      .join('');

    tfoot.innerHTML = `
      <tr>
        <td colspan="2">TOTALES CONSOLIDADOS (${bifStrata.length} Estratos)</td>
        <td class="num-col">${formatNumber(N)}</td>
        <td class="num-col">100.00%</td>
        <td class="num-col">${formatNumber(sumIni)}</td>
        <td class="num-col">--</td>
        <td class="num-col">--</td>
        <td class="num-col">${formatNumber(sumFinal)}</td>
        <td class="num-col">+${formatNumber(sumRestante)}</td>
        <td class="num-col">Total Efectivo: ${formatNumber(sumIni + sumRestante)}</td>
      </tr>
    `;

    // Eventos de edición en vivo de p_0h
    tbody.querySelectorAll('.bifasico-p-input').forEach((inp) => {
      inp.addEventListener('change', (ev) => {
        const code = ev.target.getAttribute('data-code');
        const val = Math.min(0.98, Math.max(0.02, Number(ev.target.value) || 0.50));
        STATE.bifasico.pilotPropsByStratum[code] = val;
        computeAndRenderBifasico(true);
      });
    });
  }

  // --------------------------------------------------------------------------
  // 11. ACTUALIZACIÓN DE INTERFAZ, KPIs, FÓRMULAS KaTeX Y GRÁFICOS PLOTLY
  // --------------------------------------------------------------------------
  function updateModalAndRibbonCounts() {
    const keys = ['estudiantes', 'docentes', 'administrativos', 'empleadores', 'graduados'];
    keys.forEach((k) => {
      const nPop = (STATE.filteredDatasets[k] || []).length;
      const modalEl = document.getElementById(`modalN_${k}`);
      const ribbonEl = document.getElementById(`ribbonN_${k}`);
      if (modalEl) modalEl.textContent = formatNumber(nPop);
      if (ribbonEl) ribbonEl.textContent = formatNumber(nPop);

      // Calcular muestra rápida para mostrar en la cinta superior
      const quickRes = computeSamplingForEstamento(k);
      const sampleEl = document.getElementById(`ribbonSample_${k}`);
      if (sampleEl) sampleEl.textContent = formatNumber(quickRes.nFinal);
    });

    // Etiquetas de estado en tarjetas superiores
    const nEst = (STATE.filteredDatasets.estudiantes || []).length;
    const minGpa = Number(STATE.studentGpaFilter.minGpa) || 0;
    document.getElementById('tagParamEstudiantes').textContent =
      `${STATE.params.P12 === 'FORMA_1' ? 'Forma 1' : 'Forma 2'}${minGpa > 0 ? ` · Prom ≥ ${minGpa.toFixed(1)}` : ''}`;
    const footEst = document.getElementById('footEstudiantes');
    if (footEst) {
      footEst.textContent =
        STATE.params.P12 === 'FORMA_1'
          ? `Estudiantes.xlsx (N = 29,904) · ${minGpa > 0 ? `Sorteo Prom ≥ ${minGpa.toFixed(1)}` : 'Sin filtro Prom.'}`
          : `Estudiantes_Enriquecido.xlsx (N = ${formatNumber(nEst)})`;
    }

    document.getElementById('tagParamDocentes').textContent =
      `${STATE.params.P_COMISION === 'EXCLUIR' ? 'Sin Comisión' : 'Con Comisión'} · ${STATE.params.P4 === 'EXCLUIR' ? 'Con Correo' : 'Todos'}`;
    document.getElementById('tagParamAdministrativos').textContent =
      `P9: ${STATE.params.P9 === 'MUESTREO' ? 'MAE Jerárquico' : 'Censo'}`;
    document.getElementById('tagParamEmpleadores').textContent =
      `P10: ${STATE.params.P10 === 'EMPRESA' ? 'Empresa (303)' : 'Jefe (531)'}`;
    document.getElementById('tagParamGraduados').textContent =
      `P11: ${STATE.params.P11 === 'EXCLUIR' ? 'Con Correo' : 'Todos'}`;

    const estImpact = document.getElementById('modalGroupImpactEstudiantes');
    if (estImpact) {
      estImpact.textContent = `Marco resultante: ${formatNumber(nEst)} estudiantes (${STATE.params.P12 === 'FORMA_1' ? 'Forma 1 · Estudiantes.xlsx' : 'Forma 2 · Estudiantes_Enriquecido.xlsx'})`;
    }
    const docImpact = document.getElementById('modalGroupImpactDocentes');
    if (docImpact) {
      docImpact.textContent = `Marco resultante: ${formatNumber(STATE.filteredDatasets.docentes.length)} docentes`;
    }
    const admImpact = document.getElementById('modalGroupImpactAdmin');
    if (admImpact) {
      admImpact.textContent = `Marco resultante: ${formatNumber(STATE.filteredDatasets.administrativos.length)} administrativos (${STATE.params.P9 === 'MUESTREO' ? 'MAE por Nivel Jerárquico' : 'Censo'})`;
    }
    const empImpact = document.getElementById('modalGroupImpactEmpleadores');
    if (empImpact) {
      empImpact.textContent = `Unidad activa: ${STATE.params.P10 === 'EMPRESA' ? 'Empresa única por NIT (N = 303)' : 'Jefe directo (N = 531)'}`;
    }
    const gradImpact = document.getElementById('modalGroupImpactGraduados');
    if (gradImpact) {
      gradImpact.textContent = `Marco resultante: ${formatNumber(STATE.filteredDatasets.graduados.length)} graduados`;
    }
  }

  function renderVariablesChecklist() {
    const estKey = STATE.activeEstamento;
    const meta = ESTAMENTO_META[estKey];
    const container = document.getElementById('stratVariablesChecklist');
    const subtitleEl = document.getElementById('stratVarsSubtitle');
    if (!container || !meta) return;

    // Mostrar u ocultar el panel especial de Forma 1 vs Forma 2 en Estudiantes
    const studentFormaPanel = document.getElementById('studentFormaPanel');
    const forma2FilterBox = document.getElementById('forma2FilterBox');
    if (estKey === 'estudiantes') {
      studentFormaPanel.classList.remove('hidden');
      if (STATE.params.P12 === 'FORMA_2') {
        forma2FilterBox.classList.remove('hidden');
      } else {
        forma2FilterBox.classList.add('hidden');
      }
    } else {
      studentFormaPanel.classList.add('hidden');
      forma2FilterBox.classList.add('hidden');
    }

    // Caso especial: Empleadores no tiene variables de estratificación en el dataset
    if (estKey === 'empleadores' || meta.variables.length === 0) {
      if (subtitleEl) subtitleEl.textContent = 'Sin variables de estratificación en el dataset';
      container.innerHTML = `
        <div style="padding: 12px 14px; background: #f8fafc; border: 1px dashed #94a3b8; border-radius: 8px; font-size: 12.5px; color: #334155; line-height: 1.45;">
          <strong>ℹ️ Sin variables de estratificación:</strong> El archivo oficial <code>Empleadores.xlsx</code> solo contiene datos de identificación y contacto (<em>Empresa, NIT, Jefe Directo, Teléfono y Correo</em>). El muestreo se ejecuta mediante <strong>Muestreo Aleatorio Simple Sin Reemplazo (MAS)</strong> o <strong>Censo</strong> sobre la unidad elegida en <strong>P10</strong> (${STATE.params.P10 === 'EMPRESA' ? '303 Empresas' : '531 Jefes'}).
        </div>
      `;
      updateQuantitativePanelVisibility();
      return;
    }

    // Caso de variable única obligatoria (Docentes: Tipo de Vinculación | Administrativos: Nivel Jerárquico)
    const isSingleLockedVar = (estKey === 'docentes' || estKey === 'administrativos') && meta.variables.length === 1;
    if (isSingleLockedVar) {
      STATE.selectedVarsByEstamento[estKey] = [meta.variables[0].id];
      if (subtitleEl) subtitleEl.textContent = 'Variable única oficial de estratificación para este estamento';
    } else if (estKey === 'estudiantes' && STATE.params.P12 === 'FORMA_1') {
      const forma1Ids = ESTAMENTO_META.estudiantes.defaultVarsForma1;
      STATE.selectedVarsByEstamento.estudiantes = (STATE.selectedVarsByEstamento.estudiantes || []).filter((id) =>
        forma1Ids.includes(id)
      );
      if (subtitleEl) subtitleEl.textContent = 'Forma 1: Selecciona 1, 2 o las 3 variables (Facultad, Nivel Agrupado, Modalidad)';
    } else {
      if (subtitleEl) subtitleEl.textContent = 'Selecciona las variables para construir los estratos';
    }

    const selected = STATE.selectedVarsByEstamento[estKey] || [];

    const varsToShow = meta.variables.filter((v) => {
      if (estKey === 'estudiantes' && STATE.params.P12 === 'FORMA_1') {
        return v.forma === 1;
      }
      return true;
    });

    container.innerHTML = varsToShow
      .map((v) => {
        const isChecked = selected.includes(v.id);
        const lockAttr = isSingleLockedVar ? 'disabled' : '';
        return `
        <label class="var-check-item ${isChecked ? 'checked' : ''}">
          <div class="var-check-left">
            <input type="checkbox" class="chk-strat-var" value="${v.id}" ${isChecked ? 'checked' : ''} ${lockAttr} />
            <span>${v.label}</span>
          </div>
          <span class="var-type-tag ${v.type === 'quant' ? 'tag-quant' : 'tag-cat'}">
            ${v.type === 'quant' ? 'CUANTITATIVA' : 'CATEGÓRICA'}
          </span>
        </label>
      `;
      })
      .join('');

    container.querySelectorAll('.chk-strat-var').forEach((chk) => {
      chk.addEventListener('change', () => {
        const activeVars = Array.from(container.querySelectorAll('.chk-strat-var:checked')).map(
          (el) => el.value
        );
        STATE.selectedVarsByEstamento[estKey] = activeVars;
        renderVariablesChecklist();
        updateQuantitativePanelVisibility();
        runMainCalculationAndRender();
      });
    });

    updateQuantitativePanelVisibility();
  }

  function updateQuantitativePanelVisibility() {
    const estKey = STATE.activeEstamento;
    const selected = STATE.selectedVarsByEstamento[estKey] || [];
    const quantVars = selected.filter((vId) => isVariableQuantitative(estKey, vId));

    const panel = document.getElementById('quantDiscretizationPanel');
    const histCard = document.getElementById('quantHistogramCard');
    const selectQuantVar = document.getElementById('selectQuantVarActive');

    if (quantVars.length === 0) {
      panel.classList.add('hidden');
      histCard.classList.add('hidden');
      return;
    }

    panel.classList.remove('hidden');
    histCard.classList.remove('hidden');

    const meta = ESTAMENTO_META[estKey];
    selectQuantVar.innerHTML = quantVars
      .map((vId) => {
        const vMeta = meta.variables.find((m) => m.id === vId);
        return `<option value="${vId}">${vMeta ? vMeta.label : vId}</option>`;
      })
      .join('');

    if (!quantVars.includes(STATE.quantConfig.activeVar)) {
      STATE.quantConfig.activeVar = quantVars[0];
    }
    selectQuantVar.value = STATE.quantConfig.activeVar;

    const customCuts = STATE.quantConfig.customCutsByVar[STATE.quantConfig.activeVar] || [3.0, 4.0];
    document.getElementById('inputCustomCuts').value = customCuts.join(', ');
  }

  function renderContextBanner() {
    const estKey = STATE.activeEstamento;
    const meta = ESTAMENTO_META[estKey];
    const P = STATE.params;
    const pill = document.getElementById('contextEstamentoPill');
    const desc = document.getElementById('contextRulesDescription');

    pill.textContent = `${meta.icon} ${meta.label.toUpperCase()}`;

    if (estKey === 'estudiantes') {
      const nEst = (STATE.filteredDatasets.estudiantes || []).length;
      const minGpa = Number(STATE.studentGpaFilter.minGpa) || 0;
      const gpaNote = minGpa > 0
        ? ` · <strong>Sorteo post-cálculo:</strong> Elegibles con <code>Promedio ≥ ${minGpa.toFixed(1)}</code>`
        : ' · <strong>Sorteo post-cálculo:</strong> Sin filtro de promedio mínimo';
      desc.innerHTML =
        P.P12 === 'FORMA_1'
          ? `<strong>P12 = Forma 1 Activa (Prioritaria):</strong> Base oficial <code>Estudiantes.xlsx</code> (<strong>N = 29,904</strong> estudiantes) estratificada por <code>${(STATE.selectedVarsByEstamento.estudiantes || []).join(' × ')}</code>${gpaNote}.`
          : `<strong>P12 = Forma 2 Activa (Base Enriquecida):</strong> Archivo <code>Estudiantes_Enriquecido.xlsx</code> con <strong>N = ${formatNumber(nEst)} estudiantes enlazados</strong> (ΔN = ${formatNumber(nEst - 29904)} respecto a Forma 1)${gpaNote}.`;
    } else if (estKey === 'docentes') {
      desc.innerHTML = `<strong>Reglas en Docentes Activas:</strong> En Comisión Admin = <code>${P.P_COMISION === 'EXCLUIR' ? 'Excluidos (-11 directivos = -16 filas)' : 'Incluidos'}</code> · P1 (Duplicados cátedra extra = <code>${P.P1}</code>) · P4 (239 sin correo en Cátedra = <code>${P.P4}</code>) · Estratificación única: <code>Tipo de Vinculación</code> → <strong>N = ${formatNumber(STATE.filteredDatasets.docentes.length)} docentes</strong>.`;
    } else if (estKey === 'administrativos') {
      desc.innerHTML = `<strong>Reglas P5–P9 Activas:</strong> P5 (19 sin cédula = <code>${P.P5}</code>) · P6 (Vacante 71720089 = <code>${P.P6}</code>) · P7 (70382234 = <code>${P.P7}</code>) · P8 (Conductor 71611587 = <code>${P.P8}</code>) · P9 (Estrategia = <code>${P.P9 === 'MUESTREO' ? 'MAE por Nivel Jerárquico' : 'Censo'}</code>) → <strong>N = ${formatNumber(STATE.filteredDatasets.administrativos.length)} funcionarios</strong>.`;
    } else if (estKey === 'empleadores') {
      desc.innerHTML = `<strong>Regla P10 Activa:</strong> Unidad estadística de muestreo = <code>${P.P10 === 'EMPRESA' ? 'Empresa única por NIT (N = 303)' : 'Jefe / Contacto directo (N = 531)'}</code> · Sin variables de estratificación (Muestreo Aleatorio Simple Sin Reemplazo).`;
    } else if (estKey === 'graduados') {
      desc.innerHTML = `<strong>Regla P11 Activa:</strong> Graduados sin correo electrónico = <code>${P.P11 === 'EXCLUIR' ? 'Excluir (Marco contactable con correo: N = 3,276)' : 'No Excluir (Marco histórico total: N = 45,154)'}</code>.`;
    } else {
      desc.innerHTML = `<strong>Dataset Propio Cargado:</strong> ${formatNumber(STATE.filteredDatasets.custom.length)} registros activos en memoria local.`;
    }
  }

  function renderLiveMathFormula(res) {
    const container = document.getElementById('liveFormulaMathContainer');
    const titleEl = document.getElementById('liveFormulaTitle');
    if (!container || !res) return;

    const Zstr = res.Z.toFixed(4);
    const estr = res.e.toFixed(4);
    let latex = '';

    if (res.design === 'CENSO') {
      titleEl.textContent = 'Censo Institucional Exhaustivo (N = n)';
      latex = `$$n = N = ${res.N}, \\qquad f = \\frac{n}{N} = 100.0\\%, \\qquad \\text{Error Muestral } e = 0.0\\%$$`;
    } else if (res.design === 'MAS') {
      titleEl.textContent = 'Muestreo Aleatorio Simple Sin Reemplazo (MAS con FPC)';
      latex = `$$n_0 = \\frac{(${Zstr})^2 \\cdot (${res.p.toFixed(2)})(${(1 - res.p).toFixed(2)})}{(${estr})^2}, \\qquad n_{\\text{MAS}} = \\left\\lceil \\frac{n_0}{1 + \\dfrac{n_0 - 1}{${res.N}}} \\right\\rceil = \\mathbf{${res.nFinal}}$$`;
    } else if (res.design === 'NEYMAN') {
      titleEl.textContent = `Muestreo Aleatorio Estratificado — Afijación Óptima de Neyman (H = ${res.H} estratos)`;
      latex = `$$n_{\\text{Neyman}} = \\frac{\\left(\\sum_{h=1}^{${res.H}} W_h S_h\\right)^2}{\\dfrac{(${estr})^2}{(${Zstr})^2} + \\dfrac{1}{${res.N}}\\sum_{h=1}^{${res.H}} W_h S_h^2} = \\mathbf{${res.nNet}} \\xrightarrow{\\text{Ajuste TNR / Micro-estratos}} \\mathbf{n = ${res.nFinal}}, \\qquad n_h = n \\cdot \\frac{N_h S_h}{\\sum_{k=1}^{${res.H}} N_k S_k}$$`;
    } else {
      titleEl.textContent = `Muestreo Aleatorio Estratificado — Afijación Proporcional (H = ${res.H} estratos)`;
      latex = `$$n_{\\text{Prop}} = \\frac{\\sum_{h=1}^{${res.H}} W_h S_h^2}{\\dfrac{(${estr})^2}{(${Zstr})^2} + \\dfrac{1}{${res.N}}\\sum_{h=1}^{${res.H}} W_h S_h^2} = \\mathbf{${res.nNet}} \\xrightarrow{\\text{Ajuste TNR / Micro-estratos}} \\mathbf{n = ${res.nFinal}}, \\qquad n_h = \\left\\lceil n \\cdot \\frac{N_h}{${res.N}} \\right\\rceil$$`;
    }

    container.innerHTML = latex;
    if (window.renderMathInElement) {
      window.renderMathInElement(container, {
        delimiters: [
          { left: '$$', right: '$$', display: true },
          { left: '$', right: '$', display: false }
        ]
      });
    }
  }

  // Número formateado para usarse DENTRO de KaTeX ($...$): la coma va entre llaves para que no agregue espacio
  function fmtMath(n) {
    return formatNumber(n).replace(/,/g, '{,}');
  }

  function renderMathIn(el) {
    if (el && window.renderMathInElement) {
      window.renderMathInElement(el, {
        delimiters: [
          { left: '$$', right: '$$', display: true },
          { left: '$', right: '$', display: false }
        ],
        throwOnError: false
      });
    }
  }

  function renderStudentGpaViabilityStrip(res) {
    const strip = document.getElementById('gpaViabilitySummaryStrip');
    if (!strip || !res || STATE.activeEstamento !== 'estudiantes') return;

    const minGpa = Number(STATE.studentGpaFilter.minGpa) || 0;
    const zeroRule = STATE.studentGpaFilter.zeroRule;
    const cTxt = minGpa.toFixed(1);

    const totalGeC = res.strata.reduce((acc, s) => acc + (s.NhGpaGeC ?? s.Nh), 0);
    const totalZero = res.strata.reduce((acc, s) => acc + (s.NhGpaZero ?? 0), 0);
    const totalLowPos = res.strata.reduce((acc, s) => acc + (s.NhGpaLowPos ?? 0), 0);
    const pct = (x) => (res.N > 0 ? ((x / res.N) * 100).toFixed(2) : '0.00');

    // Conteo dinámico de estudiantes con Promedio = 0.0 en la explicación de la regla
    const zeroCountEl = document.getElementById('gpaZeroCountLabel');
    if (zeroCountEl) {
      zeroCountEl.textContent = minGpa > 0 ? `${formatNumber(totalZero)} estudiantes (${pct(totalZero)}%)` : '—';
    }

    if (minGpa <= 0) {
      strip.className = 'gpa-viability-strip viable-ok';
      strip.innerHTML = `
        <p class="gpa-viab-headline">Filtro desactivado ($c = 0.0$): los <strong>${formatNumber(res.N)}</strong> estudiantes participan en el sorteo de los <strong>${res.H}</strong> estratos ($n = ${fmtMath(res.nFinal)}$).</p>`;
      renderMathIn(strip);
      return;
    }

    // Resultado por estrato según la regla elegida para Promedio = 0.0
    let totalFaltantes = 0;
    let totalCompletadosConCero = 0;
    const rows = res.strata.map((s) => {
      const ge = s.NhGpaGeC ?? 0;
      const zero = s.NhGpaZero ?? 0;
      const low = s.NhGpaLowPos ?? 0;
      const alcanzaGe = ge >= s.nh;
      let usaCero = 0;
      let faltan = 0;
      let accion = '';

      if (alcanzaGe) {
        accion = `Se sortean los ${formatNumber(s.nh)} entre los ${formatNumber(ge)} con promedio ≥ ${cTxt}`;
      } else if (zeroRule === 'STRICT_PURE') {
        faltan = s.nh - ge;
        accion = `Solo hay ${formatNumber(ge)}: el estrato queda con <strong>${formatNumber(faltan)}</strong> encuestado(s) menos`;
      } else if (zeroRule === 'INCLUDE_ZERO') {
        const disp = ge + zero;
        faltan = Math.max(0, s.nh - disp);
        usaCero = Math.min(zero, s.nh - ge);
        accion = faltan > 0
          ? `Ni sumando los de promedio 0.0 alcanza: faltan <strong>${formatNumber(faltan)}</strong>`
          : `Se sortea entre los ${formatNumber(disp)} con promedio ≥ ${cTxt} o igual a 0.0`;
      } else {
        usaCero = Math.min(zero, s.nh - ge);
        faltan = Math.max(0, s.nh - ge - usaCero);
        accion = `Se toman ${formatNumber(ge)} con promedio ≥ ${cTxt} y se completan <strong>${formatNumber(usaCero)}</strong> con promedio 0.0`
          + (faltan > 0 ? `; aún faltan <strong>${formatNumber(faltan)}</strong>` : '');
      }
      if (!alcanzaGe) totalCompletadosConCero += usaCero;
      totalFaltantes += faltan;

      const completo = faltan === 0;
      return `
        <tr class="${alcanzaGe ? '' : completo ? 'row-gpa-fallback' : 'row-gpa-deficit'}">
          <td class="num-col">${s.index}</td>
          <td>${s.name}</td>
          <td class="num-col">${formatNumber(s.Nh)}</td>
          <td class="num-col"><strong>${formatNumber(s.nh)}</strong></td>
          <td class="num-col">${formatNumber(ge)}</td>
          <td class="num-col">${formatNumber(low)}</td>
          <td class="num-col">${formatNumber(zero)}</td>
          <td class="center-col">${alcanzaGe ? '<span class="gpa-yes">Sí</span>' : '<span class="gpa-no">No</span>'}</td>
          <td>${accion}</td>
        </tr>`;
    });

    const deficitCount = res.strata.filter((s) => (s.NhGpaGeC ?? 0) < s.nh).length;
    const viableCount = res.H - deficitCount;

    let headline;
    if (deficitCount === 0) {
      strip.className = 'gpa-viability-strip viable-ok';
      headline = `<strong>Los ${res.H} estratos tienen suficientes estudiantes con promedio ≥ ${cTxt}</strong> para sortear la muestra completa de ${formatNumber(res.nFinal)} estudiantes.`;
    } else {
      strip.className = totalFaltantes > 0 ? 'gpa-viability-strip viable-bad' : 'gpa-viability-strip viable-warn';
      const reglaTxt = totalFaltantes > 0
        ? `Con la opción elegida, la muestra queda incompleta: <strong>faltan ${formatNumber(totalFaltantes)} estudiante(s)</strong>.`
        : zeroRule === 'INCLUDE_ZERO'
          ? `Con la opción elegida (incluir siempre a los de promedio 0.0), esos estratos sí alcanzan, y en todos los estratos los de promedio 0.0 también pueden ser sorteados. La muestra queda completa.`
          : `Con la opción elegida se completan usando <strong>${formatNumber(totalCompletadosConCero)} estudiante(s) con promedio 0.0</strong>, así que la muestra queda completa.`;
      headline = `<strong>${viableCount} de ${res.H} estratos</strong> tienen suficientes estudiantes con promedio ≥ ${cTxt}. En <strong>${deficitCount}</strong> (filas resaltadas) no alcanzan. ${reglaTxt}`;
    }

    strip.innerHTML = `
      <p class="gpa-viab-headline">${headline}</p>
      <div class="gpa-viab-legend">
        <span><strong>Población total</strong> usada para calcular los $n_h$: ${formatNumber(res.N)}</span>
        <span>Promedio ≥ ${cTxt}: ${formatNumber(totalGeC)} (${pct(totalGeC)}%)</span>
        <span>Promedio entre 0 y ${cTxt}: ${formatNumber(totalLowPos)} (${pct(totalLowPos)}%)</span>
        <span>Promedio = 0.0: ${formatNumber(totalZero)} (${pct(totalZero)}%)</span>
      </div>
      <div class="gpa-viab-table-wrap">
        <table class="gpa-viab-table">
          <thead>
            <tr>
              <th class="num-col">#</th>
              <th>Estrato</th>
              <th class="num-col">Estudiantes en el estrato<br/><span class="th-sym">$N_h$</span></th>
              <th class="num-col">Muestra requerida<br/><span class="th-sym">$n_h$</span></th>
              <th class="num-col">Con promedio ≥ ${cTxt}<br/><span class="th-sym">$N_h(\\text{Prom.} \\ge ${cTxt})$</span></th>
              <th class="num-col">Con promedio entre 0 y ${cTxt}<br/><span class="th-sym">$N_h(0 < \\text{Prom.} < ${cTxt})$</span></th>
              <th class="num-col">Con promedio 0.0<br/><span class="th-sym">$N_h(\\text{Prom.} = 0)$</span></th>
              <th class="center-col">¿Alcanzan los de promedio ≥ ${cTxt}?<br/><span class="th-sym">$N_h(\\text{Prom.} \\ge ${cTxt}) \\ge n_h$</span></th>
              <th>Qué se hace en el sorteo</th>
            </tr>
          </thead>
          <tbody>${rows.join('')}</tbody>
          <tfoot>
            <tr>
              <td></td>
              <td><strong>Total</strong></td>
              <td class="num-col"><strong>${formatNumber(res.N)}</strong></td>
              <td class="num-col"><strong>${formatNumber(res.nFinal)}</strong></td>
              <td class="num-col"><strong>${formatNumber(totalGeC)}</strong></td>
              <td class="num-col"><strong>${formatNumber(totalLowPos)}</strong></td>
              <td class="num-col"><strong>${formatNumber(totalZero)}</strong></td>
              <td class="center-col"><strong>${viableCount} / ${res.H}</strong></td>
              <td>${totalFaltantes > 0 ? `Faltan ${formatNumber(totalFaltantes)}` : 'Muestra completa'}</td>
            </tr>
          </tfoot>
        </table>
      </div>`;
    renderMathIn(strip);
  }

  function renderStrataAllocationTable(res) {
    const tbody = document.getElementById('tbodyStrataAllocation');
    const tfoot = document.getElementById('tfootStrataAllocation');
    const thElig = document.getElementById('thEligibleGpaCol');
    const searchTerm = (document.getElementById('inputSearchStrata').value || '').toLowerCase().trim();

    const isEst = STATE.activeEstamento === 'estudiantes';
    const minGpa = isEst ? Number(STATE.studentGpaFilter.minGpa) || 0 : 0;

    if (thElig) {
      thElig.textContent = isEst && minGpa > 0
        ? `Elegibles Prom ≥ ${minGpa.toFixed(1)}`
        : 'Elegibles Sorteo';
    }

    const filteredStrata = res.strata.filter(
      (s) =>
        !searchTerm ||
        s.name.toLowerCase().includes(searchTerm) ||
        s.code.toLowerCase().includes(searchTerm)
    );

    tbody.innerHTML = filteredStrata
      .map((st) => {
        let eligCellHtml = `${formatNumber(st.Nh)}`;
        if (isEst && minGpa > 0) {
          if (st.strictViable) {
            eligCellHtml = `<span class="badge-pill badge-emerald" title="Suficientes con Promedio ≥ ${minGpa.toFixed(1)}: ${formatNumber(st.NhGpaGeC)} ≥ n_h (${formatNumber(st.nh)})">✅ ${formatNumber(st.NhGpaGeC)}</span>`;
          } else {
            eligCellHtml = `<span class="badge-pill badge-amber" title="Déficit estricto: ${formatNumber(st.NhGpaGeC)} con Prom ≥ ${minGpa.toFixed(1)} (hay ${formatNumber(st.NhGpaZero)} con Prom = 0.0)">⚠️ ${formatNumber(st.NhGpaGeC)} <small>(+${formatNumber(st.NhGpaZero)} en 0.0)</small></span>`;
          }
        }
        return `
      <tr>
        <td class="num-col">${st.index}</td>
        <td><span class="badge-pill badge-navy">${st.code}</span></td>
        <td><strong>${st.name}</strong></td>
        <td class="num-col">${formatNumber(st.Nh)}</td>
        <td class="num-col">${(st.Wh * 100).toFixed(2)}%</td>
        <td class="num-col">${st.ph.toFixed(2)}</td>
        <td class="num-col">${st.Sh.toFixed(4)}</td>
        <td class="num-col">${formatNumber(st.n0h)}</td>
        <td class="num-col"><strong class="highlight-teal">${formatNumber(st.nh)}</strong></td>
        <td class="num-col">${eligCellHtml}</td>
        <td class="num-col">${st.fh.toFixed(2)}%</td>
        <td class="num-col">${st.expansionFactor.toFixed(2)}</td>
      </tr>
    `;
      })
      .join('');

    const globalFh = res.N > 0 ? ((res.nFinal / res.N) * 100).toFixed(2) : '0.00';
    const globalExp = res.nFinal > 0 ? (res.N / res.nFinal).toFixed(2) : '1.00';
    const totalGeC = res.strata.reduce((acc, s) => acc + (s.NhGpaGeC ?? s.Nh), 0);

    tfoot.innerHTML = `
      <tr>
        <td colspan="3">TOTAL CONSOLIDADO DEL ESTAMENTO (${res.H} Estratos)</td>
        <td class="num-col">${formatNumber(res.N)}</td>
        <td class="num-col">100.00%</td>
        <td class="num-col">${res.p.toFixed(2)}</td>
        <td class="num-col">--</td>
        <td class="num-col">${formatNumber(res.nNet)}</td>
        <td class="num-col">${formatNumber(res.nFinal)}</td>
        <td class="num-col">${formatNumber(totalGeC)}</td>
        <td class="num-col">${globalFh}%</td>
        <td class="num-col">${globalExp}</td>
      </tr>
    `;
  }

  function renderSelectedSampleTable() {
    const tbody = document.getElementById('tbodySelectedIndividuals');
    const statusBadge = document.getElementById('sampleDrawnStatusBadge');
    const footerNote = document.getElementById('samplePreviewFooterNote');
    const query = (document.getElementById('inputSearchSampleRows').value || '').toLowerCase().trim();

    const list = STATE.drawnSample || [];
    statusBadge.textContent = `${formatNumber(list.length)} Individuos Sorteados (Semilla: ${STATE.samplingConfig.seed})`;

    const filtered = query
      ? list.filter(
          (item) =>
            String(item.id_registro).toLowerCase().includes(query) ||
            String(item.hash_id).toLowerCase().includes(query) ||
            String(item.stratumName).toLowerCase().includes(query)
        )
      : list;

    const preview = filtered.slice(0, 150);
    tbody.innerHTML = preview
      .map((item) => {
        return `
        <tr>
          <td class="num-col">#${item.order}</td>
          <td><code>${item.id_registro}</code></td>
          <td><code>${item.hash_id}</code></td>
          <td><span class="badge-pill badge-teal">${item.stratumCode}</span></td>
          <td>${item.stratumName}</td>
          <td class="num-col">${(item.inclusionProb * 100).toFixed(2)}%</td>
          <td class="num-col">${item.designWeight.toFixed(2)}</td>
        </tr>
      `;
      })
      .join('');

    footerNote.innerHTML = `Mostrando <strong>${formatNumber(preview.length)}</strong> de <strong>${formatNumber(filtered.length)}</strong> unidades seleccionadas aleatoriamente. Usa los botones <strong>Exportar Muestra (.xlsx / .csv)</strong> para obtener el archivo completo con las ${formatNumber(list.length)} unidades.`;
  }

  // --------------------------------------------------------------------------
  // 12. GRÁFICOS INTERACTIVOS CON PLOTLY.JS
  // --------------------------------------------------------------------------
  function renderCharts(res) {
    if (!window.Plotly || !res) return;

    // 1. Gráfico de Barras por Estrato
    const viewMode = document.getElementById('selectChartMetricView').value || 'SAMPLE_BARS';
    const topStrata = res.strata.slice(0, 25); // Mostrar hasta 25 estratos principales
    const labels = topStrata.map((s) => `${s.code}: ${s.name.length > 34 ? s.name.slice(0, 34) + '…' : s.name}`);

    let traces = [];
    if (viewMode === 'POP_VS_SAMPLE') {
      traces = [
        {
          x: labels,
          y: topStrata.map((s) => s.Nh),
          name: 'Población (N_h)',
          type: 'bar',
          marker: { color: '#003865' }
        },
        {
          x: labels,
          y: topStrata.map((s) => s.nh),
          name: 'Muestra (n_h)',
          type: 'bar',
          marker: { color: '#0D9488' }
        }
      ];
    } else if (viewMode === 'WEIGHTS') {
      traces = [
        {
          x: labels,
          y: topStrata.map((s) => Number((s.Wh * 100).toFixed(2))),
          name: 'Peso W_h (%)',
          type: 'bar',
          marker: {
            color: topStrata.map((_, i) => PALETTE[i % PALETTE.length])
          },
          text: topStrata.map((s) => `${(s.Wh * 100).toFixed(1)}%`),
          textposition: 'outside'
        }
      ];
    } else {
      traces = [
        {
          x: labels,
          y: topStrata.map((s) => s.nh),
          name: 'Muestra Asignada (n_h)',
          type: 'bar',
          marker: {
            color: topStrata.map((_, i) => PALETTE[i % PALETTE.length]),
            line: { width: 1, color: '#ffffff' }
          },
          text: topStrata.map((s) => s.nh),
          textposition: 'outside'
        }
      ];
    }

    const layoutStrata = {
      margin: { t: 20, r: 20, b: 110, l: 55 },
      paper_bgcolor: 'rgba(0,0,0,0)',
      plot_bgcolor: 'rgba(0,0,0,0)',
      font: { family: 'Inter, sans-serif', size: 11, color: '#1E293B' },
      xaxis: { tickangle: -32, automargin: true },
      yaxis: {
        title: viewMode === 'WEIGHTS' ? 'Peso W_h (%)' : 'Unidades',
        gridcolor: '#E2E8F0'
      },
      barmode: 'group',
      legend: { orientation: 'h', y: 1.12 }
    };

    Plotly.react('plotStrataAllocation', traces, layoutStrata, { responsive: true, displayModeBar: false });

    // 2. Curva de Sensibilidad Muestral (e = 1.0% a 10.0% para 90%, 95%, 99%)
    const errors = [];
    for (let eVal = 1.0; eVal <= 10.0; eVal += 0.5) {
      errors.push(Number(eVal.toFixed(1)));
    }

    const confLevels = [
      { conf: 90, color: '#059669', name: '90% Confianza' },
      { conf: 95, color: '#003865', name: '95% Confianza' },
      { conf: 99, color: '#E11D48', name: '99% Confianza' }
    ];

    const sensTraces = confLevels.map((cl) => {
      const zVal = Z_TABLE[cl.conf];
      const yVals = errors.map((ePct) => {
        const eDec = ePct / 100.0;
        const n0 = (zVal * zVal * res.p * (1 - res.p)) / (eDec * eDec);
        return Math.min(res.N, Math.ceil(n0 / (1 + (n0 - 1) / res.N)));
      });
      return {
        x: errors,
        y: yVals,
        mode: 'lines+markers',
        name: cl.name,
        line: { color: cl.color, width: 2.5 },
        marker: { size: 5 }
      };
    });

    // Punto activo actual
    sensTraces.push({
      x: [res.e * 100],
      y: [res.nFinal],
      mode: 'markers',
      name: 'Punto Activo',
      marker: { color: '#D97706', size: 13, symbol: 'diamond', line: { color: '#ffffff', width: 2 } }
    });

    const layoutSens = {
      margin: { t: 20, r: 20, b: 50, l: 60 },
      paper_bgcolor: 'rgba(0,0,0,0)',
      plot_bgcolor: 'rgba(0,0,0,0)',
      font: { family: 'Inter, sans-serif', size: 11, color: '#1E293B' },
      xaxis: { title: 'Margen de Error Admisible e (%)', gridcolor: '#E2E8F0' },
      yaxis: { title: 'Tamaño de Muestra (n)', gridcolor: '#E2E8F0' },
      legend: { orientation: 'h', y: 1.14 }
    };

    Plotly.react('plotSensitivityCurve', sensTraces, layoutSens, { responsive: true, displayModeBar: false });

    // 3. Histograma de Variable Cuantitativa y Puntos de Corte (si está visible)
    const histCard = document.getElementById('quantHistogramCard');
    if (!histCard.classList.contains('hidden')) {
      renderQuantitativeHistogram();
    }
  }

  function renderQuantitativeHistogram() {
    const estKey = STATE.activeEstamento;
    const varId = STATE.quantConfig.activeVar;
    const records = STATE.filteredDatasets[estKey] || [];
    if (!records.length || !varId) return;

    const sepZero = STATE.quantConfig.separateZero;
    const vals = records
      .map((r) => Number(r[varId]))
      .filter((v) => !Number.isNaN(v) && (sepZero ? v > 0 : true));

    const cuts = STATE.quantConfig.computedCutsByVar[varId] || [];
    const summaryBox = document.getElementById('quantCutsSummaryBox');
    summaryBox.innerHTML =
      cuts.length > 0
        ? `<strong>Puntos de corte activos (${STATE.quantConfig.method}):</strong> <code>${cuts.join('  |  ')}</code>`
        : `Sin cortes intermedios activos.`;

    document.getElementById('quantHistogramBadge').textContent =
      cuts.length > 0 ? `Cortes: ${cuts.join(', ')}` : 'Distribución Continua';

    const shapes = cuts.map((c) => ({
      type: 'line',
      x0: c,
      x1: c,
      y0: 0,
      y1: 1,
      yref: 'paper',
      line: { color: '#E11D48', width: 2.5, dash: 'dash' }
    }));

    const traceHist = {
      x: vals,
      type: 'histogram',
      nbinsx: 45,
      marker: {
        color: 'rgba(13, 148, 136, 0.72)',
        line: { color: '#003865', width: 1 }
      },
      name: varId
    };

    const layoutHist = {
      margin: { t: 18, r: 20, b: 45, l: 55 },
      paper_bgcolor: 'rgba(0,0,0,0)',
      plot_bgcolor: 'rgba(0,0,0,0)',
      font: { family: 'Inter, sans-serif', size: 11, color: '#1E293B' },
      xaxis: { title: `Valor de ${varId}`, gridcolor: '#E2E8F0' },
      yaxis: { title: 'Frecuencia Poblacional', gridcolor: '#E2E8F0' },
      shapes
    };

    Plotly.react('plotQuantHistogram', [traceHist], layoutHist, { responsive: true, displayModeBar: false });
  }

  // --------------------------------------------------------------------------
  // 13. MATRIZ GLOBAL CONSOLIDADA DE LOS 5 ESTAMENTOS (PESTAÑA 3)
  // --------------------------------------------------------------------------
  function renderGlobalConsolidadoMatrix() {
    const tbody = document.getElementById('tbodyGlobalConsolidado');
    const tfoot = document.getElementById('tfootGlobalConsolidado');
    if (!tbody || !tfoot) return;

    const estamentos = ['estudiantes', 'docentes', 'administrativos', 'empleadores', 'graduados'];
    const scenarios = [
      { id: 'A', confidence: 95, marginError: 5.0 },
      { id: 'B', confidence: 95, marginError: 3.5 },
      { id: 'C', confidence: 95, marginError: 3.0 },
      { id: 'D', confidence: 99, marginError: 3.0 }
    ];

    let sumRaw = 0;
    let sumN = 0;
    let sumH = 0;
    const sumScen = { A: 0, B: 0, C: 0, D: 0 };
    const chartData = [];

    const rowsHtml = estamentos
      .map((k) => {
        const meta = ESTAMENTO_META[k];
        const baseRes = computeSamplingForEstamento(k);
        const scenValues = {};

        scenarios.forEach((sc) => {
          if (k === 'administrativos' && STATE.params.P9 === 'CENSO') {
            scenValues[sc.id] = baseRes.N;
          } else {
            const r = computeSamplingForEstamento(k, {
              ...STATE.samplingConfig,
              design: 'PROPORCIONAL',
              confidence: sc.confidence,
              marginError: sc.marginError
            });
            scenValues[sc.id] = r.nFinal;
          }
          sumScen[sc.id] += scenValues[sc.id];
        });

        sumRaw += meta.rawCount;
        sumN += baseRes.N;
        sumH += baseRes.H;

        chartData.push({
          label: meta.label,
          A: scenValues.A,
          B: scenValues.B,
          C: scenValues.C,
          D: scenValues.D
        });

        let rulesText = '';
        if (k === 'estudiantes') rulesText = `P12: ${STATE.params.P12}${STATE.studentGpaFilter.minGpa > 0 ? ` (Prom≥${STATE.studentGpaFilter.minGpa.toFixed(1)})` : ''}`;
        if (k === 'docentes') rulesText = `Comisión:${STATE.params.P_COMISION}, P1:${STATE.params.P1}, P4:${STATE.params.P4}`;
        if (k === 'administrativos') rulesText = `P5-P7, P8:${STATE.params.P8}, P9:${STATE.params.P9}`;
        if (k === 'empleadores') rulesText = `P10: ${STATE.params.P10}`;
        if (k === 'graduados') rulesText = `P11: ${STATE.params.P11}`;

        const varsText = (STATE.selectedVarsByEstamento[k] || []).join(' × ');

        return `
        <tr>
          <td><strong>${meta.icon} ${meta.label}</strong></td>
          <td><code>${meta.sourceFile}</code></td>
          <td><span class="badge-pill badge-navy">${rulesText}</span></td>
          <td>${varsText}</td>
          <td class="num-col">${formatNumber(meta.rawCount)}</td>
          <td class="num-col"><strong>${formatNumber(baseRes.N)}</strong></td>
          <td class="num-col">${baseRes.H}</td>
          <td class="num-col"><strong class="highlight-teal">${formatNumber(scenValues.A)}</strong></td>
          <td class="num-col">${formatNumber(scenValues.B)}</td>
          <td class="num-col"><strong class="highlight-emerald">${formatNumber(scenValues.C)}</strong></td>
          <td class="num-col">${formatNumber(scenValues.D)}</td>
        </tr>
      `;
      })
      .join('');

    tbody.innerHTML = rowsHtml;
    tfoot.innerHTML = `
      <tr>
        <td colspan="4">TOTAL INSTITUCIONAL CONSOLIDADO (5 ESTAMENTOS ITM)</td>
        <td class="num-col">${formatNumber(sumRaw)}</td>
        <td class="num-col">${formatNumber(sumN)}</td>
        <td class="num-col">${sumH}</td>
        <td class="num-col">${formatNumber(sumScen.A)}</td>
        <td class="num-col">${formatNumber(sumScen.B)}</td>
        <td class="num-col">${formatNumber(sumScen.C)}</td>
        <td class="num-col">${formatNumber(sumScen.D)}</td>
      </tr>
    `;

    // Gráfico comparativo multiescenario
    if (window.Plotly) {
      const xLabels = chartData.map((d) => d.label);
      const traces = [
        { x: xLabels, y: chartData.map((d) => d.A), name: 'Escenario A (95% / 5.0%)', type: 'bar', marker: { color: '#0D9488' } },
        { x: xLabels, y: chartData.map((d) => d.B), name: 'Escenario B (95% / 3.5%)', type: 'bar', marker: { color: '#003865' } },
        { x: xLabels, y: chartData.map((d) => d.C), name: 'Escenario C (95% / 3.0%)', type: 'bar', marker: { color: '#059669' } },
        { x: xLabels, y: chartData.map((d) => d.D), name: 'Escenario D (99% / 3.0%)', type: 'bar', marker: { color: '#4F46E5' } }
      ];
      Plotly.react(
        'plotGlobalComparison',
        traces,
        {
          margin: { t: 25, r: 20, b: 50, l: 60 },
          paper_bgcolor: 'rgba(0,0,0,0)',
          plot_bgcolor: 'rgba(0,0,0,0)',
          barmode: 'group',
          yaxis: { title: 'Tamaño de Muestra (n)', gridcolor: '#E2E8F0' },
          legend: { orientation: 'h', y: 1.12 }
        },
        { responsive: true, displayModeBar: false }
      );
    }
  }

  // --------------------------------------------------------------------------
  // 14. EJECUCIÓN PRINCIPAL DE CÁLCULO Y RENDERIZADO
  // --------------------------------------------------------------------------
  function runMainCalculationAndRender() {
    const estKey = STATE.activeEstamento;
    renderContextBanner();

    const res = computeSamplingForEstamento(estKey);
    STATE.currentResult = res;

    // Actualizar banner de cambio inmediato en Estudiantes (Forma 1 vs Forma 2)
    if (estKey === 'estudiantes') {
      const isF1 = STATE.params.P12 === 'FORMA_1';
      const headingEl = document.getElementById('studentFormaMainHeading');
      const pillEl = document.getElementById('studentFormaStatusPill');
      const diffBanner = document.getElementById('formaLiveDiffBanner');
      const diffFile = document.getElementById('diffActiveFile');
      const diffN = document.getElementById('diffActiveN');
      const diffVars = document.getElementById('diffActiveVars');

      if (headingEl) {
        headingEl.innerHTML = isF1
          ? `Forma 1 Activa: <code>Estudiantes.xlsx</code> (<strong>N = ${formatNumber(res.N)}</strong> estudiantes · <strong>${res.H} estratos</strong>)`
          : `Forma 2 Activa: <code>Estudiantes_Enriquecido.xlsx</code> (<strong>N = ${formatNumber(res.N)}</strong> estudiantes con cruce · <strong>${res.H} estratos</strong>)`;
      }
      if (pillEl) {
        pillEl.innerHTML = isF1
          ? `Base Activa: <strong>Estudiantes.xlsx (N = 29,904)</strong>`
          : `Base Activa: <strong>Estudiantes_Enriquecido.xlsx (N = ${formatNumber(res.N)} | ΔN = ${formatNumber(res.N - 29904)})</strong>`;
      }
      if (diffBanner) {
        diffBanner.classList.toggle('mode-forma-1', isF1);
        diffBanner.classList.toggle('mode-forma-2', !isF1);
      }
      if (diffFile) {
        diffFile.textContent = isF1 ? 'Estudiantes.xlsx (Base Oficial)' : 'Estudiantes_Enriquecido.xlsx (Base Enriquecida)';
      }
      if (diffN) {
        diffN.textContent = isF1
          ? '29,904 estudiantes (100% del archivo original)'
          : `${formatNumber(res.N)} estudiantes (${res.N - 29904 === 0 ? 'Todos' : formatNumber(res.N - 29904) + ' sin cruce histórico excluidos'})`;
      }
      if (diffVars) {
        const activeVarNames = (STATE.selectedVarsByEstamento.estudiantes || []).join(' × ') || 'Sin estratificar (MAS)';
        diffVars.textContent = `${activeVarNames} (H = ${res.H} estratos | Muestra n = ${formatNumber(res.nFinal)})`;
      }
      renderStudentGpaViabilityStrip(res);
    }

    // Actualizar KPIs principales con animación de destello cuando cambia N o n
    const kpiPopEl = document.getElementById('kpiPopulationN');
    const prevNText = kpiPopEl ? kpiPopEl.textContent : '';
    const newNText = formatNumber(res.N);
    if (kpiPopEl) kpiPopEl.textContent = newNText;
    document.getElementById('kpiPopulationSub').textContent =
      estKey === 'estudiantes'
        ? `${STATE.params.P12 === 'FORMA_1' ? 'Archivo: Estudiantes.xlsx' : 'Archivo: Estudiantes_Enriquecido.xlsx'}`
        : `Estamento: ${ESTAMENTO_META[estKey].label}`;

    document.getElementById('kpiSampleNet').textContent = formatNumber(res.nNet);
    const fNet = res.N > 0 ? ((res.nNet / res.N) * 100).toFixed(2) : '0.00';
    document.getElementById('kpiSampleNetSub').textContent = `Fracción muestral base f = ${fNet}%`;

    document.getElementById('kpiSampleFinal').textContent = formatNumber(res.nFinal);
    const fFinal = res.N > 0 ? ((res.nFinal / res.N) * 100).toFixed(2) : '0.00';
    document.getElementById('kpiSampleFinalSub').textContent = `Fracción final f = ${fFinal}% (TNR = ${(res.tnr * 100).toFixed(0)}%)`;

    document.getElementById('kpiStrataCount').textContent = formatNumber(res.H);
    document.getElementById('kpiStrataSub').textContent = `Factor FPC global: ${res.fpc.toFixed(4)}`;

    if (prevNText && prevNText !== newNText) {
      document.querySelectorAll('.kpi-card').forEach((card) => {
        card.classList.remove('kpi-flash-update');
        void card.offsetWidth;
        card.classList.add('kpi-flash-update');
      });
    }

    // Actualizar muestra en la tarjeta superior activa
    const ribbonSampleEl = document.getElementById(`ribbonSample_${estKey}`);
    if (ribbonSampleEl) ribbonSampleEl.textContent = formatNumber(res.nFinal);

    renderLiveMathFormula(res);
    renderStrataAllocationTable(res);
    executeRandomSampleDraw();
    renderCharts(res);
    computeAndRenderBifasico(true);
    renderGlobalConsolidadoMatrix();
  }

  // --------------------------------------------------------------------------
  // 15. EXPORTACIÓN A EXCEL (.XLSX) Y CSV (.CSV) CON SHEETJS
  // --------------------------------------------------------------------------
  function exportCurrentSampleToExcel(format = 'xlsx') {
    if (!window.XLSX || !STATE.drawnSample.length) return;
    const estKey = STATE.activeEstamento;
    const rows = STATE.drawnSample.map((item) => ({
      Orden_Sorteo: item.order,
      Estamento: ESTAMENTO_META[estKey].label,
      ID_Anonimizado: item.id_registro,
      Hash_SHA256: item.hash_id,
      Codigo_Estrato: item.stratumCode,
      Definicion_Estrato: item.stratumName,
      Probabilidad_Inclusion_Pi: Number(item.inclusionProb.toFixed(6)),
      Factor_Expansion_Peso: Number(item.designWeight.toFixed(4)),
      Semilla_PRNG: STATE.samplingConfig.seed,
      Confianza_Pct: STATE.samplingConfig.confidence,
      Error_Margen_Pct: STATE.samplingConfig.marginError
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Muestra_Seleccionada');

    const filename = `SIMER_ITM_Muestra_${estKey}_n${rows.length}_Seed${STATE.samplingConfig.seed}.${format}`;
    XLSX.writeFile(wb, filename, { bookType: format === 'csv' ? 'csv' : 'xlsx' });
  }

  function exportStrataTableToExcel() {
    if (!window.XLSX || !STATE.currentResult) return;
    const estKey = STATE.activeEstamento;
    const res = STATE.currentResult;
    const rows = res.strata.map((st) => ({
      No_Estrato: st.index,
      Codigo_Estrato: st.code,
      Definicion_Estrato: st.name,
      Poblacion_Nh: st.Nh,
      Peso_Wh_Pct: Number((st.Wh * 100).toFixed(4)),
      Proporcion_ph: st.ph,
      Cuasi_Desv_Sh: Number(st.Sh.toFixed(4)),
      Muestra_Base_n0h: st.n0h,
      Muestra_Final_nh: st.nh,
      Fraccion_Muestral_fh_Pct: Number(st.fh.toFixed(2)),
      Factor_Expansion_Fh: Number(st.expansionFactor.toFixed(4))
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Afijacion_Estratos');
    XLSX.writeFile(wb, `SIMER_ITM_Afijacion_Estratos_${estKey}.xlsx`);
  }

  function exportBifasicoWorkbook() {
    if (!window.XLSX) return;
    const estKey = STATE.activeEstamento;
    const wb = XLSX.utils.book_new();

    const wsPilot = XLSX.utils.json_to_sheet(STATE.bifasico.drawnPilot);
    const wsRem = XLSX.utils.json_to_sheet(STATE.bifasico.drawnRemaining);
    const wsConsolidated = XLSX.utils.json_to_sheet([
      ...STATE.bifasico.drawnPilot,
      ...STATE.bifasico.drawnRemaining
    ]);
    const wsStrata = XLSX.utils.json_to_sheet(
      STATE.bifasico.strataResults.map((s) => ({
        Codigo: s.code,
        Estrato: s.name,
        Poblacion_Nh: s.Nh,
        Muestra_Preliminar_n_inicial: s.nIniH,
        Proporcion_Piloto_p0h: s.p0h,
        Cuasi_Desv_Piloto_Sh: Number(s.sHatH.toFixed(4)),
        Muestra_Definitiva_n_final: s.nFinalH,
        Individuos_Faltantes_n_restante: s.nRestanteH
      }))
    );

    XLSX.utils.book_append_sheet(wb, wsStrata, '1_Resumen_Estratos');
    XLSX.utils.book_append_sheet(wb, wsPilot, '2_Muestra_Preliminar');
    XLSX.utils.book_append_sheet(wb, wsRem, '3_Muestra_Restante_Comp');
    XLSX.utils.book_append_sheet(wb, wsConsolidated, '4_Muestra_Total_Final');

    XLSX.writeFile(wb, `SIMER_ITM_Muestreo_Bifasico_${estKey}.xlsx`);
  }

  function exportGlobalSummaryWorkbook() {
    if (!window.XLSX) return;
    const table = document.getElementById('tableGlobalConsolidado');
    const wb = XLSX.utils.table_to_book(table, { sheet: 'Consolidado_5_Estamentos' });
    XLSX.writeFile(wb, 'SIMER_ITM_Consolidado_Global_5_Estamentos.xlsx');
  }

  // --------------------------------------------------------------------------
  // 16. CARGA DE DATASET PROPIO (.XLSX / .CSV) VÍA SHEETJS
  // --------------------------------------------------------------------------
  function handleCustomFileUpload(file) {
    if (!file || !window.XLSX) return;
    const reader = new FileReader();
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      const jsonRows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: '' });

      if (!jsonRows.length) return;

      const cols = Object.keys(jsonRows[0]);
      const formattedRows = jsonRows.map((r, idx) => ({
        ...r,
        id_registro: r.ID || r.id || `CUST-${String(idx + 1).padStart(5, '0')}`,
        hash_id: `LOCAL-${String(idx + 1).padStart(5, '0')}`
      }));

      // Detectar tipos de columnas automáticamente
      const detectedVars = cols.map((colName) => {
        const sampleVals = formattedRows.slice(0, 100).map((r) => r[colName]).filter((v) => v !== '');
        const numericCount = sampleVals.filter((v) => !Number.isNaN(Number(v))).length;
        const uniqueSet = new Set(sampleVals);
        const isQuant = numericCount > sampleVals.length * 0.8 && uniqueSet.size > 10;
        return {
          id: colName,
          label: `${colName} ${isQuant ? '[Cuantitativa]' : ''}`,
          type: isQuant ? 'quant' : 'cat'
        };
      });

      ESTAMENTO_META.custom.rawCount = formattedRows.length;
      ESTAMENTO_META.custom.sourceFile = file.name;
      ESTAMENTO_META.custom.variables = detectedVars;
      ESTAMENTO_META.custom.defaultVars = detectedVars.slice(0, 2).map((v) => v.id);

      STATE.rawDatasets.custom = formattedRows;
      STATE.filteredDatasets.custom = formattedRows;
      STATE.selectedVarsByEstamento.custom = ESTAMENTO_META.custom.defaultVars.slice();

      const statusBox = document.getElementById('customDatasetStatusBox');
      statusBox.classList.remove('hidden');
      statusBox.innerHTML = `
        <h4>✅ Archivo cargado exitosamente: <code>${file.name}</code></h4>
        <p>Se detectaron <strong>${formatNumber(formattedRows.length)} registros</strong> y <strong>${cols.length} columnas</strong>. Ahora estás analizando este dataset en la Pestaña 1.</p>
      `;

      // Cambiar automáticamente al estamento custom y abrir Pestaña 1
      STATE.activeEstamento = 'custom';
      switchMainTab('tab-muestreo');
      renderVariablesChecklist();
      runMainCalculationAndRender();
    };
    reader.readAsArrayBuffer(file);
  }

  // --------------------------------------------------------------------------
  // 17. INICIALIZACIÓN DE EVENTOS E INTERACCIÓN DEL DOM
  // --------------------------------------------------------------------------
  function renderAllStaticMath() {
    if (!window.renderMathInElement) return;
    const opts = {
      delimiters: [
        { left: '$$', right: '$$', display: true },
        { left: '$', right: '$', display: false }
      ],
      throwOnError: false
    };
    const headerEl = document.querySelector('.itm-header');
    const mainEl = document.querySelector('.main-workspace');
    if (headerEl) window.renderMathInElement(headerEl, opts);
    if (mainEl) window.renderMathInElement(mainEl, opts);
  }

  function switchMainTab(tabId) {
    STATE.activeTab = tabId;
    document.querySelectorAll('.nav-tab').forEach((btn) => {
      btn.classList.toggle('active', btn.getAttribute('data-tab') === tabId);
    });
    document.querySelectorAll('.tab-panel').forEach((panel) => {
      panel.classList.toggle('active', panel.id === tabId);
    });
    renderAllStaticMath();
    // Redimensionar gráficos Plotly al cambiar de pestaña
    setTimeout(() => {
      if (window.Plotly) {
        document.querySelectorAll('.plotly-chart-box').forEach((box) => {
          try { Plotly.Plots.resize(box); } catch (_) {}
        });
      }
    }, 60);
  }

  function syncModalInputsFromState() {
    Object.keys(STATE.params).forEach((pKey) => {
      const val = STATE.params[pKey];
      const el = document.getElementById(`param_${pKey}`);
      if (el) el.value = val;

      // Sincronizar tarjetas interactivas .param-choice-btn dentro del modal
      document.querySelectorAll(`.param-choice-btn[data-param="${pKey}"]`).forEach((btn) => {
        const isSelected = btn.getAttribute('data-val') === val;
        btn.classList.toggle('active', isSelected);
        const radioSpan = btn.querySelector('.choice-radio');
        if (radioSpan) {
          radioSpan.textContent = isSelected ? '●' : '○';
        }
      });
    });

    // Sincronizar tarjetas Forma 1 vs Forma 2 en Estudiantes
    const isForma1 = STATE.params.P12 === 'FORMA_1';
    const c1 = document.getElementById('cardForma1');
    const c2 = document.getElementById('cardForma2');
    if (c1) c1.classList.toggle('active', isForma1);
    if (c2) c2.classList.toggle('active', !isForma1);
    const radios = document.querySelectorAll('input[name="radioP12Forma"]');
    radios.forEach((r) => {
      r.checked = r.value === STATE.params.P12;
    });

    const scopeSel = document.getElementById('selectForma2Scope');
    if (scopeSel) {
      scopeSel.value = STATE.forma2Scope;
    }
  }

  function applyParamOptionChange(pKey, newVal) {
    if (!pKey || !newVal) return;
    STATE.params[pKey] = newVal;

    if (pKey === 'P12') {
      if (STATE.params.P12 === 'FORMA_1') {
        STATE.selectedVarsByEstamento.estudiantes = ESTAMENTO_META.estudiantes.defaultVarsForma1.slice();
      } else {
        STATE.selectedVarsByEstamento.estudiantes = ESTAMENTO_META.estudiantes.defaultVarsForma2.slice();
      }
    }
    if (pKey === 'P9' && STATE.activeEstamento === 'administrativos') {
      STATE.samplingConfig.design = STATE.params.P9 === 'CENSO' ? 'CENSO' : 'PROPORCIONAL';
      const desSel = document.getElementById('selectSamplingDesign');
      if (desSel) desSel.value = STATE.samplingConfig.design;
    }

    syncModalInputsFromState();
    applyParameterizationRules();
    renderVariablesChecklist();
    runMainCalculationAndRender();

    const feedbackEl = document.getElementById('modalLiveFeedbackText');
    if (feedbackEl) {
      feedbackEl.textContent = `✓ Parámetro ${pKey} actualizado a "${newVal}" — Poblaciones recalculadas`;
    }
  }

  function resetParamsToDefaults() {
    STATE.params = {
      P_COMISION: 'EXCLUIR',
      P1: 'CATEDRA',
      P2: 'CORREO_ITM',
      P3: 'DEJAR_UNO',
      P4: 'EXCLUIR',
      P5: 'EXCLUIR',
      P6: 'EXCLUIR',
      P7: 'NO_EXCLUIR',
      P8: 'NO_EXCLUIR',
      P9: 'MUESTREO',
      P10: 'EMPRESA',
      P11: 'EXCLUIR',
      P12: 'FORMA_1'
    };
    STATE.forma2Scope = 'LINKED_ALL';
    STATE.studentGpaFilter.minGpa = 2.0;
    STATE.studentGpaFilter.zeroRule = 'STRICT_WITH_FALLBACK';
    const inpGpa = document.getElementById('inputMinGpaThreshold');
    const rngGpa = document.getElementById('rangeMinGpaThreshold');
    if (inpGpa) inpGpa.value = '2.0';
    if (rngGpa) rngGpa.value = '2.0';
    document.querySelectorAll('input[name="gpaZeroRule"]').forEach((r) => {
      r.checked = r.value === 'STRICT_WITH_FALLBACK';
    });
    document.querySelectorAll('.btn-preset-gpa').forEach((b) => {
      b.classList.toggle('active', Math.abs(Number(b.getAttribute('data-gpa')) - 2.0) < 0.05);
    });
    STATE.selectedVarsByEstamento.estudiantes = ESTAMENTO_META.estudiantes.defaultVarsForma1.slice();
    STATE.selectedVarsByEstamento.docentes = ['tipo_vinculacion'];
    STATE.selectedVarsByEstamento.administrativos = ['nivel_jerarquico'];
    STATE.selectedVarsByEstamento.empleadores = [];
    STATE.selectedVarsByEstamento.graduados = ['facultad', 'nivel_agrupado'];
    syncModalInputsFromState();
    applyParameterizationRules();
    renderVariablesChecklist();
    runMainCalculationAndRender();

    const feedbackEl = document.getElementById('modalLiveFeedbackText');
    if (feedbackEl) {
      feedbackEl.textContent = '↺ Restaurados los valores por defecto institucionales';
    }
  }

  function bindEventListeners() {
    // Pestañas de navegación
    document.querySelectorAll('.nav-tab').forEach((btn) => {
      btn.addEventListener('click', () => {
        switchMainTab(btn.getAttribute('data-tab'));
      });
    });

    // Selección de Estamento en las 5 tarjetas superiores
    document.querySelectorAll('.estamento-card').forEach((card) => {
      card.addEventListener('click', () => {
        const est = card.getAttribute('data-estamento');
        STATE.activeEstamento = est;
        document.querySelectorAll('.estamento-card').forEach((c) => c.classList.remove('active'));
        card.classList.add('active');

        // Ajustar diseño por defecto si es Administrativos y P9 = CENSO, o Empleadores (MAS)
        if (est === 'administrativos' && STATE.params.P9 === 'CENSO') {
          document.getElementById('selectSamplingDesign').value = 'CENSO';
          STATE.samplingConfig.design = 'CENSO';
        } else if (est === 'empleadores' && STATE.samplingConfig.design !== 'CENSO') {
          document.getElementById('selectSamplingDesign').value = 'MAS';
          STATE.samplingConfig.design = 'MAS';
        } else if (STATE.samplingConfig.design === 'CENSO' || STATE.samplingConfig.design === 'MAS') {
          document.getElementById('selectSamplingDesign').value = 'PROPORCIONAL';
          STATE.samplingConfig.design = 'PROPORCIONAL';
        }

        renderVariablesChecklist();
        runMainCalculationAndRender();
      });
    });

    // Selector Forma 1 vs Forma 2 en Estudiantes (Pantalla Principal)
    document.querySelectorAll('input[name="radioP12Forma"]').forEach((radio) => {
      radio.addEventListener('change', (e) => {
        applyParamOptionChange('P12', e.target.value);
      });
    });

    const scopeSel = document.getElementById('selectForma2Scope');
    if (scopeSel) {
      scopeSel.addEventListener('change', (e) => {
        STATE.forma2Scope = e.target.value;
        applyParameterizationRules();
        runMainCalculationAndRender();
      });
    }

    // Controles del Filtro de Promedio Mínimo en Estudiantes (cambiables en cualquier momento)
    const syncMinGpa = (val) => {
      const clean = Math.min(4.5, Math.max(0.0, Number(val) || 0.0));
      STATE.studentGpaFilter.minGpa = clean;
      const inp = document.getElementById('inputMinGpaThreshold');
      const rng = document.getElementById('rangeMinGpaThreshold');
      if (inp) inp.value = clean.toFixed(1);
      if (rng) rng.value = Math.min(4.0, clean);
      document.querySelectorAll('.btn-preset-gpa').forEach((b) => {
        b.classList.toggle('active', Math.abs(Number(b.getAttribute('data-gpa')) - clean) < 0.05);
      });
      updateModalAndRibbonCounts();
      runMainCalculationAndRender();
    };

    const inpGpa = document.getElementById('inputMinGpaThreshold');
    const rngGpa = document.getElementById('rangeMinGpaThreshold');
    if (inpGpa) inpGpa.addEventListener('change', (e) => syncMinGpa(e.target.value));
    if (rngGpa) rngGpa.addEventListener('input', (e) => syncMinGpa(e.target.value));
    document.querySelectorAll('.btn-preset-gpa').forEach((btn) => {
      btn.addEventListener('click', () => syncMinGpa(btn.getAttribute('data-gpa')));
    });
    document.querySelectorAll('input[name="gpaZeroRule"]').forEach((radio) => {
      radio.addEventListener('change', (e) => {
        if (!e.target.checked) return;
        STATE.studentGpaFilter.zeroRule = e.target.value;
        runMainCalculationAndRender();
      });
    });

    const chkLinked = document.getElementById('checkOnlyLinkedStudents');
    if (chkLinked) {
      chkLinked.addEventListener('change', (e) => {
        STATE.onlyLinkedStudentsForma2 = e.target.checked;
        applyParameterizationRules();
        runMainCalculationAndRender();
      });
    }

    // Controles del Motor Estadístico
    document.getElementById('selectSamplingDesign').addEventListener('change', (e) => {
      STATE.samplingConfig.design = e.target.value;
      if (STATE.activeEstamento === 'administrativos') {
        STATE.params.P9 = e.target.value === 'CENSO' ? 'CENSO' : 'MUESTREO';
        syncModalInputsFromState();
      }
      runMainCalculationAndRender();
    });

    document.getElementById('rangeConfidence').addEventListener('input', (e) => {
      const conf = Number(e.target.value);
      STATE.samplingConfig.confidence = conf;
      const z = Z_TABLE[conf] || 1.96;
      document.getElementById('valConfidenceDisplay').textContent = `${conf}% (Z = ${z.toFixed(4)})`;
      runMainCalculationAndRender();
    });

    const syncMarginError = (val) => {
      const clean = Math.min(10.0, Math.max(1.0, Number(val) || 5.0));
      STATE.samplingConfig.marginError = clean;
      document.getElementById('rangeMarginError').value = clean;
      document.getElementById('inputMarginError').value = clean.toFixed(1);
      document.querySelectorAll('.btn-preset').forEach((b) => {
        b.classList.toggle('active', Math.abs(Number(b.getAttribute('data-error')) - clean) < 0.05);
      });
      runMainCalculationAndRender();
    };

    document.getElementById('rangeMarginError').addEventListener('input', (e) => syncMarginError(e.target.value));
    document.getElementById('inputMarginError').addEventListener('change', (e) => syncMarginError(e.target.value));
    document.querySelectorAll('.btn-preset').forEach((btn) => {
      btn.addEventListener('click', () => syncMarginError(btn.getAttribute('data-error')));
    });

    document.getElementById('inputPropP').addEventListener('change', (e) => {
      const p = Math.min(0.95, Math.max(0.05, Number(e.target.value) || 0.50));
      STATE.samplingConfig.propP = p;
      document.getElementById('valPropPDisplay').textContent = p.toFixed(2);
      runMainCalculationAndRender();
    });

    document.getElementById('inputTNR').addEventListener('change', (e) => {
      const tnr = Math.min(50, Math.max(0, Number(e.target.value) || 0));
      STATE.samplingConfig.tnr = tnr;
      document.getElementById('valTNRDisplay').textContent = `${tnr}%`;
      runMainCalculationAndRender();
    });

    document.getElementById('checkMicroStrataCensus').addEventListener('change', (e) => {
      STATE.samplingConfig.microStrataCensus = e.target.checked;
      runMainCalculationAndRender();
    });

    // Controles de Discretización de Variables Cuantitativas
    document.getElementById('selectQuantVarActive').addEventListener('change', (e) => {
      STATE.quantConfig.activeVar = e.target.value;
      const cuts = STATE.quantConfig.customCutsByVar[STATE.quantConfig.activeVar] || [3.0, 4.0];
      document.getElementById('inputCustomCuts').value = cuts.join(', ');
      runMainCalculationAndRender();
    });

    document.getElementById('selectQuantMethod').addEventListener('change', (e) => {
      STATE.quantConfig.method = e.target.value;
      const isCustom = e.target.value === 'CUSTOM';
      document.getElementById('quantCustomCutsGroup').classList.toggle('hidden', !isCustom);
      document.getElementById('quantNumBinsGroup').classList.toggle('hidden', isCustom);
      runMainCalculationAndRender();
    });

    document.getElementById('selectQuantBins').addEventListener('change', (e) => {
      STATE.quantConfig.numBins = Number(e.target.value) || 3;
      document.getElementById('valQuantBinsChip').textContent = `L = ${STATE.quantConfig.numBins} intervalos`;
      runMainCalculationAndRender();
    });

    document.getElementById('btnApplyCustomCuts').addEventListener('click', () => {
      const rawText = document.getElementById('inputCustomCuts').value || '';
      const parsed = rawText
        .split(',')
        .map((s) => Number(s.trim()))
        .filter((n) => !Number.isNaN(n))
        .sort((a, b) => a - b);
      if (parsed.length > 0) {
        STATE.quantConfig.customCutsByVar[STATE.quantConfig.activeVar] = parsed;
      }
      runMainCalculationAndRender();
    });

    document.getElementById('checkSeparateZeroQuant').addEventListener('change', (e) => {
      STATE.quantConfig.separateZero = e.target.checked;
      runMainCalculationAndRender();
    });

    // Sorteo Aleatorio y Búsqueda en Tablas
    document.getElementById('btnNewRandomSeed').addEventListener('click', () => {
      const randomSeed = Math.floor(10000000 + Math.random() * 90000000);
      document.getElementById('inputRandomSeed').value = randomSeed;
      executeRandomSampleDraw();
    });

    document.getElementById('btnRunRandomDraw').addEventListener('click', () => {
      executeRandomSampleDraw();
    });

    document.getElementById('inputSearchStrata').addEventListener('input', () => {
      if (STATE.currentResult) renderStrataAllocationTable(STATE.currentResult);
    });

    document.getElementById('inputSearchSampleRows').addEventListener('input', () => {
      renderSelectedSampleTable();
    });

    document.getElementById('selectChartMetricView').addEventListener('change', () => {
      if (STATE.currentResult) renderCharts(STATE.currentResult);
    });

    // Botones de Exportación
    document.getElementById('btnExportSampleXLSX').addEventListener('click', () => exportCurrentSampleToExcel('xlsx'));
    document.getElementById('btnExportSampleCSV').addEventListener('click', () => exportCurrentSampleToExcel('csv'));
    document.getElementById('btnExportStrataTableXLSX').addEventListener('click', exportStrataTableToExcel);
    document.getElementById('btnExportBifasicoXLSX').addEventListener('click', exportBifasicoWorkbook);
    document.getElementById('btnExportGlobalSummaryXLSX').addEventListener('click', exportGlobalSummaryWorkbook);

    // Controles de Muestreo Bifásico
    document.getElementById('inputPilotTotalN').addEventListener('change', () => computeAndRenderBifasico(true));
    document.getElementById('selectBifasicoAllocation').addEventListener('change', () => computeAndRenderBifasico(true));
    document.getElementById('btnDrawBifasicoSamples').addEventListener('click', () => computeAndRenderBifasico(true));

    document.getElementById('btnSimulateRealisticPilot').addEventListener('click', () => {
      if (!STATE.currentResult) return;
      const prng = createPRNG((Number(document.getElementById('inputRandomSeed').value) || 2026) + 99);
      STATE.currentResult.strata.forEach((st) => {
        STATE.bifasico.pilotPropsByStratum[st.code] = Number((0.68 + prng() * 0.22).toFixed(2));
      });
      computeAndRenderBifasico(true);
    });

    document.getElementById('btnResetConservativePilot').addEventListener('click', () => {
      if (!STATE.currentResult) return;
      STATE.currentResult.strata.forEach((st) => {
        STATE.bifasico.pilotPropsByStratum[st.code] = 0.50;
      });
      computeAndRenderBifasico(true);
    });

    // Modal de Parametrización Inicial P1–P12 (con delegación de eventos infalible)
    const modal = document.getElementById('modalParametrizacion');
    const openModal = () => {
      syncModalInputsFromState();
      modal.classList.add('active');
    };
    const closeModal = () => modal.classList.remove('active');

    document.getElementById('btnOpenParamsModal').addEventListener('click', openModal);
    document.getElementById('btnCloseParamsModalTop').addEventListener('click', closeModal);
    document.getElementById('btnConfirmParamsModal').addEventListener('click', closeModal);
    document.getElementById('btnResetParamsModal').addEventListener('click', resetParamsToDefaults);
    document.getElementById('btnQuickResetDefaults').addEventListener('click', resetParamsToDefaults);

    // Clic directo sobre cualquiera de las tarjetas de opción (.param-choice-btn) en P1-P12
    modal.addEventListener('click', (e) => {
      const choiceBtn = e.target.closest('.param-choice-btn');
      if (choiceBtn) {
        e.preventDefault();
        const pKey = choiceBtn.getAttribute('data-param');
        const val = choiceBtn.getAttribute('data-val');
        applyParamOptionChange(pKey, val);
        return;
      }

      // Clic sobre las pestañas de filtro por Estamento dentro del modal (.modal-filter-btn)
      const filterBtn = e.target.closest('.modal-filter-btn');
      if (filterBtn) {
        e.preventDefault();
        const filterVal = filterBtn.getAttribute('data-filter');
        document.querySelectorAll('.modal-filter-btn').forEach((b) => b.classList.remove('active'));
        filterBtn.classList.add('active');

        document.querySelectorAll('#modalParamsScrollBody .param-section-group').forEach((sec) => {
          const g = sec.getAttribute('data-group');
          sec.style.display = filterVal === 'ALL' || g === filterVal ? 'block' : 'none';
        });
      }
    });

    document.querySelectorAll('.param-input').forEach((sel) => {
      sel.addEventListener('change', (e) => {
        const pKey = e.target.id.replace('param_', '');
        applyParamOptionChange(pKey, e.target.value);
      });
    });

    // Carga de Dataset Propio (.xlsx / .csv)
    const fileInput = document.getElementById('inputCustomDatasetFile');
    const dropzone = document.getElementById('customFileDropzone');

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        handleCustomFileUpload(e.target.files[0]);
      }
    });

    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.style.background = '#99f6e4';
    });
    dropzone.addEventListener('dragleave', () => {
      dropzone.style.background = '#f0fdfa';
    });
    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.style.background = '#f0fdfa';
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        handleCustomFileUpload(e.dataTransfer.files[0]);
      }
    });
  }

  // Exponer funciones clave globalmente
  window.SIMER_APP = {
    openParamsModal: () => {
      syncModalInputsFromState();
      document.getElementById('modalParametrizacion').classList.add('active');
    },
    getState: () => STATE
  };

  // Inicialización al cargar el DOM
  document.addEventListener('DOMContentLoaded', async () => {
    await loadAllDatasets();
    bindEventListeners();
    syncModalInputsFromState();
    applyParameterizationRules();
    renderVariablesChecklist();
    runMainCalculationAndRender();
    renderAllStaticMath();
    setTimeout(renderAllStaticMath, 200);
  });
})();
