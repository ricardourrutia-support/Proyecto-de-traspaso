/**
 * ============================================================
 *  CONSOLA DE AUDITORÍA — la aplicación web
 * ============================================================
 *
 * Una sola pantalla que te va diciendo qué toca hacer ahora.
 *
 *   Paso 1  Extraer      qué tickets faltan por extraer, la lista lista
 *                        para pegar en el extractor, y la carga del CSV.
 *   Paso 2  Auditar      qué tickets esperan veredicto, el prompt para
 *                        copiar, y dónde pegar la respuesta del agente.
 *   Paso 3  Revisar      la revisión humana de lo que dijo el agente.
 *
 * Todo lo pesado ya existe y está probado: Historico.gs baja Tableau,
 * Zendesk.gs normaliza, Auditoria.gs puntúa, Revision.gs registra la
 * revisión. Este archivo es la capa que los ordena y los muestra.
 *
 * CÓMO SE INSTALA — ver conInstrucciones() al final.
 */

const CON = {
  // Nombres de los archivos HTML en el proyecto, SIN la extensión.
  // Apps Script no deja que un .gs y un .html se llamen igual, por eso la H.
  // Si los renombras en el editor, cámbialos acá y en ningún otro lado.
  ARCHIVO_PAGINA:   "ConsolaH",
  ARCHIVO_REVISION: "RevisionH",
  ARCHIVO_REPORTES: "ReportesH",

  HOJA_INBOX:    "Inbox",
  MAX_EXTRACTOR: 500,                 // tope del formulario de n8n
  URL_EXTRACTOR: "https://n8n.cabify.tools/form/zendesk_support_tools",
  URL_AGENTE:    "https://marketplace-agent.cabify.tools/c/new",
  TZ:            "America/Santiago",
  TITULO:        "Consola de auditoría — Cabify Chile",
};


// ============================================================
//  CACHÉ DE EJECUCIÓN — leer cada hoja UNA vez
// ============================================================
//
// POR QUÉ EXISTE ESTO.
//
// Abrir la consola disparaba cuatro llamadas al servidor, y entre ellas
// leían el histórico entero CUATRO veces. La pestaña de resultados lo leía
// tres veces más en una sola llamada, y mantenimiento otra vez además de sus
// nueve lecturas por columna. Con el histórico en miles de filas cada lectura
// cuesta segundos, se hacen todas en paralelo y ninguna termina: resResumen
// pasó de 19 segundos por la mañana a 577 por la tarde, y cuatro recargas en
// tres minutos apilaron diez llamadas más sobre una cola ya saturada. Eso es
// lo que se veía como "la app se quedó pegada".
//
// La regla de acá es una sola: DENTRO DE UNA EJECUCIÓN, cada hoja se lee una
// vez. No es una caché entre peticiones —no hay nada que invalidar entre una
// y otra, y una caché con vida propia miente tarde o temprano—; es la misma
// llamada compartiendo la lectura que ya hizo.
//
// Vive en una variable global del script, que en Apps Script nace y muere con
// la ejecución. Cuando algo ESCRIBE en una hoja cacheada dentro de la misma
// ejecución hay que invalidarla a mano: los puntos donde eso pasa están
// marcados con conInvalidarCache_().
var CON_CACHE_ = {};

/**
 * Una hoja leída entera, una sola vez.
 *
 * Devuelve { existe, h, filas, n }: la cabecera ya limpia, las filas SIN la
 * cabecera, y cuántas son. Una hoja que no existe y una hoja vacía se
 * distinguen —`existe`— porque no son el mismo problema.
 */
function conHojaEnMemoria_(ss, nombre, maxCols) {
  const clave = nombre + (maxCols ? "|" + maxCols : "");
  if (CON_CACHE_[clave]) return CON_CACHE_[clave];

  const sh = ss.getSheetByName(nombre);
  let out;
  if (!sh) {
    out = { existe: false, h: [], filas: [], n: 0 };
  } else if (sh.getLastRow() < 2 || sh.getLastColumn() < 1) {
    out = { existe: true, h: [], filas: [], n: 0 };
  } else {
    const cols = maxCols ? Math.min(maxCols, sh.getLastColumn()) : sh.getLastColumn();
    const vals = sh.getRange(1, 1, sh.getLastRow(), cols).getValues();
    const h = vals[0].map(x => String(x || "").trim());
    out = { existe: true, h: h, filas: vals.slice(1), n: vals.length - 1 };
  }
  CON_CACHE_[clave] = out;
  return out;
}

/** El nombre de la hoja del histórico, con el mismo respaldo en todos lados. */
function conNombreHistorico_() {
  return (typeof HIST !== "undefined" && HIST.SHEET_NAME) ? HIST.SHEET_NAME : "Historico";
}

/** El histórico entero. Es la hoja ligera: no lleva conversaciones ni prompts. */
function conHistoricoEnMemoria_(ss) {
  return conHojaEnMemoria_(ss, conNombreHistorico_());
}

/**
 * Conversaciones, SOLO hasta el corte ligero.
 *
 * Nunca las tres columnas de texto largo: son 45.000 caracteres por ticket y
 * traerlas es exactamente lo que ya colgó la pestaña de revisión una vez.
 * Como el corte depende de la cabecera, se lee la cabecera primero —una fila—
 * y recién entonces el bloque.
 */
function conConversacionesEnMemoria_(ss) {
  const nombre = ZD.HOJA_CONVER;
  const clave = nombre + "|ligero";
  if (CON_CACHE_[clave]) return CON_CACHE_[clave];

  const sh = ss.getSheetByName(nombre);
  let out;
  if (!sh) {
    out = { existe: false, h: [], filas: [], n: 0 };
  } else if (sh.getLastRow() < 2 || sh.getLastColumn() < 1) {
    out = { existe: true, h: [], filas: [], n: 0 };
  } else {
    const hTodo = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x || "").trim());
    const corte = conCorteLigero_(hTodo);
    const cols = Math.max(1, Math.min(corte - 1, sh.getLastColumn()));
    const vals = sh.getRange(1, 1, sh.getLastRow(), cols).getValues();
    out = { existe: true, h: vals[0].map(x => String(x || "").trim()),
            filas: vals.slice(1), n: vals.length - 1 };
  }
  CON_CACHE_[clave] = out;
  return out;
}

/** Las revisiones humanas, enteras. Son pocas columnas y ninguna pesada. */
function conRevisionesEnMemoria_(ss) {
  const nombre = (typeof REV !== "undefined" && REV.HOJA_REVISIONES) ? REV.HOJA_REVISIONES : "Revisiones";
  return conHojaEnMemoria_(ss, nombre);
}

/**
 * Una hoja leída POR COLUMNAS, y solo las que hacen falta.
 *
 * El histórico tiene treinta y tantas columnas de Tableau y casi nadie las
 * quiere todas: la lista de revisión mira seis, el contador de auditables tres.
 * Traer la hoja entera para leer seis columnas es pagar treinta.
 *
 * Cada columna se lee una vez por ejecución y se guarda; pedirla de nuevo sale
 * gratis. Y si alguien ya trajo la hoja completa, se sirve de ahí sin volver a
 * la red.
 *
 * Devuelve { existe, n, h, col(nombre) } donde col() da el array de valores de
 * esa columna —sin cabecera— o un array de vacíos si la columna no está. Un
 * nombre que no existe no revienta: devuelve vacíos, que es lo que el código
 * de arriba ya sabe manejar.
 */
function conColumnas_(ss, nombre, nombres) {
  const claveH = nombre + "|cols";
  let memo = CON_CACHE_[claveH];

  if (!memo) {
    const entera = CON_CACHE_[nombre];
    const sh = ss.getSheetByName(nombre);
    if (!sh) { memo = { existe: false, n: 0, h: [], datos: {} }; }
    else if (sh.getLastRow() < 2 || sh.getLastColumn() < 1) {
      memo = { existe: true, n: 0, h: [], datos: {} };
    } else {
      memo = {
        existe: true,
        n: sh.getLastRow() - 1,
        h: entera ? entera.h
                  : sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                      .map(x => String(x || "").trim()),
        datos: {},
        sh: sh,
      };
    }
    CON_CACHE_[claveH] = memo;
  }

  const entera = CON_CACHE_[nombre];
  (nombres || []).forEach(n => {
    if (memo.datos[n] !== undefined || !memo.n) return;
    const k = memo.h.indexOf(n);
    if (k < 0) { memo.datos[n] = null; return; }
    // La hoja completa ya está en memoria: sacarla de ahí en vez de pedirla.
    memo.datos[n] = entera ? entera.filas.map(f => f[k])
                           : memo.sh.getRange(2, k + 1, memo.n, 1).getValues().map(f => f[0]);
  });

  const vacio = new Array(memo.n).fill("");
  return {
    existe: memo.existe, n: memo.n, h: memo.h,
    col: n => (memo.datos[n] === undefined || memo.datos[n] === null) ? vacio : memo.datos[n],
  };
}

/**
 * Olvida lo leído. Se llama DESPUÉS de escribir en una hoja y ANTES de
 * volver a contarla en la misma ejecución.
 *
 * Sin nombre olvida todo. Es lo correcto por defecto: equivocarse de lado
 * cuesta una lectura de más; equivocarse del otro hace que la pantalla
 * informe un número viejo con toda la confianza del mundo.
 */
function conInvalidarCache_(nombre) {
  if (!nombre) { CON_CACHE_ = {}; return; }
  Object.keys(CON_CACHE_).forEach(k => {
    if (k === nombre || k.indexOf(nombre + "|") === 0) delete CON_CACHE_[k];
  });
}

// ============================================================
//  AUDIENCIAS
// ============================================================

/**
 * Las audiencias que hay HOY en el histórico, con cuántos tickets auditables
 * tiene cada una. Alimenta el selector de la cabecera.
 *
 * Se devuelven solo las que tienen tickets: un selector con una audiencia
 * vacía invita a filtrar por ella y quedarse mirando una pantalla en blanco
 * sin entender por qué.
 */
function conAudiencias() {
  const ss = histGetSpreadsheet_(true);
  const cuenta = {};
  conAuditables_(ss, null, cuenta);

  const out = audAudiencias_()
    .filter(k => cuenta[k])
    .map(k => ({ clave: k, nombre: audNombreAudiencia_(k), auditables: cuenta[k] }));

  // Si el histórico todavía no tiene la columna, no se ofrece filtro: es
  // preferible una consola sin selector que un selector que no filtra nada.
  return { audiencias: out, total: audAudiencias_().reduce((a, k) => a + (cuenta[k] || 0), 0) };
}

// ============================================================
//  SERVIR LA APP
// ============================================================

function doGet(e) {
  return conServirApp();
}

function conServirApp() {
  return HtmlService.createTemplateFromFile(CON.ARCHIVO_PAGINA)
    .evaluate()
    .setTitle(CON.TITULO)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Permite partir el HTML en varios archivos. */
function include(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

/** El nombre del archivo de la pestaña 3, para que el HTML no lo hardcodee. */
function conArchivoRevision() {
  return CON.ARCHIVO_REVISION;
}

/** Ídem para la pestaña 4. */
function conArchivoReportes() {
  return CON.ARCHIVO_REPORTES;
}

/** Las URLs que la interfaz muestra como enlaces. */
function conEnlaces() {
  return { extractor: CON.URL_EXTRACTOR, agente: CON.URL_AGENTE, tope: CON.MAX_EXTRACTOR };
}

// ============================================================
//  ESTADO — el semáforo de la cabecera
// ============================================================

/**
 * Cuenta en qué punto del ciclo está cada ticket auditable.
 * Lee solo la columna de ticket de cada hoja: no arrastra las
 * conversaciones, que son lo pesado.
 */
function conEstado(audiencia) {
  const ss = histGetSpreadsheet_(true);
  const auditables = conAuditables_(ss, audiencia);
  const extraidos  = conClaves_(ss, ZD.HOJA_CONVER);
  const conVer     = conClaves_(ss, AUD.HOJA_VEREDICTOS);
  const revisados  = conClaves_(ss, REV.HOJA_REVISIONES);

  const pendExtraccion = auditables.filter(t => !extraidos[t]);
  const auditados      = auditables.filter(t => conVer[t]);
  const pendRevision   = auditados.filter(t => !revisados[t]);

  // "Por auditar" se cuenta con LA MISMA función que llena la tabla del paso 2,
  // no con una cuenta paralela sobre el histórico. Antes eran dos cálculos
  // distintos y por eso podían no cuadrar: la cabecera decía 227 y la tabla
  // mostraba 36. Y el desajuste no era solo la columna Audiencia — el histórico
  // y Conversaciones también difieren en "Auditable", porque Conversaciones sabe
  // cosas que el histórico no puede saber (que el canal es una llamada, que no
  // hay ni un turno de agente).
  //
  // Un contador que no cuadra con su propia lista es peor que un contador
  // caro. Calcularlo desde la lista lo hace imposible por construcción.
  const pendVeredicto = conPendientesVeredicto(audiencia);

  return {
    auditables:      auditables.length,
    pendExtraccion:  pendExtraccion.length,
    pendVeredicto:   pendVeredicto.length,
    auditados:       auditados.length,
    pendRevision:    pendRevision.length,
    revisados:       auditados.filter(t => revisados[t]).length,
    tope:            CON.MAX_EXTRACTOR,
    audiencia:       audiencia || "",
    actualizado:     Utilities.formatDate(new Date(), CON.TZ, "HH:mm"),
  };
}

/**
 * TODO lo que la consola necesita para pintarse al abrir, en UNA llamada.
 *
 * Antes eran cuatro —enlaces, audiencias, estado y lista— y cada una abría su
 * propia ejecución, así que cada una releía el histórico por su cuenta: cuatro
 * lecturas de la hoja más pesada para dibujar una pantalla. Juntas comparten
 * la lectura y son una sola.
 *
 * Si una parte falla, las otras igual llegan: un enlace roto no puede dejar la
 * consola en blanco.
 */
function conArranque(audiencia) {
  const intenta = (fn, sivale) => { try { return fn(); } catch (e) { return sivale; } };
  return {
    enlaces:    intenta(() => conEnlaces(), { extractor: CON.URL_EXTRACTOR, agente: CON.URL_AGENTE }),
    audiencias: intenta(() => conAudiencias(), { audiencias: [], total: 0 }),
    estado:     conEstado(audiencia),
    lista:      intenta(() => conListaParaExtraer(audiencia), null),
  };
}

/**
 * Tickets marcados como auditables en el histórico.
 *
 * `audiencia` filtra por la columna Audiencia; vacío o nulo = todas.
 * `cuenta` es un objeto opcional donde se acumula el total por audiencia,
 * para no tener que releer la hoja entera solo para contar.
 */
function conAuditables_(ss, audiencia, cuenta) {
  // La lectura se comparte con todo lo demás que mira el histórico en esta
  // misma ejecución. Antes cada una de las cinco funciones que llaman acá se
  // traía la hoja entera por su cuenta.
  const hist = conHistoricoEnMemoria_(ss);
  if (!hist.n) return [];
  const h = hist.h;
  const cT = h.indexOf("Ticket Number"), cA = h.indexOf("Auditable"), cAud = h.indexOf("Audiencia");
  if (cT < 0) return [];

  const filtro = String(audiencia || "").trim();
  const out = [];
  hist.filas.forEach(r => {
    // Esta es la columna del HISTÓRICO: su "NO" es el veredicto de alcance y
    // es autoritativo.
    if (cA >= 0 && !audAlcanceHistoricoOk_(r[cA])) return;
    const id = zdClaveTicket_(r[cT]);
    if (!id) return;

    // Una fila sin audiencia escrita (histórico de antes de la columna) cae en
    // la por defecto, que es la que se usó para auditarla.
    const a = cAud >= 0 ? (String(r[cAud] || "").trim() || AUD_AUDIENCIA_DEFECTO) : AUD_AUDIENCIA_DEFECTO;
    if (cuenta) cuenta[a] = (cuenta[a] || 0) + 1;
    if (filtro && a !== filtro) return;
    out.push(id);
  });
  return out;
}

/**
 * Columnas de Conversaciones que pesan: la conversación completa y los dos
 * prompts. Cada una puede tener 45.000 caracteres POR TICKET.
 *
 * Traerlas cuando no se necesitan es la diferencia entre una consulta
 * instantánea y una que se cuelga: con 500 tickets son cientos de MB que
 * Apps Script tiene que serializar para nada. Devuelve el índice (1-based)
 * de la primera columna pesada, para poder leer solo lo que va antes.
 */
// Las tres columnas de texto largo. Todo lo que va a su izquierda es ligero y se
// puede leer en bloque; a partir de acá se lee fila por fila. "Señales" queda a
// propósito del lado ligero: es un JSON corto y hace falta para armar el prompt.
const CON_COLS_PESADAS = ["PROMPT AUDITORIA (copiar y pegar)", "Prompt solo viaje", "Conversacion"];

function conCorteLigero_(headers) {
  let corte = headers.length + 1;
  CON_COLS_PESADAS.forEach(n => {
    const i = headers.indexOf(n);
    if (i >= 0) corte = Math.min(corte, i + 1);
  });
  return corte;
}

/** Conjunto de tickets presentes en una hoja, leyendo solo la 1ª columna. */
function conClaves_(ss, nombre) {
  const out = {};
  const sh = ss.getSheetByName(nombre);
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => {
    const k = zdClaveTicket_(r[0]);
    if (k) out[k] = true;
  });
  return out;
}

// ============================================================
//  PASO 1 — EXTRACCIÓN
// ============================================================

/**
 * La lista de ticket_id para pegar en el extractor, separada por comas.
 * Corta en el tope del formulario y dice cuántos quedan para la vuelta
 * siguiente. Si tienes 600 pendientes, entrega 500 y avisa que restan 100.
 */
function conListaParaExtraer(audiencia) {
  const ss = histGetSpreadsheet_(true);
  const auditables = conAuditables_(ss, audiencia);
  const extraidos  = conClaves_(ss, ZD.HOJA_CONVER);
  const pendientes = auditables.filter(t => !extraidos[t]);

  const entrega = pendientes.slice(0, CON.MAX_EXTRACTOR);
  return {
    total:     pendientes.length,
    entregados: entrega.length,
    restantes: Math.max(0, pendientes.length - entrega.length),
    lista:     entrega.join(","),
    tope:      CON.MAX_EXTRACTOR,
    url:       CON.URL_EXTRACTOR,
  };
}

/** Deja la hoja Inbox vacía y con la cabecera del CSV. */
function conCargaIniciar(cabecera) {
  if (!cabecera || !cabecera.length) throw new Error("El CSV no trae cabecera.");
  const ss = histGetSpreadsheet_(true);
  let sh = ss.getSheetByName(CON.HOJA_INBOX);
  if (!sh) sh = ss.insertSheet(CON.HOJA_INBOX);
  sh.clear();
  sh.getRange(1, 1, 1, cabecera.length).setValues([cabecera]);
  return { ok: true, columnas: cabecera.length };
}

/**
 * Recibe un trozo de filas del CSV. La interfaz parte el archivo en el
 * navegador y lo manda por partes: un CSV de 500 tickets pesa varios MB
 * y mandarlo de una sola vez se cae.
 */
function conCargaTrozo(filas) {
  if (!filas || !filas.length) return { ok: true, filas: 0 };
  const ss = histGetSpreadsheet_(true);
  const sh = ss.getSheetByName(CON.HOJA_INBOX);
  if (!sh) throw new Error("Falta iniciar la carga antes de mandar filas.");

  const ancho = Math.max.apply(null, filas.map(f => f.length));
  const norm  = filas.map(f => {
    const c = f.slice();
    while (c.length < ancho) c.push("");
    return c;
  });
  sh.getRange(sh.getLastRow() + 1, 1, norm.length, ancho).setValues(norm);
  return { ok: true, filas: norm.length, total: sh.getLastRow() - 1 };
}

/**
 * Con el CSV ya en la hoja, normaliza y enriquece.
 * Reusa zdIngestar(), que acumula: los tickets ya auditados no se tocan.
 */
function conCargaProcesar() {
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const antes = conEstado();
    const r = zdIngestar();
    // OBLIGATORIO. zdIngestar acaba de escribir en Conversaciones y en el
    // histórico; sin olvidar lo leído, el "despues" se calcularía con la foto
    // de antes y "extraidosAhora" saldría 0 siempre, justo cuando la carga sí
    // funcionó. Un contador que dice cero después de un trabajo que salió bien
    // es peor que no tener contador.
    conInvalidarCache_();
    const despues = conEstado();
    return {
      ok: true,
      comentarios: r.comentarios,
      tickets: r.tickets,
      porCanal: r.porCanal,
      extraidosAhora: antes.pendExtraccion - despues.pendExtraccion,
      estado: despues,
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  PASO 2 — AUDITORÍA
// ============================================================

/**
 * Tickets con conversación pero sin veredicto. Devuelve lo justo para
 * pintar la tabla; el prompt se pide aparte porque son miles de
 * caracteres por ticket y no vale la pena traerlos todos de golpe.
 */
function conPendientesVeredicto(audiencia) {
  const ss = histGetSpreadsheet_(true);
  const filtro = String(audiencia || "").trim();
  const conVer = conClaves_(ss, AUD.HOJA_VEREDICTOS);
  // La audiencia según el Histórico, que es su fuente de verdad. La columna de
  // Conversaciones es un atajo y puede estar vacía en las filas extraídas antes
  // de que existiera: dar esas por B2B era lo que hacía que la cabecera dijera
  // 227 y la tabla mostrara 36.
  const audHist = audMapaAudiencias_();
  // El ALCANCE se consulta en vivo al histórico. La celda "Auditable" de
  // Conversaciones se escribió una vez, al ingerir, y si el histórico se
  // corrige después esa celda se queda con el valor viejo — así se quedaron
  // 191 tickets invisibles.
  const alcanceOk = {};
  conAuditables_(ss, filtro).forEach(t => { alcanceOk[t] = true; });
  // Solo las columnas ligeras, y compartidas con el resto de la ejecución: la
  // tabla no muestra ni conversación ni prompts.
  const conv = conConversacionesEnMemoria_(ss);
  if (!conv.n) return [];
  const h = conv.h;
  const c = n => h.indexOf(n);
  const cols = ["Ticket Number", "Audiencia", "Canal", "Journey Id", "Fecha del viaje",
                "Auditable", "Banderas", "Turnos", "Motivo"];
  const idx = {};
  cols.forEach(n => { idx[n] = c(n); });
  if (idx["Ticket Number"] < 0) return [];

  const out = [];
  conv.filas.forEach(r => {
    const t = zdClaveTicket_(r[idx["Ticket Number"]]);
    if (!t || conVer[t]) return;

    // Un "NO (...)" es de la CONVERSACIÓN y es permanente: una llamada sin
    // transcripción o un hilo sin turnos de agente no se pueden auditar nunca.
    // Cualquier otro valor —"SI", "PARCIAL", "FUERA DE ALCANCE"— se decide con
    // el histórico, que es el que sabe el alcance de hoy.
    // Y esta es la de CONVERSACIONES: solo bloquea lo que impide el hilo.
    const aud = idx["Auditable"] >= 0 ? String(r[idx["Auditable"]] || "").trim() : "SI";
    if (!audConversacionAuditable_(aud)) return;
    if (!alcanceOk[t]) return;
    const v = n => idx[n] >= 0 ? String(r[idx[n]] || "") : "";
    // Celda -> histórico -> por defecto. Un valor ausente no es un valor por
    // defecto: preguntarlo antes de suponerlo es toda la diferencia.
    const a = v("Audiencia") || audHist[t] || AUD_AUDIENCIA_DEFECTO;
    if (filtro && a !== filtro) return;
    out.push({
      ticket: t, audiencia: a, audienciaNombre: audNombreAudiencia_(a),
      canal: v("Canal"), journey: v("Journey Id"),
      journeyDate: v("Fecha del viaje"), banderas: v("Banderas"),
      turnos: v("Turnos"), motivo: v("Motivo"),
      sinFecha: !v("Fecha del viaje"),
    });
  });
  return out;
}

/** El prompt ya armado de un ticket, para copiar y pegar en el agente. */
function conPromptDe(ticket) {
  const t = zdClaveTicket_(ticket);

  // Se pide el prompt VIGENTE: si el guardado es de una versión anterior de
  // las reglas, se regenera solo antes de entregarlo. Lo que se copia y se
  // pega en el agente tiene que llevar las reglas de hoy sin que nadie se
  // acuerde de refrescar nada.
  const pv = zdPromptVigenteDe_(t);
  if (!pv.prompt || !pv.prompt.trim())
    throw new Error("El ticket " + t + " no tiene prompt: " +
                    (pv.motivo || "motivo desconocido") + ".");

  return { ticket: t, prompt: pv.prompt, url: CON.URL_AGENTE,
           regenerado: !!pv.regenerado, motivoRegenerado: pv.regenerado ? pv.motivo : "" };
}

/**
 * Recibe la respuesta del agente pegada tal cual, la procesa y deja la
 * nota escrita. Es exactamente la misma cadena que ya usábamos a mano.
 */
function conGuardarVeredicto(ticket, csv, json) {
  const t = zdClaveTicket_(ticket);
  if (!t) throw new Error("Falta el número de ticket.");

  // El agente responde en dos bloques y la interfaz los recibe por separado.
  // Se pegan acá porque el parser los lee juntos. Si json viene vacío, se
  // procesa igual: los veredictos son lo obligatorio, el viaje es un extra.
  const bloqueCsv  = String(csv  || "").trim();
  const bloqueJson = String(json || "").trim();

  if (!bloqueCsv && !bloqueJson) throw new Error("No pegaste nada.");
  if (!bloqueCsv)
    throw new Error("Falta el bloque CSV. Es el que trae los veredictos, " +
                    "el que empieza con la cabecera criterio_id;veredicto;…");

  const texto = bloqueJson ? bloqueCsv + "\n" + bloqueJson : bloqueCsv;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const canal = conCanalDe_(t);
    const pegado = audParsearTexto_(texto, t);
    const veredictos = pegado.veredictos[t];
    if (!veredictos || !veredictos.length)
      throw new Error("No encontré veredictos para el ticket " + t +
                      ". Revisa que hayas pegado el bloque CSV completo, con su cabecera.");

    const r = audCalcularNota_(canal, veredictos, audAudienciaDeTicket_(t));
    audEscribirVeredictos_(t, canal, veredictos, r);
    if (pegado.viajes[t]) { audEscribirViaje_(t, pegado.viajes[t]); r.viajeGuardado = true; }
    audEscribirEnHistorico_(t, r);

    return {
      ok: true, ticket: t, canal: canal,
      nota: r.nota, techo: r.techo, estado: r.estado,
      cobertura: r.cobertura, aplicables: r.aplicables, evaluados: r.evaluados,
      criticos: r.criticosIncumplidos, porConfirmar: r.criticosPorConfirmar,
      noConcluyentes: r.noConcluyentes.length,
      viajeGuardado: !!r.viajeGuardado,
      avisos: r.avisos,
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Audita UN ticket llamando a la API del agente. Es lo mismo que hacías a
 * mano, sin el copiar y pegar.
 *
 * De a uno a propósito: cada auditoría consulta BigQuery y puede tardar cerca
 * de un minuto. Un lote grande en una sola llamada chocaría con el límite de
 * ejecución de Apps Script y perderías todo el lote. La interfaz encadena las
 * llamadas y va mostrando el avance, así puedes mirar y parar cuando quieras.
 */
function conAuditarAuto(ticket) {
  const t = zdClaveTicket_(ticket);
  if (!t) throw new Error("Falta el número de ticket.");
  if (typeof audAuditarTicket !== "function")
    throw new Error("Falta el archivo Marketplace en el proyecto: es el que habla con la API del agente.");

  const r = audAuditarTicket(t);
  return {
    ok: true, ticket: t,
    nota: r.nota, techo: r.techo, estado: r.estado,
    cobertura: r.cobertura, aplicables: r.aplicables, evaluados: r.evaluados,
    criticos: r.criticosIncumplidos || [], porConfirmar: r.criticosPorConfirmar || [],
    noConcluyentes: (r.noConcluyentes || []).length,
    fueraDeAlcance: r.fueraDeAlcance || [],
    viajeGuardado: !!r.viajeGuardado,
    promptRegenerado: !!r.promptRegenerado,
    avisos: r.avisos || [],
  };
}

/** Los tickets que faltan por auditar, solo sus números, para el ciclo. */
function conColaDeAuditoria(limite, audiencia) {
  const pend = conPendientesVeredicto(audiencia).filter(x => !x.sinFecha || true);
  const n = Number(limite) > 0 ? Number(limite) : pend.length;
  return pend.slice(0, n).map(x => x.ticket);
}

function conCanalDe_(ticket) {
  const ss = histGetSpreadsheet_(true);
  const sh = ss.getSheetByName(ZD.HOJA_CONVER);
  if (!sh || sh.getLastRow() < 2) return "ticket";
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x || "").trim());
  const cC = h.indexOf("Canal");
  if (cC < 0) return "ticket";
  const nFila = zdFilaDeTicket_(sh, ticket);
  if (nFila < 0) return "ticket";
  return String(sh.getRange(nFila, cC + 1).getValue() || "ticket");
}

// ============================================================
//  MANTENIMIENTO — el tablero dice qué hay que hacer
// ============================================================

/**
 * Los arreglos que la consola puede ejecutar sola, por su id.
 *
 * Es una LISTA BLANCA a propósito. La interfaz manda un id, no un nombre de
 * función: si mandara el nombre, cualquiera con la URL de la app podría pedir
 * que se ejecute cualquier función del proyecto, incluidas las que borran.
 * Acá solo entran operaciones que no destruyen nada.
 */
const CON_REPARACIONES = {
  audiencias: {
    etiqueta: "Completar la columna Audiencia",
    fn: () => zdCompletarAudiencias(),
    hecho: r => "Audiencia completada en " + r.tocadas + " filas.",
  },
  prompts: {
    etiqueta: "Regenerar todos los prompts",
    fn: () => zdRefrescarPrompts(),
    hecho: r => "Prompts regenerados: " + r.tocados + ".",
  },
  procedimientos: {
    etiqueta: "Sincronizar Confluence",
    fn: () => confSincronizar(),
    hecho: r => "Procedimientos sincronizados: " + ((r && r.nuevos) || 0) + " nuevos.",
  },
  tags: {
    etiqueta: "Sincronizar el catálogo de tags",
    fn: () => tagsSincronizar(),
    hecho: r => r.tags + " tags y " + r.pautas + " pautas sincronizadas.",
  },
  trigger: {
    etiqueta: "Activar la carga diaria",
    fn: () => histCrearTriggerDiario(),
    hecho: () => "Carga diaria agendada.",
  },
  autoReintentar: {
    etiqueta: "Devolver los apartados a la cola",
    fn: () => autoReintentar(""),
    hecho: () => "Los tickets apartados vuelven a la cola en la próxima tanda.",
  },
  autoApagar: {
    etiqueta: "Apagar el modo automático",
    fn: () => autoDesactivar("apagado desde el panel de mantenimiento"),
    hecho: () => "Modo automático apagado.",
  },
  historico: {
    etiqueta: "Actualizar el histórico desde Tableau",
    fn: () => histActualizar(),
    hecho: r => "Histórico actualizado: " + r.nuevos + " nuevos, " + r.actualizados + " actualizados.",
  },
  auditableConver: {
    etiqueta: "Recalcular qué conversaciones son auditables",
    fn: () => zdRecalcularAuditable(),
    hecho: r => r.cambios + " filas actualizadas; " + r.desbloqueados +
                " estaban bloqueadas por un valor viejo.",
  },
  audienciasHist: {
    etiqueta: "Recalcular audiencia y alcance del histórico",
    fn: () => histCompletarAudiencias(),
    hecho: r => "Audiencia escrita en " + r.audiencias + " filas; alcance recalculado en " +
                r.alcance + " (" + r.nuevosAuditables + " pasaron a auditables).",
  },
};

/**
 * Ejecuta uno de los arreglos de la lista blanca.
 *
 * Nada de lo que hay acá borra trabajo. Lo que sí borra —rehacer auditorías,
 * volver a extraer— NO se ejecuta desde este panel: se sugiere, con la lista
 * de tickets, y la decisión la toma una persona en el paso que corresponde.
 * Un botón que borra ochenta auditorías no debería estar a un clic de
 * distancia de uno que rellena una columna.
 */
function conRepararAhora(id) {
  const r = CON_REPARACIONES[String(id || "")];
  if (!r) throw new Error("Reparación desconocida: " + id);
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const salida = r.fn();
    // Una reparación reescribe hojas. Lo que se lea después en esta misma
    // ejecución tiene que ver el resultado, no la foto anterior.
    conInvalidarCache_();
    return { ok: true, id: id, mensaje: r.hecho(salida || {}) };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ¿POR QUÉ VEO 36 Y NO 227?
 *
 * Un ticket recorre cuatro filtros antes de aparecer en el paso 2, y cada uno
 * lo puede dejar fuera por un motivo distinto. Cuando el número final no
 * cuadra con la intuición, mirar solo el número no sirve de nada: hay que ver
 * en qué escalón se cayeron.
 *
 * Esta función cuenta los cuatro escalones y, en cada uno, muestra los
 * primeros tickets que se perdieron. Es un instrumento, no un arreglo: lo que
 * hace es dejar de adivinar.
 *
 * Lee SOLO columnas ligeras. La columna de la conversación son 45.000
 * caracteres por fila y leerla es lo que cuelga estas pantallas.
 */
function conPorQueFaltan(audiencia) {
  const ss = histGetSpreadsheet_(true);
  const filtro = String(audiencia || "").trim();
  const L = [];
  const nombreHist = (typeof HIST !== "undefined" && HIST.SHEET_NAME) ? HIST.SHEET_NAME : "Historico";

  // ---- Escalón 1: el histórico ----
  const shH = ss.getSheetByName(nombreHist);
  if (!shH || shH.getLastRow() < 2)
    return { error: "El histórico está vacío." };

  const hH = shH.getRange(1, 1, 1, shH.getLastColumn()).getValues()[0].map(x => String(x || "").trim());
  const iH = n => hH.indexOf(n);
  const nH = shH.getLastRow() - 1;
  const col = i => i >= 0 ? shH.getRange(2, i + 1, nH, 1).getValues().map(f => f[0]) : new Array(nH).fill("");

  const tks     = col(iH("Ticket Number")).map(v => zdClaveTicket_(v));
  const auds    = col(iH("Audiencia")).map(v => String(v || "").trim());
  const grupos  = col(iH("Group Name")).map(v => String(v || "").trim());
  const audita  = col(iH("Auditable")).map(v => String(v || "").trim());
  const motivos = col(iH("Motivo no auditable")).map(v => String(v || "").trim());

  // La audiencia REAL de cada fila: la escrita, y si está vacía la que dice su
  // Group Name. Una celda vacía no es "b2b": es una celda vacía.
  const audReal = tks.map((t, i) => auds[i] && AUD_AUDIENCIAS[auds[i]]
    ? auds[i]
    : audAudienciaDeGrupo_(grupos[i], null));

  const idx = [];
  for (let i = 0; i < nH; i++) if (tks[i] && (!filtro || audReal[i] === filtro)) idx.push(i);

  const sinEscribir = idx.filter(i => !auds[i]).length;
  L.push({ paso: "1. En el histórico", cuenta: idx.length,
           nota: sinEscribir ? sinEscribir + " de ellos con la columna Audiencia sin escribir " +
                               "(se deducen del Group Name)" : "" });

  // ---- Escalón 2: marcados auditables ----
  const auditables = idx.filter(i => audAlcanceHistoricoOk_(audita[i]));
  const fuera = idx.filter(i => !audAlcanceHistoricoOk_(audita[i]));
  const porMotivo = {};
  fuera.forEach(i => {
    const m = motivos[i] || "(sin motivo escrito)";
    if (!porMotivo[m]) porMotivo[m] = [];
    porMotivo[m].push(tks[i]);
  });
  L.push({ paso: "2. Auditables", cuenta: auditables.length,
           perdidos: fuera.length,
           detalle: Object.keys(porMotivo).map(m => ({
             motivo: m, cuantos: porMotivo[m].length, ejemplos: porMotivo[m].slice(0, 5) })) });

  // ---- Escalón 3: extraídos (están en Conversaciones) ----
  const extraidos = conClaves_(ss, ZD.HOJA_CONVER);
  const conConver = auditables.filter(i => extraidos[tks[i]]);
  const sinExtraer = auditables.filter(i => !extraidos[tks[i]]).map(i => tks[i]);
  L.push({ paso: "3. Con conversación extraída", cuenta: conConver.length,
           perdidos: sinExtraer.length,
           nota: sinExtraer.length ? "Estos hay que extraerlos en el paso 1." : "",
           tickets: sinExtraer });

  // ---- Escalón 4: lo que la tabla del paso 2 muestra ----
  //
  // El motivo se MIDE, leyendo la celda "Auditable" de cada fila. La primera
  // versión de esto agrupaba todo lo que no llegaba a la tabla bajo "la
  // conversación los descarta (llamada, sin turnos de agente)" — una
  // explicación que no había comprobado. Eran 191 tickets bloqueados por un
  // valor viejo en esa celda, y la etiqueta me hizo perder una ronda entera
  // buscando en el lugar equivocado.
  const conVer = conClaves_(ss, AUD.HOJA_VEREDICTOS);
  const yaAuditados = conConver.filter(i => conVer[tks[i]]).map(i => tks[i]);
  const enTabla = conPendientesVeredicto(filtro);
  const enTablaSet = {};
  enTabla.forEach(x => { enTablaSet[x.ticket] = true; });

  const shC2 = ss.getSheetByName(ZD.HOJA_CONVER);
  const audDe = {};
  if (shC2 && shC2.getLastRow() > 1) {
    const hC2 = shC2.getRange(1, 1, 1, shC2.getLastColumn()).getValues()[0]
      .map(x => String(x || "").trim());
    const cT2 = hC2.indexOf("Ticket Number"), cA2 = hC2.indexOf("Auditable");
    if (cT2 >= 0 && cA2 >= 0) {
      const nC = shC2.getLastRow() - 1;
      const id2 = shC2.getRange(2, cT2 + 1, nC, 1).getValues();
      const au2 = shC2.getRange(2, cA2 + 1, nC, 1).getValues();
      for (let k = 0; k < nC; k++) {
        const t = zdClaveTicket_(id2[k][0]);
        if (t) audDe[t] = String(au2[k][0] || "").trim();
      }
    }
  }

  const porQue = {};
  const descartados = [];
  conConver.forEach(i => {
    const t = tks[i];
    if (conVer[t] || enTablaSet[t]) return;
    descartados.push(t);
    const v = audDe[t] || "(sin valor)";
    // El texto entre paréntesis ES el motivo que se escribió al ingerir: se
    // muestra tal cual en vez de resumirlo por nosotros.
    const clave = !audConversacionAuditable_(v)
      ? v
      : 'la celda "Auditable" dice "' + v + '" (valor de la ingesta, desactualizado)';
    if (!porQue[clave]) porQue[clave] = [];
    porQue[clave].push(t);
  });

  const detalle4 = [];
  if (yaAuditados.length)
    detalle4.push({ motivo: "ya tienen veredicto", cuantos: yaAuditados.length,
                    ejemplos: yaAuditados.slice(0, 5) });
  Object.keys(porQue).forEach(k => detalle4.push({
    motivo: k, cuantos: porQue[k].length, ejemplos: porQue[k].slice(0, 5) }));

  const desactualizados = descartados.filter(t => audConversacionAuditable_(audDe[t]));

  L.push({ paso: "4. En la tabla del paso 2", cuenta: enTabla.length,
           perdidos: yaAuditados.length + descartados.length,
           nota: desactualizados.length
             ? desactualizados.length + ' de estos están bloqueados por un valor viejo en la ' +
               'celda "Auditable", no por la conversación. Se arregla en Mantenimiento con ' +
               '"Recalcular qué conversaciones son auditables".'
             : "",
           detalle: detalle4,
           tickets: descartados });

  return { audiencia: filtro, pasos: L,
           actualizado: Utilities.formatDate(new Date(), CON.TZ, "HH:mm") };
}

/**
 * Revisa el estado del sistema y devuelve qué conviene hacer.
 *
 * Tres niveles, y la diferencia importa:
 *   "problema"   algo está roto o incompleto y hay un arreglo automático.
 *   "decision"   hay trabajo que rehacer, pero destruye lo hecho: se sugiere.
 *   "info"       está bien, o es el flujo normal. NO cuenta para el contador.
 *
 * El contador de la pestaña solo suma problemas y decisiones. Si sumara los
 * informativos estaría siempre encendido, y un aviso que suena siempre deja
 * de informar — es la misma razón por la que los críticos que nunca se pueden
 * comprobar salieron de la matriz automática.
 */
function conSalud(audiencia) {
  const ss = histGetSpreadsheet_(true);
  const items = [];
  const add = (nivel, id, titulo, detalle, accion, sugerencia, tickets) => items.push({
    nivel: nivel, id: id, titulo: titulo, detalle: detalle,
    accion: accion ? { id: accion, etiqueta: CON_REPARACIONES[accion].etiqueta } : null,
    sugerencia: sugerencia || "", tickets: tickets || [],
  });

  const hoja = n => ss.getSheetByName(n);
  const filas = n => { const h = hoja(n); return h ? Math.max(0, h.getLastRow() - 1) : -1; };
  const cabecera = n => {
    const h = hoja(n);
    if (!h || h.getLastColumn() < 1) return [];
    return h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(x => String(x || "").trim());
  };

  // ---- 1. El histórico ----
  const nHist = filas(HIST.SHEET_NAME || "Historico");
  if (nHist <= 0) {
    add("problema", "historico", "El histórico está vacío",
        "Sin histórico no hay tickets que auditar.", "historico");
  } else {
    const trig = ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === "histActualizar").length;
    if (!trig)
      add("problema", "trigger", "La carga diaria no está activa",
          "El histórico solo se actualiza cuando alguien lo pide a mano.", "trigger");
  }

  // ---- 2. La columna Audiencia de Conversaciones ----
  // Todas las cuentas sobre Conversaciones salen de UNA lectura de sus
  // columnas ligeras. Antes este panel leía la hoja cuatro veces, columna por
  // columna, y encima volvía a leer el histórico nueve veces más abajo: por eso
  // tardaba 541 segundos y se comía la cola de todo lo demás.
  const conv = conConversacionesEnMemoria_(ss);
  const hC = conv.h;
  if (conv.n) {
    const cA = hC.indexOf("Audiencia");
    if (cA < 0) {
      add("problema", "audiencias", 'Falta la columna "Audiencia" en Conversaciones',
          "Se crea sola en la próxima ingesta.", null,
          "Vuelve a cargar un CSV en el paso 1 y la columna aparece.");
    } else {
      const vacias = conv.filas.filter(f => !String(f[cA] || "").trim()).length;
      if (vacias)
        add("problema", "audiencias", vacias + " conversaciones sin audiencia escrita",
            "Son filas extraídas antes de que existiera la columna. Los contadores ya las " +
            "resuelven mirando el histórico, pero dejarlo escrito ahorra una lectura en cada " +
            "consulta.", "audiencias");
    }
  }

  // ---- 2b. Filas del histórico evaluadas con la audiencia equivocada ----
  //
  // Este es EL desajuste que costó 191 tickets invisibles. La vista de Tableau
  // tiene filtro de fecha relativo, así que cada corrida solo reescribe los
  // últimos días: las filas viejas conservan el alcance con el que se
  // evaluaron. Los tickets de aeropuerto cargados antes de que existieran las
  // audiencias se juzgaron con el alcance de B2B —que exige viaje— y los que
  // no tenían journey quedaron en Auditable = NO para siempre.
  //
  // Se detecta recalculando el alcance en memoria y comparándolo con lo escrito.
  const hist = conHistoricoEnMemoria_(ss);
  if (hist.n) {
    const hH = hist.h;
    const iH = x => hH.indexOf(x);
    const nF = hist.n;
    const v = (r, k) => k >= 0 ? r[k] : "";
    if (iH("Audiencia") >= 0 && iH("Auditable") >= 0) {
      const kGr = iH("Group Name"), kJo = iH("Journey Id"), kAg = iH("Assignee FullName"),
            kAu = iH("Abi/Auto Answer"), kT1 = iH("ES Output Tags 1st Level v2"),
            kT2 = iH("ES Output Tags 2nd Level v2"), kT3 = iH("ES Output Tags 3rd Level v2"),
            kAd = iH("Auditable"), kAc = iH("Audiencia");
      let desajustadas = 0, ganarian = 0;
      for (let k = 0; k < nF; k++) {
        const f = hist.filas[k];
        const t = { grupo: String(v(f, kGr) || ""), journey: String(v(f, kJo) || ""),
                    agent: String(v(f, kAg) || ""), auto: String(v(f, kAu) || ""),
                    tag1: String(v(f, kT1) || ""), tag2: String(v(f, kT2) || ""),
                    tag3: String(v(f, kT3) || ""), audiencia: String(v(f, kAc) || "").trim() };
        histEvaluarAlcance_(t);
        const antes = audAlcanceHistoricoOk_(v(f, kAd)) ? "SI" : "NO";
        if (t.auditable !== antes) {
          desajustadas++;
          if (t.auditable === "SI") ganarian++;
        }
      }
      if (desajustadas)
        add("problema", "audienciasHist",
            desajustadas + " filas del histórico tienen el alcance mal calculado",
            "Se evaluaron con la audiencia equivocada: la vista de Tableau solo reescribe los " +
            "últimos días, así que las filas viejas conservan el alcance con el que entraron. " +
            (ganarian ? ganarian + " tickets pasarían a auditables y hoy no aparecen en el paso 2."
                      : "Ninguno gana alcance, pero conviene dejarlo coherente."),
            "audienciasHist");
    }
  }

  // ---- 2c. Celdas "Auditable" de Conversaciones desactualizadas ----
  //
  // El gemelo del desajuste del histórico, y el que de verdad dejaba la tabla
  // en cero: esa celda se escribe UNA vez, al ingerir, con lo que decía el
  // histórico entonces. Corregir el histórico no la toca.
  if (conv.n) {
    const cAu = hC.indexOf("Auditable"), cTk2 = hC.indexOf("Ticket Number");
    if (cAu >= 0 && cTk2 >= 0) {
      const nC = conv.n;
      const vivos = {};
      conAuditables_(ss, "").forEach(t => { vivos[t] = true; });
      let bloqueados = 0;
      const ejemplos = [];
      for (let k = 0; k < nC; k++) {
        const f = conv.filas[k];
        const t = zdClaveTicket_(f[cTk2]);
        if (!t) continue;
        const v = String(f[cAu] || "").trim();
        // La celda lo rechaza por algo que NO es de la conversación, y el
        // histórico hoy dice que sí es auditable: valor viejo.
        if (v.toUpperCase().indexOf("NO") === 0 && vivos[t] && audConversacionAuditable_(v)) {
          bloqueados++;
          if (ejemplos.length < 5) ejemplos.push(t);
        }
      }
      if (bloqueados)
        add("problema", "auditableConver",
            bloqueados + " conversaciones bloqueadas por un valor viejo",
            'Su celda "Auditable" las rechaza por un motivo que ya no aplica, mientras el ' +
            "histórico dice que sí se pueden auditar. No aparecen en el paso 2 y no es por la " +
            "conversación. Ejemplos: " + ejemplos.join(", "), "auditableConver");
    }
  }

  // ---- 3. Prompts de una versión anterior ----
  try {
    const est = audEstadoPrompts_();
    const pend = est.viejos + est.desconocidos;
    if (pend)
      add("info", "prompts", pend + " prompts son de una versión anterior de las reglas",
          "No hace falta tocar nada: cada uno se regenera solo la primera vez que se audita " +
          "ese ticket. Este botón solo adelanta el trabajo.", "prompts");
  } catch (e) { /* sin hoja todavía */ }

  // ---- 4. Confluence ----
  const nProc = filas((typeof CONF !== "undefined" && CONF.HOJA) || "Procedimientos");
  if (typeof confSiempreDe_ === "function") {
    if (nProc <= 0) {
      add("problema", "procedimientos", "No hay procedimientos sincronizados",
          "Los prompts salen sin el procedimiento vigente, así que los criterios de proceso " +
          "se juzgan contra la intuición del modelo.", "procedimientos");
    } else {
      let hojaProc = {};
      try { hojaProc = confLeerHoja_(); } catch (e) {}
      const faltan = [];
      audAudiencias_().forEach(a => {
        confSiempreDe_(a).forEach(id => { if (!hojaProc[String(id)]) faltan.push(a + ":" + id); });
      });
      if (faltan.length)
        add("problema", "procedimientos", faltan.length + " páginas fijas sin descargar",
            "Están referenciadas pero no en la hoja: " + faltan.slice(0, 6).join(", ") +
            (faltan.length > 6 ? "…" : ""), "procedimientos");
    }
  }

  // ---- 5. Catálogo de tags ----
  if (typeof tagsLeer_ === "function") {
    const d = tagsLeer_();
    if (!d.filas.length)
      add("problema", "tags", "El catálogo de output tags no está sincronizado",
          "C23 es crítico y sin catálogo se evalúa sin conocer las tags válidas.", "tags");
  }

  // ---- 6. Conversaciones truncadas: hay que volver a extraerlas ----
  //
  // Se lee la columna BANDERAS, que es ligera. La primera version de esto leia
  // la columna de la conversacion para buscar la marca TRUNCADO al final del
  // texto: con 300 filas de 45.000 caracteres son 13 MB y el panel se quedaba
  // colgado en "Revisando el estado del sistema…" sin dar error. El mismo
  // error que ya habia costado colgar la pestaña de revision, repetido.
  if (conv.n) {
    const cBand = hC.indexOf("Banderas"), cTk = hC.indexOf("Ticket Number");
    if (cBand >= 0 && cTk >= 0) {
      const truncados = [];
      conv.filas.forEach(f => {
        if (String(f[cBand] || "").indexOf("conversacion_truncada") >= 0)
          truncados.push(zdClaveTicket_(f[cTk]));
      });
      if (truncados.length)
        add("decision", "truncadas", truncados.length + " conversaciones llegaron cortadas",
            "Su texto excede el tope de una celda. El prompt se armó con lo que había, y las " +
            "señales del hilo se reconstruyen a medias: el cierre puede quedar en DESCONOCIDO.",
            null,
            "Vuelve a extraerlas en el paso 1 pidiendo esos tickets solos, para que quepan.",
            truncados);
    }
  }

  // ---- 7. Auditorías hechas con reglas anteriores ----
  const shV = hoja(AUD.HOJA_VEREDICTOS);
  if (shV && shV.getLastRow() > 1) {
    const hV = cabecera(AUD.HOJA_VEREDICTOS);
    const cR = hV.indexOf("Version reglas"), cT = hV.indexOf("Ticket Number");
    if (cT >= 0) {
      const n = shV.getLastRow() - 1;
      const ids = shV.getRange(2, cT + 1, n, 1).getValues();
      const vers = cR >= 0 ? shV.getRange(2, cR + 1, n, 1).getValues() : null;
      const viejos = {};
      for (let i = 0; i < n; i++) {
        const v = vers ? String(vers[i][0] || "").trim() : "";
        if (v === AUD_VERSION_REGLAS) continue;
        const t = zdClaveTicket_(ids[i][0]);
        if (t) viejos[t] = true;
      }
      const lista = Object.keys(viejos);
      if (lista.length)
        add("decision", "reauditar", lista.length + " auditorías se hicieron con reglas anteriores",
            "Las reglas actuales son [" + AUD_VERSION_REGLAS + "]. Estas auditorías salieron " +
            "con una versión previa, así que pueden arrastrar falsos positivos ya corregidos.",
            null,
            "Rehacerlas BORRA los veredictos actuales, así que la decisión es tuya. En el paso 3, " +
            '"Borrar todas las auditorías" las devuelve a la cola; o borra las de un ticket suelto ' +
            "desde su ficha.",
            lista);
    }
  }

  // ---- 8. El modo automático ----
  //
  // Un modo automático que se ve encendido y no avanza es peor que uno
  // apagado: nadie va a mirar el paso 2 porque cree que está trabajando. Así
  // que acá se mira lo que puede hacer que no avance, no si está encendido.
  if (typeof autoEstado === "function") {
    let a = null;
    try { a = autoEstado(); } catch (err) { a = null; }
    if (a) {
      if (a.descuadre)
        add("problema", "autoDescuadre", "El modo automático está descuadrado", a.descuadre,
            a.activo ? null : "autoApagar");

      if (a.apartados.length)
        add("decision", "autoApartados",
            a.apartados.length + " tickets quedaron apartados por el modo automático",
            "Fallaron " + a.maxIntentos + " veces seguidas, así que salieron de la cola para no " +
            "bloquearla. El motivo de cada uno está en la hoja \"" + a.hoja + "\".",
            "autoReintentar",
            "Antes de reintentarlos vale la pena mirar el error: si es el mismo, volverán a " +
            "apartarse.",
            a.apartados.map(x => x.ticket));

      // Silencio en un modo encendido: o terminó la cola, o el disparador
      // dejó de dispararse. Son dos cosas distintas y se dicen distinto.
      if (a.activo && a.ultimaTanda) {
        const minutos = a.minutosDesdeUltima;
        if (minutos !== null && minutos > Math.max(30, (a.minutos || 5) * 4))
          add("problema", "autoDormido",
              "El modo automático no despierta desde hace " + minutos + " minutos",
              "Debería correr cada " + (a.minutos || 5) + " minutos. Suele ser cuota de " +
              "disparadores agotada por hoy.", "autoApagar",
              "Apágalo y vuelve a encenderlo mañana, o sigue a mano en el paso 2.");
        else if (a.atascado)
          // Encendido y sin cola SUENA a que terminó. Con todo apartado no
          // terminó nada: hay que mirar los errores o no se audita más.
          add("problema", "autoAtascado", "El modo automático está encendido pero atascado",
              "No le queda nada que intentar: los pendientes que había están apartados por " +
              "errores repetidos. Seguirá despertando sin auditar nada hasta que se mire qué " +
              "les pasa.", "autoReintentar",
              'Mira el motivo en la hoja "' + a.hoja + '" antes de devolverlos a la cola.');
        else if (a.enEspera)
          add("info", "autoEspera", "El modo automático está encendido y sin cola",
              "Auditó todo lo pendiente. Sigue despierto: si el paso 1 trae tickets nuevos, " +
              "los toma solo.");
        else
          add("info", "autoTrabajando", "El modo automático está trabajando",
              a.auditados + " auditados en " + a.tandas + " tandas" +
              (a.fallidos ? ", " + a.fallidos + " con error" : "") +
              (a.mediaSegundos ? " · " + a.mediaSegundos + " s por auditoría" : "") + ".");
      } else if (!a.activo && a.apagadoPor) {
        add("info", "autoApagado", "El modo automático está apagado", a.apagadoPor);
      }
    }
  }

  // ---- 8. Pendientes de extracción: el flujo normal ----
  const e = conEstado(audiencia);
  if (e.pendExtraccion)
    add("info", "extraer", e.pendExtraccion + " tickets auditables sin extraer",
        "Es el flujo normal: se extraen en el paso 1.", null,
        "Ve al paso 1 y copia la lista en el extractor.");

  const problemas = items.filter(x => x.nivel === "problema").length;
  const decisiones = items.filter(x => x.nivel === "decision").length;
  if (!items.length)
    add("info", "ok", "Todo en orden", "No hay nada pendiente de mantenimiento.");

  return {
    items: items,
    problemas: problemas,
    decisiones: decisiones,
    pendientes: problemas + decisiones,
    audiencia: audiencia || "",
    actualizado: Utilities.formatDate(new Date(), CON.TZ, "HH:mm"),
  };
}

// ============================================================
//  INSTALACIÓN
// ============================================================

/**
 * Ejecuta esto una vez y lee el registro: te dice qué falta.
 */
function conInstrucciones() {
  const ss = histGetSpreadsheet_(true);
  const faltan = [];
  [["Historico", "corre histActualizar()"],
   [ZD.HOJA_CONVER, "se crea sola al cargar el primer CSV"],
   [AUD.HOJA_VEREDICTOS, "se crea sola al guardar el primer veredicto"],
   [REV.HOJA_REVISIONES, "corre revSetup()"]].forEach(p => {
    if (!ss.getSheetByName(p[0])) faltan.push("  · falta la hoja " + p[0] + " — " + p[1]);
  });

  Logger.log([
    "CONSOLA DE AUDITORÍA",
    "",
    "Hoja: " + ss.getUrl(),
    "",
    faltan.length ? "Pendientes:\n" + faltan.join("\n") : "Todas las hojas están.",
    "",
    "Para publicar: Implementar ▸ Nueva implementación ▸ Aplicación web.",
    "  Ejecutar como: yo.   Con acceso: cualquier usuario de la organización.",
  ].join("\n"));
  return faltan;
}
