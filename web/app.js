/**
 * DietaDash — dashboard locale del registro alimentare.
 *
 * Il database e' Dieta_data.xlsx: la pagina non tiene una copia propria dei
 * dati. Ogni salvataggio va sul workbook tramite il server locale, e ogni
 * pochi secondi si controlla il timbro di modifica del file: se l'Excel viene
 * cambiato a mano, la pagina se ne accorge e si ridisegna. Il legame e' a
 * doppio mandato, e in mezzo non c'e' nessuna cache da tenere allineata.
 */

'use strict'

// ---------------------------------------------------------------- costanti

/** Alimenti che nel registro segnano uno sgarro: esclusi dalle statistiche. */
const CIBI_SGARRO = new Set(['Pizza', 'Sushi', 'Ristorante', 'Kebab', 'Curdo', 'Risotto', 'Piadina'])

const GIORNI_BREVI = ['Lun', 'Mar', 'Mer', 'Gio', 'Ven', 'Sab', 'Dom']
const MESI = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
  'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre']

/** Timbro di versione della copia pubblicata (vedi index.html). In locale
 *  vale "sviluppo" e non serve, perche' il server manda gia' no-store. */
const VERSIONE = document.querySelector('meta[name="versione"]')?.content || 'sviluppo'

const PAGINE = [
  { id: 'panoramica', etichetta: 'Panoramica', icona: '📊' },
  { id: 'calendario', etichetta: 'Calendario', icona: '📅' },
  { id: 'inserisci', etichetta: 'Inserisci', icona: '✍️' },
  { id: 'registro', etichetta: 'Registro', icona: '🗂️' },
]

const CAMPI_TESTO = [
  ['colazione', 'Colazione'],
  ['pranzoCarbo', 'Pranzo · carboidrati'],
  ['pranzoProt', 'Pranzo · proteine'],
  ['cenaCarbo', 'Cena · carboidrati'],
  ['cenaProt', 'Cena · proteine'],
  ['merendaMattina', 'Merenda mattina'],
  ['merendaPomeriggio', 'Merenda pomeriggio'],
]

// ------------------------------------------------------------------- stato

const stato = {
  giorni: [],
  misure: [],
  config: {},
  revisione: 0,
  collegato: null, // null = non si sa ancora
  solaLettura: false, // istantanea senza server (es. pubblicata online)
  pagina: leggiHash(),
  bozza: null, // giorno in modifica nella pagina Inserisci
  filtroRegistro: '',
  evidenzia: null, // giorno da mostrare nel registro, arrivando dal calendario
}

// -------------------------------------------------------------- utilitari

const $ = (sel, dove = document) => dove.querySelector(sel)
const el = (tag, attr = {}, ...figli) => {
  const n = document.createElement(tag)
  for (const [k, v] of Object.entries(attr)) {
    if (v === null || v === undefined || v === false) continue
    if (k === 'class') n.className = v
    else if (k === 'html') n.innerHTML = v
    else if (k.startsWith('on')) n.addEventListener(k.slice(2).toLowerCase(), v)
    else n.setAttribute(k, v === true ? '' : String(v))
  }
  for (const f of figli.flat()) {
    if (f === null || f === undefined || f === false) continue
    n.append(f instanceof Node ? f : document.createTextNode(String(f)))
  }
  return n
}

const oggiISO = () => {
  const d = new Date()
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}
/** Le date sono stringhe AAAA-MM-GG: si costruiscono a mezzogiorno per non
 *  farsi spostare il giorno dal fuso orario. */
const aData = (iso) => new Date(`${iso}T12:00:00`)
const dataBreve = (iso) => {
  const d = aData(iso)
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getFullYear()).slice(2)}`
}
const dataLunga = (iso) => {
  const d = aData(iso)
  return `${d.getDate()} ${MESI[d.getMonth()].toLowerCase()} ${d.getFullYear()}`
}
const giorniFra = (a, b) => Math.round((aData(b) - aData(a)) / 86400000)
const isoDi = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
/** La stessa data n mesi prima, in formato AAAA-MM-GG. */
const mesiPrima = (iso, n) => {
  const d = aData(iso)
  d.setMonth(d.getMonth() - n)
  return isoDi(d)
}
const num = (v, dec = 1) => (v === null || v === undefined || Number.isNaN(v) ? '—' : v.toFixed(dec))
const perc = (v) => (v === null || v === undefined || Number.isNaN(v) ? '—' : `${Math.round(v * 100)}%`)

/** Frase alla maniera del workbook: 2.9 -> "quasi tre volte a settimana". */
function aParole(v) {
  if (v === null || v === undefined || Number.isNaN(v)) return ''
  const nomi = ['zero', 'una', 'due', 'tre', 'quattro', 'cinque', 'sei', 'sette']
  const vicino = Math.round(v)
  if (vicino > 7) return 'più di sette volte a settimana'
  const scarto = v - vicino
  const volte = vicino === 1 ? 'volta' : 'volte'
  const nome = nomi[Math.max(0, vicino)]
  if (Math.abs(scarto) <= 0.12) return `${nome} ${volte} a settimana`
  return scarto > 0 ? `poco più di ${nome} ${volte} a settimana` : `poco meno di ${nome} ${volte} a settimana`
}

function brindisi(testo, tipo = 'ok') {
  const box = $('#brindisi')
  const n = el('div', { class: tipo }, testo)
  box.append(n)
  setTimeout(() => {
    n.style.transition = 'opacity .25s'
    n.style.opacity = '0'
    setTimeout(() => n.remove(), 250)
  }, tipo === 'ko' ? 6000 : 2800)
}

function leggiHash() {
  const h = location.hash.replace('#', '')
  return PAGINE.some((p) => p.id === h) ? h : 'panoramica'
}

// ---------------------------------------------------------------- statistiche

/** Data dell'ultimo check: l'ultimo giorno la cui nota contiene "check". */
function ultimoCheck(giorni) {
  const c = giorni.filter((g) => /check/i.test(g.note)).map((g) => g.data)
  return c.length ? c.sort()[c.length - 1] : null
}

function riepiloga(giorni) {
  const n = giorni.length
  const settimane = n / 7
  const somma = (campo) => giorni.reduce((s, g) => s + (g[campo] || 0), 0)
  const allenamenti = somma('allenamento')
  const passi = somma('passi')
  const sgarri = somma('sgarro')
  return {
    giorni: n,
    settimane,
    allenamenti,
    passi,
    sgarri,
    allenamentiSett: n ? allenamenti / settimane : null,
    passiSett: n ? passi / settimane : null,
    sgarriSett: n ? sgarri / settimane : null,
    aderenza: n ? (n - sgarri) / n : null,
    dal: n ? giorni[0].data : null,
    al: n ? giorni[n - 1].data : null,
  }
}

/**
 * Quante volte compare ogni alimento nei campi indicati.
 *
 * Si esclude la singola portata da sgarro, non l'intera giornata: se la cena
 * e' stata una pizza, il riso del pranzo resta un pasto della dieta e va
 * contato. Le celle vuote sono i pasti sostituiti dallo sgarro.
 */
function distribuzione(giorni, campi) {
  const conto = new Map()
  let portate = 0
  for (const g of giorni) {
    for (const campo of campi) {
      const v = g[campo]
      if (!v || CIBI_SGARRO.has(v)) continue
      conto.set(v, (conto.get(v) || 0) + 1)
      portate++
    }
  }
  const voci = [...conto.entries()].map(([nome, n]) => ({ nome, n })).sort((a, b) => b.n - a.n)
  return { voci, portate, settimane: giorni.length / 7 }
}

/** Serie settimanale di allenamenti e passi, per il grafico di andamento. */
function serieSettimanali(giorni) {
  const per = new Map()
  for (const g of giorni) {
    if (!per.has(g.settimana)) per.set(g.settimana, { settimana: g.settimana, inizio: g.data, allenamento: 0, passi: 0, giorni: 0 })
    const s = per.get(g.settimana)
    s.allenamento += g.allenamento
    s.passi += g.passi
    s.giorni++
    if (g.data < s.inizio) s.inizio = g.data
  }
  return [...per.values()].sort((a, b) => a.settimana - b.settimana)
}

// ------------------------------------------------------------------- rete

async function chiedi(url, opzioni) {
  const r = await fetch(url, opzioni)
  const corpo = await r.json().catch(() => null)
  if (!r.ok) throw new Error(corpo?.errore || `Richiesta fallita (${r.status}).`)
  return corpo
}

function assorbi(dati) {
  stato.giorni = dati.giorni
  stato.misure = dati.misure || []
  stato.config = dati.config || {}
  stato.revisione = dati.revisione
  stato.collegato = true
}

let battito = null

/**
 * Prima si prova il server locale. Se non c'e' — la dashboard aperta online,
 * dove il server non puo' esistere — si ripiega sull'istantanea `dati.json`
 * committata accanto alla pagina: si guardano i numeri, non si scrive.
 */
async function carica() {
  try {
    assorbi(await chiedi('api/dati'))
    stato.solaLettura = false
  } catch (e) {
    try {
      assorbi(await chiedi(`dati.json?v=${encodeURIComponent(VERSIONE)}`))
      stato.collegato = false
      stato.solaLettura = true
      // Online il server non arrivera' mai: continuare a interrogarlo sarebbe
      // solo una richiesta fallita ogni tre secondi.
      if (battito) clearInterval(battito)
    } catch {
      stato.collegato = false
      stato.erroreCollegamento = e.message
    }
  }
  disegna()
}

async function salvaGiorno(giorno) {
  const esito = await chiedi('api/giorni', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ giorni: [giorno] }),
  })
  assorbi(esito.dati)
  return esito
}

async function eliminaGiorno(data) {
  const esito = await chiedi(`api/giorno?data=${encodeURIComponent(data)}`, { method: 'DELETE' })
  assorbi(esito.dati)
  return esito
}

/**
 * Polling del timbro di modifica: e' il verso Excel -> pagina. Costa una
 * chiamata da poche decine di byte, quindi si puo' fare spesso; il registro
 * intero si rilegge solo quando il file e' davvero cambiato.
 */
async function sorveglia() {
  try {
    const r = await chiedi('api/revisione')
    if (stato.collegato === false) {
      stato.collegato = true
      await carica()
      brindisi('Collegamento al workbook ripristinato.', 'ok')
      return
    }
    if (r.revisione && r.revisione !== stato.revisione) {
      assorbi(await chiedi('api/dati'))
      disegna()
      brindisi("Excel modificato: dashboard aggiornata.", 'info')
    }
  } catch {
    // Solo al passaggio da collegato a scollegato: ridisegnare a ogni battito
    // farebbe sfarfallare la pagina finche' il server e' spento.
    if (stato.collegato !== false) {
      stato.collegato = false
      disegna()
      brindisi('Server locale non raggiungibile: per ora non posso scrivere sull\'Excel.', 'ko')
    }
  }
}

// ------------------------------------------------------------- componenti

function statCard({ etichetta, valore, sotto, tono = '', accento = false, variazione = null }) {
  return el('div', { class: `card stat${accento ? ' acc' : ''}` },
    el('div', { class: 'etichetta' }, etichetta),
    el('div', { class: 'riga-valore' },
      el('div', { class: `valore ${tono}` }, valore),
      variazione,
    ),
    sotto ? el('div', { class: 'sotto', html: sotto }) : null,
  )
}

/**
 * Quanto e' cambiato il periodo dall'ultimo check rispetto alla media di
 * sempre. Il segno da solo non basta a dire se e' una buona notizia: piu'
 * allenamenti e' un miglioramento, piu' sgarri no. Per questo il verso
 * "buono" arriva da fuori invece di essere dedotto dal segno.
 */
function distintivo(recente, generale, versoBuono = +1) {
  if (recente === null || generale === null || !generale) return null
  const scarto = (recente - generale) / generale
  if (Math.abs(scarto) < 0.005) {
    return el('span', { class: 'variazione pari', title: 'Come la media di sempre.' }, '=')
  }
  const buono = Math.sign(scarto) === Math.sign(versoBuono)
  const segno = scarto > 0 ? '+' : '−'
  return el('span', {
    class: `variazione ${buono ? 'su' : 'giu'}`,
    title: `${(scarto * 100).toFixed(1)}% rispetto alla media dall'inizio della dieta.`,
  }, `${segno}${Math.abs(Math.round(scarto * 100))}%`)
}

function sezione(titolo, suggerimento, destra, ...corpo) {
  return el('section', { class: 'card sezione' },
    el('header', {},
      el('div', {}, el('h2', {}, titolo), suggerimento ? el('p', {}, suggerimento) : null),
      destra,
    ),
    el('div', { class: 'corpo' }, corpo),
  )
}

/**
 * Andamento settimanale: due linee su una griglia 0-7, perche' allenamenti e
 * passi si contano entrambi in giorni su sette. Le settimane parziali (la
 * prima e l'ultima) restano fuori: mezza settimana disegnerebbe un crollo che
 * non c'e' stato.
 */
function graficoAndamento(serie) {
  const piene = serie.filter((s) => s.giorni === 7)
  if (piene.length < 2) return el('div', { class: 'vuoto' }, 'Servono almeno due settimane complete.')

  const NS = 'http://www.w3.org/2000/svg'
  const L = 34, R = 12, T = 12, B = 30, W = 920, H = 250
  const larg = W - L - R, alt = H - T - B
  const passo = larg / (piene.length - 1)
  const x = (i) => L + i * passo
  const y = (v) => T + (1 - v / 7) * alt

  const svg = document.createElementNS(NS, 'svg')
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`)
  svg.setAttribute('class', 'grafico')
  svg.setAttribute('role', 'img')
  svg.setAttribute('aria-label', 'Allenamenti e giorni di passi per settimana')
  const crea = (tag, attr, testo) => {
    const n = document.createElementNS(NS, tag)
    for (const [k, v] of Object.entries(attr)) n.setAttribute(k, String(v))
    if (testo !== undefined) n.textContent = testo
    return n
  }

  for (let v = 0; v <= 7; v++) {
    svg.append(crea('line', { x1: L, x2: W - R, y1: y(v), y2: y(v), stroke: '#262c38', 'stroke-width': 1 }))
    if (v % 1 === 0) svg.append(crea('text', { x: L - 7, y: y(v) + 4, 'text-anchor': 'end', 'font-size': 10, fill: '#7c8798' }, String(v)))
  }

  const salto = Math.max(1, Math.ceil(piene.length / 12))
  piene.forEach((s, i) => {
    if (i % salto) return
    svg.append(crea('text', { x: x(i), y: H - 10, 'text-anchor': 'middle', 'font-size': 10, fill: '#7c8798' }, dataBreve(s.inizio).slice(0, 5)))
  })

  const linea = (campo, colore) => {
    const punti = piene.map((s, i) => `${x(i)},${y(s[campo])}`)
    svg.append(crea('polyline', {
      points: punti.join(' '), fill: 'none', stroke: colore,
      'stroke-width': 2.2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round',
    }))
    piene.forEach((s, i) => {
      const c = crea('circle', { cx: x(i), cy: y(s[campo]), r: 3, fill: colore })
      c.append(crea('title', {}, `Settimana ${s.settimana} (dal ${dataBreve(s.inizio)}): ${s[campo]}`))
      svg.append(c)
    })
  }
  linea('allenamento', '#f97316')
  linea('passi', '#38bdf8')

  return el('div', {},
    svg,
    el('div', { class: 'legenda', style: 'margin-top:8px;justify-content:center' },
      el('span', { html: "<i style='background:#f97316'></i>Allenamenti" }),
      el('span', { html: "<i style='background:#38bdf8'></i>Giorni di passi" }),
    ),
  )
}

/**
 * Confronto fra la frequenza di sempre e quella dall'ultimo check: due barre
 * per alimento, sulla stessa scala, cosi' si vede a colpo d'occhio cos'e'
 * cambiato dopo il check invece di dover leggere due elenchi separati.
 */
function barreConfronto(totale, dalCheck) {
  if (!totale.voci.length) return el('div', { class: 'vuoto' }, 'Nessun dato.')
  const perSett = (v, s) => (s ? v / s : 0)
  const dalCheckMappa = new Map(dalCheck.voci.map((v) => [v.nome, v.n]))
  const righe = totale.voci.map((v) => ({
    nome: v.nome,
    tot: perSett(v.n, totale.settimane),
    chk: perSett(dalCheckMappa.get(v.nome) || 0, dalCheck.settimane),
  }))
  const massimo = Math.max(...righe.flatMap((r) => [r.tot, r.chk]), 0.5)

  return el('div', {},
    ...righe.map((r) =>
      el('div', { class: 'barra-riga' },
        el('div', { class: 'nome' }, r.nome),
        el('div', { class: 'barra-pista' },
          el('div', { class: 'traccia', title: `Dall'inizio: ${aParole(r.tot)}` },
            el('div', {}, el('div', { class: 'barra tot', style: `width:${Math.max(1, (r.tot / massimo) * 100)}%` })),
            el('b', {}, `${r.tot.toFixed(1)}×`)),
          el('div', { class: 'traccia', title: `Dall'ultimo check: ${aParole(r.chk)}` },
            el('div', {}, el('div', { class: 'barra chk', style: `width:${Math.max(1, (r.chk / massimo) * 100)}%` })),
            el('b', {}, `${r.chk.toFixed(1)}×`)),
        ),
      ),
    ),
    el('div', { class: 'legenda', style: 'margin-top:10px' },
      el('span', { html: "<i style='background:#f97316'></i>Dall'inizio" }),
      el('span', { html: "<i style='background:#38404f'></i>Dall'ultimo check" }),
      el('span', {}, '· volte a settimana'),
    ),
  )
}

// ------------------------------------------------------------ pagina: panoramica

function vistaPanoramica() {
  const giorni = stato.giorni
  if (!giorni.length) return el('div', { class: 'vuoto' }, 'Registro vuoto.')

  const check = ultimoCheck(giorni)
  const dalCheck = check ? giorni.filter((g) => g.data >= check) : []
  const tot = riepiloga(giorni)
  const chk = riepiloga(dalCheck)

  // Il grafico guarda solo il periodo recente: tre mesi bastano a vedere un
  // andamento, e un anno intero di settimane schiaccerebbe le differenze.
  const taglio = mesiPrima(tot.al, 3)
  const serieRecenti = serieSettimanali(giorni.filter((g) => g.data >= taglio))

  const prossimo = stato.config.prossimo_check || ''
  const mancano = prossimo ? giorniFra(oggiISO(), prossimo) : null

  // Il numero grande e' il periodo dall'ultimo check: e' quello su cui puoi
  // ancora intervenire. La media di sempre resta sotto, come metro di paragone.
  const generale = (v) => `generale <b>${v}</b>`

  const carte = el('div', { class: 'griglia-stat' },
    statCard({
      accento: true,
      etichetta: 'Prossimo check',
      valore: prossimo
        ? (mancano > 0 ? `fra ${mancano} ${mancano === 1 ? 'giorno' : 'giorni'}` : mancano === 0 ? 'oggi' : `${-mancano} gg fa`)
        : '—',
      sotto: prossimo ? dataLunga(prossimo) : 'nessuna data impostata',
    }),
    statCard({
      etichetta: 'Allenamenti a settimana',
      valore: num(chk.allenamentiSett),
      variazione: distintivo(chk.allenamentiSett, tot.allenamentiSett, +1),
      sotto: `${generale(num(tot.allenamentiSett))} · ${aParole(chk.allenamentiSett)}`,
    }),
    statCard({
      etichetta: 'Passi a settimana',
      valore: num(chk.passiSett),
      variazione: distintivo(chk.passiSett, tot.passiSett, +1),
      sotto: `${generale(num(tot.passiSett))} · ${aParole(chk.passiSett)}`,
    }),
    statCard({
      etichetta: 'Aderenza',
      valore: perc(chk.aderenza),
      tono: chk.aderenza >= 0.85 ? 'buono' : chk.aderenza >= 0.75 ? 'attenzione' : 'cattivo',
      variazione: distintivo(chk.aderenza, tot.aderenza, +1),
      sotto: `${generale(perc(tot.aderenza))} · ${chk.sgarri} sgarri su ${chk.giorni} giorni`,
    }),
    statCard({
      etichetta: 'Sgarri a settimana',
      valore: num(chk.sgarriSett),
      // Qui piu' non e' meglio: il verso buono e' verso il basso.
      variazione: distintivo(chk.sgarriSett, tot.sgarriSett, -1),
      sotto: generale(num(tot.sgarriSett)),
    }),
    statCard({
      etichetta: 'Giorni registrati',
      valore: String(chk.giorni),
      sotto: `dall'ultimo check · in tutto <b>${tot.giorni}</b> dal ${dataBreve(tot.dal)}`,
    }),
  )

  const intestazioneCheck = check
    ? `Ultimo check: ${dataLunga(check)} · ${chk.giorni} giorni (${chk.settimane.toFixed(1)} settimane)`
    : 'Nessun check registrato: scrivi "Check" nelle note di un giorno per marcarlo.'

  return el('div', { style: 'display:flex;flex-direction:column;gap:12px' },
    carte,
    sezione('Andamento settimanale',
      `Ultimi 3 mesi, dal ${dataBreve(taglio)}. Solo settimane complete: ogni linea conta i giorni su sette.`,
      null, graficoAndamento(serieRecenti)),
    el('div', { class: 'due-colonne' },
      sezione('Carboidrati', 'Pranzo e cena, portate da sgarro escluse.', null,
        barreConfronto(distribuzione(giorni, ['pranzoCarbo', 'cenaCarbo']), distribuzione(dalCheck, ['pranzoCarbo', 'cenaCarbo']))),
      sezione('Proteine', 'Pranzo e cena, portate da sgarro escluse.', null,
        barreConfronto(distribuzione(giorni, ['pranzoProt', 'cenaProt']), distribuzione(dalCheck, ['pranzoProt', 'cenaProt']))),
    ),
    sezione('Periodo di confronto', intestazioneCheck, null,
      el('p', { class: 'nota-piccola' },
        'I numeri grandi sono quelli dal giorno marcato Check nelle note fino a oggi; sotto, in piccolo, ' +
        "la media dall'inizio della dieta. La percentuale confronta i due: verde quando il periodo recente " +
        'va meglio della media, rossa quando va peggio. La data del prossimo check si cambia dalla pagina ' +
        'Inserisci ed è salvata nel foglio Config del workbook.'),
    ),
  )
}

// ------------------------------------------------------------ pagina: calendario

function vistaCalendario() {
  const giorni = stato.giorni
  if (!giorni.length) return el('div', { class: 'vuoto' }, 'Registro vuoto.')

  const mappa = new Map(giorni.map((g) => [g.data, g]))
  const oggi = oggiISO()
  const prima = aData(giorni[0].data)
  const ultima = aData(giorni[giorni.length - 1].data > oggi ? giorni[giorni.length - 1].data : oggi)

  const mesi = []
  for (let a = prima.getFullYear(), m = prima.getMonth(); a < ultima.getFullYear() || (a === ultima.getFullYear() && m <= ultima.getMonth());) {
    mesi.push([a, m])
    m++
    if (m > 11) { m = 0; a++ }
  }
  // Dal mese piu' recente: e' quello che si guarda, e cosi' e' subito in alto
  // invece che in fondo a un anno di calendari.
  mesi.reverse()

  const disegnaMese = ([anno, mese]) => {
    const inizio = new Date(anno, mese, 1)
    const quanti = new Date(anno, mese + 1, 0).getDate()
    const spazio = (inizio.getDay() + 6) % 7 // la settimana comincia di lunedì
    const celle = []
    for (let i = 0; i < spazio; i++) celle.push(el('div', { class: 'giorno fuori' }))

    let seguiti = 0, registrati = 0
    for (let d = 1; d <= quanti; d++) {
      const iso = `${anno}-${String(mese + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`
      const g = mappa.get(iso)
      let classe = 'assente'
      let titolo = `${dataLunga(iso)} — nessun dato`
      if (g) {
        registrati++
        if (g.sgarro) {
          classe = 'sgarro'
          titolo = `${dataLunga(iso)} — sgarro: ${[g.pranzoCarbo, g.cenaCarbo].filter((v) => CIBI_SGARRO.has(v)).join(', ') || 'segnato'}`
        } else {
          seguiti++
          classe = 'seguita'
          titolo = `${dataLunga(iso)} — dieta seguita\nPranzo: ${g.pranzoCarbo || '—'} + ${g.pranzoProt || '—'}\nCena: ${g.cenaCarbo || '—'} + ${g.cenaProt || '—'}`
        }
        if (g.allenamento) titolo += `\nAllenamento${g.tipoAllenamento ? ` ${g.tipoAllenamento}` : ''}`
        if (g.passi) titolo += '\nPassi ✓'
      } else if (iso > oggi) {
        classe = 'futuro'
        titolo = `${dataLunga(iso)} — ancora da venire`
      }
      const check = g && /check/i.test(g.note)
      celle.push(el('div', {
        class: `giorno ${classe}${check ? ' check' : ''}`,
        title: `${check ? `${titolo}\n★ Check` : titolo}\n\n(clicca per vederlo nel registro)`,
        'data-cliccabile': true,
        // Il registro esiste anche nella copia online, quindi qui non serve
        // la guardia di sola lettura che vale per la pagina Inserisci.
        onclick: () => {
          stato.evidenzia = iso
          stato.filtroRegistro = ''   // un filtro attivo potrebbe nascondere proprio quel giorno
          vai('registro')
        },
      }, String(d)))
    }

    return el('div', { class: 'card', style: 'padding:10px' },
      el('h3', { class: 'mese-titolo' },
        `${MESI[mese]} ${anno}`,
        el('small', {}, registrati ? `${seguiti}/${registrati} in linea` : '—')),
      el('div', { class: 'mese-griglia' },
        ...GIORNI_BREVI.map((g) => el('div', { class: 'dow' }, g[0])),
        ...celle),
    )
  }

  return sezione('Calendario', 'Dal mese più recente. Verde: dieta seguita. Rosso: sgarro. Il bordo arancione segna un check. Clicca un giorno per vederlo nel registro.',
    el('div', { class: 'legenda' },
      el('span', { html: "<i style='background:rgba(63,191,127,.5)'></i>seguita" }),
      el('span', { html: "<i style='background:rgba(244,98,111,.6)'></i>sgarro" }),
      el('span', { html: "<i style='background:#151922'></i>senza dati" }),
    ),
    el('div', { class: 'calendario' }, ...mesi.map(disegnaMese)),
  )
}

// ------------------------------------------------------------ pagina: inserisci

/** Valori gia' usati in un campo, dai piu' frequenti: alimentano i datalist. */
function opzioni(campo) {
  const conto = new Map()
  for (const g of stato.giorni) {
    const v = g[campo]
    if (v) conto.set(v, (conto.get(v) || 0) + 1)
  }
  return [...conto.entries()].sort((a, b) => b[1] - a[1]).map(([v]) => v)
}

/** L'ultimo giorno registrato che non sia uno sgarro: base per un giorno nuovo. */
function modelloGiorno() {
  const buoni = stato.giorni.filter((g) => !g.sgarro)
  const base = buoni[buoni.length - 1] || stato.giorni[stato.giorni.length - 1]
  return base ? { ...base } : null
}

function vistaInserisci() {
  if (stato.solaLettura) {
    return sezione('Inserimento non disponibile',
      'Questa e\' la copia pubblicata online.', null,
      el('p', { class: 'nota-piccola' },
        'Salvare vuol dire scrivere dentro Dieta_data.xlsx, e il workbook sta sul computer, ',
        'non qui. Scarica il progetto e avvialo: da li\' questa pagina scrive e rilegge il file. ',
        'Le istruzioni passo passo sono in INSTALLAZIONE.md.'))
  }

  const data = stato.bozza || oggiISO()
  const esistente = stato.giorni.find((g) => g.data === data)
  const modello = modelloGiorno()
  const g = esistente
    ? { ...esistente }
    : {
        data,
        colazione: modello?.colazione || '',
        pranzoCarbo: '', pranzoProt: '', cenaCarbo: '', cenaProt: '',
        merendaMattina: modello?.merendaMattina || 'Frutto',
        merendaPomeriggio: modello?.merendaPomeriggio || '',
        allenamento: 0, tipoAllenamento: '', passi: 0, sgarro: 0, note: '',
      }

  const modulo = el('form', { class: 'modulo', onsubmit: (e) => e.preventDefault() })

  const campoData = el('input', {
    type: 'date', class: 'campo', value: data, required: true,
    onchange: (e) => { stato.bozza = e.target.value; disegna() },
  })
  modulo.append(el('label', { class: 'gruppo' }, el('span', {}, 'Data'), campoData))

  modulo.append(el('label', { class: 'gruppo' },
    el('span', {}, 'Giorno'),
    el('input', { class: 'campo', value: GIORNI_BREVI[(aData(data).getDay() + 6) % 7], disabled: true }),
  ))

  const riferimenti = {}
  for (const [campo, etichetta] of CAMPI_TESTO) {
    const id = `lista-${campo}`
    const input = el('input', { class: 'campo', list: id, value: g[campo] || '', autocomplete: 'off', placeholder: '—' })
    riferimenti[campo] = input
    modulo.append(el('label', { class: 'gruppo' },
      el('span', {}, etichetta),
      input,
      el('datalist', { id }, ...opzioni(campo).map((v) => el('option', { value: v }))),
    ))
  }

  const spunta = (campo, etichetta) => {
    const input = el('input', { type: 'checkbox', checked: !!g[campo] })
    riferimenti[campo] = input
    return el('label', { class: 'riga-scelta' }, input, etichetta)
  }
  const selTipo = el('select', { class: 'campo' },
    el('option', { value: '' }, '—'),
    ...['U', 'L', 'B'].map((t) => el('option', { value: t, selected: g.tipoAllenamento === t }, t)),
  )
  riferimenti.tipoAllenamento = selTipo

  modulo.append(
    el('div', { class: 'gruppo' }, el('span', {}, 'Allenamento'), spunta('allenamento', 'Allenamento fatto')),
    el('label', { class: 'gruppo' }, el('span', {}, 'Tipo (U / L / B)'), selTipo),
    el('div', { class: 'gruppo' }, el('span', {}, 'Passi'), spunta('passi', 'Passi fatti')),
    el('div', { class: 'gruppo' }, el('span', {}, 'Sgarro'), spunta('sgarro', 'Giornata di sgarro')),
  )

  const campoNote = el('input', { class: 'campo', value: g.note || '', placeholder: 'Check, viaggio, evento…' })
  riferimenti.note = campoNote
  modulo.append(el('label', { class: 'gruppo larga' }, el('span', {}, 'Note'), campoNote))

  // Se scrivi un alimento da sgarro, la spunta si accende da sola: dimenticarla
  // falserebbe l'aderenza e le medie.
  for (const campo of ['pranzoCarbo', 'cenaCarbo']) {
    riferimenti[campo].addEventListener('change', () => {
      const c = ['pranzoCarbo', 'cenaCarbo'].some((f) => CIBI_SGARRO.has(riferimenti[f].value.trim()))
      if (c && !riferimenti.sgarro.checked) {
        riferimenti.sgarro.checked = true
        brindisi('Alimento da sgarro: ho acceso la spunta.', 'info')
      }
    })
  }

  const raccogli = () => {
    const fuori = { data: campoData.value, note: campoNote.value.trim() }
    for (const [campo] of CAMPI_TESTO) fuori[campo] = riferimenti[campo].value.trim()
    fuori.allenamento = riferimenti.allenamento.checked ? 1 : 0
    fuori.passi = riferimenti.passi.checked ? 1 : 0
    fuori.sgarro = riferimenti.sgarro.checked ? 1 : 0
    fuori.tipoAllenamento = selTipo.value
    return fuori
  }

  const bottoneSalva = el('button', { class: 'btn btn-primario', type: 'button' },
    esistente ? 'Aggiorna il giorno' : 'Salva nel workbook')
  bottoneSalva.addEventListener('click', async () => {
    bottoneSalva.disabled = true
    try {
      const esito = await salvaGiorno(raccogli())
      brindisi(esito.aggiunti ? `Giorno aggiunto — backup ${esito.backup}` : `Giorno aggiornato — backup ${esito.backup}`, 'ok')
      stato.bozza = null
      disegna()
    } catch (e) {
      brindisi(e.message, 'ko')
      bottoneSalva.disabled = false
    }
  })

  const azioni = el('div', { style: 'display:flex;gap:8px;flex-wrap:wrap;align-items:center' },
    bottoneSalva,
    el('button', { class: 'btn', type: 'button', onclick: () => { stato.bozza = oggiISO(); disegna() } }, 'Oggi'),
    esistente
      ? el('button', {
          class: 'btn btn-pericolo', type: 'button',
          onclick: async () => {
            if (!confirm(`Elimino il ${dataLunga(data)} dal workbook?`)) return
            try {
              const esito = await eliminaGiorno(data)
              brindisi(`Giorno eliminato — backup ${esito.backup}`, 'ok')
              stato.bozza = null
              disegna()
            } catch (e) { brindisi(e.message, 'ko') }
          },
        }, 'Elimina')
      : null,
    el('span', { class: 'nota-piccola' },
      esistente ? 'Giorno già presente: i campi mostrano quello che c’è nell’Excel.' : 'Giorno nuovo, precompilato con le tue abitudini.'),
  )

  // --- riquadro del prossimo check ---
  const campoCheck = el('input', { type: 'date', class: 'campo', value: stato.config.prossimo_check || '' })
  const salvaCheck = el('button', { class: 'btn', type: 'button' }, 'Salva data')
  salvaCheck.addEventListener('click', async () => {
    salvaCheck.disabled = true
    try {
      const esito = await chiedi('api/config', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ config: { prossimo_check: campoCheck.value } }),
      })
      stato.config = esito.config
      stato.revisione = esito.revisione
      brindisi('Data del prossimo check salvata nel workbook.', 'ok')
      disegna()
    } catch (e) { brindisi(e.message, 'ko'); salvaCheck.disabled = false }
  })

  return el('div', { style: 'display:flex;flex-direction:column;gap:12px' },
    sezione(esistente ? `Modifica ${dataLunga(data)}` : `Nuovo giorno · ${dataLunga(data)}`,
      'Quello che salvi finisce direttamente in Dieta_data.xlsx, con una copia di sicurezza in data/backup.',
      null, modulo, el('div', { style: 'margin-top:12px' }, azioni)),
    sezione('Prossimo check', 'Salvata nel foglio Config del workbook, così resta anche fuori dalla dashboard.',
      null,
      el('div', { style: 'display:flex;gap:8px;flex-wrap:wrap;align-items:center;max-width:420px' },
        campoCheck, salvaCheck),
    ),
  )
}

// ------------------------------------------------------------- pagina: registro

function vistaRegistro() {
  const filtro = stato.filtroRegistro.trim().toLowerCase()
  const righe = [...stato.giorni].reverse().filter((g) =>
    !filtro || Object.values(g).some((v) => String(v).toLowerCase().includes(filtro)))

  const cerca = el('input', {
    class: 'campo', placeholder: 'cerca alimento, nota, data…', value: stato.filtroRegistro,
    style: 'max-width:280px',
    oninput: (e) => { stato.filtroRegistro = e.target.value; disegna() },
  })

  const intestazioni = ['Data', 'Giorno', 'Sett.', 'Colazione', 'Pranzo', 'Cena', 'Merende', 'All.', 'Passi', 'Sgarro', 'Note']
  const tabella = el('table', { class: 'registro' },
    el('thead', {}, el('tr', {}, ...intestazioni.map((t) => el('th', {}, t)))),
    el('tbody', {}, ...righe.map((g) =>
      el('tr', {
        class: [g.sgarro ? 'sgarro' : '', g.data === stato.evidenzia ? 'evidenziato' : ''].filter(Boolean).join(' '),
        'data-giorno': g.data,
        style: stato.solaLettura ? '' : 'cursor:pointer',
        onclick: stato.solaLettura ? null : () => { stato.bozza = g.data; vai('inserisci') },
      },
        el('td', {}, dataBreve(g.data)),
        el('td', {}, g.giorno.slice(0, 3)),
        el('td', { class: 'num' }, String(g.settimana)),
        el('td', {}, g.colazione || '—'),
        el('td', {}, [g.pranzoCarbo, g.pranzoProt].filter(Boolean).join(' + ') || '—'),
        el('td', {}, [g.cenaCarbo, g.cenaProt].filter(Boolean).join(' + ') || '—'),
        el('td', {}, [g.merendaMattina, g.merendaPomeriggio].filter(Boolean).join(' · ') || '—'),
        el('td', { class: 'num' }, g.allenamento
          ? el('span', { class: 'pastiglia si' }, g.tipoAllenamento || '✓')
          : el('span', { class: 'pastiglia no' }, '—')),
        el('td', { class: 'num' }, el('span', { class: `pastiglia ${g.passi ? 'si' : 'no'}` }, g.passi ? '✓' : '—')),
        el('td', { class: 'num' }, el('span', { class: `pastiglia ${g.sgarro ? 'rossa' : 'no'}` }, g.sgarro ? '✓' : '—')),
        el('td', {}, /check/i.test(g.note) ? el('span', { class: 'pastiglia acc' }, g.note) : (g.note || '')),
      ))),
  )

  return sezione(`Registro · ${righe.length} giorni`, 'Clicca una riga per modificarla. È il contenuto del foglio Dieta.',
    cerca, el('div', { class: 'tabella-guscio', style: 'max-height:70vh;overflow-y:auto' }, tabella))
}

// ------------------------------------------------------------------ telaio

function vai(p) {
  if (p !== 'registro') stato.evidenzia = null
  location.hash = p
  stato.pagina = p
  disegna()
}

function disegnaSpia() {
  const spia = $('#spia')
  if (stato.collegato === null) {
    spia.className = 'spia'
    spia.textContent = '…'
    spia.title = 'Sto cercando il workbook…'
  } else if (stato.collegato) {
    spia.className = 'spia viva'
    spia.textContent = 'Excel'
    spia.title = 'Collegato a Dieta_data.xlsx: la pagina scrive e rilegge il workbook.'
  } else if (stato.solaLettura) {
    spia.className = 'spia'
    spia.textContent = 'sola lettura'
    spia.title = "Copia pubblicata: si guarda ma non si scrive, perche' online il workbook non c'e'."
  } else {
    spia.className = 'spia morta'
    spia.textContent = 'scollegato'
    spia.title = stato.erroreCollegamento || 'Server locale non raggiungibile.'
  }
}

function disegnaSchede() {
  const nav = $('#schede')
  const visibili = stato.solaLettura ? PAGINE.filter((p) => p.id !== 'inserisci') : PAGINE
  nav.replaceChildren(...visibili.map((p) =>
    el('button', {
      type: 'button',
      'aria-current': stato.pagina === p.id ? 'page' : null,
      onclick: () => vai(p.id),
    }, `${p.icona} ${p.etichetta}`)))
}

function disegna() {
  disegnaSchede()
  disegnaSpia()

  $('#conteggio').textContent = stato.giorni.length
    ? `${stato.giorni.length} giorni · ${(stato.giorni.length / 7).toFixed(1)} settimane`
    : ''

  const main = $('#pagina')
  main.replaceChildren()

  if (stato.solaLettura) {
    main.append(el('div', { class: 'avviso buono' },
      el('strong', {}, 'Copia di sola lettura. '),
      "Stai guardando un'istantanea del registro: i numeri sono veri, ma da qui non si scrive. ",
      "Per inserire i tuoi giorni serve il workbook sul tuo computer — le istruzioni sono nel file ",
      el('code', {}, 'INSTALLAZIONE.md'), ' del progetto.',
    ))
  } else if (stato.collegato === false) {
    main.append(el('div', { class: 'avviso' },
      el('strong', {}, 'Server locale non raggiungibile. '),
      'La dashboard legge e scrive il workbook attraverso ',
      el('code', {}, 'scripts/server.py'),
      '. Avviala con ', el('code', {}, './avvia.command'),
      ' (oppure ', el('code', {}, 'python3 scripts/server.py'), ') dalla cartella Dieta.',
      stato.erroreCollegamento ? el('div', { class: 'nota-piccola', style: 'margin-top:6px' }, stato.erroreCollegamento) : null,
    ))
    if (!stato.giorni.length) return
  }

  if (stato.collegato === null) {
    main.append(el('div', { class: 'vuoto' }, 'Carico il registro…'))
    return
  }

  main.append(
    stato.pagina === 'panoramica' ? vistaPanoramica()
      : stato.pagina === 'calendario' ? vistaCalendario()
      : stato.pagina === 'inserisci' ? vistaInserisci()
      : vistaRegistro(),
  )

  // La riga puo' essere a centinaia di posizioni di distanza dentro il
  // contenitore che scorre: arrivandoci dal calendario va portata in vista.
  if (stato.pagina === 'registro' && stato.evidenzia) {
    const riga = main.querySelector(`tr[data-giorno="${stato.evidenzia}"]`)
    if (riga) riga.scrollIntoView({ block: 'center' })
  }

  $('#pie').textContent = stato.solaLettura
    ? 'DietaDash — copia pubblicata, di sola lettura. Sul tuo computer la stessa pagina scrive dentro Dieta_data.xlsx.'
    : 'DietaDash — il database è Dieta_data.xlsx: la dashboard ci scrive dentro e si aggiorna da sola quando lo modifichi in Excel. Copie di sicurezza in data/backup.'
}

window.addEventListener('hashchange', () => { stato.pagina = leggiHash(); disegna() })

// I browser rallentano i timer nelle schede in secondo piano (anche a un
// battito al minuto), e la scheda e' proprio in secondo piano mentre stai
// modificando l'Excel: si ricontrolla appena la dashboard torna in primo
// piano, che e' il momento in cui vuoi vedere i dati nuovi.
document.addEventListener('visibilitychange', () => { if (!document.hidden) void sorveglia() })
window.addEventListener('focus', () => void sorveglia())

disegna()
carica()
battito = setInterval(sorveglia, 3000)
