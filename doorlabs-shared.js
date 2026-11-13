// Automatisch gebaut (outlook-addin/build.mjs). Nicht von Hand ändern.
(function (root) {
  "use strict";
  function load(factory) { var module = { exports: {} }; factory(module.exports, module); return module.exports; }
  root.DoorLabsShared = {
  code: load(function (exports, module) {
"use strict";
// DoorLabs-Code in Terminen: `[DoorLabs: meeting | blue | Sitzungszimmer 2]`
//
// Das Outlook-Add-in schreibt den Code ins Feld «Ort» des Termins. Die Kalender-Automatik (App und Edge Function calendar-sync)
// liest ihn aus dem ICS und setzt Status und Farbe für das passende Gerät. Diese Datei ist die einzige Stelle, die das Format kennt:
// Die App, die Edge Function und das Add-in (outlook-addin/build.mjs wandelt sie in einfaches JavaScript um) benutzen sie gemeinsam.
// Reine Funktionen, keine Deno- oder Browser-Aufrufe.
Object.defineProperty(exports, "__esModule", { value: true });
exports.STATUS_PRESET_NAMES = exports.CODE_TEXT_MAX = exports.STATUS_DEFAULT_COLOR = exports.COLOR_HEX = exports.COLOR_KEYS = exports.STATUS_KEYS = void 0;
exports.cleanCodeText = cleanCodeText;
exports.formatCode = formatCode;
exports.parseCode = parseCode;
exports.findCode = findCode;
exports.stripCode = stripCode;
exports.codeMatchesDevice = codeMatchesDevice;
exports.presetForCode = presetForCode;
/** Feste englische Schlüssel, unabhängig von der Sprache der Person, die den Termin erstellt */
exports.STATUS_KEYS = ['meeting', 'call', 'free', 'back', 'focus', 'break'];
/** Dieselben Schlüssel wie SPECTRA6_HEX in src/config/colors.ts */
exports.COLOR_KEYS = ['black', 'white', 'yellow', 'red', 'blue', 'green'];
/**
 * Gespeicherte Werte der 6 Farben (Kopie von SPECTRA6_HEX in src/config/colors.ts, weil Edge Functions nichts aus src/ laden können;
 * ein Test vergleicht beide).
 */
exports.COLOR_HEX = {
    black: '#000000',
    white: '#ffffff',
    yellow: '#ffff00',
    red: '#ff0000',
    blue: '#0000ff',
    green: '#00ff00',
};
/** Farbe, die zu einem Status gehört, wenn der Code keine nennt */
exports.STATUS_DEFAULT_COLOR = {
    meeting: 'blue',
    call: 'red',
    free: 'green',
    back: 'yellow',
    focus: 'black',
    break: 'white',
};
/** Höchste Länge von Status-Text und Gerätename im Code (wie «Ort» auf dem Schild: 24 Zeichen) */
exports.CODE_TEXT_MAX = 24;
/** Altes Format (Version 1 des Add-ins): deutsche Wörter. Wird weiterhin gelesen. */
const LEGACY_STATUS = {
    'im meeting': 'meeting',
    'im call': 'call',
    'frei': 'free',
    'bald zurück': 'back',
    'bald zurueck': 'back',
    'fokus': 'focus',
    'pause': 'break',
};
const LEGACY_COLOR = {
    schwarz: 'black',
    weiss: 'white',
    'weiß': 'white',
    gelb: 'yellow',
    rot: 'red',
    blau: 'blue',
    'grün': 'green',
    gruen: 'green',
};
/**
 * Namen der Standard-Vorlagen in den vier Sprachen (klein geschrieben). Damit findet die Automatik die Vorlage zu einem Status,
 * egal in welcher Sprache sie angelegt wurde. Ein Test vergleicht die Namen mit den Standard-Vorlagen der App.
 */
exports.STATUS_PRESET_NAMES = {
    meeting: ['in a meeting', 'im meeting', 'en réunion', 'in riunione'],
    call: ['in a call', 'im call', 'en appel', 'in chiamata'],
    free: ['free', 'frei', 'libre', 'libero'],
    back: ['away', 'bald zurück', 'de retour bientôt', 'torno subito'],
    focus: ['focus mode', 'fokus', 'concentré', 'concentrato'],
    break: ['break', 'pause', 'en pause', 'in pausa'],
};
/** Entfernt Zeichen, die das Format brechen ([ ] |), Zeilenumbrüche und doppelte Leerzeichen; höchstens CODE_TEXT_MAX Zeichen */
function cleanCodeText(s) {
    return (s !== null && s !== void 0 ? s : '').replace(/[[\]|]/g, '').replace(/\s+/g, ' ').trim().slice(0, exports.CODE_TEXT_MAX).trim();
}
/** Erzeugt den Code. status null = eigener Text (leer wird zu «Busy»). */
function formatCode(input) {
    var _a;
    const status = (_a = input.status) !== null && _a !== void 0 ? _a : (cleanCodeText(input.text) || 'Busy');
    const parts = [status, input.color];
    const device = cleanCodeText(input.device);
    if (device)
        parts.push(device);
    return `[DoorLabs: ${parts.join(' | ')}]`;
}
const CODE_RE = /\[DoorLabs:([^\]]*)\]/i;
const CODE_RE_ALL = /\s*\[DoorLabs:[^\]]*\]/gi;
/** Liest den ersten DoorLabs-Code aus einem Text. null, wenn keiner da ist oder er unbrauchbar ist. */
function parseCode(text) {
    var _a, _b;
    const m = CODE_RE.exec(text !== null && text !== void 0 ? text : '');
    if (!m)
        return null;
    const parts = m[1].split('|').map((p) => p.trim());
    const rawStatus = cleanCodeText(parts[0]);
    if (!rawStatus)
        return null;
    const key = rawStatus.toLowerCase();
    let status = null;
    if (exports.STATUS_KEYS.includes(key))
        status = key;
    else if (LEGACY_STATUS[key])
        status = LEGACY_STATUS[key];
    const rawColor = ((_a = parts[1]) !== null && _a !== void 0 ? _a : '').toLowerCase();
    const color = exports.COLOR_KEYS.includes(rawColor)
        ? rawColor
        : ((_b = LEGACY_COLOR[rawColor]) !== null && _b !== void 0 ? _b : null);
    return {
        status,
        text: status ? '' : rawStatus,
        color: color !== null && color !== void 0 ? color : (status ? exports.STATUS_DEFAULT_COLOR[status] : 'blue'),
        device: cleanCodeText(parts[2]),
    };
}
/** Erster brauchbarer Code aus mehreren Feldern (Reihenfolge: Ort, Betreff, Beschreibung) */
function findCode(fields) {
    for (const f of fields) {
        const c = parseCode(f);
        if (c)
            return c;
    }
    return null;
}
/** Text ohne DoorLabs-Code (für Anzeige und zum Ersetzen) */
function stripCode(text) {
    return (text !== null && text !== void 0 ? text : '').replace(CODE_RE_ALL, '').trim();
}
/** Gilt der Code für dieses Gerät? Ohne Gerätename im Code gilt er für alle; Vergleich ohne Gross-/Kleinschreibung. */
function codeMatchesDevice(code, deviceName) {
    if (!code.device)
        return true;
    return code.device.trim().toLowerCase() === (deviceName !== null && deviceName !== void 0 ? deviceName : '').trim().toLowerCase();
}
/**
 * Welche Vorlage gehört zum Code?
 * 1. bekannter Status → Vorlage mit dem Namen des Status (in einer der vier Sprachen)
 * 2. eigener Text → Vorlage mit genau diesem Namen
 * 3. sonst die erste Vorlage in der Farbe des Codes (hexOf wandelt den Farbschlüssel in #rrggbb um)
 * null, wenn nichts passt; dann gelten die normalen Regeln (Stichwörter).
 */
function presetForCode(code, presets, hexOf = (c) => exports.COLOR_HEX[c]) {
    var _a, _b;
    if (code.status) {
        const names = exports.STATUS_PRESET_NAMES[code.status];
        const hit = presets.find((p) => names.includes(p.name.trim().toLowerCase()));
        if (hit)
            return hit.id;
    }
    else if (code.text) {
        const wanted = code.text.trim().toLowerCase();
        const hit = presets.find((p) => p.name.trim().toLowerCase() === wanted);
        if (hit)
            return hit.id;
    }
    const hex = hexOf(code.color).toLowerCase();
    return (_b = (_a = presets.find((p) => { var _a; return ((_a = p.color) !== null && _a !== void 0 ? _a : '').toLowerCase() === hex; })) === null || _a === void 0 ? void 0 : _a.id) !== null && _b !== void 0 ? _b : null;
}

  }),
  spectra: load(function (exports, module) {
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.DEFAULT_DISPLAY_STYLE = exports.DISPLAY_STYLES = exports.SPECTRA6 = void 0;
exports.hexToRgb = hexToRgb;
exports.nearestSpectra6 = nearestSpectra6;
exports.spectra6Display = spectra6Display;
exports.textColorFor = textColorFor;
exports.isDisplayStyle = isDisplayStyle;
exports.normalizeDisplayStyle = normalizeDisplayStyle;
exports.displayColors = displayColors;
exports.wcagContrast = wcagContrast;
/**
 * Die 6 Farben des E-Paper-Displays (Spectra 6). Das Display kann nur diese
 * 6 Farben zeigen. `ref` ist der reine Referenzwert für die Zuordnung,
 * `display` ist ein abgedunkelter Anzeigewert für die Vorschau. Die echten
 * Pigmentfarben sind nicht gemessen; die Anzeigewerte sind nur eine Näherung.
 */
exports.SPECTRA6 = [
    { name: 'black', ref: [0, 0, 0], display: '#1a1a1a' },
    { name: 'white', ref: [255, 255, 255], display: '#e8e6df' },
    { name: 'yellow', ref: [255, 255, 0], display: '#e6c619' },
    { name: 'red', ref: [255, 0, 0], display: '#b3261e' },
    { name: 'blue', ref: [0, 0, 255], display: '#233f8f' },
    { name: 'green', ref: [0, 255, 0], display: '#2e7d4f' },
];
function hexToRgb(hex) {
    const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
    if (!m)
        return [0, 0, 0];
    const n = parseInt(m[1], 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
/** Nächste Spectra-6-Farbe nach quadratischem RGB-Abstand. */
function nearestSpectra6(hex) {
    const [r, g, b] = hexToRgb(hex);
    let best = 'black';
    let bestDist = Infinity;
    for (const c of exports.SPECTRA6) {
        const d = Math.pow((r - c.ref[0]), 2) + Math.pow((g - c.ref[1]), 2) + Math.pow((b - c.ref[2]), 2);
        if (d < bestDist) {
            bestDist = d;
            best = c.name;
        }
    }
    return best;
}
function spectra6Display(name) {
    return exports.SPECTRA6.find((c) => c.name === name).display;
}
/** Schwarz auf Weiss und Gelb, sonst Weiss. */
function textColorFor(name) {
    return name === 'white' || name === 'yellow' ? 'black' : 'white';
}
exports.DISPLAY_STYLES = ['full', 'minimal'];
/** Neue Vorlagen starten mit «Schlicht» (Team-Entscheid 25.09.2026: Standard-Hintergrund schwarz). */
exports.DEFAULT_DISPLAY_STYLE = 'minimal';
function isDisplayStyle(value) {
    return value === 'full' || value === 'minimal';
}
/** Unbekannte oder fehlende Werte (alter Server, alte Daten) gelten als «Schlicht». */
function normalizeDisplayStyle(value) {
    return isDisplayStyle(value) ? value : exports.DEFAULT_DISPLAY_STYLE;
}
/**
 * Die EINE Farbregel für das Türschild (App, Simulator und Geräte-Software):
 *
 *  Stil      Vorlagenfarbe  Hintergrund  Text     Symbol
 *  Vollbild  Schwarz        Schwarz      Weiss    Weiss
 *  Vollbild  Weiss          Weiss        Schwarz  Schwarz
 *  Vollbild  Gelb           Gelb         Schwarz  Schwarz
 *  Vollbild  Rot            Rot          Weiss    Weiss
 *  Vollbild  Blau           Blau         Weiss    Weiss
 *  Vollbild  Grün           Grün         Weiss    Weiss
 *  Schlicht  Schwarz        Schwarz      Weiss    Weiss     (Schwarz auf Schwarz ginge nicht)
 *  Schlicht  Weiss          Schwarz      Weiss    Weiss
 *  Schlicht  Gelb           Schwarz      Gelb     Gelb
 *  Schlicht  Rot            Schwarz      Rot      Rot
 *  Schlicht  Blau           Schwarz      Weiss    Blau      (Kontrast Blau auf Schwarz nur 2,44)
 *  Schlicht  Grün           Schwarz      Grün     Grün
 *
 * Gespiegelt in device/doorlabs-epaper/doorlabs_layout.h (displayColorsFor). Ein Test vergleicht beide Tabellen.
 */
function displayColors(style, color) {
    const c = exports.SPECTRA6.some((s) => s.name === color) ? color : nearestSpectra6(String(color));
    if (style === 'full') {
        const text = textColorFor(c);
        return { background: c, text, icon: text, textKeptWhiteForContrast: false };
    }
    if (c === 'black')
        return { background: 'black', text: 'white', icon: 'white', textKeptWhiteForContrast: false };
    if (c === 'blue')
        return { background: 'black', text: 'white', icon: 'blue', textKeptWhiteForContrast: true };
    return { background: 'black', text: c, icon: c, textKeptWhiteForContrast: false };
}
/** WCAG-Kontrast zweier Display-Farben, mit den reinen Referenzfarben gerechnet (die echten Pigmente sind dunkler). */
function wcagContrast(a, b) {
    const lum = (name) => {
        const ref = exports.SPECTRA6.find((s) => s.name === name).ref;
        const [r, g, bl] = ref.map((v) => {
            const s = v / 255;
            return s <= 0.03928 ? s / 12.92 : Math.pow(((s + 0.055) / 1.055), 2.4);
        });
        return 0.2126 * r + 0.7152 * g + 0.0722 * bl;
    };
    const [hi, lo] = [lum(a), lum(b)].sort((x, y) => y - x);
    return (hi + 0.05) / (lo + 0.05);
}

  })
  };
})(typeof window !== "undefined" ? window : globalThis);
