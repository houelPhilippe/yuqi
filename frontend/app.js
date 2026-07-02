'use strict';

// ── State ─────────────────────────────────────────────────────────────────────
const state = {
  tabs: [],
  activeTabId: null,
  nextTabId: 1,
  editingPreview: false,
  sourceMode: false,
};

let turndown = null;
let previewInputTimer = null;
let pendingInitialFile = null;

// Couleurs de surbrillance : nom CSS (stocké en markdown) → rgba allégé pour l'affichage
const _HIGHLIGHT_COLORS = {
  'Yellow':      'rgba(255,220,0,0.55)',
  'Orange':      'rgba(255,140,0,0.45)',
  'LightGreen':  'rgba(50,200,80,0.35)',
  'Cyan':        'rgba(0,200,220,0.40)',
  'LightPink':   'rgba(255,100,150,0.35)',
  'Lavender':    'rgba(140,100,255,0.35)',
  'LightSalmon': 'rgba(255,100,80,0.40)',
  'SkyBlue':     'rgba(30,144,255,0.35)',
};

// ── Commentaires (annotations liées à une sélection, façon Word) ────────────────
// Encodage markdown : [texte sélectionné]{.comment comment-id="..." comment-text="..."}
// comment-text échappe \, " et les retours à la ligne pour tenir sur un attribut markdown.
function _annotEscape(s) {
  return (s || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\r\n|\r|\n/g, '\\n');
}
function _annotUnescape(s) {
  return (s || '').replace(/\\(.)/g, (_, c) => c === 'n' ? '\n' : c);
}
// Échappement HTML dédié pour l'attribut data-annot-text (indépendant de _escHtml).
function _annotHtmlEscape(s) {
  return (s || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function _newAnnotId() {
  return 'a' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

// ── Grid table helpers ─────────────────────────────────────────────────────────
// Retourne string[][] : un tableau de lignes par colonne
function parseCells(lines, colBounds, numCols) {
  const rawCols = new Array(numCols).fill(null).map(() => []);
  for (const line of lines) {
    for (let c = 0; c < numCols; c++) {
      const s = colBounds[c] + 1;
      const e = colBounds[c + 1];
      const seg = s < line.length ? line.slice(s, Math.min(e, line.length)).trimEnd() : '';
      // On retire seulement l'espace de gauche s'il existe
      const trimmed = seg.startsWith(' ') ? seg.slice(1) : seg;
      rawCols[c].push(trimmed);
    }
  }
  // Pour chaque colonne, on retourne le tableau de lignes.
  // Le tokenizer se chargera de joindre ces lignes.
  return rawCols;
}

// ── Coloration syntaxique YAML ────────────────────────────────────────────────
function highlightYaml(raw) {
  const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  const colorValue = (val) => {
    const t = val.trim();
    if (!t) return val;
    if (/^#/.test(t))                               return `<span class="yi-comment">${esc(val)}</span>`;
    if (/^["']/.test(t))                            return `<span class="yi-string">${esc(val)}</span>`;
    if (/^(true|false|yes|no|on|off)$/i.test(t))   return `<span class="yi-bool">${esc(val)}</span>`;
    if (/^(null|~)$/i.test(t))                      return `<span class="yi-null">${esc(val)}</span>`;
    if (/^-?\d[\d.,_]*(%|[eE][+-]?\d+)?$/.test(t)) return `<span class="yi-number">${esc(val)}</span>`;
    return `<span class="yi-value">${esc(val)}</span>`;
  };

  return raw.split('\n').map(line => {
    // Commentaire seul
    if (/^\s*#/.test(line)) return `<span class="yi-comment">${esc(line)}</span>`;
    // Liste : - item
    const listM = line.match(/^(\s*-\s)(.*)$/);
    if (listM) return `${esc(listM[1])}${colorValue(listM[2])}`;
    // Clé : valeur  (y compris clés avec guillemets)
    const kvM = line.match(/^(\s*)((?:["'][^"']*["']|[\w\-. ]+)\s*)(:[ \t]*)(.*)$/);
    if (kvM) {
      const [, indent, key, sep, val] = kvM;
      // Valeur vide → simple clé (objet imbriqué)
      const valHtml = val.trim() ? colorValue(val) : '';
      return `${esc(indent)}<span class="yi-key">${esc(key)}</span><span class="yi-sep">${esc(sep)}</span>${valHtml}`;
    }
    return esc(line);
  }).join('\n');
}

// ── Helpers image ─────────────────────────────────────────────────────────────
let _currentDocBasePath = '';

// ── Citations BibTeX — état global ────────────────────────────────────────────
let _bibEntries   = null;  // cache des entrées parsées (null = non chargé)
let _citePopupIdx = -1;    // index de l'entrée sélectionnée dans le popup

/**
 * Insère `node` juste après le bloc qui contient le curseur,
 * en remontant jusqu'au fils direct de `#preview`.
 */
function _insertAtCursor(node, savedRange) {
  const preview = document.getElementById('preview');
  const sel = window.getSelection();
  const range = savedRange || (sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null);

  // Déterminer l'ancre.
  // Priorité : si le clic droit était dans un .fenced-div-content, on l'utilise
  // en priorité sur le range — le curseur peut être ailleurs (hors DIV) si
  // l'utilisateur a fait un clic droit sans avoir d'abord cliqué dans la DIV.
  let anchor = null;
  const savedAnchor = _scInsertionAnchor;
  _scInsertionAnchor = null;

  const fdivFromTarget = savedAnchor?.closest?.('.fenced-div-content');
  if (fdivFromTarget && preview.contains(savedAnchor)) {
    // Clic droit dans une DIV → conteneur = cette DIV
    const container = fdivFromTarget;
    // Si le range est aussi dans cette DIV, insérer après son nœud bloc
    let insertAfter = null;
    if (range && container.contains(range.commonAncestorContainer)) {
      insertAfter = range.commonAncestorContainer;
      if (insertAfter.nodeType === Node.TEXT_NODE) insertAfter = insertAfter.parentElement;
      while (insertAfter && insertAfter.parentElement !== container) insertAfter = insertAfter.parentElement;
    }
    if (insertAfter && insertAfter !== container) {
      insertAfter.after(node);
    } else {
      container.appendChild(node);
    }
    return;
  }

  // Pas dans une DIV : utiliser le range puis fallback sur savedAnchor.
  // On préfère startContainer (précis même pour une sélection multi-blocs) ;
  // commonAncestorContainer peut valoir #preview lui-même, ce qui ferait
  // remonter le traversal hors du DOM → insertion en fin de document.
  // Si le range cible directement #preview (clic dans une zone vide entre
  // ou après les blocs), on insère au niveau de l'enfant correspondant à
  // l'offset plutôt que de tomber dans le fallback "fin de document".
  let previewOffset = null;

  if (range) {
    const startCont = range.startContainer;
    const cac       = range.commonAncestorContainer;
    let candidate   = preview.contains(startCont) ? startCont : (preview.contains(cac) ? cac : null);
    if (candidate === preview && startCont === preview) {
      previewOffset = range.startOffset;
    }
    if (candidate) {
      if (candidate.nodeType === Node.TEXT_NODE) candidate = candidate.parentElement;
      if (candidate !== preview) anchor = candidate;
    }
  }
  if (!anchor && savedAnchor && preview.contains(savedAnchor) && savedAnchor !== preview) {
    anchor = savedAnchor;
  }

  if (anchor) {
    const container = anchor.closest?.('.fenced-div-content') || preview;
    let insertAfter = anchor;
    while (insertAfter && insertAfter.parentElement !== container) {
      insertAfter = insertAfter.parentElement;
    }
    if (insertAfter && insertAfter !== container) {
      insertAfter.after(node);
    } else {
      container.appendChild(node);
    }
    return;
  }
  if (previewOffset !== null) {
    preview.insertBefore(node, preview.childNodes[previewOffset] || null);
    return;
  }
  preview.appendChild(node);
}   // mis à jour dans updatePreview()

function _pathToUrl(p) {
  if (!p) return '';
  if (p.startsWith('http://') || p.startsWith('https://') || p.startsWith('file:///')) return p;
  if (/^[a-zA-Z]:/.test(p)) return 'file:///' + p.replace(/\\/g, '/');
  if (p.startsWith('/') || p.startsWith('\\')) return 'file://' + p.replace(/\\/g, '/');
  // Chemin relatif → résoudre par rapport au répertoire du document courant
  if (_currentDocBasePath) return 'file:///' + _currentDocBasePath + '/' + p;
  return p;
}

/** Calcule le chemin de `toFile` relatif au répertoire de `fromFile`. */
function _makeRelativePath(fromFile, toFile) {
  const norm = s => s.replace(/\\/g, '/');
  const to = norm(toFile);
  if (!fromFile) return to;   // document non sauvegardé → chemin absolu
  const fromDir = norm(fromFile).replace(/\/[^/]+$/, '');
  const fromParts = fromDir.split('/');
  const toParts   = to.split('/');
  let i = 0;
  while (i < fromParts.length && i < toParts.length &&
         fromParts[i].toLowerCase() === toParts[i].toLowerCase()) i++;
  const rel = [...Array(fromParts.length - i).fill('..'), ...toParts.slice(i)];
  return rel.length ? rel.join('/') : toParts[toParts.length - 1];
}

// ── Init ──────────────────────────────────────────────────────────────────────
window.addEventListener('pywebviewready', init);
window.addEventListener('DOMContentLoaded', () => {
  // Appliquer le thème immédiatement pour éviter le flash blanc
  try {
    const s = JSON.parse(localStorage.getItem('md_editor_settings')) || {};
    const t = s.theme || 'dark';
    if (t !== 'dark') document.documentElement.setAttribute('data-theme', t);
  } catch (_) {}
  if (!window.pywebview) setTimeout(init, 100);
});

function init() {
  if (turndown) return;

  _initTheme();

  marked.setOptions({ breaks: true, gfm: true });

  marked.use({
    extensions: [
      // ── HTML littéral : balises saisies directement dans le source ──────────
      {
        name: 'htmlliteralblock',
        level: 'block',
        start(src) { const m = src.match(/^[ \t]*<\/?[a-zA-Z]/m); return m ? m.index : src.length; },
        tokenizer(src) {
          // Une ou plusieurs lignes consécutives qui ressemblent à des balises HTML
          const match = src.match(/^(?:[ \t]*<\/?[a-zA-Z][^>\n]*(?:>|$)[ \t]*\n?)+/);
          if (!match) return;
          return { type: 'htmlliteralblock', raw: match[0], text: match[0].trimEnd() };
        },
        renderer(token) {
          const escaped = token.text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
          return `<div class="md-html-literal">${escaped}</div>\n`;
        }
      },
      {
        name: 'htmlliteralinline',
        level: 'inline',
        start(src) { return src.indexOf('<'); },
        tokenizer(src) {
          const match = src.match(/^<\/?[a-zA-Z][^>]*>/);
          if (!match) return;
          return { type: 'htmlliteralinline', raw: match[0], text: match[0] };
        },
        renderer(token) {
          const escaped = token.text.replace(/</g, '&lt;').replace(/>/g, '&gt;');
          return `<span class="md-html-literal">${escaped}</span>`;
        }
      },
      // ── Titres avec signet {#id} ─────────────────────────────────────────────
      {
        name: 'bookmarkheading',
        level: 'block',
        start(src) { const m = src.match(/^#{1,6}\s/m); return m ? m.index : src.length; },
        tokenizer(src) {
          const match = src.match(/^(#{1,6})\s+(.*?)\s*\{#([\w-]+)\}\s*(?:\n|$)/);
          if (!match) return;
          const depth   = match[1].length;
          const rawText = match[2].trim();
          const id      = match[3];
          const token   = { type: 'bookmarkheading', raw: match[0], depth, id, tokens: [] };
          this.lexer.inline(rawText, token.tokens);
          return token;
        },
        renderer(token) {
          const inner = this.parser.parseInline(token.tokens);
          return `<h${token.depth} id="${token.id}" data-bookmark="${token.id}">${inner}<span class="heading-bookmark" contenteditable="false">#${token.id}</span></h${token.depth}>\n`;
        }
      },
      // ────────────────────────────────────────────────────────────────────────
      {
        name: 'gridtable',
        level: 'block',
        start(src) { const m = src.match(/^\+[ \t]*[-=:]/m); return m ? m.index : src.length; },
        tokenizer(src) {
          if (!/^\+[ \t]*[-=:]/.test(src)) return;
          const lines = src.split('\n');
          const tableLines = [];
          for (const line of lines) {
            if (/^[|+]/.test(line)) tableLines.push(line);
            else break;
          }
          if (tableLines.length < 3) return;
          if (!/^\+[-=:]/.test(tableLines[tableLines.length - 1])) return;

          const afterTable = src.slice(tableLines.join('\n').length + 1);
          let colPcts = null;
          let colValigns = null;
          let captionText = '';
          let captionRaw = '';
          let attrRaw = '';
          let bordered = false;
          let striped  = false;

          const parseAttrContent = (inner) => {
            const cwM = inner.match(/tbl-colwidths\s*=\s*"\[([^\]]+)\]"/);
            if (cwM) colPcts = cwM[1].split(',').map(s => s.trim()).filter(Boolean);
            const cvM = inner.match(/tbl-colvaligns\s*=\s*"\[([^\]]+)\]"/);
            if (cvM) colValigns = cvM[1].split(',').map(s => s.trim()).filter(Boolean);
            if (inner.includes('.bordered')) bordered = true;
            if (inner.includes('.striped'))  striped  = true;
          };

          const captionLineMatch = afterTable.match(/^:[ \t]*([^{\n]*)[ \t]*(\{[^}]*\})?[ \t]*\n?/);
          if (captionLineMatch) {
            captionRaw  = captionLineMatch[0];
            captionText = captionLineMatch[1].trim();
            const attrStr = captionLineMatch[2] || '';
            if (attrStr) parseAttrContent(attrStr);
          } else {
            // Compatibilité : "{attr}" seul sans légende
            const attrMatch = afterTable.match(/^\{([^}]*)\}[ \t]*\n?/);
            if (attrMatch) {
              attrRaw = attrMatch[0];
              parseAttrContent(attrMatch[1]);
            }
          }

          const raw = tableLines.join('\n') + '\n' + captionRaw + attrRaw;
          const firstLine = tableLines[0];
          const colBounds = [];
          for (let i = 0; i < firstLine.length; i++) {
            if (firstLine[i] === '+') colBounds.push(i);
          }
          const numCols = colBounds.length - 1;
          if (numCols < 1) return;

          let headerSepIdx = -1;
          const aligns = new Array(numCols).fill('');
          let hasAlignMarkers = false;
          for (let i = 0; i < tableLines.length; i++) {
            const sep = tableLines[i];
            if (sep.startsWith('+')) {
              if (sep.includes(':')) {
                hasAlignMarkers = true;
                if (i > 0 && headerSepIdx === -1 && sep.includes('=')) headerSepIdx = i;
                for (let c = 0; c < numCols; c++) {
                  const seg = sep.slice(colBounds[c] + 1, colBounds[c + 1]);
                  const left = seg[0] === ':';
                  const right = seg[seg.length - 1] === ':';
                  if (left && right) aligns[c] = 'center';
                  else if (right) aligns[c] = 'right';
                  else if (left) aligns[c] = 'left';
                }
                // Si on a trouvé des marqueurs, on a les alignements pour toute la table
              } else if (i > 0 && headerSepIdx === -1 && sep.startsWith('+=')) {
                headerSepIdx = i;
              }
            }
          }

          let inHeader = headerSepIdx !== -1;
          let currentLines = [];
          const headerRows = [], bodyRows = [];
          for (let i = 1; i < tableLines.length; i++) {
            const line = tableLines[i];
            if (line.startsWith('+')) {
              if (currentLines.length > 0) {
                const cells = parseCells(currentLines, colBounds, numCols);
                if (inHeader) headerRows.push(cells);
                else bodyRows.push(cells);
                currentLines = [];
              }
              if (i === headerSepIdx) inHeader = false;
            } else {
              currentLines.push(line);
            }
          }

          const renderCell = (tag, paras, align, valign, extra = '') => {
            const styles = [];
            if (align)  styles.push(`text-align:${align}`);
            if (valign) styles.push(`vertical-align:${valign}`);
            const styleAttr = styles.length ? ` style="${styles.join(';')}"` : '';

            // On joint les lignes avec des retours à la ligne pour marked.parse.
            // 1. trimEnd() : supprime les espaces trailing pour éviter que "  \n"
            //    (produit par turndown à partir des <br> que le navigateur injecte
            //    dans les <li> vides lors d'un DOM corrompu) ne soit interprété
            //    par marked (breaks:true) comme un <br> involontaire.
            // 2. On échappe les marqueurs de liste (- / * / +) en début de ligne,
            //    y compris le marqueur seul en fin de ligne (\s|$), pour éviter
            //    qu'un texte comme "- - - - -  texte" soit rendu comme des listes
            //    imbriquées dans une cellule (marked crée sinon 1 niveau par "- ").
            //    N.B. : on n'échappe que si le marqueur est suivi d'un autre marqueur
            //    (ex. "- - texte") ; les vrais items de liste "- texte" restent intacts.
            const escapedParas = paras.map(line => {
              const stripped = line.trimEnd();
              return stripped.replace(/^(\s*)([-*+])([ \t]+[-*+])/, '$1\\$2$3');
            });
            const contentMd = escapedParas.join('\n');
            const contentHtml = marked.parse(contentMd);

            return `<${tag}${styleAttr}${extra}>${contentHtml}</${tag}>`;
          };

          let tableStyle = '';
          let colWidthsForColgroup = null;
          if (colPcts) {
            // Les % dans {tbl-colwidths} sont relatifs au conteneur.
            // La table prend totalPct% du conteneur.
            // Les <col> doivent être en % de la TABLE : col% = (valeur / total) * 100
            const nums     = colPcts.map(w => parseFloat(w));
            const sumPct   = nums.reduce((s, w) => s + w, 0);
            const totalPct = Math.min(sumPct, 100);
            tableStyle = ` style="table-layout:fixed;width:${totalPct}%"`;
            // Normaliser par la somme réelle (pas par totalPct, qui peut être
            // tronqué à 100) pour que les <col> totalisent toujours 100% de la table.
            colWidthsForColgroup = nums.map(w => ((w / sumPct) * 100).toFixed(4));
          }
          const dataAttr      = colPcts   ? ` data-col-widths="${colPcts.join(',')}"` : '';
          const dataVAttr     = colValigns ? ` data-col-valigns="${colValigns.join(',')}"` : '';
          const dataBordered  = bordered ? ` data-bordered="1"` : '';
          const dataStriped   = striped  ? ` data-striped="1"`  : '';
          const extraClasses  = (bordered ? ' bordered' : '') + (striped ? ' striped' : '');
          let html = `<table class="grid-table${extraClasses}"${tableStyle}${dataAttr}${dataVAttr}${dataBordered}${dataStriped}>\n`;
          if (captionText) html += `<caption>${captionText}</caption>\n`;
          if (colWidthsForColgroup) {
            html += '<colgroup>';
            colWidthsForColgroup.forEach(w => { html += `<col style="width:${w}%">`; });
            html += '</colgroup>\n';
          }
          if (headerRows.length > 0) {
            html += '<thead>\n';
            for (const row of headerRows) {
              html += '<tr>';
              row.forEach((paras, c) => {
                const wAttr = colPcts ? ` data-width="${colPcts[c]}%"` : '';
                html += renderCell('th', paras, aligns[c], colValigns?.[c] || '', wAttr);
              });
              html += '</tr>\n';
            }
            html += '</thead>\n';
          }
          html += '<tbody>\n';
          for (const row of bodyRows) {
            html += '<tr>';
            row.forEach((paras, c) => {
              html += renderCell('td', paras, aligns[c], colValigns?.[c] || '');
            });
            html += '</tr>\n';
          }
          html += '</tbody>\n</table>\n';
          return { type: 'gridtable', raw, html };
        },
        renderer(token) { return token.html; }
      },
      {
        name: 'figure',
        level: 'block',
        start(src) { const m = src.match(/^!\[/m); return m ? m.index : src.length; },
        tokenizer(src) {
          const match = src.match(/^!\[((?:[^\[\]]|\[[^\[\]]*\])*)\]\(([^)]+)\)(?:\{([^}]*)\})?[ \t]*\n?/);
          if (!match) return;
          const alt     = match[1] || '';
          const rawSrc  = match[2];
          const attrStr = match[3] || '';
          const alignM  = attrStr.match(/fig-align="(left|center|right)"/) || attrStr.match(/\.align-(left|center|right)/);
          const align   = alignM ? alignM[1] : 'center';
          const widthM  = attrStr.match(/width=(\d+)%/);
          const imgW    = widthM ? parseInt(widthM[1]) : 100;
          const figClass = `fig-${align}`;
          const imgStyle = imgW < 100 ? ` style="width:${imgW}%"` : '';
          const capInner = alt ? marked.parseInline(alt) : '';
          const capHtml  = alt ? `<figcaption>${capInner}</figcaption>` : '';
          const altText  = alt.replace(/\[([^\]]*)\]\([^)]*\)/g, '$1');
          const html = `<figure class="md-figure ${figClass}" data-src="${rawSrc}" data-align="${align}" data-imgwidth="${imgW}"><img src="${_pathToUrl(rawSrc)}" alt="${altText}"${imgStyle}>${capHtml}</figure>\n`;
          return { type: 'figure', raw: match[0], html };
        },
        renderer(token) { return token.html; }
      },
      {
        name: 'fenceddiv',
        level: 'block',
        start(src) { const m = src.match(/^:{3,}/m); return m ? m.index : src.length; },
        tokenizer(src) {
          // Backreference sur le délimiteur : ::::...:::: englobe :::...:::
          const match = src.match(/^(:{3,})[ \t]*(\{([^}]*)\})?[ \t]*\n([\s\S]*?)\n\1[ \t]*(?:\n|$)/);
          if (!match) return;
          const attrStr  = match[3] || '';
          const inner    = match[4] || '';
          const classM   = attrStr.match(/\.[\w-]+/g) || [];
          const divClass = classM.map(c => c.slice(1)).join(' ');
          const formatM  = attrStr.match(/when-format="([^"]+)"/);
          const format   = formatM ? formatM[1] : '';
          const styleM   = attrStr.match(/style="([^"]+)"/);
          const styleStr = styleM ? styleM[1] : '';
          const widthM   = styleStr.match(/width:\s*(\d+)%/);
          const divWidth = widthM ? widthM[1] : '';
          const divMargin = styleStr.includes('margin: auto') || styleStr.includes('margin:auto') ? '1' : '';
          const innerHtml = marked.parse(inner);
          const inlineStyle = styleStr ? ` style="${styleStr}"` : '';
          const dataWidth  = divWidth  ? ` data-div-width="${divWidth}"`   : '';
          const dataMargin = divMargin ? ` data-div-margin="1"` : '';
          const dataClass  = divClass ? ` data-div-class="${divClass}"` : '';
          const dataFormat = format   ? ` data-div-format="${format}"` : '';
          const classAttr  = divClass ? `fenced-div-wrapper ${divClass}` : 'fenced-div-wrapper';
          const html = `<div class="${classAttr}"${dataClass}${dataFormat}${dataWidth}${dataMargin}${inlineStyle}><span class="fenced-div-label" contenteditable="false">DV</span><div class="fenced-div-content">${innerHtml}</div></div>\n`;
          return { type: 'fenceddiv', raw: match[0], html };
        },
        renderer(token) { return token.html; }
      },
      {
        name: 'yamlblock',
        level: 'block',
        start(src) { return src.indexOf('---'); },
        tokenizer(src) {
          const match = src.match(/^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/);
          if (match) return { type: 'yamlblock', raw: match[0], text: match[1] };
        },
        renderer(token) {
          return `<div class="yaml-block"><div class="yaml-block-header" contenteditable="false" onclick="this.closest('.yaml-block').classList.toggle('collapsed')"><svg class="yaml-toggle-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>― YAML ―</div><pre><code>${highlightYaml(token.text)}</code></pre></div>`;
        }
      },
      {
        name: 'htmlcomment',
        level: 'block',
        start(src) { const m = src.match(/^<!--/m); return m ? m.index : src.length; },
        tokenizer(src) {
          const match = src.match(/^<!--([\s\S]*?)-->[ \t]*(?:\n|$)/);
          if (!match) return;
          const text = match[1].trim();
          const safeText = text.replace(/"/g, '&quot;');
          return {
            type: 'htmlcomment',
            raw: match[0],
            html: _buildCommentHtml(text) + '\n'
          };
        },
        renderer(token) { return token.html; }
      },
      {
        name: 'rawhtml',
        level: 'block',
        start(src) { const m = src.match(/^```\{=html\}/m); return m ? m.index : src.length; },
        tokenizer(src) {
          const match = src.match(/^```\{=html\}\r?\n([\s\S]*?)\r?\n```(?:[ \t]*(?:\n|$))/);
          if (!match) return;
          const innerHtml = match[1];
          const safeRaw   = innerHtml.replace(/"/g, '&quot;');
          const html = `<div class="md-rawhtml" data-html="${safeRaw}" contenteditable="false"><div class="md-rawhtml-content">${innerHtml}</div></div>\n`;
          return { type: 'rawhtml', raw: match[0], html };
        },
        renderer(token) { return token.html; }
      },
      {
        name: 'shortcode',
        level: 'block',
        start(src) { const m = src.match(/^\{\{</m); return m ? m.index : src.length; },
        tokenizer(src) {
          const match = src.match(/^\{\{<\s*([^>]+?)\s*>\}\}[ \t]*(?:\n|$)/);
          if (!match) return;
          const inner = match[1].trim();
          const rawText = `{{< ${inner} >}}`;
          const safeRaw = rawText.replace(/"/g, '&quot;');
          const html = `<div class="md-shortcode" data-sc-raw="${safeRaw}" contenteditable="false">`
            + `<span class="sc-brace">{{&lt;</span> <span class="sc-inner">${inner}</span> <span class="sc-brace">&gt;}}</span>`
            + `</div>\n`;
          return { type: 'shortcode', raw: match[0], html };
        },
        renderer(token) { return token.html; }
      },
      {
        name: 'underline',
        level: 'inline',
        start(src) { return src.indexOf('['); },
        tokenizer(src) {
          const match = src.match(/^\[([^\]]+)\]\{\.underline\}/);
          if (!match) return;
          const token = { type: 'underline', raw: match[0], text: match[1], tokens: [] };
          this.lexer.inline(token.text, token.tokens);
          return token;
        },
        renderer(token) { return `<u>${this.parser.parseInline(token.tokens)}</u>`; }
      },
      {
        name: 'smallcaps',
        level: 'inline',
        start(src) { return src.indexOf('['); },
        tokenizer(src) {
          const match = src.match(/^\[([^\]]+)\]\{\.smallcaps\}/);
          if (!match) return;
          const token = { type: 'smallcaps', raw: match[0], text: match[1], tokens: [] };
          this.lexer.inline(token.text, token.tokens);
          return token;
        },
        renderer(token) { return `<span class="smallcaps">${this.parser.parseInline(token.tokens)}</span>`; }
      },
      {
        name: 'highlight',
        level: 'inline',
        start(src) { return src.indexOf('['); },
        tokenizer(src) {
          if (src[0] !== '[') return;
          let depth = 0, i = 0;
          for (; i < src.length; i++) {
            if (src[i] === '[') depth++;
            else if (src[i] === ']') { if (--depth === 0) break; }
          }
          if (depth !== 0) return;
          const text = src.slice(1, i);
          const attrMatch = src.slice(i + 1).match(/^\{style="background-color:\s*([^"]+?);?"\}/);
          if (!attrMatch) return;
          const token = { type: 'highlight', raw: src.slice(0, i + 1) + attrMatch[0], text, color: attrMatch[1].trim(), tokens: [] };
          this.lexer.inline(token.text, token.tokens);
          return token;
        },
        renderer(token) {
          const displayColor = _HIGHLIGHT_COLORS[token.color] || token.color;
          return `<span class="md-highlight" style="background-color:${displayColor}" data-color="${token.color}">${this.parser.parseInline(token.tokens)}</span>`;
        }
      },
      {
        name: 'textcolor',
        level: 'inline',
        start(src) { return src.indexOf('['); },
        tokenizer(src) {
          if (src[0] !== '[') return;
          let depth = 0, i = 0;
          for (; i < src.length; i++) {
            if (src[i] === '[') depth++;
            else if (src[i] === ']') { if (--depth === 0) break; }
          }
          if (depth !== 0) return;
          const text = src.slice(1, i);
          const attrMatch = src.slice(i + 1).match(/^\{style="color:\s*([^"]+?);?"\}/);
          if (!attrMatch) return;
          const token = { type: 'textcolor', raw: src.slice(0, i + 1) + attrMatch[0], text, color: attrMatch[1].trim(), tokens: [] };
          this.lexer.inline(token.text, token.tokens);
          return token;
        },
        renderer(token) {
          return `<span class="md-textcolor" style="color:${token.color}" data-color="${token.color}">${this.parser.parseInline(token.tokens)}</span>`;
        }
      },
      {
        name: 'annotation',
        level: 'inline',
        start(src) { return src.indexOf('['); },
        tokenizer(src) {
          if (src[0] !== '[') return;
          let depth = 0, i = 0;
          for (; i < src.length; i++) {
            if (src[i] === '[') depth++;
            else if (src[i] === ']') { if (--depth === 0) break; }
          }
          if (depth !== 0) return;
          const text = src.slice(1, i);
          const attrMatch = src.slice(i + 1).match(/^\{\.comment comment-id="([^"]*)" comment-text="((?:\\.|[^"\\])*)"\}/);
          if (!attrMatch) return;
          const token = {
            type: 'annotation',
            raw: src.slice(0, i + 1) + attrMatch[0],
            text,
            annotId:   attrMatch[1],
            annotText: _annotUnescape(attrMatch[2]),
            tokens: [],
          };
          this.lexer.inline(token.text, token.tokens);
          return token;
        },
        renderer(token) {
          const safeText = _annotHtmlEscape(token.annotText);
          return `<span class="md-annotation" data-annot-id="${token.annotId}" data-annot-text="${safeText}">${this.parser.parseInline(token.tokens)}</span>`;
        }
      },
      // ── Note de bas de page (bloc) : [^label]: texte ──────────────────────
      {
        name: 'footnoteDef',
        level: 'block',
        start(src) { return src.indexOf('[^'); },
        tokenizer(src) {
          const match = src.match(/^\[\^([^\]]+)\]:\s+(.+)/);
          if (!match) return;
          return { type: 'footnoteDef', raw: match[0], label: match[1], text: match[2] };
        },
        renderer(token) {
          return `<div class="footnote-def" id="fn-${token.label}" data-label="${token.label}"><sup class="footnote-def-num">${token.label}</sup> <span class="footnote-def-text">${token.text}</span></div>\n`;
        }
      },
      // ── Référence de note (inline) : [^label] ─────────────────────────────
      {
        name: 'footnoteRef',
        level: 'inline',
        start(src) { return src.indexOf('[^'); },
        tokenizer(src) {
          const match = src.match(/^\[\^([^\]]+)\]/);
          if (!match) return;
          return { type: 'footnoteRef', raw: match[0], label: match[1] };
        },
        renderer(token) {
          return `<sup class="footnote-ref" data-label="${token.label}"><a href="#fn-${token.label}">[${token.label}]</a></sup>`;
        }
      },
      // ── Référence bibliographique Quarto/Pandoc : @clé ────────────────────
      {
        name: 'citation',
        level: 'inline',
        start(src) { return src.indexOf('@'); },
        tokenizer(src) {
          // Clé BibTeX : commence par une lettre, suivie de lettres/chiffres/- / _ / :
          const match = src.match(/^@([a-zA-Z][\w:-]*)/);
          if (!match) return;
          return { type: 'citation', raw: match[0], key: match[1] };
        },
        renderer(token) {
          return `<span class="md-citation" data-key="${token.key}">@${token.key}</span>`;
        }
      }
    ]
  });

  // Surcharge les renderers checkbox et listitem de marked.
  // checkbox : retire "disabled" pour l'interactivité, ajoute la classe CSS.
  // listitem : ajoute class="task-item" sur les éléments tâche pour que le
  //            CSS puisse masquer la puce sans dépendre de :has().
  marked.use({
    renderer: {
      checkbox(checked) {
        return `<input type="checkbox" class="task-checkbox"${checked ? ' checked' : ''}>`;
      },
      listitem(text, task, checked) {
        if (task) {
          // Tight list : '<input ...> texte'
          let m = text.match(/^(<input[^>]+>)([\s\S]*)$/);
          if (m) {
            return `<li class="task-item">${m[1]}<div class="task-label">${m[2].trimStart()}</div></li>\n`;
          }
          // Loose list : '<p><input ...> texte</p>...'
          m = text.match(/^<p>(<input[^>]+>)\s*([\s\S]*?)<\/p>([\s\S]*)$/);
          if (m) {
            const rest = m[3].trim();
            const content = m[2] + (rest ? '\n' + rest : '');
            return `<li class="task-item">${m[1]}<div class="task-label">${content}</div></li>\n`;
          }
        }
        return task ? `<li class="task-item">${text}</li>\n` : `<li>${text}</li>\n`;
      },
      list(body, ordered, start) {
        const tag = ordered ? 'ol' : 'ul';
        const startAttr = ordered && start !== 1 ? ` start="${start}"` : '';
        const cls = !ordered && body.includes('class="task-item"') ? ' class="task-list"' : '';
        return `<${tag}${startAttr}${cls}>\n${body}</${tag}>\n`;
      }
    }
  });

  turndown = new TurndownService({
    headingStyle: 'atx',
    bulletListMarker: '-',
    codeBlockStyle: 'fenced',
    emDelimiter: '*',
  });

  turndown.addRule('bookmarkheading', {
    filter: node => /^H[1-6]$/.test(node.nodeName) && node.getAttribute('data-bookmark'),
    replacement: (_, node) => {
      const level  = parseInt(node.nodeName[1]);
      const id     = node.getAttribute('data-bookmark');
      const prefix = '#'.repeat(level);
      // Clone le nœud et retire le badge pour ne pas l'inclure dans le texte
      const clone = node.cloneNode(true);
      const badge = clone.querySelector('.heading-bookmark');
      if (badge) badge.remove();
      // Convertit l'innerHTML via turndown pour préserver le formatage inline (surbrillance, gras, etc.)
      const text = turndown.turndown(clone.innerHTML).trim();
      return `\n\n${prefix} ${text} {#${id}}\n\n`;
    }
  });

  turndown.addRule('htmlliteral', {
    filter: node => (node.nodeName === 'DIV' || node.nodeName === 'SPAN') && node.classList.contains('md-html-literal'),
    replacement: (_, node) => {
      // Restaurer le HTML brut en décodant les entités
      const raw = (node.textContent || '')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>');
      return node.nodeName === 'DIV' ? `\n\n${raw}\n\n` : raw;
    }
  });

  turndown.addRule('htmlcomment', {
    filter: node => node.nodeName === 'DIV' && node.classList.contains('md-comment'),
    replacement: (_, node) => {
      const text = (node.getAttribute('data-comment') || '').replace(/&quot;/g, '"');
      return `\n\n<!-- ${text} -->\n\n`;
    }
  });

  turndown.addRule('rawhtml', {
    filter: node => node.nodeName === 'DIV' && node.classList.contains('md-rawhtml'),
    replacement: (_, node) => {
      const raw = (node.getAttribute('data-html') || '').replace(/&quot;/g, '"');
      return `\n\n\`\`\`{=html}\n${raw}\n\`\`\`\n\n`;
    }
  });

  turndown.addRule('shortcode', {
    filter: node => node.nodeName === 'DIV' && node.classList.contains('md-shortcode'),
    replacement: (_, node) => {
      const raw = (node.getAttribute('data-sc-raw') || '').replace(/&quot;/g, '"');
      return `\n\n${raw}\n\n`;
    }
  });

  turndown.addRule('highlight', {
    filter: node => node.nodeName === 'SPAN' && node.classList.contains('md-highlight'),
    replacement: (content, node) => {
      const color = node.getAttribute('data-color') || node.style.backgroundColor || 'Yellow';
      return `[${content}]{style="background-color: ${color};"}`;
    }
  });

  turndown.addRule('textcolor', {
    filter: node => node.nodeName === 'SPAN' && node.classList.contains('md-textcolor'),
    replacement: (content, node) => {
      const color = node.getAttribute('data-color') || node.style.color || 'Black';
      return `[${content}]{style="color: ${color};"}`;
    }
  });

  turndown.addRule('annotation', {
    filter: node => node.nodeName === 'SPAN' && node.classList.contains('md-annotation'),
    replacement: (content, node) => {
      const id   = node.getAttribute('data-annot-id')   || '';
      const text = node.getAttribute('data-annot-text') || '';
      return `[${content}]{.comment comment-id="${id}" comment-text="${_annotEscape(text)}"}`;
    }
  });

  turndown.addRule('fenceddiv', {
    filter: node => node.nodeName === 'DIV' && node.classList.contains('fenced-div-wrapper'),
    replacement: (_, node) => {
      const divClass   = node.getAttribute('data-div-class')  || '';
      const format     = node.getAttribute('data-div-format') || '';
      const contentEl  = node.querySelector('.fenced-div-content');
      const divWidth   = node.getAttribute('data-div-width')  || '';
      const divMargin  = node.getAttribute('data-div-margin') === '1';

      // Si la DIV contient des sous-DIVs enfants, ignorer le texte direct :
      // une DIV parente ne doit contenir que ses DIVs enfants, pas de texte.
      let innerMd = '';
      if (contentEl) {
        const childWrappers = [...contentEl.querySelectorAll(':scope > .fenced-div-wrapper')];
        if (childWrappers.length > 0) {
          innerMd = childWrappers
            .map(w => turndown.turndown(w.outerHTML).trim())
            .join('\n');
        } else {
          innerMd = turndown.turndown(contentEl.innerHTML).trim();
        }
      }

      const attrParts  = [];
      if (divClass) attrParts.push(...divClass.split(' ').map(c => '.' + c));
      if (format)   attrParts.push(`when-format="${format}"`);
      const styleParts = [];
      if (divWidth)  styleParts.push(`width: ${divWidth}%`);
      if (divMargin) styleParts.push('margin: auto');
      if (styleParts.length) attrParts.push(`style="${styleParts.join('; ')}"`);
      const attrBlock = attrParts.length ? ` {${attrParts.join(' ')}}` : '';

      // Calcule la profondeur de clôture à partir du contenu converti :
      // le délimiteur parent doit toujours avoir plus de colonnes que les délimiteurs enfants.
      const maxInnerFence = (innerMd.match(/^:{3,}/gm) || [])
        .reduce((m, s) => Math.max(m, s.length), 0);
      const fence = ':'.repeat(Math.max(3, maxInnerFence + 1));

      return `\n\n${fence}${attrBlock}\n${innerMd}\n${fence}\n\n`;
    }
  });

  turndown.addRule('gridtable', {
    filter: node => node.nodeName === 'TABLE',
    replacement: (_, node) => {
      // Extraire les paragraphes d'une cellule HTML
      // Utilise turndown pour préserver tout le formatage (gras, listes, etc.)
      const getCellParas = (cell) => {
        let md = turndown.turndown(cell.innerHTML).trim();
        // Turndown échappe le '-' en début de ligne (→ '\-') pour éviter qu'il soit
        // interprété comme un marqueur de liste. Mais des tirets consécutifs comme
        // "------" sont du texte pur (jamais des marqueurs de liste) et ne doivent
        // pas être échappés. On réintroduit le tiret non-échappé pour tout motif
        // "backslash + 2 tirets ou plus" en début de ligne.
        md = md.replace(/^\\(-{2,})/mg, '$1');
        if (!md) return [''];
        // On découpe par ligne pour le format grid table.
        // Turndown peut générer des doubles retours à la ligne pour les paragraphes.
        return md.split('\n');
      };

      // Helper pour lire l'alignement style directement (important car turndown travaille sur un clone)
      const getCellAlign = (c) => {
        const attr = c.getAttribute('align');
        if (attr === 'center' || attr === 'right' || attr === 'left') return attr;
        const a = c.style.textAlign || '';
        if (a === 'center' || a === 'right' || a === 'left') return a;
        // Fallback sur l'attribut style si le clone a perdu le style object
        const st = c.getAttribute('style') || '';
        const match = st.match(/text-align:\s*(center|right|left)/i);
        return match ? match[1].toLowerCase() : '';
      };
      const getCellValign = (c) => {
        const attr = c.getAttribute('valign');
        if (attr === 'top' || attr === 'middle' || attr === 'bottom') return attr;
        const v = c.style.verticalAlign || '';
        if (v === 'top' || v === 'middle' || v === 'bottom') return v;
        const st = c.getAttribute('style') || '';
        const match = st.match(/vertical-align:\s*(top|middle|bottom)/i);
        return match ? match[1].toLowerCase() : '';
      };

      const thead = node.querySelector('thead tr');
      const firstRow = thead || node.querySelector('tbody tr');
      const headerCells = thead
        ? [...thead.querySelectorAll('th')].map(th => ({
            paras: getCellParas(th),
            align: getCellAlign(th)
          }))
        : [];
      const valigns = firstRow
        ? [...firstRow.cells].map(c => getCellValign(c))
        : [];

      const bodyParas = [...node.querySelectorAll('tbody tr')].map(tr =>
        [...tr.querySelectorAll('td')].map(td => getCellParas(td))
      );

      const numCols = headerCells.length || (bodyParas[0] || []).length;
      
      // Récupérer les alignements des colonnes (depuis thead, ou premier rang tbody par défaut)
      const aligns = headerCells.length > 0 
        ? headerCells.map(h => h.align)
        : (firstRow ? [...firstRow.cells].map(c => getCellAlign(c)) : new Array(numCols).fill(''));

      const headerParaArrays = headerCells.map(h => h.paras);

      const colWidths = new Array(numCols).fill(5);
      headerParaArrays.forEach((paras, colIdx) => paras.forEach(p => {
        colWidths[colIdx] = Math.max(colWidths[colIdx], p.length + 2);
      }));
      bodyParas.forEach(row => row.forEach((paras, colIdx) => {
        if (colIdx < numCols) paras.forEach(p => { colWidths[colIdx] = Math.max(colWidths[colIdx], p.length + 2); });
      }));

      const makeSep = (char, includeAlign = false) => '+' + colWidths.map((w, i) => {
        if (includeAlign) {
          const a = aligns[i] || '';
          if (a === 'center') return ':' + char.repeat(w - 2) + ':';
          if (a === 'right')  return char.repeat(w - 1) + ':';
          if (a === 'left')   return ':' + char.repeat(w - 1);
        }
        return char.repeat(w);
      }).join('+') + '+';

      const makeRow = cells =>
        '|' + cells.map((t, i) => (' ' + t).padEnd(colWidths[i])).join('|') + '|';

      const buildMultilineRow = (cellParaArrays) => {
        const maxLines = Math.max(...cellParaArrays.map(p => p.length));
        const lines = [];
        for (let l = 0; l < maxLines; l++) {
          lines.push(makeRow(cellParaArrays.map(paras => paras[l] || '')));
        }
        return lines.join('\n');
      };

      // Lire les largeurs depuis data-col-widths ou depuis les <col> du colgroup
      const rawAttrWidths = node.getAttribute('data-col-widths');
      let detectedPcts = rawAttrWidths
        ? rawAttrWidths.split(',').map(w => parseFloat(w)).filter(n => !isNaN(n) && n > 0)
        : [...node.querySelectorAll('colgroup col')].map(c => parseFloat(c.style.width || '')).filter(n => !isNaN(n) && n > 0);
      const attrParts = [];
      if (detectedPcts.length === numCols) attrParts.push(`tbl-colwidths="[${detectedPcts.join(', ')}]"`);
      if (valigns.some(v => v)) attrParts.push(`tbl-colvaligns="[${valigns.map(v => v || 'top').join(', ')}]"`);
      if (node.getAttribute('data-bordered') === '1') attrParts.push('.bordered');
      if (node.getAttribute('data-striped')  === '1') attrParts.push('.striped');
      const attrBlock = attrParts.length ? `{${attrParts.join(' ')}}` : '';

      const hasHeader = headerParaArrays.length > 0 && headerParaArrays.some(p => p.length > 0);
      
      // On met l'alignement dans le premier séparateur possible si pas de header,
      // sinon dans le séparateur de header (standard Pandoc).
      let result = '\n\n' + makeSep('-', !hasHeader) + '\n';
      if (hasHeader) {
        result += buildMultilineRow(headerParaArrays) + '\n';
        result += makeSep('=', true) + '\n';
      }
      for (const row of bodyParas) {
        result += buildMultilineRow(row) + '\n';
        result += makeSep('-') + '\n';
      }
      const captionEl  = node.querySelector('caption');
      const captionTxt = captionEl ? captionEl.textContent.trim() : '';
      let suffix = '';
      if (attrBlock) {
        // Toujours ": [légende] {attr}" quand des largeurs sont définies (légende peut être vide)
        suffix = captionTxt ? `: ${captionTxt} ${attrBlock}` : `: ${attrBlock}`;
      } else if (captionTxt) {
        suffix = `: ${captionTxt}`;
      }
      return result + (suffix ? suffix + '\n\n' : '\n');
    }
  });

  turndown.addRule('figure', {
    filter: node => node.nodeName === 'FIGURE' && node.classList.contains('md-figure'),
    replacement: (_, node) => {
      const img    = node.querySelector('img');
      if (!img) return '';
      const src    = node.getAttribute('data-src') || img.getAttribute('src') || '';
      const figcap = node.querySelector('figcaption');
      const alt = figcap ? _figcapToMd(figcap) : (img.getAttribute('alt') || '');
      const align  = node.getAttribute('data-align') || 'center';
      const imgW   = parseInt(node.getAttribute('data-imgwidth') || '100');
      const wPart  = imgW < 100 ? ` width=${imgW}%` : '';
      return `\n\n![${alt}](${src}){fig-align="${align}"${wPart}}\n\n`;
    }
  });

  turndown.addRule('yaml-block', {
    filter: node => node.nodeName === 'DIV' && node.classList.contains('yaml-block'),
    replacement: (_, node) => {
      const code = node.querySelector('pre code');
      const text = code ? code.textContent : '';
      return '\n\n---\n' + text + '\n---\n\n';
    },
  });

  turndown.addRule('underline', {
    filter: ['u'],
    replacement: (content) => `[${content}]{.underline}`,
  });

  turndown.addRule('code-block-wrapper', {
    filter: node => node.nodeName === 'DIV' && node.classList.contains('code-block-wrapper'),
    replacement: (_, node) => {
      const label = node.querySelector('.code-lang-label');
      const code  = node.querySelector('pre code');
      const lang  = label ? label.textContent.trim() : 'texinfo';
      const text  = code  ? _codeElText(code) : '';
      return '\n\n```' + lang + '\n' + text + '\n```\n\n';
    },
  });

  turndown.addRule('strikethrough', {
    filter: ['s', 'del', 'strike'],
    replacement: (content) => `~~${content}~~`,
  });

  turndown.addRule('smallcaps', {
    filter: node => node.nodeName === 'SPAN' && node.classList.contains('smallcaps'),
    replacement: (content) => `[${content}]{.smallcaps}`,
  });

  turndown.addRule('citation', {
    filter: node => node.nodeName === 'SPAN' && node.classList.contains('md-citation'),
    replacement: (_, node) => '@' + (node.getAttribute('data-key') || node.textContent.replace(/^@/, '')),
  });

  // Les marques de recherche sont éphémères : les ignorer lors de la conversion
  turndown.addRule('search-highlight', {
    filter: node => node.nodeName === 'MARK' && node.classList.contains('search-hl'),
    replacement: (content) => content,
  });

  turndown.addRule('footnoteRef', {
    filter: node => node.nodeName === 'SUP' && node.classList.contains('footnote-ref'),
    replacement: (_, node) => `[^${node.getAttribute('data-label') || ''}]`,
  });

  turndown.addRule('footnoteDef', {
    filter: node => node.nodeName === 'DIV' && node.classList.contains('footnote-def'),
    replacement: (_, node) => {
      const label = node.getAttribute('data-label') || node.id.replace('fn-', '');
      const textEl = node.querySelector('.footnote-def-text');
      const text = textEl ? textEl.textContent.trim() : '';
      return `\n\n[^${label}]: ${text}`;
    },
  });

  // <br> visuels insérés à la place des md-html-literal masqués : le span caché gère le round-trip
  turndown.addRule('br-from-literal', {
    filter: node => node.nodeName === 'BR' && node.dataset && node.dataset.fromLiteral === '1',
    replacement: () => '',
  });

  turndown.addRule('tasklistcheckbox', {
    filter: node => node.nodeName === 'INPUT' && node.getAttribute('type') === 'checkbox'
                    && node.classList.contains('task-checkbox'),
    replacement: (_, node) => node.checked ? '[x]' : '[ ]',
  });

  turndown.addRule('tasklabel', {
    filter: node => node.nodeName === 'DIV' && node.classList.contains('task-label'),
    replacement: (content) => ' ' + content.trim(),
  });

  document.addEventListener('click', e => {
    if (!document.getElementById('menu-container').contains(e.target)) {
      closeMenu();
    }
    const hc = document.getElementById('heading-container');
    if (hc && !hc.contains(e.target)) {
      closeHeadingMenu();
    }
  });

  setupKeyboardShortcuts();
  setupPreviewEditing();
  _setupPasteImageListener();
  _initCiteListeners();
  _initTabsMenuClose();
  _initTabsResizeObserver();
  setupContextMenu();
  setupTableColResize();
  _setupLogMinimap();
  _setupSourceMinimap();
  _setupPumlMinimap();
  _setupYamlMinimap();
  _setupCssMinimap();
  _setupJsonMinimap();
  _setupLuaMinimap();
  _setupXmlMinimap();
  _setupPreviewWidthHandle();
  initPumlPreviewResize();
  _initPumlPreviewZoom();
  _initImageViewerZoom();
  _setupCursorTracking();
  document.querySelector('.status-sep').style.display = 'none';
  initSettings();
  renderEmpty();

  // Installer la vraie fonction (remplace le stub HTML)
  window.openInitialFile = openInitialFile;

  // Ouvrir le fichier initial s'il a été mis en file d'attente avant init()
  const queued = pendingInitialFile || window._pendingInitialFileData;
  if (queued) {
    pendingInitialFile = null;
    window._pendingInitialFileData = null;
    openInitialFile(queued);
  }

  // Raw HTML : clic droit + double-clic
  document.getElementById('preview').addEventListener('contextmenu', e => {
    const rh = e.target.closest('.md-rawhtml');
    if (!rh) return;
    e.preventDefault();
    e.stopPropagation();
    showRawHtmlContextMenu(e, rh);
  });
  document.getElementById('preview').addEventListener('dblclick', e => {
    const rh = e.target.closest('.md-rawhtml');
    if (!rh) return;
    e.preventDefault();
    openRawHtmlDialog(rh);
  });
  document.addEventListener('click', () => hideRawHtmlContextMenu());

  // Commentaire HTML : clic droit + double-clic
  document.getElementById('preview').addEventListener('contextmenu', e => {
    const cm = e.target.closest('.md-comment');
    if (!cm) return;
    e.preventDefault();
    e.stopPropagation();
    showCommentContextMenu(e, cm);
  });
  document.getElementById('preview').addEventListener('dblclick', e => {
    const cm = e.target.closest('.md-comment');
    if (!cm) return;
    e.preventDefault();
    _editCommentText(cm);
  });
  document.addEventListener('click', () => hideCommentContextMenu());

  // Commentaire (annotation) : clic sur le texte surligné → mettre en évidence
  // la carte correspondante dans le panneau « Commentaires » du volet droit.
  document.getElementById('preview').addEventListener('click', e => {
    const annot = e.target.closest('.md-annotation');
    if (!annot) return;
    const id = annot.getAttribute('data-annot-id');
    if (!id) return;
    const section = document.getElementById('annotations-section');
    if (section) section.classList.remove('collapsed');
    const item = document.querySelector(`.annot-item[data-annot-id="${id}"]`);
    if (!item) return;
    item.scrollIntoView({ block: 'nearest' });
    item.classList.add('annot-item--flash');
    setTimeout(() => item.classList.remove('annot-item--flash'), 900);
  });

  // Shortcode : clic droit + double-clic
  document.getElementById('preview').addEventListener('contextmenu', e => {
    const sc = e.target.closest('.md-shortcode');
    if (!sc) return;
    e.preventDefault();
    e.stopPropagation();
    showSCContextMenu(e, sc);
  });
  document.getElementById('preview').addEventListener('dblclick', e => {
    const sc = e.target.closest('.md-shortcode');
    if (!sc) return;
    e.preventDefault();
    openSCDialog(sc);
  });
  document.addEventListener('click', () => hideSCContextMenu());

  // Shortcode : aperçu live
  ['sc-name','sc-args'].forEach(id =>
    document.getElementById(id).addEventListener('input', _updateSCPreview));

  setupFDivContextMenu();
  setupCodeBlockContextMenu();

  // Clic droit sur une note de bas de page → menu dédié (avant le handler lien)
  document.getElementById('preview').addEventListener('contextmenu', e => {
    const supRef = e.target.closest('sup.footnote-ref');
    if (!supRef) return;
    e.preventDefault();
    e.stopPropagation();
    showFootnoteContextMenu(e, supRef);
  });
  document.addEventListener('click', () => hideFootnoteContextMenu());

  // ── Popover survol note de bas de page ────────────────────────────────────
  const fnTooltip = document.getElementById('footnote-tooltip');
  let _fnTooltipTimer = null;

  document.getElementById('preview').addEventListener('mouseover', e => {
    const sup = e.target.closest('sup.footnote-ref');
    if (!sup) return;
    clearTimeout(_fnTooltipTimer);
    const label = sup.getAttribute('data-label');
    if (!label) return;

    // Récupérer le texte de la note
    let content = '';
    if (state.sourceMode) {
      const ta = document.getElementById('source-editor');
      const m  = ta.value.match(new RegExp(`\\[\\^${label}\\]:\\s+(.+)`));
      content = m ? m[1] : '';
    } else {
      const defEl  = document.querySelector(`#preview .footnote-def[data-label="${label}"]`);
      const textEl = defEl && defEl.querySelector('.footnote-def-text');
      content = textEl ? textEl.textContent.trim() : '';
    }
    if (!content) return;

    document.getElementById('fn-tooltip-num').textContent = `[^${label}]`;
    const hasList = /^[-*+]\s|^\d+\.\s/m.test(content);
    document.getElementById('fn-tooltip-body').innerHTML = hasList
      ? marked.parse(content)
      : marked.parseInline(content);

    // Positionnement : au-dessus du sup, centré
    const rect = sup.getBoundingClientRect();
    const tw = Math.min(340, window.innerWidth - 16);
    let left = rect.left + rect.width / 2 - tw / 2;
    left = Math.max(8, Math.min(left, window.innerWidth - tw - 8));
    fnTooltip.style.width = tw + 'px';
    fnTooltip.style.left  = left + 'px';

    // Afficher au-dessus si possible, sinon en-dessous
    fnTooltip.style.visibility = 'hidden';
    fnTooltip.style.top = '0px';
    fnTooltip.classList.add('visible');
    const th = fnTooltip.offsetHeight;
    fnTooltip.classList.remove('visible');
    fnTooltip.style.visibility = '';
    const topAbove = rect.top - th - 8;
    fnTooltip.style.top = (topAbove >= 8 ? topAbove : rect.bottom + 8) + 'px';

    _fnTooltipTimer = setTimeout(() => fnTooltip.classList.add('visible'), 120);
  });

  document.getElementById('preview').addEventListener('mouseout', e => {
    const sup = e.target.closest('sup.footnote-ref');
    if (!sup) return;
    clearTimeout(_fnTooltipTimer);
    fnTooltip.classList.remove('visible');
  });

  // ── Popover survol commentaire (annotation liée à une sélection) ──────────
  const annotTooltip = document.getElementById('annot-tooltip');
  let _annotTooltipTimer = null;

  document.getElementById('preview').addEventListener('mouseover', e => {
    const span = e.target.closest('.md-annotation');
    if (!span) return;
    clearTimeout(_annotTooltipTimer);
    const text = span.getAttribute('data-annot-text') || '';
    if (!text.trim()) return;

    document.getElementById('annot-tooltip-body').textContent = text;

    // Positionnement : toujours en dessous du texte surligné, centré
    const rect = span.getBoundingClientRect();
    const tw = Math.min(320, window.innerWidth - 16);
    let left = rect.left + rect.width / 2 - tw / 2;
    left = Math.max(8, Math.min(left, window.innerWidth - tw - 8));
    annotTooltip.style.width = tw + 'px';
    annotTooltip.style.left  = left + 'px';
    annotTooltip.style.top   = (rect.bottom + 6) + 'px';

    _annotTooltipTimer = setTimeout(() => annotTooltip.classList.add('visible'), 120);
  });

  document.getElementById('preview').addEventListener('mouseout', e => {
    const span = e.target.closest('.md-annotation');
    if (!span) return;
    clearTimeout(_annotTooltipTimer);
    annotTooltip.classList.remove('visible');
  });

  // Clic droit sur un lien → menu contextuel dédié
  document.getElementById('preview').addEventListener('contextmenu', e => {
    const anchor = e.target.closest('a[href]');
    if (!anchor) return;
    // Ne pas interférer avec les notes de bas de page
    if (anchor.closest('sup.footnote-ref')) return;
    // Ne pas interférer avec le label fdiv
    if (anchor.closest('.fenced-div-label')) return;
    // Autoriser les liens dans figcaption, bloquer les autres liens dans figure
    if (anchor.closest('figure') && !anchor.closest('figcaption')) return;
    e.preventDefault();
    e.stopPropagation();
    showLinkContextMenu(e, anchor);
  });
  document.addEventListener('click', () => hideLinkContextMenu());

  // Double-clic sur un lien → ouvrir le dialogue d'édition
  document.getElementById('preview').addEventListener('dblclick', e => {
    const anchor = e.target.closest('a[href]');
    if (!anchor) return;
    e.preventDefault();
    openLinkEditDialog(anchor);
  });

  // Curseur "lien" quand Ctrl est enfoncé
  document.addEventListener('keydown', e => {
    if (e.ctrlKey || e.metaKey) document.body.classList.add('ctrl-held');
  });
  document.addEventListener('keyup', e => {
    if (!e.ctrlKey && !e.metaKey) document.body.classList.remove('ctrl-held');
  });
  window.addEventListener('blur', () => document.body.classList.remove('ctrl-held'));

  // Clic sur une case à cocher de liste de tâches → basculer l'état + barrer le texte.
  // mousedown : on toggling + preventDefault (empêche le déplacement de curseur contenteditable).
  // click     : preventDefault seul (empêche le re-toggle automatique du navigateur).
  document.getElementById('preview').addEventListener('mousedown', e => {
    if (e.target.nodeName !== 'INPUT' || e.target.getAttribute('type') !== 'checkbox') return;
    e.preventDefault();
    if (e.target.classList.contains('md-todo-checkbox')) {
      const node = e.target.closest('.md-comment--todo');
      if (!node) return;
      const isDone = node.classList.contains('md-comment--done');
      const body   = node.querySelector('.md-todo-body')?.textContent?.trim() ?? '';
      const newKind = isDone ? 'TODO' : 'DONE';
      node.classList.toggle('md-comment--done', !isDone);
      node.setAttribute('data-comment', `${newKind}: ${body}`.replace(/"/g, '&quot;'));
      node.setAttribute('data-todo-kind', newKind);
      e.target.checked = !isDone;
      const tag = node.querySelector('.md-todo-tag');
      if (tag) tag.textContent = newKind;
    } else {
      e.target.checked = !e.target.checked;
      const li = e.target.closest('li.task-item');
      if (li) li.classList.toggle('task-done', e.target.checked);
    }
    syncPreviewToContent();
  });
  document.getElementById('preview').addEventListener('click', e => {
    if (e.target.nodeName !== 'INPUT' || e.target.getAttribute('type') !== 'checkbox') return;
    e.preventDefault(); // annule le toggle automatique du navigateur (déjà fait dans mousedown)
  });

  // Ctrl+Clic sur un lien → suivre le lien
  document.getElementById('preview').addEventListener('click', e => {
    if (!e.ctrlKey && !e.metaKey) return;
    const anchor = e.target.closest('a[href]');
    if (!anchor) return;
    e.preventDefault();
    const href = anchor.getAttribute('href') || '';
    if (href.startsWith('#')) {
      // Lien interne → scroll vers le titre cible dans le preview
      const target = document.getElementById('preview').querySelector(`[id="${href.slice(1)}"]`)
                  || [...document.querySelectorAll('#preview h1,#preview h2,#preview h3,#preview h4,#preview h5,#preview h6')]
                       .find(h => _headingToId(h.textContent.trim()) === href.slice(1));
      if (target) target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } else {
      // Lien externe → ouvrir via Python (navigateur par défaut)
      window.pywebview.api.open_url(href);
    }
  });

  // Aperçu live du dialogue lien
  ['link-text','link-url','link-id'].forEach(id => {
    document.getElementById(id).addEventListener('input', _updateLinkPreview);
  });
  document.getElementById('link-heading-select').addEventListener('change', _updateLinkPreview);

  // Tab dans l'éditeur source : tab stop dans un bloc de code, indent liste ailleurs
  // Shift+Enter dans l'éditeur source : insérer <br> + saut de ligne
  document.getElementById('source-editor').addEventListener('keydown', e => {
    const ta = document.getElementById('source-editor');
    if (e.key === 'Enter' && e.shiftKey) {
      e.preventDefault();
      ta.setRangeText('<br>\n\n', ta.selectionStart, ta.selectionEnd, 'end');
      onSourceInput();
      return;
    }
    if (e.key === 'Tab') {
      e.preventDefault();
      if (_taListIndent(ta, e)) {
        onSourceInput();
      } else if (_taCursorInCodeFence(ta)) {
        // Dans un bloc de code fencé : insérer/supprimer un vrai \t
        if (e.shiftKey) {
          const pos = ta.selectionStart;
          if (pos > 0 && ta.value[pos - 1] === '\t') {
            ta.setRangeText('', pos - 1, pos, 'end');
          }
        } else {
          ta.setRangeText('\t', ta.selectionStart, ta.selectionEnd, 'end');
        }
        onSourceInput();
      } else {
        _taTabStop(ta, e.shiftKey);
        onSourceInput();
      }
      return;
    }
  });

  // Mise à jour de la visibilité de la règle sur tout déplacement de curseur
  // selectionchange est l'événement le plus fiable (inclut clavier, souris, focus)
  document.addEventListener('selectionchange', () => {
    _updateRulerVisibility();
    _updateTOCActive();
  });

  // Éditeurs texte spécialisés — 'input' + 'keyup' pour compatibilité pywebview
  [
    ['yaml-editor', onYamlInput],
    ['css-editor',  onCssInput],
    ['lua-editor',  onLuaInput],
    ['json-editor', onJsonInput],
    ['puml-editor', onPumlInput],
    ['xml-editor',  onXmlInput],
  ].forEach(([id, fn]) => {
    const ta = document.getElementById(id);
    if (!ta) return;
    ta.addEventListener('input',   fn);
    ta.addEventListener('keyup',   fn);
    ta.addEventListener('cut',     fn);
    ta.addEventListener('paste',   fn);
    // Tab / Shift+Tab : indentation par tabulation plutôt que changement de focus
    ta.addEventListener('keydown', e => {
      if (e.key !== 'Tab') return;
      e.preventDefault();
      _taCodeTab(ta, e);
      fn();
    });
  });
  document.getElementById('puml-editor').addEventListener('contextmenu', showPumlContextMenu);

  // Fermer le menu PUML sur clic hors menu
  document.addEventListener('mousedown', e => {
    if (!e.target.closest('#puml-context-menu')) hidePumlContextMenu();
  });

  // Initialiser le volet latéral gauche
  initSidebarResize();
  initCompileResize();

  // Initialiser le sommaire (droite)
  initTOCResize();
  initTocMetaResize();
  initMetaAnnotResize();
  initTOCScrollSpy();

  // Initialiser la règle de tabulation
  initTabRuler();

  // Synchroniser le décalage de la barre d'onglets avec la sidebar
  _initTabAreaOffsetSync();

  _hideSplashScreen();
}

// ── Splash screen ─────────────────────────────────────────────────────────────
// Affiché un court instant au démarrage, puis disparaît en fondu une fois
// l'initialisation terminée (durée minimale pour rester lisible).
const _SPLASH_MIN_DURATION = 1000;
const _splashStartTime = Date.now();
function _hideSplashScreen() {
  const splash = document.getElementById('splash-screen');
  if (!splash) return;
  const elapsed = Date.now() - _splashStartTime;
  setTimeout(() => {
    splash.classList.add('hidden');
    setTimeout(() => splash.remove(), 450);
  }, Math.max(0, _SPLASH_MIN_DURATION - elapsed));
}

// ── Tabs ──────────────────────────────────────────────────────────────────────
function getActiveTab() {
  return state.tabs.find(t => t.id === state.activeTabId) || null;
}

function createTab(name, path, content) {
  const s = loadSettings();
  const defaultZoom = s.zoomDefault || 100;
  return state.tabs[state.tabs.push({
    id: state.nextTabId++,
    name: name || 'Sans titre',
    path: path || null,
    content: content || '',
    savedContent: content || '',
    modified: false,
    zoom: defaultZoom,
    imageZoom: 1.0,
  }) - 1];
}

const _IMAGE_EXTS = new Set(['.png','.jpg','.jpeg','.gif','.bmp','.webp','.svg','.ico']);

function switchToTab(id) {
  // Sauvegarder le zoom de l'onglet courant avant de changer d'onglet
  const _leavingTab = getActiveTab();
  if (_leavingTab) {
    _leavingTab.zoom = parseInt(document.getElementById('zoom-slider').value) || 100;
    _leavingTab.imageZoom = _imageZoom;
  }

  state.activeTabId = id;
  _searchLastQuery = '';   // forcer une nouvelle recherche dans le nouvel onglet
  const tab = getActiveTab();
  if (tab) {
    const p   = tab.path ? tab.path.toLowerCase() : '';
    const ext = p.lastIndexOf('.') >= 0 ? p.slice(p.lastIndexOf('.')) : '';

    const isHtml  = ext === '.html' || ext === '.htm';
    const isPdf   = ext === '.pdf';
    const isImage = _IMAGE_EXTS.has(ext);

    const htmlViewer  = document.getElementById('html-viewer');
    const pdfViewer   = document.getElementById('pdf-viewer');
    const imageViewer = document.getElementById('image-viewer');

    const fileUrl = tab.path ? 'file:///' + tab.path.replace(/\\/g, '/') : '';

    if (isImage) {
      document.body.classList.add('image-mode');
      document.body.classList.remove('html-mode', 'pdf-mode');
      if (imageViewer) { imageViewer.src = fileUrl; _imageZoom = tab.imageZoom ?? 1.0; _applyImageZoom(); }
      if (htmlViewer)  htmlViewer.src  = 'about:blank';
      if (pdfViewer)   pdfViewer.src   = 'about:blank';
    } else if (isHtml) {
      document.body.classList.add('html-mode');
      document.body.classList.remove('pdf-mode', 'image-mode');
      if (htmlViewer)  htmlViewer.src  = fileUrl || 'about:blank';
      if (pdfViewer)   pdfViewer.src   = 'about:blank';
      if (imageViewer) imageViewer.src = '';
    } else if (isPdf) {
      document.body.classList.add('pdf-mode');
      document.body.classList.remove('html-mode', 'image-mode');
      if (pdfViewer)   pdfViewer.src   = fileUrl || 'about:blank';
      if (htmlViewer)  htmlViewer.src  = 'about:blank';
      if (imageViewer) imageViewer.src = '';
    } else {
      document.body.classList.remove('html-mode', 'pdf-mode', 'image-mode');
      if (htmlViewer)  htmlViewer.src  = 'about:blank';
      if (pdfViewer)   pdfViewer.src   = 'about:blank';
      if (imageViewer) imageViewer.src = '';
      const isYaml = ext === '.yml' || ext === '.yaml';
      const isLog  = ext === '.log';
      const isTex  = ext === '.tex';
      const isPuml = ext === '.puml' || ext === '.plantuml';
      const isCss  = ext === '.css' || ext === '.scss' || ext === '.less';
      const isLua  = ext === '.lua';
      const isJson = ext === '.json';
      const isXml  = ext === '.xml';
      if (isPuml) {
        document.body.classList.add('puml-mode');
        document.body.classList.remove('yaml-mode', 'log-mode', 'css-mode', 'lua-mode', 'json-mode', 'xml-mode');
        document.getElementById('puml-editor').value = tab.content;
        _editorPollingSync('puml-editor', tab.content);
        _updatePumlHighlight();
        setTimeout(() => { _drawPumlMinimap(); document.getElementById('puml-editor').focus(); }, 0);
      } else if (isYaml) {
        document.body.classList.add('yaml-mode');
        document.body.classList.remove('log-mode', 'puml-mode', 'css-mode', 'lua-mode', 'json-mode', 'xml-mode');
        document.getElementById('yaml-editor').value = tab.content;
        _editorPollingSync('yaml-editor', tab.content);
        _updateYamlHighlight();
        setTimeout(() => { _drawYamlMinimap(); document.getElementById('yaml-editor').focus(); }, 0);
      } else if (isCss) {
        document.body.classList.add('css-mode');
        document.body.classList.remove('yaml-mode', 'log-mode', 'puml-mode', 'lua-mode', 'json-mode', 'xml-mode');
        document.getElementById('css-editor').value = tab.content;
        _editorPollingSync('css-editor', tab.content);
        _updateCssHighlight();
        setTimeout(() => { _drawCssMinimap(); document.getElementById('css-editor').focus(); }, 0);
      } else if (isLua) {
        document.body.classList.add('lua-mode');
        document.body.classList.remove('yaml-mode', 'log-mode', 'puml-mode', 'css-mode', 'json-mode', 'xml-mode');
        document.getElementById('lua-editor').value = tab.content;
        _editorPollingSync('lua-editor', tab.content);
        _updateLuaHighlight();
        setTimeout(() => { _drawLuaMinimap(); document.getElementById('lua-editor').focus(); }, 0);
      } else if (isJson) {
        document.body.classList.add('json-mode');
        document.body.classList.remove('yaml-mode', 'log-mode', 'puml-mode', 'css-mode', 'lua-mode', 'xml-mode');
        document.getElementById('json-editor').value = tab.content;
        _editorPollingSync('json-editor', tab.content);
        _updateJsonHighlight();
        updateTOC();
        setTimeout(() => { _drawJsonMinimap(); document.getElementById('json-editor').focus(); }, 0);
      } else if (isXml) {
        document.body.classList.add('xml-mode');
        document.body.classList.remove('yaml-mode', 'log-mode', 'puml-mode', 'css-mode', 'lua-mode', 'json-mode');
        document.getElementById('xml-editor').value = tab.content;
        _editorPollingSync('xml-editor', tab.content);
        _updateXmlHighlight();
        setTimeout(() => { _drawXmlMinimap(); document.getElementById('xml-editor').focus(); }, 0);
      } else if (isLog || isTex) {
        document.body.classList.add('log-mode');
        document.body.classList.remove('yaml-mode', 'puml-mode', 'css-mode', 'lua-mode', 'json-mode', 'xml-mode');
        renderLogPane(tab.content, isTex ? 'tex' : 'log');
      } else {
        document.body.classList.remove('yaml-mode', 'log-mode', 'puml-mode', 'css-mode', 'lua-mode', 'json-mode', 'xml-mode');
        if (state.sourceMode) {
          document.getElementById('source-editor').value = tab.content;
          _updateSourceHighlight();
        } else {
          updatePreview(tab.content);
        }
      }
    }
    const zoomVal = tab.zoom || 100;
    document.getElementById('zoom-slider').value = zoomVal;
    setZoom(zoomVal);
    updateFileStatus(tab);
  }
  renderTabList();
  _updatePreviewWidthHandlePos();
}

function closeTab(id) {
  const tab = state.tabs.find(t => t.id === id);
  if (!tab) return;
  if (tab.modified) {
    if (!confirm(`Fermer "${tab.name}" sans enregistrer les modifications ?`)) return;
  }
  const idx = state.tabs.findIndex(t => t.id === id);
  state.tabs.splice(idx, 1);

  if (state.activeTabId === id) {
    if (state.tabs.length > 0) {
      switchToTab(state.tabs[Math.min(idx, state.tabs.length - 1)].id);
    } else {
      state.activeTabId = null;
      renderEmpty();
    }
  }
  renderTabList();
}

function renderTabList() {
  const list = document.getElementById('tab-list');
  list.innerHTML = '';
  state.tabs.forEach(tab => {
    const el = document.createElement('div');
    el.className = 'tab' + (tab.id === state.activeTabId ? ' active' : '');
    el.title = tab.path || tab.name;

    const nameSpan = document.createElement('span');
    nameSpan.className = 'tab-name';
    nameSpan.textContent = (tab.modified ? '● ' : '') + tab.name;

    const closeBtn = document.createElement('button');
    closeBtn.className = 'tab-close';
    closeBtn.textContent = '×';
    closeBtn.title = 'Fermer';
    closeBtn.addEventListener('click', e => { e.stopPropagation(); closeTab(tab.id); });

    el.appendChild(nameSpan);
    el.appendChild(closeBtn);
    el.addEventListener('click', () => switchToTab(tab.id));
    el.addEventListener('contextmenu', e => { e.preventDefault(); showTabContextMenu(e, tab); });
    list.appendChild(el);
  });

  _updateTabWidths();
  const active = list.querySelector('.tab.active');
  if (active) active.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  _updateSidebarActiveItem();
  _updateTodoPanelBtn();
}

// ── Menu contextuel onglet ────────────────────────────────────────────────────
let _tabContextTarget = null;

function showTabContextMenu(e, tab) {
  _tabContextTarget = tab;
  const menu = document.getElementById('tab-context-menu');
  const p = tab.path ? tab.path.toLowerCase() : '';
  const canCompile = p.endsWith('.qmd') || p.endsWith('.md');
  document.getElementById('tab-ctx-compile').style.display = canCompile ? '' : 'none';
  const hasSep = canCompile;
  document.getElementById('tab-ctx-sep-file').style.display = hasSep ? '' : 'none';
  document.getElementById('tab-ctx-copy-name').style.display = tab.path ? '' : 'none';
  document.getElementById('tab-ctx-rename').style.display = tab.path ? '' : 'none';
  _showMenuAt(menu, e.clientX, e.clientY, m => m.classList.add('visible'));
}

let _fileContextItem = null;

async function showSidebarFileContextMenu(e, item) {
  _fileContextItem = item;
  const canCompile = item.path.toLowerCase().endsWith('.qmd');
  document.getElementById('sfctx-compile').style.display     = canCompile ? '' : 'none';
  document.getElementById('sfctx-compile-sep').style.display = canCompile ? '' : 'none';

  const isPuml = item.path.toLowerCase().endsWith('.puml') || item.path.toLowerCase().endsWith('.plantuml');
  const pumlDisplay = isPuml ? '' : 'none';
  document.getElementById('sfctx-puml-both').style.display = pumlDisplay;
  document.getElementById('sfctx-puml-svg').style.display  = pumlDisplay;
  document.getElementById('sfctx-puml-png').style.display  = pumlDisplay;
  document.getElementById('sfctx-puml-sep').style.display  = pumlDisplay;

  const canCompilePdf = canCompile;
  const btnPdf = document.getElementById('sfctx-compile-pdf');
  const sepPdf = document.getElementById('sfctx-compile-pdf-sep');
  if (btnPdf) btnPdf.style.display = canCompilePdf ? '' : 'none';
  if (sepPdf) sepPdf.style.display = canCompilePdf ? '' : 'none';

  const isXml = item.path.toLowerCase().endsWith('.xml');
  const btnXml = document.getElementById('sfctx-xml-convert');
  const sepXml = document.getElementById('sfctx-xml-sep');
  if (btnXml) btnXml.style.display = isXml ? '' : 'none';
  if (sepXml) sepXml.style.display = isXml ? '' : 'none';

  const menu = document.getElementById('sidebar-file-context-menu');
  _showMenuAt(menu, e.clientX, e.clientY, m => m.classList.add('visible'));
}

function hideSidebarFileContextMenu() {
  document.getElementById('sidebar-file-context-menu').classList.remove('visible');
  _fileContextItem = null;
}

async function sidebarFileCtxCompile() {
  const item = _fileContextItem;
  hideSidebarFileContextMenu();
  if (!item || !window.pywebview) return;
  const settings = loadSettings();
  const projectDir = settings.projectDir || '';
  openCompilePanel(item.name);
  await window.pywebview.api.compile_quarto(item.path, projectDir, settings.compileCmdHtml || '');
}

async function sidebarFileCtxCompilePdf() {
  const item = _fileContextItem;
  hideSidebarFileContextMenu();
  if (!item || !window.pywebview) return;
  const settings = loadSettings();
  const projectDir = settings.projectDir || '';
  openCompilePanel(item.name + ' (PDF)');
  await window.pywebview.api.compile_quarto_pdf(item.path, projectDir, settings.compileCmdPdf || '');
}

async function sidebarFileCtxCompilePuml(fmt) {
  const item = _fileContextItem;
  hideSidebarFileContextMenu();
  if (!item || !window.pywebview) return;
  openCompilePanel(item.name);
  await window.pywebview.api.compile_puml(item.path, fmt || 'both');
}

function sidebarFileCtxOpen() {
  const item = _fileContextItem;
  hideSidebarFileContextMenu();
  if (item) sidebarOpenFile(item);
}

function sidebarFileCtxRename() {
  const item = _fileContextItem;
  hideSidebarFileContextMenu();
  if (!item) return;
  _tabContextTarget = { path: item.path, name: item.name };
  tabCtxRenameFile();
}

function sidebarFileCtxCopyName() {
  const item = _fileContextItem;
  hideSidebarFileContextMenu();
  if (!item) return;
  navigator.clipboard.writeText(item.name).then(() => showToast('Nom copié'));
}

async function sidebarFileCtxDuplicate() {
  const item = _fileContextItem;
  hideSidebarFileContextMenu();
  if (!item || !window.pywebview) return;
  const res = await window.pywebview.api.duplicate_file(item.path);
  if (res.error) { showToast('Erreur : ' + res.error); return; }
  showToast('Fichier dupliqué : ' + res.name);
  if (_sidebarCurrentDir) loadSidebarDir(_sidebarCurrentDir);
}

function sidebarFileCtxDelete() {
  const item = _fileContextItem;
  hideSidebarFileContextMenu();
  if (!item) return;
  const overlay = document.getElementById('delete-confirm-overlay');
  overlay.dataset.path = item.path;
  overlay.dataset.type = 'file';
  document.getElementById('delete-confirm-title').textContent = 'Supprimer le fichier ?';
  document.getElementById('delete-confirm-name').textContent = item.name;
  document.getElementById('delete-confirm-warn').textContent = 'Cette action est irréversible.';
  overlay.classList.add('open');
}

function sidebarFileCtxConvertXml() {
  const item = _fileContextItem;
  hideSidebarFileContextMenu();
  if (!item) return;
  openXmlConvertDialog(item.path);
}

// ── Boîte de dialogue Conversion XML → JSON ───────────────────────────────

function openXmlConvertDialog(xmlPath) {
  document.getElementById('xml-convert-xml-path').value = xmlPath || '';
  document.getElementById('xml-convert-xsd-path').value = '';
  document.getElementById('xml-convert-log').value = '';
  document.getElementById('xml-convert-btn').disabled = false;
  document.getElementById('xml-convert-overlay').classList.add('open');
}

function closeXmlConvertDialog() {
  document.getElementById('xml-convert-overlay').classList.remove('open');
}

function closeXmlConvertOnOverlay(e) {
  if (e.target === document.getElementById('xml-convert-overlay')) closeXmlConvertDialog();
}

async function xmlConvertBrowseXsd() {
  if (!window.pywebview || !window.pywebview.api) return;
  const s = loadSettings();
  const initialDir = s.xsdDefaultDir || '';
  const path = await window.pywebview.api.browse_xsd_for_xml(initialDir);
  if (path) document.getElementById('xml-convert-xsd-path').value = path;
}

function xmlConvertClearXsd() {
  document.getElementById('xml-convert-xsd-path').value = '';
}

async function runXmlConvert() {
  if (!window.pywebview || !window.pywebview.api) return;
  const xmlPath = document.getElementById('xml-convert-xml-path').value.trim();
  const xsdPath = document.getElementById('xml-convert-xsd-path').value.trim();
  if (!xmlPath) { showToast('Chemin XML manquant'); return; }
  const logEl = document.getElementById('xml-convert-log');
  logEl.value = '';
  document.getElementById('xml-convert-btn').disabled = true;
  document.getElementById('xml-convert-open-md').style.display = 'none';
  await window.pywebview.api.run_xml_to_json(xmlPath, xsdPath);
}

let _xmlConvertMdPath = null;

window.xmlConvertLog = function(msg) {
  const logEl = document.getElementById('xml-convert-log');
  if (!logEl) return;
  logEl.value += msg + '\n';
  logEl.scrollTop = logEl.scrollHeight;
};

window.xmlConvertFinished = function(ok, mdPath) {
  const btn = document.getElementById('xml-convert-btn');
  if (btn) btn.disabled = false;
  _xmlConvertMdPath = mdPath || null;
  const openMdBtn = document.getElementById('xml-convert-open-md');
  if (openMdBtn) openMdBtn.style.display = (!ok && mdPath) ? '' : 'none';
  if (ok) showToast('Conversion terminée avec succès !');
  else showToast('Erreur lors de la conversion — voir le journal');
};

function xmlConvertOpenMd() {
  if (!_xmlConvertMdPath) return;
  const name = _xmlConvertMdPath.replace(/\\/g, '/').split('/').pop();
  sidebarOpenFile({ path: _xmlConvertMdPath, name });
}

function cancelDeleteFile(e) {
  if (e && e.target !== document.getElementById('delete-confirm-overlay')) return;
  document.getElementById('delete-confirm-overlay').classList.remove('open');
}

// ── Menu contextuel répertoire sidebar ────────────────────────────────────────
let _dirContextTarget = null;

function showSidebarDirContextMenu(e, item) {
  _dirContextTarget = item;
  const menu = document.getElementById('sidebar-dir-context-menu');
  _showMenuAt(menu, e.clientX, e.clientY, m => m.classList.add('visible'));
}

function hideSidebarDirContextMenu() {
  document.getElementById('sidebar-dir-context-menu').classList.remove('visible');
  _dirContextTarget = null;
}

async function sidebarCreateDir(parentPath) {
  const base = parentPath || _sidebarCurrentDir;
  if (!base) return;
  const name = prompt('Nom du nouveau répertoire :', '');
  if (!name || !name.trim()) return;
  if (!window.pywebview) return;
  const res = await window.pywebview.api.create_directory(base, name.trim());
  if (res.error) { showToast('Erreur : ' + res.error); return; }
  showToast('Répertoire créé');
  loadSidebarDir(_sidebarCurrentDir);
}

function sidebarDirCtxCreateSubdir() {
  const item = _dirContextTarget;
  hideSidebarDirContextMenu();
  if (!item || !item.path) return;
  sidebarCreateDir(item.path);
}

async function compileDirQuarto() {
  const item = _dirContextTarget;
  hideSidebarDirContextMenu();
  if (!item || !item.path) return;
  if (!window.pywebview || !window.pywebview.api) return;
  openCompilePanel(item.name + '/*.qmd');
  await window.pywebview.api.compile_directory(item.path, _sidebarCurrentDir);
}

function sidebarDirCtxDelete() {
  const item = _dirContextTarget;
  hideSidebarDirContextMenu();
  if (!item || !item.path) return;
  const overlay = document.getElementById('delete-confirm-overlay');
  overlay.dataset.path = item.path;
  overlay.dataset.type = 'dir';
  document.getElementById('delete-confirm-title').textContent = 'Supprimer le répertoire ?';
  document.getElementById('delete-confirm-name').textContent = item.name;
  document.getElementById('delete-confirm-warn').textContent =
    'Tout le contenu sera supprimé. Cette action est irréversible.';
  overlay.classList.add('open');
}

async function confirmDeleteFile() {
  const overlay = document.getElementById('delete-confirm-overlay');
  const path = overlay.dataset.path;
  const type = overlay.dataset.type;
  overlay.classList.remove('open');
  if (!path || !window.pywebview) return;

  let res;
  if (type === 'dir') {
    res = await window.pywebview.api.delete_dir(path);
  } else {
    res = await window.pywebview.api.delete_file(path);
    if (!res.error) {
      const tab = state.tabs.find(t => t.path && t.path.replace(/\\/g, '/') === path.replace(/\\/g, '/'));
      if (tab) closeTab(tab.id);
    }
  }
  if (res.error) { showToast('Erreur : ' + res.error); return; }
  showToast(type === 'dir' ? 'Répertoire supprimé' : 'Fichier supprimé');
  if (_sidebarCurrentDir) loadSidebarDir(_sidebarCurrentDir);
}

function _positionContextMenu(menu, x, y) {
  menu.style.left = x + 'px';
  menu.style.top  = y + 'px';
}

// Corrige la position des sous-menus CSS qui déborderaient du bas ou de la droite de l'écran
document.addEventListener('mouseenter', e => {
  if (!e.target || typeof e.target.closest !== 'function') return;
  const item = e.target.closest('.ctx-has-sub');
  if (!item) return;
  const sub = item.querySelector(':scope > .ctx-sub-menu');
  if (!sub) return;

  // Mesure les dimensions du sous-menu sans le rendre visible
  sub.style.setProperty('display', 'flex', 'important');
  sub.style.setProperty('visibility', 'hidden', 'important');
  const subH = sub.offsetHeight;
  const subW = sub.offsetWidth;
  sub.style.removeProperty('display');
  sub.style.removeProperty('visibility');

  const rect = item.getBoundingClientRect();

  // Vertical : aligner en haut par défaut, décaler vers le haut si débordement bas
  const spaceBelow = window.innerHeight - 8 - rect.top;
  sub.style.top    = (subH <= spaceBelow ? 0 : spaceBelow - subH) + 'px';
  sub.style.bottom = 'auto';

  // Horizontal : à droite par défaut, basculer à gauche si débordement droit
  if (rect.right + subW > window.innerWidth - 8) {
    sub.style.left  = 'auto';
    sub.style.right = '100%';
  } else {
    sub.style.left  = '100%';
    sub.style.right = 'auto';
  }
}, true);

function _showMenuAt(menu, x, y, showFn) {
  menu.style.visibility = 'hidden';
  showFn(menu);
  const w = menu.offsetWidth;
  const h = menu.offsetHeight;
  menu.style.left       = Math.max(4, Math.min(x, window.innerWidth  - w - 8)) + 'px';
  menu.style.top        = Math.max(4, Math.min(y, window.innerHeight - h - 8)) + 'px';
  menu.style.visibility = '';
}

function hideTabContextMenu() {
  document.getElementById('tab-context-menu').classList.remove('visible');
  _tabContextTarget = null;
}

async function compileTab() {
  const tab = _tabContextTarget;
  hideTabContextMenu();
  if (!tab || !tab.path) return;
  if (!window.pywebview || !window.pywebview.api) return;

  const s = loadSettings();
  const projectDir = s.projectDir || null;

  openCompilePanel(tab.name);
  await window.pywebview.api.compile_quarto(tab.path, projectDir, s.compileCmdHtml || '');
}

function tabCtxCopyName() {
  const target = _tabContextTarget;
  hideTabContextMenu();
  if (!target || !target.name) return;
  navigator.clipboard.writeText(target.name).catch(() => {});
}

async function tabCtxRenameFile() {
  const target = _tabContextTarget;
  hideTabContextMenu();
  if (!target || !target.path) return;
  const currentName = target.name || target.path.replace(/\\/g, '/').split('/').pop();
  const newName = window.prompt('Nouveau nom du fichier :', currentName);
  if (!newName || newName === currentName) return;
  if (!window.pywebview || !window.pywebview.api) return;
  const result = await window.pywebview.api.rename_file(target.path, newName);
  if (!result || result.error) { alert('Erreur : ' + (result?.error || 'inconnue')); return; }
  // Mettre à jour l'onglet si le fichier est ouvert
  const normOld = target.path.replace(/\\/g, '/');
  const tab = state.tabs.find(t => t.path && t.path.replace(/\\/g, '/') === normOld);
  if (tab) {
    tab.path = result.new_path;
    tab.name = newName;
    renderTabList();
    const active = getActiveTab();
    if (active) updateFileStatus(active);
  }
  // Rafraîchir la barre latérale
  if (_sidebarCurrentDir) loadSidebarDir(_sidebarCurrentDir);
}

// ── Compilation du projet global ──────────────────────────────────────────────
function toggleCompileMenu(e) {
  if (e) e.stopPropagation();
  document.getElementById('toolbar-compile-menu').classList.toggle('open');
}

function toggleViewMenu(e) {
  if (e) e.stopPropagation();
  document.getElementById('toolbar-view-menu').classList.toggle('open');
}

async function compileProject(format) {
  document.getElementById('toolbar-compile-menu')?.classList.remove('open');
  if (!window.pywebview || !window.pywebview.api) return;
  const s = loadSettings();
  const projectDir = s.projectDir || null;
  if (!projectDir) {
    alert('Aucun répertoire de projet configuré.\nDéfinissez-le dans les Paramètres.');
    return;
  }
  openCompilePanel(`Projet — ${format.toUpperCase()}`);
  await window.pywebview.api.compile_project(format, projectDir);
}

function togglePumlMenu(e) {
  if (e) e.stopPropagation();
  document.getElementById('toolbar-puml-menu').classList.toggle('open');
}

async function compilePuml(fmt) {
  document.getElementById('toolbar-puml-menu')?.classList.remove('open');
  if (!window.pywebview || !window.pywebview.api) return;
  const tab = getActiveTab();
  if (!tab || !tab.path) {
    alert('Enregistrez d\'abord le fichier avant de le compiler.');
    return;
  }
  openCompilePanel(tab.name);
  await window.pywebview.api.compile_puml(tab.path, fmt || 'both');
}

async function choosePlantumlJar() {
  if (!window.pywebview || !window.pywebview.api) return;
  const path = await window.pywebview.api.browse_plantuml_jar();
  if (path) document.getElementById('plantuml-jar').value = path;
}

async function chooseBibFile() {
  if (!window.pywebview || !window.pywebview.api) return;
  const path = await window.pywebview.api.browse_bib_file();
  if (path) {
    document.getElementById('bib-file').value = path.replace(/\\/g, '/');
    _bibEntries = null; // forcer le rechargement au prochain @
  }
}

async function chooseXsdDefaultDir() {
  if (!window.pywebview || !window.pywebview.api) return;
  const path = await window.pywebview.api.choose_directory();
  if (path) document.getElementById('xsd-default-dir').value = path.replace(/\\/g, '/');
}

// Fermer le menu compile si clic ailleurs
document.addEventListener('click', () => {
  document.getElementById('toolbar-compile-menu')?.classList.remove('open');
  document.getElementById('toolbar-view-menu')?.classList.remove('open');
  document.getElementById('toolbar-puml-menu')?.classList.remove('open');
});

// ── Panneau de compilation ────────────────────────────────────────────────────
function initCompileResize() {
  const handle = document.getElementById('compile-resize-handle');
  const panel  = document.getElementById('compile-panel');
  let dragging = false, startY = 0, startH = 0;

  handle.addEventListener('mousedown', e => {
    dragging = true;
    startY   = e.clientY;
    startH   = panel.offsetHeight;
    handle.classList.add('dragging');
    document.body.style.cursor     = 'row-resize';
    document.body.style.userSelect = 'none';
  });
  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    const newH = Math.max(80, Math.min(window.innerHeight * 0.8, startH - (e.clientY - startY)));
    panel.style.height = newH + 'px';
  });
  document.addEventListener('mouseup', () => {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove('dragging');
    document.body.style.cursor     = '';
    document.body.style.userSelect = '';
  });
}

function openCompilePanel(fileName) {
  const panel  = document.getElementById('compile-panel');
  const status = document.getElementById('compile-status');
  const output = document.getElementById('compile-output');
  document.getElementById('compile-title').textContent =
    'Compilation' + (fileName ? ' — ' + fileName : '');
  output.innerHTML = '';
  status.textContent = 'en cours…';
  status.className = 'running';
  panel.classList.add('open');
}

function closeCompilePanel() {
  document.getElementById('compile-panel').classList.remove('open');
}

function clearCompilePanel() {
  document.getElementById('compile-output').innerHTML = '';
}

// Appelé par Python via evaluate_js pour chaque ligne de sortie
window.appendCompileOutput = function(line) {
  const output = document.getElementById('compile-output');
  const span = document.createElement('span');
  const lower = line.toLowerCase();
  if (/error|erreur/i.test(lower))   span.className = 'compile-line-error';
  else if (/warn|avertissement/i.test(lower)) span.className = 'compile-line-warn';
  else if (/^\[.*\] success|✓|done/i.test(lower)) span.className = 'compile-line-ok';
  span.textContent = line + '\n';
  output.appendChild(span);
  output.scrollTop = output.scrollHeight;
};

// Appelé par Python en fin de compilation
window.compileFinished = function(returnCode, errorMsg) {
  const status = document.getElementById('compile-status');
  if (returnCode === 0) {
    status.textContent = '✓ succès';
    status.className = 'success';
    // Rafraîchir l'aperçu PUML si le panneau est ouvert
    if (document.body.classList.contains('puml-mode')) {
      _pumlRefreshPreviewAfterCompile();
    }
  } else {
    status.textContent = errorMsg ? '✗ ' + errorMsg : '✗ échec (code ' + returnCode + ')';
    status.className = 'error';
  }
};

// ── Empty state ───────────────────────────────────────────────────────────────
function renderEmpty() {
  // Si on est en mode source, revenir au mode preview
  if (state.sourceMode) {
    state.sourceMode = false;
    const editor = document.getElementById('source-editor');
    editor.removeEventListener('input', onSourceInput);
    document.getElementById('source-pane').style.display  = 'none';
    document.getElementById('preview-pane').style.display = '';
  }

  // Vider la textarea source
  const editor = document.getElementById('source-editor');
  if (editor) {
    editor.value = '';
    _updateSourceLineNumbers();
    _drawSourceMinimap();
  }

  // Vider la visionneuse image
  const imageViewer = document.getElementById('image-viewer');
  if (imageViewer) imageViewer.src = '';
  document.body.classList.remove('image-mode');

  // Vider l'éditeur CSS
  const cssEditor = document.getElementById('css-editor');
  const cssHighlight = document.getElementById('css-highlight');
  if (cssEditor) cssEditor.value = '';
  if (cssHighlight) cssHighlight.innerHTML = '';
  document.body.classList.remove('css-mode');

  // Vider l'éditeur Lua
  const luaEditor = document.getElementById('lua-editor');
  const luaHighlight = document.getElementById('lua-highlight');
  if (luaEditor) luaEditor.value = '';
  if (luaHighlight) luaHighlight.innerHTML = '';
  document.body.classList.remove('lua-mode');

  // Vider l'éditeur XML
  const xmlEditor = document.getElementById('xml-editor');
  const xmlHighlight = document.getElementById('xml-highlight');
  if (xmlEditor) xmlEditor.value = '';
  if (xmlHighlight) xmlHighlight.innerHTML = '';
  document.body.classList.remove('xml-mode');

  // Vider l'éditeur PlantUML
  const pumlEditor = document.getElementById('puml-editor');
  const pumlHighlight = document.getElementById('puml-highlight');
  if (pumlEditor) pumlEditor.value = '';
  if (pumlHighlight) pumlHighlight.innerHTML = '';
  document.body.classList.remove('puml-mode');

  const preview = document.getElementById('preview');
  preview.contentEditable = 'false';
  preview.innerHTML = `
    <div class="empty-state">
      <div class="empty-state-header">
        <img src="splash-screen.jpeg" class="empty-state-splash" alt="YUQI">
        <p class="empty-state-subtitle">Your Unofficial Quarto IDE</p>
      </div>
      <p>Ouvrez un fichier ou créez un nouvel onglet</p>
      <p><kbd>Ctrl+O</kbd> Ouvrir &nbsp; <kbd>Ctrl+T</kbd> Nouveau</p>
      <div class="empty-state-features">
        ${[
          ['M9 5H7a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2M9 5a2 2 0 0 0 2 2h2a2 2 0 0 0 2-2M9 5a2 2 0 0 0 2-2h2a2 2 0 0 0 2 2', 'Multi-onglets', '.md, .qmd, .html\u2026'],
          ['M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z M12 9a3 3 0 1 0 0 6 3 3 0 0 0 0-6z', 'Pr\xe9visualisation', 'Rendu Markdown en direct'],
          ['M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z M14 2v6h6', 'Quarto', 'Compilation HTML / PDF'],
          ['M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z M9 22V12h6v10', 'Explorateur', 'Navigation & recherche'],
          ['M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71 M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71', 'PlantUML', 'Diagrammes UML int\xe9gr\xe9s'],
          ['M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7 M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z', '\xc9dition avanc\xe9e', 'Tableaux, YAML, shortcodes'],
        ].map(([path, title, desc]) => `
        <div class="empty-state-feature">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
            <path d="${path}"/>
          </svg>
          <span><strong style="color:var(--text);font-weight:500">${title}</strong><br>${desc}</span>
        </div>`).join('')}
      </div>
    </div>
  `;
  document.getElementById('status-file').textContent = 'Aucun fichier ouvert';
  document.getElementById('status-filetype').textContent = '';
  document.getElementById('status-cursor').textContent = '';
  const sep = document.querySelector('.status-sep');
  if (sep) sep.style.display = 'none';
  const tocBody = document.getElementById('toc-body');
  if (tocBody) tocBody.innerHTML = '<div class="toc-empty">Aucun document</div>';
  const annotBody = document.getElementById('annotations-body');
  if (annotBody) annotBody.innerHTML = '';
  const annotCounter = document.getElementById('annotations-counter');
  if (annotCounter) annotCounter.textContent = '';
}

// ── Métadonnées YAML (panneau de droite) ────────────────────────────────────────
// L'en-tête YAML (front-matter `--- ... ---` en début de document) est édité via
// un formulaire dans le volet de droite plutôt que sous forme de bloc dans le
// corps du document (voir règle CSS `#preview > .yaml-block:first-child`).
const _META_KNOWN_KEYS = ['title', 'abstract-title', 'lightbox', 'toc-depth'];

// Retire les guillemets entourant une valeur scalaire YAML (simple ou double).
function _yamlUnquote(value) {
  if (value.length >= 2 && value[0] === '"' && value[value.length - 1] === '"') {
    try { return JSON.parse(value); } catch (e) { return value.slice(1, -1); }
  }
  if (value.length >= 2 && value[0] === "'" && value[value.length - 1] === "'") {
    return value.slice(1, -1).replace(/''/g, "'");
  }
  return value;
}

// Extrait le bloc `--- ... ---` en tout début de contenu.
function _parseFrontMatter(content) {
  const m = (content || '').match(/^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/);
  if (!m) return { text: '', matchLength: 0 };
  return { text: m[1], matchLength: m[0].length };
}

// Découpe le texte YAML en champs connus (title, abstract-title, lightbox,
// toc-depth) et le reste des lignes (autres clés, conservées telles quelles).
function _splitYamlFields(text) {
  const lines = text.split('\n');
  const fields = {};
  const otherLines = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const m = line.match(/^([A-Za-z0-9_-]+):\s*(.*)$/);
    if (m && _META_KNOWN_KEYS.includes(m[1])) {
      const key = m[1];
      const value = m[2].trim();
      if (value === '|' || value === '|-' || value === '>') {
        // Bloc multiligne indenté
        const blockLines = [];
        let j = i + 1;
        let indent = null;
        while (j < lines.length) {
          const l = lines[j];
          if (l.trim() === '') { blockLines.push(''); j++; continue; }
          const im = l.match(/^(\s+)/);
          if (!im) break;
          if (indent === null) indent = im[1].length;
          if (im[1].length < indent) break;
          blockLines.push(l.slice(indent));
          j++;
        }
        while (blockLines.length && blockLines[blockLines.length - 1] === '') blockLines.pop();
        fields[key] = blockLines.join('\n');
        i = j;
        continue;
      }
      fields[key] = _yamlUnquote(value);
      i++;
      continue;
    }
    otherLines.push(line);
    i++;
  }
  while (otherLines.length && otherLines[0].trim() === '') otherLines.shift();
  while (otherLines.length && otherLines[otherLines.length - 1].trim() === '') otherLines.pop();
  return { fields, otherLines };
}

// Reconstruit le bloc `--- ... ---` à partir des champs du formulaire et des
// autres lignes YAML conservées. Retourne '' si tout est vide (pas d'en-tête).
function _buildFrontMatter(fields, otherLines) {
  const lines = [];
  if (fields.title) lines.push(`title: ${JSON.stringify(fields.title)}`);
  if (fields['abstract-title']) {
    lines.push('abstract-title: |');
    fields['abstract-title'].split('\n').forEach(l => lines.push('    ' + l));
  }
  if (fields.lightbox) lines.push('lightbox: true');
  if (fields['toc-depth']) lines.push(`toc-depth: ${parseInt(fields['toc-depth'], 10)}`);
  otherLines.forEach(l => lines.push(l));
  if (!lines.length) return '';
  return '---\n' + lines.join('\n') + '\n---\n\n';
}

// Recharge le formulaire « Métadonnées » à partir du contenu de l'onglet actif.
function _updateMetadataPanel() {
  const titleEl    = document.getElementById('meta-title');
  if (!titleEl) return;
  const abstractEl = document.getElementById('meta-abstract-title');
  const lightboxEl = document.getElementById('meta-lightbox');
  const tocDepthEl = document.getElementById('meta-toc-depth');
  const otherEl    = document.getElementById('meta-other-yaml');

  const tab = getActiveTab();
  const { text } = _parseFrontMatter(tab ? tab.content : '');
  const { fields, otherLines } = _splitYamlFields(text);

  const active = document.activeElement;
  if (active !== titleEl)    titleEl.value    = fields.title || '';
  if (active !== abstractEl) abstractEl.value = fields['abstract-title'] || '';
  if (active !== tocDepthEl) tocDepthEl.value = fields['toc-depth'] || '';
  if (active !== otherEl)    otherEl.value    = otherLines.join('\n');
  lightboxEl.checked = fields.lightbox === 'true';
}

// Appelé lors de la modification d'un champ du formulaire « Métadonnées » :
// reconstruit l'en-tête YAML et met à jour le contenu de l'onglet actif.
function onMetaFieldChange() {
  const tab = getActiveTab();
  if (!tab) return;

  const fields = {
    title: document.getElementById('meta-title').value.trim(),
    'abstract-title': document.getElementById('meta-abstract-title').value.replace(/\s+$/, ''),
    lightbox: document.getElementById('meta-lightbox').checked,
    'toc-depth': document.getElementById('meta-toc-depth').value.trim(),
  };
  const otherLines = document.getElementById('meta-other-yaml').value
    .split('\n').filter(l => l.trim() !== '');

  const { matchLength } = _parseFrontMatter(tab.content);
  const newBlock = _buildFrontMatter(fields, otherLines);
  const rest = tab.content.slice(matchLength).replace(/^\n+/, '');
  tab.content = newBlock + rest;
  tab.modified = (tab.content !== tab.savedContent);
  renderTabList();

  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    if (ta) {
      ta.value = tab.content;
      _updateSourceHighlight();
    }
  } else {
    updatePreview(tab.content);
  }
}

function toggleTocSection() {
  const section = document.getElementById('toc-section');
  const collapsed = section.classList.toggle('collapsed');
  const handle = document.getElementById('toc-meta-resize-handle');
  if (handle) handle.classList.toggle('disabled', collapsed);
}

function toggleMetadataSection() {
  const section = document.getElementById('metadata-section');
  const collapsed = section.classList.toggle('collapsed');
  const handle = document.getElementById('toc-meta-resize-handle');
  if (handle) handle.classList.toggle('disabled', collapsed);
  const metaAnnotHandle = document.getElementById('meta-annot-resize-handle');
  if (metaAnnotHandle) metaAnnotHandle.classList.toggle('disabled', collapsed);
}

function toggleMetadataOther() {
  document.getElementById('metadata-other-section').classList.toggle('collapsed');
}

// ── Preview ───────────────────────────────────────────────────────────────────
// ── TOC (Sommaire) ────────────────────────────────────────────────────────────
function updateTOC() {
  updateAnnotationsList();

  const body = document.getElementById('toc-body');
  if (!body) return;

  // JSON mode: build structure TOC from top-level keys
  if (document.body.classList.contains('json-mode')) {
    _updateJsonTOC(body);
    return;
  }

  const preview = document.getElementById('preview');
  const headings = [...preview.querySelectorAll('h1,h2,h3,h4,h5,h6')];

  body.innerHTML = '';

  if (!headings.length) {
    body.innerHTML = '<div class="toc-empty">Aucun titre</div>';
    return;
  }

  headings.forEach((h, i) => {
    if (!h.id) h.id = 'toc-h-' + i;
    const level    = parseInt(h.tagName[1]);
    const headingId = h.id;
    // Exclure le badge .heading-bookmark du texte affiché
    const badge    = h.querySelector('.heading-bookmark');
    const headText = badge
      ? h.textContent.replace(badge.textContent, '').trim()
      : h.textContent.trim();
    const div = document.createElement('div');
    div.className = `toc-item toc-h${level}`;
    div.dataset.headingId = headingId;
    div.innerHTML = `<span class="toc-bullet"></span><span class="toc-text" title="${headText}">${headText}</span>`;
    div.onclick = () => {
      if (state.sourceMode) {
        _scrollSourceToHeading(headText, level);
      } else {
        // Retrouver l'élément par ID (robuste après re-render)
        requestAnimationFrame(() => {
          const target = document.getElementById(headingId);
          if (!target) return;
          const pane = document.getElementById('preview-pane');
          if (pane) {
            const offset = target.getBoundingClientRect().top
                         - pane.getBoundingClientRect().top
                         + pane.scrollTop - 24;
            pane.scrollTo({ top: offset, behavior: 'smooth' });
          } else {
            target.scrollIntoView({ behavior: 'smooth', block: 'start' });
          }
        });
      }
    };
    body.appendChild(div);
  });
}

function _updateJsonTOC(body) {
  body.innerHTML = '';
  const ta = document.getElementById('json-editor');
  if (!ta) { body.innerHTML = '<div class="toc-empty">Aucun contenu</div>'; return; }
  let parsed;
  try { parsed = JSON.parse(ta.value); } catch (e) {
    body.innerHTML = '<div class="toc-empty">JSON invalide</div>';
    return;
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    body.innerHTML = '<div class="toc-empty">Aucune clé</div>';
    return;
  }
  const keys = Object.keys(parsed);
  if (!keys.length) { body.innerHTML = '<div class="toc-empty">Objet vide</div>'; return; }
  keys.forEach(key => {
    const val = parsed[key];
    const typeLabel = Array.isArray(val) ? `[ ${val.length} ]`
      : (val !== null && typeof val === 'object') ? `{ ${Object.keys(val).length} }`
      : String(val).slice(0, 40);
    const div = document.createElement('div');
    div.className = 'toc-item toc-h1';
    div.innerHTML = `<span class="toc-bullet"></span><span class="toc-text" title="${key}: ${typeLabel}">${key} <small style="opacity:.6">${typeLabel}</small></span>`;
    div.addEventListener('click', e => {
      if (body._jsonTocDragMoved) return;
      _jsonTocScrollToKey(key);
    });
    body.appendChild(div);
  });

  _initJsonTocDrag(body);
}

function _jsonTocScrollToKey(key) {
  const editor = document.getElementById('json-editor');
  const scrollEl = document.getElementById('json-scroll');
  if (!editor) return;
  const idx = editor.value.indexOf(`"${key}"`);
  if (idx < 0) return;
  editor.focus();
  editor.setSelectionRange(idx, idx + key.length + 2);
  const cs = getComputedStyle(editor);
  const lineH = parseFloat(cs.lineHeight) || (parseFloat(cs.fontSize) * 1.6) || 20;
  const line = editor.value.slice(0, idx).split('\n').length - 1;
  if (scrollEl) scrollEl.scrollTop = Math.max(0, line * lineH - scrollEl.clientHeight / 3);
}

function _initJsonTocDrag(body) {
  if (body._jsonDragCleanup) body._jsonDragCleanup();

  let dragging = false;
  body._jsonTocDragMoved = false;

  const onMousedown = e => {
    dragging = true;
    body._jsonTocDragMoved = false;
    body.style.cursor = 'grabbing';
    e.preventDefault();
  };

  const onMousemove = e => {
    if (!dragging) return;
    body._jsonTocDragMoved = true;
    const scrollEl = document.getElementById('json-scroll');
    if (!scrollEl) return;
    const rect = body.getBoundingClientRect();
    const ratio = Math.max(0, Math.min(1, (e.clientY - rect.top) / rect.height));
    scrollEl.scrollTop = ratio * scrollEl.scrollHeight;
  };

  const onMouseup = () => {
    if (!dragging) return;
    dragging = false;
    body.style.cursor = '';
    // Réinitialiser après que le click éventuel soit traité
    setTimeout(() => { body._jsonTocDragMoved = false; }, 50);
  };

  body.addEventListener('mousedown', onMousedown);
  document.addEventListener('mousemove', onMousemove);
  document.addEventListener('mouseup', onMouseup);

  body._jsonDragCleanup = () => {
    body.removeEventListener('mousedown', onMousedown);
    document.removeEventListener('mousemove', onMousemove);
    document.removeEventListener('mouseup', onMouseup);
  };
}

/** Positionne le curseur et scroll la source vers le titre Hx correspondant. */
function _scrollSourceToHeading(headingText, level) {
  const ta   = document.getElementById('source-editor');
  const pane = document.getElementById('source-scroll');
  const src  = ta.value;

  // Permet un suffixe optionnel de signet : ## Titre {#mon-id}
  const escaped = headingText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const suffix  = '(?:\\s+\\{[^}]*\\})?\\s*$';
  let match = new RegExp(`^#{${level}}\\s+${escaped}${suffix}`, 'mi').exec(src)
           || new RegExp(`^#{1,6}\\s+${escaped}${suffix}`, 'mi').exec(src);
  if (!match) return;

  const pos    = match.index;
  const endPos = pos + match[0].length;

  ta.setSelectionRange(pos, endPos);
  ta.focus();

  // Mesure la position Y exacte via miroir (tient compte du word-wrap)
  const cs           = getComputedStyle(ta);
  const paddingTop   = parseFloat(cs.paddingTop)   || 24;
  const paddingLeft  = parseFloat(cs.paddingLeft)  || 16;
  const paddingRight = parseFloat(cs.paddingRight) || 32;
  const lineHeight   = parseFloat(cs.lineHeight)   || (parseFloat(cs.fontSize) * 1.6);
  const contentW     = ta.clientWidth - paddingLeft - paddingRight;

  const lineNum = src.substring(0, pos).split('\n').length - 1;
  const targetScrollTop = Math.max(0, paddingTop + lineNum * lineHeight - pane.clientHeight / 4);

  // Applique après l'auto-scroll du focus pour ne pas être écrasé
  setTimeout(() => {
    if (pane) pane.scrollTop = targetScrollTop;
    _drawSourceMinimap();
  }, 0);
}

function toggleTOC() {
  const panel  = document.getElementById('toc-panel');
  const handle = document.getElementById('toc-resize-handle');
  if (panel.classList.contains('collapsed')) {
    panel.classList.remove('collapsed');
    if (panel._savedWidth) panel.style.width = panel._savedWidth;
    if (handle) handle.style.display = '';
  } else {
    panel._savedWidth = panel.style.width || (panel.offsetWidth + 'px');
    panel.style.width = '';
    panel.classList.add('collapsed');
    if (handle) handle.style.display = 'none';
  }
}

function _syncTabAreaOffset() {
  const sidebar = document.getElementById('sidebar');
  const tabArea = document.getElementById('tab-area');
  if (!sidebar || !tabArea) return;
  // Mesurer après que le navigateur a calculé le layout
  requestAnimationFrame(() => {
    const sidebarW = sidebar.getBoundingClientRect().right;
    const tabAreaL = tabArea.getBoundingClientRect().left;
    const padding  = Math.max(4, Math.round(sidebarW - tabAreaL));
    tabArea.style.paddingLeft = padding + 'px';
  });
}

function _initTabAreaOffsetSync() {
  const sidebar = document.getElementById('sidebar');
  if (!sidebar) return;
  // Se déclenche à chaque changement de taille de la sidebar (resize, collapse, paramètres)
  const ro = new ResizeObserver(_syncTabAreaOffset);
  ro.observe(sidebar);
  _syncTabAreaOffset();
}

function initSidebarResize() {
  const handle  = document.getElementById('sidebar-resize-handle');
  const sidebar = document.getElementById('sidebar');
  if (!handle || !sidebar) return;

  let dragging = false, startX = 0, startWidth = 0;

  handle.addEventListener('mousedown', e => {
    dragging   = true;
    startX     = e.clientX;
    startWidth = sidebar.offsetWidth;
    handle.classList.add('dragging');
    sidebar.style.transition       = 'none';
    document.body.style.cursor     = 'col-resize';
    document.body.style.userSelect = 'none';
    e.preventDefault();
  });

  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    const dx       = e.clientX - startX;
    const newWidth = Math.max(120, Math.min(600, startWidth + dx));
    sidebar.style.width = newWidth + 'px';
  });

  document.addEventListener('mouseup', () => {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove('dragging');
    sidebar.style.transition       = '';
    document.body.style.cursor     = '';
    document.body.style.userSelect = '';
  });
}

// Redimensionnement vertical entre les sections « Sommaire » et « Métadonnées »
// du volet de droite, via la poignée placée entre les deux.
function initTocMetaResize() {
  const handle      = document.getElementById('toc-meta-resize-handle');
  const tocSection  = document.getElementById('toc-section');
  const metaSection = document.getElementById('metadata-section');
  if (!handle || !tocSection || !metaSection) return;

  let dragging = false, startY = 0, startH = 0;

  handle.addEventListener('mousedown', e => {
    if (tocSection.classList.contains('collapsed') || metaSection.classList.contains('collapsed')) return;
    dragging = true;
    startY   = e.clientY;
    startH   = metaSection.offsetHeight;
    handle.classList.add('dragging');
    document.body.style.cursor     = 'row-resize';
    document.body.style.userSelect = 'none';
    e.preventDefault();
  });

  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    const panel = document.getElementById('toc-panel');
    const maxH  = panel.offsetHeight - handle.offsetHeight - 60; // garder un minimum pour le sommaire
    const dy    = e.clientY - startY; // glisser vers le bas = agrandir le sommaire
    const newH  = Math.max(60, Math.min(maxH, startH - dy));
    metaSection.style.height = newH + 'px';
  });

  document.addEventListener('mouseup', () => {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove('dragging');
    document.body.style.cursor     = '';
    document.body.style.userSelect = '';
  });
}

// Redimensionnement vertical entre les sections « Métadonnées » et « Commentaires »
// du volet de droite, via la poignée placée entre les deux.
function initMetaAnnotResize() {
  const handle       = document.getElementById('meta-annot-resize-handle');
  const metaSection  = document.getElementById('metadata-section');
  const annotSection = document.getElementById('annotations-section');
  if (!handle || !metaSection || !annotSection) return;

  let dragging = false, startY = 0, startH = 0;

  handle.addEventListener('mousedown', e => {
    if (metaSection.classList.contains('collapsed') || annotSection.classList.contains('collapsed')) return;
    dragging = true;
    startY   = e.clientY;
    startH   = annotSection.offsetHeight;
    handle.classList.add('dragging');
    document.body.style.cursor     = 'row-resize';
    document.body.style.userSelect = 'none';
    e.preventDefault();
  });

  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    const panel = document.getElementById('toc-panel');
    const maxH  = panel.offsetHeight - handle.offsetHeight - 120; // garder un minimum pour sommaire + métadonnées
    const dy    = e.clientY - startY; // glisser vers le bas = agrandir les métadonnées
    const newH  = Math.max(60, Math.min(maxH, startH - dy));
    annotSection.style.height = newH + 'px';
  });

  document.addEventListener('mouseup', () => {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove('dragging');
    document.body.style.cursor     = '';
    document.body.style.userSelect = '';
  });
}

function initTOCResize() {
  const handle = document.getElementById('toc-resize-handle');
  const panel  = document.getElementById('toc-panel');
  if (!handle || !panel) return;

  let dragging = false, startX = 0, startWidth = 0;

  handle.addEventListener('mousedown', e => {
    dragging   = true;
    startX     = e.clientX;
    startWidth = panel.offsetWidth;
    handle.classList.add('dragging');
    panel.style.transition  = 'none';
    document.body.style.cursor     = 'col-resize';
    document.body.style.userSelect = 'none';
    e.preventDefault();
  });

  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    const dx       = startX - e.clientX;          // drag gauche = panel plus grand
    const newWidth = Math.max(120, Math.min(600, startWidth + dx));
    panel.style.width = newWidth + 'px';
  });

  document.addEventListener('mouseup', () => {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove('dragging');
    panel.style.transition  = '';
    document.body.style.cursor     = '';
    document.body.style.userSelect = '';
  });
}

function initPumlPreviewResize() {
  const handle = document.getElementById('puml-preview-resize-handle');
  const panel  = document.getElementById('puml-preview-panel');
  if (!handle || !panel) return;

  let dragging = false, startX = 0, startWidth = 0;

  handle.addEventListener('mousedown', e => {
    dragging   = true;
    startX     = e.clientX;
    startWidth = panel.offsetWidth;
    handle.classList.add('dragging');
    panel.style.transition          = 'none';
    document.body.style.cursor     = 'col-resize';
    document.body.style.userSelect = 'none';
    e.preventDefault();
  });

  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    const dx       = startX - e.clientX;        // drag gauche = panel plus large
    const newWidth = Math.max(200, Math.min(window.innerWidth * 0.6, startWidth + dx));
    panel.style.width = newWidth + 'px';
    _drawPumlMinimap();
  });

  document.addEventListener('mouseup', () => {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove('dragging');
    panel.style.transition          = '';
    document.body.style.cursor     = '';
    document.body.style.userSelect = '';
  });
}

// Scroll-spy : met en surbrillance l'entrée TOC correspondant au titre visible
function initTOCScrollSpy() {
  const pane = document.getElementById('preview-pane');
  if (!pane) return;
  pane.addEventListener('scroll', () => {
    if (state.sourceMode) return;
    const body = document.getElementById('toc-body');
    if (!body) return;
    const preview  = document.getElementById('preview');
    const headings = [...preview.querySelectorAll('h1,h2,h3,h4,h5,h6')];
    if (!headings.length) return;
    const paneTop  = pane.getBoundingClientRect().top;
    let active = headings[0];
    for (const h of headings) {
      if (h.getBoundingClientRect().top - paneTop <= 8) active = h;
    }
    body.querySelectorAll('.toc-item').forEach(item => {
      item.classList.toggle('toc-active', item.dataset.headingId === active.id);
    });
  }, { passive: true });
}

function updatePreview(content) {
  const tab = getActiveTab();
  _currentDocBasePath = (tab && tab.path)
    ? tab.path.replace(/\\/g, '/').replace(/\/[^/]+$/, '')
    : '';
  const preview = document.getElementById('preview');
  preview.innerHTML = marked.parse(content || '');

  // Initialise la classe task-done sur les li déjà cochés au rendu
  preview.querySelectorAll('li.task-item input.task-checkbox:checked').forEach(cb => {
    cb.closest('li.task-item').classList.add('task-done');
  });

  // Injection DOM des labels de langue sur les blocs de code
  preview.querySelectorAll('pre > code').forEach(codeEl => {
    const pre = codeEl.parentElement;
    if (pre.parentElement && pre.parentElement.classList.contains('code-block-wrapper')) return;
    if (pre.closest('.yaml-block')) return;

    const langClass = [...codeEl.classList].find(c => c.startsWith('language-'));
    const lang = langClass ? langClass.replace('language-', '') : 'texinfo';

    const wrapper = document.createElement('div');
    wrapper.className = 'code-block-wrapper';

    const label = document.createElement('span');
    label.className = 'code-lang-label';
    label.textContent = lang;
    label.title = 'Cliquer pour changer le langage';
    label.addEventListener('click', e => { e.stopPropagation(); openLangEditor(label, codeEl); });

    pre.parentNode.insertBefore(wrapper, pre);
    wrapper.appendChild(label);
    wrapper.appendChild(pre);
  });

  // Centrer automatiquement les éléments avec un style inline width > 100%
  preview.querySelectorAll('[style]').forEach(el => {
    const w = el.style.width;
    if (w && w.endsWith('%')) {
      const pct = parseFloat(w);
      if (pct > 100) {
        el.style.position = 'relative';
        el.style.left = '50%';
        el.style.transform = `translateX(-50%)`;
        el.style.maxWidth = 'none';
      }
    }
  });

  // Indicateurs BR dans la marge gauche (hors blocs de code et listes)
  preview.querySelectorAll('br').forEach(br => {
    if (br.closest('pre') || br.closest('li')) return;
    if (br.dataset.fromLiteral) return; // <br> visuel inséré à la place d'un md-html-literal
    const marker = document.createElement('span');
    marker.className = 'br-marker';
    br.parentNode.insertBefore(marker, br.nextSibling);
  });
  // Balises <br> : colorer comme le fond (invisible sauf à la sélection).
  // Un vrai <br> est inséré avant pour l'effet visuel de saut de ligne.
  preview.querySelectorAll('.md-html-literal').forEach(el => {
    if (el.closest('pre')) return;
    if (/^<br\s*\/?>$/i.test(el.textContent.trim())) {
      el.classList.add('md-br-tag');
      if (!el.previousSibling || el.previousSibling.nodeName !== 'BR' || !el.previousSibling.dataset?.fromLiteral) {
        const br = document.createElement('br');
        br.dataset.fromLiteral = '1';
        el.parentNode.insertBefore(br, el);
      }
    }
  });

  preview.contentEditable = 'true';
  _annotateCitationTypes();
  updateTOC();
  _updateMetadataPanel();
}

let _annotateGeneration = 0;
async function _annotateCitationTypes() {
  const gen = ++_annotateGeneration;
  const entries = await _ensureBibEntries();
  // Si une annotation plus récente a démarré (nouveau updatePreview), on abandonne
  if (gen !== _annotateGeneration) return;
  if (!entries.length) return;
  const byKey = Object.fromEntries(entries.map(e => [e.key, e.type]));
  document.querySelectorAll('#preview .md-citation').forEach(sp => {
    const t = byKey[sp.dataset.key];
    if (t) sp.dataset.type = t;
    else delete sp.dataset.type;
  });
}

const LANGUAGES = [
  'texinfo','bash','sh','zsh','python','javascript','js','typescript','ts',
  'html','css','scss','json','yaml','toml','xml','markdown','sql','graphql',
  'c','cpp','csharp','java','rust','go','php','ruby','swift','kotlin',
  'r','julia','lua','perl','haskell','elixir','clojure','scala',
  'dockerfile','makefile','nginx','ini','diff','plaintext',
];

function openLangEditor(label, codeEl) {
  document.getElementById('lang-popup')?.remove();

  const current = label.textContent.trim();
  const popup   = document.createElement('div');
  popup.id = 'lang-popup';

  const input = document.createElement('input');
  input.type = 'text';
  input.placeholder = 'Filtrer…';
  popup.appendChild(input);

  const list = document.createElement('div');
  list.id = 'lang-popup-list';
  popup.appendChild(list);

  let focusedIdx = -1;

  const buildList = (filter) => {
    const q = filter.toLowerCase();
    const items = LANGUAGES.filter(l => l.includes(q));
    // Si le filtre ne correspond à aucun élément connu, propose-le quand même
    if (q && !LANGUAGES.includes(q)) items.unshift(q);
    list.innerHTML = '';
    focusedIdx = -1;
    items.forEach(lang => {
      const opt = document.createElement('div');
      opt.className = 'lang-option' + (lang === current ? ' current' : '');
      opt.textContent = lang;
      opt.addEventListener('mousedown', e => { e.preventDefault(); applyLang(lang); });
      list.appendChild(opt);
    });
  };

  const applyLang = (lang) => {
    label.textContent = lang;
    codeEl.className  = 'language-' + lang;
    popup.remove();
    syncPreviewToContent();
  };

  const moveFocus = (dir) => {
    const opts = list.querySelectorAll('.lang-option');
    if (!opts.length) return;
    opts[focusedIdx]?.classList.remove('focused');
    focusedIdx = Math.max(0, Math.min(opts.length - 1, focusedIdx + dir));
    opts[focusedIdx].classList.add('focused');
    opts[focusedIdx].scrollIntoView({ block: 'nearest' });
  };

  input.addEventListener('input', () => buildList(input.value));

  input.addEventListener('keydown', e => {
    if (e.key === 'ArrowDown')  { e.preventDefault(); moveFocus(1); }
    if (e.key === 'ArrowUp')    { e.preventDefault(); moveFocus(-1); }
    if (e.key === 'Enter') {
      e.preventDefault();
      const focused = list.querySelector('.lang-option.focused');
      if (focused) applyLang(focused.textContent);
      else if (input.value.trim()) applyLang(input.value.trim());
    }
    if (e.key === 'Escape') popup.remove();
  });

  input.addEventListener('blur', () => setTimeout(() => popup.remove(), 150));
  popup.addEventListener('mousedown', e => e.preventDefault());

  const rect = label.getBoundingClientRect();
  popup.style.top   = (rect.bottom + 6) + 'px';
  popup.style.right = (window.innerWidth - rect.right) + 'px';
  document.body.appendChild(popup);

  buildList('');
  // Pré-sélectionner le langage courant
  const currentOpt = [...list.querySelectorAll('.lang-option')].find(o => o.textContent === current);
  if (currentOpt) {
    focusedIdx = [...list.querySelectorAll('.lang-option')].indexOf(currentOpt);
    currentOpt.classList.add('focused');
    currentOpt.scrollIntoView({ block: 'nearest' });
  }

  input.focus();
}

// Reconstruit le texte d'un <code> en convertissant les <br> en \n
// (le navigateur insère des <br> quand on presse Entrée dans un contenteditable)
function _codeElText(node) {
  let text = '';
  node.childNodes.forEach(child => {
    if (child.nodeType === Node.TEXT_NODE) text += child.textContent;
    else if (child.nodeName === 'BR') text += '\n';
    else if (child.nodeName === 'DIV' || child.nodeName === 'P') text += '\n' + _codeElText(child);
    else text += _codeElText(child);
  });
  return text;
}

// ── Preview WYSIWYG editing ───────────────────────────────────────────────────
function setupPreviewEditing() {
  const preview = document.getElementById('preview');

  preview.addEventListener('focus', () => {
    if (!getActiveTab()) return;
    state.editingPreview = true;
  });

  preview.addEventListener('blur', () => {
    state.editingPreview = false;
    clearTimeout(previewInputTimer);
    syncPreviewToContent();
  });

  preview.addEventListener('input', () => {
    if (!state.editingPreview) return;
    clearTimeout(previewInputTimer);
    previewInputTimer = setTimeout(() => {
      syncPreviewToContent();
      updateTOC();
    }, 300);
  });

  // Tab / Shift+Tab dans un tableau : navigation entre cellules
  preview.addEventListener('keydown', e => {
    if (e.key !== 'Tab') return;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    let node = sel.getRangeAt(0).startContainer;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
    const cell = node.closest?.('td, th');
    if (!cell) return;
    e.preventDefault();

    // Collecter toutes les cellules du tableau dans l'ordre DOM
    const table = cell.closest('table');
    const cells = [...table.querySelectorAll('td, th')];
    const idx   = cells.indexOf(cell);
    const target = cells[e.shiftKey ? idx - 1 : idx + 1];

    // Dernière cellule + Tab → créer une nouvelle ligne
    if (!target && !e.shiftKey) {
      const numCols = cell.closest('tr').cells.length;
      const newRow  = document.createElement('tr');
      for (let i = 0; i < numCols; i++) newRow.appendChild(document.createElement('td'));
      const tbody = table.querySelector('tbody') || table;
      tbody.appendChild(newRow);

      const tableIdx = [...preview.querySelectorAll('table')].indexOf(table);
      _syncAndRerender();

      const updatedTable = preview.querySelectorAll('table')[tableIdx];
      if (updatedTable) {
        const newCells = [...updatedTable.querySelectorAll('td, th')];
        const firstNew = newCells[newCells.length - numCols];
        if (firstNew) {
          const r2 = document.createRange();
          r2.selectNodeContents(firstNew);
          const s2 = window.getSelection();
          s2.removeAllRanges();
          s2.addRange(r2);
          firstNew.scrollIntoView({ block: 'nearest' });
        }
      }
      return;
    }

    if (!target) return;

    // Placer le curseur dans la cellule cible (sélectionner tout le contenu)
    const r = document.createRange();
    r.selectNodeContents(target);
    sel.removeAllRanges();
    sel.addRange(r);
    // Faire défiler la cellule si nécessaire
    target.scrollIntoView({ block: 'nearest' });
  });

  // Tab partout : insérer des tabulations / espaces
  preview.addEventListener('keydown', e => {
    if (e.key !== 'Tab') return;
    if (e.defaultPrevented) return;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    let node = sel.getRangeAt(0).startContainer;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
    
    // Ignorer si tableau ou liste ET qu'on n'est pas dans un code block à l'intérieur
    if (node.closest?.('td, th')) return;
    if (node.closest?.('li') && !node.closest?.('pre, code')) return;

    e.preventDefault();
    const range = sel.getRangeAt(0);
    
    // Calcul de la colonne courante dans le bloc parent, en récupérant tout le texte qui précède
    // Cela résout le bug où le curseur traverse des noeuds TextNode successifs
    const block = node.closest('pre, p, div, li, h1, h2, h3, h4, h5, h6, blockquote') || document.getElementById('preview');
    const tempRange = document.createRange();
    tempRange.setStart(block, 0);
    tempRange.setEnd(range.startContainer, range.startOffset);
    const textSoFar = tempRange.toString();
    const lastNl = textSoFar.lastIndexOf('\n');
    const col = lastNl >= 0 ? textSoFar.length - lastNl - 1 : textSoFar.length;
    const spaces = _tabStopSpaces(col, e.shiftKey);
    
    const isCode = !!node.closest?.('pre, code');

    if (isCode) {
      // Bloc de code : insérer/supprimer un vrai caractère \t
      if (e.shiftKey) {
        if (range.startContainer.nodeType === Node.TEXT_NODE && range.startOffset > 0) {
          if (range.startContainer.textContent[range.startOffset - 1] === '\t') {
            const del = document.createRange();
            del.setStart(range.startContainer, range.startOffset - 1);
            del.setEnd(range.startContainer, range.startOffset);
            del.deleteContents();
          }
        }
      } else {
        range.deleteContents();
        const tn = document.createTextNode('\t');
        range.insertNode(tn);
        range.setStartAfter(tn);
        range.collapse(true);
        sel.removeAllRanges();
        sel.addRange(range);
      }
      syncPreviewToContent();
      return;
    }

    // Hors bloc de code : tabulation par arrêt de tabulation (espaces insécables)
    if (e.shiftKey) {
      const prevChars = textSoFar.slice(-spaces);
      if (spaces > 0 && range.startOffset >= spaces && (prevChars === ' '.repeat(spaces) || prevChars === '\u00A0'.repeat(spaces))) {
        const del = document.createRange();
        del.setStart(range.startContainer, range.startOffset - spaces);
        del.setEnd(range.startContainer, range.startOffset);
        del.deleteContents();
      }
    } else {
      range.deleteContents();
      const ins = document.createTextNode('\u00A0'.repeat(spaces));
      range.insertNode(ins);
      range.setStartAfter(ins);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    }
    syncPreviewToContent();
  });

  // Tab dans une liste du preview : indent / dé-indent
  preview.addEventListener('keydown', e => {
    if (e.key !== 'Tab') return;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    let node = sel.getRangeAt(0).startContainer;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
    const li = node.closest?.('li');
    if (!li) return;
    e.preventDefault();
    const parentList = li.parentElement; // ul ou ol
    if (!e.shiftKey) {
      // ── Indenter : imbriquer le <li> dans le <li> précédent ─────────────
      const prevLi = li.previousElementSibling;
      if (!prevLi) return; // premier item : rien à faire
      // Chercher une sous-liste existante dans prevLi
      const subTag  = parentList.tagName; // UL ou OL
      let   subList = prevLi.querySelector(':scope > ' + subTag);
      if (!subList) {
        subList = document.createElement(subTag);
        prevLi.appendChild(subList);
      }
      subList.appendChild(li);
    } else {
      // ── Dé-indenter : remonter le <li> d'un niveau ──────────────────────
      const grandParentLi = parentList.closest('li');
      if (!grandParentLi) {
        // ── Niveau racine → sortir de la liste, créer un paragraphe ──────
        // Collecter les items suivants pour les placer dans une nouvelle liste
        const afterItems = [];
        let sib = li.nextElementSibling;
        while (sib) { afterItems.push(sib); sib = sib.nextElementSibling; }

        // Créer le paragraphe à partir du contenu du <li> (hors sous-listes)
        const p = document.createElement('p');
        [...li.childNodes].forEach(child => {
          if (child.nodeName !== 'UL' && child.nodeName !== 'OL')
            p.appendChild(child.cloneNode(true));
        });
        if (!p.innerHTML.trim()) p.innerHTML = '<br>';

        // Insérer le paragraphe après la liste courante
        parentList.insertAdjacentElement('afterend', p);

        // S'il y a des items après, les regrouper dans une nouvelle liste
        if (afterItems.length > 0) {
          const newList = document.createElement(parentList.tagName);
          afterItems.forEach(item => newList.appendChild(item));
          p.insertAdjacentElement('afterend', newList);
        }

        // Supprimer le <li> d'origine (et la liste si elle est vide)
        li.remove();
        if (parentList.children.length === 0) parentList.remove();

        // Placer le curseur au début du paragraphe
        const r = document.createRange();
        r.setStart(p, 0);
        r.collapse(true);
        sel.removeAllRanges();
        sel.addRange(r);
        syncPreviewToContent();
        return;
      }
      const grandParentList = grandParentLi.parentElement;
      // Insérer après grandParentLi
      grandParentList.insertBefore(li, grandParentLi.nextSibling);
      // Nettoyer la sous-liste si elle est vide
      if (parentList.children.length === 0) parentList.remove();
    }
    // Repositionner le curseur dans le li déplacé
    const r = document.createRange();
    r.selectNodeContents(li);
    r.collapse(false);
    sel.removeAllRanges();
    sel.addRange(r);
    syncPreviewToContent();
  });

  // Shift+Enter dans un <li> : saut de ligne doux (même puce, plusieurs lignes)
  // Backspace après un <br> dans un <li> : supprime le saut doux proprement
  preview.addEventListener('keydown', e => {
    if (e.key !== 'Backspace') return;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    if (!range.collapsed) return;

    const container = range.startContainer;
    const offset    = range.startOffset;
    let prevNode = null;

    if (container.nodeType === Node.TEXT_NODE) {
      if (offset === 0) prevNode = container.previousSibling;
    } else {
      if (offset > 0) prevNode = container.childNodes[offset - 1];
    }

    if (prevNode && prevNode.nodeName === 'BR') {
      const el = container.nodeType === Node.TEXT_NODE ? container.parentElement : container;
      const li = el.closest?.('li');
      if (li) {
        e.preventDefault();
        const beforeBr = prevNode.previousSibling;
        prevNode.remove();
        const r = document.createRange();
        if (beforeBr && beforeBr.nodeType === Node.TEXT_NODE) {
          r.setStart(beforeBr, beforeBr.length);
        } else if (beforeBr) {
          r.setStartAfter(beforeBr);
        } else {
          r.setStart(li, 0);
        }
        r.collapse(true);
        sel.removeAllRanges();
        sel.addRange(r);
        syncPreviewToContent();
      }
    }
  });

  // Entrée dans un bloc de code : insérer \n au lieu du <br> du navigateur
  // Entrée sur une ligne vide dans une citation : sortir du blockquote
  preview.addEventListener('keydown', e => {
    if (e.key !== 'Enter' || e.ctrlKey || e.metaKey) return;
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    let node = sel.getRangeAt(0).startContainer;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;

    // Shift+Enter : pas de traitement spécial, comportement par défaut du navigateur
    if (e.shiftKey) return;

    // ── Enter sur paragraphe vide : insérer <br/> pour préserver la ligne ───────
    {
      const p = node.closest?.('p');
      if (p && !p.closest('blockquote') && !p.closest('pre')) {
        const kids = [...p.childNodes];
        const isEmpty = p.textContent.trim() === '' &&
          (kids.length === 0 || (kids.length === 1 && kids[0].nodeName === 'BR'));
        if (isEmpty) {
          e.preventDefault();
          p.textContent = '<br/>';
          const newP = document.createElement('p');
          newP.innerHTML = '<br>';
          p.insertAdjacentElement('afterend', newP);
          const r = document.createRange();
          r.setStart(newP, 0);
          r.collapse(true);
          sel.removeAllRanges();
          sel.addRange(r);
          syncPreviewToContent();
          return;
        }
      }
    }

    // ── Sortie de blockquote sur ligne vide ───────────────────────────────────
    const bq = node.closest?.('blockquote');
    if (bq) {
      // Remonter jusqu'au bloc enfant direct du blockquote
      let block = node;
      while (block.parentElement && block.parentElement !== bq) {
        block = block.parentElement;
      }
      // Sécurité : on doit bien avoir trouvé un enfant de bq
      if (block.parentElement === bq && block.textContent.trim() === '') {
        e.preventDefault();
        block.remove();
        if (!bq.textContent.trim()) bq.remove();
        // Nouveau paragraphe normal après le blockquote
        const p = document.createElement('p');
        p.innerHTML = '<br>';
        bq.insertAdjacentElement('afterend', p);
        const r = document.createRange();
        r.setStart(p, 0);
        r.collapse(true);
        sel.removeAllRanges();
        sel.addRange(r);
        syncPreviewToContent();
        return;
      }
    }

    // ── Entrée dans un bloc de code ───────────────────────────────────────────
    if (!node.closest?.('pre code')) return;
    e.preventDefault();
    const range = sel.getRangeAt(0);
    range.deleteContents();
    const nl = document.createTextNode('\n');
    range.insertNode(nl);
    range.setStartAfter(nl);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
  });

  document.addEventListener('selectionchange', () => {
    if (!state.editingPreview) return;
    document.querySelectorAll('#preview .fenced-div-wrapper.focused').forEach(el => el.classList.remove('focused'));
    const sel = window.getSelection();
    if (sel && sel.anchorNode && preview.contains(sel.anchorNode)) {
      const node = sel.anchorNode.nodeType === 3 ? sel.anchorNode.parentElement : sel.anchorNode;
      const wrapper = node.closest ? node.closest('.fenced-div-wrapper') : null;
      if (wrapper) wrapper.classList.add('focused');
    }
  });
}

function syncPreviewToContent() {
  if (_editingComment) return;
  const tab = getActiveTab();
  if (!tab) return;
  const preview = document.getElementById('preview');
  tab.content = turndown.turndown(preview.innerHTML);
  tab.modified = (tab.content !== tab.savedContent);
  renderTabList();
}

// ── Mode source ───────────────────────────────────────────────────────────────
// ── Synchronisation scroll preview ↔ source ───────────────────────────────────
// Construit une table de correspondance : numéro de ligne source ↔ Y dans le preview
// en utilisant les titres (Hx) comme ancres (même ordre dans les deux vues).
function _buildScrollMap(content, previewEl) {
  const lines    = content.split('\n');
  const headings = [...previewEl.querySelectorAll('h1,h2,h3,h4,h5,h6')];
  const map = [{ line: 0, y: 0 }];
  let hIdx = 0;
  for (let i = 0; i < lines.length && hIdx < headings.length; i++) {
    if (/^#{1,6}\s/.test(lines[i])) {
      map.push({ line: i, y: headings[hIdx].offsetTop });
      hIdx++;
    }
  }
  map.push({ line: lines.length, y: previewEl.scrollHeight });
  return map;
}

function _mapLineToY(line, map) {
  for (let i = 0; i < map.length - 1; i++) {
    if (line >= map[i].line && line <= map[i + 1].line) {
      const dl = map[i + 1].line - map[i].line;
      const dy = map[i + 1].y   - map[i].y;
      return map[i].y + (dl > 0 ? (line - map[i].line) / dl : 0) * dy;
    }
  }
  return 0;
}

function _mapYToLine(y, map) {
  for (let i = 0; i < map.length - 1; i++) {
    if (y >= map[i].y && y <= map[i + 1].y) {
      const dy = map[i + 1].y   - map[i].y;
      const dl = map[i + 1].line - map[i].line;
      return Math.round(map[i].line + (dy > 0 ? (y - map[i].y) / dy : 0) * dl);
    }
  }
  return 0;
}

function toggleSourceMode(enabled) {
  // Fermer la recherche lors d'un changement de mode
  _clearSearchHighlights();
  _searchMatches    = [];
  _searchSrcMatches = [];
  _searchCurrent    = -1;
  _searchLastQuery  = '';
  document.getElementById('search-count').textContent = '';
  document.getElementById('search-input').classList.remove('search-no-match');
  // Relancer la recherche dans le nouveau mode si la barre est ouverte
  const searchOpen = document.getElementById('search-bar').classList.contains('open');

  state.sourceMode = enabled;
  const previewPane = document.getElementById('preview-pane');
  const sourcePane  = document.getElementById('source-pane');
  const sourceScroll = document.getElementById('source-scroll');
  const tab = getActiveTab();

  if (enabled) {
    // ── Preview → Source ──────────────────────────────────────────────────────
    const content = tab ? tab.content : '';
    const preview = document.getElementById('preview');

    // 1. Calculer la ligne cible AVANT masquage (offsetTop invalide sur display:none)
    //    Priorité : position du nœud ancre DOM ; repli : scrollTop du panneau.
    const map = _buildScrollMap(content, preview);
    let targetLine = 0;
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0 && preview.contains(sel.anchorNode)) {
      let anchorEl = sel.anchorNode;
      if (anchorEl.nodeType === Node.TEXT_NODE) anchorEl = anchorEl.parentElement;
      // Remonter jusqu'à un enfant direct du preview (bloc de premier niveau)
      while (anchorEl && anchorEl !== preview && anchorEl.parentElement !== preview) {
        anchorEl = anchorEl.parentElement;
      }
      if (anchorEl && anchorEl !== preview) {
        const elTop   = anchorEl.getBoundingClientRect().top
                      - preview.getBoundingClientRect().top
                      + previewPane.scrollTop;
        targetLine = _mapYToLine(Math.max(0, elTop), map);
      } else {
        targetLine = _mapYToLine(previewPane.scrollTop, map);
      }
    } else {
      targetLine = _mapYToLine(previewPane.scrollTop, map);
    }

    // 2. Basculer l'affichage
    state.editingPreview = false;
    previewPane.style.display = 'none';
    sourcePane.style.display  = 'flex';

    // Mettre à jour la visibilité de la règle après rendu du layout
    setTimeout(() => _updateRulerVisibility(), 0);
    setTimeout(() => _updateRulerVisibility(), 150);

    // 3. Charger + rendre la source
    const editor = document.getElementById('source-editor');
    editor.value = content;
    editor.addEventListener('input', onSourceInput);
    _updateSourceHighlight();
    updateTOC();
    if (searchOpen) _doSearch();

    // 4. Positionner le curseur à la ligne cible
    const lines = content.split('\n');
    let charPos = 0;
    for (let i = 0; i < Math.min(targetLine, lines.length - 1); i++) {
      charPos += lines[i].length + 1;
    }
    editor.selectionStart = editor.selectionEnd = charPos;
    editor.focus();
    // Scroller dans un setTimeout pour passer après l'auto-scroll du focus
    setTimeout(() => {
      const lineH = editor.scrollHeight / Math.max(lines.length, 1);
      sourceScroll.scrollTop = Math.max(0, (targetLine - 3) * lineH);
      _drawSourceMinimap();
    }, 0);

  } else {
    // ── Source → Preview ──────────────────────────────────────────────────────
    const editor  = document.getElementById('source-editor');
    const content = tab ? tab.content : '';

    // 1. Ligne courante = position du curseur (pas du scroll)
    const currentLine = editor.value.substring(0, editor.selectionStart).split('\n').length - 1;

    // 2. Basculer l'affichage
    sourcePane.style.display  = 'none';
    previewPane.style.display = '';
    state.editingPreview = false;
    editor.removeEventListener('input', onSourceInput);

    // Masquer la règle (on quitte le mode source)
    _updateRulerVisibility();

    // 3. Rendre le preview
    if (tab) updatePreview(tab.content);
    if (searchOpen) setTimeout(_doSearch, 50);

    // 4. Scroller preview à la ligne du curseur après rendu
    setTimeout(() => {
      const prev = document.getElementById('preview');
      const map  = _buildScrollMap(content, prev);
      previewPane.scrollTop = _mapLineToY(currentLine, map);
      _updatePreviewWidthHandlePos();
    }, 30);
  }
  _updatePreviewWidthHandlePos();
}

function onSourceInput() {
  const tab = getActiveTab();
  if (!tab) return;
  tab.content = document.getElementById('source-editor').value;
  tab.modified = (tab.content !== tab.savedContent);
  renderTabList();
  _updateSourceHighlight();
  _updateMetadataPanel();
}

// ── Largeur automatique du wrapper (mode sans retour à la ligne) ──────────────
function _setEditorWrapperWidth(ta, ln, wrapperId, scrollId) {
  const wrapper = document.getElementById(wrapperId);
  const scroll  = document.getElementById(scrollId);
  if (!wrapper || !scroll) return;
  const cs    = getComputedStyle(ta);
  const tabSz = parseInt(cs.tabSize) || 2;
  const lines = ta.value.split('\n');
  let longestLine = '', longestLen = 0;
  for (const line of lines) {
    let len = 0;
    for (const ch of line) len = ch === '\t' ? Math.ceil((len + 1) / tabSz) * tabSz : len + 1;
    if (len > longestLen) { longestLen = len; longestLine = line; }
  }
  const ruler = document.createElement('span');
  ruler.style.cssText =
    'position:absolute;visibility:hidden;white-space:pre;display:inline-block;' +
    `font-family:${cs.fontFamily};font-size:${cs.fontSize};` +
    `top:-9999px;left:-9999px;padding:0;margin:0;tab-size:${cs.tabSize};`;
  ruler.textContent = longestLine || 'W';
  document.body.appendChild(ruler);
  const textW = ruler.offsetWidth;
  document.body.removeChild(ruler);
  const padL   = parseFloat(cs.paddingLeft)  || 16;
  const padR   = parseFloat(cs.paddingRight) || 32;
  const lnW    = ln ? ln.offsetWidth : 0;
  const needed = lnW + padL + textW + padR + 32;
  wrapper.style.minWidth = Math.max(needed, scroll.clientWidth) + 'px';
}

// ── Coloration syntaxique du mode source ──────────────────────────────────────
function _updateSourceHighlight() {
  const ta   = document.getElementById('source-editor');
  const pane = document.getElementById('source-scroll');
  // Sauvegarder le scroll du pane avant de réduire la textarea (évite le saut de vue)
  const savedScroll = pane ? pane.scrollTop : 0;
  // Auto-agrandit la textarea
  ta.style.height = '1px';
  ta.style.height = Math.max(ta.scrollHeight, pane ? pane.clientHeight : 0) + 'px';
  ta.scrollTop = 0;
  // Restaurer immédiatement le scroll du pane
  if (pane) pane.scrollTop = savedScroll;
  // Mettre à jour les numéros de ligne, la minimap et la coloration syntaxique
  _updateSourceLineNumbers();
  _drawSourceMinimap();
}

// ── Numéros de ligne et minimap du mode source ────────────────────────────────
function _updateSourceLineNumbers() {
  const ta = document.getElementById('source-editor');
  const ln = document.getElementById('source-line-numbers');
  if (!ln || !ta) return;
  const lines = ta.value.split('\n');
  const cs    = getComputedStyle(ta);
  const lineH = parseFloat(cs.lineHeight) || (parseFloat(cs.fontSize) * 1.6);
  ln.innerHTML = lines.map((_, i) =>
    `<span style="display:block;height:${lineH}px;line-height:${lineH}px">${i + 1}</span>`
  ).join('');
  if (ta.style.whiteSpace !== 'pre-wrap') {
    _setEditorWrapperWidth(ta, ln, 'source-wrapper', 'source-scroll');
  }
  // Synchronise la position gauche du pre overlay avec la largeur réelle des numéros de ligne
  requestAnimationFrame(() => {
    const lnW = ln.offsetWidth;
  });
}
function _sourceMinimapLineColor(line) {
  const t = line.trimStart();
  if (/^#{1,6}\s/.test(t)) {
    const level = t.match(/^(#{1,6})/)[1].length;
    const alpha = Math.max(0.55, 1 - (level - 1) * 0.12);
    return `rgba(137,180,250,${alpha})`;
  }
  if (/^`{3}/.test(t)) return 'rgba(166,227,161,0.5)';
  if (/^-{3,}$/.test(t)) return 'rgba(249,226,175,0.5)';
  if (/^>/.test(t))      return 'rgba(203,166,247,0.4)';
  return 'rgba(147,153,178,0.35)';
}

function _drawSourceMinimap() {
  const canvas   = document.getElementById('source-minimap');
  const scrollEl = document.getElementById('source-scroll');
  const ta       = document.getElementById('source-editor');
  if (!canvas || !scrollEl || !ta) return;

  const lines = ta.value.split('\n');
  if (!lines.length) return;

  const dpr = window.devicePixelRatio || 1;
  const w   = canvas.offsetWidth;
  const h   = canvas.offsetHeight;
  if (!w || !h) return;

  canvas.width  = w * dpr;
  canvas.height = h * dpr;

  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);

  ctx.fillStyle = '#1e1e2e';
  ctx.fillRect(0, 0, w, h);

  const lineH = h / lines.length;

  lines.forEach((line, i) => {
    const y  = i * lineH;
    const lh = Math.max(0.8, lineH);
    const isHeading = /^#{1,6}\s/.test(line.trimStart());
    if (isHeading) {
      ctx.fillStyle = _sourceMinimapLineColor(line);
      ctx.fillRect(0, y, w, Math.max(1.5, lh));
    } else {
      const len = Math.min(line.trimEnd().length, 120);
      if (len > 0) {
        const lw = Math.max(6, (len / 120) * (w - 8));
        ctx.fillStyle = _sourceMinimapLineColor(line);
        ctx.fillRect(4, y, lw, lh);
      }
    }
  });

  // Indicateur de fenêtre visible
  const scrollTop = scrollEl.scrollTop;
  const scrollH   = scrollEl.scrollHeight;
  const clientH   = scrollEl.clientHeight;
  if (scrollH > clientH) {
    const vpTop = (scrollTop / scrollH) * h;
    const vpH   = Math.max(16, (clientH / scrollH) * h);
    ctx.fillStyle   = 'rgba(205,214,244,0.08)';
    ctx.fillRect(0, vpTop, w, vpH);
    ctx.strokeStyle = 'rgba(205,214,244,0.28)';
    ctx.lineWidth   = 1;
    ctx.strokeRect(0.5, vpTop + 0.5, w - 1, vpH - 1);
  }
}

let _sourceMinimapDragging = false;

function _sourceMinimapScrollTo(clientY) {
  const canvas   = document.getElementById('source-minimap');
  const scrollEl = document.getElementById('source-scroll');
  if (!canvas || !scrollEl) return;
  const rect  = canvas.getBoundingClientRect();
  const ratio = Math.max(0, Math.min(1, (clientY - rect.top) / rect.height));
  scrollEl.scrollTop = ratio * scrollEl.scrollHeight - scrollEl.clientHeight / 2;
}

function _setupSourceMinimap() {
  const canvas   = document.getElementById('source-minimap');
  const scrollEl = document.getElementById('source-scroll');
  if (!canvas || !scrollEl) return;

  scrollEl.addEventListener('scroll', _drawSourceMinimap);

  canvas.addEventListener('mousedown', e => {
    e.preventDefault();
    _sourceMinimapDragging = true;
    _sourceMinimapScrollTo(e.clientY);
  });
  document.addEventListener('mousemove', e => {
    if (!_sourceMinimapDragging) return;
    _sourceMinimapScrollTo(e.clientY);
  });
  document.addEventListener('mouseup', () => { _sourceMinimapDragging = false; });

  if (window.ResizeObserver) {
    new ResizeObserver(() => {
      if (state.sourceMode) _drawSourceMinimap();
    }).observe(canvas.parentElement || canvas);
  }
}

// ── Éditeur PlantUML ─────────────────────────────────────────────────────────
const _PUML_KEYWORDS = new Set([
  'participant','actor','boundary','control','entity','database','collections',
  'queue','note','group','loop','alt','else','opt','break','critical','ref',
  'activate','deactivate','destroy','create','box','title','legend','class',
  'interface','abstract','enum','package','namespace','component','usecase',
  'node','cloud','artifact','folder','frame','agent','storage','rectangle',
  'card','file','stack','as','autonumber','hide','show','skinparam','return',
  'rnote','hnote','end','start','if','then','endif','fork','again','while',
  'repeat','stop','detach','left','right','over','on','is','not','and','or',
  'together','newpage','footbox','header','footer','scale','sprite','object',
  'map','json','yaml','salt','ditaa','latex','gantt','mindmap','wbs',
  'nwdiag','seqdiag','blockdiag','rackdiag','chronology'
]);

function highlightPuml(raw) {
  function _e(s) { return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

  return raw.split('\n').map(line => {
    const t = line.trimStart();

    // Blocs @startuml / @enduml / @startXxx
    if (/^@(start|end)\w*/.test(t))
      return `<span class="puml-block">${_e(line)}</span>`;

    // Commentaires : ' ou /'...'/ (ligne débutant par ')
    if (/^'/.test(t))
      return `<span class="puml-comment">${_e(line)}</span>`;

    // Tokenisation : strings → couleurs → flèches → mots-clés → reste
    let result = '';
    let i = 0;
    while (i < line.length) {
      const ch = line[i];

      // String "..."
      if (ch === '"') {
        let j = line.indexOf('"', i + 1);
        if (j < 0) j = line.length - 1;
        result += `<span class="puml-string">${_e(line.slice(i, j + 1))}</span>`;
        i = j + 1;
        continue;
      }

      // Couleur #xxx ou #RRGGBB
      if (ch === '#' && i + 1 < line.length && /[A-Za-z0-9]/.test(line[i + 1])) {
        let j = i + 1;
        while (j < line.length && /[\w]/.test(line[j])) j++;
        result += `<span class="puml-color">${_e(line.slice(i, j))}</span>`;
        i = j;
        continue;
      }

      // Flèches : ->, -->, ->>, <-, <--, <<-, <->, <=>, =>, ==>, ..>, <..
      const arrowM = line.slice(i).match(
        /^(-{1,3}>>?|<<?-{1,3}|={1,3}>>?|<<?={1,3}|\.{1,3}>>?|<<?\.{1,3}|<-?>|<--?>)/
      );
      if (arrowM) {
        result += `<span class="puml-arrow">${_e(arrowM[0])}</span>`;
        i += arrowM[0].length;
        continue;
      }

      // Mot (potentiellement un mot-clé)
      if (/[A-Za-z_]/.test(ch)) {
        let j = i;
        while (j < line.length && /[\w]/.test(line[j])) j++;
        const word = line.slice(i, j);
        result += _PUML_KEYWORDS.has(word.toLowerCase())
          ? `<span class="puml-kw">${_e(word)}</span>`
          : _e(word);
        i = j;
        continue;
      }

      result += _e(ch);
      i++;
    }
    return result;
  }).join('\n');
}

function _updatePumlHighlight() {
  const ta    = document.getElementById('puml-editor');
  const pre   = document.getElementById('puml-highlight');
  const area  = document.getElementById('puml-editor-area');
  const pane  = document.getElementById('puml-scroll');
  if (!ta || !pre) return;
  const lines = highlightPuml(ta.value).split('\n');
  pre.innerHTML = lines.map(l =>
    `<span class="puml-line">${l === '' ? '&#8203;' : l}</span>`
  ).join('');
  ta.style.height  = '1px';
  const newH = Math.max(ta.scrollHeight, pane ? pane.clientHeight : 0);
  ta.style.height  = newH + 'px';
  if (area) area.style.minHeight = newH + 'px';
  _updatePumlLineNumbers();
  _drawPumlMinimap();
}

function _updatePumlLineNumbers() {
  const ta = document.getElementById('puml-editor');
  const ln = document.getElementById('puml-line-numbers');
  if (!ln || !ta) return;
  const lines = ta.value.split('\n');
  const cs    = getComputedStyle(ta);
  const lineH = parseFloat(cs.lineHeight) || (parseFloat(cs.fontSize) * 1.6);
  ln.innerHTML = lines.map((_, i) =>
    `<span style="display:block;height:${lineH}px;line-height:${lineH}px">${i + 1}</span>`
  ).join('');
  _setEditorWrapperWidth(ta, ln, 'puml-wrapper', 'puml-scroll');
}
function _pumlMinimapLineColor(line) {
  const t = line.trimStart();
  if (/^@(start|end)\w*/.test(t)) return 'rgba(137,180,250,0.9)';   // blocs @start/@end → bleu vif
  if (/^'/.test(t))               return 'rgba(108,112,134,0.5)';   // commentaires → gris
  if (/^(note|title|legend|group|box|loop|alt|opt|break|if)\b/i.test(t))
                                  return 'rgba(203,166,247,0.7)';   // blocs structurants → violet
  return 'rgba(147,153,178,0.35)';
}

function _drawPumlMinimap() {
  const canvas   = document.getElementById('puml-minimap');
  const scrollEl = document.getElementById('puml-scroll');
  const ta       = document.getElementById('puml-editor');
  if (!canvas || !scrollEl || !ta) return;

  const lines = ta.value.split('\n');
  if (!lines.length) return;

  const dpr = window.devicePixelRatio || 1;
  const w   = canvas.offsetWidth;
  const h   = canvas.offsetHeight;
  if (!w || !h) return;

  canvas.width  = w * dpr;
  canvas.height = h * dpr;
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);
  ctx.fillStyle = '#1e1e2e';
  ctx.fillRect(0, 0, w, h);

  const lineH = h / lines.length;
  lines.forEach((line, i) => {
    const y  = i * lineH;
    const lh = Math.max(0.8, lineH);
    const isBlock = /^@(start|end)\w*/.test(line.trimStart());
    if (isBlock) {
      ctx.fillStyle = _pumlMinimapLineColor(line);
      ctx.fillRect(0, y, w, Math.max(1.5, lh));
    } else {
      const len = Math.min(line.trimEnd().length, 120);
      if (len > 0) {
        ctx.fillStyle = _pumlMinimapLineColor(line);
        ctx.fillRect(4, y, Math.max(6, (len / 120) * (w - 8)), lh);
      }
    }
  });

  // Indicateur viewport
  const scrollTop = scrollEl.scrollTop;
  const scrollH   = scrollEl.scrollHeight;
  const clientH   = scrollEl.clientHeight;
  if (scrollH > clientH) {
    const vpTop = (scrollTop / scrollH) * h;
    const vpH   = Math.max(16, (clientH / scrollH) * h);
    ctx.fillStyle   = 'rgba(205,214,244,0.08)';
    ctx.fillRect(0, vpTop, w, vpH);
    ctx.strokeStyle = 'rgba(205,214,244,0.28)';
    ctx.lineWidth   = 1;
    ctx.strokeRect(0.5, vpTop + 0.5, w - 1, vpH - 1);
  }
}

let _pumlMinimapDragging = false;

function _pumlMinimapScrollTo(clientY) {
  const canvas   = document.getElementById('puml-minimap');
  const scrollEl = document.getElementById('puml-scroll');
  if (!canvas || !scrollEl) return;
  const rect  = canvas.getBoundingClientRect();
  const ratio = Math.max(0, Math.min(1, (clientY - rect.top) / rect.height));
  scrollEl.scrollTop = ratio * scrollEl.scrollHeight - scrollEl.clientHeight / 2;
}

function _setupPumlMinimap() {
  const canvas   = document.getElementById('puml-minimap');
  const scrollEl = document.getElementById('puml-scroll');
  if (!canvas || !scrollEl) return;

  scrollEl.addEventListener('scroll', _drawPumlMinimap);

  canvas.addEventListener('mousedown', e => {
    e.preventDefault();
    _pumlMinimapDragging = true;
    _pumlMinimapScrollTo(e.clientY);
  });
  document.addEventListener('mousemove', e => {
    if (!_pumlMinimapDragging) return;
    _pumlMinimapScrollTo(e.clientY);
  });
  document.addEventListener('mouseup', () => { _pumlMinimapDragging = false; });

  if (window.ResizeObserver) {
    new ResizeObserver(() => {
      if (state.activeTabId && document.body.classList.contains('puml-mode')) _drawPumlMinimap();
    }).observe(canvas.parentElement || canvas);
  }
}

// ── Menu contextuel PUML ──────────────────────────────────────────────────────
function showPumlContextMenu(e) {
  e.preventDefault();
  e.stopPropagation();
  const menu = document.getElementById('puml-context-menu');
  if (!menu) return;
  // Positionner
  const vw = window.innerWidth, vh = window.innerHeight;
  let x = e.clientX, y = e.clientY;
  menu.style.display = 'block';
  const mw = menu.offsetWidth, mh = menu.offsetHeight;
  menu.style.display = '';
  if (x + mw > vw - 8) x = vw - mw - 8;
  if (y + mh > vh - 8) y = vh - mh - 8;
  menu.style.left = x + 'px';
  menu.style.top  = y + 'px';
  menu.classList.add('visible');
}

function hidePumlContextMenu() {
  document.getElementById('puml-context-menu')?.classList.remove('visible');
}

function pumlCtxCopy() {
  const ta = document.getElementById('puml-editor');
  const text = ta.value.substring(ta.selectionStart, ta.selectionEnd);
  if (text) navigator.clipboard.writeText(text);
  hidePumlContextMenu();
}

function pumlCtxCut() {
  const ta = document.getElementById('puml-editor');
  const s = ta.selectionStart, e = ta.selectionEnd;
  if (s !== e) {
    navigator.clipboard.writeText(ta.value.substring(s, e));
    ta.setRangeText('', s, e, 'start');
    onPumlInput();
  }
  hidePumlContextMenu();
}

function pumlCtxPaste() {
  hidePumlContextMenu();
  navigator.clipboard.readText().then(text => {
    if (!text) return;
    const ta = document.getElementById('puml-editor');
    ta.setRangeText(text, ta.selectionStart, ta.selectionEnd, 'end');
    onPumlInput();
  });
}

// ── Panneau aperçu PlantUML ───────────────────────────────────────────────────
let _pumlPreviewFmt = 'svg';

function pumlShowPreview(fmt) {
  hidePumlContextMenu();
  _pumlPreviewFmt = fmt || 'svg';
  const panel  = document.getElementById('puml-preview-panel');
  const handle = document.getElementById('puml-preview-resize-handle');
  panel.classList.add('open');
  if (handle) handle.classList.add('visible');
  _pumlZoom = 1.0; _applyPumlZoom();
  _pumlLoadPreviewImage(_pumlPreviewFmt);
  _updatePumlPreviewFmtBtns();
}

function closePumlPreview() {
  document.getElementById('puml-preview-panel').classList.remove('open');
  const handle = document.getElementById('puml-preview-resize-handle');
  if (handle) handle.classList.remove('visible');
}

function pumlPreviewSwitchFmt(fmt) {
  _pumlPreviewFmt = fmt;
  _pumlLoadPreviewImage(fmt);
  _updatePumlPreviewFmtBtns();
}

function _updatePumlPreviewFmtBtns() {
  document.getElementById('puml-fmt-svg')?.classList.toggle('active', _pumlPreviewFmt === 'svg');
  document.getElementById('puml-fmt-png')?.classList.toggle('active', _pumlPreviewFmt === 'png');
}

function _pumlLoadPreviewImage(fmt) {
  const tab = getActiveTab();
  if (!tab || !tab.path) return;

  // Chemin de l'image : même nom que le fichier .puml, extension changée
  const base = tab.path.replace(/\.[^.]+$/, '');
  const imgPath = base + '.' + fmt;
  const url = 'file:///' + imgPath.replace(/\\/g, '/');

  const img     = document.getElementById('puml-preview-img');
  const missing = document.getElementById('puml-preview-missing');

  // Label indique quel fichier est affiché
  document.getElementById('puml-preview-label').textContent =
    'Aperçu — ' + imgPath.split(/[\\/]/).pop();

  img.classList.remove('loaded');
  missing.classList.remove('visible');

  // Ajouter un timestamp pour forcer le rechargement après compilation
  img.src = url + '?t=' + Date.now();

  img.onload  = () => { img.classList.add('loaded'); missing.classList.remove('visible'); };
  img.onerror = () => { img.classList.remove('loaded'); missing.classList.add('visible'); };
}

// ---- Zoom aperçu PUML ----
let _pumlZoom = 1.0;

function _applyPumlZoom() {
  const img   = document.getElementById('puml-preview-img');
  const label = document.getElementById('puml-zoom-label');
  if (img)   img.style.width = (_pumlZoom * 100) + '%';
  if (label) label.textContent = Math.round(_pumlZoom * 100) + '%';
}

function pumlZoomIn()    { _pumlZoom = Math.min(4, Math.round((_pumlZoom + 0.25) * 100) / 100); _applyPumlZoom(); }
function pumlZoomOut()   { _pumlZoom = Math.max(0.25, Math.round((_pumlZoom - 0.25) * 100) / 100); _applyPumlZoom(); }
function pumlZoomReset() { _pumlZoom = 1.0; _applyPumlZoom(); }

function _initPumlPreviewZoom() {
  const content = document.getElementById('puml-preview-content');
  if (!content) return;
  content.addEventListener('wheel', e => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    if (e.deltaY < 0) pumlZoomIn(); else pumlZoomOut();
  }, { passive: false });
  // Bloquer le menu contextuel dans le volet aperçu
  document.getElementById('puml-preview-panel')
    ?.addEventListener('contextmenu', e => e.preventDefault());
}

// ---- Zoom visionneuse d'images ----
let _imageZoom = 1.0;

function _applyImageZoom() {
  const img   = document.getElementById('image-viewer');
  const label = document.getElementById('image-zoom-label');
  if (img)   img.style.width = (_imageZoom * 100) + '%';
  if (label) label.textContent = Math.round(_imageZoom * 100) + '%';
}

function imageZoomIn()    { _imageZoom = Math.min(4, Math.round((_imageZoom + 0.25) * 100) / 100); _applyImageZoom(); }
function imageZoomOut()   { _imageZoom = Math.max(0.25, Math.round((_imageZoom - 0.25) * 100) / 100); _applyImageZoom(); }
function imageZoomReset() { _imageZoom = 1.0; _applyImageZoom(); }

function _initImageViewerZoom() {
  const scroll = document.getElementById('image-scroll');
  if (!scroll) return;
  scroll.addEventListener('wheel', e => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    if (e.deltaY < 0) imageZoomIn(); else imageZoomOut();
  }, { passive: false });
  scroll.addEventListener('contextmenu', e => e.preventDefault());
}

/** Rafraîchit l'aperçu automatiquement après une compilation réussie. */
function _pumlRefreshPreviewAfterCompile() {
  const panel = document.getElementById('puml-preview-panel');
  if (panel?.classList.contains('open')) {
    _pumlLoadPreviewImage(_pumlPreviewFmt);
  }
}

function onPumlInput() {
  const tab = getActiveTab();
  if (!tab) return;
  tab.content = document.getElementById('puml-editor').value;
  tab.modified = (tab.content !== tab.savedContent);
  renderTabList();
  _updatePumlHighlight();
}

// ── Éditeur JSON ─────────────────────────────────────────────────────────────

function highlightJson(raw) {
  const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;' };
  const e = s => s.replace(/[&<>]/g, c => ESC[c]);

  // Tokeniser le JSON caractère par caractère
  let out = '';
  let i = 0;
  const src = raw;

  while (i < src.length) {
    // Chaîne
    if (src[i] === '"') {
      let j = i + 1;
      while (j < src.length) {
        if (src[j] === '\\') { j += 2; continue; }
        if (src[j] === '"') { j++; break; }
        j++;
      }
      const token = src.slice(i, j);
      // Regarder si c'est une clé (suivi de ":")
      let k = j;
      while (k < src.length && (src[k] === ' ' || src[k] === '\t')) k++;
      const isKey = src[k] === ':';
      out += `<span class="${isKey ? 'json-key' : 'json-str'}">${e(token)}</span>`;
      i = j;
      continue;
    }
    // Nombre
    const numM = src.slice(i).match(/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/);
    if (numM && (i === 0 || /[\s,\[{:]/.test(src[i - 1]))) {
      out += `<span class="json-num">${e(numM[0])}</span>`;
      i += numM[0].length;
      continue;
    }
    // true / false / null
    if (src.startsWith('true', i))  { out += `<span class="json-bool">true</span>`;  i += 4; continue; }
    if (src.startsWith('false', i)) { out += `<span class="json-bool">false</span>`; i += 5; continue; }
    if (src.startsWith('null', i))  { out += `<span class="json-null">null</span>`;  i += 4; continue; }
    // Ponctuation
    if ('{}[]:,'.includes(src[i])) { out += `<span class="json-punct">${e(src[i])}</span>`; i++; continue; }
    // Autre (whitespace, newline…)
    out += e(src[i]);
    i++;
  }
  return out;
}

function _updateJsonHighlight() {
  const ta     = document.getElementById('json-editor');
  const pre    = document.getElementById('json-highlight');
  const scroll = document.getElementById('json-scroll');
  if (!ta || !pre) return;
  const lines = highlightJson(ta.value).split('\n');
  pre.innerHTML = lines.map(l => `<span class="json-line">${l === '' ? '&#8203;' : l}</span>`).join('');
  ta.style.height = '1px';
  ta.style.height = Math.max(ta.scrollHeight, scroll ? scroll.clientHeight : 0) + 'px';
  _updateJsonLineNumbers();
  _drawJsonMinimap();
  const ln = document.getElementById('json-line-numbers');
  _setEditorWrapperWidth(ta, ln, 'json-wrapper', 'json-scroll');
}

function onJsonInput() {
  const tab = getActiveTab();
  if (!tab) return;
  tab.content = document.getElementById('json-editor').value;
  tab.modified = (tab.content !== tab.savedContent);
  renderTabList();
  _updateJsonHighlight();
}

function _updateJsonLineNumbers() {
  const ta = document.getElementById('json-editor');
  const ln = document.getElementById('json-line-numbers');
  if (!ln || !ta) return;
  const lines = ta.value.split('\n');
  const cs    = getComputedStyle(ta);
  const lineH = parseFloat(cs.lineHeight) || (parseFloat(cs.fontSize) * 1.6);
  ln.innerHTML = lines.map((_, i) =>
    `<span style="display:block;height:${lineH}px;line-height:${lineH}px">${i + 1}</span>`
  ).join('');
}

function _jsonMinimapLineColor(line) {
  const t = line.trim();
  if (!t) return null;
  if (/^"[^"]*"\s*:/.test(t))        return 'rgba(137,180,250,0.7)';   // clé
  if (/"/.test(t))                    return 'rgba(166,227,161,0.6)';   // chaîne valeur
  if (/\b(true|false)\b/.test(t))     return 'rgba(203,166,247,0.7)';   // bool
  if (/\bnull\b/.test(t))             return 'rgba(243,139,168,0.6)';   // null
  if (/^-?\d/.test(t))                return 'rgba(250,179,135,0.7)';   // nombre
  return 'rgba(108,112,134,0.35)';
}

function _drawJsonMinimap() {
  const canvas   = document.getElementById('json-minimap');
  const scrollEl = document.getElementById('json-scroll');
  const ta       = document.getElementById('json-editor');
  if (!canvas || !ta) return;
  const dpr   = window.devicePixelRatio || 1;
  const W     = canvas.offsetWidth  || 80;
  const H     = canvas.offsetHeight || 400;
  canvas.width  = W * dpr;
  canvas.height = H * dpr;
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);
  ctx.clearRect(0, 0, W, H);
  const lines    = ta.value.split('\n');
  const total    = lines.length || 1;
  const lineH    = Math.max(1, H / total);
  lines.forEach((line, idx) => {
    const col = _jsonMinimapLineColor(line);
    if (!col) return;
    ctx.fillStyle = col;
    ctx.fillRect(2, idx * lineH, W - 4, Math.max(1, lineH - 0.5));
  });
  if (scrollEl) {
    const ratio     = scrollEl.scrollTop / (scrollEl.scrollHeight || 1);
    const viewRatio = scrollEl.clientHeight / (scrollEl.scrollHeight || 1);
    ctx.fillStyle   = 'rgba(255,255,255,0.06)';
    ctx.fillRect(0, ratio * H, W, viewRatio * H);
  }
}

function _setupJsonMinimap() {
  const canvas = document.getElementById('json-minimap');
  const scroll = document.getElementById('json-scroll');
  if (!canvas || !scroll) return;
  scroll.addEventListener('scroll', _drawJsonMinimap);
  window.addEventListener('resize', _drawJsonMinimap);

  let _minimapDrag = false;
  const _minimapScroll = e => {
    const rect  = canvas.getBoundingClientRect();
    const ratio = Math.max(0, Math.min(1, (e.clientY - rect.top) / rect.height));
    scroll.scrollTop = ratio * scroll.scrollHeight;
  };
  canvas.addEventListener('mousedown', e => {
    _minimapDrag = true;
    _minimapScroll(e);
    e.preventDefault();
  });
  document.addEventListener('mousemove', e => {
    if (_minimapDrag) _minimapScroll(e);
  });
  document.addEventListener('mouseup', () => { _minimapDrag = false; });

  // Poignée de redimensionnement du minimap
  const resizeHandle = document.getElementById('json-minimap-resize-handle');
  if (resizeHandle && !resizeHandle._resizeInit) {
    resizeHandle._resizeInit = true;
    let dragging = false, startX = 0, startW = 0;
    resizeHandle.addEventListener('mousedown', e => {
      dragging = true;
      startX   = e.clientX;
      startW   = canvas.offsetWidth;
      resizeHandle.classList.add('dragging');
      document.body.style.cursor     = 'col-resize';
      document.body.style.userSelect = 'none';
      e.preventDefault();
    });
    document.addEventListener('mousemove', e => {
      if (!dragging) return;
      // déplacer la poignée vers la gauche = agrandir le minimap
      const dx = startX - e.clientX;
      const newW = Math.max(40, Math.min(300, startW + dx));
      canvas.style.width = newW + 'px';
      _drawJsonMinimap();
    });
    document.addEventListener('mouseup', () => {
      if (!dragging) return;
      dragging = false;
      resizeHandle.classList.remove('dragging');
      document.body.style.cursor     = '';
      document.body.style.userSelect = '';
    });
  }
}

// ── Éditeur XML ───────────────────────────────────────────────────────────────

function highlightXml(raw) {
  const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
  const e = s => s.replace(/[&<>"']/g, c => ESC[c]);

  let out = '';
  let i = 0;

  while (i < raw.length) {
    // Commentaire <!-- ... -->
    if (raw.startsWith('<!--', i)) {
      const end = raw.indexOf('-->', i + 4);
      const j = end >= 0 ? end + 3 : raw.length;
      out += `<span class="xml-comment">${e(raw.slice(i, j))}</span>`;
      i = j; continue;
    }
    // CDATA <![CDATA[...]]>
    if (raw.startsWith('<![CDATA[', i)) {
      const end = raw.indexOf(']]>', i + 9);
      const j = end >= 0 ? end + 3 : raw.length;
      out += `<span class="xml-cdata">${e(raw.slice(i, j))}</span>`;
      i = j; continue;
    }
    // Déclaration <?...?>
    if (raw.startsWith('<?', i)) {
      const end = raw.indexOf('?>', i + 2);
      const j = end >= 0 ? end + 2 : raw.length;
      out += `<span class="xml-decl">${e(raw.slice(i, j))}</span>`;
      i = j; continue;
    }
    // Toute balise < ... >
    if (raw[i] === '<') {
      // Trouver la fin de la balise en respectant les chaînes
      let j = i + 1;
      let inStr = false, strChar = '';
      while (j < raw.length) {
        if (inStr) {
          if (raw[j] === strChar) inStr = false;
        } else {
          if (raw[j] === '"' || raw[j] === "'") { inStr = true; strChar = raw[j]; }
          else if (raw[j] === '>') { j++; break; }
        }
        j++;
      }
      // Tokeniser la balise caractère par caractère (sans regex sur du HTML déjà produit)
      const tag = raw.slice(i, j);
      const isClose = tag[1] === '/';
      let k = 0;
      let tagOut = '';

      // < ou </
      if (isClose) {
        tagOut += `<span class="xml-tag-close">${e('</')}</span>`;
        k = 2;
      } else {
        tagOut += `<span class="xml-tag-open">${e('<')}</span>`;
        k = 1;
      }

      // Nom de la balise
      const nameStart = k;
      while (k < tag.length && !/[\s/>]/.test(tag[k])) k++;
      if (k > nameStart) {
        const cls = isClose ? 'xml-tag-close' : 'xml-tag-open';
        tagOut += `<span class="${cls}">${e(tag.slice(nameStart, k))}</span>`;
      }

      // Attributs et fermeture
      while (k < tag.length) {
        // Espaces blancs
        if (/\s/.test(tag[k])) {
          let ws = '';
          while (k < tag.length && /\s/.test(tag[k])) ws += tag[k++];
          tagOut += e(ws);
          continue;
        }
        // Fermeture />
        if (tag[k] === '/' && tag[k + 1] === '>') {
          tagOut += `<span class="xml-tag-close">${e('/>')}</span>`;
          k += 2; break;
        }
        // Fermeture >
        if (tag[k] === '>') {
          tagOut += `<span class="xml-tag-close">${e('>')}</span>`;
          k++; break;
        }
        // Nom d'attribut
        const attrStart = k;
        while (k < tag.length && !/[\s=/>]/.test(tag[k])) k++;
        tagOut += `<span class="xml-attr-name">${e(tag.slice(attrStart, k))}</span>`;
        // Espaces + signe =
        let between = '';
        while (k < tag.length && /\s/.test(tag[k])) between += tag[k++];
        if (k < tag.length && tag[k] === '=') {
          between += '='; k++;
          while (k < tag.length && /\s/.test(tag[k])) between += tag[k++];
          tagOut += e(between);
          // Valeur entre guillemets
          if (k < tag.length && (tag[k] === '"' || tag[k] === "'")) {
            const q = tag[k];
            const vStart = k; k++;
            while (k < tag.length && tag[k] !== q) k++;
            if (k < tag.length) k++; // guillemet fermant
            tagOut += `<span class="xml-attr-value">${e(tag.slice(vStart, k))}</span>`;
          }
        } else {
          tagOut += e(between);
        }
      }
      out += tagOut;
      i = j; continue;
    }
    // Texte ordinaire
    out += e(raw[i]);
    i++;
  }
  return out;
}

function _updateXmlHighlight() {
  const ta     = document.getElementById('xml-editor');
  const pre    = document.getElementById('xml-highlight');
  const scroll = document.getElementById('xml-scroll');
  if (!ta || !pre) return;
  const lines = highlightXml(ta.value).split('\n');
  pre.innerHTML = lines.map(l => `<span class="xml-line">${l === '' ? '&#8203;' : l}</span>`).join('');
  ta.style.height = '1px';
  ta.style.height = Math.max(ta.scrollHeight, scroll ? scroll.clientHeight : 0) + 'px';
  _updateXmlLineNumbers();
  _drawXmlMinimap();
  const ln = document.getElementById('xml-line-numbers');
  _setEditorWrapperWidth(ta, ln, 'xml-wrapper', 'xml-scroll');
}

function onXmlInput() {
  const tab = getActiveTab();
  if (!tab) return;
  tab.content = document.getElementById('xml-editor').value;
  tab.modified = (tab.content !== tab.savedContent);
  renderTabList();
  _updateXmlHighlight();
}

function _updateXmlLineNumbers() {
  const ta = document.getElementById('xml-editor');
  const ln = document.getElementById('xml-line-numbers');
  if (!ln || !ta) return;
  const lines = ta.value.split('\n');
  const cs    = getComputedStyle(ta);
  const lineH = parseFloat(cs.lineHeight) || (parseFloat(cs.fontSize) * 1.6);
  ln.innerHTML = lines.map((_, i) =>
    `<span style="display:block;height:${lineH}px;line-height:${lineH}px">${i + 1}</span>`
  ).join('');
}

function _xmlMinimapLineColor(line) {
  const t = line.trim();
  if (!t) return null;
  if (/^<!--/.test(t))                    return 'rgba(108,112,134,0.5)';   // commentaire
  if (/^<!\[CDATA\[/.test(t))             return 'rgba(250,179,135,0.7)';   // CDATA
  if (/^<\?/.test(t))                     return 'rgba(243,139,168,0.6)';   // déclaration
  if (/^<\//.test(t))                     return 'rgba(137,180,250,0.5)';   // balise fermante
  if (/^<[a-zA-Z]/.test(t))              return 'rgba(137,180,250,0.8)';   // balise ouvrante
  return 'rgba(205,214,244,0.3)';
}

function _drawXmlMinimap() {
  const canvas   = document.getElementById('xml-minimap');
  const scrollEl = document.getElementById('xml-scroll');
  const ta       = document.getElementById('xml-editor');
  if (!canvas || !ta) return;
  const dpr   = window.devicePixelRatio || 1;
  const W     = canvas.offsetWidth  || 80;
  const H     = canvas.offsetHeight || 400;
  canvas.width  = W * dpr;
  canvas.height = H * dpr;
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);
  ctx.clearRect(0, 0, W, H);
  const lines    = ta.value.split('\n');
  const total    = lines.length || 1;
  const lineH    = Math.max(1, H / total);
  lines.forEach((line, idx) => {
    const col = _xmlMinimapLineColor(line);
    if (!col) return;
    ctx.fillStyle = col;
    ctx.fillRect(2, idx * lineH, W - 4, Math.max(1, lineH - 0.5));
  });
  if (scrollEl) {
    const ratio     = scrollEl.scrollTop / (scrollEl.scrollHeight || 1);
    const viewRatio = scrollEl.clientHeight / (scrollEl.scrollHeight || 1);
    ctx.fillStyle   = 'rgba(255,255,255,0.06)';
    ctx.fillRect(0, ratio * H, W, viewRatio * H);
  }
}

function _setupXmlMinimap() {
  const canvas = document.getElementById('xml-minimap');
  const scroll = document.getElementById('xml-scroll');
  if (!canvas || !scroll) return;
  scroll.addEventListener('scroll', _drawXmlMinimap);
  window.addEventListener('resize', _drawXmlMinimap);

  let _minimapDrag = false;
  const _minimapScroll = e => {
    const rect  = canvas.getBoundingClientRect();
    const ratio = Math.max(0, Math.min(1, (e.clientY - rect.top) / rect.height));
    scroll.scrollTop = ratio * scroll.scrollHeight;
  };
  canvas.addEventListener('mousedown', e => {
    _minimapDrag = true;
    _minimapScroll(e);
    e.preventDefault();
  });
  document.addEventListener('mousemove', e => {
    if (_minimapDrag) _minimapScroll(e);
  });
  document.addEventListener('mouseup', () => { _minimapDrag = false; });

  const resizeHandle = document.getElementById('xml-minimap-resize-handle');
  if (resizeHandle && !resizeHandle._resizeInit) {
    resizeHandle._resizeInit = true;
    let dragging = false, startX = 0, startW = 0;
    resizeHandle.addEventListener('mousedown', e => {
      dragging = true;
      startX   = e.clientX;
      startW   = canvas.offsetWidth;
      resizeHandle.classList.add('dragging');
      document.body.style.cursor     = 'col-resize';
      document.body.style.userSelect = 'none';
      e.preventDefault();
    });
    document.addEventListener('mousemove', e => {
      if (!dragging) return;
      const dx = startX - e.clientX;
      const newW = Math.max(40, Math.min(300, startW + dx));
      canvas.style.width = newW + 'px';
      _drawXmlMinimap();
    });
    document.addEventListener('mouseup', () => {
      if (!dragging) return;
      dragging = false;
      resizeHandle.classList.remove('dragging');
      document.body.style.cursor     = '';
      document.body.style.userSelect = '';
    });
  }
}

// ── Polling éditeurs spécialisés ─────────────────────────────────────────────
// pywebview/Chromium ne déclenche pas toujours 'input'/'keyup' sur les
// textareas transparentes. On scrute la valeur toutes les 150 ms.
// ── Polling éditeurs spécialisés ─────────────────────────────────────────────
const _editorPollingMap = [
  { id: 'puml-editor', mode: 'puml-mode', fn: () => onPumlInput() },
  { id: 'yaml-editor', mode: 'yaml-mode', fn: () => onYamlInput() },
  { id: 'css-editor',  mode: 'css-mode',  fn: () => onCssInput()  },
  { id: 'lua-editor',  mode: 'lua-mode',  fn: () => onLuaInput()  },
  { id: 'json-editor', mode: 'json-mode', fn: () => onJsonInput() },
  { id: 'xml-editor',  mode: 'xml-mode',  fn: () => onXmlInput()  },
];
window._editorLastVal = {};

// Réinitialise la valeur de référence pour un éditeur (à appeler après switchToTab)
function _editorPollingSync(id, value) {
  window._editorLastVal[id] = value;
}

setInterval(() => {
  for (const { id, mode, fn } of _editorPollingMap) {
    if (!document.body.classList.contains(mode)) continue;
    const ta = document.getElementById(id);
    if (!ta) continue;
    const cur = ta.value;
    if (cur !== window._editorLastVal[id]) {
      window._editorLastVal[id] = cur;
      fn();
    }
  }
}, 150);

// ── Éditeur CSS ──────────────────────────────────────────────────────────────

function highlightCss(raw) {
  const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;' };
  const e   = s => s.replace(/[&<>]/g, c => ESC[c]);
  const sp  = (cls, txt) => `<span class="css-${cls}">${e(txt)}</span>`;

  let out = '', i = 0, depth = 0;
  const src = raw, len = src.length;
  const isWord  = c => /[\w-]/.test(c);
  const isDigit = c => /[0-9.]/.test(c);

  while (i < len) {
    const c = src[i];

    // Block comment /* ... */
    if (c === '/' && src[i+1] === '*') {
      const end = src.indexOf('*/', i + 2);
      const s   = end < 0 ? src.slice(i) : src.slice(i, end + 2);
      out += sp('comment', s); i += s.length; continue;
    }
    // Line comment //
    if (c === '/' && src[i+1] === '/') {
      let j = i; while (j < len && src[j] !== '\n') j++;
      out += sp('comment', src.slice(i, j)); i = j; continue;
    }
    // String
    if (c === '"' || c === "'") {
      let j = i + 1;
      while (j < len && src[j] !== c && src[j] !== '\n') { if (src[j] === '\\') j++; j++; }
      if (src[j] === c) j++;
      out += sp('string', src.slice(i, j)); i = j; continue;
    }
    // @-rule
    if (c === '@') {
      let j = i + 1; while (j < len && isWord(src[j])) j++;
      out += sp('atrule', src.slice(i, j)); i = j; continue;
    }
    // !important
    if (c === '!' && src.slice(i, i+10) === '!important') {
      out += sp('important', '!important'); i += 10; continue;
    }
    // Open brace
    if (c === '{') { depth++; out += sp('brace', '{'); i++; continue; }
    // Close brace
    if (c === '}') { depth = Math.max(0, depth - 1); out += sp('brace', '}'); i++; continue; }

    // Inside rule block
    if (depth > 0) {
      // Hex color value
      if (c === '#') {
        let j = i + 1; while (j < len && /[0-9a-fA-F]/.test(src[j])) j++;
        if (j - i >= 4) { out += sp('color', src.slice(i, j)); i = j; continue; }
      }
      // Number + optional unit
      if (isDigit(c)) {
        let j = i; while (j < len && isDigit(src[j])) j++;
        const UNITS = ['px','em','rem','%','vh','vw','vmin','vmax','pt','cm','mm','deg','rad','turn','s','ms','fr','ch'];
        for (const u of UNITS) {
          if (src.slice(j, j + u.length) === u && !/\w/.test(src[j + u.length] || '')) {
            j += u.length; break;
          }
        }
        out += sp('number', src.slice(i, j)); i = j; continue;
      }
      // Property name (word before ':')
      if (isWord(c)) {
        let j = i; while (j < len && isWord(src[j])) j++;
        let k = j; while (k < len && (src[k] === ' ' || src[k] === '\t')) k++;
        if (src[k] === ':' && src[k+1] !== ':') {
          out += sp('property', src.slice(i, j)); i = j; continue;
        }
        out += sp('value', src.slice(i, j)); i = j; continue;
      }
      // Colon / semicolon
      if (c === ':' || c === ';') { out += sp('punct', c); i++; continue; }
    } else {
      // Outside rule: selector tokens
      if (c === '#') {
        let j = i + 1; while (j < len && isWord(src[j])) j++;
        out += sp('sel-id', src.slice(i, j)); i = j; continue;
      }
      if (c === '.') {
        let j = i + 1; while (j < len && isWord(src[j])) j++;
        out += sp('sel-cls', src.slice(i, j)); i = j; continue;
      }
      if (c === ':') {
        // pseudo-class / pseudo-element
        let j = i; if (src[i+1] === ':') j++;
        j++; while (j < len && isWord(src[j])) j++;
        out += sp('sel-pct', src.slice(i, j)); i = j; continue;
      }
      if (isWord(c)) {
        let j = i; while (j < len && isWord(src[j])) j++;
        out += sp('selector', src.slice(i, j)); i = j; continue;
      }
    }

    out += e(c); i++;
  }
  return out;
}

function _updateCssHighlight() {
  const ta     = document.getElementById('css-editor');
  const pre    = document.getElementById('css-highlight');
  const scroll = document.getElementById('css-scroll');
  if (!ta || !pre) return;
  const lines = highlightCss(ta.value).split('\n');
  pre.innerHTML = lines.map(l => `<span class="css-line">${l === '' ? '&#8203;' : l}</span>`).join('');
  ta.style.height = '1px';
  ta.style.height = Math.max(ta.scrollHeight, scroll ? scroll.clientHeight : 0) + 'px';
  _updateCssLineNumbers();
  _drawCssMinimap();
}

function onCssInput() {
  const tab = getActiveTab();
  if (!tab) return;
  tab.content = document.getElementById('css-editor').value;
  tab.modified = (tab.content !== tab.savedContent);
  renderTabList();
  _updateCssHighlight();
}

function _updateCssLineNumbers() {
  const ta = document.getElementById('css-editor');
  const ln = document.getElementById('css-line-numbers');
  if (!ln || !ta) return;
  const lines = ta.value.split('\n');
  const cs    = getComputedStyle(ta);
  const lineH = parseFloat(cs.lineHeight) || (parseFloat(cs.fontSize) * 1.6);
  ln.innerHTML = lines.map((_, i) =>
    `<span style="display:block;height:${lineH}px;line-height:${lineH}px">${i + 1}</span>`
  ).join('');
  _setEditorWrapperWidth(ta, ln, 'css-wrapper', 'css-scroll');
}
function _cssMinimapLineColor(line) {
  const t = line.trimStart();
  if (/^\/[/*]/.test(t))          return 'rgba(108,112,134,0.5)';    // commentaire
  if (/^\s*\*/.test(t))           return 'rgba(108,112,134,0.4)';    // suite commentaire bloc
  if (/^\s*@/.test(t))            return 'rgba(203,166,247,0.75)';   // @-rule
  if (/\{/.test(t))               return 'rgba(137,180,250,0.75)';   // sélecteur + {
  if (/^\s*[\w-]+\s*:/.test(t))  return 'rgba(137,220,235,0.6)';    // propriété
  if (/^\s*\}/.test(t))           return 'rgba(137,180,250,0.4)';    // }
  return 'rgba(147,153,178,0.35)';
}

function _drawCssMinimap() {
  const canvas   = document.getElementById('css-minimap');
  const scrollEl = document.getElementById('css-scroll');
  const ta       = document.getElementById('css-editor');
  if (!canvas || !scrollEl || !ta) return;
  const lines = ta.value.split('\n');
  if (!lines.length) return;
  const dpr = window.devicePixelRatio || 1;
  const w = canvas.offsetWidth, h = canvas.offsetHeight;
  if (!w || !h) return;
  canvas.width  = w * dpr; canvas.height = h * dpr;
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);
  ctx.fillStyle = '#1e1e2e'; ctx.fillRect(0, 0, w, h);
  const lineH = h / lines.length;
  lines.forEach((line, i) => {
    const y   = i * lineH, lh = Math.max(0.8, lineH);
    const len = Math.min(line.trimEnd().length, 120);
    if (len > 0) {
      ctx.fillStyle = _cssMinimapLineColor(line);
      ctx.fillRect(4, y, Math.max(6, (len / 120) * (w - 8)), lh);
    }
  });
  const scrollTop = scrollEl.scrollTop, scrollH = scrollEl.scrollHeight, clientH = scrollEl.clientHeight;
  if (scrollH > clientH) {
    const vpTop = (scrollTop / scrollH) * h;
    const vpH   = Math.max(16, (clientH / scrollH) * h);
    ctx.fillStyle = 'rgba(205,214,244,0.08)'; ctx.fillRect(0, vpTop, w, vpH);
    ctx.strokeStyle = 'rgba(205,214,244,0.28)'; ctx.lineWidth = 1;
    ctx.strokeRect(0.5, vpTop + 0.5, w - 1, vpH - 1);
  }
}

let _cssMinimapDragging = false;

function _cssMmScrollTo(clientY) {
  const canvas = document.getElementById('css-minimap');
  const scroll = document.getElementById('css-scroll');
  if (!canvas || !scroll) return;
  const rect  = canvas.getBoundingClientRect();
  const ratio = Math.max(0, Math.min(1, (clientY - rect.top) / rect.height));
  scroll.scrollTop = ratio * scroll.scrollHeight - scroll.clientHeight / 2;
}

function _setupCssMinimap() {
  const canvas = document.getElementById('css-minimap');
  const scroll = document.getElementById('css-scroll');
  if (!canvas || !scroll) return;
  scroll.addEventListener('scroll', _drawCssMinimap);
  canvas.addEventListener('mousedown', e => {
    e.preventDefault(); _cssMinimapDragging = true; _cssMmScrollTo(e.clientY);
  });
  document.addEventListener('mousemove', e => { if (_cssMinimapDragging) _cssMmScrollTo(e.clientY); });
  document.addEventListener('mouseup',   () => { _cssMinimapDragging = false; });
  if (window.ResizeObserver) {
    new ResizeObserver(() => {
      if (state.activeTabId && document.body.classList.contains('css-mode')) _drawCssMinimap();
    }).observe(canvas.parentElement || canvas);
  }
}

// ── Coloration syntaxique du mode Lua ────────────────────────────────────────
const _LUA_KEYWORDS = new Set(['and','break','do','else','elseif','end','false','for','function','goto','if','in','local','nil','not','or','repeat','return','then','true','until','while']);
const _LUA_BUILTINS = new Set(['print','pairs','ipairs','type','tostring','tonumber','table','string','math','os','io','pcall','xpcall','require','setmetatable','getmetatable','error','assert','select','next','rawget','rawset','rawequal','unpack','coroutine','self']);

function highlightLua(raw) {
  const ESC = { '&': '&amp;', '<': '&lt;', '>': '&gt;' };
  const e   = s => s.replace(/[&<>]/g, c => ESC[c]);
  const sp  = (cls, txt) => `<span class="lua-${cls}">${e(txt)}</span>`;

  let out = '', i = 0;
  const src = raw, len = src.length;
  const isWord  = c => /[A-Za-z0-9_]/.test(c);
  const isDigit = c => /[0-9]/.test(c);

  while (i < len) {
    const c = src[i];

    // Commentaire bloc --[[ ... ]] (avec niveaux ==)
    if (c === '-' && src[i+1] === '-' && /^\[=*\[/.test(src.slice(i+2))) {
      const m     = src.slice(i+2).match(/^\[(=*)\[/);
      const close = `]${m[1]}]`;
      const end   = src.indexOf(close, i + 2 + m[0].length);
      const s     = end < 0 ? src.slice(i) : src.slice(i, end + close.length);
      out += sp('comment', s); i += s.length; continue;
    }
    // Commentaire ligne --
    if (c === '-' && src[i+1] === '-') {
      let j = i; while (j < len && src[j] !== '\n') j++;
      out += sp('comment', src.slice(i, j)); i = j; continue;
    }
    // Chaîne longue [[ ... ]] / [=[ ... ]=]
    if (c === '[' && /^\[=*\[/.test(src.slice(i))) {
      const m     = src.slice(i).match(/^\[(=*)\[/);
      const close = `]${m[1]}]`;
      const end   = src.indexOf(close, i + m[0].length);
      const s     = end < 0 ? src.slice(i) : src.slice(i, end + close.length);
      out += sp('string', s); i += s.length; continue;
    }
    // Chaîne ' ou "
    if (c === '"' || c === "'") {
      let j = i + 1;
      while (j < len && src[j] !== c && src[j] !== '\n') { if (src[j] === '\\') j++; j++; }
      if (src[j] === c) j++;
      out += sp('string', src.slice(i, j)); i = j; continue;
    }
    // Nombre
    if (isDigit(c) || (c === '.' && isDigit(src[i+1]))) {
      let j = i; while (j < len && /[0-9a-fA-Fx.]/.test(src[j])) j++;
      out += sp('number', src.slice(i, j)); i = j; continue;
    }
    // Identifiant / mot-clé / builtin / appel de fonction
    if (/[A-Za-z_]/.test(c)) {
      let j = i; while (j < len && isWord(src[j])) j++;
      const word = src.slice(i, j);
      let k = j; while (k < len && (src[k] === ' ' || src[k] === '\t')) k++;
      if (_LUA_KEYWORDS.has(word)) out += sp('keyword', word);
      else if (src[k] === '(') out += sp('func', word);
      else if (_LUA_BUILTINS.has(word)) out += sp('builtin', word);
      else out += e(word);
      i = j; continue;
    }
    // Opérateurs
    if ('+-*/%^#=<>~'.includes(c)) {
      let j = i + 1;
      if ('=<>'.includes(c) && src[j] === '=') j++;
      out += sp('operator', src.slice(i, j)); i = j; continue;
    }
    // Ponctuation
    if ('(){}[];,.:'.includes(c)) { out += sp('punct', c); i++; continue; }

    out += e(c); i++;
  }
  return out;
}

function _updateLuaHighlight() {
  const ta     = document.getElementById('lua-editor');
  const pre    = document.getElementById('lua-highlight');
  const scroll = document.getElementById('lua-scroll');
  if (!ta || !pre) return;
  const lines = highlightLua(ta.value).split('\n');
  pre.innerHTML = lines.map(l => `<span class="lua-line">${l === '' ? '&#8203;' : l}</span>`).join('');
  ta.style.height = '1px';
  ta.style.height = Math.max(ta.scrollHeight, scroll ? scroll.clientHeight : 0) + 'px';
  _updateLuaLineNumbers();
  _drawLuaMinimap();
}

function onLuaInput() {
  const tab = getActiveTab();
  if (!tab) return;
  tab.content = document.getElementById('lua-editor').value;
  tab.modified = (tab.content !== tab.savedContent);
  renderTabList();
  _updateLuaHighlight();
}

function _updateLuaLineNumbers() {
  const ta = document.getElementById('lua-editor');
  const ln = document.getElementById('lua-line-numbers');
  if (!ln || !ta) return;
  const lines = ta.value.split('\n');
  const cs    = getComputedStyle(ta);
  const lineH = parseFloat(cs.lineHeight) || (parseFloat(cs.fontSize) * 1.6);
  ln.innerHTML = lines.map((_, i) =>
    `<span style="display:block;height:${lineH}px;line-height:${lineH}px">${i + 1}</span>`
  ).join('');
  _setEditorWrapperWidth(ta, ln, 'lua-wrapper', 'lua-scroll');
}

function _luaMinimapLineColor(line) {
  const t = line.trimStart();
  if (/^--/.test(t))                    return 'rgba(108,112,134,0.5)';   // commentaire
  if (/^(local\s+)?function\b/.test(t)) return 'rgba(137,180,250,0.75)';  // function
  if (/^(local|return|if|elseif|else|for|while|repeat|end)\b/.test(t)) return 'rgba(203,166,247,0.75)'; // mot-clé
  return 'rgba(147,153,178,0.35)';
}

function _drawLuaMinimap() {
  const canvas   = document.getElementById('lua-minimap');
  const scrollEl = document.getElementById('lua-scroll');
  const ta       = document.getElementById('lua-editor');
  if (!canvas || !scrollEl || !ta) return;
  const lines = ta.value.split('\n');
  if (!lines.length) return;
  const dpr = window.devicePixelRatio || 1;
  const w = canvas.offsetWidth, h = canvas.offsetHeight;
  if (!w || !h) return;
  canvas.width  = w * dpr; canvas.height = h * dpr;
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);
  ctx.fillStyle = '#1e1e2e'; ctx.fillRect(0, 0, w, h);
  const lineH = h / lines.length;
  lines.forEach((line, i) => {
    const y   = i * lineH, lh = Math.max(0.8, lineH);
    const len = Math.min(line.trimEnd().length, 120);
    if (len > 0) {
      ctx.fillStyle = _luaMinimapLineColor(line);
      ctx.fillRect(4, y, Math.max(6, (len / 120) * (w - 8)), lh);
    }
  });
  const scrollTop = scrollEl.scrollTop, scrollH = scrollEl.scrollHeight, clientH = scrollEl.clientHeight;
  if (scrollH > clientH) {
    const vpTop = (scrollTop / scrollH) * h;
    const vpH   = Math.max(16, (clientH / scrollH) * h);
    ctx.fillStyle = 'rgba(205,214,244,0.08)'; ctx.fillRect(0, vpTop, w, vpH);
    ctx.strokeStyle = 'rgba(205,214,244,0.28)'; ctx.lineWidth = 1;
    ctx.strokeRect(0.5, vpTop + 0.5, w - 1, vpH - 1);
  }
}

let _luaMinimapDragging = false;

function _luaMmScrollTo(clientY) {
  const canvas = document.getElementById('lua-minimap');
  const scroll = document.getElementById('lua-scroll');
  if (!canvas || !scroll) return;
  const rect  = canvas.getBoundingClientRect();
  const ratio = Math.max(0, Math.min(1, (clientY - rect.top) / rect.height));
  scroll.scrollTop = ratio * scroll.scrollHeight - scroll.clientHeight / 2;
}

function _setupLuaMinimap() {
  const canvas = document.getElementById('lua-minimap');
  const scroll = document.getElementById('lua-scroll');
  if (!canvas || !scroll) return;
  scroll.addEventListener('scroll', _drawLuaMinimap);
  canvas.addEventListener('mousedown', e => {
    e.preventDefault(); _luaMinimapDragging = true; _luaMmScrollTo(e.clientY);
  });
  document.addEventListener('mousemove', e => { if (_luaMinimapDragging) _luaMmScrollTo(e.clientY); });
  document.addEventListener('mouseup',   () => { _luaMinimapDragging = false; });
  if (window.ResizeObserver) {
    new ResizeObserver(() => {
      if (state.activeTabId && document.body.classList.contains('lua-mode')) _drawLuaMinimap();
    }).observe(canvas.parentElement || canvas);
  }
}

// ── Coloration syntaxique du mode YAML ───────────────────────────────────────
function _updateYamlHighlight() {
  const ta     = document.getElementById('yaml-editor');
  const pre    = document.getElementById('yaml-highlight');
  const scroll = document.getElementById('yaml-scroll');
  if (!ta || !pre) return;
  const lines = highlightYaml(ta.value).split('\n');
  pre.innerHTML = lines.map(l => `<span class="yml-line">${l === '' ? '&#8203;' : l}</span>`).join('');
  ta.style.height = '1px';
  ta.style.height = Math.max(ta.scrollHeight, scroll ? scroll.clientHeight : 0) + 'px';
  _updateYamlLineNumbers();
  _drawYamlMinimap();
}

function onYamlInput() {
  const tab = getActiveTab();
  if (!tab) return;
  tab.content = document.getElementById('yaml-editor').value;
  tab.modified = (tab.content !== tab.savedContent);
  renderTabList();
  _updateYamlHighlight();
}

// ── Numéros de ligne YAML ─────────────────────────────────────────────────────
function _updateYamlLineNumbers() {
  const ta = document.getElementById('yaml-editor');
  const ln = document.getElementById('yaml-line-numbers');
  if (!ln || !ta) return;
  const lines = ta.value.split('\n');
  const cs    = getComputedStyle(ta);
  const lineH = parseFloat(cs.lineHeight) || (parseFloat(cs.fontSize) * 1.6);
  ln.innerHTML = lines.map((_, i) =>
    `<span style="display:block;height:${lineH}px;line-height:${lineH}px">${i + 1}</span>`
  ).join('');
  _setEditorWrapperWidth(ta, ln, 'yaml-wrapper', 'yaml-scroll');
}
// ── Minimap YAML ──────────────────────────────────────────────────────────────
function _yamlMinimapLineColor(line) {
  const t = line.trimStart();
  if (/^#/.test(t))             return 'rgba(108,112,134,0.5)';    // commentaires → gris
  if (/^[\w-]+\s*:/.test(t))   return 'rgba(137,180,250,0.75)';   // clés → bleu
  if (/^-\s/.test(t))          return 'rgba(166,227,161,0.55)';   // items liste → vert
  return 'rgba(147,153,178,0.35)';
}

function _drawYamlMinimap() {
  const canvas   = document.getElementById('yaml-minimap');
  const scrollEl = document.getElementById('yaml-scroll');
  const ta       = document.getElementById('yaml-editor');
  if (!canvas || !scrollEl || !ta) return;

  const lines = ta.value.split('\n');
  if (!lines.length) return;

  const dpr = window.devicePixelRatio || 1;
  const w   = canvas.offsetWidth;
  const h   = canvas.offsetHeight;
  if (!w || !h) return;

  canvas.width  = w * dpr;
  canvas.height = h * dpr;
  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);
  ctx.fillStyle = '#1e1e2e';
  ctx.fillRect(0, 0, w, h);

  const lineH = h / lines.length;
  lines.forEach((line, i) => {
    const y   = i * lineH;
    const lh  = Math.max(0.8, lineH);
    const len = Math.min(line.trimEnd().length, 120);
    if (len > 0) {
      ctx.fillStyle = _yamlMinimapLineColor(line);
      ctx.fillRect(4, y, Math.max(6, (len / 120) * (w - 8)), lh);
    }
  });

  // Indicateur viewport
  const scrollTop = scrollEl.scrollTop;
  const scrollH   = scrollEl.scrollHeight;
  const clientH   = scrollEl.clientHeight;
  if (scrollH > clientH) {
    const vpTop = (scrollTop / scrollH) * h;
    const vpH   = Math.max(16, (clientH / scrollH) * h);
    ctx.fillStyle   = 'rgba(205,214,244,0.08)';
    ctx.fillRect(0, vpTop, w, vpH);
    ctx.strokeStyle = 'rgba(205,214,244,0.28)';
    ctx.lineWidth   = 1;
    ctx.strokeRect(0.5, vpTop + 0.5, w - 1, vpH - 1);
  }
}

let _yamlMinimapDragging = false;

function _yamlMinimapScrollTo(clientY) {
  const canvas   = document.getElementById('yaml-minimap');
  const scrollEl = document.getElementById('yaml-scroll');
  if (!canvas || !scrollEl) return;
  const rect  = canvas.getBoundingClientRect();
  const ratio = Math.max(0, Math.min(1, (clientY - rect.top) / rect.height));
  scrollEl.scrollTop = ratio * scrollEl.scrollHeight - scrollEl.clientHeight / 2;
}

function _setupYamlMinimap() {
  const canvas   = document.getElementById('yaml-minimap');
  const scrollEl = document.getElementById('yaml-scroll');
  if (!canvas || !scrollEl) return;

  scrollEl.addEventListener('scroll', _drawYamlMinimap);

  canvas.addEventListener('mousedown', e => {
    e.preventDefault();
    _yamlMinimapDragging = true;
    _yamlMinimapScrollTo(e.clientY);
  });
  document.addEventListener('mousemove', e => {
    if (!_yamlMinimapDragging) return;
    _yamlMinimapScrollTo(e.clientY);
  });
  document.addEventListener('mouseup', () => { _yamlMinimapDragging = false; });

  if (window.ResizeObserver) {
    new ResizeObserver(() => {
      if (state.activeTabId && document.body.classList.contains('yaml-mode')) _drawYamlMinimap();
    }).observe(canvas.parentElement || canvas);
  }
}

// ── Visionneuse de fichiers Log / TeX ─────────────────────────────────────────
let _logRawContent  = '';
let _logSearchQuery = '';
let _logSearchMatches = [];   // [{line, start, end}]
let _logSearchIdx   = -1;
let _logViewerMode  = 'log';  // 'log' | 'tex'

function _escLog(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function _texLevelClass(line) {
  const t = line.trimStart();
  if (/^%/.test(t))                                           return 'log-trace';  // commentaire
  if (/^\\(chapter|section|subsection|subsubsection)\b/.test(t)) return 'log-warn';   // titres
  if (/^\\(begin|end)\s*\{/.test(t))                          return 'log-info';   // environnements
  if (/^\$\$|^\\[\[\]]/.test(t))                              return 'log-debug';  // maths display
  if (/^\\/.test(t))                                          return 'log-tex-cmd'; // autres commandes
  return '';
}

function _logLevelClass(line) {
  if (_logViewerMode === 'tex') return _texLevelClass(line);
  const up = line.toUpperCase();
  if (/\b(ERROR|SEVERE|CRITICAL|FATAL)\b/.test(up)) return 'log-error';
  if (/\b(WARN|WARNING)\b/.test(up))                return 'log-warn';
  if (/\bINFO\b/.test(up))                          return 'log-info';
  if (/\bDEBUG\b/.test(up))                         return 'log-debug';
  if (/\bTRACE\b/.test(up))                         return 'log-trace';
  if (/^\s+at\s|^\s+\.\.\.\s\d/.test(line))         return 'log-trace'; // stack trace
  return '';
}

function _buildLogHtml(lines, query, matches, currentIdx) {
  // Map : lineIndex → matches in that line
  const byLine = {};
  matches.forEach((m, mi) => {
    (byLine[m.line] = byLine[m.line] || []).push({ s: m.start, e: m.end, mi });
  });

  return lines.map((line, i) => {
    const lvl = _logLevelClass(line);
    const ms  = byLine[i];
    let content;
    if (!ms) {
      content = _escLog(line) || '&#8203;';
    } else {
      let res = '', pos = 0;
      for (const m of ms) {
        res += _escLog(line.slice(pos, m.s));
        const cls = m.mi === currentIdx ? 'log-match log-match-current' : 'log-match';
        res += `<mark class="${cls}">${_escLog(line.slice(m.s, m.e))}</mark>`;
        pos = m.e;
      }
      res += _escLog(line.slice(pos));
      content = res || '&#8203;';
    }
    return `<div class="log-row${lvl ? ' ' + lvl : ''}">`
      + `<span class="log-num">${i + 1}</span>`
      + `<span class="log-line">${content}</span>`
      + `</div>`;
  }).join('');
}

function _findLogMatches(lines, query) {
  if (!query) return [];
  const lq = query.toLowerCase();
  const matches = [];
  lines.forEach((line, li) => {
    const ll = line.toLowerCase();
    let pos = 0;
    while ((pos = ll.indexOf(lq, pos)) !== -1) {
      matches.push({ line: li, start: pos, end: pos + lq.length });
      pos += lq.length;
    }
  });
  return matches;
}

function _refreshLogContent() {
  const lines = _logRawContent.split('\n');
  document.getElementById('log-content').innerHTML =
    _buildLogHtml(lines, _logSearchQuery, _logSearchMatches,
      _logSearchMatches.length > 0 ? _logSearchIdx : -1);

  const countEl = document.getElementById('log-search-count');
  if (_logSearchQuery) {
    if (_logSearchMatches.length === 0) {
      countEl.textContent = 'Aucun résultat';
      countEl.classList.add('log-no-result');
    } else {
      countEl.textContent = `${_logSearchIdx + 1}\u202f/\u202f${_logSearchMatches.length}`;
      countEl.classList.remove('log-no-result');
      _scrollToCurrentLogMatch();
    }
  } else {
    countEl.textContent = '';
    countEl.classList.remove('log-no-result');
  }
  _drawLogMinimap();
}

function _scrollToCurrentLogMatch() {
  requestAnimationFrame(() => {
    const pane    = document.getElementById('log-pane');
    const current = document.querySelector('#log-content .log-match-current');
    if (!current || !pane) return;
    const paneRect  = pane.getBoundingClientRect();
    const matchRect = current.getBoundingClientRect();
    const offset = matchRect.top - paneRect.top - pane.clientHeight / 2 + matchRect.height / 2;
    pane.scrollTop += offset;
  });
}

function renderLogPane(content, mode) {
  _logViewerMode  = mode || 'log';
  _logRawContent  = content;
  _logSearchQuery = '';
  _logSearchMatches = [];
  _logSearchIdx   = -1;

  const searchInput = document.getElementById('log-search-input');
  const countEl     = document.getElementById('log-search-count');
  const lineCountEl = document.getElementById('log-line-count');
  if (searchInput) searchInput.value = '';
  if (countEl)     { countEl.textContent = ''; countEl.classList.remove('log-no-result'); }

  const lines = content.split('\n');
  document.getElementById('log-content').innerHTML = _buildLogHtml(lines, '', [], -1);
  if (lineCountEl) lineCountEl.textContent = `${lines.length} ligne${lines.length > 1 ? 's' : ''}`;

  // Scroll en bas (comportement naturel pour un log)
  requestAnimationFrame(() => {
    const el = document.getElementById('log-content');
    if (el) el.scrollTop = el.scrollHeight;
    _drawLogMinimap();
  });
}

// ── Minimap log ───────────────────────────────────────────────────────────────
let _logMinimapDragging = false;

function _texLevelColor(line) {
  const t = line.trimStart();
  if (/^%/.test(t))                                               return 'rgba(69,71,90,0.35)';
  if (/^\\(chapter|section|subsection|subsubsection)\b/.test(t)) return 'rgba(250,179,135,0.70)';
  if (/^\\(begin|end)\s*\{/.test(t))                             return 'rgba(137,220,235,0.55)';
  if (/^\$\$|^\\[\[\]]/.test(t))                                 return 'rgba(203,166,247,0.50)';
  if (/^\\/.test(t))                                             return 'rgba(166,227,161,0.35)';
  return 'rgba(205,214,244,0.18)';
}

function _logLevelColor(line) {
  if (_logViewerMode === 'tex') return _texLevelColor(line);
  const up = line.toUpperCase();
  if (/\b(ERROR|SEVERE|CRITICAL|FATAL)\b/.test(up)) return 'rgba(243,139,168,0.80)';
  if (/\b(WARN|WARNING)\b/.test(up))                return 'rgba(250,179,135,0.65)';
  if (/\bINFO\b/.test(up))                          return 'rgba(137,220,235,0.50)';
  if (/\bDEBUG\b/.test(up))                         return 'rgba(108,112,134,0.40)';
  if (/\bTRACE\b/.test(up) || /^\s+at\s/.test(line)) return 'rgba(69,71,90,0.35)';
  return 'rgba(205,214,244,0.18)';
}

function _drawLogMinimap() {
  const canvas   = document.getElementById('log-minimap');
  const scrollEl = document.getElementById('log-content');
  if (!canvas || !scrollEl) return;

  const lines = _logRawContent.split('\n');
  if (!lines.length) return;

  const dpr = window.devicePixelRatio || 1;
  const w   = canvas.offsetWidth;
  const h   = canvas.offsetHeight;
  if (!w || !h) return;

  canvas.width  = w * dpr;
  canvas.height = h * dpr;

  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);

  // Fond
  ctx.fillStyle = '#1e1e2e';
  ctx.fillRect(0, 0, w, h);

  const lineH     = h / lines.length;          // peut être < 1
  const minLineH  = Math.max(1, lineH);
  const matchSet  = new Set(_logSearchMatches.map(m => m.line));

  lines.forEach((line, i) => {
    const y   = i * lineH;
    const lh  = Math.max(0.8, lineH);
    // Largeur proportionnelle au contenu (max 100 chars → pleine largeur)
    const len = Math.min(line.trimEnd().length, 120);
    const lw  = len > 0 ? Math.max(6, (len / 120) * (w - 8)) : 0;

    if (matchSet.has(i)) {
      ctx.fillStyle = 'rgba(249,226,175,0.75)';
      ctx.fillRect(0, y, w, minLineH);
    } else if (lw > 0) {
      ctx.fillStyle = _logLevelColor(line);
      ctx.fillRect(4, y, lw, lh);
    }
  });

  // Indicateur de fenêtre visible
  const scrollTop = scrollEl.scrollTop;
  const scrollH   = scrollEl.scrollHeight;
  const clientH   = scrollEl.clientHeight;
  if (scrollH > clientH) {
    const vpTop = (scrollTop / scrollH) * h;
    const vpH   = Math.max(16, (clientH / scrollH) * h);
    ctx.fillStyle   = 'rgba(205,214,244,0.08)';
    ctx.fillRect(0, vpTop, w, vpH);
    ctx.strokeStyle = 'rgba(205,214,244,0.28)';
    ctx.lineWidth   = 1;
    ctx.strokeRect(0.5, vpTop + 0.5, w - 1, vpH - 1);
  }
}

function _logMinimapScrollTo(clientY) {
  const canvas   = document.getElementById('log-minimap');
  const scrollEl = document.getElementById('log-content');
  if (!canvas || !scrollEl) return;
  const rect    = canvas.getBoundingClientRect();
  const ratio   = Math.max(0, Math.min(1, (clientY - rect.top) / rect.height));
  scrollEl.scrollTop = ratio * scrollEl.scrollHeight - scrollEl.clientHeight / 2;
}

function _setupLogMinimap() {
  const canvas   = document.getElementById('log-minimap');
  const scrollEl = document.getElementById('log-content');
  if (!canvas || !scrollEl) return;

  scrollEl.addEventListener('scroll', _drawLogMinimap);

  canvas.addEventListener('mousedown', e => {
    e.preventDefault();
    _logMinimapDragging = true;
    _logMinimapScrollTo(e.clientY);
  });
  document.addEventListener('mousemove', e => {
    if (!_logMinimapDragging) return;
    _logMinimapScrollTo(e.clientY);
  });
  document.addEventListener('mouseup', () => { _logMinimapDragging = false; });

  if (window.ResizeObserver) {
    new ResizeObserver(() => { if (document.body.classList.contains('log-mode')) _drawLogMinimap(); })
      .observe(document.getElementById('log-body') || canvas.parentElement);
  }
}

function onLogSearchInput() {
  _logSearchQuery   = document.getElementById('log-search-input').value;
  _logSearchMatches = _findLogMatches(_logRawContent.split('\n'), _logSearchQuery);
  _logSearchIdx     = _logSearchMatches.length > 0 ? 0 : -1;
  _refreshLogContent();
}

function onLogSearchKeydown(e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    if (e.shiftKey) logSearchPrev(); else logSearchNext();
  } else if (e.key === 'Escape') {
    logSearchClear();
  }
}

function logSearchNext() {
  if (!_logSearchMatches.length) return;
  _logSearchIdx = (_logSearchIdx + 1) % _logSearchMatches.length;
  _refreshLogContent();
}

function logSearchPrev() {
  if (!_logSearchMatches.length) return;
  _logSearchIdx = (_logSearchIdx - 1 + _logSearchMatches.length) % _logSearchMatches.length;
  _refreshLogContent();
}

function logSearchClear() {
  document.getElementById('log-search-input').value = '';
  _logSearchQuery   = '';
  _logSearchMatches = [];
  _logSearchIdx     = -1;
  _refreshLogContent();
}


// ── Menu ──────────────────────────────────────────────────────────────────────
function toggleMenu() {
  document.getElementById('menu-dropdown').classList.toggle('open');
}

function closeMenu() {
  document.getElementById('menu-dropdown').classList.remove('open');
}

// ── Lien ──────────────────────────────────────────────────────────────────────
let _linkSelectionRange = null;   // range sauvegardée avant ouverture du dialogue
let _editingLinkNode    = null;   // <a> en cours d'édition (null = insertion)

/** Convertit un texte de titre en identifiant d'ancre (même logique que marked). */
function _headingToId(text) {
  return text.toLowerCase()
    .replace(/[^\w\s-]/g, '')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
}

/** Peuple la liste des signets ({#id}) dans le dialogue. */
let _linkFilePath = '';

function _setLinkFilePath(path) {
  _linkFilePath = path || '';
  const el = document.getElementById('link-file-path');
  if (!el) return;
  el.textContent = _linkFilePath || '—';
  el.classList.toggle('has-file', !!_linkFilePath);
  _updateLinkPreview();
}

async function browseLinkFile() {
  if (!window.pywebview || !window.pywebview.api) return;
  const tab      = getActiveTab();
  const baseDir  = tab?.path ? tab.path.replace(/\/[^/]+$/, '') : '';
  const s        = typeof loadSettings === 'function' ? loadSettings() : {};
  const startDir = s.projectDir || baseDir || '';
  try {
    const result = await window.pywebview.api.pick_include_file(startDir, baseDir);
    if (!result) return;
    _setLinkFilePath(result.rel || result.path);
    if (!document.getElementById('link-text').value.trim()) {
      const name = (result.rel || result.path).split('/').pop().replace(/\.[^.]+$/, '');
      document.getElementById('link-text').value = name;
    }
    _updateLinkPreview();
  } catch (e) {}
}

function _populateLinkBookmarks(selectedId) {
  const select = document.getElementById('link-bookmark-select');
  select.innerHTML = '';
  const none = document.createElement('option');
  none.value = ''; none.textContent = '— choisir un signet —';
  select.appendChild(none);
  document.querySelectorAll('#preview [data-bookmark]').forEach(el => {
    const id   = el.getAttribute('data-bookmark');
    const badge = el.querySelector('.heading-bookmark');
    const label = (badge ? el.textContent.replace(badge.textContent, '') : el.textContent).trim();
    const opt  = document.createElement('option');
    opt.value       = id;
    opt.textContent = `${el.tagName} – ${label} (#${id})`;
    if (id === selectedId) opt.selected = true;
    select.appendChild(opt);
  });
  const hasBookmarks = select.options.length > 1;
  document.getElementById('link-bookmark-row').style.display = hasBookmarks ? '' : 'none';
}

/** Peuple la liste des titres dans le dialogue. */
function _populateLinkHeadings(selectedId) {
  const select = document.getElementById('link-heading-select');
  select.innerHTML = '';
  document.querySelectorAll('#preview h1,#preview h2,#preview h3,#preview h4,#preview h5,#preview h6').forEach(h => {
    const level = h.tagName.toLowerCase();
    const text  = h.textContent.trim();
    const id    = _headingToId(text);
    const opt   = document.createElement('option');
    opt.value       = id;
    opt.textContent = `${level.toUpperCase()} – ${text}`;
    if (id === selectedId) opt.selected = true;
    select.appendChild(opt);
  });
  if (select.options.length === 0) {
    const opt = document.createElement('option');
    opt.value = ''; opt.textContent = '(aucun titre dans le document)';
    select.appendChild(opt);
  }
}

/** Détecte le type d'un href existant. */
function _detectLinkType(href) {
  if (!href) return 'url';
  if (href.startsWith('#')) {
    const id = href.slice(1);
    const headingIds = [...document.querySelectorAll(
      '#preview h1,#preview h2,#preview h3,#preview h4,#preview h5,#preview h6'
    )].map(h => _headingToId(h.textContent.trim()));
    return headingIds.includes(id) ? 'heading' : 'id';
  }
  if (/^https?:\/\/|^mailto:|^ftp:\/\//i.test(href)) return 'url';
  return 'file';
}

function openLinkDialog() {
  _editingLinkNode = null;
  _linkFilePath    = '';

  const sel = window.getSelection();
  _linkSelectionRange = (sel && sel.rangeCount > 0 && !sel.isCollapsed)
    ? sel.getRangeAt(0).cloneRange() : null;
  const selectedText = _linkSelectionRange ? _linkSelectionRange.toString().trim() : '';

  document.getElementById('link-dialog-title').textContent = 'Insérer un lien';
  document.getElementById('link-confirm-btn').textContent  = 'Insérer';
  document.getElementById('link-text').value = selectedText;
  document.getElementById('link-url').value  = '';
  document.getElementById('link-id').value   = '';
  _setLinkFilePath('');
  _populateLinkHeadings('');
  selectLinkType('url');
  document.getElementById('link-overlay').classList.add('open');
  setTimeout(() => document.getElementById('link-url').focus(), 50);
}

function openLinkEditDialog(anchor) {
  _editingLinkNode = anchor;
  _linkSelectionRange = null;

  const href = anchor.getAttribute('href') || '';
  const text = anchor.textContent.trim();
  const type = _detectLinkType(href);

  document.getElementById('link-dialog-title').textContent = 'Modifier le lien';
  document.getElementById('link-confirm-btn').textContent  = 'Mettre à jour';
  document.getElementById('link-text').value = text;
  document.getElementById('link-url').value  = '';
  document.getElementById('link-id').value   = '';
  _setLinkFilePath('');
  _populateLinkHeadings('');

  if (type === 'url') {
    document.getElementById('link-url').value = href;
  } else if (type === 'file') {
    _setLinkFilePath(href);
  } else if (type === 'id') {
    document.getElementById('link-id').value = href.slice(1);
  }
  // heading : _populateLinkHeadings gère la sélection
  _populateLinkHeadings(type === 'heading' ? href.slice(1) : '');
  selectLinkType(type);
  // Pour 'id', pré-sélectionner le signet dans la liste si présent
  if (type === 'id') _populateLinkBookmarks(href.slice(1));
  document.getElementById('link-overlay').classList.add('open');
}

function selectLinkType(type) {
  ['url','file','id','heading'].forEach(t => {
    document.getElementById(`link-type-${t}`).classList.toggle('active', t === type);
    document.getElementById(`link-${t}-row`).style.display = t === type ? '' : 'none';
  });
  if (type === 'id') {
    const currentId = document.getElementById('link-id').value.trim();
    _populateLinkBookmarks(currentId);
  }
  _updateLinkPreview();
}

function _updateLinkPreview() {
  const text = document.getElementById('link-text').value || 'texte';
  let   href = '';
  if (document.getElementById('link-type-url').classList.contains('active')) {
    href = document.getElementById('link-url').value.trim() || 'https://…';
  } else if (document.getElementById('link-type-file').classList.contains('active')) {
    href = _linkFilePath || 'chemin/vers/fichier';
  } else if (document.getElementById('link-type-id').classList.contains('active')) {
    const id = document.getElementById('link-id').value.trim();
    href = id ? `#${id}` : '#id';
  } else {
    const sel = document.getElementById('link-heading-select');
    href = sel.value ? `#${sel.value}` : '#titre';
  }
  document.getElementById('link-preview').textContent = `[${text}](${href})`;
}

function closeLinkDialog(e) {
  if (e && e.target !== document.getElementById('link-overlay')) return;
  document.getElementById('link-overlay').classList.remove('open');
  _editingLinkNode = null;
}

function confirmLinkDialog() {
  const text = document.getElementById('link-text').value.trim() || 'lien';
  let   href = '';
  if (document.getElementById('link-type-url').classList.contains('active')) {
    href = document.getElementById('link-url').value.trim();
  } else if (document.getElementById('link-type-file').classList.contains('active')) {
    href = _linkFilePath || '';
  } else if (document.getElementById('link-type-id').classList.contains('active')) {
    const id = document.getElementById('link-id').value.trim();
    href = id ? `#${id}` : '#';
  } else {
    const val = document.getElementById('link-heading-select').value;
    href = val ? `#${val}` : '#';
  }
  if (!href) return;

  document.getElementById('link-overlay').classList.remove('open');

  // ── Mode édition : mettre à jour le <a> existant
  if (_editingLinkNode) {
    _editingLinkNode.setAttribute('href', href);
    _editingLinkNode.textContent = text;
    _editingLinkNode = null;
    syncPreviewToContent();
    return;
  }

  // ── Mode insertion
  const tab = getActiveTab();
  if (!tab) return;
  const md = `[${text}](${href})`;

  if (state.sourceMode) {
    const editor = document.getElementById('source-editor');
    const pos    = editor.selectionEnd;
    editor.value = editor.value.slice(0, pos) + md + editor.value.slice(pos);
    editor.selectionStart = editor.selectionEnd = pos + md.length;
    editor.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    if (_linkSelectionRange && preview.contains(_linkSelectionRange.commonAncestorContainer)) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(_linkSelectionRange);
    }
    document.execCommand('insertHTML', false, `<a href="${href}">${text}</a>`);
    syncPreviewToContent();
  }
  _linkSelectionRange = null;
}

// ── Menu contextuel lien ───────────────────────────────────────────────────────
let _linkContextAnchor = null;

function _isLocalHref(href) {
  if (!href) return false;
  // Exclure les ancres internes, URLs externes, protocoles spéciaux
  if (href.startsWith('#')) return false;
  if (/^[a-z][a-z0-9+\-.]*:/i.test(href)) return false; // http:, mailto:, ftp:, etc.
  return true;
}

function showLinkContextMenu(e, anchor) {
  _linkContextAnchor = anchor;
  const href = anchor.getAttribute('href') || '';
  const isLocal = _isLocalHref(href);
  document.getElementById('link-ctx-open').style.display     = isLocal ? '' : 'none';
  document.getElementById('link-ctx-open-sep').style.display = isLocal ? '' : 'none';
  const menu = document.getElementById('link-context-menu');
  _showMenuAt(menu, e.clientX, e.clientY, m => m.classList.add('open'));
  e.preventDefault();
  e.stopPropagation();
}

async function openLinkFileFromMenu() {
  if (!_linkContextAnchor) return;
  const href = (_linkContextAnchor.getAttribute('href') || '').replace(/\?[^#]*/, '').replace(/#.*$/, '');
  _linkContextAnchor = null;
  if (!href || !_isLocalHref(href)) return;

  // Résoudre le chemin absolu à partir du fichier courant
  const tab = getActiveTab();
  if (!tab || !tab.path) { alert('Enregistrez le fichier courant avant d\'ouvrir ce lien.'); return; }

  const norm  = s => s.replace(/\\/g, '/');
  const base  = norm(tab.path).replace(/\/[^/]+$/, ''); // dossier du fichier courant
  const parts = base.split('/');
  norm(href).split('/').forEach(p => {
    if (p === '..') parts.pop();
    else if (p && p !== '.') parts.push(p);
  });
  const absPath = parts.join('/');

  try {
    const data = await window.pywebview.api.open_file_by_path(absPath);
    if (!data || data.error) { alert(`Impossible d'ouvrir : ${absPath}`); return; }
    const existing = state.tabs.find(t => norm(t.path) === norm(data.path));
    if (existing) { switchToTab(existing.id); return; }
    const newTab = createTab(data.name, data.path, data.content);
    switchToTab(newTab.id);
  } catch (err) {
    alert(`Erreur lors de l'ouverture : ${err}`);
  }
}

function copyLinkFromMenu() {
  if (!_linkContextAnchor) return;
  const href = _linkContextAnchor.getAttribute('href') || '';
  _linkContextAnchor = null;
  if (!href) return;
  navigator.clipboard.writeText(href).then(() => showToast('Lien copié'));
}

function hideLinkContextMenu() {
  document.getElementById('link-context-menu').classList.remove('open');
}

function editLinkFromMenu() {
  if (!_linkContextAnchor) return;
  openLinkEditDialog(_linkContextAnchor);
  _linkContextAnchor = null;
}

function deleteLinkFromMenu() {
  if (!_linkContextAnchor) return;
  // Remplace le <a> par son contenu texte
  const parent = _linkContextAnchor.parentNode;
  while (_linkContextAnchor.firstChild) parent.insertBefore(_linkContextAnchor.firstChild, _linkContextAnchor);
  parent.removeChild(_linkContextAnchor);
  _linkContextAnchor = null;
  syncPreviewToContent();
}

// ── Signet (heading bookmark) ─────────────────────────────────────────────────
let _bookmarkHeadingNode = null;

function openBookmarkDialog(headingEl) {
  _bookmarkHeadingNode = headingEl || null;
  const existing = headingEl ? (headingEl.getAttribute('data-bookmark') || '') : '';
  const input = document.getElementById('bookmark-id-input');
  input.value = existing;
  document.getElementById('bookmark-dialog-title').textContent = existing ? 'Modifier le signet' : 'Définir un signet';
  document.getElementById('bookmark-remove-btn').style.display = existing ? 'inline-flex' : 'none';
  document.getElementById('bookmark-overlay').classList.add('open');
  setTimeout(() => input.focus(), 50);
}

function closeBookmarkDialog(e) {
  if (e && e.target !== document.getElementById('bookmark-overlay')) return;
  document.getElementById('bookmark-overlay').classList.remove('open');
  _bookmarkHeadingNode = null;
}

function confirmBookmarkDialog() {
  if (!_bookmarkHeadingNode) return;
  const id = document.getElementById('bookmark-id-input').value.trim().replace(/\s+/g, '-');
  document.getElementById('bookmark-overlay').classList.remove('open');
  if (!id) return;
  _bookmarkHeadingNode.setAttribute('id', id);
  _bookmarkHeadingNode.setAttribute('data-bookmark', id);
  // Mettre à jour ou créer le badge
  let badge = _bookmarkHeadingNode.querySelector('.heading-bookmark');
  if (!badge) {
    badge = document.createElement('span');
    badge.className = 'heading-bookmark';
    badge.contentEditable = 'false';
    _bookmarkHeadingNode.appendChild(badge);
  }
  badge.textContent = '#' + id;
  badge.title = 'Signet : #' + id;
  _bookmarkHeadingNode = null;
  syncPreviewToContent();
}

function removeBookmark() {
  if (!_bookmarkHeadingNode) return;
  document.getElementById('bookmark-overlay').classList.remove('open');
  _bookmarkHeadingNode.removeAttribute('data-bookmark');
  _bookmarkHeadingNode.removeAttribute('id');
  const badge = _bookmarkHeadingNode.querySelector('.heading-bookmark');
  if (badge) badge.remove();
  _bookmarkHeadingNode = null;
  syncPreviewToContent();
}

function triggerBookmark() {
  const preview = document.getElementById('preview');

  if (state.sourceMode) {
    // Trouver si la ligne courante est un titre
    const ta      = document.getElementById('source-editor');
    const pos     = ta.selectionStart;
    const lines   = ta.value.split('\n');
    const lineIdx = ta.value.substring(0, pos).split('\n').length - 1;
    const line    = lines[lineIdx];
    const m       = line.match(/^(#{1,6})\s+(.+?)(\s*\{#[^}]*\})?$/);
    if (!m) return;

    const level      = m[1].length;
    const titleText  = m[2].trim();
    // Chercher le heading correspondant dans le preview (même niveau, texte proche)
    const headings   = [...preview.querySelectorAll(`h${level}`)];
    const heading    = headings.find(h => {
      const badge = h.querySelector('.heading-bookmark');
      const txt   = badge
        ? h.textContent.replace(badge.textContent, '').trim()
        : h.textContent.trim();
      return txt === titleText;
    });
    if (!heading) return;
    openBookmarkDialog(heading);
  } else {
    // Mode prévisualisation : heading sous le curseur
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    let node = sel.anchorNode;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
    const heading = node.closest('h1,h2,h3,h4,h5,h6');
    if (!heading || !preview.contains(heading)) return;
    openBookmarkDialog(heading);
  }
}

// ── Raw HTML ──────────────────────────────────────────────────────────────────
let _rhContextNode    = null;
let _rhInsertionRange = null;

function openRawHtmlDialog(node) {
  _rhContextNode = node || null;
  if (!node) {
    const sel = window.getSelection();
    const preview = document.getElementById('preview');
    _rhInsertionRange = (sel && sel.rangeCount > 0 && preview.contains(sel.anchorNode))
      ? sel.getRangeAt(0).cloneRange()
      : null;
  } else {
    _rhInsertionRange = null;
  }
  const ta = document.getElementById('rh-code');
  ta.value = node ? (node.getAttribute('data-html') || '').replace(/&quot;/g, '"') : '';
  document.getElementById('rh-dialog-title').textContent = node ? 'Modifier le bloc HTML' : 'Insérer un bloc HTML';
  document.getElementById('rh-confirm-btn').textContent  = node ? 'Modifier' : 'Insérer';
  document.getElementById('rh-overlay').classList.add('open');
  setTimeout(() => ta.focus(), 50);
}

function closeRawHtmlDialog(e) {
  if (e && e.target !== document.getElementById('rh-overlay')) return;
  document.getElementById('rh-overlay').classList.remove('open');
  _rhContextNode = null;
}

function confirmRawHtmlDialog() {
  const code = document.getElementById('rh-code').value.trim();
  if (!code) return;
  document.getElementById('rh-overlay').classList.remove('open');

  if (_rhContextNode) {
    _rhContextNode.setAttribute('data-html', code.replace(/"/g, '&quot;'));
    _rhContextNode.querySelector('.md-rawhtml-content').innerHTML = code;
    _rhContextNode = null;
    syncPreviewToContent();
    return;
  }

  const safeRaw = code.replace(/"/g, '&quot;');
  const wrapper = document.createElement('div');
  wrapper.innerHTML = `<div class="md-rawhtml" data-html="${safeRaw}" contenteditable="false"><div class="md-rawhtml-content">${code}</div></div>`;
  const node = wrapper.firstChild;
  _insertAtCursor(node, _rhInsertionRange);
  _rhInsertionRange = null;
  syncPreviewToContent();
  const _tab = getActiveTab();
  if (_tab) updatePreview(_tab.content);
}

function showRawHtmlContextMenu(e, node) {
  _rhContextNode = node;
  const menu = document.getElementById('rh-context-menu');
  _showMenuAt(menu, e.clientX, e.clientY, m => m.classList.add('open'));
}

function hideRawHtmlContextMenu() {
  document.getElementById('rh-context-menu').classList.remove('open');
}

function editRawHtmlFromMenu() {
  const node = _rhContextNode;
  _rhContextNode = null;
  if (node) openRawHtmlDialog(node);
}

function deleteRawHtmlFromMenu() {
  if (_rhContextNode) { _rhContextNode.remove(); _rhContextNode = null; syncPreviewToContent(); }
}

// ── Commentaire HTML ──────────────────────────────────────────────────────────
let _commentContextNode = null;

/** Construit le HTML de prévisualisation d'un commentaire (régulier ou TODO). */
function _buildCommentHtml(text) {
  const safeText = text.replace(/"/g, '&quot;');
  const todoMatch = text.match(/^(TODO|DONE)\s*:?\s*([\s\S]*)/i);
  if (todoMatch) {
    const kind = todoMatch[1].toUpperCase();
    const body = (todoMatch[2] || '').trim();
    const done = kind === 'DONE';
    return `<div class="md-comment md-comment--todo${done ? ' md-comment--done' : ''}" data-comment="${safeText}" data-todo-kind="${kind}" contenteditable="false">`
      + `<input type="checkbox" class="md-todo-checkbox"${done ? ' checked' : ''}>`
      + `<span class="md-todo-tag">${kind}</span>`
      + `<span class="md-todo-body">${body}</span>`
      + `</div>`;
  }
  return `<div class="md-comment" data-comment="${safeText}" contenteditable="false">`
    + `<span class="md-comment-delim"><!--</span>`
    + `<span class="md-comment-text"> ${text || ''} </span>`
    + `<span class="md-comment-delim">--></span>`
    + `</div>`;
}

function insertHtmlComment() {
  const placeholder = 'commentaire';
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const pos = ta.selectionStart;
    const lineStart = ta.value.lastIndexOf('\n', pos - 1) + 1;
    const insertion = `<!-- ${placeholder} -->\n`;
    ta.value = ta.value.substring(0, lineStart) + insertion + ta.value.substring(lineStart);
    ta.selectionStart = lineStart + 5;
    ta.selectionEnd   = lineStart + 5 + placeholder.length;
    ta.focus();
    onSourceInput();
  } else {
    const wrapper = document.createElement('div');
    wrapper.innerHTML = _buildCommentHtml(placeholder);
    const node = wrapper.firstChild;
    _insertAtCursor(node, _scInsertionRange);
    _scInsertionRange = null;
    syncPreviewToContent();
    const tab = getActiveTab();
    if (tab) updatePreview(tab.content);
  }
}

function insertTodoComment() {
  const placeholder = 'TODO: tâche à faire';
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const pos = ta.selectionStart;
    const lineStart = ta.value.lastIndexOf('\n', pos - 1) + 1;
    const insertion = `<!-- ${placeholder} -->\n`;
    ta.value = ta.value.substring(0, lineStart) + insertion + ta.value.substring(lineStart);
    // Sélectionner le texte après "TODO: "
    ta.selectionStart = lineStart + 9;
    ta.selectionEnd   = lineStart + insertion.length - 5;
    ta.focus();
    onSourceInput();
  } else {
    const wrapper = document.createElement('div');
    wrapper.innerHTML = _buildCommentHtml(placeholder);
    const node = wrapper.firstChild;
    _insertAtCursor(node, _scInsertionRange);
    _scInsertionRange = null;
    syncPreviewToContent();
    const tab = getActiveTab();
    if (tab) updatePreview(tab.content);
  }
}

// ── Panneau TODO ──────────────────────────────────────────────────────────────
const TODO_STORAGE_KEY = 'md-editor-todo-checked';
let _todoFilter   = 'all';
let _todoScanDir  = null;   // répertoire scanné
let _todoScanItems = [];    // todos issus du scan fichier [{rel,tabKey,tabName,text,done,fromScan}]
let _todoJsonPath  = null;  // chemin du todo_tracking.json

function _todoGetChecked() {
  try { return JSON.parse(localStorage.getItem(TODO_STORAGE_KEY) || '{}'); }
  catch { return {}; }
}
function _todoSetChecked(map) {
  localStorage.setItem(TODO_STORAGE_KEY, JSON.stringify(map));
}
function _todoKey(tabKey, text) {
  return tabKey + '||' + text;
}

// TODOs des onglets ouverts en mémoire
function _collectMemTodos() {
  const RE = /<!--\s*(TODO\s*:?\s*[\s\S]*?)-->/gi;
  const todos = [];
  for (const tab of state.tabs) {
    const src = tab.content || '';
    let m;
    RE.lastIndex = 0;
    while ((m = RE.exec(src)) !== null) {
      const raw = m[1].trim();
      if (/^TODO\s*:?/i.test(raw)) {
        todos.push({
          tabKey:   tab.path || tab.name,
          tabName:  tab.name,
          text:     raw.replace(/^TODO\s*:?\s*/i, '').trim() || '(sans texte)',
          fromScan: false,
        });
      }
    }
  }
  return todos;
}

// Fusion mémoire + scan (déduplique par tabKey+text)
function _collectTodos() {
  const memItems = _collectMemTodos();
  if (_todoScanItems.length === 0) return memItems;

  // Index des clés déjà présentes en mémoire
  const memKeys = new Set(memItems.map(t => _todoKey(t.tabKey, t.text)));
  // Ajoute les items scannés absents des onglets ouverts
  const extra = _todoScanItems.filter(t => !memKeys.has(_todoKey(t.tabKey, t.text)));
  return [...memItems, ...extra];
}

function _renderTodoList() {
  const container = document.getElementById('todo-list-container');
  if (!container) return;

  const checked  = _todoGetChecked();
  const allTodos = _collectTodos();

  // Synchronise done depuis localStorage pour les items mémoire,
  // et depuis _todoScanItems pour les items fichier
  const enriched = allTodos.map(t => {
    const isDone = t.fromScan
      ? t.done
      : !!checked[_todoKey(t.tabKey, t.text)];
    return { ...t, isDone };
  });

  const visible = enriched.filter(t => {
    if (_todoFilter === 'todo') return !t.isDone;
    if (_todoFilter === 'done') return t.isDone;
    return true;
  });

  const doneCount = enriched.filter(t => t.isDone).length;
  const counter   = document.getElementById('todo-counter');
  if (counter) counter.textContent = `${doneCount} / ${enriched.length} faite${doneCount !== 1 ? 's' : ''}`;

  // Info répertoire scanné
  const scanInfoId = 'todo-scan-info';
  let infoEl = document.getElementById(scanInfoId);
  if (_todoScanDir) {
    if (!infoEl) {
      infoEl = document.createElement('div');
      infoEl.id = scanInfoId;
      infoEl.className = 'todo-scan-info';
      container.before(infoEl);
    }
    const dirName = _todoScanDir.replace(/\\/g, '/').split('/').pop();
    infoEl.innerHTML = `<svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.4" width="11" height="11"><path d="M1 3.5h12M1 3.5v8h12v-8M3 3.5V2.5a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v1"/></svg>Scanné : <strong>${_escHtml(dirName)}</strong> — ${_todoScanItems.length} TODO(s) trouvé(s)`;
  } else if (infoEl) {
    infoEl.remove();
  }

  if (visible.length === 0) {
    container.innerHTML = `<div class="todo-empty">${
      enriched.length === 0
        ? 'Aucun TODO dans les fichiers ouverts.<br><small style="opacity:.6">Utilisez "Scanner répertoire" pour chercher dans tous les fichiers.</small>'
        : _todoFilter === 'done' ? 'Aucune tâche faite.' : 'Aucune tâche à faire !'
    }</div>`;
    return;
  }

  const grouped = {};
  for (const t of visible) {
    const gk = t.tabKey;
    if (!grouped[gk]) grouped[gk] = { name: t.tabName, items: [], fromScan: t.fromScan, path: t.path || t.tabKey };
    grouped[gk].items.push(t);
  }

  let html = '';
  for (const [, grp] of Object.entries(grouped)) {
    const srcTag = grp.fromScan
      ? `<span class="todo-source-tag todo-source-tag--file">fichier</span>`
      : `<span class="todo-source-tag todo-source-tag--mem">onglet</span>`;
    const filePath = _escHtml(grp.path || '');
    html += `<div class="todo-group"><div class="todo-group-label">${srcTag} <span class="todo-file-link" data-todo-open="${filePath}" title="Ouvrir ${_escHtml(grp.name)}">${_escHtml(grp.name)}</span></div>`;
    for (const t of grp.items) {
      const key       = _todoKey(t.tabKey, t.text);
      const doneClass = t.isDone ? ' done' : '';
      html += `<div class="todo-item${doneClass}" data-todo-key="${_escHtml(key)}">
        <div class="todo-item-check">
          <svg viewBox="0 0 10 10" fill="none" stroke="white" stroke-width="1.8" width="9" height="9">
            <polyline points="1.5,5 4,7.5 8.5,2"/>
          </svg>
        </div>
        <div class="todo-item-text">${_escHtml(t.text)}</div>
      </div>`;
    }
    html += `</div>`;
  }
  container.innerHTML = html;
  _updateTodoPanelBtn();
}

function _escHtml(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── Panneau Commentaires (annotations liées à une sélection) ───────────────────
function _collectAnnotations(content) {
  const RE = /\[((?:[^\[\]]|\[[^\]]*\])*)\]\{\.comment comment-id="([^"]*)" comment-text="((?:\\.|[^"\\])*)"\}/g;
  const items = [];
  let m;
  while ((m = RE.exec(content || '')) !== null) {
    items.push({ id: m[2], anchorText: m[1], text: _annotUnescape(m[3]) });
  }
  return items;
}

function _annotAnchorPreview(raw) {
  const plain = (raw || '')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/\*([^*]+)\*/g, '$1')
    .replace(/`([^`]+)`/g, '$1')
    .replace(/[_~]/g, '')
    .trim();
  const truncated = plain.length > 90 ? plain.slice(0, 90) + '…' : plain;
  return _annotHtmlEscape(truncated);
}

function updateAnnotationsList() {
  const body = document.getElementById('annotations-body');
  if (!body) return;

  const tab   = getActiveTab();
  const items = tab ? _collectAnnotations(tab.content) : [];

  const counter = document.getElementById('annotations-counter');
  if (counter) counter.textContent = items.length ? String(items.length) : '';

  if (!items.length) {
    body.innerHTML = `<div class="annot-empty">Aucun commentaire.<br>
      <small style="opacity:.6">Sélectionnez du texte, puis « Ajouter un commentaire » dans le menu contextuel.</small></div>`;
    return;
  }

  body.innerHTML = items.map(it => `
    <div class="annot-item" data-annot-id="${it.id}">
      <div class="annot-anchor" onclick="selectAnnotationAnchor('${it.id}')" title="Aller au texte sélectionné">${_annotAnchorPreview(it.anchorText)}</div>
      <textarea class="annot-text-input" oninput="onAnnotationTextInput('${it.id}', this.value)" placeholder="Écrire un commentaire…">${_annotHtmlEscape(it.text)}</textarea>
      <button class="annot-delete-btn" onclick="deleteAnnotation('${it.id}')" title="Supprimer le commentaire">×</button>
    </div>`).join('');
}

function _focusAnnotationInput(id) {
  const section = document.getElementById('annotations-section');
  if (section) section.classList.remove('collapsed');
  setTimeout(() => {
    const item = document.querySelector(`.annot-item[data-annot-id="${id}"]`);
    if (!item) return;
    item.scrollIntoView({ block: 'nearest' });
    const ta = item.querySelector('.annot-text-input');
    if (ta) ta.focus();
  }, 0);
}

function onAnnotationTextInput(id, value) {
  const tab = getActiveTab();
  if (!tab) return;
  const RE = new RegExp(`(comment-id="${id}" comment-text=")((?:\\\\.|[^"\\\\])*)(")`);
  const newContent = tab.content.replace(RE, (_m, p1, _old, p3) => p1 + _annotEscape(value) + p3);
  if (newContent === tab.content) return;
  tab.content  = newContent;
  tab.modified = (tab.content !== tab.savedContent);
  renderTabList();
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    if (ta) { ta.value = tab.content; _updateSourceHighlight(); }
  } else {
    const span = document.querySelector(`#preview .md-annotation[data-annot-id="${id}"]`);
    if (span) span.setAttribute('data-annot-text', value);
  }
}

function deleteAnnotation(id) {
  const tab = getActiveTab();
  if (!tab) return;
  if (!state.sourceMode) {
    const span = document.querySelector(`#preview .md-annotation[data-annot-id="${id}"]`);
    if (span) {
      const parent = span.parentNode;
      while (span.firstChild) parent.insertBefore(span.firstChild, span);
      parent.removeChild(span);
      parent.normalize();
    }
    syncPreviewToContent();
  } else {
    const RE = new RegExp(`\\[((?:[^\\[\\]]|\\[[^\\]]*\\])*)\\]\\{\\.comment comment-id="${id}" comment-text="(?:\\\\.|[^"\\\\])*"\\}`);
    const newContent = tab.content.replace(RE, '$1');
    if (newContent === tab.content) return;
    tab.content  = newContent;
    tab.modified = (tab.content !== tab.savedContent);
    renderTabList();
    const ta = document.getElementById('source-editor');
    if (ta) { ta.value = tab.content; _updateSourceHighlight(); }
  }
  updateAnnotationsList();
}

function selectAnnotationAnchor(id) {
  const tab = getActiveTab();
  if (!tab) return;
  if (state.sourceMode) {
    const ta   = document.getElementById('source-editor');
    const pane = document.getElementById('source-scroll');
    if (!ta) return;
    const RE = new RegExp(`\\[((?:[^\\[\\]]|\\[[^\\]]*\\])*)\\]\\{\\.comment comment-id="${id}" comment-text="(?:\\\\.|[^"\\\\])*"\\}`);
    const m = RE.exec(tab.content);
    if (!m) return;
    ta.setSelectionRange(m.index, m.index + m[0].length);
    ta.focus();
    if (pane) {
      const cs          = getComputedStyle(ta);
      const paddingTop  = parseFloat(cs.paddingTop) || 24;
      const lineHeight  = parseFloat(cs.lineHeight) || (parseFloat(cs.fontSize) * 1.6);
      const linesBefore = ta.value.substring(0, m.index).split('\n').length - 1;
      const targetScrollTop = Math.max(0, paddingTop + linesBefore * lineHeight - pane.clientHeight / 2 + lineHeight / 2);
      setTimeout(() => { pane.scrollTop = targetScrollTop; }, 0);
    }
  } else {
    const preview = document.getElementById('preview');
    const span = preview.querySelector(`.md-annotation[data-annot-id="${id}"]`);
    if (!span) return;
    span.scrollIntoView({ block: 'center', behavior: 'smooth' });
    const range = document.createRange();
    range.selectNodeContents(span);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    preview.focus();
    span.classList.add('md-annotation--flash');
    setTimeout(() => span.classList.remove('md-annotation--flash'), 900);
  }
}

function toggleAnnotationsSection() {
  const section = document.getElementById('annotations-section');
  if (!section) return;
  const collapsed = section.classList.toggle('collapsed');
  const handle = document.getElementById('meta-annot-resize-handle');
  if (handle) handle.classList.toggle('disabled', collapsed);
}

async function _todoOpenFile(path) {
  if (!path) return;
  // Si l'onglet est déjà ouvert, on bascule simplement dessus
  const norm = p => p.replace(/\\/g, '/').toLowerCase();
  const existing = state.tabs.find(t => t.path && norm(t.path) === norm(path));
  if (existing) { switchToTab(existing.id); return; }
  // Sinon on charge via l'API Python
  if (!window.pywebview || !window.pywebview.api) return;
  try {
    const data = await window.pywebview.api.open_file_by_path(path);
    if (!data || data.error) { alert(`Impossible d'ouvrir : ${path}`); return; }
    const tab = createTab(data.name, data.path, data.content);
    switchToTab(tab.id);
  } catch (e) {
    console.error('_todoOpenFile:', e);
  }
}

async function _persistScanTracking() {
  if (!_todoJsonPath || !window.pywebview) return;
  try {
    await window.pywebview.api.save_todo_tracking(_todoJsonPath, _todoScanItems.map(t => ({
      path: t.path || '',
      rel:  t.tabKey,
      name: t.tabName,
      text: t.text,
      done: t.done,
    })));
  } catch (e) { console.warn('save_todo_tracking:', e); }
}

function todoToggle(key) {
  // Cherche d'abord dans les items scannés
  const scanIdx = _todoScanItems.findIndex(t => _todoKey(t.tabKey, t.text) === key);
  if (scanIdx >= 0) {
    _todoScanItems[scanIdx].done = !_todoScanItems[scanIdx].done;
    _persistScanTracking();
  } else {
    const checked = _todoGetChecked();
    checked[key] = !checked[key];
    if (!checked[key]) delete checked[key];
    _todoSetChecked(checked);
  }
  _renderTodoList();
}

function todoSetFilter(f) {
  _todoFilter = f;
  document.querySelectorAll('.todo-filter-btn').forEach(b => b.classList.remove('active'));
  const btn = document.getElementById('todo-filter-' + f);
  if (btn) btn.classList.add('active');
  _renderTodoList();
}

function todoClearDone() {
  // Items mémoire
  const checked = _todoGetChecked();
  for (const t of _collectMemTodos()) delete checked[_todoKey(t.tabKey, t.text)];
  _todoSetChecked(checked);
  // Items scannés
  _todoScanItems.forEach(t => { t.done = false; });
  _persistScanTracking();
  _renderTodoList();
}

async function todoScanDirectory() {
  if (!window.pywebview || !window.pywebview.api) return;

  // Détermine le répertoire : sidebar actif, sinon répertoire du fichier actif, sinon choix
  let dir = _sidebarCurrentDir;
  if (!dir) {
    const tab = getActiveTab();
    if (tab && tab.path) dir = tab.path.replace(/\\/g, '/').split('/').slice(0, -1).join('/');
  }
  if (!dir) {
    try { dir = await window.pywebview.api.get_default_dir(); } catch { return; }
  }

  const scanBtn = document.getElementById('todo-scan-btn');
  if (scanBtn) scanBtn.classList.add('scanning');

  try {
    const res = await window.pywebview.api.scan_todos(dir);
    if (res && res.error) { alert('Erreur scan : ' + res.error); return; }

    _todoScanDir   = res.dir;
    _todoJsonPath  = res.json_path;
    _todoScanItems = (res.todos || []).map(t => ({
      path:     t.path,
      tabKey:   t.rel,
      tabName:  t.name,
      text:     t.text,
      done:     t.done,
      fromScan: true,
    }));
    _renderTodoList();
  } catch (e) {
    console.error('todoScanDirectory:', e);
  } finally {
    if (scanBtn) scanBtn.classList.remove('scanning');
  }
}

let _todoListenerAttached = false;

async function openTodoPanel() {
  _todoFilter = 'all';
  document.querySelectorAll('.todo-filter-btn').forEach(b => b.classList.remove('active'));
  const all = document.getElementById('todo-filter-all');
  if (all) all.classList.add('active');

  if (!_todoListenerAttached) {
    const container = document.getElementById('todo-list-container');
    if (container) {
      container.addEventListener('click', e => {
        const link = e.target.closest('.todo-file-link[data-todo-open]');
        if (link) { _todoOpenFile(link.getAttribute('data-todo-open')); return; }
        const item = e.target.closest('.todo-item[data-todo-key]');
        if (item) todoToggle(item.getAttribute('data-todo-key'));
      });
      _todoListenerAttached = true;
    }
  }

  document.getElementById('todo-overlay').classList.add('open');
  _renderTodoList();

  // Charge todo_tracking.json du répertoire courant si pas déjà scanné
  if (_todoScanItems.length === 0 && window.pywebview && window.pywebview.api) {
    let dir = _sidebarCurrentDir;
    if (!dir) {
      const tab = getActiveTab();
      if (tab && tab.path) dir = tab.path.replace(/\\/g, '/').split('/').slice(0, -1).join('/');
    }
    if (!dir) {
      try { dir = await window.pywebview.api.get_default_dir(); } catch { dir = null; }
    }
    if (dir) {
      try {
        const data = await window.pywebview.api.load_todo_json(dir);
        if (data && !data.error && Array.isArray(data.todos)) {
          _todoScanDir   = data.dir;
          _todoJsonPath  = data.json_path;
          _todoScanItems = data.todos.map(t => ({
            path:     t.path || '',
            tabKey:   t.rel,
            tabName:  t.name,
            text:     t.text,
            done:     t.done,
            fromScan: true,
          }));
          _renderTodoList();
        }
      } catch (e) { console.warn('openTodoPanel load_todo_json:', e); }
    }
  }
}

function closeTodoPanel() {
  document.getElementById('todo-overlay').classList.remove('open');
}

function _updateTodoPanelBtn() {
  const total = _collectTodos().length;
  const btn = document.getElementById('todo-panel-btn');
  if (!btn) return;
  btn.classList.toggle('has-todos', total > 0);
}

document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    const overlay = document.getElementById('todo-overlay');
    if (overlay && overlay.classList.contains('open')) { closeTodoPanel(); e.stopPropagation(); }
  }
}, true);

let _editingComment = false;

function _editCommentText(commentNode) {
  const isTodo   = commentNode.classList.contains('md-comment--todo');
  const editSpan = isTodo
    ? commentNode.querySelector('.md-todo-body')
    : commentNode.querySelector('.md-comment-text');
  if (!editSpan) return;

  const fullSaved = (commentNode.getAttribute('data-comment') || '').replace(/&quot;/g, '"');
  const kind      = commentNode.getAttribute('data-todo-kind') || 'TODO';
  const savedBody = isTodo ? fullSaved.replace(/^(?:TODO|DONE)\s*:?\s*/i, '').trim() : fullSaved;

  _editingComment = true;
  commentNode.contentEditable = 'true';
  editSpan.contentEditable = 'true';
  // Différer le focus d'un tick pour que le browser traite le changement contentEditable
  setTimeout(() => {
    if (_finished) return;
    editSpan.focus();
    const range = document.createRange();
    range.selectNodeContents(editSpan);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }, 0);

  let _finished = false;

  const finish = () => {
    if (_finished) return;
    _finished = true;
    document.removeEventListener('mousedown', onDocMousedown, true);
    editSpan.removeEventListener('keydown', onKey);

    const newBody = editSpan.textContent.trim() || (isTodo ? 'tâche' : 'commentaire');
    if (isTodo) {
      commentNode.setAttribute('data-comment', `${kind}: ${newBody}`);
      editSpan.textContent = newBody;
    } else {
      commentNode.setAttribute('data-comment', newBody);
      editSpan.textContent = ` ${newBody} `;
    }
    commentNode.contentEditable = 'false';
    editSpan.contentEditable = 'false';
    _editingComment = false;

    // Mise à jour directe de tab.content par remplacement de chaîne
    const tab = getActiveTab();
    if (tab) {
      const oldMd = `<!-- ${fullSaved} -->`;
      const newMd = isTodo ? `<!-- ${kind}: ${newBody} -->` : `<!-- ${newBody} -->`;
      if (tab.content.includes(oldMd)) {
        tab.content = tab.content.replace(oldMd, newMd);
        tab.modified = (tab.content !== tab.savedContent);
        renderTabList();
      } else {
        syncPreviewToContent();
      }
    }
  };

  // mousedown en phase de capture : fiable même si blur ne se déclenche pas dans WebView2
  const onDocMousedown = (e) => {
    if (!commentNode.contains(e.target)) finish();
  };

  const onKey = e => {
    if (e.key === 'Enter') { e.preventDefault(); finish(); }
    if (e.key === 'Escape') {
      _finished = true;
      document.removeEventListener('mousedown', onDocMousedown, true);
      editSpan.removeEventListener('keydown', onKey);
      editSpan.textContent = isTodo ? savedBody : ` ${savedBody} `;
      commentNode.contentEditable = 'false';
      editSpan.contentEditable = 'false';
      _editingComment = false;
    }
  };

  document.addEventListener('mousedown', onDocMousedown, true);
  editSpan.addEventListener('keydown', onKey);
}

function showCommentContextMenu(e, node) {
  _commentContextNode = node;
  const menu = document.getElementById('comment-context-menu');
  _showMenuAt(menu, e.clientX, e.clientY, m => m.classList.add('open'));
}

function hideCommentContextMenu() {
  document.getElementById('comment-context-menu').classList.remove('open');
}

function editCommentFromMenu() {
  const node = _commentContextNode;
  _commentContextNode = null;
  if (node) _editCommentText(node);
}

function deleteCommentFromMenu() {
  if (_commentContextNode) { _commentContextNode.remove(); _commentContextNode = null; syncPreviewToContent(); }
}

function _getCommentBodyText(node) {
  if (!node) return '';
  if (node.classList.contains('md-comment--todo'))
    return node.querySelector('.md-todo-body')?.textContent?.trim() ?? '';
  return node.querySelector('.md-comment-text')?.textContent?.trim() ?? '';
}

function _setCommentBodyText(node, text) {
  if (!node) return;
  if (node.classList.contains('md-comment--todo')) {
    const kind = node.getAttribute('data-todo-kind') || 'TODO';
    const span = node.querySelector('.md-todo-body');
    if (span) span.textContent = text;
    node.setAttribute('data-comment', `${kind}: ${text}`);
  } else {
    const span = node.querySelector('.md-comment-text');
    if (span) span.textContent = ` ${text} `;
    node.setAttribute('data-comment', text);
  }
}

function commentCtxCopy() {
  const text = _getCommentBodyText(_commentContextNode);
  if (text) navigator.clipboard.writeText(text);
}

function commentCtxCut() {
  const node = _commentContextNode;
  _commentContextNode = null;
  const text = _getCommentBodyText(node);
  if (text) navigator.clipboard.writeText(text);
  if (node) { node.remove(); syncPreviewToContent(); }
}

function commentCtxPaste() {
  const node = _commentContextNode;
  _commentContextNode = null;
  hideCommentContextMenu();
  if (!node) return;
  navigator.clipboard.readText().then(text => {
    if (!text) return;
    _setCommentBodyText(node, text);
    syncPreviewToContent();
  }).catch(() => {});
}

// ── Shortcode ─────────────────────────────────────────────────────────────────
let _scContextNode     = null;  // nœud .md-shortcode en cours d'édition
let _scInsertionRange  = null;  // range sauvegardé avant ouverture du dialogue
let _scInsertionAnchor = null;  // e.target au moment du clic droit (fallback fiable)

function _buildShortcodeHtml(inner) {
  const rawText = `{{< ${inner} >}}`;
  const safeRaw = rawText.replace(/"/g, '&quot;');
  return `<div class="md-shortcode" data-sc-raw="${safeRaw}" contenteditable="false">`
    + `<span class="sc-brace">{{&lt;</span> <span class="sc-inner">${inner}</span> <span class="sc-brace">&gt;}}</span>`
    + `</div>`;
}

function _updateSCPreview() {
  const name = document.getElementById('sc-name').value.trim();
  const isInclude = name === 'include';
  document.getElementById('sc-args-row').style.display    = isInclude ? 'none' : '';
  document.getElementById('sc-include-row').style.display = isInclude ? '' : 'none';
  const args = isInclude
    ? document.getElementById('sc-include-path').value.trim()
    : document.getElementById('sc-args').value.trim();
  const inner = args ? `${name} ${args}` : name;
  document.getElementById('sc-preview').textContent = inner ? `{{< ${inner} >}}` : '';
}

async function _scBrowseInclude() {
  if (!window.pywebview) return;
  const s = loadSettings();
  const startDir = s.projectDir || '';
  // Répertoire du fichier actif = base pour le chemin relatif
  const tab = getActiveTab();
  const filePath = tab && tab.path ? tab.path.replace(/\\/g, '/') : '';
  const baseDir  = filePath ? filePath.substring(0, filePath.lastIndexOf('/')) : startDir;
  const result = await window.pywebview.api.pick_include_file(startDir, baseDir);
  if (!result) return;
  document.getElementById('sc-include-path').value = result.rel;
  // Synchronise sc-args pour que confirmSCDialog fonctionne sans modification
  document.getElementById('sc-args').value = result.rel;
  _updateSCPreview();
}

function openSCDialog(node) {
  _scContextNode = node || null;
  // Le range d'insertion est déjà sauvegardé au moment du clic droit (contextmenu),
  // avant tout changement de focus. On l'efface seulement en mode édition (node existant).
  if (node) {
    _scInsertionRange = null;
  }
  let name = '', args = '';
  if (node) {
    const raw = (node.getAttribute('data-sc-raw') || '').replace(/&quot;/g, '"');
    const m = raw.match(/^\{\{<\s*(\S+)(.*?)>\}\}$/);
    if (m) { name = m[1].trim(); args = m[2].trim(); }
    document.getElementById('sc-dialog-title').textContent = 'Modifier le shortcode';
    document.getElementById('sc-confirm-btn').textContent  = 'Mettre à jour';
  } else {
    document.getElementById('sc-dialog-title').textContent = 'Insérer un shortcode';
    document.getElementById('sc-confirm-btn').textContent  = 'Insérer';
  }
  document.getElementById('sc-name').value = name;
  document.getElementById('sc-args').value = args;
  // Si édition d'un include existant, pré-remplir le champ chemin
  document.getElementById('sc-include-path').value = (name === 'include') ? args : '';
  _updateSCPreview();
  document.getElementById('sc-overlay').classList.add('open');
  setTimeout(() => document.getElementById('sc-name').focus(), 50);
}

function closeSCDialog(e) {
  if (e && e.target !== document.getElementById('sc-overlay')) return;
  document.getElementById('sc-overlay').classList.remove('open');
  _scContextNode = null;
}

function confirmSCDialog() {
  const name = document.getElementById('sc-name').value.trim();
  if (!name) return;
  const args  = document.getElementById('sc-args').value.trim();
  const inner = args ? `${name} ${args}` : name;

  document.getElementById('sc-overlay').classList.remove('open');

  // Mode édition : mettre à jour le nœud existant
  if (_scContextNode) {
    const rawText = `{{< ${inner} >}}`;
    _scContextNode.setAttribute('data-sc-raw', rawText.replace(/"/g, '&quot;'));
    _scContextNode.querySelector('.sc-inner').textContent = inner;
    _scContextNode = null;
    syncPreviewToContent();
    return;
  }

  // Mode insertion
  const tab = getActiveTab();
  if (!tab) return;
  const md = `{{< ${inner} >}}`;

  if (state.sourceMode) {
    const editor = document.getElementById('source-editor');
    const pos    = editor.selectionEnd;
    editor.value = editor.value.slice(0, pos) + '\n' + md + '\n' + editor.value.slice(pos);
    editor.selectionStart = editor.selectionEnd = pos + md.length + 2;
    editor.focus();
    onSourceInput();
  } else {
    const sel = window.getSelection();
    const preview = document.getElementById('preview');
    const wrapper = document.createElement('div');
    wrapper.innerHTML = _buildShortcodeHtml(inner);
    const node = wrapper.firstChild;
    _insertAtCursor(node, _scInsertionRange);
    _scInsertionRange = null;
    // Re-render depuis le markdown pour que l'extension marked génère le badge proprement
    syncPreviewToContent();
    const _tab = getActiveTab();
    if (_tab) updatePreview(_tab.content);
  }
}

function showSCContextMenu(e, node) {
  _scContextNode = node;

  // Afficher le bouton "Ouvrir" uniquement pour les shortcodes {{< include ... >}}
  const raw       = (node.getAttribute('data-sc-raw') || '').replace(/&quot;/g, '"');
  const isInclude = /^\{\{<\s*include\s+/i.test(raw);
  document.getElementById('sc-ctx-include-open').style.display = isInclude ? '' : 'none';
  document.getElementById('sc-ctx-include-sep').style.display  = isInclude ? '' : 'none';

  const menu = document.getElementById('sc-context-menu');
  _showMenuAt(menu, e.clientX, e.clientY, m => m.classList.add('open'));
  e.preventDefault();
  e.stopPropagation();
}

async function openIncludeFileFromMenu() {
  if (!_scContextNode) return;
  const raw = (_scContextNode.getAttribute('data-sc-raw') || '').replace(/&quot;/g, '"');
  const m   = raw.match(/\{\{<\s*include\s+([^\s>{}]+)/i);
  if (!m) return;
  const relPath = m[1].trim();

  // Résoudre le chemin absolu à partir du fichier courant
  const tab = getActiveTab();
  if (!tab || !tab.path) { alert('Enregistrez le fichier courant avant d\'ouvrir un include.'); return; }

  const norm   = s => s.replace(/\\/g, '/');
  const base   = norm(tab.path).replace(/\/[^/]+$/, ''); // dossier du fichier courant
  const parts  = base.split('/');
  norm(relPath).split('/').forEach(p => {
    if (p === '..') parts.pop();
    else if (p && p !== '.') parts.push(p);
  });
  const absPath = parts.join('/');

  try {
    const data = await window.pywebview.api.open_file_by_path(absPath);
    if (!data || data.error) { alert(`Impossible d'ouvrir : ${absPath}`); return; }
    const existing = state.tabs.find(t => t.path === data.path);
    if (existing) { switchToTab(existing.id); return; }
    const newTab = createTab(data.name, data.path, data.content);
    switchToTab(newTab.id);
  } catch (err) {
    alert(`Erreur lors de l'ouverture : ${err}`);
  }
}

function hideSCContextMenu() {
  document.getElementById('sc-context-menu').classList.remove('open');
}

function editSCFromMenu() {
  const node = _scContextNode;
  _scContextNode = null;
  if (node) openSCDialog(node);
}

function deleteSCFromMenu() {
  if (_scContextNode) { _scContextNode.remove(); _scContextNode = null; syncPreviewToContent(); }
}

// ── DIV conditionnelle ────────────────────────────────────────────────────────
let _fdivContextNode = null;

function applyFencedDiv() {
  if (state.sourceMode) {
    const ta  = document.getElementById('source-editor');
    const pos = ta.selectionEnd;
    const ins = `\n:::\nContenu de la DIV\n:::\n`;
    ta.value = ta.value.substring(0, pos) + ins + ta.value.substring(pos);
    ta.selectionStart = ta.selectionEnd = pos + ins.length;
    ta.focus();
    onSourceInput();
  } else {
    const wrapper = document.createElement('div');
    wrapper.className = 'fenced-div-wrapper';
    const label = document.createElement('span');
    label.className = 'fenced-div-label';
    label.contentEditable = 'false';
    label.textContent = 'DV';
    const content = document.createElement('div');
    content.className = 'fenced-div-content';
    content.innerHTML = '<p>Contenu de la DIV</p>';
    wrapper.appendChild(label);
    wrapper.appendChild(content);
    _insertAtCursor(wrapper);
    syncPreviewToContent();
  }
}

function insertNestedFDiv() {
  if (!_fdivContextNode) return;
  const container = _fdivContextNode.querySelector('.fenced-div-content');
  if (!container) return;
  const wrapper = document.createElement('div');
  wrapper.className = 'fenced-div-wrapper';
  const label = document.createElement('span');
  label.className = 'fenced-div-label';
  label.contentEditable = 'false';
  label.textContent = 'DV';
  const content = document.createElement('div');
  content.className = 'fenced-div-content';
  content.innerHTML = '<p>Contenu de la DIV</p>';
  wrapper.appendChild(label);
  wrapper.appendChild(content);
  container.appendChild(wrapper);
  _fdivClearDirectContentIfNested(_fdivContextNode);
  syncPreviewToContent();
}

async function openFDivDialog() {
  hideFDivContextMenu();
  if (!_fdivContextNode) return;

  // Détecter le niveau d'imbrication (0 = parente, 1+ = imbriquée)
  let level = 0;
  let ancestor = _fdivContextNode.parentElement;
  while (ancestor) {
    if (ancestor.classList && ancestor.classList.contains('fenced-div-wrapper')) level++;
    ancestor = ancestor.parentElement;
  }
  const badge = document.getElementById('fdiv-level-badge');
  if (badge) badge.textContent = level === 0 ? 'parente' : `niveau\u00a0${level + 1}`;

  document.getElementById('fdiv-class').value  = _fdivContextNode.getAttribute('data-div-class')  || '';
  document.getElementById('fdiv-format').value = _fdivContextNode.getAttribute('data-div-format') || '';
  const w = _fdivContextNode.getAttribute('data-div-width') || '';
  const hasWidth = !!w;
  const wVal = w || 100;
  document.getElementById('fdiv-width-enable').checked  = hasWidth;
  document.getElementById('fdiv-width-slider').value    = wVal;
  document.getElementById('fdiv-width-number').value    = wVal;
  document.getElementById('fdiv-width-pct').textContent = wVal + '%';
  _fdivWidthToggle(hasWidth);
  document.getElementById('fdiv-margin').checked = _fdivContextNode.getAttribute('data-div-margin') === '1';
  _fdivSyncChips();
  await _fdivLoadProjectClasses();

  _fdivBuildNav();
  _fdivBuildPreview(false);   // aperçu replié par défaut

  document.getElementById('fdiv-overlay').classList.add('open');
}

function closeFDivDialog(e) {
  if (e && e.target !== document.getElementById('fdiv-overlay')) return;
  document.getElementById('fdiv-overlay').classList.remove('open');
}

/** Construit la barre de navigation parent ↔ DIV courante ↔ enfants. */
function _fdivBuildNav() {
  const nav = document.getElementById('fdiv-nav');
  if (!nav || !_fdivContextNode) return;
  nav.innerHTML = '';

  // Chaîne des parents (du plus lointain au plus proche)
  const parents = [];
  let el = _fdivContextNode.parentElement;
  while (el) {
    if (el.classList && el.classList.contains('fenced-div-wrapper')) parents.unshift(el);
    el = el.parentElement;
  }

  // Boutons parents (breadcrumb)
  parents.forEach((parentNode, i) => {
    const btn = document.createElement('button');
    btn.className = 'fdiv-nav-btn fdiv-nav-up';
    const cls = parentNode.getAttribute('data-div-class') || '';
    btn.textContent = cls ? `⬆ .${cls.split(' ')[0]}` : '⬆ DIV';
    btn.title = cls || 'DIV parente';
    btn.onclick = () => { _fdivContextNode = parentNode; openFDivDialog(); };
    nav.appendChild(btn);
    const sep = document.createElement('span');
    sep.className = 'fdiv-nav-sep';
    sep.textContent = '›';
    nav.appendChild(sep);
  });

  // DIV courante
  const cur = document.createElement('span');
  cur.className = 'fdiv-nav-current';
  const curCls = _fdivContextNode.getAttribute('data-div-class') || '';
  cur.textContent = curCls ? `.${curCls.split(' ')[0]}` : 'DV';
  cur.title = curCls || 'DV active';
  nav.appendChild(cur);

  // Enfants
  const contentEl = _fdivContextNode.querySelector(':scope > .fenced-div-content');
  if (contentEl) {
    const children = [...contentEl.children].filter(el => el.classList.contains('fenced-div-wrapper'));
    if (children.length > 0) {
      const sep = document.createElement('span');
      sep.className = 'fdiv-nav-sep';
      sep.textContent = '›';
      nav.appendChild(sep);
      const strip = document.createElement('div');
      strip.className = 'fdiv-nav-children';
      children.forEach((child, i) => {
        const btn = document.createElement('button');
        btn.className = 'fdiv-nav-btn fdiv-nav-child';
        const childCls = child.getAttribute('data-div-class') || '';
        btn.textContent = childCls ? `.${childCls.split(' ')[0]}` : `Enfant ${i + 1}`;
        btn.title = childCls || `Sous-DIV ${i + 1}`;
        btn.onclick = () => { _fdivContextNode = child; openFDivDialog(); };
        strip.appendChild(btn);
      });
      nav.appendChild(strip);
    }
  }
}

/** Construit l'aperçu du contenu de la DIV courante (texte direct, sans sous-DIVs). */
function _fdivBuildPreview(open = false) {
  const section = document.getElementById('fdiv-preview-section');
  const content = document.getElementById('fdiv-preview-content');
  const arrow   = document.getElementById('fdiv-preview-arrow');
  if (!section || !content || !_fdivContextNode) return;

  // Cloner le contenu direct (sans les sous-DIVs)
  const contentEl = _fdivContextNode.querySelector(':scope > .fenced-div-content');
  if (contentEl) {
    const clone = contentEl.cloneNode(true);
    clone.querySelectorAll('.fenced-div-wrapper').forEach(el => el.remove());
    clone.querySelectorAll('.fenced-div-label').forEach(el => el.remove());
    const text = clone.textContent.trim();
    if (text) {
      content.innerHTML = clone.innerHTML;
    } else {
      content.innerHTML = '<em class="fdiv-preview-empty">Contenu vide</em>';
    }
  } else {
    content.innerHTML = '<em class="fdiv-preview-empty">—</em>';
  }

  // État ouvert/fermé
  content.style.display = open ? '' : 'none';
  if (arrow) arrow.textContent = open ? '▼' : '▶';
}

/** Bascule la visibilité de l'aperçu du contenu. */
function fdivTogglePreview() {
  const content = document.getElementById('fdiv-preview-content');
  const arrow   = document.getElementById('fdiv-preview-arrow');
  if (!content) return;
  const isOpen = content.style.display !== 'none';
  content.style.display = isOpen ? 'none' : '';
  if (arrow) arrow.textContent = isOpen ? '▶' : '▼';
}

function _updateFDivLabel(wrapper) {
  const label = wrapper.querySelector('.fenced-div-label');
  if (!label) return;
  label.textContent = 'DV';
}

function showFDivContextMenu(e, node) {
  _fdivContextNode = node;
  const menu = document.getElementById('fdiv-context-menu');
  menu.style.left = e.clientX + 'px';
  menu.style.top  = e.clientY + 'px';
  menu.classList.add('open');
  e.preventDefault();
  e.stopPropagation();
}

function hideFDivContextMenu() {
  document.getElementById('fdiv-context-menu').classList.remove('open');
}

// ── Helpers chips classes DIV ─────────────────────────────────────────────────
function _fdivGetClasses() {
  return new Set(
    document.getElementById('fdiv-class').value
      .split(/\s+/).filter(Boolean)
  );
}

const _FDIV_CALLOUT_CLASSES = new Set([
  'callout-warning', 'callout-note', 'callout-tip', 'callout-important', 'callout-caution'
]);

function _fdivToggleClass(cls) {
  const classes = _fdivGetClasses();
  if (classes.has(cls)) {
    classes.delete(cls);
  } else {
    // Les callouts sont mutuellement exclusifs
    if (_FDIV_CALLOUT_CLASSES.has(cls)) {
      _FDIV_CALLOUT_CLASSES.forEach(c => classes.delete(c));
    }
    classes.add(cls);
  }
  document.getElementById('fdiv-class').value = [...classes].join(' ');
  _fdivSyncChips();
}

function _fdivSyncChips() {
  const active = _fdivGetClasses();
  document.querySelectorAll('#fdiv-dialog .fdiv-list-item').forEach(item => {
    item.classList.toggle('active', active.has(item.dataset.cls));
  });
}

function _fdivWidthToggle(enabled) {
  document.getElementById('fdiv-width-slider').disabled = !enabled;
  document.getElementById('fdiv-width-number').disabled = !enabled;
}

async function _fdivLoadProjectClasses() {
  const s = loadSettings();
  const projectDir = s.projectDir || '';
  const container = document.getElementById('fdiv-list-project');
  const pathLabel = document.getElementById('fdiv-project-css-path');
  if (!projectDir || !window.pywebview) {
    container.innerHTML = '<div class="fdiv-list-empty">Aucun fichier<br>conf/styles.css</div>';
    return;
  }
  const cssPath = projectDir.replace(/\\/g, '/').replace(/\/$/, '') + '/conf/styles.css';
  const classes = await window.pywebview.api.get_css_classes(cssPath);
  if (!classes || classes.length === 0) {
    container.innerHTML = '<div class="fdiv-list-empty">Aucun fichier<br>conf/styles.css</div>';
    pathLabel.textContent = '';
    return;
  }
  pathLabel.textContent = 'conf/styles.css';
  const active = _fdivGetClasses();
  container.innerHTML = '';
  classes.forEach(cls => {
    const el = document.createElement('div');
    el.className = 'fdiv-list-item' + (active.has(cls) ? ' active' : '');
    el.dataset.cls = cls;
    el.textContent = cls;
    el.onclick = () => _fdivToggleClass(cls);
    container.appendChild(el);
  });
}

// Si un format est sélectionné, ajoute 'content-visible' aux classes si absent
// (et si 'content-hidden' n'est pas déjà présent)
function _fdivEnsureContentVisible(cls, fmt) {
  if (!fmt) return cls;
  const parts = cls.split(/\s+/).filter(Boolean);
  if (!parts.includes('content-visible') && !parts.includes('content-hidden')) {
    parts.unshift('content-visible');
  }
  return parts.join(' ');
}

// Applique les valeurs du dialogue sur _fdivContextNode (sans fermer ni synchro)
function _applyFDivSettings() {
  if (!_fdivContextNode) return;
  const divFormatVal = document.getElementById('fdiv-format').value;
  const divClassVal  = _fdivEnsureContentVisible(
    document.getElementById('fdiv-class').value, divFormatVal
  );
  document.getElementById('fdiv-class').value = divClassVal;
  divClassVal  ? _fdivContextNode.setAttribute('data-div-class',  divClassVal)  : _fdivContextNode.removeAttribute('data-div-class');
  divFormatVal ? _fdivContextNode.setAttribute('data-div-format', divFormatVal) : _fdivContextNode.removeAttribute('data-div-format');
  _fdivContextNode.className = 'fenced-div-wrapper';
  if (divClassVal) divClassVal.split(/\s+/).filter(Boolean).forEach(c => _fdivContextNode.classList.add(c));

  const widthEnabled = document.getElementById('fdiv-width-enable').checked;
  const widthVal     = document.getElementById('fdiv-width-number').value;
  const marginAuto   = document.getElementById('fdiv-margin').checked;
  if (widthEnabled && widthVal) {
    _fdivContextNode.setAttribute('data-div-width', widthVal);
  } else {
    _fdivContextNode.removeAttribute('data-div-width');
  }
  if (marginAuto) {
    _fdivContextNode.setAttribute('data-div-margin', '1');
  } else {
    _fdivContextNode.removeAttribute('data-div-margin');
  }
  const styleParts = [];
  if (widthEnabled && widthVal) styleParts.push(`width: ${widthVal}%`);
  if (marginAuto) styleParts.push('margin: auto');
  _fdivContextNode.style.cssText = styleParts.join('; ');
  _updateFDivLabel(_fdivContextNode);
}

// Supprime le contenu direct d'une .fenced-div-content si elle contient des sous-divs
function _fdivClearDirectContentIfNested(wrapper) {
  const contentEl = wrapper?.querySelector('.fenced-div-content');
  if (!contentEl) return;
  const hasNested = contentEl.querySelector(':scope > .fenced-div-wrapper');
  if (!hasNested) return;
  [...contentEl.childNodes].forEach(node => {
    const isWrapper = node.nodeType === Node.ELEMENT_NODE
                      && node.classList.contains('fenced-div-wrapper');
    if (!isWrapper) node.remove();
  });
}

function confirmFDivDialog() {
  _applyFDivSettings();
  _fdivClearDirectContentIfNested(_fdivContextNode);
  document.getElementById('fdiv-overlay').classList.remove('open');
  syncPreviewToContent();
}

/** Valide le dialogue courant, insère une sous-DIV et ouvre le dialogue sur elle. */
function fdivDialogAddNested() {
  if (!_fdivContextNode) return;
  _applyFDivSettings();
  const parentNode = _fdivContextNode;
  document.getElementById('fdiv-overlay').classList.remove('open');

  const container = parentNode.querySelector('.fenced-div-content');
  if (!container) { syncPreviewToContent(); return; }

  const wrapper = document.createElement('div');
  wrapper.className = 'fenced-div-wrapper';
  const label = document.createElement('span');
  label.className = 'fenced-div-label';
  label.contentEditable = 'false';
  label.textContent = 'DV';
  const content = document.createElement('div');
  content.className = 'fenced-div-content';
  content.innerHTML = '<p>Contenu de la DIV</p>';
  wrapper.appendChild(label);
  wrapper.appendChild(content);
  container.appendChild(wrapper);

  // La div parente contient maintenant une sous-div → vider son contenu direct
  _fdivClearDirectContentIfNested(parentNode);
  syncPreviewToContent();

  // Ouvrir immédiatement le dialogue complet pour la nouvelle sous-DIV
  _fdivContextNode = wrapper;
  openFDivDialog();
}

/** Déplace la DIV cible vers le haut (-1) ou vers le bas (+1) parmi ses sœurs. */
function moveFDiv(direction) {
  if (!_fdivContextNode) return;
  const parent = _fdivContextNode.parentElement;
  if (!parent) return;
  const siblings = [...parent.children].filter(el => el.classList.contains('fenced-div-wrapper'));
  const idx = siblings.indexOf(_fdivContextNode);
  if (direction < 0) {
    if (idx <= 0) return;
    parent.insertBefore(_fdivContextNode, siblings[idx - 1]);
  } else {
    if (idx >= siblings.length - 1) return;
    parent.insertBefore(siblings[idx + 1], _fdivContextNode);
  }
  syncPreviewToContent();
}

/** Met à jour la visibilité des boutons Monter/Descendre du menu contextuel. */
function _fdivUpdateMoveButtons(node) {
  const parent = node.parentElement;
  const isNested = !!(parent && parent.classList.contains('fenced-div-content')
    && parent.closest('.fenced-div-wrapper'));
  const siblings = isNested
    ? [...parent.children].filter(el => el.classList.contains('fenced-div-wrapper'))
    : [];
  const idx      = siblings.indexOf(node);
  const moveUp   = document.getElementById('fdiv-ctx-move-up');
  const moveDown = document.getElementById('fdiv-ctx-move-down');
  const moveSep  = document.getElementById('fdiv-ctx-move-sep');
  if (moveUp)   moveUp.style.display   = (isNested && idx > 0)                       ? '' : 'none';
  if (moveDown) moveDown.style.display = (isNested && idx < siblings.length - 1)     ? '' : 'none';
  if (moveSep)  moveSep.style.display  = (isNested && siblings.length > 1)            ? '' : 'none';
}

function deleteFDiv() {
  document.getElementById('fdiv-overlay').classList.remove('open');
  hideFDivContextMenu();
  if (!_fdivContextNode) return;
  _fdivContextNode.remove();
  _fdivContextNode = null;
  syncPreviewToContent();
}

function hideFDivContextMenu() {
  document.getElementById('fdiv-context-menu').classList.remove('open');
}

function setupFDivContextMenu() {
  const preview = document.getElementById('preview');

  // Double-clic sur le label → ouvre directement le dialogue
  preview.addEventListener('dblclick', e => {
    const label = e.target.closest('.fenced-div-label');
    if (!label) return;
    const wrapper = label.closest('.fenced-div-wrapper');
    if (!wrapper) return;
    _fdivContextNode = wrapper;
    openFDivDialog();
    e.preventDefault();
    e.stopPropagation();
  });

  // Clic droit sur le label → menu contextuel dédié
  preview.addEventListener('contextmenu', e => {
    const label = e.target.closest('.fenced-div-label');
    if (!label) return;
    const wrapper = label.closest('.fenced-div-wrapper');
    if (!wrapper) return;
    _fdivContextNode = wrapper;
    _fdivUpdateMoveButtons(wrapper);
    const menu = document.getElementById('fdiv-context-menu');
    _showMenuAt(menu, e.clientX, e.clientY, m => m.classList.add('open'));
    e.preventDefault();
    e.stopPropagation();
  });

  document.addEventListener('click', () => hideFDivContextMenu());

  // Survol du label → encart markdown
  let _fdivTooltipTimer = null;
  preview.addEventListener('mouseover', e => {
    const label = e.target.closest('.fenced-div-label');
    if (!label) return;
    const wrapper = label.closest('.fenced-div-wrapper');
    if (!wrapper) return;
    clearTimeout(_fdivTooltipTimer);
    _fdivTooltipTimer = setTimeout(() => {
      const md = turndown.turndown(wrapper.outerHTML).trim();
      _showFDivTooltip(label, md);
    }, 300);
  });
  preview.addEventListener('mouseout', e => {
    const label = e.target.closest('.fenced-div-label');
    if (!label) return;
    clearTimeout(_fdivTooltipTimer);
    // Laisser un court délai pour permettre de survoler le tooltip lui-même
    _fdivTooltipTimer = setTimeout(_hideFDivTooltip, 200);
  });
}

// ── Menu contextuel des blocs de code ────────────────────────────────────────
let _codeBlockContextNode = null;

function setupCodeBlockContextMenu() {
  const preview = document.getElementById('preview');

  preview.addEventListener('contextmenu', e => {
    const wrapper = e.target.closest('.code-block-wrapper');
    if (!wrapper) return;
    // Ne pas interférer si on est déjà dans un menu spécifique plus prioritaire
    if (e.target.closest('.fenced-div-label')) return;
    e.preventDefault();
    e.stopPropagation();
    _codeBlockContextNode = wrapper;
    showCodeBlockContextMenu(e, wrapper);
  });

  document.addEventListener('click', () => hideCodeBlockContextMenu());
}

function showCodeBlockContextMenu(e, wrapper) {
  const menu       = document.getElementById('code-block-context-menu');
  const wrapBtn    = document.getElementById('code-ctx-wrap-btn');
  const unwrapBtn  = document.getElementById('code-ctx-unwrap-btn');
  const insideFDiv = !!wrapper.closest('.fenced-div-wrapper');

  wrapBtn.style.display   = insideFDiv ? 'none' : '';
  unwrapBtn.style.display = insideFDiv ? ''      : 'none';

  menu.style.left = e.clientX + 'px';
  menu.style.top  = e.clientY + 'px';
  menu.classList.add('open');
}

function hideCodeBlockContextMenu() {
  document.getElementById('code-block-context-menu').classList.remove('open');
}

function wrapCodeBlockInDiv() {
  if (!_codeBlockContextNode) return;

  const fdiv    = document.createElement('div');
  fdiv.className = 'fenced-div-wrapper';

  const label   = document.createElement('span');
  label.className      = 'fenced-div-label';
  label.contentEditable = 'false';
  label.textContent    = 'DV';

  const content = document.createElement('div');
  content.className = 'fenced-div-content';

  // Déplacer le bloc de code dans le contenu de la DIV
  _codeBlockContextNode.parentNode.insertBefore(fdiv, _codeBlockContextNode);
  content.appendChild(_codeBlockContextNode);
  fdiv.appendChild(label);
  fdiv.appendChild(content);

  _codeBlockContextNode = null;
  syncPreviewToContent();
}

function unwrapCodeBlockFromDiv() {
  if (!_codeBlockContextNode) return;

  const fdivContent = _codeBlockContextNode.closest('.fenced-div-content');
  if (!fdivContent) { _codeBlockContextNode = null; return; }
  const fdivWrapper = fdivContent.closest('.fenced-div-wrapper');
  if (!fdivWrapper) { _codeBlockContextNode = null; return; }

  // Déplacer le bloc de code juste avant la DIV englobante
  fdivWrapper.parentNode.insertBefore(_codeBlockContextNode, fdivWrapper);

  // Supprimer la DIV si elle est maintenant vide (ou ne contient plus rien d'utile)
  const remaining = fdivContent.innerHTML.trim().replace(/<br\s*\/?>/gi, '').trim();
  if (!remaining) fdivWrapper.remove();

  _codeBlockContextNode = null;
  syncPreviewToContent();
}

function deleteCodeBlock() {
  if (!_codeBlockContextNode) return;
  _codeBlockContextNode.remove();
  _codeBlockContextNode = null;
  syncPreviewToContent();
}

function _showFDivTooltip(labelEl, mdText) {
  let tooltip = document.getElementById('fdiv-md-tooltip');
  if (!tooltip) return;

  const pre = tooltip.querySelector('.fdiv-tooltip-code');
  pre.textContent = mdText;

  // Réinitialiser la largeur pour mesurer librement
  tooltip.style.width = '';
  tooltip.classList.add('visible');

  // Largeur naturelle = largeur du <pre> + paddings du tooltip (bordures)
  const maxVw   = Math.floor(window.innerWidth * 0.90);
  const needed  = pre.scrollWidth + 24 + 2; // padding 12px × 2 + bordure × 2
  const tw = Math.min(Math.max(needed, 260), maxVw);
  tooltip.style.width = tw + 'px';

  // Positionner sous (ou au-dessus si pas de place) du label
  const rect = labelEl.getBoundingClientRect();
  const th = tooltip.offsetHeight || 120;
  let left = rect.right - tw;
  if (left < 8) left = 8;
  if (left + tw > window.innerWidth - 8) left = window.innerWidth - tw - 8;
  let top = rect.bottom + 6;
  if (top + th > window.innerHeight - 8) top = rect.top - th - 6;
  tooltip.style.left = left + 'px';
  tooltip.style.top  = top  + 'px';

  // Annuler la fermeture si on entre dans le tooltip
  tooltip.onmouseenter = () => clearTimeout(document.getElementById('fdiv-md-tooltip')._hideTimer);
  tooltip.onmouseleave = () => { document.getElementById('fdiv-md-tooltip')._hideTimer = setTimeout(_hideFDivTooltip, 150); };
}

function _hideFDivTooltip() {
  const tooltip = document.getElementById('fdiv-md-tooltip');
  if (tooltip) tooltip.classList.remove('visible');
}

// ── Titres ────────────────────────────────────────────────────────────────────
function toggleHeadingMenu() {
  document.getElementById('heading-dropdown').classList.toggle('open');
}

function closeHeadingMenu() {
  document.getElementById('heading-dropdown')?.classList.remove('open');
}

function applyHeading(level) {
  closeHeadingMenu();

  if (state.sourceMode) {
    const ta    = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const val   = ta.value;

    const lineStart   = val.lastIndexOf('\n', start - 1) + 1;
    const lineEnd     = val.indexOf('\n', end);
    const lineEndSafe = lineEnd === -1 ? val.length : lineEnd;
    const line        = val.substring(lineStart, lineEndSafe);
    const stripped    = line.replace(/^#{1,6}\s/, '');
    const newLine     = level === 0 ? stripped : '#'.repeat(level) + ' ' + stripped;

    ta.value = val.substring(0, lineStart) + newLine + val.substring(lineEndSafe);
    ta.selectionStart = lineStart;
    ta.selectionEnd   = lineStart + newLine.length;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    const range = sel.getRangeAt(0);

    let node = range.commonAncestorContainer;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
    const block = node.closest('h1,h2,h3,h4,h5,h6,p,div,li') || node;
    if (!preview.contains(block)) return;

    const tag    = block.tagName.toLowerCase();
    const newTag = level === 0 ? 'p' : (tag === `h${level}` ? 'p' : `h${level}`);
    const newEl  = document.createElement(newTag);
    newEl.innerHTML = block.innerHTML;
    block.replaceWith(newEl);

    // Replacer le curseur dans le nouveau nœud
    const newRange = document.createRange();
    newRange.selectNodeContents(newEl);
    newRange.collapse(false);
    sel.removeAllRanges();
    sel.addRange(newRange);

    syncPreviewToContent();
    updateTOC();
  }
}

function insertEmptyParaAbove() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const pos = ta.selectionStart;
    const lineStart = ta.value.lastIndexOf('\n', pos - 1) + 1;
    ta.value = ta.value.substring(0, lineStart) + '\n' + ta.value.substring(lineStart);
    ta.selectionStart = ta.selectionEnd = lineStart;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    let node = sel.getRangeAt(0).commonAncestorContainer;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
    // Remonter jusqu'au bloc direct enfant de #preview
    while (node && node.parentElement !== preview) node = node.parentElement;
    if (!node || !preview.contains(node)) return;

    const newP = document.createElement('p');
    newP.innerHTML = '<br>';
    preview.insertBefore(newP, node);

    // Placer le curseur dans le nouveau paragraphe
    const newRange = document.createRange();
    newRange.setStart(newP, 0);
    newRange.collapse(true);
    sel.removeAllRanges();
    sel.addRange(newRange);
    newP.focus();

    syncPreviewToContent();
  }
}

// ── Formatage ─────────────────────────────────────────────────────────────────
function applyBold() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const selected = ta.value.substring(start, end);
    ta.value = ta.value.substring(0, start) + '**' + selected + '**' + ta.value.substring(end);
    ta.selectionStart = start + 2;
    ta.selectionEnd   = end + 2;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (!preview.contains(sel.anchorNode)) return;
    document.execCommand('bold');
    syncPreviewToContent();
  }
}

// Supprime tous les spans d'une classe qui se trouvent dans la sélection
// ou qui englobent le curseur (sélection collapsed)
function _removeColorSpan(className) {
  const preview = document.getElementById('preview');
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) return false;
  if (!preview.contains(sel.anchorNode)) return false;
  const range = sel.getRangeAt(0);

  // Collecte les spans qui chevauchent la sélection
  const toRemove = [];
  preview.querySelectorAll('.' + className).forEach(span => {
    if (range.intersectsNode(span)) toRemove.push(span);
  });

  // Si rien en sélection, cherche l'ancêtre englobant le curseur
  if (toRemove.length === 0) {
    let node = range.startContainer;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
    const span = node.closest('.' + className);
    if (span) toRemove.push(span);
  }

  if (toRemove.length === 0) return false;

  toRemove.forEach(span => {
    const parent = span.parentNode;
    while (span.firstChild) parent.insertBefore(span.firstChild, span);
    parent.removeChild(span);
  });
  preview.normalize();
  syncPreviewToContent();
  return true;
}

// Supprime la syntaxe couleur en mode source autour du curseur ou dans la sélection
function _removeColorSource(pattern) {
  const ta     = document.getElementById('source-editor');
  const cursor = ta.selectionStart;
  const selEnd = ta.selectionEnd;
  const text   = ta.value;

  // Cas 1 : du texte est sélectionné → essaie de retirer le wrapper dans la sélection
  if (cursor !== selEnd) {
    const selected  = text.substring(cursor, selEnd);
    const stripped  = selected.replace(new RegExp(pattern.source, 'g'), '$1');
    if (stripped !== selected) {
      ta.value = text.substring(0, cursor) + stripped + text.substring(selEnd);
      ta.selectionStart = cursor;
      ta.selectionEnd   = cursor + stripped.length;
      ta.focus();
      onSourceInput();
      return true;
    }
  }

  // Cas 2 : cherche le pattern qui contient la position du curseur
  const searchStart = Math.max(0, cursor - 300);
  const re = new RegExp(pattern.source, 'g');
  re.lastIndex = 0;
  const segment = text.substring(searchStart);
  let match;
  while ((match = re.exec(segment)) !== null) {
    const absStart = searchStart + match.index;
    const absEnd   = absStart + match[0].length;
    if (absStart <= cursor && cursor <= absEnd) {
      const inner = match[1];
      ta.value = text.substring(0, absStart) + inner + text.substring(absEnd);
      ta.selectionStart = absStart;
      ta.selectionEnd   = absStart + inner.length;
      ta.focus();
      onSourceInput();
      return true;
    }
  }
  return false;
}

function applyHighlight(color) {
  if (color === null) {
    if (state.sourceMode) {
      _removeColorSource(/\[((?:[^\[\]]|\[[^\]]*\])*)\]\{style="background-color:[^"]+"\}/g);
    } else {
      _removeColorSpan('md-highlight');
    }
    return;
  }
  if (state.sourceMode) {
    const ta    = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const sel   = ta.value.substring(start, end);
    if (!sel) return;
    const wrapped = `[${sel}]{style="background-color: ${color};"}`;
    ta.value = ta.value.substring(0, start) + wrapped + ta.value.substring(end);
    ta.selectionStart = start;
    ta.selectionEnd   = start + wrapped.length;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return;
    if (!preview.contains(sel.anchorNode)) return;
    const range = sel.getRangeAt(0);
    const span  = document.createElement('span');
    span.className = 'md-highlight';
    span.style.backgroundColor = _HIGHLIGHT_COLORS[color] || color;
    span.setAttribute('data-color', color);
    span.appendChild(range.extractContents());
    range.insertNode(span);
    sel.removeAllRanges();
    syncPreviewToContent();
  }
}

function applyTextColor(color) {
  if (color === null) {
    if (state.sourceMode) {
      _removeColorSource(/\[((?:[^\[\]]|\[[^\]]*\])*)\]\{style="color:[^"]+"\}/g);
    } else {
      _removeColorSpan('md-textcolor');
    }
    return;
  }
  if (state.sourceMode) {
    const ta    = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const sel   = ta.value.substring(start, end);
    if (!sel) return;
    const wrapped = `[${sel}]{style="color: ${color};"}`;
    ta.value = ta.value.substring(0, start) + wrapped + ta.value.substring(end);
    ta.selectionStart = start;
    ta.selectionEnd   = start + wrapped.length;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return;
    if (!preview.contains(sel.anchorNode)) return;
    const range = sel.getRangeAt(0);
    const span  = document.createElement('span');
    span.className = 'md-textcolor';
    span.style.color = color;
    span.setAttribute('data-color', color);
    span.appendChild(range.extractContents());
    range.insertNode(span);
    sel.removeAllRanges();
    syncPreviewToContent();
  }
}

// ── Ajout d'un commentaire (annotation) sur la sélection courante ──────────────
function addAnnotation() {
  if (state.sourceMode) {
    const ta    = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const sel   = ta.value.substring(start, end);
    if (!sel) return;
    const id      = _newAnnotId();
    const wrapped = `[${sel}]{.comment comment-id="${id}" comment-text=""}`;
    ta.value = ta.value.substring(0, start) + wrapped + ta.value.substring(end);
    ta.selectionStart = start;
    ta.selectionEnd   = start + wrapped.length;
    ta.focus();
    onSourceInput();
    updateAnnotationsList();
    _focusAnnotationInput(id);
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return;
    if (!preview.contains(sel.anchorNode)) return;
    const range = sel.getRangeAt(0);
    const span  = document.createElement('span');
    span.className = 'md-annotation';
    const id = _newAnnotId();
    span.setAttribute('data-annot-id', id);
    span.setAttribute('data-annot-text', '');
    span.appendChild(range.extractContents());
    range.insertNode(span);
    sel.removeAllRanges();
    syncPreviewToContent();
    updateAnnotationsList();
    _focusAnnotationInput(id);
  }
}

function applyCodeBlock() {
  const DEFAULT_LANG = 'texinfo';
  if (state.sourceMode) {
    const ta     = document.getElementById('source-editor');
    const start  = ta.selectionStart;
    const end    = ta.selectionEnd;
    const selected = ta.value.substring(start, end);
    const prefix = '```' + DEFAULT_LANG + '\n';
    const block  = prefix + selected + '\n```';
    ta.value = ta.value.substring(0, start) + block + ta.value.substring(end);
    ta.selectionStart = start + prefix.length;
    ta.selectionEnd   = start + prefix.length + selected.length;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (!preview.contains(sel.anchorNode)) return;

    const range   = sel.getRangeAt(0);
    const wrapper = document.createElement('div');
    wrapper.className = 'code-block-wrapper';

    const label = document.createElement('span');
    label.className = 'code-lang-label';
    label.textContent = DEFAULT_LANG;
    label.title = 'Cliquer pour changer le langage';

    const pre  = document.createElement('pre');
    const code = document.createElement('code');
    code.className = 'language-' + DEFAULT_LANG;
    code.appendChild(range.extractContents());
    pre.appendChild(code);
    wrapper.appendChild(label);
    wrapper.appendChild(pre);
    range.insertNode(wrapper);

    label.addEventListener('click', e => { e.stopPropagation(); openLangEditor(label, code); });

    syncPreviewToContent();
    const tab = getActiveTab();
    if (tab) updatePreview(tab.content);
  }
}

function applyYamlBlock() {
  const DEFAULT_CONTENT = 'Title: Sans titre';
  if (state.sourceMode) {
    const ta      = document.getElementById('source-editor');
    const start   = ta.selectionStart;
    const end     = ta.selectionEnd;
    const selected = ta.value.substring(start, end) || DEFAULT_CONTENT;
    const block   = '---\n' + selected + '\n---';
    ta.value = ta.value.substring(0, start) + block + ta.value.substring(end);
    ta.selectionStart = start + 4;
    ta.selectionEnd   = start + 4 + selected.length;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel     = window.getSelection();
    const selectedText = sel && sel.rangeCount > 0 ? sel.toString() : '';
    const content = selectedText.trim() || DEFAULT_CONTENT;

    if (sel && sel.rangeCount > 0 && preview.contains(sel.anchorNode)) {
      sel.getRangeAt(0).deleteContents();
    }

    const wrapper = document.createElement('div');
    wrapper.className = 'yaml-block';
    const header = document.createElement('div');
    header.className = 'yaml-block-header';
    header.textContent = '― YAML ―';
    const pre  = document.createElement('pre');
    const code = document.createElement('code');
    code.textContent = content;
    pre.appendChild(code);
    wrapper.appendChild(header);
    wrapper.appendChild(pre);

    if (sel && sel.rangeCount > 0 && preview.contains(sel.anchorNode)) {
      sel.getRangeAt(0).insertNode(wrapper);
    } else {
      preview.appendChild(wrapper);
    }
    sel && sel.removeAllRanges();
    syncPreviewToContent();
    const tab = getActiveTab();
    if (tab) updatePreview(tab.content);
  }
}

function applyTable() {
  const DEFAULT_TABLE =
    '+---------+---------+---------+\n' +
    '| Col 1   | Col 2   | Col 3   |\n' +
    '+=========+=========+=========+\n' +
    '| cell    | cell    | cell    |\n' +
    '+---------+---------+---------+\n' +
    '| cell    | cell    | cell    |\n' +
    '+---------+---------+---------+\n';

  if (state.sourceMode) {
    const ta    = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const block = '\n' + DEFAULT_TABLE + '\n';
    ta.value = ta.value.substring(0, start) + block + ta.value.substring(start);
    ta.selectionStart = ta.selectionEnd = start + block.length;
    ta.focus();
    onSourceInput();
  } else {
    const activeTab = getActiveTab();
    if (!activeTab) return;

    // Construire l'élément table à partir du HTML parsé
    const temp = document.createElement('div');
    temp.innerHTML = marked.parse(DEFAULT_TABLE.trim());
    const tableEl = temp.querySelector('table');
    if (!tableEl) return;

    const preview = document.getElementById('preview');
    const sel = window.getSelection();

    if (sel && sel.rangeCount > 0 && preview.contains(sel.anchorNode)) {
      const range = sel.getRangeAt(0);
      range.collapse(true);
      range.insertNode(tableEl);
      range.setStartAfter(tableEl);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    } else {
      preview.appendChild(tableEl);
    }

    // Synchroniser le DOM → markdown puis re-rendre
    syncPreviewToContent();
    updatePreview(activeTab.content);
  }
}

function applyBlockquote() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const before    = ta.value.substring(0, start);
    const lineStart = before.lastIndexOf('\n') + 1;
    const selected  = ta.value.substring(lineStart, end);
    const lines = selected.split('\n');
    const allQuoted = lines.every(l => /^> /.test(l) || l === '>');
    const replaced = allQuoted
      ? lines.map(l => l.replace(/^> ?/, '')).join('\n')
      : lines.map(l => l === '' ? '>' : '> ' + l).join('\n');
    ta.value = ta.value.substring(0, lineStart) + replaced + ta.value.substring(end);
    if (allQuoted) {
      // Suppression : sélectionner le bloc résultant
      ta.selectionStart = lineStart;
      ta.selectionEnd   = lineStart + replaced.length;
    } else {
      // Ajout : placer le curseur au début du contenu de la première ligne citée
      const firstLineOffset = lines[0] === '' ? 1 : 2; // après '>' ou '> '
      ta.selectionStart = ta.selectionEnd = lineStart + firstLineOffset;
    }
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (!preview.contains(sel.anchorNode)) return;
    const range = sel.getRangeAt(0);

    // Toggle : si le curseur est déjà dans un blockquote, on le retire
    const existingBq = sel.anchorNode.parentElement?.closest?.('blockquote');
    if (existingBq && preview.contains(existingBq)) {
      // Déplacer les enfants du blockquote devant lui, puis le supprimer
      const parent = existingBq.parentNode;
      while (existingBq.firstChild) {
        parent.insertBefore(existingBq.firstChild, existingBq);
      }
      parent.removeChild(existingBq);
      sel.removeAllRanges();
    } else {
      // Envelopper la sélection dans un nouveau blockquote
      const bq = document.createElement('blockquote');
      bq.appendChild(range.extractContents());
      range.insertNode(bq);
      // Placer le curseur au début du contenu du blockquote
      const newRange = document.createRange();
      // Trouver le premier nœud de texte ou enfant dans le blockquote
      let target = bq.firstChild;
      while (target && target.nodeType === Node.ELEMENT_NODE && target.firstChild) {
        target = target.firstChild;
      }
      if (target) {
        newRange.setStart(target, 0);
        newRange.collapse(true);
      } else {
        newRange.setStart(bq, 0);
        newRange.collapse(true);
      }
      sel.removeAllRanges();
      sel.addRange(newRange);
    }
    syncPreviewToContent();
  }
}

// ── Menu contextuel note de bas de page ───────────────────────────────────────
let _footnoteContextSup = null;

function showFootnoteContextMenu(e, supEl) {
  _footnoteContextSup = supEl;
  const menu = document.getElementById('footnote-context-menu');
  _showMenuAt(menu, e.clientX, e.clientY, m => m.classList.add('open'));
  e.preventDefault();
  e.stopPropagation();
}

function hideFootnoteContextMenu() {
  document.getElementById('footnote-context-menu').classList.remove('open');
  _footnoteContextSup = null;
}

function editFootnoteFromMenu() {
  const sup = _footnoteContextSup;
  hideFootnoteContextMenu();
  if (!sup) return;
  const label = sup.getAttribute('data-label');
  if (!label) return;

  let existing = '';
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const match = ta.value.match(new RegExp(`\\[\\^${label}\\]:\\s+(.+)`));
    existing = match ? match[1] : '';
  } else {
    const defEl  = document.querySelector(`#preview .footnote-def[data-label="${label}"]`);
    const textEl = defEl ? defEl.querySelector('.footnote-def-text') : null;
    existing = textEl ? textEl.textContent.trim() : '';
  }
  _footnoteRange = null;
  _footnoteEditLabel = label;
  _fnOpenDialog(existing, true);
}

function deleteFootnoteFromMenu() {
  const sup = _footnoteContextSup;
  hideFootnoteContextMenu();
  if (!sup) return;
  const label = sup.getAttribute('data-label');
  if (!label) return;

  if (state.sourceMode) {
    const ta  = document.getElementById('source-editor');
    let src   = ta.value;
    // Supprimer la référence [^label]
    src = src.replace(new RegExp(`\\[\\^${label}\\]`, 'g'), '');
    // Supprimer la définition [^label]: texte (ligne entière)
    src = src.replace(new RegExp(`\\n?\\[\\^${label}\\]:[^\\n]*`, 'g'), '');
    ta.value = src.trimEnd();
    onSourceInput();
  } else {
    // Supprimer la référence (le <sup>)
    sup.remove();
    // Supprimer la définition correspondante
    const defEl = document.querySelector(`#preview .footnote-def[data-label="${label}"]`);
    if (defEl) defEl.remove();
    syncPreviewToContent();
    updateFileStatus(getActiveTab());
  }
}

// ── Notes de bas de page ──────────────────────────────────────────────────────
let _footnoteRange = null;
let _footnoteEditLabel = null; // null = nouvelle note, sinon label à éditer

// ── Mini éditeur footnote ─────────────────────────────────────────────────────
function _fnUpdatePreview() {
  const text = document.getElementById('footnote-text-input').value;
  if (!text.trim()) { document.getElementById('footnote-preview').innerHTML = ''; return; }
  // marked.parse pour supporter les listes (blocs), parseInline sinon
  const hasList = /^[-*+]\s|^\d+\.\s/m.test(text);
  document.getElementById('footnote-preview').innerHTML = hasList
    ? marked.parse(text)
    : marked.parseInline(text);
}

function _fnFormat(type) {
  const ta = document.getElementById('footnote-text-input');
  const start = ta.selectionStart;
  const end   = ta.selectionEnd;
  const sel   = ta.value.substring(start, end);

  // ── Listes : transformation ligne par ligne ───────────────────────────────
  if (type === 'ul' || type === 'ol') {
    // Si pas de sélection, on insère un item vide sur une nouvelle ligne
    if (start === end) {
      const before = ta.value.substring(0, start);
      const after  = ta.value.substring(end);
      const prefix = type === 'ul' ? '- ' : '1. ';
      // Ajouter une nouvelle ligne si on n'est pas déjà en début de ligne
      const needNl = before.length > 0 && !before.endsWith('\n');
      const insertion = (needNl ? '\n' : '') + prefix;
      ta.setRangeText(insertion, start, end, 'end');
      ta.focus();
      _fnUpdatePreview();
      return;
    }
    // Transformer chaque ligne sélectionnée
    const lines = sel.split('\n');
    let counter = 1;
    // Compter les items existants avant la sélection pour numérotation correcte
    if (type === 'ol') {
      const textBefore = ta.value.substring(0, start);
      const lastItems = textBefore.match(/^\d+\.\s/gm);
      if (lastItems) counter = lastItems.length + 1;
    }
    const transformed = lines.map((line, i) => {
      if (!line.trim()) return line; // garder les lignes vides telles quelles
      const already = type === 'ul' ? /^[-*+]\s/.test(line) : /^\d+\.\s/.test(line);
      if (already) return line; // déjà formaté
      return type === 'ul' ? `- ${line}` : `${counter++ + i - 1}. ${line}`;
    }).join('\n');
    ta.setRangeText(transformed, start, end, 'select');
    ta.focus();
    _fnUpdatePreview();
    return;
  }

  // ── Formatage inline (bold, italic, code, link) ───────────────────────────
  let before = '', after = '', placeholder = '';
  switch (type) {
    case 'bold':   before = '**'; after = '**';      placeholder = 'texte en gras';    break;
    case 'italic': before = '*';  after = '*';       placeholder = 'texte en italique'; break;
    case 'code':   before = '`';  after = '`';       placeholder = 'code';              break;
    case 'link':   before = '[';  after = '](url)';  placeholder = 'texte du lien';     break;
  }
  const inner = sel || placeholder;
  const replacement = before + inner + after;
  ta.setRangeText(replacement, start, end, 'select');
  ta.selectionStart = start + before.length;
  ta.selectionEnd   = start + before.length + inner.length;
  ta.focus();
  _fnUpdatePreview();
}

// ── Règle de tabulation ───────────────────────────────────────────────────────
let _tabStops      = [];   // positions en colonnes, ex. [4, 8, 12, 16, 24, 32]
let _rulerCharW    = 0;    // largeur d'un caractère (px), mesuré dynamiquement
let _rulerLeftOff  = 0;    // décalage gauche = numéros de ligne + padding éditeur

function _defaultTabStops(interval) {
  const stops = [];
  for (let c = interval; c <= 80; c += interval) stops.push(c);
  return stops;
}

function _loadTabStops() {
  const s = loadSettings();
  _tabStops = Array.isArray(s.tabStops) && s.tabStops.length
    ? s.tabStops
    : _defaultTabStops(s.tabSize || 4);
}

function _measureRulerCharW() {
  let ta;
  if (state.sourceMode) {
    ta = document.getElementById('source-editor');
  } else {
    const sel = window.getSelection();
    ta = document.getElementById('preview');
    if (sel && sel.rangeCount > 0 && ta && ta.contains(sel.anchorNode)) {
      let node = sel.anchorNode;
      if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
      const code = node.closest?.('code, pre');
      if (code) ta = code;
    }
  }
  const span = document.createElement('span');
  const cs   = getComputedStyle(ta);
  span.style.cssText = `position:absolute;visibility:hidden;white-space:pre;` +
    `font-family:${cs.fontFamily};font-size:${cs.fontSize};`;
  span.textContent = 'X'.repeat(20);
  document.body.appendChild(span);
  const w = span.getBoundingClientRect().width / 20;
  span.remove();
  return w || 8.4;
}

function _getRulerLeftOff() {
  if (state.sourceMode) {
    const ln = document.getElementById('source-line-numbers');
    return (ln ? ln.offsetWidth : 0) + 16; // 16 = padding-left du source-editor
  } else {
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      let node = sel.anchorNode;
      if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
      const pre = node.closest?.('pre');
      const ruler = document.getElementById('tab-ruler');
      if (pre && ruler) {
        const preRect = pre.getBoundingClientRect();
        const rulerRect = ruler.getBoundingClientRect();
        const style = window.getComputedStyle(pre);
        const pl = parseFloat(style.paddingLeft) || 0;
        const bl = parseFloat(style.borderLeftWidth) || 0;
        return preRect.left - rulerRect.left + pl + bl;
      }
    }
    return 0;
  }
}

function _renderRuler() {
  const canvas = document.getElementById('tab-ruler-canvas');
  const ruler  = document.getElementById('tab-ruler');
  if (!canvas || !ruler) return;

  const W = ruler.offsetWidth;
  const H = ruler.offsetHeight;

  // Dimensions nulles = layout pas encore calculé → réessayer
  if (W === 0 || H === 0) { setTimeout(_renderRuler, 50); return; }

  const dpr = window.devicePixelRatio || 1;
  canvas.width  = Math.round(W * dpr);
  canvas.height = Math.round(H * dpr);
  canvas.style.width  = W + 'px';
  canvas.style.height = H + 'px';

  const ctx = canvas.getContext('2d');
  ctx.scale(dpr, dpr);

  // Couleurs : lecture directe sur un élément rendu pour éviter les variables non résolues
  const srcEl  = document.getElementById('source-editor');
  const bodyCs = getComputedStyle(document.body);
  const bg     = getComputedStyle(document.getElementById('source-line-numbers') || document.body).backgroundColor || '#1e2030';
  const muted  = '#6c7086';
  const border = '#45475a';
  const accent = 'rgba(245,169,73,0.95)';

  // Fond
  ctx.fillStyle = bg;
  ctx.fillRect(0, 0, W, H);

  // Ligne de séparation basse
  ctx.strokeStyle = border;
  ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(0, H - 0.5); ctx.lineTo(W, H - 0.5); ctx.stroke();

  if (!_rulerCharW) _rulerCharW = _measureRulerCharW();
  if (!_rulerLeftOff) _rulerLeftOff = _getRulerLeftOff();
  const charW   = _rulerCharW;
  const leftOff = _rulerLeftOff;
  const maxCols = Math.ceil((W - leftOff) / charW) + 2;

  // Intervalle entre les arrêts (ex. 4 si tabSize=4)
  const interval = _tabStops.length >= 2
    ? _tabStops[1] - _tabStops[0]
    : (loadSettings().tabSize || 4);

  // Ticks légers à chaque tabSize (repère visuel de la grille)
  ctx.strokeStyle = muted;
  ctx.lineWidth   = 0.5;
  for (let col = interval; col <= maxCols; col += interval) {
    const x = leftOff + col * charW;
    if (x < 0 || x > W) continue;
    ctx.beginPath();
    ctx.moveTo(x, H - 5); ctx.lineTo(x, H - 1);
    ctx.stroke();
  }

  // Numéros de colonne tous les 2 intervalles
  ctx.fillStyle    = muted;
  ctx.font         = `9px 'JetBrains Mono', monospace`;
  ctx.textAlign    = 'left';
  ctx.textBaseline = 'top';
  const labelStep  = interval * 2;
  for (let col = labelStep; col <= maxCols; col += labelStep) {
    const x = leftOff + col * charW;
    if (x < 0 || x > W) continue;
    ctx.fillText(String(col), x + 2, 1);
  }

  // Marqueurs d'arrêt (L inversé, comme Word) à chaque position de tabulation
  for (const stop of _tabStops) {
    const x = Math.round(leftOff + stop * charW);
    if (x < 0 || x > W) continue;
    ctx.strokeStyle = accent;
    ctx.fillStyle   = accent;
    ctx.lineWidth   = 1.5;
    ctx.beginPath();
    ctx.moveTo(x,     H - 14);
    ctx.lineTo(x,     H - 2);
    ctx.lineTo(x + 7, H - 2);
    ctx.stroke();
  }
}

function _rulerClick(e) {
  if (!_rulerCharW) _rulerCharW = _measureRulerCharW();
  if (!_rulerLeftOff) _rulerLeftOff = _getRulerLeftOff();
  const ruler = document.getElementById('tab-ruler');
  const x   = e.clientX - ruler.getBoundingClientRect().left;
  const col = Math.round((x - _rulerLeftOff) / _rulerCharW);
  if (col < 1) return;

  const near = _tabStops.findIndex(s => Math.abs(s - col) <= 1);
  if (near >= 0) {
    _tabStops.splice(near, 1);
  } else {
    _tabStops.push(col);
    _tabStops.sort((a, b) => a - b);
  }
  const s = loadSettings();
  saveSettings({ ...s, tabStops: [..._tabStops] });
  _renderRuler();
}

function _rulerHover(e) {
  if (!_rulerCharW) _rulerCharW = _measureRulerCharW();
  if (!_rulerLeftOff) _rulerLeftOff = _getRulerLeftOff();
  const ruler = document.getElementById('tab-ruler');
  const x   = e.clientX - ruler.getBoundingClientRect().left;
  const col = Math.round((x - _rulerLeftOff) / _rulerCharW);
  if (col < 1) { ruler.title = ''; return; }
  const exists = _tabStops.some(s => Math.abs(s - col) <= 1);
  ruler.title = `Colonne ${col} — Clic pour ${exists ? 'supprimer' : 'ajouter'} un point d'arrêt`;
}

function resetTabStops() {
  const s = loadSettings();
  _tabStops = _defaultTabStops(s.tabSize || 4);
  saveSettings({ ...s, tabStops: [..._tabStops] });
  _rulerCharW   = 0;
  _rulerLeftOff = 0;
  _renderRuler();
}

// Affiche la règle uniquement quand le curseur est dans un bloc de code fencé ou <pre>
function _updateRulerVisibility() {
  const ruler = document.getElementById('tab-ruler');
  if (!ruler) return;
  
  let shouldShow = false;

  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    if (ta && _taCursorInCodeFence(ta)) {
      shouldShow = true;
    }
  } else {
    // Mode prévisualisation : cursor dans un bloc pre/code
    const sel = window.getSelection();
    const preview = document.getElementById('preview');
    if (sel && sel.rangeCount > 0 && preview && preview.contains(sel.anchorNode)) {
      let node = sel.anchorNode;
      if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
      if (node && node.closest?.('pre, code')) {
        shouldShow = true;
      }
    }
  }

  const isVisible = ruler.classList.contains('ruler-visible');
  if (shouldShow && !isVisible) {
    ruler.classList.add('ruler-visible');
    _rulerCharW = 0; _rulerLeftOff = 0;
    setTimeout(_renderRuler, 0);   // setTimeout : layout mis à jour avant le dessin
  } else if (!shouldShow && isVisible) {
    ruler.classList.remove('ruler-visible');
  } else if (shouldShow && isVisible && !state.sourceMode) {
    // L'offset gauche peut changer selon l'imbrication (ex: listes, citations)
    const newLeft = _getRulerLeftOff();
    if (_rulerLeftOff !== newLeft && newLeft > 0) {
      _rulerLeftOff = newLeft;
      _rulerCharW = 0;
      _renderRuler();
    }
  }
}

function initTabRuler() {
  _loadTabStops();
  const ruler = document.getElementById('tab-ruler');
  if (!ruler) return;
  ruler.addEventListener('click', _rulerClick);
  ruler.addEventListener('mousemove', _rulerHover);

  // ResizeObserver : re-render quand les dimensions changent (ex. redimensionnement de fenêtre)
  if (typeof ResizeObserver !== 'undefined') {
    new ResizeObserver(() => {
      if (document.getElementById('tab-ruler')?.classList.contains('ruler-visible')) {
        _rulerCharW = 0; _rulerLeftOff = 0; _renderRuler();
      }
    }).observe(ruler);
  }
}

// Calcule le nombre d'espaces à insérer (Tab) ou supprimer (Shift+Tab)
// selon la colonne courante et les points d'arrêt configurés.
function _tabStopSpaces(col, reverse) {
  if (reverse) {
    // Revenir au stop précédent
    const prev = [..._tabStops].reverse().find(s => s < col) ?? 0;
    return col - prev;
  }
  const next = _tabStops.find(s => s > col);
  if (next != null) return next - col;
  // Après le dernier stop : revenir cycliquement au premier intervalle
  const interval = _tabStops.length >= 2
    ? _tabStops[1] - _tabStops[0]
    : (loadSettings().tabSize || 4);
  const last = _tabStops.at(-1) ?? 0;
  return (last + interval) - col;
}

// Applique Tab / Shift+Tab dans une textarea à la position curseur
function _taTabStop(ta, reverse) {
  const pos       = ta.selectionStart;
  const lineStart = ta.value.lastIndexOf('\n', pos - 1) + 1;
  const col       = pos - lineStart;
  if (reverse) {
    const spaces = _tabStopSpaces(col, true);
    const before = ta.value.substring(pos - spaces, pos);
    if (spaces > 0 && /^ +$/.test(before)) {
      ta.setRangeText('', pos - spaces, pos, 'end');
    }
  } else {
    const spaces = _tabStopSpaces(col, false);
    ta.setRangeText(' '.repeat(Math.max(1, spaces)), pos, ta.selectionEnd, 'end');
  }
}

// Applique Tab / Shift+Tab dans une textarea de code (PUML, YAML, CSS, Lua, JSON) :
// insère/supprime une tabulation au curseur, ou (dé)indente chaque ligne sélectionnée.
function _taCodeTab(ta, e) {
  const val   = ta.value;
  const start = ta.selectionStart;
  const end   = ta.selectionEnd;

  if (start === end) {
    if (e.shiftKey) {
      if (start > 0 && val[start - 1] === '\t') {
        ta.setRangeText('', start - 1, start, 'end');
      }
    } else {
      ta.setRangeText('\t', start, end, 'end');
    }
    return;
  }

  // Sélection multi-lignes : (dé)indenter chaque ligne du bloc sélectionné
  const lineStart  = val.lastIndexOf('\n', start - 1) + 1;
  const lineEndIdx = val.indexOf('\n', end);
  const blockEnd   = lineEndIdx === -1 ? val.length : lineEndIdx;
  const lines      = val.substring(lineStart, blockEnd).split('\n');

  let firstLineDelta = 0, totalDelta = 0;
  const newLines = lines.map((line, i) => {
    if (e.shiftKey) {
      if (line.startsWith('\t')) {
        if (i === 0) firstLineDelta = -1;
        totalDelta -= 1;
        return line.slice(1);
      }
      return line;
    }
    if (i === 0) firstLineDelta = 1;
    totalDelta += 1;
    return '\t' + line;
  });

  ta.setRangeText(newLines.join('\n'), lineStart, blockEnd, 'preserve');
  ta.selectionStart = Math.max(lineStart, start + firstLineDelta);
  ta.selectionEnd   = end + totalDelta;
}

// ── Détection curseur dans un bloc de code fencé (``` ou ~~~) ────────────────
function _taCursorInCodeFence(ta) {
  const before = ta.value.substring(0, ta.selectionStart);
  const fences  = before.match(/^(`{3,}|~{3,})[^\n]*/mg) || [];
  // Un nombre impair de délimiteurs d'ouverture signifie qu'on est à l'intérieur
  return fences.length % 2 === 1;
}

// ── Indentation de liste dans une textarea (Tab / Shift+Tab) ─────────────────
// Retourne true si Tab a été consommé (ligne courante est un item de liste).
function _taListIndent(ta, e) {
  if (e.key !== 'Tab') return false;
  const val       = ta.value;
  const cursor    = ta.selectionStart;
  // Début de la ligne courante
  const lineStart = val.lastIndexOf('\n', cursor - 1) + 1;
  const lineEnd   = val.indexOf('\n', cursor);
  const line      = val.substring(lineStart, lineEnd === -1 ? val.length : lineEnd);
  // La ligne doit être un item de liste (avec indentation éventuelle)
  if (!/^(\s*)([-*+]|\d+\.)\s/.test(line)) return false;
  e.preventDefault();
  const indent = e.shiftKey ? -2 : 2;
  if (indent > 0) {
    // Ajouter 2 espaces en début de ligne
    ta.setRangeText('  ', lineStart, lineStart, 'preserve');
    ta.selectionStart = ta.selectionEnd = cursor + 2;
  } else {
    // Retirer jusqu'à 2 espaces en début de ligne
    const leading = line.match(/^( +)/);
    const remove  = leading ? Math.min(leading[1].length, 2) : 0;
    if (remove > 0) {
      ta.setRangeText('', lineStart, lineStart + remove, 'preserve');
      ta.selectionStart = ta.selectionEnd = Math.max(lineStart, cursor - remove);
    }
  }
  return true;
}

function _fnKeydown(e) {
  if ((e.ctrlKey || e.metaKey) && e.key === 'b') { e.preventDefault(); _fnFormat('bold'); return; }
  if ((e.ctrlKey || e.metaKey) && e.key === 'i') { e.preventDefault(); _fnFormat('italic'); return; }
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); confirmFootnoteDialog(); return; }
  if (e.key === 'Escape') { e.preventDefault(); cancelFootnoteDialog(); return; }
  // Tab dans une liste
  const ta = document.getElementById('footnote-text-input');
  if (_taListIndent(ta, e)) _fnUpdatePreview();
}

function _fnOpenDialog(text, isEdit) {
  document.getElementById('footnote-text-input').value = text || '';
  document.getElementById('footnote-confirm-btn').textContent = isEdit ? 'Mettre à jour' : 'Insérer';
  document.getElementById('footnote-dialog-title').textContent = isEdit ? 'Modifier la note' : 'Note de bas de page';
  _fnUpdatePreview();
  document.getElementById('footnote-overlay').style.display = 'flex';
  setTimeout(() => {
    const ta = document.getElementById('footnote-text-input');
    ta.focus();
    ta.selectionStart = ta.selectionEnd = ta.value.length;
  }, 50);
  // Mise à jour de l'aperçu en live
  document.getElementById('footnote-text-input').oninput = _fnUpdatePreview;
}

function applyFootnote() {
  const preview = document.getElementById('preview');
  if (!state.sourceMode) {
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0 && preview.contains(sel.anchorNode)) {
      _footnoteRange = sel.getRangeAt(0).cloneRange();
    } else {
      _footnoteRange = null;
    }
  }
  _footnoteEditLabel = null;
  _fnOpenDialog('', false);
}

function cancelFootnoteDialog() {
  document.getElementById('footnote-overlay').style.display = 'none';
  _footnoteRange = null;
  _footnoteEditLabel = null;
}

function confirmFootnoteDialog() {
  const text = document.getElementById('footnote-text-input').value.trim();
  if (!text) { cancelFootnoteDialog(); return; }

  // Mode édition : mettre à jour une note existante
  if (_footnoteEditLabel !== null) {
    const label = _footnoteEditLabel;
    if (state.sourceMode) {
      const ta  = document.getElementById('source-editor');
      ta.value  = ta.value.replace(
        new RegExp(`(\\[\\^${label}\\]:\\s+).+`),
        `$1${text}`
      );
      onSourceInput();
    } else {
      const defEl  = document.querySelector(`#preview .footnote-def[data-label="${label}"]`);
      const textEl = defEl ? defEl.querySelector('.footnote-def-text') : null;
      if (textEl) textEl.textContent = text;
      syncPreviewToContent();
      updateFileStatus(getActiveTab());
    }
    cancelFootnoteDialog();
    return;
  }

  const label = _nextFootnoteLabel();

  if (state.sourceMode) {
    const ta     = document.getElementById('source-editor');
    const pos    = ta.selectionEnd;
    const before = ta.value.substring(0, pos);
    const after  = ta.value.substring(pos);
    const ref    = `[^${label}]`;
    const def    = `\n\n[^${label}]: ${text}`;
    ta.value = before + ref + after + def;
    ta.selectionStart = ta.selectionEnd = pos + ref.length;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const range = _footnoteRange || (() => {
      const r = document.createRange();
      r.selectNodeContents(preview);
      r.collapse(false);
      return r;
    })();
    range.collapse(false);

    // Insérer la référence au curseur
    const sup = document.createElement('sup');
    sup.className = 'footnote-ref';
    sup.setAttribute('data-label', label);
    sup.innerHTML = `<a href="#fn-${label}">[${label}]</a>`;
    range.insertNode(sup);
    range.setStartAfter(sup);
    range.collapse(true);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);

    // Insérer la définition en fin de preview
    const defDiv = document.createElement('div');
    defDiv.className = 'footnote-def';
    defDiv.id = `fn-${label}`;
    defDiv.setAttribute('data-label', label);
    defDiv.innerHTML = `<sup class="footnote-def-num">${label}</sup> <span class="footnote-def-text">${text}</span>`;
    preview.appendChild(defDiv);

    syncPreviewToContent();
    updateFileStatus(getActiveTab());
  }

  cancelFootnoteDialog();
}

function _nextFootnoteLabel() {
  if (state.sourceMode) {
    const src = document.getElementById('source-editor').value;
    const nums = [...src.matchAll(/\[\^(\d+)\]/g)].map(m => parseInt(m[1]));
    return nums.length > 0 ? Math.max(...nums) + 1 : 1;
  } else {
    return document.querySelectorAll('#preview .footnote-ref').length + 1;
  }
}

function applyPagebreak() {
  if (state.sourceMode) {
    const ta     = document.getElementById('source-editor');
    const start  = ta.selectionStart;
    const before = ta.value.substring(0, start);
    const after  = ta.value.substring(ta.selectionEnd);
    const nl     = before.length > 0 && !before.endsWith('\n') ? '\n' : '';
    const ins    = nl + '\n{{< pagebreak >}}\n\n';
    ta.value = before + ins + after;
    ta.selectionStart = ta.selectionEnd = before.length + ins.length;
    ta.focus();
    onSourceInput();
  } else {
    const sc = document.createElement('div');
    sc.className = 'md-shortcode';
    sc.setAttribute('data-sc-raw', '{{< pagebreak >}}');
    sc.contentEditable = 'false';
    sc.innerHTML = '<span class="sc-brace">{{&lt;</span> <span class="sc-inner">pagebreak</span> <span class="sc-brace">&gt;}}</span>';
    _insertAtCursor(sc);
    syncPreviewToContent();
    const tab = getActiveTab();
    if (tab) updatePreview(tab.content);
  }
}

function applyHorizontalRule() {
  if (state.sourceMode) {
    const ta    = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const before = ta.value.substring(0, start);
    const after  = ta.value.substring(start);
    const nl = before.length > 0 && !before.endsWith('\n') ? '\n' : '';
    ta.value = before + nl + '\n---\n\n' + after;
    const pos = before.length + nl.length + 5;
    ta.selectionStart = ta.selectionEnd = pos;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || !preview.contains(sel.anchorNode)) return;
    const range = sel.getRangeAt(0);
    const hr = document.createElement('hr');
    range.collapse(false);
    range.insertNode(hr);
    range.setStartAfter(hr);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    syncPreviewToContent();
    const tab = getActiveTab();
    if (tab) updatePreview(tab.content);
  }
}

function applyOrderedList() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const before    = ta.value.substring(0, start);
    const lineStart = before.lastIndexOf('\n') + 1;
    const selected  = ta.value.substring(lineStart, end);
    const lines = selected.split('\n');
    const allNumbered = lines.every(l => /^\d+\. /.test(l));
    const replaced = allNumbered
      ? lines.map(l => l.replace(/^\d+\. /, '')).join('\n')
      : lines.map((l, i) => `${i + 1}. ${l}`).join('\n');
    ta.value = ta.value.substring(0, lineStart) + replaced + ta.value.substring(end);
    ta.selectionStart = lineStart;
    ta.selectionEnd   = lineStart + replaced.length;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (!preview.contains(sel.anchorNode)) return;
    document.execCommand('insertOrderedList');
    syncPreviewToContent();
  }
}

function applyBulletList() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const before  = ta.value.substring(0, start);
    const lineStart = before.lastIndexOf('\n') + 1;
    const selected  = ta.value.substring(lineStart, end);
    const lines = selected.split('\n');
    const allBulleted = lines.every(l => /^- /.test(l));
    const replaced = lines.map(l => allBulleted ? l.replace(/^- /, '') : '- ' + l).join('\n');
    ta.value = ta.value.substring(0, lineStart) + replaced + ta.value.substring(end);
    ta.selectionStart = lineStart;
    ta.selectionEnd   = lineStart + replaced.length;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (!preview.contains(sel.anchorNode)) return;
    document.execCommand('insertUnorderedList');
    syncPreviewToContent();
  }
}

function applyCheckboxList() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const before    = ta.value.substring(0, start);
    const lineStart = before.lastIndexOf('\n') + 1;
    const selected  = ta.value.substring(lineStart, end);
    const lines = selected.split('\n');
    const allCheck = lines.every(l => /^- \[[ xX]\] /.test(l));
    const replaced = allCheck
      ? lines.map(l => l.replace(/^- \[[ xX]\] /, '')).join('\n')
      : lines.map(l => '- [ ] ' + l.replace(/^(?:- \[[ xX]\] |- |\* |\+ |\d+\. )/, '')).join('\n');
    ta.value = ta.value.substring(0, lineStart) + replaced + ta.value.substring(end);
    ta.selectionStart = lineStart;
    ta.selectionEnd   = lineStart + replaced.length;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (!preview.contains(sel.anchorNode)) return;
    let node = sel.anchorNode;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
    const li = node.closest('li');
    if (li) {
      const cb = li.querySelector('input.task-checkbox');
      if (cb) {
        cb.remove();
        const labelDiv = li.querySelector('div.task-label');
        if (labelDiv) {
          while (labelDiv.firstChild) li.appendChild(labelDiv.firstChild);
          labelDiv.remove();
        }
        li.classList.remove('task-item');
        const ul = li.closest('ul');
        if (ul && !ul.querySelector('li.task-item')) ul.classList.remove('task-list');
      } else {
        const input = document.createElement('input');
        input.type = 'checkbox';
        input.className = 'task-checkbox';
        const labelDiv = document.createElement('div');
        labelDiv.className = 'task-label';
        while (li.firstChild) labelDiv.appendChild(li.firstChild);
        li.appendChild(input);
        li.appendChild(labelDiv);
        li.classList.add('task-item');
        const ul = li.closest('ul');
        if (ul) ul.classList.add('task-list');
      }
    } else {
      document.execCommand('insertUnorderedList');
      const newSel = window.getSelection();
      if (newSel && newSel.rangeCount > 0) {
        let newNode = newSel.anchorNode;
        if (newNode.nodeType === Node.TEXT_NODE) newNode = newNode.parentElement;
        const newLi = newNode.closest('li');
        if (newLi) {
          const input = document.createElement('input');
          input.type = 'checkbox';
          input.className = 'task-checkbox';
          const labelDiv = document.createElement('div');
          labelDiv.className = 'task-label';
          while (newLi.firstChild) labelDiv.appendChild(newLi.firstChild);
          newLi.appendChild(input);
          newLi.appendChild(labelDiv);
          newLi.classList.add('task-item');
          const ul = newLi.closest('ul');
          if (ul) ul.classList.add('task-list');
        }
      }
    }
    syncPreviewToContent();
  }
}

function applySmallCaps() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const selected = ta.value.substring(start, end);
    const replacement = `[${selected}]{.smallcaps}`;
    ta.value = ta.value.substring(0, start) + replacement + ta.value.substring(end);
    ta.selectionStart = start + 1;
    ta.selectionEnd   = start + 1 + selected.length;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (!preview.contains(sel.anchorNode)) return;
    const range = sel.getRangeAt(0);
    const span = document.createElement('span');
    span.className = 'smallcaps';
    span.appendChild(range.extractContents());
    range.insertNode(span);
    syncPreviewToContent();
  }
}

function applyStrikethrough() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const selected = ta.value.substring(start, end);
    ta.value = ta.value.substring(0, start) + '~~' + selected + '~~' + ta.value.substring(end);
    ta.selectionStart = start + 2;
    ta.selectionEnd   = end + 2;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (!preview.contains(sel.anchorNode)) return;
    document.execCommand('strikethrough');
    syncPreviewToContent();
  }
}

function applyCode() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const selected = ta.value.substring(start, end);
    ta.value = ta.value.substring(0, start) + '`' + selected + '`' + ta.value.substring(end);
    ta.selectionStart = start + 1;
    ta.selectionEnd   = end + 1;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (!preview.contains(sel.anchorNode)) return;
    const range = sel.getRangeAt(0);
    const code = document.createElement('code');
    code.appendChild(range.extractContents());
    range.insertNode(code);
    syncPreviewToContent();
  }
}

function clearFormatting() {
  if (state.sourceMode) {
    const ta    = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    if (start === end) return;
    let sel = ta.value.substring(start, end);
    // Supprimer les marqueurs inline : **gras**, *italique*, __souligné__, ~~barré~~,
    // `code`, [texte]{.smallcaps}, [texte]{.underline}, [texte]{style="..."}, ==texte==
    sel = sel
      .replace(/\*\*(.+?)\*\*/gs, '$1')
      .replace(/\*(.+?)\*/gs, '$1')
      .replace(/__(.+?)__/gs, '$1')
      .replace(/~~(.+?)~~/gs, '$1')
      .replace(/`(.+?)`/gs, '$1')
      .replace(/\[(.+?)\]\{[^}]*\}/gs, '$1')
      .replace(/==(.+?)==/gs, '$1');
    // Supprimer le préfixe de titre sur chaque ligne
    sel = sel.replace(/^#{1,6}\s+/gm, '');
    ta.value = ta.value.substring(0, start) + sel + ta.value.substring(end);
    ta.selectionStart = start;
    ta.selectionEnd   = start + sel.length;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel     = window.getSelection();
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return;
    if (!preview.contains(sel.anchorNode)) return;

    // Récupérer l'ancêtre commun pour limiter le nettoyage post-formatage
    const range    = sel.getRangeAt(0);
    let   ancestor = range.commonAncestorContainer;
    if (ancestor.nodeType === Node.TEXT_NODE) ancestor = ancestor.parentElement;

    // removeFormat retire le formatage inline (gras, italique, soulignement,
    // couleurs, background-color) tout en conservant le texte
    document.execCommand('removeFormat');

    // Déballer les éléments inline résiduels que removeFormat ne traite pas
    // (code, mark, span de classe/style) dans la zone affectée
    ancestor.querySelectorAll('code, mark, span[class], span[style]').forEach(el => {
      if (!el.parentNode) return;
      const parent = el.parentNode;
      while (el.firstChild) parent.insertBefore(el.firstChild, el);
      parent.removeChild(el);
    });
    ancestor.normalize();

    syncPreviewToContent();
    updateTOC();
  }
}

function applyUnderline() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const selected = ta.value.substring(start, end);
    const replacement = `[${selected}]{.underline}`;
    ta.value = ta.value.substring(0, start) + replacement + ta.value.substring(end);
    ta.selectionStart = start + 1;
    ta.selectionEnd   = start + 1 + selected.length;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (!preview.contains(sel.anchorNode)) return;
    document.execCommand('underline');
    syncPreviewToContent();
  }
}

function applyItalic() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const start = ta.selectionStart;
    const end   = ta.selectionEnd;
    const selected = ta.value.substring(start, end);
    ta.value = ta.value.substring(0, start) + '*' + selected + '*' + ta.value.substring(end);
    ta.selectionStart = start + 1;
    ta.selectionEnd   = end + 1;
    ta.focus();
    onSourceInput();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    if (!preview.contains(sel.anchorNode)) return;
    document.execCommand('italic');
    syncPreviewToContent();
  }
}

function setZoom(value) {
  const pct = parseInt(value);
  document.getElementById('zoom-label').textContent = pct + '%';
  const zoom = pct / 100;
  document.getElementById('preview').style.zoom = zoom;
  const htmlViewer = document.getElementById('html-viewer');
  if (htmlViewer) htmlViewer.style.zoom = zoom;
  const pdfViewer = document.getElementById('pdf-viewer');
  if (pdfViewer) pdfViewer.style.zoom = zoom;
  // Source : zoom via font-size pour éviter tout décalage curseur/texte
  const fs = (14 * zoom).toFixed(2) + 'px';
  document.getElementById('source-editor').style.fontSize   = fs;
  const yamlFs = (13 * zoom).toFixed(2) + 'px';
  document.getElementById('yaml-editor').style.fontSize    = yamlFs;
  document.getElementById('yaml-highlight').style.fontSize = yamlFs;
  const cssFs = (13 * zoom).toFixed(2) + 'px';
  const cEditor = document.getElementById('css-editor');
  if (cEditor) cEditor.style.fontSize = cssFs;
  const cHighlight = document.getElementById('css-highlight');
  if (cHighlight) cHighlight.style.fontSize = cssFs;
  const lFs = (13 * zoom).toFixed(2) + 'px';
  const lEditor = document.getElementById('lua-editor');
  if (lEditor) lEditor.style.fontSize = lFs;
  const lHighlight = document.getElementById('lua-highlight');
  if (lHighlight) lHighlight.style.fontSize = lFs;
  document.getElementById('log-pane').style.fontSize = (13 * zoom).toFixed(2) + 'px';
  const jsonFs = (13 * zoom).toFixed(2) + 'px';
  const jEditor = document.getElementById('json-editor');
  if (jEditor) jEditor.style.fontSize = jsonFs;
  const jHighlight = document.getElementById('json-highlight');
  if (jHighlight) jHighlight.style.fontSize = jsonFs;
  const xmlFs = (13 * zoom).toFixed(2) + 'px';
  const xEditor = document.getElementById('xml-editor');
  if (xEditor) xEditor.style.fontSize = xmlFs;
  const xHighlight = document.getElementById('xml-highlight');
  if (xHighlight) xHighlight.style.fontSize = xmlFs;
  const tab = getActiveTab();
  if (tab) tab.zoom = pct;
  if (state.sourceMode) _updateSourceHighlight();
  if (document.body.classList.contains('yaml-mode')) _updateYamlHighlight();
  if (document.body.classList.contains('css-mode')) _updateCssHighlight();
  if (document.body.classList.contains('lua-mode')) _updateLuaHighlight();
  if (document.body.classList.contains('json-mode')) _updateJsonHighlight();
  if (document.body.classList.contains('xml-mode')) _updateXmlHighlight();
}

// ── Largeur de la colonne de contenu ──────────────────────────────────────────
let _contentMaxWidth = 780;

function _isPreviewWidthActive() {
  return !state.sourceMode &&
    !document.body.classList.contains('yaml-mode') &&
    !document.body.classList.contains('json-mode') &&
    !document.body.classList.contains('xml-mode')  &&
    !document.body.classList.contains('css-mode')  &&
    !document.body.classList.contains('lua-mode')  &&
    !document.body.classList.contains('puml-mode') &&
    !document.body.classList.contains('log-mode')  &&
    !document.body.classList.contains('html-mode') &&
    !document.body.classList.contains('pdf-mode')  &&
    !document.body.classList.contains('image-mode');
}

function _applyContentMaxWidth(w) {
  const pane = document.getElementById('preview-pane');
  const min  = 300;
  const max  = pane ? Math.max(min + 1, pane.clientWidth - 40) : 1600;
  _contentMaxWidth = Math.max(min, Math.min(max, Math.round(w / 10) * 10));
  document.documentElement.style.setProperty('--content-max-width', _contentMaxWidth + 'px');
  const lbl = document.getElementById('status-width-label');
  if (lbl) lbl.textContent = _contentMaxWidth + 'px';
  const slider = document.getElementById('status-width-slider');
  if (slider) {
    slider.max = String(Math.round(max));
    slider.value = String(_contentMaxWidth);
  }
  _updatePreviewWidthHandlePos();
  const s = loadSettings();
  saveSettings({ ...s, contentMaxWidth: _contentMaxWidth });
}

function _updatePreviewWidthHandlePos() {
  const handle  = document.getElementById('preview-width-handle');
  const control = document.getElementById('status-width-control');
  const active  = _isPreviewWidthActive();

  // Contrôle dans la statusbar : visible dès qu'on est en mode prévisualisation
  if (control) control.style.display = active ? 'flex' : 'none';

  // Poignée latérale : visible seulement quand la colonne n'est pas pleine largeur
  if (!handle) return;
  if (!active) { handle.style.display = 'none'; return; }
  const preview = document.getElementById('preview');
  const pane    = document.getElementById('preview-pane');
  if (!preview || !pane || pane.offsetParent === null) { handle.style.display = 'none'; return; }
  const paneRect    = pane.getBoundingClientRect();
  const previewRect = preview.getBoundingClientRect();
  if (previewRect.left <= paneRect.left + 12) { handle.style.display = 'none'; return; }
  handle.style.display  = 'flex';
  handle.style.top      = paneRect.top    + 'px';
  handle.style.height   = paneRect.height + 'px';
  handle.style.left     = (previewRect.left - 16) + 'px';
  handle.style.width    = '18px';
}

function _setupPreviewWidthHandle() {
  const handle = document.getElementById('preview-width-handle');
  if (!handle) return;

  let dragging = false, startX = 0, startW = 0;

  handle.addEventListener('mousedown', e => {
    dragging = true;
    startX   = e.clientX;
    startW   = _contentMaxWidth;
    handle.classList.add('dragging');
    document.body.style.cursor     = 'col-resize';
    document.body.style.userSelect = 'none';
    e.preventDefault();
  });
  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    // Déplacer à gauche = élargir (symétrique via margin:auto)
    const dx = startX - e.clientX;
    _applyContentMaxWidth(startW + dx * 2);
  });
  document.addEventListener('mouseup', () => {
    if (!dragging) return;
    dragging = false;
    handle.classList.remove('dragging');
    document.body.style.cursor     = '';
    document.body.style.userSelect = '';
  });

  handle.addEventListener('dblclick', () => _applyContentMaxWidth(780));

  // Repositionnement dynamique
  const ro = new ResizeObserver(_updatePreviewWidthHandlePos);
  const pane    = document.getElementById('preview-pane');
  const preview = document.getElementById('preview');
  if (pane)    ro.observe(pane);
  if (preview) ro.observe(preview);
  window.addEventListener('resize', _updatePreviewWidthHandlePos);
  setTimeout(_updatePreviewWidthHandlePos, 200);
}

// ── Thème (sombre / clair / système) ─────────────────────────────────────────
const _THEMES = ['dark', 'light', 'system'];
const _HEADING_COLOR_DEFAULTS = { dark: '#89b4fa', light: '#1e66f5', system: '#89b4fa' };
const _THEME_LABELS = { dark: 'Sombre', light: 'Clair', system: 'Système' };
const _THEME_ICONS = {
  dark: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
    <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>
  </svg>`,
  light: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
    <circle cx="12" cy="12" r="5"/>
    <line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/>
    <line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/>
    <line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/>
    <line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>
  </svg>`,
  system: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
    <rect x="2" y="3" width="20" height="14" rx="2"/>
    <line x1="8" y1="21" x2="16" y2="21"/>
    <line x1="12" y1="17" x2="12" y2="21"/>
  </svg>`,
};

function _applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme === 'dark' ? '' : theme);
  if (theme === 'dark') {
    document.documentElement.removeAttribute('data-theme');
  } else {
    document.documentElement.setAttribute('data-theme', theme);
  }
  const btn = document.getElementById('theme-toggle-btn');
  if (btn) {
    btn.innerHTML = _THEME_ICONS[theme] + '<span>' + _THEME_LABELS[theme] + '</span>';
    btn.title = 'Thème : ' + _THEME_LABELS[theme] + ' — cliquer pour changer';
  }
  // Mettre à jour la couleur des titres pour le nouveau thème
  _applyFonts({ ...loadSettings(), theme });
}

function cycleTheme() {
  const s = loadSettings();
  const cur = s.theme || 'dark';
  const next = _THEMES[(_THEMES.indexOf(cur) + 1) % _THEMES.length];
  saveSettings({ ...s, theme: next });
  _applyTheme(next);
}

function _initTheme() {
  const theme = loadSettings().theme || 'dark';
  _applyTheme(theme);
}

async function openReadme() {
  closeMenu();
  const overlay = document.getElementById('readme-overlay');
  const content = document.getElementById('readme-content');
  overlay.classList.add('open');
  if (content.dataset.loaded) return;
  content.innerHTML = '<p style="color:var(--text-muted);padding:8px 0">Chargement…</p>';
  if (!window.pywebview) { content.innerHTML = '<p style="color:var(--text-muted)">API non disponible.</p>'; return; }
  const res = await window.pywebview.api.read_readme();
  if (res.error) {
    content.innerHTML = `<p style="color:#f38ba8">Erreur : ${res.error}</p>`;
  } else {
    content.innerHTML = marked.parse(res.content);
    content.dataset.loaded = '1';
  }
}

function closeReadme() {
  document.getElementById('readme-overlay').classList.remove('open');
}

function closeReadmeOnOverlay(e) {
  if (e.target === document.getElementById('readme-overlay')) closeReadme();
}

async function minimizeWindow() {
  if (window.pywebview) await window.pywebview.api.minimize();
}

async function toggleFullscreen() {
  if (window.pywebview) await window.pywebview.api.toggle_fullscreen();
}

async function quitApp() {
  closeMenu();
  if (window.pywebview) await window.pywebview.api.quit();
}

// ── Toast ─────────────────────────────────────────────────────────────────────
let toastTimer = null;
function showToast(message) {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('show'), 2500);
}

// ── Sidebar ───────────────────────────────────────────────────────────────────
let _sidebarCurrentDir = null;

async function initSidebar() {
  if (!window.pywebview || !window.pywebview.api) return;
  try {
    const s = loadSettings();
    const dir = s.projectDir || await window.pywebview.api.get_default_dir();
    await loadSidebarDir(dir);
  } catch (e) { console.error('initSidebar:', e); }
}

// Navigue le sidebar vers le répertoire d'un fichier donné
function sidebarNavigateTo(filePath) {
  if (!filePath) return;
  const normalized = filePath.replace(/\\/g, '/');
  const dir = normalized.substring(0, normalized.lastIndexOf('/'));
  if (dir) loadSidebarDir(dir);
}

async function loadSidebarDir(path) {
  _sidebarCurrentDir = path;
  try {
    const result = await window.pywebview.api.list_directory(path);
    if (result && !result.error) renderSidebarFiles(result);
  } catch (e) { console.error('loadSidebarDir:', e); }
}

function renderSidebarFiles(result) {
  const dirNameEl = document.getElementById('sidebar-dir-name');
  const parts = result.path.replace(/\\/g, '/').split('/').filter(Boolean);
  const lastName = parts[parts.length - 1] || result.path;
  dirNameEl.textContent = lastName;
  dirNameEl.title = result.path;

  const container = document.getElementById('sidebar-files');
  container.innerHTML = '';

  // Entrée "parent" (..)
  if (result.parent && result.parent !== result.path) {
    const up = document.createElement('div');
    up.className = 'sidebar-item sidebar-item-up';
    up.innerHTML = `
      <svg class="item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <polyline points="15 18 9 12 15 6"/>
        <line x1="9" y1="12" x2="20" y2="12"/>
      </svg>
      <span class="item-name">..</span>`;
    up.onclick = () => loadSidebarDir(result.parent);
    container.appendChild(up);
  }

  for (const item of result.items) {
    const div = document.createElement('div');
    const editable = !item.is_dir && (item.ext === '.md' || item.ext === '.qmd' || item.ext === '.html' || item.ext === '.htm' || item.ext === '.pdf' || item.ext === '.yml' || item.ext === '.yaml' || item.ext === '.log' || item.ext === '.tex' || item.ext === '.puml' || item.ext === '.plantuml' || item.ext === '.css' || item.ext === '.scss' || item.ext === '.less' || item.ext === '.json' || item.ext === '.xml' || _IMAGE_EXTS.has(item.ext));
    let cls = 'sidebar-item';
    if (item.is_dir)       cls += ' sidebar-item-dir';
    else if (!editable)    cls += ' sidebar-item-other';
    div.className = cls;
    if (!item.is_dir) div.dataset.path = item.path;

    const iconSvg = item.is_dir
      ? `<svg class="item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
           <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
         </svg>`
      : `<svg class="item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
           <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
           <polyline points="14 2 14 8 20 8"/>
         </svg>`;

    const dateStr = (!item.is_dir && item.mtime)
      ? new Date(item.mtime * 1000).toLocaleString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }).replace(',', '')
      : '';
    const datePart = dateStr ? `<span class="item-mtime">${dateStr}</span>` : '';
    div.innerHTML = `${iconSvg}<span class="item-name" title="${item.name}">${item.name}</span>${datePart}`;

    if (item.is_dir) {
      div.onclick = () => loadSidebarDir(item.path);
      div.addEventListener('contextmenu', e => {
        e.preventDefault();
        e.stopPropagation();
        showSidebarDirContextMenu(e, item);
      });
    } else if (editable) {
      div.onclick = () => sidebarOpenFile(item);
      div.addEventListener('contextmenu', e => {
        e.preventDefault();
        e.stopPropagation();
        showSidebarFileContextMenu(e, item);
      });
    } else if (!item.is_dir) {
      div.addEventListener('contextmenu', e => {
        e.preventDefault();
        e.stopPropagation();
        showSidebarFileContextMenu(e, item);
      });
    }

    container.appendChild(div);
  }
  _updateSidebarActiveItem();
}

function _updateSidebarActiveItem() {
  const activeTab  = state.tabs.find(t => t.id === state.activeTabId);
  const activePath = activeTab?.path
    ? activeTab.path.replace(/\\/g, '/').toLowerCase()
    : null;
  document.querySelectorAll('#sidebar-files .sidebar-item[data-path]').forEach(el => {
    const p = el.dataset.path.replace(/\\/g, '/').toLowerCase();
    el.classList.toggle('sidebar-item-active', !!activePath && p === activePath);
  });
}

async function openProjectQuartoFile(filename) {
  const s = loadSettings();
  if (!s.projectDir) { alert('Aucun répertoire de projet défini dans les paramètres.'); return; }
  if (!window.pywebview) return;
  const path = s.projectDir.replace(/\\/g, '/').replace(/\/$/, '') + '/' + filename;
  const data = await window.pywebview.api.open_file_by_path(path);
  if (!data || data.error) { alert(`Fichier introuvable : ${filename}`); return; }
  const existing = state.tabs.find(t => t.path === data.path);
  if (existing) { switchToTab(existing.id); closeSettings(); return; }
  const tab = createTab(data.name, data.path, data.content);
  switchToTab(tab.id);
  closeSettings();
}

async function sidebarOpenFile(item) {
  try {
    const data = await window.pywebview.api.open_file_by_path(item.path);
    if (!data || data.error) return;
    const existing = state.tabs.find(t => t.path === data.path);
    if (existing) { switchToTab(existing.id); return; }
    const tab = createTab(data.name, data.path, data.content);
    switchToTab(tab.id);
  } catch (e) { console.error('sidebarOpenFile:', e); }
}

function toggleSidebar() {
  const sidebar = document.getElementById('sidebar');
  const handle  = document.getElementById('sidebar-resize-handle');
  if (sidebar.classList.contains('collapsed')) {
    sidebar.classList.remove('collapsed');
    if (sidebar._savedWidth) sidebar.style.width = sidebar._savedWidth;
    if (handle) handle.style.display = '';
  } else {
    sidebar._savedWidth = sidebar.style.width || (sidebar.offsetWidth + 'px');
    sidebar.style.width = '';
    sidebar.classList.add('collapsed');
    if (handle) handle.style.display = 'none';
  }
}

// ── Recherche dans les fichiers du volet latéral ──────────────────────────────
let _sidebarSearchTimer = null;

function toggleSidebarSearch() {
  const bar     = document.getElementById('sidebar-search-bar');
  const btn     = document.getElementById('sidebar-search-btn');
  const isOpen  = bar.classList.contains('open');
  if (isOpen) {
    sidebarSearchClose();
  } else {
    bar.classList.add('open');
    btn.classList.add('active');
    document.getElementById('sidebar-files').style.display = '';
    document.getElementById('sidebar-search-results').classList.remove('open');
    setTimeout(() => document.getElementById('sidebar-search-input').focus(), 50);
  }
}

function sidebarSearchClose() {
  document.getElementById('sidebar-search-bar').classList.remove('open');
  document.getElementById('sidebar-search-btn').classList.remove('active');
  document.getElementById('sidebar-search-results').classList.remove('open');
  document.getElementById('sidebar-filename-search-results').classList.remove('open');
  document.getElementById('sidebar-files').style.display = '';
  document.getElementById('sidebar-search-input').value = '';
  document.getElementById('sidebar-search-status').textContent = '';
  document.getElementById('sidebar-search-status').className = '';
  document.getElementById('sidebar-filename-search-input').value = '';
  document.getElementById('sidebar-filename-search-status').textContent = '';
  document.getElementById('sidebar-filename-search-status').className = '';
  clearTimeout(_sidebarSearchTimer);
  clearTimeout(_sidebarFilenameSearchTimer);
}

function sidebarSearchClear() {
  document.getElementById('sidebar-search-input').value = '';
  document.getElementById('sidebar-search-status').textContent = '';
  document.getElementById('sidebar-search-status').className = '';
  document.getElementById('sidebar-search-results').classList.remove('open');
  document.getElementById('sidebar-files').style.display = '';
  document.getElementById('sidebar-search-input').focus();
  clearTimeout(_sidebarSearchTimer);
}

function onSidebarSearchKeydown(e) {
  if (e.key === 'Escape') { sidebarSearchClose(); }
  else if (e.key === 'Enter') { clearTimeout(_sidebarSearchTimer); runSidebarSearch(); }
}

function onSidebarSearchInput() {
  clearTimeout(_sidebarSearchTimer);
  const q = document.getElementById('sidebar-search-input').value.trim();
  if (!q) {
    sidebarSearchClear();
    return;
  }
  _silentResetFilenameSearch();
  _sidebarSearchTimer = setTimeout(runSidebarSearch, 400);
}

async function runSidebarSearch() {
  const q   = document.getElementById('sidebar-search-input').value.trim();
  const dir = _sidebarCurrentDir;
  if (!q || !dir) return;
  if (!window.pywebview || !window.pywebview.api) return;

  const statusEl  = document.getElementById('sidebar-search-status');
  const resultsEl = document.getElementById('sidebar-search-results');
  const filesEl   = document.getElementById('sidebar-files');

  statusEl.textContent = 'Recherche…';
  statusEl.className   = 'searching';
  resultsEl.classList.remove('open');
  filesEl.style.display = 'none';

  try {
    const results = await window.pywebview.api.search_in_files(dir, q);
    _renderSidebarSearchResults(results, q);
  } catch (e) {
    statusEl.textContent = 'Erreur : ' + e;
    statusEl.className   = 'no-result';
  }
}

function _escHtml(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function _highlightMatch(text, query) {
  const lq  = query.toLowerCase();
  const lt  = text.toLowerCase();
  let res = '', pos = 0;
  while (pos < text.length) {
    const idx = lt.indexOf(lq, pos);
    if (idx === -1) { res += _escHtml(text.slice(pos)); break; }
    res += _escHtml(text.slice(pos, idx));
    res += `<mark>${_escHtml(text.slice(idx, idx + query.length))}</mark>`;
    pos = idx + query.length;
  }
  return res;
}

// Stockage temporaire des résultats pour les handlers (évite les problèmes de guillemets dans onclick)
let _ssrResults = [];

function _renderSidebarSearchResults(results, query) {
  const statusEl  = document.getElementById('sidebar-search-status');
  const resultsEl = document.getElementById('sidebar-search-results');
  const filesEl   = document.getElementById('sidebar-files');

  if (!results.length) {
    statusEl.textContent = 'Aucun résultat';
    statusEl.className   = 'no-result';
    resultsEl.classList.remove('open');
    filesEl.style.display = '';
    return;
  }

  _ssrResults = results;

  const total = results.reduce((s, r) => s + r.matches.length, 0);
  statusEl.textContent = `${total} occurrence${total > 1 ? 's' : ''} dans ${results.length} fichier${results.length > 1 ? 's' : ''}`;
  statusEl.className   = '';

  const fileIconSvg = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
    <polyline points="14 2 14 8 20 8"/>
  </svg>`;

  const html = results.map((r, fi) => {
    const matchesHtml = r.matches.map((m, mi) => {
      const trimmed = m.text.trimStart();
      return `<div class="ssr-match" data-fi="${fi}" data-line="${m.line}">
        <span class="ssr-line-num">L${m.line}</span>
        <span class="ssr-line-text">${_highlightMatch(trimmed, query)}</span>
      </div>`;
    }).join('');

    const extraCount = r.matches.length >= 10
      ? `<span class="ssr-count">≥10</span>`
      : `<span class="ssr-count">${r.matches.length}</span>`;

    return `<div class="ssr-file">
      <div class="ssr-file-header" data-fi="${fi}">
        <span class="ssr-file-icon">${fileIconSvg}</span>
        <span class="ssr-file-name">${_escHtml(r.name)}</span>
        ${extraCount}
        <span class="ssr-file-rel" title="${_escHtml(r.rel)}">${_escHtml(r.rel)}</span>
        <button class="ssr-open-btn" data-fi="${fi}">Ouvrir</button>
      </div>
      <div class="ssr-matches">${matchesHtml}</div>
    </div>`;
  }).join('');

  resultsEl.innerHTML = html;

  // Event listeners via délégation — aucun guillemet dans le HTML
  resultsEl.querySelectorAll('.ssr-open-btn').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const fi = parseInt(btn.dataset.fi);
      _ssrOpen(fi);
    });
  });
  resultsEl.querySelectorAll('.ssr-file-header').forEach(hdr => {
    hdr.addEventListener('click', () => {
      const fi = parseInt(hdr.dataset.fi);
      _ssrOpen(fi);
    });
  });
  resultsEl.querySelectorAll('.ssr-match').forEach(row => {
    row.addEventListener('click', () => {
      const fi   = parseInt(row.dataset.fi);
      const line = parseInt(row.dataset.line);
      _ssrOpen(fi, line);
    });
  });

  resultsEl.classList.add('open');
  filesEl.style.display = 'none';
}

function _ssrOpen(fi, line) {
  const r = _ssrResults[fi];
  if (r) sidebarSearchOpenFile(r.path, line);
}

async function sidebarSearchOpenFile(path, line) {
  try {
    const data = await window.pywebview.api.open_file_by_path(path);
    if (!data || data.error) return;
    const existing = state.tabs.find(t => t.path === data.path || t.path?.replace(/\\/g,'/') === path);
    const tab = existing || createTab(data.name, data.path, data.content);
    switchToTab(tab.id);
    if (!existing) { /* tab already switched */ }
    // Scroll to line if provided
    if (line && !state.sourceMode) {
      // wait for preview render then scroll
      setTimeout(() => {
        const pane = document.getElementById('preview-pane');
        if (!pane) return;
        const rows = document.querySelectorAll('#preview [id], #preview h1, #preview h2, #preview h3, #preview p');
        // estimate line position via line height
        const content = tab.content || '';
        const lines   = content.split('\n');
        const ratio   = Math.min(1, (line - 1) / Math.max(1, lines.length));
        pane.scrollTop = ratio * pane.scrollHeight;
      }, 80);
    }
  } catch (e) { console.error('sidebarSearchOpenFile:', e); }
}

// ── Recherche dans les noms de fichiers du volet latéral ─────────────────────
let _sidebarFilenameSearchTimer = null;
let _sfrResults = [];

function _silentResetFilenameSearch() {
  const input = document.getElementById('sidebar-filename-search-input');
  if (!input) return;
  if (!input.value && !document.getElementById('sidebar-filename-search-results').classList.contains('open')) return;
  input.value = '';
  document.getElementById('sidebar-filename-search-status').textContent = '';
  document.getElementById('sidebar-filename-search-status').className = '';
  document.getElementById('sidebar-filename-search-results').classList.remove('open');
  document.getElementById('sidebar-filename-search-results').innerHTML = '';
  clearTimeout(_sidebarFilenameSearchTimer);
}

function _silentResetContentSearch() {
  const input = document.getElementById('sidebar-search-input');
  if (!input) return;
  if (!input.value && !document.getElementById('sidebar-search-results').classList.contains('open')) return;
  input.value = '';
  document.getElementById('sidebar-search-status').textContent = '';
  document.getElementById('sidebar-search-status').className = '';
  document.getElementById('sidebar-search-results').classList.remove('open');
  document.getElementById('sidebar-files').style.display = '';
  clearTimeout(_sidebarSearchTimer);
}

function onSidebarFilenameSearchKeydown(e) {
  if (e.key === 'Escape') { sidebarSearchClose(); }
  else if (e.key === 'Enter') { clearTimeout(_sidebarFilenameSearchTimer); runSidebarFilenameSearch(); }
}

function onSidebarFilenameSearchInput() {
  clearTimeout(_sidebarFilenameSearchTimer);
  const q = document.getElementById('sidebar-filename-search-input').value.trim();
  if (!q) {
    sidebarFilenameSearchClear();
    return;
  }
  _silentResetContentSearch();
  _sidebarFilenameSearchTimer = setTimeout(runSidebarFilenameSearch, 300);
}

function sidebarFilenameSearchClear() {
  document.getElementById('sidebar-filename-search-input').value = '';
  document.getElementById('sidebar-filename-search-status').textContent = '';
  document.getElementById('sidebar-filename-search-status').className = '';
  document.getElementById('sidebar-filename-search-results').classList.remove('open');
  document.getElementById('sidebar-filename-search-results').innerHTML = '';
  document.getElementById('sidebar-files').style.display = '';
  document.getElementById('sidebar-filename-search-input').focus();
  clearTimeout(_sidebarFilenameSearchTimer);
}

async function runSidebarFilenameSearch() {
  const q   = document.getElementById('sidebar-filename-search-input').value.trim();
  const dir = _sidebarCurrentDir;
  if (!q || !dir) return;
  if (!window.pywebview || !window.pywebview.api) return;

  const statusEl  = document.getElementById('sidebar-filename-search-status');
  const resultsEl = document.getElementById('sidebar-filename-search-results');
  const filesEl   = document.getElementById('sidebar-files');

  statusEl.textContent = 'Recherche…';
  statusEl.className   = 'searching';
  resultsEl.classList.remove('open');
  filesEl.style.display = 'none';

  try {
    const results = await window.pywebview.api.search_files_by_name(dir, q);
    _renderSidebarFilenameSearchResults(results, q);
  } catch (e) {
    statusEl.textContent = 'Erreur : ' + e;
    statusEl.className   = 'no-result';
  }
}

function _renderSidebarFilenameSearchResults(results, query) {
  const statusEl  = document.getElementById('sidebar-filename-search-status');
  const resultsEl = document.getElementById('sidebar-filename-search-results');
  const filesEl   = document.getElementById('sidebar-files');

  if (!results.length) {
    statusEl.textContent = 'Aucun résultat';
    statusEl.className   = 'no-result';
    resultsEl.classList.remove('open');
    filesEl.style.display = '';
    return;
  }

  _sfrResults = results;
  statusEl.textContent = `${results.length} fichier${results.length > 1 ? 's' : ''} trouvé${results.length > 1 ? 's' : ''}`;
  statusEl.className   = '';

  const fileIconSvg = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
    <polyline points="14 2 14 8 20 8"/>
  </svg>`;

  const html = results.map((r, fi) => `<div class="ssr-file">
    <div class="ssr-file-header sfr-file-header" data-fi="${fi}">
      <span class="ssr-file-icon">${fileIconSvg}</span>
      <span class="ssr-file-name">${_highlightMatch(r.name, query)}</span>
      <span class="ssr-file-rel" title="${_escHtml(r.rel)}">${_escHtml(r.rel)}</span>
      <button class="ssr-open-btn sfr-open-btn" data-fi="${fi}">Ouvrir</button>
    </div>
  </div>`).join('');

  resultsEl.innerHTML = html;

  resultsEl.querySelectorAll('.sfr-open-btn').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      _sfrOpen(parseInt(btn.dataset.fi));
    });
  });
  resultsEl.querySelectorAll('.sfr-file-header').forEach(hdr => {
    hdr.addEventListener('click', () => _sfrOpen(parseInt(hdr.dataset.fi)));
  });

  resultsEl.classList.add('open');
  filesEl.style.display = 'none';
}

function _sfrOpen(fi) {
  const r = _sfrResults[fi];
  if (r) sidebarSearchOpenFile(r.path);
}

async function changeSidebarDir() {
  try {
    const path = await window.pywebview.api.choose_directory();
    if (!path) return;
    await window.pywebview.api.set_default_dir(path);
    await loadSidebarDir(path);
  } catch (e) { console.error('changeSidebarDir:', e); }
}

// ── Status bar ────────────────────────────────────────────────────────────────
const _FILE_TYPE_MAP = {
  md: 'Markdown', qmd: 'Quarto',
  html: 'HTML', htm: 'HTML',
  pdf: 'PDF',
  yml: 'YAML', yaml: 'YAML',
  log: 'Log', tex: 'LaTeX',
  puml: 'PlantUML', plantuml: 'PlantUML',
  png: 'PNG', jpg: 'JPEG', jpeg: 'JPEG',
  gif: 'GIF', bmp: 'BMP', webp: 'WebP', svg: 'SVG', ico: 'ICO',
};

function updateFileStatus(tab) {
  document.getElementById('status-file').textContent =
    tab.path ? tab.path : `[${tab.name}]`;

  const ext = tab.path
    ? tab.path.toLowerCase().split('.').pop()
    : '';
  document.getElementById('status-filetype').textContent =
    _FILE_TYPE_MAP[ext] || (ext ? ext.toUpperCase() : '');

  _updateCursorStatus();
}

// ── Curseur ligne / colonne ───────────────────────────────────────────────────
function _cursorFromTextarea(ta) {
  if (!ta) return null;
  const before = ta.value.substring(0, ta.selectionStart);
  const lines  = before.split('\n');
  return { line: lines.length, col: lines[lines.length - 1].length + 1 };
}

function _updateCursorStatus() {
  const el = document.getElementById('status-cursor');
  if (!el) return;

  let pos = null;
  if (state.sourceMode) {
    pos = _cursorFromTextarea(document.getElementById('source-editor'));
  } else if (document.body.classList.contains('yaml-mode')) {
    pos = _cursorFromTextarea(document.getElementById('yaml-editor'));
  } else if (document.body.classList.contains('puml-mode')) {
    pos = _cursorFromTextarea(document.getElementById('puml-editor'));
  } else if (document.body.classList.contains('css-mode')) {
    pos = _cursorFromTextarea(document.getElementById('css-editor'));
  } else if (document.body.classList.contains('json-mode')) {
    pos = _cursorFromTextarea(document.getElementById('json-editor'));
  } else if (document.body.classList.contains('lua-mode')) {
    pos = _cursorFromTextarea(document.getElementById('lua-editor'));
  } else if (document.body.classList.contains('xml-mode')) {
    pos = _cursorFromTextarea(document.getElementById('xml-editor'));
  }

  el.textContent = pos ? `Ln ${pos.line}, Col ${pos.col}` : '';
  document.querySelector('.status-sep').style.display = pos ? '' : 'none';
  _updateTOCActive();
}

function _updateTOCActive() {
  const body = document.getElementById('toc-body');
  if (!body) return;
  const items = [...body.querySelectorAll('.toc-item')];
  if (!items.length) return;

  if (state.sourceMode) {
    // Mode source (textarea) : chercher le dernier titre au-dessus du curseur
    const ta = document.getElementById('source-editor');
    if (!ta) return;
    const lines = ta.value.substring(0, ta.selectionStart).split('\n');
    let activeNorm = null;
    for (let i = lines.length - 1; i >= 0; i--) {
      const m = lines[i].match(/^(#{1,6})\s+(.+)/);
      if (m) {
        let text = m[2]
          .replace(/\s*\{#[^}]*\}\s*$/, '')
          .replace(/\*\*(.+?)\*\*/g, '$1')
          .replace(/\*(.+?)\*/g, '$1')
          .replace(/__(.+?)__/g, '$1')
          .replace(/_(.+?)_/g, '$1')
          .replace(/`(.+?)`/g, '$1')
          .replace(/\[([^\]]+)\]\([^)]*\)/g, '$1')
          .trim();
        activeNorm = _headingToId(text);
        break;
      }
    }
    let activeItem = null;
    items.forEach(item => {
      const span = item.querySelector('.toc-text');
      const norm = _headingToId((span?.textContent || '').trim());
      const isActive = activeNorm !== null && norm === activeNorm;
      item.classList.toggle('toc-active', isActive);
      if (isActive) activeItem = item;
    });
    if (activeItem) activeItem.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  } else {
    // Mode preview (contenteditable) : trouver le dernier titre avant la position du curseur dans le DOM
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const preview = document.getElementById('preview');
    if (!preview || !preview.contains(sel.anchorNode)) return;
    const headings = [...preview.querySelectorAll('h1,h2,h3,h4,h5,h6')];
    if (!headings.length) return;
    let activeHeading = null;
    for (const h of headings) {
      const pos = h.compareDocumentPosition(sel.anchorNode);
      // Curseur à l'intérieur du titre → priorité absolue
      if (pos & Node.DOCUMENT_POSITION_CONTAINED_BY) { activeHeading = h; break; }
      // Titre précède le curseur → candidat courant
      if (pos & Node.DOCUMENT_POSITION_FOLLOWING) activeHeading = h;
    }
    let activeItem = null;
    items.forEach(item => {
      const isActive = activeHeading !== null && item.dataset.headingId === activeHeading.id;
      item.classList.toggle('toc-active', isActive);
      if (isActive) activeItem = item;
    });
    if (activeItem) activeItem.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }
}

function _setupCursorTracking() {
  // Textareas : on écoute tous les événements qui déplacent le curseur
  const textareas = [
    { id: 'source-editor',  check: () => state.sourceMode },
    { id: 'yaml-editor',    check: () => document.body.classList.contains('yaml-mode') },
    { id: 'puml-editor',    check: () => document.body.classList.contains('puml-mode') },
    { id: 'css-editor',     check: () => document.body.classList.contains('css-mode') },
    { id: 'json-editor',    check: () => document.body.classList.contains('json-mode') },
    { id: 'lua-editor',     check: () => document.body.classList.contains('lua-mode') },
    { id: 'xml-editor',     check: () => document.body.classList.contains('xml-mode') },
  ];
  for (const { id, check } of textareas) {
    const ta = document.getElementById(id);
    if (!ta) continue;
    const handler = () => { if (check()) _updateCursorStatus(); };
    ta.addEventListener('keyup',    handler);
    ta.addEventListener('mouseup',  handler);
    ta.addEventListener('click',    handler);
    ta.addEventListener('focus',    handler);
    ta.addEventListener('input',    handler);
    // selectionchange ne bubble pas sur textarea, on utilise l'event document
  }
  // selectionchange global (couvre aussi les cas non clavier)
  document.addEventListener('selectionchange', () => {
    const active = document.activeElement;
    if (active && (active.id === 'source-editor' ||
                   active.id === 'yaml-editor'   ||
                   active.id === 'puml-editor'   ||
                   active.id === 'css-editor'    ||
                   active.id === 'json-editor'   ||
                   active.id === 'lua-editor'    ||
                   active.id === 'xml-editor')) {
      _updateCursorStatus();
    }
  });
}

// ── File operations ───────────────────────────────────────────────────────────
async function openFile() {
  if (!window.pywebview) { alert('pywebview non disponible'); return; }
  const result = await window.pywebview.api.open_file();
  if (!result) return;
  if (result.error) { alert('Erreur : ' + result.error); return; }

  const existing = state.tabs.find(t => t.path === result.path);
  if (existing) { switchToTab(existing.id); return; }

  const tab = createTab(result.name, result.path, result.content);
  switchToTab(tab.id);
  sidebarNavigateTo(result.path);
}

async function saveFile() {
  const tab = getActiveTab();
  if (!tab) return;

  // Synchroniser tab.content depuis la textarea active (fallback fiable)
  const _modeMap = {
    'puml-mode': 'puml-editor',
    'yaml-mode': 'yaml-editor',
    'css-mode':  'css-editor',
    'lua-mode':  'lua-editor',
    'json-mode': 'json-editor',
    'xml-mode':  'xml-editor',
  };
  for (const [cls, taId] of Object.entries(_modeMap)) {
    if (document.body.classList.contains(cls)) {
      const ta = document.getElementById(taId);
      if (ta) tab.content = ta.value;
      break;
    }
  }

  if (tab.path) {
    if (!window.pywebview) return;
    const res = await window.pywebview.api.save_file(tab.path, tab.content);
    if (res && res.error) { alert('Erreur : ' + res.error); return; }
    tab.savedContent = tab.content;
    tab.modified = false;
    renderTabList();
    updateFileStatus(tab);
    showToast('✓ Document enregistré');
  } else {
    await saveFileAs();
  }
}

async function saveFileAs() {
  const tab = getActiveTab();
  if (!tab) return;
  if (!window.pywebview) return;
  const res = await window.pywebview.api.save_file_as(tab.content, tab.name);
  if (!res) return;
  if (res.error) { alert('Erreur : ' + res.error); return; }
  tab.path = res.path;
  tab.name = res.name;
  tab.savedContent = tab.content;
  tab.modified = false;
  renderTabList();
  updateFileStatus(tab);
  showToast('✓ Document enregistré');
}

function closeCurrentTab() {
  if (state.activeTabId !== null) closeTab(state.activeTabId);
}

function newTab() {
  const tab = createTab('Sans titre', null, '');
  switchToTab(tab.id);
}

// ── Paramètres ────────────────────────────────────────────────────────────────
const SETTINGS_KEY = 'md_editor_settings';

function loadSettings() {
  try { return JSON.parse(localStorage.getItem(SETTINGS_KEY)) || {}; } catch { return {}; }
}

function saveSettings(s) {
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(s));
  if (window.pywebview && window.pywebview.api) {
    window.pywebview.api.save_settings(s);
  }
}

function switchSettingsTab(name) {
  document.querySelectorAll('.settings-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === name));
  document.querySelectorAll('.settings-tab-panel').forEach(p => p.classList.toggle('active', p.id === 'stab-' + name));
}

function _updateSettingsProjectLabel(dir) {
  const el = document.getElementById('settings-project-path');
  if (!el) return;
  if (dir) {
    const name = dir.replace(/\\/g, '/').split('/').filter(Boolean).pop() || dir;
    el.textContent = name;
    el.title = dir;
  } else {
    el.textContent = '(aucun projet)';
    el.title = '';
  }
}

function openSettings() {
  const s = loadSettings();
  _updateSettingsProjectLabel(s.projectDir || '');
  switchSettingsTab('general');
  document.getElementById('project-dir').value  = s.projectDir  ?? '';
  document.getElementById('pdf-file-name').value = s.pdfFileName ?? '';
  document.getElementById('compile-cmd-html').value = s.compileCmdHtml ?? '';
  document.getElementById('compile-cmd-pdf').value  = s.compileCmdPdf  ?? '';
  document.getElementById('bib-file').value = s.bibFile ?? '';
  document.getElementById('spellcheck-enabled').checked = s.spellcheck ?? false;
  document.getElementById('spellcheck-lang').value = s.lang ?? 'fr';
  const zd = s.zoomDefault ?? 100;
  document.getElementById('zoom-default-slider').value = zd;
  document.getElementById('zoom-default-label').textContent = zd + '%';
  document.getElementById('sidebar-width-input').value    = s.sidebarWidth    ?? 370;
  document.getElementById('toc-width-input').value        = s.tocWidth        ?? 370;
  document.getElementById('log-minimap-width-input').value = s.logMinimapWidth ?? 100;
  document.getElementById('plantuml-jar').value = s.plantumlJar ?? '';
  document.getElementById('xsd-default-dir').value = s.xsdDefaultDir ?? '';
  document.getElementById('tab-size-input').value = s.tabSize ?? 4;
  document.getElementById('source-wrap-enabled').checked = s.sourceWrap ?? false;
  document.getElementById('justify-enabled').checked  = s.justify  ?? true;
  document.getElementById('edit-bar-enabled').checked = s.editBar  ?? true;
  const lh = s.lineHeight ?? 1.4;
  document.getElementById('line-height-slider').value = lh;
  document.getElementById('line-height-label').textContent = parseFloat(lh).toFixed(1);
  const lhCode = s.lineHeightCode ?? 1.45;
  document.getElementById('line-height-code-slider').value = lhCode;
  document.getElementById('line-height-code-label').textContent = parseFloat(lhCode).toFixed(2);
  document.getElementById('font-text').value       = s.fontText      ?? 'system-ui, sans-serif';
  document.getElementById('font-table').value      = s.fontTable     ?? 'system-ui, sans-serif';
  document.getElementById('font-code').value       = s.fontCode      ?? "'JetBrains Mono', 'Fira Code', monospace";
  document.getElementById('font-yaml').value       = s.fontYaml      ?? "'JetBrains Mono', 'Fira Code', monospace";
  document.getElementById('font-size-text').value  = s.fontSizeText  ?? '14px';
  document.getElementById('font-size-table').value = s.fontSizeTable ?? '13px';
  document.getElementById('font-size-code').value  = s.fontSizeCode  ?? '13px';
  document.getElementById('font-size-yaml').value  = s.fontSizeYaml  ?? '13px';
  _THEMES.forEach(t => {
    const key = 'headingColor' + t[0].toUpperCase() + t.slice(1);
    document.getElementById('heading-color-' + t).value =
      s[key] ?? s.headingColor ?? _HEADING_COLOR_DEFAULTS[t];
  });
  document.getElementById('settings-overlay').classList.add('open');
}

function closeSettings() {
  document.getElementById('settings-overlay').classList.remove('open');
}

function closeSettingsOnOverlay(e) {
  if (e.target === document.getElementById('settings-overlay')) closeSettings();
}

function applySettings() {
  const projectDir   = document.getElementById('project-dir').value;
  const pdfFileName  = document.getElementById('pdf-file-name').value.trim();
  const compileCmdHtml = document.getElementById('compile-cmd-html').value.trim();
  const compileCmdPdf  = document.getElementById('compile-cmd-pdf').value.trim();
  const bibFile      = document.getElementById('bib-file').value.trim();
  const enabled      = document.getElementById('spellcheck-enabled').checked;
  const lang         = document.getElementById('spellcheck-lang').value;
  const zoomDefault  = parseInt(document.getElementById('zoom-default-slider').value);
  const sidebarWidth    = parseInt(document.getElementById('sidebar-width-input').value);
  const tocWidth        = parseInt(document.getElementById('toc-width-input').value);
  const logMinimapWidth = parseInt(document.getElementById('log-minimap-width-input').value);
  const plantumlJar  = document.getElementById('plantuml-jar').value.trim();
  const xsdDefaultDir = document.getElementById('xsd-default-dir').value.trim();
  const tabSize      = Math.min(8, Math.max(1, parseInt(document.getElementById('tab-size-input').value) || 4));
  const sourceWrap   = document.getElementById('source-wrap-enabled').checked;
  const justify      = document.getElementById('justify-enabled').checked;
  const editBar      = document.getElementById('edit-bar-enabled').checked;
  const lineHeight     = parseFloat(document.getElementById('line-height-slider').value);
  const lineHeightCode = parseFloat(document.getElementById('line-height-code-slider').value);
  const fontText      = document.getElementById('font-text').value;
  const fontTable     = document.getElementById('font-table').value;
  const fontCode      = document.getElementById('font-code').value;
  const fontYaml      = document.getElementById('font-yaml').value;
  const fontSizeText  = document.getElementById('font-size-text').value;
  const fontSizeTable = document.getElementById('font-size-table').value;
  const fontSizeCode  = document.getElementById('font-size-code').value;
  const fontSizeYaml  = document.getElementById('font-size-yaml').value;
  const headingColorDark   = document.getElementById('heading-color-dark').value;
  const headingColorLight  = document.getElementById('heading-color-light').value;
  const headingColorSystem = document.getElementById('heading-color-system').value;
  _bibEntries = null; // invalider le cache si le .bib a changé
  saveSettings({ projectDir, pdfFileName, compileCmdHtml, compileCmdPdf, bibFile, plantumlJar, xsdDefaultDir, tabSize, sourceWrap, tabStops: [..._tabStops], spellcheck: enabled, lang, zoomDefault, sidebarWidth, tocWidth, logMinimapWidth, justify, editBar, lineHeight, lineHeightCode, fontText, fontTable, fontCode, fontYaml, fontSizeText, fontSizeTable, fontSizeCode, fontSizeYaml, headingColorDark, headingColorLight, headingColorSystem });
  _applyTabSize(tabSize);
  _applySourceWrap(sourceWrap);
  _rulerCharW = 0; _rulerLeftOff = 0; _renderRuler();

  const preview = document.getElementById('preview');
  const source  = document.getElementById('source-editor');
  [preview, source].forEach(el => { el.spellcheck = enabled; el.lang = lang; });
  document.documentElement.lang = lang;
  preview.classList.toggle('justified', justify);
  preview.classList.toggle('no-edit-bar', !editBar);
  preview.style.lineHeight = lineHeight;
  source.style.lineHeight  = lineHeight;
  document.documentElement.style.setProperty('--line-height-code', lineHeightCode ?? 1.45);
  _applyFonts({ fontText, fontTable, fontCode, fontYaml, fontSizeText, fontSizeTable, fontSizeCode, fontSizeYaml, headingColorDark, headingColorLight, headingColorSystem });
  document.getElementById('zoom-slider').value = zoomDefault;
  setZoom(zoomDefault);
  document.getElementById('sidebar').style.width   = sidebarWidth + 'px';
  document.getElementById('toc-panel').style.width = tocWidth + 'px';
  _applyLogMinimapWidth(logMinimapWidth);
  if (projectDir) loadSidebarDir(projectDir);
  closeSettings();
}

async function chooseProjectDir() {
  if (!window.pywebview || !window.pywebview.api) return;
  try {
    const path = await window.pywebview.api.choose_directory();
    if (path) document.getElementById('project-dir').value = path.replace(/\\/g, '/');
  } catch (e) {}
}

async function openProjectHtml() {
  const s = loadSettings();
  if (!s.projectDir) { alert('Aucun répertoire de projet défini dans les paramètres.'); return; }
  const path = s.projectDir.replace(/\\/g, '/').replace(/\/$/, '') + '/_book/index.html';
  const existing = state.tabs.find(t => t.path === path);
  if (existing) { switchToTab(existing.id); return; }
  try {
    const data = await window.pywebview.api.open_file_by_path(path);
    if (!data || data.error) { alert('Fichier introuvable :\n' + path); return; }
    const tab = createTab(data.name, data.path, data.content);
    renderTabList();
    switchToTab(tab.id);
  } catch (e) { alert('Erreur ouverture HTML : ' + e); }
}

async function openProjectPdf() {
  const s = loadSettings();
  if (!s.projectDir) { alert('Aucun répertoire de projet défini dans les paramètres.'); return; }
  if (!s.pdfFileName) { alert('Nom du fichier PDF non défini.\nRenseignez-le dans Paramètres → Répertoire du projet.'); return; }
  const path = s.projectDir.replace(/\\/g, '/').replace(/\/$/, '') + '/_book/' + s.pdfFileName;
  const existing = state.tabs.find(t => t.path === path);
  if (existing) { switchToTab(existing.id); return; }
  try {
    const data = await window.pywebview.api.open_file_by_path(path);
    if (!data || data.error) { alert('Fichier introuvable :\n' + path); return; }
    const tab = createTab(data.name, data.path, data.content);
    renderTabList();
    switchToTab(tab.id);
  } catch (e) { alert('Erreur ouverture PDF : ' + e); }
}

function _applyLogMinimapWidth(w) {
  const canvas = document.getElementById('log-minimap');
  if (!canvas) return;
  const px = Math.max(60, Math.min(300, parseInt(w) || 100));
  canvas.style.width = px + 'px';
  _drawLogMinimap();
}

function _applySourceWrap(enabled) {
  const ta     = document.getElementById('source-editor');
  const scroll = document.getElementById('source-scroll');
  if (!ta || !scroll) return;
  ta.style.whiteSpace    = enabled ? 'pre-wrap' : 'pre';
  scroll.style.overflowX = enabled ? 'hidden'   : 'auto';
  if (!enabled) {
    // Recalculer la largeur du wrapper pour le scroll horizontal
    const ln = document.getElementById('source-line-numbers');
    _setEditorWrapperWidth(ta, ln, 'source-wrapper', 'source-scroll');
  } else {
    // En mode wrap, le wrapper n'a pas besoin de dépasser la largeur du scroll
    const wrapper = document.getElementById('source-wrapper');
    if (wrapper) wrapper.style.minWidth = '';
  }
}

function _applyTabSize(n) {
  const size = Math.min(8, Math.max(1, parseInt(n) || 4));
  document.documentElement.style.setProperty('--tab-size', size);
  document.querySelectorAll('pre, code, #source-editor, #puml-editor, #yaml-editor, #css-editor, #lua-editor, #json-editor, #xml-editor').forEach(el => {
    el.style.tabSize = size;
    el.style.MozTabSize = size;
  });
}

function _applyFonts(s) {
  const root = document.documentElement;
  root.style.setProperty('--font-text',       s.fontText      ?? 'system-ui, sans-serif');
  root.style.setProperty('--font-table',      s.fontTable     ?? 'system-ui, sans-serif');
  root.style.setProperty('--font-code',       s.fontCode      ?? "'JetBrains Mono', 'Fira Code', monospace");
  root.style.setProperty('--font-yaml',       s.fontYaml      ?? "'JetBrains Mono', 'Fira Code', monospace");
  root.style.setProperty('--font-size-text',  s.fontSizeText  ?? '14px');
  root.style.setProperty('--font-size-table', s.fontSizeTable ?? '13px');
  root.style.setProperty('--font-size-code',  s.fontSizeCode  ?? '13px');
  root.style.setProperty('--font-size-yaml',  s.fontSizeYaml  ?? '13px');
  const _theme   = s.theme ?? loadSettings().theme ?? 'dark';
  const _colorKey = 'headingColor' + _theme[0].toUpperCase() + _theme.slice(1);
  const _hColor  = s[_colorKey] ?? s.headingColor ?? null;
  if (_hColor) root.style.setProperty('--heading-color', _hColor);
  else root.style.removeProperty('--heading-color');
}

function _getAccentHex() {
  return getComputedStyle(document.documentElement).getPropertyValue('--accent').trim() || '#89b4fa';
}

function resetHeadingColor(theme) {
  const s   = loadSettings();
  const key = 'headingColor' + theme[0].toUpperCase() + theme.slice(1);
  delete s[key];
  saveSettings(s);
  document.getElementById('heading-color-' + theme).value = _HEADING_COLOR_DEFAULTS[theme];
  // Si le thème réinitialisé est le thème courant, retirer la variable CSS
  if ((s.theme || 'dark') === theme) {
    document.documentElement.style.removeProperty('--heading-color');
  }
}

function _applySettingsObj(s) {
  const preview    = document.getElementById('preview');
  const source     = document.getElementById('source-editor');
  const enabled    = s.spellcheck ?? false;
  const lang       = s.lang ?? 'fr';
  const lineHeight = s.lineHeight ?? 1.4;
  [preview, source].forEach(el => { el.spellcheck = enabled; el.lang = lang; });
  document.documentElement.lang = lang;
  preview.classList.toggle('justified',   s.justify ?? true);
  preview.classList.toggle('no-edit-bar', !(s.editBar ?? true));
  preview.style.lineHeight = lineHeight;
  source.style.lineHeight  = lineHeight;
  document.documentElement.style.setProperty('--line-height-code', s.lineHeightCode ?? 1.45);
  _applySourceWrap(s.sourceWrap ?? false);
  _applyFonts(s);
  if (s.zoomDefault) {
    document.getElementById('zoom-slider').value = s.zoomDefault;
    setZoom(s.zoomDefault);
  }
  if (s.sidebarWidth)    document.getElementById('sidebar').style.width   = s.sidebarWidth + 'px';
  if (s.tocWidth)        document.getElementById('toc-panel').style.width = s.tocWidth     + 'px';
  if (s.contentMaxWidth) {
    _contentMaxWidth = s.contentMaxWidth;
    document.documentElement.style.setProperty('--content-max-width', _contentMaxWidth + 'px');
    const lbl = document.getElementById('status-width-label');
    if (lbl) lbl.textContent = _contentMaxWidth + 'px';
    const slider = document.getElementById('status-width-slider');
    if (slider) slider.value = String(_contentMaxWidth);
  }
  _applyLogMinimapWidth(s.logMinimapWidth ?? 100);
  _applyTabSize(s.tabSize ?? 4);
  if (s.projectDir && window.pywebview && window.pywebview.api) {
    loadSidebarDir(s.projectDir);
  }
}

function initSettings() {
  _applySettingsObj(loadSettings()); // paramètres en cache (localStorage) appliqués immédiatement

  const _init = async () => {
    try {
      const proj = await window.pywebview.api.get_current_project();
      if (proj && proj.exists) {
        const fs = await window.pywebview.api.load_settings();
        if (fs && Object.keys(fs).length > 0) {
          localStorage.setItem(SETTINGS_KEY, JSON.stringify(fs));
          _applySettingsObj(fs);
        }
      }
      // Toujours afficher le choix des projets, le précédent en premier
      openProjectChooser(proj?.recent || [], proj?.exists ? proj.dir : null);
    } catch (e) {}
  };

  if (window.pywebview && window.pywebview.api) {
    _init();
  } else {
    window.addEventListener('pywebviewready', _init);
  }
}

// ── Sélecteur de projet ───────────────────────────────────────────────────────
// Vrai lorsque le sélecteur de projet a été ouvert depuis la fenêtre des
// paramètres (bouton « Changer de projet… ») : permet de rouvrir les
// paramètres une fois le choix effectué (ou annulé).
let _projectChooserFromSettings = false;

// Ouvre le même sélecteur de projet qu'au lancement de l'application,
// depuis le bouton « Changer de projet… » des paramètres.
async function openProjectChooserFromSettings() {
  if (!window.pywebview) return;
  try {
    const proj = await window.pywebview.api.get_current_project();
    _projectChooserFromSettings = true;
    closeSettings();
    openProjectChooser(proj?.recent || [], proj?.exists ? proj.dir : null);
  } catch (e) {}
}

function openProjectChooser(recent, current) {
  const list = document.getElementById('proj-recent-list');
  list.innerHTML = '';
  // Le projet précédent (actuel) en premier dans la liste
  let ordered = recent || [];
  if (current) {
    ordered = [current, ...ordered.filter(dir => dir !== current)];
  }
  if (ordered.length > 0) {
    const title = document.createElement('div');
    title.className = 'proj-recent-title';
    title.textContent = 'Projets récents';
    list.appendChild(title);
    ordered.forEach(dir => {
      const isCurrent = dir === current;
      const btn = document.createElement('button');
      btn.className = 'proj-recent-item' + (isCurrent ? ' proj-recent-item--current' : '');
      btn.title = dir;
      btn.onclick = () => isCurrent ? closeProjectChooser() : openRecentProject(dir);
      const name = dir.replace(/\\/g, '/').split('/').filter(Boolean).pop() || dir;
      const suffix = isCurrent ? ' <span class="proj-recent-current-tag">(actuel)</span>' : '';
      btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg><span class="proj-recent-name">${name}${suffix}</span><span class="proj-recent-path">${dir}</span>`;
      list.appendChild(btn);
    });
    list.style.display = '';
  } else {
    list.style.display = 'none';
  }
  document.getElementById('project-chooser-overlay').classList.add('open');
}

function closeProjectChooser() {
  document.getElementById('project-chooser-overlay').classList.remove('open');
  // Si le sélecteur a été ouvert depuis les paramètres et que l'utilisateur
  // ferme/annule sans choisir de projet (bouton « Passer »), revenir aux paramètres.
  if (_projectChooserFromSettings) {
    _projectChooserFromSettings = false;
    openSettings();
  }
}

// Bouton « Parcourir… » du sélecteur de projet.
function chooseProjectFromDialog() {
  chooseProject(_projectChooserFromSettings);
}

async function chooseProject(fromSettings) {
  if (!window.pywebview) return;
  const result = await window.pywebview.api.choose_project_dir();
  if (!result) return;
  await _applyProject(result, fromSettings);
}

async function openRecentProject(dir) {
  if (!window.pywebview) return;
  const result = await window.pywebview.api.open_recent_project(dir);
  if (!result || result.error) {
    alert(result?.error || 'Impossible d\'ouvrir ce projet.');
    return;
  }
  await _applyProject(result, _projectChooserFromSettings);
}

async function _applyProject(result, fromSettings) {
  _projectChooserFromSettings = false;
  closeProjectChooser();
  const fs = result.settings || {};
  // Forcer projectDir = répertoire choisi dans les settings
  if (!fs.projectDir) fs.projectDir = result.dir;
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(fs));
  _applySettingsObj(fs);
  _updateSettingsProjectLabel(result.dir);
  if (fromSettings) {
    // Recharger l'affichage des champs dans la boîte de dialogue
    openSettings();
  }
}

// ── Menu contextuel ───────────────────────────────────────────────────────────
let _tableContextCell = null; // cellule cible du menu tableau
let _tableLinkRange   = null; // sélection sauvegardée pour l'insertion de lien dans tableau
let _ctxTightList    = null; // liste cible du menu contextuel normal
let _ctxOrderedList  = null; // <ol> cible pour numéro de départ
let _ctxOrderedListSourceLine = -1; // ligne source du premier item de la liste

// ── Correcteur orthographique ─────────────────────────────────────────────────
const _ignoredWordsSession = new Set();
let _spellCurrentWord  = '';
let _spellCurrentRange = null; // Range (preview) ou {start, end} (source)
let _spellPending      = false;

function hideContextMenu() {
  document.getElementById('context-menu').classList.remove('visible');
  document.getElementById('table-context-menu').classList.remove('visible');
  _hideCitationTooltip();
}

function ctxCopy() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const text = ta.value.substring(ta.selectionStart, ta.selectionEnd);
    if (text) navigator.clipboard.writeText(text);
  } else {
    document.execCommand('copy');
  }
  hideContextMenu();
}

function ctxCut() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const s = ta.selectionStart, e = ta.selectionEnd;
    if (s !== e) {
      navigator.clipboard.writeText(ta.value.substring(s, e));
      ta.setRangeText('', s, e, 'start');
      onSourceInput();
    }
  } else {
    document.execCommand('cut');
    syncPreviewToContent();
  }
  hideContextMenu();
}

function ctxPaste() {
  hideContextMenu();
  navigator.clipboard.readText().then(text => {
    if (!text) return;
    if (state.sourceMode) {
      const ta = document.getElementById('source-editor');
      const s = ta.selectionStart;
      ta.setRangeText(text, s, ta.selectionEnd, 'end');
      onSourceInput();
    } else {
      // Restaurer la sélection sauvegardée avant d'insérer
      if (_scInsertionRange) {
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(_scInsertionRange);
      }
      document.execCommand('insertText', false, text);
      syncPreviewToContent();
    }
  });
}

async function pasteNoFormat() {
  const text = await navigator.clipboard.readText().catch(() => '');
  if (!text) return;
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    ta.setRangeText(text, ta.selectionStart, ta.selectionEnd, 'end');
    onSourceInput();
    ta.focus();
  } else {
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    // Restaurer la sélection si le menu contextuel a pris le focus
    if (_scInsertionRange) {
      sel.removeAllRanges();
      sel.addRange(_scInsertionRange);
    }
    if (!sel || !preview.contains(sel.anchorNode)) return;
    document.execCommand('insertText', false, text);
    syncPreviewToContent();
  }
}

let _copiedStyle = null;

function ctxCopyStyle() {
  if (state.sourceMode) { hideContextMenu(); return; }
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) { hideContextMenu(); return; }

  const node = sel.getRangeAt(0).commonAncestorContainer;
  const el = node.nodeType === Node.TEXT_NODE ? node.parentElement : node;

  _copiedStyle = {
    bold:          !!el.closest('strong, b'),
    italic:        !!el.closest('em, i'),
    underline:     !!el.closest('u'),
    strikethrough: !!el.closest('s, del'),
    smallcaps:     !!el.closest('.smallcaps'),
    textColor:     el.closest('.md-textcolor')?.getAttribute('data-color') || null,
    highlight:     el.closest('.md-highlight')?.getAttribute('data-color')  || null,
  };

  // Indiquer visuellement que le style est disponible
  ['ctx-paste-style', 'table-ctx-paste-style'].forEach(id => {
    const btn = document.getElementById(id);
    if (btn) btn.classList.add('ctx-style-ready');
  });
  hideContextMenu();
}

function ctxPasteStyle() {
  if (!_copiedStyle || state.sourceMode) { hideContextMenu(); return; }
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0 || sel.isCollapsed) { hideContextMenu(); return; }

  const node = sel.getRangeAt(0).commonAncestorContainer;
  const el = node.nodeType === Node.TEXT_NODE ? node.parentElement : node;

  // Appliquer uniquement les formats absents (évite le toggle involontaire)
  if (_copiedStyle.bold          && !document.queryCommandState('bold'))          applyBold();
  if (_copiedStyle.italic        && !document.queryCommandState('italic'))        applyItalic();
  if (_copiedStyle.underline     && !document.queryCommandState('underline'))     applyUnderline();
  if (_copiedStyle.strikethrough && !el.closest('s, del'))                        applyStrikethrough();
  if (_copiedStyle.smallcaps     && !el.closest('.smallcaps'))                    applySmallCaps();
  if (_copiedStyle.textColor     && !el.closest('.md-textcolor'))                 applyTextColor(_copiedStyle.textColor);
  if (_copiedStyle.highlight     && !el.closest('.md-highlight'))                 applyHighlight(_copiedStyle.highlight);

  syncPreviewToContent();
  hideContextMenu();
}

function ctxToggleCase() {
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const s = ta.selectionStart, e = ta.selectionEnd;
    if (s === e) { hideContextMenu(); return; }
    const selected = ta.value.substring(s, e);
    const toggled = selected === selected.toUpperCase() ? selected.toLowerCase() : selected.toUpperCase();
    ta.setRangeText(toggled, s, e, 'select');
    onSourceInput();
  } else {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) { hideContextMenu(); return; }
    const selected = sel.toString();
    const toggled = selected === selected.toUpperCase() ? selected.toLowerCase() : selected.toUpperCase();
    document.execCommand('insertText', false, toggled);
    syncPreviewToContent();
  }
  hideContextMenu();
}

function hideTableContextMenu() {
  document.getElementById('table-context-menu').classList.remove('visible');
}

async function copyTableToClipboard() {
  if (!_tableContextCell) return;
  const table = _tableContextCell.closest('table');
  if (!table) return;

  const rows = Array.from(table.rows);
  const text = rows.map(r =>
    Array.from(r.cells).map(c => c.innerText.trim()).join('\t')
  ).join('\n');

  try {
    await navigator.clipboard.write([new ClipboardItem({
      'text/plain': new Blob([text], { type: 'text/plain' }),
      'text/html':  new Blob([table.outerHTML], { type: 'text/html' }),
    })]);
  } catch {
    await navigator.clipboard.writeText(text);
  }
  showToast('Tableau copié dans le presse-papiers');
  hideTableContextMenu();
}

/**
 * Normalise les cellules d'un tableau dont le DOM contient des listes imbriquées
 * artefacts (produites par "- - - texte" → N niveaux de <ul><li> vides).
 * Le navigateur (Chromium) ajoute des <br> dans les <li> vides, ce qui corrompt
 * le cycle DOM→Markdown. On reconstruit le texte "- - … texte" échappé.
 * Les listes intentionnellement imbriquées (li avec du texte direct) sont préservées.
 */
function _normalizeTableCells(table) {
  const selector = 'td, th';
  const cells = table
    ? table.querySelectorAll(selector)
    : document.getElementById('preview').querySelectorAll(selector);

  cells.forEach(cell => {
    // Agir seulement si la cellule a des listes imbriquées (li > ul/ol)
    const nestedList = cell.querySelector('li > ul, li > ol');
    if (!nestedList) return;

    // Distinguer les listes intentionnellement imbriquées des artefacts
    // "- - - texte" : dans ce cas le <li> parent n'a PAS de texte direct,
    // il ne contient qu'une sous-liste. Une vraie liste imbriquée a du texte
    // dans le <li> parent.
    const outerLi = nestedList.closest('li');
    const hasDirectText = outerLi && [...outerLi.childNodes].some(
      n => n.nodeType === Node.TEXT_NODE && n.textContent.trim()
    );
    if (hasDirectText) return;

    // Compter le niveau d'imbrication (chaque niveau = un "- " dans la source)
    let depth = 0;
    let el = cell.querySelector('li');
    while (el) { depth++; el = el.querySelector('li'); }

    // Trouver le <li> le plus profond pour récupérer le contenu textuel
    let deepLi = cell.querySelector('li');
    while (deepLi && deepLi.querySelector('li')) deepLi = deepLi.querySelector('li');
    const text = deepLi ? deepLi.textContent.trim() : cell.textContent.trim();

    // Reconstruire "- - - … texte" puis échapper le premier tiret (anti-liste)
    const dashes = Array(depth).fill('- ').join('');
    const md = (dashes + text).replace(/^(\s*)([-*+])(\s|$)/, '$1\\$2$3');
    cell.innerHTML = marked.parse(md);
  });
}

function _syncAndRerender() {
  // Normaliser les cellules avec listes imbriquées AVANT la sync DOM→Markdown
  // pour éviter que Chromium injecte des <br> dans les <li> qui se retrouveraient
  // comme balises littérales dans le Markdown.
  _normalizeTableCells();
  syncPreviewToContent();
  const tab = getActiveTab();
  if (tab) updatePreview(tab.content);
}

function tableAddRow(where, count = 1) {
  if (!_tableContextCell) return;
  const row     = _tableContextCell.closest('tr');
  const numCols = row.cells.length;
  for (let r = 0; r < count; r++) {
    const newRow = document.createElement('tr');
    for (let i = 0; i < numCols; i++) newRow.appendChild(document.createElement('td'));
    if (where === 'above') row.parentNode.insertBefore(newRow, row);
    else                   row.parentNode.insertBefore(newRow, row.nextSibling);
  }
  _syncAndRerender();
}

function tableAddRowPrompt(where) {
  const label = where === 'above' ? 'au-dessus' : 'en-dessous';
  const raw   = prompt(`Nombre de lignes à insérer ${label} :`, '1');
  if (raw === null) return;
  const count = parseInt(raw, 10);
  if (!Number.isFinite(count) || count < 1) return;
  tableAddRow(where, count);
}

function tableDeleteRow() {
  if (!_tableContextCell) return;
  const row   = _tableContextCell.closest('tr');
  const tbody = row.closest('tbody');
  if (!tbody || tbody.rows.length <= 1) return; // garder au moins 1 ligne
  row.remove();
  _tableContextCell = null;
  _syncAndRerender();
}

function tableAddCol(where) {
  if (!_tableContextCell) return;
  const colIdx = _tableContextCell.cellIndex;
  const table  = _tableContextCell.closest('table');
  table.querySelectorAll('tr').forEach(tr => {
    const inHead = !!tr.closest('thead');
    const cell   = document.createElement(inHead ? 'th' : 'td');
    const ref    = where === 'left' ? tr.cells[colIdx] : (tr.cells[colIdx + 1] || null);
    tr.insertBefore(cell, ref);
  });
  _syncAndRerender();
}

function _getTableSizeWrapper(table) {
  const parent = table.parentElement;
  if (!parent?.classList.contains('fenced-div-content')) return null;
  const wrapper = parent.parentElement;
  if (!wrapper?.classList.contains('fenced-div-wrapper')) return null;
  // Ne concerne que les wrappers de mise en page : sans classe ni format
  if (wrapper.getAttribute('data-div-class') || wrapper.getAttribute('data-div-format')) return null;
  return wrapper;
}

function _colWidthUpdateTotal() {
  const sliders = [...document.querySelectorAll('#col-width-body .cw-slider')];
  const total   = sliders.reduce((s, r) => s + parseInt(r.value || 0), 0);
  const el      = document.getElementById('cw-total');
  if (!el) return;
  el.textContent = total + ' %';
  el.style.color = Math.abs(total - 100) <= 1 ? 'var(--green)' : total > 100 ? 'var(--red)' : 'var(--accent)';
}

function openColWidthDialog() {
  if (!_tableContextCell) return;
  const table   = _tableContextCell.closest('table');
  if (!table) return;
  const numCols = table.rows[0]?.cells.length || 0;

  // Lire les largeurs depuis data-col-widths (valeurs conteneur) ; sinon répartition égale
  const storedWidths = (table.getAttribute('data-col-widths') || '')
    .split(',').map(w => parseInt(w)).filter(n => !isNaN(n) && n > 0);
  const defaultPct = Math.round(100 / numCols);
  const body = document.getElementById('col-width-body');
  body.innerHTML = '';

  for (let i = 0; i < numCols; i++) {
    const pct = storedWidths[i] || (i < numCols - 1 ? defaultPct : 100 - defaultPct * (numCols - 1));
    const safe = Math.max(1, Math.min(100, pct));
    const row  = document.createElement('div');
    row.className = 'cw-row';
    row.innerHTML = `
      <span class="cw-col-badge">${i + 1}</span>
      <div class="cw-group">
        <input class="cw-slider" type="range" min="1" max="100" value="${safe}" data-col="${i}"
               oninput="this.nextElementSibling.value=this.value; _colWidthUpdateTotal()">
        <input class="cw-number" type="number" min="1" max="100" value="${safe}"
               oninput="this.previousElementSibling.value=this.value; _colWidthUpdateTotal()">
        <span class="cw-unit">%</span>
      </div>`;
    body.appendChild(row);
  }

  const footer = document.createElement('div');
  footer.className = 'cw-total-row';
  footer.innerHTML = `<span class="cw-total-label">Total</span><span id="cw-total"></span>`;
  body.appendChild(footer);

  // Préremplir la légende depuis l'élément <caption> du tableau
  const captionEl = table.querySelector('caption');
  document.getElementById('cw-caption').value = captionEl ? captionEl.textContent.trim() : '';

  document.getElementById('cw-bordered').checked = table.getAttribute('data-bordered') === '1';
  document.getElementById('cw-striped').checked  = table.getAttribute('data-striped')  === '1';

  // Détecter un wrapper fenced-div de mise en page (style: width + margin, sans class ni format)
  const sizeWrap = _getTableSizeWrapper(table);
  const tableWidth  = sizeWrap ? (parseInt(sizeWrap.getAttribute('data-div-width')) || 100) : 100;
  const tableCenter = sizeWrap ? sizeWrap.getAttribute('data-div-margin') === '1' : false;
  document.getElementById('cw-table-width-slider').value  = tableWidth;
  document.getElementById('cw-table-width-number').value  = tableWidth;
  document.getElementById('cw-table-center').checked      = tableCenter;

  document.getElementById('col-width-overlay').classList.add('open');
  _colWidthUpdateTotal();
}

function closeColWidthDialog(e) {
  if (e && e.target !== document.getElementById('col-width-overlay')) return;
  document.getElementById('col-width-overlay').classList.remove('open');
}

function confirmColWidths() {
  if (!_tableContextCell) return;
  const table = _tableContextCell.closest('table');
  if (!table) return;

  const sliders = [...document.querySelectorAll('#col-width-body .cw-slider')];
  const widths  = sliders.map(r => parseInt(r.value, 10));

  // Stocker les largeurs en data-col-widths pour que Turndown génère {tbl-colwidths="[...]"}
  table.setAttribute('data-col-widths', widths.join(','));

  const isBordered = document.getElementById('cw-bordered').checked;
  const isStriped  = document.getElementById('cw-striped').checked;
  table.classList.toggle('bordered', isBordered);
  table.classList.toggle('striped',  isStriped);
  isBordered ? table.setAttribute('data-bordered', '1') : table.removeAttribute('data-bordered');
  isStriped  ? table.setAttribute('data-striped',  '1') : table.removeAttribute('data-striped');

  // Mettre à jour l'élément <caption> (doit être le 1er enfant de <table>)
  const existingCaption = table.querySelector('caption');
  if (existingCaption) existingCaption.remove();
  const captionVal = document.getElementById('cw-caption').value.trim();
  if (captionVal) {
    const cap = document.createElement('caption');
    cap.textContent = captionVal;
    table.insertBefore(cap, table.firstChild);
  }

  // Wrapper fenced-div : largeur + centrage du tableau
  const tableWidth  = Math.max(10, Math.min(100, parseInt(document.getElementById('cw-table-width-number').value) || 100));
  const tableCenter = document.getElementById('cw-table-center').checked;
  const needsWrap   = tableWidth < 100 || tableCenter;
  const sizeWrap    = _getTableSizeWrapper(table);

  if (!needsWrap) {
    if (sizeWrap) sizeWrap.replaceWith(table);
  } else if (sizeWrap) {
    if (tableWidth < 100) sizeWrap.setAttribute('data-div-width', tableWidth);
    else sizeWrap.removeAttribute('data-div-width');
    if (tableCenter) sizeWrap.setAttribute('data-div-margin', '1');
    else sizeWrap.removeAttribute('data-div-margin');
    sizeWrap.style.width  = tableWidth < 100 ? tableWidth + '%' : '';
    sizeWrap.style.margin = tableCenter ? 'auto' : '';
  } else {
    const wrapper = document.createElement('div');
    wrapper.className = 'fenced-div-wrapper';
    if (tableWidth < 100) { wrapper.setAttribute('data-div-width', tableWidth); wrapper.style.width = tableWidth + '%'; }
    if (tableCenter)      { wrapper.setAttribute('data-div-margin', '1');        wrapper.style.margin = 'auto'; }
    const label   = document.createElement('span');
    label.className = 'fenced-div-label';
    label.contentEditable = 'false';
    label.textContent = 'DV';
    const content = document.createElement('div');
    content.className = 'fenced-div-content';
    wrapper.appendChild(label);
    wrapper.appendChild(content);
    table.replaceWith(wrapper);
    content.appendChild(table);
  }

  document.getElementById('col-width-overlay').classList.remove('open');

  // 1) DOM → markdown : Turndown lit data-col-widths et génère {tbl-colwidths="[...]"}
  // 2) markdown → DOM : l'extension marked recrée <colgroup> + table-layout:fixed proprement
  const tab = getActiveTab();
  if (!tab) return;
  const preview = document.getElementById('preview');
  tab.content = turndown.turndown(preview.innerHTML);
  tab.modified = (tab.content !== tab.savedContent);
  renderTabList();
  updatePreview(tab.content);
}

function _doToggleTightList(list) {
  if (!list) return;
  const isLoose = !!list.querySelector('li > p');
  if (isLoose) {
    list.querySelectorAll('li').forEach(li => {
      const p = li.querySelector(':scope > p');
      if (p) {
        while (p.firstChild) li.insertBefore(p.firstChild, p);
        p.remove();
      }
    });
  } else {
    list.querySelectorAll('li').forEach(li => {
      const p = document.createElement('p');
      [...li.childNodes].forEach(child => {
        if (child.nodeName !== 'UL' && child.nodeName !== 'OL')
          p.appendChild(child);
      });
      li.insertBefore(p, li.firstChild);
    });
  }
  syncPreviewToContent();
}

function ctxToggleTightList() {
  _doToggleTightList(_ctxTightList);
}

function toggleTightList() {
  if (state.sourceMode) {
    const ta   = document.getElementById('source-editor');
    const pos  = ta.selectionStart;
    const all  = ta.value.split('\n');
    const cur  = ta.value.substring(0, pos).split('\n').length - 1;

    const isListLine = l => /^\s*([-*+]|\d+\.)\s/.test(l);
    const isBlank    = l => l.trim() === '';

    // Trouver le bloc de liste autour du curseur
    let s = cur;
    while (s > 0 && (isListLine(all[s - 1]) || isBlank(all[s - 1]))) s--;
    let e = cur;
    while (e < all.length - 1 && (isListLine(all[e + 1]) || isBlank(all[e + 1]))) e++;

    const block = all.slice(s, e + 1);
    if (!block.some(isListLine)) return; // pas dans une liste

    const isLoose = block.some(isBlank);
    let replaced;
    if (isLoose) {
      // Serré : supprimer les lignes vides
      replaced = block.filter(l => !isBlank(l));
    } else {
      // Espacé : insérer une ligne vide entre chaque item
      replaced = [];
      block.forEach((l, i) => {
        replaced.push(l);
        if (i < block.length - 1) replaced.push('');
      });
    }

    all.splice(s, e - s + 1, ...replaced);
    ta.value = ta.value = all.join('\n');
    ta.selectionStart = ta.selectionEnd = pos;
    ta.focus();
    onSourceInput();
  } else {
    // Mode prévisualisation : trouver la liste sous le curseur
    const preview = document.getElementById('preview');
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    let node = sel.anchorNode;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
    const list = node.closest('ul, ol');
    if (!list || !preview.contains(list)) return;
    _doToggleTightList(list);
  }
}

function openListStartDialog() {
  let current = 1;
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const lines = ta.value.split('\n');
    const line = lines[_ctxOrderedListSourceLine] || '';
    const m = line.match(/^\s*(\d+)\. /);
    if (m) current = parseInt(m[1]);
  } else if (_ctxOrderedList) {
    current = parseInt(_ctxOrderedList.getAttribute('start') || '1');
  }
  const input = document.getElementById('list-start-input');
  input.value = current;
  document.getElementById('list-start-overlay').classList.add('open');
  setTimeout(() => { input.focus(); input.select(); }, 50);
}

function closeListStartDialog(e) {
  if (e && e.target !== document.getElementById('list-start-overlay')) return;
  document.getElementById('list-start-overlay').classList.remove('open');
}

function applyListStart() {
  const n = parseInt(document.getElementById('list-start-input').value);
  if (isNaN(n) || n < 0) return;
  document.getElementById('list-start-overlay').classList.remove('open');

  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const lines = ta.value.split('\n');
    const idx = _ctxOrderedListSourceLine;
    if (idx < 0 || idx >= lines.length) return;
    lines[idx] = lines[idx].replace(/^(\s*)\d+(\. )/, `$1${n}$2`);
    const cursor = ta.selectionStart;
    ta.value = lines.join('\n');
    ta.selectionStart = ta.selectionEnd = cursor;
    ta.focus();
    onSourceInput();
  } else if (_ctxOrderedList) {
    if (n === 1) {
      _ctxOrderedList.removeAttribute('start');
    } else {
      _ctxOrderedList.setAttribute('start', String(n));
    }
    syncPreviewToContent();
  }
}

function tableToggleTightList() {
  if (!_tableContextCell) return;
  _doToggleTightList(_tableContextCell.querySelector('ul, ol'));
}

function tableInsertLink() {
  if (_tableLinkRange) {
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(_tableLinkRange);
    _tableLinkRange = null;
  }
  openLinkDialog();
}

function tableDelete() {
  if (!_tableContextCell) return;
  const table = _tableContextCell.closest('table');
  if (!table) return;
  const wrapper = table.closest('.code-block-wrapper') || table;
  wrapper.remove();
  _tableContextCell = null;
  _syncAndRerender();
}

function tableDeleteCol() {
  if (!_tableContextCell) return;
  const colIdx = _tableContextCell.cellIndex;
  const table  = _tableContextCell.closest('table');
  const numCols = table.rows[0]?.cells.length || 0;
  if (numCols <= 1) return; // garder au moins 1 colonne
  table.querySelectorAll('tr').forEach(tr => {
    if (tr.cells[colIdx]) tr.deleteCell(colIdx);
  });
  _tableContextCell = null;
  _syncAndRerender();
}

function alignColumn(align) {
  if (!_tableContextCell) return;
  const cell  = _tableContextCell;
  const table = cell.closest('table');
  if (!table) return;
  const colIdx = cell.cellIndex;

  // Appliquer l'alignement sur toutes les cellules de la colonne
  table.querySelectorAll('tr').forEach(tr => {
    const c = tr.cells[colIdx];
    if (!c) return;
    c.style.textAlign = align;
    c.setAttribute('align', align); // Redondance pour la sérialisation
    c.querySelectorAll('p').forEach(p => { 
      p.style.textAlign = align; 
      p.setAttribute('align', align);
    });
  });

  syncPreviewToContent();
  const tab = getActiveTab();
  if (tab) updatePreview(tab.content);
}

function alignColumnVertical(align) {
  if (!_tableContextCell) return;
  const cell  = _tableContextCell;
  const table = cell.closest('table');
  if (!table) return;
  const colIdx = cell.cellIndex;

  // Appliquer l'alignement vertical sur toutes les cellules de la colonne
  table.querySelectorAll('tr').forEach(tr => {
    const c = tr.cells[colIdx];
    if (!c) return;
    c.style.verticalAlign = align;
    c.setAttribute('valign', align); // Redondance
  });

  syncPreviewToContent();
  const tab = getActiveTab();
  if (tab) updatePreview(tab.content);
}

// ── Image ─────────────────────────────────────────────────────────────────────
let _imageContextFigure = null;
let _pastedImagePendingFigure = null; // figure insérée par paste, à retirer si l'utilisateur annule

function _blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

async function _handlePasteImage(blob) {
  if (!window.pywebview) return;
  const tab = getActiveTab();
  const ext = (blob.type.split('/')[1] || 'png').replace('jpeg', 'jpg');
  const initialDir = (tab?.path || '').replace(/[/\\][^/\\]+$/, '');

  const base64 = await _blobToBase64(blob);
  const result = await window.pywebview.api.save_pasted_image(base64, ext, initialDir);
  if (!result || !result.path) return;

  const relPath = _makeRelativePath(tab?.path || '', result.path);
  const name    = result.name.replace(/\.[^.]+$/, '');
  const md      = `![${name}](${relPath}){fig-align="center"}`;

  if (state.sourceMode) {
    const editor = document.getElementById('source-editor');
    const pos    = editor.selectionStart;
    editor.value = editor.value.slice(0, pos) + '\n' + md + '\n' + editor.value.slice(pos);
    if (tab) { tab.content = editor.value; tab.modified = true; }
    renderTabList();
  } else {
    const temp = document.createElement('div');
    temp.innerHTML = marked.parse(md);
    const figEl = temp.querySelector('figure');
    if (!figEl) return;

    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      const range = sel.getRangeAt(0);
      range.collapse(true);
      range.insertNode(figEl);
      range.setStartAfter(figEl);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    } else {
      document.getElementById('preview').appendChild(figEl);
    }

    _pastedImagePendingFigure = figEl;
    _imageContextFigure = figEl;
    openImageDialog();
  }
}

function _setupPasteImageListener() {
  document.addEventListener('paste', async e => {
    if (!e.clipboardData?.items) return;
    const imageItem = Array.from(e.clipboardData.items).find(it => it.type.startsWith('image/'));
    if (!imageItem) return;
    // Laisser passer si le focus est dans un champ de saisie (hors éditeur source)
    const active = document.activeElement;
    const tag    = active?.tagName;
    if (['INPUT', 'SELECT'].includes(tag)) return;
    if (tag === 'TEXTAREA' && active.id !== 'source-editor') return;
    e.preventDefault();
    const blob = imageItem.getAsFile();
    if (blob) await _handlePasteImage(blob);
  });
}

async function applyImage() {
  const result = await window.pywebview.api.open_image();
  if (!result || !result.path) return;
  const tab = getActiveTab();
  if (!tab) return;
  const relPath = _makeRelativePath(tab.path || '', result.path);
  const name = result.name.replace(/\.[^.]+$/, '');
  const md = `![${name}](${relPath}){fig-align="center"}`;
  if (state.sourceMode) {
    const editor = document.getElementById('source-editor');
    const pos = editor.selectionStart;
    editor.value = editor.value.slice(0, pos) + '\n' + md + '\n' + editor.value.slice(pos);
    tab.content = editor.value;
    tab.modified = true;
    renderTabList();
  } else {
    const temp = document.createElement('div');
    temp.innerHTML = marked.parse(md);
    const figEl = temp.querySelector('figure');
    if (!figEl) return;
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      const range = sel.getRangeAt(0);
      range.collapse(true);
      range.insertNode(figEl);
      range.setStartAfter(figEl);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
    } else {
      document.getElementById('preview').appendChild(figEl);
    }
    syncPreviewToContent();
    updatePreview(tab.content);
  }
}

function hideImageContextMenu() {
  document.getElementById('img-context-menu').style.display = 'none';
}

// Convertit les enfants d'une figcaption en markdown (texte + liens)
function _figcapToMd(figcap) {
  let md = '';
  figcap.childNodes.forEach(node => {
    if (node.nodeType === Node.TEXT_NODE) {
      md += node.textContent;
    } else if (node.nodeName === 'A') {
      md += `[${node.textContent}](${node.getAttribute('href') || ''})`;
    } else {
      md += node.textContent;
    }
  });
  return md.trim();
}

function insertCaptionLink() {
  const input = document.getElementById('img-caption-input');
  const start = input.selectionStart;
  const end   = input.selectionEnd;
  const sel   = input.value.slice(start, end).trim();
  const ins   = sel ? `[${sel}](url)` : '[texte](url)';
  input.value = input.value.slice(0, start) + ins + input.value.slice(end);
  const uStart = input.value.indexOf('(', start) + 1;
  const uEnd   = input.value.indexOf(')', uStart);
  input.focus();
  input.setSelectionRange(uStart, uEnd);
}

function openImageDialog() {
  if (!_imageContextFigure) return;
  const currentSrc = _imageContextFigure.getAttribute('data-src') || '';
  document.getElementById('img-file-name').textContent = currentSrc || '—';
  document.getElementById('img-file-name').title = currentSrc || '';
  const figcap = _imageContextFigure.querySelector('figcaption');
  document.getElementById('img-caption-input').value = figcap ? _figcapToMd(figcap) : '';
  const align = _imageContextFigure.getAttribute('data-align') || 'center';
  document.querySelectorAll('.img-align-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.align === align);
  });
  const imgW = parseInt(_imageContextFigure.getAttribute('data-imgwidth') || '100');
  document.getElementById('img-width-slider').value = imgW;
  document.getElementById('img-width-number').value = imgW;

  // Détection d'un wrapper content-visible existant
  const fdivContent = _imageContextFigure.parentElement;
  const fdivWrapper = fdivContent?.classList.contains('fenced-div-content') ? fdivContent.parentElement : null;
  const isContentVisible = fdivWrapper?.classList.contains('fenced-div-wrapper')
    && (fdivWrapper.getAttribute('data-div-class') || '').includes('content-visible');
  let vis = 'always';
  if (isContentVisible) {
    const fmt = fdivWrapper.getAttribute('data-div-format') || '';
    vis = fmt === 'html' ? 'html' : fmt === 'pdf' ? 'pdf' : 'always';
  }
  document.querySelectorAll('.img-align-btn[data-vis]').forEach(b => {
    b.classList.toggle('active', b.dataset.vis === vis);
  });

  document.getElementById('img-overlay').classList.add('open');
}

function closeImageDialog(e) {
  if (e && e.target !== document.getElementById('img-overlay')) return;
  document.getElementById('img-overlay').classList.remove('open');
  // Si on annule après un collage, retirer la figure temporaire
  if (_pastedImagePendingFigure) {
    _pastedImagePendingFigure.remove();
    _pastedImagePendingFigure = null;
    syncPreviewToContent();
  }
}

function selectImageAlign(align) {
  document.querySelectorAll('.img-align-btn[data-align]').forEach(b => {
    b.classList.toggle('active', b.dataset.align === align);
  });
}

function selectImageVisibility(vis) {
  document.querySelectorAll('.img-align-btn[data-vis]').forEach(b => {
    b.classList.toggle('active', b.dataset.vis === vis);
  });
}

function confirmImageDialog() {
  if (!_imageContextFigure) return;
  const caption = document.getElementById('img-caption-input').value.trim();
  const activeBtn = document.querySelector('.img-align-btn.active');
  const align = activeBtn ? activeBtn.dataset.align : 'center';
  const imgW  = Math.max(5, Math.min(100, parseInt(document.getElementById('img-width-number').value) || 100));

  // Mettre à jour l'alignement et la taille
  _imageContextFigure.className = `md-figure fig-${align}`;
  _imageContextFigure.setAttribute('data-align', align);
  _imageContextFigure.setAttribute('data-imgwidth', imgW);
  const img = _imageContextFigure.querySelector('img');
  if (img) img.style.width = imgW < 100 ? imgW + '%' : '';

  // Mettre à jour la légende (supporte texte libre + liens markdown [texte](url))
  let figcap = _imageContextFigure.querySelector('figcaption');
  if (caption) {
    if (!figcap) {
      figcap = document.createElement('figcaption');
      _imageContextFigure.appendChild(figcap);
    }
    figcap.innerHTML = marked.parseInline(caption);
    const altText = caption.replace(/\[([^\]]*)\]\([^)]*\)/g, '$1');
    if (img) img.setAttribute('alt', altText);
  } else if (figcap) {
    figcap.remove();
  }

  // Visibilité : wrapper fenced-div content-visible
  const activeVisBtn = document.querySelector('.img-align-btn[data-vis].active');
  const vis = activeVisBtn ? activeVisBtn.dataset.vis : 'always';

  const fdivContent  = _imageContextFigure.parentElement;
  const fdivWrapper  = fdivContent?.classList.contains('fenced-div-content') ? fdivContent.parentElement : null;
  const existingWrap = fdivWrapper?.classList.contains('fenced-div-wrapper')
    && (fdivWrapper.getAttribute('data-div-class') || '').includes('content-visible')
    ? fdivWrapper : null;

  if (vis === 'always') {
    if (existingWrap) {
      // Déplacer la figure hors du wrapper, supprimer le wrapper
      existingWrap.replaceWith(_imageContextFigure);
    }
  } else {
    if (existingWrap) {
      existingWrap.setAttribute('data-div-format', vis);
    } else {
      // Créer un nouveau wrapper autour de la figure
      const wrapper = document.createElement('div');
      wrapper.className = 'fenced-div-wrapper content-visible';
      wrapper.setAttribute('data-div-class', 'content-visible');
      wrapper.setAttribute('data-div-format', vis);
      const label = document.createElement('span');
      label.className = 'fenced-div-label';
      label.contentEditable = 'false';
      label.textContent = 'DV';
      const content = document.createElement('div');
      content.className = 'fenced-div-content';
      wrapper.appendChild(label);
      wrapper.appendChild(content);
      _imageContextFigure.replaceWith(wrapper);
      content.appendChild(_imageContextFigure);
    }
  }

  _pastedImagePendingFigure = null;
  document.getElementById('img-overlay').classList.remove('open');
  syncPreviewToContent();
}

async function changeImageFile() {
  const result = await window.pywebview.api.open_image();
  if (!result || !result.path) return;
  if (!_imageContextFigure) return;
  const tab = getActiveTab();
  const relPath = _makeRelativePath(tab && tab.path ? tab.path : '', result.path);
  const img = _imageContextFigure.querySelector('img');
  if (img) img.src = _pathToUrl(relPath);
  _imageContextFigure.setAttribute('data-src', relPath);
  document.getElementById('img-file-name').textContent = relPath;
  document.getElementById('img-file-name').title = relPath;
}

function copyImagePath() {
  const path = document.getElementById('img-file-name').textContent;
  if (!path || path === '—') return;
  navigator.clipboard.writeText(path).then(() => {
    const btn = document.querySelector('#img-dialog .settings-btn[onclick="copyImagePath()"]');
    if (!btn) return;
    const orig = btn.textContent;
    btn.textContent = '✓ Copié';
    setTimeout(() => { btn.textContent = orig; }, 1500);
  });
}

function deleteImage() {
  if (!_imageContextFigure) return;
  _imageContextFigure.remove();
  _imageContextFigure = null;
  syncPreviewToContent();
}

// Recharge l'image depuis le disque (contourne le cache du navigateur) : utile
// quand le fichier a été modifié en dehors de l'éditeur (ex. régénéré par un
// autre outil) sans que son chemin ait changé.
function refreshImage() {
  if (!_imageContextFigure) return;
  const img = _imageContextFigure.querySelector('img');
  const src = _imageContextFigure.getAttribute('data-src') || '';
  if (!img || !src) return;
  img.src = _pathToUrl(src) + '?t=' + Date.now();
}

// Convertit un blob image (png/jpg/gif/webp/svg…) en PNG via un <canvas>, en passant
// par une URL blob: (même origine, donc pas de canvas « tainted ») plutôt que file:/http:.
function _rasterizeToPng(blob) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width  = img.naturalWidth;
      canvas.height = img.naturalHeight;
      canvas.getContext('2d').drawImage(img, 0, 0);
      canvas.toBlob(pngBlob => {
        URL.revokeObjectURL(url);
        pngBlob ? resolve(pngBlob) : reject(new Error('toBlob a échoué'));
      }, 'image/png');
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("Chargement de l'image échoué")); };
    img.src = url;
  });
}

// Copie l'image (fichier local ou distant) dans le presse-papiers système, au format
// PNG (seul format universellement supporté par l'API Clipboard). Les images locales
// sont lues via l'API Python (et non fetch()/canvas sur une URL file:, non fiable
// selon le moteur webview) puis converties en PNG côté JS.
async function copyImage() {
  if (!_imageContextFigure) return;
  const rawSrc = _imageContextFigure.getAttribute('data-src') || '';
  if (!rawSrc) return;

  try {
    let blob;
    if (/^https?:\/\//i.test(rawSrc)) {
      const resp = await fetch(rawSrc);
      blob = await resp.blob();
    } else {
      if (!window.pywebview || !window.pywebview.api) throw new Error('API indisponible');
      const fileUrl = _pathToUrl(rawSrc);
      const absPath = decodeURIComponent(fileUrl.replace(/^file:\/\/\//, ''));
      const result  = await window.pywebview.api.read_image_base64(absPath);
      if (!result || result.error) throw new Error(result && result.error);
      const bytes = Uint8Array.from(atob(result.data), c => c.charCodeAt(0));
      blob = new Blob([bytes], { type: result.mime || 'application/octet-stream' });
    }
    const pngBlob = await _rasterizeToPng(blob);
    await navigator.clipboard.write([new ClipboardItem({ 'image/png': pngBlob })]);
    showToast('Image copiée dans le presse-papiers');
  } catch (e) {
    console.error('copyImage:', e);
    showToast("Impossible de copier l'image");
  }
}

// ── Redimensionnement colonnes tableau par glisser-déposer ────────────────────
// `widths` : pourcentages relatifs à la largeur de la TABLE (somme = 100).
// La largeur de la table par rapport au conteneur (#preview) n'est PAS modifiée
// par un redimensionnement de colonne : seule la répartition interne change.
function _applyTableColWidths(table, widths) {
  const sumPct = widths.reduce((s, w) => s + w, 0);
  const colPcts = widths.map(w => (w / sumPct) * 100);

  let colgroup = table.querySelector('colgroup');
  if (!colgroup) {
    colgroup = document.createElement('colgroup');
    table.insertBefore(colgroup, table.firstChild);
  }
  colgroup.innerHTML = '';
  colPcts.forEach(w => {
    const col = document.createElement('col');
    col.style.width = w.toFixed(3) + '%';
    colgroup.appendChild(col);
  });

  table.style.tableLayout = 'fixed';

  // data-col-widths / tbl-colwidths sont exprimés en % du conteneur : on les
  // recalcule à partir des % de colonnes (somme 100) en conservant la largeur
  // actuelle de la table par rapport au conteneur.
  const currentTableWidthPct = parseFloat(table.style.width) || 100;
  const containerPcts = colPcts.map(w => (w / 100) * currentTableWidthPct);
  const rounded = containerPcts.map(w => Math.round(w));
  table.setAttribute('data-col-widths', rounded.join(','));

  [...table.querySelectorAll('th')].forEach((th, i) => {
    if (i < rounded.length) th.setAttribute('data-width', rounded[i] + '%');
    else th.removeAttribute('data-width');
  });
}

function setupTableColResize() {
  const preview = document.getElementById('preview');
  let resizing = null;

  function thAtRightBorder(e) {
    const th = e.target && e.target.closest && e.target.closest('#preview th');
    if (!th || !th.closest('.grid-table')) return null;
    const rect = th.getBoundingClientRect();
    return Math.abs(e.clientX - rect.right) <= 6 ? th : null;
  }

  // Curseur col-resize au survol du bord droit d'un <th>
  preview.addEventListener('mousemove', e => {
    if (resizing) return;
    preview.style.cursor = thAtRightBorder(e) ? 'col-resize' : '';
  });

  // Début du glisser
  preview.addEventListener('mousedown', e => {
    const th = thAtRightBorder(e);
    if (!th) return;
    const table = th.closest('table');
    const numCols = table.rows[0] ? table.rows[0].cells.length : 0;
    if (!numCols) return;
    const colIdx = th.cellIndex;

    e.preventDefault();
    e.stopPropagation();

    // Toutes les largeurs sont exprimées en % de la largeur de la TABLE
    // elle-même (et non du conteneur #preview) : par construction leur somme
    // vaut toujours 100, quel que soit le padding/zoom du conteneur.
    const tableWidth = table.getBoundingClientRect().width;
    const rawWidths = [...table.rows[0].cells].map(
      c => (c.getBoundingClientRect().width / tableWidth) * 100
    );
    const sumRaw = rawWidths.reduce((s, w) => s + w, 0);
    const startWidths = rawWidths.map(w => (w / sumRaw) * 100);

    // Verrouiller le layout pendant le glisser
    table.style.tableLayout = 'fixed';

    resizing = { table, colIdx, numCols, startX: e.clientX, startWidths, tableWidth };
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
  });

  // Glisser
  document.addEventListener('mousemove', e => {
    if (!resizing) return;
    const { table, colIdx, numCols, startX, startWidths, tableWidth } = resizing;
    let dPct = ((e.clientX - startX) / tableWidth) * 100;
    // Colonne voisine qui compense : la suivante, ou la précédente pour la dernière colonne
    const otherIdx = colIdx < numCols - 1 ? colIdx + 1 : colIdx - 1;
    // Bornes de dPct pour que les deux colonnes restent >= 2% — la somme totale
    // (et donc le total à 100%) reste ainsi toujours strictement constante.
    const minD = -(startWidths[colIdx]   - 2);
    const maxD =   startWidths[otherIdx] - 2;
    dPct = Math.max(minD, Math.min(maxD, dPct));
    const newWidths = [...startWidths];
    newWidths[colIdx]   = startWidths[colIdx]   + dPct;
    newWidths[otherIdx] = startWidths[otherIdx] - dPct;
    _applyTableColWidths(table, newWidths);
  });

  // Fin du glisser
  document.addEventListener('mouseup', () => {
    if (!resizing) return;
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
    resizing = null;
    syncPreviewToContent();
  });
}

// ── Correcteur orthographique intégré au menu contextuel ─────────────────────

function _isSpellCheckEnabled() {
  return !!(loadSettings().spellcheck);
}

function _cleanWord(raw) {
  return raw.replace(/^[«»"'.,;:!?()\[\]{}\-–—…\d]+|[«»"'.,;:!?()\[\]{}\-–—…\d]+$/gu, '');
}

// Extrait le mot sous le pointeur dans le preview (contenteditable)
function _getWordAtPoint(x, y) {
  let node, offset;
  if (document.caretRangeFromPoint) {
    const r = document.caretRangeFromPoint(x, y);
    if (!r) return null;
    node = r.startContainer; offset = r.startOffset;
  } else if (document.caretPositionFromPoint) {
    const p = document.caretPositionFromPoint(x, y);
    if (!p) return null;
    node = p.offsetNode; offset = p.offset;
  } else { return null; }

  if (!node || node.nodeType !== Node.TEXT_NODE) return null;
  const text = node.textContent;
  const isW  = c => /[a-zA-ZÀ-ÿ'-]/.test(c);

  let pos = offset;
  if (pos > 0 && !isW(text[pos]) && isW(text[pos - 1])) pos--;
  if (pos >= text.length || !isW(text[pos])) return null;

  let s = pos, e = pos + 1;
  while (s > 0 && isW(text[s - 1])) s--;
  while (e < text.length && isW(text[e])) e++;

  const word = text.substring(s, e);
  if (word.length < 2) return null;
  const range = document.createRange();
  range.setStart(node, s);
  range.setEnd(node, e);
  return { word, range };
}

// Extrait le mot à la position dans la textarea source
function _getWordAtSourcePos(pos) {
  const ta   = document.getElementById('source-editor');
  const text = ta.value;
  const isW  = c => /[a-zA-ZÀ-ÿ'-]/.test(c);

  let p = pos;
  if (p > 0 && !isW(text[p]) && isW(text[p - 1])) p--;
  if (p >= text.length || !isW(text[p])) return null;

  let s = p, e = p + 1;
  while (s > 0 && isW(text[s - 1])) s--;
  while (e < text.length && isW(text[e])) e++;

  const word = text.substring(s, e);
  return word.length < 2 ? null : { word, start: s, end: e };
}

// Remplit la section orthographe du menu contextuel et retourne true si mot erroné
async function _fillSpellSection(x, y) {
  const section = document.getElementById('ctx-spell-section');
  section.style.display = 'none';
  _spellCurrentWord = '';
  _spellCurrentRange = null;

  if (!window.pywebview) return false;

  // Extraire le mot selon le mode
  let hit = null;
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    hit = _getWordAtSourcePos(ta.selectionStart);
    if (hit) _spellCurrentRange = { start: hit.start, end: hit.end };
  } else {
    hit = _getWordAtPoint(x, y);
    if (hit) _spellCurrentRange = hit.range;
  }
  if (!hit) { console.log('[spell] no word at point', x, y); return false; }

  const clean = _cleanWord(hit.word);
  if (!clean || clean.length < 2) return false;
  if (_ignoredWordsSession.has(clean.toLowerCase())) return false;

  // Appel Python avec délai max de 1500 ms
  const lang = loadSettings().lang || 'fr';
  const timeout = new Promise(resolve => setTimeout(() => resolve(null), 1500));
  const check   = window.pywebview.api.get_spell_suggestions(clean, lang);
  const res = await Promise.race([check, timeout]);
  console.log('[spell]', clean, '->', res);
  if (!res || res.correct) return false;

  _spellCurrentWord = clean;
  document.getElementById('ctx-spell-word').textContent = '« ' + clean + ' »';

  const items = document.getElementById('ctx-spell-items');
  items.innerHTML = '';
  if (res.suggestions.length === 0) {
    const d = document.createElement('div');
    d.className = 'spell-no-suggestion';
    d.textContent = 'Aucune suggestion';
    items.appendChild(d);
  } else {
    res.suggestions.forEach(sug => {
      const btn = document.createElement('button');
      btn.className = 'ctx-item ctx-spell-correction';
      btn.textContent = sug;
      btn.onclick = () => { applySpellCorrection(sug); hideContextMenu(); };
      items.appendChild(btn);
    });
  }
  section.style.display = '';
  return true;
}

function applySpellCorrection(correction) {
  const savedRange = _spellCurrentRange;
  _spellCurrentRange = null;
  if (!savedRange) return;

  if (typeof savedRange === 'object' && 'start' in savedRange) {
    // Mode source
    const ta = document.getElementById('source-editor');
    ta.setRangeText(correction, savedRange.start, savedRange.end, 'end');
    ta.focus();
    onSourceInput();
  } else {
    // Mode preview — Range API (pas d'execCommand)
    try {
      savedRange.deleteContents();
      const tn = document.createTextNode(correction);
      savedRange.insertNode(tn);
      savedRange.setStartAfter(tn);
      savedRange.collapse(true);
      const sel = window.getSelection();
      if (sel) { sel.removeAllRanges(); sel.addRange(savedRange); }
    } catch (_) { /* DOM modifié, on ignore */ }
    syncPreviewToContent();
  }
}

async function ctxIgnoreSpellWord() {
  const word = _spellCurrentWord;
  hideContextMenu();
  if (!word) return;
  _ignoredWordsSession.add(word.toLowerCase());
  if (window.pywebview) await window.pywebview.api.ignore_word(word);
}

function setupContextMenu() {
  const menu      = document.getElementById('context-menu');
  const tableMenu = document.getElementById('table-context-menu');
  const imgMenu   = document.getElementById('img-context-menu');

  function hideAll() { hideContextMenu(); hideTableContextMenu(); hideImageContextMenu(); hideTabContextMenu(); hideSidebarDirContextMenu(); hideSidebarFileContextMenu(); hidePumlContextMenu(); }

  document.getElementById('main').addEventListener('contextmenu', async e => {
    e.preventDefault();
    hideAll();

    const figure  = e.target.closest('figure.md-figure');
    const cell    = e.target.closest('td, th');
    const heading = e.target.closest('#preview h1,#preview h2,#preview h3,#preview h4,#preview h5,#preview h6');

    // Mémoriser le titre pour l'option signet du menu contextuel
    _bookmarkHeadingNode = heading || null;
    const bookmarkItem = document.getElementById('ctx-bookmark');
    const bookmarkSep  = document.getElementById('ctx-bookmark-sep');
    if (bookmarkItem) bookmarkItem.style.display = heading ? '' : 'none';
    if (bookmarkSep)  bookmarkSep.style.display  = heading ? '' : 'none';

    const inFigcaption = !!e.target.closest('figcaption');

    if (figure && e.target.closest('#preview') && !inFigcaption) {
      // Clic sur une image (hors légende) → menu image
      _imageContextFigure = figure;
      _showMenuAt(imgMenu, e.clientX, e.clientY, m => m.style.display = 'block');
    } else if (cell && e.target.closest('#preview')) {
      // Clic dans une cellule de tableau → menu tableau
      _tableContextCell = cell;
      // Sauvegarder la sélection courante pour l'insertion de lien
      const _selForLink = window.getSelection();
      _tableLinkRange = (_selForLink && _selForLink.rangeCount > 0)
        ? _selForLink.getRangeAt(0).cloneRange()
        : null;
      // Mettre à jour le bouton "Liste serrée/espacée"
      const cellList = cell.querySelector('ul, ol');
      const tightBtn   = document.getElementById('table-ctx-tight-list');
      const tightLabel = document.getElementById('table-ctx-tight-label');
      if (tightBtn && tightLabel) {
        tightBtn.style.display = cellList ? '' : 'none';
        if (cellList) {
          const isLoose = !!cellList.querySelector('li > p');
          tightLabel.textContent = isLoose ? 'Liste serrée' : 'Liste espacée';
        }
      }
      _showMenuAt(tableMenu, e.clientX, e.clientY, m => m.classList.add('visible'));
    } else {
      // Sinon → menu contextuel normal
      // Sauvegarder la position d'insertion avant tout changement de focus.
      // e.target est toujours fiable (indépendant de la sélection).
      const _preview = document.getElementById('preview');
      _scInsertionAnchor = e.target && _preview.contains(e.target) ? e.target : null;
      const _selNow = window.getSelection();
      _scInsertionRange = (_selNow && _selNow.rangeCount > 0 && _preview.contains(_selNow.anchorNode))
        ? _selNow.getRangeAt(0).cloneRange()
        : (document.caretRangeFromPoint
            ? (() => {
                const r = document.caretRangeFromPoint(e.clientX, e.clientY);
                return (r && _preview.contains(r.startContainer)) ? r : null;
              })()
            : null);

      // Mettre à jour le bouton "Liste serrée/espacée"
      const ctxLi   = e.target.closest?.('li');
      _ctxTightList = ctxLi ? ctxLi.closest('ul, ol') : null;
      const ctxTightBtn   = document.getElementById('ctx-tight-list');
      const ctxTightLabel = document.getElementById('ctx-tight-label');
      if (ctxTightBtn && ctxTightLabel) {
        ctxTightBtn.style.display = _ctxTightList ? '' : 'none';
        if (_ctxTightList) {
          const isLoose = !!_ctxTightList.querySelector('li > p');
          ctxTightLabel.textContent = isLoose ? 'Liste serrée' : 'Liste espacée';
        }
      }

      // Mettre à jour le bouton "Numéro de départ"
      _ctxOrderedList = ctxLi ? ctxLi.closest('ol') : null;
      _ctxOrderedListSourceLine = -1;
      const ctxStartBtn = document.getElementById('ctx-list-start');
      if (ctxStartBtn) {
        let showStart = !!_ctxOrderedList;
        if (state.sourceMode) {
          const ta = document.getElementById('source-editor');
          const cursorLine = ta.value.substring(0, ta.selectionStart).split('\n').length - 1;
          const lines = ta.value.split('\n');
          if (/^\s*\d+\. /.test(lines[cursorLine] || '')) {
            showStart = true;
            // Remonter au premier item du bloc
            let first = cursorLine;
            while (first > 0 && /^\s*\d+\. /.test(lines[first - 1])) first--;
            _ctxOrderedListSourceLine = first;
          } else {
            showStart = false;
          }
        }
        ctxStartBtn.style.display = showStart ? '' : 'none';
      }

      await _fillSpellSection(e.clientX, e.clientY);
      _showMenuAt(menu, e.clientX, e.clientY, m => m.classList.add('visible'));
    }
  });

  document.addEventListener('click', hideAll);
  let _lastEscTime = 0;
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      hideAll();
      const now = Date.now();
      if (now - _lastEscTime < 400) { openHelpDialog(); _lastEscTime = 0; }
      else _lastEscTime = now;
    }
  });
  document.addEventListener('contextmenu', e => {
    if (!e.target.closest('#main')) hideAll();
  });
}

// ── Zoom molette ──────────────────────────────────────────────────────────────
document.addEventListener('wheel', e => {
  if (e.shiftKey && !e.ctrlKey && _isPreviewWidthActive()) {
    e.preventDefault();
    const step = e.deltaY < 0 ? 20 : -20;
    _applyContentMaxWidth(_contentMaxWidth + step);
    return;
  }
  if (!e.ctrlKey) return;
  e.preventDefault();
  const slider = document.getElementById('zoom-slider');
  const step = e.deltaY < 0 ? 5 : -5;
  const newVal = Math.min(200, Math.max(50, parseInt(slider.value) + step));
  slider.value = newVal;
  setZoom(newVal);
}, { passive: false });

// ── Recherche (Ctrl+F) ────────────────────────────────────────────────────────
let _searchMatches    = [];   // mode preview : tableau de <mark>
let _searchSrcMatches = [];   // mode source  : tableau de {start, end}
let _searchCurrent    = -1;
let _searchLastQuery  = '';   // dernière requête exécutée

function openSearch() {
  const bar = document.getElementById('search-bar');
  bar.classList.add('open');
  const input = document.getElementById('search-input');
  input.select();
  input.focus();
}

function closeSearch() {
  document.getElementById('search-bar').classList.remove('open');
  _clearSearchHighlights();
  document.getElementById('search-count').textContent = '';
  document.getElementById('search-input').classList.remove('search-no-match');
  _searchMatches    = [];
  _searchSrcMatches = [];
  _searchCurrent    = -1;
  _searchLastQuery  = '';
}

function onSearchInput() {
  // Réinitialise uniquement si le champ est vidé
  const query = document.getElementById('search-input').value;
  if (!query) {
    _clearSearchHighlights();
    document.getElementById('search-count').textContent = '';
    document.getElementById('search-input').classList.remove('search-no-match');
    _searchMatches    = [];
    _searchSrcMatches = [];
    _searchCurrent    = -1;
    _searchLastQuery  = '';
  }
}

function onSearchKeydown(e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    const query = document.getElementById('search-input').value;
    if (query !== _searchLastQuery) {
      // Nouvelle requête → lancer la recherche et aller au 1er résultat
      _searchLastQuery = query;
      _doSearch();
    } else {
      // Même requête → naviguer
      e.shiftKey ? searchPrev() : searchNext();
    }
  }
  if (e.key === 'Escape') { e.preventDefault(); closeSearch(); }
}

function _doSearch() {
  const query = document.getElementById('search-input').value;
  if (state.sourceMode) _doSearchSource(query);
  else                  _doSearchPreview(query);
}

/* ── Mode source ── */
function _doSearchSource(query) {
  _searchSrcMatches = [];
  const countEl = document.getElementById('search-count');
  const input   = document.getElementById('search-input');
  if (!query) { countEl.textContent = ''; input.classList.remove('search-no-match'); return; }

  const ta    = document.getElementById('source-editor');
  const text  = ta.value;
  const lower = text.toLowerCase();
  const ql    = query.toLowerCase();
  let i = 0;
  while ((i = lower.indexOf(ql, i)) !== -1) {
    _searchSrcMatches.push({ start: i, end: i + query.length });
    i += query.length;
  }
  _searchCurrent = _searchSrcMatches.length > 0 ? 0 : -1;
  _updateSearchCount();
  _jumpToSourceMatch();
}

function _jumpToSourceMatch() {
  if (_searchCurrent < 0 || !_searchSrcMatches[_searchCurrent]) return;
  const ta   = document.getElementById('source-editor');
  const pane = document.getElementById('source-scroll');
  const m    = _searchSrcMatches[_searchCurrent];

  // 1. Placer la sélection puis donner le focus (rend la sélection visible)
  ta.setSelectionRange(m.start, m.end);
  ta.focus();

  // 2. Calculer la position Y exacte du match via un miroir hors-écran
  //    (tient compte du retour à la ligne automatique)
  const cs          = getComputedStyle(ta);
  const paddingTop  = parseFloat(cs.paddingTop)   || 24;
  const paddingLeft = parseFloat(cs.paddingLeft)  || 16;
  const paddingRight= parseFloat(cs.paddingRight) || 32;
  const lineHeight  = parseFloat(cs.lineHeight)   || (parseFloat(cs.fontSize) * 1.6);
  const linesBefore = ta.value.substring(0, m.start).split('\n').length - 1;
  const targetScrollTop = Math.max(0, paddingTop + linesBefore * lineHeight - pane.clientHeight / 2 + lineHeight / 2);

  // 3. Appliquer le scroll APRÈS l'auto-scroll du navigateur déclenché par focus()
  setTimeout(() => {
    pane.scrollTop = targetScrollTop;
    _drawSourceMinimap();
  }, 0);
}

/* ── Mode preview ── */
function _clearSearchHighlights() {
  document.getElementById('preview').querySelectorAll('mark.search-hl').forEach(el => {
    const parent = el.parentNode;
    if (!parent) return;
    while (el.firstChild) parent.insertBefore(el.firstChild, el);
    parent.removeChild(el);
  });
  document.getElementById('preview').normalize();
}

function _doSearchPreview(query) {
  _clearSearchHighlights();
  _searchMatches = [];
  const countEl = document.getElementById('search-count');
  const input   = document.getElementById('search-input');
  if (!query) { countEl.textContent = ''; input.classList.remove('search-no-match'); return; }

  const preview = document.getElementById('preview');
  const ql      = query.toLowerCase();

  // Collecter tous les nœuds texte pertinents
  const walker = document.createTreeWalker(preview, NodeFilter.SHOW_TEXT, {
    acceptNode(n) {
      const p = n.parentElement;
      if (p.classList.contains('fenced-div-label') ||
          p.classList.contains('heading-bookmark') ||
          p.closest('.md-rawhtml')) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    }
  });
  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);

  for (const node of nodes) {
    const text  = node.textContent;
    const lower = text.toLowerCase();
    let pos = 0, last = 0;
    const frag = document.createDocumentFragment();
    let found = false;
    while ((pos = lower.indexOf(ql, last)) !== -1) {
      found = true;
      if (pos > last) frag.appendChild(document.createTextNode(text.slice(last, pos)));
      const mark = document.createElement('mark');
      mark.className = 'search-hl';
      mark.textContent = text.slice(pos, pos + query.length);
      frag.appendChild(mark);
      _searchMatches.push(mark);
      last = pos + query.length;
    }
    if (found) {
      if (last < text.length) frag.appendChild(document.createTextNode(text.slice(last)));
      node.replaceWith(frag);
    }
  }

  _searchCurrent = _searchMatches.length > 0 ? 0 : -1;
  _updateSearchCount();
  _highlightCurrentPreview();
}

function _highlightCurrentPreview() {
  _searchMatches.forEach((m, i) => {
    m.className = i === _searchCurrent ? 'search-hl search-current' : 'search-hl';
  });
  if (_searchCurrent >= 0 && _searchMatches[_searchCurrent]) {
    _searchMatches[_searchCurrent].scrollIntoView({ block: 'center', behavior: 'smooth' });
  }
}

function _updateSearchCount() {
  const total  = state.sourceMode ? _searchSrcMatches.length : _searchMatches.length;
  const countEl = document.getElementById('search-count');
  const input   = document.getElementById('search-input');
  if (total === 0) {
    countEl.textContent = 'Aucun';
    input.classList.add('search-no-match');
  } else {
    countEl.textContent = `${_searchCurrent + 1} / ${total}`;
    input.classList.remove('search-no-match');
  }
}

function searchNext() {
  if (state.sourceMode) {
    if (!_searchSrcMatches.length) return;
    _searchCurrent = (_searchCurrent + 1) % _searchSrcMatches.length;
    _updateSearchCount();
    _jumpToSourceMatch();
  } else {
    if (!_searchMatches.length) return;
    _searchCurrent = (_searchCurrent + 1) % _searchMatches.length;
    _updateSearchCount();
    _highlightCurrentPreview();
  }
}

function searchPrev() {
  if (state.sourceMode) {
    if (!_searchSrcMatches.length) return;
    _searchCurrent = (_searchCurrent - 1 + _searchSrcMatches.length) % _searchSrcMatches.length;
    _updateSearchCount();
    _jumpToSourceMatch();
  } else {
    if (!_searchMatches.length) return;
    _searchCurrent = (_searchCurrent - 1 + _searchMatches.length) % _searchMatches.length;
    _updateSearchCount();
    _highlightCurrentPreview();
  }
}

// ── Keyboard shortcuts ────────────────────────────────────────────────────────
function setupKeyboardShortcuts() {
  document.addEventListener('keydown', e => {
    const ctrl = e.ctrlKey || e.metaKey;
    if (ctrl && e.key === 'b') { e.preventDefault(); applyBold(); }
    if (ctrl && e.key === 'i') { e.preventDefault(); applyItalic(); }
    if (ctrl && e.key === 'u') { e.preventDefault(); applyUnderline(); }
    if (ctrl && e.key === 'e') { e.preventDefault(); applyCode(); }
    if (ctrl && e.shiftKey && e.key === 'X') { e.preventDefault(); applyStrikethrough(); }
    if (ctrl && e.shiftKey && e.key === 'K') { e.preventDefault(); applySmallCaps(); }
    if (ctrl && e.shiftKey && e.key === 'L') { e.preventDefault(); applyBulletList(); }
    if (ctrl && e.shiftKey && e.key === 'O') { e.preventDefault(); applyOrderedList(); }
    if (ctrl && e.shiftKey && e.key === 'F') { e.preventDefault(); applyCheckboxList(); }
    if (ctrl && e.shiftKey && e.key === 'A') {
      e.preventDefault();
      if (!state.sourceMode) {
        const _prev = document.getElementById('preview');
        const _sel  = window.getSelection();
        _scInsertionRange  = (_sel && _sel.rangeCount > 0 && _prev.contains(_sel.anchorNode))
          ? _sel.getRangeAt(0).cloneRange() : null;
        _scInsertionAnchor = null;
      }
      insertTodoComment();
    }
    if (ctrl && e.shiftKey && e.key === 'C') { e.preventDefault(); applyCodeBlock(); }
    if (ctrl && e.shiftKey && e.key === 'Y') { e.preventDefault(); applyYamlBlock(); }
    if (ctrl && e.shiftKey && e.key === 'Q') { e.preventDefault(); applyBlockquote(); }
    if (ctrl && e.shiftKey && e.key === 'T') { e.preventDefault(); applyTable(); }
    if (ctrl && e.key === '0') { e.preventDefault(); applyHeading(0); }
    if (ctrl && e.key === '1') { e.preventDefault(); applyHeading(1); }
    if (ctrl && e.key === '2') { e.preventDefault(); applyHeading(2); }
    if (ctrl && e.key === '3') { e.preventDefault(); applyHeading(3); }
    if (ctrl && e.key === '4') { e.preventDefault(); applyHeading(4); }
    if (ctrl && e.key === '5') { e.preventDefault(); applyHeading(5); }
    if (ctrl && e.key === '6') { e.preventDefault(); applyHeading(6); }
    if (ctrl && e.shiftKey && e.key === 'I') { e.preventDefault(); applyImage(); }
    if (ctrl && e.shiftKey && e.key === 'D') { e.preventDefault(); applyFencedDiv(); }
    if (ctrl && e.shiftKey && e.key === 'V') { e.preventDefault(); pasteNoFormat(); }
    if (ctrl && e.shiftKey && e.key === 'P') { e.preventDefault(); toggleTightList(); }
    if (ctrl && e.shiftKey && e.key === 'B') { e.preventDefault(); triggerBookmark(); }
    if (ctrl && e.key === 'k') { e.preventDefault(); openLinkDialog(); }
    if (ctrl && e.key === 'o') { e.preventDefault(); openFile(); }
    if (ctrl && e.key === 's' && !e.shiftKey) { e.preventDefault(); saveFile(); }
    if (ctrl && e.key === 'S') { e.preventDefault(); saveFileAs(); }
    if (ctrl && e.key === 'w') { e.preventDefault(); closeCurrentTab(); }
    if (ctrl && e.key === 't') { e.preventDefault(); newTab(); }
    if (ctrl && e.key === 'f') { e.preventDefault(); openSearch(); }
    if (ctrl && e.altKey && e.key === 'ArrowLeft')  { e.preventDefault(); toggleSidebar(); }
    if (ctrl && e.altKey && e.key === 'ArrowRight') { e.preventDefault(); toggleTOC(); }
    if (ctrl && e.altKey && (e.key === 'm' || e.key === 'M')) { e.preventDefault(); addAnnotation(); }
    if (ctrl && !e.shiftKey && e.key === 'Enter') { e.preventDefault(); applyPagebreak(); }
    if (ctrl && e.shiftKey && e.key === 'Enter') { e.preventDefault(); insertEmptyParaAbove(); }
    if (ctrl && e.shiftKey && e.key === 'R') { e.preventDefault(); applyHorizontalRule(); }
    if (ctrl && e.shiftKey && e.key === 'N') { e.preventDefault(); applyFootnote(); }
    if (ctrl && e.shiftKey && e.key === 'E') { e.preventDefault(); clearFormatting(); }
    if (ctrl && e.shiftKey && e.key === 'H') {
      e.preventDefault();
      if (!state.sourceMode) {
        const sel = window.getSelection();
        const preview = document.getElementById('preview');
        _scInsertionRange = (sel && sel.rangeCount > 0 && preview.contains(sel.anchorNode))
          ? sel.getRangeAt(0).cloneRange() : null;
        _scInsertionAnchor = null;
      }
      openSCDialog(null);
    }
    if (ctrl && e.shiftKey && e.key === 'M') {
      e.preventDefault();
      const _t = getActiveTab();
      const _ext = _t?.path ? _t.path.toLowerCase().split('.').pop() : '';
      if (_ext === 'md' || _ext === 'qmd' || !_t?.path) {
        const next = !state.sourceMode;
        document.getElementById('source-toggle-input').checked = next;
        toggleSourceMode(next);
      }
    }
    if (ctrl && e.key === 'Tab') {
      e.preventDefault();
      if (state.tabs.length < 2) return;
      const idx = state.tabs.findIndex(t => t.id === state.activeTabId);
      const next = e.shiftKey
        ? (idx - 1 + state.tabs.length) % state.tabs.length
        : (idx + 1) % state.tabs.length;
      switchToTab(state.tabs[next].id);
    }
    if (ctrl && (e.key === 'PageDown' || e.key === 'PageUp')) {
      e.preventDefault();
      if (state.tabs.length < 2) return;
      const idx = state.tabs.findIndex(t => t.id === state.activeTabId);
      const next = e.key === 'PageDown'
        ? (idx + 1) % state.tabs.length
        : (idx - 1 + state.tabs.length) % state.tabs.length;
      switchToTab(state.tabs[next].id);
    }
    if (e.key === 'F11') { e.preventDefault(); toggleFullscreen(); }
    if (e.key === 'F3') {
      e.preventDefault();
      const bar = document.getElementById('search-bar');
      if (!bar.classList.contains('open')) { openSearch(); return; }
      e.shiftKey ? searchPrev() : searchNext();
    }
    if (e.key === 'Escape' && document.getElementById('search-bar').classList.contains('open')) {
      e.preventDefault(); closeSearch();
    }
  });
}

// ── Fenêtre d'aide ────────────────────────────────────────────────────────────
function openHelpDialog() {
  document.getElementById('help-overlay').classList.add('open');
}
function closeHelpDialog(e) {
  if (e && e.target !== document.getElementById('help-overlay')) return;
  document.getElementById('help-overlay').classList.remove('open');
}

// ── Ouverture fichier initial (appelé depuis Python via evaluate_js) ──────────
function openInitialFile(data) {
  if (!turndown) {
    // init() pas encore terminé (marked CDN pas chargé) — mise en file d'attente
    pendingInitialFile = data;
    return;
  }
  const tab = createTab(data.name, data.path, data.content);
  switchToTab(tab.id);
  sidebarNavigateTo(data.path);
}

// ── Citations BibTeX (@-autocomplete) ────────────────────────────────────────

async function _ensureBibEntries() {
  if (_bibEntries !== null) return _bibEntries;
  const s = loadSettings();
  if (!s.bibFile || !window.pywebview) { _bibEntries = []; return []; }
  try {
    const res = await window.pywebview.api.parse_bib_file(s.bibFile);
    _bibEntries = Array.isArray(res) ? res : [];
  } catch { _bibEntries = []; }
  return _bibEntries;
}

function _citeQueryAtCaret() {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (!range.collapsed) return null;
  const node = range.startContainer;
  if (node.nodeType !== Node.TEXT_NODE) return null;
  const text = node.textContent.substring(0, range.startOffset);
  const at = text.lastIndexOf('@');
  if (at === -1) return null;
  const q = text.substring(at + 1);
  if (!/^[\w-]*$/.test(q)) return null;
  return { query: q, atIdx: at, textNode: node, endOffset: range.startOffset };
}

function _citeQueryInTextarea(ta) {
  const pos = ta.selectionStart;
  if (pos !== ta.selectionEnd) return null;
  const text = ta.value.substring(0, pos);
  const at = text.lastIndexOf('@');
  if (at === -1) return null;
  const q = text.substring(at + 1);
  if (!/^[\w-]*$/.test(q)) return null;
  return { query: q, atIdx: at, endOffset: pos };
}

function _citeGetTextareaCaretRect(ta) {
  const cs = window.getComputedStyle(ta);
  const div = document.createElement('div');
  for (const p of ['fontFamily','fontSize','fontWeight','fontStyle','letterSpacing',
                    'lineHeight','paddingTop','paddingRight','paddingBottom','paddingLeft',
                    'borderTopWidth','borderRightWidth','borderBottomWidth','borderLeftWidth',
                    'boxSizing','tabSize','whiteSpace','wordWrap','overflowWrap']) {
    try { div.style[p] = cs[p]; } catch(_) {}
  }
  div.style.position  = 'absolute';
  div.style.visibility = 'hidden';
  div.style.overflow   = 'hidden';
  div.style.width      = ta.clientWidth + 'px';
  div.style.whiteSpace = 'pre-wrap';
  div.style.wordBreak  = 'break-all';
  const taRect = ta.getBoundingClientRect();
  div.style.left = (taRect.left + window.scrollX) + 'px';
  div.style.top  = (taRect.top + window.scrollY - ta.scrollTop) + 'px';
  div.appendChild(document.createTextNode(ta.value.substring(0, ta.selectionStart)));
  const span = document.createElement('span');
  span.textContent = '​';
  div.appendChild(span);
  document.body.appendChild(div);
  const sr = span.getBoundingClientRect();
  document.body.removeChild(div);
  const lh = parseFloat(cs.lineHeight) || parseFloat(cs.fontSize) * 1.4;
  return { top: sr.top, bottom: sr.top + lh, left: sr.left, right: sr.right };
}

function _hideCitePopup() {
  document.getElementById('cite-popup').style.display = 'none';
  _citePopupIdx = -1;
}

function _buildCitePopup(entries, query) {
  const popup = document.getElementById('cite-popup');
  const q = query.toLowerCase();
  const filtered = entries.filter(e =>
    !q ||
    e.key.toLowerCase().includes(q) ||
    e.title.toLowerCase().includes(q) ||
    e.author.toLowerCase().includes(q)
  ).slice(0, 40);
  if (!filtered.length) { _hideCitePopup(); return null; }

  popup.innerHTML = '';
  _citePopupIdx = -1;

  filtered.forEach(e => {
    const opt = document.createElement('div');
    opt.className = 'cite-option';
    opt.dataset.key = e.key;

    const keyEl = document.createElement('div');
    keyEl.className = 'cite-option-key';
    keyEl.textContent = '@' + e.key;
    opt.appendChild(keyEl);

    const parts = [];
    if (e.title)  parts.push(e.title);
    const firstAuthor = e.author ? e.author.split(/\band\b/i)[0].split(',')[0].trim() : '';
    const ay = [firstAuthor, e.year].filter(Boolean).join(', ');
    if (ay) parts.push('(' + ay + ')');
    if (parts.length) {
      const meta = document.createElement('div');
      meta.className = 'cite-option-meta';
      meta.textContent = parts.join(' — ');
      opt.appendChild(meta);
    }
    opt.addEventListener('mousedown', ev => { ev.preventDefault(); _applyCite(e.key); });
    popup.appendChild(opt);
  });
  return filtered;
}

function _positionCitePopup(rect) {
  const popup = document.getElementById('cite-popup');
  popup.style.display = 'block';
  const vw = window.innerWidth, vh = window.innerHeight;
  let top  = rect.bottom + 4;
  let left = rect.left;
  if (left + 360 > vw - 8) left = Math.max(8, vw - 360 - 8);
  // Recalculate height after display:block
  const h = popup.offsetHeight;
  if (top + h > vh - 8) top = Math.max(8, rect.top - h - 4);
  popup.style.top  = top + 'px';
  popup.style.left = left + 'px';
}

function _citeMoveFocus(dir) {
  const popup = document.getElementById('cite-popup');
  const opts  = popup.querySelectorAll('.cite-option');
  if (!opts.length) return;
  opts[_citePopupIdx]?.classList.remove('focused');
  _citePopupIdx = Math.max(0, Math.min(opts.length - 1, _citePopupIdx + dir));
  opts[_citePopupIdx].classList.add('focused');
  opts[_citePopupIdx].scrollIntoView({ block: 'nearest' });
}

function _applyCite(key) {
  _hideCitePopup();
  if (state.sourceMode) {
    const ta = document.getElementById('source-editor');
    const info = _citeQueryInTextarea(ta);
    if (!info) return;
    ta.setRangeText('@' + key, info.atIdx, info.endOffset, 'end');
    ta.focus();
    onSourceInput();
  } else {
    const info = _citeQueryAtCaret();
    if (!info) return;
    const sel = window.getSelection();
    const rng = document.createRange();
    rng.setStart(info.textNode, info.atIdx);
    rng.setEnd(info.textNode, info.endOffset);
    rng.deleteContents();
    const tn = document.createTextNode('@' + key);
    rng.insertNode(tn);
    rng.setStartAfter(tn);
    rng.collapse(true);
    sel.removeAllRanges();
    sel.addRange(rng);
    syncPreviewToContent();
  }
}

async function _checkCitePopup(isSource) {
  const info = isSource
    ? _citeQueryInTextarea(document.getElementById('source-editor'))
    : _citeQueryAtCaret();
  if (!info) { _hideCitePopup(); return; }
  const entries = await _ensureBibEntries();
  if (!_buildCitePopup(entries, info.query)) return;
  let rect;
  if (isSource) {
    rect = _citeGetTextareaCaretRect(document.getElementById('source-editor'));
  } else {
    rect = window.getSelection().getRangeAt(0).getBoundingClientRect();
  }
  _positionCitePopup(rect);
}

function _citeHandleKeydown(e) {
  const popup = document.getElementById('cite-popup');
  if (popup.style.display === 'none') return false;
  if (e.key === 'ArrowDown') { e.preventDefault(); _citeMoveFocus(1); return true; }
  if (e.key === 'ArrowUp')   { e.preventDefault(); _citeMoveFocus(-1); return true; }
  if (e.key === 'Escape')    { e.preventDefault(); _hideCitePopup(); return true; }
  if (e.key === 'Enter' || e.key === 'Tab') {
    const focused = popup.querySelector('.cite-option.focused');
    if (focused) { e.preventDefault(); _applyCite(focused.dataset.key); return true; }
  }
  return false;
}

// ── Tooltip référence bibliographique ────────────────────────────────────────

let _citeTooltipTimer = null;

function _formatCitationTooltip(entry) {
  const tip = document.getElementById('citation-tooltip');
  if (!tip) return;
  tip.innerHTML = '';

  // Auteurs formatés : "Nom, P.; Nom2, P2"
  if (entry.author) {
    const authors = entry.author.split(/\s+and\s+/i).map(a => {
      const parts = a.split(',');
      if (parts.length >= 2) return parts[0].trim() + ', ' + parts[1].trim().charAt(0) + '.';
      return a.trim();
    });
    const aEl = document.createElement('div');
    aEl.className = 'ctt-authors';
    aEl.textContent = authors.join(' ; ') + (entry.year ? ' (' + entry.year + ')' : '');
    tip.appendChild(aEl);
  }

  if (entry.title) {
    const tEl = document.createElement('div');
    tEl.className = 'ctt-title';
    tEl.textContent = entry.title;
    tip.appendChild(tEl);
  }

  // Lieu de publication : journal / booktitle / publisher
  const venue = [];
  const src = entry.journal || entry.booktitle || entry.publisher || '';
  if (src) venue.push(src);
  if (entry.volume) {
    let vol = entry.volume;
    if (entry.number) vol += '(' + entry.number + ')';
    venue.push(vol);
  }
  if (entry.pages) venue.push('p. ' + entry.pages);
  if (venue.length) {
    const vEl = document.createElement('div');
    vEl.className = 'ctt-venue';
    vEl.textContent = venue.join(', ');
    tip.appendChild(vEl);
  }

  const kEl = document.createElement('div');
  kEl.className = 'ctt-key';
  kEl.textContent = '@' + entry.key;
  tip.appendChild(kEl);
}

function _positionCitationTooltip(targetEl) {
  const tip  = document.getElementById('citation-tooltip');
  if (!tip) return;
  const rect = targetEl.getBoundingClientRect();
  const vw = window.innerWidth, vh = window.innerHeight;

  tip.style.left = '-9999px';
  tip.classList.add('visible');
  const tw = tip.offsetWidth, th = tip.offsetHeight;

  let top  = rect.bottom + 6;
  let left = rect.left;
  if (top + th > vh - 8)  top  = rect.top - th - 6;
  if (left + tw > vw - 8) left = Math.max(8, vw - tw - 8);
  if (left < 8)           left = 8;

  tip.style.top  = top + 'px';
  tip.style.left = left + 'px';
}

function _hideCitationTooltip() {
  clearTimeout(_citeTooltipTimer);
  document.getElementById('citation-tooltip')?.classList.remove('visible');
}

function _initCiteListeners() {
  const preview = document.getElementById('preview');
  const sourceEditor = document.getElementById('source-editor');

  // ── Tooltip survol des @citations ────────────────────────────────────────
  preview.addEventListener('mouseover', async e => {
    const el = e.target.closest('.md-citation');
    if (!el) return;
    clearTimeout(_citeTooltipTimer);
    _citeTooltipTimer = setTimeout(async () => {
      const key     = el.dataset.key;
      const entries = await _ensureBibEntries();
      const entry   = entries.find(x => x.key === key);
      if (!entry) return;
      _formatCitationTooltip(entry);
      _positionCitationTooltip(el);
    }, 120);
  });

  preview.addEventListener('mouseout', e => {
    if (!e.target.closest('.md-citation')) return;
    _hideCitationTooltip();
  });

  preview.addEventListener('input', () => {
    if (!state.editingPreview) return;
    _checkCitePopup(false);
  });

  preview.addEventListener('keydown', e => {
    _citeHandleKeydown(e);
  }, true); // capture phase pour être avant les autres handlers

  sourceEditor.addEventListener('input', () => {
    if (!state.sourceMode) return;
    _checkCitePopup(true);
  });

  sourceEditor.addEventListener('keydown', e => {
    _citeHandleKeydown(e);
  }, true);

  // Fermer le popup sur clic hors preview/source
  document.addEventListener('mousedown', e => {
    const popup = document.getElementById('cite-popup');
    if (!popup.contains(e.target) &&
        !preview.contains(e.target) &&
        !sourceEditor.contains(e.target)) {
      _hideCitePopup();
    }
  });
}

// ── Menu "…" onglets en overflow ─────────────────────────────────────────────

let _tabsMenuOpen = false;
let _tabsResizeObserver = null;

function _updateTabWidths() {
  const tabList = document.getElementById('tab-list');
  const btn     = document.getElementById('tabs-menu-btn');
  const label   = document.getElementById('tabs-menu-label');
  if (!tabList || !btn) return;

  const n = state.tabs.length;
  btn.style.display = 'none';
  if (n === 0) return;

  // Après rendu : si les onglets (largeur fixe) débordent du conteneur → bouton "…"
  requestAnimationFrame(() => {
    if (tabList.scrollWidth > tabList.clientWidth + 1) {
      if (label) label.textContent = '… ' + n;
      btn.style.display = '';
    }
  });
}

function _buildTabsMenuList(filter) {
  const list = document.getElementById('tabs-menu-list');
  if (!list) return;
  const q = (filter || '').toLowerCase();
  list.innerHTML = '';

  state.tabs
    .filter(t => !q || t.name.toLowerCase().includes(q))
    .forEach(tab => {
      const item = document.createElement('div');
      item.className = 'tabs-menu-item' + (tab.id === state.activeTabId ? ' active' : '');
      item.title = tab.path || tab.name;

      if (tab.modified) {
        const dot = document.createElement('span');
        dot.className = 'tabs-menu-item-dot';
        dot.textContent = '●';
        item.appendChild(dot);
      }

      const name = document.createElement('span');
      name.className = 'tabs-menu-item-name';
      name.textContent = tab.name;
      item.appendChild(name);

      const closeBtn = document.createElement('button');
      closeBtn.className = 'tabs-menu-item-close';
      closeBtn.textContent = '×';
      closeBtn.title = 'Fermer';
      closeBtn.addEventListener('click', e => {
        e.stopPropagation();
        closeTabsMenu();
        closeTab(tab.id);
      });
      item.appendChild(closeBtn);

      item.addEventListener('click', () => {
        closeTabsMenu();
        switchToTab(tab.id);
      });
      list.appendChild(item);
    });
}

function toggleTabsMenu() {
  if (_tabsMenuOpen) { closeTabsMenu(); return; }

  const btn      = document.getElementById('tabs-menu-btn');
  const dropdown = document.getElementById('tabs-menu-dropdown');
  const search   = document.getElementById('tabs-menu-search');
  if (!btn || !dropdown) return;

  _buildTabsMenuList('');
  if (search) search.value = '';

  // Positionner sous le bouton, aligné à droite
  const rect = btn.getBoundingClientRect();
  const top  = rect.bottom + 4;
  let   left = rect.left;
  const vw   = window.innerWidth;
  if (left + 280 > vw - 8) left = vw - 288;
  if (left < 8) left = 8;

  dropdown.style.top  = top + 'px';
  dropdown.style.left = left + 'px';
  dropdown.classList.add('open');
  btn.classList.add('open');
  _tabsMenuOpen = true;

  requestAnimationFrame(() => {
    const active = dropdown.querySelector('.tabs-menu-item.active');
    if (active) active.scrollIntoView({ block: 'nearest' });
    if (search) search.focus();
  });
}

function closeTabsMenu() {
  document.getElementById('tabs-menu-dropdown')?.classList.remove('open');
  document.getElementById('tabs-menu-btn')?.classList.remove('open');
  _tabsMenuOpen = false;
}

function filterTabsMenu(value) {
  _buildTabsMenuList(value);
}

function _initTabsMenuClose() {
  document.addEventListener('mousedown', e => {
    if (!_tabsMenuOpen) return;
    const dropdown = document.getElementById('tabs-menu-dropdown');
    const btn      = document.getElementById('tabs-menu-btn');
    if (dropdown && !dropdown.contains(e.target) && btn && !btn.contains(e.target)) {
      closeTabsMenu();
    }
  });
  document.addEventListener('keydown', e => {
    if (_tabsMenuOpen && e.key === 'Escape') closeTabsMenu();
  });
}

function _initTabsResizeObserver() {
  const tabArea = document.getElementById('tab-area');
  if (!tabArea || _tabsResizeObserver) return;
  _tabsResizeObserver = new ResizeObserver(() => _updateTabWidths());
  _tabsResizeObserver.observe(tabArea);
}

// ── Exports for HTML onclick attributes ───────────────────────────────────────
window.openFile = openFile;
window.saveFile = saveFile;
window.saveFileAs = saveFileAs;
window.closeCurrentTab = closeCurrentTab;
window.closeTab = closeTab;
window.switchToTab = switchToTab;
window.newTab = newTab;
window.toggleTabsMenu  = toggleTabsMenu;
window.closeTabsMenu   = closeTabsMenu;
window.filterTabsMenu  = filterTabsMenu;
window.toggleMenu = toggleMenu;
window.closeMenu = closeMenu;
window.toggleSourceMode = toggleSourceMode;
window.minimizeWindow = minimizeWindow;
window.toggleFullscreen = toggleFullscreen;
window.quitApp = quitApp;
window.setZoom = setZoom;
window.openInitialFile = openInitialFile;
window.hideContextMenu      = hideContextMenu;
window.hideTableContextMenu = hideTableContextMenu;
window.alignColumn          = alignColumn;
window.tableAddRow          = tableAddRow;
window.tableAddRowPrompt    = tableAddRowPrompt;
window.tableDeleteRow       = tableDeleteRow;
window.tableAddCol          = tableAddCol;
window.tableDeleteCol       = tableDeleteCol;
window.tableDelete          = tableDelete;
window._colWidthUpdateTotal = _colWidthUpdateTotal;
window.openColWidthDialog   = openColWidthDialog;
window.closeColWidthDialog  = closeColWidthDialog;
window.confirmColWidths     = confirmColWidths;
window.openSettings          = openSettings;
window.openProjectChooser    = openProjectChooser;
window.openProjectChooserFromSettings = openProjectChooserFromSettings;
window.closeProjectChooser   = closeProjectChooser;
window.chooseProject         = chooseProject;
window.chooseProjectFromDialog = chooseProjectFromDialog;
window.openRecentProject     = openRecentProject;
window.chooseBibFile   = chooseBibFile;
window.resetHeadingColor = resetHeadingColor;
window.openHelpDialog  = openHelpDialog;
window.closeHelpDialog = closeHelpDialog;
window.closeSettings = closeSettings;
window.closeSettingsOnOverlay = closeSettingsOnOverlay;
window.applySettings = applySettings;
window.openLinkDialog    = openLinkDialog;
window.closeLinkDialog   = closeLinkDialog;
window.selectLinkType    = selectLinkType;
window.confirmLinkDialog = confirmLinkDialog;
window.browseLinkFile    = browseLinkFile;
window._updateLinkPreview = _updateLinkPreview;
window.editLinkFromMenu  = editLinkFromMenu;
window.deleteLinkFromMenu = deleteLinkFromMenu;
window.hideLinkContextMenu = hideLinkContextMenu;
window.applyHighlight  = applyHighlight;
window.applyTextColor  = applyTextColor;
window.applyItalic = applyItalic;
window.applyUnderline = applyUnderline;
window.applyCode = applyCode;
window.applyStrikethrough = applyStrikethrough;
window.applySmallCaps  = applySmallCaps;
window.clearFormatting = clearFormatting;
window.applyHorizontalRule = applyHorizontalRule;
window.applyBulletList    = applyBulletList;
window.applyOrderedList   = applyOrderedList;
window.applyCheckboxList  = applyCheckboxList;
window.applyCodeBlock = applyCodeBlock;
window.applyYamlBlock  = applyYamlBlock;
window.applyBlockquote = applyBlockquote;
window.applyTable      = applyTable;
window.applyImage      = applyImage;
window.hideImageContextMenu = hideImageContextMenu;
window.openImageDialog      = openImageDialog;
window.closeImageDialog     = closeImageDialog;
window.selectImageAlign       = selectImageAlign;
window.selectImageVisibility  = selectImageVisibility;
window.confirmImageDialog   = confirmImageDialog;
window.deleteImage          = deleteImage;
window.refreshImage         = refreshImage;
window.copyImage            = copyImage;
window.changeImageFile      = changeImageFile;
window.copyImagePath        = copyImagePath;
window.insertCaptionLink    = insertCaptionLink;
window.toggleSidebar        = toggleSidebar;
window.changeSidebarDir     = changeSidebarDir;
window.toggleTOC            = toggleTOC;
window.toggleHeadingMenu    = toggleHeadingMenu;
window.closeHeadingMenu     = closeHeadingMenu;
window.applyHeading         = applyHeading;
window.insertEmptyParaAbove = insertEmptyParaAbove;
window.openBookmarkDialog   = openBookmarkDialog;
window.closeBookmarkDialog  = closeBookmarkDialog;
window.confirmBookmarkDialog= confirmBookmarkDialog;
window.removeBookmark       = removeBookmark;
window.openRawHtmlDialog    = openRawHtmlDialog;
window.closeRawHtmlDialog   = closeRawHtmlDialog;
window.confirmRawHtmlDialog = confirmRawHtmlDialog;
window.editRawHtmlFromMenu  = editRawHtmlFromMenu;
window.deleteRawHtmlFromMenu= deleteRawHtmlFromMenu;
window.hideRawHtmlContextMenu = hideRawHtmlContextMenu;
window.insertHtmlComment    = insertHtmlComment;
window.insertTodoComment    = insertTodoComment;
window.openTodoPanel        = openTodoPanel;
window.closeTodoPanel       = closeTodoPanel;
window.todoSetFilter        = todoSetFilter;
window.todoClearDone        = todoClearDone;
window.todoToggle           = todoToggle;
window.todoScanDirectory    = todoScanDirectory;
window.resetTabStops        = resetTabStops;
window.editCommentFromMenu  = editCommentFromMenu;
window.deleteCommentFromMenu= deleteCommentFromMenu;
window.hideCommentContextMenu=hideCommentContextMenu;
window.commentCtxCopy       = commentCtxCopy;
window.commentCtxCut        = commentCtxCut;
window.commentCtxPaste      = commentCtxPaste;
window.openSCDialog         = openSCDialog;
window.closeSCDialog        = closeSCDialog;
window.confirmSCDialog      = confirmSCDialog;
window._scBrowseInclude     = _scBrowseInclude;
window.editSCFromMenu       = editSCFromMenu;
window.deleteSCFromMenu     = deleteSCFromMenu;
window.hideSCContextMenu    = hideSCContextMenu;
window.applyFootnote        = applyFootnote;
window.cancelFootnoteDialog = cancelFootnoteDialog;
window.confirmFootnoteDialog= confirmFootnoteDialog;
window.editFootnoteFromMenu = editFootnoteFromMenu;
window.deleteFootnoteFromMenu=deleteFootnoteFromMenu;
window.showFootnoteContextMenu=showFootnoteContextMenu;
window.hideFootnoteContextMenu=hideFootnoteContextMenu;
window._fnFormat            = _fnFormat;
window._fnKeydown           = _fnKeydown;
window.applyFencedDiv       = applyFencedDiv;
window.openFDivDialog       = openFDivDialog;
window.closeFDivDialog      = closeFDivDialog;
window.confirmFDivDialog    = confirmFDivDialog;
window.deleteFDiv              = deleteFDiv;
window.hideFDivContextMenu     = hideFDivContextMenu;
window.wrapCodeBlockInDiv      = wrapCodeBlockInDiv;
window.unwrapCodeBlockFromDiv  = unwrapCodeBlockFromDiv;
window.deleteCodeBlock         = deleteCodeBlock;
window.hideCodeBlockContextMenu = hideCodeBlockContextMenu;
window._fdivToggleClass     = _fdivToggleClass;
window._fdivSyncChips       = _fdivSyncChips;
window._fdivWidthToggle     = _fdivWidthToggle;
window.openSearch           = openSearch;
window.closeSearch          = closeSearch;
window.onSearchInput        = onSearchInput;
window.onSearchKeydown      = onSearchKeydown;
window.searchNext           = searchNext;
window.searchPrev           = searchPrev;
window.toggleSidebarSearch          = toggleSidebarSearch;
window.sidebarSearchClear           = sidebarSearchClear;
window.onSidebarSearchInput         = onSidebarSearchInput;
window.onSidebarSearchKeydown       = onSidebarSearchKeydown;
window.sidebarFilenameSearchClear   = sidebarFilenameSearchClear;
window.onSidebarFilenameSearchInput = onSidebarFilenameSearchInput;
window.onSidebarFilenameSearchKeydown = onSidebarFilenameSearchKeydown;
window.onLogSearchInput     = onLogSearchInput;
window.onLogSearchKeydown   = onLogSearchKeydown;
window.logSearchNext        = logSearchNext;
window.logSearchPrev        = logSearchPrev;
window.logSearchClear       = logSearchClear;
window.hidePumlContextMenu  = hidePumlContextMenu;
window.pumlCtxCopy          = pumlCtxCopy;
window.pumlCtxCut           = pumlCtxCut;
window.pumlCtxPaste         = pumlCtxPaste;
window.pumlShowPreview      = pumlShowPreview;
window.closePumlPreview     = closePumlPreview;
window.pumlPreviewSwitchFmt = pumlPreviewSwitchFmt;
window.pumlZoomIn           = pumlZoomIn;
window.pumlZoomOut          = pumlZoomOut;
window.pumlZoomReset        = pumlZoomReset;
window.imageZoomIn          = imageZoomIn;
window.imageZoomOut         = imageZoomOut;
window.imageZoomReset       = imageZoomReset;
window.onMetaFieldChange     = onMetaFieldChange;
window.toggleMetadataSection = toggleMetadataSection;
window.toggleMetadataOther   = toggleMetadataOther;
window.toggleTocSection      = toggleTocSection;
window.addAnnotation           = addAnnotation;
window.selectAnnotationAnchor  = selectAnnotationAnchor;
window.onAnnotationTextInput   = onAnnotationTextInput;
window.deleteAnnotation        = deleteAnnotation;
window.toggleAnnotationsSection= toggleAnnotationsSection;
window._applyContentMaxWidth = _applyContentMaxWidth;
Object.defineProperty(window, '_contentMaxWidth', { get: () => _contentMaxWidth });
