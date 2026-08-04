// email-modules.js — editable email block definitions for the Missouri S&T builder.
// Each module: { id, name, desc, tags, fields[], render(values, brand) -> email-safe HTML string }
// brand (b): { FF, accent, btn, R, dept, arc, barR }
//
// Layout technique: fluid-hybrid, not media queries. BeeFree's "paste HTML" block sanitizes
// to body-only markup and strips <head>/<style>, so @media breakpoints never survive. Instead,
// every column is a `display:inline-block;width:100%;max-width:Npx` div — it fluidly shrinks
// with its container and naturally wraps to a new line once two columns can no longer fit side
// by side, with no breakpoint required. Outlook (Word engine) ignores max-width/inline-block,
// so each column is also wrapped in an `<!--[if mso]>` conditional real <table><td> of the same
// proportional width, seen only by Outlook. Everything else — layout that actually needs a
// real <table> (data rows, agendas) — stays a table, just fluid-width instead of fixed-width.

export const FF = "'Libre Franklin',Arial,Helvetica,sans-serif";

export function mkBtn(b, label, href, bg, textColor) {
  const plain = String(label).replace(/&[a-z]+;/g, 'x');
  const w = Math.max(150, plain.length * 9 + 60);
  const btnR = 25;
  const btnArc = '52%';
  return `<!--[if mso]>
              <v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word" href="${href}" style="height:50px;v-text-anchor:middle;width:${w}px;" arcsize="${btnArc}" strokecolor="${bg}" fillcolor="${bg}">
                <w:anchorlock/>
                <center style="color:${textColor};font-family:Arial,sans-serif;font-size:15px;font-weight:bold;">${plain.trim()}</center>
              </v:roundrect>
              <![endif]-->
              <!--[if !mso]><!-->
              <a href="${href}" target="_blank" style="display:inline-block;padding:15px 34px;font-family:${FF};font-size:15px;line-height:18px;font-weight:700;color:${textColor};text-decoration:none;background-color:${bg};border-radius:${btnR}px;">${label}</a>
              <!--<![endif]-->`;
}

// Data-URI SVG placeholder image — a real <img> tag the user repoints in Emma. Renders cleanly in the preview.
export function phImg(label, w, h, accent) {
  const svg = `<svg xmlns='http://www.w3.org/2000/svg' width='${w}' height='${h}'><rect width='100%' height='100%' fill='#154734'/><text x='50%' y='50%' fill='${accent}' font-family='Arial,Helvetica,sans-serif' font-size='13' font-weight='700' letter-spacing='2' text-anchor='middle' dominant-baseline='central'>${String(label || 'Image').toUpperCase()}</text></svg>`;
  return `data:image/svg+xml,${encodeURIComponent(svg)}`;
}
export function phVideo(label, w, h, accent) {
  const cx = w / 2, cy = h / 2 - 8, r = 31;
  const svg = `<svg xmlns='http://www.w3.org/2000/svg' width='${w}' height='${h}'><rect width='100%' height='100%' fill='#154734'/><circle cx='${cx}' cy='${cy}' r='${r}' fill='#ffffff'/><path d='M${cx - 9},${cy - 14} L${cx - 9},${cy + 14} L${cx + 17},${cy} Z' fill='#007A33'/><text x='50%' y='${h - 26}' fill='${accent}' font-family='Arial,Helvetica,sans-serif' font-size='11' font-weight='700' letter-spacing='2' text-anchor='middle' dominant-baseline='central'>${String(label || 'Video').toUpperCase()}</text></svg>`;
  return `data:image/svg+xml,${encodeURIComponent(svg)}`;
}

// Parse a multi-line "a | b | c" field into an array of part-arrays.
export function lines(s) {
  return String(s || '').split('\n').map(l => l.trim()).filter(Boolean).map(l => l.split('|').map(x => x.trim()));
}

// ---- Fluid-hybrid primitives ------------------------------------------------------------

// Outer shell: fixed 30px side gutter (a real edge margin, fine as a static value), fluid
// content div capped at maxWidth with an Outlook-only conditional table so old Outlook still
// centers it at a true fixed width instead of stretching to 100% of the message pane.
export function shell(innerHtml, opts = {}) {
  const maxW = opts.maxWidth || 580;
  const bg = opts.bg || '#ffffff';
  const padX = opts.padX != null ? opts.padX : 30;
  const padY = opts.padY != null ? opts.padY : 0;
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;background-color:${bg};"><tr><td align="center" style="padding:${padY}px ${padX}px;"><!--[if mso]><table role="presentation" align="center" width="${maxW}" style="width:${maxW}px;"><tr><td><![endif]--><div style="width:100%;max-width:${maxW}px;margin:0 auto;text-align:left;">${innerHtml}</div><!--[if mso]></td></tr></table><![endif]--></td></tr></table>`;
}

// N fluid columns side by side, wrapping to full width below one another once the container
// can't fit them all. cols: [{ html, width (px cap), style }]
export function fluidCols(cols) {
  const total = cols.reduce((s, c) => s + c.width, 0);
  return cols.map((c, i) => {
    const pct = (c.width / total * 100).toFixed(2);
    const msoValign = c.valign === 'middle' ? 'middle' : 'top';
    const openMso = i === 0
      ? `<!--[if mso]><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td width="${pct}%" valign="${msoValign}"><![endif]-->`
      : `<!--[if mso]><td width="${pct}%" valign="${msoValign}"><![endif]-->`;
    const closeMso = i === cols.length - 1
      ? `<!--[if mso]></td></tr></table><![endif]-->`
      : `<!--[if mso]></td><![endif]-->`;
    const safeW = Math.max(1, c.width - 3);
    // Middle-valign columns (e.g. hero image/text meant to vertically center like a table
    // row would) use inline-flex so centering works without a real table row to equalize on;
    // everything else is a plain inline-block, stacking top-aligned same as the old td[valign=top].
    const display = c.valign === 'middle'
      ? 'display:inline-flex;flex-direction:column;justify-content:center;'
      : 'display:inline-block;';
    return `${openMso}<div style="${display}width:100%;max-width:${safeW}px;vertical-align:top;box-sizing:border-box;${c.style || ''}">${c.html}</div>${closeMso}`;
  }).join('');
}

export function twoCol(aHtml, aw, aStyle, bHtml, bw, bStyle, valign) {
  return fluidCols([{ html: aHtml, width: aw, style: aStyle, valign }, { html: bHtml, width: bw, style: bStyle, valign }]);
}

// For asymmetric pairs (a small badge/number beside a much wider text column): side by side
// at their intended proportion when both fit on one line, but — unlike fluidCols — each one
// individually stretches to fill the full row once it wraps alone on mobile, instead of
// staying stuck at its narrow desktop share. Needs flex (no way to get "grow to fill when
// alone" from plain inline-block/max-width); Outlook still gets the real mso table below,
// completely unaffected by this.
export function fluidColsGrow(cols, alignItems, gap) {
  const parts = cols.map((c, i) => {
    const pct = (c.width / cols.reduce((s, x) => s + x.width, 0) * 100).toFixed(2);
    const openMso = i === 0
      ? `<!--[if mso]><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td width="${pct}%" valign="top"><![endif]-->`
      : `<!--[if mso]><td width="${pct}%" valign="top"><![endif]-->`;
    const closeMso = i === cols.length - 1
      ? `<!--[if mso]></td></tr></table><![endif]-->`
      : `<!--[if mso]></td><![endif]-->`;
    return `${openMso}<div style="flex:${c.width} 1 ${c.width - 3}px;min-width:0;box-sizing:border-box;${c.style || ''}">${c.html}</div>${closeMso}`;
  }).join('');
  return `<div style="display:flex;flex-wrap:wrap;${gap ? `gap:${gap}px;` : ''}${alignItems ? `align-items:${alignItems};` : ''}">${parts}</div>`;
}

// N items in a grid, wrapping every `perRow` into its own MSO-safe fluidCols row (real <table>
// under Outlook, so multi-row wrapping degrades correctly instead of relying on flex-wrap,
// which classic Outlook's Word engine ignores entirely).
export function gridRows(items, perRow, cellHtml, gapY, gapX, cls) {
  const rows = [];
  for (let i = 0; i < items.length; i += perRow) rows.push(items.slice(i, i + perRow));
  const gx = gapX || 12;
  const gy = gapY || 16;
  const colW = Math.floor((580 - gx * (perRow - 1)) / perRow);
  const c = cls || 'grid-cell-' + Math.random().toString(36).slice(2, 8);
  // Fixed-width cells (no flex-grow) so a short final row (e.g. 4 items at 3-per-row) stays
  // left-aligned at its normal column width on desktop instead of stretching to fill the row.
  // Cells naturally wrap every perRow (colW*perRow fills the 580px shell exactly); the media
  // query collapses everything to one full-width column per line on mobile, and the "-first"
  // modifier keeps the very first cell from getting an extra top gap.
  const style = `<style>@media screen and (max-width:480px){.${c}{max-width:100%!important;padding-left:0!important;padding-top:${gy}px!important;}.${c}-first{padding-top:0!important;}}</style>`;
  let idx = 0;
  const cellsHtml = rows.map((row) => {
    const pct = (100 / row.length).toFixed(2);
    return row.map((item, ci) => {
      const open = ci === 0
        ? `<!--[if mso]><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td width="${pct}%" valign="top"><![endif]-->`
        : `<!--[if mso]><td width="${pct}%" valign="top"><![endif]-->`;
      const close = ci === row.length - 1
        ? `<!--[if mso]></td></tr></table><![endif]-->`
        : `<!--[if mso]></td><![endif]-->`;
      const isFirst = idx === 0;
      const padLeft = ci === 0 ? '' : `padding-left:${gx}px;`;
      const padTop = (ci === 0 && idx !== 0) ? `padding-top:${gy}px;` : '';
      idx++;
      return `${open}<div class="${c}${isFirst ? ` ${c}-first` : ''}" style="display:inline-block;width:100%;max-width:${colW}px;vertical-align:top;box-sizing:border-box;${padLeft}${padTop}">${cellHtml(item)}</div>${close}`;
    }).join('');
  }).join('');
  return style + cellsHtml;
}

export const MODULES = [];

MODULES.push({
  id: 'in-this-issue', name: 'In This Issue', tags: ['Index', 'Newsletter'],
  category: 'Newsletter & Content',
  desc: 'Numbered contents index. One item per line: Title | Description.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'January Newsletter' },
    { key: 'title', label: 'Title', type: 'text', default: 'In this issue' },
    { key: 'items', label: 'Issue items', type: 'repeater', addLabel: 'Add issue item',
      parts: [
        { key: 'title', label: 'Heading', type: 'text' },
        { key: 'desc', label: 'Description', type: 'textarea' }
      ],
      default: 'A welcome for the spring term | A look at what\u2019s ahead across the department.\nThree Miners named Goldwater Scholars | National recognition for undergraduate research.\nNew makerspace opens in Toomey Hall | Reserve tools, time, and training starting now.' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', num: '#007A33', heading: '#154734', desc: '#525252', border: '#eceae6' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', num: '#007A33', heading: '#154734', desc: '#525252', border: '#e0e3da' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', num: b.accent, heading: '#ffffff', desc: 'rgba(255,255,255,0.7)', border: 'rgba(255,255,255,0.15)' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const rows = lines(v.items).map((p, i) => {
      const n = String(i + 1).padStart(2, '0');
      return `<tr><td colspan="2" style="font-size:0;line-height:0;border-top:1px solid ${pal.border};">&nbsp;</td></tr>
              <tr>
                <td width="48" valign="top" style="padding:18px 14px 18px 0;font-family:${FF};font-size:22px;font-weight:900;color:${pal.num};line-height:1.05;">${n}</td>
                <td valign="top" style="padding:18px 0;font-family:${FF};">
                  <div style="font-size:17px;font-weight:800;color:${pal.heading};line-height:1.3;letter-spacing:-0.2px;">${p[0] || ''}</div>
                  <div style="margin-top:4px;font-size:13px;line-height:1.5;color:${pal.desc};">${p[1] || ''}</div>
                </td>
              </tr>`;
    }).join('');
    const inner = `<div style="padding:38px 0 8px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:9px 0 0;font-size:26px;font-weight:900;color:${pal.title};letter-spacing:-0.5px;">${v.title}</h2>
      </div>
      <div style="padding:8px 0 34px;font-family:${FF};">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
          ${rows}
        </table>
      </div>`;
    return `<!-- ===== Missouri S&T · In This Issue ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'section-divider', name: 'Section Divider', tags: ['Structure', 'Label'],
  category: 'Newsletter & Content',
  desc: 'A labeled break to organize a long newsletter into sections.',
  fields: [
    { key: 'label', label: 'Section label (35 char max)', type: 'text', default: 'Research & Innovation', maxLength: 35 },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', label: '#154734', rule: '#e3e5e0' },
      Tint: { bg: '#f4f5f1', label: '#154734', rule: '#dfe2d9' },
      Forest: { bg: '#154734', label: '#ffffff', rule: 'rgba(255,255,255,0.25)' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const label = String(v.label || '').slice(0, 35).split(' ').join('&nbsp;');
    // Outlook desktop (Word rendering engine) unreliably honors nowrap/white-space on a
    // width:100% sibling cell and can wrap the label regardless \u2014 so Outlook gets its own
    // fixed-pixel two-column table (generous enough for the 35-char cap) via conditional
    // comments, hidden from every other client, which instead renders the fluid version.
    const swatchCells = `<td width="10" height="10" bgcolor="${b.accent}" style="width:10px;min-width:10px;height:10px;background-color:${b.accent};border-radius:2px;">&nbsp;</td>
            <td width="10" style="width:10px;min-width:10px;">&nbsp;</td>
            <td nowrap="nowrap" style="font-family:${FF};font-size:13px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;color:${pal.label};white-space:nowrap;">${label}</td>
            <td width="16" style="width:16px;min-width:16px;">&nbsp;</td>`;
    const ruleDiv = `<div style="border-top:2px solid ${pal.rule};font-size:0;line-height:0;">&nbsp;</div>`;
    const inner = `<style>@media screen and (max-width:480px){.sd-table{display:block!important;}.sd-label,.sd-rule-cell{display:block!important;width:100%!important;padding:0!important;}.sd-rule-cell{margin-top:10px;}}</style>
      <!--[if mso]>
      <table role="presentation" width="520" cellpadding="0" cellspacing="0" border="0"><tr>
        <td width="360" valign="middle">
          <table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>${swatchCells}</tr></table>
        </td>
        <td width="160" valign="middle">${ruleDiv}</td>
      </tr></table>
      <![endif]-->
      <!--[if !mso]><!-->
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="sd-table" style="border-collapse:collapse;">
      <tr>
        <td class="sd-label" valign="middle" nowrap="nowrap" style="font-family:${FF};font-size:13px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;color:${pal.label};white-space:nowrap;padding:0;">
          <table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>${swatchCells}</tr></table>
        </td>
        <td class="sd-rule-cell" valign="middle" width="100%" style="padding:0;">
          ${ruleDiv}
        </td>
      </tr>
    </table>
      <!--<![endif]-->`;
    return `<!-- ===== Missouri S&T · Section Divider ===== -->
${shell(inner, { maxWidth: 580, padY: 26, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'directory', name: 'Directory', tags: ['Contacts', 'People', 'List'],
  category: 'People & Directory',
  desc: 'A tidy contact list \u2014 names, roles, and tappable email links in one place. Ideal for a \u201cwho to reach\u201d block.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Need help?' },
    { key: 'title', label: 'Title', type: 'text', default: 'Who to contact' },
    { key: 'items', label: 'Contacts', type: 'repeater', addLabel: 'Add contact',
      parts: [
        { key: 'name', label: 'Name', type: 'text' },
        { key: 'role', label: 'Role', type: 'text' },
        { key: 'email', label: 'Email', type: 'text' }
      ],
      default: 'Dr. Lena Ortiz | Undergraduate advising | test@email.com\nJordan Blake | Financial aid & billing | test@email.com\nSam Whitfield | Housing & residential life | test@email.com' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', name: '#154734', role: '#525252', border: '#eceae6', link: '#007A33' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', name: '#154734', role: '#525252', border: '#e0e3da', link: '#007A33' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', name: '#ffffff', role: 'rgba(255,255,255,0.7)', border: 'rgba(255,255,255,0.15)', link: b.accent }
    };
    const pal = panels[v.panelBg] || panels.White;
    const rows = lines(v.items).map(p => `<tr><td style="padding:16px 0;border-top:1px solid ${pal.border};font-family:${FF};">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
                <tr>
                  <td width="38%" valign="middle" style="font-family:${FF};">
                    <div style="font-size:16px;font-weight:800;color:${pal.name};letter-spacing:-0.2px;">${p[0] || ''}</div>
                    <div style="margin-top:2px;font-size:13px;font-weight:600;color:${pal.role};">${p[1] || ''}</div>
                  </td>
                  <td width="62%" valign="middle" align="right" style="font-family:${FF};padding-left:10px;"><a href="mailto:${p[2] || ''}" style="font-size:13px;font-weight:800;color:${pal.link};text-decoration:none;white-space:nowrap;">${p[2] || ''}</a></td>
                </tr>
              </table>
            </td></tr>`).join('');
    const inner = `<div style="padding:38px 0 4px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:9px 0 0;font-size:26px;font-weight:900;color:${pal.title};letter-spacing:-0.5px;">${v.title}</h2>
      </div>
      <div style="padding:8px 0 34px;font-family:${FF};">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
          ${rows}
        </table>
      </div>`;
    return `<!-- ===== Missouri S&T · Directory ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'pull-quote', name: 'Pull Quote', tags: ['Testimonial', 'Quote'],
  category: 'Newsletter & Content',
  desc: 'A typographic quote block for testimonials, alumni, or a dean\u2019s message.',
  fields: [
    { key: 'quote', label: 'Quote', type: 'textarea', default: 'Coming to S&T changed the trajectory of my career. The faculty saw potential in me before I saw it in myself.' },
    { key: 'name', label: 'Attribution', type: 'text', default: 'Marcus Lee, \u201919' },
    { key: 'meta', label: 'Sub-attribution', type: 'text', default: 'Systems Engineer, NASA' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['Tint', 'White', 'Forest'], default: 'Tint' }
  ],
  render: (v, b) => {
    const panels = {
      Tint: { bg: '#f4f5f1', mark: b.accent, quote: '#154734', name: '#154734', meta: '#525252', rule: b.accent },
      White: { bg: '#ffffff', mark: b.accent, quote: '#154734', name: '#154734', meta: '#525252', rule: b.accent },
      Forest: { bg: '#154734', mark: b.accent, quote: '#ffffff', name: '#ffffff', meta: 'rgba(255,255,255,0.7)', rule: b.accent }
    };
    const pal = panels[v.panelBg] || panels.Tint;
    const inner = `<div style="padding:46px 0 48px;font-family:${FF};">
        <div style="font-family:Georgia,'Times New Roman',serif;font-size:58px;line-height:0.5;font-weight:700;color:${pal.mark};">&ldquo;</div>
        <p style="margin:18px 0 26px;font-size:23px;line-height:1.4;font-weight:800;color:${pal.quote};letter-spacing:-0.4px;">${v.quote}</p>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
          <tr>
            <td valign="middle" style="width:30px;padding-right:14px;"><div style="width:30px;height:3px;background-color:${pal.rule};font-size:0;line-height:0;">&nbsp;</div></td>
            <td valign="middle" style="font-family:${FF};">
              <div style="font-size:14px;font-weight:800;color:${pal.name};">${v.name}</div>
              <div style="font-size:13px;font-weight:600;color:${pal.meta};">${v.meta}</div>
            </td>
          </tr>
        </table>
      </div>`;
    return `<!-- ===== Missouri S&T · Pull Quote ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'event-ticket', name: 'Date Feature', tags: ['Event', 'Deadline', 'Date'],
  category: 'Events & Dates',
  desc: 'A bold date or number stub beside details \u2014 use it for an event with an RSVP, or a countdown deadline. The third badge line is optional (leave it blank for a deadline). Add more than one and they stack vertically.',
  fields: [
    { key: 'items', label: 'Items', type: 'repeater', addLabel: 'Add item',
      parts: [
        { key: 'day', label: 'Number / day', type: 'text' },
        { key: 'month', label: 'Unit / month', type: 'text' },
        { key: 'weekday', label: 'Third line (optional \u2014 e.g. weekday)', type: 'text' },
        { key: 'kicker', label: 'Kicker', type: 'text' },
        { key: 'title', label: 'Title', type: 'textarea' },
        { key: 'info', label: 'Detail line', type: 'text' },
        { key: 'linkText', label: 'Link text', type: 'text' },
        { key: 'linkHref', label: 'Link URL', type: 'text' }
      ],
      default: '12 | Sep | Thursday | Department Seminar | Quantum materials at room temperature | 4:00 PM  \u00b7  Schrenk Hall 120 | RSVP for this talk | mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', detail: '#56615a', link: '#007A33', dateBg: '#154734', dateNum: '#ffffff', dateUnit: b.accent },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', detail: '#56615a', link: '#007A33', dateBg: '#154734', dateNum: '#ffffff', dateUnit: b.accent },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', detail: 'rgba(255,255,255,0.75)', link: b.accent, dateBg: b.accent, dateNum: '#154734', dateUnit: '#154734' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const one = (p) => {
      const weekday = p[2] ? `<div style="margin-top:8px;font-size:11px;font-weight:600;letter-spacing:1px;text-transform:uppercase;color:${pal.dateUnit};">${p[2]}</div>` : '';
      const dateBox = `<div style="background-color:${pal.dateBg};border-radius:${b.R}px;text-align:center;padding:28px 12px;margin-bottom:20px;"><div style="font-size:42px;font-weight:900;color:${pal.dateNum};line-height:1;letter-spacing:-1px;">${p[0] || ''}</div><div style="margin-top:8px;font-size:13px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.dateUnit};">${p[1] || ''}</div>${weekday}</div>`;
      const info = `<div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${p[3] || ''}</div><h3 style="margin:8px 0 10px;font-size:20px;font-weight:900;color:${pal.title};line-height:1.22;letter-spacing:-0.3px;">${p[4] || ''}</h3><div style="font-size:13px;font-weight:600;color:${pal.detail};line-height:1.6;">${p[5] || ''}</div><div style="margin-top:14px;"><a href="${p[7] || '#'}" target="_blank" style="font-size:14px;font-weight:800;color:${pal.link};text-decoration:none;">${p[6] || ''} &rarr;</a></div>`;
      return `<div style="margin-bottom:16px;">${fluidColsGrow([{ html: dateBox, width: 180, style: `font-family:${FF};padding:0 12px;text-align:center;` }, { html: info, width: 400, style: `font-family:${FF};padding:0 12px;text-align:left;` }])}</div>`;
    };
    const tickets = lines(v.items).map(one).join('');
    return `<!-- ===== Missouri S&T · Date Feature ===== -->
${shell(`<div style="padding-top:22px;">${tickets}</div>`, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'schedule', name: 'Schedule / Agenda', tags: ['Agenda', 'Schedule'],
  category: 'Events & Dates',
  desc: 'A vertical agenda. One row per line: Time | Session | Location.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Research Conference' },
    { key: 'title', label: 'Title', type: 'text', default: 'Day one schedule' },
    { key: 'items', label: 'Sessions', type: 'repeater', addLabel: 'Add session',
      parts: [
        { key: 'time', label: 'Time', type: 'text' },
        { key: 'session', label: 'Session', type: 'text' },
        { key: 'location', label: 'Location / detail', type: 'textarea' }
      ],
      default: '9:00 AM | Opening keynote | Havener Center, St. Pat\u2019s Ballroom\n10:30 AM | Research lightning talks | Toomey Hall 199\n12:00 PM | Lunch & poster session | Havener Atrium\n2:00 PM | Industry panel | Butler-Carlton 125' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', time: '#007A33', session: '#154734', loc: '#525252', border: '#eceae6' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', time: '#007A33', session: '#154734', loc: '#525252', border: '#e0e3da' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', time: b.accent, session: '#ffffff', loc: 'rgba(255,255,255,0.7)', border: 'rgba(255,255,255,0.15)' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const rows = lines(v.items).map(p => `<tr><td colspan="2" style="font-size:0;line-height:0;border-top:1px solid ${pal.border};">&nbsp;</td></tr>
              <tr>
                <td width="96" valign="top" style="padding:16px 16px 16px 0;font-family:${FF};font-size:14px;font-weight:800;color:${pal.time};white-space:nowrap;">${p[0] || ''}</td>
                <td valign="top" style="padding:16px 0;font-family:${FF};">
                  <div style="font-size:16px;font-weight:800;color:${pal.session};line-height:1.3;letter-spacing:-0.2px;">${p[1] || ''}</div>
                  <div style="margin-top:3px;font-size:13px;color:${pal.loc};">${p[2] || ''}</div>
                </td>
              </tr>`).join('');
    const inner = `<div style="padding:38px 0 12px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:9px 0 0;font-size:26px;font-weight:900;color:${pal.title};letter-spacing:-0.5px;">${v.title}</h2>
      </div>
      <div style="padding:8px 0 38px;font-family:${FF};">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
          ${rows}
        </table>
      </div>`;
    return `<!-- ===== Missouri S&T · Schedule / Agenda ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'table', name: 'Data Table', tags: ['Table', 'Comparison', 'Data'],
  category: 'Newsletter & Content',
  desc: 'An email-safe table with a forest header row and zebra striping \u2014 for deadlines, tuition, or any tidy comparison.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'At a glance' },
    { key: 'title', label: 'Title', type: 'text', default: 'Dates & fees by program' },
    { key: 'col1', label: 'Column 1 header', type: 'text', default: 'Program' },
    { key: 'col2', label: 'Column 2 header', type: 'text', default: 'Deadline' },
    { key: 'col3', label: 'Column 3 header', type: 'text', default: 'Fee' },
    { key: 'items', label: 'Rows', type: 'repeater', addLabel: 'Add row',
      parts: [
        { key: 'c1', label: 'Column 1', type: 'text' },
        { key: 'c2', label: 'Column 2', type: 'text' },
        { key: 'c3', label: 'Column 3', type: 'text' }
      ],
      default: 'B.S. Chemistry | March 1 | $50\nM.S. Chemistry | Feb 1 | $65\nPh.D. Chemistry | Dec 15 | $65' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', border: '#e3e5e0', zebra: '#f7f8f5', row: '#ffffff' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', border: '#dfe2d9', zebra: '#ffffff', row: '#fbfcfa' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', border: 'rgba(255,255,255,0.2)', zebra: 'rgba(255,255,255,0.06)', row: 'transparent', headBg: b.accent, headText: '#154734' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const isForest = v.panelBg === 'Forest';
    const headBg = pal.headBg || '#154734';
    const headText = pal.headText || '#ffffff';
    const cols = [v.col1 || '', v.col2 || '', v.col3 || ''];
    const head = cols.map((c, ci) => `<td bgcolor="${headBg}" width="${ci === 0 ? '42%' : '29%'}" style="background-color:${headBg};padding:13px 10px;font-family:${FF};font-size:11px;font-weight:800;letter-spacing:0.4px;text-transform:uppercase;color:${headText};text-align:left;overflow-wrap:break-word;">${c}</td>`).join('');
    const body = lines(v.items).map((r, ri) => `<tr>${cols.map((_, ci) => `<td style="padding:14px 10px;font-family:${FF};font-size:13px;font-weight:${ci === 0 ? '800' : '600'};color:${ci === 0 ? (isForest ? '#ffffff' : '#154734') : (isForest ? 'rgba(255,255,255,0.75)' : '#525252')};border-top:1px solid ${pal.border};background-color:${ri % 2 ? pal.zebra : pal.row};overflow-wrap:break-word;">${r[ci] || ''}</td>`).join('')}</tr>`).join('');
    const inner = `<div style="padding:38px 0 4px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:9px 0 0;font-size:26px;font-weight:900;color:${pal.title};letter-spacing:-0.5px;">${v.title}</h2>
      </div>
      <div style="padding:18px 0 38px;font-family:${FF};">
        <div style="border:1px solid ${pal.border};border-radius:${b.R}px;overflow:hidden;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;width:100%;table-layout:fixed;">
            <tr>${head}</tr>
            ${body}
          </table>
        </div>
      </div>`;
    return `<!-- ===== Missouri S&T · Data Table ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'timeline', name: 'Timeline (Simple List)', tags: ['Milestones', 'Process', 'Dates'],
  category: 'Events & Dates',
  desc: 'A dated list of milestones for a process, roadmap, or the key dates leading up to an event \u2014 built without absolute positioning so it renders identically in every email client, including Outlook.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'How to apply' },
    { key: 'title', label: 'Title', type: 'text', default: 'From interest to enrolled' },
    { key: 'items', label: 'Milestones', type: 'repeater', addLabel: 'Add milestone',
      parts: [
        { key: 'date', label: 'Date / step', type: 'text' },
        { key: 'title', label: 'Title', type: 'text' },
        { key: 'desc', label: 'Description', type: 'textarea' }
      ],
      default: 'Jan 15 | Submit your application | Free to apply \u2014 no fee for Missouri residents.\nFeb 1 | Send transcripts | Official records due to Admissions.\nMar 1 | Decisions released | Watch your inbox and applicant portal.\nMay 1 | Confirm your spot | Reserve your place with an enrollment deposit.' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', date: '#007A33', text: '#154734', desc: '#525252', border: '#eceae6' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', date: '#007A33', text: '#154734', desc: '#525252', border: '#e0e3da' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', date: b.accent, text: '#ffffff', desc: 'rgba(255,255,255,0.7)', border: 'rgba(255,255,255,0.15)' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const items = lines(v.items);
    const head = `<div style="padding:38px 0 12px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:9px 0 0;font-size:26px;font-weight:900;color:${pal.title};letter-spacing:-0.5px;">${v.title}</h2>
      </div>`;
    const rows = items.map(p => `<div style="padding:14px 0;border-top:1px solid ${pal.border};font-family:${FF};"><span style="font-size:12px;font-weight:800;color:${pal.date};">${p[0] || ''}</span> &mdash; <strong style="color:${pal.text};font-size:14px;">${p[1] || ''}.</strong> <span style="font-size:13px;color:${pal.desc};">${p[2] || ''}</span></div>`).join('');
    const inner = `${head}<div style="padding:6px 0 38px;">${rows}</div>`;
    return `<!-- ===== Missouri S&T · Timeline (Simple List) ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'timeline-cards', name: 'Timeline (Accent Cards)', tags: ['Milestones', 'Process', 'Dates'],
  category: 'Events & Dates',
  desc: 'A dated list of milestones as stacked accent cards \u2014 built without absolute positioning so it renders identically in every email client, including Outlook.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'How to apply' },
    { key: 'title', label: 'Title', type: 'text', default: 'From interest to enrolled' },
    { key: 'items', label: 'Milestones', type: 'repeater', addLabel: 'Add milestone',
      parts: [
        { key: 'date', label: 'Date / step', type: 'text' },
        { key: 'title', label: 'Title', type: 'text' },
        { key: 'desc', label: 'Description', type: 'textarea' }
      ],
      default: 'Jan 15 | Submit your application | Free to apply \u2014 no fee for Missouri residents.\nFeb 1 | Send transcripts | Official records due to Admissions.\nMar 1 | Decisions released | Watch your inbox and applicant portal.\nMay 1 | Confirm your spot | Reserve your place with an enrollment deposit.' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', cardBg: '#f4f5f1', accent: '#007A33', cardTitle: '#154734', desc: '#525252' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', cardBg: '#ffffff', accent: '#007A33', cardTitle: '#154734', desc: '#525252' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', cardBg: 'rgba(255,255,255,0.08)', accent: b.accent, cardTitle: '#ffffff', desc: 'rgba(255,255,255,0.7)' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const items = lines(v.items);
    const head = `<div style="padding:38px 0 12px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:9px 0 0;font-size:26px;font-weight:900;color:${pal.title};letter-spacing:-0.5px;">${v.title}</h2>
      </div>`;
    const cards = items.map(p => `<div style="background-color:${pal.cardBg};border-left:4px solid ${pal.accent};border-radius:4px;padding:14px 18px;margin-top:10px;font-family:${FF};">
        <div style="font-size:12px;font-weight:800;color:${pal.accent};">${p[0] || ''}</div>
        <div style="margin-top:2px;font-size:15px;font-weight:800;color:${pal.cardTitle};">${p[1] || ''}</div>
        <div style="margin-top:3px;font-size:13px;color:${pal.desc};">${p[2] || ''}</div>
      </div>`).join('');
    const inner = `${head}<div style="padding:8px 0 38px;">${cards}</div>`;
    return `<!-- ===== Missouri S&T · Timeline (Accent Cards) ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'quick-links', name: 'Quick Links', tags: ['Links', 'Resources'],
  category: 'Info & Resources',
  desc: 'A grid of tappable links, sized to fit each label. Add a link and its URL as separate fields below.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Quick links' },
    { key: 'title', label: 'Title', type: 'text', default: 'Get things done' },
    { key: 'items', label: 'Links', type: 'repeater', addLabel: 'Add link',
      parts: [
        { key: 'label', label: 'Label', type: 'text' },
        { key: 'url', label: 'URL', type: 'text' }
      ],
      default: 'Register for classes | mst.edu\nPay your bill | mst.edu\nBook an advisor | mst.edu\nFind campus jobs | mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['Tint', 'White', 'Forest'], default: 'Tint' }
  ],
  render: (v, b) => {
    const panels = {
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', linkBg: '#ffffff', linkBorder: '#e3e5e0', linkText: '#154734', arrow: '#007A33' },
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', linkBg: '#f4f5f1', linkBorder: '#e3e5e0', linkText: '#154734', arrow: '#007A33' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', linkBg: 'rgba(255,255,255,0.08)', linkBorder: 'rgba(255,255,255,0.2)', linkText: '#ffffff', arrow: b.accent }
    };
    const pal = panels[v.panelBg] || panels.Tint;
    const link = (p) => `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:separate;margin-bottom:10px;">
                  <tr><td style="background-color:${pal.linkBg};border:1px solid ${pal.linkBorder};border-radius:${b.R}px;padding:0;">
                    <a href="${p[1] || '#'}" target="_blank" style="display:block;font-family:${FF};font-size:14px;font-weight:700;color:${pal.linkText};text-decoration:none;padding:15px 16px;text-align:left;">${p[0] || ''} <span style="color:${pal.arrow};">&rarr;</span></a>
                  </td></tr>
                </table>`;
    const all = lines(v.items);
    const left = all.filter((_, i) => i % 2 === 0).map(link).join('');
    const right = all.filter((_, i) => i % 2 === 1).map(link).join('');
    const head = `<div style="padding:34px 0 6px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:8px 0 0;font-size:23px;font-weight:900;color:${pal.title};letter-spacing:-0.4px;">${v.title}</h2>
      </div>`;
    const cols = twoCol(left, 290, `font-family:${FF};padding:0 12px;`, right, 290, `font-family:${FF};padding:0 12px;`);
    return `<!-- ===== Missouri S&T · Quick Links ===== -->
${shell(`${head}<div style="padding:16px 0 34px;text-align:center;">${cols}</div>`, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'faq', name: 'Questions & Answers', tags: ['FAQ', 'Support', 'List'],
  category: 'Info & Resources',
  desc: 'A stacked question-and-answer list for the things people always ask \u2014 clean, scannable, and email-safe with no accordion needed.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Good to know' },
    { key: 'title', label: 'Title', type: 'text', default: 'Questions, answered' },
    { key: 'items', label: 'Questions', type: 'repeater', addLabel: 'Add question',
      parts: [
        { key: 'q', label: 'Question', type: 'text' },
        { key: 'a', label: 'Answer', type: 'textarea' }
      ],
      default: 'Do I need to register for orientation? | Yes \u2014 pick a session in the portal before June 1. Sessions fill quickly.\nCan I bring a guest? | Each student may bring up to two family members at no cost.\nWhere do I park? | Visitor lots off Bishop Ave are free on orientation days; signs will direct you.' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', q: '#007A33', question: '#154734', answer: '#525252', border: '#eceae6' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', q: '#007A33', question: '#154734', answer: '#525252', border: '#e0e3da' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', q: b.accent, question: '#ffffff', answer: 'rgba(255,255,255,0.75)', border: 'rgba(255,255,255,0.15)' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const rows = lines(v.items).map(p => `<tr><td style="padding:20px 0;border-top:1px solid ${pal.border};font-family:${FF};">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
                <tr>
                  <td width="30" valign="top" style="font-family:${FF};font-size:16px;font-weight:900;color:${pal.q};line-height:1.35;">Q</td>
                  <td valign="top" style="font-family:${FF};font-size:16px;font-weight:800;color:${pal.question};line-height:1.35;letter-spacing:-0.2px;">${p[0] || ''}</td>
                </tr>
              </table>
              <div style="margin:8px 0 0 30px;font-family:${FF};font-size:14px;line-height:1.6;color:${pal.answer};">${p[1] || ''}</div>
            </td></tr>`).join('');
    const inner = `<div style="padding:38px 0 4px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:9px 0 0;font-size:26px;font-weight:900;color:${pal.title};letter-spacing:-0.5px;">${v.title}</h2>
      </div>
      <div style="padding:8px 0 34px;font-family:${FF};">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
          ${rows}
        </table>
      </div>`;
    return `<!-- ===== Missouri S&T · Questions & Answers ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'save-the-date', name: 'Save the Date', tags: ['Date', 'Feature'],
  category: 'Events & Dates',
  desc: 'A bold typographic date feature on deep forest.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Save the date' },
    { key: 'date', label: 'Date', type: 'text', default: 'April 18, 2026' },
    { key: 'meta', label: 'Detail line', type: 'text', default: 'Spring Commencement  ·  Gale Bullman Arena' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'Add to calendar' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: '' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['Forest', 'White', 'Lime'], default: 'Forest' }
  ],
  render: (v, b) => {
    const panels = {
      Forest: { bg: '#154734', kicker: b.accent, date: '#ffffff', meta: 'rgba(255,255,255,0.85)', link: b.accent },
      White: { bg: '#ffffff', kicker: '#007A33', date: '#154734', meta: '#525252', link: '#007A33' },
      Lime: { bg: b.accent, kicker: '#154734', date: '#154734', meta: 'rgba(21,71,52,0.8)', link: '#154734' }
    };
    const pal = panels[v.panelBg] || panels.Forest;
    const inner = `<div style="text-align:center;padding:46px 0 48px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <div style="margin-top:14px;font-size:46px;font-weight:900;color:${pal.date};letter-spacing:-1.5px;line-height:1;">${v.date}</div>
        <div style="margin-top:15px;font-size:15px;font-weight:600;color:${pal.meta};letter-spacing:0.2px;">${v.meta}</div>
        <div style="margin-top:24px;"><a href="${v.linkHref}" target="_blank" style="font-family:${FF};font-size:15px;font-weight:800;color:${pal.link};text-decoration:none;">${v.linkText} &rarr;</a></div>
      </div>`;
    return `<!-- ===== Missouri S&T · Save the Date ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'by-the-numbers', name: 'By the Numbers', tags: ['Stats', 'Impact'],
  category: 'Newsletter & Content',
  desc: 'A stat band on deep forest. One stat per line: Number | Label.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'By the numbers' },
    { key: 'items', label: 'Stats', type: 'repeater', addLabel: 'Add stat',
      parts: [
        { key: 'number', label: 'Number', type: 'text' },
        { key: 'label', label: 'Label', type: 'text' }
      ],
      default: '1,200+ | Undergraduate researchers\n94% | Career placement rate\n#7 | Best value public, Midwest' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['Forest', 'White', 'Tint'], default: 'Forest' }
  ],
  render: (v, b) => {
    const panels = {
      Forest: { bg: '#154734', kicker: b.accent, num: '#ffffff', label: b.accent },
      White: { bg: '#ffffff', kicker: '#007A33', num: '#154734', label: '#007A33' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', num: '#154734', label: '#007A33' }
    };
    const pal = panels[v.panelBg] || panels.Forest;
    const cells = lines(v.items).slice(0, 3).map(p => `<div style="text-align:center;padding:16px 20px;font-family:${FF};"><div style="font-size:38px;font-weight:900;color:${pal.num};line-height:1;letter-spacing:-1.5px;font-family:${FF};">${p[0] || ''}</div><div style="margin-top:11px;font-size:12px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:${pal.label};line-height:1.45;font-family:${FF};">${p[1] || ''}</div></div>`);
    while (cells.length < 3) cells.push('<div>&nbsp;</div>');
    const head = `<div style="text-align:center;padding:42px 0 4px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
      </div>`;
    const cols = fluidCols([
      { html: cells[0], width: 194 },
      { html: cells[1], width: 194 },
      { html: cells[2], width: 194 }
    ]);
    return `<!-- ===== Missouri S&T · By the Numbers ===== -->
${shell(`${head}<div style="padding:16px 0 42px;text-align:center;">${cols}</div>`, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'donation-goal', name: 'Donation Goal', tags: ['Fundraising', 'Progress'],
  category: 'Giving & Fundraising',
  desc: 'A fundraising progress block with an email-safe thermometer bar.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Annual Fund' },
    { key: 'title', label: 'Title  ·  {pct} auto-fills the percent', type: 'text', default: 'We\u2019re {pct}% to our goal' },
    { key: 'raised', label: 'Raised amount', type: 'text', default: '$340,000' },
    { key: 'goal', label: 'Goal amount', type: 'text', default: '$500,000' },
    { key: 'donorLine', label: 'Donor line', type: 'text', default: '1,204 donors and counting \u2014 help us close the gap.' },
    { key: 'buttonText', label: 'Button text', type: 'text', default: 'Give now' },
    { key: 'buttonHref', label: 'Button URL', type: 'text', default: '' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', raised: '#154734', goal: '#525252', track: '#e6e9e3', bar: b.btn, donor: '#525252', btnBg: b.btn, btnText: '#ffffff' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', raised: '#154734', goal: '#525252', track: '#e6e9e3', bar: b.btn, donor: '#525252', btnBg: b.btn, btnText: '#ffffff' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', raised: '#ffffff', goal: 'rgba(255,255,255,0.7)', track: 'rgba(255,255,255,0.2)', bar: b.accent, donor: 'rgba(255,255,255,0.75)', btnBg: b.accent, btnText: '#154734' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const num = (s) => parseFloat(String(s || '').replace(/[^0-9.]/g, '')) || 0;
    const pct = num(v.goal) > 0 ? Math.max(0, Math.min(100, Math.round(num(v.raised) / num(v.goal) * 100))) : 0;
    const title = String(v.title || '').replace(/\{pct\}/g, pct);
    const inner = `<div style="padding:38px 0 40px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:9px 0 20px;font-size:25px;font-weight:900;color:${pal.title};letter-spacing:-0.5px;">${title}</h2>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
          <tr>
            <td align="left" style="font-family:${FF};font-size:14px;font-weight:800;color:${pal.raised};padding-bottom:8px;">${v.raised} raised</td>
            <td align="right" style="font-family:${FF};font-size:13px;font-weight:600;color:${pal.goal};padding-bottom:8px;">Goal ${v.goal}</td>
          </tr>
        </table>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:separate;">
          <tr><td bgcolor="${pal.track}" style="background-color:${pal.track};border-radius:${b.barR}px;padding:0;font-size:0;line-height:0;">
            <table role="presentation" width="${pct}%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:separate;">
              <tr><td bgcolor="${pal.bar}" style="background-color:${pal.bar};border-radius:${b.barR}px;height:16px;font-size:0;line-height:0;">&nbsp;</td></tr>
            </table>
          </td></tr>
        </table>
        <div style="margin:16px 0 24px;font-family:${FF};font-size:13px;font-weight:600;color:${pal.donor};">${v.donorLine}</div>
        ${mkBtn({ ...b, R: 25, arc: '52%' }, v.buttonText, v.buttonHref, pal.btnBg, pal.btnText)}
      </div>`;
    return `<!-- ===== Missouri S&T · Donation Goal ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'cta', name: 'Call to Action', tags: ['CTA', 'Button'],
  category: 'Newsletter & Content',
  desc: 'A simple centered prompt with a button \u2014 a clean, all-purpose call to action. Pick a color combo to match the moment (a light panel, deep forest, or bold gold).',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Ready?' },
    { key: 'headline', label: 'Headline', type: 'textarea', default: 'Start your application today' },
    { key: 'body', label: 'Body', type: 'textarea', default: 'It takes about 20 minutes, and there\u2019s no fee for Missouri residents.' },
    { key: 'buttonText', label: 'Button text', type: 'text', default: 'Apply now' },
    { key: 'buttonHref', label: 'Button URL', type: 'text', default: 'mst.edu' },
    { key: 'palette', label: 'Color combo', type: 'select', options: ['Light', 'Forest', 'Lime'], default: 'Light' }
  ],
  render: (v, b) => {
    const palettes = {
      Light: { bg: '#f4f5f1', kicker: '#007A33', headline: '#154734', body: '#4a554c', btnBg: b.btn, btnText: '#ffffff' },
      Forest: { bg: '#154734', kicker: b.accent, headline: '#ffffff', body: 'rgba(255,255,255,0.85)', btnBg: b.accent, btnText: '#154734' },
      Lime: { bg: b.accent, kicker: '#154734', headline: '#154734', body: '#154734', btnBg: '#154734', btnText: '#ffffff' }
    };
    const pal = palettes[v.palette] || palettes.Light;
    const inner = `<div style="text-align:center;padding:44px 0 48px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:13px 0 0;font-size:27px;line-height:1.18;font-weight:900;color:${pal.headline};letter-spacing:-0.5px;">${v.headline}</h2>
        <p style="margin:14px 0 28px;font-size:16px;line-height:1.6;color:${pal.body};">${v.body}</p>
        ${mkBtn(b, v.buttonText, v.buttonHref, pal.btnBg, pal.btnText)}
      </div>`;
    return `<!-- ===== Missouri S&T · Call to Action ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'announcement', name: 'Announcement Banner', tags: ['Alert', 'Accent', 'Deadline'],
  category: 'Events & Dates',
  desc: 'A bold accent strip for one time-sensitive message — a deadline, a closure, a call to act. Impossible to miss.',
  fields: [
    { key: 'label', label: 'Label', type: 'text', default: 'Heads up' },
    { key: 'message', label: 'Message', type: 'textarea', default: 'Priority registration for fall closes Friday, March 28.' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'Register now' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: 'mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['Lime', 'Forest', 'White'], default: 'Lime' }
  ],
  render: (v, b) => {
    const panels = {
      Lime: { bg: b.accent, badgeBg: '#154734', badgeText: '#ffffff', text: '#154734', link: '#154734' },
      Forest: { bg: '#154734', badgeBg: b.accent, badgeText: '#154734', text: '#ffffff', link: '#ffffff' },
      White: { bg: '#ffffff', badgeBg: '#154734', badgeText: '#ffffff', text: '#154734', link: '#154734' }
    };
    const pal = panels[v.panelBg] || panels.Lime;
    const inner = `<div style="padding:24px 0 26px;font-family:${FF};">
        <span style="display:inline-block;background-color:${pal.badgeBg};color:${pal.badgeText};font-size:11px;font-weight:800;letter-spacing:1.2px;text-transform:uppercase;padding:5px 12px;border-radius:${Math.max(b.R, 20)}px;">${v.label}</span>
        <div style="margin-top:13px;font-size:18px;font-weight:800;color:${pal.text};letter-spacing:-0.3px;line-height:1.4;">${v.message}</div>
        <div style="margin-top:13px;"><a href="${v.linkHref}" target="_blank" style="font-family:${FF};font-size:14px;font-weight:800;color:${pal.link};text-decoration:underline;">${v.linkText} &rarr;</a></div>
      </div>`;
    return `<!-- ===== Missouri S&T · Announcement Banner ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'photo-textbox', name: 'Full-Bleed Photo + Textbox', tags: ['Photo', 'Hero', 'Overlay'],
  category: 'Giving & Fundraising',
  fullBleed: true,
  desc: 'A full-width photo with a white textbox overlaid near the bottom \u2014 paste your own hosted photo URL. Falls back to a plain color card if left blank.',
  fields: [
    { key: 'photoUrl', label: 'Photo URL (paste a hosted image)', type: 'text', default: '' },
    { key: 'photoAlt', label: 'Photo alt text', type: 'text', default: 'Students celebrating on campus during Giving Day' },
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Giving Day 2026' },
    { key: 'title', label: 'Title', type: 'textarea', default: 'Every gift is matched, dollar for dollar' },
    { key: 'body', label: 'Body', type: 'textarea', default: 'A group of alumni will match new gifts up to $50,000 today only.' },
    { key: 'buttonText', label: 'Button text', type: 'text', default: 'Give now' },
    { key: 'buttonHref', label: 'Button URL', type: 'text', default: '' },
    { key: 'boxColor', label: 'Textbox color', type: 'select', options: ['White', 'Forest', 'Lime'], default: 'White' }
  ],
  render: (v, b) => {
    const bg = (v.photoUrl || '').trim();
    const bgStyle = bg ? `background-image:url('${bg}');background-size:cover;background-position:center;` : `background-color:#154734;`;
    const boxPalettes = {
      White: { boxBg: '#ffffff', kicker: '#007A33', title: '#154734', body: '#4a554c', btnBg: b.btn, btnText: '#ffffff' },
      Forest: { boxBg: '#154734', kicker: b.accent, title: '#ffffff', body: 'rgba(255,255,255,0.85)', btnBg: b.accent, btnText: '#154734' },
      Lime: { boxBg: b.accent, kicker: '#154734', title: '#154734', body: '#154734', btnBg: '#154734', btnText: '#ffffff' }
    };
    const boxPal = boxPalettes[v.boxColor] || boxPalettes.White;
    const box = `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr><td bgcolor="${boxPal.boxBg}" style="background-color:${boxPal.boxBg};border-radius:${b.R}px;padding:28px 30px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${boxPal.kicker};">${v.kicker}</div>
        <h2 style="margin:12px 0 0;font-size:24px;font-weight:900;color:${boxPal.title};letter-spacing:-0.4px;line-height:1.2;max-width:380px;">${v.title}</h2>
        <p style="margin:14px 0 0;font-size:14px;line-height:1.6;color:${boxPal.body};max-width:380px;">${v.body}</p>
        <div style="margin-top:18px;">${mkBtn({ ...b, R: 25, arc: '52%' }, v.buttonText, v.buttonHref, boxPal.btnBg, boxPal.btnText)}</div>
      </td></tr></table>`;
    // No shell()/580px cap here on purpose — the photo is meant to run the full width of the
    // message pane (wider than the fixed-width body other modules use), not just to 580px.
    // role="img" + aria-label carries the alt text since a CSS background-image has no native alt.
    // Outlook desktop (Windows, Word rendering engine) ignores CSS background-image entirely —
    // the VML v:rect/v:fill block is the standard bulletproof fallback: only MSO parses the
    // conditional comments, so it paints the same photo full-bleed there while every other
    // client (which doesn't understand VML) just uses the CSS background-image as normal.
    const vmlOpen = bg
      ? `<!--[if mso]><v:rect xmlns:v="urn:schemas-microsoft-com:vml" fill="true" stroke="false" style="width:600px;"><v:fill type="frame" src="${bg}" color="#154734" /><v:textbox inset="0,0,0,0"><![endif]-->`
      : '';
    const vmlClose = bg ? `<!--[if mso]></v:textbox></v:rect><![endif]-->` : '';
    const inner = `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr><td role="img" aria-label="${v.photoAlt || ''}" style="${bgStyle}padding:280px 24px 24px;">${vmlOpen}${box}${vmlClose}</td></tr></table>`;
    return `<!-- ===== Missouri S&T · Full-Bleed Photo + Textbox ===== -->
${inner}`;
  }
});

MODULES.push({
  id: 'webinar', name: 'Virtual Event / Webinar', tags: ['Webinar', 'Event', 'Zoom'],
  category: 'Events & Dates',
  desc: 'A live webinar or virtual info session card with a photo placeholder and RSVP link.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Live Webinar' },
    { key: 'title', label: 'Title', type: 'textarea', default: 'Grad school 101: funding your degree' },
    { key: 'meta', label: 'Date / time / platform', type: 'text', default: 'Tues, Aug 4  \u00b7  6:00 PM CT  \u00b7  Zoom' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'Save your seat' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: 'mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', meta: '#525252', link: '#007A33' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', meta: '#525252', link: '#007A33' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', meta: 'rgba(255,255,255,0.75)', link: b.accent }
    };
    const pal = panels[v.panelBg] || panels.White;
    const inner = `<div style="font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:${pal.kicker};font-family:${FF};">${v.kicker}</div>
      <div style="margin-top:6px;font-size:19px;font-weight:800;color:${pal.title};line-height:1.3;font-family:${FF};">${v.title}</div>
      <div style="margin-top:4px;font-size:13px;color:${pal.meta};font-family:${FF};">${v.meta}</div>
      <div style="margin-top:14px;"><a href="${v.linkHref}" target="_blank" style="font-size:14px;font-weight:800;color:${pal.link};text-decoration:none;font-family:${FF};">${v.linkText} &rarr;</a></div>`;
    return `<!-- ===== Missouri S&T · Virtual Event / Webinar ===== -->
${shell(`<div style="padding:30px 0;">${inner}</div>`, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'campus-alert', name: 'Alert Bar', tags: ['Alert', 'Safety', 'Weather'],
  category: 'Events & Dates',
  desc: 'A high-contrast alert strip for weather closures, safety notices, or urgent operational changes.',
  fields: [
    { key: 'severity', label: 'Severity', type: 'select', options: ['Red', 'Kiwi', 'Green'], default: 'Red' },
    { key: 'label', label: 'Label', type: 'text', default: 'Weather Advisory' },
    { key: 'message', label: 'Message', type: 'textarea', default: 'Evening classes canceled \u2014 campus closes at 5pm.' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'Details' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: 'mst.edu' }
  ],
  render: (v, b) => {
    const palettes = {
      Red: { bg: '#B3261E', badgeBg: '#ffffff', badgeText: '#B3261E', linkText: '#B3261E', word: 'Urgent' },
      Kiwi: { bg: b.accent, badgeBg: '#154734', badgeText: b.accent, linkText: '#154734', word: 'Caution', textColor: '#154734' },
      Green: { bg: '#154734', badgeBg: b.accent, badgeText: '#154734', linkText: '#154734', word: 'Notice' }
    };
    const pal = palettes[v.severity] || palettes.Red;
    const msgColor = pal.textColor || '#ffffff';
    const badge = `<span style="display:inline-block;background-color:${pal.badgeBg};color:${pal.badgeText};font-size:10px;font-weight:800;letter-spacing:1px;text-transform:uppercase;padding:3px 8px;border-radius:3px;font-family:${FF};">${pal.word}</span>`;
    const text = `<div>${badge} <span style="font-size:11px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;color:${msgColor};font-family:${FF};vertical-align:middle;">${v.label}</span></div>
      <div style="margin-top:8px;font-size:15px;font-weight:800;color:${msgColor};line-height:1.4;font-family:${FF};">${v.message}</div>`;
    const link = `<div style="margin-top:16px;"><a href="${v.linkHref}" target="_blank" style="display:inline-block;font-size:13px;font-weight:800;color:${pal.linkText};text-decoration:none;background-color:#ffffff;padding:10px 18px;border-radius:20px;white-space:nowrap;font-family:${FF};">${v.linkText} &rarr;</a></div>`;
    const inner = twoCol(text, 380, `font-family:${FF};padding:0 16px 0 0;`, link, 120, `padding:0;`, 'middle');
    return `<!-- ===== Missouri S&T · Campus Alert Bar ===== -->
${shell(`<div style="padding:22px 0;">${inner}</div>`, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'progress-ring', name: 'Milestone Stepper', tags: ['Fundraising', 'Milestones', 'Goal'],
  category: 'Giving & Fundraising',
  desc: 'A campaign with several funding tiers, shown as filled/unfilled step segments instead of one continuous bar \u2014 use it when a gift unlocks named stretch goals along the way. One row per line: Milestone amount | Label.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'New Robotics Lab' },
    { key: 'raised', label: 'Raised amount', type: 'text', default: '$340,000' },
    { key: 'milestones', label: 'Milestones', type: 'repeater', addLabel: 'Add milestone',
      parts: [
        { key: 'amount', label: 'Amount', type: 'text' },
        { key: 'label', label: 'Label', type: 'text' }
      ],
      default: '$150,000 | Design complete\n$300,000 | Equipment ordered\n$450,000 | Construction begins\n$500,000 | Grand opening' },
    { key: 'body', label: 'Body', type: 'textarea', default: 'Help us close the gap before the semester starts.' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'Help close the gap' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: 'mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['Forest', 'White', 'Tint'], default: 'Forest' }
  ],
  render: (v, b) => {
    const panels = {
      Forest: { bg: '#154734', kicker: b.accent, raised: '#ffffff', segOn: b.accent, segOff: 'rgba(255,255,255,0.18)', labelOn: '#ffffff', labelOff: 'rgba(255,255,255,0.55)', descOn: 'rgba(255,255,255,0.85)', descOff: 'rgba(255,255,255,0.4)', body: 'rgba(255,255,255,0.85)', link: '#ffffff' },
      White: { bg: '#ffffff', kicker: '#007A33', raised: '#154734', segOn: b.btn, segOff: '#e6e9e3', labelOn: '#154734', labelOff: '#9aa398', descOn: '#525252', descOff: '#c2c8bd', body: '#4a554c', link: '#154734' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', raised: '#154734', segOn: b.btn, segOff: '#e0e3da', labelOn: '#154734', labelOff: '#9aa398', descOn: '#525252', descOff: '#c2c8bd', body: '#4a554c', link: '#154734' }
    };
    const pal = panels[v.panelBg] || panels.Forest;
    const num = (s) => parseFloat(String(s || '').replace(/[^0-9.]/g, '')) || 0;
    const raised = num(v.raised);
    const stones = lines(v.milestones);
    const n = Math.max(1, stones.length);
    const segWidth = Math.max(90, Math.round(560 / n));
    const cols = fluidCols(stones.map((p) => {
      const reached = raised >= num(p[0]);
      const html = `<div style="height:10px;border-radius:5px;background-color:${reached ? pal.segOn : pal.segOff};font-size:0;line-height:0;">&nbsp;</div>
        <div style="margin-top:8px;font-family:${FF};">
          <div style="font-size:12px;font-weight:800;color:${reached ? pal.labelOn : pal.labelOff};">${p[0] || ''}</div>
          <div style="margin-top:2px;font-size:10.5px;color:${reached ? pal.descOn : pal.descOff};line-height:1.4;">${p[1] || ''}</div>
        </div>`;
      return { html, width: segWidth, style: 'padding:0 8px 14px 0;' };
    }));
    const inner = `<div style="padding:36px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <div style="margin-top:10px;font-size:30px;font-weight:900;color:${pal.raised};letter-spacing:-0.8px;line-height:1.15;">${v.raised} raised so far</div>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin-top:20px;"><tr><td>${cols}</td></tr></table>
        <div style="margin-top:18px;font-size:14px;color:${pal.body};line-height:1.5;">${v.body}</div>
        <div style="margin-top:16px;"><a href="${v.linkHref}" target="_blank" style="font-size:13px;font-weight:800;color:${pal.link};text-decoration:none;border-bottom:2px solid ${pal.segOn};padding-bottom:2px;">${v.linkText} &rarr;</a></div>
      </div>`;
    return `<!-- ===== Missouri S&T · Milestone Stepper ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'gift-allocation', name: 'Where Your Gift Goes', tags: ['Fundraising', 'Breakdown'],
  category: 'Giving & Fundraising',
  desc: 'A segmented allocation bar showing how a gift is split across funds.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Fund Allocation' },
    { key: 'title', label: 'Title', type: 'text', default: 'Where a $100 gift goes' },
    { key: 'items', label: 'Allocation rows', type: 'repeater', addLabel: 'Add row',
      parts: [
        { key: 'label', label: 'Fund label', type: 'text' },
        { key: 'amount', label: 'Amount', type: 'text' },
        { key: 'pct', label: '% of bar width (0\u2013100)', type: 'text' }
      ],
      default: 'Student scholarships | $52 | 52\nResearch equipment | $28 | 28\nFaculty support | $12 | 12\nCampus operations | $8 | 8' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'See the full breakdown' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: 'mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', legend: '#4a554c', amount: '#154734', link: '#007A33', swatches: ['#154734', '#007A33', b.accent, '#cccccc'] },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', legend: '#4a554c', amount: '#154734', link: '#007A33', swatches: ['#154734', '#007A33', b.accent, '#a8afa1'] },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', legend: 'rgba(255,255,255,0.8)', amount: '#ffffff', link: b.accent, swatches: [b.accent, '#72BF44', '#ffffff', 'rgba(255,255,255,0.35)'] }
    };
    const pal = panels[v.panelBg] || panels.White;
    const swatches = pal.swatches;
    const rows = lines(v.items);
    const bar = rows.map((p, i) => `<td width="${p[2] || 0}%" bgcolor="${swatches[i % swatches.length]}" style="background-color:${swatches[i % swatches.length]};font-size:0;line-height:0;height:14px;">&nbsp;</td>`).join('');
    const legend = rows.map((p, i) => `<div style="display:table;width:100%;font-family:${FF};font-size:13px;color:${pal.legend};margin-top:8px;"><div style="display:table-cell;"><span style="display:inline-block;width:9px;height:9px;border-radius:2px;background-color:${swatches[i % swatches.length]};margin-right:8px;">&nbsp;</span>${p[0] || ''}</div><div style="display:table-cell;text-align:right;font-weight:700;color:${pal.amount};white-space:nowrap;">${p[1] || ''}</div></div>`).join('');
    const inner = `<div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:8px 0 0;font-size:22px;font-weight:900;color:${pal.title};letter-spacing:-0.3px;">${v.title}</h2>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:separate;margin-top:18px;border-radius:7px;overflow:hidden;"><tr>${bar}</tr></table>
        <div style="margin-top:16px;">${legend}</div>
        <div style="margin-top:16px;"><a href="${v.linkHref}" target="_blank" style="font-size:13px;font-weight:800;color:${pal.link};text-decoration:none;">${v.linkText} &rarr;</a></div>
      </div>`;
    return `<!-- ===== Missouri S&T · Where Your Gift Goes ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'giving-levels', name: 'Giving Levels', tags: ['Fundraising', 'Tiers'],
  category: 'Giving & Fundraising',
  desc: 'Three side-by-side giving tiers with the middle level highlighted \u2014 for a membership or recurring-gift program.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Ways to Give' },
    { key: 'title', label: 'Title', type: 'text', default: 'Pick the level that\u2019s right for you' },
    { key: 'tierAName', label: 'Tier 1 \u2014 Name', type: 'text', default: 'Miner' },
    { key: 'tierAAmount', label: 'Tier 1 \u2014 Amount', type: 'text', default: '$25/mo' },
    { key: 'tierAPerks', label: 'Tier 1 \u2014 Perks (one per line)', type: 'textarea', default: 'Newsletter\nClass updates' },
    { key: 'tierBName', label: 'Tier 2 \u2014 Name (highlighted)', type: 'text', default: 'Advocate' },
    { key: 'tierBAmount', label: 'Tier 2 \u2014 Amount', type: 'text', default: '$50/mo' },
    { key: 'tierBPerks', label: 'Tier 2 \u2014 Perks (one per line)', type: 'textarea', default: 'All Miner perks\nEvent invitations' },
    { key: 'tierCName', label: 'Tier 3 \u2014 Name', type: 'text', default: 'Founder' },
    { key: 'tierCAmount', label: 'Tier 3 \u2014 Amount', type: 'text', default: '$100/mo' },
    { key: 'tierCPerks', label: 'Tier 3 \u2014 Perks (one per line)', type: 'textarea', default: 'All Advocate perks\nNamed recognition' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'Compare all levels' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: 'mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', link: '#007A33', cardBg: '#ffffff', cardBorder: 'rgba(0,0,0,0.12)', cardName: '#525252', cardAmount: '#154734', cardPerks: '#525252', hiBg: '#154734', hiName: b.accent, hiAmount: '#ffffff', hiRule: 'rgba(255,255,255,0.25)', hiPerks: 'rgba(255,255,255,0.85)' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', link: '#007A33', cardBg: '#ffffff', cardBorder: 'rgba(0,0,0,0.12)', cardName: '#525252', cardAmount: '#154734', cardPerks: '#525252', hiBg: '#154734', hiName: b.accent, hiAmount: '#ffffff', hiRule: 'rgba(255,255,255,0.25)', hiPerks: 'rgba(255,255,255,0.85)' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', link: '#ffffff', cardBg: 'rgba(255,255,255,0.06)', cardBorder: 'rgba(255,255,255,0.25)', cardName: 'rgba(255,255,255,0.8)', cardAmount: '#ffffff', cardPerks: 'rgba(255,255,255,0.8)', hiBg: b.accent, hiName: '#154734', hiAmount: '#154734', hiRule: 'rgba(21,71,52,0.25)', hiPerks: 'rgba(21,71,52,0.8)' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const tier = (name, amount, perksRaw, hi) => {
      const perks = String(perksRaw || '').split('\n').map(s => s.trim()).filter(Boolean).join('<br/>');
      return `<div style="${hi ? `background-color:${pal.hiBg};` : `background-color:${pal.cardBg};border:1px solid ${pal.cardBorder};`}border-radius:6px;padding:16px 10px;text-align:center;font-family:${FF};">
          <div style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:${hi ? pal.hiName : pal.cardName};">${name || ''}</div>
          <div style="margin-top:6px;font-size:17px;font-weight:900;color:${hi ? pal.hiAmount : pal.cardAmount};">${amount || ''}</div>
          <div style="margin-top:8px;height:1px;background-color:${hi ? pal.hiRule : pal.cardBorder};font-size:0;">&nbsp;</div>
          <div style="margin-top:8px;font-size:11px;color:${hi ? pal.hiPerks : pal.cardPerks};line-height:1.6;">${perks}</div>
        </div>`;
    };
    const cols = fluidColsGrow([
      { html: tier(v.tierAName, v.tierAAmount, v.tierAPerks, false), width: 178, style: 'padding:0 5px 10px 0;' },
      { html: tier(v.tierBName, v.tierBAmount, v.tierBPerks, true), width: 178, style: 'padding:0 5px 10px;' },
      { html: tier(v.tierCName, v.tierCAmount, v.tierCPerks, false), width: 178, style: 'padding:0 0 10px 5px;' }
    ]);
    const inner = `<div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:8px 0 0;font-size:22px;font-weight:900;color:${pal.title};letter-spacing:-0.3px;">${v.title}</h2>
        <div style="margin-top:18px;text-align:center;">${cols}</div>
        <div style="margin-top:8px;"><a href="${v.linkHref}" target="_blank" style="font-size:13px;font-weight:800;color:${pal.link};text-decoration:none;">${v.linkText} &rarr;</a></div>
      </div>`;
    return `<!-- ===== Missouri S&T · Giving Levels ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'dept-comparison', name: 'Progress Comparison', tags: ['Fundraising', 'Comparison', 'Progress'],
  category: 'Giving & Fundraising',
  desc: 'Compare progress across any set of groups — departments, teams, or class years — with stacked mini progress bars. One row per line: Label | Amount | % of bar.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Giving Day Progress' },
    { key: 'title', label: 'Title', type: 'text', default: 'How each college is tracking today' },
    { key: 'items', label: 'Rows', type: 'repeater', addLabel: 'Add row',
      parts: [
        { key: 'label', label: 'Label', type: 'text' },
        { key: 'amount', label: 'Amount raised', type: 'text' },
        { key: 'pct', label: '% of goal (0\u2013100)', type: 'text' }
      ],
      default: 'College of Engineering | $18,420 | 92\nCollege of Arts & Sciences | $16,110 | 80\nCollege of Business | $9,760 | 48' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'See the full comparison' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: 'mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', label: '#154734', track: '#eceae6', bar: '#007A33', link: '#007A33' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', label: '#154734', track: '#e0e3da', bar: '#007A33', link: '#007A33' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', label: '#ffffff', track: 'rgba(255,255,255,0.2)', bar: b.accent, link: b.accent }
    };
    const pal = panels[v.panelBg] || panels.White;
    const rows = lines(v.items).map(p => `<div style="margin-top:14px;font-family:${FF};">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr>
          <td style="font-size:13px;font-weight:700;color:${pal.label};">${p[0] || ''}</td>
          <td align="right" style="font-size:13px;font-weight:700;color:${pal.label};white-space:nowrap;">${p[1] || ''}</td>
        </tr></table>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:separate;margin-top:6px;"><tr>
          <td bgcolor="${pal.track}" style="background-color:${pal.track};border-radius:3px;padding:0;">
            <table role="presentation" width="${p[2] || 0}%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:separate;"><tr><td bgcolor="${pal.bar}" style="background-color:${pal.bar};height:6px;border-radius:3px;font-size:0;">&nbsp;</td></tr></table>
          </td>
        </tr></table>
      </div>`).join('');
    const inner = `<div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:8px 0 0;font-size:22px;font-weight:900;color:${pal.title};letter-spacing:-0.3px;">${v.title}</h2>
        <div style="margin-top:6px;">${rows}</div>
        <div style="margin-top:18px;"><a href="${v.linkHref}" target="_blank" style="font-size:13px;font-weight:800;color:${pal.link};text-decoration:none;">${v.linkText} &rarr;</a></div>
      </div>`;
    return `<!-- ===== Missouri S&T · Department Progress Comparison ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'campus-wishlist', name: 'Gift-in-Kind Wishlist', tags: ['Fundraising', 'Wishlist'],
  category: 'Giving & Fundraising',
  desc: 'A short wishlist of small, specific needs with a fund link beside each \u2014 an easier ask than a general appeal.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Campus Wishlist' },
    { key: 'title', label: 'Title', type: 'text', default: 'Small gifts the library needs this term' },
    { key: 'items', label: 'Wishlist items', type: 'repeater', addLabel: 'Add item',
      parts: [
        { key: 'item', label: 'Item', type: 'text' },
        { key: 'linkText', label: 'Link text', type: 'text' },
        { key: 'linkHref', label: 'Link URL', type: 'text' }
      ],
      default: '2 study room whiteboards | Fund it | mst.edu\nBraille printer paper, 1 yr supply | Fund it | mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', item: '#154734', link: '#007A33', border: '#eceae6' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', item: '#154734', link: '#007A33', border: '#e0e3da' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', item: '#ffffff', link: b.accent, border: 'rgba(255,255,255,0.15)' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const rows = lines(v.items).map(p => `<tr><td style="padding:12px 0;border-top:1px solid ${pal.border};font-family:${FF};">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr>
            <td style="font-size:13px;color:${pal.item};font-weight:700;">${p[0] || ''}</td>
            <td align="right"><a href="${p[2] || '#'}" target="_blank" style="font-size:13px;font-weight:800;color:${pal.link};text-decoration:none;white-space:nowrap;">${p[1] || ''} &rarr;</a></td>
          </tr></table>
        </td></tr>`).join('');
    const inner = `<div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:8px 0 0;font-size:20px;font-weight:800;color:${pal.title};">${v.title}</h2>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin-top:14px;">${rows}</table>
      </div>`;
    return `<!-- ===== Missouri S&T · Gift-in-Kind Wishlist ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'speaker-series', name: 'Speaker Series Lineup', tags: ['Speakers', 'Series', 'Lineup'],
  category: 'People & Directory',
  desc: 'A multi-speaker lineup for a lecture or seminar series, each with a photo placeholder, topic, and date.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Fall Speaker Series' },
    { key: 'title', label: 'Title', type: 'text', default: 'Four talks, one Thursday a month' },
    { key: 'showPhotos', label: 'Show speaker photos', type: 'select', options: ['Yes', 'No'], default: 'Yes' },
    { key: 'items', label: 'Speakers', type: 'repeater', addLabel: 'Add speaker',
      parts: [
        { key: 'name', label: 'Name', type: 'text' },
        { key: 'topic', label: 'Topic', type: 'text' },
        { key: 'date', label: 'Date', type: 'text' },
        { key: 'alt', label: 'Photo alt text', type: 'text' },
        { key: 'photoUrl', label: 'Photo URL (paste a hosted image)', type: 'text' }
      ],
      default: 'Dr. Sana Iqbal | Climate modeling at scale | Sep 11 | Portrait of Dr. Sana Iqbal | \nMarcus Webb | Building resilient infrastructure | Oct 9 | Portrait of Marcus Webb | \nDr. Renee Okafor | AI in materials discovery | Nov 13 | Portrait of Dr. Renee Okafor | ' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', name: '#154734', topic: '#525252', date: '#007A33', border: '#eceae6' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', name: '#154734', topic: '#525252', date: '#007A33', border: '#e0e3da' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', name: '#ffffff', topic: 'rgba(255,255,255,0.7)', date: b.accent, border: 'rgba(255,255,255,0.15)' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const showPhotos = v.showPhotos !== 'No';
    const photoCell = showPhotos
      ? (p) => `<td width="64" valign="middle"><img src="${(p[4] || '').trim() || phImg('Photo', 160, 160, b.accent)}" width="64" height="64" alt="${p[3] || ''}" style="display:block;width:64px;height:64px;border-radius:50%;border:0;object-fit:cover;"></td>`
      : () => '';
    const rows = lines(v.items).map(p => `<tr><td style="padding:14px 0;border-top:1px solid ${pal.border};font-family:${FF};">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr>
            ${photoCell(p)}
            <td valign="middle" style="padding:${showPhotos ? '0 12px' : '0'};">
              <div style="font-size:14px;font-weight:800;color:${pal.name};">${p[0] || ''}</div>
              <div style="font-size:12px;color:${pal.topic};">${p[1] || ''}</div>
            </td>
            <td valign="middle" align="right" style="font-size:12px;font-weight:700;color:${pal.date};white-space:nowrap;">${p[2] || ''}</td>
          </tr></table>
        </td></tr>`).join('');
    const inner = `<div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:8px 0 0;font-size:20px;font-weight:800;color:${pal.title};">${v.title}</h2>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin-top:14px;">${rows}</table>
      </div>`;
    return `<!-- ===== Missouri S&T · Speaker Series Lineup ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'quick-facts', name: 'Quick Facts Spec Sheet', tags: ['Facts', 'Program', 'Spec'],
  category: 'Info & Resources',
  desc: 'A label/value spec grid for one or more programs at a glance \u2014 length, format, cost, dates, and the like. Add multiple programs to stack them vertically.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Program at a Glance' },
    { key: 'programs', label: 'Programs', type: 'repeater', addLabel: 'Add program',
      parts: [
        { key: 'title', label: 'Program title', type: 'text' },
        { key: 'facts', label: 'Facts \u2014 "Label: Value" pairs separated by ;', type: 'text' }
      ],
      default: 'M.S. in Data Science | Length: 18 months; Format: Online + on-campus; Credit hours: 30; Next start: January 2027' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'View the full program page' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: 'mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', label: '#666666', value: '#154734', border: '#eceae6', link: '#007A33' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', label: '#666666', value: '#154734', border: '#e0e3da', link: '#007A33' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', label: 'rgba(255,255,255,0.6)', value: '#ffffff', border: 'rgba(255,255,255,0.15)', link: b.accent }
    };
    const pal = panels[v.panelBg] || panels.White;
    const progs = lines(v.programs);
    const parseFacts = (s) => String(s || '').split(';').map(x => x.trim()).filter(Boolean).map(pair => {
      const idx = pair.indexOf(':');
      return idx === -1 ? [pair, ''] : [pair.slice(0, idx).trim(), pair.slice(idx + 1).trim()];
    });
    const one = (title, factsRaw) => {
      const facts = parseFacts(factsRaw);
      const cells = facts.map((p, i) => {
        const left = i % 2 === 0;
        return `<td width="50%" valign="top" style="padding:12px ${left ? '16px' : '0'} 12px ${left ? '0' : '16px'};border-top:1px solid ${pal.border};${left ? '' : `border-left:1px solid ${pal.border};`}font-family:${FF};">
            <div style="font-size:11px;color:${pal.label};text-transform:uppercase;letter-spacing:0.5px;">${p[0] || ''}</div>
            <div style="margin-top:3px;font-size:14px;font-weight:700;color:${pal.value};">${p[1] || ''}</div>
          </td>`;
      });
      const rows = [];
      for (let i = 0; i < cells.length; i += 2) rows.push(`<tr>${cells[i] || '<td width="50%"></td>'}${cells[i + 1] || '<td width="50%"></td>'}</tr>`);
      return `<div style="margin-top:20px;">
          <h3 style="margin:0;font-size:16px;font-weight:800;color:${pal.title};font-family:${FF};">${title || ''}</h3>
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin-top:10px;">${rows.join('')}</table>
        </div>`;
    };
    const allPrograms = progs.map(p => one(p[0], p[1])).join('');
    const inner = `<div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        ${allPrograms}
        <div style="margin-top:18px;"><a href="${v.linkHref}" target="_blank" style="font-size:13px;font-weight:800;color:${pal.link};text-decoration:none;">${v.linkText} &rarr;</a></div>
      </div>`;
    return `<!-- ===== Missouri S&T · Quick Facts Spec Sheet ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'sidebar-layout', name: 'Article + Resource Sidebar', tags: ['Article', 'Sidebar', 'Resources'],
  category: 'Info & Resources',
  desc: 'A short article next to a slim resource sidebar of related links.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Advising' },
    { key: 'title', label: 'Title', type: 'textarea', default: 'Fall registration opens Monday \u2014 here\u2019s how to prepare' },
    { key: 'body', label: 'Body', type: 'textarea', default: 'Meet with your advisor before your registration window opens to confirm you\u2019re on track and avoid hold delays.' },
    { key: 'links', label: 'Sidebar links', type: 'repeater', addLabel: 'Add link',
      parts: [
        { key: 'label', label: 'Link label', type: 'text' },
        { key: 'href', label: 'URL', type: 'text' }
      ],
      default: 'Book an advisor | mst.edu\nCheck your holds | mst.edu\nDegree audit tool | mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', body: '#4a554c', sidebarBg: '#f4f5f1', sidebarLabel: '#666666', link: '#007A33' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', body: '#4a554c', sidebarBg: '#ffffff', sidebarLabel: '#666666', link: '#007A33' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', body: 'rgba(255,255,255,0.85)', sidebarBg: 'rgba(255,255,255,0.08)', sidebarLabel: 'rgba(255,255,255,0.6)', link: b.accent }
    };
    const pal = panels[v.panelBg] || panels.White;
    const links = lines(v.links).map(p => `<a href="${p[1] || '#'}" target="_blank" style="display:block;font-size:12px;font-weight:700;color:${pal.link};text-decoration:none;margin-top:8px;font-family:${FF};">${p[0] || ''}</a>`).join('');
    const article = `<div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
      <div style="margin-top:6px;font-size:16px;font-weight:800;color:${pal.title};line-height:1.3;">${v.title}</div>
      <p style="margin:10px 0 0;font-size:13px;line-height:1.6;color:${pal.body};">${v.body}</p>`;
    const sidebar = `<div style="background-color:${pal.sidebarBg};border-radius:6px;padding:18px 20px;margin-top:24px;font-family:${FF};">
        <div style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:${pal.sidebarLabel};">Resources</div>
        <div style="margin-top:2px;">${links}</div>
      </div>`;
    const inner = twoCol(article, 360, `font-family:${FF};padding:0 24px 0 0;`, sidebar, 170, `font-family:${FF};padding:0;`);
    return `<!-- ===== Missouri S&T · Article + Resource Sidebar ===== -->
${shell(`<div style="padding:30px 0;">${inner}</div>`, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'how-it-works', name: 'How It Works Explainer', tags: ['Process', 'Steps', 'Explainer'],
  category: 'Info & Resources',
  desc: 'A numbered process explainer \u2014 for policies, applications, or anything with clear sequential steps.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'How Transfer Credit Works' },
    { key: 'items', label: 'Steps', type: 'repeater', addLabel: 'Add step',
      parts: [
        { key: 'bold', label: 'Bold lead-in', type: 'text' },
        { key: 'rest', label: 'Rest of sentence', type: 'textarea' }
      ],
      default: 'Submit transcripts | from every school you\u2019ve attended.\nWe evaluate | each course against our catalog, usually within 2 weeks.\nCredit posts | to your degree audit automatically.' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['Tint', 'White', 'Forest'], default: 'Tint' }
  ],
  render: (v, b) => {
    const panels = {
      Tint: { bg: '#f4f5f1', kicker: '#007A33', numBg: '#154734', numText: b.accent, bold: '#154734', rest: '#4a554c' },
      White: { bg: '#ffffff', kicker: '#007A33', numBg: '#154734', numText: b.accent, bold: '#154734', rest: '#4a554c' },
      Forest: { bg: '#154734', kicker: b.accent, numBg: b.accent, numText: '#154734', bold: '#ffffff', rest: 'rgba(255,255,255,0.85)' }
    };
    const pal = panels[v.panelBg] || panels.Tint;
    const rows = lines(v.items).map((p, i) => `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin-top:${i === 0 ? '16' : '14'}px;"><tr>
        <td width="32" valign="top"><div style="font-size:15px;font-weight:900;color:${pal.numText};background-color:${pal.numBg};width:32px;height:32px;border-radius:50%;text-align:center;line-height:32px;font-family:${FF};">${i + 1}</div></td>
        <td valign="top" style="padding-left:18px;padding-top:4px;font-size:13px;line-height:1.6;color:${pal.rest};font-family:${FF};"><strong style="color:${pal.bold};">${p[0] || ''}</strong> ${p[1] || ''}</td>
      </tr></table>`).join('');
    const inner = `<div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        ${rows}
      </div>`;
    return `<!-- ===== Missouri S&T · How It Works Explainer ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'resource-library', name: 'Resource Library Grid', tags: ['Downloads', 'Resources'],
  category: 'Info & Resources',
  desc: 'A short stack of labeled downloads \u2014 PDFs, forms, maps \u2014 with a colored type badge beside each.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Downloads' },
    { key: 'title', label: 'Title', type: 'text', default: 'Everything you need for orientation' },
    { key: 'items', label: 'Resources', type: 'repeater', addLabel: 'Add resource',
      parts: [
        { key: 'badge', label: 'Badge (short, e.g. PDF)', type: 'text' },
        { key: 'label', label: 'Label', type: 'text' },
        { key: 'href', label: 'URL', type: 'text' }
      ],
      default: 'PDF | Orientation schedule | mst.edu\nFORM | Housing preference form | mst.edu\nMAP | Campus map, printable | mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', label: '#154734', border: '#eceae6' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', label: '#154734', border: '#e0e3da' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', label: '#ffffff', border: 'rgba(255,255,255,0.2)' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const isForest = v.panelBg === 'Forest';
    const badgeColors = isForest
      ? { PDF: '#ffffff', FORM: b.accent, MAP: '#72BF44' }
      : { PDF: '#154734', FORM: '#007A33', MAP: b.accent };
    const rows = lines(v.items).map(p => {
      const badge = (p[0] || '').toUpperCase();
      const bg = badgeColors[badge] || (isForest ? '#ffffff' : '#154734');
      const fg = (bg === b.accent || bg === '#72BF44' || bg === '#ffffff') ? '#154734' : '#ffffff';
      return `<a href="${p[2] || '#'}" target="_blank" style="display:block;text-decoration:none;border:1px solid ${pal.border};border-radius:6px;padding:12px 14px;margin-top:8px;font-family:${FF};">
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr>
            <td style="font-size:10px;font-weight:800;color:${fg};background-color:${bg};border-radius:3px;padding:4px 7px;white-space:nowrap;">${badge}</td>
            <td style="padding-left:12px;font-size:13px;font-weight:700;color:${pal.label};">${p[1] || ''}</td>
          </tr></table>
        </a>`;
    }).join('');
    const inner = `<div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:8px 0 0;font-size:20px;font-weight:800;color:${pal.title};">${v.title}</h2>
        <div style="margin-top:8px;">${rows}</div>
      </div>`;
    return `<!-- ===== Missouri S&T · Resource Library Grid ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'policy-update', name: 'Before/After Comparison', tags: ['Policy', 'Update', 'Comparison'],
  category: 'Info & Resources',
  desc: 'A side-by-side before/after comparison for any update or change — a policy, a deadline, a redesign.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Policy Update' },
    { key: 'title', label: 'Title', type: 'textarea', default: 'Late add/drop deadline is changing' },
    { key: 'before', label: 'Before', type: 'textarea', default: 'Drop deadline: end of week 12, no exceptions' },
    { key: 'after', label: 'After', type: 'textarea', default: 'Drop deadline: end of week 14, advisor sign-off required' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'Read the full policy' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: 'mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', beforeLabel: '#666666', beforeText: '#525252', afterLabel: '#007A33', afterText: '#154734', link: '#007A33' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', beforeLabel: '#666666', beforeText: '#525252', afterLabel: '#007A33', afterText: '#154734', link: '#007A33' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', beforeLabel: 'rgba(255,255,255,0.55)', beforeText: 'rgba(255,255,255,0.75)', afterLabel: b.accent, afterText: '#ffffff', link: b.accent }
    };
    const pal = panels[v.panelBg] || panels.White;
    const before = `<div style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:${pal.beforeLabel};">Before</div><div style="margin-top:6px;font-size:13px;color:${pal.beforeText};line-height:1.6;">${v.before}</div>`;
    const after = `<div class="pu-after"><div style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:${pal.afterLabel};">After</div><div style="margin-top:6px;font-size:13px;color:${pal.afterText};font-weight:600;line-height:1.6;">${v.after}</div></div>`;
    const cols = twoCol(before, 260, `font-family:${FF};padding:0 16px 0 0;`, after, 260, `font-family:${FF};padding:0;`);
    const inner = `<style>@media screen and (max-width:480px){.pu-after{margin-top:20px!important;display:block!important;}}</style>
      <div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:8px 0 0;font-size:20px;font-weight:800;color:${pal.title};">${v.title}</h2>
        <div style="margin-top:18px;">${cols}</div>
        <div style="margin-top:16px;"><a href="${v.linkHref}" target="_blank" style="font-size:13px;font-weight:800;color:${pal.link};text-decoration:none;">${v.linkText} &rarr;</a></div>
      </div>`;
    return `<!-- ===== Missouri S&T · Policy Update Before/After ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'study-abroad-grid', name: 'Destination Grid', tags: ['Study Abroad', 'Grid', 'Destinations'],
  category: 'Info & Resources',
  desc: 'A 3-across grid of linked destination photo placeholders \u2014 study abroad, campus tours, or program locations.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Study Abroad' },
    { key: 'title', label: 'Title', type: 'text', default: 'Applications open for spring programs' },
    { key: 'items', label: 'Destinations', type: 'repeater', addLabel: 'Add destination',
      parts: [
        { key: 'name', label: 'Name', type: 'text' },
        { key: 'href', label: 'URL', type: 'text' },
        { key: 'alt', label: 'Photo alt text', type: 'text' },
        { key: 'photoUrl', label: 'Photo URL (paste a hosted image)', type: 'text' }
      ],
      default: 'Seoul, Korea | mst.edu | Skyline of Seoul, Korea | \nValpara\u00edso, Chile | mst.edu | Hillside houses in Valpara\u00edso, Chile | \nFlorence, Italy | mst.edu | The Duomo above the rooftops of Florence, Italy | ' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'See all destinations' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: 'mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', name: '#154734', link: '#007A33' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', name: '#154734', link: '#007A33' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', name: '#ffffff', link: b.accent }
    };
    const pal = panels[v.panelBg] || panels.White;
    const items = lines(v.items);
    const cell = (p) => `<a href="${p[1] || '#'}" target="_blank" style="text-decoration:none;display:block;font-family:${FF};">
        <img src="${(p[3] || '').trim() || phImg('Photo', 200, 160, b.accent)}" width="180" style="display:block;width:100%;height:120px;object-fit:cover;border:0;" alt="${p[2] || ''}">
        <div style="margin-top:8px;font-size:12px;font-weight:800;color:${pal.name};">${p[0] || ''}</div>
      </a>`;
    const cols = gridRows(items, 3, cell, 16, 16, 'sag-cell');
    const inner = `<div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:8px 0 0;font-size:20px;font-weight:800;color:${pal.title};">${v.title}</h2>
        <div style="margin-top:14px;">${cols}</div>
        <div style="margin-top:16px;"><a href="${v.linkHref}" target="_blank" style="font-size:13px;font-weight:800;color:${pal.link};text-decoration:none;">${v.linkText} &rarr;</a></div>
      </div>`;
    return `<!-- ===== Missouri S&T · Destination Grid ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'career-fair-grid', name: 'Employer Grid', tags: ['Career Fair', 'Employers', 'Grid'],
  category: 'Info & Resources',
  desc: 'A compact grid of employer name badges for a career fair or partner announcement.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Fall Career Fair' },
    { key: 'title', label: 'Title', type: 'text', default: '120+ employers, September 16' },
    { key: 'items', label: 'Employers', type: 'repeater', addLabel: 'Add employer',
      parts: [
        { key: 'name', label: 'Employer name', type: 'text' }
      ],
      default: 'Boeing\nAmeren\nCaterpillar\nEmerson' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'See the full employer list' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: 'mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', name: '#4a554c', border: '#eceae6', link: '#007A33' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', name: '#4a554c', border: '#dfe2d9', link: '#007A33' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', name: 'rgba(255,255,255,0.85)', border: 'rgba(255,255,255,0.25)', link: b.accent }
    };
    const pal = panels[v.panelBg] || panels.White;
    const names = lines(v.items).map(p => (p[0] || '').trim()).filter(Boolean);
    const cell = (n) => `<div style="border:1px solid ${pal.border};border-radius:6px;padding:14px 8px;text-align:center;font-size:11px;font-weight:800;color:${pal.name};font-family:${FF};box-sizing:border-box;">${n}</div>`;
    const cols = gridRows(names, 3, cell, 10, 10, 'cfg-cell');
    const inner = `<div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:8px 0 0;font-size:20px;font-weight:800;color:${pal.title};">${v.title}</h2>
        <div style="margin-top:14px;">${cols}</div>
        <div style="margin-top:16px;"><a href="${v.linkHref}" target="_blank" style="font-size:13px;font-weight:800;color:${pal.link};text-decoration:none;">${v.linkText} &rarr;</a></div>
      </div>`;
    return `<!-- ===== Missouri S&T · Employer Grid ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'admissions-steps', name: 'Step Checklist', tags: ['Checklist', 'Progress', 'Steps'],
  category: 'Info & Resources',
  desc: 'A numbered checklist with completed steps checked off \u2014 for onboarding, enrollment, or any multi-step process.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Admitted Student' },
    { key: 'title', label: 'Title', type: 'text', default: 'You\u2019re almost enrolled \u2014 2 steps left' },
    { key: 'items', label: 'Steps', type: 'repeater', addLabel: 'Add step',
      parts: [
        { key: 'label', label: 'Step label', type: 'text' },
        { key: 'meta', label: 'Detail / date', type: 'text' },
        { key: 'status', label: 'Status (done / active / upcoming)', type: 'text' },
        { key: 'linkText', label: 'Link text (active step only)', type: 'text' },
        { key: 'linkHref', label: 'Link URL (active step only)', type: 'text' }
      ],
      default: 'Application submitted | Completed March 2 | done | | \nAdmitted | Completed March 28 | done | | \nConfirm your enrollment deposit | Due by May 1 | active | Do it now | mst.edu\nRegister for orientation | Opens after deposit | upcoming | | ' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', label: '#154734', upcoming: '#666666', meta: '#666666', ringBorder: 'rgba(0,0,0,0.15)', done: '#007A33', active: '#007A33', link: '#007A33' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', label: '#154734', upcoming: '#666666', meta: '#666666', ringBorder: 'rgba(0,0,0,0.15)', done: '#007A33', active: '#007A33', link: '#007A33' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', label: '#ffffff', upcoming: 'rgba(255,255,255,0.6)', meta: 'rgba(255,255,255,0.6)', ringBorder: 'rgba(255,255,255,0.3)', done: '#72BF44', active: b.accent, link: b.accent }
    };
    const pal = panels[v.panelBg] || panels.White;
    const doneText = v.panelBg === 'Forest' ? '#154734' : '#ffffff';
    const rows = lines(v.items).map((p, i) => {
      const status = (p[2] || 'upcoming').trim();
      const marker = status === 'done'
        ? `<div style="width:24px;height:24px;border-radius:50%;background-color:${pal.done};color:${doneText};font-size:13px;font-weight:800;text-align:center;line-height:24px;">&#10003;</div>`
        : status === 'active'
          ? `<div style="width:24px;height:24px;border-radius:50%;border:2px solid ${pal.active};color:${pal.active};font-size:12px;font-weight:800;text-align:center;line-height:20px;">${i + 1}</div>`
          : `<div style="width:24px;height:24px;border-radius:50%;border:2px solid ${pal.ringBorder};color:${pal.upcoming};font-size:12px;font-weight:800;text-align:center;line-height:20px;">${i + 1}</div>`;
      const labelColor = status === 'upcoming' ? pal.upcoming : pal.label;
      const link = (status === 'active' && p[3]) ? `<a href="${p[4] || '#'}" target="_blank" style="font-size:13px;font-weight:800;color:${pal.link};text-decoration:none;white-space:nowrap;">${p[3]} &rarr;</a>` : '';
      return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin-top:16px;"><tr>
          <td width="24" valign="top">${marker}</td>
          <td valign="top" style="padding:0 12px;font-family:${FF};">
            <div style="font-size:14px;font-weight:800;color:${labelColor};">${p[0] || ''}</div>
            <div style="font-size:12px;color:${pal.meta};">${p[1] || ''}</div>
          </td>
          <td valign="top" align="right">${link}</td>
        </tr></table>`;
    }).join('');
    const inner = `<div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:8px 0 0;font-size:20px;font-weight:800;color:${pal.title};">${v.title}</h2>
        ${rows}
      </div>`;
    return `<!-- ===== Missouri S&T · Admissions Next Steps ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'then-now', name: 'Then & Now Split', tags: ['History', 'Comparison', 'Photo'],
  category: 'Info & Resources',
  desc: 'A two-photo split comparing a building, program, or tradition then and now.',
  fields: [
    { key: 'photoUrlA', label: 'Left photo URL (paste a hosted image)', type: 'text', default: '' },
    { key: 'labelA', label: 'Left label', type: 'text', default: '1974' },
    { key: 'altA', label: 'Left photo alt text', type: 'text', default: 'Toomey Hall exterior, 1974' },
    { key: 'photoUrlB', label: 'Right photo URL (paste a hosted image)', type: 'text', default: '' },
    { key: 'labelB', label: 'Right label', type: 'text', default: '2026' },
    { key: 'altB', label: 'Right photo alt text', type: 'text', default: 'Toomey Hall exterior, 2026' },
    { key: 'kicker', label: 'Kicker', type: 'text', default: '50 Years of Toomey Hall' },
    { key: 'body', label: 'Body', type: 'textarea', default: 'From a single lecture hall to the heart of engineering on campus.' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', label: '#154734', body: '#4a554c' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', label: '#154734', body: '#4a554c' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', label: '#ffffff', body: 'rgba(255,255,255,0.85)' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const urlA = (v.photoUrlA || '').trim() || phImg('Photo', 290, 200, b.accent);
    const urlB = (v.photoUrlB || '').trim() || phImg('Photo', 290, 200, b.accent);
    const strip = `<div style="position:relative;font-size:0;line-height:0;">
        <!--[if mso]><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td width="50%"><![endif]-->
        <img src="${urlA}" alt="${v.altA || v.labelA}" width="290" style="display:inline-block;width:50%;height:200px;object-fit:cover;border:0;vertical-align:top;">
        <!--[if mso]></td><td width="50%"><![endif]-->
        <img src="${urlB}" alt="${v.altB || v.labelB}" width="290" style="display:inline-block;width:50%;height:200px;object-fit:cover;border:0;vertical-align:top;">
        <!--[if mso]></td></tr></table><![endif]-->
        <div style="position:absolute;top:0;bottom:0;left:50%;width:2px;background-color:${b.accent};transform:translateX(-1px);"></div>
        <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);width:34px;height:34px;border-radius:50%;background-color:#ffffff;box-shadow:0 1px 4px rgba(0,0,0,0.25);text-align:center;line-height:34px;font-size:14px;font-weight:800;color:#154734;font-family:${FF};">&harr;</div>
      </div>`;
    const labels = `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin-top:10px;"><tr>
        <td align="left" style="font-size:14px;font-weight:800;color:${pal.label};font-family:${FF};">${v.labelA}</td>
        <td align="right" style="font-size:14px;font-weight:800;color:${pal.label};font-family:${FF};">${v.labelB}</td>
      </tr></table>`;
    const inner = `<div style="font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};">Then &amp; Now</div>
        <h2 style="margin:8px 0 0;font-size:23px;font-weight:900;color:${pal.title};letter-spacing:-0.4px;line-height:1.25;">${v.kicker}</h2>
      </div>
      <div style="margin-top:18px;">${strip}${labels}</div>
      <div style="margin-top:18px;font-family:${FF};">
        <div style="font-size:14px;line-height:1.6;color:${pal.body};">${v.body}</div>
      </div>`;
    return `<!-- ===== Missouri S&T · Then & Now Split ===== -->
${shell(inner, { maxWidth: 580, padY: 30, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'magazine-column', name: 'Magazine Column', tags: ['Editorial', 'Feature', 'Story'],
  category: 'Newsletter & Content',
  desc: 'A long-form editorial block with a typographic drop cap \u2014 for a feature story or field-notes piece.',
  fields: [
    { key: 'title', label: 'Title', type: 'text', default: 'Notes from the field' },
    { key: 'firstLetter', label: 'Drop cap letter', type: 'text', default: 'T' },
    { key: 'body', label: 'Body (starts after the drop cap)', type: 'textarea', default: 'his spring, six students spent their break mapping sinkholes across the Ozarks for a state geological survey \u2014 fieldwork that started as a class project and turned into published data. It\u2019s the kind of hands-on research this program was built for.' },
    { key: 'linkText', label: 'Link text', type: 'text', default: 'Read the full story' },
    { key: 'linkHref', label: 'Link URL', type: 'text', default: 'mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', title: '#154734', body: '#4a554c', cap: '#007A33', link: '#007A33' },
      Tint: { bg: '#f4f5f1', title: '#154734', body: '#4a554c', cap: '#007A33', link: '#007A33' },
      Forest: { bg: '#154734', title: '#ffffff', body: 'rgba(255,255,255,0.85)', cap: b.accent, link: b.accent }
    };
    const pal = panels[v.panelBg] || panels.White;
    const inner = `<div style="padding:34px 0;font-family:${FF};">
        <h2 style="margin:0 0 14px;font-size:22px;font-weight:900;color:${pal.title};letter-spacing:-0.3px;">${v.title}</h2>
        <div style="font-size:13px;line-height:1.75;color:${pal.body};">
          <span style="float:left;font-size:52px;line-height:0.8;font-weight:900;color:${pal.cap};padding:6px 8px 0 0;font-family:${FF};">${v.firstLetter}</span>
          ${v.body}
        </div>
        <div style="clear:both;"></div>
        <div style="margin-top:14px;"><a href="${v.linkHref}" target="_blank" style="font-size:13px;font-weight:800;color:${pal.link};text-decoration:none;">${v.linkText} &rarr;</a></div>
      </div>`;
    return `<!-- ===== Missouri S&T · Magazine Column ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'track-picker', name: 'Concurrent Track Picker', tags: ['Sessions', 'Tracks', 'Agenda'],
  category: 'Events & Dates',
  desc: 'Side-by-side session tracks for a single time block, each with its own color accent \u2014 use when attendees must choose one of several concurrent sessions.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Session Block' },
    { key: 'title', label: 'Title', type: 'text', default: 'Pick your track' },
    { key: 'items', label: 'Tracks', type: 'repeater', addLabel: 'Add track',
      parts: [
        { key: 'label', label: 'Track label', type: 'text' },
        { key: 'location', label: 'Location', type: 'text' },
        { key: 'session', label: 'Session title', type: 'text' }
      ],
      default: 'Track A | Toomey 199 | Robotics in the field\nTrack B | Butler-Carlton 125 | Sustainable materials\nTrack C | EECH 104 | AI ethics panel' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', cardBg: '#f4f5f1', session: '#154734', swatches: [
        { border: '#007A33', text: '#007A33' },
        { border: '#8a9420', text: '#8a9420' },
        { border: '#154734', text: '#154734' }
      ] },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', cardBg: '#ffffff', session: '#154734', swatches: [
        { border: '#007A33', text: '#007A33' },
        { border: '#8a9420', text: '#8a9420' },
        { border: '#154734', text: '#154734' }
      ] },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', cardBg: 'rgba(255,255,255,0.1)', session: '#ffffff', swatches: [
        { border: b.accent, text: b.accent },
        { border: '#72BF44', text: '#72BF44' },
        { border: '#ffffff', text: 'rgba(255,255,255,0.85)' }
      ] }
    };
    const pal = panels[v.panelBg] || panels.White;
    const swatches = pal.swatches;
    const rows = lines(v.items);
    // fluidColsGrow (not the plain fluidCols by-the-numbers/etc. use) so each card stretches
    // to the full row width once it wraps alone on mobile, instead of staying stuck at its
    // narrow desktop share with a stray gap beside it.
    const cols = fluidColsGrow(rows.map((p, i) => {
      const sw = swatches[i % swatches.length];
      const meta = [p[0] || '', p[1] || ''].filter(Boolean).join(' &middot; ');
      const html = `<div style="border-top:4px solid ${sw.border};background-color:${pal.cardBg};border-radius:0 0 8px 8px;padding:16px 14px;font-family:${FF};">
          <div style="font-size:10px;font-weight:800;letter-spacing:1px;text-transform:uppercase;color:${sw.text};">${meta}</div>
          <div style="margin-top:8px;font-size:14px;font-weight:800;color:${pal.session};line-height:1.3;">${p[2] || ''}</div>
        </div>`;
      return { html, width: 190, style: 'padding:0 6px 10px;' };
    }));
    const inner = `<div style="padding:32px 0 6px;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <h2 style="margin:9px 0 0;font-size:21px;font-weight:900;color:${pal.title};letter-spacing:-0.4px;">${v.title}</h2>
      </div>
      <div style="padding:14px 0 34px;">${cols}</div>`;
    return `<!-- ===== Missouri S&T · Concurrent Track Picker ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'mini-calendar', name: 'Mini Calendar Highlight', tags: ['Calendar', 'Date', 'Feature'],
  category: 'Events & Dates',
  desc: 'A small month calendar with one date circled \u2014 use when the date itself, in context of the month, tells the story (a break, a deadline, a recurring day).',
  fields: [
    { key: 'monthLabel', label: 'Month / year label', type: 'text', default: 'March 2026' },
    { key: 'startWeekday', label: 'Weekday the 1st falls on', type: 'select', options: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'], default: 'Sun' },
    { key: 'daysInMonth', label: 'Days in month', type: 'select', options: ['28', '29', '30', '31'], default: '31' },
    { key: 'highlightDay', label: 'Day to highlight', type: 'text', default: '14' },
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'Mark your calendar' },
    { key: 'title', label: 'Title', type: 'text', default: 'Spring Break begins' },
    { key: 'body', label: 'Body', type: 'textarea', default: 'Saturday, March 14 \u2014 classes resume Monday, March 23.' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['White', 'Tint', 'Forest'], default: 'White' }
  ],
  render: (v, b) => {
    const panels = {
      White: { bg: '#ffffff', kicker: '#007A33', title: '#154734', body: '#525252', monthLabel: '#154734', dayNum: '#525252', weekday: '#9aa398', border: '#eceae6' },
      Tint: { bg: '#f4f5f1', kicker: '#007A33', title: '#154734', body: '#525252', monthLabel: '#154734', dayNum: '#525252', weekday: '#9aa398', border: '#dfe2d9' },
      Forest: { bg: '#154734', kicker: b.accent, title: '#ffffff', body: 'rgba(255,255,255,0.8)', monthLabel: '#ffffff', dayNum: 'rgba(255,255,255,0.8)', weekday: 'rgba(255,255,255,0.5)', border: 'rgba(255,255,255,0.2)' }
    };
    const pal = panels[v.panelBg] || panels.White;
    const weekdayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const startIdx = Math.max(0, weekdayNames.indexOf(v.startWeekday));
    const daysInMonth = parseInt(v.daysInMonth, 10) || 30;
    const highlight = parseInt(v.highlightDay, 10) || 0;
    const cells = [];
    for (let i = 0; i < startIdx; i++) cells.push('<td>&nbsp;</td>');
    for (let d = 1; d <= daysInMonth; d++) {
      const isHi = d === highlight;
      cells.push(`<td align="center" style="padding:5px 0;font-family:${FF};font-size:12px;color:${isHi ? '#ffffff' : pal.dayNum};">${isHi ? `<span style="display:inline-block;width:22px;height:22px;line-height:22px;border-radius:50%;background-color:#007A33;color:#ffffff;font-weight:800;">${d}</span>` : d}</td>`);
    }
    while (cells.length % 7 !== 0) cells.push('<td>&nbsp;</td>');
    let bodyRows = '';
    for (let i = 0; i < cells.length; i += 7) bodyRows += `<tr>${cells.slice(i, i + 7).join('')}</tr>`;
    const headerRow = `<tr>${weekdayNames.map(w => `<td align="center" style="font-family:${FF};font-size:9px;font-weight:700;color:${pal.weekday};padding-bottom:6px;">${w[0]}</td>`).join('')}</tr>`;
    const calendar = `<div style="font-size:12px;font-weight:800;color:${pal.monthLabel};text-align:center;margin-bottom:8px;font-family:${FF};text-transform:uppercase;">${v.monthLabel || ''}</div>
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;max-width:220px;width:100%;margin:0 auto;">${headerRow}${bodyRows}</table>`;
    const details = `<div style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:${pal.kicker};font-family:${FF};">${v.kicker}</div>
      <h2 style="margin:8px 0 0;font-size:19px;font-weight:900;color:${pal.title};letter-spacing:-0.3px;font-family:${FF};">${v.title}</h2>
      <div style="margin-top:6px;font-size:13px;color:${pal.body};line-height:1.6;font-family:${FF};">${v.body}</div>`;
    // Hand-built two-column markup (same inline-block + mso-table technique fluidCols uses)
    // instead of the fluidCols/twoCol helper, so each column carries its own class and a
    // media query can let it go max-width:100% once it wraps alone on mobile \u2014 the helper's
    // per-column max-width is a fixed inline style with no hook for that override.
    const calCol = `<!--[if mso]><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td width="38%" valign="top"><![endif]-->
      <div class="mc-cal-col" style="display:inline-block;width:100%;max-width:217px;vertical-align:top;box-sizing:border-box;font-family:${FF};padding:0 24px 0 0;">${calendar}</div>
      <!--[if mso]></td><td width="62%" valign="top"><![endif]-->`;
    const detailsCol = `<div class="mc-details-col" style="display:inline-block;width:100%;max-width:297px;vertical-align:top;box-sizing:border-box;font-family:${FF};border-left:1px solid ${pal.border};padding-left:24px;">${details}</div>
      <!--[if mso]></td></tr></table><![endif]-->`;
    // .mc-details-col's border-left is a side-by-side desktop divider that would otherwise
    // persist as a stray vertical rule + left indent once the column wraps and stacks on
    // mobile; the media query swaps it for a plain top divider and lets both columns go
    // full width instead of staying stuck at their narrow desktop share.
    const inner = `<style>@media screen and (max-width:480px){.mc-cal-col,.mc-details-col{max-width:100%!important;}.mc-details-col{border-left:0!important;padding-left:0!important;border-top:1px solid ${pal.border};padding-top:16px;margin-top:18px;}}</style>
      <div style="padding:32px 0;">${calCol}${detailsCol}</div>`;
    return `<!-- ===== Missouri S&T · Mini Calendar Highlight ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'bold-statement', name: 'Bold Type Statement', tags: ['Stat', 'Statement', 'Impact'],
  category: 'Newsletter & Content',
  desc: 'One striking statistic set in oversized type on a solid gold panel \u2014 for a single number that deserves the whole block.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'By the numbers' },
    { key: 'statement', label: 'Statement (short \u2014 renders large)', type: 'textarea', default: '94% of grads land a job in 6 months.' },
    { key: 'sourceLine', label: 'Source line', type: 'text', default: 'Career Services placement report, 2026' },
    { key: 'palette', label: 'Color combo', type: 'select', options: ['Lime', 'Forest', 'White', 'Green'], default: 'Lime' }
  ],
  render: (v, b) => {
    const palettes = {
      Lime: { bg: b.accent, text: '#154734', source: '#154734' },
      Forest: { bg: '#154734', text: '#ffffff', source: 'rgba(255,255,255,0.75)' },
      White: { bg: '#ffffff', text: '#154734', source: '#525252' },
      Green: { bg: '#007A33', text: '#ffffff', source: 'rgba(255,255,255,0.75)' }
    };
    const pal = palettes[v.palette] || palettes.Lime;
    const inner = `<div style="padding:44px 40px;font-family:${FF};">
        <div style="font-size:13px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;color:${pal.text};">${v.kicker}</div>
        <div style="margin-top:14px;font-size:40px;font-weight:900;color:${pal.text};letter-spacing:-1.2px;line-height:1.12;">${v.statement}</div>
        <div style="margin-top:18px;font-size:13px;font-weight:700;color:${pal.source};">${v.sourceLine}</div>
      </div>`;
    return `<!-- ===== Missouri S&T · Bold Type Statement ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'media-mentions', name: 'Media Mentions', tags: ['Press', 'News', 'Coverage'],
  category: 'Newsletter & Content',
  desc: 'A short list of press clippings \u2014 outlet name and headline, each linking out to the full article.',
  fields: [
    { key: 'kicker', label: 'Kicker', type: 'text', default: 'In the News' },
    { key: 'items', label: 'Mentions', type: 'repeater', addLabel: 'Add mention',
      parts: [
        { key: 'outlet', label: 'Outlet', type: 'text' },
        { key: 'headline', label: 'Headline', type: 'textarea' },
        { key: 'href', label: 'Article URL', type: 'text' }
      ],
      default: 'St. Louis Post-Dispatch | S&amp;T researchers crack decades-old materials puzzle | mst.edu\nForbes | This public university keeps landing grads at top firms | mst.edu' },
    { key: 'panelBg', label: 'Panel background', type: 'select', options: ['Tint', 'White', 'Forest'], default: 'Tint' }
  ],
  render: (v, b) => {
    const panels = {
      Tint: { bg: '#f4f5f1', kicker: '#007A33', cardBg: '#ffffff', outlet: '#8a9484', headline: '#154734' },
      White: { bg: '#ffffff', kicker: '#007A33', cardBg: '#f4f5f1', outlet: '#8a9484', headline: '#154734' },
      Forest: { bg: '#154734', kicker: b.accent, cardBg: 'rgba(255,255,255,0.08)', outlet: 'rgba(255,255,255,0.6)', headline: '#ffffff' }
    };
    const pal = panels[v.panelBg] || panels.Tint;
    const rows = lines(v.items).map(p => `<a href="${p[2] || '#'}" target="_blank" style="display:block;text-decoration:none;background-color:${pal.cardBg};border-radius:8px;padding:14px 16px;margin-top:10px;font-family:${FF};">
        <div style="font-size:10.5px;font-weight:800;letter-spacing:0.6px;text-transform:uppercase;color:${pal.outlet};">${p[0] || ''}</div>
        <div style="margin-top:5px;font-size:14px;font-weight:800;color:${pal.headline};line-height:1.4;">${p[1] || ''}</div>
      </a>`).join('');
    const inner = `<div style="padding:30px 0;font-family:${FF};">
        <div style="font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:${pal.kicker};">${v.kicker}</div>
        <div style="margin-top:2px;">${rows}</div>
      </div>`;
    return `<!-- ===== Missouri S&T · Media Mentions ===== -->
${shell(inner, { maxWidth: 580, bg: pal.bg })}`;
  }
});

MODULES.push({
  id: 'checkerboard-stats', name: 'Checkerboard Stat Wall', tags: ['Stats', 'Impact', 'Grid'],
  category: 'Newsletter & Content',
  fullBleed: true,
  desc: 'Four stats as solid-color quadrants, no heading needed \u2014 just the numbers. One stat per line: Number | Label.',
  fields: [
    { key: 'items', label: 'Stats (exactly 4)', type: 'repeater', addLabel: 'Add stat',
      parts: [
        { key: 'number', label: 'Number', type: 'text' },
        { key: 'label', label: 'Label', type: 'text' }
      ],
      default: '1,200+ | Undergrad researchers\n94% | Career placement\n#7 | Best value, Midwest\n150+ | Student orgs' },
    { key: 'palette', label: 'Color combo', type: 'select', options: ['Bold Blocks', 'Bordered Bold', 'Monochrome Forest', 'High Contrast'], default: 'Bold Blocks' }
  ],
  render: (v, b) => {
    const boldSwatches = [
      { bg: '#154734', num: '#ffffff', label: b.accent },
      { bg: b.accent, num: '#154734', label: '#154734' },
      { bg: '#007A33', num: '#ffffff', label: '#ffffff' },
      { bg: '#154734', num: '#ffffff', label: b.accent }
    ];
    const palettes = {
      'Bold Blocks': { border: null, swatches: boldSwatches },
      'Bordered Bold': { border: '#154734', swatches: [
        { bg: '#e7f1e7', num: '#154734', label: '#007A33' },
        { bg: '#e7f1e7', num: '#154734', label: '#007A33' },
        { bg: '#e7f1e7', num: '#154734', label: '#007A33' },
        { bg: '#e7f1e7', num: '#154734', label: '#007A33' }
      ] },
      'Monochrome Forest': { border: null, swatches: [
        { bg: '#154734', num: '#ffffff', label: 'rgba(255,255,255,0.75)' },
        { bg: '#007A33', num: '#ffffff', label: 'rgba(255,255,255,0.85)' },
        { bg: '#72BF44', num: '#154734', label: '#154734' },
        { bg: '#BFD730', num: '#154734', label: '#154734' }
      ] },
      'High Contrast': { border: '#154734', swatches: [
        { bg: '#154734', num: '#ffffff', label: b.accent },
        { bg: '#ffffff', num: '#154734', label: '#48663f' },
        { bg: '#ffffff', num: '#154734', label: '#48663f' },
        { bg: '#154734', num: '#ffffff', label: b.accent }
      ] }
    };
    const pal = palettes[v.palette] || palettes['Bold Blocks'];
    const swatches = pal.swatches;
    const rows = lines(v.items).slice(0, 4);
    const cell = (p, si) => {
      const sw = swatches[si % swatches.length];
      const borderStyle = pal.border ? `border:1px solid ${pal.border};` : '';
      return `<td width="50%" valign="top" style="width:50%;box-sizing:border-box;background-color:${sw.bg};${borderStyle}padding:30px 24px;font-family:${FF};"><div style="font-size:32px;font-weight:900;color:${sw.num};letter-spacing:-1px;">${p[0] || ''}</div><div style="margin-top:6px;font-size:11px;font-weight:700;letter-spacing:0.6px;text-transform:uppercase;color:${sw.label};">${p[1] || ''}</div></td>`;
    };
    // Real <table>/<td> per row (not inline-block divs) so both cells in a row always match
    // height \u2014 inline-block left a stray gap under the shorter cell whenever one label
    // wrapped to more lines than its neighbor.
    const row = (a, ai, b, bi) => `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;table-layout:fixed;"><tr>${cell(a, ai)}${cell(b, bi)}</tr></table>`;
    const top = row(rows[0] || [], 0, rows[1] || [], 1);
    const bottom = row(rows[2] || [], 2, rows[3] || [], 3);
    return `<!-- ===== Missouri S&T · Checkerboard Stat Wall ===== -->
<div style="width:100%;">${top}${bottom}</div>`;
  }
});

// __APPEND__