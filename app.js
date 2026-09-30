(function(){
  "use strict";

  /* ---------- vocabulary ---------- */

  // State uses the reserved status palette. Priority never borrows these hues.
  var STATUS = {
    'To be Initiated': { color:'var(--st-neutral)',  order:0 },
    'In-Progress':     { color:'var(--st-warning)',  order:1 },
    'Delayed':         { color:'var(--st-critical)', order:2 },
    'On Hold':         { color:'var(--st-serious)',  order:3 },
    'Completed':       { color:'var(--st-good)',     order:4 },
    'Not set':         { color:'var(--st-unset)',    order:5 }
  };
  function statusMeta(l){ return STATUS[l] || { color:'var(--accent)', order:6 }; }

  // Priority is one hue, shallow to deep, plus a filled-bar count so it reads without colour.
  var PRIORITY = {
    'High':    { color:'var(--pri-high)', bars:3, order:0 },
    'Medium':  { color:'var(--pri-med)',  bars:2, order:1 },
    'Low':     { color:'var(--pri-low)',  bars:1, order:2 },
    'Not set': { color:'var(--pri-none)', bars:0, order:3 }
  };
  function priorityMeta(l){ return PRIORITY[l] || { color:'var(--accent)', bars:0, order:4 }; }

  var STACK_ORDER = ['To be Initiated','In-Progress','Delayed','On Hold','Completed','Not set'];
  var SHEET_WANTED = 'dashboard';

  var FILTER_DEFS = [
    { key:'workstream', label:'Workstream' },
    { key:'owner',      label:'Owner' },
    { key:'status',     label:'Status' },
    { key:'priority',   label:'Priority' },
    { key:'timeline',   label:'Timeline' }
  ];
  var DIMS = FILTER_DEFS.map(function(f){ return f.key; });

  /* ---------- state ---------- */

  var dataset = [];
  var known = {}, chosen = {};
  DIMS.forEach(function(d){ known[d] = new Set(); chosen[d] = new Set(); });

  var fileHandle = null, lastModified = null, pollTimer = null;

  /* ---------- helpers ---------- */

  function $(id){ return document.getElementById(id); }
  function el(tag, cls){ var e = document.createElement(tag); if (cls) e.className = cls; return e; }
  function txt(tag, cls, s){ var e = el(tag, cls); e.textContent = s; return e; }

  function banner(msg, kind){
    $('banner').className = 'banner show ' + (kind || 'info');
    $('bannerText').textContent = msg;
  }
  function clearBanner(){
    $('banner').className = 'banner';
    $('bannerText').textContent = '';
  }
  function setPill(mode, label){
    $('statusDot').className = 'dot' + (mode === 'live' ? ' live' : '');
    $('statusText').textContent = label;
  }
  function stamp(iso){
    var d = new Date(iso);
    if (isNaN(d.getTime())) return String(iso);
    return d.toLocaleDateString(undefined, { day:'numeric', month:'short' }) + ', ' +
           d.toLocaleTimeString(undefined, { hour:'2-digit', minute:'2-digit' });
  }

  /* ---------- workbook parsing ---------- */

  function normHeader(s){ return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g,''); }

  function headerRowIndex(rows){
    for (var i = 0; i < Math.min(rows.length, 15); i++){
      var n = (rows[i] || []).map(normHeader);
      if (n.some(function(c){ return c.indexOf('status') !== -1; }) &&
          n.some(function(c){ return c.indexOf('owner') !== -1; }) &&
          n.some(function(c){ return c.indexOf('actionitem') !== -1; })) return i;
    }
    return -1;
  }

  // Columns are found by header text, not by position, so inserting or moving a
  // column in the sheet does not break the page. The SI column is ignored by design.
  function mapColumns(headerRow){
    var n = headerRow.map(normHeader);
    function find(pred){ for (var i = 0; i < n.length; i++){ if (pred(n[i])) return i; } return -1; }
    return {
      workstream:    find(function(c){ return c.indexOf('workstream') !== -1 || c.indexOf('worksream') !== -1; }),
      subWorkstream: find(function(c){ return c.indexOf('subworkstream') !== -1 || c.indexOf('subworksream') !== -1; }),
      actionItem:    find(function(c){ return c.indexOf('actionitem') !== -1; }),
      owner:         find(function(c){ return c.indexOf('owner') !== -1 && c.indexOf('bcg') === -1; }),
      status:        find(function(c){ return c.indexOf('status') !== -1; }),
      priority:      find(function(c){ return c.indexOf('priority') !== -1; }),
      timeline:      find(function(c){ return c.indexOf('timeline') !== -1; }),
      remarks:       find(function(c){ return c.indexOf('remark') !== -1; })
    };
  }

  function cell(row, i){
    if (i == null || i < 0 || i >= row.length) return '';
    var v = row[i];
    return v == null ? '' : String(v).trim();
  }

  function canonStatus(raw){
    var s = String(raw || '').trim();
    if (!s) return 'Not set';
    var k = s.toLowerCase().replace(/[\s\-_]+/g,' ').trim();
    if (k === 'completed' || k === 'complete' || k === 'done') return 'Completed';
    if (k === 'in progress' || k === 'inprogress' || k === 'ongoing') return 'In-Progress';
    if (k === 'delayed' || k === 'delay' || k === 'overdue' || k === 'behind schedule' || k === 'slipped') return 'Delayed';
    if (k === 'on hold' || k === 'onhold' || k === 'hold') return 'On Hold';
    if (k === 'to be initiated' || k === 'not started' || k === 'to be started' || k === 'not initiated') return 'To be Initiated';
    return s;
  }

  function canonPriority(raw){
    var s = String(raw || '').trim();
    if (!s || s === '-') return 'Not set';
    var k = s.toLowerCase().replace(/[\s\-_]+/g,'');
    if (k === 'high' || k === 'h' || k === 'hi' || k === 'p1' || k === 'critical') return 'High';
    if (k === 'medium' || k === 'med' || k === 'm' || k === 'p2' || k === 'moderate') return 'Medium';
    if (k === 'low' || k === 'l' || k === 'p3') return 'Low';
    return s;
  }

  function orNotSet(v){ return (!v || v === '-' || v === '—') ? 'Not set' : v; }

  function parseItems(rows){
    var h = headerRowIndex(rows);
    if (h === -1) throw new Error('Could not find the header row - expected columns such as Workstream, Owner and Status.');
    var c = mapColumns(rows[h]);
    var out = [];
    for (var r = h + 1; r < rows.length; r++){
      var row = rows[r] || [];
      var ws = cell(row, c.workstream), sub = cell(row, c.subWorkstream), act = cell(row, c.actionItem);
      if (!ws && !sub && !act) continue;
      out.push({
        workstream:    ws || 'Not set',
        subWorkstream: sub,
        actionItem:    act,
        owner:         orNotSet(cell(row, c.owner)),
        status:        canonStatus(cell(row, c.status)),
        priority:      canonPriority(cell(row, c.priority)),
        timeline:      cell(row, c.timeline) || 'Not set',
        remarks:       cell(row, c.remarks)
      });
    }
    return out;
  }

  /* ---------- ordering ---------- */

  function timelineKey(v){
    if (!v || v === 'Not set') return [9, 9];
    if (/tbd/i.test(v)) return [3, 0];
    var wk = /wk\s*(\d+)/i.exec(v);
    if (wk){
      var months = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12 };
      var mm = /([a-z]{3})/i.exec(v);
      var mo = mm ? (months[mm[1].toLowerCase()] || 0) : 0;
      return [0, mo * 10 + parseInt(wk[1], 10)];
    }
    return [2, 0];
  }

  function sortValues(dim, values){
    var arr = Array.from(values);
    if (dim === 'status')   return arr.sort(function(a,b){ return statusMeta(a).order - statusMeta(b).order; });
    if (dim === 'priority') return arr.sort(function(a,b){ return priorityMeta(a).order - priorityMeta(b).order; });
    if (dim === 'timeline') return arr.sort(function(a,b){
      var ka = timelineKey(a), kb = timelineKey(b);
      if (ka[0] !== kb[0]) return ka[0] - kb[0];
      if (ka[1] !== kb[1]) return ka[1] - kb[1];
      return String(a).localeCompare(String(b));
    });
    return arr.sort(function(a,b){
      if (a === 'Not set') return 1;
      if (b === 'Not set') return -1;
      return String(a).localeCompare(String(b), undefined, { sensitivity:'base' });
    });
  }

  function hasRealPriorities(){
    return Array.from(known.priority).some(function(v){ return v !== 'Not set'; });
  }

  /* ---------- filter state ---------- */

  function syncFilters(items){
    DIMS.forEach(function(dim){
      var present = new Set(items.map(function(it){ return it[dim]; }));
      present.forEach(function(v){
        if (!known[dim].has(v)){ known[dim].add(v); chosen[dim].add(v); }
      });
      Array.from(known[dim]).forEach(function(v){
        if (!present.has(v)){ known[dim]['delete'](v); chosen[dim]['delete'](v); }
      });
    });
  }

  function visibleItems(){
    return dataset.filter(function(it){
      return DIMS.every(function(dim){ return chosen[dim].has(it[dim]); });
    });
  }

  // Every filter except one. A control that filters by status should keep showing the
  // full spread of statuses, otherwise selecting one zeroes the rest and there is no
  // way to switch to another without resetting first.
  function visibleItemsExcept(skipDim){
    return dataset.filter(function(it){
      return DIMS.every(function(dim){ return dim === skipDim || chosen[dim].has(it[dim]); });
    });
  }

  function resetFilters(){
    DIMS.forEach(function(dim){ known[dim].forEach(function(v){ chosen[dim].add(v); }); });
    buildFilters();
    render();
  }

  // Only counts as a live selection if there was something else to choose from -
  // otherwise a dimension with a single value would always look filtered.
  function isExclusive(dim, value){
    return known[dim].size > 1 && chosen[dim].size === 1 && chosen[dim].has(value);
  }

  // Clicking a tile, bar segment or bar label narrows the whole dashboard to that
  // slice. Clicking the same thing again widens those dimensions back out, so a
  // selection is always reversible without hunting for the reset button.
  function selectOnly(pairs){
    var alreadyThere = pairs.every(function(p){
      return chosen[p.dim].size === 1 && chosen[p.dim].has(p.value);
    });
    pairs.forEach(function(p){
      if (alreadyThere){
        known[p.dim].forEach(function(v){ chosen[p.dim].add(v); });
      } else {
        chosen[p.dim].clear();
        chosen[p.dim].add(p.value);
      }
    });
    buildFilters();
    render();
  }

  /* ---------- filter controls ---------- */

  function buildFilters(){
    var bar = $('toolbar');
    Array.prototype.slice.call(bar.querySelectorAll('.filter')).forEach(function(n){ n.remove(); });
    var anchor = $('btnReset');

    FILTER_DEFS.forEach(function(def){
      var dim = def.key;

      var wrap = el('div', 'filter');
      var btn = el('button', 'filter-btn');
      btn.type = 'button';
      btn.setAttribute('aria-haspopup', 'true');
      btn.setAttribute('aria-expanded', 'false');
      var value = txt('span', 'fval', '');
      btn.appendChild(txt('span', 'fname', def.label));
      btn.appendChild(value);
      btn.appendChild(txt('span', 'caret', '▾'));

      var panel = el('div', 'filter-panel');
      panel.hidden = true;

      var actions = el('div', 'panel-actions');
      var allBtn = txt('button', null, 'Select all'); allBtn.type = 'button';
      var noneBtn = txt('button', null, 'Clear');     noneBtn.type = 'button';
      actions.appendChild(allBtn); actions.appendChild(noneBtn);
      panel.appendChild(actions);

      var list = el('div', 'panel-options');
      panel.appendChild(list);

      function label(){
        var total = known[dim].size, sel = chosen[dim].size;
        value.textContent = sel === total ? 'All' : (sel === 0 ? 'None' : sel + ' of ' + total);
      }

      function fill(){
        list.innerHTML = '';
        sortValues(dim, known[dim]).forEach(function(v){
          var row = el('label', 'opt');
          var box = el('input');
          box.type = 'checkbox';
          box.checked = chosen[dim].has(v);
          box.addEventListener('change', function(){
            if (box.checked) chosen[dim].add(v); else chosen[dim]['delete'](v);
            label(); render();
          });
          var name = txt('span', null, v);
          name.title = v;
          row.appendChild(box); row.appendChild(name);
          list.appendChild(row);
        });
      }

      allBtn.addEventListener('click', function(){
        known[dim].forEach(function(v){ chosen[dim].add(v); });
        fill(); label(); render();
      });
      noneBtn.addEventListener('click', function(){
        chosen[dim].clear();
        fill(); label(); render();
      });
      btn.addEventListener('click', function(){
        var wasOpen = !panel.hidden;
        closePanels();
        if (!wasOpen){ panel.hidden = false; btn.setAttribute('aria-expanded','true'); }
      });

      fill(); label();
      wrap.appendChild(btn); wrap.appendChild(panel);
      bar.insertBefore(wrap, anchor);
    });
  }

  function closePanels(){
    Array.prototype.slice.call(document.querySelectorAll('.filter-panel')).forEach(function(p){ p.hidden = true; });
    Array.prototype.slice.call(document.querySelectorAll('.filter-btn')).forEach(function(b){ b.setAttribute('aria-expanded','false'); });
  }
  document.addEventListener('click', function(e){
    if (!e.target.closest || !e.target.closest('.filter')) closePanels();
  });

  /* ---------- summary tiles ---------- */

  /* ---------- summary area ---------- */
  // Three interchangeable treatments over one data shape. A version picks its treatment
  // with data-style on the container, so they can be compared without forking the logic.

  function summaryStyle(hostId){
    var h = $(hostId);
    return (h && h.getAttribute('data-style')) || 'tiles';
  }

  function pctOf(count, total){ return total > 0 ? Math.round(100 * count / total) : 0; }

  function statusGroups(){
    var pool = visibleItemsExcept('status');
    return {
      pool: pool, dim: 'status',
      groups: sortValues('status', known.status).map(function(label){
        return {
          label: label,
          count: pool.filter(function(it){ return it.status === label; }).length,
          color: statusMeta(label).color,
          meter: false
        };
      })
    };
  }

  function priorityGroups(){
    var pool = visibleItemsExcept('priority');
    var levels = ['High','Medium','Low'];
    sortValues('priority', known.priority).forEach(function(v){
      if (levels.indexOf(v) === -1) levels.push(v);
    });
    return {
      pool: pool, dim: 'priority',
      groups: levels.map(function(level){
        return {
          label: level,
          count: pool.filter(function(it){ return it.priority === level; }).length,
          color: priorityMeta(level).color,
          meter: true
        };
      })
    };
  }

  function badgeFor(g, cls){
    if (g.meter) return meterFor(g.label);
    var dot = el('i', cls);
    dot.style.background = g.color;
    return dot;
  }

  /* ---- style A (v1): draft-mark tiles ---- */
  function renderTiles(hostId, hero, data){
    var host = $(hostId);
    host.innerHTML = '';

    if (hero){
      var lead = el('button', 'stat lead');
      lead.type = 'button';
      lead.title = 'Show everything';
      lead.appendChild(txt('div', 's-label', hero.label));
      lead.appendChild(txt('div', 's-value', String(hero.value)));
      lead.appendChild(txt('div', 's-share', hero.sub));
      lead.addEventListener('click', resetFilters);
      host.appendChild(lead);
    }

    data.groups.forEach(function(g){
      var share = pctOf(g.count, data.pool.length);
      var cls = 'stat';
      if (g.count === 0) cls += ' is-empty';
      if (g.label === 'Delayed' && g.count > 0) cls += ' flagged';
      if (isExclusive(data.dim, g.label)) cls += ' is-active';

      var tile = el('button', cls);
      tile.type = 'button';
      tile.title = 'Show only ' + g.label;
      tile.style.setProperty('--tile-color', g.color);

      var fill = el('span', 'fill');
      fill.style.height = share + '%';
      var mark = el('span', 'mark');
      mark.style.bottom = share + '%';
      if (share <= 0) mark.hidden = true;
      tile.appendChild(fill);
      tile.appendChild(mark);

      var head = el('div', 's-label');
      head.appendChild(badgeFor(g, 's-dot'));
      head.appendChild(txt('span', null, g.label));
      tile.appendChild(head);
      tile.appendChild(txt('div', 's-value', String(g.count)));
      tile.appendChild(txt('div', 's-share',
        data.pool.length ? share + '% of ' + data.pool.length : 'nothing to show'));
      tile.addEventListener('click', function(){ selectOnly([{ dim:data.dim, value:g.label }]); });
      host.appendChild(tile);
    });
  }

  /* ---- style B (v2): stat rail, with the composition in one pipeline bar ---- */
  function renderRail(hostId, hero, data){
    var host = $(hostId);
    host.innerHTML = '';

    if (hero){
      var lead = el('button', 'rail-cell is-hero');
      lead.type = 'button';
      lead.title = 'Show everything';
      lead.appendChild(txt('span', 'rc-label', hero.label));
      lead.appendChild(txt('span', 'rc-value', String(hero.value)));
      lead.appendChild(txt('span', 'rc-sub', hero.sub));
      lead.addEventListener('click', resetFilters);
      host.appendChild(lead);
    }

    data.groups.forEach(function(g){
      var share = pctOf(g.count, data.pool.length);
      var cls = 'rail-cell';
      if (g.count === 0) cls += ' is-empty';
      if (isExclusive(data.dim, g.label)) cls += ' is-active';

      var cell = el('button', cls);
      cell.type = 'button';
      cell.title = 'Show only ' + g.label;
      cell.style.setProperty('--c', g.color);

      var lab = el('span', 'rc-label');
      lab.appendChild(badgeFor(g, 'rc-dot'));
      lab.appendChild(txt('span', null, g.label));
      cell.appendChild(lab);
      cell.appendChild(txt('span', 'rc-value', String(g.count)));
      cell.appendChild(txt('span', 'rc-sub', data.pool.length ? share + '%' : '–'));
      cell.addEventListener('click', function(){ selectOnly([{ dim:data.dim, value:g.label }]); });
      host.appendChild(cell);
    });

    // One pipeline bar under the headline rail carries the part-to-whole.
    var pipeId = hostId + 'Pipe';
    var existing = $(pipeId);
    if (existing) existing.parentNode.removeChild(existing);
    if (!hero) return;

    var wrap = el('div');
    wrap.id = pipeId;

    var bar = el('div', 'pipeline');
    data.groups.forEach(function(g){
      if (!g.count) return;
      var seg = el('button', 'pipe-seg');
      seg.type = 'button';
      seg.style.background = g.color;
      seg.style.flex = g.count + ' 0 auto';
      seg.setAttribute('aria-label', 'Show only ' + g.label + ', ' + g.count + ' items');
      tipFor(seg, g.count + (g.count === 1 ? ' item' : ' items'), g.label + ' — click to filter');
      seg.addEventListener('click', function(){ selectOnly([{ dim:data.dim, value:g.label }]); });
      bar.appendChild(seg);
    });
    wrap.appendChild(bar);

    var cap = el('div', 'pipe-caption');
    data.groups.forEach(function(g){
      if (!g.count) return;
      var s = el('span');
      var sw = el('i');
      sw.style.background = g.color;
      s.appendChild(sw);
      s.appendChild(txt('span', null, g.label));
      s.appendChild(txt('b', null, String(g.count)));
      cap.appendChild(s);
    });
    wrap.appendChild(cap);
    host.parentNode.appendChild(wrap);
  }

  /* ---- style C (v3): ring with the total in the middle, exact counts in rows ---- */
  function renderRing(hostId, hero, data){
    var host = $(hostId);
    host.innerHTML = '';

    var total = data.pool.length;
    var NS = 'http://www.w3.org/2000/svg';
    var r = 58, cx = 74, cy = 74, circ = 2 * Math.PI * r;
    var narrowed = chosen[data.dim].size === 1 && known[data.dim].size > 1;

    var wrap = el('div', 'ringwrap');
    var svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', '0 0 148 148');
    svg.setAttribute('aria-hidden', 'true');

    var track = document.createElementNS(NS, 'circle');
    track.setAttribute('class', 'ring-track');
    track.setAttribute('cx', cx); track.setAttribute('cy', cy); track.setAttribute('r', r);
    track.setAttribute('stroke-width', '14');
    svg.appendChild(track);

    var offset = 0;
    data.groups.forEach(function(g){
      if (!g.count || !total) return;
      var len = circ * g.count / total;
      var arc = document.createElementNS(NS, 'circle');
      arc.setAttribute('class', 'ring-seg' + (narrowed && !chosen[data.dim].has(g.label) ? ' is-dim' : ''));
      arc.setAttribute('cx', cx); arc.setAttribute('cy', cy); arc.setAttribute('r', r);
      arc.setAttribute('stroke', g.color);
      arc.setAttribute('stroke-width', '14');
      arc.setAttribute('stroke-dasharray', Math.max(len - 2, 1) + ' ' + (circ - Math.max(len - 2, 1)));
      arc.setAttribute('stroke-dashoffset', String(-offset));
      tipFor(arc, g.count + (g.count === 1 ? ' item' : ' items'), g.label + ' — click to filter');
      arc.addEventListener('click', function(){ selectOnly([{ dim:data.dim, value:g.label }]); });
      svg.appendChild(arc);
      offset += len;
    });
    wrap.appendChild(svg);

    var centre = el('div', 'ring-center');
    centre.appendChild(txt('b', null, String(hero ? hero.value : total)));
    centre.appendChild(txt('span', null, hero ? hero.label : 'items'));
    wrap.appendChild(centre);
    host.appendChild(wrap);

    var rows = el('div', 'statrows');
    data.groups.forEach(function(g){
      var share = pctOf(g.count, total);
      var cls = 'statrow';
      if (g.count === 0) cls += ' is-empty';
      if (isExclusive(data.dim, g.label)) cls += ' is-active';
      else if (narrowed) cls += ' is-dim';

      var row = el('button', cls);
      row.type = 'button';
      row.title = 'Show only ' + g.label;
      row.style.setProperty('--c', g.color);

      var name = el('span', 'sr-name');
      name.appendChild(badgeFor(g, 'sr-dot'));
      name.appendChild(txt('span', null, g.label));
      row.appendChild(name);

      var trk = el('span', 'sr-track');
      var fill = el('i');
      fill.style.width = share + '%';
      trk.appendChild(fill);
      row.appendChild(trk);

      row.appendChild(txt('span', 'sr-count', String(g.count)));
      row.appendChild(txt('span', 'sr-pct', total ? share + '%' : '–'));
      row.addEventListener('click', function(){ selectOnly([{ dim:data.dim, value:g.label }]); });
      rows.appendChild(row);
    });
    host.appendChild(rows);
  }

  function renderSummary(hostId, hero, data){
    var style = summaryStyle(hostId);
    if (style === 'rail') renderRail(hostId, hero, data);
    else if (style === 'ring') renderRing(hostId, hero, data);
    else renderTiles(hostId, hero, data);
  }

  function renderStats(items){
    renderSummary('stats', {
      label: 'Action items',
      value: items.length,
      sub: items.length === dataset.length ? 'Whole tracker' : 'of ' + dataset.length + ' in tracker'
    }, statusGroups());
  }

  function renderPriorityStats(){
    renderSummary('priorityStats', null, priorityGroups());
  }

  /* ---------- stacked bars ---------- */

  function tipFor(node, headline, detail){
    var tip = $('tip');
    function show(e){
      tip.innerHTML = '';
      tip.appendChild(txt('b', null, headline));
      tip.appendChild(txt('span', null, detail));
      tip.classList.add('show');
      place(e);
    }
    function place(e){
      var x, y;
      if (e && typeof e.clientX === 'number'){ x = e.clientX; y = e.clientY; }
      else { var r = node.getBoundingClientRect(); x = r.left + r.width / 2; y = r.top; }
      tip.style.left = Math.min(x + 14, window.innerWidth - 260) + 'px';
      tip.style.top  = Math.max(y - 44, 8) + 'px';
    }
    function hide(){ tip.classList.remove('show'); }
    node.addEventListener('mouseenter', show);
    node.addEventListener('mousemove', place);
    node.addEventListener('mouseleave', hide);
    node.addEventListener('focus', show);
    node.addEventListener('blur', hide);
  }

  function groupBy(items, dim){
    var map = {};
    items.forEach(function(it){
      if (!map[it[dim]]) map[it[dim]] = {};
      map[it[dim]][it.status] = (map[it[dim]][it.status] || 0) + 1;
    });
    return Object.keys(map).map(function(name){
      var counts = map[name];
      var total = Object.keys(counts).reduce(function(s,k){ return s + counts[k]; }, 0);
      return { name:name, counts:counts, total:total };
    });
  }

  function renderBars(hostId, groups, dim, emptyMessage){
    var host = $(hostId);
    host.innerHTML = '';

    if (!groups.length){
      host.appendChild(txt('p', 'sub', emptyMessage));
      return;
    }

    var max = groups.reduce(function(m,g){ return Math.max(m, g.total); }, 0) || 1;
    var order = STACK_ORDER.slice();
    sortValues('status', known.status).forEach(function(s){
      if (order.indexOf(s) === -1) order.push(s);
    });

    groups.forEach(function(g){
      var row = el('div', 'bar-row');
      if (known[dim].size > 1 && chosen[dim].size === 1){
        row.classList.add(chosen[dim].has(g.name) ? 'is-sel' : 'is-dim');
      }

      var name = txt('button', 'bar-name', g.name);
      name.type = 'button';
      name.title = 'Show only ' + g.name;
      name.addEventListener('click', function(){ selectOnly([{ dim:dim, value:g.name }]); });

      var track = el('div', 'bar-track');
      track.style.width = (100 * g.total / max) + '%';

      order.forEach(function(status){
        var n = g.counts[status];
        if (!n) return;
        var seg = el('button', 'seg');
        seg.type = 'button';
        seg.style.background = statusMeta(status).color;
        seg.style.flex = n + ' 0 auto';
        seg.setAttribute('aria-label', 'Show the ' + n + ' ' + status + ' items in ' + g.name);
        tipFor(seg, n + (n === 1 ? ' item' : ' items'), status + ' · ' + g.name + ' — click to filter');
        seg.addEventListener('click', function(){
          selectOnly([{ dim:dim, value:g.name }, { dim:'status', value:status }]);
        });
        track.appendChild(seg);
      });

      row.appendChild(name);
      row.appendChild(track);
      row.appendChild(txt('div', 'bar-total', String(g.total)));
      host.appendChild(row);
    });
  }

  function renderStatusLegend(){
    var host = $('statusLegend');
    host.innerHTML = '';
    sortValues('status', known.status).forEach(function(label){
      var item = el('div', 'legend-item');
      var sw = el('span', 'legend-swatch');
      sw.style.background = statusMeta(label).color;
      item.appendChild(sw);
      item.appendChild(txt('span', null, label));
      host.appendChild(item);
    });
  }

  /* ---------- priority ---------- */

  function meterFor(level){
    var meta = priorityMeta(level);
    var meter = el('span', 'meter');
    meter.setAttribute('aria-hidden', 'true');
    for (var i = 1; i <= 3; i++){
      var bar = el('i');
      if (i <= meta.bars) bar.style.background = meta.color;
      meter.appendChild(bar);
    }
    return meter;
  }

  function priorityCell(level){
    var wrap = el('span', 'pri pri-' + (level === 'Not set' ? 'none' : level.toLowerCase()));
    wrap.appendChild(meterFor(level));
    wrap.appendChild(txt('span', 'pri-label', level === 'Not set' ? '—' : level));
    return wrap;
  }

  function renderPriority(){
    var card = $('priorityCard');
    if (!hasRealPriorities()){ card.hidden = true; return; }
    card.hidden = false;
    var items = visibleItemsExcept('priority');

    var legend = $('priorityLegend');
    legend.innerHTML = '';
    sortValues('priority', known.priority).forEach(function(level){
      var item = el('div', 'legend-item');
      item.appendChild(meterFor(level));
      item.appendChild(txt('span', null, level));
      legend.appendChild(item);
    });

    var groups = groupBy(items, 'priority');
    groups.sort(function(a,b){ return priorityMeta(a.name).order - priorityMeta(b.name).order; });
    renderBars('priorityBars', groups, 'priority', 'No action items match the selected filters.');
  }

  /* ---------- sorting the table ---------- */

  // The list reads as a workstream outline by default; any column can take over,
  // and workstream/sub-workstream stay the tiebreak underneath whatever is chosen.
  var sortState = { key:'workstream', dir:'asc' };

  function isBlank(v){ return !v || v === 'Not set'; }

  function rawCmp(key, a, b){
    var av = a[key], bv = b[key];
    if (key === 'status')   return statusMeta(av).order - statusMeta(bv).order;
    if (key === 'priority') return priorityMeta(av).order - priorityMeta(bv).order;
    if (key === 'timeline'){
      var ka = timelineKey(av), kb = timelineKey(bv);
      return (ka[0] - kb[0]) || (ka[1] - kb[1]);
    }
    return String(av || '').localeCompare(String(bv || ''), undefined, { sensitivity:'base' });
  }

  function sortItems(items){
    var key = sortState.key;
    var dir = sortState.dir === 'desc' ? -1 : 1;
    return items.slice().sort(function(a, b){
      // Empty cells sit at the bottom whichever way the column is pointed.
      var ab = isBlank(a[key]), bb = isBlank(b[key]);
      if (ab !== bb) return ab ? 1 : -1;

      var r = rawCmp(key, a, b) * dir;
      if (r) return r;
      if (key !== 'workstream'){ r = rawCmp('workstream', a, b); if (r) return r; }
      if (key !== 'subWorkstream'){ r = rawCmp('subWorkstream', a, b); if (r) return r; }
      return rawCmp('actionItem', a, b);
    });
  }

  function updateSortIndicators(){
    Array.prototype.slice.call(document.querySelectorAll('thead th[data-key]')).forEach(function(th){
      var active = sortState.key === th.getAttribute('data-key');
      th.setAttribute('aria-sort', active ? (sortState.dir === 'asc' ? 'ascending' : 'descending') : 'none');
      th.classList.toggle('is-sorted', active);
      var arrow = th.querySelector('.th-arrow');
      if (arrow) arrow.textContent = active ? (sortState.dir === 'asc' ? '▲' : '▼') : '↕';
    });
  }

  function initSortHeaders(){
    Array.prototype.slice.call(document.querySelectorAll('thead th[data-key]')).forEach(function(th){
      var key = th.getAttribute('data-key');
      var label = th.textContent.trim();
      th.textContent = '';
      var btn = el('button', 'th-sort');
      btn.type = 'button';
      btn.title = 'Sort by ' + label;
      btn.appendChild(txt('span', null, label));
      btn.appendChild(txt('span', 'th-arrow', '↕'));
      btn.addEventListener('click', function(){
        if (sortState.key === key) sortState.dir = (sortState.dir === 'asc' ? 'desc' : 'asc');
        else { sortState.key = key; sortState.dir = 'asc'; }
        render();
      });
      th.appendChild(btn);
    });
    updateSortIndicators();
  }

  /* ---------- table ---------- */

  function renderTable(rawItems){
    var body = $('tableBody');
    body.innerHTML = '';
    var items = sortItems(rawItems);
    updateSortIndicators();

    if (!items.length){
      var tr = el('tr', 'empty');
      var td = el('td');
      td.colSpan = 8;
      td.textContent = 'No action items match the selected filters.';
      tr.appendChild(td);
      body.appendChild(tr);
    } else {
      items.forEach(function(it){
        var delayed = it.status === 'Delayed';
        var tr = el('tr', delayed ? 'is-delayed' : null);

        // A blank cell in the sheet reads better as a dash than as the word "Not set",
        // which is kept only as a filter value.
        function shown(v){ return (!v || v === 'Not set') ? '—' : v; }
        function cls(base, v){ return (!v || v === 'Not set') ? base + ' c-blank' : base; }

        var wsTd = txt('td', 'c-stream', it.workstream);
        wsTd.title = it.workstream;
        tr.appendChild(wsTd);
        tr.appendChild(txt('td', 'c-stream', shown(it.subWorkstream)));
        tr.appendChild(txt('td', 'c-action', shown(it.actionItem)));
        var ownerTd = txt('td', cls('c-owner', it.owner), shown(it.owner));
        ownerTd.title = it.owner;
        tr.appendChild(ownerTd);

        var stateTd = el('td');
        var state = el('span', 'state' + (delayed ? ' is-delayed' : ''));
        var dot = el('span', 'state-dot');
        dot.style.background = statusMeta(it.status).color;
        state.appendChild(dot);
        state.appendChild(txt('span', null, it.status));
        stateTd.appendChild(state);
        tr.appendChild(stateTd);

        var priTd = el('td');
        priTd.appendChild(priorityCell(it.priority));
        tr.appendChild(priTd);

        tr.appendChild(txt('td', cls('c-time', it.timeline), shown(it.timeline)));

        var remarks = el('td', 'c-remarks');
        remarks.appendChild(txt('span', null, it.remarks || '—'));
        if (it.remarks) remarks.title = it.remarks;
        tr.appendChild(remarks);

        body.appendChild(tr);
      });
    }

    $('tableSub').textContent = items.length + ' of ' + dataset.length +
      (dataset.length === 1 ? ' action item' : ' action items') + ' shown';
  }

  /* ---------- render ---------- */

  function render(){
    var items = visibleItems();
    renderStats(items);
    renderPriorityStats();
    renderBars('bars',
               groupBy(visibleItemsExcept('workstream'), 'workstream').sort(function(a,b){ return b.total - a.total; }),
               'workstream', 'No action items match the selected filters.');
    renderStatusLegend();
    renderPriority();
    renderTable(items);
    $('resultCount').textContent = 'Showing ' + items.length + ' of ' + dataset.length;
  }

  function applyRows(rows){
    dataset = parseItems(rows);
    syncFilters(dataset);
    buildFilters();
    render();
    $('dash').hidden = false;
  }

  /* ---------- data sources ---------- */

  function loadSnapshot(){
    var snap = window.__VPPL_SNAPSHOT__;
    if (!snap || !snap.rows || !snap.rows.length){
      setPill('idle', 'No data');
      banner('This page was published without any tracker data in it. It needs to be rebuilt from the workbook.', 'warning');
      return;
    }
    try {
      applyRows(snap.rows);
      setPill('idle', 'Tracker saved ' + stamp(snap.fileTime || snap.stamp));
    } catch (err){
      setPill('idle', 'Could not read data');
      banner('Could not read the published data: ' + err.message, 'error');
    }
  }

  function sheetNameIn(workbook){
    var names = workbook.SheetNames, i;
    for (i = 0; i < names.length; i++){ if (names[i].trim().toLowerCase() === SHEET_WANTED) return names[i]; }
    for (i = 0; i < names.length; i++){ if (names[i].toLowerCase().indexOf(SHEET_WANTED) !== -1) return names[i]; }
    return null;
  }

  function readWorkbook(file){
    return file.arrayBuffer().then(function(buf){
      var wb = XLSX.read(buf, { type:'array' });
      var name = sheetNameIn(wb);
      if (!name) throw new Error('That workbook has no sheet named "Dashboard".');
      var rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header:1, defval:'', raw:false });
      lastModified = file.lastModified;
      applyRows(rows);
      clearBanner();
      return true;
    })['catch'](function(err){
      banner('Could not read that workbook: ' + err.message, 'error');
      return false;
    });
  }

  function readHandle(){ return fileHandle.getFile().then(readWorkbook); }

  function startWatching(){
    stopWatching();
    pollTimer = setInterval(function(){
      if (!fileHandle) return;
      fileHandle.getFile().then(function(file){
        if (file.lastModified !== lastModified){
          return readWorkbook(file).then(function(ok){
            if (ok) setPill('live', 'Live · saved ' + stamp(new Date(file.lastModified).toISOString()));
          });
        }
      })['catch'](function(){
        stopWatching();
        fileHandle = null;
        setPill('idle', 'Live connection lost');
        banner('Lost access to your local workbook. The figures above are from the last successful read.', 'warning');
        $('btnRefresh').hidden = true;
        $('btnLive').hidden = false;
      });
    }, 4000);
  }
  function stopWatching(){ if (pollTimer){ clearInterval(pollTimer); pollTimer = null; } }

  function goLive(handle){
    fileHandle = handle;
    return readHandle().then(function(ok){
      if (!ok){ fileHandle = null; return; }
      setPill('live', 'Live · saved ' + stamp(new Date(lastModified).toISOString()));
      $('btnRefresh').hidden = false;
      $('btnLive').hidden = true;
      $('lblLoadCopy').hidden = true;
      startWatching();
    });
  }

  function pickLiveFile(){
    window.showOpenFilePicker({
      id: 'vppl-tracker',
      types: [{ description:'Excel Workbook', accept:{ 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx'] } }],
      multiple: false
    }).then(function(handles){ return goLive(handles[0]); })
      ['catch'](function(e){
        if (e && e.name === 'AbortError') return;
        $('btnLive').hidden = true;
        $('lblLoadCopy').hidden = false;
        banner('This browser will not let an embedded page watch a file for changes. Use "Load my copy" to read the workbook once instead.', 'warning');
      });
  }

  function wireFilePicker(){
    $('filePicker').addEventListener('change', function(e){
      var f = e.target.files && e.target.files[0];
      if (!f) return;
      stopWatching();
      fileHandle = null;
      $('btnRefresh').hidden = true;
      readWorkbook(f).then(function(ok){
        if (!ok) return;
        setPill('idle', 'Your copy · read ' + stamp(new Date().toISOString()));
        banner('Showing your local copy of the workbook. This was a one-off read - pick the file again to pull in later edits.', 'info');
      });
    });
  }

  function wireTheme(){
    var saved = null;
    try { saved = localStorage.getItem('vppl-theme'); } catch(e){}
    if (saved === 'light' || saved === 'dark') document.documentElement.setAttribute('data-theme', saved);
    $('btnTheme').addEventListener('click', function(){
      var cur = document.documentElement.getAttribute('data-theme');
      var next = cur === 'dark' ? 'light'
               : cur === 'light' ? null
               : (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'light' : 'dark');
      if (next) document.documentElement.setAttribute('data-theme', next);
      else document.documentElement.removeAttribute('data-theme');
      try { next ? localStorage.setItem('vppl-theme', next) : localStorage.removeItem('vppl-theme'); } catch(e){}
    });
  }

  function init(){
    var boot = $('bootWarning');
    if (boot) boot.parentNode.removeChild(boot);

    wireTheme();
    wireFilePicker();
    initSortHeaders();
    $('btnReset').addEventListener('click', resetFilters);
    $('btnLive').addEventListener('click', pickLiveFile);
    $('btnRefresh').addEventListener('click', function(){ if (fileHandle) readHandle(); });

    loadSnapshot();

    var framed = false;
    try { framed = window.self !== window.top; } catch(e){ framed = true; }
    if (('showOpenFilePicker' in window) && !framed) $('btnLive').hidden = false;
    else $('lblLoadCopy').hidden = false;
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
