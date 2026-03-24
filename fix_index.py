with open('public/index.html', 'r', encoding='utf-8') as f:
    content = f.read()

original_len = len(content)

# ============================================================
# 1. Fix renderScorecards function
# ============================================================
old_sc = '''function renderScorecards() {
  const data = getDisplayStores();
  if (!data.length) return;
  const n = parseInt(document.getElementById('hypoCount').value)||2;
  const avgIST = data.reduce((a,s)=>a+(s.in_store||0),0)/data.length;
  const avgHypo = data.reduce((a,s)=>a+calcHypo(s.in_store||0,s.deliveries,n),0)/data.length;
  const avgMk = data.reduce((a,s)=>a+parseTime(s.make),0)/data.length;
  const avgPr = data.reduce((a,s)=>a+parseTime(s.production),0)/data.length;
  const avgPct4 = data.reduce((a,s)=>a+parsePct(s.pct_lt4),0)/data.length;
  const avgPct15 = data.reduce((a,s)=>a+parsePct(s.pct_lt15),0)/data.length;
  const avgOT = data.reduce((a,s)=>a+parsePct(s.on_time),0)/data.length;
  const allScoped = scopeStores(currentData.stores);
  const saved = avgIST - avgHypo;

  document.getElementById('scorecardStrip').innerHTML = `
    <div class="score-tile"><div class="score-tile-label">Stores</div><div class="score-tile-value val-neutral">${data.length}</div><div class="score-tile-sub">${activeDay==='wtd'?'WTD view':'Day view'}</div></div>
    <div class="score-tile"><div class="score-tile-label">Avg In-Store</div><div class="score-tile-value ${istColor(avgIST)}">${avgIST.toFixed(1)}m</div><div class="score-tile-sub">Target <19m</div></div>
    <div class="score-tile"><div class="score-tile-label">If Target (-${n})</div><div class="score-tile-value val-neutral">${avgHypo.toFixed(1)}m</div><div class="score-tile-sub">${saved>0?'\\u2193 '+saved.toFixed(1)+'m possible':'On target'}</div></div>
    <div class="score-tile"><div class="score-tile-label">Avg Make</div><div class="score-tile-value ${mkColor(avgMk)}">${fmtTime(avgMk)}</div><div class="score-tile-sub">${avgPct4.toFixed(0)}% under 4:00</div></div>
    <div class="score-tile"><div class="score-tile-label">Avg Production</div><div class="score-tile-value ${prColor(avgPr)}">${fmtTime(avgPr)}</div><div class="score-tile-sub">${avgPct15.toFixed(0)}% under 15:00</div></div>
    <div class="score-tile"><div class="score-tile-label">On-Time %</div><div class="score-tile-value ${pctColor(avgOT,85,75)}">${avgOT.toFixed(1)}%</div><div class="score-tile-sub">Target >85%</div></div>
  `;
  document.getElementById('storeCount').textContent = `${data.length} stores`;
}'''

new_sc = '''function renderScorecards() {
  const data = getDisplayStores();
  if (!data.length) return;
  const avgIST = data.reduce((a,s)=>a+(s.in_store||0),0)/data.length;
  const avgMk = data.reduce((a,s)=>a+parseTime(s.make),0)/data.length;
  const avgPct4 = data.reduce((a,s)=>a+parsePct(s.pct_lt4),0)/data.length;
  const avgOT = data.reduce((a,s)=>a+parsePct(s.on_time),0)/data.length;
  const totalLt10 = data.reduce((a,s)=>a+(parseInt(s.ist_lt10)||0),0);
  const total1014 = data.reduce((a,s)=>a+(parseInt(s.ist_1014)||0),0);
  const total1518 = data.reduce((a,s)=>a+(parseInt(s.ist_1518)||0),0);
  const total1925 = data.reduce((a,s)=>a+(parseInt(s.ist_1925)||0),0);
  const totalGt25 = data.reduce((a,s)=>a+(parseInt(s.ist_gt25)||0),0);

  document.getElementById('scorecardStrip').innerHTML = `
    <div class="score-tile"><div class="score-tile-label">Stores</div><div class="score-tile-value val-neutral">${data.length}</div><div class="score-tile-sub">${activeDay==='wtd'?'WTD view':'Day view'}</div></div>
    <div class="score-tile"><div class="score-tile-label">Avg In-Store</div><div class="score-tile-value ${istColor(avgIST)}">${avgIST.toFixed(1)}m</div><div class="score-tile-sub">Target <19m</div></div>
    <div class="score-tile"><div class="score-tile-label">Avg Make</div><div class="score-tile-value ${mkColor(avgMk)}">${fmtTime(avgMk)}</div><div class="score-tile-sub">${avgPct4.toFixed(0)}% under 4:00</div></div>
    <div class="score-tile"><div class="score-tile-label">On-Time %</div><div class="score-tile-value ${pctColor(avgOT,85,75)}">${avgOT.toFixed(1)}%</div><div class="score-tile-sub">Target >85%</div></div>
    <div class="score-tile"><div class="score-tile-label"><10# / 10-14# / 15-18#</div><div class="score-tile-value val-good" style="font-size:18px">${totalLt10} / ${total1014} / ${total1518}</div><div class="score-tile-sub">Fast IST buckets</div></div>
    <div class="score-tile"><div class="score-tile-label">19-25# / >25#</div><div class="score-tile-value ${total1925+totalGt25>data.length*3?'val-bad':total1925+totalGt25>data.length?'val-warn':'val-good'}" style="font-size:18px">${total1925} / ${totalGt25}</div><div class="score-tile-sub">Slow IST buckets</div></div>
  `;
  document.getElementById('storeCount').textContent = `${data.length} stores`;
}'''

if old_sc in content:
    content = content.replace(old_sc, new_sc, 1)
    print("✓ renderScorecards replaced")
else:
    print("✗ renderScorecards NOT FOUND")
    # Find the function and print the exact characters including quotes
    idx = content.find('function renderScorecards')
    if idx >= 0:
        chunk = content[idx:idx+300]
        for i, ch in enumerate(chunk):
            if i < 50 or ord(ch) not in range(32, 127):
                print(f"  [{i}] {ord(ch):3d} {repr(ch)}")

# ============================================================
# 2. Fix toggleAlert - remove prod reference
# ============================================================
old_ta = "document.getElementById('alertBarLabel').textContent = alertFilter==='make' ? '\u26a0 Filtered: Make Time > 4:00' : '\u26a0 Filtered: Production Time > 20:00';"
new_ta = "document.getElementById('alertBarLabel').textContent = '\u26a0 Filtered: Make Time > 4:00';"
if old_ta in content:
    content = content.replace(old_ta, new_ta, 1)
    print("✓ toggleAlert replaced")
else:
    print("✗ toggleAlert NOT FOUND")

# ============================================================
# 3. Fix storeRow - remove If Target, Delta, Production, %<15m; add <10#, 10-14#, 15-18#
# ============================================================
old_sr = """  function storeRow(s) {
    const hypo = calcHypo(s.in_store||0, s.deliveries, n);
    const delta = (s.in_store||0) - hypo;
    const mm = parseTime(s.make), pm = parseTime(s.production);
    const lt4 = parsePct(s.pct_lt4), lt15 = parsePct(s.pct_lt15), ot = parsePct(s.on_time);
    const wtdBadge = isWTD && s.days_reported > 1 ? `<span style="font-size:9px;color:var(--green);margin-left:4px">${s.days_reported}d avg</span>` : '';
    return `<tr>
      <td><div class="store-name">${s.name}${wtdBadge}</div><div class="store-id">${s.store_id}</div></td>
      <td style="font-size:11px;color:var(--muted)">${s.area||''}</td>
      <td style="font-size:10px;color:var(--muted2)">${s.region_coach||''}</td>
      <td><span class="ist-val ${istColor(s.in_store||0)}">${(s.in_store||0).toFixed ? (s.in_store).toFixed(1) : s.in_store}m</span></td>
      <td><span style="color:var(--accent);font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:700">${hypo.toFixed(1)}m</span></td>
      <td><span class="delta-pill ${delta>0.5?'delta-pos':'delta-zero'}">\\u2193${delta.toFixed(1)}</span></td>
      <td><span class="${mkColor(mm)}" style="font-family:'Space Mono',monospace;font-size:11px">${s.make||'\\u2014'}</span></td>
      <td><div class="pct-bar"><span class="${pctColor(lt4,80,60)}" style="font-size:11px;min-width:30px">${s.pct_lt4||'\\u2014'}</span><div class="bar-track"><div class="bar-fill" style="width:${lt4}%;background:${lt4>=80?'var(--green)':lt4>=60?'var(--yellow)':'var(--red)'}"></div></div></div></td>
      <td><span class="${prColor(pm)}" style="font-family:'Space Mono',monospace;font-size:11px">${s.production||'\\u2014'}</span></td>
      <td><div class="pct-bar"><span class="${pctColor(lt15,80,60)}" style="font-size:11px;min-width:30px">${s.pct_lt15||'\\u2014'}</span><div class="bar-track"><div class="bar-fill" style="width:${lt15}%;background:${lt15>=80?'var(--green)':lt15>=60?'var(--yellow)':'var(--red)'}"></div></div></div></td>
      <td><span class="${(s.ist_1925||0)>3?'val-bad':(s.ist_1925||0)>1?'val-warn':'val-good'}" style="font-family:'Space Mono',monospace;font-size:11px">${s.ist_1925!=null?parseInt(s.ist_1925):0}</span></td>
      <td><span class="${(s.ist_gt25||0)>3?'val-bad':(s.ist_gt25||0)>1?'val-warn':'val-good'}" style="font-family:'Space Mono',monospace;font-size:11px">${s.ist_gt25!=null?parseInt(s.ist_gt25):0}</span></td>
      <td><span class="${pctColor(ot,85,75)}" style="font-size:11px">${s.on_time||'\\u2014'}</span></td>
      <td style="font-size:11px;color:var(--muted)">${s.deliveries||0}</td>
    </tr>`;
  }"""

new_sr = """  function storeRow(s) {
    const mm = parseTime(s.make);
    const lt4 = parsePct(s.pct_lt4), ot = parsePct(s.on_time);
    const wtdBadge = isWTD && s.days_reported > 1 ? `<span style="font-size:9px;color:var(--green);margin-left:4px">${s.days_reported}d avg</span>` : '';
    return `<tr>
      <td><div class="store-name">${s.name}${wtdBadge}</div><div class="store-id">${s.store_id}</div></td>
      <td style="font-size:11px;color:var(--muted)">${s.area||''}</td>
      <td style="font-size:10px;color:var(--muted2)">${s.region_coach||''}</td>
      <td><span class="ist-val ${istColor(s.in_store||0)}">${(s.in_store||0).toFixed ? (s.in_store).toFixed(1) : s.in_store}m</span></td>
      <td><span class="${mkColor(mm)}" style="font-family:'Space Mono',monospace;font-size:11px">${s.make||'\\u2014'}</span></td>
      <td><div class="pct-bar"><span class="${pctColor(lt4,80,60)}" style="font-size:11px;min-width:30px">${s.pct_lt4||'\\u2014'}</span><div class="bar-track"><div class="bar-fill" style="width:${lt4}%;background:${lt4>=80?'var(--green)':lt4>=60?'var(--yellow)':'var(--red)'}"></div></div></div></td>
      <td><span class="${(s.ist_lt10||0)>3?'val-warn':(s.ist_lt10||0)>0?'val-good':'val-neutral'}" style="font-family:'Space Mono',monospace;font-size:11px">${s.ist_lt10!=null?parseInt(s.ist_lt10):0}</span></td>
      <td><span class="${(s.ist_1014||0)>3?'val-warn':(s.ist_1014||0)>0?'val-good':'val-neutral'}" style="font-family:'Space Mono',monospace;font-size:11px">${s.ist_1014!=null?parseInt(s.ist_1014):0}</span></td>
      <td><span class="${(s.ist_1518||0)>3?'val-warn':(s.ist_1518||0)>0?'val-good':'val-neutral'}" style="font-family:'Space Mono',monospace;font-size:11px">${s.ist_1518!=null?parseInt(s.ist_1518):0}</span></td>
      <td><span class="${(s.ist_1925||0)>3?'val-bad':(s.ist_1925||0)>1?'val-warn':'val-good'}" style="font-family:'Space Mono',monospace;font-size:11px">${s.ist_1925!=null?parseInt(s.ist_1925):0}</span></td>
      <td><span class="${(s.ist_gt25||0)>3?'val-bad':(s.ist_gt25||0)>1?'val-warn':'val-good'}" style="font-family:'Space Mono',monospace;font-size:11px">${s.ist_gt25!=null?parseInt(s.ist_gt25):0}</span></td>
      <td><span class="${pctColor(ot,85,75)}" style="font-size:11px">${s.on_time||'\\u2014'}</span></td>
      <td style="font-size:11px;color:var(--muted)">${s.deliveries||0}</td>
    </tr>`;
  }"""

if old_sr in content:
    content = content.replace(old_sr, new_sr, 1)
    print("✓ storeRow replaced")
else:
    print("✗ storeRow NOT FOUND")
    idx = content.find('function storeRow')
    if idx >= 0:
        print(repr(content[idx:idx+200]))

# ============================================================
# 4. Fix groupRow - remove If Target, Delta, Production, %<15m; add <10#, 10-14#, 15-18#
# ============================================================
old_gr = """  function groupRow(stores, label, cls) {
    if (!stores.length) return '';
    const avgIST=stores.reduce((a,s)=>a+(s.in_store||0),0)/stores.length;
    const avgHypo=stores.reduce((a,s)=>a+calcHypo(s.in_store||0,s.deliveries,n),0)/stores.length;
    const avgMk=stores.reduce((a,s)=>a+parseTime(s.make),0)/stores.length;
    const avgPr=stores.reduce((a,s)=>a+parseTime(s.production),0)/stores.length;
    const avgLt4=stores.reduce((a,s)=>a+parsePct(s.pct_lt4),0)/stores.length;
    const avgLt15=stores.reduce((a,s)=>a+parsePct(s.pct_lt15),0)/stores.length;
    const avgOT=stores.reduce((a,s)=>a+parsePct(s.on_time),0)/stores.length; 
    const total1925=stores.reduce((a,s)=>a+(parseInt(s.ist_1925)||0),0); const totalGt25=stores.reduce((a,s)=>a+(parseInt(s.ist_gt25)||0),0);
    const mf=stores.filter(s=>parseTime(s.make)>4).length;
    const pf=stores.filter(s=>parseTime(s.production)>20).length;
    const delta=avgIST-avgHypo;
    return `<tr class="group-row ${cls}">
      <td colspan="3">${label} <span style="opacity:.4;font-size:9px">${stores.length} stores</span></td>
      <td><span class="ist-val ${istColor(avgIST)}">${avgIST.toFixed(1)}m</span></td>
      <td><span style="color:var(--accent);font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:700">${avgHypo.toFixed(1)}m</span></td>
      <td><span class="delta-pill delta-pos">\\u2193${delta.toFixed(1)}</span></td>
      <td class="${mkColor(avgMk)}">${fmtTime(avgMk)}</td>
      <td class="${pctColor(avgLt4,80,60)}">${avgLt4.toFixed(0)}%</td>
      <td class="${prColor(avgPr)}">${fmtTime(avgPr)}</td>
      <td class="${pctColor(avgLt15,80,60)}">${avgLt15.toFixed(0)}%</td> 
      <td class="${total1925>stores.length*3?'val-bad':total1925>stores.length?'val-warn':'val-good'}">${total1925}</td>
      <td class="${totalGt25>stores.length*2?'val-bad':totalGt25>stores.length?'val-warn':'val-good'}">${totalGt25}</td>
      <td class="${pctColor(avgOT,85,75)}">${avgOT.toFixed(1)}%</td>
      <td>${stores.reduce((a,s)=>a+(s.deliveries||0),0)}</td>
      <td style="font-size:9px;color:var(--muted)">${mf}mk/${pf}pr</td>
    </tr>`;
  }"""

new_gr = """  function groupRow(stores, label, cls) {
    if (!stores.length) return '';
    const avgIST=stores.reduce((a,s)=>a+(s.in_store||0),0)/stores.length;
    const avgMk=stores.reduce((a,s)=>a+parseTime(s.make),0)/stores.length;
    const avgLt4=stores.reduce((a,s)=>a+parsePct(s.pct_lt4),0)/stores.length;
    const avgOT=stores.reduce((a,s)=>a+parsePct(s.on_time),0)/stores.length;
    const totalLt10=stores.reduce((a,s)=>a+(parseInt(s.ist_lt10)||0),0);
    const total1014=stores.reduce((a,s)=>a+(parseInt(s.ist_1014)||0),0);
    const total1518=stores.reduce((a,s)=>a+(parseInt(s.ist_1518)||0),0);
    const total1925=stores.reduce((a,s)=>a+(parseInt(s.ist_1925)||0),0);
    const totalGt25=stores.reduce((a,s)=>a+(parseInt(s.ist_gt25)||0),0);
    const mf=stores.filter(s=>parseTime(s.make)>4).length;
    return `<tr class="group-row ${cls}">
      <td colspan="3">${label} <span style="opacity:.4;font-size:9px">${stores.length} stores</span></td>
      <td><span class="ist-val ${istColor(avgIST)}">${avgIST.toFixed(1)}m</span></td>
      <td class="${mkColor(avgMk)}">${fmtTime(avgMk)}</td>
      <td class="${pctColor(avgLt4,80,60)}">${avgLt4.toFixed(0)}%</td>
      <td>${totalLt10}</td>
      <td>${total1014}</td>
      <td>${total1518}</td>
      <td class="${total1925>stores.length*3?'val-bad':total1925>stores.length?'val-warn':'val-good'}">${total1925}</td>
      <td class="${totalGt25>stores.length*2?'val-bad':totalGt25>stores.length?'val-warn':'val-good'}">${totalGt25}</td>
      <td class="${pctColor(avgOT,85,75)}">${avgOT.toFixed(1)}%</td>
      <td>${stores.reduce((a,s)=>a+(s.deliveries||0),0)}</td>
      <td style="font-size:9px;color:var(--muted)">${mf}mk slow</td>
    </tr>`;
  }"""

if old_gr in content:
    content = content.replace(old_gr, new_gr, 1)
    print("✓ groupRow replaced")
else:
    print("✗ groupRow NOT FOUND")
    idx = content.find('function groupRow')
    if idx >= 0:
        print(repr(content[idx:idx+200]))

# ============================================================
# 5. Fix table headers (cols) - remove If Target, Delta, Production, %<15m; add <10#, 10-14#, 15-18#
# ============================================================
old_cols = """  const cols = `<tr>
    <th onclick="sort('name')">Store</th><th onclick="sort('area')">Area</th><th onclick="sort('region_coach')">Region</th>
    <th onclick="sort('in_store')" class="${sortCol==='in_store'?'sort-'+sortDir:''}">${isWTD?'WTD IST':'In-Store'}</th>
    <th>If Target (-${n})</th><th>Delta</th>
    <th onclick="sort('make')" class="${sortCol==='make'?'sort-'+sortDir:''}">${isWTD?'WTD Make':'Make'}</th>
    <th onclick="sort('pct_lt4')" class="${sortCol==='pct_lt4'?'sort-'+sortDir:''}">%<4m</th>
    <th onclick="sort('production')" class="${sortCol==='production'?'sort-'+sortDir:''}">${isWTD?'WTD Prod':'Production'}</th>
    <th onclick="sort('pct_lt15')" class="${sortCol==='pct_lt15'?'sort-'+sortDir:''}">%<15m</th>
    <th onclick="sort('ist_1925')" class="${sortCol==='ist_1925'?'sort-'+sortDir:''}">19-25#</th>
    <th onclick="sort('ist_gt25')" class="${sortCol==='ist_gt25'?'sort-'+sortDir:''}">\\u003e25#</th>
    <th onclick="sort('on_time')" class="${sortCol==='on_time'?'sort-'+sortDir:''}">On-Time</th>
    <th>Del</th>
  </tr>`;"""

new_cols = """  const cols = `<tr>
    <th onclick="sort('name')">Store</th><th onclick="sort('area')">Area</th><th onclick="sort('region_coach')">Region</th>
    <th onclick="sort('in_store')" class="${sortCol==='in_store'?'sort-'+sortDir:''}">${isWTD?'WTD IST':'In-Store'}</th>
    <th onclick="sort('make')" class="${sortCol==='make'?'sort-'+sortDir:''}">${isWTD?'WTD Make':'Make'}</th>
    <th onclick="sort('pct_lt4')" class="${sortCol==='pct_lt4'?'sort-'+sortDir:''}">%<4m</th>
    <th onclick="sort('ist_lt10')" class="${sortCol==='ist_lt10'?'sort-'+sortDir''}"><10#</th>
    <th onclick="sort('ist_1014')" class="${sortCol==='ist_1014'?'sort-'+sortDir:''}">10-14#</th>
    <th onclick="sort('ist_1518')" class="${sortCol==='ist_1518'?'sort-'+sortDir:''}">15-18#</th>
    <th onclick="sort('ist_1925')" class="${sortCol==='ist_1925'?'sort-'+sortDir:''}">19-25#</th>
    <th onclick="sort('ist_gt25')" class="${sortCol==='ist_gt25'?'sort-'+sortDir:''}">\\u003e25#</th>
    <th onclick="sort('on_time')" class="${sortCol==='on_time'?'sort-'+sortDir:''}">On-Time</th>
    <th>Del</th>
  </tr>`;"""

if old_cols in content:
    content = content.replace(old_cols, new_cols, 1)
    print("✓ table headers replaced")
else:
    print("✗ table headers NOT FOUND")
    idx = content.find('const cols =')
    if idx >= 0:
        print(repr(content[idx:idx+300]))

# ============================================================
# 6. Fix sort comparator - add ist_lt10/1014/1518; remove production/pct_lt15
# ============================================================
old_sort = "    if(['in_store','deliveries','ist_1925','ist_gt25'].includes(sortCol)){av=a[sortCol]||0;bv=b[sortCol]||0;}\n    else if(['make','production'].includes(sortCol)){av=parseTime(a[sortCol]);bv=parseTime(b[sortCol]);}\n    else if(['pct_lt4','pct_lt15','on_time'].includes(sortCol)){av=parsePct(a[sortCol]);bv=parsePct(b[sortCol]);}"

new_sort = "    if(['in_store','deliveries','ist_lt10','ist_1014','ist_1518','ist_1925','ist_gt25'].includes(sortCol)){av=a[sortCol]||0;bv=b[sortCol]||0;}\n    else if(['make'].includes(sortCol)){av=parseTime(a[sortCol]);bv=parseTime(b[sortCol]);}\n    else if(['pct_lt4','on_time'].includes(sortCol)){av=parsePct(a[sortCol]);bv=parsePct(b[sortCol]);}"

if old_sort in content:
    content = content.replace(old_sort, new_sort, 1)
    print("✓ sort comparator replaced")
else:
    print("✗ sort comparator NOT FOUND")
    idx = content.find("['in_store','deliveries'")
    if idx >= 0:
        print(repr(content[idx:idx+200]))

# ============================================================
# 7. Remove hypo box from toolbar
# ============================================================
old_hypo = """  <div class="hypo-box">
    <span class="hypo-label">⚡ Replace</span>
    <input type="number" id="hypoCount" value="2" min="1" max="10" onchange="renderTable()">
    <span class="hypo-label">slow orders → target</span>
  </div>"""

new_hypo = """  <input type="hidden" id="hypoCount" value="2">"""

if old_hypo in content:
    content = content.replace(old_hypo, new_hypo, 1)
    print("✓ hypo box removed from toolbar")
else:
    print("✗ hypo box NOT FOUND")
    idx = content.find('hypo-box')
    if idx >= 0:
        print(repr(content[idx-20:idx+100]))

# ============================================================
# 8. Update table subtitle hint text
# ============================================================
old_hint = '<span style="font-size:9px;color:var(--muted)">Make <4:00 · Prod <15:00 · IST target <19m</span>'
new_hint = '<span style="font-size:9px;color:var(--muted)">Make <4:00 · In-Store target <19m · PDF buckets</span>'

if old_hint in content:
    content = content.replace(old_hint, new_hint, 1)
    print("✓ subtitle hint updated")
else:
    print("✗ subtitle hint NOT FOUND")

# ============================================================
# 9. Fix autoAnalyze - remove production references
# ============================================================
old_ai = "  const wPr=[...data].sort((a,b)=>parseTime(b.production)-parseTime(a.production)).slice(0,5).map(s=>`${s.name}(${s.production})`).join(', ');\n  const days=currentData.days.length;\n  const period=currentData.period||'';\n  const dayLabel=activeDay==='wtd'?`${days}-day WTD average`:'single day';\n  runAI(`You are Velocity, a Pizza Hut Speed of Service agent. ${currentUser} is viewing ${dayLabel} data for ${period}. Be direct, ops-focused, no fluff.\\n\\nDATA (${data.length} stores, ${dayLabel}):\\n- Worst In-Store: ${wIST}\\n- Worst Make (threshold <4:00): ${wMk}\\n- Worst Production (threshold <15:00): ${wPr}\\n- Hypothetical: replacing ${n} slowest orders with 19-min target\\n\\nGive: 1) Top 3 critical findings, 2) Immediate coaching priorities, 3) Make vs Production pattern, 4) What's working.`);"

new_ai = "  const days=currentData.days.length;\n  const period=currentData.period||'';\n  const dayLabel=activeDay==='wtd'?`${days}-day WTD average`:'single day';\n  const totalGt25=data.reduce((a,s)=>a+(parseInt(s.ist_gt25)||0),0);\n  const total1925=data.reduce((a,s)=>a+(parseInt(s.ist_1925)||0),0);\n  runAI(`You are Velocity, a Pizza Hut Speed of Service agent. ${currentUser} is viewing ${dayLabel} data for ${period}. Be direct, ops-focused, no fluff.\\n\\nDATA (${data.length} stores, ${dayLabel}):\\n- Worst In-Store Time: ${wIST}\\n- Worst Make Time (threshold <4:00): ${wMk}\\n- Stores with >25min IST orders: ${totalGt25} total, 19-25min: ${total1925} total\\n\\nGive: 1) Top 3 critical findings, 2) Immediate coaching priorities, 3) Make time pattern, 4) What's working.`);"

if old_ai in content:
    content = content.replace(old_ai, new_ai, 1)
    print("✓ autoAnalyze replaced")
else:
    print("✗ autoAnalyze NOT FOUND")
    idx = content.find('wPr=')
    if idx >= 0:
        print(repr(content[idx-50:idx+200]))

# ============================================================
# 10. Fix getFilteredStores - remove production/pct_lt15; add ist_lt10/1014/1518
# ============================================================
old_gfs = """  if (activeDay === 'wtd') {
    stores = currentData.stores.map(s => ({
      ...s,
      in_store: s.wtd_in_store, make: s.wtd_make, pct_lt4: s.wtd_pct_lt4,
      production: s.wtd_production, pct_lt15: s.wtd_pct_lt15,
      on_time: s.wtd_on_time, deliveries: s.wtd_deliveries,
      ist_1925: s.wtd_ist_1925||0, ist_gt25: s.wtd_ist_gt25||0,
    }));
  } else {
    stores = currentData.stores.map(s => {
      const day = s.daily && s.daily[activeDay];
      if (!day) return null;
      return { ...s, in_store: day.in_store, make: day.make,
        pct_lt4: day.pct_lt4||s.pct_lt4, production: day.production,
        pct_lt15: day.pct_lt15||s.pct_lt15, on_time: day.on_time,
        deliveries: day.deliveries, ist_1925: day.ist_1925||0, ist_gt25: day.ist_gt25||0 };
    }).filter(Boolean);
  }"""

new_gfs = """  if (activeDay === 'wtd') {
    stores = currentData.stores.map(s => ({
      ...s,
      in_store: s.wtd_in_store, make: s.wtd_make, pct_lt4: s.wtd_pct_lt4,
      on_time: s.wtd_on_time, deliveries: s.wtd_deliveries,
      ist_lt10: s.wtd_ist_lt10||0, ist_1014: s.wtd_ist_1014||0, ist_1518: s.wtd_ist_1518||0,
      ist_1925: s.wtd_ist_1925||0, ist_gt25: s.wtd_ist_gt25||0,
    }));
  } else {
    stores = currentData.stores.map(s => {
      const day = s.daily && s.daily[activeDay];
      if (!day) return null;
      return { ...s, in_store: day.in_store, make: day.make,
        pct_lt4: day.pct_lt4||s.pct_lt4, on_time: day.on_time,
        deliveries: day.deliveries,
        ist_lt10: day.ist_lt10||0, ist_1014: day.ist_1014||0, ist_1518: day.ist_1518||0,
        ist_1925: day.ist_1925||0, ist_gt25: day.ist_gt25||0 };
    }).filter(Boolean);
  }"""

if old_gfs in content:
    content = content.replace(old_gfs, new_gfs, 1)
    print("✓ getFilteredStores replaced")
else:
    print("✗ getFilteredStores NOT FOUND")
    idx = content.find('function getFilteredStores')
    if idx >= 0:
        print(repr(content[idx:idx+300]))

# ============================================================
# 11. Fix Excel export - update column widths, headers, area row, store row
# ============================================================
# Fix column widths (14 cols -> 13 cols: remove 2 cols for If Target+Delta+Prod+%<15, add 3 for <10/10-14/15-18)
old_widths = "  [22,18,18,10,10,8,9,8,11,8,8,7,10,7].forEach((w,i) => { ws.getColumn(i+1).width = w; });"
new_widths = "  [22,18,18,10,9,8,7,7,7,7,7,10,7].forEach((w,i) => { ws.getColumn(i+1).width = w; });"
if old_widths in content:
    content = content.replace(old_widths, new_widths, 1)
    print("✓ Excel col widths replaced")
else:
    print("✗ Excel col widths NOT FOUND")

# Fix title row merge
old_merge = "  ws.mergeCells('A1:N1');"
new_merge = "  ws.mergeCells('A1:M1');"
if old_merge in content:
    content = content.replace(old_merge, new_merge, 1)
    print("✓ Excel merge cells replaced")
else:
    print("✗ Excel merge cells NOT FOUND")

# Fix title text
old_title_text = "  t.value = `${stores.length} stores  |  ${period} ${dateRange}  |  Make <4:00 - Prod <15:00 - IST target <19m`;"
new_title_text = "  t.value = `${stores.length} stores  |  ${period} ${dateRange}  |  Make <4:00 - IST target <19m`;"
if old_title_text in content:
    content = content.replace(old_title_text, new_title_text, 1)
    print("✓ Excel title text replaced")
else:
    print("✗ Excel title text NOT FOUND")

# Fix headers row
old_hdrs = "  const hdr = ws.addRow(['Store','Area','Region','In Store','IF Target','Delta','Make','%<4m','Production','%<15m','19-25','> 25','%On Time','Del']);"
new_hdrs = "  const hdr = ws.addRow(['Store','Area','Region','In Store','Make','%<4m','<10#','10-14#','15-18#','19-25#','> 25#','%On Time','Del']);"
if old_hdrs in content:
    content = content.replace(old_hdrs, new_hdrs, 1)
    print("✓ Excel headers replaced")
else:
    print("✗ Excel headers NOT FOUND")

# Fix area summary row
old_area_row = """    const gr = ws.addRow([
      `${coach}  ${aStores.length} stores`, '', '',
      avgIST.toFixed(1)+'m', avgHypo.toFixed(1)+'m', String.fromCharCode(8595)+grpDelta.toFixed(1),
      fmtTime(avgMk), avgLt4/100, fmtTime(avgPr), avgLt15/100,
      tot1925, totGt25, avgOT/100, totDel, `${mf}mk/${pf}pr`
    ]);
    gr.eachCell(c => { c.fill = BLUE_FILL; c.font = { bold: true, size: 10 }; });
    gr.getCell(4).font  = { bold:true, size:10, color: istFontColor(avgIST) };
    gr.getCell(5).font  = { bold:true, size:10, color: ORANGE };
    gr.getCell(6).font  = { bold:true, size:10, color: GREEN };
    gr.getCell(8).numFmt  = '0%';
    gr.getCell(10).numFmt = '0%';
    gr.getCell(13).numFmt = '0%';
    gr.getCell(15).font = { color: GRAY, size: 9 };"""

new_area_row = """    const totLt10 = aStores.reduce((a,s)=>a+(parseInt(s.ist_lt10)||0),0);
    const tot1014 = aStores.reduce((a,s)=>a+(parseInt(s.ist_1014)||0),0);
    const tot1518 = aStores.reduce((a,s)=>a+(parseInt(s.ist_1518)||0),0);
    const gr = ws.addRow([
      `${coach}  ${aStores.length} stores`, '', '',
      avgIST.toFixed(1)+'m', fmtTime(avgMk), avgLt4/100,
      totLt10, tot1014, tot1518, tot1925, totGt25, avgOT/100, totDel
    ]);
    gr.eachCell(c => { c.fill = BLUE_FILL; c.font = { bold: true, size: 10 }; });
    gr.getCell(4).font  = { bold:true, size:10, color: istFontColor(avgIST) };
    gr.getCell(5).font  = { bold:true, size:10, color: mkFont ? mkFont(parseTime(avgMk)) : GREEN };
    gr.getCell(6).numFmt  = '0%';
    gr.getCell(12).numFmt = '0%';"""

if old_area_row in content:
    content = content.replace(old_area_row, new_area_row, 1)
    print("✓ Excel area row replaced")
else:
    print("✗ Excel area row NOT FOUND")
    idx = content.find('const gr = ws.addRow')
    if idx >= 0:
        print(repr(content[idx:idx+200]))

# Fix store row in export
old_store_row_xl = """      const sr = ws.addRow([
        s.name||s.store_id, s.area||s.area_coach||'', s.region_coach||'',
        ist.toFixed(1)+'m', hypo.toFixed(1)+'m', String.fromCharCode(8595)+sdelta.toFixed(1),
        s.make||'', lt4>0?lt4/100:null,
        s.production||'', lt15>0?lt15/100:null,
        c1925, cgt25, ot>0?ot/100:null, s.deliveries||0
      ]);
      sr.getCell(4).font  = { bold:true, size:10, color: istFontColor(ist) };
      sr.getCell(5).font  = { bold:true, size:10, color: ORANGE };
      sr.getCell(6).font  = { bold:true, size:10, color: GREEN };
      sr.getCell(7).font  = { size:10, color: mk>4?RED:mk>3.5?ORANGE:GREEN };
      sr.getCell(8).font  = { size:10, color: pctFont(lt4) };
      sr.getCell(8).numFmt  = '0%';
      sr.getCell(9).font  = { size:10, color: prFont(pr) };
      sr.getCell(10).font = { size:10, color: pctFont(lt15) };
      sr.getCell(10).numFmt = '0%';
      sr.getCell(11).font = { size:10, color: countFont(c1925) };
      sr.getCell(12).font = { size:10, color: countFont(cgt25) };
      sr.getCell(13).font = { size:10, color: pctFont(ot,85,75) };
      sr.getCell(13).numFmt = '0%';
      sr.getCell(14).font = { size:10, color: GRAY };"""

new_store_row_xl = """      const clt10 = parseInt(s.ist_lt10)||0;
      const c1014 = parseInt(s.ist_1014)||0;
      const c1518 = parseInt(s.ist_1518)||0;
      const sr = ws.addRow([
        s.name||s.store_id, s.area||s.area_coach||'', s.region_coach||'',
        ist.toFixed(1)+'m', s.make||'', lt4>0?lt4/100:null,
        clt10, c1014, c1518, c1925, cgt25, ot>0?ot/100:null, s.deliveries||0
      ]);
      sr.getCell(4).font  = { bold:true, size:10, color: istFontColor(ist) };
      sr.getCell(5).font  = { size:10, color: mk>4?RED:mk>3.5?ORANGE:GREEN };
      sr.getCell(6).font  = { size:10, color: pctFont(lt4) };
      sr.getCell(6).numFmt  = '0%';
      sr.getCell(7).font  = { size:10, color: GREEN };
      sr.getCell(8).font  = { size:10, color: GREEN };
      sr.getCell(9).font  = { size:10, color: YELLOW };
      sr.getCell(10).font = { size:10, color: countFont(c1925) };
      sr.getCell(11).font = { size:10, color: countFont(cgt25) };
      sr.getCell(12).font = { size:10, color: pctFont(ot,85,75) };
      sr.getCell(12).numFmt = '0%';
      sr.getCell(13).font = { size:10, color: GRAY };"""

if old_store_row_xl in content:
    content = content.replace(old_store_row_xl, new_store_row_xl, 1)
    print("✓ Excel store row replaced")
else:
    print("✗ Excel store row NOT FOUND")

# Fix excel export - remove production/hypo vars from store loop
old_xl_vars = """      const ist   = s.in_store || 0;
      const hypo  = calcHypo(ist, s.deliveries, n);
      const mk    = parseTime(s.make);
      const pr    = parseTime(s.production);
      const lt4   = parsePct(s.pct_lt4);
      const lt15  = parsePct(s.pct_lt15);
      const ot    = parsePct(s.on_time);
      const c1925 = parseInt(s.ist_1925)||0;
      const cgt25 = parseInt(s.ist_gt25)||0;
      const sdelta = ist - hypo;"""

new_xl_vars = """      const ist   = s.in_store || 0;
      const mk    = parseTime(s.make);
      const lt4   = parsePct(s.pct_lt4);
      const ot    = parsePct(s.on_time);
      const c1925 = parseInt(s.ist_1925)||0;
      const cgt25 = parseInt(s.ist_gt25)||0;"""

if old_xl_vars in content:
    content = content.replace(old_xl_vars, new_xl_vars, 1)
    print("✓ Excel store vars replaced")
else:
    print("✗ Excel store vars NOT FOUND")

# Fix excel export - remove avgHypo/avgPr/avgLt15/grpDelta from area loop
old_xl_area_vars = """    const avgIST  = aStores.reduce((a,s)=>a+(s.in_store||0),0)/aStores.length;
    const avgHypo = aStores.reduce((a,s)=>a+calcHypo(s.in_store||0,s.deliveries,n),0)/aStores.length;
    const avgMk   = aStores.reduce((a,s)=>a+parseTime(s.make),0)/aStores.length;
    const avgPr   = aStores.reduce((a,s)=>a+parseTime(s.production),0)/aStores.length;
    const avgLt4  = aStores.reduce((a,s)=>a+parsePct(s.pct_lt4),0)/aStores.length;
    const avgLt15 = aStores.reduce((a,s)=>a+parsePct(s.pct_lt15),0)/aStores.length;
    const avgOT   = aStores.reduce((a,s)=>a+parsePct(s.on_time),0)/aStores.length;
    const tot1925 = aStores.reduce((a,s)=>a+(parseInt(s.ist_1925)||0),0);
    const totGt25 = aStores.reduce((a,s)=>a+(parseInt(s.ist_gt25)||0),0);
    const totDel  = aStores.reduce((a,s)=>a+(s.deliveries||0),0);
    const mf = aStores.filter(s=>parseTime(s.make)>4).length;
    const pf = aStores.filter(s=>parseTime(s.production)>=20).length;
    const grpDelta = avgIST - avgHypo;"""

new_xl_area_vars = """    const avgIST  = aStores.reduce((a,s)=>a+(s.in_store||0),0)/aStores.length;
    const avgMk   = aStores.reduce((a,s)=>a+parseTime(s.make),0)/aStores.length;
    const avgLt4  = aStores.reduce((a,s)=>a+parsePct(s.pct_lt4),0)/aStores.length;
    const avgOT   = aStores.reduce((a,s)=>a+parsePct(s.on_time),0)/aStores.length;
    const tot1925 = aStores.reduce((a,s)=>a+(parseInt(s.ist_1925)||0),0);
    const totGt25 = aStores.reduce((a,s)=>a+(parseInt(s.ist_gt25)||0),0);
    const totDel  = aStores.reduce((a,s)=>a+(s.deliveries||0),0);
    const mf = aStores.filter(s=>parseTime(s.make)>4).length;"""

if old_xl_area_vars in content:
    content = content.replace(old_xl_area_vars, new_xl_area_vars, 1)
    print("✓ Excel area vars replaced")
else:
    print("✗ Excel area vars NOT FOUND")

# Fix frozen row count (was 2, still 2, headers are now row 2)
# ws.views is fine as-is

with open('public/index.html', 'w', encoding='utf-8') as f:
    f.write(content)

print(f"\nDone. Length: {original_len} -> {len(content)}")