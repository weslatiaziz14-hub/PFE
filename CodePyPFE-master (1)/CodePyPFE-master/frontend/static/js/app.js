// =============================================================================
// KPI Dashboard - app.js  (updated)
// =============================================================================

'use strict';

// ── Global state ──────────────────────────────────────────────────────────────
let dashboardData  = {};
let charts         = {};
let currentYear    = null;
let currentMonth   = null;
let refreshInterval= null;
let availableYears = [];

function gel(id) { return document.getElementById(id); }
function qs(sel) { return document.querySelector(sel); }
function qsa(sel){ return document.querySelectorAll(sel); }

document.addEventListener('DOMContentLoaded', initializeDashboard);

// =============================================================================
// INITIALIZATION
// =============================================================================

async function initializeDashboard() {
    try {
        await loadAvailableYears();
        await loadDashboardData();
        startAutoRefresh();
        console.log('[Dashboard] Initialized successfully');
    } catch (err) {
        console.error('[Dashboard] Initialization error:', err);
    }
}

// =============================================================================
// DATA LOADING
// =============================================================================

async function loadAvailableYears() {
    try {
        const res = await fetch('/api/available-years', { credentials: 'include' });
        if (!res.ok) return;
        availableYears = await res.json();

        const ycs = gel('ycs');
        if (ycs && availableYears.length) {
            ycs.innerHTML = '';
            availableYears.forEach(y => {
                const b = document.createElement('button');
                b.className = 'yc' + (y === currentYear ? ' active' : '');
                b.textContent = String(y);
                b.onclick = () => {
                    currentYear = y;
                    ycs.querySelectorAll('.yc').forEach(x => x.classList.remove('active'));
                    b.classList.add('active');
                    loadDashboardData(y);
                };
                ycs.appendChild(b);
            });
            if (!currentYear && availableYears.length) {
                currentYear = availableYears[0];
                ycs.querySelector('.yc')?.classList.add('active');
            }
        }
        console.log('[Dashboard] Years loaded:', availableYears);
    } catch (err) {
        console.warn('[Dashboard] Could not load years:', err);
    }
}

async function loadDashboardData(year = currentYear) {
    try {
        const url = year ? `/api/dashboard-data?year=${year}` : '/api/dashboard-data';
        const res  = await fetch(url, { credentials: 'include' });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        dashboardData = await res.json();
        console.log('[Dashboard] Data loaded:', Object.keys(dashboardData));
        renderActiveTab();
        renderAdditionalTables();
        updateRefreshInfo();
    } catch (err) {
        console.error('[Dashboard] Error loading data:', err);
    }
}

// =============================================================================
// TAB RENDERING
// =============================================================================

function renderActiveTab() {
    const active = qs('.sec.active');
    if (!active) return;
    const id = active.id.replace('tab-', '');
    renderTab(id);
}

function renderTab(name) {
    switch(name) {
        case 'global':  renderGlobalCharts();    break;
        case 'cnc':     renderCNCCharts();        break;
        case 'asm':     renderAssemblyCharts();   break;
        case 'fg':      renderFGCharts();         break;
        case 'daily':   renderJournalierCharts(); break;
        case 'extra':   renderAdditionalTables(); break;
    }
}

// =============================================================================
// CHART HELPERS
// =============================================================================

function mkChart(id, cfg) {
    const canvas = gel(id);
    if (!canvas) { console.warn('Canvas not found:', id); return null; }
    if (charts[id]) { charts[id].destroy(); }
    charts[id] = new Chart(canvas.getContext('2d'), cfg);
    return charts[id];
}

// Target line — ALWAYS BLUE (#3b82f6)
function tgtLine(val, label) {
    return {
        type: 'line', label, data: Array(200).fill(val),
        borderColor: '#3b82f6', borderWidth: 2, borderDash: [5, 4],
        pointRadius: 0, tension: 0, fill: false, order: 0
    };
}

// Bar dataset with per-bar color logic
// colorMode: 'rft' = green above target / red below  |  'taux' = red above / green below  |  'fixed' = single color
function barDsColored(label, data, target, colorMode, fixedColor) {
    let bgColors, bdColors;
    if (colorMode === 'rft') {
        bgColors = data.map(v => v === null ? 'rgba(100,100,100,0.3)' : (v >= target ? '#10b981aa' : '#ef4444aa'));
        bdColors = data.map(v => v === null ? '#666' : (v >= target ? '#10b981' : '#ef4444'));
    } else if (colorMode === 'taux') {
        bgColors = data.map(v => v === null ? 'rgba(100,100,100,0.3)' : (v <= target ? '#10b981aa' : '#ef4444aa'));
        bdColors = data.map(v => v === null ? '#666' : (v <= target ? '#10b981' : '#ef4444'));
    } else {
        const c = fixedColor || '#6366f1';
        bgColors = c + 'aa';
        bdColors = c;
    }
    return {
        type: 'bar', label, data,
        backgroundColor: bgColors,
        borderColor: bdColors,
        borderWidth: 1.5, borderRadius: 4, borderSkipped: false
    };
}

// Simple fixed-color bar dataset (for non-KPI charts like FG, pareto)
function barDs(label, data, color) {
    return {
        type: 'bar', label, data,
        backgroundColor: color + 'aa', borderColor: color,
        borderWidth: 1.5, borderRadius: 4, borderSkipped: false
    };
}

function baseOpts(sfx = '%', mn = 0, mx = 105) {
    return {
        responsive: true, maintainAspectRatio: false,
        plugins: {
            legend: { display: true, position: 'top', labels: { font: { size: 11 }, padding: 12 } },
            tooltip: {
                backgroundColor: '#191d2b', borderColor: '#2d3450', borderWidth: 1,
                titleColor: '#e2e8f0', bodyColor: '#94a3b8', padding: 9,
                callbacks: { label: c => ` ${c.dataset.label}: ${c.parsed.y !== undefined ? c.parsed.y + sfx : ''}` }
            }
        },
        scales: {
            y: { min: mn, max: mx, grid: { color: 'rgba(37,42,61,.6)' }, ticks: { color: '#4a5568', callback: v => v + sfx } },
            x: { grid: { color: 'rgba(37,42,61,.6)' }, ticks: { color: '#4a5568', maxRotation: 45 } }
        }
    };
}

function lineCfg(labels, datasets, yMin, yMax) {
    return {
        type: 'line', data: { labels, datasets },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: true, position: 'top' } },
            scales: {
                y: { min: yMin ?? 0, max: yMax, grid: { color: 'rgba(37,42,61,.6)' }, ticks: { color: '#4a5568' } },
                x: { grid: { color: 'rgba(37,42,61,.6)' }, ticks: { color: '#4a5568', maxRotation: 45 } }
            }
        }
    };
}

// =============================================================================
// GLOBAL CHARTS
// =============================================================================

function renderGlobalCharts() {
    const data = dashboardData.rft_global || [];
    if (!data.length) return;

    const filtered = filterByMonth(data);
    const labels   = filtered.map(r => r.period || `${r.mois} ${r.annee}`);
    // [7] Renamed: RFT MMC (formula = 1 - Taux erreur MMC)
    const rftMMC   = filtered.map(r => parseFloat(r.rft_mmc_global_pct)  || 0);
    const rftCNC   = filtered.map(r => parseFloat(r.rft_cnc_pct)         || 0);
    const rftAsm   = filtered.map(r => parseFloat(r.rft_assembly_pct)    || 0);
    const fg       = filtered.map(r => parseFloat(r.finish_good)         || 0);
    // [6] Taux erreur MMC = ((Nbre_déf + Quantité) / Qte) * 100
    const tauxMMC  = filtered.map(r => {
        const qte = parseFloat(r.quantite_produite || r.qte || 0);
        const def = parseFloat(r.nbre_defauts_interne || 0);
        const qty = parseFloat(r.nbre_defauts_externe || 0);
        return qte > 0 ? Math.round((def + qty) / qte * 1000) / 10 : 0;
    });

    // [1][2] RFT chart: blue target, green/red bars
    mkChart('c-gm', {
        type: 'bar',
        data: { labels, datasets: [
            barDsColored('RFT MMC', rftMMC, 80, 'rft'),
            tgtLine(80, 'Objectif 80%')
        ]},
        options: baseOpts('%', 0, 105)
    });

    mkChart('c-gc', {
        type: 'bar',
        data: { labels, datasets: [
            barDsColored('RFT CNC', rftCNC, 94, 'rft'),
            tgtLine(94, 'Objectif 94%')
        ]},
        options: baseOpts('%', 0, 105)
    });

    const aMin = Math.max(0, Math.min(...rftAsm.filter(v => v > 0)) - 2);
    mkChart('c-ga', {
        type: 'bar',
        data: { labels, datasets: [
            barDsColored('RFT Assembly', rftAsm, 90, 'rft'),
            tgtLine(90, 'Objectif 90%')
        ]},
        options: baseOpts('%', aMin, 100.5)
    });

    // [6] Taux erreur MMC chart
    const maxT = Math.max(...tauxMMC.filter(v => !isNaN(v)), 10);
    mkChart('c-taux-mmc', {
        type: 'bar',
        data: { labels, datasets: [
            // [3] Taux erreur: green below target, red above
            barDsColored('Taux Erreur MMC', tauxMMC, 20, 'taux'),
            tgtLine(20, 'Objectif ≤ 20%')
        ]},
        options: baseOpts('%', 0, maxT + 5)
    });

    // Update KPI cards
    if (rftMMC.length)  updateCard('g-mmc',  'kb-mmc',  avg(rftMMC), 80, true);
    if (rftCNC.length)  updateCard('g-cnc',  'kb-cnc',  avg(rftCNC), 94, true);
    if (rftAsm.length)  updateCard('g-asm',  'kb-asm',  avg(rftAsm), 90, true);
    // [6] Taux erreur MMC card
    updateTauxCard('g-taux-mmc', avg(tauxMMC), 20);
    const totalFG = fg.reduce((a, b) => a + b, 0);
    if (gel('g-fg')) gel('g-fg').textContent = totalFG.toLocaleString('fr-FR');

    // Summary table
    const tb = gel('g-tbl');
    if (tb) {
        tb.innerHTML = '';
        filtered.forEach((r, i) => {
            const m  = parseFloat(r.rft_mmc_global_pct) || 0;
            const c  = parseFloat(r.rft_cnc_pct)        || 0;
            const a  = parseFloat(r.rft_assembly_pct)   || 0;
            const tc = parseFloat(r.taux_erreur_cnc_pct)|| 0;
            const ta = parseFloat(r.taux_erreur_assembly_pct) || 0;
            const tm = tauxMMC[i] || 0;
            tb.innerHTML += `<tr>
                <td style="color:var(--txt);font-weight:600">${labels[i]}</td>
                <td>${(parseFloat(r.finish_good)||0).toLocaleString('fr-FR')}</td>
                <td>${r.nbre_defauts_interne||0}</td>
                <td>${r.nbre_defauts_externe||0}</td>
                <td>${(parseFloat(r.finish_good)||0).toLocaleString('fr-FR')}</td>
                <td><span class="pill ${pc(m,80)}">${m.toFixed(1)}%</span></td>
                <td><span class="pill ${pc(c,94)}">${c.toFixed(1)}%</span></td>
                <td><span class="pill ${pc(a,90)}">${a.toFixed(2)}%</span></td>
                <td><span class="pill ${pc(tc,6,true)}">${tc.toFixed(1)}%</span></td>
                <td><span class="pill ${pc(ta,10,true)}">${ta.toFixed(3)}%</span></td>
                <td><span class="pill ${pc(tm,20,true)}">${tm.toFixed(1)}%</span></td>
            </tr>`;
        });
    }
}

// =============================================================================
// CNC CHARTS
// =============================================================================

function renderCNCCharts() {
    const data = dashboardData.cnc_mensuel || [];
    const filtered = filterByMonth(data);
    if (!filtered.length) return;

    const labels  = filtered.map(r => r.period || `${r.mois} ${r.annee}`);
    const rftD    = filtered.map(r => parseFloat(r.rft_cnc_pct)        || 0);
    const tauxD   = filtered.map(r => parseFloat(r.taux_erreur_cnc_pct)|| 0);
    const effD    = filtered.map(r => parseFloat(r.efficacite_cnc_pct) || 0);

    // [2] RFT: green above target, red below — [1] blue target line
    mkChart('c-cnc-rft', {
        type: 'bar',
        data: { labels, datasets: [
            barDsColored('RFT CNC', rftD, 94, 'rft'),
            tgtLine(94, 'Objectif 94%')
        ]},
        options: baseOpts('%', 0, 105)
    });

    // [3] Taux erreur: green below target, red above — [1] blue target line
    const maxT = Math.max(...tauxD.filter(v => !isNaN(v)), 10);
    mkChart('c-cnc-taux', {
        type: 'bar',
        data: { labels, datasets: [
            barDsColored('Taux Erreur CNC', tauxD, 6, 'taux'),
            tgtLine(6, 'Objectif ≤ 6%')
        ]},
        options: baseOpts('%', 0, maxT + 5)
    });

    // Efficacité — fixed color, blue target
    const validEff = effD.filter(v => v > 0);
    mkChart('c-cnc-eff', {
        type: 'bar',
        data: { labels, datasets: [
            barDsColored('Efficacité CNC', effD, 94, 'rft'),
            tgtLine(94, 'Objectif 94%')
        ]},
        options: baseOpts('%', validEff.length ? Math.min(...validEff) - 2 : 88, 100)
    });

    if (gel('c-rft'))  updateCard('c-rft',  'kb-crft', avg(rftD),  94, true);
    if (gel('c-taux')) updateTauxCard('c-taux', avg(tauxD), 6);
    if (gel('c-eff'))  updateCard('c-eff',  'kb-ceff', avg(effD),  94, true);

    // Pareto
    const pd = (dashboardData.cnc_pareto_defauts || []).slice(0, 10);
    if (pd.length) mkChart('c-cnc-pareto', {
        type: 'bar',
        data: { labels: pd.map(r => (r.description_defaut || 'N/A').substring(0, 25)),
                datasets: [barDs('Défauts', pd.map(r => parseFloat(r.nombre_defauts) || 0), '#6366f1')] },
        options: { ...baseOpts(' def', 0), indexAxis: 'y' }
    });

    const pa = (dashboardData.cnc_pareto_actions || []).slice(0, 10);
    if (pa.length) mkChart('c-cnc-act', {
        type: 'bar',
        data: { labels: pa.map(r => (r.action_correction || 'N/A').substring(0, 25)),
                datasets: [barDs('Défauts', pa.map(r => parseFloat(r.nombre_defauts) || 0), '#f59e0b')] },
        options: { ...baseOpts(' def', 0), indexAxis: 'y' }
    });

    // [4] RFT Rate per Process chart
    renderRFTPerProcess(filtered, labels);
}

// =============================================================================
// [4] RFT RATE PER PROCESS — Duplications + Prototypes in same chart
// Data from kpi_cnc_par_statut or computed from cnc data
// =============================================================================

function renderRFTPerProcess(filtered, labels) {
    // Try DB data first (kpi_cnc_par_statut), fall back to RAW computation
    const statData = dashboardData.cnc_par_statut || [];

    // Main CAO processes matching screenshot: Duplication + Prototype
    // We use the inline RAW data for this chart (computed from Suivi_defauts_CNC.xlsx structure)
    // The data is grouped by period × Statu × CAO_norm and stored in cnc_rft_par_process
    const processData = dashboardData.cnc_rft_par_process || dashboardData.rft_par_process || [];

    if (!processData.length) {
        // Fallback: use computed data from cnc_mensuel for now, split arbitrarily
        renderRFTPerProcessFallback(filtered, labels);
        return;
    }

    // Build chart from DB data
    const periods = [...new Set(processData.map(r => r.period || r.mois))].sort();
    const dupData  = periods.map(p => {
        const rows = processData.filter(r => (r.period || r.mois) === p && (r.statu || '').toLowerCase().includes('dup'));
        return rows.length ? Math.round(rows.reduce((s, r) => s + (parseFloat(r.rft_pct) || 0), 0) / rows.length * 10) / 10 : null;
    });
    const protoData = periods.map(p => {
        const rows = processData.filter(r => (r.period || r.mois) === p && (r.statu || '').toLowerCase().includes('proto'));
        return rows.length ? Math.round(rows.reduce((s, r) => s + (parseFloat(r.rft_pct) || 0), 0) / rows.length * 10) / 10 : null;
    });

    buildRFTPerProcessChart(periods, dupData, protoData);
}

function renderRFTPerProcessFallback(filtered, labels) {
    // Use RAW static data arrays (already computed in dashboard.html script block)
    // These are window-level variables set from the static RAW data
    const dupRFT   = (window.PROCESS_DUP_RFT   || []).slice(0, labels.length);
    const protoRFT = (window.PROCESS_PROTO_RFT  || []).slice(0, labels.length);
    buildRFTPerProcessChart(labels, dupRFT, protoRFT);
}

function buildRFTPerProcessChart(labels, dupData, protoData) {
    const tgt94 = 94, tgt85 = 85;
    const canvas = gel('c-rft-process');
    if (!canvas) return;

    // Color per bar based on target
    const dupColors   = dupData.map(v   => v === null ? '#66666688' : (v >= tgt94 ? '#10b981aa' : '#ef4444aa'));
    const protoColors = protoData.map(v => v === null ? '#66666688' : (v >= tgt85 ? '#10b981aa' : '#ef4444aa'));
    const dupBorders   = dupData.map(v   => v === null ? '#666' : (v >= tgt94 ? '#10b981' : '#ef4444'));
    const protoBorders = protoData.map(v => v === null ? '#666' : (v >= tgt85 ? '#10b981' : '#ef4444'));

    mkChart('c-rft-process', {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    type: 'bar', label: 'Duplication',
                    data: dupData,
                    backgroundColor: dupColors, borderColor: dupBorders,
                    borderWidth: 1.5, borderRadius: 4, borderSkipped: false
                },
                {
                    type: 'bar', label: 'Prototype',
                    data: protoData,
                    backgroundColor: protoColors, borderColor: protoBorders,
                    borderWidth: 1.5, borderRadius: 4, borderSkipped: false
                },
                tgtLine(tgt94, `Objectif Dup. ${tgt94}%`),
                { ...tgtLine(tgt85, `Objectif Proto ${tgt85}%`), borderDash: [3, 3], borderColor: '#3b82f6', borderWidth: 1.5 }
            ]
        },
        options: {
            ...baseOpts('%', 0, 105),
            plugins: {
                ...baseOpts('%', 0, 105).plugins,
                legend: { display: true, position: 'top', labels: { font: { size: 11 }, padding: 12 } }
            }
        }
    });
}

// =============================================================================
// ASSEMBLY CHARTS
// =============================================================================

function renderAssemblyCharts() {
    const data     = dashboardData.cf_mensuel || [];
    const filtered = filterByMonth(data);
    if (!filtered.length) return;

    const labels = filtered.map(r => r.period || `${r.mois} ${r.annee}`);
    const rftD   = filtered.map(r => parseFloat(r.rft_assembly_pct)         || 0);
    const tauxD  = filtered.map(r => parseFloat(r.taux_erreur_assembly_pct) || 0);
    const defD   = filtered.map(r => parseFloat(r.nbre_defauts_externe)     || 0);

    const rftMin = Math.max(0, Math.min(...rftD.filter(v => v > 0)) - 2);
    // [2] RFT green/red + [1] blue target
    mkChart('c-asm-rft', {
        type: 'bar',
        data: { labels, datasets: [
            barDsColored('RFT Assembly', rftD, 90, 'rft'),
            tgtLine(90, 'Objectif 90%')
        ]},
        options: baseOpts('%', rftMin, 100.5)
    });

    // [3] Taux erreur green/red + [1] blue target
    const maxT = Math.max(...tauxD.filter(v => !isNaN(v)), 10);
    mkChart('c-asm-taux', {
        type: 'bar',
        data: { labels, datasets: [
            barDsColored('Taux Erreur Assembly', tauxD, 10, 'taux'),
            tgtLine(10, 'Objectif ≤ 10%')
        ]},
        options: baseOpts('%', 0, maxT + 2)
    });

    mkChart('c-asm-def', {
        type: 'bar',
        data: { labels, datasets: [barDs('Défauts abs.', defD, '#f59e0b')] },
        options: baseOpts(' pcs', 0, Math.max(...defD) + 2)
    });

    if (gel('a-rft'))  updateCard('a-rft',  'kb-arft', avg(rftD),  90, true);
    if (gel('a-taux')) updateTauxCard('a-taux', avg(tauxD), 10);
    if (gel('a-def'))  { const e = gel('a-def'); if (e) e.textContent = defD.reduce((a, b) => a + b, 0); }
}

// =============================================================================
// FINISH GOOD CHARTS
// =============================================================================

function renderFGCharts() {
    const daily   = dashboardData.mmc_journalier         || [];
    const monthly = dashboardData.mmc_finishgood_mensuel || [];

    if (daily.length) {
        const labs = daily.map(r => r.date ? new Date(r.date).toLocaleDateString('fr-FR', { day: '2-digit', month: '2-digit' }) : '');
        const vals = daily.map(r => parseFloat(r.qte_fg) || 0);
        mkChart('c-fg-d', {
            type: 'bar',
            data: { labels: labs, datasets: [
                barDs('FG pcs/jour', vals, '#34d399'),
                tgtLine(40, 'Target 40'),
                { ...tgtLine(45, 'Stretch 45'), borderDash: [3, 3] }
            ]},
            options: baseOpts(' pcs', 0, Math.max(...vals, 45) + 5)
        });
    }

    if (monthly.length) {
        const filt = filterByMonth(monthly);
        const labs = filt.map(r => r.period || `${r.mois} ${r.annee}`);
        const vals = filt.map(r => parseFloat(r.finish_good) || 0);
        mkChart('c-fg-m', {
            type: 'bar',
            data: { labels: labs, datasets: [barDs('FG Mensuel', vals, '#34d399')] },
            options: baseOpts(' pcs', 0)
        });

        const agg = {};
        monthly.forEach(r => { agg[r.annee] = (agg[r.annee] || 0) + (parseFloat(r.finish_good) || 0); });
        const yrs  = Object.keys(agg).sort();
        const yVals = yrs.map(y => agg[y]);
        mkChart('c-fg-y', {
            type: 'bar',
            data: { labels: yrs, datasets: [barDs('FG Annuel', yVals, '#34d399')] },
            options: baseOpts(' pcs', 0)
        });
    }
}

// =============================================================================
// JOURNALIER CHARTS
// =============================================================================

function renderJournalierCharts() {
    const cnc = filterByDate(dashboardData.cnc_journalier || []);
    const asm = filterByDate(dashboardData.cf_journalier  || []);

    if (cnc.length) {
        const labs   = cnc.map(r => fmt(r.date));
        const rftCNC = cnc.map(r => parseFloat(r.rft_cnc_pct)         || 0);
        const tauCNC = cnc.map(r => parseFloat(r.taux_erreur_cnc_pct) || 0);
        const effCNC = cnc.map(r => parseFloat(r.efficacite_cnc_pct)  || 0);

        // [2] RFT charts: green/red + blue target
        mkChart('c-d-mmc', {
            type: 'bar',
            data: { labels: labs, datasets: [
                barDsColored('RFT MMC', cnc.map(r => parseFloat(r.rft_cnc_pct) || 0), 80, 'rft'),
                tgtLine(80, 'Objectif')
            ]},
            options: baseOpts('%', 0, 105)
        });

        mkChart('c-d-cnc', {
            type: 'bar',
            data: { labels: labs, datasets: [
                barDsColored('RFT CNC', rftCNC, 94, 'rft'),
                tgtLine(94, 'Objectif')
            ]},
            options: baseOpts('%', 0, 105)
        });

        // [3] Taux: green/red + blue target
        mkChart('c-d-tcnc', {
            type: 'bar',
            data: { labels: labs, datasets: [
                barDsColored('Taux Err CNC', tauCNC, 6, 'taux'),
                tgtLine(6, 'Objectif')
            ]},
            options: baseOpts('%', 0)
        });

        mkChart('c-d-eff', {
            type: 'bar',
            data: { labels: labs, datasets: [
                barDsColored('Efficacité CNC', effCNC, 94, 'rft'),
                tgtLine(94, 'Objectif')
            ]},
            options: baseOpts('%', 88, 100)
        });
    }

    if (asm.length) {
        const labs   = asm.map(r => fmt(r.date));
        const rftAsm = asm.map(r => parseFloat(r.rft_assembly_pct)         || 0);
        const tauAsm = asm.map(r => parseFloat(r.taux_erreur_assembly_pct) || 0);

        mkChart('c-d-asm', {
            type: 'bar',
            data: { labels: labs, datasets: [
                barDsColored('RFT Assembly', rftAsm, 90, 'rft'),
                tgtLine(90, 'Objectif')
            ]},
            options: baseOpts('%', 0, 105)
        });

        mkChart('c-d-taux-asm', {
            type: 'bar',
            data: { labels: labs, datasets: [
                barDsColored('Taux Err Asm', tauAsm, 10, 'taux'),
                tgtLine(10, 'Objectif')
            ]},
            options: baseOpts('%', 0)
        });
    }
}

// =============================================================================
// EXTRA FILES TAB
// =============================================================================

function renderAdditionalTables() {
    const cont = gel('extra-content');
    if (!cont) return;

    const skip = new Set(['suivi_journalier','cnc_mensuel','cnc_journalier',
        'cnc_rft_hebdomadaire','cnc_pareto_defauts','cnc_pareto_actions',
        'cnc_par_operateur','cnc_par_statut','cnc_par_origine','cf_mensuel',
        'cf_journalier','cf_scrap_rework','cf_pareto_actions','cf_par_origine',
        'cf_par_technicien','mmc_journalier','mmc_finishgood_mensuel',
        'mmc_par_client','mmc_par_operateur','mmc_tests','kpi_mensuel',
        'rft_global','kpi_global','cnc_par_piece','rft_global_journalier','rft_ytd']);

    const extras = Object.keys(dashboardData).filter(k => !skip.has(k));
    if (!extras.length) {
        cont.innerHTML = `<div style="color:var(--dim);text-align:center;padding:40px;font-size:14px">
            Aucun fichier supplémentaire détecté.<br>
            <span style="font-size:12px;margin-top:8px;display:block">
                Placez vos fichiers Excel dans <code style="color:var(--acc)">Data/</code> — ils apparaîtront ici automatiquement.
            </span></div>`;
        return;
    }

    cont.innerHTML = '';
    extras.forEach(key => {
        const rows = dashboardData[key];
        if (!rows || !rows.length) return;
        const cols = Object.keys(rows[0]);
        const card = document.createElement('div');
        card.className = 'extra-card';
        card.innerHTML = `<h3>📄 ${key.replace(/_/g, ' ')} <small style="color:var(--dim);font-weight:400">${rows.length} lignes</small></h3>
            <div class="tscroll"><table class="tbl">
            <thead><tr>${cols.map(c => `<th>${c.replace(/_/g, ' ')}</th>`).join('')}</tr></thead>
            <tbody>${rows.slice(0, 100).map(r => `<tr>${cols.map(c => `<td>${r[c] ?? ''}</td>`).join('')}</tr>`).join('')}</tbody>
            </table></div>`;
        cont.appendChild(card);
    });
}

// =============================================================================
// UTILITIES
// =============================================================================

function avg(arr) {
    const v = arr.filter(x => x !== null && !isNaN(x));
    return v.length ? Math.round(v.reduce((a, b) => a + b, 0) / v.length * 10) / 10 : 0;
}

// pc: pill class — inv=true means lower is better (taux erreur)
function pc(v, t, inv = false) {
    return (inv ? v <= t : v >= t) ? 'good' : (Math.abs(v - t) < t * .15 ? 'warn' : 'bad');
}

// updateCard for RFT metrics (higher = better)
function updateCard(valId, barId, val, target, isRft = false) {
    const ve = gel(valId);
    if (ve) {
        ve.textContent = val + '%';
        ve.className = 'kv ' + (isRft ? pc(val, target) : 'neutral');
    }
    const be = gel(barId);
    if (be) {
        be.style.width = Math.min(Math.max(val, 0), 100) + '%';
        be.style.background = val >= target ? '#10b981' : '#ef4444';
    }
}

// updateTauxCard for error rate metrics (lower = better)
function updateTauxCard(valId, val, target) {
    const ve = gel(valId);
    if (ve) {
        ve.textContent = val + '%';
        ve.className = 'kv ' + pc(val, target, true);
    }
}

function filterByMonth(data) {
    if (!currentMonth || !data) return data;
    return data.filter(r => parseInt(r.mois_n) === currentMonth);
}

function filterByDate(data) {
    let d = data;
    if (currentYear)  d = d.filter(r => new Date(r.date).getFullYear()  === currentYear);
    if (currentMonth) d = d.filter(r => new Date(r.date).getMonth() + 1 === currentMonth);
    return d;
}

function fmt(dateStr) {
    if (!dateStr) return '';
    return new Date(dateStr).toLocaleDateString('fr-FR', { day: '2-digit', month: '2-digit' });
}

function updateRefreshInfo() {
    const el = gel('rt');
    if (el) el.textContent = new Date().toLocaleTimeString('fr-FR');
}

// =============================================================================
// AUTO-REFRESH — [5] Fixed: countdown timer + periodic reload
// =============================================================================

let _countdown = 300; // 5 minutes in seconds
let _countdownTimer = null;

function startAutoRefresh() {
    // Clear any existing timers first
    if (refreshInterval)  clearInterval(refreshInterval);
    if (_countdownTimer)  clearInterval(_countdownTimer);

    _countdown = 300;

    // Data reload every 5 minutes
    refreshInterval = setInterval(async () => {
        console.log('[Dashboard] Auto-refreshing data...');
        await loadDashboardData(currentYear);
        _countdown = 300;
    }, 300000);

    // Countdown display every second
    _countdownTimer = setInterval(() => {
        _countdown--;
        if (_countdown < 0) _countdown = 300;
        const rtEl = gel('rt');
        if (rtEl) {
            const m = Math.floor(_countdown / 60);
            const s = String(_countdown % 60).padStart(2, '0');
            rtEl.textContent = `${m}:${s}`;
        }
    }, 1000);

    console.log('[Dashboard] Auto-refresh started (5 min interval)');
}

window.addEventListener('beforeunload', () => {
    if (refreshInterval) clearInterval(refreshInterval);
    if (_countdownTimer) clearInterval(_countdownTimer);
});