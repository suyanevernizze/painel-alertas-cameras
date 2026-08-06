/**
 * parser.js
 * Responsável por ler o .xlsx exportado e transformar a estrutura
 * "intercalada" (alerta + bloco de histórico) em uma lista de objetos
 * planos, prontos para os cálculos e para a tabela de dados.
 */

function parseDate(str) {
  if (!str || str === '-') return null;
  const s = String(str).trim();
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/);
  if (!m) return null;
  const [, d, mo, y, h, mi, se] = m;
  return new Date(+y, +mo - 1, +d, +h, +mi, +se);
}

/**
 * Recebe a matriz bruta (array de arrays) lida do Excel via SheetJS
 * e devolve a lista de alertas já estruturada.
 *
 * Regra de negócio importante:
 * "Tratado" = o alerta teve QUALQUER interação no histórico
 * (falso positivo, assistido ou finalizado contam como tratativa).
 */
function parseAlerts(raw) {
  const rows = raw.slice(1); // remove cabeçalho da planilha
  const alerts = [];
  const isId = v => v !== null && v !== undefined && /^\d+$/.test(String(v).trim());

  let i = 0;
  const n = rows.length;

  while (i < n) {
    const row = rows[i] || [];
    if (isId(row[0])) {
      const alert = {
        id: String(row[0]).trim(),
        estado: row[1],
        placa: row[2],
        motorista: row[3],
        dataAlerta: parseDate(row[4]),
        risco: row[5],
        tipoEvento: row[6],
        endereco: row[7],
      };

      // coleta o bloco de histórico, se existir
      const events = [];
      let j = i + 1;
      if (j < n && rows[j] && String(rows[j][0] || '').trim() === 'Histórico de Eventos') {
        j++;
        if (j < n && rows[j] && String(rows[j][0] || '').trim() === 'Tipo de evento') j++;
        while (
          j < n && rows[j] && rows[j][0] !== null && rows[j][0] !== undefined &&
          String(rows[j][0]).trim() !== 'Histórico de Eventos' && !isId(rows[j][0])
        ) {
          events.push({ tipo: rows[j][0], usuario: rows[j][1], data: parseDate(rows[j][3]) });
          j++;
        }
      }

      const tiposTexto = events.map(e => String(e.tipo || '').toLowerCase());
      alert.nEventos = events.length;
      alert.falsoPositivo = tiposTexto.some(t => t.includes('falso positivo'));
      alert.assistido = tiposTexto.some(t => t.includes('assistido'));
      alert.finalizado = tiposTexto.some(t => t.includes('finalizado'));
      // CORRIGIDO: tratado = teve qualquer interação no histórico, não só assistido/finalizado.
      // Sem essa correção, alertas marcados só como "falso positivo" ficavam de fora da contagem.
      alert.tratado = events.length > 0;

      const usuarios = events.map(e => e.usuario).filter(u => u && u !== '-');
      alert.usuario = usuarios.length ? usuarios[0] : null;

      const datas = events.map(e => e.data).filter(d => d);
      alert.dataPrimeiraTratativa = datas.length ? new Date(Math.min(...datas)) : null;

      if (alert.dataAlerta && alert.dataPrimeiraTratativa) {
        const diffMin = (alert.dataPrimeiraTratativa - alert.dataAlerta) / 60000;
        alert.tempoResposta = diffMin >= 0 ? diffMin : null;
      } else {
        alert.tempoResposta = null;
      }

      alerts.push(alert);
      i = j;
    } else {
      i++;
    }
  }
  return alerts;
}

// exposto globalmente para os outros módulos (sem bundler neste projeto)
window.AppParser = { parseAlerts, parseDate };
/**
 * format.js — helpers de formatação usados nas duas telas.
 */
function fmtMin(min) {
  if (min === null || min === undefined || isNaN(min)) return '—';
  if (min < 60) return Math.round(min) + ' min';
  const h = Math.floor(min / 60), m = Math.round(min % 60);
  return h + 'h' + (m ? ' ' + m + 'min' : '');
}

function fmtPct(x) {
  return (x * 100).toFixed(1) + '%';
}

function fmtDateTime(d) {
  if (!d) return '—';
  return d.toLocaleDateString('pt-BR') + ' ' + d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
}

window.AppFormat = { fmtMin, fmtPct, fmtDateTime };
/**
 * stats.js
 * Agregações usadas no dashboard: KPIs gerais, por tipo de evento,
 * por risco, por veículo, por usuário e por dia.
 */

function average(arr) {
  return arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;
}

function computeKpis(data) {
  const total = data.length;
  const tratados = data.filter(a => a.tratado).length;
  const fpCount = data.filter(a => a.falsoPositivo).length;
  const pendentes = total - tratados;
  const tempos = data.filter(a => a.tempoResposta !== null).map(a => a.tempoResposta);
  return {
    total,
    tratados,
    pctTratados: total ? tratados / total : 0,
    pendentes,
    fpCount,
    taxaFp: total ? fpCount / total : 0,
    tempoMedio: average(tempos),
  };
}

function groupByTipoEvento(data) {
  const map = {};
  data.forEach(a => {
    const t = a.tipoEvento || 'Não informado';
    if (!map[t]) map[t] = { count: 0, fp: 0, proc: 0, tempos: [] };
    map[t].count++;
    if (a.falsoPositivo) map[t].fp++;
    if (a.tratado && !a.falsoPositivo) map[t].proc++;
    if (a.tempoResposta !== null) map[t].tempos.push(a.tempoResposta);
  });
  return Object.entries(map)
    .map(([nome, v]) => ({ nome, count: v.count, fp: v.fp, proc: v.proc, tempoMedio: average(v.tempos) }))
    .sort((a, b) => b.count - a.count);
}

function groupByRisco(data) {
  const map = {};
  data.forEach(a => {
    const r = a.risco || 'Não informado';
    if (!map[r]) map[r] = { count: 0, fp: 0, proc: 0 };
    map[r].count++;
    if (a.falsoPositivo) map[r].fp++;
    if (a.tratado && !a.falsoPositivo) map[r].proc++;
  });
  const ordem = ['Alto', 'Médio', 'Baixo'];
  const mk = nome => ({ nome, count: map[nome].count, fp: map[nome].fp, proc: map[nome].proc });
  const conhecidos = ordem.filter(r => map[r]).map(mk);
  const outros = Object.keys(map).filter(k => !ordem.includes(k)).map(mk);
  return conhecidos.concat(outros);
}

function topPlacas(data, limit = 10) {
  const map = {};
  data.forEach(a => {
    const p = a.placa || '—';
    if (!map[p]) map[p] = { count: 0, proc: 0 };
    map[p].count++;
    if (a.tratado && !a.falsoPositivo) map[p].proc++;
  });
  return Object.entries(map).map(([nome, v]) => ({ nome, count: v.count, proc: v.proc }))
    .sort((a, b) => b.count - a.count).slice(0, limit);
}

function groupByUsuario(data) {
  const map = {};
  data.forEach(a => {
    if (!a.usuario) return;
    if (!map[a.usuario]) map[a.usuario] = { count: 0, fp: 0, proc: 0, tempos: [], tipoMap: {}, diaMap: {} };
    map[a.usuario].count++;
    if (a.falsoPositivo) map[a.usuario].fp++;
    if (a.tratado && !a.falsoPositivo) map[a.usuario].proc++;
    if (a.tempoResposta !== null) map[a.usuario].tempos.push(a.tempoResposta);
    const t = a.tipoEvento || 'Outros';
    map[a.usuario].tipoMap[t] = (map[a.usuario].tipoMap[t] || 0) + 1;
    if (a.dataPrimeiraTratativa) {
      const dia = new Date(a.dataPrimeiraTratativa).toLocaleDateString('pt-BR');
      map[a.usuario].diaMap[dia] = (map[a.usuario].diaMap[dia] || 0) + 1;
    }
  });
  return Object.entries(map)
    .map(([nome, v]) => ({
      nome,
      count: v.count,
      fp: v.fp,
      proc: v.proc,
      taxaFp: v.count ? v.fp / v.count : 0,
      taxaProc: v.count ? v.proc / v.count : 0,
      tempoMedio: average(v.tempos),
      tempoMin: v.tempos.length ? Math.min(...v.tempos) : null,
      tempoMax: v.tempos.length ? Math.max(...v.tempos) : null,
      tipoMaisFrequente: Object.entries(v.tipoMap).sort((a, b) => b[1] - a[1])[0]?.[0] || '—',
      tipoMap: v.tipoMap,
      diaMap: v.diaMap,
    }))
    .sort((a, b) => b.count - a.count);
}

function groupByDia(data) {
  const map = {};
  data.forEach(a => {
    if (!a.dataAlerta) return;
    const key = a.dataAlerta.toLocaleDateString('pt-BR');
    if (!map[key]) map[key] = { count: 0, tratados: 0 };
    map[key].count++;
    if (a.tratado) map[key].tratados++;
  });
  return Object.entries(map).map(([data_, v]) => ({ data: data_, count: v.count, tratados: v.tratados }))
    .sort((a, b) => {
      const [da, ma, ya] = a.data.split('/');
      const [db, mb, yb] = b.data.split('/');
      return new Date(ya, ma - 1, da) - new Date(yb, mb - 1, db);
    });
}

window.AppStats = { computeKpis, groupByTipoEvento, groupByRisco, topPlacas, groupByUsuario, groupByDia, average };
/**
 * dashboard.js
 * Renderiza a aba "Indicadores": KPIs, tabelas agregadas e gráficos.
 */
let _charts = {};

function renderDashboard(data) {
  const { computeKpis, groupByTipoEvento, groupByRisco, topPlacas, groupByUsuario, groupByDia } = window.AppStats;
  const { fmtMin, fmtPct } = window.AppFormat;

  const kpis = computeKpis(data);
  document.getElementById('kpiTotal').textContent = kpis.total.toLocaleString('pt-BR');
  document.getElementById('kpiTratados').textContent = fmtPct(kpis.pctTratados);
  document.getElementById('kpiTempo').textContent = fmtMin(kpis.tempoMedio);
  document.getElementById('kpiFP').textContent = fmtPct(kpis.taxaFp);
  document.getElementById('kpiPend').textContent = kpis.pendentes.toLocaleString('pt-BR');
  const procedentesRows = data.filter(a => a.tratado && !a.falsoPositivo);
  document.getElementById('kpiProcedentes').textContent = procedentesRows.length.toLocaleString('pt-BR');
  window.__procedentesRows = procedentesRows;

  // --- tabela por tipo de evento ---
  const tipoArr = groupByTipoEvento(data);
  document.getElementById('tblTipo').innerHTML = tipoArr.map(t => {
    const pctTotal = t.count / kpis.total;
    return `<tr>
      <td>${t.nome}</td>
      <td class="num">${t.count}</td>
      <td class="num"><div class="bar-cell"><div class="bar-track"><div class="bar-fill" style="width:${pctTotal * 100}%;background:var(--blue)"></div></div>${fmtPct(pctTotal)}</div></td>
      <td class="num">${fmtMin(t.tempoMedio)}</td>
      <td class="num">${fmtPct(t.count ? t.fp / t.count : 0)}</td>
      <td class="num">${fmtPct(t.count ? t.proc / t.count : 0)}</td>
    </tr>`;
  }).join('');

  // --- tabela por risco ---
  const riscoArr = groupByRisco(data);
  document.getElementById('tblRisco').innerHTML = riscoArr.map(r => {
    const colorVar = r.nome === 'Alto' ? 'var(--red)' : r.nome === 'Médio' ? 'var(--amber)' : 'var(--teal)';
    const pctTotal = r.count / kpis.total;
    return `<tr>
      <td class="risk-${r.nome}">${r.nome}</td>
      <td class="num">${r.count}</td>
      <td class="num"><div class="bar-cell"><div class="bar-track"><div class="bar-fill" style="width:${pctTotal * 100}%;background:${colorVar}"></div></div>${fmtPct(pctTotal)}</div></td>
      <td class="num">${fmtPct(r.count ? r.fp / r.count : 0)}</td>
      <td class="num">${fmtPct(r.count ? r.proc / r.count : 0)}</td>
    </tr>`;
  }).join('');

  // --- top veículos ---
  const placaArr = topPlacas(data, 10);

  // --- tabela por usuário ---
  const userArr = groupByUsuario(data);
  document.getElementById('tblUsuario').innerHTML = userArr.map(u => `<tr>
      <td>${u.nome}</td>
      <td class="num">${u.count}</td>
      <td class="num">${fmtMin(u.tempoMedio)}</td>
      <td class="num">${fmtPct(u.count ? u.fp / u.count : 0)}</td>
      <td class="num">${fmtPct(u.count ? u.proc / u.count : 0)}</td>
    </tr>`).join('');

  // --- por dia ---
  const diaArr = groupByDia(data);

  drawCharts(tipoArr, placaArr, diaArr, groupByTipoEvento(procedentesRows));
}

function drawCharts(tipoArr, placaArr, diaArr, tipoVerdArr) {
  Object.values(_charts).forEach(c => c.destroy());
  const palette = ['#5B8DEF', '#36C2B4', '#F2A33C', '#E5484D', '#9B7BFF', '#3BC9DB', '#F783AC', '#94D82D', '#FFA94D'];
  const rootStyles = getComputedStyle(document.documentElement);
  const mutedColor = rootStyles.getPropertyValue('--muted').trim() || '#7C8AA5';
  const panelColor = rootStyles.getPropertyValue('--panel').trim() || '#111A2E';
  Chart.defaults.font.family = "'JetBrains Mono', monospace";
  Chart.defaults.color = mutedColor;

  const tipoDataLabels = {
    id: 'tipoDataLabels',
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      const dataset = chart.data.datasets[0];
      const total = dataset.data.reduce((a, b) => a + b, 0);
      if (!total) return;
      const meta = chart.getDatasetMeta(0);
      ctx.save();
      ctx.font = "bold 11px 'JetBrains Mono', monospace";
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      meta.data.forEach((arc, i) => {
        const value = dataset.data[i];
        const pct = value / total;
        if (pct < 0.02) return; // evita poluir fatias muito pequenas
        const { x, y, innerRadius, outerRadius, startAngle, endAngle } = arc.getProps(
          ['x', 'y', 'innerRadius', 'outerRadius', 'startAngle', 'endAngle'], true
        );
        const midAngle = (startAngle + endAngle) / 2;
        const radius = (innerRadius + outerRadius) / 2;
        const lx = x + Math.cos(midAngle) * radius;
        const ly = y + Math.sin(midAngle) * radius;
        const label = (pct * 100).toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + '%';
        ctx.lineWidth = 3;
        ctx.strokeStyle = 'rgba(0,0,0,0.55)';
        ctx.strokeText(label, lx, ly);
        ctx.fillStyle = '#fff';
        ctx.fillText(label, lx, ly);
      });
      ctx.restore();
    }
  };

  // cor fixa por tipo (mesma nas duas pizzas e na legenda compartilhada)
  const colorByTipo = {};
  tipoArr.forEach((t, i) => { colorByTipo[t.nome] = palette[i % palette.length]; });

  _charts.tipo = new Chart(document.getElementById('chartTipo'), {
    type: 'doughnut',
    data: { labels: tipoArr.map(t => t.nome), datasets: [{ data: tipoArr.map(t => t.count), backgroundColor: tipoArr.map(t => colorByTipo[t.nome]), borderColor: panelColor, borderWidth: 2 }] },
    options: { plugins: { legend: { display: false } }, maintainAspectRatio: false },
    plugins: [tipoDataLabels]
  });

  _charts.tipoVerd = new Chart(document.getElementById('chartTipoVerd'), {
    type: 'doughnut',
    data: { labels: tipoVerdArr.map(t => t.nome), datasets: [{ data: tipoVerdArr.map(t => t.count), backgroundColor: tipoVerdArr.map(t => colorByTipo[t.nome]), borderColor: panelColor, borderWidth: 2 }] },
    options: { plugins: { legend: { display: false } }, maintainAspectRatio: false },
    plugins: [tipoDataLabels]
  });

  // legenda única embaixo, compartilhada pelas duas pizzas
  const distLegendEl = document.getElementById('distLegend');
  if (distLegendEl) distLegendEl.innerHTML = tipoArr.map(t =>
    `<span class="leg-item"><span class="leg-dot" style="background:${colorByTipo[t.nome]}"></span>${t.nome}</span>`
  ).join('');

  _charts.placas = new Chart(document.getElementById('chartPlacas'), {
    type: 'bar',
    data: { labels: placaArr.map(p => p.nome), datasets: [
      { label: 'Total', data: placaArr.map(p => p.count), backgroundColor: '#F2A33C', borderRadius: 4 },
      { label: 'Positivos', data: placaArr.map(p => p.proc), backgroundColor: '#36C2B4', borderRadius: 4 }
    ] },
    options: {
      plugins: { legend: { display: true, position: 'top', labels: { boxWidth: 10, font: { size: 11 } } } },
      scales: { x: { grid: { display: false } }, y: { grid: { color: rootStyles.getPropertyValue('--line').trim() || 'rgba(255,255,255,0.05)' } } },
      maintainAspectRatio: false,
      animation: { onComplete: function() {
        var chart = this, ctx = chart.ctx;
        ctx.font = '9px JetBrains Mono, monospace';
        ctx.fillStyle = rootStyles.getPropertyValue('--muted').trim() || '#7C8AA5';
        ctx.textAlign = 'center';
        chart.data.datasets.forEach(function(ds, di) {
          chart.getDatasetMeta(di).data.forEach(function(bar, i) {
            ctx.fillText(ds.data[i].toLocaleString('pt-BR'), bar.x, bar.y - 5);
          });
        });
      }}
    }
  });

  const diaDataLabels = {
    id: 'diaDataLabels',
    afterDatasetsDraw(chart) {
      const { ctx } = chart;
      ctx.save();
      ctx.font = "10px 'JetBrains Mono', monospace";
      ctx.textAlign = 'center';
      chart.data.datasets.forEach((ds, di) => {
        const meta = chart.getDatasetMeta(di);
        if (meta.hidden) return;
        // Alertas (índice 0) fica com o rótulo acima do ponto; Tratados (índice 1) fica abaixo,
        // para não sobrepor quando as duas linhas se cruzam.
        const dy = di === 0 ? -8 : 14;
        ctx.fillStyle = ds.borderColor;
        meta.data.forEach((point, i) => {
          const val = ds.data[i];
          if (val === null || val === undefined) return;
          ctx.fillText(val.toLocaleString('pt-BR'), point.x, point.y + dy);
        });
      });
      ctx.restore();
    }
  };

  _charts.dia = new Chart(document.getElementById('chartDia'), {
    type: 'line',
    data: {
      labels: diaArr.map(d => d.data),
      datasets: [
        { label: 'Alertas', data: diaArr.map(d => d.count), borderColor: '#5B8DEF', backgroundColor: 'rgba(91,141,239,0.12)', fill: true, tension: 0.3, pointRadius: 3 },
        { label: 'Tratados', data: diaArr.map(d => d.tratados), borderColor: '#36C2B4', backgroundColor: 'rgba(54,194,180,0.12)', fill: true, tension: 0.3, pointRadius: 3 },
      ]
    },
    options: {
      interaction: { mode: 'index', intersect: false },
      layout: { padding: { top: 16, bottom: 10 } },
      plugins: {
        legend: { display: true, position: 'top', labels: { boxWidth: 10, font: { size: 11 } } },
        tooltip: { mode: 'index', intersect: false }
      },
      scales: { x: { grid: { display: false } }, y: { grid: { color: rootStyles.getPropertyValue('--line').trim() || 'rgba(255,255,255,0.05)' } } },
      maintainAspectRatio: false
    },
    plugins: [diaDataLabels]
  });
}

window.AppDashboard = { renderDashboard };
/**
 * dadosTable.js
 * Renderiza a aba "Dados": a base de alertas linha a linha,
 * equivalente à aba "Dados" criada na planilha Excel — com busca
 * e paginação porque aqui são milhares de linhas.
 */
const PAGE_SIZE = 100;
/* ===== Modal de alertas procedentes (abre ao clicar no card) ===== */
const PROC_MODAL_LIMIT = 1000;
function renderProcedentesModal() {
  const rows = window.__procedentesRows || [];
  const fmt = window.AppFormat;
  const shown = rows.slice(0, PROC_MODAL_LIMIT);
  document.getElementById('procedentesModalCount').textContent = rows.length > PROC_MODAL_LIMIT
    ? `(exibindo ${PROC_MODAL_LIMIT.toLocaleString('pt-BR')} de ${rows.length.toLocaleString('pt-BR')} — refine o período)`
    : `(${rows.length.toLocaleString('pt-BR')})`;
  document.getElementById('procedentesModalBody').innerHTML = shown.map(a =>
    `<tr><td>${a.id}</td><td class="col-placa">${a.placa||'—'}</td><td>${fmt.fmtDateTime(a.dataAlerta)}</td>` +
    `<td class="risk-${a.risco}">${a.risco||'—'}</td><td>${a.tipoEvento||'—'}</td><td>${a.endereco||'—'}</td>` +
    `<td>${a.usuario||'—'}</td><td class="num">${fmt.fmtMin(a.tempoResposta)}</td></tr>`
  ).join('') || '<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:20px;">nenhum alerta procedente no período</td></tr>';
}
function openProcedentesModal() { renderProcedentesModal(); const m = document.getElementById('procedentesModal'); if (m) m.hidden = false; }
function closeProcedentesModal() { const m = document.getElementById('procedentesModal'); if (m) m.hidden = true; }
(function () {
  function setup() {
    const card = document.getElementById('cardProcedentes');
    if (card) {
      card.addEventListener('click', openProcedentesModal);
      card.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openProcedentesModal(); } });
    }
    const closeBtn = document.getElementById('procedentesModalClose');
    if (closeBtn) closeBtn.addEventListener('click', closeProcedentesModal);
    const overlay = document.getElementById('procedentesModal');
    if (overlay) overlay.addEventListener('click', e => { if (e.target === overlay) closeProcedentesModal(); });
    document.addEventListener('keydown', e => { if (e.key === 'Escape') closeProcedentesModal(); });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', setup);
  else setup();
})();

let _dadosState = { rows: [], filtered: [], page: 1 };

function badge(value, kind) {
  if (kind === 'tratado') {
    return value ? '<span class="badge yes">Tratado</span>' : '<span class="badge no">Pendente</span>';
  }
  if (kind === 'fp') {
    return value ? '<span class="badge fp">Sim</span>' : '<span class="badge no">Não</span>';
  }
  return value;
}

function buildDadosToolbar() {
  const wrap = document.getElementById('dadosToolbar');
  wrap.innerHTML = `
    <input type="text" id="dadosSearch" placeholder="buscar por placa, motorista, endereço, usuário…">
    <select id="dadosFiltroTratado">
      <option value="">todos os status</option>
      <option value="tratado">tratados</option>
      <option value="pendente">pendentes</option>
    </select>
    <span class="dados-count" id="dadosCount"></span>
  `;
  document.getElementById('dadosSearch').addEventListener('input', () => { _dadosState.page = 1; applyFilters(); });
  document.getElementById('dadosFiltroTratado').addEventListener('change', () => { _dadosState.page = 1; applyFilters(); });
}

function applyFilters() {
  const term = (document.getElementById('dadosSearch').value || '').toLowerCase().trim();
  const statusFiltro = document.getElementById('dadosFiltroTratado').value;
  _dadosState.filtered = _dadosState.rows.filter(a => {
    if (statusFiltro === 'tratado' && !a.tratado) return false;
    if (statusFiltro === 'pendente' && a.tratado) return false;
    if (!term) return true;
    const haystack = [a.placa, a.motorista, a.endereco, a.usuario, a.tipoEvento, a.id]
      .map(v => String(v || '').toLowerCase()).join(' ');
    return haystack.includes(term);
  });
  renderPage();
}

function renderPage() {
  const { fmtMin, fmtDateTime } = window.AppFormat;
  const total = _dadosState.filtered.length;
  const totalPages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  if (_dadosState.page > totalPages) _dadosState.page = totalPages;
  const start = (_dadosState.page - 1) * PAGE_SIZE;
  const pageRows = _dadosState.filtered.slice(start, start + PAGE_SIZE);

  document.getElementById('dadosCount').textContent = `${total.toLocaleString('pt-BR')} alertas encontrados`;

  const tbody = document.getElementById('tblDados');
  tbody.innerHTML = pageRows.map(a => `
    <tr>
      <td>${a.id}</td>
      <td class="col-placa">${a.placa || '—'}</td>
      <td>${fmtDateTime(a.dataAlerta)}</td>
      <td class="risk-${a.risco}">${a.risco || '—'}</td>
      <td>${a.tipoEvento || '—'}</td>
      <td>${a.endereco || '—'}</td>
      <td>${badge(a.tratado, 'tratado')}</td>
      <td>${badge(a.falsoPositivo, 'fp')}</td>
      <td>${a.usuario || '—'}</td>
      <td class="num">${fmtMin(a.tempoResposta)}</td>
    </tr>
  `).join('');

  document.getElementById('pageInfo').textContent = `página ${_dadosState.page} de ${totalPages}`;
  document.getElementById('prevPage').disabled = _dadosState.page <= 1;
  document.getElementById('nextPage').disabled = _dadosState.page >= totalPages;
}

function renderDadosTable(data) {
  _dadosState.rows = data;
  _dadosState.page = 1;
  buildDadosToolbar();
  applyFilters();

  document.getElementById('prevPage').onclick = () => { if (_dadosState.page > 1) { _dadosState.page--; renderPage(); } };
  document.getElementById('nextPage').onclick = () => {
    const totalPages = Math.max(1, Math.ceil(_dadosState.filtered.length / PAGE_SIZE));
    if (_dadosState.page < totalPages) { _dadosState.page++; renderPage(); }
  };
}

window.AppDadosTable = { renderDadosTable };
/**
 * app.js — ponto de entrada: drag-and-drop, leitura do .xlsx,
 * controle das abas (Indicadores / Dados) e estado da aplicação.
 */
(function () {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const statusEl = document.getElementById('status');
  const filehint = document.getElementById('filehint');
  const tabs = document.getElementById('tabs');

  let currentData = [];

  // Verificação defensiva: se alguma dependência não carregou (CDN bloqueado,
  // bloqueador de anúncios, offline, ou um dos módulos não foi publicado
  // junto), mostra uma mensagem clara em vez de a página ficar "sem reagir".
  function checkDependencies() {
    const missing = [];
    if (window.__libError === 'xlsx' || typeof XLSX === 'undefined') missing.push('biblioteca de leitura de Excel (xlsx)');
    if (window.__libError === 'chart' || typeof Chart === 'undefined') missing.push('biblioteca de gráficos (Chart.js)');
    if (!window.AppParser) missing.push('módulo parser');
    if (!window.AppStats) missing.push('módulo stats');
    if (!window.AppFormat) missing.push('módulo format');
    if (!window.AppDashboard) missing.push('módulo dashboard');
    if (!window.AppDadosTable) missing.push('módulo dadosTable');
    if (missing.length) {
      setStatus('não foi possível carregar: ' + missing.join(', ') + '. Verifique sua conexão ou se o app.js está completo, e recarregue a página (Ctrl+F5).', 'err');
      dropzone.style.opacity = '0.4';
      dropzone.style.pointerEvents = 'none';
      return false;
    }
    return true;
  }

  ['dragenter', 'dragover'].forEach(ev => dropzone.addEventListener(ev, e => {
    e.preventDefault(); dropzone.classList.add('drag');
  }));
  ['dragleave', 'drop'].forEach(ev => dropzone.addEventListener(ev, e => {
    e.preventDefault(); dropzone.classList.remove('drag');
  }));
  dropzone.addEventListener('drop', e => {
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  });
  dropzone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', e => {
    if (e.target.files[0]) handleFile(e.target.files[0]);
  });
  document.getElementById('resetBtn').addEventListener('click', resetApp);

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });

  checkDependencies();

  function switchTab(name) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === name));
    document.querySelectorAll('.view').forEach(v => v.classList.toggle('active', v.id === 'view-' + name));
    document.querySelector('main').classList.add('wide');
  }

  function setStatus(msg, type) {
    statusEl.textContent = msg;
    statusEl.className = type || '';
  }

  function resetApp() {
    currentData = [];
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    tabs.classList.remove('show');
    dropzone.style.display = 'block';
    setStatus('');
    filehint.textContent = 'nenhum arquivo carregado';
    fileInput.value = '';
  }

  function handleFile(file) {
    if (!checkDependencies()) return;
    setStatus('lendo arquivo…');
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: null });
        const parsed = window.AppParser.parseAlerts(raw);

        if (parsed.length === 0) {
          setStatus('não encontrei alertas reconhecíveis nessa planilha. Confira se é o export de "Relatório de tratativas".', 'err');
          return;
        }

        currentData = parsed;
        window.__currentAlertData = parsed;
        populateFiltroPlaca(parsed);
        populateDateBounds(parsed);
        resetFiltrosInputs();
        window.AppDashboard.renderDashboard(parsed);
        window.AppDadosTable.renderDadosTable(parsed);
        if (window.AppUsuarios) window.AppUsuarios.renderUsuarios(parsed);

        dropzone.style.display = 'none';
        tabs.classList.add('show');
        switchTab('indicadores');

        filehint.textContent = file.name + ' · ' + parsed.length + ' alertas';
        setStatus('');
      } catch (err) {
        console.error(err);
        setStatus('erro ao processar o arquivo: ' + err.message, 'err');
      }
    };
    reader.readAsArrayBuffer(file);
  }
})();

/**
 * Fundo animado: bolhas subindo com mini ícones de câmera de caminhão,
 * alerta e a marca "B" da Binotto, substituindo a antiga linha de scan.
 */
(function () {
  const container = document.getElementById('bubbleBg');
  if (!container) return;
  // tipos de bolha: emoji (câmera de caminhão / alerta) ou a marca "B"
  const types = [
    { kind: 'emoji', value: '📹' },
    { kind: 'emoji', value: '🚛' },
    { kind: 'emoji', value: '⚠️' },
    { kind: 'emoji', value: '😴' },
    { kind: 'emoji', value: '🚨' },
    { kind: 'emoji', value: '📊' },
    { kind: 'emoji', value: '🧐' },
    { kind: 'emoji', value: '🥱' },
    { kind: 'emoji', value: '📵' },
    { kind: 'brand', value: 'B' },
    { kind: 'emoji', value: '⚠️' },
  ];
  const MAX_BUBBLES = 18;

  function spawnBubble() {
    const b = document.createElement('div');
    b.className = 'bubble';
    const size = 26 + Math.random() * 46; // 26–72px
    const left = Math.random() * 100; // %
    const duration = 14 + Math.random() * 14; // 14–28s
    const delay = -Math.random() * duration; // entra já em movimento
    const drift = (Math.random() * 60 - 30) + 'px';
    b.style.width = size + 'px';
    b.style.height = size + 'px';
    b.style.left = left + '%';
    b.style.setProperty('--drift', drift);
    b.style.animationDuration = duration + 's';
    b.style.animationDelay = delay + 's';

    const type = types[Math.floor(Math.random() * types.length)];
    const icon = document.createElement('span');
    if (type.kind === 'brand') {
      icon.className = 'icon icon-brand';
      icon.style.fontSize = (size * 0.4) + 'px';
    } else {
      icon.className = 'icon';
      icon.style.fontSize = (size * 0.45) + 'px';
    }
    icon.textContent = type.value;
    b.appendChild(icon);

    b.addEventListener('click', () => popBubble(b));

    container.appendChild(b);
  }

  function popBubble(b) {
    if (b.classList.contains('pop')) return;

    // captura a posição/tamanho REAIS da bolha antes de qualquer alteração,
    // pra garantir que a explosão nasça exatamente onde ela estava.
    const rect = b.getBoundingClientRect();
    const parentRect = container.getBoundingClientRect();
    const cx = rect.left - parentRect.left + rect.width / 2;
    const cy = rect.top - parentRect.top + rect.height / 2;
    const size = rect.width;

    b.classList.add('pop');
    // trava a bolha na posição capturada (em vez de deixar a animação de
    // subida continuar/saltar) pra ela encolher exatamente onde foi clicada
    b.style.left = (cx - size / 2) + 'px';
    b.style.bottom = 'auto';
    b.style.top = (cy - size / 2) + 'px';

    // flash central (clarão da explosão)
    const flash = document.createElement('div');
    flash.className = 'pop-flash';
    flash.style.width = size + 'px';
    flash.style.height = size + 'px';
    flash.style.left = (cx - size / 2) + 'px';
    flash.style.top = (cy - size / 2) + 'px';
    container.appendChild(flash);
    setTimeout(() => flash.remove(), 420);

    // anel de onda
    const ripple = document.createElement('div');
    ripple.className = 'pop-ripple';
    ripple.style.width = size + 'px';
    ripple.style.height = size + 'px';
    ripple.style.left = (cx - size / 2) + 'px';
    ripple.style.top = (cy - size / 2) + 'px';
    container.appendChild(ripple);
    setTimeout(() => ripple.remove(), 570);

    // fragmentos estilhaçando pra fora
    const fragCount = 10 + Math.floor(Math.random() * 4);
    for (let i = 0; i < fragCount; i++) {
      const frag = document.createElement('div');
      frag.className = 'pop-fragment';
      const fragSize = 5 + Math.random() * 7;
      const angle = (Math.PI * 2 * i) / fragCount + Math.random() * 0.5;
      const dist = size * 0.9 + Math.random() * size * 0.7;
      frag.style.width = fragSize + 'px';
      frag.style.height = fragSize + 'px';
      frag.style.left = (cx - fragSize / 2) + 'px';
      frag.style.top = (cy - fragSize / 2) + 'px';
      frag.style.setProperty('--fx', Math.cos(angle) * dist + 'px');
      frag.style.setProperty('--fy', Math.sin(angle) * dist + 'px');
      container.appendChild(frag);
      setTimeout(() => frag.remove(), 570);
    }

    setTimeout(() => {
      b.remove();
      spawnBubble(); // mantém o número de bolhas no ar
    }, 250);
  }

  for (let i = 0; i < MAX_BUBBLES; i++) spawnBubble();
})();



/**
 * Alternância de tema claro/escuro, com preferência salva no navegador.
 */
(function () {
  const toggleBtn = document.getElementById('themeToggle');
  if (!toggleBtn) return;
  const root = document.documentElement;

  function applyTheme(theme) {
    const moonIcon = document.getElementById('themeIconMoon');
    const sunIcon = document.getElementById('themeIconSun');
    if (theme === 'light') {
      root.setAttribute('data-theme', 'light');
      if (moonIcon) moonIcon.style.display = 'none';
      if (sunIcon) sunIcon.style.display = 'block';
    } else {
      root.removeAttribute('data-theme');
      if (moonIcon) moonIcon.style.display = 'block';
      if (sunIcon) sunIcon.style.display = 'none';
    }
  }

  let saved = null;
  try { saved = localStorage.getItem('painel-tema'); } catch (e) { /* sem acesso a storage */ }
  const preferLight = window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches;
  applyTheme(saved || (preferLight ? 'light' : 'dark'));

  toggleBtn.addEventListener('click', () => {
    const isLight = root.getAttribute('data-theme') === 'light';
    const next = isLight ? 'dark' : 'light';
    applyTheme(next);
    try { localStorage.setItem('painel-tema', next); } catch (e) { /* ignora se bloqueado */ }
    // re-renderiza os gráficos pra pegarem as cores do novo tema (mantendo os filtros ativos)
    if (window.AppFiltrosIndicadores) {
      window.AppFiltrosIndicadores.applyFilters();
      if (window.AppUsuarios && window.__currentAlertData) window.AppUsuarios.renderUsuarios(window.__currentAlertData);
    } else if (window.__currentAlertData && window.AppDashboard) {
      window.AppDashboard.renderDashboard(window.__currentAlertData);
    }
  });
})();

/**
 * Filtros do painel de Indicadores: período (data início/fim) e placa.
 * Filtra sobre os dados completos (window.__currentAlertData) e
 * re-renderiza indicadores e a tabela de dados.
 */
(function () {
  const dataIni = document.getElementById('filtroDataIni');
  const dataFim = document.getElementById('filtroDataFim');
  const selectPlaca = document.getElementById('filtroPlaca');
  const btnLimpar = document.getElementById('filtroLimpar');
  if (!dataIni || !dataFim || !selectPlaca || !btnLimpar) return;

  function parseInputDate(value, endOfDay) {
    if (!value) return null;
    const [y, m, d] = value.split('-').map(Number);
    return endOfDay ? new Date(y, m - 1, d, 23, 59, 59, 999) : new Date(y, m - 1, d, 0, 0, 0, 0);
  }

  function applyFilters() {
    const all = window.__currentAlertData || [];
    const ini = parseInputDate(dataIni.value, false);
    const fim = parseInputDate(dataFim.value, true);
    const placa = (selectPlaca.value || '').trim().toUpperCase();

    const filtered = all.filter(a => {
      if (ini && (!a.dataAlerta || a.dataAlerta < ini)) return false;
      if (fim && (!a.dataAlerta || a.dataAlerta > fim)) return false;
      if (placa && !String(a.placa || '').toUpperCase().includes(placa)) return false;
      return true;
    });

    if (window.AppDashboard) window.AppDashboard.renderDashboard(filtered);
    if (window.AppDadosTable) window.AppDadosTable.renderDadosTable(filtered);
    if (window.AppUsuarios) window.AppUsuarios.renderUsuarios(filtered);
  }

  dataIni.addEventListener('change', applyFilters);
  dataFim.addEventListener('change', applyFilters);
  selectPlaca.addEventListener('input', applyFilters);
  btnLimpar.addEventListener('click', () => {
    dataIni.value = '';
    dataFim.value = '';
    selectPlaca.value = '';
    applyFilters();
  });


  window.AppFiltrosIndicadores = { applyFilters };
})();

/**
 * Preenche as sugestões (datalist) do campo digitável de placa
 * com as placas únicas da planilha carregada.
 */
/**
 * AppRangePicker — seletor de INTERVALO com dois meses lado a lado e atalhos
 * laterais. A seleção (2 cliques) ou um atalho aplica e fecha na hora, sem
 * botões. Clicar fora fecha sem aplicar. Substitui os dois campos De/Até por um
 * único campo, mas por baixo continua alimentando os inputs originais
 * (mesmos ids, agora hidden) e disparando "change", então toda a lógica de
 * filtro e o "limpar filtros" seguem valendo.
 *
 * Navegação travada no intervalo de dados: o painel esquerdo vai de min até
 * (max-1) e o direito é sempre esquerdo+1, cobrindo todos os meses com dados
 * sem exibir meses vazios. Atalhos que caem fora dos dados são recortados; se
 * ficarem totalmente fora, aparecem desabilitados.
 */
(function () {
  // ---------- lógica pura (idêntica à testada em rangelogic.js) ----------
  function ymd(d){ return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
  function fromYmd(s){ if(!s) return null; var p=String(s).split('-'); if(p.length!==3) return null; var d=new Date(+p[0],+p[1]-1,+p[2]); d.setHours(0,0,0,0); return isNaN(d)?null:d; }
  function displayBr(s){ var d=fromYmd(s); if(!d) return ''; return String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0')+'/'+d.getFullYear(); }
  function monthIndex(y,m){ return y*12+m; }
  function idxToYM(idx){ return { year:Math.floor(idx/12), month:((idx%12)+12)%12 }; }
  function startOfWeek(d){ var r=new Date(d); var diff=(r.getDay()+6)%7; r.setDate(r.getDate()-diff); r.setHours(0,0,0,0); return r; }

  function buildMonth(year,month,minStr,maxStr,startStr,endStr,hoverStr){
    var startDow=new Date(year,month,1).getDay();
    var dim=new Date(year,month+1,0).getDate();
    var min=fromYmd(minStr),max=fromYmd(maxStr),start=fromYmd(startStr),end=fromYmd(endStr),hover=fromYmd(hoverStr);
    var rA=start, rB=end||(start&&hover?hover:null);
    if(rA&&rB&&rB<rA){ var t=rA; rA=rB; rB=t; }
    var cells=[];
    for(var i=0;i<42;i++){
      var dn=i-startDow+1;
      if(dn<1||dn>dim){ cells.push(null); continue; }
      var d=new Date(year,month,dn); d.setHours(0,0,0,0);
      cells.push({
        day:dn, ymd:ymd(d),
        disabled:(min&&d<min)||(max&&d>max),
        isStart:!!(start&&d.getTime()===start.getTime()),
        isEnd:!!(end&&d.getTime()===end.getTime()),
        inRange:!!(rA&&rB&&d>=rA&&d<=rB)
      });
    }
    return cells;
  }
  function leftRange(minStr,maxStr){
    var min=fromYmd(minStr),max=fromYmd(maxStr); if(!min||!max) return null;
    var lo=monthIndex(min.getFullYear(),min.getMonth()), hi=monthIndex(max.getFullYear(),max.getMonth());
    return { lo:lo, hi: hi>lo?hi-1:hi };
  }
  function navState(ly,lm,minStr,maxStr){
    var lr=leftRange(minStr,maxStr), cur=monthIndex(ly,lm);
    if(!lr) return { canPrev:true, canNext:true };
    return { canPrev:cur>lr.lo, canNext:cur<lr.hi };
  }
  function initialLeft(startStr,endStr,minStr,maxStr){
    var base=fromYmd(startStr)||fromYmd(endStr)||fromYmd(maxStr)||fromYmd(minStr)||new Date();
    var idx=monthIndex(base.getFullYear(),base.getMonth()), lr=leftRange(minStr,maxStr);
    if(lr){ if(idx<lr.lo) idx=lr.lo; if(idx>lr.hi) idx=lr.hi; }
    return idxToYM(idx);
  }
  function clampRange(iniStr,fimStr,minStr,maxStr){
    var ini=fromYmd(iniStr),fim=fromYmd(fimStr),min=fromYmd(minStr),max=fromYmd(maxStr);
    if(!ini||!fim) return null;
    if(fim<ini){ var t=ini; ini=fim; fim=t; }
    if(min&&ini<min) ini=min; if(max&&fim>max) fim=max;
    if(ini>fim) return null;
    return { ini:ymd(ini), fim:ymd(fim) };
  }
  function quickRange(key,now){
    now=now||new Date(); var y=now.getFullYear(),mo=now.getMonth(),ini,fim;
    switch(key){
      case 'hoje': ini=new Date(y,mo,now.getDate()); fim=new Date(ini); break;
      case 'ontem': ini=new Date(y,mo,now.getDate()-1); fim=new Date(ini); break;
      case 'semana-atual': ini=startOfWeek(now); fim=new Date(ini); fim.setDate(fim.getDate()+6); break;
      case 'semana-anterior': fim=new Date(startOfWeek(now)); fim.setDate(fim.getDate()-1); ini=new Date(fim); ini.setDate(ini.getDate()-6); break;
      case 'mes-atual': ini=new Date(y,mo,1); fim=new Date(y,mo+1,0); break;
      case 'mes-anterior': ini=new Date(y,mo-1,1); fim=new Date(y,mo,0); break;
      case 'trimestre-atual': { var q=Math.floor(mo/3); ini=new Date(y,q*3,1); fim=new Date(y,q*3+3,0); break; }
      case 'trimestre-anterior': { var q2=Math.floor(mo/3)-1,yy=y; if(q2<0){q2=3;yy=y-1;} ini=new Date(yy,q2*3,1); fim=new Date(yy,q2*3+3,0); break; }
      case 'ano-atual': ini=new Date(y,0,1); fim=new Date(y,11,31); break;
      case 'ano-anterior': ini=new Date(y-1,0,1); fim=new Date(y-1,11,31); break;
      default: return null;
    }
    return { ini:ymd(ini), fim:ymd(fim) };
  }

  // ---------- constantes de UI ----------
  var MES_ABBR=['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var DOW=['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];
  var SHORTCUTS=[
    { key:'', label:'Limpar' },
    { key:'hoje', label:'Hoje' },
    { key:'ontem', label:'Ontem' },
    { key:'semana-atual', label:'Esta semana' },
    { key:'semana-anterior', label:'Semana passada' },
    { key:'mes-atual', label:'Este mês' },
    { key:'mes-anterior', label:'Mês passado' },
    { key:'trimestre-atual', label:'Este trimestre' },
    { key:'trimestre-anterior', label:'Trimestre passado' },
    { key:'ano-atual', label:'Este ano' },
    { key:'ano-anterior', label:'Ano passado' }
  ];

  var _bounds={};      // iniId -> {min,max}
  var _openClose=null; // fecha o popup aberto atualmente

  function setBounds(iniId,minStr,maxStr){ _bounds[iniId]={min:minStr,max:maxStr}; }

  function attach(iniInput, fimInput){
    if(!iniInput||!fimInput||iniInput.__rpAttached) return;
    iniInput.__rpAttached=true;

    iniInput.type='hidden'; fimInput.type='hidden';
    var iniField=iniInput.parentNode;           // .filter-field do "De"
    var fimField=fimInput.parentNode;           // .filter-field do "Até"
    var lbl=iniField.querySelector('label'); if(lbl) lbl.textContent='Período';
    if(fimField && fimField!==iniField) fimField.style.display='none';

    var wrap=document.createElement('div'); wrap.className='rp-wrap';
    var field=document.createElement('button'); field.type='button'; field.className='rp-field';
    field.innerHTML='<span class="rp-text"></span><span class="rp-ico">📅</span>';
    var popup=document.createElement('div'); popup.className='rp-popup'; popup.style.display='none';
    iniField.appendChild(wrap); wrap.appendChild(field); wrap.appendChild(popup);

    // sincroniza o texto do campo quando .value muda por código (dropdown/limpar)
    function hookValue(input){
      var proto=Object.getPrototypeOf(input);
      var desc=Object.getOwnPropertyDescriptor(proto,'value');
      Object.defineProperty(input,'value',{
        configurable:true,
        get:function(){ return desc.get.call(this); },
        set:function(v){ desc.set.call(this,v); syncField(); }
      });
      input.__rpRaw=desc; // acesso ao value cru sem re-disparar sync
    }
    function raw(input){ return input.__rpRaw.get.call(input); }
    hookValue(iniInput); hookValue(fimInput);

    var view={ ly:0, lm:0 };
    var pStart=null, pEnd=null, hover=null;

    function bounds(){ return _bounds[iniInput.id]||{}; }

    function syncField(){
      var t=field.querySelector('.rp-text');
      var a=raw(iniInput), b=raw(fimInput);
      if(a&&b){ t.textContent=displayBr(a)+' — '+displayBr(b); field.classList.add('has-val'); }
      else if(a){ t.textContent=displayBr(a); field.classList.add('has-val'); }
      else { t.textContent='selecione o período'; field.classList.remove('has-val'); }
    }

    function yearOptions(rLo,rHi,curYear){
      var y0=idxToYM(rLo).year, y1=idxToYM(rHi).year, html='';
      for(var y=y0;y<=y1;y++) html+='<option value="'+y+'"'+(y===curYear?' selected':'')+'>'+y+'</option>';
      return html;
    }
    function mesOptions(rLo,rHi,year,curMonth){
      var html='';
      for(var m=0;m<12;m++){
        var idx=monthIndex(year,m), dis=(idx<rLo||idx>rHi);
        html+='<option value="'+m+'"'+(m===curMonth?' selected':'')+(dis?' disabled':'')+'>'+MES_ABBR[m]+'</option>';
      }
      return html;
    }

    function gridHtml(year,month){
      var b=bounds();
      var cells=buildMonth(year,month,b.min,b.max,pStart,pEnd,hover);
      return cells.map(function(c){
        if(!c) return '<span class="rp-day empty"></span>';
        var cls='rp-day';
        if(c.disabled) cls+=' disabled';
        if(c.inRange) cls+=' in-range';
        if(c.isStart) cls+=' start';
        if(c.isEnd) cls+=' end';
        return '<button type="button" class="'+cls+'" data-ymd="'+c.ymd+'"'+(c.disabled?' disabled':'')+'>'+c.day+'</button>';
      }).join('');
    }

    function renderAll(){
      var b=bounds();
      var lr=leftRange(b.min,b.max)||{lo:monthIndex(view.ly,view.lm),hi:monthIndex(view.ly,view.lm)};
      var leftIdx=monthIndex(view.ly,view.lm);
      var rightIdx=leftIdx+1;
      var L=idxToYM(leftIdx), R=idxToYM(rightIdx);
      var ns=navState(view.ly,view.lm,b.min,b.max);

      var side='<div class="rp-side">'+SHORTCUTS.map(function(s){
        var dis=false;
        if(s.key){ var qr=quickRange(s.key,new Date()); dis=!clampRange(qr.ini,qr.fim,b.min,b.max); }
        return '<button type="button" class="rp-shortcut'+(s.key===''?' rp-clear':'')+'" data-key="'+s.key+'"'+(dis?' disabled':'')+'>'+s.label+'</button>';
      }).join('')+'</div>';

      var calL='<div class="rp-cal">'+
        '<div class="rp-cal-head">'+
          '<button type="button" class="rp-nav rp-prev"'+(ns.canPrev?'':' disabled')+'>‹</button>'+
          '<span class="rp-selects"><select class="rp-mes" data-side="L">'+mesOptions(lr.lo,lr.hi,L.year,L.month)+'</select>'+
          '<select class="rp-ano" data-side="L">'+yearOptions(lr.lo,lr.hi,L.year)+'</select></span>'+
          '<span class="rp-nav-spacer"></span>'+
        '</div>'+
        '<div class="rp-dow">'+DOW.map(function(d){return '<span>'+d+'</span>';}).join('')+'</div>'+
        '<div class="rp-grid" data-side="L">'+gridHtml(L.year,L.month)+'</div>'+
      '</div>';

      var calR='<div class="rp-cal">'+
        '<div class="rp-cal-head">'+
          '<span class="rp-nav-spacer"></span>'+
          '<span class="rp-selects"><select class="rp-mes" data-side="R">'+mesOptions(lr.lo+1,lr.hi+1,R.year,R.month)+'</select>'+
          '<select class="rp-ano" data-side="R">'+yearOptions(lr.lo+1,lr.hi+1,R.year)+'</select></span>'+
          '<button type="button" class="rp-nav rp-next"'+(ns.canNext?'':' disabled')+'>›</button>'+
        '</div>'+
        '<div class="rp-dow">'+DOW.map(function(d){return '<span>'+d+'</span>';}).join('')+'</div>'+
        '<div class="rp-grid" data-side="R">'+gridHtml(R.year,R.month)+'</div>'+
      '</div>';

      popup.innerHTML=side+'<div class="rp-main"><div class="rp-cals">'+calL+calR+'</div></div>';
      wire();
    }

    // Repinta só as classes dos dias existentes (NÃO recria os elementos),
    // pra o clique da data-fim nunca ser perdido por recriação no hover.
    function paintRange(){
      var A=fromYmd(pStart), B=fromYmd(pEnd || (pStart && hover ? hover : null));
      if(A&&B&&B<A){ var t=A; A=B; B=t; }
      popup.querySelectorAll('.rp-day[data-ymd]').forEach(function(btn){
        var d=fromYmd(btn.getAttribute('data-ymd'));
        btn.classList.remove('in-range','start','end');
        if(!d) return;
        if(A&&B&&d>=A&&d<=B) btn.classList.add('in-range');
        if(A&&d.getTime()===A.getTime()) btn.classList.add('start');
        if(B&&d.getTime()===B.getTime()) btn.classList.add('end');
      });
    }

    function onDay(y){
      if(!pStart || (pStart&&pEnd)){ pStart=y; pEnd=null; hover=null; paintRange(); }
      else { pEnd=y; hover=null; if(fromYmd(pEnd)<fromYmd(pStart)){ var t=pStart; pStart=pEnd; pEnd=t; } apply(); }
    }
    function onHover(y){ if(pStart&&!pEnd){ hover=y; paintRange(); } }

    function applyShortcut(key){
      var b=bounds();
      if(key===''){ pStart=null; pEnd=null; hover=null; apply(); return; }
      var qr=quickRange(key,new Date());
      var cr=clampRange(qr.ini,qr.fim,b.min,b.max);
      if(!cr) return;
      pStart=cr.ini; pEnd=cr.fim; hover=null; apply();
    }

    function stepPar(delta){
      var idx=monthIndex(view.ly,view.lm)+delta;
      var ym=idxToYM(idx); view.ly=ym.year; view.lm=ym.month; renderAll();
    }
    function setLeftFromSelect(side,year,month){
      var idx=monthIndex(year,month); if(side==='R') idx-=1;
      var b=bounds(), lr=leftRange(b.min,b.max);
      if(lr){ if(idx<lr.lo) idx=lr.lo; if(idx>lr.hi) idx=lr.hi; }
      var ym=idxToYM(idx); view.ly=ym.year; view.lm=ym.month; renderAll();
    }

    function apply(){
      var a=pStart, bb=pEnd;
      if(a&&bb&&fromYmd(bb)<fromYmd(a)){ var t=a; a=bb; bb=t; }
      if(a&&bb){ iniInput.value=a; fimInput.value=bb; }
      else if(a){ iniInput.value=a; fimInput.value=a; }
      else { iniInput.value=''; fimInput.value=''; }
      fimInput.dispatchEvent(new Event('change',{bubbles:true}));
      close();
    }

    function wireDays(){
      popup.querySelectorAll('.rp-day:not(.empty):not(.disabled)').forEach(function(btn){
        btn.onclick=function(){ onDay(btn.getAttribute('data-ymd')); };
        btn.onmouseenter=function(){ onHover(btn.getAttribute('data-ymd')); };
      });
    }
    function wire(){
      var prev=popup.querySelector('.rp-prev'); if(prev) prev.onclick=function(){ if(!prev.disabled) stepPar(-1); };
      var next=popup.querySelector('.rp-next'); if(next) next.onclick=function(){ if(!next.disabled) stepPar(1); };
      popup.querySelectorAll('.rp-shortcut').forEach(function(b){ b.onclick=function(){ if(!b.disabled) applyShortcut(b.getAttribute('data-key')); }; });
      popup.querySelectorAll('.rp-mes').forEach(function(sel){
        sel.onchange=function(){
          var side=sel.getAttribute('data-side');
          var ano=+popup.querySelector('.rp-ano[data-side="'+side+'"]').value;
          setLeftFromSelect(side, ano, +sel.value);
        };
      });
      popup.querySelectorAll('.rp-ano').forEach(function(sel){
        sel.onchange=function(){
          var side=sel.getAttribute('data-side');
          var mes=+popup.querySelector('.rp-mes[data-side="'+side+'"]').value;
          setLeftFromSelect(side, +sel.value, mes);
        };
      });
      wireDays();
    }

    function open(){
      if(_openClose && _openClose!==close) _openClose();
      pStart=raw(iniInput)||null; pEnd=raw(fimInput)||null; hover=null;
      var b=bounds();
      var v=initialLeft(pStart,pEnd,b.min,b.max); view.ly=v.year; view.lm=v.month;
      renderAll();
      popup.style.display='flex';
      _openClose=close;
    }
    function close(){ popup.style.display='none'; if(_openClose===close) _openClose=null; }

    field.addEventListener('click',function(e){ e.stopPropagation(); if(popup.style.display==='none') open(); else close(); });
    popup.addEventListener('click',function(e){ e.stopPropagation(); });

    syncField();
  }

  document.addEventListener('click',function(){ if(_openClose) _openClose(); });

  window.AppRangePicker={ attach:attach, setBounds:setBounds };
})();

/**
 * Liga o range picker nos pares De/Até (Indicadores e Analistas).
 */
(function () {
  function initRangePickers() {
    if (!window.AppRangePicker) return;
    [['filtroDataIni','filtroDataFim'], ['uDataIni','uDataFim']].forEach(function (par) {
      var a = document.getElementById(par[0]), b = document.getElementById(par[1]);
      if (a && b) window.AppRangePicker.attach(a, b);
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', initRangePickers);
  else initRangePickers();
})();

/**
 * Limita a navegação do range picker às datas que realmente existem na
 * planilha carregada (o painel esquerdo vai até max-1 e o direito é +1).
 */
function populateDateBounds(data) {
  function bounds(list) {
    let min = null, max = null;
    list.forEach(d => { if (!d) return; const t = new Date(d); if (isNaN(t)) return; if (!min || t < min) min = t; if (!max || t > max) max = t; });
    return { min, max };
  }
  function toInputValue(d) {
    const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, '0'), day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }
  function apply(iniId, range) {
    if (!range.min || !range.max || !window.AppRangePicker) return;
    window.AppRangePicker.setBounds(iniId, toInputValue(range.min), toInputValue(range.max));
  }
  apply('filtroDataIni', bounds(data.map(a => a.dataAlerta)));      // Indicadores: data do alerta
  apply('uDataIni', bounds(data.map(a => a.dataPrimeiraTratativa))); // Analistas: 1ª tratativa
}


function populateFiltroPlaca(data) {
  const list = document.getElementById('placaList');
  if (!list) return;
  const placas = Array.from(new Set(data.map(a => a.placa).filter(Boolean))).sort();
  list.innerHTML = placas.map(p => `<option value="${p}"></option>`).join('');
}

function resetFiltrosInputs() {
  const dataIni = document.getElementById('filtroDataIni');
  const dataFim = document.getElementById('filtroDataFim');
  const selectPlaca = document.getElementById('filtroPlaca');
  if (dataIni) dataIni.value = '';
  if (dataFim) dataFim.value = '';
  if (selectPlaca) selectPlaca.value = '';
}

/**
 * AppUsuarios — aba Usuários com KPIs, linha do tempo, tipos por usuário e tabela detalhada.
 */
(function () {
  let _chartsU = {};
  const PALETTE = ['#5B8DEF','#36C2B4','#F2A33C','#E5484D','#9B7BFF','#3BC9DB','#F783AC','#94D82D','#FFA94D','#20C997'];

  function renderUsuarios(data) {
    const { groupByUsuario } = window.AppStats;
    const { fmtMin, fmtPct } = window.AppFormat;
    const userArr = groupByUsuario(data);
    const totalTratativas = userArr.reduce(function(s, u) { return s + u.count; }, 0);

    // --- KPIs ---
    document.getElementById('kpiUsuariosAtivos').textContent = userArr.length;
    var totalProcedentes = data.filter(function(a){ return a.tratado && !a.falsoPositivo; }).length;
    document.getElementById('kpiUsuarioProcedentes').textContent = totalProcedentes.toLocaleString('pt-BR');
    var maisRapido = userArr.filter(function(u) { return u.tempoMedio !== null; }).sort(function(a,b){ return a.tempoMedio - b.tempoMedio; })[0];
    document.getElementById('kpiUsuarioMaisRapido').textContent = maisRapido ? fmtMin(maisRapido.tempoMedio) : '—';
    var menorFP = userArr.filter(function(u){ return u.count >= 5; }).sort(function(a,b){ return a.taxaFp - b.taxaFp; })[0];
    document.getElementById('kpiUsuarioMenorFP').textContent = menorFP ? fmtPct(menorFP.taxaFp) : '—';

    // --- Tabela detalhada ---
    document.getElementById('tblUsuariosDetalhe').innerHTML = userArr.map(function(u) {
      return '<tr>' +
        '<td>' + u.nome + '</td>' +
        '<td class="num">' + u.count + '</td>' +
        '<td class="num"><div class="bar-cell"><div class="bar-track"><div class="bar-fill" style="width:' + (u.count/Math.max(userArr[0].count,1)*100) + '%;background:var(--blue)"></div></div>' + fmtPct(totalTratativas ? u.count/totalTratativas : 0) + '</div></td>' +
        '<td class="num">' + fmtMin(u.tempoMedio) + '</td>' +
        '<td class="num">' + fmtMin(u.tempoMin) + '</td>' +
        '<td class="num">' + fmtMin(u.tempoMax) + '</td>' +
        '<td class="num">' + u.fp + '</td>' +
        '<td class="num">' + fmtPct(u.taxaFp) + '</td>' +
        '<td>' + u.tipoMaisFrequente + '</td>' +
      '</tr>';
    }).join('');

    // --- Tabela tipos por usuário ---
    var todostipos = Array.from(new Set(userArr.flatMap(function(u){ return Object.keys(u.tipoMap); }))).sort();
    var tblEl = document.getElementById('tblUsuariosTipos');
    var header = '<thead><tr><th>Usuário</th>' + todostipos.map(function(t){ return '<th class="num" style="font-size:10px;">' + t.replace('Aviso de ','').replace('Falta de uso do ','') + '</th>'; }).join('') + '<th class="num">Total</th></tr></thead>';
    var body = '<tbody>' + userArr.map(function(u) {
      return '<tr><td>' + u.nome.split(' ')[0] + '</td>' + todostipos.map(function(t) {
        var cnt = u.tipoMap[t] || 0;
        var pct = u.count ? (cnt/u.count*100).toFixed(0) : 0;
        return '<td class="num" style="font-size:11px;">' + (cnt ? cnt + '<br><span style="color:var(--muted);font-size:9.5px;">' + pct + '%</span>' : '—') + '</td>';
      }).join('') + '<td class="num"><strong>' + u.count + '</strong></td></tr>';
    }).join('') + '</tbody>';
    tblEl.innerHTML = header + body;

    // --- Gráficos ---
    Object.values(_chartsU).forEach(function(c){ c.destroy(); });
    var rootStyles = getComputedStyle(document.documentElement);
    var gridColor = rootStyles.getPropertyValue('--line').trim();
    var top10 = userArr.slice(0, 10);

    _chartsU.bar = new Chart(document.getElementById('chartUsuariosBar'), {
      type: 'bar',
      data: {
        labels: top10.map(function(u){ return u.nome.split(' ')[0]; }),
        datasets: [{ label: 'Tratativas', data: top10.map(function(u){ return u.count; }), backgroundColor: '#5B8DEF', borderRadius: 5 }]
      },
      options: {
        indexAxis: 'y',
        plugins: {
          legend: { display: false },
          datalabels: { display: false }
        },
        layout: { padding: { right: 48 } },
        scales: { x: { grid: { color: gridColor } }, y: { grid: { display: false } } },
        maintainAspectRatio: false,
        animation: { onComplete: function() {
          var chart = this;
          var ctx = chart.ctx;
          ctx.font = '11px JetBrains Mono, monospace';
          ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--muted').trim() || '#7C8AA5';
          ctx.textAlign = 'left';
          chart.data.datasets.forEach(function(ds, di) {
            var meta = chart.getDatasetMeta(di);
            meta.data.forEach(function(bar, i) {
              var val = ds.data[i];
              ctx.fillText(val.toLocaleString('pt-BR'), bar.x + 6, bar.y + 4);
            });
          });
        }}
      }
    });

    _chartsU.fp = new Chart(document.getElementById('chartUsuariosFP'), {
      type: 'bar',
      data: {
        labels: top10.map(function(u){ return u.nome.split(' ')[0]; }),
        datasets: [
          { label: 'Falso +', data: top10.map(function(u){ return +(u.taxaFp * 100).toFixed(1); }), backgroundColor: '#E5484D', borderRadius: 5 },
          { label: 'Verdadeiros', data: top10.map(function(u){ return +(u.taxaProc * 100).toFixed(1); }), backgroundColor: '#36C2B4', borderRadius: 5 }
        ]
      },
      options: {
        indexAxis: 'y',
        layout: { padding: { right: 52 } },
        plugins: { legend: { display: true, position: 'top', labels: { boxWidth: 10, font: { size: 11 } } }, tooltip: { callbacks: { label: function(ctx){ return ctx.dataset.label + ': ' + ctx.parsed.x.toFixed(1) + '%'; } } } },
        scales: { x: { grid: { color: gridColor }, ticks: { callback: function(v){ return v + '%'; } } }, y: { grid: { display: false } } },
        maintainAspectRatio: false,
        animation: { onComplete: function() {
          var chart = this;
          var ctx = chart.ctx;
          ctx.font = '11px JetBrains Mono, monospace';
          ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--muted').trim() || '#7C8AA5';
          ctx.textAlign = 'left';
          chart.data.datasets.forEach(function(ds, di) {
            var meta = chart.getDatasetMeta(di);
            meta.data.forEach(function(bar, i) {
              ctx.fillText(ds.data[i].toFixed(1) + '%', bar.x + 6, bar.y + 4);
            });
          });
        }}
      }
    });

    // --- Linha do tempo: tratativas por usuário por dia ---
    var todasDias = Array.from(new Set(data.filter(function(a){ return a.dataPrimeiraTratativa; }).map(function(a){
      return new Date(a.dataPrimeiraTratativa).toLocaleDateString('pt-BR');
    }))).sort(function(a, b){
      var pa = a.split('/'), pb = b.split('/');
      return new Date(pa[2],pa[1]-1,pa[0]) - new Date(pb[2],pb[1]-1,pb[0]);
    });

    var datasets = top10.map(function(u, i) {
      return {
        label: u.nome.split(' ')[0],
        data: todasDias.map(function(d){ return u.diaMap[d] || 0; }),
        borderColor: PALETTE[i % PALETTE.length],
        backgroundColor: 'transparent',
        tension: 0.3,
        pointRadius: 2,
        borderWidth: 2,
      };
    });

    _chartsU.linha = new Chart(document.getElementById('chartUsuariosLinha'), {
      type: 'line',
      data: { labels: todasDias, datasets: datasets },
      options: {
        interaction: { mode: 'index', intersect: false },
        plugins: { legend: { position: 'top', labels: { boxWidth: 10, font: { size: 10 } } } },
        scales: {
          x: { grid: { display: false } },
          y: { grid: { color: gridColor }, beginAtZero: true }
        },
        maintainAspectRatio: false
      }
    });
  }

  window.AppUsuarios = { renderUsuarios };
})();

/**
 * Filtro de período da aba Usuários — independente do filtro da aba Indicadores.
 */
(function () {
  function parseD(value, endOfDay) {
    if (!value) return null;
    var parts = value.split('-');
    return endOfDay
      ? new Date(+parts[0], +parts[1]-1, +parts[2], 23, 59, 59, 999)
      : new Date(+parts[0], +parts[1]-1, +parts[2], 0, 0, 0, 0);
  }

  function applyUsuarioFilter() {
    var all = window.__currentAlertData || [];
    var ini = parseD((document.getElementById('uDataIni')||{}).value, false);
    var fim = parseD((document.getElementById('uDataFim')||{}).value, true);
    var filtered = all.filter(function(a) {
      if (ini && (!a.dataPrimeiraTratativa || new Date(a.dataPrimeiraTratativa) < ini)) return false;
      if (fim && (!a.dataPrimeiraTratativa || new Date(a.dataPrimeiraTratativa) > fim)) return false;
      return true;
    });
    if (window.AppUsuarios) window.AppUsuarios.renderUsuarios(filtered);
  }

  document.addEventListener('DOMContentLoaded', function() {
    var uIni = document.getElementById('uDataIni');
    var uFim = document.getElementById('uDataFim');
    var uLimpar = document.getElementById('uLimpar');
    if (!uIni) return;

    uIni.addEventListener('change', applyUsuarioFilter);
    uFim.addEventListener('change', applyUsuarioFilter);
    uLimpar.addEventListener('click', function() {
      uIni.value=''; uFim.value='';
      applyUsuarioFilter();
    });
  });
})();
