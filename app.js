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
    if (!map[t]) map[t] = { count: 0, fp: 0, tempos: [] };
    map[t].count++;
    if (a.falsoPositivo) map[t].fp++;
    if (a.tempoResposta !== null) map[t].tempos.push(a.tempoResposta);
  });
  return Object.entries(map)
    .map(([nome, v]) => ({ nome, count: v.count, fp: v.fp, tempoMedio: average(v.tempos) }))
    .sort((a, b) => b.count - a.count);
}

function groupByRisco(data) {
  const map = {};
  data.forEach(a => {
    const r = a.risco || 'Não informado';
    map[r] = (map[r] || 0) + 1;
  });
  const ordem = ['Alto', 'Médio', 'Baixo'];
  const conhecidos = ordem.filter(r => map[r]).map(r => ({ nome: r, count: map[r] }));
  const outros = Object.entries(map).filter(([k]) => !ordem.includes(k)).map(([nome, count]) => ({ nome, count }));
  return conhecidos.concat(outros);
}

function topPlacas(data, limit = 10) {
  const map = {};
  data.forEach(a => {
    const p = a.placa || '—';
    map[p] = (map[p] || 0) + 1;
  });
  return Object.entries(map).map(([nome, count]) => ({ nome, count }))
    .sort((a, b) => b.count - a.count).slice(0, limit);
}

function groupByUsuario(data) {
  const map = {};
  data.forEach(a => {
    if (!a.usuario) return;
    if (!map[a.usuario]) map[a.usuario] = { count: 0, fp: 0, tempos: [] };
    map[a.usuario].count++;
    if (a.falsoPositivo) map[a.usuario].fp++;
    if (a.tempoResposta !== null) map[a.usuario].tempos.push(a.tempoResposta);
  });
  return Object.entries(map)
    .map(([nome, v]) => ({ nome, count: v.count, fp: v.fp, tempoMedio: average(v.tempos) }))
    .sort((a, b) => b.count - a.count);
}

function groupByDia(data) {
  const map = {};
  data.forEach(a => {
    if (!a.dataAlerta) return;
    const key = a.dataAlerta.toLocaleDateString('pt-BR');
    map[key] = (map[key] || 0) + 1;
  });
  return Object.entries(map).map(([data_, count]) => ({ data: data_, count }))
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
    </tr>`).join('');

  // --- por dia ---
  const diaArr = groupByDia(data);

  drawCharts(tipoArr, placaArr, diaArr);
}

function drawCharts(tipoArr, placaArr, diaArr) {
  Object.values(_charts).forEach(c => c.destroy());
  const palette = ['#5B8DEF', '#36C2B4', '#F2A33C', '#E5484D', '#9B7BFF', '#3BC9DB', '#F783AC', '#94D82D', '#FFA94D'];
  const rootStyles = getComputedStyle(document.documentElement);
  const mutedColor = rootStyles.getPropertyValue('--muted').trim() || '#7C8AA5';
  const panelColor = rootStyles.getPropertyValue('--panel').trim() || '#111A2E';
  Chart.defaults.font.family = "'JetBrains Mono', monospace";
  Chart.defaults.color = mutedColor;

  _charts.tipo = new Chart(document.getElementById('chartTipo'), {
    type: 'doughnut',
    data: { labels: tipoArr.map(t => t.nome), datasets: [{ data: tipoArr.map(t => t.count), backgroundColor: palette, borderColor: panelColor, borderWidth: 2 }] },
    options: { plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } } }, maintainAspectRatio: false }
  });

  _charts.placas = new Chart(document.getElementById('chartPlacas'), {
    type: 'bar',
    data: { labels: placaArr.map(p => p.nome), datasets: [{ data: placaArr.map(p => p.count), backgroundColor: '#F2A33C', borderRadius: 4 }] },
    options: { plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } }, y: { grid: { color: rootStyles.getPropertyValue('--line').trim() || 'rgba(255,255,255,0.05)' } } }, maintainAspectRatio: false }
  });

  _charts.dia = new Chart(document.getElementById('chartDia'), {
    type: 'line',
    data: { labels: diaArr.map(d => d.data), datasets: [{ data: diaArr.map(d => d.count), borderColor: '#36C2B4', backgroundColor: 'rgba(54,194,180,0.15)', fill: true, tension: 0.3, pointRadius: 3 }] },
    options: { plugins: { legend: { display: false } }, scales: { x: { grid: { display: false } }, y: { grid: { color: rootStyles.getPropertyValue('--line').trim() || 'rgba(255,255,255,0.05)' } } }, maintainAspectRatio: false }
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
      <td>${a.placa || '—'}</td>
      <td>${a.motorista || '—'}</td>
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
        resetFiltrosInputs();
        window.AppDashboard.renderDashboard(parsed);
        window.AppDadosTable.renderDadosTable(parsed);

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
    const placa = selectPlaca.value;

    const filtered = all.filter(a => {
      if (ini && (!a.dataAlerta || a.dataAlerta < ini)) return false;
      if (fim && (!a.dataAlerta || a.dataAlerta > fim)) return false;
      if (placa && a.placa !== placa) return false;
      return true;
    });

    if (window.AppDashboard) window.AppDashboard.renderDashboard(filtered);
    if (window.AppDadosTable) window.AppDadosTable.renderDadosTable(filtered);
  }

  dataIni.addEventListener('change', () => { const sel = document.getElementById('filtroPeriodoRapido'); if (sel) sel.value = ''; applyFilters(); });
  dataFim.addEventListener('change', () => { const sel = document.getElementById('filtroPeriodoRapido'); if (sel) sel.value = ''; applyFilters(); });
  selectPlaca.addEventListener('change', applyFilters);
  btnLimpar.addEventListener('click', () => {
    dataIni.value = '';
    dataFim.value = '';
    selectPlaca.value = '';
    const sel = document.getElementById('filtroPeriodoRapido');
    if (sel) sel.value = '';
    applyFilters();
  });

  function toInputValue(d) {
    const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, '0'), day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function computeRange(key) {
    const now = new Date();
    let ini, fim;
    switch (key) {
      case 'mes-atual':
        ini = new Date(now.getFullYear(), now.getMonth(), 1);
        fim = new Date(now.getFullYear(), now.getMonth() + 1, 0);
        break;
      case 'mes-anterior':
        ini = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        fim = new Date(now.getFullYear(), now.getMonth(), 0);
        break;
      case 'trimestre-atual': {
        const tStart = Math.floor(now.getMonth() / 3) * 3;
        ini = new Date(now.getFullYear(), tStart, 1);
        fim = new Date(now.getFullYear(), tStart + 3, 0);
        break;
      }
      case 'semestre-atual': {
        const sStart = now.getMonth() < 6 ? 0 : 6;
        ini = new Date(now.getFullYear(), sStart, 1);
        fim = new Date(now.getFullYear(), sStart + 6, 0);
        break;
      }
      case 'ano-atual':
        ini = new Date(now.getFullYear(), 0, 1);
        fim = new Date(now.getFullYear(), 11, 31);
        break;
      case 'ano-anterior':
        ini = new Date(now.getFullYear() - 1, 0, 1);
        fim = new Date(now.getFullYear() - 1, 11, 31);
        break;
      default:
        return null;
    }
    return { ini, fim };
  }

  const selectPeriodoRapido = document.getElementById('filtroPeriodoRapido');
  if (selectPeriodoRapido) {
    selectPeriodoRapido.addEventListener('change', () => {
      const range = computeRange(selectPeriodoRapido.value);
      if (!range) {
        dataIni.value = '';
        dataFim.value = '';
      } else {
        dataIni.value = toInputValue(range.ini);
        dataFim.value = toInputValue(range.fim);
      }
      applyFilters();
    });
  }

  window.AppFiltrosIndicadores = { applyFilters };
})();

/**
 * Preenche o seletor de placas com as placas únicas da planilha carregada.
 */
function populateFiltroPlaca(data) {
  const select = document.getElementById('filtroPlaca');
  if (!select) return;
  const placas = Array.from(new Set(data.map(a => a.placa).filter(Boolean))).sort();
  select.innerHTML = '<option value="">Todas</option>' + placas.map(p => `<option value="${p}">${p}</option>`).join('');
}

function resetFiltrosInputs() {
  const dataIni = document.getElementById('filtroDataIni');
  const dataFim = document.getElementById('filtroDataFim');
  const selectPlaca = document.getElementById('filtroPlaca');
  if (dataIni) dataIni.value = '';
  if (dataFim) dataFim.value = '';
  if (selectPlaca) selectPlaca.value = '';
}
