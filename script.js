/**
 * Precificador Dinâmico — Platinum Edition
 * Lógica de cálculos em tempo real, simulador reverso, gestão de histórico e exportação CSV/Excel.
 */

const STORAGE_KEY = 'precificador_platinum_produtos_v1';

// Dados originais exatos da planilha fornecida pelo usuário
const INITIAL_PRODUCTS = [
  { id: 'init-1', sku: '322', precoFinal: 32, precoInicial: 15.52, pct12: 12, pct16: 16, embalagem: 0.2, erpCusto: 2.15, frete: null },
  { id: 'init-2', sku: '10',  precoFinal: 24, precoInicial: 11.9,  pct12: 12, pct16: 16, embalagem: 0.2, erpCusto: 2.15, frete: null },
  { id: 'init-3', sku: '1',   precoFinal: 42, precoInicial: 27.47, pct12: 12, pct16: 16, embalagem: 0.2, erpCusto: 0,    frete: null },
  { id: 'init-4', sku: '323', precoFinal: 179.9, precoInicial: 62.08, pct12: 12, pct16: 16, embalagem: 1, erpCusto: 4.3, frete: 26.35 },
  { id: 'init-5', sku: '10',  precoFinal: 35, precoInicial: 15.98, pct12: 12, pct16: 16, embalagem: 0.2, erpCusto: 2.15, frete: null },
  { id: 'init-6', sku: '9',   precoFinal: 25, precoInicial: 9.45,  pct12: 12, pct16: 16, embalagem: 0.2, erpCusto: 2.15, frete: null },
  { id: 'init-7', sku: '70',  precoFinal: 40.99, precoInicial: 19.05, pct12: 12, pct16: 16, embalagem: 0.5, erpCusto: 2.15, frete: null },
  { id: 'init-8', sku: '11',  precoFinal: 33, precoInicial: 14,    pct12: 12, pct16: 16, embalagem: 0.2, erpCusto: 2.15, frete: null },
  { id: 'init-9', sku: '12',  precoFinal: 37.9, precoInicial: 17.31, pct12: 12, pct16: 16, embalagem: 0.2, erpCusto: 2.15, frete: null },
  { id: 'init-10', sku: '13', precoFinal: 28.99, precoInicial: 10.41, pct12: 12, pct16: 16, embalagem: 0.2, erpCusto: 2.15, frete: null }
];

// Estado da Aplicação
let products = [];
let editingProductId = null;
let sortState = { column: null, direction: 'asc' };

// ============================================================================
// FUNÇÕES MATEMÁTICAS E FORMATAÇÃO
// ============================================================================

/**
 * Arredonda evitando artefatos de ponto flutuante (ex: 33.177200000000005 -> 33.1772)
 */
function cleanFloat(num, decimals = 4) {
  const factor = Math.pow(10, decimals);
  return Math.round((Number(num) + Number.EPSILON) * factor) / factor;
}

/**
 * Calcula todos os campos derivados de um produto seguindo a fórmula da planilha:
 * 12% PF = Preço Final * (pct12 / 100)
 * 16% PF = Preço Final * (pct16 / 100)
 * Custo Total = Preço Inicial + 12% PF + 16% PF + Embalagem + ERP Custo + Frete
 * Margem = Preço Final - Custo Total
 */
function computeProductMetrics(raw) {
  const precoFinal = Number(raw.precoFinal) || 0;
  const precoInicial = Number(raw.precoInicial) || 0;
  const pct12 = raw.pct12 !== undefined ? Number(raw.pct12) : 12;
  const pct16 = raw.pct16 !== undefined ? Number(raw.pct16) : 16;
  const embalagem = Number(raw.embalagem) || 0;
  const erpCusto = Number(raw.erpCusto) || 0;
  const hasFrete = raw.frete !== null && raw.frete !== '' && !isNaN(Number(raw.frete)) && Number(raw.frete) > 0;
  const freteVal = hasFrete ? Number(raw.frete) : 0;

  const taxa12 = cleanFloat(precoFinal * (pct12 / 100), 4);
  const taxa16 = cleanFloat(precoFinal * (pct16 / 100), 4);
  const custoTotal = cleanFloat(precoInicial + taxa12 + taxa16 + embalagem + erpCusto + freteVal, 4);
  const margem = cleanFloat(precoFinal - custoTotal, 4);
  const margemPct = precoFinal > 0 ? cleanFloat((margem / precoFinal) * 100, 2) : 0;
  const markupPct = precoInicial > 0 ? cleanFloat(((precoFinal - precoInicial) / precoInicial) * 100, 1) : 0;

  return {
    ...raw,
    precoFinal: cleanFloat(precoFinal, 4),
    precoInicial: cleanFloat(precoInicial, 4),
    pct12,
    pct16,
    taxa12,
    taxa16,
    embalagem: cleanFloat(embalagem, 4),
    erpCusto: cleanFloat(erpCusto, 4),
    frete: hasFrete ? cleanFloat(freteVal, 4) : null,
    custoTotal,
    margem,
    margemPct,
    markupPct
  };
}

/**
 * Formata valor em Moeda Real (R$)
 * Se o valor tiver mais de 2 casas decimais significativas (ex: 4.9188), exibe até 4 casas quando solicitado
 */
function formatBRL(value, allowFourDecimals = false) {
  if (value === null || value === undefined || isNaN(value)) return '—';
  const num = Number(value);
  const hasExtraDecimals = allowFourDecimals && Math.abs(Math.round(num * 100) - num * 100) > 0.001;
  return num.toLocaleString('pt-BR', {
    style: 'currency',
    currency: 'BRL',
    minimumFractionDigits: 2,
    maximumFractionDigits: hasExtraDecimals ? 4 : 2
  });
}

/**
 * Formata número para exportação CSV original (ponto decimal, sem zeros desnecessários)
 */
function formatRawCSVNumber(value) {
  if (value === null || value === undefined || value === '') return '';
  return String(cleanFloat(value, 4));
}

// ============================================================================
// PERSISTÊNCIA (LOCALSTORAGE)
// ============================================================================

function loadProductsFromStorage() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        products = parsed.map(computeProductMetrics);
        return;
      }
    }
  } catch (err) {
    console.warn('Erro ao ler localStorage:', err);
  }
  products = INITIAL_PRODUCTS.map(computeProductMetrics);
  saveProductsToStorage();
}

function saveProductsToStorage() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(products));
  } catch (err) {
    console.warn('Erro ao salvar no localStorage:', err);
  }
  updateHistoryBadge();
}

function updateHistoryBadge() {
  const badge = document.getElementById('history-count-badge');
  if (badge) {
    badge.textContent = products.length;
  }
}

// ============================================================================
// NAVEGAÇÃO ENTRE GUIAS (TABS)
// ============================================================================

function switchTab(tabName) {
  const btnCalc = document.getElementById('tab-btn-calc');
  const btnHistory = document.getElementById('tab-btn-history');
  const panelCalc = document.getElementById('panel-calc');
  const panelHistory = document.getElementById('panel-history');

  if (tabName === 'calc') {
    btnCalc.classList.add('active');
    btnCalc.setAttribute('aria-selected', 'true');
    btnHistory.classList.remove('active');
    btnHistory.setAttribute('aria-selected', 'false');

    panelCalc.classList.add('active');
    panelCalc.hidden = false;
    panelHistory.classList.remove('active');
    panelHistory.hidden = true;
  } else {
    btnHistory.classList.add('active');
    btnHistory.setAttribute('aria-selected', 'true');
    btnCalc.classList.remove('active');
    btnCalc.setAttribute('aria-selected', 'false');

    panelHistory.classList.add('active');
    panelHistory.hidden = false;
    panelCalc.classList.remove('active');
    panelCalc.hidden = true;

    renderHistoryTable();
    renderCatalogKPIs();
  }
}

// ============================================================================
// MODO CÁLCULO — ATUALIZAÇÃO EM TEMPO REAL E SIMULADOR
// ============================================================================

function getFormValues() {
  const sku = document.getElementById('input-sku').value.trim();
  const precoInicial = parseFloat(document.getElementById('input-preco-inicial').value) || 0;
  const precoFinal = parseFloat(document.getElementById('input-preco-final').value) || 0;
  const pct12 = parseFloat(document.getElementById('input-taxa-12').value) || 0;
  const pct16 = parseFloat(document.getElementById('input-taxa-16').value) || 0;
  const embalagem = parseFloat(document.getElementById('input-embalagem').value) || 0;
  const erpCusto = parseFloat(document.getElementById('input-erp').value) || 0;
  const freteRaw = document.getElementById('input-frete').value.trim();
  const frete = freteRaw === '' ? null : parseFloat(freteRaw);

  return computeProductMetrics({
    sku,
    precoInicial,
    precoFinal,
    pct12,
    pct16,
    embalagem,
    erpCusto,
    frete
  });
}

function updateLiveCalculation() {
  const m = getFormValues();

  // 1. Atualiza caixas de taxas 12% e 16%
  document.getElementById('display-val-12').textContent = formatBRL(m.taxa12, true);
  document.getElementById('display-val-16').textContent = formatBRL(m.taxa16, true);
  document.getElementById('label-pct-12').textContent = m.pct12;
  document.getElementById('label-pct-16').textContent = m.pct16;

  // 2. Cabeçalho do Card de Resultado
  document.getElementById('result-sku-display').textContent = m.sku ? `SKU ${m.sku}` : 'SKU — (Simulação)';

  // Status da Margem
  const statusBadge = document.getElementById('margin-status-badge');
  const heroMarginBox = document.getElementById('hero-margin-box');
  heroMarginBox.classList.remove('metric-highlight', 'metric-warn', 'metric-neg');
  statusBadge.classList.remove('status-positive', 'status-warning', 'status-negative');

  if (m.margem <= 0) {
    statusBadge.textContent = 'Prejuízo / Negativa';
    statusBadge.classList.add('status-negative');
    heroMarginBox.classList.add('metric-neg');
  } else if (m.margemPct < 10) {
    statusBadge.textContent = 'Margem Apertada (<10%)';
    statusBadge.classList.add('status-warning');
    heroMarginBox.classList.add('metric-warn');
  } else {
    statusBadge.textContent = `Margem Saudável (${m.margemPct.toFixed(1).replace('.', ',')}%)`;
    statusBadge.classList.add('status-positive');
    heroMarginBox.classList.add('metric-highlight');
  }

  // 3. KPIs Principais
  document.getElementById('res-preco-final').textContent = formatBRL(m.precoFinal);
  document.getElementById('res-markup-info').textContent =
    m.precoInicial > 0
      ? `Markup: ${m.markupPct >= 0 ? '+' : ''}${m.markupPct.toFixed(1).replace('.', ',')}%`
      : 'Markup: —';

  document.getElementById('res-custo-total').textContent = formatBRL(m.custoTotal, true);
  const custoShare = m.precoFinal > 0 ? ((m.custoTotal / m.precoFinal) * 100).toFixed(1).replace('.', ',') : '0,0';
  document.getElementById('res-custo-pct').textContent = `${custoShare}% do Preço Final`;

  document.getElementById('res-margem-valor').textContent = formatBRL(m.margem, true);
  document.getElementById('res-margem-pct').textContent = `${m.margemPct.toFixed(2).replace('.', ',')}% sobre a venda`;

  // 4. Barra de Decomposição Visual
  const pf = m.precoFinal > 0 ? m.precoFinal : 1;
  const opCosts = m.embalagem + m.erpCusto + (m.frete || 0);

  const pInicial = Math.max(0, Math.min(100, (m.precoInicial / pf) * 100));
  const p12 = Math.max(0, Math.min(100, (m.taxa12 / pf) * 100));
  const p16 = Math.max(0, Math.min(100, (m.taxa16 / pf) * 100));
  const pOp = Math.max(0, Math.min(100, (opCosts / pf) * 100));
  const pMargem = m.margem > 0 ? Math.max(0, Math.min(100, (m.margem / pf) * 100)) : 0;

  document.getElementById('bar-seg-inicial').style.width = `${pInicial}%`;
  document.getElementById('bar-seg-12').style.width = `${p12}%`;
  document.getElementById('bar-seg-16').style.width = `${p16}%`;
  document.getElementById('bar-seg-op').style.width = `${pOp}%`;
  document.getElementById('bar-seg-margem').style.width = `${pMargem}%`;

  document.getElementById('leg-inicial').textContent = `${pInicial.toFixed(1).replace('.', ',')}%`;
  document.getElementById('leg-label-12').textContent = `${m.pct12}% PF`;
  document.getElementById('leg-12').textContent = `${p12.toFixed(1).replace('.', ',')}%`;
  document.getElementById('leg-label-16').textContent = `${m.pct16}% PF`;
  document.getElementById('leg-16').textContent = `${p16.toFixed(1).replace('.', ',')}%`;
  document.getElementById('leg-op').textContent = `${pOp.toFixed(1).replace('.', ',')}%`;
  document.getElementById('leg-margem').textContent = `${m.margemPct.toFixed(1).replace('.', ',')}%`;

  document.getElementById('composition-summary-text').textContent =
    m.margem >= 0
      ? `Lucro retido: R$ ${m.margem.toFixed(2).replace('.', ',')} por unidade`
      : `Déficit de R$ ${Math.abs(m.margem).toFixed(2).replace('.', ',')} por unidade`;

  // 5. Demonstrativo Detalhado
  document.getElementById('bd-preco-inicial').textContent = formatBRL(m.precoInicial, true);
  document.getElementById('bd-label-12').textContent = `Taxa ${m.pct12}% PF`;
  document.getElementById('bd-val-12').textContent = formatBRL(m.taxa12, true);
  document.getElementById('bd-label-16').textContent = `Taxa ${m.pct16}% PF`;
  document.getElementById('bd-val-16').textContent = formatBRL(m.taxa16, true);
  document.getElementById('bd-embalagem').textContent = formatBRL(m.embalagem, true);
  document.getElementById('bd-erp').textContent = formatBRL(m.erpCusto, true);
  document.getElementById('bd-frete').textContent = m.frete !== null ? formatBRL(m.frete, true) : '—';
  document.getElementById('bd-custo-total').textContent = formatBRL(m.custoTotal, true);
  document.getElementById('bd-margem').textContent = formatBRL(m.margem, true);

  // 6. Atualiza Simulador Reverso (Metas de Margem 0%, 10%, 15%, 20%, 25%)
  updateReverseSimulator(m);
  highlightActiveChips();
}

/**
 * Calcula o Preço Final necessário para atingir determinada % de margem sobre a venda:
 * PF = (Custos Fixos) / (1 - (pct12 + pct16 + targetMarginPct)/100)
 */
function updateReverseSimulator(m) {
  const fixedCosts = m.precoInicial + m.embalagem + m.erpCusto + (m.frete || 0);
  const taxRateFraction = (m.pct12 + m.pct16) / 100;

  const targets = [0, 10, 15, 20, 25];
  targets.forEach((tPct) => {
    const el = document.getElementById(`sim-price-${tPct}`);
    if (!el) return;
    const denominator = 1 - taxRateFraction - tPct / 100;
    if (denominator <= 0 || fixedCosts <= 0) {
      el.textContent = '—';
    } else {
      const suggestedPF = fixedCosts / denominator;
      el.textContent = formatBRL(suggestedPF);
    }
  });
}

function highlightActiveChips() {
  document.querySelectorAll('.chip').forEach((chip) => {
    const targetId = chip.getAttribute('data-target');
    const chipVal = chip.getAttribute('data-value');
    const inputEl = document.getElementById(targetId);
    if (!inputEl) return;

    const currentRaw = inputEl.value.trim();
    if (chipVal === '' && currentRaw === '') {
      chip.classList.add('active-chip');
    } else if (chipVal !== '' && currentRaw !== '' && Math.abs(parseFloat(currentRaw) - parseFloat(chipVal)) < 0.001) {
      chip.classList.add('active-chip');
    } else {
      chip.classList.remove('active-chip');
    }
  });
}

// ============================================================================
// SALVAR, EDITAR, DUPLICAR E EXCLUIR PRODUTOS
// ============================================================================

function handleFormSubmit(event) {
  event.preventDefault();
  const m = getFormValues();

  if (!m.sku) {
    showToast('Informe o SKU do produto antes de salvar.');
    document.getElementById('input-sku').focus();
    return;
  }

  if (m.precoFinal <= 0) {
    showToast('O Preço Final deve ser maior que zero.');
    document.getElementById('input-preco-final').focus();
    return;
  }

  if (editingProductId) {
    const idx = products.findIndex((p) => p.id === editingProductId);
    if (idx !== -1) {
      products[idx] = { ...m, id: editingProductId };
      showToast(`Produto SKU ${m.sku} atualizado no histórico!`);
    }
    exitEditMode();
  } else {
    const newProduct = {
      ...m,
      id: 'prod-' + Date.now() + '-' + Math.random().toString(36).substring(2, 6)
    };
    products.unshift(newProduct);
    showToast(`Produto SKU ${m.sku} salvo no histórico!`);
  }

  saveProductsToStorage();
  renderHistoryTable();
  renderCatalogKPIs();
}

function loadProductIntoCalculator(id, isDuplicate = false) {
  const prod = products.find((p) => p.id === id);
  if (!prod) return;

  document.getElementById('input-sku').value = prod.sku;
  document.getElementById('input-preco-inicial').value = prod.precoInicial;
  document.getElementById('input-preco-final').value = prod.precoFinal;
  document.getElementById('input-taxa-12').value = prod.pct12 !== undefined ? prod.pct12 : 12;
  document.getElementById('input-taxa-16').value = prod.pct16 !== undefined ? prod.pct16 : 16;
  document.getElementById('input-embalagem').value = prod.embalagem;
  document.getElementById('input-erp').value = prod.erpCusto;
  document.getElementById('input-frete').value = prod.frete !== null ? prod.frete : '';

  if (isDuplicate) {
    exitEditMode();
    showToast(`Dados do SKU ${prod.sku} carregados como novo cálculo.`);
  } else {
    editingProductId = prod.id;
    document.getElementById('edit-mode-banner').classList.remove('hidden');
    document.getElementById('editing-sku-label').textContent = prod.sku;
    document.getElementById('btn-save-text').textContent = `Atualizar Produto (SKU ${prod.sku})`;
  }

  updateLiveCalculation();
  switchTab('calc');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function exitEditMode() {
  editingProductId = null;
  document.getElementById('edit-mode-banner').classList.add('hidden');
  document.getElementById('btn-save-text').textContent = 'Salvar Produto no Histórico';
}

function resetCalculatorForm() {
  exitEditMode();
  document.getElementById('input-sku').value = '';
  document.getElementById('input-preco-inicial').value = '';
  document.getElementById('input-preco-final').value = '';
  document.getElementById('input-taxa-12').value = '12';
  document.getElementById('input-taxa-16').value = '16';
  document.getElementById('input-embalagem').value = '0.20';
  document.getElementById('input-erp').value = '2.15';
  document.getElementById('input-frete').value = '';
  document.getElementById('input-target-brl').value = '';
  updateLiveCalculation();
  document.getElementById('input-sku').focus();
}

function deleteProduct(id) {
  const prod = products.find((p) => p.id === id);
  if (!prod) return;
  products = products.filter((p) => p.id !== id);
  if (editingProductId === id) {
    exitEditMode();
  }
  saveProductsToStorage();
  renderHistoryTable();
  renderCatalogKPIs();
  showToast(`Item SKU ${prod.sku} removido do histórico.`);
}

// ============================================================================
// HISTÓRICO DE PRODUTOS — TABELA, FILTROS E KPIS
// ============================================================================

function getFilteredAndSortedProducts() {
  const searchTerm = (document.getElementById('search-sku')?.value || '').trim().toLowerCase();
  const marginFilter = document.getElementById('filter-margin')?.value || 'all';

  let list = products.filter((item) => {
    const matchesSearch =
      !searchTerm ||
      String(item.sku).toLowerCase().includes(searchTerm) ||
      String(item.precoFinal).includes(searchTerm) ||
      String(item.precoInicial).includes(searchTerm);

    if (!matchesSearch) return false;

    if (marginFilter === 'high') return item.margemPct >= 20;
    if (marginFilter === 'medium') return item.margemPct >= 10 && item.margemPct < 20;
    if (marginFilter === 'low') return item.margemPct < 10;
    return true;
  });

  if (sortState.column) {
    const col = sortState.column;
    const dir = sortState.direction === 'asc' ? 1 : -1;
    list.sort((a, b) => {
      if (col === 'sku') {
        const numA = parseFloat(a.sku);
        const numB = parseFloat(b.sku);
        if (!isNaN(numA) && !isNaN(numB)) return (numA - numB) * dir;
        return String(a.sku).localeCompare(String(b.sku)) * dir;
      }
      const valA = a[col] === null ? -Infinity : Number(a[col]);
      const valB = b[col] === null ? -Infinity : Number(b[col]);
      return (valA - valB) * dir;
    });
  }

  return list;
}

function renderCatalogKPIs() {
  const total = products.length;
  document.getElementById('kpi-total-items').textContent = total;

  if (total === 0) {
    document.getElementById('kpi-avg-price').textContent = 'R$ 0,00';
    document.getElementById('kpi-avg-cost').textContent = 'R$ 0,00';
    document.getElementById('kpi-avg-margin').textContent = 'R$ 0,00';
    document.getElementById('kpi-avg-margin-pct').textContent = '0,0% de rentabilidade média';
    return;
  }

  const sumPrice = products.reduce((acc, p) => acc + p.precoFinal, 0);
  const sumCost = products.reduce((acc, p) => acc + p.custoTotal, 0);
  const sumMargin = products.reduce((acc, p) => acc + p.margem, 0);
  const avgMarginPct = sumPrice > 0 ? (sumMargin / sumPrice) * 100 : 0;

  document.getElementById('kpi-avg-price').textContent = formatBRL(sumPrice / total);
  document.getElementById('kpi-avg-cost').textContent = formatBRL(sumCost / total);
  document.getElementById('kpi-avg-margin').textContent = formatBRL(sumMargin / total);
  document.getElementById('kpi-avg-margin-pct').textContent = `${avgMarginPct.toFixed(2).replace('.', ',')}% de rentabilidade média`;
}

function renderHistoryTable() {
  const tbody = document.getElementById('products-tbody');
  const emptyState = document.getElementById('empty-state');
  const list = getFilteredAndSortedProducts();

  tbody.innerHTML = '';

  if (list.length === 0) {
    emptyState.classList.remove('hidden');
    return;
  }
  emptyState.classList.add('hidden');

  list.forEach((item) => {
    const tr = document.createElement('tr');

    let marginClass = 'margin-val-pos';
    let badgeClass = 'status-positive';
    if (item.margem <= 0) {
      marginClass = 'margin-val-neg';
      badgeClass = 'status-negative';
    } else if (item.margemPct < 10) {
      marginClass = 'margin-val-warn';
      badgeClass = 'status-warning';
    }

    tr.innerHTML = `
      <td><span class="sku-badge">${escapeHtml(item.sku)}</span></td>
      <td class="num-col"><strong>${formatBRL(item.precoFinal, true)}</strong></td>
      <td class="num-col">${formatBRL(item.precoInicial, true)}</td>
      <td class="num-col">${formatBRL(item.taxa12, true)}</td>
      <td class="num-col">${formatBRL(item.taxa16, true)}</td>
      <td class="num-col">${formatBRL(item.embalagem, true)}</td>
      <td class="num-col">${formatBRL(item.erpCusto, true)}</td>
      <td class="num-col">${item.frete !== null ? formatBRL(item.frete, true) : '—'}</td>
      <td class="num-col cell-cost-total">${formatBRL(item.custoTotal, true)}</td>
      <td class="num-col ${marginClass}">${formatBRL(item.margem, true)}</td>
      <td class="num-col">
        <span class="margin-pct-badge ${badgeClass}">${item.margemPct.toFixed(1).replace('.', ',')}%</span>
      </td>
      <td class="actions-col">
        <div class="row-actions">
          <button type="button" class="action-icon-btn" data-action="edit" data-id="${item.id}" title="Editar no Modo Cálculo">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
              <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
            </svg>
          </button>
          <button type="button" class="action-icon-btn" data-action="duplicate" data-id="${item.id}" title="Simular variação deste SKU">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
              <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
            </svg>
          </button>
          <button type="button" class="action-icon-btn delete-btn" data-action="delete" data-id="${item.id}" title="Excluir produto">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <polyline points="3 6 5 6 21 6"></polyline>
              <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
            </svg>
          </button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ============================================================================
// EXPORTAÇÃO E IMPORTAÇÃO DE DADOS (CSV / EXCEL / CLIPBOARD)
// ============================================================================

/**
 * Gera o CSV exatamente no formato original da planilha:
 * SKU,Preço Final,Preço Inicial,12% PF,16% PF,Embalagem,ERP Custo,Frete,Custo Total,Margem
 */
function exportStandardCSV() {
  if (products.length === 0) {
    showToast('Não há produtos para exportar.');
    return;
  }

  const headers = [
    'SKU',
    'Preço Final',
    'Preço Inicial',
    '12% PF',
    '16% PF',
    'Embalagem',
    'ERP Custo',
    'Frete',
    'Custo Total',
    'Margem'
  ];

  const rows = products.map((p) => [
    p.sku,
    formatRawCSVNumber(p.precoFinal),
    formatRawCSVNumber(p.precoInicial),
    formatRawCSVNumber(p.taxa12),
    formatRawCSVNumber(p.taxa16),
    formatRawCSVNumber(p.embalagem),
    formatRawCSVNumber(p.erpCusto),
    p.frete !== null ? formatRawCSVNumber(p.frete) : '',
    formatRawCSVNumber(p.custoTotal),
    formatRawCSVNumber(p.margem)
  ]);

  const csvContent = [headers.join(','), ...rows.map((r) => r.join(','))].join('\r\n');
  triggerDownload(csvContent, 'precificacao_produtos.csv', 'text/csv;charset=utf-8;');
  showToast('Planilha CSV padrão exportada com sucesso!');
}

/**
 * Gera CSV formatado para Excel em Português (separador ';' e decimal ',')
 */
function exportExcelPtBrCSV() {
  if (products.length === 0) {
    showToast('Não há produtos para exportar.');
    return;
  }

  const headers = [
    'SKU',
    'Preço Final',
    'Preço Inicial',
    '12% PF',
    '16% PF',
    'Embalagem',
    'ERP Custo',
    'Frete',
    'Custo Total',
    'Margem'
  ];

  const fmtPt = (v) => (v === null || v === undefined || v === '' ? '' : String(cleanFloat(v, 4)).replace('.', ','));

  const rows = products.map((p) => [
    p.sku,
    fmtPt(p.precoFinal),
    fmtPt(p.precoInicial),
    fmtPt(p.taxa12),
    fmtPt(p.taxa16),
    fmtPt(p.embalagem),
    fmtPt(p.erpCusto),
    p.frete !== null ? fmtPt(p.frete) : '',
    fmtPt(p.custoTotal),
    fmtPt(p.margem)
  ]);

  // BOM UTF-8 (\uFEFF) garante acentuação perfeita ao abrir direto no Excel
  const csvContent = '\uFEFF' + [headers.join(';'), ...rows.map((r) => r.join(';'))].join('\r\n');
  triggerDownload(csvContent, 'precificacao_produtos_excel_ptbr.csv', 'text/csv;charset=utf-8;');
  showToast('Arquivo para Excel (PT-BR) exportado com sucesso!');
}

/**
 * Copia a tabela em formato TSV para colar diretamente no Excel ou Google Sheets
 */
function copyTableToClipboard() {
  if (products.length === 0) {
    showToast('Não há produtos para copiar.');
    return;
  }

  const headers = [
    'SKU',
    'Preço Final',
    'Preço Inicial',
    '12% PF',
    '16% PF',
    'Embalagem',
    'ERP Custo',
    'Frete',
    'Custo Total',
    'Margem'
  ];

  const fmtPt = (v) => (v === null || v === undefined || v === '' ? '' : String(cleanFloat(v, 4)).replace('.', ','));

  const rows = products.map((p) => [
    p.sku,
    fmtPt(p.precoFinal),
    fmtPt(p.precoInicial),
    fmtPt(p.taxa12),
    fmtPt(p.taxa16),
    fmtPt(p.embalagem),
    fmtPt(p.erpCusto),
    p.frete !== null ? fmtPt(p.frete) : '',
    fmtPt(p.custoTotal),
    fmtPt(p.margem)
  ]);

  const tsv = [headers.join('\t'), ...rows.map((r) => r.join('\t'))].join('\n');
  navigator.clipboard.writeText(tsv).then(() => {
    showToast('Tabela copiada! Use Ctrl+V no Excel ou Google Planilhas.');
  });
}

function triggerDownload(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Converte texto bruto (CSV com ',', CSV PT-BR com ';', TSV copiado com '\t' ou JSON)
 * em uma lista de produtos validados.
 */
function parseProductsFromText(rawText) {
  const cleaned = String(rawText || '').replace(/^\uFEFF/, '').trim();
  if (!cleaned) return [];

  // Suporte caso o usuário cole ou importe um JSON
  if (cleaned.startsWith('[')) {
    try {
      const arr = JSON.parse(cleaned);
      if (Array.isArray(arr)) {
        return arr.map((item, idx) =>
          computeProductMetrics({
            id: 'imp-' + Date.now() + '-' + idx,
            sku: String(item.sku || ''),
            precoFinal: Number(item.precoFinal) || 0,
            precoInicial: Number(item.precoInicial) || 0,
            pct12: item.pct12 !== undefined ? Number(item.pct12) : 12,
            pct16: item.pct16 !== undefined ? Number(item.pct16) : 16,
            embalagem: item.embalagem !== undefined ? Number(item.embalagem) : 0.2,
            erpCusto: item.erpCusto !== undefined ? Number(item.erpCusto) : 2.15,
            frete: item.frete !== undefined && item.frete !== null && item.frete !== '' ? Number(item.frete) : null
          })
        );
      }
    } catch (_) {
      // Continua para o parser de linhas CSV/TSV
    }
  }

  const lines = cleaned.split(/\r?\n/).filter((l) => l.trim().length > 0);
  if (lines.length === 0) return [];

  // Detecta automaticamente o separador: Tab ('\t'), Ponto-e-vírgula (';') ou Vírgula (',')
  const firstLine = lines[0];
  let delimiter = ',';
  if (firstLine.includes('\t')) {
    delimiter = '\t';
  } else if (firstLine.includes(';')) {
    delimiter = ';';
  }

  // Verifica se a primeira linha é cabeçalho (contém "SKU" ou "Preço")
  const hasHeader = /sku|pre[cç]o/i.test(firstLine);
  const startIdx = hasHeader ? 1 : 0;

  const parseNum = (val) => {
    if (val === undefined || val === null) return null;
    const str = String(val).replace(/R\$/gi, '').trim();
    if (str === '' || str === '—' || str === '-') return null;
    // Se usa ',' como decimal (ex: Excel PT-BR ou TSV), converte para '.'
    const normalized = delimiter === ';' || delimiter === '\t' ? str.replace(/\./g, '').replace(',', '.') : str.replace(',', '.');
    const n = parseFloat(normalized);
    return isNaN(n) ? null : n;
  };

  const imported = [];
  for (let i = startIdx; i < lines.length; i++) {
    const cols = lines[i].split(delimiter);
    if (cols.length < 3) continue;

    const sku = cols[0].trim();
    const precoFinal = parseNum(cols[1]);
    const precoInicial = parseNum(cols[2]);
    if (!sku || precoFinal === null) continue;

    const embalagem = cols[5] !== undefined && parseNum(cols[5]) !== null ? parseNum(cols[5]) : 0.2;
    const erpCusto = cols[6] !== undefined && parseNum(cols[6]) !== null ? parseNum(cols[6]) : 2.15;
    const frete = cols[7] !== undefined ? parseNum(cols[7]) : null;

    imported.push(
      computeProductMetrics({
        id: 'imp-' + Date.now() + '-' + i + '-' + Math.random().toString(36).substring(2, 5),
        sku,
        precoFinal,
        precoInicial: precoInicial || 0,
        pct12: 12,
        pct16: 16,
        embalagem,
        erpCusto,
        frete
      })
    );
  }

  return imported;
}

function getSelectedImportMode() {
  const checked = document.querySelector('input[name="import-mode"]:checked');
  return checked ? checked.value : 'merge';
}

function applyImportedProducts(importedList) {
  if (!importedList || importedList.length === 0) {
    showToast('Não foi possível identificar produtos válidos nos dados enviados.');
    return false;
  }

  const mode = getSelectedImportMode();
  if (mode === 'replace') {
    products = importedList;
    showToast(`Lista substituída por ${importedList.length} produtos importados!`);
  } else {
    // Mesclar: adiciona novos itens evitando duplicatas exatas (mesmo SKU + mesmo Preço Final)
    let addedCount = 0;
    importedList.forEach((imp) => {
      const exists = products.some(
        (p) => String(p.sku) === String(imp.sku) && Math.abs(p.precoFinal - imp.precoFinal) < 0.001
      );
      if (!exists) {
        products.push(imp);
        addedCount++;
      }
    });
    showToast(`${addedCount} novos produtos adicionados (${importedList.length - addedCount} já existiam)!`);
  }

  saveProductsToStorage();
  renderHistoryTable();
  renderCatalogKPIs();
  closeImportModal();
  switchTab('history');
  return true;
}

function handleImportFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    const text = String(e.target?.result || '');
    const imported = parseProductsFromText(text);
    applyImportedProducts(imported);
  };
  reader.readAsText(file);
}

function openImportModal() {
  const modal = document.getElementById('import-modal');
  if (modal) {
    modal.classList.remove('hidden');
  }
}

function closeImportModal() {
  const modal = document.getElementById('import-modal');
  if (modal) {
    modal.classList.add('hidden');
  }
  const textarea = document.getElementById('textarea-import-paste');
  if (textarea) textarea.value = '';
}

// ============================================================================
// TOAST DE FEEDBACK
// ============================================================================

function showToast(message) {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.textContent = message;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.25s ease';
    setTimeout(() => toast.remove(), 260);
  }, 3200);
}

// ============================================================================
// INICIALIZAÇÃO DE EVENTOS
// ============================================================================

document.addEventListener('DOMContentLoaded', () => {
  loadProductsFromStorage();

  // Carrega o primeiro item da planilha (SKU 322) como exemplo inicial na calculadora
  if (products.length > 0) {
    const first = products[0];
    document.getElementById('input-sku').value = first.sku;
    document.getElementById('input-preco-inicial').value = first.precoInicial;
    document.getElementById('input-preco-final').value = first.precoFinal;
    document.getElementById('input-embalagem').value = first.embalagem;
    document.getElementById('input-erp').value = first.erpCusto;
    document.getElementById('input-frete').value = first.frete !== null ? first.frete : '';
  }

  updateLiveCalculation();
  renderHistoryTable();
  renderCatalogKPIs();

  // Eventos das Guias (Tabs)
  document.getElementById('tab-btn-calc').addEventListener('click', () => switchTab('calc'));
  document.getElementById('tab-btn-history').addEventListener('click', () => switchTab('history'));
  document.getElementById('btn-go-to-calc').addEventListener('click', () => switchTab('calc'));

  // Eventos dos Inputs para Cálculo Dinâmico em Tempo Real
  const liveInputs = [
    'input-sku',
    'input-preco-inicial',
    'input-preco-final',
    'input-taxa-12',
    'input-taxa-16',
    'input-embalagem',
    'input-erp',
    'input-frete'
  ];

  liveInputs.forEach((id) => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', updateLiveCalculation);
    }
  });

  // Atalhos rápidos (Chips de Embalagem, ERP e Frete)
  document.querySelectorAll('.chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      const targetId = chip.getAttribute('data-target');
      const value = chip.getAttribute('data-value');
      const inputEl = document.getElementById(targetId);
      if (inputEl) {
        inputEl.value = value;
        updateLiveCalculation();
      }
    });
  });

  // Simulador Reverso: Botões de % de Margem
  document.querySelectorAll('.btn-preset').forEach((btn) => {
    btn.addEventListener('click', () => {
      const targetPct = parseFloat(btn.getAttribute('data-target-margin-pct'));
      const m = getFormValues();
      const fixedCosts = m.precoInicial + m.embalagem + m.erpCusto + (m.frete || 0);
      const denominator = 1 - (m.pct12 + m.pct16 + targetPct) / 100;

      if (denominator > 0 && fixedCosts > 0) {
        const idealPrice = cleanFloat(fixedCosts / denominator, 2);
        document.getElementById('input-preco-final').value = idealPrice.toFixed(2);
        updateLiveCalculation();
        showToast(`Preço Final ajustado para R$ ${idealPrice.toFixed(2).replace('.', ',')} (Margem de ${targetPct}%)`);
      }
    });
  });

  // Simulador Reverso: Meta em R$
  document.getElementById('btn-apply-target-brl').addEventListener('click', () => {
    const targetBrl = parseFloat(document.getElementById('input-target-brl').value);
    if (isNaN(targetBrl)) {
      showToast('Digite um valor de margem em R$ para calcular.');
      return;
    }
    const m = getFormValues();
    const fixedCosts = m.precoInicial + m.embalagem + m.erpCusto + (m.frete || 0);
    const denominator = 1 - (m.pct12 + m.pct16) / 100;

    if (denominator > 0) {
      const idealPrice = cleanFloat((fixedCosts + targetBrl) / denominator, 2);
      document.getElementById('input-preco-final').value = idealPrice.toFixed(2);
      updateLiveCalculation();
      showToast(`Preço Final ajustado para R$ ${idealPrice.toFixed(2).replace('.', ',')} (Lucro de ${formatBRL(targetBrl)})`);
    }
  });

  // Submissão do formulário, Reset e Cancelar Edição
  document.getElementById('pricing-form').addEventListener('submit', handleFormSubmit);
  document.getElementById('btn-reset-form').addEventListener('click', resetCalculatorForm);
  document.getElementById('btn-cancel-edit').addEventListener('click', () => {
    exitEditMode();
    showToast('Modo de edição cancelado.');
  });

  // Ações nas linhas da tabela do Histórico (Delegação de eventos)
  document.getElementById('products-tbody').addEventListener('click', (e) => {
    const btn = e.target.closest('.action-icon-btn');
    if (!btn) return;
    const action = btn.getAttribute('data-action');
    const id = btn.getAttribute('data-id');

    if (action === 'edit') {
      loadProductIntoCalculator(id, false);
    } else if (action === 'duplicate') {
      loadProductIntoCalculator(id, true);
    } else if (action === 'delete') {
      deleteProduct(id);
    }
  });

  // Busca, Filtro e Ordenação no Histórico
  document.getElementById('search-sku').addEventListener('input', renderHistoryTable);
  document.getElementById('filter-margin').addEventListener('change', renderHistoryTable);

  document.querySelectorAll('#products-table th.sortable').forEach((th) => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-sort');
      if (sortState.column === col) {
        sortState.direction = sortState.direction === 'asc' ? 'desc' : 'asc';
      } else {
        sortState.column = col;
        sortState.direction = 'asc';
      }

      document.querySelectorAll('#products-table th.sortable .sort-indicator').forEach((el) => {
        el.textContent = '';
      });
      th.querySelector('.sort-indicator').textContent = sortState.direction === 'asc' ? ' ▲' : ' ▼';

      renderHistoryTable();
    });
  });

  // Botões de Exportação, Cópia, Importação e Restauração
  document.getElementById('btn-quick-export').addEventListener('click', exportStandardCSV);
  document.getElementById('btn-export-csv').addEventListener('click', exportStandardCSV);
  document.getElementById('btn-export-excel-csv').addEventListener('click', exportExcelPtBrCSV);
  document.getElementById('btn-copy-table').addEventListener('click', copyTableToClipboard);

  // Eventos do Modal de Importação / Transferência entre Aparelhos
  document.getElementById('btn-open-import-modal')?.addEventListener('click', openImportModal);
  document.getElementById('btn-history-import')?.addEventListener('click', openImportModal);
  document.getElementById('btn-close-import-modal')?.addEventListener('click', closeImportModal);

  document.getElementById('import-modal')?.addEventListener('click', (e) => {
    if (e.target.id === 'import-modal') closeImportModal();
  });

  // Alternância visual dos cards de modo (Mesclar vs Substituir)
  document.querySelectorAll('input[name="import-mode"]').forEach((radio) => {
    radio.addEventListener('change', () => {
      document.getElementById('label-mode-merge')?.classList.toggle('active', radio.value === 'merge' && radio.checked);
      document.getElementById('label-mode-replace')?.classList.toggle('active', radio.value === 'replace' && radio.checked);
    });
  });

  // Seleção e Drag-and-Drop de Arquivo
  const fileInput = document.getElementById('input-import-csv');
  const dropzone = document.getElementById('import-dropzone');

  document.getElementById('btn-trigger-file')?.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput?.click();
  });

  dropzone?.addEventListener('click', () => fileInput?.click());

  dropzone?.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('drag-over');
  });

  dropzone?.addEventListener('dragleave', () => {
    dropzone.classList.remove('drag-over');
  });

  dropzone?.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
    const file = e.dataTransfer?.files?.[0];
    if (file) handleImportFile(file);
  });

  fileInput?.addEventListener('change', (e) => {
    const file = e.target.files?.[0];
    if (file) handleImportFile(file);
    e.target.value = '';
  });

  // Importar a partir de texto colado
  document.getElementById('btn-import-from-text')?.addEventListener('click', () => {
    const raw = document.getElementById('textarea-import-paste')?.value || '';
    if (!raw.trim()) {
      showToast('Cole o conteúdo do CSV ou da tabela antes de importar.');
      return;
    }
    const imported = parseProductsFromText(raw);
    applyImportedProducts(imported);
  });

  document.getElementById('btn-restore-defaults').addEventListener('click', () => {
    products = INITIAL_PRODUCTS.map(computeProductMetrics);
    saveProductsToStorage();
    renderHistoryTable();
    renderCatalogKPIs();
    showToast('Planilha original com os 10 produtos restaurada!');
  });
});
