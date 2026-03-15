/**
 * サロン家計簿 - メインアプリケーション
 * Version: 1.5.0 (GAS URL定数埋め込み・設定UI簡略化)
 *
 * ─ アーキテクチャ概要 ─────────────────────────
 *  GasAPI    … Google Apps Script Web API との HTTP 通信
 *  Storage   … localStorage の読み書き（GAS取得後のキャッシュ）
 *  Data      … CRUD 操作（GAS優先・成功後にlocalStorage更新）
 *  Templates … テンプレート管理（localStorage のみ）
 *  Calculator … 残高・集計ロジック
 *  Format    … 表示用フォーマット変換
 *  Charts    … Chart.js グラフ描画
 *  UI        … タブ切り替え・画面描画・イベント処理・オフライン検知
 * ─────────────────────────────────────────────
 */

'use strict';

/* =============================================
   定数・マスターデータ
   ============================================= */

const APP_VERSION   = '1.5.0';
const STORAGE_KEY   = 'salon_kaikei_v1_transactions';
const SETTINGS_KEY  = 'salon_kaikei_v1_settings';
const TEMPLATES_KEY = 'salon_kaikei_v1_templates';

const GAS_URL = 'https://script.google.com/macros/s/AKfycbwPJ4vodIXXLiozy8HnytXyrSh6PpuqvKCUGpFS-I_Jm-e2HdWVtl_UuukUOeFtCwif/exec';

/** 勘定科目マスター（青色申告決算書準拠） */
const CATEGORIES = {
  income:  ['売上'],
  expense: [
    '材料費',
    '消耗品費',
    '広告宣伝費',
    '水道光熱費',
    '通信費',
    '地代家賃',
    '外注費',
    '接待交際費',
    '雑費',
  ],
};

/** グラフ用カラー */
const CHART_COLORS = {
  income:    '#10B981',
  expense:   '#EF4444',
  incomeAlpha:  'rgba(16, 185, 129, 0.72)',
  expenseAlpha: 'rgba(239, 68,  68,  0.72)',
  categories: [
    '#F472B6', '#A78BFA', '#60A5FA', '#34D399',
    '#FBBF24', '#F87171', '#818CF8', '#2DD4BF',
    '#FB923C', '#4ADE80',
  ],
};

/* =============================================
   GasAPI モジュール
   Google Apps Script Web アプリとの HTTP 通信
   ============================================= */
const GasAPI = {
  isConfigured() {
    return typeof GAS_URL === 'string' && GAS_URL.length > 0 && !GAS_URL.includes('★');
  },

  async _post(body) {
    const res = await fetch(GAS_URL, {
      method:   'POST',
      headers:  { 'Content-Type': 'text/plain;charset=UTF-8' },
      body:     JSON.stringify(body),
      redirect: 'follow',
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (json.error) throw new Error(json.error);
    return json;
  },

  async fetchAll() {
    const res = await fetch(`${GAS_URL}?action=getAll`, { redirect: 'follow' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (json.error) throw new Error(json.error);
    return json;
  },

  async addTransaction(t) {
    return this._post({ action: 'addTransaction', data: t });
  },

  async updateTransaction(t) {
    return this._post({ action: 'updateTransaction', data: t });
  },

  async deleteTransaction(id) {
    return this._post({ action: 'deleteTransaction', id });
  },

  async saveSettings(s) {
    return this._post({ action: 'saveSettings', data: s });
  },

};

/* =============================================
   Storage モジュール
   localStorage の読み書き（GAS キャッシュ兼用）
   ============================================= */
const Storage = {
  getTransactions() {
    try {
      return JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    } catch {
      return [];
    }
  },

  saveTransactions(list) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
  },

  getSettings() {
    const defaults = { businessName: 'マイサロン', initialCash: 0, initialBank: 0 };
    try {
      return { ...defaults, ...JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}') };
    } catch {
      return defaults;
    }
  },

  saveSettings(obj) {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(obj));
  },

  async loadFromGas() {
    try {
      const data = await GasAPI.fetchAll();
      if (Array.isArray(data.transactions)) {
        this.saveTransactions(data.transactions);
      }
      if (data.settings && typeof data.settings === 'object') {
        this.saveSettings({ ...this.getSettings(), ...data.settings });
      }
      return true;
    } catch (e) {
      console.warn('[Storage] GASからのデータ取得に失敗（ローカルキャッシュを使用）:', e.message);
      return false;
    }
  },
};

/* =============================================
   Data モジュール（CRUD）
   ============================================= */
const Data = {
  /** 取引を追加（GAS優先・成功後にlocalStorageを更新） */
  async add(transaction) {
    if (!GasAPI.isConfigured()) throw new Error('GAS URLが設定されていません。');
    const record = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`,
      ...transaction,
      createdAt: new Date().toISOString(),
    };
    await GasAPI.addTransaction(record); // 失敗時は例外をスロー
    const list = Storage.getTransactions();
    list.unshift(record);
    Storage.saveTransactions(list);
    return record;
  },

  /** ID で更新（GAS優先・成功後にlocalStorageを更新） */
  async update(updated) {
    if (!GasAPI.isConfigured()) throw new Error('GAS URLが設定されていません。');
    const list = Storage.getTransactions();
    const idx  = list.findIndex(t => t.id === updated.id);
    if (idx === -1) throw new Error('取引が見つかりません');
    const merged = { ...list[idx], ...updated, updatedAt: new Date().toISOString() };
    await GasAPI.updateTransaction(merged); // 失敗時は例外をスロー
    list[idx] = merged;
    Storage.saveTransactions(list);
    return merged;
  },

  /** ID で削除（GAS優先・成功後にlocalStorageを更新） */
  async remove(id) {
    if (!GasAPI.isConfigured()) throw new Error('GAS URLが設定されていません。');
    await GasAPI.deleteTransaction(id); // 失敗時は例外をスロー
    const list = Storage.getTransactions().filter(t => t.id !== id);
    Storage.saveTransactions(list);
  },

  /** 全件取得 */
  getAll() {
    return Storage.getTransactions();
  },

  /** 年月で絞り込み */
  getByYearMonth(year, month) {
    return this.getAll().filter(t => {
      const d = new Date(t.date);
      return d.getFullYear() === year && (d.getMonth() + 1) === month;
    });
  },

  /** 年で絞り込み */
  getByYear(year) {
    return this.getAll().filter(t => new Date(t.date).getFullYear() === year);
  },

  /** ID で1件取得 */
  getById(id) {
    return this.getAll().find(t => t.id === id) || null;
  },
};

/* =============================================
   Templates モジュール
   よく使う取引テンプレートの管理（localStorage のみ）
   ============================================= */
const Templates = {
  getAll() {
    try {
      return JSON.parse(localStorage.getItem(TEMPLATES_KEY) || '[]');
    } catch {
      return [];
    }
  },

  _save(list) {
    localStorage.setItem(TEMPLATES_KEY, JSON.stringify(list));
  },

  /** テンプレートを追加 */
  add(tpl) {
    const list = this.getAll();
    const record = {
      id: `tpl-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      name:          tpl.name          || tpl.description || '無題',
      type:          tpl.type          || 'expense',
      amount:        Number(tpl.amount) || 0,
      category:      tpl.category      || '',
      description:   tpl.description   || '',
      paymentMethod: tpl.paymentMethod || 'cash',
      createdAt:     new Date().toISOString(),
    };
    list.push(record);
    this._save(list);
    return record;
  },

  /** ID でテンプレートを削除 */
  remove(id) {
    this._save(this.getAll().filter(t => t.id !== id));
  },
};

/* =============================================
   Calculator モジュール
   ============================================= */
const Calculator = {
  getBalances() {
    const { initialCash, initialBank } = Storage.getSettings();
    let cash = Number(initialCash) || 0;
    let bank = Number(initialBank) || 0;

    Data.getAll().forEach(t => {
      const amt = Number(t.amount);
      if (t.type === 'income') {
        t.paymentMethod === 'cash' ? (cash += amt) : (bank += amt);
      } else {
        t.paymentMethod === 'cash' ? (cash -= amt) : (bank -= amt);
      }
    });
    return { cash, bank };
  },

  getCurrentMonthSummary() {
    const now = new Date();
    return this._sumTransactions(
      Data.getByYearMonth(now.getFullYear(), now.getMonth() + 1)
    );
  },

  getAnnualSummary(year) {
    return this._sumTransactions(Data.getByYear(year));
  },

  getMonthlySummary(year) {
    return Array.from({ length: 12 }, (_, i) => {
      const m = i + 1;
      const s = this._sumTransactions(Data.getByYearMonth(year, m));
      return { month: m, ...s };
    });
  },

  getCategorySummary(year) {
    const expenses = Data.getByYear(year).filter(t => t.type === 'expense');
    const map = {};
    expenses.forEach(t => {
      map[t.category] = (map[t.category] || 0) + Number(t.amount);
    });
    const total = Object.values(map).reduce((a, b) => a + b, 0);
    return Object.entries(map)
      .sort(([, a], [, b]) => b - a)
      .map(([category, amount]) => ({
        category,
        amount,
        ratio: total > 0 ? ((amount / total) * 100).toFixed(1) : '0.0',
      }));
  },

  getAvailableYears() {
    const years = new Set(
      Data.getAll().map(t => new Date(t.date).getFullYear())
    );
    years.add(new Date().getFullYear());
    return [...years].sort((a, b) => b - a);
  },

  _sumTransactions(list) {
    const income  = list.filter(t => t.type === 'income')
                        .reduce((s, t) => s + Number(t.amount), 0);
    const expense = list.filter(t => t.type === 'expense')
                        .reduce((s, t) => s + Number(t.amount), 0);
    return { income, expense, profit: income - expense };
  },
};

/* =============================================
   Format ユーティリティ
   ============================================= */
const Format = {
  currency(amount) {
    return '¥' + Math.abs(Number(amount)).toLocaleString('ja-JP');
  },

  date(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${y}/${m}/${dd}`;
  },
};

/* =============================================
   Charts モジュール
   ============================================= */
const Charts = {
  _monthly:  null,
  _category: null,

  renderMonthly() {
    const ctx = document.getElementById('monthlyChart');
    if (!ctx) return;
    if (this._monthly) { this._monthly.destroy(); this._monthly = null; }

    const year = new Date().getFullYear();
    const data = Calculator.getMonthlySummary(year);

    this._monthly = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: data.map(d => `${d.month}月`),
        datasets: [
          {
            label: '売上',
            data: data.map(d => d.income),
            backgroundColor: CHART_COLORS.incomeAlpha,
            borderColor: CHART_COLORS.income,
            borderWidth: 1,
            borderRadius: 4,
          },
          {
            label: '経費',
            data: data.map(d => d.expense),
            backgroundColor: CHART_COLORS.expenseAlpha,
            borderColor: CHART_COLORS.expense,
            borderWidth: 1,
            borderRadius: 4,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top', labels: { font: { size: 11 }, boxWidth: 14 } },
          tooltip: {
            callbacks: {
              label: ctx => ` ${ctx.dataset.label}：¥${ctx.raw.toLocaleString()}`,
            },
          },
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: {
              font: { size: 10 },
              callback: v => v === 0 ? '0' : `¥${(v / 10000).toFixed(0)}万`,
            },
            grid: { color: '#F3F4F6' },
          },
          x: { ticks: { font: { size: 10 } }, grid: { display: false } },
        },
      },
    });
  },

  renderCategory() {
    const ctx = document.getElementById('categoryChart');
    if (!ctx) return;
    if (this._category) { this._category.destroy(); this._category = null; }

    const now = new Date();
    const expenses = Data.getByYearMonth(now.getFullYear(), now.getMonth() + 1)
                        .filter(t => t.type === 'expense');
    const map = {};
    expenses.forEach(t => {
      map[t.category] = (map[t.category] || 0) + Number(t.amount);
    });
    const entries = Object.entries(map);

    if (entries.length === 0) {
      this._category = new Chart(ctx, {
        type: 'doughnut',
        data: {
          labels: ['経費の記録なし'],
          datasets: [{ data: [1], backgroundColor: ['#E5E7EB'], borderWidth: 0 }],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false }, tooltip: { enabled: false } },
        },
      });
      return;
    }

    this._category = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels: entries.map(([k]) => k),
        datasets: [{
          data: entries.map(([, v]) => v),
          backgroundColor: CHART_COLORS.categories,
          borderWidth: 2,
          borderColor: 'white',
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'right',
            labels: { font: { size: 10 }, padding: 8, boxWidth: 11 },
          },
          tooltip: {
            callbacks: {
              label: ctx => ` ${ctx.label}：¥${ctx.raw.toLocaleString()}`,
            },
          },
        },
      },
    });
  },
};

/* =============================================
   UI モジュール
   ============================================= */
const UI = {
  currentType:  'income',
  _pendingDeleteId: null,

  /** 編集モード状態 */
  _editMode:   false,
  _editingId:  null,

  /* ─── タブ ─────────────────────────────── */

  initTabs() {
    document.querySelectorAll('.nav-btn').forEach(btn => {
      btn.addEventListener('click', () => this.showTab(btn.dataset.tab));
    });
  },

  showTab(name) {
    document.querySelectorAll('.nav-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.tab === name);
    });
    document.querySelectorAll('.tab-content').forEach(s => {
      s.classList.toggle('active', s.id === `tab-${name}`);
    });

    if (name === 'dashboard') this.renderDashboard();
    if (name === 'list')      this.renderList();
    if (name === 'summary')   this.renderSummary();
    if (name === 'settings')  this.renderSettings();
  },

  /* ─── ダッシュボード ─────────────────── */

  renderDashboard() {
    const { cash, bank } = Calculator.getBalances();
    const { income, expense } = Calculator.getCurrentMonthSummary();

    document.getElementById('cashBalance').textContent    = Format.currency(cash);
    document.getElementById('bankBalance').textContent    = Format.currency(bank);
    document.getElementById('monthlyIncome').textContent  = Format.currency(income);
    document.getElementById('monthlyExpense').textContent = Format.currency(expense);

    const listEl = document.getElementById('recentList');
    listEl.innerHTML = '';
    const recent = Data.getAll().slice(0, 5);
    if (recent.length === 0) {
      listEl.innerHTML = '<p style="text-align:center;color:#9CA3AF;padding:24px 0">まだ記録がありません</p>';
    } else {
      recent.forEach(t => listEl.appendChild(this._buildItem(t, false)));
    }

    Charts.renderMonthly();
    Charts.renderCategory();
  },

  /* ─── 収支一覧 ────────────────────────── */

  renderList() {
    this._rebuildFilterOptions();
    this._applyFilters();
  },

  _rebuildFilterOptions() {
    const months = new Set();
    Data.getAll().forEach(t => {
      const d = new Date(t.date);
      months.add(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
    });
    const monthSel = document.getElementById('filterMonth');
    const prevMonth = monthSel.value;
    monthSel.innerHTML = '<option value="">全期間</option>';
    [...months].sort().reverse().forEach(key => {
      const [y, m] = key.split('-');
      const opt = document.createElement('option');
      opt.value = key;
      opt.textContent = `${y}年 ${parseInt(m)}月`;
      monthSel.appendChild(opt);
    });
    monthSel.value = prevMonth;

    const catSel = document.getElementById('filterCategory');
    const prevCat = catSel.value;
    catSel.innerHTML = '<option value="">科目すべて</option>';
    [...CATEGORIES.income, ...CATEGORIES.expense].forEach(cat => {
      const opt = document.createElement('option');
      opt.value = cat;
      opt.textContent = cat;
      catSel.appendChild(opt);
    });
    catSel.value = prevCat;
  },

  _applyFilters() {
    let list = Data.getAll();
    const month    = document.getElementById('filterMonth').value;
    const type     = document.getElementById('filterType').value;
    const category = document.getElementById('filterCategory').value;

    if (month) {
      const [y, m] = month.split('-').map(Number);
      list = list.filter(t => {
        const d = new Date(t.date);
        return d.getFullYear() === y && (d.getMonth() + 1) === m;
      });
    }
    if (type)     list = list.filter(t => t.type === type);
    if (category) list = list.filter(t => t.category === category);

    const listEl  = document.getElementById('transactionList');
    const emptyEl = document.getElementById('listEmpty');
    listEl.innerHTML = '';

    if (list.length === 0) {
      emptyEl.style.display = 'block';
    } else {
      emptyEl.style.display = 'none';
      list.forEach(t => listEl.appendChild(this._buildItem(t, true)));
    }
  },

  /* ─── 取引アイテム DOM 生成 ──────────── */

  _buildItem(t, showActions) {
    const div = document.createElement('div');
    div.className = 'transaction-item';

    const payTag = t.paymentMethod === 'cash'
      ? '<span class="tag cash">💴 現金</span>'
      : '<span class="tag card">💳 カード</span>';
    const sign = t.type === 'income' ? '+' : '−';

    const actionBtns = showActions ? `
      <button class="edit-btn" title="編集" style="
        background:none;border:1px solid #D1D5DB;border-radius:6px;
        padding:4px 10px;font-size:12px;cursor:pointer;color:#374151;
        margin-right:4px;white-space:nowrap;
      ">編集</button>
      <button class="delete-btn" title="削除" aria-label="${this._esc(t.description)}を削除">✕</button>
    ` : '';

    div.innerHTML = `
      <div class="transaction-type-dot ${t.type}" aria-hidden="true"></div>
      <div class="transaction-info">
        <div class="transaction-desc">${this._esc(t.description)}</div>
        <div class="transaction-meta">
          <span>${Format.date(t.date)}</span>
          <span class="tag">${this._esc(t.category)}</span>
          ${payTag}
        </div>
      </div>
      <div class="transaction-amount ${t.type}" aria-label="${t.type === 'income' ? '収入' : '支出'} ${Format.currency(t.amount)}">
        ${sign}${Format.currency(t.amount)}
      </div>
      ${actionBtns}
    `;

    if (showActions) {
      const editBtn = div.querySelector('.edit-btn');
      const delBtn  = div.querySelector('.delete-btn');
      if (!navigator.onLine) {
        editBtn.disabled = true;
        editBtn.style.opacity = '0.4';
        editBtn.style.cursor  = 'not-allowed';
        delBtn.disabled  = true;
        delBtn.style.opacity  = '0.4';
        delBtn.style.cursor   = 'not-allowed';
      }
      editBtn.addEventListener('click', () => { this._loadForEdit(t.id); });
      delBtn.addEventListener('click', () => { this._openDeleteModal(t.id, t.description); });
    }
    return div;
  },

  /* ─── 集計タブ ────────────────────────── */

  renderSummary() {
    const years = Calculator.getAvailableYears();
    const sel = document.getElementById('summaryYear');
    const prev = sel.value || String(new Date().getFullYear());
    sel.innerHTML = '';
    years.forEach(y => {
      const opt = document.createElement('option');
      opt.value = y;
      opt.textContent = `${y}年`;
      sel.appendChild(opt);
    });
    sel.value = prev || String(years[0]);

    // CSV エクスポートセクション注入（初回のみ）
    this._injectCsvSection();

    this._updateSummary(Number(sel.value));
  },

  _updateSummary(year) {
    const ann = Calculator.getAnnualSummary(year);
    document.getElementById('annualIncome').textContent  = Format.currency(ann.income);
    document.getElementById('annualExpense').textContent = Format.currency(ann.expense);
    const profEl = document.getElementById('annualProfit');
    profEl.textContent = Format.currency(ann.profit);
    profEl.className = `card-value ${ann.profit >= 0 ? 'income-text' : 'expense-text'}`;

    const monthly = Calculator.getMonthlySummary(year);
    const tbody = document.getElementById('monthlyTableBody');
    tbody.innerHTML = '';
    monthly.forEach(m => {
      const tr = document.createElement('tr');
      const noData = m.income === 0 && m.expense === 0;
      const cls    = m.profit >= 0 ? 'profit-positive' : 'profit-negative';
      const sign   = m.profit >= 0 ? '+' : '−';
      tr.innerHTML = `
        <td>${m.month}月</td>
        <td>${m.income  > 0 ? Format.currency(m.income)  : '−'}</td>
        <td>${m.expense > 0 ? Format.currency(m.expense) : '−'}</td>
        <td class="${noData ? '' : cls}">
          ${noData ? '−' : sign + Format.currency(m.profit)}
        </td>
      `;
      tbody.appendChild(tr);
    });

    const cats = Calculator.getCategorySummary(year);
    const cTbody = document.getElementById('categoryTableBody');
    cTbody.innerHTML = '';
    if (cats.length === 0) {
      cTbody.innerHTML = `
        <tr>
          <td colspan="3" style="text-align:center;color:#9CA3AF;padding:20px">
            経費の記録がありません
          </td>
        </tr>`;
    } else {
      cats.forEach(c => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${this._esc(c.category)}</td>
          <td>${Format.currency(c.amount)}</td>
          <td>${c.ratio}%</td>
        `;
        cTbody.appendChild(tr);
      });
    }

    // CSV 月セレクターの年を更新
    const csvMonth = document.getElementById('csvExportMonth');
    if (csvMonth) csvMonth.dataset.year = String(year);
  },

  /* ─── CSV エクスポート ────────────────── */

  _injectCsvSection() {
    if (document.getElementById('csvExportSection')) return;
    const summaryTab = document.getElementById('tab-summary');
    if (!summaryTab) return;

    const section = document.createElement('div');
    section.id        = 'csvExportSection';
    section.className = 'card';
    section.style.cssText = 'margin-top:16px;';
    section.innerHTML = `
      <h3 class="card-title">📥 CSVエクスポート（青色申告用）</h3>
      <p style="font-size:13px;color:#6B7280;margin-bottom:12px;line-height:1.6">
        Excelで開ける形式（UTF-8 BOM付き）でダウンロードできます。
      </p>
      <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
        <select id="csvExportMonth" class="form-input" style="flex:1;min-width:120px;max-width:160px;">
          <option value="">年間（全月）</option>
          <option value="1">1月</option>
          <option value="2">2月</option>
          <option value="3">3月</option>
          <option value="4">4月</option>
          <option value="5">5月</option>
          <option value="6">6月</option>
          <option value="7">7月</option>
          <option value="8">8月</option>
          <option value="9">9月</option>
          <option value="10">10月</option>
          <option value="11">11月</option>
          <option value="12">12月</option>
        </select>
        <button type="button" id="exportCsvBtn" class="btn btn-primary" style="flex:1;min-width:140px;">
          CSVダウンロード
        </button>
      </div>
    `;
    summaryTab.appendChild(section);

    document.getElementById('exportCsvBtn').addEventListener('click', () => {
      const yearSel  = document.getElementById('summaryYear');
      const monthSel = document.getElementById('csvExportMonth');
      const year  = Number(yearSel.value);
      const month = monthSel.value ? Number(monthSel.value) : null;
      this._exportCsv(year, month);
    });
  },

  _exportCsv(year, month) {
    const list = month
      ? Data.getByYearMonth(year, month)
      : Data.getByYear(year);

    // 日付順にソート（古い順）
    const sorted = [...list].sort((a, b) => a.date.localeCompare(b.date));

    const typeLabel = t => t.type === 'income' ? '収入' : '支出';
    const payLabel  = t => t.paymentMethod === 'cash' ? '現金' : 'カード';

    const header = '日付,種別,勘定科目,内容,金額,支払方法';
    const rows   = sorted.map(t =>
      [
        t.date,
        typeLabel(t),
        `"${t.category}"`,
        `"${String(t.description).replace(/"/g, '""')}"`,
        t.amount,
        payLabel(t),
      ].join(',')
    );

    // UTF-8 BOM 付き（Excel 文字化け防止）
    const bom  = '\uFEFF';
    const csv  = bom + [header, ...rows].join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    const suffix = month ? `${year}年${month}月` : `${year}年_年間`;
    a.download = `salon_kaikei_${suffix}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  },

  /* ─── 設定タブ ────────────────────────── */

  renderSettings() {
    const s = Storage.getSettings();
    document.getElementById('businessName').value = s.businessName || '';
    document.getElementById('initialCash').value  = s.initialCash  || '';
    document.getElementById('initialBank').value  = s.initialBank  || '';
  },

  /* ─── 入力フォーム ───────────────────── */

  initForm() {
    document.getElementById('inputDate').value = this._todayStr();

    document.querySelectorAll('.type-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        this.currentType = btn.dataset.type;
        document.querySelectorAll('.type-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        this._rebuildCategoryOptions();
      });
    });

    document.querySelectorAll('input[name="paymentMethod"]').forEach(radio => {
      radio.addEventListener('change', () => this._updatePaymentStyle());
    });
    this._updatePaymentStyle();
    this._rebuildCategoryOptions();

    document.getElementById('recordForm').addEventListener('submit', e => {
      e.preventDefault();
      this._submitRecord();
    });
  },

  /** テンプレート & 編集バナーを記録タブに注入（初回のみ） */
  initTemplates() {
    const recordTab = document.getElementById('tab-record');
    if (!recordTab || document.getElementById('editModeBanner')) return;

    // 1) 編集モードバナー
    const banner = document.createElement('div');
    banner.id = 'editModeBanner';
    banner.style.cssText = [
      'display:none',
      'background:#FEF3C7',
      'border:1px solid #FCD34D',
      'border-radius:8px',
      'padding:10px 14px',
      'margin-bottom:12px',
      'display:none',
      'align-items:center',
      'justify-content:space-between',
      'gap:8px',
      'flex-wrap:wrap',
    ].join(';');
    banner.innerHTML = `
      <span id="editModeBannerText" style="font-size:13px;font-weight:600;color:#92400E;">
        ✏️ 編集モード
      </span>
      <button type="button" id="cancelEditBtn" class="btn btn-secondary"
        style="padding:4px 14px;font-size:12px;">
        キャンセル
      </button>
    `;
    recordTab.insertBefore(banner, recordTab.firstChild);

    document.getElementById('cancelEditBtn').addEventListener('click', () => {
      this._cancelEdit();
    });

    // 2) テンプレートセクション
    const tplSection = document.createElement('div');
    tplSection.id        = 'templateSection';
    tplSection.className = 'card';
    tplSection.style.cssText = 'margin-bottom:12px;';
    tplSection.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center;cursor:pointer;"
           id="templateToggleRow">
        <h3 class="card-title" style="margin:0;">📋 テンプレートから入力</h3>
        <span id="templateToggleIcon" style="font-size:12px;color:#6B7280;">▼ 開く</span>
      </div>
      <div id="templateListArea" style="display:none;margin-top:12px;">
        <div id="templateListItems"></div>
        <button type="button" id="saveAsTemplateBtn" class="btn btn-secondary"
          style="width:100%;margin-top:8px;font-size:13px;">
          📌 現在の入力をテンプレートに保存
        </button>
        <p id="templateMsg" class="form-message" style="display:none;margin-top:6px;"></p>
      </div>
    `;

    const form = document.getElementById('recordForm');
    recordTab.insertBefore(tplSection, form);

    // テンプレート開閉
    document.getElementById('templateToggleRow').addEventListener('click', () => {
      const area = document.getElementById('templateListArea');
      const icon = document.getElementById('templateToggleIcon');
      const open = area.style.display === 'none';
      area.style.display = open ? 'block' : 'none';
      icon.textContent   = open ? '▲ 閉じる' : '▼ 開く';
      if (open) this._renderTemplateList();
    });

    // テンプレートに保存ボタン
    document.getElementById('saveAsTemplateBtn').addEventListener('click', () => {
      const date          = document.getElementById('inputDate').value;
      const amountRaw     = document.getElementById('inputAmount').value;
      const category      = document.getElementById('inputCategory').value;
      const description   = document.getElementById('inputDescription').value.trim();
      const paymentMethod = document.querySelector('input[name="paymentMethod"]:checked').value;

      if (!category || !description) {
        this._showMsg('templateMsg', '⚠️ 科目と内容を入力してください', 'error');
        return;
      }
      const amount = Number(amountRaw) || 0;

      // テンプレート名をダイアログで入力
      const name = window.prompt(
        'テンプレート名を入力してください\n（例：ジェルネイル施術、交通費など）',
        description
      );
      if (name === null) return; // キャンセル

      Templates.add({
        name:          name.trim() || description,
        type:          this.currentType,
        amount,
        category,
        description,
        paymentMethod,
      });
      this._showMsg('templateMsg', '✅ テンプレートに保存しました', 'success');
      this._renderTemplateList();
      setTimeout(() => {
        const el = document.getElementById('templateMsg');
        if (el) el.style.display = 'none';
      }, 2000);
    });
  },

  _renderTemplateList() {
    const container = document.getElementById('templateListItems');
    if (!container) return;
    container.innerHTML = '';

    const list = Templates.getAll();
    if (list.length === 0) {
      container.innerHTML = '<p style="color:#9CA3AF;text-align:center;padding:8px 0;font-size:13px;">テンプレートがありません</p>';
      return;
    }

    list.forEach(tpl => {
      const row = document.createElement('div');
      row.style.cssText = 'display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid #F3F4F6;flex-wrap:wrap;';
      const typeLabel = tpl.type === 'income' ? '収入' : '支出';
      const typeColor = tpl.type === 'income' ? '#10B981' : '#EF4444';
      row.innerHTML = `
        <div style="flex:1;min-width:0;">
          <div style="font-size:13px;font-weight:600;">${this._esc(tpl.name)}</div>
          <div style="font-size:11px;color:#6B7280;">
            <span style="color:${typeColor}">${typeLabel}</span>
            ／${this._esc(tpl.category)}
            ${tpl.amount > 0 ? '／' + Format.currency(tpl.amount) : ''}
          </div>
        </div>
        <button class="tpl-apply-btn btn btn-primary" style="padding:4px 12px;font-size:12px;white-space:nowrap;">
          呼び出し
        </button>
        <button class="tpl-delete-btn" style="background:none;border:none;color:#9CA3AF;cursor:pointer;font-size:16px;padding:0 4px;" title="削除">✕</button>
      `;

      row.querySelector('.tpl-apply-btn').addEventListener('click', () => {
        this._applyTemplate(tpl);
      });
      row.querySelector('.tpl-delete-btn').addEventListener('click', () => {
        if (window.confirm(`「${tpl.name}」を削除しますか？`)) {
          Templates.remove(tpl.id);
          this._renderTemplateList();
        }
      });

      container.appendChild(row);
    });
  },

  _applyTemplate(tpl) {
    // 種別を切り替え
    this.currentType = tpl.type;
    document.querySelectorAll('.type-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.type === tpl.type);
    });
    this._rebuildCategoryOptions();

    // フォームに値をセット
    if (tpl.amount > 0) document.getElementById('inputAmount').value = tpl.amount;
    document.getElementById('inputCategory').value    = tpl.category;
    document.getElementById('inputDescription').value = tpl.description;

    // 支払い方法
    const radio = document.querySelector(`input[name="paymentMethod"][value="${tpl.paymentMethod}"]`);
    if (radio) { radio.checked = true; this._updatePaymentStyle(); }

    // テンプレートを閉じる
    const area = document.getElementById('templateListArea');
    const icon = document.getElementById('templateToggleIcon');
    if (area) area.style.display = 'none';
    if (icon) icon.textContent   = '▼ 開く';

    this._showMsg('formMessage', `📋 「${this._esc(tpl.name)}」を呼び出しました`, 'success');
    setTimeout(() => {
      const el = document.getElementById('formMessage');
      if (el) el.style.display = 'none';
    }, 2000);
  },

  /* ─── 編集モード ─────────────────────── */

  _loadForEdit(id) {
    const t = Data.getById(id);
    if (!t) return;

    // 記録タブに切り替え
    this.showTab('record');

    // 編集モードフラグ
    this._editMode  = true;
    this._editingId = id;

    // バナー表示
    const banner = document.getElementById('editModeBanner');
    if (banner) {
      banner.style.display = 'flex';
      const bannerText = document.getElementById('editModeBannerText');
      if (bannerText) bannerText.textContent = `✏️ 編集中：${t.description}`;
    }

    // 種別セット
    this.currentType = t.type;
    document.querySelectorAll('.type-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.type === t.type);
    });
    this._rebuildCategoryOptions();

    // フォームに値をセット
    document.getElementById('inputDate').value        = t.date;
    document.getElementById('inputAmount').value      = t.amount;
    document.getElementById('inputCategory').value    = t.category;
    document.getElementById('inputDescription').value = t.description;

    const radio = document.querySelector(`input[name="paymentMethod"][value="${t.paymentMethod}"]`);
    if (radio) { radio.checked = true; this._updatePaymentStyle(); }

    // 送信ボタンのラベルを変更
    const submitBtn = document.querySelector('#recordForm button[type="submit"]');
    if (submitBtn) submitBtn.textContent = '更新する';

    // フォームにスクロール
    document.getElementById('recordForm').scrollIntoView({ behavior: 'smooth', block: 'start' });
  },

  _cancelEdit() {
    this._editMode  = false;
    this._editingId = null;

    const banner = document.getElementById('editModeBanner');
    if (banner) banner.style.display = 'none';

    const submitBtn = document.querySelector('#recordForm button[type="submit"]');
    if (submitBtn) submitBtn.textContent = '記録する';

    // フォームリセット
    document.getElementById('inputAmount').value      = '';
    document.getElementById('inputDescription').value = '';
    document.getElementById('inputDate').value        = this._todayStr();
  },

  _rebuildCategoryOptions() {
    const sel  = document.getElementById('inputCategory');
    const cats = this.currentType === 'income' ? CATEGORIES.income : CATEGORIES.expense;
    sel.innerHTML = '<option value="">-- 科目を選んでください --</option>';
    cats.forEach(c => {
      const opt = document.createElement('option');
      opt.value = c;
      opt.textContent = c;
      sel.appendChild(opt);
    });
    if (this.currentType === 'income') sel.value = '売上';
  },

  _updatePaymentStyle() {
    const checked = document.querySelector('input[name="paymentMethod"]:checked');
    document.getElementById('payOptCash').classList.toggle(
      'selected', checked && checked.value === 'cash'
    );
    document.getElementById('payOptCard').classList.toggle(
      'selected', checked && checked.value === 'card'
    );
  },

  async _submitRecord() {
    const date          = document.getElementById('inputDate').value;
    const amountRaw     = document.getElementById('inputAmount').value;
    const category      = document.getElementById('inputCategory').value;
    const description   = document.getElementById('inputDescription').value.trim();
    const paymentMethod = document.querySelector('input[name="paymentMethod"]:checked').value;

    if (!date || !amountRaw || !category || !description) {
      this._showMsg('formMessage', '⚠️ 未入力の項目があります', 'error');
      return;
    }
    const amount = Number(amountRaw);
    if (!Number.isFinite(amount) || amount <= 0) {
      this._showMsg('formMessage', '⚠️ 金額は1円以上の数値を入力してください', 'error');
      return;
    }

    const payload    = { date, amount, category, description, paymentMethod, type: this.currentType };
    const submitBtn  = document.querySelector('#recordForm button[type="submit"]');
    const origLabel  = submitBtn ? submitBtn.textContent : '';
    if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = '保存中...'; }

    try {
      if (this._editMode && this._editingId) {
        // 編集モード → 更新
        await Data.update({ id: this._editingId, ...payload });
        this._cancelEdit();
        this._showMsg('formMessage', '✅ 内容を更新しました！', 'success');
      } else {
        // 新規追加
        const label = this.currentType === 'income' ? '売上' : '経費';
        await Data.add(payload);
        document.getElementById('inputAmount').value      = '';
        document.getElementById('inputDescription').value = '';
        document.getElementById('inputDate').value        = this._todayStr();
        this._showMsg('formMessage', `✅ ${label}を記録しました！`, 'success');
      }
      this.renderDashboard();
    } catch (e) {
      this._showMsg('formMessage', `❌ 保存に失敗しました: ${e.message}`, 'error');
    } finally {
      if (submitBtn) { submitBtn.disabled = false; submitBtn.textContent = origLabel; }
    }

    setTimeout(() => {
      const el = document.getElementById('formMessage');
      if (el) el.style.display = 'none';
    }, 3000);
  },

  /* ─── 削除モーダル ───────────────────── */

  initModal() {
    // キャンセルボタン・オーバーレイ・Esc はここで登録（固定動作）
    document.getElementById('modalCancel').addEventListener('click', () => {
      this._closeModal();
    });

    document.getElementById('modal').addEventListener('click', e => {
      if (e.target === document.getElementById('modal')) this._closeModal();
    });

    document.addEventListener('keydown', e => {
      if (e.key === 'Escape') this._closeModal();
    });
  },

  _openDeleteModal(id, desc) {
    this._pendingDeleteId = id;
    document.getElementById('modalMessage').textContent =
      `「${desc}」を削除します。この操作は取り消せません。`;

    // 確定ボタンを cloneNode でリセット → 古いリスナーの蓄積・競合を防ぐ
    const oldBtn = document.getElementById('modalConfirm');
    const newBtn = oldBtn.cloneNode(true);
    oldBtn.replaceWith(newBtn);

    newBtn.addEventListener('click', async () => {
      this._closeModal();
      try {
        await Data.remove(id);
        this._applyFilters();
        if (document.getElementById('tab-dashboard').classList.contains('active')) {
          this.renderDashboard();
        }
      } catch (e) {
        this._showToast(`❌ 削除に失敗しました: ${e.message}`);
      }
    });

    document.getElementById('modal').style.display = 'flex';
    newBtn.focus();
  },

  _closeModal() {
    document.getElementById('modal').style.display = 'none';
    this._pendingDeleteId = null;
  },

  /** 画面上部にトースト通知を表示（4秒後に自動消去） */
  _showToast(msg, type = 'error') {
    const styles = {
      error: { bg: '#FEE2E2', color: '#DC2626', border: '#FCA5A5' },
      success: { bg: '#D1FAE5', color: '#065F46', border: '#6EE7B7' },
      warn:  { bg: '#FEF3C7', color: '#92400E', border: '#FCD34D' },
    };
    const s = styles[type] || styles.error;
    const toast = document.createElement('div');
    toast.style.cssText = [
      'position:fixed',
      'top:70px',
      'left:50%',
      'transform:translateX(-50%)',
      'max-width:90vw',
      `background:${s.bg}`,
      `color:${s.color}`,
      `border:1px solid ${s.border}`,
      'border-radius:8px',
      'padding:10px 20px',
      'z-index:2000',
      'font-size:14px',
      'text-align:center',
      'box-shadow:0 4px 12px rgba(0,0,0,0.15)',
    ].join(';');
    toast.textContent = msg;
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), 4000);
  },

  /* ─── 設定フォーム ───────────────────── */

  initSettings() {
    document.getElementById('settingsForm').addEventListener('submit', e => {
      e.preventDefault();
      const settings = {
        businessName: document.getElementById('businessName').value.trim(),
        initialCash:  Number(document.getElementById('initialCash').value)  || 0,
        initialBank:  Number(document.getElementById('initialBank').value) || 0,
      };
      Storage.saveSettings(settings);

      if (GasAPI.isConfigured()) {
        GasAPI.saveSettings(settings).catch(e => {
          console.warn('[Settings] GAS への保存に失敗:', e.message);
        });
      }

      if (settings.businessName) {
        document.getElementById('headerSubtitle').textContent =
          settings.businessName + '  ┊  青色申告対応';
      }
      this._showMsg('settingsMessage', '✅ 設定を保存しました', 'success');
      setTimeout(() => {
        const el = document.getElementById('settingsMessage');
        if (el) el.style.display = 'none';
      }, 3000);
    });

    document.getElementById('exportData').addEventListener('click', () => {
      const payload = {
        appVersion:   APP_VERSION,
        exportedAt:   new Date().toISOString(),
        settings:     Storage.getSettings(),
        transactions: Data.getAll(),
        templates:    Templates.getAll(),
      };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      a.href     = url;
      a.download = `salon_kaikei_backup_${this._todayStr()}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    });

    document.getElementById('deleteAllData').addEventListener('click', () => {
      const ok = window.confirm(
        '⚠️ 端末上のすべてのデータを削除します。\n' +
        'この操作は取り消せません。\n\n' +
        '※ Google スプレッドシート上のデータは削除されません。\n\n' +
        '本当に削除しますか？'
      );
      if (ok) {
        localStorage.removeItem(STORAGE_KEY);
        localStorage.removeItem(SETTINGS_KEY);
        localStorage.removeItem(TEMPLATES_KEY);
        window.location.reload();
      }
    });
  },

  /* ─── フィルター ─────────────────────── */

  initFilters() {
    ['filterMonth', 'filterType', 'filterCategory'].forEach(id => {
      document.getElementById(id).addEventListener('change', () => this._applyFilters());
    });
  },

  /* ─── 集計年変更 ─────────────────────── */

  initSummaryYear() {
    document.getElementById('summaryYear').addEventListener('change', e => {
      this._updateSummary(Number(e.target.value));
    });
  },

  /* ─── ユーティリティ ─────────────────── */

  _esc(str) {
    const el = document.createElement('div');
    el.textContent = String(str || '');
    return el.innerHTML;
  },

  _todayStr() {
    return new Date().toISOString().split('T')[0];
  },

  _showMsg(id, text, type) {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent   = text;
    el.className     = `form-message ${type}`;
    el.style.display = 'block';
  },

  /* ─── オフライン検知 ────────────────────── */

  initOfflineDetection() {
    window.addEventListener('online',  async () => {
      this._setOnlineState(true);
      // オンライン復帰時に GAS からデータを再取得してキャッシュを更新
      if (GasAPI.isConfigured()) {
        await Storage.loadFromGas();
        this.renderDashboard();
        if (document.getElementById('tab-list')?.classList.contains('active')) {
          this.renderList();
        }
      }
    });
    window.addEventListener('offline', () => this._setOnlineState(false));
    if (!navigator.onLine) this._setOnlineState(false);
  },

  _setOnlineState(isOnline) {
    // オフラインバナーの表示・非表示
    let banner = document.getElementById('offlineBanner');
    if (!isOnline && !banner) {
      banner = document.createElement('div');
      banner.id = 'offlineBanner';
      banner.style.cssText = [
        'background:#FEF3C7',
        'color:#92400E',
        'border-bottom:2px solid #FCD34D',
        'text-align:center',
        'padding:10px 16px',
        'font-size:13px',
        'font-weight:600',
        'width:100%',
        'box-sizing:border-box',
      ].join(';');
      banner.textContent = '📵 オフライン中は入力・編集・削除ができません';
      const nav = document.querySelector('nav') || document.querySelector('header');
      if (nav && nav.parentNode) {
        nav.parentNode.insertBefore(banner, nav.nextSibling);
      } else {
        document.body.prepend(banner);
      }
    } else if (isOnline && banner) {
      banner.remove();
    }

    // フォーム送信ボタン
    const submitBtn = document.querySelector('#recordForm button[type="submit"]');
    if (submitBtn) submitBtn.disabled = !isOnline;

    // 動的生成済みの編集・削除ボタン
    document.querySelectorAll('.edit-btn, .delete-btn').forEach(btn => {
      btn.disabled     = !isOnline;
      btn.style.opacity = isOnline ? '' : '0.4';
      btn.style.cursor  = isOnline ? '' : 'not-allowed';
    });
  },
};

/* =============================================
   アプリ起動
   ============================================= */
async function init() {
  UI.initTabs();
  UI.initForm();
  UI.initTemplates();
  UI.initModal();
  UI.initSettings();
  UI.initFilters();
  UI.initSummaryYear();
  UI.initOfflineDetection();

  const settings = Storage.getSettings();
  if (settings.businessName && settings.businessName !== 'マイサロン') {
    document.getElementById('headerSubtitle').textContent =
      settings.businessName + '  ┊  青色申告対応';
  }

  // キャッシュを使って即時表示
  UI.renderDashboard();

  if (!GasAPI.isConfigured()) {
    // GAS URL未設定（app.js の GAS_URL 定数を設定してください）
    UI._showToast('⚠️ GAS URL が未設定です。app.js の GAS_URL 定数を設定してください。', 'warn');
    return;
  }

  if (navigator.onLine) {
    // GAS からデータを取得してキャッシュを更新
    const loaded = await Storage.loadFromGas();
    if (loaded) UI.renderDashboard();
  }
}

document.addEventListener('DOMContentLoaded', init);
