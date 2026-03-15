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

const APP_VERSION      = '2.0.0';
const STORAGE_KEY      = 'salon_kaikei_v1_transactions';
const SETTINGS_KEY     = 'salon_kaikei_v1_settings';
const TEMPLATES_KEY    = 'salon_kaikei_v1_templates';
const ALLOCATION_KEY   = 'salon_kaikei_v1_allocation';
const AUTO_RULES_KEY   = 'salon_kaikei_v1_auto_rules';
const MASTER_KEY       = 'salon_kaikei_v1_master';
const CUSTOMERS_KEY     = 'charm_plus_customers';
const APPOINTMENTS_KEY  = 'charm_plus_appointments';

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

  async getMaster() {
    const res = await fetch(`${GAS_URL}?action=getMaster`, { redirect: 'follow' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (json.error) throw new Error(json.error);
    return json;
  },

  async saveMaster(data) {
    return this._post({ action: 'saveMaster', data });
  },

  async getCustomers() {
    const res = await fetch(`${GAS_URL}?action=getCustomers`, { redirect: 'follow' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (json.error) throw new Error(json.error);
    return json.customers || [];
  },

  async addCustomer(data) {
    return this._post({ action: 'addCustomer', data });
  },

  async updateCustomer(data) {
    return this._post({ action: 'updateCustomer', data });
  },

  async deleteCustomer(id) {
    return this._post({ action: 'deleteCustomer', id });
  },

  async getAppointments(from, to) {
    const url = `${GAS_URL}?action=getAppointments&from=${from}&to=${to}`;
    const res = await fetch(url, { redirect: 'follow' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (json.error) throw new Error(json.error);
    return json.appointments || [];
  },

  async updateAppointmentStatus(id, status, staffNote) {
    return this._post({ action: 'updateAppointmentStatus', id, status, staffNote });
  },

  async cancelAppointment(id) {
    return this._post({ action: 'cancelAppointment', id });
  },

  /** 領収書画像をGoogle Driveにアップロードし、ファイルIDを返す */
  async uploadReceipt(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const base64 = e.target.result.split(',')[1];
          const result = await this._post({
            action: 'uploadReceipt',
            fileName: file.name,
            mimeType: file.type,
            base64Data: base64,
          });
          resolve(result.fileId || result.id || '');
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = () => reject(new Error('ファイル読み込みに失敗しました'));
      reader.readAsDataURL(file);
    });
  },

  /** GoogleドライブのファイルURLを生成 */
  getDriveViewUrl(fileId) {
    return `https://drive.google.com/file/d/${fileId}/view`;
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

  async loadMasterFromGas() {
    try {
      const data = await GasAPI.getMaster();
      if (data && typeof data === 'object') {
        Master.loadFromData(data);
      }
      return true;
    } catch (e) {
      console.warn('[Storage] GASからマスタデータ取得に失敗（ローカルキャッシュを使用）:', e.message);
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
      if (!t.date) return false;
      const [y, m] = t.date.substring(0, 7).split('-').map(Number);
      return y === year && m === month;
    });
  },

  /** 年で絞り込み */
  getByYear(year) {
    return this.getAll().filter(t => {
      if (!t.date) return false;
      return Number(t.date.substring(0, 4)) === year;
    });
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
   Allocation モジュール
   家事按分：科目ごとの事業使用割合を管理
   ============================================= */
const Allocation = {
  /** 全按分設定を取得 { 水道光熱費: 50, 地代家賃: 30, ... } */
  getAll() {
    try {
      return JSON.parse(localStorage.getItem(ALLOCATION_KEY) || '{}');
    } catch { return {}; }
  },

  save(map) {
    localStorage.setItem(ALLOCATION_KEY, JSON.stringify(map));
  },

  /** 科目の事業使用割合（%）を取得。設定なければ100 */
  getRatio(category) {
    const map = this.getAll();
    const v = map[category];
    return (v !== undefined && v !== '') ? Number(v) : 100;
  },

  /** 按分後の金額を計算（支出のみ対象） */
  applyRatio(amount, category) {
    const ratio = this.getRatio(category);
    return Math.round(Number(amount) * ratio / 100);
  },

  /** 按分が設定されている（100%未満）か */
  hasAllocation(category) {
    return this.getRatio(category) < 100;
  },
};

/* =============================================
   AutoRules モジュール
   自動登録ルール：支払方法＋金額帯で科目を自動入力
   ============================================= */
const AutoRules = {
  getAll() {
    try {
      return JSON.parse(localStorage.getItem(AUTO_RULES_KEY) || '[]');
    } catch { return []; }
  },

  _save(list) {
    localStorage.setItem(AUTO_RULES_KEY, JSON.stringify(list));
  },

  add(rule) {
    const list = this.getAll();
    const record = {
      id: `rule-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      name: (rule.name || '').trim() || 'ルール',
      type: rule.type || 'expense',
      paymentMethod: rule.paymentMethod || '',   // '' = どちらでも
      amountMin: Number(rule.amountMin) || 0,
      amountMax: rule.amountMax !== '' && rule.amountMax !== undefined
        ? Number(rule.amountMax)
        : 9_999_999_999,
      category: rule.category || '',
    };
    list.push(record);
    this._save(list);
    return record;
  },

  remove(id) {
    this._save(this.getAll().filter(r => r.id !== id));
  },

  /** 条件に合う最初のルールのcategoryを返す。なければnull */
  match(type, paymentMethod, amount) {
    const num = Number(amount) || 0;
    for (const rule of this.getAll()) {
      if (rule.type !== type) continue;
      if (rule.paymentMethod && rule.paymentMethod !== paymentMethod) continue;
      if (num < rule.amountMin) continue;
      if (num > rule.amountMax) continue;
      if (!rule.category) continue;
      return rule.category;
    }
    return null;
  },
};

/* =============================================
   Master モジュール
   勘定科目・支払方法・タグのマスタデータ管理
   ============================================= */
const Master = {
  _raw() {
    try { return JSON.parse(localStorage.getItem(MASTER_KEY) || '{}'); } catch { return {}; }
  },

  _setRaw(obj) {
    localStorage.setItem(MASTER_KEY, JSON.stringify(obj));
  },

  /** 勘定科目一覧（income/expense 両方） */
  getCategories() {
    const raw = this._raw();
    return {
      income:  raw.income  ? [...raw.income]  : [...CATEGORIES.income],
      expense: raw.expense ? [...raw.expense] : [...CATEGORIES.expense],
    };
  },

  saveCategories(cats) {
    const raw = this._raw();
    this._setRaw({ ...raw, income: cats.income, expense: cats.expense });
  },

  /** 支払方法一覧 */
  getPaymentMethods() {
    const raw = this._raw();
    return raw.paymentMethods || [
      { value: 'cash', label: '現金' },
      { value: 'card', label: 'カード／振込' },
    ];
  },

  savePaymentMethods(methods) {
    this._setRaw({ ...this._raw(), paymentMethods: methods });
  },

  /** タグ一覧 */
  getTags() {
    return this._raw().tags || [];
  },

  saveTags(tags) {
    this._setRaw({ ...this._raw(), tags });
  },

  /** メニュー一覧（名前・価格・カテゴリ） */
  getMenus() {
    return this._raw().menus || [
      { id: 'menu_default_1', name: 'ジェルネイル（ワンカラー）', price: 7000,  category: '売上', duration: 90 },
      { id: 'menu_default_2', name: 'ジェルネイル（アート）',     price: 10000, category: '売上', duration: 120 },
      { id: 'menu_default_3', name: 'まつエク（120本）',          price: 8000,  category: '売上', duration: 90 },
      { id: 'menu_default_4', name: 'まつエク（160本）',          price: 11000, category: '売上', duration: 120 },
    ];
  },

  /** 営業時間設定 */
  getBusinessHours() {
    return this._raw().businessHours || {
      mon: { open: '10:00', close: '19:00' },
      tue: { open: '10:00', close: '19:00' },
      wed: null,
      thu: { open: '10:00', close: '19:00' },
      fri: { open: '10:00', close: '19:00' },
      sat: { open: '10:00', close: '18:00' },
      sun: null,
    };
  },

  saveBusinessHours(hours) {
    this._setRaw({ ...this._raw(), businessHours: hours });
  },

  /** 予約可能日数（今日から何日先まで受け付けるか） */
  getBookingWindowDays() {
    return this._raw().bookingWindowDays || 60;
  },

  saveBookingWindowDays(days) {
    this._setRaw({ ...this._raw(), bookingWindowDays: Number(days) || 60 });
  },

  saveMenus(menus) {
    this._setRaw({ ...this._raw(), menus });
  },

  /** GASから取得したデータを反映 */
  loadFromData(data) {
    if (!data || typeof data !== 'object') return;
    const existing = this._raw();
    this._setRaw({ ...existing, ...data });
  },

  /** GAS保存用に全データを返す */
  getAllData() {
    return this._raw();
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
      const adjusted = Allocation.applyRatio(Number(t.amount), t.category);
      map[t.category] = (map[t.category] || 0) + adjusted;
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
      Data.getAll()
        .filter(t => t.date && t.date.length >= 4)
        .map(t => Number(t.date.substring(0, 4)))
        .filter(y => !isNaN(y))
    );
    years.add(new Date().getFullYear());
    return [...years].sort((a, b) => b - a);
  },

  /**
   * 収支の合計を計算。
   * income はそのまま、expense は按分後の金額を使用（税務上の経費計上額）。
   */
  _sumTransactions(list) {
    const income  = list.filter(t => t.type === 'income')
                        .reduce((s, t) => s + Number(t.amount), 0);
    const expense = list.filter(t => t.type === 'expense')
                        .reduce((s, t) => s + Allocation.applyRatio(Number(t.amount), t.category), 0);
    return { income, expense, profit: income - expense };
  },

  /** 按分前の経費合計（実際の支出金額） */
  _sumExpenseRaw(list) {
    return list.filter(t => t.type === 'expense')
               .reduce((s, t) => s + Number(t.amount), 0);
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
    if (name === 'ledger')    this.renderLedger();
    if (name === 'settings')  this.renderSettings();
    if (name === 'customers')    CustomerUI.renderCustomers();
    if (name === 'appointments') AppointmentUI.renderAppointments();
  },

  /* ─── ダッシュボード ─────────────────── */

  renderDashboard() {
    const { cash, bank } = Calculator.getBalances();
    const { income, expense } = Calculator.getCurrentMonthSummary();

    document.getElementById('cashBalance').textContent    = Format.currency(cash);
    document.getElementById('bankBalance').textContent    = Format.currency(bank);
    document.getElementById('monthlyIncome').textContent  = Format.currency(income);
    document.getElementById('monthlyExpense').textContent = Format.currency(expense);

    // 今月の予約売上見込み（confirmed の合計）
    const yearMonth = new Date().toISOString().substring(0, 7);
    const forecast = AppointmentData.getAll()
      .filter(a => (a.status === 'confirmed' || a.status === 'pending') &&
                   (a.dateTime || '').startsWith(yearMonth))
      .reduce((sum, a) => sum + Number(a.price || 0), 0);
    const forecastEl = document.getElementById('forecastIncome');
    if (forecastEl) forecastEl.textContent = Format.currency(forecast);

    this._renderTodaySchedule();

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

  _renderTodaySchedule() {
    const el = document.getElementById('todaySchedule');
    if (!el) return;
    const today = new Date().toISOString().split('T')[0];
    const appts = AppointmentData.getByDate(today)
      .filter(a => a.status === 'pending' || a.status === 'confirmed')
      .sort((a, b) => (a.dateTime || '').localeCompare(b.dateTime || ''));

    if (!appts.length) {
      el.innerHTML = '<p style="text-align:center;color:var(--text-light);font-size:13px;padding:20px 0;">今日の予約はありません</p>';
      return;
    }
    const badge = { pending: '🟡', confirmed: '🟢' };
    el.innerHTML = appts.map(a => {
      const time = (a.dateTime || '').substring(11, 16);
      return `<div style="display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid var(--border);">
        <span style="font-family:var(--font-serif);font-size:1.1rem;font-weight:700;min-width:48px;">${time}</span>
        <span style="font-size:14px;">${badge[a.status] || ''}</span>
        <div style="flex:1;min-width:0;">
          <div style="font-weight:700;font-size:14px;">${UI._esc(a.customerName || '')}</div>
          <div style="font-size:12px;color:var(--text-sub);">${UI._esc(a.menuName || '')} ／ ¥${Number(a.price || 0).toLocaleString()}</div>
        </div>
        <span style="font-size:12px;color:var(--text-sub);flex-shrink:0;">${Number(a.duration || 0)}分</span>
      </div>`;
    }).join('');
  },

  /* ─── 収支一覧 ────────────────────────── */

  renderList() {
    this._rebuildFilterOptions();
    this._applyFilters();
  },

  _rebuildFilterOptions() {
    const months = new Set();
    const tags   = new Set();
    Data.getAll().forEach(t => {
      // 文字列から直接 YYYY-MM を取り出す（タイムゾーン問題を回避）
      if (t.date && t.date.length >= 7) {
        months.add(t.date.substring(0, 7));
      }
      if (Array.isArray(t.tags)) t.tags.forEach(tag => tag && tags.add(tag));
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

    const cats = Master.getCategories();
    const catSel = document.getElementById('filterCategory');
    const prevCat = catSel.value;
    catSel.innerHTML = '<option value="">科目すべて</option>';
    [...cats.income, ...cats.expense].forEach(cat => {
      const opt = document.createElement('option');
      opt.value = cat;
      opt.textContent = cat;
      catSel.appendChild(opt);
    });
    catSel.value = prevCat;

    const tagSel = document.getElementById('filterTag');
    if (tagSel) {
      const prevTag = tagSel.value;
      tagSel.innerHTML = '<option value="">タグすべて</option>';
      [...tags].sort().forEach(tag => {
        const opt = document.createElement('option');
        opt.value = tag;
        opt.textContent = tag;
        tagSel.appendChild(opt);
      });
      tagSel.value = prevTag;
    }
  },

  _applyFilters() {
    let list = Data.getAll();
    const month    = document.getElementById('filterMonth').value;
    const type     = document.getElementById('filterType').value;
    const category = document.getElementById('filterCategory').value;
    const tagSel   = document.getElementById('filterTag');
    const tag      = tagSel ? tagSel.value : '';

    if (month) {
      list = list.filter(t => t.date && t.date.startsWith(month));
    }
    if (type)     list = list.filter(t => t.type === type);
    if (category) list = list.filter(t => t.category === category);
    if (tag)      list = list.filter(t => Array.isArray(t.tags) && t.tags.includes(tag));

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

    const methods = Master.getPaymentMethods();
    const pm = methods.find(m => m.value === t.paymentMethod);
    const pmIcon = t.paymentMethod === 'cash' ? '💴' : '💳';
    const pmLabel = pm ? pm.label : (t.paymentMethod || '');
    const payTag = `<span class="tag ${t.paymentMethod === 'cash' ? 'cash' : 'card'}">${pmIcon} ${this._esc(pmLabel)}</span>`;

    const sign = t.type === 'income' ? '+' : '−';

    // 按分後金額の表示
    const adjusted = t.type === 'expense' ? Allocation.applyRatio(Number(t.amount), t.category) : Number(t.amount);
    const amountDisplay = Format.currency(adjusted);
    const hasAlloc = t.type === 'expense' && Allocation.hasAllocation(t.category);
    const allocNote = hasAlloc
      ? `<small style="font-size:10px;color:#6B7280;display:block;">入力額 ${Format.currency(t.amount)}</small>`
      : '';

    // タグ表示
    const tagHtml = Array.isArray(t.tags) && t.tags.length > 0
      ? t.tags.map(tag => `<span class="tag" style="background:#EDE9FE;color:#5B21B6;font-size:10px;">${this._esc(tag)}</span>`).join('')
      : '';

    // 領収書アイコン
    const receiptBtn = t.receiptId
      ? `<button class="receipt-btn" title="領収書を見る" data-receipt-id="${this._esc(t.receiptId)}"
           style="background:none;border:1px solid #D1D5DB;border-radius:6px;padding:4px 8px;
                  font-size:12px;cursor:pointer;color:#374151;margin-right:2px;white-space:nowrap;">📎</button>`
      : '';

    const actionBtns = showActions ? `
      ${receiptBtn}
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
          ${tagHtml}
        </div>
      </div>
      <div class="transaction-amount ${t.type}" aria-label="${t.type === 'income' ? '収入' : '支出'} ${amountDisplay}">
        ${sign}${amountDisplay}
        ${allocNote}
      </div>
      ${actionBtns}
    `;

    if (showActions) {
      const editBtn    = div.querySelector('.edit-btn');
      const delBtn     = div.querySelector('.delete-btn');
      const rcptBtn    = div.querySelector('.receipt-btn');

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

      if (rcptBtn) {
        rcptBtn.addEventListener('click', () => {
          const fileId = rcptBtn.dataset.receiptId;
          this._openReceiptModal(fileId);
        });
      }
    }
    return div;
  },

  _openReceiptModal(fileId) {
    const modal   = document.getElementById('receiptModal');
    const content = document.getElementById('receiptModalContent');
    const url     = GasAPI.getDriveViewUrl(fileId);
    content.innerHTML = `
      <p style="font-size:13px;color:#6B7280;margin-bottom:12px;">
        Googleドライブで画像を確認できます。
      </p>
      <a href="${url}" target="_blank" rel="noopener noreferrer"
         class="btn btn-primary" style="display:inline-block;">
        📎 Googleドライブで開く
      </a>
    `;
    modal.style.display = 'flex';
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
    const methods   = Master.getPaymentMethods();
    const payLabel  = t => {
      const pm = methods.find(m => m.value === t.paymentMethod);
      return pm ? pm.label : t.paymentMethod;
    };

    const header = '日付,種別,勘定科目,内容,金額（入力）,金額（按分後）,按分率(%),支払方法,タグ';
    const rows   = sorted.map(t => {
      const ratio    = t.type === 'expense' ? Allocation.getRatio(t.category) : 100;
      const adjusted = t.type === 'expense' ? Allocation.applyRatio(Number(t.amount), t.category) : Number(t.amount);
      const tagsStr  = Array.isArray(t.tags) ? t.tags.join('／') : '';
      return [
        t.date,
        typeLabel(t),
        `"${t.category}"`,
        `"${String(t.description).replace(/"/g, '""')}"`,
        t.amount,
        adjusted,
        ratio,
        `"${payLabel(t)}"`,
        `"${tagsStr}"`,
      ].join(',');
    });

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

    this._injectAllocationSection();
    this._injectAutoRulesSection();
    this._injectMenuSection();
    this._injectBusinessHoursSection();
    this._injectQrSection();
    this._injectCategorySection();
    this._injectPaymentMethodSection();
    this._injectTagSection();
  },

  /* ─── 家事按分設定 ──────────────────── */

  _injectAllocationSection() {
    if (document.getElementById('allocationSection')) {
      this._renderAllocationSection();
      return;
    }
    const danger = document.querySelector('.danger-zone');
    const section = document.createElement('div');
    section.id        = 'allocationSection';
    section.className = 'card form-card';
    section.style.cssText = 'margin-top:16px;';
    danger.parentNode.insertBefore(section, danger);
    this._renderAllocationSection();
  },

  _renderAllocationSection() {
    const section = document.getElementById('allocationSection');
    if (!section) return;
    const allocation = Allocation.getAll();
    const cats = Master.getCategories().expense;

    section.innerHTML = `
      <h2 class="section-title">家事按分の設定</h2>
      <p style="font-size:13px;color:#6B7280;margin-bottom:12px;line-height:1.6">
        自宅サロンや家事と兼用の経費は、事業で使っている割合（%）を設定すると、
        集計・CSVに按分後の金額が使われます。
      </p>
      <div id="allocationRows"></div>
      <button type="button" id="saveAllocationBtn" class="btn btn-primary" style="margin-top:12px;width:100%;">
        按分設定を保存する
      </button>
      <div id="allocationMsg" class="form-message" style="display:none;margin-top:8px;"></div>
    `;

    const rowsEl = document.getElementById('allocationRows');
    cats.forEach(cat => {
      const ratio = allocation[cat] !== undefined ? allocation[cat] : 100;
      const row = document.createElement('div');
      row.style.cssText = 'display:flex;align-items:center;gap:12px;padding:6px 0;border-bottom:1px solid #F3F4F6;';
      row.innerHTML = `
        <div style="flex:1;font-size:14px;">${this._esc(cat)}</div>
        <div style="display:flex;align-items:center;gap:6px;">
          <input type="number" class="allocation-input" data-cat="${this._esc(cat)}"
            value="${ratio}" min="0" max="100" step="1"
            style="width:70px;border:1px solid #E5E7EB;border-radius:6px;padding:4px 8px;font-size:14px;text-align:right;">
          <span style="font-size:13px;color:#6B7280;">%</span>
        </div>
      `;
      rowsEl.appendChild(row);
    });

    document.getElementById('saveAllocationBtn').addEventListener('click', () => {
      const map = {};
      document.querySelectorAll('.allocation-input').forEach(input => {
        const cat = input.dataset.cat;
        const val = Number(input.value);
        if (cat) map[cat] = Math.min(100, Math.max(0, isNaN(val) ? 100 : val));
      });
      Allocation.save(map);
      this._showMsg('allocationMsg', '✅ 按分設定を保存しました', 'success');
      setTimeout(() => {
        const el = document.getElementById('allocationMsg');
        if (el) el.style.display = 'none';
      }, 2500);
      // 関連UIを更新
      this._updateAllocationDisplay();
    });
  },

  /* ─── 自動登録ルール ─────────────────── */

  _injectAutoRulesSection() {
    if (document.getElementById('autoRulesSection')) {
      this._renderAutoRulesSection();
      return;
    }
    const danger = document.querySelector('.danger-zone');
    const section = document.createElement('div');
    section.id        = 'autoRulesSection';
    section.className = 'card form-card';
    section.style.cssText = 'margin-top:16px;';
    danger.parentNode.insertBefore(section, danger);
    this._renderAutoRulesSection();
  },

  _renderAutoRulesSection() {
    const section = document.getElementById('autoRulesSection');
    if (!section) return;
    const cats = Master.getCategories();
    const methods = Master.getPaymentMethods();

    const catOptions = (type) => [
      ...( type === 'expense' ? cats.expense : cats.income ),
    ].map(c => `<option value="${this._esc(c)}">${this._esc(c)}</option>`).join('');

    const pmOptions = `
      <option value="">どちらでも</option>
      ${methods.map(m => `<option value="${this._esc(m.value)}">${this._esc(m.label)}</option>`).join('')}
    `;

    section.innerHTML = `
      <h2 class="section-title">自動登録ルール</h2>
      <p style="font-size:13px;color:#6B7280;margin-bottom:12px;line-height:1.6">
        支払い方法と金額帯が条件に合うとき、科目を自動で入力します。
      </p>
      <div id="autoRulesList" style="margin-bottom:12px;"></div>
      <div class="card" style="background:#F9FAFB;border:1px solid #E5E7EB;padding:14px;">
        <h3 style="font-size:14px;font-weight:bold;margin-bottom:10px;">ルールを追加</h3>
        <div class="form-row">
          <div class="form-group">
            <label style="font-size:12px;">種別</label>
            <select id="newRuleType" class="filter-select" style="width:100%;">
              <option value="expense">支出（経費）</option>
              <option value="income">収入（売上）</option>
            </select>
          </div>
          <div class="form-group">
            <label style="font-size:12px;">支払い方法</label>
            <select id="newRulePm" class="filter-select" style="width:100%;">
              ${pmOptions}
            </select>
          </div>
        </div>
        <div class="form-row">
          <div class="form-group">
            <label style="font-size:12px;">金額（下限・円）</label>
            <input type="number" id="newRuleMin" placeholder="0" min="0" style="width:100%;border:1px solid #E5E7EB;border-radius:6px;padding:6px 10px;font-size:14px;">
          </div>
          <div class="form-group">
            <label style="font-size:12px;">金額（上限・円）</label>
            <input type="number" id="newRuleMax" placeholder="上限なし" min="0" style="width:100%;border:1px solid #E5E7EB;border-radius:6px;padding:6px 10px;font-size:14px;">
          </div>
        </div>
        <div class="form-group">
          <label style="font-size:12px;">科目（自動入力する科目）</label>
          <select id="newRuleCat" class="filter-select" style="width:100%;">
            <option value="">-- 選んでください --</option>
            ${catOptions('expense')}
          </select>
        </div>
        <div class="form-group">
          <label style="font-size:12px;">ルール名（メモ）</label>
          <input type="text" id="newRuleName" placeholder="例：小額現金は消耗品費" maxlength="40"
            style="width:100%;border:1px solid #E5E7EB;border-radius:6px;padding:6px 10px;font-size:14px;">
        </div>
        <button type="button" id="addRuleBtn" class="btn btn-primary" style="width:100%;">ルールを追加する</button>
        <div id="ruleMsg" class="form-message" style="display:none;margin-top:6px;"></div>
      </div>
    `;

    // 種別変更で科目リストを更新
    document.getElementById('newRuleType').addEventListener('change', e => {
      const catSel = document.getElementById('newRuleCat');
      catSel.innerHTML = `<option value="">-- 選んでください --</option>${catOptions(e.target.value)}`;
    });

    document.getElementById('addRuleBtn').addEventListener('click', () => {
      const type = document.getElementById('newRuleType').value;
      const pm   = document.getElementById('newRulePm').value;
      const min  = document.getElementById('newRuleMin').value;
      const max  = document.getElementById('newRuleMax').value;
      const cat  = document.getElementById('newRuleCat').value;
      const name = document.getElementById('newRuleName').value.trim();
      if (!cat) {
        this._showMsg('ruleMsg', '⚠️ 科目を選んでください', 'error');
        return;
      }
      AutoRules.add({ type, paymentMethod: pm, amountMin: min || 0, amountMax: max || '', category: cat, name });
      this._showMsg('ruleMsg', '✅ ルールを追加しました', 'success');
      setTimeout(() => { const el = document.getElementById('ruleMsg'); if (el) el.style.display = 'none'; }, 2000);
      this._renderAutoRulesList();
    });

    this._renderAutoRulesList();
  },

  _renderAutoRulesList() {
    const container = document.getElementById('autoRulesList');
    if (!container) return;
    container.innerHTML = '';
    const rules = AutoRules.getAll();
    if (rules.length === 0) {
      container.innerHTML = '<p style="color:#9CA3AF;font-size:13px;text-align:center;padding:8px 0;">ルールがありません</p>';
      return;
    }
    rules.forEach(rule => {
      const row = document.createElement('div');
      row.style.cssText = 'display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid #F3F4F6;flex-wrap:wrap;';
      const pmLabel = rule.paymentMethod
        ? (Master.getPaymentMethods().find(m => m.value === rule.paymentMethod)?.label || rule.paymentMethod)
        : 'どちらでも';
      const maxStr  = rule.amountMax >= 9_999_999_999 ? '上限なし' : `${Number(rule.amountMax).toLocaleString()}円`;
      row.innerHTML = `
        <div style="flex:1;min-width:0;">
          <div style="font-size:13px;font-weight:600;">${this._esc(rule.name)}</div>
          <div style="font-size:11px;color:#6B7280;">
            ${rule.type === 'expense' ? '支出' : '収入'} ／ ${this._esc(pmLabel)}
            ／ ${Number(rule.amountMin).toLocaleString()}円〜${maxStr}
            → <span style="color:#7C3AED;font-weight:600;">${this._esc(rule.category)}</span>
          </div>
        </div>
        <button class="del-rule-btn" data-id="${rule.id}" style="background:none;border:1px solid #FECACA;
          border-radius:6px;padding:4px 10px;font-size:12px;cursor:pointer;color:#EF4444;white-space:nowrap;">
          削除
        </button>
      `;
      row.querySelector('.del-rule-btn').addEventListener('click', e => {
        AutoRules.remove(e.target.dataset.id);
        this._renderAutoRulesList();
      });
      container.appendChild(row);
    });
  },

  /* ─── 勘定科目管理 ──────────────────── */

  _injectCategorySection() {
    if (document.getElementById('categoryMgmtSection')) {
      this._renderCategorySection();
      return;
    }
    const danger = document.querySelector('.danger-zone');
    const section = document.createElement('div');
    section.id        = 'categoryMgmtSection';
    section.className = 'card form-card';
    section.style.cssText = 'margin-top:16px;';
    danger.parentNode.insertBefore(section, danger);
    this._renderCategorySection();
  },

  _renderCategorySection() {
    const section = document.getElementById('categoryMgmtSection');
    if (!section) return;
    const cats = Master.getCategories();
    const usedCats = new Set(Data.getAll().map(t => t.category));

    const renderList = (type) => {
      const items = cats[type];
      return items.map(c => {
        const used = usedCats.has(c);
        const isDefault = CATEGORIES[type].includes(c);
        return `
          <div style="display:flex;align-items:center;gap:8px;padding:4px 0;border-bottom:1px solid #F9FAFB;">
            <span style="flex:1;font-size:13px;">${this._esc(c)}</span>
            ${isDefault ? '<span style="font-size:11px;color:#9CA3AF;">既定</span>' : ''}
            ${used
              ? '<span style="font-size:11px;color:#FCA5A5;" title="この科目は取引データに使用されているため削除できません">使用中</span>'
              : `<button class="del-cat-btn" data-type="${type}" data-cat="${this._esc(c)}"
                   style="background:none;border:1px solid #FECACA;border-radius:6px;padding:2px 8px;
                          font-size:11px;cursor:pointer;color:#EF4444;">削除</button>`
            }
          </div>`;
      }).join('');
    };

    section.innerHTML = `
      <h2 class="section-title">勘定科目の管理</h2>
      <div style="margin-bottom:16px;">
        <h3 style="font-size:13px;font-weight:bold;color:#10B981;margin-bottom:6px;">収入科目</h3>
        <div id="incomeCatList">${renderList('income')}</div>
        <div style="display:flex;gap:8px;margin-top:8px;">
          <input type="text" id="newIncomeCat" placeholder="新しい収入科目名" maxlength="20"
            style="flex:1;border:1px solid #E5E7EB;border-radius:6px;padding:6px 10px;font-size:13px;">
          <button type="button" id="addIncomeCatBtn" class="btn btn-primary" style="font-size:13px;padding:6px 14px;">追加</button>
        </div>
      </div>
      <div>
        <h3 style="font-size:13px;font-weight:bold;color:#EF4444;margin-bottom:6px;">支出科目（経費）</h3>
        <div id="expenseCatList">${renderList('expense')}</div>
        <div style="display:flex;gap:8px;margin-top:8px;">
          <input type="text" id="newExpenseCat" placeholder="新しい経費科目名" maxlength="20"
            style="flex:1;border:1px solid #E5E7EB;border-radius:6px;padding:6px 10px;font-size:13px;">
          <button type="button" id="addExpenseCatBtn" class="btn btn-primary" style="font-size:13px;padding:6px 14px;">追加</button>
        </div>
      </div>
      <div id="catMsg" class="form-message" style="display:none;margin-top:8px;"></div>
    `;

    // 削除ボタン
    section.querySelectorAll('.del-cat-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const type = btn.dataset.type;
        const cat  = btn.dataset.cat;
        const cats2 = Master.getCategories();
        cats2[type] = cats2[type].filter(c => c !== cat);
        Master.saveCategories(cats2);
        this._renderCategorySection();
        this._renderAllocationSection();
        this._rebuildCategoryOptions();
      });
    });

    // 追加ボタン（収入）
    document.getElementById('addIncomeCatBtn').addEventListener('click', () => {
      const val = document.getElementById('newIncomeCat').value.trim();
      if (!val) { this._showMsg('catMsg', '⚠️ 科目名を入力してください', 'error'); return; }
      const cats2 = Master.getCategories();
      if (cats2.income.includes(val)) { this._showMsg('catMsg', '⚠️ 既に存在する科目名です', 'error'); return; }
      cats2.income.push(val);
      Master.saveCategories(cats2);
      document.getElementById('newIncomeCat').value = '';
      this._renderCategorySection();
      this._rebuildCategoryOptions();
      this._showMsg('catMsg', `✅「${val}」を追加しました`, 'success');
      this._saveMasterToGas();
    });

    // 追加ボタン（支出）
    document.getElementById('addExpenseCatBtn').addEventListener('click', () => {
      const val = document.getElementById('newExpenseCat').value.trim();
      if (!val) { this._showMsg('catMsg', '⚠️ 科目名を入力してください', 'error'); return; }
      const cats2 = Master.getCategories();
      if (cats2.expense.includes(val)) { this._showMsg('catMsg', '⚠️ 既に存在する科目名です', 'error'); return; }
      cats2.expense.push(val);
      Master.saveCategories(cats2);
      document.getElementById('newExpenseCat').value = '';
      this._renderCategorySection();
      this._rebuildCategoryOptions();
      this._renderAllocationSection();
      this._showMsg('catMsg', `✅「${val}」を追加しました`, 'success');
      this._saveMasterToGas();
    });
  },

  /* ─── 支払方法管理 ──────────────────── */

  _injectPaymentMethodSection() {
    if (document.getElementById('paymentMethodSection')) {
      this._renderPaymentMethodSection();
      return;
    }
    const danger = document.querySelector('.danger-zone');
    const section = document.createElement('div');
    section.id        = 'paymentMethodSection';
    section.className = 'card form-card';
    section.style.cssText = 'margin-top:16px;';
    danger.parentNode.insertBefore(section, danger);
    this._renderPaymentMethodSection();
  },

  _renderPaymentMethodSection() {
    const section = document.getElementById('paymentMethodSection');
    if (!section) return;
    const methods  = Master.getPaymentMethods();
    const defaults = ['cash', 'card'];

    const listHtml = methods.map(m => {
      const isDef = defaults.includes(m.value);
      return `
        <div style="display:flex;align-items:center;gap:8px;padding:4px 0;border-bottom:1px solid #F9FAFB;">
          <span style="flex:1;font-size:13px;">${this._esc(m.label)}</span>
          <span style="font-size:11px;color:#9CA3AF;">${this._esc(m.value)}</span>
          ${isDef ? '<span style="font-size:11px;color:#9CA3AF;">既定</span>' : `
            <button class="del-pm-btn" data-val="${this._esc(m.value)}"
              style="background:none;border:1px solid #FECACA;border-radius:6px;padding:2px 8px;
                     font-size:11px;cursor:pointer;color:#EF4444;">削除</button>`}
        </div>
      `;
    }).join('');

    section.innerHTML = `
      <h2 class="section-title">支払い方法の管理</h2>
      <div id="pmList" style="margin-bottom:12px;">${listHtml}</div>
      <div style="display:flex;gap:8px;">
        <input type="text" id="newPmLabel" placeholder="表示名（例：PayPay）" maxlength="20"
          style="flex:1;border:1px solid #E5E7EB;border-radius:6px;padding:6px 10px;font-size:13px;">
        <input type="text" id="newPmValue" placeholder="識別子（例：paypay）" maxlength="20"
          style="width:110px;border:1px solid #E5E7EB;border-radius:6px;padding:6px 10px;font-size:13px;">
        <button type="button" id="addPmBtn" class="btn btn-primary" style="font-size:13px;padding:6px 14px;">追加</button>
      </div>
      <p style="font-size:11px;color:#9CA3AF;margin-top:4px;">識別子は半角英数字（小文字）で入力してください</p>
      <div id="pmMsg" class="form-message" style="display:none;margin-top:8px;"></div>
    `;

    // 削除
    section.querySelectorAll('.del-pm-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const val      = btn.dataset.val;
        const methods2 = Master.getPaymentMethods().filter(m => m.value !== val);
        Master.savePaymentMethods(methods2);
        this._renderPaymentMethodSection();
        this._saveMasterToGas();
      });
    });

    // 追加
    document.getElementById('addPmBtn').addEventListener('click', () => {
      const label = document.getElementById('newPmLabel').value.trim();
      const value = document.getElementById('newPmValue').value.trim().toLowerCase().replace(/[^a-z0-9_]/g, '');
      if (!label || !value) {
        this._showMsg('pmMsg', '⚠️ 表示名と識別子を入力してください', 'error');
        return;
      }
      const methods2 = Master.getPaymentMethods();
      if (methods2.find(m => m.value === value)) {
        this._showMsg('pmMsg', '⚠️ 既に存在する識別子です', 'error');
        return;
      }
      methods2.push({ value, label });
      Master.savePaymentMethods(methods2);
      document.getElementById('newPmLabel').value = '';
      document.getElementById('newPmValue').value = '';
      this._renderPaymentMethodSection();
      this._rebuildPaymentMethodOptions();
      this._showMsg('pmMsg', `✅「${label}」を追加しました`, 'success');
      this._saveMasterToGas();
    });
  },

  /* ─── タグ管理 ──────────────────────── */

  _injectTagSection() {
    if (document.getElementById('tagMgmtSection')) {
      this._renderTagSection();
      return;
    }
    const danger = document.querySelector('.danger-zone');
    const section = document.createElement('div');
    section.id        = 'tagMgmtSection';
    section.className = 'card form-card';
    section.style.cssText = 'margin-top:16px;';
    danger.parentNode.insertBefore(section, danger);
    this._renderTagSection();
  },

  _renderTagSection() {
    const section = document.getElementById('tagMgmtSection');
    if (!section) return;
    const tags    = Master.getTags();
    const tagHtml = tags.map(tag => `
      <span style="display:inline-flex;align-items:center;gap:4px;background:#EDE9FE;
                   color:#5B21B6;border:1px solid #C4B5FD;border-radius:20px;
                   padding:3px 10px;font-size:12px;margin:3px;">
        ${this._esc(tag)}
        <button class="del-tag-btn" data-tag="${this._esc(tag)}"
          style="background:none;border:none;cursor:pointer;color:#7C3AED;font-size:14px;line-height:1;">×</button>
      </span>
    `).join('');

    section.innerHTML = `
      <h2 class="section-title">タグの管理</h2>
      <div id="tagList" style="margin-bottom:12px;min-height:28px;">${tagHtml || '<p style="color:#9CA3AF;font-size:13px;">タグがありません</p>'}</div>
      <div style="display:flex;gap:8px;">
        <input type="text" id="newTagInput" placeholder="新しいタグ名" maxlength="20"
          style="flex:1;border:1px solid #E5E7EB;border-radius:6px;padding:6px 10px;font-size:13px;">
        <button type="button" id="addTagBtn" class="btn btn-primary" style="font-size:13px;padding:6px 14px;">追加</button>
      </div>
      <div id="tagMsg" class="form-message" style="display:none;margin-top:8px;"></div>
    `;

    section.querySelectorAll('.del-tag-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const tag   = btn.dataset.tag;
        const tags2 = Master.getTags().filter(t => t !== tag);
        Master.saveTags(tags2);
        this._renderTagSection();
        this._rebuildTagSuggestions();
        this._saveMasterToGas();
      });
    });

    document.getElementById('addTagBtn').addEventListener('click', () => {
      const val = document.getElementById('newTagInput').value.trim();
      if (!val) { this._showMsg('tagMsg', '⚠️ タグ名を入力してください', 'error'); return; }
      const tags2 = Master.getTags();
      if (tags2.includes(val)) { this._showMsg('tagMsg', '⚠️ 既に存在するタグです', 'error'); return; }
      tags2.push(val);
      Master.saveTags(tags2);
      document.getElementById('newTagInput').value = '';
      this._renderTagSection();
      this._rebuildTagSuggestions();
      this._showMsg('tagMsg', `✅「${val}」を追加しました`, 'success');
      this._saveMasterToGas();
    });
  },

  /* ─── メニュー管理 ──────────────────── */

  _injectMenuSection() {
    if (document.getElementById('menuMgmtSection')) {
      this._renderMenuSection();
      return;
    }
    const danger = document.querySelector('.danger-zone');
    const section = document.createElement('div');
    section.id        = 'menuMgmtSection';
    section.className = 'card form-card';
    section.style.cssText = 'margin-top:16px;';
    danger.parentNode.insertBefore(section, danger);
    this._renderMenuSection();
  },

  _renderMenuSection() {
    const section = document.getElementById('menuMgmtSection');
    if (!section) return;
    const menus = Master.getMenus();

    const listHtml = menus.map(m => `
      <div style="display:flex;align-items:center;gap:8px;padding:8px 0;border-bottom:1px solid var(--border);">
        <span style="flex:1;font-size:13px;font-weight:600;">${this._esc(m.name)}</span>
        <span style="font-size:12px;color:var(--text-sub);white-space:nowrap;">${m.duration || 60}分</span>
        <span style="font-size:13px;font-family:var(--font-serif);color:var(--accent);white-space:nowrap;">¥${Number(m.price).toLocaleString()}</span>
        <button class="del-menu-btn" data-id="${this._esc(m.id)}"
          style="background:none;border:1px solid #FECACA;border-radius:6px;padding:2px 8px;font-size:11px;cursor:pointer;color:#EF4444;flex-shrink:0;">削除</button>
      </div>`).join('');

    const inputStyle = 'border:1.5px solid var(--border-normal);border-radius:var(--radius-sm);padding:8px 10px;font-size:13px;font-family:var(--font-sans);width:100%;';

    section.innerHTML = `
      <h2 class="section-title">メニュー管理</h2>
      <p style="font-size:12px;color:var(--text-sub);margin-bottom:12px;">記録フォームで選択すると金額・科目を自動入力します。施術時間は予約フォームの空き枠計算に使用します。</p>
      <div id="menuList">${listHtml || '<p style="font-size:13px;color:var(--text-light);">メニューがありません</p>'}</div>
      <div style="display:grid;grid-template-columns:1fr 90px 90px auto;gap:8px;margin-top:12px;align-items:end;">
        <div>
          <label style="font-size:11px;color:var(--text-sub);display:block;margin-bottom:4px;">メニュー名</label>
          <input type="text" id="newMenuName" placeholder="例：ジェルネイル" maxlength="30" style="${inputStyle}">
        </div>
        <div>
          <label style="font-size:11px;color:var(--text-sub);display:block;margin-bottom:4px;">金額（円）</label>
          <input type="number" id="newMenuPrice" placeholder="7000" min="1" step="1" style="${inputStyle}">
        </div>
        <div>
          <label style="font-size:11px;color:var(--text-sub);display:block;margin-bottom:4px;">施術時間（分）</label>
          <input type="number" id="newMenuDuration" placeholder="60" min="15" step="15" value="60" style="${inputStyle}">
        </div>
        <button type="button" id="addMenuBtn" class="btn btn-primary" style="font-size:13px;padding:8px 14px;white-space:nowrap;align-self:end;">追加</button>
      </div>
      <div id="menuMsg" class="form-message" style="display:none;margin-top:8px;"></div>
    `;

    section.querySelectorAll('.del-menu-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const menus2 = Master.getMenus().filter(m => m.id !== btn.dataset.id);
        Master.saveMenus(menus2);
        this._renderMenuSection();
        this._rebuildMenuOptions();
        this._saveMasterToGas();
      });
    });

    document.getElementById('addMenuBtn').addEventListener('click', () => {
      const name     = document.getElementById('newMenuName').value.trim();
      const price    = Number(document.getElementById('newMenuPrice').value);
      const duration = Number(document.getElementById('newMenuDuration').value) || 60;
      if (!name)            { this._showMsg('menuMsg', '⚠️ メニュー名を入力してください', 'error'); return; }
      if (!price || price <= 0) { this._showMsg('menuMsg', '⚠️ 金額を正しく入力してください', 'error'); return; }
      const menus2 = Master.getMenus();
      if (menus2.find(m => m.name === name)) { this._showMsg('menuMsg', '⚠️ 同じ名前のメニューが既にあります', 'error'); return; }
      menus2.push({ id: `menu_${Date.now()}`, name, price, duration, category: '売上' });
      Master.saveMenus(menus2);
      document.getElementById('newMenuName').value     = '';
      document.getElementById('newMenuPrice').value    = '';
      document.getElementById('newMenuDuration').value = '60';
      this._renderMenuSection();
      this._rebuildMenuOptions();
      this._saveMasterToGas();
      this._showMsg('menuMsg', `✅「${name}」を追加しました`, 'success');
    });
  },

  /* ─── 営業時間・定休日設定 ──────────────── */

  _injectBusinessHoursSection() {
    if (document.getElementById('businessHoursSection')) {
      this._renderBusinessHoursSection();
      return;
    }
    const danger = document.querySelector('.danger-zone');
    const section = document.createElement('div');
    section.id        = 'businessHoursSection';
    section.className = 'card form-card';
    section.style.cssText = 'margin-top:16px;';
    danger.parentNode.insertBefore(section, danger);
    this._renderBusinessHoursSection();
  },

  _renderBusinessHoursSection() {
    const section = document.getElementById('businessHoursSection');
    if (!section) return;

    const hours  = Master.getBusinessHours();
    const window = Master.getBookingWindowDays();

    const days = [
      { key: 'mon', label: '月' },
      { key: 'tue', label: '火' },
      { key: 'wed', label: '水' },
      { key: 'thu', label: '木' },
      { key: 'fri', label: '金' },
      { key: 'sat', label: '土' },
      { key: 'sun', label: '日' },
    ];

    const inputStyle = 'border:1.5px solid var(--border-normal);border-radius:var(--radius-sm);padding:6px 8px;font-size:13px;font-family:var(--font-sans);width:80px;';

    const rowsHtml = days.map(d => {
      const h       = hours[d.key];
      const isClosed = !h;
      return `
        <div class="bh-row" data-day="${d.key}"
          style="display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid var(--border);">
          <span style="width:20px;font-weight:700;font-size:13px;text-align:center;">${d.label}</span>
          <label style="display:flex;align-items:center;gap:5px;font-size:13px;cursor:pointer;color:var(--text-sub);">
            <input type="checkbox" class="bh-closed" data-day="${d.key}" ${isClosed ? 'checked' : ''}
              style="width:15px;height:15px;accent-color:var(--accent);">
            定休日
          </label>
          <div class="bh-times" style="display:${isClosed ? 'none' : 'flex'};align-items:center;gap:6px;">
            <input type="time" class="bh-open" data-day="${d.key}" value="${h?.open || '10:00'}" style="${inputStyle}">
            <span style="font-size:12px;color:var(--text-sub);">〜</span>
            <input type="time" class="bh-close" data-day="${d.key}" value="${h?.close || '19:00'}" style="${inputStyle}">
          </div>
          <span class="bh-closed-label" style="display:${isClosed ? 'inline' : 'none'};font-size:12px;color:var(--text-light);">定休日</span>
        </div>`;
    }).join('');

    section.innerHTML = `
      <h2 class="section-title">営業時間・予約設定</h2>
      <p style="font-size:12px;color:var(--text-sub);margin-bottom:14px;">予約フォームの空き枠計算と受付可能日に使用します</p>

      <h3 style="font-size:12px;font-weight:700;color:var(--text);margin-bottom:8px;">営業時間・定休日</h3>
      <div id="bhRows">${rowsHtml}</div>

      <div style="display:flex;align-items:center;gap:10px;margin-top:16px;padding-top:14px;border-top:1px solid var(--border);">
        <label style="font-size:13px;font-weight:700;white-space:nowrap;">予約受付日数</label>
        <input type="number" id="bookingWindowDays" value="${window}" min="7" max="365" step="1"
          style="${inputStyle} width:70px;">
        <span style="font-size:12px;color:var(--text-sub);">日先まで受け付ける（7〜365）</span>
      </div>

      <div id="bhMsg" class="form-message" style="display:none;margin-top:10px;"></div>
      <button type="button" id="saveBusinessHoursBtn" class="btn btn-primary" style="margin-top:14px;">保存する</button>
    `;

    // 定休日チェックで時間入力の表示切替
    section.querySelectorAll('.bh-closed').forEach(cb => {
      cb.addEventListener('change', () => {
        const row = section.querySelector(`.bh-row[data-day="${cb.dataset.day}"]`);
        row.querySelector('.bh-times').style.display        = cb.checked ? 'none' : 'flex';
        row.querySelector('.bh-closed-label').style.display = cb.checked ? 'inline' : 'none';
      });
    });

    // 保存
    document.getElementById('saveBusinessHoursBtn').addEventListener('click', () => {
      const newHours = {};
      days.forEach(d => {
        const closed = section.querySelector(`.bh-closed[data-day="${d.key}"]`).checked;
        if (closed) {
          newHours[d.key] = null;
        } else {
          const open  = section.querySelector(`.bh-open[data-day="${d.key}"]`).value  || '10:00';
          const close = section.querySelector(`.bh-close[data-day="${d.key}"]`).value || '19:00';
          if (open >= close) {
            this._showMsg('bhMsg', `⚠️ ${d.label}曜日の終了時間は開始より後に設定してください`, 'error');
            return;
          }
          newHours[d.key] = { open, close };
        }
      });
      const windowDays = Number(document.getElementById('bookingWindowDays').value) || 60;
      Master.saveBusinessHours(newHours);
      Master.saveBookingWindowDays(windowDays);
      this._saveMasterToGas();
      this._showMsg('bhMsg', '✅ 営業時間設定を保存しました', 'success');
    });
  },

  /* ─── 予約フォームQRコード ─────────────────── */

  _injectQrSection() {
    if (document.getElementById('qrSection')) {
      this._renderQrSection();
      return;
    }
    const danger = document.querySelector('.danger-zone');
    const section = document.createElement('div');
    section.id        = 'qrSection';
    section.className = 'card form-card';
    section.style.cssText = 'margin-top:16px;';
    danger.parentNode.insertBefore(section, danger);
    this._renderQrSection();
  },

  _renderQrSection() {
    const section = document.getElementById('qrSection');
    if (!section) return;
    const saved = Storage.getSettings().bookingUrl || '';
    section.innerHTML = `
      <h2 class="section-title">予約フォーム QRコード</h2>
      <p style="font-size:12px;color:var(--text-sub);margin-bottom:14px;">
        booking.html のURLを入力すると、ショップカード用QRコードを生成できます。
      </p>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
        <input type="url" id="bookingUrlInput" value="${UI._esc(saved)}"
          placeholder="https://example.com/booking.html"
          style="flex:1;min-width:200px;border:1.5px solid var(--border-normal);border-radius:var(--radius-sm);
                 padding:8px 12px;font-size:13px;font-family:var(--font-sans);">
        <button type="button" id="generateQrBtn" class="btn btn-primary">QR生成</button>
        <button type="button" id="copyUrlBtn" class="btn btn-secondary"
          style="border:1.5px solid var(--border-normal);background:none;color:var(--text);">URLコピー</button>
      </div>
      <div id="qrCanvas" style="margin-top:16px;display:flex;flex-direction:column;align-items:center;gap:10px;"></div>
    `;

    document.getElementById('generateQrBtn').addEventListener('click', () => {
      const url = document.getElementById('bookingUrlInput').value.trim();
      if (!url) { UI._showToast('URLを入力してください', 'warn'); return; }
      // 設定に保存
      const s = Storage.getSettings();
      Storage.saveSettings({ ...s, bookingUrl: url });
      // QR生成
      const container = document.getElementById('qrCanvas');
      container.innerHTML = '';
      const canvas = document.createElement('canvas');
      container.appendChild(canvas);
      if (typeof QRCode !== 'undefined') {
        QRCode.toCanvas(canvas, url, { width: 200, margin: 2, color: { dark: '#1C1C1E', light: '#FFFFFF' } }, err => {
          if (err) { container.innerHTML = '<p style="color:#EF4444;font-size:13px;">QR生成に失敗しました</p>'; }
          else {
            const dl = document.createElement('a');
            dl.href = canvas.toDataURL('image/png');
            dl.download = 'booking-qr.png';
            dl.className = 'btn btn-secondary';
            dl.style.cssText = 'border:1.5px solid var(--border-normal);background:none;color:var(--text);padding:6px 18px;border-radius:var(--radius-sm);font-size:13px;text-decoration:none;display:inline-block;';
            dl.textContent = '📥 PNG保存';
            container.appendChild(dl);
          }
        });
      } else {
        container.innerHTML = '<p style="font-size:13px;color:var(--text-sub);">QRライブラリの読み込み中です。ページを再読み込みしてください。</p>';
      }
    });

    document.getElementById('copyUrlBtn').addEventListener('click', () => {
      const url = document.getElementById('bookingUrlInput').value.trim();
      if (!url) { UI._showToast('URLを入力してください', 'warn'); return; }
      navigator.clipboard.writeText(url).then(() => UI._showToast('✅ URLをコピーしました', 'success'))
        .catch(() => UI._showToast('コピーに失敗しました', 'warn'));
    });

    // 保存済みURLがあれば自動生成
    if (saved) document.getElementById('generateQrBtn').click();
  },

  /** 支払方法の radio ボタンを動的に再構築 */
  _rebuildPaymentMethodOptions() {
    const group = document.getElementById('paymentMethodGroup');
    if (!group) return;
    const methods = Master.getPaymentMethods();
    const checked = document.querySelector('input[name="paymentMethod"]:checked')?.value || 'cash';

    group.innerHTML = '';
    methods.forEach((m, i) => {
      const icon  = m.value === 'cash' ? '💴' : (m.value === 'card' ? '💳' : '💰');
      const label = document.createElement('label');
      label.className = 'payment-option';
      label.id = `payOpt${m.value}`;
      label.innerHTML = `
        <input type="radio" name="paymentMethod" value="${this._esc(m.value)}" ${i === 0 || m.value === checked ? 'checked' : ''}>
        <span>${icon} ${this._esc(m.label)}</span>
      `;
      group.appendChild(label);
    });

    // イベント再登録
    group.querySelectorAll('input[name="paymentMethod"]').forEach(radio => {
      radio.addEventListener('change', () => {
        this._updatePaymentStyle();
        this._tryAutoRule();
      });
    });
    this._updatePaymentStyle();
  },

  _updatePaymentStyle() {
    const checked = document.querySelector('input[name="paymentMethod"]:checked');
    // 既存のselectedクラスをリセット
    document.querySelectorAll('.payment-option').forEach(el => el.classList.remove('selected'));
    if (checked) {
      checked.closest('.payment-option')?.classList.add('selected');
    }
  },

  /** マスタデータをGASに非同期で保存（失敗しても無視） */
  _saveMasterToGas() {
    if (GasAPI.isConfigured() && navigator.onLine) {
      GasAPI.saveMaster(Master.getAllData()).catch(e => {
        console.warn('[Master] GAS保存失敗:', e.message);
      });
    }
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
        this._updateAllocationDisplay();
        const menuRow = document.getElementById('menuSelectRow');
        if (menuRow) menuRow.style.display = this.currentType === 'income' ? 'block' : 'none';
        const menuSel = document.getElementById('inputMenuSelect');
        if (menuSel) menuSel.value = '';
      });
    });

    this._rebuildPaymentMethodOptions();
    this._rebuildCategoryOptions();
    this._rebuildMenuOptions();

    // メニュー選択 → 金額・科目・説明を自動入力
    const menuSel = document.getElementById('inputMenuSelect');
    if (menuSel) {
      menuSel.addEventListener('change', () => {
        const opt = menuSel.options[menuSel.selectedIndex];
        if (!opt.value) return;
        document.getElementById('inputAmount').value      = opt.dataset.price;
        document.getElementById('inputDescription').value = opt.dataset.name;
        const catSel = document.getElementById('inputCategory');
        if (catSel) catSel.value = opt.dataset.category;
        this._updateAllocationDisplay();
      });
    }

    document.getElementById('inputAmount').addEventListener('input', () => {
      this._updateAllocationDisplay();
      this._tryAutoRule();
    });

    document.getElementById('inputCategory').addEventListener('change', () => {
      this._updateAllocationDisplay();
    });

    document.getElementById('recordForm').addEventListener('submit', e => {
      e.preventDefault();
      this._submitRecord();
    });

    // タグ候補のクリック登録
    this._rebuildTagSuggestions();

    // 領収書プレビュー
    const receiptInput = document.getElementById('inputReceipt');
    if (receiptInput) {
      receiptInput.addEventListener('change', () => {
        const file = receiptInput.files[0];
        const preview  = document.getElementById('receiptPreview');
        const previewImg = document.getElementById('receiptPreviewImg');
        const previewName = document.getElementById('receiptPreviewName');
        if (file) {
          previewName.textContent = file.name;
          if (file.type.startsWith('image/')) {
            const url = URL.createObjectURL(file);
            previewImg.src = url;
            previewImg.style.display = 'block';
          } else {
            previewImg.style.display = 'none';
          }
          preview.style.display = 'flex';
          preview.style.alignItems = 'center';
        } else {
          preview.style.display = 'none';
        }
      });
    }
  },

  /** 按分表示を更新 */
  _updateAllocationDisplay() {
    const display  = document.getElementById('allocationDisplay');
    const textEl   = document.getElementById('allocationText');
    if (!display || !textEl) return;

    if (this.currentType !== 'expense') {
      display.style.display = 'none';
      return;
    }

    const cat    = document.getElementById('inputCategory').value;
    const amount = Number(document.getElementById('inputAmount').value) || 0;

    if (!cat || !Allocation.hasAllocation(cat)) {
      display.style.display = 'none';
      return;
    }

    const ratio    = Allocation.getRatio(cat);
    const adjusted = Allocation.applyRatio(amount, cat);
    textEl.textContent =
      `家事按分あり：事業使用割合 ${ratio}% → 経費計上額 ${Format.currency(adjusted)}（集計・CSVはこの金額で出力されます）`;
    display.style.display = 'block';
  },

  /** 自動登録ルールを適用 */
  _tryAutoRule() {
    const pm     = document.querySelector('input[name="paymentMethod"]:checked');
    const amount = Number(document.getElementById('inputAmount').value) || 0;
    if (!pm || amount <= 0) return;

    const matched = AutoRules.match(this.currentType, pm.value, amount);
    if (!matched) return;

    const catSel = document.getElementById('inputCategory');
    const options = [...catSel.options].map(o => o.value);
    if (options.includes(matched)) {
      catSel.value = matched;
      this._updateAllocationDisplay();
    }
  },

  /** タグ候補チップを再構築 */
  _rebuildTagSuggestions() {
    const container = document.getElementById('tagSuggestions');
    if (!container) return;
    container.innerHTML = '';
    const tags = Master.getTags();
    tags.forEach(tag => {
      const chip = document.createElement('button');
      chip.type = 'button';
      chip.textContent = tag;
      chip.style.cssText = [
        'background:#EDE9FE',
        'color:#5B21B6',
        'border:1px solid #C4B5FD',
        'border-radius:20px',
        'padding:2px 10px',
        'font-size:12px',
        'cursor:pointer',
      ].join(';');
      chip.addEventListener('click', () => {
        const input = document.getElementById('inputTags');
        const current = input.value.split(',').map(s => s.trim()).filter(Boolean);
        if (!current.includes(tag)) {
          current.push(tag);
          input.value = current.join(', ');
        }
      });
      container.appendChild(chip);
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

    // タグをセット
    const tagsInput = document.getElementById('inputTags');
    if (tagsInput) tagsInput.value = Array.isArray(t.tags) ? t.tags.join(', ') : '';

    // 按分表示を更新
    this._updateAllocationDisplay();

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
    const masterCats = Master.getCategories();
    const cats = this.currentType === 'income' ? masterCats.income : masterCats.expense;
    sel.innerHTML = '<option value="">-- 科目を選んでください --</option>';
    cats.forEach(c => {
      const opt = document.createElement('option');
      opt.value = c;
      opt.textContent = c;
      sel.appendChild(opt);
    });
    if (this.currentType === 'income') sel.value = masterCats.income[0] || '売上';
  },

  _rebuildMenuOptions() {
    const sel = document.getElementById('inputMenuSelect');
    if (!sel) return;
    const menus = Master.getMenus();
    sel.innerHTML = '<option value="">-- メニューから選択（省略可） --</option>';
    menus.forEach(m => {
      const opt = document.createElement('option');
      opt.value = m.id;
      opt.textContent = `${m.name}（¥${Number(m.price).toLocaleString()} / ${m.duration || 60}分）`;
      opt.dataset.price    = m.price;
      opt.dataset.category = m.category;
      opt.dataset.name     = m.name;
      opt.dataset.duration = m.duration || 60;
      sel.appendChild(opt);
    });
  },

  async _submitRecord() {
    const date          = document.getElementById('inputDate').value;
    const amountRaw     = document.getElementById('inputAmount').value;
    const category      = document.getElementById('inputCategory').value;
    const description   = document.getElementById('inputDescription').value.trim();
    const paymentMethod = document.querySelector('input[name="paymentMethod"]:checked').value;
    const tagsRaw       = (document.getElementById('inputTags')?.value || '').trim();
    const tags          = tagsRaw
      ? tagsRaw.split(',').map(s => s.trim()).filter(Boolean)
      : [];

    if (!date || !amountRaw || !category || !description) {
      this._showMsg('formMessage', '⚠️ 未入力の項目があります', 'error');
      return;
    }
    const amount = Number(amountRaw);
    if (!Number.isFinite(amount) || amount <= 0) {
      this._showMsg('formMessage', '⚠️ 金額は1円以上の数値を入力してください', 'error');
      return;
    }

    const payload   = { date, amount, category, description, paymentMethod, type: this.currentType, tags };
    const submitBtn = document.querySelector('#recordForm button[type="submit"]');
    const origLabel = submitBtn ? submitBtn.textContent : '';
    if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = '保存中...'; }

    try {
      // 領収書アップロード（新規のみ・編集時は既存IDを維持）
      let receiptId = '';
      const receiptInput = document.getElementById('inputReceipt');
      if (!this._editMode && receiptInput && receiptInput.files[0] && GasAPI.isConfigured()) {
        const statusEl = document.getElementById('receiptUploadStatus');
        if (statusEl) statusEl.textContent = '📤 領収書をアップロード中...';
        try {
          receiptId = await GasAPI.uploadReceipt(receiptInput.files[0]);
        } catch (err) {
          console.warn('[Receipt] アップロード失敗:', err.message);
          this._showToast(`⚠️ 領収書のアップロードに失敗しましたが、取引は保存されます`, 'warn');
        }
        if (statusEl) statusEl.textContent = '';
      } else if (this._editMode && this._editingId) {
        // 編集時は既存のreceiptIdを引き継ぐ
        const existing = Data.getById(this._editingId);
        receiptId = existing?.receiptId || '';
      }
      if (receiptId) payload.receiptId = receiptId;

      if (this._editMode && this._editingId) {
        await Data.update({ id: this._editingId, ...payload });
        this._cancelEdit();
        this._showMsg('formMessage', '✅ 内容を更新しました！', 'success');
      } else {
        const label = this.currentType === 'income' ? '売上' : '経費';
        await Data.add(payload);
        document.getElementById('inputAmount').value      = '';
        document.getElementById('inputDescription').value = '';
        document.getElementById('inputDate').value        = this._todayStr();
        if (document.getElementById('inputTags')) document.getElementById('inputTags').value = '';
        if (receiptInput) {
          receiptInput.value = '';
          document.getElementById('receiptPreview').style.display = 'none';
        }
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

  initLedger() {
    document.getElementById('ledgerYear').addEventListener('change', e => {
      this._updateLedger(Number(e.target.value));
    });
    document.getElementById('ledgerMonth').addEventListener('change', () => {
      const y = Number(document.getElementById('ledgerYear').value);
      this._updateLedger(y);
    });
    document.getElementById('ledgerAccount').addEventListener('change', () => {
      const y    = Number(document.getElementById('ledgerYear').value);
      const mVal = document.getElementById('ledgerMonth').value;
      const m    = mVal ? Number(mVal) : null;
      let list   = m ? Data.getByYearMonth(y, m) : Data.getByYear(y);
      list = [...list].sort((a, b) => a.date.localeCompare(b.date));
      this._renderGeneralLedger(document.getElementById('ledgerAccount').value, list);
    });
    document.getElementById('printJournalBtn').addEventListener('click', () => {
      this._printSection('journalCard', '仕訳帳');
    });
    document.getElementById('printLedgerBtn').addEventListener('click', () => {
      this._printSection('generalLedgerCard', '総勘定元帳');
    });
    // 領収書モーダルの閉じるボタン
    document.getElementById('receiptModalClose').addEventListener('click', () => {
      document.getElementById('receiptModal').style.display = 'none';
    });
    document.getElementById('receiptModal').addEventListener('click', e => {
      if (e.target === document.getElementById('receiptModal')) {
        document.getElementById('receiptModal').style.display = 'none';
      }
    });
  },

  _printSection(cardId, title) {
    const card = document.getElementById(cardId);
    if (!card) return;
    const w = window.open('', '_blank', 'width=900,height=700');
    w.document.write(`
      <!DOCTYPE html>
      <html lang="ja">
      <head>
        <meta charset="UTF-8">
        <title>${title}</title>
        <style>
          body { font-family: 'Hiragino Sans', sans-serif; font-size: 12px; padding: 20px; }
          table { border-collapse: collapse; width: 100%; }
          th, td { border: 1px solid #ccc; padding: 6px 10px; text-align: left; }
          th { background: #f5f5f5; font-weight: bold; }
          @media print { button { display: none !important; } }
        </style>
      </head>
      <body>
        <h2>${title}</h2>
        ${card.querySelector('table')?.outerHTML || ''}
        <script>window.onload = function(){ window.print(); }<\/script>
      </body>
      </html>
    `);
    w.document.close();
  },

  /* ─── フィルター ─────────────────────── */

  initFilters() {
    ['filterMonth', 'filterType', 'filterCategory', 'filterTag'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', () => this._applyFilters());
    });
  },

  /* ─── 帳簿タブ ───────────────────────── */

  renderLedger() {
    const years = Calculator.getAvailableYears();
    const yearSel = document.getElementById('ledgerYear');
    const prev = yearSel.value || String(new Date().getFullYear());
    yearSel.innerHTML = '';
    years.forEach(y => {
      const opt = document.createElement('option');
      opt.value = y;
      opt.textContent = `${y}年`;
      yearSel.appendChild(opt);
    });
    yearSel.value = prev || String(years[0]);

    // 総勘定元帳の科目セレクター
    this._rebuildLedgerAccountOptions();

    this._updateLedger(Number(yearSel.value));
  },

  _rebuildLedgerAccountOptions() {
    const sel = document.getElementById('ledgerAccount');
    if (!sel) return;
    const prev = sel.value;
    sel.innerHTML = '';
    const cats = Master.getCategories();
    const { initialCash, initialBank } = Storage.getSettings();

    const allAccounts = [
      { value: '現金', label: '現金' },
      { value: '普通預金', label: '普通預金（カード・振込）' },
      ...cats.income.map(c => ({ value: c, label: `${c}（収益）` })),
      ...cats.expense.map(c => ({ value: c, label: `${c}（費用）` })),
    ];
    allAccounts.forEach(a => {
      const opt = document.createElement('option');
      opt.value = a.value;
      opt.textContent = a.label;
      sel.appendChild(opt);
    });
    sel.value = prev || '現金';
  },

  _updateLedger(year) {
    const monthVal = document.getElementById('ledgerMonth')?.value;
    const month    = monthVal ? Number(monthVal) : null;

    let list = month
      ? Data.getByYearMonth(year, month)
      : Data.getByYear(year);
    list = [...list].sort((a, b) => a.date.localeCompare(b.date));

    this._renderJournal(list);
    this._renderGeneralLedger(document.getElementById('ledgerAccount')?.value || '現金', list);
  },

  /** 仕訳帳 */
  _renderJournal(list) {
    const tbody  = document.getElementById('journalBody');
    const empty  = document.getElementById('journalEmpty');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (list.length === 0) {
      if (empty) empty.style.display = 'block';
      return;
    }
    if (empty) empty.style.display = 'none';

    list.forEach(t => {
      const isIncome  = t.type === 'income';
      const isCash    = t.paymentMethod === 'cash';
      const amount    = isIncome
        ? Number(t.amount)
        : Allocation.applyRatio(Number(t.amount), t.category);
      const hasAlloc  = !isIncome && Allocation.hasAllocation(t.category);

      // 複式簿記の仕訳
      // 収入・現金: 借方=現金 / 貸方=売上
      // 収入・カード: 借方=普通預金 / 貸方=売上
      // 支出・現金: 借方=[費用科目] / 貸方=現金
      // 支出・カード: 借方=[費用科目] / 貸方=普通預金
      const debit  = isIncome ? (isCash ? '現金' : '普通預金') : t.category;
      const credit = isIncome ? t.category : (isCash ? '現金' : '普通預金');

      const amtNote = hasAlloc
        ? `${Format.currency(amount)}<small style="font-size:10px;color:#9CA3AF;"> (${Allocation.getRatio(t.category)}%)</small>`
        : Format.currency(amount);

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="white-space:nowrap;">${Format.date(t.date)}</td>
        <td>${this._esc(t.description)}</td>
        <td style="color:#3B82F6;">${this._esc(debit)}</td>
        <td style="color:#EF4444;">${this._esc(credit)}</td>
        <td style="text-align:right;">${amtNote}</td>
      `;
      tbody.appendChild(tr);
    });
  },

  /** 総勘定元帳 */
  _renderGeneralLedger(account, allList) {
    const tbody = document.getElementById('ledgerBody');
    const empty = document.getElementById('ledgerEmpty');
    if (!tbody) return;
    tbody.innerHTML = '';

    const { initialCash, initialBank } = Storage.getSettings();
    const cats = Master.getCategories();
    const isAsset   = account === '現金' || account === '普通預金';
    const isIncomeCat = cats.income.includes(account);
    const isExpCat    = cats.expense.includes(account);

    // この科目に関係する取引を抽出
    const related = allList.filter(t => {
      if (account === '現金') {
        return t.paymentMethod === 'cash';
      }
      if (account === '普通預金') {
        return t.paymentMethod === 'card';
      }
      if (isIncomeCat) return t.type === 'income' && t.category === account;
      if (isExpCat)    return t.type === 'expense' && t.category === account;
      return false;
    });

    if (related.length === 0) {
      if (empty) empty.style.display = 'block';
      return;
    }
    if (empty) empty.style.display = 'none';

    // 期首残高
    let balance = 0;
    if (account === '現金')    balance = Number(initialCash) || 0;
    if (account === '普通預金') balance = Number(initialBank) || 0;

    // 期首行
    const initTr = document.createElement('tr');
    initTr.style.color = '#9CA3AF';
    initTr.innerHTML = `
      <td>―</td>
      <td>期首残高</td>
      <td></td>
      <td></td>
      <td style="text-align:right;font-weight:bold;">${Format.currency(balance)}</td>
    `;
    tbody.appendChild(initTr);

    related.forEach(t => {
      const isIncome = t.type === 'income';
      const isCash   = t.paymentMethod === 'cash';
      const adjAmt   = !isIncome ? Allocation.applyRatio(Number(t.amount), t.category) : Number(t.amount);
      let debitAmt   = 0;
      let creditAmt  = 0;

      if (isAsset) {
        // 資産科目: 収入でお金が入る(借方+)、支出でお金が出る(貸方−)
        if (isIncome) {
          debitAmt  = Number(t.amount);
          balance  += Number(t.amount);
        } else {
          creditAmt = adjAmt;
          balance  -= adjAmt;
        }
      } else if (isIncomeCat) {
        // 収益科目: 売上は貸方に増加
        creditAmt = Number(t.amount);
        balance  += Number(t.amount);
      } else if (isExpCat) {
        // 費用科目: 経費は借方に増加
        debitAmt = adjAmt;
        balance += adjAmt;
      }

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="white-space:nowrap;">${Format.date(t.date)}</td>
        <td>${this._esc(t.description)}</td>
        <td style="text-align:right;color:#3B82F6;">${debitAmt > 0 ? Format.currency(debitAmt) : ''}</td>
        <td style="text-align:right;color:#EF4444;">${creditAmt > 0 ? Format.currency(creditAmt) : ''}</td>
        <td style="text-align:right;font-weight:bold;">${Format.currency(balance)}</td>
      `;
      tbody.appendChild(tr);
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
   AppointmentData モジュール
   予約データのローカルキャッシュ＋GAS同期
   ============================================= */
const AppointmentData = {
  _cache: null,

  getAll() {
    if (this._cache) return this._cache;
    try { this._cache = JSON.parse(localStorage.getItem(APPOINTMENTS_KEY) || '[]'); } catch { this._cache = []; }
    return this._cache;
  },

  saveAll(list) {
    this._cache = list;
    localStorage.setItem(APPOINTMENTS_KEY, JSON.stringify(list));
  },

  getByDate(dateStr) {
    return this.getAll().filter(a => (a.dateTime || '').startsWith(dateStr));
  },

  async loadFromGas(from, to) {
    if (!GasAPI.isConfigured()) return false;
    try {
      const list = await GasAPI.getAppointments(from, to);
      // 既存キャッシュと日付範囲をマージ（単純に全上書き）
      this.saveAll(list);
      return true;
    } catch(e) {
      console.warn('[AppointmentData] load failed:', e.message);
      return false;
    }
  },

  async updateStatus(id, status, staffNote) {
    if (GasAPI.isConfigured()) await GasAPI.updateAppointmentStatus(id, status, staffNote);
    const list = this.getAll().map(a =>
      a.id === id ? { ...a, status, staffNote: staffNote ?? a.staffNote } : a
    );
    this.saveAll(list);
  },

  async cancel(id) {
    if (GasAPI.isConfigured()) await GasAPI.cancelAppointment(id);
    const list = this.getAll().map(a => a.id === id ? { ...a, status: 'cancelled' } : a);
    this.saveAll(list);
  },
};

/* =============================================
   AppointmentUI モジュール
   予約タブ（日次ビュー・ステータス管理）
   ============================================= */
const AppointmentUI = {
  init() {
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('apptDate').value = today;

    document.getElementById('apptPrevDay').addEventListener('click', () => {
      const d = new Date(document.getElementById('apptDate').value + 'T00:00:00');
      d.setDate(d.getDate() - 1);
      document.getElementById('apptDate').value = d.toISOString().split('T')[0];
      this.renderAppointments();
    });

    document.getElementById('apptNextDay').addEventListener('click', () => {
      const d = new Date(document.getElementById('apptDate').value + 'T00:00:00');
      d.setDate(d.getDate() + 1);
      document.getElementById('apptDate').value = d.toISOString().split('T')[0];
      this.renderAppointments();
    });

    document.getElementById('apptToday').addEventListener('click', () => {
      document.getElementById('apptDate').value = new Date().toISOString().split('T')[0];
      this.renderAppointments();
    });

    document.getElementById('apptDate').addEventListener('change', () => this.renderAppointments());
  },

  async renderAppointments() {
    const dateStr = document.getElementById('apptDate').value;
    const el      = document.getElementById('appointmentList');
    if (!el) return;

    // 日付ラベル
    const [y, m, d] = dateStr.split('-');
    const weekDays  = ['日','月','火','水','木','金','土'];
    const dow       = weekDays[new Date(dateStr + 'T00:00:00').getDay()];
    const label     = `${y}年${Number(m)}月${Number(d)}日（${dow}）`;

    el.innerHTML = `<div style="font-family:var(--font-serif);font-size:0.95rem;font-weight:700;color:var(--text);margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid var(--border);">${label}の予約</div>
      <div style="text-align:center;padding:24px;color:var(--text-sub);font-size:13px;">読み込み中...</div>`;

    // GASから当日の予約を取得
    if (GasAPI.isConfigured() && navigator.onLine) {
      await AppointmentData.loadFromGas(dateStr, dateStr);
    }

    const appts = AppointmentData.getByDate(dateStr)
      .sort((a, b) => (a.dateTime || '').localeCompare(b.dateTime || ''));

    const statusOrder = { pending: 0, confirmed: 1, completed: 2, noshow: 3, cancelled: 4 };
    appts.sort((a, b) => (statusOrder[a.status] ?? 9) - (statusOrder[b.status] ?? 9) ||
                         (a.dateTime || '').localeCompare(b.dateTime || ''));

    if (!appts.length) {
      el.innerHTML = `<div style="font-family:var(--font-serif);font-size:0.95rem;font-weight:700;color:var(--text);margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid var(--border);">${label}の予約</div>
        <p style="text-align:center;color:var(--text-light);font-size:13px;padding:32px 0;">予約はありません</p>`;
      return;
    }

    const statusBadge = { pending: '🟡 未確認', confirmed: '🟢 確定', completed: '✅ 完了', cancelled: '❌ キャンセル', noshow: '🔴 NC' };
    const statusColor = { pending: '#FEF3C7', confirmed: '#EAF3EC', completed: '#F0F0F0', cancelled: '#F5F5F5', noshow: '#F5EDED' };

    const cardsHtml = appts.map(a => {
      const time = (a.dateTime || '').substring(11, 16);
      const badge = statusBadge[a.status] || a.status;
      const bg    = statusColor[a.status] || '#F5F5F5';
      const isDimmed = a.status === 'cancelled' || a.status === 'completed' || a.status === 'noshow';

      const actions = [];
      if (a.status === 'pending') {
        actions.push(`<button class="appt-action-btn" data-id="${UI._esc(a.id)}" data-action="confirm"
          style="background:var(--income);color:white;border:none;border-radius:var(--radius-xs);padding:5px 12px;font-size:12px;cursor:pointer;font-family:var(--font-sans);font-weight:700;">確定する</button>`);
        actions.push(`<button class="appt-action-btn" data-id="${UI._esc(a.id)}" data-action="cancel"
          style="background:none;border:1px solid #FECACA;border-radius:var(--radius-xs);padding:5px 12px;font-size:12px;cursor:pointer;color:#EF4444;font-family:var(--font-sans);">キャンセル</button>`);
      } else if (a.status === 'confirmed') {
        actions.push(`<button class="appt-action-btn" data-id="${UI._esc(a.id)}" data-action="complete"
          style="background:var(--primary);color:var(--accent);border:none;border-radius:var(--radius-xs);padding:5px 12px;font-size:12px;cursor:pointer;font-family:var(--font-sans);font-weight:700;">完了</button>`);
        actions.push(`<button class="appt-action-btn" data-id="${UI._esc(a.id)}" data-action="noshow"
          style="background:none;border:1px solid #FECACA;border-radius:var(--radius-xs);padding:5px 12px;font-size:12px;cursor:pointer;color:#EF4444;font-family:var(--font-sans);">NC</button>`);
        actions.push(`<button class="appt-action-btn" data-id="${UI._esc(a.id)}" data-action="cancel"
          style="background:none;border:1px solid var(--border-normal);border-radius:var(--radius-xs);padding:5px 12px;font-size:12px;cursor:pointer;color:var(--text-sub);font-family:var(--font-sans);">取消</button>`);
      }

      return `
        <div style="background:${isDimmed ? '#F9F9F8' : 'white'};border:1px solid var(--border);border-radius:var(--radius-sm);
          padding:14px 16px;margin-bottom:10px;opacity:${isDimmed ? '0.65' : '1'};">
          <div style="display:flex;align-items:flex-start;gap:12px;flex-wrap:wrap;">
            <div style="font-family:var(--font-serif);font-size:1.3rem;font-weight:700;color:var(--text);min-width:52px;">${time || '--:--'}</div>
            <div style="flex:1;min-width:0;">
              <div style="font-weight:700;font-size:14px;">${UI._esc(a.customerName || '')}</div>
              <div style="font-size:12px;color:var(--text-sub);margin-top:2px;">${UI._esc(a.menuName || '')} ／ ¥${Number(a.price).toLocaleString()}</div>
              ${a.phone ? `<div style="font-size:12px;color:var(--text-sub);">${UI._esc(a.phone)}</div>` : ''}
              ${a.notes ? `<div style="font-size:12px;color:var(--text-sub);margin-top:4px;">備考: ${UI._esc(a.notes)}</div>` : ''}
            </div>
            <span style="background:${bg};padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;white-space:nowrap;">${badge}</span>
          </div>
          ${actions.length ? `<div style="display:flex;gap:8px;margin-top:10px;flex-wrap:wrap;">${actions.join('')}</div>` : ''}
        </div>`;
    }).join('');

    el.innerHTML = `
      <div style="font-family:var(--font-serif);font-size:0.95rem;font-weight:700;color:var(--text);margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid var(--border);">${label}の予約 (${appts.length}件)</div>
      ${cardsHtml}`;

    // ステータス変更ボタン
    el.querySelectorAll('.appt-action-btn').forEach(btn => {
      btn.addEventListener('click', () => this._handleAction(btn.dataset.id, btn.dataset.action));
    });
  },

  // 予約完了時: 売上トランザクションをローカル＋GASに自動作成
  async _autoRecordTransaction(appt) {
    const dateOnly = (appt.dateTime || '').substring(0, 10) || new Date().toISOString().split('T')[0];
    const record = {
      id:        `appt_${appt.id}_${Date.now()}`,
      date:      dateOnly,
      type:      'income',
      amount:    Number(appt.price || 0),
      category:  '売上',
      memo:      `${appt.menuName || 'メニュー'}（予約 #${appt.id.substring(0, 8)}）`,
      paymentMethod: Master.getPaymentMethods()[0] || '現金',
      tags:      [],
      appointmentId: appt.id,
      createdAt: new Date().toISOString(),
    };
    // GAS同期（設定済みの場合のみ、失敗してもローカル保存は続行）
    if (GasAPI.isConfigured()) {
      try { await GasAPI.addTransaction(record); } catch(e) { console.warn('[AutoRecord] GAS sync failed:', e.message); }
    }
    const list = Storage.getTransactions();
    list.unshift(record);
    Storage.saveTransactions(list);
  },

  // 予約完了時: 顧客の来店回数・最終来店日・累計売上を更新
  async _updateCustomerStats(appt) {
    if (!appt.customerId) return;
    const cust = CustomerData.getById(appt.customerId);
    if (!cust) return;
    const dateOnly = (appt.dateTime || '').substring(0, 10) || new Date().toISOString().split('T')[0];
    await CustomerData.update({
      ...cust,
      lastVisit:  dateOnly,
      visitCount: (cust.visitCount || 0) + 1,
      totalSpend: (cust.totalSpend || 0) + Number(appt.price || 0),
    });
  },

  async _handleAction(id, action) {
    const statusMap = { confirm: 'confirmed', complete: 'completed', noshow: 'noshow', cancel: 'cancelled' };
    const labelMap  = { confirm: '確定', complete: '完了にする', noshow: '無断欠席（NC）にする', cancel: 'キャンセルにする' };
    const status    = statusMap[action];
    if (!status) return;

    // 確認
    const appt = AppointmentData.getAll().find(a => a.id === id);
    if (!appt) return;
    document.getElementById('modalTitle').textContent   = `予約を${labelMap[action]}しますか？`;
    document.getElementById('modalMessage').textContent = `${appt.customerName} 様 / ${appt.menuName}`;
    document.getElementById('modal').style.display      = 'flex';

    const confirmBtn = document.getElementById('modalConfirm');
    const newBtn     = confirmBtn.cloneNode(true);
    newBtn.textContent = labelMap[action];
    confirmBtn.parentNode.replaceChild(newBtn, confirmBtn);
    newBtn.addEventListener('click', async () => {
      document.getElementById('modal').style.display = 'none';
      try {
        await AppointmentData.updateStatus(id, status);

        // 完了時: 売上自動記録 + 顧客情報更新
        if (action === 'complete') {
          await this._autoRecordTransaction(appt);
          await this._updateCustomerStats(appt);
        }

        UI._showToast(`✅ 予約を${labelMap[action]}にしました`, 'success');
        this.renderAppointments();
        // ダッシュボードが表示中なら再描画
        if (document.getElementById('tab-dashboard')?.classList.contains('active')) {
          UI.renderDashboard();
        }
      } catch(e) {
        UI._showToast('⚠️ 更新に失敗しました: ' + e.message, 'warn');
      }
    });
  },
};

/* =============================================
   CustomerData モジュール
   顧客データのローカルキャッシュ＋GAS同期
   ============================================= */
const CustomerData = {
  getAll() {
    try { return JSON.parse(localStorage.getItem(CUSTOMERS_KEY) || '[]'); } catch { return []; }
  },

  saveAll(list) {
    localStorage.setItem(CUSTOMERS_KEY, JSON.stringify(list));
  },

  getById(id) {
    return this.getAll().find(c => c.id === id) || null;
  },

  async add(data) {
    const customer = {
      id:           'cust_' + Date.now(),
      createdAt:    new Date().toISOString().split('T')[0],
      visitCount:   0,
      totalSpend:   0,
      lastVisit:    '',
      ...data,
    };
    if (GasAPI.isConfigured()) await GasAPI.addCustomer(customer);
    const list = this.getAll();
    list.push(customer);
    this.saveAll(list);
    return customer;
  },

  async update(data) {
    if (GasAPI.isConfigured()) await GasAPI.updateCustomer(data);
    const list = this.getAll().map(c => c.id === data.id ? { ...c, ...data } : c);
    this.saveAll(list);
  },

  async remove(id) {
    if (GasAPI.isConfigured()) await GasAPI.deleteCustomer(id);
    this.saveAll(this.getAll().filter(c => c.id !== id));
  },

  async loadFromGas() {
    try {
      const customers = await GasAPI.getCustomers();
      this.saveAll(customers);
      return true;
    } catch (e) {
      console.warn('[CustomerData] GAS load failed:', e.message);
      return false;
    }
  },
};

/* =============================================
   CustomerUI モジュール
   顧客タブの描画・検索・追加/編集
   ============================================= */
const CustomerUI = {
  _editingId: null,

  init() {
    // 検索
    document.getElementById('customerSearch').addEventListener('input', e => {
      this._renderList(e.target.value.trim());
    });

    // 新規登録ボタン
    document.getElementById('addCustomerBtn').addEventListener('click', () => {
      this._openModal(null);
    });

    // モーダルキャンセル
    document.getElementById('customerModalCancel').addEventListener('click', () => {
      this._closeModal();
    });
    document.getElementById('customerModal').addEventListener('click', e => {
      if (e.target === document.getElementById('customerModal')) this._closeModal();
    });

    // フォーム送信
    document.getElementById('customerForm').addEventListener('submit', e => {
      e.preventDefault();
      this._submitForm();
    });
  },

  renderCustomers() {
    this._renderList(document.getElementById('customerSearch')?.value.trim() || '');
  },

  _renderList(query = '') {
    const container = document.getElementById('customerList');
    if (!container) return;

    let customers = CustomerData.getAll();

    if (query) {
      const q = query.toLowerCase();
      customers = customers.filter(c =>
        (c.name     || '').toLowerCase().includes(q) ||
        (c.nameKana || '').toLowerCase().includes(q) ||
        (c.phone    || '').includes(q)
      );
    }

    // フリガナ→氏名 順にソート
    customers.sort((a, b) =>
      (a.nameKana || a.name || '').localeCompare(b.nameKana || b.name || '', 'ja')
    );

    if (!customers.length) {
      container.innerHTML = `<p style="color:var(--text-light);font-size:14px;text-align:center;padding:40px 0;">${query ? '該当する顧客が見つかりません' : '顧客が登録されていません'}</p>`;
      return;
    }

    container.innerHTML = customers.map(c => `
      <div class="customer-card" data-id="${this._esc(c.id)}"
        style="display:flex;align-items:center;gap:14px;padding:14px 16px;background:white;
               border:1px solid var(--border);border-radius:var(--radius-sm);margin-bottom:8px;
               cursor:pointer;transition:border-color 0.15s,box-shadow 0.15s;">
        <div style="width:40px;height:40px;border-radius:50%;background:var(--accent-dim);
                    display:flex;align-items:center;justify-content:center;flex-shrink:0;
                    font-family:var(--font-serif);font-size:1.1rem;color:var(--accent);font-weight:700;">
          ${this._esc((c.name || '？').charAt(0))}
        </div>
        <div style="flex:1;min-width:0;">
          <div style="font-size:0.9rem;font-weight:700;color:var(--text);">${this._esc(c.name || '')}</div>
          <div style="font-size:0.72rem;color:var(--text-sub);">${this._esc(c.nameKana || '')}</div>
          <div style="font-size:0.72rem;color:var(--text-sub);margin-top:2px;">${this._esc(c.phone || '')}</div>
        </div>
        <div style="text-align:right;font-size:0.72rem;color:var(--text-sub);flex-shrink:0;">
          ${c.lastVisit ? `最終来店<br><span style="color:var(--text);font-weight:600;">${this._esc(c.lastVisit)}</span>` : '<span style="color:var(--text-light);">未来店</span>'}
          <br><span style="margin-top:4px;display:inline-block;">来店 <strong>${c.visitCount || 0}</strong> 回</span>
        </div>
        <div style="flex-shrink:0;display:flex;flex-direction:column;gap:4px;">
          <button class="cust-edit-btn edit-btn" data-id="${this._esc(c.id)}"
            style="background:none;border:1px solid var(--border-normal);border-radius:var(--radius-xs);
                   padding:3px 10px;font-size:11px;cursor:pointer;color:var(--text);font-family:var(--font-sans);">編集</button>
          <button class="cust-del-btn delete-btn" data-id="${this._esc(c.id)}"
            style="background:none;border:1px solid #FECACA;border-radius:var(--radius-xs);
                   padding:3px 10px;font-size:11px;cursor:pointer;color:#EF4444;font-family:var(--font-sans);">削除</button>
        </div>
      </div>`).join('');

    container.querySelectorAll('.cust-edit-btn').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        this._openModal(btn.dataset.id);
      });
    });

    container.querySelectorAll('.cust-del-btn').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        this._confirmDelete(btn.dataset.id);
      });
    });
  },

  _openModal(id) {
    this._editingId = id;
    const isEdit = !!id;
    document.getElementById('customerModalTitle').textContent = isEdit ? '顧客情報を編集' : '顧客を登録する';
    document.getElementById('customerModalSave').textContent = isEdit ? '更新する' : '保存する';
    UI._showMsg('customerFormMsg', '', '');
    document.getElementById('customerFormMsg').style.display = 'none';

    if (isEdit) {
      const c = CustomerData.getById(id);
      if (!c) return;
      document.getElementById('cName').value         = c.name        || '';
      document.getElementById('cNameKana').value     = c.nameKana    || '';
      document.getElementById('cPhone').value        = c.phone       || '';
      document.getElementById('cEmail').value        = c.email       || '';
      document.getElementById('cBirthday').value     = c.birthday    || '';
      document.getElementById('cTags').value         = c.tags        || '';
      document.getElementById('cAllergyNotes').value = c.allergyNotes || '';
      document.getElementById('cMemo').value         = c.memo        || '';
    } else {
      document.getElementById('customerForm').reset();
    }

    document.getElementById('customerModal').style.display = 'flex';
    document.getElementById('cName').focus();
  },

  _closeModal() {
    document.getElementById('customerModal').style.display = 'none';
    this._editingId = null;
  },

  async _submitForm() {
    const name = document.getElementById('cName').value.trim();
    if (!name) {
      UI._showMsg('customerFormMsg', '⚠️ 氏名を入力してください', 'error');
      return;
    }

    const data = {
      name:         name,
      nameKana:     document.getElementById('cNameKana').value.trim(),
      phone:        document.getElementById('cPhone').value.trim(),
      email:        document.getElementById('cEmail').value.trim(),
      birthday:     document.getElementById('cBirthday').value.trim(),
      tags:         document.getElementById('cTags').value.trim(),
      allergyNotes: document.getElementById('cAllergyNotes').value.trim(),
      memo:         document.getElementById('cMemo').value.trim(),
    };

    const saveBtn = document.getElementById('customerModalSave');
    saveBtn.disabled = true;
    saveBtn.textContent = '保存中...';

    try {
      if (this._editingId) {
        await CustomerData.update({ id: this._editingId, ...data });
        UI._showToast('✅ 顧客情報を更新しました', 'success');
      } else {
        await CustomerData.add(data);
        UI._showToast('✅ 顧客を登録しました', 'success');
      }
      this._closeModal();
      this.renderCustomers();
    } catch (e) {
      UI._showMsg('customerFormMsg', '⚠️ 保存に失敗しました: ' + e.message, 'error');
    } finally {
      saveBtn.disabled = false;
      saveBtn.textContent = this._editingId ? '更新する' : '保存する';
    }
  },

  _confirmDelete(id) {
    const c = CustomerData.getById(id);
    if (!c) return;
    document.getElementById('modalTitle').textContent   = '顧客を削除しますか？';
    document.getElementById('modalMessage').textContent = `「${c.name}」を削除します。この操作は元に戻せません。`;
    document.getElementById('modal').style.display      = 'flex';

    const confirmBtn = document.getElementById('modalConfirm');
    const newBtn = confirmBtn.cloneNode(true);
    confirmBtn.parentNode.replaceChild(newBtn, confirmBtn);
    newBtn.addEventListener('click', async () => {
      document.getElementById('modal').style.display = 'none';
      await CustomerData.remove(id);
      this.renderCustomers();
      UI._showToast('🗑️ 顧客を削除しました', 'info');
    });
  },

  _esc(str) {
    const el = document.createElement('div');
    el.textContent = String(str || '');
    return el.innerHTML;
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
  UI.initLedger();
  UI.initOfflineDetection();
  CustomerUI.init();
  AppointmentUI.init();

  const settings = Storage.getSettings();
  if (settings.businessName && settings.businessName !== 'マイサロン') {
    document.getElementById('headerSubtitle').textContent =
      settings.businessName + '  ┊  青色申告対応';
  }

  // キャッシュを使って即時表示
  UI.renderDashboard();

  if (!GasAPI.isConfigured()) {
    UI._showToast('⚠️ GAS URL が未設定です。app.js の GAS_URL 定数を設定してください。', 'warn');
    return;
  }

  if (navigator.onLine) {
    // GAS からデータ・マスタを取得してキャッシュを更新
    const [loaded] = await Promise.all([
      Storage.loadFromGas(),
      Storage.loadMasterFromGas(),
      CustomerData.loadFromGas(),
      AppointmentData.loadFromGas(),
    ]);
    if (loaded) {
      UI.renderDashboard();
      // フォームの支払方法・科目も更新
      UI._rebuildPaymentMethodOptions();
      UI._rebuildCategoryOptions();
      // 収支一覧が表示中なら再描画（GAS同期後に最新データを反映）
      if (document.getElementById('tab-list')?.classList.contains('active')) {
        UI.renderList();
      }
    }
  }
}

document.addEventListener('DOMContentLoaded', init);
