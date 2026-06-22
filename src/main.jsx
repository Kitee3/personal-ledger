import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import "../styles.css";
import "./react.css";

const storageKey = "personal-ledger-v1";
const firebaseConfigKey = "personal-ledger-firebase-config";
const autoCloudSyncKey = "personal-ledger-auto-cloud-sync";

const defaultState = {
  accounts: [
    { id: "wechat", name: "微信", type: "wallet", initialBalance: 0 },
    { id: "alipay", name: "支付宝", type: "wallet", initialBalance: 0 },
    { id: "bank", name: "银行卡", type: "debit_card", initialBalance: 0 },
    { id: "cash", name: "现金", type: "cash", initialBalance: 0 },
    { id: "credit", name: "信用卡", type: "credit_card", initialBalance: 0 },
  ],
  expenseCategories: ["餐饮", "交通", "购物", "住房", "通讯", "医疗", "娱乐", "学习", "人情", "差旅", "其他"],
  incomeCategories: ["工资", "奖金", "补贴", "报销到账", "副业", "其他"],
  transactions: [],
  trips: [],
  attachments: [],
  budgets: {
    total: 0,
    categories: {},
  },
  lastAccount: "wechat",
};

const views = [
  ["dashboard", "首页"],
  ["quick", "记一笔"],
  ["transactions", "流水"],
  ["trips", "差旅"],
  ["budget", "预算"],
  ["import", "导入"],
  ["settings", "设置"],
];

function today() {
  return new Date().toISOString().slice(0, 10);
}

function loadState() {
  const raw = localStorage.getItem(storageKey);
  if (!raw) return structuredClone(defaultState);
  try {
    return { ...structuredClone(defaultState), ...JSON.parse(raw) };
  } catch {
    return structuredClone(defaultState);
  }
}

function saveLocalState(nextState) {
  localStorage.setItem(storageKey, JSON.stringify(nextState));
}

function money(value) {
  return `¥${Number(value || 0).toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}

function getMonthKey(date = new Date()) {
  return date.toISOString().slice(0, 7);
}

function inCurrentMonth(transaction) {
  return String(transaction.occurredAt || "").startsWith(getMonthKey());
}

function byDateDesc(a, b) {
  return `${b.occurredAt} ${b.updatedAt || ""}`.localeCompare(`${a.occurredAt} ${a.updatedAt || ""}`);
}

function toCsv(rows) {
  return rows
    .map((row) => row.map((cell) => `"${String(cell ?? "").replaceAll('"', '""')}"`).join(","))
    .join("\n");
}

function download(filename, content, type = "text/plain;charset=utf-8") {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

async function buildAttachments(files) {
  const selected = [...files].slice(0, 3);
  const validFiles = selected.filter((file) => file.size <= 2 * 1024 * 1024);
  return Promise.all(
    validFiles.map(async (file) => ({
      id: crypto.randomUUID(),
      name: file.name,
      type: file.type || "application/octet-stream",
      size: file.size,
      dataUrl: await readFileAsDataUrl(file),
      createdAt: new Date().toISOString(),
    }))
  );
}

function parseCsv(text) {
  const lines = text.replace(/^\ufeff/, "").split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];
  const parseLine = (line) => {
    const cells = [];
    let current = "";
    let quote = false;
    for (let index = 0; index < line.length; index += 1) {
      const char = line[index];
      const next = line[index + 1];
      if (char === '"' && quote && next === '"') {
        current += '"';
        index += 1;
      } else if (char === '"') {
        quote = !quote;
      } else if (char === "," && !quote) {
        cells.push(current);
        current = "";
      } else {
        current += char;
      }
    }
    cells.push(current);
    return cells.map((cell) => cell.trim());
  };
  const headerIndex = Math.max(
    0,
    lines.findIndex((line) => /日期|时间|金额|交易/.test(parseLine(line).join("")))
  );
  const headers = parseLine(lines[headerIndex]);
  return lines.slice(headerIndex + 1).map((line) => {
    const cells = parseLine(line);
    return Object.fromEntries(headers.map((header, index) => [header, cells[index] || ""]));
  });
}

async function parseBillFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const XLSX = await import("xlsx");
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return [];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
    const headerIndex = Math.max(
      0,
      rows.findIndex((row) => /日期|时间|金额|交易/.test(row.join("")))
    );
    const headers = rows[headerIndex].map((cell) => String(cell).trim());
    return rows.slice(headerIndex + 1).map((row) =>
      Object.fromEntries(headers.map((header, index) => [header, String(row[index] ?? "").trim()]))
    );
  }
  return parseCsv(await file.text());
}

function getFirstValue(row, keys) {
  for (const key of keys) {
    if (row[key] !== undefined && String(row[key]).trim() !== "") return String(row[key]).trim();
  }
  return "";
}

function parseAmount(value) {
  const cleaned = String(value || "").replace(/[￥¥,\s]/g, "").replace(/[()]/g, "").trim();
  const match = cleaned.match(/-?\d+(\.\d+)?/);
  return match ? Math.abs(Number(match[0])) : 0;
}

function guessCategory({ merchant, note, category }, type) {
  const text = `${merchant || ""} ${note || ""}`.toLowerCase();
  if (type === "income") return category || "其他";
  if (/酒店|宾馆|航旅|携程|飞猪/.test(text)) return "差旅";
  if (/滴滴|打车|地铁|公交|火车|机票|高铁|taxi/.test(text)) return "交通";
  if (/美团|饿了么|餐|咖啡|饭|奶茶/.test(text)) return "餐饮";
  if (/京东|淘宝|天猫|拼多多|超市/.test(text)) return "购物";
  return category || "其他";
}

function guessAccount({ merchant, note, fallback }) {
  const text = `${merchant || ""} ${note || ""}`.toLowerCase();
  if (/支付宝|alipay|花呗|余额宝/.test(text)) return "alipay";
  if (/微信|wechat|零钱/.test(text)) return "wechat";
  if (/信用卡|visa|mastercard|白条|分期/.test(text)) return "credit";
  return fallback || "wechat";
}

function applyQuickGuess(form, type, fallbackAccount) {
  const merchant = form.elements.merchant?.value || "";
  const note = form.elements.note?.value || "";
  const category = guessCategory({ merchant, note, category: "" }, type);
  const account = guessAccount({ merchant, note, fallback: fallbackAccount });
  if (form.elements.category && [...form.elements.category.options].some((option) => option.value === category)) {
    form.elements.category.value = category;
  }
  if (form.elements.account && [...form.elements.account.options].some((option) => option.value === account)) {
    form.elements.account.value = account;
  }
}

function shouldAttachToTrip({ merchant, note, category, type, mode }) {
  if (type !== "expense") return false;
  if (mode === "all-expense") return true;
  const text = `${merchant || ""} ${note || ""} ${category || ""}`.toLowerCase();
  return /酒店|宾馆|航旅|携程|飞猪|机票|高铁|火车|动车|打车|滴滴|出租车|机场|车站/.test(text);
}

function App() {
  const requestedView = new URLSearchParams(window.location.search).get("view");
  const initialView = requestedView === "sync" ? "settings" : requestedView;
  const [activeView, setActiveView] = useState(views.some(([key]) => key === initialView) ? initialView : "dashboard");
  const [state, setState] = useState(loadState);
  const [entryType, setEntryType] = useState("expense");
  const [filters, setFilters] = useState({ search: "", type: "all", trip: "all" });
  const [pendingImportRows, setPendingImportRows] = useState([]);
  const [importLog, setImportLog] = useState("等待导入。");
  const [importTripId, setImportTripId] = useState("");
  const [importTripMode, setImportTripMode] = useState("strong");
  const [firebaseConfigText, setFirebaseConfigText] = useState(() => {
    const raw = localStorage.getItem(firebaseConfigKey);
    return raw ? JSON.stringify(JSON.parse(raw), null, 2) : "";
  });
  const [syncStatus, setSyncStatus] = useState({ title: "未连接", text: "当前只保存在本机浏览器。", tone: "offline" });
  const [syncLog, setSyncLog] = useState("等待连接。");
  const [syncAuth, setSyncAuth] = useState({ email: "", password: "" });
  const [autoCloudSync, setAutoCloudSync] = useState(localStorage.getItem(autoCloudSyncKey) !== "false");
  const [firebaseClient, setFirebaseClient] = useState(null);
  const [cloudUser, setCloudUser] = useState(null);

  const openTrips = state.trips.filter((trip) => trip.status !== "reimbursed");

  function getAccountName(id) {
    return state.accounts.find((account) => account.id === id)?.name || id || "未设置";
  }

  function getTripName(id) {
    return state.trips.find((trip) => trip.id === id)?.name || (id ? "未关联项目" : "-");
  }

  function commitState(nextState, { sync = true } = {}) {
    setState(nextState);
    saveLocalState(nextState);
    if (sync && autoCloudSync && cloudUser && firebaseClient) {
      pushCloudState(nextState, { silent: true });
    }
  }

  const monthTransactions = useMemo(
    () => state.transactions.filter((transaction) => !transaction.deletedAt && inCurrentMonth(transaction)),
    [state.transactions]
  );
  const budgetSettings = state.budgets || defaultState.budgets;

  const summary = useMemo(
    () =>
      monthTransactions.reduce(
        (acc, transaction) => {
          if (transaction.type === "income") acc.income += Number(transaction.amount);
          if (transaction.type === "expense") acc.expense += Number(transaction.amount);
          if (transaction.isTrip && transaction.type === "expense" && ["pending", "submitted"].includes(transaction.reimbursementStatus)) {
            acc.pending += Number(transaction.amount);
          }
          return acc;
        },
        { income: 0, expense: 0, pending: 0 }
      ),
    [monthTransactions]
  );

  const expenseByCategory = useMemo(() => Object.fromEntries(groupExpenseByCategory(monthTransactions)), [monthTransactions]);
  const totalBudget = Number(budgetSettings.total || 0);
  const totalBudgetRemaining = totalBudget ? totalBudget - summary.expense : 0;
  const totalBudgetRate = totalBudget ? (summary.expense / totalBudget) * 100 : 0;

  function groupExpenseByCategory(transactions) {
    const groups = new Map();
    transactions
      .filter((transaction) => transaction.type === "expense")
      .forEach((transaction) => groups.set(transaction.category, (groups.get(transaction.category) || 0) + Number(transaction.amount)));
    return [...groups.entries()].sort((a, b) => b[1] - a[1]);
  }

  function exportJson() {
    download(`简账助手备份-${today()}.json`, JSON.stringify(state, null, 2), "application/json;charset=utf-8");
  }

  async function addTransaction(event) {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const amount = Number(form.get("amount"));
    if (!amount || amount <= 0) return;
    const isTrip = form.get("isTrip") === "on";
    const account = form.get("account");
    const attachments = await buildAttachments(event.currentTarget.elements.attachments.files || []);
    const transaction = {
      id: crypto.randomUUID(),
      amount,
      type: entryType,
      category: form.get("category"),
      account,
      occurredAt: form.get("occurredAt") || today(),
      merchant: String(form.get("merchant") || "").trim(),
      note: String(form.get("note") || "").trim(),
      isTrip,
      tripId: isTrip ? form.get("tripId") || "" : "",
      reimbursementStatus: isTrip ? form.get("reimbursementStatus") || "pending" : "none",
      attachmentIds: attachments.map((attachment) => attachment.id),
      source: "manual",
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    commitState({
      ...state,
      attachments: [...(state.attachments || []), ...attachments],
      transactions: [...state.transactions, transaction],
      lastAccount: account,
    });
    event.currentTarget.reset();
    setEntryType("expense");
    setActiveView("dashboard");
  }

  function deleteTransaction(id) {
    commitState({
      ...state,
      transactions: state.transactions.map((transaction) =>
        transaction.id === id ? { ...transaction, deletedAt: new Date().toISOString(), updatedAt: new Date().toISOString() } : transaction
      ),
    });
  }

  function updateReimbursementStatus(id, reimbursementStatus) {
    const now = new Date().toISOString();
    commitState({
      ...state,
      transactions: state.transactions.map((transaction) =>
        transaction.id === id
          ? {
              ...transaction,
              isTrip: reimbursementStatus === "none" ? transaction.isTrip : true,
              reimbursementStatus,
              updatedAt: now,
            }
          : transaction
      ),
    });
  }

  function updateTripReimbursementStatus(tripId, reimbursementStatus) {
    const now = new Date().toISOString();
    commitState({
      ...state,
      trips: state.trips.map((trip) =>
        trip.id === tripId
          ? {
              ...trip,
              status: reimbursementStatus === "reimbursed" ? "reimbursed" : "active",
              updatedAt: now,
            }
          : trip
      ),
      transactions: state.transactions.map((transaction) =>
        !transaction.deletedAt && transaction.tripId === tripId && transaction.type === "expense"
          ? {
              ...transaction,
              isTrip: true,
              reimbursementStatus,
              updatedAt: now,
            }
          : transaction
      ),
    });
  }

  function addTrip(event) {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const name = String(form.get("name") || "").trim();
    if (!name) return;
    commitState({
      ...state,
      trips: [
        ...state.trips,
        {
          id: crypto.randomUUID(),
          name,
          destination: String(form.get("destination") || "").trim(),
          startDate: form.get("startDate") || "",
          endDate: form.get("endDate") || "",
          purpose: String(form.get("purpose") || "").trim(),
          status: "active",
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        },
      ],
    });
    event.currentTarget.reset();
  }

  function exportTripCsv() {
    const rows = [["差旅项目", "日期", "分类", "商户", "备注", "账户", "报销状态", "金额"]];
    state.transactions
      .filter((transaction) => !transaction.deletedAt && transaction.isTrip)
      .sort(byDateDesc)
      .forEach((transaction) =>
        rows.push([
          getTripName(transaction.tripId),
          transaction.occurredAt,
          transaction.category,
          transaction.merchant,
          transaction.note,
          getAccountName(transaction.account),
          transaction.reimbursementStatus,
          transaction.amount,
        ])
      );
    download(`差旅报销明细-${today()}.csv`, `\ufeff${toCsv(rows)}`, "text/csv;charset=utf-8");
  }

  function saveBudget(event) {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const nextCategories = {};
    state.expenseCategories.forEach((category) => {
      const value = Number(form.get(`budget-${category}`) || 0);
      if (value > 0) nextCategories[category] = value;
    });
    commitState({
      ...state,
      budgets: {
        total: Number(form.get("totalBudget") || 0),
        categories: nextCategories,
        updatedAt: new Date().toISOString(),
      },
    });
  }

  function normalizeImportRow(row) {
    const amountRaw = getFirstValue(row, ["金额", "金额(元)", "交易金额", "支出金额", "收入金额"]);
    const amount = parseAmount(amountRaw);
    const marker = getFirstValue(row, ["类型", "收/支", "收支", "交易类型", "资金流向"]);
    const type = /收入|收款|入账|\+|转入/.test(marker) ? "income" : "expense";
    const occurredAt = getFirstValue(row, ["日期", "交易时间", "时间", "付款时间", "创建时间"]).slice(0, 10) || today();
    const merchant = getFirstValue(row, ["商户", "交易对方", "对方", "收款方", "付款方", "商品", "商品说明"]);
    const note = getFirstValue(row, ["备注", "说明", "交易分类", "交易类型", "商品", "商品说明"]);
    const category = guessCategory({ merchant, note, category: getFirstValue(row, ["分类"]) }, type);
    const account = /支付宝|花呗/.test(JSON.stringify(row))
      ? "alipay"
      : /微信|零钱/.test(JSON.stringify(row))
        ? "wechat"
        : state.lastAccount || "wechat";
    const tripMatched = importTripId && shouldAttachToTrip({ merchant, note, category, type, mode: importTripMode });
    const importKey = [occurredAt, type, amount.toFixed(2), merchant, note].join("|").toLowerCase();
    return {
      amount,
      type,
      category,
      account,
      occurredAt,
      merchant,
      note,
      importKey,
      source: "import",
      isTrip: Boolean(tripMatched) || category === "差旅",
      tripId: tripMatched ? importTripId : "",
      reimbursementStatus: tripMatched || category === "差旅" ? "pending" : "none",
    };
  }

  function isDuplicateImport(row) {
    return state.transactions.some((transaction) => {
      if (transaction.deletedAt) return false;
      const existingKey =
        transaction.importKey ||
        [transaction.occurredAt, transaction.type, Number(transaction.amount).toFixed(2), transaction.merchant, transaction.note]
          .join("|")
          .toLowerCase();
      return existingKey === row.importKey;
    });
  }

  async function previewCsv(file) {
    if (!file) {
      setImportLog("请先选择 CSV 文件。");
      return;
    }
    const rows = (await parseBillFile(file)).map(normalizeImportRow);
    const nextRows = rows.map((row) => ({
      ...row,
      status: row.amount > 0 ? (isDuplicateImport(row) ? "duplicate" : "ready") : "invalid",
    }));
    setPendingImportRows(nextRows);
    setImportLog(`预览完成：${file.name}。确认无误后点击“确认导入”。`);
  }

  function confirmImportRows() {
    const rows = pendingImportRows.filter((row) => row.status === "ready");
    const now = new Date().toISOString();
    commitState({
      ...state,
      transactions: [
        ...state.transactions,
        ...rows.map((row) => ({
          id: crypto.randomUUID(),
          amount: row.amount,
          type: row.type,
          category: row.category,
          account: row.account,
          occurredAt: row.occurredAt,
          merchant: row.merchant,
          note: row.note,
          isTrip: row.isTrip,
          tripId: row.tripId || "",
          reimbursementStatus: row.reimbursementStatus,
          attachmentIds: [],
          importKey: row.importKey,
          source: row.source,
          createdAt: now,
          updatedAt: now,
        })),
      ],
    });
    setPendingImportRows(pendingImportRows.map((row) => (row.status === "ready" ? { ...row, status: "duplicate" } : row)));
    setImportLog(`导入完成：${rows.length} 条流水。重复记录已自动跳过。`);
  }

  function downloadTemplate() {
    const rows = [
      ["日期", "金额", "类型", "分类", "商户", "备注", "账户"],
      [today(), "25.5", "支出", "餐饮", "示例商户", "午餐", "微信"],
    ];
    download("账单导入模板.csv", `\ufeff${toCsv(rows)}`, "text/csv;charset=utf-8");
  }

  function getStoredFirebaseConfig() {
    const raw = localStorage.getItem(firebaseConfigKey);
    return raw ? JSON.parse(raw) : null;
  }

  async function initFirebaseClient() {
    if (firebaseClient) return firebaseClient;
    const config = getStoredFirebaseConfig();
    if (!config?.apiKey || !config?.projectId) throw new Error("请先保存有效的 Firebase 配置 JSON。");
    const [{ getApps, initializeApp }, authModule, firestoreModule] = await Promise.all([
      import("https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js"),
      import("https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js"),
      import("https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js"),
    ]);
    const app = getApps()[0] || initializeApp(config);
    const auth = authModule.getAuth(app);
    const db = firestoreModule.getFirestore(app);
    authModule.onAuthStateChanged(auth, (user) => {
      setCloudUser(user);
      setSyncStatus(
        user
          ? { title: "已连接", text: `当前账号：${user.email || user.uid}`, tone: "online" }
          : { title: "未登录", text: "Firebase 已配置，但还没有登录账号。", tone: "offline" }
      );
    });
    const client = {
      auth,
      db,
      createUserWithEmailAndPassword: authModule.createUserWithEmailAndPassword,
      signInWithEmailAndPassword: authModule.signInWithEmailAndPassword,
      signOut: authModule.signOut,
      doc: firestoreModule.doc,
      getDoc: firestoreModule.getDoc,
      serverTimestamp: firestoreModule.serverTimestamp,
      setDoc: firestoreModule.setDoc,
    };
    setFirebaseClient(client);
    return client;
  }

  async function saveFirebaseConfig() {
    try {
      const config = JSON.parse(firebaseConfigText.trim());
      if (!config.apiKey || !config.projectId || !config.authDomain) throw new Error("配置至少需要 apiKey、authDomain 和 projectId。");
      localStorage.setItem(firebaseConfigKey, JSON.stringify(config));
      setFirebaseClient(null);
      setSyncLog("Firebase 配置已保存。下一步可以注册或登录。");
      await initFirebaseClient();
    } catch (error) {
      setSyncStatus({ title: "配置错误", text: error.message, tone: "error" });
      setSyncLog(`配置保存失败：${error.message}`);
    }
  }

  async function loginCloud(mode) {
    try {
      const client = await initFirebaseClient();
      if (mode === "register") {
        await client.createUserWithEmailAndPassword(client.auth, syncAuth.email.trim(), syncAuth.password);
      } else {
        await client.signInWithEmailAndPassword(client.auth, syncAuth.email.trim(), syncAuth.password);
      }
    } catch (error) {
      setSyncStatus({ title: mode === "register" ? "注册失败" : "登录失败", text: error.message, tone: "error" });
      setSyncLog(error.message);
    }
  }

  async function pushCloudState(nextState = state, options = {}) {
    try {
      const client = await initFirebaseClient();
      const user = client.auth.currentUser;
      if (!user) throw new Error("请先登录 Firebase 账号。");
      await client.setDoc(client.doc(client.db, "users", user.uid, "ledger", "main"), {
        state: nextState,
        schemaVersion: 1,
        clientUpdatedAt: new Date().toISOString(),
        serverUpdatedAt: client.serverTimestamp(),
      });
      if (!options.silent) setSyncLog(`上传完成：${new Date().toLocaleString("zh-CN")}`);
    } catch (error) {
      if (!options.silent) {
        setSyncStatus({ title: "上传失败", text: error.message, tone: "error" });
        setSyncLog(error.message);
      }
    }
  }

  async function pullCloudState() {
    try {
      const client = await initFirebaseClient();
      const user = client.auth.currentUser;
      if (!user) throw new Error("请先登录 Firebase 账号。");
      const snapshot = await client.getDoc(client.doc(client.db, "users", user.uid, "ledger", "main"));
      if (!snapshot.exists()) {
        setSyncLog("云端还没有账本。可以先上传本地数据。");
        return;
      }
      const nextState = { ...structuredClone(defaultState), ...snapshot.data().state };
      commitState(nextState, { sync: false });
      setSyncLog(`拉取完成：${snapshot.data().clientUpdatedAt || "未知时间"}`);
    } catch (error) {
      setSyncStatus({ title: "拉取失败", text: error.message, tone: "error" });
      setSyncLog(error.message);
    }
  }

  useEffect(() => {
    const config = localStorage.getItem(firebaseConfigKey);
    if (config) {
      initFirebaseClient().catch((error) => setSyncLog(`初始化失败：${error.message}`));
    }
  }, []);

  const categories = entryType === "expense" ? state.expenseCategories : state.incomeCategories;
  const recent = state.transactions.filter((transaction) => !transaction.deletedAt).sort(byDateDesc).slice(0, 10);
  const filteredTransactions = state.transactions
    .filter((transaction) => !transaction.deletedAt)
    .filter((transaction) => (filters.type === "all" ? true : transaction.type === filters.type))
    .filter((transaction) => (filters.trip === "all" ? true : filters.trip === "trip" ? transaction.isTrip : !transaction.isTrip))
    .filter((transaction) => {
      const search = filters.search.trim().toLowerCase();
      if (!search) return true;
      return [transaction.category, transaction.merchant, transaction.note, getAccountName(transaction.account)].join(" ").toLowerCase().includes(search);
    })
    .sort(byDateDesc);

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark">账</div>
          <div>
            <h1>简账助手</h1>
            <p>个人财务与差旅报销</p>
          </div>
        </div>
        <nav className="nav" aria-label="主导航">
          {views.map(([key, label]) => (
            <button key={key} className={`nav-item ${activeView === key ? "active" : ""}`} type="button" onClick={() => setActiveView(key)}>
              {label}
            </button>
          ))}
        </nav>
        <div className="sync-card">
          <span className="sync-dot"></span>
          <div>
            <strong>{cloudUser ? "云同步已连接" : "本地原型"}</strong>
            <p>{cloudUser ? cloudUser.email : "数据保存在当前浏览器"}</p>
          </div>
        </div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div>
            <p className="eyebrow">本月概览</p>
            <h2>{views.find(([key]) => key === activeView)?.[1]}</h2>
          </div>
          <div className="topbar-actions">
            <button className="ghost-btn" type="button" onClick={exportJson}>
              导出 JSON
            </button>
            <button className="primary-btn" type="button" onClick={() => setActiveView("quick")}>
              记一笔
            </button>
          </div>
        </header>

        {activeView === "dashboard" && (
          <>
            <div className="metric-grid">
              <Metric label="本月收入" value={money(summary.income)} />
              <Metric label="本月支出" value={money(summary.expense)} />
              <Metric label="本月结余" value={money(summary.income - summary.expense)} />
              <Metric label="待报销" value={money(summary.pending)} accent />
            </div>
            <BudgetSnapshot
              totalBudget={totalBudget}
              spent={summary.expense}
              remaining={totalBudgetRemaining}
              rate={totalBudgetRate}
              onOpen={() => setActiveView("budget")}
            />
            <div className="two-column">
              <section className="panel">
                <div className="panel-head">
                  <h3>最近流水</h3>
                  <button className="text-btn" type="button" onClick={() => setActiveView("transactions")}>
                    查看全部
                  </button>
                </div>
                <div className="list">{recent.length ? recent.map((item) => <TransactionCard key={item.id} item={item} attachments={state.attachments || []} getAccountName={getAccountName} getTripName={getTripName} />) : <div className="empty">还没有流水，先记一笔。</div>}</div>
              </section>
              <section className="panel">
                <h3>支出分类 Top 5</h3>
                <Bars rows={groupExpenseByCategory(monthTransactions).slice(0, 5)} />
              </section>
            </div>
          </>
        )}

        {activeView === "quick" && (
          <form className="entry-form" onSubmit={addTransaction}>
            <div className="amount-row">
              <label htmlFor="amount">金额</label>
              <input id="amount" name="amount" inputMode="decimal" placeholder="0.00" required />
            </div>
            <div className="segmented">
              <button className={`segment ${entryType === "expense" ? "active" : ""}`} type="button" onClick={() => setEntryType("expense")}>
                支出
              </button>
              <button className={`segment ${entryType === "income" ? "active" : ""}`} type="button" onClick={() => setEntryType("income")}>
                收入
              </button>
            </div>
            <div className="form-grid">
              <label>
                分类
                <select name="category">{categories.map((category) => <option key={category}>{category}</option>)}</select>
              </label>
              <label>
                账户
                <select name="account" defaultValue={state.lastAccount}>{state.accounts.map((account) => <option key={account.id} value={account.id}>{account.name}</option>)}</select>
              </label>
              <label>
                日期
                <input name="occurredAt" type="date" defaultValue={today()} />
              </label>
              <label>
                商户
                <input
                  name="merchant"
                  placeholder="例如：滴滴、美团、酒店"
                  onBlur={(event) => applyQuickGuess(event.currentTarget.form, entryType, state.lastAccount)}
                />
              </label>
            </div>
            <label>
              备注
              <textarea
                name="note"
                rows="3"
                placeholder="可选"
                onBlur={(event) => applyQuickGuess(event.currentTarget.form, entryType, state.lastAccount)}
              ></textarea>
            </label>
            <button
              className="text-btn quick-guess-btn"
              type="button"
              onClick={(event) => applyQuickGuess(event.currentTarget.form, entryType, state.lastAccount)}
            >
              智能识别分类和账户
            </button>
            <label className="attachment-input">
              发票/凭证
              <input name="attachments" type="file" accept="image/*,.pdf" multiple />
              <span className="transaction-meta">最多 3 个文件，单个 2MB 以内。当前版本随账本本地保存并参与云同步。</span>
            </label>
            <div className="trip-box">
              <label className="checkbox-line">
                <input name="isTrip" type="checkbox" />
                <span>这是一笔差旅相关支出</span>
              </label>
              <div className="form-grid trip-fields active">
                <label>
                  差旅项目
                  <select name="tripId">
                    <option value="">选择差旅项目</option>
                    {openTrips.map((trip) => <option key={trip.id} value={trip.id}>{trip.name}</option>)}
                  </select>
                </label>
                <label>
                  报销状态
                  <select name="reimbursementStatus" defaultValue="pending">
                    <option value="pending">待整理</option>
                    <option value="submitted">已提交</option>
                    <option value="reimbursed">已报销</option>
                    <option value="none">不报销</option>
                  </select>
                </label>
              </div>
            </div>
            <div className="form-actions">
              <button className="primary-btn" type="submit">保存</button>
              <button className="ghost-btn" type="reset">清空</button>
            </div>
          </form>
        )}

        {activeView === "transactions" && (
          <>
            <div className="toolbar">
              <input placeholder="搜索商户、备注、分类" value={filters.search} onChange={(event) => setFilters({ ...filters, search: event.target.value })} />
              <select value={filters.type} onChange={(event) => setFilters({ ...filters, type: event.target.value })}>
                <option value="all">全部类型</option>
                <option value="expense">支出</option>
                <option value="income">收入</option>
              </select>
              <select value={filters.trip} onChange={(event) => setFilters({ ...filters, trip: event.target.value })}>
                <option value="all">全部流水</option>
                <option value="trip">仅差旅</option>
                <option value="normal">非差旅</option>
              </select>
            </div>
            <TransactionTable
              rows={filteredTransactions}
              attachments={state.attachments || []}
              getAccountName={getAccountName}
              onDelete={deleteTransaction}
              onStatusChange={updateReimbursementStatus}
            />
          </>
        )}

        {activeView === "trips" && (
          <div className="split-layout">
            <form className="panel compact-form" onSubmit={addTrip}>
              <h3>新建差旅项目</h3>
              <label>项目名称<input name="name" required placeholder="例如：2026-07 北京出差" /></label>
              <label>地点<input name="destination" placeholder="北京" /></label>
              <div className="form-grid">
                <label>开始日期<input name="startDate" type="date" /></label>
                <label>结束日期<input name="endDate" type="date" /></label>
              </div>
              <label>事由<textarea name="purpose" rows="3"></textarea></label>
              <button className="primary-btn" type="submit">创建项目</button>
            </form>
            <section className="panel">
              <div className="panel-head">
                <h3>差旅项目</h3>
                <button className="ghost-btn" type="button" onClick={exportTripCsv}>导出报销 CSV</button>
              </div>
              <TripList trips={state.trips} transactions={state.transactions} onBatchStatusChange={updateTripReimbursementStatus} />
            </section>
          </div>
        )}

        {activeView === "budget" && (
          <BudgetPage
            categories={state.expenseCategories}
            budgets={budgetSettings}
            expenseByCategory={expenseByCategory}
            totalSpent={summary.expense}
            onSave={saveBudget}
          />
        )}

        {activeView === "import" && (
          <section className="panel">
            <h3>账单导入</h3>
            <p className="muted">支持微信/支付宝导出的 Excel，也支持 CSV。先预览确认，再批量写入账本。</p>
            <div className="form-grid import-options">
              <label>关联差旅项目<select value={importTripId} onChange={(event) => setImportTripId(event.target.value)}>
                <option value="">不自动关联差旅</option>
                {openTrips.map((trip) => <option key={trip.id} value={trip.id}>{trip.name}</option>)}
              </select></label>
              <label>差旅识别<select value={importTripMode} onChange={(event) => setImportTripMode(event.target.value)}>
                <option value="strong">仅强相关支出</option>
                <option value="all-expense">所有支出都关联</option>
              </select></label>
            </div>
            <input id="csvInputReact" type="file" accept=".xlsx,.xls,.csv,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" />
            <div className="import-actions">
              <button className="primary-btn" type="button" onClick={() => previewCsv(document.querySelector("#csvInputReact").files[0])}>预览</button>
              <button className="ghost-btn" type="button" disabled={!pendingImportRows.some((row) => row.status === "ready")} onClick={confirmImportRows}>确认导入</button>
              <button className="ghost-btn" type="button" onClick={downloadTemplate}>下载模板</button>
            </div>
            <ImportPreview rows={pendingImportRows} getAccountName={getAccountName} getTripName={getTripName} />
            <pre className="import-log">{importLog}</pre>
          </section>
        )}

        {activeView === "settings" && (
          <div className="two-column">
            <section className="panel">
              <h3>云同步设置</h3>
              <p className="muted">配置 Firebase 后，手机和电脑会使用同一份账本。自动上传默认开启。</p>
              <label>Firebase 配置 JSON<textarea rows="9" value={firebaseConfigText} onChange={(event) => setFirebaseConfigText(event.target.value)} /></label>
              <div className="form-actions">
                <button className="primary-btn" type="button" onClick={saveFirebaseConfig}>保存配置</button>
                <button className="ghost-btn" type="button" onClick={() => { localStorage.removeItem(firebaseConfigKey); setFirebaseConfigText(""); setSyncStatus({ title: "未连接", text: "当前只保存在本机浏览器。", tone: "offline" }); }}>清除配置</button>
              </div>
            </section>
            <section className="panel">
              <h3>账号与数据</h3>
              <div className="form-grid">
                <label>邮箱<input type="email" value={syncAuth.email} onChange={(event) => setSyncAuth({ ...syncAuth, email: event.target.value })} /></label>
                <label>密码<input type="password" value={syncAuth.password} onChange={(event) => setSyncAuth({ ...syncAuth, password: event.target.value })} /></label>
              </div>
              <div className="form-actions">
                <button className="primary-btn" type="button" onClick={() => loginCloud("login")}>登录</button>
                <button className="ghost-btn" type="button" onClick={() => loginCloud("register")}>注册</button>
                <button className="ghost-btn" type="button" onClick={async () => firebaseClient?.signOut(firebaseClient.auth)}>退出</button>
              </div>
              <div className="sync-status-panel">
                <span className={`sync-dot ${syncStatus.tone === "error" ? "error" : syncStatus.tone === "offline" ? "offline" : ""}`}></span>
                <div><strong>{syncStatus.title}</strong><p>{syncStatus.text}</p></div>
              </div>
              <label className="checkbox-line">
                <input type="checkbox" checked={autoCloudSync} onChange={(event) => { setAutoCloudSync(event.target.checked); localStorage.setItem(autoCloudSyncKey, event.target.checked ? "true" : "false"); }} />
                <span>保存记账数据后自动上传到云端</span>
              </label>
              <div className="form-actions">
                <button className="primary-btn" type="button" onClick={() => pushCloudState()}>上传本地数据</button>
                <button className="ghost-btn" type="button" onClick={pullCloudState}>拉取云端数据</button>
              </div>
              <pre className="import-log">{syncLog}</pre>
            </section>
          </div>
        )}
      </main>
    </div>
  );
}

function Metric({ label, value, accent }) {
  return <article className={`metric ${accent ? "accent" : ""}`}><span>{label}</span><strong>{value}</strong></article>;
}

function BudgetSnapshot({ totalBudget, spent, remaining, rate, onOpen }) {
  if (!totalBudget) {
    return (
      <section className="panel budget-snapshot">
        <div>
          <h3>本月预算</h3>
          <p className="muted">还没有设置预算。设置后首页会显示使用率和剩余额度。</p>
        </div>
        <button className="primary-btn" type="button" onClick={onOpen}>设置预算</button>
      </section>
    );
  }
  const over = remaining < 0;
  return (
    <section className="panel budget-snapshot">
      <div className="budget-snapshot-main">
        <div className="panel-head">
          <h3>本月预算</h3>
          <button className="text-btn" type="button" onClick={onOpen}>调整</button>
        </div>
        <div className="budget-line">
          <span>已用 {money(spent)} / {money(totalBudget)}</span>
          <strong className={over ? "money-expense" : "money-income"}>{over ? "超出" : "剩余"} {money(Math.abs(remaining))}</strong>
        </div>
        <ProgressBar value={rate} danger={over} />
      </div>
    </section>
  );
}

function ProgressBar({ value, danger }) {
  const width = Math.max(0, Math.min(value, 100));
  return (
    <div className="budget-progress" aria-label={`预算使用率 ${Math.round(value)}%`}>
      <div className={`budget-progress-fill ${danger ? "danger" : ""}`} style={{ width: `${width}%` }}></div>
    </div>
  );
}

function BudgetPage({ categories, budgets, expenseByCategory, totalSpent, onSave }) {
  const totalBudget = Number(budgets.total || 0);
  const remaining = totalBudget ? totalBudget - totalSpent : 0;
  return (
    <form className="panel budget-page" onSubmit={onSave}>
      <div className="panel-head">
        <div>
          <h3>月度预算</h3>
          <p className="muted">预算会按每个月循环使用。分类预算不填则不限制该分类。</p>
        </div>
        <button className="primary-btn" type="submit">保存预算</button>
      </div>

      <div className="budget-total-grid">
        <label>
          月度总预算
          <input name="totalBudget" type="number" min="0" step="0.01" defaultValue={totalBudget || ""} placeholder="例如：5000" />
        </label>
        <article className="metric">
          <span>本月已用</span>
          <strong>{money(totalSpent)}</strong>
        </article>
        <article className="metric accent">
          <span>{remaining < 0 ? "已超出" : "总预算剩余"}</span>
          <strong className={remaining < 0 ? "money-expense" : ""}>{money(Math.abs(remaining))}</strong>
        </article>
      </div>

      <div className="budget-category-list">
        {categories.map((category) => {
          const budget = Number(budgets.categories?.[category] || 0);
          const spent = Number(expenseByCategory[category] || 0);
          const categoryRemaining = budget ? budget - spent : 0;
          const rate = budget ? (spent / budget) * 100 : 0;
          const over = budget > 0 && categoryRemaining < 0;
          return (
            <div className="budget-category-row" key={category}>
              <div>
                <strong>{category}</strong>
                <p className="transaction-meta">
                  已用 {money(spent)}
                  {budget ? ` · ${over ? "超出" : "剩余"} ${money(Math.abs(categoryRemaining))}` : " · 未设置预算"}
                </p>
                {budget ? <ProgressBar value={rate} danger={over} /> : null}
              </div>
              <label>
                分类预算
                <input name={`budget-${category}`} type="number" min="0" step="0.01" defaultValue={budget || ""} placeholder="不限" />
              </label>
            </div>
          );
        })}
      </div>
    </form>
  );
}

function getTransactionAttachments(item, attachments) {
  const ids = item.attachmentIds || [];
  return attachments.filter((attachment) => ids.includes(attachment.id));
}

function AttachmentLinks({ files }) {
  if (!files.length) return <span>-</span>;
  return (
    <div className="attachment-links">
      {files.map((file, index) => (
        <a key={file.id} href={file.dataUrl} target="_blank" rel="noreferrer" title={file.name}>
          凭证{index + 1}
        </a>
      ))}
    </div>
  );
}

function TransactionCard({ item, attachments, getAccountName, getTripName }) {
  const cls = item.type === "income" ? "money-income" : "money-expense";
  const files = getTransactionAttachments(item, attachments);
  return (
    <article className="transaction-item">
      <div className="transaction-row"><strong>{item.category}</strong><strong className={cls}>{item.type === "income" ? "+" : "-"}{money(item.amount)}</strong></div>
      <div className="transaction-meta">{item.occurredAt} · {getAccountName(item.account)}{item.isTrip ? ` · ${getTripName(item.tripId)}` : ""}</div>
      <div className="transaction-meta">{item.merchant || item.note || "无备注"}{files.length ? ` · ${files.length} 个凭证` : ""}</div>
      {files.length ? <AttachmentLinks files={files} /> : null}
    </article>
  );
}

function Bars({ rows }) {
  if (!rows.length) return <div className="empty">暂无支出数据。</div>;
  const max = Math.max(...rows.map(([, amount]) => amount), 1);
  return <div className="bars">{rows.map(([label, amount]) => <div className="bar-row" key={label}><div className="bar-label"><span>{label}</span><strong>{money(amount)}</strong></div><div className="bar-track"><div className="bar-fill" style={{ width: `${Math.max((amount / max) * 100, 4)}%` }}></div></div></div>)}</div>;
}

function ReimbursementStatusSelect({ value, onChange }) {
  return (
    <select className="status-select" value={value || "none"} onChange={(event) => onChange(event.target.value)}>
      <option value="pending">待整理</option>
      <option value="submitted">已提交</option>
      <option value="reimbursed">已报销</option>
      <option value="none">不报销</option>
    </select>
  );
}

function TransactionTable({ rows, attachments, getAccountName, onDelete, onStatusChange }) {
  return (
    <div className="table-wrap"><table><thead><tr><th>日期</th><th>类型</th><th>分类</th><th>商户/备注</th><th>账户</th><th>报销</th><th>凭证</th><th className="num">金额</th><th></th></tr></thead><tbody>
      {rows.length ? rows.map((row) => {
        const files = getTransactionAttachments(row, attachments);
        return <tr key={row.id}><td>{row.occurredAt}</td><td>{row.type === "income" ? "收入" : "支出"}</td><td>{row.category}</td><td>{row.merchant || row.note || "-"}</td><td>{getAccountName(row.account)}</td><td>{row.type === "expense" ? <ReimbursementStatusSelect value={row.reimbursementStatus} onChange={(status) => onStatusChange(row.id, status)} /> : "-"}</td><td><AttachmentLinks files={files} /></td><td className={`num ${row.type === "income" ? "money-income" : "money-expense"}`}>{row.type === "income" ? "+" : "-"}{money(row.amount)}</td><td><button className="delete-btn" type="button" onClick={() => onDelete(row.id)}>删除</button></td></tr>;
      }) : <tr><td colSpan="9">没有匹配的流水。</td></tr>}
    </tbody></table></div>
  );
}

function TripList({ trips, transactions, onBatchStatusChange }) {
  if (!trips.length) return <div className="empty">还没有差旅项目。</div>;
  return <div className="trip-list">{trips.map((trip) => {
    const rows = transactions.filter((item) => !item.deletedAt && item.tripId === trip.id);
    const total = rows.filter((item) => item.type === "expense").reduce((sum, item) => sum + Number(item.amount), 0);
    const pending = rows.filter((item) => item.type === "expense" && ["pending", "submitted"].includes(item.reimbursementStatus)).reduce((sum, item) => sum + Number(item.amount), 0);
    const reimbursed = rows.filter((item) => item.type === "expense" && item.reimbursementStatus === "reimbursed").reduce((sum, item) => sum + Number(item.amount), 0);
    const hasExpenses = rows.some((item) => item.type === "expense");
    return <article className="trip-item" key={trip.id}><div className="transaction-row"><strong>{trip.name}</strong><span className="transaction-meta">{trip.startDate || "-"} 至 {trip.endDate || "-"}</span></div><div className="transaction-meta">{trip.destination || "未填地点"} · {trip.purpose || "未填事由"}</div><div className="trip-stats"><div className="trip-stat"><span>总支出</span><strong>{money(total)}</strong></div><div className="trip-stat"><span>待报销</span><strong>{money(pending)}</strong></div><div className="trip-stat"><span>已报销</span><strong>{money(reimbursed)}</strong></div><div className="trip-stat"><span>流水数</span><strong>{rows.length}</strong></div></div><div className="trip-actions"><button className="ghost-btn" type="button" disabled={!hasExpenses} onClick={() => onBatchStatusChange(trip.id, "submitted")}>全部标记已提交</button><button className="primary-btn" type="button" disabled={!hasExpenses} onClick={() => onBatchStatusChange(trip.id, "reimbursed")}>全部标记已报销</button></div></article>;
  })}</div>;
}

function AccountBalances({ state }) {
  return <div className="account-list">{state.accounts.map((account) => {
    const delta = state.transactions.filter((transaction) => !transaction.deletedAt && transaction.account === account.id).reduce((sum, transaction) => sum + (transaction.type === "income" ? Number(transaction.amount) : -Number(transaction.amount)), 0);
    return <div className="account-item" key={account.id}><div className="transaction-row"><strong>{account.name}</strong><strong>{money(Number(account.initialBalance || 0) + delta)}</strong></div><div className="transaction-meta">{account.type}</div></div>;
  })}</div>;
}

function MonthChart({ transactions }) {
  const months = [];
  const now = new Date();
  for (let index = 5; index >= 0; index -= 1) {
    const date = new Date(now.getFullYear(), now.getMonth() - index, 1);
    months.push({ key: date.toISOString().slice(0, 7), label: `${date.getMonth() + 1}月`, income: 0, expense: 0 });
  }
  transactions.filter((transaction) => !transaction.deletedAt).forEach((transaction) => {
    const bucket = months.find((month) => month.key === String(transaction.occurredAt).slice(0, 7));
    if (bucket) bucket[transaction.type === "income" ? "income" : "expense"] += Number(transaction.amount);
  });
  const max = Math.max(...months.flatMap((month) => [month.income, month.expense]), 1);
  return <div className="month-chart">{months.map((month) => <div className="month-bar" key={month.key}><div className="month-stack"><div className="income-bar" style={{ height: `${Math.max((month.income / max) * 170, month.income ? 4 : 0)}px` }}></div><div className="expense-bar" style={{ height: `${Math.max((month.expense / max) * 170, month.expense ? 4 : 0)}px` }}></div></div><div className="month-label">{month.label}</div></div>)}</div>;
}

function ImportPreview({ rows, getAccountName, getTripName }) {
  const ready = rows.filter((row) => row.status === "ready").length;
  const duplicate = rows.filter((row) => row.status === "duplicate").length;
  const invalid = rows.filter((row) => row.status === "invalid").length;
  return <>
    <div className="import-summary">{rows.length ? `预览 ${rows.length} 条：可导入 ${ready} 条，重复 ${duplicate} 条，无效 ${invalid} 条。` : ""}</div>
    <div className="table-wrap import-preview"><table><thead><tr><th>状态</th><th>日期</th><th>类型</th><th>分类</th><th>差旅</th><th>商户/说明</th><th>账户</th><th className="num">金额</th></tr></thead><tbody>
      {rows.length ? rows.slice(0, 100).map((row, index) => <tr key={`${row.importKey}-${index}`}><td><span className={row.status === "ready" ? "status-pill" : "status-pill skip"}>{row.status === "ready" ? "可导入" : row.status === "duplicate" ? "重复" : "无效"}</span></td><td>{row.occurredAt || "-"}</td><td>{row.type === "income" ? "收入" : "支出"}</td><td>{row.category || "-"}</td><td>{row.isTrip ? getTripName(row.tripId) : "-"}</td><td>{row.merchant || row.note || "-"}</td><td>{getAccountName(row.account)}</td><td className="num">{money(row.amount)}</td></tr>) : <tr><td colSpan="8">选择账单文件后点击预览。</td></tr>}
    </tbody></table></div>
  </>;
}

createRoot(document.getElementById("root")).render(<App />);
