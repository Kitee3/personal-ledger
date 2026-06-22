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
  transactions: [
    {
      id: crypto.randomUUID(),
      amount: 38,
      type: "expense",
      category: "餐饮",
      account: "wechat",
      occurredAt: today(),
      merchant: "午餐",
      note: "示例流水，可删除",
      isTrip: false,
      tripId: "",
      reimbursementStatus: "none",
      source: "manual",
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    },
  ],
  trips: [],
  lastAccount: "wechat",
};

let state = loadState();
let activeType = "expense";
let firebaseClient = null;
let cloudUser = null;
let cloudBusy = false;
let pendingImportRows = [];

const views = {
  dashboard: document.querySelector("#dashboardView"),
  quick: document.querySelector("#quickView"),
  transactions: document.querySelector("#transactionsView"),
  trips: document.querySelector("#tripsView"),
  reports: document.querySelector("#reportsView"),
  import: document.querySelector("#importView"),
  sync: document.querySelector("#syncView"),
};

const viewTitles = {
  dashboard: "首页",
  quick: "快速记账",
  transactions: "流水",
  trips: "差旅",
  reports: "报表",
  import: "账单导入",
  sync: "云同步",
};

function loadState() {
  const raw = localStorage.getItem(storageKey);
  if (!raw) return structuredClone(defaultState);
  try {
    return { ...structuredClone(defaultState), ...JSON.parse(raw) };
  } catch {
    return structuredClone(defaultState);
  }
}

function saveState() {
  localStorage.setItem(storageKey, JSON.stringify(state));
  if (shouldAutoSync()) {
    pushCloudState({ silent: true });
  }
}

function getStoredFirebaseConfig() {
  const raw = localStorage.getItem(firebaseConfigKey);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function shouldAutoSync() {
  return localStorage.getItem(autoCloudSyncKey) === "true" && cloudUser && firebaseClient && !cloudBusy;
}

function setSyncLog(message) {
  const log = document.querySelector("#syncLog");
  if (log) log.textContent = message;
}

function setCloudStatus(title, text, tone = "offline") {
  const titleEl = document.querySelector("#cloudStatusTitle");
  const textEl = document.querySelector("#cloudStatusText");
  const dot = document.querySelector("#cloudStatusDot");
  if (!titleEl || !textEl || !dot) return;
  titleEl.textContent = title;
  textEl.textContent = text;
  dot.classList.toggle("offline", tone === "offline");
  dot.classList.toggle("error", tone === "error");
}

async function initFirebaseClient() {
  if (firebaseClient) return firebaseClient;
  const config = getStoredFirebaseConfig();
  if (!config?.apiKey || !config?.projectId) {
    throw new Error("请先保存有效的 Firebase 配置 JSON。");
  }

  const [{ getApps, initializeApp }, authModule, firestoreModule] = await Promise.all([
    import("https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js"),
    import("https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js"),
    import("https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js"),
  ]);

  const app = getApps()[0] || initializeApp(config);
  const auth = authModule.getAuth(app);
  const db = firestoreModule.getFirestore(app);

  authModule.onAuthStateChanged(auth, (user) => {
    cloudUser = user;
    if (user) {
      setCloudStatus("已连接", `当前账号：${user.email || user.uid}`, "online");
      setSyncLog("已登录，可以上传或拉取云端账本。");
    } else {
      setCloudStatus("未登录", "Firebase 已配置，但还没有登录账号。", "offline");
    }
  });

  firebaseClient = {
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
  return firebaseClient;
}

function getCloudDocRef(client) {
  if (!cloudUser) throw new Error("请先登录 Firebase 账号。");
  return client.doc(client.db, "users", cloudUser.uid, "ledger", "main");
}

function getCloudPayload() {
  return {
    state,
    schemaVersion: 1,
    clientUpdatedAt: new Date().toISOString(),
  };
}

function normalizeCloudState(nextState) {
  return {
    ...structuredClone(defaultState),
    ...nextState,
    accounts: Array.isArray(nextState?.accounts) ? nextState.accounts : structuredClone(defaultState.accounts),
    transactions: Array.isArray(nextState?.transactions) ? nextState.transactions : [],
    trips: Array.isArray(nextState?.trips) ? nextState.trips : [],
  };
}

async function saveFirebaseConfig() {
  const input = document.querySelector("#firebaseConfig");
  try {
    const config = JSON.parse(input.value.trim());
    if (!config.apiKey || !config.projectId || !config.authDomain) {
      throw new Error("配置至少需要 apiKey、authDomain 和 projectId。");
    }
    localStorage.setItem(firebaseConfigKey, JSON.stringify(config));
    firebaseClient = null;
    setSyncLog("Firebase 配置已保存。下一步可以注册或登录。");
    await initFirebaseClient();
  } catch (error) {
    setCloudStatus("配置错误", error.message, "error");
    setSyncLog(`配置保存失败：${error.message}`);
  }
}

function clearFirebaseConfig() {
  localStorage.removeItem(firebaseConfigKey);
  firebaseClient = null;
  cloudUser = null;
  document.querySelector("#firebaseConfig").value = "";
  setCloudStatus("未连接", "当前只保存在本机浏览器。", "offline");
  setSyncLog("Firebase 配置已清除。");
}

async function loginCloud() {
  try {
    const client = await initFirebaseClient();
    const email = document.querySelector("#syncEmail").value.trim();
    const password = document.querySelector("#syncPassword").value;
    await client.signInWithEmailAndPassword(client.auth, email, password);
  } catch (error) {
    setCloudStatus("登录失败", error.message, "error");
    setSyncLog(`登录失败：${error.message}`);
  }
}

async function registerCloud() {
  try {
    const client = await initFirebaseClient();
    const email = document.querySelector("#syncEmail").value.trim();
    const password = document.querySelector("#syncPassword").value;
    await client.createUserWithEmailAndPassword(client.auth, email, password);
  } catch (error) {
    setCloudStatus("注册失败", error.message, "error");
    setSyncLog(`注册失败：${error.message}`);
  }
}

async function logoutCloud() {
  try {
    const client = await initFirebaseClient();
    await client.signOut(client.auth);
    cloudUser = null;
    setCloudStatus("未登录", "Firebase 已配置，但还没有登录账号。", "offline");
    setSyncLog("已退出云同步账号。");
  } catch (error) {
    setSyncLog(`退出失败：${error.message}`);
  }
}

async function pushCloudState(options = {}) {
  if (cloudBusy) return;
  cloudBusy = true;
  try {
    const client = await initFirebaseClient();
    const ref = getCloudDocRef(client);
    await client.setDoc(ref, {
      ...getCloudPayload(),
      serverUpdatedAt: client.serverTimestamp(),
    });
    if (!options.silent) setSyncLog(`上传完成：${new Date().toLocaleString("zh-CN")}`);
  } catch (error) {
    if (!options.silent) {
      setCloudStatus("上传失败", error.message, "error");
      setSyncLog(`上传失败：${error.message}`);
    }
  } finally {
    cloudBusy = false;
  }
}

async function pullCloudState() {
  if (cloudBusy) return;
  cloudBusy = true;
  try {
    const client = await initFirebaseClient();
    const snapshot = await client.getDoc(getCloudDocRef(client));
    if (!snapshot.exists()) {
      setSyncLog("云端还没有账本。可以先上传本地数据。");
      return;
    }
    const payload = snapshot.data();
    state = normalizeCloudState(payload.state);
    localStorage.setItem(storageKey, JSON.stringify(state));
    render();
    setSyncLog(`拉取完成：${payload.clientUpdatedAt || "未知时间"}`);
  } catch (error) {
    setCloudStatus("拉取失败", error.message, "error");
    setSyncLog(`拉取失败：${error.message}`);
  } finally {
    cloudBusy = false;
  }
}

function initSyncView() {
  const config = getStoredFirebaseConfig();
  document.querySelector("#firebaseConfig").value = config ? JSON.stringify(config, null, 2) : "";
  document.querySelector("#autoCloudSync").checked = localStorage.getItem(autoCloudSyncKey) === "true";
  if (config) {
    setCloudStatus("未登录", "Firebase 已配置，但还没有登录账号。", "offline");
    initFirebaseClient().catch((error) => {
      setCloudStatus("初始化失败", error.message, "error");
      setSyncLog(`初始化失败：${error.message}`);
    });
  } else {
    setCloudStatus("未连接", "当前只保存在本机浏览器。", "offline");
  }
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

function money(value) {
  return `¥${Number(value || 0).toLocaleString("zh-CN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}

function byDateDesc(a, b) {
  return `${b.occurredAt} ${b.updatedAt || ""}`.localeCompare(`${a.occurredAt} ${a.updatedAt || ""}`);
}

function currentMonthRange() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const end = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  return [start, end];
}

function inCurrentMonth(transaction) {
  const [start, end] = currentMonthRange();
  const date = new Date(`${transaction.occurredAt}T00:00:00`);
  return date >= start && date < end;
}

function getAccountName(id) {
  return state.accounts.find((account) => account.id === id)?.name || id || "未设置";
}

function getTripName(id) {
  return state.trips.find((trip) => trip.id === id)?.name || "未关联项目";
}

function switchView(name) {
  Object.entries(views).forEach(([key, element]) => element.classList.toggle("active", key === name));
  document.querySelectorAll(".nav-item").forEach((button) => {
    button.classList.toggle("active", button.dataset.view === name);
  });
  document.querySelector("#pageTitle").textContent = viewTitles[name];
  render();
}

function setActiveType(type) {
  activeType = type;
  document.querySelectorAll(".segment").forEach((button) => {
    button.classList.toggle("active", button.dataset.type === type);
  });
  renderCategoryOptions();
}

function renderCategoryOptions() {
  const select = document.querySelector("#category");
  const categories = activeType === "expense" ? state.expenseCategories : state.incomeCategories;
  select.innerHTML = categories.map((category) => `<option value="${category}">${category}</option>`).join("");
}

function renderAccountOptions() {
  const accountSelect = document.querySelector("#account");
  accountSelect.innerHTML = state.accounts
    .map((account) => `<option value="${account.id}">${account.name}</option>`)
    .join("");
  accountSelect.value = state.lastAccount || state.accounts[0].id;
}

function renderTripOptions() {
  const tripSelect = document.querySelector("#tripId");
  const openTrips = state.trips.filter((trip) => trip.status !== "reimbursed");
  tripSelect.innerHTML = [
    `<option value="">选择差旅项目</option>`,
    ...openTrips.map((trip) => `<option value="${trip.id}">${trip.name}</option>`),
  ].join("");

  const importTripSelect = document.querySelector("#importTripId");
  if (importTripSelect) {
    const previousValue = importTripSelect.value;
    importTripSelect.innerHTML = [
      `<option value="">不自动关联差旅</option>`,
      ...openTrips.map((trip) => `<option value="${trip.id}">${trip.name}</option>`),
    ].join("");
    importTripSelect.value = openTrips.some((trip) => trip.id === previousValue) ? previousValue : "";
  }
}

function getMonthTransactions() {
  return state.transactions.filter((transaction) => !transaction.deletedAt && inCurrentMonth(transaction));
}

function summarizeTransactions(transactions) {
  return transactions.reduce(
    (summary, transaction) => {
      if (transaction.type === "income") summary.income += Number(transaction.amount);
      if (transaction.type === "expense") summary.expense += Number(transaction.amount);
      if (
        transaction.isTrip &&
        transaction.type === "expense" &&
        ["pending", "submitted"].includes(transaction.reimbursementStatus)
      ) {
        summary.pending += Number(transaction.amount);
      }
      return summary;
    },
    { income: 0, expense: 0, pending: 0 }
  );
}

function groupExpenseByCategory(transactions) {
  const groups = new Map();
  transactions
    .filter((transaction) => transaction.type === "expense")
    .forEach((transaction) => {
      groups.set(transaction.category, (groups.get(transaction.category) || 0) + Number(transaction.amount));
    });
  return [...groups.entries()].sort((a, b) => b[1] - a[1]);
}

function renderDashboard() {
  const monthTransactions = getMonthTransactions();
  const summary = summarizeTransactions(monthTransactions);
  document.querySelector("#monthIncome").textContent = money(summary.income);
  document.querySelector("#monthExpense").textContent = money(summary.expense);
  document.querySelector("#monthBalance").textContent = money(summary.income - summary.expense);
  document.querySelector("#pendingReimburse").textContent = money(summary.pending);

  const recent = state.transactions.filter((transaction) => !transaction.deletedAt).sort(byDateDesc).slice(0, 10);
  document.querySelector("#recentTransactions").innerHTML = recent.length
    ? recent.map(renderTransactionCard).join("")
    : `<div class="empty">还没有流水，先记一笔。</div>`;

  renderCategoryBars("#topCategories", groupExpenseByCategory(monthTransactions).slice(0, 5));
}

function renderTransactionCard(transaction) {
  const cls = transaction.type === "income" ? "money-income" : "money-expense";
  const sign = transaction.type === "income" ? "+" : "-";
  const trip = transaction.isTrip ? ` · ${getTripName(transaction.tripId)}` : "";
  return `
    <article class="transaction-item">
      <div class="transaction-row">
        <strong>${transaction.category}</strong>
        <strong class="${cls}">${sign}${money(transaction.amount)}</strong>
      </div>
      <div class="transaction-meta">${transaction.occurredAt} · ${getAccountName(transaction.account)}${trip}</div>
      <div class="transaction-meta">${transaction.merchant || transaction.note || "无备注"}</div>
    </article>
  `;
}

function renderCategoryBars(selector, groups) {
  const container = document.querySelector(selector);
  if (!groups.length) {
    container.innerHTML = `<div class="empty">暂无支出数据。</div>`;
    return;
  }
  const max = Math.max(...groups.map(([, amount]) => amount), 1);
  container.innerHTML = groups
    .map(([category, amount]) => {
      const width = Math.max((amount / max) * 100, 4);
      return `
        <div class="bar-row">
          <div class="bar-label"><span>${category}</span><strong>${money(amount)}</strong></div>
          <div class="bar-track"><div class="bar-fill" style="width:${width}%"></div></div>
        </div>
      `;
    })
    .join("");
}

function renderTransactions() {
  const search = document.querySelector("#searchInput").value.trim().toLowerCase();
  const type = document.querySelector("#typeFilter").value;
  const trip = document.querySelector("#tripFilter").value;
  const rows = state.transactions
    .filter((transaction) => !transaction.deletedAt)
    .filter((transaction) => (type === "all" ? true : transaction.type === type))
    .filter((transaction) => (trip === "all" ? true : trip === "trip" ? transaction.isTrip : !transaction.isTrip))
    .filter((transaction) => {
      if (!search) return true;
      return [transaction.category, transaction.merchant, transaction.note, getAccountName(transaction.account)]
        .join(" ")
        .toLowerCase()
        .includes(search);
    })
    .sort(byDateDesc);

  document.querySelector("#transactionTable").innerHTML = rows.length
    ? rows.map(renderTransactionRow).join("")
    : `<tr><td colspan="8">没有匹配的流水。</td></tr>`;
}

function renderTransactionRow(transaction) {
  const typeText = transaction.type === "income" ? "收入" : "支出";
  const cls = transaction.type === "income" ? "money-income" : "money-expense";
  const statusMap = { none: "-", pending: "待整理", submitted: "已提交", reimbursed: "已报销" };
  return `
    <tr>
      <td>${transaction.occurredAt}</td>
      <td>${typeText}</td>
      <td>${transaction.category}</td>
      <td>${transaction.merchant || transaction.note || "-"}</td>
      <td>${getAccountName(transaction.account)}</td>
      <td>${transaction.isTrip ? statusMap[transaction.reimbursementStatus] || "待整理" : "-"}</td>
      <td class="num ${cls}">${transaction.type === "income" ? "+" : "-"}${money(transaction.amount)}</td>
      <td><button class="delete-btn" data-delete-id="${transaction.id}" type="button">删除</button></td>
    </tr>
  `;
}

function renderTrips() {
  renderTripOptions();
  const container = document.querySelector("#tripList");
  if (!state.trips.length) {
    container.innerHTML = `<div class="empty">还没有差旅项目。</div>`;
    return;
  }
  container.innerHTML = state.trips
    .map((trip) => {
      const transactions = state.transactions.filter((item) => !item.deletedAt && item.tripId === trip.id);
      const total = transactions
        .filter((item) => item.type === "expense")
        .reduce((sum, item) => sum + Number(item.amount), 0);
      const pending = transactions
        .filter((item) => item.type === "expense" && ["pending", "submitted"].includes(item.reimbursementStatus))
        .reduce((sum, item) => sum + Number(item.amount), 0);
      const reimbursed = transactions
        .filter((item) => item.type === "expense" && item.reimbursementStatus === "reimbursed")
        .reduce((sum, item) => sum + Number(item.amount), 0);
      return `
        <article class="trip-item">
          <div class="transaction-row">
            <strong>${trip.name}</strong>
            <span class="transaction-meta">${trip.startDate || "-"} 至 ${trip.endDate || "-"}</span>
          </div>
          <div class="transaction-meta">${trip.destination || "未填地点"} · ${trip.purpose || "未填事由"}</div>
          <div class="trip-stats">
            <div class="trip-stat"><span>总支出</span><strong>${money(total)}</strong></div>
            <div class="trip-stat"><span>待报销</span><strong>${money(pending)}</strong></div>
            <div class="trip-stat"><span>已报销</span><strong>${money(reimbursed)}</strong></div>
            <div class="trip-stat"><span>流水数</span><strong>${transactions.length}</strong></div>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderReports() {
  const monthTransactions = getMonthTransactions();
  renderCategoryBars("#reportCategoryBars", groupExpenseByCategory(monthTransactions));

  const balances = state.accounts.map((account) => {
    const delta = state.transactions
      .filter((transaction) => !transaction.deletedAt && transaction.account === account.id)
      .reduce((sum, transaction) => {
        return sum + (transaction.type === "income" ? Number(transaction.amount) : -Number(transaction.amount));
      }, 0);
    return { ...account, balance: Number(account.initialBalance || 0) + delta };
  });

  document.querySelector("#accountBalances").innerHTML = balances
    .map(
      (account) => `
        <div class="account-item">
          <div class="transaction-row">
            <strong>${account.name}</strong>
            <strong>${money(account.balance)}</strong>
          </div>
          <div class="transaction-meta">${account.type}</div>
        </div>
      `
    )
    .join("");

  renderMonthChart();
}

function renderMonthChart() {
  const months = [];
  const now = new Date();
  for (let index = 5; index >= 0; index -= 1) {
    const date = new Date(now.getFullYear(), now.getMonth() - index, 1);
    const key = date.toISOString().slice(0, 7);
    months.push({ key, label: `${date.getMonth() + 1}月`, income: 0, expense: 0 });
  }

  state.transactions
    .filter((transaction) => !transaction.deletedAt)
    .forEach((transaction) => {
      const bucket = months.find((month) => month.key === transaction.occurredAt.slice(0, 7));
      if (!bucket) return;
      bucket[transaction.type === "income" ? "income" : "expense"] += Number(transaction.amount);
    });

  const max = Math.max(...months.flatMap((month) => [month.income, month.expense]), 1);
  document.querySelector("#monthChart").innerHTML = months
    .map((month) => {
      const incomeHeight = Math.max((month.income / max) * 170, month.income ? 4 : 0);
      const expenseHeight = Math.max((month.expense / max) * 170, month.expense ? 4 : 0);
      return `
        <div class="month-bar">
          <div class="month-stack">
            <div class="income-bar" title="收入 ${money(month.income)}" style="height:${incomeHeight}px"></div>
            <div class="expense-bar" title="支出 ${money(month.expense)}" style="height:${expenseHeight}px"></div>
          </div>
          <div class="month-label">${month.label}</div>
        </div>
      `;
    })
    .join("");
}

function render() {
  renderCategoryOptions();
  renderAccountOptions();
  renderTripOptions();
  renderDashboard();
  renderTransactions();
  renderTrips();
  renderReports();
}

function resetTransactionForm() {
  document.querySelector("#transactionForm").reset();
  document.querySelector("#occurredAt").value = today();
  document.querySelector("#account").value = state.lastAccount || "wechat";
  document.querySelector("#tripFields").classList.remove("active");
  setActiveType("expense");
}

function addTransaction(event) {
  event.preventDefault();
  const amount = Number(document.querySelector("#amount").value);
  if (!amount || amount <= 0) return;

  const isTrip = document.querySelector("#isTrip").checked;
  const account = document.querySelector("#account").value;
  const transaction = {
    id: crypto.randomUUID(),
    amount,
    type: activeType,
    category: document.querySelector("#category").value,
    account,
    occurredAt: document.querySelector("#occurredAt").value || today(),
    merchant: document.querySelector("#merchant").value.trim(),
    note: document.querySelector("#note").value.trim(),
    isTrip,
    tripId: isTrip ? document.querySelector("#tripId").value : "",
    reimbursementStatus: isTrip ? document.querySelector("#reimbursementStatus").value : "none",
    source: "manual",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
  state.transactions.push(transaction);
  state.lastAccount = account;
  saveState();
  resetTransactionForm();
  switchView("dashboard");
}

function addTrip(event) {
  event.preventDefault();
  const name = document.querySelector("#tripName").value.trim();
  if (!name) return;
  state.trips.push({
    id: crypto.randomUUID(),
    name,
    destination: document.querySelector("#destination").value.trim(),
    startDate: document.querySelector("#startDate").value,
    endDate: document.querySelector("#endDate").value,
    purpose: document.querySelector("#purpose").value.trim(),
    status: "active",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  });
  saveState();
  document.querySelector("#tripForm").reset();
  render();
}

function deleteTransaction(id) {
  const item = state.transactions.find((transaction) => transaction.id === id);
  if (!item) return;
  item.deletedAt = new Date().toISOString();
  item.updatedAt = new Date().toISOString();
  saveState();
  render();
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

function exportJson() {
  download(`简账助手备份-${today()}.json`, JSON.stringify(state, null, 2), "application/json;charset=utf-8");
}

function toCsv(rows) {
  return rows
    .map((row) =>
      row
        .map((cell) => `"${String(cell ?? "").replaceAll('"', '""')}"`)
        .join(",")
    )
    .join("\n");
}

function exportTripCsv() {
  const rows = [["差旅项目", "日期", "分类", "商户", "备注", "账户", "报销状态", "金额"]];
  state.transactions
    .filter((transaction) => !transaction.deletedAt && transaction.isTrip)
    .sort(byDateDesc)
    .forEach((transaction) => {
      rows.push([
        getTripName(transaction.tripId),
        transaction.occurredAt,
        transaction.category,
        transaction.merchant,
        transaction.note,
        getAccountName(transaction.account),
        transaction.reimbursementStatus,
        transaction.amount,
      ]);
    });
  download(`差旅报销明细-${today()}.csv`, `\ufeff${toCsv(rows)}`, "text/csv;charset=utf-8");
}

function downloadTemplate() {
  const rows = [
    ["日期", "金额", "类型", "分类", "商户", "备注", "账户"],
    [today(), "25.5", "支出", "餐饮", "示例商户", "午餐", "微信"],
  ];
  download("账单导入模板.csv", `\ufeff${toCsv(rows)}`, "text/csv;charset=utf-8");
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
    lines.findIndex((line) => {
      const joined = parseLine(line).join("");
      return /日期|时间|金额|交易/.test(joined);
    })
  );
  const headers = parseLine(lines[headerIndex]);
  return lines.slice(headerIndex + 1).map((line) => {
    const cells = parseLine(line);
    return Object.fromEntries(headers.map((header, index) => [header, cells[index] || ""]));
  });
}

function getFirstValue(row, keys) {
  for (const key of keys) {
    if (row[key] !== undefined && String(row[key]).trim() !== "") return String(row[key]).trim();
  }
  return "";
}

function parseAmount(value) {
  const cleaned = String(value || "")
    .replace(/[￥¥,\s]/g, "")
    .replace(/[()]/g, "")
    .trim();
  const match = cleaned.match(/-?\d+(\.\d+)?/);
  return match ? Math.abs(Number(match[0])) : 0;
}

function normalizeType(row, amountRaw) {
  const marker = getFirstValue(row, ["类型", "收/支", "收支", "交易类型", "资金流向"]);
  if (/收入|收款|入账|\+|转入/.test(marker)) return "income";
  if (/支出|付款|消费|-|转出/.test(marker)) return "expense";
  if (String(amountRaw || "").trim().startsWith("-")) return "expense";
  return "expense";
}

function normalizeAccount(row, source) {
  const account = getFirstValue(row, ["账户", "支付方式", "收/付款方式", "付款方式", "支付渠道"]);
  if (/微信|零钱/.test(account) || source === "wechat") return "wechat";
  if (/支付宝|余额宝|花呗/.test(account) || source === "alipay") return "alipay";
  if (/信用卡|贷记卡/.test(account)) return "credit";
  if (/现金/.test(account)) return "cash";
  if (/银行|储蓄卡|借记卡|招商|建设|工商|农业|中国银行|交通|邮储|中信|浦发|民生|兴业|广发|平安/.test(account)) {
    return "bank";
  }
  return resolveAccount(account);
}

function detectImportSource(row) {
  const keys = Object.keys(row).join(" ");
  if (/微信|交易对方|商品|支付方式|当前状态/.test(keys)) return "wechat";
  if (/支付宝|商品说明|收\/付款方式|交易订单号|商家订单号/.test(keys)) return "alipay";
  return "generic";
}

function normalizeImportRow(row) {
  const source = detectImportSource(row);
  const selectedTripId = document.querySelector("#importTripId")?.value || "";
  const tripMode = document.querySelector("#importTripMode")?.value || "strong";
  const amountRaw = getFirstValue(row, ["金额", "金额(元)", "交易金额", "支出金额", "收入金额"]);
  const amount = parseAmount(amountRaw);
  const type = normalizeType(row, amountRaw);
  const occurredAt = getFirstValue(row, ["日期", "交易时间", "时间", "付款时间", "创建时间"]).slice(0, 10) || today();
  const merchant = getFirstValue(row, ["商户", "交易对方", "对方", "收款方", "付款方", "商品", "商品说明"]);
  const note = getFirstValue(row, ["备注", "说明", "交易分类", "交易类型", "商品", "商品说明"]);
  const category = guessCategory({ 商户: merchant, 备注: note, 分类: getFirstValue(row, ["分类"]) }, type);
  const account = normalizeAccount(row, source);
  const importKey = [occurredAt, type, amount.toFixed(2), merchant, note].join("|").toLowerCase();
  const tripMatched = selectedTripId && shouldAttachToTrip({ merchant, note, category, type, mode: tripMode });
  return {
    amount,
    type,
    category,
    account,
    occurredAt,
    merchant,
    note,
    source,
    importKey,
    isTrip: Boolean(tripMatched) || category === "差旅",
    tripId: tripMatched ? selectedTripId : "",
    reimbursementStatus: tripMatched || category === "差旅" ? "pending" : "none",
  };
}

function shouldAttachToTrip({ merchant, note, category, type, mode }) {
  if (type !== "expense") return false;
  if (mode === "all-expense") return true;
  const text = `${merchant || ""} ${note || ""} ${category || ""}`.toLowerCase();
  return /酒店|宾馆|航旅|携程|飞猪|机票|高铁|火车|动车|打车|滴滴|出租车|机场|车站/.test(text);
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

function guessCategory(row, type) {
  const text = `${row["商户"] || ""} ${row["备注"] || ""}`.toLowerCase();
  if (type === "income") return row["分类"] || "其他";
  if (/滴滴|打车|地铁|公交|火车|机票|高铁|taxi/.test(text)) return "交通";
  if (/酒店|宾馆|航旅|携程|飞猪/.test(text)) return "差旅";
  if (/美团|饿了么|餐|咖啡|饭|奶茶/.test(text)) return "餐饮";
  if (/京东|淘宝|天猫|拼多多|超市/.test(text)) return "购物";
  return row["分类"] || "其他";
}

function resolveAccount(name) {
  const account = state.accounts.find((item) => item.name === name || item.id === name);
  return account?.id || state.lastAccount || "wechat";
}

function renderImportPreview(rows) {
  const table = document.querySelector("#importPreviewTable");
  const summary = document.querySelector("#importSummary");
  const validCount = rows.filter((row) => row.status === "ready").length;
  const duplicateCount = rows.filter((row) => row.status === "duplicate").length;
  const invalidCount = rows.filter((row) => row.status === "invalid").length;
  summary.textContent = `预览 ${rows.length} 条：可导入 ${validCount} 条，重复 ${duplicateCount} 条，无效 ${invalidCount} 条。`;
  document.querySelector("#confirmImportBtn").disabled = validCount === 0;
  if (!rows.length) {
    table.innerHTML = `<tr><td colspan="8">没有识别到可预览的流水。</td></tr>`;
    return;
  }
  table.innerHTML = rows
    .slice(0, 100)
    .map((row) => {
      const statusText = row.status === "ready" ? "可导入" : row.status === "duplicate" ? "重复" : "无效";
      const pillClass = row.status === "ready" ? "status-pill" : "status-pill skip";
      return `
        <tr>
          <td><span class="${pillClass}">${statusText}</span></td>
          <td>${row.occurredAt || "-"}</td>
          <td>${row.type === "income" ? "收入" : "支出"}</td>
          <td>${row.category || "-"}</td>
          <td>${row.isTrip ? getTripName(row.tripId) : "-"}</td>
          <td>${row.merchant || row.note || "-"}</td>
          <td>${getAccountName(row.account)}</td>
          <td class="num">${money(row.amount)}</td>
        </tr>
      `;
    })
    .join("");
}

async function importCsv() {
  const file = document.querySelector("#csvInput").files[0];
  const log = document.querySelector("#importLog");
  if (!file) {
    log.textContent = "请先选择 CSV 文件。";
    return;
  }
  const text = await file.text();
  const rows = parseCsv(text).map(normalizeImportRow);
  pendingImportRows = rows.map((row) => ({
    ...row,
    status: row.amount > 0 ? (isDuplicateImport(row) ? "duplicate" : "ready") : "invalid",
  }));
  renderImportPreview(pendingImportRows);
  log.textContent = `预览完成：${file.name}。确认无误后点击“确认导入”。`;
}

function confirmImportRows() {
  const log = document.querySelector("#importLog");
  const rows = pendingImportRows.filter((row) => row.status === "ready");
  rows.forEach((row) => {
    state.transactions.push({
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
      importKey: row.importKey,
      source: row.source,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
  });
  saveState();
  render();
  pendingImportRows = pendingImportRows.map((row) => (row.status === "ready" ? { ...row, status: "duplicate" } : row));
  renderImportPreview(pendingImportRows);
  log.textContent = `导入完成：${rows.length} 条流水。重复记录已自动跳过，差旅相关记录会标记为待整理。`;
}

document.querySelectorAll(".nav-item").forEach((button) => {
  button.addEventListener("click", () => switchView(button.dataset.view));
});

document.querySelectorAll("[data-view-target]").forEach((button) => {
  button.addEventListener("click", () => switchView(button.dataset.viewTarget));
});

document.querySelectorAll(".segment").forEach((button) => {
  button.addEventListener("click", () => setActiveType(button.dataset.type));
});

document.querySelector("#quickAddBtn").addEventListener("click", () => switchView("quick"));
document.querySelector("#transactionForm").addEventListener("submit", addTransaction);
document.querySelector("#tripForm").addEventListener("submit", addTrip);
document.querySelector("#resetFormBtn").addEventListener("click", resetTransactionForm);
document.querySelector("#exportJsonBtn").addEventListener("click", exportJson);
document.querySelector("#exportTripCsvBtn").addEventListener("click", exportTripCsv);
document.querySelector("#downloadTemplateBtn").addEventListener("click", downloadTemplate);
document.querySelector("#importCsvBtn").addEventListener("click", importCsv);
document.querySelector("#confirmImportBtn").addEventListener("click", confirmImportRows);
document.querySelector("#importTripId").addEventListener("change", () => {
  if (document.querySelector("#csvInput").files[0]) importCsv();
});
document.querySelector("#importTripMode").addEventListener("change", () => {
  if (document.querySelector("#csvInput").files[0]) importCsv();
});
document.querySelector("#saveFirebaseConfigBtn").addEventListener("click", saveFirebaseConfig);
document.querySelector("#clearFirebaseConfigBtn").addEventListener("click", clearFirebaseConfig);
document.querySelector("#syncLoginBtn").addEventListener("click", loginCloud);
document.querySelector("#syncRegisterBtn").addEventListener("click", registerCloud);
document.querySelector("#syncLogoutBtn").addEventListener("click", logoutCloud);
document.querySelector("#pushCloudBtn").addEventListener("click", () => pushCloudState());
document.querySelector("#pullCloudBtn").addEventListener("click", pullCloudState);
document.querySelector("#autoCloudSync").addEventListener("change", (event) => {
  localStorage.setItem(autoCloudSyncKey, event.target.checked ? "true" : "false");
  setSyncLog(event.target.checked ? "已开启自动上传。" : "已关闭自动上传。");
});
document.querySelector("#isTrip").addEventListener("change", (event) => {
  document.querySelector("#tripFields").classList.toggle("active", event.target.checked);
});

["#searchInput", "#typeFilter", "#tripFilter"].forEach((selector) => {
  document.querySelector(selector).addEventListener("input", renderTransactions);
});

document.querySelector("#transactionTable").addEventListener("click", (event) => {
  const button = event.target.closest("[data-delete-id]");
  if (button) deleteTransaction(button.dataset.deleteId);
});

resetTransactionForm();
initSyncView();
const initialView = new URLSearchParams(window.location.search).get("view");
if (initialView && views[initialView]) {
  switchView(initialView);
} else {
  render();
}
