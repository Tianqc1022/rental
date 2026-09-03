/* viewer.js — 租房管家·云端查看器：上传 / 下载 / 只读查看最新 Excel 数据 */
'use strict';
(function () {
  var cfg = window.APP_CONFIG || {};
  var API = 'https://api.jsonbin.io/v3/b';
  var BIN_ID = cfg.BIN_ID || '';

  var STATUS_CN = { rented: '已出租', vacant: '空置' };
  var STATUS_EN = { '已出租': 'rented', '空置': 'vacant', '出租中': 'rented' };
  var TYPE_CN = { rent: '房租', deposit: '押金', other: '其他' };
  var TYPE_EN = { '房租': 'rent', '押金': 'deposit', '其他': 'other' };
  var METHOD_CN = { wechat: '微信', alipay: '支付宝', cash: '现金', bank: '银行转账' };
  var METHOD_EN = { '微信': 'wechat', '支付宝': 'alipay', '现金': 'cash', '银行转账': 'bank' };

  var state = null;   // { houses, payments }
  var tab = 'home';
  var els = {};

  function esc(s) { return utils.escapeHtml(s); }
  function $(sel) { return document.querySelector(sel); }
  function str(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
  function num(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
  function uid(p) { return utils.uid(p); }
  function isConfigured() { return !!(cfg.ACCESS_KEY || cfg.MASTER_KEY); }

  function authHeaders() {
    var h = {};
    if (cfg.ACCESS_KEY) h['X-Access-Key'] = cfg.ACCESS_KEY;
    else if (cfg.MASTER_KEY) h['X-Master-Key'] = cfg.MASTER_KEY;
    return h;
  }

  /* ================= jsonbin 云端 ================= */
  function fetchBin() {
    return fetch(API + '/' + BIN_ID + '/latest', { headers: authHeaders() })
      .then(function (r) {
        if (!r.ok) throw new Error('云端读取失败(' + r.status + ')');
        return r.json();
      })
      .then(function (j) { return j.record || null; });
  }

  function createBin(data) {
    return fetch(API, {
      method: 'POST',
      headers: Object.assign(authHeaders(), { 'Content-Type': 'application/json' }),
      body: JSON.stringify(data)
    }).then(function (r) {
      if (!r.ok) throw new Error('云端创建失败(' + r.status + ')');
      return r.json();
    }).then(function (j) { return (j.metadata && j.metadata.id) || ''; });
  }

  function updateBin(data) {
    return fetch(API + '/' + BIN_ID, {
      method: 'PUT',
      headers: Object.assign(authHeaders(), { 'Content-Type': 'application/json' }),
      body: JSON.stringify(data)
    }).then(function (r) {
      if (!r.ok) throw new Error('云端上传失败(' + r.status + ')');
      return r.json();
    });
  }

  /* ================= Excel 解析（与已验证逻辑一致） ================= */
  function fmtDate(v) {
    if (v === null || v === undefined || v === '') return '';
    if (v instanceof Date && !isNaN(v.getTime())) {
      return utils.fmtDate(new Date(v.getFullYear(), v.getMonth(), v.getDate(), 12));
    }
    if (typeof v === 'number') {
      var d = new Date(Math.round((v - 25569) * 86400000));
      d.setMinutes(d.getMinutes() + d.getTimezoneOffset());
      return utils.fmtDate(d);
    }
    var m = /^(\d{4})[^\d]+(\d{1,2})[^\d]+(\d{1,2})/.exec(String(v).trim());
    if (m) return m[1] + '-' + utils.pad2(parseInt(m[2], 10)) + '-' + utils.pad2(parseInt(m[3], 10));
    return String(v).trim();
  }

  function clampPayDay(v) {
    var n = parseInt(v, 10);
    if (isNaN(n)) return 1;
    return Math.min(31, Math.max(1, n));
  }

  /* ---- 付款周期 ---- */
  function periodText(m) {
    if (m === 3) return '每3个月';
    if (m === 6) return '每6个月';
    if (m === 12) return '每12个月';
    return '每月';
  }
  function periodFromText(t) {
    var s = str(t);
    if (s === '每3个月') return 3;
    if (s === '每6个月') return 6;
    if (s === '每12个月') return 12;
    return 1;
  }
  function isDateLike(v) { return /^\d{4}/.test(str(v)); }

  /* ---- 日期字符串推算（提醒用） ---- */
  function addMonthsStr(dateStr, nMonths) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateStr || '');
    if (!m) return '';
    var y = parseInt(m[1], 10), mo = parseInt(m[2], 10), d = parseInt(m[3], 10);
    var total = y * 12 + (mo - 1) + nMonths;
    var ny = Math.floor(total / 12), nm = (total % 12) + 1;
    var nd = Math.min(d, utils.lastDayOfMonth(ny, nm));
    return ny + '-' + utils.pad2(nm) + '-' + utils.pad2(nd);
  }
  function firstDueStr(startStr, payDay) {
    var s = utils.parseDate(startStr);
    if (!s) return '';
    var y = s.getFullYear(), m = s.getMonth() + 1, d = s.getDate();
    var eff = Math.min(payDay, utils.lastDayOfMonth(y, m));
    if (d <= eff) return y + '-' + utils.pad2(m) + '-' + utils.pad2(eff);
    var ny = (m === 12) ? y + 1 : y, nm = (m === 12) ? 1 : m + 1;
    return ny + '-' + utils.pad2(nm) + '-' + utils.pad2(Math.min(payDay, utils.lastDayOfMonth(ny, nm)));
  }
  function diffDaysStr(aStr, bStr) {
    var a = utils.parseDate(aStr), b = utils.parseDate(bStr);
    if (!a || !b) return 0;
    return Math.round((b.getTime() - a.getTime()) / 86400000);
  }

  function parseXlsx(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () {
        try {
          var wb = XLSX.read(new Uint8Array(reader.result), { type: 'array', cellDates: true });
          resolve(extract(wb));
        } catch (e) { reject(e); }
      };
      reader.onerror = function () { reject(new Error('读取文件失败')); };
      reader.readAsArrayBuffer(file);
    });
  }

  function extract(wb) {
    var houses = [], payments = [], skipped = 0;
    var hsSheet = wb.Sheets['房源'] || wb.Sheets[wb.SheetNames[0]];
    if (!hsSheet) throw new Error('未找到「房源」工作表');
    var psSheet = wb.Sheets['收款记录'] || (wb.SheetNames.length > 1 ? wb.Sheets[wb.SheetNames[1]] : null);

    var hs = XLSX.utils.sheet_to_json(hsSheet, { header: 1, raw: true, defval: '' });
    for (var i = 1; i < hs.length; i++) {
      var row = hs[i] || [];
      var name = str(row[0]);
      if (!name) continue;
      var st = STATUS_EN[str(row[2])] || 'vacant';
      var contract = null;
      if (st === 'rented') {
        // 新模板：第9列=付款周期；旧模板：第9列=合同开始（自动识别兼容）
        var oldLayout = isDateLike(row[8]);
        contract = {
          tenantName: str(row[3]), tenantPhone: str(row[4]),
          monthlyRent: num(row[5]), deposit: num(row[6]),
          payDay: clampPayDay(row[7]),
          payPeriod: oldLayout ? 1 : periodFromText(row[8]),
          startDate: fmtDate(oldLayout ? row[8] : row[9]),
          endDate: fmtDate(oldLayout ? row[9] : row[10]),
          remark: str(oldLayout ? row[10] : row[11])
        };
      }
      houses.push({ id: uid('h'), name: name, address: str(row[1]), status: st, contract: contract });
    }
    if (!houses.length) throw new Error('「房源」工作表中没有数据');

    var idmap = {};
    houses.forEach(function (h) { idmap[h.name] = h.id; });
    var ps = psSheet ? XLSX.utils.sheet_to_json(psSheet, { header: 1, raw: true, defval: '' }) : [];
    for (var j = 1; j < ps.length; j++) {
      var pr = ps[j] || [];
      var pname = str(pr[1]);
      var amt = num(pr[3]);
      if (!pname || !(amt > 0)) continue;
      var hid = idmap[pname];
      if (!hid) { skipped += 1; continue; }
      payments.push({
        id: uid('p'), houseId: hid,
        type: TYPE_EN[str(pr[2])] || 'rent',
        amount: amt, payDate: fmtDate(pr[0]),
        method: METHOD_EN[str(pr[4])] || 'wechat',
        remark: str(pr[5])
      });
    }
    return { houses: houses, payments: payments, skipped: skipped };
  }

  /* ================= Excel 生成（下载用） ================= */
  function buildXlsx(data) {
    var wb = XLSX.utils.book_new();
    var hr = [['房源名称', '地址', '状态', '租客姓名', '联系电话', '月租金(元)', '押金(元)',
      '每月几号收租', '付款周期', '合同开始', '合同结束', '备注']];
    data.houses.forEach(function (h) {
      var c = h.contract;
      hr.push([
        h.name, h.address, STATUS_CN[h.status] || h.status,
        c ? c.tenantName : '', c ? c.tenantPhone : '',
        c ? c.monthlyRent : '', c ? c.deposit : '',
        c ? c.payDay : '', c ? periodText(c.payPeriod) : '',
        c ? c.startDate : '', c ? c.endDate : '',
        c ? c.remark : ''
      ]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(hr), '房源');

    var pr = [['收款日期', '房源名称', '类型', '金额(元)', '收款方式', '备注']];
    data.payments.forEach(function (p) {
      var h = data.houses.find(function (x) { return x.id === p.houseId; });
      pr.push([
        p.payDate, h ? h.name : '', TYPE_CN[p.type] || p.type,
        p.amount, METHOD_CN[p.method] || p.method, p.remark || ''
      ]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(pr), '收款记录');
    return wb;
  }

  /* ================= 提醒计算 ================= */
  function computeReminders(data, now) {
    var items = [];
    (data.houses || []).forEach(function (h) {
      if (h.status !== 'rented' || !h.contract) return;
      var c = h.contract;
      var start = utils.parseDate(c.startDate);
      var end = utils.parseDate(c.endDate);

      if (end) {
        if (utils.dayKey(now) > utils.dayKey(end)) {
          items.push({ house: h, type: 'contract_expired', days: utils.diffDays(end, now), label: '合同已到期 ' + utils.diffDays(end, now) + ' 天' });
        } else if (utils.diffDays(now, end) <= 30) {
          items.push({ house: h, type: 'contract_expiring', days: utils.diffDays(now, end), label: utils.diffDays(now, end) === 0 ? '合同今天到期' : '合同 ' + utils.diffDays(now, end) + ' 天后到期' });
        }
      }

      // 房租：按「上次房租收款日 + 付款周期」推下次应缴日；没收到过按合同首期应缴日
      var period = c.payPeriod || 1;
      var lastPay = '';
      (data.payments || []).forEach(function (p) {
        if (p.houseId === h.id && p.type === 'rent' && p.payDate && p.payDate > lastPay) lastPay = p.payDate;
      });
      var nextDue = lastPay ? addMonthsStr(lastPay, period) : firstDueStr(c.startDate, c.payDay);
      if (nextDue && start) {
        var dToDue = diffDaysStr(utils.fmtDate(now), nextDue);
        if (dToDue < 0) {
          items.push({ house: h, type: 'rent_overdue', days: -dToDue, label: '房租已逾期 ' + (-dToDue) + ' 天' });
        } else if (dToDue <= 7) {
          items.push({ house: h, type: 'rent_due_soon', days: dToDue, label: dToDue === 0 ? '今天应收租' : dToDue + ' 天后收租' });
        }
      }
    });
    var rank = { rent_overdue: 0, contract_expired: 1, rent_due_soon: 2, contract_expiring: 3 };
    items.sort(function (a, b) { return (rank[a.type] - rank[b.type]) || (b.days - a.days); });
    return items;
  }

  /* ================= 渲染 ================= */
  function render() {
    if (!isConfigured()) { renderSetup(); return; }
    if (!state) { renderEmpty(); return; }
    els.lastUpdate.textContent = state.updatedAt ? '上次更新：' + fmtFull(state.updatedAt) : '';
    if (tab === 'home') renderHome();
    else if (tab === 'houses') renderHouses();
    else renderPayments();
  }

  function fmtFull(iso) {
    var d = new Date(iso);
    if (isNaN(d.getTime())) return String(iso);
    return (d.getMonth() + 1) + '月' + d.getDate() + '日 ' + utils.pad2(d.getHours()) + ':' + utils.pad2(d.getMinutes());
  }

  function renderSetup() {
    els.lastUpdate.textContent = '';
    $('#content').innerHTML =
      '<div class="card empty-state"><span class="emoji">🔑</span>' +
      '<div style="font-size:15px;color:#262626">还没有配置云端</div>' +
      '<div style="font-size:13px;line-height:1.8;margin-top:6px">请在 <b>js/config.js</b> 里填好 jsonbin 的密钥和仓库 ID。<br>具体步骤看项目里的 <b>README.md</b>。</div></div>';
  }

  function renderEmpty() {
    els.lastUpdate.textContent = '';
    $('#content').innerHTML =
      '<div class="card empty-state"><span class="emoji">📭</span>' +
      '<div style="font-size:15px;color:#262626">云端还没有数据</div>' +
      '<div style="font-size:13px;color:#8c8c8c;margin-top:6px">点上方「📤 上传 Excel」上传第一份数据（会自动创建云端仓库）。</div></div>';
  }

  function renderHome() {
    var now = new Date();
    var reminders = computeReminders(state, now);
    var mk = utils.monthKeyOf(now);
    var monthIncome = (state.payments || []).filter(function (p) {
      return p.type === 'rent' && p.payDate && p.payDate.indexOf(mk) === 0;
    }).reduce(function (s, p) { return s + p.amount; }, 0);
    var rented = (state.houses || []).filter(function (h) { return h.status === 'rented'; }).length;
    var vacant = (state.houses || []).length - rented;

    var html = '';
    html += '<div class="stats-grid">' +
      stat('', state.houses.length, '总房源') +
      stat('green', rented, '出租中') +
      stat('orange', vacant, '空置') +
      stat('primary', '¥' + utils.fmtMoney(monthIncome).replace(/\.00$/, ''), '本月已收') +
      '</div>';

    html += '<div class="section-label">待办提醒</div>';
    if (!reminders.length) {
      html += '<div class="card empty-state"><span class="emoji">😌</span>暂无待办，一切妥当</div>';
    } else {
      reminders.forEach(function (r) {
        var color = (r.type === 'rent_overdue' || r.type === 'contract_expired') ? 'danger' : (r.type === 'rent_due_soon' ? 'orange' : 'blue');
        html += '<div class="todo-card color-' + color + '">' +
          '<div class="todo-icon">' + (r.type.indexOf('rent') === 0 ? '¥' : '📄') + '</div>' +
          '<div class="todo-body">' +
            '<div class="todo-title">' + esc(r.house.name) + '</div>' +
            '<div class="todo-text">' + esc(r.label) + '</div>' +
            (r.house.contract ? '<div class="todo-sub">' + esc(periodText(r.house.contract.payPeriod)) + '收租 · 月租 ¥' + utils.fmtMoney(r.house.contract.monthlyRent) + '</div>' : '') +
          '</div>' +
        '</div>';
      });
    }

    html += '<div class="section-label">房源速览</div>';
    (state.houses || []).forEach(function (h) {
      var sub = (h.status === 'rented' && h.contract)
        ? (h.contract.tenantName || '') + ' · 月租 ¥' + utils.fmtMoney(h.contract.monthlyRent)
        : '暂无租客';
      html += '<div class="house-mini">' +
        '<span class="badge ' + h.status + '">' + STATUS_CN[h.status] + '</span>' +
        '<div class="body"><div class="name">' + esc(h.name) + '</div><div class="sub">' + esc(sub) + '</div></div>' +
        '</div>';
    });
    $('#content').innerHTML = html;
  }

  function stat(cls, n, label) {
    return '<div class="stat-card ' + cls + '"><div class="num">' + esc(n) + '</div><div class="label">' + esc(label) + '</div></div>';
  }

  function renderHouses() {
    var html = '<div class="section-label">房源（' + state.houses.length + '）</div>';
    html += '<div class="table-wrap"><table class="table">' +
      '<thead><tr><th>房源</th><th>状态</th><th>租客</th><th>月租</th><th>收租日</th><th>周期</th><th>合同至</th></tr></thead><tbody>';
    state.houses.forEach(function (h) {
      var c = h.contract;
      html += '<tr>' +
        '<td>' + esc(h.name) + '</td>' +
        '<td><span class="badge ' + h.status + '">' + STATUS_CN[h.status] + '</span></td>' +
        '<td>' + esc(c ? c.tenantName : '—') + '</td>' +
        '<td>' + (c ? '¥' + utils.fmtMoney(c.monthlyRent) : '—') + '</td>' +
        '<td>' + (c ? c.payDay + '号' : '—') + '</td>' +
        '<td>' + (c ? esc(periodText(c.payPeriod)) : '—') + '</td>' +
        '<td>' + (c ? esc(c.endDate || '') : '—') + '</td>' +
        '</tr>';
    });
    html += '</tbody></table></div>';
    $('#content').innerHTML = html;
  }

  function renderPayments() {
    var pays = (state.payments || []).slice().sort(function (a, b) {
      if (a.payDate !== b.payDate) return a.payDate < b.payDate ? 1 : -1;
      return (b.createdAt || 0) - (a.createdAt || 0);
    });
    var html = '<div class="section-label">收款记录（' + pays.length + '）</div>';
    if (!pays.length) {
      html += '<div class="card empty-state"><span class="emoji">💰</span>暂无收款记录</div>';
    } else {
      html += '<div class="table-wrap"><table class="table">' +
        '<thead><tr><th>日期</th><th>房源</th><th>类型</th><th>金额</th><th>方式</th><th>备注</th></tr></thead><tbody>';
      var nameOf = {};
      state.houses.forEach(function (h) { nameOf[h.id] = h.name; });
      pays.forEach(function (p) {
        html += '<tr>' +
          '<td>' + utils.fmtDateCN(p.payDate) + '</td>' +
          '<td>' + esc(nameOf[p.houseId] || '') + '</td>' +
          '<td><span class="badge type-' + p.type + '">' + TYPE_CN[p.type] + '</span></td>' +
          '<td>¥' + utils.fmtMoney(p.amount) + '</td>' +
          '<td>' + esc(METHOD_CN[p.method] || '') + '</td>' +
          '<td>' + esc(p.remark || '') + '</td>' +
          '</tr>';
      });
      html += '</tbody></table></div>';
    }
    $('#content').innerHTML = html;
  }

  /* ================= 上传 / 下载 ================= */
  function doUpload(file) {
    if (!isConfigured()) { toast('请先在 js/config.js 里配置 jsonbin（见 README）', 'error'); return; }
    if (typeof XLSX === 'undefined') { toast('Excel 组件未加载，请刷新后重试', 'error'); return; }
    parseXlsx(file).then(function (res) {
      if (!res.houses.length) { toast('Excel 里没有「房源」数据', 'error'); return; }
      var payload = {
        schema: 1,
        updatedAt: new Date().toISOString(),
        houses: res.houses,
        payments: res.payments
      };
      var size = JSON.stringify(payload).length;
      if (size > 90000) {
        toast('数据较大（约 ' + Math.round(size / 1024) + 'KB），已接近 jsonbin 免费上限，请精简后再传', 'error');
        return;
      }
      saveToCloud(payload).then(function () {
        toast(res.skipped ? '上传成功（' + res.skipped + ' 笔收款因房源不匹配被跳过）' : '上传成功');
        state = payload;
        render();
      }).catch(function (e) {
        toast('上传失败：' + e.message, 'error');
      });
    }).catch(function (e) {
      toast('Excel 解析失败：' + e.message, 'error');
    });
  }

  function saveToCloud(payload) {
    if (BIN_ID) return updateBin(payload);
    return createBin(payload).then(function (id) {
      BIN_ID = id;
      promptSaveBinId(id);
    });
  }

  function promptSaveBinId(id) {
    var msg = '已自动创建云端仓库，仓库 ID：<br><b style="word-break:break-all">' + esc(id) + '</b>' +
      '<br><br>请把这段 ID 填到 <b>js/config.js</b> 的 BIN_ID 里，重新上传到 GitHub。' +
      '<br>这样所有设备（A、B）都指向同一个仓库。';
    var d = document.createElement('div');
    d.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:100;display:flex;align-items:center;justify-content:center;padding:24px';
    d.innerHTML = '<div style="background:#fff;border-radius:14px;max-width:420px;width:100%;padding:20px;text-align:center">' +
      '<div style="font-size:16px;font-weight:700;margin-bottom:10px">请保存仓库 ID</div>' +
      '<div style="font-size:13px;color:#595959;line-height:1.7">' + msg + '</div>' +
      '<button id="save-id-ok" style="margin-top:16px;width:100%;height:42px;border:none;border-radius:10px;background:#1677ff;color:#fff;font-size:15px">我知道了</button></div>';
    document.body.appendChild(d);
    d.querySelector('#save-id-ok').onclick = function () { d.remove(); };
  }

  function doDownload() {
    if (!state) { toast('云端还没有数据，请先上传', 'error'); return; }
    if (typeof XLSX === 'undefined') { toast('Excel 组件未加载，请刷新后重试', 'error'); return; }
    var wb = buildXlsx(state);
    var filename = '租房管家数据_' + utils.todayStr() + '.xlsx';
    try {
      var array = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      var blob = new Blob([array], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      if (navigator.share && typeof File !== 'undefined') {
        var file = new File([blob], filename, { type: blob.type });
        if (navigator.canShare && navigator.canShare({ files: [file] })) {
          navigator.share({ files: [file], title: '租房管家数据' }).then(function () {}, function () {});
          return;
        }
      }
      XLSX.writeFile(wb, filename);
      toast('已下载，用 Excel 打开编辑');
    } catch (e) {
      toast('下载失败：' + e.message, 'error');
    }
  }

  function loadLatest() {
    if (!isConfigured()) { render(); return; }
    fetchBin().then(function (record) {
      state = record;
      render();
    }).catch(function () {
      state = null;
      render();
    });
  }

  /* ================= toast ================= */
  var toastTimer = null;
  function toast(msg, type) {
    var t = $('#toast');
    t.textContent = msg;
    t.className = 'toast' + (type === 'error' ? ' error' : '');
    t.style.display = 'block';
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.style.display = 'none'; }, 2200);
  }

  /* ================= 初始化 ================= */
  function init() {
    els.lastUpdate = $('#last-update');
    $('#content').innerHTML = '<div class="empty-state"><span class="emoji">⏳</span>正在读取云端…</div>';

    document.querySelector('.tabs').addEventListener('click', function (e) {
      var b = e.target.closest('.tab');
      if (!b) return;
      tab = b.dataset.tab;
      document.querySelectorAll('.tab').forEach(function (x) { x.classList.toggle('active', x === b); });
      render();
    });

    var fileInput = $('#excel-file');
    document.querySelector('[data-act="upload"]').onclick = function () { fileInput.click(); };
    fileInput.onchange = function () {
      if (fileInput.files && fileInput.files[0]) doUpload(fileInput.files[0]);
      fileInput.value = '';
    };
    document.querySelector('[data-act="download"]').onclick = doDownload;

    loadLatest();
  }

  document.addEventListener('DOMContentLoaded', init);
})();
