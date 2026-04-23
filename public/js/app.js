document.addEventListener('DOMContentLoaded', () => {
  // === State ===
  const tabs = ['pending', 'started', 'finished'];
  const data = { pending: [], started: [], finished: [] };
  let currentTab = 'pending';
  let activeJob = null;

  // Selection & Filter state per tab
  const selectedKeys = { pending: new Set(), started: new Set(), finished: new Set() };
  const searchTerms = { pending: '', started: '', finished: '' };
  const searchTags = { pending: [], started: [], finished: [] };
  const activeBranch = { pending: null, started: null, finished: null };
  let finishedMonth = '';
  let finishedPage = 1;
  let finishedHasMore = false;
  let isFetchingFinished = false;
  let totalFinishedCount = 0;
  let sortOrderPending = 'desc';
  let filterTodayPending = false;

  // === Monitor / Data-Load Log ===
  const MONITOR_PIN = '638796796';
  const MONITOR_LS_KEY = 'docdelivery_loadlog';
  const MONITOR_RETENTION_MS = 7 * 24 * 60 * 60 * 1000; // 7 วัน

  function loadLogFromStorage() {
    try {
      const raw = localStorage.getItem(MONITOR_LS_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      const cutoff = Date.now() - MONITOR_RETENTION_MS;
      return parsed.filter(e => e.ts >= cutoff);
    } catch { return []; }
  }

  function saveLogToStorage(log) {
    try {
      localStorage.setItem(MONITOR_LS_KEY, JSON.stringify(log));
    } catch { /* storage full — skip */ }
  }

  const loadLog = loadLogFromStorage();

  function snapshotRecords(rows, deptField = 'แผนก ปลายทาง') {
    return rows.map(r => ({
      k: r.Key || '',
      s: r['ชื่อผู้ส่ง'] || '',
      d: r[deptField] || r['แผนก ต้นทาง'] || r['แผนก ปลายทาง'] || '',
      t: r['เวลาทำรายการ'] || r['Dropoff'] || ''
    }));
  }

  function logLoad(name, count, success, error = null, records = [], groupByDept = false) {
    loadLog.unshift({ name, count, ts: Date.now(), success, error, records, groupByDept });
    const cutoff = Date.now() - MONITOR_RETENTION_MS;
    while (loadLog.length && loadLog[loadLog.length - 1].ts < cutoff) loadLog.pop();
    saveLogToStorage(loadLog);
  }

  // === DOM Elements ===
  const navBtns = document.querySelectorAll('.nav-btn');
  const panels = document.querySelectorAll('.tab-panel');
  const refreshBtn = document.getElementById('refreshBtn');
  const syncDot = document.getElementById('syncDot');
  const syncLabel = document.getElementById('syncLabel');

  // Modal
  const modalBtn = document.getElementById('confirmBtn');
  const modalLabel = document.getElementById('confirmBtnLabel');
  const modalOverlay = document.getElementById('modalOverlay');
  const modalCloseBtn = document.getElementById('modalCloseBtn');

  // Batch Bar
  const batchBar = document.getElementById('batchBar');
  const batchCount = document.getElementById('batchCount');
  const batchCancelBtn = document.getElementById('batchCancelBtn');
  const batchConfirmBtn = document.getElementById('batchConfirmBtn');
  const batchConfirmLabel = document.getElementById('batchConfirmLabel');

  // Batch Signature Modal
  const batchSigOverlay = document.getElementById('batchSigOverlay');
  const batchSigCloseBtn = document.getElementById('batchSigCloseBtn');
  const batchSigClearBtn = document.getElementById('batchSigClearBtn');
  const batchSigConfirmBtn = document.getElementById('batchSigConfirmBtn');
  const batchSigConfirmLabel = document.getElementById('batchSigConfirmLabel');
  const batchSigInfo = document.getElementById('batchSigInfo');

  // Theme
  const themeToggle = document.getElementById('themeToggle');

  // Signature (main modal)
  let canvas, ctx;
  let isDrawing = false;
  let hasSignature = false;
  let lastX = 0;
  let lastY = 0;

  // Signature (batch modal)
  let bCanvas, bCtx;
  let bIsDrawing = false;
  let bHasSignature = false;
  let bLastX = 0;
  let bLastY = 0;

  // Signature Viewer
  let sigViewer;

  // === Presence (multi-user lock) ===
  let userId = localStorage.getItem('docDelivery_userId');
  if (!userId) {
    userId = 'u_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    localStorage.setItem('docDelivery_userId', userId);
  }
  const displayName = localStorage.getItem('docDelivery_displayName') || ('ผู้ใช้-' + userId.slice(-4));
  let othersPresence = {}; // { key: { userId, name } }
  let viewingKey = null;   // key being viewed in modal (also claimed)

  function getMyClaims() {
    const claims = new Set();
    selectedKeys.pending.forEach(k => claims.add(k));
    selectedKeys.started.forEach(k => claims.add(k));
    if (viewingKey) claims.add(viewingKey);
    return Array.from(claims);
  }

  async function sendHeartbeat() {
    try {
      const res = await fetch('/api/presence/heartbeat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ userId, name: displayName, claims: getMyClaims() })
      });
      const json = await res.json();
      if (json.success) {
        const prev = JSON.stringify(othersPresence);
        othersPresence = json.others || {};
        if (prev !== JSON.stringify(othersPresence)) {
          updatePresenceIcons();
        }
      }
    } catch (_) { /* ignore */ }
  }

  function showConflictToast(msg) {
    let t = document.getElementById('conflict-toast');
    if (!t) {
      t = document.createElement('div');
      t.id = 'conflict-toast';
      t.className = 'conflict-toast';
      document.body.appendChild(t);
    }
    t.innerHTML = '<i class="bx bxs-lock-alt"></i> ' + msg;
    t.classList.add('show');
    clearTimeout(t._hideTimer);
    t._hideTimer = setTimeout(() => t.classList.remove('show'), 3500);
  }

  function updatePresenceIcons() {
    document.querySelectorAll('.job-card[data-key]').forEach(card => {
      const k = card.getAttribute('data-key');
      const left = card.querySelector('.card-left');
      if (!left) return;
      const existing = left.querySelector('.presence-lock');
      const owner = othersPresence[k];
      if (owner) {
        const html = '<i class="bx bxs-user-circle"></i><span class="presence-name">' + owner.name + '</span>';
        if (!existing) {
          const icon = document.createElement('span');
          icon.className = 'presence-lock';
          icon.title = 'กำลังถูกใช้งานโดย ' + owner.name;
          icon.innerHTML = html;
          left.appendChild(icon);
        } else {
          existing.title = 'กำลังถูกใช้งานโดย ' + owner.name;
          existing.innerHTML = html;
        }
        card.classList.add('is-locked');
      } else {
        if (existing) existing.remove();
        card.classList.remove('is-locked');
      }
    });
  }

  // Send on unload to release immediately
  window.addEventListener('beforeunload', () => {
    try {
      const blob = new Blob([JSON.stringify({ userId, name: displayName, claims: [] })], { type: 'application/json' });
      navigator.sendBeacon('/api/presence/heartbeat', blob);
    } catch (_) {}
  });

  // === Init ===
  initTheme();
  initTabs();
  initSearch();
  initDateFilter();
  initModalHandlers();
  initBatchHandlers();
  initBatchSigModal();
  initSignaturePad();
  createSignatureViewer();
  fetchData();
  fetchFinishedJobs(true);
  initRealtimeStream();
  sendHeartbeat();
  setInterval(sendHeartbeat, 8000);

  refreshBtn.addEventListener('click', () => {
    const icon = refreshBtn.querySelector('svg') || refreshBtn.querySelector('i');
    if (icon) {
      icon.style.transform = 'rotate(180deg)';
      icon.style.transition = 'transform 0.3s';
    }
    Promise.all([fetchData(true), fetchFinishedJobs(true)]).finally(() => {
      setTimeout(() => {
        if (icon) {
          icon.style.transform = 'none';
          icon.style.transition = 'none';
        }
      }, 300);
    });
  });

  // ========================
  // THEME
  // ========================
  function initTheme() {
    const saved = localStorage.getItem('docdelivery-theme') || 'dark';
    document.documentElement.setAttribute('data-theme', saved);
    updateThemeIcons(saved);

    themeToggle.addEventListener('click', () => {
      const current = document.documentElement.getAttribute('data-theme');
      const next = current === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      localStorage.setItem('docdelivery-theme', next);
      updateThemeIcons(next);
      const meta = document.querySelector('meta[name="theme-color"]');
      if (meta) meta.setAttribute('content', next === 'dark' ? '#050b18' : '#f0f2f5');
    });
  }

  function updateThemeIcons(theme) {
    const darkIcon = themeToggle.querySelector('.theme-icon-dark');
    const lightIcon = themeToggle.querySelector('.theme-icon-light');
    if (theme === 'dark') {
      darkIcon.classList.remove('hidden');
      lightIcon.classList.add('hidden');
    } else {
      darkIcon.classList.add('hidden');
      lightIcon.classList.remove('hidden');
    }
  }

  // ========================
  // API & DATA
  // ========================
  async function fetchData(force = false) {
    setSyncStatus('syncing', 'กำลังซิงค์...');
    try {
      const url = `/api/jobs${force ? '?refresh=1' : ''}`;
      const res = await fetch(url);
      const json = await res.json();
      if (!json.success) throw new Error(json.error);
      
      processRawData(json.data);
      renderBranchChips('pending'); renderTab('pending');
      renderBranchChips('started'); renderTab('started');
      setSyncStatus('online', 'ออนไลน์');
      logLoad('งานรอรับ (Pending)', data.pending.length, true, null, snapshotRecords(data.pending, 'แผนก ต้นทาง'), true);
      logLoad('งานกำลังส่ง (Started)', data.started.length, true, null, snapshotRecords(data.started, 'แผนก ปลายทาง'), false);
    } catch (err) {
      console.error(err);
      setSyncStatus('error', 'ออฟไลน์');
      logLoad('งานรอรับ / กำลังส่ง', 0, false, err.message);
      if (force) showToast('ไม่สามารถดึงข้อมูลได้', 'error');
    }
  }

  async function fetchFinishedJobs(reset = false) {
    if (isFetchingFinished) return;
    isFetchingFinished = true;
    if (reset) {
      finishedPage = 1;
      data.finished = [];
      const listEl = document.getElementById('list-finished');
      if (listEl) listEl.innerHTML = '';
      totalFinishedCount = 0;
    }
    
    document.getElementById('loading-finished').classList.remove('hidden');
    document.getElementById('empty-finished').classList.add('hidden');
    document.getElementById('load-more-container').classList.add('hidden');

    try {
      let url = `/api/jobs/finished?page=${finishedPage}&limit=20`;
      if (finishedMonth) url += `&month=${finishedMonth}`;
      
      const res = await fetch(url);
      const json = await res.json();
      if (!json.success) throw new Error(json.error);

      if (reset) {
        data.finished = json.data;
      } else {
        data.finished = [...data.finished, ...json.data];
      }
      
      finishedHasMore = json.hasMore;
      totalFinishedCount = json.total || data.finished.length;

      renderBranchChips('finished');
      renderTab('finished');
      logLoad(`ประวัติส่งสำเร็จ (หน้า ${finishedPage})`, json.data.length, true, null, snapshotRecords(json.data));

    } catch (err) {
      console.error('fetchFinishedJobs error:', err);
      logLoad('ประวัติส่งสำเร็จ', 0, false, err.message);
      showToast('ไม่สามารถดึงข้อมูลประวัติได้', 'error');
    } finally {
      isFetchingFinished = false;
      document.getElementById('loading-finished').classList.add('hidden');
      if (finishedHasMore) {
        document.getElementById('load-more-container').classList.remove('hidden');
      }
    }
  }

  function setSyncStatus(state, label) {
    syncDot.className = `sync-dot ${state}`;
    syncLabel.textContent = label;
  }

  // ========================
  // REAL-TIME STREAM (SSE)
  // ========================
  // var (hoisted) so initRealtimeStream() called earlier in DOMContentLoaded
  // can reference these without hitting the TDZ of let-declarations below.
  var sseSource = null;
  var sseBackoff = 1000;
  var sseFallbackTimer = null;

  function initRealtimeStream() {
    if (typeof EventSource === 'undefined') {
      // Fallback to polling for old browsers
      setInterval(() => fetchData(false), 15000);
      return;
    }
    connectSSE();
  }

  function connectSSE() {
    try {
      if (sseSource) sseSource.close();
      sseSource = new EventSource('/api/events');

      sseSource.addEventListener('hello', () => {
        sseBackoff = 1000;
        if (sseFallbackTimer) { clearInterval(sseFallbackTimer); sseFallbackTimer = null; }
      });

      sseSource.addEventListener('jobs', () => {
        fetchData(false);
        if (currentTab === 'finished') fetchFinishedJobs(true);
      });

      sseSource.onerror = () => {
        if (sseSource) sseSource.close();
        sseSource = null;
        // Start short-interval polling fallback until reconnect
        if (!sseFallbackTimer) {
          sseFallbackTimer = setInterval(() => fetchData(false), 10000);
        }
        sseBackoff = Math.min(sseBackoff * 2, 15000);
        setTimeout(connectSSE, sseBackoff);
      };
    } catch (e) {
      console.warn('SSE init failed, using polling', e);
      setInterval(() => fetchData(false), 15000);
    }
  }

  function processRawData(rows) {
    data.pending = [];
    data.started = [];
    
    rows.forEach(row => {
      if (!row['Key']) return;
      const statusRaw = String(row['Status'] || '').trim().toLowerCase();
      if (statusRaw === 'started' || statusRaw === '2') {
        data.started.push(row);
      } else {
        data.pending.push(row);
      }
    });

    data.started.sort((a, b) => (b['เวลาทำรายการ'] || '').localeCompare(a['เวลาทำรายการ'] || ''));
    data.pending.sort((a, b) => (b['เวลาทำรายการ'] || '').localeCompare(a['เวลาทำรายการ'] || ''));
  }

  async function updateJobStatus(jobKey, newStatus, extraData = {}) {
    const payload = { key: jobKey, status: newStatus, ...extraData };
    try {
      modalBtn.disabled = true;
      modalLabel.textContent = 'กำลังบันทึก...';
      
      const res = await fetch('/api/jobs/update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ...payload, _userId: userId })
      });
      const json = await res.json();
      if (res.status === 409 || json.conflict) {
        showConflictToast(json.error || 'รายการนี้กำลังถูกใช้งานโดยผู้ใช้งานท่านอื่น');
        throw new Error(json.error || 'Conflict');
      }
      if (!json.success && json.status !== 'success') throw new Error(json.error || 'Update failed');
      
      closeModal();
      showToast('ทำรายการสำเร็จ', 'success');
      fetchData(true);
    } catch (err) {
      console.error(err);
      showToast('เกิดข้อผิดพลาด: ' + err.message, 'error');
      modalBtn.disabled = false;
      modalLabel.textContent = 'ลองใหม่';
    }
  }

  async function batchUpdateStatus(keys, newStatus, extraData = {}) {
    let successCount = 0;
    let failCount = 0;
    
    // Send all requests in parallel for speed
    const promises = keys.map(key => {
      const payload = { key, status: newStatus, ...extraData };
      return fetch('/api/jobs/update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ...payload, _userId: userId })
      }).then(async r => {
        const json = await r.json();
        if (r.status === 409 || json.conflict) {
          showConflictToast(json.error || ('รายการ ' + key + ' กำลังถูกใช้งานโดยผู้ใช้งานท่านอื่น'));
        }
        return json;
      }).then(json => {
        if (json.success || json.status === 'success') successCount++;
        else failCount++;
      }).catch(() => { failCount++; });
    });

    await Promise.all(promises);
    
    selectedKeys[currentTab].clear();
    updateBatchBar();
    
    if (successCount > 0) showToast(`อัปเดตสำเร็จ ${successCount} รายการ`, 'success');
    if (failCount > 0) showToast(`ล้มเหลว ${failCount} รายการ`, 'error');
    
    fetchData(true);
  }

  // ========================
  // SEARCH & FILTER
  // ========================
  function initSearch() {
    tabs.forEach(tab => {
      const input = document.getElementById(`search-${tab}`);
      if (!input) return;

      const wrap = input.parentElement;
      if (wrap) {
        wrap.classList.add('has-tags');
        // Tags container (chips) — placed before input
        if (!wrap.querySelector('.search-tags')) {
          const tagsEl = document.createElement('div');
          tagsEl.className = 'search-tags';
          tagsEl.id = `tags-${tab}`;
          wrap.insertBefore(tagsEl, input);
        }
        // Suggestions dropdown
        if (!wrap.querySelector('.search-suggestions')) {
          const sug = document.createElement('div');
          sug.className = 'search-suggestions hidden';
          sug.id = `suggestions-${tab}`;
          wrap.appendChild(sug);
        }
      }

      input.addEventListener('input', () => {
        searchTerms[tab] = input.value.trim().toLowerCase();
        renderSuggestions(tab);
        renderTab(tab);
      });
      input.addEventListener('focus', () => renderSuggestions(tab));
      input.addEventListener('blur', () => {
        setTimeout(() => hideSuggestions(tab), 150);
      });
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
          hideSuggestions(tab);
          input.blur();
        }
        // Backspace on empty input removes last tag
        if (e.key === 'Backspace' && input.value === '' && searchTags[tab].length > 0) {
          searchTags[tab].pop();
          renderSearchTags(tab);
          renderTab(tab);
        }
      });
    });
  }

  function renderSearchTags(tab) {
    const el = document.getElementById(`tags-${tab}`);
    if (!el) return;
    el.innerHTML = searchTags[tab].map((t, i) => (
      '<span class="search-tag" data-index="' + i + '">' +
      '<span class="search-tag-text">' + escapeHtml(t) + '</span>' +
      '<button type="button" class="search-tag-remove" aria-label="ลบ" data-index="' + i + '">&times;</button>' +
      '</span>'
    )).join('');
    el.querySelectorAll('.search-tag-remove').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const idx = parseInt(btn.getAttribute('data-index'), 10);
        searchTags[tab].splice(idx, 1);
        renderSearchTags(tab);
        renderTab(tab);
      });
    });
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }

  function highlightTerm(text, term) {
    const safe = escapeHtml(text);
    if (!term) return safe;
    const idx = safe.toLowerCase().indexOf(term.toLowerCase());
    if (idx === -1) return safe;
    return safe.slice(0, idx) +
      '<mark>' + safe.slice(idx, idx + term.length) + '</mark>' +
      safe.slice(idx + term.length);
  }

  function hideSuggestions(tab) {
    const box = document.getElementById(`suggestions-${tab}`);
    if (box) box.classList.add('hidden');
  }

  function renderSuggestions(tab) {
    const box = document.getElementById(`suggestions-${tab}`);
    const input = document.getElementById(`search-${tab}`);
    if (!box || !input) return;
    const term = (input.value || '').trim().toLowerCase();
    if (!term) {
      box.classList.add('hidden');
      box.innerHTML = '';
      return;
    }

    // Collect unique field values that match the term, grouped by field type
    const fieldLabels = {
      'ชื่อผู้ส่ง': 'ผู้ส่ง',
      'ชื่อผู้รับ': 'ผู้รับ',
      'แผนก ต้นทาง': 'แผนกต้นทาง',
      'แผนก ปลายทาง': 'แผนกปลายทาง',
      'ส่งจากสาขา': 'สาขา',
      'รายละเอียด': 'รายละเอียด',
      'Key': 'รหัส',
      'วันที่ส่งเอกสาร/พัสดุ': 'วันที่'
    };
    const fieldKeys = Object.keys(fieldLabels);

    const seen = new Set(); // dedupe by "field|value"
    const matches = [];
    let startsWithCount = 0;

    // Pass 1: starts-with matches (higher priority)
    for (const item of data[tab]) {
      for (const fk of fieldKeys) {
        const val = item[fk];
        if (!val) continue;
        const v = String(val).trim();
        if (!v) continue;
        if (!v.toLowerCase().startsWith(term)) continue;
        const id = fk + '|' + v.toLowerCase();
        if (seen.has(id)) continue;
        seen.add(id);
        matches.push({ field: fk, label: fieldLabels[fk], value: v });
        startsWithCount++;
        if (matches.length >= 10) break;
      }
      if (matches.length >= 10) break;
    }

    // Pass 2: contains matches (fill remaining slots)
    if (matches.length < 10) {
      for (const item of data[tab]) {
        for (const fk of fieldKeys) {
          const val = item[fk];
          if (!val) continue;
          const v = String(val).trim();
          if (!v) continue;
          const lower = v.toLowerCase();
          if (lower.startsWith(term)) continue; // already in pass 1
          if (!lower.includes(term)) continue;
          const id = fk + '|' + lower;
          if (seen.has(id)) continue;
          seen.add(id);
          matches.push({ field: fk, label: fieldLabels[fk], value: v });
          if (matches.length >= 10) break;
        }
        if (matches.length >= 10) break;
      }
    }

    if (!matches.length) {
      box.classList.add('hidden');
      box.innerHTML = '';
      return;
    }

    const headerHtml =
      '<div class="suggestions-header">' +
      '<span><i class="bx bx-list-ul"></i> ตัวเลือกที่พบ</span>' +
      '<span>' + matches.length + ' รายการ</span>' +
      '</div>';

    box.innerHTML = headerHtml + matches.map((m, i) => {
      return (
        '<button type="button" class="suggestion-item" data-value="' + escapeHtml(m.value) + '" data-index="' + i + '">' +
        '<i class="bx bx-search suggestion-icon"></i>' +
        '<div class="suggestion-text">' +
        '<div class="suggestion-main">' + highlightTerm(m.value, term) + '</div>' +
        '<div class="suggestion-sub">' + escapeHtml(m.label) + '</div>' +
        '</div>' +
        '<span class="suggestion-key">' + escapeHtml(m.label) + '</span>' +
        '</button>'
      );
    }).join('');
    box.classList.remove('hidden');

    box.querySelectorAll('.suggestion-item').forEach(btn => {
      btn.addEventListener('mousedown', (e) => e.preventDefault()); // keep focus
      btn.addEventListener('click', () => {
        const value = btn.getAttribute('data-value');
        // Add as tag chip (avoid duplicates)
        if (!searchTags[tab].includes(value)) {
          searchTags[tab].push(value);
        }
        input.value = '';
        searchTerms[tab] = '';
        renderSearchTags(tab);
        hideSuggestions(tab);
        renderTab(tab);
        input.focus();
      });
    });
  }

  function initDateFilter() {
    const monthInput = document.getElementById('month-finished');
    const monthClearBtn = document.getElementById('month-clear-finished');
    const loadMoreBtn = document.getElementById('load-more-finished');
    if (!monthInput) return;

    // Default to current month
    const now = new Date();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    finishedMonth = `${now.getFullYear()}-${mm}`;
    monthInput.value = finishedMonth;
    monthClearBtn.classList.remove('hidden');

    monthInput.addEventListener('change', () => {
      finishedMonth = monthInput.value;
      if (finishedMonth) {
        monthClearBtn.classList.remove('hidden');
      } else {
        monthClearBtn.classList.add('hidden');
      }
      fetchFinishedJobs(true);
    });

    monthClearBtn.addEventListener('click', () => {
      monthInput.value = '';
      finishedMonth = '';
      monthClearBtn.classList.add('hidden');
      fetchFinishedJobs(true);
    });

    if (loadMoreBtn) {
      loadMoreBtn.addEventListener('click', () => {
        if (!isFetchingFinished && finishedHasMore) {
          finishedPage++;
          fetchFinishedJobs(false);
        }
      });
    }

    const sortSelect = document.getElementById('sort-date-pending');
    if (sortSelect) {
      sortSelect.addEventListener('change', () => {
        sortOrderPending = sortSelect.value;
        renderTab('pending');
      });
    }

    const todayCheckbox = document.getElementById('filter-today-pending');
    if (todayCheckbox) {
      todayCheckbox.addEventListener('change', () => {
        filterTodayPending = todayCheckbox.checked;
        renderTab('pending');
      });
    }
  }

  function getFilteredItems(tab) {
    let items = data[tab];
    const term = searchTerms[tab];
    const branch = activeBranch[tab];

    // Branch filter
    if (branch) {
      items = items.filter(item => (item['ส่งจากสาขา'] || '') === branch);
    }

    // Today filter
    if (tab === 'pending' && filterTodayPending) {
      const now = new Date();
      let d = String(now.getDate()).padStart(2, '0');
      let m = String(now.getMonth() + 1).padStart(2, '0');
      let ceYear = now.getFullYear();
      let thYear = ceYear + 543;
      items = items.filter(item => {
        const itemDate = item['วันที่ส่งเอกสาร/พัสดุ'] || '';
        return itemDate === `${d}/${m}/${ceYear}` || itemDate === `${d}/${m}/${thYear}`;
      });
    }

    // Search filter (tags AND live term, all must match)
    const tags = searchTags[tab] || [];
    if (term || tags.length) {
      items = items.filter(item => {
        const searchable = [
          item['Key'],
          item['ชื่อผู้ส่ง'],
          item['ชื่อผู้รับ'],
          item['แผนก ต้นทาง'],
          item['แผนก ปลายทาง'],
          item['รายละเอียด'],
          item['ส่งจากสาขา'],
          item['วันที่ส่งเอกสาร/พัสดุ']
        ].filter(Boolean).join(' ').toLowerCase();
        if (term && !searchable.includes(term)) return false;
        for (const t of tags) {
          if (!searchable.includes(String(t).toLowerCase())) return false;
        }
        return true;
      });
    }

    return items;
  }

  function getBranches(tab) {
    const branches = new Set();
    data[tab].forEach(item => {
      const b = (item['ส่งจากสาขา'] || '').trim();
      if (b) branches.add(b);
    });
    return Array.from(branches).sort();
  }

  function renderBranchChips(tab) {
    const container = document.getElementById(`branch-${tab}`);
    if (!container) return;
    const branches = getBranches(tab);
    
    if (branches.length <= 1) {
      container.innerHTML = '';
      return;
    }

    container.innerHTML = '';
    
    const allChip = document.createElement('span');
    allChip.className = `branch-chip${!activeBranch[tab] ? ' active' : ''}`;
    allChip.textContent = 'ทั้งหมด';
    allChip.addEventListener('click', () => {
      activeBranch[tab] = null;
      renderBranchChips(tab);
      renderTab(tab);
    });
    container.appendChild(allChip);

    branches.forEach(branch => {
      const count = data[tab].filter(i => i['ส่งจากสาขา'] === branch).length;
      const chip = document.createElement('span');
      chip.className = `branch-chip${activeBranch[tab] === branch ? ' active' : ''}`;
      chip.textContent = `${branch} (${count})`;
      chip.addEventListener('click', () => {
        activeBranch[tab] = activeBranch[tab] === branch ? null : branch;
        renderBranchChips(tab);
        renderTab(tab);
      });
      container.appendChild(chip);
    });
  }

  // ========================
  // UI RENDERING
  // ========================
  function renderAllTabs() {
    tabs.forEach(tab => {
      renderBranchChips(tab);
      renderTab(tab);
    });
  }

  function renderTab(tab) {
    const items = getFilteredItems(tab);
    const totalCount = tab === 'finished' ? totalFinishedCount : data[tab].length;
    const filteredCount = items.length;

    const hasFilter = searchTerms[tab] || activeBranch[tab];
    const countText = hasFilter
      ? `${filteredCount}/${totalCount} รายการ`
      : `${totalCount} รายการ`;
    document.getElementById(`count-${tab}`).textContent = countText;

    const badge = document.getElementById(`badge-${tab}`);
    if (totalCount > 0 && tab !== 'finished') {
      badge.textContent = totalCount > 99 ? '99+' : totalCount;
      badge.classList.remove('hidden');
    } else {
      badge.classList.add('hidden');
    }

    if (tab === 'pending' || tab === 'started') {
      renderGroupedList(tab, items);
    } else {
      renderSimpleList(tab, items);
    }
  }

  function renderGroupedList(tab, items) {
    const listEl = document.getElementById(`list-${tab}`);
    document.getElementById(`loading-${tab}`).classList.add('hidden');
    
    if (data[tab].length === 0) {
      listEl.innerHTML = '';
      document.getElementById(`empty-${tab}`).classList.remove('hidden');
      return;
    }
    
    document.getElementById(`empty-${tab}`).classList.add('hidden');
    listEl.innerHTML = '';

    if (items.length === 0) {
      listEl.innerHTML = '<div class="no-results"><i class="no-results-icon bx bx-search"></i><p>ไม่พบรายการที่ค้นหา</p></div>';
      return;
    }

    const groups = {};
    items.forEach(item => {
      const date = item['วันที่ส่งเอกสาร/พัสดุ'] || 'ไม่ระบุวันที่';
      const dept = tab === 'started' ? (item['แผนก ปลายทาง'] || 'ไม่ระบุแผนก') : (item['แผนก ต้นทาง'] || 'ไม่ระบุแผนก');
      if (!groups[date]) groups[date] = {};
      if (!groups[date][dept]) groups[date][dept] = [];
      groups[date][dept].push(item);
    });

    const parseDateToISO = (dateStr) => {
      if (!dateStr || dateStr === 'ไม่ระบุวันที่') return '0000-00-00';
      const m = dateStr.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (m) {
        let yr = parseInt(m[3]);
        if (yr > 2500) yr -= 543;
        return `${yr}-${m[2]}-${m[1]}`;
      }
      return dateStr;
    };

    const dates = Object.keys(groups).sort((a, b) => {
      const dateA = parseDateToISO(a);
      const dateB = parseDateToISO(b);
      if (tab === 'pending') {
        if (sortOrderPending === 'desc') return dateB.localeCompare(dateA);
        return dateA.localeCompare(dateB);
      }
      return dateB.localeCompare(dateA); // Default DESC for started
    });

    dates.forEach(date => {
      const dateHeader = document.createElement('div');
      dateHeader.className = 'group-date-header';
      dateHeader.innerHTML = date === 'ไม่ระบุวันที่' ? date : `<i class="bx bx-calendar-event"></i> ${date}`;
      listEl.appendChild(dateHeader);

      const depts = Object.keys(groups[date]).sort();
      depts.forEach(dept => {
        const deptHeader = document.createElement('div');
        deptHeader.className = 'group-dept-header';
        deptHeader.textContent = `${dept} (${groups[date][dept].length})`;
        listEl.appendChild(deptHeader);
        
        groups[date][dept].forEach(job => {
          listEl.appendChild(createJobCard(job, tab));
        });
      });
    });
  }

  function renderSimpleList(tab, items) {
    const listEl = document.getElementById(`list-${tab}`);
    document.getElementById(`loading-${tab}`).classList.add('hidden');
    
    if (data[tab].length === 0) {
      listEl.innerHTML = '';
      document.getElementById(`empty-${tab}`).classList.remove('hidden');
      return;
    }
    
    document.getElementById(`empty-${tab}`).classList.add('hidden');
    listEl.innerHTML = '';

    if (items.length === 0) {
      listEl.innerHTML = '<div class="no-results"><i class="no-results-icon bx bx-search"></i><p>ไม่พบรายการที่ค้นหา</p></div>';
      return;
    }
    
    items.forEach(job => {
      listEl.appendChild(createJobCard(job, tab));
    });
  }

  function createJobCard(job, type) {
    const card = document.createElement('div');
    card.className = 'job-card';
    const key = job.Key;
    card.setAttribute('data-key', key);
    const isSelected = selectedKeys[type].has(key);
    if (isSelected) card.classList.add('selected');
    const lockedBy = othersPresence[key];
    if (lockedBy) card.classList.add('is-locked');
    
    const sender = job['ชื่อผู้ส่ง'] || '-';
    const receiver = job['ชื่อผู้รับ'] || '-';
    const qty = job['จำนวน'] || '1';
    // รายละเอียดเอกสาร = คอลัมน์ F ของ Google Sheet (index 5, นับจาก A=0)
    const detail = Object.values(job)[5] || '';
    const branch = job['ส่งจากสาขา'] || '';
    const sigUrl = job['Txt_01'] || job['Dropoff Signature'] || '';
    
    let infoHtml = '';
    if (type === 'pending' || type === 'started') {
      infoHtml = `
        <div class="info-row">
          <svg class="info-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>
          <div class="info-text">
            <div class="user-name">จาก: ${sender}</div>
            <div class="dept-name">${job['แผนก ต้นทาง'] || ''}</div>
          </div>
        </div>
        <div class="info-row">
          <svg class="info-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><polyline points="20 6 9 17 4 12"/></svg>
          <div class="info-text">
            <div class="user-name">ถึง: ${receiver}</div>
            <div class="dept-name">${job['แผนก ปลายทาง'] || ''}</div>
          </div>
        </div>
      `;
    } else {
      // Finished: show both sender AND receiver
      infoHtml = `
        <div class="info-row">
          <svg class="info-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>
          <div class="info-text">
            <div class="user-name">จาก: ${sender}</div>
            <div class="dept-name">${job['แผนก ต้นทาง'] || ''}</div>
          </div>
        </div>
        <div class="info-row">
          <svg class="info-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
          <div class="info-text">
            <div class="user-name">ผู้รับ: ${receiver}</div>
            <div class="dept-name">${job['แผนก ปลายทาง'] || ''}</div>
          </div>
        </div>
      `;
    }

    const statusPill = type === 'pending' ? '<span class="status-pill status-pending">รอรับ</span>' :
      type === 'started' ? '<span class="status-pill status-started">นำส่ง</span>' :
      '<span class="status-pill status-finished">รับแล้ว</span>';

    card.innerHTML = `
      ${type !== 'finished' ? '<button class="card-delete-bg" type="button" aria-label="ลบรายการ"><i class="bx bx-trash"></i><span>ลบ</span></button>' : ''}
      <div class="card-content">
      <div class="card-top">
        <div class="card-left">
          ${type !== 'finished' ? `<input type="checkbox" class="card-checkbox" data-key="${key}" ${isSelected ? 'checked' : ''} />` : ''}
          <span class="card-key">${key}</span>
          ${lockedBy ? `<span class="presence-lock" title="กำลังถูกใช้งานโดย ${lockedBy.name}"><i class="bx bxs-user-circle"></i><span class="presence-name">${lockedBy.name}</span></span>` : ''}
        </div>
        <span class="card-time">${type === 'finished' ? (job['Dropoff'] || job['เวลาทำรายการ'] || '') : (job['เวลาทำรายการ'] || '')}</span>
      </div>
      <div class="card-body">
        ${infoHtml}
      </div>
      <div class="card-footer">
        <div class="item-count">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="7" width="20" height="14" rx="2" ry="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg>
          ${detail ? `${detail} · ` : ''}จำนวน: <strong>${qty}</strong>
        </div>
        ${branch ? `<span class="branch-tag">${branch}</span>` : ''}
        ${statusPill}
      </div>
      ${type === 'finished' && isValidSignatureSrc(sigUrl) ? `
        <div class="sig-preview-wrap" data-receiver="${receiver}">
          <img class="sig-preview-img" src="${sigUrl}" alt="ลายเซ็น" loading="lazy" onerror="this.parentElement.style.display='none'" />
          <span class="sig-preview-text"><i class="bx bx-edit-alt"></i> ลายเซ็นผู้รับ — กดเพื่อดู</span>
        </div>
      ` : ''}
      </div>
    `;

    // Swipe-to-delete (mobile/tablet only) for non-finished cards
    if (type !== 'finished') {
      enableSwipeDelete(card, key);
    }

    // Signature preview click (for finished cards)
    const sigPreview = card.querySelector('.sig-preview-wrap');
    if (sigPreview) {
      sigPreview.addEventListener('click', (e) => {
        e.stopPropagation();
        const img = sigPreview.querySelector('img');
        const src = img ? img.src : '';
        showSignatureViewer(src, sigPreview.dataset.receiver);
      });
    }

    // Selection only for pending & started
    if (type !== 'finished') {
      const checkbox = card.querySelector('.card-checkbox');
      card.addEventListener('click', (e) => {
        if (e.target.closest('.sig-preview-wrap')) return;
        // Block selection if another user is currently working on this item
        const owner = othersPresence[key];
        if (owner && !checkbox.checked) {
          e.preventDefault();
          if (e.target.classList.contains('card-checkbox')) checkbox.checked = false;
          showConflictToast('รายการ ' + key + ' กำลังถูกดำเนินการโดย ' + owner.name);
          return;
        }
        if (!e.target.classList.contains('card-checkbox')) {
          checkbox.checked = !checkbox.checked;
        }
        if (checkbox.checked) {
          selectedKeys[type].add(key);
          card.classList.add('selected');
        } else {
          selectedKeys[type].delete(key);
          card.classList.remove('selected');
        }
        updateBatchBar();
        sendHeartbeat();
      });
    }

    return card;
  }

  // ========================
  // SIGNATURE VIEWER (full screen)
  // ========================
  function createSignatureViewer() {
    sigViewer = document.createElement('div');
    sigViewer.className = 'sig-viewer-overlay';
    sigViewer.innerHTML = `
      <img class="sig-viewer-img" id="sigViewerImg" src="" alt="ลายเซ็น" />
      <div class="sig-viewer-label" id="sigViewerLabel"></div>
    `;
    document.body.appendChild(sigViewer);
    sigViewer.addEventListener('click', () => sigViewer.classList.remove('active'));
  }

  function showSignatureViewer(url, receiverName) {
    const img = document.getElementById('sigViewerImg');
    const label = document.getElementById('sigViewerLabel');
    img.src = url;
    label.textContent = `ลายเซ็นผู้รับ: ${receiverName}`;
    sigViewer.classList.add('active');
  }

  // ========================
  // BATCH ACTIONS
  // ========================
  function initBatchHandlers() {
    batchCancelBtn.addEventListener('click', () => {
      selectedKeys[currentTab].clear();
      updateBatchBar();
      renderTab(currentTab);
    });

    batchConfirmBtn.addEventListener('click', () => {
      const keys = Array.from(selectedKeys[currentTab]);
      if (keys.length === 0) return;

      if (currentTab === 'pending') {
        batchConfirmBtn.disabled = true;
        batchConfirmLabel.textContent = 'กำลังบันทึก...';
        batchUpdateStatus(keys, 'Started').finally(() => {
          batchConfirmBtn.disabled = false;
          batchConfirmLabel.textContent = 'รับเข้าระบบ';
        });
      } else if (currentTab === 'started') {
        // Open batch signature modal for signing
        openBatchSigModal(keys);
      }
    });
  }

  function updateBatchBar() {
    const count = selectedKeys[currentTab].size;
    if (count > 0 && currentTab !== 'finished') {
      batchBar.classList.remove('hidden');
      batchCount.textContent = count;

      // Sum QTY (column E = index 4) of selected items
      let totalQty = 0;
      data[currentTab].forEach(item => {
        if (selectedKeys[currentTab].has(item.Key)) {
          const raw = item['จำนวน'] != null && item['จำนวน'] !== ''
            ? item['จำนวน']
            : Object.values(item)[4];
          const n = parseInt(String(raw).replace(/[^\d-]/g, ''), 10);
          if (!isNaN(n)) totalQty += n;
        }
      });
      const qtyEl = document.getElementById('batchQty');
      const qtyVal = document.getElementById('batchQtyValue');
      if (qtyEl && qtyVal) {
        qtyVal.textContent = totalQty;
        qtyEl.classList.toggle('hidden', totalQty <= 0);
      }
      
      if (currentTab === 'pending') {
        batchConfirmLabel.textContent = 'รับเข้าระบบ';
        batchConfirmBtn.className = 'batch-btn batch-confirm';
      } else if (currentTab === 'started') {
        batchConfirmLabel.textContent = 'ส่งมอบ';
        batchConfirmBtn.className = 'batch-btn batch-confirm started-action';
      }
    } else {
      batchBar.classList.add('hidden');
    }
  }

  // ========================
  // BATCH SIGNATURE MODAL
  // ========================
  let pendingBatchKeys = [];

  function initBatchSigModal() {
    batchSigCloseBtn.addEventListener('click', closeBatchSigModal);
    batchSigOverlay.addEventListener('click', e => {
      if (e.target === batchSigOverlay) closeBatchSigModal();
    });

    batchSigClearBtn.addEventListener('click', clearBatchSig);

    batchSigConfirmBtn.addEventListener('click', async () => {
      if (!bHasSignature) {
        showToast('กรุณาเซ็นชื่อรับเอกสาร', 'error');
        return;
      }
      
      batchSigConfirmBtn.disabled = true;
      batchSigConfirmLabel.textContent = 'กำลังบันทึก...';
      
      const sigBase64 = compressSignature(bCanvas);
      const now = new Date().toLocaleString('th-TH');
      
      await batchUpdateStatus(pendingBatchKeys, 'Finished', {
        signature: sigBase64,
        dropoff: now
      });
      
      closeBatchSigModal();
      batchSigConfirmBtn.disabled = false;
      batchSigConfirmLabel.textContent = 'ส่งมอบ';
    });

    // Init batch sig canvas
    bCanvas = document.getElementById('batchSigCanvas');
    if (bCanvas) {
      bCtx = bCanvas.getContext('2d', { willReadFrequently: true });

      bCanvas.addEventListener('mousedown', bStartDrawing);
      bCanvas.addEventListener('mousemove', bDraw);
      bCanvas.addEventListener('mouseup', bStopDrawing);
      bCanvas.addEventListener('mouseout', bStopDrawing);

      bCanvas.addEventListener('touchstart', bHandleTouch, { passive: false });
      bCanvas.addEventListener('touchmove', bHandleTouch, { passive: false });
      bCanvas.addEventListener('touchend', bStopDrawing);
    }
  }

  function openBatchSigModal(keys) {
    pendingBatchKeys = keys;
    batchSigInfo.textContent = `กำลังยืนยัน ${keys.length} รายการ`;
    clearBatchSig();
    batchSigOverlay.classList.add('active');
    document.body.style.overflow = 'hidden';
    
    // Resize canvas after modal is visible
    setTimeout(() => resizeBatchCanvas(), 100);
  }

  function closeBatchSigModal() {
    batchSigOverlay.classList.remove('active');
    document.body.style.overflow = 'auto';
    pendingBatchKeys = [];
  }

  function clearBatchSig() {
    if (!bCanvas || !bCtx) return;
    resizeBatchCanvas();
    bHasSignature = false;
    const ph = document.getElementById('batchSigPlaceholder');
    if (ph) ph.style.opacity = '1';
  }

  function resizeBatchCanvas() {
    if (!bCanvas) return;
    const wrap = bCanvas.parentElement;
    const rect = wrap.getBoundingClientRect();
    bCanvas.width = rect.width;
    bCanvas.height = rect.height;
    bCtx.fillStyle = '#fff';
    bCtx.fillRect(0, 0, bCanvas.width, bCanvas.height);
  }

  function bHandleTouch(e) {
    if (e.type !== 'touchend') e.preventDefault();
    const touch = e.touches[0];
    const mouseEvent = new MouseEvent(e.type === 'touchstart' ? 'mousedown' : 'mousemove', {
      clientX: touch.clientX,
      clientY: touch.clientY
    });
    bCanvas.dispatchEvent(mouseEvent);
  }

  function bStartDrawing(e) {
    bIsDrawing = true;
    const rect = bCanvas.getBoundingClientRect();
    bLastX = e.clientX - rect.left;
    bLastY = e.clientY - rect.top;
    if (!bHasSignature) {
      const ph = document.getElementById('batchSigPlaceholder');
      if (ph) ph.style.opacity = '0';
      bHasSignature = true;
    }
  }

  function bDraw(e) {
    if (!bIsDrawing) return;
    const rect = bCanvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    bCtx.beginPath();
    bCtx.moveTo(bLastX, bLastY);
    bCtx.lineTo(x, y);
    bCtx.strokeStyle = '#000';
    bCtx.lineWidth = 3;
    bCtx.lineCap = 'round';
    bCtx.lineJoin = 'round';
    bCtx.stroke();
    bLastX = x;
    bLastY = y;
  }

  function bStopDrawing() { bIsDrawing = false; }

  // ========================
  // MODAL (single item)
  // ========================
  function openJobModal(job, type) {
    activeJob = job;
    const grid = document.getElementById('modalDetailGrid');
    
    const details = [
      { label: 'รหัสเอกสาร', val: job.Key },
      { label: 'ผู้ส่ง / ต้นทาง', val: `${job['ชื่อผู้ส่ง'] || '-'} (${job['แผนก ต้นทาง'] || '-'})` },
      { label: 'ผู้รับ / ปลายทาง', val: `${job['ชื่อผู้รับ'] || '-'} (${job['แผนก ปลายทาง'] || '-'})` },
      { label: 'สาขา', val: job['ส่งจากสาขา'] || '-' },
      { label: 'รายละเอียด', val: job['รายละเอียด'] || '-' },
      { label: 'จำนวน', val: job['จำนวน'] || '1' },
      { label: 'เวลาทำรายการ', val: job['เวลาทำรายการ'] || '-' }
    ];

    if (type === 'finished') {
      details.push({ label: 'เวลาที่รับ', val: job['Dropoff'] || '-' });
      if (job['Remark']) details.push({ label: 'หมายเหตุ', val: job['Remark'] });
    }

    grid.innerHTML = details.map(d => `
      <div class="detail-item">
        <div class="detail-label">${d.label}</div>
        <div class="detail-value">${d.val}</div>
      </div>
    `).join('');

    const sigSection = document.getElementById('sigSection');
    const remarkSection = document.getElementById('remarkSection');
    const footer = document.getElementById('modalFooter');
    
    sigSection.style.display = 'none';
    remarkSection.style.display = 'none';
    footer.style.display = 'flex';
    modalBtn.disabled = false;
    modalBtn.className = 'btn btn-primary';
    
    if (type === 'pending') {
      document.getElementById('modalTitle').textContent = 'ตรวจสอบการรับเข้าระบบ';
      modalLabel.textContent = 'รับเข้าระบบ (เพื่อนำส่ง)';
    } else if (type === 'started') {
      document.getElementById('modalTitle').textContent = 'ยืนยันการส่งมอบ';
      modalLabel.textContent = 'บันทึกการส่งมอบ';
      modalBtn.classList.add('started');
      sigSection.style.display = 'flex';
      remarkSection.style.display = 'block';
      document.getElementById('remarkInput').value = '';
      clearSignature();
    } else if (type === 'finished') {
      document.getElementById('modalTitle').textContent = 'รายละเอียดการรับ';
      footer.style.display = 'none';
      
      const sigUrl = job['Txt_01'] || job['Dropoff Signature'] || '';
      if (isValidSignatureSrc(sigUrl)) {
        sigSection.style.display = 'flex';
        const wrap = document.getElementById('sigCanvasWrap');
        const receiverName = (job['ชื่อผู้รับ'] || '').replace(/'/g, "\\'");
        wrap.innerHTML = `<img src="${sigUrl}" style="width:100%; height:100%; object-fit:contain; background:white; cursor:pointer;" data-receiver="${receiverName}" />`;
        const img = wrap.querySelector('img');
        img.addEventListener('click', () => {
          document.getElementById('sigViewerImg').src = sigUrl;
          document.getElementById('sigViewerLabel').textContent = 'ลายเซ็นผู้รับ: ' + (job['ชื่อผู้รับ'] || '');
          document.querySelector('.sig-viewer-overlay').classList.add('active');
        });
        document.getElementById('sigClearBtn').style.display = 'none';
      }
    }

    modalOverlay.classList.add('active');
    document.body.style.overflow = 'hidden';
    viewingKey = job && job.Key ? job.Key : null;
    sendHeartbeat();

    if (type === 'started') resizeCanvas();
  }

  function closeModal() {
    modalOverlay.classList.remove('active');
    document.body.style.overflow = 'auto';
    activeJob = null;
    viewingKey = null;
    sendHeartbeat();
    
    const wrap = document.getElementById('sigCanvasWrap');
    if (wrap.querySelector('img')) {
      wrap.innerHTML = `
        <canvas id="sigCanvas" aria-label="พื้นที่เซ็นชื่อ"></canvas>
        <div class="sig-placeholder" id="sigPlaceholder">
          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" opacity="0.4"><path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/></svg>
          <p>เซ็นชื่อในพื้นที่นี้</p>
        </div>
      `;
      document.getElementById('sigClearBtn').style.display = 'block';
      initSignaturePad();
    }
  }

  function initModalHandlers() {
    modalCloseBtn.addEventListener('click', closeModal);
    modalOverlay.addEventListener('click', e => {
      if (e.target === modalOverlay) closeModal();
    });

    modalBtn.addEventListener('click', () => {
      if (!activeJob) return;
      
      if (currentTab === 'pending') {
        updateJobStatus(activeJob.Key, 'Started');
      } else if (currentTab === 'started') {
        if (!hasSignature) {
          showToast('กรุณาเซ็นชื่อรับเอกสาร', 'error');
          return;
        }
        
        const sigBase64 = getSignatureData();
        const remark = document.getElementById('remarkInput').value;
        const now = new Date().toLocaleString('th-TH');
        
        updateJobStatus(activeJob.Key, 'Finished', {
          signature: sigBase64,
          dropoff: now,
          remark: remark
        });
      }
    });
  }

  // ========================
  // SIGNATURE (main modal)
  // ========================
  function initSignaturePad() {
    canvas = document.getElementById('sigCanvas');
    if (!canvas) return;
    ctx = canvas.getContext('2d', { willReadFrequently: true });
    
    document.getElementById('sigClearBtn').addEventListener('click', clearSignature);

    canvas.addEventListener('mousedown', startDrawing);
    canvas.addEventListener('mousemove', draw);
    canvas.addEventListener('mouseup', stopDrawing);
    canvas.addEventListener('mouseout', stopDrawing);

    canvas.addEventListener('touchstart', handleTouch, { passive: false });
    canvas.addEventListener('touchmove', handleTouch, { passive: false });
    canvas.addEventListener('touchend', stopDrawing);

    window.addEventListener('resize', resizeCanvas);
  }

  function handleTouch(e) {
    if (e.type !== 'touchend') e.preventDefault();
    const touch = e.touches[0];
    const mouseEvent = new MouseEvent(e.type === 'touchstart' ? 'mousedown' : 'mousemove', {
      clientX: touch.clientX,
      clientY: touch.clientY
    });
    canvas.dispatchEvent(mouseEvent);
  }

  function startDrawing(e) {
    isDrawing = true;
    const rect = canvas.getBoundingClientRect();
    lastX = e.clientX - rect.left;
    lastY = e.clientY - rect.top;
    if (!hasSignature) {
      document.getElementById('sigPlaceholder').style.opacity = '0';
      hasSignature = true;
    }
  }

  function draw(e) {
    if (!isDrawing) return;
    const rect = canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    ctx.beginPath();
    ctx.moveTo(lastX, lastY);
    ctx.lineTo(x, y);
    ctx.strokeStyle = '#000';
    ctx.lineWidth = 3;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.stroke();
    lastX = x;
    lastY = y;
  }

  function stopDrawing() { isDrawing = false; }

  function clearSignature() {
    if (!canvas) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    hasSignature = false;
    const placeholder = document.getElementById('sigPlaceholder');
    if (placeholder) placeholder.style.opacity = '1';
  }

  function getSignatureData() {
    return compressSignature(canvas);
  }

  function isValidSignatureSrc(src) {
    if (!src) return false;
    const s = String(src).trim();
    return s.startsWith('http') || s.startsWith('data:image/');
  }

  function escapeAttr(s) {
    return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  // Downscale signature canvas to ≤600px width + export as JPEG 0.82
  // Reduces payload from ~50KB PNG → ~5-10KB JPEG, making Drive upload much faster
  function compressSignature(srcCanvas) {
    const MAX_W = 600;
    const scale = Math.min(1, MAX_W / srcCanvas.width);
    const w = Math.round(srcCanvas.width * scale);
    const h = Math.round(srcCanvas.height * scale);
    const tmp = document.createElement('canvas');
    tmp.width = w; tmp.height = h;
    const tctx = tmp.getContext('2d');
    tctx.fillStyle = '#fff';
    tctx.fillRect(0, 0, w, h);
    tctx.drawImage(srcCanvas, 0, 0, w, h);
    return tmp.toDataURL('image/jpeg', 0.82);
  }

  function resizeCanvas() {
    if (!canvas) return;
    const wrap = canvas.parentElement;
    const rect = wrap.getBoundingClientRect();
    
    let temp = null;
    if (hasSignature) {
      temp = ctx.getImageData(0, 0, canvas.width, canvas.height);
    }
    
    canvas.width = rect.width;
    canvas.height = rect.height;
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    if (temp) ctx.putImageData(temp, 0, 0);
  }

  // ========================
  // NAVIGATION
  // ========================
  function initTabs() {
    navBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        const target = btn.dataset.panel;
        if (target === currentTab) return;
        
        navBtns.forEach(b => {
          b.classList.remove('active');
          b.setAttribute('aria-selected', 'false');
        });
        btn.classList.add('active');
        btn.setAttribute('aria-selected', 'true');
        
        panels.forEach(p => {
          if (p.id === `panel-${target}`) {
            p.classList.remove('hidden');
          } else {
            p.classList.add('hidden');
          }
        });
        
        currentTab = target;
        updateBatchBar();
        window.scrollTo(0, 0);
      });
    });
  }

  // ========================
  // TOASTS
  // ========================
  function showToast(message, type = 'success') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    const icon = type === 'success' ? 
      '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="var(--accent-emerald)" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>' :
      '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="var(--accent-red)" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>';
      
    toast.innerHTML = `${icon}<span>${message}</span>`;
    container.appendChild(toast);
    
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(-10px)';
      toast.style.transition = 'all 0.3s';
      setTimeout(() => toast.remove(), 300);
    }, 3000);
  }

  // ====== Swipe to delete (mobile/tablet) ======
  function enableSwipeDelete(card, key) {
    const isTouch = window.matchMedia('(hover: none) and (pointer: coarse)').matches;
    if (!isTouch) return;

    const content = card.querySelector('.card-content');
    const deleteBtn = card.querySelector('.card-delete-bg');
    if (!content || !deleteBtn) return;

    const REVEAL = 88;
    const THRESHOLD = 40;
    let startX = 0;
    let startY = 0;
    let dx = 0;
    let dragging = false;
    let locked = null;

    content.addEventListener('touchstart', (e) => {
      const t = e.touches[0];
      startX = t.clientX;
      startY = t.clientY;
      dx = 0;
      dragging = true;
      locked = null;
      card.classList.add('swiping');
    }, { passive: true });

    content.addEventListener('touchmove', (e) => {
      if (!dragging) return;
      const t = e.touches[0];
      const mx = t.clientX - startX;
      const my = t.clientY - startY;
      if (locked === null) {
        if (Math.abs(mx) > 8 || Math.abs(my) > 8) {
          locked = Math.abs(mx) > Math.abs(my) ? 'x' : 'y';
        }
      }
      if (locked !== 'x') return;
      dx = Math.min(0, mx);
      if (dx < -REVEAL) dx = -REVEAL + (dx + REVEAL) * 0.2;
      content.style.transform = 'translateX(' + dx + 'px)';
    }, { passive: true });

    content.addEventListener('touchend', () => {
      if (!dragging) return;
      dragging = false;
      card.classList.remove('swiping');
      content.style.transform = '';
      if (dx < -THRESHOLD) {
        card.classList.add('swipe-revealed');
      } else {
        card.classList.remove('swipe-revealed');
      }
    });

    deleteBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      showDeleteDialog(key, card);
    });
  }

  function showDeleteDialog(key, card) {
    if (document.getElementById('delete-dialog')) return;
    const overlay = document.createElement('div');
    overlay.id = 'delete-dialog';
    overlay.style.cssText =
      'position:fixed;inset:0;background:rgba(0,0,0,0.6);backdrop-filter:blur(8px);' +
      '-webkit-backdrop-filter:blur(8px);z-index:9999;display:flex;align-items:center;' +
      'justify-content:center;padding:16px;';

    const box = document.createElement('div');
    box.style.cssText =
      'background:var(--panel-bg,#0e1729);color:var(--text-main,#fff);border:1px solid var(--border,#1f2937);' +
      'border-radius:16px;padding:20px;max-width:420px;width:100%;box-shadow:0 20px 60px rgba(0,0,0,0.5);';

    box.innerHTML =
      '<div style="display:flex;align-items:center;gap:12px;margin-bottom:12px;">' +
      '<div style="width:44px;height:44px;border-radius:12px;background:linear-gradient(135deg,#dc2626,#b91c1c);display:flex;align-items:center;justify-content:center;font-size:24px;color:#fff;"><i class="bx bx-trash"></i></div>' +
      '<div style="flex:1;"><div style="font-weight:700;font-size:16px;">ลบรายการ ' + key + '</div>' +
      '<div style="font-size:12px;color:var(--text-muted,#94a3b8);">กรุณาระบุเหตุผลการลบ</div></div></div>' +
      '<textarea id="delete-note" rows="3" placeholder="หมายเหตุ (จำเป็น)" style="width:100%;padding:10px;border:1px solid var(--border,#1f2937);background:var(--card-bg,#0a1220);color:var(--text-main,#fff);border-radius:10px;font-family:inherit;font-size:16px;resize:vertical;outline:none;box-sizing:border-box;"></textarea>' +
      '<div style="display:flex;gap:8px;margin-top:14px;">' +
      '<button id="del-cancel" style="flex:1;padding:10px;border:1px solid var(--border,#1f2937);background:transparent;color:var(--text-main,#fff);border-radius:10px;font-weight:600;cursor:pointer;font-family:inherit;">ยกเลิก</button>' +
      '<button id="del-confirm" style="flex:1;padding:10px;border:none;background:linear-gradient(135deg,#dc2626,#b91c1c);color:#fff;border-radius:10px;font-weight:600;cursor:pointer;font-family:inherit;">ยืนยันลบ</button>' +
      '</div>';

    overlay.appendChild(box);
    document.body.appendChild(overlay);
    const prevBodyOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    setTimeout(() => document.getElementById('delete-note').focus({ preventScroll: true }), 50);

    const close = () => {
      document.body.style.overflow = prevBodyOverflow;
      overlay.remove();
    };
    document.getElementById('del-cancel').onclick = () => {
      close();
      card.classList.remove('swipe-revealed');
    };
    document.getElementById('del-confirm').onclick = async () => {
      const note = document.getElementById('delete-note').value.trim();
      if (!note) {
        document.getElementById('delete-note').style.borderColor = '#dc2626';
        return;
      }
      const btn = document.getElementById('del-confirm');
      btn.disabled = true;
      btn.textContent = 'กำลังลบ...';
      try {
        const res = await fetch('/api/jobs/update', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ key, action: 'cancel', note, _userId: userId })
        });
        const json = await res.json();
        if (res.status === 409 || json.conflict) {
          showConflictToast(json.error || 'รายการนี้กำลังถูกใช้งานโดยผู้ใช้งานท่านอื่น');
          throw new Error(json.error || 'Conflict');
        }
        if (!json.success) throw new Error(json.error || 'ลบไม่สำเร็จ');
        close();
        card.style.transition = 'opacity 0.3s, transform 0.3s';
        card.style.opacity = '0';
        card.style.transform = 'translateX(-100%)';
        setTimeout(() => {
          card.remove();
          fetchData(true);
        }, 300);
      } catch (err) {
        alert('ลบไม่สำเร็จ: ' + err.message);
        btn.disabled = false;
        btn.textContent = 'ยืนยันลบ';
      }
    };
  }

  // ====== PWA: register service worker (required for install prompt) ======
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('/sw.js').catch((err) => {
      console.warn('SW registration failed:', err);
    });
  }

  // ====== PWA Install Prompt ======
  setupInstallPrompt();

  function setupInstallPrompt() {
    const STORAGE_KEY = 'pwa-install-dismissed';
    const isStandalone =
      window.matchMedia('(display-mode: standalone)').matches ||
      window.navigator.standalone === true;
    if (isStandalone) return;
    if (localStorage.getItem(STORAGE_KEY) === '1') return;

    const ua = window.navigator.userAgent;
    const isIOS = /iPad|iPhone|iPod/.test(ua) && !window.MSStream;

    let deferredPrompt = null;

    window.addEventListener('beforeinstallprompt', (e) => {
      e.preventDefault();
      deferredPrompt = e;
      setTimeout(() => showInstallDialog(false), 1500);
    });

    window.addEventListener('appinstalled', () => {
      localStorage.setItem(STORAGE_KEY, '1');
      const d = document.getElementById('pwa-install-dialog');
      if (d) d.remove();
    });

    if (isIOS) {
      setTimeout(() => showInstallDialog(true), 1500);
    }

    function showInstallDialog(iosMode) {
      if (document.getElementById('pwa-install-dialog')) return;

      const overlay = document.createElement('div');
      overlay.id = 'pwa-install-dialog';
      overlay.style.cssText =
        'position:fixed;inset:0;background:rgba(0,0,0,0.6);backdrop-filter:blur(8px);' +
        '-webkit-backdrop-filter:blur(8px);z-index:9999;display:flex;align-items:flex-end;' +
        'justify-content:center;padding:16px;animation:fadeIn 0.3s;';

      const card = document.createElement('div');
      card.style.cssText =
        'background:var(--panel-bg,#0e1729);color:var(--text-main,#fff);border:1px solid var(--border,#1f2937);' +
        'border-radius:16px;padding:20px;max-width:420px;width:100%;box-shadow:0 20px 60px rgba(0,0,0,0.5);' +
        'margin-bottom:calc(env(safe-area-inset-bottom,0px) + 16px);animation:slideUp 0.3s;';

      const iosHelp = iosMode
        ? '<div style="font-size:13px;color:var(--text-muted,#94a3b8);margin-top:10px;line-height:1.6;">' +
          'แตะปุ่ม <b>แชร์</b> <i class="bx bx-share"></i> ที่แถบล่างของ Safari แล้วเลือก <b>"เพิ่มที่หน้าจอโฮม"</b></div>'
        : '';

      card.innerHTML =
        '<div style="display:flex;align-items:center;gap:12px;margin-bottom:8px;">' +
        '<div style="width:44px;height:44px;border-radius:12px;background:linear-gradient(135deg,#3b82f6,#8b5cf6);' +
        'display:flex;align-items:center;justify-content:center;font-size:24px;color:#fff;">' +
        '<i class="bx bx-download"></i></div>' +
        '<div style="flex:1;"><div style="font-weight:700;font-size:16px;">ติดตั้ง DocDelivery</div>' +
        '<div style="font-size:12px;color:var(--text-muted,#94a3b8);">เปิดแบบเต็มหน้าจอ ไม่มีแถบเบราว์เซอร์</div></div>' +
        '</div>' +
        '<div style="font-size:14px;color:var(--text-main,#e2e8f0);margin-top:6px;">' +
        'คุณต้องการสร้างไอคอน DocDelivery บนหน้าจอหลักหรือไม่?</div>' +
        iosHelp +
        '<div style="display:flex;gap:8px;margin-top:16px;">' +
        '<button id="pwa-cancel" style="flex:1;padding:10px;border:1px solid var(--border,#1f2937);' +
        'background:transparent;color:var(--text-main,#fff);border-radius:10px;font-weight:600;cursor:pointer;font-family:inherit;">ไม่ใช่ตอนนี้</button>' +
        (iosMode
          ? '<button id="pwa-ok" style="flex:1;padding:10px;border:none;background:linear-gradient(135deg,#3b82f6,#8b5cf6);color:#fff;border-radius:10px;font-weight:600;cursor:pointer;font-family:inherit;">เข้าใจแล้ว</button>'
          : '<button id="pwa-ok" style="flex:1;padding:10px;border:none;background:linear-gradient(135deg,#3b82f6,#8b5cf6);color:#fff;border-radius:10px;font-weight:600;cursor:pointer;font-family:inherit;">ติดตั้งเลย</button>') +
        '</div>';

      overlay.appendChild(card);
      document.body.appendChild(overlay);

      const close = () => overlay.remove();
      document.getElementById('pwa-cancel').onclick = () => {
        localStorage.setItem(STORAGE_KEY, '1');
        close();
      };
      document.getElementById('pwa-ok').onclick = async () => {
        if (iosMode) {
          close();
          return;
        }
        if (deferredPrompt) {
          close();
          deferredPrompt.prompt();
          const { outcome } = await deferredPrompt.userChoice;
          if (outcome === 'dismissed') {
            localStorage.setItem(STORAGE_KEY, '1');
          }
          deferredPrompt = null;
        }
      };
    }
  }

  // ========================
  // MONITOR / DATA-LOAD LOG
  // ========================
  function timeAgo(ts) {
    const diff = Math.floor((Date.now() - ts) / 1000);
    if (diff < 60) return `${diff} วินาทีที่แล้ว`;
    if (diff < 3600) return `${Math.floor(diff / 60)} นาทีที่แล้ว`;
    if (diff < 86400) return `${Math.floor(diff / 3600)} ชั่วโมงที่แล้ว`;
    return `${Math.floor(diff / 86400)} วันที่แล้ว`;
  }

  function formatTs(ts) {
    return new Date(ts).toLocaleString('th-TH', {
      day: '2-digit', month: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit', second: '2-digit'
    });
  }

  function renderMonitorPanel() {
    const summaryEl = document.getElementById('monitorSummary');
    const listEl = document.getElementById('monitorLogList');

    const successLogs = loadLog.filter(l => l.success);
    const totalRecords = data.pending.length + data.started.length + totalFinishedCount;
    const lastLoad = loadLog.find(l => l.success);

    summaryEl.innerHTML = `
      <div class="monitor-stat">
        <div class="monitor-stat-label">รายการทั้งหมดบนหน้าจอ</div>
        <div class="monitor-stat-value">${totalRecords.toLocaleString()}</div>
        <div class="monitor-stat-sub">รอรับ ${data.pending.length} · ส่ง ${data.started.length} · สำเร็จ ${totalFinishedCount}</div>
      </div>
      <div class="monitor-stat">
        <div class="monitor-stat-label">โหลดสำเร็จทั้งหมด</div>
        <div class="monitor-stat-value">${successLogs.length}</div>
        <div class="monitor-stat-sub">ล้มเหลว ${loadLog.length - successLogs.length} ครั้ง</div>
      </div>
      <div class="monitor-stat">
        <div class="monitor-stat-label">โหลดล่าสุดเมื่อ</div>
        <div class="monitor-stat-value" style="font-size:14px;line-height:1.4">${lastLoad ? timeAgo(lastLoad.ts) : '—'}</div>
        <div class="monitor-stat-sub">${lastLoad ? formatTs(lastLoad.ts) : 'ยังไม่มีข้อมูล'}</div>
      </div>
    `;

    if (loadLog.length === 0) {
      listEl.innerHTML = '<div class="monitor-empty">ยังไม่มีบันทึกการโหลดข้อมูล</div>';
      return;
    }

    listEl.innerHTML = loadLog.map((log, idx) => {
      const badge = log.success
        ? '<span class="monitor-log-badge badge-ok">✓ สำเร็จ</span>'
        : '<span class="monitor-log-badge badge-err">✗ ล้มเหลว</span>';
      const countClass = log.count === 0 ? 'log-zero' : '';
      const itemClass = log.success ? 'log-success' : 'log-error';
      const errorDetail = !log.success && log.error
        ? `<div class="monitor-log-errtext">${log.error}</div>` : '';
      const recs = log.records || [];
      const expandBtn = recs.length
        ? `<button type="button" class="monitor-expand-btn" data-idx="${idx}">
             <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
             ดูรายการ (${recs.length})
           </button>` : '';
      const recRows = recs.map(r =>
        `<tr>
          <td class="mrt-key">${r.k || '—'}</td>
          <td class="mrt-sender">${r.s || '—'}</td>
          <td class="mrt-dept">${r.d || '—'}</td>
          <td class="mrt-time">${r.t || '—'}</td>
        </tr>`
      ).join('');
      const recTable = recs.length ? `
        <div class="monitor-rec-wrap hidden" id="mrec-${idx}">
          <table class="monitor-rec-table">
            <thead><tr><th>รหัส</th><th>ผู้ส่ง</th><th>แผนกปลายทาง</th><th>เวลา</th></tr></thead>
            <tbody>${recRows}</tbody>
          </table>
        </div>` : '';
      return `
        <div class="monitor-log-item ${itemClass}">
          <div class="monitor-log-name">${log.name}${badge}</div>
          <div class="monitor-log-count ${countClass}">${log.success ? log.count + ' รายการ' : '—'}</div>
          <div class="monitor-log-time">${formatTs(log.ts)}</div>
          <div class="monitor-log-ago">${timeAgo(log.ts)}</div>
          ${errorDetail}
          ${expandBtn}
          ${recTable}
        </div>`;
    }).join('');

    // expand/collapse toggle
    listEl.querySelectorAll('.monitor-expand-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const wrap = listEl.querySelector(`#mrec-${btn.dataset.idx}`);
        if (!wrap) return;
        const open = !wrap.classList.contains('hidden');
        wrap.classList.toggle('hidden', open);
        btn.classList.toggle('expanded', !open);
      });
    });
  }

  // PIN flow
  const monitorBtn = document.getElementById('monitorBtn');
  const monitorPinOverlay = document.getElementById('monitorPinOverlay');
  const monitorPinInputs = document.querySelectorAll('.pin-digit');
  const monitorPinError = document.getElementById('monitorPinError');
  const monitorPinCancel = document.getElementById('monitorPinCancel');
  const monitorPanel = document.getElementById('monitorPanel');
  const monitorPanelClose = document.getElementById('monitorPanelClose');

  function openPinModal() {
    monitorPinError.classList.add('hidden');
    monitorPinInputs.forEach(i => { i.value = ''; i.classList.remove('pin-error'); });
    monitorPinOverlay.classList.remove('hidden');
    monitorPinInputs[0].focus();
  }

  function closePinModal() {
    monitorPinOverlay.classList.add('hidden');
  }

  function checkPin() {
    const entered = Array.from(monitorPinInputs).map(i => i.value).join('');
    if (entered === MONITOR_PIN) {
      closePinModal();
      renderMonitorPanel();
      monitorPanel.classList.remove('hidden');
    } else {
      monitorPinInputs.forEach(i => {
        i.value = '';
        i.classList.add('pin-error');
        setTimeout(() => i.classList.remove('pin-error'), 400);
      });
      monitorPinError.classList.remove('hidden');
      monitorPinInputs[0].focus();
    }
  }

  monitorPinInputs.forEach((input, idx) => {
    input.addEventListener('input', () => {
      input.value = input.value.replace(/\D/g, '').slice(-1);
      monitorPinError.classList.add('hidden');
      if (input.value && idx < monitorPinInputs.length - 1) {
        monitorPinInputs[idx + 1].focus();
      }
      if (idx === monitorPinInputs.length - 1 && input.value) {
        checkPin();
      }
    });
    input.addEventListener('keydown', e => {
      if (e.key === 'Backspace' && !input.value && idx > 0) {
        monitorPinInputs[idx - 1].focus();
      }
    });
  });

  monitorBtn.addEventListener('click', openPinModal);
  monitorPinCancel.addEventListener('click', closePinModal);
  monitorPinOverlay.addEventListener('click', e => { if (e.target === monitorPinOverlay) closePinModal(); });
  monitorPanelClose.addEventListener('click', () => monitorPanel.classList.add('hidden'));
});
