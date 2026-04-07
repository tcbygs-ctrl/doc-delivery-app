document.addEventListener('DOMContentLoaded', () => {
  // Elements
  const tabs = ['pending', 'started', 'finished'];
  const data = { pending: [], started: [], finished: [] };
  let currentTab = 'pending';
  let activeJob = null;

  const navBtns = document.querySelectorAll('.nav-btn');
  const panels = document.querySelectorAll('.tab-panel');
  const refreshBtn = document.getElementById('refreshBtn');
  const syncDot = document.getElementById('syncDot');
  const syncLabel = document.getElementById('syncLabel');
  
  // Modal Elements
  const modalBtn = document.getElementById('confirmBtn');
  const modalLabel = document.getElementById('confirmBtnLabel');
  const modalOverlay = document.getElementById('modalOverlay');
  const modalCloseBtn = document.getElementById('modalCloseBtn');

  // Initialization
  initTabs();
  initModalHandlers();
  initSignaturePad();
  fetchData();
  
  // Auto refresh every 15s
  setInterval(() => fetchData(false), 15000);
  
  refreshBtn.addEventListener('click', () => {
    const icon = refreshBtn.querySelector('svg');
    icon.style.transform = 'rotate(180deg)';
    icon.style.transition = 'transform 0.3s';
    fetchData(true).finally(() => {
      setTimeout(() => {
        icon.style.transform = 'none';
        icon.style.transition = 'none';
      }, 300);
    });
  });

  // --- API & Data ---
  async function fetchData(force = false) {
    setSyncStatus('syncing', 'กำลังซิงค์...');
    try {
      const url = `/api/jobs${force ? '?refresh=1' : ''}`;
      const res = await fetch(url);
      const json = await res.json();
      if (!json.success) throw new Error(json.error);
      
      processRawData(json.data);
      renderAllTabs();
      setSyncStatus('online', 'ออนไลน์');
    } catch (err) {
      console.error(err);
      setSyncStatus('error', 'ออฟไลน์');
      if (force) showToast('ไม่สามารถดึงข้อมูลได้', 'error');
    }
  }

  function setSyncStatus(state, label) {
    syncDot.className = `sync-dot ${state}`;
    syncLabel.textContent = label;
  }

  function processRawData(rows) {
    data.pending = [];
    data.started = [];
    data.finished = [];
    
    rows.forEach(row => {
      if (!row['Key']) return;
      const status = String(row['Status'] || '').trim();
      const cancel = String(row['Cancel'] || '').trim().toLowerCase();
      
      if (cancel === 'yes' || cancel === 'true') return; // Skip cancelled

      // Determine array based on status
      if (status === 'pending' || status === '1' || status === '') {
        data.pending.push(row);
      } else if (status === 'Started' || status === '2') {
        data.started.push(row);
      } else if (status === 'Finished' || status === '3') {
        data.finished.push(row);
      }
    });

    data.started.sort((a, b) => (b['เวลาทำรายการ'] || '').localeCompare(a['เวลาทำรายการ'] || ''));
    data.finished.sort((a, b) => (b['Dropoff'] || b['เวลาทำรายการ'] || '').localeCompare(a['Dropoff'] || a['เวลาทำรายการ'] || ''));
  }

  async function updateJobStatus(jobKey, newStatus, extraData = {}) {
    const payload = { key: jobKey, status: newStatus, ...extraData };
    try {
      modalBtn.disabled = true;
      modalLabel.textContent = 'กำลังบันทึก...';
      
      const res = await fetch('/api/jobs/update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const json = await res.json();
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

  // --- UI Rendering ---
  function renderAllTabs() {
    renderPendingList(data.pending);
    renderSimpleList('started', data.started);
    renderSimpleList('finished', data.finished);
    
    // Update Counts & Badges
    tabs.forEach(tab => {
      const count = data[tab].length;
      document.getElementById(`count-${tab}`).textContent = `${count} รายการ`;
      
      const badge = document.getElementById(`badge-${tab}`);
      if (count > 0 && tab !== 'finished') {
        badge.textContent = count > 99 ? '99+' : count;
        badge.classList.remove('hidden');
      } else {
        badge.classList.add('hidden');
      }
    });
  }

  function renderPendingList(items) {
    const listEl = document.getElementById('list-pending');
    document.getElementById('loading-pending').classList.add('hidden');
    
    if (items.length === 0) {
      listEl.innerHTML = '';
      document.getElementById('empty-pending').classList.remove('hidden');
      return;
    }
    
    document.getElementById('empty-pending').classList.add('hidden');
    listEl.innerHTML = '';

    // Group by Date, then by Dept
    const groups = {};
    items.forEach(item => {
      const date = item['วันที่ส่งเอกสาร/พัสดุ'] || 'ไม่ระบุวันที่';
      const dept = item['แผนก ต้นทาง'] || 'ไม่ระบุแผนก';
      if (!groups[date]) groups[date] = {};
      if (!groups[date][dept]) groups[date][dept] = [];
      groups[date][dept].push(item);
    });

    const dates = Object.keys(groups).sort((a, b) => b.localeCompare(a));
    
    dates.forEach(date => {
      const depts = Object.keys(groups[date]).sort();
      depts.forEach(dept => {
        const groupEl = document.createElement('div');
        groupEl.className = 'group-header';
        groupEl.textContent = `${date} • ${dept} (${groups[date][dept].length})`;
        listEl.appendChild(groupEl);
        
        groups[date][dept].forEach(job => {
          listEl.appendChild(createJobCard(job, 'pending'));
        });
      });
    });
  }

  function renderSimpleList(tab, items) {
    const listEl = document.getElementById(`list-${tab}`);
    document.getElementById(`loading-${tab}`).classList.add('hidden');
    
    if (items.length === 0) {
      listEl.innerHTML = '';
      document.getElementById(`empty-${tab}`).classList.remove('hidden');
      return;
    }
    
    document.getElementById(`empty-${tab}`).classList.add('hidden');
    listEl.innerHTML = '';
    
    items.forEach(job => {
      listEl.appendChild(createJobCard(job, tab));
    });
  }

  function createJobCard(job, type) {
    const card = document.createElement('div');
    card.className = 'job-card';
    
    const sender = job['ชื่อผู้ส่ง'] || '-';
    const receiver = job['ชื่อผู้รับ'] || '-';
    const qty = job['จำนวน'] || '1';
    
    let infoHtml = '';
    if (type === 'pending' || type === 'started') {
      infoHtml = `
        <div class="info-row">
          <svg class="info-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>
          <div class="info-text">
            <div class="user-name">จาก: ${sender}</div>
            <div class="dept-name">${job['แผนก ต้นทาง'] || ''} ${job['ส่งจากสาขา'] ? `(${job['ส่งจากสาขา']})` : ''}</div>
          </div>
        </div>
        <div class="info-row mt-2">
          <svg class="info-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><polyline points="20 6 9 17 4 12"/></svg>
          <div class="info-text">
            <div class="user-name">ถึง: ${receiver}</div>
            <div class="dept-name">${job['แผนก ปลายทาง'] || ''}</div>
          </div>
        </div>
      `;
    } else {
      infoHtml = `
        <div class="info-row">
          <svg class="info-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
          <div class="info-text">
            <div class="user-name">ผู้รับ: ${receiver}</div>
            <div class="dept-name">เวลา: ${job['Dropoff'] || job['เวลาทำรายการ'] || '-'}</div>
          </div>
        </div>
      `;
    }

    card.innerHTML = `
      <div class="card-top">
        <span class="card-key">${job.Key}</span>
        <span class="card-time">${job['เวลาทำรายการ'] ? job['เวลาทำรายการ'].split(' ')[0] : ''}</span>
      </div>
      <div class="card-body">
        ${infoHtml}
      </div>
      <div class="card-footer">
        <div class="item-count">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="7" width="20" height="14" rx="2" ry="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg>
          พัสดุ: <strong>${qty}</strong>
        </div>
        ${type === 'pending' ? '<span class="status-pill status-pending">รอรับ</span>' : 
          type === 'started' ? '<span class="status-pill status-started">นำส่ง</span>' : 
          '<span class="status-pill status-finished">รับแล้ว</span>'}
      </div>
    `;

    card.addEventListener('click', () => openJobModal(job, type));
    return card;
  }

  // --- Modal Logic ---
  function openJobModal(job, type) {
    activeJob = job;
    const grid = document.getElementById('modalDetailGrid');
    
    // Setup Content
    const details = [
      { label: 'รหัสเอกสาร', val: job.Key },
      { label: 'ผู้ส่ง / ต้นทาง', val: `${job['ชื่อผู้ส่ง'] || '-'} (${job['แผนก ต้นทาง'] || '-'})` },
      { label: 'ผู้รับ / ปลายทาง', val: `${job['ชื่อผู้รับ'] || '-'} (${job['แผนก ปลายทาง'] || '-'})` },
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

    // Configure specific sections based on type
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
      footer.style.display = 'none'; // No action needed for finished
      
      // Show signature if exists
      if (job['Dropoff Signature']) {
        sigSection.style.display = 'flex';
        // Need to render image instead of canvas
        const wrap = document.getElementById('sigCanvasWrap');
        wrap.innerHTML = `<img src="${job['Dropoff Signature'].startsWith('data:') ? job['Dropoff Signature'] : 'data:image/png;base64,' + job['Dropoff Signature']}" style="width:100%; height:100%; object-fit:contain; background:white;">`;
        document.getElementById('sigClearBtn').style.display = 'none';
      }
    }

    modalOverlay.classList.add('active');
    document.body.style.overflow = 'hidden'; // Prevent background scrolling
    
    if (type === 'started') resizeCanvas();
  }

  function closeModal() {
    modalOverlay.classList.remove('active');
    document.body.style.overflow = 'auto';
    activeJob = null;
    
    // Restore canvas HTML if overridden by finished state
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
      initSignaturePad(); // Rebind events
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
        updateJobStatus(activeJob.Key, 'Started'); // Move to step 2
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

  // --- Signature Logic ---
  let canvas, ctx;
  let isDrawing = false;
  let hasSignature = false;
  let lastX = 0;
  let lastY = 0;

  function initSignaturePad() {
    canvas = document.getElementById('sigCanvas');
    if (!canvas) return;
    ctx = canvas.getContext('2d', { willReadFrequently: true });
    
    document.getElementById('sigClearBtn').addEventListener('click', clearSignature);

    // Mouse Events
    canvas.addEventListener('mousedown', startDrawing);
    canvas.addEventListener('mousemove', draw);
    canvas.addEventListener('mouseup', stopDrawing);
    canvas.addEventListener('mouseout', stopDrawing);

    // Touch Events
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

  function stopDrawing() {
    isDrawing = false;
  }

  function clearSignature() {
    if (!canvas) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    hasSignature = false;
    const placeholder = document.getElementById('sigPlaceholder');
    if(placeholder) placeholder.style.opacity = '1';
  }

  function getSignatureData() {
    // Return base64 without prefix if needed, or keep prefix based on backend requirements
    // Here we strip prefix to save space, but you can adjust if needed
    const dataUrl = canvas.toDataURL('image/png');
    return dataUrl; // data:image/png;base64,....
  }

  function resizeCanvas() {
    if (!canvas) return;
    const wrap = canvas.parentElement;
    const rect = wrap.getBoundingClientRect();
    
    // Save image
    let temp = null;
    if (hasSignature) {
      temp = ctx.getImageData(0, 0, canvas.width, canvas.height);
    }
    
    canvas.width = rect.width;
    canvas.height = rect.height;
    
    // Provide white background
    ctx.fillStyle = "#fff";
    ctx.fillRect(0,0,canvas.width,canvas.height);
    
    // Restore
    if (temp) {
      ctx.putImageData(temp, 0, 0);
    } else {
      ctx.fillStyle = "#fff";
      ctx.fillRect(0,0,canvas.width,canvas.height);
    }
  }

  // --- Navigation ---
  function initTabs() {
    navBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        const target = btn.dataset.panel;
        if (target === currentTab) return;
        
        // Update UI
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
        window.scrollTo(0, 0);
      });
    });
  }

  // --- Toasts ---
  function showToast(message, type = 'success') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    const icon = type === 'success' ? 
      '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="var(--accent-emerald)" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>' :
      '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="var(--accent-red)" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>';
      
    toast.innerHTML = `
      ${icon}
      <span>${message}</span>
    `;
    
    container.appendChild(toast);
    
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(-10px)';
      toast.style.transition = 'all 0.3s';
      setTimeout(() => toast.remove(), 300);
    }, 3000);
  }
});
