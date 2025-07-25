const input = document.getElementById('track-input');
const btn = document.getElementById('track-btn');
const timeline = document.getElementById('timeline');
const searchSection = document.getElementById('search-section');
const resultSection = document.getElementById('result-section');
const backBtn = document.getElementById('back-btn');
const printControls = document.getElementById('print-controls');
const previewBtn = document.getElementById('preview-btn');
const printBtn = document.getElementById('print-btn');
const settingBtn = document.getElementById('setting-btn');
const printPreviewModal = document.getElementById('print-preview-modal');
const printPreviewPages = document.getElementById('print-preview-pages');
const closePreviewBtn = document.getElementById('close-preview-btn');

let lastResults = [];

function fadeOutIn(hideEl, showEl, cb) {
  hideEl.classList.remove('fade-in');
  hideEl.classList.add('fade-out');
  setTimeout(() => {
    hideEl.style.display = 'none';
    hideEl.classList.remove('fade-out');
    showEl.style.display = '';
    showEl.classList.add('fade-in');
    setTimeout(() => showEl.classList.remove('fade-in'), 500);
    if (cb) cb();
  }, 400);
}

const stepMap = [
  { key: 'รับฝาก', label: 'รับเข้าระบบ', icon: 'insert.png' },
  { key: 'ระหว่างขนส่ง', label: 'ระหว่างขนส่ง', icon: 'transport.png' },
  { key: 'อยู่ระหว่างการนำจ่าย', label: 'ออกไปนำจ่าย', icon: 'shipping.png' },
  { key: 'นำจ่ายสำเร็จ', label: 'นำจ่ายสำเร็จ', icon: 'shippingsuccess.png' },
];

btn.onclick = async () => {
  timeline.innerHTML = '<div class="loader"></div>';
  const raw = input.value.trim();
  if (!raw) {
    timeline.innerHTML = '<div class="alert">กรุณาใส่เลขพัสดุ</div>';
    return;
  }
  fadeOutIn(searchSection, resultSection);
  printControls.style.display = '';
  // รองรับหลายเลข (split ด้วย space หรือ comma)
  const barcodes = raw.split(/[,\s]+/).filter(Boolean);
  let allResults = [];
  for (const barcode of barcodes) {
    try {
      const result = await window.thaipost.trackParcel(barcode);
      const itemsObj = result.response.items;
      let found = false;
      for (const realBarcode in itemsObj) {
        const items = itemsObj[realBarcode] || [];
        allResults.push({ barcode: realBarcode, items });
        found = true;
      }
      if (!found) {
        allResults.push({ barcode, items: [] });
      }
    } catch (e) {
      allResults.push({ barcode, error: e });
    }
  }
  lastResults = allResults;
  timeline.innerHTML = allResults.map((r, idx) => renderParcelTimeline(r, idx + 1, allResults.length)).join('');
  bindRefreshButtons();
};

function bindRefreshButtons() {
  document.querySelectorAll('.refresh-btn').forEach(btn => {
    btn.onclick = async function() {
      const barcode = btn.getAttribute('data-barcode');
      btn.disabled = true;
      btn.textContent = 'กำลังโหลด...';
      try {
        const result = await window.thaipost.trackParcel(barcode);
        const itemsObj = result.response.items;
        let found = false;
        for (const realBarcode in itemsObj) {
          const items = itemsObj[realBarcode] || [];
          const idx = lastResults.findIndex(r => r.barcode === barcode);
          if (idx !== -1) lastResults[idx] = { barcode: realBarcode, items };
          const block = btn.closest('.parcel-block');
          block.outerHTML = renderParcelTimeline({ barcode: realBarcode, items }, idx + 1, lastResults.length);
          found = true;
        }
        if (!found) {
          btn.textContent = 'ไม่พบข้อมูล';
        }
      } catch (e) {
        btn.textContent = 'เกิดข้อผิดพลาด';
      }
      bindRefreshButtons();
    };
  });
}

backBtn.onclick = () => {
  fadeOutIn(resultSection, searchSection, () => {
    timeline.innerHTML = '';
    input.value = '';
    input.focus();
    printControls.style.display = 'none';
    lastResults = [];
  });
};

previewBtn.onclick = () => {
  printPreviewPages.innerHTML = lastResults.map((r, idx) => renderPrintPage(r, idx + 1, lastResults.length)).join('');
  printPreviewModal.style.display = '';
};

closePreviewBtn.onclick = () => {
  printPreviewModal.style.display = 'none';
};

printBtn.onclick = () => {
  // สร้าง print-page ชั่วคราวแล้วสั่ง print
  const printWindow = window.open('', '', 'width=900,height=1200');
  printWindow.document.write('<html><head><title>Print</title><style>' +
    document.querySelector('style').innerHTML +
    '@media print { body *:not(.print-page) { display: none !important; } .print-page { display: block !important; page-break-after: always; } }' +
    '</style></head><body>' +
    lastResults.map((r, idx) => renderPrintPage(r, idx + 1, lastResults.length)).join('') +
    '</body></html>');
  printWindow.document.close();
  printWindow.focus();
  setTimeout(() => { printWindow.print(); printWindow.close(); }, 500);
};

settingBtn.onclick = () => {
  alert('Setting: เลือกเครื่องพิมพ์, คุณภาพ, ขนาดกระดาษ ได้ที่ dialog ของระบบเมื่อสั่ง Print (window.print)\n\nหากต้องการบันทึกค่าตั้งค่าเพิ่มเติม สามารถพัฒนาเพิ่มได้ในอนาคต');
};

function renderParcelTimeline({ barcode, items, error }, order, total) {
  if (error) {
    return `<div class="parcel-block" data-barcode="${barcode}"><div class="parcel-order">#${order}</div><h3>เลขพัสดุ: ${barcode}</h3><button class="refresh-btn" data-barcode="${barcode}">เช็คใหม่</button><button class="print-one-btn" data-barcode="${barcode}">ปริ้น</button><div class="alert error">เกิดข้อผิดพลาด: ${error}</div></div>${order < total ? '<hr class="parcel-divider"/>' : ''}`;
  }
  if (!items || !items.length) {
    return `<div class="parcel-block" data-barcode="${barcode}"><div class="parcel-order">#${order}</div><h3>เลขพัสดุ: ${barcode}</h3><button class="refresh-btn" data-barcode="${barcode}">เช็คใหม่</button><button class="print-one-btn" data-barcode="${barcode}">ปริ้น</button><div class="alert">ไม่พบข้อมูล</div></div>${order < total ? '<hr class="parcel-divider"/>' : ''}`;
  }
  // หาสถานะที่ถึงล่าสุด
  const completedSet = new Set(items.map(i => i.status_description));
  // หาค่า index สถานะล่าสุดที่มีใน items
  const latestStatusIdx = stepMap.reduce((acc, step, idx) => (
    items.some(i => i.status_description === step.key) ? idx : acc
  ), -1);
  // หาสถานะนำจ่ายสำเร็จ (ถ้ามี)
  const delivered = items.find(i => i.status_description === 'นำจ่ายสำเร็จ');
  const latest = items[items.length - 1] || {};
  let summaryBox = '';
  if (delivered) {
    summaryBox = `
      <div class="summary-box delivered">
        <div class="summary-main">${delivered.status_detail}</div>
        <div class="summary-meta">ชื่อผู้รับ : ${delivered.receiver_name || '-'}</div>
        <div class="summary-meta">สถานะ : ${delivered.delivery_description || '-'}</div>
      </div>
    `;
  } else if (latest) {
    summaryBox = `
      <div class="summary-box">
        <div class="summary-main">${latest.status_detail}</div>
        <div class="summary-meta">ชื่อผู้รับ : ${latest.receiver_name || '-'}</div>
        <div class="summary-meta">สถานะ : ${latest.delivery_description || latest.status_description || '-'}</div>
      </div>
    `;
  }
  let signatureBox = '';
  if (delivered && delivered.signature) {
    signatureBox = `<div class="signature-box">
      <div class="signature-title">ลายเซ็นผู้รับ</div>
      <a href="${delivered.signature}" target="_blank">
        <img src="${delivered.signature}" alt="signature" />
      </a>
    </div>`;
  }
  // Timeline bar inline style + watermark
  const connectorMeta = [
    { color: 'linear-gradient(90deg, #F44336 60%, #FF9800 100%)', label: 'ระหว่างขนส่ง', labelColor: '#F44336' },
    { color: 'linear-gradient(90deg, #4CAF50 60%, #009900 100%)', label: 'ออกไปนำจ่าย', labelColor: '#009900' },
    { color: 'linear-gradient(90deg, #4CAF50 60%, #009900 100%)', label: 'นำจ่ายสำเร็จ', labelColor: '#009900' }
  ];
  let bar = `<div class="timeline-bar-inline">`;
  stepMap.forEach((step, idx) => {
    const completed = completedSet.has(step.key);
    const isWatermark = idx > latestStatusIdx;
    bar += `<div class="timeline-step-inline${completed ? ' completed' : ''}${isWatermark ? ' watermark' : ''}">
      <div class="timeline-icon-inline" style="border-color:${completed ? '#4CAF50' : '#F44336'}; opacity:${isWatermark ? 0.3 : 1}">
        <img src="../${step.icon}" style="opacity:${isWatermark ? 0.3 : 1}" />
        ${isWatermark ? '<span class="icon-cross">✖</span>' : ''}
      </div>
      <div class="timeline-label-inline${completed ? ' completed' : ''}${isWatermark ? ' label-watermark' : ''}" style="opacity:${isWatermark ? 0.3 : 1}">${step.label}</div>
    </div>`;
    if (idx < stepMap.length - 1) {
      const meta = connectorMeta[idx];
      bar += `<div class="timeline-connector-inline${isWatermark ? ' connector-watermark' : ''}" style="background:${meta.color}; opacity:${(idx+1) > latestStatusIdx ? 0.2 : 1}"></div>`;
    }
  });
  bar += '</div>';
  let timelineRow = bar;
  let list = '<div class="status-list">';
  items.slice().reverse().forEach(item => {
    list += `<div class="status-item${item.status_description === 'นำจ่ายสำเร็จ' ? ' delivered' : ''}">
      <div class="date">${item.status_date} | ${item.status_description}</div>
      <div class="desc">${item.status_detail}</div>
      ${item.location ? `<div class="meta"><span>สถานที่:</span> ${item.location}</div>` : ''}
      ${item.receiver_name ? `<div class="meta"><span>ผู้รับ:</span> ${item.receiver_name}</div>` : ''}
      ${item.delivery_description ? `<div class="meta">${item.delivery_description}</div>` : ''}
    </div>`;
  });
  list += '</div>';
  let statusSignatureRow = `<div class="status-signature-row">
    <div class="status-list-col">${list}</div>
    ${signatureBox ? `<div class="signature-list-col">${signatureBox}</div>` : ''}
  </div>`;
  return `<div class="parcel-block" data-barcode="${barcode}"><div class="parcel-order">#${order}</div><h3>เลขพัสดุ: ${barcode}</h3><button class="refresh-btn" data-barcode="${barcode}">เช็คใหม่</button><button class="print-one-btn" data-barcode="${barcode}">ปริ้น</button>${summaryBox}${timelineRow}${statusSignatureRow}</div>${order < total ? '<hr class="parcel-divider"/>' : ''}`;
}

function renderPrintPage(r, order, total) {
  // สามารถปรับแต่งหน้าตาใบปริ้นได้ตามต้องการ
  // ตัวอย่าง: แสดงเลขพัสดุ, สรุป, timeline สั้น, รายละเอียดล่าสุด, ลายเซ็น
  let content = `<div class="print-page">
    <h2>ใบตรวจสอบสถานะพัสดุ</h2>
    <div><b>ลำดับ:</b> ${order} / ${total}</div>
    <div><b>เลขพัสดุ:</b> ${r.barcode}</div>
    <hr/>
  `;
  if (r.items && r.items.length) {
    const latest = r.items[r.items.length - 1];
    content += `<div><b>สถานะล่าสุด:</b> ${latest.status_description} (${latest.status_date})</div>
      <div>${latest.status_detail}</div>
      <div>สถานที่: ${latest.location || '-'}</div>
      <div>ผู้รับ: ${latest.receiver_name || '-'}</div>
      <div>ผลการนำจ่าย: ${latest.delivery_description || '-'}</div>
    `;
    if (latest.signature) {
      content += `<div style="margin-top:18px;"><b>ลายเซ็นผู้รับ</b><br/><img src="${latest.signature}" style="max-width:260px;max-height:120px;border:2px solid #388e3c;border-radius:8px;"/></div>`;
    }
  } else if (r.error) {
    content += `<div style="color:#c62828;">เกิดข้อผิดพลาด: ${r.error}</div>`;
  } else {
    content += `<div>ไม่พบข้อมูล</div>`;
  }
  content += '</div>';
  return content;
}

function showPrintPreviewModal(r, order, total) {
  // Render HTML จริงของ .parcel-block
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = renderParcelTimeline(r, order, total);
  tempDiv.querySelectorAll('.refresh-btn, .print-one-btn, .paper-preview-btn').forEach(el => el.remove());
  const parcelBlockHtml = tempDiv.querySelector('.parcel-block').outerHTML;

  // Modal preview
  const modal = document.createElement('div');
  modal.style = 'position:fixed;left:0;top:0;right:0;bottom:0;background:rgba(0,0,0,0.25);z-index:2000;display:flex;align-items:center;justify-content:center;';
  modal.innerHTML = `
    <div style="position:relative;">
      <button id="close-print-preview" style="position:absolute;right:12px;top:12px;z-index:10;">ปิด</button>
      <button id="confirm-print" style="position:absolute;right:120px;top:12px;z-index:10;background:#009900;color:#fff;padding:8px 24px;border-radius:8px;border:none;font-weight:bold;box-shadow:0 1px 4px #00990011;">ยืนยัน</button>
      <div class="print-preview-paper">${parcelBlockHtml}</div>
    </div>
  `;
  document.body.appendChild(modal);

  document.getElementById('close-print-preview').onclick = () => modal.remove();
  document.getElementById('confirm-print').onclick = () => {
    modal.remove();
    printSingleParcelWithDialog(r, order, total);
  };
}

function printSingleParcelWithDialog(r, order, total) {
  // เปิด window ใหม่ทันที (ปลอดภัยกับ popup blocker)
  const printWindow = window.open('', '', 'width=900,height=1200');
  if (!printWindow) {
    alert('กรุณาอนุญาต popup หรือปิด popup blocker');
    return;
  }

  // ดึง style.css จาก <link> เดิม
  const styleHref = Array.from(document.querySelectorAll('link[rel="stylesheet"]'))
    .map(link => link.href)
    .find(href => href.includes('style.css'));
  const styleLink = styleHref ? `<link rel="stylesheet" href="${styleHref}">` : '';

  // Render HTML จริงของ .parcel-block
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = renderParcelTimeline(r, order, total);
  tempDiv.querySelectorAll('.refresh-btn, .print-one-btn, .paper-preview-btn').forEach(el => el.remove());
  const parcelBlockHtml = tempDiv.querySelector('.parcel-block').outerHTML;

  // inject HTML ทันที
  printWindow.document.write(`
    <html>
      <head>
        <title>Print Preview</title>
        ${styleLink}
        <style>
          body { background: #fff !important; }
          .print-preview-paper {
            background: #fff;
            width: 794px;
            min-height: 1123px;
            box-shadow: 0 4px 32px #0003;
            margin: 0 auto;
            padding: 32px 32px 32px 32px;
            box-sizing: border-box;
            border-radius: 8px;
            position: relative;
            overflow: hidden;
          }
          @media print {
            #print-btn { display: none !important; }
            body *:not(.parcel-block) { display: none !important; }
            .parcel-block { display: block !important; page-break-after: always; }
            .parcel-block { transform-origin: top left; }
          }
        </style>
      </head>
      <body>
        <div style="text-align:right;margin-bottom:16px;">
          <button id="print-btn" style="font-size:1em;padding:8px 24px;border-radius:8px;background:#009900;color:#fff;border:none;font-weight:bold;">Print</button>
        </div>
        <div class="print-preview-paper">${parcelBlockHtml}</div>
        <script>
          document.getElementById('print-btn').onclick = function() {
            window.print();
          };
        <\/script>
      </body>
    </html>
  `);
  printWindow.document.close();
  printWindow.focus();
}

// ปรับ bindPrintOneButtons ให้เรียก showPrintPreviewModal
function bindPrintOneButtons() {
  document.querySelectorAll('.print-one-btn').forEach(btn => {
    btn.onclick = () => {
      // เปิด window ใหม่ทันที (ปลอดภัยกับ popup blocker)
      const printWindow = window.open('', '', 'width=900,height=1200');
      if (!printWindow) {
        alert('กรุณาอนุญาต popup หรือปิด popup blocker');
        return;
      }
      const barcode = btn.getAttribute('data-barcode');
      const idx = lastResults.findIndex(r => r.barcode === barcode);
      if (idx !== -1) {
        // ดึง style.css
        const styleHref = Array.from(document.querySelectorAll('link[rel="stylesheet"]'))
          .map(link => link.href)
          .find(href => href.includes('style.css'));
        const styleLink = styleHref ? `<link rel="stylesheet" href="${styleHref}">` : '';
        // Render HTML จริง
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = renderParcelTimeline(lastResults[idx], idx + 1, lastResults.length);
        tempDiv.querySelectorAll('.refresh-btn, .print-one-btn, .paper-preview-btn').forEach(el => el.remove());
        const parcelBlockHtml = tempDiv.querySelector('.parcel-block').outerHTML;
        // inject HTML
        printWindow.document.write(`
          <html>
            <head>
              <title>Print Preview</title>
              ${styleLink}
              <style>
                body { background: #fff !important; }
                .print-preview-paper {
                  background: #fff;
                  width: 794px;
                  min-height: 1123px;
                  box-shadow: 0 4px 32px #0003;
                  margin: 0 auto;
                  padding: 32px 32px 32px 32px;
                  box-sizing: border-box;
                  border-radius: 8px;
                  position: relative;
                  overflow: hidden;
                }
                @media print {
                  #print-btn { display: none !important; }
                  body *:not(.parcel-block) { display: none !important; }
                  .parcel-block { display: block !important; page-break-after: always; }
                  .parcel-block { transform-origin: top left; }
                }
              </style>
            </head>
            <body>
              <div style="text-align:right;margin-bottom:16px;">
                <button id="print-btn" style="font-size:1em;padding:8px 24px;border-radius:8px;background:#009900;color:#fff;border:none;font-weight:bold;">Print</button>
              </div>
              <div class="print-preview-paper">${parcelBlockHtml}</div>
              <script>
                document.getElementById('print-btn').onclick = function() {
                  window.print();
                };
              <\/script>
            </body>
          </html>
        `);
        printWindow.document.close();
        printWindow.focus();
      }
    };
  });
}
// เรียก bindPrintOneButtons หลัง render timeline และ refresh
const origBindRefreshButtons = bindRefreshButtons;
bindRefreshButtons = function() {
  origBindRefreshButtons();
  bindPrintOneButtons();
}; 