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
const importExcelBtn = document.getElementById('import-excel-btn');
const importPDFBtn = document.getElementById('import-pdf-btn');

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

printBtn.onclick = async () => {
  try {
    // Show loading indicator
    const originalText = printBtn.textContent;
    printBtn.textContent = 'กำลังจับภาพหน้าจอ...';
    printBtn.disabled = true;
    
    // Show loading overlay
    const loadingOverlay = document.createElement('div');
    loadingOverlay.id = 'pdf-loading-overlay';
    loadingOverlay.innerHTML = `
      <div style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); display: flex; justify-content: center; align-items: center; z-index: 9999;">
        <div style="background: white; padding: 30px; border-radius: 10px; text-align: center; box-shadow: 0 4px 20px rgba(0,0,0,0.3);">
          <div style="border: 4px solid #f3f3f3; border-top: 4px solid #009900; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto 20px;"></div>
          <p style="font-size: 18px; margin: 0;">กำลังจับภาพหน้าจอ...</p>
          <p style="font-size: 14px; margin: 10px 0 0; color: #666;">กรุณารอสักครู่</p>
        </div>
      </div>
      <style>
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      </style>
    `;
    document.body.appendChild(loadingOverlay);
    
    // Call the main process to capture all parcels and save as PDF
    const result = await window.thaipost.captureAllParcels(lastResults.length);
    
    // Remove loading overlay
    document.body.removeChild(loadingOverlay);
    
    // Restore button state
    printBtn.textContent = originalText;
    printBtn.disabled = false;
    
    if (!result.success) {
      alert(result.message);
    } else {
      // Show success message
      const successMessage = document.createElement('div');
      successMessage.innerHTML = `
        <div style="position: fixed; top: 20px; right: 20px; background: #4CAF50; color: white; padding: 15px; border-radius: 5px; z-index: 10000; box-shadow: 0 2px 10px rgba(0,0,0,0.2);">
          บันทึก PDF สำเร็จ!
        </div>
      `;
      document.body.appendChild(successMessage);
      
      // Remove success message after 3 seconds
      setTimeout(() => {
        if (successMessage.parentNode) {
          successMessage.parentNode.removeChild(successMessage);
        }
      }, 3000);
    }
  } catch (error) {
    // Remove loading overlay if it exists
    const loadingOverlay = document.getElementById('pdf-loading-overlay');
    if (loadingOverlay && loadingOverlay.parentNode) {
      loadingOverlay.parentNode.removeChild(loadingOverlay);
    }
    
    // Restore button state
    printBtn.textContent = 'บันทึกเป็น PDF';
    printBtn.disabled = false;
    
    console.error('Error capturing page:', error);
    alert('เกิดข้อผิดพลาดในการจับภาพหน้าจอ: ' + error.message);
  }
};

settingBtn.onclick = () => {
  alert('Setting: เลือกเครื่องพิมพ์, คุณภาพ, ขนาดกระดาษ ได้ที่ dialog ของระบบเมื่อสั่ง Print (window.print)\n\nหากต้องการบันทึกค่าตั้งค่าเพิ่มเติม สามารถพัฒนาเพิ่มได้ในอนาคต');
};

importExcelBtn.onclick = async () => {
  try {
    console.log('Import Excel button clicked');

    // Show loading indicator
    const originalText = importExcelBtn.textContent;
    importExcelBtn.textContent = 'กำลังโหลดไฟล์...';
    importExcelBtn.disabled = true;

    console.log('Calling window.thaipost.importExcel()');
    const result = await window.thaipost.importExcel();
    console.log('Import result:', result);

    // Restore button state
    importExcelBtn.textContent = originalText;
    importExcelBtn.disabled = false;

    if (result.success) {
      console.log('Tracking numbers found:', result.trackingNumbers);
      // Put the tracking numbers in the input field
      if (input) {
        input.value = result.trackingNumbers.join('\n');
        console.log('Input field updated with tracking numbers');
      } else {
        console.error('Input element not found');
      }

      // Show success message
      alert(`นำเข้าเลขพัสดุ ${result.trackingNumbers.length} รายการเรียบร้อยแล้ว`);
    } else if (result.sheetSelectionRequired) {
      // Sheet selection dialog is already opened, do nothing here
      console.log('Sheet selection dialog opened');
    } else {
      console.log('Import error:', result.message);
      alert(result.message || 'ไม่สามารถอ่านไฟล์ Excel ได้');
    }
  } catch (error) {
    console.error('Error importing Excel:', error);
    importExcelBtn.textContent = 'นำเข้าไฟล์ Excel';
    importExcelBtn.disabled = false;
    alert('เกิดข้อผิดพลาดในการนำเข้าไฟล์ Excel: ' + error.message);
  }
};

// Set up excel import result listener immediately
console.log('Setting up excel-import-result listener');
console.log('electronAPI available:', !!window.electronAPI);
console.log('ipcRenderer available:', !!(window.electronAPI && window.electronAPI.ipcRenderer));

if (window.electronAPI && window.electronAPI.ipcRenderer) {
  try {
    console.log('ipcRenderer.on function available:', typeof window.electronAPI.ipcRenderer.on);
    window.electronAPI.ipcRenderer.on('excel-import-result', (event, result) => {
      console.log('Received excel import result from main process:', result);

      if (result.success) {
        console.log('Tracking numbers found:', result.trackingNumbers);
        const inputElement = document.getElementById('track-input');
        console.log('Input element:', inputElement);
        console.log('Input element value before:', inputElement ? inputElement.value : 'no input');

        // Put the tracking numbers in the input field
        if (inputElement) {
          inputElement.value = result.trackingNumbers.join('\n');
          console.log('Input element value after:', inputElement.value);
          console.log('Successfully updated input field with tracking numbers');
        } else {
          console.error('Input element not found');
        }

        // Show success message
        alert(`นำเข้าเลขพัสดุ ${result.trackingNumbers.length} รายการเรียบร้อยแล้ว`);
      } else if (result.sheetSelectionRequired) {
        // Sheet selection dialog is already opened, do nothing here
        console.log('Sheet selection dialog opened');
      } else {
        console.log('Import error:', result.message);
        alert(result.message || 'ไม่สามารถอ่านไฟล์ Excel ได้');
      }
    });
    console.log('excel-import-result listener set up successfully');
  } catch (error) {
    console.error('Error setting up excel-import-result listener:', error);
  }
} else {
  console.log('electronAPI not available for excel-import-result listener');
}

importPDFBtn.onclick = async () => {
  try {
    console.log('Import PDF button clicked');
    
    // Show loading indicator
    const originalText = importPDFBtn.textContent;
    importPDFBtn.textContent = 'กำลังโหลดไฟล์...';
    importPDFBtn.disabled = true;
    
    console.log('Calling window.thaipost.importPDF()');
    const result = await window.thaipost.importPDF();
    console.log('Import result:', result);
    
    // Restore button state
    importPDFBtn.textContent = originalText;
    importPDFBtn.disabled = false;
    
    if (result.success) {
      console.log('Tracking numbers found:', result.trackingNumbers);
      // Put the tracking numbers in the input field
      input.value = result.trackingNumbers.join('\n');
      
      // Show success message
      alert(`นำเข้าเลขพัสดุ ${result.trackingNumbers.length} รายการเรียบร้อยแล้ว`);
    } else {
      console.log('Import error:', result.message);
      alert(result.message || 'ไม่สามารถอ่านไฟล์ PDF ได้');
    }
  } catch (error) {
    console.error('Error importing PDF:', error);
    importPDFBtn.textContent = 'นำเข้าไฟล์ PDF';
    importPDFBtn.disabled = false;
    alert('เกิดข้อผิดพลาดในการนำเข้าไฟล์ PDF: ' + error.message);
  }
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
    // Removed the connector logic since we're using a single line in CSS
  });
  bar += '</div>';
  let timelineRow = bar;
  // Limit status items to fit on one page in PDF
  let limitedItems = items.slice().reverse();
  if (limitedItems.length > 8) {
    limitedItems = limitedItems.slice(0, 8);
  }
  let list = '<div class="status-list">';
  limitedItems.forEach(item => {
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
      content += `<div style="margin-top:18px;page-break-inside:avoid;"><b>ลายเซ็นผู้รับ</b><br/><img src="${latest.signature}" style="max-width:260px;max-height:120px;border:2px solid #388e3c;border-radius:8px;page-break-inside:avoid;"/></div>`;
    }
  } else if (r.error) {
    content += `<div style="color:#c62828;">เกิดข้อผิดพลาด: ${r.error}</div>`;
  } else {
    content += `<div>ไม่พบข้อมูล</div>`;
  }
  content += '</div>';
  return content;
}

// เรียก bindRefreshButtons หลัง render timeline และ refresh
const origBindRefreshButtons = bindRefreshButtons;
bindRefreshButtons = function() {
  origBindRefreshButtons();
  bindPrintOneButtons();
};

function bindPrintOneButtons() {
  document.querySelectorAll('.print-one-btn').forEach(btn => {
    btn.onclick = () => {
      alert('ฟังก์ชันนี้ถูกปิดใช้งานชั่วคราว');
    };
  });
}