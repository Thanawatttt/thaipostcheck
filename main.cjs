const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');
const { trackParcel } = require('./thailandpost.cjs');

// Global variables to store workbook data during sheet selection
let currentWorkbook = null;
let currentFilePath = null;
let sheetSelectorWindow = null;

function createWindow() {
  const win = new BrowserWindow({
    width: 900,
    height: 700,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      nodeIntegration: false,
      contextIsolation: true,
    }
  });
  win.loadFile('renderer/index.html');
}

app.whenReady().then(createWindow);

ipcMain.handle('track-parcel', async (event, trackingNumber) => {
  return await trackParcel(trackingNumber);
});

ipcMain.handle('get-sheet-names', async () => {
  console.log('Getting sheet names, currentWorkbook:', currentWorkbook);
  if (currentWorkbook && currentWorkbook.SheetNames) {
    console.log('Returning sheet names:', currentWorkbook.SheetNames);
    return currentWorkbook.SheetNames;
  }
  console.log('No current workbook found');
  return [];
});

ipcMain.handle('select-sheet', async (event, sheetName) => {
  console.log('Sheet selected:', sheetName);

  // Resolve the promise with the selected sheet name
  if (sheetSelectorWindow && sheetSelectorWindow.resolveSheetSelection) {
    sheetSelectorWindow.resolveSheetSelection(sheetName);
  }

  // Close the sheet selector window
  if (sheetSelectorWindow) {
    sheetSelectorWindow.close();
    sheetSelectorWindow = null;
  }

  return { success: true };
});

ipcMain.handle('cancel-sheet-selection', async () => {
  console.log('Sheet selection cancelled');

  // Reject the promise
  if (sheetSelectorWindow && sheetSelectorWindow.rejectSheetSelection) {
    sheetSelectorWindow.rejectSheetSelection(new Error('Sheet selection cancelled'));
  }

  // Close the sheet selector window
  if (sheetSelectorWindow) {
    sheetSelectorWindow.close();
    sheetSelectorWindow = null;
  }

  // Clear workbook data
  currentWorkbook = null;
  currentFilePath = null;
  return { success: false, message: 'ยกเลิกการเลือกไฟล์' };
});

ipcMain.on('request-sheet-names', (event) => {
  console.log('Received request for sheet names from renderer');
  console.log('Current workbook state:', !!currentWorkbook);
  console.log('Current workbook sheet names:', currentWorkbook ? currentWorkbook.SheetNames : 'No workbook');
  
  if (currentWorkbook && currentWorkbook.SheetNames && currentWorkbook.SheetNames.length > 0) {
    console.log('Sending sheet names to selector window:', currentWorkbook.SheetNames);
    try {
      event.sender.send('sheet-names', currentWorkbook.SheetNames);
      console.log('Sheet names sent successfully');
    } catch (error) {
      console.error('Error sending sheet names:', error);
    }
  } else {
    console.log('No valid workbook or sheet names available');
    try {
      event.sender.send('sheet-names', []);
    } catch (error) {
      console.error('Error sending empty sheet names array:', error);
    }
  }
});

ipcMain.handle('import-excel', async (event) => {
  try {
    console.log('Import Excel function called');

    // Show file dialog to select Excel file
    const { canceled, filePaths } = await dialog.showOpenDialog({
      title: 'เลือกไฟล์ Excel',
      filters: [
        { name: 'Excel Files', extensions: ['xlsx', 'xls'] }
      ],
      properties: ['openFile']
    });

    console.log('File dialog result:', { canceled, filePaths });

    if (canceled || !filePaths || filePaths.length === 0) {
      console.log('File selection canceled or no file selected');
      return { success: false, message: 'ยกเลิกการเลือกไฟล์' };
    }

    const filePath = filePaths[0];
    console.log('Selected file:', filePath);

    // Read the Excel file
    const workbook = XLSX.readFile(filePath);
    console.log('Workbook loaded, sheet names:', workbook.SheetNames);

    // If there's only one sheet, process it directly
    if (workbook.SheetNames.length === 1) {
      const sheetName = workbook.SheetNames[0];
      console.log('Only one sheet found, using sheet:', sheetName);

      const worksheet = workbook.Sheets[sheetName];

      // Convert to JSON
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      console.log('Sheet data (first 5 rows):', jsonData.slice(0, 5));

      // Extract tracking numbers from all cells (not just first column)
      const trackingNumbers = [];

      // Process all rows and columns
      for (let i = 0; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (row) {
          // Process all columns in the row
          for (let j = 0; j < row.length; j++) {
            const cellValue = row[j];
            console.log('Processing cell value:', cellValue);

            // Check if it looks like a tracking number
            if (typeof cellValue === 'string') {
              // More flexible pattern matching for Thai Post tracking numbers
              const trackingPattern = /^[A-Z]{2}[0-9]{9}[A-Z]{2}$/i;
              if (trackingPattern.test(cellValue.trim())) {
                console.log('Found tracking number:', cellValue.trim());
                trackingNumbers.push(cellValue.trim());
              } else {
                console.log('Cell value is string but does not match tracking pattern:', cellValue);
              }
            } else if (typeof cellValue === 'number') {
              console.log('Cell value is number:', cellValue);
            } else {
              console.log('Cell value type:', typeof cellValue, 'Value:', cellValue);
            }
          }
        }
      }

      console.log('Final tracking numbers list:', trackingNumbers);

      // Remove duplicates
      const uniqueTrackingNumbers = [...new Set(trackingNumbers)];
      console.log('Unique tracking numbers list:', uniqueTrackingNumbers);

      if (uniqueTrackingNumbers.length === 0) {
        return { success: false, message: 'ไม่พบเลขพัสดุในไฟล์ Excel ที่เลือก' };
      }

      return { success: true, trackingNumbers: uniqueTrackingNumbers };
    }
    // If there are multiple sheets, open sheet selection dialog and wait for selection
    else {
      console.log('Multiple sheets found, opening sheet selection dialog');

      // Store workbook data for later use
      currentWorkbook = workbook;
      currentFilePath = filePath;

      // Get the sender window
      const senderWindow = BrowserWindow.fromWebContents(event.sender);

      // Create sheet selection window
      sheetSelectorWindow = new BrowserWindow({
        width: 500,
        height: 400,
        parent: senderWindow,
        modal: true,
        webPreferences: {
          preload: path.join(__dirname, 'preload.cjs'),
          nodeIntegration: false,
          contextIsolation: true,
        }
      });

      // Create a promise that will resolve when a sheet is selected
      const sheetSelectionPromise = new Promise((resolve, reject) => {
        sheetSelectorWindow.resolveSheetSelection = resolve;
        sheetSelectorWindow.rejectSheetSelection = reject;

        // Handle window close without selection
        sheetSelectorWindow.on('closed', () => {
          reject(new Error('Sheet selection cancelled'));
        });
      });

      // Send sheet names to the selector window after it finishes loading
      sheetSelectorWindow.webContents.once('did-finish-load', () => {
        console.log('Sheet selector window loaded, sending sheet names:', workbook.SheetNames);
        if (sheetSelectorWindow && !sheetSelectorWindow.isDestroyed() && workbook) {
          try {
            // Send sheet names with a small delay to ensure renderer is ready
            setTimeout(() => {
              if (sheetSelectorWindow && !sheetSelectorWindow.isDestroyed()) {
                sheetSelectorWindow.webContents.send('sheet-names', workbook.SheetNames);
                console.log('Sheet names sent to selector window successfully');
              } else {
                console.log('Sheet selector window was destroyed before sending sheet names');
              }
            }, 200);
          } catch (error) {
            console.error('Error sending sheet names to selector window:', error);
          }
        } else {
          console.log('Sheet selector window was destroyed before sending sheet names or no workbook');
        }
      });

      sheetSelectorWindow.loadFile('renderer/sheet-selector.html');

      // Wait for sheet selection
      try {
        const selectedSheet = await sheetSelectionPromise;
        console.log('Sheet selected:', selectedSheet);

        // Process the selected sheet
        const worksheet = currentWorkbook.Sheets[selectedSheet];

        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        console.log('Sheet data (first 5 rows):', jsonData.slice(0, 5));

        // Extract tracking numbers from all cells (not just first column)
        const trackingNumbers = [];

        // Process all rows and columns
        for (let i = 0; i < jsonData.length; i++) {
          const row = jsonData[i];
          if (row) {
            // Process all columns in the row
            for (let j = 0; j < row.length; j++) {
              const cellValue = row[j];
              console.log('Processing cell value:', cellValue);

              // Check if it looks like a tracking number
              if (typeof cellValue === 'string') {
                // More flexible pattern matching for Thai Post tracking numbers
                const trackingPattern = /^[A-Z]{2}[0-9]{9}[A-Z]{2}$/i;
                if (trackingPattern.test(cellValue.trim())) {
                  console.log('Found tracking number:', cellValue.trim());
                  trackingNumbers.push(cellValue.trim());
                } else {
                  console.log('Cell value is string but does not match tracking pattern:', cellValue);
                }
              } else if (typeof cellValue === 'number') {
                console.log('Cell value is number:', cellValue);
              } else {
                console.log('Cell value type:', typeof cellValue, 'Value:', cellValue);
              }
            }
          }
        }

        console.log('Final tracking numbers list:', trackingNumbers);

        // Remove duplicates
        const uniqueTrackingNumbers = [...new Set(trackingNumbers)];
        console.log('Unique tracking numbers list:', uniqueTrackingNumbers);

        // Clear workbook data
        currentWorkbook = null;
        currentFilePath = null;
        sheetSelectorWindow = null;

        if (uniqueTrackingNumbers.length === 0) {
          return { success: false, message: 'ไม่พบเลขพัสดุในไฟล์ Excel ที่เลือก' };
        }

        return { success: true, trackingNumbers: uniqueTrackingNumbers };
      } catch (error) {
        console.log('Sheet selection cancelled or failed:', error.message);
        // Clear workbook data
        currentWorkbook = null;
        currentFilePath = null;
        sheetSelectorWindow = null;
        return { success: false, message: 'ยกเลิกการเลือกไฟล์' };
      }
    }
  } catch (error) {
    console.error('Error importing Excel file:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการอ่านไฟล์ Excel: ' + error.message };
  }
});

ipcMain.handle('import-pdf', async () => {
  try {
    console.log('Import PDF function called');
    
    // Show file dialog to select PDF file
    const { canceled, filePaths } = await dialog.showOpenDialog({
      title: 'เลือกไฟล์ PDF',
      filters: [
        { name: 'PDF Files', extensions: ['pdf'] }
      ],
      properties: ['openFile']
    });
    
    console.log('File dialog result:', { canceled, filePaths });
    
    if (canceled || !filePaths || filePaths.length === 0) {
      console.log('File selection canceled or no file selected');
      return { success: false, message: 'ยกเลิกการเลือกไฟล์' };
    }

    const filePath = filePaths[0];
    console.log('Selected PDF file:', filePath);
    
    // Read the PDF file
    const dataBuffer = fs.readFileSync(filePath);
    const pdfData = await pdfParse(dataBuffer);
    
    console.log('PDF loaded, number of pages:', pdfData.numpages);
    console.log('Extracted text content:', pdfData.text.substring(0, 100) + '...');
    
    // Find tracking numbers in the text
    // Pattern for Thai Post tracking numbers like RB032239401TH
    const trackingPattern = /[A-Z]{2}[0-9]{9}[A-Z]{2}/g;
    const matches = pdfData.text.match(trackingPattern);
    
    console.log('Tracking number matches:', matches);
    
    if (!matches || matches.length === 0) {
      return { success: false, message: 'ไม่พบเลขพัสดุในไฟล์ PDF ที่เลือก' };
    }
    
    // Remove duplicates
    const trackingNumbers = [...new Set(matches)];
    
    console.log('Final tracking numbers list:', trackingNumbers);
    
    return { success: true, trackingNumbers };
  } catch (error) {
    console.error('Error importing PDF file:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการอ่านไฟล์ PDF: ' + error.message };
  }
});

ipcMain.handle('save-pdf', async (event, htmlContent) => {
  try {
    // Create a new browser window for PDF generation
    const printWindow = new BrowserWindow({
      show: false,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
      }
    });
    
    // Load HTML content into the window
    printWindow.loadURL('data:text/html;charset=UTF-8,' + encodeURIComponent(htmlContent));
    
    // Wait for the page to load
    await new Promise(resolve => {
      printWindow.webContents.once('did-finish-load', resolve);
    });
    
    // Generate PDF
    const pdfData = await printWindow.webContents.printToPDF({
      marginsType: 0,
      pageSize: 'A4',
      printBackground: true,
      printSelectionOnly: false,
    });
    
    // Close the print window
    printWindow.destroy();
    
    // Show save dialog
    const { canceled, filePath } = await dialog.showSaveDialog({
      title: 'บันทึกพัสดุเป็น PDF',
      defaultPath: `พัสดุไปรษณีย์ไทย_${new Date().toISOString().slice(0, 10)}.pdf`,
      filters: [
        { name: 'PDF Files', extensions: ['pdf'] }
      ]
    });
    
    if (canceled) {
      return { success: false, message: 'ยกเลิกการบันทึก' };
    }
    
    // Save the PDF
    await fs.promises.writeFile(filePath, pdfData);
    return { success: true, message: 'บันทึก PDF สำเร็จ' };
  } catch (error) {
    console.error('Error generating PDF:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการสร้าง PDF: ' + error.message };
  }
});

// Function to convert image to base64
function imageToBase64(imagePath) {
  try {
    const imageBuffer = fs.readFileSync(imagePath);
    return imageBuffer.toString('base64');
  } catch (error) {
    console.error('Error reading image:', error);
    return null;
  }
}

ipcMain.handle('capture-all-parcels', async (event, parcelCount) => {
  try {
    // Get the main window
    const win = BrowserWindow.getFocusedWindow() || BrowserWindow.getAllWindows()[0];
    
    if (!win) {
      throw new Error('ไม่พบหน้าต่างของแอปพลิเคชัน');
    }
    
    // Convert images to base64
    const insertImage = imageToBase64(path.join(__dirname, 'insert.png'));
    const transportImage = imageToBase64(path.join(__dirname, 'transport.png'));
    const shippingImage = imageToBase64(path.join(__dirname, 'shipping.png'));
    const shippingsuccessImage = imageToBase64(path.join(__dirname, 'shippingsuccess.png'));
    
    // Create HTML content that will hold all parcels
    let htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>หน้าจอบันทึกพัสดุ</title>
        <style>
          body {
            font-family: 'Sarabun', 'Prompt', 'Segoe UI', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background: white;
          }
          .page {
            page-break-after: always;
            box-sizing: border-box;
            max-height: 100vh;
            overflow: hidden;
          }
          .parcel-block {
            background: #fff;
            border-radius: 18px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            margin-bottom: 30px;
            padding: 20px;
            border: 1px solid #e0e0e0;
            page-break-inside: avoid;
            break-inside: avoid;
            max-height: 90vh;
            overflow: hidden;
          }
          .parcel-order {
            font-size: 1.1em;
            color: #fff;
            background: linear-gradient(90deg, #009900 60%, #4CAF50 100%);
            display: inline-block;
            border-radius: 50%;
            width: 36px;
            height: 36px;
            text-align: center;
            line-height: 36px;
            font-weight: bold;
            margin-bottom: 10px;
          }
          h1 {
            text-align: center;
            color: #009900;
          }
          h3 {
            color: #333;
            margin-top: 0;
            font-size: 1.3em;
          }
          .summary-box {
            background: #e3fbe3;
            border: 2px solid #4CAF50;
            border-radius: 12px;
            padding: 15px;
            margin: 15px 0;
          }
          .summary-box.delivered {
            background: linear-gradient(90deg, #e8f5e9 60%, #f1f8e9 100%);
            border-color: #388e3c;
          }
          .timeline-bar-inline {
            display: flex;
            align-items: flex-end;
            justify-content: center;
            margin: 20px 0;
            gap: 0;
            flex-wrap: wrap;
            position: relative;
          }
          /* Add line between timeline icons */
          .timeline-bar-inline::before {
            content: '';
            position: absolute;
            top: 20px;
            left: 10%;
            width: 80%;
            height: 4px;
            background: linear-gradient(90deg, #F44336 0%, #FF9800 50%, #4CAF50 100%);
            border-radius: 2px;
            z-index: 1;
          }
          .timeline-step-inline {
            display: flex;
            flex-direction: column;
            align-items: center;
            min-width: 70px;
            position: relative;
            z-index: 2;
            margin: 0 10px; /* Add horizontal spacing */
          }
          .timeline-icon-inline {
            background: #fff;
            border-radius: 50%;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border: 4px solid #F44336;
            width: 40px;
            height: 40px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 8px; /* Increase spacing below icons */
          }
          .timeline-step-inline.completed .timeline-icon-inline {
            border-color: #4CAF50;
          }
          .timeline-label-inline {
            font-weight: bold;
            margin-top: 2px;
            color: #F44336;
            font-size: 0.75em;
            text-align: center;
          }
          .timeline-label-inline.completed {
            color: #009900;
          }
          .status-item {
            padding: 6px;
            border-bottom: 1px solid #eee;
            font-size: 0.85em;
          }
          .status-item:last-child {
            border-bottom: none;
          }
          .status-item.delivered {
            background-color: #e8f5e9;
            border-radius: 5px;
          }
          .status-signature-row {
            display: flex;
            flex-direction: row;
            align-items: flex-start;
            gap: 15px;
            margin-top: 10px;
            page-break-inside: avoid;
            break-inside: avoid;
          }
          .status-list-col {
            flex: 2 1 0;
            min-width: 0;
            max-height: 300px;
            overflow-y: auto;
          }
          .signature-list-col {
            flex: 1 1 0;
            min-width: 0;
            display: flex;
            align-items: flex-start;
            justify-content: flex-end;
          }
          .signature-box {
            min-width: 90px;
            max-width: 150px;
            margin: 0 auto;
            background: #fff;
            border: 2.5px solid #4CAF50;
            border-radius: 12px;
            box-shadow: 0 2px 12px #4caf5022;
            padding: 8px 6px 6px 6px;
            page-break-inside: avoid;
            break-inside: avoid;
          }
          .signature-box img {
            max-width: 150px;
            max-height: 80px;
            margin-top: 4px;
            border: 2px solid #388e3c;
            border-radius: 8px;
            background: #fff;
            box-shadow: 0 1px 8px #0001;
            page-break-inside: avoid;
            break-inside: avoid;
          }
          .signature-title {
            color: #388e3c;
            font-weight: bold;
            margin-bottom: 4px;
            font-size: 0.9em;
            letter-spacing: 0.5px;
          }
          .timeline-icon-inline img {
            width: 24px;
            height: 24px;
          }
          .summary-main {
            font-size: 0.95em;
          }
          .summary-meta {
            font-size: 0.85em;
          }
          .date {
            font-weight: bold;
          }
          .desc {
            margin-top: 4px;
          }
          .meta {
            margin-top: 4px;
            font-size: 0.85em;
          }
        </style>
      </head>
      <body>
    `;
    
    // Get each parcel HTML separately
    for (let i = 0; i < parcelCount; i++) {
      const parcelHtml = await win.webContents.executeJavaScript(`
        (() => {
          const parcelBlocks = document.querySelectorAll('.parcel-block');
          if (parcelBlocks[${i}] && parcelBlocks[${i}].outerHTML) {
            // Clone the element to avoid modifying the original
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = parcelBlocks[${i}].outerHTML;
            
            // Remove buttons
            const buttons = tempDiv.querySelectorAll('.refresh-btn, .print-one-btn');
            buttons.forEach(button => button.remove());
            
            // Replace relative image paths with base64 data
            const images = tempDiv.querySelectorAll('img');
            images.forEach(img => {
              const src = img.src;
              if (src.includes('insert.png')) {
                img.src = 'data:image/png;base64,${insertImage}';
              } else if (src.includes('transport.png')) {
                img.src = 'data:image/png;base64,${transportImage}';
              } else if (src.includes('shipping.png')) {
                img.src = 'data:image/png;base64,${shippingImage}';
              } else if (src.includes('shippingsuccess.png')) {
                img.src = 'data:image/png;base64,${shippingsuccessImage}';
              }
            });
            
            return '<div class="page">' + tempDiv.innerHTML + '</div>';
          } else {
            return '<div class="page"><div>ไม่พบข้อมูลพัสดุ</div></div>';
          }
        })()
      `);
      
      htmlContent += parcelHtml;
    }
    
    htmlContent += `
      </body>
      </html>
    `;
    
    // Create a new browser window for PDF generation
    const printWindow = new BrowserWindow({
      show: false,
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
      }
    });
    
    // Load HTML content into the window
    printWindow.loadURL('data:text/html;charset=UTF-8,' + encodeURIComponent(htmlContent));
    
    // Wait for the page to load
    await new Promise(resolve => {
      printWindow.webContents.once('did-finish-load', resolve);
    });
    
    // Generate PDF
    const pdfData = await printWindow.webContents.printToPDF({
      marginsType: 0,
      pageSize: 'A4',
      printBackground: true,
      printSelectionOnly: false,
    });
    
    // Close the print window
    printWindow.destroy();
    
    // Show save dialog
    const { canceled, filePath } = await dialog.showSaveDialog({
      title: 'บันทึกหน้าจอเป็น PDF',
      defaultPath: `หน้าจอบันทึกพัสดุ_${new Date().toISOString().slice(0, 10)}.pdf`,
      filters: [
        { name: 'PDF Files', extensions: ['pdf'] }
      ]
    });
    
    if (canceled) {
      return { success: false, message: 'ยกเลิกการบันทึก' };
    }
    
    // Save the PDF
    await fs.promises.writeFile(filePath, pdfData);
    return { success: true, message: 'บันทึก PDF สำเร็จ' };
  } catch (error) {
    console.error('Error capturing all parcels:', error);
    return { success: false, message: 'เกิดข้อผิดพลาดในการจับภาพหน้าจอ: ' + error.message };
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});