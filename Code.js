function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const action = data.action;

    // 1. Ambil Semua Data
    if (action === 'GET_ALL') {
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) return responseJSON({ success: true, data: [] });
      
      const result = rows.slice(1).map(r => ({
        doNumber: String(r[0]),
        clientName: String(r[1]),
        boxQty: r[2] !== "" ? Number(r[2]) : "Pending",
        scannedQty: Number(r[3]) || 0,
        status: String(r[4] || 'Processing'),
        remark: String(r[5] || ''),
        updatedAt: r[6]
      }));
      return responseJSON({ success: true, data: result });
    }

    // 2. Import Pukal (Tanpa Perlu Jumlah Kotak Awal)
    if (action === 'BULK_IMPORT') {
      const items = data.items;
      const timestamp = new Date();
      items.forEach(item => {
        sheet.appendRow([
          item.doNumber,
          item.clientName,
          "", // Box Qty dibiarkan kosong dahulu
          0,  // Scanned Qty bermula 0
          'Processing',
          item.remark || '',
          timestamp
        ]);
      });
      return responseJSON({ success: true, message: 'Data imported' });
    }

    // 3. Mod Imbasan Dinamik
    if (action === 'UPDATE_SCAN') {
      const barcode = String(data.barcode);
      const mode = data.mode; // 'COUNT', 'FINISH_READY', atau 'HANDOVER'
      const rows = sheet.getDataRange().getValues();
      const timestamp = new Date();

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === barcode) {
          let currentStatus = rows[i][4];
          let currentScanned = Number(rows[i][3]) || 0;

          // MOD 1: Tambah Imbasan Kotak (+1)
          if (mode === 'COUNT') {
            if (currentStatus !== 'Processing') {
              return responseJSON({ success: false, message: `DO [${barcode}] sudah siap/berstatus ${currentStatus}!` });
            }
            let newScanned = currentScanned + 1;
            sheet.getRange(i + 1, 4).setValue(newScanned);
            sheet.getRange(i + 1, 7).setValue(timestamp);

            return responseJSON({ success: true, doNumber: barcode, scanned: newScanned });
          }
          
          // MOD 2: Selesaikan & Konfigurasi Jumlah Akhir Kotak ➔ Ready
          else if (mode === 'FINISH_READY') {
            if (currentStatus !== 'Processing') {
              return responseJSON({ success: false, message: `DO [${barcode}] bukan berstatus Processing!` });
            }
            if (currentScanned === 0) {
              return responseJSON({ success: false, message: `Imbas sekurang-kurangnya 1 kotak dahulu sebelum tukar ke Ready!` });
            }

            // Kunci kuantiti imbasan sebagai Box Qty rasmi
            sheet.getRange(i + 1, 3).setValue(currentScanned); // Column C (Box Qty)
            sheet.getRange(i + 1, 5).setValue('Ready');        // Column E (Status)
            sheet.getRange(i + 1, 7).setValue(timestamp);

            return responseJSON({ success: true, doNumber: barcode, scanned: currentScanned, boxQty: currentScanned, status: 'Ready' });
          }

          // MOD 3: Handover Ke Transporter
          else if (mode === 'HANDOVER') {
            if (currentStatus !== 'Ready') {
              return responseJSON({ success: false, message: `DO [${barcode}] mesti berstatus Ready sebelum Handover!` });
            }
            sheet.getRange(i + 1, 5).setValue('Hand over to transporter');
            sheet.getRange(i + 1, 7).setValue(timestamp);

            return responseJSON({ success: true, doNumber: barcode, status: 'Hand over to transporter' });
          }
        }
      }
      return responseJSON({ success: false, message: `DO [${barcode}] tidak dijumpai dalam Google Sheets!` });
    }

  } catch (error) {
    return responseJSON({ success: false, error: error.toString() });
  }
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}