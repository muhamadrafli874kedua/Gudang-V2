// ============================================================
// STOK GUDANG — Google Apps Script Backend v1.1
// PERBAIKAN: izin UrlFetchApp, token diperbarui, error handling
//
// Sheet: Master_Stok | Barang_Masuk | Barang_Keluar
//
// Master_Stok kolom:
//   A: Kode    B: Nama    C: Stok Awal
//   D: Total Masuk   E: Total Keluar   F: Stok Akhir
//   G: Kadaluarsa    H: Posisi Rak     I: Kategori   J: No. Batch
//
// Barang_Masuk / Barang_Keluar kolom:
//   A: Tanggal   B: Barcode   C: Nama   D: Qty   E: Catatan   F: No. Batch
//
// CARA DEPLOY YANG BENAR:
//   1. Klik Deploy → New deployment
//   2. Type: Web app
//   3. Execute as: Me (your Google account)
//   4. Who has access: Anyone  ← WAJIB agar frontend bisa akses
//   5. Klik Deploy → salin URL → tempel ke index.html di const API = '...'
//   6. Jika update kode: Deploy → Manage deployments → Edit → New version → Deploy
// ============================================================

const SHEET_MASTER = 'Master_Stok'
const SHEET_MASUK  = 'Barang_Masuk'
const SHEET_KELUAR = 'Barang_Keluar'

// ---- ROUTER ----

function doGet(e) {
  const action = e.parameter.action
  try {
    let result
    if      (action === 'ping')        result = { status: 'ok' }
    else if (action === 'search')      result = searchByBarcode(e.parameter.barcode)
    else if (action === 'getAllStock')  result = getAllStock()
    else if (action === 'getHistory')  result = getHistory(e.parameter)
    else if (action === 'getStats')    result = getStats(e.parameter)
    else if (action === 'getSettings') result = getSettings()
    else                               result = { error: 'Unknown action: ' + action }
    return jsonResponse(result)
  } catch (err) {
    return jsonResponse({ error: err.message })
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents)

    // Telegram webhook update
    if (data.update_id !== undefined) {
      _handleTelegramUpdate(data)
      return jsonResponse({ ok: true })
    }

    const action = data.action
    let result
    if      (action === 'stockIn')            result = stockIn(data)
    else if (action === 'stockOut')           result = stockOut(data)
    else if (action === 'updateItem')         result = updateItem(data)
    else if (action === 'addItem')            result = addItem(data)
    else if (action === 'deleteItem')         result = deleteItem(data)
    else if (action === 'saveSettings')       result = saveSettings(data)
    else if (action === 'sendDailyReport')    result = sendDailyReport()
    else if (action === 'sendWelcomeMessage') result = sendWelcomeMessage()
    else                                       result = { error: 'Unknown action: ' + action }
    return jsonResponse(result)
  } catch (err) {
    return jsonResponse({ error: err.message })
  }
}

function _handleTelegramUpdate(update) {
  const props = PropertiesService.getScriptProperties()
  const token = props.getProperty('telegram_bot_token')
  if (!token) return
  const msg = update.message
  if (!msg)  return
  const text = (msg.text || '').trim()
  if (!text.startsWith('/start')) return
  _sendTelegramMessage(token, msg.chat.id,
    '<b>Stok Gudang — Sistem Notifikasi</b>\n\n' +
    'Selamat datang! Anda akan menerima laporan harian otomatis.\n\n' +
    'Laporan dikirim setiap hari pukul 07.00 WIB dan mencakup:\n' +
    '  - Ringkasan total stok\n' +
    '  - Pergerakan barang hari ini\n' +
    '  - Daftar item dengan stok sedikit\n' +
    '  - Daftar item yang mendekati kadaluarsa\n\n' +
    '<i>Stok Gudang  |  by M Rafli</i>'
  )
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON)
}

// ---- HELPER ----

function _makeKey(barcode, batch) {
  const b = (batch || '').toString().trim()
  return b ? `${barcode}|${b}` : barcode.toString().trim()
}

function _sumByKey(sheet) {
  const result = {}
  if (!sheet || sheet.getLastRow() < 2) return result
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues()
    .forEach(r => {
      const bc = r[1].toString().trim()
      if (!bc) return
      const key = _makeKey(bc, r[5])
      result[key] = (result[key] || 0) + (Number(r[3]) || 0)
    })
  return result
}

// ---- SEARCH ----

function searchByBarcode(barcode) {
  if (!barcode) return { found: false, batches: [] }
  const ss    = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(SHEET_MASTER)
  if (!sheet || sheet.getLastRow() < 2) return { found: false, batches: [] }

  const totMasuk  = _sumByKey(ss.getSheetByName(SHEET_MASUK))
  const totKeluar = _sumByKey(ss.getSheetByName(SHEET_KELUAR))

  const batches = []
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues().forEach(r => {
    if (r[0].toString().trim() !== barcode.toString().trim()) return
    const bc    = r[0].toString().trim()
    const batch = r[9] ? r[9].toString().trim() : ''
    const key   = _makeKey(bc, batch)
    const qty   = (Number(r[2]) || 0) + (totMasuk[key] || 0) - (totKeluar[key] || 0)
    batches.push({
      barcode:  bc,
      nama:     r[1].toString(),
      stokAwal: Number(r[2]) || 0,
      qty:      qty.toString(),
      exp:      r[6] ? formatTgl(r[6]) : '',
      posisi:   r[7].toString(),
      kategori: r[8] ? r[8].toString() : '',
      batch:    batch,
    })
  })

  if (batches.length === 0) return { found: false, batches: [] }

  const totalQty = batches.reduce((s, b) => s + Number(b.qty), 0)
  return {
    found: true, batches,
    barcode:  batches[0].barcode,
    nama:     batches[0].nama,
    qty:      totalQty.toString(),
    exp:      batches[0].exp,
    posisi:   batches[0].posisi,
    kategori: batches[0].kategori,
  }
}

// ---- ALL STOCK ----

function getAllStock() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(SHEET_MASTER)
  if (!sheet || sheet.getLastRow() < 2) return { items: [], total: 0 }

  const rows  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues()
  const items = rows
    .filter(r => r[0])
    .map(r => ({
      barcode:  r[0].toString().trim(),
      nama:     r[1].toString(),
      stokAwal: Number(r[2]) || 0,
      qty:      (Number(r[5]) || 0).toString(),
      exp:      r[6] ? formatTgl(r[6]) : '',
      posisi:   r[7].toString(),
      kategori: r[8] ? r[8].toString() : '',
      batch:    r[9] ? r[9].toString().trim() : '',
    }))
  return { items, total: items.length }
}

// ---- HISTORY ----

function getHistory(params) {
  const barcode   = params && params.barcode   ? params.barcode.toString().trim() : ''
  const startDate = params && params.startDate ? new Date(params.startDate) : null
  const endDate   = params && params.endDate
    ? (() => { const d = new Date(params.endDate); d.setHours(23,59,59,999); return d })()
    : null
  const limit = barcode ? 50 : 500

  const ss       = SpreadsheetApp.getActiveSpreadsheet()
  const sheetIn  = ss.getSheetByName(SHEET_MASUK)
  const sheetOut = ss.getSheetByName(SHEET_KELUAR)
  const history  = []

  function rowToObj(r, tipe) {
    const d = r[0] ? new Date(r[0]) : null
    if (startDate && d && d < startDate) return null
    if (endDate   && d && d > endDate)   return null
    if (barcode && r[1].toString().trim() !== barcode) return null
    return {
      tanggal: d ? d.toISOString() : '',
      tipe, barcode: r[1].toString(),
      nama: r[2].toString(), qty: r[3].toString(),
      catatan: r[4].toString(), batch: r[5] ? r[5].toString() : '',
    }
  }

  if (sheetIn && sheetIn.getLastRow() > 1)
    sheetIn.getRange(2, 1, sheetIn.getLastRow() - 1, 6).getValues()
      .forEach(r => { const obj = rowToObj(r, 'MASUK'); if (obj) history.push(obj) })

  if (sheetOut && sheetOut.getLastRow() > 1)
    sheetOut.getRange(2, 1, sheetOut.getLastRow() - 1, 6).getValues()
      .forEach(r => { const obj = rowToObj(r, 'KELUAR'); if (obj) history.push(obj) })

  history.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal))
  return { history: history.slice(0, limit) }
}

// ---- STOCK IN ----

function stockIn(data) {
  const { barcode, nama, qty, exp, posisi, catatan, tanggal, kategori, stokAwal, batch } = data
  const ss      = SpreadsheetApp.getActiveSpreadsheet()
  const sheetIn = ss.getSheetByName(SHEET_MASUK)
  if (!sheetIn) return { success: false, error: 'Sheet Barang_Masuk tidak ditemukan.' }

  if (Number(qty) > 0) {
    sheetIn.appendRow([
      tanggal ? new Date(tanggal) : new Date(),
      barcode, nama || '', Number(qty), catatan || '', batch || ''
    ])
    SpreadsheetApp.flush()
  }

  upsertMaster(barcode, nama, exp, posisi, stokAwal !== undefined ? Number(stokAwal) : undefined, kategori, batch || '')
  return { success: true }
}

// ---- STOCK OUT ----

function stockOut(data) {
  const { barcode, qty, catatan, tanggal, batch } = data
  const ss       = SpreadsheetApp.getActiveSpreadsheet()
  const sheetOut = ss.getSheetByName(SHEET_KELUAR)
  if (!sheetOut) return { success: false, error: 'Sheet Barang_Keluar tidak ditemukan.' }

  const info = searchByBarcode(barcode)
  if (!info.found) return { success: false, error: 'Barang tidak ditemukan.' }

  let targetBatch = (batch || '').trim()
  let itemName    = info.nama

  if (batch) {
    const batchInfo = info.batches.find(b => b.batch === targetBatch)
    if (!batchInfo) return { success: false, error: `Batch "${batch}" tidak ditemukan.` }
    if (Number(batchInfo.qty) < Number(qty))
      return { success: false, error: `Stok batch ini tidak cukup. Tersedia: ${batchInfo.qty} pcs.` }
    itemName = batchInfo.nama
  } else {
    const available = info.batches
      .filter(b => Number(b.qty) > 0)
      .sort((a, b) => {
        if (!a.exp && !b.exp) return 0
        if (!a.exp) return 1
        if (!b.exp) return -1
        return new Date(a.exp) - new Date(b.exp)
      })
    if (available.length === 0) return { success: false, error: 'Stok tidak tersedia.' }
    if (Number(available[0].qty) < Number(qty))
      return { success: false, error: `Stok tidak cukup. Tersedia: ${available[0].qty} pcs (batch ${available[0].batch || 'default'}).` }
    targetBatch = available[0].batch
    itemName    = available[0].nama
  }

  sheetOut.appendRow([
    tanggal ? new Date(tanggal) : new Date(),
    barcode, itemName, Number(qty) || 0, catatan || '', targetBatch
  ])
  SpreadsheetApp.flush()

  upsertMaster(barcode, null, null, null, undefined, undefined, targetBatch)
  return { success: true }
}

// ---- UPDATE / ADD ITEM ----

function updateItem(data) {
  return upsertMaster(
    data.barcode, data.nama, data.exp, data.posisi,
    data.stokAwal !== undefined ? Number(data.stokAwal) : undefined,
    data.kategori, data.batch || ''
  )
}

function addItem(data) {
  return upsertMaster(
    data.barcode, data.nama || 'BARANG BARU', data.exp, data.posisi,
    Number(data.qty) || 0, data.kategori, data.batch || ''
  )
}

// ---- UPSERT MASTER ----

function upsertMaster(barcode, nama, exp, posisi, stokAwal, kategori, batch) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MASTER)
  if (!sheet) return { success: false, error: 'Master_Stok tidak ditemukan.' }

  const batchVal = (batch || '').toString().trim()
  const lastRow  = sheet.getLastRow()

  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues()
    for (let i = 0; i < data.length; i++) {
      const rowBC    = data[i][0].toString().trim()
      const rowBatch = data[i][9] ? data[i][9].toString().trim() : ''
      if (rowBC === barcode.toString().trim() && rowBatch === batchVal) {
        const row = i + 2
        if (nama)                     sheet.getRange(row, 2).setValue(nama)
        if (stokAwal !== undefined)   sheet.getRange(row, 3).setValue(Number(stokAwal) || 0)
        if (exp !== undefined && exp) sheet.getRange(row, 7).setValue(exp)
        if (posisi)                   sheet.getRange(row, 8).setValue(posisi)
        if (kategori !== undefined)   sheet.getRange(row, 9).setValue(kategori || '')
        _updateRowTotals(sheet, barcode.toString().trim(), batchVal, row)
        return { success: true }
      }
    }
  }

  const saVal = stokAwal !== undefined ? Number(stokAwal) : 0
  sheet.appendRow([barcode, nama || 'BARANG BARU', saVal, 0, 0, saVal, exp || '', posisi || '', kategori || '', batchVal])
  _updateRowTotals(sheet, barcode.toString().trim(), batchVal, sheet.getLastRow())
  return { success: true }
}

function _updateRowTotals(sheet, barcode, batch, row) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet()
  const key       = _makeKey(barcode, batch)
  const totMasuk  = _sumByKey(ss.getSheetByName(SHEET_MASUK))
  const totKeluar = _sumByKey(ss.getSheetByName(SHEET_KELUAR))
  const stokAwal  = Number(sheet.getRange(row, 3).getValue()) || 0
  const masuk     = totMasuk[key]  || 0
  const keluar    = totKeluar[key] || 0
  sheet.getRange(row, 4).setValue(masuk)
  sheet.getRange(row, 5).setValue(keluar)
  sheet.getRange(row, 6).setValue(stokAwal + masuk - keluar)
}

function _refreshAllTotalsBatch(ss) {
  const master = ss.getSheetByName(SHEET_MASTER)
  if (!master || master.getLastRow() < 2) return 0

  const totMasuk  = _sumByKey(ss.getSheetByName(SHEET_MASUK))
  const totKeluar = _sumByKey(ss.getSheetByName(SHEET_KELUAR))

  const lastRow = master.getLastRow()
  const data    = master.getRange(2, 1, lastRow - 1, 10).getValues()

  const updates = data.map(r => {
    if (!r[0]) return [0, 0, 0]
    const key    = _makeKey(r[0].toString().trim(), r[9] ? r[9].toString().trim() : '')
    const sa     = Number(r[2]) || 0
    const masuk  = totMasuk[key]  || 0
    const keluar = totKeluar[key] || 0
    return [masuk, keluar, sa + masuk - keluar]
  })

  master.getRange(2, 4, updates.length, 3).setValues(updates)
  return updates.length
}

function refreshAllFormulas() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet()
  const updated = _refreshAllTotalsBatch(ss)
  SpreadsheetApp.getUi().alert(`✅ Total stok diperbarui untuk ${updated} baris.`)
}

// ---- DELETE ----

function deleteItem(data) {
  const { barcode, batch } = data
  if (!barcode) return { success: false, error: 'Barcode diperlukan.' }
  const batchVal = (batch || '').toString().trim()

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MASTER)
  if (!sheet) return { success: false, error: 'Master_Stok tidak ditemukan.' }
  const lastRow = sheet.getLastRow()
  if (lastRow < 2) return { success: false, error: 'Barang tidak ditemukan.' }

  const rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues()
  for (let i = 0; i < rows.length; i++) {
    const rowBC    = rows[i][0].toString().trim()
    const rowBatch = rows[i][9] ? rows[i][9].toString().trim() : ''
    if (rowBC === barcode.toString().trim() && rowBatch === batchVal) {
      sheet.deleteRow(i + 2)
      return { success: true }
    }
  }
  return { success: false, error: 'Barang tidak ditemukan.' }
}

// ---- SETTINGS ----

function getSettings() {
  const props   = PropertiesService.getScriptProperties()
  const expDays = props.getProperty('exp_threshold_days')
  return {
    expDays:          expDays != null ? Number(expDays) : 30,
    telegramBotToken: props.getProperty('telegram_bot_token') || '',
    telegramChatId:   props.getProperty('telegram_chat_id')   || '',
    telegramHour:     Number(props.getProperty('telegram_hour') || 7),
  }
}

// PERBAIKAN: saveSettings hanya simpan data, TIDAK memanggil Telegram
function saveSettings(data) {
  const props = PropertiesService.getScriptProperties()
  if (data.expDays !== undefined && data.expDays !== null && data.expDays !== '') {
    props.setProperty('exp_threshold_days', String(Math.max(1, Math.min(365, Number(data.expDays)))))
  }
  if (data.telegramBotToken !== undefined)
    props.setProperty('telegram_bot_token', data.telegramBotToken.toString().trim())
  if (data.telegramChatId !== undefined)
    props.setProperty('telegram_chat_id', data.telegramChatId.toString().trim())
  if (data.telegramHour !== undefined)
    props.setProperty('telegram_hour', String(Math.max(0, Math.min(23, Number(data.telegramHour)))))
  return { success: true }
}

// ---- TELEGRAM ----

function _sendTelegramMessage(token, chatId, text) {
  // Pastikan GAS sudah diizinkan mengakses URL eksternal:
  // File → Project Settings → centang "Show 'appsscript.json' in editor"
  // Lalu di appsscript.json tambahkan: "oauthScopes": ["https://www.googleapis.com/auth/script.external_request"]
  // Atau cukup jalankan fungsi ini sekali dari editor agar trigger OAuth popup
  try {
    return UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: text, parse_mode: 'HTML' }),
      muteHttpExceptions: true,
    })
  } catch (e) {
    Logger.log('Telegram error: ' + e.message)
    return null
  }
}

function _sendTelegramDocument(token, chatId, filename, content, caption) {
  try {
    var blob    = Utilities.newBlob(content, 'text/plain', filename)
    var payload = { chat_id: String(chatId), document: blob }
    if (caption) payload.caption = caption
    return UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendDocument', {
      method: 'POST', payload: payload, muteHttpExceptions: true,
    })
  } catch (e) {
    Logger.log('Telegram doc error: ' + e.message)
    return null
  }
}

function sendWelcomeMessage() {
  var props  = PropertiesService.getScriptProperties()
  var token  = props.getProperty('telegram_bot_token')
  var chatId = props.getProperty('telegram_chat_id')
  if (!token || !chatId) return { success: false, error: 'Token atau Chat ID belum dikonfigurasi. Simpan dulu pengaturan.' }
  try {
    var resp   = _sendTelegramMessage(token, chatId,
      '<b>Stok Gudang — Sistem Notifikasi</b>\n\n' +
      'Konfigurasi notifikasi berhasil!\n\n' +
      'Laporan dikirim otomatis setiap hari pukul 07.00 WIB.\n\n' +
      '<i>Stok Gudang  |  by M Rafli</i>'
    )
    if (!resp) return { success: false, error: 'Tidak dapat terhubung ke Telegram. Pastikan GAS sudah mendapat izin UrlFetchApp.' }
    var result = JSON.parse(resp.getContentText())
    if (!result.ok) return { success: false, error: result.description || 'Telegram error. Cek token & chat ID.' }
    return { success: true }
  } catch (e) {
    return { success: false, error: e.message }
  }
}

function sendDailyReport() {
  const props  = PropertiesService.getScriptProperties()
  const token  = props.getProperty('telegram_bot_token')
  const chatId = props.getProperty('telegram_chat_id')
  if (!token || !chatId) return { success: false, error: 'Token atau Chat ID belum dikonfigurasi.' }

  try {
    const expDays = Number(props.getProperty('exp_threshold_days') || 30)
    const stats   = getStats({ expDays: expDays })
    const details = _getReportDetails(expDays)
    const now     = new Date()
    const dateStr = Utilities.formatDate(now, 'Asia/Jakarta', 'dd MMMM yyyy')
    const timeStr = Utilities.formatDate(now, 'Asia/Jakarta', 'HH:mm')

    const L = []
    L.push('<b>LAPORAN HARIAN STOK GUDANG</b>')
    L.push(dateStr + '   ' + timeStr + ' WIB')
    L.push('─────────────────────────')
    L.push('')
    L.push('<b>RINGKASAN</b>')
    L.push('Total Item        : <b>' + stats.totalItem   + '</b>')
    L.push('Total Qty Stok    : <b>' + stats.totalStok   + '</b>')
    L.push('Masuk Hari Ini    : <b>' + stats.todayMasuk  + '</b>')
    L.push('Keluar Hari Ini   : <b>' + stats.todayKeluar + '</b>')

    if (stats.lowStock > 0) {
      L.push('')
      L.push('<b>STOK SEDIKIT  (' + stats.lowStock + ' item)</b>')
      details.lowItems.slice(0, 10).forEach(function(it) {
        L.push('  ' + it.nama + '  —  ' + it.qty + ' pcs')
      })
      if (stats.lowStock > 10) L.push('  ... daftar lengkap terlampir')
    }

    if (stats.expiringSoon > 0) {
      L.push('')
      L.push('<b>SEGERA KADALUARSA  (' + stats.expiringSoon + ' item, &lt;= ' + expDays + ' hari)</b>')
      details.expiringItems.slice(0, 10).forEach(function(it) {
        L.push('  ' + it.nama + '  —  ' + it.exp)
      })
      if (stats.expiringSoon > 10) L.push('  ... daftar lengkap terlampir')
    }

    if (stats.expired > 0) {
      L.push('')
      L.push('<b>SUDAH KADALUARSA  (' + stats.expired + ' item)</b>')
      details.expiredItems.slice(0, 10).forEach(function(it) {
        L.push('  ' + it.nama + '  —  ' + it.exp)
      })
      if (stats.expired > 10) L.push('  ... daftar lengkap terlampir')
    }

    if (stats.lowStock === 0 && stats.expiringSoon === 0 && stats.expired === 0) {
      L.push('')
      L.push('✅ Semua stok dalam kondisi aman.')
    }

    L.push('')
    L.push('─────────────────────────')
    L.push('<i>Stok Gudang  |  by M Rafli</i>')

    var resp   = _sendTelegramMessage(token, chatId, L.join('\n'))
    if (!resp) return { success: false, error: 'Tidak dapat terhubung ke Telegram.' }
    var result = JSON.parse(resp.getContentText())
    if (!result.ok) return { success: false, error: result.description || 'Telegram API error' }

    if (details.lowItems.length > 10) {
      var lines = ['STOK SEDIKIT — ' + dateStr, 'Total: ' + details.lowItems.length + ' item', '']
      details.lowItems.forEach(function(it, i) { lines.push((i+1) + '. ' + it.nama + ' — ' + it.qty + ' pcs') })
      Utilities.sleep(300)
      _sendTelegramDocument(token, chatId,
        'stok_sedikit_' + dateStr.replace(/ /g,'_') + '.txt',
        lines.join('\n'), 'Stok Sedikit — Daftar Lengkap')
    }

    if (details.expiringItems.length > 10) {
      var lines2 = ['SEGERA KADALUARSA — ' + dateStr, 'Total: ' + details.expiringItems.length + ' item', '']
      details.expiringItems.forEach(function(it, i) { lines2.push((i+1) + '. ' + it.nama + ' — Exp: ' + it.exp) })
      Utilities.sleep(300)
      _sendTelegramDocument(token, chatId,
        'segera_exp_' + dateStr.replace(/ /g,'_') + '.txt',
        lines2.join('\n'), 'Segera Kadaluarsa — Daftar Lengkap')
    }

    return { success: true }
  } catch (e) {
    return { success: false, error: e.message }
  }
}

function _getReportDetails(expDays) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet()
  const sheet     = ss.getSheetByName(SHEET_MASTER)
  const totMasuk  = _sumByKey(ss.getSheetByName(SHEET_MASUK))
  const totKeluar = _sumByKey(ss.getSheetByName(SHEET_KELUAR))
  const lowItems = [], expiringItems = [], expiredItems = []
  const today = new Date(); today.setHours(0,0,0,0)
  const inExp = new Date(today); inExp.setDate(today.getDate() + expDays)
  const tz    = 'Asia/Jakarta'

  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues()
      .filter(r => r[0])
      .forEach(r => {
        const bc    = r[0].toString().trim()
        const batch = r[9] ? r[9].toString().trim() : ''
        const key   = _makeKey(bc, batch)
        const qty   = (Number(r[2]) || 0) + (totMasuk[key] || 0) - (totKeluar[key] || 0)
        const nama  = r[1].toString()
        if (qty <= 3) lowItems.push({ nama, barcode: bc, qty })
        if (r[6]) {
          const exp = new Date(r[6]); exp.setHours(0,0,0,0)
          const expStr = Utilities.formatDate(exp, tz, 'dd MMM yyyy')
          if (exp < today)       expiredItems.push({ nama, exp: expStr })
          else if (exp <= inExp) expiringItems.push({ nama, exp: expStr })
        }
      })
  }

  lowItems.sort((a, b) => a.qty - b.qty)
  expiringItems.sort((a, b) => new Date(a.exp) - new Date(b.exp))
  return { lowItems, expiringItems, expiredItems }
}

// ============================================================
// SETUP TELEGRAM — Jalankan SEKALI dari GAS editor
// Gunakan token & chat ID baru dari pengaturan
// ============================================================
function setupTelegram() {
  // ← Token dan Chat ID baru (diupdate sesuai info user)
  var token  = '8866238150:AAEKZSGwqWGodSPvOQkg_ADtiNtQcl9DmAU'
  var chatId = '5047350228'
  var hour   = 7

  var props = PropertiesService.getScriptProperties()
  props.setProperty('telegram_bot_token', token)
  props.setProperty('telegram_chat_id',   chatId)
  props.setProperty('telegram_hour',      String(hour))
  props.setProperty('exp_threshold_days', '30')
  Logger.log('[1/4] Config tersimpan.')

  // Set webhook
  var gasUrl = ScriptApp.getService().getUrl()
  var whResp = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/setWebhook', {
    method: 'POST', contentType: 'application/json',
    payload: JSON.stringify({ url: gasUrl }), muteHttpExceptions: true,
  })
  Logger.log('[2/4] Webhook: ' + whResp.getContentText())

  // Hapus trigger lama & buat baru
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendDailyReport') ScriptApp.deleteTrigger(t)
  })
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased().atHour(hour).everyDays(1).inTimezone('Asia/Jakarta').create()
  Logger.log('[3/4] Trigger harian jam ' + hour + ':00 WIB aktif.')

  // Test kirim pesan welcome
  var result = sendWelcomeMessage()
  Logger.log('[4/4] Test kirim: ' + JSON.stringify(result))
  Logger.log('✅ Setup selesai. Cek Telegram kamu.')
}

// ---- STATS ----

function getStats(params) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet()
  const sheet    = ss.getSheetByName(SHEET_MASTER)
  const sheetIn  = ss.getSheetByName(SHEET_MASUK)
  const sheetOut = ss.getSheetByName(SHEET_KELUAR)
  const expDays  = params && params.expDays ? Math.max(1, Number(params.expDays)) : 30

  const totMasuk  = _sumByKey(sheetIn)
  const totKeluar = _sumByKey(sheetOut)

  let totalItem = 0, totalStok = 0, lowStock = 0, expiringSoon = 0, expired = 0
  const today = new Date(); today.setHours(0,0,0,0)
  const in30  = new Date(today); in30.setDate(today.getDate() + expDays)

  if (sheet && sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues().filter(r => r[0])
    totalItem = rows.length
    rows.forEach(r => {
      const bc    = r[0].toString().trim()
      const batch = r[9] ? r[9].toString().trim() : ''
      const key   = _makeKey(bc, batch)
      const qty   = (Number(r[2]) || 0) + (totMasuk[key] || 0) - (totKeluar[key] || 0)
      totalStok += qty
      if (qty <= 3) lowStock++
      if (r[6]) {
        const exp = new Date(r[6]); exp.setHours(0,0,0,0)
        if (exp < today)      expired++
        else if (exp <= in30) expiringSoon++
      }
    })
  }

  return {
    totalItem, totalStok, lowStock, expiringSoon, expired,
    todayMasuk:  _sumToday(sheetIn),
    todayKeluar: _sumToday(sheetOut),
    weeklyChart: _weeklyChart(sheetIn, sheetOut),
  }
}

function _sumToday(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return 0
  const today = new Date(); today.setHours(0,0,0,0)
  let sum = 0
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues().forEach(r => {
    if (!r[0]) return
    const d = new Date(r[0]); d.setHours(0,0,0,0)
    if (d.getTime() === today.getTime()) sum += (Number(r[3]) || 0)
  })
  return sum
}

function _weeklyChart(sheetIn, sheetOut) {
  const tz  = 'Asia/Jakarta'
  const end = new Date(); end.setHours(23,59,59,999)
  const start = new Date(end); start.setDate(end.getDate() - 6); start.setHours(0,0,0,0)

  const days = []
  const cur  = new Date(start)
  while (cur <= end) {
    days.push({ date: Utilities.formatDate(new Date(cur), tz, 'dd/MM'), masuk: 0, keluar: 0 })
    cur.setDate(cur.getDate() + 1)
  }

  function fill(sheet, key) {
    if (!sheet || sheet.getLastRow() < 2) return
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues().forEach(r => {
      if (!r[0]) return
      const d = new Date(r[0]); d.setHours(0,0,0,0)
      if (d < start || d > end) return
      const label = Utilities.formatDate(d, tz, 'dd/MM')
      const slot  = days.find(x => x.date === label)
      if (slot) slot[key] += (Number(r[3]) || 0)
    })
  }
  fill(sheetIn,  'masuk')
  fill(sheetOut, 'keluar')
  return days
}

// ---- UTIL ----

function formatTgl(val) {
  try { return Utilities.formatDate(new Date(val), Session.getScriptTimeZone(), 'yyyy-MM-dd') }
  catch(e) { return val.toString() }
}

// ============================================================
// SETUP TEMPLATE SPREADSHEET
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Stok Gudang')
    .addItem('✅ Format & Perbaiki Semua (Data Aman)', 'perbaikiSemua')
    .addItem('🔄 Refresh Semua Total Stok', 'refreshAllFormulas')
    .addSeparator()
    .addItem('⚠️ Setup Template BARU (Hapus Semua Data)', 'setupSpreadsheet')
    .addToUi()
}

function perbaikiSemua() {
  const ui   = SpreadsheetApp.getUi()
  const resp = ui.alert(
    '✅ Format & Perbaiki Semua',
    'Tindakan ini akan:\n' +
    '1. Terapkan format warna & kolom\n' +
    '2. Isi kolom Batch yang kosong dengan "NO001"\n' +
    '3. Hitung ulang Total Masuk / Keluar / Stok Akhir\n\n' +
    'Data tidak akan dihapus. Lanjutkan?',
    ui.ButtonSet.YES_NO
  )
  if (resp !== ui.Button.YES) return

  const ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.setSpreadsheetTimeZone('Asia/Jakarta')

  _formatMasterOnly(ss)
  _formatTransaksiOnly(ss, SHEET_MASUK,  '#15803D', '#DCFCE7', '#F0FDF4')
  _formatTransaksiOnly(ss, SHEET_KELUAR, '#B91C1C', '#FEE2E2', '#FFF1F2')
  applyConditionalFormatting()

  let count = 0
  function batchFillBatch(sheet, col) {
    if (!sheet || sheet.getLastRow() < 2) return
    const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1)
    const vals  = range.getValues()
    let changed = false
    vals.forEach((r, i) => {
      if (r[0].toString().trim() === '') { vals[i][0] = 'NO001'; count++; changed = true }
    })
    if (changed) range.setValues(vals)
  }

  const master = ss.getSheetByName(SHEET_MASTER)
  batchFillBatch(master, 10)
  batchFillBatch(ss.getSheetByName(SHEET_MASUK),  6)
  batchFillBatch(ss.getSheetByName(SHEET_KELUAR), 6)

  SpreadsheetApp.flush()
  const updated = _refreshAllTotalsBatch(ss)
  ui.alert(`✅ Selesai!\n${count} baris batch diisi "NO001".\n${updated} baris stok diperbarui.`)
}

function applyFormatOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.setSpreadsheetTimeZone('Asia/Jakarta')
  _formatMasterOnly(ss)
  _formatTransaksiOnly(ss, SHEET_MASUK,  '#15803D', '#DCFCE7', '#F0FDF4')
  _formatTransaksiOnly(ss, SHEET_KELUAR, '#B91C1C', '#FEE2E2', '#FFF1F2')
  applyConditionalFormatting()
  SpreadsheetApp.getUi().alert('✅ Format warna berhasil diterapkan.')
}

function _formatMasterOnly(ss) {
  const sheet = ss.getSheetByName(SHEET_MASTER)
  if (!sheet) return
  const headerColor = '#1E3A5F'
  const numCols = 10
  sheet.getRange(1, 1, 1, numCols)
    .setBackground(headerColor).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true,true,true,true,true,true,'#FFFFFF',SpreadsheetApp.BorderStyle.SOLID)
  sheet.setRowHeight(1, 38)
  sheet.setFrozenRows(1)
  sheet.setColumnWidth(1,150); sheet.setColumnWidth(2,220); sheet.setColumnWidth(3,90)
  sheet.setColumnWidth(4,110); sheet.setColumnWidth(5,110); sheet.setColumnWidth(6,100)
  sheet.setColumnWidth(7,120); sheet.setColumnWidth(8,160); sheet.setColumnWidth(9,130)
  sheet.setColumnWidth(10,120)
  sheet.getRange('C:F').setNumberFormat('#,##0')
  sheet.getRange('G:G').setNumberFormat('dd MMM yyyy')
  try {
    sheet.getBandings().forEach(b => b.remove())
    const rows = Math.max(sheet.getLastRow(), 2)
    const banding = sheet.getRange(1, 1, rows, numCols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
    banding.setHeaderRowColor(headerColor)
    banding.setFirstRowColor('#EFF6FF')
    banding.setSecondRowColor('#FFFFFF')
  } catch(e) {}
}

function _formatTransaksiOnly(ss, name, headerColor, rowColor1, rowColor2) {
  const sheet = ss.getSheetByName(name)
  if (!sheet) return
  const numCols = 6
  sheet.getRange(1, 1, 1, numCols)
    .setBackground(headerColor).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
  sheet.setRowHeight(1, 38)
  sheet.setFrozenRows(1)
  const e1 = sheet.getRange(1, 5).getValue().toString().trim()
  if (e1 === 'Keterangan' || e1 === '') sheet.getRange(1, 5).setValue('Catatan')
  sheet.setColumnWidth(1,160); sheet.setColumnWidth(2,150); sheet.setColumnWidth(3,220)
  sheet.setColumnWidth(4,90);  sheet.setColumnWidth(5,200); sheet.setColumnWidth(6,130)
  sheet.getRange('A:A').setNumberFormat('dd MMM yyyy HH:mm')
  sheet.getRange('D:D').setNumberFormat('#,##0')
  try {
    sheet.getBandings().forEach(b => b.remove())
    const rows = Math.max(sheet.getLastRow(), 2)
    const banding = sheet.getRange(1, 1, rows, numCols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
    banding.setHeaderRowColor(headerColor)
    banding.setFirstRowColor(rowColor1)
    banding.setSecondRowColor(rowColor2)
  } catch(e) {}
}

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  ss.setSpreadsheetTimeZone('Asia/Jakarta')
  _setupMaster(ss)
  _setupMasuk(ss)
  _setupKeluar(ss)
  const order = [SHEET_MASTER, SHEET_MASUK, SHEET_KELUAR]
  order.forEach((name, i) => {
    const s = ss.getSheetByName(name)
    if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1) }
  })
  ss.getSheetByName(SHEET_MASTER).activate()
  SpreadsheetApp.getUi().alert('✅ Template berhasil diterapkan!')
}

function _setupMaster(ss) {
  let sheet = ss.getSheetByName(SHEET_MASTER)
  if (!sheet) sheet = ss.insertSheet(SHEET_MASTER)
  sheet.clear(); sheet.clearConditionalFormatRules()
  const headers = ['Kode Barcode','Nama Barang','Stok Awal','Total Masuk','Total Keluar','Stok Akhir','Kadaluarsa','Posisi Rak','Kategori','No. Batch']
  const hRange  = sheet.getRange(1, 1, 1, headers.length)
  hRange.setValues([headers])
       .setBackground('#1E3A5F').setFontColor('#FFFFFF')
       .setFontWeight('bold').setFontSize(11)
       .setHorizontalAlignment('center').setVerticalAlignment('middle')
  sheet.setRowHeight(1, 38); sheet.setFrozenRows(1)
  sheet.setColumnWidth(1,150); sheet.setColumnWidth(2,220); sheet.setColumnWidth(3,90)
  sheet.setColumnWidth(4,110); sheet.setColumnWidth(5,110); sheet.setColumnWidth(6,100)
  sheet.setColumnWidth(7,120); sheet.setColumnWidth(8,160); sheet.setColumnWidth(9,130)
  sheet.setColumnWidth(10,120)
  sheet.getRange('C:F').setNumberFormat('#,##0')
  sheet.getRange('G:G').setNumberFormat('dd MMM yyyy')
  try {
    sheet.getBandings().forEach(b => b.remove())
    const banding = sheet.getRange(1, 1, 500, headers.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
    banding.setHeaderRowColor('#1E3A5F')
    banding.setFirstRowColor('#EFF6FF')
    banding.setSecondRowColor('#FFFFFF')
  } catch(e) {}
  applyConditionalFormatting()
  hRange.setBorder(true,true,true,true,true,true,'#FFFFFF',SpreadsheetApp.BorderStyle.SOLID)
  sheet.getRange('A1:J1').protect().setDescription('Header terkunci').setWarningOnly(true)
}

function applyConditionalFormatting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MASTER)
  if (!sheet) return
  sheet.clearConditionalFormatRules()
  const rules = []
  const maxRow = 500
  const fullRow = `A2:J${maxRow}`

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(3)
    .setBackground('#FEE2E2').setFontColor('#991B1B').setBold(true)
    .setRanges([sheet.getRange(`F2:F${maxRow}`)]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$F2<=3`)
    .setBackground('#FEF2F2')
    .setRanges([sheet.getRange(fullRow)]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(4, 10)
    .setBackground('#FEF9C3').setFontColor('#854D0E')
    .setRanges([sheet.getRange(`F2:F${maxRow}`)]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenDateBefore(SpreadsheetApp.RelativeDate.TODAY)
    .setBackground('#F1F5F9').setFontColor('#94A3B8').setStrikethrough(true)
    .setRanges([sheet.getRange(`G2:G${maxRow}`)]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($G2<>"", $G2<TODAY())`)
    .setBackground('#F8FAFC').setFontColor('#94A3B8')
    .setRanges([sheet.getRange(fullRow)]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND(G2>=TODAY(), G2<=TODAY()+30)`)
    .setBackground('#FFF7ED').setFontColor('#C2410C').setBold(true)
    .setRanges([sheet.getRange(`G2:G${maxRow}`)]).build())
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($G2>=TODAY(), $G2<=TODAY()+30)`)
    .setBackground('#FFFBEB')
    .setRanges([sheet.getRange(fullRow)]).build())
  sheet.setConditionalFormatRules(rules)
}

function _setupSheet(ss, name, headers, headerColor, rowColor1, rowColor2) {
  let sheet = ss.getSheetByName(name)
  if (!sheet) sheet = ss.insertSheet(name)
  sheet.clear()
  const hRange = sheet.getRange(1, 1, 1, headers.length)
  hRange.setValues([headers])
       .setBackground(headerColor).setFontColor('#FFFFFF')
       .setFontWeight('bold').setFontSize(11)
       .setHorizontalAlignment('center').setVerticalAlignment('middle')
  sheet.setRowHeight(1, 38); sheet.setFrozenRows(1)
  sheet.setColumnWidth(1,160); sheet.setColumnWidth(2,150); sheet.setColumnWidth(3,220)
  sheet.setColumnWidth(4,90);  sheet.setColumnWidth(5,200); sheet.setColumnWidth(6,130)
  sheet.getRange('A:A').setNumberFormat('dd MMM yyyy HH:mm')
  sheet.getRange('D:D').setNumberFormat('#,##0')
  try {
    const banding = sheet.getRange(1, 1, 500, headers.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
    banding.setHeaderRowColor(headerColor)
    banding.setFirstRowColor(rowColor1)
    banding.setSecondRowColor(rowColor2)
  } catch(e) {}
}

function _setupMasuk(ss) {
  _setupSheet(ss, SHEET_MASUK,
    ['Tanggal & Waktu','Kode Barcode','Nama Barang','Qty Masuk','Catatan','No. Batch'],
    '#15803D', '#DCFCE7', '#F0FDF4')
}

function _setupKeluar(ss) {
  _setupSheet(ss, SHEET_KELUAR,
    ['Tanggal & Waktu','Kode Barcode','Nama Barang','Qty Keluar','Catatan','No. Batch'],
    '#B91C1C', '#FEE2E2', '#FFF1F2')
}
