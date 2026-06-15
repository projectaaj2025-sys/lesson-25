// ============================================
// КОД ДЛЯ GOOGLE APPS SCRIPT
// ============================================
// 1. Откройте https://script.google.com
// 2. Создайте новый проект
// 3. Удалите весь код по умолчанию, вставьте этот
// 4. Разверните → Новое развёртывание → Веб-приложение
// 5. Доступ: "Все" → Выполнить от моего имени
// 6. Скопируйте URL и дайте ученикам
// ============================================

const SHEET_NAME = 'Результаты';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Очистка таблицы
    if (data.action === 'clear') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(SHEET_NAME);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.deleteRows(2, lastRow - 1);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, cleared: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Добавление результата
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Время', 'ФИО', 'Класс', 'Вариант', 
        'Балл', 'Макс', 'Процент', 'Время_выполнения', 'Ошибки'
      ]);
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#2563eb').setFontColor('white');
    }

    sheet.appendRow([
      data.timestamp,
      data.name,
      data.class,
      data.variant,
      data.totalScore,
      data.maxScore,
      data.percentage,
      data.timeSpent,
      data.wrongItems
    ]);

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
