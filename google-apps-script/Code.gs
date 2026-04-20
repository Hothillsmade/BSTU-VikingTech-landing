// ═══════════════════════════════════════════════════════
//  VikingTech — Google Apps Script
//  1. Приём заявок с лендинга         (doPost)
//  2. Карточка клиента / смена статуса (doGet)
//  3. Программа лояльности + скидки
//  4. Рекомендации товаров
//  5. Telegram-уведомления менеджеру
//  6. Welcome-email клиенту
//  7. Повторные касания через 7 и 30 дней
//  8. Дашборд с графиками
// ═══════════════════════════════════════════════════════

var SPREADSHEET_ID  = '1SUe1g9YbzNlwRA1MjWOOgjuPuZwhcW05orxkwQVcPmk';
var SHEET_NAME      = 'Заявки';
var DASHBOARD_NAME  = 'Дашборд';
var SECRET_KEY      = 'vikingtech2024';

// ── Telegram: @BotFather → /newbot → токен ──────────────
// chat_id: открой https://api.telegram.org/bot<TOKEN>/getUpdates после /start
var TELEGRAM_TOKEN   = 'ВСТАВИТЬ_TOKEN_БОТА';
var TELEGRAM_CHAT_ID = 'ВСТАВИТЬ_CHAT_ID';

// ─────────────────────────────────────────────────────────
//  Таблица скидок (программа лояльности)
//  [скидка за 2-й заказ, за 3-й, за 4-й+]  %
// ─────────────────────────────────────────────────────────
var DISCOUNT_TABLE = {
  'Холодильник':          [2, 3, 5],
  'Стиральная машина':    [2, 3, 5],
  'Посудомоечная машина': [2, 3, 5],
  'Климат-техника':       [2, 3, 5],
  'Телевизор':            [5, 7, 10],
  'Кухонная техника':     [5, 7, 10],
  'Пылесос':              [10, 15, 20],
  'Другое':               [10, 15, 20]
};

// ─────────────────────────────────────────────────────────
//  Кросс-продажи: что предложить к каждой категории
// ─────────────────────────────────────────────────────────
var RECOMMENDATIONS = {
  'Холодильник':          ['Стиральная машина', 'Посудомоечная машина', 'Кухонная техника'],
  'Стиральная машина':    ['Посудомоечная машина', 'Пылесос', 'Холодильник'],
  'Посудомоечная машина': ['Стиральная машина', 'Холодильник', 'Кухонная техника'],
  'Климат-техника':       ['Пылесос', 'Телевизор', 'Кухонная техника'],
  'Телевизор':            ['Кухонная техника', 'Климат-техника', 'Пылесос'],
  'Кухонная техника':     ['Холодильник', 'Посудомоечная машина', 'Пылесос'],
  'Пылесос':              ['Климат-техника', 'Стиральная машина', 'Кухонная техника'],
  'Другое':               ['Холодильник', 'Стиральная машина', 'Телевизор']
};

// Колонки листа "Заявки"
// A:Дата  B:Имя  C:Телефон  D:Город  E:Техника  F:Комментарий
// G:Статус  H:Менеджер  I:Примечание  J:Источник  K:Карточка
// L:Email  M:Дата закрытия

// ═══════════════════════════════════════════════════════
//  1. doPost — приём заявки с лендинга
// ═══════════════════════════════════════════════════════
function doPost(e) {
  try {
    var data  = JSON.parse(e.postData.contents);
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); setupHeaders(sheet); }
    if (sheet.getLastRow() === 0) setupHeaders(sheet);

    var rawPhone   = (data.phone || '').replace(/\D/g, '');
    var prevOrders = countPreviousOrders(sheet, rawPhone);

    sheet.appendRow([
      data.timestamp || new Date().toLocaleString('ru-RU'), // A
      data.name      || '',                                  // B
      '',                                                    // C — телефон пишем отдельно ниже
      data.city      || '',                                  // D
      data.category  || '',                                  // E
      data.comment   || '',                                  // F
      'Новая',                                               // G
      '',                                                    // H
      '',                                                    // I
      data.source    || 'landing',                           // J
      '',                                                    // K (карточка — ниже)
      data.email     || '',                                  // L
      ''                                                     // M (дата закрытия — ниже)
    ]);

    var lastRow = sheet.getLastRow();
    // Устанавливаем формат текста ДО записи номера — иначе Sheets парсит +375... как формулу
    var phoneCell = sheet.getRange(lastRow, 3);
    phoneCell.setNumberFormat('@');
    phoneCell.setValue(data.phone || '');
    sheet.getRange(lastRow, 1, 1, 13)
         .setBackground(prevOrders > 0 ? '#E3F2FD' : '#FFF9C4');

    // Ссылка на карточку (столбец K)
    var scriptUrl = ScriptApp.getService().getUrl();
    var cardUrl   = scriptUrl + '?action=view&row=' + lastRow + '&secret=' + SECRET_KEY;
    sheet.getRange(lastRow, 11).setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText('📱 Карточка').setLinkUrl(cardUrl).build()
    );

    // Уведомления (ошибки не ломают приём заявки)
    try { sendTelegram(data, lastRow, cardUrl); }       catch(ex) {}
    try { if (data.email) sendWelcomeEmail(data); }     catch(ex) {}
    try { updateDashboard(ss); }                        catch(ex) {}

    return ok('Заявка принята, строка ' + lastRow);
  } catch (err) {
    return ok('Ошибка doPost: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════
//  2. doGet — карточка клиента или смена статуса
// ═══════════════════════════════════════════════════════
function doGet(e) {
  var action = e.parameter.action || '';
  var secret = e.parameter.secret || '';
  var row    = parseInt(e.parameter.row || '0');

  if (secret !== SECRET_KEY) {
    return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;color:#e8350e;padding:40px">⛔ Доступ запрещён</h2>');
  }
  if (!row || row < 2) {
    return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px">⚠️ Неверный номер строки</h2>');
  }

  if (action === 'view')   return showClientCard(row);
  if (action === 'update') return updateClientStatus(row, e.parameter.status || '');

  return ContentService.createTextOutput('VikingTech API — работает! ' + new Date().toLocaleString('ru-RU'));
}

// ═══════════════════════════════════════════════════════
//  3. Карточка клиента
// ═══════════════════════════════════════════════════════
function showClientCard(row) {
  try {
    var sheet  = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    var vals   = sheet.getRange(row, 1, 1, 13).getValues()[0];
    var phone    = (vals[2] || '').toString().replace(/^'/, '');
    var category = (vals[4] || '').toString();
    var rawPhone = phone.replace(/\D/g, '');

    var prevOrders = countPreviousOrders(sheet, rawPhone, row);
    var discount   = getDiscount(prevOrders, category);
    var recs       = getRecommendations(sheet, rawPhone, category, row);

    var data = {
      row:          row,
      timestamp:    vals[0] ? vals[0].toLocaleString ? vals[0].toLocaleString('ru-RU') : String(vals[0]) : '',
      name:         (vals[1] || '').toString(),
      phone:        phone,
      city:         (vals[3] || '').toString(),
      category:     category,
      comment:      (vals[5] || '').toString(),
      status:       (vals[6] || 'Новая').toString(),
      orderNum:     prevOrders + 1,
      discount:     discount > 0 ? discount + '%' : '',
      recs:         recs,     // массив строк — рекомендуемые товары
      scriptUrl:    ScriptApp.getService().getUrl(),
      secret:       SECRET_KEY
    };

    var tpl = HtmlService.createTemplateFromFile('ClientCard');
    tpl.data = data;
    return tpl.evaluate()
      .setTitle('Клиент: ' + data.name)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

  } catch (err) {
    return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px">Ошибка: ' + err.message + '</h2>');
  }
}

// ═══════════════════════════════════════════════════════
//  4. Смена статуса (вызывается через google.script.run)
// ═══════════════════════════════════════════════════════
function updateStatusFromCard(row, status) {
  var allowed = ['В работе', 'Перезвонить', 'Сделка закрыта', 'Отказ'];
  if (!status || allowed.indexOf(status) === -1) throw new Error('Недопустимый статус');

  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  sheet.getRange(row, 7).setValue(status);

  // Сохраняем дату закрытия (для повторных касаний)
  if (status === 'Сделка закрыта') {
    sheet.getRange(row, 13).setValue(new Date().toLocaleDateString('ru-RU'));
  }

  var colors = {
    'В работе':       '#FFF9C4',
    'Перезвонить':    '#FFE0B2',
    'Сделка закрыта': '#C8E6C9',
    'Отказ':          '#FFCDD2'
  };
  sheet.getRange(row, 1, 1, 13).setBackground(colors[status] || '#ffffff');

  try { updateDashboard(SpreadsheetApp.openById(SPREADSHEET_ID)); } catch(ex) {}

  return status;
}

// ═══════════════════════════════════════════════════════
//  5. Программа лояльности
// ═══════════════════════════════════════════════════════
function countPreviousOrders(sheet, rawPhone, excludeRow) {
  if (!rawPhone || sheet.getLastRow() < 2) return 0;
  var phones = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < phones.length; i++) {
    if (excludeRow && (i + 2) === excludeRow) continue;
    if ((phones[i][0] || '').toString().replace(/\D/g, '') === rawPhone) count++;
  }
  return count;
}

function getDiscount(prevOrders, category) {
  if (prevOrders < 1) return 0;
  var tiers = DISCOUNT_TABLE[category] || DISCOUNT_TABLE['Другое'];
  return tiers[Math.min(prevOrders - 1, tiers.length - 1)];
}

// ═══════════════════════════════════════════════════════
//  6. Рекомендации товаров
//  Смотрим все прошлые категории клиента, исключаем то,
//  что уже покупал, и добавляем кросс-продажи
// ═══════════════════════════════════════════════════════
function getRecommendations(sheet, rawPhone, currentCategory, excludeRow) {
  if (!rawPhone || sheet.getLastRow() < 2) return [];

  // Собираем все категории, которые клиент уже покупал
  var allRows  = sheet.getRange(2, 3, sheet.getLastRow() - 1, 3).getValues(); // C-E
  var bought   = {};
  for (var i = 0; i < allRows.length; i++) {
    if (excludeRow && (i + 2) === excludeRow) continue;
    var cellPhone = (allRows[i][0] || '').toString().replace(/\D/g, '');
    if (cellPhone === rawPhone) {
      var cat = (allRows[i][2] || '').toString();
      if (cat) bought[cat] = true;
    }
  }
  bought[currentCategory] = true; // текущую тоже исключаем

  // Собираем рекомендации из всех категорий клиента + текущей
  var recSet = {};
  var allCategories = Object.keys(bought).concat([currentCategory]);
  allCategories.forEach(function(cat) {
    var list = RECOMMENDATIONS[cat] || [];
    list.forEach(function(r) {
      if (!bought[r]) recSet[r] = true;
    });
  });

  return Object.keys(recSet).slice(0, 4); // максимум 4 рекомендации
}

// ═══════════════════════════════════════════════════════
//  7. Telegram-уведомление о новой заявке
// ═══════════════════════════════════════════════════════
function sendTelegram(data, row, cardUrl) {
  if (!TELEGRAM_TOKEN || TELEGRAM_TOKEN === 'ВСТАВИТЬ_TOKEN_БОТА') return;

  var lines = [
    '🔔 *Новая заявка — VikingTech*', '',
    '👤 *Имя:* '    + esc(data.name     || '—'),
    '📞 *Телефон:* ' + esc(data.phone    || '—'),
    '🏙️ *Город:* '  + esc(data.city     || '—'),
    '🔧 *Техника:* ' + esc(data.category || '—')
  ];
  if (data.comment) lines.push('💬 *Комментарий:* ' + esc(data.comment));
  lines.push('', '🕐 ' + (data.timestamp || new Date().toLocaleString('ru-RU')));
  lines.push('[📋 Открыть карточку](' + cardUrl + ')');

  UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_TOKEN + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text: lines.join('\n'), parse_mode: 'Markdown' })
  });
}

// ═══════════════════════════════════════════════════════
//  8. Welcome-email клиенту
// ═══════════════════════════════════════════════════════
function sendWelcomeEmail(data) {
  var html =
    '<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto">' +
    '<div style="background:#1a3a6e;padding:24px 32px"><h1 style="color:#fff;margin:0">Viking<span style="color:#fbbf24">Tech</span></h1></div>' +
    '<div style="padding:32px">' +
    '<h2 style="color:#1a3a6e;margin-top:0">Здравствуйте, ' + (data.name || '') + '!</h2>' +
    '<p>Ваша заявка на консультацию по теме <strong>«' + (data.category || '') + '»</strong> принята.</p>' +
    '<div style="background:#f0f4ff;border-left:4px solid #1a3a6e;padding:16px 20px;border-radius:4px;margin:20px 0">' +
    '⏱️ Менеджер свяжется с вами <strong>в течение 15 минут</strong> (9:00–21:00).</div>' +
    '<p style="margin-bottom:0">С уважением,<br><strong>Команда VikingTech</strong></p></div>' +
    '<div style="background:#f5f7fa;padding:12px 32px;font-size:0.8rem;color:#6b7280">Автоматическое письмо — отвечать не нужно.</div></div>';

  MailApp.sendEmail({ to: data.email, subject: 'Ваша заявка принята — VikingTech', htmlBody: html });
}

// ═══════════════════════════════════════════════════════
//  9. Повторные касания (запускать по триггеру ежедневно)
//     Настройка: выполни setupTriggers() один раз вручную
// ═══════════════════════════════════════════════════════
function checkFollowUps() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;

  var today    = new Date();
  today.setHours(0, 0, 0, 0);
  var rows     = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  var scriptUrl = ScriptApp.getService().getUrl();

  rows.forEach(function(row, idx) {
    var status    = (row[6] || '').toString();
    var closeStr  = (row[12] || '').toString(); // M: Дата закрытия
    if (status !== 'Сделка закрыта' || !closeStr) return;

    var closeDate = new Date(closeStr.split('.').reverse().join('-'));
    closeDate.setHours(0, 0, 0, 0);
    var diffDays = Math.round((today - closeDate) / 86400000);

    if (diffDays !== 7 && diffDays !== 30) return;

    var name     = (row[1] || '').toString();
    var phone    = (row[2] || '').toString().replace(/^'/, '');
    var category = (row[4] || '').toString();
    var email    = (row[11] || '').toString();
    var sheetRow = idx + 2;
    var cardUrl  = scriptUrl + '?action=view&row=' + sheetRow + '&secret=' + SECRET_KEY;
    var discount = getDiscount(countPreviousOrders(sheet, phone.replace(/\D/g,'')), category);
    var recs     = (RECOMMENDATIONS[category] || []).slice(0, 2).join(', ');

    // Telegram менеджеру
    try {
      var label = diffDays === 7 ? '7 дней' : '30 дней';
      var lines = [
        '⏰ *Повторное касание — ' + label + '*', '',
        '👤 ' + esc(name) + '  📞 ' + esc(phone),
        '🔧 Покупка: ' + esc(category),
        discount > 0 ? '🎁 Скидка для клиента: *' + discount + '%*' : '',
        recs ? '💡 Предложить: ' + esc(recs) : '',
        '', '[📋 Открыть карточку](' + cardUrl + ')'
      ].filter(Boolean);
      sendTelegramRaw(lines.join('\n'));
    } catch(ex) {}

    // Email клиенту (если есть)
    if (email) {
      try { sendFollowUpEmail(name, email, category, discount, diffDays); } catch(ex) {}
    }
  });
}

function sendFollowUpEmail(name, email, category, discount, days) {
  var subject = days === 7
    ? 'Как вам техника? Специальное предложение от VikingTech'
    : 'Для вас — персональная скидка от VikingTech';

  var discountBlock = discount > 0
    ? '<div style="background:#f0fdf4;border-left:4px solid #16a34a;padding:16px 20px;border-radius:4px;margin:20px 0">' +
      '🎁 Как постоянному клиенту, вам доступна скидка <strong>' + discount + '%</strong> на следующую покупку.</div>'
    : '';

  var recs = (RECOMMENDATIONS[category] || []).slice(0, 3);
  var recsHtml = recs.length
    ? '<p>Возможно, вас заинтересует: <strong>' + recs.join(', ') + '</strong>.</p>'
    : '';

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto">' +
    '<div style="background:#1a3a6e;padding:24px 32px"><h1 style="color:#fff;margin:0">Viking<span style="color:#fbbf24">Tech</span></h1></div>' +
    '<div style="padding:32px">' +
    '<h2 style="color:#1a3a6e;margin-top:0">Здравствуйте, ' + name + '!</h2>' +
    (days === 7
      ? '<p>Прошла неделя с момента вашей покупки <strong>«' + category + '»</strong>. Надеемся, всё работает отлично!</p>'
      : '<p>Прошёл месяц — самое время подумать о следующей покупке.</p>') +
    discountBlock + recsHtml +
    '<p style="margin-bottom:0">С уважением,<br><strong>Команда VikingTech</strong></p></div>' +
    '<div style="background:#f5f7fa;padding:12px 32px;font-size:0.8rem;color:#6b7280">Автоматическое письмо.</div></div>';

  MailApp.sendEmail({ to: email, subject: subject, htmlBody: html });
}

// Установка триггеров (выполнить один раз вручную из редактора)
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkFollowUps') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkFollowUps').timeBased().atHour(10).everyDays(1).create();
  Logger.log('✅ Триггер установлен: checkFollowUps запускается каждый день в 10:00');
}

// ═══════════════════════════════════════════════════════
//  10. Дашборд
// ═══════════════════════════════════════════════════════
function updateDashboard(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var leadsSheet = ss.getSheetByName(SHEET_NAME);
  if (!leadsSheet || leadsSheet.getLastRow() < 2) return;

  var dash = ss.getSheetByName(DASHBOARD_NAME);
  if (!dash) dash = ss.insertSheet(DASHBOARD_NAME);
  dash.clearContents();
  dash.clearFormats();

  var rows = leadsSheet.getRange(2, 1, leadsSheet.getLastRow() - 1, 13).getValues();
  var now  = new Date();

  // ── Считаем метрики ──────────────────────────────────
  var total     = rows.length;
  var week7     = 0, month30 = 0, closed = 0, rejected = 0;
  var statusMap = {}, phoneMap = {}, weekMap = {};

  rows.forEach(function(r) {
    var dateVal  = r[0];
    var status   = (r[6] || 'Новая').toString();
    var rawPhone = (r[2] || '').toString().replace(/\D/g, '');

    // Дата заявки
    var d = dateVal instanceof Date ? dateVal : new Date(String(dateVal).split('.').reverse().join('-'));
    var diffDays = isNaN(d) ? 999 : Math.round((now - d) / 86400000);
    if (diffDays <= 7)  week7++;
    if (diffDays <= 30) month30++;

    // Статусы
    statusMap[status] = (statusMap[status] || 0) + 1;
    if (status === 'Сделка закрыта') closed++;
    if (status === 'Отказ')          rejected++;

    // Уникальные номера (для лояльности)
    if (rawPhone) phoneMap[rawPhone] = (phoneMap[rawPhone] || 0) + 1;

    // Заявки по неделям (ISO-неделя)
    if (!isNaN(d)) {
      var weekKey = getWeekLabel(d);
      weekMap[weekKey] = (weekMap[weekKey] || 0) + 1;
    }
  });

  var uniquePhones  = Object.keys(phoneMap).length;
  var repeatClients = Object.keys(phoneMap).filter(function(p) { return phoneMap[p] > 1; }).length;
  var loyaltyPct    = uniquePhones > 0 ? Math.round(repeatClients / uniquePhones * 100) : 0;
  var convPct       = total > 0 ? Math.round(closed / total * 100) : 0;

  // ── Шапка дашборда ───────────────────────────────────
  dash.getRange('A1').setValue('VikingTech — Дашборд интернет-маркетинга');
  dash.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#1a3a6e');
  dash.getRange('A2').setValue('Обновлено: ' + now.toLocaleString('ru-RU'))
      .setFontColor('#6b7280').setFontSize(10);

  // ── Ключевые метрики ─────────────────────────────────
  var metrics = [
    ['КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ', ''],
    ['Всего заявок',           total],
    ['За последние 7 дней',    week7],
    ['За последние 30 дней',   month30],
    ['Сделок закрыто',         closed],
    ['Конверсия',              convPct + '%'],
    ['Повторных клиентов',     loyaltyPct + '%'],
    ['Отказов',                rejected]
  ];

  dash.getRange(4, 1, metrics.length, 2).setValues(metrics);
  dash.getRange(4, 1, 1, 2).setBackground('#1a3a6e').setFontColor('#fff').setFontWeight('bold');
  dash.getRange(5, 2, metrics.length - 1, 1).setFontWeight('bold').setFontSize(13);
  dash.getRange(5, 1, metrics.length - 1, 2).setBackground('#f5f7fa');
  dash.setColumnWidth(1, 200); dash.setColumnWidth(2, 120);

  // ── Данные: заявки по неделям (последние 8 недель) ───
  var weekRow = 14;
  dash.getRange(weekRow, 1).setValue('ЗАЯВКИ ПО НЕДЕЛЯМ').setFontWeight('bold').setFontColor('#1a3a6e');
  dash.getRange(weekRow + 1, 1).setValue('Неделя');
  dash.getRange(weekRow + 1, 2).setValue('Заявки');
  dash.getRange(weekRow + 1, 1, 1, 2).setBackground('#1a3a6e').setFontColor('#fff').setFontWeight('bold');

  var weekKeys = Object.keys(weekMap).sort().slice(-8);
  weekKeys.forEach(function(k, i) {
    dash.getRange(weekRow + 2 + i, 1).setValue(k);
    dash.getRange(weekRow + 2 + i, 2).setValue(weekMap[k]);
  });
  var weekDataRange = dash.getRange(weekRow + 2, 1, Math.max(weekKeys.length, 1), 2);

  // ── Данные: статусы ──────────────────────────────────
  var statRow = weekRow + 2 + weekKeys.length + 2;
  dash.getRange(statRow, 1).setValue('СТАТУСЫ ЗАЯВОК').setFontWeight('bold').setFontColor('#1a3a6e');
  dash.getRange(statRow + 1, 1).setValue('Статус');
  dash.getRange(statRow + 1, 2).setValue('Количество');
  dash.getRange(statRow + 1, 1, 1, 2).setBackground('#1a3a6e').setFontColor('#fff').setFontWeight('bold');

  var statuses = ['Новая', 'В работе', 'Перезвонить', 'Сделка закрыта', 'Отказ'];
  statuses.forEach(function(s, i) {
    dash.getRange(statRow + 2 + i, 1).setValue(s);
    dash.getRange(statRow + 2 + i, 2).setValue(statusMap[s] || 0);
  });
  var statDataRange = dash.getRange(statRow + 1, 1, statuses.length + 1, 2);

  // ── Удаляем старые графики, создаём новые ────────────
  dash.getCharts().forEach(function(c) { dash.removeChart(c); });

  if (weekKeys.length > 0) {
    var barChart = dash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dash.getRange(weekRow + 1, 1, weekKeys.length + 1, 2))
      .setOption('title', 'Заявки по неделям')
      .setOption('colors', ['#1a3a6e'])
      .setOption('legend', { position: 'none' })
      .setPosition(4, 4, 10, 10)
      .build();
    dash.insertChart(barChart);
  }

  var pieChart = dash.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(statDataRange)
    .setOption('title', 'Распределение по статусам')
    .setOption('colors', ['#FFF9C4', '#FFF9C4', '#FFE0B2', '#C8E6C9', '#FFCDD2'])
    .setOption('pieHole', 0.4)
    .setPosition(statRow, 4, 10, 10)
    .build();
  dash.insertChart(pieChart);

  // Активируем лист заявок обратно
  ss.setActiveSheet(leadsSheet);
}

function getWeekLabel(date) {
  var d   = new Date(date);
  var day = d.getDay() || 7;
  d.setDate(d.getDate() + 4 - day);
  var year    = d.getFullYear();
  var week    = Math.ceil(((d - new Date(year, 0, 1)) / 86400000 + 1) / 7);
  return year + '-W' + (week < 10 ? '0' : '') + week;
}

// ═══════════════════════════════════════════════════════
//  Заголовки листа "Заявки"
// ═══════════════════════════════════════════════════════
function setupHeaders(sheet) {
  var headers = [
    'Дата и время', 'Имя клиента', 'Телефон', 'Город', 'Тип техники',
    'Комментарий', 'Статус', 'Менеджер', 'Примечание', 'Источник',
    'Карточка', 'Email', 'Дата закрытия'
  ];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length)
       .setBackground('#1a3a6e').setFontColor('#fff').setFontWeight('bold').setFontSize(11);
  sheet.setFrozenRows(1);

  [150,150,160,120,170,250,130,130,200,100,110,180,140]
    .forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

  sheet.getRange('C2:C1000').setNumberFormat('@');

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Новая','В работе','Перезвонить','Сделка закрыта','Отказ'], true)
    .setAllowInvalid(false).build();
  sheet.getRange('G2:G1000').setDataValidation(rule);
}

// ═══════════════════════════════════════════════════════
//  Вспомогательные
// ═══════════════════════════════════════════════════════
function ok(msg) {
  return ContentService.createTextOutput(JSON.stringify({status:'ok',message:msg}))
    .setMimeType(ContentService.MimeType.JSON);
}

function esc(text) {
  return String(text).replace(/[_*[\]()~`>#+=|{}.!-]/g, '\\$&');
}

function sendTelegramRaw(text) {
  if (!TELEGRAM_TOKEN || TELEGRAM_TOKEN === 'ВСТАВИТЬ_TOKEN_БОТА') return;
  UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_TOKEN + '/sendMessage', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text: text, parse_mode: 'Markdown' })
  });
}

// ═══════════════════════════════════════════════════════
//  Тест: запусти из редактора скриптов
// ═══════════════════════════════════════════════════════
function testInsert() {
  doPost({ postData: { contents: JSON.stringify({
    timestamp: new Date().toLocaleString('ru-RU'),
    name: 'Тест Тестов', phone: '+375 (29) 123-45-67',
    city: 'Минск', category: 'Холодильник',
    comment: 'Тест', source: 'test', email: ''
  })}});
  Logger.log('Готово');
}
