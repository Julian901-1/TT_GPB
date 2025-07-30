// ================== НАСТРОЙКИ СКРИПТА ==================
const CONFIG = {
  // ID таблиц
  REPORT_TT_BUILDER_ID: '1YBri89xtf_uvECGRUhMoxh0i9nREuXzLg04hpy08jj0',
  MP_STATUS_ID: '1sHqFzXKlLW-3PKglNNxV4RvS5GXsXzEPq34QZ2tetjE',
  
  // Настройки пользователя
  TARGET_MONTH: 'Август',  // Месяц для обработки
  TARGET_PRODUCT: 'ДК',    // Продукт: 'ДК' или 'ПДС'
  
  // Общие вкладки в REPORT_TT_BUILDER
  COMMON_SHEETS: [
    'Общие требования от 23.10.24',
    'ПРОДУКТ'
  ],
  
  // Столбцы в MP_STATUS
  MP_COLUMNS: {
    PRODUCT: 'B',
    PLATFORM: 'D', 
    FORMAT: 'E',
    DATE: 'G'
  },
  
  // Блоки размещений
  PLACEMENT_BLOCKS: {
    'OLV': ['Яндекс Видео', 'Яндекс.Видео'],
    'Тематические сайты': ['Avito', 'Авто.ру', 'Коммерсант', 'РБК'],
    'Баннерная реклама': ['Яндекс Медийные баннеры', 'Яндекс.Медийные баннеры'],
    'Социальные сети': ['VK Ads'],
    'Фин. агрегаторы': ['Sravni.ru', 'Bankiros', 'VBR']
  }
};

// ================== ОСНОВНАЯ ФУНКЦИЯ ==================
function buildReport() {
  try {
    console.log('Начинаем сборку отчета...');
    
    // 1. Получаем данные из MP_STATUS
    const mpData = getMPStatusData();
    console.log('Данные из MP_STATUS получены:', mpData);
    
    // 2. Находим подходящие вкладки в REPORT_TT_BUILDER
    const matchingSheets = findMatchingSheets(mpData);
    console.log('Найденные вкладки:', matchingSheets.map(m => ({
      originalName: m.originalName,
      newName: m.newName,
      platform: m.placement.platform
    })));
    
    // 3. Создаем копию REPORT_TT_BUILDER
    const newSpreadsheet = createReportCopy(matchingSheets);
    console.log('Создана копия отчета:', newSpreadsheet.getUrl());
    
    // 4. Обновляем лист ПРОДУКТ
    updateProductSheet(newSpreadsheet, mpData);
    
    console.log('Отчет успешно создан!');
    return newSpreadsheet.getUrl();
    
  } catch (error) {
    console.error('Ошибка при создании отчета:', error);
    throw error;
  }
}

// ================== ПОЛУЧЕНИЕ ДАННЫХ ==================
function getMPStatusData() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.MP_STATUS_ID);
  const sheet = spreadsheet.getSheetByName(CONFIG.TARGET_MONTH);
  
  if (!sheet) {
    throw new Error(`Вкладка "${CONFIG.TARGET_MONTH}" не найдена в MP_STATUS`);
  }
  
  const data = sheet.getDataRange().getValues();
  const placements = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const product = row[1]; // Столбец B
    const platform = row[3]; // Столбец D
    const format = row[4]; // Столбец E
    const date = row[6]; // Столбец G
    
    if (product === CONFIG.TARGET_PRODUCT && platform && format) {
      placements.push({
        platform: platform.toString().trim(),
        format: format.toString().trim(),
        date: date,
        row: i + 1
      });
    }
  }
  
  return placements;
}

// ================== ПОИСК ПОДХОДЯЩИХ ВКЛАДОК ==================
function findMatchingSheets(mpData) {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.REPORT_TT_BUILDER_ID);
  const allSheets = spreadsheet.getSheets();
  const matchingSheets = [];
  
  for (const placement of mpData) {
    const targetNames = [
      `${placement.platform} (${placement.format})`,
      placement.platform
    ];
    
    let bestMatch = null;
    let bestScore = 0;
    
    for (const sheet of allSheets) {
      const sheetName = sheet.getName();
      
      for (const targetName of targetNames) {
        const score = calculateSimilarity(sheetName, targetName);
        if (score > bestScore && score > 0.7) {
          bestScore = score;
          bestMatch = {
            sheet: sheet,
            originalName: sheetName,
            newName: placement.platform,
            placement: placement
          };
        }
      }
    }
    
    if (bestMatch && !matchingSheets.find(m => m.originalName === bestMatch.originalName)) {
      matchingSheets.push(bestMatch);
    }
  }
  
  return matchingSheets;
}

// ================== РАСЧЕТ СХОЖЕСТИ СТРОК ==================
function calculateSimilarity(str1, str2) {
  const s1 = str1.toLowerCase().replace(/[.\s]/g, '');
  const s2 = str2.toLowerCase().replace(/[.\s]/g, '');
  
  if (s1 === s2) return 1;
  if (s1.includes(s2) || s2.includes(s1)) return 0.9;
  
  const longer = s1.length > s2.length ? s1 : s2;
  const shorter = s1.length > s2.length ? s2 : s1;
  
  if (longer.length === 0) return 1;
  
  const distance = levenshteinDistance(longer, shorter);
  return (longer.length - distance) / longer.length;
}

function levenshteinDistance(str1, str2) {
  const matrix = [];
  
  for (let i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }
  
  for (let j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }
  
  for (let i = 1; i <= str2.length; i++) {
    for (let j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }
  
  return matrix[str2.length][str1.length];
}

// ================== СОЗДАНИЕ КОПИИ ОТЧЕТА ==================
function createReportCopy(matchingSheets) {
  const sourceSpreadsheet = SpreadsheetApp.openById(CONFIG.REPORT_TT_BUILDER_ID);
  const newSpreadsheet = sourceSpreadsheet.copy(`Отчет ${CONFIG.TARGET_PRODUCT} - ${CONFIG.TARGET_MONTH}`);
  
  // Получаем все листы новой таблицы
  const allSheets = newSpreadsheet.getSheets();
  const sheetsToKeep = new Set();
  const sheetsToRename = new Map();
  
  console.log('Все листы в копии:', allSheets.map(s => s.getName()));
  
  // Добавляем общие листы
  for (const commonSheetName of CONFIG.COMMON_SHEETS) {
    const sheet = newSpreadsheet.getSheetByName(commonSheetName);
    if (sheet) {
      sheetsToKeep.add(sheet.getName());
      if (commonSheetName === 'ПРОДУКТ') {
        sheetsToRename.set(sheet.getName(), CONFIG.TARGET_PRODUCT);
      }
      console.log('Оставляем общий лист:', sheet.getName());
    }
  }
  
  // Добавляем найденные листы
  for (const matchInfo of matchingSheets) {
    const sheet = newSpreadsheet.getSheetByName(matchInfo.originalName);
    if (sheet) {
      sheetsToKeep.add(sheet.getName());
      sheetsToRename.set(sheet.getName(), matchInfo.newName);
      console.log('Оставляем найденный лист:', sheet.getName(), '-> будет переименован в:', matchInfo.newName);
    }
  }
  
  console.log('Листы к сохранению:', Array.from(sheetsToKeep));
  console.log('Листы к переименованию:', Array.from(sheetsToRename.entries()));
  
  // Удаляем ненужные листы (но не все!)
  const sheetsToDelete = [];
  for (const sheet of allSheets) {
    if (!sheetsToKeep.has(sheet.getName())) {
      sheetsToDelete.push(sheet);
    }
  }
  
  console.log('Листы к удалению:', sheetsToDelete.map(s => s.getName()));
  
  // Проверяем, что мы не удаляем все листы
  if (sheetsToDelete.length >= allSheets.length) {
    throw new Error('Попытка удалить все листы! Проверьте логику поиска листов.');
  }
  
  // Удаляем ненужные листы
  for (const sheet of sheetsToDelete) {
    try {
      console.log('Удаляем лист:', sheet.getName());
      newSpreadsheet.deleteSheet(sheet);
    } catch (error) {
      console.error('Ошибка при удалении листа', sheet.getName(), ':', error);
    }
  }
  
  // Переименовываем листы
  for (const [oldName, newName] of sheetsToRename.entries()) {
    const sheet = newSpreadsheet.getSheetByName(oldName);
    if (sheet && oldName !== newName) {
      console.log('Переименовываем лист:', oldName, '->', newName);
      sheet.setName(newName);
    }
  }
  
  return newSpreadsheet;
}

// ================== ОБНОВЛЕНИЕ ЛИСТА ПРОДУКТ ==================
function updateProductSheet(newSpreadsheet, mpData) {
  const productSheet = newSpreadsheet.getSheetByName(CONFIG.TARGET_PRODUCT);
  if (!productSheet) {
    throw new Error(`Лист ${CONFIG.TARGET_PRODUCT} не найден`);
  }
  
  // Очищаем данные начиная с 3 строки
  const lastRow = productSheet.getLastRow();
  if (lastRow >= 3) {
    productSheet.getRange(3, 1, lastRow - 2, productSheet.getLastColumn()).clearContent();
  }
  
  // Группируем размещения по блокам
  const groupedPlacements = groupPlacementsByBlocks(mpData);
  console.log('Группировка по блокам:', Object.keys(groupedPlacements));
  
  let currentRow = 3;
  let placementNumber = 1;
  
  for (const [blockName, placements] of Object.entries(groupedPlacements)) {
    console.log(`Обрабатываем блок: ${blockName}, размещений: ${placements.length}`);
    
    // Добавляем заголовок блока
    productSheet.getRange(currentRow, 1).setValue(blockName);
    currentRow++;
    
    // Добавляем размещения в блоке
    for (const placement of placements) {
      const mpStatusData = getMPStatusPlacementData(placement);
      const comment = generateComment(placement);
      const hyperlink = generateHyperlink(newSpreadsheet, placement.platform);
      
      console.log(`Добавляем размещение: ${placement.platform}`);
      
      // A - номер размещения
      productSheet.getRange(currentRow, 1).setValue(placementNumber);
      
      // B - название площадки
      productSheet.getRange(currentRow, 2).setValue(placement.platform);
      
      // C - гиперссылка на вкладку
      productSheet.getRange(currentRow, 3).setFormula(hyperlink);
      
      // D - данные из столбца E в MP_STATUS
      productSheet.getRange(currentRow, 4).setValue(mpStatusData.format);
      
      // E - данные из столбца G в MP_STATUS
      productSheet.getRange(currentRow, 5).setValue(mpStatusData.date);
      
      // F - комментарий
      productSheet.getRange(currentRow, 6).setValue(comment);
      
      // G - дата минус 2 рабочие недели
      const deadlineDate = subtractWorkingDays(mpStatusData.date, 14);
      productSheet.getRange(currentRow, 7).setValue(deadlineDate);
      
      placementNumber++;
      currentRow++;
    }
  }
  
  console.log('Лист ПРОДУКТ обновлен успешно');
}

// ================== ГРУППИРОВКА ПО БЛОКАМ ==================
function groupPlacementsByBlocks(mpData) {
  const grouped = {};
  
  for (const placement of mpData) {
    let blockFound = false;
    
    for (const [blockName, platforms] of Object.entries(CONFIG.PLACEMENT_BLOCKS)) {
      for (const platform of platforms) {
        if (calculateSimilarity(placement.platform, platform) > 0.8) {
          if (!grouped[blockName]) {
            grouped[blockName] = [];
          }
          grouped[blockName].push(placement);
          blockFound = true;
          console.log(`Размещение ${placement.platform} добавлено в блок ${blockName}`);
          break;
        }
      }
      if (blockFound) break;
    }
    
    if (!blockFound) {
      // Если блок не найден, создаем новый
      const newBlockName = 'Прочие';
      if (!grouped[newBlockName]) {
        grouped[newBlockName] = [];
      }
      grouped[newBlockName].push(placement);
      console.log(`Размещение ${placement.platform} добавлено в блок ${newBlockName}`);
    }
  }
  
  return grouped;
}

// ================== ПОЛУЧЕНИЕ ДАННЫХ РАЗМЕЩЕНИЯ ==================
function getMPStatusPlacementData(placement) {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.MP_STATUS_ID);
  const sheet = spreadsheet.getSheetByName(CONFIG.TARGET_MONTH);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === CONFIG.TARGET_PRODUCT && 
        row[3] === placement.platform && 
        row[4] === placement.format) {
      return {
        format: row[4],
        date: row[6]
      };
    }
  }
  
  return { format: placement.format, date: placement.date };
}

// ================== ГЕНЕРАЦИЯ КОММЕНТАРИЯ ==================
function generateComment(placement) {
  const previousMonth = getPreviousMonth(CONFIG.TARGET_MONTH);
  const wasPreviouslyLaunched = checkPreviousLaunch(placement, previousMonth);
  
  if (wasPreviouslyLaunched) {
    const mpStatusData = getMPStatusPlacementData(placement);
    const deadlineDate = subtractWorkingDays(mpStatusData.date, 14);
    const deadlineDateStr = Utilities.formatDate(deadlineDate, Session.getScriptTimeZone(), 'dd.MM');
    
    return `В наличии будут креативы ${previousMonth.toLowerCase()}.В случае замены просим предоставить апдейт до ${deadlineDateStr}
В случае пролонгации на ${CONFIG.TARGET_MONTH.toLowerCase()} - сообщите дополнительно.`;
  } else {
    return 'Ожидаем креативы';
  }
}

// ================== ПРОВЕРКА ПРЕДЫДУЩЕГО ЗАПУСКА ==================
function checkPreviousLaunch(placement, previousMonth) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.MP_STATUS_ID);
    const sheet = spreadsheet.getSheetByName(previousMonth);
    
    if (!sheet) return false;
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === CONFIG.TARGET_PRODUCT && 
          row[3] === placement.platform && 
          row[4] === placement.format) {
        return true;
      }
    }
    
    return false;
  } catch (error) {
    console.log(`Не удалось проверить предыдущий месяц ${previousMonth}:`, error);
    return false;
  }
}

// ================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==================
function getPreviousMonth(currentMonth) {
  const months = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 
                  'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  
  const currentIndex = months.indexOf(currentMonth);
  if (currentIndex === -1) return 'Июль'; // По умолчанию
  
  return currentIndex === 0 ? months[11] : months[currentIndex - 1];
}

function subtractWorkingDays(date, days) {
  const result = new Date(date);
  let workingDaysLeft = days;
  
  while (workingDaysLeft > 0) {
    result.setDate(result.getDate() - 1);
    const dayOfWeek = result.getDay();
    
    // Пропускаем выходные (суббота = 6, воскресенье = 0)
    if (dayOfWeek !== 0 && dayOfWeek !== 6) {
      workingDaysLeft--;
    }
  }
  
  return result;
}

function generateHyperlink(spreadsheet, platformName) {
  const sheet = spreadsheet.getSheetByName(platformName);
  if (sheet) {
    const gid = sheet.getSheetId();
    return `=ГИПЕРССЫЛКА("#gid=${gid}";"${platformName}")`;
  }
  return platformName;
}

// ================== ФУНКЦИЯ ДЛЯ ТЕСТИРОВАНИЯ ==================
function testScript() {
  console.log('Тестирование скрипта...');
  console.log('Настройки:', CONFIG);
  
  try {
    const url = buildReport();
    console.log('Отчет создан успешно:', url);
  } catch (error) {
    console.error('Ошибка тестирования:', error);
  }
}