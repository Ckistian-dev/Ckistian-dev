function importLatestDataFromDrive() {
  // ID da pasta e planilha
  var folderId = '1-412sTH7LKXgcTFLjNMhvUUYYeJbuUO3';
  var spreadsheetId = '1vITmZvs-JskK3xVyIaO4HtfS5rZZ-CsGL055U8Qq3sM';
  var sheetName = 'Dados';
  
  // Importar dados CSV com as modificações nas colunas
  importLatestCSVFromDrive(folderId, spreadsheetId, sheetName);
  
  // Importar dados XLS/XLSX sem modificar as colunas
  importLatestXLSFromDrive(folderId, spreadsheetId, sheetName);
  
  // Remover linhas com a palavra "Devolvido"
  removeRowsContainingWord(spreadsheetId, sheetName, 'Devolvido');
  
  // Remover duplicatas
  removeDuplicateRows(spreadsheetId, sheetName);
}

function importLatestCSVFromDrive(folderId, spreadsheetId, sheetName) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.CSV);
  
  if (!files.hasNext()) {
    Logger.log('Nenhum arquivo CSV encontrado');
    return;
  }
  
  var latestFile = findLatestFile(files);
  
  var csvContent = latestFile.getBlob().getDataAsString();
  var csvData = parseCsvWithDelimiter(csvContent, ';');

  // Remove a primeira linha do CSV (cabeçalho)
  csvData.shift();

  // Colunas a serem removidas (somente para o CSV)
  var columnsToRemove = [0, 1, 4, 5, 6, 7, 10, 11, 13, 14, 16];
  var cleanedData = cleanData(csvData, columnsToRemove);
  
  // Adiciona coluna em branco entre E e F (índice 4 e 5)
  cleanedData = cleanedData.map(function(row) {
    var newRow = [];
    for (var i = 0; i < row.length; i++) {
      newRow.push(row[i]);
      if (i === 4) newRow.push(''); // Adiciona coluna em branco
    }
    return newRow;
  });
  
  updateSpreadsheet(spreadsheetId, sheetName, cleanedData, true);
  Logger.log('Dados CSV importados: ' + latestFile.getName());
}

function importLatestXLSFromDrive(folderId, spreadsheetId, sheetName) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  
  if (!files.hasNext()) {
    Logger.log('Nenhum arquivo XLS/XLSX encontrado');
    return;
  }
  
  var latestFile = findLatestXlsFile(files);
  if (!latestFile) return;

  var tempFile = Drive.Files.copy({
    title: latestFile.getName().replace('.xls', ''),
    mimeType: MimeType.GOOGLE_SHEETS
  }, latestFile.getId());
  
  var tempSpreadsheetId = tempFile.id;
  var newSheetData = SpreadsheetApp.openById(tempSpreadsheetId).getSheets()[0].getDataRange().getValues();
  
  newSheetData.shift(); // Remove a primeira linha dos dados importados
  
  updateSpreadsheet(spreadsheetId, sheetName, newSheetData, false);
  Logger.log('Dados XLS importados: ' + latestFile.getName());
  
  DriveApp.getFileById(tempSpreadsheetId).setTrashed(true);
}

function removeRowsContainingWord(spreadsheetId, sheetName, word) {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  
  // Filtra as linhas que não contêm a palavra especificada
  var filteredData = data.filter(function(row) {
    return !row.some(function(cell) {
      return typeof cell === 'string' && cell.includes(word);
    });
  });

  // Limpa a planilha e reinsere os dados filtrados
  sheet.clear();
  if (filteredData.length > 0) {
    sheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
  }

  Logger.log('Linhas com a palavra "' + word + '" removidas.');
}

function removeDuplicateRows(spreadsheetId, sheetName) {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();
  
  var uniqueData = [];
  var seenRows = new Set();
  
  data.forEach(function(row) {
    // Altere aqui as colunas para verificar duplicatas
    var key = row[0] + row[3]; // Exemplo: considera as colunas A e D
    if (!seenRows.has(key)) {
      seenRows.add(key);
      uniqueData.push(row);
    }
  });
  
  sheet.clear();
  if (uniqueData.length > 0) {
    sheet.getRange(1, 1, uniqueData.length, uniqueData[0].length).setValues(uniqueData);
  }
  
  Logger.log('Linhas duplicadas removidas.');
}

function findLatestFile(files) {
  var latestFile = null;
  var latestDate = null;
  
  while (files.hasNext()) {
    var file = files.next();
    var fileDate = file.getLastUpdated();
    if (!latestDate || fileDate > latestDate) {
      latestFile = file;
      latestDate = fileDate;
    }
  }
  
  return latestFile;
}

function findLatestXlsFile(files) {
  var latestFile = null;
  var latestDate = null;
  
  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().match(/\.xls(x)?$/i)) {
      var fileDate = file.getLastUpdated();
      if (!latestDate || fileDate > latestDate) {
        latestFile = file;
        latestDate = fileDate;
      }
    }
  }
  
  return latestFile;
}

function cleanData(data, columnsToRemove) {
  return data.map(function(row) {
    return row.filter(function(_, index) {
      return !columnsToRemove.includes(index);
    });
  }).filter(function(row) {
    return row.length > 0 && row.some(cell => cell !== ''); // Filtra linhas vazias
  });
}

function updateSpreadsheet(spreadsheetId, sheetName, data, clearSheet) {
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  
  if (clearSheet) sheet.clear();
  
  // Filtrar dados não vazios
  data = data.filter(function(row) {
    return row.some(function(cell) {
      return cell !== '';
    });
  });

  if (data.length > 0) {
    var existingData = sheet.getDataRange().getValues();
    if (existingData.length > 0) {
      sheet.getRange(existingData.length + 1, 1, data.length, data[0].length).setValues(data);
    } else {
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }
  } else {
    Logger.log('Nenhum dado válido para adicionar.');
  }
}

function parseCsvWithDelimiter(csvContent, delimiter) {
  return csvContent.split('\n').map(function(row) {
    return row.split(delimiter);
  });
}
