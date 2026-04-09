// OnlyOffice макрос - адаптировано из VBA
// Скопируй этот код в редактор скриптов OnlyOffice

// Получение листа по имени
function getSheet(sheetName) {
    var doc = Api.GetDocument();
    var sheets = doc.Sheets;
    for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].Name === sheetName) {
            return sheets[i];
        }
    }
    return null;
}

// Получение значения ячейки
function getCell(sheet, row, col) {
    var colLetter = String.fromCharCode(64 + col);
    return sheet.GetRange(colLetter + row).GetValue();
}

// Установка значения ячейки
function setCell(sheet, row, col, value) {
    var colLetter = String.fromCharCode(64 + col);
    sheet.GetRange(colLetter + row).SetValue(value);
}

// Главная функция AUTOSET
function AUTOSET() {
    var sheetBase = getSheet("Base");
    var sheetEmp = getSheet("Сотрудники");
    var sheetDeals = getSheet("Сделки");
    
    if (!sheetBase || !sheetEmp || !sheetDeals) {
        Api.ShowMessage("Листы не найдены!");
        return;
    }
    
    // Очистка
    sheetBase.GetRange("A2:R900").ClearContents();
    
    var firstRow = 2;
    var n = 2;
    
    // Количество строк
    var maxRowEmp = sheetEmp.GetUsedRange().GetRows().GetCount();
    var maxRowDeals = sheetDeals.GetUsedRange().GetRows().GetCount();
    
    // Цикл по сделкам
    for (var j = firstRow; j <= maxRowDeals; j++) {
        while (getCell(sheetDeals, j, 4) > 10 && getCell(sheetEmp, 2, 10) > 0) {
            
            for (var i = firstRow; i <= maxRowEmp; i++) {
                if (getCell(sheetDeals, j, 4) < 10 || getCell(sheetEmp, 2, 10) === 0) {
                    break;
                }
                
                var availHours = getCell(sheetEmp, i, 6);
                if (availHours > 0.01) {
                    var stavka = getCell(sheetEmp, i, 2);
                    var dealValue = getCell(sheetDeals, j, 4);
                    var hour = Math.round(dealValue / stavka * 100) / 100;
                    
                    if (hour > availHours) {
                        // Запись в Base
                        setCell(sheetBase, n, 1, getCell(sheetEmp, 2, 8));   // Период
                        setCell(sheetBase, n, 2, getCell(sheetDeals, j, 1));  // Сделка
                        setCell(sheetBase, n, 3, getCell(sheetDeals, j, 2));  // Услуга
                        setCell(sheetBase, n, 4, getCell(sheetEmp, i, 1));    // ФИО
                        setCell(sheetBase, n, 5, getCell(sheetEmp, i, 4));    // Таб номер
                        setCell(sheetBase, n, 7, getCell(sheetEmp, i, 5));    // Подразделение
                        setCell(sheetBase, n, 9, getCell(sheetDeals, i, 5));  // РВ
                        setCell(sheetBase, n, 8, Math.round(availHours * 100) / 100); // Часы
                        
                        // Обновление остатков
                        var newDealValue = Math.round((dealValue - availHours * stavka) * 100) / 100;
                        setCell(sheetDeals, j, 4, newDealValue);
                        setCell(sheetEmp, i, 6, 0);
                        
                        n++;
                    } else {
                        setCell(sheetBase, n, 1, getCell(sheetEmp, 2, 8));
                        setCell(sheetBase, n, 2, getCell(sheetDeals, j, 1));
                        setCell(sheetBase, n, 3, getCell(sheetDeals, j, 2));
                        setCell(sheetBase, n, 4, getCell(sheetEmp, i, 1));
                        setCell(sheetBase, n, 5, getCell(sheetEmp, i, 4));
                        setCell(sheetBase, n, 7, getCell(sheetEmp, i, 5));
                        setCell(sheetBase, n, 9, getCell(sheetDeals, i, 5));
                        setCell(sheetBase, n, 8, hour);
                        
                        setCell(sheetEmp, i, 6, availHours - hour);
                        var newDealValue2 = Math.round((dealValue - hour * stavka) * 100) / 100;
                        setCell(sheetDeals, j, 4, newDealValue2);
                    }
                }
            }
            n++;
        }
    }
    
    Api.ShowMessage("Готово! Записей: " + (n - 2));
}

// Для запуска просто вызови AUTOSET()
//AUTOSET();
