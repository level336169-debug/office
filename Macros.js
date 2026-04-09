// OnlyOffice макрос
// Адаптировано из VBA

(function() {
    // Вспомогательная функция для получения значения ячейки
    function getCellValue(sheet, row, col) {
        var colLetter = String.fromCharCode(64 + col); // A=1 -> A, B=2 -> B
        return sheet.GetRange(colLetter + row).GetValue();
    }
    
    // Вспомогательная функция для установки значения ячейки
    function setCellValue(sheet, row, col, value) {
        var colLetter = String.fromCharCode(64 + col);
        sheet.GetRange(colLetter + row).SetValue(value);
    }
    
    // Основная функция AUTOSET
    function autoSet() {
        var oSheetBase = Api.GetSheet("Base");
        var oSheetEmployees = Api.GetSheet("Сотрудники");
        var oSheetDeals = Api.GetSheet("Сделки");
        
        if (!oSheetBase || !oSheetEmployees || !oSheetDeals) {
            console.log("Не найдены нужные листы!");
            return;
        }
        
        // Очистка листа Base
        oSheetBase.GetRange("A2:R900").ClearContents();
        
        var n = 2;
        var firstRow = 2;
        
        // Получаем последние заполненные строки
        var maxUsedRow = oSheetEmployees.GetUsedRange().GetRows().GetCount();
        var maxUsedRowServ = oSheetDeals.GetUsedRange().GetRows().GetCount();
        
        var summServ = getCellValue(oSheetDeals, 2, 3);
        var summServItog = getCellValue(oSheetDeals, 2, 4);
        
        // Цикл по сделкам
        for (var j = firstRow; j <= maxUsedRowServ; j++) {
            while (getCellValue(oSheetDeals, j, 4) > 10 && 
                   getCellValue(oSheetEmployees, 2, 10) > 0) {
                
                for (var i = firstRow; i <= maxUsedRow; i++) {
                    if (getCellValue(oSheetDeals, j, 4) < 10 || 
                        getCellValue(oSheetEmployees, 2, 10) === 0) {
                        break;
                    }
                    
                    if (getCellValue(oSheetEmployees, i, 6) > 0.01) {
                        var stavka = getCellValue(oSheetEmployees, i, 2);
                        var hours = Math.round(getCellValue(oSheetDeals, j, 4) / stavka * 100) / 100;
                        var availableHours = getCellValue(oSheetEmployees, i, 6);
                        
                        if (hours > availableHours) {
                            // Запись в Base
                            setCellValue(oSheetBase, n, 1, getCellValue(oSheetEmployees, 2, 8));
                            setCellValue(oSheetBase, n, 2, getCellValue(oSheetDeals, j, 1));
                            setCellValue(oSheetBase, n, 3, getCellValue(oSheetDeals, j, 2));
                            setCellValue(oSheetBase, n, 4, getCellValue(oSheetEmployees, i, 1));
                            setCellValue(oSheetBase, n, 5, getCellValue(oSheetEmployees, i, 4));
                            setCellValue(oSheetBase, n, 7, getCellValue(oSheetEmployees, i, 5));
                            setCellValue(oSheetBase, n, 9, getCellValue(oSheetDeals, i, 5));
                            setCellValue(oSheetBase, n, 8, Math.round(availableHours * 100) / 100);
                            
                            // Обновление остатков
                            var newServValue = Math.round((getCellValue(oSheetDeals, j, 4) - availableHours * stavka) * 100) / 100;
                            setCellValue(oSheetDeals, j, 4, newServValue);
                            setCellValue(oSheetEmployees, i, 6, 0);
                            
                            n++;
                        } else {
                            setCellValue(oSheetBase, n, 1, getCellValue(oSheetEmployees, 2, 8));
                            setCellValue(oSheetBase, n, 2, getCellValue(oSheetDeals, j, 1));
                            setCellValue(oSheetBase, n, 3, getCellValue(oSheetDeals, j, 2));
                            setCellValue(oSheetBase, n, 4, getCellValue(oSheetEmployees, i, 1));
                            setCellValue(oSheetBase, n, 5, getCellValue(oSheetEmployees, i, 4));
                            setCellValue(oSheetBase, n, 7, getCellValue(oSheetEmployees, i, 5));
                            setCellValue(oSheetBase, n, 9, getCellValue(oSheetDeals, i, 5));
                            setCellValue(oSheetBase, n, 8, hours);
                            
                            setCellValue(oSheetEmployees, i, 6, 
                                getCellValue(oSheetEmployees, i, 6) - hours);
                            
                            var newServValue2 = Math.round((getCellValue(oSheetDeals, j, 4) - hours * stavka) * 100) / 100;
                            setCellValue(oSheetDeals, j, 4, newServValue2);
                        }
                    }
                }
                n++;
            }
        }
    }
    
    // Обработчик события изменения листа
    function onSheetChange(oSheet, oRange) {
        var sheetName = oSheet.GetName();
        
        if (sheetName === "Сотрудники" || sheetName === "Сделки") {
            var oSheetEmployees = Api.GetSheet("Сотрудники");
            var oSheetDeals = Api.GetSheet("Сделки");
            
            // Проверяем, изменилась ли ячейка I2
            if (oRange.GetCells().GetCell(1, 1).GetAddress() === "$I$2") {
                var maxUsedRow = oSheetEmployees.GetUsedRange().GetRows().GetCount();
                var maxUsedRowServ = oSheetDeals.GetUsedRange().GetRows().GetCount();
                var valueI2 = getCellValue(oSheetEmployees, 2, 9);
                
                // Заполняем колонку F значением из I2
                for (var i = 2; i <= maxUsedRow; i++) {
                    setCellValue(oSheetEmployees, i, 6, valueI2);
                }
                
                // Копируем значения из колонки C в D
                for (var i = 2; i <= maxUsedRowServ; i++) {
                    setCellValue(oSheetDeals, i, 4, getCellValue(oSheetDeals, i, 3));
                }
            }
        }
    }
    
    // Регистрация обработчика (для автоматического запуска)
    // Api.attachEvent("sheetChange", onSheetChange);
    
    // Экспорт
    return {
        AUTOSET: autoSet,
        onSheetChange: onSheetChange
    };
})();