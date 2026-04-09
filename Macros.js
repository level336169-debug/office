// OnlyOffice макрос
// Адаптировано из VBA

(function() {
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
        
        var summServ = oSheetDeals.GetCells().GetCell(2, 3).GetValue();
        var summServItog = oSheetDeals.GetCells().GetCell(2, 4).GetValue();
        
        // Цикл по сделкам
        for (var j = firstRow; j <= maxUsedRowServ; j++) {
            while (oSheetDeals.GetCells().GetCell(j, 4).GetValue() > 10 && 
                   oSheetEmployees.GetCells().GetCell(2, 10).GetValue() > 0) {
                
                for (var i = firstRow; i <= maxUsedRow; i++) {
                    if (oSheetDeals.GetCells().GetCell(j, 4).GetValue() < 10 || 
                        oSheetEmployees.GetCells().GetCell(2, 10).GetValue() === 0) {
                        break;
                    }
                    
                    if (oSheetEmployees.GetCells().GetCell(i, 6).GetValue() > 0.01) {
                        var stavka = oSheetEmployees.GetCells().GetCell(i, 2).GetValue();
                        var hours = Math.round(oSheetDeals.GetCells().GetCell(j, 4).GetValue() / stavka * 100) / 100;
                        var availableHours = oSheetEmployees.GetCells().GetCell(i, 6).GetValue();
                        
                        if (hours > availableHours) {
                            // Запись в Base
                            oSheetBase.GetCells().GetCell(n, 1).SetValue(oSheetEmployees.GetCells().GetCell(2, 8).GetValue());
                            oSheetBase.GetCells().GetCell(n, 2).SetValue(oSheetDeals.GetCells().GetCell(j, 1).GetValue());
                            oSheetBase.GetCells().GetCell(n, 3).SetValue(oSheetDeals.GetCells().GetCell(j, 2).GetValue());
                            oSheetBase.GetCells().GetCell(n, 4).SetValue(oSheetEmployees.GetCells().GetCell(i, 1).GetValue());
                            oSheetBase.GetCells().GetCell(n, 5).SetValue(oSheetEmployees.GetCells().GetCell(i, 4).GetValue());
                            oSheetBase.GetCells().GetCell(n, 7).SetValue(oSheetEmployees.GetCells().GetCell(i, 5).GetValue());
                            oSheetBase.GetCells().GetCell(n, 9).SetValue(oSheetDeals.GetCells().GetCell(i, 5).GetValue());
                            oSheetBase.GetCells().GetCell(n, 8).SetValue(Math.round(availableHours * 100) / 100);
                            
                            // Обновление остатков
                            var newServValue = Math.round((oSheetDeals.GetCells().GetCell(j, 4).GetValue() - availableHours * stavka) * 100) / 100;
                            oSheetDeals.GetCells().GetCell(j, 4).SetValue(newServValue);
                            oSheetEmployees.GetCells().GetCell(i, 6).SetValue(0);
                            
                            n++;
                        } else {
                            oSheetBase.GetCells().GetCell(n, 1).SetValue(oSheetEmployees.GetCells().GetCell(2, 8).GetValue());
                            oSheetBase.GetCells().GetCell(n, 2).SetValue(oSheetDeals.GetCells().GetCell(j, 1).GetValue());
                            oSheetBase.GetCells().GetCell(n, 3).SetValue(oSheetDeals.GetCells().GetCell(j, 2).GetValue());
                            oSheetBase.GetCells().GetCell(n, 4).SetValue(oSheetEmployees.GetCells().GetCell(i, 1).GetValue());
                            oSheetBase.GetCells().GetCell(n, 5).SetValue(oSheetEmployees.GetCells().GetCell(i, 4).GetValue());
                            oSheetBase.GetCells().GetCell(n, 7).SetValue(oSheetEmployees.GetCells().GetCell(i, 5).GetValue());
                            oSheetBase.GetCells().GetCell(n, 9).SetValue(oSheetDeals.GetCells().GetCell(i, 5).GetValue());
                            oSheetBase.GetCells().GetCell(n, 8).SetValue(hours);
                            
                            oSheetEmployees.GetCells().GetCell(i, 6).SetValue(
                                oSheetEmployees.GetCells().GetCell(i, 6).GetValue() - hours
                            );
                            
                            var newServValue2 = Math.round((oSheetDeals.GetCells().GetCell(j, 4).GetValue() - hours * stavka) * 100) / 100;
                            oSheetDeals.GetCells().GetCell(j, 4).SetValue(newServValue2);
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
            
            var oCell = oSheetEmployees.GetRange("I2");
            
            // Проверяем, изменилась ли ячейка I2
            if (oRange.GetCells().GetCell(1, 1).GetAddress() === "$I$2") {
                var maxUsedRow = oSheetEmployees.GetUsedRange().GetRows().GetCount();
                var maxUsedRowServ = oSheetDeals.GetUsedRange().GetRows().GetCount();
                var valueI2 = oSheetEmployees.GetCells().GetCell(2, 9).GetValue();
                
                // Заполняем колонку F значением из I2
                for (var i = 2; i <= maxUsedRow; i++) {
                    oSheetEmployees.GetCells().GetCell(i, 6).SetValue(valueI2);
                }
                
                // Копируем значения из колонки C в D
                for (var i = 2; i <= maxUsedRowServ; i++) {
                    oSheetDeals.GetCells().GetCell(i, 4).SetValue(
                        oSheetDeals.GetCells().GetCell(i, 3).GetValue()
                    );
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