const { doc } = require('./src/Structures/Untils/googlesheet.js');

// Функция для получения индекса столбца по его буквенному обозначению
function getColumnIndex(column) {
    const ACode = 'A'.charCodeAt(0);
    return column.charCodeAt(0) - ACode;
}

// Асинхронная функция, которая выполняет действия с таблицей
async function copyPasteColumns(doc, sheetId, sourceColumn, targetColumn, startRow, endRow) {
    try {
        await doc.loadInfo();
        const sheet = doc.sheetsById[sheetId];

        // Загрузка ячеек и проверка их загрузки
        await sheet.loadCells(`${sourceColumn}${startRow}:${sourceColumn}${endRow}`);
        await sheet.loadCells(`${targetColumn}${startRow}:${targetColumn}${endRow}`);

        // Проверка, что ячейки загружены
        const sourceCell = sheet.getCell(startRow, getColumnIndex(sourceColumn));
        const targetCell = sheet.getCell(startRow, getColumnIndex(targetColumn));

        if (!sourceCell.isLoaded() || !targetCell.isLoaded()) {
            console.error('Ошибка: Ячейки не загружены полностью.');
            return;
        }

        // Копирование значений из sourceColumn в targetColumn
        for (let i = startRow; i <= endRow; i++) {
            sheet.getCell(i, getColumnIndex(targetColumn)).value = sheet.getCell(i, getColumnIndex(sourceColumn)).value;
        }

        // Сохранение изменений
        await sheet.saveUpdatedCells();
    } catch (error) {
        console.error("Произошла ошибка:", error);
    }
}

// Основная функция
async function main() {
    const sheetId = 1162940648;
    const sourceColumn = 'BM';
    const targetColumn = 'BN';
    const startRow = 4;
    const endRow = 28;

    try {
        await doc.loadInfo();

        // Вызов функции copyPasteColumns
        await copyPasteColumns(doc, sheetId, sourceColumn, targetColumn, startRow, endRow);

        console.log('Значения скопированы успешно.');
    } catch (error) {
        console.error('Произошла ошибка:', error);
    }
}

main();
