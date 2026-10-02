// R7-Office macro: build summary report sheet.
// Counts rows before/after, errors by type, deleted duplicates.
(function () {
    function getOrCreate(name) {
        var sheets = Api.GetAllSheets();
        for (var i = 0; i < sheets.length; i++) {
            if (sheets[i].GetName() === name) return sheets[i];
        }
        return Api.AddSheet(name);
    }

    var rep = getOrCreate('Итоговый_отчет');
    rep.GetRange('A1:D1').SetValue([['Файл', 'До очистки', 'После очистки', 'Удалено дубликатов']]);

    var files = [
        ['Товары',    'Товары_исходные',    'Товары_очищенные'],
        ['Поставщики','Поставщики_исходные','Поставщики_очищенные'],
        ['Продажи',   'Продажи_исходные',   'Продажи_очищенные'],
        ['Клиенты',   'Клиенты_исходные',   'Клиенты_очищенные']
    ];

    for (var i = 0; i < files.length; i++) {
        var row = 2 + i;
        rep.GetRange('A' + row).SetValue(files[i][0]);
        rep.GetRange('B' + row).SetFormula('=COUNTA(\'' + files[i][1] + '\'!A:A)-1');
        rep.GetRange('C' + row).SetFormula('=COUNTA(\'' + files[i][2] + '\'!A:A)-1');
        rep.GetRange('D' + row).SetFormula('=B' + row + '-C' + row);
    }

    rep.GetRange('A8').SetValue('Ошибки по типам');
    rep.GetRange('A9').SetValue('Ошибка ИНН');
    rep.GetRange('B9').SetFormula('=COUNTIF(\'Поставщики_очищенные\'!J:J,"Ошибка ИНН")');
    rep.GetRange('A10').SetValue('Ошибка Email (поставщики)');
    rep.GetRange('B10').SetFormula('=COUNTIF(\'Поставщики_очищенные\'!K:K,"Ошибка Email")');
    rep.GetRange('A11').SetValue('Ошибка Email (клиенты)');
    rep.GetRange('B11').SetFormula('=COUNTIF(\'Клиенты_очищенные\'!J:J,"Ошибка Email")');
    rep.GetRange('A12').SetValue('Ошибка даты рождения');
    rep.GetRange('B12').SetFormula('=COUNTIF(\'Клиенты_очищенные\'!I:I,"Ошибка даты рождения")');
    rep.GetRange('A13').SetValue('Розница ниже закупки');
    rep.GetRange('B13').SetFormula('=COUNTIF(\'Товары_очищенные\'!H:H,"Розница ниже закупки")');
    rep.GetRange('A14').SetValue('Дубликаты артикулов');
    rep.GetRange('B14').SetFormula('=COUNTIF(\'Товары_очищенные\'!I:I,"Дубликат")');
})();
