// R7-Office macro: clean sales.
// Source: 'Продажи_исходные' A=Дата B=Артикул C=Количество D=Цена_продажи
//          E=ID_клиента F=ID_поставщика G=Статус
// Target: H=Розничная_цена, I=Категория (VLOOKUP in Товары_очищенные)
(function () {
    var SRC = 'Продажи_исходные';
    var DST = 'Продажи_очищенные';

    function getOrCreate(name) {
        var sheets = Api.GetAllSheets();
        for (var i = 0; i < sheets.length; i++) {
            if (sheets[i].GetName() === name) return sheets[i];
        }
        return Api.AddSheet(name);
    }

    var src = getOrCreate(SRC);
    var dst = getOrCreate(DST);
    var lr = src.GetUsedRange().GetLastRow();
    if (lr < 1) return;

    dst.GetRange('A1:G' + (lr + 1)).SetValue(src.GetRange('A1:G' + (lr + 1)).GetValue());
    dst.GetRange('H1').SetValue('Розничная_цена');
    dst.GetRange('I1').SetValue('Категория_товара');
    dst.GetRange('J1').SetValue('Проверка_цены');

    var S = "'" + SRC + "'!";

    for (var r = 1; r <= lr; r++) {
        var n = r + 1;

        // Дата MM/DD/YYYY -> DATE(yyyy, mm, dd)
        dst.GetRange('A' + n).SetFormula(
            '=DATE(RIGHT(' + S + 'A' + n + ',4),LEFT(' + S + 'A' + n + ',2),MID(' + S + 'A' + n + ',4,2))'
        );

        // Артикул -> АРТ-XXX (same cleaning as products)
        dst.GetRange('B' + n).SetFormula(
            '="АРТ-"&LEFT(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(' + S + 'B' + n +
            ',"№",""),"АРТ-",""),"АРТ_",""),"/","-"),3)'
        );

        // Количество -> number (strips "шт", spaces, etc.)
        dst.GetRange('C' + n).SetFormula(
            '=IFERROR(VALUE(TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(' +
            S + 'C' + n + '," шт.","")," шт",""),"шт.",""),"шт",""))),0)'
        );

        // Цена_продажи -> RUB. Rates: $=95, €=100, ₽=1.
        dst.GetRange('D' + n).SetFormula(
            '=IF(LEFT(' + S + 'D' + n + ',1)="$",VALUE(MID(' + S + 'D' + n + ',2,255))*95,' +
            'IF(LEFT(' + S + 'D' + n + ',1)="€",VALUE(MID(' + S + 'D' + n + ',2,255))*100,' +
            'VALUE(SUBSTITUTE(SUBSTITUTE(' + S + 'D' + n + ',"₽","")," ",""))))'
        );

        // Статус через VLOOKUP
        dst.GetRange('G' + n).SetFormula(
            '=IFERROR(VLOOKUP(' + S + 'G' + n +
            ',\'Справочники\'!$G$2:$H$15,2,FALSE),' + S + 'G' + n + ')'
        );

        // VLOOKUP в каталог: розничная цена = колонка 6, категория = колонка 3
        dst.GetRange('H' + n).SetFormula(
            '=IFERROR(VLOOKUP(B' + n + ',\'Товары_очищенные\'!$A$1:$G$1000,6,FALSE),"нет в каталоге")'
        );
        dst.GetRange('I' + n).SetFormula(
            '=IFERROR(VLOOKUP(B' + n + ',\'Товары_очищенные\'!$A$1:$G$1000,3,FALSE),"нет в каталоге")'
        );

        // Валидация: цена продажи >= закупочной не проверяется (закупки нет в продажах),
        // поэтому проверяем только что цена положительна
        dst.GetRange('J' + n).SetFormula(
            '=IF(D' + n + '>0,"OK","Некорректная цена")'
        );
    }
})();
