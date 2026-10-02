// R7-Office macro: clean product catalog with formulas.
// Source: 'Товары_исходные' (A=Артикул, B=Наименование, C=Категория,
//          D=Производитель, E=Цена_закупки, F=Цена_розница, G=ЕдИзм)
// Target: 'Товары_очищенные' (same layout, formulas in each cell)
(function () {
    var SRC = 'Товары_исходные';
    var DST = 'Товары_очищенные';

    function getOrCreate(name) {
        var sheets = Api.GetAllSheets();
        for (var i = 0; i < sheets.length; i++) {
            if (sheets[i].GetName() === name) return sheets[i];
        }
        return Api.AddSheet(name);
    }

    var src = getOrCreate(SRC);
    var dst = getOrCreate(DST);
    var lr = src.GetUsedRange().GetLastRow();     // 0-indexed
    if (lr < 1) return;

    // Copy raw snapshot (header + rows). Formulas overwrite in next loop.
    dst.GetRange('A1:G' + (lr + 1)).SetValue(src.GetRange('A1:G' + (lr + 1)).GetValue());

    var S = "'" + SRC + "'!";

    for (var r = 1; r <= lr; r++) {
        var n = r + 1;

        // Артикул -> АРТ-XXX
        dst.GetRange('A' + n).SetFormula(
            '="АРТ-"&LEFT(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(' + S + 'A' + n +
            ',"№",""),"АРТ-",""),"АРТ_",""),"/","-"),3)'
        );

        // Наименование: strip (...) and Title Case
        dst.GetRange('B' + n).SetFormula(
            '=PROPER(TRIM(IFERROR(LEFT(' + S + 'B' + n + ',FIND("(",' + S + 'B' + n +
            ')-1),' + S + 'B' + n + ')))'
        );

        // Категория via VLOOKUP -> Справочники!A:B
        dst.GetRange('C' + n).SetFormula(
            '=IFERROR(VLOOKUP(' + S + 'C' + n +
            ',\'Справочники\'!$A$2:$B$20,2,FALSE),' + S + 'C' + n + ')'
        );

        // Производитель: Title Case
        dst.GetRange('D' + n).SetFormula(
            '=PROPER(TRIM(' + S + 'D' + n + '))'
        );

        // Цена_закупки: strip spaces/руб, comma -> dot
        dst.GetRange('E' + n).SetFormula(
            '=IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(' +
            S + 'E' + n + '," ",""),"руб.",""),"руб",""),",",".")),0)'
        );

        // Цена_розница
        dst.GetRange('F' + n).SetFormula(
            '=IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(' +
            S + 'F' + n + '," ",""),"руб.",""),"руб",""),",",".")),0)'
        );

        // Единица_измерения -> шт.
        dst.GetRange('G' + n).SetFormula(
            '=IFERROR(VLOOKUP(' + S + 'G' + n +
            ',\'Справочники\'!$D$2:$E$10,2,FALSE),"шт.")'
        );

        // Валидация: розница >= закупки
        dst.GetRange('H1').SetValue('Проверка_цены');
        dst.GetRange('H' + n).SetFormula(
            '=IF(F' + n + '>=E' + n + ',"OK","Розница ниже закупки")'
        );

        // Валидация: артикул уникален
        dst.GetRange('I1').SetValue('Проверка_уникальности');
        dst.GetRange('I' + n).SetFormula(
            '=IF(COUNTIF($A$2:$A$' + (lr + 1) + ',A' + n + ')=1,"OK","Дубликат")'
        );
    }
})();
