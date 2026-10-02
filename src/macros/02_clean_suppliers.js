// R7-Office macro: clean suppliers.
// Source: 'Поставщики_исходные' A=ID B=Название C=ИНН D=КПП E=Телефон
//          F=Email G=Адрес
// Target: A..G cleaned, H=Индекс, I=Адрес_остаток, J/K = validations.
(function () {
    var SRC = 'Поставщики_исходные';
    var DST = 'Поставщики_очищенные';

    function getOrCreate(name) {
        var sheets = Api.GetAllSheets();
        for (var i = 0; i < sheets.length; i++) {
            if (sheets[i].GetName() === name) return sheets[i];
        }
        return Api.AddSheet(name);
    }

    // classic Excel trick: extract digits from a text cell
    function digitsFormula(ref) {
        return 'SUMPRODUCT(MID(0&' + ref + ',LARGE(INDEX(ISNUMBER(--MID(' + ref +
               ',ROW(INDIRECT("1:"&LEN(' + ref + '))),1))*ROW(INDIRECT("1:"&LEN(' + ref +
               '))),0),ROW(INDIRECT("1:"&LEN(' + ref + '))))+1,1)*10^ROW(INDIRECT("1:"&LEN(' + ref + ')))/10)';
    }

    var src = getOrCreate(SRC);
    var dst = getOrCreate(DST);
    var lr = src.GetUsedRange().GetLastRow();
    if (lr < 1) return;

    dst.GetRange('A1:G' + (lr + 1)).SetValue(src.GetRange('A1:G' + (lr + 1)).GetValue());
    dst.GetRange('H1').SetValue('Индекс');
    dst.GetRange('I1').SetValue('Адрес_остаток');
    dst.GetRange('J1').SetValue('Проверка_ИНН');
    dst.GetRange('K1').SetValue('Проверка_Email');

    var S = "'" + SRC + "'!";

    for (var r = 1; r <= lr; r++) {
        var n = r + 1;

        // ИНН: digits only (formula). Note: leading zeros are lost, INN never start with 0.
        dst.GetRange('C' + n).SetFormula('=' + digitsFormula(S + 'C' + n));

        // Телефон: digits -> +7 (XXX) XXX-XX-XX
        var ph = S + 'E' + n;
        var phDigits = 'TEXT(' + digitsFormula(ph) + ',"00000000000")';
        dst.GetRange('E' + n).SetFormula(
            '=IF(LEN(' + phDigits + ')=11,"+7 ("&MID(' + phDigits + ',2,3)&") "&MID(' +
            phDigits + ',5,3)&"-"&MID(' + phDigits + ',8,2)&"-"&MID(' + phDigits + ',10,2),' + ph + ')'
        );

        // Email: strip spaces
        dst.GetRange('F' + n).SetFormula('=TRIM(' + S + 'F' + n + ')');

        // Адрес -> Индекс (6 digits) + остаток
        dst.GetRange('H' + n).SetFormula(
            '=IF(ISNUMBER(VALUE(LEFT(TRIM(' + S + 'G' + n + '),6))),LEFT(TRIM(' + S + 'G' + n + '),6),"")'
        );
        dst.GetRange('I' + n).SetFormula(
            '=TRIM(MID(TRIM(' + S + 'G' + n + '),7,255))'
        );

        // Валидация ИНН: длина 10 или 12
        dst.GetRange('J' + n).SetFormula(
            '=IF(OR(LEN(C' + n + ')=10,LEN(C' + n + ')=12),"OK","Ошибка ИНН")'
        );

        // Валидация Email: есть @ и точка после @
        dst.GetRange('K' + n).SetFormula(
            '=IF(AND(ISNUMBER(FIND("@",F' + n + ')),ISNUMBER(FIND(".",MID(F' + n +
            ',FIND("@",F' + n + ')+1,255)))),"OK","Ошибка Email")'
        );
    }
})();
