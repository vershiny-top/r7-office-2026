// R7-Office macro: clean clients.
// Source: 'Клиенты_исходные' A=ID B=ФИО C=Дата_рождения D=Телефон
//          E=Email F=Город G=Дата_регистрации
(function () {
    var SRC = 'Клиенты_исходные';
    var DST = 'Клиенты_очищенные';

    function getOrCreate(name) {
        var sheets = Api.GetAllSheets();
        for (var i = 0; i < sheets.length; i++) {
            if (sheets[i].GetName() === name) return sheets[i];
        }
        return Api.AddSheet(name);
    }

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
    dst.GetRange('H1').SetValue('Возраст');
    dst.GetRange('I1').SetValue('Проверка_возраста');
    dst.GetRange('J1').SetValue('Проверка_Email');

    var S = "'" + SRC + "'!";

    for (var r = 1; r <= lr; r++) {
        var n = r + 1;

        // ФИО -> Title Case
        dst.GetRange('B' + n).SetFormula('=PROPER(TRIM(' + S + 'B' + n + '))');

        // Дата рождения: "DD.MM.YYYY" or "YYYY-MM-DD"
        var bd = S + 'C' + n;
        dst.GetRange('C' + n).SetFormula(
            '=IFERROR(IF(ISNUMBER(FIND(".",' + bd + ')),' +
            'DATE(RIGHT(' + bd + ',4),MID(' + bd + ',4,2),LEFT(' + bd + ',2)),' +
            'DATE(LEFT(' + bd + ',4),MID(' + bd + ',6,2),RIGHT(' + bd + ',2))),' + bd + ')'
        );

        // Телефон
        var ph = S + 'D' + n;
        var phDigits = 'TEXT(' + digitsFormula(ph) + ',"00000000000")';
        dst.GetRange('D' + n).SetFormula(
            '=IF(LEN(' + phDigits + ')=11,"+7 ("&MID(' + phDigits + ',2,3)&") "&MID(' +
            phDigits + ',5,3)&"-"&MID(' + phDigits + ',8,2)&"-"&MID(' + phDigits + ',10,2),' + ph + ')'
        );

        // Email
        dst.GetRange('E' + n).SetFormula('=TRIM(' + S + 'E' + n + ')');

        // Город: VLOOKUP (МСК/СПб -> Москва/СПб)
        dst.GetRange('F' + n).SetFormula(
            '=IFERROR(VLOOKUP(' + S + 'F' + n +
            ',\'Справочники\'!$J$2:$K$50,2,FALSE),' + S + 'F' + n + ')'
        );

        // Дата регистрации
        var rd = S + 'G' + n;
        dst.GetRange('G' + n).SetFormula(
            '=IFERROR(IF(ISNUMBER(FIND(".",' + rd + ')),' +
            'DATE(RIGHT(' + rd + ',4),MID(' + rd + ',4,2),LEFT(' + rd + ',2)),' +
            'DATE(LEFT(' + rd + ',4),MID(' + rd + ',6,2),RIGHT(' + rd + ',2))),' + rd + ')'
        );

        // Возраст на сегодня
        dst.GetRange('H' + n).SetFormula('=DATEDIF(C' + n + ',TODAY(),"Y")');

        // Проверка возраста: 18..80
        dst.GetRange('I' + n).SetFormula(
            '=IF(OR(H' + n + '<18,H' + n + '>80),"Ошибка даты рождения","OK")'
        );

        // Проверка Email
        dst.GetRange('J' + n).SetFormula(
            '=IF(AND(ISNUMBER(FIND("@",E' + n + ')),ISNUMBER(FIND(".",MID(E' + n +
            ',FIND("@",E' + n + ')+1,255)))),"OK","Ошибка Email")'
        );
    }
})();
