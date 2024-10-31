******************************************************************** PROCEDURE start ************************************************************

PROCEDURE Start
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

f_IdDepartmentPRK = GetGlob('DepartPRK')  && Номер подразделения ПРК по справочнику Department
f_IdDepartmentKP = GetGlob('DepartKP')  && Номер подразделения КП по справочнику Department

IdShiftCFT = 0

new_data_odb_cft = SUBSTR(DTOC(DATE()), 1, 2) + SUBSTR(DTOC(DATE()), 4, 2) + SUBSTR(DTOC(DATE()), 9, 2)

DIMENSION ar[1,2]

PRIVATE DateSmenaPRK

STORE SPACE(0) TO date_exp_file, time_exp_file, exp_file_dat, exp_file_txt, exp_file_cft, exp_file_title

* f_OpDate = CursOperDate.OpDate  && Дата открытого операционного дня КП
* f_OpDate_Den = CursOperDate.OpDate  && Дата открытого операционного дня КП

DateSmenaPRK = OpenDatePRK  && Формат DateTime - Календарная дата смены из таблицы Shift поле ReportDate
DateSmenaKP = OpenDateKP  && Формат DateTime - Календарная дата смены из таблицы Shift поле ReportDate

* WAIT WINDOW 'f_OpDate = ' + TTOC(f_OpDate) TIMEOUT 3
* WAIT WINDOW 'f_OpDate_Den = ' + TTOC(f_OpDate_Den) TIMEOUT 3

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

dtB = new_data_odb  && Дата открытой смены в КП
f_CurDate = new_data_odb  && Дата открытой смены в КП

* WAIT WINDOW 'dtB = ' + DTOC(dtB) TIMEOUT 3
* WAIT WINDOW '1 - f_CurDate = ' + DTOC(f_CurDate) TIMEOUT 3

* ------------------------------------ *
DO FORM GetDate1
* ------------------------------------ *

IF NOT LASTKEY() = 27

  f_CurDate = DTOT(dtb)

ELSE
  RETURN
ENDIF

* WAIT WINDOW '2 - f_CurDate = ' + DTOC(f_CurDate) TIMEOUT 3

* ---------------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

* ---------------------------------------- Находим смену КП за дату f_CurDate ------------------------------------- *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT DISTINCT Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate, Od.OpenDate AS OpenDateDen
  FROM Shift AS Sh (READPAST)
  INNER JOIN OperDate AS Od (READPAST) ON Od.Id = Sh.IdOperDate
  INNER JOIN PaketKasOrderATM AS pko (READPAST) ON pko.IdShift = sh.Id AND pko.State = 0 AND pko.Flg_Bic = 0 AND pko.BuhSymbol_eko <> SPACE(0)
  WHERE Sh.ReportDate = ?f_CurDate AND Sh.IdDepartment = ?f_IdDepartmentKP
  ORDER BY Sh.Id DESC

ENDTEXT

IF RunSQL_Load('Shift_KP', @ar) <> 1
  MESSAGEBOX('Ошибка работы со справочником Смен ... ', 16, 'Смена не найдена.')
  RETURN -1
ENDIF

* BROWSE WINDOW brows

IF USED('Shift_KP') = .T. AND RECCOUNT('Shift_KP') > 0

  SELECT Shift_KP
  colvo_zap = RECCOUNT()
  GO TOP

  IF colvo_zap > 1

    BROWSE ;
    FIELDS Id, IdDepartment, Status, ReportDate, IdOperDate, OpDate, OpenDateDen ;
    WINDOW smena ;
    TITLE 'Вниманние! Доступно для обработки 2 смены. Поставьте стрелками курсор на нужную Вам смену. Для выхода нажмите ESC'

  ENDIF

  colvo_denCFT = TTOD(Shift_KP.OpDate) - TTOD(Shift_KP.OpenDateDen)

  OperDateCFT = Shift_KP.OpDate  && Дата открытого операционного дня КП
  OpenDateCFT = Shift_KP.OpenDate  && Дата открытия операционного дня КП
  DateSmenaCFT = TTOD(Shift_KP.ReportDate)  && Дата смены КП
  ShiftDateSmena = TTOD(Shift_KP.ShiftDate)  && Дата смены КП
  ReportDateSmena = TTOD(Shift_KP.ReportDate)  && Дата смены КП

  f_OpDateCFT = Shift_KP.OpDate  && Дата открытого операционного дня КП
  f_OpenDateCFT = Shift_KP.OpenDate  && Дата открытия операционного дня КП
  f_OpDate_DenCFT = Shift_KP.OpDate  && Дата открытого операционного дня КП

  IdOperDateCFT = Shift_KP.IdOperDate  && Id открытого опердня
  IdShiftCFT = Shift_KP.Id

ELSE
  =err('Внимание! За выбранную дату доступных смен в справочнике не обнаружено ... ')
  RETURN
ENDIF

* BROWSE WINDOW brows

* WAIT WINDOW 'colvo_denCFT = ' + ALLTRIM(STR(colvo_denCFT)) TIMEOUT 3
* WAIT WINDOW 'OperDateCFT = ' + DTOC(OperDateCFT) TIMEOUT 3
* WAIT WINDOW 'IdShiftCFT = ' + ALLTRIM(STR(IdShiftCFT)) TIMEOUT 3

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

f_State = 0

put_select = 1

DO WHILE .T.

  tex1 = 'Выбрать реестр платежей в ПОЛНОМ объеме'
  tex2 = 'Выбрать реестр платежей по ПИН клиента'
  tex3 = 'Выбрать документ по номеру СУМКИ'
  tex4 = 'Выбрать документ по номеру счета в Инкахран'
  tex5 = 'Выбрать документ по номеру счета в Банке'
  tex6 = 'Выбрать документ по сумме ДОКУМЕНТА'
  l_bar = 11
  =popup_big(tex1,tex2,tex3,tex4,tex5,tex6,tex7,tex8,l_bar)

  ACTIVATE POPUP vibor BAR put_select
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    put_select = IIF(BAR() <= 0, 1, BAR())

    DO CASE
      CASE put_select = 1  && Выбрать реестр платежей в ПОЛНОМ объеме

        TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT pko.Id, pko.IsDriver, pko.IdFilial, pko.FilialCode, pko.BudgetNum, pko.IdShift, shf.IdPart AS IdPartKP, pko.DateTransfer, pko.IdOperDate, pko.BuhSymbol_eko, pko.State, pko.Ref, pko.RefDtime, pko.Ccy, pko.Acc20, pko.LizNumCln,
          pko.TotalSumKP, pko.Kurs, pko.TotalSumKP_RUB, pko.Comment, pko.INNCln, pko.Fio_Kontroler, pko.Pass_Type, pko.Pass_numer, pko.Pass_data, pko.Pass_mesto, pko.Flg_Bic
          FROM PaketKasOrderATM AS pko (READPAST)
          INNER JOIN Shift AS shf (READPAST) ON shf.Id = pko.IdShift
          WHERE pko.IdShift = ?IdShiftCFT AND pko.State = ?f_State AND pko.Flg_Bic = 0 AND pko.BuhSymbol_eko <> SPACE(0)
          ORDER BY pko.BuhSymbol_eko, pko.Acc20, pko.LizNumCln, TotalSumKP

        ENDTEXT


      CASE put_select = 3  && Выбрать реестр платежей по ПИН клиента

        f_FilialCode = ''
*      f_FilialCode = '09-320'

        DO FORM Vibor_pin_client

        IF NOT LASTKEY() = 27 AND pr_lastkey_27 = .T.

          WAIT WINDOW 'ПИН клиента = ' + ALLTRIM(f_FilialCode) TIMEOUT 2

          TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT pko.Id, pko.IsDriver, pko.IdFilial, pko.FilialCode, pko.BudgetNum, pko.IdShift, shf.IdPart AS IdPartKP, pko.DateTransfer, pko.IdOperDate, pko.BuhSymbol_eko, pko.State, pko.Ref, pko.RefDtime, pko.Ccy, pko.Acc20, pko.LizNumCln,
            pko.TotalSumKP, pko.Kurs, pko.TotalSumKP_RUB, pko.Comment, pko.INNCln, pko.Fio_Kontroler, pko.Pass_Type, pko.Pass_numer, pko.Pass_data, pko.Pass_mesto, pko.Flg_Bic
            FROM PaketKasOrderATM AS pko (READPAST)
            INNER JOIN Shift AS shf (READPAST) ON shf.Id = pko.IdShift
            WHERE pko.IdShift = ?IdShiftCFT AND pko.State = ?f_State AND pko.Flg_Bic = 0 AND pko.FilialCode = ?f_FilialCode AND pko.BuhSymbol_eko <> SPACE(0)
            ORDER BY pko.BuhSymbol_eko, pko.Acc20, pko.LizNumCln, TotalSumKP

          ENDTEXT

        ELSE
          LOOP
        ENDIF


      CASE put_select = 5  && Выбрать документ по номеру СУМКИ

        f_BudgetNum = ''
*      f_BudgetNum = 'KM459763'

        DO FORM Vibor_sumka_client

        IF NOT LASTKEY() = 27 AND pr_lastkey_27 = .T.

          WAIT WINDOW 'Номер СУМКИ = ' + ALLTRIM(f_BudgetNum) TIMEOUT 2

          TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT pko.Id, pko.IsDriver, pko.IdFilial, pko.FilialCode, pko.BudgetNum, pko.IdShift, shf.IdPart AS IdPartKP, pko.DateTransfer, pko.IdOperDate, pko.BuhSymbol_eko, pko.State, pko.Ref, pko.RefDtime, pko.Ccy, pko.Acc20, pko.LizNumCln,
            pko.TotalSumKP, pko.Kurs, pko.TotalSumKP_RUB, pko.Comment, pko.INNCln, pko.Fio_Kontroler, pko.Pass_Type, pko.Pass_numer, pko.Pass_data, pko.Pass_mesto, pko.Flg_Bic
            FROM PaketKasOrderATM AS pko (READPAST)
            INNER JOIN Shift AS shf (READPAST) ON shf.Id = pko.IdShift
            WHERE pko.IdShift = ?IdShiftCFT AND pko.State = ?f_State AND pko.Flg_Bic = 0 AND pko.BudgetNum = ?f_BudgetNum AND pko.BuhSymbol_eko <> SPACE(0)
            ORDER BY pko.BuhSymbol_eko, pko.Acc20, pko.LizNumCln, TotalSumKP

          ENDTEXT

        ELSE
          LOOP
        ENDIF


      CASE put_select = 7  && Выбрать документ по номеру счета в Инкахран

        f_NumLiz = ''
*      f_NumLiz = '30109810655990000039'

        DO FORM Vibor_numliz_client

        IF NOT LASTKEY() = 27 AND pr_lastkey_27 = .T.

          WAIT WINDOW 'Номер счета в Инкахран = ' + ALLTRIM(f_NumLiz) TIMEOUT 2

          TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT pko.Id, pko.IsDriver, pko.IdFilial, pko.FilialCode, pko.BudgetNum, pko.IdShift, shf.IdPart AS IdPartKP, pko.DateTransfer, pko.IdOperDate, pko.BuhSymbol_eko, pko.State, pko.Ref, pko.RefDtime, pko.Ccy, pko.Acc20, pko.LizNumCln,
            pko.TotalSumKP, pko.Kurs, pko.TotalSumKP_RUB, pko.Comment, pko.INNCln, pko.Fio_Kontroler, pko.Pass_Type, pko.Pass_numer, pko.Pass_data, pko.Pass_mesto, pko.Flg_Bic
            FROM PaketKasOrderATM AS pko (READPAST)
            INNER JOIN Shift AS shf (READPAST) ON shf.Id = pko.IdShift
            WHERE pko.IdShift = ?IdShiftCFT AND pko.State = ?f_State AND pko.Flg_Bic = 0 AND pko.Acc20 = ?f_NumLiz AND pko.BuhSymbol_eko <> SPACE(0)
            ORDER BY pko.BuhSymbol_eko, pko.Acc20, pko.LizNumCln, TotalSumKP

          ENDTEXT

        ELSE
          LOOP
        ENDIF


      CASE put_select = 9  && Выбрать документ по номеру счета в Банке

        f_NumLiz = ''
*      f_NumLiz = '40702810487280000334'

        DO FORM Vibor_numliz_client

        IF NOT LASTKEY() = 27 AND pr_lastkey_27 = .T.

          WAIT WINDOW 'Номер счета в Банке = ' + ALLTRIM(f_NumLiz) TIMEOUT 2

          TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT pko.Id, pko.IsDriver, pko.IdFilial, pko.FilialCode, pko.BudgetNum, pko.IdShift, shf.IdPart AS IdPartKP, pko.DateTransfer, pko.IdOperDate, pko.BuhSymbol_eko, pko.State, pko.Ref, pko.RefDtime, pko.Ccy, pko.Acc20, pko.LizNumCln,
            pko.TotalSumKP, pko.Kurs, pko.TotalSumKP_RUB, pko.Comment, pko.INNCln, pko.Fio_Kontroler, pko.Pass_Type, pko.Pass_numer, pko.Pass_data, pko.Pass_mesto, pko.Flg_Bic
            FROM PaketKasOrderATM AS pko (READPAST)
            INNER JOIN Shift AS shf (READPAST) ON shf.Id = pko.IdShift
            WHERE pko.IdShift = ?IdShiftCFT AND pko.State = ?f_State AND pko.Flg_Bic = 0 AND pko.LizNumCln = ?f_NumLiz AND pko.BuhSymbol_eko <> SPACE(0)
            ORDER BY pko.BuhSymbol_eko, pko.Acc20, pko.LizNumCln, TotalSumKP

          ENDTEXT

        ELSE
          LOOP
        ENDIF


      CASE put_select = 11  && Выбрать документ по сумме ДОКУМЕНТА

        f_TotalSum = 0.00
*      f_TotalSum = 130000.00

        DO FORM Vibor_summa_client

        IF NOT LASTKEY() = 27 AND pr_lastkey_27 = .T.

          WAIT WINDOW 'Сумма документа = ' + ALLTRIM(TRANSFORM(f_TotalSum, '999 999 999 999.99')) TIMEOUT 2

          TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT pko.Id, pko.IsDriver, pko.IdFilial, pko.FilialCode, pko.BudgetNum, pko.IdShift, shf.IdPart AS IdPartKP, pko.DateTransfer, pko.IdOperDate, pko.BuhSymbol_eko, pko.State, pko.Ref, pko.RefDtime, pko.Ccy, pko.Acc20, pko.LizNumCln,
            pko.TotalSumKP, pko.Kurs, pko.TotalSumKP_RUB, pko.Comment, pko.INNCln, pko.Fio_Kontroler, pko.Pass_Type, pko.Pass_numer, pko.Pass_data, pko.Pass_mesto, pko.Flg_Bic
            FROM PaketKasOrderATM AS pko (READPAST)
            INNER JOIN Shift AS shf (READPAST) ON shf.Id = pko.IdShift
            WHERE pko.IdShift = ?IdShiftCFT AND pko.State = ?f_State AND pko.Flg_Bic = 0 AND pko.TotalSumKP = ?f_TotalSum AND pko.BuhSymbol_eko <> SPACE(0)
            ORDER BY pko.BuhSymbol_eko, pko.Acc20, pko.LizNumCln, TotalSumKP

          ENDTEXT

        ELSE
          LOOP
        ENDIF

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    IF RunSQL_Load('PaketKasOrderATM', @ar) <> 1
      MESSAGEBOX('Ошибка выборки данных из таблицы FROM PaketKasOrderATM', 16, 'Обратитесь к программистам по тел: 243, 303, 170, 256')
      DO exit_sql_server IN RunSQL_Load
      RETURN -1
    ENDIF

* BROWSE WINDOW brows

    IF ar[1,2] = 0
      IF USED('PaketKasOrderATM')
        USE IN PaketKasOrderATM
      ENDIF
      MESSAGEBOX('Сумок за дату ' + DTOC(new_data_odb) + ' и номером смены ' + ALLTRIM(STR(IdShiftKP)) + CHR(13) + '   доступных для зачисления НЕ ОБНАРУЖЕНО', 16, 'Обратитесь к начальнику смены Кассы Пересчета')
      DO exit_sql_server IN RunSQL_Load
      RETURN
    ENDIF

* BROWSE WINDOW brows

    IF USED('PaketKasOrderATM') = .T.

      SELECT PaketKasOrderATM
      colvo_zap = RECCOUNT()

* BROWSE WINDOW brows

      IF colvo_zap <> 0

        tex1 = 'Формирование данных для загрузки в ЦФТ с разбивкой по БУХГАЛТЕРСКОМУ СИМВОЛУ'
        tex2 = 'Формирование данных для загрузки в ЦФТ с разбивкой по НОМЕРУ СЧЁТА в ИНКАХРАН'
        l_bar = 3
        =popup_9(tex1,tex2,tex3,tex4,l_bar)
        ACTIVATE POPUP vibor
        RELEASE POPUPS vibor

        IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

          type_vivod_reestr = IIF(BAR() <= 0, 1, BAR())

          DO CASE
            CASE type_vivod_reestr = 1  && Формирование данных для загрузки в ЦФТ с разбивкой по БУХГАЛТЕРСКОМУ СИМВОЛУ

              DO export_reestr_buhsimvol IN PlatAtmToCFT


            CASE type_vivod_reestr = 3  && Формирование данных для загрузки в ЦФТ с разбивкой по НОМЕРУ СЧЁТА в ИНКАХРАН

              DO export_reestr_numliz IN PlatAtmToCFT

          ENDCASE

        ELSE
          EXIT
        ENDIF

      ELSE
        =soob('Внимание! В сформированной таблице PaketKasOrderATM из базы SQL Server записей не обнаружено ... ')
      ENDIF

    ELSE
      =soob('Внимание! Выборка данных из базы SQL Server PaketKasOrderATM закончилась аварийно ... ')
    ENDIF

  ELSE
    EXIT
  ENDIF

ENDDO

* ----------------------------------------------------- *
DO exit_sql_server IN RunSQL_Load
* ----------------------------------------------------- *

SELECT(sel)
RETURN


****************************************************************************** PROCEDURE export_reestr_buhsimvol *************************************************************************

PROCEDURE export_reestr_buhsimvol  && Формирование данных для загрузки в ЦФТ с разбивкой по БУХГАЛТЕРСКОМУ СИМВОЛУ

DO WHILE .T.

  tex1 = 'Только просмотр выгружаемых данных для загрузки в ЦФТ'
  tex2 = 'Только формирование текстового файла для загрузки в ЦФТ'
  l_bar = 3
  =popup_9(tex1,tex2,tex3,tex4,l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    put_vivod = IIF(BAR() <= 0, 1, BAR())

    SELECT DISTINCT BuhSymbol_eko AS BuhSymbol ;
      FROM PaketKasOrderATM ;
      INTO CURSOR Scan_BuhSymbol ;
      ORDER BY BuhSymbol_eko

* BROWSE WINDOW brows

    IF _TALLY <> 0  && Сначала сканируем данные по полю BuhSymbol, данные по разным Бухгалтерским символам должны находиться в разных выходных файлах

      SELECT Scan_BuhSymbol
      GO TOP

      SCAN

        num_BuhSymbol = ALLTRIM(Scan_BuhSymbol.BuhSymbol)

        time_exp_dir = LEFT(TIME(), 2) + SUBSTR(TIME(), 4, 2) + RIGHT(TIME(), 2)

        dir_login_time = ALLTRIM(dir_login) + '_' + ALLTRIM(time_exp_dir)
        put_login_time = ALLTRIM(dir_login) + '_' + ALLTRIM(time_exp_dir) + '\'

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        IF put_vivod = 3  && Только формирование текстового файла для загрузки в ЦФТ

          put_select = MESSAGEBOX("Вы действительно желаете сформировать" + CHR(13) + ;
            "выходной текстовый файл для ЦФТ?"+ CHR(13) + 'Выгрузка по бухгалтерскому символу - ' + ALLTRIM(num_BuhSymbol), 32+16+1, 'Сделайте правильный и осознанный выбор')

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          IF put_select = 1  && Файлы для ЦФТ формируем

            =soob('Файлы для ЦФТ создаются в папку с именем "Логин + текущее время" - ' + ALLTRIM(put_tmp_CftAtm) + ALLTRIM(dir_login_time))

            IF DIRECTORY(ALLTRIM(put_tmp_CftAtm) + ALLTRIM(put_login_time), 1) = .F.

              old_path = SYS(5) + SYS(2003)

              SET DEFAULT TO (put_tmp_CftAtm) && Переходим в каталог временных файлов
              MKDIR (put_tmp_CftAtm) + (dir_login_time) && Создаем каталог

              SET DEFAULT TO (old_path)

            ENDIF

          ELSE
            LOOP
          ENDIF

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        IF put_vivod = 3  && Только формирование текстового файла для загрузки в ЦФТ

          IF DIRECTORY(ALLTRIM(put_tmp_CftAtm) + ALLTRIM(put_login_time), 1) = .F.

            =err('Папка, необходимая для выгрузки файлов ЦФТ по пути - ' + ALLTRIM(put_tmp_CftAtm) + ALLTRIM(dir_login_time) + ' НЕ СОЗДАНА!!!')
            RETURN

          ENDIF

        ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        SELECT pko.Id, pko.IsDriver, pko.DateTransfer, pko.IdOperDate, pko.IdShift, pko.IdPartKP, pko.State, pko.Ref, pko.RefDtime, pko.BuhSymbol_eko, pko.Acc20, pko.LizNumCln,;
          pko.TotalSumKP, pko.Kurs, pko.TotalSumKP_RUB, pko.IdFilial, pko.FilialCode, pko.BudgetNum,;
          pko.INNCln, pko.Ccy, pko.Comment, pko.Fio_Kontroler, pko.Pass_Type, pko.Pass_numer, pko.Pass_data, pko.Pass_mesto, pko.Flg_Bic ;
          FROM PaketKasOrderATM AS pko ;
          WHERE ALLTRIM(pko.BuhSymbol_eko) == ALLTRIM(num_BuhSymbol) ;
          INTO CURSOR Plat_PaketKasOrderATM ;
          ORDER BY pko.BuhSymbol_eko, pko.Acc20, pko.TotalSumKP

* BROWSE WINDOW brows

        IF USED('Plat_PaketKasOrderATM') = .T.

          SELECT Plat_PaketKasOrderATM

          colvo_zap = RECCOUNT()

          IF colvo_zap <> 0

            date_exp_file = LEFT(DTOC(DATE()), 2) + SUBSTR(DTOC(DATE()), 4, 2) + RIGHT(DTOC(DATE()), 2)
            time_exp_file = LEFT(TIME(), 2) + SUBSTR(TIME(), 4, 2) + RIGHT(TIME(), 2)

            exp_file_cft = 'FORMDOC_' + ALLTRIM(date_exp_file) + '_' + ALLTRIM(time_exp_file) + '_приход_' + ALLTRIM(num_BuhSymbol) + '.txt'

            exp_file_title = 'Реестр платежей: ' + ALLTRIM(date_exp_file) + '_' + ALLTRIM(time_exp_file) + '_приход_' + ALLTRIM(num_BuhSymbol)

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            SELECT Plat_PaketKasOrderATM
            GO TOP

            DO CASE
              CASE put_vivod = 1  && Только просмотр выгружаемых данных для загрузки в ЦФТ

                BROWSE FIELDS ;
                  DateTransfer:H = 'Дата',;
                  BuhSymbol_eko:H = 'Символ ЕКО',;
                  Acc20:H = 'Счет клиента в Инкахран',;
                  TotalSumKP:H = 'Сумма проводки':P = '99 999 999.99',;
                  Kurs:H = 'Курс':P = '99.9999',;
                  TotalSumKP_RUB:H = 'Сумма проводки RUB':P = '999 999 999.99',;
                  FilialCode:H = 'ПИН Код',;
                  BudgetNum:H = 'Номер сумки',;
                  Comment:H = 'Комментарий',;
                  Pass_Type:H = 'Код документа',;
                  Pass_Numer:H = 'Серия и номер',;
                  Pass_Data:H = 'Дата выдачи',;
                  Pass_Mesto:H = 'Место выдачи' ;
                  WINDOW BROWS TITLE ALLTRIM(exp_file_title)

*                  LizNumCln:H = 'Счет клиента в Банке',;

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

              CASE put_vivod = 3  && Только формирование текстового файла для загрузки в ЦФТ


                DO CASE
                  CASE put_select = 1  && Файлы для ЦФТ формируем

                    SELECT Plat_PaketKasOrderATM
                    GO TOP

                    SCAN  && Установка статуса "Платеж Выгружен" для отправки платежа в ЦФТ

                      f_Id = Plat_PaketKasOrderATM.ID
                      f_BuhSymbol = Plat_PaketKasOrderATM.BuhSymbol_eko

                      TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                        UPDATE PaketKasOrderATM SET
                        Flg_Bic = 1, File_Name_Txt = ?exp_file_cft
                        WHERE Id = ?f_Id AND State = ?f_State AND Flg_Bic = 0 AND BuhSymbol_eko = ?f_BuhSymbol

                      ENDTEXT

                      IF RunSQL_Load('', @ar) <> 1
                        MESSAGEBOX('Ошибка установки статуса "Платеж Выгружен"' + CHR(13) + 'Для отправки платежа в ЦФТ', 16, 'Обратитесь к программистам по тел: 243, 303, 170, 256')
                        RETURN -1
                      ENDIF

                      SELECT Plat_PaketKasOrderATM
                    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
                    DO zapis_cft IN PlatAtmToCFT && Процедура записи данных кассового ордера в текстовый файл, в дальнейшем используемую для передачи в ЦФТ пакетных кассовых ордеров.
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *


                  CASE put_select = 2  && Файлы для ЦФТ не формируем

                    RETURN

                ENDCASE

            ENDCASE

          ELSE
            =soob('Внимание! В сформированной таблице Plat_PaketKasOrderATM из базы SQL Server записей не обнаружено ... ')
          ENDIF

        ELSE
          =soob('Внимание! Выборка данных из базы SQL Server Plat_PaketKasOrderATM закончилась аварийно ... ')
        ENDIF

        SELECT Scan_BuhSymbol
      ENDSCAN

    ENDIF

  ELSE
    EXIT
  ENDIF

  IF put_vivod = 3  && Только формирование текстового файла для загрузки в ЦФТ
    EXIT
  ENDIF

ENDDO

RETURN


****************************************************************************** PROCEDURE export_reestr_numliz ****************************************************************************

PROCEDURE export_reestr_numliz  && Формирование данных для загрузки в ЦФТ с разбивкой по НОМЕРУ СЧЁТА в ИНКАХРАН

DO WHILE .T.

  tex1 = 'Только просмотр выгружаемых данных для загрузки в ЦФТ'
  tex2 = 'Только формирование текстового файла для загрузки в ЦФТ'
  l_bar = 3
  =popup_9(tex1,tex2,tex3,tex4,l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    put_vivod = IIF(BAR() <= 0, 1, BAR())

    SELECT DISTINCT BuhSymbol_eko AS BuhSymbol ;
      FROM PaketKasOrderATM ;
      INTO CURSOR Scan_BuhSymbol ;
      ORDER BY BuhSymbol_eko

    IF _TALLY <> 0  && Сначала сканируем данные по полю BuhSymbol_eko, данные по разным Бухгалтерским символам должны находиться в разных выходных файлах

      SELECT Scan_BuhSymbol
      GO TOP

      SCAN

        num_BuhSymbol = ALLTRIM(Scan_BuhSymbol.BuhSymbol)

        SELECT DISTINCT BuhSymbol_eko AS BuhSymbol, Acc20 ;
          FROM PaketKasOrderATM ;
          WHERE ALLTRIM(BuhSymbol_eko) == ALLTRIM(num_BuhSymbol) ;
          INTO CURSOR Scan_LizNumCln ;
          ORDER BY BuhSymbol_eko, Acc20

        IF _TALLY <> 0  && Теперь сканируем данные по полю Acc20, данные по разным Бухгалтерским символам и разным счетам по дебету должны находиться в разных выходных файлах

          time_exp_dir = LEFT(TIME(), 2) + SUBSTR(TIME(), 4, 2) + RIGHT(TIME(), 2)

          dir_login_time = ALLTRIM(dir_login) + '_' + ALLTRIM(time_exp_dir)
          put_login_time = ALLTRIM(dir_login) + '_' + ALLTRIM(time_exp_dir) + '\'

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          IF put_vivod = 3  && Только формирование текстового файла для загрузки в ЦФТ

            put_select = MESSAGEBOX("Вы действительно желаете сформировать" + CHR(13) + ;
              "выходной текстовый файл для ЦФТ?"+ CHR(13) + 'Выгрузка по бухгалтерскому символу - ' + ALLTRIM(num_BuhSymbol), 32+16+1, "Сделайте правильный и осознанный выбор")

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

            IF put_select = 1  && Файлы для ЦФТ формируем

              =soob('Файлы для ЦФТ создаются в папку с именем "Логин + текущее время" - ' + ALLTRIM(put_tmp_CftAtm) + ALLTRIM(dir_login_time))

              IF DIRECTORY(ALLTRIM(put_tmp_CftAtm) + ALLTRIM(put_login_time), 1) = .F.

                old_path = SYS(5) + SYS(2003)

                SET DEFAULT TO (put_tmp_CftAtm) && Переходим в каталог временных файлов
                MKDIR (put_tmp_CftAtm) + (dir_login_time) && Создаем каталог

                SET DEFAULT TO (old_path)

              ENDIF

            ELSE
              LOOP
            ENDIF

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          IF put_vivod = 3  && Только формирование текстового файла для загрузки в ЦФТ

            IF DIRECTORY(ALLTRIM(put_tmp_CftAtm) + ALLTRIM(put_login_time), 1) = .F.

              =err('Папка, необходимая для выгрузки файлов ЦФТ по пути - ' + ALLTRIM(put_tmp_CftAtm) + ALLTRIM(dir_login_time) + ' НЕ СОЗДАНА!!!')
              RETURN

            ENDIF

          ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

          SELECT Scan_LizNumCln
          GO TOP

          SCAN

            num_Acc20 = ALLTRIM(Scan_LizNumCln.Acc20)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

            SELECT pko.Id, pko.IsDriver, pko.DateTransfer, pko.IdOperDate, pko.IdShift, pko.IdPartKP, pko.State, pko.Ref, pko.RefDtime, pko.BuhSymbol_eko, pko.Acc20, pko.LizNumCln, pko.TotalSumKP, pko.Kurs, pko.TotalSumKP_RUB, pko.IdFilial, pko.FilialCode, pko.BudgetNum,;
              pko.INNCln, pko.Ccy, pko.Comment, pko.Fio_Kontroler, pko.Pass_Type, pko.Pass_numer, pko.Pass_data, pko.Pass_mesto, pko.Flg_Bic ;
              FROM PaketKasOrderATM AS pko ;
              WHERE pko.State = f_State AND pko.Flg_Bic = 0 AND ALLTRIM(pko.BuhSymbol_eko) == ALLTRIM(num_BuhSymbol) AND ALLTRIM(Acc20) == ALLTRIM(num_Acc20) ;
              INTO CURSOR Plat_PaketKasOrderATM ;
              ORDER BY pko.BuhSymbol_eko, pko.Acc20, pko.TotalSumKP

* BROWSE WINDOW brows

            IF USED('Plat_PaketKasOrderATM') = .T.

              SELECT Plat_PaketKasOrderATM
              colvo_zap = RECCOUNT()

* BROWSE WINDOW brows

              IF colvo_zap <> 0

                date_exp_file = LEFT(DTOC(DATE()), 2) + SUBSTR(DTOC(DATE()), 4, 2) + RIGHT(DTOC(DATE()), 2)
                time_exp_file = LEFT(TIME(), 2) + SUBSTR(TIME(), 4, 2) + RIGHT(TIME(), 2)

                exp_file_cft = 'FORMDOC_' + ALLTRIM(date_exp_file) + '_' + ALLTRIM(time_exp_file) + '_' + ALLTRIM(num_BuhSymbol) + '_приход_' + ALLTRIM(num_Acc20) + '.txt'

                exp_file_title = 'Реестр платежей: ' + ALLTRIM(date_exp_file) + '_' + ALLTRIM(time_exp_file) + '_' + ALLTRIM(num_BuhSymbol) + '_приход_' + ALLTRIM(num_Acc20)

                SELECT Plat_PaketKasOrderATM
                GO TOP

                DO CASE
                  CASE put_vivod = 1  && Только просмотр выгружаемых данных для загрузки в ЦФТ

                    BROWSE FIELDS ;
                      DateTransfer:H = 'Дата',;
                      BuhSymbol_eko:H = 'Символ ЕКО',;
                      Acc20:H = 'Счет клиента в Инкахран',;
                      TotalSumKP:H = 'Сумма проводки':P = '99 999 999.99',;
                      Kurs:H = 'Курс':P = '99.9999',;
                      TotalSumKP_RUB:H = 'Сумма проводки RUB':P = '999 999 999.99',;
                      FilialCode:H = 'ПИН Код',;
                      BudgetNum:H = 'Номер сумки',;
                      Comment:H = 'Комментарий',;
                      Fio_Kontroler:H = 'ФИО Контролера',;
                      Pass_Type:H = 'Код документа',;
                      Pass_Numer:H = 'Серия и номер',;
                      Pass_Data:H = 'Дата выдачи',;
                      Pass_Mesto:H = 'Место выдачи' ;
                      WINDOW BROWS TITLE ALLTRIM(exp_file_title)

*                      LizNumCln:H = 'Счет клиента в Банке',;

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                  CASE put_vivod = 3  && Только формирование текстового файла для загрузки в ЦФТ


                    DO CASE
                      CASE put_select = 1  && Файлы для ЦФТ формируем

                        SELECT Plat_PaketKasOrderATM
                        GO TOP

                        SCAN  && Установка статуса "Платеж Выгружен" для отправки платежа в ЦФТ

                          f_Id = Plat_PaketKasOrderATM.ID
                          f_BuhSymbol = Plat_PaketKasOrderATM.BuhSymbol_eko

                          TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                            UPDATE PaketKasOrderATM SET
                            Flg_Bic = 1, File_Name_Txt = ?exp_file_cft
                            WHERE Id = ?f_Id AND State = ?f_State AND Flg_Bic = 0 AND BuhSymbol_eko = ?f_BuhSymbol

                          ENDTEXT

                          IF RunSQL_Load('', @ar) <> 1
                            MESSAGEBOX('Ошибка установки статуса "Платеж Выгружен"' + CHR(13) + 'Для отправки платежа в ЦФТ', 16, 'Обратитесь к программистам по тел: 243, 303, 170, 256')
                            RETURN -1
                          ENDIF

                          SELECT Plat_PaketKasOrderATM
                        ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
                        DO zapis_cft IN PlatAtmToCFT && Процедура записи данных кассового ордера в текстовый файл, в дальнейшем используемую для передачи в ЦФТ пакетных кассовых ордеров.
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *


                      CASE put_select = 2  && Файлы для ЦФТ не формируем

                        RETURN

                    ENDCASE

                ENDCASE

              ELSE
                =soob('Внимание! В сформированной таблице Plat_PaketKasOrderATM из базы SQL Server записей не обнаружено ... ')
              ENDIF

            ELSE
              =soob('Внимание! Выборка данных из базы SQL Server Plat_PaketKasOrderATM закончилась аварийно ... ')
            ENDIF

          ENDSCAN

        ENDIF

      ENDSCAN

    ENDIF

  ELSE
    EXIT
  ENDIF

  IF put_vivod = 3  && Только формирование текстового файла для загрузки в ЦФТ
    EXIT
  ENDIF

ENDDO

RETURN


************************************************************************************** PROCEDURE zapis_cft ***********************************************************************************************************

PROCEDURE zapis_cft

SET MARK TO '/' && тип разделителя в формате даты

SELECT Plat_PaketKasOrderATM
GO TOP

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------- Основное операционное время, в КП смена УТРО --------------------------------------------------------------------------------------------------------------- *

Inkahran_Kassa_Acc_RUB = ALLTRIM(GetSqlGlob_Mesto('InkhranAccRUB', f_MestoCode))  && Формирование счета основной кассы в зависимости от площадки
Inkahran_Kassa_Acc_USD = ALLTRIM(GetSqlGlob_Mesto('InkhranAccUSD', f_MestoCode))  && Формирование счета основной кассы в зависимости от площадки
Inkahran_Kassa_Acc_EUR = ALLTRIM(GetSqlGlob_Mesto('InkhranAccEUR', f_MestoCode))  && Формирование счета основной кассы в зависимости от площадки

* ---------------------------------------------------------------------------------- После операционное время, в КП смена ДЕНЬ и НОЧЬ ----------------------------------------------------------------------------------------------------------- *

Inkahran_Kassa_Acc_Pos_RUB = ALLTRIM(GetSqlGlob_Mesto('InkhranAccPosRUB', f_MestoCode))  && Формирование счета после операционной кассы в зависимости от площадки
Inkahran_Kassa_Acc_Pos_USD = ALLTRIM(GetSqlGlob_Mesto('InkhranAccPosUSD', f_MestoCode))  && Формирование счета после операционной кассы в зависимости от площадки
Inkahran_Kassa_Acc_Pos_EUR = ALLTRIM(GetSqlGlob_Mesto('InkhranAccPosEUR', f_MestoCode))  && Формирование счета после операционной кассы в зависимости от площадки

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW 'Inkahran_Kassa_Acc_RUB = ' + ALLTRIM(Inkahran_Kassa_Acc_RUB) TIMEOUT 3
* WAIT WINDOW 'Inkahran_Kassa_Acc_USD = ' + ALLTRIM(Inkahran_Kassa_Acc_USD) TIMEOUT 3
* WAIT WINDOW 'Inkahran_Kassa_Acc_EUR = ' + ALLTRIM(Inkahran_Kassa_Acc_EUR) TIMEOUT 3

* WAIT WINDOW 'Inkahran_Kassa_Acc_Pos_RUB = ' + ALLTRIM(Inkahran_Kassa_Acc_Pos_RUB) TIMEOUT 3
* WAIT WINDOW 'Inkahran_Kassa_Acc_Pos_USD = ' + ALLTRIM(Inkahran_Kassa_Acc_Pos_USD) TIMEOUT 3
* WAIT WINDOW 'Inkahran_Kassa_Acc_Pos_EUR = ' + ALLTRIM(Inkahran_Kassa_Acc_Pos_EUR) TIMEOUT 3

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

ACTIVATE WINDOW poisk
@ WROWS()/3-0.2,3 SAY PADC('Начато формирование данных для пакетных кассовых платежей в "ЦФТ" ... ',WCOLS())

* ----------------------------------------------------------------------------------------------- Создаем временную таблицу Kassa_CFT --------------------------------------------------------------------------------------------------------------------------------------- *

CREATE CURSOR Kassa_CFT ;
  (ICODE CHAR(25),;  && Уникальный код документа
FILIAL CHAR(3),;  && Код филиала из справочника "Филиалы"
DEPART CHAR(20),;  && Код департамента из справочника "Справочник подразделений"
ACC_DT CHAR(25),;  && Счет плательщика, счет по дебету
ACC_KT CHAR(25),;  && Счет получателя, счет по кредиту
DATE_DOC DATE,;  && Дата документа
DATE_PROV DATE,;  && Дата проводки
NAZN CHAR(250),;  && Назначение платежа
PRIORITET INTEGER(4),;  && Приоритет, допустимые значения:1,2,3,4,5,6
SUM Numeric(17, 2),;  && Сумма дебета
SUM_PO Numeric(17, 2),;  && Сумма кредита
SUM_NT Numeric(17, 2),;  && Сумма национального покрытия
SUM_KSPL CHAR(250),;  && Сумма документа в разбиении по символам кассового плана в специальном формате: символ кассового плана, сумма, забалансовый символ касс.плана
VID_DOC CHAR(16),;  && Код из справочника "Наименования платежных документов" КАСС_ПРИХОД_ОРД кассовый приходный ордер, БЕЗН_МЕМО_ОРД - мемориальный ордер; ВНЕБ_ПРИХ_РАСХ - внебалансовый мемориальный ордер
KL_DT_INTERNAL CHAR(5),;  && Признак - внутренний клиент или внешний, в нашем случае всегда True.
KL_DT_CLIENT CHAR(25),;  && Клиент- владелец
KL_DT_ACC CHAR(25),;  && Счет клиента
KL_KT_INTERNAL CHAR(5),;  && Признак - внутренний клиент или внешний, в нашем случае всегда True.
KL_KT_CLIENT CHAR(25),;  && Клиент- владелец
KL_KT_ACC CHAR(25),;  && Счет получателя
REQ_MAN_CLIENT CHAR(25),;  && Внутренний клиент. Заполняется для кассового документа.
REQ_MAN_NAME CHAR(160),;  && Наименование  клиента- вносителя/получателя средств. Заполняется для кассового документа.
REQ_MAN_COUNTRY CHAR(100),;  && Код  страны из справочника "Общероссийский классификатор стран мира"
REQ_MAN_ADDRESS CHAR(250),;  && Адрес строкой клиента- вносителя/получателя средств. Заполняется для кассового документа.
REQ_MAN_REZIDENT CHAR(5),;  && Свойство резидентности клиента- вносителя/получателя средств из удостоверения. Заполняется для кассового документа, в нашем случае всегда True.
REQ_MAN_DOC_BIRTH_PLACE CHAR(250),;  && Место рождения клиента- вносителя/получателя средств из удостоверения. Заполняется для кассового документа.
REQ_MAN_DOC_DATE DATE,;  && Дата выдачи удостоверения клиента- вносителя/получателя средств. Заполняется для кассового документа.
REQ_MAN_DOC_TYPE CHAR(50),;  && Тип удостоверения (строкой) клиента- вносителя/получателя средств. Заполняется для кассового документа.
REQ_MAN_DOC_SER CHAR(16),;  && Серия удостоверения (строкой) клиента- вносителя/получателя средств. Заполняется для кассового документа.
REQ_MAN_DOC_NUM CHAR(16),;  && Номер удостоверения клиента- вносителя/получателя средств. Заполняется для кассового документа.
REQ_MAN_DOC_WHO CHAR(100),;  && Кем выдано удостоверение клиента- вносителя/получателя средств. Заполняется для кассового документа.
REQ_MAN_DOC_DEPART_CODE CHAR(20))  && Код подразделения, выдавшего удостоверение клиенту- вносителю/получателю средств. Заполняется для кассового документа.

* BROWSE WINDOW brows

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

old_path = SYS(5) + SYS(2003)

* WAIT WINDOW 'put_tmp_CftAtm - ' + ALLTRIM(put_tmp_CftAtm) TIMEOUT 5

SET TEXTMERGE ON
SET TEXTMERGE NOSHOW

IF FILE((put_tmp_CftAtm) + (put_login_time) + (exp_file_cft))
  DELETE FILE (put_tmp_CftAtm) + (put_login_time) + (exp_file_cft)
ENDIF

_TEXT = FCREATE((put_tmp_CftAtm) + (put_login_time) + (exp_file_cft))

IF _TEXT < 1
  =err('Внимание! Ошибка при создании файла данных для пакетных кассовых платежей в "ЦФТ" ... ')
  =FCLOSE(_TEXT)
ELSE

* ----------------------------------------------------------------------------------------------------------------------- Формирование строк детализации --------------------------------------------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW 'Дата операционного дня - ' + DT(skip_data_odb) + '  Дата открытой смены в ПРК на Нагорной - ' + DT(new_data_odb) TIMEOUT 2

  delta_data = 1

  delta_data = skip_data_odb - new_data_odb  && Вычисляем количество дней между датой опердня и датой смены из поля ReportDate

  DO CASE
    CASE delta_data < 10
      delta_data_cft = ALLTRIM(STR(delta_data))

    CASE delta_data = 10
      delta_data_cft = 'a'

    CASE delta_data = 11
      delta_data_cft = 'b'

    CASE delta_data = 12
      delta_data_cft = 'c'

    CASE delta_data = 13
      delta_data_cft = 'd'

    CASE delta_data = 14
      delta_data_cft = 'e'

    CASE delta_data = 15
      delta_data_cft = 'f'

  ENDCASE

  new_data_odb_cft = SUBSTR(DTOC(new_data_odb), 1, 2) + SUBSTR(DTOC(new_data_odb), 4, 2) + SUBSTR(DTOC(new_data_odb), 9, 2)  && Приводим дату к формату ddmmgg

* WAIT WINDOW 'Дата операционного дня - ' + DT(skip_data_odb) + '  Дата открытой смены в ПРК на Нагорной - ' + DT(new_data_odb) + '  new_data_odb_cft - ' + ALLTRIM(new_data_odb_cft) TIMEOUT 5

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW 'colvo_denCFT = ' + ALLTRIM(STR(colvo_denCFT)) TIMEOUT 3

  SELECT Plat_PaketKasOrderATM
  GO TOP

  SCAN

    SELECT Kassa_CFT
    APPEND BLANK

*  ddmmyySXXFNNNNN - длинна блока 15 символов - вариант № 1, где
*  ddmmyy - дата смены из справочника смен (то есть дата фактической инкассации), например 080317
*  S - кол-во дней между датой фактической инкассации и оперднём зачисления, например 0, 1, 2, …, 9, A (это 10), B (это 11), C (это 12), ….

* XX - '12' - ПКО разгрузка и пересчет остатков в кассетах в КП - АРМ Транзит УР RUB
* XX - '13' - ПКО разгрузка и пересчет остатков в кассетах в КП - АРМ Транзит УР USD
* XX - '14' - ПКО разгрузка и пересчет остатков в кассетах в КП - АРМ Транзит УР EUR

* XX - '15' - ПКО разгрузка и пересчет остатков в кассетах в КП - АРМ Транзит УР RUB (для двух одинаковых номеров АТМ по разгрузке кассет)
* XX - '16' - ПКО разгрузка и пересчет остатков в кассетах в КП - АРМ Транзит УР USD (для двух одинаковых номеров АТМ по разгрузке кассет)
* XX - '17' - ПКО разгрузка и пересчет остатков в кассетах в КП - АРМ Транзит УР EUR (для двух одинаковых номеров АТМ по разгрузке кассет)

*  F - код подразделения: '0' - код Каспера для Москвы, '4' - код Каспера для Питера

*  NNNNN - пять правых символов из поля - первичного ключа таблицы в Каспер, где содержится первичный документ.

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    DO CASE
      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'RUB' AND ALLTRIM(STR(Plat_PaketKasOrderATM.IsDriver)) == '0'

        REPLACE Kassa_CFT.ICODE WITH ALLTRIM(new_data_odb_cft) + ALLTRIM(delta_data_cft) + '12' + '0' + RIGHT(ALLTRIM(STR(Plat_PaketKasOrderATM.Id)), 5)   && Уникальный код документа

      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'RUB' AND ALLTRIM(STR(Plat_PaketKasOrderATM.IsDriver)) == '1'

        REPLACE Kassa_CFT.ICODE WITH ALLTRIM(new_data_odb_cft) + ALLTRIM(delta_data_cft) + '15' + '0' + RIGHT(ALLTRIM(STR(Plat_PaketKasOrderATM.Id)), 5)   && Уникальный код документа


      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'USD' AND ALLTRIM(STR(Plat_PaketKasOrderATM.IsDriver)) == '0'

        REPLACE Kassa_CFT.ICODE WITH ALLTRIM(new_data_odb_cft) + ALLTRIM(delta_data_cft) + '13' + '0' + RIGHT(ALLTRIM(STR(Plat_PaketKasOrderATM.Id)), 5)   && Уникальный код документа

      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'USD' AND ALLTRIM(STR(Plat_PaketKasOrderATM.IsDriver)) == '1'

        REPLACE Kassa_CFT.ICODE WITH ALLTRIM(new_data_odb_cft) + ALLTRIM(delta_data_cft) + '16' + '0' + RIGHT(ALLTRIM(STR(Plat_PaketKasOrderATM.Id)), 5)   && Уникальный код документа


      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'EUR' AND ALLTRIM(STR(Plat_PaketKasOrderATM.IsDriver)) == '0'

        REPLACE Kassa_CFT.ICODE WITH ALLTRIM(new_data_odb_cft) + ALLTRIM(delta_data_cft) + '14' + '0' + RIGHT(ALLTRIM(STR(Plat_PaketKasOrderATM.Id)), 5)   && Уникальный код документа

      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'EUR' AND ALLTRIM(STR(Plat_PaketKasOrderATM.IsDriver)) == '1'

        REPLACE Kassa_CFT.ICODE WITH ALLTRIM(new_data_odb_cft) + ALLTRIM(delta_data_cft) + '17' + '0' + RIGHT(ALLTRIM(STR(Plat_PaketKasOrderATM.Id)), 5)   && Уникальный код документа


    ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    REPLACE Kassa_CFT.FILIAL WITH PADR(ALLTRIM(kod_filial), 3)  && Код филиала из справочника "Филиалы", определяется в головном модуле.

    REPLACE Kassa_CFT.DEPART WITH PADR(ALLTRIM(kod_depart), 20)  && Код департамента из справочника "Подразделения", определяется в головном модуле.

* ------------------------------------------ Формирование по дебету счета кассы: либо счет операционной кассы, либо счет после операционной кассы -------------------------------------------------------------------------- *

*** Inkahran_Kassa_Acc - счет операционной кассы
*** Inkahran_Kassa_Acc_Pos_RUB - счет после операционной кассы

* WAIT WINDOW 'colvo_denCFT = ' + ALLTRIM(STR(colvo_denCFT)) TIMEOUT 3
* WAIT WINDOW 'OperDateCFT = ' + TTOC(OperDateCFT) TIMEOUT 3

    DO CASE
      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'RUB'

        DO CASE
          CASE colvo_denCFT = 1 AND DATE() <= TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт = 1 дню. Схема рабочего дня.

            DO CASE
              CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_RUB), 25)  && Счет плательщика, счет по дебету


              CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_RUB), 25)  && Счет плательщика, счет по дебету


              CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_RUB), 25)  && Счет плательщика, счет по дебету

            ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          CASE colvo_denCFT > 1 AND DATE() < TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня.

            REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_RUB), 25)  && Счет плательщика, счет по дебету


          CASE colvo_denCFT > 1 AND DATE() = TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня. Утро первого рабочего дня

            DO CASE
              CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_RUB), 25)  && Счет плательщика, счет по дебету


              CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_RUB), 25)  && Счет плательщика, счет по дебету


              CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_RUB), 25)  && Счет плательщика, счет по дебету

            ENDCASE


        ENDCASE

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'USD'

        DO CASE
          CASE colvo_denCFT = 1 AND DATE() <= TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт = 1 дню. Схема рабочего дня.

            DO CASE
              CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_USD), 25)  && Счет плательщика, счет по дебету


              CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_USD), 25)  && Счет плательщика, счет по дебету


              CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_USD), 25)  && Счет плательщика, счет по дебету

            ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          CASE colvo_denCFT > 1 AND DATE() < TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня.

            REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_USD), 25)  && Счет плательщика, счет по дебету


          CASE colvo_denCFT > 1 AND DATE() = TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня. Утро первого рабочего дня

            DO CASE
              CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_USD), 25)  && Счет плательщика, счет по дебету


              CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_USD), 25)  && Счет плательщика, счет по дебету


              CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_USD), 25)  && Счет плательщика, счет по дебету

            ENDCASE


        ENDCASE

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'EUR'

        DO CASE
          CASE colvo_denCFT = 1 AND DATE() <= TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт = 1 дню. Схема рабочего дня.

            DO CASE
              CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_EUR), 25)  && Счет плательщика, счет по дебету


              CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_EUR), 25)  && Счет плательщика, счет по дебету


              CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_EUR), 25)  && Счет плательщика, счет по дебету

            ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          CASE colvo_denCFT > 1 AND DATE() < TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня.

            REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_EUR), 25)  && Счет плательщика, счет по дебету


          CASE colvo_denCFT > 1 AND DATE() = TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня. Утро первого рабочего дня

            DO CASE
              CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_EUR), 25)  && Счет плательщика, счет по дебету


              CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_EUR), 25)  && Счет плательщика, счет по дебету


              CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

                REPLACE Kassa_CFT.ACC_DT WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_EUR), 25)  && Счет плательщика, счет по дебету

            ENDCASE


        ENDCASE


    ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    DO CASE
      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'RUB'

        REPLACE Kassa_CFT.ACC_KT WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.Acc20), 25)  && Счет получателя, счет по кредиту

      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'USD'

        REPLACE Kassa_CFT.ACC_KT WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.Acc20), 25)  && Счет получателя, счет по кредиту

      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'EUR'

        REPLACE Kassa_CFT.ACC_KT WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.Acc20), 25)  && Счет получателя, счет по кредиту

    ENDCASE

* --------------------------------------------------------- Формирование значения поля "Дата документа" ------------------------------------------------------------------------------------------------------------------------------------------------ *

*  ShiftDateSmena = TTOD(CursOpShift_2.ShiftDate)  && Дата открытой смены - поле ShiftDate
*  ReportDateSmena = TTOD(CursOpShift_2.ReportDate)  && Дата открытой смены - поле ReportDate
*  OperDate_KP = TTOD(Shift_KP.OpDate)  && Дата открытого опердня
*  DateSmenaKP = TTOD(Shift_KP.ReportDate) + 1  && Дата смены КП

*  WAIT WINDOW 'colvo_denCFT = ' + ALLTRIM(STR(colvo_denCFT)) TIMEOUT 5
*  WAIT WINDOW 'f_OpDate = ' + TTOC(f_OpDate) TIMEOUT 5
*  WAIT WINDOW 'OperDateCFT = ' + TTOC(OperDateCFT) TIMEOUT 2
*  WAIT WINDOW 'ShiftDateSmena = ' + DTOC(ShiftDateSmena) TIMEOUT 2
*  WAIT WINDOW 'ReportDateSmena = ' + DTOC(ReportDateSmena) TIMEOUT 2
*  WAIT WINDOW 'Plat_PaketKasOrder_IdPartKP = ' + ALLTRIM(STR(Plat_PaketKasOrderATM.IdPartKP)) TIMEOUT 5

    DO CASE
      CASE colvo_denCFT = 1 AND DATE() <= TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт = 1 дню. Схема рабочего дня.

        DO CASE
          CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

            REPLACE Kassa_CFT.DATE_DOC WITH DATE()  && Дата документа, пишется Дата открытой смены

          CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

            REPLACE Kassa_CFT.DATE_DOC WITH ReportDateSmena  && Дата документа, пишется Дата открытой смены

          CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

            REPLACE Kassa_CFT.DATE_DOC WITH TTOD(f_OpDate_Den)  && Дата документа, пишется Дата открытого операционного дня  TTOD(f_OpDate_Den)

        ENDCASE


      CASE colvo_denCFT > 1 AND DATE() <= TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня.

        DO CASE
          CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

            REPLACE Kassa_CFT.DATE_DOC WITH DATE()  && Дата документа, пишется календарная дата

          CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

            REPLACE Kassa_CFT.DATE_DOC WITH DATE()  && Дата документа, пишется календарная дата

          CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

            REPLACE Kassa_CFT.DATE_DOC WITH DATE()  && Дата документа, пишется календарная дата

        ENDCASE

    ENDCASE

* --------------------------------------------------------- Формирование значения поля "Дата валютирования" ------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE colvo_denCFT = 1 AND DATE() <= TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт = 1 дню. Схема рабочего дня.

        DO CASE
          CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

            REPLACE Kassa_CFT.DATE_PROV WITH CTOD('  /  /    ')  && Дата проводки, пишется Дата открытого операционного дня

          CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

            REPLACE Kassa_CFT.DATE_PROV WITH ReportDateSmena  && Дата документа, пишется Дата открытой смены

          CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

            REPLACE Kassa_CFT.DATE_PROV WITH CTOD('  /  /    ')  && Дата проводки, пишется Дата открытого операционного дня

        ENDCASE


      CASE colvo_denCFT > 1 AND DATE() <= TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня.

        DO CASE
          CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

            REPLACE Kassa_CFT.DATE_PROV WITH CTOD('  /  /    ')  && Дата проводки, пишется Дата открытого операционного дня

          CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

            REPLACE Kassa_CFT.DATE_PROV WITH CTOD('  /  /    ')  && Дата проводки, пишется Дата открытого операционного дня

          CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

            REPLACE Kassa_CFT.DATE_PROV WITH CTOD('  /  /    ')  && Дата проводки, пишется Дата открытого операционного дня

        ENDCASE


    ENDCASE

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    f_Comment = STRTRAN(Plat_PaketKasOrderATM.Comment, CHR(10), SPACE(0))
    f_Comment = STRTRAN(f_Comment, CHR(13), SPACE(0))

    REPLACE Kassa_CFT.NAZN WITH PADR(ALLTRIM(f_Comment), 250)  && Назначение платежа

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    REPLACE Kassa_CFT.PRIORITET WITH 1 && Приоритет, допустимые значения:1,2,3,4,5,6

    REPLACE Kassa_CFT.SUM WITH Plat_PaketKasOrderATM.TotalSumKP  && Сумма дебета

    REPLACE Kassa_CFT.SUM_PO WITH Plat_PaketKasOrderATM.TotalSumKP  && Сумма кредита

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    DO CASE
      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'RUB'

        REPLACE Kassa_CFT.SUM_NT WITH Plat_PaketKasOrderATM.TotalSumKP_RUB  && Сумма национального покрытия

      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'USD'

        REPLACE Kassa_CFT.SUM_NT WITH Plat_PaketKasOrderATM.TotalSumKP_RUB  && Сумма национального покрытия

      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'EUR'

        REPLACE Kassa_CFT.SUM_NT WITH Plat_PaketKasOrderATM.TotalSumKP_RUB  && Сумма национального покрытия

    ENDCASE

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
* -------------------------------------------------------------------------------------------------------------------------------- num_branch == '01'  && Номер филиала г. Москва ------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    DO CASE
      CASE ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) == '02'

        IF DAY(TTOD(OperDateCFT)) = 1  && Дата открытого операционного дня
          REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_PO, 17, 2)) + ', ' + '89' + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
        ELSE
          REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_PO, 17, 2)) + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
        ENDIF


      CASE ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) == '11' AND ALLTRIM(Plat_PaketKasOrderATM.Acc20) == '40821810855990000000'

        IF ALLTRIM(Plat_PaketKasOrderATM.Acc20) == '40821810855990000000'
          REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_PO, 17, 2)) + ', ' + '100' + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
        ELSE
          REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_PO, 17, 2)) + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
        ENDIF


      CASE ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) == '11' AND ALLTRIM(Plat_PaketKasOrderATM.Acc20) == '40821810955990000007'

        IF ALLTRIM(Plat_PaketKasOrderATM.Acc20) == '40821810955990000007'
          REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_PO, 17, 2)) + ', ' + '100' + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
        ELSE
          REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_PO, 17, 2)) + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
        ENDIF


      CASE ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) == '11' AND ALLTRIM(Plat_PaketKasOrderATM.Acc20) == '40821810055990100014'

        IF ALLTRIM(Plat_PaketKasOrderATM.Acc20) == '40821810055990100014'
          REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_NT, 17, 2)) + ', ' + '100' + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
        ELSE
          REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_NT, 17, 2)) + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
        ENDIF


      CASE ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) == '13' AND ALLTRIM(Plat_PaketKasOrderATM.Acc20) == '40821810155990000001'

        IF ALLTRIM(Plat_PaketKasOrderATM.Acc20) == '40821810155990000001'
          REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_PO, 17, 2)) + ', ' + '99' + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
        ELSE
          REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_PO, 17, 2)) + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
        ENDIF


      CASE ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) == '13' AND ALLTRIM(Plat_PaketKasOrderATM.Acc20) == '40821810255990200015'

        IF ALLTRIM(Plat_PaketKasOrderATM.Acc20) == '40821810255990200015'
          REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_NT, 17, 2)) + ', ' + '99' + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
        ELSE
          REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_NT, 17, 2)) + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
        ENDIF


      OTHERWISE


        DO CASE
          CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'RUB'

            REPLACE Kassa_CFT.SUM_KSPL WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.BuhSymbol_eko) + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_PO, 17, 2)) + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>

          CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'USD'

            REPLACE Kassa_CFT.SUM_KSPL WITH PADR('10' + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_PO, 17, 2)) + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>

          CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'EUR'

            REPLACE Kassa_CFT.SUM_KSPL WITH PADR('10' + ', ' + ALLTRIM(STR(Kassa_CFT.SUM_PO, 17, 2)) + ';', 250)  && Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>

        ENDCASE


    ENDCASE

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    REPLACE Kassa_CFT.VID_DOC WITH 'КАСС_ПРИХОД_ОРД' && Код из справочника "Наименования платежных документов" БЕЗН_МЕМО_ОРД - мемориальный ордер; ВНЕБ_ПРИХ_РАСХ - внебалансовый мемориальный ордер

    REPLACE Kassa_CFT.KL_DT_INTERNAL WITH PADR('True', 5)  && Признак - внутренний клиент или внешний, в нашем случае всегда True.
    REPLACE Kassa_CFT.KL_DT_CLIENT WITH PADR('', 25)  && Клиент- владелец

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* ------------------------------------------ Формирование счета кассы для клиента: либо счет операционной кассы, либо счет после операционной кассы -------------------------------------------------------------------------- *

*** Inkahran_Kassa_Acc - счет операционной кассы
*** Inkahran_Kassa_Acc_Pos_RUB - счет после операционной кассы

    DO CASE
      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'RUB'

        DO CASE
          CASE colvo_denCFT = 1 AND DATE() <= TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт = 1 дню. Схема рабочего дня.

            DO CASE
              CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_RUB), 25)  && Счет кассы для клиента


              CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_RUB), 25)  && Счет кассы для клиента


              CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_RUB), 25)  && Счет кассы для клиента

            ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          CASE colvo_denCFT > 1 AND DATE() < TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня.

            REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_RUB), 25)  && Счет кассы для клиента


          CASE colvo_denCFT > 1 AND DATE() = TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня. Утро первого рабочего дня

            DO CASE
              CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_RUB), 25)  && Счет кассы для клиента


              CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_RUB), 25)  && Счет кассы для клиента


              CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_RUB), 25)  && Счет кассы для клиента

            ENDCASE


        ENDCASE

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'USD'

        DO CASE
          CASE colvo_denCFT = 1 AND DATE() <= TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт = 1 дню. Схема рабочего дня.

            DO CASE
              CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_USD), 25)  && Счет кассы для клиента


              CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_USD), 25)  && Счет кассы для клиента


              CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_USD), 25)  && Счет кассы для клиента

            ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          CASE colvo_denCFT > 1 AND DATE() < TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня.

            REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_USD), 25)  && Счет кассы для клиента


          CASE colvo_denCFT > 1 AND DATE() = TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня. Утро первого рабочего дня

            DO CASE
              CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_USD), 25)  && Счет кассы для клиента


              CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_USD), 25)  && Счет кассы для клиента


              CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_USD), 25)  && Счет кассы для клиента

            ENDCASE


        ENDCASE

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'EUR'

        DO CASE
          CASE colvo_denCFT = 1 AND DATE() <= TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт = 1 дню. Схема рабочего дня.

            DO CASE
              CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_EUR), 25)  && Счет кассы для клиента


              CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_EUR), 25)  && Счет кассы для клиента


              CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_EUR), 25)  && Счет кассы для клиента

            ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          CASE colvo_denCFT > 1 AND DATE() < TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня.

            REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_EUR), 25)  && Счет кассы для клиента


          CASE colvo_denCFT > 1 AND DATE() = TTOD(OperDateCFT)  && Разница между датой открытого опердня и датой, в которой он был открыт > 1 дня. Схема выходного дня. Утро первого рабочего дня

            DO CASE
              CASE Plat_PaketKasOrderATM.IdPartKP = 1  && В КП сумка пересчитана в смене УТРО

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_EUR), 25)  && Счет кассы для клиента


              CASE Plat_PaketKasOrderATM.IdPartKP = 2  && В КП сумка пересчитана в смене ДЕНЬ

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_EUR), 25)  && Счет кассы для клиента


              CASE Plat_PaketKasOrderATM.IdPartKP = 3  && В КП сумка пересчитана в смене НОЧЬ

                REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_Pos_EUR), 25)  && Счет кассы для клиента

            ENDCASE


        ENDCASE


    ENDCASE


*      DO CASE
*        CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'RUB'

*          REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_RUB), 25)  && Счет кассы для клиента

*        CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'USD'

*          REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_USD), 25)  && Счет кассы для клиента

*        CASE ALLTRIM(Plat_PaketKasOrderATM.Ccy) == 'EUR'

*          REPLACE Kassa_CFT.KL_DT_ACC WITH PADR(ALLTRIM(Inkahran_Kassa_Acc_EUR), 25)  && Счет кассы для клиента

*      ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    REPLACE Kassa_CFT.KL_KT_INTERNAL WITH PADR('True', 25)  && Признак - внутренний клиент или внешний, в нашем случае всегда True.
    REPLACE Kassa_CFT.KL_KT_CLIENT WITH PADR('', 25)  && Клиент- владелец
    REPLACE Kassa_CFT.KL_KT_ACC WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.Acc20), 25)  && Счет получателя
    REPLACE Kassa_CFT.REQ_MAN_CLIENT WITH PADR('', 25)  && CHAR(25) Внутренний клиент. Заполняется для кассового документа.

*    REPLACE Kassa_CFT.REQ_MAN_NAME WITH PADR(Plat_PaketKasOrderATM.Fio_Kontroler, 160)  && CHAR(160) Наименование  клиента- вносителя/получателя средств. Заполняется для кассового документа.
    REPLACE Kassa_CFT.REQ_MAN_NAME WITH PADR('', 160)  && CHAR(160) Наименование  клиента- вносителя/получателя средств. Заполняется для кассового документа.

    REPLACE Kassa_CFT.REQ_MAN_COUNTRY WITH PADR('', 100)  && CHAR(100) Код  страны из справочника "Общероссийский классификатор стран мира"
    REPLACE Kassa_CFT.REQ_MAN_ADDRESS WITH PADR('', 250)  && CHAR(250) Адрес строкой клиента- вносителя/получателя средств. Заполняется для кассового документа.
    REPLACE Kassa_CFT.REQ_MAN_REZIDENT WITH PADR('True', 5)  && CHAR(5) Свойство резидентности клиента- вносителя/получателя средств из удостоверения. Заполняется для кассового документа, в нашем случае всегда True.
    REPLACE Kassa_CFT.REQ_MAN_DOC_BIRTH_PLACE WITH PADR('', 250)  && CHAR(250) Место рождения клиента- вносителя/получателя средств из удостоверения. Заполняется для кассового документа.

* ------------------------------------------------ Со слов Заировой Е.Т. паспортные данные в приходных кассовых ордерах не требуются в ЦФТ ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*      REPLACE Kassa_CFT.REQ_MAN_DOC_DATE WITH TTOD(Plat_PaketKasOrderATM.Pass_data)  && Дата выдачи документа

*      REPLACE Kassa_CFT.REQ_MAN_DOC_TYPE WITH PADR(IIF(ALLTRIM(Plat_PaketKasOrderATM.Pass_Type) == '21', 'Паспорт РФ', ALLTRIM(Plat_PaketKasOrderATM.Pass_Type)), 50)  && Код удостоверяющего документа

*      REPLACE Kassa_CFT.REQ_MAN_DOC_NUM WITH PADR(RIGHT(ALLTRIM(Plat_PaketKasOrderATM.Pass_numer), 6), 16)  && Серия документа

*      REPLACE Kassa_CFT.REQ_MAN_DOC_SER WITH PADR(LEFT(ALLTRIM(Plat_PaketKasOrderATM.Pass_numer), 4), 16)  && Номер документа

*      REPLACE Kassa_CFT.REQ_MAN_DOC_WHO WITH PADR(ALLTRIM(Plat_PaketKasOrderATM.Pass_mesto), 100)  && CHAR(100) Кем выдан документ

*      REPLACE Kassa_CFT.REQ_MAN_DOC_DEPART_CODE WITH SPACE(20)  && CHAR(20)  Код подразделения, выдавшего удостоверение клиенту- вносителю/получателю средств. Заполняется для кассового документа.

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    REPLACE Kassa_CFT.REQ_MAN_DOC_DATE WITH CTOD('  /  /    ')  && Дата выдачи документа

    REPLACE Kassa_CFT.REQ_MAN_DOC_TYPE WITH PADR('', 50)  && Код удостоверяющего документа

    REPLACE Kassa_CFT.REQ_MAN_DOC_NUM WITH PADR('', 16)  && Серия документа

    REPLACE Kassa_CFT.REQ_MAN_DOC_SER WITH PADR('', 16)  && Номер документа

    REPLACE Kassa_CFT.REQ_MAN_DOC_WHO WITH PADR('', 100)  && CHAR(100) Кем выдан документ

    REPLACE Kassa_CFT.REQ_MAN_DOC_DEPART_CODE WITH SPACE(20)  && CHAR(20)  Код подразделения, выдавшего удостоверение клиенту- вносителю/получателю средств. Заполняется для кассового документа.

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SELECT Plat_PaketKasOrderATM
  ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  SELECT Kassa_CFT
  GO TOP

* BROWSE WINDOW brows

  stroka_titul = ;
    PADR('<ICODE', 25 - 1) + '>' + SPACE(1) + ;  && CHAR(25) Уникальный код документа
  PADR('<FILIAL', 8 - 1) + '>' + SPACE(1) + ;  && CHAR(3) Код филиала из справочника "Филиалы"
  PADR('<DEPART', 20 - 1) + '>' + SPACE(1) + ;  && CHAR(20) Код филиала из справочника "Подразделения"
  PADR('<ACC_DT', 25 - 1) + '>' + SPACE(1) + ;  && CHAR(25) Счет плательщика, счет по дебету
  PADR('<ACC_KT', 25 - 1) + '>' + SPACE(1) + ;  && CHAR(25) Счет получателя, счет по кредиту
  PADR('<DATE_DOC', 10 - 1) + '>' + SPACE(1) + ;  && DATETIME Дата документа
  PADR('<DATE_PROV', 11 - 1) + '>' + SPACE(1) + ;  && DATETIME Дата проводки
  PADR('<NAZN', 250 - 1) + '>' + SPACE(1) + ;  && CHAR(250) Назначение платежа
  PADR('<PRIORITET', 11 - 1) + '>' + SPACE(1) + ;  && INTEGER(4) Приоритет, допустимые значения:1,2,3,4,5,6
  PADR('<SUM', 20 - 1) + '>' + SPACE(1) + ;  && Numeric(17, 2) Сумма дебета
  PADR('<SUM_PO', 20 - 1) + '>' + SPACE(1) + ;  && Numeric(17, 2) Сумма кредита
  PADR('<SUM_NT', 20 - 1) + '>' + SPACE(1) + ;  && Numeric(17, 2) Сумма национального покрытия
  PADR('<SUM_KSPL', 320 - 1) + '>' + SPACE(1) + ;  && CHAR(250) Сумма документа в разбиении по символам кассового плана в специальном формате: <символ кассового плана>, <сумма>, <забалансовый символ касс.плана>
  PADR('<VID_DOC', 16 - 1) + '>' + SPACE(1) + ;  && CHAR(16) Код из справочника "Наименования платежных документов" КАСС_ПРИХОД_ОРД кассовый приходный ордер, БЕЗН_МЕМО_ОРД - мемориальный ордер; ВНЕБ_ПРИХ_РАСХ - внебалансовый мемориальный ордер
  PADR('<KL_DT_INTERNAL', 16 - 1) + '>' + SPACE(1) + ;  && CHAR(5) Признак - внутренний клиент или внешний, в нашем случае всегда True.
  PADR('<KL_DT_CLIENT', 25 - 1) + '>' + SPACE(1) + ;  && CHAR(25) Клиент- владелец
  PADR('<KL_DT_ACC', 25 - 1) + '>' + SPACE(1) + ;  && CHAR(25) Счет кассы для клиента
  PADR('<KL_KT_INTERNAL', 16 - 1) + '>' + SPACE(1) + ;  && CHAR(5)  Признак - внутренний клиент или внешний, в нашем случае всегда True.
  PADR('<KL_KT_CLIENT', 25 - 1) + '>' + SPACE(1) + ;  && CHAR(25) Клиент- владелец
  PADR('<KL_KT_ACC', 25 - 1) + '>' + SPACE(1) + ;  && CHAR(25) Счет получателя
  PADR('<REQ_MAN_CLIENT', 26 - 1) + '>' + SPACE(1) + ;  && CHAR(25),;  && Внутренний клиент. Заполняется для кассового документа.
  PADR('<REQ_MAN_NAME', 160 - 1) + '>' + SPACE(1) + ;  && CHAR(160) Наименование  клиента- вносителя/получателя средств. Заполняется для кассового документа.
  PADR('<REQ_MAN_COUNTRY', 100 - 1) + '>' + SPACE(1) + ;  && CHAR(100) Код  страны из справочника "Общероссийский классификатор стран мира"
  PADR('<REQ_MAN_ADDRESS', 250 - 1) + '>' + SPACE(1) + ;  && CHAR(250) Адрес строкой клиента- вносителя/получателя средств. Заполняется для кассового документа.
  PADR('<REQ_MAN_REZIDENT', 18 - 1) + '>' + SPACE(1) + ;  && CHAR(5) Свойство резидентности клиента- вносителя/получателя средств из удостоверения. Заполняется для кассового документа, в нашем случае всегда True.
  PADR('<REQ_MAN_DOC_BIRTH_PLACE', 250 - 1) + '>' + SPACE(1) + ;  && CHAR(250) Место рождения клиента- вносителя/получателя средств из удостоверения. Заполняется для кассового документа.
  PADR('<REQ_MAN_DOC_DATE', 18 - 1) + '>' + SPACE(1) + ;  && DATETIME Дата выдачи удостоверения клиента- вносителя/получателя средств. Заполняется для кассового документа.
  PADR('<REQ_MAN_DOC_TYPE', 50 - 1) + '>' + SPACE(1) + ;  && CHAR(50) Тип удостоверения (строкой) клиента- вносителя/получателя средств. Заполняется для кассового документа.
  PADR('<REQ_MAN_DOC_SER', 17 - 1) + '>' + SPACE(1) + ;  && CHAR(16) Серия удостоверения (строкой) клиента- вносителя/получателя средств. Заполняется для кассового документа.
  PADR('<REQ_MAN_DOC_NUM', 17 - 1) + '>' + SPACE(1) + ;  && CHAR(16) Номер удостоверения клиента- вносителя/получателя средств. Заполняется для кассового документа.
  PADR('<REQ_MAN_DOC_WHO', 100 - 1) + '>' + SPACE(1) + ;  && CHAR(100) Кем выдано удостоверение клиента- вносителя/получателя средств. Заполняется для кассового документа.
  PADR('<REQ_MAN_DOC_DEPART_CODE', 25 - 1) + '>' + SPACE(1) + 'END_HEAD'  && CHAR(20)) Код подразделения, выдавшего удостоверение клиенту- вносителю/получателю средств. Заполняется для кассового документа.

  =FPUTS(_TEXT, (stroka_titul))

  SCAN

    f_ICODE = Kassa_CFT.ICODE + SPACE(1)
    f_FILIAL = Kassa_CFT.FILIAL + SPACE(5) + SPACE(1)
    f_DEPART = Kassa_CFT.DEPART + SPACE(1)
    f_ACC_DT = Kassa_CFT.ACC_DT + SPACE(1)
    f_ACC_KT = Kassa_CFT.ACC_KT + SPACE(1)
    f_DATE_DOC = PADR(ALLTRIM(DTOC(Kassa_CFT.DATE_DOC)), 10) + SPACE(1)
    f_DATE_PROV = PADR(DTOC(Kassa_CFT.DATE_PROV), 11) + SPACE(1)
    f_NAZN = Kassa_CFT.NAZN + SPACE(1)
    f_PRIORITET = ALLTRIM(STR(Kassa_CFT.PRIORITET)) + SPACE(10) + SPACE(1)

    f_SUM = PADR(ALLTRIM(STR(Kassa_CFT.SUM, 17, 2)), 20) + SPACE(1)
    f_SUM_PO = PADR(ALLTRIM(STR(Kassa_CFT.SUM_PO, 17, 2)), 20) + SPACE(1)
    f_SUM_NT = PADR(ALLTRIM(STR(Kassa_CFT.SUM_NT, 17, 2)), 20) + SPACE(1)
    f_SUM_KSPL = PADR(ALLTRIM(Kassa_CFT.SUM_KSPL), 320) + SPACE(1)
    f_VID_DOC = Kassa_CFT.VID_DOC + SPACE(1)

    f_KL_DT_INTERNAL = Kassa_CFT.KL_DT_INTERNAL + SPACE(11) + SPACE(1)
    f_KL_DT_CLIENT = Kassa_CFT.KL_DT_CLIENT + SPACE(1)
    f_KL_DT_ACC = Kassa_CFT.KL_DT_ACC + SPACE(1)
    f_KL_KT_INTERNAL = Kassa_CFT.KL_KT_INTERNAL + SPACE(11) + SPACE(1)
    f_KL_KT_CLIENT = Kassa_CFT.KL_KT_CLIENT + SPACE(1)
    f_KL_KT_ACC = Kassa_CFT.KL_KT_ACC + SPACE(1)

    f_REQ_MAN_CLIENT = Kassa_CFT.REQ_MAN_CLIENT + SPACE(1)

    f_REQ_MAN_NAME = SPACE(1) + Kassa_CFT.REQ_MAN_NAME

    f_REQ_MAN_COUNTRY = Kassa_CFT.REQ_MAN_COUNTRY + SPACE(1)
    f_REQ_MAN_ADDRESS = Kassa_CFT.REQ_MAN_ADDRESS + SPACE(1) + SPACE(1)
    f_REQ_MAN_REZIDENT = Kassa_CFT.REQ_MAN_REZIDENT + SPACE(12) + SPACE(1)
    f_REQ_MAN_DOC_BIRTH_PLACE = PADR(ALLTRIM(Kassa_CFT.REQ_MAN_DOC_BIRTH_PLACE), 250) + SPACE(2)
    f_REQ_MAN_DOC_DATE = PADR(DTOC(Kassa_CFT.REQ_MAN_DOC_DATE), 18) + SPACE(1)
    f_REQ_MAN_DOC_TYPE = Kassa_CFT.REQ_MAN_DOC_TYPE + SPACE(1)
    f_REQ_MAN_DOC_SER = PADR(ALLTRIM(Kassa_CFT.REQ_MAN_DOC_SER), 17) + SPACE(1)
    f_REQ_MAN_DOC_NUM = PADR(ALLTRIM(Kassa_CFT.REQ_MAN_DOC_NUM), 17) + SPACE(1)
    f_REQ_MAN_DOC_WHO = Kassa_CFT.REQ_MAN_DOC_WHO + SPACE(1)
    f_REQ_MAN_DOC_DEPART_CODE = PADR(ALLTRIM(Kassa_CFT.REQ_MAN_DOC_DEPART_CODE), 25) + SPACE(1)

    stroka_detal_record = ;
      f_ICODE + ;
      f_FILIAL + ;
      f_DEPART + ;
      f_ACC_DT + ;
      f_ACC_KT + ;
      f_DATE_DOC + ;
      f_DATE_PROV + ;
      f_NAZN + ;
      f_PRIORITET + ;
      f_SUM + ;
      f_SUM_PO + ;
      f_SUM_NT + ;
      f_SUM_KSPL + ;
      f_VID_DOC + ;
      f_KL_DT_INTERNAL + ;
      f_KL_DT_CLIENT + ;
      f_KL_DT_ACC + ;
      f_KL_KT_INTERNAL + ;
      f_KL_KT_CLIENT + ;
      f_KL_KT_ACC + ;
      f_REQ_MAN_CLIENT + ;
      f_REQ_MAN_NAME + ;
      f_REQ_MAN_COUNTRY + ;
      f_REQ_MAN_ADDRESS + ;
      f_REQ_MAN_REZIDENT + ;
      f_REQ_MAN_DOC_BIRTH_PLACE + ;
      f_REQ_MAN_DOC_DATE + ;
      f_REQ_MAN_DOC_TYPE + ;
      f_REQ_MAN_DOC_SER + ;
      f_REQ_MAN_DOC_NUM + ;
      f_REQ_MAN_DOC_WHO + ;
      f_REQ_MAN_DOC_DEPART_CODE

    =FPUTS(_TEXT, (stroka_detal_record))

  ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  =FCLOSE(_TEXT)

ENDIF

SET TEXTMERGE OFF

SET DEFAULT TO (old_path)

@ WROWS()/3-0.2,3 SAY PADC('Файл для ЦФТ - ' + ALLTRIM(exp_file_cft) + ' успешно сформирован ... ',WCOLS())
=INKEY(1)
DEACTIVATE WINDOW poisk

SET MARK TO '.' && тип разделителя в формате даты

RETURN


***********************************************************************************************************************************************************************************************************************


