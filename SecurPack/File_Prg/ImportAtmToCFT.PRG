**************************************************
*** Создание пакета платежей из сведенных сумок ***
**************************************************

PROCEDURE start
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

f_IdDepartmentPRK = GetGlob('DepartPRK')  && Номер подразделения ПРК по справочнику Department

DIMENSION ar[1,2]

* f_OperDate_OpDate - дата открытого операционного дня, из таблицы OperDate
* f_IdOperDate - Id открытого операционного дня, из таблицы Shift
* f_IdShift_2 - Id открытой смены Демартамент = 2, это КП (касса пересчета), из таблицы Shift

f_CurDate = new_data_odb  && Дата открытой смены в КП

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF f_IdUser_BossKassir = 0

  DO CASE
    CASE num_branch == '01'  && Номер филиала г. Москва

      f_IdUser_BossKassir = 298  && Данные по сотруднику УКО, от имени которого будут уходить платежи в ЦФТ
      f_Fio_BossKassir = 'Кипченко Наталия Викторовна'  && Сотрудник УКО, от имени которого будут уходить платежи в ЦФТ

*        f_IdUser_BossKassir = 2018  && Данные по сотруднику УКО, от имени которого будут уходить платежи в ЦФТ
*        f_Fio_BossKassir = 'Румянцева Елена Петровна'  && Сотрудник УКО, от имени которого будут уходить платежи в ЦФТ

*        f_IdUser_BossKassir = 351  && Данные по сотруднику УКО, от имени которого будут уходить платежи в ЦФТ
*        f_Fio_BossKassir = 'Шелестенко Елена Александровна'  && Сотрудник УКО, от имени которого будут уходить платежи в ЦФТ


    CASE num_branch == '06'  && Номер филиала г. Питер

      f_IdUser_BossKassir = 15  && Данные по сотруднику УКО, от имени которого будут уходить платежи в ЦФТ
      f_Fio_BossKassir = 'Гацук Елена Васильевна'  && Сотрудник УКО, от имени которого будут уходить платежи в ЦФТ

*        f_IdUser_BossKassir = 53  && Данные по сотруднику УКО, от имени которого будут уходить платежи в ЦФТ
*        f_Fio_BossKassir = 'Пермякова Татьяна Ивановна'  && Сотрудник УКО, от имени которого будут уходить платежи в ЦФТ

*        f_IdUser_BossKassir = 58  && Данные по сотруднику УКО, от имени которого будут уходить платежи в ЦФТ
*        f_Fio_BossKassir = 'Радионова Марина Леонидовна'  && Сотрудник УКО, от имени которого будут уходить платежи в ЦФТ

  ENDCASE

ENDIF

* WAIT WINDOW 'f_IdUser_BossKassir = ' + ALLTRIM(STR(f_IdUser_BossKassir)) TIMEOUT 3

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* ----------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ----------------------------------------------------- *

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT Id, IdPost, IdDepartment, IdUser, Fio, TabNumber, IsActive,
  Pass_Type, Pass_Seria, Pass_Numer, Pass_Data, Pass_Mesto
  FROM CashKontroller (READPAST)
  WHERE IdUser = ?f_IdUser_BossKassir
  ORDER BY IdUser

ENDTEXT

IF RunSQL_Load('CashKontroller', @ar) <> 1
  MESSAGEBOX('Ошибка выборки данных из таблицы FROM CashKontroller', 16, 'Обратитесь к программистам по тел: 243, 303, 170, 256')
  f_IdUser_BossKassir = 2018
ENDIF

* BROWSE WINDOW brows TITLE 'Данные по сотруднику УКО, от имени которого будут уходить платежи в ЦФТ' + '   f_IdUser_BossKassir = ' + ALLTRIM(STR(f_IdUser_BossKassir))

IF ar[1,2] = 0

  IF USED('CashKontroller') = .T.
    USE IN CashKontroller
  ENDIF

* MESSAGEBOX('Данных по сотруднику УКО, открывшему смену КП' + CHR(13) + 'за дату - ' + DTOC(new_data_odb) + ' и номером смены ' + ALLTRIM(STR(f_IdShift_2)) + CHR(13) + 'НЕ ОБНАРУЖЕНО', 16, 'Обратитесь к начальнику смены Кассы Пересчета')

  f_IdUser_BossKassir = 2018  && f_IdUser_BossKassir = 888  Не существующий сотрудник, номер для тестирования

  TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

    SELECT IdUser, Fio, TabNumber, Pass_Type, Pass_Seria, Pass_Numer, Pass_Data, Pass_Mesto
    FROM CashKontroller (READPAST)
    WHERE IdUser = ?f_IdUser_BossKassir
    ORDER BY IdUser

  ENDTEXT

  IF RunSQL_Load('CashKontroller', @ar) <> 1
    MESSAGEBOX('Ошибка выборки данных из таблицы FROM CashKontroller', 16, 'Обратитесь к программистам по тел: 243, 303, 170, 256')
    f_IdUser_BossKassir = 351
  ENDIF

* BROWSE WINDOW brows TITLE 'Данные универсального сотрудника УКО, от имени которого будут уходить платежи в ЦФТ' + '   f_IdUser_BossKassir = ' + ALLTRIM(STR(f_IdUser_BossKassir))

  IF ar[1,2] = 0

    IF USED('CashKontroller') = .T.
      USE IN CashKontroller
    ENDIF

    f_IdUser_BossKassir = 351  && f_IdUser_BossKassir = 888  Не существующий сотрудник, номер для тестирования

    TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

    SELECT IdUser, Fio, TabNumber, Pass_Type, Pass_Seria, Pass_Numer, Pass_Data, Pass_Mesto
    FROM CashKontroller (READPAST)
    WHERE IdUser = ?f_IdUser_BossKassir
    ORDER BY IdUser

    ENDTEXT

    IF RunSQL_Load('CashKontroller', @ar) <> 1
      MESSAGEBOX('Ошибка выборки данных из таблицы FROM CashKontroller', 16, 'Обратитесь к программистам по тел: 243, 303, 170, 256')
      RETURN
    ENDIF

* BROWSE WINDOW brows TITLE 'Данные универсального сотрудника УКО, от имени которого будут уходить платежи в ЦФТ' + '   f_IdUser_BossKassir = ' + ALLTRIM(STR(f_IdUser_BossKassir))

    IF ar[1,2] = 0

      MESSAGEBOX('Данных по сотруднику УКО IdUser = ' + ALLTRIM(STR(f_IdUser_BossKassir)) + ', имеющем право' + CHR(13) + 'отправлять приходные кассовые платежи в ЦФТ' + CHR(13) + 'НЕ ОБНАРУЖЕНО', 16, 'Обратитесь к начальнику смены Кассы Пересчета')
      RETURN

    ENDIF

  ENDIF

ENDIF

IF USED('CashKontroller') = .T.

  SELECT CashKontroller
  GO TOP
  f_Fio_BossKassir = ALLTRIM(CashKontroller.Fio)  && Сотрудник УКО, от имени которого будут уходить платежи в ЦФТ

ELSE
  f_Fio_BossKassir = 'Кипченко Наталия Викторовна'  && Сотрудник УКО, от имени которого будут уходить платежи в ЦФТ
ENDIF

* WAIT WINDOW 'Данные сотруднику УКО, от имени, которого будут уходить платежи в ЦФТ:' + CHR(13) + 'ФИО - ' + ALLTRIM(CashKontroller.Fio) TIMEOUT 5

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

put_select = 1

OperDate_PRK = DATE() - 1
IdOperDate_PRK = 0
IdShift_PRK = ''

f_FormName = 'Информация'

dtB = new_data_odb  && Дата открытой смены в КП
f_CurDate = new_data_odb  && Дата открытой смены в КП

* WAIT WINDOW 'dtB = ' + DTOC(dtB) TIMEOUT 3
* WAIT WINDOW '1 - f_CurDate = ' + DTOC(f_CurDate) TIMEOUT 3

* ------------------------------------ *
DO FORM GetDate1
* ------------------------------------ *

IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

  f_CurDate = dtb

*   WAIT WINDOW '2 - f_CurDate = ' + DTOC(f_CurDate) TIMEOUT 3

* ---------------------------------------- Находим смену ПРК за дату f_CurDate ------------------------------------- *

  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

    SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate, Od.OpenDate
    FROM Shift AS Sh (READPAST)
    INNER JOIN OperDate AS Od ON Od.Id = Sh.IdOperDate
    WHERE Sh.ReportDate = ?f_CurDate AND Sh.IdDepartment = ?f_IdDepartmentPRK
    ORDER BY Sh.Id DESC

  ENDTEXT

  IF RunSQL_Load('Shift_PRK', @ar) <> 1
    MESSAGEBOX('Ошибка работы со справочником Смен ... ', 16, 'Смена не найдена.')
    RETURN -1
  ENDIF

  IF USED('Shift_PRK') = .T. AND RECCOUNT('Shift_PRK') > 0

    SELECT Shift_PRK
    GO TOP

* BROWSE WINDOW brows

    OperDatePRK = TTOD(Shift_PRK.OpDate)  && Дата открытого опердня
    IdOperDatePRK = Shift_PRK.IdOperDate  && Id открытого опердня
    IdShiftPRK_CFT = Shift_PRK.Id
    ReportDateSmenaPRK_CFT = Shift_PRK.ReportDate && Дата открытой суточной смены ПРК для двух подсмен.

    f_OperDate_OpDatePRK = OperDatePRK  && дата открытого операционного дня, из таблицы OperDate
    f_IdOperDatePRK = IdOperDatePRK  && Id открытого операционного дня, из таблицы Shift

  ELSE
    =err('Внимание! За выбранную дату доступных смен ПРК в справочнике не обнаружено ... ')
    RETURN
  ENDIF

* ---------------------------------------- Находим смену КП за дату f_CurDate ------------------------------------- *

  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

    SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate, Od.OpenDate, Od.OpenDate AS OpenDateDen
    FROM Shift AS Sh (READPAST)
    INNER JOIN OperDate AS Od ON Od.Id = Sh.IdOperDate
    WHERE Sh.ReportDate = ?f_CurDate AND Sh.IdDepartment = ?f_IdDepartmentKP
    ORDER BY Sh.Id DESC

  ENDTEXT

  IF RunSQL_Load('Shift_KP', @ar) <> 1
    MESSAGEBOX('Ошибка работы со справочником Смен ... ', 16, 'Смена не найдена.')
    RETURN -1
  ENDIF

  IF USED('Shift_KP') = .T. AND RECCOUNT('Shift_KP') > 0

    SELECT Shift_KP
    GO TOP

* BROWSE WINDOW brows

    OperDateKP_CFT = TTOD(Shift_KP.OpDate)  && Дата открытого опердня
    IdOperDateKP_CFT = Shift_KP.IdOperDate  && Id открытого опердня
    IdShiftKP_CFT = Shift_KP.Id
    ReportDateSmenaKP_CFT = Shift_KP.ReportDate && Дата открытой суточной смены ПРК для двух подсмен.

    f_OperDate_OpDateKP_CFT = OperDateKP_CFT  && дата открытого операционного дня, из таблицы OperDate
    f_IdOperDateKP_CFT = IdOperDateKP_CFT  && Id открытого операционного дня, из таблицы Shift

  ELSE
    =err('Внимание! За выбранную дату доступных смен КП в справочнике не обнаружено ... ')
    RETURN
  ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk

  @ WROWS()/3, 3 SAY PADC('Уважаемый Пользователь! Начат импорт пересчетных АТМ за дату - ' + DTOC(f_CurDate), WCOLS())

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* IdBudgetHistory
* IdBudgetKPAtm

* WHERE atm.IdShift IN (<<IdShift_PRK>>)

*    TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

*      SELECT atm.Id, atm.IsDriver, atm.IdShift, atm.FilialCode, atm.BudgetNum, atm.TotalSum_RUB_Out AS TotalSumKP, 'RUB' AS Ccy, 1.0000 AS Kurs, atm.CurDate, atm.DateInkas,
*      cln.TypeAcc, cln.ClientName, ISNULL(cln.IncomeAccount, ''+SPACE(30)+'') AS IncomeAccount, bnk.BankName, cln.Account AS ClAcc,
*      ?f_OperDate_OpDatePRK AS OpDate, ?f_IdOperDatePRK AS IdOperDate, fil.Id AS IdFilial, cln.INN, cln.Id AS Cl_Id,
*      ISNULL(IIF(grp_atm.TypeAcc = '10', grp_atm.Account30109, IIF(grp_atm.TypeAcc = '11', grp_atm.Account30222, '')), ''+SPACE(20)+'') AS Account,
*      ISNULL(bnk.AccTranzit, ''+SPACE(20)+'') AS AccTranzit,
*      ISNULL(bnk.AccLoro, ''+SPACE(20)+'') AS AccLoro,
*      cln.ClientFullName, cln.BuhSymbol AS BuhSymbol_eko, fil.FilialName, cln.FactAddress,
*      knt.Fio, knt.IdPost, knt.IdDepartment,
*      LTRIM(RTRIM(atm.Comment)) AS Comment,
*      LTRIM(RTRIM(knt.Pass_Type)) AS Pass_Type, LTRIM(RTRIM(knt.Pass_Seria)) + LTRIM(RTRIM(knt.Pass_Numer)) AS Pass_Numer, knt.Pass_Data, knt.Pass_Mesto, atm.IdBudgetHistory
*      FROM Tranzit_ATM_PER AS atm (READPAST)
*      INNER JOIN Filial AS fil (READPAST) ON fil.FilialCode = atm.FilialCode
*      INNER JOIN Client AS cln (READPAST) ON cln.Id = fil.IdClient
*      INNER JOIN ClientGroup AS clg (READPAST) ON cln.IdClientGroup = clg.Id
*      INNER JOIN ClientGroupAccounts AS grp_atm (READPAST) ON grp_atm.IdClientGroup = clg.Id AND grp_atm.IdTypeValue = 1
*      INNER JOIN Bank AS bnk (READPAST) ON clg.IdBank = bnk.Id
*      INNER JOIN CashKontroller AS knt (READPAST) ON knt.IdUser = ?f_IdUser_BossKassir
*      INNER JOIN Shift AS shf (READPAST) ON shf.Id = atm.IdShift
*      WHERE shf.ReportDate = ?ReportDateSmenaPRK_CFT AND atm.Id  NOT IN (SELECT Id FROM PaketKasOrderATM (READPAST) WHERE Id = atm.Id AND IsDriver = atm.IsDriver AND Ccy = 'RUB') AND atm.TotalSum_RUB_Out > 0.00 trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
*      UNION ALL
*      SELECT atm.Id, atm.IsDriver, atm.IdShift, atm.FilialCode, atm.BudgetNum, atm.TotalSum_USD_Out AS TotalSumKP, 'USD' AS Ccy, atm.Kurs_USD AS Kurs, atm.CurDate, atm.DateInkas,
*      cln.TypeAcc, cln.ClientName, ISNULL(cln.IncomeAccount, ''+SPACE(30)+'') AS IncomeAccount, bnk.BankName, cln.Account AS ClAcc,
*      ?f_OperDate_OpDatePRK AS OpDate, ?f_IdOperDatePRK AS IdOperDate, fil.Id AS IdFilial, cln.INN, cln.Id AS Cl_Id,
*      ISNULL(IIF(grp_atm.TypeAcc = '10', grp_atm.Account30109, IIF(grp_atm.TypeAcc = '11', grp_atm.Account30222, '')), ''+SPACE(20)+'') AS Account,
*      ISNULL(bnk.AccTranzit, ''+SPACE(20)+'') AS AccTranzit,
*      ISNULL(bnk.AccLoro, ''+SPACE(20)+'') AS AccLoro,
*      cln.ClientFullName, cln.BuhSymbol AS BuhSymbol_eko, fil.FilialName, cln.FactAddress,
*      knt.Fio, knt.IdPost, knt.IdDepartment,
*      LTRIM(RTRIM(atm.Comment)) AS Comment,
*      LTRIM(RTRIM(knt.Pass_Type)) AS Pass_Type, LTRIM(RTRIM(knt.Pass_Seria)) + LTRIM(RTRIM(knt.Pass_Numer)) AS Pass_Numer, knt.Pass_Data, knt.Pass_Mesto, atm.IdBudgetHistory
*      FROM Tranzit_ATM_PER AS atm (READPAST)
*      INNER JOIN Filial AS fil (READPAST) ON fil.FilialCode = atm.FilialCode
*      INNER JOIN Client AS cln (READPAST) ON cln.Id = fil.IdClient
*      INNER JOIN ClientGroup AS clg (READPAST) ON cln.IdClientGroup = clg.Id
*      INNER JOIN ClientGroupAccounts AS grp_atm (READPAST) ON grp_atm.IdClientGroup = clg.Id AND grp_atm.IdTypeValue = 2
*      INNER JOIN Bank AS bnk (READPAST) ON clg.IdBank = bnk.Id
*      INNER JOIN CashKontroller AS knt (READPAST) ON knt.IdUser = ?f_IdUser_BossKassir
*      INNER JOIN Shift AS shf (READPAST) ON shf.Id = atm.IdShift
*      WHERE shf.ReportDate = ?ReportDateSmenaPRK_CFT AND atm.Id NOT IN (SELECT Id FROM PaketKasOrderATM (READPAST) WHERE Id = atm.Id AND IsDriver = atm.IsDriver AND Ccy = 'USD') AND atm.TotalSum_USD_Out > 0.00 trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
*      UNION ALL
*      SELECT atm.Id, atm.IsDriver, atm.IdShift, atm.FilialCode, atm.BudgetNum, atm.TotalSum_EUR_Out AS TotalSumKP, 'EUR' AS Ccy, atm.Kurs_EUR AS Kurs, atm.CurDate, atm.DateInkas,
*      cln.TypeAcc, cln.ClientName, ISNULL(cln.IncomeAccount, ''+SPACE(30)+'') AS IncomeAccount, bnk.BankName, cln.Account AS ClAcc,
*      ?f_OperDate_OpDatePRK AS OpDate, ?f_IdOperDatePRK AS IdOperDate, fil.Id AS IdFilial, cln.INN, cln.Id AS Cl_Id,
*      ISNULL(IIF(grp_atm.TypeAcc = '10', grp_atm.Account30109, IIF(grp_atm.TypeAcc = '11', grp_atm.Account30222, '')), ''+SPACE(20)+'') AS Account,
*      ISNULL(bnk.AccTranzit, ''+SPACE(20)+'') AS AccTranzit,
*      ISNULL(bnk.AccLoro, ''+SPACE(20)+'') AS AccLoro,
*      cln.ClientFullName, cln.BuhSymbol AS BuhSymbol_eko, fil.FilialName, cln.FactAddress,
*      knt.Fio, knt.IdPost, knt.IdDepartment,
*      LTRIM(RTRIM(atm.Comment)) AS Comment,
*      LTRIM(RTRIM(knt.Pass_Type)) AS Pass_Type, LTRIM(RTRIM(knt.Pass_Seria)) + LTRIM(RTRIM(knt.Pass_Numer)) AS Pass_Numer, knt.Pass_Data, knt.Pass_Mesto, atm.IdBudgetHistory
*      FROM Tranzit_ATM_PER AS atm (READPAST)
*      INNER JOIN Filial AS fil (READPAST) ON fil.FilialCode = atm.FilialCode
*      INNER JOIN Client AS cln (READPAST) ON cln.Id = fil.IdClient
*      INNER JOIN ClientGroup AS clg (READPAST) ON cln.IdClientGroup = clg.Id
*      INNER JOIN ClientGroupAccounts AS grp_atm (READPAST) ON grp_atm.IdClientGroup = clg.Id AND grp_atm.IdTypeValue = 3
*      INNER JOIN Bank AS bnk (READPAST) ON clg.IdBank = bnk.Id
*      INNER JOIN CashKontroller AS knt (READPAST) ON knt.IdUser = ?f_IdUser_BossKassir
*      INNER JOIN Shift AS shf (READPAST) ON shf.Id = atm.IdShift
*      WHERE shf.ReportDate = ?ReportDateSmenaPRK_CFT AND atm.Id NOT IN (SELECT Id FROM PaketKasOrderATM (READPAST) WHERE Id = atm.Id AND IsDriver = atm.IsDriver AND Ccy = 'EUR') AND atm.TotalSum_EUR_Out > 0.00 trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
*      ORDER BY atm.IdShift, atm.FilialCode

*    ENDTEXT

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*    TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
*     EXECUTE Select_Tranzit_ATM_PER_ReportDate ?f_OperDate_OpDatePRK, ?f_IdOperDatePRK, ?f_IdUser_BossKassir, ?ReportDateSmenaPRK_CFT
*    ENDTEXT

  TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
    EXECUTE Select_Tranzit_ATM_PER_ReportDate_IdBudget ?f_OperDate_OpDatePRK, ?f_IdOperDatePRK, ?f_IdUser_BossKassir, ?ReportDateSmenaPRK_CFT
  ENDTEXT

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF RunSQL_Load('Tranzit_ATM_PER', @ar) <> 1
    DEACTIVATE WINDOW poisk
    MESSAGEBOX('Ошибка выборки данных из таблицы FROM Tranzit_ATM_PER', 16, 'Обратитесь к программистам по тел: 243, 303, 170, 256')
    DO exit_sql_server IN RunSQL_Load
    RETURN
  ENDIF

  IF ar[1,2] = 0
    IF USED('Tranzit_ATM_PER') = .T.
      USE IN Tranzit_ATM_PER
    ENDIF
    DEACTIVATE WINDOW poisk
    MESSAGEBOX('Приходных кассовых ордеров за дату ' + DTOC(dtb) + CHR(13) + 'созданных и ещё не зачисленных в ЦФТ - НЕ ОБНАРУЖЕНО', 16, 'Обратитесь к начальнику смены КП')
    DO exit_sql_server IN RunSQL_Load
    RETURN
  ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF USED('Tranzit_ATM_PER') = .T.  && В результате выборки таблица с платежами сформирована

    colvo_zap = RECCOUNT()

    IF colvo_zap <> 0  && Количество записей в ней не равно нулю, можно работать дальше

      SELECT Tranzit_ATM_PER
      GO TOP

* BROWSE WINDOW brows

* --------------------------------------------- *
      DO start IN ImportAtmKassa
* --------------------------------------------- *

      tim2 = SECONDS()

      result_time = ROUND(tim2 - tim1,  2)

      @ WROWS()/3, 3 SAY PADC('Импорт данных завершен успешо.' + '  (время = ' + ALLTRIM(TRANSFORM(result_time,  '999.99')) + ' сек.)', WCOLS())
      = INKEY(2)
      DEACTIVATE WINDOW poisk

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      IF USED('Tranzit_ATM_PER') = .T.
        SELECT Tranzit_ATM_PER
        USE
      ENDIF

    ELSE
      =soob('Внимание! В сформированной таблице из базы SQL Server записей не обнаружено ... ')
    ENDIF

  ELSE
    =soob('Внимание! Выборка данных из базы SQL Server закончилась аварийно ... ')
  ENDIF

ENDIF

* ----------------------------------------------------- *
DO exit_sql_server IN RunSQL_Load
* ----------------------------------------------------- *

SELECT(sel)
RETURN


**************************************************************************************************************************************************


