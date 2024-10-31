*****************************************************
*** Формирование ведомости по приему Сумок в ПРК ***
*****************************************************

************************************************************************ PROCEDURE start *************************************************************

PROCEDURE start

IF USED('SecurPack') = .F.
  MESSAGEBOX('Рабочая таблица - SecurPack - НЕ ОТКРЫТА!!!' + CHR(13) + 'Нет готовых сумок для передачи в КП', 16, 'Прием ведомости')
  RETURN -1
ENDIF

SELECT SecurPack
* BROWSE WINDOW brows

DIMENSION ar[1,2]

PRIVATE msg

msg = ''

sel_Sost = '1'

f_CurDate = f_ShiftDate
f_EmptyQty = 0    &&Кол-во порожниж сумок
f_EmptyNumbers = ''  &&Номера порожниж сумок

* --------------------------------------------------------------------------------------------------------------------------------------  *
f_IdShiftPrn = f_IdShift  && Эти переменные нужны для выборки порожних явочных
* f_IdRoute =   && Эти переменные нужны для выборки порожних явочных
* f_RecIdRoute =   && Эти переменные нужны для выборки порожних явочных
* --------------------------------------------------------------------------------------------------------------------------------------  *

IF USED('SecurPack') = .F.
  MESSAGEBOX('Рабочая таблица - SecurPack - НЕ ОТКРЫТА!!!' + CHR(13) + 'Нет готовых сумок для передачи в КП', 16, 'Прием ведомости')
  RETURN -1
ENDIF

SELECT * ;
  FROM SecurPack ;
  WHERE ALLTRIM(Sost) = ALLTRIM(sel_Sost) ;
  INTO CURSOR SecurPack_1 ;
  ORDER BY Flg2PIN ASC

IF USED('SecurPack_1') = .F.
  MESSAGEBOX('Рабочая таблица - SecurPack_1 - НЕ ОТКРЫТА!!!' + CHR(13) + 'Нет готовых сумок для передачи в КП', 16, 'Прием ведомости')
  RETURN -1
ENDIF

SELECT SecurPack_1
GO TOP

* BROWSE WINDOW brows

colvo_zap = _TALLY

IF colvo_zap = 0
  MESSAGEBOX('Нет готовых сумок для передачи в КП', 16, 'Прием ведомости')
  RETURN -1
ENDIF

IF USED('SecurPack') = .F.
  MESSAGEBOX('Рабочая таблица - SecurPack - НЕ ОТКРЫТА!!!' + CHR(13) + 'Нет готовых сумок для передачи в КП', 16, 'Прием ведомости')
  RETURN -1
ENDIF

SELECT SecurPack
* BROWSE WINDOW brows

* WAIT WINDOW 'Запуск процедуры - VedomPRK Процедура записи данных в таблицы' TIMEOUT 2

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

ret = VedomPRK()  && Процедура записи данных в таблицы: RouteHistory - приёмы партий сумок от инкассаторов, BudgetHistoryPRK - записи о принятых и переданных сумках

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW 'Выход из процедуры - VedomPRK Процедура записи данных в таблицы' + '   Ret = ' + ALLTRIM(STR(Ret)) TIMEOUT 5

IF USED('SecurPack_1') = .T.
  SELECT SecurPack_1
  USE
ENDIF

IF USED('SecurPack') = .F.
  MESSAGEBOX('Рабочая таблица - SecurPack - НЕ ОТКРЫТА!!!' + CHR(13) + 'Нет готовых сумок для передачи в КП', 16, 'Прием ведомости')
  RETURN -1
ENDIF

* ---------------------------------------------------------  *
DO Empty_Bauls_11 IN TakeVedomPrn
* ---------------------------------------------------------  *

SELECT SecurPack

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF ret = 1

  MESSAGEBOX('Ведомость успешно принята в ПРК', 48, 'Передача Ведомости в ПРК')
  RELEASE WINDOWS

ELSE

  IF ret = -1
&& Скорее всего будет повтор номера сумки. Но, как минимум, в части ошибочных ситуаций в msg записывается причина ошибки. Если такое произошло - выводим именно его.
    MESSAGEBOX(IIF(!EMPTY(msg),msg,'При выполнении приема ведомости - произошла ОШИБКА'), 16, 'Операция приема ведомости в ПРК Прекращена')
  ENDIF

  IF ret = -2
    MESSAGEBOX('Инкасатор с табельным номером ' + ALLTRIM(f_TabNumber) + ' в базе не найден', 16, 'Операция приема ведомости в ПРК Прекращена')
  ENDIF

  IF ret = -3
    MESSAGEBOX('Данная сумка уже обработана другим оператором', 16, 'Операция создания ведомости по секюрПакам Прекращена')
  ENDIF

ENDIF

* -------------------------------------------------------------------------------------------------------------------------------- *

status_pak = '2'

* f_IdRouteHistory - номер только что вставленной записи в таблицу RouteHistory

* WAIT WINDOW 'f_RecIdRoute = ' + ALLTRIM(STR(f_RecIdRoute)) TIMEOUT 3

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* f_IdPost = 1
* f_IdUser = 1198

*  DO CASE
*    CASE f_IdPost = 1  && Работают пользователи по должности Кассир ПРК

*      TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*        SELECT LEFT(CONVERT(CHAR(8), sp.CurDate, 108), 2) AS CurTime, sp.Id, sp.IdShift, sp.FilialCode, sp.BudgetNum, sp.IdRoute, sp.RecIdRoute, sp.IdUser, sp.TabNumber, sp.IdTypeValue, sp.TypeValueKind,
*        sp.TotalSum, sp.Pr_60322, sp.TotalSum_60322, sp.DateInkas, sp.DtIncome, sp.CurDate, us.UserName, bgh.IdRouteHistory, bgh.IdRouteHistoryMemo, rh.IdCashMessenger, cm.FullName, cm.Fio, sp.Sost, sp.IsTransit, sp.FileName, cfb.FilialName, 0 AS numDef,
*        sp.Flg2PIN, sp.WorkType, sp.SuperBudget, sp.SuperQty, sp.IdJk, cfb.TypeACC, cfb.Account, cfb.Account60322, cfb.UIFactAddress,
*        IIF(sp.IdKeyATM > 0, (SELECT DISTINCT sub.IdKeyATMBank
*        FROM dbo.KeyATMDocSub AS sub
*        INNER JOIN dbo.KeyATM AS ket ON ket.Id = sub.IdKeyATM
*        INNER JOIN dbo.KeyATMBank AS bnk ON bnk.Id = sub.IdKeyATMBank
*        WHERE ket.Id = sp.IdKeyATM), 0) AS IdBank,
*        IIF(sp.IdKeyATM > 0, (SELECT DISTINCT bnk.AtmBankName
*        FROM dbo.KeyATMDocSub AS sub
*        INNER JOIN dbo.KeyATM AS ket ON ket.Id = sub.IdKeyATM
*        INNER JOIN dbo.KeyATMBank AS bnk ON bnk.Id = sub.IdKeyATMBank
*        WHERE ket.Id = sp.IdKeyATM), SPACE(50)) AS BankName,
*        ISNULL((SELECT 1
*        FROM SecurPackDefect AS spd (READPAST)
*        INNER JOIN BudgetDefectNew AS bd (READPAST) ON bd.DefectCode = spd.IdBudgetDefect
*        WHERE spd.IdSecurPack = sp.Id AND bd.IsOpen = 1), 0) AS IsOpen,
*        bud.DefectCode, bud.DefectName
*        FROM SecurPack AS sp (READPAST)
*        INNER JOIN BudgetHistoryPRK AS bgh (READPAST) ON bgh.IdSecurPack = sp.Id
*        INNER JOIN RouteHistory AS rh (READPAST) ON rh.Id = bgh.IdRouteHistory
*        INNER JOIN [User] AS us (READPAST) ON us.Id = rh.IdUser
*        INNER JOIN CashMessenger AS cm (READPAST) ON cm.Id = rh.IdCashMessenger
*        INNER JOIN View_ClientGroupMax_Filial_Bank AS cfb (READPAST) ON cfb.FilialCode = sp.FilialCode
*        LEFT OUTER JOIN SecurPackDefect AS spd (READPAST) ON spd.IdSecurPack = sp.Id
*        INNER JOIN BudgetDefectNew AS bud (READPAST) ON bud.DefectCode = spd.IdBudgetDefect
*        WHERE sp.IdShift = ?f_IdShift AND sp.IdRoute = ?f_IdRoute AND rh.Id = ?f_IdRouteHistory AND sp.RecIdRoute = ?f_RecIdRoute AND sp.State = 0 AND sp.OperCode = 1 AND sp.Sost = ?status_pak AND sp.FilialCode <> ?skldCode AND sp.IdUser = ?f_IdUser
*        ORDER BY sp.IdUser, IdBank, sp.BudgetNum, sp.TypeValueKind

*      ENDTEXT

*  * ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*    CASE f_IdPost <> 1  && Работают пользователи по должности выше чем Кассир ПРК, например старший кассир, контролер, старший смены

*      TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*        SELECT LEFT(CONVERT(CHAR(8), sp.CurDate, 108), 2) AS CurTime, sp.Id, sp.IdShift, sp.FilialCode, sp.BudgetNum, sp.IdRoute, sp.RecIdRoute, sp.IdUser, sp.TabNumber, sp.IdTypeValue, sp.TypeValueKind,
*        sp.TotalSum, sp.Pr_60322, sp.TotalSum_60322, sp.DateInkas, sp.DtIncome, sp.CurDate, us.UserName, bgh.IdRouteHistory, bgh.IdRouteHistoryMemo, rh.IdCashMessenger, cm.FullName, cm.Fio, sp.Sost, sp.IsTransit, sp.FileName, cfb.FilialName, 0 AS numDef,
*        sp.Flg2PIN, sp.WorkType, sp.SuperBudget, sp.SuperQty, sp.IdJk, cfb.TypeACC, cfb.Account, cfb.Account60322, cfb.UIFactAddress,
*        IIF(sp.IdKeyATM > 0, (SELECT DISTINCT sub.IdKeyATMBank
*        FROM dbo.KeyATMDocSub AS sub
*        INNER JOIN dbo.KeyATM AS ket ON ket.Id = sub.IdKeyATM
*        INNER JOIN dbo.KeyATMBank AS bnk ON bnk.Id = sub.IdKeyATMBank
*        WHERE ket.Id = sp.IdKeyATM), 0) AS IdBank,
*        IIF(sp.IdKeyATM > 0, (SELECT DISTINCT bnk.AtmBankName
*        FROM dbo.KeyATMDocSub AS sub
*        INNER JOIN dbo.KeyATM AS ket ON ket.Id = sub.IdKeyATM
*        INNER JOIN dbo.KeyATMBank AS bnk ON bnk.Id = sub.IdKeyATMBank
*        WHERE ket.Id = sp.IdKeyATM), SPACE(50)) AS BankName,
*        ISNULL((SELECT 1
*        FROM SecurPackDefect AS spd (READPAST)
*        INNER JOIN BudgetDefectNew AS bd (READPAST) ON bd.DefectCode = spd.IdBudgetDefect
*        WHERE spd.IdSecurPack = sp.Id AND bd.IsOpen = 1), 0) AS IsOpen,
*        bud.DefectCode, bud.DefectName
*        FROM SecurPack AS sp (READPAST)
*        INNER JOIN BudgetHistoryPRK AS bgh (READPAST) ON bgh.IdSecurPack = sp.Id
*        INNER JOIN RouteHistory AS rh (READPAST) ON rh.Id = bgh.IdRouteHistory
*        INNER JOIN [User] AS us (READPAST) ON us.Id = rh.IdUser
*        INNER JOIN CashMessenger AS cm (READPAST) ON cm.Id = rh.IdCashMessenger
*        INNER JOIN View_ClientGroupMax_Filial_Bank AS cfb (READPAST) ON cfb.FilialCode = sp.FilialCode
*        LEFT OUTER JOIN SecurPackDefect AS spd (READPAST) ON spd.IdSecurPack = sp.Id
*        INNER JOIN BudgetDefectNew AS bud (READPAST) ON bud.DefectCode = spd.IdBudgetDefect
*        WHERE sp.IdShift = ?f_IdShift AND sp.IdRoute = ?f_IdRoute AND rh.Id = ?f_IdRouteHistory AND sp.RecIdRoute = ?f_RecIdRoute AND sp.State = 0 AND sp.OperCode = 1 AND sp.Sost = ?status_pak AND sp.FilialCode <> ?skldCode
*        ORDER BY sp.IdUser, IdBank, sp.BudgetNum, sp.TypeValueKind

*      ENDTEXT

*  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*** Реализуем печать журналов/актов приёма с разных RecIdRoute (Если сумки ввели, но ведомость не приняли, то потом при добавлении сумок с того же маршрута будут разные RecIdRoute)

SELECT DISTINCT RecIdRoute ;
  FROM SecurPack ;
  ORDER BY RecIdRoute ;
  INTO CURSOR _RecIdRoutes

SELECT _RecIdRoutes
GO TOP

SCAN   && Бегаем по заездам данного маршрута, выбираем сумки и печатаем журналы. ВНУТРИ SCAN'ов НИКАКИХ ВХОДОВ/ВЫХОДОВ!!!!

  f_RecIdRoute = _RecIdRoutes.RecIdRoute    && Где-то внутри идёт вызов проверки даты-смены, которая отрубает нас от сервера.

* ---------------------------------------------------------------- *
  DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*    TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
*      EXECUTE Formir_Priem_Vedom_PRK_New ?f_IdShift, ?f_IdRoute, ?f_IdRouteHistory, ?f_RecIdRoute, ?status_pak, ?f_SkldCode, ?f_IdUser, ?f_IdPost
*    ENDTEXT

*    TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
*      EXECUTE Formir_Priem_Vedom_PRK_60322 ?f_IdShift, ?f_IdRoute, ?f_IdRouteHistory, ?f_RecIdRoute, ?status_pak, ?f_SkldCode, ?f_IdUser, ?f_IdPost
*    ENDTEXT

  TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
    EXECUTE Formir_Priem_Vedom_PRK_Key ?f_IdShift, ?f_IdRoute, ?f_IdRouteHistory, ?f_RecIdRoute, ?status_pak, ?f_SkldCode, ?f_IdUser, ?f_IdPost
  ENDTEXT

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF RunSQL_Load('SecurPack_Sum', @ar) <> 1
    msg = 'По введенным Вами параметрам данные не найдены ... '
    MESSAGEBOX('Таблица - SecurPack_Sum, необходимая для работы НЕ СФОРМИРОВАНА', 16, 'Операция создания ведомости по секюрПакам Прекращена')
    DO exit_sql_server IN RunSQL_Load
    =CloseDB()
    RETURN -1
  ENDIF

* BROWSE WINDOW brows TITLE 'SELECT SecurPack_Sum'

  IF USED('SecurPack_Sum') = .F.
    MESSAGEBOX('Таблица - SecurPack_Sum, необходимая для работы НЕ ОТКРЫТА', 16, 'Операция создания ведомости по секюрПакам Прекращена')
    DO exit_sql_server IN RunSQL_Load
    =CloseDB()
    RETURN -1
  ENDIF

* ---------------------------------------------------------------- *
  DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

  SELECT SecurPack_Sum

  colvo_zap = RECCOUNT()

  GO TOP
* BROWSE WINDOW brows TITLE 'SELECT SecurPack_Sum'

  IF colvo_zap <> 0

    SELECT DISTINCT IdUser ;
      FROM SecurPack_Sum ;
      INTO CURSOR Scan_IdUser ;
      ORDER BY IdUser

* BROWSE WINDOW brows

    SELECT Scan_IdUser
    GO TOP

    SCAN

      Numer_IdUser = Scan_IdUser.IdUser

      SELECT * ;
        FROM SecurPack_Sum ;
        WHERE IdUser = Numer_IdUser ;
        INTO CURSOR SecurPack ;
        ORDER BY IdBank, Flg2PIN, BudgetNum, TypeValueKind

      SELECT SecurPack
      GO TOP

* BROWSE WINDOW brows

      f_TabNumber = ALLTRIM(SecurPack.TabNumber) && Берем из последней строки

* WAIT WINDOW 'f_TabNumber = ' + ALLTRIM(f_TabNumber) TIMEOUT 3

      poisk_TabNumber = ALLTRIM(f_TabNumber)

*      TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*        SELECT chm.*, pos.PostName
*        FROM CashMessenger AS chm (READPAST)
*        INNER JOIN Post AS pos (READPAST) ON pos.Id = chm.IdPost
*        WHERE chm.TabNumber = ?poisk_TabNumber AND chm.IsActive = 1

*      ENDTEXT

* ---------------------------------------------------------------- *
      DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

* -------------------------------------------------------------------------------------------------------------- *

      TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
        EXECUTE Formir_CashMessenger_Post_TabNumber ?poisk_TabNumber
      ENDTEXT

* -------------------------------------------------------------------------------------------------------------- *

      IF RunSQL_Load('CashMessenger', @ar) <> 1
        MESSAGEBOX('Таблица - CashMessenger, необходимая для работы НЕ СФОРМИРОВАНА', 16, 'Операция создания ведомости по секюрПакам Прекращена')
        DO exit_sql_server IN RunSQL_Load
        =CloseDB()
        RETURN -1
      ENDIF

* ---------------------------------------------------------------- *
      DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

      IF USED('CashMessenger') = .T.

        SELECT CashMessenger
        GO TOP

* BROWSE WINDOW brows TITLE 'CashMessenger'

        IncFullName = ALLTRIM(CashMessenger.FullName)
        f_IncFullName = ALLTRIM(CashMessenger.FullName)
        f_IncPostName = ALLTRIM(CashMessenger.PostName)

        USE

      ELSE

        MESSAGEBOX('Таблица - CashMessenger, необходимая для работы ПУСТАЯ', 16, 'Операция отгрузки ведомости по секюрПакам Прекращена')
        DO exit_sql_server IN RunSQL_Load
        =CloseDB()
        RETURN -1

      ENDIF

      msg = ''

      sel_Sost = status_pak

      Vedom_prn = .T.  && Признак выдачи отчета на принтер, когда идет первичная печать
      Memo_prn = .T.  && Признак выдачи отчета на принтер, когда идет первичная печать

      pr_rabota_arxiv = .F.  && Признак работы в текущей смене и опердне

* ------------------------------------------------------------------------------------ *
      DO Formir_Table_SecurPack IN Priem_Vedom_PRK
* ------------------------------------------------------------------------------------ *

      SELECT Scan_IdUser
    ENDSCAN

    f_IdShiftList = ALLTRIM(STR(f_IdShift_1)) + ", ''"

* ------------------------------------------------------------------------ *
    DO VedomTsd_Inkass_0 IN TakeVedomPrn
* ------------------------------------------------------------------------ *

  ELSE

    MESSAGEBOX('Таблица - SecurPack_Sum, необходимая для работы ПУСТАЯ', 16, 'Операция создания ведомости по секюрПакам Прекращена')
    =CloseDB()
    RETURN -1

  ENDIF

  SELECT _RecIdRoutes
ENDSCAN

Stop_DateTimeRoute = DATETIME()

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7
  EXECUTE sp_Finish_RoutePRK_Time ?f_IdRouteHistory, ?Stop_DateTimeRoute
ENDTEXT

=RunSQL('', @ar) <> 1

=CloseDB()

RETURN


**************************************************************************** PROCEDURE Formir_Table_SecurPack ********************************************************************

PROCEDURE Formir_Table_SecurPack

SELECT SecurPack
GO TOP

* BROWSE WINDOW brows TITLE 'SecurPack'

* -------------------------------------------------------------------- Обрабатываем МНОГОПИНОВЫЕ и ОДНОПИНОВЫЕ Секьюрпаки с ПЕРЕСЧЁТОМ рублевые IdRoute <> 999 ----------------------------------------------------------------------- *

SELECT SecurPack
GO TOP

* BROWSE WINDOW brows TITLE 'Обрабатываем МНОГОПИНОВЫЕ Секьюрпаки с ПЕРЕСЧЁТОМ рублевые'

SELECT '1' AS GrpCoarse, ICASE(sp.IdBank = 1, 'X5  ', sp.IdBank = 6, 'МФ  ', sp.IdBank = 7, 'ВБРР', 'П   ') AS GrpShort,;
  PADR('Секьюрпаки с ПЕРЕСЧЁТОМ', 80) AS grName, f_IncFullName AS IncFullName, sp.* ;
  FROM SecurPack AS sp ;
  WHERE ALLTRIM(sp.Sost) = ALLTRIM(sel_Sost) AND sp.IsTransit = '0' AND sp.WorkType = '0' AND ALLTRIM(sp.TypeValueKind) == 'RUB' AND sp.IdRoute <> 999 AND sp.Flg2PIN < 3 ;
  INTO CURSOR Sum_SecurPackPRK READWRITE ;
  ORDER BY sp.IdBank, sp.RecIdRoute, sp.BudgetNum, sp.TypeValueKind

* BROWSE WINDOW brows TITLE 'Обрабатываем МНОГОПИНОВЫЕ Секьюрпаки с ПЕРЕСЧЁТОМ рублевые'

IF _TALLY <> 0

  SELECT Sum_SecurPackPRK

  REPLACE ALL IdBank WITH 17

  GO TOP

  f_IsTransit = Sum_SecurPackPRK.IsTransit

*  f_IdKeyATMDocMain = Sum_SecurPackPRK.IdKeyATMDocMain
*  f_IdKeyATM = Sum_SecurPackPRK.IdKeyATM

* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки с ПЕРЕСЧЁТОМ рублевые#1'

*  В этом и следующем SELECT замена PADL(ALLTRIM(STR(IdBank)), 5, '0') на GrpCoarse

  SELECT DISTINCT GrpCoarse + ' ' + ALLTRIM(TypeValueKind) + ' ' + PADL(ALLTRIM(STR(IdBank)), 2, '0') AS IdBank ;
    FROM Sum_SecurPackPRK AS sp ;
    INTO CURSOR Scan_IdBank ;
    ORDER BY 1

  SELECT Scan_IdBank
  GO TOP

* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 2'

  SCAN

    rec_IdBank = Scan_IdBank.IdBank

    SELECT sp.* ;
      FROM Sum_SecurPackPRK AS sp ;
      WHERE GrpCoarse + ' ' + ALLTRIM(TypeValueKind) + ' ' + PADL(ALLTRIM(STR(IdBank)), 2, '0') == ALLTRIM(rec_IdBank)  ;
      INTO CURSOR SecurPackPRK ;
      ORDER BY sp.IdBank, sp.grName, sp.BudgetNum

    SELECT SecurPackPRK
* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 3'

    IF _TALLY <> 0

      SELECT DISTINCT Flg2PIN ;
        FROM SecurPackPRK ;
        INTO CURSOR Flg2PIN_SecurPack ;
        ORDER BY Flg2PIN ASC

* BROWSE WINDOW brows

      colvo_Flg2PIN = RECCOUNT()

      DO CASE
        CASE colvo_Flg2PIN = 1

          pole_Flg2PIN = Flg2PIN_SecurPack.Flg2PIN

          DO Formir_Table_Prn_2PIN IN Priem_Vedom_PRK WITH pole_Flg2PIN

        CASE colvo_Flg2PIN > 1

          DO Formir_Table_Prn_2PIN IN Priem_Vedom_PRK WITH 1

      ENDCASE

    ENDIF

    SELECT Scan_IdBank
  ENDSCAN

  f_TypeValueKind = 'RUB'

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
  DO start IN Memoorder_Tranzit_Key  &&   && Формируем таблицу для сканирования, по ключам АТМ будут формироватся и записываться мемоордера в отдельную таблицу RouteHistoryMemoKEY
* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

ELSE

  colvo_key_4 = 0
  colvo_key_3 = 0

  SELECT SecurPack
  GO TOP

  COUNT TO colvo_key_4 FOR Flg2Pin = 4 AND WorkType = '0'
  COUNT TO colvo_key_3 FOR Flg2Pin = 3 AND WorkType = '0'

  DO CASE
    CASE colvo_key_4 > 0 AND colvo_key_3 = 0

      =err('Внимание! в Журнале имеются записи только по потерянным ключам ... ')

    CASE colvo_key_4 = 0 AND colvo_key_3 > 0

      =err('Внимание! в Журнале имеются записи только по сданным ключам без сумок ... ')
      DO start IN Memoorder_Tranzit_Key

    CASE colvo_key_4 > 0 AND colvo_key_3 > 0

      =err('Внимание! в Журнале имеются записи по потерянным ключам и по сданным ключам без сумок ... ')
      DO start IN Memoorder_Tranzit_Key

  ENDCASE

ENDIF


* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* -------------------------------------------------------------------- Обрабатываем ОДНОПИНОВЫЕ Секьюрпаки с ПЕРЕСЧЁТОМ рублевые IdRoute = 999 -------------------------------------------------------------------------------------------- *

SELECT SecurPack
GO TOP

* BROWSE WINDOW brows TITLE 'Монетная инкассация'

SELECT IIF(sp.IdBank = 3, '0', '1') AS GrpCoarse, ICASE(sp.IdBank = 1, 'X5  ', sp.IdBank = 6, 'МФ  ', sp.IdBank = 7, 'ВБРР', 'П   ') AS GrpShort,;
  PADR('Секьюрпаки с ПЕРЕСЧЁТОМ' + ' - ' + ALLTRIM(sp.BankName), 80) AS grName, f_IncFullName AS IncFullName, sp.* ;
  FROM SecurPack AS sp ;
  WHERE ALLTRIM(sp.Sost) = ALLTRIM(sel_Sost) AND sp.IsTransit = '0' AND sp.WorkType = '0' AND ALLTRIM(sp.TypeValueKind) == 'RUB' AND sp.IdRoute = 999 ;
  INTO CURSOR Sum_SecurPackPRK READWRITE ;
  ORDER BY sp.IdBank, sp.RecIdRoute, sp.BudgetNum, sp.TypeValueKind

IF _TALLY <> 0

  SELECT Sum_SecurPackPRK

  REPLACE ALL IdBank WITH 17

  GO TOP

  f_IsTransit = Sum_SecurPackPRK.IsTransit

* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки с ПЕРЕСЧЁТОМ рублевые#1'

*  В этом и следующем SELECT замена PADL(ALLTRIM(STR(IdBank)), 5, '0') на GrpCoarse

  SELECT DISTINCT GrpCoarse + ' ' + ALLTRIM(TypeValueKind) + ' ' + PADL(ALLTRIM(STR(RecIdRoute)), 2, '0') AS IdBank ;
    FROM Sum_SecurPackPRK AS sp ;
    INTO CURSOR Scan_IdBank ;
    ORDER BY 1

  SELECT Scan_IdBank
  GO TOP

* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 2'

  SCAN

    rec_IdBank = Scan_IdBank.IdBank

    SELECT sp.* ;
      FROM Sum_SecurPackPRK AS sp ;
      WHERE GrpCoarse + ' ' + ALLTRIM(TypeValueKind) + ' ' + PADL(ALLTRIM(STR(RecIdRoute)), 2, '0') == ALLTRIM(rec_IdBank)  ;
      INTO CURSOR SecurPackPRK ;
      ORDER BY sp.IdBank, sp.grName, sp.BudgetNum

    SELECT SecurPackPRK
* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 3'

    IF _TALLY <> 0

      DO Formir_Table_Prn_999 IN Priem_Vedom_PRK

    ENDIF

    SELECT Scan_IdBank
  ENDSCAN

  f_TypeValueKind = 'RUB'

ENDIF


* ---------------------------------------------------------------------------- Обрабатываем СТАНДАРТНЫЕ Супер Сумки с ПЕРЕСЧЁТОМ рублевые ------------------------------------------------------------------------------------ *

SELECT SecurPack
GO TOP

* BROWSE WINDOW brows TITLE 'Обрабатываем Супер Сумки'

*** Pr_60322 - актуален на март 2017 только для суперсумок. Если = 1 то это россыпь с пола АДМ, идущая отдельной сумкой, не учитываемой в журнале, но на которую печатается приходник
*** 26.02.2018 - Сумки с россыпью учитываются в 301 журнале с признаком (ПКО)

SELECT PADR('Секьюрпаки с ПЕРЕСЧЁТОМ стандартные' + ' - ' + ALLTRIM(sp.BankName), 80) AS grName, f_IncFullName AS IncFullName, sp.* ;
  FROM SecurPack AS sp ;
  WHERE ALLTRIM(sp.Sost) = ALLTRIM(sel_Sost) AND sp.IsTransit = '0' AND INLIST(sp.WorkType, '1', '2') AND ALLTRIM(sp.TypeValueKind) == 'RUB'  AND sp.Flg2PIN < 3 ;   &&AND sp.Pr_60322 = 0
INTO CURSOR Sum_SecurPackPRK READWRITE ;
  ORDER BY sp.IdBank, sp.RecIdRoute, sp.WorkType, sp.BudgetNum, sp.TypeValueKind

colvo_zap = _TALLY  && Считаем ВСЕ выбранные записи по условию INLIST(sp.WorkType, '1', '2') AND sp.Flg2PIN < 3

* BROWSE WINDOW brows TITLE 'Обрабатываем Супер Сумки'

SELECT DISTINCT WorkType ;
  FROM Sum_SecurPackPRK ;
  INTO CURSOR Dis_WorkType ;
  ORDER BY WorkType

colvo_zap_WorkType = _TALLY  && Считаем ВСЕ выбранные записи по условию INLIST(sp.WorkType, '1', '2'), для правильного построения отчета должно быть записей >= 2, то есть оба значения WorkType, '1', '2'

* BROWSE WINDOW brows

IF colvo_zap > 0 AND colvo_zap_WorkType > 1

  SELECT Sum_SecurPackPRK
  GO TOP

  f_IsTransit = Sum_SecurPackPRK.IsTransit

* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 1'

  SELECT DISTINCT PADL(ALLTRIM(STR(IdBank)), 5, '0') + ' ' + ALLTRIM(TypeValueKind) + ' ' + PADL(ALLTRIM(STR(RecIdRoute)), 2, '0') AS IdBank ;
    FROM Sum_SecurPackPRK AS sp ;
    INTO CURSOR Scan_IdBank ;
    ORDER BY 1

  SELECT Scan_IdBank
  GO TOP
* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 2'

  SCAN

    rec_IdBank = Scan_IdBank.IdBank

    SELECT sp.* ;
      FROM Sum_SecurPackPRK AS sp ;
      WHERE PADL(ALLTRIM(STR(IdBank)), 5, '0') + ' ' + ALLTRIM(TypeValueKind) + ' ' + PADL(ALLTRIM(STR(RecIdRoute)), 2, '0') == ALLTRIM(rec_IdBank) ;  && AND Flg2PIN < 3
    INTO CURSOR SecurPackPRK ;
      ORDER BY sp.IdBank, sp.RecIdRoute, sp.Superbudget,sp.WorkType,sp.BudgetNum, sp.TypeValueKind

    SELECT SecurPackPRK
* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 3'

    IF _TALLY <> 0
      DO Formir_Table_Prn IN Priem_Vedom_PRK WITH 3
    ENDIF

    SELECT Scan_IdBank
  ENDSCAN

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
  DO start IN Memoorder_Tranzit_Key  &&   && Формируем таблицу для сканирования, по ключам АТМ будут формироватся и записываться мемоордера в отдельную таблицу RouteHistoryMemoKEY
* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

ELSE

  colvo_key_4 = 0
  colvo_key_3 = 0

  SELECT SecurPack
  GO TOP

  COUNT TO colvo_key_4 FOR Flg2Pin = 4 AND WorkType > '0'
  COUNT TO colvo_key_3 FOR Flg2Pin = 3 AND WorkType > '0'

  DO CASE
    CASE colvo_key_4 > 0 AND colvo_key_3 = 0

      =err('Внимание! в Журнале имеются записи только по потерянным ключам ... ')

    CASE colvo_key_4 = 0 AND colvo_key_3 > 0

      =err('Внимание! в Журнале имеются записи только по сданным ключам без сумок ... ')
      DO start IN Memoorder_Tranzit_Key

    CASE colvo_key_4 > 0 AND colvo_key_3 > 0

      =err('Внимание! в Журнале имеются записи по потерянным ключам и по сданным ключам без сумок ... ')
      DO start IN Memoorder_Tranzit_Key

  ENDCASE

ENDIF


* ---------------------------------------------------------------------------- Обрабатываем изьятые купюры с пола АДМ с ПЕРЕСЧЁТОМ рублевые ------------------------------------------------------------------------------------ *

SELECT SecurPack
GO TOP
* BROWSE WINDOW brows TITLE 'Обрабатываем изьятые купюры с пола АДМ с ПЕРЕСЧЁТОМ рублевые'

SELECT PADR('Секьюрпаки с ПЕРЕСЧЁТОМ стандартные' + ' - ' + ALLTRIM(sp.BankName), 80) AS grName, f_IncFullName AS IncFullName, sp.* ;
  FROM SecurPack AS sp ;
  WHERE ALLTRIM(sp.Sost) = ALLTRIM(sel_Sost) AND sp.IsTransit = '0'  AND ALLTRIM(sp.TypeValueKind) == 'RUB' AND sp.Pr_60322 = 1 ;  &&AND sp.WorkType = '2'
INTO CURSOR Sum_SecurPackPRK READWRITE ;
  ORDER BY sp.IdBank, sp.RecIdRoute, sp.BudgetNum, sp.TypeValueKind

* BROWSE WINDOW brows TITLE 'Обрабатываем изьятые купюры с пола АДМ с ПЕРЕСЧЁТОМ рублевые'

IF _TALLY <> 0

  SELECT Sum_SecurPackPRK
  GO TOP

* REPLACE Account WITH NVL(Account60322,'') ALL   && ??

* REPLACE Account WITH IIF(ALLTRIM(TypeAcc) == '11', ALLTRIM(Account60322), ALLTRIM(Account)) ALL

  REPLACE FOR ALLTRIM(Sum_SecurPackPRK.TypeAcc) == '11' Account WITH ALLTRIM(Sum_SecurPackPRK.Account60322)

  GO TOP

  f_IsTransit = Sum_SecurPackPRK.IsTransit

* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 1'

  SELECT DISTINCT PADL(ALLTRIM(STR(IdBank)), 5, '0') + ' ' + ALLTRIM(TypeValueKind) + ' ' + PADL(ALLTRIM(STR(RecIdRoute)), 2, '0') AS IdBank ;
    FROM Sum_SecurPackPRK AS sp ;
    INTO CURSOR Scan_IdBank ;
    ORDER BY 1

  SELECT Scan_IdBank
  GO TOP
* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 2'

  SCAN

    rec_IdBank = Scan_IdBank.IdBank

    SELECT sp.* ;
      FROM Sum_SecurPackPRK AS sp ;
      WHERE PADL(ALLTRIM(STR(IdBank)), 5, '0') + ' ' + ALLTRIM(TypeValueKind) + ' ' + PADL(ALLTRIM(STR(RecIdRoute)), 2, '0') == ALLTRIM(rec_IdBank)  ;
      INTO CURSOR SecurPackPRK ;
      ORDER BY sp.IdBank, sp.RecIdRoute, sp.Superbudget, sp.WorkType, sp.BudgetNum, sp.TypeValueKind

    SELECT SecurPackPRK
* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 3'

    IF _TALLY <> 0
      DO Formir_Table_Prn_ADM IN Priem_Vedom_PRK WITH 0
    ENDIF

    SELECT Scan_IdBank
  ENDSCAN

  f_TypeValueKind = 'RUB'

ENDIF


* ---------------------------------------------------------------------------- Обрабатываем пакеты из спец. контейнера залитые краской ------------------------------------------------------------------------------------ *

SELECT SecurPack
GO TOP
* BROWSE WINDOW brows TITLE 'Обрабатываем изьятые купюры с пола АДМ с ПЕРЕСЧЁТОМ рублевые'

SELECT PADR('Секьюрпаки из спец. контейнера залитые краской' + ' - ' + ALLTRIM(sp.BankName), 80) AS grName, f_IncFullName AS IncFullName, sp.* ;
  FROM SecurPack AS sp ;
  WHERE ALLTRIM(sp.Sost) = ALLTRIM(sel_Sost) AND sp.IsTransit = '0' AND sp.IdPackDefect = 23 AND INLIST(sp.WorkType, '0', '2') = .T. ;
  INTO CURSOR Sum_SecurPackPRK READWRITE ;
  ORDER BY sp.IdBank, sp.WorkType, sp.BudgetNum, sp.TypeValueKind

* BROWSE WINDOW brows TITLE 'Обрабатываем изьятые купюры с пола АДМ с ПЕРЕСЧЁТОМ рублевые'

IF _TALLY <> 0

  SELECT Sum_SecurPackPRK
  GO TOP

  f_IsTransit = Sum_SecurPackPRK.IsTransit
  f_IdRouteHistory = Sum_SecurPackPRK.IdRouteHistory

* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 1'

  SELECT DISTINCT PADL(ALLTRIM(STR(IdBank)), 5, '0') + ' ' + ALLTRIM(TypeValueKind) AS IdBank ;
    FROM Sum_SecurPackPRK AS sp ;
    INTO CURSOR Scan_IdBank ;
    ORDER BY 1

  SELECT Scan_IdBank
  GO TOP
* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 2'

  SCAN

    rec_IdBank = Scan_IdBank.IdBank

    SELECT sp.* ;
      FROM Sum_SecurPackPRK AS sp ;
      WHERE PADL(ALLTRIM(STR(IdBank)), 5, '0') + ' ' + ALLTRIM(TypeValueKind) == ALLTRIM(rec_IdBank)  ;
      INTO CURSOR SecurPackPRK ;
      ORDER BY sp.IdBank, sp.WorkType, sp.BudgetNum, sp.TypeValueKind

    SELECT SecurPackPRK
* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки ПЕРЕСЧЁТОМ рублевые № 3'

    IF _TALLY <> 0
      DO Formir_Table_Prn_Kraska IN Priem_Vedom_PRK WITH 0
    ENDIF

    SELECT Scan_IdBank
  ENDSCAN

  f_TypeValueKind = 'RUB'

ENDIF


* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF USED('Sum_SecurPackPRK') = .T.
  SELECT Sum_SecurPackPRK
  USE
ENDIF

IF USED('SecurPackPRK') = .T.
  SELECT SecurPackPRK
  USE
ENDIF

IF USED('Scan_IdBank') = .T.
  SELECT Scan_IdBank
  USE
ENDIF

RETURN


***************************************************************** PROCEDURE Formir_Table_Prn_2PIN **********************************************************

PROCEDURE Formir_Table_Prn_2PIN
PARAMETERS Priznak_2PIN

IF TYPE('Priznak_2PIN') = 'U'
  Priznak_2PIN = 0
ENDIF

* WAIT WINDOW 'Priznak_2PIN = ' + ALLTRIM(STR(Priznak_2PIN)) TIMEOUT 3

PRIVATE f_IsOpen_1_pred, f_DefectName_1_pred, f_IsOpen_2_pred, f_DefectName_2_pred

STORE 0 TO f_IsOpen_1_pred, f_IsOpen_2_pred
STORE SPACE(0) TO f_DefectName_1_pred, f_DefectName_2_pred

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

Itog_IdTypeValue = SecurPackPRK.IdTypeValue

lenBNum=LEN(SecurPackPRK.BudgetNum)

SELECT DISTINCT IIF(AT('/', sp.BudgetNum) > 0, PADR(LEFT(sp.BudgetNum, AT('/', sp.BudgetNum) - 1),lenBNum), sp.BudgetNum) ;
  FROM SecurPackPRK sp ;
  INTO CURSOR Count_BudgetNum

Itog_BudgetNum=_TALLY
**COUNT TO Itog_BudgetNum  && Подсчет количества сумок

* WAIT WINDOW 'Itog_BudgetNum = ' + ALLTRIM(TRANSFORM(Itog_BudgetNum, '999.99')) &&TIMEOUT 3

* BROWSE WINDOW brows

SELECT DISTINCT BudgetNum ;
  FROM SecurPackPRK ;
  INTO CURSOR Count_BuhVedom NOFILTER

Itog_BuhVedom=_TALLY
**COUNT TO Itog_BuhVedom  && Подсчет количества сопроводительных ведомостей

* WAIT WINDOW 'Itog_BuhVedom = ' + ALLTRIM(TRANSFORM(Itog_BuhVedom, '999.99')) &&TIMEOUT 3

* BROWSE WINDOW brows

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

SUM TotalSum TO Itog_TotalSum  && Подсчет общей суммы
SUM TotalSum TO f_TotalSum60322 FOR Pr_60322=1

DO start IN Propis_chislo WITH Itog_BudgetNum  && Формирование количества сумок прописью
Itog_BudgetNum_Prop = Chislo_Prop

DO start IN Propis_chislo WITH Itog_BuhVedom  && Формирование количества ведомостей прописью
Itog_BuhVedom_Prop = Chislo_Prop

DO start IN PropisVal WITH Itog_TotalSum, Itog_IdTypeValue  && Формирование общей суммы прописью
Itog_TotalSum_Prop = Summa_Prop

CREATE CURSOR SecurPackPRK_cursor ;
  (GrName CHAR(80) NULL, IncFullName CHAR(40) NULL, CURDATE DATETIME, CURTIME CHAR(2) NULL, PacketNum_1 INTEGER, ;
  BudgetNum_1 CHAR(20) NULL, TotalSum_1 Numeric(15, 2), TypeValueKind_1 CHAR(80) NULL, BdgType_1 CHAR(1), GrpShort_1 CHAR(4),;
  PacketNum_2 INTEGER, BudgetNum_2 CHAR(20) NULL, TotalSum_2 Numeric(15, 2), TypeValueKind_2 CHAR(80) NULL, BdgType_2 CHAR(1), GrpShort_2 CHAR(4), ;
  Sost INTEGER, IdBank INTEGER, UserName CHAR(20) NULL, IdRouteHistory INTEGER, ;
  IsOpen_1 INTEGER DEFAULT 0, DefectName_1 CHAR(50), IsOpen_2 INTEGER DEFAULT 0, DefectName_2 CHAR(50), PKO_1 INTEGER, PKO_2 INTEGER)

SELECT SecurPackPRK
colvo_zap = RECCOUNT()
GO TOP
* BROWSE WINDOW brows

colvo_zap_polovina = ROUND(colvo_zap / 2, 0)

* WAIT WINDOW 'colvo_zap_polovina = ' + ALLTRIM(TRANSFORM(colvo_zap_polovina, '999.99')) TIMEOUT 3

DO CASE
  CASE MOD(colvo_zap_polovina, 2) > 0

* WAIT WINDOW 'Добавляются строки по НЕ ЧЕТНОМУ варианту' TIMEOUT 2

    SELECT SecurPackPRK_cursor

    FOR I = 1 TO colvo_zap_polovina + 1
      APPEND BLANK
    ENDFOR

    colvo_zap_polovina = colvo_zap_polovina + 1


  CASE MOD(colvo_zap_polovina, 2) = 0

* WAIT WINDOW 'Добавляются строки по ЧЕТНОМУ варианту' TIMEOUT 2

    SELECT SecurPackPRK_cursor

    FOR I = 1 TO colvo_zap_polovina
      APPEND BLANK
    ENDFOR

    colvo_zap_polovina = colvo_zap_polovina + 0

ENDCASE

SELECT SecurPackPRK_cursor
* BROWSE WINDOW brows

f_IdRoute = SecurPackPRK.IdRoute

* ------------------------------------------------------------------------------------------------ *

SELECT SecurPackPRK_cursor
GO TOP
* BROWSE WINDOW brows

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

f_BudgetNum_pred = ''

rec_num = 0

m.CurrencyName = SecurPackPRK.TypeValueKind    &&Текущий вид валюты для печати в начале отчёта
m.CurrentGrpCoarse = SecurPackPRK.GrpCoarse    &&Текущая общая группа (гипермаркеты или нет), необходимо для печати "Гипермаркеты" в отдельном отчёте

SCAN

  rec_num = rec_num + 1

  SELECT SecurPackPRK_cursor
* BROWSE WINDOW brows TITLE 'SCAN'

  SCATTER MEMVAR

  m.GrName = SecurPackPRK.GrName
  m.IncFullName = SecurPackPRK.IncFullName
  m.CurDate = SecurPackPRK.CURDATE

  m.Sost = VAL(ALLTRIM(SecurPackPRK.Sost))

  m.IdBank = IIF(SecurPackPRK.IdBank = 0, '17', PADL(ALLTRIM(STR(SecurPackPRK.IdBank)), 2, '0'))

  m.CurTime = SecurPackPRK.CURTIME

  m.BdgType_1 = 'М'  && В любом случае (Flg2PIN > 0) Сумка многопиновая
  m.GrpShort_1 = SecurPackPRK.GrpShort  && Краткое название группы
  m.PKO_1=SecurPackPRK.Pr_60322    &&  Признак россыпи

  DO CASE
    CASE SecurPackPRK.Flg2PIN = 0

      m.BdgType_1 = 'С'  && Стандартная сумка
      m.PacketNum_1 = rec_num

      m.BudgetNum_1 = SecurPackPRK.BudgetNum

      m.TotalSum_1 = SecurPackPRK.TotalSum

      m.TypeValueKind_1 = ' (' + ALLTRIM(SecurPackPRK.FilialCode)  && Закрывающая скобка будет проставлена в отдельном SCAN'е по табличке

      m.IsOpen_1 = IIF((m.IsOpen_1 = 1) OR (SecurPackPRK.IsOpen = 1), 1, 0)  && Если хоть на одну из записей многопиновой сумки есть акт вскрытия, то пишем Акт
      m.DefectName_1 = IIF(m.IsOpen_1 = 0, ALLTRIM(SecurPackPRK.DefectName), '')

      IF VARTYPE(SecurPackPRK.IdJk) = 'N' AND SecurPackPRK.IdJk = (-1)
        m.TypeValueKind_1 = 'ЯK отсутствует. ' + ALLTRIM(m.TypeValueKind_1)
      ENDIF



    CASE SecurPackPRK.Flg2PIN = 1

      m.PacketNum_1 = rec_num

      m.BudgetNum_1 = SecurPackPRK.BudgetNum

      m.PacketNum_1 = IIF(ALLTRIM(m.BudgetNum_1) <> ALLTRIM(f_BudgetNum_pred), rec_num, 0)

      m.TotalSum_1 = SecurPackPRK.TotalSum

      m.TypeValueKind_1 = ' (' + ALLTRIM(SecurPackPRK.FilialCode)

      m.IsOpen_1 = IIF((m.IsOpen_1 = 1) OR (SecurPackPRK.IsOpen = 1), 1, 0)  && Если хоть на одну из записей многопиновой сумки есть акт вскрытия, то пишем Акт
      m.DefectName_1 = IIF(m.IsOpen_1 = 0, ALLTRIM(SecurPackPRK.DefectName), '')

      f_IsOpen_1_pred = m.IsOpen_1
      f_DefectName_1_pred = m.DefectName_1

      IF VARTYPE(SecurPackPRK.IdJk) = 'N' AND SecurPackPRK.IdJk = (-1)
        m.TypeValueKind_1 = 'ЯK отсутствует. ' + ALLTRIM(m.TypeValueKind_1)
      ENDIF



    CASE SecurPackPRK.Flg2PIN > 1

      m.BudgetNum_1 = ALLTRIM(SecurPackPRK.BudgetNum)

      m.BudgetNum_1 = STRTRAN(m.BudgetNum_1, '/2', '')
      m.BudgetNum_1 = STRTRAN(m.BudgetNum_1, '/3', '')

* WAIT WINDOW 'm.BudgetNum_1 = ' + ALLTRIM(m.BudgetNum_1) + '   f_BudgetNum_pred = ' + ALLTRIM(f_BudgetNum_pred) TIMEOUT 3

      DO CASE
        CASE ALLTRIM(m.BudgetNum_1) == ALLTRIM(f_BudgetNum_pred)

          rec_num = rec_num - 1

          m.PacketNum_1 = IIF(ALLTRIM(m.BudgetNum_1) == ALLTRIM(f_BudgetNum_pred), rec_num, 0)  && Смысл? мы тут будем только при истинности условия

          m.TotalSum_1 = f_TotalSum_pred + SecurPackPRK.TotalSum

          m.TypeValueKind_1 = IIF(EMPTY(ALLTRIM(f_TypeValueKind_pred)) = .F., ALLTRIM(f_TypeValueKind_pred) + ', ' + ALLTRIM(SecurPackPRK.FilialCode), '(' + ALLTRIM(SecurPackPRK.FilialCode))

          m.IsOpen_1 = f_IsOpen_1_pred
          m.DefectName_1 = f_DefectName_1_pred


          IF BOF() = .F.

            SKIP -1

            IF BOF() = .T.
              SKIP +1
            ENDIF

          ENDIF

        CASE ALLTRIM(m.BudgetNum_1) <> ALLTRIM(f_BudgetNum_pred)

          m.PacketNum_1 = rec_num

          m.TotalSum_1 = SecurPackPRK.TotalSum

          m.TypeValueKind_1 = ' (' + ALLTRIM(SecurPackPRK.FilialCode)

          IF VARTYPE(SecurPackPRK.IdJk) = 'N' AND SecurPackPRK.IdJk = (-1)
            m.TypeValueKind_1 = 'ЯK отсутствует. ' + ALLTRIM(m.TypeValueKind_1)
          ENDIF

      ENDCASE

  ENDCASE

* WAIT WINDOW 'm.PacketNum_1 = ' + ALLTRIM(TRANSFORM(m.PacketNum_1, '999.99')) TIMEOUT 1

  m.UserName = SecurPackPRK.UserName

  m.IdRouteHistory = SecurPackPRK.IdRouteHistory

  SELECT SecurPackPRK_cursor

  IF BOF() = .F. AND EOF() = .F.
    GATHER MEMVAR
  ENDIF

* BROWSE WINDOW brows

  SELECT SecurPackPRK
  rec_new = RECNO()

  SELECT SecurPackPRK_cursor

* BROWSE WINDOW brows

  IF EOF() = .F.
    SKIP
  ENDIF

* WAIT WINDOW 'rec_new = ' + ALLTRIM(TRANSFORM(rec_new, '999.99')) + '   colvo_zap_polovina = ' + ALLTRIM(TRANSFORM(colvo_zap_polovina, '999.99')) TIMEOUT 3

*  BROWSE WINDOW brows

  IF rec_new >= colvo_zap_polovina

    SELECT SecurPackPRK
    SKIP

*  BROWSE WINDOW brows

    DO CASE
      CASE LIKE(ALLTRIM(m.BudgetNum_1) + '??',  ALLTRIM(SecurPackPRK.BudgetNum)) = .T.

        SKIP -1

      CASE LIKE(ALLTRIM(m.BudgetNum_1) + '??',  ALLTRIM(SecurPackPRK.BudgetNum)) = .F.

        SKIP -1
        EXIT

    ENDCASE

  ENDIF

  IF EOF() = .T.
    EXIT
  ENDIF

  f_BudgetNum_pred = ALLTRIM(m.BudgetNum_1)

  f_TotalSum_pred = m.TotalSum_1

  f_TypeValueKind_pred = ALLTRIM(m.TypeValueKind_1)

  SELECT SecurPackPRK
ENDSCAN

*      m.TypeValueKind_1 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_1)) = .F., ALLTRIM(m.TypeValueKind_1) + '; ', '') + ;
IIF(EMPTY(ALLTRIM(m.BdgType_1)) = .F., '; ' + m.BdgType_1, '') + ;
IIF(EMPTY(ALLTRIM(m.GrpShort_1)) = .F., '; ' + m.GrpShort_1, '') + ;
IIF(m.IsOpen_1 = 1, '; ' + 'Акт', IIF(EMPTY(ALLTRIM(m.DefectName_1)) = .F., '; ' + ALLTRIM(m.DefectName_1), ''))

SELECT SecurPackPRK_cursor
* BROWSE WINDOW brows

f_BudgetNum_pred = ''
f_TotalSum_pred = 0.00
f_TypeValueKind_pred = ''

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

* WAIT WINDOW 'Остановились на записи в левой части = ' + ALLTRIM(STR(rec_new)) TIMEOUT 3
* WAIT WINDOW 'Порядковый номер записи = ' + ALLTRIM(STR(rec_num)) TIMEOUT 3

SELECT SecurPackPRK_cursor
GO TOP
* BROWSE WINDOW brows

SELECT SecurPackPRK
* BROWSE WINDOW brows

SCAN FOR RECNO() >= rec_new + 1

  rec_num = rec_num + 1

  SELECT SecurPackPRK_cursor

  SCATTER MEMVAR FIELDS PacketNum_2, BudgetNum_2, TotalSum_2, TypeValueKind_2, BdgType_2, GrpShort_2, IsOpen_2,PKO_2

  m.BdgType_2 = 'М' && В любом случае (Flg2PIN > 0) Сумка многопиновая
  m.GrpShort_2 = SecurPackPRK.GrpShort  && Краткое название группы
  m.PKO_2=SecurPackPRK.Pr_60322

  DO CASE
    CASE SecurPackPRK.Flg2PIN = 0

      m.BdgType_2 = 'С'  && Стандартная сумка

      m.PacketNum_2 = rec_num

      m.BudgetNum_2 = SecurPackPRK.BudgetNum

      m.TotalSum_2 = SecurPackPRK.TotalSum

      m.TypeValueKind_2 = ' (' + ALLTRIM(SecurPackPRK.FilialCode)  && Закрывающая скобка будет проставлена в отдельном SCAN'е по табличке

      m.IsOpen_2 = IIF((m.IsOpen_2 = 1) OR (SecurPackPRK.IsOpen = 1), 1, 0)
      m.DefectName_2 = IIF(m.IsOpen_2 = 0, ALLTRIM(SecurPackPRK.DefectName), '')

      IF VARTYPE(SecurPackPRK.IdJk) = 'N' AND SecurPackPRK.IdJk = (-1)
        m.TypeValueKind_2 = 'ЯK отсутствует. ' + ALLTRIM(m.TypeValueKind_2)
      ENDIF



    CASE SecurPackPRK.Flg2PIN = 1

      m.PacketNum_2 = rec_num

      m.BudgetNum_2 = SecurPackPRK.BudgetNum

      m.PacketNum_2 = IIF(ALLTRIM(m.BudgetNum_2) <> ALLTRIM(f_BudgetNum_pred), rec_num, 0)

      m.TotalSum_2 = SecurPackPRK.TotalSum

      m.TypeValueKind_2 = ' (' + ALLTRIM(SecurPackPRK.FilialCode)

      m.IsOpen_2 = IIF((m.IsOpen_2 = 1) OR (SecurPackPRK.IsOpen = 1), 1, 0)
      m.DefectName_2 = IIF(m.IsOpen_2 = 0, ALLTRIM(SecurPackPRK.DefectName), '')

      f_IsOpen_2_pred = m.IsOpen_2
      f_DefectName_2_pred = m.DefectName_2

      IF VARTYPE(SecurPackPRK.IdJk) = 'N' AND SecurPackPRK.IdJk = (-1)
        m.TypeValueKind_2 = 'ЯK отсутствует. ' + ALLTRIM(m.TypeValueKind_2)
      ENDIF



    CASE SecurPackPRK.Flg2PIN > 1

      m.BudgetNum_2 = ALLTRIM(SecurPackPRK.BudgetNum)

      m.BudgetNum_2 = STRTRAN(m.BudgetNum_2, '/2', '')
      m.BudgetNum_2 = STRTRAN(m.BudgetNum_2, '/3', '')

      DO CASE
        CASE ALLTRIM(m.BudgetNum_2) == ALLTRIM(f_BudgetNum_pred)

          rec_num = rec_num - 1

          m.PacketNum_2 = IIF(ALLTRIM(m.BudgetNum_2) == ALLTRIM(f_BudgetNum_pred), rec_num, 0)

          m.TotalSum_2 = f_TotalSum_pred + SecurPackPRK.TotalSum

          m.TypeValueKind_2 = IIF(EMPTY(ALLTRIM(f_TypeValueKind_pred)) = .F., ALLTRIM(f_TypeValueKind_pred) + ', ' + ALLTRIM(SecurPackPRK.FilialCode), '(' + ALLTRIM(SecurPackPRK.FilialCode))

          m.IsOpen_2 = f_IsOpen_2_pred
          m.DefectName_2 = f_DefectName_2_pred

          IF BOF() = .F.

            SKIP -1

            IF BOF() = .T.
              SKIP +1
            ENDIF

          ENDIF

        CASE ALLTRIM(m.BudgetNum_2) <> ALLTRIM(f_BudgetNum_pred)

          m.PacketNum_2 = rec_num

          m.TotalSum_2 = SecurPackPRK.TotalSum

          m.TypeValueKind_2 = ' (' + ALLTRIM(SecurPackPRK.FilialCode)

          IF VARTYPE(SecurPackPRK.IdJk) = 'N' AND SecurPackPRK.IdJk = (-1)
            m.TypeValueKind_2 = 'ЯK отсутствует. ' + ALLTRIM(m.TypeValueKind_2)
          ENDIF

      ENDCASE

  ENDCASE

  SELECT SecurPackPRK_cursor

  IF BOF() = .F. AND EOF() = .F.
    GATHER MEMVAR FIELDS PacketNum_2, BudgetNum_2, TotalSum_2, TypeValueKind_2,BdgType_2,GrpShort_2, IsOpen_2, DefectName_2, PKO_2
  ENDIF

* BROWSE WINDOW brows

  SELECT SecurPackPRK
  rec_new = RECNO()

* WAIT WINDOW 'rec_new = ' + ALLTRIM(STR(rec_new)) + '   colvo_zap_polovina = ' + ALLTRIM(STR(colvo_zap_polovina)) TIMEOUT 2

  SELECT SecurPackPRK_cursor

  IF EOF() = .F.
    SKIP
  ENDIF

  IF EOF() = .T.
    EXIT
  ENDIF

  f_BudgetNum_pred = ALLTRIM(m.BudgetNum_2)

  f_TotalSum_pred = m.TotalSum_2

  f_TypeValueKind_pred = ALLTRIM(m.TypeValueKind_2)

  SELECT SecurPackPRK
ENDSCAN

SELECT SecurPackPRK_cursor
* BROWSE WINDOW brows

*DELETE FOR SecurPackPRK_cursor.TotalSum_1 = 0 AND SecurPackPRK_cursor.TotalSum_2 = 0
DELETE FOR EMPTY(SecurPackPRK_cursor.BudgetNum_1) AND EMPTY(SecurPackPRK_cursor.BudgetNum_2)

colvo_zap = RECCOUNT()

SELECT SecurPackPRK_cursor
GO TOP

* BROWSE WINDOW brows

SCAN

  SCATTER MEMVAR FIELDS TypeValueKind_1, BdgType_1, GrpShort_1, IsOpen_1, DefectName_1, PKO_1, TypeValueKind_2, BdgType_2, GrpShort_2, IsOpen_2, DefectName_2,PKO_2

  m.TypeValueKind_1 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_1)) = .F., ALLTRIM(m.TypeValueKind_1) + ')', SPACE(0))

  m.TypeValueKind_1 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_1)) = .F., ALLTRIM(m.TypeValueKind_1), '') + ;
    IIF(EMPTY(ALLTRIM(m.BdgType_1)) = .F., '; ' + ALLTRIM(m.BdgType_1), '') + ;
    IIF(EMPTY(ALLTRIM(m.GrpShort_1)) = .F., '; ' + ALLTRIM(m.GrpShort_1), '') + ;
    IIF(m.IsOpen_1 = 1, '; ' + 'Акт', IIF(EMPTY(ALLTRIM(m.DefectName_1)) = .F., '; ' + ALLTRIM(m.DefectName_1), '')) + ;
    IIF(m.PKO_1 = 1,'; ПКО','')

  m.TypeValueKind_2 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_2)) = .F., ALLTRIM(m.TypeValueKind_2) + ')', SPACE(0))

  m.TypeValueKind_2 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_2)) = .F., ALLTRIM(m.TypeValueKind_2), '') + ;
    IIF(EMPTY(ALLTRIM(m.BdgType_2)) = .F., '; ' + ALLTRIM(m.BdgType_2), '') + ;
    IIF(EMPTY(ALLTRIM(m.GrpShort_2)) = .F., '; ' + ALLTRIM(m.GrpShort_2), '') + ;
    IIF(m.IsOpen_2 = 1, '; ' + 'Акт', IIF(EMPTY(ALLTRIM(m.DefectName_2)) = .F., '; ' + ALLTRIM(m.DefectName_2), ''))+ ;
    IIF(m.PKO_2 = 1,'; ПКО','')
  GATHER MEMVAR FIELDS TypeValueKind_1, BdgType_1, GrpShort_1, IsOpen_1, DefectName_1, PKO_1, TypeValueKind_2, BdgType_2, GrpShort_2, IsOpen_2, DefectName_2, PKO_2

ENDSCAN

SELECT SecurPackPRK_cursor
GO TOP

SCAN FOR IdBank = 0

  SCATTER MEMVAR FIELDS GrName, IncFullName, CurDate, Sost, IdBank, CurTime, UserName, IdRouteHistory

  m.GrName = SecurPackPRK.GrName
  m.IncFullName = SecurPackPRK.IncFullName
  m.CurDate = SecurPackPRK.CurDate

  m.Sost = VAL(ALLTRIM(SecurPackPRK.Sost))

  m.IdBank = IIF(SecurPackPRK.IdBank = 0, '17', PADL(ALLTRIM(STR(SecurPackPRK.IdBank)), 2, '0'))
  m.CurTime = SecurPackPRK.CurTime

  m.UserName = SecurPackPRK.UserName
  m.IdRouteHistory = SecurPackPRK.IdRouteHistory

  GATHER MEMVAR FIELDS GrName, IncFullName, CurDate, Sost, IdBank, CurTime, UserName, IdRouteHistory

ENDSCAN

SELECT SecurPackPRK_cursor
GO TOP

* BROWSE WINDOW brows

* WAIT WINDOW 'colvo_zap = ' + ALLTRIM(STR(colvo_zap)) TIMEOUT 3

* -------------------------------------------------------------------- *
DO SecurPack_print_2Pin IN Priem_Vedom_PRK
* -------------------------------------------------------------------- *

RETURN


******************************************************************************* PROCEDURE SecurPack_print_2Pin ***********************************************************

PROCEDURE SecurPack_print_2Pin

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW 'f_curDate = ' + DTOC(f_curDate) TIMEOUT 3

IF pr_rabota_arxiv = .F.  && Признак работы в текущей смене pr_rabota_arxiv = .F., признак работы в архивной смене pr_rabota_arxiv = .T.

  Test_ShiftDate = TTOD(Proverka_DataSmena_PRK())  && Вычисляем дату открытой смены

*** new_data_odb - дата открытой смены, используется в печатных отчетах

* Test_ShiftDate = Test_ShiftDate + 1  && Для тестирования

* WAIT WINDOW 'До правки - new_data_odb = ' + DTOC(new_data_odb) + '   Test_ShiftDate = ' + DTOC(Test_ShiftDate) + '   skip_data_odb = ' + DTOC(skip_data_odb) + '   skip_data_odb_test = ' + DTOC(skip_data_odb_test) TIMEOUT 5

  IF new_data_odb <> Test_ShiftDate  && Проверяем равенство дат, если работа происходит в текущей смене и опердне

    skip_data_odb = skip_data_odb_test

    poisk_date = Test_ShiftDate
    new_data_odb = Test_ShiftDate
    prn_data_odb = Test_ShiftDate

    f_CurDate = Test_ShiftDate
    f_poisk_date = DTOC(Test_ShiftDate)

  ENDIF

* WAIT WINDOW 'После правки - new_data_odb = ' + DTOC(new_data_odb) + '   Test_ShiftDate = ' + DTOC(Test_ShiftDate) + '   skip_data_odb = ' + DTOC(skip_data_odb) + '   skip_data_odb_test = ' + DTOC(skip_data_odb_test) TIMEOUT 5

ENDIF

* WAIT WINDOW 'f_IdShift = ' + ALLTRIM(STR(f_IdShift)) TIMEOUT 3

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SELECT SecurPackPRK_cursor
GO TOP

* BROWSE WINDOW brows

IF colvo_zap <> 0

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  def_patch = SYS(2003)

  IF DIRECTORY(ALLTRIM(put_tmp_loc) + ALLTRIM(f_Login)) = .F.

    SET DEFAULT TO ALLTRIM(put_tmp_loc)
    MD ALLTRIM(f_Login)

  ENDIF

  QRCodeFileName = ALLTRIM(put_tmp_loc) + ALLTRIM(f_Login) + '\QRCodeFileReport' + ALLTRIM(SYS(2015)) + '.bmp'

  DECLARE GenerateBMP IN quricol32.dll STRING @, STRING @, INTEGER, INTEGER, INTEGER

* Itog_BudgetNum  && Подсчет количества сумок
* Itog_TotalSum  && Подсчет общей суммы

*    WAIT WINDOW 'SecurPackPRK_cursor_IdRouteHistory = ' + ALLTRIM(STR(SecurPackPRK_cursor.IdRouteHistory)) &&TIMEOUT 3
*    WAIT WINDOW 'f_IdRoute = ' + ALLTRIM(STR(f_IdRoute)) &&TIMEOUT 3
*    WAIT WINDOW 'Itog_TotalSum = ' + ALLTRIM(STR(Itog_TotalSum)) &&TIMEOUT 3
*    WAIT WINDOW 'Itog_BudgetNum = ' + ALLTRIM(STR(Itog_BudgetNum)) &&TIMEOUT 3

  SELECT SecurPackPRK_cursor
  GO TOP

*  BROWSE WINDOW brows

  QRCodeText_1 = ALLTRIM(STR(SecurPackPRK_cursor.IdRouteHistory)) + ';' + ALLTRIM(STR(f_IdRoute))
  QRCodeText_2 = ALLTRIM(STR(Itog_TotalSum, 15, 2)) + ';' + ALLTRIM(STR(Itog_BudgetNum)) + ';' + TTOC(SecurPackPRK_cursor.CurDate)
  QRCodeText = QRCodeText_1 + ';' + QRCodeText_2

*  QRCodeText_1 = ALLTRIM(STR(SecurPackPRK_cursor.IdRouteHistory)) + CHR(9) + ALLTRIM(STR(f_IdRoute))
*  QRCodeText_2 = ALLTRIM(STR(Itog_TotalSum, 15, 2)) + CHR(9) + ALLTRIM(STR(Itog_BudgetNum)) + CHR(9) + TTOC(SecurPackPRK_cursor.CurDate)
*  QRCodeText = QRCodeText_1 + CHR(9) + QRCodeText_2

  =GenerateBMP(QRCodeFileName, QRCodeText, 2, 5, 0)

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  SELECT SecurPackPRK_cursor
  GO TOP

* BROWSE WINDOW brows

  DO CASE
    CASE sergeizhuk_flag = .T.

      REPORT FORM pakpriem_group PREVIEW

    CASE sergeizhuk_flag = .F.

      DO CASE
        CASE Vedom_prn = .T.  && Признак выдачи отчета на принтер, когда идет первичная печать

          REPORT FORM pakpriem_group TO PRINTER NOCONSOLE
          REPORT FORM pakpriem_group TO PRINTER NOCONSOLE

        CASE Vedom_prn = .F.  && Признак выдачи отчета на экран, когда идет повторная печать

          REPORT FORM pakpriem_group PREVIEW

      ENDCASE

  ENDCASE

ENDIF

IF DIRECTORY(ALLTRIM(put_tmp_loc) + ALLTRIM(f_Login)) = .T.

  IF FILE(ALLTRIM(QRCodeFileName)) = .T.
    ERASE ALLTRIM(QRCodeFileName)
  ENDIF

  SET DEFAULT TO ALLTRIM(put_tmp_loc)
  RD ALLTRIM(f_Login)
  SET DEFAULT TO ALLTRIM(def_patch)

ENDIF

SELECT SecurPackPRK_cursor
USE

RETURN


***************************************************************** PROCEDURE Formir_Table_Prn **********************************************************

PROCEDURE Formir_Table_Prn
PARAMETERS Priznak_2PIN

IF TYPE('Priznak_2PIN') = 'U'
  Priznak_2PIN = 0
ENDIF

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

Itog_IdTypeValue = SecurPackPRK.IdTypeValue

DO CASE
  CASE Priznak_2PIN < 3

    SELECT DISTINCT BudgetNum ;
      FROM SecurPackPRK ;
      WHERE WorkType = '0' AND Flg2PIN = Priznak_2PIN ;
      INTO CURSOR Count_BudgetNum

    COUNT TO Itog_BudgetNum  && Подсчет количества сумок

    DO Start IN Propis_chislo WITH Itog_BudgetNum  && Формирование количества сумок прописью
    Itog_BudgetNum_Prop = Chislo_Prop


  CASE Priznak_2PIN = 3

    SELECT DISTINCT BudgetNum ;
      FROM SecurPackPRK ;
      WHERE WorkType = '1' AND Flg2PIN = 0 ;
      INTO CURSOR Count_BudgetNum

    COUNT TO Itog_BudgetNum_1  && Подсчет количества супер сумок

    SELECT DISTINCT BudgetNum ;
      FROM SecurPackPRK ;
      WHERE WorkType = '2' AND Flg2PIN = 0 ;
      INTO CURSOR Count_BudgetNum

    COUNT TO Itog_BudgetNum_2  && Подсчет количества вложений в супер сумки

    DO Start IN Propis_chislo WITH Itog_BudgetNum_2  && Формирование количества сумок прописью
    Itog_BudgetNum_Prop = Chislo_Prop

ENDCASE

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

DO CASE
  CASE Priznak_2PIN < 3

    SUM TotalSum TO Itog_TotalSum FOR WorkType = '0' AND Flg2PIN = Priznak_2PIN  && Подсчет общей суммы

  CASE Priznak_2PIN = 3

    SUM TotalSum TO Itog_TotalSum FOR (WorkType = '1' AND Flg2PIN = 0) OR (WorkType = '2' AND Pr_60322 = 1)  && Подсчет общей суммы

ENDCASE

SUM TotalSum TO f_Total60322 FOR WorkType = '2' AND Pr_60322=1 && Подсчет суммы по россыпи

DO Start IN PropisVal WITH Itog_TotalSum, Itog_IdTypeValue  && Формирование общей суммы прописью

Itog_TotalSum_Prop = Summa_Prop

CREATE CURSOR SecurPackPRK_cursor ;
  (GrName CHAR(80) NULL, IncFullName CHAR(40) NULL, CURDATE DATETIME, CURTIME CHAR(2) NULL, PacketNum_1 INTEGER,;
  BudgetNum_1 CHAR(20) NULL, TotalSum_1 Numeric(15, 2), TypeValueKind_1 CHAR(80) NULL,;
  PacketNum_2 INTEGER, BudgetNum_2 CHAR(20) NULL, TotalSum_2 Numeric(15, 2), TypeValueKind_2 CHAR(80) NULL,;
  Sost INTEGER, IdBank INTEGER, UserName CHAR(20) NULL, IdRouteHistory INTEGER,;
  IsOpen_1 INTEGER DEFAULT 0, IdPackDefect_1 INTEGER DEFAULT 0, PackDefectName_1 CHAR(30) DEFAULT SPACE(0), IsOpen_2 INTEGER DEFAULT 0, IdPackDefect_2 INTEGER DEFAULT 0, PackDefectName_2 CHAR(30) DEFAULT SPACE(0))

SELECT SecurPackPRK
colvo_zap = RECCOUNT()
GO TOP
* BROWSE WINDOW brows

colvo_zap_polovina = ROUND(colvo_zap / 2, 0)

* WAIT WINDOW 'colvo_zap = ' + ALLTRIM(STR(colvo_zap)) + '   colvo_zap_polovina = ' + ALLTRIM(STR(colvo_zap_polovina)) TIMEOUT 3

DO CASE
  CASE MOD(colvo_zap_polovina, 2) > 0

* WAIT WINDOW 'Добавляются строки по НЕ ЧЕТНОМУ варианту' TIMEOUT 2

    SELECT SecurPackPRK_cursor

    FOR I = 1 TO colvo_zap_polovina + 1
      APPEND BLANK
    ENDFOR

    colvo_zap_polovina = colvo_zap_polovina + 1


  CASE MOD(colvo_zap_polovina, 2) = 0

* WAIT WINDOW 'Добавляются строки по ЧЕТНОМУ варианту' TIMEOUT 2

    SELECT SecurPackPRK_cursor

    FOR I = 1 TO colvo_zap_polovina
      APPEND BLANK
    ENDFOR

    colvo_zap_polovina = colvo_zap_polovina + 0

ENDCASE

SELECT SecurPackPRK_cursor
* BROWSE WINDOW brows

f_IdRoute = SecurPackPRK.IdRoute

* ------------------------------------------------------------------------------------------------ *

SELECT SecurPackPRK_cursor
GO TOP
* BROWSE WINDOW brows

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------ *

f_BudgetNum_pred = ''

SCAN

  rec_num = RECNO()

  SELECT SecurPackPRK_cursor

  SCATTER MEMVAR

  m.GrName = SecurPackPRK.GrName
  m.IncFullName = SecurPackPRK.IncFullName
  m.CurDate = SecurPackPRK.CURDATE

  m.Sost = VAL(ALLTRIM(SecurPackPRK.Sost))

  m.IdBank = IIF(SecurPackPRK.IdBank = 0, '17', PADL(ALLTRIM(STR(SecurPackPRK.IdBank)), 2, '0'))

  m.CurTime = SecurPackPRK.CURTIME

  m.PacketNum_1 = IIF(ALLTRIM(SecurPackPRK.BudgetNum) <> ALLTRIM(f_BudgetNum_pred), rec_num, 0)

  m.BudgetNum_1 = SecurPackPRK.BudgetNum

  m.TotalSum_1 = SecurPackPRK.TotalSum

  m.TypeValueKind_1 = ALLTRIM(SecurPackPRK.TypeValueKind) + '  (' + ALLTRIM(SecurPackPRK.FilialCode) + ')'

  m.IsOpen_1 = IIF((m.IsOpen_1 = 1) OR (SecurPackPRK.IsOpen = 1), 1, 0)   && Единообразие

  m.IdPackDefect_1 = SecurPackPRK.DefectCode  && Добавил 08.09.2016 Ефимов А.В.

  m.PackDefectName_1 = ALLTRIM(SecurPackPRK.DefectName)

  IF VARTYPE(SecurPackPRK.IdJk) = 'N' AND SecurPackPRK.IdJk = (-1)
    m.TypeValueKind_1 = 'ЯK отсутствует. ' + ALLTRIM(m.TypeValueKind_1)
  ENDIF
  m.TypeValueKind_1 = m.TypeValueKind_1 + IIF(SecurPackPRK.Pr_60322==1,' (ПКО)','')

*    IF EMPTY(ALLTRIM(m.PackDefectName_1)) = .F.
*      m.TypeValueKind_1 = ALLTRIM(m.TypeValueKind_1) + CHR(13) + ALLTRIM(m.PackDefectName_1)
*    ENDIF

  m.UserName = SecurPackPRK.UserName

  m.IdRouteHistory = SecurPackPRK.IdRouteHistory

  GATHER MEMVAR

  SELECT SecurPackPRK
  rec_new = RECNO()

  IF rec_new = colvo_zap_polovina
    EXIT
  ENDIF

  SELECT SecurPackPRK_cursor
  SKIP

  IF EOF() = .T.
    EXIT
  ENDIF

  f_BudgetNum_pred = ALLTRIM(m.BudgetNum_1)

  SELECT SecurPackPRK
ENDSCAN

f_BudgetNum_pred = ''

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

* WAIT WINDOW 'rec_new = ' + ALLTRIM(STR(rec_new)) TIMEOUT 3

SELECT SecurPackPRK_cursor
GO TOP
* BROWSE WINDOW brows

SELECT SecurPackPRK
* BROWSE WINDOW brows

SCAN FOR RECNO() >= rec_new + 1

  rec_num = RECNO()

  SELECT SecurPackPRK_cursor

  SCATTER MEMVAR FIELDS PacketNum_2, BudgetNum_2, TotalSum_2, TypeValueKind_2, IsOpen_2

  m.PacketNum_2 = IIF(ALLTRIM(SecurPackPRK.BudgetNum) <> ALLTRIM(f_BudgetNum_pred), rec_num, 0)

  m.BudgetNum_2 = SecurPackPRK.BudgetNum

  m.TotalSum_2 = SecurPackPRK.TotalSum

  m.TypeValueKind_2 = ALLTRIM(SecurPackPRK.TypeValueKind) + '  (' + ALLTRIM(SecurPackPRK.FilialCode) + ')'

  m.IsOpen_2 = IIF((m.IsOpen_2 = 1) OR (SecurPackPRK.IsOpen = 1), 1, 0)  && Единообразие

  m.IdPackDefect_2 = SecurPackPRK.DefectCode  && Добавил 08.09.2016 Ефимов А.В.

  m.PackDefectName_2 = ALLTRIM(SecurPackPRK.DefectName)

  IF VARTYPE(SecurPackPRK.IdJk) = 'N' AND SecurPackPRK.IdJk = (-1)
    m.TypeValueKind_2 = 'ЯK отсутствует. ' + ALLTRIM(m.TypeValueKind_2)
  ENDIF
  m.TypeValueKind_2 = m.TypeValueKind_2 + IIF(SecurPackPRK.Pr_60322==1,' (ПКО)','')
*    IF EMPTY(ALLTRIM(m.PackDefectName_2)) = .F.
*      m.TypeValueKind_2 = ALLTRIM(m.TypeValueKind_2) + CHR(13) + ALLTRIM(m.PackDefectName_2)
*    ENDIF

  GATHER MEMVAR FIELDS PacketNum_2, BudgetNum_2, TotalSum_2, TypeValueKind_2, IsOpen_2, IdPackDefect_2

  SKIP

  IF EOF() = .T.
    EXIT
  ENDIF

  f_BudgetNum_pred = ALLTRIM(m.BudgetNum_2)

  SELECT SecurPackPRK
ENDSCAN

SELECT SecurPackPRK_cursor
colvo_zap = RECCOUNT()
GO TOP
* BROWSE WINDOW brows

* WAIT WINDOW 'colvo_zap = ' + ALLTRIM(STR(colvo_zap)) TIMEOUT 3

IF Priznak_2PIN < 3 &&AND NOT 'Документы Мегафон' $ GrName
  DELETE FOR SecurPackPRK_cursor.TotalSum_1 = 0 AND SecurPackPRK_cursor.TotalSum_2 = 0
ENDIF

colvo_zap = RECCOUNT()

SELECT SecurPackPRK_cursor
GO TOP

* BROWSE WINDOW brows

SCAN

  SCATTER MEMVAR FIELDS TypeValueKind_1, IsOpen_1, PackDefectName_1, TypeValueKind_2, IsOpen_2, PackDefectName_2

  m.TypeValueKind_1 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_1)) = .F., ALLTRIM(m.TypeValueKind_1) + ')', SPACE(0))

  m.TypeValueKind_1 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_1)) = .F., ALLTRIM(m.TypeValueKind_1), '') + ;
    IIF(m.IsOpen_1 = 1, '; ' + 'Акт', IIF(EMPTY(ALLTRIM(m.PackDefectName_1)) = .F., '; ' + ALLTRIM(m.PackDefectName_1), ''))

  m.TypeValueKind_2 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_2)) = .F., ALLTRIM(m.TypeValueKind_2) + ')', SPACE(0))

  m.TypeValueKind_2 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_2)) = .F., ALLTRIM(m.TypeValueKind_2), '') + ;
    IIF(m.IsOpen_2 = 1, '; ' + 'Акт', IIF(EMPTY(ALLTRIM(m.PackDefectName_2)) = .F., '; ' + ALLTRIM(m.PackDefectName_2), ''))

  GATHER MEMVAR FIELDS TypeValueKind_1, IsOpen_1, PackDefectName_1, TypeValueKind_2, IsOpen_2, PackDefectName_2

ENDSCAN

SELECT SecurPackPRK_cursor
GO TOP
* BROWSE WINDOW brows

SCAN FOR IdBank = 0

  SCATTER MEMVAR FIELDS GrName, IncFullName, CURDATE, Sost, IdBank, CURTIME, UserName, IdRouteHistory

  m.GrName = SecurPackPRK.GrName
  m.IncFullName = SecurPackPRK.IncFullName
  m.CurDate = SecurPackPRK.CURDATE

  m.Sost = VAL(ALLTRIM(SecurPackPRK.Sost))

  m.IdBank = IIF(SecurPackPRK.IdBank = 0, '17', PADL(ALLTRIM(STR(SecurPackPRK.IdBank)), 2, '0'))

  m.CurTime = SecurPackPRK.CURTIME

  m.UserName = SecurPackPRK.UserName
  m.IdRouteHistory = SecurPackPRK.IdRouteHistory

  GATHER MEMVAR FIELDS GrName, IncFullName, CURDATE, Sost, IdBank, CURTIME, UserName, IdRouteHistory

ENDSCAN

SELECT SecurPackPRK_cursor
GO TOP
* BROWSE WINDOW brows

* ------------------------------------------------------------ *
DO SecurPack_print IN Priem_Vedom_PRK
* ------------------------------------------------------------ *

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

RETURN


******************************************************************************* PROCEDURE SecurPack_print ***********************************************************

PROCEDURE SecurPack_print

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW 'f_curDate = ' + DTOC(f_curDate) TIMEOUT 3

IF pr_rabota_arxiv = .F.  && Признак работы в текущей смене pr_rabota_arxiv = .F., признак работы в архивной смене pr_rabota_arxiv = .T.

  Test_ShiftDate = TTOD(Proverka_DataSmena_PRK())  && Вычисляем дату открытой смены

*** new_data_odb - дата открытой смены, используется в печатных отчетах

* Test_ShiftDate = Test_ShiftDate + 1  && Для тестирования

* WAIT WINDOW 'До правки - new_data_odb = ' + DTOC(new_data_odb) + '   Test_ShiftDate = ' + DTOC(Test_ShiftDate) + '   skip_data_odb = ' + DTOC(skip_data_odb) + '   skip_data_odb_test = ' + DTOC(skip_data_odb_test) TIMEOUT 5

  IF new_data_odb <> Test_ShiftDate  && Проверяем равенство дат, если работа происходит в текущей смене и опердне

    skip_data_odb = skip_data_odb_test

    poisk_date = Test_ShiftDate
    new_data_odb = Test_ShiftDate
    prn_data_odb = Test_ShiftDate

    f_CurDate = Test_ShiftDate
    f_poisk_date = DTOC(Test_ShiftDate)

  ENDIF

* WAIT WINDOW 'После правки - new_data_odb = ' + DTOC(new_data_odb) + '   Test_ShiftDate = ' + DTOC(Test_ShiftDate) + '   skip_data_odb = ' + DTOC(skip_data_odb) + '   skip_data_odb_test = ' + DTOC(skip_data_odb_test) TIMEOUT 5

ENDIF

* WAIT WINDOW 'f_IdShift = ' + ALLTRIM(STR(f_IdShift)) TIMEOUT 3

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

def_patch = SYS(2003)

IF DIRECTORY(ALLTRIM(put_tmp_loc) + ALLTRIM(f_Login)) = .F.

  SET DEFAULT TO ALLTRIM(put_tmp_loc)
  MD ALLTRIM(f_Login)

ENDIF

QRCodeFileName = ALLTRIM(put_tmp_loc) + ALLTRIM(f_Login) + '\QRCodeFileReport' + ALLTRIM(SYS(2015)) + '.bmp'

DECLARE GenerateBMP IN quricol32.dll STRING @, STRING @, INTEGER, INTEGER, INTEGER

* Itog_BudgetNum  && Подсчет количества сумок
* Itog_TotalSum  && Подсчет общей суммы

SELECT SecurPackPRK_cursor
GO TOP

* BROWSE WINDOW brows

DO CASE
  CASE Priznak_2PIN < 3

    QRCodeText_1 = ALLTRIM(STR(SecurPackPRK_cursor.IdRouteHistory)) + ';' + ALLTRIM(STR(f_IdRoute))
    QRCodeText_2 = ALLTRIM(STR(Itog_TotalSum, 15, 2)) + ';' + ALLTRIM(STR(Itog_BudgetNum)) + ';' + TTOC(SecurPackPRK_cursor.CurDate)
    QRCodeText = QRCodeText_1 + ';' + QRCodeText_2

*  QRCodeText_2 = ALLTRIM(STR(Itog_TotalSum, 15, 2)) + CHR(9) + ALLTRIM(STR(Itog_BudgetNum)) + CHR(9) + TTOC(SecurPackPRK_cursor.CurDate)
*  QRCodeText_1 = ALLTRIM(STR(SecurPackPRK_cursor.IdRouteHistory)) + CHR(9) + ALLTRIM(STR(f_IdRoute))
*  QRCodeText = QRCodeText_1 + CHR(9) + QRCodeText_2


  CASE Priznak_2PIN = 3

    QRCodeText_1 = ALLTRIM(STR(SecurPackPRK_cursor.IdRouteHistory)) + ';' + ALLTRIM(STR(f_IdRoute))
    QRCodeText_2 = ALLTRIM(STR(Itog_TotalSum, 15, 2)) + ';' + ALLTRIM(STR(Itog_BudgetNum_2)) + ';' + TTOC(SecurPackPRK_cursor.CurDate)
    QRCodeText = QRCodeText_1 + ';' + QRCodeText_2

*  QRCodeText_2 = ALLTRIM(STR(Itog_TotalSum, 15, 2)) + CHR(9) + ALLTRIM(STR(Itog_BudgetNum_2)) + CHR(9) + TTOC(SecurPackPRK_cursor.CurDate)
*  QRCodeText_1 = ALLTRIM(STR(SecurPackPRK_cursor.IdRouteHistory)) + CHR(9) + ALLTRIM(STR(f_IdRoute))
*  QRCodeText = QRCodeText_1 + CHR(9) + QRCodeText_2


ENDCASE

=GenerateBMP(QRCodeFileName, QRCodeText, 2, 5, 0)

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SELECT SecurPackPRK_cursor
GO TOP

* BROWSE WINDOW brows

IF colvo_zap <> 0

  DO CASE
    CASE Priznak_2PIN < 3  && Печать простых сумок и секьюрпаков

      DO CASE
        CASE sergeizhuk_flag = .T.

          DO CASE
            CASE f_IsTransit = '0'
              REPORT FORM pakpriem_blank_rub_prk PREVIEW

            CASE f_IsTransit = '1'
              REPORT FORM pakpriem_blank_rub_trn_prk PREVIEW

          ENDCASE

        CASE sergeizhuk_flag = .F.

          DO CASE
            CASE Vedom_prn = .T.  && Признак выдачи отчета на принтер, когда идет первичная печать

              DO CASE
                CASE f_IsTransit = '0'

                  REPORT FORM pakpriem_blank_rub_prk TO PRINTER NOCONSOLE
                  REPORT FORM pakpriem_blank_rub_prk TO PRINTER NOCONSOLE

                CASE f_IsTransit = '1'

                  REPORT FORM pakpriem_blank_rub_trn_prk TO PRINTER NOCONSOLE
                  REPORT FORM pakpriem_blank_rub_trn_prk TO PRINTER NOCONSOLE

              ENDCASE


            CASE Vedom_prn = .F.  && Признак выдачи отчета на экран, когда идет повторная печать

              DO CASE
                CASE f_IsTransit = '0'

                  REPORT FORM pakpriem_blank_rub_prk PREVIEW

                CASE f_IsTransit = '1'

                  REPORT FORM pakpriem_blank_rub_trn_prk PREVIEW

              ENDCASE

          ENDCASE

      ENDCASE


    CASE Priznak_2PIN = 3  && Печать супер сумок с вложениями

      DO CASE
        CASE sergeizhuk_flag = .T.

          REPORT FORM supersumka_blank_rub_prk PREVIEW

        CASE sergeizhuk_flag = .F.

          DO CASE
            CASE Vedom_prn = .T.  && Признак выдачи отчета на принтер, когда идет первичная печать

              REPORT FORM supersumka_blank_rub_prk TO PRINTER NOCONSOLE
              REPORT FORM supersumka_blank_rub_prk TO PRINTER NOCONSOLE

            CASE Vedom_prn = .F.  && Признак выдачи отчета на экран, когда идет повторная печать

              REPORT FORM supersumka_blank_rub_prk PREVIEW

          ENDCASE

      ENDCASE

  ENDCASE

ENDIF

IF DIRECTORY(ALLTRIM(put_tmp_loc) + ALLTRIM(f_Login)) = .T.

  IF FILE(ALLTRIM(QRCodeFileName)) = .T.
    ERASE ALLTRIM(QRCodeFileName)
  ENDIF

  SET DEFAULT TO ALLTRIM(put_tmp_loc)
  RD ALLTRIM(f_Login)
  SET DEFAULT TO ALLTRIM(def_patch)

ENDIF

SELECT SecurPackPRK_cursor
USE

RETURN


***************************************************************** PROCEDURE Formir_Table_Prn_999 **********************************************************

PROCEDURE Formir_Table_Prn_999

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

Itog_IdTypeValue = SecurPackPRK.IdTypeValue

SELECT DISTINCT BudgetNum ;
  FROM SecurPackPRK ;
  WHERE WorkType = '0' ;
  INTO CURSOR Count_BudgetNum

COUNT TO Itog_BudgetNum  && Подсчет количества сумок

DO Start IN Propis_chislo WITH Itog_BudgetNum  && Формирование количества сумок прописью
Itog_BudgetNum_Prop = Chislo_Prop

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

SUM TotalSum TO Itog_TotalSum FOR WorkType = '0'  && Подсчет общей суммы

DO Start IN PropisVal WITH Itog_TotalSum, Itog_IdTypeValue  && Формирование общей суммы прописью
Itog_TotalSum_Prop = Summa_Prop

* Добавил в курсор поля  IdPackDefect_1 integer, IdPackDefect_2 integer  08.09.2016 Ефимов А.В.

CREATE CURSOR SecurPackPRK_cursor ;
  (GrName CHAR(80) NULL, IncFullName CHAR(40) NULL, CURDATE DATETIME, CURTIME CHAR(2) NULL, PacketNum_1 INTEGER,;
  BudgetNum_1 CHAR(20) NULL, TotalSum_1 Numeric(15, 2), TypeValueKind_1 CHAR(80) NULL,;
  PacketNum_2 INTEGER, BudgetNum_2 CHAR(20) NULL, TotalSum_2 Numeric(15, 2), TypeValueKind_2 CHAR(80) NULL,;
  Sost INTEGER, IdBank INTEGER, UserName CHAR(20) NULL, IdRouteHistory INTEGER,;
  IsOpen_1 INTEGER DEFAULT 0, IdPackDefect_1 INTEGER DEFAULT 0, PackDefectName_1 CHAR(30) DEFAULT SPACE(0), IsOpen_2 INTEGER DEFAULT 0, IdPackDefect_2 INTEGER DEFAULT 0, PackDefectName_2 CHAR(30) DEFAULT SPACE(0))

SELECT SecurPackPRK
colvo_zap = RECCOUNT()
GO TOP
* BROWSE WINDOW brows

colvo_zap_polovina = ROUND(colvo_zap / 2, 0)

* WAIT WINDOW 'colvo_zap = ' + ALLTRIM(STR(colvo_zap)) + '   colvo_zap_polovina = ' + ALLTRIM(STR(colvo_zap_polovina)) TIMEOUT 3

DO CASE
  CASE MOD(colvo_zap_polovina, 2) > 0

* WAIT WINDOW 'Добавляются строки по НЕ ЧЕТНОМУ варианту' TIMEOUT 2

    SELECT SecurPackPRK_cursor

    FOR I = 1 TO colvo_zap_polovina + 1
      APPEND BLANK
    ENDFOR

    colvo_zap_polovina = colvo_zap_polovina + 1


  CASE MOD(colvo_zap_polovina, 2) = 0

* WAIT WINDOW 'Добавляются строки по ЧЕТНОМУ варианту' TIMEOUT 2

    SELECT SecurPackPRK_cursor

    FOR I = 1 TO colvo_zap_polovina
      APPEND BLANK
    ENDFOR

    colvo_zap_polovina = colvo_zap_polovina + 0

ENDCASE

SELECT SecurPackPRK_cursor
* BROWSE WINDOW brows

f_IdRoute = SecurPackPRK.IdRoute

* ------------------------------------------------------------------------------------------------ *

SELECT SecurPackPRK_cursor
GO TOP
* BROWSE WINDOW brows

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------ *

f_BudgetNum_pred = ''

SCAN

  rec_num = RECNO()

  SELECT SecurPackPRK_cursor

  SCATTER MEMVAR

  m.GrName = SecurPackPRK.GrName
  m.IncFullName = SecurPackPRK.IncFullName
  m.CurDate = SecurPackPRK.CURDATE

  m.Sost = VAL(ALLTRIM(SecurPackPRK.Sost))

  m.IdBank = IIF(SecurPackPRK.IdBank = 0, '17', PADL(ALLTRIM(STR(SecurPackPRK.IdBank)), 2, '0'))

  m.CurTime = SecurPackPRK.CURTIME

  m.PacketNum_1 = IIF(ALLTRIM(SecurPackPRK.BudgetNum) <> ALLTRIM(f_BudgetNum_pred), rec_num, 0)

  m.BudgetNum_1 = SecurPackPRK.BudgetNum

  m.TotalSum_1 = SecurPackPRK.TotalSum

  m.TypeValueKind_1 = ALLTRIM(SecurPackPRK.TypeValueKind) + '  (' + ALLTRIM(SecurPackPRK.FilialCode) + ')'

  m.IsOpen_1 = IIF((m.IsOpen_1 = 1) OR (SecurPackPRK.IsOpen = 1), 1, 0)   && Единообразие

  m.IdPackDefect_1 = SecurPackPRK.DefectCode  && Добавил 08.09.2016 Ефимов А.В.

  m.PackDefectName_1 = ALLTRIM(SecurPackPRK.DefectName)

  IF VARTYPE(SecurPackPRK.IdJk) = 'N' AND SecurPackPRK.IdJk = (-1)
    m.TypeValueKind_1 = 'ЯK отсутствует. ' + ALLTRIM(m.TypeValueKind_1)
  ENDIF

*    IF EMPTY(ALLTRIM(m.PackDefectName_1)) = .F.
*      m.TypeValueKind_1 = ALLTRIM(m.TypeValueKind_1) + CHR(13) + ALLTRIM(m.PackDefectName_1)
*    ENDIF

  m.UserName = SecurPackPRK.UserName

  m.IdRouteHistory = SecurPackPRK.IdRouteHistory

  GATHER MEMVAR

  SELECT SecurPackPRK
  rec_new = RECNO()

  IF rec_new = colvo_zap_polovina
    EXIT
  ENDIF

  SELECT SecurPackPRK_cursor
  SKIP

  IF EOF() = .T.
    EXIT
  ENDIF

  f_BudgetNum_pred = ALLTRIM(m.BudgetNum_1)

  SELECT SecurPackPRK
ENDSCAN

f_BudgetNum_pred = ''

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

* WAIT WINDOW 'rec_new = ' + ALLTRIM(STR(rec_new)) TIMEOUT 3

SELECT SecurPackPRK_cursor
GO TOP
* BROWSE WINDOW brows

SELECT SecurPackPRK
* BROWSE WINDOW brows

SCAN FOR RECNO() >= rec_new + 1

  rec_num = RECNO()

  SELECT SecurPackPRK_cursor

  SCATTER MEMVAR FIELDS PacketNum_2, BudgetNum_2, TotalSum_2, TypeValueKind_2, IsOpen_2

  m.PacketNum_2 = IIF(ALLTRIM(SecurPackPRK.BudgetNum) <> ALLTRIM(f_BudgetNum_pred), rec_num, 0)

  m.BudgetNum_2 = SecurPackPRK.BudgetNum

  m.TotalSum_2 = SecurPackPRK.TotalSum

  m.TypeValueKind_2 = ALLTRIM(SecurPackPRK.TypeValueKind) + '  (' + ALLTRIM(SecurPackPRK.FilialCode) + ')'

  m.IsOpen_2 = IIF((m.IsOpen_2 = 1) OR (SecurPackPRK.IsOpen = 1), 1, 0)  && Единообразие

  m.IdPackDefect_2 = SecurPackPRK.DefectCode  && Добавил 08.09.2016 Ефимов А.В.

  m.PackDefectName_2 = ALLTRIM(SecurPackPRK.DefectName)

  IF VARTYPE(SecurPackPRK.IdJk) = 'N' AND SecurPackPRK.IdJk = (-1)
    m.TypeValueKind_2 = 'ЯK отсутствует. ' + ALLTRIM(m.TypeValueKind_2)
  ENDIF

*    IF EMPTY(ALLTRIM(m.PackDefectName_2)) = .F.
*      m.TypeValueKind_2 = ALLTRIM(m.TypeValueKind_2) + CHR(13) + ALLTRIM(m.PackDefectName_2)
*    ENDIF

  GATHER MEMVAR FIELDS PacketNum_2, BudgetNum_2, TotalSum_2, TypeValueKind_2, IsOpen_2, IdPackDefect_2

  SKIP

  IF EOF() = .T.
    EXIT
  ENDIF

  f_BudgetNum_pred = ALLTRIM(m.BudgetNum_2)

  SELECT SecurPackPRK
ENDSCAN

SELECT SecurPackPRK_cursor
colvo_zap = RECCOUNT()
GO TOP
* BROWSE WINDOW brows

* WAIT WINDOW 'colvo_zap = ' + ALLTRIM(STR(colvo_zap)) TIMEOUT 3

DELETE FOR SecurPackPRK_cursor.TotalSum_1 = 0 AND SecurPackPRK_cursor.TotalSum_2 = 0

colvo_zap = RECCOUNT()

SELECT SecurPackPRK_cursor
GO TOP

* BROWSE WINDOW brows

SCAN

  SCATTER MEMVAR FIELDS TypeValueKind_1, IsOpen_1, PackDefectName_1, TypeValueKind_2, IsOpen_2, PackDefectName_2

  m.TypeValueKind_1 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_1)) = .F., ALLTRIM(m.TypeValueKind_1) + ')', SPACE(0))

  m.TypeValueKind_1 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_1)) = .F., ALLTRIM(m.TypeValueKind_1), '') + ;
    IIF(m.IsOpen_1 = 1, '; ' + 'Акт', IIF(EMPTY(ALLTRIM(m.PackDefectName_1)) = .F., '; ' + ALLTRIM(m.PackDefectName_1), ''))

  m.TypeValueKind_2 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_2)) = .F., ALLTRIM(m.TypeValueKind_2) + ')', SPACE(0))

  m.TypeValueKind_2 = IIF(EMPTY(ALLTRIM(m.TypeValueKind_2)) = .F., ALLTRIM(m.TypeValueKind_2), '') + ;
    IIF(m.IsOpen_2 = 1, '; ' + 'Акт', IIF(EMPTY(ALLTRIM(m.PackDefectName_2)) = .F., '; ' + ALLTRIM(m.PackDefectName_2), ''))

  GATHER MEMVAR FIELDS TypeValueKind_1, IsOpen_1, PackDefectName_1, TypeValueKind_2, IsOpen_2, PackDefectName_2

ENDSCAN

SELECT SecurPackPRK_cursor
GO TOP
* BROWSE WINDOW brows

SCAN FOR IdBank = 0

  SCATTER MEMVAR FIELDS GrName, IncFullName, CURDATE, Sost, IdBank, CURTIME, UserName, IdRouteHistory

  m.GrName = SecurPackPRK.GrName
  m.IncFullName = SecurPackPRK.IncFullName
  m.CurDate = SecurPackPRK.CURDATE

  m.Sost = VAL(ALLTRIM(SecurPackPRK.Sost))

  m.IdBank = IIF(SecurPackPRK.IdBank = 0, '17', PADL(ALLTRIM(STR(SecurPackPRK.IdBank)), 2, '0'))

  m.CurTime = SecurPackPRK.CURTIME

  m.UserName = SecurPackPRK.UserName
  m.IdRouteHistory = SecurPackPRK.IdRouteHistory

  GATHER MEMVAR FIELDS GrName, IncFullName, CURDATE, Sost, IdBank, CURTIME, UserName, IdRouteHistory

ENDSCAN

SELECT SecurPackPRK_cursor
GO TOP
* BROWSE WINDOW brows

* ------------------------------------------------------------------- *
DO SecurPack_print_999 IN Priem_Vedom_PRK
* ------------------------------------------------------------------- *

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

RETURN


******************************************************************************* PROCEDURE SecurPack_print_999 ***********************************************************

PROCEDURE SecurPack_print_999

* WAIT WINDOW 'f_curDate = ' + DTOC(f_curDate) TIMEOUT 3

IF pr_rabota_arxiv = .F.  && Признак работы в текущей смене pr_rabota_arxiv = .F., признак работы в архивной смене pr_rabota_arxiv = .T.

  Test_ShiftDate = TTOD(Proverka_DataSmena_PRK())  && Вычисляем дату открытой смены

*** new_data_odb - дата открытой смены, используется в печатных отчетах

* Test_ShiftDate = Test_ShiftDate + 1  && Для тестирования

* WAIT WINDOW 'До правки - new_data_odb = ' + DTOC(new_data_odb) + '   Test_ShiftDate = ' + DTOC(Test_ShiftDate) + '   skip_data_odb = ' + DTOC(skip_data_odb) + '   skip_data_odb_test = ' + DTOC(skip_data_odb_test) TIMEOUT 5

  IF new_data_odb <> Test_ShiftDate  && Проверяем равенство дат, если работа происходит в текущей смене и опердне

    skip_data_odb = skip_data_odb_test

    poisk_date = Test_ShiftDate
    new_data_odb = Test_ShiftDate
    prn_data_odb = Test_ShiftDate

    f_CurDate = Test_ShiftDate
    f_poisk_date = DTOC(Test_ShiftDate)

  ENDIF

* WAIT WINDOW 'После правки - new_data_odb = ' + DTOC(new_data_odb) + '   Test_ShiftDate = ' + DTOC(Test_ShiftDate) + '   skip_data_odb = ' + DTOC(skip_data_odb) + '   skip_data_odb_test = ' + DTOC(skip_data_odb_test) TIMEOUT 5

ENDIF

* WAIT WINDOW 'f_IdShift = ' + ALLTRIM(STR(f_IdShift)) TIMEOUT 3

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

def_patch = SYS(2003)

IF DIRECTORY(ALLTRIM(put_tmp_loc) + ALLTRIM(f_Login)) = .F.

  SET DEFAULT TO ALLTRIM(put_tmp_loc)
  MD ALLTRIM(f_Login)

ENDIF

SELECT SecurPackPRK_cursor
GO TOP

* BROWSE WINDOW brows

QRCodeFileName = ALLTRIM(put_tmp_loc) + ALLTRIM(f_Login) + '\QRCodeFileReport' + ALLTRIM(SYS(2015)) + '.bmp'

DECLARE GenerateBMP IN quricol32.dll STRING @, STRING @, INTEGER, INTEGER, INTEGER

* Itog_BudgetNum  && Подсчет количества сумок
* Itog_TotalSum  && Подсчет общей суммы

QRCodeText_1 = ALLTRIM(STR(SecurPackPRK_cursor.IdRouteHistory)) + ';' + ALLTRIM(STR(f_IdRoute))
QRCodeText_2 = ALLTRIM(STR(Itog_TotalSum, 15, 2)) + ';' + ALLTRIM(STR(Itog_BudgetNum)) + ';' + TTOC(SecurPackPRK_cursor.CurDate)
QRCodeText = QRCodeText_1 + ';' + QRCodeText_2

*  QRCodeText_1 = ALLTRIM(STR(SecurPackPRK_cursor.IdRouteHistory)) + CHR(9) + ALLTRIM(STR(f_IdRoute))
*  QRCodeText_2 = ALLTRIM(STR(Itog_TotalSum, 15, 2)) + CHR(9) + ALLTRIM(STR(Itog_BudgetNum)) + CHR(9) + TTOC(SecurPackPRK_cursor.CurDate)
*  QRCodeText = QRCodeText_1 + CHR(9) + QRCodeText_2

=GenerateBMP(QRCodeFileName, QRCodeText, 2, 5, 0)

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SELECT SecurPackPRK_cursor
GO TOP

* BROWSE WINDOW brows

IF colvo_zap <> 0

  DO CASE
    CASE sergeizhuk_flag = .T.

      REPORT FORM pakpriem_blank_moneta PREVIEW

    CASE sergeizhuk_flag = .F.

      DO CASE
        CASE Vedom_prn = .T.  && Признак выдачи отчета на принтер, когда идет первичная печать

          REPORT FORM pakpriem_blank_moneta TO PRINTER NOCONSOLE
          REPORT FORM pakpriem_blank_moneta TO PRINTER NOCONSOLE

        CASE Vedom_prn = .F.  && Признак выдачи отчета на экран, когда идет повторная печать

          REPORT FORM pakpriem_blank_moneta PREVIEW

      ENDCASE

  ENDCASE

ENDIF

IF DIRECTORY(ALLTRIM(put_tmp_loc) + ALLTRIM(f_Login)) = .T.

  IF FILE(ALLTRIM(QRCodeFileName)) = .T.
    ERASE ALLTRIM(QRCodeFileName)
  ENDIF

  SET DEFAULT TO ALLTRIM(put_tmp_loc)
  RD ALLTRIM(f_Login)
  SET DEFAULT TO ALLTRIM(def_patch)

ENDIF

SELECT SecurPackPRK_cursor
USE

RETURN


***************************************************************** PROCEDURE Formir_Table_Prn_ADM **********************************************************

PROCEDURE Formir_Table_Prn_ADM
PARAMETERS Priznak_2PIN

IF TYPE('Priznak_2PIN') = 'U'
  Priznak_2PIN = 0
ENDIF

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

Itog_IdTypeValue = SecurPackPRK.IdTypeValue

*********************   Это печать россыпи из АДМ количество сумок не считается. Только выводится ПКО (приходник), а сумка скорее всего всегда одна.

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

CREATE CURSOR SecurPackPRK_cursor ;
  (GrName CHAR(80) NULL, IncFullName CHAR(40) NULL, IncFio CHAR(50), CURDATE DATETIME, CURTIME CHAR(2) NULL, PacketNum INTEGER, ;
  BudgetNum CHAR(20) NULL, TotalSum Numeric(15, 2), TypeValueKind CHAR(80) NULL,;
  Sost INTEGER, IdBank INTEGER, UserName CHAR(20) NULL, IdRouteHistory INTEGER, TypeACC CHAR(2) NULL, Account CHAR(20) NULL, FilialName CHAR(100) NULL, BankName CHAR(100) NULL, UIFactAddress CHAR(100),;
  IsOpen INTEGER DEFAULT 0, IdPackDefect INTEGER DEFAULT 0, IdTypeValue int, FiCode CHAR(20))

SELECT SecurPackPRK

colvo_zap = RECCOUNT()

GO TOP

* BROWSE WINDOW brows

* WAIT WINDOW 'colvo_zap = ' + ALLTRIM(STR(colvo_zap)) TIMEOUT 3

*  SELECT SecurPackPRK_cursor

* BROWSE WINDOW brows

f_IdRoute = SecurPackPRK.IdRoute

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------ *

SCAN

  rec_num = RECNO()

  SELECT SecurPackPRK_cursor

  SCATTER MEMVAR BLANK

  m.FiCode = SecurPackPRK.FilialCode
  m.GrName = SecurPackPRK.GrName

  m.IncFullName = SecurPackPRK.IncFullName

  m.UIFactAddress = SecurPackPRK.UIFactAddress

  m.IncFio = SecurPackPRK.Fio

  m.CurDate = SecurPackPRK.CURDATE

  m.Sost = VAL(ALLTRIM(SecurPackPRK.Sost))

  m.IdBank = SecurPackPRK.IdBank
  m.CurTime = SecurPackPRK.CURTIME

  m.PacketNum = rec_num

  m.BudgetNum = SecurPackPRK.BudgetNum

  m.TotalSum = SecurPackPRK.TotalSum

  m.TypeValueKind = ALLTRIM(SecurPackPRK.TypeValueKind) + '  (' + ALLTRIM(SecurPackPRK.FilialCode) + ')'

  m.UserName = SecurPackPRK.UserName

  m.IdRouteHistory = SecurPackPRK.IdRouteHistory

  m.TypeACC = ALLTRIM(SecurPackPRK.TypeACC)

  m.Account = ALLTRIM(SecurPackPRK.Account)

  m.FilialName = ALLTRIM(SecurPackPRK.FilialName)

  m.BankName = ALLTRIM(SecurPackPRK.BankName)

  m.TypeValueKind= ALLTRIM(SecurPackPRK.TypeValueKind)
  m.IdTypeValue= SecurPackPRK.IdTypeValue
  INSERT INTO SecurPackPRK_cursor FROM MEMVAR

  SELECT SecurPackPRK
ENDSCAN

SELECT SecurPackPRK_cursor
colvo_zap = RECCOUNT()
GO TOP
* BROWSE WINDOW brows

SCAN FOR IdBank = 0

  SCATTER MEMVAR FIELDS GrName, IncFullName, CURDATE, Sost, IdBank, CURTIME, UserName, IdRouteHistory

  m.GrName = SecurPackPRK.GrName
  m.IncFullName = SecurPackPRK.IncFullName
  m.CurDate = SecurPackPRK.CURDATE

  m.Sost = VAL(ALLTRIM(SecurPackPRK.Sost))

  m.IdBank = IIF(SecurPackPRK.IdBank = 0, '17', PADL(ALLTRIM(STR(SecurPackPRK.IdBank)), 2, '0'))

  m.CurTime = SecurPackPRK.CURTIME

  m.UserName = SecurPackPRK.UserName
  m.IdRouteHistory = SecurPackPRK.IdRouteHistory

  GATHER MEMVAR FIELDS GrName, IncFullName, CURDATE, Sost, IdBank, CURTIME, UserName, IdRouteHistory

ENDSCAN

SELECT SecurPackPRK_cursor
GO TOP
* BROWSE WINDOW brows

* ----------------------------------------------------------------------- *
DO SecurPack_print_ADM IN Priem_Vedom_PRK
* ----------------------------------------------------------------------- *

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

RETURN


******************************************************************************* PROCEDURE SecurPack_print_ADM ***********************************************************

PROCEDURE SecurPack_print_ADM

**  Тут печатаем приходные ордеры по деньгам вне сейфа

* WAIT WINDOW 'f_curDate = ' + DTOC(f_curDate) TIMEOUT 3

IF pr_rabota_arxiv = .F.  && Признак работы в текущей смене pr_rabota_arxiv = .F., признак работы в архивной смене pr_rabota_arxiv = .T.

  Test_ShiftDate = TTOD(Proverka_DataSmena_PRK())  && Вычисляем дату открытой смены

*** new_data_odb - дата открытой смены, используется в печатных отчетах

* Test_ShiftDate = Test_ShiftDate + 1  && Для тестирования

* WAIT WINDOW 'До правки - new_data_odb = ' + DTOC(new_data_odb) + '   Test_ShiftDate = ' + DTOC(Test_ShiftDate) + '   skip_data_odb = ' + DTOC(skip_data_odb) + '   skip_data_odb_test = ' + DTOC(skip_data_odb_test) TIMEOUT 5

  IF new_data_odb <> Test_ShiftDate  && Проверяем равенство дат, если работа происходит в текущей смене и опердне

    skip_data_odb = skip_data_odb_test

    poisk_date = Test_ShiftDate
    new_data_odb = Test_ShiftDate
    prn_data_odb = Test_ShiftDate

    f_CurDate = Test_ShiftDate
    f_poisk_date = DTOC(Test_ShiftDate)

  ENDIF

* WAIT WINDOW 'После правки - new_data_odb = ' + DTOC(new_data_odb) + '   Test_ShiftDate = ' + DTOC(Test_ShiftDate) + '   skip_data_odb = ' + DTOC(skip_data_odb) + '   skip_data_odb_test = ' + DTOC(skip_data_odb_test) TIMEOUT 5

ENDIF

* WAIT WINDOW 'f_IdShift = ' + ALLTRIM(STR(f_IdShift)) TIMEOUT 3

TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT IdTypeValue, POVAccount AS Acc
  FROM KassaAccounts
  WHERE MestoCode = '<<f_MestoCode>>'    -- ПРК всегда работает в послеоперационном времени.

ENDTEXT

* -------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* -------------------------------------------------------- *

IF RunSQL_Load('NonOperationAccounts', @ar) <> 1
  MESSAGEBOX('Таблица - NonOperationAccounts, необходимая для работы НЕ СФОРМИРОВАНА', 16, 'Операция ПЕЧАТИ ПКО Прекращена')
  DO exit_sql_server IN RunSQL_Load
  RETURN -1
ENDIF

* -------------------------------------------------------- *
DO exit_sql_server IN RunSQL_Load
* -------------------------------------------------------- *

SELECT SecurPackPRK_cursor
GO TOP
* BROWSE WINDOW brows
IF colvo_zap <> 0

  SCAN

    num_recno = RECNO()

    SELECT NonOperationAccounts

    LOCATE FOR NonOperationAccounts.IdTypeValue = SecurPackPRK_cursor.IdTypeValue

    f_FiCode=ALLTRIM(SecurPackPRK_cursor.FiCode)
    InkahranAcc = ALLTRIM(NonOperationAccounts.Acc)

    f_OrderNum = ''

    f_Acc20 = ALLTRIM(SecurPackPRK_cursor.Account)
    f_TypeAcc = ALLTRIM(SecurPackPRK_cursor.TypeAcc)
    f_TypeValueKind = ALLTRIM(SecurPackPRK_cursor.TypeValueKind)  && для вывода в ПКО      &&  Перенесено вверх на несколько строк, дабы f_Osnovanie было с ПИНом

    f_Osnovanie = 'Поступление наличных денег, выявленных вне кассеты из ПТС расположенного по адресу: (' ;
      + ALLTRIM(SecurPackPRK_cursor.UIFactAddress) + ')  от ' + DTOC(OpenDatePRK) + ', ПИН: ' + ALLTRIM(f_FiCode) + ' ' + ALLTRIM(f_TypeValueKind)

    f_BankName = ALLTRIM(SecurPackPRK_cursor.BankName)
    f_TotalSum = SecurPackPRK_cursor.TotalSum
    f_KassaKod = '33'
    f_KassaSumma = SecurPackPRK_cursor.TotalSum

    f_IncFio = ALLTRIM(SecurPackPRK_cursor.IncFio)  && для вывода в ПКО

    DO start IN PropisVal WITH f_TotalSum, Itog_IdTypeValue   && внутри иницилизируются одинаковым значением 2 переменные: summa_prop,propis_summa

    DO CASE
      CASE sergeizhuk_flag = .T.

        REPORT FORM prihod_kassa_order NEXT 1 PREVIEW

      CASE sergeizhuk_flag = .F.

        DO CASE
          CASE Vedom_prn = .T.  && Признак выдачи отчета на принтер, когда идет первичная печать

            REPORT FORM prihod_kassa_order NEXT 1 TO PRINTER NOCONSOLE

          CASE Vedom_prn = .F.  && Признак выдачи отчета на экран, когда идет повторная печать

            REPORT FORM prihod_kassa_order NEXT 1 PREVIEW

        ENDCASE

    ENDCASE

  ENDSCAN

ENDIF

SELECT SecurPackPRK_cursor
USE

RETURN


***************************************************************** PROCEDURE Formir_Table_Prn_Kraska **********************************************************

PROCEDURE Formir_Table_Prn_Kraska
PARAMETERS Priznak_2PIN

IF TYPE('Priznak_2PIN') = 'U'
  Priznak_2PIN = 0
ENDIF

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

Itog_IdTypeValue = SecurPackPRK.IdTypeValue

SUM TotalSum TO Itog_TotalSum

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

CREATE CURSOR SecurPackPRK_cursor ;
  (PacketNum  INTEGER, GrName CHAR(80) NULL, IncFullName CHAR(40) NULL, IncFio CHAR(50), FilialName CHAR(200) NULL,;
  BudgetNum CHAR(20) NULL, TypeValueKind CHAR(80) NULL, TotalSum Numeric(15, 2), TotalSumProp CHAR(200) NULL,;
  Sost INTEGER, IdBank INTEGER, UserName CHAR(20) NULL, BankName CHAR(100) NULL, ;
  IsOpen INTEGER DEFAULT 0, IdTypeValue INTEGER)

SELECT SecurPackPRK

colvo_zap = RECCOUNT()

GO TOP

* BROWSE WINDOW brows

* WAIT WINDOW 'colvo_zap = ' + ALLTRIM(STR(colvo_zap)) TIMEOUT 3

*  SELECT SecurPackPRK_cursor

* BROWSE WINDOW brows

f_IdRoute = SecurPackPRK.IdRoute

SELECT SecurPackPRK
GO TOP

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------ *

SCAN

  SELECT SecurPackPRK_cursor

  SCATTER MEMVAR BLANK

  m.GrName = SecurPackPRK.GrName

  m.IncFullName = SecurPackPRK.IncFullName

  m.IncFio = SecurPackPRK.Fio

  m.FilialName = ALLTRIM(SecurPackPRK.FilialName)

  m.BudgetNum = ALLTRIM(SecurPackPRK.BudgetNum)

  m.IdTypeValue = SecurPackPRK.IdTypeValue

  m.TypeValueKind = ALLTRIM(SecurPackPRK.TypeValueKind)  && + '  (' + ALLTRIM(SecurPackPRK.FilialCode) + ')'

  m.TotalSum = SecurPackPRK.TotalSum

  DO start IN PropisVal WITH m.TotalSum, m.IdTypeValue  && Формирование суммы прописью
  m.TotalSumProp = Summa_Prop

  m.UserName = SecurPackPRK.UserName

  m.Sost = VAL(ALLTRIM(SecurPackPRK.Sost))

  m.IdBank = SecurPackPRK.IdBank

  m.BankName = ALLTRIM(SecurPackPRK.BankName)

  INSERT INTO SecurPackPRK_cursor FROM MEMVAR

  SELECT SecurPackPRK
ENDSCAN

*  DO start IN Propis_chislo WITH Itog_BuhVedom  && Формирование количества ведомостей прописью
*  Itog_BuhVedom_Prop = Chislo_Prop

SELECT SecurPackPRK_cursor

colvo_zap = RECCOUNT()

GO TOP

* BROWSE WINDOW brows

SCAN FOR IdBank = 0

  SCATTER MEMVAR FIELDS GrName, IncFullName, Sost, IdBank, UserName

  m.GrName = SecurPackPRK.GrName
  m.IncFullName = SecurPackPRK.IncFullName

  m.Sost = VAL(ALLTRIM(SecurPackPRK.Sost))

  m.IdBank = IIF(SecurPackPRK.IdBank = 0, '17', PADL(ALLTRIM(STR(SecurPackPRK.IdBank)), 2, '0'))

  m.UserName = SecurPackPRK.UserName

  GATHER MEMVAR FIELDS GrName, IncFullName, Sost, IdBank, UserName

ENDSCAN

SELECT SecurPackPRK_cursor
GO TOP

* BROWSE WINDOW brows

* ----------------------------------------------------------------------- *
DO SecurPack_Print_Kraska IN Priem_Vedom_PRK
* ----------------------------------------------------------------------- *

SELECT SecurPackPRK
GO TOP

* BROWSE WINDOW brows

RETURN


******************************************************************************* PROCEDURE SecurPack_Print_Kraska ***********************************************************

PROCEDURE SecurPack_Print_Kraska

* WAIT WINDOW 'f_curDate = ' + DTOC(f_curDate) TIMEOUT 3

IF pr_rabota_arxiv = .F.  && Признак работы в текущей смене pr_rabota_arxiv = .F., признак работы в архивной смене pr_rabota_arxiv = .T.

  Test_ShiftDate = TTOD(Proverka_DataSmena_PRK())  && Вычисляем дату открытой смены

*** new_data_odb - дата открытой смены, используется в печатных отчетах

* Test_ShiftDate = Test_ShiftDate + 1  && Для тестирования

* WAIT WINDOW 'До правки - new_data_odb = ' + DTOC(new_data_odb) + '   Test_ShiftDate = ' + DTOC(Test_ShiftDate) + '   skip_data_odb = ' + DTOC(skip_data_odb) + '   skip_data_odb_test = ' + DTOC(skip_data_odb_test) TIMEOUT 5

  IF new_data_odb <> Test_ShiftDate  && Проверяем равенство дат, если работа происходит в текущей смене и опердне

    skip_data_odb = skip_data_odb_test

    poisk_date = Test_ShiftDate
    new_data_odb = Test_ShiftDate
    prn_data_odb = Test_ShiftDate

    f_CurDate = Test_ShiftDate
    f_poisk_date = DTOC(Test_ShiftDate)

  ENDIF

* WAIT WINDOW 'После правки - new_data_odb = ' + DTOC(new_data_odb) + '   Test_ShiftDate = ' + DTOC(Test_ShiftDate) + '   skip_data_odb = ' + DTOC(skip_data_odb) + '   skip_data_odb_test = ' + DTOC(skip_data_odb_test) TIMEOUT 5

ENDIF

* WAIT WINDOW 'f_IdShift = ' + ALLTRIM(STR(f_IdShift)) TIMEOUT 3

SELECT SecurPackPRK_cursor

REPLACE ALL PacketNum WITH RECNO()

GO TOP

* BROWSE WINDOW brows

Num_IdTranSheet = f_IdRouteHistory

IF colvo_zap <> 0

  DO CASE
    CASE sergeizhuk_flag = .T.

      REPORT FORM Akt_prk_kraska PREVIEW

    CASE sergeizhuk_flag = .F.

      DO CASE
        CASE Vedom_prn = .T.  && Признак выдачи отчета на принтер, когда идет первичная печать

          REPORT FORM Akt_prk_kraska TO PRINTER NOCONSOLE

        CASE Vedom_prn = .F.  && Признак выдачи отчета на экран, когда идет повторная печать

          REPORT FORM Akt_prk_kraska PREVIEW

      ENDCASE

  ENDCASE

ENDIF

SELECT SecurPackPRK_cursor
USE

RETURN


************************************************************************************************* FUNCTION CloseDB **********************************************************************

FUNCTION CloseDB

IF USED('SecurPack') = .T.
  SELECT SecurPack
  USE
ENDIF

ENDFUNC


******************************************************************************************************************************************************************************************




