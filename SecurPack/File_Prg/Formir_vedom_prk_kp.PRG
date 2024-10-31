********************************************************************************************************************************
*** Формирование ведомости принятых пакетов в результате выполнения пункта меню "Передача обработанных сумок из ПРК в КП" ***
********************************************************************************************************************************

*********************************************************************** PROCEDURE Start_ClientGroupMax *********************************************************

PROCEDURE Start_ClientGroupMax
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)
HIDE POPUP vidacha_sumka_prk

Vedom_prn = .F.  && Признак выдачи отчета на принтер, когда идет повторная печать

* ---------------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

DIMENSION ar[1,2]

f_IncFullName = ''
perinp1 = ''
msg = ''

*  IF sergeizhuk_flag = .T.
*    f_IdShift_1 = 7305
*  ENDIF

* WAIT WINDOW 'f_IdShift = ' + ALLTRIM(STR(f_IdShift)) TIMEOUT 3

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT trn.Id, clm.IdClnInk, clm.ClientMaxGroup, COUNT(bgh.BudgetNum) AS BudgetsAmount, trn.ChangeDatePRK, trn.State, trn.IdGrup, CAST(0 AS bit) AS Fl
  FROM TransferSheet AS trn (READPAST)
  INNER JOIN BudgetHistoryPRK AS bgh (READPAST) ON bgh.IdTranSheet = trn.Id
  INNER JOIN View_ClientGroupMax_Filial_Bank AS clm (READPAST) ON clm.IdFilial = bgh.IdFilial
  WHERE trn.IdShift = ?f_IdShift_1
  GROUP BY trn.Id, clm.IdClnInk, clm.ClientMaxGroup, trn.ChangeDatePRK, trn.State, trn.IdGrup
  ORDER BY clm.IdClnInk ASC

ENDTEXT

IF RunSQL_Load('TransferSheet', @ar) <> 1
  DO exit_sql_server IN RunSQL_Load
  RETURN -1
ENDIF

* BROWSE WINDOW brows

IF USED('TransferSheet') = .T.

  SELECT TransferSheet

  colvo_zap = RECCOUNT()

  IF colvo_zap <> 0

* --------------------------------------------------- *
    DO FORM Export_PRK_KP_VIP
* --------------------------------------------------- *

    IF NOT LASTKEY() = 27 AND Pr_Exit_Form = .F.

*** Собираем выбранные передаточные ведомости по переданным сумкам в КП

      SELECT TransferSheet

* BROWSE WINDOW brows

      f_IdTransferSheet = ''

      SCAN FOR TransferSheet.Fl = .T.
        f_IdTransferSheet = f_IdTransferSheet + ALLTRIM(STR(TransferSheet.Id)) + [,]
      ENDSCAN

      f_IdTransferSheet = f_IdTransferSheet + ['']  && Полный список выбранных оператором сумок

*      WAIT WINDOW 'f_IdTransferSheet = ' + ALLTRIM(f_IdTransferSheet) TIMEOUT 3

* ----------------------------------------------------------- Находим код сумки или пакета с пересчетом в таблице SignProc ---------------------------------------------------------------------- *

      TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

        SELECT Id
        FROM SignProc (READPAST)
        WHERE Idplace = 2

      ENDTEXT

      IF RunSQL_Load('SignProc', @ar) <> 1
        DO exit_sql_server IN RunSQL_Load
        RETURN -1
      ENDIF

* BROWSE WINDOW brows

      IF USED('SignProc') = .T.

        SELECT SignProc

        colvo_zap = RECCOUNT()

        IF colvo_zap <> 1
          DO exit_sql_server IN RunSQL_Load
          =err('Внимание! Не верно заполнен справочник Signproc.SignName для рублевой ПРК')
          RETURN -1
        ENDIF

        f_IdSignProc = SignProc.Id  && f_IdSignProc = 15  && Код сумки или пакета с пересчетом для Москвы - f_IdSignProc = 1  && Код сумки или пакета с пересчетом для Питера

        USE

      ENDIF

* ------------------------------------------------------------------------- *
      DO start_vedom_vip IN Formir_Vedom_PRK_KP
* ------------------------------------------------------------------------- *

    ELSE

      DO exit_sql_server IN RunSQL_Load
      RETURN

    ENDIF

  ELSE
    =soob('Внимание! В сформированной таблице из базы SQL Server TransferSheet записей не обнаружено ... ')
  ENDIF

ELSE
  =soob('Внимание! Выборка данных из базы SQL Server таблица TransferSheet закончилась аварийно ... ')
ENDIF

* ---------------------------------------------------------------- *
DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

SELECT(sel)
RETURN


*********************************************************************** PROCEDURE Start_RouteHistory ************************************************************

PROCEDURE Start_RouteHistory
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)
HIDE POPUP vidacha_sumka_prk

Vedom_prn = .F.  && Признак выдачи отчета на принтер, когда идет повторная печать

* ---------------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

DIMENSION ar[1,2]

f_IncFullName = ''
perinp1 = ''
msg = ''

*  IF sergeizhuk_flag = .T.
*    f_IdShift_1 = 7305
*  ENDIF

* WAIT WINDOW 'f_IdShift = ' + ALLTRIM(STR(f_IdShift)) TIMEOUT 3

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT rth.Id, rth.IdShift, rth.IdRoute, rth.IdCashMessenger, CONVERT(CHAR(8), rth.DtEnd, 114) + ' - ' + chm.Fio AS Fio, rth.IdUser, usr.UserName, COUNT(bgh.BudgetNum) AS BudgetsAmount, rth.DtEnd, CAST(0 AS bit) AS Fl
  FROM RouteHistory AS rth (READPAST)
  INNER JOIN BudgetHistoryPRK AS bgh (READPAST) ON bgh.IdRouteHistory = rth.Id
  INNER JOIN CashMessenger AS chm (READPAST) ON chm.Id = rth.IdCashMessenger
  INNER JOIN [User] AS usr (READPAST) ON usr.Id = rth.IdUser
  INNER JOIN TransferSheet AS trs (READPAST) ON trs.Id = bgh.IdTranSheet
  WHERE rth.IdShift = ?f_IdShift_1
  GROUP BY rth.Id, rth.IdShift, rth.IdRoute, rth.IdCashMessenger, chm.Fio, rth.IdUser, usr.UserName, rth.BudgetsAmount, rth.DtEnd
  ORDER BY rth.IdRoute ASC, BudgetsAmount DESC

ENDTEXT

IF RunSQL_Load('RouteHistory', @ar) <> 1
  DO exit_sql_server IN RunSQL_Load
  RETURN -1
ENDIF

* BROWSE WINDOW brows

IF USED('RouteHistory') = .T.

  SELECT RouteHistory

  colvo_zap = RECCOUNT()

  IF colvo_zap <> 0

* --------------------------------------------- *
    DO FORM Export_PRK_KP
* --------------------------------------------- *

    IF NOT LASTKEY() = 27 AND Pr_Exit_Form = .F.

*** Собираем выбранные передаточные ведомости по переданным сумкам в КП

      SELECT RouteHistory

      f_RecordRouteHistory = ''

      SCAN FOR RouteHistory.fl = .T.
        f_RecordRouteHistory = f_RecordRouteHistory + ALLTRIM(STR(RouteHistory.Id)) + [,]
      ENDSCAN

      f_RecordRouteHistory = f_RecordRouteHistory + ['']  && Полный список выбранных оператором сумок

* WAIT WINDOW 'f_RecordRouteHistory = ' + ALLTRIM(f_RecordRouteHistory) TIMEOUT 3

      TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

        SELECT DISTINCT rth.Id, rth.IdShift, rth.IdRoute, rth.IdCashMessenger, chm.Fio, rth.IdUser, usr.UserName, rth.BudgetsAmount, rth.DtEnd, bgh.IdTranSheet
        FROM RouteHistory AS rth (READPAST)
        INNER JOIN BudgetHistoryPRK AS bgh (READPAST) ON bgh.IdRouteHistory = rth.Id
        INNER JOIN CashMessenger AS chm (READPAST) ON chm.Id = rth.IdCashMessenger
        INNER JOIN [User] AS usr (READPAST) ON usr.Id = rth.IdUser
        INNER JOIN TransferSheet AS trs (READPAST) ON trs.Id = bgh.IdTranSheet
        WHERE rth.IdShift = ?f_IdShift_1 AND rth.Id IN (<<f_RecordRouteHistory>>)
        ORDER BY rth.Id

      ENDTEXT

      IF RunSQL_Load('RouteHistory', @ar) <> 1
        DO exit_sql_server IN RunSQL_Load
        RETURN -1
      ENDIF

* BROWSE WINDOW brows

      IF USED('RouteHistory') = .T.

        SELECT RouteHistory

        colvo_zap = RECCOUNT()

        IF colvo_zap <> 0

          f_IdTransferSheet = ''

          SELECT RouteHistory

          SCAN
            f_IdTransferSheet = f_IdTransferSheet + ALLTRIM(STR(RouteHistory.IdTranSheet)) + [,]
          ENDSCAN

          f_IdTransferSheet = f_IdTransferSheet + ['']  && Полный список выбранных оператором сумок

*          WAIT WINDOW 'f_IdTransferSheet = ' + ALLTRIM(f_IdTransferSheet) TIMEOUT 3

* ----------------------------------------------------------- Находим код сумки или пакета с пересчетом в таблице SignProc ---------------------------------------------------------------------- *

          TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT Id
            FROM SignProc (READPAST)
            WHERE Idplace = 2

          ENDTEXT

          IF RunSQL_Load('SignProc', @ar) <> 1
            DO exit_sql_server IN RunSQL_Load
            RETURN -1
          ENDIF

* BROWSE WINDOW brows

          IF USED('SignProc') = .T.

            SELECT SignProc

            colvo_zap = RECCOUNT()

            IF colvo_zap <> 1
              DO exit_sql_server IN RunSQL_Load
              =err('Внимание! Не верно заполнен справочник Signproc.SignName для рублевой ПРК')
              RETURN -1
            ENDIF

            f_IdSignProc = SignProc.Id  && f_IdSignProc = 15  && Код сумки или пакета с пересчетом для Москвы - f_IdSignProc = 1  && Код сумки или пакета с пересчетом для Питера

            USE

          ENDIF

* ----------------------------------------------------------------------------- *
          DO start_vedom_prk IN Formir_Vedom_PRK_KP
* ----------------------------------------------------------------------------- *

        ELSE
          =soob('Внимание! В сформированной таблице из базы SQL Server RouteHistory записей не обнаружено ... ')
        ENDIF

      ELSE
        =soob('Внимание! Выборка данных из базы SQL Server таблица RouteHistory закончилась аварийно ... ')
      ENDIF

    ELSE

      DO exit_sql_server IN RunSQL_Load
      RETURN

    ENDIF

  ELSE
    =soob('Внимание! В сформированной таблице из базы SQL Server RouteHistory записей не обнаружено ... ')
  ENDIF

ELSE
  =soob('Внимание! Выборка данных из базы SQL Server таблица RouteHistory закончилась аварийно ... ')
ENDIF

* ---------------------------------------------------------------- *
DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

SELECT(sel)
RETURN


******************************************************************* PROCEDURE start_vedom_prk *******************************************************************

PROCEDURE start_vedom_prk

tim1 = SECONDS()

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Формирования печатного отчета об переданных сумках из ПРК в КП ...',WCOLS())

* WAIT WINDOW 'f_IdShift = ' + ALLTRIM(STR(f_IdShift)) TIMEOUT 3
* WAIT WINDOW 'f_RecordRouteHistory = ' + ALLTRIM(f_RecordRouteHistory) TIMEOUT 3
* WAIT WINDOW 'f_IdTransferSheet = ' + ALLTRIM(STR(f_IdTransferSheet)) TIMEOUT 3

* ---------------------------------------------------------------------------- Обрабатываем Секьюрпаки с ПЕРЕСЧЁТОМ рублевые ------------------------------------------------------------------------------------ *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT bgh.*, trs.*, rth.*, f.FilialCode, f.FilialName, 1 AS Sost, usr.UserName
  FROM BudgetHistoryPRK AS bgh (READPAST)
  INNER JOIN TransferSheet AS trs (READPAST) ON trs.Id = bgh.IdTranSheet
  INNER JOIN RouteHistory AS rth (READPAST) ON rth.Id = bgh.IdRouteHistory
  INNER JOIN Filial AS f (READPAST) ON bgh.IdFilial = f.Id
  INNER JOIN [User] AS usr (READPAST) ON usr.Id = bgh.IdUser
  WHERE rth.IdShift = ?f_IdShift_1 AND bgh.WorkType = '0' AND bgh.Flg2PIN = 0 AND bgh.SelDat = 1 AND rth.Id IN (<<f_RecordRouteHistory>>) AND bgh.IdTranSheet IN (<<f_IdTransferSheet>>)

ENDTEXT

=RunSQL_Load('Priem_PRK_KP', @ar)

SELECT Priem_PRK_KP

colvo_zap = RECCOUNT()

* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки с ПЕРЕСЧЁТОМ рублевые'

IF colvo_zap <> 0
  SELECT IdTranSheet ;
      FROM Priem_PRK_KP ;
      GROUP BY IdTranSheet ;
      INTO CURSOR _TranSheets
      
  SELECT _TranSheets
  
  SCAN ALL

      f_IncFullName = ''

      SELECT f_IncFullName AS IncFullName, sp.* ;
        FROM Priem_PRK_KP AS sp ;
        WHERE sp.IdSignProc = ?f_IdSignProc  AND sp.IdTranSheet=_TranSheets.IdTranSheet ;
        INTO CURSOR SecurPackPRK READWRITE ;
        ORDER BY SuperBudget, WorkType, BudgetNum

      colvo_zap_prn = _TALLY

      Text_grName = 'Секьюрпаки стандартные принятые из ПРК в КП'

      SELECT SecurPackPRK

      REPLACE ALL CurDate WITH new_data_odb

    * BROWSE WINDOW brows

    * ---------------------------------------------------------------------------------------- *

      IF colvo_zap_prn <> 0
        DO Formir_Table_Prn IN Formir_Vedom_PRK_KP WITH 0                                                                         && Печать однопиновых ?
      ENDIF
  ENDSCAN 
  USE IN _TranSheets
* ---------------------------------------------------------------------------------------- *

ENDIF

* -------------------------------------------------------------------- Обрабатываем МНОГОПИНОВЫЕ Секьюрпаки ПЕРЕСЧЁТОМ рублевые ----------------------------------------------------------------------- *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT bgh.*, trs.*, rth.*, f.FilialCode, f.FilialName, 1 AS Sost, usr.UserName
  FROM BudgetHistoryPRK AS bgh (READPAST)
  INNER JOIN TransferSheet AS trs (READPAST) ON trs.Id = bgh.IdTranSheet
  INNER JOIN RouteHistory AS rth (READPAST) ON rth.Id = bgh.IdRouteHistory
  INNER JOIN Filial AS f (READPAST) ON bgh.IdFilial = f.Id
  INNER JOIN [User] AS usr (READPAST) ON usr.Id = bgh.IdUser
  WHERE rth.IdShift = ?f_IdShift_1 AND bgh.WorkType = '0' AND bgh.Flg2PIN > 0 AND bgh.SelDat = 1 AND rth.Id IN (<<f_RecordRouteHistory>>) AND bgh.IdTranSheet IN (<<f_IdTransferSheet>>)
  ORDER BY bgh.BudgetNum, bgh.TypeValueKind

ENDTEXT

=RunSQL_Load('Priem_PRK_KP', @ar)

SELECT Priem_PRK_KP

colvo_zap = RECCOUNT()

* BROWSE WINDOW brows TITLE 'Обрабатываем МНОГОПИНОВЫЕ Секьюрпаки ПЕРЕСЧЁТОМ рублевые'

IF colvo_zap <> 0

  f_IncFullName = ''
  
  SELECT IdTranSheet ;
      FROM Priem_PRK_KP ;
      GROUP BY IdTranSheet ;
      INTO CURSOR _TranSheets
      
  SELECT _TranSheets
  
  SCAN ALL

      SELECT BudgetNum AS BudgetNum2Pin, f_IncFullName AS IncFullName, sp.* ;
        FROM Priem_PRK_KP AS sp ;
        WHERE sp.IdSignProc = ?f_IdSignProc  AND sp.IdTranSheet=_TranSheets.IdTranSheet ;
        INTO CURSOR SecurPackPRK READWRITE ;
        ORDER BY sp.BudgetNum, sp.TypeValueKind

      colvo_zap_prn = _TALLY

    * BROWSE WINDOW brows TITLE 'INTO CURSOR SecurPackPRK READWRITE'

      Text_grName = 'Ведомость сдачи инкассаторских сумок (мешков)'

      SELECT SecurPackPRK

      REPLACE ALL CurDate WITH new_data_odb

      REPLACE BudgetNum2Pin WITH STRTRAN(BudgetNum2Pin, '/2', SPACE(0)) FOR '/2' $ BudgetNum2Pin

      REPLACE BudgetNum2Pin WITH STRTRAN(BudgetNum2Pin, '/3', SPACE(0)) FOR '/3' $ BudgetNum2Pin

    * BROWSE WINDOW brows TITLE 'REPLACE BudgetNum2Pin'

      IF colvo_zap_prn <> 0
        DO Formir_table_prn IN Formir_vedom_prk_kp WITH 1                                                            && Печать многопиновых
      ENDIF
   ENDSCAN 
   USE IN _TranSheets

ENDIF

* ---------------------------------------------------------------------------- Обрабатываем Супер сумки с ПЕРЕСЧЁТОМ рублевые ------------------------------------------------------------------------------------ *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT bgh.*, trs.*, rth.*, f.FilialCode, f.FilialName, 1 AS Sost, usr.UserName
  FROM BudgetHistoryPRK AS bgh (READPAST)
  INNER JOIN TransferSheet AS trs (READPAST) ON trs.Id = bgh.IdTranSheet
  INNER JOIN RouteHistory AS rth (READPAST) ON rth.Id = bgh.IdRouteHistory
  INNER JOIN Filial AS f (READPAST) ON bgh.IdFilial = f.Id
  INNER JOIN [User] AS usr (READPAST) ON usr.Id = bgh.IdUser
  WHERE rth.IdShift = ?f_IdShift_1 AND bgh.WorkType IN ('1', '2') AND bgh.Flg2PIN = 0 AND bgh.SelDat = 1 AND rth.Id IN (<<f_RecordRouteHistory>>) AND bgh.IdTranSheet IN (<<f_IdTransferSheet>>)

ENDTEXT

=RunSQL_Load('Priem_PRK_KP', @ar)

SELECT Priem_PRK_KP

colvo_zap = RECCOUNT()

* BROWSE WINDOW brows TITLE 'Обрабатываем Супер сумки с ПЕРЕСЧЁТОМ рублевые'

IF colvo_zap <> 0
  SELECT IdTranSheet ;
      FROM Priem_PRK_KP ;
      GROUP BY IdTranSheet ;
      INTO CURSOR _TranSheets
      
  SELECT _TranSheets
  
  SCAN ALL
      f_IncFullName = ''

      SELECT SPACE(20) AS Budget2Pin, f_IncFullName AS IncFullName, sp.* ;
        FROM Priem_PRK_KP AS sp ;
        WHERE sp.IdSignProc = ?f_IdSignProc AND sp.IdTranSheet=_TranSheets.IdTranSheet ;
        INTO CURSOR SecurPackPRK READWRITE ;
        ORDER BY SuperBudget, WorkType, BudgetNum

      colvo_zap_prn = _TALLY

      Text_grName = 'Секьюрпаки стандартные принятые из ПРК в КП'

      SELECT SecurPackPRK

      REPLACE ALL CurDate WITH new_data_odb

    * BROWSE WINDOW brows

    * ---------------------------------------------------------------------------------------- *

      IF colvo_zap_prn <> 0
        DO Formir_Table_Prn IN Formir_Vedom_PRK_KP WITH 2                                                            && Печать суперсумок
      ENDIF
   ENDSCAN 
   USE IN _TranSheets
* ---------------------------------------------------------------------------------------- *

ENDIF

DEACTIVATE WINDOW poisk

RETURN


******************************************************************* PROCEDURE start_vedom_vip *******************************************************************

PROCEDURE start_vedom_vip

tim1 = SECONDS()

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Формирования печатного отчета об переданных сумках из ПРК в КП ...',WCOLS())

* WAIT WINDOW 'f_IdShift = ' + ALLTRIM(STR(f_IdShift)) TIMEOUT 3
* WAIT WINDOW 'f_RecordRouteHistory = ' + ALLTRIM(f_RecordRouteHistory) TIMEOUT 3
* WAIT WINDOW 'f_IdTransferSheet = ' + ALLTRIM(STR(f_IdTransferSheet)) TIMEOUT 3

* ---------------------------------------------------------------------------- Обрабатываем Секьюрпаки с ПЕРЕСЧЁТОМ рублевые ------------------------------------------------------------------------------------ *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT bgh.*, trs.*, rth.*, f.FilialCode, f.FilialName, 1 AS Sost, usr.UserName
  FROM BudgetHistoryPRK AS bgh (READPAST)
  INNER JOIN TransferSheet AS trs (READPAST) ON trs.Id = bgh.IdTranSheet
  INNER JOIN RouteHistory AS rth (READPAST) ON rth.Id = bgh.IdRouteHistory
  INNER JOIN Filial AS f (READPAST) ON bgh.IdFilial = f.Id
  INNER JOIN [User] AS usr (READPAST) ON usr.Id = bgh.IdUser
  WHERE rth.IdShift = ?f_IdShift_1 AND bgh.WorkType = '0' AND bgh.Flg2PIN = 0 AND bgh.SelDat = 1 AND bgh.IdTranSheet IN (<<f_IdTransferSheet>>)

ENDTEXT

=RunSQL_Load('Priem_PRK_KP', @ar)

SELECT Priem_PRK_KP

colvo_zap = RECCOUNT()

* BROWSE WINDOW brows TITLE 'Обрабатываем Секьюрпаки с ПЕРЕСЧЁТОМ рублевые'

IF colvo_zap <> 0

  f_IncFullName = ''

  SELECT f_IncFullName AS IncFullName, sp.* ;
    FROM Priem_PRK_KP AS sp ;
    WHERE sp.IdSignProc = ?f_IdSignProc ;
    INTO CURSOR SecurPackPRK READWRITE ;
    ORDER BY SuperBudget, WorkType, BudgetNum

  colvo_zap_prn = _TALLY

  Text_grName = 'Секьюрпаки стандартные принятые из ПРК в КП'

  SELECT SecurPackPRK

  REPLACE ALL CurDate WITH new_data_odb

* BROWSE WINDOW brows

* ---------------------------------------------------------------------------------------- *

  IF colvo_zap_prn <> 0
    DO Formir_Table_Prn IN Formir_Vedom_PRK_KP WITH 0
  ENDIF

* ---------------------------------------------------------------------------------------- *

ENDIF

* -------------------------------------------------------------------- Обрабатываем МНОГОПИНОВЫЕ Секьюрпаки ПЕРЕСЧЁТОМ рублевые ----------------------------------------------------------------------- *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT bgh.*, trs.*, rth.*, f.FilialCode, f.FilialName, 1 AS Sost, usr.UserName
  FROM BudgetHistoryPRK AS bgh (READPAST)
  INNER JOIN TransferSheet AS trs (READPAST) ON trs.Id = bgh.IdTranSheet
  INNER JOIN RouteHistory AS rth (READPAST) ON rth.Id = bgh.IdRouteHistory
  INNER JOIN Filial AS f (READPAST) ON bgh.IdFilial = f.Id
  INNER JOIN [User] AS usr (READPAST) ON usr.Id = bgh.IdUser
  WHERE rth.IdShift = ?f_IdShift_1 AND bgh.WorkType = '0' AND bgh.Flg2PIN > 0 AND bgh.SelDat = 1 AND bgh.IdTranSheet IN (<<f_IdTransferSheet>>)
  ORDER BY bgh.BudgetNum, bgh.TypeValueKind

ENDTEXT

=RunSQL_Load('Priem_PRK_KP', @ar)

SELECT Priem_PRK_KP

colvo_zap = RECCOUNT()

* BROWSE WINDOW brows TITLE 'Обрабатываем МНОГОПИНОВЫЕ Секьюрпаки ПЕРЕСЧЁТОМ рублевые'

IF colvo_zap <> 0

  SELECT BudgetNum AS BudgetNum2Pin, f_IncFullName AS IncFullName, sp.* ;
    FROM Priem_PRK_KP AS sp ;
    WHERE sp.IdSignProc = ?f_IdSignProc ;
    INTO CURSOR SecurPackPRK READWRITE ;
    ORDER BY sp.BudgetNum, sp.TypeValueKind

  colvo_zap_prn = _TALLY

* BROWSE WINDOW brows TITLE 'INTO CURSOR SecurPackPRK READWRITE'

  Text_grName = 'Ведомость сдачи инкассаторских сумок (мешков)'

  SELECT SecurPackPRK

  REPLACE ALL CurDate WITH new_data_odb

  REPLACE BudgetNum2Pin WITH STRTRAN(BudgetNum2Pin, '/2', SPACE(0)) FOR '/2' $ BudgetNum2Pin

  REPLACE BudgetNum2Pin WITH STRTRAN(BudgetNum2Pin, '/3', SPACE(0)) FOR '/3' $ BudgetNum2Pin

* BROWSE WINDOW brows TITLE 'REPLACE BudgetNum2Pin'

* ---------------------------------------------------------------------------------------- *

  IF colvo_zap_prn <> 0
    DO Formir_Table_Prn IN Formir_Vedom_PRK_KP WITH 1
  ENDIF

    * ---------------------------------------------------------------------------------------- *

ENDIF

* ---------------------------------------------------------------------------- Обрабатываем Супер сумки с ПЕРЕСЧЁТОМ рублевые ------------------------------------------------------------------------------------ *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT bgh.*, trs.*, rth.*, f.FilialCode, f.FilialName, 1 AS Sost, usr.UserName
  FROM BudgetHistoryPRK AS bgh (READPAST)
  INNER JOIN TransferSheet AS trs (READPAST) ON trs.Id = bgh.IdTranSheet
  INNER JOIN RouteHistory AS rth (READPAST) ON rth.Id = bgh.IdRouteHistory
  INNER JOIN Filial AS f (READPAST) ON bgh.IdFilial = f.Id
  INNER JOIN [User] AS usr (READPAST) ON usr.Id = bgh.IdUser
  WHERE rth.IdShift = ?f_IdShift_1 AND bgh.WorkType IN ('1', '2') AND bgh.Flg2PIN = 0 AND bgh.SelDat = 1 AND bgh.IdTranSheet IN (<<f_IdTransferSheet>>)

ENDTEXT

=RunSQL_Load('Priem_PRK_KP', @ar)

SELECT Priem_PRK_KP

colvo_zap = RECCOUNT()

* BROWSE WINDOW brows TITLE 'Обрабатываем Супер сумки с ПЕРЕСЧЁТОМ рублевые'

IF colvo_zap <> 0

  f_IncFullName = ''

  SELECT SPACE(20) AS Budget2Pin, f_IncFullName AS IncFullName, sp.* ;
    FROM Priem_PRK_KP AS sp ;
    WHERE sp.IdSignProc = ?f_IdSignProc ;
    INTO CURSOR SecurPackPRK READWRITE ;
    ORDER BY SuperBudget, WorkType, BudgetNum

  colvo_zap_prn = _TALLY

  Text_grName = 'Секьюрпаки стандартные принятые из ПРК в КП'

  SELECT SecurPackPRK

  REPLACE ALL CurDate WITH new_data_odb

* BROWSE WINDOW brows

* ---------------------------------------------------------------------------------------- *

  IF colvo_zap_prn <> 0
    DO Formir_Table_Prn IN Formir_Vedom_PRK_KP WITH 2
  ENDIF

* ---------------------------------------------------------------------------------------- *

ENDIF

DEACTIVATE WINDOW poisk

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

Num_IdTranSheet = SecurPackPRK.IdTranSheet

* ------------------------------------------------------------------------------------------------------------------------------------------------------ *

SELECT DISTINCT IdRoute ;
  FROM SecurPackPRK ;
  INTO CURSOR Select_IdRoute ;
  ORDER BY IdRoute

*  WHERE Flg2PIN = Priznak_2PIN ;

Colvo_IdRoute = _TALLY

SELECT Select_IdRoute
GO TOP
* BROWSE WINDOW brows

PRIVATE Itog_IdRoute

Itog_IdRoute = ''  && Подсчет количества сумок

DO CASE
  CASE Colvo_IdRoute = 1

    Itog_IdRoute = ALLTRIM(STR(Select_IdRoute.IdRoute))

  CASE Colvo_IdRoute > 1

    Itog_IdRoute = ALLTRIM(STR(Select_IdRoute.IdRoute))

    SCAN FOR RECNO() > 1

      Itog_IdRoute = ALLTRIM(Itog_IdRoute) + ', ' + ALLTRIM(STR(Select_IdRoute.IdRoute))

    ENDSCAN

ENDCASE

* WAIT WINDOW 'Itog_IdRoute = ' + ALLTRIM(Itog_IdRoute) TIMEOUT 3

* ------------------------------------------------------------------------------------------------------------------------------------------------------ *

DO CASE
  CASE Priznak_2PIN = 0

    SELECT DISTINCT BudgetNum ;
      FROM SecurPackPRK ;
      WHERE Flg2PIN = Priznak_2PIN AND WorkType = '0' ;
      INTO CURSOR Count_BudgetNum

    COUNT TO Itog_BudgetNum  && Подсчет количества сумок

    DO start IN Propis_chislo WITH Itog_BudgetNum
    Itog_BudgetNum_Prop = Chislo_Prop


  CASE Priznak_2PIN = 1

    SELECT DISTINCT BudgetNum2Pin ;
      FROM SecurPackPRK ;
      WHERE Flg2PIN > 0 AND WorkType = '0' ;
      INTO CURSOR Count_BudgetNum

    COUNT TO Itog_BudgetNum  && Подсчет количества сумок

    DO start IN Propis_chislo WITH Itog_BudgetNum
    Itog_BudgetNum_Prop = Chislo_Prop


  CASE Priznak_2PIN = 2

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

    DO start IN Propis_chislo WITH Itog_BudgetNum_2
    Itog_BudgetNum_Prop = Chislo_Prop

ENDCASE

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

DO CASE
  CASE Priznak_2PIN = 0

    SUM TotalSum TO Itog_TotalSum FOR WorkType = '0' AND Flg2PIN = 0  && Подсчет общей суммы

  CASE Priznak_2PIN = 1

    SUM TotalSum TO Itog_TotalSum FOR WorkType = '0' AND INLIST(Flg2PIN, 1, 2)  && Подсчет общей суммы

  CASE Priznak_2PIN = 2

    SUM TotalSum TO Itog_TotalSum FOR WorkType = '1' AND Flg2PIN = 0  && Подсчет общей суммы

ENDCASE

DO start IN Propis WITH Itog_TotalSum
Itog_TotalSum_Prop = Summa_Prop

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

CREATE CURSOR SecurPackPRK_cursor ;
  (IdBank Integer, CurDate DateTime, PacketNum Integer, BudgetNum Char(25), FilialCode Char(25), TotalSum Numeric(15, 2), TypeValueKind Char(40))

SELECT SecurPackPRK
colvo_zap = RECCOUNT()
GO TOP
* BROWSE WINDOW brows

Export_DatePRK = TTOC(SecurPackPRK.ChangeDatePRK)

SELECT SecurPackPRK_cursor

FOR I = 1 TO colvo_zap
  APPEND BLANK
ENDFOR

GO TOP
* BROWSE WINDOW brows

SELECT SecurPackPRK
GO TOP
* BROWSE WINDOW brows

f_BudgetNum_pred = ''

rec_num = 0

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SCAN

  rec_num = rec_num + 1

  SELECT SecurPackPRK_cursor

  SCATTER MEMVAR

  m.IdBank = 1
  m.CurDate = SecurPackPRK.CurDate

  DO CASE
    CASE SecurPackPRK.Flg2PIN = 0

      m.BudgetNum = ALLTRIM(SecurPackPRK.BudgetNum)

      m.FilialCode = ALLTRIM(SecurPackPRK.FilialCode)

      m.PacketNum = IIF(ALLTRIM(m.BudgetNum) <> ALLTRIM(f_BudgetNum_pred), rec_num, 0)

    CASE SecurPackPRK.Flg2PIN = 1

      m.BudgetNum = ALLTRIM(SecurPackPRK.BudgetNum)

      m.FilialCode = ALLTRIM(SecurPackPRK.FilialCode)

      m.PacketNum = IIF(ALLTRIM(m.BudgetNum) <> ALLTRIM(f_BudgetNum_pred), rec_num, 0)

    CASE SecurPackPRK.Flg2PIN > 1

      m.BudgetNum = STRTRAN(SecurPackPRK.BudgetNum, '/2', '')

      m.BudgetNum = STRTRAN(m.BudgetNum, '/3', '')

      m.FilialCode = ALLTRIM(SecurPackPRK.FilialCode)

      IF ALLTRIM(m.BudgetNum) == ALLTRIM(f_BudgetNum_pred)
        rec_num = rec_num - 1
      ENDIF

      m.PacketNum = IIF(ALLTRIM(m.BudgetNum) <> ALLTRIM(f_BudgetNum_pred), rec_num, 0)

  ENDCASE

  m.TotalSum = SecurPackPRK.TotalSum

  m.TypeValueKind = IIF(ISNULL(SecurPackPRK.TypeValueKind) = .F., ALLTRIM(SecurPackPRK.TypeValueKind), '')

  GATHER MEMVAR

  SELECT SecurPackPRK
  rec_new = RECNO()

  SELECT SecurPackPRK_cursor
  SKIP

  IF EOF() = .T.
    EXIT
  ENDIF

  f_BudgetNum_pred = ALLTRIM(m.BudgetNum)

  SELECT SecurPackPRK
ENDSCAN

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SELECT SecurPackPRK_cursor
colvo_zap = RECCOUNT()
GO TOP
* BROWSE WINDOW brows

* WAIT WINDOW 'colvo_zap = ' + ALLTRIM(STR(colvo_zap)) TIMEOUT 3

tim2 = SECONDS()

result_time = ROUND(tim2 - tim1, 2)

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Формирования печатного отчета успешно завершено.' + ' (время = ' + ALLTRIM(TRANSFORM(result_time, '999.99')) + ' сек.)',WCOLS())
=INKEY(2)
DEACTIVATE WINDOW poisk

WAIT CLEAR

* ---------------------------------------------------------------------------- *
DO SecurPackPRK_Print IN Formir_Vedom_PRK_KP
* ---------------------------------------------------------------------------- *

RETURN


************************************************************************* PROCEDURE SecurPackPRK_Print ***************************************************

PROCEDURE SecurPackPRK_Print

IF colvo_zap <> 0

  DO CASE
    CASE Priznak_2PIN < 2

      DO CASE
        CASE sergeizhuk_flag = .T.

          REPORT FORM KPpakpriem_blank_1 PREVIEW

        CASE sergeizhuk_flag = .F.

          DO CASE
            CASE Vedom_prn = .T.  && Признак выдачи отчета на принтер, когда идет первичная печать TO PRINTER NOCONSOLE

              REPORT FORM KPpakpriem_blank_1 PREVIEW

            CASE Vedom_prn = .F.  && Признак выдачи отчета на экран, когда идет повторная печать

              REPORT FORM KPpakpriem_blank_1 PREVIEW

          ENDCASE

      ENDCASE


    CASE Priznak_2PIN = 2

      DO CASE
        CASE sergeizhuk_flag = .T.

          REPORT FORM KPsupersumka_blank_1 PREVIEW

        CASE sergeizhuk_flag = .F.

          DO CASE
            CASE Vedom_prn = .T.  && Признак выдачи отчета на принтер, когда идет первичная печать TO PRINTER NOCONSOLE

              REPORT FORM KPsupersumka_blank_1 PREVIEW

            CASE Vedom_prn = .F.  && Признак выдачи отчета на экран, когда идет повторная печать

              REPORT FORM KPsupersumka_blank_1 PREVIEW

          ENDCASE

      ENDCASE

  ENDCASE

ENDIF

SELECT SecurPackPRK_cursor
USE

RETURN


***************************************************************************** FUNCTION CloseDB *****************************************************************

FUNCTION CloseDB

IF USED('Vedom') = .T.
  SELECT Vedom
  USE
ENDIF

ENDFUNC


*****************************************************************************************************************************************************************


