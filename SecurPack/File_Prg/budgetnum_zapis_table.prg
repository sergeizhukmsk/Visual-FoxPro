******************************************************************** PROCEDURE start ************************************************************

PROCEDURE start

sel = SELECT()

* f_IdBudgetKP = LstBudget.Id  && Запоминаем Id выбранной сумки для отправки в первой экранной форме

*** Локальный запрос к таблице используемой в гриде экранной формы, по Id выбранной сумки, сумка должна быть строго одна, статус у нее State = 0 ***

* WAIT WINDOW 'f_WorkType = ' + ALLTRIM(f_WorkType) TIMEOUT 1
* WAIT WINDOW 'f_IdBudgetKP = ' + ALLTRIM(STR(f_IdBudgetKP)) TIMEOUT 1
* WAIT WINDOW 'f_SuperBudget = ' + ALLTRIM(STR(f_SuperBudget)) TIMEOUT 3
* WAIT WINDOW 'f_Flg2Pin = ' + ALLTRIM(STR(f_Flg2Pin)) TIMEOUT 3

DO CASE
  CASE f_WorkType = '0' AND f_Flg2Pin = 0  && ТИП выбранной сумки - стандартная

    SELECT * ;
      FROM LstBudget ;
      WHERE Id = f_IdBudgetKP ;
      INTO CURSOR LstBugCl

  CASE f_WorkType = '0' AND f_Flg2Pin = 1  && ТИП выбранной сумки - главная многопиновая

    SELECT * ;
      FROM LstBudget ;
      WHERE SuperBudget = f_SuperBudget ;
      INTO CURSOR LstBugCl

  CASE f_WorkType = '0' AND f_Flg2Pin = 2  && ТИП выбранной сумки - дополнительная многопиновая

    SELECT * ;
      FROM LstBudget ;
      WHERE SuperBudget = f_SuperBudget ;
      INTO CURSOR LstBugCl

  CASE f_WorkType = '1' AND f_Flg2Pin = 0  && ТИП выбранной сумки - суперсумка

    SELECT * ;
      FROM LstBudget ;
      WHERE SuperBudget = f_SuperBudget ;
      INTO CURSOR LstBugCl

  CASE f_WorkType = '2' AND f_Flg2Pin = 0  && ТИП выбранной сумки - вложенная сумка

    SELECT * ;
      FROM LstBudget ;
      WHERE Id = f_IdBudgetKP ;
      INTO CURSOR LstBugCl

ENDCASE

SELECT LstBugCl
GO TOP

* BROWSE WINDOW brows

IF _TALLY = 0
  MESSAGEBOX('Выбранная сумка не доступна для зачисления ', 64, 'Сообщение')
  RETURN
ENDIF

SELECT LstBugCL
GO TOP
* BROWSE WINDOW brows

* ---------------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
  BEGIN TRANSACTION
ENDTEXT

rc = RunSQL_Load('', @ar)

IF rc = 1

  SCAN

    f_IdBudgetKP = LstBugCL.Id

    f_IdFilial = LstBugCL.IdFilial
    f_FilialCode = ALLTRIM(LstBugCL.FilialCode)

    f_BudgetNum = ALLTRIM(LstBugCL.BudgetNum)
    f_TotalSumPRK = LstBugCL.TotalSumPRK && Сумма
    f_TotalSumKP = LstBugCL.TotalSumKP && Сумма

    f_WorkType = LstBugCL.WorkType
    f_SuperBudget = LstBugCL.SuperBudget
    f_SuperQty = LstBugCL.SuperQty
    f_Flg2Pin = LstBugCL.Flg2Pin

    f_DtIncome = TTOD(LstBugCL.DtIncome)    && Дата выручки
    f_DateTransfer = LstBugCL.DateTransfer

    f_Account = ALLTRIM(LstBugCL.Account)
    f_IncomeAccount = ALLTRIM(LstBugCL.IncomeAccount)

* WAIT WINDOW 'f_IdBudgetKP = ' + ALLTRIM(STR(f_IdBudgetKP)) TIMEOUT 3

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

      BEGIN

      IF NOT EXISTS(SELECT Id FROM ExportNewSmenaKP WHERE Id = ?f_IdBudgetKP)

        INSERT INTO ExportNewSmenaKP
        (Id, IdFilial, IdOperDate, IdShiftIn, IdShiftOut, DateTransfer, FilialCode, BudgetNum, TotalSumPRK, TotalSumKP, State, WorkType, SuperBudget, SuperQty, Flg2Pin, Account, IncomeAccount, ImpPref)
        VALUES
        (?f_IdBudgetKP, ?f_IdFilial, ?f_IdOperDate, ?f_StrIdSmena, ?f_NewIdSmena, ?f_DateTransfer, ?f_FilialCode, ?f_BudgetNum, ?f_TotalSumPRK, ?f_TotalSumKP, ?f_State, ?f_WorkType, ?f_SuperBudget, ?f_SuperQty, ?f_Flg2Pin, ?f_Account, ?f_IncomeAccount, 0)

      ELSE

        UPDATE ExportNewSmenaKP SET
        Id = ?f_IdBudgetKP,
        IdFilial = ?f_IdFilial,
        IdShiftIn = ?f_StrIdSmena,
        IdShiftOut = ?f_NewIdSmena,
        DateTransfer = ?f_DateTransfer,
        IdOperDate = ?f_IdOperDate,
        FilialCode = ?f_FilialCode,
        BudgetNum = ?f_BudgetNum,
        TotalSumPRK = ?f_TotalSumPRK,
        TotalSumKP = ?f_TotalSumKP,
        State = ?f_State,
        WorkType = ?f_WorkType,
        SuperBudget = ?f_SuperBudget,
        SuperQty = ?f_SuperQty,
        Flg2Pin = ?f_Flg2Pin,
        Account = ?f_Account,
        IncomeAccount = ?f_IncomeAccount,
        ImpPref = 0,
        DtRecord = GetDate()
        WHERE Id = ?f_IdBudgetKP

        SELECT *
        FROM ExportNewSmenaKP
        WHERE Id = ?f_IdBudgetKP

       END

    ENDTEXT

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*      TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

*        EXECUTE Insert_Update_ExportNewSmenaKP
*        ?f_IdBudgetKP, ?f_IdFilial, ?f_IdOperDate, ?f_StrIdSmena, ?f_NewIdSmena, ?f_DateTransfer, ?f_FilialCode, ?f_BudgetNum, ?f_TotalSumKP, ?f_State, ?f_Account, ?f_IncomeAccount

*      ENDTEXT

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    rc = RunSQL_Load('ExportNewSmenaKP', @ar)

    IF rc <> 1
      MESSAGEBOX('Таблица ExportNewSmenaKP не сформирована', 16, 'Обратитесь к Программистам по тел: 102, 267, 158')
      RETURN -1
    ENDIF

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    colvo_zap = 0

    IF USED('ExportNewSmenaKP') = .T.

      SELECT ExportNewSmenaKP

      colvo_zap = RECCOUNT()

      GO TOP

*      BROWSE WINDOW brows

      IF colvo_zap > 0

        SELECT LstBudget
        DELETE FOR LstBudget.Id = f_IdBudgetKP
        GO TOP

      ENDIF

    ENDIF

    SELECT LstBugCL
  ENDSCAN

  DO CASE
    CASE rc = 1

      TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
        COMMIT TRANSACTION
      ENDTEXT

    CASE rc <> 1

      TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
        ROLLBACK TRANSACTION
      ENDTEXT

  ENDCASE

  rc = RunSQL_Load('', @ar)

* ---------------------------------------------------------------- *
  DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

ENDIF

SELECT(sel)
RETURN


**************************************************************************************************************************************************





