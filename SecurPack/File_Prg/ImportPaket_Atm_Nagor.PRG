*********************************************************************** PROCEDURE start *****************************************************

PROCEDURE start
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

DIMENSION ar[1,2]

* WAIT WINDOW 'f_IdDepartmentPRK = ' + ALLTRIM(STR(f_IdDepartmentPRK)) TIMEOUT 3

put_select = 1

dtB = new_data_odb  && Дата открытой смены в ПРК
f_CurDate = new_data_odb  && Дата открытой смены в ПРК

* WAIT WINDOW 'dtB = ' + DTOC(dtB) TIMEOUT 3
* WAIT WINDOW '1 - f_CurDate = ' + DTOC(f_CurDate) TIMEOUT 3

f_FormName = 'Информация'

* ------------------------------------ *
DO FORM GetDate1
* ------------------------------------ *

* ---------------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

  f_CurDate = dtb

* WAIT WINDOW '2 - f_CurDate = ' + DTOC(f_CurDate) TIMEOUT 3

* ---------------------------------------- Находим смены ПРК за дату f_DtBegin ------------------------------------- *

  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

    SELECT Id, IdOperDate, IdDepartment, IdUser, Status, ShiftDate, ReportDate, OpenDate, CloseDate, Blok
    FROM Shift (READPAST)
    WHERE ReportDate = ?f_CurDate AND IdDepartment = ?f_IdDepartmentPRK
    ORDER BY Id DESC

  ENDTEXT

  IF RunSql_Load('Shift_PRK', @ar) <> 1
    DO exit_sql_server IN RunSQL_Load
    MESSAGEBOX('Ошибка работы со справочником Смен ... ', 16, 'Смена не найдена.')
    RETURN -1
  ENDIF

  DO CASE
    CASE USED('Shift_PRK') = .T. AND RECCOUNT('Shift_PRK') > 0

      SELECT Shift_PRK
      GO TOP

      IdShift_CurDate = Shift_PRK.Id

      ReportDate_PRK = TTOD(Shift_PRK.ReportDate)
      ReportDate_ATM = TTOD(Shift_PRK.ReportDate)

      ShiftDate_PRK = TTOD(Shift_PRK.ShiftDate)
      ShiftDate_ATM = TTOD(Shift_PRK.ShiftDate)

*      BROWSE WINDOW brows

      IdShift_PRK = ''

      SCAN

        IdShift_PRK = IdShift_PRK + ALLTRIM(STR(Shift_PRK.Id)) + [,]  && Список Id открытых смен

      ENDSCAN

      IdShift_PRK = IdShift_PRK + ['']  && Список Id открытых смен

*      WAIT WINDOW 'IdShift_PRK = ' + ALLTRIM(IdShift_PRK) TIMEOUT 3

      SELECT Shift_PRK
      USE

    CASE USED('Shift_PRK') = .F.

      IdShift_PRK = ''

  ENDCASE

* ---------------------------------------------------------------------------------------- Проверка правильности выборки номера смены из справочника ------------------------------------------------------------------------------- *

  IF EMPTY(IdShift_PRK) = .T.

    DO exit_sql_server IN RunSQL_Load
    MESSAGEBOX('Внимание! По введенной Вами дате - ' + DTOC(f_CurDate) + CHR(13) + 'номер смены в справочнике - НЕ НАЙДЕН', 16, 'Выборка номера смены по дате')
    RETURN

  ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  f_IdSignProc = RunSignProc(PlacePRK)  && SignProc.Id = 15

* WAIT WINDOW 'f_IdShift_1 = ' + ALLTRIM(STR(f_IdShift_1)) TIMEOUT 3
* WAIT WINDOW 'IdShift_PRK = ' + ALLTRIM(IdShift_PRK) TIMEOUT 3
* WAIT WINDOW 'f_IdSignProc = ' + ALLTRIM(STR(f_IdSignProc)) TIMEOUT 3
* WAIT WINDOW 'f_CurDate 1 = ' + DTOC(f_CurDate) TIMEOUT 3

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk

  @ WROWS()/3, 3 SAY PADC('Уважаемый Пользователь! Начат импорт баулов с кассетами АТМ за дату - ' + DTOC(f_CurDate), WCOLS())


* ----------------  Заполняем таблицу BudgetHistoryATM_Out данными по пересчетным баулам переданным из цеха АТМ в ПРК для выдачи на маршрут------------------------------- *

*  DO start IN SendBudget_Atm
  DO start IN SendBudget_Atm_Str

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW 'f_CurDate 2 = ' + DTOC(f_CurDate) TIMEOUT 3
* WAIT WINDOW 'f_IdSignProc = ' + ALLTRIM(STR(f_IdSignProc)) TIMEOUT 3
* WAIT WINDOW 'IdShift_PRK = ' + ALLTRIM(IdShift_PRK) TIMEOUT 5

  TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT pack_prk.Id AS IdSecurPack, pack_prk.IdShift, pack_prk.IdUser, pack_prk.IdRoute, atm.Id AS IdATM, pack_prk.BudgetNum, pack_prk.WorkType, bdg_prk.SuperBudget, bdg_prk.SuperQty,
      bdg_prk.Id AS IdBudgetHistory, bdg_prk.IdFilial, pack_prk.FilialCode, bdg_atm.Id AS IdBudgetKPAtm, pack_prk.IdLoad, atm.NumerATM, atm.NumerATMReq,
      pack_prk.TabNumber, pack_prk.DateInkas, pack_prk.DtIncome, pack_prk.CurDate,
      bdg_atm.TotalSumKP_RUB AS TotalSum_RUB_Out,
      bdg_atm.TotalSumKP_USD AS TotalSum_USD_Out,
      bdg_atm.TotalSumKP_EUR AS TotalSum_EUR_Out,
      cln.ClientName, bdg_prk.IdVedList, bdg_prk.IdSignProc, pack_prk.IsTransit, pack_prk.DtRecord
      FROM SecurPack_ATM_Nagor AS pack_prk (READPAST)
      INNER JOIN BudgetHistory_ATM_Nagor AS bdg_prk (READPAST) ON bdg_prk.IdSecurPack = pack_prk.Id
      INNER JOIN R71BudgetKP_Atm AS bdg_atm (READPAST) ON bdg_atm.IdBudgetHistoryPRK = bdg_prk.Id
      INNER JOIN Filial AS fil (READPAST) ON fil.Id = bdg_prk.IdFilial
      INNER JOIN Client AS cln (READPAST) ON cln.Id = fil.IdClient
      INNER JOIN BaseATM AS atm (READPAST) ON atm.IdFilial = bdg_prk.IdFilial
      WHERE pack_prk.IdShift IN (<<IdShift_PRK>>) AND bdg_prk.IdSignProc = ?f_IdSignProc AND pack_prk.WorkType = '1' AND pack_prk.IdShift <> 0
      UNION
      SELECT pack_prk.Id AS IdSecurPack, pack_prk.IdShift, pack_prk.IdUser, pack_prk.IdRoute, atm.Id AS IdATM, pack_prk.BudgetNum, pack_prk.WorkType, bdg_prk.SuperBudget, bdg_prk.SuperQty,
      bdg_prk.Id AS IdBudgetHistory, bdg_prk.IdFilial, pack_prk.FilialCode, bdg_prk.Id AS IdBudgetKPAtm, pack_prk.IdLoad, atm.NumerATM, atm.NumerATMReq,
      pack_prk.TabNumber, pack_prk.DateInkas, pack_prk.DtIncome, pack_prk.CurDate,
      pack_prk.CountedSum_RUB AS TotalSum_RUB_Out,
      pack_prk.CountedSum_USD AS TotalSum_USD_Out,
      pack_prk.CountedSum_EUR AS TotalSum_EUR_Out,
      cln.ClientName, bdg_prk.IdVedList, bdg_prk.IdSignProc, pack_prk.IsTransit, pack_prk.DtRecord
      FROM SecurPack_ATM_Nagor AS pack_prk (READPAST)
      INNER JOIN BudgetHistory_ATM_Nagor AS bdg_prk (READPAST) ON bdg_prk.IdSecurPack = pack_prk.Id
      INNER JOIN Filial AS fil (READPAST) ON fil.Id = bdg_prk.IdFilial
      INNER JOIN Client AS cln (READPAST) ON cln.Id = fil.IdClient
      INNER JOIN BaseATM AS atm (READPAST) ON atm.IdFilial = bdg_prk.IdFilial
      WHERE pack_prk.IdShift IN (<<IdShift_PRK>>) AND bdg_prk.IdSignProc = ?f_IdSignProc AND pack_prk.WorkType = '2' AND pack_prk.IdShift <> 0 AND pack_prk.OutOfSafe = 1
      UNION
      SELECT pack_prk.Id AS IdSecurPack, pack_prk.IdShift, pack_prk.IdUser, pack_prk.IdRoute, ISNULL(atm.Id, 0) AS IdATM, pack_prk.BudgetNum, pack_prk.WorkType, bdg_prk.SuperBudget, bdg_prk.SuperQty,
      ISNULL(bdg_prk.Id, 0) AS IdBudgetHistory, ISNULL(bdg_prk.IdFilial, 0) AS IdFilial, pack_prk.FilialCode, ISNULL(bdg_prk.Id, 0) AS IdBudgetKPAtm, pack_prk.IdLoad, atm.NumerATM, atm.NumerATMReq,
      pack_prk.TabNumber, pack_prk.DateInkas, pack_prk.DtIncome, pack_prk.CurDate,
      pack_prk.CountedSum_RUB AS TotalSum_RUB_Out,
      pack_prk.CountedSum_USD AS TotalSum_USD_Out,
      pack_prk.CountedSum_EUR AS TotalSum_EUR_Out,
      cln.ClientName, ISNULL(bdg_prk.IdVedList, 0) AS IdVedList, ISNULL(bdg_prk.IdSignProc, 0) AS IdSignProc, pack_prk.IsTransit, pack_prk.DtRecord
      FROM SecurPack_ATM_Nagor AS pack_prk (READPAST)
      LEFT OUTER JOIN BudgetHistory_ATM_Nagor AS bdg_prk (READPAST) ON bdg_prk.IdSecurPack = pack_prk.Id
      LEFT OUTER JOIN Filial AS fil (READPAST) ON fil.Id = bdg_prk.IdFilial
      LEFT OUTER JOIN Client AS cln (READPAST) ON cln.Id = fil.IdClient
      LEFT OUTER JOIN BaseATM AS atm (READPAST) ON atm.IdFilial = bdg_prk.IdFilial
      WHERE pack_prk.IdShift IN (<<IdShift_PRK>>) AND pack_prk.WorkType = '2' AND pack_prk.IdShift <> 0 AND pack_prk.OutOfSafe = 1
      ORDER BY IdShift, DtRecord

  ENDTEXT

  IF RunSql_Load('SecurPack_ATM', @ar) <> 1
    DEACTIVATE WINDOW poisk
    DO exit_sql_server IN RunSQL_Load
    MESSAGEBOX('Ошибка выборки данных из таблицы FROM SecurPack_ATM', 16, 'Обратитесь к Программистам по тел: 1243, 1303')
    RETURN -1
  ENDIF

* BROWSE WINDOW brows

* -------------------------------------------------------------- Выполняем запрос без участия таблицы R71BudgetKP_Atm, это старый алгоритм ------------------------------------------------------------------------------------------------- *

  IF USED('SecurPack_ATM') = .T.

    colvo_zap = RECCOUNT()

    IF colvo_zap = 0

      TEXT TO SqlReq NOSHOW TEXTMERGE && PRETEXT 7

        SELECT pack_prk.Id AS IdSecurPack, pack_prk.IdShift, pack_prk.IdUser, pack_prk.IdRoute, atm.Id AS IdATM, pack_prk.BudgetNum, fil.MestoCode, pack_prk.WorkType, bdg_prk.SuperBudget, bdg_prk.SuperQty,
        bdg_prk.Id AS IdBudgetHistory, bdg_prk.IdFilial, pack_prk.FilialCode, bdg_prk.Id AS IdBudgetKPAtm, pack_prk.IdLoad, atm.NumerATM, atm.NumerATMReq,
        pack_prk.TabNumber, pack_prk.DateInkas, pack_prk.DtIncome, pack_prk.CurDate,
        pack_prk.TotalSum_RUB AS TotalSum_RUB_Out,
        pack_prk.TotalSum_USD AS TotalSum_USD_Out,
        pack_prk.TotalSum_EUR AS TotalSum_EUR_Out,
        cln.ClientName, bdg_prk.IdVedList, bdg_prk.IdSignProc, pack_prk.IsTransit, pack_prk.DtRecord
        FROM SecurPack_ATM_Nagor AS pack_prk (READPAST)
        INNER JOIN BudgetHistory_ATM_Nagor AS bdg_prk (READPAST) ON bdg_prk.IdSecurPack = pack_prk.Id
        INNER JOIN Filial AS fil (READPAST) ON fil.Id = bdg_prk.IdFilial
        INNER JOIN Client AS cln (READPAST) ON cln.Id = fil.IdClient
        INNER JOIN BaseATM AS atm (READPAST) ON atm.IdFilial = bdg_prk.IdFilial
        WHERE pack_prk.IdShift IN (<<IdShift_PRK>>) AND bdg_prk.IdSignProc = ?f_IdSignProc AND pack_prk.WorkType = '1' AND pack_prk.IdShift <> 0
        UNION
        SELECT pack_prk.Id AS IdSecurPack, pack_prk.IdShift, pack_prk.IdUser, pack_prk.IdRoute, atm.Id AS IdATM, pack_prk.BudgetNum, fil.MestoCode, pack_prk.WorkType, bdg_prk.SuperBudget, bdg_prk.SuperQty,
        bdg_prk.Id AS IdBudgetHistory, bdg_prk.IdFilial, pack_prk.FilialCode, bdg_prk.Id AS IdBudgetKPAtm, pack_prk.IdLoad, atm.NumerATM, atm.NumerATMReq,
        pack_prk.TabNumber, pack_prk.DateInkas, pack_prk.DtIncome, pack_prk.CurDate,
        pack_prk.CountedSum_RUB AS TotalSum_RUB_Out,
        pack_prk.CountedSum_USD AS TotalSum_USD_Out,
        pack_prk.CountedSum_EUR AS TotalSum_EUR_Out,
        cln.ClientName, bdg_prk.IdVedList, bdg_prk.IdSignProc, pack_prk.IsTransit, pack_prk.DtRecord
        FROM SecurPack_ATM_Nagor AS pack_prk (READPAST)
        INNER JOIN BudgetHistory_ATM_Nagor AS bdg_prk (READPAST) ON bdg_prk.IdSecurPack = pack_prk.Id
        INNER JOIN Filial AS fil (READPAST) ON fil.Id = bdg_prk.IdFilial
        INNER JOIN Client AS cln (READPAST) ON cln.Id = fil.IdClient
        INNER JOIN BaseATM AS atm (READPAST) ON atm.IdFilial = bdg_prk.IdFilial
        WHERE pack_prk.IdShift IN (<<IdShift_PRK>>) AND bdg_prk.IdSignProc = ?f_IdSignProc AND pack_prk.WorkType = '2' AND pack_prk.IdShift <> 0 AND pack_prk.OutOfSafe = 1
        UNION
        SELECT pack_prk.Id AS IdSecurPack, pack_prk.IdShift, pack_prk.IdUser, pack_prk.IdRoute, ISNULL(atm.Id, 0) AS IdATM, pack_prk.BudgetNum, fil.MestoCode, pack_prk.WorkType, bdg_prk.SuperBudget, bdg_prk.SuperQty,
        ISNULL(bdg_prk.Id, 0) AS IdBudgetHistory, ISNULL(bdg_prk.IdFilial, 0) AS IdFilial, pack_prk.FilialCode, ISNULL(bdg_prk.Id, 0) AS IdBudgetKPAtm, pack_prk.IdLoad, atm.NumerATM, atm.NumerATMReq,
        pack_prk.TabNumber, pack_prk.DateInkas, pack_prk.DtIncome, pack_prk.CurDate,
        pack_prk.CountedSum_RUB AS TotalSum_RUB_Out,
        pack_prk.CountedSum_USD AS TotalSum_USD_Out,
        pack_prk.CountedSum_EUR AS TotalSum_EUR_Out,
        cln.ClientName, ISNULL(bdg_prk.IdVedList, 0) AS IdVedList, ISNULL(bdg_prk.IdSignProc, 0) AS IdSignProc, pack_prk.IsTransit, pack_prk.DtRecord
        FROM SecurPack_ATM_Nagor AS pack_prk (READPAST)
        LEFT OUTER JOIN BudgetHistory_ATM_Nagor AS bdg_prk (READPAST) ON bdg_prk.IdSecurPack = pack_prk.Id
        LEFT OUTER JOIN Filial AS fil (READPAST) ON fil.Id = bdg_prk.IdFilial
        LEFT OUTER JOIN Client AS cln (READPAST) ON cln.Id = fil.IdClient
        LEFT OUTER JOIN BaseATM AS atm (READPAST) ON atm.IdFilial = bdg_prk.IdFilial
        WHERE pack_prk.IdShift IN (<<IdShift_PRK>>) AND pack_prk.WorkType = '2' AND pack_prk.IdShift <> 0 AND pack_prk.OutOfSafe = 1
        ORDER BY pack_prk.IdShift, pack_prk.DtRecord

      ENDTEXT

      IF RunSql_Load('SecurPack_ATM', @ar) <> 1
        DEACTIVATE WINDOW poisk
        DO exit_sql_server IN RunSQL_Load
        MESSAGEBOX('Ошибка выборки данных из таблицы FROM SecurPack_ATM', 16, 'Обратитесь к Программистам по тел: 243, 303')
        RETURN -1
      ENDIF

* BROWSE WINDOW brows

    ENDIF
  ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  IF USED('SecurPack_ATM') = .T.
    IF ar[1,2] = 0
      USE IN SecurPack_ATM
      DEACTIVATE WINDOW poisk
      DO exit_sql_server IN RunSQL_Load
      MESSAGEBOX('Данных за дату ' + DTOC(f_CurDate) + CHR(13) + 'доступных для контроля НЕ ОБНАРУЖЕНО', 0, 'Внимательно проверьте введенные Вами данные')
      RETURN -1
    ENDIF
  ENDIF

* BROWSE WINDOW brows

* WAIT WINDOW 'f_CurDate = ' + DTOC(f_CurDate) + '  new_data_odb = ' + DTOC(new_data_odb) TIMEOUT 5

  IF USED('SecurPack_ATM') = .T.

    SELECT SecurPack_ATM

    colvo_zap = RECCOUNT()

* BROWSE WINDOW brows

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SELECT SecurPack_ATM

    colvo_zap = RECCOUNT()

    GO TOP

*    BROWSE WINDOW brows

    IF colvo_zap > 0

      colvo_zap_prk = 0

      COUNT TO colvo_zap_prk

      WAIT WINDOW 'Всего за выбранную дату пересчетных Баулов АТМ из ПРК = ' + ALLTRIM(TRANSFORM(colvo_zap_prk, '999 999 999')) TIMEOUT 5

      SELECT SecurPack_ATM
      GO TOP

*    BROWSE WINDOW brows

      SCAN

        trn_IdSecurPack = SecurPack_ATM.IdSecurPack
        trn_IdBudgetHistory = SecurPack_ATM.IdBudgetHistory
        trn_IdBudgetKPAtm = SecurPack_ATM.IdBudgetKPAtm

        trn_IdShift_PRK = SecurPack_ATM.IdShift
        trn_IdUser = SecurPack_ATM.IdUser
        trn_IdRoute = SecurPack_ATM.IdRoute
        trn_IdATM = SecurPack_ATM.IdATM
        trn_IdLoad =  SecurPack_ATM.IdLoad
        trn_BudgetNum = SecurPack_ATM.BudgetNum

        trn_WorkType = SecurPack_ATM.WorkType
        trn_SuperBudget = SecurPack_ATM.SuperBudget
        trn_SuperQty = SecurPack_ATM.SuperQty

        trn_IdFilial = SecurPack_ATM.IdFilial
        trn_FilialCode = SecurPack_ATM.FilialCode
        trn_TabNumber = SecurPack_ATM.TabNumber
        trn_DateInkas = SecurPack_ATM.DateInkas
        trn_DtIncome = SecurPack_ATM.DtIncome
        trn_CurDate = SecurPack_ATM.CurDate

        trn_ClientName = ALLTRIM(SecurPack_ATM.ClientName)

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        f_KursValue = 1.0000
        f_Kurs = 1.0000
        trn_Kurs = 1.0000

        f_Nominal = 1.0

        trn_TotalSum_RUB_Out = SecurPack_ATM.TotalSum_RUB_Out

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        trn_TotalSum_USD_Out = SecurPack_ATM.TotalSum_USD_Out

        DO start IN KursValueLoc WITH 'USD', f_CurDate

        f_Nominal_usd = f_Nominal
        f_Kurs_usd = f_KursValue
        trn_Kurs_usd = f_KursValue

        trn_TotalSum_USD_RUB_Out = ROUND(trn_TotalSum_USD_Out * f_Nominal_usd * f_Kurs_usd, 2)

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        trn_TotalSum_EUR_Out = SecurPack_ATM.TotalSum_EUR_Out

        DO start IN KursValueLoc WITH 'EUR', f_CurDate

        f_Nominal_eur = f_Nominal
        f_Kurs_eur = f_KursValue
        trn_Kurs_eur = f_KursValue

        trn_TotalSum_Eur_RUB_Out = ROUND(trn_TotalSum_EUR_Out * f_Nominal_eur * f_Kurs_eur, 2)

* WAIT WINDOW 'IdShift_PRK = ' + ALLTRIM(IdShift_PRK) TIMEOUT 5

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
* WAIT WINDOW 'trn_TotalSum_USD_RUB_Out = ' + ALLTRIM(TRANSFORM(trn_TotalSum_USD_RUB_Out, '999 999 999.99')) + '   trn_TotalSum_EUR_RUB_Out = ' + ALLTRIM(TRANSFORM(trn_TotalSum_EUR_RUB_Out, '999 999 999 999.99')) TIMEOUT 3
* WAIT WINDOW '   trn_TotalSum_USD_RUB_Out = ' + ALLTRIM(TRANSFORM(trn_TotalSum_USD_RUB_Out, '999 999 999 999.99')) TIMEOUT 2
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        trn_IsTransit = SecurPack_ATM.IsTransit
        trn_DtRecord = SecurPack_ATM.DtRecord
        trn_Sost = 0
        trn_IdKontrol = NULL
        trn_KontrolDate = NULL
        trn_IsActive = 1

        trn_NumerATM = ALLTRIM(SecurPack_ATM.NumerATM)
        trn_NumerATMReq = ALLTRIM(SecurPack_ATM.NumerATMReq)

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        IF 'Тинькофф Банк' $ trn_ClientName
          trn_Comment = 'АТМ' + ALLTRIM(trn_NumerATM) + ' РАЗГРУЗКА ' + DTOC(TTOD(trn_DateInkas) + 1) + SPACE(1) + ALLTRIM(trn_FilialCode)  && ATMXXXXXX TYPE   ДД.MM.ГГГГ.ПИН
        ELSE
          trn_Comment = 'Разгрузка ' + LTRIM(RTRIM(trn_BudgetNum)) + ', ' + LTRIM(RTRIM(trn_ClientName)) + ', ' + 'АКТ от ' + DTOC(TTOD(trn_DateInkas) + 1)
        ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        sqlReq = 'BEGIN TRANSACTION' && Начало транзакции

        rc = RunSql_Load('', @ar)

        TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

          IF NOT EXISTS(SELECT IdSecurPack FROM Tranzit_ATM_PER WHERE IdShift IN (<<IdShift_PRK>>) AND IdSecurPack = ?trn_IdSecurPack)

            INSERT INTO Tranzit_ATM_PER
              (IdSecurPack,
              IdBudgetHistory,
              IdBudgetKPAtm,
              IdShift,
              IdUser,
              IdRoute,
              IdATM,
              IdLoad,
              BudgetNum,
              WorkType,
              SuperBudget,
              SuperQty,
              IdFilial,
              FilialCode,
              TabNumber,
              DateInkas,
              DtIncome,
              CurDate,
              TotalSum_RUB_Out,
              TotalSum_USD_Out,
              Kurs_USD,
              TotalSum_USD_RUB_Out,
              TotalSum_EUR_Out,
              Kurs_EUR,
              TotalSum_EUR_RUB_Out,
              [Comment],
              IsTransit,
              Sost,
              IdKontrol,
              KontrolDate,
              IsActive)
              VALUES
              (?trn_IdSecurPack,
              ?trn_IdBudgetHistory,
              ?trn_IdBudgetKPAtm,
              ?trn_IdShift_PRK,
              ?trn_IdUser,
              ?trn_IdRoute,
              ?trn_IdATM,
              ?trn_IdLoad,
              ?trn_BudgetNum,
              ?trn_WorkType,
              ?trn_SuperBudget,
              ?trn_SuperQty,
              ?trn_IdFilial,
              ?trn_FilialCode,
              ?trn_TabNumber,
              ?trn_DateInkas,
              ?trn_DtIncome,
              ?trn_CurDate,
              ?trn_TotalSum_RUB_Out,
              ?trn_TotalSum_USD_Out,
              ?trn_Kurs_USD,
              ?trn_TotalSum_USD_RUB_Out,
              ?trn_TotalSum_EUR_Out,
              ?trn_Kurs_EUR,
              ?trn_TotalSum_EUR_RUB_Out,
              ?trn_Comment,
              ?trn_IsTransit,
              ?trn_Sost,
              ?trn_IdKontrol,
              ?trn_KontrolDate,
              ?trn_IsActive)

           ELSE

            BEGIN

            UPDATE Tranzit_ATM_PER SET
              IdSecurPack = ?trn_IdSecurPack,
              IdBudgetHistory = ?trn_IdBudgetHistory,
              IdBudgetKPAtm = ?trn_IdBudgetKPAtm,
              IdShift = ?trn_IdShift_PRK,
              IdUser = ?trn_IdUser,
              IdRoute = ?trn_IdRoute,
              IdATM = ?trn_IdATM,
              IdLoad = ?trn_IdLoad,
              BudgetNum = ?trn_BudgetNum,
              WorkType = ?trn_WorkType,
              SuperBudget = ?trn_SuperBudget,
              SuperQty = ?trn_SuperQty,
              IdFilial = ?trn_IdFilial,
              FilialCode = ?trn_FilialCode,
              TabNumber = ?trn_TabNumber,
              DateInkas = ?trn_DateInkas,
              DtIncome = ?trn_DtIncome,
              CurDate = ?trn_CurDate,
              TotalSum_RUB_Out = ?trn_TotalSum_RUB_Out,
              TotalSum_USD_Out = ?trn_TotalSum_USD_Out,
              TotalSum_EUR_Out = ?trn_TotalSum_EUR_Out,
              Kurs_USD = ?trn_Kurs_USD,
              TotalSum_USD_RUB_Out = ROUND(TotalSum_USD_Out * ?f_Nominal_usd * ?trn_Kurs_USD, 2),
              Kurs_EUR = ?trn_Kurs_EUR,
              TotalSum_EUR_RUB_Out = ROUND(TotalSum_EUR_Out * ?f_Nominal_eur * ?trn_Kurs_EUR, 2),
              Comment = ?trn_Comment
            WHERE IdShift IN (<<IdShift_PRK>>) AND IdSecurPack = ?trn_IdSecurPack AND Sost = 0 AND IsActive = 1

            UPDATE Tranzit_ATM_PER SET
              Kurs_USD = ?trn_Kurs_USD,
              TotalSum_USD_RUB_Out = ROUND(TotalSum_USD_Out * ?f_Nominal_usd * ?trn_Kurs_USD, 2),
              Kurs_EUR = ?trn_Kurs_EUR,
              TotalSum_EUR_RUB_Out = ROUND(TotalSum_EUR_Out * ?f_Nominal_eur * ?trn_Kurs_EUR, 2)
            WHERE IdShift IN (<<IdShift_PRK>>) AND IdSecurPack = ?trn_IdSecurPack

          END

        ENDTEXT

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        rc = RunSql_Load('', @ar)

        DO CASE
          CASE rc <> 1

            sqlReq = 'ROLLBACK TRANSACTION'

            rc = RunSql_Load('', @ar)

            MESSAGEBOX('Ошибка записи проверенных данных в таблицу Tranzit_ATM' + CHR(13) + ;
              'ПИН - ' + ALLTRIM(trn_FilialCode) + '   Номер АТМ - ' + ALLTRIM(trn_BudgetNum), 16, 'Обратитесь к Программистам по тел: 1243, 1303')
            LOOP

          CASE rc = 1

            sqlReq = 'COMMIT TRANSACTION' && Завершение транзакции

            rc = RunSql_Load('', @ar)

            WAIT WINDOW 'Данные по РАЗГРУЗКЕ: обслуживаемый АТМ: ' + 'ПИН - ' + ALLTRIM(trn_FilialCode) + ;
              '   Номер АТМ - ' + ALLTRIM(trn_BudgetNum) + ' успешно записана в таблицу для проведения КОНТРОЛЯ' NOWAIT

        ENDCASE

        SELECT SecurPack_ATM
      ENDSCAN

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      WAIT CLEAR

      tim2 = SECONDS()

      result_time = ROUND(tim2 - tim1,  2)

      @ WROWS()/3, 3 SAY PADC('Импорт данных по РАЗГРУЗКЕ успешо завершен.' + '  (время = ' + ALLTRIM(TRANSFORM(result_time,  '999.99')) + ' сек.)', WCOLS())
      = INKEY(5)
      @ WROWS()/3, 3 SAY PADC('Начинаем импорт данных по ЗАГРУЗКЕ.', WCOLS())
      = INKEY(3)

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      colvo_zap_trn = 0
      colvo_zap_bdg = 0
      colvo_zap_fil = 0
      colvo_zap_bdg_fil = 0

* AND IdFilial = 20706
* AND FilialCode = '46-044'
* f_IdFilial = 20706

      TEXT TO SqlReq NOSHOW TEXTMERGE && PRETEXT 7

        SELECT DISTINCT trn.IdFilial, trn.FilialCode
        FROM Tranzit_ATM_PER AS trn (READPAST)
        INNER JOIN Filial AS fil (READPAST) ON fil.FilialCode = trn.FilialCode
        WHERE trn.DtIncome = ?f_CurDate AND fil.MestoCode = ?f_MestoCode
        ORDER BY trn.IdFilial

      ENDTEXT

      IF RunSql_Load('Scan_Tranzit_ATM', @ar) <> 1
        DEACTIVATE WINDOW poisk
        MESSAGEBOX('Ошибка выборки данных из таблицы FROM Scan_Tranzit_ATM', 16, 'Обратитесь к Программистам по тел: 1243, 1303')
      ENDIF

*      BROWSE WINDOW brows

      IF USED('Scan_Tranzit_ATM') = .T.

        SELECT Scan_Tranzit_ATM
        GO TOP

        SCAN

          f_IdFilial = Scan_Tranzit_ATM.IdFilial

*  WAIT WINDOW 'f_CurDate = ' + DTOC(f_CurDate) TIMEOUT 3
*  WAIT WINDOW 'f_IdFilial = ' + ALLTRIM(STR(f_IdFilial)) TIMEOUT 3

          TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT Id, IdShift, IdFilial, FilialCode, IdATM, IdLoad, BudgetNum, DtIncome, CurDate,
            TotalSum_RUB_In, TotalSum_USD_In, TotalSum_EUR_In, Kurs_USD, Kurs_EUR,
            IsTransit, IsDriver, IsActive, DtRecord
            FROM Tranzit_ATM_PER
            WHERE DtIncome = ?f_CurDate AND IdFilial = ?f_IdFilial
            ORDER BY IdFilial

          ENDTEXT

          rec = RunSql_Load('Tranzit_ATM_Data_IdFilial', @ar) <> 1

          IF USED('Tranzit_ATM_Data_IdFilial') = .T.

            SELECT Tranzit_ATM_Data_IdFilial

            colvo_zap_trn = RECCOUNT()
            GO TOP

            WAIT WINDOW 'Данные по ЗАГРУЗКЕ: обслуживаемый АТМ: ' + 'ПИН - ' + ALLTRIM(Tranzit_ATM_Data_IdFilial.FilialCode) + ;
              '   Номер АТМ - ' + ALLTRIM(Tranzit_ATM_Data_IdFilial.BudgetNum) + ' успешно записана в таблицу для проведения КОНТРОЛЯ' NOWAIT

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

              SELECT IdFilial, TypeValueKind, COUNT(IdFilial) AS Colvo
              FROM BudgetHistoryATM_Out
              WHERE DataLoad = ?f_CurDate AND IdFilial = ?f_IdFilial
              GROUP BY IdFilial, TypeValueKind
              ORDER BY IdFilial, TypeValueKind

            ENDTEXT

            rec = RunSql_Load('BudgetHistoryATM_Out_Data_IdFilial', @ar) <> 1

*        BROWSE WINDOW brows TITLE 'DataLoad = f_CurDate AND IdFilial = f_IdFilial'

            IF USED('BudgetHistoryATM_Out_Data_IdFilial') = .T.

              SELECT BudgetHistoryATM_Out_Data_IdFilial

              colvo_zap_bdg = BudgetHistoryATM_Out_Data_IdFilial.Colvo

            ENDIF

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*        WAIT WINDOW 'colvo_zap_trn = ' + ALLTRIM(STR(colvo_zap_trn)) TIMEOUT 3
*        WAIT WINDOW 'colvo_zap_bdg = ' + ALLTRIM(STR(colvo_zap_bdg)) TIMEOUT 3

            IF colvo_zap_trn = 0 AND colvo_zap_bdg > 0

              sqlReq = 'BEGIN TRANSACTION' && Начало транзакции

              rc = RunSql_Load('', @ar)


              TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                  SELECT bdg_atm.IdRequestATM AS IdSecurPack, bdg_atm.Id AS IdBudgetHistory, bdg_atm.IdShift, bdg_atm.IdUser, bdg_atm.IdRoute, atm.Id AS IdATM,
                  RTRIM(fil.FilialCode) + '/ATM' + RTRIM(bdg_atm.NumerATM) AS BudgetNum,
                  bdg_atm.WorkType, bdg_atm.SuperBudget, bdg_atm.SuperQty,
                  bdg_atm.IdFilial, fil.FilialCode, bdg_atm.IdLoadAtm AS IdBudgetKPAtm, bdg_atm.IdLoad,
                  0 AS TabNumber, bdg_atm.DataLoad AS DateInkas, bdg_atm.DataLoad AS DtIncome, bdg_atm.DataLoad AS CurDate,
                  cln.ClientName, bdg_atm.IdVedList, 15 AS IdSignProc, 0 AS IsTransit, bdg_atm.DataLoad, bdg_atm.DtRecord
                  FROM BudgetHistoryATM_Out AS bdg_atm (READPAST)
                  INNER JOIN Filial AS fil (READPAST) ON fil.Id = bdg_atm.IdFilial
                  INNER JOIN Client AS cln (READPAST) ON cln.Id = fil.IdClient
                  INNER JOIN BaseATM AS atm (READPAST) ON atm.IdFilial = bdg_atm.IdFilial
                  WHERE bdg_atm.DataLoad = ?f_CurDate AND bdg_atm.IdFilial = ?f_IdFilial
                  ORDER BY bdg_atm.IdShift, bdg_atm.IdFilial, bdg_atm.DataLoad

              ENDTEXT

              rec = RunSql_Load('BudgetHistoryATM_Out_Data_Fil', @ar) <> 1

              SELECT BudgetHistoryATM_Out_Data_Fil

              colvo_zap_bdg_fil = RECCOUNT()

              GO TOP

*    BROWSE WINDOW brows

              IF colvo_zap_bdg_fil > 0

                trn_IdSecurPack = BudgetHistoryATM_Out_Data_Fil.IdSecurPack
                trn_IdBudgetHistory = BudgetHistoryATM_Out_Data_Fil.IdBudgetHistory
                trn_IdBudgetKPAtm = BudgetHistoryATM_Out_Data_Fil.IdBudgetKPAtm

                trn_IdShift_PRK = BudgetHistoryATM_Out_Data_Fil.IdShift
                trn_IdUser = BudgetHistoryATM_Out_Data_Fil.IdUser
                trn_IdRoute = BudgetHistoryATM_Out_Data_Fil.IdRoute
                trn_IdATM = BudgetHistoryATM_Out_Data_Fil.IdATM
                trn_IdLoad =  BudgetHistoryATM_Out_Data_Fil.IdLoad
                trn_BudgetNum = BudgetHistoryATM_Out_Data_Fil.BudgetNum

                trn_WorkType = BudgetHistoryATM_Out_Data_Fil.WorkType
                trn_SuperBudget = BudgetHistoryATM_Out_Data_Fil.SuperBudget
                trn_SuperQty = BudgetHistoryATM_Out_Data_Fil.SuperQty

                trn_IdFilial = BudgetHistoryATM_Out_Data_Fil.IdFilial
                trn_FilialCode = BudgetHistoryATM_Out_Data_Fil.FilialCode
                trn_TabNumber = BudgetHistoryATM_Out_Data_Fil.TabNumber
                trn_DateInkas = BudgetHistoryATM_Out_Data_Fil.DateInkas
                trn_DtIncome = BudgetHistoryATM_Out_Data_Fil.DtIncome
                trn_CurDate = BudgetHistoryATM_Out_Data_Fil.CurDate

                trn_ClientName = ALLTRIM(BudgetHistoryATM_Out_Data_Fil.ClientName)

                trn_IsTransit = BudgetHistoryATM_Out_Data_Fil.IsTransit
                trn_DtRecord = BudgetHistoryATM_Out_Data_Fil.DtRecord
                trn_Sost = 0
                trn_IdKontrol = NULL
                trn_KontrolDate = NULL
                trn_IsActive = 1

              ENDIF


              TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

              IF NOT EXISTS(SELECT Id FROM Tranzit_ATM_PER WHERE DtIncome = ?f_CurDate AND IdFilial = ?f_IdFilial)

                INSERT INTO Tranzit_ATM_PER
                (IdSecurPack,
                IdBudgetHistory,
                IdBudgetKPAtm,
                IdShift,
                IdUser,
                IdRoute,
                IdATM,
                IdLoad,
                BudgetNum,
                WorkType,
                SuperBudget,
                SuperQty,
                IdFilial,
                FilialCode,
                TabNumber,
                DateInkas,
                DtIncome,
                CurDate,
                [Comment],
                IsTransit,
                Sost,
                IdKontrol,
                KontrolDate,
                IsActive)
                VALUES
                (?trn_IdSecurPack,
                ?trn_IdBudgetHistory,
                ?trn_IdBudgetKPAtm,
                ?trn_IdShift_PRK,
                ?trn_IdUser,
                ?trn_IdRoute,
                ?trn_IdATM,
                ?trn_IdLoad,
                ?trn_BudgetNum,
                ?trn_WorkType,
                ?trn_SuperBudget,
                ?trn_SuperQty,
                ?trn_IdFilial,
                ?trn_FilialCode,
                ?trn_TabNumber,
                ?trn_DateInkas,
                ?trn_DtIncome,
                ?trn_CurDate,
                ?trn_Comment,
                ?trn_IsTransit,
                ?trn_Sost,
                ?trn_IdKontrol,
                ?trn_KontrolDate,
                ?trn_IsActive)

              ENDTEXT

              rc = RunSql_Load('', @ar)

              DO CASE
                CASE rc <> 1

                  sqlReq = 'ROLLBACK TRANSACTION'

                  rc = RunSql_Load('', @ar)

                CASE rc = 1

                  sqlReq = 'COMMIT TRANSACTION' && Завершение транзакции

                  rc = RunSql_Load('', @ar)

              ENDCASE

            ENDIF

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

              SELECT Id, IdShift, IdFilial, FilialCode, BudgetNum, DtIncome, CurDate,
              TotalSum_RUB_In, TotalSum_USD_In, TotalSum_EUR_In, Kurs_USD, Kurs_EUR,
              IsTransit, IsDriver, IsActive, IdATM, IdLoad, DtRecord
              FROM Tranzit_ATM_PER
              WHERE DtIncome = ?f_CurDate AND IdFilial = ?f_IdFilial
              ORDER BY IdFilial

            ENDTEXT

            rec = RunSql_Load('Tranzit_ATM_Data_IdFilial', @ar) <> 1

*        BROWSE WINDOW brows TITLE 'DtIncome = f_CurDate AND IdFilial = f_IdFilial NoInsert'

            IF USED('Tranzit_ATM_Data_IdFilial') = .T.

              SELECT Tranzit_ATM_Data_IdFilial

              colvo_zap_fil = RECCOUNT()

            ENDIF

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            IF colvo_zap > 0

*          BROWSE WINDOW brows

              SELECT Tranzit_ATM_Data_IdFilial
              GO TOP

              SCAN

*  WAIT WINDOW 'f_IdFilial = ' + ALLTRIM(STR(f_IdFilial)) TIMEOUT 3

                TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                  SELECT IdFilial, TypeValueKind, DataLoad, IdLoad, SUM(TotalSum) AS TotalSum
                  FROM BudgetHistoryATM_Out
                  WHERE DataLoad = ?f_CurDate AND IdFilial = ?f_IdFilial
                  GROUP BY IdFilial, TypeValueKind, DataLoad, IdLoad

                ENDTEXT

                rec = RunSql_Load('BudgetHistoryATM_Out_Sum', @ar) <> 1

* BROWSE WINDOW brows TITLE 'DataLoad = f_CurDate AND IdFilial = f_IdFilial SUM_TotalSum'

                IF USED('BudgetHistoryATM_Out_Sum') = .T.

                  SELECT BudgetHistoryATM_Out_Sum

                  colvo_zap = RECCOUNT()

                  IF colvo_zap = 0

                    TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                      SELECT IdFilial, TypeValueKind, DataLoad, IdLoad, SUM(TotalSum) AS TotalSum
                      FROM BudgetHistoryATM_Out
                      WHERE IdShiftAtm = ?IdShift_ATM AND IdFilial = ?f_IdFilial
                      GROUP BY IdFilial, TypeValueKind, DataLoad, IdLoad

                    ENDTEXT

                    rec = RunSql_Load('BudgetHistoryATM_Out_Sum', @ar) <> 1

*                BROWSE WINDOW brows TITLE 'IdShiftAtm = ?IdShift_ATM AND IdFilial = f_IdFilial SUM_TotalSum'

                  ENDIF
                ENDIF


                IF USED('BudgetHistoryATM_Out_Sum') = .T.

                  SELECT BudgetHistoryATM_Out_Sum

                  colvo_zap = RECCOUNT()

                  IF colvo_zap > 0

*               BROWSE WINDOW brows TITLE 'DataLoad = f_CurDate AND IdFilial = f_IdFilial SUM_TotalSum'

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                    SELECT BudgetHistoryATM_Out_Sum
                    GO TOP

                    SCAN

                      f_TotalSum = BudgetHistoryATM_Out_Sum.TotalSum
                      f_DataLoad = BudgetHistoryATM_Out_Sum.DataLoad
                      f_IdLoad = BudgetHistoryATM_Out_Sum.IdLoad

                      DO CASE
                        CASE ALLTRIM(BudgetHistoryATM_Out_Sum.TypeValueKind) == 'RUB'

                          TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                            UPDATE Tranzit_ATM_PER SET
                            TotalSum_RUB_In = 0.00, DataLoad = ?f_DataLoad, IdLoad = ?f_IdLoad
                            WHERE DtIncome = ?f_CurDate AND IdFilial = ?f_IdFilial AND WorkType = '1' AND CHARINDEX('*', BudgetNum) > 0

                            UPDATE Tranzit_ATM_PER SET
                            TotalSum_RUB_In = ?f_TotalSum, DataLoad = ?f_DataLoad, IdLoad = ?f_IdLoad
                            WHERE DtIncome = ?f_CurDate AND IdFilial = ?f_IdFilial AND WorkType = '1' AND CHARINDEX('*', BudgetNum) = 0

                            UPDATE Tranzit_ATM_PER SET
                            TotalSum_RUB_In = 0.00, DataLoad = ?f_DataLoad, IdLoad = ?f_IdLoad
                            WHERE DtIncome = ?f_CurDate AND IdFilial = ?f_IdFilial AND WorkType = '2'

                          ENDTEXT

                          rec = RunSql_Load('', @ar) <> 1


                        CASE ALLTRIM(BudgetHistoryATM_Out_Sum.TypeValueKind) == 'USD'

                          DO start IN KursValueLoc WITH 'USD', f_CurDate

                          f_Kurs = f_KursValue

                          trn_TotalSum_USD_RUB_In = ROUND(f_TotalSum * f_Nominal * f_Kurs, 2)

                          TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                            UPDATE Tranzit_ATM_PER SET
                            TotalSum_USD_In = 0.00,
                            Kurs_USD = ?f_Kurs,
                            TotalSum_USD_RUB_In = ?trn_TotalSum_USD_RUB_In,
                            DataLoad = ?f_DataLoad, IdLoad = ?f_IdLoad
                            WHERE DtIncome = ?f_CurDate AND IdFilial = ?f_IdFilial AND WorkType = '1' AND CHARINDEX('*', BudgetNum) > 0

                            UPDATE Tranzit_ATM_PER SET
                            TotalSum_USD_In = ?f_TotalSum,
                            Kurs_USD = ?f_Kurs,
                            TotalSum_USD_RUB_In = 0.00,
                            DataLoad = ?f_DataLoad, IdLoad = ?f_IdLoad
                            WHERE DtIncome = ?f_CurDate AND IdFilial = ?f_IdFilial AND WorkType = '1' AND CHARINDEX('*', BudgetNum) = 0

                            UPDATE Tranzit_ATM_PER SET
                            TotalSum_USD_In = 0.00,
                            Kurs_USD = ?f_Kurs,
                            TotalSum_USD_RUB_In = 0.00,
                            DataLoad = ?f_DataLoad, IdLoad = ?f_IdLoad
                            WHERE DtIncome = ?f_CurDate AND IdFilial = ?f_IdFilial AND WorkType = '2'

                          ENDTEXT

                          rec = RunSql_Load('', @ar) <> 1


                        CASE ALLTRIM(BudgetHistoryATM_Out_Sum.TypeValueKind) == 'EUR'

                          DO start IN KursValueLoc WITH 'EUR', f_CurDate

                          f_Kurs = f_KursValue

                          trn_TotalSum_EUR_RUB_In = ROUND(f_TotalSum * f_Nominal * f_Kurs, 2)

                          TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                            UPDATE Tranzit_ATM_PER SET
                            TotalSum_EUR_In = 0.00,
                            Kurs_EUR = ?f_Kurs,
                            TotalSum_EUR_RUB_In = ?trn_TotalSum_EUR_RUB_In,
                            DataLoad = ?f_DataLoad, IdLoad = ?f_IdLoad
                            WHERE DtIncome = ?f_CurDate AND IdFilial = ?f_IdFilial AND WorkType = '1' AND CHARINDEX('*', BudgetNum) > 0

                            UPDATE Tranzit_ATM_PER SET
                            TotalSum_EUR_In = ?f_TotalSum,
                            Kurs_EUR = ?f_Kurs,
                            TotalSum_EUR_RUB_In = 0.00,
                            DataLoad = ?f_DataLoad, IdLoad = ?f_IdLoad
                            WHERE DtIncome = ?f_CurDate AND IdFilial = ?f_IdFilial AND WorkType = '1' AND CHARINDEX('*', BudgetNum) = 0

                            UPDATE Tranzit_ATM_PER SET
                            TotalSum_EUR_In = 0.00,
                            Kurs_EUR = ?f_Kurs,
                            TotalSum_EUR_RUB_In = 0.00,
                            DataLoad = ?f_DataLoad, IdLoad = ?f_IdLoad
                            WHERE DtIncome = ?f_CurDate AND IdFilial = ?f_IdFilial AND WorkType = '2'

                          ENDTEXT

                      ENDCASE

                      rec = RunSql_Load('', @ar) <> 1

                      SELECT BudgetHistoryATM_Out_Sum
                    ENDSCAN

                  ENDIF

                ENDIF

                SELECT Scan_Tranzit_ATM
              ENDSCAN

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            ENDIF

          ENDIF

          SELECT Scan_Tranzit_ATM
        ENDSCAN

      ENDIF

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      IF USED('Tranzit_ATM_Data_IdFilial') = .T.
        SELECT Tranzit_ATM_Data_IdFilial
        USE
      ENDIF

      IF USED('BudgetHistoryATM_Out_Data_Fil') = .T.
        SELECT BudgetHistoryATM_Out_Data_Fil
        USE
      ENDIF

      IF USED('BudgetHistoryATM_Out_Sum') = .T.
        SELECT BudgetHistoryATM_Out_Sum
        USE
      ENDIF

      IF USED('BudgetHistoryATM_Out') = .T.
        SELECT BudgetHistoryATM_Out
        USE
      ENDIF

      IF USED('Tranzit_ATM') = .T.
        SELECT Tranzit_ATM
        USE
      ENDIF

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      WAIT CLEAR

      tim2 = SECONDS()

      result_time = ROUND(tim2 - tim1,  2)

      @ WROWS()/3, 3 SAY PADC('Импорт данных завершен успешо.' + '  (время = ' + ALLTRIM(TRANSFORM(result_time,  '999.99')) + ' сек.)', WCOLS())
      = INKEY(2)
      DEACTIVATE WINDOW poisk

    ELSE
      DEACTIVATE WINDOW poisk
      =soob('Внимание! В сформированной таблице из базы SQL Server записей не обнаружено ... ')
      LOOP
    ENDIF

  ELSE
    DEACTIVATE WINDOW poisk
    =soob('Внимание! Выборка данных из базы SQL Server закончилась аварийно ... ')
    LOOP
  ENDIF

ENDIF

* ---------------------------------------------------------------- *
DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

SELECT(sel)
RETURN


***********************************************************************************************************************************************


