********************************************************************** PROCEDURE start ********************************************************

PROCEDURE start
PARAMETERS Check1_Value

sel = SELECT()

tim1 = SECONDS()

* WAIT WINDOW 'Odbcsrc = ' + ALLTRIM(Odbcsrc) TIMEOUT 5

nRepType = 0    && Тип формируемого отчёта

&& Получаем вид получаемого отчёта: 0 - стандартный, 1 - сколько прошло сумок по кассирам ПРК (из пересчитанных в КП), 3- сколько прошло сумок по кассирам КП

* --------------------------------------------------------------- *
DO FORM SearchBdgrepType TO nRepType
* --------------------------------------------------------------- *

ACTIVATE WINDOW poisk

@ WROWS()/3, 3 SAY PADC('Минуточку терпения Уважаемый Пользователь! Начата выборка данных по введенным параметрам ... ', WCOLS())

DIMENSION ar[1,2]

* -------------------------------------------------------------------------------------------------- *
SET STEP ON 
TblName = '#InfBudget'
* TblName = 'aTmp_InfBudget'

*  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*    DROP TABLE <<TblName>>

*  ENDTEXT

*  =RunSQL_Load('', @ar)

* -------------------------------------------------------------------------------------------------- *

SomNepl = '#SomNepl'
* SomNepl = 'aTmp_SomNepl'

*  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*    DROP TABLE <<SomNepl>>

*  ENDTEXT

*  =RunSQL_Load('', @ar)

* -------------------------------------------------------------------------------------------------- *

sqlReq = 'CREATE TABLE ' + TblName + ' (' + ;
  ' BudgetId [int],' + ;
  ' BudgetPRKId [int],' + ;
  ' ShiftDate [char](11),' + ;
  ' UserName  [char](50),' + ;
  ' UserNamePRK [char](50),' + ;
  ' UserNameInkas [char](50),' + ;
  ' RouteNum [int],' + ;
  ' ClientGroup [char](100),' + ;
  ' ClientGroupName [char](200),' + ;
  ' FilialCode [char](10),' + ;
  ' ClientCode [char](20),' + ;
  ' BudgetNum [char](25),' + ;
  ' SuperBudget [int],' + ;
  ' WorkType [char](1),' + ;
  ' SuperQty [int],' + ;
  ' Flg2PIN tinyint,' + ;
  ' TotalSumPRK [money],' + ;
  ' TotalSumKP [money],' + ;
  ' DtCalcBegin [datetime],' + ;
  ' DtCalcEnd [datetime],' + ;
  ' DtInventoryKP [datetime],' + ;
  ' Defect [char] (30),' + ;
  ' Delta [money], ' + ;
  ' Somnit [money],' + ;
  ' Neplat [money],' + ;
  ' NumCard [char](15),' + ;
  ' DtPacketCBS [datetime],' + ;
  ' Wt_Flag [int]) ON [PRIMARY]'

* WAIT WINDOW 'CREATE TABLE TblName' TIMEOUT 2

=RunSQL_Load('', @ar)

* Кроме супер сумок и ошибочных сумок

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

* WAIT WINDOW 'OdbcSrc = ' + ALLTRIM(OdbcSrc) TIMEOUT 3

Type_OdbcSrc = PROPER(LOWER(ALLTRIM(OdbcSrc)))

* WAIT WINDOW 'Type_OdbcSrc = ' + ALLTRIM(Type_OdbcSrc) TIMEOUT 3

* WAIT WINDOW 'God_Date_Begin = ' + ALLTRIM(STR(YEAR(DATE()) - YEAR(f_Date_Begin))) TIMEOUT 3
* WAIT WINDOW 'God_Date_End = ' + ALLTRIM(STR(YEAR(DATE()) - YEAR(f_Date_End))) TIMEOUT 3

DO CASE
  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 0 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 0 AND ;
      (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ;
      ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko' OR ALLTRIM(Type_OdbcSrc) == 'Recount_yar')

* WAIT WINDOW 'FROM BudgetKP' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<TblName>>
      (BudgetId, BudgetPRKId, ShiftDate, UserName, ClientGroup, ClientGroupName, FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, Flg2PIN, TotalSumPRK, TotalSumKP, DtCalcBegin, DtCalcEnd, Delta, Somnit, Neplat, Defect, NumCard, DtPacketCBS, Wt_flag)
      SELECT
      BudgetKp.Id AS Id, BudgetKp.IdBudgetHistoryPRK, CONVERT(Char(10), Shift.ReportDate, 104) AS ShiftDate, UserName, ClientGroup.ClientGroupName, Filial.FilialName, Filial.FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, BudgetKP.Flg2PIN, TotalSumPRK, TotalSumKP, DtToCbs, DtFromCbs, (TotalSumKP - TotalSumPRK), 0.00, 0.00, '', NumCard, DtPacketCBS, Wt_flag
      FROM BudgetKP (READPAST)
      LEFT OUTER JOIN [User] (READPAST) ON BudgetKP.IdUser = [User].Id
      INNER JOIN Shift (READPAST) ON BudgetKP.IdShift = Shift.Id
      INNER JOIN Filial (READPAST) ON BudgetKP.IdFilial = Filial.Id
      INNER JOIN Client (READPAST) ON Filial.IdClient = Client.Id
      INNER JOIN ClientGroup (READPAST) ON Client.IdClientGroup = ClientGroup.Id
      WHERE CAST(Shift.ReportDate AS Date) >= ?f_Date_Begin AND CAST(Shift.ReportDate AS Date) <= ?f_Date_End AND Shift.IdDepartment = ?f_IdDepartmentKP

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 1 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 0 AND ;
      (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ;
      ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko' OR ALLTRIM(Type_OdbcSrc) == 'Recount_yar')

* WAIT WINDOW 'FROM BudgetKP' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<TblName>>
      (BudgetId, BudgetPRKId, ShiftDate, UserName, ClientGroup, ClientGroupName, FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, Flg2PIN, TotalSumPRK, TotalSumKP, DtCalcBegin, DtCalcEnd, Delta, Somnit, Neplat, Defect, NumCard, DtPacketCBS, Wt_flag)
      SELECT
      BudgetKp.Id AS Id, BudgetKp.IdBudgetHistoryPRK, CONVERT(char(11), CAST(Shift.ReportDate AS Date), 104) AS ShiftDate, UserName, ClientGroup.ClientGroupName, Filial.FilialName, Filial.FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, BudgetKP.Flg2PIN, TotalSumPRK, TotalSumKP, DtToCbs, DtFromCbs, (TotalSumKP - TotalSumPRK), 0.00, 0.00, '', NumCard, DtPacketCBS, Wt_flag
      FROM BudgetKP (READPAST)
      LEFT OUTER JOIN [User] (READPAST) ON BudgetKP.IdUser = [User].Id
      INNER JOIN Shift (READPAST) ON BudgetKP.IdShift = Shift.Id
      INNER JOIN Filial (READPAST) ON BudgetKP.IdFilial = Filial.Id
      INNER JOIN Client (READPAST) ON Filial.IdClient = Client.Id
      INNER JOIN ClientGroup (READPAST) ON Client.IdClientGroup = ClientGroup.Id
      WHERE CAST(Shift.ReportDate AS Date) >= ?f_Date_Begin AND CAST(Shift.ReportDate AS Date) <= ?f_Date_End AND Shift.IdDepartment = ?f_IdDepartmentKP

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 1 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 1 AND ;
      (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko')

* WAIT WINDOW 'FROM BudgetKP' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<TblName>>
      (BudgetId, BudgetPRKId, ShiftDate, UserName, ClientGroup, ClientGroupName, FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, Flg2PIN, TotalSumPRK, TotalSumKP, DtCalcBegin, DtCalcEnd, Delta, Somnit, Neplat, Defect, NumCard, DtPacketCBS, Wt_flag)
      SELECT
      BudgetKp.Id AS Id, BudgetKp.IdBudgetHistoryPRK, CONVERT(char(11), CAST(Shift.ReportDate AS Date), 104) AS ShiftDate, UserName, ClientGroup.ClientGroupName, Filial.FilialName, Filial.FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, BudgetKP.Flg2PIN, TotalSumPRK, TotalSumKP, DtToCbs, DtFromCbs, (TotalSumKP - TotalSumPRK), 0.00, 0.00, '', NumCard, DtPacketCBS, Wt_flag
      FROM BudgetKP (READPAST)
      LEFT OUTER JOIN [User] (READPAST) ON BudgetKP.IdUser = [User].Id
      INNER JOIN Shift (READPAST) ON BudgetKP.IdShift = Shift.Id
      INNER JOIN Filial (READPAST) ON BudgetKP.IdFilial = Filial.Id
      INNER JOIN Client (READPAST) ON Filial.IdClient = Client.Id
      INNER JOIN ClientGroup (READPAST) ON Client.IdClientGroup = ClientGroup.Id
      WHERE CAST(Shift.ReportDate AS Date) >= ?f_Date_Begin AND CAST(Shift.ReportDate AS Date) <= ?f_Date_End

    ENDTEXT

* ======================================================================================================================================================================================================== *

  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 2 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 2 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

* WAIT WINDOW 'FROM BudgetKP_2016' + ' - ' + ALLTRIM(Type_OdbcSrc) TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<TblName>>
      (BudgetId, BudgetPRKId, ShiftDate, UserName, ClientGroup, ClientGroupName, FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, Flg2PIN, TotalSumPRK, TotalSumKP, DtCalcBegin, DtCalcEnd, Delta, Somnit, Neplat, Defect, NumCard, DtPacketCBS, Wt_flag)
      SELECT
      BudgetKp.Id AS Id, BudgetKp.IdBudgetHistoryPRK, CONVERT(char(11), CAST(Shift.ReportDate AS Date), 104) AS ShiftDate, UserName, ClientGroup.ClientGroupName, Filial.FilialName, Filial.FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, BudgetKP.Flg2PIN, TotalSumPRK, TotalSumKP, DtToCbs, DtFromCbs, (TotalSumKP - TotalSumPRK), 0.00, 0.00, '', NumCard, DtPacketCBS, Wt_flag
      FROM BudgetKP_2016 AS BudgetKP (READPAST)
      LEFT OUTER JOIN [User] (READPAST) ON BudgetKP.IdUser = [User].Id
      INNER JOIN Shift (READPAST) ON BudgetKP.IdShift = Shift.Id
      INNER JOIN Filial (READPAST) ON BudgetKP.IdFilial = Filial.Id
      INNER JOIN Client (READPAST) ON Filial.IdClient = Client.Id
      INNER JOIN ClientGroup (READPAST) ON Client.IdClientGroup = ClientGroup.Id
      WHERE CAST(Shift.ReportDate AS Date) >= ?f_Date_Begin AND CAST(Shift.ReportDate AS Date) <= ?f_Date_End

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 3 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 3 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

* WAIT WINDOW 'FROM BudgetKP_2015' + ' - ' + ALLTRIM(Type_OdbcSrc) TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<TblName>>
      (BudgetId, BudgetPRKId, ShiftDate, UserName, ClientGroup, ClientGroupName, FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, Flg2PIN, TotalSumPRK, TotalSumKP, DtCalcBegin, DtCalcEnd, Delta, Somnit, Neplat, Defect, NumCard, DtPacketCBS, Wt_flag)
      SELECT
      BudgetKp.Id AS Id, BudgetKp.IdBudgetHistoryPRK, CONVERT(char(11), CAST(Shift.ReportDate AS Date), 104) AS ShiftDate, UserName, ClientGroup.ClientGroupName, Filial.FilialName, Filial.FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, BudgetKP.Flg2PIN, TotalSumPRK, TotalSumKP, DtToCbs, DtFromCbs, (TotalSumKP - TotalSumPRK), 0.00, 0.00, '', NumCard, DtPacketCBS, Wt_flag
      FROM BudgetKP_2015 AS BudgetKP (READPAST)
      LEFT OUTER JOIN [User] (READPAST) ON BudgetKP.IdUser = [User].Id
      INNER JOIN Shift (READPAST) ON BudgetKP.IdShift = Shift.Id
      INNER JOIN Filial (READPAST) ON BudgetKP.IdFilial = Filial.Id
      INNER JOIN Client (READPAST) ON Filial.IdClient = Client.Id
      INNER JOIN ClientGroup (READPAST) ON Client.IdClientGroup = ClientGroup.Id
      WHERE CAST(Shift.ReportDate AS Date) >= ?f_Date_Begin AND CAST(Shift.ReportDate AS Date) <= ?f_Date_End

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 4 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 4 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

* WAIT WINDOW 'FROM BudgetKP_2014' + ' - ' + ALLTRIM(Type_OdbcSrc) TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<TblName>>
      (BudgetId, BudgetPRKId, ShiftDate, UserName, ClientGroup, ClientGroupName, FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, Flg2PIN, TotalSumPRK, TotalSumKP, DtCalcBegin, DtCalcEnd, Delta, Somnit, Neplat, Defect, NumCard, DtPacketCBS, Wt_flag)
      SELECT
      BudgetKp.Id AS Id, BudgetKp.IdBudgetHistoryPRK, CONVERT(char(11), CAST(Shift.ReportDate AS Date), 104) AS ShiftDate, UserName, ClientGroup.ClientGroupName, Filial.FilialName, Filial.FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, BudgetKP.Flg2PIN, TotalSumPRK, TotalSumKP, DtToCbs, DtFromCbs, (TotalSumKP - TotalSumPRK), 0.00, 0.00, '', NumCard, DtPacketCBS, Wt_flag
      FROM BudgetKP_2014 AS BudgetKP (READPAST)
      LEFT OUTER JOIN [User] (READPAST) ON BudgetKP.IdUser = [User].Id
      INNER JOIN Shift (READPAST) ON BudgetKP.IdShift = Shift.Id
      INNER JOIN Filial (READPAST) ON BudgetKP.IdFilial = Filial.Id
      INNER JOIN Client (READPAST) ON Filial.IdClient = Client.Id
      INNER JOIN ClientGroup (READPAST) ON Client.IdClientGroup = ClientGroup.Id
      WHERE CAST(Shift.ReportDate AS Date) >= ?f_Date_Begin AND CAST(Shift.ReportDate AS Date) <= ?f_Date_End

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 5 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 5 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

* WAIT WINDOW 'FROM BudgetKP_2013' + ' - ' + ALLTRIM(Type_OdbcSrc) TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<TblName>>
      (BudgetId, BudgetPRKId, ShiftDate, UserName, ClientGroup, ClientGroupName, FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, Flg2PIN, TotalSumPRK, TotalSumKP, DtCalcBegin, DtCalcEnd, Delta, Somnit, Neplat, Defect, NumCard, DtPacketCBS, Wt_flag)
      SELECT
      BudgetKp.Id AS Id, BudgetKp.IdBudgetHistoryPRK, CONVERT(char(11), CAST(Shift.ReportDate AS Date), 104) AS ShiftDate, UserName, ClientGroup.ClientGroupName, Filial.FilialName, Filial.FilialCode, ClientCode, BudgetNum, SuperBudget, WorkType, SuperQty, BudgetKP.Flg2PIN, TotalSumPRK, TotalSumKP, DtToCbs, DtFromCbs, (TotalSumKP - TotalSumPRK), 0.00, 0.00, '', NumCard, DtPacketCBS, Wt_flag
      FROM BudgetKP_2013 AS BudgetKP (READPAST)
      LEFT OUTER JOIN [User] (READPAST) ON BudgetKP.IdUser = [User].Id
      INNER JOIN Shift (READPAST) ON BudgetKP.IdShift = Shift.Id
      INNER JOIN Filial (READPAST) ON BudgetKP.IdFilial = Filial.Id
      INNER JOIN Client (READPAST) ON Filial.IdClient = Client.Id
      INNER JOIN ClientGroup (READPAST) ON Client.IdClientGroup = ClientGroup.Id
      WHERE CAST(Shift.ReportDate AS Date) >= ?f_Date_Begin AND CAST(Shift.ReportDate AS Date) <= ?f_Date_End

    ENDTEXT

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

* WAIT WINDOW 'INSERT INTO TblName' TIMEOUT 2

* --------------------------------------------------------------------------------------------------------------------------------------------------------------- *

StrOrder = ' ORDER BY ShiftDate, SuperBudget, WorkType, SuperQty, BudgetNum'

* --------------------------------------------------------------------------------------------------------------------------------------------------------------- *

sqlReq1 = ''

IF EMPTY(f_KassirKP) = .F.

  sqlReq1 = sqlReq1 + " AND UserName LIKE '%" + ALLTRIM(f_KassirKP) + "%'"

ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF EMPTY(f_ClientGroup) = .F.

  sqlReq1 = sqlReq1 + " AND ClientGroupName LIKE '%" + ALLTRIM(f_ClientGroup) + "%'"

ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF EMPTY(f_FilialCode) = .F.

  sqlReq1 = sqlReq1 + ' AND FilialCode = ' + "'" + ALLTRIM(f_FilialCode) + "'"

ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF EMPTY(f_Sum1) = .F.

  sqlReq1 = sqlReq1 + ' AND TotalSumKP >=  ' + ALLTRIM(STR20(f_Sum1, 1))

ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF EMPTY(f_Sum2) = .F.

  sqlReq1 = sqlReq1 + ' AND TotalSumKP <=  ' + ALLTRIM(STR20(f_Sum2, 1))

ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF EMPTY(sqlReq1) = .F.

  sqlReq = sqlReq + sqlReq1

ENDIF

sqlReq = sqlReq + StrOrder

* WAIT WINDOW 'Выполнение - INSERT INTO TblName' TIMEOUT 2

rc = RunSQL_Load('', @ar)

* WAIT WINDOW 'rc = ' + ALLTRIM(STR(rc)) TIMEOUT 3

IF rc <> 1
  MESSAGEBOX('Ошибка при запросе к базе данных', 16, 'Поиск информации по сумкам')
  RETURN
ENDIF

IF ar[1,2] = 0
  MESSAGEBOX('Данных по выбранному критерию не найдено', 16, 'Поиск информации по сумкам')
  RETURN
ENDIF

* WAIT WINDOW 'Выполнение - INSERT INTO TblName' TIMEOUT 2

m.allbg = ar[1,2]


IF (nRepType = 0) AND (ar[1,2] > 65535)  && Количество строк в Excel , проверка только для оригинального отчёта, остальные будут короче

  MESSAGEBOX('Слишком большой объем данных, количество записей = ' + STR20(ar[1,2], 1), 16, 'Операция прекращена')
  RETURN

ENDIF

* Очистить сумму ПРК у вложенных сумок

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  UPDATE <<TblName>> SET
  TotalSumPRK = 0
  WHERE WorkType = '2' AND Wt_flag <> 1

ENDTEXT

IF RunSQL_Load('', @ar) <> 1
  MESSAGEBOX('Ошибка при запросе к базе данных 222', 16, 'Поиск информации по сумкам')
  RETURN
ENDIF

* WAIT WINDOW 'Очистить сумму КП у супер-сумок' TIMEOUT 2

* Очистить сумму КП у супер-сумок

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  UPDATE <<TblName>> SET
  TotalSumKP = 0
  WHERE WorkType = '1'

ENDTEXT

IF RunSQL_Load('', @ar) <> 1
  MESSAGEBOX('Ошибка при запросе к базе данных 333', 16, 'Поиск информации по сумкам')
  RETURN
ENDIF

* WAIT WINDOW 'Очистить сумму избыток-недостачу у супер-сумок Леруа' TIMEOUT 2

* Очистить сумму избыток-недостачу у супер-сумок Леруа

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  UPDATE <<TblName>> SET
  Delta = 0, TotalSumPRK = 0
  WHERE WorkType = '1' AND Wt_flag = '1'

ENDTEXT

IF RunSQL_Load('', @ar) <> 1
  MESSAGEBOX('Ошибка при запросе к базе данных 444', 16, 'Поиск информации по сумкам')
  RETURN
ENDIF

* WAIT WINDOW 'Заполнение полей по ПРК' TIMEOUT 2

* Заполнение полей по ПРК

* ========================================================================================================================================================================================================== *

DO CASE
  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 0 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 0 AND ;
      (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko')

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      UPDATE <<TblName>> SET
      UserNamePRK = [user].UserName, RouteNum = RouteNumber, UserNameInkas = FullName
      FROM BudgetHistoryPRK (READPAST)
      INNER JOIN [User] (READPAST) ON BudgetHistoryPRK.IdUser = [user].Id
      INNER JOIN [RouteHistory] (READPAST) ON BudgetHistoryPRK.IdRouteHistory = RouteHistory.Id
      INNER JOIN CashMessenger (READPAST) ON IdCashMessenger = CashMessenger.Id
      INNER JOIN [Route] (READPAST) ON RouteHistory.IdRoute = Route.Id
      WHERE BudgetPRKId = BudgetHistoryPRK.Id

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 1 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 0 AND ;
      (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko')

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      UPDATE <<TblName>> SET
      UserNamePRK = [user].UserName, RouteNum = RouteNumber, UserNameInkas = FullName
      FROM BudgetHistoryPRK (READPAST)
      INNER JOIN [User] (READPAST) ON BudgetHistoryPRK.IdUser = [user].Id
      INNER JOIN [RouteHistory] (READPAST) ON BudgetHistoryPRK.IdRouteHistory = RouteHistory.Id
      INNER JOIN CashMessenger (READPAST) ON IdCashMessenger = CashMessenger.Id
      INNER JOIN [Route] (READPAST) ON RouteHistory.IdRoute = Route.Id
      WHERE BudgetPRKId = BudgetHistoryPRK.Id

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 1 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 1 AND ;
      (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko')

*    WAIT WINDOW '2 - Recount - FROM BudgetHistoryPRK_2017' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      UPDATE <<TblName>> SET
      UserNamePRK = [user].UserName, RouteNum = RouteNumber, UserNameInkas = FullName
      FROM BudgetHistoryPRK (READPAST)
      INNER JOIN [User] (READPAST) ON BudgetHistoryPRK.IdUser = [user].Id
      INNER JOIN [RouteHistory] (READPAST) ON BudgetHistoryPRK.IdRouteHistory = RouteHistory.Id
      INNER JOIN CashMessenger (READPAST) ON IdCashMessenger = CashMessenger.Id
      INNER JOIN [Route] (READPAST) ON RouteHistory.IdRoute = Route.Id
      WHERE BudgetPRKId = BudgetHistoryPRK.Id

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 2 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 2 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

*    WAIT WINDOW '2 - Recount - FROM BudgetHistoryPRK_2016' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      UPDATE <<TblName>> SET
      UserNamePRK = [user].UserName, RouteNum = RouteNumber, UserNameInkas = FullName
      FROM BudgetHistoryPRK_2016 AS BudgetHistoryPRK (READPAST)
      INNER JOIN [User] (READPAST) ON BudgetHistoryPRK.IdUser = [user].Id
      INNER JOIN [RouteHistory] (READPAST) ON BudgetHistoryPRK.IdRouteHistory = RouteHistory.Id
      INNER JOIN CashMessenger (READPAST) ON IdCashMessenger = CashMessenger.Id
      INNER JOIN [Route] (READPAST) ON RouteHistory.IdRoute = Route.Id
      WHERE BudgetPRKId = BudgetHistoryPRK.Id

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 3 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 3 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

*    WAIT WINDOW '2 - Recount - FROM BudgetHistoryPRK_2015' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      UPDATE <<TblName>> SET
      UserNamePRK = [user].UserName, RouteNum = RouteNumber, UserNameInkas = FullName
      FROM BudgetHistoryPRK_2015 AS BudgetHistoryPRK (READPAST)
      INNER JOIN [User] (READPAST) ON BudgetHistoryPRK.IdUser = [user].Id
      INNER JOIN [RouteHistory] (READPAST) ON BudgetHistoryPRK.IdRouteHistory = RouteHistory.Id
      INNER JOIN CashMessenger (READPAST) ON IdCashMessenger = CashMessenger.Id
      INNER JOIN [Route] (READPAST) ON RouteHistory.IdRoute = Route.Id
      WHERE BudgetPRKId = BudgetHistoryPRK.Id

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 4 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 4 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

*    WAIT WINDOW '2 - Recount - FROM BudgetHistoryPRK_2014' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      UPDATE <<TblName>> SET
      UserNamePRK = [user].UserName, RouteNum = RouteNumber, UserNameInkas = FullName
      FROM BudgetHistoryPRK_2014 AS BudgetHistoryPRK (READPAST)
      INNER JOIN [User] (READPAST) ON BudgetHistoryPRK.IdUser = [user].Id
      INNER JOIN [RouteHistory] (READPAST) ON BudgetHistoryPRK.IdRouteHistory = RouteHistory.Id
      INNER JOIN CashMessenger (READPAST) ON IdCashMessenger = CashMessenger.Id
      INNER JOIN [Route] (READPAST) ON RouteHistory.IdRoute = Route.Id
      WHERE BudgetPRKId = BudgetHistoryPRK.Id

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 5 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 5 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

*    WAIT WINDOW '2 - Recount - FROM BudgetHistoryPRK_2013' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      UPDATE <<TblName>> SET
      UserNamePRK = [user].UserName, RouteNum = RouteNumber, UserNameInkas = FullName
      FROM BudgetHistoryPRK_2013 AS BudgetHistoryPRK (READPAST)
      INNER JOIN [User] (READPAST) ON BudgetHistoryPRK.IdUser = [user].Id
      INNER JOIN [RouteHistory] (READPAST) ON BudgetHistoryPRK.IdRouteHistory = RouteHistory.Id
      INNER JOIN CashMessenger (READPAST) ON IdCashMessenger = CashMessenger.Id
      INNER JOIN [Route] (READPAST) ON RouteHistory.IdRoute = Route.Id
      WHERE BudgetPRKId = BudgetHistoryPRK.Id

    ENDTEXT


ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

IF RunSQL_Load('', @ar) <> 1
  MESSAGEBOX('Ошибка при запросе к базе данных', 16, 'Поиск информации по сумкам')
  RETURN
ENDIF

* WAIT WINDOW 'Выполнение - UPDATE TblName' TIMEOUT 2

* WAIT WINDOW 'Поиск дефектов' TIMEOUT 2
* Поиск дефектов

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT BudgetId, IdBudgetDefect
  FROM <<TblName>> (READPAST), BudgetHistoryPRKDefectLink (READPAST)
  WHERE IdBudgetHistory = BudgetPRKId
  ORDER BY BudgetId

ENDTEXT

IF RunSQL_Load('DefectLink', @ar) <> 1
  MESSAGEBOX('Ошибка при запросе к базе данных 666', 16, 'Поиск информации по сумкам')
  RETURN
ENDIF

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  CREATE TABLE <<SomNepl>> (IdBudgetKP [int], State [tinyint], Summ [money]) ON [PRIMARY]

ENDTEXT

= RunSQL_Load('', @ar)

* WAIT WINDOW 'Сомнительные и Неплатежные' TIMEOUT 2

* Сомнительные и Неплатежные

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO CASE
  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 0 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 0 AND ;
      (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ;
      ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko' OR ALLTRIM(Type_OdbcSrc) == 'Recount_yar')

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<SomNepl>>
      (IdBudgetKP, State, Summ)
      SELECT
      IdBudgetKP, State, SUM(FaceValue) AS Summ
      FROM MistakeLogKP (READPAST)
      INNER JOIN <<TblName>> (READPAST) ON BudgetId = IdBudgetKP
      INNER JOIN Valuable (READPAST) ON MistakeLogKP.Idvaluable = Valuable.Id
      INNER JOIN FaceValue (READPAST) ON Valuable.IdFaceValue = FaceValue.Id
      WHERE State < 3
      GROUP BY State, MistakeLogKP.IdBudgetKP
      ORDER BY State

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 1 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 0 AND ;
      (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ;
      ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko' OR ALLTRIM(Type_OdbcSrc) == 'Recount_yar')

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<SomNepl>>
      (IdBudgetKP, State, Summ)
      SELECT
      IdBudgetKP, State, SUM(FaceValue) AS Summ
      FROM MistakeLogKP (READPAST)
      INNER JOIN <<TblName>> (READPAST) ON BudgetId = IdBudgetKP
      INNER JOIN Valuable (READPAST) ON MistakeLogKP.Idvaluable = Valuable.Id
      INNER JOIN FaceValue (READPAST) ON Valuable.IdFaceValue = FaceValue.Id
      WHERE State < 3
      GROUP BY State, MistakeLogKP.IdBudgetKP
      ORDER BY State

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 1 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 1 AND ;
      (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR  ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko')

*    WAIT WINDOW '2 - Recount - FROM MistakeLogKP_2016' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<SomNepl>>
      (IdBudgetKP, State, Summ)
      SELECT
      IdBudgetKP, State, SUM(FaceValue) AS Summ
      FROM MistakeLogKP (READPAST)
      INNER JOIN <<TblName>> (READPAST) ON BudgetId = IdBudgetKP
      INNER JOIN Valuable (READPAST) ON MistakeLogKP.Idvaluable = Valuable.Id
      INNER JOIN FaceValue (READPAST) ON Valuable.IdFaceValue = FaceValue.Id
      WHERE State < 3
      GROUP BY State, MistakeLogKP.IdBudgetKP
      ORDER BY State

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 2 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 2 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

*    WAIT WINDOW '2 - Recount - FROM MistakeLogKP_2016' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<SomNepl>>
      (IdBudgetKP, State, Summ)
      SELECT
      IdBudgetKP, State, SUM(FaceValue) AS Summ
      FROM MistakeLogKP_2016 AS MistakeLogKP (READPAST)
      INNER JOIN <<TblName>> (READPAST) ON BudgetId = IdBudgetKP
      INNER JOIN Valuable (READPAST) ON MistakeLogKP.Idvaluable = Valuable.Id
      INNER JOIN FaceValue (READPAST) ON Valuable.IdFaceValue = FaceValue.Id
      WHERE State < 3
      GROUP BY State, MistakeLogKP.IdBudgetKP
      ORDER BY State

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 3 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 3 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

*    WAIT WINDOW '2 - Recount - FROM MistakeLogKP_2015' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<SomNepl>>
      (IdBudgetKP, State, Summ)
      SELECT
      IdBudgetKP, State, SUM(FaceValue) AS Summ
      FROM MistakeLogKP_2015 AS MistakeLogKP (READPAST)
      INNER JOIN <<TblName>> (READPAST) ON BudgetId = IdBudgetKP
      INNER JOIN Valuable (READPAST) ON MistakeLogKP.Idvaluable = Valuable.Id
      INNER JOIN FaceValue (READPAST) ON Valuable.IdFaceValue = FaceValue.Id
      WHERE State < 3
      GROUP BY State, MistakeLogKP.IdBudgetKP
      ORDER BY State

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 4 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 4 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

*    WAIT WINDOW '2 - Recount - FROM MistakeLogKP_2014' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<SomNepl>>
      (IdBudgetKP, State, Summ)
      SELECT
      IdBudgetKP, State, SUM(FaceValue) AS Summ
      FROM MistakeLogKP_2014 AS MistakeLogKP (READPAST)
      INNER JOIN <<TblName>> (READPAST) ON BudgetId = IdBudgetKP
      INNER JOIN Valuable (READPAST) ON MistakeLogKP.Idvaluable = Valuable.Id
      INNER JOIN FaceValue (READPAST) ON Valuable.IdFaceValue = FaceValue.Id
      WHERE State < 3
      GROUP BY State, MistakeLogKP.IdBudgetKP
      ORDER BY State

    ENDTEXT


  CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 5 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 5 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

*    WAIT WINDOW '2 - Recount - FROM MistakeLogKP_2013' TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO <<SomNepl>>
      (IdBudgetKP, State, Summ)
      SELECT
      IdBudgetKP, State, SUM(FaceValue) AS Summ
      FROM MistakeLogKP_2013 AS MistakeLogKP (READPAST)
      INNER JOIN <<TblName>> (READPAST) ON BudgetId = IdBudgetKP
      INNER JOIN Valuable (READPAST) ON MistakeLogKP.Idvaluable = Valuable.Id
      INNER JOIN FaceValue (READPAST) ON Valuable.IdFaceValue = FaceValue.Id
      WHERE State < 3
      GROUP BY State, MistakeLogKP.IdBudgetKP
      ORDER BY State

    ENDTEXT


ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

IF RunSQL_Load('', @ar) <> 1
  RETURN -1
ENDIF

* WAIT WINDOW 'Выполнение - INSERT INTO <<SomNepl>>' TIMEOUT 2

* WAIT WINDOW 'Сомнительные' TIMEOUT 2

* Сомнительные
* Неплатежные

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  UPDATE <<TblName>> SET
  Somnit = <<SomNepl>>.Summ
  FROM <<TblName>>
  INNER JOIN <<SomNepl>> ON <<SomNepl>>.IdBudgetKP=<<TblName>>.BudgetId
  WHERE <<SomNepl>>.State = 1

  UPDATE <<TblName>> SET
  Neplat = Summ
  FROM <<TblName>>
  INNER JOIN <<SomNepl>> ON <<SomNepl>>.IdBudgetKP=<<TblName>>.BudgetId
  WHERE  <<SomNepl>>.State = 2

  DECLARE @Somnit Numeric(15, 2)
  SET @Somnit = 0.00

  SELECT @Somnit = SUM(SomNepl.Summ)
  FROM <<TblName>> AS TblName
  INNER JOIN <<SomNepl>> AS SomNepl ON SomNepl.IdBudgetKP = TblName.BudgetId
  WHERE TblName.WorkType = 2

  UPDATE <<TblName>> SET
  Delta =  -Somn.Somnit        -- Delta==0 @ this moment   IIF(Delta < 0, (Delta + @Somnit), (Delta - @Somnit))
  , Somnit = Somn.Somnit
  FROM <<TblName>> as TblName
  INNER JOIN (select TblNameInner.SuperBudget,SUM(SomNepl.Summ) as Somnit
                  from <<SomNepl>> AS SomNepl 
                  inner join <<TblName>> as TblNameInner on TblNameInner.BudgetId =SomNepl.IdBudgetKP 
                  where TblNameInner.WorkType='2'
                  group by TblNameInner.SuperBudget
              ) as Somn on somn.superBudget=TblName.BudgetId


ENDTEXT

IF RunSQL_Load('', @ar) <> 1
  RETURN -1
ENDIF

* ---------------------------------------------------------------------------------------- *

*  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*    SELECT SomNepl.Summ, TblName.*
*    FROM <<TblName>> AS TblName
*    INNER JOIN <<SomNepl>> AS SomNepl ON SomNepl.IdBudgetKP = TblName.BudgetId
*    WHERE TblName.WorkType = 2

*  ENDTEXT

*  IF RunSQL_Load('TblName', @ar) <> 1
*    RETURN -1
*  ENDIF

*  BROWSE WINDOW brows



* ---------------------------------------------------------------------------------------- *

* WAIT WINDOW 'Неплатежные' TIMEOUT 2
* Неплатежные

*  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*    UPDATE <<TblName>> SET
*    Neplat = Summ
*    FROM <<SomNepl>>
*    WHERE BudgetId = IdBudgetKP AND State = 2

*  ENDTEXT

*  IF RunSQL_Load('', @ar) <> 1
*    RETURN -1
*  ENDIF

* ---------------------------------------------------------------------------------------- *

*  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*    SELECT *
*    FROM <<SomNepl>>

*  ENDTEXT

*  IF RunSQL_Load('SomNepl', @ar) <> 1
*    RETURN -1
*  ENDIF

*  BROWSE WINDOW brows

*  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*    SELECT *
*    FROM <<TblName>>

*  ENDTEXT

*  IF RunSQL_Load('TblName', @ar) <> 1
*    RETURN -1
*  ENDIF

*  BROWSE WINDOW brows

* ---------------------------------------------------------------------------------------- *

*  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*    DROP TABLE <<SomNepl>>

*  ENDTEXT

*  =RunSQL_Load('', @ar)

* ----------------------------------------------------------------------------------------- *

* WAIT WINDOW 'Создаем локальный курсор с развернутой покупюркой для каждой сумки' TIMEOUT 2

* Основная выборка - к ней потом будут , при необходимости подцепляться дополнительные, имхо Макаров КВ.
* Создаем локальный курсор с развернутой покупюркой для каждой сумки

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

DO CASE
  CASE nRepType == 4	&& Итоговый отчет по суммам ПРК (маршруты)

  	TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT Convert(DateTime, ShiftDate, 104) AS _ShiftDate,STR(RouteNum,5) as Route, SUM(TotalSumPRK) AS TotalSumPRK
        FROM <<TblName>>
        WHERE WorkType = '0' OR (WorkType = '1' AND Flg2Pin = 0)
        GROUP BY shiftdate,RouteNum
        ORDER BY 2

  	ENDTEXT
  	
    IF RunSQL_Load('AllResultp', @ar) <> 1  &&выбрали в предварительную таблицу
      RETURN -1
    ENDIF

    SELECT DTOC(_shiftdate) AS ShiftDate, Route, TotalSumPRK ;
      FROM AllResultp ;
      ORDER BY _ShiftDate, Route ;
      INTO CURSOR AllResult
      
  CASE nRepType = 3  && итоговый отчёт по кассирам КП

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT Convert(DateTime, ShiftDate, 104) AS _ShiftDate, UserName, COUNT(*) AS BudgCount, Convert(varchar(15), 'Стандартные') AS BudgType, 1 AS SortOrd
        FROM <<TblName>>
        WHERE WorkType = '0' AND Flg2Pin = 0
        GROUP BY shiftdate,UserName
      UNION
      SELECT Convert(DateTime, ShiftDate, 104) AS _ShiftDate, UserName, COUNT(*) AS BudgCount, Convert(varchar(15), 'Многопиновые') AS BudgType, 2 AS SortOrd
        FROM <<TblName>>
        WHERE WorkType = '0' AND Flg2Pin = 1
        GROUP BY shiftdate,UserName
      UNION
      SELECT Convert(DateTime, ShiftDate, 104) AS _ShiftDate, UserName, COUNT(*) AS BudgCount, Convert(varchar(15), 'В суперсумках') AS BudgType, 3 AS SortOrd
        FROM <<TblName>>
        WHERE WorkType = '2' AND Flg2Pin = 0
        GROUP BY ShiftDate, UserName

    ENDTEXT

*      SELECT Convert(DateTime, ShiftDate, 104) AS _ShiftDate, UserName, SUM(SuperQty) AS BudgCount, Convert(varchar(15), 'В суперсумках') AS BudgType, 3 AS SortOrd

    IF RunSQL_Load('AllResultp', @ar) <> 1  &&выбрали в предварительную таблицу
      RETURN -1
    ENDIF

*** перевыбираем: делаем гарантированную сортировку, убираем лишние поля ***

    SELECT DTOC(_shiftdate) AS ShiftDate, UserName, BudgCount, BudgType ;
      FROM AllResultp ;
      ORDER BY _ShiftDate, UserName, SortOrd ;
      INTO CURSOR AllResult

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE nRepType = 2  && итоговый отчёт по кассирам ПРК

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT Convert(DateTime, ShiftDate, 104) AS _ShiftDate, UserNamePrk, COUNT(*) AS BudgCount, Convert(varchar(15), 'Стандартные') AS BudgType, 1 AS SortOrd
        FROM <<TblName>>
        WHERE WorkType = '0' AND Flg2Pin = 0
        GROUP BY ShiftDate,UserNameprk
      UNION
      SELECT Convert(DateTime, ShiftDate, 104) AS _ShiftDate, UserNamePrk, COUNT(*) AS BudgCount, Convert(varchar(15), 'Многопиновые') AS BudgType, 2 AS SortOrd
        FROM <<TblName>>
        WHERE WorkType = '0' AND Flg2Pin = 1
        GROUP BY ShiftDate,UserNameprk
      UNION
      SELECT Convert(DateTime, ShiftDate, 104) AS _ShiftDate, UserNamePrk, COUNT(*) AS BudgCount, Convert(varchar(15), 'В суперсумках') AS BudgType, 3 AS SortOrd
        FROM <<TblName>>
        WHERE WorkType = '2' AND Flg2Pin = 0
        GROUP BY ShiftDate,UserNameprk

    ENDTEXT

    IF RunSQL_Load('AllResultp', @ar) <> 1  &&выбрали в предварительную таблицу
      RETURN -1
    ENDIF

*** перевыбираем: делаем гарантированную сортировку, убираем лишние поля ***

    SELECT DTOC(_ShiftDate) AS ShiftDate, UserNamePrk, BudgCount, BudgType ;
      FROM AllResultp ;
      ORDER BY _ShiftDate, UserNamePrk, SortOrd ;
      INTO CURSOR AllResult

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE nRepType = 1  && Оригинальный отчёт по кассирам КП

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT *, 0 AS _1к, 0 AS _5к, 0 AS _10к, 0 AS _50к, 0 AS _1р, 0 AS _2р, 0 AS _5рм, 0 AS _10рм, 0 AS _5р, 0 AS _10р, 0 AS _50р, 0 AS _100р, 0 AS _200р, 0 AS _500р, 0 AS _1000р, 0 AS _2000р, 0 AS _5000р, SPACE(60) AS numACT
      FROM <<TblName>>

    ENDTEXT

    IF RunSQL_Load('AllResult', @ar) <> 1
      RETURN -1
    ENDIF

* BROWSE WINDOW brows

* WAIT WINDOW 'Поиск АКТОВ' TIMEOUT 2

* Поиск АКТОВ

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 0 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 0 AND ;
          (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko')

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT BudgetId, NumAct
          FROM <<TblName>> (READPAST), MistakeLogKP (READPAST)
          WHERE BudgetId = IdBudgetKP
          ORDER BY BudgetId

        ENDTEXT


      CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 1 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 0 AND ;
          (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko')

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT BudgetId, NumAct
          FROM <<TblName>> (READPAST), MistakeLogKP (READPAST)
          WHERE BudgetId = IdBudgetKP
          ORDER BY BudgetId

        ENDTEXT


      CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 1 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 1 AND ;
          (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko')

* WAIT WINDOW '2 - Recount - FROM TblName_MistakeLogKP_2017' TIMEOUT 3

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT BudgetId, NumAct
          FROM <<TblName>> (READPAST), MistakeLogKP (READPAST)
          WHERE BudgetId = IdBudgetKP
          ORDER BY BudgetId

        ENDTEXT


      CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 2 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 2 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

* WAIT WINDOW '2 - Recount - FROM TblName_MistakeLogKP_2016' TIMEOUT 3

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT BudgetId, NumAct
          FROM <<TblName>> (READPAST), MistakeLogKP_2016 AS MistakeLogKP (READPAST)
          WHERE BudgetId = IdBudgetKP
          ORDER BY BudgetId

        ENDTEXT


      CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 3 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 3 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

* WAIT WINDOW '2 - Recount - FROM TblName_MistakeLogKP_2015' TIMEOUT 3

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT BudgetId, NumAct
          FROM <<TblName>> (READPAST), MistakeLogKP_2015 AS MistakeLogKP (READPAST)
          WHERE BudgetId = IdBudgetKP
          ORDER BY BudgetId

        ENDTEXT


      CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 4 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 4 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

* WAIT WINDOW '2 - Recount - FROM TblName_MistakeLogKP_2014' TIMEOUT 3

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT BudgetId, NumAct
          FROM <<TblName>> (READPAST), MistakeLogKP_2014 AS MistakeLogKP (READPAST)
          WHERE BudgetId = IdBudgetKP
          ORDER BY BudgetId

        ENDTEXT


      CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 5 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 5 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

* WAIT WINDOW '2 - Recount - FROM TblName_MistakeLogKP_2013' TIMEOUT 3

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT BudgetId, NumAct
          FROM <<TblName>> (READPAST), MistakeLogKP_2013 AS MistakeLogKP (READPAST)
          WHERE BudgetId = IdBudgetKP
          ORDER BY BudgetId

        ENDTEXT


    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    IF RunSQL_Load('MistakeLogKP', @ar) <> 1
      MESSAGEBOX('Ошибка при запросе к базе данных 777', 16, 'Поиск информации по сумкам')
      RETURN
    ENDIF

* WAIT WINDOW 'Выполнение - SELECT BudgetId, NumAct' TIMEOUT 2

* WAIT WINDOW 'Запись АКТОВ' TIMEOUT 2

* BROWSE WINDOW brows

* Запись АКТОВ

    SELECT AllResult

    SCAN

      SELECT MistakeLogKP

      SCAN FOR AllResult.BudgetId = MistakeLogKP.BudgetId  && Для текущей сумки
        REPLACE AllResult.numAct WITH ALLTRIM(AllResult.numAct) + STR20(MistakeLogKP.numAct, 1) + ';' IN 'AllResult'
      ENDSCAN

    ENDSCAN

    SELECT MistakeLogKP

    USE

* ----------------------------------------------------------------- Переключатель "Покупюрка" ВКЛЮЧЕН ---------------------------------------------------------------------------------------------- *

    IF Check1_Value = 1

* WAIT WINDOW 'Переключатель Покупюрка ВКЛЮЧЕН' TIMEOUT 2

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      DO CASE
        CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 0 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 0 AND ;
            (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko')

          TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT BudgetId, IdValuable, Amount, IdDetailValue, FaceValue, InventoryKP.DtRecord
            FROM <<TblName>> (READPAST)
            INNER JOIN InventoryKP (READPAST) ON InventoryKP.IdBudgetKP = BudgetId
            INNER JOIN Valuable (READPAST) ON valuable.Id = IdValuable
            INNER JOIN FaceValue (READPAST) ON FaceValue.Id = IdFaceValue
            ORDER BY ShiftDate, SuperBudget, WorkType, BudgetNum, IdValuable

          ENDTEXT


        CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 1 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 0 AND ;
            (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko')

          TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT BudgetId, IdValuable, Amount, IdDetailValue, FaceValue, InventoryKP.DtRecord
            FROM <<TblName>> (READPAST)
            INNER JOIN InventoryKP (READPAST) ON InventoryKP.IdBudgetKP = BudgetId
            INNER JOIN Valuable (READPAST) ON valuable.Id = IdValuable
            INNER JOIN FaceValue (READPAST) ON FaceValue.Id = IdFaceValue
            ORDER BY ShiftDate, SuperBudget, WorkType, BudgetNum, IdValuable

          ENDTEXT


        CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 1 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 1 AND ;
            (ALLTRIM(Type_OdbcSrc) == 'Recount' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05' OR ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t' OR ALLTRIM(Type_OdbcSrc) == 'Recount_05_t2' OR ALLTRIM(Type_OdbcSrc) == 'Controller_kp' OR ALLTRIM(Type_OdbcSrc) == 'Recount_11' OR ALLTRIM(Type_OdbcSrc) == 'Recount_30' OR ALLTRIM(Type_OdbcSrc) == 'Recount_eko')

          TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT BudgetId, IdValuable, Amount, IdDetailValue, FaceValue, InventoryKP.DtRecord
            FROM <<TblName>> (READPAST)
            INNER JOIN InventoryKP (READPAST) ON InventoryKP.IdBudgetKP = BudgetId
            INNER JOIN Valuable (READPAST) ON valuable.Id = IdValuable
            INNER JOIN FaceValue (READPAST) ON FaceValue.Id = IdFaceValue
            ORDER BY ShiftDate, SuperBudget, WorkType, BudgetNum, IdValuable

          ENDTEXT


        CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 2 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 2 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

*   WAIT WINDOW '2 - Recount - FROM InventoryKP_2016' TIMEOUT 3

          TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT BudgetId, IdValuable, Amount, IdDetailValue, FaceValue, InventoryKP.DtRecord
            FROM <<TblName>> (READPAST)
            INNER JOIN InventoryKP_2016 AS InventoryKP (READPAST) ON InventoryKP.IdBudgetKP = BudgetId
            INNER JOIN Valuable (READPAST) ON Valuable.Id = IdValuable
            INNER JOIN FaceValue (READPAST) ON FaceValue.Id = IdFaceValue
            ORDER BY ShiftDate, SuperBudget, WorkType, BudgetNum, IdValuable

          ENDTEXT


        CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 3 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 3 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

*   WAIT WINDOW '2 - Recount - FROM InventoryKP_2015' TIMEOUT 3

          TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT BudgetId, IdValuable, Amount, IdDetailValue, FaceValue, InventoryKP.DtRecord
            FROM <<TblName>> (READPAST)
            INNER JOIN InventoryKP_2015 AS InventoryKP (READPAST) ON InventoryKP.IdBudgetKP = BudgetId
            INNER JOIN Valuable (READPAST) ON Valuable.Id = IdValuable
            INNER JOIN FaceValue (READPAST) ON FaceValue.Id = IdFaceValue
            ORDER BY ShiftDate, SuperBudget, WorkType, BudgetNum, IdValuable

          ENDTEXT


        CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 4 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 4 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

*   WAIT WINDOW '2 - Recount - FROM InventoryKP_2014' TIMEOUT 3

          TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT BudgetId, IdValuable, Amount, IdDetailValue, FaceValue, InventoryKP.DtRecord
            FROM <<TblName>> (READPAST)
            INNER JOIN InventoryKP_2014 AS InventoryKP (READPAST) ON InventoryKP.IdBudgetKP = BudgetId
            INNER JOIN Valuable (READPAST) ON Valuable.Id = IdValuable
            INNER JOIN FaceValue (READPAST) ON FaceValue.Id = IdFaceValue
            ORDER BY ShiftDate, SuperBudget, WorkType, BudgetNum, IdValuable

          ENDTEXT


        CASE (YEAR(DATE()) - YEAR(f_Date_Begin)) = 5 AND (YEAR(DATE()) - YEAR(f_Date_End)) = 5 AND ALLTRIM(Type_OdbcSrc) == 'Recount_arxiv_5'

*   WAIT WINDOW '2 - Recount - FROM InventoryKP_2013' TIMEOUT 3

          TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT BudgetId, IdValuable, Amount, IdDetailValue, FaceValue, InventoryKP.DtRecord
            FROM <<TblName>> (READPAST)
            INNER JOIN InventoryKP_2013 AS InventoryKP (READPAST) ON InventoryKP.IdBudgetKP = BudgetId
            INNER JOIN Valuable (READPAST) ON Valuable.Id = IdValuable
            INNER JOIN FaceValue (READPAST) ON FaceValue.Id = IdFaceValue
            ORDER BY ShiftDate, SuperBudget, WorkType, BudgetNum, IdValuable

          ENDTEXT


      ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      IF RunSQL_Load('InventBudget', @ar) <> 1
        MESSAGEBOX('Ошибка при запросе к базе данных 888', 16, 'Поиск информации по сумкам')
        RETURN
      ENDIF

* WAIT WINDOW 'Выполнение - SELECT BudgetId, IdValuable, Amount' TIMEOUT 2

*  BROWSE WINDOW brows

      IF ar[1,2] = 0
        MESSAGEBOX('Данных по выбранному критерию не найдено', 16, 'Поиск информации по сумкам')
        RETURN
      ENDIF

* WAIT WINDOW 'Переключатель Покупюрка ВКЛЮЧЕН' TIMEOUT 2

* -------------------------------------------------------------------------------------------------- *

*        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*          DROP TABLE <<TblName>>

*        ENDTEXT

*        =RunSQL_Load('', @ar)

* -------------------------------------------------------------------------------------------------- *

* Для каждой сумки найти данные по номиналам

      iii = 0

      SELECT AllResult

      SCAN

        SELECT InventBudget

        SCAN FOR AllResult.BudgetId = InventBudget.BudgetId  && Для текущей сумки

          REPLACE DtInventoryKP WITH InventBudget.DtRecord IN 'AllResult'

          DO CASE
            CASE InventBudget.Idvaluable =  1 &&    0.01
              REPLACE _1к WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable =  2 &&    0.05
              REPLACE _5к WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable =  3 &&    0.10
              REPLACE _10к WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable =  4 &&    0.50
              REPLACE _50к WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable =  5 &&    1.00
              REPLACE _1р WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable =  6 &&    2.00
              REPLACE _2р WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable =  7 &&    5.00
              REPLACE _5рм WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable =  8 &&   10.00
              REPLACE _10рм WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable =  9 &&    5.00 (банкноты)
              REPLACE _5р WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable = 10 &&   10.00
              REPLACE _10р WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable = 11 &&   50.00
              REPLACE _50р WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable = 12 &&  100.00
              REPLACE _100р WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable = 13 &&  500.00
              REPLACE _500р WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable = 14 && 1000.00
              REPLACE _1000р WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable = 22 && 5000.00
              REPLACE _5000р WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable = 25 &&  100.00
              REPLACE _200р WITH InventBudget.amount IN 'AllResult'

            CASE InventBudget.Idvaluable = 26 &&  100.00
              REPLACE _2000р WITH InventBudget.amount IN 'AllResult'

          ENDCASE

        ENDSCAN

        m.iii = m.iii + 1

        IF MOD(m.iii, 1000) = 0
          WAIT 'Осталось ' + STR20((m.allbg-m.iii),1) + ' записей' WINDOW NOWAIT
        ENDIF

      ENDSCAN

    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

* WAIT WINDOW 'Записываем в локальный курсор данные о имеющихся дефектов' TIMEOUT 2

* Записываем в локальный курсор данные о имеющихся дефектов

    SELECT DefectLink

    SCAN

      SELECT AllResult

      LOCATE FOR defectLink.BudgetId  = AllResult.BudgetId  && Для текущей сумки

      IF FOUND() = .T.
        REPLACE Defect WITH ALLTRIM(Defect) + STR20(defectLink.IdBudgetDefect, 1) + ';'
      ENDIF

    ENDSCAN

  OTHERWISE  &&nRepType не 1,2 или 3
    MESSAGEBOX("Неизвестный вид выбора",64)
    RETURN

ENDCASE      &&по nRepType

* WAIT WINDOW 'Папка по пути - ' + ALLTRIM(put_tmp) + 'Temp\' + ' существует ... ' TIMEOUT 3

NewFileName = ALLTRIM(put_tmp) + 'Temp\' + 'InventBudget' + ALLTRIM(SYS(2015)) + '.xls'

IF DIRECTORY(ALLTRIM(put_tmp) + 'Temp\', 1) = .T.  && Проверяем доступность сетевой папки, в которую ниже будет создаваться отчет

  IF FILE(NewFileName) = .T.
    DELETE FILE (NewFileName)
  ENDIF

ELSE
  =err('Вам не доступна сетевая папка - ' + ALLTRIM(put_tmp) + 'Temp\' + ' для записи отчета ... ')
ENDIF

WAIT CLEAR

tim2 = SECONDS()

result_time = ROUND(tim2 - tim1,  2)

@ WROWS()/3, 3 SAY PADC('Выборка данных успешо завершена.' + '  (время = ' + ALLTRIM(TRANSFORM(result_time,  '999.99')) + ' сек.)', WCOLS())
= INKEY(2)
DEACTIVATE WINDOW poisk

SELECT AllResult

COUNT TO colvo_zap
GO TOP

* BROWSE WINDOW brows

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW 'EXPORT TO NewFileName XL5' TIMEOUT 2
* WAIT WINDOW 'Папка по пути - ' + ALLTRIM(put_tmp) + 'Temp\' + ' существует ... ' TIMEOUT 3

IF colvo_zap > 0

  IF DIRECTORY(ALLTRIM(put_tmp) + 'Temp\', 1) = .T.  && Проверяем доступность сетевой папки, в которую ниже будет создаваться отчет

    EXPORT TO (NewFileName) XL5

    IF FILE(NewFileName) = .T.

      MESSAGEBOX('Операция выполнена успешно файл: ' + ALLTRIM(NewFileName), 64, 'Поиск информации по сумкам')

      loExcel = CREATEOBJECT('Excel.Application')
      loExcel.WorkBooks.OPEN(NewFileName,.T.,.T.)
      loExcel.Visible= .T.

    ELSE
      =err('Файл, для работы, по пути  - ' + ALLTRIM(NewFileName) + ' не сформирован ... ')
    ENDIF

  ELSE
    =err('Вам не доступна сетевая папка - ' + ALLTRIM(put_tmp) + 'Temp\' + ' для записи отчета ... ')
  ENDIF

ELSE
  =err('Внимание! Записей для формирования выходного файла в формате Excel не обнаружено ... ')
ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF USED('AllResult') = .T.
  SELECT AllResult
  USE
ENDIF

IF USED('InventBudget') = .T.
  SELECT InventBudget
  USE
ENDIF

IF USED('DefectLink') = .T.
  SELECT DefectLink
  USE
ENDIF

IF USED('MistakeLogKP') = .T.
  SELECT MistakeLogKP
  USE
ENDIF

IF USED('AllResult')
  USE IN AllResult
ENDIF

SELECT(sel)
RETURN


************************************************************************************************************************************************


