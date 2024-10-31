********************************************************************** PROCEDURE start ********************************************************

PROCEDURE start
PARAMETERS Param_IdShift, pShiftManagerUserName

sel = SELECT()

tim1 = SECONDS()

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет выборка данных по ПРИХОДНЫМ операциям ... ',WCOLS())

* ---------------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

*  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*    SELECT bdg.Id, bdg.BudgetNum, SUM(okp.Amount) AS Amount, (SUM(fcv.FaceValue * okp.Amount)) AS TotalSumKP, okp.IdUser, usr.UserName, okp.CdOper, okp.IdUser2, usr2.UserName AS UserName2
*    FROM BudgetKP AS bdg (READPAST)
*    INNER JOIN OperKP AS okp (READPAST) ON okp.IdBudgetKP = bdg.Id
*    INNER JOIN Valuable AS Vlb ON Vlb.Id = okp.IdValuable
*    INNER JOIN FaceValue AS Fcv ON Fcv.Id = Vlb.IdFaceValue
*    INNER JOIN DetailValue AS Dtv ON Dtv.Id = Vlb.IdDetailValue
*    INNER JOIN [User] AS usr ON usr.Id = okp.IdUser
*    INNER JOIN Shift AS shf (READPAST) ON shf.Id = bdg.IdShift
*    LEFT OUTER JOIN [User] AS usr2 ON usr2.Id = okp.IdUser2
*    WHERE YEAR(shf.ReportDate) = YEAR(GetDate()) AND okp.IdShift = ?Param_IdShift AND okp.CdOper IN (1, 3, 7, 9, 13)
*    GROUP BY bdg.Id, bdg.BudgetNum, okp.CdOper, okp.IdUser, usr.UserName, okp.IdUser2, usr2.UserName
*    UNION
*    SELECT okp.Id, SPACE(20) AS BudgetNum, SUM(okp.Amount) AS Amount, (SUM(fcv.FaceValue * okp.Amount)) AS TotalSumKP, okp.IdUser, usr.UserName, okp.CdOper, okp.IdUser2, usr2.UserName AS UserName2
*    FROM OperKP AS okp (READPAST)
*    INNER JOIN Valuable AS Vlb ON Vlb.Id = okp.IdValuable
*    INNER JOIN FaceValue AS Fcv ON Fcv.Id = Vlb.IdFaceValue
*    INNER JOIN DetailValue AS Dtv ON Dtv.Id = Vlb.IdDetailValue
*    INNER JOIN [User] AS usr ON usr.Id = okp.IdUser
*    INNER JOIN Shift AS shf (READPAST) ON shf.Id = okp.IdShift
*    LEFT OUTER JOIN [User] AS usr2 ON usr2.Id = okp.IdUser2
*    WHERE YEAR(shf.ReportDate) = YEAR(GetDate()) AND okp.IdShift = ?Param_IdShift AND okp.CdOper IN (1, 3, 7, 9, 13) AND okp.IdBudgetKP IS NULL
*    GROUP BY okp.Id, okp.CdOper, okp.IdUser, usr.UserName, okp.IdUser2, usr2.UserName
*    ORDER BY 5, 7, 2

*  ENDTEXT

*** AND okp.IdUser = 1719

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  EXECUTE Select_OperKP_BudgetKP_ControlList_Prixod ?Param_IdShift

ENDTEXT

IF RunSQL_Load('Prixod', @ar) <> 1
  DO exit_sql_server IN RunSQL_Load
  MESSAGEBOX('НЕТ ДАННЫХ по ПРИХОДУ по выбранному кассиру', 64, 'Отказ от Операции')
  RETURN
ENDIF

* BROWSE WINDOW brows TITLE 'Приходные операции'

* ---------------------------------------------------------------- *
DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

tim2 = SECONDS()

result_time = ROUND(tim2 - tim1, 2)

@ WROWS()/3,3 SAY PADC('Выборка данных по ПРИХОДНЫМ операциям успешно завершена.' + ' (время = ' + ALLTRIM(TRANSFORM(result_time, '999.99')) + ' сек.)',WCOLS())
=INKEY(3)
DEACTIVATE WINDOW poisk

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SELECT Prixod
GO TOP

SELECT Kassirs
GO TOP

* BROWSE WINDOW brows

* Выбрать данные только для отмеченных пользователем на экране (Kassirs.fl_fl = 1)

* SELECT '01' AS Group, kso.BudgetNum, kso.Idvaluable, kso.FaceValueName, kso.FaceValue, kso.Amount, kso.TotalSumKP, kso.IdUser, kso.UserName, kso.IdUser2, kso.UserName2 ;

SELECT '01' AS Group, kso.BudgetNum, kso.Amount, kso.TotalSumKP, kso.IdUser, kso.UserName, kso.IdUser2, kso.UserName2 ;
  FROM Prixod AS kso ;
  INNER JOIN Kassirs AS kas ON kas.IdUser = kso.IdUser ;
  WHERE kas.fl_fl = 1 ;
  INTO CURSOR For_Rep_Prixod READWRITE ;
  ORDER BY Id

* BROWSE WINDOW brows

IF _TALLY = 0
  MESSAGEBOX('НЕТ ДАННЫХ по ПРИХОДУ по выбранному кассиру', 64, 'Данные по ПРИХОДНЫМ операциям')
  RETURN
ENDIF

* BROWSE WINDOW brows

SELECT For_Rep_Prixod

DO CASE
  CASE sergeizhuk_flag = .F.

    REPLACE UserName2 WITH 'Нач. смены КП - ' + ALLTRIM(pShiftManagerUserName) FOR ISNULL(UserName2) = .T.

  CASE sergeizhuk_flag = .T.

    REPLACE UserName2 WITH IIF(EMPTY(ALLTRIM(BudgetNum)) = .F., 'Нач. смены КП - ' + ALLTRIM(pShiftManagerUserName) + ' - ' + ALLTRIM(BudgetNum), 'Клиентская инкассация') FOR ISNULL(UserName2) = .T.

ENDCASE

GO TOP

* BROWSE WINDOW brows

SELECT(sel)
RETURN


************************************************************************************************************************************************





