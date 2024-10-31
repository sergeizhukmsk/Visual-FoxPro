
USE RecountCC

DECLARE @IdShift Int
SET @IdShift = 25875

--DELETE dbo.BudgetHistoryPRK WHERE IdSecurPack IN (3503875, 3503874)

--UPDATE dbo.BudgetHistoryPRK SET IdRecShift = 10760, IdCurShift = 10760 WHERE IdSecurPack IN (3123723,3123722,3123721,3123720)
--UPDATE dbo.BudgetHistoryPRK SET IdRecShift = 10852, IdCurShift = 10852 WHERE Id IN (8082299, 8082298, 8082301, 8082300)
--UPDATE dbo.BudgetHistoryPRK SET IdRecShift = 10938, IdCurShift = 10938 WHERE Id = 8119683
--UPDATE dbo.BudgetHistoryPRK SET IdRecShift = 12463, IdCurShift = 12463 WHERE IdCurShift = 12460
--UPDATE dbo.BudgetHistoryPRK SET IdRecShift = 11110, IdCurShift = 11110 WHERE IdTranSheet = 216714
--UPDATE dbo.BudgetHistoryPRK SET IdRecShift = 10760, IdCurShift = 10760 WHERE IdSecurPack >= 3801401 AND IdSecurPack <= 3801431
--UPDATE dbo.BudgetHistoryPRK SET TotalSum = 1111100.00 WHERE Id = 11753711
--UPDATE dbo.BudgetHistoryPRK SET TotalSum = 1000.00 WHERE Id = 8122470
--UPDATE dbo.BudgetHistoryPRK SET TotalSum = 452758.40 WHERE IdSecurPack = 3854434
--UPDATE dbo.BudgetHistoryPRK SET TotalSum = 8630878.53 WHERE IdSecurPack = 3852959
--UPDATE dbo.BudgetHistoryPRK SET SelDat = 0 WHERE IdCurShift = 24547 AND SelDat = 1 AND IdTranSheet IS NULL
--UPDATE dbo.BudgetHistoryPRK SET SelDat = 0 WHERE IdSecurPack BETWEEN 7172526 AND 7172548  --7172550
--UPDATE dbo.BudgetHistoryPRK SET SelDat = 0 WHERE IdSecurPack = 7172550
--UPDATE dbo.BudgetHistoryPRK SET SelDat = 0 WHERE SuperBudget = 11793040
--UPDATE dbo.BudgetHistoryPRK SET SelDat = 0 WHERE IdCurShift = @IdShift AND BudgetNum IN (SELECT BudgetNum FROM dbo.SecurPack WHERE IdShift = @IdShift AND IdRoute = 43) AND IdTranSheet IS NULL
--UPDATE dbo.BudgetHistoryPRK SET BudgetNum = '06-027091102' WHERE Id = 8156518
--UPDATE dbo.BudgetHistoryPRK SET BudgetNum = '06-009091103' WHERE Id = 8156520
--UPDATE dbo.BudgetHistoryPRK SET Pr_Exp = 0 WHERE Id = 11069896
--UPDATE dbo.BudgetHistoryPRK SET State = 1 WHERE SuperBudget = 11793040
--UPDATE dbo.BudgetHistoryPRK SET IdTypeValue = 1, TypeValueKind = 'RUB' WHERE Id = 12858997

SELECT bdg.Id
      ,bdg.IdSecurPack
      ,sec.IdTsd
      ,sec.IdRoute
      ,bdg.IdRouteHistory
      ,bdg.IdUser
      ,bdg.SelDat
      ,bdg.Pr_Exp
      ,bdg.IdFilial
      ,fil.FilialCode
      ,bdg.BudgetNum
      ,bdg.TotalSum
      ,sec.Lait
      ,bdg.State
      ,bdg.IdCurShift
      ,bdg.IsTransferred
      ,bdg.DateTransfer
      ,bdg.IdTranSheet
      ,shf.ReportDate
      ,bdg.DtRecord
      ,bdg.TypeValueKind
      ,bdg.SuperBudget
      ,bdg.IdRecShift
      ,fil.FilialName
      ,bdg.WorkType
      ,bdg.SuperQty
      ,bdg.IdSignProc
      ,bdg.DtIncome
      ,bdg.Ref
      ,bdg.Wt_flag
      ,bdg.Pin2Sum
      ,bdg.CurDate
      ,bdg.Flg2PIN
      ,usr.UserName
      ,bdg.IdVedList
  FROM dbo.BudgetHistoryPRK AS bdg (READPAST)
  INNER JOIN dbo.SecurPack AS sec (READPAST) ON sec.Id = bdg.IdSecurPack
  INNER JOIN dbo.Filial AS fil (READPAST) ON fil.Id = bdg.IdFilial
  INNER JOIN  dbo.[User] AS usr (READPAST) ON usr.Id = bdg.IdUser
  INNER JOIN dbo.Shift AS shf (READPAST) ON shf.Id = bdg.IdCurShift
  WHERE shf.Status > 0 AND shf.IdDepartment = 1 --AND sec.IdRoute = 20 --AND bdg.IdTranSheet IS NULL --AND sec.Lait = 1 --AND bdg.State = 1 --AND bdg.IsTransferred = 0 AND bdg.IdTranSheet IS NULL  -- NOT 
  --WHERE shf.Status > 0 AND shf.IdDepartment IN (1, 1) AND bdg.IdTranSheet = 430302
  --WHERE shf.Status > 0 AND shf.IdDepartment IN (1, 1) AND bdg.IdTranSheet IN (430272,430273,430274,430241,430240,430239,430275,430276,430277,430121,430216,430215,430221,430222,430189,430190,430322,430323,430324,430280,430281)
  --WHERE shf.Status > 0 AND shf.IdDepartment IN (1, 1) AND bdg.BudgetNum LIKE '%36122%'
  --WHERE bdg.IdCurShift = @IdShift AND sec.IdRoute = 20
  --WHERE bdg.BudgetNum LIKE '%91-5%'
  --WHERE bdg.IdFilial = 40256
  --WHERE shf.Status > 0 AND shf.IdDepartment IN (1, 1) AND bdg.SelDat = 1 AND bdg.IdTranSheet IS NULL
  --WHERE shf.Status > 0 AND shf.IdDepartment IN (1, 1) AND bdg.IdTranSheet IS NULL AND bdg.SelDat = 1 --AND sec.IdRoute = 59
  --WHERE bdg.IdSecurPack BETWEEN 7172526 AND 7172550
  --WHERE shf.Status > 0 AND shf.IdDepartment IN (1,41,51) AND bdg.SelDat = 1 AND bdg.IdTranSheet IS NULL
  --WHERE bdg.IdCurShift = @IdShift --AND sec.Lait = 1 --AND bdg.State = 0 --AND bdg.WorkType > '0' --AND bdg.Flg2PIN = 1
  --WHERE bdg.IdCurShift = @IdShift AND bdg.BudgetNum IN (SELECT BudgetNum FROM dbo.SecurPack WHERE IdShift = @IdShift AND IdRoute = 43) AND bdg.IdTranSheet IS NULL
  --WHERE bdg.IdCurShift = @IdShift AND bdg.BudgetNum LIKE '%41430481%'
  --WHERE bdg.IdSecurPack >= 3801401 AND bdg.IdSecurPack <= 3801431
  --WHERE bdg.IdCurShift IN (10128, 10134) AND bdg.BudgetNum LIKE '%06-027%'
  --WHERE bdg.BudgetNum LIKE '%36995703%'
  --WHERE bdg.BudgetNum = '36995703'
  --WHERE bdg.Id IN (8681782,8681780,8681779,8681777)
  --WHERE bdg.Id = 12863827
  --WHERE bdg.IdTranSheet = 353757
  --WHERE bdg.IdTranSheet IN (402541,402540,402539)
  --WHERE bdg.IdSecurPack = 8647253
  --WHERE bdg.SuperBudget = 1029476
  --WHERE bdg.Id = 12523954
  --WHERE bdg.IdSecurPack IN (8681782,8681780,8681779,8681777)
  --WHERE bdg.IdCurShift = @IdShift AND bdg.BudgetNum IN ('AE288087','AF111005','KO873092')
  --WHERE bdg.IdCurShift = @IdShift AND fil.FilialName LIKE '%Русская телефонная компания%'
  --WHERE bdg.IdCurShift = @IdShift AND fil.FilialName LIKE '%РН-Москва%'
  --WHERE bdg.IdCurShift = @IdShift AND bdg.SelDat = 1 AND bdg.IdTranSheet IS NULL
  --WHERE bdg.IdCurShift IN (7962, 7963) AND bdg.BudgetNum NOT IN (SELECT BudgetNum FROM dbo.BudgetKP WHERE IdShift IN (7965, 7969))
  --WHERE bdg.IdCurShift IN (7962, 7963) AND bdg.Id NOT IN (SELECT IdBudgetHistoryPRK FROM dbo.BudgetKP WHERE IdShift IN (7965, 7969))
  --WHERE bdg.IdCurShift = @IdShift AND bdg.BudgetNum LIKE '%09-482%'
  --WHERE bdg.IdCurShift = @IdShift AND bdg.SelDat = 1
  --WHERE bdg.IdCurShift = @IdShift AND bdg.IdSecurPack = 0 AND bdg.WorkType = 0
  --WHERE bdg.IdCurShift = @IdShift AND bdg.IdTranSheet IN (179337,179336,179335,179334,179333)
  --WHERE bdg.IdCurShift = @IdShift AND (bdg.BudgetNum LIKE '%/2' OR bdg.BudgetNum LIKE '%/3') AND bdg.IdSecurPack = 0 --AND bdg.Flg2PIN > 0
  --WHERE bdg.IdCurShift = @IdShift AND fil.FilialName LIKE '%РН-Москва%' AND bdg.Id > 6299110  --6297255
  --ORDER BY bdg.BudgetNum DESC
  ORDER BY bdg.IdSecurPack DESC
  --ORDER BY bdg.DtRecord DESC
  --ORDER BY bdg.IdUser ASC

  --WHERE bdg.IdCurShift IN (7962, 7963) AND bdg.BudgetNum NOT IN (SELECT BudgetNum FROM dbo.BudgetKP WHERE IdShift IN (7965, 7969))
  --WHERE bdg.IdCurShift IN (7962, 7963) AND bdg.Id NOT IN (SELECT IdBudgetHistoryPRK FROM dbo.BudgetKP WHERE IdShift IN (7965, 7969))
  
  --WHERE bdg.IdCurShift IN (7665, 7668) AND (bdg.BudgetNum LIKE '%AF116781%' OR bdg.BudgetNum LIKE '%KO848021%' OR bdg.BudgetNum LIKE '%KO851836%')
  --WHERE bdg.IdCurShift = 7665 AND (bdg.BudgetNum LIKE '%AF116781%' OR bdg.BudgetNum LIKE '%KO848021%' OR bdg.BudgetNum LIKE '%KO851836%')
  --WHERE bdg.IdSecurPack IN (2648258,2648257,2648244,2648256,2648255,2648248,2648254,2648253)
  --WHERE bdg.IdCurShift IN (7665, 7668) AND bdg.BudgetNum LIKE '%AF116781%'

  --WHERE bdg.IdCurShift = 7665 AND bdg.BudgetNum IN (SELECT BudgetNum FROM dbo.BudgetHistoryPRK WHERE IdCurShift = 7599)
  --WHERE bdg.IdCurShift IN (7665, 7668) AND bdg.BudgetNum IN ('KO882981', 'KO882479', 'KO881541', 'KO874926')

  --WHERE bdg.IdCurShift = 7566 AND bdg.BudgetNum NOT IN (SELECT BudgetNum FROM dbo.BudgetKP WHERE IdShift = 7569)
  --WHERE bdg.IdCurShift = 7566 AND bdg.Flg2PIN <> 2 AND bdg.Id NOT IN (SELECT IdBudgetHistoryPRK FROM dbo.BudgetKP WHERE IdShift = 7569)

  -- AND bdg.BudgetNum LIKE '%09-482%'
  --AND bdg.BudgetNum IN ('AE288087','AF111005','KO873092')
  --WHERE bdg.Id = 21113852
  --AND bdg.BudgetNum NOT IN (SELECT BudgetNum FROM dbo.BudgetKP WHERE IdShift = @IdShift)
  --AND bdg.BudgetNum IN (SELECT BudgetNum FROM dbo.SecurPack WHERE IdShift = @IdShift)
  --AND bdg.BudgetNum IN (SELECT BudgetNum FROM dbo.BudgetHistoryPRK WHERE IdCurShift = @IdShift)

--DELETE dbo.BudgetHistoryPRK WHERE IdSecurPack = 2535996
--DELETE dbo.RouteHistory WHERE Id = 568735
--DELETE dbo.SecurPack WHERE Id = 2535996
--DELETE dbo.BudgetHistoryPRK  WHERE IdSecurPack IN (2648258,2648257,2648244,2648256,2648255,2648248,2648254,2648253)

--UPDATE dbo.BudgetHistoryPRK SET IdRecShift = 8419, IdCurShift = 8419 WHERE IdSecurPack IN (3123723,3123722,3123721,3123720,3123719,3123718,3123717,3123716,3123715,3123714,3123713,3123712)
--UPDATE  dbo.BudgetHistoryPRK SET IdSecurPack = 2818293 WHERE Id = 7162996

