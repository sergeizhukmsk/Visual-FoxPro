
USE RecountCC

DECLARE @IdShift Int
SET @IdShift = 30592

--DELETE dbo.SecurPack WHERE Id IN (3503875, 3503874)

--UPDATE dbo.SecurPack SET IdShift = 10760 WHERE Id IN (3123723,3123722,3123721,3123720)
--UPDATE dbo.SecurPack SET IdShift = 10938 WHERE Id >= 3887613
--UPDATE dbo.SecurPack SET IdShift = 10760 WHERE Id >= 3801401 AND Id <= 3801431
--UPDATE dbo.SecurPack SET TotalSum = 566230.00 WHERE Id = 3890556
--UPDATE dbo.SecurPack SET Sost = 1 WHERE Id = 7740994
--UPDATE dbo.SecurPack SET BudgetNum = '06-027091102' WHERE Id = 3927938
--UPDATE dbo.SecurPack SET BudgetNum = '06-009091103' WHERE Id = 8647253
--UPDATE dbo.SecurPack SET IdTypeValue = 1, TypeValueKind = 'RUB' WHERE Id = 8647253
--UPDATE dbo.SecurPack SET IdShift = 12463 WHERE IdShift = 12460

SELECT sec.Id
      ,sec.IdShift
      ,sec.IdTsd
      ,sec.Lait
      ,sec.IdRoute
      ,sec.IdUser
      ,sec.TabNumber
      ,sec.DtRecord
      ,shf.ReportDate
      ,sec.BudgetNum
      ,sec.FilialCode
      ,sec.TotalSum
      --,sec.ColvoKey
      --,sec.IdKeyATMDocMain
      --,sec.IdKeyATM
      ,sec.Flg2PIN
      ,sec.RecIdRoute
      ,sec.Sost
      ,fil.FilialName
      ,sec.TypeValueKind
      ,sec.WorkType
      ,sec.SuperBudget
      ,sec.SuperQty
      ,sec.DateInkas
      ,sec.DtIncome
      ,sec.CurDate
      ,sec.IsError
      ,sec.OperCode
      ,sec.IdJk
      ,sec.IdPackDefect
      ,sec.State
      ,sec.IsTransit
      ,sec.Pr_60322
      ,sec.TotalSum_60322
  FROM dbo.SecurPack AS sec (READPAST)
  INNER JOIN dbo.Filial AS fil (READPAST) ON fil.FilialCode = sec.FilialCode
  INNER JOIN dbo.Shift AS shf (READPAST) ON shf.Id = sec.IdShift
  WHERE shf.Status > 0 AND shf.IdDepartment = 1 --AND sec.IdRoute = 20  --AND sec.WorkType > 0 --AND sec.FilialCode LIKE '%25-2147%' --AND sec.IdRoute = 302
  --WHERE shf.Status > 0 AND shf.IdDepartment = 1 AND sec.IdTsd > 0 AND sec.Lait > 0
  --WHERE shf.Status > 0 AND shf.IdDepartment = 41 AND sec.IdTsd = 0
  --WHERE shf.Status > 0 AND shf.IdDepartment IN (1, 1) AND sec.IdTsd > 0 AND fil.FilialName LIKE '%ÌÎÑÊÎÂÑÊÈÉ ÊÐÅÄÈÒÍÛÉ ÁÀÍÊ%'
  --WHERE sec.IdShift = @IdShift --AND sec.ColvoKey > 0  --AND sec.WorkType > 0  --AND sec.IdPackDefect > 0 --AND sec.IdRoute = 1 --AND sec.IdRoute = 122
  --WHERE sec.Id >= 4740412 AND sec.Id <= 4740533
  --WHERE sec.IdShift = @IdShift AND sec.IdTsd > 0 AND sec.Lait > 0
  --WHERE sec.IdShift = @IdShift AND sec.Flg2PIN > 0
  --WHERE sec.IdShift = @IdShift AND sec.FilialCode LIKE '%06-07%'
  --WHERE sec.WorkType = '2' AND sec.TotalSum > 0
  --WHERE sec.FilialCode LIKE '%91-785%' AND MONTH(shf.ReportDate) = 10
  --WHERE sec.IdShift = @IdShift AND sec.BudgetNum IN ('KN071351', 'KN915023', 'KO011398')
  --WHERE sec.IdShift = @IdShift AND sec.BudgetNum LIKE '%09-482%' --AND sec.WorkType = 1
  --WHERE sec.BudgetNum LIKE '%09-482%' AND sec.WorkType = 1
  --WHERE sec.BudgetNum LIKE '%0296843%'
  --WHERE sec.SuperBudget = '1N5YU024YN'
  --WHERE sec.BudgetNum   IN ('37386429','36176722','36167070','35058537')
  --WHERE sec.IdShift = @IdShift AND fil.FilialName LIKE '%ÐÍ-Ìîñêâà%' AND sec.Flg2PIN = 2
  --WHERE sec.IdShift = @IdShift AND fil.FilialName LIKE '%ÏÑÊÁ%'
  --WHERE fil.FilialName LIKE '%ÏÑÊÁ%' AND sec.IdShift = 20528 --AND sec.IdPackDefect = 21    
  --WHERE fil.FilialName LIKE '%ÏÑÊÁ%' AND sec.TotalSum = 36400
  --WHERE sec.IdShift = @IdShift AND fil.FilialCode = '25-357'
  --WHERE YEAR(sec.DtRecord) = 2018 AND sec.WorkType IN (1, 2) AND sec.IdKeyATM > 0
  --WHERE LEFT(sec.FilialCode, 2) IN ('59', '62', '76') AND sec.State = 2
  --WHERE sec.Id = 7740994
  --WHERE sec.Id IN (7580246,7580245,7580244)
  --WHERE sec.IdTsd = 3095150
  --WHERE sec.IdTsd IN (626864,626865,626866,626867,626868,626869,626891,626892,626893,626894,626895,626896)
  ORDER BY sec.Id DESC
  --ORDER BY sec.FilialCode DESC
  --ORDER BY sec.IdRoute, sec.BudgetNum DESC
  --ORDER BY sec.BudgetNum ASC
  --ORDER BY sec.DtRecord DESC
  
  --UPDATE dbo.SecurPack SET RecIdRoute = 1  WHERE IdShift = @IdShift AND IdRoute = 115
  --WHERE LEFT(sec.FilialCode, 2) = '62'
  --WHERE sec.IdPackDefect > 0
  --WHERE sec.IdShift = @IdShift AND sec.BudgetNum IN (SELECT BudgetNum FROM [RecountCC].[dbo].[SecurPack] WHERE sec.IdShift = @IdShift)
  --WHERE sec.BudgetNum IN (SELECT BudgetNum FROM [RecountCC].[dbo].[BudgetHistoryTRN] WHERE IdRecShift = 5948 AND IdCurShift = @IdShift)
  --WHERE sec.IdShift = @IdShift AND sec.BudgetNum IN (SELECT BudgetNum FROM [RecountCC].[dbo].[BudgetHistoryTRN] WHERE IdSecurPack = 0 AND IdRecShift = @IdShift AND IdCurShift = @IdShift)
  --WHERE sec.IdShift = @IdShift AND sec.IdRoute IN (32, 76)
  --WHERE sec.IdShift = @IdShift AND sec.BudgetNum LIKE '%AD347528%'
  --WHERE sec.IdShift = @IdShift AND sec.FileName LIKE '%EXCHANGE%' AND Flg2PIN > 0
  --WHERE sec.IdShift = @IdShift AND sec.BudgetNum IN ('AE138524', 'KN657139', 'KN722873', 'AE110859', 'KN709833', 'AD557284', 'AD892714', 'AA005189')
  
  --UPDATE [[dbo.[SecurPack SET Flg2PIN = 1 WHERE sec.IdShift = @IdShift AND FileName NOT LIKE '%EXCHANGE%' AND Flg2PIN = 2
  --WHERE sec.IdShift = @IdShift AND (sec.FileName LIKE '%EXCHANGE%' OR sec.FileName LIKE '%Ðó÷íîé ââîä íîìåðà ïàêåòà%')  AND sec.Flg2PIN > 0
  --DELETE dbo.SecurPack WHERE IdShift = @IdShift AND Sost = 0 AND IdRoute = 76
  --UPDATE dbo.SecurPack SET IdShift = 7879 WHERE Id IN (3123723,3123722,3123721,3123720,3123719,3123718,3123717,3123716,3123715,3123714,3123713,3123712)
