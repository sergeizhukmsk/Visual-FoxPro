
USE RecountCC

 --UPDATE dbo.Filial SET MestoCode = 'KTL' WHERE FilialCode = '25-6947'
 --UPDATE dbo.Filial SET MestoCode = 'YAR' WHERE FilialCode = '25-5535'
 --UPDATE dbo.Filial SET MestoCode = 'YAR' WHERE FilialCode IN ('97-0068', '97-0069', '97-087')
 --UPDATE dbo.Filial SET MestoCode = 'KTL' WHERE FilialCode IN ('97-0268','97-0059')
 --UPDATE dbo.Filial SET MestoCode = 'KTL' WHERE Id IN (52426,52420)
 --UPDATE dbo.Filial SET MestoCode = 'YAR' WHERE Id IN (54853,54855)
 --UPDATE dbo.Filial SET IsActive = 1 WHERE Id = 38814
 --UPDATE dbo.Filial SET FlgRecAlc = 0 WHERE Id IN (54853,54855)
 
SELECT fil.Id
      ,fil.IdClient
--      ,fil.SuperSumka
      ,clg.Id AS IdClientGroup
--      ,con.Id AS IdContract
--      ,ISNULL(con.IsJKServe, 1) AS IsJKServe
      ,fil.IdRoute
      ,fil.IsActive AS IsActive_Fil
      ,cln.IsActive AS IsActive_Cln
      ,fil.FilialCode
   	  ,cln.Flg2PIN
--      ,fil.MestoCode
      ,fil.BasePIN
      ,fil.FlgRecalc
      ,cln.PointType
      ,fil.FilialName
      ,fil.UiFilialName
      ,fil.UiFactAddress
      ,fil.DtBeginPer
      ,fil.DtEndPer
      ,fil.DtClose
--  	  ,fil.DtUpdUI
  FROM dbo.Filial AS fil (READPAST)
  INNER JOIN Client AS cln (READPAST) ON cln.Id = fil.IdClient
  INNER JOIN ClientGroup AS clg (READPAST) ON clg.Id = cln.IdClientGroup
  --INNER JOIN [Contract] AS con (READPAST) ON con.Id = clg.IdContract
  --WHERE fil.FilialCode = '25-75210'
  --WHERE con.IsJKServe = 0 AND fil.FlgRecalc = 1
  --WHERE fil.FilialCode IN ('25-6260', '25-6264', '25-333', '25-347')
  --WHERE LEFT(fil.FilialCode, 3) = '09-' AND fil.IsActive = 1 AND fil.FlgRecalc = 0
  --WHERE fil.IdClient = 34523
  --WHERE fil.FilialCode LIKE '25-7521%' --AND fil.IsActive = 1 AND FlgRecalc = 1
  --WHERE fil.FilialCode = '16-470' --AND IsActive = 1 AND FlgRecalc = 1
  --WHERE fil.FilialName LIKE '%Росбанк%' AND fil.IsActive = 1 AND FlgRecalc = 0 --AND MestoCode = 'KTL'
  --WHERE fil.IsActive = 1 AND FlgRecalc = 0 AND cln.PointType = 1
  --WHERE MestoCode = 'YAR' AND fil.FlgRecalc = 0 AND fil.IsActive = 1
  --WHERE fil.FilialName LIKE '%дроп-слот%' --AND MestoCode = 'KTL'
  --WHERE fil.FilialName LIKE '%дроп-слот%' AND clg.Id = 2115
  --WHERE LEFT(fil.FilialCode, 4) = '91-4'
  --WHERE fil.IdClient IN (33263,32012)
  --ORDER BY fil.Id DESC
  ORDER BY fil.FilialCode ASC

-- изменен статус по клиентам ООО «ВЕМЕС» и ООО «НБА-СЕВЕРО-ЗАПАД».

  --25-6705, 46420//Россия АЗС Shell 2019 МКАД 7 км 6 дроп слот
  --AND fil.FilialName LIKE '%Shell%' AND NOT fil.FilialName LIKE '%дроп-слот%'
  --NOT fil.FilialName LIKE '%Shell%' AND NOT fil.FilialName LIKE '%дроп-слот%'

--МОСКЛИРИНГЦЕНТР
--Русский Стандарт
--ООО КАРИ
--WHERE FilialName LIKE '%МИЦАР%'
--WHERE FilialName LIKE '%Элекснет%'  


