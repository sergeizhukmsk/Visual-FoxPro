 
USE RecountCC

--DELETE dbo.OperDate WHERE Id = 2852

--UPDATE dbo.OperDate SET Status = 2, CloseDate = NULL, LastOrder = NULL WHERE Id = 2322
--UPDATE dbo.OperDate SET Status = 0, CloseDate = '20191018 20:15:30', LastOrder = 99999 WHERE Id = 2856
--UPDATE dbo.OperDate SET Status = 0, CloseDate = GetDate(), LastOrder = 99999 WHERE Id = 2903
--UPDATE dbo.OperDate SET Status = 0, CloseDate = '20190222' WHERE Id = 2612
--UPDATE dbo.OperDate SET Status = 1, CloseDate = NULL, LastOrder = NULL WHERE Id = 2903
--UPDATE dbo.OperDate SET Status = 1 WHERE Id = 2904
--UPDATE dbo.OperDate SET LastOrder = NULL WHERE Id = 1937
--UPDATE dbo.OperDate SET OpDate = '20180109' WHERE Id = 2313
--UPDATE dbo.OperDate SET IdDepartment = 21 WHERE Id = 2529

SELECT TOP 5
      opd.Id
      ,opd.OpDate
      ,opd.Status
      ,opd.IdDepartment
      ,opd.OpenDate
      ,opd.CloseDate
      ,opd.LastOrder
      ,opd.Blok
  FROM dbo.OperDate AS opd (READPAST)
  --WHERE opd.Id = 2803
  ORDER BY opd.Id DESC

    

