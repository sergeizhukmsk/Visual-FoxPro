 
USE RecountCC

--DELETE dbo.Shift WHERE Id = 29201

--UPDATE dbo.Shift SET Status = 1, CloseDate = NULL WHERE Id = 30175
--UPDATE dbo.Shift SET Status = 0, CloseDate = GetDate() WHERE Id = 30164
--UPDATE dbo.Shift SET Status = 3, CloseDate = NULL WHERE Id = 25881

--UPDATE dbo.Shift SET Status = 1 WHERE Id = 30178
--UPDATE dbo.Shift SET Status = 2 WHERE Id = 25875

--UPDATE dbo.Shift SET IdOperDate = 2860 WHERE Id = 29203

--UPDATE dbo.Shift SET IdPart = 2 WHERE Id = 29469
--UPDATE dbo.Shift SET IdDay = 5 WHERE Id = 21110
--UPDATE dbo.Shift SET IdDay = DAY(ShiftDate) WHERE IdDay <> DAY(ShiftDate)

--UPDATE dbo.Shift SET IdUser = 2018 WHERE Id = 25802
--UPDATE dbo.Shift SET IdDepartment = 13 WHERE Id = 9947

--UPDATE dbo.Shift SET Blok = 0 WHERE Id = 29217

--UPDATE dbo.Shift SET ShiftDate = '20170607' WHERE Id = 14424
--UPDATE dbo.Shift SET ReportDate = '20190831' WHERE Id = 26014

--UPDATE dbo.Shift SET IdDay = DAY(ReportDate) WHERE Id = 23622

SELECT TOP 15
       shf.Id
      ,shf.IdOperDate
      ,shf.IdDepartment
      ,usr.IdPost
      ,shf.IdUser
      ,usr.UserName
      ,shf.Status
      ,shf.IdPart
      ,shf.IdDay
      ,shf.ShiftDate
      ,shf.ReportDate
      ,shf.OpenDate
      ,shf.CloseDate
      ,shf.Blok
      ,shf.MestoCode
--      ,pst.PostName
      ,dpt.DepartmentName
  FROM dbo.Shift AS shf (READPAST)
  INNER JOIN dbo.[User] AS usr (READPAST) ON usr.Id = shf.IdUser
  INNER JOIN dbo.Post AS pst ON pst.Id = usr.IdPost
  INNER JOIN dbo.Department AS dpt ON dpt.Id = usr.IdDepartment
  --WHERE shf.IdDepartment = 42 --AND CloseDate IS NULL --AND shf.ReportDate = '05.01.2019'
  --WHERE shf.IdDepartment IN (1, 41)
  --WHERE shf.IdDepartment IN (41, 51)
  --WHERE shf.IdDepartment IN (2, 42)
  --WHERE shf.IdDepartment IN (6, 45) AND CloseDate IS NULL
  --WHERE shf.IdDepartment IN (21, 48) AND CloseDate IS NULL
  --WHERE shf.IdDepartment IN (18, 49) --AND CloseDate IS NULL
  WHERE shf.IdDepartment IN (1, 2, 41, 42, 21, 48) AND CloseDate IS NULL
  --WHERE shf.Status = 1
  --WHERE shf.IdDepartment = 6 --AND IdPart = 1
  --WHERE shf.Id IN (22059, 22071, 22080)
  --WHERE shf.Id = 26014
  --WHERE shf.IdDepartment = 2 AND shf.IdPart = 2
  --WHERE shf.IdDepartment = 2 AND shf.Status = 3 AND CloseDate IS NULL
  --WHERE shf.IdUser = 819
  --WHERE IdOperDate = 2859
  --WHERE YEAR(shf.ShiftDate) = 2015 AND MONTH(shf.ShiftDate) = 12 AND DAY(shf.ShiftDate) = 31 AND shf.IdDepartment = 1
  ORDER BY shf.Id DESC
  
  --WHERE shf.ReportDate <> shf.ShiftDate 
  
  GO

  --DECLARE @ReportDate DateTime
  --SET @ReportDate = '20191103 00:00:00'

  --DECLARE @IdPart Int
  ----SET @IdPart = 1

  --SELECT @IdPart = (SELECT COUNT(Id) FROM dbo.Shift WHERE ReportDate = @ReportDate AND IdDepartment = 1) + 1

  --SELECT * FROM dbo.Shift WHERE ReportDate = @ReportDate AND IdDepartment = 1
  --SELECT * FROM dbo.Shift WHERE ReportDate = @ReportDate AND IdPart = @IdPart AND IdDepartment = 1

  --GO
