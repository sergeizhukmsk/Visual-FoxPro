********************************************************************************* FUNCTION ChkShift ***************************************************************

FUNCTION ChkShift
PARAMETERS NoShiftOpn, PartShift

* WAIT WINDOW 'NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 3

* ------------------------------------------------------------------------- �������� ���������� ����������� �������� ���� ----------------------------------------------------------------------------------------------------- *

* ---------------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.IdPart, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Sh.Blok
  FROM Shift AS Sh
  WHERE Sh.IdDepartment = ?f_IdDepartmentKP AND Sh.Status = 3 AND Sh.CloseDate IS NULL AND  Sh.Id < ?f_IdShift_2
  ORDER BY Sh.Id DESC

ENDTEXT

IF RunSQL_Load('', @ar) <> 1
  MESSAGEBOX('������ SELECT TOP 1 Shift_Id ��� ��������� �������� ������� ����� ��', 16, '�������� ����� ��')
  RETURN .F.
ENDIF

* BROWSE WINDOW brows

colvo_zap = RECCOUNT()

IF colvo_zap = 1  && ������� ���������� ����� �� �������� 3

  MESSAGEBOX('������ � �� ������� ��� �����!!!' + CHR(13) + '��������� ��� ����� ���������!!!', 16, '������ ��� �������� ����� ����� ��')
  RETURN .F.

ELSE

  DO CASE
    CASE f_IdPart = 1 AND INLIST(PartShift, 1, 3) = .T.  && ����� - ����

      MESSAGEBOX('������ � �� ������� ����� ����!!!' + CHR(13) + '������� ����� ������ ����� ����!!!', 16, '������ ��� �������� ����� ����� ��')
      RETURN .F.

    CASE f_IdPart = 2 AND INLIST(PartShift, 1, 2) = .T.  && ����� - ����

      MESSAGEBOX('������ � �� ������� ����� ����!!!' + CHR(13) + '������� ����� ������ ����� ����!!!', 16, '������ ��� �������� ����� ����� ��')
      RETURN .F.

    CASE f_IdPart = 3 AND INLIST(PartShift, 2, 3) = .T.  && ����� - ����

      MESSAGEBOX('������ � �� ������� ����� ����!!!' + CHR(13) + '������� ����� ������ ����� ����!!!', 16, '������ ��� �������� ����� ����� ��')
      RETURN .F.

  ENDCASE

ENDIF

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO CASE
  CASE NoShiftOpn = .F.

    IF MESSAGEBOX('��������� ����� ����� �� ������ ������!!!' + CHR(13) + '������� ��������� ��!!!' + CHR(13) + ;
        '�� ������������� ������ ������� ����� �����' + IIF(PartShift = 1, ' ����', IIF(PartShift = 2, ' ����', IIF(PartShift = 3, ' ����', ''))) + '?', 33, '�������� ����� ����� ��') = 2
      RETURN .F.
    ENDIF

  CASE NoShiftOpn = .T.

    IF MESSAGEBOX('������������� ����� �� ������ ������!!!' + CHR(13) + '������� ��������� ��!!!' + CHR(13) + ;
        '�� ������������� ������ ����������� �����' + IIF(PartShift = 1, ' ����', IIF(PartShift = 2, ' ����', IIF(PartShift = 3, ' ����', ''))) + '?', 33, '������������ ����� ��') = 2
      RETURN .F.
    ENDIF

ENDCASE

f_PartShift = PartShift

* WAIT WINDOW 'NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 3
* WAIT WINDOW 'PartShift = ' + ALLTRIM(STR(PartShift)) TIMEOUT 3

DIMENSION ar[1,2]

rc = 0

* ����� ��������� �����

f_IdDepartmentPRK = GetGlob('DepartPRK')

f_IdDepartmentPRK_1 = GetGlob('DepartPRK_1')
f_IdDepartmentPRK_2 = GetGlob('DepartPRK_2')
f_IdDepartmentPRK_3 = GetGlob('DepartPRK_3')

*   WAIT WINDOW 'f_IdDepartmentKP = ' + ALLTRIM(STR(f_IdDepartmentKP)) TIMEOUT 3
*   WAIT WINDOW 'f_IdDepartmentPRK = ' + ALLTRIM(STR(f_IdDepartmentPRK)) TIMEOUT 3
*   WAIT WINDOW 'f_IdDepartmentPRK_1 = ' + ALLTRIM(STR(f_IdDepartmentPRK_1)) TIMEOUT 3
*   WAIT WINDOW 'f_IdDepartmentPRK_2 = ' + ALLTRIM(STR(f_IdDepartmentPRK_2)) TIMEOUT 3
*   WAIT WINDOW 'f_IdDepartmentPRK_3 = ' + ALLTRIM(STR(f_IdDepartmentPRK_3)) TIMEOUT 3

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT
   usr.Id
  ,usr.IdDepartment
  ,dpt.DepartmentName
  ,usr.IdPost
  ,pst.PostName
  ,usr.UserName
  ,usr.UserId
  ,usr.UserPassword
  ,usr.IsActive
  FROM [User] AS usr
  INNER JOIN Post AS pst ON pst.Id = usr.IdPost
  INNER JOIN Department AS dpt ON dpt.Id = usr.IdDepartment
  WHERE usr.Id = ?f_IdUser
  ORDER BY usr.Id DESC

ENDTEXT

rc = RunSql_Load('SpravUser', @ar)

* BROWSE WINDOW brows TITLE '������� ������ �� Id �� ����������� �������������'

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF rc = 1  && ������� ������ �� ���� SQL Server ������� User ����������� �������

  IF USED('SpravUser') = .T.

    SELECT SpravUser

    colvo_zap = RECCOUNT()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    DO CASE
      CASE colvo_zap = 0  && ���� ������� � �������� ������� �� �������

        =err('��������! � ����������� ������������� ������� �� ���������� ... ')
        DO exit_sql_server IN RunSQL_Load
        RETURN .F.

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE colvo_zap = 1  && ���� � �������� ������� ������� 1 ������

        f_PostName = ALLTRIM(SpravUser.PostName)
        f_IdPost = SpravUser.IdPost
        f_IdControler = SpravUser.Id

        IF sergeizhuk_flag = .T.
          f_IdPost = 11
        ENDIF

        IF f_IdPost <> 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) <> '������� ��������� ��'

          MESSAGEBOX( + ;
            '��������� ������������ - ' + ALLTRIM(SpravUser.UserName) + CHR(13) + ;
            '� ��� ��� ���������� ��������� ����� ����� �� !!!' + CHR(13) + ;
            '��������� ����������, ������� ��������� ����� �� - ������������' + CHR(13) + ;
            '���������� ����� �� ���� � ��� ��� ����� � ������� "������� ��������� ��"' + CHR(13) + ;
            '��������� ����� ����� ��������� !!!', 16, '�������� ���������� ���������� ��')

          DO exit_sql_server IN RunSQL_Load
          RETURN .F.

        ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE colvo_zap > 1  && ���� � �������� ������� ������� ����� 1 ������

        =err('��������! � ����������� ������������� ���������� ����� ����� ������ ... ')
        DO exit_sql_server IN RunSQL_Load
        RETURN .F.

    ENDCASE

    IF USED('SpravUser') = .T.
      SELECT SpravUser
      USE
    ENDIF

  ELSE
    =err('��������! ������� User �������������� �������� ... ')
    DO exit_sql_server IN RunSQL_Load
    RETURN .F.
  ENDIF

ELSE
  =err('��������! ������� ������ �� ���� SQL Server ������� User ����������� �������� ... ')
  DO exit_sql_server IN RunSQL_Load
  RETURN .F.
ENDIF

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO CASE
  CASE NoShiftOpn = .F.  && ������� �������� ����� �����

    DO CASE
      CASE PartShift = 1  && �������� ����� C���� �� - ����

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT TOP 1 Id, OpDate, Status, IdDepartment, OpenDate, CloseDate, LastOrder
          FROM OperDate (READPAST)
          WHERE IdDepartment IN (?f_IdDepartmentPRK, ?f_IdDepartmentPRK_1, ?f_IdDepartmentPRK_2, ?f_IdDepartmentPRK_3) AND Status IN (1, 2) AND CloseDate IS NULL
          ORDER BY Id ASC

        ENDTEXT


      CASE PartShift = 2  && �������� ����� C���� �� - ����

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT TOP 1 Id, OpDate, Status, IdDepartment, OpenDate, CloseDate, LastOrder
          FROM OperDate (READPAST)
          WHERE IdDepartment IN (?f_IdDepartmentPRK, ?f_IdDepartmentPRK_1, ?f_IdDepartmentPRK_2, ?f_IdDepartmentPRK_3) AND Status IN (1, 2) AND CloseDate IS NULL
          ORDER BY Id DESC

        ENDTEXT


      CASE PartShift = 3  && �������� ����� C���� �� - ����

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT TOP 1 Id, OpDate, Status, IdDepartment, OpenDate, CloseDate, LastOrder
          FROM OperDate (READPAST)
          WHERE IdDepartment IN (?f_IdDepartmentPRK, ?f_IdDepartmentPRK_1, ?f_IdDepartmentPRK_2, ?f_IdDepartmentPRK_3) AND Status IN (1, 2) AND CloseDate IS NULL
          ORDER BY Id DESC

        ENDTEXT

    ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE NoShiftOpn = .T.  && ������� �������������� ����� �������� �����

    DO CASE
      CASE PartShift = 1  && �������� ����� C���� �� - ����

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT TOP 1 Id, OpDate, Status, IdDepartment, OpenDate, CloseDate, LastOrder
          FROM OperDate (READPAST)
          WHERE IdDepartment IN (?f_IdDepartmentPRK, ?f_IdDepartmentPRK_1, ?f_IdDepartmentPRK_2, ?f_IdDepartmentPRK_3) AND Status = 0 AND CloseDate IS NOT NULL
          ORDER BY Id DESC

        ENDTEXT


      CASE PartShift = 2  && �������� ����� C���� �� - ����

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT TOP 1 Id, OpDate, Status, IdDepartment, OpenDate, CloseDate, LastOrder
          FROM OperDate (READPAST)
          WHERE IdDepartment IN (?f_IdDepartmentPRK, ?f_IdDepartmentPRK_1, ?f_IdDepartmentPRK_2, ?f_IdDepartmentPRK_3) AND Status IN (1, 2) AND CloseDate IS NULL
          ORDER BY Id DESC

        ENDTEXT


      CASE PartShift = 3  && �������� ����� C���� �� - ����

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

          SELECT TOP 1 Id, OpDate, Status, IdDepartment, OpenDate, CloseDate, LastOrder
          FROM OperDate (READPAST)
          WHERE IdDepartment IN (?f_IdDepartmentPRK, ?f_IdDepartmentPRK_1, ?f_IdDepartmentPRK_2, ?f_IdDepartmentPRK_3) AND Status = 0 AND CloseDate IS NOT NULL
          ORDER BY Id DESC

        ENDTEXT

    ENDCASE


ENDCASE

=RunSql_Load('CursOperDate', @ar)

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF USED('CursOperDate') = .T.

  SELECT CursOperDate

  colvo_zap = RECCOUNT()

  DO CASE
    CASE colvo_zap = 1  && ���� �������� ���� ������� �������, �� �������� �� ������������ ����

      f_IdOperDate = CursOperDate.Id

* WAIT WINDOW 'f_IdOperDate = ' + ALLTRIM(STR(f_IdOperDate)) TIMEOUT 3
* WAIT WINDOW 'PartShift = ' + ALLTRIM(STR(PartShift)) TIMEOUT 3

      DO CASE
        CASE NoShiftOpn = .F.  && ������� �������� ����� �����

          DO CASE
            CASE PartShift = 1  && �������� ����� C���� �� - ����

              TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

                SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.IdPart, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate
                FROM Shift AS Sh (READPAST)
                INNER JOIN OperDate AS Od (READPAST) ON Od.Id = Sh.IdOperDate
                WHERE Sh.IdDepartment = ?f_IdDepartmentKP AND Sh.Status = 1 AND Sh.IdPart = 3 AND Sh.CloseDate IS NULL
                ORDER BY Sh.Id DESC

              ENDTEXT


            CASE PartShift = 2  && �������� ����� C���� �� - ����

              TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

                SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.IdPart, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate
                FROM Shift AS Sh (READPAST)
                INNER JOIN OperDate AS Od (READPAST) ON Od.Id = Sh.IdOperDate
                WHERE Sh.IdDepartment = ?f_IdDepartmentKP AND Sh.Status = 1 AND Sh.IdPart = 1 AND Sh.CloseDate IS NULL
                ORDER BY Sh.Id DESC

              ENDTEXT


            CASE PartShift = 3  && �������� ����� C���� �� - ����

              TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

                SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.IdPart, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate
                FROM Shift AS Sh (READPAST)
                INNER JOIN OperDate AS Od (READPAST) ON Od.Id = Sh.IdOperDate
                WHERE Sh.IdDepartment = ?f_IdDepartmentKP AND Sh.Status IN (1, 2) AND Sh.IdPart = 2 AND Sh.CloseDate IS NULL
                ORDER BY Sh.Id DESC

              ENDTEXT

          ENDCASE

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        CASE NoShiftOpn = .T.  && ������� �������������� ����� �������� �����

* WAIT WINDOW 'f_IdOperDate = ' + ALLTRIM(STR(f_IdOperDate)) TIMEOUT 3
* WAIT WINDOW 'PartShift = ' + ALLTRIM(STR(PartShift)) TIMEOUT 3

          TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.IdPart, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate
            FROM Shift AS Sh (READPAST)
            INNER JOIN OperDate AS Od (READPAST) ON Od.Id = Sh.IdOperDate
            WHERE Od.Id = ?f_IdOperDate AND Sh.IdDepartment = ?f_IdDepartmentKP AND Sh.Status IN (0, 3) AND Sh.IdPart = ?PartShift
            ORDER BY Sh.Id DESC

          ENDTEXT

      ENDCASE

      =RunSql_Load('CursOperDate', @ar)

* BROWSE WINDOW brows

      IF USED('CursOperDate') = .T.

        SELECT CursOperDate

        colvo_zap = RECCOUNT()

        DO CASE
          CASE colvo_zap > 0  && ���� �������� ���� ������� � �������� ����� ��� �������, �� ��������� � �������� ����� ��

            f_IdShift_2 = CursOperDate.Id

* ----------------------------------------------- *
            DO Start_Shift IN ChkShift
* ----------------------------------------------- *

*          WAIT WINDOW 'NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 5

          CASE colvo_zap = 0  && ���� �������� ����� �� �� �������

* ----------------------------------------------- *
            DO Start_Shift IN ChkShift
* ----------------------------------------------- *

        ENDCASE

      ELSE
        =soob('��������! � �������������� ������� �� ���� SQL Server CursOperDate ������� �� ���������� ...')
      ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE colvo_zap > 1  && ���� �������� ��� �������� ������� ���������, �� ������� ��������� ���������

      MESSAGEBOX('� ����������� �������� ������� ��������� �������� ���' + CHR(13) + '��������� � ��� ����� ��� ��������� ����� ����:' + CHR(13) + '����������� ������� �������� �� ������� � ���', 16, '�������� ����� ����� � ��')

      MESSAGEBOX('����� ���������� ����� � ���, ��������� ��������:' + CHR(13) + '�������� ����� ����� � ��', 64, '�������� ����� ����� � ��')

  ENDCASE

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF USED('CursShift') = .T.
    USE IN CursShift
  ENDIF

  IF USED('CursOperDate') = .T.
    USE IN CursOperDate
  ENDIF

ELSE
  =soob('��������! ������� ������ �� ���� SQL Server ������� CursOperDate ����������� �������� ... ')
ENDIF

* WAIT WINDOW 'NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 5

RETURN NoShiftOpn


************************************************************************* PROCEDURE Start_Shift *******************************************************************

PROCEDURE Start_Shift

* BROWSE WINDOW brows

IF f_IdPart = 3

  f_ReportDate = TTOD(CursOperDate.OpenDate)
  f_IdDay = DAY(f_ReportDate)

ELSE

  f_ReportDate = IIF(TTOD(CursOperDate.OpenDate) < DATE(), DATE(), TTOD(CursOperDate.OpenDate))
  f_IdDay = DAY(f_ReportDate)

ENDIF

f_OpenDate = IIF(TTOD(CursOperDate.OpenDate) < DATE(), DATE(), TTOD(CursOperDate.OpenDate))

* WAIT WINDOW 'f_IdDepartmentKP = ' + ALLTRIM(STR(f_IdDepartmentKP)) + '   f_IdOperDate = ' + ALLTRIM(STR(f_IdOperDate)) TIMEOUT 3
* WAIT WINDOW 'f_IdOperDate = ' + ALLTRIM(STR(f_IdOperDate)) TIMEOUT 3
* WAIT WINDOW 'NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 3

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7
  BEGIN TRANSACTION
ENDTEXT

=RunSql_Load('', @ar)

* WAIT WINDOW 'NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 3

DO CASE
  CASE NoShiftOpn = .F.  && ������� �������� ����� �����

* WAIT WINDOW '������� �������� ����� �����' TIMEOUT 3
* WAIT WINDOW 'CursOperDate_OpenDate = ' + DTOC(TTOD(CursOperDate.OpenDate)) TIMEOUT 3
* WAIT WINDOW 'f_OpenDate = ' + DTOC(f_OpenDate) TIMEOUT 3

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.IdPart, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate, Sh.Blok
      FROM Shift AS Sh (READPAST)
      INNER JOIN OperDate AS Od (READPAST) ON Od.Id = Sh.IdOperDate
      WHERE Sh.IdDepartment = ?f_IdDepartmentKP AND Sh.CloseDate IS NULL
      ORDER BY Sh.Id DESC

    ENDTEXT

    rc = RunSql_Load('NewCursShift', @ar)

*  WAIT WINDOW 'IDENTITY - rc = ' + ALLTRIM(STR(rc)) TIMEOUT 3

* BROWSE WINDOW brows

    IF USED('NewCursShift') = .T. AND RECCOUNT() > 1

      MESSAGEBOX('� ����������� ���� ������� ����� ����� �������� ���� ��' + CHR(13) + '��������� ����� ����� ��������� !!!' + CHR(13) + '���������� ������� ���������� �����.', 16, '�������� ���������� �������� ���� � ��')

      SELECT NewCursShift
      GO TOP
*      BROWSE WINDOW brows
      USE
      RETURN .F.

    ENDIF

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW 'f_IdDepartmentKP = ' + ALLTRIM(STR(f_IdDepartmentKP)) + '   f_IdDepartmentPRK = ' + ALLTRIM(STR(f_IdDepartmentPRK)) TIMEOUT 3

    DO CASE
      CASE f_PartShift = 1  && �������� ����� C���� �� - ����

* WAIT WINDOW '�������� ����� C���� �� - ����' TIMEOUT 3

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

            DECLARE @IdOperDate_3 Int, @IdShift_3 Int

            SET @IdOperDate_3 = 0
            SET @IdShift_3 = 0

            SELECT TOP 1 @IdOperDate_3 = Sh.IdOperDate, @IdShift_3 = Sh.Id
            FROM Shift AS Sh (READPAST)
            WHERE Sh.IdDepartment = ?f_IdDepartmentKP AND Sh.Status = 1 AND Sh.IdPart = 3 AND Sh.CloseDate IS NULL
            ORDER BY Sh.Id DESC

            BEGIN

            IF @IdOperDate_3 > 0 AND @IdShift_3 > 0

              UPDATE Shift SET
              Status = 3,
              Blok = 1
              WHERE IdOperDate = @IdOperDate_3 AND Id = @IdShift_3

              IF EXISTS(SELECT TOP 1 Id FROM Shift (READPAST) WHERE IdOperDate = @IdOperDate_3 AND Id = @IdShift_3 AND Status = 3 AND CloseDate IS NULL) AND
              NOT EXISTS(SELECT TOP 1 Id FROM Shift (READPAST) WHERE IdOperDate = ?f_IdOperDate AND IdDepartment = ?f_IdDepartmentKP AND Status = 1 AND IdPart = 1 AND CloseDate IS NULL)

              INSERT INTO Shift
              (IdOperDate, IdDepartment, IdUser, Status, IdPart, IdDay, ShiftDate, ReportDate, OpenDate, CloseDate, Blok, MestoCode)
              VALUES
              (?f_IdOperDate, ?f_IdDepartmentKP, ?f_IdUser, 1, 1, ?f_IdDay, ?f_OpenDate, ?f_ReportDate, GetDate(), NULL, 1, ?f_MestoCode)

            END

            BEGIN

             IF @IdOperDate_3 = 0 AND @IdShift_3 = 0

              INSERT INTO Shift
              (IdOperDate, IdDepartment, IdUser, Status, IdPart, IdDay, ShiftDate, ReportDate, OpenDate, CloseDate, Blok, MestoCode)
              VALUES
              (?f_IdOperDate, ?f_IdDepartmentKP, ?f_IdUser, 1, 1, ?f_IdDay, ?f_OpenDate, ?f_ReportDate, GetDate(), NULL, 0, ?f_MestoCode)

             END

            IF EXISTS(SELECT TOP 1 Id FROM Shift (READPAST) WHERE Id = @@IDENTITY)

              SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.IdPart, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate, Sh.Blok, sh.MestoCode
              FROM Shift AS Sh (READPAST)
              INNER JOIN OperDate AS Od (READPAST) ON Od.Id = Sh.IdOperDate
              WHERE Sh.Id = @@IDENTITY AND Sh.IdPart = 1 AND Sh.CloseDate IS NULL
              ORDER BY Sh.Id DESC

        ENDTEXT

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE f_PartShift = 2  && �������� ����� C���� �� - ����

* WAIT WINDOW '�������� ����� C���� �� - ����' TIMEOUT 3

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

            DECLARE @IdOperDate_3 Int, @IdShift_3 Int

            SET @IdOperDate_3 = 0
            SET @IdShift_3 = 0

            SELECT TOP 1 @IdOperDate_3 = Sh.IdOperDate, @IdShift_3 = Sh.Id
            FROM Shift AS Sh (READPAST)
            WHERE Sh.IdDepartment = ?f_IdDepartmentKP AND Sh.Status = 1 AND Sh.IdPart = 1 AND Sh.CloseDate IS NULL
            ORDER BY Sh.Id DESC

            BEGIN

            IF @IdOperDate_3 > 0 AND @IdShift_3 > 0

              UPDATE Shift SET
              Status = 3,
              Blok = 1
              WHERE IdOperDate = @IdOperDate_3 AND Id = @IdShift_3

            IF EXISTS(SELECT TOP 1 Id FROM Shift (READPAST) WHERE IdOperDate = @IdOperDate_3 AND Id = @IdShift_3 AND Status = 3 AND CloseDate IS NULL) AND
            NOT EXISTS(SELECT TOP 1 Id FROM Shift (READPAST) WHERE IdOperDate = ?f_IdOperDate AND IdDepartment = ?f_IdDepartmentKP AND Status = 1 AND IdPart = 2 AND CloseDate IS NULL)

              INSERT INTO Shift
              (IdOperDate, IdDepartment, IdUser, Status, IdPart, IdDay, ShiftDate, ReportDate, OpenDate, CloseDate, Blok, MestoCode)
              VALUES
              (?f_IdOperDate, ?f_IdDepartmentKP, ?f_IdUser, 1, 2, ?f_IdDay, ?f_OpenDate, ?f_ReportDate, GetDate(), NULL, 1, ?f_MestoCode)

             END

            BEGIN

            IF @IdOperDate_3 = 0 AND @IdShift_3 = 0

              INSERT INTO Shift
              (IdOperDate, IdDepartment, IdUser, Status, IdPart, IdDay, ShiftDate, ReportDate, OpenDate, CloseDate, Blok, MestoCode)
              VALUES
              (?f_IdOperDate, ?f_IdDepartmentKP, ?f_IdUser, 1, 2, ?f_IdDay, ?f_OpenDate, ?f_ReportDate, GetDate(), NULL, 0, ?f_MestoCode)

             END

            IF EXISTS(SELECT TOP 1 Id FROM Shift (READPAST) WHERE Id = @@IDENTITY)

              SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.IdPart, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate, Sh.Blok, sh.MestoCode
              FROM Shift AS Sh (READPAST)
              INNER JOIN OperDate AS Od (READPAST) ON Od.Id = Sh.IdOperDate
              WHERE Sh.Id = @@IDENTITY AND Sh.IdPart = 2 AND Sh.CloseDate IS NULL
              ORDER BY Sh.Id DESC

        ENDTEXT

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE f_PartShift = 3  && �������� ����� C���� �� - ����

* WAIT WINDOW '�������� ����� C���� �� - ����' TIMEOUT 3

        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

            DECLARE @IdOperDate_3 Int, @IdShift_3 Int

            SET @IdOperDate_3 = 0
            SET @IdShift_3 = 0

            SELECT TOP 1 @IdOperDate_3 = Sh.IdOperDate, @IdShift_3 = Sh.Id
            FROM Shift AS Sh (READPAST)
            WHERE Sh.IdDepartment = ?f_IdDepartmentKP AND Sh.Status = 1 AND Sh.IdPart = 2 AND Sh.CloseDate IS NULL
            ORDER BY Sh.Id DESC

            BEGIN

            IF @IdOperDate_3 > 0 AND @IdShift_3 > 0

              UPDATE Shift SET
              Status = 3,
              Blok = 1
              WHERE IdOperDate = @IdOperDate_3 AND Id = @IdShift_3

              IF EXISTS(SELECT TOP 1 Id FROM Shift (READPAST) WHERE IdOperDate = @IdOperDate_3 AND Id = @IdShift_3 AND Status = 3 AND CloseDate IS NULL) AND
              NOT EXISTS(SELECT TOP 1 Id FROM Shift (READPAST) WHERE IdOperDate = ?f_IdOperDate AND IdDepartment = ?f_IdDepartmentKP AND Status = 1 AND IdPart = 3 AND CloseDate IS NULL)

              INSERT INTO Shift
              (IdOperDate, IdDepartment, IdUser, Status, IdPart, IdDay, ShiftDate, ReportDate, OpenDate, CloseDate, Blok, MestoCode)
              VALUES
              (?f_IdOperDate, ?f_IdDepartmentKP, ?f_IdUser, 1, 3, ?f_IdDay, ?f_OpenDate, ?f_ReportDate, GetDate(), NULL, 1, ?f_MestoCode)

             END

            BEGIN

            IF @IdOperDate_3 = 0 AND @IdShift_3 = 0

              INSERT INTO Shift
              (IdOperDate, IdDepartment, IdUser, Status, IdPart, IdDay, ShiftDate, ReportDate, OpenDate, CloseDate, Blok, MestoCode)
              VALUES
              (?f_IdOperDate, ?f_IdDepartmentKP, ?f_IdUser, 1, 3, ?f_IdDay, ?f_OpenDate, ?f_ReportDate, GetDate(), NULL, 0, ?f_MestoCode)

             END

            IF EXISTS(SELECT TOP 1 Id FROM Shift (READPAST) WHERE Id = @@IDENTITY)

              SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.IdPart, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate, Sh.Blok, sh.MestoCode
              FROM Shift AS Sh (READPAST)
              INNER JOIN OperDate AS Od (READPAST) ON Od.Id = Sh.IdOperDate
              WHERE Sh.Id = @@IDENTITY AND Sh.IdPart = 3 AND Sh.CloseDate IS NULL
              ORDER BY Sh.Id DESC

        ENDTEXT

    ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE NoShiftOpn = .T.  && ������� �������������� ����� �������� �����

*  WAIT WINDOW 'f_IdShift_2 = ' + ALLTRIM(STR(f_IdShift_2)) TIMEOUT 3
*  WAIT WINDOW 'f_IdDepartmentKP = ' + ALLTRIM(STR(f_IdDepartmentKP)) TIMEOUT 3
*  WAIT WINDOW 'PartShift = ' + ALLTRIM(STR(PartShift)) TIMEOUT 3
*  WAIT WINDOW 'f_IdPart = ' + ALLTRIM(STR(f_IdPart)) + '   PartShift = ' + ALLTRIM(STR(PartShift)) TIMEOUT 3

    IF f_IdPart = PartShift

      MESSAGEBOX('��������� ������������!' + CHR(13) + '����������� ����� �� ����� ���� ���� ����� � �����.', 16, '�������������� ����� �������� ����� ��')
      NoShiftOpn = .F.
      RETURN

    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*      put_select = 1

*      tex1 = '�������� � ������ ����������� ����� �������, ���������� ����� ����������, � ��� �������� ����� ��������'
*      tex2 = '�������� � ������ ����������� ����� �������, ���������� ����� �������, � ��� �������� ����� ��������'
*      l_bar = 3
*      =popup_9(tex1, tex2, tex3, tex4, l_bar)

*      ACTIVATE POPUP vibor BAR put_select

*      RELEASE POPUPS vibor

*      IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

*        arm_version = IIF(BAR() >= 1, BAR(), 1)

*        DO CASE
*          CASE put_select = 1  && �������� � ������ ����������� ����� �������, ���������� ����� ����������, � ��� �������� ����� ��������


*  * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*          CASE put_select = 3  && �������� � ������ ����������� ����� �������, ���������� ����� �������, � ��� �������� ����� ��������


*        ENDCASE

*      ENDIF

*  * ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

        IF EXISTS(SELECT TOP 1 Id FROM Shift (READPAST) WHERE IdDepartment = ?f_IdDepartmentKP AND Status = 0 AND IdPart = ?PartShift AND CloseDate IS NOT NULL)

          UPDATE Shift SET
          IdUser = ?f_IdUser,
          Status = 1,
          Blok = 1,
          CloseDate = NULL
          WHERE Id = ?f_IdShift_2 AND IdDepartment = ?f_IdDepartmentKP AND Status IN (0, 3) AND IdPart = ?PartShift

          UPDATE Shift SET
          IdUser = ?f_IdUser,
          Status = 0,
          Blok = 1,
          CloseDate = GetDate()
          WHERE Id > ?f_IdShift_2 AND IdDepartment = ?f_IdDepartmentKP AND Status = 1

          SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.IdPart, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate, Sh.Blok, sh.MestoCode
          FROM Shift AS Sh (READPAST)
          INNER JOIN OperDate AS Od (READPAST) ON Od.Id = Sh.IdOperDate
          WHERE Sh.Id = ?f_IdShift_2 AND Sh.IdDepartment = ?f_IdDepartmentKP AND IdPart = ?PartShift
          ORDER BY Sh.Id DESC

    ENDTEXT

ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

rc = RunSql_Load('NewCursShift', @ar)

* WAIT WINDOW 'rc = ' + ALLTRIM(STR(rc)) TIMEOUT 3

* BROWSE WINDOW brows

IF USED('NewCursShift') = .T.

  SELECT NewCursShift
  GO TOP

* BROWSE WINDOW brows

ENDIF

DO CASE
  CASE rc = 1

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7
      COMMIT TRANSACTION
    ENDTEXT

    =RunSql_Load('', @ar)

  CASE rc <> 1

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7
      ROLLBACK TRANSACTION
    ENDTEXT

    =RunSql_Load('', @ar)

    RETURN

ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*  BROWSE WINDOW brows

IF USED('NewCursShift') = .T.

  SELECT NewCursShift

  colvo_zap = RECCOUNT()

  GO TOP

* BROWSE WINDOW brows

  IF colvo_zap > 0

    SELECT NewCursShift

    GO TOP

* BROWSE WINDOW brows

    f_IdOperDate = NewCursShift.IdOperDate
    f_IdShift_2 = NewCursShift.Id

    f_ShiftDate = NewCursShift.ShiftDate
    f_OpenDate = NewCursShift.OpenDate
    f_OpDate = NewCursShift.OpDate

    DO CASE
      CASE NewCursShift.Status = 1
        NoShiftOpn = .T.

      CASE NewCursShift.Status = 2
        NoShiftOpn = .T.

      CASE NewCursShift.Status = 0
        NoShiftOpn = .F.

    ENDCASE

* WAIT WINDOW 'NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 3

  ENDIF

ELSE
  =soob('��������! � �������������� ������� �� ���� SQL Server CursOperDate ������� �� ���������� ... ')
ENDIF

* -------------------------------------------------------- *
DO exit_sql_server IN RunSQL_Load
* -------------------------------------------------------- *

* -------------------------------------------------------------------------------------------------------------------------------------------------- *

IF PartShift = 3  && �������� ����� C���� �� - ����  &&  AND NOT INLIST(DOW(DATE()), 7, 1)
  DO start IN ReadShift_PRK
ENDIF

* -------------------------------------------------------------------------------------------------------------------------------------------------- *

ShiftDate_ = f_ShiftDate
OpenDate_ = f_OpenDate

skip_data_odb = TTOD(f_OpDate)
new_data_odb = TTOD(f_ShiftDate)

title_program_full = SPACE(5) + ALLTRIM(title_program) + '  ���� ����� �� - ' + DT(new_data_odb) + '  ���� ������� - ' + DT(skip_data_odb) + '  ' + ALLTRIM(f_NameSmena) + f_MestoCodeName + '  ����� ����� �� = ' + ALLTRIM(STR(f_IdShift_2))

_SCREEN.CAPTION = ALLTRIM(title_program_full)

RETURN


*************************************************************************************************************************************************************************************************


