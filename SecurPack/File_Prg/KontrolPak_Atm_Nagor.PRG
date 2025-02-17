*********************************************************************** PROCEDURE start *****************************************************

PROCEDURE start
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

DIMENSION ar[1,2]

text_titul_prn_sost_0 = '������ �� ������ � ��������� ��� �� ��������� �������� ���������� ��������'
text_titul_prn_sost_1 = '������ �� ������ � ��������� ��� ��������� �������� ���������� ��������'

f_IdDepartmentPRK = GetGlob('DepartPRK')  && ����� ������������� ��� �� ����������� Department

* WAIT WINDOW 'f_IdDepartmentPRK = ' + ALLTRIM(STR(f_IdDepartmentPRK)) TIMEOUT 3

put_select = 1

dtB = new_data_odb  && ���� �������� ����� � ��
f_CurDate = new_data_odb  && ���� �������� ����� � ��

* WAIT WINDOW 'dtB = ' + DTOC(dtB) TIMEOUT 3
* WAIT WINDOW '1 - f_CurDate = ' + DTOC(f_CurDate) TIMEOUT 3

f_FormName = '����������'

* ------------------------------------ *
DO FORM GetDate1
* ------------------------------------ *

* ---------------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  f_CurDate = dtb

* WAIT WINDOW '2 - f_CurDate = ' + DTOC(f_CurDate) TIMEOUT 3

* ---------------------------------------- ������� ����� ��� �� ���� f_DtBegin ------------------------------------- *

  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

    SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.ShiftDate, Sh.ReportDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate, Od.OpenDate
    FROM Shift AS Sh (READPAST)
    INNER JOIN OperDate AS Od ON Od.Id = Sh.IdOperDate
    WHERE Sh.ReportDate = ?f_CurDate AND Sh.IdDepartment = ?f_IdDepartmentPRK
    ORDER BY Sh.Id DESC

  ENDTEXT

  IF RunSql_Load('Shift_PRK', @ar) <> 1
    DO exit_sql_server IN RunSQL_Load
    MESSAGEBOX('������ ������ �� ������������ ���� ... ', 16, '����� �� �������.')
    RETURN -1
  ENDIF

  IF USED('Shift_PRK') = .T. AND RECCOUNT('Shift_PRK') > 0

    SELECT Shift_PRK
    GO TOP

* BROWSE WINDOW brows

    OperDatePRK = TTOD(Shift_PRK.OpDate)  && ���� ��������� �������
    IdOperDatePRK = Shift_PRK.IdOperDate  && Id ��������� �������
    IdShiftPRK = Shift_PRK.Id
    ReportDateSmenaPRK = Shift_PRK.ReportDate && ���� �������� �������� ����� ��� ��� ���� �������.

    f_OperDate_OpDatePRK = OperDatePRK  && ���� ��������� ������������� ���, �� ������� OperDate
    f_IdOperDatePRK = IdOperDatePRK  && Id ��������� ������������� ���, �� ������� Shift

  ELSE
    =err('��������! �� ��������� ���� ��������� ���� ��� � ����������� �� ���������� ... ')
    RETURN
  ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  f_IdSignProc = RunSignProc(PlacePRK)  && SignProc.Id = 15

* WAIT WINDOW 'f_IdSignProc = ' + ALLTRIM(STR(f_IdSignProc)) TIMEOUT 3
* WAIT WINDOW 'f_IdShift_1 = ' + ALLTRIM(STR(f_IdShift_1)) TIMEOUT 3
*  WAIT WINDOW 'IdShift_PRK = ' + ALLTRIM(IdShift_PRK) TIMEOUT 3

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  DO WHILE .T.

    tex1 = '������� ������ �� ������ �����'
    tex2 = '������� ������ �� ��� �������'
    tex3 = '������� ������ �� �������� �������'
    tex4 = '������� ������ �� ������ �������'
    tex5 = '������� ������ �� ����� �����'
    tex6 = '������� ������ �� �������� �����'
    tex7 = '������� ������ �� �������� ������'
    tex8 = '������� ������ � ������ ������ �� ����'
    l_bar = 15
    =popup_big(tex1, tex2, tex3, tex4, tex5, tex6, tex7, tex8, l_bar)

    ACTIVATE POPUP vibor BAR put_select
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

      put_select = IIF(BAR() <= 0, 1, BAR())

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      DO CASE
        CASE put_select = 1  && ������� ������ �� ������ �����

          f_BudgetNum = ''
*          f_BudgetNum = 'AC709137'

          DO FORM Vibor_sumka_client

          IF NOT LASTKEY() = 27 AND pr_lastkey_27 = .T.

* WAIT WINDOW '����� ����� = ' + ALLTRIM(f_BudgetNum) TIMEOUT 1

            DO CASE
              CASE f_TypeNumer = 0 && � ������� ����� ������ ������ ������ ����� �������� ������ ��� �����

                TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                  SELECT trn.BudgetNum, trn.FilialCode, trn.WorkType, trn.SuperBudget, trn.Sost, trn.DateInkas, trn.Kurs_USD, trn.Kurs_EUR, trn.Comment, trn.IsTransit,
                  trn.TotalSum_RUB_Out, trn.TotalSum_USD_Out, trn.TotalSum_EUR_Out,
                  trn.TotalSum_RUB_In, trn.TotalSum_USD_In, trn.TotalSum_EUR_In,
                  trn.Id, trn.IdSecurPack, trn.IdShift, trn.IdUser, trn.IdRoute, trn.TabNumber, trn.DtIncome, trn.CurDate, trn.DtRecord,
                  fil.ClientFullName, fil.FilialName, fil.FlgRecAlc, fil.BankName, fil.BIC_b AS Bik, fil.Account_b AS Korschet, fil.PostAddress, fil.FactAddress, fil.UIFilialName, fil.UIFactAddress, fil.MestoCode,
                  trn.IdKontrol, trn.KontrolDate, trn.IsActive, trn.IsDriver
                  FROM Tranzit_ATM_PER AS trn (READPAST)
                  INNER JOIN View_Client_Filial_Bank AS fil (READPAST) ON fil.FilialCode = trn.FilialCode
                  INNER JOIN Shift AS shf (READPAST) ON shf.Id = trn.IdShift
                  WHERE shf.ReportDate = ?ReportDateSmenaPRK AND trn.BudgetNum = ?f_BudgetNum AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
                  ORDER BY trn.FilialCode, trn.WorkType, trn.BudgetNum

                ENDTEXT

*                  WHERE trn.IdShift IN (<<IdShift_PRK>>) AND trn.BudgetNum = ?f_BudgetNum AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode


              CASE f_TypeNumer = 1 && � ������� ����� ������ ������� ����� ������ �������� ������ ��� �����

                TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                  SELECT trn.BudgetNum, trn.FilialCode, trn.WorkType, trn.SuperBudget, trn.Sost, trn.DateInkas, trn.Kurs_USD, trn.Kurs_EUR, trn.Comment, trn.IsTransit,
                  trn.TotalSum_RUB_Out, trn.TotalSum_USD_Out, trn.TotalSum_EUR_Out,
                  trn.TotalSum_RUB_In, trn.TotalSum_USD_In, trn.TotalSum_EUR_In,
                  trn.Id, trn.IdSecurPack, trn.IdShift, trn.IdUser, trn.IdRoute, trn.TabNumber, trn.DtIncome, trn.CurDate, trn.DtRecord,
                  fil.ClientFullName, fil.FilialName, fil.FlgRecAlc, fil.BankName, fil.BIC_b AS Bik, fil.Account_b AS Korschet, fil.PostAddress, fil.FactAddress, fil.UIFilialName, fil.UIFactAddress, fil.MestoCode,
                  trn.IdKontrol, trn.KontrolDate, trn.IsActive, trn.IsDriver
                  FROM Tranzit_ATM_PER AS trn (READPAST)
                  INNER JOIN View_Client_Filial_Bank AS fil (READPAST) ON fil.FilialCode = trn.FilialCode
                  INNER JOIN Shift AS shf (READPAST) ON shf.Id = trn.IdShift
                  WHERE shf.ReportDate = ?ReportDateSmenaPRK AND trn.BudgetNum LIKE '%<<f_BudgetNum>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
                  ORDER BY trn.FilialCode, trn.WorkType, trn.BudgetNum

                ENDTEXT

            ENDCASE

*                  WHERE trn.IdShift IN (<<IdShift_PRK>>) AND trn.BudgetNum LIKE '%<<f_BudgetNum>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode

          ELSE
            LOOP
          ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        CASE put_select = 3  && ������� ������ �� ��� �������

          f_FilialCode = ''
*          f_FilialCode = '66-157'

          DO FORM Vibor_pin_client

          IF NOT LASTKEY() = 27 AND pr_lastkey_27 = .T.

* WAIT WINDOW '��� ������� = ' + ALLTRIM(f_FilialCode) TIMEOUT 1

            TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

              SELECT trn.BudgetNum, trn.FilialCode, trn.WorkType, trn.SuperBudget, trn.Sost, trn.DateInkas, trn.Kurs_USD, trn.Kurs_EUR, trn.Comment, trn.IsTransit,
              trn.TotalSum_RUB_Out, trn.TotalSum_USD_Out, trn.TotalSum_EUR_Out,
              trn.TotalSum_RUB_In, trn.TotalSum_USD_In, trn.TotalSum_EUR_In,
              trn.Id, trn.IdSecurPack, trn.IdShift, trn.IdUser, trn.IdRoute, trn.TabNumber, trn.DtIncome, trn.CurDate, trn.DtRecord,
              fil.ClientFullName, fil.FilialName, fil.FlgRecAlc, fil.BankName, fil.BIC_b AS Bik, fil.Account_b AS Korschet, fil.PostAddress, fil.FactAddress, fil.UIFilialName, fil.UIFactAddress, fil.MestoCode,
              trn.IdKontrol, trn.KontrolDate, trn.IsActive, trn.IsDriver
              FROM Tranzit_ATM_PER AS trn (READPAST)
              INNER JOIN View_Client_Filial_Bank AS fil (READPAST) ON fil.FilialCode = trn.FilialCode
              INNER JOIN Shift AS shf (READPAST) ON shf.Id = trn.IdShift
              WHERE shf.ReportDate = ?ReportDateSmenaPRK AND trn.FilialCode = ?f_FilialCode AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
              ORDER BY trn.FilialCode, trn.WorkType, trn.BudgetNum

            ENDTEXT

*              WHERE trn.IdShift IN (<<IdShift_PRK>>) AND trn.FilialCode = ?f_FilialCode AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode

          ELSE
            LOOP
          ENDIF


* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        CASE put_select = 5  && ������� ������ �� �������� �������

          f_FilialName = ''
*          f_FilialName = '���'

          DO FORM Vibor_name_client

          IF NOT LASTKEY() = 27 AND pr_lastkey_27 = .T.

* WAIT WINDOW '����������� ������� = ' + ALLTRIM(f_FilialName) TIMEOUT 1

            TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

              SELECT trn.BudgetNum, trn.FilialCode, trn.WorkType, trn.SuperBudget, trn.Sost, trn.DateInkas, trn.Kurs_USD, trn.Kurs_EUR, trn.Comment, trn.IsTransit,
              trn.TotalSum_RUB_Out, trn.TotalSum_USD_Out, trn.TotalSum_EUR_Out,
              trn.TotalSum_RUB_In, trn.TotalSum_USD_In, trn.TotalSum_EUR_In,
              trn.Id, trn.IdSecurPack, trn.IdShift, trn.IdUser, trn.IdRoute, trn.TabNumber, trn.DtIncome, trn.CurDate, trn.DtRecord,
              fil.ClientFullName, fil.FilialName, fil.FlgRecAlc, fil.BankName, fil.BIC_b AS Bik, fil.Account_b AS Korschet, fil.PostAddress, fil.FactAddress, fil.UIFilialName, fil.UIFactAddress, fil.MestoCode,
              trn.IdKontrol, trn.KontrolDate, trn.IsActive, trn.IsDriver
              FROM Tranzit_ATM_PER AS trn (READPAST)
              INNER JOIN View_Client_Filial_Bank AS fil (READPAST) ON fil.FilialCode = trn.FilialCode
              INNER JOIN Shift AS shf (READPAST) ON shf.Id = trn.IdShift
              WHERE shf.ReportDate = ?ReportDateSmenaPRK AND fil.FilialName LIKE '%<<f_FilialName>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
              ORDER BY trn.FilialCode, trn.WorkType, trn.BudgetNum

            ENDTEXT

*              WHERE trn.IdShift IN (<<IdShift_PRK>>) AND fil.FilialName LIKE '%<<f_FilialName>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode

          ELSE
            LOOP
          ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        CASE put_select = 7  && ������� ������ �� ������ �������

          f_FilialAddress = ''
          f_PostAddress = ''
          f_FactAddress = ''
          f_UIFactAddress = ''

          DO FORM Vibor_adress_client

          IF NOT LASTKEY() = 27 AND pr_lastkey_27 = .T.

            DO CASE
              CASE f_FilialAddress = 1 AND EMPTY(ALLTRIM(f_PostAddress)) = .F.

                TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                  SELECT trn.BudgetNum, trn.FilialCode, trn.WorkType, trn.SuperBudget, trn.Sost, trn.DateInkas, trn.Kurs_USD, trn.Kurs_EUR, trn.Comment, trn.IsTransit,
                  trn.TotalSum_RUB_Out, trn.TotalSum_USD_Out, trn.TotalSum_EUR_Out,
                  trn.TotalSum_RUB_In, trn.TotalSum_USD_In, trn.TotalSum_EUR_In,
                  trn.Id, trn.IdSecurPack, trn.IdShift, trn.IdUser, trn.IdRoute, trn.TabNumber, trn.DtIncome, trn.CurDate, trn.DtRecord,
                  fil.ClientFullName, fil.FilialName, fil.FlgRecAlc, fil.BankName, fil.BIC_b AS Bik, fil.Account_b AS Korschet, fil.PostAddress, fil.FactAddress, fil.UIFilialName, fil.UIFactAddress, fil.MestoCode,
                  trn.IdKontrol, trn.KontrolDate, trn.IsActive, trn.IsDriver
                  FROM Tranzit_ATM_PER AS trn (READPAST)
                  INNER JOIN View_Client_Filial_Bank AS fil (READPAST) ON fil.FilialCode = trn.FilialCode
                  INNER JOIN Shift AS shf (READPAST) ON shf.Id = trn.IdShift
                  WHERE shf.ReportDate = ?ReportDateSmenaPRK AND fil.PostAddress LIKE '%<<f_PostAddress>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
                  ORDER BY trn.FilialCode, trn.WorkType, trn.BudgetNum

                ENDTEXT

*                  WHERE trn.IdShift IN (<<IdShift_PRK>>) AND fil.PostAddress LIKE '%<<f_PostAddress>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode


              CASE f_FilialAddress = 2 AND EMPTY(ALLTRIM(f_FactAddress)) = .F.

                TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                  SELECT trn.BudgetNum, trn.FilialCode, trn.WorkType, trn.SuperBudget, trn.Sost, trn.DateInkas, trn.Kurs_USD, trn.Kurs_EUR, trn.Comment, trn.IsTransit,
                  trn.TotalSum_RUB_Out, trn.TotalSum_USD_Out, trn.TotalSum_EUR_Out,
                  trn.TotalSum_RUB_In, trn.TotalSum_USD_In, trn.TotalSum_EUR_In,
                  trn.Id, trn.IdSecurPack, trn.IdShift, trn.IdUser, trn.IdRoute, trn.TabNumber, trn.DtIncome, trn.CurDate, trn.DtRecord,
                  fil.ClientFullName, fil.FilialName, fil.FlgRecAlc, fil.BankName, fil.BIC_b AS Bik, fil.Account_b AS Korschet, fil.PostAddress, fil.FactAddress, fil.UIFilialName, fil.UIFactAddress, fil.MestoCode,
                  trn.IdKontrol, trn.KontrolDate, trn.IsActive, trn.IsDriver
                  FROM Tranzit_ATM_PER AS trn (READPAST)
                  INNER JOIN View_Client_Filial_Bank AS fil (READPAST) ON fil.FilialCode = trn.FilialCode
                  INNER JOIN Shift AS shf (READPAST) ON shf.Id = trn.IdShift
                  WHERE shf.ReportDate = ?ReportDateSmenaPRK AND fil.FactAddress LIKE '%<<f_FactAddress>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
                  ORDER BY trn.FilialCode, trn.WorkType, trn.BudgetNum

                ENDTEXT

*                  WHERE trn.IdShift IN (<<IdShift_PRK>>) AND fil.FactAddress LIKE '%<<f_FactAddress>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode


              CASE f_FilialAddress = 3 AND EMPTY(ALLTRIM(f_UIFactAddress)) = .F.

                TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                  SELECT trn.BudgetNum, trn.FilialCode, trn.WorkType, trn.SuperBudget, trn.Sost, trn.DateInkas, trn.Kurs_USD, trn.Kurs_EUR, trn.Comment, trn.IsTransit,
                  trn.TotalSum_RUB_Out, trn.TotalSum_USD_Out, trn.TotalSum_EUR_Out,
                  trn.TotalSum_RUB_In, trn.TotalSum_USD_In, trn.TotalSum_EUR_In,
                  trn.Id, trn.IdSecurPack, trn.IdShift, trn.IdUser, trn.IdRoute, trn.TabNumber, trn.DtIncome, trn.CurDate, trn.DtRecord,
                  fil.ClientFullName, fil.FilialName, fil.FlgRecAlc, fil.BankName, fil.BIC_b AS Bik, fil.Account_b AS Korschet, fil.PostAddress, fil.FactAddress, fil.UIFilialName, fil.UIFactAddress, fil.MestoCode,
                  trn.IdKontrol, trn.KontrolDate, trn.IsActive, trn.IsDriver
                  FROM Tranzit_ATM_PER AS trn (READPAST)
                  INNER JOIN View_Client_Filial_Bank AS fil (READPAST) ON fil.FilialCode = trn.FilialCode
                  INNER JOIN Shift AS shf (READPAST) ON shf.Id = trn.IdShift
                  WHERE shf.ReportDate = ?ReportDateSmenaPRK AND fil.UIFactAddress LIKE '%<<f_UIFactAddress>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
                  ORDER BY trn.FilialCode, trn.WorkType, trn.BudgetNum

                ENDTEXT

            ENDCASE

*                  WHERE trn.IdShift IN (<<IdShift_PRK>>) AND fil.UIFactAddress LIKE '%<<f_UIFactAddress>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode

          ELSE
            LOOP
          ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        CASE put_select = 9  && ������� ������ �� ����� �����

          f_TotalSum = 0
*          f_TotalSum = 100000.00

          DO FORM Vibor_summa_client

          IF NOT LASTKEY() = 27 AND pr_lastkey_27 = .T.

* WAIT WINDOW '����� ������ = ' + ALLTRIM(TRANSFORM(f_TotalSum, '999 999 999.99')) TIMEOUT 1

            TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

              SELECT trn.BudgetNum, trn.FilialCode, trn.WorkType, trn.SuperBudget, trn.Sost, trn.DateInkas, trn.Kurs_USD, trn.Kurs_EUR, trn.Comment, trn.IsTransit,
              trn.TotalSum_RUB_Out, trn.TotalSum_USD_Out, trn.TotalSum_EUR_Out,
              trn.TotalSum_RUB_In, trn.TotalSum_USD_In, trn.TotalSum_EUR_In,
              trn.Id, trn.IdSecurPack, trn.IdShift, trn.IdUser, trn.IdRoute, trn.TabNumber, trn.DtIncome, trn.CurDate, trn.DtRecord,
              fil.ClientFullName, fil.FilialName, fil.FlgRecAlc, fil.BankName, fil.BIC_b AS Bik, fil.Account_b AS Korschet, fil.PostAddress, fil.FactAddress, fil.UIFilialName, fil.UIFactAddress, fil.MestoCode,
              trn.IdKontrol, trn.KontrolDate, trn.IsActive, trn.IsDriver
              FROM Tranzit_ATM_PER AS trn (READPAST)
              INNER JOIN View_Client_Filial_Bank AS fil (READPAST) ON fil.FilialCode = trn.FilialCode
              INNER JOIN Shift AS shf (READPAST) ON shf.Id = trn.IdShift
              WHERE shf.ReportDate = ?ReportDateSmenaPRK AND trn.TotalSum_RUB_Out = ?f_TotalSum AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
              ORDER BY trn.FilialCode, trn.WorkType, trn.BudgetNum

            ENDTEXT

*              WHERE trn.IdShift IN (<<IdShift_PRK>>) AND trn.TotalSum_RUB_Out = ?f_TotalSum AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode

          ELSE
            LOOP
          ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        CASE put_select = 11  && ������� ������ �� �������� �����

          f_BankName = ''
*          f_BankName = '�����������'

          DO FORM Vibor_name_bank

          IF NOT LASTKEY() = 27 AND pr_lastkey_27 = .T.

            TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

              SELECT trn.BudgetNum, trn.FilialCode, trn.WorkType, trn.SuperBudget, trn.Sost, trn.DateInkas, trn.Kurs_USD, trn.Kurs_EUR, trn.Comment, trn.IsTransit,
              trn.TotalSum_RUB_Out, trn.TotalSum_USD_Out, trn.TotalSum_EUR_Out,
              trn.TotalSum_RUB_In, trn.TotalSum_USD_In, trn.TotalSum_EUR_In,
              trn.Id, trn.IdSecurPack, trn.IdShift, trn.IdUser, trn.IdRoute, trn.TabNumber, trn.DtIncome, trn.CurDate, trn.DtRecord,
              fil.ClientFullName, fil.FilialName, fil.FlgRecAlc, fil.BankName, fil.BIC_b AS Bik, fil.Account_b AS Korschet, fil.PostAddress, fil.FactAddress, fil.UIFilialName, fil.UIFactAddress, fil.MestoCode,
              trn.IdKontrol, trn.KontrolDate, trn.IsActive, trn.IsDriver
              FROM Tranzit_ATM_PER AS trn (READPAST)
              INNER JOIN View_Client_Filial_Bank AS fil (READPAST) ON fil.FilialCode = trn.FilialCode
              INNER JOIN Shift AS shf (READPAST) ON shf.Id = trn.IdShift
              WHERE shf.ReportDate = ?ReportDateSmenaPRK AND fil.BankName LIKE '%<<f_BankName>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
              ORDER BY trn.FilialCode, trn.WorkType, trn.BudgetNum

            ENDTEXT

*              WHERE trn.IdShift IN (<<IdShift_PRK>>) AND fil.BankName LIKE '%<<f_BankName>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode

          ELSE
            LOOP
          ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        CASE put_select = 13  && ������� ������ �� �������� ������

          f_GroupName = ''
*          f_GroupName = '�������'

          DO FORM Vibor_name_group

          IF NOT LASTKEY() = 27 AND pr_lastkey_27 = .T.

* WAIT WINDOW '����������� ������ = ' + ALLTRIM(f_GroupName) TIMEOUT 1

            TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

              SELECT trn.BudgetNum, trn.FilialCode, trn.WorkType, trn.SuperBudget, trn.Sost, trn.DateInkas, trn.Kurs_USD, trn.Kurs_EUR, trn.Comment, trn.IsTransit,
              trn.TotalSum_RUB_Out, trn.TotalSum_USD_Out, trn.TotalSum_EUR_Out,
              trn.TotalSum_RUB_In, trn.TotalSum_USD_In, trn.TotalSum_EUR_In,
              trn.Id, trn.IdSecurPack, trn.IdShift, trn.IdUser, trn.IdRoute, trn.TabNumber, trn.DtIncome, trn.CurDate, trn.DtRecord,
              fil.ClientFullName, fil.FilialName, fil.FlgRecAlc, fil.BankName, fil.BIC_b AS Bik, fil.Account_b AS Korschet, fil.PostAddress, fil.FactAddress, fil.UIFilialName, fil.UIFactAddress, fil.MestoCode,
              trn.IdKontrol, trn.KontrolDate, trn.IsActive, trn.IsDriver
              FROM Tranzit_ATM_PER AS trn (READPAST)
              INNER JOIN View_Client_Filial_Bank AS fil (READPAST) ON fil.FilialCode = trn.FilialCode
              INNER JOIN Shift AS shf (READPAST) ON shf.Id = trn.IdShift
              WHERE shf.ReportDate = ?ReportDateSmenaPRK AND fil.ClientGroupName LIKE '%<<f_GroupName>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
              ORDER BY trn.FilialCode, trn.WorkType, trn.BudgetNum

            ENDTEXT

*              WHERE trn.IdShift IN (<<IdShift_PRK>>) AND fil.ClientGroupName LIKE '%<<f_GroupName>>%' AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode

          ELSE
            LOOP
          ENDIF


* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        CASE put_select = 15  && ������� ������ � ������ ������ �� ����

          TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

            SELECT trn.BudgetNum, trn.FilialCode, trn.WorkType, trn.SuperBudget, trn.Sost, trn.DateInkas, trn.Kurs_USD, trn.Kurs_EUR, trn.Comment, trn.IsTransit,
            trn.TotalSum_RUB_Out, trn.TotalSum_USD_Out, trn.TotalSum_EUR_Out,
            trn.TotalSum_RUB_In, trn.TotalSum_USD_In, trn.TotalSum_EUR_In,
            trn.Id, trn.IdSecurPack, trn.IdShift, trn.IdUser, trn.IdRoute, trn.TabNumber, trn.DtIncome, trn.CurDate, trn.DtRecord,
            fil.ClientFullName, fil.FilialName, fil.FlgRecAlc, fil.BankName, fil.BIC_b AS Bik, fil.Account_b AS Korschet, fil.PostAddress, fil.FactAddress, fil.UIFilialName, fil.UIFactAddress, fil.MestoCode,
            trn.IdKontrol, trn.KontrolDate, trn.IsActive, trn.IsDriver
            FROM Tranzit_ATM_PER AS trn (READPAST)
            INNER JOIN View_Client_Filial_Bank AS fil (READPAST) ON fil.FilialCode = trn.FilialCode
            INNER JOIN Shift AS shf (READPAST) ON shf.Id = trn.IdShift
            WHERE shf.ReportDate = ?ReportDateSmenaPRK AND trn.IsActive = 1 AND fil.MestoCode = ?f_MestoCode
            ORDER BY trn.FilialCode, trn.WorkType, trn.BudgetNum

          ENDTEXT

*            WHERE trn.IdShift IN (<<IdShift_PRK>>)

      ENDCASE

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

        SqlReqRefresh = ALLTRIM(SqlReq)

        IF RunSql_Load('Table_Kontrol_ATM', @ar) <> 1
          DO exit_sql_server IN RunSQL_Load
          MESSAGEBOX('������ ������� ������ �� ������� FROM Table_Kontrol_ATM', 16, '���������� � ������������� �� ���: 243, 303, 170, 256')
          RETURN -1
        ENDIF

        IF ar[1,2] = 0
          IF USED('Table_Kontrol_ATM') = .T.
            USE IN Table_Kontrol_ATM
          ENDIF
          MESSAGEBOX('����� �� ���� ' + DTOC(f_CurDate) + CHR(13) + '��������� ��� �������� �� ����������', 0, '����������� ��������� ��������� ���� ������')
          LOOP
        ENDIF

* BROWSE WINDOW brows

        IF USED('Table_Kontrol_ATM') = .T.

          SELECT Table_Kontrol_ATM

          colvo_zap = RECCOUNT()

          GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          IF colvo_zap <> 0

            SELECT Table_Kontrol_ATM
            GO TOP
*          BROWSE WINDOW brows

* ------------------------------------------------------- *
            DO FORM KontrolPaket_ATM
* ------------------------------------------------------- *

            IF pr_vvod = .T. AND pr_vvod_sql = .T.

              IF USED('Table_Kontrol_ATM') = .T.

                SELECT Table_Kontrol_ATM
                GO TOP

                SCAN && FOR Table_Kontrol_ATM.Sost = 1  && AND Table_Kontrol_ATM.IsActive = 1

                  f_IdTable_Kontrol_ATM = Table_Kontrol_ATM.Id

                  f_FilialCode = ALLTRIM(Table_Kontrol_ATM.FilialCode)
                  f_BudgetNum = ALLTRIM(Table_Kontrol_ATM.BudgetNum)

                  f_Sost = Table_Kontrol_ATM.Sost

                  f_Comment = ALLTRIM(Table_Kontrol_ATM.Comment)

                  f_IsTransit = Table_Kontrol_ATM.IsTransit

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                  f_TotalSum_RUB_Out = Table_Kontrol_ATM.TotalSum_RUB_Out

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                  f_TotalSum_USD_Out = Table_Kontrol_ATM.TotalSum_USD_Out

                  DO start IN KursValueLoc WITH 'USD', f_CurDate

                  f_Nominal_usd = f_Nominal
                  f_Kurs_usd = f_KursValue
                  trn_Kurs_usd = f_KursValue

                  f_TotalSum_USD_RUB_Out = ROUND(f_TotalSum_USD_Out * f_Nominal_usd * f_Kurs_usd, 2)

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                  f_TotalSum_EUR_Out = Table_Kontrol_ATM.TotalSum_EUR_Out

                  DO start IN KursValueLoc WITH 'EUR', f_CurDate

                  f_Nominal_eur = f_Nominal
                  f_Kurs_eur = f_KursValue
                  trn_Kurs_eur = f_KursValue

                  f_TotalSum_EUR_RUB_Out = ROUND(f_TotalSum_EUR_Out * f_Nominal_eur * f_Kurs_eur, 2)

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                  f_TotalSum_RUB_In = Table_Kontrol_ATM.TotalSum_RUB_In

                  f_TotalSum_USD_In = Table_Kontrol_ATM.TotalSum_USD_In
                  f_TotalSum_USD_RUB_In = ROUND(f_TotalSum_USD_In * f_Nominal_usd * f_Kurs_usd, 2)

                  f_TotalSum_EUR_In = Table_Kontrol_ATM.TotalSum_EUR_In
                  f_TotalSum_Eur_RUB_In = ROUND(f_TotalSum_EUR_In * f_Nominal_eur * f_Kurs_eur, 2)

                  f_IdKontrol = f_IdUser

                  f_KontrolDate = Table_Kontrol_ATM.KontrolDate

                  f_IsActive = 1  && Table_Kontrol_ATM.IsActive

                  f_DtRecord = TTOC(Table_Kontrol_ATM.DtRecord)

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                  TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                    UPDATE Tranzit_ATM_PER SET
                    Sost = ?f_Sost,
                    IsTransit = ?f_IsTransit,
                    TotalSum_RUB_Out = ?f_TotalSum_RUB_Out,
                    TotalSum_USD_Out = ?f_TotalSum_USD_Out,
                    TotalSum_USD_RUB_Out = ?f_TotalSum_USD_RUB_Out,
                    TotalSum_EUR_Out = ?f_TotalSum_EUR_Out,
                    TotalSum_EUR_RUB_Out = ?f_TotalSum_EUR_RUB_Out,
                    TotalSum_RUB_In = ?f_TotalSum_RUB_In,
                    TotalSum_USD_In = ?f_TotalSum_USD_In,
                    TotalSum_USD_RUB_In = ?f_TotalSum_USD_RUB_In,
                    TotalSum_EUR_In = ?f_TotalSum_EUR_In,
                    TotalSum_EUR_RUB_In = ?f_TotalSum_EUR_RUB_In,
                    Comment = ?f_Comment,
                    IdKontrol = ?f_IdKontrol,
                    KontrolDate = ?f_KontrolDate,
                    IsActive = ?f_IsActive
                    WHERE Id = ?f_IdTable_Kontrol_ATM AND RTRIM(CONVERT(Char(20), DtRecord, 104)) + ' ' + RTRIM(CONVERT(Char(20), DtRecord, 108)) = ?f_DtRecord

                  ENDTEXT

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                  IF RunSql_Load('', @ar) <> 1
                    MESSAGEBOX('������ ������ ����������� ������ � ������� Table_Kontrol_ATM' + CHR(13) + ;
                      '��� - ' + ALLTRIM(f_FilialCode) + '   ����� ����� - ' + ALLTRIM(f_BudgetNum) + '   ����� = ' + ALLTRIM(TRANSFORM(f_TotalSum, '999 999 999 999.99')), 16, '���������� � ������������� �� ���: 243, 303, 170, 256')
                    LOOP
                  ENDIF

                  SELECT Table_Kontrol_ATM
                ENDSCAN

              ELSE
                =err('��������! ������� Table_Kontrol_ATM, ����������� ��� ������ �� ������� ... ')
                LOOP
              ENDIF

            ENDIF

          ELSE
            =err('��������! � �������������� ������� �� ���� SQL Server ������� �� ���������� ... ')
            LOOP
          ENDIF

        ELSE
          =err('��������! ������� ������ �� ���� SQL Server ����������� �������� ... ')
          LOOP
        ENDIF

      ENDIF

    ELSE
      EXIT
    ENDIF

  ENDDO

* WAIT WINDOW 'DO exit_sql_server IN RunSQL_Load'

ENDIF

* ---------------------------------------------------------------- *
DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

SELECT(sel)
RETURN


***********************************************************************************************************************************************


