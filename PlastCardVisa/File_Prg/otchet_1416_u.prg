**************************************************
*** ������ � ���������� �������� - ����� 1416 � ***
**************************************************

****************************************************************** PROCEDURE start *********************************************************************

PROCEDURE start
sel = SELECT()
popup_ima=LOWER(POPUP())
prompt_ima=LOWER(PROMPT())
bar_num=BAR()
HIDE POPUP (popup_ima)

put_vivod = 1

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO WHILE .T.

  tex1='������������ ������ �� ������� ���� ������������� ���'
  tex2='������������ ������ �� �������� ���� ������������� ���'
  l_bar=3
  =popup_9(tex1,tex2,tex3,tex4,l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor
  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

    put_vivod = BAR()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    DO CASE
      CASE put_vivod = 1  && ������������ ������ �� ������� ���� ������������� ���

        STORE new_data_odb TO dt1, dt2

        DO formir_forma_1416_u IN Otchet_1416_U

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE put_vivod = 3  && ������������ ������ �� �������� ���� ������������� ���

        STORE arx_data_odb TO dt1, dt2

        DO WHILE .T.

          SELECT DISTINCT data ;
            FROM data_odb ;
            INTO CURSOR tov ;
            ORDER BY 1 DESC

          IF _TALLY <> 0

            DO FORM (put_scr)+('seldata1.scx')

            IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

              dt1 = dt2

              DO formir_forma_1416_u IN Otchet_1416_U

            ELSE
              EXIT
            ENDIF

          ELSE
            =soob('��������! ������������� ���� ������ �� ���������� ...')
            EXIT
          ENDIF

        ENDDO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    ENDCASE

  ELSE
    EXIT
  ENDIF

ENDDO

SELECT(sel)
RETURN


****************************************************************** PROCEDURE start *********************************************************************

PROCEDURE formir_forma_1416_u
sel = SELECT()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

put_prn = 1

DO WHILE .T.

  tex1='���������� ������ � ����� ������������ �������� ������'
  tex2='���������� ������� ����� �������������� �������� ������'
  tex3='���������� �������� ����� �������������� ������ �� ��������� ����'
  tex4='���������� ������� ����� �������������� ������ � DOS ���������'
  l_bar=7
  =popup_9(tex1,tex2,tex3,tex4,l_bar)
  ACTIVATE POPUP vibor BAR put_prn
  RELEASE POPUPS vibor
  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

    put_prn = BAR()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    DO CASE
      CASE BAR() = 1  && ���������� ������ � ����� ������������ �������� ������

        DO scan_forma_1416_u IN Otchet_1416_U

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE BAR() = 3  && ���������� ������� ����� �������������� �������� ������

        DO select_forma_1416_u IN Otchet_1416_U

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE BAR() = 5  && ���������� �������� ����� �������������� ������ �� ��������� ����

        =soob('��������! ���������� ������ ��������� ������� ������������ ������ ������ � ��������� ... ')

        tex1 = '� � � � � � � � �  �������� ������ �� ��������� ����'
        tex2 = '� � � � � � � � � �  �� ���������� ������ ���������'
        l_bar = 3
        =popup_9(tex1,tex2,tex3,tex4,l_bar)
        ACTIVATE POPUP vibor
        RELEASE POPUPS vibor

        IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

          IF BAR() = 1

            IF USED('forma_1416_u')

              SELECT forma_1416_u

              COUNT FOR BETWEEN(date_ostat, dt1, dt2)

              IF _TALLY <> 0

                tim1 = SECONDS()

                ACTIVATE WINDOW poisk
                @ WROWS()/3,3 SAY PADC('��������! ������ �������� ������� � ������������ ��������.',WCOLS())

                SET EXCLUSIVE ON

                SELECT forma_1416_u
                USE
                USE (put_dbf)+('forma_1416_u.dbf') ALIAS forma_1416_u ORDER TAG date_ostat EXCLUSIVE

                SELECT forma_1416_u
                DELETE FOR BETWEEN(date_ostat, dt1, dt2)
                PACK
                REINDEX

                SELECT forma_1416_u
                USE
                USE (put_dbf)+('forma_1416_u.dbf') ALIAS forma_1416_u ORDER TAG date_ostat SHARE

                SET EXCLUSIVE OFF

                tim2=SECONDS()

                CLEAR
                @ WROWS()/3,3 SAY PADC('�������� ��������� ������ ������� ���������.'+' (�����='+ALLTRIM(STR(tim2 - tim1, 8, 4))+' ���.)',WCOLS())
                =INKEY(2)
                DEACTIVATE WINDOW poisk

              ELSE
                =soob('��������! ������������� ���� ������ �� ���������� ...')
              ENDIF

            ENDIF

          ENDIF
        ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE BAR() = 7  && ���������� ������� ����� �������������� ������ � DOS ���������

        tim1 = SECONDS()

        SELECT forma_1416_u

        SELECT * ;
          FROM forma_1416_u ;
          WHERE BETWEEN(date_ostat, dt1, dt2) ;
          INTO CURSOR sel_forma_1416_u ;
          ORDER BY date_ostat

        IF _TALLY <> 0

          ACTIVATE WINDOW poisk
          @ WROWS()/3,3 SAY PADC('��������! ����� ������� ������ � DOS ���������.',WCOLS())

          COPY TO (put_tmp)+('forma_1416_u.dbf') AS 866

          tim2=SECONDS()

          CLEAR
          @ WROWS()/3,3 SAY PADC('���� '+(put_tmp)+('forma_1416_u.dbf')+' ������� �������.'+' (�����='+ALLTRIM(STR(tim2 - tim1, 8, 4))+' ���.)',WCOLS())
          =INKEY(4)
          DEACTIVATE WINDOW poisk

        ELSE
          =soob('��������! ������������� ���� ������ �� ���������� ...')
        ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    ENDCASE

  ELSE
    EXIT
  ENDIF

ENDDO

SELECT(sel)
RETURN


******************************************************* PROCEDURE select_forma_1416_u *******************************************************************

PROCEDURE select_forma_1416_u

SELECT forma_1416_u

SELECT * ;
  FROM forma_1416_u ;
  WHERE BETWEEN(date_ostat, dt1, dt2) ;
  INTO CURSOR sel_count ;
  ORDER BY 1

IF _TALLY <> 0

  DO WHILE .T.

    tex1='��������� �������� ������'
    tex2='�������� ����� �� �������'
    l_bar=3
    =popup_9(tex1,tex2,tex3,tex4,l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor
    IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      DO CASE
        CASE BAR() = 1  && ��������� �������� ������

          SELECT branch, date_ostat,;
            bal_40803, bal_40806, bal_40809,;
            bal_40812, bal_40813, bal_40814, bal_40815, bal_40817,;
            bal_423, bal_426,;
            bal_47603, bal_47605,;
            bal_522, bal_52404, itog_summa ;
            FROM forma_1416_u ;
            WHERE BETWEEN(date_ostat, dt1, dt2) ;
            INTO CURSOR sel_ostatok ;
            ORDER BY branch, date_ostat

          IF _TALLY <> 0

            DO brw_ostat_vedom IN Otchet_1416_U

          ELSE
            =soob('��������! ������ � ������� �������� �� ���������� ... ')
          ENDIF

          LOOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        CASE BAR() = 3  && �������� ����� �� �������

          PUBLIC colvo_nls_sum, colvo_nls_100, sum_nls_100

          DO poisk_den_kurs IN Servis WITH dt1, 'USD'

          DO poisk_den_kurs IN Servis WITH dt1, 'EUR'

          SELECT * ;
            FROM account ;
            WHERE LEFT(ALLTRIM(card_nls), 5) <> '47411' ;
            INTO CURSOR  sel_account

          SELECT COUNT(*) AS colvo_nls ;
            FROM sel_account ;
            INTO CURSOR  colvo_account

          colvo_nls_sum = colvo_account.colvo_nls

          SELECT COUNT(*) AS colvo_nls ;
            FROM sel_account ;
            WHERE end_bal > 100000.00 AND ALLTRIM(val_card) == 'RUR' ;
            UNION ALL ;
            SELECT COUNT(*) AS colvo_nls ;
            FROM sel_account ;
            WHERE ROUND(end_bal * data_kurs_usd, 2) > 100000.00 AND ALLTRIM(val_card) == 'USD' ;
            UNION ALL ;
            SELECT COUNT(*) AS colvo_nls ;
            FROM sel_account ;
            WHERE ROUND(end_bal * data_kurs_eur, 2) > 100000.00 AND ALLTRIM(val_card) == 'EUR' ;
            INTO CURSOR  colvo_account_3

* BROWSE WINDOW brows

          SELECT SUM(colvo_nls) AS colvo_nls ;
            FROM colvo_account_3 ;
            INTO CURSOR  colvo_account

* BROWSE WINDOW brows

          colvo_nls_100 = colvo_account.colvo_nls

          SELECT ROUND(SUM(end_bal) / 1000, 2) AS end_bal ;
            FROM sel_account ;
            WHERE end_bal <= 100000.00 AND ALLTRIM(val_card) == 'RUR' ;
            UNION ALL ;
            SELECT ROUND((SUM(end_bal) * data_kurs_usd) / 1000, 2) AS end_bal ;
            FROM sel_account ;
            WHERE ROUND(end_bal * data_kurs_usd, 2) <= 100000.00 AND ALLTRIM(val_card) == 'USD' ;
            UNION ALL ;
            SELECT ROUND((SUM(end_bal) * data_kurs_eur) / 1000, 2) AS end_bal ;
            FROM sel_account ;
            WHERE ROUND(end_bal * data_kurs_eur, 2) <= 100000.00 AND ALLTRIM(val_card) == 'EUR' ;
            INTO CURSOR  colvo_account_3

* BROWSE WINDOW brows

          SELECT SUM(end_bal) AS sum_nls ;
            FROM colvo_account_3 ;
            INTO CURSOR  colvo_account

* BROWSE WINDOW brows

          sum_nls_100 = colvo_account.sum_nls

          SELECT branch, date_ostat,;
            bal_40803, bal_40806, bal_40809,;
            bal_40812, bal_40813, bal_40814, bal_40815, bal_40817,;
            bal_423, bal_426,;
            bal_47603, bal_47605,;
            bal_522, bal_52404, itog_summa ;
            FROM forma_1416_u ;
            WHERE BETWEEN(date_ostat, dt1, dt2) ;
            INTO CURSOR sel_ostatok ;
            ORDER BY branch, date_ostat

* BROWSE WINDOW brows

          DO WHILE .T.
            tex1=' �������� ����� �� ����� '
            tex2=' �������� ����� �� ������� '
            l_bar=3
            =popup_9(tex1,tex2,tex3,tex4,l_bar)
            ACTIVATE POPUP vibor
            RELEASE POPUPS vibor
            IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

              put_prn=BAR()

              SELECT DISTINCT branch ;
                FROM sel_ostatok ;
                INTO CURSOR  sel_branch ;
                ORDER BY 1

              SELECT sel_branch

              SCAN

                rec=ALLTRIM(sel_branch.branch)

                SELECT branch, date_ostat,;
                  bal_40803, bal_40806, bal_40809,;
                  bal_40812, bal_40813, bal_40814, bal_40815, bal_40817,;
                  bal_423, bal_426,;
                  bal_47603, bal_47605,;
                  bal_522, bal_52404, itog_summa ;
                  FROM sel_ostatok ;
                  WHERE ALLTRIM(branch) == rec ;
                  INTO CURSOR  prn_ostatok ;
                  ORDER BY branch, date_ostat

                IF _TALLY <> 0

                  dt1 = new_data_odb

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                  tex1='������� �������� ����� �� ������ ������� �4 �� 2 ������'
                  tex2='������� �������� ����� �� ������ ������� �3 �� 1 �����'
                  l_bar=3
                  =popup_9(tex1,tex2,tex3,tex4,l_bar)
                  ACTIVATE POPUP vibor
                  RELEASE POPUPS vibor
                  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

                    put_bar = BAR()

                    DO CASE
                      CASE put_bar = 1  && ������� �������� ����� �� ������ ������� �4 �� 2 ������

                        ima_frx_1 = (put_frx)+('forma_1416_u_1.frx')
                        ima_frx_2 = (put_frx)+('forma_1416_u_2.frx')

                        DO prn_form IN Otchet_1416_U

                      CASE put_bar = 3  && ������� �������� ����� �� ������ ������� �3 �� 1 �����

                        ima_frx = (put_frx)+('forma_1416_u.frx')

                        DO prn_form IN Otchet_1416_U

                    ENDCASE

                  ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                  RELEASE colvo_nls_sum, colvo_nls_100, sum_nls_100

                ELSE
                  =soob('��������! ������ � ������� �������� �� ���������� ... ')
                ENDIF

                SELECT sel_branch
              ENDSCAN

              EXIT

            ELSE
              EXIT
            ENDIF

          ENDDO

      ENDCASE

    ELSE
      EXIT
    ENDIF

  ENDDO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

ELSE
  =soob('��������! ������ � ������� �������� �� ���������� ... ')
ENDIF

RETURN


****************************************************** PROCEDURE brw_ostat_vedom ***********************************************************************

PROCEDURE brw_ostat_vedom

BROWSE FIELDS ;
  branch:H='������',;
  date_ostat:H='����',;
  bal_40803:H='���� � 40803':P='999 999 999 999.99',;
  bal_40806:H='���� � 40806':P='999 999 999 999.99',;
  bal_40809:H='���� � 40809':P='999 999 999 999.99',;
  bal_40812:H='���� � 40812':P='999 999 999 999.99',;
  bal_40813:H='���� � 40813':P='999 999 999 999.99',;
  bal_40814:H='���� � 40814':P='999 999 999 999.99',;
  bal_40815:H='���� � 40815':P='999 999 999 999.99',;
  bal_40817:H='���� � 40817':P='999 999 999 999.99',;
  bal_423:H='���� � 423':P='999 999 999 999.99',;
  bal_426:H='���� � 426':P='999 999 999 999.99',;
  bal_47603:H='���� � 47603':P='999 999 999 999.99',;
  bal_47605:H='���� � 47605':P='999 999 999 999.99',;
  bal_522:H='���� � 522':P='999 999 999 999.99',;
  bal_52404:H='���� � 52404':P='999 999 999 999.99',;
  itog_summa:H='����� �� ������':P='999 999 999 999.99' ;
  TITLE ' ������� ������� �� ���������� ������ ��������� �� ���������� ������ ' ;
  WINDOW brows

RETURN


************************************************************** PROCEDURE scan_forma_1416_u *************************************************************

PROCEDURE scan_forma_1416_u

STORE SPACE(0) TO num_balan
STORE 0 TO len_num_balan

SELECT forma_1416_u

FOR I = 3 TO 17  && �������� ������ � ����� ������������ ����� ������� forma_1416_u.dbf. �� �������� ���� ����������� ����� ����������� �����

  name_pole_scan = LOWER(ALLTRIM(FIELDS(I)))

*  WAIT WINDOW '����� ����������� ����� = ' + ALLTRIM(name_pole_scan) TIMEOUT 1

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  DO CASE
    CASE ALLTRIM(name_pole_scan) == 'bal_40803'

      num_balan = '40803'
      len_num_balan = 5

    CASE ALLTRIM(name_pole_scan) == 'bal_40806'

      num_balan = '40806'
      len_num_balan = 5

    CASE ALLTRIM(name_pole_scan) == 'bal_40809'

      num_balan = '40809'
      len_num_balan = 5

    CASE ALLTRIM(name_pole_scan) == 'bal_40812'

      num_balan = '40812'
      len_num_balan = 5

    CASE ALLTRIM(name_pole_scan) == 'bal_40813'

      num_balan = '40813'
      len_num_balan = 5

    CASE ALLTRIM(name_pole_scan) == 'bal_40814'

      num_balan = '40814'
      len_num_balan = 5

    CASE ALLTRIM(name_pole_scan) == 'bal_40815'

      num_balan = '40815'
      len_num_balan = 5

    CASE ALLTRIM(name_pole_scan) == 'bal_40817'

      num_balan = '40817'
      len_num_balan = 5

    CASE ALLTRIM(name_pole_scan) == 'bal_423'

      num_balan = '423'
      len_num_balan = 3

    CASE ALLTRIM(name_pole_scan) == 'bal_426'

      num_balan = '426'
      len_num_balan = 3

    CASE ALLTRIM(name_pole_scan) == 'bal_47603'

      num_balan = '47603'
      len_num_balan = 5

    CASE ALLTRIM(name_pole_scan) == 'bal_47605'

      num_balan = '47605'
      len_num_balan = 5

    CASE ALLTRIM(name_pole_scan) == 'bal_522'

      num_balan = '522'
      len_num_balan = 3

    CASE ALLTRIM(name_pole_scan) == 'bal_52404'

      num_balan = '52404'
      len_num_balan = 5

  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  DO CASE
    CASE put_vivod = 1  && ������������ ������ �� ������� ���� ������������� ���

      WAIT WINDOW ' ������� �� ����������� ����� '+ALLTRIM(num_balan)+' � �� "CARD-���� ������" ' TIMEOUT 1

      SELECT DISTINCT branch, dt1 AS date_vedom,;
        ROUND(SUM(begin_bal) / 1000, 2) AS begin_bal, ROUND(SUM(debit) / 1000, 2) AS debit, ROUND(SUM(credit) / 1000, 2) AS credit, ROUND(SUM(end_bal) / 1000, 2) AS end_bal ;
        WHERE LEFT(ALLTRIM(card_nls), len_num_balan) == ALLTRIM(num_balan) ;
        FROM account ;
        INTO CURSOR sum_ostat_bal ;
        ORDER BY 1

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      IF _TALLY = 0

        IF USED('rem_account_rub') = .F.
          USE visa!rem_account_rub IN 0
        ENDIF

        WAIT WINDOW ' ������� �� ����������� ����� '+ALLTRIM(num_balan)+' � ��� RsBank ' TIMEOUT 1

        SELECT DISTINCT num_branch AS branch, dt1 AS date_vedom,;
          account.balance,;
          account.account,;
          account.kind_account,;
          ROUND(SUM(Dbt_data.LDmoney2Currency(account.d0_l,account.d0_h,account.d0_p)) / 1000, 2) AS debit,;
          ROUND(SUM(Dbt_data.LDmoney2Currency(account.k0_l,account.k0_h,account.k0_p)) / 1000, 2) AS credit,;
          ROUND(SUM(Dbt_data.LDmoney2Currency(account.r0_l,account.r0_h,account.r0_p)) / 1000, 2) AS end_bal,;
          account.nameaccount ;
          FROM rem_account_rub account ;
          WHERE LEFT(ALLTRIM(account.account), len_num_balan) == ALLTRIM(num_balan) ;
          INTO CURSOR sum_ostat_bal ;
          ORDER BY 1

* BROWSE WINDOW brows

      ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE put_vivod = 3  && ������������ ������ �� �������� ���� ������������� ���

      SELECT DISTINCT branch, date_vedom,;
        ROUND(SUM(begin_bal) / 1000, 2) AS begin_bal, ROUND(SUM(debit) / 1000, 2) AS debit, ROUND(SUM(credit) / 1000, 2) AS credit, ROUND(SUM(end_bal) / 1000, 2) AS end_bal ;
        WHERE BETWEEN(date_vedom, dt1, dt2) AND LEFT(ALLTRIM(card_nls), len_num_balan) == ALLTRIM(num_balan) ;
        FROM vedom_ost ;
        INTO CURSOR sum_ostat_bal ;
        GROUP BY branch, date_vedom ;
        ORDER BY branch, date_vedom

  ENDCASE

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF _TALLY <> 0

    SELECT forma_1416_u
    SET ORDER TO date_ostat
    GO TOP

    SELECT sum_ostat_bal

    SCAN

      poisk_data = sum_ostat_bal.date_vedom

      WAIT WINDOW '������� ������ �� ����������� ����� = ' + ALLTRIM(num_balan) + '. ���� ��������� = ' + DTOC(poisk_data) TIMEOUT 0.5

      SELECT forma_1416_u

      poisk = SEEK(DTOS(poisk_data))

      IF poisk = .T.
        SCATTER MEMVAR
      ELSE
        SCATTER MEMVAR BLANK
      ENDIF

      m.branch = ALLTRIM(num_branch)

      m.date_ostat = poisk_data

      m.data_tran = DATETIME()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      DO CASE
        CASE ALLTRIM(name_pole_scan) == 'bal_40803'

          m.bal_40803 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_40806'

          m.bal_40806 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_40809'

          m.bal_40809 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_40812'

          m.bal_40812 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_40813'

          m.bal_40813 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_40814'

          m.bal_40814 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_40815'

          m.bal_40815 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_40817'

          m.bal_40817 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_423'

          m.bal_423 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_426'

          m.bal_426 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_47603'

          m.bal_47603 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_47605'

          m.bal_47605 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_522'

          m.bal_522 = sum_ostat_bal.end_bal

        CASE ALLTRIM(name_pole_scan) == 'bal_52404'

          m.bal_52404 = sum_ostat_bal.end_bal

      ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      m.itog_summa = m.bal_40803 + m.bal_40806 + m.bal_40809 + m.bal_40812 + m.bal_40813 + m.bal_40814 + m.bal_40815 + m.bal_40817 + ;
        m.bal_423 + m.bal_426 + m.bal_47603 + m.bal_47605 + m.bal_522 + m.bal_52404

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT forma_1416_u

      IF poisk = .T.
        GATHER MEMVAR
      ELSE
        INSERT INTO forma_1416_u FROM MEMVAR
      ENDIF

      SELECT sum_ostat_bal
    ENDSCAN

  ELSE
    WAIT WINDOW '��������! ������ �� ��������� �� ����������� ����� = ' + ALLTRIM(num_balan) + ' �� ���������� ... ' TIMEOUT 1
  ENDIF

  SELECT forma_1416_u

ENDFOR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF USED('rem_account_rub') = .T.
  SELECT rem_account_rub
  USE
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

RETURN


************************************************************ PROCEDURE prn_form ***********************************************************************

PROCEDURE prn_form

DO CASE
  CASE put_prn = 1

    DO CASE
      CASE put_bar = 1  && ������� �������� ����� �� ������ ������� �4 �� 2 ������

        REPORT FORM (ima_frx_1) NOEJECT PREVIEW
        REPORT FORM (ima_frx_2) NOEJECT PREVIEW

      CASE put_bar = 3  && ������� �������� ����� �� ������ ������� �3 �� 1 �����

        REPORT FORM (ima_frx) NOEJECT PREVIEW

    ENDCASE

  CASE put_prn = 3

    DO WHILE .T.
      IF PRINTSTATUS()
        log=.T.
        EXIT
      ELSE
        =err('��������! �p����p �� �����. �p���p��� ���p������.')
        tex1=' �p�������� ������ '
        tex2=' ��������� ������ '
        l_bar=3
        =popup_9(tex1,tex2,tex3,tex4,l_bar)
        ACTIVATE POPUP vibor
        RELEASE POPUPS vibor
        IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC
          DO CASE
            CASE BAR()=1
              LOOP
            CASE BAR()=3
              log=.F.
              EXIT
          ENDCASE
        ELSE
          log=.F.
          EXIT
        ENDIF
      ENDIF
    ENDDO

    IF log=.T.

      colvo=1

      DO FORM (put_scr)+('colcopy.scx')

      IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

        FOR I=1 TO colvo

          DO CASE
            CASE par_prompt = .T.  && ��� ������ ������ �� ������ ���������� ������ ������ ������������ ���������

              DO CASE
                CASE put_bar = 1  && ������� �������� ����� �� ������ ������� �4 �� 2 ������

                  REPORT FORM (ima_frx_1) NOEJECT NOCONSOLE TO PRINTER PROMPT
                  REPORT FORM (ima_frx_2) NOEJECT NOCONSOLE TO PRINTER PROMPT

                CASE put_bar = 3  && ������� �������� ����� �� ������ ������� �3 �� 1 �����

                  REPORT FORM (ima_frx) NOEJECT NOCONSOLE TO PRINTER PROMPT

              ENDCASE

            CASE par_prompt = .F.  && ��� ������ ������ �� ������ �� ���������� ������ ������ ������������ ���������

              DO CASE
                CASE put_bar = 1  && ������� �������� ����� �� ������ ������� �4 �� 2 ������

                  REPORT FORM (ima_frx_1) NOEJECT NOCONSOLE TO PRINTER
                  REPORT FORM (ima_frx_2) NOEJECT NOCONSOLE TO PRINTER

                CASE put_bar = 3  && ������� �������� ����� �� ������ ������� �3 �� 1 �����

                  REPORT FORM (ima_frx) NOEJECT NOCONSOLE TO PRINTER

              ENDCASE

          ENDCASE

        ENDFOR

      ENDIF

    ENDIF

ENDCASE

RETURN


********************************************************************************************************************************************************








