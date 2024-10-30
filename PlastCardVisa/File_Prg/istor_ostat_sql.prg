************************************************************************************
*** Формирование истории остатков по карточному счету клиента из ПО "Transmaster" ***
************************************************************************************

******************************************************************** PROCEDURE start *********************************************************************

PROCEDURE start
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

DO WHILE .T.

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
  DO FORM (put_scr) + ('poisk_client.scx')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет выборка данных ... ',WCOLS())

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    stroka_sql = "EXECUTE select_account_fio ?len_client, ?fam_client"

    select_data = SQLEXEC(sql_num, stroka_sql, 'sel_account')

    DO WHILE SQLMORERESULTS(sql_num) < 2
    ENDDO

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SELECT COUNT(*) AS count_zap ;
      FROM sel_account ;
      INTO CURSOR count_colvo_zap ;
      ORDER BY 1

    colvo_zap = count_colvo_zap.count_zap

    =INKEY(1)
    DEACTIVATE WINDOW poisk

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    IF colvo_zap <> 0

      SELECT DISTINCT SPACE(1) + PADR(ALLTRIM(fio_rus), 45) + ' - ' + card_nls AS fio_rus ;
        FROM sel_account ;
        INTO CURSOR tov ;
        ORDER BY 1

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      DO FORM (put_scr) + ('selfio_nls.scx')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

        DO poisk_client IN Istor_ostat_sql

      ELSE
        EXIT
      ENDIF

    ELSE
      =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
    ENDIF

  ELSE
    EXIT
  ENDIF

ENDDO

SELECT(sel)
RETURN


**************************************************************** PROCEDURE poisk_client ******************************************************************

PROCEDURE poisk_client

DO WHILE .T.

  STORE new_data_odb TO dt1, dt2

  tex1 = 'Выборка данных по реальным датам операций'
  tex2 = 'Выборка данных по датам введенными оператором'
  l_bar = 3
  =popup_9(tex1,tex2,tex3,tex4,l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    put_vivod = BAR()

    DO CASE
      CASE put_vivod = 1  && Выборка данных по реальным датам операций

        ACTIVATE WINDOW poisk
        @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет выборка данных ... ',WCOLS())

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        select_data = SQLEXEC(sql_num,"EXECUTE select_istor_ost_cln ?sel_card_acct","istor_ost")

        DO WHILE SQLMORERESULTS(sql_num) < 2
        ENDDO

* WAIT WINDOW 'sel_card_acct = ' + ALLTRIM(sel_card_acct) + '  Результат выбора данных select_data_data = ' + ALLTRIM(STR(select_data)) TIMEOUT 2

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        SELECT DISTINCT TTOD(date_ost_p) AS data ;
          FROM istor_ost ;
          INTO CURSOR tov ;
          ORDER BY 1 DESC

        =INKEY(1)
        DEACTIVATE WINDOW poisk

        IF _TALLY <> 0

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
          DO FORM (put_scr) + ('seldata2.scx')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

            DO vibor_data IN Istor_ostat_sql

          ENDIF

        ELSE
          =soob('Внимание! Запрашиваемых Вами данных по счету - ' + ALLTRIM(sel_card_acct) + ' не обнаружено ...')
        ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE put_vivod = 3  && Выборка данных по датам введенными оператором

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
        DO FORM (put_scr) + ('sel_data_2.scx')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

          DO vibor_data IN Istor_ostat_sql

        ENDIF

    ENDCASE

  ELSE
    EXIT
  ENDIF

ENDDO

RETURN


******************************************************************* PROCEDURE vibor_data ****************************************************************

PROCEDURE vibor_data

DO WHILE .T.

  tex1 = 'Табличный просмотр данных'
  tex2 = 'Печатный отчет на принтер'
  l_bar = 3
  =popup_9(tex1,tex2,tex3,tex4,l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    DO CASE
      CASE BAR() = 1  && Табличный просмотр данных

        ACTIVATE WINDOW poisk
        @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет выборка данных ... ',WCOLS())

        SELECT branch, date_ost_p AS date_ostat, ref_client, card_nls, card_acct, val_card, fio_rus, name_bank_P AS name_p,;
          begin_bal, debit, credit, end_bal ;
          FROM istor_ost ;
          WHERE ALLTRIM(card_acct) == ALLTRIM(sel_card_acct) AND BETWEEN(DTOS(TTOD(date_ost_p)), DTOS(dt1), DTOS(dt2)) ;
          INTO CURSOR sel_ostatok ;
          ORDER BY branch, val_card, card_nls, date_ost_p

        =INKEY(1)
        DEACTIVATE WINDOW poisk

        SELECT sel_ostatok
        GO TOP

        IF _TALLY <> 0

          DO brw_ostat IN Istor_ostat_sql

        ELSE
          =soob('Внимание! Данных в таблице остатков не обнаружено ... ')
        ENDIF

        LOOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE BAR() = 3  && Печатный отчет на принтер

        ACTIVATE WINDOW poisk
        @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет выборка данных ... ',WCOLS())

        SELECT branch, ref_client, card_nls, card_acct, val_card, fio_rus, name_bank_P AS name_p,;
          date_ost_p AS date_ostat, begin_bal, debit, credit, end_bal,;
          SUBSTR(ALLTRIM(card_nls), 1, 8) + SUBSTR(ALLTRIM(card_nls), 10, 11) AS sort_nls ;
          FROM istor_ost ;
          WHERE ALLTRIM(card_acct) == ALLTRIM(sel_card_acct) AND BETWEEN(DTOS(TTOD(date_ost_p)), DTOS(dt1), DTOS(dt2)) ;
          INTO CURSOR sel_ostatok ;
          ORDER BY branch, val_card, sort_nls, date_ost_p

        =INKEY(1)
        DEACTIVATE WINDOW poisk

        DO WHILE .T.

          tex1 = 'Печатный отчет на экран'
          tex2 = 'Печатный отчет на принтер'
          l_bar = 3
          =popup_9(tex1,tex2,tex3,tex4,l_bar)
          ACTIVATE POPUP vibor
          RELEASE POPUPS vibor

          IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

            put_prn = BAR()

            SELECT DISTINCT branch ;
              FROM sel_ostatok ;
              INTO CURSOR sel_branch ;
              ORDER BY 1

            SELECT sel_branch

            SCAN

              rec = ALLTRIM(sel_branch.branch)

              SELECT branch, ref_client, card_nls, card_acct, val_card, fio_rus,;
                date_ostat, begin_bal, debit, credit, end_bal,;
                name_p, sort_nls ;
                FROM sel_ostatok ;
                WHERE ALLTRIM(branch) == ALLTRIM(rec) ;
                INTO CURSOR prn_ostatok ;
                ORDER BY val_card, date_ostat, sort_nls

              IF _TALLY <> 0

                DO CASE
                  CASE LEN(ALLTRIM(prn_ostatok.card_acct)) < 20

                    ima_frx = (put_frx) + ('istr_ostatok_1.frx')

                  CASE LEN(ALLTRIM(prn_ostatok.card_acct)) = 20

                    ima_frx = (put_frx) + ('istr_ostatok_2.frx')

                ENDCASE

                SELECT prn_ostatok
                GO TOP

                DO prn_form_branch IN Visa

              ELSE
                =soob('Внимание! Данных в таблице остатков не обнаружено ... ')
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

RETURN


******************************************************************* PROCEDURE brw_ostat *****************************************************************

PROCEDURE brw_ostat

BROWSE FIELDS ;
  branch:H = 'Филиал',;
  date_ostat:H = 'Дата остатка',;
  card_nls:H = 'Карточный счет',;
  card_acct:H = 'Счет вклада',;
  val_card:H = 'Валюта карты',;
  fio_rus:H = 'Фамилия Имя Отчество',;
  begin_bal:H = 'Вход. остаток.':P = '999 999 999.99',;
  debit:H = 'Дебет':P = '999 999 999.99',;
  credit:H = 'Кредит':P = '999 999 999.99',;
  end_bal:H = 'Исх. остаток':P = '999 999 999.99',;
  name_p:H = 'Название филиала' ;
  TITLE ' История остатков ' ;
  WINDOW brows

RETURN


*********************************************************************************************************************************************************



