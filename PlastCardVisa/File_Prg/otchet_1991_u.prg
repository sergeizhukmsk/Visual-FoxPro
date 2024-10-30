**************************************************
*** Данные о ежедневных остатках - форма 1991 У ***
**************************************************

****************************************************************** PROCEDURE start *********************************************************************

PROCEDURE start
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

DO formirov_menu IN Otchet_1991_U

ACTIVATE POPUP otchet_1991_U

RELEASE MENUS otchet_1991_U

SELECT(sel)
RETURN


*************************************************************** PROCEDURE select_account ****************************************************************

PROCEDURE select_account
PARAMETERS filtr_data

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, начата выборка данных ... ',WCOLS())

SELECT acc_del
COUNT TO colvo_acc_del

SELECT account
COUNT TO colvo_account

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF colvo_account <> 0

  SELECT DISTINCT branch, ref_client, fio_rus, num_dog, date_nls, stavka, card_acct AS num_nls, tran_42301 AS tran_nls,;
    begin_bal AS ostat_val, val_card AS val_nls, name_card AS name_nls ;
    FROM account ;
    WHERE del_card <> 2 AND ALLTRIM(vid_card) == '1' AND date_nls < filtr_data ;
    INTO CURSOR sel_account_O ;
    ORDER BY 3

  colvo_account = _TALLY

* BROWSE WINDOW brows TITLE 'Количество открытых счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

ELSE
  RETURN
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF colvo_acc_del <> 0

  SELECT DISTINCT branch, ref_client, fio_rus, num_dog, date_nls, stavka, card_acct AS num_nls, tran_42301 AS tran_nls,;
    begin_bal AS ostat_val, val_card AS val_nls, name_card AS name_nls ;
    FROM acc_del ;
    WHERE del_card = 2 AND ALLTRIM(vid_card) == '1' AND  date_nls < filtr_data ;
    INTO CURSOR sel_account_Z ;
    ORDER BY 3

  colvo_acc_del = _TALLY

* BROWSE WINDOW brows TITLE 'Количество закрытых счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO CASE
  CASE colvo_account <> 0 AND colvo_acc_del = 0

    SELECT DISTINCT branch, ref_client, fio_rus, num_dog, date_nls, stavka, num_nls, tran_nls, ostat_val, val_nls, name_nls ;
      FROM sel_account_O ;
      INTO CURSOR sel_account_pol ;
      ORDER BY 3

    colvo_zap = _TALLY

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

  CASE colvo_account <> 0 AND colvo_acc_del <> 0

    SELECT DISTINCT branch, ref_client, fio_rus, num_dog, date_nls, stavka, num_nls, tran_nls, ostat_val, val_nls, name_nls ;
      FROM sel_account_O ;
      UNION ;
      SELECT DISTINCT branch, ref_client, fio_rus, num_dog, date_nls, stavka, num_nls, tran_nls, ostat_val, val_nls, name_nls ;
      FROM sel_account_Z ;
      INTO CURSOR sel_account_pol ;
      ORDER BY 3

    colvo_zap = _TALLY

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF colvo_zap <> 0

  SELECT DISTINCT branch, ref_client, fio_rus, num_dog, date_nls, stavka, num_nls, tran_nls, ostat_val, val_nls, name_nls ;
    FROM sel_account_pol ;
    INTO CURSOR sel_account ;
    ORDER BY 3

  colvo_zap = _TALLY

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

  =INKEY(2)
  DEACTIVATE WINDOW poisk

ELSE

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')
ENDIF

RETURN


************************************************************ PROCEDURE forma_1991_U_den ****************************************************************

PROCEDURE forma_1991_U_den

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, начата выборка данных ... ',WCOLS())

DO poisk_den_kurs_avt IN Servis WITH new_data_odb, 'USD'
DO poisk_den_kurs_avt IN Servis WITH new_data_odb, 'EUR'

SELECT DISTINCT branch, new_data_odb AS date_vedom, ref_client, fio_rus, num_nls, tran_nls, num_dog, date_nls, stavka, ostat_val,;
  IIF(ALLTRIM(val_nls) == 'RUR', ROUND(1.00 * 1.00, 4),;
  IIF(ALLTRIM(val_nls) == 'USD', ROUND(poisk_kurs_usd * 1.00, 4),;
  IIF(ALLTRIM(val_nls) == 'EUR', ROUND(poisk_kurs_eur * 1.00, 4), 000.0000))) AS kurs,;
  IIF(ALLTRIM(val_nls) == 'RUR', ostat_val,;
  IIF(ALLTRIM(val_nls) == 'USD', ROUND(ostat_val * poisk_kurs_usd, 2 ),;
  IIF(ALLTRIM(val_nls) == 'EUR', ROUND(ostat_val * poisk_kurs_eur, 2 ), 000000.00))) AS ostat_rur ;
  FROM sel_account ;
  INTO CURSOR sel_reestr ;
  ORDER BY 1, 2, 4

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

=INKEY(2)
DEACTIVATE WINDOW poisk

RETURN


********************************************************* PROCEDURE forma_1991_U_arx *******************************************************************

PROCEDURE forma_1991_U_arx

SELECT DISTINCT data_home AS data ;
  FROM calendar ;
  WHERE data_home >= CTOD('01.06.2007') AND DAY(data_home) = 1 AND DTOS(data_home) < DTOS(new_data_odb) AND YEAR(data_home) <= YEAR(new_data_odb) ;
  INTO CURSOR tov ;
  ORDER BY 1 DESC

* BROWSE WINDOW brows

IF _TALLY <> 0

  =soob('Необходимо выбрать самую первую дату С КОТОРОЙ НЕОБХОДИМО СФОРМИРОВАТЬ ФОРМУ 1991-У')

  DO FORM (put_scr) + ('seldata1.scx')

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    DO poisk_den_kurs_avt IN Servis WITH new_data_odb, 'USD'
    DO poisk_den_kurs_avt IN Servis WITH new_data_odb, 'EUR'

* ------------------------------------------------------------------------------------ Выбор данных из таблицы ведомость остатков -------------------------------------------------------------------------------------- *

    text1 = 'Вы желаете использовать ведомость остатков из ПО "TRANSMASTER"'
    text2 = 'Вы желаете использовать ведомость остатков из ПО "CARD-VISA филиал"'
    l_bar = 3
    =popup_9(text1, text2, text3, text4, l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      ACTIVATE WINDOW poisk
      @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, начата выборка данных ... ',WCOLS())

      DO CASE
        CASE BAR() = 1  && Вы желаете использовать ведомость остатков из ПО "TRANSMASTER"

          DO start IN Select_vedom_ost_tran

        CASE BAR() = 3  && Вы желаете использовать ведомость остатков из ПО "CARD-VISA филиал"

          DO start IN Select_vedom_ost_mak

      ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      IF _TALLY <> 0

* WAIT WINDOW 'Выбранная дата - ' + DTOC(dt1)

        SELECT DISTINCT branch, date_ostat, date_vedom, ref_client, fio_rus, val_card, card_acct AS num_nls, tran_42301 AS tran_nls,;
          SPACE(10) AS num_dog, CTOD('  .  .    ') AS date_nls, 0.00 AS stavka, ostat_val,;
          IIF(ALLTRIM(val_card) == 'RUR', ROUND(1.00 * 1.00, 4),;
          IIF(ALLTRIM(val_card) == 'USD', ROUND(poisk_kurs_usd * 1.00, 4),;
          IIF(ALLTRIM(val_card) == 'EUR', ROUND(poisk_kurs_eur * 1.00, 4), 000.0000))) AS kurs,;
          IIF(ALLTRIM(val_card) == 'RUR', ostat_val,;
          IIF(ALLTRIM(val_card) == 'USD', ROUND(ostat_val * poisk_kurs_usd, 2 ),;
          IIF(ALLTRIM(val_card) == 'EUR', ROUND(ostat_val * poisk_kurs_eur, 2 ), 000000.00))) AS ostat_rur ;
          FROM sum_vedom_ost ;
          WHERE date_vedom >= dt1 ;
          INTO CURSOR sel_reestr ;
          ORDER BY 1, 3, 5

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

      ELSE

        =INKEY(2)
        DEACTIVATE WINDOW poisk

        =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')

      ENDIF

      =INKEY(2)
      DEACTIVATE WINDOW poisk

    ENDIF

    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ENDIF

ELSE

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')

ENDIF

RETURN


************************************************************** PROCEDURE start_reestr *******************************************************************

PROCEDURE start_reestr
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

=soob('Формирование детализированной ведомости остатков по лицевым счетам для формы 1991 У')

DEFINE WINDOW otchet_zb AT 0,0 SIZE colvo_rows - 2, colvo_cols - 4 FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW otchet_zb CENTER

put_select = 1

DO WHILE .T.

  tex1 = 'Формирование отчета за текущую дату операционного дня'
  tex2 = 'Формирование отчета за архивную дату операционного дня'
  l_bar = 3
  =popup_9(tex1,tex2,tex3,tex4,l_bar)
  ACTIVATE POPUP vibor BAR put_select
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    put_select = IIF(BAR() = 0, 1, BAR())

    DO CASE
      CASE put_select = 1  && Формирование отчета за текущую дату операционного дня

        STORE new_data_odb TO dt1, dt2

        DO select_account IN Otchet_1991_U WITH dt1

        IF colvo_zap <> 0

          DO forma_1991_U_den IN Otchet_1991_U

        ELSE

          =INKEY(2)
          DEACTIVATE WINDOW poisk

          =err('Внимание! Данных по лицевым счетам не обнаружено ...')
        ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE put_select = 3  && Формирование отчета за архивную дату операционного дня

        STORE arx_data_odb TO dt1, dt2

        DO forma_1991_U_arx IN Otchet_1991_U

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    IF _TALLY <> 0

      =INKEY(2)
      DEACTIVATE WINDOW poisk

      put_vivod = 1

      DO WHILE .T.

        tex1 = 'Произвести заполнение данными отчетной таблицы за выбранную дату'
        tex2 = 'Печатный отчет по ранее сформированным данным за выбранную дату'
        tex3 = 'Произвести удаление ранее сформированных данных за выбранную дату'
        tex4 = 'Произвести экспорт ранее сформированных данных в DOS кодировке'
        l_bar = 7
        =popup_9(tex1,tex2,tex3,tex4,l_bar)
        ACTIVATE POPUP vibor BAR put_vivod
        RELEASE POPUPS vibor

        IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

          put_vivod = IIF(BAR() = 0, 1, BAR())

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          DO CASE
            CASE put_vivod = 1  && Произвести заполнение данными отчетной таблицы за выбранную дату

              SELECT sel_reestr
              GO TOP

              BROWSE FIELDS ;
                branch:H = 'Номер филиала',;
                date_vedom:H = 'Дата',;
                ref_client:H = 'Номер клиента',;
                fio_rus:H = 'Фамилия Имя Отчество',;
                num_dog:H = 'Номер договора',;
                date_nls:H = 'Дата договора',;
                stavka:H = 'Ставка',;
                num_nls:H = 'Лицевой счет',;
                tran_nls:H = 'Транзитный счет',;
                ostat_val:H = 'Вход. остаток в валюте вклада':P = '999 999 999.99',;
                kurs:H = 'Курс валюты':P = '999.9999',;
                ostat_rur:H = 'Вход. остаток по курсу ЦБ':P = '999 999 999.99' ;
                TITLE ' Детализированная ведомость остатков по лицевым счетам -  отчет 1991 У' ;
                WINDOW otchet_zb

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

              SELECT DISTINCT SUBSTR(ALLTRIM(num_nls), 6, 3) AS kod_val, SUM(ostat_val) AS summa ;
                FROM sel_reestr ;
                INTO CURSOR sel_1991 ;
                GROUP BY 1 ;
                ORDER BY 1

              BROWSE FIELDS ;
                kod_val:H = 'Код валюты',;
                summa:H = 'С у м м а':P = '999 999 999 999.99' ;
                TITLE ' Суммарные остатки на счетах по типам валют до формирования отчета 1991-У на дату ' + DT(dt1)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
              DO zapis_forma_1991 IN Otchet_1991_U
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

              SELECT forma_1991_u
              GO TOP

              SELECT branch, date_vedom, ref_client, fio_rus, num_dog, date_nls, stavka, num_nls, tran_nls, ostat_val, kurs, ostat_rur ;
                FROM forma_1991_u ;
                WHERE EMPTY(ALLTRIM(num_dog)) = .T. AND date_vedom = dt1 ;
                INTO CURSOR sel_forma_1991_u ;
                ORDER BY 1

* BROWSE WINDOW brows

              IF _TALLY <> 0
                =err('Внимание! В форме 1991-У обнаружены записи с не заполненным полем "Номер договора"')
                =soob('Выполните пункт меню "Заполнение регистрационных данных счетов в УЖЕ сформированной форме 1991 У"')
              ENDIF

              SELECT branch, date_vedom, ref_client, fio_rus, num_dog, date_nls, stavka, num_nls, tran_nls, ostat_val, kurs, ostat_rur ;
                FROM forma_1991_u ;
                WHERE EMPTY(date_nls) = .T. AND date_vedom = dt1 ;
                INTO CURSOR sel_forma_1991_u ;
                ORDER BY 1

* BROWSE WINDOW brows

              IF _TALLY <> 0
                =err('Внимание! В форме 1991-У обнаружены записи с не заполненным полем "Дата договора"')
                =soob('Выполните пункт меню "Заполнение регистрационных данных счетов в УЖЕ сформированной форме 1991 У"')
              ENDIF

              SELECT branch, date_vedom, ref_client, fio_rus, num_dog, date_nls, stavka, num_nls, tran_nls, ostat_val, kurs, ostat_rur ;
                FROM forma_1991_u ;
                WHERE EMPTY(ALLTRIM(tran_nls)) = .T. AND date_vedom = dt1 ;
                INTO CURSOR sel_forma_1991_u ;
                ORDER BY 1

* BROWSE WINDOW brows

              IF _TALLY <> 0
                =err('Внимание! В форме 1991-У обнаружены записи с не заполненным полем "Транзитный счет"')
                =soob('Выполните пункт меню "Заполнение регистрационных данных счетов в УЖЕ сформированной форме 1991 У"')
              ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            CASE put_vivod = 3  && Печатный отчет по ранее сформированным данным за выбранную дату

              SELECT forma_1991_u

              SELECT branch, date_vedom, ref_client, fio_rus, num_dog, date_nls, stavka, num_nls, tran_nls, ostat_val, kurs, ostat_rur ;
                FROM forma_1991_u ;
                WHERE BETWEEN(DTOS(date_vedom), DTOS(dt1), DTOS(dt2)) ;
                INTO CURSOR sel_forma_1991_u ;
                ORDER BY date_vedom, fio_rus

              IF _TALLY <> 0

                GO TOP

                ima_frx = (put_frx) + ('forma_1991_u.frx')

                DO prn_form IN Visa

              ELSE
                =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')
              ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            CASE put_vivod = 5  && Произвести удаление ранее сформированных данных за выбранную дату

              =soob('Выполнение данной процедуры требует монопольного режима работы с таблицами ... ')

              tex1 = 'В ы п о л н и т ь  удаление данных за выбранную дату'
              tex2 = 'О т к а з а т ь с я  от выполнения данной процедуры'
              l_bar = 3
              =popup_9(tex1,tex2,tex3,tex4,l_bar)
              ACTIVATE POPUP vibor
              RELEASE POPUPS vibor

              IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

                IF BAR() = 1

                  IF USED('forma_1991_u')

                    SELECT forma_1991_u

                    COUNT FOR BETWEEN(date_vedom, dt1, dt2)

                    IF _TALLY <> 0

                      tim1 = SECONDS()

                      ACTIVATE WINDOW poisk
                      @ WROWS()/3,3 SAY PADC('Внимание! Начата упаковка таблицы и перестроение индексов.',WCOLS())

                      SET EXCLUSIVE ON

                      SELECT forma_1991_u
                      USE
                      USE (put_dbf) + ('forma_1991_u.dbf') ALIAS forma_1991_u ORDER TAG date_fio EXCLUSIVE

                      SELECT forma_1991_u
                      DELETE FOR BETWEEN(date_vedom, dt1, dt2)
                      PACK
                      REINDEX

                      SELECT forma_1991_u
                      USE
                      USE (put_dbf) + ('forma_1991_u.dbf') ALIAS forma_1991_u ORDER TAG date_fio SHARE

                      SET EXCLUSIVE OFF

                      tim2 = SECONDS()

                      CLEAR
                      @ WROWS()/3,3 SAY PADC('Удаление выбранных данных успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
                      =INKEY(2)
                      DEACTIVATE WINDOW poisk

                    ELSE
                      =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')
                    ENDIF

                  ENDIF

                ENDIF
              ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            CASE put_vivod = 7  && Произвести экспорт ранее сформированных данных в DOS кодировке

              ACTIVATE WINDOW poisk
              @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт данных в DOS кодировке.',WCOLS())

              SELECT forma_1991_u

              SELECT branch, date_vedom, ref_client, fio_rus, num_dog, date_nls, stavka, srok_vklad, num_nls, tran_nls, ostat_val, kurs, ostat_rur ;
                FROM forma_1991_u ;
                WHERE BETWEEN(DTOS(date_vedom), DTOS(dt1), DTOS(dt2)) ;
                INTO CURSOR sel_forma_1991_u ;
                ORDER BY date_vedom, fio_rus

              IF _TALLY <> 0

                COPY TO (put_tmp) + ('forma_1991_u.dbf') FOX2X AS 866
                COPY TO (put_tmp) + ('forma_1991_u.xls') XL5 AS 1251

                @ WROWS()/3,3 SAY PADC('Файлы - ' + ('forma_1991_u.dbf') + ' и ' + ('forma_1991_u.xls') + ' успешно записаны по пути ' + (put_tmp),WCOLS())
                =INKEY(4)
                DEACTIVATE WINDOW poisk

              ELSE

                =INKEY(2)
                DEACTIVATE WINDOW poisk

                =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')

              ENDIF

          ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        ELSE
          EXIT
        ENDIF

      ENDDO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    ELSE

      =INKEY(2)
      DEACTIVATE WINDOW poisk

      =err('Внимание! Данных по реестру обязательств банка перед вкладчиками не обнаружено ...')
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  ELSE
    EXIT
  ENDIF

ENDDO

RELEASE WINDOW otchet_zb

RETURN


******************************************************** PROCEDURE zapis_registr_nls *********************************************************************

PROCEDURE zapis_registr_nls
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

SELECT DISTINCT fio_rus, num_nls, num_dog, date_nls ;
  FROM forma_1991_u ;
  WHERE EMPTY(ALLTRIM(num_dog)) = .T. ;
  UNION ;
  SELECT DISTINCT fio_rus, num_nls, num_dog, date_nls ;
  FROM forma_1991_u ;
  WHERE EMPTY(date_nls) = .T. ;
  INTO CURSOR sel_num_nls ;
  ORDER BY fio_rus

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

IF _TALLY <> 0

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Начато заполнение формы 1991-У регистрационными данными счетов ... ',WCOLS())
  =INKEY(2)

  SELECT sel_num_nls
  GO TOP

  SCAN

    sel_card_acct = ALLTRIM(sel_num_nls.num_nls)
    sel_fio_rus = ALLTRIM(sel_num_nls.fio_rus)

    WAIT WINDOW 'Клиент - ' + ALLTRIM(sel_num_nls.fio_rus) + '  Лицевой счет - ' + ALLTRIM(sel_num_nls.num_nls) NOWAIT

    SELECT DISTINCT ref_client, fio_rus, card_acct, num_dog, date_nls, stavka ;
      FROM account ;
      WHERE ALLTRIM(card_acct) == ALLTRIM(sel_card_acct) ;
      UNION ;
      SELECT DISTINCT ref_client, fio_rus, card_acct, num_dog, date_nls, stavka ;
      FROM acc_del ;
      WHERE ALLTRIM(card_acct) == ALLTRIM(sel_card_acct) ;
      INTO CURSOR sel_account ;
      ORDER BY 2, 3

    IF _TALLY = 0  && Искомый счет в справочнике счетов не найден

      SELECT DISTINCT fio_rus, card_nls, RIGHT(ALLTRIM(ref_client), 6) AS num_dog, data_sost AS date_nls, cln_stavka AS stavka ;
        FROM zajava_b ;
        WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_acct) ;
        INTO CURSOR sel_zajava_b ;
        ORDER BY fio_rus, card_acct

      IF _TALLY = 0  && Искомый счет в справочнике составленных заявлений не найден

        SELECT DISTINCT fio_rus, card_nls, RIGHT(ALLTRIM(ref_client), 6) AS num_dog, data_sost AS date_nls, cln_stavka AS stavka ;
          FROM zajava_b ;
          WHERE ALLTRIM(fio_rus) == ALLTRIM(sel_fio_rus) ;
          INTO CURSOR sel_zajava_b ;
          ORDER BY fio_rus, card_acct

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

        IF _TALLY <> 0

          SELECT sel_zajava_b
          GO TOP

          SELECT forma_1991_u

          UPDATE forma_1991_u SET ;
            num_dog = sel_zajava_b.num_dog,;
            date_nls = sel_zajava_b.date_nls,;
            stavka = sel_zajava_b.stavka ;
            WHERE ALLTRIM(num_nls) == ALLTRIM(sel_card_acct) ;

        ENDIF

      ELSE

        SELECT sel_zajava_b
        GO TOP

        SELECT forma_1991_u

        UPDATE forma_1991_u SET ;
          num_dog = sel_zajava_b.num_dog,;
          date_nls = sel_zajava_b.date_nls,;
          stavka = sel_zajava_b.stavka ;
          WHERE ALLTRIM(num_nls) == ALLTRIM(sel_card_acct) ;

      ENDIF

    ELSE

      SELECT sel_account
      GO TOP

      SELECT forma_1991_u

      UPDATE forma_1991_u SET ;
        num_dog = IIF(EMPTY(ALLTRIM(sel_account.num_dog)) = .F., ALLTRIM(sel_account.num_dog), RIGHT(ALLTRIM(sel_account.ref_client), 6)),;
        date_nls = sel_account.date_nls,;
        stavka = sel_account.stavka ;
        WHERE ALLTRIM(num_nls) == ALLTRIM(sel_card_acct) ;

    ENDIF

    SELECT sel_num_nls
  ENDSCAN

  WAIT CLEAR

  @ WROWS()/3,3 SAY PADC('Заполнение данными таблицы успешно завершено ... ',WCOLS())
  =INKEY(3)
  DEACTIVATE WINDOW poisk

ELSE
  =err('Внимание! В форме 1991-У записей с пустым значением полей "Номер договора" и "Дата договора" не обнаружено ...')
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SELECT DISTINCT fio_rus, num_nls, tran_nls ;
  FROM forma_1991_u ;
  WHERE EMPTY(ALLTRIM(tran_nls)) = .T. ;
  INTO CURSOR sel_tran_nls ;
  ORDER BY fio_rus

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

IF _TALLY <> 0

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Начато заполнение формы 1991-У транзитными счетами ... ',WCOLS())
  =INKEY(2)

  SELECT sel_tran_nls
  GO TOP

  SCAN

    sel_card_acct = ALLTRIM(sel_tran_nls.num_nls)
    sel_fio_rus = ALLTRIM(sel_tran_nls.fio_rus)

    WAIT WINDOW 'Клиент - ' + ALLTRIM(sel_tran_nls.fio_rus) + '  Лицевой счет - ' + ALLTRIM(sel_tran_nls.num_nls) NOWAIT

    SELECT DISTINCT fio_rus, card_acct, tran_42301 ;
      FROM account ;
      WHERE ALLTRIM(card_acct) == ALLTRIM(sel_card_acct) ;
      UNION ;
      SELECT DISTINCT fio_rus, card_acct, tran_42301 ;
      FROM acc_del ;
      WHERE ALLTRIM(card_acct) == ALLTRIM(sel_card_acct) ;
      INTO CURSOR sel_account ;
      ORDER BY fio_rus, card_acct

    IF _TALLY = 0  && Искомый счет в справочнике счетов не найден

      SELECT DISTINCT fio_rus, card_nls, tran_42301 ;
        FROM zajava_b ;
        WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_acct) ;
        INTO CURSOR sel_zajava_b ;
        ORDER BY fio_rus, card_acct

      IF _TALLY = 0  && Искомый счет в справочнике составленных заявлений не найден

        SELECT DISTINCT fio_rus, card_nls, tran_42301 ;
          FROM zajava_b ;
          WHERE ALLTRIM(fio_rus) == ALLTRIM(sel_fio_rus) ;
          INTO CURSOR sel_zajava_b ;
          ORDER BY fio_rus, card_acct

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

        IF _TALLY <> 0

          SELECT sel_zajava_b
          GO TOP

          SELECT forma_1991_u

          UPDATE forma_1991_u SET ;
            tran_nls = sel_zajava_b.tran_42301 ;
            WHERE ALLTRIM(num_nls) == ALLTRIM(sel_card_acct) ;

          SELECT vedom_ost

          UPDATE vedom_ost SET ;
            tran_42301 = sel_zajava_b.tran_42301 ;
            WHERE ALLTRIM(card_acct) == ALLTRIM(sel_card_acct) ;

        ENDIF

      ELSE

        SELECT sel_zajava_b
        GO TOP

        SELECT forma_1991_u

        UPDATE forma_1991_u SET ;
          tran_nls = sel_zajava_b.tran_42301 ;
          WHERE ALLTRIM(num_nls) == ALLTRIM(sel_card_acct) ;

        SELECT vedom_ost

        UPDATE vedom_ost SET ;
          tran_42301 = sel_zajava_b.tran_42301 ;
          WHERE ALLTRIM(card_acct) == ALLTRIM(sel_card_acct) ;

      ENDIF

    ELSE

      SELECT sel_account
      GO TOP

      SELECT forma_1991_u

      UPDATE forma_1991_u SET ;
        tran_nls = sel_account.tran_42301 ;
        WHERE ALLTRIM(num_nls) == ALLTRIM(sel_card_acct) ;

      SELECT vedom_ost

      UPDATE vedom_ost SET ;
        tran_42301 = sel_account.tran_42301 ;
        WHERE ALLTRIM(card_acct) == ALLTRIM(sel_card_acct) ;

    ENDIF

    SELECT sel_tran_nls
  ENDSCAN

  WAIT CLEAR

  @ WROWS()/3,3 SAY PADC('Заполнение данными таблицы успешно завершено ... ',WCOLS())
  =INKEY(3)
  DEACTIVATE WINDOW poisk

ELSE
  =err('Внимание! В форме 1991-У записей с пустым значением поля "Транзитный счет" не обнаружено ...')
ENDIF

SELECT (sel)
RETURN


******************************************************** PROCEDURE zapis_forma_1991 ********************************************************************

PROCEDURE zapis_forma_1991

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Внимание! Выполняется заполнение таблицы Форма 1991 У данными по лицевым счетам ... ',WCOLS())

SELECT sel_reestr
GO TOP

SCAN

  WAIT WINDOW 'Дата обработки - ' + DTOC(sel_reestr.date_vedom) + '  Клиент - ' + ALLTRIM(sel_reestr.fio_rus) + '  Лицевой счет - ' + ALLTRIM(sel_reestr.num_nls) NOWAIT

  rec_zap = DTOS(sel_reestr.date_vedom) + ALLTRIM(sel_reestr.num_nls)

  SELECT forma_1991_u
  SET ORDER TO date_nls
  GO TOP

  poisk = SEEK(rec_zap)

  IF poisk = .T.
    SCATTER MEMVAR MEMO
  ELSE
    SCATTER MEMVAR MEMO BLANK
  ENDIF

  m.branch = ALLTRIM(sel_reestr.branch)
  m.date_vedom = sel_reestr.date_vedom

  m.ref_client = ALLTRIM(sel_reestr.ref_client)

  m.fio_rus = ALLTRIM(sel_reestr.fio_rus)

  m.num_dog = ALLTRIM(sel_reestr.num_dog)

  m.date_nls = sel_reestr.date_nls
  m.srok_vklad = 0
  m.stavka = sel_reestr.stavka

  m.num_nls = sel_reestr.num_nls
  m.tran_nls = sel_reestr.tran_nls

  m.ostat_val = sel_reestr.ostat_val
  m.kurs = sel_reestr.kurs
  m.ostat_rur = sel_reestr.ostat_rur

  m.pr_dnei = 'X'

  IF EMPTY(ALLTRIM(m.num_dog)) = .T. OR EMPTY(m.date_nls) = .T.

    SELECT DISTINCT ref_client, fio_rus, card_acct, num_dog, date_nls ;
      FROM account ;
      WHERE ALLTRIM(card_acct) == ALLTRIM(m.num_nls) ;
      UNION ;
      SELECT DISTINCT ref_client, fio_rus, card_acct, num_dog, date_nls ;
      FROM acc_del ;
      WHERE ALLTRIM(card_acct) == ALLTRIM(m.num_nls) ;
      INTO CURSOR sel_account ;
      ORDER BY 2, 3

    IF _TALLY <> 0

      SELECT forma_1991_u

      m.num_dog = ALLTRIM(sel_account.num_dog)
      m.date_nls = sel_account.date_nls

    ENDIF

  ENDIF

  SELECT forma_1991_u

  IF poisk = .T.
    GATHER MEMVAR MEMO
  ELSE
    INSERT INTO forma_1991_u FROM MEMVAR
  ENDIF

  SELECT sel_reestr
ENDSCAN

WAIT CLEAR

@ WROWS()/3,3 SAY PADC('Заполнение данными таблицы успешно завершено ... ',WCOLS())
=INKEY(3)
DEACTIVATE WINDOW poisk

RETURN


************************************************************** PROCEDURE start_obazat ******************************************************************

PROCEDURE start_obazat
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

=soob('Формирование детализированной ведомости остатков по лицевым счетам -  отчет 1991 У')

DEFINE WINDOW otchet_zb AT 0,0 SIZE colvo_rows - 2, colvo_cols - 4 FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW otchet_zb CENTER

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, начата выборка данных ... ',WCOLS())

SELECT DISTINCT date_vedom AS data ;
  FROM forma_1991_u ;
  INTO CURSOR tov ;
  ORDER BY 1 DESC

IF _TALLY <> 0

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  DO FORM (put_scr) + ('seldata2.scx')

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
    DO select_obazat IN Otchet_1991_U
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  ENDIF

ELSE

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')

ENDIF

RELEASE WINDOW otchet_zb
RETURN


*************************************************************** PROCEDURE select_obazat *****************************************************************

PROCEDURE select_obazat

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, начата выборка данных ... ',WCOLS())

SELECT branch, date_vedom, ref_client, fio_rus, num_dog, num_nls, tran_nls, ostat_val, kurs, ostat_rur, pr_dnei, date_nls, srok_vklad, stavka ;
  FROM forma_1991_u ;
  WHERE BETWEEN(date_vedom, dt1, dt2) ;
  INTO CURSOR sel_forma_1991_u ;
  ORDER BY date_vedom, num_nls

IF _TALLY <> 0

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  SELECT sel_forma_1991_u
  GO TOP

  DO WHILE .T.

    tex1 = 'Табличный просмотр сформированных данных на экране'
    tex2 = 'Формирование текстового файла для головного банка'
    l_bar = 3
    =popup_9(tex1,tex2,tex3,tex4,l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      DO CASE
        CASE BAR() = 1  && Табличный просмотр сформированных данных на экране

          SELECT * ;
            FROM sel_forma_1991_u ;
            INTO CURSOR sel_forma ;
            ORDER BY date_vedom, fio_rus

          GO TOP

          BROWSE FIELDS ;
            branch:H = 'Номер филиала',;
            date_vedom:H = 'Дата',;
            ref_client:H = 'Номер клиента',;
            fio_rus:H = 'Фамилия Имя Отчество',;
            num_dog:H = 'Номер договора',;
            date_nls:H = 'Дата договора',;
            srok_vklad:H = 'Срок',;
            stavka:H = 'Ставка',;
            num_nls:H = 'Лицевой счет',;
            tran_nls:H = 'Транзитный счет',;
            ostat_val:H = 'Вход. остаток в валюте вклада':P = '999 999 999.99',;
            kurs:H = 'Курс валюты':P = '999.9999',;
            ostat_rur:H = 'Вход. остаток по курсу ЦБ':P = '999 999 999.99',;
            pr_dnei:H = 'Признак' ;
            TITLE ' Детализированная ведомость остатков по лицевым счетам -  отчет 1991 У' ;
            WINDOW otchet_zb

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

        CASE BAR() = 3  && Формирование текстового файла для головного банка

          ACTIVATE WINDOW poisk
          @ WROWS()/3-0.2,3 SAY PADC('Внимание! Начато формирование текстового файла из сформированных данных ... ',WCOLS())

          SELECT * ;
            FROM sel_forma_1991_u ;
            INTO CURSOR sel_forma ;
            ORDER BY date_vedom, num_nls

          GO TOP

          SET TEXTMERGE ON
          SET TEXTMERGE NOSHOW

          exp_file = 'ost' + RIGHT(ALLTRIM(STR(YEAR(dt2))), 1) + PADL(ALLTRIM(STR(MONTH(dt2))), 2, '0') + PADL(ALLTRIM(STR(DAY(dt2))), 2, '0') + '.' + ALLTRIM(num_branch)

          IF FILE((put_tmp) + (exp_file))
            DELETE FILE (put_tmp) + (exp_file)
          ENDIF

          _text = FCREATE((put_tmp) + (exp_file))

          IF _text < 1

            DEACTIVATE WINDOW poisk
            =err('Внимание! Ошибка при создании файла экспорта реестра ... ')
            =FCLOSE(_text)

          ELSE

            SELECT sel_forma_1991_u
            GO TOP

            SCAN

              f_num_nls = ALLTRIM(sel_forma_1991_u.num_nls) + CHR(59)
              f_tran_nls = ALLTRIM(sel_forma_1991_u.tran_nls) + CHR(59)
              f_ostat_val = ALLTRIM(TRANSFORM(sel_forma_1991_u.ostat_val, '999999999999.99')) + CHR(59)
              f_date_vedom = DTOC(sel_forma_1991_u.date_vedom) + CHR(59)
              f_priznak = ALLTRIM(sel_forma_1991_u.pr_dnei) + CHR(59)
              f_date_nls = DTOC(sel_forma_1991_u.date_nls) + CHR(59)
              f_srok_vklad = ALLTRIM(STR(sel_forma_1991_u.srok_vklad)) + CHR(59)
              f_stavka = ALLTRIM(TRANSFORM(sel_forma_1991_u.stavka, '999.99'))

              stroka_detal_record = f_num_nls + f_tran_nls + f_ostat_val + f_date_vedom + f_priznak + f_date_nls + f_srok_vklad + f_stavka

              =FPUTS(_text, (stroka_detal_record))

            ENDSCAN

            =FCLOSE(_text)

            SET TEXTMERGE OFF

            CLEAR
            @ WROWS()/3-0.2,3 SAY PADC('Формирование текстового файла успешно завершено ... ',WCOLS())
            =INKEY(3)
            DEACTIVATE WINDOW poisk

            =soob('Экспортный файл - ' + ALLTRIM(exp_file) + ' - в директорию ' + ALLTRIM(put_tmp) + ' сформирован ... ')

          ENDIF

      ENDCASE

    ELSE
      EXIT
    ENDIF

  ENDDO

ELSE

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  =err('Внимание! Данных об обязательствах банка перед вкладчиком не обнаружено ...')
ENDIF

RETURN


************************************************************** PROCEDURE start_filial ********************************************************************

PROCEDURE start_filial
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

DEFINE WINDOW otchet_zb AT 0,0 SIZE 6, colvo_cols - 40 FONT 'Times New Roman', num_schrift_v + 1  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW otchet_zb CENTER

SELECT branch, ALLTRIM(ima_p) AS ima_p, ALLTRIM(adres) AS adres ;
  FROM rekvizit ;
  INTO CURSOR sel_filial

IF _TALLY <> 0

  EDIT FIELDS ;
    branch:H = 'Номер филиала :',;
    ima_p:H = 'Полное наименование филиала :',;
    adres:H = 'Почтовый адрес филиала :' ;
    TITLE ' Сведения о филиале банка' ;
    WINDOW otchet_zb

ELSE
  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')
ENDIF

RELEASE WINDOW otchet_zb
RETURN


********************************************************** PROCEDURE scan_insert_day_01 ****************************************************************

PROCEDURE scan_insert_day_01
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

SET ESCAPE ON

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Начато формирование недостающих дат в ведомости остатков для формы 1991 У ... ',WCOLS())

SELECT DISTINCT data_home AS data ;
  FROM calendar ;
  WHERE data_home >= CTOD('01.06.2007') AND DAY(data_home) = 1 AND ;
  DTOS(data_home) < DTOS(new_data_odb) AND YEAR(data_home) <= YEAR(new_data_odb) ;
  INTO CURSOR sel_calendar ;
  ORDER BY 1

*  AND YEAR(data_home) = 2007

* BROWSE WINDOW brows

SELECT sel_calendar
GO TOP

SCAN

  dt1 = sel_calendar.data
  scan_date = sel_calendar.data

  SELECT DISTINCT date_vedom ;
    FROM vedom_ost ;
    WHERE BETWEEN(DTOS(date_vedom), DTOS(dt1), DTOS(dt1)) ;
    INTO CURSOR sel_vedom_ost ;
    ORDER BY 1

  colvo_zap = _TALLY

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  DO CASE
    CASE colvo_zap = 0 AND YEAR(dt1) = 2007

      SELECT DISTINCT date_vedom ;
        FROM vedom_ost ;
        WHERE BETWEEN(date_vedom, CTOD('01.06.2007'), CTOD('31.12.2007')) ;
        INTO CURSOR sel_vedom_ost ;
        ORDER BY 1

* BROWSE WINDOW brows

      IF _TALLY = 0

        IF FILE((put_tmp) + ('vedom_ost_2007.dbf')) AND FILE((put_tmp) + ('vedom_ost_2007.cdx'))

          SELECT DISTINCT date_vedom ;
            FROM (put_tmp) + ('vedom_ost_2007.dbf') ;
            WHERE BETWEEN(date_vedom, CTOD('01.06.2007'), CTOD('31.12.2007')) ;
            INTO CURSOR sel_vedom_ost ;
            ORDER BY 1

* BROWSE WINDOW brows

          IF _TALLY <> 0

            SELECT * ;
              FROM (put_tmp) + ('vedom_ost_2007.dbf') ;
              WHERE BETWEEN(DTOS(date_vedom), DTOS(dt1), DTOS(dt1)) ;
              INTO CURSOR sel_vedom_ost ;
              ORDER BY 1

* BROWSE WINDOW brows

            IF _TALLY = 0

              WAIT WINDOW 'Дата ' + DTOC(dt1) + ' необходимая для формы 1991-У в ведомости остатков НЕ НАЙДЕНА ... ' TIMEOUT 2

              scan_date = scan_date + 1

              DO WHILE .T.

                SELECT * ;
                  FROM (put_tmp) + ('vedom_ost_2007.dbf') ;
                  WHERE BETWEEN(DTOS(date_vedom), DTOS(scan_date), DTOS(scan_date)) ;
                  INTO CURSOR sel_vedom_ost ;
                  ORDER BY date_vedom, fio_rus

                IF _TALLY <> 0
                  EXIT
                ELSE
                  scan_date = scan_date + 1
                ENDIF

              ENDDO

              DO scan_vedom_ost IN Otchet_1991_U

            ELSE

              WAIT WINDOW 'Дата ' + DTOC(dt1) + ' необходимая для формы 1991-У в ведомости остатков НЕ НАЙДЕНА ... ' TIMEOUT 2

              DO scan_vedom_ost IN Otchet_1991_U

            ENDIF

          ELSE

            @ WROWS()/3,3 SAY PADC('Внимание! В ведомости остатков нет данных за 2007 год ... ',WCOLS())
            =INKEY(6)
            DEACTIVATE WINDOW poisk
            RETURN

          ENDIF

        ELSE

          @ WROWS()/3,3 SAY PADC('В папке \Tmp\ работающей программы файлики vedom_ost_2007.dbf и  vedom_ost_2007.cdx НЕ НАЙДЕНЫ ... ',WCOLS())
          =INKEY(6)
          @ WROWS()/3,3 SAY PADC('Скопируйте в папку \Tmp\ работающей программы файлики vedom_ost_2007.dbf и  vedom_ost_2007.cdx ... ',WCOLS())
          =INKEY(6)
          DEACTIVATE WINDOW poisk
          RETURN

        ENDIF

      ELSE

        WAIT WINDOW 'Дата ' + DTOC(dt1) + ' необходимая для формы 1991-У в ведомости остатков НЕ НАЙДЕНА ... ' TIMEOUT 2

        scan_date = scan_date + 1

        DO WHILE .T.

          SELECT * ;
            FROM vedom_ost ;
            WHERE BETWEEN(DTOS(date_vedom), DTOS(scan_date), DTOS(scan_date)) ;
            INTO CURSOR sel_vedom_ost ;
            ORDER BY date_vedom, fio_rus

          IF _TALLY <> 0
            EXIT
          ELSE
            scan_date = scan_date + 1
          ENDIF

        ENDDO

        DO scan_vedom_ost IN Otchet_1991_U

      ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE colvo_zap = 0 AND YEAR(dt1) = 2008

      WAIT WINDOW 'Дата ' + DTOC(scan_date) + ' необходимая для формы 1991-У в ведомости остатков НЕ НАЙДЕНА ... ' TIMEOUT 2

      scan_date = dt1 + 1

      DO WHILE .T.

        SELECT * ;
          FROM vedom_ost ;
          WHERE BETWEEN(DTOS(date_vedom), DTOS(scan_date), DTOS(scan_date)) ;
          INTO CURSOR sel_vedom_ost ;
          ORDER BY date_vedom, fio_rus

        IF _TALLY <> 0
          EXIT
        ELSE
          scan_date = scan_date + 1
        ENDIF

      ENDDO

      DO scan_vedom_ost IN Otchet_1991_U

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE colvo_zap = 0 AND YEAR(dt1) = 2009

      WAIT WINDOW 'Дата ' + DTOC(scan_date) + ' необходимая для формы 1991-У в ведомости остатков НЕ НАЙДЕНА ... ' TIMEOUT 2

      scan_date = dt1 + 1

      DO WHILE .T.

        SELECT * ;
          FROM vedom_ost ;
          WHERE BETWEEN(DTOS(date_vedom), DTOS(scan_date), DTOS(scan_date)) ;
          INTO CURSOR sel_vedom_ost ;
          ORDER BY date_vedom, fio_rus

        IF _TALLY <> 0
          EXIT
        ELSE
          scan_date = scan_date + 1
        ENDIF

      ENDDO

      DO scan_vedom_ost IN Otchet_1991_U

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE colvo_zap <> 0

      WAIT WINDOW 'Дата ' + DTOC(dt1) + ' необходимая для формы 1991-У в ведомости остатков НАЙДЕНА ... ' TIMEOUT 2

  ENDCASE

  SELECT sel_calendar
ENDSCAN

@ WROWS()/3,3 SAY PADC('Формирование недостающих дат в ведомости остатков успешно завершено ... ',WCOLS())
=INKEY(3)
DEACTIVATE WINDOW poisk

SET ESCAPE OFF

SELECT (sel)
RETURN


************************************************************* PROCEDURE scan_vedom_ost ****************************************************************

PROCEDURE scan_vedom_ost

SELECT sel_vedom_ost
GO TOP

SCAN

  WAIT WINDOW 'Дата ' + DTOC(dt1) + ' Клиент - ' + ALLTRIM(sel_vedom_ost.fio_rus) + ' Номер счета - ' + ALLTRIM(sel_vedom_ost.card_acct) NOWAIT

  rec = ALLTRIM(sel_vedom_ost.card_acct) + DTOS(dt1)

  SELECT vedom_ost
  SET ORDER TO acct_date
  GO TOP

  poisk = SEEK(rec)

  IF poisk = .F.

    SCATTER MEMVAR BLANK

    m.branch = ALLTRIM(sel_vedom_ost.branch)
    m.ref_client = ALLTRIM(sel_vedom_ost.ref_client)

    m.card_nls = ALLTRIM(sel_vedom_ost.card_nls)
    m.card_acct = ALLTRIM(sel_vedom_ost.card_acct)
    m.tran_42301 = ALLTRIM(sel_vedom_ost.tran_42301)

    m.val_card = ALLTRIM(sel_vedom_ost.val_card)

    m.fio_rus = ALLTRIM(sel_vedom_ost.fio_rus)

    m.date_vedom = dt1
    m.date_ostat = dt1

    m.begin_bal = sel_vedom_ost.begin_bal
    m.debit = 0.00
    m.credit = 0.00
    m.end_bal = (m.begin_bal + m.credit) - m.debit

    m.date_ost_m = dt1
    m.pr_izm_ost = .T.

    m.vxd_ost_m = sel_vedom_ost.vxd_ost_m
    m.debit_m = 0.00
    m.credit_m = 0.00
    m.isx_ost_m = (m.vxd_ost_m + m.credit_m) - m.debit_m

    INSERT INTO vedom_ost FROM MEMVAR

  ENDIF

  SELECT sel_vedom_ost
ENDSCAN

WAIT CLEAR

RETURN


************************************************************* PROCEDURE prover_vedom_ost **************************************************************

PROCEDURE prover_vedom_ost
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

SET ESCAPE ON

SELECT vedom_ost
SET ORDER TO date_fio

date_filtr = CTOD('01.06.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.07.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.08.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.09.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.10.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.11.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.12.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.01.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.02.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.03.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.04.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.05.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.06.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.07.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.08.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.09.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.10.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.11.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.12.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

date_filtr = CTOD('01.01.2009')
SET FILTER TO date_vedom = date_filtr
DO browse_vedom_ost IN Otchet_1991_U

SET ESCAPE OFF

SELECT (sel)
RETURN


************************************************************* PROCEDURE prover_forma_1991 *************************************************************

PROCEDURE prover_forma_1991
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

SET ESCAPE ON

SELECT forma_1991_u
SET ORDER TO date_fio

date_filtr = CTOD('01.06.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.07.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.08.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.09.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.10.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.11.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.12.2007')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.01.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.02.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.03.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.04.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.05.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.06.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.07.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.08.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.09.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.10.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.11.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.12.2008')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

date_filtr = CTOD('01.01.2009')
SET FILTER TO date_vedom = date_filtr
DO browse_forma_1991 IN Otchet_1991_U

SET ESCAPE OFF

SELECT (sel)
RETURN


************************************************************* PROCEDURE browse_vedom_ost **************************************************************

PROCEDURE browse_vedom_ost

COUNT TO colvo_zap
GO TOP

BROWSE FIELDS ;
  branch:H = 'Филиал',;
  date_vedom:H = 'Дата',;
  card_acct:H = 'Лицевой счет',;
  fio_rus:H = 'Фамилия Имя Отчество',;
  val_card:H = 'Валюта карты',;
  begin_bal:H = 'Вход. остаток.':P = '999 999 999.99',;
  debit:H = 'Дебет':P = '999 999 999.99',;
  credit:H = 'Кредит':P = '999 999 999.99',;
  end_bal:H = 'Исх. остаток':P = '999 999 999.99' ;
  WINDOW brows TITLE 'Количество записей за дату ' + DTOC(date_filtr) + ' равно = ' + ALLTRIM(STR(colvo_zap)) + ' записей' NODELETE NOEDIT

RETURN


*************************************************************** PROCEDURE browse_forma_1991 ***********************************************************

PROCEDURE browse_forma_1991

COUNT TO colvo_zap
GO TOP

BROWSE FIELDS ;
  branch:H = 'Филиал',;
  date_vedom:H = 'Дата',;
  num_nls:H = 'Лицевой счет',;
  tran_nls:H = 'Транзитный счет',;
  fio_rus:H = 'Фамилия Имя Отчество',;
  date_nls:H = 'Дата счета',;
  ostat_val:H = 'Вход. остаток.':P = '999 999 999.99' ;
  WINDOW brows TITLE 'Количество записей за дату ' + DTOC(date_filtr) + ' равно = ' + ALLTRIM(STR(colvo_zap)) + ' записей' NODELETE NOEDIT

RETURN


*************************************************************** PROCEDURE vedom_ost_detal **************************************************************

PROCEDURE vedom_ost_detal
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO select_forma_1991 IN Otchet_1991_U
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF _TALLY <> 0

  SELECT DISTINCT branch, date_vedom, ref_client, fio_rus, num_nls, tran_nls,;
    num_dog, date_nls, ostat_val, kurs, ostat_rur ;
    FROM vedom_ost_1991 ;
    INTO CURSOR sel_forma_1991 ;
    ORDER BY branch, date_vedom, tran_nls, fio_rus

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  DO WHILE .T.

    SELECT DISTINCT branch, date_vedom, ref_client, fio_rus, num_nls, tran_nls,;
      num_dog, date_nls, ostat_val, kurs, ostat_rur ;
      FROM sel_forma_1991 ;
      INTO CURSOR select_forma ;
      ORDER BY branch, date_vedom, tran_nls, fio_rus

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

    text1 = 'Табличный просмотр данных'
    text2 = 'Печатный отчет на принтер'
    l_bar = 3
    =popup_9(text1,text2,text3,text4,l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      DO CASE
        CASE BAR() = 1  && Табличный просмотр данных

          COUNT TO colvo_zap
          GO TOP

          BROWSE FIELDS ;
            branch:H = 'Филиал',;
            date_vedom:H = 'Дата',;
            fio_rus:H = 'Фамилия Имя Отчество',;
            num_nls:H = 'Лицевой счет',;
            tran_nls:H = 'Транзитный счет',;
            date_nls:H = 'Дата счета',;
            ostat_val:H = 'Вход. остаток.':P = '999 999 999.99' ;
            WINDOW brows TITLE 'Количество записей за дату ' + DTOC(dt1) + ' равно = ' + ALLTRIM(STR(colvo_zap)) + ' записей' NODELETE NOEDIT

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        CASE BAR() = 3  && Печатный отчет на принтер

          put_prn = IIF(BAR() <= 0, 1, BAR())

          SELECT DISTINCT branch ;
            FROM select_forma ;
            INTO CURSOR  sel_branch ;
            ORDER BY 1

          SELECT sel_branch

          SCAN

            rec = ALLTRIM(sel_branch.branch)

            SELECT DISTINCT branch, date_vedom, ref_client, fio_rus, num_nls, tran_nls,;
              num_dog, date_nls, ostat_val, kurs, ostat_rur ;
              FROM select_forma ;
              WHERE ALLTRIM(branch) == ALLTRIM(rec) ;
              INTO CURSOR  prn_ostatok ;
              ORDER BY branch, date_vedom, tran_nls, fio_rus

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

            IF _TALLY <> 0

              ima_frx = (put_frx) + ('forma_1991_u.frx')

              DO prn_form IN Visa

            ELSE
              =soob('Внимание! Данных в таблице остатков не обнаружено ... ')
            ENDIF

            SELECT sel_branch
          ENDSCAN

      ENDCASE

    ELSE
      EXIT
    ENDIF

  ENDDO

ELSE

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SELECT (sel)
RETURN


**************************************************************** PROCEDURE vedom_ost_tran **************************************************************

PROCEDURE vedom_ost_tran
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO select_forma_1991 IN Otchet_1991_U
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF _TALLY <> 0

  SELECT DISTINCT A.branch, A.date_vedom, B.name_card AS name_tran, A.tran_nls, SUM(A.ostat_val) AS ostat_val, SUM(A.ostat_rur) AS ostat_rur ;
    FROM vedom_ost_1991 A, type_liz_card B ;
    WHERE ALLTRIM(A.tran_nls) == ALLTRIM(B.tran_42301) ;
    INTO CURSOR sel_forma_1991 ;
    GROUP BY A.branch, A.date_vedom, A.tran_nls ;
    ORDER BY A.branch, A.date_vedom, A.tran_nls

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  DO WHILE .T.

    SELECT branch, date_vedom, name_tran, tran_nls, ostat_val, ostat_rur ;
      FROM sel_forma_1991 ;
      INTO CURSOR select_forma ;
      ORDER BY 1, 2, 4

    text1 = 'Табличный просмотр данных'
    text2 = 'Печатный отчет на принтер'
    l_bar = 3
    =popup_9(text1,text2,text3,text4,l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      DO CASE
        CASE BAR() = 1  && Табличный просмотр данных

          COUNT TO colvo_zap
          GO TOP

          BROWSE FIELDS ;
            branch:H = 'Филиал',;
            date_vedom:H = 'Дата',;
            tran_nls:H = 'Транзитный счет',;
            ostat_val:H = 'Вход. остаток.':P = '999 999 999.99' ;
            WINDOW brows TITLE 'Количество записей за дату ' + DTOC(dt1) + ' равно = ' + ALLTRIM(STR(colvo_zap)) + ' записей' NODELETE NOEDIT

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        CASE BAR() = 3  && Печатный отчет на принтер

          put_prn = IIF(BAR() <= 0, 1, BAR())

          SELECT DISTINCT branch ;
            FROM select_forma ;
            INTO CURSOR  sel_branch ;
            ORDER BY 1

          SELECT sel_branch

          SCAN

            rec = ALLTRIM(sel_branch.branch)

            SELECT branch, date_vedom, name_tran, tran_nls, ostat_val, ostat_rur ;
              FROM select_forma ;
              WHERE ALLTRIM(branch) == ALLTRIM(rec) ;
              INTO CURSOR  prn_ostatok ;
              ORDER BY 1, 2, 4

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

            IF _TALLY <> 0

              ima_frx = (put_frx) + ('forma_1991_tran.frx')

              DO prn_form IN Visa

            ELSE
              =soob('Внимание! Данных в таблице остатков не обнаружено ... ')
            ENDIF

            SELECT sel_branch
          ENDSCAN

          EXIT

      ENDCASE

    ELSE
      EXIT
    ENDIF

  ENDDO

ELSE

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SELECT (sel)
RETURN


**************************************************************** PROCEDURE select_forma_1991 ***********************************************************

PROCEDURE select_forma_1991

STORE arx_data_odb TO dt1, dt2

SELECT DISTINCT data_home AS data ;
  FROM calendar ;
  WHERE data_home >= CTOD('01.06.2007') AND DAY(data_home) = 1 AND DTOS(data_home) < DTOS(new_data_odb) AND YEAR(data_home) <= YEAR(new_data_odb) ;
  INTO CURSOR tov ;
  ORDER BY 1 DESC

* BROWSE WINDOW brows

IF _TALLY <> 0

  DO FORM (put_scr) + ('seldata1.scx')

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, начата выборка данных ... ',WCOLS())

    SELECT DISTINCT branch, date_vedom, ref_client, fio_rus, num_nls, tran_nls, num_dog, date_nls, ostat_val, kurs, ostat_rur ;
      FROM forma_1991_u ;
      WHERE BETWEEN(DTOS(date_vedom), DTOS(dt1), DTOS(dt2)) AND DAY(date_vedom) = 1 ;
      INTO CURSOR vedom_ost_1991 ;
      ORDER BY branch, date_vedom, tran_nls, fio_rus

    BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

  ELSE

    =INKEY(2)
    DEACTIVATE WINDOW poisk

    =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')

  ENDIF

ELSE
  =err('Внимание! Данных по лицевым счетам не обнаружено ...')
ENDIF

RETURN


************************************************************ PROCEDURE formirov_menu ******************************************************************

PROCEDURE formirov_menu

text_otchet_1991_1 = 'Формирование детализированной ведомости остатков по лицевым счетам для формы 1991 У'
text_otchet_1991_2 = 'Формирование текстового файла "Сведения об остатках на счетах банка для формы 1991 У"'
text_otchet_1991_3 = 'Формирование данных "Сведения о филиале банка, составившем отчет 1991 У"'
text_otchet_1991_4 = 'Проверка и формирование недостающих дат в ведомости остатков для формы 1991 У'
text_otchet_1991_5 = 'Проверка наличия первых чисел месяца в ведомости остатков  в диапазоне 01.06.2007 - ' + '01.' + PADL(ALLTRIM(STR(MONTH(new_data_odb))), 2, '0') + '.' + ALLTRIM(STR(YEAR(new_data_odb)))
text_otchet_1991_6 = 'Проверка наличия первых чисел месяца в таблице 1991-У  в диапазоне 01.06.2007 - ' + '01.' + PADL(ALLTRIM(STR(MONTH(new_data_odb))), 2, '0') + '.' + ALLTRIM(STR(YEAR(new_data_odb)))
text_otchet_1991_7 = 'Ведомость остатков на дату по карточным счетам в УЖЕ сформированной форме 1991 У'
text_otchet_1991_8 = 'Ведомость остатков на дату по транзитным счетам в УЖЕ сформированной форме 1991 У'
text_otchet_1991_9 = 'Заполнение регистрационных данных счетов в УЖЕ сформированной форме 1991 У'

len_otchet_1991_1 = TXTWIDTH(text_otchet_1991_1,'Times New Roman', num_schrift_v,'B')
len_otchet_1991_2 = TXTWIDTH(text_otchet_1991_2,'Times New Roman', num_schrift_v,'B')
len_otchet_1991_3 = TXTWIDTH(text_otchet_1991_3,'Times New Roman', num_schrift_v,'B')
len_otchet_1991_4 = TXTWIDTH(text_otchet_1991_4,'Times New Roman', num_schrift_v,'B')
len_otchet_1991_5 = TXTWIDTH(text_otchet_1991_5,'Times New Roman', num_schrift_v,'B')
len_otchet_1991_6 = TXTWIDTH(text_otchet_1991_6,'Times New Roman', num_schrift_v,'B')
len_otchet_1991_7 = TXTWIDTH(text_otchet_1991_7,'Times New Roman', num_schrift_v,'B')
len_otchet_1991_8 = TXTWIDTH(text_otchet_1991_8,'Times New Roman', num_schrift_v,'B')
len_otchet_1991_9 = TXTWIDTH(text_otchet_1991_9,'Times New Roman', num_schrift_v,'B')

l_bar = 22.0
row = IIF(l_bar<=25, l_bar, 30)

V_row = (WROWS()-row)/2 + 1
Lmenu = MAX(len_otchet_1991_1, len_otchet_1991_2, len_otchet_1991_3, len_otchet_1991_4, len_otchet_1991_5, len_otchet_1991_6, len_otchet_1991_7, len_otchet_1991_8, len_otchet_1991_9)
L_col = (WCOLS()-Lmenu)/2 - 0

DEFINE POPUP otchet_1991_U FROM V_row,L_col ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1 OF otchet_1991_U PROMPT text_otchet_1991_1 ;
  MESSAGE 'Формирование детализированной ведомости остатков по лицевым счетам для формы 1991 У'
DEFINE BAR 2 OF otchet_1991_U PROMPT '\-'
DEFINE BAR 3 OF otchet_1991_U PROMPT text_otchet_1991_2 ;
  MESSAGE 'Формирование текстового файла "Сведения об остатках на счетах банка для формы 1991 У"'
DEFINE BAR 4 OF otchet_1991_U PROMPT '\-'
DEFINE BAR 5 OF otchet_1991_U PROMPT text_otchet_1991_3 ;
  MESSAGE 'Формирование данных "Сведения о филиале банка, составившем отчет 1991 У"'
DEFINE BAR 6 OF otchet_1991_U PROMPT '\-'
DEFINE BAR 7 OF otchet_1991_U PROMPT text_otchet_1991_4 ;
  MESSAGE 'Проверка и формирование недостающих дат в ведомости остатков для формы 1991 У'
DEFINE BAR 8 OF otchet_1991_U PROMPT '\-'
DEFINE BAR 9 OF otchet_1991_U PROMPT text_otchet_1991_5 ;
  MESSAGE 'Проверка наличия первых чисел месяца в ведомости остатков  в диапазоне 01.06.2007 - ' + '01.' + PADL(ALLTRIM(STR(MONTH(new_data_odb))), 2, '0') + '.' + ALLTRIM(STR(YEAR(new_data_odb)))
DEFINE BAR 10 OF otchet_1991_U PROMPT '\-'
DEFINE BAR 11 OF otchet_1991_U PROMPT text_otchet_1991_6 ;
  MESSAGE 'Проверка наличия первых чисел месяца в таблице 1991-У  в диапазоне 01.06.2007 - ' + '01.' + PADL(ALLTRIM(STR(MONTH(new_data_odb))), 2, '0') + '.' + ALLTRIM(STR(YEAR(new_data_odb)))
DEFINE BAR 12 OF otchet_1991_U PROMPT '\-'
DEFINE BAR 13 OF otchet_1991_U PROMPT text_otchet_1991_7 ;
  MESSAGE 'Ведомость остатков на дату по карточным счетам в УЖЕ сформированной форме 1991 У'
DEFINE BAR 14 OF otchet_1991_U PROMPT '\-'
DEFINE BAR 15 OF otchet_1991_U PROMPT text_otchet_1991_8 ;
  MESSAGE 'Ведомость остатков на дату по транзитным счетам в УЖЕ сформированной форме 1991 У'
DEFINE BAR 16 OF otchet_1991_U PROMPT '\-'
DEFINE BAR 17 OF otchet_1991_U PROMPT text_otchet_1991_9 ;
  MESSAGE 'Заполнение регистрационных данных счетов в УЖЕ сформированной форме 1991 У'

ON SELECTION BAR 1 OF otchet_1991_U DO start_reestr IN Otchet_1991_U
ON SELECTION BAR 3 OF otchet_1991_U DO start_obazat IN Otchet_1991_U
ON SELECTION BAR 5 OF otchet_1991_U DO start_filial IN Otchet_1991_U
ON SELECTION BAR 7 OF otchet_1991_U DO scan_insert_day_01 IN Otchet_1991_U
ON SELECTION BAR 9 OF otchet_1991_U DO prover_vedom_ost IN Otchet_1991_U
ON SELECTION BAR 11 OF otchet_1991_U DO prover_forma_1991 IN Otchet_1991_U
ON SELECTION BAR 13 OF otchet_1991_U DO vedom_ost_detal IN Otchet_1991_U
ON SELECTION BAR 15 OF otchet_1991_U DO vedom_ost_tran IN Otchet_1991_U
ON SELECTION BAR 17 OF otchet_1991_U DO zapis_registr_nls IN Otchet_1991_U

RETURN


********************************************************************************************************************************************************

* SELECT DISTINCT date_vedom FROM vedom_ost ORDER BY 1 WHERE YEAR(date_vedom) = 2007
* SELECT DISTINCT date_vedom FROM vedom_ost ORDER BY 1 WHERE YEAR(date_vedom) = 2008 AND DAY(date_vedom) = 1












