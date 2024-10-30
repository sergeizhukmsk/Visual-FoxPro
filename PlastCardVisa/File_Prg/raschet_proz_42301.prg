*******************************************************
*** Процедура расчета процентов по карточным счетам ***
*******************************************************

* Используются две локальных таблицы acc_cln_dt.dbf и acc_cln_sum.dbf
* acc_cln_dt.dbf - это таблица содержит начисленные суммы по карточным счетам за каждый день.
* acc_cln_sum.dbf - это таблица содержит свернутые начисленные суммы по карточным счетам за период когда остаток по счету не менялся.

* Используются процедуры
* sel_fio - для выполнения режима редактирования данных по конкретному клиенту, оператор выбирает нужного
* клиента из предлагаемого списка, который формируется в sel_fio на основе данных содержащихся в таблицах account.dbf и acc_del.dbf
* Для выбора используется экран DO FORM (put_scr) + ('selfio_nls.scx')
* red_data - выполнение редактирования данных
* Для редактирования используется экран DO FORM (put_scr) + ('red_proz_42301.scx')
* В процедуре read_sum оператору предоставляется возможность ручного редактирования рассчитанных сумм и установки признака
* отменяющего пересчет данных в режиме расчета


****************************************************************** PROCEDURE sel_fio ********************************************************************

PROCEDURE sel_fio
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO FORM (put_scr) + ('poisk_client.scx')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO WHILE .T.

  SELECT DISTINCT SPACE(1) + fio_rus + ' - ' + val_card + ' - ' + card_acct AS fio_rus ;
    FROM account ;
    WHERE LEFT(ALLTRIM(fio_rus), len_client) == ALLTRIM(fam_client) ;
    UNION ;
    SELECT DISTINCT SPACE(1) + fio_rus + ' - ' + val_card + ' - ' + card_acct AS fio_rus ;
    FROM acc_del ;
    WHERE (end_bal <> 0 OR isx_ost_m <> 0) AND  LEFT(ALLTRIM(fio_rus), len_client) == ALLTRIM(fam_client) ;
    INTO CURSOR tov ;
    ORDER BY 1

* BROWSE WINDOW brows

* INTO CURSOR tov ;
* INTO TABLE (put_dbf) + ('tov.dbf') ;

  IF _TALLY <> 0

    DO FORM (put_scr) + ('selfio_nls.scx')

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      SELECT account
      SET ORDER TO card_acct

      poisk = SEEK(sel_card_acct)

      DO CASE
        CASE poisk = .T.  && Запись найдена в таблице открытых карт, запускаем процедуру редактирвоания карточки клиента

          ima_table = 'account'

          DO red_data IN Raschet_proz_42301

        CASE poisk = .F.  && Если запись не найдена в таблице действующих карт, то ищем в таблице закрытых карт

          SELECT acc_del
          SET ORDER TO card_acct

          poisk = SEEK(sel_card_acct)

          IF poisk = .T.  && Если запись найдена в таблице закрытых карт, то запускаем процедуру редактирвоания карточки клиента

            ima_table = 'acc_del'

            DO red_data IN Raschet_proz_42301

          ELSE  && Если запись не найдена в таблице закрытых карт, то выводим сообщение об ошибке в поиске

            =err('Внимание! Выбранный Вами клиент не найден ни в открытых, ни в закрытых картах.')

          ENDIF

      ENDCASE

    ELSE
      EXIT
    ENDIF

  ELSE
    =err('Внимание! По введенному Вами сочетанию символов клиентов не найдено ...')
    EXIT
  ENDIF

ENDDO

SELECT (sel)
RETURN


******************************************************************** PROCEDURE red_data ****************************************************************

PROCEDURE red_data

Poisk_1 = SPACE(20)
mBal_1 = SPACE(5)
mVal_1 = SPACE(3)
mKl_1 = SPACE(1)
mLiz_1 = SPACE(11)

Poisk_2 = SPACE(20)
mBal_2 = SPACE(5)
mVal_2 = SPACE(3)
mKl_2 = SPACE(1)
mLiz_2 = SPACE(11)

SELECT (ima_table)

DO WHILE .T.

  SCATTER MEMVAR MEMO

  m.data = DATETIME()
  m.ima_pk = SYS(0)

  Poisk_1 = ALLTRIM(m.card_nls)
  mBal_1 = SUBSTR(ALLTRIM(m.card_nls), 1, 5)
  mVal_1 = SUBSTR(ALLTRIM(m.card_nls), 6, 3)
  mKl_1 = SUBSTR(ALLTRIM(m.card_nls), 9, 1)
  mLiz_1 = SUBSTR(ALLTRIM(m.card_nls), 10, 11)

  Poisk_2 = ALLTRIM(m.num_47411)
  mBal_2 = SUBSTR(ALLTRIM(m.num_47411), 1, 5)
  mVal_2 = SUBSTR(ALLTRIM(m.num_47411), 6, 3)
  mKl_2 = SUBSTR(ALLTRIM(m.num_47411), 9, 1)
  mLiz_2 = SUBSTR(ALLTRIM(m.num_47411), 10, 11)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
  DO FORM (put_scr) + ('red_proz_42301.scx')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    text1 = 'Ввод пpавильный и можно выйти'
    text2 = 'Веpнуться в pежим ввода данных'
    text3 = 'Отменить введенные данные'
    l_bar = 5
    =popup_9(text1, text2, text3, text4, l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC
      DO CASE
        CASE BAR() = 1

          SELECT (ima_table)

          BEGIN TRANSACTION

          GATHER MEMVAR MEMO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          SELECT (ima_table)

          UPDATE (ima_table) SET ;
            date_nls = m.date_nls,;
            card_nls = m.card_nls,;
            stavka = m.stavka,;
            num_dog = m.num_dog,;
            data_dog = m.data_dog,;
            data_proz = m.data_proz,;
            num_47411 = m.num_47411 ;
            WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_nls)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          SELECT acc_cln_dt_pol

          UPDATE acc_cln_dt_pol SET ;
            card_nls = m.card_nls,;
            num_dog = m.num_dog,;
            data_dog = m.data_dog,;
            data_proz = m.data_proz ;
            WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_nls)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          SELECT acc_cln_sum_pol

          UPDATE acc_cln_sum_pol SET ;
            card_nls = m.card_nls,;
            num_dog = m.num_dog,;
            data_dog = m.data_dog,;
            data_proz = m.data_proz ;
            WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_nls)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          END TRANSACTION

          EXIT

        CASE BAR() = 3
          LOOP

        CASE BAR() = 5
          EXIT

      ENDCASE
    ENDIF

  ELSE
    EXIT
  ENDIF
ENDDO

RETURN


**************************************************************** PROCEDURE scan_data_dog ***************************************************************

PROCEDURE scan_data_dog
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

SELECT account
SET ORDER TO fio_rus
GO TOP

SCAN

  WAIT WINDOW 'Клиент - ' + ALLTRIM(account.fio_rus) + '  Счет клиента - ' + ALLTRIM(account.card_nls) NOWAIT

  SCATTER MEMVAR

  sel_card_nls = ALLTRIM(m.card_nls)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  BEGIN TRANSACTION

  UPDATE acc_cln_dt_pol SET ;
    num_dog = m.num_dog,;
    data_dog = m.data_dog,;
    data_proz = m.data_proz ;
    WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_nls)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  UPDATE acc_cln_sum_pol SET ;
    num_dog = m.num_dog,;
    data_dog = m.data_dog,;
    data_proz = m.data_proz ;
    WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_nls)

  END TRANSACTION

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  SELECT account
ENDSCAN

WAIT CLEAR

SELECT (sel)
RETURN


********************************************************************* PROCEDURE read_sum **************************************************************

PROCEDURE read_sum
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

text1 = 'Выбор данных для редактирования по коду валюты RUR'
text2 = 'Выбор данных для редактирования по коду валюты USD'
text3 = 'Выбор данных для редактирования по коду валюты EUR'
l_bar = 5
=popup_9(text1, text2, text3, text4, l_bar)
ACTIVATE POPUP vibor
RELEASE POPUPS vibor

IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

  DO CASE
    CASE BAR() = 1  && Выбор данных для редактирования по коду валюты RUR

      sel_val_card = 'RUR'

    CASE BAR() = 3  && Выбор данных для редактирования по коду валюты USD

      sel_val_card = 'USD'

    CASE BAR() = 5  && Выбор данных для редактирования по коду валюты EUR

      sel_val_card = 'EUR'

  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
  DO FORM (put_scr) + ('poisk_client.scx')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  DO WHILE .T.

    SELECT DISTINCT SPACE(1) + fio_rus + ' - ' + ALLTRIM(val_card) + ' - ' + ALLTRIM(card_nls) AS fio_rus ;
      FROM acc_cln_sum_pol ;
      WHERE stavka <> 0 AND LEFT(ALLTRIM(fio_rus), len_client) == ALLTRIM(fam_client) AND ALLTRIM(val_card) == ALLTRIM(sel_val_card) ;
      INTO CURSOR tov ;
      ORDER BY 1

* BROWSE WINDOW brows

    DO FORM (put_scr) + ('selfio_nls.scx')

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      SELECT DISTINCT data_pak AS data ;
        FROM acc_cln_sum_pol ;
        WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_nls) ;
        INTO CURSOR tov ;
        ORDER BY 1 DESC

      IF _TALLY <> 0

        DO FORM (put_scr) + ('seldata1.scx')

        IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

          SELECT acc_cln_sum_pol
          SET ORDER TO fio
          SET FILTER TO data_pak = dt1 AND ALLTRIM(card_nls) == ALLTRIM(sel_card_nls)
          GO TOP

          DEFINE WINDOW brow_proz  AT 0,0 SIZE  colvo_rows - 4, colvo_cols - 30 FONT "Courier New Cyr", 10  STYLE 'B' ;
            NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
          MOVE WINDOW brow_proz CENTER

          SET CURSOR ON

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

          EDIT FIELDS ;
            sum_dobav:H = ' Сумма процентов :':P = '9 999.99',;
            pr:H = ' Запрет на изменение данных :',;
            card_nls:H = ' Номер счета :',;
            fio_rus:H = ' Фамилия Имя Отчество :',;
            val_card:H = ' Тип карты :',;
            num_dog:H = ' Номер договора :',;
            data_proz:H = ' Дата поступления средств :',;
            kurs:H = ' Курс:':P='999.9999',;
            stavka:H = ' Ставка :':P='999.99',;
            sum_ostat:H = ' Сумма остатка по счету :':P='999 999.99',;
            kvartal:H = ' Номер квартала :',;
            mes_kvart:H = ' Номер месяца в квартале :',;
            dt_hom_end:H = ' Первая дата в квартале :',;
            data_home:H = ' Начальная дата расчета :',;
            data_end:H = ' Конечная дата расчета :',;
            den_mes:H = ' Дней в месяце :' ;
            TITLE ' Редактирование данных по рассчитанным процентам ' ;
            WINDOW brow_proz

*     sum_rasch:H = ' Сумма на счет расчетов :':P='9 999.99',;
*     sum_rasxod:H = ' Сумма на расходы банка :':P='9 999.99',;
*     sum_nalog:H = ' Сумма подоходного налога :':P='9 999.99',;
*     sum_viplat:H = ' Сумма процентов к выплате :':P='9 999.99' ;

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

          REPLACE ALL sum_viplat WITH sum_dobav

          RELEASE WINDOW brow_proz

          SELECT acc_cln_sum_pol
          SET FILTER TO
          SET ORDER TO card_nls
          GO TOP

          SET CURSOR OFF

        ENDIF

      ELSE
        =soob('Внимание! В таблице рассчитанных процентов не обнаружено.')
      ENDIF

    ELSE
      EXIT
    ENDIF

  ENDDO

ENDIF

SELECT(sel)
RETURN


****************************************************************** PROCEDURE del_data ******************************************************************

PROCEDURE del_data
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

=soob('Внимание! Выполнение данной процедуры требует монопольного режима работы с таблицами ... ')

tex1 = 'В ы п о л н и т ь  удаление данных по рассчитанных процентам'
tex2 = 'О т к а з а т ь с я  от удаления данных по рассчитанных процентам'
l_bar = 3
=popup_9(tex1,tex2,tex3,tex4,l_bar)
ACTIVATE POPUP vibor
RELEASE POPUPS vibor

IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

  IF BAR() = 1  && В ы п о л н и т ь  удаление данных по рассчитанных процентам

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      text1 = 'Удаление данных выборочно по одному клиенту и диапазону дат расчета'
      text2 = 'Удаление данных списком по всем клиентам и диапазону дат расчета'
      l_bar = 3
      =popup_9(text1, text2, text3, text4, l_bar)
      ACTIVATE POPUP vibor
      RELEASE POPUPS vibor

      IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

        put_vibor = BAR()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

        DO CASE
          CASE put_vibor = 1  && Выборка данных выборочно по одному клиенту и диапазону дат расчета

            num_mesaz = MONTH(new_data_odb)
            num_god = ALLTRIM(STR(YEAR(new_data_odb)))

            DO vid_kvartal_1 IN Raschet_proz_42301

            vvod_data_end = new_data_odb

            DO FORM (put_scr) + ('sel_data_proz.scx')

            IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

              num_mesaz = MONTH(vvod_data_end)
              num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

              DO vid_kvartal_2 IN Raschet_proz_42301

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

              DO WHILE .T.

                DO FORM (put_scr) + ('poisk_client.scx')

                SELECT DISTINCT SPACE(1) + fio_rus + ' - ' + val_card + ' - ' + card_nls AS fio_rus ;
                  FROM acc_cln_sum_pol ;
                  WHERE BETWEEN(data_end, vvod_data_home, vvod_data_end) AND LEFT(ALLTRIM(fio_rus), len_client) == ALLTRIM(fam_client) ;
                  INTO CURSOR tov ;
                  ORDER BY 1

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                IF _TALLY <> 0

                  DO FORM (put_scr) + ('selfio_nls.scx')

                  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

                    tim1 = SECONDS()

                    ACTIVATE WINDOW poisk
                    @ WROWS()/3,3 SAY PADC('Минуточку терпения, начато удаление данных в таблице с расчетами по дням ... ',WCOLS())

                    SELECT acc_cln_dt_pol
                    SET ORDER TO card_nls

                    SELECT DISTINCT branch, ref_client, fio_rus, card_nls, data_home, data_end ;
                      FROM acc_cln_dt_pol ;
                      WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_nls) AND BETWEEN(data_end, vvod_data_home, vvod_data_end) ;
                      INTO CURSOR sel_prozent ;
                      ORDER BY fio_rus, data_home, data_end

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                    colvo_zap = _TALLY

                    IF colvo_zap <> 0  && Если количество найденных записей в таблицах не равно нулю, то запускаем процедуру пометки на удаление рассчитанных процентов

                      SELECT acc_cln_dt_pol
                      USE
                      USE (put_dbf) + ('acc_cln_dt.dbf') ALIAS acc_cln_dt_pol EXCLUSIVE
                      SET ORDER TO card_nls

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                      SELECT sel_prozent

                      SCAN

                        sel_fio_rus = ALLTRIM(sel_prozent.fio_rus)
                        sel_card_nls = ALLTRIM(sel_prozent.card_nls)

                        WAIT WINDOW 'Клиент - ' + ALLTRIM(sel_fio_rus) + '  Счет клиента - ' + ALLTRIM(sel_card_nls) + '.  Данные в диапазоне ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) NOWAIT

                        rec = ALLTRIM(sel_prozent.card_nls)  +  DTOS(sel_prozent.data_home) + DTOS(sel_prozent.data_end)

                        SELECT acc_cln_dt_pol

                        poisk = SEEK(rec)

                        IF poisk = .T.
                          DELETE NEXT 1
                        ENDIF

                        SELECT sel_prozent
                      ENDSCAN

                      WAIT CLEAR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                      SELECT acc_cln_dt_pol
                      PACK
                      USE
                      USE (put_dbf) + ('acc_cln_dt.dbf') ALIAS acc_cln_dt_pol SHARED
                      SET ORDER TO card_nls

                      ACTIVATE WINDOW poisk
                      @ WROWS()/3,3 SAY PADC('Клиент ' + ALLTRIM(sel_fio_rus) + ' из таблицы с расчетом по дням удален ...',WCOLS())
                      =INKEY(2)

                      tim2 = SECONDS()

                      ACTIVATE WINDOW poisk
                      @ WROWS()/3,3 SAY PADC('Работа с таблицей с расчетом по дням успешно завершена.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
                      =INKEY(2)
                      DEACTIVATE WINDOW poisk

                    ELSE

                      DEACTIVATE WINDOW poisk

                      =err('Внимание! Данных в таблице с расчетом по дням не обнаружено ...')

                    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                    ACTIVATE WINDOW poisk
                    @ WROWS()/3,3 SAY PADC('Минуточку терпения, начато удаление данных в таблице с итоговыми расчетами ... ',WCOLS())
                    =INKEY(2)

                    tim1 = SECONDS()

                    SELECT acc_cln_sum_pol
                    SET ORDER TO card_nls

                    SELECT DISTINCT branch, ref_client, fio_rus, card_nls, data_home, data_end ;
                      FROM acc_cln_sum_pol ;
                      WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_nls) AND BETWEEN(data_end, vvod_data_home, vvod_data_end) ;
                      INTO CURSOR sel_prozent ;
                      ORDER BY fio_rus, data_home, data_end

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                    colvo_zap = _TALLY

                    IF colvo_zap <> 0  && Если количество найденных записей в таблицах не равно нулю, то запускаем процедуру пометки на удаление рассчитанных процентов

                      SELECT acc_cln_sum_pol
                      USE
                      USE (put_dbf) + ('acc_cln_sum.dbf') ALIAS acc_cln_sum_pol EXCLUSIVE
                      SET ORDER TO card_nls

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                      SELECT sel_prozent

                      SCAN

                        sel_fio_rus = ALLTRIM(sel_prozent.fio_rus)
                        sel_card_nls = ALLTRIM(sel_prozent.card_nls)

                        WAIT WINDOW 'Клиент - ' + ALLTRIM(sel_fio_rus) + '  Счет клиента - ' + ALLTRIM(sel_card_nls) + '.  Данные в диапазоне ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) NOWAIT

                        rec = ALLTRIM(sel_prozent.card_nls)  +  DTOS(sel_prozent.data_home) + DTOS(sel_prozent.data_end)

                        SELECT acc_cln_sum_pol

                        poisk = SEEK(rec)

                        IF poisk = .T.
                          DELETE NEXT 1
                        ENDIF

                        SELECT sel_prozent
                      ENDSCAN

                      WAIT CLEAR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                      SELECT acc_cln_sum_pol
                      PACK
                      USE
                      USE (put_dbf) + ('acc_cln_sum.dbf') ALIAS acc_cln_sum_pol SHARED
                      SET ORDER TO card_nls

                      ACTIVATE WINDOW poisk
                      @ WROWS()/3,3 SAY PADC('Клиент ' + ALLTRIM(sel_fio_rus) + ' из таблицы с итоговым расчетом удален ...',WCOLS())
                      =INKEY(2)

                      tim2 = SECONDS()

                      ACTIVATE WINDOW poisk
                      @ WROWS()/3,3 SAY PADC('Работа с таблицей с итоговым расчетом успешно завершена.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
                      =INKEY(2)
                      DEACTIVATE WINDOW poisk

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
                      DO del_acc_cln_null IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                    ELSE

                      DEACTIVATE WINDOW poisk

                      =err('Внимание! Данных в таблице с итоговым расчетом не обнаружено ...')

                    ENDIF

                  ELSE
                    EXIT
                  ENDIF

                ELSE

                  DEACTIVATE WINDOW poisk

                  EXIT
                ENDIF

              ENDDO

            ENDIF

* ====================================================================================================================================== *

          CASE put_vibor = 3  && Выборка данных списком по всем клиентам и диапазону дат расчета

            num_mesaz = MONTH(new_data_odb)
            num_god = ALLTRIM(STR(YEAR(new_data_odb)))

            DO vid_kvartal_1 IN Raschet_proz_42308

            DO FORM (put_scr) + ('sel_data_proz.scx')

            IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

              num_mesaz = MONTH(vvod_data_end)
              num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

              DO vid_kvartal_2 IN Raschet_proz_42308

              tim1 = SECONDS()

              ACTIVATE WINDOW poisk
              @ WROWS()/3,3 SAY PADC('Минуточку терпения, идет удаление данных в таблице с расчетами по дням ... ',WCOLS())

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

              SELECT acc_cln_dt_pol
              SET ORDER TO card_nls

              SELECT DISTINCT branch, ref_client, fio_rus, card_nls, data_home, data_end ;
                FROM acc_cln_dt_pol ;
                WHERE BETWEEN(data_home, vvod_data_home, vvod_data_end) ;
                INTO CURSOR sel_prozent ;
                ORDER BY fio_rus, data_home, data_end

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

              colvo_zap = _TALLY

              IF colvo_zap <> 0  && Если количество найденных записей в таблицах не равно нулю, то запускаем процедуру пометки на удаление рассчитанных процентов

                SELECT acc_cln_dt_pol
                USE
                USE (put_dbf) + ('acc_cln_dt.dbf') ALIAS acc_cln_dt_pol EXCLUSIVE
                SET ORDER TO card_nls

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                SELECT sel_prozent

                SCAN

                  sel_fio_rus = ALLTRIM(sel_prozent.fio_rus)
                  sel_card_nls = ALLTRIM(sel_prozent.card_nls)

                  WAIT WINDOW 'Клиент - ' + ALLTRIM(sel_fio_rus) + '  Счет клиента - ' + ALLTRIM(sel_card_nls) + '.  Данные в диапазоне ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) NOWAIT

                  rec = ALLTRIM(sel_prozent.card_nls)  +  DTOS(sel_prozent.data_home) + DTOS(sel_prozent.data_end)

                  SELECT acc_cln_dt_pol

                  poisk = SEEK(rec)

                  IF poisk = .T.
                    DELETE NEXT 1
                  ENDIF

                  SELECT sel_prozent
                ENDSCAN

                WAIT CLEAR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                SELECT acc_cln_dt_pol
                PACK
                USE
                USE (put_dbf) + ('acc_cln_dt.dbf') ALIAS acc_cln_dt_pol SHARED
                SET ORDER TO card_nls

                ACTIVATE WINDOW poisk
                @ WROWS()/3,3 SAY PADC('Данные в диапазоне ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) + ' из таблицы с расчетом по дням удалены ...',WCOLS())
                =INKEY(2)

                tim2 = SECONDS()

                ACTIVATE WINDOW poisk
                @ WROWS()/3,3 SAY PADC('Работа с таблицей с расчетом по дням успешно завершена.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
                =INKEY(2)
                DEACTIVATE WINDOW poisk

              ELSE

                DEACTIVATE WINDOW poisk

                =err('Внимание! Данных в таблице с расчетом по дням не обнаружено ...')

              ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

              ACTIVATE WINDOW poisk
              @ WROWS()/3,3 SAY PADC('Минуточку терпения, начато удаление данных в таблице с итоговыми расчетами ... ',WCOLS())
              =INKEY(2)

              tim1 = SECONDS()

              SELECT acc_cln_sum_pol
              SET ORDER TO card_nls

              SELECT DISTINCT branch, ref_client, fio_rus, card_nls, data_home, data_end ;
                FROM acc_cln_sum_pol ;
                WHERE BETWEEN(data_home, vvod_data_home, vvod_data_end) ;
                INTO CURSOR sel_prozent ;
                ORDER BY fio_rus, data_home, data_end

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

              colvo_zap = _TALLY

              IF colvo_zap <> 0  && Если количество найденных записей в таблицах не равно нулю, то запускаем процедуру пометки на удаление рассчитанных процентов

                SELECT acc_cln_sum_pol
                USE
                USE (put_dbf) + ('acc_cln_sum.dbf') ALIAS acc_cln_sum_pol EXCLUSIVE
                SET ORDER TO card_nls

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                SELECT sel_prozent

                SCAN

                  sel_fio_rus = ALLTRIM(sel_prozent.fio_rus)
                  sel_card_nls = ALLTRIM(sel_prozent.card_nls)

                  WAIT WINDOW 'Клиент - ' + ALLTRIM(sel_fio_rus) + '  Счет клиента - ' + ALLTRIM(sel_card_nls) + '.  Данные в диапазоне ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) NOWAIT

                  rec = ALLTRIM(sel_prozent.card_nls)  +  DTOS(sel_prozent.data_home) + DTOS(sel_prozent.data_end)

                  SELECT acc_cln_sum_pol

                  poisk = SEEK(rec)

                  IF poisk = .T.
                    DELETE NEXT 1
                  ENDIF

                  SELECT sel_prozent
                ENDSCAN

                WAIT CLEAR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                SELECT acc_cln_sum_pol
                PACK

                USE
                USE (put_dbf) + ('acc_cln_sum.dbf') ALIAS acc_cln_sum_pol SHARED
                SET ORDER TO card_nls

                ACTIVATE WINDOW poisk
                @ WROWS()/3,3 SAY PADC('Данные в диапазоне ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) + ' из таблицы с итоговым расчетом удалены ...',WCOLS())
                =INKEY(2)

                tim2 = SECONDS()

                ACTIVATE WINDOW poisk
                @ WROWS()/3,3 SAY PADC('Работа с таблицей с итоговым расчетом успешно завершена.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
                =INKEY(2)
                DEACTIVATE WINDOW poisk

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
                DO del_acc_cln_null IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

              ELSE

                DEACTIVATE WINDOW poisk

                =err('Внимание! Данных в таблице с итоговым расчетом не обнаружено ...')

              ENDIF
            ENDIF

        ENDCASE

      ENDIF
    ENDIF
  ENDIF
ENDIF

SELECT(sel)
RETURN


*************************************************************** PROCEDURE del_acc_cln_null **************************************************************

PROCEDURE del_acc_cln_null

=soob('Начинается подготовка данных к удалению из временных таблиц ...')

IF FILE((put_dbf) + ('acc_cln_dt_null.dbf'))

  IF FILE((put_dbf) + ('acc_cln_dt_null.cdx'))

    IF NOT USED('acc_cln_dt_null_baze')

      SELECT 0
      USE (put_dbf) + ('acc_cln_dt_null.dbf') ALIAS acc_cln_dt_null_baze EXCLUSIVE
      DELETE ALL
      PACK
      USE

    ELSE

      SELECT acc_cln_dt_null_baze
      USE

    ENDIF

  ELSE
    =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('acc_cln_dt_null.cdx') + ' не найден.')
    RETURN
  ENDIF

ELSE
  =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('acc_cln_dt_null.dbf') + ' не найден.')
  RETURN
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

=soob('Данные из временной таблицы с расчетом по дням удалены ...')

IF FILE((put_dbf) + ('acc_cln_sum_null.dbf'))

  IF FILE((put_dbf) + ('acc_cln_sum_null.cdx'))

    IF NOT USED('acc_cln_sum_null_baze')

      SELECT 0
      USE (put_dbf) + ('acc_cln_sum_null.dbf') ALIAS acc_cln_sum_null_baze EXCLUSIVE
      DELETE ALL
      PACK
      USE

    ELSE

      SELECT acc_cln_sum_null_baze
      USE

    ENDIF

  ELSE
    =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('acc_cln_sum_null.cdx') + ' не найден.')
    RETURN
  ENDIF

ELSE
  =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('acc_cln_sum_null.dbf') + ' не найден.')
  RETURN
ENDIF

=soob('Данные из временной таблицы с итоговым расчетом удалены ...')

RETURN


*************************************************************** PROCEDURE vid_kvartal_1 *****************************************************************

PROCEDURE vid_kvartal_1

DO CASE
  CASE BETWEEN(num_mesaz, 1, 3) = .T.

    num_kvartal = 1
    data_home_end = CTOD('01.01.' + num_god)

    DO CASE
      CASE num_mesaz = 1
        num_mesaz_kvartal = 1
        vvod_data_home = CTOD('01.01.' + num_god)
        vvod_data_end = CTOD('31.01.' + num_god)

      CASE num_mesaz = 2
        num_mesaz_kvartal = 2
        vvod_data_home = CTOD('01.02.' + num_god)
        vvod_data_end = IIF(colvo_den_god = 365,CTOD('28.02.' + num_god),IIF(colvo_den_god = 366,CTOD('29.02.' + num_god),DATE()))

      CASE num_mesaz = 3
        num_mesaz_kvartal = 3
        vvod_data_home = CTOD('01.03.' + num_god)
        vvod_data_end = CTOD('31.03.' + num_god)

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE BETWEEN(num_mesaz, 4, 6) = .T.

    num_kvartal = 2
    data_home_end = CTOD('01.04.' + num_god)

    DO CASE
      CASE num_mesaz = 4
        num_mesaz_kvartal = 1
        vvod_data_home = CTOD('01.04.' + num_god)
        vvod_data_end = CTOD('30.04.' + num_god)

      CASE num_mesaz = 5
        num_mesaz_kvartal = 2
        vvod_data_home = CTOD('01.05.' + num_god)
        vvod_data_end = CTOD('31.05.' + num_god)

      CASE num_mesaz = 6
        num_mesaz_kvartal = 3
        vvod_data_home = CTOD('01.06.' + num_god)
        vvod_data_end = CTOD('30.06.' + num_god)

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE BETWEEN(num_mesaz, 7, 9) = .T.
    num_kvartal = 3
    data_home_end = CTOD('01.07.' + num_god)

    DO CASE
      CASE num_mesaz = 7
        num_mesaz_kvartal = 1
        vvod_data_home = CTOD('01.07.' + num_god)
        vvod_data_end = CTOD('31.07.' + num_god)

      CASE num_mesaz = 8
        num_mesaz_kvartal = 2
        vvod_data_home = CTOD('01.08.' + num_god)
        vvod_data_end = CTOD('31.08.' + num_god)

      CASE num_mesaz = 9
        num_mesaz_kvartal = 3
        vvod_data_home = CTOD('01.09.' + num_god)
        vvod_data_end = CTOD('30.09.' + num_god)

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE BETWEEN(num_mesaz, 10, 12) = .T.
    num_kvartal = 4
    data_home_end = CTOD('01.10.' + num_god)

    DO CASE
      CASE num_mesaz = 10
        num_mesaz_kvartal = 1
        vvod_data_home = CTOD('01.10.' + num_god)
        vvod_data_end = CTOD('31.10.' + num_god)

      CASE num_mesaz = 11
        num_mesaz_kvartal = 2
        vvod_data_home = CTOD('01.11.' + num_god)
        vvod_data_end = CTOD('30.11.' + num_god)

      CASE num_mesaz = 12
        num_mesaz_kvartal = 3
        vvod_data_home = CTOD('01.12.' + num_god)
        vvod_data_end = CTOD('31.12.' + num_god)

    ENDCASE

ENDCASE

RETURN


**************************************************************** PROCEDURE vid_kvartal_2 ****************************************************************

PROCEDURE vid_kvartal_2

DO CASE
  CASE BETWEEN(num_mesaz, 1, 3) = .T.

    num_kvartal = 1
    data_home_end = CTOD('01.01.' + num_god)

    DO CASE
      CASE num_mesaz = 1
        num_mesaz_kvartal = 1

      CASE num_mesaz = 2
        num_mesaz_kvartal = 2

      CASE num_mesaz = 3
        num_mesaz_kvartal = 3

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE BETWEEN(num_mesaz, 4, 6) = .T.
    num_kvartal = 2
    data_home_end = CTOD('01.04.' + num_god)

    DO CASE
      CASE num_mesaz = 4
        num_mesaz_kvartal = 1

      CASE num_mesaz = 5
        num_mesaz_kvartal = 2

      CASE num_mesaz = 6
        num_mesaz_kvartal = 3

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE BETWEEN(num_mesaz, 7, 9) = .T.
    num_kvartal = 3
    data_home_end = CTOD('01.07.' + num_god)

    DO CASE
      CASE num_mesaz = 7
        num_mesaz_kvartal = 1

      CASE num_mesaz = 8
        num_mesaz_kvartal = 2

      CASE num_mesaz = 9
        num_mesaz_kvartal = 3

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE BETWEEN(num_mesaz, 10, 12) = .T.
    num_kvartal = 4
    data_home_end = CTOD('01.10.' + num_god)

    DO CASE
      CASE num_mesaz = 10
        num_mesaz_kvartal = 1

      CASE num_mesaz = 11
        num_mesaz_kvartal = 2

      CASE num_mesaz = 12
        num_mesaz_kvartal = 3

    ENDCASE

ENDCASE

RETURN


************************************************************** PROCEDURE vibor_proz *********************************************************************

PROCEDURE vibor_proz
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

DO CASE
  CASE pr_table_proz = .T.  && Временные процентные таблицы создаются на сервере

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Временные таблицы будут созданы на СЕРВЕРЕ.',WCOLS())
    =INKEY(3)
    DEACTIVATE WINDOW poisk

  CASE pr_table_proz = .F.  && Временные процентные таблицы создаются на локальной машине

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Временные таблицы будут созданы на ЛОКАЛЬНОЙ МАШИНЕ.',WCOLS())
    =INKEY(3)
    DEACTIVATE WINDOW poisk

    IF ROUND(VAL(SYS(2020)) / 1000000, 0) < 100.00

      ACTIVATE WINDOW poisk
      @ WROWS()/3,3 SAY PADC('На диске "С" должно быть свободного места примерно 100 Мбайт.',WCOLS())
      =INKEY(3)
      DEACTIVATE WINDOW poisk

    ENDIF

ENDCASE

vvod_data_pak = MaxDataMes(DATE())
put_bar = 1

num_mesaz = MONTH(new_data_odb)
num_god = ALLTRIM(STR(YEAR(new_data_odb)))

DO vid_kvartal_1 IN Raschet_proz_42301

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
DO FORM (put_scr) + ('poisk_client.scx')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

SELECT DISTINCT SPACE(1) + fio_rus + ' - ' + val_card + ' - ' + card_nls AS fio_rus ;
  FROM account ;
  WHERE stavka <> 0 AND LEFT(ALLTRIM(fio_rus), len_client) == ALLTRIM(fam_client) ;
  UNION ;
  SELECT DISTINCT SPACE(1) + fio_rus + ' - ' + val_card + ' - ' + card_nls AS fio_rus ;
  FROM acc_del ;
  WHERE stavka <> 0 AND (end_bal <> 0 OR isx_ost_m <> 0) AND LEFT(ALLTRIM(fio_rus), len_client) == ALLTRIM(fam_client) ;
  INTO CURSOR tov ;
  ORDER BY 1

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

IF _TALLY <> 0

  DO FORM (put_scr) + ('selfio_nls.scx')

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    vvod_data_end = new_data_odb

    DO FORM (put_scr) + ('vvod_data_proz.scx')

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

* WAIT WINDOW 'sel_card_nls = ' + ALLTRIM(sel_card_nls)

      num_mesaz = MONTH(vvod_data_end)
      num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

      DO vid_kvartal_2 IN Raschet_proz_42301

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      colvo_den_god = GOMONTH(CTOD('01.01.' + ALLTRIM(STR(YEAR(vvod_data_end)))), 12) - CTOD('01.01.' + ALLTRIM(STR(YEAR(vvod_data_end))))
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW DTOC(vvod_data_home) + '   ' + DTOC(vvod_data_end)

      SELECT DISTINCT branch, ref_client, ref_card, fio_rus, data_proz, card_nls, card_acct,;
        val_card, name_card, stavka, 000.00 AS sum_ostat, num_47411, pr_odb, num_dog, data_dog ;
        FROM account ;
        WHERE data_proz <= vvod_data_end AND stavka <> 0 AND ALLTRIM(card_nls) == ALLTRIM(sel_card_nls) AND ALLTRIM(vid_card) == '1' ;
        INTO CURSOR sel_account ;
        ORDER BY branch, fio_rus

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      IF _TALLY = 0  && Если запись не найдена в таблице действующих карт, то ищем в таблице закрытых карт

        SELECT DISTINCT branch, ref_client, ref_card, fio_rus, data_proz, card_nls, card_acct,;
          val_card, name_card, stavka, 000.00 AS sum_ostat, num_47411, pr_odb, num_dog, data_dog ;
          FROM acc_del ;
          WHERE data_proz <= vvod_data_end AND stavka <> 0 AND ALLTRIM(card_nls) == ALLTRIM(sel_card_nls) AND ALLTRIM(vid_card) == '1' ;
          INTO CURSOR sel_account ;
          ORDER BY branch, fio_rus

        IF _TALLY <> 0  && Если запись найдена в таблице закрытых карт, то запускаем расчет процентов

          DO raschet IN Raschet_proz_42301

        ELSE  && Если запись не найдена в таблице закрытых карт, то выводим сообщение об ошибке в поиске

          =err('Внимание! Выбранный Вами клиент не найден ни в открытых, ни в закрытых картах.')

        ENDIF

      ELSE  && Если запись найдена в таблице действующих карт, то запускаем расчет процентов

        DO raschet IN Raschet_proz_42301

      ENDIF
    ENDIF
  ENDIF

ELSE
  =err('Внимание! По введенному Вами сочетанию символов клиентов не найдено ...')
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
colvo_den_god = GOMONTH(CTOD('01.01.' + ALLTRIM(STR(YEAR(new_data_odb)))), 12) - CTOD('01.01.' + ALLTRIM(STR(YEAR(new_data_odb))))
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SELECT(sel)
RETURN


**************************************************************** PROCEDURE sum_proz ********************************************************************

PROCEDURE sum_proz
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

DO CASE
  CASE pr_table_proz = .T.  && Временные процентные таблицы создаются на сервере

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Временные таблицы будут созданы на СЕРВЕРЕ.',WCOLS())
    =INKEY(3)
    DEACTIVATE WINDOW poisk

  CASE pr_table_proz = .F.  && Временные процентные таблицы создаются на локальной машине

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Временные таблицы будут созданы на ЛОКАЛЬНОЙ МАШИНЕ.',WCOLS())
    =INKEY(3)
    DEACTIVATE WINDOW poisk

    IF ROUND(VAL(SYS(2020)) / 1000000, 0) < 100.00

      ACTIVATE WINDOW poisk
      @ WROWS()/3,3 SAY PADC('На диске "С" должно быть свободного места примерно 100 Мбайт.',WCOLS())
      =INKEY(3)
      DEACTIVATE WINDOW poisk

    ENDIF

ENDCASE

put_bar = 1

num_mesaz = MONTH(new_data_odb)
num_god = ALLTRIM(STR(YEAR(new_data_odb)))

DO vid_kvartal_1 IN Raschet_proz_42301

DO FORM (put_scr) + ('vvod_data_proz.scx')

IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

  vvod_data_pak = MaxDataMes(vvod_data_end)

  num_mesaz = MONTH(vvod_data_end)
  num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

  DO vid_kvartal_2 IN Raschet_proz_42301

* WAIT WINDOW DTOC(vvod_data_home) + '   ' + DTOC(vvod_data_end)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
  colvo_den_god = GOMONTH(CTOD('01.01.' + ALLTRIM(STR(YEAR(vvod_data_end)))), 12) - CTOD('01.01.' + ALLTRIM(STR(YEAR(vvod_data_end))))
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

* ----------------------------------------------------- Выбор схемы рассчета процентов полным списком или отдельно по проектам ------------------------------------------------------------------------ *

  colvo_zap = 0

  DO CASE
    CASE pr_zarpl = .T.  && Признак наличия зарплатного проекта ВКЛЮЧЕН

      text1 = 'Схема рассчета процентов полным списком'
      text2 = 'Схема рассчета процентов отдельно по проектам'
      l_bar = 3
      =popup_9(text1, text2, text3, text4, l_bar)
      ACTIVATE POPUP vibor
      RELEASE POPUPS vibor

      IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

        DO CASE
          CASE BAR() = 1  && Схема рассчета процентов полным списком

            =soob('Внимание! Вами выбрана схема расчета процентов - полным списком')

            SELECT DISTINCT branch, ref_client, ref_card, fio_rus, data_proz, card_nls, card_acct,;
              val_card, name_card, stavka, 000.00 AS sum_ostat, num_47411, pr_odb, num_dog, data_dog ;
              FROM account ;
              WHERE data_proz <= vvod_data_end AND stavka <> 0 AND del_card <> 2 AND ALLTRIM(vid_card) == '1' ;
              UNION ;
              SELECT DISTINCT branch, ref_client, ref_card, fio_rus, data_proz, card_nls, card_acct,;
              val_card, name_card, stavka, 000.00 AS sum_ostat, num_47411, pr_odb, num_dog, data_dog ;
              FROM acc_del ;
              WHERE data_proz <= vvod_data_end AND stavka <> 0 AND ALLTRIM(vid_card) == '1' AND (end_bal <> 0 OR isx_ost_m <> 0) ;
              INTO CURSOR sel_account ;
              ORDER BY 1, 4

* BROWSE WINDOW brows TITLE 'Схема рассчета процентов полным списком + pr_zarpl = T'

          CASE BAR() = 3  && Схема рассчета процентов отдельно по проектам

            =soob('Внимание! Вами выбрана схема расчета процентов - отдельно по проектам')

            =soob('Вами выбран код проекта - ' + ALLTRIM(num_proekt))

            SELECT DISTINCT branch, ref_client, ref_card, fio_rus, data_proz, card_nls, card_acct,;
              val_card, name_card, stavka, 000.00 AS sum_ostat, num_47411, pr_odb, num_dog, data_dog ;
              FROM account ;
              WHERE data_proz <= vvod_data_end AND stavka <> 0 AND del_card <> 2 AND ALLTRIM(vid_card) == '1' AND ALLTRIM(kod_proekt) == ALLTRIM(num_proekt) ;
              UNION ;
              SELECT DISTINCT branch, ref_client, ref_card, fio_rus, data_proz, card_nls, card_acct,;
              val_card, name_card, stavka, 000.00 AS sum_ostat, num_47411, pr_odb, num_dog, data_dog ;
              FROM acc_del ;
              WHERE data_proz <= vvod_data_end AND stavka <> 0 AND ALLTRIM(vid_card) == '1' AND ALLTRIM(kod_proekt) == ALLTRIM(num_proekt) AND ;
              (end_bal <> 0 OR isx_ost_m <> 0) ;
              INTO CURSOR sel_account ;
              ORDER BY 1, 4

* BROWSE WINDOW brows TITLE 'Схема рассчета процентов отдельно по проектам + pr_zarpl = T'

        ENDCASE

      ELSE

        =err('Внимание! Вами не выбрана схема расчета процентов ....')

        RETURN

      ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE pr_zarpl = .F.  && Признак наличия зарплатного проекта ВЫКЛЮЧЕН

      SELECT DISTINCT branch, ref_client, ref_card, fio_rus, data_proz, card_nls, card_acct,;
        val_card, name_card, stavka, 000.00 AS sum_ostat, num_47411, pr_odb, num_dog, data_dog ;
        FROM account ;
        WHERE data_proz <= vvod_data_end AND stavka <> 0 AND del_card <> 2 AND ALLTRIM(vid_card) == '1' ;
        UNION ;
        SELECT DISTINCT branch, ref_client, ref_card, fio_rus, data_proz, card_nls, card_acct,;
        val_card, name_card, stavka, 000.00 AS sum_ostat, num_47411, pr_odb, num_dog, data_dog ;
        FROM acc_del ;
        WHERE data_proz <= vvod_data_end AND stavka <> 0 AND ALLTRIM(vid_card) == '1' AND (end_bal <> 0 OR isx_ost_m <> 0) ;
        INTO CURSOR sel_account ;
        ORDER BY 1, 4

* BROWSE WINDOW brows TITLE 'Схема рассчета процентов полным списком + pr_zarpl = F'

  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  colvo_zap = _TALLY

  IF colvo_zap <> 0  && Если количество найденных записей в таблицах не равно нулю, то запускаем расчет процентов

    WAIT WINDOW 'Количество счетов по которым будет произведен расчет процентов равно - ' + ALLTRIM(STR(colvo_zap)) TIMEOUT 3

    put_bar = 1

    DO raschet IN Raschet_proz_42301

  ELSE
    =err('Внимание! В базе данных записей не обнаружено.')
  ENDIF

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
colvo_den_god = GOMONTH(CTOD('01.01.' + ALLTRIM(STR(YEAR(new_data_odb)))), 12) - CTOD('01.01.' + ALLTRIM(STR(YEAR(new_data_odb))))
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SELECT(sel)
RETURN


************************************************************** PROCEDURE use_acc_cln_null ***************************************************************

PROCEDURE use_acc_cln_null

DO CASE
  CASE pr_table_proz = .T.  && Временные процентные таблицы создаются на сервере

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Идет создание и открытие временных таблиц на СЕРВЕРЕ ... ',WCOLS())

    tim1 = SECONDS()

    IF FILE((put_dbf) + ('acc_cln_dt_null.dbf'))

      IF FILE((put_dbf) + ('acc_cln_dt_null.cdx'))

        IF NOT USED('acc_cln_dt_null')

          SELECT 0
          USE (put_dbf) + ('acc_cln_dt_null.dbf') ALIAS acc_cln_dt_null ORDER TAG card_nls

          new_ima_dt_null = ALLTRIM(('acc_cln_dt_null_') + LOWER(SYS(2015)))

          COPY STRUCTURE TO (put_tmp) + ALLTRIM(new_ima_dt_null) + ('.dbf') CDX
          USE
          USE (put_tmp) + ALLTRIM(new_ima_dt_null) + ('.dbf') ALIAS acc_cln_dt_null ORDER TAG card_nls

        ENDIF

      ELSE
        =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('acc_cln_dt_null.cdx') + ' не найден.')
        RETURN
      ENDIF

    ELSE
      =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('acc_cln_dt_null.dbf') + ' не найден.')
      RETURN
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    IF FILE((put_dbf) + ('acc_cln_sum_null.dbf'))

      IF FILE((put_dbf) + ('acc_cln_sum_null.cdx'))

        IF NOT USED('acc_cln_sum_null')

          SELECT 0
          USE (put_dbf) + ('acc_cln_sum_null.dbf') ALIAS acc_cln_sum_null ORDER TAG card_nls

          new_ima_sum_null = ALLTRIM(('acc_cln_sum_null_') + LOWER(SYS(2015)))

          COPY STRUCTURE TO (put_tmp) + ALLTRIM(new_ima_sum_null) + ('.dbf') CDX
          USE
          USE (put_tmp) + ALLTRIM(new_ima_sum_null) + ('.dbf') ALIAS acc_cln_sum_null ORDER TAG card_nls

        ENDIF

      ELSE
        =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('acc_cln_sum_null.cdx') + ' не найден.')
        RETURN
      ENDIF

    ELSE
      =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('acc_cln_sum_null.dbf') + ' не найден.')
      RETURN
    ENDIF

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Временные таблицы созданы на СЕРВЕРЕ.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

* ===================================================================================================================================== *

  CASE pr_table_proz = .F.  && Временные процентные таблицы создаются на локальной машине

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Идет создание и открытие временных таблиц на ЛОКАЛЬНОЙ МАШИНЕ ... ',WCOLS())

    tim1 = SECONDS()

    path_str = SYS(5) + SYS(2003)

    IF FILE((put_dbf) + ('acc_cln_dt_null.dbf'))

      IF FILE((put_dbf) + ('acc_cln_dt_null.cdx'))

        IF NOT USED('acc_cln_dt_null')

          SELECT 0
          USE (put_dbf) + ('acc_cln_dt_null.dbf') ALIAS acc_cln_dt_null ORDER TAG card_nls

          new_ima_dt_null = ALLTRIM(('acc_cln_dt_null_') + LOWER(SYS(2015)))

          COPY STRUCTURE TO (put_tmp_dbf) + ALLTRIM(new_ima_dt_null) + ('.dbf') CDX
          USE
          USE (put_tmp_dbf) + ALLTRIM(new_ima_dt_null) + ('.dbf') ALIAS acc_cln_dt_null ORDER TAG card_nls

        ENDIF

      ELSE
        =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('acc_cln_dt_null.cdx') + ' не найден.')
        RETURN
      ENDIF

    ELSE
      =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('acc_cln_dt_null.dbf') + ' не найден.')
      RETURN
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    IF FILE((put_dbf) + ('acc_cln_sum_null.dbf'))

      IF FILE((put_dbf) + ('acc_cln_sum_null.cdx'))

        IF NOT USED('acc_cln_sum_null')

          SELECT 0
          USE (put_dbf) + ('acc_cln_sum_null.dbf') ALIAS acc_cln_sum_null ORDER TAG card_nls

          new_ima_sum_null = ALLTRIM(('acc_cln_sum_null_') + LOWER(SYS(2015)))

          COPY STRUCTURE TO (put_tmp_dbf) + ALLTRIM(new_ima_sum_null) + ('.dbf') CDX
          USE
          USE (put_tmp_dbf) + ALLTRIM(new_ima_sum_null) + ('.dbf') ALIAS acc_cln_sum_null ORDER TAG card_nls

        ENDIF

      ELSE
        =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('acc_cln_sum_null.cdx') + ' не найден.')
        RETURN
      ENDIF

    ELSE
      =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('acc_cln_sum_null.dbf') + ' не найден.')
      RETURN
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    IF FILE((put_dbf) + ('istor_ost.dbf'))

      IF FILE((put_dbf) + ('istor_ost.cdx'))

        IF NOT USED('istor_ost')

          SELECT 0
          USE (put_dbf) + ('istor_ost.dbf') ALIAS istor_ost ORDER TAG card_nls

        ELSE
          SELECT istor_ost
        ENDIF

        new_ima_sum_null = 'istor_ost_local'

        COPY TO (put_tmp_dbf) + ALLTRIM(new_ima_sum_null) + ('.dbf') CDX
        USE
        USE (put_tmp_dbf) + ALLTRIM(new_ima_sum_null) + ('.dbf') ALIAS istor_ost ORDER TAG card_nls

      ELSE
        =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('istor_ost.cdx') + ' не найден.')
        RETURN
      ENDIF

    ELSE
      =err('Внимание! Нужный для работы файл - ' + (put_dbf) + ('istor_ost.dbf') + ' не найден.')
      RETURN
    ENDIF

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Временные таблицы созданы на ЛОКАЛЬНОЙ МАШИНЕ.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

ENDCASE

RETURN


************************************************************ PROCEDURE close_acc_cln_null ***************************************************************

PROCEDURE close_acc_cln_null

DO CASE
  CASE pr_table_proz = .T.  && Временные процентные таблицы создаются на сервере

    IF USED('acc_cln_dt_null')
      SELECT acc_cln_dt_null
      USE
      ERASE (put_tmp) + ('acc_cln_dt_null_') + ('*.*')
    ENDIF

    IF USED('acc_cln_sum_null')
      SELECT acc_cln_sum_null
      USE
      ERASE (put_tmp) + ('acc_cln_sum_null_') + ('*.*')
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE pr_table_proz = .F.  && Временные процентные таблицы создаются на локальной машине

    IF USED('acc_cln_dt_null')
      SELECT acc_cln_dt_null
      USE
      ERASE (put_tmp_dbf) + ('acc_cln_dt_null_') + ('*.*')
    ENDIF

    IF USED('acc_cln_sum_null')
      SELECT acc_cln_sum_null
      USE
      ERASE (put_tmp_dbf) + ('acc_cln_sum_null_') + ('*.*')
    ENDIF

    IF USED('istor_ost')
      SELECT istor_ost
      USE
      USE (put_dbf) + ('istor_ost.dbf') ALIAS istor_ost ORDER TAG card_nls
      ERASE (put_tmp_dbf) + ('istor_ost_local') + ('.*')
    ENDIF

ENDCASE

RETURN


************************************************************* PROCEDURE raschet ************************************************************************

PROCEDURE raschet

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO use_acc_cln_null IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PRIVATE vid_istor_ost
STORE SPACE(0) TO vid_istor_ost

STORE 00.0000 TO tim230, tim231, tim232, tim233, tim234, tim235, tim236, tim237

sum_ostat_prev = 0
ref_ost_prev = SYS(2015)

SELECT kurs_val
SET ORDER TO data_kurs
GO TOP

SELECT acc_cln_dt_null
SET ORDER TO card_nls
GO TOP

SELECT acc_cln_sum_null
SET ORDER TO card_nls
GO TOP

* ------------------------------------------------------------------------------------------ Выбор таблицы с историей остатков ------------------------------------------------------------------------------------------------ *

text1 = 'Вы желаете использовать историю остатков из ПО "TRANSMASTER"'
text2 = 'Вы желаете использовать историю остатков из ПО "CARD-VISA филиал"'
l_bar = 3
=popup_9(text1, text2, text3, text4, l_bar)
ACTIVATE POPUP vibor
RELEASE POPUPS vibor

IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

  DO CASE
    CASE BAR() = 1  && Вы желаете использовать историю остатков из ПО "TRANSMASTER"

      vid_istor_ost = 'istor_ost'

      SELECT istor_ost
      SET ORDER TO card_date  && INDEX ON ALLTRIM(card_nls) + DTOS(TTOD(date_ost_p)) TAG card_date

    CASE BAR() = 3  && Вы желаете использовать историю остатков из ПО "CARD-VISA филиал"

      vid_istor_ost = 'istor_mak'

      SELECT istor_mak
      SET ORDER TO card_date  && INDEX ON ALLTRIM(card_nls) + DTOS(TTOD(date_ost_p)) TAG card_date

  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Внимание! Начат расчет процентов по карточным счетам клиентов ...',WCOLS())

  SELECT sel_account  && Переходим в сформированную таблицу выбранных клиентов

  SCAN  && Начало сканирования таблицы выбранных клиентов

    sel_card_nls = ALLTRIM(sel_account.card_nls)  && Присваиваем переменной номер сканируемого карточного счета

* ------------------------------------------------------------------------- Берем один карточный счет и сканируем таблицу с датами --------------------------------------------------------------------------------- *
    DO scan_calendar IN Raschet_proz_42301
* ---------------------------------------------------------------------- Сворачивание данных рассчитанных по дням в итоги за месяц ------------------------------------------------------------------------------ *
    DO svert_sum_den IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT WINDOW 'Клиент - ' + ALLTRIM(sel_account.fio_rus) + '  Счет клиента - ' + ALLTRIM(sel_card_nls) + ' (время = ' + ALLTRIM(STR(tim237 - tim230, 6, 3)) + ' сек.)' NOWAIT

    SELECT sel_account

  ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  WAIT CLEAR

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Расчет процентов успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)

  @ WROWS()/3,3 SAY PADC('Внимание! Начат перенос рассчитанных процентов по датам в главную таблицу ... ',WCOLS())
  =INKEY(2)

  tim2 = SECONDS()

* -------------------------------------------------------------------- Переносим рассчитанные данные по дням в основные таблицы ----------------------------------------------------------------------------- *
  DO export_data_proz_dt IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  tim3 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Перенос рассчитанных процентов в главную таблицу успешно завершен' + ' (время = ' + ALLTRIM(STR(tim3 - tim2, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)

  @ WROWS()/3,3 SAY PADC('Внимание! Начат перенос рассчитанных итоговых процентов в главную таблицу ... ',WCOLS())
  =INKEY(2)

  tim2 = SECONDS()

* -------------------------------------------------------------------- Переносим рассчитанные данные за месяц в основные таблицы ---------------------------------------------------------------------------- *
  DO export_data_proz_sum IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  tim3 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Перенос рассчитанных процентов в главную таблицу успешно завершен' + ' (время = ' + ALLTRIM(STR(tim3 - tim2, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* ------------------------------------------------------------------------- Очистка рассчитанных процентов во временных таблицах -------------------------------------------------------------------------------- *
* DO export_del_proz IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  WAIT CLEAR

  tim3 = SECONDS()

  RELEASE tim230, tim231, tim232, tim233, tim234, tim235, tim236, tim237

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Процедура по расчету процентов успешно завершена.' + ' (время = ' + ALLTRIM(STR(tim3 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO close_acc_cln_null IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

RETURN


******************************************************** PROCEDURE scan_calendar ************************************************************************

PROCEDURE scan_calendar

tim230 = SECONDS()

SELECT calendar

* Сканирование таблицы calendar.dbf в диапазоне введенных дат: начала диапазона vvod_data_home и конец диапазона vvod_data_end
* Также проверяется условие: если карта заведена позже введенной даты начала диапазона vvod_data_home,
* то есть vvod_data_home меньше чем дата начала расчета процентов sel_account.data_proz, то за начало диапазона
* берется дата начала расчета процентов sel_account.data_proz

SCAN FOR BETWEEN(calendar.data_home, IIF(vvod_data_home < sel_account.data_proz, sel_account.data_proz, vvod_data_home), vvod_data_end)

*SCAN FOR IIF(vvod_data_home < sel_account.data_proz,;
BETWEEN(calendar.data_home, sel_account.data_proz, vvod_data_end),;
BETWEEN(calendar.data_home, vvod_data_home, vvod_data_end))

  tim231 = SECONDS()

  IF INLIST(SUBSTR(ALLTRIM(sel_card_nls), 6, 3), kod_val_usd, kod_val_eur)  && Работа с курсами только для валютных счетов

* ------------------------------------------------------------------------------------------------------------ Поиск курсов ---------------------------------------------------------------------------------------------------------------- *
    SET NEAR ON
* ------------------------------------------------------------------------------------------------ Ищем курс валюты в долларах -------------------------------------------------------------------------------------------------- *

    SELECT kurs_val
    GO TOP

    poisk_usd = SEEK(DTOS(calendar.data_end) + ALLTRIM('USD'))

    IF poisk_usd = .T.

      poisk_kurs_usd = kurs_val.kurs

    ELSE

      DO WHILE .T.
        IF ALLTRIM(kod_str) == 'USD'
          poisk_kurs_usd = kurs_val.kurs
          EXIT
        ELSE
          SKIP -1
          LOOP
        ENDIF
      ENDDO

    ENDIF

* ------------------------------------------------------------------------------------------------------- Ищем курс валюты в евро ------------------------------------------------------------------------------------------------- *

    SELECT kurs_val
    GO TOP

    poisk_eur = SEEK(DTOS(calendar.data_end) + ALLTRIM('EUR'))

    IF poisk_eur = .T.

      poisk_kurs_eur = kurs_val.kurs

    ELSE

      DO WHILE .T.
        IF ALLTRIM(kod_str) == 'EUR'
          poisk_kurs_eur = kurs_val.kurs
          EXIT
        ELSE
          SKIP -1
          LOOP
        ENDIF
      ENDDO

    ENDIF

    SET NEAR OFF

* ---------------------------------------------------------------------------------- Пишем в таблицу с датами расчета найденные курсы ------------------------------------------------------------------------------- *

    SELECT calendar

    SCATTER MEMVAR FIELDS kod_usd, kurs_usd, kod_eur, kurs_eur

    m.kod_usd = kod_val_usd
    m.kurs_usd = poisk_kurs_usd

    m.kod_eur = kod_val_eur
    m.kurs_eur = poisk_kurs_eur

    GATHER MEMVAR FIELDS kod_usd, kurs_usd, kod_eur, kurs_eur

  ENDIF  && Работа с курсами только для валютных счетов

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  rec = ALLTRIM(sel_card_nls) + DTOS(calendar.data_home) + DTOS(calendar.data_end)

  SELECT acc_cln_dt_null

  poisk = SEEK(rec)

  IF poisk = .T. AND acc_cln_dt_null.pr = .F.   && Данная запись отредактирована вручную, пропустить ее, не обсчитывать.

    LOOP

  ELSE

    IF poisk = .T.
      SCATTER MEMVAR MEMO
    ELSE
      SCATTER MEMVAR MEMO BLANK
      m.pr = .T.
    ENDIF

    m.branch = ALLTRIM(sel_account.branch)
    m.ref_client = ALLTRIM(sel_account.ref_client)
    m.ref_card = sel_account.ref_card

    m.card_nls = ALLTRIM(sel_account.card_nls)  && Номер карточного счета
    m.val_card = ALLTRIM(sel_account.val_card)  && Тип карты

    m.fio_rus = ALLTRIM(sel_account.fio_rus)  && ФИО клиента

    m.num_dog = sel_account.num_dog  && Номер договора
    m.data_dog = sel_account.data_dog  && Дата начала заключения договора

    m.data_proz = sel_account.data_proz  && Дата начала начисления процентов
    m.dt_hom_end = data_home_end  && Дата начала квартала за который происходит расчет
    m.data_home = calendar.data_home  && Начальная дата диапазона расчета, вводится оператором
    m.data_end = calendar.data_end  && Конечная дата диапазона расчета, вводится оператором
    m.data_pak = vvod_data_pak  && Конечная дата диапазона расчета при пакетном расчете процентов, рассчитыается автоматически

* --------------------------------------------------------------------------------- Пишем в таблицу с датами расчета найденные курсы -------------------------------------------------------------------------------- *

    DO CASE
      CASE SUBSTR(ALLTRIM(m.card_nls), 6, 3) == ALLTRIM(kod_val_rur)
        m.kurs = 1  && Курс валюты

      CASE SUBSTR(ALLTRIM(m.card_nls), 6, 3) == ALLTRIM(kod_val_usd)
        m.kurs = vvod_kurs_usd  && Курс валюты

      CASE SUBSTR(ALLTRIM(m.card_nls), 6, 3) == ALLTRIM(kod_val_eur)
        m.kurs = vvod_kurs_eur  && Курс валюты

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    m.kvartal = num_kvartal  && Номер квартала
    m.mes_kvart = num_mesaz_kvartal  && Номер месяца в квартале

    m.stavka = sel_account.stavka  && Процентная ставка

* -------------------------------------------------------------------------------- Получение входящего остатка за сканируемую дату ----------------------------------------------------------------------------------- *

    DO CASE
      CASE m.data_home = new_data_odb  && Если рассчетная дата равна дате открытого операционного дня

        DO vxd_ost_vxod_ost IN Raschet_proz_42301

      CASE m.data_home > new_data_odb AND BETWEEN((m.data_home - new_data_odb), 1, 3) && Если рассчетная дата больше даты открытого операционного дня

        DO vxd_ost_isxod_ost IN Raschet_proz_42301

      CASE m.data_home < new_data_odb  && Если рассчетная дата меньше даты операционного дня

        DO vxd_ost_istor_ost IN Raschet_proz_42301

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    IF m.sum_ostat <> sum_ostat_prev
      m.ref_ost = SYS(2015)
    ELSE
      m.ref_ost = ALLTRIM(ref_ost_prev)
    ENDIF

    sum_ostat_prev = m.sum_ostat  && Сумма остатка на предыдущую дату
    ref_ost_prev = m.ref_ost  && Метка остатка на предыдущую дату

    m.colvo_den = IIF(calendar.data_home < sel_account.data_proz, (calendar.data_end - sel_account.data_proz), calendar.colvo_den)

* -------------------------------------------------------------------------------------------- Параметр велечины порядка округления -------------------------------------------------------------------------------------- *
    por_okrugl = 5 && Порядок округления для формулы ROUND()
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    sum_dobav_1 = (m.sum_ostat / colvo_den_god) * m.colvo_den

    sum_dobav_2 = ROUND(((sum_dobav_1 * sel_account.stavka) / 100), por_okrugl)

    m.sum_dobav = IIF(m.pr = .T., sum_dobav_2, m.sum_dobav)  && Сумма расчетная

    sum_rasch_1 = (m.sum_ostat / colvo_den_god) * m.colvo_den

    sum_rasch_2 = ROUND(((sum_rasch_1 * m.stavka) / 100), por_okrugl)

    m.sum_rasch = IIF(m.pr = .T.,sum_rasch_2, m.sum_rasch)  && Сумма расчетная

    m.sum_rasxod = IIF(m.pr = .T., ROUND((m.sum_rasch * m.kurs), por_okrugl), m.sum_rasxod)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE INLIST(ALLTRIM(m.val_card), 'USD', 'EUR')  && Валютный вклад, необходимо провести расчет налога

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

        DO CASE
          CASE m.stavka > stavka_ref AND INLIST(LEFT(ALLTRIM(m.card_nls), 5), '40817')  && Для резидентов

* Если у клиента ставка больше 9.00 %, то с него нужно взять подоходный налог в размере 35%
* Сначало считаем сумму по его ставке = sum_nalog_1
* Потом считаем сумму по ставке 9.0 % = sum_nalog_2
* Потом считаем разницу (sum_nalog_1 - sum_nalog_2)
* И берем с нее подоходный налог в размере 35% (((sum_nalog_1 - sum_nalog_2) * 35) / 100) = sum_nalog_3

            sum_nalog_1 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * m.stavka) / 100

            sum_nalog_2 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * 9) / 100

            sum_nalog_3 = ROUND((((sum_nalog_1 - sum_nalog_2) * 35) / 100), por_okrugl)

            m.sum_nalog = IIF(m.pr = .T., sum_nalog_3, m.sum_nalog)  && Сумма налога

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

          CASE m.stavka > 9.00 AND INLIST(LEFT(ALLTRIM(m.card_nls), 5), '40820')  && Для резидентов

* Если у клиента ставка больше 9.00 %, то с него нужно взять подоходный налог в размере 35%
* Сначало считаем сумму по его ставке = sum_nalog_1
* Потом считаем сумму по ставке 9.0 % = sum_nalog_2
* Потом считаем разницу (sum_nalog_1 - sum_nalog_2)
* И берем с нее подоходный налог в размере 35% (((sum_nalog_1 - sum_nalog_2) * 35) / 100) = sum_nalog_3

            sum_nalog_1 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * m.stavka) / 100

            sum_nalog_2 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * 9) / 100

            sum_nalog_3 = ROUND((((sum_nalog_1 - sum_nalog_2) * 35) / 100), por_okrugl)

            m.sum_nalog = IIF(m.pr = .T., sum_nalog_3, m.sum_nalog)  && Сумма налога

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

*            CASE INLIST(LEFT(ALLTRIM(m.card_nls), 5), '40820')  && Для не резидентов    Эта секция для старой редакции налового кодекса

*  * Если клиент не резидент , то с него нужно взять подоходный налог в размере 30%
*  * Сначало считаем сумму по его ставке = sum_nalog_1
*  * Потом берем с нее подоходный налог в размере 30% ((sum_nalog_1 * 30) / 100) = sum_nalog_2

*              sum_nalog_1 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * m.stavka) / 100

*              sum_nalog_2 = ROUND(((sum_nalog_1 * 30) / 100), por_okrugl)

*              m.sum_nalog = IIF(m.pr = .T., sum_nalog_2, m.sum_nalog)  && Сумма налога

          OTHERWISE

            m.sum_nalog = 0

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE INLIST(ALLTRIM(m.val_card), 'RUR')  && Рублевой вклад вклад, расчет налога не производится

        m.sum_nalog = 0

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    m.sum_viplat = m.sum_rasch - m.sum_nalog  && Сумма к выплате

    m.data = DATETIME()
    m.ima_pk = SYS(0)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    SELECT acc_cln_dt_null

    IF poisk = .T.
      GATHER MEMVAR MEMO
    ELSE
      INSERT INTO acc_cln_dt_null FROM MEMVAR
    ENDIF

  ENDIF

  tim232 = SECONDS()

  IF pr_time_rachet_proz = .T.
    SCATTER MEMVAR FIELDS time_den
    m.time_den = ROUND(tim232 - tim231, 3)
    GATHER MEMVAR FIELDS time_den
  ENDIF

  SELECT calendar
ENDSCAN

tim233 = SECONDS()

RETURN


************************************************************** PROCEDURE vxd_ost_istor_ost ***************************************************************

PROCEDURE vxd_ost_istor_ost

* ----------------------------------------------------------------------------- Использовать историю остатков из ПО "TRANSMASTER" ----------------------------------------------------------------------------- *

DO CASE
  CASE ALLTRIM(vid_istor_ost) == 'istor_ost'

    SELECT istor_ost

    SET NEAR ON

    poisk_ostat = SEEK(ALLTRIM(m.card_nls) + DTOS(m.data_home))  && Ищем остаток по счету на заданную дату

    SET NEAR OFF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

        SELECT acc_cln_dt_null

        m.sum_ostat = istor_ost.begin_bal  && Сумма остатка на дату

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE poisk_ostat = .F.  AND m.data_home < TTOD(istor_ost.date_ost_p) AND ALLTRIM(istor_ost.card_nls) == ALLTRIM(m.card_nls)

        SELECT istor_ost
        SKIP -1

        IF ALLTRIM(istor_ost.card_nls) == ALLTRIM(m.card_nls) AND m.data_home > TTOD(istor_ost.date_ost_p)

          SELECT acc_cln_dt_null

          m.sum_ostat = istor_ost.end_bal  && Сумма остатка на дату

        ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE poisk_ostat = .F. AND ALLTRIM(istor_ost.card_nls) <> ALLTRIM(m.card_nls)

        SELECT istor_ost
        SKIP -1

        IF ALLTRIM(istor_ost.card_nls) == ALLTRIM(m.card_nls) AND m.data_home > TTOD(istor_ost.date_ost_p)

          m.sum_ostat = istor_ost.end_bal  && Сумма остатка на дату

        ELSE

          SELECT MAX(date_ost_p) AS date_ost_p, branch, ref_client, card_nls, card_acct, begin_bal, debit, credit, end_bal ;
            FROM istor_ost ;
            WHERE ALLTRIM(card_nls) == ALLTRIM(m.card_nls) AND date_ost_p < m.data_home ;
            INTO CURSOR  sel_ostatok

          SELECT acc_cln_dt_null

          IF _TALLY <> 0

            m.sum_ostat = sel_ostatok.end_bal  && Сумма остатка на дату

          ELSE

            SELECT account
            SET ORDER TO card_nls

            poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && Ищем остаток по счету на заданную дату

            DO CASE
              CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

                SELECT acc_cln_dt_null

                m.sum_ostat = account.begin_bal  && Сумма остатка на дату

              CASE poisk_ostat = .F.  && Поиск не успешный, по счету на рассчитанную дату

                SELECT acc_del
                SET ORDER TO card_nls

                poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && Ищем остаток по счету на заданную дату

                DO CASE
                  CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

                    SELECT acc_cln_dt_null

                    m.sum_ostat = acc_del.begin_bal  && Сумма остатка на дату

                  CASE poisk_ostat = .F.  && Поиск не успешный, по счету на рассчитанную дату

                    m.sum_ostat = 0.00  && Сумма остатка на дату

                ENDCASE

            ENDCASE

          ENDIF

        ENDIF

    ENDCASE

* --------------------------------------------------------------------------- Использовать историю остатков из ПО "CARD-VISA филиал" --------------------------------------------------------------------------- *

  CASE ALLTRIM(vid_istor_ost) == 'istor_mak'

    SELECT istor_mak

    SET NEAR ON

    poisk_ostat = SEEK(ALLTRIM(m.card_nls) + DTOS(m.data_home))  && Ищем остаток по счету на заданную дату

    SET NEAR OFF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

        SELECT acc_cln_dt_null

        m.sum_ostat = istor_mak.vxd_ost_m  && Сумма остатка на дату

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE poisk_ostat = .F.  AND m.data_home < TTOD(istor_mak.date_ost_m) AND ALLTRIM(istor_mak.card_nls) == ALLTRIM(m.card_nls)

        SELECT istor_mak
        SKIP -1

        IF ALLTRIM(istor_mak.card_nls)  == ALLTRIM(m.card_nls) AND m.data_home > TTOD(istor_mak.date_ost_m)

          SELECT acc_cln_dt_null

          m.sum_ostat = istor_mak.isx_ost_m  && Сумма остатка на дату

        ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE poisk_ostat = .F. AND ALLTRIM(card_nls) <> ALLTRIM(m.card_nls)

        SELECT istor_mak
        SKIP -1

        IF ALLTRIM(istor_mak.card_nls) == ALLTRIM(m.card_nls) AND m.data_home > TTOD(istor_mak.date_ost_m)

          SELECT acc_cln_dt_null

          m.sum_ostat = istor_mak.isx_ost_m  && Сумма остатка на дату

        ELSE

          SELECT MAX(date_ost_m) AS date_ost_m, branch, ref_client, card_nls, card_acct, vxd_ost_m, debit_m, credit_m, isx_ost_m ;
            FROM istor_mak ;
            WHERE ALLTRIM(card_nls) == ALLTRIM(m.card_nls) AND TTOD(date_ost_m) < m.data_home ;
            INTO CURSOR  sel_ostatok

          SELECT acc_cln_dt_null

          IF _TALLY <> 0

            m.sum_ostat = sel_ostatok.isx_ost_m  && Сумма остатка на дату

          ELSE

            SELECT account
            SET ORDER TO card_nls

            poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && Ищем остаток по счету на заданную дату

            DO CASE
              CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

                SELECT acc_cln_dt_null

                m.sum_ostat = account.vxd_ost_m  && Сумма остатка на дату

              CASE poisk_ostat = .F.  && Поиск не успешный, по счету на рассчитанную дату

                SELECT acc_del
                SET ORDER TO card_nls

                poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && Ищем остаток по счету на заданную дату

                DO CASE
                  CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

                    SELECT acc_cln_dt_null

                    m.sum_ostat = acc_del.vxd_ost_m  && Сумма остатка на дату

                  CASE poisk_ostat = .F.  && Поиск не успешный, по счету на рассчитанную дату

                    m.sum_ostat = 0.00  && Сумма остатка на дату

                ENDCASE
            ENDCASE


          ENDIF

        ENDIF

    ENDCASE

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

RETURN


******************************************************* PROCEDURE vxd_ost_vxod_ost *********************************************************************

PROCEDURE vxd_ost_vxod_ost  && Если рассчетная дата равна дате открытого операционного дня

* ----------------------------------------------------------------------------- Использовать историю остатков из ПО "TRANSMASTER" ----------------------------------------------------------------------------- *

DO CASE
  CASE ALLTRIM(vid_istor_ost) == 'istor_ost'

    SELECT account
    SET ORDER TO card_nls

    poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && Ищем остаток по счету на заданную дату

    DO CASE
      CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

        SELECT acc_cln_dt_null

        m.sum_ostat = account.begin_bal  && Сумма остатка на дату

      CASE poisk_ostat = .F.  && Поиск не успешный, по счету на рассчитанную дату

        SELECT acc_del
        SET ORDER TO card_nls

        poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && Ищем остаток по счету на заданную дату

        DO CASE
          CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

            SELECT acc_cln_dt_null

            m.sum_ostat = acc_del.begin_bal  && Сумма остатка на дату

          CASE poisk_ostat = .F.  && Поиск не успешный, по счету на рассчитанную дату

            m.sum_ostat = 0.00  && Сумма остатка на дату

        ENDCASE
    ENDCASE

* --------------------------------------------------------------------------- Использовать историю остатков из ПО "CARD-VISA филиал" --------------------------------------------------------------------------- *

  CASE ALLTRIM(vid_istor_ost) == 'istor_mak'

    SELECT account
    SET ORDER TO card_nls

    poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && Ищем остаток по счету на заданную дату

    DO CASE
      CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

        SELECT acc_cln_dt_null

        m.sum_ostat = account.vxd_ost_m  && Сумма остатка на дату

      CASE poisk_ostat = .F.  && Поиск не успешный, по счету на рассчитанную дату

        SELECT acc_del
        SET ORDER TO card_nls

        poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && Ищем остаток по счету на заданную дату

        DO CASE
          CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

            SELECT acc_cln_dt_null

            m.sum_ostat = acc_del.vxd_ost_m  && Сумма остатка на дату

          CASE poisk_ostat = .F.  && Поиск не успешный, по счету на рассчитанную дату

            m.sum_ostat = 0.00  && Сумма остатка на дату

        ENDCASE
    ENDCASE

ENDCASE

RETURN


******************************************************* PROCEDURE vxd_ost_isxod_ost *********************************************************************

PROCEDURE vxd_ost_isxod_ost  && Если рассчетная дата больше даты открытого операционного дня

* ----------------------------------------------------------------------------- Использовать историю остатков из ПО "TRANSMASTER" ----------------------------------------------------------------------------- *

DO CASE
  CASE ALLTRIM(vid_istor_ost) == 'istor_ost'

    SELECT account
    SET ORDER TO card_nls

    poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && Ищем остаток по счету на заданную дату

    DO CASE
      CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

        SELECT acc_cln_dt_null

        m.sum_ostat = account.end_bal  && Сумма остатка на дату

      CASE poisk_ostat = .F.  && Поиск не успешный, по счету на рассчитанную дату

        SELECT acc_del
        SET ORDER TO card_nls

        poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && Ищем остаток по счету на заданную дату

        DO CASE
          CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

            SELECT acc_cln_dt_null

            m.sum_ostat = acc_del.end_bal  && Сумма остатка на дату

          CASE poisk_ostat = .F.  && Поиск не успешный, по счету на рассчитанную дату

            m.sum_ostat = 0.00  && Сумма остатка на дату

        ENDCASE
    ENDCASE

* --------------------------------------------------------------------------- Использовать историю остатков из ПО "CARD-VISA филиал" --------------------------------------------------------------------------- *

  CASE ALLTRIM(vid_istor_ost) == 'istor_mak'

    SELECT account
    SET ORDER TO card_nls

    poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && Ищем остаток по счету на заданную дату

    DO CASE
      CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

        SELECT acc_cln_dt_null

        m.sum_ostat = account.isx_ost_m  && Сумма остатка на дату

      CASE poisk_ostat = .F.  && Поиск не успешный, по счету на рассчитанную дату

        SELECT acc_del
        SET ORDER TO card_nls

        poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && Ищем остаток по счету на заданную дату

        DO CASE
          CASE poisk_ostat = .T.  && Поиск успешный, по счету на рассчитанную дату

            SELECT acc_cln_dt_null

            m.sum_ostat = acc_del.isx_ost_m  && Сумма остатка на дату

          CASE poisk_ostat = .F.  && Поиск не успешный, по счету на рассчитанную дату

            m.sum_ostat = 0.00  && Сумма остатка на дату

        ENDCASE
    ENDCASE

ENDCASE

RETURN


****************************************************** PROCEDURE svert_sum_den ***********************************************************************

PROCEDURE svert_sum_den

tim234 = SECONDS()

* ---------------------------------------------------------------------------------------- Параметр велечины порядка округления ------------------------------------------------------------------------------------------- *

por_okrugl_det = 4  && Порядок округления для формулы ROUND()

* ---------------------------------------------------------- Производим выборку данных из таблицы процентов за день и сворачиваем в иоговую -------------------------------------------------------- *

SELECT branch, ref_client, ref_card, ref_ost, card_nls, val_card, fio_rus, data_proz, dt_hom_end,;
  MIN(data_home) AS data_home,;
  MAX(data_end) - 1 AS data_end, data_pak,;
  kurs, kvartal, mes_kvart, stavka, sum_ostat,;
  SUM(colvo_den) AS den_mes,;
  ROUND(SUM(sum_dobav), por_okrugl_det) AS sum_dobav,;
  ROUND(SUM(sum_rasch), por_okrugl_det) AS sum_rasch,;
  ROUND(SUM(sum_rasxod), por_okrugl_det) AS sum_rasxod,;
  ROUND(SUM(sum_nalog), por_okrugl_det) AS sum_nalog,;
  ROUND(SUM(sum_viplat), por_okrugl_det) AS sum_viplat,;
  num_dog, data_dog, data, ima_pk ;
  FROM acc_cln_dt_null ;
  WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_nls) AND BETWEEN(data_home, vvod_data_home, vvod_data_end) ;
  INTO CURSOR sel_proz_sum ;
  GROUP BY card_nls, sum_ostat, ref_ost ;
  ORDER BY card_nls, data_home, data_end

* BROWSE WINDOW brows

* INTO CURSOR sel_proz_sum ;
* INTO TABLE (put_dbf) + ('sel_proz_sum.dbf') ;

* -------------------------------------------------------------------- Начинаем сканировать таблицу процентов сверную в итоговую ------------------------------------------------------------------------------ *

SELECT sel_proz_sum
GO TOP

SCAN

  tim235 = SECONDS()

  rec = ALLTRIM(sel_proz_sum.card_nls) + DTOS(sel_proz_sum.data_home) + DTOS(sel_proz_sum.data_end)

  SELECT acc_cln_sum_null

  poisk = SEEK(rec)

  IF poisk = .T. AND acc_cln_sum_null.pr = .F.   && Данная запись отредактирована вручную, пропустить ее, не обсчитывать.

    LOOP

  ELSE

    IF poisk = .T.
      SCATTER MEMVAR MEMO
    ELSE
      SCATTER MEMVAR MEMO BLANK
      m.pr = .T.
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    m.branch = ALLTRIM(sel_proz_sum.branch)

    m.ref_client = ALLTRIM(sel_proz_sum.ref_client)
    m.ref_card = sel_proz_sum.ref_card
    m.ref_ost = sel_proz_sum.ref_ost

    m.card_nls = ALLTRIM(sel_proz_sum.card_nls)  && Номер карточного счета
    m.val_card = ALLTRIM(sel_proz_sum.val_card)  && Тип карты

    m.fio_rus = ALLTRIM(sel_proz_sum.fio_rus)  && ФИО клиента

    m.num_dog = sel_proz_sum.num_dog  && Номер договора
    m.data_dog = sel_proz_sum.data_dog  && Дата начала заключения договора

    m.data_proz = sel_proz_sum.data_proz  && Дата начала начисления процентов

    m.dt_hom_end = sel_proz_sum.dt_hom_end  && Дата начала квартала за который происходит расчет
    m.data_home = sel_proz_sum.data_home  && Начальная дата диапазона расчета, вводится оператором
    m.data_end = sel_proz_sum.data_end  && Конечная дата диапазона расчета, вводится оператором
    m.data_pak = sel_proz_sum.data_pak  && Конечная дата диапазона расчета при пакетном расчете процентов, рассчитыается автоматически

    DO CASE
      CASE SUBSTR(ALLTRIM(m.card_nls),6,3) == ALLTRIM(kod_val_rur)
        m.kurs = 1  && Курс валюты

      CASE SUBSTR(ALLTRIM(m.card_nls),6,3) == ALLTRIM(kod_val_usd)
        m.kurs = vvod_kurs_usd  && Курс валюты

      CASE SUBSTR(ALLTRIM(m.card_nls),6,3) == ALLTRIM(kod_val_eur)
        m.kurs = vvod_kurs_eur  && Курс валюты

    ENDCASE

    m.kvartal = sel_proz_sum.kvartal  && Номер квартала
    m.mes_kvart = sel_proz_sum.mes_kvart  && Номер месяца в квартале

    m.stavka = sel_proz_sum.stavka  && Процентная ставка
    m.sum_ostat = sel_proz_sum.sum_ostat  && Сумма остатка по счету

    m.den_mes = sel_proz_sum.den_mes

* ---------------------------------------------------------------------------------------------- Параметр велечины порядка округления ------------------------------------------------------------------------------------- *
    por_okrugl_sum = 2  && Порядок округления для формулы ROUND()
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    m.sum_dobav = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_dobav, por_okrugl_sum), m.sum_dobav)
    m.sum_rasch = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_rasch, por_okrugl_sum), m.sum_rasch)
    m.sum_rasxod = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_rasxod, por_okrugl_sum), m.sum_rasxod)
    m.sum_nalog = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_nalog, por_okrugl_sum), m.sum_nalog)
    m.sum_viplat = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_viplat, por_okrugl_sum), m.sum_viplat)

* ----------------------------------------------------------------- Проверяем данные свернутые по дням в итоги за месяц на ошибку округления ----------------------------------------------------------- *
    por_okrugl_prover = 2  && Порядок округления для формулы ROUND()


    prover_sum_dobav = ROUND(((((m.sum_ostat / colvo_den_god) * m.den_mes) * m.stavka) / 100), por_okrugl_prover)  && Сумма расчетная для проверки

    m.sum_dobav = IIF(m.sum_dobav = prover_sum_dobav, m.sum_dobav, prover_sum_dobav)

    prover_sum_rasch = ROUND(((((m.sum_ostat / colvo_den_god) * m.den_mes) * m.stavka) / 100), por_okrugl_prover)  && Сумма расчетная для проверки

    m.sum_rasch = IIF(m.sum_rasch = prover_sum_rasch, m.sum_rasch, prover_sum_rasch)

    prover_sum_rasxod = ROUND((prover_sum_rasch * m.kurs), por_okrugl_prover)

* m.sum_rasxod = IIF(m.sum_rasxod = prover_sum_rasxod, m.sum_rasxod, prover_sum_rasxod)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE INLIST(ALLTRIM(m.val_card), 'USD', 'EUR')  && Валютный вклад, необходимо провести расчет налога

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

        DO CASE
          CASE m.stavka > 9 AND LEFT(ALLTRIM(m.card_nls), 5) == '40817'  && Для резидентов

* Если у клиента ставка больше 9.0 %, то с него нужно взять подоходный налог в размере 35%
* Сначало считаем сумму по его ставке = sum_nalog_1
* Потом считаем сумму по ставке 9.0 % = sum_nalog_2
* Потом считаем разницу (sum_nalog_1 - sum_nalog_2)
* И берем с нее подоходный налог в размере 35% (((sum_nalog_1 - sum_nalog_2) * 35) / 100) = sum_nalog_3

            sum_nalog_1 = (((m.sum_ostat / colvo_den_god) * m.den_mes) * m.stavka) / 100

            sum_nalog_2 = (((m.sum_ostat / colvo_den_god) * m.den_mes) * 9) / 100

            prover_sum_nalog = ROUND((((sum_nalog_1 - sum_nalog_2) * 35) / 100), por_okrugl_prover)  && Сумма налога

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

          CASE m.stavka > 9.00 AND INLIST(LEFT(ALLTRIM(m.card_nls), 5), '40820')  && Для НЕрезидентов

* Если у клиента ставка больше 9.00 %, то с него нужно взять подоходный налог в размере 35%
* Сначало считаем сумму по его ставке = sum_nalog_1
* Потом считаем сумму по ставке 9.0 % = sum_nalog_2
* Потом считаем разницу (sum_nalog_1 - sum_nalog_2)
* И берем с нее подоходный налог в размере 35% (((sum_nalog_1 - sum_nalog_2) * 35) / 100) = sum_nalog_3

            sum_nalog_1 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * m.stavka) / 100

            sum_nalog_2 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * 9) / 100

            sum_nalog_3 = ROUND((((sum_nalog_1 - sum_nalog_2) * 35) / 100), por_okrugl)

            m.sum_nalog = IIF(m.pr = .T., sum_nalog_3, m.sum_nalog)  && Сумма налога


          OTHERWISE

            m.sum_nalog = 0

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE INLIST(ALLTRIM(m.val_card), 'RUR')  && Рублевой вклад вклад, расчет налога не производится

        m.sum_nalog = 0

    ENDCASE

    m.sum_viplat = m.sum_rasch - m.sum_nalog  && Сумма к выплате

    m.data = sel_proz_sum.data
    m.ima_pk = sel_proz_sum.ima_pk

    SELECT acc_cln_sum_null

    IF poisk = .T.
      GATHER MEMVAR MEMO
    ELSE
      INSERT INTO acc_cln_sum_null FROM MEMVAR
    ENDIF

  ENDIF

  tim236 = SECONDS()

  IF pr_time_rachet_proz = .T.
    SCATTER MEMVAR FIELDS time_svert
    m.time_svert = ROUND(tim236 - tim235, 3)
    GATHER MEMVAR FIELDS time_svert
  ENDIF

  SELECT sel_proz_sum
ENDSCAN

tim237 = SECONDS()

RETURN


************************************************************** PROCEDURE sel_data **********************************************************************

PROCEDURE sel_data
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

vvod_data_pak = MaxDataMes(DATE())

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

DO WHILE .T.

  text1 = 'Выборка данных выборочно по одному клиенту и диапазону дат расчета'
  text2 = 'Выборка данных списком по всем клиентам и диапазону дат расчета'
  l_bar = 3
  =popup_9(text1, text2, text3, text4, l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    put_vibor = BAR()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE put_vibor = 1  && Выборка данных выборочно по одному клиенту и диапазону дат расчета

        num_mesaz = MONTH(new_data_odb)
        num_god = ALLTRIM(STR(YEAR(new_data_odb)))

        DO vid_kvartal_1 IN Raschet_proz_42301

        DO WHILE .T.

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
          DO FORM (put_scr) + ('poisk_client.scx')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

          SELECT DISTINCT SPACE(1) + fio_rus + ' - ' + val_card + ' - ' + card_nls AS fio_rus ;
            FROM account ;
            WHERE stavka <> 0 AND LEFT(ALLTRIM(fio_rus), len_client) == ALLTRIM(fam_client) ;
            UNION ;
            SELECT DISTINCT SPACE(1) + fio_rus + ' - ' + val_card + ' - ' + card_nls AS fio_rus ;
            FROM acc_del ;
            WHERE stavka <> 0 AND (end_bal <> 0 OR isx_ost_m <> 0) AND LEFT(ALLTRIM(fio_rus), len_client) == ALLTRIM(fam_client) ;
            INTO CURSOR tov ;
            ORDER BY 1

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

          IF _TALLY <> 0

            DO FORM (put_scr) + ('selfio_nls.scx')

            IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

              vvod_data_end = new_data_odb

              DO FORM (put_scr) + ('sel_data_proz.scx')

              IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

                num_mesaz = MONTH(vvod_data_end)
                num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

                DO vid_kvartal_2 IN Raschet_proz_42301

*          WAIT WINDOW DTOC(vvod_data_home) + '   ' + DTOC(vvod_data_end)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                DO WHILE .T.

                  text1 = 'Выборка данных с детализацией рассчетных периодов'
                  text2 = 'Выборка данных с суммированием рассчетных периодов'
                  l_bar = 3
                  =popup_9(text1, text2, text3, text4, l_bar)
                  ACTIVATE POPUP vibor
                  RELEASE POPUPS vibor

                  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

                    put_prn = BAR()

                    SELECT DISTINCT fio_rus, card_nls, card_acct, tran_42301, vklad_acct, val_card, name_card, pr_odb, kod_proekt ;
                      FROM account ;
                      WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_nls) ;
                      UNION ;
                      SELECT DISTINCT fio_rus, card_nls, card_acct, tran_42301, vklad_acct, val_card, name_card, pr_odb, kod_proekt ;
                      FROM acc_del ;
                      WHERE (end_bal <> 0 OR isx_ost_m <> 0) AND ALLTRIM(card_nls) == ALLTRIM(sel_card_nls) ;
                      INTO CURSOR sel_account ;
                      ORDER BY 1, 2

* BROWSE WINDOW brows

                    DO CASE
                      CASE put_prn = 1  && Выборка данных с детализацией рассчетных периодов

                        DO sel_data_det_cln IN Raschet_proz_42301

                      CASE put_prn = 3  && Выборка данных с суммированием рассчетных периодов

                        DO sel_data_sum_cln IN Raschet_proz_42301

                    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                    IF _TALLY <> 0

                      SELECT sel_prozent

                      DO CASE
                        CASE bar_num = 11  && Выбран пункт меню для просмотра данных на экране

                          DO brows_proz IN Raschet_proz_42301

                        CASE bar_num = 13  && Выбран пункт меню для просмотра данных на принтере

                          DO prn_proz IN Raschet_proz_42301

                      ENDCASE

                    ELSE
                      =err('Внимание! За выбранные даты данных не обнаружено ...')
                    ENDIF

                  ELSE
                    EXIT
                  ENDIF

                ENDDO

              ENDIF

            ELSE
              EXIT
            ENDIF

          ELSE
            =err('Внимание! По введенному Вами сочетанию букв фамилий не найдено ... ')
          ENDIF

        ENDDO

* ====================================================================================================================================== *

      CASE put_vibor = 3  && Выборка данных списком по всем клиентам и диапазону дат расчета

        num_mesaz = MONTH(new_data_odb)
        num_god = ALLTRIM(STR(YEAR(new_data_odb)))

        DO vid_kvartal_1 IN Raschet_proz_42308

* vvod_data_home = CTOD('01.05.2010')
* vvod_data_end = CTOD('31.05.2010')

        DO FORM (put_scr) + ('sel_data_proz.scx')

        IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

          num_mesaz = MONTH(vvod_data_end)
          num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

          DO vid_kvartal_2 IN Raschet_proz_42308

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

          DO WHILE .T.

            DO CASE
              CASE bar_num = 11  && Выбран пункт меню для просмотра данных на экране

                text1 = 'Отчет с детализацией рассчетных периодов'
                text2 = 'Отчет с суммированием рассчетных периодов'
                text3 = 'Отчет с сумированием по транзитным счетам'
                l_bar = 5
                =popup_9(text1, text2, text3, text4, l_bar)

              CASE bar_num = 13  && Выбран пункт меню для просмотра данных на принтере

                text1 = 'Отчет с детализацией рассчетных периодов - книжная'
                text2 = 'Отчет с детализацией рассчетных периодов - альбомная'
                text3 = 'Отчет с суммированием периодов коротких счетов нет - альбомная'
                text4 = 'Отчет с суммированием периодов короткие счета есть - альбомная'
                text5 = 'Отчет с сумированием по транзитным счетам - альбомная'
                l_bar = 9
                =popup_big(text1, text2, text3, text4, text5, text6, text7, text8, l_bar)

            ENDCASE

            ACTIVATE POPUP vibor
            RELEASE POPUPS vibor

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

            IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

              put_prn = BAR()

              SELECT DISTINCT fio_rus, card_nls, tran_42301, vklad_acct, val_card, name_card, pr_odb, kod_proekt ;
                FROM account ;
                WHERE stavka <> 0 ;
                UNION ;
                SELECT DISTINCT fio_rus, card_nls, tran_42301, vklad_acct, val_card, name_card, pr_odb, kod_proekt ;
                FROM acc_del ;
                WHERE stavka <> 0 AND (end_bal <> 0 OR isx_ost_m <> 0) ;
                INTO CURSOR sel_account ;
                ORDER BY 1, 2

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

              DO CASE
                CASE bar_num = 11  && Выбран пункт меню для просмотра данных на экране

                  DO CASE
                    CASE put_prn = 1  && Отчет с детализацией рассчетных периодов - форма № 1

                      DO sel_data_det_all IN Raschet_proz_42301

                    CASE put_prn = 3  && Отчет с суммированием рассчетных периодов - форма № 2

                      DO sel_data_sum_all IN Raschet_proz_42301

                    CASE put_prn = 5  && Выборка данных с суммированием по транзитным счетам - форма № 3

                      DO sel_data_sum_tran IN Raschet_proz_42301

                  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                CASE bar_num = 13  && Выбран пункт меню для просмотра данных на принтере

                  DO CASE
                    CASE put_prn = 1  && Отчет с детализацией рассчетных периодов - книжная - форма № 1

                      DO sel_data_det_all IN Raschet_proz_42301

                    CASE put_prn = 3  && Отчет с детализацией периодов коротких счетов нет - альбомная - форма № 2

                      DO sel_data_det_all IN Raschet_proz_42301

                    CASE put_prn = 5  && Отчет с суммированием периодов коротких счетов нет - альбомная - форма № 3

                      DO sel_data_sum_all IN Raschet_proz_42301

                    CASE put_prn = 7  && Отчет с суммированием периодов короткие счета есть - альбомная - форма № 4

                      DO sel_data_sum_all IN Raschet_proz_42301

                    CASE put_prn = 9  && Отчет с сумированием по транзитным счетам - альбомная - форма № 5

                      DO sel_data_sum_tran IN Raschet_proz_42301

                  ENDCASE

              ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

              IF _TALLY <> 0

                SELECT sel_prozent
                GO TOP

                DO CASE
                  CASE bar_num = 11  && Выбран пункт меню для просмотра данных на экране
                    DO brows_proz IN Raschet_proz_42301

                  CASE bar_num = 13  && Выбран пункт меню для печати данных на принтере
                    DO prn_proz IN Raschet_proz_42301

                ENDCASE

              ELSE
                =err('Внимание! За выбранные даты данных не обнаружено ...')
              ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

            ELSE
              EXIT
            ENDIF

          ENDDO

        ENDIF

    ENDCASE

  ELSE
    EXIT
  ENDIF

ENDDO

SELECT(sel)
RETURN


************************************************************* PROCEDURE sel_data_det_cln ****************************************************************

PROCEDURE sel_data_det_cln

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO CASE
  CASE pr_zarpl = .T.  && Признак наличия зарплатного проекта ВКЛЮЧЕН

    DO CASE
      CASE pr_summa_null = .T.  && Если признак=.T., то нулевые суммы выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          B.data_home, B.data_end, B.den_mes,;
          B.sum_dobav, B.sum_rasch, B.sum_rasxod, B.sum_nalog, B.sum_viplat,;
          A.tran_42301, A.vklad_acct, B.num_dog, B.data_dog,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ALLTRIM(B.card_nls) == ALLTRIM(sel_card_nls) AND ;
          ALLTRIM(A.kod_proekt) == ALLTRIM(num_proekt) ;
          INTO CURSOR sel_prozent ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

      CASE pr_summa_null = .F.  && Если признак=.F., то нулевые суммы не выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          B.data_home, B.data_end, B.den_mes,;
          B.sum_dobav, B.sum_rasch, B.sum_rasxod, B.sum_nalog, B.sum_viplat,;
          A.tran_42301, A.vklad_acct, B.num_dog, B.data_dog,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ALLTRIM(B.card_nls) == ALLTRIM(sel_card_nls) AND ;
          ALLTRIM(A.kod_proekt) == ALLTRIM(num_proekt) AND B.sum_dobav > 0.00 ;
          INTO CURSOR sel_prozent ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE pr_zarpl = .F.  && Признак наличия зарплатного проекта ВЫКЛЮЧЕН

    DO CASE
      CASE pr_summa_null = .T.  && Если признак=.T., то нулевые суммы выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          B.data_home, B.data_end, B.den_mes,;
          B.sum_dobav, B.sum_rasch, B.sum_rasxod, B.sum_nalog, B.sum_viplat,;
          A.tran_42301, A.vklad_acct, B.num_dog, B.data_dog,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ;
          ALLTRIM(B.card_nls) == ALLTRIM(sel_card_nls) ;
          INTO CURSOR sel_prozent ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

      CASE pr_summa_null = .F.  && Если признак=.F., то нулевые суммы не выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          B.data_home, B.data_end, B.den_mes,;
          B.sum_dobav, B.sum_rasch, B.sum_rasxod, B.sum_nalog, B.sum_viplat,;
          A.tran_42301, A.vklad_acct, B.num_dog, B.data_dog,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ;
          ALLTRIM(B.card_nls) == ALLTRIM(sel_card_nls) AND B.sum_dobav > 0.00 ;
          INTO CURSOR sel_prozent ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

    ENDCASE

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* BROWSE WINDOW brows

* INTO CURSOR sel_prozent ;
* INTO TABLE (put_dbf) + ('sel_prozent.dbf') ;

RETURN


************************************************************* PROCEDURE sel_data_sum_cln ****************************************************************

PROCEDURE sel_data_sum_cln

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO CASE
  CASE pr_zarpl = .T.  && Признак наличия зарплатного проекта ВКЛЮЧЕН

    DO CASE
      CASE pr_summa_null = .T.  && Если признак=.T., то нулевые суммы выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          IIF(vvod_data_home < B.data_proz, B.data_proz, vvod_data_home) AS data_home, vvod_data_end AS data_end,;
          SUM(B.den_mes) AS den_mes,;
          SUM(B.sum_dobav) AS sum_dobav,;
          SUM(B.sum_rasch) AS sum_rasch,;
          SUM(B.sum_rasxod) AS sum_rasxod,;
          SUM(B.sum_nalog) AS sum_nalog,;
          SUM(B.sum_viplat) AS sum_viplat,;
          A.tran_42301, A.vklad_acct, B.num_dog, B.data_dog,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ALLTRIM(B.card_nls) == ALLTRIM(sel_card_nls) AND ;
          ALLTRIM(A.kod_proekt) == ALLTRIM(num_proekt) ;
          INTO CURSOR sel_prozent ;
          GROUP BY B.card_nls ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

      CASE pr_summa_null = .F.  && Если признак=.F., то нулевые суммы не выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          IIF(vvod_data_home < B.data_proz, B.data_proz, vvod_data_home) AS data_home, vvod_data_end AS data_end,;
          SUM(B.den_mes) AS den_mes,;
          SUM(B.sum_dobav) AS sum_dobav,;
          SUM(B.sum_rasch) AS sum_rasch,;
          SUM(B.sum_rasxod) AS sum_rasxod,;
          SUM(B.sum_nalog) AS sum_nalog,;
          SUM(B.sum_viplat) AS sum_viplat,;
          A.tran_42301, A.vklad_acct, B.num_dog, B.data_dog,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ALLTRIM(B.card_nls) == ALLTRIM(sel_card_nls) AND ;
          ALLTRIM(A.kod_proekt) == ALLTRIM(num_proekt) AND B.sum_dobav > 0.00 ;
          INTO CURSOR sel_prozent ;
          GROUP BY B.card_nls ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE pr_zarpl = .F.  && Признак наличия зарплатного проекта ВЫКЛЮЧЕН

    DO CASE
      CASE pr_summa_null = .T.  && Если признак=.T., то нулевые суммы выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          IIF(vvod_data_home < B.data_proz, B.data_proz, vvod_data_home) AS data_home, vvod_data_end AS data_end,;
          SUM(B.den_mes) AS den_mes,;
          SUM(B.sum_dobav) AS sum_dobav,;
          SUM(B.sum_rasch) AS sum_rasch,;
          SUM(B.sum_rasxod) AS sum_rasxod,;
          SUM(B.sum_nalog) AS sum_nalog,;
          SUM(B.sum_viplat) AS sum_viplat,;
          A.tran_42301, A.vklad_acct, B.num_dog, B.data_dog,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ;
          ALLTRIM(B.card_nls) == ALLTRIM(sel_card_nls) ;
          INTO CURSOR sel_prozent ;
          GROUP BY B.card_nls ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

      CASE pr_summa_null = .F.  && Если признак=.F., то нулевые суммы не выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          IIF(vvod_data_home < B.data_proz, B.data_proz, vvod_data_home) AS data_home, vvod_data_end AS data_end,;
          SUM(B.den_mes) AS den_mes,;
          SUM(B.sum_dobav) AS sum_dobav,;
          SUM(B.sum_rasch) AS sum_rasch,;
          SUM(B.sum_rasxod) AS sum_rasxod,;
          SUM(B.sum_nalog) AS sum_nalog,;
          SUM(B.sum_viplat) AS sum_viplat,;
          A.tran_42301, A.vklad_acct, B.num_dog, B.data_dog,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ;
          ALLTRIM(B.card_nls) == ALLTRIM(sel_card_nls) AND B.sum_dobav > 0.00 ;
          INTO CURSOR sel_prozent ;
          GROUP BY B.card_nls ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

    ENDCASE

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* BROWSE WINDOW brows

* INTO CURSOR sel_prozent ;
* INTO TABLE (put_dbf) + ('sel_prozent.dbf') ;

RETURN


************************************************************** PROCEDURE sel_data_det_all ****************************************************************

PROCEDURE sel_data_det_all

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO CASE
  CASE pr_zarpl = .T.  && Признак наличия зарплатного проекта ВКЛЮЧЕН

    DO CASE
      CASE pr_summa_null = .T.  && Если признак=.T., то нулевые суммы выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          B.data_home, B.data_end, B.den_mes,;
          B.sum_dobav, B.sum_rasch, B.sum_rasxod, B.sum_nalog, B.sum_viplat,;
          A.tran_42301, B.num_dog, B.data_dog,;
          IIF(EMPTY(ALLTRIM(A.vklad_acct)) = .F. AND LEN(ALLTRIM(A.vklad_acct)) < 20, PADC(ALLTRIM(A.vklad_acct), 20), ALLTRIM(B.card_nls)) AS vklad_acct,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ;
          ALLTRIM(A.kod_proekt) == ALLTRIM(num_proekt) ;
          INTO CURSOR sel_prozent ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

      CASE pr_summa_null = .F.  && Если признак=.F., то нулевые суммы не выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          B.data_home, B.data_end, B.den_mes,;
          B.sum_dobav, B.sum_rasch, B.sum_rasxod, B.sum_nalog, B.sum_viplat,;
          A.tran_42301, B.num_dog, B.data_dog,;
          IIF(EMPTY(ALLTRIM(A.vklad_acct)) = .F. AND LEN(ALLTRIM(A.vklad_acct)) < 20, PADC(ALLTRIM(A.vklad_acct), 20), ALLTRIM(B.card_nls)) AS vklad_acct,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ;
          ALLTRIM(A.kod_proekt) == ALLTRIM(num_proekt) AND B.sum_dobav > 0.00 ;
          INTO CURSOR sel_prozent ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE pr_zarpl = .F.  && Признак наличия зарплатного проекта ВЫКЛЮЧЕН

    DO CASE
      CASE pr_summa_null = .T.  && Если признак=.T., то нулевые суммы выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          B.data_home, B.data_end, B.den_mes,;
          B.sum_dobav, B.sum_rasch, B.sum_rasxod, B.sum_nalog, B.sum_viplat,;
          A.tran_42301, B.num_dog, B.data_dog,;
          IIF(EMPTY(ALLTRIM(A.vklad_acct)) = .F. AND LEN(ALLTRIM(A.vklad_acct)) < 20, PADC(ALLTRIM(A.vklad_acct), 20), ALLTRIM(B.card_nls)) AS vklad_acct,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) ;
          INTO CURSOR sel_prozent ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

      CASE pr_summa_null = .F.  && Если признак=.F., то нулевые суммы не выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          B.data_home, B.data_end, B.den_mes,;
          B.sum_dobav, B.sum_rasch, B.sum_rasxod, B.sum_nalog, B.sum_viplat,;
          A.tran_42301, B.num_dog, B.data_dog,;
          IIF(EMPTY(ALLTRIM(A.vklad_acct)) = .F. AND LEN(ALLTRIM(A.vklad_acct)) < 20, PADC(ALLTRIM(A.vklad_acct), 20), ALLTRIM(B.card_nls)) AS vklad_acct,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND B.sum_dobav > 0.00 ;
          INTO CURSOR sel_prozent ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

    ENDCASE

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* BROWSE WINDOW brows

* INTO CURSOR sel_prozent ;
* INTO TABLE (put_dbf) + ('sel_prozent.dbf') ;

IF INLIST(ALLTRIM(num_branch), '03', '12')
 COPY TO (put_tmp) + ('sel_prozent.xls') XL5
ENDIF

RETURN


************************************************************* PROCEDURE sel_data_sum_all ****************************************************************

PROCEDURE sel_data_sum_all

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO CASE
  CASE pr_zarpl = .T.  && Признак наличия зарплатного проекта ВКЛЮЧЕН

    DO CASE
      CASE pr_summa_null = .T.  && Если признак=.T., то нулевые суммы выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          IIF(vvod_data_home < B.data_proz, B.data_proz, vvod_data_home) AS data_home, vvod_data_end AS data_end,;
          SUM(B.den_mes) AS den_mes,;
          SUM(B.sum_dobav) AS sum_dobav,;
          SUM(B.sum_rasch) AS sum_rasch,;
          SUM(B.sum_rasxod) AS sum_rasxod,;
          SUM(B.sum_nalog) AS sum_nalog,;
          SUM(B.sum_viplat) AS sum_viplat,;
          A.tran_42301, B.num_dog, B.data_dog,;
          IIF(EMPTY(ALLTRIM(A.vklad_acct)) = .F. AND LEN(ALLTRIM(A.vklad_acct)) < 20, PADC(ALLTRIM(A.vklad_acct), 20), ALLTRIM(B.card_nls)) AS vklad_acct,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ;
          ALLTRIM(A.kod_proekt) == ALLTRIM(num_proekt) ;
          INTO CURSOR sel_prozent ;
          GROUP BY B.card_nls ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

      CASE pr_summa_null = .F.  && Если признак=.F., то нулевые суммы не выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          IIF(vvod_data_home < B.data_proz, B.data_proz, vvod_data_home) AS data_home, vvod_data_end AS data_end,;
          SUM(B.den_mes) AS den_mes,;
          SUM(B.sum_dobav) AS sum_dobav,;
          SUM(B.sum_rasch) AS sum_rasch,;
          SUM(B.sum_rasxod) AS sum_rasxod,;
          SUM(B.sum_nalog) AS sum_nalog,;
          SUM(B.sum_viplat) AS sum_viplat,;
          A.tran_42301, B.num_dog, B.data_dog,;
          IIF(EMPTY(ALLTRIM(A.vklad_acct)) = .F. AND LEN(ALLTRIM(A.vklad_acct)) < 20, PADC(ALLTRIM(A.vklad_acct), 20), ALLTRIM(B.card_nls)) AS vklad_acct,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ;
          ALLTRIM(A.kod_proekt) == ALLTRIM(num_proekt) AND B.sum_dobav > 0.00 ;
          INTO CURSOR sel_prozent ;
          GROUP BY B.card_nls ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE pr_zarpl = .F.  && Признак наличия зарплатного проекта ВЫКЛЮЧЕН

    DO CASE
      CASE pr_summa_null = .T.  && Если признак=.T., то нулевые суммы выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          IIF(vvod_data_home < B.data_proz, B.data_proz, vvod_data_home) AS data_home, vvod_data_end AS data_end,;
          SUM(B.den_mes) AS den_mes,;
          SUM(B.sum_dobav) AS sum_dobav,;
          SUM(B.sum_rasch) AS sum_rasch,;
          SUM(B.sum_rasxod) AS sum_rasxod,;
          SUM(B.sum_nalog) AS sum_nalog,;
          SUM(B.sum_viplat) AS sum_viplat,;
          A.tran_42301, B.num_dog, B.data_dog,;
          IIF(EMPTY(ALLTRIM(A.vklad_acct)) = .F. AND LEN(ALLTRIM(A.vklad_acct)) < 20, PADC(ALLTRIM(A.vklad_acct), 20), ALLTRIM(B.card_nls)) AS vklad_acct,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) ;
          INTO CURSOR sel_prozent ;
          GROUP BY B.card_nls ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

      CASE pr_summa_null = .F.  && Если признак=.F., то нулевые суммы не выводятся:

        SELECT B.branch, B.ref_client, B.ref_card, A.val_card, A.name_card, B.fio_rus, B.data_proz, B.kurs,;
          IIF(A.pr_odb = .F., A.tran_42301, B.card_nls) AS card_nls,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, B.stavka, B.sum_ostat,;
          IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
          IIF(vvod_data_home < B.data_proz, B.data_proz, vvod_data_home) AS data_home, vvod_data_end AS data_end,;
          SUM(B.den_mes) AS den_mes,;
          SUM(B.sum_dobav) AS sum_dobav,;
          SUM(B.sum_rasch) AS sum_rasch,;
          SUM(B.sum_rasxod) AS sum_rasxod,;
          SUM(B.sum_nalog) AS sum_nalog,;
          SUM(B.sum_viplat) AS sum_viplat,;
          A.tran_42301, B.num_dog, B.data_dog,;
          IIF(EMPTY(ALLTRIM(A.vklad_acct)) = .F. AND LEN(ALLTRIM(A.vklad_acct)) < 20, PADC(ALLTRIM(A.vklad_acct), 20), ALLTRIM(B.card_nls)) AS vklad_acct,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND B.sum_dobav > 0.00 ;
          INTO CURSOR sel_prozent ;
          GROUP BY B.card_nls ;
          ORDER BY B.branch, grup, A.val_card, A.name_card, B.fio_rus, B.data_home

    ENDCASE

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* BROWSE WINDOW brows

* INTO CURSOR sel_prozent ;
* INTO TABLE (put_dbf) + ('sel_prozent.dbf') ;

* COPY TO (put_tmp) + ('sel_proz.dbf')

IF INLIST(ALLTRIM(num_branch), '03', '12')
 COPY TO (put_tmp) + ('sel_prozent.xls') XL5
ENDIF

RETURN


**************************************************************** PROCEDURE sel_data_sum_tran ************************************************************

PROCEDURE sel_data_sum_tran

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO CASE
  CASE pr_zarpl = .T.  && Признак наличия зарплатного проекта ВКЛЮЧЕН

    DO CASE
      CASE pr_summa_null = .T.  && Если признак=.T., то нулевые суммы выводятся:

        SELECT B.branch, A.val_card, A.name_card, B.kurs, A.tran_42301, A.vklad_acct,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, vvod_data_home AS data_home, vvod_data_end AS data_end,;
          SUM(B.sum_dobav) AS sum_dobav,;
          SUM(B.sum_rasch) AS sum_rasch,;
          SUM(B.sum_rasxod) AS sum_rasxod,;
          SUM(B.sum_nalog) AS sum_nalog,;
          SUM(B.sum_viplat) AS sum_viplat,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ;
          ALLTRIM(A.kod_proekt) == ALLTRIM(num_proekt) ;
          INTO CURSOR sel_prozent ;
          GROUP BY A.tran_42301 ;
          ORDER BY B.branch, grup, A.val_card, A.tran_42301, B.data_home

      CASE pr_summa_null = .F.  && Если признак=.F., то нулевые суммы не выводятся:

        SELECT B.branch, A.val_card, A.name_card, B.kurs, A.tran_42301, A.vklad_acct,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, vvod_data_home AS data_home, vvod_data_end AS data_end,;
          SUM(B.sum_dobav) AS sum_dobav,;
          SUM(B.sum_rasch) AS sum_rasch,;
          SUM(B.sum_rasxod) AS sum_rasxod,;
          SUM(B.sum_nalog) AS sum_nalog,;
          SUM(B.sum_viplat) AS sum_viplat,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND ;
          ALLTRIM(A.kod_proekt) == ALLTRIM(num_proekt) AND B.sum_dobav > 0.00 ;
          INTO CURSOR sel_prozent ;
          GROUP BY A.tran_42301 ;
          ORDER BY B.branch, grup, A.val_card, A.tran_42301, B.data_home

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE pr_zarpl = .F.  && Признак наличия зарплатного проекта ВЫКЛЮЧЕН

    DO CASE
      CASE pr_summa_null = .T.  && Если признак=.T., то нулевые суммы выводятся:

        SELECT B.branch, A.val_card, A.name_card, B.kurs, A.tran_42301, A.vklad_acct,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, vvod_data_home AS data_home, vvod_data_end AS data_end,;
          SUM(B.sum_dobav) AS sum_dobav,;
          SUM(B.sum_rasch) AS sum_rasch,;
          SUM(B.sum_rasxod) AS sum_rasxod,;
          SUM(B.sum_nalog) AS sum_nalog,;
          SUM(B.sum_viplat) AS sum_viplat,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) ;
          INTO CURSOR sel_prozent ;
          GROUP BY A.tran_42301 ;
          ORDER BY B.branch, grup, A.val_card, A.tran_42301, B.data_home

      CASE pr_summa_null = .F.  && Если признак=.F., то нулевые суммы не выводятся:

        SELECT B.branch, A.val_card, A.name_card, B.kurs, A.tran_42301, A.vklad_acct,;
          IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
          B.kvartal, B.mes_kvart, vvod_data_home AS data_home, vvod_data_end AS data_end,;
          SUM(B.sum_dobav) AS sum_dobav,;
          SUM(B.sum_rasch) AS sum_rasch,;
          SUM(B.sum_rasxod) AS sum_rasxod,;
          SUM(B.sum_nalog) AS sum_nalog,;
          SUM(B.sum_viplat) AS sum_viplat,;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
          IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
          IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS num_47411 ;
          FROM sel_account A, acc_cln_sum_pol B ;
          WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND B.sum_dobav > 0.00 ;
          INTO CURSOR sel_prozent ;
          GROUP BY A.tran_42301 ;
          ORDER BY B.branch, grup, A.val_card, A.tran_42301, B.data_home

    ENDCASE

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* BROWSE WINDOW brows

* INTO CURSOR sel_prozent ;
* INTO TABLE (put_dbf) + ('sel_prozent.dbf') ;

RETURN


******************************************************************* PROCEDURE brows_proz ***************************************************************

PROCEDURE brows_proz

DEFINE WINDOW brow_proz  AT 2, 2 SIZE  colvo_rows - 4, colvo_cols - 50 FONT "Courier New Cyr", 10  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW brow_proz CENTER

DO CASE
  CASE put_prn  =  1  && Выборка данных с детализацией рассчетных периодов

    EDIT FIELDS ;
      branch:H = ' Филиал :',;
      ref_client:H = ' Клиент :',;
      ref_card:H = ' Карта :',;
      num_dog:H = ' Номер договора :',;
      fio_rus:H = ' Фамилия Имя Отчество :',;
      dt_hom_end:H = ' Первая дата в квартале :',;
      data_home:H = ' Начальная дата расчета :',;
      data_end:H = ' Конечная дата расчета :',;
      card_nls:H = ' Карточный счет СКС :',;
      name_card:H = ' Тип карты :',;
      kurs:H = ' Курс:':P='999.9999',;
      stavka:H = ' Ставка :':P='999.99',;
      sum_ostat:H = ' Сумма остатка по счету :':P='999 999.99',;
      kvartal:H = ' Номер квартала :',;
      mes_kvart:H = ' Номер месяца в квартале :',;
      den_mes:H = ' Дней в месяце :',;
      sum_dobav:H = ' Сумма к доначислению :':P='9 999.99',;
      sum_viplat:H = ' Сумма процентов к выплате :':P='9 999.99',;
      num_47411:H = ' Счет начисленных процентов :' ;
      TITLE ' Просмотр данных по рассчитанным процентам ' ;
      WINDOW brow_proz

*         sum_rasch:H = ' Сумма отнесенная на счет расчетов по выплате :':P='9 999.99',;
*         sum_rasxod:H = ' Сумма отнесенная на расходы банка :':P='9 999.99',;
*         sum_nalog:H = ' Сумма подоходного налога :':P='9 999.99',;

  CASE put_prn = 3  && Выборка данных с суммированием рассчетных периодов

    EDIT FIELDS ;
      branch:H = ' Филиал :',;
      ref_client:H = ' Клиент :',;
      ref_card:H = ' Карта :',;
      num_dog:H = ' Номер договора :',;
      fio_rus:H = ' Фамилия Имя Отчество :',;
      dt_hom_end:H = ' Первая дата в квартале :',;
      data_home:H = ' Начальная дата расчета :',;
      data_end:H = ' Конечная дата расчета :',;
      card_nls:H = ' Карточный счет СКС :',;
      name_card:H = ' Тип карты :',;
      kurs:H = ' Курс:':P='999.9999',;
      stavka:H = ' Ставка :':P='999.99',;
      kvartal:H = ' Номер квартала :',;
      mes_kvart:H = ' Номер месяца в квартале :',;
      den_mes:H = ' Дней в месяце :',;
      sum_dobav:H = ' Сумма к доначислению :':P='9 999.99',;
      sum_viplat:H = ' Сумма процентов к выплате :':P='9 999.99',;
      num_47411:H = ' Счет начисленных процентов :' ;
      TITLE ' Просмотр данных по рассчитанным процентам ' ;
      WINDOW brow_proz

*         sum_rasch:H = ' Сумма отнесенная на счет расчетов по выплате :':P='9 999.99',;
*         sum_rasxod:H = ' Сумма отнесенная на расходы банка :':P='9 999.99',;
*         sum_nalog:H = ' Сумма подоходного налога :':P='9 999.99',;

  CASE put_prn = 5  && Выборка данных с суммированием по транзитным счетам

    EDIT FIELDS ;
      branch:H = ' Филиал :',;
      data_home:H = ' Начальная дата расчета :',;
      data_end:H = ' Конечная дата расчета :',;
      tran_42301:H = ' Транзитный счет :',;
      name_card:H = ' Тип карты :',;
      name_card:H = ' Вид карты :',;
      kurs:H = ' Курс:':P='999.9999',;
      kvartal:H = ' Номер квартала :',;
      mes_kvart:H = ' Номер месяца в квартале :',;
      sum_dobav:H = ' Сумма к доначислению :':P='9 999.99',;
      sum_viplat:H = ' Сумма процентов к выплате :':P='9 999.99',;
      num_47411:H = ' Счет начисленных процентов :' ;
      TITLE ' Просмотр данных по рассчитанным процентам ' ;
      WINDOW brow_proz

*         sum_rasch:H = ' Сумма отнесенная на счет расчетов по выплате :':P='9 999.99',;
*         sum_rasxod:H = ' Сумма отнесенная на расходы банка :':P='9 999.99',;
*         sum_nalog:H = ' Сумма подоходного налога :':P='9 999.99',;

ENDCASE

RELEASE WINDOW brow_proz

RETURN


******************************************************************** PROCEDURE prn_proz ****************************************************************

PROCEDURE prn_proz

DO FORM (put_scr) + ('data_prn.scx')

IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

  text1 = 'Печатный отчет на экран'
  text2 = 'Печатный отчет на принтер'
  l_bar = 3
  =popup_9(text1, text2, text3, text4, l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE put_prn = 1  && Отчет с детализацией рассчетных периодов - книжная - форма № 1

        DO CASE
          CASE BAR() = 1

            REPORT FORM (put_frx) + ('proz_mesaz_42301_d_v.frx') NOEJECT PREVIEW
            * REPORT FORM (put_frx) + ('proz_mesaz_42301_d_v.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_d_v.doc') ASCII

          CASE BAR() = 3

            DO CASE
              CASE par_prompt = .T.  && При выдаче отчета на печать предлагать панель выбора подключенных принтеров

                REPORT FORM (put_frx) + ('proz_mesaz_42301_d_v.frx') NOEJECT NOCONSOLE TO PRINTER PROMPT
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_d_v.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_d_v.doc') ASCII

              CASE par_prompt = .F.  && При выдаче отчета на печать не предлагать панель выбора подключенных принтеров

                REPORT FORM (put_frx) + ('proz_mesaz_42301_d_v.frx') NOEJECT NOCONSOLE TO PRINTER
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_d_v.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_d_v.doc') ASCII

            ENDCASE

        ENDCASE

      CASE put_prn = 3  && Отчет с детализацией периодов коротких счетов нет - альбомная - форма № 2

        DO CASE
          CASE BAR() = 1

            REPORT FORM (put_frx) + ('proz_mesaz_42301_d_g.frx') NOEJECT PREVIEW
            * REPORT FORM (put_frx) + ('proz_mesaz_42301_d_g.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_d_g.doc') ASCII

          CASE BAR() = 3

            DO CASE
              CASE par_prompt = .T.  && При выдаче отчета на печать предлагать панель выбора подключенных принтеров

                REPORT FORM (put_frx) + ('proz_mesaz_42301_d_g.frx') NOEJECT NOCONSOLE TO PRINTER PROMPT
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_d_g.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_d_g.doc') ASCII

              CASE par_prompt = .F.  && При выдаче отчета на печать не предлагать панель выбора подключенных принтеров

                REPORT FORM (put_frx) + ('proz_mesaz_42301_d_g.frx') NOEJECT NOCONSOLE TO PRINTER
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_d_g.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_d_g.doc') ASCII

            ENDCASE

        ENDCASE

      CASE put_prn = 5  && Отчет с суммированием периодов коротких счетов нет - альбомная - форма № 4

        DO CASE
          CASE BAR() = 1

            REPORT FORM (put_frx) + ('proz_mesaz_42301_pn.frx') NOEJECT PREVIEW
            * REPORT FORM (put_frx) + ('proz_mesaz_42301_pn.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_pn.doc') ASCII

          CASE BAR() = 3

            DO CASE
              CASE par_prompt = .T.  && При выдаче отчета на печать предлагать панель выбора подключенных принтеров

                REPORT FORM (put_frx) + ('proz_mesaz_42301_pn.frx') NOEJECT NOCONSOLE TO PRINTER PROMPT
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_pn.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_pn.doc') ASCII

              CASE par_prompt = .F.  && При выдаче отчета на печать не предлагать панель выбора подключенных принтеров

                REPORT FORM (put_frx) + ('proz_mesaz_42301_pn.frx') NOEJECT NOCONSOLE TO PRINTER
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_pn.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_pn.doc') ASCII

            ENDCASE

        ENDCASE

      CASE put_prn = 7  && Отчет с суммированием периодов короткие счета есть - альбомная - форма № 5

        DO CASE
          CASE BAR() = 1

            REPORT FORM (put_frx) + ('proz_mesaz_42301_pv.frx') NOEJECT PREVIEW
            * REPORT FORM (put_frx) + ('proz_mesaz_42301_pv.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_pv.doc') ASCII

          CASE BAR() = 3

            DO CASE
              CASE par_prompt = .T.  && При выдаче отчета на печать предлагать панель выбора подключенных принтеров

                REPORT FORM (put_frx) + ('proz_mesaz_42301_pv.frx') NOEJECT NOCONSOLE TO PRINTER PROMPT
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_pv.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_pv.doc') ASCII

              CASE par_prompt = .F.  && При выдаче отчета на печать не предлагать панель выбора подключенных принтеров

                REPORT FORM (put_frx) + ('proz_mesaz_42301_pv.frx') NOEJECT NOCONSOLE TO PRINTER
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_pv.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_pv.doc') ASCII

            ENDCASE

        ENDCASE

      CASE put_prn = 9  && Отчет с сумированием по транзитным счетам - альбомная - форма № 6

        DO CASE
          CASE BAR() = 1

            REPORT FORM (put_frx) + ('proz_mesaz_42301_tr.frx') NOEJECT PREVIEW
            * REPORT FORM (put_frx) + ('proz_mesaz_42301_tr.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_tr.doc') ASCII

          CASE BAR() = 3

            DO CASE
              CASE par_prompt = .T.  && При выдаче отчета на печать предлагать панель выбора подключенных принтеров

                REPORT FORM (put_frx) + ('proz_mesaz_42301_tr.frx') NOEJECT NOCONSOLE TO PRINTER PROMPT
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_tr.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_tr.doc') ASCII

              CASE par_prompt = .F.  && При выдаче отчета на печать не предлагать панель выбора подключенных принтеров

                REPORT FORM (put_frx) + ('proz_mesaz_42301_tr.frx') NOEJECT NOCONSOLE TO PRINTER
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_tr.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_tr.doc') ASCII

            ENDCASE

        ENDCASE

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  ENDIF

ENDIF

RETURN


***************************************************************** PROCEDURE formirov_pb ****************************************************************

PROCEDURE formirov_pb
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

vvod_data_pak = MaxDataMes(DATE())
put_bar = 1

sel_1 = SELECT()

num_mesaz = MONTH(new_data_odb)
num_god = ALLTRIM(STR(YEAR(new_data_odb)))

DO vid_kvartal_1 IN Raschet_proz_42301

DO FORM (put_scr) + ('sel_data_proz.scx')

IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

  num_mesaz = MONTH(vvod_data_end)
  num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

  DO vid_kvartal_2 IN Raschet_proz_42301

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
*    IF pr_zarpl = .T.
*      =soob('Внимание! Формирование файла для - ' + ALLTRIM(title_zarpl))
*    ENDIF
* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  text1 = 'Формирование данных для выгрузки в УПК по всем видам валют'
  text2 = 'Формирование данных для выгрузки в УПК по коду валюты RUR'
  text3 = 'Формирование данных для выгрузки в УПК по коду валюты USD'
  text4 = 'Формирование данных для выгрузки в УПК по коду валюты EUR'
  l_bar = 7
  =popup_9(text1, text2, text3, text4, l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    DO CASE
      CASE BAR() = 1  && Формирование данных для выгрузки в УПК по всем видам валют
        sel_val_card = SPACE(0)

      CASE BAR() = 3  && Формирование данных для выгрузки в УПК по коду валюты RUR
        sel_val_card = 'RUR'

      CASE BAR() = 5  && Формирование данных для выгрузки в УПК по коду валюты USD
        sel_val_card = 'USD'

      CASE BAR() = 7  && Формирование данных для выгрузки в УПК по коду валюты EUR
        sel_val_card = 'EUR'

    ENDCASE

  ELSE
    RETURN
  ENDIF

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  DO CASE
    CASE EMPTY(ALLTRIM(sel_val_card)) = .T.

      SELECT DISTINCT card_nls, card_acct, tran_42301, val_card, name_card, kod_proekt,;
        IIF(LEFT(card_nls, 5) == '40817', '1', IIF(LEFT(card_nls, 5) == '40820', '1', '9')) AS grup ;
        FROM account ;
        WHERE ALLTRIM(vid_card) == '1' ;
        UNION ;
        SELECT DISTINCT card_nls, card_acct, tran_42301, val_card, name_card, kod_proekt,;
        IIF(LEFT(card_nls, 5) == '40817', '1', IIF(LEFT(card_nls, 5) == '40820', '1', '9')) AS grup ;
        FROM acc_del ;
        WHERE ALLTRIM(vid_card) == '1' AND (end_bal <> 0 OR isx_ost_m <> 0) ;
        INTO CURSOR sel_account ;
        ORDER BY 1

    CASE EMPTY(ALLTRIM(sel_val_card)) = .F.

      SELECT DISTINCT card_nls, card_acct, tran_42301, val_card, name_card, kod_proekt,;
        IIF(LEFT(card_nls, 5) == '40817', '1', IIF(LEFT(card_nls, 5) == '40820', '1', '9')) AS grup ;
        FROM account ;
        WHERE ALLTRIM(vid_card) == '1' AND ALLTRIM(val_card) == ALLTRIM(sel_val_card) ;
        UNION ;
        SELECT DISTINCT card_nls, card_acct, tran_42301, val_card, name_card, kod_proekt,;
        IIF(LEFT(card_nls, 5) == '40817', '1', IIF(LEFT(card_nls, 5) == '40820', '1', '9')) AS grup ;
        FROM acc_del ;
        WHERE ALLTRIM(vid_card) == '1' AND (end_bal <> 0 OR isx_ost_m <> 0) AND ALLTRIM(val_card) == ALLTRIM(sel_val_card) ;
        INTO CURSOR sel_account ;
        ORDER BY 1

  ENDCASE

  IF _TALLY <> 0

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SELECT B.branch, RIGHT(ALLTRIM( B.card_nls), 10) AS record, B.ref_card AS nn_doc, B.fio_rus, B.data_proz, B.kurs, A.val_card, A.name_card,;
      A.card_acct, B.card_nls, vvod_data_home AS data_home, vvod_data_end AS data_end,;
      PADL(DAY(B.data_pak), 2, '0') + PADL(MONTH(B.data_pak), 2, '0') + RIGHT(ALLTRIM(STR(YEAR(B.data_pak))),1) AS pril,;
      SUM(B.sum_viplat) AS sum_viplat, A.kod_proekt, A.grup ;
      FROM sel_account A, acc_cln_sum_pol B ;
      WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND sum_viplat <> 0 ;
      INTO CURSOR sel_prozent ;
      GROUP BY B.card_nls ;
      ORDER BY B.branch, A.grup, A.val_card, A.name_card, B.fio_rus, B.data_home

    SELECT branch, record, pril, data_end AS data_val, nn_doc, ALLTRIM(val_card) AS kodval, name_card,;
      '516' AS kod_oper, fio_rus, card_acct, data_home,;
      IIF(INLIST(LEFT(ALLTRIM(card_nls), 5), '40817'),;
      IIF(ALLTRIM(val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
      IIF(INLIST(LEFT(ALLTRIM(card_nls), 5), '40820'),;
      IIF(ALLTRIM(val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS debit,;
      card_nls AS credit, sum_viplat AS summa, SPACE(0) AS summa_prop, .T. AS export, .T. AS zapis,;
      'Начислены проценты' AS prim, kod_proekt, grup ;
      FROM sel_prozent ;
      INTO CURSOR sel_export ;
      ORDER BY branch, grup, val_card, name_card, fio_rus, data_home

    colvo_zap = _TALLY

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    IF colvo_zap <> 0

      SELECT popol_ckc

*!*  Имя файла:  Pb02FFFK.NNN
*!*  P - идентификатор файла
*!*  b - признак карт VISA
*!*  02 - код банка
*!*  fff - Отделение (филиал) банка

*!*  001-Москва,
*!*  002-Ленск,
*!*  003-Орел,
*!*  004-Мирный,
*!*  005-Удачный,
*!*  006-Якутск,
*!*  007-Айхал,
*!*  008-Архангельск,
*!*  009-Краснодар,
*!*  010-Новосибирск,
*!*  011-Иркутск,
*!*  013-Головной офис

*!*  k - Номер файла за день (0,1,2,3…)
*!*  nnn - Юлианская дата (количество дней с начала года)

      colvo_ckc = 7
      data_ckc = DATE()
      explog_ckc = .F.

      DO WHILE .T.

        DO FORM (put_scr) + ('popol_ckc.scx')

        IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

          exp_file = 'PB02' + (fff) + ALLTRIM(STR(colvo_ckc)) + '.' + (nnn)

          IF NOT FILE((put_out_pb) + (exp_file))

            explog_ckc = .T.
            EXIT

          ELSE

            =err('Внимание! Файл - ' + (put_out_pb) + (exp_file) + ' - в директории уже существует !!! ')

            text1 = 'Вы желаете перезаписать файл с данным номером'
            text2 = 'Вы желаете ввести новый номер для файла экспорта'
            l_bar = 3
            =popup_9(text1, text2, text3, text4, l_bar)
            ACTIVATE POPUP vibor
            RELEASE POPUPS vibor

            IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

              DO CASE
                CASE BAR() = 1
                  explog_ckc = .T.
                  EXIT

                CASE BAR() = 3
                  explog_ckc = .F.
                  LOOP

              ENDCASE

            ENDIF
          ENDIF

        ELSE
          EXIT
        ENDIF

      ENDDO

      IF explog_ckc = .T.

        ACTIVATE WINDOW poisk
        @ WROWS()/3-0.2,3 SAY PADC('Внимание! Начато формирование данных для комплекса "TRANSMASTER" ... ',WCOLS())

        SET TEXTMERGE ON
        SET TEXTMERGE NOSHOW

        IF FILE((put_out_pb) + (exp_file))
          DELETE FILE (put_out_pb) + (exp_file)
        ENDIF

        _text = FCREATE((put_out_pb) + (exp_file))

        IF _text < 1

          =err('Внимание! Ошибка при создании файла экспорта проводок ... ')
          =FCLOSE(_text)

        ELSE

          data_ckc = DATE()

          SELECT sel_export
          GO TOP

* ---------------------------------------------------------------------------------------------- Формирование заголовка файла -------------------------------------------------------------------------------------------------- *

          SELECT SUM(summa) AS summa_all ;
            FROM sel_export ;
            INTO CURSOR sel_summa

          s_code = '00'
          s_file_date = ALLTRIM(STR(YEAR(data_ckc))) + PADL(ALLTRIM(STR(MONTH(data_ckc))), 2, '0') + PADL(ALLTRIM(STR(DAY(data_ckc))), 2, '0')
          s_file_time = SUBSTR(TIME(), 1, 2) + SUBSTR(TIME(), 4, 2) + SUBSTR(TIME(), 7, 2)
          s_source = 'K'
          s_file_name = (exp_file)
          s_paym_ord_no = SPACE(8)
          s_paym_ord_date = (s_file_date)
          s_paym_ord_sum = PADL(ALLTRIM(STR(sel_summa.summa_all, 12, 2)), 12, '0')
          s_p_file_edition = '02'

          stroka_header_record = s_code + s_file_date + s_file_time + s_source + s_file_name + s_paym_ord_no + s_paym_ord_date + s_paym_ord_sum + s_p_file_edition

          =FPUTS(_text,(stroka_header_record))

* --------------------------------------------------------------------------------------------- Формирование строк детализации ------------------------------------------------------------------------------------------------ *

          SELECT sel_export
          GO TOP

          SCAN

            s_code = '11'
            s_card_no = SPACE(19)
            s_tran_date = ALLTRIM(STR(YEAR(data_ckc))) + PADL(ALLTRIM(STR(MONTH(data_ckc))), 2, '0') + PADL(ALLTRIM(STR(DAY(data_ckc))), 2, '0')
            s_tran_type = ALLTRIM(sel_export.kod_oper)
            s_amount = PADL(ALLTRIM(STR(sel_export.summa, 12, 2)), 12)
            s_currency = ALLTRIM(sel_export.kodval)
            s_bank_code = '02'
            s_branch_code = new_branch(num_branch)
            s_wrpl_no = '000'
            s_tab_no = '000'
            s_tran_no = PADL(ALLTRIM(STR(RECNO())), 8, '0')
            s_counter_party = SPACE(200)
            s_deal_desc = SPACE(140)

            s_card_acct = PADR(ALLTRIM(sel_export.card_acct), 20)

            stroka_detal_record = s_code + s_card_no + s_tran_date + s_tran_type + s_amount + s_currency + s_bank_code + s_branch_code + ;
              s_wrpl_no + s_tab_no + s_tran_no + s_counter_party + s_deal_desc + s_card_acct

            =FPUTS(_text,(stroka_detal_record))

          ENDSCAN

* ------------------------------------------------------------------------------------------------- Формирование окончания файла ---------------------------------------------------------------------------------------------- *

          SELECT COUNT(*) AS colvo_all ;
            FROM sel_export ;
            INTO CURSOR sel_colvo

          s_code = '99'
          s_record_cnt = PADL(ALLTRIM(STR(sel_colvo.colvo_all)), 10)
          s_paym_ord_sum = PADL(ALLTRIM(STR(sel_summa.summa_all, 12, 2)), 12)
          s_check_sum = SPACE(12)
          s_eof = CHR(26)
          stroka_footer_record = s_code + s_record_cnt + s_paym_ord_sum + s_check_sum + s_eof

          =FPUTS(_text,(stroka_footer_record))

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

          =FCLOSE(_text)

        ENDIF

        SET TEXTMERGE OFF

        =INKEY(2)
        ACTIVATE WINDOW poisk
        @ WROWS()/3-0.2,3 SAY PADC('Формирование данных для комплекса "TRANSMASTER" успешно завершено ... ',WCOLS())
        =INKEY(2)
        DEACTIVATE WINDOW poisk

        =soob('Экспортный файл - ' + (exp_file) + ' - в директорию ' + (put_out_pb) + ' сформирован ... ')

        ACTIVATE WINDOW poisk
        @ WROWS()/3-0.2,3 SAY PADC('Количество сформированных проводок - ' + ALLTRIM(TRANSFORM(sel_colvo.colvo_all, '999 999 999')) + ;
          '  На общую сумму - ' + ALLTRIM(TRANSFORM(sel_summa.summa_all, '999 999 999 999.99')), WCOLS())
        =INKEY(10)
        DEACTIVATE WINDOW poisk

        SELECT popol_arx
        SET ORDER TO record

        SELECT sel_export

        SCAN

          rec = ALLTRIM(sel_export.record) + ALLTRIM(sel_export.pril)

          SELECT popol_arx

          poisk = SEEK(rec)

          IF poisk = .T.

            UPDATE popol_arx SET ;
              record = sel_export.record,;
              pril = sel_export.pril,;
              export = sel_export.export,;
              zapis = sel_export.zapis,;
              data_val = sel_export.data_val,;
              nn_doc = sel_export.nn_doc,;
              kodval = sel_export.kodval,;
              kod_oper = sel_export.kod_oper,;
              fio_rus = sel_export.fio_rus,;
              card_acct = sel_export.card_acct,;
              debit = sel_export.debit,;
              credit = sel_export.credit,;
              summa = sel_export.summa,;
              summa_prop = sel_export.summa_prop,;
              prim = sel_export.prim,;
              kod_proekt = sel_export.kod_proekt ;
              WHERE ALLTRIM(record) + ALLTRIM(pril) == rec

          ELSE

            INSERT INTO popol_arx ;
              (record, pril, export, zapis, data_val, nn_doc, kodval, kod_oper, fio_rus, card_acct, debit, credit, summa, summa_prop, prim, kod_proekt) ;
              VALUES ;
              (sel_export.record, sel_export.pril, sel_export.export, sel_export.zapis, sel_export.data_val, sel_export.nn_doc,;
              sel_export.kodval, sel_export.kod_oper, sel_export.fio_rus, sel_export.card_acct,;
              sel_export.debit, sel_export.credit, sel_export.summa, sel_export.summa_prop, sel_export.prim, sel_export.kod_proekt)

          ENDIF

          SELECT sel_export
        ENDSCAN

      ENDIF

    ELSE
      =err('Внимание! По выбранному Вами условию данных не обнаружено ...')
    ENDIF

  ELSE
    =err('Внимание! Вы отказались от формирования выходного файла для УПК ...')
  ENDIF

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

IF INLIST(ALLTRIM(num_branch), '04')  && Процедура копирования выходного файла для УПК разработана по просьбе Ардашева Дмитрия Филиал № 4 (24.03.2004)

  text1='В ы п о л н и т ь  режим копирования файла в почтовую директорию'
  text2='Н е  в ы п о л н я т ь  режим копирования файла в почтовую директорию'
  l_bar=3
  =popup_9(text1,text2,text3,text4,l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    DO CASE
      CASE BAR() = 1  && В ы п о л н и т ь  режим копирования файла в почтовую директорию

        IF FILE((put_out_pb) + (exp_file))

          =soob('Внимание! Файл в директории ' + (put_out_pb) + (exp_file) + ' найден ...')

          IF FILE((put_upk_1) + (exp_file))
            DELETE FILE (put_upk_1) + (exp_file)
          ENDIF

          IF FILE((put_upk_2) + (exp_file))
            DELETE FILE (put_upk_2) + (exp_file)
          ENDIF

          COPY FILE (put_out_pb) + (exp_file) TO (put_upk_1) + (exp_file)
          COPY FILE (put_out_pb) + (exp_file) TO (put_upk_2) + (exp_file)

          =soob('Внимание! Файл - ' + (put_out_pb) + (exp_file) + ' - скопирован по первому пути ' + (put_upk_1))
          =soob('Внимание! Файл - ' + (put_out_pb) + (exp_file) + ' - скопирован по второму пути ' + (put_upk_2))

          IF FILE((put_out_pb) + (exp_file))
            DELETE FILE (put_out_pb) + (exp_file)
          ENDIF

        ELSE

          =err('Внимание! Файл - ' + (put_out_pb) + (exp_file) + ' - в директории не существует !!! ')

        ENDIF

      CASE BAR() = 3  && Н е  в ы п о л н я т ь  режим копирования файла в почтовую директорию

        =err('Вы отказались копировать файл - ' + (put_out_pb) + (exp_file) + ' в почтовую директорию.')

    ENDCASE

  ENDIF

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

SELECT (sel_1)
RETURN


********************************************************************** PROCEDURE export_odb ************************************************************

PROCEDURE export_odb
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

vvod_data_pak = MaxDataMes(DATE())
put_bar = 1

sel_1 = SELECT()

num_mesaz = MONTH(new_data_odb)
num_god = ALLTRIM(STR(YEAR(new_data_odb)))

DO vid_kvartal_1 IN Raschet_proz_42301

DO FORM (put_scr) + ('sel_data_proz.scx')

IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

  num_mesaz = MONTH(vvod_data_end)
  num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

  DO vid_kvartal_2 IN Raschet_proz_42301

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  SELECT DISTINCT A.fio_rus, A.card_nls, A.tran_42301, B.tran_47411, A.vklad_acct, A.val_card, A.name_card, A.pr_odb, A.kod_proekt ;
    FROM account A, type_liz_card B ;
    WHERE A.stavka <> 0 AND ALLTRIM(A.tran_42301) == ALLTRIM(B.tran_42301) ;
    UNION ;
    SELECT DISTINCT A.fio_rus, A.card_nls, A.tran_42301, B.tran_47411, A.vklad_acct, A.val_card, A.name_card, A.pr_odb, A.kod_proekt ;
    FROM acc_del A, type_liz_card B ;
    WHERE A.stavka <> 0 AND (A.end_bal <> 0 OR A.isx_ost_m <> 0) AND ALLTRIM(A.tran_42301) == ALLTRIM(B.tran_42301) ;
    INTO CURSOR sel_account ;
    ORDER BY 1, 2

*    BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  SELECT B.branch, B.ref_client, B.ref_card, A.val_card, B.fio_rus, B.data_proz, B.kurs,;
    IIF(LEFT(B.card_nls, 5) == '40817', '1', IIF(LEFT(B.card_nls, 5) == '40820', '1', '9')) AS grup,;
    A.card_nls, A.name_card, A.vklad_acct, A.tran_42301, A.tran_47411,;
    IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40817',;
    IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_proz, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_proz, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_proz, SPACE(20)))),;
    IIF(LEFT(ALLTRIM(A.card_nls), 5) == '40820',;
    IIF(ALLTRIM(A.val_card) == 'RUR', num_70203_rur_norez, IIF(ALLTRIM(A.val_card) == 'USD', num_70203_usd_norez, IIF(ALLTRIM(A.val_card) == 'EUR', num_70203_eur_norez, SPACE(20)))), SPACE(20))) AS nls_deb,;
    IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
    IIF(vvod_data_home < B.data_proz, B.data_proz, vvod_data_home) AS data_home, vvod_data_end AS data_end,;
    SUM(B.sum_dobav) AS sum_cred,;
    data_pak AS data_val,;
    '516' AS kod_oper,;
    'Начислены проценты за период ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) AS kod_name ;
    FROM sel_account A, acc_cln_sum_pol B ;
    WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND B.sum_dobav <> 0 AND ;
    ALLTRIM(A.kod_proekt) == ALLTRIM(num_proekt) ;
    INTO CURSOR sel_prozent ;
    GROUP BY B.card_nls ;
    ORDER BY B.branch, A.val_card, B.fio_rus

*     BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  IF _TALLY <> 0

    ACTIVATE WINDOW poisk
    @ WROWS()/3-0.2,3 SAY PADC('Внимание! Начат экспорт данных из комплекса "CARD-банк VISA" ... ',WCOLS())
    =INKEY(2)

        DO expsum_rsbank IN Raschet_proz_42301

    ACTIVATE WINDOW poisk
    @ WROWS()/3-0.2,3 SAY PADC('Экспорт данных из комплекса "CARD-банк VISA" успешно завершен ... ',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ENDIF

SELECT(sel)

RETURN


**************************************************************** PROCEDURE expsum_rsbank **************************************************************

PROCEDURE expsum_rsbank

IF flag_prichislen = .T.

  SELECT branch, data_val, val_card, name_card, nls_deb, card_nls, nls_deb AS debit, tran_47411 AS kredit, SUM(sum_cred) AS sum_cred, '1' AS order_oper ;
    FROM sel_prozent ;
    GROUP BY val_card, tran_47411, nls_deb ;
    UNION ;
    SELECT branch, data_val, val_card, name_card, nls_deb, card_nls, tran_47411 AS debit, tran_42301 AS kredit, SUM(sum_cred) AS sum_cred, '2' AS order_oper ;
    FROM sel_prozent ;
    GROUP BY val_card, tran_42301, tran_47411 ;
    INTO CURSOR rs_vip_doc ;
    ORDER BY nls_deb, card_nls, order_oper

  GO TOP

*   BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
  DO scan_name_bank IN Visa
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  SELECT ;
    PADL(ALLTRIM(STR(RECNO())),5) AS numdoc,;
    data_val,;
    val_card,;
    SPACE(9) AS bik_plat,;
    SPACE(20) AS kor_plat,;
    debit AS liz_plat,;
    SPACE(9) AS bik_polu,;
    SPACE(20) AS kor_polu,;
    kredit AS liz_polu,;
    sum_cred AS summa,;
    SPACE(80) AS name_plat,;
    SPACE(80) AS name_polu,;
    PADR(name_bank,80) AS bank_plat,;
    PADR(name_bank,80) AS bank_polu,;
    PADR(IIF(LEFT(ALLTRIM(debit), 3) = '706' AND INLIST(LEFT(ALLTRIM(kredit), 3), '474'), (txt_ps_702_423) + SPACE(1) + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end),;
    IIF(LEFT(ALLTRIM(debit), 3) = '474' AND INLIST(LEFT(ALLTRIM(kredit), 3), '408'), (txt_pd_702_423) + SPACE(1) + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end), SPACE(0))), 160) AS osnov,;
    IIF(val_card == 'RUR', ROUND(VAL(num_oper_rur_v), 0), ROUND(VAL(num_oper_val_v), 0)) AS num_oper,;
    IIF(val_card == 'RUR', ROUND(VAL(num_pach_rur_v), 0), ROUND(VAL(num_pach_val_v), 0)) AS num_pach,;
    PADR(IIF(INLIST(LEFT(ALLTRIM(debit), 3), '706') AND INLIST(LEFT(ALLTRIM(kredit), 3), '474'), '09',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '474') AND INLIST(LEFT(ALLTRIM(kredit), 3), '408'), '09',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '408', '302', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '706'), '02',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '301', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '408', '302', '303'), '09', '09')))), 2, '0') AS shifr,;
    PADR(IIF(INLIST(LEFT(ALLTRIM(debit), 3), '302', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '202'), ' 55',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '202') AND INLIST(LEFT(ALLTRIM(kredit), 3), '302', '303'), ' 16', '  0')), 3) AS symbol,;
    '  0' AS symbnotbal,;
    IIF(SUBSTR(ALLTRIM(debit), 6, 3) == ALLTRIM(kod_val_rur) AND INLIST(SUBSTR(ALLTRIM(kredit), 6, 3), kod_val_usd, kod_val_eur) AND INLIST(ALLTRIM(val_card), 'USD', 'EUR'),;
    kod_carry_R, SPACE(10)) AS kind_carry,;
    PADR(IIF(INLIST(LEFT(ALLTRIM(debit), 3), '302', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '202'), ' 3',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '202') AND INLIST(LEFT(ALLTRIM(kredit), 3), '302', '303'), ' 3',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '302', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '706'), ' 1',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '301', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '408', '302', '303'), ' 6', ' 6')))), 2) AS kind_oper ;
    FROM rs_vip_doc ;
    INTO CURSOR expdocum ;
    ORDER BY 1

*   BROWSE WINDOW brows

* INTO CURSOR expdocum ;
* INTO TABLE (put_dbf) + ('expdocum.dbf') ;
*  ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

ELSE        &&  Если flag_prichislen = .F.

  SELECT branch, data_val, val_card, name_card, nls_deb AS debit, tran_42301 AS kredit, SUM(sum_cred) AS sum_cred ;
    FROM sel_prozent ;
    INTO CURSOR rs_vip_doc ;
    GROUP BY nls_deb, tran_42301 ;
    ORDER BY nls_deb, tran_42301

  GO TOP

* BROWSE WINDOW brows

* INTO CURSOR rs_vip_doc ;
* INTO TABLE (put_dbf) + ('rs_vip_doc.dbf') ;

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
  DO scan_name_bank IN Visa
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  SELECT ;
    PADL(ALLTRIM(STR(RECNO())),5) AS numdoc,;
    data_val,;
    val_card,;
    SPACE(9) AS bik_plat,;
    SPACE(20) AS kor_plat,;
    debit AS liz_plat,;
    SPACE(9) AS bik_polu,;
    SPACE(20) AS kor_polu,;
    kredit AS liz_polu,;
    sum_cred AS summa,;
    SPACE(80) AS name_plat,;
    SPACE(80) AS name_polu,;
    PADR(name_bank,80) AS bank_plat,;
    PADR(name_bank,80) AS bank_polu,;
    PADR(IIF(LEFT(ALLTRIM(debit), 3) = '706' AND INLIST(LEFT(ALLTRIM(kredit), 3), '408', '302', '303'), (txt_ps_702_423) + SPACE(1) + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end), SPACE(0)), 160) AS osnov,;
    IIF(val_card == 'RUR', ROUND(VAL(num_oper_rur_o), 0), ROUND(VAL(num_oper_val_o), 0)) AS num_oper,;
    IIF(val_card == 'RUR', ROUND(VAL(num_pach_rur_o), 0), ROUND(VAL(num_pach_val_o), 0)) AS num_pach,;
    PADR(IIF(INLIST(LEFT(ALLTRIM(debit), 3), '408', '302', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '202'), '09',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '202') AND INLIST(LEFT(ALLTRIM(kredit), 3), '408', '302', '303'), '09',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '408', '302', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '706'), '02',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '301', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '408', '302', '303'), '09', '09')))), 2, '0') AS shifr,;
    PADR(IIF(INLIST(LEFT(ALLTRIM(debit), 3), '408', '302', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '202'), ' 55',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '202') AND INLIST(LEFT(ALLTRIM(kredit), 3), '408', '302', '303'), ' 16', '  0')), 3) AS symbol,;
    '  0' AS symbnotbal,;
    IIF(SUBSTR(ALLTRIM(debit), 6, 3) == ALLTRIM(kod_val_rur) AND INLIST(SUBSTR(ALLTRIM(kredit), 6, 3), kod_val_usd, kod_val_eur) AND INLIST(ALLTRIM(val_card), 'USD', 'EUR'),;
    kod_carry_R, SPACE(10)) AS kind_carry,;
    PADR(IIF(INLIST(LEFT(ALLTRIM(debit), 3), '408', '302', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '202'), ' 3',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '202') AND INLIST(LEFT(ALLTRIM(kredit), 3), '408', '302', '303'), ' 3',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '408', '302', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '706'), ' 1',;
    IIF(INLIST(LEFT(ALLTRIM(debit), 3), '301', '303') AND INLIST(LEFT(ALLTRIM(kredit), 3), '408', '302', '303'), ' 6', ' 6')))), 2) AS kind_oper ;
    FROM rs_vip_doc ;
    INTO CURSOR expdocum ;
    ORDER BY 3, 10

* BROWSE WINDOW brows

* INTO CURSOR expdocum ;
* INTO TABLE (put_dbf) + ('expdocum.dbf') ;

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

imavisadoc = ('vissum.dbf')

COPY TO (put_exp) + (imavisadoc) FOX2X AS 866 && Выгрузка сформированных и свернутых по транзитным счетам в RsBank

IF sergeizhuk_flag = .F.
  COPY TO (put_rsbank) + (imavisadoc) FOX2X AS 866 && Выгрузка сформированных и свернутых по транзитным счетам в RsBank
ENDIF

RETURN


******************************************************** PROCEDURE export_data_proz_dt ******************************************************************

PROCEDURE export_data_proz_dt

SELECT acc_cln_dt_pol
SET ORDER TO card_nls
GO TOP

SELECT acc_cln_dt_null
SET ORDER TO fio
GO TOP

* ----------------------------------------------------------------- Начинаем сканировать таблицу процентов расчитанную по датам ------------------------------------------------------------------------------ *

SCAN

  WAIT WINDOW 'Клиент - ' + ALLTRIM(acc_cln_dt_null.fio_rus) + '  Счет клиента - ' + ALLTRIM(acc_cln_dt_null.card_nls) NOWAIT

  rec = ALLTRIM(acc_cln_dt_null.card_nls)  +  DTOS(acc_cln_dt_null.data_home)  +  DTOS(acc_cln_dt_null.data_end)

  SELECT acc_cln_dt_pol

  poisk = SEEK(rec)

  IF poisk = .T.
    SCATTER MEMVAR MEMO
  ELSE
    SCATTER MEMVAR MEMO BLANK
  ENDIF

  m.branch = ALLTRIM(acc_cln_dt_null.branch)
  m.ref_client = ALLTRIM(acc_cln_dt_null.ref_client)
  m.ref_card = acc_cln_dt_null.ref_card

  m.ref_ost = acc_cln_dt_null.ref_ost
  m.num_dog = acc_cln_dt_null.num_dog  && Номер договора
  m.data_dog = acc_cln_dt_null.data_dog  && Дата начала заключения договора

  m.fio_rus = ALLTRIM(acc_cln_dt_null.fio_rus)  && ФИО клиента
  m.card_nls = ALLTRIM(acc_cln_dt_null.card_nls)  && Номер карточного счета
  m.val_card = ALLTRIM(acc_cln_dt_null.val_card)  && Тип карты
  m.pr = acc_cln_dt_null.pr

  m.data_proz = acc_cln_dt_null.data_proz  && Дата начала начисления процентов
  m.dt_hom_end = acc_cln_dt_null.dt_hom_end  && Дата начала квартала за который происходит расчет
  m.data_home = acc_cln_dt_null.data_home  && Начальная дата диапазона расчета, вводится оператором
  m.data_end = acc_cln_dt_null.data_end  && Конечная дата диапазона расчета, вводится оператором
  m.data_pak = acc_cln_dt_null.data_pak  && Конечная дата диапазона расчета при пакетном расчете процентов, рассчитыается автоматически

  m.colvo_den = acc_cln_dt_null.colvo_den

  m.kvartal = acc_cln_dt_null.kvartal  && Номер квартала
  m.mes_kvart = acc_cln_dt_null.mes_kvart  && Номер месяца в квартале

  m.kurs = acc_cln_dt_null.kurs  && Курс валюты

  m.stavka = acc_cln_dt_null.stavka  && Процентная ставка
  m.sum_ostat = acc_cln_dt_null.sum_ostat  && Сумма остатка по счету

  m.sum_dobav = acc_cln_dt_null.sum_dobav  && Сумма расчетная
  m.sum_rasch = acc_cln_dt_null.sum_rasch  && Сумма расчетная
  m.sum_rasxod = acc_cln_dt_null.sum_rasxod
  m.sum_nalog = acc_cln_dt_null.sum_nalog  && Сумма налога
  m.sum_viplat = acc_cln_dt_null.sum_viplat  && Сумма к выплате

  IF pr_time_rachet_proz = .T.
    m.time_den = acc_cln_dt_null.time_den
  ENDIF

  m.data = acc_cln_dt_null.data

  m.ima_pk = ALLTRIM(acc_cln_dt_null.ima_pk)

  IF poisk = .T.
    GATHER MEMVAR MEMO
  ELSE
    INSERT INTO acc_cln_dt_pol FROM MEMVAR
  ENDIF

  SELECT acc_cln_dt_null
ENDSCAN

RETURN


******************************************************* PROCEDURE export_data_proz_sum ******************************************************************

PROCEDURE export_data_proz_sum

SELECT acc_cln_sum_pol
SET ORDER TO card_nls
GO TOP

SELECT acc_cln_sum_null
SET ORDER TO fio
GO TOP

* ----------------------------------------------------------------- Начинаем сканировать таблицу процентов сверную в итоговую ---------------------------------------------------------------------------------- *

SCAN

  WAIT WINDOW 'Клиент - ' + ALLTRIM(acc_cln_sum_null.fio_rus) + '  Счет клиента - ' + ALLTRIM(acc_cln_sum_null.card_nls) NOWAIT

  rec = ALLTRIM(acc_cln_sum_null.card_nls)  +  DTOS(acc_cln_sum_null.data_home) + DTOS(acc_cln_sum_null.data_end)

  SELECT acc_cln_sum_pol

  poisk = SEEK(rec)

  IF poisk = .T.
    SCATTER MEMVAR MEMO
  ELSE
    SCATTER MEMVAR MEMO BLANK
  ENDIF

  m.branch = ALLTRIM(acc_cln_sum_null.branch)
  m.ref_client = ALLTRIM(acc_cln_sum_null.ref_client)
  m.ref_card = acc_cln_sum_null.ref_card

  m.ref_ost = acc_cln_sum_null.ref_ost

  m.num_dog = acc_cln_sum_null.num_dog  && Номер договора
  m.data_dog = acc_cln_sum_null.data_dog  && Дата начала заключения договора

  m.fio_rus = ALLTRIM(acc_cln_sum_null.fio_rus)  && ФИО клиента

  m.card_nls = ALLTRIM(acc_cln_sum_null.card_nls)  && Номер карточного счета

  m.val_card = ALLTRIM(acc_cln_sum_null.val_card)  && Тип карты

  m.pr = acc_cln_sum_null.pr

  m.data_proz = acc_cln_sum_null.data_proz  && Дата начала начисления процентов
  m.dt_hom_end = acc_cln_sum_null.dt_hom_end  && Дата начала квартала за который происходит расчет
  m.data_home = acc_cln_sum_null.data_home  && Начальная дата диапазона расчета, вводится оператором
  m.data_end = acc_cln_sum_null.data_end  && Конечная дата диапазона расчета, вводится оператором
  m.data_pak = acc_cln_sum_null.data_pak  && Конечная дата диапазона расчета при пакетном расчете процентов, рассчитыается автоматически

  m.den_mes = acc_cln_sum_null.den_mes

  m.kvartal = acc_cln_sum_null.kvartal  && Номер квартала
  m.mes_kvart = acc_cln_sum_null.mes_kvart  && Номер месяца в квартале

  m.kurs = acc_cln_sum_null.kurs  && Курс валюты

  m.stavka = acc_cln_sum_null.stavka  && Процентная ставка
  m.sum_ostat = acc_cln_sum_null.sum_ostat  && Сумма остатка по счету

  m.sum_dobav = acc_cln_sum_null.sum_dobav
  m.sum_rasch = acc_cln_sum_null.sum_rasch
  m.sum_rasxod = acc_cln_sum_null.sum_rasxod
  m.sum_nalog = acc_cln_sum_null.sum_nalog
  m.sum_viplat = acc_cln_sum_null.sum_viplat

  IF pr_time_rachet_proz = .T.
    m.time_svert = acc_cln_sum_null.time_svert
  ENDIF

  m.data = acc_cln_sum_null.data

  m.ima_pk = ALLTRIM(acc_cln_sum_null.ima_pk)

  IF poisk = .T.
    GATHER MEMVAR MEMO
  ELSE
    INSERT INTO acc_cln_sum_pol FROM MEMVAR
  ENDIF

  SELECT acc_cln_sum_null
ENDSCAN

RETURN


*********************************************************** PROCEDURE export_del_proz *******************************************************************

PROCEDURE export_del_proz

=soob('Внимание! Выполнение данной процедуры требует монопольного режима работы с таблицами ... ')

tex1 = 'В ы п о л н и т ь  удаление данных во временных таблицах'
tex2 = 'О т к а з а т ь с я  от удаления данных во временных таблицах'
l_bar = 3
=popup_9(tex1,tex2,tex3,tex4,l_bar)
ACTIVATE POPUP vibor
RELEASE POPUPS vibor
IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

  IF BAR() = 1

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начата очистка рассчитанных процентов во временных таблицах ... ',WCOLS())
    =INKEY(2)

    tim2 = SECONDS()

    SELECT branch, ref_client, fio_rus, card_nls ;
      FROM acc_cln_dt_null ;
      WHERE BETWEEN(data_home, vvod_data_home, vvod_data_end) ;
      INTO CURSOR sel_prozent ;
      ORDER BY 1, 2

    IF _TALLY <> 0

      SELECT acc_cln_dt_null
      USE
      USE (put_dbf) + ('acc_cln_dt_null.dbf') ALIAS acc_cln_dt_null EXCLUSIVE
      DELETE FOR BETWEEN(data_home, vvod_data_home, vvod_data_end)
      PACK
      USE
      USE (put_dbf) + ('acc_cln_dt_null.dbf') ALIAS acc_cln_dt_null SHARED
      SET ORDER TO card_nls

      =soob('Данные в диапазоне ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) + ' из таблицы с расчетом по дням удалены ...')

      SELECT branch, ref_client, fio_rus, card_nls ;
        FROM acc_cln_sum_null ;
        WHERE BETWEEN(data_home, vvod_data_home, vvod_data_end) ;
        INTO CURSOR sel_prozent ;
        ORDER BY 1, 2

      IF _TALLY <> 0

        SELECT acc_cln_sum_null
        USE
        USE (put_dbf) + ('acc_cln_sum_null.dbf') ALIAS acc_cln_sum_null EXCLUSIVE
        DELETE FOR BETWEEN(data_home, vvod_data_home, vvod_data_end)
        PACK
        USE
        USE (put_dbf) + ('acc_cln_sum_null.dbf') ALIAS acc_cln_sum_null SHARED
        SET ORDER TO card_nls

        =soob('Данные в диапазоне ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) + ' из таблицы с итоговым расчетом удалены ...')

      ELSE
        =err('Внимание! За выбранные даты данных не обнаружено ...')
      ENDIF

    ELSE
      =err('Внимание! За выбранные даты данных не обнаружено ...')
    ENDIF

    tim3 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Очистка временных таблиц успешно завершена' + ' (время = ' + ALLTRIM(STR(tim3 - tim2, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ENDIF
ENDIF

RETURN


*********************************************************************************************************************************************************






































































