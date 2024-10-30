**************************************************
*** Формирование справочника карточных счетов ***
**************************************************

******************************************************************* PROCEDURE start ********************************************************************

PROCEDURE start
sel = SELECT()
popup_ima=LOWER(POPUP())
prompt_ima=LOWER(PROMPT())
bar_num=BAR()
HIDE POPUP (popup_ima)

tex1 = 'Выгрузка данных по ЗАРПЛАТНЫМ балансовым карточным счетам 40817 и 40820'
tex2 = 'Выгрузка данных по КЛИЕНТСКИМ балансовым карточным счетам 40817 и 40820'
l_bar = 3
=popup_9(tex1, tex2, tex3, tex4, l_bar)
ACTIVATE POPUP vibor
RELEASE POPUPS vibor

IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

  put_select = BAR()

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  DO CASE
    CASE put_select = 1  && Выгрузка данных по ЗАРПЛАТНЫМ балансовым карточным счетам 40817 и 40820

      =soob('Внимание! Выберите номер проекта, в который будет перемещен карточный счет клиента ... ')

      ACTIVATE POPUP zarpl_proekt

      IF NOT LASTKEY() = 27

        IF INLIST(LASTKEY(), 4, 19)
          LOOP
        ELSE
          DO zarpl_proekt IN Visa
        ENDIF

      ELSE

        num_proekt = '00'  && Номер зарплатного проекта

        =soob('Внимание! Ваш выбор - Главный справочник карточных счетов филиала - код проекта "00"')

      ENDIF

      SELECT DISTINCT branch, num_proekt AS kod_proekt, client_b, fio_rus,;
        LEFT(ALLTRIM(card_acct), 8) + '0' + '4' + SUBSTR(ALLTRIM(card_acct), 11, 1) + ALLTRIM(num_branch) + ALLTRIM(num_proekt) + RIGHT(ALLTRIM(card_acct), 5) AS card_nls, card_acct,;
        type_card, val_card, tran_42301, name_card ;
        FROM account ;
        WHERE INLIST(LEFT(ALLTRIM(card_acct), 5), '42301', '42601', '40817', '40820') AND INLIST(ALLTRIM(type_card), '003', '006', '008') AND ;
        SUBSTR(ALLTRIM(card_acct), 10, 1) <> '4' AND ALLTRIM(val_card) == 'RUR' ;
        UNION ;
        SELECT DISTINCT branch, num_proekt AS kod_proekt, client_b, fio_rus,;
        LEFT(ALLTRIM(card_acct), 8) + '0' + '4' + SUBSTR(ALLTRIM(card_acct), 11, 1) + ALLTRIM(num_branch) + ALLTRIM(num_proekt) + RIGHT(ALLTRIM(card_acct), 5) AS card_nls, card_acct,;
        type_card, val_card, tran_42301, name_card ;
        FROM acc_del ;
        WHERE INLIST(LEFT(ALLTRIM(card_acct), 5), '42301', '42601', '40817', '40820') AND INLIST(ALLTRIM(type_card), '003', '006', '008') AND ;
        SUBSTR(ALLTRIM(card_acct), 10, 1) <> '4' AND ALLTRIM(val_card) == 'RUR' ;
        INTO CURSOR konvert READWRITE ;
        ORDER BY 1, 2, 6, 7, 3

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE put_select = 3  && Выгрузка данных по КЛИЕНТСКИМ балансовым карточным счетам 40817 и 40820

      SELECT DISTINCT branch, kod_proekt, client_b, fio_rus,;
        SUBSTR(ALLTRIM(card_acct), 1, 8) + '0' + '5' + SUBSTR(ALLTRIM(card_acct), 11, 10) AS card_nls, card_acct,;
        type_card, val_card, tran_42301, name_card ;
        FROM account ;
        WHERE INLIST(LEFT(ALLTRIM(card_acct), 5), '40817', '40820') AND INLIST(ALLTRIM(type_card), '001', '005', '007', '012', '014', '015', '020') ;
        UNION ;
        SELECT DISTINCT branch, kod_proekt, client_b, fio_rus,;
        SUBSTR(ALLTRIM(card_acct), 1, 8) + '0' + '5' + SUBSTR(ALLTRIM(card_acct), 11, 10) AS card_nls, card_acct,;
        type_card, val_card, tran_42301, name_card ;
        FROM acc_del ;
        WHERE INLIST(LEFT(ALLTRIM(card_acct), 5), '40817', '40820') AND INLIST(ALLTRIM(type_card), '001', '005', '007', '012', '014', '015', '020') ;
        INTO CURSOR konvert READWRITE ;
        ORDER BY 1, 2, 6, 7, 3

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  ENDCASE

  colvo_zap = _TALLY

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF colvo_zap <> 0

    SELECT konvert
    GO TOP

    SCAN

      sel_card_nls = ALLTRIM(konvert.card_nls)
      sel_card_acct = ALLTRIM(konvert.card_acct)

      zSHET = ALLTRIM(sel_card_nls)

      DO R_key IN Rasch_kl WITH zUCH, zSHET && В данной процедуре происходит расчет ключа в лицевом счете, номер позиции 9.

      sel_card_nls = ALLTRIM(kSHET)

      REPLACE card_nls WITH sel_card_nls

      SELECT konvert
    ENDSCAN

    SELECT konvert
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    BROWSE FIELDS ;
      branch:H = 'Филиал',;
      kod_proekt:H = 'Проект',;
      client_b:H = '№ клиента',;
      fio_rus:H = 'Фамилия Имя Отчество',;
      card_nls:H = 'Картсчет в МАК Visa',;
      card_acct:H = 'Картсчет в УПК',;
      type_card:H = 'Тип',;
      val_card:H = 'Валюта',;
      tran_42301:H = 'Транзитный счет',;
      name_card:H = 'Тип карты' ;
      WINDOWS brows ;
      TITLE ' Справочные данные по карточным счетам. Количество найденных счетов составляет - ' + ALLTRIM(STR(colvo_zap)) + ' штук.'

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начато формирование текстового файла для конвертации счетов.',WCOLS())

    SELECT konvert
    GO TOP

    SET TEXTMERGE ON
    SET TEXTMERGE NOSHOW

    _text = FCREATE((put_tmp) + ('konvert.txt'))

    IF _text < 1

      =err('Внимание! Ошибка при создании файла ' + (put_tmp) + ('konvert.txt'))

      =FCLOSE(_text)

    ELSE

      SELECT konvert
      GO TOP

      SCAN

        sel_card_nls = ALLTRIM(konvert.card_nls)
        sel_card_acct = ALLTRIM(konvert.card_acct)

        stroka = ALLTRIM(sel_card_acct) + SPACE(3) + ALLTRIM(sel_card_nls)

        =FPUTS(_text, (stroka))

        SELECT konvert
      ENDSCAN

      =FCLOSE(_text)

    ENDIF

    SET TEXTMERGE OFF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
    COPY TO (put_dbf) + ('konvert.dbf') AS 1251 FIELDS branch, kod_proekt, client_b, fio_rus, card_nls, card_acct, type_card, val_card, tran_42301, name_card
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    IF USED('konvert')
      SELECT konvert
      USE
    ENDIF

    IF FILE((put_dbf) + ('konvert.dbf'))

      IF NOT USED('new_konvert')
        SELECT 0
        USE (put_dbf) + ('konvert.dbf') ALIAS new_konvert
      ENDIF

      SELECT new_konvert
      GO TOP

      SCAN

        sel_card_nls = ALLTRIM(card_nls)

        zSHET = ALLTRIM(sel_card_nls)

        DO R_key IN Rasch_kl WITH zUCH, zSHET && В данной процедуре происходит расчет ключа в лицевом счете, номер позиции 9.

        REPLACE card_nls WITH ALLTRIM(kSHET)

        SELECT new_konvert
      ENDSCAN

      IF USED('new_konvert')
        SELECT new_konvert
        USE
      ENDIF

    ELSE
      =err('Необходимый для работы файл по пути ' + (put_dbf) + ('konvert.dbf') + ' не найден ... ')
    ENDIF

    @ WROWS()/3,3 SAY PADC('Файл ' + (put_dbf) + ('konvert.dbf') + ' успешно записан ... ',WCOLS())
    =INKEY(4)

    @ WROWS()/3,3 SAY PADC('Файл ' + (put_tmp) + ('konvert.txt') + ' успешно записан ... ',WCOLS())
    =INKEY(4)

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Формирование текстового файла для конвертации счетов успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ENDIF

RETURN


***********************************************************************************************************************************************************


