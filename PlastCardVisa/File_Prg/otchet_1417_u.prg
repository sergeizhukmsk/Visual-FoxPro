**************************************************
*** Данные о ежедневных остатках - форма 1417 У ***
**************************************************

****************************************************************** PROCEDURE start *********************************************************************

PROCEDURE start
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

DO formirov_menu IN Otchet_1417_U

ACTIVATE POPUP otchet_1417_U

RELEASE MENUS otchet_1417_U

SELECT(sel)
RETURN


*************************************************************** PROCEDURE select_account ****************************************************************

PROCEDURE select_account
PARAMETERS filtr_data

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, начата выборка данных ... ',WCOLS())

SELECT client
COUNT TO colvo_client

SELECT acc_del
COUNT TO colvo_acc_del

SELECT account
COUNT TO colvo_account

SELECT client

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF colvo_client <> 0

  SELECT DISTINCT A.ref_client, A.client_b, A.fio_rus,;
    IIF(EMPTY(ALLTRIM(A.kladr_dom)) = .F., PADR(ALLTRIM(A.kladr_dom), 250), SPACE(250)) AS adres_dom,;
    IIF(EMPTY(ALLTRIM(A.kladr_reg)) = .F., PADR(ALLTRIM(A.kladr_reg), 250), SPACE(250)) AS adres_reg,;
    A.telef_dom, A.e_mail AS e_mail,;
    A.doc_type, A.pass_seria, A.pass_num, A.pass_data,;
    IIF(EMPTY(ALLTRIM(A.pass_mesto)) = .F., PADR(ALLTRIM(A.pass_mesto), 160), SPACE(160)) AS pass_mesto, A.birth_date,;
    A.risk, A.date_risk, A.dat_reestr, A.date_read, A.pr_rekviz,;
    B.name_P AS name_p ;
    FROM client A, sprfirma B ;
    WHERE ALLTRIM(A.branch) == ALLTRIM(B.branch) ;
    INTO CURSOR  sel_client ;
    ORDER BY fio_rus

  colvo_client = _TALLY

* BROWSE WINDOW brows TITLE 'Количество клиентов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

ELSE
  RETURN
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF colvo_account <> 0

  SELECT DISTINCT A.branch, A.ref_client, A.fio_rus, A.card_nls, A.card_acct,;
    A.num_dog, A.data_dog, A.card_nls AS num_liz_d,;
    A.begin_bal, A.vxd_ost_m, B.name_p, A.val_card, A.name_card,;
    B.adres_reg, B.adres_dom, B.telef_dom, B.e_mail,;
    B.doc_type, B.pass_seria, B.pass_num, B.pass_data, B.pass_mesto, B.birth_date,;
    B.risk, B.date_risk, B.dat_reestr, B.date_read, B.pr_rekviz ;
    FROM account A, sel_client B ;
    WHERE ALLTRIM(A.ref_client) == ALLTRIM(B.ref_client) AND A.del_card <> 2 AND ALLTRIM(A.vid_card) == '1' AND ;
    A.date_nls < filtr_data AND ALLTRIM(A.status_tra) <> '4' AND NOT INLIST(LEFT(ALLTRIM(A.card_acct), 5), '40702', '40802') ;
    INTO CURSOR sel_account_O ;
    ORDER BY 3

  colvo_account = _TALLY

* BROWSE WINDOW brows TITLE 'Количество открытых счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

ELSE
  RETURN
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF colvo_acc_del <> 0

  SELECT DISTINCT A.branch, A.ref_client, A.fio_rus, A.card_nls, A.card_acct,;
    A.num_dog, A.data_dog, A.card_nls AS num_liz_d,;
    A.begin_bal, A.vxd_ost_m, B.name_p, A.val_card, A.name_card,;
    B.adres_reg, B.adres_dom, B.telef_dom, B.e_mail,;
    B.doc_type, B.pass_seria, B.pass_num, B.pass_data, B.pass_mesto, B.birth_date,;
    B.risk, B.date_risk, B.dat_reestr, B.date_read, B.pr_rekviz ;
    FROM acc_del A, sel_client B ;
    WHERE ALLTRIM(A.ref_client) == ALLTRIM(B.ref_client) AND A.del_card = 2 AND ALLTRIM(A.vid_card) == '1' AND ;
    A.date_nls < filtr_data AND ALLTRIM(A.status_tra) <> '4' AND NOT INLIST(LEFT(ALLTRIM(A.card_acct), 5), '40702', '40802') ;
    INTO CURSOR sel_account_Z ;
    ORDER BY 3

  colvo_acc_del = _TALLY

* BROWSE WINDOW brows TITLE 'Количество закрытых счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO CASE
  CASE colvo_account <> 0 AND colvo_acc_del <> 0 AND colvo_client <> 0

    SELECT DISTINCT branch, ref_client, fio_rus, card_nls, card_acct,;
      num_dog, data_dog, num_liz_d,;
      begin_bal, vxd_ost_m, name_p, val_card, name_card,;
      adres_reg, adres_dom, telef_dom, e_mail,;
      doc_type, pass_seria, pass_num, pass_data, pass_mesto, birth_date,;
      risk, date_risk, dat_reestr, date_read, pr_rekviz ;
      FROM sel_account_O ;
      UNION ;
      SELECT DISTINCT branch, ref_client, fio_rus, card_nls, card_acct,;
      num_dog, data_dog, num_liz_d,;
      begin_bal, vxd_ost_m, name_p, val_card, name_card,;
      adres_reg, adres_dom, telef_dom, e_mail,;
      doc_type, pass_seria, pass_num, pass_data, pass_mesto, birth_date,;
      risk, date_risk, dat_reestr, date_read, pr_rekviz ;
      FROM sel_account_Z ;
      INTO CURSOR sel_account_pol ;
      ORDER BY 3

    colvo_zap = _TALLY

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

  CASE colvo_account <> 0 AND colvo_client <> 0

    SELECT DISTINCT branch, ref_client, fio_rus, card_nls, card_acct,;
      num_dog, data_dog, num_liz_d,;
      begin_bal, vxd_ost_m, name_p, val_card, name_card,;
      adres_reg, adres_dom, telef_dom, e_mail,;
      doc_type, pass_seria, pass_num, pass_data, pass_mesto, birth_date,;
      risk, date_risk, dat_reestr, date_read, pr_rekviz ;
      FROM sel_account_O ;
      INTO CURSOR sel_account_pol ;
      ORDER BY 3

    colvo_zap = _TALLY

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF colvo_zap <> 0

  SELECT DISTINCT branch, ref_client, fio_rus, card_nls, card_acct,;
    num_dog, data_dog, num_liz_d,;
    begin_bal, vxd_ost_m, name_p, val_card, name_card,;
    adres_reg, adres_dom, telef_dom, e_mail,;
    doc_type, pass_seria, pass_num, pass_data, pass_mesto, birth_date,;
    risk, date_risk, dat_reestr, date_read, pr_rekviz ;
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

  =err('Внимание! Данных для реестра обязательств не обнаружено ...')
ENDIF

RETURN


************************************************************ PROCEDURE forma_1417_U_den ****************************************************************

PROCEDURE forma_1417_U_den

STORE new_data_odb TO dt1

colvo_zap = 0

DO poisk_den_kurs_avt IN Servis WITH dt1 - 1, 'USD'
DO poisk_den_kurs_avt IN Servis WITH dt1 - 1, 'EUR'

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

      SELECT DISTINCT branch, new_data_odb AS date_vedom, ref_client, fio_rus, birth_date, card_nls, card_acct, adres_reg, adres_dom, telef_dom, e_mail,;
        doc_type, pass_seria, pass_num, pass_data, pass_mesto,;
        num_dog, data_dog, num_liz_d, name_p,;
        begin_bal AS vxd_ostat_val,;
        IIF(ALLTRIM(val_card) == 'RUR', ROUND(1.00 * 1.00, 4),;
        IIF(ALLTRIM(val_card) == 'USD', ROUND(poisk_kurs_usd * 1.00, 4),;
        IIF(ALLTRIM(val_card) == 'EUR', ROUND(poisk_kurs_eur * 1.00, 4), 000.0000))) AS kurs,;
        IIF(ALLTRIM(val_card) == 'RUR', begin_bal,;
        IIF(ALLTRIM(val_card) == 'USD', ROUND(begin_bal * poisk_kurs_usd, 2 ),;
        IIF(ALLTRIM(val_card) == 'EUR', ROUND(begin_bal * poisk_kurs_eur, 2 ), 000000.00))) AS vxd_ostat_rur,;
        risk, date_risk, dat_reestr, date_read, pr_rekviz ;
        FROM sel_account_O ;
        INTO CURSOR sel_reestr ;
        ORDER BY branch, fio_rus, birth_date

    CASE BAR() = 3  && Вы желаете использовать ведомость остатков из ПО "CARD-VISA филиал"

      SELECT DISTINCT branch, new_data_odb AS date_vedom, ref_client, fio_rus, birth_date, card_nls, card_acct, adres_reg, adres_dom, telef_dom, e_mail,;
        doc_type, pass_seria, pass_num, pass_data, pass_mesto,;
        num_dog, data_dog, num_liz_d, name_p,;
        vxd_ost_m AS vxd_ostat_val,;
        IIF(ALLTRIM(val_card) == 'RUR', ROUND(1.00 * 1.00, 4),;
        IIF(ALLTRIM(val_card) == 'USD', ROUND(poisk_kurs_usd * 1.00, 4),;
        IIF(ALLTRIM(val_card) == 'EUR', ROUND(poisk_kurs_eur * 1.00, 4), 000.0000))) AS kurs,;
        IIF(ALLTRIM(val_card) == 'RUR', vxd_ost_m,;
        IIF(ALLTRIM(val_card) == 'USD', ROUND(vxd_ost_m * poisk_kurs_usd, 2 ),;
        IIF(ALLTRIM(val_card) == 'EUR', ROUND(vxd_ost_m * poisk_kurs_eur, 2 ), 000000.00))) AS vxd_ostat_rur,;
        risk, date_risk, dat_reestr, date_read, pr_rekviz ;
        FROM sel_account_O ;
        INTO CURSOR sel_reestr ;
        ORDER BY branch, fio_rus, birth_date

  ENDCASE

  colvo_zap = _TALLY

  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

ENDIF

RETURN


********************************************************* PROCEDURE forma_1417_U_arx *******************************************************************

PROCEDURE forma_1417_U_arx

STORE arx_data_odb TO dt1

colvo_zap = 0

SELECT DISTINCT data_home AS data ;
  FROM calendar ;
  WHERE DTOS(data_home) < DTOS(new_data_odb) ;
  INTO CURSOR tov ;
  ORDER BY 1 DESC

IF _TALLY <> 0

  =soob('Из списка дат нужно выбрать дату НА КОТОРУЮ НЕОБХОДИМО ВЫГРУЗИТЬ РЕЕСТР')

  DO FORM (put_scr) + ('seldata1.scx')

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    DO poisk_den_kurs_avt IN Servis WITH dt1 - 1, 'USD'
    DO poisk_den_kurs_avt IN Servis WITH dt1 - 1, 'EUR'

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

          SELECT DISTINCT A.branch, B.date_vedom, A.ref_client, A.fio_rus, A.birth_date, A.card_nls, A.card_acct,;
            A.adres_reg, A.adres_dom, A.telef_dom, A.e_mail,;
            A.doc_type, A.pass_seria, A.pass_num, A.pass_data, A.pass_mesto,;
            A.num_dog, A.data_dog, A.num_liz_d, A.name_p,;
            B.begin_bal AS vxd_ostat_val,;
            IIF(ALLTRIM(B.val_card) == 'RUR', ROUND(1.00 * 1.00, 4),;
            IIF(ALLTRIM(B.val_card) == 'USD', ROUND(poisk_kurs_usd * 1.00, 4),;
            IIF(ALLTRIM(B.val_card) == 'EUR', ROUND(poisk_kurs_eur * 1.00, 4), 000.0000))) AS kurs,;
            IIF(ALLTRIM(B.val_card) == 'RUR', B.begin_bal,;
            IIF(ALLTRIM(B.val_card) == 'USD', ROUND(B.begin_bal * poisk_kurs_usd, 2 ),;
            IIF(ALLTRIM(B.val_card) == 'EUR', ROUND(B.begin_bal * poisk_kurs_eur, 2 ), 000000.00))) AS vxd_ostat_rur,;
            A.risk, A.date_risk, A.dat_reestr, A.date_read, A.pr_rekviz ;
            FROM sel_account A, vedom_ost B ;
            WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_acct) AND BETWEEN(DTOS(B.date_vedom), DTOS(dt1), DTOS(dt1)) ;
            INTO CURSOR sel_reestr ;
            ORDER BY A.branch, A.fio_rus, A.birth_date

        CASE BAR() = 3  && Вы желаете использовать ведомость остатков из ПО "CARD-VISA филиал"

          SELECT DISTINCT A.branch, B.date_vedom, A.ref_client, A.fio_rus, A.birth_date, A.card_nls, A.card_acct,;
            A.adres_reg, A.adres_dom, A.telef_dom, A.e_mail,;
            A.doc_type, A.pass_seria, A.pass_num, A.pass_data, A.pass_mesto,;
            A.num_dog, A.data_dog, A.num_liz_d, A.name_p,;
            B.vxd_ost_m AS vxd_ostat_val,;
            IIF(ALLTRIM(B.val_card) == 'RUR', ROUND(1.00 * 1.00, 4),;
            IIF(ALLTRIM(B.val_card) == 'USD', ROUND(poisk_kurs_usd * 1.00, 4),;
            IIF(ALLTRIM(B.val_card) == 'EUR', ROUND(poisk_kurs_eur * 1.00, 4), 000.0000))) AS kurs,;
            IIF(ALLTRIM(B.val_card) == 'RUR', B.vxd_ost_m,;
            IIF(ALLTRIM(B.val_card) == 'USD', ROUND(B.vxd_ost_m * poisk_kurs_usd, 2 ),;
            IIF(ALLTRIM(B.val_card) == 'EUR', ROUND(B.vxd_ost_m * poisk_kurs_eur, 2 ), 000000.00))) AS vxd_ostat_rur,;
            A.risk, A.date_risk, A.dat_reestr, A.date_read, A.pr_rekviz ;
            FROM sel_account A, vedom_ost B ;
            WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_acct) AND BETWEEN(DTOS(B.date_vedom), DTOS(dt1), DTOS(dt1)) ;
            INTO CURSOR sel_reestr ;
            ORDER BY A.branch, A.fio_rus, A.birth_date

      ENDCASE

      colvo_zap = _TALLY

      =INKEY(2)
      DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows TITLE 'Количество всех счетов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

    ENDIF

  ENDIF

ELSE

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')

ENDIF

RETURN


********************************************************** PROCEDURE brows_zajava_b_adres_dom **********************************************************

PROCEDURE brows_zajava_b_adres_dom

BROWSE FIELDS ;
  kod_proekt:H = 'Проект: ',;
  fio_rus:H = 'Фамилия Имя Отчество: ',;
  pr_rekviz:H = 'Признак наличия',;
  adres_dom:H = 'Полный адрес ДОМАШНИЙ: ',;
  ind_dom:H = 'Индекс дом.: ',;
  region_dom:H = 'Регион дом.: ',;
  rajon_dom:H = 'Район дом.: ',;
  city_dom:H = 'Город дом.: ',;
  punkt_dom:H = 'Населен. пункт дом.: ',;
  street_dom:H = 'Улица дом.: ',;
  dom_dom:H = 'Дом дом.: ',;
  korpus_dom:H = 'Корпус дом.: ',;
  stroen_dom:H = 'Строение дом.: ',;
  kvart_dom:H = 'Квартира дом.: ',;
  pass_type:H = 'Вид паспорта: ',;
  pass_seria:H = 'Серия паспорта: ',;
  pass_num:H = 'Номер паспорта: ',;
  pass_data:H = 'Дата выдачи: ',;
  pass_mesto:H = 'Место выдачи: ' ;
  WINDOW brows ;
  TITLE ' Просомтр адресов в справочнике заявлений на соответствие формату КЛАДР'

RETURN


********************************************************** PROCEDURE brows_zajava_b_adres_reg **********************************************************

PROCEDURE brows_zajava_b_adres_reg

BROWSE FIELDS ;
  kod_proekt:H = 'Проект: ',;
  fio_rus:H = 'Фамилия Имя Отчество: ',;
  pr_rekviz:H = 'Признак наличия',;
  adres_reg:H = 'Полный адрес РЕГИСТРАЦИИ: ',;
  ind_reg:H = 'Индекс рег.: ',;
  region_reg:H = 'Регион рег.: ',;
  rajon_reg:H = 'Район рег.: ',;
  city_reg:H = 'Город рег.: ',;
  punkt_reg:H = 'Населен. пункт рег.: ',;
  street_reg:H = 'Улица рег.: ',;
  dom_reg:H = 'Дом рег.: ',;
  korpus_reg:H = 'Корпус рег.: ',;
  stroen_reg:H = 'Строения рег.: ',;
  kvart_reg:H = 'Квартира рег.: ',;
  pass_type:H = 'Вид паспорта: ',;
  pass_seria:H = 'Серия паспорта: ',;
  pass_num:H = 'Номер паспорта: ',;
  pass_data:H = 'Дата выдачи: ',;
  pass_mesto:H = 'Место выдачи: ' ;
  WINDOW brows ;
  TITLE ' Просомтр адресов в справочнике заявлений на соответствие формату КЛАДР'

RETURN


************************************************************** PROCEDURE start_reestr *******************************************************************

PROCEDURE start_reestr
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

pr_start_reestr = .F.

DO CASE
  CASE ALLTRIM(num_branch) == '03'

*    DO new_adres IN Adres_client_03  && Формируем адресные поля по формату КЛАДР и паспортные поля в справочнике заявлений

    DO red_client IN Adres_client_03  && Формируем адресные поля по формату КЛАДР и паспортные поля в справочнике клиентов

  CASE ALLTRIM(num_branch) <> '03'

    DO scan_zajava_b IN Adres_client  && Формируем адресные поля по формату КЛАДР и паспортные поля в справочнике заявлений и справочнике клиентов

ENDCASE

=soob('Внимание! Запущена процедура проверки количества запятых в адресных полях ... ')

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SELECT branch, kod_proekt, client_b, fio_rus, pr_rekviz,;
  ind_dom, ind_reg, region_dom, region_reg, rajon_dom, rajon_reg, city_dom, city_reg, punkt_dom, punkt_reg, street_dom, street_reg,;
  dom_dom, dom_reg, korpus_dom, korpus_reg, stroen_dom, stroen_reg, kvart_dom, kvart_reg, adres_dom, kladr_dom,;
  pass_type, pass_seria, pass_num, pass_data, pass_mesto ;
  FROM zajava_b ;
  WHERE OCCURS(',', adres_dom) <> 9 ;
  INTO CURSOR sel_zajava_b ;
  ORDER BY fio_rus

colvo_zap_dom =_TALLY

* BROWSE WINDOW brows

IF colvo_zap_dom <> 0

  =err('Внимание! Обнаружены записи с неверным количеством запятых в адресе ДОМАШНЕМ ... ')

  DO brows_zajava_b_adres_dom IN Otchet_1417_u

ENDIF

SELECT branch, kod_proekt, client_b, fio_rus, pr_rekviz,;
  ind_dom, ind_reg, region_dom, region_reg, rajon_dom, rajon_reg, city_dom, city_reg, punkt_dom, punkt_reg, street_dom, street_reg,;
  dom_dom, dom_reg, korpus_dom, korpus_reg, stroen_dom, stroen_reg, kvart_dom, kvart_reg, adres_reg, kladr_reg,;
  pass_type, pass_seria, pass_num, pass_data, pass_mesto ;
  FROM zajava_b ;
  WHERE OCCURS(',', adres_reg) <> 9 ;
  INTO CURSOR sel_zajava_b ;
  ORDER BY fio_rus

colvo_zap_reg =_TALLY

* BROWSE WINDOW brows

IF colvo_zap_reg <> 0

  =err('Внимание! Обнаружены записи с неверным количеством запятых в адресе РЕГИСТРАЦИИ ... ')

  DO brows_zajava_b_adres_reg IN Otchet_1417_u

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF colvo_zap_dom <> 0 OR colvo_zap_reg <> 0

  put_prn = 1

  tex1 = 'Формирование реестра с ошибками в формате адресов ВЫПОЛНИТЬ'
  tex2 = 'Формирование реестра с ошибками в формате адресов НЕ ВЫПОЛНЯТЬ'
  l_bar = 3
  =popup_9(tex1,tex2,tex3,tex4,l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    DO CASE
      CASE BAR() = 1  && Формирование реестра с ошибками в формате адресов ВЫПОЛНИТЬ
        pr_start_reestr = .T.

      CASE BAR() = 1  && Формирование реестра с ошибками в формате адресов НЕ ВЫПОЛНЯТЬ
        pr_start_reestr = .F.

    ENDCASE

  ELSE
    pr_start_reestr = .F.
  ENDIF

ELSE
  =soob('Внимание! Записей с неверным количеством запятых в адресных полях не обнаружено ... ')
  pr_start_reestr = .T.
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF pr_start_reestr = .T.

  =soob('Внимание! Формирование реестра обязательств банка перед вкладчиками.')

  DEFINE WINDOW otchet_zb AT 0,0 SIZE colvo_rows - 2, colvo_cols - 4 FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
    NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
  MOVE WINDOW otchet_zb CENTER

  put_prn = 1

  DO WHILE .T.

    tex1 = 'Формирование отчета за текущую дату операционного дня'
    tex2 = 'Формирование отчета за архивную дату операционного дня'
    l_bar = 3
    =popup_9(tex1,tex2,tex3,tex4,l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      put_prn = BAR()

      DO CASE
        CASE BAR() = 1  && Формирование отчета за текущую дату операционного дня

          DO select_account IN Otchet_1417_U WITH new_data_odb

          IF colvo_zap <> 0

            DO forma_1417_U_den IN Otchet_1417_U

          ELSE

            =INKEY(2)
            DEACTIVATE WINDOW poisk

            =err('Внимание! Данных по реестру обязательств банка перед вкладчиками не обнаружено ...')
          ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        CASE BAR() = 3  && Формирование отчета за архивную дату операционного дня

          DO select_account IN Otchet_1417_U WITH dt1

          IF colvo_zap <> 0

            DO forma_1417_U_arx IN Otchet_1417_U

          ELSE

            =INKEY(2)
            DEACTIVATE WINDOW poisk

            =err('Внимание! Данных по реестру обязательств банка перед вкладчиками не обнаружено ...')
          ENDIF

      ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      IF colvo_zap <> 0

        =INKEY(2)
        DEACTIVATE WINDOW poisk

        put_prn = 1

        DO WHILE .T.

          tex1 = 'Произвести заполнение данными отчетной таблицы за выбранную дату'
          tex2 = 'Печатный отчет по ранее сформированным данным за выбранную дату'
          tex3 = 'Произвести удаление ранее сформированных данных за выбранную дату'
          tex4 = 'Произвести экспорт ранее сформированных данных в DOS кодировке'
          l_bar = 7
          =popup_9(tex1,tex2,tex3,tex4,l_bar)
          ACTIVATE POPUP vibor BAR put_prn
          RELEASE POPUPS vibor

          IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

            put_prn = BAR()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            DO CASE
              CASE BAR() = 1  && Произвести заполнение данными отчетной таблицы за выбранную дату

                SELECT sel_reestr
                GO TOP

                BROWSE FIELDS ;
                  date_vedom:H = 'Дата',;
                  branch:H = 'Номер филиала',;
                  ref_client:H = 'Номер вкладчика',;
                  fio_rus:H = 'Фамилия Имя Отчество',;
                  pr_rekviz:H = 'Признак наличия',;
                  birth_date:H = 'Дата рожденья',;
                  adres_reg:H = 'Адрес регистрации',;
                  adres_dom:H = 'Адрес проживания',;
                  telef_dom:H = 'Телефон для связи',;
                  e_mail:H = 'Адрес E-mail',;
                  doc_type:H = 'Тип документа',;
                  pass_seria:H = 'Серия паспорта',;
                  pass_num:H = 'Номер паспорта',;
                  pass_data:H = 'Дата выдачи',;
                  pass_mesto:H = 'Место выдачи',;
                  num_dog:H = 'Номер договора',;
                  data_dog:H = 'Дата договора',;
                  num_liz_d:H = 'Лицевой счет вкладчика',;
                  vxd_ostat_val:H = 'Вход. остаток в валюте вклада':P = '999 999 999.99',;
                  kurs:H = 'Курс валюты':P = '999.9999',;
                  vxd_ostat_rur:H = 'Вход. остаток по курсу ЦБ':P = '999 999 999.99',;
                  risk:H = 'Риск',;
                  date_risk:H = 'Дата риска',;
                  dat_reestr:H = 'Дата реестра',;
                  date_read:H = 'Дата правки' ;
                  TITLE ' Реестр обязательств банка перед вкладчиками' ;
                  WINDOW otchet_zb

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                SELECT DISTINCT SUBSTR(ALLTRIM(num_liz_d), 6, 3) AS kod_val, SUM(vxd_ostat_val) AS summa ;
                  FROM sel_reestr ;
                  INTO CURSOR sel_1417 ;
                  GROUP BY 1 ;
                  ORDER BY 1

                BROWSE FIELDS ;
                  kod_val:H = 'Код валюты',;
                  summa:H = 'С у м м а':P = '999 999 999.99' ;
                  TITLE ' Суммарные остатки на счетах по типам валют до формирования реестра 1417-У на дату ' + DT(dt1)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                =soob('Начинается заполнение данными таблицы реестра обязательств банка перед вкладчиками.')

                SELECT sel_reestr
                GO TOP

                SCAN

                  WAIT WINDOW 'Дата обработки - ' + DTOC(sel_reestr.date_vedom) + '  Клиент - ' + ALLTRIM(sel_reestr.fio_rus) NOWAIT

                  rec_zap = DTOC(sel_reestr.date_vedom) + ALLTRIM(sel_reestr.ref_client) + ALLTRIM(sel_reestr.fio_rus) + ALLTRIM(sel_reestr.num_liz_d)

                  SELECT forma_1417_u
                  SET ORDER TO ref_client
                  GO TOP

                  poisk = SEEK(rec_zap)

                  IF poisk = .T.
                    SCATTER MEMVAR MEMO
                  ELSE
                    SCATTER MEMVAR MEMO BLANK
                  ENDIF

                  m.date_vedom = sel_reestr.date_vedom

                  m.risk = sel_reestr.risk
                  m.date_risk = sel_reestr.date_risk
                  m.dat_reestr = new_data_odb
                  m.date_read = sel_reestr.date_read

                  m.branch = ALLTRIM(sel_reestr.branch)
                  m.ref_client = ALLTRIM(sel_reestr.ref_client)

                  m.fio_rus = ALLTRIM(sel_reestr.fio_rus)

                  m.fio_1 = PROPER(SUBSTR(ALLTRIM(sel_reestr.fio_rus), 1, (AT(SPACE(1), ALLTRIM(sel_reestr.fio_rus), 1) - 1)))

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                  DO CASE
                    CASE AT(SPACE(1), ALLTRIM(sel_reestr.fio_rus), 2) > 0  && Значение в поле fio_rus состоит из 3 слов

                      m.fio_2 = PROPER(SUBSTR(ALLTRIM(sel_reestr.fio_rus), (AT(SPACE(1), ALLTRIM(sel_reestr.fio_rus), 1) + 1), ;
                        (AT(SPACE(1), ALLTRIM(sel_reestr.fio_rus), 2) - 1) - (AT(SPACE(1), ALLTRIM(sel_reestr.fio_rus), 1))))

                      m.fio_3 = PROPER(SUBSTR(ALLTRIM(sel_reestr.fio_rus), (AT(SPACE(1), ALLTRIM(sel_reestr.fio_rus), 2) + 1), ;
                        (LEN(ALLTRIM(sel_reestr.fio_rus))) - (AT(SPACE(1), ALLTRIM(sel_reestr.fio_rus), 2))))

                    CASE AT(SPACE(1), ALLTRIM(sel_reestr.fio_rus), 2) = 0  && Значение в поле fio_rus состоит из 2 слов

                      m.fio_2 = PROPER(SUBSTR(ALLTRIM(sel_reestr.fio_rus), (AT(SPACE(1), ALLTRIM(sel_reestr.fio_rus), 1) + 1), ;
                        (LEN(ALLTRIM(sel_reestr.fio_rus))) - (AT(SPACE(1), ALLTRIM(sel_reestr.fio_rus), 2))))

                      m.fio_3 = ''

                  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                  m.birth_date = IIF(EMPTY(sel_reestr.birth_date) = .F., DTOC(sel_reestr.birth_date), '00.00.0000')

                  m.srok_vklad = 0.00

                  m.adres_dom = IIF(LEFT(ALLTRIM(sel_reestr.adres_dom), 1) == ',', SUBSTR(ALLTRIM(sel_reestr.adres_dom), 2, LEN(ALLTRIM(sel_reestr.adres_dom))), ALLTRIM(sel_reestr.adres_dom))
                  m.adres_dom = IIF(LEFT(ALLTRIM(m.adres_dom), 1) == ',', SUBSTR(ALLTRIM(m.adres_dom), 2, LEN(ALLTRIM(m.adres_dom))), ALLTRIM(m.adres_dom))

                  m.adres_reg = IIF(LEFT(ALLTRIM(sel_reestr.adres_reg), 1) == ',', SUBSTR(ALLTRIM(sel_reestr.adres_reg), 2, LEN(ALLTRIM(sel_reestr.adres_reg))), ALLTRIM(sel_reestr.adres_reg))
                  m.adres_reg = IIF(LEFT(ALLTRIM(m.adres_reg), 1) == ',', SUBSTR(ALLTRIM(m.adres_reg), 2, LEN(ALLTRIM(m.adres_reg))), ALLTRIM(m.adres_reg))

                  m.telef_dom = ALLTRIM(sel_reestr.telef_dom)
                  m.e_mail = ALLTRIM(sel_reestr.e_mail)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                  DO CASE
                    CASE EMPTY(ALLTRIM(sel_reestr.doc_type)) = .T.
                      m.pass_kod = '14'

                    CASE ALLTRIM(sel_reestr.doc_type) = '01'
                      m.pass_kod = '1'

                    CASE ALLTRIM(sel_reestr.doc_type) = '001'
                      m.pass_kod = '1'

                    CASE ALLTRIM(sel_reestr.doc_type) = '002'
                      m.pass_kod = '1'

                    CASE ALLTRIM(sel_reestr.doc_type) = '003'
                      m.pass_kod = '6'

                    CASE ALLTRIM(sel_reestr.doc_type) = '004'
                      m.pass_kod = '6'

                    CASE ALLTRIM(sel_reestr.doc_type) = '005'
                      m.pass_kod = '8'

                    OTHERWISE
                      m.pass_kod = '14'

                  ENDCASE

                  m.pass_seria = ALLTRIM(sel_reestr.pass_seria)
                  m.pass_num = ALLTRIM(sel_reestr.pass_num)
                  m.pass_data = sel_reestr.pass_data
                  m.pass_mesto = ALLTRIM(sel_reestr.pass_mesto)

                  m.pr_rekviz = ALLTRIM(sel_reestr.pr_rekviz)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                  DO CASE
                    CASE sel_reestr.vxd_ostat_val >= 0

                      m.num_dog_d = sel_reestr.num_dog
                      m.data_dog_d = sel_reestr.data_dog
                      m.num_liz_d = sel_reestr.num_liz_d

                      m.ost_val_d = sel_reestr.vxd_ostat_val
                      m.kurs = sel_reestr.kurs
                      m.ost_rur_d = sel_reestr.vxd_ostat_rur

                    CASE sel_reestr.vxd_ostat_val < 0

                      m.num_dog_d = sel_reestr.num_dog
                      m.data_dog_d = sel_reestr.data_dog
                      m.num_liz_d = sel_reestr.num_liz_d

                      m.ost_val_d = 0.00
                      m.ost_rur_d = 0.00

                      m.num_dog_k = sel_reestr.num_dog
                      m.data_dog_k = sel_reestr.data_dog
                      m.num_liz_k = sel_reestr.num_liz_d

                      m.ost_val_k = sel_reestr.vxd_ostat_val * -1
                      m.kurs = sel_reestr.kurs
                      m.ost_rur_k = sel_reestr.vxd_ostat_rur * -1

                  ENDCASE

                  DO CASE
                    CASE m.ost_rur_d > m.ost_rur_k

                      m.itog_ostat = m.ost_rur_d - m.ost_rur_k

                    CASE m.ost_rur_d < m.ost_rur_k

                      m.itog_ostat = m.ost_rur_k - m.ost_rur_d

                    CASE m.ost_rur_d = m.ost_rur_k

                      m.itog_ostat = m.ost_rur_d - m.ost_rur_k

                  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                  m.pr_soft = 'В'

                  IF poisk = .T.
                    GATHER MEMVAR MEMO
                  ELSE
                    INSERT INTO forma_1417_u FROM MEMVAR
                  ENDIF

                  SELECT sel_reestr
                ENDSCAN

                WAIT CLEAR

                SELECT client
                GO TOP
                REPLACE ALL dat_reestr WITH new_data_odb

                =soob('Заполнение данными таблицы реестра успешно завершено.')

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                SELECT DISTINCT SUBSTR(ALLTRIM(num_liz_d), 6, 3) AS kod_val, SUM(ost_val_d) AS summa_deb, SUM(ost_val_k) AS summa_cred ;
                  FROM forma_1417_u ;
                  INTO CURSOR sel_forma_1417 ;
                  WHERE date_vedom = dt1 ;
                  GROUP BY 1 ;
                  ORDER BY 1

                SELECT A.kod_val, A.summa AS summa_sum, B.summa_deb, B.summa_cred ;
                  FROM sel_1417 A, sel_forma_1417 B ;
                  WHERE ALLTRIM(A.kod_val) == ALLTRIM(B.kod_val) ;
                  INTO CURSOR sum_forma_1417 ;
                  ORDER BY 1

                BROWSE FIELDS ;
                  kod_val:H = 'Код валюты',;
                  summa_sum:H = 'Сумма общая':P = '999 999 999 999.99',;
                  summa_deb:H = 'Сумма по дебету':P = '999 999 999 999.99',;
                  summa_cred:H = 'Сумма по кредиту':P = '999 999 999 999.99' ;
                  TITLE ' Суммарные остатки на счетах по типам валют после формирования реестра 1417-У на дату ' + DT(dt1)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

              CASE BAR() = 3  && Печатный отчет по ранее сформированным данным за выбранную дату

                SELECT forma_1417_u

                SELECT date_vedom, branch, ref_client, fio_rus, fio_1, fio_2, fio_3, birth_date, srok_vklad, adres_reg, adres_dom, telef_dom, adres_eml,;
                  pass_kod, pass_seria, pass_num, pass_data, pass_mesto,;
                  num_dog_d, data_dog_d, num_liz_d, ost_val_d, kurs, ost_rur_d,;
                  num_dog_k, data_dog_k, num_liz_k, ost_val_k, ost_rur_k, itog_ostat,;
                  risk, date_risk, dat_reestr, date_read, pr_soft, pr_rekviz ;
                  FROM forma_1417_u ;
                  WHERE BETWEEN(DTOS(date_vedom), DTOS(dt1), DTOS(dt1)) ;
                  INTO CURSOR sel_forma_1417_u ;
                  ORDER BY fio_rus

                IF _TALLY <> 0

                  GO TOP

                  ima_frx = (put_frx) + ('forma_1417_u.frx')

                  =soob('Внимание! Печатный отчет сформирован для формата бумаги А3 ...')

                  DO prn_form IN Visa

                ELSE
                  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')
                ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

              CASE BAR() = 5  && Произвести удаление ранее сформированных данных за выбранную дату

                =soob('Выполнение данной процедуры требует монопольного режима работы с таблицами ... ')

                tex1 = 'В ы п о л н и т ь  удаление данных за выбранную дату'
                tex2 = 'О т к а з а т ь с я  от выполнения данной процедуры'
                l_bar = 3
                =popup_9(tex1,tex2,tex3,tex4,l_bar)
                ACTIVATE POPUP vibor
                RELEASE POPUPS vibor

                IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

                  IF BAR() = 1

                    IF USED('forma_1417_u')

                      SELECT forma_1417_u

                      COUNT FOR BETWEEN(date_vedom, dt1, dt1)

                      IF _TALLY <> 0

                        tim1 = SECONDS()

                        ACTIVATE WINDOW poisk
                        @ WROWS()/3,3 SAY PADC('Внимание! Начата упаковка таблицы и перестроение индексов.',WCOLS())

                        SET EXCLUSIVE ON

                        SELECT forma_1417_u
                        USE
                        USE (put_dbf) + ('forma_1417_u.dbf') ALIAS forma_1417_u ORDER TAG ref_client EXCLUSIVE

                        SELECT forma_1417_u
                        DELETE FOR BETWEEN(date_vedom, dt1, dt1)
                        PACK
                        REINDEX

                        SELECT forma_1417_u
                        USE
                        USE (put_dbf) + ('forma_1417_u.dbf') ALIAS forma_1417_u ORDER TAG ref_client SHARE

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

              CASE BAR() = 7  && Произвести экспорт ранее сформированных данных в DOS кодировке

                ACTIVATE WINDOW poisk
                @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт данных в DOS кодировке.',WCOLS())

                SELECT forma_1417_u

                SELECT date_vedom, branch, ref_client, fio_rus, fio_1, fio_2, fio_3, birth_date, srok_vklad, adres_reg, adres_dom, telef_dom, adres_eml,;
                  pass_kod, pass_seria, pass_num, pass_data, pass_mesto,;
                  num_dog_d, data_dog_d, num_liz_d, ost_val_d, kurs, ost_rur_d,;
                  num_dog_k, data_dog_k, num_liz_k, ost_val_k, ost_rur_k, itog_ostat,;
                  risk, date_risk, dat_reestr, date_read, pr_soft, pr_rekviz ;
                  FROM forma_1417_u ;
                  WHERE BETWEEN(DTOS(date_vedom), DTOS(dt1), DTOS(dt1)) ;
                  INTO CURSOR sel_forma_1417_u ;
                  ORDER BY fio_rus

                IF _TALLY <> 0

                  COPY TO (put_tmp) + ('forma_1417_u.dbf') FOX2X AS 866
                  COPY TO (put_tmp) + ('forma_1417_u.xls') XL5 AS 1251

                  @ WROWS()/3,3 SAY PADC('Файлы - ' + ('forma_1417_u.dbf') + ' и ' + ('forma_1417_u.xls') + ' успешно записаны по пути ' + (put_tmp),WCOLS())
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

ENDIF

RETURN


************************************************************** PROCEDURE start_vipiska ******************************************************************

PROCEDURE start_vipiska
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

=soob('Внимание! Формирование выписки из реестра обязательств банка перед вкладчиками.')

DEFINE WINDOW otchet_zb AT 0,0 SIZE colvo_rows - 2, colvo_cols - 4 FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW otchet_zb CENTER

SELECT DISTINCT date_vedom AS data ;
  FROM forma_1417_u ;
  INTO CURSOR tov ;
  ORDER BY 1 DESC

IF _TALLY <> 0

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  DO FORM (put_scr) + ('seldata1.scx')

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    =soob('Внимание! Формирование данных "Сведения о вкладчике, заключившем договор"')

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
    DO select_client IN Otchet_1417_U
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    =soob('Внимание! Формирование данных "Сведения об обязательствах банка перед вкладчиком"')

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
    DO select_obazat IN Otchet_1417_U
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    =soob('Внимание! Формирование данных "Сведения о встречных требованиях банка к вкладчику"')

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
    DO select_trebov IN Otchet_1417_U
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  ENDIF

ELSE

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')

ENDIF

RELEASE WINDOW otchet_zb
RETURN


************************************************************* PROCEDURE start_client ********************************************************************

PROCEDURE start_client
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

=soob('Внимание! Формирование текстового файла "Сведения о вкладчике, заключившем договор".')

DEFINE WINDOW otchet_zb AT 0,0 SIZE colvo_rows - 2, colvo_cols - 4 FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW otchet_zb CENTER

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, начата выборка данных ... ',WCOLS())

SELECT DISTINCT date_vedom AS data ;
  FROM forma_1417_u ;
  INTO CURSOR tov ;
  ORDER BY 1 DESC

IF _TALLY <> 0

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  DO FORM (put_scr) + ('seldata1.scx')

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
    DO select_client IN Otchet_1417_U
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  ENDIF

ELSE

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')

ENDIF

RELEASE WINDOW otchet_zb
RETURN


*************************************************************** PROCEDURE select_client *****************************************************************

PROCEDURE select_client

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, начата выборка данных ... ',WCOLS())

SELECT DISTINCT branch, date_vedom, ref_client, fio_rus, fio_1, fio_2, fio_3, birth_date,;
  UPPER(adres_reg) AS adres_reg, UPPER(adres_dom) AS adres_dom, adres_eml,;
  pass_kod, ALLTRIM(pass_seria) + ',' + ALLTRIM(pass_num) + ',' + DTOC(pass_data) + ',' + ALLTRIM(pass_mesto) AS passport,;
  MIN(data_dog_d) AS data_dog_d, risk, date_read, pr_rekviz ;
  FROM forma_1417_u ;
  WHERE BETWEEN(date_vedom, dt1, dt1) ;
  INTO CURSOR sel_forma_1417_u ;
  GROUP BY fio_rus, birth_date, ref_client ;
  ORDER BY fio_rus, birth_date, ref_client

* BROWSE WINDOW brows TITLE 'Количество клиентов = ' + ALLTRIM(TRANSFORM(_TALLY, '999 999 999'))

IF _TALLY <> 0

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  SELECT sel_forma_1417_u
  GO TOP

  DO WHILE .T.

    tex1 = 'Табличный просмотр сформированного реестра на экране'
    tex2 = 'Формирование текстового файла из сформированного реестра'
    l_bar = 3
    =popup_9(tex1,tex2,tex3,tex4,l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      DO CASE
        CASE BAR() = 1  && Табличный просмотр сформированного реестра на экране

          BROWSE FIELDS ;
            date_vedom:H = 'Дата',;
            branch:H = 'Номер филиала',;
            fio_1:H = 'Фамилия',;
            fio_2:H = 'Имя',;
            fio_3:H = 'Отчество',;
            pr_rekviz:H = 'Признак_наличия',;
            adres_reg:H = 'Адрес регистрации',;
            adres_dom:H = 'Адрес проживания',;
            adres_eml:H = 'Адрес E-mail',;
            pass_kod:H = 'Тип документа',;
            passport:H = 'Данные паспорта',;
            risk:H = 'Риск',;
            data_dog_d:H = 'Дата договора',;
            date_read:H = 'Дата правки' ;
            TITLE ' Сведения о вкладчике, заключившем договор' ;
            WINDOW otchet_zb

*              pass_seria:H = 'Серия паспорта',;
*              pass_num:H = 'Номер паспорта',;
*              pass_data:H = 'Дата выдачи',;
*              pass_mesto:H = 'Место выдачи',;

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

        CASE BAR() = 3  && Формирование текстового файла из сформированного реестра

          ACTIVATE WINDOW poisk
          @ WROWS()/3-0.2,3 SAY PADC('Внимание! Начато формирование текстового файла из сформированного реестра ... ',WCOLS())

          SET TEXTMERGE ON
          SET TEXTMERGE NOSHOW

          exp_file = 'inv' + RIGHT(ALLTRIM(STR(YEAR(dt1))), 1) + PADL(ALLTRIM(STR(MONTH(dt1))), 2, '0') + PADL(ALLTRIM(STR(DAY(dt1))), 2, '0') + '.' + ALLTRIM(num_branch)

          IF FILE((put_tmp) + (exp_file))
            DELETE FILE (put_tmp) + (exp_file)
          ENDIF

          _text = FCREATE((put_tmp) + (exp_file))

          IF _text < 1

            DEACTIVATE WINDOW poisk
            =err('Внимание! Ошибка при создании файла экспорта реестра ... ')
            =FCLOSE(_text)

          ELSE

            SELECT sel_forma_1417_u
            GO TOP

            SCAN

              f_ref_client = ALLTRIM(sel_forma_1417_u.ref_client) + CHR(94)  && PADL(ALLTRIM(STR(RECNO())), 10, '0')

              f_fio_1 = ALLTRIM(sel_forma_1417_u.fio_1) + CHR(94)
              f_fio_2 = ALLTRIM(sel_forma_1417_u.fio_2) + CHR(94)
              f_fio_3 = ALLTRIM(sel_forma_1417_u.fio_3) + CHR(94)

              f_adres_reg = ALLTRIM(sel_forma_1417_u.adres_reg) + CHR(94)
              f_adres_dom = ALLTRIM(sel_forma_1417_u.adres_dom) + CHR(94)

              f_adres_eml = ALLTRIM(sel_forma_1417_u.adres_eml) + CHR(94)

              f_pass_kod = ALLTRIM(sel_forma_1417_u.pass_kod) + CHR(94)

              f_passport = ALLTRIM(sel_forma_1417_u.passport) + CHR(94)

              f_birth_date = ALLTRIM(sel_forma_1417_u.birth_date) + CHR(94)

              f_risk = ALLTRIM(sel_forma_1417_u.risk) + CHR(94)

              f_data_dog_d = DTOC(sel_forma_1417_u.data_dog_d) + CHR(94)
              f_date_read = DTOC(sel_forma_1417_u.date_read) + CHR(94)

              f_pr_rekviz = ALLTRIM(sel_forma_1417_u.pr_rekviz)

              stroka_detal_record = f_ref_client + f_fio_1 + f_fio_2 + f_fio_3 + f_adres_reg + f_adres_dom + f_adres_eml + f_pass_kod + f_passport + f_birth_date + f_risk +  f_data_dog_d + f_date_read + f_pr_rekviz

              =FPUTS(_text, (stroka_detal_record))

            ENDSCAN

            =FCLOSE(_text)

            SET TEXTMERGE OFF

*              f_passport = ALLTRIM(sel_forma_1417_u.pass_seria) + ',' + ALLTRIM(sel_forma_1417_u.pass_num) + ',' + ;
DTOC(sel_forma_1417_u.pass_data) + ',' + ALLTRIM(sel_forma_1417_u.pass_mesto) + CHR(94)

*                f_date_risk = DTOC(sel_forma_1417_u.date_risk) + CHR(94)
*                f_dat_reestr = DTOC(sel_forma_1417_u.dat_reestr) + CHR(94)

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

  =err('Внимание! Данных о вкладчиках, заключившем договор не обнаружено ...')
ENDIF

RETURN


************************************************************** PROCEDURE start_obazat ******************************************************************

PROCEDURE start_obazat
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

=soob('Внимание! Формирование текстового файла "Сведения об обязательствах банка перед вкладчиком".')

DEFINE WINDOW otchet_zb AT 0,0 SIZE colvo_rows - 2, colvo_cols - 4 FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW otchet_zb CENTER

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, начата выборка данных ... ',WCOLS())

SELECT DISTINCT date_vedom AS data ;
  FROM forma_1417_u ;
  INTO CURSOR tov ;
  ORDER BY 1 DESC

IF _TALLY <> 0

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  DO FORM (put_scr) + ('seldata1.scx')

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
    DO select_obazat IN Otchet_1417_U
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

SELECT date_vedom, branch, ref_client, fio_rus, fio_1, fio_2, fio_3, birth_date, srok_vklad, adres_reg, adres_dom, telef_dom, adres_eml,;
  pass_kod, pass_seria, pass_num, pass_data, pass_mesto,;
  num_dog_d, data_dog_d, num_liz_d, ost_val_d, kurs, ost_rur_d,;
  num_dog_k, data_dog_k, num_liz_k, ost_val_k, ost_rur_k, itog_ostat,;
  risk, date_risk, dat_reestr, date_read, pr_soft ;
  FROM forma_1417_u ;
  WHERE BETWEEN(date_vedom, dt1, dt1) AND ost_val_d >= 0 ;
  INTO CURSOR sel_forma_1417_u ;
  ORDER BY fio_rus

IF _TALLY <> 0

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  SELECT sel_forma_1417_u
  GO TOP

  DO WHILE .T.

    tex1 = 'Табличный просмотр сформированного реестра на экране'
    tex2 = 'Формирование текстового файла из сформированного реестра'
    l_bar = 3
    =popup_9(tex1,tex2,tex3,tex4,l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      DO CASE
        CASE BAR() = 1  && Табличный просмотр сформированного реестра на экране

          BROWSE FIELDS ;
            date_vedom:H = 'Дата',;
            branch:H = 'Номер филиала',;
            ref_client:H = 'Номер вкладчика',;
            fio_1:H = 'Фамилия',;
            fio_2:H = 'Имя',;
            fio_3:H = 'Отчество',;
            birth_date:H = 'Дата рожденья',;
            srok_vklad:H = 'Срок вклада',;
            adres_reg:H = 'Адрес регистрации',;
            adres_dom:H = 'Адрес проживания',;
            telef_dom:H = 'Телефон для связи',;
            adres_eml:H = 'Адрес E-mail',;
            pass_kod:H = 'Тип документа',;
            pass_seria:H = 'Серия паспорта',;
            pass_num:H = 'Номер паспорта',;
            pass_data:H = 'Дата выдачи',;
            pass_mesto:H = 'Место выдачи',;
            num_dog_d:H = 'Номер договора',;
            data_dog_d:H = 'Дата договора',;
            num_liz_d:H = 'Лицевой счет вкладчика',;
            ost_val_d:H = 'Вход. остаток в валюте вклада':P = '999 999 999.99',;
            kurs:H = 'Курс валюты':P = '999.99',;
            ost_rur_d:H = 'Вход. остаток по курсу ЦБ':P = '999 999 999.99',;
            risk:H = 'Риск',;
            dat_reestr:H = 'Дата реестра',;
            date_read:H = 'Дата правки' ;
            TITLE ' Сведения об обязательствах банка перед вкладчиком' ;
            WINDOW otchet_zb

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

        CASE BAR() = 3  && Формирование текстового файла из сформированного реестра

          ACTIVATE WINDOW poisk
          @ WROWS()/3-0.2,3 SAY PADC('Внимание! Начато формирование текстового файла из сформированного реестра ... ',WCOLS())

          SET TEXTMERGE ON
          SET TEXTMERGE NOSHOW

          exp_file = 'dep' + RIGHT(ALLTRIM(STR(YEAR(dt1))), 1) + PADL(ALLTRIM(STR(MONTH(dt1))), 2, '0') + PADL(ALLTRIM(STR(DAY(dt1))), 2, '0') + '.' + ALLTRIM(num_branch)

          IF FILE((put_tmp) + (exp_file))
            DELETE FILE (put_tmp) + (exp_file)
          ENDIF

          _text = FCREATE((put_tmp) + (exp_file))

          IF _text < 1

            DEACTIVATE WINDOW poisk
            =err('Внимание! Ошибка при создании файла экспорта реестра ... ')
            =FCLOSE(_text)

          ELSE

            SELECT sel_forma_1417_u
            GO TOP

            SCAN

              f_ref_client = ALLTRIM(sel_forma_1417_u.ref_client) + CHR(94)
              f_branch = ALLTRIM(sel_forma_1417_u.branch) + CHR(94)
              f_num_dog_d = IIF(EMPTY(ALLTRIM(sel_forma_1417_u.num_dog_d)) = .F., ALLTRIM(sel_forma_1417_u.num_dog_d), 'б/н') + CHR(94)
              f_data_dog_d = DTOC(sel_forma_1417_u.data_dog_d) + CHR(94)
              f_num_liz_d = ALLTRIM(sel_forma_1417_u.num_liz_d) + CHR(94)
              f_ost_val_d = ALLTRIM(TRANSFORM(sel_forma_1417_u.ost_val_d, '999999999999.99')) + CHR(94)
              f_pr_soft = ALLTRIM(sel_forma_1417_u.pr_soft) + CHR(94)
              f_srok_vklad = ALLTRIM(STR(sel_forma_1417_u.srok_vklad))

              stroka_detal_record = f_ref_client + f_branch + f_num_dog_d + f_data_dog_d + f_num_liz_d + f_ost_val_d + f_pr_soft + f_srok_vklad

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


************************************************************** PROCEDURE start_trebov ******************************************************************

PROCEDURE start_trebov
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

=soob('Внимание! Формирование текстового файла "Сведения о встречных требованиях банка к вкладчику".')

DEFINE WINDOW otchet_zb AT 0,0 SIZE colvo_rows - 2, colvo_cols - 4 FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW otchet_zb CENTER

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, начата выборка данных ... ',WCOLS())

SELECT DISTINCT date_vedom AS data ;
  FROM forma_1417_u ;
  INTO CURSOR tov ;
  ORDER BY 1 DESC

IF _TALLY <> 0

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  DO FORM (put_scr) + ('seldata1.scx')

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
    DO select_trebov IN Otchet_1417_U
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  ENDIF

ELSE

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')

ENDIF

RELEASE WINDOW otchet_zb
RETURN


*************************************************************** PROCEDURE select_trebov *****************************************************************

PROCEDURE select_trebov

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, начата выборка данных ... ',WCOLS())

SELECT date_vedom, branch, ref_client, fio_rus, fio_1, fio_2, fio_3, birth_date, srok_vklad, adres_reg, adres_dom, telef_dom, adres_eml,;
  pass_kod, pass_seria, pass_num, pass_data, pass_mesto,;
  num_dog_d, data_dog_d, num_liz_d, ost_val_d, kurs, ost_rur_d,;
  num_dog_k, data_dog_k, num_liz_k, ost_val_k, ost_rur_k, itog_ostat,;
  risk, date_risk, dat_reestr, date_read, pr_soft ;
  FROM forma_1417_u ;
  WHERE BETWEEN(date_vedom, dt1, dt1) AND ost_val_k <> 0 ;
  INTO CURSOR sel_forma_1417_u ;
  ORDER BY fio_rus

IF _TALLY <> 0

  =INKEY(2)
  DEACTIVATE WINDOW poisk

  SELECT sel_forma_1417_u
  GO TOP

  DO WHILE .T.

    tex1 = 'Табличный просмотр сформированного реестра на экране'
    tex2 = 'Формирование текстового файла из сформированного реестра'
    l_bar = 3
    =popup_9(tex1,tex2,tex3,tex4,l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      DO CASE
        CASE BAR() = 1  && Табличный просмотр сформированного реестра на экране

          BROWSE FIELDS ;
            date_vedom:H = 'Дата',;
            branch:H = 'Номер филиала',;
            ref_client:H = 'Номер вкладчика',;
            fio_1:H = 'Фамилия',;
            fio_2:H = 'Имя',;
            fio_3:H = 'Отчество',;
            birth_date:H = 'Дата рожденья',;
            srok_vklad:H = 'Срок вклада',;
            adres_reg:H = 'Адрес регистрации',;
            adres_dom:H = 'Адрес проживания',;
            telef_dom:H = 'Телефон для связи',;
            adres_eml:H = 'Адрес E-mail',;
            pass_kod:H = 'Тип документа',;
            pass_seria:H = 'Серия паспорта',;
            pass_num:H = 'Номер паспорта',;
            pass_data:H = 'Дата выдачи',;
            pass_mesto:H = 'Место выдачи',;
            num_dog_k:H = 'Номер договора',;
            data_dog_k:H = 'Дата договора',;
            num_liz_k:H = 'Лицевой счет вкладчика',;
            ost_val_k:H = 'Вход. остаток в валюте вклада':P = '999 999 999.99',;
            kurs:H = 'Курс валюты':P = '999.99',;
            ost_rur_k:H = 'Вход. остаток по курсу ЦБ':P = '999 999 999.99',;
            risk:H = 'Риск',;
            dat_reestr:H = 'Дата реестра',;
            date_read:H = 'Дата правки' ;
            TITLE ' Сведения о встречных требованиях банка к вкладчику' ;
            WINDOW otchet_zb

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

        CASE BAR() = 3  && Формирование текстового файла из сформированного реестра

          ACTIVATE WINDOW poisk
          @ WROWS()/3-0.2,3 SAY PADC('Внимание! Начато формирование текстового файла из сформированного реестра ... ',WCOLS())

          SET TEXTMERGE ON
          SET TEXTMERGE NOSHOW

          exp_file = 'crd' + RIGHT(ALLTRIM(STR(YEAR(dt1))), 1) + PADL(ALLTRIM(STR(MONTH(dt1))), 2, '0') + PADL(ALLTRIM(STR(DAY(dt1))), 2, '0') + '.' + ALLTRIM(num_branch)

          IF FILE((put_tmp) + (exp_file))
            DELETE FILE (put_tmp) + (exp_file)
          ENDIF

          _text = FCREATE((put_tmp) + (exp_file))

          IF _text < 1

            DEACTIVATE WINDOW poisk
            =err('Внимание! Ошибка при создании файла экспорта реестра ... ')
            =FCLOSE(_text)

          ELSE

            SELECT sel_forma_1417_u
            GO TOP

            SCAN

              f_ref_client = ALLTRIM(sel_forma_1417_u.ref_client) + CHR(94)
              f_branch = ALLTRIM(sel_forma_1417_u.branch) + CHR(94)
              f_num_dog_k = IIF(EMPTY(ALLTRIM(sel_forma_1417_u.num_dog_k)) = .F., ALLTRIM(sel_forma_1417_u.num_dog_k), 'б/н') + CHR(94)
              f_data_dog_k = DTOC(sel_forma_1417_u.data_dog_k) + CHR(94)
              f_num_liz_k = ALLTRIM(sel_forma_1417_u.num_liz_k) + CHR(94)
              f_ost_val_k = ALLTRIM(TRANSFORM(sel_forma_1417_u.ost_val_k, '999999999999.99')) + CHR(94)
              f_pr_soft = ALLTRIM(sel_forma_1417_u.pr_soft) + CHR(94)
              f_srok_vklad = ALLTRIM(STR(sel_forma_1417_u.srok_vklad))

              stroka_detal_record = f_ref_client + f_branch + f_num_dog_k + f_data_dog_k + f_num_liz_k + f_ost_val_k + f_pr_soft + f_srok_vklad

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

  =err('Внимание! Данных о встречных требованиях банка к вкладчику не обнаружено ...')
ENDIF

RETURN


*************************************************************** PROCEDURE start_bank *******************************************************************

PROCEDURE start_bank
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

=soob('Внимание! Формирование текстового файла "Сведения о банке, составившем реестр".')

DEFINE WINDOW otchet_zb AT 0,0 SIZE 4, colvo_cols - 40 FONT 'Times New Roman', num_schrift_v + 1  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW otchet_zb CENTER

SELECT ALLTRIM(bank) AS name_bank, LEFT(ALLTRIM(reg_number), 4) AS reg_number ;
  FROM rekvizit ;
  INTO CURSOR sel_bank

IF _TALLY <> 0

  EDIT FIELDS ;
    name_bank:H = 'Полное наименование банка :',;
    reg_number:H = 'Регистрационный номер банка :' ;
    TITLE ' Сведения о банке, составившем реестр' ;
    WINDOW otchet_zb

ELSE
  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')
ENDIF

RELEASE WINDOW otchet_zb
RETURN


************************************************************** PROCEDURE start_filial ********************************************************************

PROCEDURE start_filial
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

=soob('Внимание! Формирование текстового файла "Сведения о филиале банка, заключившем договор".')

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
    TITLE ' Сведения о филиале банка, заключившем договор' ;
    WINDOW otchet_zb

ELSE
  =err('Внимание! Запрашиваемых Вами данных не обнаружено ...')
ENDIF

RELEASE WINDOW otchet_zb
RETURN


************************************************************** PROCEDURE start_kontrol ******************************************************************

PROCEDURE start_kontrol
sel = SELECT()
popup_text = LOWER(POPUP())
HIDE POPUP (popup_text)

=soob('Внимание! Формирование текстового контрольного файла.')

=soob('Внимание! Данный пункт меню находится в стадии разработки ... ')

RETURN


************************************************************* PROCEDURE formirov_menu ******************************************************************

PROCEDURE formirov_menu

text_otchet_1417_1 = 'Формирование реестра обязательств банка перед вкладчиками'
text_otchet_1417_2 = 'Формирование выписки из реестра обязательств банка перед вкладчиками'
text_otchet_1417_3 = 'Формирование текстового файла "Сведения о банке, составившем реестр"'
text_otchet_1417_4 = 'Формирование текстового файла "Сведения о филиале банка, заключившем договор"'
text_otchet_1417_5 = 'Формирование текстового файла "Сведения о вкладчике, заключившем договор"'
text_otchet_1417_6 = 'Формирование текстового файла "Сведения об обязательствах банка перед вкладчиком"'
text_otchet_1417_7 = 'Формирование текстового файла "Сведения о встречных требованиях банка к вкладчику"'
text_otchet_1417_8 = 'Формирование текстового контрольного файла'

len_otchet_1417_1 = TXTWIDTH(text_otchet_1417_1,'Times New Roman', num_schrift_v,'B')
len_otchet_1417_2 = TXTWIDTH(text_otchet_1417_2,'Times New Roman', num_schrift_v,'B')
len_otchet_1417_3 = TXTWIDTH(text_otchet_1417_3,'Times New Roman', num_schrift_v,'B')
len_otchet_1417_4 = TXTWIDTH(text_otchet_1417_4,'Times New Roman', num_schrift_v,'B')
len_otchet_1417_5 = TXTWIDTH(text_otchet_1417_5,'Times New Roman', num_schrift_v,'B')
len_otchet_1417_6 = TXTWIDTH(text_otchet_1417_6,'Times New Roman', num_schrift_v,'B')
len_otchet_1417_7 = TXTWIDTH(text_otchet_1417_7,'Times New Roman', num_schrift_v,'B')
len_otchet_1417_8 = TXTWIDTH(text_otchet_1417_8,'Times New Roman', num_schrift_v,'B')

l_bar = 24.0
row = IIF(l_bar<=25, l_bar, 25)

V_row = (WROWS()-row)/2 + 1
Lmenu = MAX(len_otchet_1417_1, len_otchet_1417_2, len_otchet_1417_3, len_otchet_1417_4, len_otchet_1417_5, len_otchet_1417_6, len_otchet_1417_7, len_otchet_1417_8)
L_col = (WCOLS()-Lmenu)/2 - 0

DEFINE POPUP otchet_1417_U FROM V_row,L_col ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1 OF otchet_1417_U PROMPT text_otchet_1417_1 ;
  MESSAGE 'Расчет и формирование реестра обязательств банка перед вкладчиками'
DEFINE BAR 2 OF otchet_1417_U PROMPT '\-'
DEFINE BAR 3 OF otchet_1417_U PROMPT text_otchet_1417_2 ;
  MESSAGE 'Расчет и формирование выписки из реестра обязательств банка перед вкладчиками'
DEFINE BAR 4 OF otchet_1417_U PROMPT '\-'
DEFINE BAR 5 OF otchet_1417_U PROMPT text_otchet_1417_3 ;
  MESSAGE 'Формирование текстового файла "Сведения о банке, составившем реестр"'
DEFINE BAR 6 OF otchet_1417_U PROMPT '\-'
DEFINE BAR 7 OF otchet_1417_U PROMPT text_otchet_1417_4 ;
  MESSAGE 'Формирование текстового файла "Сведения о филиале банка, заключившем договор"'
DEFINE BAR 8 OF otchet_1417_U PROMPT '\-'
DEFINE BAR 9 OF otchet_1417_U PROMPT text_otchet_1417_5 ;
  MESSAGE 'Формирование текстового файла "Сведения о вкладчике, заключившем договор"'
DEFINE BAR 10 OF otchet_1417_U PROMPT '\-'
DEFINE BAR 11 OF otchet_1417_U PROMPT text_otchet_1417_6 ;
  MESSAGE 'Формирование текстового файла "Сведения об обязательствах банка перед вкладчиком"'
DEFINE BAR 12 OF otchet_1417_U PROMPT '\-'
DEFINE BAR 13 OF otchet_1417_U PROMPT text_otchet_1417_7 ;
  MESSAGE 'Формирование текстового файла "Сведения о встречных требованиях банка к вкладчику"'
DEFINE BAR 14 OF otchet_1417_U PROMPT '\-'
DEFINE BAR 15 OF otchet_1417_U PROMPT text_otchet_1417_8 ;
  MESSAGE 'Формирование текстового контрольного файла'
ON SELECTION BAR 1 OF otchet_1417_U DO start_reestr IN Otchet_1417_U
ON SELECTION BAR 3 OF otchet_1417_U DO start_vipiska IN Otchet_1417_U
ON SELECTION BAR 5 OF otchet_1417_U DO start_bank IN Otchet_1417_U
ON SELECTION BAR 7 OF otchet_1417_U DO start_filial IN Otchet_1417_U
ON SELECTION BAR 9 OF otchet_1417_U DO start_client IN Otchet_1417_U
ON SELECTION BAR 11 OF otchet_1417_U DO start_obazat IN Otchet_1417_U
ON SELECTION BAR 13 OF otchet_1417_U DO start_trebov IN Otchet_1417_U
ON SELECTION BAR 15 OF otchet_1417_U DO start_kontrol IN Otchet_1417_U

RETURN


********************************************************************************************************************************************************

* SET FILTER TO BETWEEN(date_vedom, CTOD('01.12.2008'), CTOD('11.01.2009')) AND ALLTRIM(card_acct) == '40817810841030100817'
* BROWSE FIELDS card_acct, fio_rus, date_vedom, begin_bal, vxd_ost_m, debit, debit_m, credit, credit_m, end_bal, isx_ost_m
* SET FILTER TO BETWEEN(date_vedom, CTOD('01.12.2008'), CTOD('11.01.2009')) AND end_bal <> isx_ost_m




