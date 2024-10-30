************************************************************************************
*** Экспорт данных из локальных таблиц в таблицы SQL Server 2000 и SQL Server 2005 ***
************************************************************************************

**************************************************************** PROCEDURE start ************************************************************************

PROCEDURE start
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

put_vibor = 1

text_exp_sql_01 = 'Экспорт данных из таблицы "Справочник клиентов в ПО "CARD-виза""'
text_exp_sql_02 = 'Экспорт данных из таблицы "Справочник по открытым картам в ПО "CARD-виза""'
text_exp_sql_03 = 'Экспорт данных из таблицы "Справочник по закрытым картам в ПО "CARD-виза""'
text_exp_sql_04 = 'Экспорт данных из таблицы "Проведенные документы из ПО "Трансмастер""'
text_exp_sql_05 = 'Экспорт данных из таблицы "Проведенные документы из ПО "CARD-виза""'
text_exp_sql_06 = 'Экспорт данных из таблицы "Документы управления счетом отправленные в ПО "Трансмастер""'
text_exp_sql_07 = 'Экспорт данных из таблицы "История остатков из ПО "Трансмастер""'
text_exp_sql_08 = 'Экспорт данных из таблицы "История остатков из ПО "CARD-виза""'
text_exp_sql_09 = 'Экспорт данных из таблицы "Ведомость остатков в ПО "CARD-виза""'
text_exp_sql_10 = 'Экспорт данных из таблицы "Ведомость остатков по счетам импортированная из ПО "Трансмастер""'
text_exp_sql_11 = 'Экспорт данных из таблицы "Расчитанные проценты по дням в ПО "CARD-виза""'
text_exp_sql_12 = 'Экспорт данных из таблицы "Расчитанные проценты суммарные в ПО "CARD-виза""'
text_exp_sql_13 = 'Экспорт данных из таблицы "Составленные мемоордера по рублевым счетам в ПО "CARD-виза""'
text_exp_sql_14 = 'Экспорт данных из таблицы "Составленные мемоордера по валютным счетам в ПО "CARD-виза""'
text_exp_sql_15 = 'Экспорт данных из таблицы "Приходные кассовые ордера по рублевым счетам в ПО "CARD-виза""'
text_exp_sql_16 = 'Экспорт данных из таблицы "Приходные кассовые ордера по валютным счетам в ПО "CARD-виза""'
text_exp_sql_17 = 'Экспорт данных из таблицы "Расходные кассовые ордера по рублевым счетам в ПО "CARD-виза""'
text_exp_sql_18 = 'Экспорт данных из таблицы "Расходные кассовые ордера по валютным счетам в ПО "CARD-виза""'
text_exp_sql_19 = 'Экспорт данных из таблицы "Составленные по счетам платежные поручения в ПО "CARD-виза""'
text_exp_sql_20 = 'Экспорт данных из таблицы "Выданные на печать платежные поручения в ПО "CARD-виза""'
text_exp_sql_21 = 'Экспорт справочных данных по операциям в ПОЛНОМ ОБЪЕМЕ из всех таблиц'

len_exp_sql_01 = TXTWIDTH(text_exp_sql_01, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_02 = TXTWIDTH(text_exp_sql_02, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_03 = TXTWIDTH(text_exp_sql_03, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_04 = TXTWIDTH(text_exp_sql_04, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_05 = TXTWIDTH(text_exp_sql_05, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_06 = TXTWIDTH(text_exp_sql_06, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_07 = TXTWIDTH(text_exp_sql_07, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_08 = TXTWIDTH(text_exp_sql_08, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_09 = TXTWIDTH(text_exp_sql_09, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_10 = TXTWIDTH(text_exp_sql_10, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_11 = TXTWIDTH(text_exp_sql_11, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_12 = TXTWIDTH(text_exp_sql_12, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_13 = TXTWIDTH(text_exp_sql_13, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_14 = TXTWIDTH(text_exp_sql_14, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_15 = TXTWIDTH(text_exp_sql_15, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_16 = TXTWIDTH(text_exp_sql_16, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_17 = TXTWIDTH(text_exp_sql_17, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_18 = TXTWIDTH(text_exp_sql_18, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_19 = TXTWIDTH(text_exp_sql_19, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_20 = TXTWIDTH(text_exp_sql_20, 'Times New Roman', num_schrift_v, 'B')
len_exp_sql_21 = TXTWIDTH(text_exp_sql_20, 'Times New Roman', num_schrift_v, 'B')

l_bar = 42.0
row = IIF(l_bar<=25, l_bar, 25)

V_row = (WROWS()-row)/2 - 7
Lmenu = MAX(len_exp_sql_01, len_exp_sql_02, len_exp_sql_03, len_exp_sql_04, len_exp_sql_05, len_exp_sql_06, len_exp_sql_07, len_exp_sql_08, len_exp_sql_09, len_exp_sql_10,;
  len_exp_sql_11, len_exp_sql_12, len_exp_sql_13, len_exp_sql_14, len_exp_sql_15, len_exp_sql_16, len_exp_sql_17, len_exp_sql_18, len_exp_sql_19, len_exp_sql_20, len_exp_sql_21)
L_col = (WCOLS()-Lmenu)/2 - 0

DEFINE POPUP export_sql FROM V_row,L_col ;
  FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
  MARGIN COLOR SCHEME 2 MESSAGE text_mess
DEFINE BAR 1  OF export_sql PROMPT text_exp_sql_01
DEFINE BAR 2  OF export_sql PROMPT '\-'
DEFINE BAR 3  OF export_sql PROMPT text_exp_sql_02
DEFINE BAR 4  OF export_sql PROMPT '\-'
DEFINE BAR 5  OF export_sql PROMPT text_exp_sql_03
DEFINE BAR 6  OF export_sql PROMPT '\-'
DEFINE BAR 7  OF export_sql PROMPT text_exp_sql_04
DEFINE BAR 8  OF export_sql PROMPT '\-'
DEFINE BAR 9  OF export_sql PROMPT text_exp_sql_05
DEFINE BAR 10 OF export_sql PROMPT '\-'
DEFINE BAR 11 OF export_sql PROMPT text_exp_sql_06
DEFINE BAR 12  OF export_sql PROMPT '\-'
DEFINE BAR 13 OF export_sql PROMPT text_exp_sql_07
DEFINE BAR 14 OF export_sql PROMPT '\-'
DEFINE BAR 15 OF export_sql PROMPT text_exp_sql_08
DEFINE BAR 16 OF export_sql PROMPT '\-'
DEFINE BAR 17 OF export_sql PROMPT text_exp_sql_09
DEFINE BAR 18 OF export_sql PROMPT '\-'
DEFINE BAR 19 OF export_sql PROMPT text_exp_sql_10
DEFINE BAR 20 OF export_sql PROMPT '\-'
DEFINE BAR 21  OF export_sql PROMPT text_exp_sql_11
DEFINE BAR 22  OF export_sql PROMPT '\-'
DEFINE BAR 23  OF export_sql PROMPT text_exp_sql_12
DEFINE BAR 24  OF export_sql PROMPT '\-'
DEFINE BAR 25  OF export_sql PROMPT text_exp_sql_13
DEFINE BAR 26  OF export_sql PROMPT '\-'
DEFINE BAR 27  OF export_sql PROMPT text_exp_sql_14
DEFINE BAR 28  OF export_sql PROMPT '\-'
DEFINE BAR 29  OF export_sql PROMPT text_exp_sql_15
DEFINE BAR 30 OF export_sql PROMPT '\-'
DEFINE BAR 31 OF export_sql PROMPT text_exp_sql_16
DEFINE BAR 32  OF export_sql PROMPT '\-'
DEFINE BAR 33 OF export_sql PROMPT text_exp_sql_17
DEFINE BAR 34 OF export_sql PROMPT '\-'
DEFINE BAR 35 OF export_sql PROMPT text_exp_sql_18
DEFINE BAR 36 OF export_sql PROMPT '\-'
DEFINE BAR 37 OF export_sql PROMPT text_exp_sql_19
DEFINE BAR 38 OF export_sql PROMPT '\-'
DEFINE BAR 39 OF export_sql PROMPT text_exp_sql_20
DEFINE BAR 40 OF export_sql PROMPT '\-'
DEFINE BAR 41 OF export_sql PROMPT text_exp_sql_21

ON SELECTION POPUP export_sql DO start_export_sql IN Export_data_sql

DO WHILE .T.

  ACTIVATE POPUP export_sql

  punkt = IIF(BAR() = 0, 1, BAR())

  IF LASTKEY() = 27
    EXIT
  ENDIF

ENDDO

DEACTIVATE POPUP export_sql
RELEASE POPUP export_sql


************************************************************** PROCEDURE start_export_sql ****************************************************************

PROCEDURE start_export_sql
HIDE POPUP export_sql

put_vibor = BAR()

scan_date_home = GOMONTH(new_data_odb, colvo_mes_arxiv)
scan_date_end = new_data_odb
*scan_date_end = GOMONTH(new_data_odb, -22)

*scan_date_home = GOMONTH(new_data_odb, - 15)
*scan_date_end = new_data_odb

=soob('Начальная дата экспорта - ' + DTOC(scan_date_home) + '   Конечная дата экспорта - ' + DTOC(scan_date_end))

* WAIT WINDOW 'put_vibor = ' + ALLTRIM(STR(put_vibor))

DO CASE
  CASE put_vibor = 1  && Экспорт данных из таблицы "Справочник клиентов в ПО "CARD-виза""

    DO export_client IN Export_data_sql WITH ima_sql_server

  CASE put_vibor = 3  && Экспорт данных из таблицы "Справочник по открытым картам в ПО "CARD-виза""

    DO export_account IN Export_data_sql WITH ima_sql_server

  CASE put_vibor = 5  && Экспорт данных из таблицы "Справочник по закрытым картам в ПО "CARD-виза""

    DO export_acc_del IN Export_data_sql WITH ima_sql_server

  CASE put_vibor = 7  && Экспорт данных из таблицы "Проведенные документы из ПО "Трансмастер""

    DO export_docum_slip IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 9  && Экспорт данных из таблицы "Проведенные документы из ПО "CARD-виза""

    DO export_docum_mak IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 11  && Экспорт данных из таблицы "Документы управления счетом отправленные в ПО "Трансмастер""

    DO export_popol_arx IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 13  && Экспорт данных из таблицы "История остатков из ПО "Трансмастер""

    DO export_istor_ost IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 15  && Экспорт данных из таблицы "История остатков из ПО "CARD-виза""

    DO export_istor_mak IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 17  && Экспорт данных из таблицы "Ведомость остатков в ПО "CARD-виза""

    DO export_vedom_ost IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 19  && Экспорт данных из таблицы "Ведомость остатков по счетам импортированная из ПО "Трансмастер""

    DO export_ostat_upk IN Export_data_sql WITH ima_sql_server

  CASE put_vibor = 21  && Экспорт данных из таблицы "Расчитанные проценты по дням в ПО "CARD-виза""

    DO export_acc_cln_dt IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 23  && Экспорт данных из таблицы "Расчитанные проценты суммарные в ПО "CARD-виза""

    DO export_acc_cln_sum IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 25  && Экспорт данных из таблицы "Составленные мемоордера по рублевым счетам в ПО "CARD-виза""

    DO export_memordrub IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 27  && Экспорт данных из таблицы "Составленные мемоордера по валютным счетам в ПО "CARD-виза""

    DO export_memordval IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 29  && Экспорт данных из таблицы "Приходные кассовые ордера по рублевым счетам в ПО "CARD-виза""

    DO export_prixodrub IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 31  && Экспорт данных из таблицы "Приходные кассовые ордера по валютным счетам в ПО "CARD-виза""

    DO export_prixodval IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 33  && Экспорт данных из таблицы "Расходные кассовые ордера по рублевым счетам в ПО "CARD-виза""

    DO export_rasxodrub IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 35  && Экспорт данных из таблицы "Расходные кассовые ордера по валютным счетам в ПО "CARD-виза""

    DO export_rasxodval IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 37  && Экспорт данных из таблицы "Составленные по счетам платежные поручения в ПО "CARD-виза""

    DO export_plateg IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 39  && Экспорт данных из таблицы "Составленные по счетам платежные поручения в ПО "CARD-виза""

    DO export_platprn IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

  CASE put_vibor = 41  && Экспорт справочных данных по операциям в ПОЛНОМ ОБЪЕМЕ из всех таблиц

    DO export_client IN Export_data_sql WITH ima_sql_server

    DO export_account IN Export_data_sql WITH ima_sql_server
    DO export_acc_del IN Export_data_sql WITH ima_sql_server

    DO export_docum_slip IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
    DO export_docum_mak IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
    DO export_popol_arx IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

    DO export_istor_ost IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
    DO export_istor_mak IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

    DO export_vedom_ost IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
    DO export_ostat_upk IN Export_data_sql WITH ima_sql_server

    DO export_acc_cln_dt IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
    DO export_acc_cln_sum IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

    DO export_memordrub IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
    DO export_memordval IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

    DO export_prixodrub IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
    DO export_prixodval IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
    DO export_rasxodrub IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
    DO export_rasxodval IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

    DO export_plateg IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
    DO export_platprn IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

ENDCASE

SELECT(sel)
RETURN


************************************************************************ PROCEDURE export_client ********************************************************

PROCEDURE export_client
PARAMETERS name_sql_server

SET ESCAPE ON

IF FILE((put_dbf) + ('client.dbf'))

  IF NOT USED('client')

    SELECT 0
    USE (put_dbf) + ('client.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT client
  SET ORDER TO fio_rus

  COUNT TO colvo_zap

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт справочника клиентов  в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT client
    GO TOP

    SCAN

      time_home = SECONDS()

      SCATTER MEMVAR MEMO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      DO scan_client IN Export_data_sql
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      WAIT WINDOW 'Филиал № ' + ALLTRIM(m.branch) + ' Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер клиента - ' + ALLTRIM(m.ref_client) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT client
    ENDSCAN

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт справочника клиентов в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('client.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


**************************************************************** PROCEDURE scan_client *******************************************************************

PROCEDURE scan_client

f_branch = ALLTRIM(m.branch)
f_ref_client = ALLTRIM(m.ref_client)
f_data_clien = DTOT(m.data_clien)
f_rec_date = DTOT(m.rec_date)
f_client_b = ALLTRIM(m.client_b)
f_inn = ALLTRIM(m.inn)
f_fio_rus = ALLTRIM(m.fio_rus)
f_cln_cat = ALLTRIM(m.cln_cat)
f_resident = ALLTRIM(m.resident)
f_tip_cln = ALLTRIM(m.tip_cln)
f_dolg_rab = ALLTRIM(m.dolg_rab)
f_mesto_rab = ALLTRIM(m.mesto_rab)
f_adres_rab = ALLTRIM(m.adres_rab)
f_telef_rab = ALLTRIM(m.telef_rab)
f_telef_dom = ALLTRIM(m.telef_dom)
f_type_sms = ALLTRIM(m.type_sms)
f_mob_telef = ALLTRIM(m.mob_telef)
f_vid_doc = ALLTRIM(m.vid_doc)
f_doc_type = ALLTRIM(m.doc_type)
f_pass_seria = ALLTRIM(m.pass_seria)
f_pass_num = ALLTRIM(m.pass_num)
f_pass_data = DTOT(m.pass_data)
f_pass_mesto = ALLTRIM(m.pass_mesto)
f_birth_date = DTOT(m.birth_date)
f_birth_city = ALLTRIM(m.birth_city)
f_doveren = ALLTRIM(m.doveren)
f_e_mail = ALLTRIM(m.e_mail)
f_risk = ALLTRIM(m.risk)
f_dat_reestr = DTOT(m.dat_reestr)
f_date_read = DTOT(m.date_read)
f_ind_dom = ALLTRIM(m.ind_dom)
f_ind_reg = ALLTRIM(m.ind_reg)
f_city_dom = ALLTRIM(m.city_dom)
f_city_reg = ALLTRIM(m.city_reg)
f_street_dom = ALLTRIM(m.street_dom)
f_street_reg = ALLTRIM(m.street_reg)
f_adres_dom = ALLTRIM(m.adres_dom)
f_adres_reg = ALLTRIM(m.adres_reg)
f_kladr_dom = ALLTRIM(m.kladr_dom)
f_kladr_reg = ALLTRIM(m.kladr_reg)
f_migr_seria = ALLTRIM(m.migr_seria)
f_migr_num = ALLTRIM(m.migr_num)
f_migr_data1 = DTOT(m.migr_data1)
f_migr_data2 = DTOT(m.migr_data2)
f_data = m.data
f_ima_pk = ALLTRIM(m.ima_pk)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

stroka_sql = "EXECUTE export_client"

stroka_sql = stroka_sql + " ?f_branch, ?f_ref_client, ?f_data_clien, ?f_rec_date, ?f_client_b, ?f_inn, ?f_fio_rus, ?f_cln_cat, ?f_resident, ?f_tip_cln, ?f_dolg_rab, ?f_mesto_rab,"
stroka_sql = stroka_sql + " ?f_adres_rab, ?f_telef_rab, ?f_telef_dom, ?f_type_sms, ?f_mob_telef, ?f_vid_doc, ?f_doc_type, ?f_pass_seria, ?f_pass_num, ?f_pass_data, ?f_pass_mesto,"
stroka_sql = stroka_sql + " ?f_birth_date, ?f_birth_city, ?f_doveren, ?f_e_mail, ?f_risk, ?f_dat_reestr, ?f_date_read, ?f_ind_dom, ?f_ind_reg, ?f_city_dom, ?f_city_reg,"
stroka_sql = stroka_sql + " ?f_street_dom, ?f_street_reg, ?f_adres_dom, ?f_adres_reg, ?f_kladr_dom, ?f_kladr_reg, ?f_migr_seria,"
stroka_sql = stroka_sql + " ?f_migr_num, ?f_migr_data1, ?f_migr_data2, ?f_data, ?f_ima_pk"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

export_data = SQLEXEC(sql_num, stroka_sql, '')

DO WHILE SQLMORERESULTS(sql_num) < 2
ENDDO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

RETURN


********************************************************************** PROCEDURE export_account ********************************************************

PROCEDURE export_account
PARAMETERS name_sql_server

SET ESCAPE ON

IF FILE((put_dbf) + ('account.dbf'))

  IF NOT USED('account')

    SELECT 0
    USE (put_dbf) + ('account.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT account
  SET ORDER TO fio_rus

  COUNT TO colvo_zap

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт справочника по открытым картам  в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT account
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN

      time_home = SECONDS()

      SCATTER MEMVAR MEMO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      DO scan_account IN Export_data_sql
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Филиал № ' + ALLTRIM(m.branch) + ' Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT account
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт справочника по открытым картам в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('account.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


************************************************************************ PROCEDURE scan_account ********************************************************

PROCEDURE scan_account

f_branch = ALLTRIM(m.branch)
f_ref_client = ALLTRIM(m.ref_client)
f_ref_card = m.ref_card
f_ref_acc = m.ref_acc
f_ref_ost = m.ref_ost
f_date_card = m.date_card
f_card_num = ALLTRIM(m.card_num)
f_vid_card = ALLTRIM(m.vid_card)
f_kod_proekt = ALLTRIM(m.kod_proekt)
f_ref_poisk = m.ref_poisk
f_del_card = m.del_card
f_date_nls = m.date_nls
f_date_close = m.date_close
f_pr_nls = m.pr_nls
f_pr_odb = m.pr_odb
f_resident = ALLTRIM(m.resident)
f_card_nls = ALLTRIM(m.card_nls)
f_card_acct = ALLTRIM(m.card_acct)
f_vklad_acct = ALLTRIM(m.vklad_acct)
f_status_tra = ALLTRIM(m.status_tra)
f_status_mak = ALLTRIM(m.status_mak)
f_fio_rus = ALLTRIM(m.fio_rus)
f_fio_eng = ALLTRIM(m.fio_eng)
f_kod_slovo = ALLTRIM(m.kod_slovo)
f_tip_cln = ALLTRIM(m.tip_cln)
f_type_card = ALLTRIM(m.type_card)
f_num_card = ALLTRIM(m.num_card)
f_val_card = ALLTRIM(m.val_card)
f_name_card = ALLTRIM(m.name_card)
f_date_ostat = m.date_ostat
f_read_ost_p = m.read_ost_p
f_begin_bal = m.begin_bal
f_debit = m.debit
f_credit = m.credit
f_end_bal = m.end_bal
f_read_ost_m = m.read_ost_m
f_pr_izm_ost = m.pr_izm_ost
f_vxd_ost_m = m.vxd_ost_m
f_debit_m = m.debit_m
f_credit_m = m.credit_m
f_isx_ost_m = m.isx_ost_m
f_card_strah = ALLTRIM(m.card_strah)
f_tran_42301 = ALLTRIM(m.tran_42301)
f_tran_42308 = ALLTRIM(m.tran_42308)
f_num_dog = ALLTRIM(m.num_dog)
f_data_dog = m.data_dog
f_data_proz = m.data_proz
f_stavka = m.stavka
f_num_47411 = ALLTRIM(m.num_47411)
f_pass_seria = ALLTRIM(m.pass_seria)
f_pass_num = ALLTRIM(m.pass_num)
f_pass_data = m.pass_data
f_pass_mesto = ALLTRIM(m.pass_mesto)
f_data = m.data
f_ima_pk = ALLTRIM(m.ima_pk)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

stroka_sql = "EXECUTE export_account"

stroka_sql = stroka_sql + " ?f_branch, ?f_ref_client, ?f_ref_card, ?f_ref_acc, ?f_ref_ost, ?f_date_card, ?f_card_num, ?f_vid_card, ?f_kod_proekt, ?f_ref_poisk, ?f_del_card,"
stroka_sql = stroka_sql + " ?f_date_nls, ?f_date_close, ?f_pr_nls, ?f_pr_odb, ?f_resident, ?f_card_nls, ?f_card_acct, ?f_vklad_acct, ?f_status_tra, ?f_status_mak, ?f_fio_rus,"
stroka_sql = stroka_sql + " ?f_fio_eng, ?f_kod_slovo, ?f_tip_cln, ?f_type_card, ?f_num_card, ?f_val_card, ?f_name_card, ?f_date_ostat, ?f_read_ost_p, ?f_begin_bal, ?f_debit, ?f_credit, ?f_end_bal,"
stroka_sql = stroka_sql + " ?f_read_ost_m, ?f_pr_izm_ost, ?f_vxd_ost_m, ?f_debit_m, ?f_credit_m, ?f_isx_ost_m, ?f_card_strah, ?f_tran_42301, ?f_tran_42308, ?f_num_dog, ?f_data_dog,"
stroka_sql = stroka_sql + " ?f_data_proz, ?f_stavka, ?f_num_47411, ?f_pass_seria, ?f_pass_num, ?f_pass_data, ?f_pass_mesto, ?f_data, ?f_ima_pk"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

export_data = SQLEXEC(sql_num, stroka_sql, '')

DO WHILE SQLMORERESULTS(sql_num) < 2
ENDDO

RETURN


******************************************************************* PROCEDURE export_acc_del ********************************************************

PROCEDURE export_acc_del
PARAMETERS name_sql_server

SET ESCAPE ON

IF FILE((put_dbf) + ('acc_del.dbf'))

  IF NOT USED('acc_del')

    SELECT 0
    USE (put_dbf) + ('acc_del.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT acc_del
  SET ORDER TO fio_rus

  COUNT TO colvo_zap

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт справочника по закрытым картам  в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT acc_del
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN

      time_home = SECONDS()

      SCATTER MEMVAR MEMO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      DO scan_acc_del IN Export_data_sql
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Филиал № ' + ALLTRIM(m.branch) + ' Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT acc_del
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт справочника по закрытым картам в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('acc_del.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


************************************************************************ PROCEDURE scan_acc_del ********************************************************

PROCEDURE scan_acc_del

f_branch = ALLTRIM(m.branch)
f_ref_client = ALLTRIM(m.ref_client)
f_ref_card = m.ref_card
f_ref_acc = m.ref_acc
f_ref_ost = m.ref_ost
f_date_card = m.date_card
f_card_num = ALLTRIM(m.card_num)
f_vid_card = ALLTRIM(m.vid_card)
f_kod_proekt = ALLTRIM(m.kod_proekt)
f_ref_poisk = m.ref_poisk
f_del_card = m.del_card
f_date_nls = m.date_nls
f_date_close = m.date_close
f_pr_nls = m.pr_nls
f_pr_odb = m.pr_odb
f_resident = ALLTRIM(m.resident)
f_card_nls = ALLTRIM(m.card_nls)
f_card_acct = ALLTRIM(m.card_acct)
f_vklad_acct = ALLTRIM(m.vklad_acct)
f_status_tra = ALLTRIM(m.status_tra)
f_status_mak = ALLTRIM(m.status_mak)
f_fio_rus = ALLTRIM(m.fio_rus)
f_fio_eng = ALLTRIM(m.fio_eng)
f_kod_slovo = ALLTRIM(m.kod_slovo)
f_tip_cln = ALLTRIM(m.tip_cln)
f_type_card = ALLTRIM(m.type_card)
f_num_card = ALLTRIM(m.num_card)
f_val_card = ALLTRIM(m.val_card)
f_name_card = ALLTRIM(m.name_card)
f_date_ostat = m.date_ostat
f_read_ost_p = m.read_ost_p
f_begin_bal = m.begin_bal
f_debit = m.debit
f_credit = m.credit
f_end_bal = m.end_bal
f_read_ost_m = m.read_ost_m
f_pr_izm_ost = m.pr_izm_ost
f_vxd_ost_m = m.vxd_ost_m
f_debit_m = m.debit_m
f_credit_m = m.credit_m
f_isx_ost_m = m.isx_ost_m
f_card_strah = ALLTRIM(m.card_strah)
f_tran_42301 = ALLTRIM(m.tran_42301)
f_tran_42308 = ALLTRIM(m.tran_42308)
f_num_dog = ALLTRIM(m.num_dog)
f_data_dog = m.data_dog
f_data_proz = m.data_proz
f_stavka = m.stavka
f_num_47411 = ALLTRIM(m.num_47411)
f_pass_seria = ALLTRIM(m.pass_seria)
f_pass_num = ALLTRIM(m.pass_num)
f_pass_data = m.pass_data
f_pass_mesto = ALLTRIM(m.pass_mesto)
f_data = m.data
f_ima_pk = ALLTRIM(m.ima_pk)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

stroka_sql = "EXECUTE export_acc_del"

stroka_sql = stroka_sql + " ?f_branch, ?f_ref_client, ?f_ref_card, ?f_ref_acc, ?f_ref_ost, ?f_date_card, ?f_card_num, ?f_vid_card, ?f_kod_proekt, ?f_ref_poisk, ?f_del_card,"
stroka_sql = stroka_sql + " ?f_date_nls, ?f_date_close, ?f_pr_nls, ?f_pr_odb, ?f_resident, ?f_card_nls, ?f_card_acct, ?f_vklad_acct, ?f_status_tra, ?f_status_mak, ?f_fio_rus,"
stroka_sql = stroka_sql + " ?f_fio_eng, ?f_kod_slovo, ?f_tip_cln, ?f_type_card, ?f_num_card, ?f_val_card, ?f_name_card, ?f_date_ostat, ?f_read_ost_p, ?f_begin_bal, ?f_debit, ?f_credit, ?f_end_bal,"
stroka_sql = stroka_sql + " ?f_read_ost_m, ?f_pr_izm_ost, ?f_vxd_ost_m, ?f_debit_m, ?f_credit_m, ?f_isx_ost_m, ?f_card_strah, ?f_tran_42301, ?f_tran_42308, ?f_num_dog, ?f_data_dog,"
stroka_sql = stroka_sql + " ?f_data_proz, ?f_stavka, ?f_num_47411, ?f_pass_seria, ?f_pass_num, ?f_pass_data, ?f_pass_mesto, ?f_data, ?f_ima_pk"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

export_data = SQLEXEC(sql_num, stroka_sql, '')

DO WHILE SQLMORERESULTS(sql_num) < 2
ENDDO

RETURN


******************************************************************* PROCEDURE export_vedom_ost *********************************************************

PROCEDURE export_vedom_ost
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('vedom_ost.dbf'))

  IF NOT USED('vedom_ost')

    SELECT 0
    USE (put_dbf) + ('vedom_ost.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT vedom_ost

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('vedom_ost')
        SELECT vedom_ost
        USE
        USE (put_dbf) + ('vedom_ost.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('vedom_ost.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(date_vedom) + ALLTRIM(fio_rus) TAG date_fio

      SELECT vedom_ost
      USE
      USE (put_dbf) + ('vedom_ost.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(date_vedom, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт ведомости остатков из ПО "CARD-виза" в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT vedom_ost
* SET FILTER TO BETWEEN(date_vedom, date_home, date_end)
    GO TOP

    SCAN FOR BETWEEN(date_vedom, date_home, date_end)

      SCATTER MEMVAR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      DO scan_vedom_ost IN Export_data_sql
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Филиал № ' + ALLTRIM(m.branch) + ' Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + ;
        ' Дата операции - ' + DTOC(m.date_vedom) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT vedom_ost
    ENDSCAN

    SELECT vedom_ost
    SET FILTER TO

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт ведомости остатков из ПО "CARD-виза" в СУБД SQL Server успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('vedom_ost.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


****************************************************************** PROCEDURE scan_vedom_ost ***********************************************************

PROCEDURE scan_vedom_ost

f_branch = ALLTRIM(m.branch)
f_ref_client = ALLTRIM(m.ref_client)
f_val_card = ALLTRIM(m.val_card)
f_fio_rus = ALLTRIM(m.fio_rus)

f_card_nls = ALLTRIM(m.card_nls)
f_card_acct = ALLTRIM(m.card_acct)

f_date_vedom = m.date_vedom
f_date_ostat = m.date_ostat

f_begin_bal = m.begin_bal
f_debit = m.debit
f_credit = m.credit
f_end_bal = m.end_bal

f_pr_izm_ost = m.pr_izm_ost
f_date_ost_m = m.date_ost_m

f_vxd_ost_m = m.vxd_ost_m
f_debit_m = m.debit_m
f_credit_m = m.credit_m
f_isx_ost_m = m.isx_ost_m

time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

stroka_sql = "EXECUTE export_vedom_ost"

stroka_sql = stroka_sql + " ?f_branch, ?f_ref_client, ?f_val_card, ?f_fio_rus, ?f_card_nls, ?f_card_acct,"
stroka_sql = stroka_sql + " ?f_date_vedom, ?f_date_ostat, ?f_begin_bal, ?f_debit, ?f_credit, ?f_end_bal,"
stroka_sql = stroka_sql + " ?f_pr_izm_ost, ?f_date_ost_m, ?f_vxd_ost_m, ?f_debit_m, ?f_credit_m, ?f_isx_ost_m"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

export_data = SQLEXEC(sql_num, stroka_sql, '')

DO WHILE SQLMORERESULTS(sql_num) < 2
ENDDO

RETURN


****************************************************************** PROCEDURE export_istor_ost **********************************************************

PROCEDURE export_istor_ost
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('istor_ost.dbf'))

  IF NOT USED('istor_ost')

    SELECT 0
    USE (put_dbf) + ('istor_ost.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT istor_ost

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('istor_ost')
        SELECT istor_ost
        USE
        USE (put_dbf) + ('istor_ost.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('istor_ost.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(date_ost_p) + ALLTRIM(fio_rus) TAG date_fio

      SELECT istor_ost
      USE
      USE (put_dbf) + ('istor_ost.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(date_ost_p, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт истории остатков из ПО "Трансмастер" в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT istor_ost
* SET FILTER TO BETWEEN(date_ost_p, date_home, date_end)
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR BETWEEN(date_ost_p, date_home, date_end)

      SCATTER MEMVAR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      DO scan_istor_ost IN Export_data_sql
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Филиал № ' + ALLTRIM(m.branch) + ' Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + ;
        ' Дата операции - ' + DTOC(m.date_ost_p) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT istor_ost
    ENDSCAN

    SELECT istor_ost
    SET FILTER TO

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт истории остатков из ПО "Трансмастер" в СУБД SQL Server успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('istor_ost.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


********************************************************************* PROCEDURE scan_istor_ost **********************************************************

PROCEDURE scan_istor_ost

f_branch = ALLTRIM(m.branch)
f_ref_client = ALLTRIM(m.ref_client)
f_ref_acc = m.ref_acc
f_ref_ost = m.ref_ost
f_val_card = ALLTRIM(m.val_card)
f_fio_rus = ALLTRIM(m.fio_rus)

f_card_nls = ALLTRIM(m.card_nls)
f_card_acct = ALLTRIM(m.card_acct)

f_date_ost_p = m.date_ost_p

f_begin_bal = m.begin_bal
f_debit = m.debit
f_credit = m.credit
f_end_bal = m.end_bal

time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

stroka_sql = "EXECUTE export_istor_ost"

stroka_sql = stroka_sql + " ?f_branch, ?f_ref_client, ?f_ref_acc, ?f_ref_ost, ?f_val_card, ?f_fio_rus, ?f_date_ost_p,"
stroka_sql = stroka_sql + " ?f_card_nls, ?f_card_acct, ?f_begin_bal, ?f_debit, ?f_credit, ?f_end_bal"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

export_data = SQLEXEC(sql_num, stroka_sql, '')

DO WHILE SQLMORERESULTS(sql_num) < 2
ENDDO

RETURN


****************************************************************** PROCEDURE export_istor_mak **********************************************************

PROCEDURE export_istor_mak
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('istor_mak.dbf'))

  IF NOT USED('istor_mak')

    SELECT 0
    USE (put_dbf) + ('istor_mak.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT istor_mak

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('istor_mak')
        SELECT istor_mak
        USE
        USE (put_dbf) + ('istor_mak.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('istor_mak.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(date_ost_m) + ALLTRIM(fio_rus) TAG date_fio

      SELECT istor_mak
      USE
      USE (put_dbf) + ('istor_mak.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(date_ost_m, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт истории остатков из ПО "CARD-виза" в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT istor_mak
* SET FILTER TO BETWEEN(date_ost_m, date_home, date_end)
    GO TOP

    SCAN FOR BETWEEN(date_ost_m, date_home, date_end)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SCATTER MEMVAR

      f_branch = ALLTRIM(m.branch)
      f_date_ost_m = m.date_ost_m
      f_ref_client = ALLTRIM(m.ref_client)
      f_ref_acc = m.ref_acc
      f_ref_ost = m.ref_ost
      f_val_card = ALLTRIM(m.val_card)
      f_fio_rus = ALLTRIM(m.fio_rus)

      f_card_nls = ALLTRIM(m.card_nls)
      f_card_acct = ALLTRIM(m.card_acct)

      f_vxd_ost_m = m.vxd_ost_m
      f_debit_m = m.debit_m
      f_credit_m = m.credit_m
      f_isx_ost_m = m.isx_ost_m

      time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      stroka_sql = "EXECUTE export_istor_mak"

      stroka_sql = stroka_sql + " ?f_branch, ?f_date_ost_m, ?f_ref_client, ?f_ref_acc, ?f_ref_ost, ?f_val_card, ?f_fio_rus,"
      stroka_sql = stroka_sql + " ?f_card_nls, ?f_card_acct, ?f_vxd_ost_m, ?f_debit_m, ?f_credit_m, ?f_isx_ost_m"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      export_data = SQLEXEC(sql_num, stroka_sql, '')

      DO WHILE SQLMORERESULTS(sql_num) < 2
      ENDDO

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Филиал № ' + ALLTRIM(m.branch) + ' Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + ;
        ' Дата операции - ' + DTOC(m.date_ost_m) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT istor_mak
    ENDSCAN

    SELECT istor_mak
    SET FILTER TO

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт истории остатков из ПО "CARD-виза" в СУБД SQL Server успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('istor_mak.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


****************************************************************** PROCEDURE export_docum_slip *********************************************************

PROCEDURE export_docum_slip
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('docum_slip.dbf'))

  IF NOT USED('docum_slip')

    SELECT 0
    USE (put_dbf) + ('docum_slip.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())
  =INKEY(2)
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  @ WROWS()/3,3 SAY PADC('Начата индексация исходной таблицы с данными ... ',WCOLS())

  IF USED('docum_slip')

    SELECT docum_slip
    USE
    USE (put_dbf) + ('docum_slip.dbf') EXCLUSIVE

    DO pack_docum_slip IN Pack_table

    SELECT docum_slip
    USE
    USE (put_dbf) + ('docum_slip.dbf') SHARE

  ENDIF

  @ WROWS()/3,3 SAY PADC('Индексация исходной таблицы завершена ... ',WCOLS())

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  SELECT docum_slip
  SET ORDER TO date_fio
  GO TOP

  COUNT TO colvo_zap FOR BETWEEN(date_prov, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт проведенных документов из ПО "Трансмастер" в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT docum_slip
* SET FILTER TO BETWEEN(date_prov, date_home, date_end)
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN  FOR BETWEEN(date_prov, date_home, date_end) AND EOF() = .F.

      time_home = SECONDS()

      SCATTER MEMVAR MEMO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
      DO scan_docum_slip IN Export_data_sql
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
      WAIT WINDOW 'Филиал № ' + ALLTRIM(m.branch) + ' Клиент - ' + ALLTRIM(m.fio_rus) + ' Дата операции - ' + DTOC(m.date_prov) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      SELECT docum_slip
    ENDSCAN

    SELECT docum_slip
    SET FILTER TO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт проведенных документов из ПО "Трансмастер"  в СУБД SQL Server успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Запрашиваемых Вами данных за диапазон дат с ' + DTOC(date_home) + ' по ' + DTOC(date_end) + ' не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('docum_slip.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


******************************************************************** PROCEDURE scan_docum_slip *********************************************************

PROCEDURE scan_docum_slip

f_branch = ALLTRIM(m.branch)
f_record = m.record
f_deal_desc = ALLTRIM(m.deal_desc)
f_date_oper = m.date_oper
f_date_bank = m.date_bank
f_date_prov = m.date_prov
f_ref_client = ALLTRIM(m.ref_client)
f_fio_rus = ALLTRIM(m.fio_rus)
f_card_num = ALLTRIM(m.card_num)
f_kod_oper = ALLTRIM(m.kod_oper)
f_del_kod = m.del_kod
f_card_nls = ALLTRIM(m.card_nls)
f_card_acct = ALLTRIM(m.card_acct)
f_val_tran = ALLTRIM(m.val_tran)
f_sum_tran = m.sum_tran
f_type_oper = ALLTRIM(m.type_oper)
f_kurs = m.kurs
f_nls_deb = ALLTRIM(m.nls_deb)
f_sum_deb = m.sum_deb
f_val_deb = ALLTRIM(m.val_deb)
f_nls_cred = ALLTRIM(m.nls_cred)
f_sum_cred = m.sum_cred
f_val_cred = ALLTRIM(m.val_cred)
f_type_card = ALLTRIM(m.type_card)
f_num_card = ALLTRIM(m.num_card)
f_val_card = ALLTRIM(m.val_card)
f_name_card = ALLTRIM(m.name_card)
f_card_strah = ALLTRIM(m.card_strah)
f_tran_42301 = ALLTRIM(m.tran_42301)
f_tran_42308 = ALLTRIM(m.tran_42308)
f_tran_name = ALLTRIM(m.tran_name)
f_tip_tran = ALLTRIM(m.tip_tran)
f_name_term = ALLTRIM(m.name_term)
f_city = ALLTRIM(m.city)
f_kod_name = ALLTRIM(m.kod_name)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

stroka_sql = "EXECUTE export_docum_slip"

stroka_sql = stroka_sql + " ?f_branch, ?f_record, ?f_deal_desc, ?f_date_oper, ?f_date_bank, ?f_date_prov, ?f_ref_client, ?f_fio_rus, ?f_card_num, ?f_kod_oper, ?f_del_kod,"
stroka_sql = stroka_sql + " ?f_card_nls, ?f_card_acct, ?f_val_tran, ?f_sum_tran, ?f_type_oper, ?f_kurs, ?f_nls_deb, ?f_sum_deb, ?f_val_deb, ?f_nls_cred, ?f_sum_cred,"
stroka_sql = stroka_sql + " ?f_val_cred, ?f_type_card, ?f_num_card, ?f_val_card, ?f_name_card, ?f_card_strah, ?f_tran_42301, ?f_tran_42308,"
stroka_sql = stroka_sql + " ?f_tran_name, ?f_tip_tran, ?f_name_term, ?f_city, ?f_kod_name"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

export_data = SQLEXEC(sql_num, stroka_sql, '')

DO WHILE SQLMORERESULTS(sql_num) < 2
ENDDO

RETURN


****************************************************************** PROCEDURE export_docum_mak *********************************************************

PROCEDURE export_docum_mak
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('docum_mak.dbf'))

  IF NOT USED('docum_mak')

    SELECT 0
    USE (put_dbf) + ('docum_mak.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  @ WROWS()/3,3 SAY PADC('Начата индексация исходной таблицы с данными ... ',WCOLS())

  IF USED('docum_mak')

    SELECT docum_mak
    USE
    USE (put_dbf) + ('docum_mak.dbf') EXCLUSIVE

    DO pack_docum_mak IN Pack_table

    SELECT docum_mak
    USE
    USE (put_dbf) + ('docum_mak.dbf') SHARE

  ENDIF

  @ WROWS()/3,3 SAY PADC('Индексация исходной таблицы завершена ... ',WCOLS())

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  SELECT docum_mak
  SET ORDER TO date_fio
  GO TOP

  COUNT TO colvo_zap FOR BETWEEN(TTOD(date_oper), date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт проведенных документов из ПО "CARD-виза"  в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT docum_mak
* SET FILTER TO BETWEEN(TTOD(date_oper), date_home, date_end)
    GO TOP

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR BETWEEN(TTOD(date_oper), date_home, date_end) AND EOF() = .F.

      SCATTER MEMVAR MEMO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
      DO scan_docum_mak IN Export_data_sql
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
      WAIT WINDOW 'Филиал № ' + ALLTRIM(m.branch) + ' Клиент - ' + ALLTRIM(m.fio_rus) + ' Дата операции - ' + DTOC(TTOD(m.date_oper)) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      SELECT docum_mak
    ENDSCAN

    SELECT docum_mak
    SET FILTER TO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт проведенных документов из ПО "CARD-виза" в СУБД SQL Server успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Запрашиваемых Вами данных за диапазон дат с ' + DTOC(date_home) + ' по ' + DTOC(date_end) + ' не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('docum_mak.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


***************************************************************** PROCEDURE scan_docum_mak ***********************************************************

PROCEDURE scan_docum_mak

f_branch = ALLTRIM(m.branch)
f_record = ALLTRIM(m.record)
f_date_oper = m.date_oper
f_ref_client = ALLTRIM(m.ref_client)
f_fio_rus = ALLTRIM(m.fio_rus)
f_card_num = ALLTRIM(m.card_num)
f_kod_oper = ALLTRIM(m.kod_oper)
f_card_acct = ALLTRIM(m.card_acct)
f_nls_deb = ALLTRIM(m.nls_deb)
f_sum_deb = m.sum_deb
f_val_deb = ALLTRIM(m.val_deb)
f_nls_cred = ALLTRIM(m.nls_cred)
f_sum_cred = m.sum_cred
f_val_cred = ALLTRIM(m.val_cred)
f_type_card = ALLTRIM(m.type_card)
f_num_card = ALLTRIM(m.num_card)
f_val_card = ALLTRIM(m.val_card)
f_name_card = ALLTRIM(m.name_card)
f_card_strah = ALLTRIM(m.card_strah)
f_tran_42301 = ALLTRIM(m.tran_42301)
f_tran_42308 = ALLTRIM(m.tran_42308)
f_tran_name = ALLTRIM(m.tran_name)
f_tip_tran = ALLTRIM(m.tip_tran)
f_kod_name = ALLTRIM(m.kod_name)

time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

stroka_sql = "EXECUTE export_docum_mak"

stroka_sql = stroka_sql + " ?f_branch, ?f_record, ?f_date_oper, ?f_ref_client, ?f_fio_rus, ?f_card_num, ?f_kod_oper,"
stroka_sql = stroka_sql + " ?f_card_acct, ?f_nls_deb, ?f_sum_deb, ?f_val_deb, ?f_nls_cred, ?f_sum_cred,"
stroka_sql = stroka_sql + " ?f_val_cred, ?f_type_card, ?f_num_card, ?f_val_card, ?f_name_card, ?f_card_strah, ?f_tran_42301, ?f_tran_42308,"
stroka_sql = stroka_sql + " ?f_tran_name, ?f_tip_tran, ?f_kod_name"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

export_data = SQLEXEC(sql_num, stroka_sql, '')

DO WHILE SQLMORERESULTS(sql_num) < 2
ENDDO

RETURN


****************************************************************** PROCEDURE export_acc_cln_dt **********************************************************

PROCEDURE export_acc_cln_dt
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('acc_cln_dt.dbf'))

  IF NOT USED('acc_cln_dt')

    SELECT 0
    USE (put_dbf) + ('acc_cln_dt.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT acc_cln_dt

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('acc_cln_dt')
        SELECT acc_cln_dt
        USE
        USE (put_dbf) + ('acc_cln_dt.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('acc_cln_dt.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(data_pak) + ALLTRIM(fio_rus) + DTOS(data_home) + DTOS(data_end) TAG date_fio

      SELECT acc_cln_dt
      USE
      USE (put_dbf) + ('acc_cln_dt.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(data_home, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт расчета процентов по дням в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT acc_cln_dt
* SET FILTER TO BETWEEN(data_home, date_home, date_end)
    GO TOP

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR  BETWEEN(data_home, date_home, date_end)

      SCATTER MEMVAR

      f_branch = ALLTRIM(m.branch)
      f_ref_client = ALLTRIM(m.ref_client)
      f_ref_card = m.ref_card
      f_ref_ost = ALLTRIM(m.ref_ost)
      f_num_dog = ALLTRIM(m.num_dog)
      f_data_dog = m.data_dog
      f_card_nls = ALLTRIM(m.card_nls)
      f_val_card = ALLTRIM(m.val_card)
      f_fio_rus = ALLTRIM(m.fio_rus)
      f_pr = m.pr
      f_data_proz = m.data_proz
      f_dt_hom_end = m.dt_hom_end
      f_data_home = m.data_home
      f_data_end = m.data_end
      f_data_pak = m.data_pak
      f_colvo_den = m.colvo_den
      f_kvartal = m.kvartal
      f_mes_kvart = m.mes_kvart
      f_kurs = m.kurs
      f_stavka = m.stavka
      f_sum_ostat = m.sum_ostat
      f_sum_dobav = m.sum_dobav
      f_sum_rasch = m.sum_rasch
      f_sum_rasxod = m.sum_rasxod
      f_sum_nalog = m.sum_nalog
      f_sum_viplat = m.sum_viplat

      time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      stroka_sql = "EXECUTE export_acc_cln_dt"

      stroka_sql = stroka_sql + " ?f_branch, ?f_ref_client ,?f_ref_card, ?f_ref_ost, ?f_num_dog, ?f_data_dog, ?f_fio_rus, ?f_card_nls, ?f_val_card, ?f_pr,"
      stroka_sql = stroka_sql + " ?f_data_proz, ?f_dt_hom_end, ?f_data_home, ?f_data_end, ?f_data_pak, ?f_colvo_den, ?f_kvartal, ?f_mes_kvart, ?f_kurs, ?f_stavka,"
      stroka_sql = stroka_sql + " ?f_sum_ostat, ?f_sum_dobav, ?f_sum_rasch, ?f_sum_rasxod, ?f_sum_nalog, ?f_sum_viplat"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      export_data = SQLEXEC(sql_num, stroka_sql, '')

      DO WHILE SQLMORERESULTS(sql_num) < 2
      ENDDO

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Филиал № ' + ALLTRIM(m.branch) + ' Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_nls) + CHR(13) + ;
        ' Дата начала операции - ' + DTOC(m.data_home) + ' Дата окончания операции - ' + DTOC(m.data_end) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT acc_cln_dt
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SELECT acc_cln_dt
    SET FILTER TO

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт таблицы по процентам в СУБД SQL Server успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('acc_cln_dt.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


******************************************************************* PROCEDURE export_acc_cln_sum ********************************************************

PROCEDURE export_acc_cln_sum
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('acc_cln_sum.dbf'))

  IF NOT USED('acc_cln_sum')

    SELECT 0
    USE (put_dbf) + ('acc_cln_sum.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT acc_cln_sum

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('acc_cln_sum')
        SELECT acc_cln_sum
        USE
        USE (put_dbf) + ('acc_cln_sum.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('acc_cln_sum.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(data_pak) + ALLTRIM(fio_rus) + DTOS(data_home) + DTOS(data_end) TAG date_fio

      SELECT acc_cln_sum
      USE
      USE (put_dbf) + ('acc_cln_sum.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(data_home, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт расчета процентов свернутого по дням в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT acc_cln_sum
* SET FILTER TO BETWEEN(data_home, date_home, date_end)
    GO TOP

* BROWSE WINDOW brows

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR  BETWEEN(data_home, date_home, date_end)

      SCATTER MEMVAR

      f_branch = ALLTRIM(m.branch)
      f_ref_client = ALLTRIM(m.ref_client)
      f_ref_card = m.ref_card
      f_ref_ost = ALLTRIM(m.ref_ost)
      f_num_dog = ALLTRIM(m.num_dog)
      f_data_dog = m.data_dog
      f_card_nls = ALLTRIM(m.card_nls)
      f_val_card = ALLTRIM(m.val_card)
      f_fio_rus = ALLTRIM(m.fio_rus)
      f_pr = m.pr
      f_data_proz = m.data_proz
      f_dt_hom_end = m.dt_hom_end
      f_data_home = m.data_home
      f_data_end = m.data_end
      f_data_pak = m.data_pak
      f_den_mes = m.den_mes
      f_kvartal = m.kvartal
      f_mes_kvart = m.mes_kvart
      f_kurs = m.kurs
      f_stavka = m.stavka
      f_sum_ostat = m.sum_ostat
      f_sum_dobav = m.sum_dobav
      f_sum_rasch = m.sum_rasch
      f_sum_rasxod = m.sum_rasxod
      f_sum_nalog = m.sum_nalog
      f_sum_viplat = m.sum_viplat

      time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      stroka_sql = "EXECUTE export_acc_cln_sum"

      stroka_sql = stroka_sql + " ?f_branch, ?f_ref_client ,?f_ref_card, ?f_ref_ost, ?f_num_dog, ?f_data_dog, ?f_fio_rus, ?f_card_nls, ?f_val_card, ?f_pr,"
      stroka_sql = stroka_sql + " ?f_data_proz, ?f_dt_hom_end, ?f_data_home, ?f_data_end, ?f_data_pak, ?f_den_mes, ?f_kvartal, ?f_mes_kvart, ?f_kurs, ?f_stavka,"
      stroka_sql = stroka_sql + " ?f_sum_ostat, ?f_sum_dobav, ?f_sum_rasch, ?f_sum_rasxod, ?f_sum_nalog, ?f_sum_viplat"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      export_data = SQLEXEC(sql_num, stroka_sql, '')

      DO WHILE SQLMORERESULTS(sql_num) < 2
      ENDDO

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Филиал № ' + ALLTRIM(m.branch) + ' Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_nls) + CHR(13) + ;
        ' Дата начала операции - ' + DTOC(m.data_home) + ' Дата окончания операции - ' + DTOC(m.data_end) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT acc_cln_sum
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SELECT acc_cln_sum
    SET FILTER TO

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт таблицы по процентам свернутого по дням в СУБД SQL Server успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('acc_cln_sum.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


******************************************************************* PROCEDURE export_popol_arx ********************************************************

PROCEDURE export_popol_arx
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('popol_arx.dbf'))

  IF NOT USED('popol_arx')

    SELECT 0
    USE (put_dbf) + ('popol_arx.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT popol_arx

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('popol_arx')
        SELECT popol_arx
        USE
        USE (put_dbf) + ('popol_arx.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('popol_arx.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(data_val) + ALLTRIM(fio_rus) TAG date_fio

      SELECT popol_arx
      USE
      USE (put_dbf) + ('popol_arx.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(data_val, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт составленных документов по управлению карточным счетом в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT popol_arx
* SET FILTER TO BETWEEN(data_val, date_home, date_end)
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR  BETWEEN(data_val, date_home, date_end)

      SCATTER MEMVAR MEMO

      f_record = ALLTRIM(m.record)
      f_pril = ALLTRIM(m.pril)
      f_export = m.export
      f_zapis = m.zapis
      f_kod_proekt = ALLTRIM(m.kod_proekt)
      f_data_val = m.data_val
      f_nn_doc = m.nn_doc
      f_kodval = ALLTRIM(m.kodval)
      f_kod_oper = ALLTRIM(m.kod_oper)
      f_fio_rus = ALLTRIM(m.fio_rus)
      f_card_acct = ALLTRIM(m.card_acct)
      f_debit = ALLTRIM(m.debit)
      f_credit = ALLTRIM(m.credit)
      f_summa = m.summa
      f_summa_prop = ALLTRIM(m.summa_prop)
      f_prim = ALLTRIM(m.prim)
      f_data_tran = m.data_tran
      f_ima_pk = ALLTRIM(m.ima_pk)
      f_num_oper = ALLTRIM(m.num_oper)

      time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      stroka_sql = "EXECUTE export_popol_arx"

      stroka_sql = stroka_sql + " ?f_record, ?f_pril, ?f_export, ?f_zapis, ?f_kod_proekt, ?f_data_val, ?f_nn_doc, ?f_kodval, ?f_kod_oper, ?f_fio_rus, ?f_card_acct,"
      stroka_sql = stroka_sql + " ?f_debit, ?f_credit, ?f_summa, ?f_summa_prop, ?f_prim, ?f_data_tran, ?f_ima_pk, ?f_num_oper"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      export_data = SQLEXEC(sql_num, stroka_sql, '')

      DO WHILE SQLMORERESULTS(sql_num) < 2
      ENDDO

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + ' Дата документа - ' + DTOC(data_val) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT popol_arx
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт документов по управлению карточным счетом в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('popol_arx.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


******************************************************************* PROCEDURE export_ostat_upk ********************************************************

PROCEDURE export_ostat_upk
PARAMETERS name_sql_server

SET ESCAPE ON

IF FILE((put_dbf) + ('ostat_upk.dbf'))

  IF NOT USED('ostat_upk')

    SELECT 0
    USE (put_dbf) + ('ostat_upk.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT ostat_upk

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'card_acct'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('ostat_upk')
        SELECT ostat_upk
        USE
        USE (put_dbf) + ('ostat_upk.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('ostat_upk.dbf') EXCLUSIVE
      ENDIF

      INDEX ON ALLTRIM(card_acct) TAG card_acct

      SELECT ostat_upk
      USE
      USE (put_dbf) + ('ostat_upk.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO card_acct

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт сравнительной ведомости остатков в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT ostat_upk
    GO TOP

    SCAN

      time_home = SECONDS()

      SCATTER MEMVAR MEMO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      DO scan_ostat_upk IN Export_data_sql
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Номер счета - ' + ALLTRIM(m.card_nls) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT ostat_upk
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт сравнительной ведомости остатков в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('ostat_upk.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


********************************************************************** PROCEDURE scan_ostat_upk ********************************************************

PROCEDURE scan_ostat_upk

f_branch = ALLTRIM(m.branch)
f_date_ostat = m.date_ostat
f_ref_client = ALLTRIM(m.ref_client)
f_ref_acc = m.ref_acc
f_card_nls = ALLTRIM(m.card_nls)
f_card_acct = ALLTRIM(m.card_acct)
f_val_card = ALLTRIM(m.val_card)
f_pr_prover = m.pr_prover
f_begin_bal = m.begin_bal
f_vxd_ostat = m.vxd_ostat
f_debit = m.debit
f_debit_f = m.debit_f
f_credit = m.credit
f_credit_f = m.credit_f
f_end_bal = m.end_bal
f_isx_ostat = m.isx_ostat

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

stroka_sql = "EXECUTE export_ostat_upk"

stroka_sql = stroka_sql + " ?f_branch, ?f_date_ostat, ?f_ref_client, ?f_ref_acc, ?f_card_nls, ?f_card_acct, ?f_val_card, ?f_pr_prover,"
stroka_sql = stroka_sql + " ?f_begin_bal, ?f_vxd_ostat, ?f_debit, ?f_debit_f, ?f_credit, ?f_credit_f, ?f_end_bal, ?f_isx_ostat"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

export_data = SQLEXEC(sql_num, stroka_sql, '')

DO WHILE SQLMORERESULTS(sql_num) < 2
ENDDO

RETURN


******************************************************************* PROCEDURE export_memordrub *********************************************************

PROCEDURE export_memordrub
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('memordrub.dbf'))

  IF NOT USED('memordrub')

    SELECT 0
    USE (put_dbf) + ('memordrub.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT memordrub

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('memordrub')
        SELECT memordrub
        USE
        USE (put_dbf) + ('memordrub.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('memordrub.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(data_doc) + ALLTRIM(fio_rus) TAG date_fio

      SELECT memordrub
      USE
      USE (put_dbf) + ('memordrub.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(data_doc, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт составленных рублевых мемориальных ордеров в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT memordrub
* SET FILTER TO BETWEEN(data_doc, date_home, date_end)
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR  BETWEEN(data_doc, date_home, date_end)

      SCATTER MEMVAR MEMO

      f_record = ALLTRIM(m.record)
      f_kodval = m.kodval
      f_vid = ALLTRIM(m.vid)
      f_kod = ALLTRIM(m.kod)
      f_kod_oper = ALLTRIM(m.kod_oper)
      f_num_ord = m.num_ord
      f_kod_proekt = ALLTRIM(m.kod_proekt)
      f_card_acct = ALLTRIM(m.card_acct)
      f_fio_rus = ALLTRIM(m.fio_rus)
      f_data_doc = m.data_doc
      f_debit = ALLTRIM(m.debit)
      f_name_dbt = ALLTRIM(m.name_dbt)
      f_credit = ALLTRIM(m.credit)
      f_name_crt = ALLTRIM(m.name_crt)
      f_summa = m.summa
      f_summa_prop = ALLTRIM(m.summa_prop)
      f_naznach = ALLTRIM(m.naznach)
      f_osnovanie = ALLTRIM(m.osnovanie)
      f_data_tran = m.data_tran
      f_ima_pk = ALLTRIM(m.ima_pk)
      f_num_oper = ALLTRIM(m.num_oper)

      time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      stroka_sql = "EXECUTE export_memordrub"

      stroka_sql = stroka_sql + " ?f_record, ?f_kodval, ?f_vid, ?f_kod, ?f_kod_oper, ?f_num_ord, ?f_kod_proekt, ?f_card_acct, ?f_fio_rus, ?f_data_doc,"
      stroka_sql = stroka_sql + " ?f_debit, ?f_name_dbt, ?f_credit, ?f_name_crt, ?f_summa, ?f_summa_prop, ?f_naznach, ?f_osnovanie, ?f_data_tran, ?f_ima_pk, ?f_num_oper"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      export_data = SQLEXEC(sql_num, stroka_sql, '')

      DO WHILE SQLMORERESULTS(sql_num) < 2
      ENDDO

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + ' Дата документа - ' + DTOC(data_doc) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT memordrub
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт рублевых мемориальных ордеров в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('memordrub.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


****************************************************************** PROCEDURE export_memordval *********************************************************

PROCEDURE export_memordval
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('memordval.dbf'))

  IF NOT USED('memordval')

    SELECT 0
    USE (put_dbf) + ('memordval.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT memordval

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('memordval')
        SELECT memordval
        USE
        USE (put_dbf) + ('memordval.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('memordval.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(data_doc) + ALLTRIM(fio_rus) TAG date_fio

      SELECT memordval
      USE
      USE (put_dbf) + ('memordval.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(data_doc, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт составленных валютных мемориальных ордеров в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT memordval
* SET FILTER TO BETWEEN(data_doc, date_home, date_end)
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR  BETWEEN(data_doc, date_home, date_end)

      SCATTER MEMVAR MEMO

      f_record = ALLTRIM(m.record)
      f_kodval = m.kodval
      f_kurs_zb = m.kurs_zb
      f_vid = ALLTRIM(m.vid)
      f_kod = ALLTRIM(m.kod)
      f_kod_oper = ALLTRIM(m.kod_oper)
      f_num_ord = m.num_ord
      f_kod_proekt = ALLTRIM(m.kod_proekt)
      f_card_acct = ALLTRIM(m.card_acct)
      f_fio_rus = ALLTRIM(m.fio_rus)
      f_data_doc = m.data_doc
      f_debit = ALLTRIM(m.debit)
      f_name_dbt = ALLTRIM(m.name_dbt)
      f_credit = ALLTRIM(m.credit)
      f_name_crt = ALLTRIM(m.name_crt)
      f_summa_val = m.summa_val
      f_summa_rub = m.summa_rub
      f_sum_pr_val = ALLTRIM(m.sum_pr_val)
      f_sum_pr_rub = ALLTRIM(m.sum_pr_rub)
      f_naznach = ALLTRIM(m.naznach)
      f_osnovanie = ALLTRIM(m.osnovanie)
      f_data_tran = m.data_tran
      f_ima_pk = ALLTRIM(m.ima_pk)
      f_num_oper = ALLTRIM(m.num_oper)

      time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      stroka_sql = "EXECUTE export_memordval"

      stroka_sql = stroka_sql + " ?f_record, ?f_kodval, ?f_kurs_zb, ?f_vid, ?f_kod, ?f_kod_oper, ?f_num_ord, ?f_kod_proekt, ?f_card_acct, ?f_fio_rus, ?f_data_doc,"
      stroka_sql = stroka_sql + " ?f_debit, ?f_name_dbt, ?f_credit, ?f_name_crt, ?f_summa_val, ?f_summa_rub, ?f_sum_pr_val, ?f_sum_pr_rub,"
      stroka_sql = stroka_sql + " ?f_naznach, ?f_osnovanie, ?f_data_tran, ?f_ima_pk, ?f_num_oper"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      export_data = SQLEXEC(sql_num, stroka_sql, '')

      DO WHILE SQLMORERESULTS(sql_num) < 2
      ENDDO

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + ' Дата документа - ' + DTOC(data_doc) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT memordval
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт валютных мемориальных ордеров в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('memordval.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


*********************************************************************** PROCEDURE export_prixodrub ******************************************************

PROCEDURE export_prixodrub
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('prixodrub.dbf'))

  IF NOT USED('prixodrub')

    SELECT 0
    USE (put_dbf) + ('prixodrub.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT prixodrub

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('prixodrub')
        SELECT prixodrub
        USE
        USE (put_dbf) + ('prixodrub.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('prixodrub.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(data_sost) + ALLTRIM(fio_rus) TAG date_fio

      SELECT prixodrub
      USE
      USE (put_dbf) + ('prixodrub.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(data_sost, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт рублевых приходных кассовых ордеров в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT prixodrub
* SET FILTER TO BETWEEN(data_sost, date_home, date_end)
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR  BETWEEN(data_sost, date_home, date_end)

      SCATTER MEMVAR MEMO

      f_record = ALLTRIM(m.record)
      f_kodval = m.kodval
      f_kod_oper = ALLTRIM(m.kod_oper)
      f_nn_plat = m.nn_plat
      f_data_sost = m.data_sost
      f_kod = ALLTRIM(m.kod)
      f_vid = ALLTRIM(m.vid)
      f_kod_proekt = ALLTRIM(m.kod_proekt)
      f_card_acct = ALLTRIM(m.card_acct)
      f_fio_rus = ALLTRIM(m.fio_rus)
      f_debit = ALLTRIM(m.debit)
      f_credit = ALLTRIM(m.credit)
      f_summa = m.summa
      f_summa_prop = ALLTRIM(m.summa_prop)
      f_prim = ALLTRIM(m.prim)
      f_name_doc = ALLTRIM(m.name_doc)
      f_pass_seria = ALLTRIM(m.pass_seria)
      f_pass_num = ALLTRIM(m.pass_num)
      f_pass_data = m.pass_data
      f_pass_mesto = ALLTRIM(m.pass_mesto)
      f_data_tran = m.data_tran
      f_ima_pk = ALLTRIM(m.ima_pk)
      f_num_oper = ALLTRIM(m.num_oper)

      time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      stroka_sql = "EXECUTE export_prixodrub"

      stroka_sql = stroka_sql + " ?f_record, ?f_kodval, ?f_kod_oper, ?f_nn_plat, ?f_data_sost, ?f_kod, ?f_vid, ?f_kod_proekt, ?f_card_acct, ?f_fio_rus, ?f_debit, ?f_credit,"
      stroka_sql = stroka_sql + " ?f_summa, ?f_summa_prop, ?f_prim, ?f_name_doc, ?f_pass_seria, ?f_pass_num, ?f_pass_data, ?f_pass_mesto, ?f_data_tran, ?f_ima_pk, ?f_num_oper"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      export_data = SQLEXEC(sql_num, stroka_sql, '')

      DO WHILE SQLMORERESULTS(sql_num) < 2
      ENDDO

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + ' Дата документа - ' + DTOC(data_sost) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT prixodrub
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт рублевых приходных кассовых ордеров в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('prixodrub.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


*********************************************************************** PROCEDURE export_rasxodrub *****************************************************

PROCEDURE export_rasxodrub
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('rasxodrub.dbf'))

  IF NOT USED('rasxodrub')

    SELECT 0
    USE (put_dbf) + ('rasxodrub.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT rasxodrub

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('rasxodrub')
        SELECT rasxodrub
        USE
        USE (put_dbf) + ('rasxodrub.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('rasxodrub.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(data_sost) + ALLTRIM(fio_rus) TAG date_fio

      SELECT rasxodrub
      USE
      USE (put_dbf) + ('rasxodrub.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(data_sost, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт рублевых расходных кассовых ордеров в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT rasxodrub
* SET FILTER TO BETWEEN(data_sost, date_home, date_end)
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR  BETWEEN(data_sost, date_home, date_end)

      SCATTER MEMVAR MEMO

      f_record = ALLTRIM(m.record)
      f_kodval = m.kodval
      f_kod_oper = ALLTRIM(m.kod_oper)
      f_nn_plat = m.nn_plat
      f_data_sost = m.data_sost
      f_kod = ALLTRIM(m.kod)
      f_vid = ALLTRIM(m.vid)
      f_kod_proekt = ALLTRIM(m.kod_proekt)
      f_card_acct = ALLTRIM(m.card_acct)
      f_fio_rus = ALLTRIM(m.fio_rus)
      f_debit = ALLTRIM(m.debit)
      f_credit = ALLTRIM(m.credit)
      f_summa = m.summa
      f_summa_prop = ALLTRIM(m.summa_prop)
      f_prim = ALLTRIM(m.prim)
      f_name_doc = ALLTRIM(m.name_doc)
      f_pass_seria = ALLTRIM(m.pass_seria)
      f_pass_num = ALLTRIM(m.pass_num)
      f_pass_data = m.pass_data
      f_pass_mesto = ALLTRIM(m.pass_mesto)
      f_data_tran = m.data_tran
      f_ima_pk = ALLTRIM(m.ima_pk)
      f_num_oper = ALLTRIM(m.num_oper)

      time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      stroka_sql = "EXECUTE export_rasxodrub"

      stroka_sql = stroka_sql + " ?f_record, ?f_kodval, ?f_kod_oper, ?f_nn_plat, ?f_data_sost, ?f_kod, ?f_vid, ?f_kod_proekt, ?f_card_acct, ?f_fio_rus, ?f_debit, ?f_credit,"
      stroka_sql = stroka_sql + " ?f_summa, ?f_summa_prop, ?f_prim, ?f_name_doc, ?f_pass_seria, ?f_pass_num, ?f_pass_data, ?f_pass_mesto, ?f_data_tran, ?f_ima_pk, ?f_num_oper"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      export_data = SQLEXEC(sql_num, stroka_sql, '')

      DO WHILE SQLMORERESULTS(sql_num) < 2
      ENDDO

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + ' Дата документа - ' + DTOC(data_sost) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT rasxodrub
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт рублевых расходных кассовых ордеров в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('rasxodrub.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


*********************************************************************** PROCEDURE export_prixodval ******************************************************

PROCEDURE export_prixodval
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('prixodval.dbf'))

  IF NOT USED('prixodval')

    SELECT 0
    USE (put_dbf) + ('prixodval.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT prixodval

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('prixodval')
        SELECT prixodval
        USE
        USE (put_dbf) + ('prixodval.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('prixodval.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(data_sost) + ALLTRIM(fio_rus) TAG date_fio

      SELECT prixodval
      USE
      USE (put_dbf) + ('prixodval.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(data_sost, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт валютных приходных кассовых ордеров в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT prixodval
* SET FILTER TO BETWEEN(data_sost, date_home, date_end)
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR  BETWEEN(data_sost, date_home, date_end)

      SCATTER MEMVAR MEMO

      f_record = ALLTRIM(m.record)
      f_kodval = m.kodval
      f_kurs_zb = m.kurs_zb
      f_kod_oper = ALLTRIM(m.kod_oper)
      f_nn_plat = m.nn_plat
      f_data_sost = m.data_sost
      f_kod = ALLTRIM(m.kod)
      f_vid = ALLTRIM(m.vid)
      f_kod_proekt = ALLTRIM(m.kod_proekt)
      f_card_acct = ALLTRIM(m.card_acct)
      f_fio_rus = ALLTRIM(m.fio_rus)
      f_debit = ALLTRIM(m.debit)
      f_credit = ALLTRIM(m.credit)
      f_summa_val = m.summa_val
      f_summa_rub = m.summa_rub
      f_sum_pr_val = ALLTRIM(m.sum_pr_val)
      f_sum_pr_rub = ALLTRIM(m.sum_pr_rub)
      f_prim = ALLTRIM(m.prim)
      f_name_doc = ALLTRIM(m.name_doc)
      f_pass_seria = ALLTRIM(m.pass_seria)
      f_pass_num = ALLTRIM(m.pass_num)
      f_pass_data = m.pass_data
      f_pass_mesto = ALLTRIM(m.pass_mesto)
      f_data_tran = m.data_tran
      f_ima_pk = ALLTRIM(m.ima_pk)
      f_num_oper = ALLTRIM(m.num_oper)

      time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      stroka_sql = "EXECUTE export_prixodval"

      stroka_sql = stroka_sql + " ?f_record, ?f_kodval, ?f_kurs_zb, ?f_kod_oper, ?f_nn_plat, ?f_data_sost, ?f_kod, ?f_vid, ?f_kod_proekt, ?f_card_acct, ?f_fio_rus, ?f_debit, ?f_credit,"
      stroka_sql = stroka_sql + " ?f_summa_val, ?f_summa_rub, ?f_sum_pr_val, ?f_sum_pr_rub, ?f_prim,"
      stroka_sql = stroka_sql + " ?f_name_doc, ?f_pass_seria, ?f_pass_num, ?f_pass_data, ?f_pass_mesto, ?f_data_tran, ?f_ima_pk, ?f_num_oper"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      export_data = SQLEXEC(sql_num, stroka_sql, '')

      DO WHILE SQLMORERESULTS(sql_num) < 2
      ENDDO

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + CHR(13) + ' Дата документа - ' + DTOC(data_sost) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT prixodval
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт валютных приходных кассовых ордеров в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('prixodval.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


*********************************************************************** PROCEDURE export_rasxodval *****************************************************

PROCEDURE export_rasxodval
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('rasxodval.dbf'))

  IF NOT USED('rasxodval')

    SELECT 0
    USE (put_dbf) + ('rasxodval.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT rasxodval

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('rasxodval')
        SELECT rasxodval
        USE
        USE (put_dbf) + ('rasxodval.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('rasxodval.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(data_sost) + ALLTRIM(fio_rus) TAG date_fio

      SELECT rasxodval
      USE
      USE (put_dbf) + ('rasxodval.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(data_sost, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт валютных расходных кассовых ордеров в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT rasxodval
* SET FILTER TO BETWEEN(data_sost, date_home, date_end)
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR  BETWEEN(data_sost, date_home, date_end)

      SCATTER MEMVAR MEMO

      f_record = ALLTRIM(m.record)
      f_kodval = m.kodval
      f_kurs_zb = m.kurs_zb
      f_kod_oper = ALLTRIM(m.kod_oper)
      f_nn_plat = m.nn_plat
      f_data_sost = m.data_sost
      f_kod = ALLTRIM(m.kod)
      f_vid = ALLTRIM(m.vid)
      f_kod_proekt = ALLTRIM(m.kod_proekt)
      f_card_acct = ALLTRIM(m.card_acct)
      f_fio_rus = ALLTRIM(m.fio_rus)
      f_debit = ALLTRIM(m.debit)
      f_credit = ALLTRIM(m.credit)
      f_summa_val = m.summa_val
      f_summa_rub = m.summa_rub
      f_sum_pr_val = ALLTRIM(m.sum_pr_val)
      f_sum_pr_rub = ALLTRIM(m.sum_pr_rub)
      f_prim = ALLTRIM(m.prim)
      f_name_doc = ALLTRIM(m.name_doc)
      f_pass_seria = ALLTRIM(m.pass_seria)
      f_pass_num = ALLTRIM(m.pass_num)
      f_pass_data = m.pass_data
      f_pass_mesto = ALLTRIM(m.pass_mesto)
      f_data_tran = m.data_tran
      f_ima_pk = ALLTRIM(m.ima_pk)
      f_num_oper = ALLTRIM(m.num_oper)

      time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      stroka_sql = "EXECUTE export_rasxodval"

      stroka_sql = stroka_sql + " ?f_record, ?f_kodval, ?f_kurs_zb, ?f_kod_oper, ?f_nn_plat, ?f_data_sost, ?f_kod, ?f_vid, ?f_kod_proekt, ?f_card_acct, ?f_fio_rus, ?f_debit, ?f_credit,"
      stroka_sql = stroka_sql + " ?f_summa_val, ?f_summa_rub, ?f_sum_pr_val, ?f_sum_pr_rub, ?f_prim,"
      stroka_sql = stroka_sql + " ?f_name_doc, ?f_pass_seria, ?f_pass_num, ?f_pass_data, ?f_pass_mesto, ?f_data_tran, ?f_ima_pk, ?f_num_oper"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      export_data = SQLEXEC(sql_num, stroka_sql, '')

      DO WHILE SQLMORERESULTS(sql_num) < 2
      ENDDO

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + ' Дата документа - ' + DTOC(data_sost) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT rasxodval
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт валютных расходных кассовых ордеров в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('rasxodval.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


*********************************************************************** PROCEDURE export_plateg *********************************************************

PROCEDURE export_plateg
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('plateg.dbf'))

  IF NOT USED('plateg')

    SELECT 0
    USE (put_dbf) + ('plateg.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT plateg

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('plateg')
        SELECT plateg
        USE
        USE (put_dbf) + ('plateg.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('plateg.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(data_s) + ALLTRIM(fio_rus) TAG date_fio

      SELECT plateg
      USE
      USE (put_dbf) + ('plateg.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(data_s, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт валютных расходных кассовых ордеров в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT plateg
* SET FILTER TO BETWEEN(data_s, date_home, date_end)
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR  BETWEEN(data_s, date_home, date_end)

      SCATTER MEMVAR MEMO

      f_record = ALLTRIM(m.record)
      f_data_s = m.data_s
      f_data_p = m.data_p
      f_kodval = m.kodval
      f_kod = ALLTRIM(m.kod)
      f_kod_oper = ALLTRIM(m.kod_oper)
      f_card_acct = ALLTRIM(m.card_acct)
      f_kod_proekt = ALLTRIM(m.kod_proekt)
      f_fio_rus = ALLTRIM(m.fio_rus)
      f_nameplat = ALLTRIM(m.nameplat)
      f_nn_plat = m.nn_plat
      f_tlg = ALLTRIM(m.tlg)
      f_summa = m.summa
      f_summa_prop = ALLTRIM(m.summa_prop)
      f_name_plat = ALLTRIM(m.name_plat)
      f_inn_plat = ALLTRIM(m.inn_plat)
      f_kpp_plat = ALLTRIM(m.kpp_plat)
      f_bank_plat = ALLTRIM(m.bank_plat)
      f_bik_plat = ALLTRIM(m.bik_plat)
      f_rkz_plat = ALLTRIM(m.rkz_plat)
      f_liz_plat = ALLTRIM(m.liz_plat)
      f_name_polu = ALLTRIM(m.name_polu)
      f_inn_polu = ALLTRIM(m.inn_polu)
      f_kpp_polu = ALLTRIM(m.kpp_polu)
      f_bank_polu = ALLTRIM(m.bank_polu)
      f_bik_polu = ALLTRIM(m.bik_polu)
      f_rkz_polu = ALLTRIM(m.rkz_polu)
      f_liz_polu = ALLTRIM(m.liz_polu)
      f_vidopl = ALLTRIM(m.vidopl)
      f_grup = m.grup
      f_prim = ALLTRIM(m.prim)
      f_vid = ALLTRIM(m.vid)
      f_data_tran = m.data_tran
      f_ima_pk = ALLTRIM(m.ima_pk)
      f_num_oper = ALLTRIM(m.num_oper)

      time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      stroka_sql = "EXECUTE export_plateg"

      stroka_sql = stroka_sql + " ?f_record, ?f_data_s, ?f_data_p, ?f_kodval, ?f_kod, ?f_kod_oper, ?f_card_acct, ?f_kod_proekt, ?f_fio_rus,"
      stroka_sql = stroka_sql + " ?f_nameplat, ?f_nn_plat, ?f_tlg, ?f_summa, ?f_summa_prop,"
      stroka_sql = stroka_sql + " ?f_name_plat, ?f_inn_plat, ?f_kpp_plat, ?f_bank_plat, ?f_bik_plat, ?f_rkz_plat, ?f_liz_plat,"
      stroka_sql = stroka_sql + " ?f_name_polu, ?f_inn_polu, ?f_kpp_polu, ?f_bank_polu, ?f_bik_polu, ?f_rkz_polu, ?f_liz_polu,"
      stroka_sql = stroka_sql + " ?f_vidopl, ?f_grup, ?f_prim, ?f_vid, ?f_data_tran, ?f_ima_pk, ?f_num_oper"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      export_data = SQLEXEC(sql_num, stroka_sql, '')

      DO WHILE SQLMORERESULTS(sql_num) < 2
      ENDDO

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + ' Дата документа - ' + DTOC(data_s) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT plateg
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт валютных расходных кассовых ордеров в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('plateg.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


********************************************************************** PROCEDURE export_platprn *********************************************************

PROCEDURE export_platprn
PARAMETERS date_home, date_end, name_sql_server

SET ESCAPE ON

* date_home = IIF(date_home >= CTOD('01.01.2006'), date_home, CTOD('01.01.2006'))

* WAIT WINDOW 'date_home = ' + DTOC(date_home) + '  date_end = ' + DTOC(date_end) TIMEOUT 3

IF FILE((put_dbf) + ('platprn.dbf'))

  IF NOT USED('platprn')

    SELECT 0
    USE (put_dbf) + ('platprn.dbf')

  ENDIF

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет открытие исходной таблицы с данными ... ',WCOLS())

  SELECT platprn

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  log_tag = .F.

  FOR I = 1 TO 10

    log_tag = LOWER(TAG(I)) == 'date_fio'

    IF log_tag = .T.
      EXIT
    ENDIF

  ENDFOR

  IF log_tag = .F.

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      WAIT WINDOW 'Внимание! Начато построения необходимого для работы индекса ... ' NOWAIT

      IF USED('platprn')
        SELECT platprn
        USE
        USE (put_dbf) + ('platprn.dbf') EXCLUSIVE
      ELSE
        SELECT 0
        USE (put_dbf) + ('platprn.dbf') EXCLUSIVE
      ENDIF

      INDEX ON DTOS(data_s) + ALLTRIM(fio_rus) TAG date_fio

      SELECT platprn
      USE
      USE (put_dbf) + ('platprn.dbf') SHARE

      WAIT CLEAR

    ENDIF
  ENDIF

  SET ORDER TO date_fio

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  COUNT TO colvo_zap FOR BETWEEN(data_s, date_home, date_end)

  WAIT WINDOW 'Количество записей, по которым будет произведен экспорт в СУБД SQL Server равно - '+ALLTRIM(STR(colvo_zap)) TIMEOUT 3

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('Открытие исходной таблицы с данными успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* BROWSE WINDOW brows

  IF colvo_zap <> 0

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('Внимание! Начат экспорт валютных расходных кассовых ордеров в СУБД SQL Server ... ' + SPACE(2) + ALLTRIM(name_sql_server),WCOLS())

    SELECT platprn
* SET FILTER TO BETWEEN(data_s, date_home, date_end)
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    SCAN FOR  BETWEEN(data_s, date_home, date_end)

      SCATTER MEMVAR MEMO

      f_record = ALLTRIM(m.record)
      f_data_s = m.data_s
      f_data_p = m.data_p
      f_kodval = m.kodval
      f_kod = ALLTRIM(m.kod)
      f_kod_oper = ALLTRIM(m.kod_oper)
      f_card_acct = ALLTRIM(m.card_acct)
      f_kod_proekt = ALLTRIM(m.kod_proekt)
      f_fio_rus = ALLTRIM(m.fio_rus)
      f_nn_plat = m.nn_plat
      f_tlg = ALLTRIM(m.tlg)
      f_summa = m.summa
      f_summa_prop = ALLTRIM(m.summa_prop)
      f_name_plat = ALLTRIM(m.name_plat)
      f_inn_plat = ALLTRIM(m.inn_plat)
      f_kpp_plat = ALLTRIM(m.kpp_plat)
      f_bank_plat = ALLTRIM(m.bank_plat)
      f_bik_plat = ALLTRIM(m.bik_plat)
      f_rkz_plat = ALLTRIM(m.rkz_plat)
      f_liz_plat = ALLTRIM(m.liz_plat)
      f_name_polu = ALLTRIM(m.name_polu)
      f_inn_polu = ALLTRIM(m.inn_polu)
      f_kpp_polu = ALLTRIM(m.kpp_polu)
      f_bank_polu = ALLTRIM(m.bank_polu)
      f_bik_polu = ALLTRIM(m.bik_polu)
      f_rkz_polu = ALLTRIM(m.rkz_polu)
      f_liz_polu = ALLTRIM(m.liz_polu)
      f_vidopl = ALLTRIM(m.vidopl)
      f_grup = m.grup
      f_prim = ALLTRIM(m.prim)
      f_vid = ALLTRIM(m.vid)
      f_data_tran = m.data_tran
      f_ima_pk = ALLTRIM(m.ima_pk)
      f_num_oper = ALLTRIM(m.num_oper)

      time_home = SECONDS()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      stroka_sql = "EXECUTE export_platprn"

      stroka_sql = stroka_sql + " ?f_record, ?f_data_s, ?f_data_p, ?f_kodval, ?f_kod, ?f_kod_oper, ?f_card_acct, ?f_kod_proekt, ?f_fio_rus,"
      stroka_sql = stroka_sql + " ?f_nn_plat, ?f_tlg, ?f_summa, ?f_summa_prop,"
      stroka_sql = stroka_sql + " ?f_name_plat, ?f_inn_plat, ?f_kpp_plat, ?f_bank_plat, ?f_bik_plat, ?f_rkz_plat, ?f_liz_plat,"
      stroka_sql = stroka_sql + " ?f_name_polu, ?f_inn_polu, ?f_kpp_polu, ?f_bank_polu, ?f_bik_polu, ?f_rkz_polu, ?f_liz_polu,"
      stroka_sql = stroka_sql + " ?f_vidopl, ?f_grup, ?f_prim, ?f_vid, ?f_data_tran, ?f_ima_pk, ?f_num_oper"

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      export_data = SQLEXEC(sql_num, stroka_sql, '')

      DO WHILE SQLMORERESULTS(sql_num) < 2
      ENDDO

      time_end = SECONDS()

      f_result_sql = ROUND(time_end - time_home, 3)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
      WAIT WINDOW 'Клиент - ' + ALLTRIM(m.fio_rus) + ' Номер счета - ' + ALLTRIM(m.card_acct) + ' Дата документа - ' + DTOC(data_s) + CHR(13) + ;
        'Результат выполнения операции export_data = ' + ALLTRIM(STR(export_data)) + CHR(13) + ;
        'Время выполнения операции = ' + ALLTRIM(STR(f_result_sql, 8, 3)) + SPACE(2) NOWAIT
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      SELECT platprn
    ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT CLEAR

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('Экспорт валютных расходных кассовых ордеров в СУБД SQL Server успешно завершен.' + ' (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' сек.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('Внимание! Запрашиваемых Вами данных не обнаружено ...')
  ENDIF

ELSE
  =err('Внимание! Файл - ' + (put_dbf) + ('platprn.dbf') + ' по данному пути не найден ... ')
ENDIF

SET ESCAPE OFF

RETURN


*********************************************************************************************************************************************************

































