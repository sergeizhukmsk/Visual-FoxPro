**********************************************
*** Карточный комплекс "CARD-VISA филиал ***
**********************************************

* ---------------------------------------------------------------------------------------------------------- Hачальные установки ----------------------------------------------------------------------------------------------------- *

CLEAR ALL  && Очистить экран

SET STATUS BAR ON && выключить статус строку
SET OPTIMIZE ON  && Включить режим оптимизации
SET CURSOR OFF && курсор включен или выключен
SET ESCAPE OFF && запрет на выход по esc
SET TALK OFF && запрет на эхо
SET SAFETY OFF && запрет на вывод сообщения о перезаписи файла
SET MESSAGE TO 24 CENTER && Задает сообщение, которое будет отображаться.
SET DELETED ON  && запрет на работу с удаленными записями
SET NEAR OFF && Устанавливает указатель записи на конец таблицы, если поиск записи с помощью команды FIND или SEEK завершился неудачным.
SET EXACT ON && Указывает, что выражения будут эквивалентны, если они совпадают посимвольно вплоть до конца выражения, расположенного справа.
SET ANSI ON  && Указывает на совместимость с кодирвкой ANSI
SET NOTIFY OFF && Разрешает или отменяет отображение некоторых системных сообщений.
SET STRICTDATE TO 0  && Совместимость с 2000 годом
SET MEMOWIDTH TO 30 && ширина вывода мемо файла (колонки)
SET WINDOW OF MEMO TO nazvan && окно для вывода мемо файла во время просмотра
SET BELL ON && Включает или выключает звуковой сигнал компьютера, а также устанавливает атрибуты сигнала.
SET HELP OFF  && Отключить системную помощь
SET CLOCK OFF  && Отключить системные часы
SET SYSMENU OFF  && Отключить системное меню
SET CONFIRM ON && Указывает, что вы не можете выходить из текстового поля, вводя данные правее его последнего символа.
SET COLOR TO && Задает цвета в цветовой схеме (цвета по умолчанию)
SET INTENSITY OFF && Задает атрибут обычного экрана для отображения полей редактирования.
SET COLLATE TO 'RUSSIAN' && порядок сортировки
SET COMPATIBLE OFF && разрешает программам, созданным в FoxBASE+, работать в Visual FoxPro без изменений.
SET UDFPARMS TO VALUE && Передача параметров по значению.
SET DEBUG OFF && Делает окна отладки и трассировки доступными или недоступными
SET STEP OFF && Закрытие окна трассировки
SET ECHO OFF && Запрет трассировки
SET CURRENCY RIGHT && Выравнивание по правому краю
SET UNIQUE OFF  && Признак построения уникальных индексов
SET TABLEVALIDATE TO 2  && Уровень проверки таблицы
SET REPORTBEHAVIOR 80  && Возможные значения 80/90. Изменение окна предварительного просмотра печатного отчета, используя возможности Visual FoxPro 9.0

SET SYSFORMATS OFF && настройки системы выключены

SET HOURS TO 24  && Установка часов в сутках
SET DATE TO GERMAN  && Формат даты
SET FDOW TO 2 && первый день недели понедельник
SET FWEEK TO 2 && первая неделя года
SET CENTURY ON && полный формат года
SET CURRENCY TO ' руб' && тип валюты по умолчанию
SET POINT TO '.' && тип разделителя целой и дробной части.
SET MARK TO '.' && тип разделителя в формате даты
SET DECIMALS TO 2 && количество знаков дробной части
SET FIXED OFF && разрешен вывод дробной части (копейки)

PUSH KEY CLEAR

* ----------------------------------------------------------------------------------- Установки для многопользовательской работы в сети ------------------------------------------------------------------------------ *

SET EXCLUSIVE OFF && Запрет монопольного захвата записи
SET LOCK OFF  && Запрет блокировки записи
SET MULTILOCKS ON && Буферизаци таблиц
SET REPROCESS TO AUTOMATIC  && Автоматическое обновление
SET REFRESH TO 5 && Обновление Чтение/Запись (сек)

* ------------------------------------------------- Установка дополнительных SET параметров для Visual FoxPro 9.0 для совместимости с версией 7.0 ---------------------------------------------- *

IF LEFT(ALLTRIM(VERSION()), 16) == 'Visual FoxPro 09'
  SET ENGINEBEHAVIOR 70
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC direktoria, I, A

sleg = 1
I = 3

DO WHILE sleg <> 0
  sleg = AT('\', SYS(16), I)
  IF sleg <> 0
    I = I + 1
  ELSE
    I = I - 2
  ENDIF
ENDDO

A = AT('\', SYS(16), I)

direktoria = LEFT(SYS(16), A)

* WAIT WINDOW 'direktoria - ' + ALLTRIM(direktoria)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
SET PATH TO (direktoria) + ('EXE\') + ';' + (direktoria) + ('DBF\') + ';' + (direktoria) + ('IZD\') + ';' + (direktoria) + ('PRG\') + ';' + (direktoria) + ('FRX\') + ';' + (direktoria) + ('BMP\') + ';' + (direktoria) + ('SCR\')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC popup_ima, prompt_ima, bar_num

popup_ima = 'Головной модуль'
prompt_ima = 'Головной модуль'
bar_num = 0

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
ON ERROR DO ErrTrap WITH PROGRAM(), LINENO(), ERROR(), MESSAGE(), MESSAGE(1), LOWER(ALLTRIM(popup_ima)) + ' - ' + ALLTRIM(STR(bar_num))
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC data, col, row, text_mess, per, tiplog, log, log_quit, vipdata, punkt, dbfname, cdxname, fail, faildbf, failcdx

PUBLIC tex1, tex2, tex3, tex4, tex5, tex6, tex7, tex8, tex9, tex10, tex11, tex12, text1, text2, text3, text4, text5, text6, text7, text8, text9, text10, text11, text12, L_bar

PUBLIC par_fio, par_ima, par_kod, par_par, par_prim, par_prompt, par_grup, fio_user, adm_kod, poisk_card_num, colvo_den_god, num_den_god, pr_time_rachet_proz, date_filtr

PUBLIC bar, bar_glmenu, analbar, putnew, put_bar, put_prn, put_vivod, put_select, put_select_data, put_type_select, err_text, poisk, colvo, colvo_zap, personfio, personima, personparol, text_titul, predlog

PUBLIC sel_client, sel_fio, sel_fio_rus, sel_ref_client, sel_inn_client, sel_punkt, sel_card_num, sel_record, sel_summa, vid_frx, vid_ostat, export_data, export_data_1, export_data_2, time_home, time_end

PUBLIC sel_date_vedom, sel_date_ostat, sel_date_ost_p, sel_date_ost_m, ddmm, ima_sql_server, put_spr, pr_rezident, poisk_val_card, poisk_date_card, sel_client_b, sel_birth_date

PUBLIC sel_pass_seria, sel_pass_num, pass_seria_str, pass_num_str, pass_data_str, pass_mesto_str, pass_seria_new, pass_num_new, pass_data_new, pass_mesto_new

PUBLIC colzap, rown, title_text, sprbaze, baze, bazenum, vidscr, vidfrx, vidord, per, logvid, numer, dt, summer, person, pr_export_golova, title_report, proz_cred, pr_zarpl

PUBLIC dbfname, cdxname, colcopy, numord, barord, barbaze, client, clientfam, clientord, clientrec, mbal, mval, mkl, mliz, bazenumliz, sel_pr_dover, num_recno, sergeizhuk_sql_flag, lh3001_sql_flag

PUBLIC put, barput, mes, imames, Isort, xListSort, xFuncSort, ns_plat, ns_polu, bar_plateg, err_num, kodtext, exp_base, exp_file, exp_file_txt, flag_add_table, sergeizhuk_flag, sergeizhuk_name_pk

PUBLIC S_branch, S_fio_rus, S_card_nls, S_dokum, S_pass_seria, S_pass_num, S_pass_mesto, S_pass_data, propis_summa, read_ost_time, new_read_ostat

PUBLIC sum_tarif_r, sum_doxod_r, sum_nds_r, ima_file, num_client, poisk_numliz_term, sel_begin_bal, sel_debit, sel_credit, sel_end_bal, max_date_import

PUBLIC S_liz_plat, S_inn_plat, S_kpp_plat, S_bik_plat, S_rkz_plat, S_name_plat, S_bank_plat, S_liz_polu, S_inn_polu, S_kpp_polu, S_bik_polu, S_rkz_polu, S_name_polu, S_bank_polu, S_kodval

PUBLIC NumKr_pk1, sel_kod_oper, rec, num_zap, text_end, summa_val, summa_rub, sel_branch, poisk_branch, kor_bank_810, kor_bank_840, kor_bank_978, numvid, num_balan

PUBLIC sel_type_card, sel_num_card, sel_val_card, sel_name_card, import_flag, export_flag, erase_flag, read_flag, sel_kod_proekt, pr_kod_proekt, connect_sql

PUBLIC sel_tran_42301, sel_tran_42308, sel_card_acct, sel_card_nls, sel_card_tran, str_card_acct, str_card_nls, new_card_acct, new_card_nls, scan_data

PUBLIC zUCH, kSHET, pSHET, zSHET, vvod_stavka, vvod_sum_strah, vvod_sum_ostat, god_tek, god_arx, ima_table, num_id, ima_arh, poisk_kod_nds, put_kod, save_put, sel_deal_desc

PUBLIC Poisk_1, mBal_1, mVal_1, mKl_1, mLiz_1, Poisk_2, mBal_2, mVal_2, mKl_2, mLiz_2, Poisk_3, mBal_3, mVal_3, mKl_3, mLiz_3, Poisk_4, mBal_4, mVal_4, mKl_4, mLiz_4

PUBLIC poisk_loc_term, poisk_exp_term, poisk_city, fam_client, len_client, summa_str, summa_new, kod_val_doc, title_zarpl, num_mes_zarpl, new_data_odb, arx_data_odb, name_bank

PUBLIC pop_ord, bar_ord_gl, pop_plateg, bar_plateg, error_1705, error_1683, error_1108, error_monopol, god1, poisk_sum_begin_bal, begin_bal_cln, end_bal_cln, poisk_sum_end_bal, new_card_liz, ima_glav_file

PUBLIC sel_doc_record, poisk_kod_oper, sel_data_doc, sel_data_sost, poisk_debit, poisk_credit, sel_doc_summa, sel_tip_oper, sel_card_strah, ima_slovar_baza, name_bmp

PUBLIC pr_rabota_sql, text_vid_doc, scan_date_home, scan_date_end, pr_rabota_sql_str, pr_rabota_sql_new, sum_monitoring, pr_monitoring, type_monitoring

PUBLIC mod_city_dom, mod_city_reg, mod_region_dom, mod_region_reg, mod_rajon_dom, mod_rajon_reg

PUBLIC vvod_kurs_usd, vvod_kurs_eur, poisk_kurs_usd, poisk_kurs_eur, sel_kurs_val, sel_god, date_new_tarif, numbal, sel_numbal, select_osnov, num_bal_doxod, num_bal_rasxod

PUBLIC put_excel, put_unzip, put_rar, put_winrar, put_spr, colvo_select_data, status_printer, colvo_zap_dom, colvo_zap_reg, pr_start_reestr, pr_zakritie_odb

PUBLIC new_ind_dom, new_region_dom, new_rajon_dom, new_city_dom, new_punkt_dom, new_street_dom, new_dom_dom, new_korpus_dom, new_stroen_dom, new_kvart_dom, new_adres_dom, new_kladr_dom

PUBLIC new_ind_reg, new_region_reg, new_rajon_reg, new_city_reg, new_punkt_reg, new_street_reg, new_dom_reg, new_korpus_reg, new_stroen_reg, new_kvart_reg, new_adres_reg, new_kladr_reg

PUBLIC str_ind_dom, str_domion_dom, str_rajon_dom, str_city_dom, str_punkt_dom, str_street_dom, str_dom_dom, str_korpus_dom, str_stroen_dom, str_kvart_dom, str_adres_dom, str_kladr_dom

PUBLIC str_ind_reg, str_region_reg, str_rajon_reg, str_city_reg, str_punkt_reg, str_street_reg, str_dom_reg, str_korpus_reg, str_stroen_reg, str_kvart_reg, str_adres_reg, str_kladr_reg

PUBLIC new_pass_type, new_vid_doc, new_pass_seria, new_pass_num, new_pass_data, new_pass_mesto, seek_kurs_den, f_result_sql, date_select_rem_view_home, date_select_rem_view

PUBLIC str_pass_type, str_vid_doc, str_pass_seria, str_pass_num, str_pass_data, str_pass_mesto, sel_doveren, poisk_vedom_ost, sql_num, pervasive_num

PUBLIC istor_ost_date_ostat, istor_ost_begin_bal, istor_ost_debit, istor_ost_credit, istor_ost_end_bal, summa_home, summa_end

PUBLIC istor_mak_date_ostat, istor_mak_begin_bal, istor_mak_debit, istor_mak_credit, istor_mak_end_bal

PUBLIC str_account_begin_bal, str_account_debit, str_account_credit, str_account_end_bal, str_account_vxd_ost_m, str_account_debit_m, str_account_credit_m, str_account_isx_ost_m

PUBLIC tim1, tim2, tim3, tim4, tim5, tim6, tim7, tim8, tim9, tim10, tim11, tim12, tim13, tim14, tim15, tim16, tim17, tim18, tim19, tim20, tim21, tim22, tim23, tim24, tim25

PUBLIC time1, time2, time3, time4, time5, time6, time7, time8, time9, time10, time11, time12, time13, time14, time15, time16, time17, time18, time19, time20, time21, time22, time23, time24, time25

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC put_bmp, put_arx, put_dbf, put_frx, put_scr, put_prg, put_exe, put_exp,;
  put_out_dbf, put_out_upk, put_out_exp, put_out_imp, put_out_pb, put_in, put_out_excelens,;
  put_tmp, put_doc, put_zrp, put_god, put_bmp_net

put_bmp = (direktoria) + ('BMP\')  && Директория для графических файлов
put_arx = (direktoria) + ('ARX\')  && Директория для архивных копий таблиц, архивных копий индексов
put_dbf = (direktoria) + ('DBF\')  && Директория для таблиц, индексов
put_frx = (direktoria) + ('FRX\')  && Директория для печатных отчетов
put_scr = (direktoria) + ('SCR\')  && Директория для форм
put_prg = (direktoria) + ('PRG\')  && Директория для программных файлов
put_exe = (direktoria) + ('EXE\')  && Директория для исполняемых файлов
put_exp = (direktoria) + ('EXP\')  && Директория для файлов экспорта для ОДБ
put_out_dbf = (direktoria) + ('OUT\DBF\')  && Директория для входящих таблиц, приходящих из филиалов по межфилиальной почте
put_out_upk = (direktoria) + ('OUT\UPK\')  &&  Директория для исходящей информации о заявлениях на сдачу карты, об отказе и утере карты для головного банка
put_out_exp = (direktoria) + ('OUT\EXP\')  && Директория для исходящей информации для головного банка, таблицы экспорта
put_out_imp = (direktoria) + ('OUT\IMP\')  && Директория для исходящей информации в процессинг, данные для открытия карт
put_out_pb = (direktoria) + ('OUT\PB\')  && Директория для исходящей информации в процессинг, данные для файлов пополнения, списания средств
put_out_excelens = (direktoria) + ('OUT\EXCELENS\')  && Директория для обработки файлов с электронной подписью
put_in = (direktoria) + ('IN\')  && Директория для входящей информации из головного банка
put_vipis = (direktoria) + ('IN\VIPIS\')  && Директория для выписки из головного банка в текстовом формате
put_tmp = (direktoria) + ('TMP\')  && Директория для временных файлов
put_doc = (direktoria) + ('DOC\')  && Директория для отчетов конвертируемых в Word
put_zrp = (direktoria) + ('ZRP\Kod_')  && Директория для файлов начисления по зарплатному проекту
put_god = (direktoria) + ('GOD\Kod_')  && Директория для файлов списания за годовое обслуживание

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

B = AT('\', SYS(16), I - 1)
put_bmp_net = LEFT(SYS(16), B) + ('\Bmp_Files\')  && Директория для графических файлов в сетевом варинте для филиалов № 1, 3

*    WAIT WINDOW 'SYS_16 - ' + SYS(16)
*    WAIT WINDOW 'put_bmp_net - ' + ALLTRIM(put_bmp_net)

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

erase_flag = .T.
import_flag = .T.
export_flag = .F.
aktual = .F.

* flag_add_table = .T.  && Признак лечения баз данных, сначала таблица исключается из базы, затем добавляется.

proz_cred = 40.00  && Процентная ставка по овердрафтным счетам
wes = 5
newbar = .F.
vid_memo = .F.
imp_sprav_dop = .T.
impdoctest = .F.  && Если impdoctest = .T., то на экран выводится временная тестовая информация скорости выборки данных
impsprdop = .F.
impsprbik = .F.

predlog = 'за '

date_select_rem_view_home = '2011-01-01'

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
pr_time_rachet_proz = .T.  && Если pr_time_rachet_proz = .T., то в процентные таблицы пишется время расчета каждой записи
* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
path_str = SYS(5) + SYS(2003)
put_tmp_dbf = SYS(2023) + '\'
* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

STORE SPACE(0) TO sprbaze, baze, bazenum, vidscr, vidfrx, vidord, numer, person, dbfname, cdxname, fail, faildbf, failcdx
STORE SPACE(0) TO numord, barord, client, clientfam, clientord, clientrec, bazekod, bazetip, bazeprim, bazenaznach, numrec, sel_god, sel_kod_proekt
STORE SPACE(0) TO balns, recprim, recnaznach, bazenumliz, bazenumlizDB, bazenumlizKR, summa_prop_text, sel_ref_client, sel_inn_client, sel_client, sel_card_num, sel_record, sel_card_pr, sel_card_tip
STORE SPACE(0) TO tex1, tex2, tex3, tex4, tex5, tex6, tex7, tex8, numbal, sel_numbal, numbal1, numbal2, numacc, numacc1, numacc2, kodtext, kodvid, title_text, sel_fio, sel_fio_rus
STORE SPACE(0) TO P_debit, P_credit, P_naznach, P_osnovanie, P_prilozh, text_memo, prn_memo_1, prn_memo_2, pop_ord, pop_plateg, ddmm
STORE SPACE(0) TO den, mes, god, ima_frx, ima_frx_det, ima_frx_sum, ima_table, sel_kod_oper, num_balan, new_card_liz
STORE SPACE(0) TO sel_tran_42301, sel_tran_42308, sel_card_acct, sel_card_nls, sel_card_tran, str_card_acct, str_card_nls, new_card_acct, new_card_nls
STORE SPACE(0) TO sel_type_card, sel_num_card, sel_val_card, sel_name_card, propis_summa, fam_client, kod_val_doc
STORE SPACE(0) TO S_liz_plat, S_inn_plat, S_kpp_plat, S_bik_plat, S_rkz_plat, S_name_plat, S_bank_plat, S_liz_polu, S_inn_polu, S_kpp_polu, S_bik_polu, S_rkz_polu, S_name_polu, S_bank_polu, S_kodval
STORE SPACE(0) TO Poisk_1, mBal_1, mVal_1, mKl_1, mLiz_1, Poisk_2, mBal_2, mVal_2, mKl_2, mLiz_2, Poisk_3, mBal_3, mVal_3, mKl_3, mLiz_3, Poisk_4, mBal_4, mVal_4, mKl_4, mLiz_4
STORE SPACE(0) TO title_zarpl, name_bank, ima_file, num_client, poisk_numliz_term, text_vid_doc, text_titul, ima_sql_server, pr_rezident
STORE SPACE(0) TO text1, text2, text3, text4, text5, text6, text7, text8, title_visa, exp_file_txt, poisk_val_card, sel_client_b, num_bal_doxod, num_bal_rasxod
STORE SPACE(0) TO personfio, personima, personparol, par_kod, par_fio, par_ima, par_grup, par_prim, numer, poisk, numrec, rec_plat, rec_polu, ns_plat, ns_polu
STORE SPACE(0) TO client, clientfam, clientord, clientrec, Poisk_Plat, Poisk_Polu, sel_deal_desc, select_osnov, sel_doveren, read_ost_time
STORE SPACE(0) TO mod_city_dom, mod_city_reg, mod_region_dom, mod_region_reg, mod_rajon_dom, mod_rajon_reg
STORE SPACE(0) TO sel_pass_seria, sel_pass_num, pass_seria_str, pass_num_str, pass_mesto_str, pass_seria_new, pass_num_new, pass_mesto_new, name_bmp

STORE SPACE(0) TO new_ind_dom, new_region_dom, new_rajon_dom, new_city_dom, new_punkt_dom, new_street_dom, new_dom_dom, new_korpus_dom, new_stroen_dom, new_kvart_dom, new_adres_dom, new_kladr_dom
STORE SPACE(0) TO new_ind_reg, new_region_reg, new_rajon_reg, new_city_reg, new_punkt_reg, new_street_reg, new_dom_reg, new_korpus_reg, new_stroen_reg, new_kvart_reg, new_adres_reg, new_kladr_reg
STORE SPACE(0) TO str_ind_dom, str_domion_dom, str_rajon_dom, str_city_dom, str_punkt_dom, str_street_dom, str_dom_dom, str_korpus_dom, str_stroen_dom, str_kvart_dom, str_adres_dom, str_kladr_dom
STORE SPACE(0) TO str_ind_reg, str_region_reg, str_rajon_reg, str_city_reg, str_punkt_reg, str_street_reg, str_dom_reg, str_korpus_reg, str_stroen_reg, str_kvart_reg, str_adres_reg, str_kladr_reg

STORE SPACE(0) TO new_pass_type, new_vid_doc, new_pass_seria, new_pass_num, new_pass_mesto
STORE SPACE(0) TO str_pass_type, str_vid_doc, str_pass_seria, str_pass_num, str_pass_mesto

STORE 1 TO barord, barbaze, recprn, numrec, rownum, colnum, punkt, bar, colvo, barput, analbar, Isort, xListSort, xFuncSort
STORE 1 TO put_vivod, put_vibor, put_prov_09, colvo_ckc, colvo, num_id, put_kod, numvid, put_select, put_type_select, vid_frx, vid_ostat, type_monitoring, put_select_data

STORE 0 TO export_data, export_data_1, export_data_2
STORE 0 TO bar_glmenu, bar_num, bar_ord_gl, bar_plateg, poisk_sum_begin_bal, begin_bal_cln, end_bal_cln, poisk_sum_end_bal, num_recno, time_home, time_end
STORE 0 TO colzap, rown, l_bar,  per, summer, colvo_zap, put_bar, num_kvartal, num_mesaz_kvartal, num_mesaz, sum_tarif_r, sum_doxod_r, sum_nds_r
STORE 0 TO vvod_kurs_usd, vvod_kurs_eur, poisk_kurs_usd, poisk_kurs_eur, sel_kurs_val, raschet_kurs, num_reis, ima_arh, len_client
STORE 0 TO data_kurs_usd, data_kurs_eur, summa_val, summa_rub, vvod_numord, konrlsum, newnumer, poisk_kod_nds, summa_str, summa_new
STORE 0 TO vvod_stavka, vvod_sum_strah, vvod_sum_ostat, itog_summa, num_oper, sum_atm, sum_atm_1, sum_atm_2, sum_atm_3, sum_atm_4, kod_branch, poisk_branch, kod_val
STORE 0 TO sel_begin_bal, sel_debit, sel_credit, sel_end_bal, sql_num, colvo_select_data, colvo_zap_dom, colvo_zap_reg, summa_home, summa_end

STORE SECONDS() TO tim1, tim2, tim3, tim4, tim5, tim6, tim7, tim8, tim9, tim10, tim11, tim12, tim13, tim14, tim15, tim16, tim17, tim18, tim19, tim20, tim21, tim22, tim23, tim24, tim25

STORE DATETIME() TO time1, time2, time3, time4, time5, time6, time7, time8, time9, time10, time11, time12, time13, time14, time15, time16, time17, time18, time19, time20, time21, time22, time23, time24, time25
STORE DATETIME() TO new_read_ostat

STORE DATE() TO dt, dt1, dt2, dt3, vvod_data_home, vvod_data_end, vvod_data_pak, data_home_end, data_frx, vvod_data, data_ckc, new_data_odb, arx_data_odb
STORE DATE() TO sel_date_vedom, sel_date_ostat, sel_date_ost_p, sel_date_ost_m, poisk_date_card, date_new_tarif, pass_data_str, pass_data_new, sel_birth_date
STORE DATE() TO new_pass_data, str_pass_data, date_filtr

STORE .F. TO read_flag, poisk_loc_term, poisk_exp_term, sel_pr_dover, poisk_card_num, pr_export_golova, pr_rabota_sql, pr_rabota_sql_str, pr_rabota_sql_new, log, log_quit, tiplog
STORE .F. TO poisk_city, status_printer, pr_start_reestr, pr_zakritie_odb, pr_kod_proekt, pr_zarpl, connect_sql, seek_kurs_den, poisk_vedom_ost
STORE .F. TO error_1683, error_1705, error_1108, error_monopol

STORE YEAR(DATE()) TO god_tek, god1
STORE YEAR(DATE()) - 1 TO god_arx

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
* Dbt_data = CREATEOBJECT('RS.LIB')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

num_bal_doxod = '70601'
num_bal_rasxod = '70606'

xdate = DATE()

date_new_tarif = CTOD('01.07.2008')

kod_val_rur = '810'
kod_val_usd = '840'
kod_val_eur = '978'

zUCH = SPACE(3)
STORE SPACE(20) TO kSHET, pSHET, zSHET

imaplat = SPACE(30)
TextDate = SPACE(5)

pr_arxiv = .F.

_THROTTLE = 1
STORE 3  TO L_bar

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
ima_glav_file = ALLTRIM(LOWER(RIGHT(ALLTRIM(SYS(16)), 8)))
* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

text_mess = 'Движение по меню - клавишами управления курсором ;  Выбоp пункта - ENTER ;  Выход - ESC'

title_visa = ' Программный комплекс "CARD-VISA филиал" - формирование документов по карточному обслуживанию'

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

erase_flag = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'PKSBZHUK', 'NBSBZHUK'), .F., .T.)

sergeizhuk_flag = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'PKSBZHUK', 'NBSBZHUK'), .T., .F.)

connect_sql = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'PKSBZHUK', 'NBSBZHUK'), .T., .F.)

sergeizhuk_sql_flag = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'PKSBZHUK', 'NBSBZHUK'), .T., .F.)

lh3001_sql_flag = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'PKSBZHUK', 'NBSBZHUK'), .T., .F.)

sergeizhuk_name_pk = ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1))   && 'PKSBZHUK', 'NBSBZHUK'

* sergeizhuk_flag = .F.

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO start IN Scan_rekvizit
* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF tiplog = .T.

  DO rasvert IN Visa

  DO WHILE .T.

    DO CASE
      CASE pr_dostup = .T.  && Признак разрешения на вход в ПО, если pr_dostup = .T., то вход разрешен, если pr_dostup = .F., то вход запрещен

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
        DO usebase IN Visa
* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        IF tiplog = .T.

          DO proverka IN Scan_rekvizit

          DO start IN Menu_visa  &&   && Запуск процедуры формирования главного меню программы VisaCard
          DO start IN Menu_zarpl  && Запуск процедуры формирования разных меню для зарплатных модулей программы

          SELECT spravliz

          SCAN
            PUBLIC (ALLTRIM(spravliz.ima))
            STORE (ALLTRIM(spravliz.numliz)) TO (ALLTRIM(spravliz.ima))
          ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          DO select_screen_oboi IN Visa

          IF error_1108 = .T.
            name_bmp = 'Vodopad.bmp'
            @ 0.000,0.000 SAY LOCFILE((put_bmp) + (name_bmp)) BITMAP SIZE colvo_rows + 1.5, colvo_cols + 1 STRETCH STYLE 'T'
          ENDIF

          DO select_data_odb IN Visa

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* sergeizhuk_flag = .F.

          DO CASE
            CASE LOWER(RIGHT(ALLTRIM(SYS(16)), 8)) == 'visa.exe' OR LOWER(RIGHT(ALLTRIM(SYS(16)), 8)) == 'visa.fxp'

              DO start_rabota_visa IN Visa

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            CASE LOWER(RIGHT(ALLTRIM(SYS(16)), 12)) == 'visa_new.exe' OR LOWER(RIGHT(ALLTRIM(SYS(16)), 12)) == 'visa_new.fxp'

              DO start_rabota_visa IN Visa

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            CASE LOWER(RIGHT(ALLTRIM(SYS(16)), 12)) == 'visa_sql.exe' OR LOWER(RIGHT(ALLTRIM(SYS(16)), 12)) == 'visa_sql.fxp'

              DO vxod_servis IN Visa

              DO CASE
                CASE ALLTRIM(sergeizhuk_name_pk) == 'PKSBZHUK' OR ALLTRIM(sergeizhuk_name_pk) == 'NBSBZHUK'

                  DO export_data_sql_avt_loc_2005 IN Vxod_sql  && Загрузка данных на локальный SQL Server 2005 по адресу 192.168.2.7
                  DO export_data_sql_avt_loc_2008 IN Vxod_sql  && Загрузка данных на локальный SQL Server 2008 по адресу 192.168.2.7


                CASE ALLTRIM(sergeizhuk_name_pk) == 'PKSERV2003'

                  DO export_data_sql_avt_loc_2005 IN Vxod_sql  && Загрузка данных на локальный SQL Server 2005 по адресу 192.168.2.7


                CASE ALLTRIM(sergeizhuk_name_pk) == 'PKWIN7'

* DO export_data_sql_avt_loc_2008 IN Vxod_sql  && Загрузка данных на локальный SQL Server 2005 по адресу 192.168.2.7

              ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            CASE LOWER(RIGHT(ALLTRIM(SYS(16)), 12)) == 'visa_imp.exe' OR LOWER(RIGHT(ALLTRIM(SYS(16)), 12)) == 'visa_imp.fxp'

              DO vxod_servis IN Visa
              DO start_avt IN Import_sprav

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            CASE LOWER(RIGHT(ALLTRIM(SYS(16)), 12)) == 'visa_ind.exe' OR LOWER(RIGHT(ALLTRIM(SYS(16)), 12)) == 'visa_ind.fxp'

              DO vxod_servis IN Visa
              DO start_avt IN Pack_table

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            CASE LOWER(RIGHT(ALLTRIM(SYS(16)), 12)) == 'visa_avt.exe' OR LOWER(RIGHT(ALLTRIM(SYS(16)), 12)) == 'visa_avt.fxp'

              DO vxod_servis IN Visa
              DO rabota_avt IN Visa

          ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          DO exit IN Visa
          EXIT

        ELSE
          WAIT WINDOW TIMEOUT 3 ('Внимание! Нужный для работы файл - ' + ALLTRIM(fail) + ' не найден.')
        ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE pr_dostup = .F.  && Признак разрешения на вход в ПО, если pr_dostup = .T., то вход разрешен, если pr_dostup = .F., то вход запрещен

        IF NOT USED('parol')

          SELECT 0
          USE (put_dbf) + ('parol.dbf') ALIAS parol SHARE
          SELECT parol
          GO TOP

        ELSE

          SELECT parol
          GO TOP

        ENDIF

        SELECT IIF(use_soft = .T., 'Активен', 'Не активен') AS use_soft, colvo_soft, kod, fio, grup ;
          FROM parol ;
          WHERE use_soft = .T. AND colvo_soft >= 1  ;
          INTO CURSOR sel_monopol

        IF _TALLY = 0

          SELECT rekvizit
          REPLACE dostup WITH .T.
          pr_dostup = rekvizit.dostup  && Признак разрешения на вход в ПО

          SELECT parol
          USE
          LOOP

        ELSE

          =err('Программный комплекс используется в монопольном режиме, ВХОД ЗАПРЕЩЕН ... ')

          DEFINE WINDOW vidparol AT 0, 0 SIZE colvo_rows - 25, colvo_cols - 60 FONT 'Times New Roman', num_schrift_v + 1  STYLE 'B' ;
            NOFLOAT NOCLOSE NOMINIMIZE DOUBLE TITLE ' Окно  редактирования  паролей ' COLOR RGB(0,0,0,192,192,192)
          MOVE WINDOW vidparol CENTER

          SELECT sel_monopol
          GO TOP

          BROWSE FIELDS ;
            use_soft:H = 'Активность',;
            colvo_soft:H = 'Кол-во копий',;
            kod:H = 'Код пользователя',;
            fio:H = 'Пользователь системы',;
            grup:H = 'Доступ к счетам' ;
            WINDOW vidparol ;
            TITLE '  Список пользователей работающих с программным обеспечением'

          RELEASE WINDOW vidparol

          text1 = 'Выполнить процедуру снятия запрета на вход в программу'
          text2 = 'Отказаться от выполнения процедуру снятия запрета'
          l_bar = 3
          =popup_9(text1, text2, text3, text4, l_bar)
          ACTIVATE POPUP vibor
          RELEASE POPUPS vibor

          IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

            DO CASE
              CASE BAR() = 1  && Выполнить процедуру снятия запрета на вход в программу

                IF USED('rekvizit')

                  SELECT rekvizit
                  GO TOP
                  REPLACE dostup WITH .T.
                  pr_dostup = rekvizit.dostup

                ELSE

                  SELECT 0
                  USE (put_dbf) + ('rekvizit.dbf') SHARE
                  GO TOP
                  REPLACE dostup WITH .T.
                  pr_dostup = rekvizit.dostup
                  USE

                ENDIF

                LOOP

              CASE BAR() = 3  && Отказаться от выполнения процедуру снятия запрета

                DO exit IN Visa
                EXIT

            ENDCASE

          ENDIF

        ENDIF

    ENDCASE

  ENDDO

ELSE
  WAIT WINDOW TIMEOUT 3 ('Внимание! Нужный для работы файл - ' + ALLTRIM(fail) + ' не найден.')
ENDIF


* ====================================================================================================================================== *


************************************************************ PROCEDURE start_rabota_visa ****************************************************************

PROCEDURE start_rabota_visa

IF sergeizhuk_flag = .T.

  DO CASE
    CASE pr_rabota_sql = .F.  && Выключен признак экспорта данных в таблицы SQL Server

      DO CASE
        CASE pr_import_loc = .T. AND pr_import_sql = .F. AND pr_import_sum = .F.  && Загрузка данных только в локальные таблицы

* WAIT WINDOW 'Выключен признак экспорта данных в таблицы SQL Server' + CHR(13) + ;
'Загрузка данных только в локальные таблицы' TIMEOUT 2

        CASE pr_import_loc = .F. AND pr_import_sql = .T. AND pr_import_sum = .F.  && Загрузка данных только в таблицы SQL Server

          WAIT WINDOW 'Выключен признак экспорта данных в таблицы SQL Server' + CHR(13) + ;
            'Загрузка данных только в таблицы SQL Server' TIMEOUT 2

        CASE pr_import_loc = .F. AND pr_import_sql = .F. AND pr_import_sum = .T.  && Загрузка данных в локальные таблицы и таблицы SQL Server

          WAIT WINDOW 'Выключен признак экспорта данных в таблицы SQL Server' + CHR(13) + ;
            'Загрузка данных в локальные таблицы и таблицы SQL Server' TIMEOUT 2

        OTHERWISE

          =err('В спарвочнике реквизитов неправильно настроен способ загрузки данных из головного банка ')

      ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE pr_rabota_sql = .T.  && Включен признак экспорта данных в таблицы SQL Server

      DO CASE
        CASE pr_import_loc = .T. AND pr_import_sql = .F. AND pr_import_sum = .F.  && Загрузка данных только в локальные таблицы

          WAIT WINDOW 'Включен признак экспорта данных в таблицы SQL Server' + CHR(13) + ;
            'Загрузка данных только в локальные таблицы' TIMEOUT 2

        CASE pr_import_loc = .F. AND pr_import_sql = .T. AND pr_import_sum = .F.  && Загрузка данных только в таблицы SQL Server

          WAIT WINDOW 'Включен признак экспорта данных в таблицы SQL Server' + CHR(13) + ;
            'Загрузка данных только в таблицы SQL Server' TIMEOUT 2

        CASE pr_import_loc = .F. AND pr_import_sql = .F. AND pr_import_sum = .T.  && Загрузка данных в локальные таблицы и таблицы SQL Server

          WAIT WINDOW 'Включен признак экспорта данных в таблицы SQL Server' + CHR(13) + ;
            'Загрузка данных в локальные таблицы и таблицы SQL Server' TIMEOUT 2

        OTHERWISE

          =err('В спарвочнике реквизитов неправильно настроен способ загрузки данных из головного банка ')

      ENDCASE

  ENDCASE

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* pr_registr = .T.
* sergeizhuk_flag = .F.

IF pr_registr = .T. AND sergeizhuk_flag = .F.
  DO vxod_parol IN Visa
ELSE
  DO vxod_admin IN Visa
ENDIF

RETURN


************************************************************ PROCEDURE rasvert *************************************************************************

PROCEDURE rasvert

title_visa = SPACE(2) + ALLTRIM(title_visa) + ' (филиал № ' + (num_branch) + ')'

dat = DT(DATE())

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*  AND sysmetric_cols = 1920 AND sysmetric_rows = 1080
*    num_schrift_menu_vert = 14
*    num_schrift_menu_gor = 12
*    num_schrift_popup = 12
*    num_schrift_windows = 12
*    colvo_rows = 45
*    colvo_cols = 250

IF sergeizhuk_flag = .T.

  DO CASE
    CASE ALLTRIM(sergeizhuk_name_pk) == 'PKSBZHUK'

      num_schrift_v = 14
      num_schrift_g = 12

      colvo_rows = 45
      colvo_cols = 265

    CASE ALLTRIM(sergeizhuk_name_pk) == 'NBSBZHUK'

      num_schrift_v = 12
      num_schrift_g = 11

      colvo_rows = 50
      colvo_cols = 210

  ENDCASE

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO formir_screen IN Visa
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DEFINE WINDOW poisk  AT 0,0 SIZE 2.8, colvo_cols - 10 FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE TITLE ' И н ф о р м а ц и я ' COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW poisk CENTER

DEFINE WINDOW brows  AT 0,0 SIZE  colvo_rows - 2, colvo_cols - 4 FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW brows CENTER

DEFINE WINDOW win_edit  AT 0,0 SIZE  colvo_rows - 2, colvo_cols - 60 FONT 'Courier New', num_schrift_g - 1  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW win_edit CENTER

DEFINE WINDOW doverenost  AT 0,0 SIZE  colvo_rows - 30, colvo_cols - 80 FONT 'Courier New', num_schrift_v - 1 STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW doverenost CENTER

DEFINE WINDOW brows_ostat  AT 0,0 SIZE  colvo_rows - 2, colvo_cols - 4 FONT 'Times New Roman', num_schrift_g - 1  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW brows_ostat CENTER

DEFINE WINDOW nazvan AT 0,0 SIZE 10, colvo_cols - 60 FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE TITLE ' Окно  просмотра дополнительной информации ' ;
  COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW nazvan CENTER

RETURN


************************************************************* PROCEDURE formir_screen *******************************************************************

PROCEDURE formir_screen

MODIF WINDOW SCREEN AT 0, 0 SIZE colvo_rows, colvo_cols ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' COLOR RGB(0,0,0,220,220,220)

_Screen.LockScreen = .T.

_Screen.ShowTips = .F.

_Screen.Caption = ALLTRIM(title_visa)
_Screen.Icon = (put_exe) + ('Files04.ico')

_Screen.AutoCenter = .T.
_Screen.WindowState = 0
_Screen.Visible = .T.
_Screen.Movable = .T.

_Screen.ControlBox = .T.
_Screen.MinButton = .T.
_Screen.MaxButton = .T.
_Screen.Closable = .F.

_Screen.LockScreen = .F.

CLEAR

RETURN


************************************************************** PROCEDURE select_data_odb ****************************************************************

PROCEDURE select_data_odb

SELECT data_odb
SET ORDER TO data

poisk_data_odb = DATE()
new_data_odb = DATE()

rec = DTOS(poisk_data_odb)

poisk = SEEK(rec)

IF poisk = .T.

  GO BOTTOM

  new_data_odb = data_odb.data

  SKIP -1
  arx_data_odb = data_odb.data

  SKIP  + 1

ELSE

  GO TOP
  GO BOTTOM

  new_data_odb = data_odb.data

  SKIP -1
  arx_data_odb = data_odb.data

  SKIP  + 1

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
* new_data_odb = CTOD('01.11.2011')
* arx_data_odb = CTOD('31.10.2011')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
colvo_den_god = GOMONTH(CTOD('01.01.' + ALLTRIM(STR(YEAR(new_data_odb)))), 12) - CTOD('01.01.' + ALLTRIM(STR(YEAR(new_data_odb))))
* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO poisk_den_kurs IN Servis WITH new_data_odb, 'USD'
DO poisk_den_kurs IN Servis WITH new_data_odb, 'EUR'

RETURN


**************************************************************** PROCEDURE select_screen_oboi ***********************************************************

PROCEDURE select_screen_oboi

DO CASE
  CASE INLIST(ALLTRIM(num_branch), '01', '03')

    poisk_bmp = .F.
    STORE SPACE(0) TO name_bmp

    num_den_god = new_data_odb - CTOD('01.01.' + ALLTRIM(STR(YEAR(new_data_odb))))  && Количество дней с начала года

    DO CASE
      CASE FILE((put_dbf) + ('screen_table.dbf')) = .T.  && Таблица с перечнем файлов для экранной заставки существует

        IF NOT USED('screen_table')

          SELECT 0
          USE (put_dbf) + ('screen_table.dbf')

        ENDIF

        SELECT screen_table
        SET ORDER TO num_den
        GO TOP

        poisk_bmp = SEEK(num_den_god)  && Ищем в таблице экранных файлов по номеру дня в течении года

* BROWSE WINDOW brows TITLE 'num_den_god = ' + ALLTRIM(STR(num_den_god))

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        DO CASE
          CASE poisk_bmp = .T.  && Запись по номеру дня года НАЙДЕНА

            DO CASE
              CASE EMPTY(ALLTRIM(screen_table.ima_file)) = .F.  && Запись найдена и поле с именем файла НЕ ПУСТОЕ

                IF BETWEEN(new_data_odb, CTOD('05.03.2010'), CTOD('09.03.2010'))

                  DO CASE
                    CASE BETWEEN(HOUR(DATETIME()), 11, 12)
                      name_bmp = 'Красивые цветы № 7.bmp'

                    CASE BETWEEN(HOUR(DATETIME()), 12, 13)
                      name_bmp = 'Красивые цветы № 6.bmp'

                    CASE BETWEEN(HOUR(DATETIME()), 13, 14)
                      name_bmp = 'Красивые цветы № 5.bmp'

                    CASE BETWEEN(HOUR(DATETIME()), 14, 15)
                      name_bmp = 'Красивые цветы № 4.bmp'

                    CASE BETWEEN(HOUR(DATETIME()), 15, 16)
                      name_bmp = 'Красивые цветы № 3.bmp'

                  ENDCASE

                ELSE
                  name_bmp = ALLTRIM(screen_table.ima_file) + ('.bmp')

                ENDIF

              CASE EMPTY(ALLTRIM(screen_table.ima_file)) = .T.  && Запись найдена но поле с именем файла ПУСТОЕ

                FOR nummer = 2 TO 50  && Запускаем цикл деления номера дня на 2, 3, 4, и т.д. и каждый раз делаем повторный поиск

                  SELECT screen_table
                  GO TOP

                  poisk_bmp = SEEK(ROUND(num_den_god / nummer, 0))

* BROWSE WINDOW brows TITLE 'num_den_god = ' + ALLTRIM(STR(num_den_god))

                  IF poisk_bmp = .T. AND EMPTY(ALLTRIM(screen_table.ima_file)) = .F.  && Если запись найдена и поле с именем файла НЕ ПУСТОЕ запоминаем его и выходим из цикла

                    name_bmp = ALLTRIM(screen_table.ima_file) + ('.bmp')
                    EXIT

                  ENDIF

                ENDFOR

* WAIT WINDOW 'name_bmp = ' + name_bmp TIMEOUT 3

                IF EMPTY(ALLTRIM(name_bmp)) = .T. && Если же после цикла переменная не содержит имени файла, то присваиваем имя Vodopad.bmp
                  name_bmp = 'Vodopad.bmp'
                ENDIF

            ENDCASE

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          CASE poisk_bmp = .F.  && Запись по номеру дня года НЕ НАЙДЕНА

            poisk_bmp = SEEK(DAY(new_data_odb))  && Ищем по номеру дня месяца

            DO CASE
              CASE poisk_bmp = .T.  && Запись по номеру дня НАЙДЕНА

                DO CASE
                  CASE EMPTY(ALLTRIM(screen_table.ima_file)) = .F.  && Запись найдена и поле с именем файла НЕ ПУСТОЕ

                    name_bmp = ALLTRIM(screen_table.ima_file) + ('.bmp')

                  CASE EMPTY(ALLTRIM(screen_table.ima_file)) = .T.  && Запись найдена но поле с именем файла ПУСТОЕ

                    name_bmp = 'Vodopad.bmp' && Если же поле не содержит имени файла, то присваиваем имя Vodopad.bmp

                ENDCASE

              CASE poisk_bmp = .F.  && Запись по номеру дня НЕ НАЙДЕНА

                name_bmp = 'Vodopad.bmp' && Если же запись не найдена, то присваиваем имя Vodopad.bmp

            ENDCASE

        ENDCASE

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE FILE((put_dbf) + ('screen_table.dbf')) = .F.  && Таблица с перечнем файлов для экранной заставки существует

        name_bmp = 'Vodopad.bmp'

    ENDCASE

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    DO CASE
      CASE FILE((put_bmp_net) + (name_bmp)) = .T.  && Проверяем физическое наличие выбранного файла на диске, файл СУЩЕСТВУЕТ

        @ 0.000,0.000 SAY LOCFILE((put_bmp_net) + (name_bmp)) BITMAP SIZE colvo_rows + 1.5, colvo_cols + 1 STRETCH STYLE 'T'

      CASE FILE((put_bmp_net) + (name_bmp)) = .F.  && Проверяем физическое наличие выбранного файла на диске, файл НЕ СУЩЕСТВУЕТ

        name_bmp = 'Vodopad.bmp' && Если же файл не существует, то присваиваем имя Vodopad.bmp

        DO CASE
          CASE FILE((put_bmp) + (name_bmp)) = .T.  && Проверяем физическое наличие выбранного файла на диске, файл СУЩЕСТВУЕТ

            @ 0.000,0.000 SAY LOCFILE((put_bmp) + (name_bmp)) BITMAP SIZE colvo_rows + 1.5, colvo_cols + 1 STRETCH STYLE 'T'

          CASE FILE((put_bmp) + (name_bmp)) = .F.  && Проверяем физическое наличие выбранного файла на диске, файл НЕ СУЩЕСТВУЕТ
            CLEAR  && В случае отсутствия файла экранную заставку не формируем, а просто главное окно очищаем
        ENDCASE

    ENDCASE

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE NOT INLIST(ALLTRIM(num_branch), '01', '03')

    name_bmp = 'Vodopad.bmp'

    @ 0.000,0.000 SAY LOCFILE((put_bmp) + (name_bmp)) BITMAP SIZE colvo_rows + 1.5, colvo_cols + 1 STRETCH STYLE 'T'

ENDCASE

RETURN


************************************************************** PROCEDURE svert ************************************************************************

PROCEDURE svert

IF RIGHT(SYS(16), 3) <> 'EXE'
  MODIFY WINDOW SCREEN
ENDIF

RETURN


********************************************************** PROCEDURE vxod_admin **********************************************************************

PROCEDURE vxod_admin

SELECT * ;
  FROM parol ;
  WHERE ALLTRIM(parol.kod) == '99' ;
  INTO CURSOR sel_parol

SELECT sel_parol
GO TOP

wes = sel_parol.pravo
adm_kod = ALLTRIM(sel_parol.kod)
par_kod = ALLTRIM(sel_parol.kod)
fio_user = ALLTRIM(sel_parol.fio)
par_fio = ALLTRIM(sel_parol.fio)
par_ima = ALLTRIM(sel_parol.ima)
par_par = ALLTRIM(sel_parol.parol)
par_prim = ALLTRIM(sel_parol.prim)
par_prompt = sel_parol.pr_prompt
par_grup = ALLTRIM(sel_parol.grup)
kod_kassa = ALLTRIM(sel_parol.kassa)

pr_registr = .F.  && Признак регистрации в ПО "CARD-виза" выключается

*  par_grup = 'RUB'
*  par_grup = 'VAL'
*  par_grup = 'ALL'

* BROWSE WINDOW brows

SELECT parol
SET ORDER TO ima

IF SEEK(ALLTRIM(par_ima))
  REPLACE use_soft WITH .T., colvo_soft WITH colvo_soft + 1
ENDIF

DO vxod IN Visa

RETURN


*********************************************************** PROCEDURE vxod_servis **********************************************************************

PROCEDURE vxod_servis

SELECT parol
REPLACE ALL use_soft WITH .F., colvo_soft WITH 0
GO TOP

SELECT * ;
  FROM parol ;
  WHERE ALLTRIM(parol.kod) == '99' ;
  INTO CURSOR sel_parol

SELECT sel_parol
GO TOP

wes = sel_parol.pravo
adm_kod = ALLTRIM(sel_parol.kod)
par_kod = ALLTRIM(sel_parol.kod)
fio_user = ALLTRIM(sel_parol.fio)
par_fio = ALLTRIM(sel_parol.fio)
par_ima = ALLTRIM(sel_parol.ima)
par_par = ALLTRIM(sel_parol.parol)
par_prim = ALLTRIM(sel_parol.prim)
par_prompt = sel_parol.pr_prompt
par_grup = ALLTRIM(sel_parol.grup)
kod_kassa = ALLTRIM(sel_parol.kassa)

pr_registr = .F.  && Признак регистрации в ПО "CARD-виза" выключается

SELECT parol
SET ORDER TO ima
GO TOP

IF SEEK(ALLTRIM(par_ima))
  REPLACE use_soft WITH .T., colvo_soft WITH colvo_soft + 1
ENDIF

RETURN


**************************************************************** PROCEDURE vxod ***********************************************************************

PROCEDURE vxod

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO zapis_dolg_fio_ispolnit IN Scan_rekvizit  && В зависимости от состояния переменной признак регистрации происходит перепресвоение значения
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Внимание! Дата текущего опердня ' + DT(new_data_odb) + '  Дата архивного опердня ' + DT(arx_data_odb),WCOLS())
=INKEY(2)
DEACTIVATE WINDOW poisk

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC fff, nnn

*  001-Москва,
*  002-Ленск,
*  003-Орел,
*  004-Мирный,
*  005-Удачный,
*  006-Якутск,
*  007-Айхал,
*  008-Архангельск,
*  009-Краснодар,
*  010-Новосибирск,
*  011-Иркутск,
*  013-Головной офис

fff = new_branch(num_branch)  && fff - Отделение (филиал) банка

nnn = PADL(ALLTRIM(STR((new_data_odb - CTOD('01.01.' + ALLTRIM(STR(YEAR(new_data_odb)))) + 1))), 3, '0')  && nnn - Юлианская дата (количество дней с начала года)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
tit = SPACE(2) + ALLTRIM(title_visa) + '  Дата операционного дня ' + DT(new_data_odb) + '  Версия ПО - ' + LEFT(ALLTRIM(VERSION()), 16)

_Screen.Caption = tit
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF pr_rabota_sql = .T.  && Если .T., то ПО работает с таблицами SQL Server, если .F., то ПО работает с локальными таблицами

  IF sergeizhuk_flag = .T.

    text1 = 'Подключиться к локальному SQL Server 2005'
    text2 = 'Подключиться к локальному SQL Server 2008'
    l_bar = 3
    =popup_9(text1,text2,text3,text4,l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      put_select = IIF(BAR() <= 0, 1, BAR())

      DO CASE
        CASE put_select = 1  && Подключиться к локальному SQL Server 2005

          DO vxod_sql_server_loc_2005 IN Vxod_sql

        CASE put_select = 3  && Подключиться к локальному SQL Server 2008

          DO vxod_sql_server_loc_2008 IN Vxod_sql

      ENDCASE

    ELSE
      DO vxod_sql_server_loc_2005 IN Vxod_sql
    ENDIF

  ENDIF
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* sergeizhuk_flag = .F.

IF sergeizhuk_flag = .F.
* DO export IN Obsluga_god
ENDIF

* sergeizhuk_flag = .T.

IF sergeizhuk_flag = .F.
  DO start_sms IN Smshome
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO WHILE .T.

  ACTIVATE POPUP Glmenu

  punkt = IIF(BAR() <= 0, 1, BAR())

  IF pr_dostup = .F.  && Признак разрешения на вход в ПО
    EXIT
  ENDIF

  IF punkt = 21
    EXIT
  ENDIF

  IF LASTKEY() = 27 AND sergeizhuk_flag = .T.
    EXIT
  ENDIF

ENDDO

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
IF pr_rabota_sql = .T.  && Если .T., то ПО работает с таблицами SQL Server, если .F., то ПО работает с локальными таблицами
  DO exit_sql_server IN Visa
ENDIF
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

RETURN


************************************************************ PROCEDURE exit_sql_server *******************************************************************

PROCEDURE exit_sql_server

IF sql_num > 0

  DO WHILE SQLDISCONNECT(sql_num) < 0
    =SQLCANCEL(sql_num)
  ENDDO

ENDIF

RETURN


************************************************************* PROCEDURE Glmenu ***********************************************************************

PROCEDURE glmenu
popup_ima = LOWER(POPUP())
bar_glmenu = BAR()
HIDE POPUP (popup_ima)

SELECT rekvizit
GO TOP

DO CASE
  CASE ALLTRIM(par_kod) == '99'
    pr_dostup = .T.

  CASE ALLTRIM(par_kod) <> '99'
    pr_dostup = rekvizit.dostup

ENDCASE

* WAIT WINDOW 'pr_dostup = ' + IIF(pr_dostup = .T., '.T.', '.F.') TIMEOUT 3

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO CASE
  CASE bar_glmenu = 1  && Движение средств по пластиковой карте

    IF pr_dostup = .T.  && Признак разрешения на вход в ПО

      tit = ' Возможности: печать распоряжений, экспорт в ОДБ, ведомость остатков, история остатков, аналитика, справочники'
      _Screen.Caption = tit

      ACTIVATE MENU document

    ELSE
      =err('Программный комплекс используется в монопольном режиме, ВХОД ЗАПРЕЩЕН ... ')
      DEACTIVATE POPUP glmenu
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE bar_glmenu = 3  && Выдача карт и управление карточным счетом

    IF pr_dostup = .T.  && Признак разрешения на вход в ПО

      tit = ' Возможности: заявление на карту, пополнение, списание, закрытие, управление, экспорт в ОДБ, справочники'
      _Screen.Caption = tit

      IF pr_zarpl = .T.

        DO WHILE .T.

          ACTIVATE POPUP zarpl_proekt

          IF NOT LASTKEY() = 27

            IF INLIST(LASTKEY(), 4, 19)
              LOOP
            ELSE
              DO zarpl_proekt IN Visa
              EXIT
            ENDIF

          ELSE

            num_proekt = '00'  && Номер зарплатного проекта

            =soob('Внимание! Ваш выбор - Главный справочник карточных счетов филиала - код проекта "00"')

            EXIT

          ENDIF

        ENDDO

      ENDIF

      ACTIVATE MENU client

    ELSE
      =err('Программный комплекс используется в монопольном режиме, ВХОД ЗАПРЕЩЕН ... ')
      DEACTIVATE POPUP glmenu
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE bar_glmenu = 5  && Создание и формирование отчетов в УПК

    IF pr_dostup = .T.  && Признак разрешения на вход в ПО

      SET ESCAPE ON

      tit = ' Возможности: формирование отчетной документации для УПК, справочники'
      _Screen.Caption = tit

* ------ Ограничение общего количества записей по номеру выбранного кода проекта (вставил Крогульский М.А.) --------------------------------------

      IF pr_zarpl = .T.

        DO WHILE .T.

          ACTIVATE POPUP zarpl_proekt

          IF NOT LASTKEY() = 27

            IF INLIST(LASTKEY(), 4, 19)
              LOOP
            ELSE
              DO zarpl_proekt IN visa
              EXIT
            ENDIF

          ELSE

            num_proekt = '00'  && Номер зарплатного проекта

            =soob('Внимание! Ваш выбор - Главный справочник карточных счетов филиала - код проекта "00"')

            EXIT

          ENDIF

        ENDDO

      ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      ACTIVATE MENU otchet_upk

    ELSE
      =err('Программный комплекс используется в монопольном режиме, ВХОД ЗАПРЕЩЕН ... ')
      DEACTIVATE POPUP glmenu
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE bar_glmenu = 7  && Начисление процентов по карточному счету

    IF pr_dostup = .T.  && Признак разрешения на вход в ПО

      SET ESCAPE ON

      tit = ' Возможности: расчет процентов по счетам 42301, 42308, печать, отправка в УПК, экспорт в ОДБ, справочники'
      _Screen.Caption = tit

      PUBLIC new_ima_dt_null, new_ima_sum_null

      IF pr_zarpl = .T.

        DO WHILE .T.

          ACTIVATE POPUP zarpl_proekt

          IF NOT LASTKEY() = 27

            IF INLIST(LASTKEY(), 4, 19)
              LOOP
            ELSE
              DO zarpl_proekt IN Visa
              EXIT
            ENDIF

          ELSE

            num_proekt = '00'  && Номер зарплатного проекта

            =soob('Внимание! Ваш выбор - Главный справочник карточных счетов филиала - код проекта "00"')

            EXIT

          ENDIF

        ENDDO

      ENDIF

      DO use_table_proz IN Visa

      IF tiplog = .T.

        ACTIVATE MENU proz_ckc

      ENDIF

      DO close_table_proz IN Visa

      RELEASE new_ima_dt_null, new_ima_sum_null

      SET ESCAPE OFF

    ELSE
      =err('Программный комплекс используется в монопольном режиме, ВХОД ЗАПРЕЩЕН ... ')
      DEACTIVATE POPUP glmenu
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE bar_glmenu = 9  && Загрузка данных из головного банка

    SET ESCAPE ON

    IF pr_dostup = .T.  && Признак разрешения на вход в ПО

      tit = ' Возможности: проверка подписи, снятие подписи, импорт данных из головного банка об операциях по картам'
      _Screen.Caption = tit

      DO start IN Import

    ELSE
      =err('Программный комплекс используется в монопольном режиме, ВХОД ЗАПРЕЩЕН ... ')
      DEACTIVATE POPUP glmenu
    ENDIF

    SET ESCAPE OFF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE bar_glmenu = 11  && Выгрузка данных в головной банк

    IF pr_dostup = .T.  && Признак разрешения на вход в ПО

      tit = ' Возможности: экспорт основных таблиц с данными составляющих систему в головной банк'
      _Screen.Caption = tit

      SET ESCAPE ON

      log_quit = .F.

      =soob('Внимание! Выполнение данной процедуры требует монопольного режима работы с таблицами ... ')

      tex1 = 'В ы п о л н и т ь  процедуру выгрузки данных в головной банк'
      tex2 = 'О т к а з а т ь с я  от процедуры выгрузки данных в головной банк'
      l_bar = 3
      =popup_9(tex1,tex2,tex3,tex4,l_bar)
      ACTIVATE POPUP vibor
      RELEASE POPUPS vibor

      IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

        IF BAR() = 1  && В ы п о л н и т ь  процедуру закрытия опердня

          DO vxod_monopol IN Visa

          IF error_monopol = .F.

            DO start IN Export_golova

          ENDIF
        ENDIF
      ENDIF

      SET ESCAPE OFF

    ELSE
      =err('Программный комплекс используется в монопольном режиме, ВХОД ЗАПРЕЩЕН ... ')
      DEACTIVATE POPUP glmenu
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE bar_glmenu = 13  && Сервисные и системные процедуры

    IF pr_dostup = .T.  && Признак разрешения на вход в ПО

      tit = ' Возможности: работа со справочниками, импорт справочника БИК, импорт курсов валют, обслуживание и изменение таблиц'
      _Screen.Caption = tit

      ACTIVATE MENU servis_glav

    ELSE
      =err('Программный комплекс используется в монопольном режиме, ВХОД ЗАПРЕЩЕН ... ')
      DEACTIVATE POPUP glmenu
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE bar_glmenu = 15  && Резервное копирование данных

    IF pr_dostup = .T.  && Признак разрешения на вход в ПО

      tit = ' Возможности: выполнение резервного копирования составляющих таблиц в резервную директорию'
      _Screen.Caption = tit

      tim_1_all = DATETIME()

      ACTIVATE WINDOW poisk
      @ WROWS()/3,3 SAY PADC('Внимание! Начато резервное копирование таблиц ПО "CARD-VISA филиал" ...',WCOLS())

      DO usebase IN Order_kassa
      DO arxiv  IN Order_kassa

      DO usebase IN Order_memo
      DO arxiv IN Order_memo

      DO usebase IN Plateg
      DO arxiv IN Plateg

      DO arxiv IN Visa

      tim_2_all = DATETIME()

      @ WROWS()/3,3 SAY PADC('Резервное копирование таблиц ПО "CARD-VISA филиал" успешно завершено.' + ' (время = ' + ALLTRIM(STR(tim_2_all - tim_1_all)) + ' сек.)',WCOLS())
      =INKEY(2)
      DEACTIVATE WINDOW poisk

    ELSE
      =err('Программный комплекс используется в монопольном режиме, ВХОД ЗАПРЕЩЕН ... ')
      DEACTIVATE POPUP glmenu
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE bar_glmenu = 17  && Закрытие операционного дня в "CARD-VISA"

    IF pr_dostup = .T.  && Признак разрешения на вход в ПО

      tit=' Возможности: закрытие операционного дня банка в ПО "CARD-VISA и переход на работу с новой датой"'
      _Screen.Caption = tit

      DO start IN Zakritie_odb

    ELSE
      =err('Программный комплекс используется в монопольном режиме, ВХОД ЗАПРЕЩЕН ... ')
      DEACTIVATE POPUP glmenu
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE bar_glmenu = 19  && Доступ к справочнику паролей

    IF pr_dostup = .T.  && Признак разрешения на вход в ПО

      tit = ' Возможности: работа со справочником паролей, изменение, добавление, удаление из списка операторов'
      _Screen.Caption = tit

      DO vidparol IN Visa

    ELSE
      =err('Программный комплекс используется в монопольном режиме, ВХОД ЗАПРЕЩЕН ... ')
      DEACTIVATE POPUP glmenu
    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE bar_glmenu = 21  && Завершение работы с программой

    DEACTIVATE POPUP glmenu

ENDCASE

tit = SPACE(2) + ALLTRIM(title_visa) + '  Дата операционного дня ' + DT(new_data_odb) + '  Версия ПО - ' + LEFT(ALLTRIM(VERSION()), 16)

_Screen.Caption = tit

RETURN


************************************************************* PROCEDURE zarpl_proekt *******************************************************************

PROCEDURE zarpl_proekt

DO CASE
  CASE ALLTRIM(num_branch) == '01'


* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE ALLTRIM(num_branch) == '02'

    DO CASE
      CASE BAR() = 1  && Главный справочник карточных счетов филиала - код проекта "00"
        num_proekt = '00'

      CASE BAR() = 3 && Корпоративный проект - Visa Business (RUB) - "ПУ Алмаздортранс АК Алроса" - код проекта "21"
        num_proekt = '21'

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE ALLTRIM(num_branch) == '03'

    DO CASE
      CASE BAR() = 1  && Главный справочник карточных счетов филиала - код проекта "00"
        num_proekt = '00'

      CASE BAR() = 3  && Номер зарплатного проекта - счета по зарплатному проекту "ЕСО" - код проекта "01"
        num_proekt = '01'

      CASE  BAR() = 5  && Номер зарплатного проекта - счета по зарплатному проекту "Алроса - торг" - код проекта "02"
        num_proekt = '02'

      CASE  BAR() = 7  && Номер зарплатного проекта - счета по зарплатному проекту "Бриллианты России" - код проекта "03"
        num_proekt = '03'

      CASE  BAR() = 9  && Номер зарплатного проекта - счета по зарплатному проекту "Алроса - представительство" - код проекта "04"
        num_proekt = '04'

      CASE  BAR() = 11  && Номер зарплатного проекта - счета по зарплатному проекту "Главный офис ЗАО Гостиницы АЛРОСА" - код проекта "05"
        num_proekt = '05'

      CASE  BAR() = 13  && Номер зарплатного проекта - счета по зарплатному проекту "Филиал ЗАО Гостиницы АЛРОСА" - код проекта "06"
        num_proekt = '06'

      CASE  BAR() = 15  && Номер зарплатного проекта - счета по зарплатному проекту "ЗАО АЛРОСА- АФРИКА" - код проекта "07"
        num_proekt = '07'

      CASE  BAR() = 17  && Номер зарплатного проекта - счета по корпоративному проекту - Visa Business (RUB) - НПФ "Алмазная Осень" - код проекта "21"
        num_proekt = '21'

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE ALLTRIM(num_branch) == '04'


* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE ALLTRIM(num_branch) == '05'

    DO CASE
      CASE BAR() = 1  && Главный справочник карточных счетов филиала - код проекта "00"
        num_proekt = '00'

      CASE BAR() = 3  && Счета по зарплатному проекту "ГУ "ДРСО ЖКХиЭ"" - код проекта "01"
        num_proekt = '01'

      CASE BAR() = 5  && Счета по зарплатному проекту "ИП Ушаков Алексей Владимирович" - код проекта "02"
        num_proekt = '02'

      CASE  BAR() = 7  && Счета по зарплатному проекту "Якутский филиал ООО СК "АЛРОСА" - код проекта "03"
        num_proekt = '03'

      CASE  BAR() = 9  && Счета по зарплатному проекту "ЗАО "АЛРОСА" - код проекта "04"
        num_proekt = '04'

      CASE  BAR() = 11  && Счета по зарплатному проекту "Якутское предприятие по торговле алмазами АК "Алроса"" - код проекта "05"
        num_proekt = '05'

      CASE  BAR() = 13  && Счета по корпоративному проекту Якутский филиал ООО СК "АЛРОСА" - код проекта "26"
        num_proekt = '26'

      CASE  BAR() = 15  && Счета по корпоративному проекту ООО "Благодать" - код проекта "27"
        num_proekt = '27'

      CASE  BAR() = 17  && Счета по корпоративному проекту ООО "Строй Мастер" - код проекта "28"
        num_proekt = '28'

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE ALLTRIM(num_branch) == '06'


* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE ALLTRIM(num_branch) == '07'


* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE ALLTRIM(num_branch) == '08'

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE ALLTRIM(num_branch) == '09'

    DO CASE
      CASE BAR() = 1  && Главный справочник карточных счетов филиала - код проекта "00"
        num_proekt = '00'

      CASE BAR() = 3 && Счета по зарплатному проекту ООО "Новый Дом" - код проекта "01"
        num_proekt = '01'

      CASE BAR() = 5 && Счета по зарплатному проекту "Представительство АК АЛРОСА в г. Орле" - код проекта "02"
        num_proekt = '02'

      CASE BAR() = 7 && Счета по зарплатному проекту ООО "Снабсервис - Т" - код проекта "03"
        num_proekt = '03'

      CASE BAR() = 9 && Счета по зарплатному проекту ООО "НПП "АСТРОН ЭЛЕКТРОНИКА"
        num_proekt = '04'

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE ALLTRIM(num_branch) == '10'

    DO CASE
      CASE BAR() = 1  && Главный справочник карточных счетов филиала - код проекта "00"
        num_proekt = '00'

      CASE BAR() = 3  && Номер зарплатного проекта - счета по зарплатному проекту "ОАО "Пур-Наволок Отель"" - код проекта "01"
        num_proekt = '01'

      CASE  BAR() = 5  && Номер корпоративного проекта - счета по корпоративному проекту - Visa Business (RUB) - код проекта "21"
        num_proekt = '21'

      CASE  BAR() = 7  && Номер корпоративного проекта - счета по корпоративному проекту - Visa Business (RUB) - код проекта "22"
        num_proekt = '22'

      CASE BAR() = 9  && Номер зарплатного проекта - счета по зарплатному проекту "ОАО "Строительно монтажный трест № 1" - код проекта "02"
        num_proekt = '02'

      CASE BAR() = 11  && Номер зарплатного проекта - счета по зарплатному проекту "ОАО "Строительно монтажный трест № 2" - код проекта "03"
        num_proekt = '03'

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE ALLTRIM(num_branch) == '11'

    DO CASE
      CASE BAR() = 1  && Главный справочник карточных счетов филиала - код проекта "00"
        num_proekt = '00'

      CASE BAR() = 3  && Номер зарплатного проекта - счета по зарплатному проекту ООО "Юг-Автомаг" - код проекта "01"
        num_proekt = '01'

      CASE  BAR() = 5  && Счета по "ИП Кочубей Юрий Владимирович" - Visa Business (RUB) - код проекта "21"
        num_proekt = '21'

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE ALLTRIM(num_branch) == '12'

    DO CASE
      CASE BAR() = 1  && Главный справочник карточных счетов филиала - код проекта "00"
        num_proekt = '00'

      CASE BAR() = 3  && Счета по зарплатному проекту ООО "Графика+" - код проекта "01"
        num_proekt = '01'

      CASE BAR() = 5 && Корпоративный проект - Visa Business (RUB) - ООО "Графика+" - код проекта "21"
        num_proekt = '21'

      CASE BAR() = 7 && Корпоративный проект - Visa Business (RUB) - "Контора № 2" - код проекта "22"
        num_proekt = '22'

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

ENDCASE

title_zarpl = ALLTRIM(PROMPT())

* WAIT WINDOW 'num_proekt = ' + (num_proekt)

=soob('Внимание! Ваш выбор - ' + ALLTRIM(title_zarpl))

RETURN


************************************************************ PROCEDURE use_table_proz *******************************************************************

PROCEDURE use_table_proz

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF NOT USED('acc_cln_dt_pol')

  IF FILE((put_dbf) + ('acc_cln_dt.dbf'))

    IF FILE((put_dbf) + ('acc_cln_dt.cdx'))

      SELECT 0
      USE (put_dbf) + ('acc_cln_dt.dbf') ALIAS acc_cln_dt_pol ORDER TAG card_nls

    ELSE
      fail = (put_dbf) + ('acc_cln_dt.cdx')
      =err('Внимание! Нужный для работы файл - ' + ALLTRIM(fail) + ' не найден.')
      tiplog = .F.
    ENDIF

  ELSE
    fail = (put_dbf) + ('acc_cln_dt.dbf')
    =err('Внимание! Нужный для работы файл - ' + ALLTRIM(fail) + ' не найден.')
    tiplog = .F.
  ENDIF

ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF NOT USED('acc_cln_sum_pol')

  IF FILE((put_dbf) + ('acc_cln_sum.dbf'))

    IF FILE((put_dbf) + ('acc_cln_sum.cdx'))

      SELECT 0
      USE (put_dbf) + ('acc_cln_sum.dbf') ALIAS acc_cln_sum_pol ORDER TAG card_nls

    ELSE
      fail = (put_dbf) + ('acc_cln_sum.cdx')
      =err('Внимание! Нужный для работы файл - ' + ALLTRIM(fail) + ' не найден.')
      tiplog = .F.
    ENDIF

  ELSE
    fail = (put_dbf) + ('acc_cln_sum.dbf')
    =err('Внимание! Нужный для работы файл - ' + ALLTRIM(fail) + ' не найден.')
    tiplog = .F.
  ENDIF

ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF NOT USED('acc_strah_dt')

  IF FILE((put_dbf) + ('acc_strah_dt.dbf'))

    IF FILE((put_dbf) + ('acc_strah_dt.cdx'))

      SELECT 0
      USE (put_dbf) + ('acc_strah_dt.dbf') ALIAS acc_strah_dt ORDER TAG card_data

    ELSE
      fail = (put_dbf) + ('acc_strah_dt.cdx')
      =err('Внимание! Нужный для работы файл - ' + ALLTRIM(fail) + ' не найден.')
      tiplog = .F.
    ENDIF

  ELSE
    fail = (put_dbf) + ('acc_strah_dt.dbf')
    =err('Внимание! Нужный для работы файл - ' + ALLTRIM(fail) + ' не найден.')
    tiplog = .F.
  ENDIF

ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF NOT USED('acc_strah_sum')

  IF FILE((put_dbf) + ('acc_strah_sum.dbf'))

    IF FILE((put_dbf) + ('acc_strah_sum.cdx'))

      SELECT 0
      USE (put_dbf) + ('acc_strah_sum.dbf') ALIAS acc_strah_sum ORDER TAG card_data

    ELSE
      fail = (put_dbf) + ('acc_strah_sum.cdx')
      =err('Внимание! Нужный для работы файл - ' + ALLTRIM(fail) + ' не найден.')
      tiplog = .F.
    ENDIF

  ELSE
    fail = (put_dbf) + ('acc_strah_sum.dbf')
    =err('Внимание! Нужный для работы файл - ' + ALLTRIM(fail) + ' не найден.')
    tiplog = .F.
  ENDIF

ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

RETURN


************************************************************* PROCEDURE close_table_proz *****************************************************************

PROCEDURE close_table_proz

IF USED('acc_cln_dt_pol')
  SELECT acc_cln_dt_pol
  USE
ENDIF

IF USED('acc_cln_sum_pol')
  SELECT acc_cln_sum_pol
  USE
ENDIF

IF USED('acc_strah_dt')
  SELECT acc_strah_dt
  USE
ENDIF

IF USED('acc_strah_sum')
  SELECT acc_strah_sum
  USE
ENDIF

RETURN


************************************************************* PROCEDURE vidparol **********************************************************************

PROCEDURE vidparol

SELECT parol

SELECT * ;
  FROM parol ;
  INTO CURSOR sel_parol

IF _TALLY <> 0

  DEFINE WINDOW vidparol AT 0, 0 SIZE colvo_rows - 10, colvo_cols - 4 FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
    NOFLOAT NOCLOSE NOMINIMIZE DOUBLE TITLE ' Окно  редактирования  паролей ' COLOR RGB(0,0,0,192,192,192)
  MOVE WINDOW vidparol CENTER

  PUSH KEY CLEAR

  ON KEY LABEL CTRL+INS DO ins IN Visa
  ON KEY LABEL CTRL+DEL DO del IN Visa

  SET CURSOR ON

  IF wes <> 5

    SELECT parol
    SET FILTER TO ALLTRIM(kod) == ALLTRIM(par_kod)
    GO TOP

    BROWSE FIELDS ;
      kod:H = 'Код пользователя':R,;
      fio:H = 'Пользователь системы',;
      ima:H = 'Системное имя',;
      parol:H = 'Системный пароль',;
      use_soft:H = 'Активность',;
      colvo_soft:H = 'Кол-во копий',;
      pravo:H = 'Право доступа',;
      grup:H = 'Доступ к счетам',;
      pr_prompt:H = 'Признак печати',;
      kassa:H = 'Касса',;
      prim:H = 'Должность' ;
      WINDOW vidparol ;
      TITLE '  Переход по полям - Tab  Запись - Enter ' NOAPPEND NODELETE

    SET FILTER TO

  ELSE

    SELECT parol
    GO TOP

    BROWSE FIELDS ;
      kod:H = 'Код пользователя',;
      fio:H = 'Пользователь системы',;
      ima:H = 'Системное имя',;
      parol:H = 'Системный пароль',;
      use_soft:H = 'Активность',;
      colvo_soft:H = 'Кол-во копий',;
      pravo:H = 'Право доступа',;
      grup:H = 'Доступ к счетам',;
      pr_prompt:H = 'Признак печати',;
      kassa:H = 'Касса',;
      prim:H = 'Должность' ;
      WINDOW vidparol ;
      TITLE '  Переход по полям - Tab  Запись - Enter ' + ' Добавить запись - Ctrl + Insert   Удалить запись - Ctrl + Delete  '

  ENDIF

  POP KEY ALL

  RELEASE WINDOW vidparol

  SELECT * ;
    FROM parol ;
    WHERE ALLTRIM(kod) == ALLTRIM(par_kod) ;
    INTO CURSOR sel_parol

* BROWSE WINDOW brows

  wes = sel_parol.pravo
  adm_kod = ALLTRIM(sel_parol.kod)
  par_kod = ALLTRIM(sel_parol.kod)
  fio_user = ALLTRIM(sel_parol.fio)
  par_fio = ALLTRIM(sel_parol.fio)
  par_ima = ALLTRIM(sel_parol.ima)
  par_par = ALLTRIM(sel_parol.parol)
  par_prim = ALLTRIM(sel_parol.prim)
  par_prompt = sel_parol.pr_prompt
  par_grup = ALLTRIM(sel_parol.grup)
  kod_kassa = ALLTRIM(sel_parol.kassa)

ELSE
  =err('Внимание! В базе данных записей не обнаружено.')
ENDIF

SET CURSOR OFF

RETURN


********************************************************* PROCEDURE vxod_monopol **********************************************************************

PROCEDURE vxod_monopol

SELECT IIF(use_soft = .T., 'Активен', 'Не активен') AS use_soft, colvo_soft, kod, fio, grup ;
  FROM parol ;
  WHERE use_soft = .T. AND ALLTRIM(ima) <> ALLTRIM(par_ima) ;
  UNION ;
  SELECT IIF(use_soft = .T., 'Активен', 'Не активен') AS use_soft, colvo_soft, kod, fio, grup ;
  FROM parol ;
  WHERE use_soft = .T. AND ALLTRIM(ima) == ALLTRIM(par_ima) AND colvo_soft > 1 ;
  INTO CURSOR sel_monopol

* BROWSE WINDOW brows

IF _TALLY <> 0

  DEFINE WINDOW vidparol AT 0, 0 SIZE colvo_rows - 25, colvo_cols - 60 FONT 'Times New Roman', num_schrift_v + 1  STYLE 'B' ;
    NOFLOAT NOCLOSE NOMINIMIZE DOUBLE TITLE ' Окно  редактирования  паролей ' COLOR RGB(0,0,0,192,192,192)
  MOVE WINDOW vidparol CENTER

  error_monopol = .T.

  =err('Внимание! Монопольный доступ не возможен, приложение открыто на других станциях ... ')

  SELECT sel_monopol
  GO TOP

  BROWSE FIELDS ;
    use_soft:H = 'Активность',;
    colvo_soft:H = 'Кол-во копий',;
    kod:H = 'Код пользователя',;
    fio:H = 'Пользователь системы',;
    grup:H = 'Доступ к счетам' ;
    WINDOW vidparol ;
    TITLE '  Список пользователей работающих с программным обеспечением'

  RELEASE WINDOW vidparol

ELSE
  error_monopol = .F.
ENDIF

RETURN


**************************************************************** PROCEDURE ins ************************************************************************

PROCEDURE ins
APPEND BLANK
KEYBOARD '{ENTER}'
DEFINE WINDOW ins AT 0,0 SIZE 2,60 FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE TITLE ' Ввод данных в справочник ' COLOR RGB(0,0,0,192,192,192)
MOVE WINDOW ins CENTER
ACTIVATE WINDOW ins
zap = '     Внимание! Запись добавлена в базу.      '
@ WROWS()/3-0.2 ,6 GET zap
READ
=INKEY(2)
CLEAR GETS
RELEASE WINDOW ins
RETURN


*************************************************************** PROCEDURE del *************************************************************************

PROCEDURE del
DELETE NEXT 1
KEYBOARD '{ENTER}'
DEFINE WINDOW del AT 0,0 SIZE 2,60 FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE TITLE ' Удаление данных из справочника ' COLOR RGB(0,0,0,192,192,192)
MOVE WINDOW del CENTER
ACTIVATE WINDOW del
zap = '      Внимание! Запись помечена на удаление.      '
@ WROWS()/3-0.2 ,6 GET zap
READ
=INKEY(2)
CLEAR GETS
RELEASE WINDOW del
RETURN


************************************************************* PROCEDURE usebase ***********************************************************************

PROCEDURE usebase

IF NOT DBUSED('visa')
  OPEN DATABASE (put_dbf) + ('visa.dbc')
ENDIF

DO WHILE TXNLEVEL() > 0
  ROLLBACK
ENDDO

faildbfima = 'slovar_visa'
aliasscan = 'slovar_visa'
faildbf = (direktoria) + ('dbf\') + (faildbfima) + ('.dbf')

IF FILE(faildbf)

  IF NOT USED(aliasscan)
    SELECT 0
    USE (faildbfima) ALIAS (aliasscan)
  ENDIF

  SELECT (aliasscan)

  IF error_1705 = .T.
    DO exit IN Visa
    CANCEL
  ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  SCAN FOR pr_use = .T. AND loc_sql = .T.

    faildbfima = ALLTRIM(slovar_visa.dbfname)
    aliasima = ALLTRIM(slovar_visa.alias)
    failcdxima = ALLTRIM(slovar_visa.cdxname)

    faildbf = (put_dbf) + (faildbfima) + ('.dbf')
    failcdx = (put_dbf) + (faildbfima) + ('.cdx')

    IF FILE(faildbf) = .T.
      IF FILE(failcdx) = .T.

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        IF flag_add_table = .T. AND NOT INLIST(ALLTRIM(faildbfima), 'rekvizit', 'ostat_upk', 'obsluga_god', 'oborot_ved', 'oborot_sum', 'overdraft', 'istor_overdraft', ;
            'istor_ost', 'istor_mak', 'forma_1416_u', 'forma_1417_u', 'forma_1991_u', 'vedom_ost', 'zajava_sdacha', 'tarif_visa')  && Признак лечения баз данных, сначала таблица исключается из базы, затем добавляется.
          REMOVE TABLE (faildbfima)
          ADD TABLE (put_dbf) + (faildbfima) + ('.dbf')
        ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
* WAIT WINDOW TIMEOUT 1 '  Имя таблицы - ' + (faildbfima) + '.dbf' + '  Имя индексного тэга - ' + (failcdxima) + SPACE(2)
* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        IF NOT USED(aliasima)
          SELECT 0
          USE (faildbf) ALIAS (aliasima)
        ELSE
          SELECT (aliasima)
        ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        SELECT (aliasima)
        GO TOP

        FOR I = 1 TO 254

          IF EMPTY(TAG(I)) = .F.

            IF LOWER(TAG(I)) == LOWER(failcdxima)
*            WAIT WINDOW TIMEOUT 2 '  Имя таблицы - ' + (faildbfima) + '.dbf' + '  Имя индексного тэга - ' + (failcdxima) + SPACE(2)
              USE (faildbf) ALIAS (aliasima) ORDER TAG (failcdxima)
              EXIT
            ENDIF

          ELSE

            USE (faildbf) ALIAS (aliasima)
            =err('Внимание! Необходимый индексный тэг ' + (failcdxima) + ' для таблицы ' + (faildbfima) + '.dbf' + ' не найден.')
            EXIT
          ENDIF

        ENDFOR

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*          sel = (aliasima)
*          I = 1

*          DO WHILE .T.

*            strokakey = ALLTRIM(KEY(I))
*            strokatag = ALLTRIM(LOWER(TAG(I)))

*            IF EMPTY(strokakey) = .T.

*              EXIT

*            ELSE

*              DO CASE
*                CASE CANDIDATE(I) = .T.
*                  stroka_index = 'INDEX ON ' + (strokakey) + ' TAG ' + (strokatag) + ' CANDIDATE'

*                CASE UNIQUE(I) = .T.
*                  stroka_index = 'INDEX ON ' + (strokakey) + ' TAG ' + (strokatag) + ' UNIQUE'

*                CASE PRIMARY(I) = .T.
*                  stroka_index = 'INDEX ON ' + (strokakey) + ' TAG ' + (strokatag) + ' UNIQUE'

*                OTHERWISE
*                  stroka_index = 'INDEX ON ' + (strokakey) + ' TAG ' + (strokatag)

*              ENDCASE

*              SELECT slovar_visa

*              IF I = 1
*                REPLACE cdxkey WITH (stroka_index)
*              ELSE
*                REPLACE cdxkey WITH ALLTRIM(cdxkey) + CHR(13) + (stroka_index)
*              ENDIF

*            ENDIF

*            SELECT (sel)

*            I = I + 1

*          ENDDO

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        tiplog = .T.

      ELSE
        fail = ALLTRIM(failcdx)
        tiplog = .F.
        RETURN tiplog
      ENDIF

    ELSE
      fail = ALLTRIM(faildbf)
      tiplog = .F.
      RETURN tiplog
    ENDIF

    SELECT (aliasscan)
  ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  tiplog = .T.

ELSE
  fail = ALLTRIM(faildbf)
  tiplog = .F.
  RETURN tiplog
ENDIF

WAIT CLEAR

RETURN


************************************************************** FUNCTION popup_9 ***********************************************************************

FUNCTION popup_9
PARAMETERS m.tex1,m.tex2,m.tex3,m.tex4,m.bar
DO CASE
  CASE m.bar = 3
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1 OF vibor PROMPT m.tex1
    DEFINE BAR 2 OF vibor PROMPT '\-'
    DEFINE BAR 3 OF vibor PROMPT m.tex2
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

  CASE m.bar = 5
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1 OF vibor PROMPT m.tex1
    DEFINE BAR 2 OF vibor PROMPT '\-'
    DEFINE BAR 3 OF vibor PROMPT m.tex2
    DEFINE BAR 4 OF vibor PROMPT '\-'
    DEFINE BAR 5 OF vibor PROMPT m.tex3
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

  CASE m.bar = 7
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1 OF vibor PROMPT m.tex1
    DEFINE BAR 2 OF vibor PROMPT '\-'
    DEFINE BAR 3 OF vibor PROMPT m.tex2
    DEFINE BAR 4 OF vibor PROMPT '\-'
    DEFINE BAR 5 OF vibor PROMPT m.tex3
    DEFINE BAR 6 OF vibor PROMPT '\-'
    DEFINE BAR 7 OF vibor PROMPT m.tex4
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

ENDCASE
RETURN


********************************************************** FUNCTION popup_big **************************************************************************

FUNCTION popup_big
PARAMETERS m.tex1,m.tex2,m.tex3,m.tex4,m.tex5,m.tex6,m.tex7,m.tex8,m.bar
DO CASE
  CASE m.bar = 9
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4,m.tex5))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1 OF vibor PROMPT m.tex1
    DEFINE BAR 2 OF vibor PROMPT '\-'
    DEFINE BAR 3 OF vibor PROMPT m.tex2
    DEFINE BAR 4 OF vibor PROMPT '\-'
    DEFINE BAR 5 OF vibor PROMPT m.tex3
    DEFINE BAR 6 OF vibor PROMPT '\-'
    DEFINE BAR 7 OF vibor PROMPT m.tex4
    DEFINE BAR 8 OF vibor PROMPT '\-'
    DEFINE BAR 9 OF vibor PROMPT m.tex5
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

  CASE m.bar = 11
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4,m.tex5,m.tex6))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1   OF vibor PROMPT m.tex1
    DEFINE BAR 2   OF vibor PROMPT '\-'
    DEFINE BAR 3   OF vibor PROMPT m.tex2
    DEFINE BAR 4   OF vibor PROMPT '\-'
    DEFINE BAR 5   OF vibor PROMPT m.tex3
    DEFINE BAR 6   OF vibor PROMPT '\-'
    DEFINE BAR 7   OF vibor PROMPT m.tex4
    DEFINE BAR 8   OF vibor PROMPT '\-'
    DEFINE BAR 9   OF vibor PROMPT m.tex5
    DEFINE BAR 10 OF vibor PROMPT '\-'
    DEFINE BAR 11 OF vibor PROMPT m.tex6
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

  CASE m.bar = 13
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4,m.tex5,m.tex6,m.tex7))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1   OF vibor PROMPT m.tex1
    DEFINE BAR 2   OF vibor PROMPT '\-'
    DEFINE BAR 3   OF vibor PROMPT m.tex2
    DEFINE BAR 4   OF vibor PROMPT '\-'
    DEFINE BAR 5   OF vibor PROMPT m.tex3
    DEFINE BAR 6   OF vibor PROMPT '\-'
    DEFINE BAR 7   OF vibor PROMPT m.tex4
    DEFINE BAR 8   OF vibor PROMPT '\-'
    DEFINE BAR 9   OF vibor PROMPT m.tex5
    DEFINE BAR 10 OF vibor PROMPT '\-'
    DEFINE BAR 11 OF vibor PROMPT m.tex6
    DEFINE BAR 12 OF vibor PROMPT '\-'
    DEFINE BAR 13 OF vibor PROMPT m.tex7
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

  CASE m.bar = 15
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4,m.tex5,m.tex6,m.tex7,m.tex8))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1   OF vibor PROMPT m.tex1
    DEFINE BAR 2   OF vibor PROMPT '\-'
    DEFINE BAR 3   OF vibor PROMPT m.tex2
    DEFINE BAR 4   OF vibor PROMPT '\-'
    DEFINE BAR 5   OF vibor PROMPT m.tex3
    DEFINE BAR 6   OF vibor PROMPT '\-'
    DEFINE BAR 7   OF vibor PROMPT m.tex4
    DEFINE BAR 8   OF vibor PROMPT '\-'
    DEFINE BAR 9   OF vibor PROMPT m.tex5
    DEFINE BAR 10 OF vibor PROMPT '\-'
    DEFINE BAR 11 OF vibor PROMPT m.tex6
    DEFINE BAR 12 OF vibor PROMPT '\-'
    DEFINE BAR 13 OF vibor PROMPT m.tex7
    DEFINE BAR 14 OF vibor PROMPT '\-'
    DEFINE BAR 15 OF vibor PROMPT m.tex8
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

ENDCASE
RETURN


******************************************************* FUNCTION popup_big_big **************************************************************************

FUNCTION popup_big_big
PARAMETERS m.tex1,m.tex2,m.tex3,m.tex4,m.tex5,m.tex6,m.tex7,m.tex8,m.tex9,m.tex10,m.tex11,m.tex12,m.bar
DO CASE
  CASE m.bar = 9
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4,m.tex5))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1 OF vibor PROMPT m.tex1
    DEFINE BAR 2 OF vibor PROMPT '\-'
    DEFINE BAR 3 OF vibor PROMPT m.tex2
    DEFINE BAR 4 OF vibor PROMPT '\-'
    DEFINE BAR 5 OF vibor PROMPT m.tex3
    DEFINE BAR 6 OF vibor PROMPT '\-'
    DEFINE BAR 7 OF vibor PROMPT m.tex4
    DEFINE BAR 8 OF vibor PROMPT '\-'
    DEFINE BAR 9 OF vibor PROMPT m.tex5
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

  CASE m.bar = 11
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4,m.tex5,m.tex6))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1   OF vibor PROMPT m.tex1
    DEFINE BAR 2   OF vibor PROMPT '\-'
    DEFINE BAR 3   OF vibor PROMPT m.tex2
    DEFINE BAR 4   OF vibor PROMPT '\-'
    DEFINE BAR 5   OF vibor PROMPT m.tex3
    DEFINE BAR 6   OF vibor PROMPT '\-'
    DEFINE BAR 7   OF vibor PROMPT m.tex4
    DEFINE BAR 8   OF vibor PROMPT '\-'
    DEFINE BAR 9   OF vibor PROMPT m.tex5
    DEFINE BAR 10 OF vibor PROMPT '\-'
    DEFINE BAR 11 OF vibor PROMPT m.tex6
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

  CASE m.bar = 13
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4,m.tex5,m.tex6,m.tex7))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1   OF vibor PROMPT m.tex1
    DEFINE BAR 2   OF vibor PROMPT '\-'
    DEFINE BAR 3   OF vibor PROMPT m.tex2
    DEFINE BAR 4   OF vibor PROMPT '\-'
    DEFINE BAR 5   OF vibor PROMPT m.tex3
    DEFINE BAR 6   OF vibor PROMPT '\-'
    DEFINE BAR 7   OF vibor PROMPT m.tex4
    DEFINE BAR 8   OF vibor PROMPT '\-'
    DEFINE BAR 9   OF vibor PROMPT m.tex5
    DEFINE BAR 10 OF vibor PROMPT '\-'
    DEFINE BAR 11 OF vibor PROMPT m.tex6
    DEFINE BAR 12 OF vibor PROMPT '\-'
    DEFINE BAR 13 OF vibor PROMPT m.tex7
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

  CASE m.bar = 15
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4,m.tex5,m.tex6,m.tex7,m.tex8))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1   OF vibor PROMPT m.tex1
    DEFINE BAR 2   OF vibor PROMPT '\-'
    DEFINE BAR 3   OF vibor PROMPT m.tex2
    DEFINE BAR 4   OF vibor PROMPT '\-'
    DEFINE BAR 5   OF vibor PROMPT m.tex3
    DEFINE BAR 6   OF vibor PROMPT '\-'
    DEFINE BAR 7   OF vibor PROMPT m.tex4
    DEFINE BAR 8   OF vibor PROMPT '\-'
    DEFINE BAR 9   OF vibor PROMPT m.tex5
    DEFINE BAR 10 OF vibor PROMPT '\-'
    DEFINE BAR 11 OF vibor PROMPT m.tex6
    DEFINE BAR 12 OF vibor PROMPT '\-'
    DEFINE BAR 13 OF vibor PROMPT m.tex7
    DEFINE BAR 14 OF vibor PROMPT '\-'
    DEFINE BAR 15 OF vibor PROMPT m.tex8
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

  CASE m.bar = 17
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4,m.tex5,m.tex6,m.tex7,m.tex8,m.tex9))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1   OF vibor PROMPT m.tex1
    DEFINE BAR 2   OF vibor PROMPT '\-'
    DEFINE BAR 3   OF vibor PROMPT m.tex2
    DEFINE BAR 4   OF vibor PROMPT '\-'
    DEFINE BAR 5   OF vibor PROMPT m.tex3
    DEFINE BAR 6   OF vibor PROMPT '\-'
    DEFINE BAR 7   OF vibor PROMPT m.tex4
    DEFINE BAR 8   OF vibor PROMPT '\-'
    DEFINE BAR 9   OF vibor PROMPT m.tex5
    DEFINE BAR 10 OF vibor PROMPT '\-'
    DEFINE BAR 11 OF vibor PROMPT m.tex6
    DEFINE BAR 12 OF vibor PROMPT '\-'
    DEFINE BAR 13 OF vibor PROMPT m.tex7
    DEFINE BAR 14 OF vibor PROMPT '\-'
    DEFINE BAR 15 OF vibor PROMPT m.tex8
    DEFINE BAR 16 OF vibor PROMPT '\-'
    DEFINE BAR 17 OF vibor PROMPT m.tex9
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

  CASE m.bar = 19
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4,m.tex5,m.tex6,m.tex7,m.tex8,m.tex9,m.tex10))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1   OF vibor PROMPT m.tex1
    DEFINE BAR 2   OF vibor PROMPT '\-'
    DEFINE BAR 3   OF vibor PROMPT m.tex2
    DEFINE BAR 4   OF vibor PROMPT '\-'
    DEFINE BAR 5   OF vibor PROMPT m.tex3
    DEFINE BAR 6   OF vibor PROMPT '\-'
    DEFINE BAR 7   OF vibor PROMPT m.tex4
    DEFINE BAR 8   OF vibor PROMPT '\-'
    DEFINE BAR 9   OF vibor PROMPT m.tex5
    DEFINE BAR 10 OF vibor PROMPT '\-'
    DEFINE BAR 11 OF vibor PROMPT m.tex6
    DEFINE BAR 12 OF vibor PROMPT '\-'
    DEFINE BAR 13 OF vibor PROMPT m.tex7
    DEFINE BAR 14 OF vibor PROMPT '\-'
    DEFINE BAR 15 OF vibor PROMPT m.tex8
    DEFINE BAR 16 OF vibor PROMPT '\-'
    DEFINE BAR 17 OF vibor PROMPT m.tex9
    DEFINE BAR 18 OF vibor PROMPT '\-'
    DEFINE BAR 19 OF vibor PROMPT m.tex10
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

  CASE m.bar = 21
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4,m.tex5,m.tex6,m.tex7,m.tex8,m.tex9,m.tex10,m.tex11))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1   OF vibor PROMPT m.tex1
    DEFINE BAR 2   OF vibor PROMPT '\-'
    DEFINE BAR 3   OF vibor PROMPT m.tex2
    DEFINE BAR 4   OF vibor PROMPT '\-'
    DEFINE BAR 5   OF vibor PROMPT m.tex3
    DEFINE BAR 6   OF vibor PROMPT '\-'
    DEFINE BAR 7   OF vibor PROMPT m.tex4
    DEFINE BAR 8   OF vibor PROMPT '\-'
    DEFINE BAR 9   OF vibor PROMPT m.tex5
    DEFINE BAR 10 OF vibor PROMPT '\-'
    DEFINE BAR 11 OF vibor PROMPT m.tex6
    DEFINE BAR 12 OF vibor PROMPT '\-'
    DEFINE BAR 13 OF vibor PROMPT m.tex7
    DEFINE BAR 14 OF vibor PROMPT '\-'
    DEFINE BAR 15 OF vibor PROMPT m.tex8
    DEFINE BAR 16 OF vibor PROMPT '\-'
    DEFINE BAR 17 OF vibor PROMPT m.tex9
    DEFINE BAR 18 OF vibor PROMPT '\-'
    DEFINE BAR 19 OF vibor PROMPT m.tex10
    DEFINE BAR 20 OF vibor PROMPT '\-'
    DEFINE BAR 21 OF vibor PROMPT m.tex11
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

  CASE m.bar = 23
    rownum = (WROWS()-m.bar)/2 - 2
    colnum = (WCOLS()-LEN(MAX(m.tex1,m.tex2,m.tex3,m.tex4,m.tex5,m.tex6,m.tex7,m.tex8,m.tex9,m.tex10,m.tex11,m.tex12))) / 3

    DEFINE POPUP vibor FROM rownum,colnum  MARGIN SHADOW  ;
      FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
      COLOR SCHEME 2 MESSAGE text_mess
    DEFINE BAR 1   OF vibor PROMPT m.tex1
    DEFINE BAR 2   OF vibor PROMPT '\-'
    DEFINE BAR 3   OF vibor PROMPT m.tex2
    DEFINE BAR 4   OF vibor PROMPT '\-'
    DEFINE BAR 5   OF vibor PROMPT m.tex3
    DEFINE BAR 6   OF vibor PROMPT '\-'
    DEFINE BAR 7   OF vibor PROMPT m.tex4
    DEFINE BAR 8   OF vibor PROMPT '\-'
    DEFINE BAR 9   OF vibor PROMPT m.tex5
    DEFINE BAR 10 OF vibor PROMPT '\-'
    DEFINE BAR 11 OF vibor PROMPT m.tex6
    DEFINE BAR 12 OF vibor PROMPT '\-'
    DEFINE BAR 13 OF vibor PROMPT m.tex7
    DEFINE BAR 14 OF vibor PROMPT '\-'
    DEFINE BAR 15 OF vibor PROMPT m.tex8
    DEFINE BAR 16 OF vibor PROMPT '\-'
    DEFINE BAR 17 OF vibor PROMPT m.tex9
    DEFINE BAR 18 OF vibor PROMPT '\-'
    DEFINE BAR 19 OF vibor PROMPT m.tex10
    DEFINE BAR 20 OF vibor PROMPT '\-'
    DEFINE BAR 21 OF vibor PROMPT m.tex11
    DEFINE BAR 22 OF vibor PROMPT '\-'
    DEFINE BAR 23 OF vibor PROMPT m.tex12
    ON SELECTION POPUP vibor DEACTIVATE POPUP vibor

ENDCASE
RETURN


*************************************************************** FUNCTION err ***************************************************************************

FUNCTION err
PARAMETERS xyz
DO FORM (put_scr) + ('infoerr.scx')
RETURN


************************************************************** FUNCTION soob **************************************************************************

FUNCTION soob
PARAMETERS xyz
DO FORM (put_scr) + ('infosoob.scx')
RETURN


********************************************************** FUNCTION MinDataMes **********************************************************************

FUNCTION MinDataMes
PARAMETERS par_data

den = '01'

mes = PADL(ALLTRIM(STR(MONTH(par_data))), 2, '0')

god = ALLTRIM(STR(YEAR(par_data)))

colvo_den = CTOD(den + '.' + mes + '.' + god)

RETURN colvo_den


********************************************************** FUNCTION MaxDataMes **********************************************************************

FUNCTION MaxDataMes
PARAMETERS par_data

den = ALLTRIM(STR(GOMONTH(CTOD('01.' + PADL(ALLTRIM(STR(MONTH(par_data))), 2, '0') + '.' + ;
  ALLTRIM(STR(YEAR(par_data)))), 1) - CTOD('01.' + PADL(ALLTRIM(STR(MONTH(par_data))), 2, '0') + '.' + ;
  ALLTRIM(STR(YEAR(par_data))))))

mes = PADL(ALLTRIM(STR(MONTH(par_data))), 2, '0')

god = ALLTRIM(STR(YEAR(par_data)))

colvo_den = CTOD(den + '.' + mes + '.' + god)

RETURN colvo_den


************************************************************ FUNCTION rusmes **************************************************************************

FUNCTION RusMes
PARAMETERS par_mes
DO CASE
  CASE MONTH(par_mes) <> 0
    imames = ALLTRIM(Mes(par_mes))

  CASE MONTH(par_mes) = 0
    imames = ALLTRIM(Mes(DATE()))

ENDCASE
RETURN imames


****************************************************************** FUNCTION Mes ***********************************************************************

FUNCTION Mes
PARAMETERS par_data
DIMENSION mesaz(12)
mesaz (1) = 'январь'
mesaz (2) = 'февраль'
mesaz (3) = 'март'
mesaz (4) = 'апрель'
mesaz (5) = 'май'
mesaz (6) = 'июнь'
mesaz (7) = 'июль'
mesaz (8) = 'август'
mesaz (9) = 'сентябрь'
mesaz (10) = 'октябрь'
mesaz (11) = 'ноябрь'
mesaz (12) = 'декабрь'
mes = mesaz(MONTH(par_data))
RETURN mes


****************************************************************** FUNCTION dt_kv **********************************************************************

FUNCTION dt_kv
PARAMETERS par_data

IF EMPTY(par_data) = .T.
  par_data = DATE()
ENDIF

DIMENSION mesaz(12)
mesaz (1) = 'январь'
mesaz (2) = 'февраль'
mesaz (3) = 'март'
mesaz (4) = 'апрель'
mesaz (5) = 'май'
mesaz (6) = 'июнь'
mesaz (7) = 'июль'
mesaz (8) = 'август'
mesaz (9) = 'сентябрь'
mesaz (10) = 'октябрь'
mesaz (11) = 'ноябрь'
mesaz (12) = 'декабрь'

IF mesaz(MONTH(par_data)) = 'март' .OR. mesaz(MONTH(par_data)) = 'август'
  mes1 = mesaz(MONTH(par_data)) + 'а'
ELSE
  mes1 = SUBSTR(mesaz(MONTH(par_data)),1,LEN(mesaz(MONTH(par_data)))-1) + 'я'
ENDIF

mes = ALLTRIM('"' + PADL(ALLTRIM(STR(DAY(par_data))),2,'0') + '"' + ' ' + mes1 + ' ' + ALLTRIM(STR(YEAR(par_data))) + ' ' + 'г.')

RETURN mes


**************************************************************** FUNCTION dt_padeg *********************************************************************

FUNCTION dt_padeg
PARAMETERS par_data

IF EMPTY(par_data) = .F.

  DIMENSION mesaz(12)
  mesaz (1) = 'январь'
  mesaz (2) = 'февраль'
  mesaz (3) = 'март'
  mesaz (4) = 'апрель'
  mesaz (5) = 'май'
  mesaz (6) = 'июнь'
  mesaz (7) = 'июль'
  mesaz (8) = 'август'
  mesaz (9) = 'сентябрь'
  mesaz (10) = 'октябрь'
  mesaz (11) = 'ноябрь'
  mesaz (12) = 'декабрь'

  IF mesaz(MONTH(par_data)) = 'март' .OR. mesaz(MONTH(par_data)) = 'август'
    mes1 = mesaz(MONTH(par_data)) + 'а'
  ELSE
    mes1 = SUBSTR(mesaz(MONTH(par_data)),1,LEN(mesaz(MONTH(par_data)))-1) + 'я'
  ENDIF

  mes = ALLTRIM('до ' + mes1 + ' ' + ALLTRIM(STR(YEAR(par_data))) + ' ' + 'г.')

ELSE

  mes = SPACE(0)

ENDIF

RETURN mes


****************************************************************** FUNCTION dt *************************************************************************

FUNCTION dt
PARAMETERS par_data

IF EMPTY(par_data) = .F.

  DIMENSION mesaz(12)
  mesaz (1) = 'январь'
  mesaz (2) = 'февраль'
  mesaz (3) = 'март'
  mesaz (4) = 'апрель'
  mesaz (5) = 'май'
  mesaz (6) = 'июнь'
  mesaz (7) = 'июль'
  mesaz (8) = 'август'
  mesaz (9) = 'сентябрь'
  mesaz (10) = 'октябрь'
  mesaz (11) = 'ноябрь'
  mesaz (12) = 'декабрь'

  IF mesaz(MONTH(par_data)) = 'март' .OR. mesaz(MONTH(par_data)) = 'август'
    mes1 = mesaz(MONTH(par_data)) + 'а'
  ELSE
    mes1 = SUBSTR(mesaz(MONTH(par_data)),1,LEN(mesaz(MONTH(par_data)))-1) + 'я'
  ENDIF

  mes = ALLTRIM(PADL(ALLTRIM(STR(DAY(par_data))),2,'0') + ' ' + mes1 + ' ' + ALLTRIM(STR(YEAR(par_data))) + ' ' + 'г.')

ELSE

  mes = SPACE(0)

ENDIF

RETURN mes


*************************************************************** FUNCTION inizial_fio *********************************************************************

FUNCTION inizial_fio
PARAMETERS par_fio
fio = SUBSTR(ALLTRIM(par_fio),1,AT(SPACE(1),par_fio,1) - 1) + SPACE(1) + SUBSTR(ALLTRIM(par_fio),AT(SPACE(1),par_fio,1) + 1,1) + '.' + SPACE(1) + SUBSTR(ALLTRIM(par_fio),AT(SPACE(1),par_fio,2) + 1,1) + '.'
RETURN fio


*************************************************************** FUNCTION inizial_iof *********************************************************************

FUNCTION inizial_iof
PARAMETERS par_fio
fio = SUBSTR(ALLTRIM(par_fio),AT(SPACE(1),par_fio,1) + 1,1) + '.' + SPACE(1) + SUBSTR(ALLTRIM(par_fio),AT(SPACE(1),par_fio,2) + 1,1) + '.' + SPACE(1) + SUBSTR(ALLTRIM(par_fio),1,AT(SPACE(1),par_fio,1) - 1)
RETURN fio


************************************************************* PROCEDURE vxod_parol ********************************************************************

PROCEDURE vxod_parol

STORE SPACE(0) TO personfio

=soob('Внимание! Необходимо зарегистрироваться.')

SELECT * ;
  FROM parol ;
  INTO CURSOR prl ;
  WHERE EMPTY(fio) = .F. ;
  ORDER BY fio

colvo_zap = _TALLY

IF colvo_zap <> 0

  coln = IIF(colvo_zap <= 25, colvo_zap, 25)
  v_rov = INT((WROWS()-coln)/3) + 2
  n_rov = v_rov + coln + 3
  l_col = INT((WCOLS()-LEN(prl.fio) + 2)/2) - 5
  p_col = l_col + LEN(prl.fio) + 30

  DEFINE POPUP parol ;
    FROM v_rov, l_col TO n_rov, p_col ;
    PROMPT FIELD prl.fio ;
    FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
    SCROLL MARGIN SHADOW  COLOR SCHEME 2 ;
    MESSAGE ' Для pаботы в пpогpамме нужно заpегистpиpоваться, выбpав свою фамилию из списка'
  ON SELECTION POPUP parol DEACTIVATE POPUP parol

  ACTIVATE POPUP parol

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    personfio = ALLTRIM(PROMPT())

    SELECT parol
    SET ORDER TO fio

    IF SEEK(personfio)

* WAIT WINDOW 'parol.fio = ' + ALLTRIM(parol.fio) + '   personfio = ' + ALLTRIM(personfio)

* BROWSE WINDOW brows

      DO WHILE .T.

        STORE SPACE(0) TO personima

        DO FORM (put_scr) + ('sysima.scx')

        personima = UPPER(ALLTRIM(personima))

        IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

          SELECT parol
          SET ORDER TO ima
          SET FILTER TO LOWER(ALLTRIM(parol.fio)) == LOWER(ALLTRIM(personfio))
          GO TOP

          IF SEEK(personima)

* WAIT WINDOW 'parol.ima = ' + ALLTRIM(parol.ima) + '   personima = ' + ALLTRIM(personima)

* BROWSE WINDOW brows

            REPLACE use_soft WITH .T., colvo_soft WITH colvo_soft + 1

            IF EMPTY(parol.parol) = .T.

              DO WHILE .T.

                =soob('В справочнике паролей у данного пользователя пароль не введен! Его длина не менее 5 знаков.')

                STORE SPACE(0) TO personparol

                DO FORM (put_scr) + ('parol.scx')

                personparol = UPPER(ALLTRIM(personparol))

                IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

                  IF LEN(ALLTRIM(personparol)) < 5
                    =err('Длина пароля должна быть не менеее 5 знаков.')
                    LOOP
                  ELSE
                    personparol = UPPER(ALLTRIM(personparol))
                    REPLACE NEXT 1 parol.parol WITH personparol
                    EXIT
                  ENDIF

                ELSE
                  EXIT
                ENDIF

              ENDDO
            ENDIF

            IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

              FOR I = 1 TO 2

                STORE SPACE(0) TO personparol

                DO FORM (put_scr) + ('parol.scx')

                personparol = UPPER(ALLTRIM(personparol))

                IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

                  personparol = UPPER(ALLTRIM(personparol))

                  SELECT parol

* WAIT WINDOW 'UPPER(ALLTRIM(parol.parol)) = ' + UPPER(ALLTRIM(parol.parol)) + '   ALLTRIM(personparol) = ' + ALLTRIM(personparol)

* BROWSE WINDOW brows

                  IF UPPER(ALLTRIM(parol.parol)) == ALLTRIM(personparol)

                    REPLACE NEXT 1 parol.parol WITH personparol

                    wes = parol.pravo
                    adm_kod = ALLTRIM(parol.kod)
                    par_kod = ALLTRIM(parol.kod)
                    fio_user = ALLTRIM(parol.fio)
                    par_fio = ALLTRIM(parol.fio)
                    par_ima = ALLTRIM(parol.ima)
                    par_par = ALLTRIM(parol.parol)
                    par_prim = ALLTRIM(parol.prim)
                    par_prompt = parol.pr_prompt
                    par_grup = ALLTRIM(parol.grup)
                    kod_kassa = ALLTRIM(parol.kassa)

                    SELECT parol
                    SET FILTER TO

                    DO vxod IN Visa

                    EXIT

                  ELSE

                    =err('Внимание! Введенный Вами пароль не верен ... ')

                    IF I = 1
                      =err('У Вас есть еще одна попытка, повторите ввод пароля.')
                    ELSE
                      =err('Ваша попытка неудачна, обратитесь к администратору БД.')
                    ENDIF

                  ENDIF
                  LOOP

                ELSE
                  EXIT
                ENDIF

              ENDFOR

            ELSE
              EXIT
            ENDIF

          ELSE
            =err('Внимание! Введенное Вами имя в системе не зарегистрировано ... ')
            LOOP
          ENDIF

        ENDIF
        EXIT
      ENDDO

    ELSE
      =err('Внимание! Поиск прошел неудачно, фамилия не найдена')
    ENDIF

  ENDIF
ENDIF

RELEASE POPUP parol

RETURN


***************************************************************** PROCEDURE exit ***********************************************************************

PROCEDURE exit

IF pr_arxiv = .T.
  DO arxiv IN Visa
ENDIF

DO svert IN Visa

=LoadKeyboard('ENG')  && Включаем английскую раскладку клавиатуры

IF sergeizhuk_flag = .T.
  =CAPSLOCK(.T.)
ELSE
  =CAPSLOCK(.F.)
ENDIF

POP KEY ALL
SET HELP ON
SET TALK ON
SET NOTIFY ON
SET WINDOW OF MEMO TO
SET STATUS BAR ON
SET SAFETY ON
SET CURSOR ON
SET EXCLUSIVE ON
SET SYSFORMATS ON && настройки системы включены

IF sergeizhuk_flag = .T.
  SET PATH TO d:\bank_rep\
ELSE
  SET PATH TO
ENDIF

IF pr_dostup = .T.  && Признак разрешения на вход в ПО

  SELECT parol
  SET ORDER TO ima

  IF SEEK(ALLTRIM(par_ima))

    DO CASE
      CASE parol.colvo_soft = 1
        REPLACE use_soft WITH .F., colvo_soft WITH colvo_soft - 1

      CASE parol.colvo_soft > 1
        REPLACE colvo_soft WITH colvo_soft - 1

      CASE parol.colvo_soft < 0
        REPLACE use_soft WITH .F., colvo_soft WITH 0

    ENDCASE

  ENDIF

ENDIF

ON ERROR

CLOSE DATA ALL

RETURN


******************************************************************* PROCEDURE arxiv ********************************************************************

PROCEDURE arxiv

faildbfima = 'slovar_visa'
aliasscan = 'slovar_visa'

faildbf = (put_dbf) + (faildbfima) + ('.dbf')
faildbfarx = (put_arx) + (faildbfima) + ('.dbf')

IF FILE(faildbf)

  SELECT (aliasscan)
  COPY TO (faildbfarx)

  SCAN FOR pr_copy = .T.

    faildbfima = ALLTRIM(slovar_visa.dbfname)
    aliasima = ALLTRIM(slovar_visa.alias)
    failcdxima = ALLTRIM(slovar_visa.cdxname)

    faildbf = (put_dbf) + (faildbfima) + ('.dbf')
    failcdx = (put_dbf) + (faildbfima) + ('.cdx')
    failarx = (put_arx) + (faildbfima) + ('.dbf')

    IF FILE(faildbf)
      IF FILE(failcdx)

        IF NOT USED(aliasima)
          SELECT 0
          USE (faildbfima) ALIAS (aliasima)
        ELSE
          SELECT (aliasima)
        ENDIF

        SELECT (aliasima)
        GO TOP

        tim1 = DATETIME()

        WAIT WINDOW ' Начато копирование таблицы - ' + (faildbfima) + '.dbf' + ' и индека - ' + (failcdxima) NOWAIT

        COPY TO (failarx) CDX

        tim2 = DATETIME()

        IF (tim2 - tim1) <> 0
          WAIT WINDOW ' Копирование таблицы - ' + (faildbfima) + '.dbf' + ' и индека - ' + (failcdxima) + ' завершено. (время = ' + ALLTRIM(STR(tim2 - tim1, 8, 4)) + ' сек.)' TIMEOUT 2
        ENDIF

        WAIT CLEAR

        tiplog = .T.

      ELSE
        fail = ALLTRIM(failcdx)
        tiplog = .F.
        RETURN tiplog
      ENDIF

    ELSE
      fail = ALLTRIM(faildbf)
      tiplog = .F.
      RETURN tiplog
    ENDIF

    SELECT (aliasscan)
  ENDSCAN

  WAIT CLEAR

  SELECT (aliasscan)
  tiplog = .T.

ELSE
  fail = ALLTRIM(faildbf)
  tiplog = .F.
  RETURN tiplog
ENDIF

RETURN


******************************************************************** FUNCTION new_branch **************************************************************

FUNCTION new_branch
PARAMETERS num_kod

STORE SPACE(3) TO new_branch

DO CASE
  CASE ALLTRIM(num_kod) == '01'
    new_branch='013'
  CASE ALLTRIM(num_kod) == '02'
    new_branch='002'
  CASE ALLTRIM(num_kod) == '03'
    new_branch='001'
  CASE ALLTRIM(num_kod) == '04'
    new_branch='005'
  CASE ALLTRIM(num_kod) == '05'
    new_branch='006'
  CASE ALLTRIM(num_kod) == '06'
    new_branch='004'
  CASE ALLTRIM(num_kod) == '07'
    new_branch='007'
  CASE ALLTRIM(num_kod) == '08'
    new_branch='010'
  CASE ALLTRIM(num_kod) == '09'
    new_branch='003'
  CASE ALLTRIM(num_kod) == '10'
    new_branch='008'
  CASE ALLTRIM(num_kod) == '11'
    new_branch='009'
  CASE ALLTRIM(num_kod) == '12'
    new_branch='011'

ENDCASE

RETURN new_branch


************************************************************ PROCEDURE ima_arh_upk *****************************************************************

PROCEDURE ima_arh_upk

*   Название архивных файлов, формируемых филиалом для УПК :

*   gddmmnnv.zip  - Головной банк для УПК
*   lddmmnnv.zip  - Ленск для УПК
*   fddmmnnv.zip  - Москва для УПК
*   uddmmnnv.zip  - Удачный для УПК
*   yddmmnnv.zip  - Якутск для УПК
*   mddmmnnv.zip  - Мирный для УПК
*   addmmnnv.zip  - Айхал для УПК
*   nddmmnnv.zip  - Новосибирск для УПК
*   oddmmnnv.zip  - Орел для УПК
*   hddmmnnv.zip  - Архангельск для УПК
*   kddmmnnv.zip  - Краснодар для УПК
*   iddmmnnv.zip  - Иркутск для УПК

DO CASE
  CASE ALLTRIM(num_branch) == '01'  && gddmmnnv.zip  - Головной банк для УПК
    ima_arh=('g') + (ddmm) + (num_reis) + 'v' + ('.zip')

  CASE ALLTRIM(num_branch) == '02'  && lddmmnnv.zip  - Ленск для УПК
    ima_arh=('l') + (ddmm) + (num_reis) + 'v' + ('.zip')

  CASE ALLTRIM(num_branch) == '03'  && fddmmnnv.zip  - Москва для УПК
    ima_arh=('f') + (ddmm) + (num_reis) + 'v' + ('.zip')

  CASE ALLTRIM(num_branch) == '04'  && uddmmnnv.zip  - Удачный для УПК
    ima_arh=('u') + (ddmm) + (num_reis) + 'v' + ('.zip')

  CASE ALLTRIM(num_branch) == '05'  && yddmmnnv.zip  - Якутск для УПК
    ima_arh=('y') + (ddmm) + (num_reis) + 'v' + ('.zip')

  CASE ALLTRIM(num_branch) == '06'  && mddmmnnv.zip  - Мирный для УПК
    ima_arh=('m') + (ddmm) + (num_reis) + 'v' + ('.zip')

  CASE ALLTRIM(num_branch) == '07'  && addmmnnv.zip  - Айхал для УПК
    ima_arh=('a') + (ddmm) + (num_reis) + 'v' + ('.zip')

  CASE ALLTRIM(num_branch) == '08'  && nddmmnnv.zip  - Новосибирск для УПК
    ima_arh=('n') + (ddmm) + (num_reis) + 'v' + ('.zip')

  CASE ALLTRIM(num_branch) == '09'  && oddmmnnv.zip  - Орел для УПК
    ima_arh=('o') + (ddmm) + (num_reis) + 'v' + ('.zip')

  CASE ALLTRIM(num_branch) == '10'  && hddmmnnv.zip  - Архангельск для УПК
    ima_arh=('h') + (ddmm) + (num_reis) + 'v' + ('.zip')

  CASE ALLTRIM(num_branch) == '11'  && kddmmnnv.zip  - Краснодар для УПК
    ima_arh=('k') + (ddmm) + (num_reis) + 'v' + ('.zip')

  CASE ALLTRIM(num_branch) == '12'  && iddmmnnv.zip  - Иркутск для УПК
    ima_arh=('i') + (ddmm) + (num_reis) + 'v' + ('.zip')

ENDCASE

RETURN


************************************************************ PROCEDURE ima_arh_zhuk *****************************************************************

PROCEDURE ima_arh_zhuk

*   Название архивных файлов, формируемых филиалом для Управления Автоматизации, для Жука С.Б. :

*   gddmmnnz.zip  - Головной банк для УА
*   lddmmnnz.zip  - Ленск для УА
*   fddmmnnz.zip  - Москва для УА
*   uddmmnnz.zip  - Удачный для УА
*   yddmmnnz.zip  - Якутск для УА
*   mddmmnnz.zip  - Мирный для УА
*   addmmnnz.zip  - Айхал для УА
*   nddmmnnz.zip  - Новосибирск для УА
*   oddmmnnz.zip  - Орел для УА
*   hddmmnnz.zip  - Архангельск для УА
*   kddmmnnz.zip  - Краснодар для УА
*   iddmmnnz.zip  - Иркутск для УА

DO CASE
  CASE ALLTRIM(num_branch) == '01'  && gddmmnnz.zip  - Головной банк для УА
    ima_arh=('g') + (ddmm) + (num_reis) + 'z' + ('.zip')

  CASE ALLTRIM(num_branch) == '02'  && lddmmnnz.zip  - Ленск для УА
    ima_arh=('l') + (ddmm) + (num_reis) + 'z' + ('.zip')

  CASE ALLTRIM(num_branch) == '03'  && fddmmnnz.zip  - Москва для УА
    ima_arh=('f') + (ddmm) + (num_reis) + 'z' + ('.zip')

  CASE ALLTRIM(num_branch) == '04'  && uddmmnnz.zip  - Удачный для УА
    ima_arh=('u') + (ddmm) + (num_reis) + 'z' + ('.zip')

  CASE ALLTRIM(num_branch) == '05'  && yddmmnnz.zip  - Якутск для УА
    ima_arh=('y') + (ddmm) + (num_reis) + 'z' + ('.zip')

  CASE ALLTRIM(num_branch) == '06'  && mddmmnnz.zip  - Мирный для УА
    ima_arh=('m') + (ddmm) + (num_reis) + 'z' + ('.zip')

  CASE ALLTRIM(num_branch) == '07'  && addmmnnz.zip  - Айхал для УА
    ima_arh=('a') + (ddmm) + (num_reis) + 'z' + ('.zip')

  CASE ALLTRIM(num_branch) == '08'  && nddmmnnz.zip  - Новосибирск для УА
    ima_arh=('n') + (ddmm) + (num_reis) + 'z' + ('.zip')

  CASE ALLTRIM(num_branch) == '09'  && oddmmnnz.zip  - Орел для УА
    ima_arh=('o') + (ddmm) + (num_reis) + 'z' + ('.zip')

  CASE ALLTRIM(num_branch) == '10'  && hddmmnnz.zip  - Архангельск для УА
    ima_arh=('h') + (ddmm) + (num_reis) + 'z' + ('.zip')

  CASE ALLTRIM(num_branch) == '11'  && kddmmnnz.zip  - Краснодар для УА
    ima_arh=('k') + (ddmm) + (num_reis) + 'z' + ('.zip')

  CASE ALLTRIM(num_branch) == '12'  && iddmmnnz.zip  - Иркутск для УА
    ima_arh=('i') + (ddmm) + (num_reis) + 'z' + ('.zip')

ENDCASE

RETURN


*************************************************************** FUNCTION LoadKeyboard *****************************************************************

FUNCTION LoadKeyboard
PARAMETERS kod_klav

LOCAL lcLayoutID

DECLARE INTEGER LoadKeyboardLayout IN win32api STRING @, INTEGER

DO CASE
  CASE ALLTRIM(kod_klav) == 'ENG'

    =CAPSLOCK(.F.)
    lcLayoutID = '00000409'  && US English Layout

  CASE ALLTRIM(kod_klav) == 'RUS'

    =CAPSLOCK(.F.)
    lcLayoutID = '00000419'  && Russian Layout

ENDCASE

LoadKeyboardLayout(@lcLayoutID,1)

CLEAR DLLS

RETURN


************************************************************* PROCEDURE scan_name_bank ***************************************************************

PROCEDURE scan_name_bank

IF INLIST(ALLTRIM(num_branch), '04', '07')

  SELECT name_p, name_k ;
    FROM sprfirma ;
    WHERE ALLTRIM(bik_rur) + ALLTRIM(rkz_rur) + ALLTRIM(branch) == ALLTRIM(bik_bank_rur) + ALLTRIM(kor_bank_rur) + '06' ;
    INTO CURSOR sel_sprfirma ;
    ORDER BY 1

  name_bank = ALLTRIM(sel_sprfirma.name_p)

ELSE

  name_bank = ALLTRIM(name_bank_P)

ENDIF

RETURN name_bank


********************************************************** PROCEDURE brw_cln_dover *****************************************************************

PROCEDURE brw_cln_dover

SELECT DISTINCT ref_client, fio_rus, fio_rus_d, doveren, pass_seria, pass_ser_d, pass_num, pass_num_d, pass_data,;
  pass_dat_d, pass_mesto, pass_mes_d, dolg_rab, mesto_rab, adres_dom, telef_dom ;
  FROM client ;
  WHERE ALLTRIM(ref_client) == ALLTRIM(sel_account.ref_client) ;
  INTO CURSOR sel_client

IF _TALLY <> 0

  IF EMPTY(ALLTRIM(sel_client.doveren)) = .F.

    EDIT FIELDS ;
      fio_rus:H='ФИО доверителя',;
      fio_rus_d:H='ФИО доверяемого',;
      doveren:H='Доверенность',;
      pass_seria:H='Серия доверителя',;
      pass_ser_d:H='Серия доверяемого',;
      pass_num:H='Номер доверителя',;
      pass_num_d:H='Номер доверяемого',;
      pass_data:H='Дата доверителя',;
      pass_dat_d:H='Дата доверяемого',;
      pass_mesto:H='Место доверителя',;
      pass_mes_d:H='Место доверяемого',;
      adres_dom:H='Адрес доверителя',;
      telef_dom:H='Телефон доверителя' ;
      WINDOWS doverenost TITLE ' Справочные данные по клиенту - просмотр Memo поля по клавише Ctrl + Home; Выход по клавише Esc'

  ELSE
    =err('Внимание! Данные в поле доверенности не внесены, оно пустое.')
  ENDIF

ELSE
  =err('Внимание! Данных по клиенту в справочнике клиентов не обнаружено.')
ENDIF

RETURN


**************************************************************** FUNCTION vid_doc **********************************************************************

FUNCTION vid_doc
PARAMETERS seek_vid_doc

DO CASE
  CASE ALLTRIM(seek_vid_doc) == '001'
    text_vid_doc = 'Паспорт РФ'

  CASE ALLTRIM(seek_vid_doc) == '002'
    text_vid_doc = 'Паспорт РФ'

  CASE ALLTRIM(seek_vid_doc) == '003'
    text_vid_doc = 'Удостоверение личности'

  CASE ALLTRIM(seek_vid_doc) == '004'
    text_vid_doc = 'Военный билет'

  CASE ALLTRIM(seek_vid_doc) == '006'
    text_vid_doc = 'Иностранный паспорт'

  CASE ALLTRIM(seek_vid_doc) == '008'
    text_vid_doc = 'Разрешение на проживание'

  CASE ALLTRIM(seek_vid_doc) == '009'
    text_vid_doc = 'Вид на жительство'

  CASE ALLTRIM(seek_vid_doc) == '014'
    text_vid_doc = 'Иные выданные документы'

ENDCASE

RETURN (text_vid_doc)


*************************************************************** PROCEDURE vibor_proekt *****************************************************************

PROCEDURE vibor_proekt

IF pr_zarpl = .T.

  tex1 = 'Выборка данных отдельно по каждому проекту'
  tex2 = 'Выборка данных суммарно по всем проектам'
  l_bar = 3
  =popup_9(tex1,tex2,tex3,tex4,l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    DO CASE
      CASE BAR() = 1  && Выборка данных отдельно по каждому проекту
        pr_kod_proekt = .T.

      CASE BAR() = 3  && Выборка данных суммарно по всем проектам
        pr_kod_proekt = .F.

      OTHERWISE
        pr_kod_proekt = .F.
    ENDCASE

  ELSE
    pr_kod_proekt = .F.
  ENDIF

ELSE
  pr_kod_proekt = .F.
ENDIF

RETURN


*************************************************************** PROCEDURE vibor_zarpl_proekt ************************************************************

PROCEDURE vibor_zarpl_proekt

IF pr_zarpl = .T.

  DO WHILE .T.

    ACTIVATE POPUP zarpl_proekt

    IF NOT LASTKEY() = 27

      IF INLIST(LASTKEY(), 4, 19)
        LOOP
      ELSE
        DO zarpl_proekt IN Visa
        EXIT
      ENDIF

    ELSE

      num_proekt = '00'  && Номер зарплатного проекта

      =soob('Внимание! Ваш выбор - Главный справочник карточных счетов филиала - код проекта "00"')

      EXIT

    ENDIF

  ENDDO

ENDIF

RETURN


************************************************************ PROCEDURE vibor_zarpl_proekt_val ************************************************************

PROCEDURE vibor_zarpl_proekt_val

IF pr_zarpl = .T.

  DO CASE
    CASE INLIST(ALLTRIM(par_grup), 'RUB', 'ALL')

      DO WHILE .T.

        ACTIVATE POPUP zarpl_proekt

        IF NOT LASTKEY() = 27

          IF INLIST(LASTKEY(), 4, 19)
            LOOP
          ELSE
            DO zarpl_proekt IN Visa
            EXIT
          ENDIF

        ELSE

          num_proekt = '00'  && Номер зарплатного проекта

          =soob('Внимание! Ваш выбор - Главный справочник карточных счетов филиала - код проекта "00"')

          EXIT

        ENDIF

      ENDDO

    CASE INLIST(ALLTRIM(par_grup), 'VAL')

      num_proekt = '00'  && Номер зарплатного проекта

    OTHERWISE

      num_proekt = '00'  && Номер зарплатного проекта

  ENDCASE

ENDIF

RETURN


************************************************************** PROCEDURE select_osnov *******************************************************************

PROCEDURE select_osnov
PARAMETERS sel_num_liz

sel = SELECT()

STORE SPACE(0) TO select_osnov

SELECT branch, ima, text, numliz, osnovanie ;
  FROM spravliz ;
  WHERE ALLTRIM(numliz) == ALLTRIM(sel_num_liz) AND ALLTRIM(kodval) == SUBSTR(ALLTRIM(sel_num_liz), 6, 3) ;
  INTO CURSOR sel_spravliz ;
  ORDER BY numliz

* BROWSE WINDOW brows

IF _TALLY <> 0
  select_osnov = ALLTRIM(sel_spravliz.osnovanie)
ENDIF

SELECT(sel)
RETURN


************************************************ PROCEDURE formir_date_close ****************************************************************************

PROCEDURE formir_date_close

DO CASE
  CASE m.data_sost < CTOD('01.07.2008')  && C 01.07.2007 введены новые тарифы

    DO CASE
      CASE 'Visa Electron' $ name_card  && Карты Visa Electron выпускаются сроком на 2 года
        m.date_close = GOMONTH(m.data_sost, 24) - 1

      CASE NOT 'Visa Electron' $ name_card  && Остальные карты выпускаются сроком на 1 год
        m.date_close = GOMONTH(m.data_sost, 12) - 1

    ENDCASE

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE m.data_sost >= CTOD('01.07.2008')  && C 01.07.2008 введены новые тарифы, в которых все карты выпускаются сроком на 2 года

    m.date_close = GOMONTH(m.data_sost, 24) - 1

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE m.data_sost >= CTOD('01.06.2009')  && C 01.06.2009 введены новые тарифы, в которых все карты выпускаются сроком на 2 года

    m.date_close = GOMONTH(m.data_sost, 24) - 1

ENDCASE

RETURN


************************************************ PROCEDURE formir_close_card ****************************************************************************

PROCEDURE formir_close_card

DO CASE
  CASE m.date_card < CTOD('01.07.2008')  && C 01.07.2007 введены новые тарифы

    DO CASE
      CASE 'Visa Electron' $ name_card  && Карты Visa Electron выпускаются сроком на 2 года
        m.close_card = GOMONTH(m.date_card, 24) - 1

      CASE NOT 'Visa Electron' $ name_card  && Остальные карты выпускаются сроком на 1 год
        m.close_card = GOMONTH(m.date_card, 12) - 1

    ENDCASE

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE m.date_card >= CTOD('01.07.2008')  && C 01.07.2008 введены новые тарифы, в которых все карты выпускаются сроком на 2 года

    m.close_card = GOMONTH(m.date_card, 24) - 1

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  CASE m.date_card >= CTOD('01.06.2009')  && C 01.06.2009 введены новые тарифы, в которых все карты выпускаются сроком на 2 года

    m.close_card = GOMONTH(m.date_card, 24) - 1

ENDCASE

RETURN


****************************************************************** PROCEDURE select_prn_form_filial *******************************************************

PROCEDURE select_prn_form_filial

DO CASE

  CASE NOT INLIST(ALLTRIM(num_branch), '03')


  CASE  INLIST(ALLTRIM(num_branch), '03')


ENDCASE

RETURN


************************************************************** PROCEDURE prn_form ***********************************************************************

PROCEDURE prn_form

put_prn = 1

* WAIT WINDOW 'ima_frx = ' + ALLTRIM(ima_frx)

DO WHILE .T.

  tex1 = 'Печатный отчет на экран'
  tex2 = 'Печатный отчет на принтер'
  l_bar = 3
  =popup_9(tex1,tex2,tex3,tex4,l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

    put_prn = IIF(BAR() <= 0, 1, BAR())

    DO CASE
      CASE put_prn = 1  && Формирование и выдача печатного отчета на экран монитора

        REPORT FORM (ima_frx) PREVIEW

      CASE put_prn = 3  && Формирование и выдача печатного отчета на принтер

        colvo = 1

        DO FORM (put_scr) + ('colcopy.scx')  && Выбор количества копий при выдаче печатного отчета на принтер

        IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

          FOR I = 1 TO colvo

            DO CASE
              CASE par_prompt = .T.  && При выдаче отчета на печать предлагать панель выбора подключенных принтеров

                REPORT FORM (ima_frx) TO PRINTER PROMPT NOCONSOLE

              CASE par_prompt = .F.  && При выдаче отчета на печать не предлагать панель выбора подключенных принтеров

                REPORT FORM (ima_frx) TO PRINTER NOCONSOLE

            ENDCASE

          ENDFOR

        ENDIF

    ENDCASE

    EXIT

  ELSE
    EXIT
  ENDIF

ENDDO

RETURN


********************************************************** PROCEDURE prn_form_branch *******************************************************************

PROCEDURE prn_form_branch

DO CASE
  CASE put_prn = 1  && Печатный отчет на экран

    REPORT FORM (ima_frx) PREVIEW

  CASE put_prn = 3  && Печатный отчет на принтер

    colvo = 1

    DO FORM (put_scr) + ('colcopy.scx')  && Выбор количества копий при выдаче печатного отчета на принтер

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      FOR I = 1 TO colvo

        DO CASE
          CASE par_prompt = .T.  && При выдаче отчета на печать предлагать панель выбора подключенных принтеров

            REPORT FORM (ima_frx) TO PRINTER PROMPT NOCONSOLE

          CASE par_prompt = .F.  && При выдаче отчета на печать не предлагать панель выбора подключенных принтеров

            REPORT FORM (ima_frx) TO PRINTER NOCONSOLE

        ENDCASE

      ENDFOR

    ENDIF

ENDCASE

RETURN


*********************************************************** PROCEDURE vxod_servis **********************************************************************

PROCEDURE vxod_servis

SELECT parol
REPLACE ALL use_soft WITH .F., colvo_soft WITH 0
GO TOP

SELECT * ;
  FROM parol ;
  WHERE ALLTRIM(parol.kod) == '99' ;
  INTO CURSOR sel_parol

SELECT sel_parol
GO TOP

wes = sel_parol.pravo
adm_kod = ALLTRIM(sel_parol.kod)
par_kod = ALLTRIM(sel_parol.kod)
fio_user = ALLTRIM(sel_parol.fio)
par_fio = ALLTRIM(sel_parol.fio)
par_ima = ALLTRIM(sel_parol.ima)
par_par = ALLTRIM(sel_parol.parol)
par_prim = ALLTRIM(sel_parol.prim)
par_prompt = sel_parol.pr_prompt
par_grup = ALLTRIM(sel_parol.grup)
kod_kassa = ALLTRIM(sel_parol.kassa)

pr_registr = .F.  && Признак регистрации в ПО "CARD-виза филиал" выключается

SELECT parol
SET ORDER TO ima
GO TOP

IF SEEK(ALLTRIM(par_ima))
  REPLACE use_soft WITH .T., colvo_soft WITH colvo_soft + 1
ENDIF

RETURN


************************************************************* PROCEDURE rabota_avt *********************************************************************

PROCEDURE rabota_avt

SELECT data_odb
SET ORDER TO data

poisk_data_odb = DATE()
new_data_odb = DATE()

rec = DTOS(poisk_data_odb)

poisk = SEEK(rec)

IF poisk = .T.

  GO BOTTOM

  new_data_odb = data_odb.data

  SKIP -1
  arx_data_odb = data_odb.data

  SKIP  + 1

ELSE

  GO TOP
  GO BOTTOM

  new_data_odb = data_odb.data

  SKIP -1
  arx_data_odb = data_odb.data

  SKIP  + 1

ENDIF

ACTIVATE WINDOW poisk
@ WROWS()/3,3 SAY PADC('Внимание! Дата текущего опердня ' + DT(new_data_odb) + '  Дата архивного опердня ' + DT(arx_data_odb),WCOLS())
=INKEY(2)
DEACTIVATE WINDOW poisk

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SET ESCAPE ON

IF pr_dostup = .T.  && Признак разрешения на вход в ПО

  tit = ' Возможности: проверка подписи, снятие подписи, импорт данных из головного банка об операциях по картам'
  _Screen.Caption = tit

  DO start_avt IN Import

ELSE
  =err('Программный комплекс используется в монопольном режиме, ВХОД ЗАПРЕЩЕН ... ')
  DEACTIVATE POPUP glmenu
ENDIF

SET ESCAPE OFF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

tit = ' Возможности: закрытие операционного дня банка в ПО "CARD-VISA филиал и переход на работу с новой датой"'
_Screen.Caption = tit

DO start_avt IN Zakritie_odb

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

tit = ' Программный комплекс "CARD-VISA филиал" - формирование бухгалтерских документов по карточному обслуживанию'
_Screen.Caption = tit

RETURN


********************************************************************************************************************************************************




