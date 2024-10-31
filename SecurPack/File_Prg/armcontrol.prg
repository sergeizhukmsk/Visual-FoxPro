***********************
*** Головной модуль ***
***********************

* ----------------------------------------------------------------------------------------------------------------------------------------- Hачальные установки ----------------------------------------------------------------------------------------------------- *

CLOSE DATABASES  && Закрытие открытых таблиц.
CLEAR ALL  && Очистить экран

SET STATUS BAR ON && выключить статус строку
SET OPTIMIZE ON  && Включить режим оптимизации
SET CURSOR OFF && курсор включен или выключен
SET ESCAPE OFF && запрет на выход по esc
SET TALK OFF && Не выводит на экран результаты исполнения команд
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
SET OLEOBJECT ON  && Определяет, что при невозможности определить описание Объекта, система Visual FoxPro выполняет поиск необходимого описания в Реестре Windows

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SET HELP OFF  && Отключить системную помощь

IF FILE(SYS(5) + CURDIR() + ('Help.chm'))

  SET HELP TO SYS(5) + CURDIR() + ('Help.chm')
  SET HELP ON  && Включить системную помощь

ENDIF

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SET SYSFORMATS OFF && настройки системы выключены

SET HOURS TO 24  && Установка часов в сутках
SET DATE TO GERMAN  && Формат даты
SET FDOW TO 2 && первый день недели понедельник
SET FWEEK TO 2 && первая неделя года
SET CENTURY ON && полный формат года
SET CURRENCY TO ' руб.' && тип валюты по умолчанию
SET POINT TO '.' && тип разделителя целой и дробной части.
SET MARK TO '.' && тип разделителя в формате даты
SET DECIMALS TO 2 && количество знаков дробной части
SET FIXED OFF && разрешен вывод дробной части (копейки)

PUSH KEY CLEAR

* --------------------------------------------------------------------------------------------------------- Установки для многопользовательской работы в сети ---------------------------------------------------------------------------------------- *

SET EXCLUSIVE OFF && Запрет монопольного захвата записи
SET LOCK OFF  && Запрет блокировки записи
SET MULTILOCKS ON && Разрешает выполнять попытку блокировки нескольких записей
SET REPROCESS TO AUTOMATIC  && Автоматическое обновление
SET REFRESH TO 3 && Обновление Чтение/Запись (сек)

_SCREEN.ICON = "logoink1.ico"
_SCREEN.WINDOWSTATE = 2    && Максимимзация главного окна программы

* ---------------------------------------------------------------------- Установка дополнительных SET параметров для Visual FoxPro 9.0 для совместимости с версией 7.0 --------------------------------------------------------- *

IF LEFT(ALLTRIM(VERSION()), 16) == 'Visual FoxPro 09'
  SET ENGINEBEHAVIOR 70
ENDIF

* -------------------------------------------------------------------------------------------------------------- Код нового разработчика ПО ------------------------------------------------------------------------------------------------------------------- *

PUBLIC Direktoria, i, a

sleg = 1
i = 3

DO WHILE sleg <> 0
  sleg = AT('\', SYS(16), i)
  IF sleg <> 0
    i = i + 1
  ELSE
    i = i - 2
  ENDIF
ENDDO

a = AT('\', SYS(16), i)  && SYS(16) - Имя файла выполняемой программы с полным путем

Direktoria = LEFT(SYS(16), a)

* WAIT WINDOW 'Direktoria - ' + ALLTRIM(Direktoria)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
SET PATH TO (Direktoria) + ('EXE\') + ';' + (Direktoria) + ('DBF\') + ';' + (Direktoria) + ('PRG\') + ';' + (Direktoria) + ('SCR\') + ';' + (Direktoria) + ('VCX\') + ';' + (Direktoria) + ('FRX\') + ';' + (Direktoria) + ('BMP\')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC menu_ima, popup_ima, prompt_ima, bar_num, colvo_wrows, colvo_wcols, sysmetric_cols, sysmetric_rows

colvo_wrows = WROWS()
colvo_wcols = WCOLS()

sysmetric_cols = SYSMETRIC(1)
sysmetric_rows = SYSMETRIC(2)

*sysmetric_cols = 1920
*sysmetric_rows = 1280

*sysmetric_cols = 1280
*sysmetric_rows = 1024

*sysmetric_cols = 1024
*sysmetric_rows = 768

*sysmetric_cols = 800
*sysmetric_rows = 600

menu_ima = 'Головной модуль'
popup_ima = 'Головной модуль'
prompt_ima = 'Головной модуль'

bar_num = 0

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
ON ERROR DO Errtrap WITH PROGRAM(), LINENO(), ERROR(), MESSAGE(), MESSAGE(1), LOWER(ALLTRIM(popup_ima)) + ' - ' + ALLTRIM(STR(bar_num))
* ON ERROR DO Errtrap_Sql WITH PROGRAM(), LINENO(), ERROR(), MESSAGE(), MESSAGE(1), LOWER(ALLTRIM(popup_ima)) + ' - ' + ALLTRIM(STR(bar_num))
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

PUBLIC DATA, COL, ROW, text_mess, per, tiplog, LOG, log_quit, dbfname, cdxname, fail, faildbf, failcdx, flag_add_table, num_branch, num_den_god, new_data_odb, arx_data_odb, skip_data_odb, oboi_data_odb

PUBLIC tex1, tex2, tex3, tex4, tex5, tex6, tex7, tex8, tex9, tex10, tex11, tex12, text1, text2, text3, text4, text5, text6, text7, text8, text9, text10, text11, text12, l_bar

PUBLIC BAR, bar_glmenu, analbar, putnew, put_bar, put_prn, put_vivod, put_select, err_text, poisk, colvo, colvo_zap, text_titul, predlog, simvol_klav, soob_time

PUBLIC dbfname, cdxname, colcopy, numord, barord, barbaze, client, mbal, mval, mkl, mliz, num_recno, pr_error, pr_lastkey_27, propis_summa_new

PUBLIC error_monopol, poisk_debit, poisk_credit, sel_doc_summa, name_bmp

PUBLIC tim1, tim2, tim3, tim4, tim5, tim6, tim7, tim8, tim9, tim10

PUBLIC time1, time2, time3, time4, time5, time6, time7, time8, time9, time10

PUBLIC text_mess, title_progam, erase_flag, sergeizhuk_sql_flag, sergeizhuk_flag, sergeizhuk_name_pk, ima_glav_file, servis_flag

PUBLIC num_schrift_menu_vert, num_schrift_menu_gor_1, num_schrift_menu_gor_2, num_schrift_popup_1, num_schrift_popup_2, num_schrift_windows, colvo_rows, colvo_cols

PUBLIC title_program_full, title_program, dat, sql_num, tit, stroka, sql_stroka, ARMvers, ARMname

PUBLIC f_Login, f_Password

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

PUBLIC put_bmp, put_arx, put_dbf, put_frx, put_scr, put_vcx, put_prg, put_exe, put_doc, put_xls

put_bmp = (Direktoria) + ('BMP\')  && Директория для графических файлов
put_arx = (Direktoria) + ('ARX\')  && Директория для архивных копий таблиц, архивных копий индексов
put_dbf = (Direktoria) + ('DBF\')  && Директория для таблиц, индексов
put_frx = (Direktoria) + ('FRX\')  && Директория для печатных отчетов
put_scr = (Direktoria) + ('SCR\')  && Директория для форм
put_vcx = (Direktoria) + ('VCX\')  && Директория для библиотек форм
put_prg = (Direktoria) + ('PRG\')  && Директория для программных файлов
put_exe = (Direktoria) + ('EXE\')  && Директория для исполняемых файлов
put_doc = (Direktoria) + ('DOC\')  && Директория для отчетов конвертируемых в Word
put_xls = (Direktoria) + ('XLS\')  && Директория для отчетов конвертируемых в XLS

PUBLIC ima_sql_server

ima_sql_server = 'Рабочий сервер SQL Server 2012'

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC ARRAY azQEmptyGoKard(3), axPrkQEmptyGoKard(3) &&Для ПРК - кол-во порожних карт
PUBLIC ARRAY azNumPRK(3),azSumPRK(3) &&2 смены ПРК
PUBLIC ARRAY azBHKPSumTotal(3), azNumKP(3),azSumKP(3),azNumPRK_KP(3),azSumPRK_KP(3) &&три смены КП
PUBLIC ARRAY azSomnit(3),azNeplat(3),azKpSumPlus(3),azKpSumMinus(3)
PUBLIC ARRAY azPRKRoute(3)
PUBLIC ARRAY azRzAfterOpKassa(3) &&2017-04 Признак режима оперкассы
PUBLIC ARRAY azDtAfterOpKassa (3) &&2017-04 Дата валютирования в режиме послеоп.кассы
PUBLIC ARRAY azDtPKO(3) &&2017-04 Дата окумента (ПКО или пакетника)
PUBLIC ARRAY azPrkPr60322Count(3) && count для сумок BH Pr_60322=1 (две смены ПРК, поэтому третий элемент массива не используется
PUBLIC ARRAY azPrkPr60322Sum(3) && Sum для сумок BH Pr_60322=1 (две смены ПРК, поэтому третий элемент массива не используется

PUBLIC ARRAY axPRK(3),axKP(4)
PUBLIC ARRAY axKPQEmptyGoKard(6)
PUBLIC ARRAY axBHKPSumTotal(3)
PUBLIC ARRAY axPrkBudCount(3), axPrkBudSum(3)
PUBLIC ARRAY axKpBudCount1(4), axKpBudSum1(4) &&2017-04 GSL добавляю поля для итога по операц и послеоперац кассе (для КП)
PUBLIC ARRAY axKpBudCount2(4), axKpBudSum2(6)
PUBLIC ARRAY axKpSumMinus(6), axKpSumPlus(6)
PUBLIC ARRAY axSomnit(6), axNeplat(6)
PUBLIC ARRAY axPrkPr60322Count(3), axPrkPr60322Sum(3)

* --------------------------------------------------- 01.02.2017 Добавляем дату смены ----------------------------------------------------------------------------------------- *

PUBLIC axReportDate, axMiniOperDate
PUBLIC ARRAY axShiftDatePRK(3)
PUBLIC ARRAY axShiftDateKP(3)

* ------------------- 02.03.2019 Параметры сеанса для возмещения за размен и контрольной ведомости (GSL)
PUBLIC g_oSess &&создаем в этой проц после Login
* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
simvol_klav = 'RUS'
erase_flag = .T.
import_flag = .T.
export_flag = .F.
* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
path_str = SYS(5) + SYS(2003)
put_tmp_dbf = SYS(2023) + '\'
num_branch = '01'
* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

STORE SPACE(0) TO tex1, tex2, tex3, tex4, tex5, tex6, tex7, tex8, tex9, tex10, tex11, tex12, text1, text2, text3, text4, text5, text6, text7, text8, text9, text10, text11, text12, l_bar

STORE SPACE(0) TO f_Login, f_Password, title_program_full, title_program, propis_summa_new

STORE 0 TO bar_glmenu, bar_num, bar_ord_gl, num_recno, time_home, time_end, colzap, rown, l_bar,  colvo_zap, put_bar, result_time, soob_time

STORE SECONDS() TO tim1, tim2, tim3, tim4, tim5, tim6, tim7, tim8, tim9, tim10

STORE DATETIME() TO time1, time2, time3, time4, time5, time6, time7, time8, time9, time10
STORE DATE() TO dt, dt1, dt2, data_frx, vvod_data, new_data_odb, arx_data_odb, skip_data_odb, oboi_data_odb

STORE .F. TO read_flag, pr_rabota_sql, log, log_quit, tiplog, pr_error, pr_lastkey_27

STORE YEAR(DATE()) TO god_tek, god1
STORE YEAR(DATE()) - 1 TO god_arx

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
ima_glav_file = ALLTRIM(LOWER(RIGHT(ALLTRIM(SYS(16)), 8)))
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

text_mess = 'Движение по меню - клавишами управления курсором ;  Выбоp пункта - ENTER ;  Выход - ESC'

* title_program = 'Программный комплекс "КасПер" - АРМ Контролер кассы Пересчета (Версия 05.00  от 05.02.2014)'

title_program = 'АРМ Контролер кассы Пересчета'

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

erase_flag = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'NKO-M-0069', 'NKO-M-0070', 'NKO-M-00071', 'NKO-M-0126', 'SERGEIZHUK', 'PKSBZHUK', 'NBSBZHUK'), .F., .T.)

sergeizhuk_flag = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'NKO-M-0069', 'NKO-M-0070', 'NKO-M-00071', 'NKO-M-0126', 'SERGEIZHUK', 'PKSBZHUK', 'NBSBZHUK','NKO-M-WIN528'), .T., .F.)

connect_sql = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'NKO-M-0069', 'NKO-M-0070', 'NKO-M-00071', 'NKO-M-0126', 'SERGEIZHUK', 'PKSBZHUK', 'NBSBZHUK'), .T., .F.)

sergeizhuk_sql_flag = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'NKO-M-0069', 'NKO-M-0070', 'NKO-M-00071', 'NKO-M-0126', 'SERGEIZHUK', 'PKSBZHUK', 'NBSBZHUK'), .T., .F.)

servis_flag = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'NKO-M-0069', 'NKO-M-0070', 'NKO-M-00071', 'NKO-M-0126', 'SERGEIZHUK', 'PKSBZHUK', 'NBSBZHUK'), .T., .F.)

sergeizhuk_name_pk = ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1))

* WAIT WINDOW 'sergeizhuk_name_pk = ' + ALLTRIM(sergeizhuk_name_pk)

* sergeizhuk_flag = .F.

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF sergeizhuk_flag = .T.
  ON KEY LABEL Ctrl+F12 DO otlad_pr
  ON KEY LABEL ALT+F12  DO otlad_pr WITH 1
  ON KEY LABEL ALT+F11  POP MENU _MSYSMENU
ENDIF

* ----------------------------------------------------------------------------------------- Новая реализация пути доступа к общей папки FileReport -------------------------------------------------------------------------------------------------- *

num_branch = '01'  && Номер филиала г. Москва
* num_branch = '06'  && Номер филиала г. Питер

PUBLIC disk_tmp, put_tmp, TmpPath, retDir, HistorySQL, DesktopPath, OldPath, FilePath, WorkPath, disk_tmp_dns, disk_tmp_ip, disk_tmp_loc
STORE SPACE(0) TO disk_tmp, put_tmp, TmpPath, retDir, DesktopPath, OldPath, FilePath, WorkPath, disk_tmp_dns, disk_tmp_ip, disk_tmp_loc

* WAIT WINDOW 'Запуск АРМ проишел с диска - ' + LEFT(ALLTRIM(SYS(5)), 1) TIMEOUT 5
* WAIT WINDOW 'sergeizhuk_flag = ' + IIF(sergeizhuk_flag = .T., '.T.', '.F.') TIMEOUT 5

disk_tmp_dns = '\\kasper-nag.inkakhran.local'
disk_tmp_ip = '\\10.12.1.33'
disk_tmp_loc = 'D:'

* \\kasper-nag.inkakhran.local\FileReport\ - без RDP
* \\10.12.1.33\FileReport\ - без RDP
* D:\FileReport\ - для RDP

DO CASE
CASE LEFT(ALLTRIM(SYS(5)), 1) == 'D'

  DO CASE
  CASE sergeizhuk_flag = .T.

    disk_tmp = 'D:'
    put_tmp = (disk_tmp_loc) + ('\FileReport\')  && Директория для временных файлов


  CASE sergeizhuk_flag = .F.

    DO CASE
    CASE DIRECTORY(ALLTRIM(disk_tmp_dns) + ('\FileReport\'), 1) = .T.

      put_tmp = (disk_tmp_dns) + ('\FileReport\')  && Директория для временных файлов

    CASE DIRECTORY(ALLTRIM(disk_tmp_ip) + ('\FileReport\'), 1) = .T.

      put_tmp = (disk_tmp_ip) + ('\FileReport\')  && Директория для временных файлов

    OTHERWISE

      IF ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)) == 'KASPER-NAG'
        put_tmp = (disk_tmp_loc) + ('\FileReport\')  && Директория для временных файлов
      ELSE
        =err('Внимание !!! Необходимая для работы папка - FileReport - не доступна!!!')
      ENDIF

    ENDCASE


  ENDCASE


CASE LEFT(ALLTRIM(SYS(5)), 1) <> 'D'

  DO CASE
  CASE DIRECTORY(ALLTRIM(disk_tmp_dns) + ('\FileReport\'), 1) = .T.

    put_tmp = (disk_tmp_dns) + ('\FileReport\')  && Директория для временных файлов

  CASE DIRECTORY(ALLTRIM(disk_tmp_ip) + ('\FileReport\'), 1) = .T.

    put_tmp = (disk_tmp_ip) + ('\FileReport\')  && Директория для временных файлов

  OTHERWISE

    IF ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)) == 'KASPER-NAG'
      put_tmp = (disk_tmp_loc) + ('\FileReport\')  && Директория для временных файлов
    ELSE
      =err('Внимание !!! Необходимая для работы папка - FileReport - не доступна!!!')
    ENDIF

  ENDCASE

ENDCASE

DesktopPath = ALLTRIM(put_tmp)

* WAIT WINDOW 'put_tmp = ' + ALLTRIM(put_tmp) TIMEOUT 5
* WAIT WINDOW 'Запуск АРМ проишел с диска - ' + LEFT(ALLTRIM(SYS(5)), 1) TIMEOUT 5
* WAIT WINDOW 'put_tmp = ' + ALLTRIM(put_tmp) TIMEOUT 5
* WAIT WINDOW 'DesktopPath = ' + ALLTRIM(DesktopPath) TIMEOUT 5

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
b = AT('\', SYS(16), i - 1)  && SYS(16) - Имя файла выполняемой программы с полным путем
* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
SET PROCEDURE TO (put_prg) + ('prog_sbzhuk.prg')
* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
SET PROCEDURE TO (put_prg) + ('FunClib.prg') ADDITIVE && GSL 2016-03 ТЕПЕРЬ МОЖНО ВЫЗЫВАТЬ функции из FUNCLIB
* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
SET PROCEDURE TO (put_prg) + ('libvfp_excel.prg') ADDITIVE &&
* -------------------------------------------------------------------------------------------------------- Код первичного разработчика ПО ------------------------------------------------------------------------------------------------------------------- *

* Отладочные значения глобальных переменных. Убрать после создания glSyst.dbf

IdShift_ = 1
IdOperDate = 1
IdShiftPRK_ = 1
ShiftDate_ = DATE()
OpenDate_ = DATE()
Regim_ = "" &&"o"
Idpost_ = 20
IdControler_ = 1
m.IdControler_ = 1
NameControler_ = ""
NoShiftOpn = .F.    && Признак - смена не открыта. Влияет на меню
HistorySQL = 1

* Список должностей и их привилегий

*DIMENSION usrPrivileg[2]
PUBLIC usrPrivileg[2]
usrPrivileg[1] = 5  && Контролер КП
usrPrivileg[2] = 21  && Специалист СБ

* Список привилегий для пунктов МЕНЮ

PUBLIC ar[1,2]

* -------------------------------------------------------------------------------------------------------------------------------------------------- Код действующего разработчика ПО ------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

PUBLIC f_TabNumber, f_OperCode, f_CurState, f_IdRoute, f_date_begin, f_date_end, f_BudgetNum, f_SignProc, f_IdTypeValue, f_SuperBudget, f_SuperQty, f_RecordNum, f_MyIdShift, f_NumBudget, f_TotalSum, f_TotalSumPRK, f_TotalSumKP, f_Wt_flag, AllOk

PUBLIC Itog_TotalSum_Prop, Summa_Prop, Chislo_Prop, f_IdShift, f_IdShift_1, f_IdShift_2, f_IdShift_3, f_IdShiftPRK, f_IdShiftKP, IdShift_KP, IdShift_PRK, f_IdCurShift, f_ShiftDate, f_NumRec, f_IdSignProc, f_IdRecord, f_NameOwner, f_IdOwner, f_Smena, f_PartShift

PUBLIC f_curDate, f_dtB, f_dtE, f_curDate_1, f_curDate_2, f_DatInkass, f_DtIncome, pr_TypeBox, pr_TypeFlg2Pin, f_State, f_IdClient, StatusSchift, f_Flg2PIN, f_ScanIdBudgetHistoryPRK, f_KassirKP, f_ClientGroup

PUBLIC f_IdFilial, f_FilialCode, f_FilialName, f_SignProc, f_IdTypeValue, f_IdDepartKP, f_IdDepartPRK, f_IdTranSheet, f_CurRecNo, Kasper_IdShift, selNumber, selNumber1, Retsql, SuperQty, Mes, Perinp1, Msg, LoginOk

PUBLIC f_OpDate, f_OpenDate, f_ReportDate, f_OpenTime, f_InkName, f_TblName, f_Sum1, f_Sum2, Detail_BudgetNum, Recnm, RecnmAll, ExactFlg, SumPack, f_SumPack, IdRec, f_IdRec

PUBLIC f_IdControler, f_IdUser, f_IdUserCBS, f_UserName, f_UserPassword, f_IdPost, f_IdDepartment, f_Login, f_UserId, f_IdBudgetKP, f_IdBudgetNumKP, f_IdPacketKP, Ret, SuperBudgetNum, CurRec, f_CurRec, SqlReq

PUBLIC load_forma, f_PacketNumKP, f_Wt_flag, f_StateNum, f_WorkType, f_CurRecPK, InnerSum, f_StateBudgetKP, f_SuperNum, f_SortRepeat, f_HardSum, f_IdPacketKP_2, f_IdPacketKP_3

PUBLIC f_IdDepartmentPRK, f_IdDepartmentPRK_1, f_IdDepartmentPRK_2, f_IdDepartmentPRK_3, f_IdDepartmentPRK_4, f_IdDepartmentKP, f_IdDepartmentKP_1, f_IdDepartmentKP_2, f_IdDepartmentKP_3, f_IdDepartmentKP_4

PUBLIC f_PackNum, f_NoteNum1, f_NoteNum2, curret, rc, IdShift_PRK, f_OperDate_OpDate, f_IdOperDate, f_StrShiftDate, f_StrIdSmena, f_NewShiftDate, f_NewIdSmena, f_IdUserIn, f_UserIn, f_IdUserOut, f_UserOut, f_IdPart, f_IdDay, f_IdShift_Arxiv, f_ClientGroup1

PUBLIC f_IdCashMessenger, f_IdRouteHistory, f_IdTransferSheet, f_BankName, f_NameSmena, f_IdBudgetHistoryPRK, rec_num, rec_zap, Type_Status, f_PostOwnerShifh, f_PostOwnerShifhKP, f_PostOwnerShifhPRK, Type_WorkType, Nach_Smena_KP, Kassir_UserName

PUBLIC Svod_Clic_UR, Type_Svod_Clic_UR, f_NumMagaz, pr_otladka, pr_tester, poisk_zap, f_StateKP, f_LocIdUser, f_IdGrup, Spisok_IdTranSheet, colvo_grid_zap, num_poz_sumka, f_NumerATM, f_MestoCode, f_MestoCodeName


STORE 0 TO f_IdFilial, f_IdShift, f_IdShift_1, f_IdShift_2, f_IdShift_3, f_IdCurShift, f_numRec, f_IdSignProcTrn, f_IdRecord, f_IdVedList, f_IdVedList_Id, f_IdShiftPRK, f_IdShiftKP, f_CurRecNo, Kasper_IdShift, f_SuperBudget, f_TotalSum, f_TotalSumPRK, f_TotalSumKP

STORE 0 TO f_IdCashMessenger, f_IdBank, f_FlRet, f_InkId, f_RouteId, f_RouteNum, Itog_BudgetNum, Itog_TotalSum, f_IdDepartKP, f_IdDepartPRK, f_IdTranSheet, pr_TypeBox, pr_TypeFlg2Pin, f_State, f_SuperQty, SumPack, f_SumPack, curret, rc

STORE 0 TO f_Flg2PIN, f_ScanIdBudgetHistoryPRK, f_Sum1, f_Sum2, f_IdBudgetKP, f_IdPacketKP, Ret, SuperBudgetNum, CurRec, Recnm, RecnmAll, ExactFlg, Retsql, SuperQty, IdRec, f_IdRec, colvo_grid_zap, num_poz_sumka

STORE 0 TO f_Wt_flag, f_StateNum, InnerSum, f_StateBudgetKP, f_SuperNum, f_SortRepeat, f_HardSum, f_IdPacketKP_2, f_IdPacketKP_3, f_OperDate_OpDate, f_IdOperDate

STORE 0 TO f_StrIdSmena, f_NewIdSmena, f_IdUserIn, f_IdUserOut, f_IdBudgetHistoryPRK, rec_num, rec_zap, f_PostOwnerShifh, f_PostOwnerShifhKP, f_PostOwnerShifhPRK, f_IdShift_Arxiv, f_IdOwner, f_StateKP, f_LocIdUser, f_IdGrup, f_IdUserCBS

STORE SPACE(0) TO Summa_Prop, Chislo_Prop, f_IdClient, selNumber, selNumber1,  f_UserName, f_UserPassword, f_Login, f_UserId, f_KassirKP, f_FilialCode, f_ClientGroup, f_InkName, f_TblName, SqlReq, Msg, f_BudgetNum, f_PacketNumKP, f_ClientGroup1

STORE SPACE(0) TO Spisok_IdTranSheet, f_NumerATM, f_MestoCode, f_MestoCodeName

STORE SPACE(0) TO f_WorkType, Detail_BudgetNum, f_UserIn, f_UserOut, f_NameSmena, Type_Status, IdShift_KP, IdShift_PRK, Type_WorkType, Nach_Smena_KP, Kassir_UserName, f_NameOwner, f_Smena, f_IdBudgetNumKP, f_NumMagaz, Perinp1

STORE 1 TO f_IdRouteHistory, f_IdTransferSheet, f_CurRecPK, f_IdPart, f_IdDay, f_IdControler, f_IdUser, f_IdPost, f_PartShift, Type_Svod_Clic_UR

STORE 1 TO f_IdDepartmentPRK, f_IdDepartmentPRK_1, f_IdDepartmentPRK_2, f_IdDepartmentPRK_3, f_IdDepartmentPRK_4

STORE 2 TO f_IdDepartment, f_IdDepartmentKP, f_IdDepartmentKP_1, f_IdDepartmentKP_2, f_IdDepartmentKP_3, f_IdDepartmentKP_4

STORE .F. TO StatusSchift, LoginOk, load_forma, AllOk, Svod_Clic_UR, pr_otladka, pr_tester, poisk_zap

STORE DATE() TO f_curDate, f_dtB, f_dtE, f_curDate_1, f_curDate_2

STORE DATETIME() TO f_OpDate, f_OpenDate, f_ReportDate, f_OpenTime, f_ShiftDate, f_DatInkass, f_DtIncome, f_StrShiftDate, f_NewShiftDate

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC dop_name_filial, name_bank, full_name_bank, full_name_bank_ktl, full_name_bank_yar

* dop_name_filial = 'Санкт-Петербургский ф-л'  && Добавка для наименования филиала
dop_name_filial = ''
name_bank_1 = 'Операционный офис 3454-К'  && Название головной организации
name_bank_2 = '/2 НКО «ИНКАХРАН» (АО)'  && Название головной организации
name_bank = name_bank_1 + name_bank_2  && Название головной организации
full_name_bank = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && Полное наименование, переменная вставляется в отчеты
full_name_bank_ktl = 'Помещение для совершения операций с ценностями НКО " ИНКАХРАН" (АО)  115201 г. Москва ул. Котляковская , д. 8'  && Полное наименование, переменная вставляется в отчеты, для Котляковки
full_name_bank_yar = 'Помещение для совершения операций с ценностями НКО " ИНКАХРАН" (АО)  129366 г. Москва ул. Ярославская , д. 11'  && Полное наименование, переменная вставляется в отчеты, для Ярославки

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC put_ObmenToCBS, put_ObmenFromCBS, put_ObmenToGdasuc, put_ObmenFromGdasuc

put_ObmenFromCBS = 'ObmenFromCBS\'
put_ObmenToCBS = 'ObmenToCBS\'

put_ObmenFromGdasuc = 'ObmenFromGdasuc\'
put_ObmenToGdasuc = 'ObmenToGdasuc\'

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

Msg = 'Здесь должно быть системное сообщение об ошибке!!!'

PrivKK  = 'KK'  && Пункт доступен только кассиру Контролеру
PrivAll = 'All' && Пункт доступен ВСЕМ

IdPost_ = 5 && Должность пользователя
IdPost = 5 && Должность пользователя
IdDepartment_ = 5 && Департамент пользователя
IdDepartment = 5 && Департамент пользователя

* Состояние сумки

* DIMENSION WorkStat[6]
PUBLIC WorkStat[6]
WorkStat[1] = "не вскрыта     "
WorkStat[2] = "пересчитывается"
WorkStat[3] = "пересчитана    "
WorkStat[4] = "сортируется    "
WorkStat[5] = "сведена        "
WorkStat[6] = "зачислена      "

* Порядок сортировки сумок

* DIMENSION OrdStat[6]
PUBLIC OrdStat[6]
OrdStat[3] = 1
OrdStat[4] = 2
OrdStat[2] = 3
OrdStat[1] = 4
OrdStat[5] = 5
OrdStat[6] = 6

* Тип сумки

* DIMENSION ArBdgtType[4]
PUBLIC ArBdgtType[4]
ArBdgtType[1] = "Стандартная сумка"
ArBdgtType[2] = "Супер сумка"
ArBdgtType[3] = "Вложенная сумка"
ArBdgtType[4] = "Многопиновая сумка"

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

num_branch = '01'  && Номер филиала г. Москва
* num_branch = '06'  && Номер филиала г. Питер

PUBLIC InkahranBIK, InkahranKorSchet, InkahranAcc, Kod_Filial, Boss_Kassa

STORE SPACE(0) TO InkahranBIK, InkahranKorSchet, InkahranAcc, Kod_LizNum, Boss_Kassa
DO CASE
CASE num_branch == '01'  && Номер филиала г. Москва

  InkahranBIK = '044525934'
  InkahranKorSchet = ''
  InkahranAcc = '20202810955950000000'
  Kod_LizNum = '5595'

  Boss_Kassa = 'Кипченко Н.В.'  && ФИО начальника кассового узла

CASE num_branch == '06'  && Номер филиала г. Питер

  InkahranBIK = '044030302'
  InkahranKorSchet = ''
  InkahranAcc = '20202810140590000001'
  Kod_LizNum = '4059'

  Boss_Kassa = 'Альчинова М.Ю.'  && ФИО начальника кассового узла

ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DIMENSION ar[1,2]

PUBLIC IniFile, RcValue, ServerName, BdConn, BdConn_Load, RegIme, OdbcSrc, BdLogin, BdPsw, SqlQueryTimeOut, SqlWaitTime

STORE 0 TO ServerName, BdConn, BdConn_Load, SqlQueryTimeOut, SqlWaitTime
STORE '' TO ServerName, IniFile, RcValue, RegIme, OdbcSrc, BdLogin, BdPsw

xx = NEWOBJECT("oldinireg", "registry.vcx")

inifile = SYS(5) + CURDIR() + "inkohr_BS.ini"

* WAIT WINDOW 'inifile = ' + ALLTRIM(inifile) TIMEOUT 3

*  recountclient.inkakhran.local - SQL сервер - прописать вместо 10.68.145.5
*  sql-dev.inkakhran.local - SQL сервер 2019 - для тестирования
*  sqlkotcl.inkakhran.ru - 10.12.39.5  - новый SQL сервер - вместо 10.68.145.5

* ServerName=recountclient.inkakhran.local

RcValue = ""
xx.getinientry(@RcValue, "ConnectSql", "ServerName", inifile)
ServerName = RcValue

* WAIT WINDOW 'ServerName = ' + ALLTRIM(ServerName) TIMEOUT 3

rcvalue = ""
xx.getinientry(@rcvalue, "ConnectSql", "ServerUserId", Inifile)
bdlogin = rcvalue

* WAIT WINDOW 'bdlogin = ' + ALLTRIM(bdlogin) TIMEOUT 2

rcvalue = ""
xx.getinientry(@rcvalue, "ConnectSql", "ServerUserPsw", Inifile)
bdpsw = rcvalue

* WAIT WINDOW 'bdpsw = ' + ALLTRIM(bdpsw) TIMEOUT 2

rcvalue = ""
xx.getinientry(@rcvalue, "ConnectSql", "OdbcName", Inifile)
odbcsrc = rcvalue

* WAIT WINDOW 'Odbcsrc = ' + ALLTRIM(Odbcsrc) TIMEOUT 3

rcvalue = ""
xx.getinientry(@rcvalue, "ConnectSql", "QueryTimeOut", Inifile)
SqlQueryTimeOut = VAL(rcvalue)

* WAIT WINDOW 'SqlQueryTimeOut = ' + ALLTRIM(STR(SqlQueryTimeOut)) TIMEOUT 2

rcvalue = ""
xx.getinientry(@rcvalue, "ConnectSql", "SqlWaitTime", Inifile)
SqlWaitTime = VAL(rcvalue)

* WAIT WINDOW 'SqlWaitTime = ' + ALLTRIM(STR(SqlWaitTime)) TIMEOUT 2

*** Указывает количество секунд, в течение которых инструкция ожидает снятия блокировки

IF SqlQueryTimeOut = 0
  SqlQueryTimeOut = 180
ENDIF

*** WaitTime - Время в секундах через которое Visual FoxPro проверяет завершилось ли выполнение инструкции SQL. Значение по умолчанию - 15 секунд. Чтение/запись.

IF SqlWaitTime = 0
  SqlWaitTime = 15
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

*    WAIT WINDOW 'ENDCASE - Запускающий модуль - ' + LOWER(RIGHT(ALLTRIM(SYS(16)), 14)) TIMEOUT 4
*    WAIT WINDOW 'ServerName = ' + ALLTRIM(ServerName) TIMEOUT 3
*    WAIT WINDOW 'ENDCASE - Odbcsrc = ' + ALLTRIM(Odbcsrc) TIMEOUT 10
*    WAIT WINDOW 'Bdlogin = ' + ALLTRIM(Bdlogin) TIMEOUT 4
*    WAIT WINDOW 'Bdpsw = ' + ALLTRIM(Bdpsw) TIMEOUT 4

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

IF SqlWaitTime = 0
  SqlWaitTime = 100
ENDIF

IF EMPTY(odbcsrc) = .T. OR EMPTY(bdlogin) = .T.
  MESSAGEBOX('Значение odbcName или serverUserId в файле inkohr_BS.ini указаны не верно!', 16, 'Так жить нельзя ...')
  RETURN
ENDIF

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
IF EnvironmentProc(3,SYS(16)) &&& 1
  DO vxod IN Armcontrol
ENDIF &&& 1
* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO Exit IN Prog_sbzhuk
* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

RETURN


************************************************************************************ PROCEDURE vxod ***********************************************************************

PROCEDURE vxod

* ------------------------------------------------ *
DO vxod_sql_server IN Vxod_sql
* ------------------------------------------------ *

* ------------------------------------------------ *
DO rasvert IN Prog_sbzhuk
* ------------------------------------------------ *

LoginOK = .F.

bar_office = 1

=LoadKeyboard('ENG')  && Включаем английскую раскладку клавиатуры

* WAIT WINDOW 'simvol_klav = ' + ALLTRIM(simvol_klav) TIMEOUT 3

connect_sql = .F.

DO CASE
CASE connect_sql = .F.

  TbLogin_Value = '790'
  TbPSW_Value = '790'

* ----------------------------- *
  DO FORM Login
* ----------------------------- *

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  DO CASE
  CASE f_UserId = '790' AND f_UserPassword = '790'
    f_IdPost = 11

  CASE f_UserId = '790_2' AND f_UserPassword = '790'
    f_IdPost = 11

  CASE f_UserId = '790_3' AND f_UserPassword = '790'
    f_IdPost = 11

  ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF LoginOK = .T.

    DO CASE
    CASE ALLTRIM(f_MestoCode) == 'KTL'

      name_bank_1 = 'Помещение для совершения операций с ценностями НКО "ИНКАХРАН" АО' + CHR(13)  && Название головной организации
      name_bank_2 = ' по адресу: 115201, г. Москва, ул. Котляковская, д. 8'  && Название головной организации
      name_bank = name_bank_1 + name_bank_2  && Название головной организации
      full_name_bank = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && Полное наименование, переменная вставляется в отчеты, для Котляковки
      full_name_bank_ktl = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && Полное наименование, переменная вставляется в отчеты, для Котляковки

      f_MestoCodeName = ' - площадка на Котляковской'

      put_ObmenFromCBS = 'ObmenFromCBS\'
      put_ObmenToCBS = 'ObmenToCBS\'
      put_ObmenFromGdasuc = 'ObmenFromGdasuc\'
      put_ObmenToGdasuc = 'ObmenToGdasuc\'
      Kod_LizNum = '5595'


    CASE ALLTRIM(f_MestoCode) == 'YAR'

      name_bank_1 = 'Помещение для совершения операций с ценностями НКО "ИНКАХРАН" АО' + CHR(13)  && Название головной организации
      name_bank_2 = ' по адресу: 129366, г. Москва, ул. Ярославская, д. 11'  && Название головной организации
      name_bank = name_bank_1 + name_bank_2  && Название головной организации
      full_name_bank = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && Полное наименование, переменная вставляется в отчеты, для Ярославки
      full_name_bank_yar = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && Полное наименование, переменная вставляется в отчеты, для Ярославки

      f_MestoCodeName = ' - площадка на Ярославке'

      put_ObmenFromCBS = 'ObmenFromCBSYAR\'
      put_ObmenToCBS = 'ObmenToCBSYAR\'
      Kod_LizNum = '5594'


    CASE ALLTRIM(f_MestoCode) == 'ELV'

      f_MestoCodeName = ' - площадка на Элеваторной'
      put_ObmenFromCBS = 'ObmenFromCBS\'
      put_ObmenToCBS = 'ObmenToCBS\'
      Kod_LizNum = '5593'

    ENDCASE

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

        SELECT mst.Code, mst.OpOffice
        FROM Mesto AS mst
        WHERE mst.IsActive = 1 --AND mst.Code = 'KTL'
        ORDER BY mst.Nn

    ENDTEXT

    ret = RunSQL('Scan_Mesto', @ar)

    IF USED('Scan_Mesto') AND RECCOUNT() > 0

      Pr_Use_Scan_Mesto = .T.

      INDEX ON Code TAG Mesto

      IF SEEK(f_MestoCode) = .T.
        full_name_bank = ALLTRIM(OpOffice)
      ENDIF

    ENDIF

* --------------------------------------------------------------------------- *
    OpenDate_ = GetOpShiftKP()  && DATETIME()
* --------------------------------------------------------------------------- *

* --------------------------------------------------------------------------- *

* Читаем из базы параметры сеанса для возм за размен и контр ведомости

    g_oSess = CREATEOBJECT("EMPTY")

    IF !F6GetSessInfo(g_oSess)
      MESSAGE("Печать справки 112 и Контр.ведомости невозможна")
    ENDIF
* --------------------------------------------------------------------------- *

* WAIT WINDOW '1 - Golova - Номер именованного соединения BdConn_Load = ' + ALLTRIM(STR(BdConn_Load)) TIMEOUT 3

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    title_program_full = SPACE(5) + ALLTRIM(title_program) + '  Дата смены КП - ' + DT(new_data_odb) + '  Дата опердня - ' + DT(skip_data_odb) + '  ' + ALLTRIM(f_NameSmena) + f_MestoCodeName + '  Номер смены КП = ' + ALLTRIM(STR(f_IdShift_2))

    _SCREEN.CAPTION = ALLTRIM(title_program_full)

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    ShiftDate_ = OpenDate_
    m.ShiftDate_ = OpenDate_

    f_State = 0

* ------------ f_IdShift -  - это Id уникальный номер открытой смены из справочника Shift, условия открытой смены WHERE IdDepartment = 1 AND Status = 1 AND CloseDate IS NULL ---------------- *

* WAIT WINDOW 'NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 5
* WAIT WINDOW 'OpenDate_ - ' + TTOC(OpenDate_) TIMEOUT 3
* WAIT WINDOW 'f_MestoCode = ' + ALLTRIM(f_MestoCode) TIMEOUT 3

* WAIT WINDOW 'f_IdShift_1 = ' + ALLTRIM(STR(f_IdShift_1)) TIMEOUT 3
* WAIT WINDOW 'f_IdShift_2 = ' + ALLTRIM(STR(f_IdShift_2)) TIMEOUT 3
* WAIT WINDOW 'f_IdShift_3 = ' + ALLTRIM(STR(f_IdShift_3)) TIMEOUT 3
* WAIT WINDOW 'f_IdShift_4 = ' + ALLTRIM(STR(f_IdShift_4)) TIMEOUT 3

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    poisk_date = TTOD(OpenDate_)
    new_data_odb = TTOD(OpenDate_)

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
* WAIT WINDOW 'Дата операционного дня - ' + DT(skip_data_odb) + '  Дата открытой смены в КП - ' + DT(new_data_odb) TIMEOUT 3
* WAIT WINDOW 'skip_data_odb - ' + DTOC(skip_data_odb) + '  new_data_odb - ' + DTOC(new_data_odb) TIMEOUT 3
* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
* WAIT WINDOW 'IdPost_ = ' + ALLTRIM(STR(IdPost_)) + '   IdDepartment_ = ' + ALLTRIM(STR(IdDepartment_)) + '   NameControler_ = ' + ALLTRIM(NameControler_) TIMEOUT 3
* WAIT WINDOW 'IdControler_ = ' + ALLTRIM(STR(IdControler_)) TIMEOUT 8
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* arm_version = 1

* Count_CursOpShift_3 = 0
* Count_CursOpShift_2 = 1

*      WAIT WINDOW 'Count_CursOpShift_3 = ' + ALLTRIM(STR(Count_CursOpShift_3)) + '   Count_CursOpShift_2 = ' + ALLTRIM(STR(Count_CursOpShift_2)) TIMEOUT 3
*      WAIT WINDOW 'f_IdPost = ' + ALLTRIM(STR(f_IdPost)) + '   NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 3

*** Count_CursOpShift_3 - кол-во записей в выборке из таблицы смен для статуса "3" Count_CursOpShift_2 -  - кол-во записей в выборке из таблицы смен для статуса "1"

    DO CASE
    CASE Count_CursOpShift_3 = 0 AND Count_CursOpShift_2 > 0  && Открыты 1 смена КП, более ранние все в статусе "0". Count_CursOpShift_3 - для статуса "3" Count_CursOpShift_2 - для статуса "1"

      DO CASE
      CASE f_IdPost <> 11  && 11 это должность Главный контролер КП  && ALLTRIM(f_PostName) <> 'Главный контролер КП'
        NoShiftOpn = .T.

      CASE f_IdPost = 11  && 11 это должность Главный контролер КП  && ALLTRIM(f_PostName) = 'Главный контролер КП'
        NoShiftOpn = .T.

      ENDCASE


    CASE Count_CursOpShift_3 > 0 AND Count_CursOpShift_2 > 0  && Открыты 2 смены КП, более рання в статусе "3", предзакрытая, действия по меню ограничены.

      DO CASE
      CASE arm_version = 1  && Работать с версией АРМа Контролер КП - пересчет сумок закрыт, подготовка к закрытию смены. Статус = "3"

        DO CASE
        CASE f_IdPost <> 11  && 11 это должность Главный контролер КП  && ALLTRIM(f_PostName) <> 'Главный контролер КП'
          NoShiftOpn = .F.

        CASE f_IdPost = 11  && 11 это должность Главный контролер КП  && ALLTRIM(f_PostName) = 'Главный контролер КП'
          NoShiftOpn = .T.

        ENDCASE

      CASE arm_version = 3  && Работать с версией АРМа Контролер КП - выполняется пересчет сумок переданных из ПРК в КП. Статус = "1"

        DO CASE
        CASE f_IdPost <> 11  && 11 это должность Главный контролер КП  && ALLTRIM(f_PostName) <> 'Главный контролер КП'
          NoShiftOpn = .T.

        CASE f_IdPost = 11  && 11 это должность Главный контролер КП  && ALLTRIM(f_PostName) = 'Главный контролер КП'
          NoShiftOpn = .T.

        ENDCASE

      ENDCASE


    ENDCASE

*      WAIT WINDOW 'f_IdPost = ' + ALLTRIM(STR(f_IdPost)) + '   NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 8

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* ------------------------------------------------ *
    DO Start IN Armcontrol_menu
* ------------------------------------------------ *

    IF rc > 0

      DO WHILE .T.

        ACTIVATE MENU armcontrol

        IF LOWER(PAD()) == 'quit'
          EXIT
        ELSE
          LOOP
        ENDIF

      ENDDO

    ENDIF

  ENDIF


* ===================================================================================================================================================================================== *


CASE connect_sql = .T.

  pr_otladka = .T.
  pr_tester = .T.

  DO CASE
  CASE num_branch == '01'  && Номер филиала г. Москва

    IdControler_ = 819
    IdPost_ = 5
    f_MestoCode = 'KTL'

  CASE num_branch == '06'  && Номер филиала г. Питер

    IdControler_ = 7
    IdPost_ = 10
    f_MestoCode = 'SPB'

  ENDCASE

  IdDepartment_ = 2

  f_IdControler = IdControler_
  f_IdUser = IdControler_
  f_IdPost = IdPost_
  f_IdDepartment = IdDepartment_

  NameControler_ = 'Жук С.Б.'
  f_UserName = 'Жук С.Б.'
  f_UserId = 790
  f_UserPassword = 790

* ------------------------------------------------------------------- *
  OpenDate_ = GetOpShiftKP()  && DATETIME()
* ------------------------------------------------------------------- *

* --------------------------------------------------------------------------------------------------------------------------- *
  DO CASE
  CASE ALLTRIM(f_MestoCode) == 'KTL'

    f_MestoCodeName = ' - площадка на Котляковской'

  CASE ALLTRIM(f_MestoCode) == 'NAG'

    f_MestoCodeName = ' - площадка на Нагорной'

  CASE ALLTRIM(f_MestoCode) == 'YAR'

    f_MestoCodeName = ' - площадка на Ярославке'

  CASE ALLTRIM(f_MestoCode) == 'ELV'

    f_MestoCodeName = ' - площадка на Элеваторной'

  ENDCASE


* --------------------------------------------------------------------------------------------------------------------------- *

  title_program_full = SPACE(5) + ALLTRIM(title_program) + '  Дата смены КП - ' + DT(new_data_odb) + '  Номер смены КП = ' + ALLTRIM(STR(f_IdShift_2)) + '  Дата опердня - ' + DT(skip_data_odb) + '  ' + ALLTRIM(f_NameSmena) + f_MestoCodeName

  _SCREEN.CAPTION = ALLTRIM(title_program_full)

  ShiftDate_ = OpenDate_
  m.ShiftDate_ = OpenDate_

  f_State = 0

* ------------ f_IdShift -  - это Id уникальный номер открытой смены из справочника Shift, условия открытой смены WHERE IdDepartment = 1 AND Status = 1 AND CloseDate IS NULL ---------------- *

* WAIT WINDOW 'NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 5
* WAIT WINDOW 'OpenDate_ - ' + TTOC(OpenDate_) TIMEOUT 3
* WAIT WINDOW 'f_MestoCode = ' + ALLTRIM(f_MestoCode) TIMEOUT 3

* WAIT WINDOW 'f_IdShift_1 = ' + ALLTRIM(STR(f_IdShift_1)) TIMEOUT 3
* WAIT WINDOW 'f_IdShift_2 = ' + ALLTRIM(STR(f_IdShift_2)) TIMEOUT 3
* WAIT WINDOW 'f_IdShift_3 = ' + ALLTRIM(STR(f_IdShift_3)) TIMEOUT 3

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  poisk_date = TTOD(OpenDate_)

  new_data_odb = TTOD(OpenDate_)

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
* WAIT WINDOW 'Дата операционного дня - ' + DT(skip_data_odb) + '  Дата открытой смены в КП - ' + DT(new_data_odb) TIMEOUT 3
* WAIT WINDOW 'skip_data_odb - ' + DTOC(skip_data_odb) + '  new_data_odb - ' + DTOC(new_data_odb) TIMEOUT 3
* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
* WAIT WINDOW 'IdPost_ = ' + ALLTRIM(STR(IdPost_)) + '   IdDepartment_ = ' + ALLTRIM(STR(IdDepartment_)) + '   NameControler_ = ' + ALLTRIM(NameControler_) TIMEOUT 3
* WAIT WINDOW 'IdControler_ = ' + ALLTRIM(STR(IdControler_)) TIMEOUT 8
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* f_IdPost = 5

* arm_version = 1

* Count_CursOpShift_3 = 0
* Count_CursOpShift_2 = 1

*      WAIT WINDOW 'Count_CursOpShift_3 = ' + ALLTRIM(STR(Count_CursOpShift_3)) + '   Count_CursOpShift_2 = ' + ALLTRIM(STR(Count_CursOpShift_2)) TIMEOUT 3
*      WAIT WINDOW 'f_IdPost = ' + ALLTRIM(STR(f_IdPost)) + '   NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 3

*** Count_CursOpShift_3 - кол-во записей в выборке из таблицы смен для статуса "3" Count_CursOpShift_2 -  - кол-во записей в выборке из таблицы смен для статуса "1"

  DO CASE
  CASE Count_CursOpShift_3 = 0 AND Count_CursOpShift_2 > 0  && Открыты 1 смена КП, более ранние все в статусе "0". Count_CursOpShift_3 - для статуса "3" Count_CursOpShift_2 - для статуса "1"

    DO CASE
    CASE f_IdPost <> 11  && 11 это должность Главный контролер КП  && ALLTRIM(f_PostName) <> 'Главный контролер КП'
      NoShiftOpn = .T.

    CASE f_IdPost = 11  && 11 это должность Главный контролер КП  && ALLTRIM(f_PostName) = 'Главный контролер КП'
      NoShiftOpn = .T.

    ENDCASE


  CASE Count_CursOpShift_3 > 0 AND Count_CursOpShift_2 > 0  && Открыты 2 смены КП, более рання в статусе "3", предзакрытая, действия по меню ограничены.

    DO CASE
    CASE arm_version = 1  && Работать с версией АРМа Контролер КП - пересчет сумок закрыт, подготовка к закрытию смены. Статус = "3"

      DO CASE
      CASE f_IdPost <> 11  && 11 это должность Главный контролер КП  && ALLTRIM(f_PostName) <> 'Главный контролер КП'
        NoShiftOpn = .F.

      CASE f_IdPost = 11  && 11 это должность Главный контролер КП  && ALLTRIM(f_PostName) = 'Главный контролер КП'
        NoShiftOpn = .T.

      ENDCASE

    CASE arm_version = 3  && Работать с версией АРМа Контролер КП - выполняется пересчет сумок переданных из ПРК в КП. Статус = "1"

      DO CASE
      CASE f_IdPost <> 11  && 11 это должность Главный контролер КП  && ALLTRIM(f_PostName) <> 'Главный контролер КП'
        NoShiftOpn = .T.

      CASE f_IdPost = 11  && 11 это должность Главный контролер КП  && ALLTRIM(f_PostName) = 'Главный контролер КП'
        NoShiftOpn = .T.

      ENDCASE

    ENDCASE


  ENDCASE

*      WAIT WINDOW 'f_IdPost = ' + ALLTRIM(STR(f_IdPost)) + '   NoShiftOpn = ' + IIF(NoShiftOpn = .F., '.F.', '.T.') TIMEOUT 3

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* ------------------------------------------------ *
  DO Start IN Armcontrol_menu
* ------------------------------------------------ *

  DO WHILE .T.

    ACTIVATE MENU armcontrol

    IF LOWER(PAD()) == 'quit'
      EXIT
    ELSE
      LOOP
    ENDIF

  ENDDO

ENDCASE

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO exit_sql_server IN Vxod_sql
* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

RETURN


************************************************************************** PROCEDURE Show_Help ******************************************************************************

PROCEDURE Show_Help

IF TYPE("_SCREEN.ActiveForm.HelpContextId") = "N" AND  _SCREEN.ACTIVEFORM.HELPCONTEXTID > 0

* WAIT WINDOW 'HELP ID _SCREEN' TIMEOUT 4

  HELP ID _SCREEN.ACTIVEFORM.HELPCONTEXTID

ELSE

* WAIT WINDOW 'HELP' TIMEOUT 4

  HELP

ENDIF

ENDFUNC


*****************************************************************************************************************************************************************************

