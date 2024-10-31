***********************
*** �������� ������ ***
***********************

* ----------------------------------------------------------------------------------------------------------------------------------------- H�������� ��������� ----------------------------------------------------------------------------------------------------- *

CLOSE DATABASES  && �������� �������� ������.
CLEAR ALL  && �������� �����

SET STATUS BAR ON && ��������� ������ ������
SET OPTIMIZE ON  && �������� ����� �����������
SET CURSOR OFF && ������ ������� ��� ��������
SET ESCAPE OFF && ������ �� ����� �� esc
SET TALK OFF && �� ������� �� ����� ���������� ���������� ������
SET SAFETY OFF && ������ �� ����� ��������� � ���������� �����
SET MESSAGE TO 24 CENTER && ������ ���������, ������� ����� ������������.
SET DELETED ON  && ������ �� ������ � ���������� ��������
SET NEAR OFF && ������������� ��������� ������ �� ����� �������, ���� ����� ������ � ������� ������� FIND ��� SEEK ���������� ���������.
SET EXACT ON && ���������, ��� ��������� ����� ������������, ���� ��� ��������� ����������� ������ �� ����� ���������, �������������� ������.
SET ANSI ON  && ��������� �� ������������� � ��������� ANSI
SET NOTIFY OFF && ��������� ��� �������� ����������� ��������� ��������� ���������.
SET STRICTDATE TO 0  && ������������� � 2000 �����
SET MEMOWIDTH TO 30 && ������ ������ ���� ����� (�������)
SET WINDOW OF MEMO TO nazvan && ���� ��� ������ ���� ����� �� ����� ���������
SET BELL ON && �������� ��� ��������� �������� ������ ����������, � ����� ������������� �������� �������.
SET CLOCK OFF  && ��������� ��������� ����
SET SYSMENU OFF  && ��������� ��������� ����
SET CONFIRM ON && ���������, ��� �� �� ������ �������� �� ���������� ����, ����� ������ ������ ��� ���������� �������.
SET COLOR TO && ������ ����� � �������� ����� (����� �� ���������)
SET INTENSITY OFF && ������ ������� �������� ������ ��� ����������� ����� ��������������.
SET COLLATE TO 'RUSSIAN' && ������� ����������
SET COMPATIBLE OFF && ��������� ����������, ��������� � FoxBASE+, �������� � Visual FoxPro ��� ���������.
SET UDFPARMS TO VALUE && �������� ���������� �� ��������.
SET DEBUG OFF && ������ ���� ������� � ����������� ���������� ��� ������������
SET STEP OFF && �������� ���� �����������
SET ECHO OFF && ������ �����������
SET CURRENCY RIGHT && ������������ �� ������� ����
SET UNIQUE OFF  && ������� ���������� ���������� ��������
SET TABLEVALIDATE TO 2  && ������� �������� �������
SET REPORTBEHAVIOR 80  && ��������� �������� 80/90. ��������� ���� ���������������� ��������� ��������� ������, ��������� ����������� Visual FoxPro 9.0
SET OLEOBJECT ON  && ����������, ��� ��� ������������� ���������� �������� �������, ������� Visual FoxPro ��������� ����� ������������ �������� � ������� Windows

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SET HELP OFF  && ��������� ��������� ������

IF FILE(SYS(5) + CURDIR() + ('Help.chm'))

  SET HELP TO SYS(5) + CURDIR() + ('Help.chm')
  SET HELP ON  && �������� ��������� ������

ENDIF

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SET SYSFORMATS OFF && ��������� ������� ���������

SET HOURS TO 24  && ��������� ����� � ������
SET DATE TO GERMAN  && ������ ����
SET FDOW TO 2 && ������ ���� ������ �����������
SET FWEEK TO 2 && ������ ������ ����
SET CENTURY ON && ������ ������ ����
SET CURRENCY TO ' ���.' && ��� ������ �� ���������
SET POINT TO '.' && ��� ����������� ����� � ������� �����.
SET MARK TO '.' && ��� ����������� � ������� ����
SET DECIMALS TO 2 && ���������� ������ ������� �����
SET FIXED OFF && �������� ����� ������� ����� (�������)

PUSH KEY CLEAR

* --------------------------------------------------------------------------------------------------------- ��������� ��� ��������������������� ������ � ���� ---------------------------------------------------------------------------------------- *

SET EXCLUSIVE OFF && ������ ������������ ������� ������
SET LOCK OFF  && ������ ���������� ������
SET MULTILOCKS ON && ��������� ��������� ������� ���������� ���������� �������
SET REPROCESS TO AUTOMATIC  && �������������� ����������
SET REFRESH TO 3 && ���������� ������/������ (���)

_SCREEN.ICON = "logoink1.ico"
_SCREEN.WINDOWSTATE = 2    && ������������� �������� ���� ���������

* ---------------------------------------------------------------------- ��������� �������������� SET ���������� ��� Visual FoxPro 9.0 ��� ������������� � ������� 7.0 --------------------------------------------------------- *

IF LEFT(ALLTRIM(VERSION()), 16) == 'Visual FoxPro 09'
  SET ENGINEBEHAVIOR 70
ENDIF

* -------------------------------------------------------------------------------------------------------------- ��� ������ ������������ �� ------------------------------------------------------------------------------------------------------------------- *

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

a = AT('\', SYS(16), i)  && SYS(16) - ��� ����� ����������� ��������� � ������ �����

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

menu_ima = '�������� ������'
popup_ima = '�������� ������'
prompt_ima = '�������� ������'

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

put_bmp = (Direktoria) + ('BMP\')  && ���������� ��� ����������� ������
put_arx = (Direktoria) + ('ARX\')  && ���������� ��� �������� ����� ������, �������� ����� ��������
put_dbf = (Direktoria) + ('DBF\')  && ���������� ��� ������, ��������
put_frx = (Direktoria) + ('FRX\')  && ���������� ��� �������� �������
put_scr = (Direktoria) + ('SCR\')  && ���������� ��� ����
put_vcx = (Direktoria) + ('VCX\')  && ���������� ��� ��������� ����
put_prg = (Direktoria) + ('PRG\')  && ���������� ��� ����������� ������
put_exe = (Direktoria) + ('EXE\')  && ���������� ��� ����������� ������
put_doc = (Direktoria) + ('DOC\')  && ���������� ��� ������� �������������� � Word
put_xls = (Direktoria) + ('XLS\')  && ���������� ��� ������� �������������� � XLS

PUBLIC ima_sql_server

ima_sql_server = '������� ������ SQL Server 2012'

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC ARRAY azQEmptyGoKard(3), axPrkQEmptyGoKard(3) &&��� ��� - ���-�� �������� ����
PUBLIC ARRAY azNumPRK(3),azSumPRK(3) &&2 ����� ���
PUBLIC ARRAY azBHKPSumTotal(3), azNumKP(3),azSumKP(3),azNumPRK_KP(3),azSumPRK_KP(3) &&��� ����� ��
PUBLIC ARRAY azSomnit(3),azNeplat(3),azKpSumPlus(3),azKpSumMinus(3)
PUBLIC ARRAY azPRKRoute(3)
PUBLIC ARRAY azRzAfterOpKassa(3) &&2017-04 ������� ������ ���������
PUBLIC ARRAY azDtAfterOpKassa (3) &&2017-04 ���� ������������� � ������ �������.�����
PUBLIC ARRAY azDtPKO(3) &&2017-04 ���� �������� (��� ��� ���������)
PUBLIC ARRAY azPrkPr60322Count(3) && count ��� ����� BH Pr_60322=1 (��� ����� ���, ������� ������ ������� ������� �� ������������
PUBLIC ARRAY azPrkPr60322Sum(3) && Sum ��� ����� BH Pr_60322=1 (��� ����� ���, ������� ������ ������� ������� �� ������������

PUBLIC ARRAY axPRK(3),axKP(4)
PUBLIC ARRAY axKPQEmptyGoKard(6)
PUBLIC ARRAY axBHKPSumTotal(3)
PUBLIC ARRAY axPrkBudCount(3), axPrkBudSum(3)
PUBLIC ARRAY axKpBudCount1(4), axKpBudSum1(4) &&2017-04 GSL �������� ���� ��� ����� �� ������ � ����������� ����� (��� ��)
PUBLIC ARRAY axKpBudCount2(4), axKpBudSum2(6)
PUBLIC ARRAY axKpSumMinus(6), axKpSumPlus(6)
PUBLIC ARRAY axSomnit(6), axNeplat(6)
PUBLIC ARRAY axPrkPr60322Count(3), axPrkPr60322Sum(3)

* --------------------------------------------------- 01.02.2017 ��������� ���� ����� ----------------------------------------------------------------------------------------- *

PUBLIC axReportDate, axMiniOperDate
PUBLIC ARRAY axShiftDatePRK(3)
PUBLIC ARRAY axShiftDateKP(3)

* ------------------- 02.03.2019 ��������� ������ ��� ���������� �� ������ � ����������� ��������� (GSL)
PUBLIC g_oSess &&������� � ���� ���� ����� Login
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

text_mess = '�������� �� ���� - ��������� ���������� �������� ;  ����p ������ - ENTER ;  ����� - ESC'

* title_program = '����������� �������� "������" - ��� ��������� ����� ��������� (������ 05.00  �� 05.02.2014)'

title_program = '��� ��������� ����� ���������'

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

* ----------------------------------------------------------------------------------------- ����� ���������� ���� ������� � ����� ����� FileReport -------------------------------------------------------------------------------------------------- *

num_branch = '01'  && ����� ������� �. ������
* num_branch = '06'  && ����� ������� �. �����

PUBLIC disk_tmp, put_tmp, TmpPath, retDir, HistorySQL, DesktopPath, OldPath, FilePath, WorkPath, disk_tmp_dns, disk_tmp_ip, disk_tmp_loc
STORE SPACE(0) TO disk_tmp, put_tmp, TmpPath, retDir, DesktopPath, OldPath, FilePath, WorkPath, disk_tmp_dns, disk_tmp_ip, disk_tmp_loc

* WAIT WINDOW '������ ��� ������� � ����� - ' + LEFT(ALLTRIM(SYS(5)), 1) TIMEOUT 5
* WAIT WINDOW 'sergeizhuk_flag = ' + IIF(sergeizhuk_flag = .T., '.T.', '.F.') TIMEOUT 5

disk_tmp_dns = '\\kasper-nag.inkakhran.local'
disk_tmp_ip = '\\10.12.1.33'
disk_tmp_loc = 'D:'

* \\kasper-nag.inkakhran.local\FileReport\ - ��� RDP
* \\10.12.1.33\FileReport\ - ��� RDP
* D:\FileReport\ - ��� RDP

DO CASE
CASE LEFT(ALLTRIM(SYS(5)), 1) == 'D'

  DO CASE
  CASE sergeizhuk_flag = .T.

    disk_tmp = 'D:'
    put_tmp = (disk_tmp_loc) + ('\FileReport\')  && ���������� ��� ��������� ������


  CASE sergeizhuk_flag = .F.

    DO CASE
    CASE DIRECTORY(ALLTRIM(disk_tmp_dns) + ('\FileReport\'), 1) = .T.

      put_tmp = (disk_tmp_dns) + ('\FileReport\')  && ���������� ��� ��������� ������

    CASE DIRECTORY(ALLTRIM(disk_tmp_ip) + ('\FileReport\'), 1) = .T.

      put_tmp = (disk_tmp_ip) + ('\FileReport\')  && ���������� ��� ��������� ������

    OTHERWISE

      IF ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)) == 'KASPER-NAG'
        put_tmp = (disk_tmp_loc) + ('\FileReport\')  && ���������� ��� ��������� ������
      ELSE
        =err('�������� !!! ����������� ��� ������ ����� - FileReport - �� ��������!!!')
      ENDIF

    ENDCASE


  ENDCASE


CASE LEFT(ALLTRIM(SYS(5)), 1) <> 'D'

  DO CASE
  CASE DIRECTORY(ALLTRIM(disk_tmp_dns) + ('\FileReport\'), 1) = .T.

    put_tmp = (disk_tmp_dns) + ('\FileReport\')  && ���������� ��� ��������� ������

  CASE DIRECTORY(ALLTRIM(disk_tmp_ip) + ('\FileReport\'), 1) = .T.

    put_tmp = (disk_tmp_ip) + ('\FileReport\')  && ���������� ��� ��������� ������

  OTHERWISE

    IF ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)) == 'KASPER-NAG'
      put_tmp = (disk_tmp_loc) + ('\FileReport\')  && ���������� ��� ��������� ������
    ELSE
      =err('�������� !!! ����������� ��� ������ ����� - FileReport - �� ��������!!!')
    ENDIF

  ENDCASE

ENDCASE

DesktopPath = ALLTRIM(put_tmp)

* WAIT WINDOW 'put_tmp = ' + ALLTRIM(put_tmp) TIMEOUT 5
* WAIT WINDOW '������ ��� ������� � ����� - ' + LEFT(ALLTRIM(SYS(5)), 1) TIMEOUT 5
* WAIT WINDOW 'put_tmp = ' + ALLTRIM(put_tmp) TIMEOUT 5
* WAIT WINDOW 'DesktopPath = ' + ALLTRIM(DesktopPath) TIMEOUT 5

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
b = AT('\', SYS(16), i - 1)  && SYS(16) - ��� ����� ����������� ��������� � ������ �����
* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
SET PROCEDURE TO (put_prg) + ('prog_sbzhuk.prg')
* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
SET PROCEDURE TO (put_prg) + ('FunClib.prg') ADDITIVE && GSL 2016-03 ������ ����� �������� ������� �� FUNCLIB
* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
SET PROCEDURE TO (put_prg) + ('libvfp_excel.prg') ADDITIVE &&
* -------------------------------------------------------------------------------------------------------- ��� ���������� ������������ �� ------------------------------------------------------------------------------------------------------------------- *

* ���������� �������� ���������� ����������. ������ ����� �������� glSyst.dbf

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
NoShiftOpn = .F.    && ������� - ����� �� �������. ������ �� ����
HistorySQL = 1

* ������ ���������� � �� ����������

*DIMENSION usrPrivileg[2]
PUBLIC usrPrivileg[2]
usrPrivileg[1] = 5  && ��������� ��
usrPrivileg[2] = 21  && ���������� ��

* ������ ���������� ��� ������� ����

PUBLIC ar[1,2]

* -------------------------------------------------------------------------------------------------------------------------------------------------- ��� ������������ ������������ �� ------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

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

* dop_name_filial = '�����-������������� �-�'  && ������� ��� ������������ �������
dop_name_filial = ''
name_bank_1 = '������������ ���� 3454-�'  && �������� �������� �����������
name_bank_2 = '/2 ��� ��������ͻ (��)'  && �������� �������� �����������
name_bank = name_bank_1 + name_bank_2  && �������� �������� �����������
full_name_bank = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && ������ ������������, ���������� ����������� � ������
full_name_bank_ktl = '��������� ��� ���������� �������� � ���������� ��� " ��������" (��)  115201 �. ������ ��. ������������ , �. 8'  && ������ ������������, ���������� ����������� � ������, ��� ����������
full_name_bank_yar = '��������� ��� ���������� �������� � ���������� ��� " ��������" (��)  129366 �. ������ ��. ����������� , �. 11'  && ������ ������������, ���������� ����������� � ������, ��� ���������

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC put_ObmenToCBS, put_ObmenFromCBS, put_ObmenToGdasuc, put_ObmenFromGdasuc

put_ObmenFromCBS = 'ObmenFromCBS\'
put_ObmenToCBS = 'ObmenToCBS\'

put_ObmenFromGdasuc = 'ObmenFromGdasuc\'
put_ObmenToGdasuc = 'ObmenToGdasuc\'

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

Msg = '����� ������ ���� ��������� ��������� �� ������!!!'

PrivKK  = 'KK'  && ����� �������� ������ ������� ����������
PrivAll = 'All' && ����� �������� ����

IdPost_ = 5 && ��������� ������������
IdPost = 5 && ��������� ������������
IdDepartment_ = 5 && ����������� ������������
IdDepartment = 5 && ����������� ������������

* ��������� �����

* DIMENSION WorkStat[6]
PUBLIC WorkStat[6]
WorkStat[1] = "�� �������     "
WorkStat[2] = "���������������"
WorkStat[3] = "�����������    "
WorkStat[4] = "�����������    "
WorkStat[5] = "�������        "
WorkStat[6] = "���������      "

* ������� ���������� �����

* DIMENSION OrdStat[6]
PUBLIC OrdStat[6]
OrdStat[3] = 1
OrdStat[4] = 2
OrdStat[2] = 3
OrdStat[1] = 4
OrdStat[5] = 5
OrdStat[6] = 6

* ��� �����

* DIMENSION ArBdgtType[4]
PUBLIC ArBdgtType[4]
ArBdgtType[1] = "����������� �����"
ArBdgtType[2] = "����� �����"
ArBdgtType[3] = "��������� �����"
ArBdgtType[4] = "������������ �����"

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

num_branch = '01'  && ����� ������� �. ������
* num_branch = '06'  && ����� ������� �. �����

PUBLIC InkahranBIK, InkahranKorSchet, InkahranAcc, Kod_Filial, Boss_Kassa

STORE SPACE(0) TO InkahranBIK, InkahranKorSchet, InkahranAcc, Kod_LizNum, Boss_Kassa
DO CASE
CASE num_branch == '01'  && ����� ������� �. ������

  InkahranBIK = '044525934'
  InkahranKorSchet = ''
  InkahranAcc = '20202810955950000000'
  Kod_LizNum = '5595'

  Boss_Kassa = '�������� �.�.'  && ��� ���������� ��������� ����

CASE num_branch == '06'  && ����� ������� �. �����

  InkahranBIK = '044030302'
  InkahranKorSchet = ''
  InkahranAcc = '20202810140590000001'
  Kod_LizNum = '4059'

  Boss_Kassa = '��������� �.�.'  && ��� ���������� ��������� ����

ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DIMENSION ar[1,2]

PUBLIC IniFile, RcValue, ServerName, BdConn, BdConn_Load, RegIme, OdbcSrc, BdLogin, BdPsw, SqlQueryTimeOut, SqlWaitTime

STORE 0 TO ServerName, BdConn, BdConn_Load, SqlQueryTimeOut, SqlWaitTime
STORE '' TO ServerName, IniFile, RcValue, RegIme, OdbcSrc, BdLogin, BdPsw

xx = NEWOBJECT("oldinireg", "registry.vcx")

inifile = SYS(5) + CURDIR() + "inkohr_BS.ini"

* WAIT WINDOW 'inifile = ' + ALLTRIM(inifile) TIMEOUT 3

*  recountclient.inkakhran.local - SQL ������ - ��������� ������ 10.68.145.5
*  sql-dev.inkakhran.local - SQL ������ 2019 - ��� ������������
*  sqlkotcl.inkakhran.ru - 10.12.39.5  - ����� SQL ������ - ������ 10.68.145.5

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

*** ��������� ���������� ������, � ������� ������� ���������� ������� ������ ����������

IF SqlQueryTimeOut = 0
  SqlQueryTimeOut = 180
ENDIF

*** WaitTime - ����� � �������� ����� ������� Visual FoxPro ��������� ����������� �� ���������� ���������� SQL. �������� �� ��������� - 15 ������. ������/������.

IF SqlWaitTime = 0
  SqlWaitTime = 15
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

*    WAIT WINDOW 'ENDCASE - ����������� ������ - ' + LOWER(RIGHT(ALLTRIM(SYS(16)), 14)) TIMEOUT 4
*    WAIT WINDOW 'ServerName = ' + ALLTRIM(ServerName) TIMEOUT 3
*    WAIT WINDOW 'ENDCASE - Odbcsrc = ' + ALLTRIM(Odbcsrc) TIMEOUT 10
*    WAIT WINDOW 'Bdlogin = ' + ALLTRIM(Bdlogin) TIMEOUT 4
*    WAIT WINDOW 'Bdpsw = ' + ALLTRIM(Bdpsw) TIMEOUT 4

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

IF SqlWaitTime = 0
  SqlWaitTime = 100
ENDIF

IF EMPTY(odbcsrc) = .T. OR EMPTY(bdlogin) = .T.
  MESSAGEBOX('�������� odbcName ��� serverUserId � ����� inkohr_BS.ini ������� �� �����!', 16, '��� ���� ������ ...')
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

=LoadKeyboard('ENG')  && �������� ���������� ��������� ����������

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

      name_bank_1 = '��������� ��� ���������� �������� � ���������� ��� "��������" ��' + CHR(13)  && �������� �������� �����������
      name_bank_2 = ' �� ������: 115201, �. ������, ��. ������������, �. 8'  && �������� �������� �����������
      name_bank = name_bank_1 + name_bank_2  && �������� �������� �����������
      full_name_bank = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && ������ ������������, ���������� ����������� � ������, ��� ����������
      full_name_bank_ktl = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && ������ ������������, ���������� ����������� � ������, ��� ����������

      f_MestoCodeName = ' - �������� �� ������������'

      put_ObmenFromCBS = 'ObmenFromCBS\'
      put_ObmenToCBS = 'ObmenToCBS\'
      put_ObmenFromGdasuc = 'ObmenFromGdasuc\'
      put_ObmenToGdasuc = 'ObmenToGdasuc\'
      Kod_LizNum = '5595'


    CASE ALLTRIM(f_MestoCode) == 'YAR'

      name_bank_1 = '��������� ��� ���������� �������� � ���������� ��� "��������" ��' + CHR(13)  && �������� �������� �����������
      name_bank_2 = ' �� ������: 129366, �. ������, ��. �����������, �. 11'  && �������� �������� �����������
      name_bank = name_bank_1 + name_bank_2  && �������� �������� �����������
      full_name_bank = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && ������ ������������, ���������� ����������� � ������, ��� ���������
      full_name_bank_yar = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && ������ ������������, ���������� ����������� � ������, ��� ���������

      f_MestoCodeName = ' - �������� �� ���������'

      put_ObmenFromCBS = 'ObmenFromCBSYAR\'
      put_ObmenToCBS = 'ObmenToCBSYAR\'
      Kod_LizNum = '5594'


    CASE ALLTRIM(f_MestoCode) == 'ELV'

      f_MestoCodeName = ' - �������� �� �����������'
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

* ������ �� ���� ��������� ������ ��� ���� �� ������ � ����� ���������

    g_oSess = CREATEOBJECT("EMPTY")

    IF !F6GetSessInfo(g_oSess)
      MESSAGE("������ ������� 112 � �����.��������� ����������")
    ENDIF
* --------------------------------------------------------------------------- *

* WAIT WINDOW '1 - Golova - ����� ������������ ���������� BdConn_Load = ' + ALLTRIM(STR(BdConn_Load)) TIMEOUT 3

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    title_program_full = SPACE(5) + ALLTRIM(title_program) + '  ���� ����� �� - ' + DT(new_data_odb) + '  ���� ������� - ' + DT(skip_data_odb) + '  ' + ALLTRIM(f_NameSmena) + f_MestoCodeName + '  ����� ����� �� = ' + ALLTRIM(STR(f_IdShift_2))

    _SCREEN.CAPTION = ALLTRIM(title_program_full)

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    ShiftDate_ = OpenDate_
    m.ShiftDate_ = OpenDate_

    f_State = 0

* ------------ f_IdShift -  - ��� Id ���������� ����� �������� ����� �� ����������� Shift, ������� �������� ����� WHERE IdDepartment = 1 AND Status = 1 AND CloseDate IS NULL ---------------- *

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
* WAIT WINDOW '���� ������������� ��� - ' + DT(skip_data_odb) + '  ���� �������� ����� � �� - ' + DT(new_data_odb) TIMEOUT 3
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

*** Count_CursOpShift_3 - ���-�� ������� � ������� �� ������� ���� ��� ������� "3" Count_CursOpShift_2 -  - ���-�� ������� � ������� �� ������� ���� ��� ������� "1"

    DO CASE
    CASE Count_CursOpShift_3 = 0 AND Count_CursOpShift_2 > 0  && ������� 1 ����� ��, ����� ������ ��� � ������� "0". Count_CursOpShift_3 - ��� ������� "3" Count_CursOpShift_2 - ��� ������� "1"

      DO CASE
      CASE f_IdPost <> 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) <> '������� ��������� ��'
        NoShiftOpn = .T.

      CASE f_IdPost = 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) = '������� ��������� ��'
        NoShiftOpn = .T.

      ENDCASE


    CASE Count_CursOpShift_3 > 0 AND Count_CursOpShift_2 > 0  && ������� 2 ����� ��, ����� ����� � ������� "3", ������������, �������� �� ���� ����������.

      DO CASE
      CASE arm_version = 1  && �������� � ������� ���� ��������� �� - �������� ����� ������, ���������� � �������� �����. ������ = "3"

        DO CASE
        CASE f_IdPost <> 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) <> '������� ��������� ��'
          NoShiftOpn = .F.

        CASE f_IdPost = 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) = '������� ��������� ��'
          NoShiftOpn = .T.

        ENDCASE

      CASE arm_version = 3  && �������� � ������� ���� ��������� �� - ����������� �������� ����� ���������� �� ��� � ��. ������ = "1"

        DO CASE
        CASE f_IdPost <> 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) <> '������� ��������� ��'
          NoShiftOpn = .T.

        CASE f_IdPost = 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) = '������� ��������� ��'
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
  CASE num_branch == '01'  && ����� ������� �. ������

    IdControler_ = 819
    IdPost_ = 5
    f_MestoCode = 'KTL'

  CASE num_branch == '06'  && ����� ������� �. �����

    IdControler_ = 7
    IdPost_ = 10
    f_MestoCode = 'SPB'

  ENDCASE

  IdDepartment_ = 2

  f_IdControler = IdControler_
  f_IdUser = IdControler_
  f_IdPost = IdPost_
  f_IdDepartment = IdDepartment_

  NameControler_ = '��� �.�.'
  f_UserName = '��� �.�.'
  f_UserId = 790
  f_UserPassword = 790

* ------------------------------------------------------------------- *
  OpenDate_ = GetOpShiftKP()  && DATETIME()
* ------------------------------------------------------------------- *

* --------------------------------------------------------------------------------------------------------------------------- *
  DO CASE
  CASE ALLTRIM(f_MestoCode) == 'KTL'

    f_MestoCodeName = ' - �������� �� ������������'

  CASE ALLTRIM(f_MestoCode) == 'NAG'

    f_MestoCodeName = ' - �������� �� ��������'

  CASE ALLTRIM(f_MestoCode) == 'YAR'

    f_MestoCodeName = ' - �������� �� ���������'

  CASE ALLTRIM(f_MestoCode) == 'ELV'

    f_MestoCodeName = ' - �������� �� �����������'

  ENDCASE


* --------------------------------------------------------------------------------------------------------------------------- *

  title_program_full = SPACE(5) + ALLTRIM(title_program) + '  ���� ����� �� - ' + DT(new_data_odb) + '  ����� ����� �� = ' + ALLTRIM(STR(f_IdShift_2)) + '  ���� ������� - ' + DT(skip_data_odb) + '  ' + ALLTRIM(f_NameSmena) + f_MestoCodeName

  _SCREEN.CAPTION = ALLTRIM(title_program_full)

  ShiftDate_ = OpenDate_
  m.ShiftDate_ = OpenDate_

  f_State = 0

* ------------ f_IdShift -  - ��� Id ���������� ����� �������� ����� �� ����������� Shift, ������� �������� ����� WHERE IdDepartment = 1 AND Status = 1 AND CloseDate IS NULL ---------------- *

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
* WAIT WINDOW '���� ������������� ��� - ' + DT(skip_data_odb) + '  ���� �������� ����� � �� - ' + DT(new_data_odb) TIMEOUT 3
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

*** Count_CursOpShift_3 - ���-�� ������� � ������� �� ������� ���� ��� ������� "3" Count_CursOpShift_2 -  - ���-�� ������� � ������� �� ������� ���� ��� ������� "1"

  DO CASE
  CASE Count_CursOpShift_3 = 0 AND Count_CursOpShift_2 > 0  && ������� 1 ����� ��, ����� ������ ��� � ������� "0". Count_CursOpShift_3 - ��� ������� "3" Count_CursOpShift_2 - ��� ������� "1"

    DO CASE
    CASE f_IdPost <> 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) <> '������� ��������� ��'
      NoShiftOpn = .T.

    CASE f_IdPost = 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) = '������� ��������� ��'
      NoShiftOpn = .T.

    ENDCASE


  CASE Count_CursOpShift_3 > 0 AND Count_CursOpShift_2 > 0  && ������� 2 ����� ��, ����� ����� � ������� "3", ������������, �������� �� ���� ����������.

    DO CASE
    CASE arm_version = 1  && �������� � ������� ���� ��������� �� - �������� ����� ������, ���������� � �������� �����. ������ = "3"

      DO CASE
      CASE f_IdPost <> 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) <> '������� ��������� ��'
        NoShiftOpn = .F.

      CASE f_IdPost = 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) = '������� ��������� ��'
        NoShiftOpn = .T.

      ENDCASE

    CASE arm_version = 3  && �������� � ������� ���� ��������� �� - ����������� �������� ����� ���������� �� ��� � ��. ������ = "1"

      DO CASE
      CASE f_IdPost <> 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) <> '������� ��������� ��'
        NoShiftOpn = .T.

      CASE f_IdPost = 11  && 11 ��� ��������� ������� ��������� ��  && ALLTRIM(f_PostName) = '������� ��������� ��'
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