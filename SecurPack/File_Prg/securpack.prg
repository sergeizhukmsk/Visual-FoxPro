***********************
*** �������� ������ ***
***********************

* ----------------------------------------------------------------------------------------------------------------------------------------- H�������� ��������� ----------------------------------------------------------------------------------------------------- *

CLOSE DATABASES  && �������� �������� ������.
CLEAR ALL  && �������� �����

SET STATUS BAR ON  && ��������� ������ ������
SET OPTIMIZE ON  && �������� ����� �����������
SET CURSOR OFF  && ������ ������� ��� ��������
SET ESCAPE OFF  && ������ �� ����� �� esc
SET TALK OFF  && �� ������� �� ����� ���������� ���������� ������
SET SAFETY OFF  && ������ �� ����� ��������� � ���������� �����
SET MESSAGE TO 24 CENTER  && ������ ���������, ������� ����� ������������.
SET DELETED ON  && ������ �� ������ � ���������� ��������
SET NEAR OFF  && ������������� ��������� ������ �� ����� �������, ���� ����� ������ � ������� ������� FIND ��� SEEK ���������� ���������.
SET EXACT ON  && ���������, ��� ��������� ����� ������������, ���� ��� ��������� ����������� ������ �� ����� ���������, �������������� ������.
SET ANSI ON  && ��������� �� ������������� � ��������� ANSI
SET NOTIFY OFF  && ��������� ��� �������� ����������� ��������� ��������� ���������.
SET STRICTDATE TO 0  && ������������� � 2000 �����
SET MEMOWIDTH TO 30  && ������ ������ ���� ����� (�������)
SET WINDOW OF MEMO TO nazvan  && ���� ��� ������ ���� ����� �� ����� ���������
SET BELL ON  && �������� ��� ��������� �������� ������ ����������, � ����� ������������� �������� �������.
SET HELP OFF  && ��������� ��������� ������
SET CLOCK OFF  && ��������� ��������� ����
SET SYSMENU OFF  && ��������� ��������� ����
SET CONFIRM ON  && ���������, ��� �� �� ������ �������� �� ���������� ����, ����� ������ ������ ��� ���������� �������.
SET COLOR TO  && ������ ����� � �������� ����� (����� �� ���������)
SET INTENSITY OFF  && ������ ������� �������� ������ ��� ����������� ����� ��������������.
SET COLLATE TO 'RUSSIAN'  && ������� ����������
SET COMPATIBLE OFF  && ��������� ����������, ��������� � FoxBASE+, �������� � Visual FoxPro ��� ���������.
SET UDFPARMS TO VALUE  && �������� ���������� �� ��������.
SET DEBUG OFF  && ������ ���� ������� � ����������� ���������� ��� ������������
SET STEP OFF  && �������� ���� �����������
SET ECHO OFF  && ������ �����������
SET CURRENCY RIGHT  && ������������ �� ������� ����
SET UNIQUE OFF  && ������� ���������� ���������� ��������
SET TABLEVALIDATE TO 2  && ������� �������� �������
SET REPORTBEHAVIOR 80  && ��������� �������� 80/90. ��������� ���� ���������������� ��������� ��������� ������, ��������� ����������� Visual FoxPro 9.0

SET SYSFORMATS OFF  && ��������� ������� ���������

SET HOURS TO 24  && ��������� ����� � ������
SET DATE TO GERMAN  && ������ ����
SET FDOW TO 2  && ������ ���� ������ �����������
SET FWEEK TO 2  && ������ ������ ����
SET CENTURY ON  && ������ ������ ����
SET CURRENCY TO ' ���.'  && ��� ������ �� ���������
SET POINT TO '.'  && ��� ����������� ����� � ������� �����.
SET NULLDISPLAY TO ''  && ���������� �����, ������������ ��� ��������� null-��������
SET MARK TO '.'  && ��� ����������� � ������� ����
SET DECIMALS TO 2  && ���������� ������ ������� �����
SET FIXED OFF  && �������� ����� ������� ����� (�������)
SET ASSERTS ON

PUSH KEY CLEAR

* --------------------------------------------------------------------------------------------------------- ��������� ��� ��������������������� ������ � ���� ---------------------------------------------------------------------------------------- *

SET EXCLUSIVE OFF  && ������ ������������ ������� ������
SET LOCK OFF  && ������ ���������� ������
SET MULTILOCKS ON  && ���������� ������
SET REPROCESS TO AUTOMATIC  && �������������� ����������
SET REFRESH TO 5  && ���������� ������/������ (���)

* ---------------------------------------------------------------------- ��������� �������������� SET ���������� ��� Visual FoxPro 9.0 ��� ������������� � ������� 7.0 --------------------------------------------------------- *

IF LEFT(ALLTRIM(VERSION()), 16) == 'Visual FoxPro 09'
  SET ENGINEBEHAVIOR 70
ENDIF

* ------------------------------------------------------------------------------------------------------------------ ��� ������ ������������ �� ------------------------------------------------------------------------------------------------------------------- *

PUBLIC menu_ima, popup_ima, prompt_ima, bar_num, colvo_wrows, colvo_wcols, sysmetric_cols, sysmetric_rows

colvo_wrows = WROWS()
colvo_wcols = WCOLS()

sysmetric_cols = SYSMETRIC(1)
sysmetric_rows = SYSMETRIC(2)

menu_ima = '�������� ������'
popup_ima = '�������� ������'
prompt_ima = '�������� ������'

bar_num = 0

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
IF VERSION(2) = 0
  ON ERROR DO errtrap WITH PROGRAM(), LINENO(), ERROR(), MESSAGE(), MESSAGE(1), LOWER(ALLTRIM(popup_ima)) + ' - ' + ALLTRIM(STR(bar_num))
ENDIF
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC Direktoria, i, a, b

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

a = AT('\', SYS(16), i)
b = IIF(i <= 1, AT('\', SYS(16), 1), AT('\', SYS(16), i - 1))

Direktoria = LEFT(SYS(16), a)

* WAIT WINDOW 'Direktoria - ' + ALLTRIM(Direktoria)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
SET PATH TO (Direktoria) + ('EXE\') + ';' + (Direktoria) + ('DBF\') + ';' + (Direktoria) + ('PRG\') + ';' + (Direktoria) + ('SCR\') + ';' + (Direktoria) + ('FRX\') + ';' + (Direktoria) + ('MPR\') + ';' + (Direktoria) + ('BMP\')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

PUBLIC data, col, row, text_mess, per, tiplog, log, log_quit, dbfname, cdxname, fail, faildbf, failcdx, flag_add_table, num_branch, gorod_branch, result_time, skip_data_odb, fact_data_odb, skip_data_odb_test, new_data_odb, arx_data_odb, oboi_data_odb

PUBLIC tex1, tex2, tex3, tex4, tex5, tex6, tex7, tex8, tex9, tex10, tex11, tex12, text1, text2, text3, text4, text5, text6, text7, text8, text9, text10, text11, text12, l_bar

PUBLIC bar, bar_glmenu, bar_office, analbar, putnew, put_bar, put_prn, put_vivod, put_select, err_text, poisk, colvo, colvo_zap, text_titul, predlog, sqlReq, msg, status_pak, sel

PUBLIC dbfname, cdxname, colcopy, numord, barord, barbaze, client, mbal, mval, mkl, mliz, num_recno, simvol_klav, Text_grName, colvo_RecNum, colvo_IdRoute

PUBLIC error_monopol, poisk_debit, poisk_credit, sel_doc_summa, name_bmp, poisk_locate, pr_vvod, pr_error, Pr_Otladka, pr_vvod_sql, pr_lastkey_27

PUBLIC tim1, tim2, tim3, tim4, tim5, tim6, tim7, tim8, tim9, tim10

PUBLIC time1, time2, time3, time4, time5, time6, time7, time8, time9, time10

PUBLIC text_mess, title_progam, erase_flag, sergeizhuk_sql_flag, sergeizhuk_flag, sergeizhuk_name_pk, ima_glav_file, servis_flag, retval, sel_Sost

PUBLIC title_program, title_program_full, dat, sql_num, tit, rc, status_printer, ima_dbf, ima_scr, ima_frx, par_prompt, NewPhotoFileName, poisk_date, otvet, Text_BankName, CurIndex

PUBLIC nDetailsOfPaymentLen

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC put_bmp, put_arx, put_dbf, put_frx, put_scr, put_prg, put_exe, put_mpr, put_doc, put_xls, put_bmp_loc, put_bmp_net

put_bmp = (Direktoria) + ('BMP\')  && ���������� ��� ����������� ������
put_arx = (Direktoria) + ('ARX\')  && ���������� ��� �������� ����� ������, �������� ����� ��������
put_dbf = (Direktoria) + ('DBF\')  && ���������� ��� ������, ��������
put_frx = (Direktoria) + ('FRX\')  && ���������� ��� �������� �������
put_scr = (Direktoria) + ('SCR\')  && ���������� ��� ����
put_prg = (Direktoria) + ('PRG\')  && ���������� ��� ����������� ������
put_mpr = (Direktoria) + ('MPR\')  && ���������� ��� ������ ����
put_exe = (Direktoria) + ('EXE\')  && ���������� ��� ����������� ������
put_doc = (Direktoria) + ('DOC\')  && ���������� ��� ������� �������������� � Word
put_xls = (Direktoria) + ('XLS\')  && ���������� ��� ������� �������������� � Excel

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
SET PROCEDURE TO (put_prg) + ('prog_sbzhuk.prg')
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
simvol_klav = 'RUS'
erase_flag = .T.
import_flag = .T.
export_flag = .F.
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
path_str = SYS(5) + SYS(2003)
put_tmp_dbf = SYS(2023) + '\'
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

STORE SPACE(0) TO tex1, tex2, tex3, tex4, tex5, tex6, tex7, tex8, tex9, tex10, tex11, tex12, text1, text2, text3, text4, text5, text6, text7, text8, text9, text10, text11, text12, l_bar

STORE SPACE(0) TO ima_dbf, ima_scr, ima_frx, sqlReq, NewPhotoFileName, Text_grName, Text_BankName, status_pak, sel_Sost, title_program, title_program_full, num_branch, gorod_branch, sel

STORE 0 TO bar_glmenu, bar_num, bar_ord_gl, num_recno, time_home, time_end, colzap, rown, l_bar,  colvo_zap, put_bar, result_time, rc, bar_office, retval, otvet, CurIndex, colvo_RecNum, colvo_IdRoute

STORE SECONDS() TO tim1, tim2, tim3, tim4, tim5, tim6, tim7, tim8, tim9, tim10

STORE DATETIME() TO time1, time2, time3, time4, time5, time6, time7, time8, time9, time10
STORE DATE() TO dt, dt1, dt2, data_frx, vvod_data, skip_data_odb, fact_data_odb, skip_data_odb_test, new_data_odb, arx_data_odb, oboi_data_odb, poisk_date

STORE .F. TO read_flag, pr_rabota_sql, LOG, log_quit, tiplog, status_printer, poisk_locate, pr_error, Pr_Otladka, pr_vvod, pr_vvod_sql, pr_lastkey_27
STORE .T. TO par_prompt

STORE YEAR(DATE()) TO god_tek, god1
STORE YEAR(DATE()) - 1 TO god_arx

STORE 250 to nDetailsOfPaymentLen

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
ima_glav_file = ALLTRIM(LOWER(RIGHT(ALLTRIM(SYS(16)), 8)))
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

text_mess = '�������� �� ���� - ��������� ���������� �������� ;  ����p ������ - ENTER ;  ����� - ESC'

title_program = ' ��� "�������� �����������"'

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

erase_flag = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'NKO-M-0069', 'NKO-M-0070', 'NKO-M-00071', 'NKO-M-0126', 'NKO-M-300', 'SERGEIZHUK', 'PKSBZHUK', 'NBSBZHUK'), .F., .T.)

sergeizhuk_flag = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'NKO-M-0069', 'NKO-M-0070', 'NKO-M-00071', 'NKO-M-0126', 'NKO-M-300', 'SERGEIZHUK', 'PKSBZHUK', 'NBSBZHUK'), .T., .F.)

connect_sql = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'NKO-M-0069', 'NKO-M-0070', 'NKO-M-00071', 'NKO-M-0126', 'NKO-M-300', 'SERGEIZHUK', 'PKSBZHUK', 'NBSBZHUK'), .T., .F.)

sergeizhuk_sql_flag = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'NKO-M-0069', 'NKO-M-0070', 'NKO-M-00071', 'NKO-M-0126', 'NKO-M-300', 'SERGEIZHUK', 'PKSBZHUK', 'NBSBZHUK'), .T., .F.)

servis_flag = IIF(INLIST(ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1)), 'NKO-M-0069', 'NKO-M-0070', 'NKO-M-00071', 'NKO-M-0126', 'NKO-M-300', 'SERGEIZHUK', 'PKSBZHUK', 'NBSBZHUK'), .T., .F.)

sergeizhuk_name_pk = ALLTRIM(SUBSTR(SYS(0), 1, AT('#',SYS(0),1) -1))

* WAIT WINDOW 'sergeizhuk_name_pk = ' + ALLTRIM(sergeizhuk_name_pk)

* sergeizhuk_flag = .F.

* ------------------------------------------------ ������ ������ ������������ ��� ���������� ������������ ������������ ������ ����  ------------------------------------------------------------------------- *

PUBLIC modul_start, Pr_Forma_Lait

modul_start = 'securpack'

Pr_Forma_Lait = .F.

* securpack.exe - securpack_nag.exe - securpack_sev.exe - securpack_new.exe - securpack_upr.exe

* WAIT WINDOW 'modul_start = ' + ALLTRIM(modul_start) TIMEOUT 3

DO CASE
  CASE LOWER(RIGHT(ALLTRIM(SYS(16)), 13)) == ALLTRIM(modul_start) + '.exe' OR LOWER(RIGHT(ALLTRIM(SYS(16)), 13)) == ALLTRIM(modul_start) + '.fxp'

    DO CASE
      CASE sergeizhuk_flag = .T.

        Pr_Forma_Lait = .T.

      CASE sergeizhuk_flag = .F.

        Pr_Forma_Lait = .F.

    ENDCASE

  CASE LOWER(RIGHT(ALLTRIM(SYS(16)), 17)) == ALLTRIM(modul_start) + '_nag.exe' OR LOWER(RIGHT(ALLTRIM(SYS(16)), 14)) == ALLTRIM(modul_start) + '_nag.fxp'

    Pr_Forma_Lait = .T.

  CASE LOWER(RIGHT(ALLTRIM(SYS(16)), 17)) == ALLTRIM(modul_start) + '_sev.exe' OR LOWER(RIGHT(ALLTRIM(SYS(16)), 14)) == ALLTRIM(modul_start) + '_sev.fxp'

    Pr_Forma_Lait = .T.

  CASE LOWER(RIGHT(ALLTRIM(SYS(16)), 17)) == ALLTRIM(modul_start) + '_new.exe' OR LOWER(RIGHT(ALLTRIM(SYS(16)), 14)) == ALLTRIM(modul_start) + '_new.fxp'

    Pr_Forma_Lait = .T.

  CASE LOWER(RIGHT(ALLTRIM(SYS(16)), 17)) == ALLTRIM(modul_start) + '_upr.exe' OR LOWER(RIGHT(ALLTRIM(SYS(16)), 14)) == ALLTRIM(modul_start) + '_upr.fxp'

    Pr_Forma_Lait = .T.

  CASE LOWER(RIGHT(ALLTRIM(SYS(16)), 18)) == ALLTRIM(modul_start) + '_test.exe' OR LOWER(RIGHT(ALLTRIM(SYS(16)), 18)) == ALLTRIM(modul_start) + '_test.fxp'

    Pr_Forma_Lait = .T.

  OTHERWISE

    Pr_Forma_Lait = .F.

ENDCASE

* WAIT WINDOW 'Pr_Forma_Lait = ' + IIF(Pr_Forma_Lait = .T., '.T.', '.F.') TIMEOUT 5

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

num_branch = '01'  && ����� ������� �. ������
* num_branch = '06'  && ����� ������� �. �����

DO CASE
  CASE num_branch == '01'  && ����� ������� �. ������

    gorod_branch = '�. ������'

  CASE num_branch == '06'  && ����� ������� �. �����

    gorod_branch = '�. �����-���������'

ENDCASE

* ----------------------------------------------------------------------------------------- ����� ���������� ���� ������� � ����� ����� FileReport -------------------------------------------------------------------------------------------------- *

PUBLIC disk_tmp, put_tmp, tmpPath, retDir, DesktopPath, WorkPath, DNS_IP_disk_tmp, disk_tmp_dns, disk_tmp_ip, disk_tmp_loc
STORE SPACE(0) TO disk_tmp, put_tmp, tmpPath, retDir, DesktopPath, WorkPath, disk_tmp_dns, disk_tmp_ip, disk_tmp_loc

* DNS_IP_disk_tmp = .T.  && ���� ���������� DNS_IP_disk_tmp = .T., �� �������� ����� ������� �������, ��������� ����� DNS �����
* DNS_IP_disk_tmp = .F.  && ���� ���������� DNS_IP_disk_tmp = .F., �� �������� ������ �������, ��������� ����� IP �����

*  DO CASE
*    CASE LEFT(ALLTRIM(SYS(5)), 1) == 'D'

*      DO CASE
*        CASE sergeizhuk_flag = .T.

*          disk_tmp = 'D:'
*          put_tmp = (disk_tmp) + ('\FileReport\')  && ���������� ��� ��������� ������

*        CASE sergeizhuk_flag = .F.

*          DO CASE
*            CASE DNS_IP_disk_tmp = .F.

*              disk_tmp = '\\10.12.1.33'
*              put_tmp = (disk_tmp) + ('\FileReport\')  && ���������� ��� ��������� ������

*            CASE DNS_IP_disk_tmp = .T.

*              disk_tmp = '\\kasper-nag.inkakhran.local'
*              put_tmp = (disk_tmp) + ('\FileReport\')  && ���������� ��� ��������� ������

*          ENDCASE

*      ENDCASE


*    CASE LEFT(ALLTRIM(SYS(5)), 1) <> 'D'

*      DO CASE
*        CASE DNS_IP_disk_tmp = .F.

*          disk_tmp = '\\10.12.1.33'
*          put_tmp = (disk_tmp) + ('\FileReport\')  && ���������� ��� ��������� ������

*        CASE DNS_IP_disk_tmp = .T.

*          disk_tmp = '\\kasper-nag.inkakhran.local'
*          put_tmp = (disk_tmp) + ('\FileReport\')  && ���������� ��� ��������� ������

*      ENDCASE


*    OTHERWISE

*      DO CASE
*        CASE DNS_IP_disk_tmp = .F.

*          disk_tmp = '\\10.12.1.33'
*          put_tmp = (disk_tmp) + ('\FileReport\')  && ���������� ��� ��������� ������

*        CASE DNS_IP_disk_tmp = .T.

*          disk_tmp = '\\kasper-nag.inkakhran.local'
*          put_tmp = (disk_tmp) + ('\FileReport\')  && ���������� ��� ��������� ������

*      ENDCASE


*  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

PUBLIC disk_tmp_dns, disk_tmp_ip, disk_tmp_loc
STORE SPACE(0) TO disk_tmp_dns, disk_tmp_ip, disk_tmp_loc

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

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

PUBLIC PlaceTranzit, PlacePRK, IdOperDate, IdShiftPRK_, ShiftDate_, OpenDate_, Regim_, IdShift_, IdControler_, NameControler_
PUBLIC PrkSMCode, PrkNGCode, SkldCode, NoShiftOpn, rc, IdPost_, IdDepartment_, i, j, s, T, Name_file

PlaceTranzit = 6 && ��� ��������
PlacePRK = 2 && ��� ���

IdShift_ = 0
IdOperDate = 0
IdShiftPRK_ = 0
ShiftDate_ = DATETIME()
OpenDate_ = DATETIME()
Regim_ = '' &&"o"
IdControler_ = 0
IdPost_ = 0
IdDepartment_ = 0
NameControler_ = ''
PrkSMCode = '0000000000'  && ��� ��� ��� ��������
PrkNGCode = '1111111111'  && ��� ��� ��� ��������
SkldCode = '9999999999'  && ��� ��� ������
NumRowInsName_file = ''

noShiftOpn=.F.    && ������� - ����� �� �������. ������ �� ����
rc=1

PUBLIC ima_sql_server

ima_sql_server = '������� ������ SQL Server 2012'

* ��� �����

DIMENSION arBdgtType[3]

arBdgtType[1]="����������� �����"
arBdgtType[2]="����� �����"
arBdgtType[3]="��������� �����"

* --------------------------------------------------------------------------------------------------------------------------------------- ��� ������������ ������������ �� ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

PUBLIC curFile, numRowF, numRowErr, numRow, fvDate, f_TabNumber, f_OperCode, f_CurState, f_State, f_IdRoute, f_FileName, f_date_begin, f_date_end, f_fildName, f_fildData, f_flagAll, f_totalMess, f_Ch_FilialCode, f_PMDomination, f_QRDomination, f_IdShift_Begin, f_IdShift_End

PUBLIC f_IdFilial, f_IdFilialIn, f_IdFilialOut, f_FilialName, f_FilialNameIn, f_FilialNameOut, f_SignProc, f_DefBud, f_IdTypeValue, f_TypeValueKind, f_IsTransit, f_DatInkass, f_newBudget, f_TrnSheet_Id, f_RecIdRoute

PUBLIC f_IdShiftMemo, f_IncomeAccount, f_DateInkas, f_DtIncome, f_CurDate, f_BudgetNum, f_BudgetNum_dop, f_BudgetNum_poisk, f_BudgetNum_str, pr_vsp_mkb

PUBLIC table_sql, StrLimit, f_IdTrnSheet, f_IdSecurPack, f_IdGrup, f_IdBudgetHistory, f_DtBeginPer, f_IdShiftKP, f_ShiftDateKP, f_IdShiftPRK, f_IdPart, f_IdDay, f_Plomba, f_Osnovanie, f_Account, f_IncomeAccount, f_BuhSymbol, f_IdKeyATMDocMain, f_IdKeyATM, Pr_DtLost

PUBLIC f_UserId, f_Login, f_IdControler, f_IdUser, f_IdPost, f_IdDepartment, f_IdDepartmentKP, f_IdDepartmentPRK, f_IdDepartmentPRK_1, f_IdDepartmentPRK_2, f_IdDepartmentPRK_3, f_IdDepartmentUI, f_IdDepartmentATM, f_IdDepartmentKeyATM, f_UserPassword

PUBLIC f_IdFilial_str, f_FilialCode_str, f_FilialName_str, f_TypeValueKind_str, f_FlgRecalc_str, f_IsTransit_str, f_RecIdRoute, f_OperDateState, OpenDatePRK, flag_Shift_Blok, f_NumBaul, f_NumPlomba, f_EnablePlomba, Pr_Jk_Flg2PIN

PUBLIC f_UserPassword, StatusSchift, zag, zag_dop, ret, ret_mes, colvo_den, f_BudgetsAmount, f_MestoCode, f_MestoCodeName, f_IdOpenDepartment_OperDate, f_MkbRegId, SostPinString, f_IdAllert, f_AllertName, f_FilialNameKraska, f_TmAllarm

PUBLIC f_FilialCode, f_FilialCodeIn, f_FilialCodeOut, f_Filial2Code, f_Filial3Code, f_Filial4Code, f_Filial5Code, f_OpDate, f_OpenDate, f_ShiftDate, f_ReportDate, f_NewShiftDate, f_ArxShiftDate, f_IdShiftList, f_IdShiftPrn, f_IdShiftRab, f_IdShiftArx, f_Status_OperDate

PUBLIC colvo_zap_Shift_Rab, colvo_zap_Shift_Arx, perinp1, perinp2, f_IdPackDefect, f_IdBudgetDefect, f_IdRouteHistoryMemo, f_RecordTransferSheet, f_kkss_Id, f_kkss_UserId, f_kkss_UserName, f_IdClientInk, f_IdClient, pMode, f_ColvoKey, data_tsd, data_tsd_sql

PUBLIC f_TotalSum, f_TotalSum_dop, f_TotalSumGlPin, f_Pin1Sum, f_Pin2Sum, f_Pin3Sum, f_Pin4Sum, f_Pin5Sum, f_BudgetNum_pred, f_TotalSum_pred, f_TypeValueKind_pred, poisk_BudgetNum, f_IdVedomostiOutMemo, f_UserName, f_UserName_Prn, f_FullName, f_IdDefectCode

PUBLIC Pr_NumRowInsTsd, ColvoRowTsdRoute, IdTsdKasper, PoiskTSD_IdBudgetNum, num_mag, f_OpAllarmType, OpenShiftPRK, f_IdJk_FilialCode, f_IdGoKard, Pr_Lait, TsdFilialCode, f_Nominal_dop, StatusSchift, zag, zag_dop, ret, ret_mes, colvo_den, f_BudgetsAmount

PUBLIC f_TotalSum_RUB, f_TotalSum_USD, f_TotalSum_EUR, f_Select_Type, f_NumerATM, f_IdClientGroupMax, f_UpdateRec, f_PostName, f_Nominal, f_Colvo, f_Summa, Poisk_bank, Poisk_vsp, Pr_Poisk_vsp, Pr_IsJKServe

PUBLIC IdShift_PRK, IdShift_UI, IdShift_ATM, IdShift_KeyATM, IdOperDate_PRK, OperDate_PRK, ShiftDate_KP, ShiftDate_PRK, ReportDate_PRK, ShiftDate_ATM, ShiftDate_KeyATM, Pr_Exit_Form

PUBLIC Itog_BudgetNum, Itog_BudgetNum_1, Itog_BudgetNum_2, Itog_BuhVedom, Itog_TotalSum, Itog_BudgetNum_Prop, Itog_BuhVedom_Prop, Itog_TotalSum_Prop, Summa_Prop, Chislo_Prop, colvo_SecurPack, Test_ShiftDate, pr_rabota_arxiv

PUBLIC BudgetList_Cb_Ins_Visible, BudgetList_Cb_Del_Visible, f_poisk_date, f_InkName, f_OpenSmena, f_IdVedList_Id, f_IdCashMessenger, f_IdRouteHistory, f_IdTransferSheet, f_IdVedListMemo

PUBLIC f_IdControler, f_IdOperDate, f_IdOperDateStatus_1, f_IdOperDateStatus_2, f_IdOperDateArx, f_IdShift, f_IdShift_1, f_IdShift_2, f_IdCurShift, f_numRec, f_IdSignProc, f_IdSignProcTrn, f_IdSignProcPer, f_IdUser, f_IdRecord, pr_IdVedList, f_IdVedList

PUBLIC f_CurDate, f_dtB, f_dtE, f_CurDate_1, f_CurDate_2, f_curRec, f_idBank, f_flRet, f_InkName, f_InkId, f_RouteId, f_PrkSMCode, f_PrkNGCode, f_SkldCode, f_NoShiftOpn, colvo_DoublePin, stroka, pr_sumka_no_kp

PUBLIC f_newBudget, f_Transfer_Id, f_IncFullName, f_IncPostName, f_FormName, f_OldData, f_IdPack, f_FlgRecalc, f_SProcNm, f_Flg2PIN, f_Flg2PIN_str, colvo_pin_grup, f_curRec_Del, Pr_MestoCode, Pr_PointType, Pr_CodeTypePntNew

PUBLIC f_SuperBudget, f_WorkType, f_Qty, f_BHSuperBudgetId, f_MessFullName, f_Num_Memo, l_s, l_dt, l_err, l_mainBudgetNum, Memo_prn, Vedom_prn, name_office, f_NameBank, f_RouteNum, f_MemoNum, f_BankName, f_BankName_str, Value_Time

PUBLIC Pr_Otladka_Tsd, Pr_Blok_Tsd, Pr_NumRowInsTsd, PoiskTSD_IdBudgetNum, num_mag, pMode, Start_DateTimeRoute, Stop_DateTimeRoute

PUBLIC f_TotalSum_Flg2PIN_Full, f_TotalSum_Flg2PIN_2, f_TotalSum_Flg2PIN_3, TypeForma_DoublePin, f_AccountPIN_1, f_AccountPIN_2, f_AccountPIN_3

PUBLIC f_BudgetNum_1, f_BudgetNum_2, f_BudgetNum_3, f_FilialCode_1, f_FilialCode_2, f_FilialCode_3, f_BuxSimvolPIN_1, f_BuxSimvolPIN_2, f_BuxSimvolPIN_3, f_SummaPIN_All, f_SummaPIN_1, f_SummaPIN_2, f_SummaPIN_3

PUBLIC m.ScatteredADM, BDConnNesting, StatefullSuperbudget, f_IdKeyATMDocMain, f_IdKeyATM, f_Numer, f_Ch_KeyATM, f_Ch_KeyATMLost, f_IdReqATMUnLoad, f_FilialNameKraska, f_DtAllarm, ret_err


STORE SPACE(0) TO curFile, f_formName, Itog_BudgetNum_Prop, Itog_BuhVedom_Prop, Itog_TotalSum_Prop, Summa_Prop, Chislo_Prop, f_TypeValueKind, f_TypeValueKind_str, f_poisk_date, zag, zag_dop, poisk_BudgetNum, stroka, f_UserName, f_FullName

STORE SPACE(0) TO poisk_BudgetNum, stroka, f_UserName, f_FullName, f_IncFullName, TsdFilialCode, f_BudgetNum, f_BudgetNum_str, f_BudgetNum_poisk, f_AccountPIN_1, f_AccountPIN_2, f_AccountPIN_3

STORE SPACE(0) TO l_mainBudgetNum, name_office, f_NameBank, f_RouteNum, f_MemoNum, f_BankName, f_BankName_str, f_PostName, f_Select_Type, f_NumerATM, f_IdShiftList, f_Plomba, f_Account, f_IncomeAccount, f_BuhSymbol, Poisk_bank, Poisk_vsp, num_mag, f_IncomeAccount

STORE SPACE(0) TO f_FilialCode, f_FilialCodeIn, f_FilialCodeOut, f_Filial2Code, f_Filial3Code, f_Filial4Code, f_Filial5Code, table_sql, f_FilialCode_str, f_FilialName_str, f_NumBaul, f_InkName, f_NumPlomba, StatefullSuperbudget

STORE SPACE(0) TO f_BudgetNum_pred, f_TypeValueKind_pred, perinp1, perinp2, f_IdControler, f_FilialNameIn, f_FilialNameOut, f_kkss_UserName, f_kkss_UserId, f_IsTransit, f_IsTransit_str, f_Osnovanie, f_Numer, f_UserName_Prn, f_Login, f_UserId, f_UserPassword

STORE SPACE(0) TO f_InkName, f_NumPlomba, num_mag, f_MestoCode, f_MestoCodeName, f_FilialNameKraska, Pr_MestoCode, f_AllertName, f_FilialNameKraska, f_TmAllarm

STORE SPACE(0) TO f_BudgetNum_1, f_BudgetNum_2, f_BudgetNum_3, f_FilialCode_1, f_FilialCode_2, f_FilialCode_3, f_BuxSimvolPIN_1, f_BuxSimvolPIN_2, f_BuxSimvolPIN_3


STORE 0 TO f_SummaPIN_All, f_SummaPIN_1, f_SummaPIN_2, f_SummaPIN_3, TypeForma_DoublePin

STORE 0 TO f_IdFilial, f_IdFilialIn, f_IdFilialOut, f_IdOperDate, f_IdOperDateStatus_1, f_IdOperDateStatus_2, f_IdOperDateArx, f_IdShift, f_IdShift_1, f_IdShift_2, f_IdCurShift, f_NumRec, f_IdSignProc, f_IdSignProcTrn, f_IdSignProcPer, f_IdUser, f_IdRecord,BDConnNesting

STORE 0 TO Itog_BudgetNum, Itog_BudgetNum_1, Itog_BudgetNum_2, Itog_BuhVedom, Itog_TotalSum, f_TrnSheet_Id, numRowF, numRowErr, numRow, f_SProcNm, f_SignProc, f_IdRoute, f_IdRouteHistory, f_RecIdRoute, f_TabNumber, ret, ret_mes, colvo_den

STORE 0 TO f_TotalSum, f_TotalSum_dop, f_TotalSumGlPin, f_Pin1Sum, f_Pin2Sum, f_Pin3Sum, f_Pin4Sum, f_Pin5Sum, f_FlgRecalc_str, f_IdFilial_str, f_RecIdRoute, StrLimit, f_IdTrnSheet, f_Flg2PIN, f_Flg2PIN_str, colvo_DoublePin, Colvo_FilialCode, colvo_SecurPack

STORE 0 TO f_TotalSum_RUB, f_TotalSum_USD, f_TotalSum_EUR, f_IdShiftKP, f_IdShiftPRK, f_kkss_Id, f_Nominal, f_Colvo, f_Summa, Pr_Poisk_vsp, f_IdShiftMemo, f_QRDomination, f_IdKeyATMDocMain, f_IdKeyATM

STORE 0 TO f_IdClientInk, f_IdClient, pMode, f_ColvoKey, f_IdReqATMUnLoad, f_IdShiftPrn, f_IdShiftRab, f_IdShiftArx, colvo_zap_Shift_Rab, colvo_zap_Shift_Arx, Colvo_FilialCode

STORE 0 TO f_IdDepartment, f_IdDepartmentKP, f_IdDepartmentPRK, f_IdDepartmentPRK_1, f_IdDepartmentPRK_2, f_IdDepartmentPRK_3, f_IdDepartmentUI, f_IdDepartmentATM, f_IdDepartmentKeyATM, Pr_Jk_Flg2PIN, f_OperDateState

STORE 0 TO f_BudgetsAmount, f_IdRouteHistoryMemo, f_PMDomination, f_Status_OperDate, f_Status_OperDate, f_IdVedomostiOutMemo, flag_Shift_Blok, f_TotalSum_pred, f_IdCashMessenger, f_IdBank, f_FlRet, f_InkId, f_RouteId, f_IdGrup, f_IdBudgetHistory, f_IdVedListMemo

STORE 0 TO IdShift_PRK, IdShift_UI, IdShift_ATM, IdShift_KeyATM, IdOperDate_PRK, f_IdDefectCode, f_IdPackDefect, f_IdBudgetDefect, f_DefBud, f_RecordTransferSheet, colvo_pin_grup, f_IdClientGroupMax, f_UpdateRec

STORE 0 TO f_IdShift_Begin, f_IdShift_End, f_Ch_FilialCode, f_IdSecurPack, f_Num_Memo, f_IdVedList, f_IdVedList_Id, Value_Time, pMode, f_IdKeyATMDocMain, f_IdKeyATM, f_Ch_KeyATM, f_Ch_KeyATMLost

STORE 0 TO f_OperDateState, f_IdShiftPrn, f_IdShiftRab, f_IdShiftArx, colvo_zap_Shift_Rab, colvo_zap_Shift_Arx, Colvo_FilialCode, f_MkbRegId, Pr_PointType, Pr_CodeTypePntNew, f_IdAllert, f_IdJk_FilialCode, f_IdGoKard

STORE 0 TO IdTsdKasper, PoiskTSD_IdBudgetNum, f_IdOpenDepartment_OperDate, SostPinString, ret_err, Pr_Lait, ColvoRowTsdRoute, f_Nominal_dop, data_tsd_sql, Pr_IsJKServe

STORE 0 TO f_TotalSum_Flg2PIN_Full, f_TotalSum_Flg2PIN_2, f_TotalSum_Flg2PIN_3


STORE 1 TO f_curRec, f_IdTransferSheet, f_IdTypeValue, f_curRec_Del, f_IdPart, f_IdDay, f_OpAllarmType

STORE .T. TO BudgetList_Cb_Ins_Visible, BudgetList_Cb_Del_Visible, Memo_prn, Vedom_prn, pr_sumka_no_kp

STORE .F. TO pr_IdVedList, StatusSchift, f_EnablePlomba, Pr_Exit_Form, pr_rabota_arxiv, Pr_DtLost, pr_vsp_mkb

STORE .F. TO Pr_NumRowInsTsd, Pr_Otladka_Tsd, Pr_Blok_Tsd, OpenShiftPRK

STORE DATE() TO f_CurDate, f_dtB, f_dtE, f_CurDate_1, f_CurDate_2, f_OpenSmena, f_DtAllarm

STORE DATETIME() TO OperDate_PRK, ShiftDate_KP, ShiftDate_PRK, ReportDate_PRK, ShiftDate_ATM, ShiftDate_KeyATM, Test_ShiftDate, f_ShiftDateKP, f_DateInkas, f_DtIncome, f_CurDate

STORE DATETIME() TO f_DtIncome, f_ShiftDate, f_ReportDate, f_NewShiftDate, f_ArxShiftDate, OpenDatePRK, f_OpDate, f_OpenDate, f_DtBeginPer, data_tsd

STORE DATETIME() TO Start_DateTimeRoute, Stop_DateTimeRoute


* --------------------------------------------------------------- ���������� ���������� ��� �������� �����, ����������� �������� �������� � ��� ---------------------------------------------------------------------------------------------------------------------------------- *

PRIVATE pr_sumka_priem_kp, pr_sumka_priem_trn, pr_sumka_atm_per, pr_sumka_atm_trn, pr_sumka_vozm_per, pr_sumka_bank_vsp, pr_sumka_atm_vsp, pr_sumka_prk_kp, ;
  pr_sumka_prk_trn, pr_sumka_vozm_xran, pr_sumka_atm_prk_per, pr_sumka_atm_prk_trn, pr_sumka_atm_skl_trn, pr_sumka_prk_trn_bank, pr_sumka_atm_prk_trn_ved, pr_sumka_vozm_xran_ved

STORE .F. TO pr_sumka_priem_kp, pr_sumka_priem_trn, pr_sumka_atm_per, pr_sumka_atm_trn, pr_sumka_vozm_per, pr_sumka_bank_vsp, pr_sumka_atm_vsp, pr_sumka_prk_kp, ;
  pr_sumka_prk_trn, pr_sumka_vozm_xran, pr_sumka_atm_prk_per, pr_sumka_atm_prk_trn, pr_sumka_atm_skl_trn, pr_sumka_prk_trn_bank, pr_sumka_atm_prk_trn_ved, pr_sumka_vozm_xran_ved

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

f_PrkSMCode = PrkSMCode  && '0000000000'  && ��� ��� ��� ��������
f_PrkNGCode = PrkNGCode  && '1111111111'  && ��� ��� ��� ��������
f_SkldCode = SkldCode  && SkldCode = '9999999999'

f_SuperBudget = NULL
f_WorkType = '0'
f_SuperQty = 0
f_BHSuperBudgetId = 0
f_MessFullName = ''

PUBLIC TotalSum_WorkType_1, TotalSum_WorkType_2

TotalSum_WorkType_1 = 0.00
TotalSum_WorkType_2 = 0.00

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
Value_Time = VAL(LEFT(TIME(), 2))  && ���������� �������� ���� � ������� �������� ����� ���, ������������ ��� �������� ���� ��������� ����: ������� 12 ����� � ����� ���� 12 �����
* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

type_modul_arm = '01'  && ������ ������ ����� ����������� � ���
type_modul_arm = '02'  && ������ ������ ����� � ����������� � ���
type_modul_arm = '03'  && ������ ������ ������ ��� ���������� � ���
type_modul_arm = '04'  && ������ ������ ��
type_modul_arm = '01'  && ��������� ������������� ����������

*** ---------------------------------------------------------------------------------------- ������ ��� ��������, �����, ����������� ---------------------------------------------------------------------------------------------------------- ***

PUBLIC naznach, dolg_rukovod, fio_boss, fio_boss_key, post_boss, dolg_ispolnit, fio_spez, liz_num_plat_prk, liz_num_polu_prk, debit
PUBLIC liz_num_polu_kredit_prk, liz_num_polu_debit_prk, credit
PUBLIC prn_memo, num_ord, summa_memo, memo_GrName, name_liz_plat, name_dbt, name_crt, name_liz_polu, memo_osnovanie, type_memo, text_memo

PUBLIC dop_name_filial, name_bank, full_name_bank, full_name_bank_yar, full_name_bank_memo, InkahranBIK, InkahranKorSchet, InkahranAcc, liz_num_60322_ADM

dop_name_filial = ''  && ������� ��� ������������ �������
name_bank_1 = '������������ ���� 3454-�'  && �������� �������� �����������
name_bank_2 = '/2 ��� ��������ͻ (��)'  && �������� �������� �����������
name_bank = name_bank_1 + name_bank_2  && �������� �������� �����������
full_name_bank = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && ������ ������������, ���������� ����������� � ������
full_name_bank_yar = '��������� ��� ���������� �������� � ���������� ��� " ��������" (��)  129366 �. ������ ��. ����������� , �. 11'  && ������ ������������, ���������� ����������� � ������, ��� ���������
full_name_bank_memo = ALLTRIM(full_name_bank)

InkahranBIK = '044583934'
InkahranKorSchet = ''
InkahranAcc = '20202810955950000000'

summa_memo = 0.00
num_ord = '001'
type_memo = ''
text_memo = ''

prn_memo = '����� �� �������� ��������� �'  && ����� �� �������� ���������

liz_num_60322_ADM = '60322810055990000001'

liz_num_plat_prk = '91202810755950000103'  && 91202.810.7.5595.0000103

liz_num_polu_prk = '99999810055990000000'  && 99999.810.0.5599.0000000

liz_num_polu_kredit_prk = '99999810055990000000'  && 99999.810.0.5599.0000000
liz_num_polu_debit_prk = '99999810055990000000'  && 99999.810.0.5599.0000000

name_liz_plat = ALLTRIM(full_name_bank)  && '��� "��������" (���)'
name_liz_polu = ALLTRIM(full_name_bank)  && '��� "��������" (���)'

memo_GrName = ''
naznach = ''

dolg_rukovod = ''
fio_boss = ''
fio_boss_key = ''
post_boss = ''

dolg_ispolnit = ''
fio_spez = ''

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

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

RcValue = ""
xx.getinientry(@RcValue, "ConnectSql", "ServerUserId", inifile)
BdLogin = RcValue

* WAIT WINDOW 'bdlogin = ' + ALLTRIM(bdlogin) TIMEOUT 3

RcValue = ""
xx.getinientry(@RcValue, "ConnectSql", "ServerUserPsw", inifile)
BdPsw = RcValue

* WAIT WINDOW 'bdpsw = ' + ALLTRIM(bdpsw) TIMEOUT 3

RcValue = ""
xx.getinientry(@RcValue, "ConnectSql", "OdbcName", inifile)
OdbcSrc = RcValue

* WAIT WINDOW 'odbcsrc = ' + ALLTRIM(odbcsrc) TIMEOUT 3

RcValue = ""
xx.getinientry(@RcValue, "ConnectSql", "QueryTimeOut", inifile)
SqlQueryTimeOut = VAL(RcValue)

* WAIT WINDOW 'SqlQueryTimeOut = ' + ALLTRIM(STR(SqlQueryTimeOut)) TIMEOUT 2

RcValue = ""
xx.getinientry(@RcValue, "ConnectSql", "SqlWaitTime", inifile)
SqlWaitTime = VAL(RcValue)

* WAIT WINDOW 'SqlWaitTime = ' + ALLTRIM(STR(SqlWaitTime)) TIMEOUT 3

*** ��������� ���������� ������, � ������� ������� ���������� ������� ������ ����������

IF SqlQueryTimeOut = 0
  SqlQueryTimeOut = 180
ENDIF

*** WaitTime - ����� � �������� ����� ������� Visual FoxPro ��������� ����������� �� ���������� ���������� SQL. �������� �� ��������� - 15 ������. ������/������.

IF SqlWaitTime = 0
  SqlWaitTime = 15
ENDIF

*   WAIT WINDOW 'ServerName = ' + ALLTRIM(ServerName) TIMEOUT 3
*   WAIT WINDOW 'odbcsrc = ' + ALLTRIM(odbcsrc) TIMEOUT 3
*   WAIT WINDOW 'bdlogin = ' + ALLTRIM(bdlogin) TIMEOUT 3
*   WAIT WINDOW 'bdpsw = ' + ALLTRIM(bdpsw) TIMEOUT 3

IF EMPTY(OdbcSrc) = .T. OR EMPTY(BdLogin) = .T.
  MESSAGEBOX("�������� OdbcName ��� ServerUserId � ����� inkohr_BS.ini ������� �� �����!", 16, "��� ���� ������.")
  RETURN
ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO vxod IN SecurPack
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
IF TYPE("loExcel.ActiveWorkbook") = 'O'
  loExcel.ActiveWorkbook.CLOSE()
ENDIF
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO exit_sql_server IN Vxod_sql
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO Exit IN Prog_sbzhuk
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

RETURN


************************************************************************************ PROCEDURE vxod ************************************************************************

PROCEDURE vxod

* -------------------------------------------------*
DO vxod_sql_server IN Vxod_sql
* ------------------------------------------------- *

* WAIT WINDOW 'f_IdDepartment = ' + ALLTRIM(STR(f_IdDepartment)) &&TIMEOUT 3

Pr_Otladka = .F.  && ������� ������ � ������ ������������ � �������

IF Pr_Otladka = .T.
  sergeizhuk_flag = .T.  && ��� ������������� �� �������� �� �����������, ��������� �� ����������� � ��������� ������
ENDIF

loginok = .F.

OpenDatePRK = OpenDate_

poisk_date = TTOD(OpenDate_)
new_data_odb = TTOD(OpenDate_)

f_CurDate = TTOD(OpenDate_)
f_poisk_date = DTOC(new_data_odb)

data_tsd = DTOT(new_data_odb)
data_tsd_sql = DTOC(TTOD(data_tsd))

bar_office = 1

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
DO rasvert IN Prog_sbzhuk
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

DO start IN SecurPack_menu_inkas  &&  ������ ��������� ������������ �������� ���� ��������� - ����� ���������������
DO start IN SecurPack_menu_vozm  &&  ������ ��������� ������������ �������� ���� ��������� - ����� � �����������
DO start IN SecurPack_menu_atm  &&  ������ ��������� ������������ �������� ���� ��������� - ����� � ��������� ���
DO start IN SecurPack_menu_mkb_vsp  &&  ������ ��������� ������������ �������� ���� ��������� - ����� ����������� ��� - ���
DO start IN SecurPack_menu_vozm_mkb_vsp  &&  ������ ��������� ������������ �������� ���� ��������� - ����� � ����������� ��� - ���

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

=LoadKeyboard('ENG')  && �������� ���������� ��������� ����������

* WAIT WINDOW 'simvol_klav = ' + ALLTRIM(simvol_klav) TIMEOUT 3

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

PUBLIC Kasper_IdRecShift, Kasper_IdCurShift, Kasper_IdShift, Kasper_IdUser

Kasper_IdShift = f_IdShift  && ������������ ��� ������ ������ � ������� BudgetHistoryPRK, RouteHistory � IdShift - ��� Id ���������� ����� �������� ����� �� ����������� Shift, ������� �������� ����� WHERE IdDepartment = 1 AND Status = 1 AND CloseDate IS NULL
Kasper_IdRecShift = f_IdShift  && ������������ ��� ������ ������ � ������� BudgetHistoryPRK � IdRecShift - ��� Id ���������� ����� �������� ����� �� ����������� Shift
Kasper_IdCurShift = f_IdShift  && ������������ ��� ������ ������ � ������� BudgetHistoryPRK � IdCurShift - ��� Id ���������� ����� �������� ����� �� ����������� Shift

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

connect_sql = .F.

* f_IdPost = 2  && �������� �� ���������� ����������� � �������� ����� ���
* f_IdPost = 52  && �������� �� ���������� ����������� � ����� ���������
* f_IdPost = 20  && � ����� ����������� �������� � ���� ������ ���������� ���������� ���������� ������ ����� - ��������� �����������

DO CASE
  CASE connect_sql = .F.

  TbLogin_Value = '770'
  TbPSW_Value = '770'

* ------------------------------- *
    DO FORM Login
* ------------------------------- *

*  WAIT WINDOW 'f_IdDepartment = ' + ALLTRIM(STR(f_IdDepartment)) TIMEOUT 3
*  WAIT WINDOW 'f_IdPost = ' + ALLTRIM(STR(f_IdPost)) TIMEOUT 3
*  WAIT WINDOW 'f_IdUser = ' + ALLTRIM(STR(f_IdUser)) TIMEOUT 3
*  WAIT WINDOW 'f_MestoCode = ' + ALLTRIM(f_MestoCode) TIMEOUT 3

* loginok = .T.

    IF loginok = .T.

      =WriteLogOwner_Local("loginID_ = " + ALLTRIM(STR(f_IdControler)), "L")

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      OpenDate_ = GetStartShiftPRK()  && DATETIME()

      ShiftDate_ = OpenDate_

* ------------ f_IdShift - ��� Id ���������� ����� �������� ����� �� ����������� Shift, ������� �������� ����� WHERE IdDepartment = 1 AND Status = 1, 2 AND CloseDate IS NULL ------------- *
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
*   WAIT WINDOW 'OpenDate_ = ' + TTOC(OpenDate_) TIMEOUT 5
*   WAIT WINDOW 'f_IdOperDate = ' + ALLTRIM(STR(f_IdOperDate)) + '   f_IdShift = ' + ALLTRIM(STR(f_IdShift)) + '   f_IdShift_1 = ' + ALLTRIM(STR(f_IdShift_1)) TIMEOUT 5
*   WAIT WINDOW 'IdPost_ = ' + ALLTRIM(STR(IdPost_)) + '   IdDepartment_ = ' + ALLTRIM(STR(IdDepartment_)) + '   NameControler_ = ' + ALLTRIM(NameControler_) TIMEOUT 5
*   WAIT WINDOW 'IdControler_ = ' + ALLTRIM(STR(IdControler_)) TIMEOUT 5
* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      Kasper_IdUser = f_IdUser

      DO CASE
        CASE f_UserId = '770' AND f_UserPassword = '770'
          f_IdPost = 71

        CASE f_UserId = '770_2' AND f_UserPassword = '770'
          f_IdPost = 71

      ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      DO CASE
        CASE ALLTRIM(f_MestoCode) == 'KTL'

          name_bank_1 = '��������� ��� ���������� �������� � ���������� ��� "��������" �� �� ������:' + SPACE(1)  && �������� �������� �����������
          name_bank_2 = '115201, �. ������, ��. ������������, �. 8'  && �������� �������� �����������
          f_MestoCodeName = ' - �������� �� ������������'
          full_name_bank_memo = ALLTRIM(full_name_bank)

        CASE ALLTRIM(f_MestoCode) == 'YAR'

          name_bank_1 = '��������� ��� ���������� �������� � ���������� ��� "��������" �� �� ������:' + SPACE(1)  && �������� �������� �����������
          name_bank_2 = '129366, �. ������, ��. �����������, �. 11'  && �������� �������� �����������
          f_MestoCodeName = ' - �������� �� ���������'
          full_name_bank_memo = ALLTRIM(full_name_bank_yar)

      ENDCASE

      name_bank = name_bank_1 + name_bank_2  && �������� �������� �����������
      full_name_bank = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && ������ ������������, ���������� ����������� � ������

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*          SELECT mst.Code, mst.OpOffice
*          FROM Mesto AS mst
*          WHERE mst.IsActive = 1
*          ORDER BY mst.Nn

*        ENDTEXT

*        ret = RunSQL('Scan_Mesto', @ar)

*  * BROWSE WINDOW brows

*        IF USED('Scan_Mesto') AND RECCOUNT() > 0

*          Pr_Use_Scan_Mesto = .T.

*          INDEX ON Code TAG Mesto

*          IF SEEK(f_MestoCode) = .T.
*            full_name_bank = ALLTRIM(OpOffice)
*          ENDIF

*        ENDIF

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      prn_memo = '����� �� �������� ��������� �'  && ����� �� �������� ���������

      DO CASE
        CASE ALLTRIM(f_MestoCode) == 'KTL'

          liz_num_60322_ADM = '60322810055990000001'
          liz_num_plat_prk = '91202810755950000103'  && 91202.810.7.5595.0000103
          liz_num_polu_prk = '99999810055990000000'  && 99999.810.0.5599.0000000
          liz_num_polu_kredit_prk = '99999810055990000000'  && 99999.810.0.5599.0000000
          liz_num_polu_debit_prk = '99999810055990000000'  && 99999.810.0.5599.0000000

        CASE ALLTRIM(f_MestoCode) == 'NAG'

          liz_num_60322_ADM = '60322810055990000001'
          liz_num_plat_prk = '91202810755950000103'  && 91202.810.7.5595.0000103
          liz_num_polu_prk = '99999810055990000000'  && 99999.810.0.5599.0000000
          liz_num_polu_kredit_prk = '99999810055990000000'  && 99999.810.0.5599.0000000
          liz_num_polu_debit_prk = '99999810055990000000'  && 99999.810.0.5599.0000000

        CASE ALLTRIM(f_MestoCode) == 'YAR'

          liz_num_60322_ADM = '60322810055990000001'
          liz_num_plat_prk = '91202810755950000103'  && 91202.810.7.5595.0000103
          liz_num_polu_prk = '99999810055990000000'  && 99999.810.0.5599.0000000
          liz_num_polu_kredit_prk = '99999810055990000000'  && 99999.810.0.5599.0000000
          liz_num_polu_debit_prk = '99999810055990000000'  && 99999.810.0.5599.0000000

        CASE ALLTRIM(f_MestoCode) == 'ELV'

          liz_num_60322_ADM = '60322810055990000001'
          liz_num_plat_prk = '91202810755950000103'  && 91202.810.7.5595.0000103
          liz_num_polu_prk = '99999810055990000000'  && 99999.810.0.5599.0000000
          liz_num_polu_kredit_prk = '99999810055990000000'  && 99999.810.0.5599.0000000
          liz_num_polu_debit_prk = '99999810055990000000'  && 99999.810.0.5599.0000000

      ENDCASE

      name_liz_plat = ALLTRIM(full_name_bank)  && '��� "��������" (���)'
      name_liz_polu = ALLTRIM(full_name_bank)  && '��� "��������" (���)'

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      Value_Time = VAL(LEFT(TIME(), 2))  && ���������� �������� ���� � ������� �������� ����� ���, ������������ ��� �������� ���� ��������� ����: ������� 12 ����� � ����� ���� 12 �����

*    Value_Time = 8  && ������������ ��� �����, ����� ����� ������������ ���� ��������������, ��� ������ ���� ��� � �� ������ ������� ����� ������� ����� �� 14:00 �����.

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* sergeizhuk_flag = .F.

*        IF sergeizhuk_flag = .F.
*          WAIT WINDOW '��������� ������������! �� ��������� � �������� ������� ����!' TIMEOUT 5
*        ENDIF

      DO CASE
        CASE INLIST(f_IdPost, 1, 2, 44, 49, 52, 56, 68, 69, 70, 71) = .T.  && �������� �� ���������� ����������� � �������� �����, �������� Id ���������� �� ������� Post (1, 2, 44, 49, 52, 56, 68, 69, 70, 71)

          bar_office = 1

          DO WHILE .T.

            pr_vsp_mkb = .F.

            DO CASE
              CASE INLIST(f_IdPost, 1, 1) = .T.  && �������� �� ���������� ����������� � �������� �����, �������� Id ���������� �� ������� Post (1, 1)

                tex1 = '�������� � ���� ����� ��������������� � ���' + f_MestoCodeName
                tex2 = '�������� � ���� ����� � ����������� �� ������ � ���' + f_MestoCodeName
                tex3 = '�������� � ���� ����� � ��������� ��� � ���' + f_MestoCodeName
                tex4 = '�������� � ���� ����� ��������������� � ��� - ��� � ���' + f_MestoCodeName
                l_bar = 7
                =popup_9(tex1, tex2, tex3, tex4, l_bar)


              CASE INLIST(f_IdPost, 2, 44, 49, 52, 56, 68, 69, 70, 71) = .T.  && �������� �� ���������� ����������� � �������� �����, �������� Id ���������� �� ������� Post (2, 44, 49, 52, 56, 68, 69, 70, 71)

                tex1 = '�������� � ���� ����� ��������������� � ���' + f_MestoCodeName
                tex2 = '�������� � ���� ����� � ����������� �� ������ � ���' + f_MestoCodeName
                tex3 = '�������� � ���� ����� � ��������� ��� � ���' + f_MestoCodeName
                tex4 = '�������� � ���� ����� ��������������� � ��� - ��� � ���' + f_MestoCodeName
                tex5 = '�������� � ���� ����� � ����������� � ��� - ��� � ���' + f_MestoCodeName
                l_bar = 9
                =popup_big(tex1, tex2, tex3, tex4, tex5, tex6, tex7, tex8, l_bar)


            ENDCASE

            ACTIVATE POPUP vibor BAR bar_office

            RELEASE POPUPS vibor

            IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

              bar_office = IIF(BAR() >= 1, BAR(), 1)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

              name_office = '���'

              OpenDate_ = GetOpShiftPRK()  && DATETIME()

              ShiftDate_ = OpenDate_

              poisk_date = TTOD(OpenDate_)
              new_data_odb = TTOD(OpenDate_)

              data_tsd = DTOT(new_data_odb)
              data_tsd_sql = DTOC(TTOD(data_tsd))

              f_curDate = TTOD(OpenDate_)
              f_poisk_date = DTOC(new_data_odb)

              WAIT WINDOW '���� ������������� ��� - ' + DT(skip_data_odb) + '  ���� �������� ����� � ��� - ' + DT(new_data_odb) + f_MestoCodeName TIMEOUT 2

              title_program_full = SPACE(2) + ALLTRIM(title_program) + '  ���� ������� - ' + DT(skip_data_odb) + '  ���� ����� � ��� - ' + DT(new_data_odb) + f_MestoCodeName + '   ����� ����� = ' + ALLTRIM(STR(f_IdShift))

              _SCREEN.CAPTION = ALLTRIM(title_program_full)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

              DO CASE
                CASE flag_Shift_Blok = 1 AND f_IdPost = 71  && �������� ���������� �� ���� � ���, � ������������ ��������� ������� ����� ��� ��� Id � ������� Post = 71

                  MESSAGEBOX('� ����������� ���� ��� ��� ���������� ����������' + CHR(13) + '������ � ���� ��������� ������ �������� ����� ���', 16, '������ � ������������� �����������')

                CASE flag_Shift_Blok = 1 AND INLIST(f_IdPost, 1, 2, 49, 52, 68, 69, 70)  && �������� ���������� �� ���� � ���, � ������������ ��������� ������, ������� ������, ������-���������, ��������� ��� ��� Id � ������� Post = 1, 2, 49, 52, 68, 69, 70

                  MESSAGEBOX('� ����������� ���� ��� ��� ���������� ����������' + CHR(13) + '������ � ����  - ���������', 16, '���������� � �������� ����� ���')
                  RETURN

              ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

              DO CASE
                CASE bar_office = 1  && ���� ����� ��������������� � ���

                  type_modul_arm = '01'  && ������ ������ ����� �����������
                  pr_vsp_mkb = .F.

                  ACTIVATE MENU SecurPack_prk


                CASE bar_office = 3  && ���� ����� � ����������� �� ������ � ���

                  type_modul_arm = '02'  && ������ ������ ����� � �����������
                  pr_vsp_mkb = .F.

                  ACTIVATE MENU SecurPack_vozm


                CASE bar_office = 5  && ���� ����� � ��������� ��� � ���

                  type_modul_arm = '03'  && ������ ������ ������ ��� ����������
                  pr_vsp_mkb = .F.

                  ACTIVATE MENU SecurPack_atm


                CASE bar_office = 7  && ���� ����� ��������������� � ��� - ��� � ���

                  type_modul_arm = '04'  && ������ ������ ����� ��������������� � ��� - ��� � ���
                  pr_vsp_mkb = .T.

                  ACTIVATE MENU SecurPack_mkb_vsp


                CASE bar_office = 9  && ���� ����� � ����������� � ��� - ��� � ���

                  type_modul_arm = '05'  && ������ ������ ����� � ����������� � ��� - ��� � ���
                  pr_vsp_mkb = .T.

                  ACTIVATE MENU SecurPack_vozm_mkb_vsp


              ENDCASE

              pr_vsp_mkb = .F.

              LOOP

              IF LOWER(PAD()) == 'quit'
                EXIT
              ELSE
                LOOP
              ENDIF

            ELSE
              EXIT
            ENDIF

          ENDDO

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        OTHERWISE

          =err('��������! � ��� ��� ����� ������� ��� ������ � ��� SecurPack')

      ENDCASE

    ENDIF


* ============================================================================================================================================================================================== *


  CASE connect_sql = .T.

*    sergeizhuk_flag = .F.

    Pr_Otladka = .F.

    IdControler_ = 819

    IdPost_ = 71

    IdDepartment_ = 10

    f_PostName = '��������� ������'

    f_IdControler = IdControler_
    f_IdUser = IdControler_
    f_IdPost = IdPost_
    f_IdDepartment = IdDepartment_
    f_IdDepartmentPRK = IdDepartment_

    f_UserName = '��� �.�.'
    NameControler_ = f_UserName

    f_UserId = 780
    f_UserPassword = 780

    f_IdUser = 819

    Kasper_IdUser = f_IdUser

    bar_office = 1

    f_MestoCode = 'KTL'
*    f_MestoCode = 'YAR'

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    OpenDate_ = GetStartShiftPRK()  && DATETIME()

    ShiftDate_ = OpenDate_

* ------------ f_IdShift -  - ��� Id ���������� ����� �������� ����� �� ����������� Shift, ������� �������� ����� WHERE IdDepartment = 1 AND Status = 1, 2 AND CloseDate IS NULL ------------- *

*   WAIT WINDOW 'OpenDate_ = ' + TTOC(OpenDate_) TIMEOUT 3
*   WAIT WINDOW 'ShiftDate_ = ' + TTOC(ShiftDate_) TIMEOUT 3
*   WAIT WINDOW 'f_IdOperDate = ' + ALLTRIM(STR(f_IdOperDate)) + '   f_IdShift = ' + ALLTRIM(STR(f_IdShift)) + '   f_IdShift_1 = ' + ALLTRIM(STR(f_IdShift_1)) TIMEOUT 5

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    DO CASE
      CASE ALLTRIM(f_MestoCode) == 'KTL'

        name_bank_1 = '��������� ��� ���������� �������� � ���������� ��� "��������" �� �� ������:' + SPACE(1)  && �������� �������� �����������
        name_bank_2 = '115201, �. ������, ��. ������������, �. 8'  && �������� �������� �����������
        f_MestoCodeName = ' - �������� �� ������������'
        full_name_bank_memo = ALLTRIM(full_name_bank)

      CASE ALLTRIM(f_MestoCode) == 'YAR'

        name_bank_1 = '��������� ��� ���������� �������� � ���������� ��� "��������" �� �� ������:' + SPACE(1)  && �������� �������� �����������
        name_bank_2 = '129366, �. ������, ��. �����������, �. 11'  && �������� �������� �����������
        f_MestoCodeName = ' - �������� �� ���������'
        full_name_bank_memo = ALLTRIM(full_name_bank_yar)

      CASE ALLTRIM(f_MestoCode) == 'ELV'

        name_bank_1 = '��������� ��� ���������� �������� � ���������� ��� "��������" �� �� ������:' + SPACE(1)  && �������� �������� �����������
        name_bank_2 = '/2 ��� ��������ͻ ��'  && �������� �������� �����������
        f_MestoCodeName = ' - �������� �� �����������'
        full_name_bank_memo = ALLTRIM(full_name_bank)

    ENDCASE

    name_bank = name_bank_1 + name_bank_2  && �������� �������� �����������
    full_name_bank = ALLTRIM(ALLTRIM(dop_name_filial) + SPACE(1) + ALLTRIM(name_bank))  && ������ ������������, ���������� ����������� � ������

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

    SELECT mst.Code, mst.OpOffice
    FROM Mesto AS mst
    WHERE mst.IsActive = 1 --AND mst.Code = 'KTL'
    ORDER BY mst.Nn

    ENDTEXT

    ret = RunSQL('Scan_Mesto', @ar)

    IF USED('Scan_Mesto') AND RECCOUNT() > 0
      Pr_Use_Scan_Mesto = .T.
      INDEX on code TAG mesto
      IF SEEK(f_MestoCode)
        full_name_bank=ALLTRIM(OpOffice)
      endif
    ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    DO WHILE .T.

      pr_vsp_mkb = .F.

      tex1 = '�������� � ���� ����� ��������������� � ���' + f_MestoCodeName
      tex2 = '�������� � ���� ����� � ����������� �� ������ � ���' + f_MestoCodeName
      tex3 = '�������� � ���� ����� � ��������� ��� � ���' + f_MestoCodeName
      tex4 = '�������� � ���� ����� ��������������� � ��� - ��� � ���' + f_MestoCodeName
      tex5 = '�������� � ���� ����� � ����������� � ��� - ��� � ���' + f_MestoCodeName
      l_bar = 9
      =popup_big(tex1, tex2, tex3, tex4, tex5, tex6, tex7, tex8, l_bar)

      ACTIVATE POPUP vibor BAR bar_office

      RELEASE POPUPS vibor

      IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

        bar_office = IIF(BAR() >= 1, BAR(), 1)

        name_office = '���'

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        Value_Time = VAL(LEFT(TIME(), 2))  && ���������� �������� ���� � ������� �������� ����� ���, ������������ ��� �������� ���� ��������� ����: ������� 12 ����� � ����� ���� 12 �����

*          Value_Time = 9  && ������������ ��� �����, ����� ����� ������������ ���� ��������������, ��� ������ ���� ��� � �� ������ ������� ����� ������� ����� �� 11:00 �����.

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        OpenDate_ = GetOpShiftPRK()  && DATETIME()

* WAIT WINDOW 'f_IdOperDate = ' + ALLTRIM(STR(f_IdOperDate)) + '   f_IdShift = ' + ALLTRIM(STR(f_IdShift)) + '   f_IdShift_1 = ' + ALLTRIM(STR(f_IdShift_1)) TIMEOUT 5

        ShiftDate_ = OpenDate_

        poisk_date = TTOD(OpenDate_)
        new_data_odb = TTOD(OpenDate_)

* WAIT WINDOW 'new_data_odb = ' + DTOC(new_data_odb) TIMEOUT 3

        data_tsd = DTOT(new_data_odb)
        data_tsd_sql = DTOC(TTOD(data_tsd))

        f_curDate = TTOD(OpenDate_)
        f_poisk_date = DTOC(new_data_odb)

        WAIT WINDOW '���� ������������� ��� - ' + DT(skip_data_odb) + '  ���� �������� ����� � ��� - ' + DT(new_data_odb) + f_MestoCodeName TIMEOUT 2

        title_program_full = SPACE(2) + ALLTRIM(title_program) + '  ���� ������� - ' + DT(skip_data_odb) + '  ���� ����� � ��� - ' + DT(new_data_odb) + f_MestoCodeName + '   ����� ����� = ' + ALLTRIM(STR(f_IdShift))

        _SCREEN.CAPTION = ALLTRIM(title_program_full)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        DO CASE
          CASE flag_Shift_Blok = 1 AND f_IdPost = 71  && �������� ���������� �� ���� � ���, � ������������ ��������� ������� ����� ��� ��� Id � ������� Post = 71

            MESSAGEBOX('� ����������� ���� ��� ��� ���������� ����������' + CHR(13) + '������ � ���� ��������� ������ �������� ����� ���', 16, '������ � ������������� �����������')

          CASE flag_Shift_Blok = 1 AND INLIST(f_IdPost, 1, 2, 49, 52, 68, 69, 70)  && �������� ���������� �� ���� � ���, � ������������ ��������� ������, ������� ������, ������-���������, ��������� ��� ��� Id � ������� Post = 1, 2, 49, 52, 68, 69, 70

            MESSAGEBOX('� ����������� ���� ��� ��� ���������� ����������' + CHR(13) + '������ � ����  - ���������', 16, '���������� � �������� ����� ���')
            RETURN

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        DO CASE
          CASE bar_office = 1  && ���� ����� ��������������� � ���

            type_modul_arm = '01'  && ������ ������ ����� �����������
            pr_vsp_mkb = .F.

            ACTIVATE MENU SecurPack_prk


          CASE bar_office = 3  && ���� ����� � ����������� �� ������ � ���

            type_modul_arm = '02'  && ������ ������ ����� � �����������
            pr_vsp_mkb = .F.

            ACTIVATE MENU SecurPack_vozm


          CASE bar_office = 5  && ���� ����� � ��������� ��� � ���

            type_modul_arm = '03'  && ������ ������ ������ ��� �����������
            pr_vsp_mkb = .F.

            ACTIVATE MENU SecurPack_atm


          CASE bar_office = 7  && ���� ����� ��������������� � ��� - ��� � ���

            type_modul_arm = '04'  && ������ ������ ����� ��������������� � ��� - ��� � ���
            pr_vsp_mkb = .T.

            ACTIVATE MENU SecurPack_mkb_vsp


          CASE bar_office = 9  && ���� ����� � ����������� � ��� - ��� � ���

            type_modul_arm = '05'  && ������ ������ ����� � ����������� � ��� - ��� � ���
            pr_vsp_mkb = .T.

            ACTIVATE MENU SecurPack_vozm_mkb_vsp


        ENDCASE

        pr_vsp_mkb = .F.

        LOOP

        IF LOWER(PAD()) == 'quit'
          EXIT
        ELSE
          LOOP
        ENDIF

      ELSE
        EXIT
      ENDIF

    ENDDO

ENDCASE

RETURN


***************************************************************************************************************************************************************************************************************************





