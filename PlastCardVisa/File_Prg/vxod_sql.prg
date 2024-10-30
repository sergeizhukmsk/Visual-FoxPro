**********************************************************************
*** ��������� ��������� �� ������ � ���� SQL Server � ������� � ��� ***
**********************************************************************

****************************************************** PROCEDURE vxod_sql_server_loc_2005 ****************************************************************

PROCEDURE vxod_sql_server_loc_2005

SET ESCAPE ON

I = 0

FOR I = 1 TO 10

  sql_num = 0

  DO CASE
    CASE sergeizhuk_flag = .T.

      sql_num = SQLCONNECT('VisaCardLoc2005_' + ALLTRIM(num_branch), 's_zhuk', 'edcxsw45')

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE sergeizhuk_flag = .F.

      DO CASE
        CASE ALLTRIM(num_branch) == '01'  && �������� ����

          sql_num = SQLCONNECT('VisaCardLoc2005_01', 'visacard_01', '123456')

        CASE ALLTRIM(num_branch) == '02'  && �����

          sql_num = SQLCONNECT('VisaCardLoc2005_02', 'visacard_02', '123456')

        CASE ALLTRIM(num_branch) == '03'  && ������

          sql_num = SQLCONNECT('VisaCardLoc2005_03', 'visacard_03', '123456')

        CASE ALLTRIM(num_branch) == '04'  && �������

          sql_num = SQLCONNECT('VisaCardLoc2005_04', 'visacard_04', '123456')

        CASE ALLTRIM(num_branch) == '05'  && ������

          sql_num = SQLCONNECT('VisaCardLoc2005_05', 'visacard_05', '123456')

        CASE ALLTRIM(num_branch) == '06'  && ������

          sql_num = SQLCONNECT('VisaCardLoc2005_06', 'visacard_06', '123456')

        CASE ALLTRIM(num_branch) == '07'  && �����

          sql_num = SQLCONNECT('VisaCardLoc2005_07', 'visacard_07', '123456')

        CASE ALLTRIM(num_branch) == '08'  && �����������

          sql_num = SQLCONNECT('VisaCardLoc2005_08', 'visacard_08', '123456')

        CASE ALLTRIM(num_branch) == '09'  && ����

          sql_num = SQLCONNECT('VisaCardLoc2005_09', 'visacard_09', '123456')

        CASE ALLTRIM(num_branch) == '10'  && �����������

          sql_num = SQLCONNECT('VisaCardLoc2005_10', 'visacard_10', '123456')

        CASE ALLTRIM(num_branch) == '11'  && ���������

          sql_num = SQLCONNECT('VisaCardLoc2005_11', 'visacard_11', '123456')

        CASE ALLTRIM(num_branch) == '12'  && �������

          sql_num = SQLCONNECT('VisaCardLoc2005_12', 'visacard_12', '123456')

      ENDCASE

  ENDCASE

* WAIT WINDOW '����� ������������ ���������� sql_num = ' + ALLTRIM(STR(sql_num)) TIMEOUT 2

  IF sql_num > 0
    EXIT
  ELSE
    =INKEY(2)
  ENDIF

ENDFOR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF sql_num > 0

  IF SQLGETPROP(sql_num, 'Asynchronous') = .T.
    =SQLSETPROP(sql_num, 'Asynchronous', .F.)
  ENDIF

  IF SQLGETPROP(sql_num, 'BatchMode') = .T.
    =SQLSETPROP(sql_num, 'BatchMode', .F.)
  ENDIF

  ima_sql_server = '(��������� ������ SQL Server 2005)'

  WAIT WINDOW TIMEOUT 2 '��������! ��������� ����������� � ���������� SQL Server 2005'

ELSE
  =err('��������! ����������� � ���������� SQL Server 2005 �� ��������� ... ')
ENDIF

SET ESCAPE OFF

RETURN


******************************************************** PROCEDURE vxod_sql_server_net_2005 *************************************************************

PROCEDURE vxod_sql_server_net_2005

SET ESCAPE ON

I = 0

FOR I = 1 TO 10

  sql_num = 0

  DO CASE
    CASE sergeizhuk_flag = .T.

      sql_num = SQLCONNECT('VisaCardNet2005_' + ALLTRIM(num_branch), 's_zhuk', 'edcxsw45')

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE sergeizhuk_flag = .F.

      DO CASE
        CASE ALLTRIM(num_branch) == '01'  && �������� ����

          sql_num = SQLCONNECT('VisaCardNet2005_01', 'visacard_01', '123456')

        CASE ALLTRIM(num_branch) == '02'  && �����

          sql_num = SQLCONNECT('VisaCardNet2005_02', 'visacard_02', '123456')

        CASE ALLTRIM(num_branch) == '03'  && ������

          sql_num = SQLCONNECT('VisaCardNet2005_03', 'visacard_03', '123456')

        CASE ALLTRIM(num_branch) == '04'  && �������

          sql_num = SQLCONNECT('VisaCardNet2005_04', 'visacard_04', '123456')

        CASE ALLTRIM(num_branch) == '05'  && ������

          sql_num = SQLCONNECT('VisaCardNet2005_05', 'visacard_05', '123456')

        CASE ALLTRIM(num_branch) == '06'  && ������

          sql_num = SQLCONNECT('VisaCardNet2005_06', 'visacard_06', '123456')

        CASE ALLTRIM(num_branch) == '07'  && �����

          sql_num = SQLCONNECT('VisaCardNet2005_07', 'visacard_07', '123456')

        CASE ALLTRIM(num_branch) == '08'  && �����������

          sql_num = SQLCONNECT('VisaCardNet2005_08', 'visacard_08', '123456')

        CASE ALLTRIM(num_branch) == '09'  && ����

          sql_num = SQLCONNECT('VisaCardNet2005_09', 'visacard_09', '123456')

        CASE ALLTRIM(num_branch) == '10'  && �����������

          sql_num = SQLCONNECT('VisaCardNet2005_10', 'visacard_10', '123456')

        CASE ALLTRIM(num_branch) == '11'  && ���������

          sql_num = SQLCONNECT('VisaCardNet2005_11', 'visacard_11', '123456')

        CASE ALLTRIM(num_branch) == '12'  && �������

          sql_num = SQLCONNECT('VisaCardNet2005_12', 'visacard_12', '123456')

      ENDCASE

  ENDCASE

* WAIT WINDOW '����� ������������ ���������� sql_num = ' + ALLTRIM(STR(sql_num)) TIMEOUT 2

  IF sql_num > 0
    EXIT
  ELSE
    =INKEY(2)
  ENDIF

ENDFOR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF sql_num > 0

  IF SQLGETPROP(sql_num, 'Asynchronous') = .T.
    =SQLSETPROP(sql_num, 'Asynchronous', .F.)
  ENDIF

  IF SQLGETPROP(sql_num, 'BatchMode') = .T.
    =SQLSETPROP(sql_num, 'BatchMode', .F.)
  ENDIF

  ima_sql_server = '(������� ������ SQL Server 2005)'

  WAIT WINDOW TIMEOUT 2 '��������! ��������� ����������� � �������� SQL Server 2005'

ELSE
  =err('��������! ����������� � �������� SQL Server 2005 �� ��������� ... ')
ENDIF

SET ESCAPE OFF

RETURN


****************************************************** PROCEDURE vxod_sql_server_loc_2008 ****************************************************************

PROCEDURE vxod_sql_server_loc_2008

SET ESCAPE ON

I = 0

FOR I = 1 TO 10

  sql_num = 0

  DO CASE
    CASE sergeizhuk_flag = .T.

      sql_num = SQLCONNECT('VisaCardLoc2008_' + ALLTRIM(num_branch), 's_zhuk', 'edcxsw45')

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE sergeizhuk_flag = .F.

      DO CASE
        CASE ALLTRIM(num_branch) == '01'  && �������� ����

          sql_num = SQLCONNECT('VisaCardLoc2008_01', 'visacard_01', '12345')

        CASE ALLTRIM(num_branch) == '02'  && �����

          sql_num = SQLCONNECT('VisaCardLoc2008_02', 'visacard_02', '12345')

        CASE ALLTRIM(num_branch) == '03'  && ������

          sql_num = SQLCONNECT('VisaCardLoc2008_03', 'visacard_03', '12345')

        CASE ALLTRIM(num_branch) == '04'  && �������

          sql_num = SQLCONNECT('VisaCardLoc2008_04', 'visacard_04', '12345')

        CASE ALLTRIM(num_branch) == '05'  && ������

          sql_num = SQLCONNECT('VisaCardLoc2008_05', 'visacard_05', '12345')

        CASE ALLTRIM(num_branch) == '06'  && ������

          sql_num = SQLCONNECT('VisaCardLoc2008_06', 'visacard_06', '12345')

        CASE ALLTRIM(num_branch) == '07'  && �����

          sql_num = SQLCONNECT('VisaCardLoc2008_07', 'visacard_07', '12345')

        CASE ALLTRIM(num_branch) == '08'  && �����������

          sql_num = SQLCONNECT('VisaCardLoc2008_08', 'visacard_08', '12345')

        CASE ALLTRIM(num_branch) == '09'  && ����

          sql_num = SQLCONNECT('VisaCardLoc2008_09', 'visacard_09', '12345')

        CASE ALLTRIM(num_branch) == '10'  && �����������

          sql_num = SQLCONNECT('VisaCardLoc2008_10', 'visacard_10', '12345')

        CASE ALLTRIM(num_branch) == '11'  && ���������

          sql_num = SQLCONNECT('VisaCardLoc2008_11', 'visacard_11', '12345')

        CASE ALLTRIM(num_branch) == '12'  && �������

          sql_num = SQLCONNECT('VisaCardLoc2008_12', 'visacard_12', '12345')

      ENDCASE

  ENDCASE

* WAIT WINDOW '����� ������������ ���������� sql_num = ' + ALLTRIM(STR(sql_num)) TIMEOUT 2

  IF sql_num > 0
    EXIT
  ELSE
    =INKEY(2)
  ENDIF

ENDFOR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF sql_num > 0

  IF SQLGETPROP(sql_num, 'Asynchronous') = .T.
    =SQLSETPROP(sql_num, 'Asynchronous', .F.)
  ENDIF

  IF SQLGETPROP(sql_num, 'BatchMode') = .T.
    =SQLSETPROP(sql_num, 'BatchMode', .F.)
  ENDIF

  ima_sql_server = '(��������� ������ SQL Server 2008)'

  WAIT WINDOW TIMEOUT 2 '��������! ��������� ����������� � ���������� SQL Server 2008'

ELSE
  =err('��������! ����������� � ���������� SQL Server 2008 �� ��������� ... ')
ENDIF

SET ESCAPE OFF

RETURN


*********************************************************** PROCEDURE export_data_sql_avt_loc_2008 *******************************************************

PROCEDURE export_data_sql_avt_loc_2008  &&  ������� ���������� ������ � ������ ������ �� ���� ������ �� "���-����" � ���� SQL Server 2008

SET ESCAPE ON

=soob('��������! �������� ������ �� ��������� SQL Server 2008 ... ')

sql_num = 0
colvo_zap = 0

ima_sql_server = '(��������� ������ SQL Server 2008)'

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

max_data_date_prov = new_data_odb - 3

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF connect_sql = .T.

  I = 0

  FOR I = 1 TO 10

    DO CASE
      CASE ALLTRIM(num_branch) == '01'  && �������� ����

        sql_num = SQLCONNECT('VisaCardLoc2008_01', 'visacard_01', '123456')

      CASE ALLTRIM(num_branch) == '02'  && �����

        sql_num = SQLCONNECT('VisaCardLoc2008_02', 'visacard_02', '123456')

      CASE ALLTRIM(num_branch) == '03'  && ������

        sql_num = SQLCONNECT('VisaCardLoc2008_03', 'visacard_03', '123456')

      CASE ALLTRIM(num_branch) == '04'  && �������

        sql_num = SQLCONNECT('VisaCardLoc2008_04', 'visacard_04', '123456')

      CASE ALLTRIM(num_branch) == '05'  && ������

        sql_num = SQLCONNECT('VisaCardLoc2008_05', 'visacard_05', '123456')

      CASE ALLTRIM(num_branch) == '06'  && ������

        sql_num = SQLCONNECT('VisaCardLoc2008_06', 'visacard_06', '123456')

      CASE ALLTRIM(num_branch) == '07'  && �����

        sql_num = SQLCONNECT('VisaCardLoc2008_07', 'visacard_07', '123456')

      CASE ALLTRIM(num_branch) == '08'  && �����������

        sql_num = SQLCONNECT('VisaCardLoc2008_08', 'visacard_08', '123456')

      CASE ALLTRIM(num_branch) == '09'  && ����

        sql_num = SQLCONNECT('VisaCardLoc2008_09', 'visacard_09', '123456')

      CASE ALLTRIM(num_branch) == '10'  && �����������

        sql_num = SQLCONNECT('VisaCardLoc2008_10', 'visacard_10', '123456')

      CASE ALLTRIM(num_branch) == '11'  && ���������

        sql_num = SQLCONNECT('VisaCardLoc2008_11', 'visacard_11', '123456')

      CASE ALLTRIM(num_branch) == '12'  && �������

        sql_num = SQLCONNECT('VisaCardLoc2008_12', 'visacard_12', '123456')

      OTHERWISE

        RETURN

    ENDCASE

    IF sql_num > 0
      EXIT
    ELSE
      =INKEY(5)
    ENDIF

  ENDFOR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF sql_num > 0

    WAIT WINDOW '����� ������������ ���������� sql_num = ' + ALLTRIM(STR(sql_num)) TIMEOUT 2

    IF SQLGETPROP(sql_num, 'Asynchronous') = .T.
      =SQLSETPROP(sql_num, 'Asynchronous', .F.)
    ENDIF

    IF SQLGETPROP(sql_num, 'BatchMode') = .T.
      =SQLSETPROP(sql_num, 'BatchMode', .F.)
    ENDIF

    export_data = SQLEXEC(sql_num,"EXECUTE select_docum_slip_max_data","max_data")

    DO WHILE SQLMORERESULTS(sql_num) < 2
    ENDDO

* BROWSE WINDOW brows

    WAIT WINDOW '��������� ���������� �������� export_data = ' + ALLTRIM(STR(export_data)) TIMEOUT 3

    IF USED('max_data')

      SELECT max_data

      COUNT TO colvo_zap
      GO TOP

    ELSE

      =err('��������! ������� � ����������� ������� "max_data" �� ������� ... ')

      colvo_zap = 0

    ENDIF

    IF colvo_zap <> 0

      SELECT max_data
      GO TOP

*  BROWSE WINDOW brows

      max_data_date_prov = IIF(ISNULL(max_data.date_prov) = .F., TTOD(max_data.date_prov), CTOD('01.01.' + ALLTRIM(STR(YEAR(new_data_odb)))))

* WAIT WINDOW 'max_data_date_prov = ' + DTOC(max_data_date_prov) TIMEOUT 4

      IF EMPTY(max_data_date_prov) = .T.
        max_data_date_prov = new_data_odb - 5
      ENDIF

      colvo_den_export = new_data_odb - max_data_date_prov

      IF colvo_den_export = 0
        colvo_den_export = 5
      ENDIF

    ELSE
      colvo_den_export = 30
      max_data_date_prov = new_data_odb - colvo_den_export
    ENDIF

    WAIT WINDOW 'colvo_den_export = ' + ALLTRIM(STR(colvo_den_export)) + ;
      '  new_data_odb = ' + DTOC(new_data_odb) + '  max_data_date_prov = ' + DTOC(max_data_date_prov) TIMEOUT 5

    scan_date_home = new_data_odb - colvo_den_export
    scan_date_end = new_data_odb

    DO export_data_table_filial IN Vxod_sql

    DO exit_sql_server IN Visa

  ELSE
    =err('��������! ����������� � ���������� SQL Server 2008 �� ��������� ... ')
  ENDIF

ELSE
  =err('��������! � ��� ��� ���������� ��� ���������� �������� ������ � ��������� SQL Server 2008 ... ')
ENDIF

SET ESCAPE OFF

RETURN


*********************************************************** PROCEDURE export_data_sql_avt_loc_2005 *******************************************************

PROCEDURE export_data_sql_avt_loc_2005  &&  ������� ���������� ������ � ������ ������ �� ���� ������ �� "���-����" � ���� SQL Server 2005

SET ESCAPE ON

=soob('��������! �������� ������ �� ��������� SQL Server 2005 ... ')

sql_num = 0
colvo_zap = 0

ima_sql_server = '(��������� ������ SQL Server 2005)'

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

max_data_date_prov = new_data_odb - 3

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF connect_sql = .T.

  I = 0

  FOR I = 1 TO 10

    DO CASE
      CASE ALLTRIM(num_branch) == '01'  && �������� ����

        sql_num = SQLCONNECT('VisaCardLoc2005_01', 'visacard_01', '123456')

      CASE ALLTRIM(num_branch) == '02'  && �����

        sql_num = SQLCONNECT('VisaCardLoc2005_02', 'visacard_02', '123456')

      CASE ALLTRIM(num_branch) == '03'  && ������

        sql_num = SQLCONNECT('VisaCardLoc2005_03', 'visacard_03', '123456')

      CASE ALLTRIM(num_branch) == '04'  && �������

        sql_num = SQLCONNECT('VisaCardLoc2005_04', 'visacard_04', '123456')

      CASE ALLTRIM(num_branch) == '05'  && ������

        sql_num = SQLCONNECT('VisaCardLoc2005_05', 'visacard_05', '123456')

      CASE ALLTRIM(num_branch) == '06'  && ������

        sql_num = SQLCONNECT('VisaCardLoc2005_06', 'visacard_06', '123456')

      CASE ALLTRIM(num_branch) == '07'  && �����

        sql_num = SQLCONNECT('VisaCardLoc2005_07', 'visacard_07', '123456')

      CASE ALLTRIM(num_branch) == '08'  && �����������

        sql_num = SQLCONNECT('VisaCardLoc2005_08', 'visacard_08', '123456')

      CASE ALLTRIM(num_branch) == '09'  && ����

        sql_num = SQLCONNECT('VisaCardLoc2005_09', 'visacard_09', '123456')

      CASE ALLTRIM(num_branch) == '10'  && �����������

        sql_num = SQLCONNECT('VisaCardLoc2005_10', 'visacard_10', '123456')

      CASE ALLTRIM(num_branch) == '11'  && ���������

        sql_num = SQLCONNECT('VisaCardLoc2005_11', 'visacard_11', '123456')

      CASE ALLTRIM(num_branch) == '12'  && �������

        sql_num = SQLCONNECT('VisaCardLoc2005_12', 'visacard_12', '123456')

    ENDCASE

    IF sql_num > 0
      EXIT
    ELSE
      =INKEY(5)
    ENDIF

  ENDFOR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF sql_num > 0

    WAIT WINDOW '����� ������������ ���������� sql_num = ' + ALLTRIM(STR(sql_num)) TIMEOUT 2

    IF SQLGETPROP(sql_num, 'Asynchronous') = .T.
      =SQLSETPROP(sql_num, 'Asynchronous', .F.)
    ENDIF

    IF SQLGETPROP(sql_num, 'BatchMode') = .T.
      =SQLSETPROP(sql_num, 'BatchMode', .F.)
    ENDIF

    export_data = SQLEXEC(sql_num,"EXECUTE select_docum_slip_max_data","max_data")

    DO WHILE SQLMORERESULTS(sql_num) < 2
    ENDDO

    WAIT WINDOW '��������� ���������� �������� export_data = ' + ALLTRIM(STR(export_data)) TIMEOUT 3

* BROWSE WINDOW brows

    IF USED('max_data')

      SELECT max_data

      COUNT TO colvo_zap
      GO TOP

    ELSE

      =err('��������! ������� � ����������� ������� "max_data" �� ������� ... ')

      colvo_zap = 0

    ENDIF

    IF colvo_zap <> 0

      SELECT max_data
      GO TOP

*  BROWSE WINDOW brows

      max_data_date_prov = IIF(ISNULL(max_data.date_prov) = .F., TTOD(max_data.date_prov), CTOD('01.01.' + ALLTRIM(STR(YEAR(new_data_odb)))))

* WAIT WINDOW 'max_data_date_prov = ' + DTOC(max_data_date_prov) TIMEOUT 4

      IF EMPTY(max_data_date_prov) = .T.
        max_data_date_prov = new_data_odb - 5
      ENDIF

      colvo_den_export = new_data_odb - max_data_date_prov

      IF colvo_den_export = 0
        colvo_den_export = 5
      ENDIF

    ELSE
      colvo_den_export = 30
      max_data_date_prov = new_data_odb - colvo_den_export
    ENDIF

    WAIT WINDOW 'colvo_den_export = ' + ALLTRIM(STR(colvo_den_export)) + ;
      '  new_data_odb = ' + DTOC(new_data_odb) + '  max_data_date_prov = ' + DTOC(max_data_date_prov) TIMEOUT 5

    scan_date_home = new_data_odb - colvo_den_export
    scan_date_end = new_data_odb

    DO export_data_table_filial IN Vxod_sql

    DO exit_sql_server IN Visa

  ELSE
    =err('��������! ����������� � ���������� SQL Server 2005 �� ��������� ... ')
  ENDIF

ELSE
  =err('��������! � ��� ��� ���������� ��� ���������� �������� ������ � ��������� SQL Server 2005 ... ')
ENDIF

SET ESCAPE OFF

RETURN


*********************************************************** PROCEDURE export_data_sql_avt_net_2005 *******************************************************

PROCEDURE export_data_sql_avt_net_2005  &&  ������� ���������� ������ � ������ ������ �� ���� ������ �� "���-����" � ���� SQL Server 2005

SET ESCAPE ON

=soob('��������! �������� ������ �� ������� SQL Server 2005 ... ')

sql_num = 0
colvo_zap = 0

ima_sql_server = '(������� ������ SQL Server 2005)'

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

max_data_date_prov = new_data_odb - 3

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF connect_sql = .T.

  I = 0

  FOR I = 1 TO 10

    DO CASE
      CASE ALLTRIM(num_branch) == '01'  && �������� ����

        sql_num = SQLCONNECT('VisaCardNet2005_01', 'visacard_01', '123456')

      CASE ALLTRIM(num_branch) == '02'  && �����

        sql_num = SQLCONNECT('VisaCardNet2005_02', 'visacard_02', '123456')

      CASE ALLTRIM(num_branch) == '03'  && ������

        sql_num = SQLCONNECT('VisaCardNet2005_03', 'visacard_03', '123456')

      CASE ALLTRIM(num_branch) == '04'  && �������

        sql_num = SQLCONNECT('VisaCardNet2005_04', 'visacard_04', '123456')

      CASE ALLTRIM(num_branch) == '05'  && ������

        sql_num = SQLCONNECT('VisaCardNet2005_05', 'visacard_05', '123456')

      CASE ALLTRIM(num_branch) == '06'  && ������

        sql_num = SQLCONNECT('VisaCardNet2005_06', 'visacard_06', '123456')

      CASE ALLTRIM(num_branch) == '07'  && �����

        sql_num = SQLCONNECT('VisaCardNet2005_07', 'visacard_07', '123456')

      CASE ALLTRIM(num_branch) == '08'  && �����������

        sql_num = SQLCONNECT('VisaCardNet2005_08', 'visacard_08', '123456')

      CASE ALLTRIM(num_branch) == '09'  && ����

        sql_num = SQLCONNECT('VisaCardNet2005_09', 'visacard_09', '123456')

      CASE ALLTRIM(num_branch) == '10'  && �����������

        sql_num = SQLCONNECT('VisaCardNet2005_10', 'visacard_10', '123456')

      CASE ALLTRIM(num_branch) == '11'  && ���������

        sql_num = SQLCONNECT('VisaCardNet2005_11', 'visacard_11', '123456')

      CASE ALLTRIM(num_branch) == '12'  && �������

        sql_num = SQLCONNECT('VisaCardNet2005_12', 'visacard_12', '123456')

    ENDCASE

    IF sql_num > 0
      EXIT
    ELSE
      =INKEY(5)
    ENDIF

  ENDFOR

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF sql_num > 0

    WAIT WINDOW '����� ������������ ���������� sql_num = ' + ALLTRIM(STR(sql_num)) TIMEOUT 2

    IF SQLGETPROP(sql_num, 'Asynchronous') = .T.
      =SQLSETPROP(sql_num, 'Asynchronous', .F.)
    ENDIF

    IF SQLGETPROP(sql_num, 'BatchMode') = .T.
      =SQLSETPROP(sql_num, 'BatchMode', .F.)
    ENDIF

    export_data = SQLEXEC(sql_num,"EXECUTE select_docum_slip_max_data","max_data")

    DO WHILE SQLMORERESULTS(sql_num) < 2
    ENDDO

    WAIT WINDOW '��������� ���������� �������� export_data = ' + ALLTRIM(STR(export_data)) TIMEOUT 3

    IF USED('max_data')

      SELECT max_data

      COUNT TO colvo_zap
      GO TOP

    ELSE

      =err('��������! ������� � ����������� ������� "max_data" �� ������� ... ')

      colvo_zap = 0

    ENDIF

    IF colvo_zap <> 0

      SELECT max_data
      GO TOP

      max_data_date_prov = IIF(ISNULL(max_data.date_prov) = .F., TTOD(max_data.date_prov), CTOD('01.01.' + ALLTRIM(STR(YEAR(new_data_odb)))))

* WAIT WINDOW 'max_data_date_prov = ' + DTOC(max_data_date_prov) TIMEOUT 5

      IF EMPTY(max_data_date_prov) = .T.
        max_data_date_prov = new_data_odb - 5
      ENDIF

      colvo_den_export = new_data_odb - max_data_date_prov

      IF colvo_den_export = 0
        colvo_den_export = 5
      ENDIF

    ELSE
      colvo_den_export = 30
      max_data_date_prov = new_data_odb - colvo_den_export
    ENDIF

    WAIT WINDOW 'colvo_den_export = ' + ALLTRIM(STR(colvo_den_export)) + ;
      '  new_data_odb = ' + DTOC(new_data_odb) + '  max_data_date_prov = ' + DTOC(max_data_date_prov) TIMEOUT 5

    scan_date_home = new_data_odb - colvo_den_export
    scan_date_end = new_data_odb

    DO export_data_table_filial IN Vxod_sql

    DO exit_sql_server IN Visa

  ELSE
    =err('��������! ����������� � �������� SQL Server 2005 �� ��������� ... ')
  ENDIF

ELSE
  =err('��������! � ��� ��� ���������� ��� ���������� �������� ������ � ������� SQL Server 2005 ... ')
ENDIF

SET ESCAPE OFF

RETURN


*********************************************************** PROCEDURE export_data_table_filial ************************************************************

PROCEDURE export_data_table_filial

DO export_client IN Export_data_sql WITH ima_sql_server

DO export_account IN Export_data_sql WITH ima_sql_server
DO export_acc_del IN Export_data_sql WITH ima_sql_server

DO export_docum_slip IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
DO export_docum_mak IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

DO export_istor_ost IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
DO export_istor_mak IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

DO export_acc_cln_dt IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
DO export_acc_cln_sum IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

DO export_vedom_ost IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

DO export_popol_arx IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
DO export_ostat_upk IN Export_data_sql WITH ima_sql_server

DO export_memordrub IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
DO export_memordval IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

DO export_prixodrub IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
DO export_prixodval IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
DO export_rasxodrub IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
DO export_rasxodval IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

DO export_plateg IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server
DO export_platprn IN Export_data_sql WITH scan_date_home, scan_date_end,  ima_sql_server

RETURN


*********************************************************************************************************************************************************

