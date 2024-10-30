**************************************************
*** ������������ ����������� ��������� ������ ***
**************************************************

******************************************************************* PROCEDURE start ********************************************************************

PROCEDURE start
sel = SELECT()
popup_ima=LOWER(POPUP())
prompt_ima=LOWER(PROMPT())
bar_num=BAR()
HIDE POPUP (popup_ima)

tex1 = '�������� ������ �� ���������� ���������� ��������� ������ 40817 � 40820'
tex2 = '�������� ������ �� ���������� ���������� ��������� ������ 40817 � 40820'
l_bar = 3
=popup_9(tex1, tex2, tex3, tex4, l_bar)
ACTIVATE POPUP vibor
RELEASE POPUPS vibor

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  put_select = BAR()

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  DO CASE
    CASE put_select = 1  && �������� ������ �� ���������� ���������� ��������� ������ 40817 � 40820

      =soob('��������! �������� ����� �������, � ������� ����� ��������� ��������� ���� ������� ... ')

      ACTIVATE POPUP zarpl_proekt

      IF NOT LASTKEY() = 27

        IF INLIST(LASTKEY(), 4, 19)
          LOOP
        ELSE
          DO zarpl_proekt IN Visa
        ENDIF

      ELSE

        num_proekt = '00'  && ����� ����������� �������

        =soob('��������! ��� ����� - ������� ���������� ��������� ������ ������� - ��� ������� "00"')

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

    CASE put_select = 3  && �������� ������ �� ���������� ���������� ��������� ������ 40817 � 40820

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

      DO R_key IN Rasch_kl WITH zUCH, zSHET && � ������ ��������� ���������� ������ ����� � ������� �����, ����� ������� 9.

      sel_card_nls = ALLTRIM(kSHET)

      REPLACE card_nls WITH sel_card_nls

      SELECT konvert
    ENDSCAN

    SELECT konvert
    GO TOP

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    BROWSE FIELDS ;
      branch:H = '������',;
      kod_proekt:H = '������',;
      client_b:H = '� �������',;
      fio_rus:H = '������� ��� ��������',;
      card_nls:H = '�������� � ��� Visa',;
      card_acct:H = '�������� � ���',;
      type_card:H = '���',;
      val_card:H = '������',;
      tran_42301:H = '���������� ����',;
      name_card:H = '��� �����' ;
      WINDOWS brows ;
      TITLE ' ���������� ������ �� ��������� ������. ���������� ��������� ������ ���������� - ' + ALLTRIM(STR(colvo_zap)) + ' ����.'

    tim1 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('��������! ������ ������������ ���������� ����� ��� ����������� ������.',WCOLS())

    SELECT konvert
    GO TOP

    SET TEXTMERGE ON
    SET TEXTMERGE NOSHOW

    _text = FCREATE((put_tmp) + ('konvert.txt'))

    IF _text < 1

      =err('��������! ������ ��� �������� ����� ' + (put_tmp) + ('konvert.txt'))

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

        DO R_key IN Rasch_kl WITH zUCH, zSHET && � ������ ��������� ���������� ������ ����� � ������� �����, ����� ������� 9.

        REPLACE card_nls WITH ALLTRIM(kSHET)

        SELECT new_konvert
      ENDSCAN

      IF USED('new_konvert')
        SELECT new_konvert
        USE
      ENDIF

    ELSE
      =err('����������� ��� ������ ���� �� ���� ' + (put_dbf) + ('konvert.dbf') + ' �� ������ ... ')
    ENDIF

    @ WROWS()/3,3 SAY PADC('���� ' + (put_dbf) + ('konvert.dbf') + ' ������� ������� ... ',WCOLS())
    =INKEY(4)

    @ WROWS()/3,3 SAY PADC('���� ' + (put_tmp) + ('konvert.txt') + ' ������� ������� ... ',WCOLS())
    =INKEY(4)

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('������������ ���������� ����� ��� ����������� ������ ������� ���������.' + ' (����� = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' ���.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('��������! ������������� ���� ������ �� ���������� ...')
  ENDIF

ENDIF

RETURN


***********************************************************************************************************************************************************


