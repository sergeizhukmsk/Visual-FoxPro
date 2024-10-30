*******************************************************
*** ��������� ������� ��������� �� ��������� ������ ***
*******************************************************

* ������������ ��� ��������� ������� acc_cln_dt.dbf � acc_cln_sum.dbf
* acc_cln_dt.dbf - ��� ������� �������� ����������� ����� �� ��������� ������ �� ������ ����.
* acc_cln_sum.dbf - ��� ������� �������� ��������� ����������� ����� �� ��������� ������ �� ������ ����� ������� �� ����� �� �������.

* ������������ ���������
* sel_fio - ��� ���������� ������ �������������� ������ �� ����������� �������, �������� �������� �������
* ������� �� ������������� ������, ������� ����������� � sel_fio �� ������ ������ ������������ � �������� account.dbf � acc_del.dbf
* ��� ������ ������������ ����� DO FORM (put_scr) + ('selfio_nls.scx')
* red_data - ���������� �������������� ������
* ��� �������������� ������������ ����� DO FORM (put_scr) + ('red_proz_42301.scx')
* � ��������� read_sum ��������� ��������������� ����������� ������� �������������� ������������ ���� � ��������� ��������
* ����������� �������� ������ � ������ �������


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

    IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

      SELECT account
      SET ORDER TO card_acct

      poisk = SEEK(sel_card_acct)

      DO CASE
        CASE poisk = .T.  && ������ ������� � ������� �������� ����, ��������� ��������� �������������� �������� �������

          ima_table = 'account'

          DO red_data IN Raschet_proz_42301

        CASE poisk = .F.  && ���� ������ �� ������� � ������� ����������� ����, �� ���� � ������� �������� ����

          SELECT acc_del
          SET ORDER TO card_acct

          poisk = SEEK(sel_card_acct)

          IF poisk = .T.  && ���� ������ ������� � ������� �������� ����, �� ��������� ��������� �������������� �������� �������

            ima_table = 'acc_del'

            DO red_data IN Raschet_proz_42301

          ELSE  && ���� ������ �� ������� � ������� �������� ����, �� ������� ��������� �� ������ � ������

            =err('��������! ��������� ���� ������ �� ������ �� � ��������, �� � �������� ������.')

          ENDIF

      ENDCASE

    ELSE
      EXIT
    ENDIF

  ELSE
    =err('��������! �� ���������� ���� ��������� �������� �������� �� ������� ...')
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

  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

    text1 = '���� �p�������� � ����� �����'
    text2 = '��p������ � p���� ����� ������'
    text3 = '�������� ��������� ������'
    l_bar = 5
    =popup_9(text1, text2, text3, text4, l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC
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

  WAIT WINDOW '������ - ' + ALLTRIM(account.fio_rus) + '  ���� ������� - ' + ALLTRIM(account.card_nls) NOWAIT

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

text1 = '����� ������ ��� �������������� �� ���� ������ RUR'
text2 = '����� ������ ��� �������������� �� ���� ������ USD'
text3 = '����� ������ ��� �������������� �� ���� ������ EUR'
l_bar = 5
=popup_9(text1, text2, text3, text4, l_bar)
ACTIVATE POPUP vibor
RELEASE POPUPS vibor

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  DO CASE
    CASE BAR() = 1  && ����� ������ ��� �������������� �� ���� ������ RUR

      sel_val_card = 'RUR'

    CASE BAR() = 3  && ����� ������ ��� �������������� �� ���� ������ USD

      sel_val_card = 'USD'

    CASE BAR() = 5  && ����� ������ ��� �������������� �� ���� ������ EUR

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

    IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

      SELECT DISTINCT data_pak AS data ;
        FROM acc_cln_sum_pol ;
        WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_nls) ;
        INTO CURSOR tov ;
        ORDER BY 1 DESC

      IF _TALLY <> 0

        DO FORM (put_scr) + ('seldata1.scx')

        IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

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
            sum_dobav:H = ' ����� ��������� :':P = '9 999.99',;
            pr:H = ' ������ �� ��������� ������ :',;
            card_nls:H = ' ����� ����� :',;
            fio_rus:H = ' ������� ��� �������� :',;
            val_card:H = ' ��� ����� :',;
            num_dog:H = ' ����� �������� :',;
            data_proz:H = ' ���� ����������� ������� :',;
            kurs:H = ' ����:':P='999.9999',;
            stavka:H = ' ������ :':P='999.99',;
            sum_ostat:H = ' ����� ������� �� ����� :':P='999 999.99',;
            kvartal:H = ' ����� �������� :',;
            mes_kvart:H = ' ����� ������ � �������� :',;
            dt_hom_end:H = ' ������ ���� � �������� :',;
            data_home:H = ' ��������� ���� ������� :',;
            data_end:H = ' �������� ���� ������� :',;
            den_mes:H = ' ���� � ������ :' ;
            TITLE ' �������������� ������ �� ������������ ��������� ' ;
            WINDOW brow_proz

*     sum_rasch:H = ' ����� �� ���� �������� :':P='9 999.99',;
*     sum_rasxod:H = ' ����� �� ������� ����� :':P='9 999.99',;
*     sum_nalog:H = ' ����� ����������� ������ :':P='9 999.99',;
*     sum_viplat:H = ' ����� ��������� � ������� :':P='9 999.99' ;

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
        =soob('��������! � ������� ������������ ��������� �� ����������.')
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

=soob('��������! ���������� ������ ��������� ������� ������������ ������ ������ � ��������� ... ')

tex1 = '� � � � � � � � �  �������� ������ �� ������������ ���������'
tex2 = '� � � � � � � � � �  �� �������� ������ �� ������������ ���������'
l_bar = 3
=popup_9(tex1,tex2,tex3,tex4,l_bar)
ACTIVATE POPUP vibor
RELEASE POPUPS vibor

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  IF BAR() = 1  && � � � � � � � � �  �������� ������ �� ������������ ���������

    DO vxod_monopol IN Visa

    IF error_monopol = .F.

      text1 = '�������� ������ ��������� �� ������ ������� � ��������� ��� �������'
      text2 = '�������� ������ ������� �� ���� �������� � ��������� ��� �������'
      l_bar = 3
      =popup_9(text1, text2, text3, text4, l_bar)
      ACTIVATE POPUP vibor
      RELEASE POPUPS vibor

      IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

        put_vibor = BAR()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

        DO CASE
          CASE put_vibor = 1  && ������� ������ ��������� �� ������ ������� � ��������� ��� �������

            num_mesaz = MONTH(new_data_odb)
            num_god = ALLTRIM(STR(YEAR(new_data_odb)))

            DO vid_kvartal_1 IN Raschet_proz_42301

            vvod_data_end = new_data_odb

            DO FORM (put_scr) + ('sel_data_proz.scx')

            IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

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

                  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

                    tim1 = SECONDS()

                    ACTIVATE WINDOW poisk
                    @ WROWS()/3,3 SAY PADC('��������� ��������, ������ �������� ������ � ������� � ��������� �� ���� ... ',WCOLS())

                    SELECT acc_cln_dt_pol
                    SET ORDER TO card_nls

                    SELECT DISTINCT branch, ref_client, fio_rus, card_nls, data_home, data_end ;
                      FROM acc_cln_dt_pol ;
                      WHERE ALLTRIM(card_nls) == ALLTRIM(sel_card_nls) AND BETWEEN(data_end, vvod_data_home, vvod_data_end) ;
                      INTO CURSOR sel_prozent ;
                      ORDER BY fio_rus, data_home, data_end

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                    colvo_zap = _TALLY

                    IF colvo_zap <> 0  && ���� ���������� ��������� ������� � �������� �� ����� ����, �� ��������� ��������� ������� �� �������� ������������ ���������

                      SELECT acc_cln_dt_pol
                      USE
                      USE (put_dbf) + ('acc_cln_dt.dbf') ALIAS acc_cln_dt_pol EXCLUSIVE
                      SET ORDER TO card_nls

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                      SELECT sel_prozent

                      SCAN

                        sel_fio_rus = ALLTRIM(sel_prozent.fio_rus)
                        sel_card_nls = ALLTRIM(sel_prozent.card_nls)

                        WAIT WINDOW '������ - ' + ALLTRIM(sel_fio_rus) + '  ���� ������� - ' + ALLTRIM(sel_card_nls) + '.  ������ � ��������� ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) NOWAIT

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
                      @ WROWS()/3,3 SAY PADC('������ ' + ALLTRIM(sel_fio_rus) + ' �� ������� � �������� �� ���� ������ ...',WCOLS())
                      =INKEY(2)

                      tim2 = SECONDS()

                      ACTIVATE WINDOW poisk
                      @ WROWS()/3,3 SAY PADC('������ � �������� � �������� �� ���� ������� ���������.' + ' (����� = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' ���.)',WCOLS())
                      =INKEY(2)
                      DEACTIVATE WINDOW poisk

                    ELSE

                      DEACTIVATE WINDOW poisk

                      =err('��������! ������ � ������� � �������� �� ���� �� ���������� ...')

                    ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                    ACTIVATE WINDOW poisk
                    @ WROWS()/3,3 SAY PADC('��������� ��������, ������ �������� ������ � ������� � ��������� ��������� ... ',WCOLS())
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

                    IF colvo_zap <> 0  && ���� ���������� ��������� ������� � �������� �� ����� ����, �� ��������� ��������� ������� �� �������� ������������ ���������

                      SELECT acc_cln_sum_pol
                      USE
                      USE (put_dbf) + ('acc_cln_sum.dbf') ALIAS acc_cln_sum_pol EXCLUSIVE
                      SET ORDER TO card_nls

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                      SELECT sel_prozent

                      SCAN

                        sel_fio_rus = ALLTRIM(sel_prozent.fio_rus)
                        sel_card_nls = ALLTRIM(sel_prozent.card_nls)

                        WAIT WINDOW '������ - ' + ALLTRIM(sel_fio_rus) + '  ���� ������� - ' + ALLTRIM(sel_card_nls) + '.  ������ � ��������� ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) NOWAIT

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
                      @ WROWS()/3,3 SAY PADC('������ ' + ALLTRIM(sel_fio_rus) + ' �� ������� � �������� �������� ������ ...',WCOLS())
                      =INKEY(2)

                      tim2 = SECONDS()

                      ACTIVATE WINDOW poisk
                      @ WROWS()/3,3 SAY PADC('������ � �������� � �������� �������� ������� ���������.' + ' (����� = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' ���.)',WCOLS())
                      =INKEY(2)
                      DEACTIVATE WINDOW poisk

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
                      DO del_acc_cln_null IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                    ELSE

                      DEACTIVATE WINDOW poisk

                      =err('��������! ������ � ������� � �������� �������� �� ���������� ...')

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

          CASE put_vibor = 3  && ������� ������ ������� �� ���� �������� � ��������� ��� �������

            num_mesaz = MONTH(new_data_odb)
            num_god = ALLTRIM(STR(YEAR(new_data_odb)))

            DO vid_kvartal_1 IN Raschet_proz_42308

            DO FORM (put_scr) + ('sel_data_proz.scx')

            IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

              num_mesaz = MONTH(vvod_data_end)
              num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

              DO vid_kvartal_2 IN Raschet_proz_42308

              tim1 = SECONDS()

              ACTIVATE WINDOW poisk
              @ WROWS()/3,3 SAY PADC('��������� ��������, ���� �������� ������ � ������� � ��������� �� ���� ... ',WCOLS())

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

              IF colvo_zap <> 0  && ���� ���������� ��������� ������� � �������� �� ����� ����, �� ��������� ��������� ������� �� �������� ������������ ���������

                SELECT acc_cln_dt_pol
                USE
                USE (put_dbf) + ('acc_cln_dt.dbf') ALIAS acc_cln_dt_pol EXCLUSIVE
                SET ORDER TO card_nls

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                SELECT sel_prozent

                SCAN

                  sel_fio_rus = ALLTRIM(sel_prozent.fio_rus)
                  sel_card_nls = ALLTRIM(sel_prozent.card_nls)

                  WAIT WINDOW '������ - ' + ALLTRIM(sel_fio_rus) + '  ���� ������� - ' + ALLTRIM(sel_card_nls) + '.  ������ � ��������� ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) NOWAIT

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
                @ WROWS()/3,3 SAY PADC('������ � ��������� ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) + ' �� ������� � �������� �� ���� ������� ...',WCOLS())
                =INKEY(2)

                tim2 = SECONDS()

                ACTIVATE WINDOW poisk
                @ WROWS()/3,3 SAY PADC('������ � �������� � �������� �� ���� ������� ���������.' + ' (����� = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' ���.)',WCOLS())
                =INKEY(2)
                DEACTIVATE WINDOW poisk

              ELSE

                DEACTIVATE WINDOW poisk

                =err('��������! ������ � ������� � �������� �� ���� �� ���������� ...')

              ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

              ACTIVATE WINDOW poisk
              @ WROWS()/3,3 SAY PADC('��������� ��������, ������ �������� ������ � ������� � ��������� ��������� ... ',WCOLS())
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

              IF colvo_zap <> 0  && ���� ���������� ��������� ������� � �������� �� ����� ����, �� ��������� ��������� ������� �� �������� ������������ ���������

                SELECT acc_cln_sum_pol
                USE
                USE (put_dbf) + ('acc_cln_sum.dbf') ALIAS acc_cln_sum_pol EXCLUSIVE
                SET ORDER TO card_nls

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                SELECT sel_prozent

                SCAN

                  sel_fio_rus = ALLTRIM(sel_prozent.fio_rus)
                  sel_card_nls = ALLTRIM(sel_prozent.card_nls)

                  WAIT WINDOW '������ - ' + ALLTRIM(sel_fio_rus) + '  ���� ������� - ' + ALLTRIM(sel_card_nls) + '.  ������ � ��������� ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) NOWAIT

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
                @ WROWS()/3,3 SAY PADC('������ � ��������� ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) + ' �� ������� � �������� �������� ������� ...',WCOLS())
                =INKEY(2)

                tim2 = SECONDS()

                ACTIVATE WINDOW poisk
                @ WROWS()/3,3 SAY PADC('������ � �������� � �������� �������� ������� ���������.' + ' (����� = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' ���.)',WCOLS())
                =INKEY(2)
                DEACTIVATE WINDOW poisk

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
                DO del_acc_cln_null IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

              ELSE

                DEACTIVATE WINDOW poisk

                =err('��������! ������ � ������� � �������� �������� �� ���������� ...')

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

=soob('���������� ���������� ������ � �������� �� ��������� ������ ...')

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
    =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('acc_cln_dt_null.cdx') + ' �� ������.')
    RETURN
  ENDIF

ELSE
  =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('acc_cln_dt_null.dbf') + ' �� ������.')
  RETURN
ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

=soob('������ �� ��������� ������� � �������� �� ���� ������� ...')

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
    =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('acc_cln_sum_null.cdx') + ' �� ������.')
    RETURN
  ENDIF

ELSE
  =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('acc_cln_sum_null.dbf') + ' �� ������.')
  RETURN
ENDIF

=soob('������ �� ��������� ������� � �������� �������� ������� ...')

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
  CASE pr_table_proz = .T.  && ��������� ���������� ������� ��������� �� �������

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('��������� ������� ����� ������� �� �������.',WCOLS())
    =INKEY(3)
    DEACTIVATE WINDOW poisk

  CASE pr_table_proz = .F.  && ��������� ���������� ������� ��������� �� ��������� ������

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('��������� ������� ����� ������� �� ��������� ������.',WCOLS())
    =INKEY(3)
    DEACTIVATE WINDOW poisk

    IF ROUND(VAL(SYS(2020)) / 1000000, 0) < 100.00

      ACTIVATE WINDOW poisk
      @ WROWS()/3,3 SAY PADC('�� ����� "�" ������ ���� ���������� ����� �������� 100 �����.',WCOLS())
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

  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

    vvod_data_end = new_data_odb

    DO FORM (put_scr) + ('vvod_data_proz.scx')

    IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

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

      IF _TALLY = 0  && ���� ������ �� ������� � ������� ����������� ����, �� ���� � ������� �������� ����

        SELECT DISTINCT branch, ref_client, ref_card, fio_rus, data_proz, card_nls, card_acct,;
          val_card, name_card, stavka, 000.00 AS sum_ostat, num_47411, pr_odb, num_dog, data_dog ;
          FROM acc_del ;
          WHERE data_proz <= vvod_data_end AND stavka <> 0 AND ALLTRIM(card_nls) == ALLTRIM(sel_card_nls) AND ALLTRIM(vid_card) == '1' ;
          INTO CURSOR sel_account ;
          ORDER BY branch, fio_rus

        IF _TALLY <> 0  && ���� ������ ������� � ������� �������� ����, �� ��������� ������ ���������

          DO raschet IN Raschet_proz_42301

        ELSE  && ���� ������ �� ������� � ������� �������� ����, �� ������� ��������� �� ������ � ������

          =err('��������! ��������� ���� ������ �� ������ �� � ��������, �� � �������� ������.')

        ENDIF

      ELSE  && ���� ������ ������� � ������� ����������� ����, �� ��������� ������ ���������

        DO raschet IN Raschet_proz_42301

      ENDIF
    ENDIF
  ENDIF

ELSE
  =err('��������! �� ���������� ���� ��������� �������� �������� �� ������� ...')
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
  CASE pr_table_proz = .T.  && ��������� ���������� ������� ��������� �� �������

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('��������� ������� ����� ������� �� �������.',WCOLS())
    =INKEY(3)
    DEACTIVATE WINDOW poisk

  CASE pr_table_proz = .F.  && ��������� ���������� ������� ��������� �� ��������� ������

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('��������� ������� ����� ������� �� ��������� ������.',WCOLS())
    =INKEY(3)
    DEACTIVATE WINDOW poisk

    IF ROUND(VAL(SYS(2020)) / 1000000, 0) < 100.00

      ACTIVATE WINDOW poisk
      @ WROWS()/3,3 SAY PADC('�� ����� "�" ������ ���� ���������� ����� �������� 100 �����.',WCOLS())
      =INKEY(3)
      DEACTIVATE WINDOW poisk

    ENDIF

ENDCASE

put_bar = 1

num_mesaz = MONTH(new_data_odb)
num_god = ALLTRIM(STR(YEAR(new_data_odb)))

DO vid_kvartal_1 IN Raschet_proz_42301

DO FORM (put_scr) + ('vvod_data_proz.scx')

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  vvod_data_pak = MaxDataMes(vvod_data_end)

  num_mesaz = MONTH(vvod_data_end)
  num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

  DO vid_kvartal_2 IN Raschet_proz_42301

* WAIT WINDOW DTOC(vvod_data_home) + '   ' + DTOC(vvod_data_end)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
  colvo_den_god = GOMONTH(CTOD('01.01.' + ALLTRIM(STR(YEAR(vvod_data_end)))), 12) - CTOD('01.01.' + ALLTRIM(STR(YEAR(vvod_data_end))))
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

* ----------------------------------------------------- ����� ����� �������� ��������� ������ ������� ��� �������� �� �������� ------------------------------------------------------------------------ *

  colvo_zap = 0

  DO CASE
    CASE pr_zarpl = .T.  && ������� ������� ����������� ������� �������

      text1 = '����� �������� ��������� ������ �������'
      text2 = '����� �������� ��������� �������� �� ��������'
      l_bar = 3
      =popup_9(text1, text2, text3, text4, l_bar)
      ACTIVATE POPUP vibor
      RELEASE POPUPS vibor

      IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

        DO CASE
          CASE BAR() = 1  && ����� �������� ��������� ������ �������

            =soob('��������! ���� ������� ����� ������� ��������� - ������ �������')

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

* BROWSE WINDOW brows TITLE '����� �������� ��������� ������ ������� + pr_zarpl = T'

          CASE BAR() = 3  && ����� �������� ��������� �������� �� ��������

            =soob('��������! ���� ������� ����� ������� ��������� - �������� �� ��������')

            =soob('���� ������ ��� ������� - ' + ALLTRIM(num_proekt))

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

* BROWSE WINDOW brows TITLE '����� �������� ��������� �������� �� �������� + pr_zarpl = T'

        ENDCASE

      ELSE

        =err('��������! ���� �� ������� ����� ������� ��������� ....')

        RETURN

      ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE pr_zarpl = .F.  && ������� ������� ����������� ������� ��������

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

* BROWSE WINDOW brows TITLE '����� �������� ��������� ������ ������� + pr_zarpl = F'

  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  colvo_zap = _TALLY

  IF colvo_zap <> 0  && ���� ���������� ��������� ������� � �������� �� ����� ����, �� ��������� ������ ���������

    WAIT WINDOW '���������� ������ �� ������� ����� ���������� ������ ��������� ����� - ' + ALLTRIM(STR(colvo_zap)) TIMEOUT 3

    put_bar = 1

    DO raschet IN Raschet_proz_42301

  ELSE
    =err('��������! � ���� ������ ������� �� ����������.')
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
  CASE pr_table_proz = .T.  && ��������� ���������� ������� ��������� �� �������

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('��������! ���� �������� � �������� ��������� ������ �� ������� ... ',WCOLS())

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
        =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('acc_cln_dt_null.cdx') + ' �� ������.')
        RETURN
      ENDIF

    ELSE
      =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('acc_cln_dt_null.dbf') + ' �� ������.')
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
        =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('acc_cln_sum_null.cdx') + ' �� ������.')
        RETURN
      ENDIF

    ELSE
      =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('acc_cln_sum_null.dbf') + ' �� ������.')
      RETURN
    ENDIF

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('��������� ������� ������� �� �������.' + ' (����� = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' ���.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

* ===================================================================================================================================== *

  CASE pr_table_proz = .F.  && ��������� ���������� ������� ��������� �� ��������� ������

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('��������! ���� �������� � �������� ��������� ������ �� ��������� ������ ... ',WCOLS())

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
        =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('acc_cln_dt_null.cdx') + ' �� ������.')
        RETURN
      ENDIF

    ELSE
      =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('acc_cln_dt_null.dbf') + ' �� ������.')
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
        =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('acc_cln_sum_null.cdx') + ' �� ������.')
        RETURN
      ENDIF

    ELSE
      =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('acc_cln_sum_null.dbf') + ' �� ������.')
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
        =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('istor_ost.cdx') + ' �� ������.')
        RETURN
      ENDIF

    ELSE
      =err('��������! ������ ��� ������ ���� - ' + (put_dbf) + ('istor_ost.dbf') + ' �� ������.')
      RETURN
    ENDIF

    tim2 = SECONDS()

    @ WROWS()/3,3 SAY PADC('��������� ������� ������� �� ��������� ������.' + ' (����� = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' ���.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

ENDCASE

RETURN


************************************************************ PROCEDURE close_acc_cln_null ***************************************************************

PROCEDURE close_acc_cln_null

DO CASE
  CASE pr_table_proz = .T.  && ��������� ���������� ������� ��������� �� �������

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

  CASE pr_table_proz = .F.  && ��������� ���������� ������� ��������� �� ��������� ������

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

* ------------------------------------------------------------------------------------------ ����� ������� � �������� �������� ------------------------------------------------------------------------------------------------ *

text1 = '�� ������� ������������ ������� �������� �� �� "TRANSMASTER"'
text2 = '�� ������� ������������ ������� �������� �� �� "CARD-VISA ������"'
l_bar = 3
=popup_9(text1, text2, text3, text4, l_bar)
ACTIVATE POPUP vibor
RELEASE POPUPS vibor

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  DO CASE
    CASE BAR() = 1  && �� ������� ������������ ������� �������� �� �� "TRANSMASTER"

      vid_istor_ost = 'istor_ost'

      SELECT istor_ost
      SET ORDER TO card_date  && INDEX ON ALLTRIM(card_nls) + DTOS(TTOD(date_ost_p)) TAG card_date

    CASE BAR() = 3  && �� ������� ������������ ������� �������� �� �� "CARD-VISA ������"

      vid_istor_ost = 'istor_mak'

      SELECT istor_mak
      SET ORDER TO card_date  && INDEX ON ALLTRIM(card_nls) + DTOS(TTOD(date_ost_p)) TAG card_date

  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('��������! ����� ������ ��������� �� ��������� ������ �������� ...',WCOLS())

  SELECT sel_account  && ��������� � �������������� ������� ��������� ��������

  SCAN  && ������ ������������ ������� ��������� ��������

    sel_card_nls = ALLTRIM(sel_account.card_nls)  && ����������� ���������� ����� ������������ ���������� �����

* ------------------------------------------------------------------------- ����� ���� ��������� ���� � ��������� ������� � ������ --------------------------------------------------------------------------------- *
    DO scan_calendar IN Raschet_proz_42301
* ---------------------------------------------------------------------- ������������ ������ ������������ �� ���� � ����� �� ����� ------------------------------------------------------------------------------ *
    DO svert_sum_den IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    WAIT WINDOW '������ - ' + ALLTRIM(sel_account.fio_rus) + '  ���� ������� - ' + ALLTRIM(sel_card_nls) + ' (����� = ' + ALLTRIM(STR(tim237 - tim230, 6, 3)) + ' ���.)' NOWAIT

    SELECT sel_account

  ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  WAIT CLEAR

  tim2 = SECONDS()

  @ WROWS()/3,3 SAY PADC('������ ��������� ������� ��������.' + ' (����� = ' + ALLTRIM(STR(tim2 - tim1, 8, 3)) + ' ���.)',WCOLS())
  =INKEY(2)

  @ WROWS()/3,3 SAY PADC('��������! ����� ������� ������������ ��������� �� ����� � ������� ������� ... ',WCOLS())
  =INKEY(2)

  tim2 = SECONDS()

* -------------------------------------------------------------------- ��������� ������������ ������ �� ���� � �������� ������� ----------------------------------------------------------------------------- *
  DO export_data_proz_dt IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  tim3 = SECONDS()

  @ WROWS()/3,3 SAY PADC('������� ������������ ��������� � ������� ������� ������� ��������' + ' (����� = ' + ALLTRIM(STR(tim3 - tim2, 8, 3)) + ' ���.)',WCOLS())
  =INKEY(2)

  @ WROWS()/3,3 SAY PADC('��������! ����� ������� ������������ �������� ��������� � ������� ������� ... ',WCOLS())
  =INKEY(2)

  tim2 = SECONDS()

* -------------------------------------------------------------------- ��������� ������������ ������ �� ����� � �������� ������� ---------------------------------------------------------------------------- *
  DO export_data_proz_sum IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  tim3 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('������� ������������ ��������� � ������� ������� ������� ��������' + ' (����� = ' + ALLTRIM(STR(tim3 - tim2, 8, 3)) + ' ���.)',WCOLS())
  =INKEY(2)
  DEACTIVATE WINDOW poisk

* ------------------------------------------------------------------------- ������� ������������ ��������� �� ��������� �������� -------------------------------------------------------------------------------- *
* DO export_del_proz IN Raschet_proz_42301
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  WAIT CLEAR

  tim3 = SECONDS()

  RELEASE tim230, tim231, tim232, tim233, tim234, tim235, tim236, tim237

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('��������� �� ������� ��������� ������� ���������.' + ' (����� = ' + ALLTRIM(STR(tim3 - tim1, 8, 3)) + ' ���.)',WCOLS())
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

* ������������ ������� calendar.dbf � ��������� ��������� ���: ������ ��������� vvod_data_home � ����� ��������� vvod_data_end
* ����� ����������� �������: ���� ����� �������� ����� ��������� ���� ������ ��������� vvod_data_home,
* �� ���� vvod_data_home ������ ��� ���� ������ ������� ��������� sel_account.data_proz, �� �� ������ ���������
* ������� ���� ������ ������� ��������� sel_account.data_proz

SCAN FOR BETWEEN(calendar.data_home, IIF(vvod_data_home < sel_account.data_proz, sel_account.data_proz, vvod_data_home), vvod_data_end)

*SCAN FOR IIF(vvod_data_home < sel_account.data_proz,;
BETWEEN(calendar.data_home, sel_account.data_proz, vvod_data_end),;
BETWEEN(calendar.data_home, vvod_data_home, vvod_data_end))

  tim231 = SECONDS()

  IF INLIST(SUBSTR(ALLTRIM(sel_card_nls), 6, 3), kod_val_usd, kod_val_eur)  && ������ � ������� ������ ��� �������� ������

* ------------------------------------------------------------------------------------------------------------ ����� ������ ---------------------------------------------------------------------------------------------------------------- *
    SET NEAR ON
* ------------------------------------------------------------------------------------------------ ���� ���� ������ � �������� -------------------------------------------------------------------------------------------------- *

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

* ------------------------------------------------------------------------------------------------------- ���� ���� ������ � ���� ------------------------------------------------------------------------------------------------- *

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

* ---------------------------------------------------------------------------------- ����� � ������� � ������ ������� ��������� ����� ------------------------------------------------------------------------------- *

    SELECT calendar

    SCATTER MEMVAR FIELDS kod_usd, kurs_usd, kod_eur, kurs_eur

    m.kod_usd = kod_val_usd
    m.kurs_usd = poisk_kurs_usd

    m.kod_eur = kod_val_eur
    m.kurs_eur = poisk_kurs_eur

    GATHER MEMVAR FIELDS kod_usd, kurs_usd, kod_eur, kurs_eur

  ENDIF  && ������ � ������� ������ ��� �������� ������

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  rec = ALLTRIM(sel_card_nls) + DTOS(calendar.data_home) + DTOS(calendar.data_end)

  SELECT acc_cln_dt_null

  poisk = SEEK(rec)

  IF poisk = .T. AND acc_cln_dt_null.pr = .F.   && ������ ������ ��������������� �������, ���������� ��, �� �����������.

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

    m.card_nls = ALLTRIM(sel_account.card_nls)  && ����� ���������� �����
    m.val_card = ALLTRIM(sel_account.val_card)  && ��� �����

    m.fio_rus = ALLTRIM(sel_account.fio_rus)  && ��� �������

    m.num_dog = sel_account.num_dog  && ����� ��������
    m.data_dog = sel_account.data_dog  && ���� ������ ���������� ��������

    m.data_proz = sel_account.data_proz  && ���� ������ ���������� ���������
    m.dt_hom_end = data_home_end  && ���� ������ �������� �� ������� ���������� ������
    m.data_home = calendar.data_home  && ��������� ���� ��������� �������, �������� ����������
    m.data_end = calendar.data_end  && �������� ���� ��������� �������, �������� ����������
    m.data_pak = vvod_data_pak  && �������� ���� ��������� ������� ��� �������� ������� ���������, ������������� �������������

* --------------------------------------------------------------------------------- ����� � ������� � ������ ������� ��������� ����� -------------------------------------------------------------------------------- *

    DO CASE
      CASE SUBSTR(ALLTRIM(m.card_nls), 6, 3) == ALLTRIM(kod_val_rur)
        m.kurs = 1  && ���� ������

      CASE SUBSTR(ALLTRIM(m.card_nls), 6, 3) == ALLTRIM(kod_val_usd)
        m.kurs = vvod_kurs_usd  && ���� ������

      CASE SUBSTR(ALLTRIM(m.card_nls), 6, 3) == ALLTRIM(kod_val_eur)
        m.kurs = vvod_kurs_eur  && ���� ������

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    m.kvartal = num_kvartal  && ����� ��������
    m.mes_kvart = num_mesaz_kvartal  && ����� ������ � ��������

    m.stavka = sel_account.stavka  && ���������� ������

* -------------------------------------------------------------------------------- ��������� ��������� ������� �� ����������� ���� ----------------------------------------------------------------------------------- *

    DO CASE
      CASE m.data_home = new_data_odb  && ���� ���������� ���� ����� ���� ��������� ������������� ���

        DO vxd_ost_vxod_ost IN Raschet_proz_42301

      CASE m.data_home > new_data_odb AND BETWEEN((m.data_home - new_data_odb), 1, 3) && ���� ���������� ���� ������ ���� ��������� ������������� ���

        DO vxd_ost_isxod_ost IN Raschet_proz_42301

      CASE m.data_home < new_data_odb  && ���� ���������� ���� ������ ���� ������������� ���

        DO vxd_ost_istor_ost IN Raschet_proz_42301

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    IF m.sum_ostat <> sum_ostat_prev
      m.ref_ost = SYS(2015)
    ELSE
      m.ref_ost = ALLTRIM(ref_ost_prev)
    ENDIF

    sum_ostat_prev = m.sum_ostat  && ����� ������� �� ���������� ����
    ref_ost_prev = m.ref_ost  && ����� ������� �� ���������� ����

    m.colvo_den = IIF(calendar.data_home < sel_account.data_proz, (calendar.data_end - sel_account.data_proz), calendar.colvo_den)

* -------------------------------------------------------------------------------------------- �������� �������� ������� ���������� -------------------------------------------------------------------------------------- *
    por_okrugl = 5 && ������� ���������� ��� ������� ROUND()
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    sum_dobav_1 = (m.sum_ostat / colvo_den_god) * m.colvo_den

    sum_dobav_2 = ROUND(((sum_dobav_1 * sel_account.stavka) / 100), por_okrugl)

    m.sum_dobav = IIF(m.pr = .T., sum_dobav_2, m.sum_dobav)  && ����� ���������

    sum_rasch_1 = (m.sum_ostat / colvo_den_god) * m.colvo_den

    sum_rasch_2 = ROUND(((sum_rasch_1 * m.stavka) / 100), por_okrugl)

    m.sum_rasch = IIF(m.pr = .T.,sum_rasch_2, m.sum_rasch)  && ����� ���������

    m.sum_rasxod = IIF(m.pr = .T., ROUND((m.sum_rasch * m.kurs), por_okrugl), m.sum_rasxod)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE INLIST(ALLTRIM(m.val_card), 'USD', 'EUR')  && �������� �����, ���������� �������� ������ ������

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

        DO CASE
          CASE m.stavka > stavka_ref AND INLIST(LEFT(ALLTRIM(m.card_nls), 5), '40817')  && ��� ����������

* ���� � ������� ������ ������ 9.00 %, �� � ���� ����� ����� ���������� ����� � ������� 35%
* ������� ������� ����� �� ��� ������ = sum_nalog_1
* ����� ������� ����� �� ������ 9.0 % = sum_nalog_2
* ����� ������� ������� (sum_nalog_1 - sum_nalog_2)
* � ����� � ��� ���������� ����� � ������� 35% (((sum_nalog_1 - sum_nalog_2) * 35) / 100) = sum_nalog_3

            sum_nalog_1 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * m.stavka) / 100

            sum_nalog_2 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * 9) / 100

            sum_nalog_3 = ROUND((((sum_nalog_1 - sum_nalog_2) * 35) / 100), por_okrugl)

            m.sum_nalog = IIF(m.pr = .T., sum_nalog_3, m.sum_nalog)  && ����� ������

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

          CASE m.stavka > 9.00 AND INLIST(LEFT(ALLTRIM(m.card_nls), 5), '40820')  && ��� ����������

* ���� � ������� ������ ������ 9.00 %, �� � ���� ����� ����� ���������� ����� � ������� 35%
* ������� ������� ����� �� ��� ������ = sum_nalog_1
* ����� ������� ����� �� ������ 9.0 % = sum_nalog_2
* ����� ������� ������� (sum_nalog_1 - sum_nalog_2)
* � ����� � ��� ���������� ����� � ������� 35% (((sum_nalog_1 - sum_nalog_2) * 35) / 100) = sum_nalog_3

            sum_nalog_1 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * m.stavka) / 100

            sum_nalog_2 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * 9) / 100

            sum_nalog_3 = ROUND((((sum_nalog_1 - sum_nalog_2) * 35) / 100), por_okrugl)

            m.sum_nalog = IIF(m.pr = .T., sum_nalog_3, m.sum_nalog)  && ����� ������

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

*            CASE INLIST(LEFT(ALLTRIM(m.card_nls), 5), '40820')  && ��� �� ����������    ��� ������ ��� ������ �������� �������� �������

*  * ���� ������ �� �������� , �� � ���� ����� ����� ���������� ����� � ������� 30%
*  * ������� ������� ����� �� ��� ������ = sum_nalog_1
*  * ����� ����� � ��� ���������� ����� � ������� 30% ((sum_nalog_1 * 30) / 100) = sum_nalog_2

*              sum_nalog_1 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * m.stavka) / 100

*              sum_nalog_2 = ROUND(((sum_nalog_1 * 30) / 100), por_okrugl)

*              m.sum_nalog = IIF(m.pr = .T., sum_nalog_2, m.sum_nalog)  && ����� ������

          OTHERWISE

            m.sum_nalog = 0

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE INLIST(ALLTRIM(m.val_card), 'RUR')  && �������� ����� �����, ������ ������ �� ������������

        m.sum_nalog = 0

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    m.sum_viplat = m.sum_rasch - m.sum_nalog  && ����� � �������

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

* ----------------------------------------------------------------------------- ������������ ������� �������� �� �� "TRANSMASTER" ----------------------------------------------------------------------------- *

DO CASE
  CASE ALLTRIM(vid_istor_ost) == 'istor_ost'

    SELECT istor_ost

    SET NEAR ON

    poisk_ostat = SEEK(ALLTRIM(m.card_nls) + DTOS(m.data_home))  && ���� ������� �� ����� �� �������� ����

    SET NEAR OFF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

        SELECT acc_cln_dt_null

        m.sum_ostat = istor_ost.begin_bal  && ����� ������� �� ����

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE poisk_ostat = .F.  AND m.data_home < TTOD(istor_ost.date_ost_p) AND ALLTRIM(istor_ost.card_nls) == ALLTRIM(m.card_nls)

        SELECT istor_ost
        SKIP -1

        IF ALLTRIM(istor_ost.card_nls) == ALLTRIM(m.card_nls) AND m.data_home > TTOD(istor_ost.date_ost_p)

          SELECT acc_cln_dt_null

          m.sum_ostat = istor_ost.end_bal  && ����� ������� �� ����

        ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE poisk_ostat = .F. AND ALLTRIM(istor_ost.card_nls) <> ALLTRIM(m.card_nls)

        SELECT istor_ost
        SKIP -1

        IF ALLTRIM(istor_ost.card_nls) == ALLTRIM(m.card_nls) AND m.data_home > TTOD(istor_ost.date_ost_p)

          m.sum_ostat = istor_ost.end_bal  && ����� ������� �� ����

        ELSE

          SELECT MAX(date_ost_p) AS date_ost_p, branch, ref_client, card_nls, card_acct, begin_bal, debit, credit, end_bal ;
            FROM istor_ost ;
            WHERE ALLTRIM(card_nls) == ALLTRIM(m.card_nls) AND date_ost_p < m.data_home ;
            INTO CURSOR  sel_ostatok

          SELECT acc_cln_dt_null

          IF _TALLY <> 0

            m.sum_ostat = sel_ostatok.end_bal  && ����� ������� �� ����

          ELSE

            SELECT account
            SET ORDER TO card_nls

            poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && ���� ������� �� ����� �� �������� ����

            DO CASE
              CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

                SELECT acc_cln_dt_null

                m.sum_ostat = account.begin_bal  && ����� ������� �� ����

              CASE poisk_ostat = .F.  && ����� �� ��������, �� ����� �� ������������ ����

                SELECT acc_del
                SET ORDER TO card_nls

                poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && ���� ������� �� ����� �� �������� ����

                DO CASE
                  CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

                    SELECT acc_cln_dt_null

                    m.sum_ostat = acc_del.begin_bal  && ����� ������� �� ����

                  CASE poisk_ostat = .F.  && ����� �� ��������, �� ����� �� ������������ ����

                    m.sum_ostat = 0.00  && ����� ������� �� ����

                ENDCASE

            ENDCASE

          ENDIF

        ENDIF

    ENDCASE

* --------------------------------------------------------------------------- ������������ ������� �������� �� �� "CARD-VISA ������" --------------------------------------------------------------------------- *

  CASE ALLTRIM(vid_istor_ost) == 'istor_mak'

    SELECT istor_mak

    SET NEAR ON

    poisk_ostat = SEEK(ALLTRIM(m.card_nls) + DTOS(m.data_home))  && ���� ������� �� ����� �� �������� ����

    SET NEAR OFF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

        SELECT acc_cln_dt_null

        m.sum_ostat = istor_mak.vxd_ost_m  && ����� ������� �� ����

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE poisk_ostat = .F.  AND m.data_home < TTOD(istor_mak.date_ost_m) AND ALLTRIM(istor_mak.card_nls) == ALLTRIM(m.card_nls)

        SELECT istor_mak
        SKIP -1

        IF ALLTRIM(istor_mak.card_nls)  == ALLTRIM(m.card_nls) AND m.data_home > TTOD(istor_mak.date_ost_m)

          SELECT acc_cln_dt_null

          m.sum_ostat = istor_mak.isx_ost_m  && ����� ������� �� ����

        ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE poisk_ostat = .F. AND ALLTRIM(card_nls) <> ALLTRIM(m.card_nls)

        SELECT istor_mak
        SKIP -1

        IF ALLTRIM(istor_mak.card_nls) == ALLTRIM(m.card_nls) AND m.data_home > TTOD(istor_mak.date_ost_m)

          SELECT acc_cln_dt_null

          m.sum_ostat = istor_mak.isx_ost_m  && ����� ������� �� ����

        ELSE

          SELECT MAX(date_ost_m) AS date_ost_m, branch, ref_client, card_nls, card_acct, vxd_ost_m, debit_m, credit_m, isx_ost_m ;
            FROM istor_mak ;
            WHERE ALLTRIM(card_nls) == ALLTRIM(m.card_nls) AND TTOD(date_ost_m) < m.data_home ;
            INTO CURSOR  sel_ostatok

          SELECT acc_cln_dt_null

          IF _TALLY <> 0

            m.sum_ostat = sel_ostatok.isx_ost_m  && ����� ������� �� ����

          ELSE

            SELECT account
            SET ORDER TO card_nls

            poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && ���� ������� �� ����� �� �������� ����

            DO CASE
              CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

                SELECT acc_cln_dt_null

                m.sum_ostat = account.vxd_ost_m  && ����� ������� �� ����

              CASE poisk_ostat = .F.  && ����� �� ��������, �� ����� �� ������������ ����

                SELECT acc_del
                SET ORDER TO card_nls

                poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && ���� ������� �� ����� �� �������� ����

                DO CASE
                  CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

                    SELECT acc_cln_dt_null

                    m.sum_ostat = acc_del.vxd_ost_m  && ����� ������� �� ����

                  CASE poisk_ostat = .F.  && ����� �� ��������, �� ����� �� ������������ ����

                    m.sum_ostat = 0.00  && ����� ������� �� ����

                ENDCASE
            ENDCASE


          ENDIF

        ENDIF

    ENDCASE

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

RETURN


******************************************************* PROCEDURE vxd_ost_vxod_ost *********************************************************************

PROCEDURE vxd_ost_vxod_ost  && ���� ���������� ���� ����� ���� ��������� ������������� ���

* ----------------------------------------------------------------------------- ������������ ������� �������� �� �� "TRANSMASTER" ----------------------------------------------------------------------------- *

DO CASE
  CASE ALLTRIM(vid_istor_ost) == 'istor_ost'

    SELECT account
    SET ORDER TO card_nls

    poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && ���� ������� �� ����� �� �������� ����

    DO CASE
      CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

        SELECT acc_cln_dt_null

        m.sum_ostat = account.begin_bal  && ����� ������� �� ����

      CASE poisk_ostat = .F.  && ����� �� ��������, �� ����� �� ������������ ����

        SELECT acc_del
        SET ORDER TO card_nls

        poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && ���� ������� �� ����� �� �������� ����

        DO CASE
          CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

            SELECT acc_cln_dt_null

            m.sum_ostat = acc_del.begin_bal  && ����� ������� �� ����

          CASE poisk_ostat = .F.  && ����� �� ��������, �� ����� �� ������������ ����

            m.sum_ostat = 0.00  && ����� ������� �� ����

        ENDCASE
    ENDCASE

* --------------------------------------------------------------------------- ������������ ������� �������� �� �� "CARD-VISA ������" --------------------------------------------------------------------------- *

  CASE ALLTRIM(vid_istor_ost) == 'istor_mak'

    SELECT account
    SET ORDER TO card_nls

    poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && ���� ������� �� ����� �� �������� ����

    DO CASE
      CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

        SELECT acc_cln_dt_null

        m.sum_ostat = account.vxd_ost_m  && ����� ������� �� ����

      CASE poisk_ostat = .F.  && ����� �� ��������, �� ����� �� ������������ ����

        SELECT acc_del
        SET ORDER TO card_nls

        poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && ���� ������� �� ����� �� �������� ����

        DO CASE
          CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

            SELECT acc_cln_dt_null

            m.sum_ostat = acc_del.vxd_ost_m  && ����� ������� �� ����

          CASE poisk_ostat = .F.  && ����� �� ��������, �� ����� �� ������������ ����

            m.sum_ostat = 0.00  && ����� ������� �� ����

        ENDCASE
    ENDCASE

ENDCASE

RETURN


******************************************************* PROCEDURE vxd_ost_isxod_ost *********************************************************************

PROCEDURE vxd_ost_isxod_ost  && ���� ���������� ���� ������ ���� ��������� ������������� ���

* ----------------------------------------------------------------------------- ������������ ������� �������� �� �� "TRANSMASTER" ----------------------------------------------------------------------------- *

DO CASE
  CASE ALLTRIM(vid_istor_ost) == 'istor_ost'

    SELECT account
    SET ORDER TO card_nls

    poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && ���� ������� �� ����� �� �������� ����

    DO CASE
      CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

        SELECT acc_cln_dt_null

        m.sum_ostat = account.end_bal  && ����� ������� �� ����

      CASE poisk_ostat = .F.  && ����� �� ��������, �� ����� �� ������������ ����

        SELECT acc_del
        SET ORDER TO card_nls

        poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && ���� ������� �� ����� �� �������� ����

        DO CASE
          CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

            SELECT acc_cln_dt_null

            m.sum_ostat = acc_del.end_bal  && ����� ������� �� ����

          CASE poisk_ostat = .F.  && ����� �� ��������, �� ����� �� ������������ ����

            m.sum_ostat = 0.00  && ����� ������� �� ����

        ENDCASE
    ENDCASE

* --------------------------------------------------------------------------- ������������ ������� �������� �� �� "CARD-VISA ������" --------------------------------------------------------------------------- *

  CASE ALLTRIM(vid_istor_ost) == 'istor_mak'

    SELECT account
    SET ORDER TO card_nls

    poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && ���� ������� �� ����� �� �������� ����

    DO CASE
      CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

        SELECT acc_cln_dt_null

        m.sum_ostat = account.isx_ost_m  && ����� ������� �� ����

      CASE poisk_ostat = .F.  && ����� �� ��������, �� ����� �� ������������ ����

        SELECT acc_del
        SET ORDER TO card_nls

        poisk_ostat = SEEK(ALLTRIM(m.card_nls))  && ���� ������� �� ����� �� �������� ����

        DO CASE
          CASE poisk_ostat = .T.  && ����� ��������, �� ����� �� ������������ ����

            SELECT acc_cln_dt_null

            m.sum_ostat = acc_del.isx_ost_m  && ����� ������� �� ����

          CASE poisk_ostat = .F.  && ����� �� ��������, �� ����� �� ������������ ����

            m.sum_ostat = 0.00  && ����� ������� �� ����

        ENDCASE
    ENDCASE

ENDCASE

RETURN


****************************************************** PROCEDURE svert_sum_den ***********************************************************************

PROCEDURE svert_sum_den

tim234 = SECONDS()

* ---------------------------------------------------------------------------------------- �������� �������� ������� ���������� ------------------------------------------------------------------------------------------- *

por_okrugl_det = 4  && ������� ���������� ��� ������� ROUND()

* ---------------------------------------------------------- ���������� ������� ������ �� ������� ��������� �� ���� � ����������� � ������� -------------------------------------------------------- *

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

* -------------------------------------------------------------------- �������� ����������� ������� ��������� ������� � �������� ------------------------------------------------------------------------------ *

SELECT sel_proz_sum
GO TOP

SCAN

  tim235 = SECONDS()

  rec = ALLTRIM(sel_proz_sum.card_nls) + DTOS(sel_proz_sum.data_home) + DTOS(sel_proz_sum.data_end)

  SELECT acc_cln_sum_null

  poisk = SEEK(rec)

  IF poisk = .T. AND acc_cln_sum_null.pr = .F.   && ������ ������ ��������������� �������, ���������� ��, �� �����������.

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

    m.card_nls = ALLTRIM(sel_proz_sum.card_nls)  && ����� ���������� �����
    m.val_card = ALLTRIM(sel_proz_sum.val_card)  && ��� �����

    m.fio_rus = ALLTRIM(sel_proz_sum.fio_rus)  && ��� �������

    m.num_dog = sel_proz_sum.num_dog  && ����� ��������
    m.data_dog = sel_proz_sum.data_dog  && ���� ������ ���������� ��������

    m.data_proz = sel_proz_sum.data_proz  && ���� ������ ���������� ���������

    m.dt_hom_end = sel_proz_sum.dt_hom_end  && ���� ������ �������� �� ������� ���������� ������
    m.data_home = sel_proz_sum.data_home  && ��������� ���� ��������� �������, �������� ����������
    m.data_end = sel_proz_sum.data_end  && �������� ���� ��������� �������, �������� ����������
    m.data_pak = sel_proz_sum.data_pak  && �������� ���� ��������� ������� ��� �������� ������� ���������, ������������� �������������

    DO CASE
      CASE SUBSTR(ALLTRIM(m.card_nls),6,3) == ALLTRIM(kod_val_rur)
        m.kurs = 1  && ���� ������

      CASE SUBSTR(ALLTRIM(m.card_nls),6,3) == ALLTRIM(kod_val_usd)
        m.kurs = vvod_kurs_usd  && ���� ������

      CASE SUBSTR(ALLTRIM(m.card_nls),6,3) == ALLTRIM(kod_val_eur)
        m.kurs = vvod_kurs_eur  && ���� ������

    ENDCASE

    m.kvartal = sel_proz_sum.kvartal  && ����� ��������
    m.mes_kvart = sel_proz_sum.mes_kvart  && ����� ������ � ��������

    m.stavka = sel_proz_sum.stavka  && ���������� ������
    m.sum_ostat = sel_proz_sum.sum_ostat  && ����� ������� �� �����

    m.den_mes = sel_proz_sum.den_mes

* ---------------------------------------------------------------------------------------------- �������� �������� ������� ���������� ------------------------------------------------------------------------------------- *
    por_okrugl_sum = 2  && ������� ���������� ��� ������� ROUND()
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    m.sum_dobav = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_dobav, por_okrugl_sum), m.sum_dobav)
    m.sum_rasch = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_rasch, por_okrugl_sum), m.sum_rasch)
    m.sum_rasxod = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_rasxod, por_okrugl_sum), m.sum_rasxod)
    m.sum_nalog = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_nalog, por_okrugl_sum), m.sum_nalog)
    m.sum_viplat = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_viplat, por_okrugl_sum), m.sum_viplat)

* ----------------------------------------------------------------- ��������� ������ ��������� �� ���� � ����� �� ����� �� ������ ���������� ----------------------------------------------------------- *
    por_okrugl_prover = 2  && ������� ���������� ��� ������� ROUND()


    prover_sum_dobav = ROUND(((((m.sum_ostat / colvo_den_god) * m.den_mes) * m.stavka) / 100), por_okrugl_prover)  && ����� ��������� ��� ��������

    m.sum_dobav = IIF(m.sum_dobav = prover_sum_dobav, m.sum_dobav, prover_sum_dobav)

    prover_sum_rasch = ROUND(((((m.sum_ostat / colvo_den_god) * m.den_mes) * m.stavka) / 100), por_okrugl_prover)  && ����� ��������� ��� ��������

    m.sum_rasch = IIF(m.sum_rasch = prover_sum_rasch, m.sum_rasch, prover_sum_rasch)

    prover_sum_rasxod = ROUND((prover_sum_rasch * m.kurs), por_okrugl_prover)

* m.sum_rasxod = IIF(m.sum_rasxod = prover_sum_rasxod, m.sum_rasxod, prover_sum_rasxod)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE INLIST(ALLTRIM(m.val_card), 'USD', 'EUR')  && �������� �����, ���������� �������� ������ ������

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

        DO CASE
          CASE m.stavka > 9 AND LEFT(ALLTRIM(m.card_nls), 5) == '40817'  && ��� ����������

* ���� � ������� ������ ������ 9.0 %, �� � ���� ����� ����� ���������� ����� � ������� 35%
* ������� ������� ����� �� ��� ������ = sum_nalog_1
* ����� ������� ����� �� ������ 9.0 % = sum_nalog_2
* ����� ������� ������� (sum_nalog_1 - sum_nalog_2)
* � ����� � ��� ���������� ����� � ������� 35% (((sum_nalog_1 - sum_nalog_2) * 35) / 100) = sum_nalog_3

            sum_nalog_1 = (((m.sum_ostat / colvo_den_god) * m.den_mes) * m.stavka) / 100

            sum_nalog_2 = (((m.sum_ostat / colvo_den_god) * m.den_mes) * 9) / 100

            prover_sum_nalog = ROUND((((sum_nalog_1 - sum_nalog_2) * 35) / 100), por_okrugl_prover)  && ����� ������

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

          CASE m.stavka > 9.00 AND INLIST(LEFT(ALLTRIM(m.card_nls), 5), '40820')  && ��� ������������

* ���� � ������� ������ ������ 9.00 %, �� � ���� ����� ����� ���������� ����� � ������� 35%
* ������� ������� ����� �� ��� ������ = sum_nalog_1
* ����� ������� ����� �� ������ 9.0 % = sum_nalog_2
* ����� ������� ������� (sum_nalog_1 - sum_nalog_2)
* � ����� � ��� ���������� ����� � ������� 35% (((sum_nalog_1 - sum_nalog_2) * 35) / 100) = sum_nalog_3

            sum_nalog_1 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * m.stavka) / 100

            sum_nalog_2 = (((m.sum_ostat / colvo_den_god) * m.colvo_den) * 9) / 100

            sum_nalog_3 = ROUND((((sum_nalog_1 - sum_nalog_2) * 35) / 100), por_okrugl)

            m.sum_nalog = IIF(m.pr = .T., sum_nalog_3, m.sum_nalog)  && ����� ������


          OTHERWISE

            m.sum_nalog = 0

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE INLIST(ALLTRIM(m.val_card), 'RUR')  && �������� ����� �����, ������ ������ �� ������������

        m.sum_nalog = 0

    ENDCASE

    m.sum_viplat = m.sum_rasch - m.sum_nalog  && ����� � �������

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

  text1 = '������� ������ ��������� �� ������ ������� � ��������� ��� �������'
  text2 = '������� ������ ������� �� ���� �������� � ��������� ��� �������'
  l_bar = 3
  =popup_9(text1, text2, text3, text4, l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

    put_vibor = BAR()

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE put_vibor = 1  && ������� ������ ��������� �� ������ ������� � ��������� ��� �������

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

            IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

              vvod_data_end = new_data_odb

              DO FORM (put_scr) + ('sel_data_proz.scx')

              IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

                num_mesaz = MONTH(vvod_data_end)
                num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

                DO vid_kvartal_2 IN Raschet_proz_42301

*          WAIT WINDOW DTOC(vvod_data_home) + '   ' + DTOC(vvod_data_end)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                DO WHILE .T.

                  text1 = '������� ������ � ������������ ���������� ��������'
                  text2 = '������� ������ � ������������� ���������� ��������'
                  l_bar = 3
                  =popup_9(text1, text2, text3, text4, l_bar)
                  ACTIVATE POPUP vibor
                  RELEASE POPUPS vibor

                  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

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
                      CASE put_prn = 1  && ������� ������ � ������������ ���������� ��������

                        DO sel_data_det_cln IN Raschet_proz_42301

                      CASE put_prn = 3  && ������� ������ � ������������� ���������� ��������

                        DO sel_data_sum_cln IN Raschet_proz_42301

                    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                    IF _TALLY <> 0

                      SELECT sel_prozent

                      DO CASE
                        CASE bar_num = 11  && ������ ����� ���� ��� ��������� ������ �� ������

                          DO brows_proz IN Raschet_proz_42301

                        CASE bar_num = 13  && ������ ����� ���� ��� ��������� ������ �� ��������

                          DO prn_proz IN Raschet_proz_42301

                      ENDCASE

                    ELSE
                      =err('��������! �� ��������� ���� ������ �� ���������� ...')
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
            =err('��������! �� ���������� ���� ��������� ���� ������� �� ������� ... ')
          ENDIF

        ENDDO

* ====================================================================================================================================== *

      CASE put_vibor = 3  && ������� ������ ������� �� ���� �������� � ��������� ��� �������

        num_mesaz = MONTH(new_data_odb)
        num_god = ALLTRIM(STR(YEAR(new_data_odb)))

        DO vid_kvartal_1 IN Raschet_proz_42308

* vvod_data_home = CTOD('01.05.2010')
* vvod_data_end = CTOD('31.05.2010')

        DO FORM (put_scr) + ('sel_data_proz.scx')

        IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

          num_mesaz = MONTH(vvod_data_end)
          num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

          DO vid_kvartal_2 IN Raschet_proz_42308

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

          DO WHILE .T.

            DO CASE
              CASE bar_num = 11  && ������ ����� ���� ��� ��������� ������ �� ������

                text1 = '����� � ������������ ���������� ��������'
                text2 = '����� � ������������� ���������� ��������'
                text3 = '����� � ������������ �� ���������� ������'
                l_bar = 5
                =popup_9(text1, text2, text3, text4, l_bar)

              CASE bar_num = 13  && ������ ����� ���� ��� ��������� ������ �� ��������

                text1 = '����� � ������������ ���������� �������� - �������'
                text2 = '����� � ������������ ���������� �������� - ���������'
                text3 = '����� � ������������� �������� �������� ������ ��� - ���������'
                text4 = '����� � ������������� �������� �������� ����� ���� - ���������'
                text5 = '����� � ������������ �� ���������� ������ - ���������'
                l_bar = 9
                =popup_big(text1, text2, text3, text4, text5, text6, text7, text8, l_bar)

            ENDCASE

            ACTIVATE POPUP vibor
            RELEASE POPUPS vibor

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

            IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

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
                CASE bar_num = 11  && ������ ����� ���� ��� ��������� ������ �� ������

                  DO CASE
                    CASE put_prn = 1  && ����� � ������������ ���������� �������� - ����� � 1

                      DO sel_data_det_all IN Raschet_proz_42301

                    CASE put_prn = 3  && ����� � ������������� ���������� �������� - ����� � 2

                      DO sel_data_sum_all IN Raschet_proz_42301

                    CASE put_prn = 5  && ������� ������ � ������������� �� ���������� ������ - ����� � 3

                      DO sel_data_sum_tran IN Raschet_proz_42301

                  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

                CASE bar_num = 13  && ������ ����� ���� ��� ��������� ������ �� ��������

                  DO CASE
                    CASE put_prn = 1  && ����� � ������������ ���������� �������� - ������� - ����� � 1

                      DO sel_data_det_all IN Raschet_proz_42301

                    CASE put_prn = 3  && ����� � ������������ �������� �������� ������ ��� - ��������� - ����� � 2

                      DO sel_data_det_all IN Raschet_proz_42301

                    CASE put_prn = 5  && ����� � ������������� �������� �������� ������ ��� - ��������� - ����� � 3

                      DO sel_data_sum_all IN Raschet_proz_42301

                    CASE put_prn = 7  && ����� � ������������� �������� �������� ����� ���� - ��������� - ����� � 4

                      DO sel_data_sum_all IN Raschet_proz_42301

                    CASE put_prn = 9  && ����� � ������������ �� ���������� ������ - ��������� - ����� � 5

                      DO sel_data_sum_tran IN Raschet_proz_42301

                  ENDCASE

              ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

              IF _TALLY <> 0

                SELECT sel_prozent
                GO TOP

                DO CASE
                  CASE bar_num = 11  && ������ ����� ���� ��� ��������� ������ �� ������
                    DO brows_proz IN Raschet_proz_42301

                  CASE bar_num = 13  && ������ ����� ���� ��� ������ ������ �� ��������
                    DO prn_proz IN Raschet_proz_42301

                ENDCASE

              ELSE
                =err('��������! �� ��������� ���� ������ �� ���������� ...')
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
  CASE pr_zarpl = .T.  && ������� ������� ����������� ������� �������

    DO CASE
      CASE pr_summa_null = .T.  && ���� �������=.T., �� ������� ����� ���������:

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

      CASE pr_summa_null = .F.  && ���� �������=.F., �� ������� ����� �� ���������:

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

  CASE pr_zarpl = .F.  && ������� ������� ����������� ������� ��������

    DO CASE
      CASE pr_summa_null = .T.  && ���� �������=.T., �� ������� ����� ���������:

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

      CASE pr_summa_null = .F.  && ���� �������=.F., �� ������� ����� �� ���������:

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
  CASE pr_zarpl = .T.  && ������� ������� ����������� ������� �������

    DO CASE
      CASE pr_summa_null = .T.  && ���� �������=.T., �� ������� ����� ���������:

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

      CASE pr_summa_null = .F.  && ���� �������=.F., �� ������� ����� �� ���������:

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

  CASE pr_zarpl = .F.  && ������� ������� ����������� ������� ��������

    DO CASE
      CASE pr_summa_null = .T.  && ���� �������=.T., �� ������� ����� ���������:

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

      CASE pr_summa_null = .F.  && ���� �������=.F., �� ������� ����� �� ���������:

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
  CASE pr_zarpl = .T.  && ������� ������� ����������� ������� �������

    DO CASE
      CASE pr_summa_null = .T.  && ���� �������=.T., �� ������� ����� ���������:

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

      CASE pr_summa_null = .F.  && ���� �������=.F., �� ������� ����� �� ���������:

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

  CASE pr_zarpl = .F.  && ������� ������� ����������� ������� ��������

    DO CASE
      CASE pr_summa_null = .T.  && ���� �������=.T., �� ������� ����� ���������:

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

      CASE pr_summa_null = .F.  && ���� �������=.F., �� ������� ����� �� ���������:

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
  CASE pr_zarpl = .T.  && ������� ������� ����������� ������� �������

    DO CASE
      CASE pr_summa_null = .T.  && ���� �������=.T., �� ������� ����� ���������:

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

      CASE pr_summa_null = .F.  && ���� �������=.F., �� ������� ����� �� ���������:

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

  CASE pr_zarpl = .F.  && ������� ������� ����������� ������� ��������

    DO CASE
      CASE pr_summa_null = .T.  && ���� �������=.T., �� ������� ����� ���������:

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

      CASE pr_summa_null = .F.  && ���� �������=.F., �� ������� ����� �� ���������:

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
  CASE pr_zarpl = .T.  && ������� ������� ����������� ������� �������

    DO CASE
      CASE pr_summa_null = .T.  && ���� �������=.T., �� ������� ����� ���������:

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

      CASE pr_summa_null = .F.  && ���� �������=.F., �� ������� ����� �� ���������:

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

  CASE pr_zarpl = .F.  && ������� ������� ����������� ������� ��������

    DO CASE
      CASE pr_summa_null = .T.  && ���� �������=.T., �� ������� ����� ���������:

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

      CASE pr_summa_null = .F.  && ���� �������=.F., �� ������� ����� �� ���������:

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
  CASE put_prn  =  1  && ������� ������ � ������������ ���������� ��������

    EDIT FIELDS ;
      branch:H = ' ������ :',;
      ref_client:H = ' ������ :',;
      ref_card:H = ' ����� :',;
      num_dog:H = ' ����� �������� :',;
      fio_rus:H = ' ������� ��� �������� :',;
      dt_hom_end:H = ' ������ ���� � �������� :',;
      data_home:H = ' ��������� ���� ������� :',;
      data_end:H = ' �������� ���� ������� :',;
      card_nls:H = ' ��������� ���� ��� :',;
      name_card:H = ' ��� ����� :',;
      kurs:H = ' ����:':P='999.9999',;
      stavka:H = ' ������ :':P='999.99',;
      sum_ostat:H = ' ����� ������� �� ����� :':P='999 999.99',;
      kvartal:H = ' ����� �������� :',;
      mes_kvart:H = ' ����� ������ � �������� :',;
      den_mes:H = ' ���� � ������ :',;
      sum_dobav:H = ' ����� � ������������ :':P='9 999.99',;
      sum_viplat:H = ' ����� ��������� � ������� :':P='9 999.99',;
      num_47411:H = ' ���� ����������� ��������� :' ;
      TITLE ' �������� ������ �� ������������ ��������� ' ;
      WINDOW brow_proz

*         sum_rasch:H = ' ����� ���������� �� ���� �������� �� ������� :':P='9 999.99',;
*         sum_rasxod:H = ' ����� ���������� �� ������� ����� :':P='9 999.99',;
*         sum_nalog:H = ' ����� ����������� ������ :':P='9 999.99',;

  CASE put_prn = 3  && ������� ������ � ������������� ���������� ��������

    EDIT FIELDS ;
      branch:H = ' ������ :',;
      ref_client:H = ' ������ :',;
      ref_card:H = ' ����� :',;
      num_dog:H = ' ����� �������� :',;
      fio_rus:H = ' ������� ��� �������� :',;
      dt_hom_end:H = ' ������ ���� � �������� :',;
      data_home:H = ' ��������� ���� ������� :',;
      data_end:H = ' �������� ���� ������� :',;
      card_nls:H = ' ��������� ���� ��� :',;
      name_card:H = ' ��� ����� :',;
      kurs:H = ' ����:':P='999.9999',;
      stavka:H = ' ������ :':P='999.99',;
      kvartal:H = ' ����� �������� :',;
      mes_kvart:H = ' ����� ������ � �������� :',;
      den_mes:H = ' ���� � ������ :',;
      sum_dobav:H = ' ����� � ������������ :':P='9 999.99',;
      sum_viplat:H = ' ����� ��������� � ������� :':P='9 999.99',;
      num_47411:H = ' ���� ����������� ��������� :' ;
      TITLE ' �������� ������ �� ������������ ��������� ' ;
      WINDOW brow_proz

*         sum_rasch:H = ' ����� ���������� �� ���� �������� �� ������� :':P='9 999.99',;
*         sum_rasxod:H = ' ����� ���������� �� ������� ����� :':P='9 999.99',;
*         sum_nalog:H = ' ����� ����������� ������ :':P='9 999.99',;

  CASE put_prn = 5  && ������� ������ � ������������� �� ���������� ������

    EDIT FIELDS ;
      branch:H = ' ������ :',;
      data_home:H = ' ��������� ���� ������� :',;
      data_end:H = ' �������� ���� ������� :',;
      tran_42301:H = ' ���������� ���� :',;
      name_card:H = ' ��� ����� :',;
      name_card:H = ' ��� ����� :',;
      kurs:H = ' ����:':P='999.9999',;
      kvartal:H = ' ����� �������� :',;
      mes_kvart:H = ' ����� ������ � �������� :',;
      sum_dobav:H = ' ����� � ������������ :':P='9 999.99',;
      sum_viplat:H = ' ����� ��������� � ������� :':P='9 999.99',;
      num_47411:H = ' ���� ����������� ��������� :' ;
      TITLE ' �������� ������ �� ������������ ��������� ' ;
      WINDOW brow_proz

*         sum_rasch:H = ' ����� ���������� �� ���� �������� �� ������� :':P='9 999.99',;
*         sum_rasxod:H = ' ����� ���������� �� ������� ����� :':P='9 999.99',;
*         sum_nalog:H = ' ����� ����������� ������ :':P='9 999.99',;

ENDCASE

RELEASE WINDOW brow_proz

RETURN


******************************************************************** PROCEDURE prn_proz ****************************************************************

PROCEDURE prn_proz

DO FORM (put_scr) + ('data_prn.scx')

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  text1 = '�������� ����� �� �����'
  text2 = '�������� ����� �� �������'
  l_bar = 3
  =popup_9(text1, text2, text3, text4, l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    DO CASE
      CASE put_prn = 1  && ����� � ������������ ���������� �������� - ������� - ����� � 1

        DO CASE
          CASE BAR() = 1

            REPORT FORM (put_frx) + ('proz_mesaz_42301_d_v.frx') NOEJECT PREVIEW
            * REPORT FORM (put_frx) + ('proz_mesaz_42301_d_v.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_d_v.doc') ASCII

          CASE BAR() = 3

            DO CASE
              CASE par_prompt = .T.  && ��� ������ ������ �� ������ ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_mesaz_42301_d_v.frx') NOEJECT NOCONSOLE TO PRINTER PROMPT
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_d_v.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_d_v.doc') ASCII

              CASE par_prompt = .F.  && ��� ������ ������ �� ������ �� ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_mesaz_42301_d_v.frx') NOEJECT NOCONSOLE TO PRINTER
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_d_v.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_d_v.doc') ASCII

            ENDCASE

        ENDCASE

      CASE put_prn = 3  && ����� � ������������ �������� �������� ������ ��� - ��������� - ����� � 2

        DO CASE
          CASE BAR() = 1

            REPORT FORM (put_frx) + ('proz_mesaz_42301_d_g.frx') NOEJECT PREVIEW
            * REPORT FORM (put_frx) + ('proz_mesaz_42301_d_g.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_d_g.doc') ASCII

          CASE BAR() = 3

            DO CASE
              CASE par_prompt = .T.  && ��� ������ ������ �� ������ ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_mesaz_42301_d_g.frx') NOEJECT NOCONSOLE TO PRINTER PROMPT
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_d_g.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_d_g.doc') ASCII

              CASE par_prompt = .F.  && ��� ������ ������ �� ������ �� ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_mesaz_42301_d_g.frx') NOEJECT NOCONSOLE TO PRINTER
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_d_g.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_d_g.doc') ASCII

            ENDCASE

        ENDCASE

      CASE put_prn = 5  && ����� � ������������� �������� �������� ������ ��� - ��������� - ����� � 4

        DO CASE
          CASE BAR() = 1

            REPORT FORM (put_frx) + ('proz_mesaz_42301_pn.frx') NOEJECT PREVIEW
            * REPORT FORM (put_frx) + ('proz_mesaz_42301_pn.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_pn.doc') ASCII

          CASE BAR() = 3

            DO CASE
              CASE par_prompt = .T.  && ��� ������ ������ �� ������ ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_mesaz_42301_pn.frx') NOEJECT NOCONSOLE TO PRINTER PROMPT
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_pn.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_pn.doc') ASCII

              CASE par_prompt = .F.  && ��� ������ ������ �� ������ �� ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_mesaz_42301_pn.frx') NOEJECT NOCONSOLE TO PRINTER
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_pn.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_pn.doc') ASCII

            ENDCASE

        ENDCASE

      CASE put_prn = 7  && ����� � ������������� �������� �������� ����� ���� - ��������� - ����� � 5

        DO CASE
          CASE BAR() = 1

            REPORT FORM (put_frx) + ('proz_mesaz_42301_pv.frx') NOEJECT PREVIEW
            * REPORT FORM (put_frx) + ('proz_mesaz_42301_pv.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_pv.doc') ASCII

          CASE BAR() = 3

            DO CASE
              CASE par_prompt = .T.  && ��� ������ ������ �� ������ ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_mesaz_42301_pv.frx') NOEJECT NOCONSOLE TO PRINTER PROMPT
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_pv.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_pv.doc') ASCII

              CASE par_prompt = .F.  && ��� ������ ������ �� ������ �� ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_mesaz_42301_pv.frx') NOEJECT NOCONSOLE TO PRINTER
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_pv.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_pv.doc') ASCII

            ENDCASE

        ENDCASE

      CASE put_prn = 9  && ����� � ������������ �� ���������� ������ - ��������� - ����� � 6

        DO CASE
          CASE BAR() = 1

            REPORT FORM (put_frx) + ('proz_mesaz_42301_tr.frx') NOEJECT PREVIEW
            * REPORT FORM (put_frx) + ('proz_mesaz_42301_tr.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_tr.doc') ASCII

          CASE BAR() = 3

            DO CASE
              CASE par_prompt = .T.  && ��� ������ ������ �� ������ ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_mesaz_42301_tr.frx') NOEJECT NOCONSOLE TO PRINTER PROMPT
                * REPORT FORM (put_frx) + ('proz_mesaz_42301_tr.frx') NOEJECT NOCONSOLE TO FILE (put_tmp) + ('proz_mesaz_42301_tr.doc') ASCII

              CASE par_prompt = .F.  && ��� ������ ������ �� ������ �� ���������� ������ ������ ������������ ���������

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

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  num_mesaz = MONTH(vvod_data_end)
  num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

  DO vid_kvartal_2 IN Raschet_proz_42301

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
*    IF pr_zarpl = .T.
*      =soob('��������! ������������ ����� ��� - ' + ALLTRIM(title_zarpl))
*    ENDIF
* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  text1 = '������������ ������ ��� �������� � ��� �� ���� ����� �����'
  text2 = '������������ ������ ��� �������� � ��� �� ���� ������ RUR'
  text3 = '������������ ������ ��� �������� � ��� �� ���� ������ USD'
  text4 = '������������ ������ ��� �������� � ��� �� ���� ������ EUR'
  l_bar = 7
  =popup_9(text1, text2, text3, text4, l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

    DO CASE
      CASE BAR() = 1  && ������������ ������ ��� �������� � ��� �� ���� ����� �����
        sel_val_card = SPACE(0)

      CASE BAR() = 3  && ������������ ������ ��� �������� � ��� �� ���� ������ RUR
        sel_val_card = 'RUR'

      CASE BAR() = 5  && ������������ ������ ��� �������� � ��� �� ���� ������ USD
        sel_val_card = 'USD'

      CASE BAR() = 7  && ������������ ������ ��� �������� � ��� �� ���� ������ EUR
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
      '��������� ��������' AS prim, kod_proekt, grup ;
      FROM sel_prozent ;
      INTO CURSOR sel_export ;
      ORDER BY branch, grup, val_card, name_card, fio_rus, data_home

    colvo_zap = _TALLY

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    IF colvo_zap <> 0

      SELECT popol_ckc

*!*  ��� �����:  Pb02FFFK.NNN
*!*  P - ������������� �����
*!*  b - ������� ���� VISA
*!*  02 - ��� �����
*!*  fff - ��������� (������) �����

*!*  001-������,
*!*  002-�����,
*!*  003-����,
*!*  004-������,
*!*  005-�������,
*!*  006-������,
*!*  007-�����,
*!*  008-�����������,
*!*  009-���������,
*!*  010-�����������,
*!*  011-�������,
*!*  013-�������� ����

*!*  k - ����� ����� �� ���� (0,1,2,3�)
*!*  nnn - ��������� ���� (���������� ���� � ������ ����)

      colvo_ckc = 7
      data_ckc = DATE()
      explog_ckc = .F.

      DO WHILE .T.

        DO FORM (put_scr) + ('popol_ckc.scx')

        IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

          exp_file = 'PB02' + (fff) + ALLTRIM(STR(colvo_ckc)) + '.' + (nnn)

          IF NOT FILE((put_out_pb) + (exp_file))

            explog_ckc = .T.
            EXIT

          ELSE

            =err('��������! ���� - ' + (put_out_pb) + (exp_file) + ' - � ���������� ��� ���������� !!! ')

            text1 = '�� ������� ������������ ���� � ������ �������'
            text2 = '�� ������� ������ ����� ����� ��� ����� ��������'
            l_bar = 3
            =popup_9(text1, text2, text3, text4, l_bar)
            ACTIVATE POPUP vibor
            RELEASE POPUPS vibor

            IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

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
        @ WROWS()/3-0.2,3 SAY PADC('��������! ������ ������������ ������ ��� ��������� "TRANSMASTER" ... ',WCOLS())

        SET TEXTMERGE ON
        SET TEXTMERGE NOSHOW

        IF FILE((put_out_pb) + (exp_file))
          DELETE FILE (put_out_pb) + (exp_file)
        ENDIF

        _text = FCREATE((put_out_pb) + (exp_file))

        IF _text < 1

          =err('��������! ������ ��� �������� ����� �������� �������� ... ')
          =FCLOSE(_text)

        ELSE

          data_ckc = DATE()

          SELECT sel_export
          GO TOP

* ---------------------------------------------------------------------------------------------- ������������ ��������� ����� -------------------------------------------------------------------------------------------------- *

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

* --------------------------------------------------------------------------------------------- ������������ ����� ����������� ------------------------------------------------------------------------------------------------ *

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

* ------------------------------------------------------------------------------------------------- ������������ ��������� ����� ---------------------------------------------------------------------------------------------- *

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
        @ WROWS()/3-0.2,3 SAY PADC('������������ ������ ��� ��������� "TRANSMASTER" ������� ��������� ... ',WCOLS())
        =INKEY(2)
        DEACTIVATE WINDOW poisk

        =soob('���������� ���� - ' + (exp_file) + ' - � ���������� ' + (put_out_pb) + ' ����������� ... ')

        ACTIVATE WINDOW poisk
        @ WROWS()/3-0.2,3 SAY PADC('���������� �������������� �������� - ' + ALLTRIM(TRANSFORM(sel_colvo.colvo_all, '999 999 999')) + ;
          '  �� ����� ����� - ' + ALLTRIM(TRANSFORM(sel_summa.summa_all, '999 999 999 999.99')), WCOLS())
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
      =err('��������! �� ���������� ���� ������� ������ �� ���������� ...')
    ENDIF

  ELSE
    =err('��������! �� ���������� �� ������������ ��������� ����� ��� ��� ...')
  ENDIF

ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

IF INLIST(ALLTRIM(num_branch), '04')  && ��������� ����������� ��������� ����� ��� ��� ����������� �� ������� �������� ������� ������ � 4 (24.03.2004)

  text1='� � � � � � � � �  ����� ����������� ����� � �������� ����������'
  text2='� �  � � � � � � � � �  ����� ����������� ����� � �������� ����������'
  l_bar=3
  =popup_9(text1,text2,text3,text4,l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

    DO CASE
      CASE BAR() = 1  && � � � � � � � � �  ����� ����������� ����� � �������� ����������

        IF FILE((put_out_pb) + (exp_file))

          =soob('��������! ���� � ���������� ' + (put_out_pb) + (exp_file) + ' ������ ...')

          IF FILE((put_upk_1) + (exp_file))
            DELETE FILE (put_upk_1) + (exp_file)
          ENDIF

          IF FILE((put_upk_2) + (exp_file))
            DELETE FILE (put_upk_2) + (exp_file)
          ENDIF

          COPY FILE (put_out_pb) + (exp_file) TO (put_upk_1) + (exp_file)
          COPY FILE (put_out_pb) + (exp_file) TO (put_upk_2) + (exp_file)

          =soob('��������! ���� - ' + (put_out_pb) + (exp_file) + ' - ���������� �� ������� ���� ' + (put_upk_1))
          =soob('��������! ���� - ' + (put_out_pb) + (exp_file) + ' - ���������� �� ������� ���� ' + (put_upk_2))

          IF FILE((put_out_pb) + (exp_file))
            DELETE FILE (put_out_pb) + (exp_file)
          ENDIF

        ELSE

          =err('��������! ���� - ' + (put_out_pb) + (exp_file) + ' - � ���������� �� ���������� !!! ')

        ENDIF

      CASE BAR() = 3  && � �  � � � � � � � � �  ����� ����������� ����� � �������� ����������

        =err('�� ���������� ���������� ���� - ' + (put_out_pb) + (exp_file) + ' � �������� ����������.')

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

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

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
    '��������� �������� �� ������ ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) AS kod_name ;
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
    @ WROWS()/3-0.2,3 SAY PADC('��������! ����� ������� ������ �� ��������� "CARD-���� VISA" ... ',WCOLS())
    =INKEY(2)

        DO expsum_rsbank IN Raschet_proz_42301

    ACTIVATE WINDOW poisk
    @ WROWS()/3-0.2,3 SAY PADC('������� ������ �� ��������� "CARD-���� VISA" ������� �������� ... ',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ELSE
    =soob('��������! ������������� ���� ������ �� ���������� ...')
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

ELSE        &&  ���� flag_prichislen = .F.

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

COPY TO (put_exp) + (imavisadoc) FOX2X AS 866 && �������� �������������� � ��������� �� ���������� ������ � RsBank

IF sergeizhuk_flag = .F.
  COPY TO (put_rsbank) + (imavisadoc) FOX2X AS 866 && �������� �������������� � ��������� �� ���������� ������ � RsBank
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

* ----------------------------------------------------------------- �������� ����������� ������� ��������� ����������� �� ����� ------------------------------------------------------------------------------ *

SCAN

  WAIT WINDOW '������ - ' + ALLTRIM(acc_cln_dt_null.fio_rus) + '  ���� ������� - ' + ALLTRIM(acc_cln_dt_null.card_nls) NOWAIT

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
  m.num_dog = acc_cln_dt_null.num_dog  && ����� ��������
  m.data_dog = acc_cln_dt_null.data_dog  && ���� ������ ���������� ��������

  m.fio_rus = ALLTRIM(acc_cln_dt_null.fio_rus)  && ��� �������
  m.card_nls = ALLTRIM(acc_cln_dt_null.card_nls)  && ����� ���������� �����
  m.val_card = ALLTRIM(acc_cln_dt_null.val_card)  && ��� �����
  m.pr = acc_cln_dt_null.pr

  m.data_proz = acc_cln_dt_null.data_proz  && ���� ������ ���������� ���������
  m.dt_hom_end = acc_cln_dt_null.dt_hom_end  && ���� ������ �������� �� ������� ���������� ������
  m.data_home = acc_cln_dt_null.data_home  && ��������� ���� ��������� �������, �������� ����������
  m.data_end = acc_cln_dt_null.data_end  && �������� ���� ��������� �������, �������� ����������
  m.data_pak = acc_cln_dt_null.data_pak  && �������� ���� ��������� ������� ��� �������� ������� ���������, ������������� �������������

  m.colvo_den = acc_cln_dt_null.colvo_den

  m.kvartal = acc_cln_dt_null.kvartal  && ����� ��������
  m.mes_kvart = acc_cln_dt_null.mes_kvart  && ����� ������ � ��������

  m.kurs = acc_cln_dt_null.kurs  && ���� ������

  m.stavka = acc_cln_dt_null.stavka  && ���������� ������
  m.sum_ostat = acc_cln_dt_null.sum_ostat  && ����� ������� �� �����

  m.sum_dobav = acc_cln_dt_null.sum_dobav  && ����� ���������
  m.sum_rasch = acc_cln_dt_null.sum_rasch  && ����� ���������
  m.sum_rasxod = acc_cln_dt_null.sum_rasxod
  m.sum_nalog = acc_cln_dt_null.sum_nalog  && ����� ������
  m.sum_viplat = acc_cln_dt_null.sum_viplat  && ����� � �������

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

* ----------------------------------------------------------------- �������� ����������� ������� ��������� ������� � �������� ---------------------------------------------------------------------------------- *

SCAN

  WAIT WINDOW '������ - ' + ALLTRIM(acc_cln_sum_null.fio_rus) + '  ���� ������� - ' + ALLTRIM(acc_cln_sum_null.card_nls) NOWAIT

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

  m.num_dog = acc_cln_sum_null.num_dog  && ����� ��������
  m.data_dog = acc_cln_sum_null.data_dog  && ���� ������ ���������� ��������

  m.fio_rus = ALLTRIM(acc_cln_sum_null.fio_rus)  && ��� �������

  m.card_nls = ALLTRIM(acc_cln_sum_null.card_nls)  && ����� ���������� �����

  m.val_card = ALLTRIM(acc_cln_sum_null.val_card)  && ��� �����

  m.pr = acc_cln_sum_null.pr

  m.data_proz = acc_cln_sum_null.data_proz  && ���� ������ ���������� ���������
  m.dt_hom_end = acc_cln_sum_null.dt_hom_end  && ���� ������ �������� �� ������� ���������� ������
  m.data_home = acc_cln_sum_null.data_home  && ��������� ���� ��������� �������, �������� ����������
  m.data_end = acc_cln_sum_null.data_end  && �������� ���� ��������� �������, �������� ����������
  m.data_pak = acc_cln_sum_null.data_pak  && �������� ���� ��������� ������� ��� �������� ������� ���������, ������������� �������������

  m.den_mes = acc_cln_sum_null.den_mes

  m.kvartal = acc_cln_sum_null.kvartal  && ����� ��������
  m.mes_kvart = acc_cln_sum_null.mes_kvart  && ����� ������ � ��������

  m.kurs = acc_cln_sum_null.kurs  && ���� ������

  m.stavka = acc_cln_sum_null.stavka  && ���������� ������
  m.sum_ostat = acc_cln_sum_null.sum_ostat  && ����� ������� �� �����

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

=soob('��������! ���������� ������ ��������� ������� ������������ ������ ������ � ��������� ... ')

tex1 = '� � � � � � � � �  �������� ������ �� ��������� ��������'
tex2 = '� � � � � � � � � �  �� �������� ������ �� ��������� ��������'
l_bar = 3
=popup_9(tex1,tex2,tex3,tex4,l_bar)
ACTIVATE POPUP vibor
RELEASE POPUPS vibor
IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  IF BAR() = 1

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('��������! ������ ������� ������������ ��������� �� ��������� �������� ... ',WCOLS())
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

      =soob('������ � ��������� ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) + ' �� ������� � �������� �� ���� ������� ...')

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

        =soob('������ � ��������� ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) + ' �� ������� � �������� �������� ������� ...')

      ELSE
        =err('��������! �� ��������� ���� ������ �� ���������� ...')
      ENDIF

    ELSE
      =err('��������! �� ��������� ���� ������ �� ���������� ...')
    ENDIF

    tim3 = SECONDS()

    ACTIVATE WINDOW poisk
    @ WROWS()/3,3 SAY PADC('������� ��������� ������ ������� ���������' + ' (����� = ' + ALLTRIM(STR(tim3 - tim2, 8, 3)) + ' ���.)',WCOLS())
    =INKEY(2)
    DEACTIVATE WINDOW poisk

  ENDIF
ENDIF

RETURN


*********************************************************************************************************************************************************






































































