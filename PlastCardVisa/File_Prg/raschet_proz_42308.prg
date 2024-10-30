************************************************
*** ��������� ������������ ������� ���������***
************************************************

* ������������ ��� ��������� ������� acc_strah.dbf � acc_strah_dt.dbf
* acc_strah.dbf - ��� ������� �������� ������ �� �������� ������, �� ������� ���� ��������� ��������
* �� ������� �� �����. ����� ������ ��������� � ������ �������, ��������� �������� ���������� �� ������ �������� � ��������.
* acc_strah_dt.dbf - ��� ������� �������� ����������� ����� �� �������� ������ � ������� ��������� ���.

* ������������ ���������
* sel_fio - ��� ���������� ������ �������������� ������ �� ����������� �������, �������� �������� �������
* ������� �� ������������� ������, ������� ����������� � sel_fio �� ������ ������ ������������ � ������� acc_strah.dbf
* ��� ������ ������������ ����� DO FORM (put_scr) + ('selfio_card.scx')
* red_data - ���������� �������������� ������
* ��� �������������� ������������ ����� DO FORM (put_scr) + ('red_proz_42308.scx')
* � ��������� read_sum ��������� ��������������� ����������� ������� �������������� ������������ ���� � ��������� ��������
* ����������� �������� ������ � ������ ������� �������


************************************************************** PROCEDURE sel_fio ************************************************************************

PROCEDURE sel_fio
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

SELECT * ;
  FROM acc_strah ;
  INTO CURSOR count_zapis

IF _TALLY <> 0

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *
  DO FORM (put_scr)+('poisk_client.scx')
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  DO WHILE .T.

    SELECT DISTINCT SPACE(1) + fio_rus + ' - ' + ALLTRIM(val_card) + ' - ' + ALLTRIM(card_nls) AS fio_rus ;
      FROM acc_strah ;
      WHERE LEFT(ALLTRIM(fio_rus), len_client) == ALLTRIM(fam_client) ;
      INTO CURSOR tov ;
      ORDER BY 1

* BROWSE WINDOW brows

* INTO CURSOR tov ;
* INTO TABLE (put_dbf) + ('tov.dbf') ;

    DO FORM (put_scr) + ('selfio_nls.scx')

    IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

      SELECT acc_strah
      SET ORDER TO card_nls

      IF SEEK(sel_card_nls)

        DO red_data IN Raschet_proz_42308

      ELSE
        =err('��������! � ���� ������ ������� �� ����������.')
      ENDIF

    ELSE
      EXIT
    ENDIF

  ENDDO

ELSE
  =err('��������! � ���� ������ ������� �� ����������.')
ENDIF

SELECT (sel)
RETURN


****************************************************************** PROCEDURE red_data ******************************************************************

PROCEDURE red_data

Poisk_1 = SPACE(20)
mBal_1 = SPACE(5)
mVal_1 = '840'
mKl_1 = SPACE(1)
mLiz_1 = SPACE(11)

Poisk_2 = SPACE(20)
mBal_2 = SPACE(5)
mVal_2 = '840'
mKl_2 = SPACE(1)
mLiz_2 = SPACE(11)

DO WHILE .T.

  SCATTER MEMVAR MEMO

  m.data = DATETIME()
  m.ima_pk = SYS(0)

  Poisk_1 = ALLTRIM(m.num_42308)
  mBal_1 = SUBSTR(ALLTRIM(m.num_42308),1,5)
  mVal_1 = SUBSTR(ALLTRIM(m.num_42308),6,3)
  mKl_1 = SUBSTR(ALLTRIM(m.num_42308),9,1)
  mLiz_1 = SUBSTR(ALLTRIM(m.num_42308),10,11)

  Poisk_2 = ALLTRIM(m.num_47411)
  mBal_2 = SUBSTR(ALLTRIM(m.num_47411),1,5)
  mVal_2 = SUBSTR(ALLTRIM(m.num_47411),6,3)
  mKl_2 = SUBSTR(ALLTRIM(m.num_47411),9,1)
  mLiz_2 = SUBSTR(ALLTRIM(m.num_47411),10,11)

  DO FORM (put_scr) + ('red_proz_42308.scx')

  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC
    text1 = ' ���� �p�������� � ����� ����� '
    text2 = ' ��p������ � p���� ����� ������ '
    text3 = ' �������� ��������� ������ '
    l_bar = 5
    =popup_9(text1,text2,text3,text4,l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC
      DO CASE
        CASE BAR() = 1

          BEGIN TRANSACTION

          GATHER MEMVAR MEMO

          UPDATE acc_strah SET ;
            date_nls = m.date_nls,;
            card_nls = m.card_nls,;
            num_dog = m.num_dog,;
            data_dog = m.data_dog,;
            data_proz = m.data_proz,;
            stavka = m.stavka,;
            sum_strah = m.sum_strah,;
            num_47411 = m.num_47411 ;
            WHERE ALLTRIM(card_nls) == sel_card_nls

          UPDATE acc_strah_dt SET ;
            card_nls = m.card_nls,;
            num_dog = m.num_dog,;
            data_dog = m.data_dog,;
            data_proz = m.data_proz ;
            WHERE ALLTRIM(card_nls) == sel_card_nls

          UPDATE acc_strah_sum SET ;
            card_nls = m.card_nls,;
            num_dog = m.num_dog,;
            data_dog = m.data_dog,;
            data_proz = m.data_proz ;
            WHERE ALLTRIM(card_nls) == sel_card_nls

          END TRANSACTION

          EXIT

        CASE BAR()=3
          LOOP

        CASE BAR()=5
          EXIT

      ENDCASE
    ENDIF

  ELSE
    EXIT
  ENDIF
ENDDO

RETURN


******************************************************************* PROCEDURE del_data *****************************************************************

PROCEDURE del_data
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

=soob('��������! ���������� ������ ��������� ������� ������������ ������ ������ � ��������� ... ')

tex1 = '������ ��������� �������� ������ �� ������������ ���������'
tex2 = '���������� �� ���������� ��������� �������� ������'
l_bar = 3
=popup_9(tex1,tex2,tex3,tex4,l_bar)
ACTIVATE POPUP vibor
RELEASE POPUPS vibor

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  IF BAR() = 1

    text1 = '������� ������ ��������� �� ������ ������� � ��������� ��� �������'
    text2 = '������� ������ ������� �� ���� �������� � ��������� ��� �������'
    l_bar = 3
    =popup_9(text1,text2,text3,text4,l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

      put_vibor = BAR()

      DO CASE
        CASE put_vibor = 1  && ������� ������ ��������� �� ������ ������� � ��������� ��� �������

          num_mesaz = MONTH(DATE())
          num_god = ALLTRIM(STR(YEAR(DATE())))

          DO vid_kvartal_1 IN Raschet_proz_42301

          DO FORM (put_scr) + ('sel_data_proz.scx')

          IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

            num_mesaz = MONTH(vvod_data_end)
            num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

            DO vid_kvartal_2 IN Raschet_proz_42301

            DO WHILE .T.

              SELECT DISTINCT MAX(LEN(ALLTRIM(fio_rus))) AS len_fio ;
                FROM acc_strah_dt ;
                WHERE BETWEEN(data_end, vvod_data_home, vvod_data_end) ;
                UNION ;
                SELECT DISTINCT MAX(LEN(ALLTRIM(fio_rus))) AS len_fio ;
                FROM acc_strah_sum ;
                WHERE BETWEEN(data_end, vvod_data_home, vvod_data_end) ;
                INTO CURSOR sel_len_fio

              SELECT DISTINCT SPACE(1) + PADR(ALLTRIM(fio_rus),sel_len_fio.len_fio) + ' - ' + ALLTRIM(val_card) + ' - ' + ALLTRIM(card_nls) AS fio_rus ;
                FROM acc_strah_dt ;
                WHERE BETWEEN(data_end, vvod_data_home, vvod_data_end) ;
                UNION ;
                SELECT DISTINCT SPACE(1) + PADR(ALLTRIM(fio_rus),sel_len_fio.len_fio) + ' - ' + ALLTRIM(val_card) + ' - ' + ALLTRIM(card_nls) AS fio_rus ;
                FROM acc_strah_sum ;
                WHERE BETWEEN(data_end, vvod_data_home, vvod_data_end) ;
                INTO CURSOR tov ;
                ORDER BY 1

              IF _TALLY <> 0

                DO FORM (put_scr) + ('selfio_nls.scx')

                IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

                  SELECT branch, ref_client, fio_rus, card_nls ;
                    FROM acc_strah_dt ;
                    WHERE BETWEEN(data_end, vvod_data_home, vvod_data_end) AND ALLTRIM(card_nls) == sel_card_nls ;
                    INTO CURSOR sel_prozent ;
                    ORDER BY 1, 2

                  IF _TALLY <> 0

                    SELECT acc_strah_dt
                    USE
                    USE (put_dbf) + ('acc_strah_dt.dbf') EXCLUSIVE
                    DELETE FOR BETWEEN(data_end, vvod_data_home, vvod_data_end) AND ALLTRIM(card_nls) == sel_card_nls
                    PACK
                    USE
                    USE (put_dbf) + ('acc_strah_dt.dbf') SHARED
                    SET ORDER TO card_data

                    =soob('������ ' + ALLTRIM(sel_prozent.fio_rus) + ' �� ������� � �������� �� ���� ������ ...')

                    SELECT branch, ref_client, fio_rus, card_nls ;
                      FROM acc_strah_sum ;
                      WHERE BETWEEN(data_end, vvod_data_home, vvod_data_end) AND ALLTRIM(card_nls) == sel_card_nls ;
                      INTO CURSOR sel_prozent ;
                      ORDER BY 1, 2

                    IF _TALLY <> 0

                      SELECT acc_strah_sum
                      USE
                      USE (put_dbf) + ('acc_strah_sum.dbf') EXCLUSIVE
                      DELETE FOR BETWEEN(data_end, vvod_data_home, vvod_data_end) AND ALLTRIM(card_nls) == sel_card_nls
                      PACK
                      USE
                      USE (put_dbf) + ('acc_strah_sum.dbf') SHARED
                      SET ORDER TO card_data

                      =soob('������ ' + ALLTRIM(sel_prozent.fio_rus) + ' �� ������� � �������� �������� ������ ...')

                    ELSE
                      =err('��������! �� ��������� ���� ������ �� ���������� ...')
                    ENDIF

                  ELSE
                    =err('��������! �� ��������� ���� ������ �� ���������� ...')
                  ENDIF

                ELSE
                  EXIT
                ENDIF

              ELSE
                =err('��������! �� ��������� ���� ������ �� ���������� ...')
                EXIT
              ENDIF

            ENDDO

          ENDIF

        CASE put_vibor = 3  && ������� ������ ������� �� ���� �������� � ��������� ��� �������

          num_mesaz = MONTH(DATE())
          num_god = ALLTRIM(STR(YEAR(DATE())))

          DO vid_kvartal_1 IN Raschet_proz_42308

          DO FORM (put_scr) + ('sel_data_proz.scx')

          IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

            num_mesaz = MONTH(vvod_data_end)
            num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

            DO vid_kvartal_2 IN Raschet_proz_42308

            SELECT branch, ref_client, fio_rus, card_nls ;
              FROM acc_strah_dt ;
              WHERE BETWEEN(data_end, vvod_data_home, vvod_data_end) ;
              INTO CURSOR sel_prozent ;
              ORDER BY 1, 2

            IF _TALLY <> 0

              SELECT acc_strah_dt
              USE
              USE (put_dbf) + ('acc_strah_dt.dbf') EXCLUSIVE
              DELETE FOR BETWEEN(data_end, vvod_data_home, vvod_data_end)
              PACK
              USE
              USE (put_dbf) + ('acc_strah_dt.dbf') SHARED
              SET ORDER TO card_data

              =soob('������ � ��������� ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) + ' �� ������� � �������� �� ���� ������� ...')

              SELECT branch, ref_client, fio_rus, card_nls ;
                FROM acc_strah_sum ;
                WHERE BETWEEN(data_end, vvod_data_home, vvod_data_end) ;
                INTO CURSOR sel_prozent ;
                ORDER BY 1, 2

              IF _TALLY <> 0

                SELECT acc_strah_sum
                USE
                USE (put_dbf) + ('acc_strah_sum.dbf') EXCLUSIVE
                DELETE FOR BETWEEN(data_end, vvod_data_home, vvod_data_end)
                PACK
                USE
                USE (put_dbf) + ('acc_strah_sum.dbf') SHARED
                SET ORDER TO card_data

                =soob('������ � ��������� ' + DTOC(vvod_data_home) + ' - ' + DTOC(vvod_data_end) + ' �� ������� � �������� �������� ������� ...')

              ELSE
                =err('��������! �� ��������� ���� ������ �� ���������� ...')
              ENDIF

            ELSE
              =err('��������! �� ��������� ���� ������ �� ���������� ...')
            ENDIF

          ENDIF

      ENDCASE

    ENDIF

  ENDIF
ENDIF

SELECT(sel)
RETURN


**************************************************** PROCEDURE read_sum *********************************************

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
    CASE BAR() = 1
      sel_val_card = 'RUR'

    CASE BAR() = 3
      sel_val_card = 'USD'

    CASE BAR() = 5
      sel_val_card = 'EUR'

  ENDCASE

ELSE
  RETURN
ENDIF

DO WHILE .T.

  SELECT MAX(LEN(ALLTRIM(fio_rus))) AS len_fio ;
    FROM acc_strah_sum ;
    WHERE ALLTRIM(val_card) == ALLTRIM(sel_val_card) ;
    INTO CURSOR sel_len_fio

  IF _TALLY <> 0

    SELECT DISTINCT SPACE(1) + PADR(ALLTRIM(fio_rus),sel_len_fio.len_fio) + ' - ' + ALLTRIM(val_card) + ' - ' + ALLTRIM(card_nls) AS fio_rus ;
      FROM acc_strah_sum ;
      WHERE stavka <> 0 AND ALLTRIM(val_card) == ALLTRIM(sel_val_card) ;
      INTO CURSOR tov ;
      ORDER BY 1

* BROWSE WINDOW brows

    DO FORM (put_scr) + ('selfio_nls.scx')

    IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

      SELECT DISTINCT data_pak AS data ;
        FROM acc_strah_sum ;
        WHERE ALLTRIM(card_nls) == sel_card_nls ;
        INTO CURSOR tov ;
        ORDER BY 1 DESC

      IF _TALLY <> 0

        DO FORM (put_scr) + ('seldata1.scx')

        IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

          SELECT acc_strah_sum
          SET ORDER TO fio
          SET FILTER TO data_pak = dt1 AND ALLTRIM(card_nls) == sel_card_nls
          GO TOP

          DEFINE WINDOW brow_proz  AT 0,0 SIZE  colvo_rows - 4, colvo_cols - 30 FONT "Courier New Cyr", 10  STYLE 'B' ;
            NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
          MOVE WINDOW brow_proz CENTER

          SET CURSOR ON

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
            sum_strah:H=' ����� ���������� �������� :':P='999 999.99',;
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

          REPLACE ALL sum_viplat WITH sum_dobav

          RELEASE WINDOW brow_proz

          SELECT acc_strah_sum
          SET FILTER TO
          SET ORDER TO card_data
          GO TOP

          SET CURSOR OFF

        ENDIF

      ELSE
        =soob('��������! � ������� ������������ ��������� �� ����������.')
      ENDIF

    ELSE
      EXIT
    ENDIF

  ELSE
    =soob('��������! � ������� ������������ ��������� �� ����������.')
    EXIT
  ENDIF

ENDDO

SELECT(sel)
RETURN


*************************************************************** PROCEDURE vid_kvartal_1 *****************************************************************

PROCEDURE vid_kvartal_1

DO CASE
  CASE BETWEEN(num_mesaz,1,3) = .T.
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
        vvod_data_end = IIF(colvo_den_god = 365, CTOD('28.02.' + num_god), IIF(colvo_den_god = 366, CTOD('29.02.' + num_god), DATE()))

      CASE num_mesaz = 3
        num_mesaz_kvartal = 3
        vvod_data_home = CTOD('01.03.' + num_god)
        vvod_data_end = CTOD('31.03.' + num_god)

    ENDCASE

  CASE BETWEEN(num_mesaz,4,6) = .T.
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

  CASE BETWEEN(num_mesaz,7,9) = .T.
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

  CASE BETWEEN(num_mesaz,10,12) = .T.
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


*************************************************************** PROCEDURE vid_kvartal_2 *****************************************************************

PROCEDURE vid_kvartal_2

DO CASE
  CASE BETWEEN(num_mesaz,1,3) = .T.
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

  CASE BETWEEN(num_mesaz,4,6) = .T.
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

  CASE BETWEEN(num_mesaz,7,9) = .T.
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

  CASE BETWEEN(num_mesaz,10,12) = .T.
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


*************************************************************** PROCEDURE sum_proz ********************************************************************

PROCEDURE sum_proz
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

put_bar = 1

num_mesaz = MONTH(DATE())
num_god = ALLTRIM(STR(YEAR(DATE())))

DO vid_kvartal_1 IN Raschet_proz_42308

DO FORM (put_scr) + ('vvod_data_proz.scx')

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  vvod_data_pak = MaxDataMes(vvod_data_end)

  num_mesaz = MONTH(vvod_data_end)
  num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

  DO vid_kvartal_2 IN Raschet_proz_42308

  SELECT branch, ref_client, ref_card, fio_rus, data_proz, card_nls, val_card, stavka, sum_strah, num_42308, num_47411, num_dog, data_dog ;
    FROM acc_strah ;
    WHERE data_proz <= vvod_data_end AND stavka <> 0 AND sum_strah <> 0 AND pr = .T. ;
    INTO CURSOR sel_acc_strah ;
    ORDER BY fio_rus

* BROWSE WINDOW brows

  IF _TALLY <> 0
    put_bar = 1
    DO raschet IN Raschet_proz_42308
  ELSE
    =err('��������! � ���� ������ ������� �� ����������.')
  ENDIF

ENDIF

SELECT(sel)
RETURN


**************************************************************** PROCEDURE vibor_proz *******************************************************************

PROCEDURE vibor_proz
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

vvod_data_pak = MaxDataMes(DATE())
put_bar = 1

num_mesaz = MONTH(DATE())
num_god = ALLTRIM(STR(YEAR(DATE())))

DO vid_kvartal_1 IN Raschet_proz_42308

DO WHILE .T.

  DO FORM (put_scr) + ('poisk_client.scx')

  SELECT MAX(LEN(ALLTRIM(fio_rus))) AS len_fio ;
    FROM acc_strah ;
    WHERE LEFT(ALLTRIM(fio_rus), len_client) == fam_client ;
    INTO CURSOR sel_len_fio

  SELECT DISTINCT SPACE(1) + PADR(ALLTRIM(fio_rus),sel_len_fio.len_fio) + ' - ' + ALLTRIM(val_card) + ' - ' + ALLTRIM(card_nls) AS fio_rus ;
    FROM acc_strah ;
    WHERE stavka <> 0 AND sum_strah <> 0 AND LEFT(ALLTRIM(fio_rus), len_client) == fam_client ;
    INTO CURSOR tov ;
    ORDER BY 1

* BROWSE WINDOW brows

  IF _TALLY <> 0

    DO FORM (put_scr) + ('selfio_nls.scx')

    IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

      DO FORM (put_scr) + ('vvod_data_proz_vib.scx')

      IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

        num_mesaz = MONTH(vvod_data_end)
        num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

        DO vid_kvartal_2 IN Raschet_proz_42308

        SELECT branch, ref_client, ref_card, fio_rus, data_proz, card_nls, val_card,;
          vvod_stavka AS stavka, vvod_sum_strah AS sum_strah, num_42308, num_47411, num_dog, data_dog ;
          FROM acc_strah ;
          WHERE data_proz <= vvod_data_end AND ALLTRIM(card_nls) == sel_card_nls ;
          INTO CURSOR sel_acc_strah ;
          ORDER BY fio_rus

* BROWSE WINDOW brows

        IF _TALLY <> 0
          DO raschet IN Raschet_proz_42308
        ELSE
          =err('��������! � ���� ������ ������� �� ����������.')
        ENDIF

      ENDIF

    ELSE
      EXIT
    ENDIF

  ELSE
    =err('��������! � ���� ������ ������� �� ����������.')
    EXIT
  ENDIF

ENDDO

SELECT(sel)
RETURN


******************************************************************** PROCEDURE raschet *****************************************************************

PROCEDURE raschet

sum_strah_prev = 0
ref_ost_prev = SPACE(10)

SELECT kurs_val
SET ORDER TO data_kurs

SELECT acc_strah_dt
SET ORDER TO card_data

SELECT acc_strah_sum
SET ORDER TO card_data

SELECT sel_acc_strah
SCAN

  SELECT calendar

* ������������ ������� calendar.dbf � ��������� ��������� ���: ������ ��������� vvod_data_home � ����� ��������� vvod_data_end
* ����� ����������� �������: ���� ����� �������� ����� ��������� ���� ������ ��������� vvod_data_home,
* �� ���� vvod_data_home ������ ��� ���� ������ ������� ��������� sel_acc_strah.data_proz, �� �� ������ ���������
* ������� ���� ������ ������� ��������� sel_acc_strah.data_proz

  SCAN FOR IIF(vvod_data_home < sel_acc_strah.data_proz,;
      BETWEEN(calendar.data_home, sel_acc_strah.data_proz, vvod_data_end),;
      BETWEEN(calendar.data_home, vvod_data_home, vvod_data_end))

* ----------------------------------------------------------------------------------------------------- ����� ������ ---------------------------------------------------------------------------------------------------------------------- *

    SET NEAR ON

* ------------------------------------------------------------------------------------------- ���� ���� ������ � �������� ------------------------------------------------------------------------------------------------------- *

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

* ------------------------------------------------------------------------------------------------ ���� ���� ������ � ���� --------------------------------------------------------------------------------------------------------- *

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

* ---------------------------------------------------------------------------------- ����� � ������� � ������ ������� ��������� ����� ------------------------------------------------------------------------------ *

    SELECT calendar

    SCATTER MEMVAR FIELDS kod_usd, kurs_usd, kod_eur, kurs_eur

    m.kod_usd = kod_val_usd
    m.kurs_usd = poisk_kurs_usd

    m.kod_eur = kod_val_eur
    m.kurs_eur = poisk_kurs_eur

    GATHER MEMVAR FIELDS kod_usd, kurs_usd, kod_eur, kurs_eur

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

    rec = ALLTRIM(sel_acc_strah.card_nls) + DTOS(calendar.data_home) + DTOS(calendar.data_end)

    WAIT WINDOW '������ - ' + ALLTRIM(sel_acc_strah.fio_rus) + '  ���� ������� - ' + ALLTRIM(sel_acc_strah.card_nls) NOWAIT

    SELECT acc_strah_dt

    poisk = SEEK(rec)

    IF poisk = .T. AND acc_strah_dt.pr = .F.   && ������ ������ ��������������� �������, ���������� ��, �� �����������.

      LOOP

    ELSE

      IF poisk = .T.
        SCATTER MEMVAR MEMO
      ELSE
        SCATTER MEMVAR MEMO BLANK
        m.pr = .T.
      ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      m.branch = ALLTRIM(sel_acc_strah.branch)
      m.ref_client = ALLTRIM(sel_acc_strah.ref_client)
      m.ref_card = sel_acc_strah.ref_card

      m.fio_rus = ALLTRIM(sel_acc_strah.fio_rus)  && ��� �������

      m.card_nls = ALLTRIM(sel_acc_strah.card_nls)  && ����� ���������� �����
      m.val_card = ALLTRIM(sel_acc_strah.val_card)  && ��� �����

      m.num_dog = sel_acc_strah.num_dog  && ����� ��������
      m.data_dog = sel_acc_strah.data_dog  && ���� ������ ���������� ��������

      m.data_proz = sel_acc_strah.data_proz  && ���� ������ ���������� ���������
      m.dt_hom_end = data_home_end  && ���� ������ �������� �� ������� ���������� ������
      m.data_home = calendar.data_home  && ��������� ���� ��������� �������, �������� ����������
      m.data_end = calendar.data_end  && �������� ���� ��������� �������, �������� ����������
      m.data_pak = vvod_data_pak  && �������� ���� ��������� ������� ��� �������� ������� ���������, ������������� �������������

      DO CASE
        CASE SUBSTR(ALLTRIM(m.card_nls),6,3) == kod_val_rur
          m.kurs=1  && ���� ������
        CASE SUBSTR(ALLTRIM(m.card_nls),6,3) == kod_val_usd
          m.kurs=vvod_kurs_usd  && ���� ������
        CASE SUBSTR(ALLTRIM(m.card_nls),6,3) == kod_val_eur
          m.kurs=vvod_kurs_eur  && ���� ������
      ENDCASE

      m.kvartal = num_kvartal  && ����� ��������
      m.mes_kvart = num_mesaz_kvartal  && ����� ������ � ��������

      m.stavka = sel_acc_strah.stavka  && ���������� ������

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
      m.sum_strah = sel_acc_strah.sum_strah  && ����� ���������� ��������
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      IF m.sum_strah <> sum_strah_prev
        m.ref_ost = SYS(2015)
      ELSE
        m.ref_ost = ref_ost_prev
      ENDIF

      sum_strah_prev = m.sum_strah  && ����� ������� �� ���������� ����
      ref_ost_prev = m.ref_ost  && ����� ������� �� ���������� ����

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
      m.colvo_den = IIF(calendar.data_home < sel_acc_strah.data_proz,(calendar.data_end - sel_acc_strah.data_proz),calendar.colvo_den)
* --------------------------------------------------------------------------------------------- �������� �������� ������� ���������� ------------------------------------------------------------------------------------- *
      por_okrugl = 5 && ������� ���������� ��� ������� ROUND()
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      sum_dobav_1 = (m.sum_strah / colvo_den_god) * m.colvo_den

      sum_dobav_2 = ROUND(((sum_dobav_1 * sel_acc_strah.stavka) / 100), por_okrugl)

      m.sum_dobav = IIF(m.pr = .T., sum_dobav_2, m.sum_dobav)

      sum_rasch_1 = (m.sum_strah / colvo_den_god) * m.colvo_den

      sum_rasch_2 = ROUND(((sum_rasch_1 * m.stavka) / 100), por_okrugl)

      m.sum_rasch = IIF(m.pr = .T., sum_rasch_2, m.sum_rasch)

      m.sum_rasxod = IIF(m.pr = .T.,ROUND((m.sum_rasch * m.kurs), por_okrugl), m.sum_rasxod)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      DO CASE
        CASE INLIST(ALLTRIM(m.val_card) , 'USD', 'EUR')  && �������� �����, ���������� �������� ������ ������

          DO CASE
            CASE m.stavka > 9 AND SUBSTR(ALLTRIM(m.card_nls),1,5) == '42301'  && ��� ����������

* ���� � ������� ������ ������ 9.0 %, �� � ���� ����� ����� ���������� ����� � ������� 35%
* ������� ������� ����� �� ��� ������ = sum_nalog_1
* ����� ������� ����� �� ������ 9.0 % = sum_nalog_2
* ����� ������� ������� (sum_nalog_1 - sum_nalog_2)
* � ����� � ��� ���������� ����� � ������� 35% (((sum_nalog_1 - sum_nalog_2) * 35) / 100) = sum_nalog_3

              sum_nalog_1 = (((m.sum_strah / colvo_den_god) * m.colvo_den) * m.stavka) / 100

              sum_nalog_2 = (((m.sum_strah / colvo_den_god) * m.colvo_den) * 9) / 100

              sum_nalog_3 = ROUND((((sum_nalog_1 - sum_nalog_2) * 35) / 100), por_okrugl)

              m.sum_nalog = IIF(m.pr = .T., sum_nalog_3, m.sum_nalog)

            CASE SUBSTR(ALLTRIM(m.card_nls),1,5) == '42601'  && ��� �� ����������

* ���� ������ �� �������� , �� � ���� ����� ����� ���������� ����� � ������� 30%
* ������� ������� ����� �� ��� ������ = sum_nalog_1
* ����� ����� � ��� ���������� ����� � ������� 30% ((sum_nalog_1 * 30) / 100) = sum_nalog_2

              sum_nalog_1 = (((m.sum_strah / colvo_den_god) * m.colvo_den) * m.stavka) / 100

              sum_nalog_2 = ROUND(((sum_nalog_1 * 30) / 100), por_okrugl)

              m.sum_nalog = IIF(m.pr = .T., sum_nalog_2, m.sum_nalog)

            OTHERWISE

              m.sum_nalog = 0

          ENDCASE

        CASE INLIST(ALLTRIM(m.val_card), 'RUR')  && �������� ����� �����, ������ ������ �� ������������

          m.sum_nalog = 0

      ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      m.sum_viplat = ROUND((m.sum_rasch - m.sum_nalog), por_okrugl)

      m.data = DATETIME()
      m.ima_pk = SYS(0)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      IF poisk = .T.
        GATHER MEMVAR MEMO
      ELSE
        INSERT INTO acc_strah_dt FROM MEMVAR
      ENDIF

    ENDIF

    SELECT calendar
  ENDSCAN

* -------------------------------------------------------------------------------------- �������� �������� ������� ���������� ---------------------------------------------------------------------------------------------- *
  por_okrugl_det = 4  && ������� ���������� ��� ������� ROUND()
* ---------------------------------------------------------- ���������� ������� ������ �� ������� ��������� �� ���� � ����������� � ������� --------------------------------------------------------- *

  SELECT branch, ref_client, ref_card, val_card, fio_rus, card_nls, data_proz, dt_hom_end,;
    IIF(vvod_data_home < data_proz, data_proz, vvod_data_home) AS data_home,;
    vvod_data_end AS data_end, data_pak,;
    kurs, kvartal, mes_kvart, stavka, sum_strah,;
    SUM(colvo_den) AS den_mes,;
    ROUND(SUM(sum_dobav), por_okrugl_det) AS sum_dobav,;
    ROUND(SUM(sum_rasch), por_okrugl_det) AS sum_rasch,;
    ROUND(SUM(sum_rasxod), por_okrugl_det) AS sum_rasxod,;
    ROUND(SUM(sum_nalog), por_okrugl_det) AS sum_nalog,;
    ROUND(SUM(sum_viplat), por_okrugl_det) AS sum_viplat,;
    num_dog, data_dog, data, ima_pk ;
    FROM acc_strah_dt ;
    WHERE ALLTRIM(card_nls) == ALLTRIM(sel_acc_strah.card_nls) AND BETWEEN(data_home, vvod_data_home, vvod_data_end) ;
    INTO CURSOR sel_proz_sum ;
    GROUP BY card_nls, sum_strah, ref_ost ;
    ORDER BY card_nls, data_home, data_end

* BROWSE WINDOW brows

* INTO CURSOR sel_proz_sum ;
* INTO TABLE (put_dbf)+('sel_proz_sum.dbf') ;

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  SELECT sel_proz_sum

  SCAN
    rec = ALLTRIM(sel_proz_sum.card_nls) + DTOS(sel_proz_sum.data_home) + DTOS(sel_proz_sum.data_end)

    SELECT acc_strah_sum

    poisk = SEEK(rec)

    IF poisk = .T. AND acc_strah_dt.pr = .F.   && ������ ������ ��������������� �������, ���������� ��, �� �����������.

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

      m.fio_rus = ALLTRIM(sel_proz_sum.fio_rus)  && ��� �������

      m.card_nls = ALLTRIM(sel_proz_sum.card_nls)  && ����� ���������� �����
      m.val_card = ALLTRIM(sel_proz_sum.val_card)  && ��� �����

      m.num_dog  =  sel_proz_sum.num_dog  && ����� ��������
      m.data_dog  =  sel_proz_sum.data_dog  && ���� ������ ���������� ��������

      m.data_proz = sel_proz_sum.data_proz  && ���� ������ ���������� ���������
      m.dt_hom_end = sel_proz_sum.dt_hom_end  && ���� ������ �������� �� ������� ���������� ������
      m.data_home = sel_proz_sum.data_home  && ��������� ���� ��������� �������, �������� ����������
      m.data_end = sel_proz_sum.data_end  && �������� ���� ��������� �������, �������� ����������
      m.data_pak = sel_proz_sum.data_pak  && �������� ���� ��������� ������� ��� �������� ������� ���������, ������������� �������������

      DO CASE
        CASE SUBSTR(ALLTRIM(m.card_nls),6,3) == kod_val_rur
          m.kurs = 1  && ���� ������
        CASE SUBSTR(ALLTRIM(m.card_nls),6,3) == kod_val_usd
          m.kurs = vvod_kurs_usd  && ���� ������
        CASE SUBSTR(ALLTRIM(m.card_nls),6,3) == kod_val_eur
          m.kurs = vvod_kurs_eur  && ���� ������
      ENDCASE

      m.kvartal = sel_proz_sum.kvartal  && ����� ��������
      m.mes_kvart = sel_proz_sum.mes_kvart  && ����� ������ � ��������

      m.stavka = sel_proz_sum.stavka  && ���������� ������
      m.sum_strah = sel_proz_sum.sum_strah  && ����� ���������� ��������

      m.den_mes = sel_proz_sum.den_mes

      m.sum_dobav = sel_proz_sum.sum_dobav

* ------------------------------------------------------------------------------------- �������� �������� ������� ���������� ---------------------------------------------------------------------------------------------- *
      por_okrugl_sum = 2  && ������� ���������� ��� ������� ROUND()
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      m.sum_dobav = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_dobav, por_okrugl_sum), m.sum_dobav)
      m.sum_rasch = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_rasch, por_okrugl_sum), m.sum_rasch)
      m.sum_rasxod = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_rasxod, por_okrugl_sum), m.sum_rasxod)
      m.sum_nalog = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_nalog, por_okrugl_sum), m.sum_nalog)
      m.sum_viplat = IIF(m.pr = .T., ROUND(sel_proz_sum.sum_viplat, por_okrugl_sum), m.sum_viplat)

* ------------------------------------------------------------- ��������� ������ ��������� �� ���� � ����� �� ����� �� ������ ���������� -------------------------------------------------------------- *
      por_okrugl_prover = 2  && ������� ���������� ��� ������� ROUND()
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      prover_sum_dobav = ROUND(((((m.sum_strah / colvo_den_god) * m.den_mes) * m.stavka) / 100), por_okrugl_prover)  && ����� ��������� ��� ��������

      m.sum_dobav = IIF(m.sum_dobav = prover_sum_dobav, m.sum_dobav, prover_sum_dobav)

      prover_sum_rasch = ROUND(((((m.sum_strah / colvo_den_god) * m.den_mes) * m.stavka) / 100), por_okrugl_prover)  && ����� ��������� ��� ��������

      m.sum_rasch = IIF(m.sum_rasch = prover_sum_rasch, m.sum_rasch, prover_sum_rasch)

      prover_sum_rasxod = ROUND((prover_sum_rasch * m.kurs), por_okrugl_prover)

* m.sum_rasxod = IIF(m.sum_rasxod = prover_sum_rasxod, m.sum_rasxod, prover_sum_rasxod)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      DO CASE
        CASE INLIST(ALLTRIM(m.val_card), 'USD', 'EUR')  && �������� �����, ���������� �������� ������ ������

          DO CASE
            CASE m.stavka > 9 AND SUBSTR(ALLTRIM(m.card_nls), 1, 5) == '42301'  && ��� ����������

* ���� � ������� ������ ������ 9.0 %, �� � ���� ����� ����� ���������� ����� � ������� 35%
* ������� ������� ����� �� ��� ������ = sum_nalog_1
* ����� ������� ����� �� ������ 9.0 % = sum_nalog_2
* ����� ������� ������� (sum_nalog_1 - sum_nalog_2)
* � ����� � ��� ���������� ����� � ������� 35% (((sum_nalog_1 - sum_nalog_2) * 35) / 100) = sum_nalog_3

              sum_nalog_1 = (((m.sum_strah / colvo_den_god) * m.colvo_den) * m.stavka) / 100

              sum_nalog_2 = (((m.sum_strah / colvo_den_god) * m.colvo_den) * 9) / 100

              prover_sum_nalog = ROUND((((sum_nalog_1 - sum_nalog_2) * 35) / 100), por_okrugl_prover)  && ����� ������

            CASE SUBSTR(ALLTRIM(m.card_nls), 1, 5) == '42601'  && ��� �� ����������

* ���� ������ �� �������� , �� � ���� ����� ����� ���������� ����� � ������� 30%
* ������� ������� ����� �� ��� ������ = sum_nalog_1
* ����� ����� � ��� ���������� ����� � ������� 30% ((sum_nalog_1 * 30) / 100) = sum_nalog_2

              sum_nalog_1 = (((m.sum_strah / colvo_den_god) * m.colvo_den) * m.stavka) / 100

              prover_sum_nalog = ROUND(((sum_nalog_1 * 30) / 100), por_okrugl_prover)  && ����� ������

            OTHERWISE

              m.prv_nalog = 0
              prover_sum_nalog = 0

          ENDCASE

        CASE INLIST(ALLTRIM(m.val_card), 'RUR')  && �������� ����� �����, ������ ������ �� ������������

          m.prv_nalog = 0
          prover_sum_nalog = 0

      ENDCASE

      m.sum_nalog = IIF(m.sum_nalog = prover_sum_nalog, m.sum_nalog, prover_sum_nalog)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      prover_sum_viplat = prover_sum_rasch - prover_sum_nalog  && ����� � �������

      m.sum_viplat = IIF(m.sum_viplat = prover_sum_viplat, m.sum_viplat, prover_sum_viplat)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      m.data = sel_proz_sum.data
      m.ima_pk = sel_proz_sum.ima_pk

      IF poisk = .T.
        GATHER MEMVAR MEMO
      ELSE
        INSERT INTO acc_strah_sum FROM MEMVAR
      ENDIF

    ENDIF

    SELECT sel_proz_sum
  ENDSCAN

  SELECT sel_acc_strah
ENDSCAN

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

WAIT CLEAR

SELECT acc_strah_sum
SET ORDER TO fio
SET FILTER TO

RETURN


************************************************************************* PROCEDURE sel_data ***********************************************************

PROCEDURE sel_data
sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

vvod_data_pak = MaxDataMes(DATE())

num_mesaz = MONTH(DATE())
num_god = ALLTRIM(STR(YEAR(DATE())))

DO vid_kvartal_1 IN Raschet_proz_42308

DO FORM (put_scr) + ('sel_data_proz.scx')

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  num_mesaz = MONTH(vvod_data_end)
  num_god = ALLTRIM(STR(YEAR(vvod_data_end)))

  DO vid_kvartal_2 IN Raschet_proz_42308

  SELECT * ;
    FROM acc_strah ;
    WHERE pr = .F. ;
    INTO CURSOR sel_cln_det

  IF _TALLY = 0

    DO sel_data_det IN Raschet_proz_42308

  ELSE

    WAIT WINDOW AT 5,20 NOWAIT ;
      PADC('��������! � ������������ ������ ������� �������� �� �������� � ������� ����������� ���������.',100) + CHR(13) + ;
      PADC('����� �� ��� ����� ������ � ������������ ��� � ����������� ���� ... ',100)

    text1 = ' ������� ������ � ������������ ���������� �������� '
    text2 = ' ������� ������ � ������������� ���������� �������� '
    l_bar = 3
    =popup_9(text1,text2,text3,text4,l_bar)
    ACTIVATE POPUP vibor
    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC
      DO CASE
        CASE BAR() = 1  && ������� ������ � ������������ ���������� ��������

          DO sel_data_det IN Raschet_proz_42308

        CASE BAR() = 3  && ������� ������ � ������������� ���������� ��������

          DO sel_data_sum IN Raschet_proz_42308

      ENDCASE

    ENDIF

  ENDIF

  IF _TALLY <> 0

    SELECT sel_prozent

    DO CASE
      CASE bar_num = 11
        DO brows_proz IN Raschet_proz_42308

      CASE bar_num = 13
        DO prn_proz IN Raschet_proz_42308

    ENDCASE

  ELSE
    =err('��������! �� ��������� ���� ������ �� ���������� ...')
  ENDIF

ENDIF

SELECT(sel)
RETURN


************************************************************** PROCEDURE sel_data_det *******************************************************************

PROCEDURE sel_data_det

SELECT B.branch, B.ref_client, B.ref_card, B.val_card, B.fio_rus, B.num_dog, B.data_dog, B.data_proz, B.kurs,;
  A.card_nls, IIF(SUBSTR(A.card_nls,1,5) == '42301', '1', IIF(SUBSTR(A.card_nls,1,5) == '42601', '2', '9')) AS grup,;
  B.kvartal, B.mes_kvart, B.stavka, B.sum_strah,;
  IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
  IIF(vvod_data_home < B.data_proz, B.data_proz, B.data_home) AS data_home, B.data_end AS data_end, B.den_mes,;
  B.sum_dobav, B.sum_rasch, B.sum_rasxod, B.sum_nalog, B.sum_viplat,;
  A.num_42308, A.num_47411 ;
  FROM acc_strah A, acc_strah_sum B ;
  WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND A.sum_strah <> 0 ;
  INTO CURSOR sel_prozent ;
  ORDER BY B.branch, B.val_card, grup, B.fio_rus, A.card_nls, data_home

* BROWSE WINDOW brows

* INTO CURSOR sel_prozent ;
* INTO TABLE (put_dbf) + ('sel_prozent.dbf') ;

RETURN


******************************************************************** PROCEDURE sel_data_sum ************************************************************

PROCEDURE sel_data_sum

SELECT B.branch, B.ref_client, B.ref_card, B.val_card, B.fio_rus, B.num_dog, B.data_dog, B.data_proz, B.kurs,;
  A.card_nls, IIF(SUBSTR(A.card_nls,1,5) == '42301', '1', IIF(SUBSTR(A.card_nls,1,5) == '42601', '2', '9')) AS grup,;
  B.kvartal, B.mes_kvart, A.stavka, A.sum_strah,;
  IIF(B.dt_hom_end < B.data_proz, B.data_proz, B.dt_hom_end) AS dt_hom_end,;
  IIF(vvod_data_home < B.data_proz, B.data_proz, vvod_data_home) AS data_home, vvod_data_end AS data_end,;
  SUM(B.den_mes) AS den_mes,;
  SUM(B.sum_dobav) AS sum_dobav,;
  SUM(B.sum_rasch) AS sum_rasch,;
  SUM(B.sum_rasxod) AS sum_rasxod,;
  SUM(B.sum_nalog) AS sum_nalog,;
  SUM(B.sum_viplat) AS sum_viplat,;
  A.num_42308, A.num_47411 ;
  FROM acc_strah A, acc_strah_sum B ;
  WHERE ALLTRIM(A.card_nls) == ALLTRIM(B.card_nls) AND BETWEEN(B.data_end, vvod_data_home, vvod_data_end) AND A.sum_strah <> 0 ;
  INTO CURSOR sel_prozent ;
  GROUP BY B.card_nls ;
  ORDER BY B.branch, B.val_card, grup, B.fio_rus, data_home

* BROWSE WINDOW brows

* INTO CURSOR sel_prozent ;
* INTO TABLE (put_dbf) + ('sel_prozent.dbf') ;

RETURN


**************************************************************** PROCEDURE brows_proz ******************************************************************

PROCEDURE brows_proz

DEFINE WINDOW brow_proz  AT 0,0 SIZE  colvo_rows - 2, colvo_cols - 35 FONT "Courier New Cyr", 10  STYLE 'B' ;
  NOFLOAT NOCLOSE NOMINIMIZE DOUBLE COLOR RGB(0,0,0,255,255,255)
MOVE WINDOW brow_proz CENTER

EDIT FIELDS ;
  branch:H = ' ������ :',;
  ref_client:H = ' ������ :',;
  ref_card:H = ' ����� :',;
  fio_rus:H = ' ������� ��� �������� :',;
  num_dog:H = ' ����� �������� :',;
  data_dog:H = ' ���� �������� :',;
  data_proz:H = ' ���� ����������� ������� :',;
  dt_hom_end:H = ' ������ ���� � �������� :',;
  data_home:H = ' ��������� ���� ������� :',;
  data_end:H = ' �������� ���� ������� :',;
  card_nls:H = ' ��������� ���� ��� :',;
  val_card:H = ' ��� ����� :',;
  kurs:H = ' ����:':P = '999.9999',;
  stavka:H = ' ������ :':P = '999.99',;
  sum_strah:H = ' ����� ���������� �������� :':P = '999 999.99',;
  kvartal:H = ' ����� �������� :',;
  mes_kvart:H = ' ����� ������ � �������� :',;
  den_mes:H = ' ���� � ������ :',;
  sum_dobav:H = ' ����� � ������������ :':P = '9 999.99',;
  sum_rasch:H = ' ����� ���������� �� ���� �������� �� ������� :':P = '9 999.99',;
  sum_rasxod:H = ' ����� ���������� �� ������� ����� :':P = '9 999.99',;
  sum_nalog:H = ' ����� ����������� ������ :':P = '9 999.99',;
  sum_viplat:H = ' ����� ��������� � ������� :':P = '9 999.99',;
  num_42308:H = ' ���� ���������� �������� :',;
  num_47411:H = ' ���� ����������� ��������� �� ���������� �������� :' ;
  TITLE ' �������� ������ �� ������������ ��������� ' ;
  WINDOW brow_proz

RELEASE WINDOW brow_proz

RETURN


********************************************************************** PROCEDURE prn_proz **************************************************************

PROCEDURE prn_proz

DO FORM (put_scr) + ('data_prn.scx')

IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

  text1 = '�������� ����� �� �����'
  text2 = '�������� ����� �� �������'
  l_bar = 3
  =popup_9(text1,text2,text3,text4,l_bar)
  ACTIVATE POPUP vibor
  RELEASE POPUPS vibor

  IF NOT LASTKEY() = 27  && ���� �� ���� ������ �� ������� ESC

    DO CASE
      CASE  sel_prozent.mes_kvart = 1 OR sel_prozent.mes_kvart = 2

        DO CASE
          CASE BAR() = 1
            REPORT FORM (put_frx) + ('proz_mesaz_42308.frx') NOEJECT PREVIEW

          CASE BAR() = 3

            DO CASE
              CASE par_prompt = .T.  && ��� ������ ������ �� ������ ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_mesaz_42308.frx') NOEJECT NOCONSOLE TO PRINTER PROMPT

              CASE par_prompt = .F.  && ��� ������ ������ �� ������ �� ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_mesaz_42308.frx') NOEJECT NOCONSOLE TO PRINTER

            ENDCASE

        ENDCASE

      CASE  sel_prozent.mes_kvart = 3

        DO CASE
          CASE BAR() = 1

            REPORT FORM (put_frx) + ('proz_kvart_42308.frx') NOEJECT PREVIEW

          CASE BAR() = 3

            DO CASE
              CASE par_prompt = .T.  && ��� ������ ������ �� ������ ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_kvart_42308.frx') NOEJECT NOCONSOLE TO PRINTER PROMPT

              CASE par_prompt = .F.  && ��� ������ ������ �� ������ �� ���������� ������ ������ ������������ ���������

                REPORT FORM (put_frx) + ('proz_kvart_42308.frx') NOEJECT NOCONSOLE TO PRINTER

            ENDCASE

        ENDCASE

    ENDCASE

  ENDIF

ENDIF

RETURN


*********************************************************************************************************************************************************





