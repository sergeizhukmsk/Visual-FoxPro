*******************************************************************************************************
*** ��������� ������������ ������� ������ �� ���������� ������� �� ����������� �������� �� �������� ***
*******************************************************************************************************

* ��������� � ����������� ���� nls_cred
* ���� ���������� ���������� �������� tip_term=mak_atm, mak_pos ���� ����������, �� ��������
* � ����������� �� ������ �����, ���� 47422810*******2****, ���� 47422840*******2****
* ���� ����� ���������� �������� �� ���� ����������, �� ��������
* � ����������� �� ������ �����, ���� 47422810*******1****, ���� 47422840*******1****
* ������ ������ �������� � ����������, ������� � ������ ��������� ����������� � � ��� ��������� ������
* �� ���������� ������� spravliz.dbf. ������ ������ � ��� ������� ����� ������� � ��������� ����������.
* ����� ���� ���������� '�������� � �������������� ����������� ����������'

* num_47422_no_rur_atm - ���������� �������� ����� ����� ��� ��������� �������� �� ������ ��� �� �������: ����� "RUR" ��������� ��������
* num_47422_no_usd_atm - ���������� �������� ����� ����� ��� ��������� �������� �� ������ ��� �� �������: ����� "USD" ��������� ��������
* num_47422_no_eur_atm - ���������� �������� ����� ����� ��� ��������� �������� �� ������ ��� �� �������: ����� "EUR" ��������� ��������

* num_47422_no_rur_pos - ���������� �������� ����� ����� ��� ��������� �������� �� ������ POS �� �������: ����� "RUR" ��������� ��������
* num_47422_no_usd_pos - ���������� �������� ����� ����� ��� ��������� �������� �� ������ POS �� �������: ����� "USD" ��������� ��������
* num_47422_no_eur_pos - ���������� �������� ����� ����� ��� ��������� �������� �� ������ POS �� �������: ����� "EUR" ��������� ��������

* num_47422_beznal_rur - ���������� �������� ����� ����� ��� ��������� �������� �� ����������� ��������� � ������
* num_47422_beznal_usd - ���������� �������� ����� ����� ��� ��������� �������� �� ����������� ��������� � ��������
* num_47422_beznal_eur - ���������� �������� ����� ����� ��� ��������� �������� �� ����������� ��������� � ����

* num_47422_no_rur_mag - ���������� �������� ����� ����� ��� ��������� �������� �� ������ TERM �� �������: ����� "RUR" ������ �������
* num_47422_no_usd_mag - ���������� �������� ����� ����� ��� ��������� �������� �� ������ TERM �� �������: ����� "USD" ������ �������
* num_47422_no_eur_mag - ���������� �������� ����� ����� ��� ��������� �������� �� ������ TERM �� �������: ����� "EUR" ������ �������


************************************************************** PROCEDURE start_liz ***********************************************************************

PROCEDURE start_liz

STORE SPACE(0) TO poisk_numliz_term

poisk_loc_term = .F.
poisk_exp_term = .F.

sel = SELECT()

SELECT loc_term
SET ORDER TO terminal

SELECT exp_term
SET ORDER TO terminal

SELECT (sel)

DO CASE
  CASE ALLTRIM(m.kod_oper) == '32N'  && ��� �������� �������� �� ������� ������ � ���������� �����

    DO start_spis_komis IN Docum_47422

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE ALLTRIM(m.kod_oper) == '52Z'  && ��� �������� ����� ���� ������� � ���������� ����� ������ ��������� �� ���������� ����

    DO CASE
      CASE pr_imp_70601 = .T.  && ������������ ����� �� ������� �� ������ ���������

        DO start_zakr_liz_new IN Docum_47422

      CASE pr_imp_70601 = .F.  && ������������ ����� �� ������� �� ������� ���������

        DO start_zakr_liz_str IN Docum_47422

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE INLIST(ALLTRIM(m.kod_oper), '52T', '52M', '52L', '52J', '52F', '52U', '52W')

    DO CASE
      CASE pr_imp_70601 = .T.  && ������������ ����� �� ������� �� ������ ���������

        DO start_tarif_pr_rasch_new IN Docum_47422

      CASE pr_imp_70601 = .F.  && ������������ ����� �� ������� �� ������� ���������

        DO start_tarif_pr_rasch_str IN Docum_47422

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE INLIST(ALLTRIM(m.kod_oper), '12A', '12B', '52R', '52O', '52S', '52K', '52P', '52N', '52D')

    DO CASE
      CASE pr_imp_70601 = .T.  && ������������ ����� �� ������� �� ������ ���������

        DO start_70107_52_new IN Docum_47422

      CASE pr_imp_70601 = .F.  && ������������ ����� �� ������� �� ������� ���������

        DO start_70107_52_str IN Docum_47422

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE (INLIST(LEFT(ALLTRIM(m.kod_oper), 2), '32') OR INLIST(ALLTRIM(m.kod_oper), '314', '31G', '31E', '31J', '31K')) AND NOT INLIST(ALLTRIM(m.kod_oper), '129', '52Z', '32N')

    DO CASE
      CASE flag_474_303 = .F.  && ����������� �������� 40817 - 30301

        IF INLIST(ALLTRIM(m.kod_oper), '314', '31G')
          DO start_30302_obr IN Docum_47422
        ELSE
          DO start_30301 IN Docum_47422
        ENDIF

      CASE flag_474_303 = .T.  && ����������� �������� 40817 - 47422

        DO CASE
          CASE EMPTY(ALLTRIM(m.name_term)) = .F.  && ������������ ���������� ����� ������� ��������� �������� �� ������

            SELECT loc_term  && ���������� ����� ���������: ���������, ������������, �������� ���������
            SET ORDER TO terminal

            poisk_loc_term = SEEK(ALLTRIM(m.name_term))  && ���������� ����� � ����������� ����� ���������

            DO CASE
              CASE poisk_loc_term = .T.  && ����� � ����������� ����� ��������� ������ �������

                poisk_city = IIF(LOWER(ALLTRIM(m.city)) $ LOWER(ALLTRIM(loc_term.city)), .T., .F.)  && ���������� �������� ������ � ��������� � ����������� ����� ���������

                IF poisk_city = .T.  && ������������ ������ ������������� �������� � ����������� ����� ���������
                  poisk_numliz_term = ALLTRIM(loc_term.numliz)
                ENDIF

                DO start_47422_loc_term IN Docum_47422

              CASE poisk_loc_term = .F.  && ����� � ����������� ����� ��������� ������ �� �������

                SELECT exp_term  && ���������� ����� ���������: ���������, ������������, �������� ���������

                poisk_exp_term = SEEK(ALLTRIM(m.name_term))  && ���������� ����� � ����������� ����� ���������

                DO CASE
                  CASE poisk_exp_term = .T.  && ����� � ����������� ����� ��������� ������ �������

                    DO start_47422_exp_term_yes IN Docum_47422

                  CASE poisk_exp_term = .F.  && ����� � ����������� ����� ��������� ������ �� �������

                    DO start_47422_exp_term_no IN Docum_47422

                ENDCASE

            ENDCASE

          CASE EMPTY(ALLTRIM(m.name_term)) = .T.  && ������������ ���������� ����� ������� ��������� �������� ������

            DO start_47422_exp_term_no IN Docum_47422

        ENDCASE

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE (INLIST(LEFT(ALLTRIM(m.kod_oper), 2), '20', '22') OR INLIST(ALLTRIM(m.kod_oper), '12A', '12B', '12C', '12S', '12T', '12V', '129', '312', '31M', '31N')) ;
      AND NOT INLIST(ALLTRIM(m.kod_oper), '206')

    DO CASE
      CASE flag_474_303 = .F.  && ����������� �������� 42301 - 30301

        IF INLIST(ALLTRIM(m.kod_oper), '225', '227')
          DO start_30302_obr IN Docum_47422
        ELSE
          DO start_30301 IN Docum_47422
        ENDIF

      CASE flag_474_303 = .T.  && ����������� �������� 42301 - 47422

        DO CASE
          CASE EMPTY(ALLTRIM(m.name_term)) = .F.  && ������������ ���������� ����� ������� ��������� �������� �� ������

            SELECT loc_term  && ���������� ����� ���������: ���������, ������������, �������� ���������
            GO TOP

            poisk_loc_term = SEEK(ALLTRIM(m.name_term))  && ���������� ����� � ����������� ����� ���������

            DO CASE
              CASE poisk_loc_term = .T.  && ����� � ����������� ����� ��������� ������ �������

                poisk_city = IIF(LOWER(ALLTRIM(m.city)) $ LOWER(ALLTRIM(loc_term.city)), .T., .F.)  && ���������� �������� ������ � ��������� � ����������� ����� ���������

                IF poisk_city = .T.  && ������������ ������ ������������� �������� � ����������� ����� ���������
                  poisk_numliz_term = ALLTRIM(loc_term.numliz)
                ENDIF

                DO start_47422_loc_term IN Docum_47422

              CASE poisk_loc_term = .F.  && ����� � ����������� ����� ��������� ������ �� �������

                SELECT exp_term  && ���������� ����� ���������: ���������, ������������, �������� ���������
                GO TOP

                poisk_exp_term = SEEK(ALLTRIM(m.name_term))  && ���������� ����� � ����������� ����� ���������

                DO CASE
                  CASE poisk_exp_term = .T.  && ����� � ����������� ����� ��������� ������ �������

                    DO start_47422_exp_term_yes IN Docum_47422

                  CASE poisk_exp_term = .F.  && ����� � ����������� ����� ��������� ������ �� �������

                    DO start_47422_exp_term_no IN Docum_47422

                ENDCASE

            ENDCASE

          CASE EMPTY(ALLTRIM(m.name_term)) = .T.  && ������������ ���������� ����� ������� ��������� �������� ������

            DO start_47422_exp_term_no IN Docum_47422

        ENDCASE

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE INLIST(ALLTRIM(m.kod_oper), '113', '115', '51P')

    DO start_beznal IN Docum_47422

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE INLIST(ALLTRIM(m.kod_oper), '206', '51A', '51W', '51P', '51N')

    DO start_beznal IN Docum_47422

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE ALLTRIM(m.kod_oper) == '516'

    DO start_70203 IN Docum_47422

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE ALLTRIM(m.kod_oper) == '110'

    DO start_20202 IN Docum_47422

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE ALLTRIM(m.kod_oper) == '995' OR ALLTRIM(m.kod_oper) == '999'

    DO CASE
      CASE tip_474 = .T.
        DO start_30301 IN Docum_47422

      CASE tip_474 = .F.
        DO start_30302 IN Docum_47422

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE pr_rasch_nds = .F. AND ALLTRIM(m.kod_oper) == '991'

    DO start_60322_70107 IN Docum_47422

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE pr_rasch_nds = .F. AND ALLTRIM(m.kod_oper) == '992'

    DO start_60322_60309 IN Docum_47422

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE pr_rasch_nds = .T. AND ALLTRIM(m.kod_oper) == '992'

    DO start_42301_60309 IN Docum_47422

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *
IF ALLTRIM(num_branch) <> '01'
  DO select_kod_oper IN Docum_47422
ENDIF
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

SELECT (sel)
RETURN


***************************************************************** PROCEDURE select_kod_oper *************************************************************

PROCEDURE select_kod_oper
sel = SELECT()

SELECT * ;
  FROM kod_oper ;
  WHERE ALLTRIM(kod_oper) == ALLTRIM(m.kod_oper) ;
  INTO CURSOR sel_kod_oper ;
  ORDER BY kod_oper

IF sel_kod_oper.pr_imp = .T.
  m.kod_name = IIF(EMPTY(ALLTRIM(sel_kod_oper.name_oper)) = .F., ALLTRIM(sel_kod_oper.name_oper), ALLTRIM(sel_kod_oper.kod_name))
ENDIF

SELECT(sel)
RETURN


********************************************************** PROCEDURE start_zakr_liz_new ******************************************************************

PROCEDURE start_zakr_liz_new

SELECT (sel)

DO CASE
  CASE ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52Z_rur)

  CASE ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52Z_usd)

  CASE ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52Z_eur)

ENDCASE

RETURN


*********************************************************** PROCEDURE start_zakr_liz_str ******************************************************************

PROCEDURE start_zakr_liz_str

SELECT (sel)

DO CASE
  CASE INLIST(ALLTRIM(m.branch), '10')

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'  && �������� �������� �����
        m.nls_cred = ALLTRIM(num_kassa810)

      CASE ALLTRIM(m.val_card) == 'USD'  && �������� �������� �����
        m.nls_cred = ALLTRIM(num_kassa840)

      CASE ALLTRIM(m.val_card) == 'EUR'  && �������� �������� �����
        m.nls_cred = ALLTRIM(num_kassa978)

    ENDCASE

  CASE NOT INLIST(ALLTRIM(m.branch), '10')

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'
        m.nls_cred = ALLTRIM(num_47422_beznal_rur)

      CASE ALLTRIM(m.val_card) == 'USD'
        m.nls_cred = ALLTRIM(num_47422_beznal_usd)

      CASE ALLTRIM(m.val_card) == 'EUR'
        m.nls_cred = ALLTRIM(num_47422_beznal_eur)

    ENDCASE

ENDCASE

RETURN


************************************************************* PROCEDURE start_spis_komis ****************************************************************

PROCEDURE start_spis_komis

SELECT (sel)

DO CASE
  CASE ALLTRIM(m.val_card) == 'RUR'

    m.nls_cred = ALLTRIM(num_32N_rur)

  CASE ALLTRIM(m.val_card) == 'USD'

    m.nls_cred = ALLTRIM(num_32N_usd)

  CASE ALLTRIM(m.val_card) == 'EUR'

    m.nls_cred = ALLTRIM(num_32N_eur)

ENDCASE

RETURN


************************************************************** PROCEDURE start_beznal *******************************************************************

PROCEDURE start_beznal

SELECT (sel)

DO CASE
  CASE ALLTRIM(m.val_card) == 'RUR'
    m.nls_deb = ALLTRIM(num_47422_beznal_rur)

  CASE ALLTRIM(m.val_card) == 'USD'
    m.nls_deb = ALLTRIM(num_47422_beznal_usd)

  CASE ALLTRIM(m.val_card) == 'EUR'
    m.nls_deb = ALLTRIM(num_47422_beznal_eur)

ENDCASE

RETURN


********************************************************** PROCEDURE start_tarif_pr_rasch_nds_T ***********************************************************

PROCEDURE start_tarif_pr_rasch_nds_t

STORE 0 TO sum_tarif_r, sum_doxod_r, sum_nds_r

******************** pr_rasch_nds = .T.   && ������� ��������� ����� ��� �� ����� ������ �� ����� �������, ������� ������� *****************************

SELECT kod_oper ;
  FROM kod_oper ;
  WHERE INLIST(ALLTRIM(m.kod_oper), '12A', '12B', '52D', '52J', '52L', '52M', '52N', '52P', '52O', '52S', '52K', '52R', '52W') AND ;
  ALLTRIM(kod_oper) == ALLTRIM(m.kod_oper) AND nds = .T. ;
  INTO CURSOR sel_kod_oper

poisk_kod_nds = _TALLY

IF poisk_kod_nds <> 0

  SELECT (sel)

  sum_tarif_r = m.sum_cred

  m.sum_cred = ROUND(((sum_tarif_r * 100) / (100 + summa_nds)), 2)

  m.sum_deb = m.sum_cred

  sum_doxod_r = m.sum_cred

ENDIF

SELECT (sel)

DO CASE
  CASE ALLTRIM(m.kod_oper) == '52J' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52J_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52J' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52J_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52J' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52J_eur)  && ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52L' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52L_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52L' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52L_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52L' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52L_eur)  && ALLTRIM(num_70107�_eur)


  OTHERWISE

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'

        m.nls_cred = ALLTRIM(num_70107�_rur)

      CASE ALLTRIM(m.val_card) == 'USD'

        m.nls_cred = ALLTRIM(num_70107�_usd)

      CASE ALLTRIM(m.val_card) == 'EUR'

        m.nls_cred = ALLTRIM(num_70107�_eur)

    ENDCASE

ENDCASE


***************************************************** PROCEDURE start_tarif_pr_rasch_new ******************************************************************

PROCEDURE start_tarif_pr_rasch_new

DO CASE
  CASE ALLTRIM(m.kod_oper) == '52T' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52T_rur)

  CASE ALLTRIM(m.kod_oper) == '52T' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52T_usd)

  CASE ALLTRIM(m.kod_oper) == '52T' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52T_eur)


  CASE ALLTRIM(m.kod_oper) == '52M' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52M_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52M' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52M_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52M' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52M_eur)  && ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52J' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52J_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52J' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52J_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52J' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52J_eur)  && ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52L' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52L_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52L' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52L_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52L' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52L_eur)  && ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52F' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52F_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52F' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52F_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52F' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52F_eur)  && ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52U' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_70107sms_rur)

  CASE ALLTRIM(m.kod_oper) == '52U' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_70107sms_usd)

  CASE ALLTRIM(m.kod_oper) == '52U' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_70107sms_eur)


  CASE ALLTRIM(m.kod_oper) == '52W' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52W_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52W' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52W_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52W' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52W_eur)  && ALLTRIM(num_70107�_eur)


  OTHERWISE

    m.nls_cred = ALLTRIM(num_60322)

ENDCASE

RETURN


****************************************************** PROCEDURE start_tarif_pr_rasch_str ******************************************************************

PROCEDURE start_tarif_pr_rasch_str

DO CASE
  CASE ALLTRIM(m.kod_oper) == '52T' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52T_rur)

  CASE ALLTRIM(m.kod_oper) == '52T' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52T_usd)

  CASE ALLTRIM(m.kod_oper) == '52T' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52T_eur)


  CASE ALLTRIM(m.kod_oper) == '52M' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52M' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52M' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_70107�_eur)


  CASE INLIST(ALLTRIM(m.kod_oper), '52L', '52J') AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_70107�_rur)

  CASE INLIST(ALLTRIM(m.kod_oper), '52L', '52J') AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_70107�_usd)

  CASE INLIST(ALLTRIM(m.kod_oper), '52L', '52J') AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52F' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52F' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52F' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52U' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_70107sms_rur)

  CASE ALLTRIM(m.kod_oper) == '52U' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_70107sms_usd)

  CASE ALLTRIM(m.kod_oper) == '52U' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_70107sms_eur)


  CASE ALLTRIM(m.kod_oper) == '52W' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52W' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52W' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_70107�_eur)


  OTHERWISE

    m.nls_cred = ALLTRIM(num_60322)

ENDCASE

RETURN


************************************************************** PROCEDURE start_42301_60309 **************************************************************

PROCEDURE start_42301_60309

SELECT (sel)

sum_nds_r = sum_tarif_r - sum_doxod_r

m.sum_cred = sum_nds_r

m.sum_deb = m.sum_cred

m.nls_cred = ALLTRIM(num_60309_���)

RETURN


************************************************************* PROCEDURE 60322_70107 *******************************************************************

PROCEDURE start_60322_70107

STORE 0 TO sum_tarif_r, sum_doxod_r, sum_nds_r

SELECT (sel)

m.nls_deb = ALLTRIM(nls_deb_cred)

sum_tarif_r = m.sum_cred

m.sum_cred = ROUND(((sum_tarif_r * 100) / (100 + summa_nds)), 2)

DO CASE
  CASE ALLTRIM(m.val_card) == 'RUR'

    m.nls_cred = ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.val_card) == 'USD'

    m.nls_cred = ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.val_card) == 'EUR'

    m.nls_cred = ALLTRIM(num_70107�_eur)

ENDCASE

m.sum_deb = m.sum_cred

sum_doxod_r = m.sum_cred

RETURN


************************************************************** PROCEDURE start_60322_60309 **************************************************************

PROCEDURE start_60322_60309

SELECT (sel)

m.nls_deb = ALLTRIM(nls_deb_cred)

sum_nds_r = sum_tarif_r - sum_doxod_r

m.sum_cred = sum_nds_r

m.nls_cred = ALLTRIM(num_60309_���)

m.sum_deb = m.sum_cred

RETURN


************************************************** PROCEDURE start_70107_52_new ************************************************************************

PROCEDURE start_70107_52_new

SELECT (sel)

DO CASE
  CASE ALLTRIM(m.kod_oper) == '52K' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52K_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52K' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52K_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52K' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52K_eur)  && ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52O' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52O_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52O' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52O_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52O' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52O_eur)  && ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52S' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52S_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52S' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52S_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52S' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52S_eur)  && ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52R' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52R_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52R' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52R_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52R' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52R_eur)  && ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52P' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52P_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52P' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52P_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52P' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52P_eur)  && ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52N' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52N_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52N' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52N_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52N' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52N_eur)  && ALLTRIM(num_70107�_eur)


  CASE ALLTRIM(m.kod_oper) == '52D' AND ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_52D_rur)  && ALLTRIM(num_70107�_rur)

  CASE ALLTRIM(m.kod_oper) == '52D' AND ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_52D_usd)  && ALLTRIM(num_70107�_usd)

  CASE ALLTRIM(m.kod_oper) == '52D' AND ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_52D_eur)  && ALLTRIM(num_70107�_eur)


  OTHERWISE  && ���� ��� ����� ������� - ������ ������ �� �������� � ��

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'

        m.nls_cred = ALLTRIM(num_70107�_rur)

      CASE ALLTRIM(m.val_card) == 'USD'

        m.nls_cred = ALLTRIM(num_70107�_usd)

      CASE ALLTRIM(m.val_card) == 'EUR'

        m.nls_cred = ALLTRIM(num_70107�_eur)

    ENDCASE

ENDCASE

RETURN


********************************************************** PROCEDURE start_70107_52_str *****************************************************************

PROCEDURE start_70107_52_str

SELECT (sel)

DO CASE
  CASE INLIST(ALLTRIM(m.kod_oper), '52K', '52O', '52S', '52R')

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'

        m.nls_cred = ALLTRIM(num_70107�_rur)

      CASE ALLTRIM(m.val_card) == 'USD'

        m.nls_cred = ALLTRIM(num_70107�_usd)

      CASE ALLTRIM(m.val_card) == 'EUR'

        m.nls_cred = ALLTRIM(num_70107�_eur)

    ENDCASE

  OTHERWISE  && ���� ��� ����� ������� - ������ ������ �� �������� � ��

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'

        m.nls_cred = ALLTRIM(num_70107�_rur)

      CASE ALLTRIM(m.val_card) == 'USD'

        m.nls_cred = ALLTRIM(num_70107�_usd)

      CASE ALLTRIM(m.val_card) == 'EUR'

        m.nls_cred = ALLTRIM(num_70107�_eur)

    ENDCASE

ENDCASE

RETURN


***************************************************************** PROCEDURE start_70107 *****************************************************************

PROCEDURE start_70107

SELECT (sel)

DO CASE
  CASE INLIST(ALLTRIM(m.kod_oper), '314', '324', '31E', '31G', '31J', '31K', '32G')  && ���� �������� ��������

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'
        m.nls_cred = ALLTRIM(num_70107_rur)

      CASE ALLTRIM(m.val_card) == 'USD'
        m.nls_cred = ALLTRIM(num_70107_usd)

      CASE ALLTRIM(m.val_card) == 'EUR'
        m.nls_cred = ALLTRIM(num_70107_eur)

    ENDCASE

  CASE INLIST(LEFT(ALLTRIM(m.kod_oper), 2), '12', '52') AND NOT INLIST(ALLTRIM(m.kod_oper), '129', '52Z')

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'

        m.nls_cred = ALLTRIM(num_70107�_rur)

      CASE ALLTRIM(m.val_card) == 'USD'

        m.nls_cred = ALLTRIM(num_70107�_usd)

      CASE ALLTRIM(m.val_card) == 'EUR'

        m.nls_cred = ALLTRIM(num_70107�_eur)

    ENDCASE

  OTHERWISE  && ���� ��� ����� ������� - ������ ������ �� �������� � ��

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'

        m.nls_cred = ALLTRIM(num_70107�_rur)

      CASE ALLTRIM(m.val_card) == 'USD'

        m.nls_cred = ALLTRIM(num_70107�_usd)

      CASE ALLTRIM(m.val_card) == 'EUR'

        m.nls_cred = ALLTRIM(num_70107�_eur)

    ENDCASE

ENDCASE

RETURN


************************************************************** PROCEDURE start_70107_obr ****************************************************************

PROCEDURE start_70107_obr

SELECT (sel)

m.nls_cred = ALLTRIM(m.nls_deb)

DO CASE
  CASE ALLTRIM(m.val_card) == 'RUR'
    m.nls_deb = ALLTRIM(num_70107_rur)

  CASE ALLTRIM(m.val_card) == 'USD'
    m.nls_deb = ALLTRIM(num_70107_usd)

  CASE ALLTRIM(m.val_card) == 'EUR'
    m.nls_deb = ALLTRIM(num_70107_eur)

ENDCASE

RETURN


**************************************************************** PROCEDURE start_70203 ******************************************************************

PROCEDURE start_70203

SELECT (sel)

DO CASE
  CASE ALLTRIM(m.val_card) == 'RUR'
    m.nls_deb = ALLTRIM(num_70203_rur_proz)

  CASE ALLTRIM(m.val_card) == 'USD'
    m.nls_deb = ALLTRIM(num_70203_usd_proz)

  CASE ALLTRIM(m.val_card) == 'EUR'
    m.nls_deb = ALLTRIM(num_70203_eur_proz)

ENDCASE

RETURN


**************************************************************** PROCEDURE start_20202 ******************************************************************

PROCEDURE start_20202

SELECT (sel)

DO CASE
  CASE ALLTRIM(m.val_card) == 'RUR'
    m.nls_deb = ALLTRIM(num_kassa810)

  CASE ALLTRIM(m.val_card) == 'USD'
    m.nls_deb = ALLTRIM(num_kassa840)

  CASE ALLTRIM(m.val_card) == 'EUR'
    m.nls_deb = ALLTRIM(num_kassa978)

ENDCASE

RETURN


************************************************************** PROCEDURE start_30302_obr ****************************************************************

PROCEDURE start_30302_obr

SELECT (sel)

m.nls_cred = ALLTRIM(m.nls_deb)

DO CASE
  CASE ALLTRIM(m.val_card) == 'RUR'
    m.nls_deb = ALLTRIM(num_30302_01_rur)

  CASE ALLTRIM(m.val_card) == 'USD'
    m.nls_deb = ALLTRIM(num_30302_01_usd)

  CASE ALLTRIM(m.val_card) == 'EUR'
    m.nls_deb = ALLTRIM(num_30302_01_eur)

ENDCASE

RETURN


**************************************************************** PROCEDURE start_30302 ******************************************************************

PROCEDURE start_30302

SELECT (sel)

DO CASE
  CASE ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_30302_01_rur)

  CASE ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_30302_01_usd)

  CASE ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_30302_01_eur)

ENDCASE

RETURN

**************************************************************** PROCEDURE start_30301 ******************************************************************

PROCEDURE start_30301

SELECT (sel)

DO CASE
  CASE ALLTRIM(m.val_card) == 'RUR'
    m.nls_cred = ALLTRIM(num_30301_01_rur)

  CASE ALLTRIM(m.val_card) == 'USD'
    m.nls_cred = ALLTRIM(num_30301_01_usd)

  CASE ALLTRIM(m.val_card) == 'EUR'
    m.nls_cred = ALLTRIM(num_30301_01_eur)

ENDCASE

RETURN


**************************************************** PROCEDURE start_47422_exp_term_yes *****************************************************************

PROCEDURE start_47422_exp_term_yes

SELECT (sel)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

DO CASE
  CASE INLIST(ALLTRIM(exp_term.term), 'A') AND INLIST(ALLTRIM(m.kod_oper), '207', '209', '225', '227', '324', '314', '31G', '32G')

    DO CASE
      CASE INLIST(ALLTRIM(m.kod_oper), '207', '209', '324', '32G')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_cred = ALLTRIM(num_47422_no_rur_atm)

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_cred = ALLTRIM(num_47422_no_usd_atm)

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_cred = ALLTRIM(num_47422_no_eur_atm)

        ENDCASE

      CASE INLIST(ALLTRIM(m.kod_oper), '227', '229', '314', '31G')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_deb = ALLTRIM(num_47422_no_rur_atm)

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_deb = ALLTRIM(num_47422_no_usd_atm)

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_deb = ALLTRIM(num_47422_no_eur_atm)

        ENDCASE

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE INLIST(ALLTRIM(exp_term.term), 'P', 'N') AND INLIST(ALLTRIM(m.kod_oper), '207', '209', '225', '227', '324', '314', '31G', '32G')

    DO CASE
      CASE INLIST(ALLTRIM(m.kod_oper), '207', '209', '324', '32G')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_cred = ALLTRIM(num_47422_no_rur_pos)

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_cred = ALLTRIM(num_47422_no_usd_pos)

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_cred = ALLTRIM(num_47422_no_eur_pos)

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE INLIST(ALLTRIM(m.kod_oper), '225')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_deb = ALLTRIM(num_47422_no_rur_mag)

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_deb = ALLTRIM(num_47422_no_usd_mag)

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_deb = ALLTRIM(num_47422_no_eur_mag)

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE INLIST(ALLTRIM(m.kod_oper), '227', '229', '314', '31G')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_deb = ALLTRIM(num_47422_no_rur_pos)

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_deb = ALLTRIM(num_47422_no_usd_pos)

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_deb = ALLTRIM(num_47422_no_eur_pos)

        ENDCASE

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE INLIST(ALLTRIM(exp_term.term), 'A', 'N') AND INLIST(ALLTRIM(m.kod_oper), '205')

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'
        m.nls_cred = ALLTRIM(num_47422_no_rur_mag)

      CASE ALLTRIM(m.val_card) == 'USD'
        m.nls_cred = ALLTRIM(num_47422_no_usd_mag)

      CASE ALLTRIM(m.val_card) == 'EUR'
        m.nls_cred = ALLTRIM(num_47422_no_eur_mag)

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE INLIST(ALLTRIM(exp_term.term), 'P') AND INLIST(ALLTRIM(m.kod_oper), '205')

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'
        m.nls_cred = ALLTRIM(num_47422_no_rur_pos)

      CASE ALLTRIM(m.val_card) == 'USD'
        m.nls_cred = ALLTRIM(num_47422_no_usd_pos)

      CASE ALLTRIM(m.val_card) == 'EUR'
        m.nls_cred = ALLTRIM(num_47422_no_eur_pos)

    ENDCASE

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

RETURN


**************************************************** PROCEDURE start_47422_exp_term_no *****************************************************************

PROCEDURE start_47422_exp_term_no

SELECT (sel)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

DO CASE
  CASE INLIST(ALLTRIM(m.kod_oper), '207', '209', '225', '227', '324', '314', '31G', '32G')

    DO CASE
      CASE INLIST(ALLTRIM(m.kod_oper), '207', '209', '324', '32G')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_cred = ALLTRIM(num_47422_no_rur_atm)

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_cred = ALLTRIM(num_47422_no_usd_atm)

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_cred = ALLTRIM(num_47422_no_eur_atm)

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE INLIST(ALLTRIM(m.kod_oper), '225')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_deb = ALLTRIM(num_47422_no_rur_mag)

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_deb = ALLTRIM(num_47422_no_usd_mag)

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_deb = ALLTRIM(num_47422_no_eur_mag)

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE INLIST(ALLTRIM(m.kod_oper), '227', '229', '314', '31G')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_deb = ALLTRIM(num_47422_no_rur_atm)

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_deb = ALLTRIM(num_47422_no_usd_atm)

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_deb = ALLTRIM(num_47422_no_eur_atm)

        ENDCASE

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE INLIST(ALLTRIM(m.kod_oper), '205')

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'
        m.nls_cred = ALLTRIM(num_47422_no_rur_mag)

      CASE ALLTRIM(m.val_card) == 'USD'
        m.nls_cred = ALLTRIM(num_47422_no_usd_mag)

      CASE ALLTRIM(m.val_card) == 'EUR'
        m.nls_cred = ALLTRIM(num_47422_no_eur_mag)

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE INLIST(ALLTRIM(m.kod_oper), '129')

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'
        m.nls_cred = ALLTRIM(num_47422_beznal_rur)

      CASE ALLTRIM(m.val_card) == 'USD'
        m.nls_cred = ALLTRIM(num_47422_beznal_usd)

      CASE ALLTRIM(m.val_card) == 'EUR'
        m.nls_cred = ALLTRIM(num_47422_beznal_eur)

    ENDCASE

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

RETURN


********************************************************** PROCEDURE start_47422_loc_term ***************************************************************

PROCEDURE start_47422_loc_term

SELECT (sel)

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

DO CASE
  CASE INLIST(ALLTRIM(loc_term.term), 'A') AND INLIST(ALLTRIM(m.kod_oper), '207', '209', '225', '227', '324', '314', '31G', '32G')

    DO CASE
      CASE INLIST(ALLTRIM(m.kod_oper), '207', '209', '324', '32G')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_cred = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_rur_atm))

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_cred = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_usd_atm))

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_cred = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_eur_atm))

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE INLIST(ALLTRIM(m.kod_oper), '225')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_deb = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_rur_mag))

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_deb = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_usd_mag))

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_deb = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_eur_mag))

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE INLIST(ALLTRIM(m.kod_oper), '227', '229', '314', '31G')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_deb = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_rur_atm))

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_deb = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_usd_atm))

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_deb = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_eur_atm))

        ENDCASE

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE INLIST(ALLTRIM(loc_term.term), 'P', 'N') AND INLIST(ALLTRIM(m.kod_oper), '207', '209', '225', '227', '324', '314', '31G', '32G')

    DO CASE
      CASE INLIST(ALLTRIM(m.kod_oper), '207', '209', '324', '32G')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_cred = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_rur_pos))

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_cred = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_usd_pos))

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_cred = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_eur_pos))

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE INLIST(ALLTRIM(m.kod_oper), '225')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_deb = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_rur_mag))

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_deb = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_usd_mag))

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_deb = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_eur_mag))

        ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

      CASE INLIST(ALLTRIM(m.kod_oper), '227', '229', '314', '31G')

        DO CASE
          CASE ALLTRIM(m.val_card) == 'RUR'
            m.nls_deb = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_rur_pos))

          CASE ALLTRIM(m.val_card) == 'USD'
            m.nls_deb = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_usd_pos))

          CASE ALLTRIM(m.val_card) == 'EUR'
            m.nls_deb = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_eur_pos))

        ENDCASE

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE INLIST(ALLTRIM(loc_term.term), 'A', 'N') AND INLIST(ALLTRIM(m.kod_oper), '205')

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'
        m.nls_cred = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_rur_mag))

      CASE ALLTRIM(m.val_card) == 'USD'
        m.nls_cred = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_usd_mag))

      CASE ALLTRIM(m.val_card) == 'EUR'
        m.nls_cred = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_eur_mag))

    ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

  CASE INLIST(ALLTRIM(loc_term.term), 'P') AND INLIST(ALLTRIM(m.kod_oper), '205')

    DO CASE
      CASE ALLTRIM(m.val_card) == 'RUR'
        m.nls_cred = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_rur_pos))

      CASE ALLTRIM(m.val_card) == 'USD'
        m.nls_cred = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_usd_pos))

      CASE ALLTRIM(m.val_card) == 'EUR'
        m.nls_cred = IIF(INLIST(ALLTRIM(m.branch), '02', '06', '07'), ALLTRIM(poisk_numliz_term), ALLTRIM(num_47422_yes_eur_pos))

    ENDCASE

ENDCASE

RETURN


**********************************************************************************************************************************************************





