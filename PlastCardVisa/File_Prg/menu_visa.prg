***********************************************************
*** ��������� ������������ ���� ��� ��������� ������� ***
***********************************************************

PROCEDURE start

* ================================================ ������������ �������� ���� ��������� ================================================== *

text_visa_1_1 = '�������� ������� �� ����������� �����'
text_visa_1_2 = '������ ���� � ���������� ��������� ������'
text_visa_1_3 = '�������� � ������������ ������� � ���'
text_visa_1_4 = '���������� ��������� �� ���������� �����'
text_visa_1_5 = '�������� ������ �� ��������� �����'
text_visa_1_6 = '�������� ������ � �������� ����'
text_visa_1_7 = '��������� � ��������� ���������'
text_visa_1_8 = '��������� ����������� ������'
text_visa_1_9 = '�������� ������������� ��� � "CARD-VISA"'
text_visa_1_10 = '������ � ����������� �������'
text_visa_1_11 = '���������� ������ � ����������'

len_visa_1_1 = TXTWIDTH(text_visa_1_1, 'Times New Roman', num_schrift_v, 'B')
len_visa_1_2 = TXTWIDTH(text_visa_1_2, 'Times New Roman', num_schrift_v, 'B')
len_visa_1_3 = TXTWIDTH(text_visa_1_3, 'Times New Roman', num_schrift_v, 'B')
len_visa_1_4 = TXTWIDTH(text_visa_1_4, 'Times New Roman', num_schrift_v, 'B')
len_visa_1_5 = TXTWIDTH(text_visa_1_5, 'Times New Roman', num_schrift_v, 'B')
len_visa_1_6 = TXTWIDTH(text_visa_1_6, 'Times New Roman', num_schrift_v, 'B')
len_visa_1_7 = TXTWIDTH(text_visa_1_7, 'Times New Roman', num_schrift_v, 'B')
len_visa_1_8 = TXTWIDTH(text_visa_1_8, 'Times New Roman', num_schrift_v, 'B')
len_visa_1_9 = TXTWIDTH(text_visa_1_9, 'Times New Roman', num_schrift_v, 'B')
len_visa_1_10 = TXTWIDTH(text_visa_1_10, 'Times New Roman', num_schrift_v, 'B')
len_visa_1_11 = TXTWIDTH(text_visa_1_11, 'Times New Roman', num_schrift_v, 'B')

l_bar = 26.0
row = IIF(l_bar<=25, l_bar, 25)

V_row = (WROWS()-row)/2 + 1
Lmenu = MAX(len_visa_1_1, len_visa_1_2, len_visa_1_3, len_visa_1_4, len_visa_1_5, len_visa_1_6, len_visa_1_7, len_visa_1_8, len_visa_1_9, len_visa_1_10, len_visa_1_11)
L_col = (WCOLS()-Lmenu)/2 - 0

DEFINE POPUP glmenu FROM V_row,L_col ;
  FONT 'Times New Roman', num_schrift_v  STYLE 'B' ;
  MARGIN COLOR SCHEME 2 MESSAGE text_mess
DEFINE BAR 1  OF glmenu PROMPT text_visa_1_1 ;
  MESSAGE '�p� ����p� ������� ������ ������������ ��������� �������� ����� �������� ������� �� �����'
DEFINE BAR 2  OF glmenu PROMPT '\-'
DEFINE BAR 3  OF glmenu PROMPT text_visa_1_2 SKIP FOR wes < 4 ;
  MESSAGE '�p� ����p� ������� ������ ������������ ��������� �������� �� ������ � ��������'
DEFINE BAR 4  OF glmenu PROMPT '\-'
DEFINE BAR 5  OF glmenu PROMPT text_visa_1_3 SKIP FOR wes < 4 ;
  MESSAGE '�p� ����p� ������� ������ ������������ ������� � ��������� ������ � ���'
DEFINE BAR 6  OF glmenu PROMPT '\-'
DEFINE BAR 7  OF glmenu PROMPT text_visa_1_4 SKIP FOR wes < 4 ;
  MESSAGE '�p� ����p� ������� ������ ������������ ��������� �������� �� ���������� ��������� �� ����������� �����'
DEFINE BAR 8  OF glmenu PROMPT '\-'
DEFINE BAR 9  OF glmenu PROMPT text_visa_1_5 SKIP FOR wes < 5 ;
  MESSAGE '�p� ����p� ������� ������ �p�������� ������ ������ �� ��������, ������ � �������� ������� �� ������'
DEFINE BAR 10 OF glmenu PROMPT '\-'
DEFINE BAR 11 OF glmenu PROMPT text_visa_1_6 SKIP FOR wes < 5 ;
  MESSAGE '�p� ����p� ������� ������ �p�������� �������� ������ �� ������ ��� �������� � �������� ����'
DEFINE BAR 12  OF glmenu PROMPT '\-'
DEFINE BAR 13 OF glmenu PROMPT text_visa_1_7 ;
  MESSAGE '�p� ����p� ������� ������ ������������ ��������� ��������� � ��������� ���������'
DEFINE BAR 14 OF glmenu PROMPT '\-'
DEFINE BAR 15 OF glmenu PROMPT text_visa_1_8 SKIP FOR wes < 5 ;
  MESSAGE '�p� ����p� ������� ������ �p�������� ��������� ����������� ������� ������ � �������'
DEFINE BAR 16 OF glmenu PROMPT '\-'
DEFINE BAR 17 OF glmenu PROMPT text_visa_1_9 SKIP FOR wes < 5 ;
  MESSAGE '�p� ����p� ������� ������ �p�������� �������� ������������� ��� � �� "CARD-Visa"'
DEFINE BAR 18 OF glmenu PROMPT '\-'
DEFINE BAR 19 OF glmenu PROMPT text_visa_1_10 ;
  MESSAGE '�p� ����p� ������� ������ �p�������� �������� ����������� �������'
DEFINE BAR 20 OF glmenu PROMPT '\-'
DEFINE BAR 21 OF glmenu PROMPT text_visa_1_11 ;
  MESSAGE '�p� ����p� ������� ������ �p�������� ����� � ��������� ����� �������'
ON SELECTION POPUP glmenu DO glmenu IN Visa


* ===================== ������������ ��������������� ���� ��� ������ �������� ���� "�������� ������� �� ����������� �����" ====================== *


************************************************************** DEFINE MENU document ********************************************************************

DEFINE MENU document COLOR SCHEME 2 ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B'

DEFINE PAD document OF document PROMPT '�������� �� �����' COLOR SCHEME 2
DEFINE PAD vedomostat OF document PROMPT '��������� ��������' COLOR SCHEME 2
DEFINE PAD vipiska OF document PROMPT '������� �� ������' COLOR SCHEME 2
DEFINE PAD analit_info OF document PROMPT '���������' COLOR SCHEME 2
DEFINE PAD otchet_zb OF document PROMPT '������ ��' COLOR SCHEME 2
DEFINE PAD servis_doc OF document PROMPT '�����������' COLOR SCHEME 2
DEFINE PAD quit OF document PROMPT '�����' COLOR SCHEME 2 ;
  MESSAGE '����� � ������� ����'

ON PAD document OF document ACTIVATE POPUP document
ON PAD vedomostat OF document ACTIVATE POPUP vedomostat
ON PAD vipiska OF document ACTIVATE POPUP vipiska
ON PAD analit_info OF document ACTIVATE POPUP analit_info
ON PAD otchet_zb OF document ACTIVATE POPUP otchet_zb
ON PAD servis_doc OF document ACTIVATE POPUP servis_doc
ON SELECTION PAD quit OF document DEACTIVATE MENU document


************************************************************ DEFINE POPUP document   *********************************************************************

DEFINE POPUP document  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF document PROMPT '����� �� ��������� �� �������� ������� �� �� "Transmaster"' ;
  MESSAGE '������������ ������ �� ��������� �� ������������ ������� �� ��������� ������ � �������� ������� �� �� "Transmaster"'
DEFINE BAR 2  OF document PROMPT '\-'
DEFINE BAR 3  OF document PROMPT '����� �� ��������� �� ���������� ������� �� �� "Transmaster"' ;
  MESSAGE '������������ ������ �� ��������� �� ���������� ������� �� �� �� ��������� ������ � �������� ������� �� �� "Transmaster"'
DEFINE BAR 4  OF document PROMPT '\-'
DEFINE BAR 5  OF document PROMPT '����� �� ��������� �� �������� ������� �� �� "CARD-Visa ������"' ;
  MESSAGE '������������ ������ �� ��������� �� ������������ ������� �� ��������� ������ � �������� ������� �� �� "CARD-Visa ������"'
DEFINE BAR 6  OF document PROMPT '\-'
DEFINE BAR 7  OF document PROMPT '����� �� ��������� �� ���������� ������� �� �� "CARD-Visa ������"' ;
  MESSAGE '������������ ������ �� ��������� �� ���������� ������� �� �� �� ��������� ������ � �������� ������� �� �� "CARD-Visa ������"'
DEFINE BAR 8  OF document PROMPT '\-'
DEFINE BAR 9  OF document PROMPT '����� � ������������� ��������� �� ����������� ���������' ;
  MESSAGE '����� � ������������� ��������� �� ����������� ���������, �� ���� ����� ������ ����� ���������� �� ���� ������ ����������'
DEFINE BAR 10 OF document PROMPT '\-'
DEFINE BAR 11 OF document PROMPT '������ ����� ����������� �������� �� �� "Transmaster"' ;
  MESSAGE '������������ ������ � ����������� ��������� �� ��������� ������ � �������� ������� �� �� "Transmaster"'
DEFINE BAR 12 OF document PROMPT '\-'
DEFINE BAR 13 OF document PROMPT '������ ����� ����������� �������� �� �� "CARD-Visa ������"' ;
  MESSAGE '������������ ������ � ����������� ��������� �� ��������� ������ � �������� ������� �� �� "CARD-Visa ������"'
DEFINE BAR 14 OF document PROMPT '\-'
DEFINE BAR 15 OF document PROMPT '�������� ��������������� ���������� �� ���� �� �� "Transmaster"' ;
  MESSAGE '������ ��������� ��������� �������� ����� ��������������� ���������� �� ��������� ���� �� �� "Transmaster"'
DEFINE BAR 16 OF document PROMPT '\-'
DEFINE BAR 17 OF document PROMPT '������� ���������� �� ������ �� ������� �� �� "Transmaster"' ;
  MESSAGE '����������� ���������� ����������� ���������� �� ������: ���������� �������, ���� �� ������, ���� �� �������, ����� ��������'
DEFINE BAR 18 OF document PROMPT '\-'
DEFINE BAR 19 OF document PROMPT '������������ ������ �� ����������� ����������� ��������' ;
  MESSAGE '����������� ������� ����������� ���������� �� ������ �� ������ ������ ��� ����������� 600 000.00 ���.'

ON SELECTION BAR 1   OF document DO start_rasx  IN Doc_rasx_slip
ON SELECTION BAR 3   OF document DO start_prix   IN Doc_prix_slip
ON SELECTION BAR 5   OF document DO start_rasx  IN Doc_rasx_mak
ON SELECTION BAR 7   OF document DO start_prix   IN Doc_prix_mak
ON SELECTION BAR 9   OF document DO start_konv IN Doc_konv_slip
ON SELECTION BAR 11  OF document DO start_opis   IN Doc_opis_slip
ON SELECTION BAR 13  OF document DO start_opis   IN Doc_opis_mak
ON SELECTION BAR 15  OF document DO start_del   IN Doc_del_slip
ON SELECTION BAR 17  OF document DO start IN Doc_filtr_slip
ON SELECTION BAR 19  OF document DO start IN Sel_legaliz_slip


*************************************************************** DEFINE POPUP vedomostat *****************************************************************

DEFINE POPUP vedomostat ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1 OF vedomostat PROMPT '��������� �������� �� ���� �� ��������� ������' ;
  MESSAGE '������������ ����������� ��������� �������� �� ��������� ���� �� ��������� ������'
DEFINE BAR 2 OF vedomostat PROMPT '\-'
DEFINE BAR 3 OF vedomostat PROMPT '��������� ��������� �� ���� �� ��������� ������' ;
  MESSAGE '������������ ��������� ��������� �� ��������� ������ �� ��������� ����'
DEFINE BAR 4 OF vedomostat PROMPT '\-'
DEFINE BAR 5 OF vedomostat PROMPT '��������� ��������� �� ������ �� ������ �� �� "Transmaster"' ;
  MESSAGE '������������ ��������� ��������� �� ��������� ������ �� �������� ��������� ��� �� �� "Transmaster"'
DEFINE BAR 6 OF vedomostat PROMPT '\-'
DEFINE BAR 7 OF vedomostat PROMPT '��������� �������� �� ����� ���� �� �� "Transmaster"' ;
  MESSAGE '������ ����������� ��������� �������� �� ���������� ���� �� �� "Transmaster"'
DEFINE BAR 8 OF vedomostat PROMPT '\-'
DEFINE BAR 9 OF vedomostat PROMPT '��������� �������� �� ����� ���� �� �� "CARD-Visa"' ;
  MESSAGE '������ ����������� ��������� �������� �� ���������� ���� �� �� "CARD-Visa"'
DEFINE BAR 10 OF vedomostat PROMPT '\-'
DEFINE BAR 11 OF vedomostat PROMPT '��������� �������� �� ����� ���� �� ��������� �������� ���' ;
  MESSAGE '������ ����������� ��������� �������� �� ���������� ���� �� ��������� ��������, ���������� �� ���'
DEFINE BAR 12 OF vedomostat PROMPT '\-'
DEFINE BAR 13 OF vedomostat PROMPT '������� �������� �� ����� ������� �� �� "Transmaster"' ;
  MESSAGE '������������ ������� �������� �� ���������� ����� ������� �� �� "Transmaster"'
DEFINE BAR 14 OF vedomostat PROMPT '\-'
DEFINE BAR 15 OF vedomostat PROMPT '������� �������� �� ����� ������� �� �� "CARD-Visa"' ;
  MESSAGE '������������ ������� �������� �� ���������� ����� ������� �� �� "CARD-Visa"'
DEFINE BAR 16 OF vedomostat PROMPT '\-'
DEFINE BAR 17 OF vedomostat PROMPT '��������� �������� �� ����� ������� �� �� "CARD-Visa"' ;
  MESSAGE '������������ ��������� �������� �� ���������� ����� ������� �� �� "CARD-Visa"'
DEFINE BAR 18 OF vedomostat PROMPT '\-'
DEFINE BAR 19 OF vedomostat PROMPT '��������� �������� �� 10 ������ � ���������� ��������' ;
  MESSAGE '������������ ��������� �������� �� 10 ������ � ���������� ��������'
DEFINE BAR 20 OF vedomostat PROMPT '\-'
DEFINE BAR 21 OF vedomostat PROMPT '������� �������� � �������� �� ������ �� �������� ����' ;
  MESSAGE '��������� ������� �������� � �������� �� ������ ������� �������� ����������� ���������� �� �������� ����'
DEFINE BAR 22 OF vedomostat PROMPT '\-'
DEFINE BAR 23 OF vedomostat PROMPT '������� �������� � �������� �� ������ �� �������� ����' ;
  MESSAGE '��������� ������� �������� � �������� �� ������ ������� �������� ����������� ���������� �� �������� ����'
DEFINE BAR 24 OF vedomostat PROMPT '\-'
DEFINE BAR 25 OF vedomostat PROMPT '������ ��������� �������� �� ������� ��������' ;
  MESSAGE '������ ��������� �������� �� ������� �������� �������������� �����'
DEFINE BAR 26 OF vedomostat PROMPT '\-'
DEFINE BAR 27 OF vedomostat PROMPT '������ ������� �������� �� "CARD-����" �� ����������� ����������' ;
  MESSAGE '������ ������� �������� �� "CARD-����" �� ����������� ����������'
DEFINE BAR 28 OF vedomostat PROMPT '\-'
DEFINE BAR 29 OF vedomostat PROMPT '����� � ���������� ��������� � ��������� �� ��� ��� ���� ��������' ;
  MESSAGE '������������ ������ �� ������� �������� �� ������ � ���������� ��������� � ��������� �� ��� �� �������� ��� ��� ���� ��������'
DEFINE BAR 30 OF vedomostat PROMPT '\-'
DEFINE BAR 31 OF vedomostat PROMPT '����� � ���������� ��������� (�����������, ������������) � ��������� �� ���' ;
  MESSAGE '������������ ������ �� ������� �������� �� ������ � ���������� ��������� (�����������, ������������) � ��������� �� ��� �� �������� ���'
DEFINE BAR 32 OF vedomostat PROMPT '\-'
DEFINE BAR 33 OF vedomostat PROMPT '����� � ���������� ��������� ��� ����� ����� � ��������� �� ���' ;
  MESSAGE '������������ ������ �� ������� �������� �� ������ � ���������� ��������� ��� ����� ����� � ��������� �� ��� �� �������� ���'

ON SELECTION BAR 1 OF vedomostat DO ostat_trans IN Docum_ostat
ON SELECTION BAR 3 OF vedomostat DO ostat_trans IN Docum_ostat
ON SELECTION BAR 5 OF vedomostat DO oborot_trans IN Oborot_period
ON SELECTION BAR 7 OF vedomostat DO cardtip_ostat IN Oborot_trans
ON SELECTION BAR 9 OF vedomostat DO cardtip_ostat IN Oborot_makvisa
ON SELECTION BAR 11 OF vedomostat DO cardtip_ostat IN Oborot_upk

DO CASE
  CASE pr_rabota_sql = .T.  && ������ ���������� � ��������� SQL Server

    ON SELECTION BAR 13 OF vedomostat DO start IN Istor_ostat_sql
    ON SELECTION BAR 15 OF vedomostat DO start IN Istor_mak_sql
    ON SELECTION BAR 17 OF vedomostat DO vedom_ost IN Vedom_ostat_sql

  CASE pr_rabota_sql = .F.  && ������ ���������� � ���������� ���������

    ON SELECTION BAR 13 OF vedomostat DO start IN Istor_ostat
    ON SELECTION BAR 15 OF vedomostat DO start IN Istor_mak
    ON SELECTION BAR 17 OF vedomostat DO vedom_ost IN Vedom_ostat

ENDCASE

ON SELECTION BAR 19 OF vedomostat DO ostat_10_max IN Vedom_ostat_max
ON SELECTION BAR 21 OF vedomostat DO start_ostat_den IN Prav_vedom_ost
ON SELECTION BAR 23 OF vedomostat DO start_ostat_arx IN Prav_vedom_ost
ON SELECTION BAR 25 OF vedomostat DO start IN Prav_vedom_ost
ON SELECTION BAR 27 OF vedomostat DO prav_istor_mak IN Prav_vedom_ost
ON SELECTION BAR 29 OF vedomostat DO start IN Oborot_debet_period
ON SELECTION BAR 31 OF vedomostat DO start IN Oborot_debet_monitor
ON SELECTION BAR 33 OF vedomostat DO start IN Oborot_debet_no_kassa


*************************************************************** DEFINE POPUP vipiska ********************************************************************

DEFINE POPUP vipiska ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1 OF vipiska PROMPT '������� �� ������ ���������� ���������� �����' ;
  MESSAGE '��������� ������������ ������� �� ������ ���������� ����� ������� � �������� �� ���'
DEFINE BAR 2 OF vipiska PROMPT '\-'
DEFINE BAR 3 OF vipiska PROMPT '������� �� ��������� ������ � ��������� �� ������� ����' ;
  MESSAGE '��������� ������������ ������� �� ���� ��������� ������ ��������, � ������� ���� �������� �� ����� �� ������� ���� ������� � ��������� �������'
DEFINE BAR 4 OF vipiska PROMPT '\-'
DEFINE BAR 5 OF vipiska PROMPT '������� �� ��������� ������ � ��������� �� �������� ����' ;
  MESSAGE '��������� ������������ ������� �� ���� ��������� ������ ��������, � ������� ���� �������� �� ����� �� �������� ���� ������� � ��������� �������'
DEFINE BAR 6 OF vipiska PROMPT '\-'
DEFINE BAR 7 OF vipiska PROMPT '����� �� ����� ������������ ��� �������� ������ � �����������' ;
  MESSAGE '����� �� ����� ������������ ��� �������� �� ��������� ����� ������ � ����������� �� ������� ���� ���'
DEFINE BAR 8 OF vipiska PROMPT '\-'
DEFINE BAR 9 OF vipiska PROMPT '������� ������������� ������ �� ������������� ���������' ;
  MESSAGE '������ � �������� ������������� ������ �� ������������� ��������� �� ���� ������ �� ������'

DO CASE
  CASE pr_rabota_sql = .T.  && ������ ���������� � ��������� SQL Server

    ON SELECTION BAR 1 OF vipiska DO start_vipis_cln IN Vipiska_sql
    ON SELECTION BAR 3 OF vipiska DO start_vipis_all_den IN Vipiska_sql
    ON SELECTION BAR 5 OF vipiska DO start_vipis_all_arx IN Vipiska_sql
    ON SELECTION BAR 7 OF vipiska DO start IN Overdraft
    ON SELECTION BAR 9 OF vipiska DO start IN Istor_overdraft

  CASE pr_rabota_sql = .F.  && ������ ���������� � ���������� ���������

    ON SELECTION BAR 1 OF vipiska DO start_vipis_cln IN Vipiska
    ON SELECTION BAR 3 OF vipiska DO start_vipis_all_den IN Vipiska
    ON SELECTION BAR 5 OF vipiska DO start_vipis_all_arx IN Vipiska
    ON SELECTION BAR 7 OF vipiska DO start IN Overdraft
    ON SELECTION BAR 9 OF vipiska DO start IN Istor_overdraft

ENDCASE


************************************************************* DEFINE POPUP analit_info ********************************************************************

DEFINE POPUP analit_info ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1 OF analit_info PROMPT '��������� �� ���������� ������� (������)' ;
  MESSAGE '������ ������ ��������� �� ���������� �������, ������� �������� ���� �� ������ ������� � ����� �����������'
DEFINE BAR 2 OF analit_info PROMPT '\-'
DEFINE BAR 3 OF analit_info PROMPT '��������� �� ���������� ������� (�� ������)' ;
  MESSAGE '������ ������ ��������� �� ���������� �������, ����������� �������� ���� �� ������ ������� � ����� ����������� ��� ���'
DEFINE BAR 4 OF analit_info PROMPT '\-'
DEFINE BAR 5 OF analit_info PROMPT '��������� �� ������������� �������� (������)' ;
  MESSAGE '������ ��������� �� ������ �������� ��� ���'
DEFINE BAR 6 OF analit_info PROMPT '\-'
DEFINE BAR 7 OF analit_info PROMPT '����� �� ��������� ������ ���������� � ����������� "CARD-����"' ;
  MESSAGE '����� �� ����������� ��������� ������ ���������� � ���������� � ���������� ������������� �� "CARD-����", ��� �������� 207.'
DEFINE BAR 8 OF analit_info PROMPT '\-'
DEFINE BAR 9 OF analit_info PROMPT '����� �� ��������� ������ ���������� � ����������� ����� ������' ;
  MESSAGE '����� �� ����������� ��������� ������ ���������� � ���������� � ���������� ������������� ��������� ������, ��� �������� 207.'
DEFINE BAR 10 OF analit_info PROMPT '\-'
DEFINE BAR 11 OF analit_info PROMPT '����� �� ��������� ������ ����� � ������� � ����������� "CARD-����"' ;
  MESSAGE '����� �� ����������� ��������� ������ ����� � ������� � ���������� � ���������� ������������� �� "CARD-����", ��� �������� 205.'
DEFINE BAR 12 OF analit_info PROMPT '\-'
DEFINE BAR 13 OF analit_info PROMPT '����� �� ��������� ������ ����� � ������� � ����������� ����� ������' ;
  MESSAGE '����� �� ����������� ��������� ������ ����� � ������� � ���������� � ���������� ������������� ��������� ������, ��� �������� 205.'
DEFINE BAR 14 OF analit_info PROMPT '\-'
DEFINE BAR 15 OF analit_info PROMPT '��������� �������� ��������� ������ �� �� VISA' ;
  MESSAGE '��������� �������� ��������� ������ �� �� VISA'
DEFINE BAR 16 OF analit_info PROMPT '\-'
DEFINE BAR 17 OF analit_info PROMPT '��������� �������� ��������� ������ �� �� VISA' ;
  MESSAGE '��������� �������� ��������� ������ �� �� VISA'
DEFINE BAR 18 OF analit_info PROMPT '\-'
DEFINE BAR 19 OF analit_info PROMPT '���������� ������� ��������� ��������� �� ������ ������� � ���' ;
  MESSAGE '���������� ������� ��������� ��������� �� ������ ������� � ���'
DEFINE BAR 20 OF analit_info PROMPT '\-'
DEFINE BAR 21 OF analit_info PROMPT '����� �� �������� � �������� ������ �� ������' ;
  MESSAGE '����� �� �������� � �������� ������, ���� �������� ������� �������� � ��������� ������'
DEFINE BAR 22 OF analit_info PROMPT '\-'
DEFINE BAR 23 OF analit_info PROMPT '����� �� �������� � �������� ������ �� ����' ;
  MESSAGE '����� �� �������� � �������� ������, ���� �������� ������� ����� ��������� ����'
DEFINE BAR 24 OF analit_info PROMPT '\-'
DEFINE BAR 25 OF analit_info PROMPT '����� �� �������� � �������� ������ �������� �� ������� ����' ;
  MESSAGE '����� �� �������� � �������� ������ �������� �� ������� ����'
DEFINE BAR 26 OF analit_info PROMPT '\-'
DEFINE BAR 27 OF analit_info PROMPT '����� �� �������� � �������� ������ �������� �� ����' ;
  MESSAGE '����� �� �������� � �������� ������ ��������, ���� �������� ������� ����� ��������� ����'
DEFINE BAR 28 OF analit_info PROMPT '\-'
DEFINE BAR 29 OF analit_info PROMPT '����� �� �������� � �������� ������ �������� �� ����' ;
  MESSAGE '����� �� �������� � �������� ������ ��������, ���� �������� ������� ����� ��������� ����'
DEFINE BAR 30 OF analit_info PROMPT '\-'
DEFINE BAR 31 OF analit_info PROMPT '����� �� �������� � �������� ������ �������� �� ������' ;
  MESSAGE '����� �� �������� � �������� ������ ��������, ���� �������� ������� �������� � ��������� ������'
DEFINE BAR 32 OF analit_info PROMPT '\-'
DEFINE BAR 33 OF analit_info PROMPT '����� �� ���������� ��������, ������, ���� �� ����' ;
  MESSAGE '����� �� ���������� ��������, ������, ����, ���� �������� ������� ����� ��������� ����'
DEFINE BAR 34 OF analit_info PROMPT '\-'
DEFINE BAR 35 OF analit_info PROMPT '����� �� ���������� ��������, ������, ���� �� ������' ;
  MESSAGE '����� �� ���������� ��������, ������, ���� �� ������, ���� �������� ������� �������� � ��������� ������'
DEFINE BAR 36 OF analit_info PROMPT '\-'
DEFINE BAR 37 OF analit_info PROMPT '����� � �������� ����� ��������� ����� � ������� ����� ��������' ;
  MESSAGE '����� �� ������ � �������� ����� ��������� ����� � ����� �������� ������� ���������'
DEFINE BAR 38 OF analit_info PROMPT '\-'
DEFINE BAR 39 OF analit_info PROMPT '����� � �������� ����� ��������� ����� � ������� ����� ��������' ;
  MESSAGE '����� �� ������ � �������� ����� ��������� ����� � ����� �������� ������� ���������'
DEFINE BAR 40 OF analit_info PROMPT '\-'
DEFINE BAR 41 OF analit_info PROMPT '����� ����������� ������� ������, � ����� �������� ������� ��������� ����' ;
  MESSAGE '����� �� �������� � �������� ������, � ����� �������� ������� ��������� ����'
DEFINE BAR 42 OF analit_info PROMPT '\-'
DEFINE BAR 43 OF analit_info PROMPT '����� �� ��������� ������ �� ��������� ������� �������� �� ����' ;
  MESSAGE '����� �� ��������� ������ �� ��������� ������� �������� �� ����'
DEFINE BAR 44 OF analit_info PROMPT '\-'
DEFINE BAR 45 OF analit_info PROMPT '����� �� ��������� ������ �� ������� �������� ����� 345 �� ����' ;
  MESSAGE '����� �� ��������� ������ �� ������� �������� ����� 345 �� ����'

ON SELECTION BAR 1 OF analit_info DO start_all IN Docum_doxod
ON SELECTION BAR 3 OF analit_info DO start_52 IN Docum_doxod
ON SELECTION BAR 5 OF analit_info DO start IN Docum_rasxod
ON SELECTION BAR 7 OF analit_info DO start_yes IN Mesto_tran
ON SELECTION BAR 9 OF analit_info DO start_no IN Mesto_tran
ON SELECTION BAR 11 OF analit_info DO start_yes IN Oplata_tovar
ON SELECTION BAR 13 OF analit_info DO start_no IN Oplata_tovar
ON SELECTION BAR 15 OF analit_info DO start_otkr_acc IN Reestr_account
ON SELECTION BAR 17 OF analit_info DO start_zakr_acc IN Reestr_account
ON SELECTION BAR 19 OF analit_info DO sverka_trans_visa IN Sverka_ostat
ON SELECTION BAR 21 OF analit_info DO sum_card_period IN Analitik_account
ON SELECTION BAR 23 OF analit_info DO sum_card_na_data IN Analitik_account
ON SELECTION BAR 25 OF analit_info DO type_card_nls IN Analitik_account
ON SELECTION BAR 27 OF analit_info DO type_card_nls_na_data IN Analitik_account
ON SELECTION BAR 29 OF analit_info DO type_card_nls_za_data IN Analitik_account
ON SELECTION BAR 31 OF analit_info DO type_card_nls_period IN Analitik_account
ON SELECTION BAR 33 OF analit_info DO cln_nls_card_na_data IN Analitik_account
ON SELECTION BAR 35 OF analit_info DO cln_nls_card_period IN Analitik_account
ON SELECTION BAR 37 OF analit_info DO card_nls_max_100000 IN Analitik_account
ON SELECTION BAR 39 OF analit_info DO card_nls_min_100000 IN Analitik_account
ON SELECTION BAR 41 OF analit_info DO start IN Registrat_card_data
ON SELECTION BAR 43 OF analit_info DO card_acct_filtr_summa IN Analitik_account
ON SELECTION BAR 45 OF analit_info DO card_acct_filtr_summa_345 IN Analitik_account


************************************************************* DEFINE POPUP otchet_zb *********************************************************************

DEFINE POPUP otchet_zb ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1 OF otchet_zb PROMPT '������ �� �������� �� ������ - ����� 1416 �' ;
  MESSAGE '������ � ������������ ������ � ���������� �������� �� ��������� ������ � ���������� ����������� - ����� 1416 �'
DEFINE BAR 2 OF otchet_zb PROMPT '\-'
DEFINE BAR 3 OF otchet_zb PROMPT '������ ������������ ����� - ����� 1417 �' ;
  MESSAGE '������ � ������������ ������� ������������ ����� ����� ����������� - ����� 1417 �'
DEFINE BAR 4 OF otchet_zb PROMPT '\-'
DEFINE BAR 5 OF otchet_zb PROMPT '������ ����������� ����� - ����� 1991 �' ;
  MESSAGE '������������ ���������������� ��������� �������� ����� - ����� 1991 �'

ON SELECTION BAR 1 OF otchet_zb DO start IN Otchet_1416_U
ON SELECTION BAR 3 OF otchet_zb DO start IN Otchet_1417_U
ON SELECTION BAR 5 OF otchet_zb DO start IN Otchet_1991_U


**************************************************************** DEFINE POPUP servis_doc *****************************************************************

DEFINE POPUP servis_doc  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF servis_doc PROMPT '���������� �������� ������ �������' ;
  MESSAGE '�������� � �������������� ����������� �������� ������ �������'
DEFINE BAR 2  OF servis_doc PROMPT '\-'
DEFINE BAR 3  OF servis_doc PROMPT '���������� ��������� ������ �������' ;
  MESSAGE '�������� � �������������� ��������� ������ �������'
DEFINE BAR 4  OF servis_doc PROMPT '\-'
DEFINE BAR 5  OF servis_doc PROMPT '���������� �� �������� ������' ;
  MESSAGE '�������� ����������� �� �������� ������'
DEFINE BAR 6  OF servis_doc PROMPT '\-'
DEFINE BAR 7  OF servis_doc PROMPT '���������� ����� ��������' ;
  MESSAGE '�������� ����������� ����� ��������'
DEFINE BAR 8  OF servis_doc PROMPT '\-'
DEFINE BAR 9  OF servis_doc PROMPT '���������� ����������' ;
  MESSAGE '�������� � �������������� ����������� ����������'
DEFINE BAR 10 OF servis_doc PROMPT '\-'
DEFINE BAR 11 OF servis_doc PROMPT '���������� ������ ����� �� ����' ;
  MESSAGE '�������� � �������������� ����������� ������ ����� �� ��������� ����'
DEFINE BAR 12 OF servis_doc PROMPT '\-'
DEFINE BAR 13 OF servis_doc PROMPT '���������� ���������� ������' ;
  MESSAGE '�������� � �������������� ����������� ���������� ������'
DEFINE BAR 14 OF servis_doc PROMPT '\-'
DEFINE BAR 15 OF servis_doc PROMPT '���������� �������� ������' ;
  MESSAGE '�������� � �������������� ����������� �������� ������ �� ������ � ������������ ���� Visa'
DEFINE BAR 16 OF servis_doc PROMPT '\-'
DEFINE BAR 17 OF servis_doc PROMPT '���������� ������ ����������� ���� �������' ;

DO CASE
  CASE pr_rabota_sql = .T.  && ������ ���������� � ��������� SQL Server

    ON SELECTION BAR 1   OF servis_doc DO spr_client IN Servis_sql
    ON SELECTION BAR 3   OF servis_doc DO read_sks IN Servis_sql
    ON SELECTION BAR 5   OF servis_doc DO spr_karta IN Servis_sql
    ON SELECTION BAR 7   OF servis_doc DO spr_kod_oper  IN Servis_sql
    ON SELECTION BAR 9   OF servis_doc DO spravliz  IN Servis_sql
    ON SELECTION BAR 11 OF servis_doc DO red_spr_val IN Servis_sql
    ON SELECTION BAR 13 OF servis_doc DO read_tran IN Servis_sql
    ON SELECTION BAR 15 OF servis_doc DO read_tarif IN Servis_sql
    ON SELECTION BAR 17 OF servis_doc DO read_pass_dover IN Servis

  CASE pr_rabota_sql = .F.  && ������ ���������� � ���������� ���������

    ON SELECTION BAR 1   OF servis_doc DO spr_client IN Servis
    ON SELECTION BAR 3   OF servis_doc DO read_sks IN Servis
    ON SELECTION BAR 5   OF servis_doc DO spr_karta IN Servis
    ON SELECTION BAR 7   OF servis_doc DO spr_kod_oper  IN Servis
    ON SELECTION BAR 9   OF servis_doc DO spravliz  IN Servis
    ON SELECTION BAR 11 OF servis_doc DO red_spr_val IN Servis
    ON SELECTION BAR 13 OF servis_doc DO read_tran IN Servis
    ON SELECTION BAR 15 OF servis_doc DO read_tarif IN Servis
    ON SELECTION BAR 17 OF servis_doc DO read_pass_dover IN Servis

ENDCASE


* ================== ������������ ��������������� ���� ��� ������ �������� ���� "������ ���� � ���������� ��������� ������" ====================== *


*********************************************************** DEFINE MENU client **************************************************************************

DEFINE MENU client COLOR SCHEME 2 ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B'

DEFINE PAD cardclient OF client PROMPT '���������' COLOR SCHEME 2
DEFINE PAD popolclient OF client PROMPT '����������' COLOR SCHEME 2
DEFINE PAD spisclient OF client PROMPT '���������' COLOR SCHEME 2
DEFINE PAD popol_ckc OF client PROMPT '����������' COLOR SCHEME 2
DEFINE PAD zakrclient OF client PROMPT '��������' COLOR SCHEME 2
DEFINE PAD analitika OF client PROMPT '���������' COLOR SCHEME 2
DEFINE PAD servis_cln OF client PROMPT '�����������' COLOR SCHEME 2
DEFINE PAD quit OF client PROMPT '�����' COLOR SCHEME 2 ;
  MESSAGE '����� � ������� ����'

ON PAD cardclient OF client ACTIVATE POPUP cardclient
ON PAD popolclient OF client ACTIVATE POPUP popolclient
ON PAD spisclient OF client ACTIVATE POPUP spisclient
ON PAD popol_ckc OF client ACTIVATE POPUP popol_ckc
ON PAD zakrclient OF client ACTIVATE POPUP zakrclient
ON PAD analitika OF client ACTIVATE POPUP analitika
ON PAD servis_cln OF client ACTIVATE POPUP servis_cln
ON SELECTION PAD quit OF client DEACTIVATE MENU client


***************************************************************** DEFINE POPUP cardclient *****************************************************************

DEFINE POPUP cardclient  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF cardclient PROMPT '�������� ����� - ���������� ������ �� ������' ;
  MESSAGE '������ ����� ���� ��������� ������ ������ �� ������� ��� ���������� ��������� �� ������ �������� ��'
DEFINE BAR 2  OF cardclient PROMPT '\-'
DEFINE BAR 3  OF cardclient PROMPT '�������� ����� - ������������ ������ ��� �����������' ;
  MESSAGE '������ ����� ���� ��������� ����������� ������ �� ����� �������� ��� �������� � �������������� ����� �����'
DEFINE BAR 4  OF cardclient PROMPT '\-'
DEFINE BAR 5  OF cardclient PROMPT '�������� ����� - ������� ������ �� ������ ����' ;
  MESSAGE '������ ����� ���� ��������� ����������� ������� ������ �� ������ ���� �� ����� �������� ��� �������� � �������������� ����� �����'
DEFINE BAR 6  OF cardclient PROMPT '\-'
DEFINE BAR 7  OF cardclient PROMPT '�������������� ����� - ���������� ������ �� ������' ;
  MESSAGE '������ ����� ���� ��������� ������ ������ �� ������� ��� ���������� ��������� �� ������ �������������� ��'
DEFINE BAR 8  OF cardclient PROMPT '\-'
DEFINE BAR 9  OF cardclient PROMPT '�������������� ����� - ������������ ������ ��� �����������' ;
  MESSAGE '������ ����� ���� ��������� ����������� ������ �� ����� �������� ��� �������� � �������������� ����� �����'
DEFINE BAR 10  OF cardclient PROMPT '\-'
DEFINE BAR 11  OF cardclient PROMPT '�������������� ����� - ������� ������ �� ������ ����' ;
  MESSAGE '������ ����� ���� ��������� ����������� ������� ������ �� ������ ���� �� ����� �������� ��� �������� � �������������� ����� �����'
DEFINE BAR 12  OF cardclient PROMPT '\-'
DEFINE BAR 13  OF cardclient PROMPT '������������ ���������� ������ ��� ����� ��������� ��������' ;
  MESSAGE '������ ����� ���� ��������� ��������� ���������� ������ �� ����� ��������� �������� � �������������� ����� �����'
DEFINE BAR 14  OF cardclient PROMPT '\-'
DEFINE BAR 15  OF cardclient PROMPT '������ ������� ������ ��� �������� ���� �������' ;
  MESSAGE '������ ����� ���� ��������� ������������� ������� �������������� ������ ��� �������� ���� �������'
DEFINE BAR 16  OF cardclient PROMPT '\-'
DEFINE BAR 17  OF cardclient PROMPT '�������� ��������� ����� ��� �������� � ���' ;
  MESSAGE '������ ����� ���� ��������� ������� � ��������� �������� ���� � ���'

ON SELECTION BAR 1    OF cardclient DO start IN Zajava_blank_osn
ON SELECTION BAR 3    OF cardclient DO start IN Zajava_export_osn
ON SELECTION BAR 5    OF cardclient DO zajava_svod IN Zajava_blank_osn
ON SELECTION BAR 7    OF cardclient DO start IN Zajava_blank_dop
ON SELECTION BAR 9    OF cardclient DO start IN Zajava_export_dop
ON SELECTION BAR 11  OF cardclient DO zajava_svod IN Zajava_blank_dop
ON SELECTION BAR 13  OF cardclient DO start IN Client_export_red
ON SELECTION BAR 15  OF cardclient DO import IN Zajava_blank_osn
ON SELECTION BAR 17  OF cardclient DO create_zip IN Zajava_export_osn


************************************************************** DEFINE POPUP popolclient *******************************************************************

DEFINE POPUP popolclient  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF popolclient PROMPT '���������� ����� ��������� - �������� �����' ;
  MESSAGE '������������ ������������� ���������� �� ���������� ����������� ����� ��������� - ����������� ���������� ��������� ������'
DEFINE BAR 2  OF popolclient PROMPT '\-'
DEFINE BAR 3  OF popolclient PROMPT '���������� ����� ������������ - ���������' ;
  MESSAGE '������������ ������������� ���������� �� ���������� ����������� ����� ������������ - ����������� ������������� ������'
DEFINE BAR 4  OF popolclient PROMPT '\-'
DEFINE BAR 5  OF popolclient PROMPT '���������� ����� ������������ - ��������� ���������' ;
  MESSAGE '������������ ������������� ���������� �� ���������� ����������� ����� ������������ - ����������� ���������� ���������'
DEFINE BAR 6 OF popolclient PROMPT '\-'
DEFINE BAR 7 OF popolclient PROMPT '���������� ������� �� ����������� ������� - ���������' ;
  MESSAGE '������������ ������������� ���������� ��� ���������� ������� �� ����������� ������� ������������ - ����������� ������������� ������'
DEFINE BAR 8 OF popolclient PROMPT '\-'
DEFINE BAR 9 OF popolclient PROMPT '������ ������ �� ����������� ������� - ��������� - ������ DBF' ;
  MESSAGE '������������ ������������� ���������� ��� ���������� ������� �� ����������� ������� ����� ������� ������ � ����������� ����'
DEFINE BAR 10 OF popolclient PROMPT '\-'
DEFINE BAR 11 OF popolclient PROMPT '������ ������ �� ����������� ������� - ��������� - ������ DCL' ;
  MESSAGE '������������ ������������� ���������� ��� ���������� ������� �� ����������� ������� ����� ������� ������ � ����������� ����'

ON SELECTION BAR 1   OF popolclient DO start  IN Order_kassa
ON SELECTION BAR 3   OF popolclient DO start  IN Order_memo
ON SELECTION BAR 5   OF popolclient DO start IN Plateg
ON SELECTION BAR 7   OF popolclient DO start  IN Order_memo
ON SELECTION BAR 9   OF popolclient DO import_cred_dbf  IN Dokum_memo
ON SELECTION BAR 11 OF popolclient DO import_cred_dcl  IN Dokum_memo


*************************************************************** DEFINE POPUP spisclient *******************************************************************

DEFINE POPUP spisclient  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF spisclient PROMPT '�������� �� ����� ��������� ������������ ����' ;
  MESSAGE '������ ����� ���� ��������� ��������� ��������� �� �������� �� ����� ��������� ������������ ����'
DEFINE BAR 2  OF spisclient PROMPT '\-'
DEFINE BAR 3  OF spisclient PROMPT '�������� �� ����� ��������� ���� ���������� � ����' ;
  MESSAGE '������ ����� ���� ��������� ��������� ��������� �� �������� �� ����� ��������� ���� ���������� � ����'
DEFINE BAR 4  OF spisclient PROMPT '\-'
DEFINE BAR 5  OF spisclient PROMPT '�������� �� ����� ��������� �������� ����' ;
  MESSAGE '������ ����� ���� ��������� ��������� ��������� �� �������� �� ����� ��������� �������� ����'
DEFINE BAR 6  OF spisclient PROMPT '\-'
DEFINE BAR 7  OF spisclient PROMPT '���������� �� ���� ��������� �� ������������' ;
  MESSAGE '������ ����� ���� ��������� ��������� ��������� �� ���������� �� ���� ��������� �� ������������'
DEFINE BAR 8 OF spisclient PROMPT '\-'
DEFINE BAR 9 OF spisclient PROMPT '������ ������ �� �������� � ��������� ������ �������' ;
  MESSAGE '������������ ������������� ���������� �� �������� ������� � ��������� ������ ����� ������� ������ ������� � ����������� ����'

ON SELECTION BAR 1  OF spisclient DO start IN Order_memo
ON SELECTION BAR 3  OF spisclient DO start IN Order_memo
ON SELECTION BAR 5  OF spisclient DO start IN Order_memo
ON SELECTION BAR 7  OF spisclient DO start IN Order_memo
ON SELECTION BAR 9 OF spisclient DO import_deb  IN Dokum_memo


************************************************************* DEFINE POPUP popol_ckc   *******************************************************************

DEFINE POPUP popol_ckc  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF popol_ckc PROMPT '������������ ����� ���������� � �������� ��� ���' ;
  MESSAGE '������ ����� ���� ��������� ����������� ����������� ���� ���������� � �������� ��� ���������� ���������� ��� �������� ��� � ���'
DEFINE BAR 2  OF popol_ckc PROMPT '\-'
DEFINE BAR 3  OF popol_ckc PROMPT '�������� ��������� ����� ��� �������� � ���' ;
  MESSAGE '������ ����� ���� ��������� ��������� �������� ���� �� ������������� ������� ��� �������� ��� � ���'
DEFINE BAR 4 OF popol_ckc PROMPT '\-'
DEFINE BAR 5 OF popol_ckc PROMPT '������� ���� ���������� � �������� � ��������' ;
  MESSAGE '������ ����� ���� ��������� ��������� ���� � ������� ���������� � �������� ��� �������� � ������������ ���� �����'

ON SELECTION BAR 1  OF popol_ckc DO start IN Popol_ckc
ON SELECTION BAR 3  OF popol_ckc DO create_zip IN Popol_ckc
ON SELECTION BAR 5  OF popol_ckc DO export_odb IN Popol_doc


*************************************************************** DEFINE POPUP zakrclient *******************************************************************

DEFINE POPUP zakrclient  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF zakrclient PROMPT '���������� ����� ��������� - �������� �����' ;
  MESSAGE '������������ ������������� ���������� �� ���������� ����������� ����� ��������� - ����������� ���������� ��������� ������'
DEFINE BAR 2  OF zakrclient PROMPT '\-'
DEFINE BAR 3  OF zakrclient PROMPT '���������� ����� ������������ - ���������' ;
  MESSAGE '������������ ������������� ���������� �� ���������� ����������� ����� ������������ - ����������� ������������� ������'
DEFINE BAR 4  OF zakrclient PROMPT '\-'
DEFINE BAR 5  OF zakrclient PROMPT '���������� ����� ������������ - ��������� ���������' ;
  MESSAGE '������������ ������������� ���������� �� ���������� ����������� ����� ������������ - ����������� ���������� ���������'

ON SELECTION BAR 1 OF zakrclient DO start  IN Order_kassa
ON SELECTION BAR 3 OF zakrclient DO start  IN Order_memo
ON SELECTION BAR 5 OF zakrclient DO start IN Plateg


*************************************************************** DEFINE POPUP analitika *******************************************************************

DEFINE POPUP analitika  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF analitika PROMPT '����������� � ���������� ������� �� ������' ;
  MESSAGE '������ ����� ���� ��������� ����������� ������� ����������� � ���������� ������� �� ������ ��� �������� � ���'
DEFINE BAR 2  OF analitika PROMPT '\-'
DEFINE BAR 3  OF analitika PROMPT '����������� � �������� ������� �� ������' ;
  MESSAGE '������ ����� ���� ��������� ����������� ������� ����������� � �������� ������� �� ������ ��� �������� � ���'
DEFINE BAR 4 OF analitika PROMPT '\-'
DEFINE BAR 5 OF analitika PROMPT '������� �������� � ��������� ��� ��� ������-������' ;
  MESSAGE '������ ����� ���� ��������� ���������� ������� ����������� �������� � ��������� ��� ��� ������-������'
DEFINE BAR 6 OF analitika PROMPT '\-'
DEFINE BAR 7 OF analitika PROMPT '������ ����, ������� ��������������� � ������� ������' ;
  MESSAGE '������ �� �������, � ������� ���������� ������� ����� �� ������� ������������ �� ������� ����'
DEFINE BAR 8 OF analitika PROMPT '\-'
DEFINE BAR 9 OF analitika PROMPT '����� ������ �� ���������� ���� � ������� ������' ;
  MESSAGE '�����, � ������� ���������� ������� ����� �� ���������� ���� � ������� ������'

ON SELECTION BAR 1  OF analitika DO start_popol IN Popol_doc
ON SELECTION BAR 3  OF analitika DO start_spis IN Popol_doc
ON SELECTION BAR 5  OF analitika DO faktura IN Dokum_memo
ON SELECTION BAR 7  OF analitika DO start IN Obsluga_god
ON SELECTION BAR 9  OF analitika DO start IN Obsl_god_proekt


*************************************************************** DEFINE POPUP servis_cln *******************************************************************

DEFINE POPUP servis_cln  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF servis_cln PROMPT '���������� �������� ������ �������' ;
  MESSAGE '�������� ����������� �������� ������ �������'
DEFINE BAR 2  OF servis_cln PROMPT '\-'
DEFINE BAR 3  OF servis_cln PROMPT '���������� ��������� ������ �������' ;
  MESSAGE '�������������� ��������� ������ �������'
DEFINE BAR 4  OF servis_cln PROMPT '\-'
DEFINE BAR 5  OF servis_cln PROMPT '���������� �� �������� ������' ;
  MESSAGE '�������� ����������� �� ������'
DEFINE BAR 6  OF servis_cln PROMPT '\-'
DEFINE BAR 7  OF servis_cln PROMPT '���������� ����� ��������' ;
  MESSAGE '�������� ����������� ����� ��������'
DEFINE BAR 8  OF servis_cln PROMPT '\-'
DEFINE BAR 9  OF servis_cln PROMPT '���������� ����������' ;
  MESSAGE '�������� � �������������� ����������� ����������'
DEFINE BAR 10 OF servis_cln PROMPT '\-'
DEFINE BAR 11 OF servis_cln PROMPT '���������� ������ �����' ;
  MESSAGE '�������� � �������������� ����������� ������ ����� �� ��������� ����'
DEFINE BAR 12 OF servis_cln PROMPT '\-'
DEFINE BAR 13 OF servis_cln PROMPT '���������� ���������� ������' ;
  MESSAGE '�������������� ����������� ���������� ������'
DEFINE BAR 14 OF servis_cln PROMPT '\-'
DEFINE BAR 15 OF servis_cln PROMPT '������ ��������������� ���������� - �������� ����������' ;
  MESSAGE '������ ��������������� ���������� - �������� ����������'
DEFINE BAR 16 OF servis_cln PROMPT '\-'
DEFINE BAR 17 OF servis_cln PROMPT '������ ��������������� ���������� - ��������� �����' ;
  MESSAGE '������ ��������������� ���������� - ��������� �����'
DEFINE BAR 18 OF servis_cln PROMPT '\-'
DEFINE BAR 19 OF servis_cln PROMPT '������ ��������� �� ��������� �����' ;
  MESSAGE '������ ��������� �� ��������� �����'
DEFINE BAR 20 OF servis_cln PROMPT '\-'
DEFINE BAR 21 OF servis_cln PROMPT '���������� �������� ������' ;
  MESSAGE '�������������� ����������� �������� ������ �� ������ � ������������ ���� Visa'
DEFINE BAR 22 OF servis_cln PROMPT '\-'
DEFINE BAR 23 OF servis_cln PROMPT '���������� ������ ����������� ���� �������' ;

DO CASE
  CASE pr_rabota_sql = .T.  && ������ ���������� � ��������� SQL Server

    ON SELECTION BAR 1   OF servis_cln DO spr_client IN Servis_sql
    ON SELECTION BAR 3   OF servis_cln DO read_sks IN Servis_sql
    ON SELECTION BAR 5   OF servis_cln DO spr_karta IN Servis_sql
    ON SELECTION BAR 7   OF servis_cln DO spr_kod_oper  IN Servis_sql
    ON SELECTION BAR 9   OF servis_cln DO spravliz  IN Servis_sql
    ON SELECTION BAR 11 OF servis_cln DO red_spr_val IN Servis_sql
    ON SELECTION BAR 13 OF servis_cln DO read_tran IN Servis_sql
    ON SELECTION BAR 15 OF servis_cln DO start IN Prn_dop_sogl
    ON SELECTION BAR 17 OF servis_cln DO start IN Prn_new_card_nls
    ON SELECTION BAR 19 OF servis_cln DO start IN Prn_zajava
    ON SELECTION BAR 21 OF servis_cln DO read_tarif IN Servis_sql
    ON SELECTION BAR 23 OF servis_cln DO read_pass_dover IN Servis

  CASE pr_rabota_sql = .F.  && ������ ���������� � ���������� ���������

    DEFINE BAR 24 OF servis_cln PROMPT '\-'
    DEFINE BAR 25 OF servis_cln PROMPT '���������� ����� ���������� �������� � ������� ���������' ;
      MESSAGE '���������� ����� ���������� �������� � ������� ��������� �� ��������� ����������� �����'

    ON SELECTION BAR 1   OF servis_cln DO spr_client IN Servis
    ON SELECTION BAR 3   OF servis_cln DO read_sks IN Servis
    ON SELECTION BAR 5   OF servis_cln DO spr_karta IN Servis
    ON SELECTION BAR 7   OF servis_cln DO spr_kod_oper  IN Servis
    ON SELECTION BAR 9   OF servis_cln DO spravliz  IN Servis
    ON SELECTION BAR 11 OF servis_cln DO red_spr_val IN Servis
    ON SELECTION BAR 13 OF servis_cln DO read_tran IN Servis
    ON SELECTION BAR 15 OF servis_cln DO start IN Prn_dop_sogl
    ON SELECTION BAR 17 OF servis_cln DO start IN Prn_new_card_nls
    ON SELECTION BAR 19 OF servis_cln DO start IN Prn_zajava
    ON SELECTION BAR 21 OF servis_cln DO read_tarif IN Servis
    ON SELECTION BAR 23 OF servis_cln DO read_pass_dover IN Servis
    ON SELECTION BAR 25 OF servis_cln DO start IN Dobav_cln_zajava

ENDCASE


* ===================== ������������ ��������������� ���� ��� ������ �������� ���� "�������� � ������������ ������� � ���" ====================== *


************************************************************* DEFINE MENU otchet_upk ********************************************************************

DEFINE MENU otchet_upk COLOR SCHEME 2 ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B'

DEFINE PAD otchet_upk OF otchet_upk PROMPT '������ � ���' COLOR SCHEME 2
DEFINE PAD servis_upk OF otchet_upk PROMPT '�����������' COLOR SCHEME 2
DEFINE PAD quit OF otchet_upk PROMPT '�����' COLOR SCHEME 2 ;
  MESSAGE '����� � ������� ����'

ON PAD otchet_upk OF otchet_upk ACTIVATE POPUP otchet_upk
ON PAD servis_upk  OF otchet_upk ACTIVATE POPUP servis_upk
ON SELECTION PAD quit OF otchet_upk DEACTIVATE MENU otchet_upk


***************************************************************** DEFINE POPUP otchet_upk   **************************************************************

DEFINE POPUP otchet_upk  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF otchet_upk PROMPT '��������� � ����������� � ����� �����' ;
  MESSAGE '������ ����� ���� ��������� ������������ ��������� � ����������� � ����� �����'
DEFINE BAR 2  OF otchet_upk PROMPT '\-'
DEFINE BAR 3  OF otchet_upk PROMPT '��������� �� ������ ����������� �����' ;
  MESSAGE '������ ����� ���� ��������� ������������ ��������� �� ������ �����'
DEFINE BAR 4  OF otchet_upk PROMPT '\-'
DEFINE BAR 5  OF otchet_upk PROMPT '��������� �� ������  ����������� �����' ;
  MESSAGE '������ ����� ���� ��������� ������������ ��������� �� ������ �����'
DEFINE BAR 6  OF otchet_upk PROMPT '\-'
DEFINE BAR 7  OF otchet_upk PROMPT '�������� ������� � ��������� ����� ����' ;
  MESSAGE '������ ����� ���� ��������� ������������ �������� � ��������� �����'
DEFINE BAR 8  OF otchet_upk PROMPT '\-'
DEFINE BAR 9  OF otchet_upk PROMPT '����������� � ������ ���� �������' ;
  MESSAGE '������ ����� ���� ��������� ������������ ����������� � ������ ���� �������'
DEFINE BAR 10 OF otchet_upk PROMPT '\-'
DEFINE BAR 11 OF otchet_upk PROMPT '����������� � ����� ���� �������' ;
  MESSAGE '������ ����� ���� ��������� ������������ ����������� � ����� ���� �������'

ON SELECTION BAR 1   OF otchet_upk DO start IN Zajava_sdacha
ON SELECTION BAR 3   OF otchet_upk DO otkaz_card IN Otchet_upk
ON SELECTION BAR 5   OF otchet_upk DO utrata_card IN Otchet_upk
ON SELECTION BAR 7   OF otchet_upk DO poluch_card IN Otchet_upk
ON SELECTION BAR 9   OF otchet_upk DO vidacha_card IN Otchet_upk
ON SELECTION BAR 11 OF otchet_upk DO sdacha_card IN Otchet_upk


**************************************************************** DEFINE POPUP servis_upk ****************************************************************

DEFINE POPUP servis_upk  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF servis_upk PROMPT '���������� �������� ������ �������' ;
  MESSAGE '�������� ����������� �������� ������ �������'
DEFINE BAR 2  OF servis_upk PROMPT '\-'
DEFINE BAR 3  OF servis_upk PROMPT '���������� ��������� ������ �������' ;
  MESSAGE '�������������� ��������� ������ �������'
DEFINE BAR 4  OF servis_upk PROMPT '\-'
DEFINE BAR 5  OF servis_upk PROMPT '���������� �� �������� ������' ;
  MESSAGE '�������� ����������� �� ������'
DEFINE BAR 6  OF servis_upk PROMPT '\-'
DEFINE BAR 7  OF servis_upk PROMPT '���������� ����� ��������' ;
  MESSAGE '�������� ����������� ����� ��������'
DEFINE BAR 8  OF servis_upk PROMPT '\-'
DEFINE BAR 9  OF servis_upk PROMPT '���������� ����������' ;
  MESSAGE '�������� � �������������� ����������� ����������'
DEFINE BAR 10 OF servis_upk PROMPT '\-'
DEFINE BAR 11 OF servis_upk PROMPT '���������� ������ �����' ;
  MESSAGE '�������� � �������������� ����������� ������ ����� �� ��������� ����'
DEFINE BAR 12 OF servis_upk PROMPT '\-'
DEFINE BAR 13 OF servis_upk PROMPT '���������� ���������� ������' ;
  MESSAGE '�������������� ����������� ���������� ������'
DEFINE BAR 14 OF servis_upk PROMPT '\-'
DEFINE BAR 15 OF servis_upk PROMPT '���������� �������� ������' ;
  MESSAGE '�������������� ����������� �������� ������ �� ������ � ������������ ���� Visa'

DO CASE
  CASE pr_rabota_sql = .T.  && ������ ���������� � ��������� SQL Server

    ON SELECTION BAR 1   OF servis_upk DO spr_client IN Servis_sql
    ON SELECTION BAR 3   OF servis_upk DO read_sks IN Servis_sql
    ON SELECTION BAR 5   OF servis_upk DO spr_karta IN Servis_sql
    ON SELECTION BAR 7   OF servis_upk DO spr_kod_oper  IN Servis_sql
    ON SELECTION BAR 9   OF servis_upk DO spravliz  IN Servis_sql
    ON SELECTION BAR 11 OF servis_upk DO red_spr_val IN Servis_sql
    ON SELECTION BAR 13 OF servis_upk DO read_tran IN Servis_sql
    ON SELECTION BAR 15 OF servis_upk DO read_tarif IN Servis_sql

  CASE pr_rabota_sql = .F.  && ������ ���������� � ���������� ���������

    ON SELECTION BAR 1   OF servis_upk DO spr_client IN Servis
    ON SELECTION BAR 3   OF servis_upk DO read_sks IN Servis
    ON SELECTION BAR 5   OF servis_upk DO spr_karta IN Servis
    ON SELECTION BAR 7   OF servis_upk DO spr_kod_oper  IN Servis
    ON SELECTION BAR 9   OF servis_upk DO spravliz  IN Servis
    ON SELECTION BAR 11 OF servis_upk DO red_spr_val IN Servis
    ON SELECTION BAR 13 OF servis_upk DO read_tran IN Servis
    ON SELECTION BAR 15 OF servis_upk DO read_tarif IN Servis

ENDCASE


* ===================== ������������ ��������������� ���� ��� ������ �������� ���� "���������� ��������� �� ���������� �����" ====================== *


************************************************************* DEFINE MENU proz_ckc *********************************************************************

DEFINE MENU proz_ckc COLOR SCHEME 2 ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B'

DEFINE PAD raschet_proz_42301 OF proz_ckc PROMPT '������ �� ��������� ������' COLOR SCHEME 2
DEFINE PAD raschet_proz_42308 OF proz_ckc PROMPT '������ �� ��������� ���������' COLOR SCHEME 2
DEFINE PAD blank_proz OF proz_ckc PROMPT '������������' COLOR SCHEME 2
DEFINE PAD servis_proz OF proz_ckc PROMPT '�����������' COLOR SCHEME 2
DEFINE PAD quit OF proz_ckc PROMPT '�����' COLOR SCHEME 2 ;
  MESSAGE '����� � ������� ����'

ON PAD raschet_proz_42301  OF proz_ckc ACTIVATE POPUP raschet_proz_42301
ON PAD raschet_proz_42308  OF proz_ckc ACTIVATE POPUP raschet_proz_42308
ON PAD blank_proz    OF proz_ckc ACTIVATE POPUP blank_proz
ON PAD servis_proz  OF proz_ckc ACTIVATE POPUP servis_proz
ON SELECTION PAD quit OF proz_ckc DEACTIVATE MENU proz_ckc


********************************************************* DEFINE POPUP raschet_proz_42301 ****************************************************************

DEFINE POPUP raschet_proz_42301  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF raschet_proz_42301 PROMPT '��������� ���������� ���������' ;
  MESSAGE '���� � �������������� ������� ������ � ���������� ������, ����������� ��� ������� ���������'
DEFINE BAR 2  OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 3  OF raschet_proz_42301 PROMPT '������ ��������� ������ �������' ;
  MESSAGE '���� ��� ������ � ��������� �������, ���� ����� ������ ��� ������� � ������ ���������'
DEFINE BAR 4  OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 5  OF raschet_proz_42301 PROMPT '������ ��������� ����������' ;
  MESSAGE '���� ��� ������ � ��������� �������, ���� ����� ������, ����� ��������� ����� � ������ ���������'
DEFINE BAR 6  OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 7  OF raschet_proz_42301 PROMPT '�������������� ������ �������' ;
  MESSAGE '����� ��� ������ � ��������� ������� ������� �� ������ � �������������� ������������ ������'
DEFINE BAR 8  OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 9  OF raschet_proz_42301 PROMPT '�������� ������ �� ������������ ���������' ;
  MESSAGE '�������� ������ �� ����� ������������ ��������� ��� �� ���������� �������, ��� � ���� ������ �� ��������� �������� ���'
DEFINE BAR 10 OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 11 OF raschet_proz_42301 PROMPT '�������� ������������ ���������' ;
  MESSAGE '����� ��� ������ � ��������� ������� ������� �� ������ � �������� ������������ ���������'
DEFINE BAR 12  OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 13  OF raschet_proz_42301 PROMPT '������ ������������ ���������' ;
  MESSAGE '����� ��� ������ � ��������� ������� ������� �� ������ � �������� ������������ ���������'
DEFINE BAR 14  OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 15  OF raschet_proz_42301 PROMPT '������������ ����� Pb �� ���������� ���������' ;
  MESSAGE '������������ ����� Pb �� ���������� ��������� � ����� 516'
DEFINE BAR 16 OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 17 OF raschet_proz_42301 PROMPT '�������� ��������� ����� ��� �������� � ���' ;
  MESSAGE '������ ����� ���� ��������� ��������� �������� ���� �� ������������� ������� ��� �������� � ���'
DEFINE BAR 18 OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 19 OF raschet_proz_42301 PROMPT '������� ����������� ��������� � ������������ ���� �����' ;
  MESSAGE '������ ����� ���� ��������� ��������� ���� � ������������ ���������� ��� ������� � ������������ ���� �����'

ON SELECTION BAR 1  OF raschet_proz_42301 DO  sel_fio IN Raschet_proz_42301
ON SELECTION BAR 3  OF raschet_proz_42301 DO  sum_proz IN Raschet_proz_42301
ON SELECTION BAR 5  OF raschet_proz_42301 DO  vibor_proz IN Raschet_proz_42301
ON SELECTION BAR 7  OF raschet_proz_42301 DO  read_sum IN Raschet_proz_42301
ON SELECTION BAR 9  OF raschet_proz_42301 DO  del_data IN Raschet_proz_42301
ON SELECTION BAR 11 OF raschet_proz_42301 DO  sel_data IN Raschet_proz_42301
ON SELECTION BAR 13 OF raschet_proz_42301 DO sel_data IN Raschet_proz_42301
ON SELECTION BAR 15 OF raschet_proz_42301 DO formirov_pb IN Raschet_proz_42301
ON SELECTION BAR 17 OF raschet_proz_42301 DO create_zip IN Popol_ckc
ON SELECTION BAR 19 OF raschet_proz_42301 DO export_odb IN Raschet_proz_42301


************************************************************* DEFINE POPUP raschet_proz_42308 ************************************************************

DEFINE POPUP raschet_proz_42308  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF raschet_proz_42308 PROMPT '��������� ������� ������' ;
  MESSAGE '���� � �������������� ������� ������, ����������� ��� ������� ���������'
DEFINE BAR 2  OF raschet_proz_42308 PROMPT '\-'
DEFINE BAR 3  OF raschet_proz_42308 PROMPT '������ ��������� ������ �������' ;
  MESSAGE '���� ��� ������ � ��������� �������, ���� ����� ������ ��� ������� � ������ ���������'
DEFINE BAR 4  OF raschet_proz_42308 PROMPT '\-'
DEFINE BAR 5  OF raschet_proz_42308 PROMPT '������ ��������� ����������' ;
  MESSAGE '���� ��� ������ � ��������� �������, ���� ����� ������, ����� ��������� ����� � ������ ���������'
DEFINE BAR 6  OF raschet_proz_42308 PROMPT '\-'
DEFINE BAR 7  OF raschet_proz_42308 PROMPT '�������������� ������ �������' ;
  MESSAGE '����� ��� ������ � ��������� ������� ������� �� ������ � �������������� ������������ ������'
DEFINE BAR 8  OF raschet_proz_42308 PROMPT '\-'
DEFINE BAR 9  OF raschet_proz_42308 PROMPT '�������� ������ �� ������������ ���������' ;
  MESSAGE '�������� ������ �� ����� ������������ ��������� ��� �� ���������� �������, ��� � ���� ������ �� ��������� �������� ���'
DEFINE BAR 10 OF raschet_proz_42308 PROMPT '\-'
DEFINE BAR 11 OF raschet_proz_42308 PROMPT '�������� ������������ ���������' ;
  MESSAGE '����� ��� ������ � ��������� ������� ������� �� ������ � �������� ������������ ���������'
DEFINE BAR 12  OF raschet_proz_42308 PROMPT '\-'
DEFINE BAR 13  OF raschet_proz_42308 PROMPT '������ ������������ ���������' ;
  MESSAGE '����� ��� ������ � ��������� ������� ������� �� ������ � ������ ������������ ���������'

ON SELECTION BAR 1  OF raschet_proz_42308 DO  sel_fio IN Raschet_proz_42308
ON SELECTION BAR 3  OF raschet_proz_42308 DO  sum_proz IN Raschet_proz_42308
ON SELECTION BAR 5  OF raschet_proz_42308 DO  vibor_proz IN Raschet_proz_42308
ON SELECTION BAR 7  OF raschet_proz_42308 DO  read_sum IN Raschet_proz_42308
ON SELECTION BAR 9  OF raschet_proz_42308 DO  del_data IN Raschet_proz_42308
ON SELECTION BAR 11 OF raschet_proz_42308 DO  sel_data IN Raschet_proz_42308
ON SELECTION BAR 13  OF raschet_proz_42308 DO  sel_data IN Raschet_proz_42308


**************************************************************** DEFINE POPUP blank_proz ****************************************************************

DEFINE POPUP blank_proz  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF blank_proz PROMPT '��������� �� �������� ����� 47411' ;
  MESSAGE '������������ ������ ��������� �� �������� ����� 47411'
ON SELECTION BAR 1  OF blank_proz DO start IN Zajava_proz


***************************************************************** DEFINE POPUP servis_proz   **************************************************************

DEFINE POPUP servis_proz  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF servis_proz PROMPT '���������� �������� ������ �������' ;
  MESSAGE '�������� ����������� �������� ������ �������'
DEFINE BAR 2  OF servis_proz PROMPT '\-'
DEFINE BAR 3  OF servis_proz PROMPT '���������� ��������� ������ �������' ;
  MESSAGE '�������������� ��������� ������ �������'
DEFINE BAR 4  OF servis_proz PROMPT '\-'
DEFINE BAR 5  OF servis_proz PROMPT '���������� �� �������� ������' ;
  MESSAGE '�������� ����������� �� ������'
DEFINE BAR 6  OF servis_proz PROMPT '\-'
DEFINE BAR 7  OF servis_proz PROMPT '���������� ����� ��������' ;
  MESSAGE '�������� ����������� ����� ��������'
DEFINE BAR 8  OF servis_proz PROMPT '\-'
DEFINE BAR 9  OF servis_proz PROMPT '���������� ����������' ;
  MESSAGE '�������� � �������������� ����������� ����������'
DEFINE BAR 10 OF servis_proz PROMPT '\-'
DEFINE BAR 11 OF servis_proz PROMPT '���������� ������ �����' ;
  MESSAGE '�������� � �������������� ����������� ������ ����� �� ��������� ����'
DEFINE BAR 12 OF servis_proz PROMPT '\-'
DEFINE BAR 13 OF servis_proz PROMPT '���������� ���������� ������' ;
  MESSAGE '�������������� ����������� ���������� ������'
DEFINE BAR 14 OF servis_proz PROMPT '\-'
DEFINE BAR 15 OF servis_proz PROMPT '������� ���������� ���������� �� ��������' ;
  MESSAGE '������������ ����������� ��������� ������ � �������������� ��������� � �������� ��������� ���������� ���������� �� ��������'
DEFINE BAR 16 OF servis_proz PROMPT '\-'
DEFINE BAR 17 OF servis_proz PROMPT '���������� �������� ������' ;
  MESSAGE '�������������� ����������� �������� ������ �� ������ � ������������ ���� Visa'

DO CASE
  CASE pr_rabota_sql = .T.  && ������ ���������� � ��������� SQL Server

    ON SELECTION BAR 1   OF servis_proz DO spr_client IN Servis_sql
    ON SELECTION BAR 3   OF servis_proz DO read_sks IN Servis_sql
    ON SELECTION BAR 5   OF servis_proz DO spr_karta IN Servis_sql
    ON SELECTION BAR 7   OF servis_proz DO spr_kod_oper  IN Servis_sql
    ON SELECTION BAR 9   OF servis_proz DO spravliz  IN Servis_sql
    ON SELECTION BAR 11 OF servis_proz DO red_spr_val IN Servis_sql
    ON SELECTION BAR 13 OF servis_proz DO read_tran IN Servis_sql
    ON SELECTION BAR 15 OF servis_proz DO scan_data_dog IN Raschet_proz_42301
    ON SELECTION BAR 17 OF servis_proz DO read_tarif IN Servis_sql

  CASE pr_rabota_sql = .F.  && ������ ���������� � ���������� ���������

    ON SELECTION BAR 1   OF servis_proz DO spr_client IN Servis
    ON SELECTION BAR 3   OF servis_proz DO read_sks IN Servis
    ON SELECTION BAR 5   OF servis_proz DO spr_karta IN Servis
    ON SELECTION BAR 7   OF servis_proz DO spr_kod_oper  IN Servis
    ON SELECTION BAR 9   OF servis_proz DO spravliz  IN Servis
    ON SELECTION BAR 11 OF servis_proz DO red_spr_val IN Servis
    ON SELECTION BAR 13 OF servis_proz DO read_tran IN Servis
    ON SELECTION BAR 15 OF servis_proz DO scan_data_dog IN Raschet_proz_42301
    ON SELECTION BAR 17 OF servis_proz DO read_tarif IN Servis

ENDCASE


* ========================== ������������ ��������������� ���� ��� ������ �������� ���� "��������� � ��������� ���������" ====================== *


************************************************************* DEFINE MENU servis_glav *******************************************************************

DEFINE MENU servis_glav COLOR SCHEME 2 ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B'

DEFINE PAD sprav_serv OF servis_glav PROMPT '�����������' COLOR SCHEME 2
DEFINE PAD servis_bik OF servis_glav PROMPT '���������� ���' COLOR SCHEME 2 SKIP FOR wes < 5
DEFINE PAD import_kurs OF servis_glav PROMPT '����� �����' COLOR SCHEME 2 SKIP FOR wes < 5
DEFINE PAD read_data OF servis_glav PROMPT '������ ������' COLOR SCHEME 2 SKIP FOR wes < 5
DEFINE PAD read_card_nls OF servis_glav PROMPT '������ �� �������' COLOR SCHEME 2 SKIP FOR wes < 5
DEFINE PAD servis_serv OF servis_glav PROMPT '������������' COLOR SCHEME 2 SKIP FOR wes < 5
DEFINE PAD quit OF servis_glav PROMPT '�����' COLOR SCHEME 2 ;
  MESSAGE '����� � ������� ����'

ON PAD sprav_serv OF servis_glav ACTIVATE POPUP sprav_serv
ON PAD servis_bik OF servis_glav ACTIVATE POPUP servis_bik
ON PAD import_kurs OF servis_glav ACTIVATE POPUP import_kurs
ON PAD read_data OF servis_glav ACTIVATE POPUP read_data
ON PAD read_card_nls OF servis_glav ACTIVATE POPUP read_card_nls
ON PAD servis_serv OF servis_glav ACTIVATE POPUP servis_serv
ON SELECTION PAD quit OF servis_glav DEACTIVATE MENU servis_glav



***************************************************************** DEFINE POPUP sprav_serv ****************************************************************

DEFINE POPUP sprav_serv ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF sprav_serv PROMPT '���������� ����� ��������' ;
  MESSAGE '�������� ����������� ����� ��������'
DEFINE BAR 2  OF sprav_serv PROMPT '\-'
DEFINE BAR 3  OF sprav_serv PROMPT '���������� ����������' ;
  MESSAGE '�������� � �������������� ����������� ����������'
DEFINE BAR 4  OF sprav_serv PROMPT '\-'
DEFINE BAR 5  OF sprav_serv PROMPT '���������� ���������� �������' ;
  MESSAGE '�������� � �������������� ����������� ���������� �������'
DEFINE BAR 6  OF sprav_serv PROMPT '\-'
DEFINE BAR 7  OF sprav_serv PROMPT '���������� ������ �����' ;
  MESSAGE '�������� � �������������� ����������� ������ ����� �� ��������� ����'
DEFINE BAR 8  OF sprav_serv PROMPT '\-'
DEFINE BAR 9  OF sprav_serv PROMPT '���������� ���������� ������' ;
  MESSAGE '�������������� ����������� ���������� ������'
DEFINE BAR 10  OF sprav_serv PROMPT '\-'
DEFINE BAR 11  OF sprav_serv PROMPT '���������� �������� ������ �������' ;
  MESSAGE '�������� ����������� �������� ������ �������'
DEFINE BAR 12 OF sprav_serv PROMPT '\-'
DEFINE BAR 13 OF sprav_serv PROMPT '���������� ��������� ������ �������' ;
  MESSAGE '�������������� ��������� ������ �������'
DEFINE BAR 14 OF sprav_serv PROMPT '\-'
DEFINE BAR 15 OF sprav_serv PROMPT '���������� �� �������� ������' ;
  MESSAGE '�������� ����������� �� ������'
DEFINE BAR 16 OF sprav_serv PROMPT '\-'
DEFINE BAR 17 OF sprav_serv PROMPT '���������� ������ �� ��������' ;
  MESSAGE '�������� ����������� �� ���� ��������'
DEFINE BAR 18  OF sprav_serv PROMPT '\-'
DEFINE BAR 19  OF sprav_serv PROMPT '���������� ������ �� ���� ��������' ;
  MESSAGE '�������� ����������� ���������� ���� �������� �������� � �� "CARD-����"'
DEFINE BAR 20  OF sprav_serv PROMPT '\-'
DEFINE BAR 21  OF sprav_serv PROMPT '���������� ������� ������ � ��� RsBank' ;
  MESSAGE '�������� ����������� ������� ������ � ��� RsBank'
DEFINE BAR 22 OF sprav_serv PROMPT '\-'
DEFINE BAR 23 OF sprav_serv PROMPT '���������� ������ �� ����������� "CARD-����"' ;
  MESSAGE '�������� � �������������� ����������� ������ �� ����������� "CARD-����", ����� ��� ��� � ���'
DEFINE BAR 24 OF sprav_serv PROMPT '\-'
DEFINE BAR 25 OF sprav_serv PROMPT '���������� ����������� ������ � �� "CARD-����"' ;
  MESSAGE '���������� ����������� ������ � �� "CARD-����"'
DEFINE BAR 26 OF sprav_serv PROMPT '\-'
DEFINE BAR 27 OF sprav_serv PROMPT '���������� �������� ������' ;
  MESSAGE '�������������� ����������� �������� ������ �� ������ � ������������ ���� Visa'
DEFINE BAR 28 OF sprav_serv PROMPT '\-'
DEFINE BAR 29 OF sprav_serv PROMPT '���������� ������ ����������� ���� �������' ;
  MESSAGE '���������� ������ ����������� ���� �������'
DEFINE BAR 30 OF sprav_serv PROMPT '\-'
DEFINE BAR 31 OF sprav_serv PROMPT '���������� �����������' ;
  MESSAGE '�������� ����������� �����������'

DO CASE
  CASE pr_rabota_sql = .T.  && ������ ���������� � ��������� SQL Server

* WAIT WINDOW '������ ���������� � ��������� SQL Server'

    ON SELECTION BAR 1 OF sprav_serv DO spr_kod_oper  IN Servis_sql
    ON SELECTION BAR 3 OF sprav_serv DO spravliz  IN Servis_sql
    ON SELECTION BAR 5 OF sprav_serv DO rekvizit IN Servis_sql
    ON SELECTION BAR 7 OF sprav_serv DO red_spr_val IN Servis_sql
    ON SELECTION BAR 9 OF sprav_serv DO read_tran IN Servis_sql
    ON SELECTION BAR 11 OF sprav_serv DO spr_client  IN Servis_sql
    ON SELECTION BAR 13 OF sprav_serv DO read_sks IN Servis_sql
    ON SELECTION BAR 15 OF sprav_serv DO spr_karta IN Servis_sql
    ON SELECTION BAR 17 OF sprav_serv DO list_client IN Servis_sql
    ON SELECTION BAR 19 OF sprav_serv DO red_sprfirma IN Servis_sql
    ON SELECTION BAR 21 OF sprav_serv DO vid_liz_visa IN Servis_sql
    ON SELECTION BAR 23 OF sprav_serv DO spravlizterm  IN Servis_sql
    ON SELECTION BAR 25 OF sprav_serv DO read_slovar  IN Servis
    ON SELECTION BAR 27 OF sprav_serv DO read_tarif IN Servis
    ON SELECTION BAR 29 OF sprav_serv DO read_pass_dover IN Servis
    ON SELECTION BAR 31 OF sprav_serv DO sprav_terror  IN Servis

  CASE pr_rabota_sql = .F.  && ������ ���������� � ���������� ���������

* WAIT WINDOW '������ ���������� � ���������� ���������'

    ON SELECTION BAR 1 OF sprav_serv DO spr_kod_oper  IN Servis
    ON SELECTION BAR 3 OF sprav_serv DO spravliz  IN Servis
    ON SELECTION BAR 5 OF sprav_serv DO rekvizit IN Servis
    ON SELECTION BAR 7 OF sprav_serv DO red_spr_val IN Servis
    ON SELECTION BAR 9 OF sprav_serv DO read_tran IN Servis
    ON SELECTION BAR 11 OF sprav_serv DO spr_client  IN Servis
    ON SELECTION BAR 13 OF sprav_serv DO read_sks IN Servis
    ON SELECTION BAR 15 OF sprav_serv DO spr_karta IN Servis
    ON SELECTION BAR 17 OF sprav_serv DO list_client IN Servis
    ON SELECTION BAR 19 OF sprav_serv DO red_sprfirma IN Servis
    ON SELECTION BAR 21 OF sprav_serv DO vid_liz_visa IN Servis
    ON SELECTION BAR 23 OF sprav_serv DO spravlizterm  IN Servis
    ON SELECTION BAR 25 OF sprav_serv DO read_slovar  IN Servis
    ON SELECTION BAR 27 OF sprav_serv DO read_tarif IN Servis
    ON SELECTION BAR 29 OF sprav_serv DO read_pass_dover IN Servis
    ON SELECTION BAR 31 OF sprav_serv DO sprav_terror  IN Servis

ENDCASE


**************************************************************** DEFINE POPUP servis_bik *****************************************************************

DEFINE POPUP servis_bik ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF servis_bik PROMPT '������ ����������� ��� �� Rs-Bank' ;
  MESSAGE '������ ������ ����������� ��� �� Rs-Bank � ������������ ����������� ����������� ���'
DEFINE BAR 2  OF servis_bik PROMPT '\-'
DEFINE BAR 3  OF servis_bik PROMPT '�������� ���������� ����������� ���' ;
  MESSAGE '��������� �������� ���������� ����������� ���'
DEFINE BAR 4  OF servis_bik PROMPT '\-'
DEFINE BAR 5  OF servis_bik PROMPT '������� ������ � ����������� ��� � RsBank' ;
  MESSAGE '�������� ���������� ������������� � ����������� ��� � RsBank'

ON SELECTION BAR 1 OF servis_bik DO import_bik IN Bnkseek
ON SELECTION BAR 3 OF servis_bik DO brows_bik IN Bnkseek
ON SELECTION BAR 5 OF servis_bik DO create_bankdprt IN Bnkseek

**************************************************************** DEFINE POPUP import_kurs ****************************************************************

DEFINE POPUP import_kurs ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF import_kurs PROMPT '������ ������ ����� �� Rs-Bank' ;
  MESSAGE '������ ������ ����������� ������ ����� �� Rs-Bank � ���������� ����������� ����������� ������ �����'
DEFINE BAR 2  OF import_kurs PROMPT '\-'
DEFINE BAR 3  OF import_kurs PROMPT '�������� ���������� ����������� ������' ;
  MESSAGE '��������� �������� ���������� ����������� ������ �����'
DEFINE BAR 4  OF import_kurs PROMPT '\-'
DEFINE BAR 5  OF import_kurs PROMPT '������� ������ � ����������� ������ ����� � RsBank' ;
  MESSAGE '�������� ���������� ������������� � ����������� ������ ����� � RsBank'

ON SELECTION BAR 1 OF import_kurs DO import_kurs IN Impkursval
ON SELECTION BAR 3 OF import_kurs DO brows_kurs IN Impkursval
ON SELECTION BAR 5 OF import_kurs DO create_kurs IN Impkursval


****************************************************************** DEFINE POPUP read_data **************************************************************

DEFINE POPUP read_data ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1 OF read_data PROMPT '������� �������� � �������� �� ������ CARD-���� � �����������' ;
  MESSAGE '������� �������� � �������� �� ������ CARD-���� � �����������'
DEFINE BAR 2 OF read_data PROMPT '\-'
DEFINE BAR 3 OF read_data PROMPT '������� �������� � �������� �� ������ � ��������� ��������' ;
  MESSAGE '������� �������� � �������� �� ������ � ��������� ��������'
DEFINE BAR 4 OF read_data PROMPT '\-'
DEFINE BAR 5 OF read_data PROMPT '������������ �������� � �������� �� ��������� ������' ;
  MESSAGE '������������ �������� � �������� �� ��������� ������'
DEFINE BAR 6 OF read_data PROMPT '\-'
DEFINE BAR 7 OF read_data PROMPT '������� ��������� �������� �� ������ ��������� ��������� �� ���' ;
  MESSAGE '������� ��������� �������� �� ������ ����� ���������� � ������� ��������� ��������� �� ���'
DEFINE BAR 8 OF read_data PROMPT '\-'
DEFINE BAR 9 OF read_data PROMPT '������� ����������� ������ �� ������ ��������� ��������� �� ���' ;
  MESSAGE '������� ����������� ������ �� ������ ����� ���������� � ������� ��������� ��������� �� ���'
DEFINE BAR 10 OF read_data PROMPT '\-'
DEFINE BAR 11 OF read_data PROMPT '������ ���������� ��������� ������ � ����������� �������� ����' ;
  MESSAGE '������ ���������� ��������� ������ � ����������� �������� ���� ������� ������������ ����������� �������� ����'
DEFINE BAR 12 OF read_data PROMPT '\-'
DEFINE BAR 13 OF read_data PROMPT '��������� �� ��������� �� 12 ������� ���� �� ������������ ����� � 01.07.2006 ����' ;
  MESSAGE '������������ ��������� �� ��������� �� 12 ������� ���� �� ������������ ����� � 01.07.2006 ����'
DEFINE BAR 14 OF read_data PROMPT '\-'
DEFINE BAR 15 OF read_data PROMPT '��������� �� ��������� �� 12 ������� ���� �� ������������ ����� � 01.01.2006 ����' ;
  MESSAGE '������������ ��������� �� ��������� �� 12 ������� ���� �� ������������ ����� � 01.01.2006 ����'
DEFINE BAR 16 OF read_data PROMPT '\-'
DEFINE BAR 17 OF read_data PROMPT '���������� ���������� ����� � ����������� ���������' ;
  MESSAGE '���������� ���������� ����� � ����������� ��������� ������� ������������ ����������� �������� � ������'
DEFINE BAR 18 OF read_data PROMPT '\-'
DEFINE BAR 19 OF read_data PROMPT '���������� ��������� ����������� ���������� ����� ���� �� ������������ ������' ;
  MESSAGE '������� ��������� ��������� ��� ������������� ������ ���������'
DEFINE BAR 20 OF read_data PROMPT '\-'
DEFINE BAR 21 OF read_data PROMPT '�������������� �������� ����� � ����������� ��������� � ������ �����' ;
  MESSAGE '�������������� �������� ����� � ����������� ��������� � ������ ����� ������� ������������ ����������� ������������ ���������'
DEFINE BAR 22 OF read_data PROMPT '\-'
DEFINE BAR 23 OF read_data PROMPT '�������� ����������� ��������� ��� �������������� ������� �� �����' ;
  MESSAGE '�������� ����������� ��������� �� �������� ����� ��� �������� �������������� �������� ����� �� �����'
DEFINE BAR 24 OF read_data PROMPT '\-'
DEFINE BAR 25 OF read_data PROMPT '�������������� ������� � ����������� ����������� ��������� �� �����' ;
  MESSAGE '�������������� ������� � ����������� ����������� ��������� �� �����'
DEFINE BAR 26 OF read_data PROMPT '\-'
DEFINE BAR 27 OF read_data PROMPT '�������� ����������� ��������� ����� �������������� ������� �� �����' ;
  MESSAGE '�������� ����������� ��������� �� �������� ����� ����� �������� �������������� �������� ����� �� �����'
DEFINE BAR 28 OF read_data PROMPT '\-'
DEFINE BAR 29 OF read_data PROMPT '������ ������� ��������� ������ � ����������� ���������' ;
  MESSAGE '������ ������� ��������� ������ � ����������� ��������� ������� ������������ ����������� ����'
DEFINE BAR 30 OF read_data PROMPT '\-'
DEFINE BAR 31 OF read_data PROMPT '������� ������� �������� �� ������ ���������� �� ��������� �����' ;
  MESSAGE '������� ������� �������� ������� ������������ ������ ���������� �� ��������� �����'
DEFINE BAR 32 OF read_data PROMPT '\-'
DEFINE BAR 33 OF read_data PROMPT '������� ��������� �������� �� ������ ���������� �� ��������� �����' ;
  MESSAGE '������� ��������� �������� ������� ������������ ������ ���������� �� ��������� �����'
DEFINE BAR 34 OF read_data PROMPT '\-'
DEFINE BAR 35 OF read_data PROMPT '���������� ������ �� ����������� ������ �������� ���� � ��������� ��������' ;
  MESSAGE '���������� ������������� ������ �� �������� ��� � ��������� �������� � ��������� ���������� �������� ���'
DEFINE BAR 36 OF read_data PROMPT '\-'
DEFINE BAR 37 OF read_data PROMPT '�������� ������� � ������� "��������� ��������" �� ��������� ����' ;
  MESSAGE '������ ��������� ��������� �������� ������� � ������� "��������� ��������" �� ��������� ����'
DEFINE BAR 38 OF read_data PROMPT '\-'
DEFINE BAR 39 OF read_data PROMPT '�������� ���� ������ �� ���������� ������� � ����� �� ���� ��������' ;
  MESSAGE '�������� ���� ������ �� ���������� ������� � ����� �� ���� ��������'
DEFINE BAR 40 OF read_data PROMPT '\-'
DEFINE BAR 41 OF read_data PROMPT '�������� �������� �������� ����� ���� � ����������� ���� � ��������� ������' ;
  MESSAGE '�������� �������� �������� ����� ���� � ����������� ���� � ��������� ������'
DEFINE BAR 42 OF read_data PROMPT '\-'
DEFINE BAR 43 OF read_data PROMPT '�������� ���� ������ � ������������ �������� � ��������� ������' ;
  MESSAGE '�������� ���� ������ � ����������� �������� � ����������� ���� � ��������� ������'

ON SELECTION BAR 1 OF read_data DO scan_acc_ostat IN Servis
ON SELECTION BAR 3 OF read_data DO scan_ved_ostat IN Servis
ON SELECTION BAR 5 OF read_data DO scan_ostat_data IN Servis
ON SELECTION BAR 7 OF read_data DO scan_oborot_ved IN Prav_vedom_ost
ON SELECTION BAR 9 OF read_data DO scan_account IN Prav_account
ON SELECTION BAR 11 OF read_data DO scan_account IN Prav_acc_del
ON SELECTION BAR 13 OF read_data DO start IN Formir_12_obslug
ON SELECTION BAR 15 OF read_data DO start IN Formir_12_obslug_pol
ON SELECTION BAR 17 OF read_data DO start_read_client IN Import_cln
ON SELECTION BAR 19 OF read_data DO red_obsluga_god IN Servis
ON SELECTION BAR 21 OF read_data DO read_adres IN Adres_client
ON SELECTION BAR 23 OF read_data DO export_zajava_b IN Adres_client
ON SELECTION BAR 25 OF read_data DO brows_zajava_b IN Adres_client
ON SELECTION BAR 27 OF read_data DO import_zajava_b IN Adres_client
ON SELECTION BAR 29 OF read_data DO read_card_acct IN Prav_card_acct
ON SELECTION BAR 31 OF read_data DO start IN Scan_istor_ost
ON SELECTION BAR 33 OF read_data DO start IN Scan_vedom_ost
ON SELECTION BAR 35 OF read_data DO start IN Vihodnoy_vedom_ost
ON SELECTION BAR 37 OF read_data DO start_del IN Del_vedom_ost
ON SELECTION BAR 39 OF read_data DO delete_card_acct_all_tables IN Servis
ON SELECTION BAR 41 OF read_data DO acc_del_pack IN Import_acc
ON SELECTION BAR 43 OF read_data DO del_client_account IN Pack_table


*************************************************************** DEFINE POPUP read_card_nls **************************************************************

DEFINE POPUP read_card_nls ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1 OF read_card_nls PROMPT '��������� ���������� ������ �� ������������ � 154 �� �����' ;
  MESSAGE '��������� ���������� ������ � ������������� �� ������������ � 154 �� �����'
DEFINE BAR 2 OF read_card_nls PROMPT '\-'
DEFINE BAR 3 OF read_card_nls PROMPT '��������� ���������� ������ � ����� �� ������������ �����' ;
  MESSAGE '��������� ���������� ������ � ����� �� ������������ �����'
DEFINE BAR 4 OF read_card_nls PROMPT '\-'
DEFINE BAR 5 OF read_card_nls PROMPT '������������ ��������� ������ �� ������ ���������' ;
  MESSAGE '������������ ��������� ������ �� ������ ���������'
DEFINE BAR 6 OF read_card_nls PROMPT '\-'
DEFINE BAR 7 OF read_card_nls PROMPT '�������� �� ������������ ��������� ������' ;
  MESSAGE '�������� �� ������������ ��������� ������'
DEFINE BAR 8 OF read_card_nls PROMPT '\-'
DEFINE BAR 9 OF read_card_nls PROMPT '�������� ���������� ����� ��� ����������� ��������� ������' ;
  MESSAGE '�������� ���������� ����� ��� ����������� ��������� ������'
DEFINE BAR 10 OF read_card_nls PROMPT '\-'
DEFINE BAR 11 OF read_card_nls PROMPT '��������� ������ ������ �� ����� �� ���� ��������' ;
  MESSAGE '��������� ������ ������ �� ����� �� ���� �������� ������� ������������'
DEFINE BAR 12 OF read_card_nls PROMPT '\-'
DEFINE BAR 13 OF read_card_nls PROMPT '�������� �������� �� ��������� ������ � ������������ SMS-�������' ;
  MESSAGE '�������� �������� �� ��������� ������ � ������������ SMS-�������, ���� ������ �������, ��������� ���'
DEFINE BAR 14 OF read_card_nls PROMPT '\-'
DEFINE BAR 15 OF read_card_nls PROMPT '���������� ���� ����������� ����� � ������� "��������� ��������"' ;
  MESSAGE '���������� ���� ����������� ����� � ������� "��������� ��������" ������� ������������ ����������� ������'
DEFINE BAR 16 OF read_card_nls PROMPT '\-'
DEFINE BAR 17 OF read_card_nls PROMPT '���������� �� �������� ������� ��������� � ������� ������� ���������' ;
  MESSAGE '���������� �� �������� ������� ��������� � ������� ������� ��������� �� ��������� ����������� �����'

DO CASE
  CASE pr_rabota_sql = .T.  && ������ ���������� � ��������� SQL Server

    ON SELECTION BAR 1 OF read_card_nls DO red_bal_num_new IN Servis_sql
    ON SELECTION BAR 3 OF read_card_nls DO red_bal_num_str IN Servis_sql
    ON SELECTION BAR 5 OF read_card_nls DO scan_formir_card_nls IN Formir_card_nls
    ON SELECTION BAR 7 OF read_card_nls DO scan_dubl_card_nls IN Dubl_card_nls
    ON SELECTION BAR 9 OF read_card_nls DO start IN Konvert_card_nls
    ON SELECTION BAR 11 OF read_card_nls DO start IN Update_new_card_acct
    ON SELECTION BAR 13 OF read_card_nls DO start IN Smshome
    ON SELECTION BAR 15 OF read_card_nls DO start IN Tran_vedom_ost
    ON SELECTION BAR 17 OF read_card_nls DO start IN Dobav_zajava_arx

  CASE pr_rabota_sql = .F.  && ������ ���������� � ���������� ���������

    ON SELECTION BAR 1 OF read_card_nls DO red_bal_num_new IN Servis
    ON SELECTION BAR 3 OF read_card_nls DO red_bal_num_str IN Servis
    ON SELECTION BAR 5 OF read_card_nls DO scan_formir_card_nls IN Formir_card_nls
    ON SELECTION BAR 7 OF read_card_nls DO scan_dubl_card_nls IN Dubl_card_nls
    ON SELECTION BAR 9 OF read_card_nls DO start IN Konvert_card_nls
    ON SELECTION BAR 11 OF read_card_nls DO start IN Update_new_card_acct
    ON SELECTION BAR 13 OF read_card_nls DO start IN Smshome
    ON SELECTION BAR 15 OF read_card_nls DO start IN Tran_vedom_ost
    ON SELECTION BAR 17 OF read_card_nls DO start IN Dobav_zajava_arx

ENDCASE


**************************************************************** DEFINE POPUP servis_serv *****************************************************************

DEFINE POPUP servis_serv ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF servis_serv PROMPT '������� �������� ���� � �����' ;
  MESSAGE '��������� �������� �������� ���� � �����'
DEFINE BAR 2  OF servis_serv PROMPT '\-'
DEFINE BAR 3  OF servis_serv PROMPT '��������� ��������� ����� � ��������' ;
  MESSAGE '��������� ��������� ����� � ��������'
DEFINE BAR 4  OF servis_serv PROMPT '\-'
DEFINE BAR 5  OF servis_serv PROMPT '�������� � ������������ ������ �� "CARD-����"' ;
  MESSAGE '�������� � ������������ ������ �� "CARD-����"'
DEFINE BAR 6  OF servis_serv PROMPT '\-'
DEFINE BAR 7  OF servis_serv PROMPT '�������� �������� ���� ����� � ����� �����' SKIP ;
  MESSAGE '�������� �������� ���� ����� � ����� ����� Electron ��� 1; Classic ��� 2; Gold ��� 3; Business ��� 4'
DEFINE BAR 8  OF servis_serv PROMPT '\-'
DEFINE BAR 9  OF servis_serv PROMPT '������� �������� � �������� �� ������ �� �������� ����' ;
  MESSAGE '��������� ������� �������� � �������� �� ������ ������� �������� ����������� ���������� �� �������� ����'
DEFINE BAR 10  OF servis_serv PROMPT '\-'
DEFINE BAR 11  OF servis_serv PROMPT '�������� ������ ����������� ��������� � �������� ����' ;
  MESSAGE '�������� ������ ����������� ��������� � �������� ����'
DEFINE BAR 12 OF servis_serv PROMPT '\-'
DEFINE BAR 13 OF servis_serv PROMPT '�������� ����� ���������� ����� �����' ;
  MESSAGE '� �������� ������ ����������� ����� ���� � ������� �������, ��������� � ������� modify.dbf'
DEFINE BAR 14 OF servis_serv PROMPT '\-'
DEFINE BAR 15 OF servis_serv PROMPT '�������� ������� � ��������� ������' ;
  MESSAGE '�������� ������� � ��������� ������ errorlog.dbf'
DEFINE BAR 16 OF servis_serv PROMPT '\-'
DEFINE BAR 17 OF servis_serv PROMPT '��������� �������� ���������� ������ �������' ;
  MESSAGE '��������� �������� ���������� ������ ������� �� ����������� �����'
DEFINE BAR 18 OF servis_serv PROMPT '\-'
DEFINE BAR 19 OF servis_serv PROMPT '������� ������ �� ��������� ���� � �������� �������' ;
  MESSAGE '������� ������ �� ��������� ���� �� ������� ������ � �������� �������'
DEFINE BAR 20 OF servis_serv PROMPT '\-'
DEFINE BAR 21 OF servis_serv PROMPT '������ ������ �� ��������� ���� �� �������� ������' ;
  MESSAGE '������ ������ �� ��������� ���� �� �������� ������ � ������� �������'
DEFINE BAR 22 OF servis_serv PROMPT '\-'
DEFINE BAR 23 OF servis_serv PROMPT '������ ��������� �������� �� ������� ��������' ;
  MESSAGE '������ ��������� �������� �� ������� �������� �������������� �����'
DEFINE BAR 24 OF servis_serv PROMPT '\-'
DEFINE BAR 25 OF servis_serv PROMPT '�������� ������������� ������� � ��������' ;
  MESSAGE '�������� ������������� ������� � ��������'
DEFINE BAR 26 OF servis_serv PROMPT '\-'
DEFINE BAR 27 OF servis_serv PROMPT '��������� �������� ������� ����� � ���������' ;
  MESSAGE '��������� �������� ������� ����� � ���������, ������������ � ���������� ��������� ������������ ������� � ��������'
DEFINE BAR 28 OF servis_serv PROMPT '\-'
DEFINE BAR 29 OF servis_serv PROMPT '���������� ����������� ��� ������������ � ��' ;
  MESSAGE '���������� ����������� ��� ������������ � ��'
DEFINE BAR 30 OF servis_serv PROMPT '\-'
DEFINE BAR 31 OF servis_serv PROMPT '������ ����� ������� ������������� ���� �����' ;
  MESSAGE '������ ����� ������� ������������� ������������� ���� CARD-�����'
DEFINE BAR 32 OF servis_serv PROMPT '\-'
DEFINE BAR 33 OF servis_serv PROMPT '��������� ��������� ������ � ����������� ��������' ;
  MESSAGE '��������� ��������� �� ������ � �������, ������������ ����������� ��������'
DEFINE BAR 34 OF servis_serv PROMPT '\-'
DEFINE BAR 35 OF servis_serv PROMPT '�������������� ������������ ������ ��' ;
  MESSAGE '�������������� ������������ ������ �� "CARD-����"'

DO CASE
  CASE pr_rabota_sql = .T.  && ������ ���������� � ��������� SQL Server

    DEFINE BAR 36 OF servis_serv PROMPT '\-'
    DEFINE BAR 37 OF servis_serv PROMPT '������� ������ �� ��������� ������ � ������� SQL Server' ;
      MESSAGE '������� ������ �� ��������� ������ � ������� SQL Server 2000 �  SQL Server 2005'

    ON SELECTION BAR 1 OF servis_serv DO acc_del_pack IN Servis_sql
    ON SELECTION BAR 3 OF servis_serv DO start IN Read_table
    ON SELECTION BAR 5 OF servis_serv DO start IN Pack_table
    ON SELECTION BAR 7 OF servis_serv DO start IN Type_card_acct
    ON SELECTION BAR 9 OF servis_serv DO start_ostat_sql IN Prav_vedom_ost
    ON SELECTION BAR 11 OF servis_serv DO start IN Import_sprav
    ON SELECTION BAR 13 OF servis_serv DO start_pak IN Read_table
    ON SELECTION BAR 15 OF servis_serv DO start_del IN Errtrap
    ON SELECTION BAR 17 OF servis_serv DO red_stavka IN Servis_sql
    ON SELECTION BAR 19 OF servis_serv DO start IN Export_table_arxiv
    ON SELECTION BAR 21 OF servis_serv DO start IN Import_table_arxiv
    ON SELECTION BAR 23 OF servis_serv DO start IN Prav_vedom_ost
    ON SELECTION BAR 25 OF servis_serv DO start IN Viver_table
    ON SELECTION BAR 27 OF servis_serv DO monopol IN Servis_sql
    ON SELECTION BAR 29 OF servis_serv DO start IN Calendar
    ON SELECTION BAR 31 OF servis_serv DO start IN Usl_card
    ON SELECTION BAR 33 OF servis_serv DO start IN Excelens
    ON SELECTION BAR 35 OF servis_serv DO start_avt_menu IN Pack_table
    ON SELECTION BAR 37 OF servis_serv DO start IN Export_data_sql

  CASE pr_rabota_sql = .F.  && ������ ���������� � ���������� ���������

    ON SELECTION BAR 1 OF servis_serv DO acc_del_pack IN Servis
    ON SELECTION BAR 3 OF servis_serv DO start IN Read_table
    ON SELECTION BAR 5 OF servis_serv DO start IN Pack_table
    ON SELECTION BAR 7 OF servis_serv DO start IN Type_card_acct
    ON SELECTION BAR 9 OF servis_serv DO start_ostat_den IN Prav_vedom_ost
    ON SELECTION BAR 11 OF servis_serv DO start IN Import_sprav
    ON SELECTION BAR 13 OF servis_serv DO start_pak IN Read_table
    ON SELECTION BAR 15 OF servis_serv DO start_del IN Errtrap
    ON SELECTION BAR 17 OF servis_serv DO red_stavka IN Servis
    ON SELECTION BAR 19 OF servis_serv DO start IN Export_table_arxiv
    ON SELECTION BAR 21 OF servis_serv DO start IN Import_table_arxiv
    ON SELECTION BAR 23 OF servis_serv DO start IN Prav_vedom_ost
    ON SELECTION BAR 25 OF servis_serv DO start IN Viver_table
    ON SELECTION BAR 27 OF servis_serv DO monopol IN Servis
    ON SELECTION BAR 29 OF servis_serv DO start IN Calendar
    ON SELECTION BAR 31 OF servis_serv DO start IN Usl_card
    ON SELECTION BAR 33 OF servis_serv DO start IN Excelens
    ON SELECTION BAR 35 OF servis_serv DO start_avt_menu IN Pack_table

ENDCASE


RETURN

*********************************************************************************************************************************************************


*  * ������ ����� �������� ������� ����
*  DEFINE PAD _msm_edit OF _MSYSMENU PROMPT "������" COLOR SCHEME 3
*  * ������ ������ �������� ������� ����

*  DEFINE POPUP _medit MARGIN RELATIVE SHADOW COLOR SCHEME 4
*  DEFINE BAR _MED_UNDO OF _medit PROMPT "��������" ;
*  KEY CTRL+Z, "CTRL+Z"
*  DEFINE BAR _MED_REDO OF _medit PROMPT "�������" ;
*  KEY CTRL+R, "CTRL+R"
*  DEFINE BAR _MED_SP100 OF _medit PROMPT "\-"
*  DEFINE BAR _MED_CUT OF _medit PROMPT "��������" ;
*  KEY CTRL+X, "CTRL+X"
*  DEFINE BAR _MED_COPY OF _medit PROMPT "����������" ;
*  KEY CTRL+C, "CTRL+C"
*  DEFINE BAR _MED_PASTE OF _medit PROMPT "��������" ;
*  KEY CTRL+V, "CTRL+V"
*  DEFINE BAR _MED_CLEAR OF _medit PROMPT "��������"
*  DEFINE BAR _MED_SP200 OF _medit PROMPT "\-"
*  DEFINE BAR _MED_SLCTA OF _medit PROMPT "�������� ���" ;
*  KEY CTRL+A, "CTRL+A"
*  [/code]


*********************************************************************************************************************************************************


