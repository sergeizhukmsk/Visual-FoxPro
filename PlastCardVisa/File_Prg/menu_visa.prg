***********************************************************
*** Процедура формирования меню для карточной системы ***
***********************************************************

PROCEDURE start

* ================================================ Формирование главного меню программы ================================================== *

text_visa_1_1 = 'Движение средств по пластиковой карте'
text_visa_1_2 = 'Выдача карт и управление карточным счетом'
text_visa_1_3 = 'Создание и формирование отчетов в УПК'
text_visa_1_4 = 'Начисление процентов по карточному счету'
text_visa_1_5 = 'Загрузка данных из головного банка'
text_visa_1_6 = 'Выгрузка данных в головной банк'
text_visa_1_7 = 'Сервисные и системные процедуры'
text_visa_1_8 = 'Резервное копирование данных'
text_visa_1_9 = 'Закрытие операционного дня в "CARD-VISA"'
text_visa_1_10 = 'Доступ к справочнику паролей'
text_visa_1_11 = 'Завершение работы с программой'

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
  MESSAGE 'Пpи выбоpе данного пункта пользователь выполняет операции учета движения средств по карте'
DEFINE BAR 2  OF glmenu PROMPT '\-'
DEFINE BAR 3  OF glmenu PROMPT text_visa_1_2 SKIP FOR wes < 4 ;
  MESSAGE 'Пpи выбоpе данного пункта пользователь выполняет операции по работе с клиентом'
DEFINE BAR 4  OF glmenu PROMPT '\-'
DEFINE BAR 5  OF glmenu PROMPT text_visa_1_3 SKIP FOR wes < 4 ;
  MESSAGE 'Пpи выбоpе данного пункта пользователь создает и формирует отчеты в УПК'
DEFINE BAR 6  OF glmenu PROMPT '\-'
DEFINE BAR 7  OF glmenu PROMPT text_visa_1_4 SKIP FOR wes < 4 ;
  MESSAGE 'Пpи выбоpе данного пункта пользователь выполняет операции по начислению процентов по пластиковой карте'
DEFINE BAR 8  OF glmenu PROMPT '\-'
DEFINE BAR 9  OF glmenu PROMPT text_visa_1_5 SKIP FOR wes < 5 ;
  MESSAGE 'Пpи выбоpе данного пункта пpоисходит импорт данных по клиентам, счетам и движению средств по картам'
DEFINE BAR 10 OF glmenu PROMPT '\-'
DEFINE BAR 11 OF glmenu PROMPT text_visa_1_6 SKIP FOR wes < 5 ;
  MESSAGE 'Пpи выбоpе данного пункта пpоисходит выгрузка данных из таблиц для передачи в головной банк'
DEFINE BAR 12  OF glmenu PROMPT '\-'
DEFINE BAR 13 OF glmenu PROMPT text_visa_1_7 ;
  MESSAGE 'Пpи выбоpе данного пункта пользователь выполняет сервисные и системные процедуры'
DEFINE BAR 14 OF glmenu PROMPT '\-'
DEFINE BAR 15 OF glmenu PROMPT text_visa_1_8 SKIP FOR wes < 5 ;
  MESSAGE 'Пpи выбоpе данного пункта пpоисходит резервное копирование рабочих таблиц с данными'
DEFINE BAR 16 OF glmenu PROMPT '\-'
DEFINE BAR 17 OF glmenu PROMPT text_visa_1_9 SKIP FOR wes < 5 ;
  MESSAGE 'Пpи выбоpе данного пункта пpоисходит закрытие операционного дня в ПО "CARD-Visa"'
DEFINE BAR 18 OF glmenu PROMPT '\-'
DEFINE BAR 19 OF glmenu PROMPT text_visa_1_10 ;
  MESSAGE 'Пpи выбоpе данного пункта пpоисходит просмотр справочника паролей'
DEFINE BAR 20 OF glmenu PROMPT '\-'
DEFINE BAR 21 OF glmenu PROMPT text_visa_1_11 ;
  MESSAGE 'Пpи выбоpе данного пункта пpоисходит выход в начальную точку запуска'
ON SELECTION POPUP glmenu DO glmenu IN Visa


* ===================== Формирование горизонтального меню для пункта главного меню "Движение средств по пластиковой карте" ====================== *


************************************************************** DEFINE MENU document ********************************************************************

DEFINE MENU document COLOR SCHEME 2 ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B'

DEFINE PAD document OF document PROMPT 'Операции по карте' COLOR SCHEME 2
DEFINE PAD vedomostat OF document PROMPT 'Ведомость остатков' COLOR SCHEME 2
DEFINE PAD vipiska OF document PROMPT 'Выписки по счетам' COLOR SCHEME 2
DEFINE PAD analit_info OF document PROMPT 'Аналитика' COLOR SCHEME 2
DEFINE PAD otchet_zb OF document PROMPT 'Отчеты ЦБ' COLOR SCHEME 2
DEFINE PAD servis_doc OF document PROMPT 'Справочники' COLOR SCHEME 2
DEFINE PAD quit OF document PROMPT 'Выход' COLOR SCHEME 2 ;
  MESSAGE 'Выход в главное меню'

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
DEFINE BAR 1  OF document PROMPT 'Отчет об операциях по списанию средств из ПО "Transmaster"' ;
  MESSAGE 'Формирование отчета об операциях по расходованию средств на основании данных о движении средств из ПО "Transmaster"'
DEFINE BAR 2  OF document PROMPT '\-'
DEFINE BAR 3  OF document PROMPT 'Отчет об операциях по зачислению средств из ПО "Transmaster"' ;
  MESSAGE 'Формирование отчета об операциях по зачислению средств на ПК на основании данных о движении средств из ПО "Transmaster"'
DEFINE BAR 4  OF document PROMPT '\-'
DEFINE BAR 5  OF document PROMPT 'Отчет об операциях по списанию средств из ПО "CARD-Visa филиал"' ;
  MESSAGE 'Формирование отчета об операциях по расходованию средств на основании данных о движении средств из ПО "CARD-Visa филиал"'
DEFINE BAR 6  OF document PROMPT '\-'
DEFINE BAR 7  OF document PROMPT 'Отчет об операциях по зачислению средств из ПО "CARD-Visa филиал"' ;
  MESSAGE 'Формирование отчета об операциях по зачислению средств на ПК на основании данных о движении средств из ПО "CARD-Visa филиал"'
DEFINE BAR 8  OF document PROMPT '\-'
DEFINE BAR 9  OF document PROMPT 'Отчет о конверсионных операциях по пластиковым карточкам' ;
  MESSAGE 'Отчет о конверсионных операциях по пластиковым карточкам, то есть когда валюта карты отличается от кода валюты транзакции'
DEFINE BAR 10 OF document PROMPT '\-'
DEFINE BAR 11 OF document PROMPT 'Полная опись проведенных операций из ПО "Transmaster"' ;
  MESSAGE 'Формирование отчета о проведенных операциях на основании данных о движении средств из ПО "Transmaster"'
DEFINE BAR 12 OF document PROMPT '\-'
DEFINE BAR 13 OF document PROMPT 'Полная опись проведенных операций из ПО "CARD-Visa филиал"' ;
  MESSAGE 'Формирование отчета о проведенных операциях на основании данных о движении средств из ПО "CARD-Visa филиал"'
DEFINE BAR 14 OF document PROMPT '\-'
DEFINE BAR 15 OF document PROMPT 'Удаление импортированных документов за дату из ПО "Transmaster"' ;
  MESSAGE 'Данная процедура выполняет удаление ранее импортированных документов за выбранную дату из ПО "Transmaster"'
DEFINE BAR 16 OF document PROMPT '\-'
DEFINE BAR 17 OF document PROMPT 'Выборка документов за период по фильтру из ПО "Transmaster"' ;
  MESSAGE 'Возможность фильтрации проведенных документов за период: назначение платежа, счет по дебету, счет по кредиту, сумма проводки'
DEFINE BAR 18 OF document PROMPT '\-'
DEFINE BAR 19 OF document PROMPT 'Формирование отчета по легализации проведенных операций' ;
  MESSAGE 'Возможность выборки проведенных документов за период по суммам равным или превышающим 600 000.00 руб.'

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
DEFINE BAR 1 OF vedomostat PROMPT 'Ведомость остатков за дату по карточным счетам' ;
  MESSAGE 'Формирование контрольной ведомости остатков за выбранную дату по карточным счетам'
DEFINE BAR 2 OF vedomostat PROMPT '\-'
DEFINE BAR 3 OF vedomostat PROMPT 'Оборотная ведомость за дату по карточным счетам' ;
  MESSAGE 'Формирование оборотной ведомости по карточным счетам за выбранную дату'
DEFINE BAR 4 OF vedomostat PROMPT '\-'
DEFINE BAR 5 OF vedomostat PROMPT 'Оборотная ведомость за период по счетам из ПО "Transmaster"' ;
  MESSAGE 'Формирование оборотной ведомости по карточным счетам за диапазон указанных дат из ПО "Transmaster"'
DEFINE BAR 6 OF vedomostat PROMPT '\-'
DEFINE BAR 7 OF vedomostat PROMPT 'Ведомость остатков по типам карт из ПО "Transmaster"' ;
  MESSAGE 'Расчет контрольной ведомости остатков по категориям карт из ПО "Transmaster"'
DEFINE BAR 8 OF vedomostat PROMPT '\-'
DEFINE BAR 9 OF vedomostat PROMPT 'Ведомость остатков по типам карт из ПО "CARD-Visa"' ;
  MESSAGE 'Расчет контрольной ведомости остатков по категориям карт из ПО "CARD-Visa"'
DEFINE BAR 10 OF vedomostat PROMPT '\-'
DEFINE BAR 11 OF vedomostat PROMPT 'Ведомость остатков по типам карт из текстовой оборотки УПК' ;
  MESSAGE 'Расчет контрольной ведомости остатков по категориям карт из текстовой оборотки, присланной из УПК'
DEFINE BAR 12 OF vedomostat PROMPT '\-'
DEFINE BAR 13 OF vedomostat PROMPT 'История остатков по счету клиента из ПО "Transmaster"' ;
  MESSAGE 'Формирование истории остатков по карточному счету клиента из ПО "Transmaster"'
DEFINE BAR 14 OF vedomostat PROMPT '\-'
DEFINE BAR 15 OF vedomostat PROMPT 'История остатков по счету клиента из ПО "CARD-Visa"' ;
  MESSAGE 'Формирование истории остатков по карточному счету клиента из ПО "CARD-Visa"'
DEFINE BAR 16 OF vedomostat PROMPT '\-'
DEFINE BAR 17 OF vedomostat PROMPT 'Ведомость остатков по счету клиента из ПО "CARD-Visa"' ;
  MESSAGE 'Формирование ведомости остатков по карточному счету клиента из ПО "CARD-Visa"'
DEFINE BAR 18 OF vedomostat PROMPT '\-'
DEFINE BAR 19 OF vedomostat PROMPT 'Ведомость остатков по 10 счетам с наибольшим остатком' ;
  MESSAGE 'Формирование ведомости остатков по 10 счетам с наибольшим остатком'
DEFINE BAR 20 OF vedomostat PROMPT '\-'
DEFINE BAR 21 OF vedomostat PROMPT 'Выверка остатков и оборотов по счетам за открытую дату' ;
  MESSAGE 'Процедура выверки остатков и оборотов по счетам методом просчета проведенных документов за открытую дату'
DEFINE BAR 22 OF vedomostat PROMPT '\-'
DEFINE BAR 23 OF vedomostat PROMPT 'Выверка остатков и оборотов по счетам за архивную дату' ;
  MESSAGE 'Процедура выверки остатков и оборотов по счетам методом просчета проведенных документов за архивную дату'
DEFINE BAR 24 OF vedomostat PROMPT '\-'
DEFINE BAR 25 OF vedomostat PROMPT 'Правка ведомости остатков по истории остатков' ;
  MESSAGE 'Правка ведомости остатков по истории остатков сформированной ранее'
DEFINE BAR 26 OF vedomostat PROMPT '\-'
DEFINE BAR 27 OF vedomostat PROMPT 'Правка истории остатков ПО "CARD-виза" по проведенным документам' ;
  MESSAGE 'Правка истории остатков ПО "CARD-виза" по проведенным документам'
DEFINE BAR 28 OF vedomostat PROMPT '\-'
DEFINE BAR 29 OF vedomostat PROMPT 'Счета с дебетовыми оборотами и остатками по ним все коды операций' ;
  MESSAGE 'Формирование отчета по истории остатков по счетам с дебетовыми оборотами и остатками по ним за диапазон дат все коды операций'
DEFINE BAR 30 OF vedomostat PROMPT '\-'
DEFINE BAR 31 OF vedomostat PROMPT 'Счета с дебетовыми оборотами (конвертация, перечисления) и остатками по ним' ;
  MESSAGE 'Формирование отчета по истории остатков по счетам с дебетовыми оборотами (конвертация, перечисления) и остатками по ним за диапазон дат'
DEFINE BAR 32 OF vedomostat PROMPT '\-'
DEFINE BAR 33 OF vedomostat PROMPT 'Счета с дебетовыми оборотами ВСЕ кроме кассы и остатками по ним' ;
  MESSAGE 'Формирование отчета по истории остатков по счетам с дебетовыми оборотами ВСЕ кроме кассы и остатками по ним за диапазон дат'

ON SELECTION BAR 1 OF vedomostat DO ostat_trans IN Docum_ostat
ON SELECTION BAR 3 OF vedomostat DO ostat_trans IN Docum_ostat
ON SELECTION BAR 5 OF vedomostat DO oborot_trans IN Oborot_period
ON SELECTION BAR 7 OF vedomostat DO cardtip_ostat IN Oborot_trans
ON SELECTION BAR 9 OF vedomostat DO cardtip_ostat IN Oborot_makvisa
ON SELECTION BAR 11 OF vedomostat DO cardtip_ostat IN Oborot_upk

DO CASE
  CASE pr_rabota_sql = .T.  && Работа происходит с таблицами SQL Server

    ON SELECTION BAR 13 OF vedomostat DO start IN Istor_ostat_sql
    ON SELECTION BAR 15 OF vedomostat DO start IN Istor_mak_sql
    ON SELECTION BAR 17 OF vedomostat DO vedom_ost IN Vedom_ostat_sql

  CASE pr_rabota_sql = .F.  && Работа происходит с локальными таблицами

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
DEFINE BAR 1 OF vipiska PROMPT 'Выписка по одному выбранному карточному счету' ;
  MESSAGE 'Процедура формирования выписки по одному карточному счету клиента с выборкой по ФИО'
DEFINE BAR 2 OF vipiska PROMPT '\-'
DEFINE BAR 3 OF vipiska PROMPT 'Выписки по карточным счетам с движением за текущий день' ;
  MESSAGE 'Процедура формирования выписки по всем карточным счетам клиентов, у которых было движение по счету на текущую дату опердня и потоковой печатью'
DEFINE BAR 4 OF vipiska PROMPT '\-'
DEFINE BAR 5 OF vipiska PROMPT 'Выписки по карточным счетам с движением за архивный день' ;
  MESSAGE 'Процедура формирования выписки по всем карточным счетам клиентов, у которых было движение по счету за архивную дату опердня и потоковой печатью'
DEFINE BAR 6 OF vipiska PROMPT '\-'
DEFINE BAR 7 OF vipiska PROMPT 'Вывод на экран обнаруженных при загрузке счетов с овердрафтом' ;
  MESSAGE 'Вывод на экран обнаруженных при загрузке из Головного банка счетов с овердрафтом на текущую дату ОДБ'
DEFINE BAR 8 OF vipiska PROMPT '\-'
DEFINE BAR 9 OF vipiska PROMPT 'История возникновения долгов за неразрешенный овердрафт' ;
  MESSAGE 'Работа с историей возникновения долгов за неразрешенный овердрафт по всем счетам за период'

DO CASE
  CASE pr_rabota_sql = .T.  && Работа происходит с таблицами SQL Server

    ON SELECTION BAR 1 OF vipiska DO start_vipis_cln IN Vipiska_sql
    ON SELECTION BAR 3 OF vipiska DO start_vipis_all_den IN Vipiska_sql
    ON SELECTION BAR 5 OF vipiska DO start_vipis_all_arx IN Vipiska_sql
    ON SELECTION BAR 7 OF vipiska DO start IN Overdraft
    ON SELECTION BAR 9 OF vipiska DO start IN Istor_overdraft

  CASE pr_rabota_sql = .F.  && Работа происходит с локальными таблицами

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
DEFINE BAR 1 OF analit_info PROMPT 'Ведомость по полученным доходам (полная)' ;
  MESSAGE 'Расчет полной ведомости по полученным доходам, включая штрафные коды за снятие средств в чужих устройствах'
DEFINE BAR 2 OF analit_info PROMPT '\-'
DEFINE BAR 3 OF analit_info PROMPT 'Ведомость по полученным доходам (не полная)' ;
  MESSAGE 'Расчет полной ведомости по полученным доходам, исключаячая штрафные коды за снятие средств в чужих устройствах без НДС'
DEFINE BAR 4 OF analit_info PROMPT '\-'
DEFINE BAR 5 OF analit_info PROMPT 'Ведомость по произведенным расходам (полная)' ;
  MESSAGE 'Расчет ведомости по счетам расходов для УПК'
DEFINE BAR 6 OF analit_info PROMPT '\-'
DEFINE BAR 7 OF analit_info PROMPT 'Отчет по операциям снятия наличности в устройствах "CARD-банк"' ;
  MESSAGE 'Отчет по проведенным операциям снятия наличности в банкоматах и терминалах принадлежащих КБ "CARD-банк", код операции 207.'
DEFINE BAR 8 OF analit_info PROMPT '\-'
DEFINE BAR 9 OF analit_info PROMPT 'Отчет по операциям снятия наличности в устройствах чужих банков' ;
  MESSAGE 'Отчет по проведенным операциям снятия наличности в банкоматах и терминалах принадлежащих сторонним банкам, код операции 207.'
DEFINE BAR 10 OF analit_info PROMPT '\-'
DEFINE BAR 11 OF analit_info PROMPT 'Отчет по операциям оплаты услуг и товаров в устройствах "CARD-банк"' ;
  MESSAGE 'Отчет по проведенным операциям оплаты услуг и товаров в банкоматах и терминалах принадлежащих КБ "CARD-банк", код операции 205.'
DEFINE BAR 12 OF analit_info PROMPT '\-'
DEFINE BAR 13 OF analit_info PROMPT 'Отчет по операциям оплаты услуг и товаров в устройствах чужих банков' ;
  MESSAGE 'Отчет по проведенным операциям оплаты услуг и товаров в банкоматах и терминалах принадлежащих сторонним банкам, код операции 205.'
DEFINE BAR 14 OF analit_info PROMPT '\-'
DEFINE BAR 15 OF analit_info PROMPT 'Ведомость открытых карточных счетов по ПК VISA' ;
  MESSAGE 'Ведомость открытых карточных счетов по ПК VISA'
DEFINE BAR 16 OF analit_info PROMPT '\-'
DEFINE BAR 17 OF analit_info PROMPT 'Ведомость закрытых карточных счетов по ПК VISA' ;
  MESSAGE 'Ведомость закрытых карточных счетов по ПК VISA'
DEFINE BAR 18 OF analit_info PROMPT '\-'
DEFINE BAR 19 OF analit_info PROMPT 'Проведение выверки оборотной ведомости по счетам филиала и УПК' ;
  MESSAGE 'Проведение выверки оборотной ведомости по счетам филиала и УПК'
DEFINE BAR 20 OF analit_info PROMPT '\-'
DEFINE BAR 21 OF analit_info PROMPT 'Отчет по открытым и закрытым картам за период' ;
  MESSAGE 'Отчет по открытым и закрытым картам, дата открытия которых попадает в выбранный период'
DEFINE BAR 22 OF analit_info PROMPT '\-'
DEFINE BAR 23 OF analit_info PROMPT 'Отчет по открытым и закрытым картам на дату' ;
  MESSAGE 'Отчет по открытым и закрытым картам, дата открытия которых менее выбранной даты'
DEFINE BAR 24 OF analit_info PROMPT '\-'
DEFINE BAR 25 OF analit_info PROMPT 'Отчет по открытым и закрытым счетам клиентов на текущую дату' ;
  MESSAGE 'Отчет по открытым и закрытым счетам клиентов на текущую дату'
DEFINE BAR 26 OF analit_info PROMPT '\-'
DEFINE BAR 27 OF analit_info PROMPT 'Отчет по открытым и закрытым счетам клиентов на дату' ;
  MESSAGE 'Отчет по открытым и закрытым счетам клиентов, дата открытия которых менее выбранной даты'
DEFINE BAR 28 OF analit_info PROMPT '\-'
DEFINE BAR 29 OF analit_info PROMPT 'Отчет по открытым и закрытым счетам клиентов за дату' ;
  MESSAGE 'Отчет по открытым и закрытым счетам клиентов, дата открытия которых равна выбранной дате'
DEFINE BAR 30 OF analit_info PROMPT '\-'
DEFINE BAR 31 OF analit_info PROMPT 'Отчет по открытым и закрытым счетам клиентов за период' ;
  MESSAGE 'Отчет по открытым и закрытым счетам клиентов, дата открытия которых попадает в выбранный период'
DEFINE BAR 32 OF analit_info PROMPT '\-'
DEFINE BAR 33 OF analit_info PROMPT 'Отчет по количеству клиентов, счетов, карт на дату' ;
  MESSAGE 'Отчет по количеству клиентов, счетов, карт, дата открытия которых менее выбранной даты'
DEFINE BAR 34 OF analit_info PROMPT '\-'
DEFINE BAR 35 OF analit_info PROMPT 'Отчет по количеству клиентов, счетов, карт за период' ;
  MESSAGE 'Отчет по количеству клиентов, счетов, карт за период, дата открытия которых попадает в выбранный период'
DEFINE BAR 36 OF analit_info PROMPT '\-'
DEFINE BAR 37 OF analit_info PROMPT 'Счета с остатком более указанной суммы и меньшей датой открытия' ;
  MESSAGE 'Отчет по счетам с остатком более указанной суммы и датой открытия меньшей выбранной'
DEFINE BAR 38 OF analit_info PROMPT '\-'
DEFINE BAR 39 OF analit_info PROMPT 'Счета с остатком менее указанной суммы и меньшей датой открытия' ;
  MESSAGE 'Отчет по счетам с остатком менее указанной суммы и датой открытия меньшей выбранной'
DEFINE BAR 40 OF analit_info PROMPT '\-'
DEFINE BAR 41 OF analit_info PROMPT 'Книга регистрации лицевых счетов, с датой открытия меньшей введенной даты' ;
  MESSAGE 'Отчет по открытым и закрытым счетам, с датой открытия меньшей введенной даты'
DEFINE BAR 42 OF analit_info PROMPT '\-'
DEFINE BAR 43 OF analit_info PROMPT 'Отчет по карточным счетам по заданному фильтру остатков за дату' ;
  MESSAGE 'Отчет по карточным счетам по заданному фильтру остатков за дату'
DEFINE BAR 44 OF analit_info PROMPT '\-'
DEFINE BAR 45 OF analit_info PROMPT 'Отчет по карточным счетам по фильтру остатков форма 345 за дату' ;
  MESSAGE 'Отчет по карточным счетам по фильтру остатков форма 345 за дату'

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
DEFINE BAR 1 OF otchet_zb PROMPT 'Данные об остатках на счетах - форма 1416 У' ;
  MESSAGE 'Расчет и формирование данных о ежедневных остатках на карточных счетах и подлежащих страхованию - форма 1416 У'
DEFINE BAR 2 OF otchet_zb PROMPT '\-'
DEFINE BAR 3 OF otchet_zb PROMPT 'Реестр обязательств банка - форма 1417 У' ;
  MESSAGE 'Расчет и формирование реестра обязательств банка перед вкладчиками - форма 1417 У'
DEFINE BAR 4 OF otchet_zb PROMPT '\-'
DEFINE BAR 5 OF otchet_zb PROMPT 'Расчет ликвидности банка - форма 1991 У' ;
  MESSAGE 'Формирование детализированной ведомости остатков банка - форма 1991 У'

ON SELECTION BAR 1 OF otchet_zb DO start IN Otchet_1416_U
ON SELECTION BAR 3 OF otchet_zb DO start IN Otchet_1417_U
ON SELECTION BAR 5 OF otchet_zb DO start IN Otchet_1991_U


**************************************************************** DEFINE POPUP servis_doc *****************************************************************

DEFINE POPUP servis_doc  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF servis_doc PROMPT 'Справочник анкетных данных клиента' ;
  MESSAGE 'Просмотр и редактирование справочника анкетных данных клиента'
DEFINE BAR 2  OF servis_doc PROMPT '\-'
DEFINE BAR 3  OF servis_doc PROMPT 'Справочник карточных счетов клиента' ;
  MESSAGE 'Просмотр и редактирование карточных счетов клиента'
DEFINE BAR 4  OF servis_doc PROMPT '\-'
DEFINE BAR 5  OF servis_doc PROMPT 'Справочник по выданным картам' ;
  MESSAGE 'Просмотр справочника по выданным картам'
DEFINE BAR 6  OF servis_doc PROMPT '\-'
DEFINE BAR 7  OF servis_doc PROMPT 'Справочник кодов операций' ;
  MESSAGE 'Просмотр справочника кодов операций'
DEFINE BAR 8  OF servis_doc PROMPT '\-'
DEFINE BAR 9  OF servis_doc PROMPT 'Справочник переменных' ;
  MESSAGE 'Просмотр и редактирование справочника переменных'
DEFINE BAR 10 OF servis_doc PROMPT '\-'
DEFINE BAR 11 OF servis_doc PROMPT 'Справочник курсов валют за дату' ;
  MESSAGE 'Просмотр и редактирование справочника курсов валют за выбранную дату'
DEFINE BAR 12 OF servis_doc PROMPT '\-'
DEFINE BAR 13 OF servis_doc PROMPT 'Справочник транзитных счетов' ;
  MESSAGE 'Просмотр и редактирование справочника транзитных счетов'
DEFINE BAR 14 OF servis_doc PROMPT '\-'
DEFINE BAR 15 OF servis_doc PROMPT 'Справочник тарифных ставок' ;
  MESSAGE 'Просмотр и редактирование справочника тарифных ставок по выдаче и обслуживанию карт Visa'
DEFINE BAR 16 OF servis_doc PROMPT '\-'
DEFINE BAR 17 OF servis_doc PROMPT 'Паспортные данные доверенного лица клиента' ;

DO CASE
  CASE pr_rabota_sql = .T.  && Работа происходит с таблицами SQL Server

    ON SELECTION BAR 1   OF servis_doc DO spr_client IN Servis_sql
    ON SELECTION BAR 3   OF servis_doc DO read_sks IN Servis_sql
    ON SELECTION BAR 5   OF servis_doc DO spr_karta IN Servis_sql
    ON SELECTION BAR 7   OF servis_doc DO spr_kod_oper  IN Servis_sql
    ON SELECTION BAR 9   OF servis_doc DO spravliz  IN Servis_sql
    ON SELECTION BAR 11 OF servis_doc DO red_spr_val IN Servis_sql
    ON SELECTION BAR 13 OF servis_doc DO read_tran IN Servis_sql
    ON SELECTION BAR 15 OF servis_doc DO read_tarif IN Servis_sql
    ON SELECTION BAR 17 OF servis_doc DO read_pass_dover IN Servis

  CASE pr_rabota_sql = .F.  && Работа происходит с локальными таблицами

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


* ================== Формирование горизонтального меню для пункта главного меню "Выдача карт и управление карточным счетом" ====================== *


*********************************************************** DEFINE MENU client **************************************************************************

DEFINE MENU client COLOR SCHEME 2 ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B'

DEFINE PAD cardclient OF client PROMPT 'Заявление' COLOR SCHEME 2
DEFINE PAD popolclient OF client PROMPT 'Пополнение' COLOR SCHEME 2
DEFINE PAD spisclient OF client PROMPT 'Коррекции' COLOR SCHEME 2
DEFINE PAD popol_ckc OF client PROMPT 'Управление' COLOR SCHEME 2
DEFINE PAD zakrclient OF client PROMPT 'Закрытие' COLOR SCHEME 2
DEFINE PAD analitika OF client PROMPT 'Аналитика' COLOR SCHEME 2
DEFINE PAD servis_cln OF client PROMPT 'Справочники' COLOR SCHEME 2
DEFINE PAD quit OF client PROMPT 'Выход' COLOR SCHEME 2 ;
  MESSAGE 'Выход в главное меню'

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
DEFINE BAR 1  OF cardclient PROMPT 'Основная карта - заполнение данных на выдачу' ;
  MESSAGE 'Данный пункт меню позволяет внести данные по клиенту для заполнения заявления на выдачу основной ПК'
DEFINE BAR 2  OF cardclient PROMPT '\-'
DEFINE BAR 3  OF cardclient PROMPT 'Основная карта - формирование данных для процессинга' ;
  MESSAGE 'Данный пункт меню позволяет формировать данные по новым клиентам для отправки в процессинговый центр банка'
DEFINE BAR 4  OF cardclient PROMPT '\-'
DEFINE BAR 5  OF cardclient PROMPT 'Основная карта - сводная заявка на выпуск карт' ;
  MESSAGE 'Данный пункт меню позволяет формировать сводную заявку на выпуск карт по новым клиентам для отправки в процессинговый центр банка'
DEFINE BAR 6  OF cardclient PROMPT '\-'
DEFINE BAR 7  OF cardclient PROMPT 'Дополнительная карта - заполнение данных на выдачу' ;
  MESSAGE 'Данный пункт меню позволяет внести данные по клиенту для заполнения заявления на выдачу дополнительной ПК'
DEFINE BAR 8  OF cardclient PROMPT '\-'
DEFINE BAR 9  OF cardclient PROMPT 'Дополнительная карта - формирование данных для процессинга' ;
  MESSAGE 'Данный пункт меню позволяет формировать данные по новым клиентам для отправки в процессинговый центр банка'
DEFINE BAR 10  OF cardclient PROMPT '\-'
DEFINE BAR 11  OF cardclient PROMPT 'Дополнительная карта - сводная заявка на выпуск карт' ;
  MESSAGE 'Данный пункт меню позволяет формировать сводную заявку на выпуск карт по новым клиентам для отправки в процессинговый центр банка'
DEFINE BAR 12  OF cardclient PROMPT '\-'
DEFINE BAR 13  OF cardclient PROMPT 'Формирование измененных данных для ранее введенных клиентов' ;
  MESSAGE 'Данный пункт меню позволяет отправить измененные данные по ранее введенным клиентам в процессинговый центр банка'
DEFINE BAR 14  OF cardclient PROMPT '\-'
DEFINE BAR 15  OF cardclient PROMPT 'Импорт внешних данных для открытия карт списком' ;
  MESSAGE 'Данный пункт меню позволяет импортировать заранее подготовленные данные для открытия карт списком'
DEFINE BAR 16  OF cardclient PROMPT '\-'
DEFINE BAR 17  OF cardclient PROMPT 'Создание архивного файла для отправки в УПК' ;
  MESSAGE 'Данный пункт меню позволяет создать и отправить архивный файл в УПК'

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
DEFINE BAR 1  OF popolclient PROMPT 'Карточного счета наличными - кассовый ордер' ;
  MESSAGE 'Формирование бухгалтерских документов на пополнение пластиковой карты наличными - составление приходного кассового ордера'
DEFINE BAR 2  OF popolclient PROMPT '\-'
DEFINE BAR 3  OF popolclient PROMPT 'Карточного счета безналичными - мемоордер' ;
  MESSAGE 'Формирование бухгалтерских документов на пополнение пластиковой карты безналичными - составление мемориального ордера'
DEFINE BAR 4  OF popolclient PROMPT '\-'
DEFINE BAR 5  OF popolclient PROMPT 'Карточного счета безналичными - платежное поручение' ;
  MESSAGE 'Формирование бухгалтерских документов на пополнение пластиковой карты безналичными - составление платежного поручения'
DEFINE BAR 6 OF popolclient PROMPT '\-'
DEFINE BAR 7 OF popolclient PROMPT 'Начисление средств по зарплатному проекту - мемоордер' ;
  MESSAGE 'Формирование бухгалтерских документов для начисления средств по зарплатному проекту безналичными - составление мемориального ордера'
DEFINE BAR 8 OF popolclient PROMPT '\-'
DEFINE BAR 9 OF popolclient PROMPT 'Импорт данных по зарплатному проекту - мемоордер - формат DBF' ;
  MESSAGE 'Формирование бухгалтерских документов для начисления средств по зарплатному проекту путем импорта данных в электронном виде'
DEFINE BAR 10 OF popolclient PROMPT '\-'
DEFINE BAR 11 OF popolclient PROMPT 'Импорт данных по зарплатному проекту - мемоордер - формат DCL' ;
  MESSAGE 'Формирование бухгалтерских документов для начисления средств по зарплатному проекту путем импорта данных в электронном виде'

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
DEFINE BAR 1  OF spisclient PROMPT 'Списание со счета различных комиссионных сумм' ;
  MESSAGE 'Данный пункт меню позволяет составить мемоордер по списанию со счета различных комиссионных сумм'
DEFINE BAR 2  OF spisclient PROMPT '\-'
DEFINE BAR 3  OF spisclient PROMPT 'Списание со счета различных сумм отчислений и плат' ;
  MESSAGE 'Данный пункт меню позволяет составить мемоордер по списанию со счета различных сумм отчислений и плат'
DEFINE BAR 4  OF spisclient PROMPT '\-'
DEFINE BAR 5  OF spisclient PROMPT 'Списание со счета различных штрафных сумм' ;
  MESSAGE 'Данный пункт меню позволяет составить мемоордер по списанию со счета различных штрафных сумм'
DEFINE BAR 6  OF spisclient PROMPT '\-'
DEFINE BAR 7  OF spisclient PROMPT 'Начисление на счет процентов за обслуживание' ;
  MESSAGE 'Данный пункт меню позволяет составить мемоордер по начисленым на счет процентов за обслуживание'
DEFINE BAR 8 OF spisclient PROMPT '\-'
DEFINE BAR 9 OF spisclient PROMPT 'Импорт данных на списание с карточных счетов списком' ;
  MESSAGE 'Формирование бухгалтерских документов на списание средств с карточных счетов путем импорта данных списком в электронном виде'

ON SELECTION BAR 1  OF spisclient DO start IN Order_memo
ON SELECTION BAR 3  OF spisclient DO start IN Order_memo
ON SELECTION BAR 5  OF spisclient DO start IN Order_memo
ON SELECTION BAR 7  OF spisclient DO start IN Order_memo
ON SELECTION BAR 9 OF spisclient DO import_deb  IN Dokum_memo


************************************************************* DEFINE POPUP popol_ckc   *******************************************************************

DEFINE POPUP popol_ckc  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF popol_ckc PROMPT 'Формирование файла начислений и списаний для УПК' ;
  MESSAGE 'Данный пункт меню позволяет формировать специальный файл начислений и списаний для управления картсчетом для отправки его в УПК'
DEFINE BAR 2  OF popol_ckc PROMPT '\-'
DEFINE BAR 3  OF popol_ckc PROMPT 'Создание архивного файла для отправки в УПК' ;
  MESSAGE 'Данный пункт меню позволяет создавать архивный файл по утвержденному шаблону для отправки его в УПК'
DEFINE BAR 4 OF popol_ckc PROMPT '\-'
DEFINE BAR 5 OF popol_ckc PROMPT 'Экспорт сумм пополнения и списания в опердень' ;
  MESSAGE 'Данный пункт меню позволяет создавать файл с суммами пополнения и списания для экспорта в операционный день банка'

ON SELECTION BAR 1  OF popol_ckc DO start IN Popol_ckc
ON SELECTION BAR 3  OF popol_ckc DO create_zip IN Popol_ckc
ON SELECTION BAR 5  OF popol_ckc DO export_odb IN Popol_doc


*************************************************************** DEFINE POPUP zakrclient *******************************************************************

DEFINE POPUP zakrclient  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF zakrclient PROMPT 'Карточного счета наличными - кассовый ордер' ;
  MESSAGE 'Формирование бухгалтерских документов на пополнение пластиковой карты наличными - составление расходного кассового ордера'
DEFINE BAR 2  OF zakrclient PROMPT '\-'
DEFINE BAR 3  OF zakrclient PROMPT 'Карточного счета безналичными - мемоордер' ;
  MESSAGE 'Формирование бухгалтерских документов на пополнение пластиковой карты безналичными - составление мемориального ордера'
DEFINE BAR 4  OF zakrclient PROMPT '\-'
DEFINE BAR 5  OF zakrclient PROMPT 'Карточного счета безналичными - платежное поручение' ;
  MESSAGE 'Формирование бухгалтерских документов на пополнение пластиковой карты безналичными - составление платежного поручения'

ON SELECTION BAR 1 OF zakrclient DO start  IN Order_kassa
ON SELECTION BAR 3 OF zakrclient DO start  IN Order_memo
ON SELECTION BAR 5 OF zakrclient DO start IN Plateg


*************************************************************** DEFINE POPUP analitika *******************************************************************

DEFINE POPUP analitika  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF analitika PROMPT 'Уведомление о пополнении средств на счетах' ;
  MESSAGE 'Данный пункт меню позволяет формировать сводное уведомление о пополнении средств на счетах для отправки в УПК'
DEFINE BAR 2  OF analitika PROMPT '\-'
DEFINE BAR 3  OF analitika PROMPT 'Уведомление о списании средств на счетах' ;
  MESSAGE 'Данный пункт меню позволяет формировать сводное уведомление о списании средств на счетах для отправки в УПК'
DEFINE BAR 4 OF analitika PROMPT '\-'
DEFINE BAR 5 OF analitika PROMPT 'Выборка операций с признаком НДС для счетов-фактур' ;
  MESSAGE 'Данный пункт меню позволяет произвести выборку проведенных операций с признаком НДС для счетов-фактур'
DEFINE BAR 6 OF analitika PROMPT '\-'
DEFINE BAR 7 OF analitika PROMPT 'Список карт, которые перевыпускаются в текущем месяце' ;
  MESSAGE 'Работа со счетами, с которых необходимо списать плату за годовое обслуживание на текущую дату'
DEFINE BAR 8 OF analitika PROMPT '\-'
DEFINE BAR 9 OF analitika PROMPT 'Сумма оплаты за перевыпуск карт в текущем месяце' ;
  MESSAGE 'Счета, с которых необходимо списать плату за перевыпуск карт в текущем месяце'

ON SELECTION BAR 1  OF analitika DO start_popol IN Popol_doc
ON SELECTION BAR 3  OF analitika DO start_spis IN Popol_doc
ON SELECTION BAR 5  OF analitika DO faktura IN Dokum_memo
ON SELECTION BAR 7  OF analitika DO start IN Obsluga_god
ON SELECTION BAR 9  OF analitika DO start IN Obsl_god_proekt


*************************************************************** DEFINE POPUP servis_cln *******************************************************************

DEFINE POPUP servis_cln  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF servis_cln PROMPT 'Справочник анкетных данных клиента' ;
  MESSAGE 'Просмотр справочника анкетных данных клиента'
DEFINE BAR 2  OF servis_cln PROMPT '\-'
DEFINE BAR 3  OF servis_cln PROMPT 'Справочник карточных счетов клиента' ;
  MESSAGE 'Редактирование карточных счетов клиента'
DEFINE BAR 4  OF servis_cln PROMPT '\-'
DEFINE BAR 5  OF servis_cln PROMPT 'Справочник по выданным картам' ;
  MESSAGE 'Просмотр справочника по картам'
DEFINE BAR 6  OF servis_cln PROMPT '\-'
DEFINE BAR 7  OF servis_cln PROMPT 'Справочник кодов операций' ;
  MESSAGE 'Просмотр справочника кодов операций'
DEFINE BAR 8  OF servis_cln PROMPT '\-'
DEFINE BAR 9  OF servis_cln PROMPT 'Справочник переменных' ;
  MESSAGE 'Просмотр и редактирование справочника переменных'
DEFINE BAR 10 OF servis_cln PROMPT '\-'
DEFINE BAR 11 OF servis_cln PROMPT 'Справочник курсов валют' ;
  MESSAGE 'Просмотр и редактирование справочника курсов валют за выбранную дату'
DEFINE BAR 12 OF servis_cln PROMPT '\-'
DEFINE BAR 13 OF servis_cln PROMPT 'Справочник транзитных счетов' ;
  MESSAGE 'Редактирование справочника транзитных счетов'
DEFINE BAR 14 OF servis_cln PROMPT '\-'
DEFINE BAR 15 OF servis_cln PROMPT 'Печать дополнительного соглашения - списание овердрафта' ;
  MESSAGE 'Печать дополнительного соглашения - списание овердрафта'
DEFINE BAR 16 OF servis_cln PROMPT '\-'
DEFINE BAR 17 OF servis_cln PROMPT 'Печать дополнительного соглашения - изменение счета' ;
  MESSAGE 'Печать дополнительного соглашения - изменение счета'
DEFINE BAR 18 OF servis_cln PROMPT '\-'
DEFINE BAR 19 OF servis_cln PROMPT 'Печать заявления на получение карты' ;
  MESSAGE 'Печать заявления на получение карты'
DEFINE BAR 20 OF servis_cln PROMPT '\-'
DEFINE BAR 21 OF servis_cln PROMPT 'Справочник тарифных ставок' ;
  MESSAGE 'Редактирование справочника тарифных ставок по выдаче и обслуживанию карт Visa'
DEFINE BAR 22 OF servis_cln PROMPT '\-'
DEFINE BAR 23 OF servis_cln PROMPT 'Паспортные данные доверенного лица клиента' ;

DO CASE
  CASE pr_rabota_sql = .T.  && Работа происходит с таблицами SQL Server

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

  CASE pr_rabota_sql = .F.  && Работа происходит с локальными таблицами

    DEFINE BAR 24 OF servis_cln PROMPT '\-'
    DEFINE BAR 25 OF servis_cln PROMPT 'Добавление ранее заведенных клиентов в таблицу заявлений' ;
      MESSAGE 'Добавление ранее заведенных клиентов в таблицу заявлений на получение пластиковой карты'

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


* ===================== Формирование горизонтального меню для пункта главного меню "Создание и формирование отчетов в УПК" ====================== *


************************************************************* DEFINE MENU otchet_upk ********************************************************************

DEFINE MENU otchet_upk COLOR SCHEME 2 ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B'

DEFINE PAD otchet_upk OF otchet_upk PROMPT 'Отчеты в УПК' COLOR SCHEME 2
DEFINE PAD servis_upk OF otchet_upk PROMPT 'Справочники' COLOR SCHEME 2
DEFINE PAD quit OF otchet_upk PROMPT 'Выход' COLOR SCHEME 2 ;
  MESSAGE 'Выход в главное меню'

ON PAD otchet_upk OF otchet_upk ACTIVATE POPUP otchet_upk
ON PAD servis_upk  OF otchet_upk ACTIVATE POPUP servis_upk
ON SELECTION PAD quit OF otchet_upk DEACTIVATE MENU otchet_upk


***************************************************************** DEFINE POPUP otchet_upk   **************************************************************

DEFINE POPUP otchet_upk  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF otchet_upk PROMPT 'Заявление и уведомление о сдаче карты' ;
  MESSAGE 'Данный пункт меню позволяет сформировать заявление и уведомление о сдаче карты'
DEFINE BAR 2  OF otchet_upk PROMPT '\-'
DEFINE BAR 3  OF otchet_upk PROMPT 'Заявление об отказе пластиковой карты' ;
  MESSAGE 'Данный пункт меню позволяет сформировать заявление об отказе карты'
DEFINE BAR 4  OF otchet_upk PROMPT '\-'
DEFINE BAR 5  OF otchet_upk PROMPT 'Заявление об утрате  пластиковой карты' ;
  MESSAGE 'Данный пункт меню позволяет сформировать заявление об утрате карты'
DEFINE BAR 6  OF otchet_upk PROMPT '\-'
DEFINE BAR 7  OF otchet_upk PROMPT 'Расписка клиента в получении новых карт' ;
  MESSAGE 'Данный пункт меню позволяет сформировать расписку в получении карты'
DEFINE BAR 8  OF otchet_upk PROMPT '\-'
DEFINE BAR 9  OF otchet_upk PROMPT 'Уведомление о выдаче карт списком' ;
  MESSAGE 'Данный пункт меню позволяет сформировать уведомление о выдаче карт списком'
DEFINE BAR 10 OF otchet_upk PROMPT '\-'
DEFINE BAR 11 OF otchet_upk PROMPT 'Уведомление о сдаче карт списком' ;
  MESSAGE 'Данный пункт меню позволяет сформировать уведомление о сдаче карт списком'

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
DEFINE BAR 1  OF servis_upk PROMPT 'Справочник анкетных данных клиента' ;
  MESSAGE 'Просмотр справочника анкетных данных клиента'
DEFINE BAR 2  OF servis_upk PROMPT '\-'
DEFINE BAR 3  OF servis_upk PROMPT 'Справочник карточных счетов клиента' ;
  MESSAGE 'Редактирование карточных счетов клиента'
DEFINE BAR 4  OF servis_upk PROMPT '\-'
DEFINE BAR 5  OF servis_upk PROMPT 'Справочник по выданным картам' ;
  MESSAGE 'Просмотр справочника по картам'
DEFINE BAR 6  OF servis_upk PROMPT '\-'
DEFINE BAR 7  OF servis_upk PROMPT 'Справочник кодов операций' ;
  MESSAGE 'Просмотр справочника кодов операций'
DEFINE BAR 8  OF servis_upk PROMPT '\-'
DEFINE BAR 9  OF servis_upk PROMPT 'Справочник переменных' ;
  MESSAGE 'Просмотр и редактирование справочника переменных'
DEFINE BAR 10 OF servis_upk PROMPT '\-'
DEFINE BAR 11 OF servis_upk PROMPT 'Справочник курсов валют' ;
  MESSAGE 'Просмотр и редактирование справочника курсов валют за выбранную дату'
DEFINE BAR 12 OF servis_upk PROMPT '\-'
DEFINE BAR 13 OF servis_upk PROMPT 'Справочник транзитных счетов' ;
  MESSAGE 'Редактирование справочника транзитных счетов'
DEFINE BAR 14 OF servis_upk PROMPT '\-'
DEFINE BAR 15 OF servis_upk PROMPT 'Справочник тарифных ставок' ;
  MESSAGE 'Редактирование справочника тарифных ставок по выдаче и обслуживанию карт Visa'

DO CASE
  CASE pr_rabota_sql = .T.  && Работа происходит с таблицами SQL Server

    ON SELECTION BAR 1   OF servis_upk DO spr_client IN Servis_sql
    ON SELECTION BAR 3   OF servis_upk DO read_sks IN Servis_sql
    ON SELECTION BAR 5   OF servis_upk DO spr_karta IN Servis_sql
    ON SELECTION BAR 7   OF servis_upk DO spr_kod_oper  IN Servis_sql
    ON SELECTION BAR 9   OF servis_upk DO spravliz  IN Servis_sql
    ON SELECTION BAR 11 OF servis_upk DO red_spr_val IN Servis_sql
    ON SELECTION BAR 13 OF servis_upk DO read_tran IN Servis_sql
    ON SELECTION BAR 15 OF servis_upk DO read_tarif IN Servis_sql

  CASE pr_rabota_sql = .F.  && Работа происходит с локальными таблицами

    ON SELECTION BAR 1   OF servis_upk DO spr_client IN Servis
    ON SELECTION BAR 3   OF servis_upk DO read_sks IN Servis
    ON SELECTION BAR 5   OF servis_upk DO spr_karta IN Servis
    ON SELECTION BAR 7   OF servis_upk DO spr_kod_oper  IN Servis
    ON SELECTION BAR 9   OF servis_upk DO spravliz  IN Servis
    ON SELECTION BAR 11 OF servis_upk DO red_spr_val IN Servis
    ON SELECTION BAR 13 OF servis_upk DO read_tran IN Servis
    ON SELECTION BAR 15 OF servis_upk DO read_tarif IN Servis

ENDCASE


* ===================== Формирование горизонтального меню для пункта главного меню "Начисление процентов по карточному счету" ====================== *


************************************************************* DEFINE MENU proz_ckc *********************************************************************

DEFINE MENU proz_ckc COLOR SCHEME 2 ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B'

DEFINE PAD raschet_proz_42301 OF proz_ckc PROMPT 'Расчет по карточным счетам' COLOR SCHEME 2
DEFINE PAD raschet_proz_42308 OF proz_ckc PROMPT 'Расчет по страховым депозитам' COLOR SCHEME 2
DEFINE PAD blank_proz OF proz_ckc PROMPT 'Документация' COLOR SCHEME 2
DEFINE PAD servis_proz OF proz_ckc PROMPT 'Справочники' COLOR SCHEME 2
DEFINE PAD quit OF proz_ckc PROMPT 'Выход' COLOR SCHEME 2 ;
  MESSAGE 'Выход в главное меню'

ON PAD raschet_proz_42301  OF proz_ckc ACTIVATE POPUP raschet_proz_42301
ON PAD raschet_proz_42308  OF proz_ckc ACTIVATE POPUP raschet_proz_42308
ON PAD blank_proz    OF proz_ckc ACTIVATE POPUP blank_proz
ON PAD servis_proz  OF proz_ckc ACTIVATE POPUP servis_proz
ON SELECTION PAD quit OF proz_ckc DEACTIVATE MENU proz_ckc


********************************************************* DEFINE POPUP raschet_proz_42301 ****************************************************************

DEFINE POPUP raschet_proz_42301  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF raschet_proz_42301 PROMPT 'Коррекция параметров картсчета' ;
  MESSAGE 'Ввод и редактирование номеров счетов и процентной ставки, участвующих при расчете процентов'
DEFINE BAR 2  OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 3  OF raschet_proz_42301 PROMPT 'Расчет процентов полным списком' ;
  MESSAGE 'Ввод дат начала и окончания периода, ввод курса валюты для расчета и расчет процентов'
DEFINE BAR 4  OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 5  OF raschet_proz_42301 PROMPT 'Расчет процентов выборочный' ;
  MESSAGE 'Ввод дат начала и окончания периода, ввод курса валюты, выбор владельца карты и расчет процентов'
DEFINE BAR 6  OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 7  OF raschet_proz_42301 PROMPT 'Редактирование данных расчета' ;
  MESSAGE 'Выбор дат начала и окончания периода расчета из списка и редактирование рассчитанных данных'
DEFINE BAR 8  OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 9  OF raschet_proz_42301 PROMPT 'Удаление данных по рассчитанным процентам' ;
  MESSAGE 'Удаление данных по ранее рассчитанным процентам как по отдельному клиенту, так и весь список за выбранный диапазон дат'
DEFINE BAR 10 OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 11 OF raschet_proz_42301 PROMPT 'Просмотр рассчитанных процентов' ;
  MESSAGE 'Выбор дат начала и окончания периода расчета из списка и просмотр рассчитанных процентов'
DEFINE BAR 12  OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 13  OF raschet_proz_42301 PROMPT 'Печать рассчитанных процентов' ;
  MESSAGE 'Выбор дат начала и окончания периода расчета из списка и просмотр рассчитанных процентов'
DEFINE BAR 14  OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 15  OF raschet_proz_42301 PROMPT 'Формирование файла Pb по начислению процентов' ;
  MESSAGE 'Формирование файла Pb по начислению процентов с кодом 516'
DEFINE BAR 16 OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 17 OF raschet_proz_42301 PROMPT 'Создание архивного файла для отправки в УПК' ;
  MESSAGE 'Данный пункт меню позволяет создавать архивный файл по утвержденному шаблону для отправки в УПК'
DEFINE BAR 18 OF raschet_proz_42301 PROMPT '\-'
DEFINE BAR 19 OF raschet_proz_42301 PROMPT 'Экспорт начисленных процентов в операционный день банка' ;
  MESSAGE 'Данный пункт меню позволяет создавать файл с начисленными процентами для эксорта в операционный день банка'

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
DEFINE BAR 1  OF raschet_proz_42308 PROMPT 'Коррекция номеров счетов' ;
  MESSAGE 'Ввод и редактирование номеров счетов, участвующих при расчете процентов'
DEFINE BAR 2  OF raschet_proz_42308 PROMPT '\-'
DEFINE BAR 3  OF raschet_proz_42308 PROMPT 'Расчет процентов полным списком' ;
  MESSAGE 'Ввод дат начала и окончания периода, ввод курса валюты для расчета и расчет процентов'
DEFINE BAR 4  OF raschet_proz_42308 PROMPT '\-'
DEFINE BAR 5  OF raschet_proz_42308 PROMPT 'Расчет процентов выборочный' ;
  MESSAGE 'Ввод дат начала и окончания периода, ввод курса валюты, выбор владельца карты и расчет процентов'
DEFINE BAR 6  OF raschet_proz_42308 PROMPT '\-'
DEFINE BAR 7  OF raschet_proz_42308 PROMPT 'Редактирование данных расчета' ;
  MESSAGE 'Выбор дат начала и окончания периода расчета из списка и редактирование рассчитанных данных'
DEFINE BAR 8  OF raschet_proz_42308 PROMPT '\-'
DEFINE BAR 9  OF raschet_proz_42308 PROMPT 'Удаление данных по рассчитанным процентам' ;
  MESSAGE 'Удаление данных по ранее рассчитанным процентам как по отдельному клиенту, так и весь список за выбранный диапазон дат'
DEFINE BAR 10 OF raschet_proz_42308 PROMPT '\-'
DEFINE BAR 11 OF raschet_proz_42308 PROMPT 'Просмотр рассчитанных процентов' ;
  MESSAGE 'Выбор дат начала и окончания периода расчета из списка и просмотр рассчитанных процентов'
DEFINE BAR 12  OF raschet_proz_42308 PROMPT '\-'
DEFINE BAR 13  OF raschet_proz_42308 PROMPT 'Печать рассчитанных процентов' ;
  MESSAGE 'Выбор дат начала и окончания периода расчета из списка и печать рассчитанных процентов'

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
DEFINE BAR 1  OF blank_proz PROMPT 'Заявление на открытие счета 47411' ;
  MESSAGE 'Формирование бланка заявления на открытие счета 47411'
ON SELECTION BAR 1  OF blank_proz DO start IN Zajava_proz


***************************************************************** DEFINE POPUP servis_proz   **************************************************************

DEFINE POPUP servis_proz  ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF servis_proz PROMPT 'Справочник анкетных данных клиента' ;
  MESSAGE 'Просмотр справочника анкетных данных клиента'
DEFINE BAR 2  OF servis_proz PROMPT '\-'
DEFINE BAR 3  OF servis_proz PROMPT 'Справочник карточных счетов клиента' ;
  MESSAGE 'Редактирование карточных счетов клиента'
DEFINE BAR 4  OF servis_proz PROMPT '\-'
DEFINE BAR 5  OF servis_proz PROMPT 'Справочник по выданным картам' ;
  MESSAGE 'Просмотр справочника по картам'
DEFINE BAR 6  OF servis_proz PROMPT '\-'
DEFINE BAR 7  OF servis_proz PROMPT 'Справочник кодов операций' ;
  MESSAGE 'Просмотр справочника кодов операций'
DEFINE BAR 8  OF servis_proz PROMPT '\-'
DEFINE BAR 9  OF servis_proz PROMPT 'Справочник переменных' ;
  MESSAGE 'Просмотр и редактирование справочника переменных'
DEFINE BAR 10 OF servis_proz PROMPT '\-'
DEFINE BAR 11 OF servis_proz PROMPT 'Справочник курсов валют' ;
  MESSAGE 'Просмотр и редактирование справочника курсов валют за выбранную дату'
DEFINE BAR 12 OF servis_proz PROMPT '\-'
DEFINE BAR 13 OF servis_proz PROMPT 'Справочник транзитных счетов' ;
  MESSAGE 'Редактирование справочника транзитных счетов'
DEFINE BAR 14 OF servis_proz PROMPT '\-'
DEFINE BAR 15 OF servis_proz PROMPT 'Выверка договорных параметров по клиентам' ;
  MESSAGE 'Сканирование справочника карточных счетов и автоматическое изменение в таблицах процентов договорных параметров по клиентам'
DEFINE BAR 16 OF servis_proz PROMPT '\-'
DEFINE BAR 17 OF servis_proz PROMPT 'Справочник тарифных ставок' ;
  MESSAGE 'Редактирование справочника тарифных ставок по выдаче и обслуживанию карт Visa'

DO CASE
  CASE pr_rabota_sql = .T.  && Работа происходит с таблицами SQL Server

    ON SELECTION BAR 1   OF servis_proz DO spr_client IN Servis_sql
    ON SELECTION BAR 3   OF servis_proz DO read_sks IN Servis_sql
    ON SELECTION BAR 5   OF servis_proz DO spr_karta IN Servis_sql
    ON SELECTION BAR 7   OF servis_proz DO spr_kod_oper  IN Servis_sql
    ON SELECTION BAR 9   OF servis_proz DO spravliz  IN Servis_sql
    ON SELECTION BAR 11 OF servis_proz DO red_spr_val IN Servis_sql
    ON SELECTION BAR 13 OF servis_proz DO read_tran IN Servis_sql
    ON SELECTION BAR 15 OF servis_proz DO scan_data_dog IN Raschet_proz_42301
    ON SELECTION BAR 17 OF servis_proz DO read_tarif IN Servis_sql

  CASE pr_rabota_sql = .F.  && Работа происходит с локальными таблицами

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


* ========================== Формирование горизонтального меню для пункта главного меню "Сервисные и системные процедуры" ====================== *


************************************************************* DEFINE MENU servis_glav *******************************************************************

DEFINE MENU servis_glav COLOR SCHEME 2 ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B'

DEFINE PAD sprav_serv OF servis_glav PROMPT 'Справочники' COLOR SCHEME 2
DEFINE PAD servis_bik OF servis_glav PROMPT 'Справочник БИК' COLOR SCHEME 2 SKIP FOR wes < 5
DEFINE PAD import_kurs OF servis_glav PROMPT 'Курсы валют' COLOR SCHEME 2 SKIP FOR wes < 5
DEFINE PAD read_data OF servis_glav PROMPT 'Правка данных' COLOR SCHEME 2 SKIP FOR wes < 5
DEFINE PAD read_card_nls OF servis_glav PROMPT 'Работа со счетами' COLOR SCHEME 2 SKIP FOR wes < 5
DEFINE PAD servis_serv OF servis_glav PROMPT 'Обслуживание' COLOR SCHEME 2 SKIP FOR wes < 5
DEFINE PAD quit OF servis_glav PROMPT 'Выход' COLOR SCHEME 2 ;
  MESSAGE 'Выход в главное меню'

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
DEFINE BAR 1  OF sprav_serv PROMPT 'Справочник кодов операций' ;
  MESSAGE 'Просмотр справочника кодов операций'
DEFINE BAR 2  OF sprav_serv PROMPT '\-'
DEFINE BAR 3  OF sprav_serv PROMPT 'Справочник переменных' ;
  MESSAGE 'Просмотр и редактирование справочника переменных'
DEFINE BAR 4  OF sprav_serv PROMPT '\-'
DEFINE BAR 5  OF sprav_serv PROMPT 'Справочник реквизитов филиала' ;
  MESSAGE 'Просмотр и редактирование справочника реквизитов филиала'
DEFINE BAR 6  OF sprav_serv PROMPT '\-'
DEFINE BAR 7  OF sprav_serv PROMPT 'Справочник курсов валют' ;
  MESSAGE 'Просмотр и редактирование справочника курсов валют за выбранную дату'
DEFINE BAR 8  OF sprav_serv PROMPT '\-'
DEFINE BAR 9  OF sprav_serv PROMPT 'Справочник транзитных счетов' ;
  MESSAGE 'Редактирование справочника транзитных счетов'
DEFINE BAR 10  OF sprav_serv PROMPT '\-'
DEFINE BAR 11  OF sprav_serv PROMPT 'Справочник анкетных данных клиента' ;
  MESSAGE 'Просмотр справочника анкетных данных клиента'
DEFINE BAR 12 OF sprav_serv PROMPT '\-'
DEFINE BAR 13 OF sprav_serv PROMPT 'Справочник карточных счетов клиента' ;
  MESSAGE 'Редактирование карточных счетов клиента'
DEFINE BAR 14 OF sprav_serv PROMPT '\-'
DEFINE BAR 15 OF sprav_serv PROMPT 'Справочник по выданным картам' ;
  MESSAGE 'Просмотр справочника по картам'
DEFINE BAR 16 OF sprav_serv PROMPT '\-'
DEFINE BAR 17 OF sprav_serv PROMPT 'Справочник данных по клиентам' ;
  MESSAGE 'Просмотр справочника по всем клиентам'
DEFINE BAR 18  OF sprav_serv PROMPT '\-'
DEFINE BAR 19  OF sprav_serv PROMPT 'Справочник данных по всем филиалов' ;
  MESSAGE 'Просмотр справочника реквизитов всех филиалов входящих в КБ "CARD-банк"'
DEFINE BAR 20  OF sprav_serv PROMPT '\-'
DEFINE BAR 21  OF sprav_serv PROMPT 'Справочник лицевых счетов в ОДБ RsBank' ;
  MESSAGE 'Просмотр справочника лицевых счетов в ОДБ RsBank'
DEFINE BAR 22 OF sprav_serv PROMPT '\-'
DEFINE BAR 23 OF sprav_serv PROMPT 'Справочник счетов по устройствам "CARD-банк"' ;
  MESSAGE 'Просмотр и редактирование справочника счетов по устройствам "CARD-банк", счета для ПОС и АТМ'
DEFINE BAR 24 OF sprav_serv PROMPT '\-'
DEFINE BAR 25 OF sprav_serv PROMPT 'Справочник открываемых таблиц в ПО "CARD-виза"' ;
  MESSAGE 'Справочник открываемых таблиц в ПО "CARD-виза"'
DEFINE BAR 26 OF sprav_serv PROMPT '\-'
DEFINE BAR 27 OF sprav_serv PROMPT 'Справочник тарифных ставок' ;
  MESSAGE 'Редактирование справочника тарифных ставок по выдаче и обслуживанию карт Visa'
DEFINE BAR 28 OF sprav_serv PROMPT '\-'
DEFINE BAR 29 OF sprav_serv PROMPT 'Паспортные данные доверенного лица клиента' ;
  MESSAGE 'Паспортные данные доверенного лица клиента'
DEFINE BAR 30 OF sprav_serv PROMPT '\-'
DEFINE BAR 31 OF sprav_serv PROMPT 'Справочник террористов' ;
  MESSAGE 'Просмотр справочника террористов'

DO CASE
  CASE pr_rabota_sql = .T.  && Работа происходит с таблицами SQL Server

* WAIT WINDOW 'Работа происходит с таблицами SQL Server'

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

  CASE pr_rabota_sql = .F.  && Работа происходит с локальными таблицами

* WAIT WINDOW 'Работа происходит с локальными таблицами'

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
DEFINE BAR 1  OF servis_bik PROMPT 'Импорт справочника БИК из Rs-Bank' ;
  MESSAGE 'Импорт данных справочника БИК из Rs-Bank и формирование внутреннего справочника БИК'
DEFINE BAR 2  OF servis_bik PROMPT '\-'
DEFINE BAR 3  OF servis_bik PROMPT 'Просмотр локального справочника БИК' ;
  MESSAGE 'Табличный просмотр локального справочника БИК'
DEFINE BAR 4  OF servis_bik PROMPT '\-'
DEFINE BAR 5  OF servis_bik PROMPT 'Создать доступ к справочнику БИК в RsBank' ;
  MESSAGE 'Создание удаленного представления к справочнику БИК в RsBank'

ON SELECTION BAR 1 OF servis_bik DO import_bik IN Bnkseek
ON SELECTION BAR 3 OF servis_bik DO brows_bik IN Bnkseek
ON SELECTION BAR 5 OF servis_bik DO create_bankdprt IN Bnkseek

**************************************************************** DEFINE POPUP import_kurs ****************************************************************

DEFINE POPUP import_kurs ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1  OF import_kurs PROMPT 'Импорт курсов валют из Rs-Bank' ;
  MESSAGE 'Импорт данных справочника курсов валют из Rs-Bank и обновление внутреннего справочника курсов валют'
DEFINE BAR 2  OF import_kurs PROMPT '\-'
DEFINE BAR 3  OF import_kurs PROMPT 'Просмотр локального справочника курсов' ;
  MESSAGE 'Табличный просмотр локального справочника курсов валют'
DEFINE BAR 4  OF import_kurs PROMPT '\-'
DEFINE BAR 5  OF import_kurs PROMPT 'Создать доступ к справочнику курсов валют в RsBank' ;
  MESSAGE 'Создание удаленного представления к справочнику курсов валют в RsBank'

ON SELECTION BAR 1 OF import_kurs DO import_kurs IN Impkursval
ON SELECTION BAR 3 OF import_kurs DO brows_kurs IN Impkursval
ON SELECTION BAR 5 OF import_kurs DO create_kurs IN Impkursval


****************************************************************** DEFINE POPUP read_data **************************************************************

DEFINE POPUP read_data ;
  FONT 'Times New Roman', num_schrift_g  STYLE 'B' ;
  MARGIN COLOR SCHEME 2
DEFINE BAR 1 OF read_data PROMPT 'Выверка остатков и оборотов по счетам CARD-виза и Трансмастер' ;
  MESSAGE 'Выверка остатков и оборотов по счетам CARD-виза и Трансмастер'
DEFINE BAR 2 OF read_data PROMPT '\-'
DEFINE BAR 3 OF read_data PROMPT 'Выверка остатков и оборотов по счетам в ведомости остатков' ;
  MESSAGE 'Выверка остатков и оборотов по счетам в ведомости остатков'
DEFINE BAR 4 OF read_data PROMPT '\-'
DEFINE BAR 5 OF read_data PROMPT 'Выравнивания остатков и оборотов по карточным счетам' ;
  MESSAGE 'Выравнивания остатков и оборотов по карточным счетам'
DEFINE BAR 6 OF read_data PROMPT '\-'
DEFINE BAR 7 OF read_data PROMPT 'Выверка ведомости остатков по данным текстовой ведомости из УПК' ;
  MESSAGE 'Выверка ведомости остатков по данным ранее закачанной в таблицу текстовой ведомости из УПК'
DEFINE BAR 8 OF read_data PROMPT '\-'
DEFINE BAR 9 OF read_data PROMPT 'Выверка справочника счетов по данным текстовой ведомости из УПК' ;
  MESSAGE 'Выверка справочника счетов по данным ранее закачанной в таблицу текстовой ведомости из УПК'
DEFINE BAR 10 OF read_data PROMPT '\-'
DEFINE BAR 11 OF read_data PROMPT 'Правка параметров карточных счетов в справочнике закрытых карт' ;
  MESSAGE 'Правка параметров карточных счетов в справочнике закрытых карт методом сканирования справочника открытых карт'
DEFINE BAR 12 OF read_data PROMPT '\-'
DEFINE BAR 13 OF read_data PROMPT 'Ведомость по разбиению на 12 месяцев плат за обслуживание счета с 01.07.2006 года' ;
  MESSAGE 'Формирование ведомости по разбиению на 12 месяцев плат за обслуживание счета с 01.07.2006 года'
DEFINE BAR 14 OF read_data PROMPT '\-'
DEFINE BAR 15 OF read_data PROMPT 'Ведомость по разбиению на 12 месяцев плат за обслуживание счета с 01.01.2006 года' ;
  MESSAGE 'Формирование ведомости по разбиению на 12 месяцев плат за обслуживание счета с 01.01.2006 года'
DEFINE BAR 16 OF read_data PROMPT '\-'
DEFINE BAR 17 OF read_data PROMPT 'Обновление паспортных полей в справочнике заявлений' ;
  MESSAGE 'Обновление паспортных полей в справочнике заявлений методом сканирования справочника клиентов и счетов'
DEFINE BAR 18 OF read_data PROMPT '\-'
DEFINE BAR 19 OF read_data PROMPT 'Заполнение долгового справочника значениями типов карт из справочников счетов' ;
  MESSAGE 'Разовая служебная процедура для корректировки работы комплекса'
DEFINE BAR 20 OF read_data PROMPT '\-'
DEFINE BAR 21 OF read_data PROMPT 'Преобразование адресных полей в справочнике заявлений в формат КЛАДР' ;
  MESSAGE 'Преобразование адресных полей в справочнике заявлений в формат КЛАДР методом сканирования справочника составленных заявлений'
DEFINE BAR 22 OF read_data PROMPT '\-'
DEFINE BAR 23 OF read_data PROMPT 'Выгрузка справочника заявлений для редактирования адресов по КЛАДР' ;
  MESSAGE 'Выгрузка справочника заявлений на открытие карты для внешнего редактирования адресных полей по КЛАДР'
DEFINE BAR 24 OF read_data PROMPT '\-'
DEFINE BAR 25 OF read_data PROMPT 'Редактирование адресов в выгруженном справочнике заявлений по КЛАДР' ;
  MESSAGE 'Редактирование адресов в выгруженном справочнике заявлений по КЛАДР'
DEFINE BAR 26 OF read_data PROMPT '\-'
DEFINE BAR 27 OF read_data PROMPT 'Загрузка справочника заявлений после редактирования адресов по КЛАДР' ;
  MESSAGE 'Загрузка справочника заявлений на открытие карты после внешнего редактирования адресных полей по КЛАДР'
DEFINE BAR 28 OF read_data PROMPT '\-'
DEFINE BAR 29 OF read_data PROMPT 'Правка номеров карточных счетов в справочнике заявлений' ;
  MESSAGE 'Правка номеров карточных счетов в справочнике заявлений методом сканирования справочника карт'
DEFINE BAR 30 OF read_data PROMPT '\-'
DEFINE BAR 31 OF read_data PROMPT 'Выверка истории остатков по данным присланным из Головного банка' ;
  MESSAGE 'Выверка истории остатков методом сканирования данных присланных из Головного банка'
DEFINE BAR 32 OF read_data PROMPT '\-'
DEFINE BAR 33 OF read_data PROMPT 'Выверка ведомости остатков по данным присланным из Головного банка' ;
  MESSAGE 'Выверка ведомости остатков методом сканирования данных присланных из Головного банка'
DEFINE BAR 34 OF read_data PROMPT '\-'
DEFINE BAR 35 OF read_data PROMPT 'Добавление данных за непрерывную череду выходных дней в ведомость остатков' ;
  MESSAGE 'Добавление отсутствующих данных за выходные дни в ведомость остатков с указанием последнего рабочего дня'
DEFINE BAR 36 OF read_data PROMPT '\-'
DEFINE BAR 37 OF read_data PROMPT 'Удаление записей в таблице "Ведомость остатков" за выбранную дату' ;
  MESSAGE 'Данная процедура выполняет удаление записей в таблице "Ведомость остатков" за выбранную дату'
DEFINE BAR 38 OF read_data PROMPT '\-'
DEFINE BAR 39 OF read_data PROMPT 'Удаление ВСЕХ данных по выбранному клиенту и счету во ВСЕХ таблицах' ;
  MESSAGE 'Удаление ВСЕХ данных по выбранному клиенту и счету во ВСЕХ таблицах'
DEFINE BAR 40 OF read_data PROMPT '\-'
DEFINE BAR 41 OF read_data PROMPT 'Удаление ошибочно открытых номер карт в справочнике карт и карточных счетов' ;
  MESSAGE 'Удаление ошибочно открытых номер карт в справочнике карт и карточных счетов'
DEFINE BAR 42 OF read_data PROMPT '\-'
DEFINE BAR 43 OF read_data PROMPT 'Удаление ВСЕХ данных в справочниках клиентов и карточных счетов' ;
  MESSAGE 'Удаление ВСЕХ данных в справочнике клиентов и справочнике карт и карточных счетов'

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
DEFINE BAR 1 OF read_card_nls PROMPT 'Изменение балансовых счетов по распоряжению № 154 на новые' ;
  MESSAGE 'Изменение балансовых счетов в справочничках по распоряжению № 154 на новые'
DEFINE BAR 2 OF read_card_nls PROMPT '\-'
DEFINE BAR 3 OF read_card_nls PROMPT 'Изменение балансовых счетов с новых на используемые ранее' ;
  MESSAGE 'Изменение балансовых счетов с новых на используемые ранее'
DEFINE BAR 4 OF read_card_nls PROMPT '\-'
DEFINE BAR 5 OF read_card_nls PROMPT 'Формирование карточных счетов по новому алгоритму' ;
  MESSAGE 'Формирование карточных счетов по новому алгоритму'
DEFINE BAR 6 OF read_card_nls PROMPT '\-'
DEFINE BAR 7 OF read_card_nls PROMPT 'Проверка на дублирование карточных счетов' ;
  MESSAGE 'Проверка на дублирование карточных счетов'
DEFINE BAR 8 OF read_card_nls PROMPT '\-'
DEFINE BAR 9 OF read_card_nls PROMPT 'Выгрузка текстового файла для конвертации карточных счетов' ;
  MESSAGE 'Выгрузка текстового файла для конвертации карточных счетов'
DEFINE BAR 10 OF read_card_nls PROMPT '\-'
DEFINE BAR 11 OF read_card_nls PROMPT 'Изменение старых счетов на новые во всех таблицах' ;
  MESSAGE 'Изменение старых счетов на новые во всех таблицах методом сканирования'
DEFINE BAR 12 OF read_card_nls PROMPT '\-'
DEFINE BAR 13 OF read_card_nls PROMPT 'Просмотр остатков на карточных счетах с подключенной SMS-услугой' ;
  MESSAGE 'Просмотр остатков на карточных счетах с подключенной SMS-услугой, срок оплаты которой, ближайшие дни'
DEFINE BAR 14 OF read_card_nls PROMPT '\-'
DEFINE BAR 15 OF read_card_nls PROMPT 'Заполнение поля транзитного счета в таблице "Ведомость остатков"' ;
  MESSAGE 'Заполнение поля транзитного счета в таблице "Ведомость остатков" методом сканирования справочника счетов'
DEFINE BAR 16 OF read_card_nls PROMPT '\-'
DEFINE BAR 17 OF read_card_nls PROMPT 'Добавление из архивной таблицы заявлений в рабочую таблицу заявлений' ;
  MESSAGE 'Добавление из архивной таблицы заявлений в рабочую таблицу заявлений на получение пластиковой карты'

DO CASE
  CASE pr_rabota_sql = .T.  && Работа происходит с таблицами SQL Server

    ON SELECTION BAR 1 OF read_card_nls DO red_bal_num_new IN Servis_sql
    ON SELECTION BAR 3 OF read_card_nls DO red_bal_num_str IN Servis_sql
    ON SELECTION BAR 5 OF read_card_nls DO scan_formir_card_nls IN Formir_card_nls
    ON SELECTION BAR 7 OF read_card_nls DO scan_dubl_card_nls IN Dubl_card_nls
    ON SELECTION BAR 9 OF read_card_nls DO start IN Konvert_card_nls
    ON SELECTION BAR 11 OF read_card_nls DO start IN Update_new_card_acct
    ON SELECTION BAR 13 OF read_card_nls DO start IN Smshome
    ON SELECTION BAR 15 OF read_card_nls DO start IN Tran_vedom_ost
    ON SELECTION BAR 17 OF read_card_nls DO start IN Dobav_zajava_arx

  CASE pr_rabota_sql = .F.  && Работа происходит с локальными таблицами

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
DEFINE BAR 1  OF servis_serv PROMPT 'Перевод закрытых карт в архив' ;
  MESSAGE 'Процедура перевода закрытых карт в архив'
DEFINE BAR 2  OF servis_serv PROMPT '\-'
DEFINE BAR 3  OF servis_serv PROMPT 'Изменение структуры полей в таблицах' ;
  MESSAGE 'Изменение структуры полей в таблицах'
DEFINE BAR 4  OF servis_serv PROMPT '\-'
DEFINE BAR 5  OF servis_serv PROMPT 'Упаковка и обслуживание таблиц ПО "CARD-виза"' ;
  MESSAGE 'Упаковка и обслуживание таблиц ПО "CARD-виза"'
DEFINE BAR 6  OF servis_serv PROMPT '\-'
DEFINE BAR 7  OF servis_serv PROMPT 'Введение признака типа карты в номер счета' SKIP ;
  MESSAGE 'Введение признака типа карты в номер счета Electron тип 1; Classic тип 2; Gold тип 3; Business тип 4'
DEFINE BAR 8  OF servis_serv PROMPT '\-'
DEFINE BAR 9  OF servis_serv PROMPT 'Выверка остатков и оборотов по счетам за открытую дату' ;
  MESSAGE 'Процедура выверки остатков и оборотов по счетам методом просчета проведенных документов за открытую дату'
DEFINE BAR 10  OF servis_serv PROMPT '\-'
DEFINE BAR 11  OF servis_serv PROMPT 'Загрузка данных выгруженных филиалами в головной банк' ;
  MESSAGE 'Загрузка данных выгруженных филиалами в головной банк'
DEFINE BAR 12 OF servis_serv PROMPT '\-'
DEFINE BAR 13 OF servis_serv PROMPT 'Пакетный режим добавления новых полей' ;
  MESSAGE 'В пакетном режиме добавляются новые поля в таблицы проекта, указанные в таблице modify.dbf'
DEFINE BAR 14 OF servis_serv PROMPT '\-'
DEFINE BAR 15 OF servis_serv PROMPT 'Удаление записей в протоколе ошибок' ;
  MESSAGE 'Удаление записей в протоколе ошибок errorlog.dbf'
DEFINE BAR 16 OF servis_serv PROMPT '\-'
DEFINE BAR 17 OF servis_serv PROMPT 'Изменение величины процентной ставки списком' ;
  MESSAGE 'Изменение величины процентной ставки списком по балансовому счету'
DEFINE BAR 18 OF servis_serv PROMPT '\-'
DEFINE BAR 19 OF servis_serv PROMPT 'Экспорт данных за прошедшие года в архивные таблицы' ;
  MESSAGE 'Экспорт данных за прошедшие года из рабочих таблиц в архивные таблицы'
DEFINE BAR 20 OF servis_serv PROMPT '\-'
DEFINE BAR 21 OF servis_serv PROMPT 'Импорт данных за прошедшие года из архивных таблиц' ;
  MESSAGE 'Импорт данных за прошедшие года из архивных таблиц в рабочие таблицы'
DEFINE BAR 22 OF servis_serv PROMPT '\-'
DEFINE BAR 23 OF servis_serv PROMPT 'Правка ведомости остатков по истории остатков' ;
  MESSAGE 'Правка ведомости остатков по истории остатков сформированной ранее'
DEFINE BAR 24 OF servis_serv PROMPT '\-'
DEFINE BAR 25 OF servis_serv PROMPT 'Удаление дублирующихся записей в таблицах' ;
  MESSAGE 'Удаление дублирующихся записей в таблицах'
DEFINE BAR 26 OF servis_serv PROMPT '\-'
DEFINE BAR 27 OF servis_serv PROMPT 'Установка признака запрета входа в программу' ;
  MESSAGE 'Установка признака запрета входа в программу, используется в процедурах требующих монопольного доступа к таблицам'
DEFINE BAR 28 OF servis_serv PROMPT '\-'
DEFINE BAR 29 OF servis_serv PROMPT 'Справочник календарных дат используемых в ПО' ;
  MESSAGE 'Справочник календарных дат используемых в ПО'
DEFINE BAR 30 OF servis_serv PROMPT '\-'
DEFINE BAR 31 OF servis_serv PROMPT 'Импорт новых условий использования карт банка' ;
  MESSAGE 'Импорт новых УСЛОВИЙ ИСПОЛЬЗОВАНИЯ МЕЖДУНАРОДНЫХ КАРТ CARD-БАНКА'
DEFINE BAR 32 OF servis_serv PROMPT '\-'
DEFINE BAR 33 OF servis_serv PROMPT 'Процедуры обработки файлов с электронной подписью' ;
  MESSAGE 'Сервисные процедуры по работе с файлами, подписанными электронной подписью'
DEFINE BAR 34 OF servis_serv PROMPT '\-'
DEFINE BAR 35 OF servis_serv PROMPT 'Автоматическое обслуживание таблиц ПО' ;
  MESSAGE 'Автоматическое обслуживание таблиц ПО "CARD-виза"'

DO CASE
  CASE pr_rabota_sql = .T.  && Работа происходит с таблицами SQL Server

    DEFINE BAR 36 OF servis_serv PROMPT '\-'
    DEFINE BAR 37 OF servis_serv PROMPT 'Экспорт данных из локальных таблиц в таблицы SQL Server' ;
      MESSAGE 'Экспорт данных из локальных таблиц в таблицы SQL Server 2000 и  SQL Server 2005'

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

  CASE pr_rabota_sql = .F.  && Работа происходит с локальными таблицами

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


*  * Первый пункт основной линейки меню
*  DEFINE PAD _msm_edit OF _MSYSMENU PROMPT "Правка" COLOR SCHEME 3
*  * Прочие пункты основной линейки меню

*  DEFINE POPUP _medit MARGIN RELATIVE SHADOW COLOR SCHEME 4
*  DEFINE BAR _MED_UNDO OF _medit PROMPT "Отменить" ;
*  KEY CTRL+Z, "CTRL+Z"
*  DEFINE BAR _MED_REDO OF _medit PROMPT "Вернуть" ;
*  KEY CTRL+R, "CTRL+R"
*  DEFINE BAR _MED_SP100 OF _medit PROMPT "\-"
*  DEFINE BAR _MED_CUT OF _medit PROMPT "Вырезать" ;
*  KEY CTRL+X, "CTRL+X"
*  DEFINE BAR _MED_COPY OF _medit PROMPT "Копировать" ;
*  KEY CTRL+C, "CTRL+C"
*  DEFINE BAR _MED_PASTE OF _medit PROMPT "Вставить" ;
*  KEY CTRL+V, "CTRL+V"
*  DEFINE BAR _MED_CLEAR OF _medit PROMPT "Очистить"
*  DEFINE BAR _MED_SP200 OF _medit PROMPT "\-"
*  DEFINE BAR _MED_SLCTA OF _medit PROMPT "Выделить все" ;
*  KEY CTRL+A, "CTRL+A"
*  [/code]


*********************************************************************************************************************************************************


