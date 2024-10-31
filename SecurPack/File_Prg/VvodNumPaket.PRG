
***************************************************************************
*** Импорт секьюрпаков полученных на точках клиента и сдаваемых в кассу ***
***************************************************************************

******************************************************************************* PROCEDURE start ****************************************************************

PROCEDURE start
LPARAMETERS pMode

*** pMode = 0 режим стандартной и многопиновой сумки
*** pMode = 1 режим суперсумки

* WAIT WINDOW 'f_IdShift = ' + ALLTRIM(STR(f_IdShift)) TIMEOUT 3

PRIVATE l_mode, l_qparams

l_qparams = PCOUNT()  && Возвращает количество параметров, переданных текущей программе, процедуре или пользовательской функции

IF l_qparams = 0
  l_mode = 0
ELSE
  l_mode = pMode
ENDIF

* WAIT WINDOW 'pMode = ' + ALLTRIM(STR(pMode)) TIMEOUT 3
* WAIT WINDOW '1 - l_mode = ' + ALLTRIM(STR(l_mode)) TIMEOUT 3

sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()
HIDE POPUP (popup_ima)

num_recno = 0

PoiskTSD_IdBudgetNum = 0

* --------------------- В запускаемой процедуре проверяется не установлен ли Старшим смены флаг блокировки на работу в АРМе --------------------- *

DO start IN Kontrol_flag_blok

IF flag_Shift_Blok = 1
  RETURN
ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF USED('SecurPack') = .T.
  USE IN SELECT('SecurPack')
ENDIF

IF USED('TypeValue') = .T.
  USE IN SELECT('TypeValue')
ENDIF

IF USED('BudgetDefect') = .T.
  USE IN SELECT('BudgetDefect')
ENDIF

IF USED('SignProc') = .T.
  USE IN SELECT('SignProc')
ENDIF

DIMENSION ar[1,2]

STORE SPACE(0) TO DesktopPath, CurFile, workPath, IdRoute
STORE 0 TO numRowF, numRow, fvDate, IdRoute

f_Ch_FilialCode = 0  && Флаг отсутствия дополнительного ПИНа в сопроводительной ведомости

f_Flg2PIN = 0  && Флаг наличия дополнительных ПИНов, принудительно устонавливаем в "0"

* WAIT WINDOW 'ShiftDate_ = ' + DTOC(TTOD(ShiftDate_)) TIMEOUT 3

dt_god = ALLTRIM(STR(YEAR(DATE())))
dt_mes = ALLTRIM(STR(MONTH(DATE())))
dt_den = ALLTRIM(STR(DAY(DATE())))

f_IncFullName = ''
f_BudgetNum = ''
f_DatInkass = {}

f_FilialCode = ''
f_TotalSum = 0.0

f_FilialName = ''
f_SignProc = 'Касса пересчета'
f_DefBud = .F.  && Дефекты
f_IdTypeValue = 0 && Тип ценности
f_IsTransit = '0'
f_IdRoute = ''

colvo_IdRoute = 0

f_SuperBudget = SPACE(0)
f_SuperQty = 0
f_WorkType = 0

perinp1 = ''

f_Ch_KeyATM = 0
f_Ch_KeyATMLost = 0

* ------------------------------------------------------------------------------------------------------------------------------- *

perinp1 = ''

DO CASE
  CASE sergeizhuk_flag = .F.

    perinp1 = ''

  CASE sergeizhuk_flag = .T.

    perinp1 = '63'

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------- *

* ------------------------------------------------------------------------------------------------------------------------------- *
DO FORM FormInp1 WITH 'Ввод номера маршрута', 'Введите номер МАРШРУТА:'
* ------------------------------------------------------------------------------------------------------------------------------- *

* Проверка значения.

IF EMPTY(perinp1) = .T.
  RETURN -1
ENDIF

f_IdRoute = ALLTRIM(perinp1)

* WAIT WINDOW 'Введен номер маршрута f_IdRoute = ' + f_IdRoute TIMEOUT 3

* ---------------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

IF TYPE('f_IdRoute') == 'C'
  f_IdRoute = VAL(ALLTRIM(f_IdRoute))
ENDIF

*  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*    SELECT *
*    FROM Route (READPAST)
*    WHERE RouteNumber = ?f_IdRoute

*  ENDTEXT

* ------------------------------------------------------------------------------------ *

TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
  EXECUTE Formir_Route ?f_IdRoute
ENDTEXT

* ------------------------------------------------------------------------------------ *

IF RunSQL_Load('Route', @ar)  <>  1
  DO exit_sql_server IN RunSQL_Load
  MESSAGEBOX('Введенный маршрут не существует ... ', 16, 'Операция движения секьюрпаков')
  RETURN -1
ENDIF

* BROWSE WINDOW brows TITLE 'RunSQL Route'

IF ar[1, 2] <> 1
  DO exit_sql_server IN RunSQL_Load
  MESSAGEBOX('Маршрут номер ' + ALLTRIM(STR(f_IdRoute)) + ' в базе не зарегистрирован', 16, 'Операция Отменена')
  RETURN -1
ENDIF

SELECT Route
GO TOP

f_IdRoute = Route.Id && Номер маршрута

* BROWSE WINDOW brows TITLE 'Route'

* ------------------------------------------------------------------------------------------------------------------------------- *

DO CASE
  CASE sergeizhuk_flag = .F.

    perinp1 = ''

  CASE sergeizhuk_flag = .T.

    perinp1 = '5555'

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------- *
DO FORM FormInp1 WITH 'Ввод табельного номера', 'Введите табельный номер:'
* ------------------------------------------------------------------------------------------------------------------------------- *

* Проверка значения.

IF EMPTY(perinp1) = .T.
  RETURN -1
ENDIF

f_TabNumber = ALLTRIM(perinp1)

* WAIT WINDOW 'Введен табельный номер f_TabNumber = ' + ALLTRIM(f_TabNumber) TIMEOUT 3

* ---------------------------------------------------------- Операция создания ведомости по секюрПакам ------------------------------------------------------------------------------------- *

* PlaceTranzit = 6 && код транзита
* PlacePRK = 2 && код ПРК

f_PlaceTranzit = PlaceTranzit  && 6 - код транзита, определяется в головном модуле

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT Id, IdPlace, IdRegister, SignName
  FROM SignProc (READPAST)
  WHERE IdPlace = ?f_PlaceTranzit

ENDTEXT

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF RunSQL_Load('SignProc', @ar) <> 1 OR ar[1,2] < 1
  DO exit_sql_server IN RunSQL_Load
  MESSAGEBOX('Таблица - SignProc, необходимая для работы НЕ СФОРМИРОВАНА', 16, 'Операция создания ведомости по секюрПакам Прекращена')
  RETURN -1
ENDIF

f_IdSProcTrn = SignProc.Id
f_SProcNmTrn = SignProc.SignName

* BROWSE WINDOW brows TITLE 'RunSQL SignProc'

IF USED('SignProc') = .T.
  USE IN SELECT('SignProc')
ENDIF

* --------------------------------------------------------- Создаем курсор дефектов для сумки ------------------------------------------------------------------------------------------------------ *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT CAST(0 AS bit) AS fl, bd.DefectName, bd.IsOpen, bd.DefectNameFull, bd.DefectCode, bd.Id, bd.AllowZero, bd.ChkJKOnAct
  FROM BudgetDefectNew AS bd (READPAST)
  WHERE bd.IsActive = 1
  ORDER BY Id ASC

ENDTEXT

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF RunSQL_Load('BudgetDefect', @ar) <> 1
  DO exit_sql_server IN RunSQL_Load
  MESSAGEBOX('Таблица - BudgetDefect, необходимая для работы НЕ СФОРМИРОВАНА', 16, 'Операция создания ведомости по секюрПакам Прекращена')
  RETURN -1
ENDIF

IF ar[1,2] < 1
  DO exit_sql_server IN RunSQL_Load
  MESSAGEBOX('Нет списка доступных дефектов!!!', 16, 'Операция создания ведомости по секюрПакам Прекращена')
  =CloseDB()
  RETURN -1
ENDIF

* BROWSE WINDOW brows TITLE 'RunSQL BudgetDefect'

* ---------------------------------------------------------------------- Заполняем справочник Типов Ценностей --------------------------------------------------------------------------------- *

TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT Id, TypeValueName, TypeValueKindNew AS TypeValueKind
  FROM TypeValue (READPAST)
  WHERE IsActive = 1
  ORDER BY Id

ENDTEXT

* WAIT WINDOW 'Заполняем справочник Типов Ценностей' TIMEOUT 1

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

IF RunSQL_Load('TypeValue', @ar) <> 1
  DO exit_sql_server IN RunSQL_Load
  MESSAGEBOX('Таблица - TypeValue, необходимая для работы НЕ СФОРМИРОВАНА', 16, 'Операция создания ведомости по секюрПакам Прекращена')
  =CloseDB()
  RETURN -1
ENDIF

* BROWSE WINDOW brows TITLE 'RunSQL TypeValue'

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*    SELECT *
*    FROM CashMessenger (READPAST)
*    WHERE TabNumber = ?f_TabNumber AND IsActive = 1

*  ENDTEXT

* ------------------------------------------------------------------------------------------------ *

TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
  EXECUTE Formir_CashMessenger_TabNumber ?f_TabNumber
ENDTEXT

* ------------------------------------------------------------------------------------------------ *

IF RunSQL_Load('CashMessenger', @ar)  <>  1
  DO exit_sql_server IN RunSQL_Load
  MESSAGEBOX('Введенный номер не существует ... ', 16, 'Операция движения секьюрпаков')
  RETURN -1
ENDIF

* BROWSE WINDOW brows TITLE 'RunSQL CashMessenger'

IF ar[1, 2] = 0
  DO exit_sql_server IN RunSQL_Load
  MESSAGEBOX('Табельный номер ' + ALLTRIM(f_TabNumber) + ' в базе не зарегистрирован', 16, 'Операция Отменена')
  RETURN -1
ENDIF

IF ar[1, 2] > 1
  DO exit_sql_server IN RunSQL_Load
  MESSAGEBOX('Табельный номер ' + ALLTRIM(f_TabNumber) + ' зарегистрирован в базе более одного раза.' + CHR(13) + ;
    'Табельный номер должен быть уникальным значением, свяжитесь с отделом кадров.', 16, 'Операция Отменена')
  RETURN -1
ENDIF

SELECT CashMessenger
GO TOP

f_IdCashMessenger = CashMessenger.Id  && Id инкассатора
f_TabNumber = ALLTRIM(CashMessenger.TabNumber) && Табельный номер
f_MessFullName = ALLTRIM(CashMessenger.FullName) && ФИО инкассатора
f_FullName = ALLTRIM(CashMessenger.FullName) && ФИО инкассатора
f_IncFullName = ALLTRIM(CashMessenger.FullName) && ФИО инкассатора

* BROWSE WINDOW brows TITLE 'CashMessenger'

* ------------------------------------------------ Выбираем номиналы банкнот/монет ----------------------------------------------- *

TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

  SELECT Valuable.*,
  TypeValue.TypeValueName,TypeValue.TypeValueKindNew,TypeValue.ValCode,
  FaceValue.FaceValue,FaceValue.FaceValueName,
  DetailValue.DetailValueName
  FROM Valuable (READPAST)
  LEFT JOIN TypeValue (READPAST) ON TypeValue.Id = Valuable.IdTypeValue
  LEFT JOIN FaceValue (READPAST) ON FaceValue.Id = Valuable.IdFaceValue
  LEFT JOIN DetailValue (READPAST) ON DetailValue.Id = Valuable.IdDetailValue
  WHERE Valuable.IsActive = 1
  ORDER BY FaceValue.FaceValue DESC

ENDTEXT

rc = RunSQL_Load('TypeFaceValue', @ar)

* BROWSE WINDOW brows TITLE 'TypeFaceValue'

IF rc <> 1
  MESSAGEBOX('Ошибка загрузки номиналов - работа с дефектами будет невозможна', 16, 'Информация об ошибке ... ')
ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

CREATE CURSOR SecurPack ( ;
  IdShift INTEGER NOT NULL,;
  BudgetNum CHAR(25) NOT NULL,;
  IdUser INTEGER NOT NULL,;
  IdTsd INTEGER NULL,;
  FilialCode CHAR(15) NOT NULL,;
  TabNumber CHAR(10) NOT NULL,;
  OperCode CHAR(1) NULL,;
  State CHAR(1) NULL,;
  IdRoute INTEGER NOT NULL,;
  DateInkas DATETIME NULL,;
  DtIncome DATETIME NULL,;
  CurDate DATETIME NULL,;
  FileName CHAR(150) NULL,;
  TotalSum NUMERIC(18, 2) NULL,;
  Pin2Sum NUMERIC(18, 2) NULL DEFAULT 0,;
  Pr_60322 INTEGER,;
  TotalSum_60322 NUMERIC(18, 2) NULL,;
  TotalSumGlPin NUMERIC(18, 2) NULL,;
  IdTypeValue INTEGER NOT NULL,;
  TypeValueKind CHAR(10) NOT NULL,;
  IsTransit CHAR(1) NOT NULL,;
  Flg2PIN INTEGER NOT NULL,;
  Sost CHAR(1) NOT NULL,;
  WorkType CHR(1) NOT NULL DEFAULT '0',;
  SuperBudget CHAR(10) NULL DEFAULT SPACE(0),;
  SuperQty INT NOT NULL DEFAULT 0,;
  IdPackDefect INTEGER NULL,;
  DtBeginPer DATETIME NULL,;
  IdJk INTEGER NULL DEFAULT 0,;
  LostJk INTEGER NULL DEFAULT 0,;
  IdFilial INTEGER NULL,;
  IdKeyATMDocMain INTEGER NULL DEFAULT 0,;
  IdKeyATM INTEGER NULL DEFAULT 0,;
  ColvoKey INTEGER NULL DEFAULT 0)


CREATE CURSOR SecurPackDefect ( ;
  IdBudgetDefect int NOT null,;    &&Код дефекта
NumPacketsByDocs  int null,;    && Кол-во пакетов по документам
TotalSum    currency null,;    && количество денег
BudgetNumReal char(20) null,;    && номер сумки
PlombaNumAct  char(20) null,;    && Фактический номер пломбы
PlombaNumDocs char(20) null,;    && Номер пломбы по документам
PacketsNumbers Memo null,;      && Номера вложенных пакетов
IdShift INTEGER NOT NULL,;      && Поля внешнего ключа- должны соответствовать одноимённым полям SecurPack'а
BudgetNum CHAR(20) NOT NULL,;
  IdTypeValue INTEGER NOT NULL)


CREATE CURSOR MisLog_Cursor ( ;
  Id int,;      &&Будущий Id записи
IdSecurPack int,;  &&Пока не понятно - будет ли ссылка на SecurPack или всё же на отсутствующий SecurPackDefect.Id
IdValueable int,;  &&код купюры/монеты ссылка на Valuable.Id м.б. =0
ValueableName char(20),; &&Текстовое представление, если IdValueable==0 - значит был ввод вручную
Year int,;        && образца банкноты / год чеканки монеты
SomeData char(15),;  && Серия+номер банкноты / монетный двор
SomeData2 char(15),;  && Серия+номер банкноты / монетный двор
DefType int,;      && тип дефекта
IdShift int,;      && номер смены ПРК
IdTypeValue int,;    && Код валюты сумки - НЕ ПУТАТЬ с кодом валюты банкноты
BudgetNum char(20))    && номер сумки


CREATE CURSOR LocKeyATMSecurPackMain ( ;
  Id int NOT NULL,;
  Dup char(1) NOT NULL,;
  IdKeyATMDocMain int NOT NULL,;
  IdShiftPRK int NOT NULL,;
  IdRoutePRK int NOT NULL,;
  IdGuardOut int NULL,;
  IdCashMessengerOut int NULL,;
  DtCreate datetime2(0) NOT NULL,;
  IdUserPRK int NOT NULL,;
  DtAcceptHK datetime2(0) NULL,;
  IdShiftHK int NULL,;
  IdUserHK int NULL)


CREATE CURSOR LocKeyATMSecurPackSub ( ;
  Id NOT NULL,;
  IdKeyAtmSecurPackMain int NOT NULL,;
  IdKeyAtmDocMain int NOT NULL,;
  IdKeyATM int NOT NULL,;
  DtIn datetime2(0) NULL,;
  DtLost datetime2(0) NULL)


* ------------------------------ Обращение к таблице TSDKasper и формирование временного курсора с данными по маршруту, в котором будет искаться сумка ---------------------------- *

DO CASE
  CASE bar_num = 1 AND l_mode = 0  && Работа с ТСД ведется только в части простых сумок.

    DO start IN ImportKassaTsdKP WITH l_Mode

  CASE bar_num = 3 AND l_mode = 1  && Работа с ТСД ведется в части суперсумок.

    DO start IN ImportKassaTsdKP WITH l_Mode

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *


* ------------------------------------------------ Проведение проверки правильности даты смены ------------------------------------------------------------------------------------------------------------------------------------------------------ *

IF sergeizhuk_flag = .F.  && Для разработчиков ПО проверка не выполняется, поскольку ПО запускается с архивными датами

* WAIT WINDOW 'f_IdPost = ' + ALLTRIM(STR(f_IdPost)) TIMEOUT 3

  IF pr_rabota_arxiv = .F.  && Если значение переменной  = .F., то сознательно ведется работа в архивной дате

    IF f_IdPost <> 71  && У пользователя должность Старший смены ПРК его Id в таблице Post = 71

*      TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*        SELECT TOP 1 Sh.Id, Sh.IdOperDate, Sh.IdDepartment, Sh.IdUser, Sh.Status, Sh.ShiftDate, Sh.OpenDate, Sh.CloseDate, Od.OpDate, Sh.Blok
*        FROM Shift AS Sh
*        INNER JOIN OperDate AS Od ON Od.Id = Sh.IdOperDate
*        WHERE Sh.Id = ?f_IdShift
*        ORDER BY Sh.Id DESC

*      ENDTEXT

* ------------------------------------------------------------------------------------------------ *

      TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
        EXECUTE Formir_Kontrol_Blok ?f_IdShift
      ENDTEXT

* ------------------------------------------------------------------------------------------------ *

      =RunSQL_Load('CursOpShift', @ar)

* ---------------------------------------------------------------- *
      DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

      Pr_Use_CursOpShift = USED('CursOpShift') = .T.  && Признак сформировалась ли таблица CursOperDate

* BROWSE WINDOW brows

      DO CASE
        CASE Pr_Use_CursOpShift = .T.

          SELECT CursOpShift

          IF RECCOUNT() > 0

            GO TOP

            IF ISNULL(CursOpShift.CloseDate) = .F.

              MESSAGEBOX('Внимание! Вы работаете с ЗАКРЫТОЙ сменой!!!' + CHR(13) + 'Выйдите из АРМа и затем войдите повторно.' + CHR(13) + 'Проверьте дату смены и дату опердня.', 16, 'Операция проверки номера смены')
              RETURN

            ENDIF

          ENDIF

        CASE Pr_Use_CursOpShift = .F.

          MESSAGEBOX('Внимание! Не проведена проверка правильности даты смены!!!' + CHR(13) + 'Выйдите из АРМа и затем войдите повторно.' + CHR(13) + 'Проверьте дату смены и дату опердня.', 16, 'Операция проверки номера смены')
          RETURN

      ENDCASE

      IF USED('CursOpShift') = .T.
        SELECT CursOpShift
        USE
      ENDIF

    ENDIF

  ENDIF

ENDIF

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

SELECT SecurPack

pr_vvod = .F.

SCATTER MEMVAR BLANK

m.IdShift = f_IdShift
m.IdUser = f_IdUser

m.FilialCode = ''

m.TabNumber = ALLTRIM(f_TabNumber)

m.OperCode = '1'
m.State = '0'
m.IdRoute = f_IdRoute

m.DateInkas = ShiftDate_  && Возвращает дату открытой смены в формате 2019-11-25 00:00:00.000
f_DateInkas = DATETIME()

m.DtIncome = f_ShiftDate && Возвращает дату открытой смены в формате DATETIME()
f_DtIncome = f_ShiftDate && Возвращает дату открытой смены в формате DATETIME()

m.CurDate = DATETIME()

m.FileName = ''
f_FilialName = ''
f_BankName = ''

m.TotalSum = 0.00

m.Pin2Sum = 0.00
f_Pin2Sum = 0.00

m.TotalSumGlPin = 0.00

m.Sost = '1'

m.IdTypeValue = 1
m.TypeValueKind = 'RUB'
f_TypeValueKind = 'RUB'

m.Flg2PIN = 0

m.SuperQty = 0
m.SuperBudget = SPACE(0)
m.WorkType = '0'

m.IdJk = 0
m.LostJk = 0

m.IdKeyATMDocMain = 0
m.IdKeyATM = 0

m.ColvoKey = 0  && Кол-во ключей в связке

f_IdKeyATMDocMain = 0
f_IdKeyATM = 0

f_ColvoKey = 0  && Кол-во ключей в связке

Pr_DtLost = .F.

m.IdTsd = 0
f_IdTsd = 0

bar_office = 1
Pr_Lait = 0

* WAIT WINDOW '2 - l_mode = ' + ALLTRIM(STR(l_mode)) TIMEOUT 3

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

Start_DateTimeRoute = DATETIME()

* WAIT WINDOW 'Pr_Forma_Lait = ' + IIF(Pr_Forma_Lait = .T., '.T.', '.F.') TIMEOUT 3

DO CASE
  CASE Pr_Forma_Lait = .F.

    DO FORM VvodNewPaket WITH l_mode
    Pr_Lait = 0


  CASE Pr_Forma_Lait = .T.

    tex1 = 'Работать с формой - ввод данных с полной проверкой реквизитов'
    tex2 = 'Работать с формой - ввод данных на основании реквизитов из ТСД'
    l_bar = 3
    =popup_9(tex1, tex2, tex3, tex4, l_bar)

    ACTIVATE POPUP vibor BAR bar_office

    RELEASE POPUPS vibor

    IF NOT LASTKEY() = 27  && Если не было выхода по клавише ESC

      bar_office = IIF(BAR() >= 1, BAR(), 1)

      DO CASE
        CASE bar_office = 1  && Работать с формой - ввод данных с полной проверкой реквизитов

          DO FORM VvodNewPaket WITH l_mode
          Pr_Lait = 0

        CASE bar_office = 3  && Работать с формой - ввод данных на основании реквизитов из ТСД

          DO FORM VvodNewPaket_Lait WITH l_mode
          Pr_Lait = 1

      ENDCASE

    ENDIF

ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW 'SecurPack_IdFilial = ' + ALLTRIM(STR(SecurPack.IdFilial)) TIMEOUT 5
* WAIT WINDOW 'LASTKEY = ' + ALLTRIM(STR(LASTKEY())) TIMEOUT 3

IF NOT LASTKEY() = 27 AND pr_vvod = .T. AND pr_vvod_sql = .T.

  SELECT SecurPack
  GO TOP

*  При приёме ведомости необходимо правильно формировать BudgetHistoryPRK.Superbudget Для Питера - это младший ПИН (/3)
*  На данный момент superbudget, передаваемый в sp_Add_BudgetHistoryPRK_Pin2Sum  игнорируется для младших ПИНов

  LOCAL lAllowZero

  SCAN

    SELECT AllowZero ;
      FROM SecurPackDefect ;
      INNER JOIN BudgetDefect ON BudgetDefect.Id = SecurPackDefect.IdBudgetDefect ;
      WHERE SecurPackDefect.BudgetNum == SecurPack.BudgetNum AND SecurPackDefect.IdShift = SecurPack.IdShift AND SecurPackDefect.IdTypeValue = SecurPack.IdTypeValue ;
      INTO ARRAY t1

    lAllowZero = IIF(_TALLY <> 0 AND t1(1), .T., .F.)    &&lAllowZero будет только при наличии у сумки дефекта, допускающего нулевые суммы ПРК

    IF SecurPack.WorkType = '0' AND (EMPTY(ALLTRIM(SecurPack.BudgetNum)) = .T. OR EMPTY(ALLTRIM(SecurPack.FilialCode)) = .T. OR (SecurPack.TotalSum = 0 AND !lAllowZero))
      DELETE NEXT 1
    ENDIF

  ENDSCAN

  COUNT TO numRow
  GO TOP

* BROWSE WINDOW brows

* WAIT WINDOW 'f_IdShift = ' + ALLTRIM(STR(f_IdShift)) TIMEOUT 3
* WAIT WINDOW 'f_IdControler = ' + ALLTRIM(STR(f_IdControler)) TIMEOUT 3
* WAIT WINDOW 'f_ShiftDate = ' + ALLTRIM(TTOC(ShiftDate_)) TIMEOUT 3

* -------------------------------------------------------------------------------------------------------------------------------------- *
  CurFile = 'Ручной ввод данных в таблицу SecurPack' + ' - ' + LEFT(ALLTRIM(TIME()), 5)
* -------------------------------------------------------------------------------------------------------------------------------------- *

  SELECT SecurPack
  COUNT TO colvo_zap

  IF colvo_zap = 0
    =err('Уважаемый Пользователь, таблица пустая записывать НЕЧЕГО ... ')
    RETURN
  ENDIF

* ---------------------------------------------------------------- *
  DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

  tim1 = SECONDS()

  ACTIVATE WINDOW poisk
  @ WROWS()/3,3 SAY PADC('Минуточку терпения Уважаемый Пользователь, идет запись данных в SQL Server ... ',WCOLS())

  numRowF = 0

* ------------------------------------------------------------------------ Сканируем временную таблицу SELECT SecurPack ------------------------------------------------------ *

  SELECT SecurPack
  GO TOP

* BROWSE WINDOW brows

  SCAN FOR EMPTY(ALLTRIM(FilialCode)) = .F.

    f_BudgetNum = ALLTRIM(SecurPack.BudgetNum)

    f_FilialCode = ALLTRIM(SecurPack.FilialCode)

    f_IdFilial = SecurPack.IdFilial

    f_SkldCode = SkldCode

* ------------ Поле содержит дату открытой смены в формате 2014-08-16 00:00:00.000 при ручном вводе пакетов в Каспер и дату инкассации прописанную в ТСД  в формате 2014-08-16 21:15:30.000, при вводе данных с ТСД ------------- *

    f_DateInkas = SecurPack.DateInkas
    f_DtIncome = SecurPack.DtIncome

* ------------------------------------------ Эта дата может меняться кассиром при вводе реквизитов сумки на основании первичного документа к сумке, например БИЛЛА сдает сумки в один день, но сформированные за разные даты -------------------------- *

    f_CurDate = SecurPack.CurDate

    f_TabNumber = ALLTRIM(SecurPack.TabNumber)
    f_OperCode = ALLTRIM(SecurPack.OperCode)
    f_State = ALLTRIM(SecurPack.State)

    f_IdRoute = SecurPack.IdRoute

    f_TotalSum = SecurPack.TotalSum
    f_Pin2Sum = SecurPack.Pin2Sum
    f_TotalSumGlPin = SecurPack.TotalSumGlPin

    f_Pr_60322 = SecurPack.Pr_60322
    f_TotalSum_60322 = SecurPack.TotalSum_60322

    f_IdTypeValue = SecurPack.IdTypeValue
    f_TypeValueKind = SecurPack.TypeValueKind

    f_Sost = ALLTRIM(SecurPack.Sost)
    f_IsTransit = ALLTRIM(SecurPack.IsTransit)

    f_Flg2PIN = SecurPack.Flg2PIN  && Флаг наличия дополнительных ПИНов

    f_WorkType = SecurPack.WorkType
    f_SuperQty = SecurPack.SuperQty
    f_SuperBudget = SecurPack.SuperBudget

    f_DtBeginPer = SecurPack.DtBeginPer  && avg 2015-05-25 Дата 1й инкассации (пересчета) для точки  (Filial.DtBeginPer)

    f_IdPackDefect = SecurPack.IdPackDefect  && Поле для записи номера дефекта состояния сумки или ее документов из справочника BudgetDefect
    f_IdJk = SecurPack.IdJk  && Поле для записи Id явочной

    f_IdKeyATMDocMain = SecurPack.IdKeyATMDocMain
    f_IdKeyATM = SecurPack.IdKeyATM

    f_ColvoKey = SecurPack.ColvoKey  && Кол-во ключей в связке

    Pr_DtLost = IIF(SecurPack.Flg2PIN = 4, .T., .F.)

    m.IdTsd = SecurPack.IdTsd
    f_IdTsd = SecurPack.IdTsd

* WAIT WINDOW 'f_IdPackDefect = ' + ALLTRIM(STR(f_IdPackDefect)) TIMEOUT 3


* ------------------------------------- Проверяем имеются ли уже данные по конкретному номеру маршрута, инкассировался ли он ранее ------------------ *

*    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*      SELECT DISTINCT IdShift, IdRoute, ?CurFile AS FileName
*      FROM SecurPack (READPAST)
*      WHERE IdShift = ?f_IdShift AND IdRoute = ?f_IdRoute
*      UNION
*      SELECT DISTINCT IdShift, IdRoute, FileName
*      FROM SecurPack (READPAST)
*      WHERE IdShift = ?f_IdShift AND IdRoute = ?f_IdRoute

*    ENDTEXT

* ----------------------------------------------------------------------------------------------------- *

    TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
      EXECUTE Formir_Count_SecurPack ?f_IdShift, ?f_IdRoute, ?CurFile
    ENDTEXT

* ----------------------------------------------------------------------------------------------------- *

    IF RunSQL_Load('Count_SecurPack', @ar)  <>  1
      MESSAGEBOX('Таблица - Count_SecurPack, необходимая для работы НЕ СФОРМИРОВАНА', 16, 'Операция движения секьюрпаков')
    ENDIF

* BROWSE WINDOW brows TITLE 'RunSQL Count_SecurPack'

    IF USED('Count_SecurPack') = .T.

      SELECT Count_SecurPack
      GO TOP

      colvo_IdRoute = RECCOUNT()
      colvo_IdRoute = IIF(colvo_IdRoute = 0, 1, colvo_IdRoute)

      USE

    ELSE
      colvo_IdRoute = 15
    ENDIF

* WAIT WINDOW 'colvo_IdRoute = ' + ALLTRIM(STR(colvo_IdRoute)) TIMEOUT 3

    IF USED('Count_SecurPack') = .T.
      USE IN SELECT('Count_SecurPack')
    ENDIF

* ---------------------------------------------------------------------------------------------------------------------------- Запись в Базу -------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW 'f_IdTypeValue = ' + ALLTRIM(STR(f_IdTypeValue)) TIMEOUT 3

* --------------------------------------------------- Начало транзакции -------------------------------------------------------------------- *

    sqlReq = 'BEGIN TRANSACTION'

    rc = RunSQL_Load('')

* WAIT WINDOW 'BEGIN TRANSACTION = ' + ALLTRIM(STR(rc)) TIMEOUT 3

* ------------------------------------------------------------------------ Производим поиск в справочнике АТМ на возможную принадлежность ПИНа к АТМ или терминалу --------------------------------------------------------------------------- *

    DO CASE
      CASE Pr_DtLost = .F.

        DO start IN InsertKeyAtm  && Пишем данные в промежуточные таблицы KeyATMSecurPackMain и KeyATMSecurPackSub при штатной сдаче ключа, вместе с сейф пакетом или без него

      CASE Pr_DtLost = .T.

        DO start IN InsertKeyAtmLost  && Пишем данные в промежуточные таблицы KeyATMSecurPackMain и KeyATMSecurPackSub при вне штатной сдаче ключа, ключ утерян

    ENDCASE

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* WAIT WINDOW 'f_IdPackDefect = ' + ALLTRIM(STR(f_IdPackDefect)) TIMEOUT 3

*        TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*          INSERT INTO SecurPack
*          (IdShift, BudgetNum, IdUser, IdTsd, FilialCode,
*          TabNumber, OperCode, State, IdRoute, DateInkas, DtIncome, CurDate, FileName,
*          TotalSum, Pin2Sum, IdTypeValue, TypeValueKind, Sost, IsTransit, IsError, RecIdRoute, Flg2PIN,
*          WorkType,SuperBudget, SuperQty, IdPackDefect, IdJk, Pr_60322, TotalSum_60322, IdKeyATMDocMain, IdKeyATM, ColvoKey, Lait)
*          VALUES
*          (?f_IdShift, ?f_BudgetNum, ?f_IdControler, ?f_IdTsd, ?f_FilialCode,
*          ?f_TabNumber, ?f_OperCode, ?f_State, ?f_IdRoute, ?f_DateInkas, ?f_DtIncome, ?f_CurDate, ?CurFile,
*          ?f_TotalSum, ?f_Pin2Sum, ?f_IdTypeValue, ?f_TypeValueKind, ?f_Sost, ?f_IsTransit, 0, ?colvo_IdRoute, ?f_Flg2PIN,
*          ?f_WorkType, ?f_SuperBudget, ?f_SuperQty, ?f_IdPackDefect, ?f_IdJk, ?f_Pr_60322, ?f_TotalSum_60322, ?f_IdKeyATMDocMain, ?f_IdKeyATM, ?f_ColvoKey, ?Pr_Lait)

*        ENDTEXT

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*      TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

*          EXECUTE Formir_Insert_SecurPack_60322_KeyAtm_ColvoKey
*          ?f_IdShift, ?f_BudgetNum, ?f_IdControler, ?f_FilialCode, ?f_TabNumber, ?f_OperCode, ?f_State, ?f_IdRoute, ?f_DateInkas, ?f_DtIncome, ?f_CurDate, ?CurFile, ?f_TotalSum, ?f_Pin2Sum, ?f_IdTypeValue, ?f_TypeValueKind, ?f_Sost,
*          ?f_IsTransit, ?colvo_IdRoute, ?f_Flg2PIN, ?f_WorkType, ?f_SuperBudget, ?f_SuperQty, ?f_IdPackDefect, ?f_IdJk, ?f_Pr_60322, ?f_TotalSum_60322, ?f_IdKeyATMDocMain, ?f_IdKeyATM, ?f_ColvoKey

*      ENDTEXT

*      TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

*          EXECUTE Formir_Insert_SecurPack_KeyAtm_TSD
*          ?f_IdShift, ?f_BudgetNum, ?f_IdControler, ?f_IdTsd, ?f_FilialCode, ?f_TabNumber, ?f_OperCode, ?f_State, ?f_IdRoute, ?f_DateInkas, ?f_DtIncome, ?f_CurDate, ?CurFile, ?f_TotalSum, ?f_Pin2Sum, ?f_IdTypeValue, ?f_TypeValueKind, ?f_Sost,
*          ?f_IsTransit, ?colvo_IdRoute, ?f_Flg2PIN, ?f_WorkType, ?f_SuperBudget, ?f_SuperQty, ?f_IdPackDefect, ?f_IdJk, ?f_Pr_60322, ?f_TotalSum_60322, ?f_IdKeyATMDocMain, ?f_IdKeyATM, ?f_ColvoKey

*      ENDTEXT


    TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

        EXECUTE Formir_Insert_SecurPack_KeyAtm_TSD_Lait
        ?f_IdShift, ?f_BudgetNum, ?f_IdControler, ?f_IdTsd, ?f_FilialCode, ?f_TabNumber, ?f_OperCode, ?f_State, ?f_IdRoute, ?f_DateInkas, ?f_DtIncome, ?f_CurDate, ?CurFile, ?f_TotalSum, ?f_Pin2Sum, ?f_IdTypeValue, ?f_TypeValueKind, ?f_Sost,
        ?f_IsTransit, ?colvo_IdRoute, ?f_Flg2PIN, ?f_WorkType, ?f_SuperBudget, ?f_SuperQty, ?f_IdPackDefect, ?f_IdJk, ?f_Pr_60322, ?f_TotalSum_60322, ?f_IdKeyATMDocMain, ?f_IdKeyATM, ?f_ColvoKey, ?Pr_Lait

    ENDTEXT

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    rc = RunSQL_Load('SpId', @ar)

* WAIT WINDOW 'INSERT INTO SecurPack = ' + ALLTRIM(STR(rc)) TIMEOUT 5

    DO CASE
      CASE rc <> 1  && Вставка записи о новой введенной сумке прошла не удачно

        sqlReq = 'ROLLBACK TRANSACTION'  && Откатываем начатую транзакцию и пишем протокол  &&вынес из блока взаимодействия с пользователем Макаров КВ 11.09.2016

        rc = RunSQL_Load('')

        IF MESSAGEBOX('Продолжить обработку файла по маршруту - ' + ALLTRIM(STR(f_IdRoute)), 33, 'Ошибка на строке № ' + STR20(RECNO(),1)) = 1

          curmess = 'Отказ от обработки файла ' + ALLTRIM(CurFile) + ' из-за ошибок при вводе данны в базу'
          ret = WriteLog(curmess, 'L', PROGRAM())

          LOOP  && Уходим в начало цикла и продолжаем работать со следующей записью

        ENDIF

        DO exit_sql_server IN RunSQL_Load
        curmess = 'Отказ от обработки файла ' + ALLTRIM(CurFile) + ' из-за ошибок при вводе данны в базу'
        ret = WriteLog(curmess, 'L', PROGRAM())
        MESSAGEBOX(curmess, 16, 'Операция движения секьюрпаков')
        RETURN -1  && Выходим аварийно из модуля записи

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

      CASE rc = 1  && Вставка записи о новой введенной сумке прошла удачно


*            BROWSE WINDOW brows



* ----------------------------------------------------- Дата 1 й инкассации (пересчета) для точки.  (Filial.DtBeginPer) ----------------------------------------------------------------------------------------- *

* WAIT WINDOW '1 - Дата 1 й инкассации (пересчета) для точки'

        IF ISNULL(f_DtBeginPer) = .F. OR EMPTY(f_DtBeginPer) = .F.

* WAIT WINDOW '2 - Дата 1 й инкассации (пересчета) для точки'

*            TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*              UPDATE Filial SET
*              DtBeginPer = ?f_DtBeginPer
*              WHERE Id = ?f_IdFilial

*            ENDTEXT

* ----------------------------------------------------------------------------------------------------- *

          TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
            EXECUTE Formir_Filial_DtBeginPer ?f_IdFilial, ?f_DtBeginPer
          ENDTEXT

* ----------------------------------------------------------------------------------------------------- *

          msg = ''

          rc = RunSQL_Load('', @ar)

        ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*            TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*              SELECT Id, IdShift, BudgetNum, IdFilial, FilialCode, IdRoute, TotalSum, TypeValueKind, RecIdRoute, WorkType, SuperBudget, SuperQty, IdPackDefect
*              FROM SecurPack
*              WHERE IdShift = ?f_IdShift AND BudgetNum = ?f_BudgetNum AND TypeValueKind = ?f_TypeValueKind

*            ENDTEXT

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
          EXECUTE Formir_SecurPack_BudgetNum_TypeValueKind ?f_IdShift, ?f_BudgetNum, ?f_TypeValueKind
        ENDTEXT

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------- *

        rc = RunSQL_Load('Insert_SecurPack', @ar)

        IF rc = 1  && Выборка записи о новой введенной сумке прошла удачно

          IF USED('Insert_SecurPack') = .T. AND RECCOUNT() = 1

*            BROWSE WINDOW brows TITLE 'Insert_SecurPack'

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------- *

            f_IdSecurPack = Insert_SecurPack.Id  && Id Только что вставленной записи

* WAIT WINDOW 'f_IdSecurPack = ' + ALLTRIM(STR(f_IdSecurPack)) TIMEOUT 5
* WAIT WINDOW 'f_BudgetNum = ' + ALLTRIM(f_BudgetNum) + '   f_IdTsd = ' + ALLTRIM(STR(f_IdTsd)) &&TIMEOUT 5

            IF USED('TSDKasperPRK') = .T.

              SELECT TSDKasperPRK
              GO TOP
              LOCATE FOR Id = f_IdTsd

*              BROWSE WINDOW brows

* WAIT WINDOW 'FOUND = ' + IIF(FOUND() = .T., '.T.', '.F.') TIMEOUT 5

              IF FOUND() = .T.
                REPLACE IdSecurPack WITH f_IdSecurPack
              ENDIF

*                BROWSE WINDOW brows FOR Id = f_IdTsd

            ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

* --------------------------------------------------------------------------------- Здесь запись дефектов ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ *

            IF f_IdPackDefect > 0  && Дефект по введенной сумке зафиксирован

              SELECT SecurPackDefect
              GO TOP

              LOCATE FOR ALLTRIM(BudgetNum) == ALLTRIM(f_BudgetNum) AND IdTypeValue = f_IdTypeValue AND IdShift = f_IdShift

              IF FOUND('SecurPackDefect') = .T.  && Пока мы внутри  "IF f_IdPackDefect > 0  && Дефект по введенной сумке зафиксирован" - это будет всегда .Т.

                f_IdSecurPack = Insert_SecurPack.Id  && Перенёс инициализацию пром.переменных после LOCATE внутрь IF FOUND Минус промежуточных переменных - по ошибке проинициализировали раньше, чем вышли на нужную запись
                f_IdBudgetDefect = Insert_SecurPack.IdPackDefect && такой вид ошибки иногда сложно обнаружить.
                f_NumPacketsByDocs = SecurPackDefect.NumPacketsByDocs
                f_TotalSum = SecurPackDefect.TotalSum
                f_BudgetNumReal = SecurPackDefect.BudgetNumReal
                f_PlombaNumAct = SecurPackDefect.PlombaNumAct
                f_PlombaNumDocs = SecurPackDefect.PlombaNumDocs
                f_PacketsNumbers = SecurPackDefect.PacketsNumbers

*                  TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

*                    INSERT INTO SecurPackDefect (IdSecurPack, IdBudgetDefect, NumPacketsByDocs, TotalSumAct, BudgetNumReal, PlombaNumReal, PlombaNumByDocs, PacketsNumbers)
*                    VALUES (?f_IdSecurPack, ?f_IdBudgetDefect, ?f_NumPacketsByDocs, ?f_TotalSum, ?f_BudgetNumReal, ?f_PlombaNumAct, ?f_PlombaNumDocs, ?f_PacketsNumbers)

*                  ENDTEXT

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
                  EXECUTE Formir_Insert_SecurPack_Defect_New ?f_IdSecurPack, ?f_IdBudgetDefect, ?f_NumPacketsByDocs, ?f_TotalSum, ?f_BudgetNumReal, ?f_PlombaNumAct, ?f_PlombaNumDocs, ?f_PacketsNumbers
                ENDTEXT

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*                TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
*                  EXECUTE Formir_Insert_SecurPack_Defect ?f_IdSecurPack, ?f_IdBudgetDefect
*                ENDTEXT

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------- *

                rc = RunSQL_Load('', @ar)

              ENDIF

              SELECT MisLog_Cursor
              GO TOP

* BROWSE WINDOW brows

              SCAN FOR ALLTRIM(BudgetNum) == ALLTRIM(f_BudgetNum) AND IdTypeValue = f_IdTypeValue AND IdShift = f_IdShift

                f_IdSecurPack = Insert_SecurPack.Id
                f_IdValueable = MisLog_Cursor.IdValueable
                f_ValueableName = MisLog_Cursor.ValueableName
                f_Year = MisLog_Cursor.Year
                f_SomeData = MisLog_Cursor.SomeData
                f_SomeData2 = MisLog_Cursor.SomeData2
                f_DefType = MisLog_Cursor.DefType

                TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

                  INSERT INTO MisLogPRK
                  (IdSecurPack, IdValueable, ValueableName, Year, SomeData, DefType,SomeData2)
                  VALUES
                  (?f_IdSecurPack, ?f_IdValueable, ?f_ValueableName, ?f_Year, ?f_SomeData, ?f_DefType,?f_SomeData2)

                ENDTEXT

                rc = RunSQL_Load('', @ar)

              ENDSCAN

*----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

              IF rc <> 1  && Вставка записи о дефекте по новой введенной сумке прошла не удачно

                sqlReq = 'ROLLBACK TRANSACTION'  && Откатываем начатую транзакцию и пишем протокол
                rc = RunSQL_Load('')

                IF MESSAGEBOX('Продолжить обработку файла по маршруту - ' + ALLTRIM(STR(f_IdRoute)), 33, 'Ошибка на строке № ' + STR20(RECNO(), 1)) = 1

                  curmess = 'Отказ от обработки файла ' + ALLTRIM(CurFile) + ' из-за ошибок при вводе данны в базу'
                  ret = WriteLog(curmess, 'L', PROGRAM())

                  LOOP  && Уходим в начало цикла и продолжаем работать со следующей записью

                ENDIF

                DO exit_sql_server IN RunSQL_Load

                curmess = 'Отказ от обработки файла ' + ALLTRIM(CurFile) + ' из-за ошибок при вводе данны в базу'
                ret = WriteLog(curmess, 'L', PROGRAM())
                DEACTIVATE WINDOW poisk
                MESSAGEBOX(curmess, 16, 'Операция движения секьюрпаков')
                RETURN -1  && Выходим аварийно из модуля записи

              ENDIF

            ENDIF

          ENDIF

        ENDIF

        IF USED('Securpack_Empty')

          PRIVATE sal

          sal=ALIAS()

          SELECT Securpack_Empty

          SCAN ALL FOR NVL(Securpack_Empty.IdFilial,0)<>0      &&  пользователь мог завести пустую запись "в автомате" надо их пропускать.

            TEXT TO SqlReq NOSHOW TEXTMERGE

              INSERT INTO Securpack_Empty
              (IdFilial,BudgetNum,IdShift,IdRoute,IdUser,IdCashMessenger,dtrecord,IdCartridgeReturnReason)
              VALUES
              (?Securpack_Empty.IdFilial,?Securpack_Empty.BudgetNum,?Securpack_Empty.IdShift,?Securpack_Empty.IdRoute,?Securpack_Empty.IdUser,?Securpack_Empty.IdCashMessenger,GETDATE(),?Securpack_Empty.IdCartridgeReturnReason)

            ENDTEXT

            rc = RunSQL_Load('', @ar)

            IF rc<>1
              ASSERT .F.
            ENDIF

          ENDSCAN

          USE IN Securpack_Empty

          IF !EMPTY(Sal)
            SELECT (sal)
          ENDIF

        ENDIF

* ----------------------------------------------- Если ВСЕ операции по вводу параметров новой сумки прошли УСПЕШНО, то фиксируем транзакцию --------------------------------------------------- *

        sqlReq = 'COMMIT TRANSACTION' && Завершение транзакции

        rc = RunSQL_Load('')

* WAIT WINDOW 'COMMIT TRANSACTION = ' + ALLTRIM(STR(rc)) TIMEOUT 2

        numRowF = numRowF  +  1

    ENDCASE

    SELECT SecurPack
  ENDSCAN

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF USED('TSDKasperPRK') = .T.

    SELECT Id,;
      IdSecurPack,;
      IdShift,;
      Branch,;
      FilialCode,;
      FilialName,;
      IdRoute,;
      InkassGuid,;
      DateInkas,;
      BudgetNum,;
      WorkType,;
      Nominal,;
      TotalSum,;
      TypeValueKind,;
      QRCode,;
      InkassType,;
      FlgErr,;
      AccountQR,;
      AccountURK,;
      BuhSymbolQR,;
      BuhSymbolURK,;
      DtRecord ;
      FROM TSDKasperPRK ;
      WHERE IdShift = f_IdShift AND IdRoute = f_IdRoute AND IdSecurPack > 0 ;
      INTO CURSOR Scan_TSDKasperPRK ;
      ORDER BY IdSecurPack

    SELECT Scan_TSDKasperPRK

* BROWSE WINDOW brows

    IF USED('Scan_TSDKasperPRK') = .T. AND RECCOUNT() > 0

      SELECT Scan_TSDKasperPRK
      GO TOP

* BROWSE WINDOW brows

      SCAN

        sql_Id = Scan_TSDKasperPRK.Id
        sql_IdSecurPack = Scan_TSDKasperPRK.IdSecurPack
        sql_IdShift = Scan_TSDKasperPRK.IdShift
        sql_Branch = Scan_TSDKasperPRK.Branch
        sql_FilialCode = Scan_TSDKasperPRK.FilialCode
        sql_IdRoute = Scan_TSDKasperPRK.IdRoute
        sql_InkassGuid = Scan_TSDKasperPRK.InkassGuid
        sql_DateInkas = Scan_TSDKasperPRK.DateInkas
        sql_BudgetNum = Scan_TSDKasperPRK.BudgetNum
        sql_WorkType = Scan_TSDKasperPRK.WorkType
        sql_Nominal = Scan_TSDKasperPRK.Nominal
        sql_TotalSum = Scan_TSDKasperPRK.TotalSum
        sql_TypeValueKind = Scan_TSDKasperPRK.TypeValueKind
        sql_QRCode = Scan_TSDKasperPRK.QRCode
        sql_InkassType = Scan_TSDKasperPRK.InkassType
        sql_FlgErr = Scan_TSDKasperPRK.FlgErr
        sql_AccountQR = Scan_TSDKasperPRK.AccountQR
        sql_AccountURK = Scan_TSDKasperPRK.AccountURK
        sql_BuhSymbolQR = Scan_TSDKasperPRK.BuhSymbolQR
        sql_BuhSymbolURK = Scan_TSDKasperPRK.BuhSymbolURK
        sql_DtRecord = Scan_TSDKasperPRK.DtRecord

        TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

          IF NOT EXISTS(SELECT Id FROM TSDKasperPRK WHERE Id = ?sql_Id AND IdSecurPack = ?sql_IdSecurPack)

            INSERT INTO TSDKasperPRK
            (Id, IdSecurPack, IdShift, Branch, FilialCode, IdRoute, InkassGuid, DateInkas, BudgetNum, WorkType, Nominal, TotalSum, TypeValueKind, QRCode, InkassType, FlgErr, AccountQR, AccountURK, BuhSymbolQR, BuhSymbolURK, DtRecord)
            VALUES
            (?sql_Id, ?sql_IdSecurPack, ?sql_IdShift, ?sql_Branch, ?sql_FilialCode, ?sql_IdRoute, ?sql_InkassGuid, ?sql_DateInkas, ?sql_BudgetNum, ?sql_WorkType, ?sql_Nominal, ?sql_TotalSum, ?sql_TypeValueKind, ?sql_QRCode, ?sql_InkassType, ?sql_FlgErr, ?sql_AccountQR, ?sql_AccountURK, ?sql_BuhSymbolQR, ?sql_BuhSymbolURK, ?sql_DtRecord)

        ENDTEXT

        rc = RunSQL_Load('', @ar)

        SELECT Scan_TSDKasperPRK
      ENDSCAN

      USE IN SELECT('Scan_TSDKasperPRK')

    ENDIF

  ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF USED('Jk_Empty') = .T. AND USED('SpId') = .T. AND SpId.Id > 0 AND RECCOUNT('Jk_Empty') > 0

* Здесь пишем пустые порожние сумки в Jk_History
* 11-отметка ПРК пустая Явочная.

    f_IdCashMes = CashMessenger.Id
    f_SpId = SpId.Id  && Пурвый введеный

    TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7

      INSERT INTO Jk_History
      (IdJk, Op,IdUser, IdCashMess, IdSecurPack) VALUES

    ENDTEXT

    SELECT Jk_Empty

    SCAN

      TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7 ADDITIVE
        (<<bar>>, 11, <<f_IdControler>>, <<f_IdCashMes>>, <<f_SpId>> ),
      ENDTEXT

    ENDSCAN

    SqlReq = SUBSTR(SqlReq, 1, LEN(SqlReq) - 1) + ';'

    rc = RunSQL_Load('')

  ENDIF

  USE IN SELECT('SpId')
  USE IN SELECT('Jk')
  USE IN SELECT('Jk_Empty')
  USE IN SELECT('Prn_Jk')

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  tim2 = SECONDS()

  result_time = ROUND(tim2 - tim1, 2)

  @ WROWS()/3,3 SAY PADC('Запись данных в SQL Server успешно завершена.' + ' (время = ' + ALLTRIM(TRANSFORM(result_time, '999.99')) + ' сек.)',WCOLS())
  =INKEY(3)
  DEACTIVATE WINDOW poisk

  WAIT CLEAR

* --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  IF numRow <> numRowF

    curmess = 'Количество строк в файле - ' + ALLTRIM(CurFile) + ' - ' + STR20(numRow, 1) + CHR(13) + 'Не совпадает с введенными в базу = ' +  STR20(numRowF, 1)

    ret = WriteLog(curmess, 'L', PROGRAM())

    MESSAGEBOX(curmess, 64, 'Операция завершилась с ошибками')

  ELSE

    curmess = 'Операция прошла успешно.' + CHR(13) + 'Количество обработанных строк = ' + STR20(numRow, 1) + CHR(13) + 'Файл - ' + ALLTRIM(CurFile)

    ret = WriteLog(curmess, 'L', PROGRAM())

    MESSAGEBOX('Операция прошла успешно.' + CHR(13) + 'Количество обработанных строк = ' + STR20(numRow, 1), 48, 'Успешное выполнение операции')

  ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*** UPDATE № 3 Пометить сумки сдаваемые в ПРК как 'ИСПОЛЬЗОВАННЫЕ' на складе (SecurSklad) полученные от клиента с деньгами ***

*** dtToKl = ?f_DateInkas - поле дата передачи клиенту
*** OperCode = 3 - код операции (3-ИСПОЛЬЗОВАННЫЕ)

*    sqlReq = 'BEGIN TRANSACTION' && Начало транзакции

*    rc = RunSQL_Load('')

*    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

*      UPDATE SecurSklad SET
*      dtFromKl = ?f_DateInkas,
*      OperCode = 3
*      FROM SecurPack AS sp (READPAST)
*      WHERE sp.IdShift = ?f_IdShift AND sp.BudgetNum = SecurSklad.BudgetNum AND sp.Sost = 0 AND
*      sp.OperCode = 1 AND sp.FilialCode <> ?f_skldCode  AND FileName = ?CurFile

*    ENDTEXT

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*    TEXT TO SqlReq NOSHOW TEXTMERGE PRETEXT 7
*      EXECUTE Formir_SecurSklad ?f_IdShift, ?f_DateInkas, ?f_skldCode, ?CurFile
*    ENDTEXT

*    rc = RunSQL_Load('', @ar)

*    DO CASE
*      CASE rc <> 1  && Выполнение команды на стороне SQL Server прошло АВАРИЙНО

*        sqlReq = 'ROLLBACK TRANSACTION'  && Откат транзакции  && Вынес сюда, дабы не происходила блокировка, откат происходил в обеих ветках
*        rc = RunSQL_Load('')

*        IF MESSAGEBOX('Продолжить обработку файла по маршруту ' + ALLTRIM(STR(f_IdRoute)), 33, 'Ошибка записи признака "Выдано со склада"') = 2

*          DO exit_sql_server IN RunSQL_Load

*          curmess = 'Отказ от обработки файла ' + ALLTRIM(CurFile) + ' из-за ошибок при вводе данны в базу'

*          ret = WriteLog(curmess, 'L', PROGRAM())

*          MESSAGEBOX(curmess, 16, 'Операция движения секьюрпаков')

*          RETURN -1

*        ENDIF


*      CASE rc = 1  && Выполнение команды на стороне SQL Server прошло ШТАТНО

*        sqlReq = 'COMMIT TRANSACTION' && Завершение транзакции

*        rc = RunSQL_Load('')

*    ENDCASE

* ---------------------------------------------------------------- *
  DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

&& Отсюда два выхода - или принимать ведомость или делать что-либо ещё (в т.ч. заново запустить ввод пакетов). В любом случае обработку пакетов на данном
&& этапе закончили. Я закрываю курсор SecurPackDefect  - Макаров К.В.

  IF USED('SecurPackDefect') = .T.

    USE IN SecurPackDefect

  ENDIF

  IF MESSAGEBOX('Вы действительно закончили оформлять ведомость?', 33, 'Ручной ввод секьюрпаков в ПРК') <> 1
    RETURN
  ENDIF

* ----------------------------------------------------- *
  DO start_potok IN TakeVedom
* ----------------------------------------------------- *

ENDIF

SELECT(sel)
RETURN


******************************************************************************************************************************************************************************



