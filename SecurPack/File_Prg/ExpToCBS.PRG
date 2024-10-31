******************************************************************************* PROCEDURE start ****************************************************************

PROCEDURE start
PARAMETERS TypeBox, TypeFlg2Pin

* BudgetKP->WorkType
* TypeBox - Тип обрабатываемой сумки (0-просто сумка, 1-супер сумка, 2-вложенная сумка)

IF PARAMETERS() = 0
  TypeBox = 0
ENDIF

IF PARAMETERS() = 0
  TypeFlg2Pin = 1
ENDIF

pr_TypeBox = TypeBox
pr_TypeFlg2Pin = TypeFlg2Pin

sel = SELECT()
popup_ima = LOWER(POPUP())
prompt_ima = LOWER(PROMPT())
bar_num = BAR()

HIDE POPUP (popup_ima)

HIDE POPUP oper_cbs_One
HIDE POPUP oper_cbs

HIDE POPUP oper_cbs_Sup_Vip
HIDE POPUP oper_cbs_Sup

PRIVATE msg, selNumber, selNumber1, sqlLS1

DIMENSION ar[1,2]

selNumber = 0

pr_error = .T.

f_State = 0

* f_IdShift_2 = 3577

* WAIT WINDOW 'f_IdShift_1 = ' + ALLTRIM(STR(f_IdShift_1)) + '   f_IdShift_2 = ' + ALLTRIM(STR(f_IdShift_2)) TIMEOUT 3

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

LOCAL _i,_str

_i=Select_LstBudget()

IF _i!=0

  _str=GetOpenedAliases()        && её бы в обработчик ошибок поставить

  MESSAGEBOX('При подготовке данных к передаче на CBS произошла ошибка №'+transform(-_i)+CHR(13)+'Запишите, пожалуйста, номер ошибки и передайте программистам/дежурной техподдержке.',16)
  STRTOFILE(TIME() + GETENV("USERNAME")+'   ---  Error in export CBS'+CHR(13)+CHR(10)+_str+CHR(13)+CHR(10),'aliases.log',1)

  =CloseDB()
  RETURN

ENDIF

* ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* -------------------------------------------------- *

IF pr_error = .T.
  DO FORM ExpToCBS
ENDIF

* -------------------------------------------------- *

=CloseDB()

SELECT(sel)
RETURN


********************************************************************************************* PROCEDURE Select_LstBudget ********************************************************

PROCEDURE Select_LstBudget

* ---------------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

DO CASE
  CASE pr_TypeBox = 0 AND pr_TypeFlg2Pin = 1  && Выборка данных по обычным сумкам

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT LTRIM(RTRIM(STR(cfb.IdClnInk))) + ' - ' + LEFT(RTRIM(cfb.ClientMaxGroup), 50) AS ClientName, COUNT(bkp.Id) AS Colvo_pak, CAST(0 AS bit) AS Fl, cfb.IdClnInk AS IdClient
      FROM View_ClientGroupMax_Filial_Bank AS cfb (READPAST)
      INNER JOIN BudgetKP AS bkp (READPAST) ON bkp.IdFilial = cfb.IdFilial
      WHERE bkp.IdShift = ?f_IdShift_2  AND bkp.Worktype = '0'  AND bkp.State = 0 AND (bkp.SuperBudget = 0 OR bkp.SuperBudget IS NULL)
      GROUP BY cfb.IdClnInk, cfb.ClientMaxGroup
      ORDER BY cfb.IdClnInk

    ENDTEXT


  CASE pr_TypeBox = 0 AND pr_TypeFlg2Pin = 3  && Выборка данных по многопиновым сумкам

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT LTRIM(RTRIM(STR(cfb.IdClnInk))) + ' - ' + LEFT(RTRIM(cfb.ClientMaxGroup), 50) AS ClientName, COUNT(bkp.Id) AS Colvo_pak, CAST(0 AS bit) AS Fl, cfb.IdClnInk AS IdClient
      FROM View_ClientGroupMax_Filial_Bank AS cfb (READPAST)
      INNER JOIN BudgetKP AS bkp (READPAST) ON bkp.IdFilial = cfb.IdFilial
      WHERE bkp.IdShift = ?f_IdShift_2 AND bkp.Worktype = '0' AND bkp.State = 0 AND bkp.SuperBudget > 0 AND bkp.SuperBudget = bkp.IdBudgetHistoryPRK
      GROUP BY cfb.IdClnInk, cfb.ClientMaxGroup
      ORDER BY cfb.IdClnInk

    ENDTEXT


  CASE pr_TypeBox = 1 AND pr_TypeFlg2Pin = 1  && Выборка данных по супер сумкам

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT LTRIM(RTRIM(STR(cfb.IdClnInk))) + ' - ' + LEFT(RTRIM(cfb.ClientMaxGroup), 50) AS ClientName, COUNT(bkp.Id) AS Colvo_pak, CAST(0 AS bit) AS Fl, cfb.IdClnInk AS IdClient
      FROM View_ClientGroupMax_Filial_Bank AS cfb (READPAST)
      INNER JOIN BudgetKP AS bkp (READPAST) ON bkp.IdFilial = cfb.IdFilial
      WHERE bkp.IdShift = ?f_IdShift_2 AND bkp.Worktype = '1' AND bkp.State IN (0, 1) AND bkp.SuperQty >= 1
      GROUP BY cfb.IdClnInk, cfb.ClientMaxGroup
      ORDER BY cfb.IdClnInk

    ENDTEXT

ENDCASE

IF RunSQL_Load('GroupBudgetKP', @ar) <> 1
  DO exit_sql_server IN RunSQL_Load
*  =CloseDB()
  RETURN -1
ENDIF

* BROWSE WINDOW brows TITLE 'RunSQL GroupBudgetKP'

* ---------------------------------------------------------------- *
DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

* ------------------------------------- *
DO FORM Sprav_cln
* ------------------------------------- *

* Собираем выбранные сумки

f_IdClient = ''

SELECT GroupBudgetKP

SCAN FOR GroupBudgetKP.Fl = .T.
  f_IdClient = f_IdClient + ALLTRIM(STR(GroupBudgetKP.IdClient)) + [,]
ENDSCAN

f_IdClient = f_IdClient + ['']  && Полный список выбранных оператором сумок

* WAIT WINDOW 'f_IdClient = ' + ALLTRIM(f_IdClient) TIMEOUT 3
* WAIT WINDOW 'pr_TypeBox = ' + ALLTRIM(STR(pr_TypeBox)) TIMEOUT 3
* WAIT WINDOW 'f_State = ' + ALLTRIM(STR(f_State)) TIMEOUT 3
* WAIT WINDOW 'f_IdShift_2 = ' + ALLTRIM(STR(f_IdShift_2)) TIMEOUT 3

* ---------------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

DO CASE
  CASE pr_TypeBox = 0 AND pr_TypeFlg2Pin = 1  && Выборка данных по обычным сумкам

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT CAST(0 AS bit) AS Fl, cfb.FilialCode, bkp.BudgetNum, bkp.TotalSumPRK, bkp.State, bkp.WorkType, cfb.IdFilial, bkp.Id
      FROM View_ClientGroupMax_Filial_Bank AS cfb (READPAST)
      INNER JOIN BudgetKP AS bkp (READPAST) ON bkp.IdFilial = cfb.IdFilial
      WHERE bkp.IdShift = ?f_IdShift_2 AND bkp.Worktype = '0' AND bkp.State = 0 AND (bkp.SuperBudget = 0 OR bkp.SuperBudget IS NULL) AND cfb.IdClnInk IN (<<f_IdClient>>)
      ORDER BY cfb.FilialCode, bkp.BudgetNum

    ENDTEXT


  CASE pr_TypeBox = 0 AND pr_TypeFlg2Pin = 3  && Выборка данных по многопиновым сумкам

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT CAST(0 AS bit) AS Fl, cfb.FilialCode, bkp.BudgetNum, (SELECT SUM(TotalSumPRK) AS TotalSumPRK FROM BudgetKP WHERE bkp.IdBudgetHistoryPRK = SuperBudget) AS TotalSumPRK, bkp.State, bkp.WorkType, cfb.IdFilial, bkp.Id
      FROM View_ClientGroupMax_Filial_Bank AS cfb (READPAST)
      INNER JOIN BudgetKP AS bkp (READPAST) ON bkp.IdFilial = cfb.IdFilial
      WHERE bkp.IdShift = ?f_IdShift_2 AND bkp.Worktype = '0' AND bkp.State = 0 AND bkp.SuperBudget > 0 AND bkp.SuperBudget = bkp.IdBudgetHistoryPRK AND cfb.IdClnInk IN (<<f_IdClient>>)
      ORDER BY cfb.FilialCode, bkp.BudgetNum

    ENDTEXT


  CASE pr_TypeBox = 1 AND pr_TypeFlg2Pin = 1  && Выборка данных по супер сумкам

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT CAST(0 AS bit) AS Fl, cfb.FilialCode, bkp.BudgetNum, bkp.TotalSumPRK, bkp.State, bkp.WorkType, bkp.SuperBudget, bkp.SuperQty, cfb.IdFilial, bkp.Id
      FROM View_ClientGroupMax_Filial_Bank AS cfb (READPAST)
      INNER JOIN BudgetKP AS bkp (READPAST) ON bkp.IdFilial = cfb.IdFilial
      WHERE bkp.IdShift = ?f_IdShift_2 AND bkp.Worktype = '1' AND bkp.State IN (0, 1) AND cfb.IdClnInk IN (<<f_IdClient>>)
      ORDER BY cfb.FilialCode, bkp.BudgetNum

    ENDTEXT

ENDCASE

IF RunSQL_Load('LstBudget', @ar) <> 1
  DO exit_sql_server IN RunSQL_Load
*  =CloseDB()
  RETURN -1
ENDIF

SELECT LstBudget

colvo_zap = RECCOUNT()

GO TOP

* BROWSE WINDOW brows TITLE 'RunSQL LstBudget'

* ---------------------------------------------------------------- *
DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

RETURN 0


******************************************************************************** PROCEDURE LstBudget_Refresh ********************************************************

PROCEDURE LstBudget_Refresh

* WAIT WINDOW 'f_IdClient = ' + ALLTRIM(f_IdClient) TIMEOUT 3

DIMENSION ar[1,2]

* f_SuperBudget = STR20(LstBudget.SuperBudget, 1)

* WAIT WINDOW 'pr_TypeBox = ' + ALLTRIM(STR(pr_TypeBox)) TIMEOUT 3
* WAIT WINDOW 'f_State = ' + ALLTRIM(STR(f_State)) TIMEOUT 3
* WAIT WINDOW 'f_IdShift_2 = ' + ALLTRIM(STR(f_IdShift_2)) TIMEOUT 3

* ---------------------------------------------------------------- *
DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

DO CASE
  CASE pr_TypeBox = 0 AND pr_TypeFlg2Pin = 1  && Выборка данных по обычным сумкам

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT CAST(0 AS bit) AS Fl, cfb.FilialCode, bkp.BudgetNum, bkp.TotalSumPRK, SPACE(30) AS State, bkp.State AS StateNum, bkp.WorkType, cfb.IdFilial, bkp.Id
      FROM View_ClientGroupMax_Filial_Bank AS cfb (READPAST)
      INNER JOIN BudgetKP AS bkp (READPAST) ON bkp.IdFilial = cfb.IdFilial
      WHERE bkp.IdShift = ?f_IdShift_2 AND bkp.Worktype = '0' AND bkp.State = 0 AND (bkp.SuperBudget = 0 OR bkp.SuperBudget IS NULL) AND cfb.IdClnInk IN (<<f_IdClient>>)
      ORDER BY cfb.FilialCode, bkp.BudgetNum

    ENDTEXT


  CASE pr_TypeBox = 0 AND pr_TypeFlg2Pin = 3  && Выборка данных по многопиновым сумкам

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT CAST(0 AS bit) AS Fl, cfb.FilialCode, bkp.BudgetNum, (SELECT SUM(TotalSumPRK) AS TotalSumPRK FROM BudgetKP WHERE bkp.IdBudgetHistoryPRK = SuperBudget) AS TotalSumPRK, bkp.State, bkp.WorkType, cfb.IdFilial, bkp.Id
      FROM View_ClientGroupMax_Filial_Bank AS cfb (READPAST)
      INNER JOIN BudgetKP AS bkp (READPAST) ON bkp.IdFilial = cfb.IdFilial
      WHERE bkp.IdShift = ?f_IdShift_2 AND bkp.Worktype = '0' AND bkp.State = 0 AND bkp.SuperBudget > 0 AND bkp.SuperBudget = bkp.IdBudgetHistoryPRK AND cfb.IdClnInk IN (<<f_IdClient>>)
      ORDER BY cfb.FilialCode, bkp.BudgetNum

    ENDTEXT


  CASE pr_TypeBox = 1 AND pr_TypeFlg2Pin = 1  && Выборка данных по супер сумкам

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      SELECT CAST(0 AS bit) AS Fl, cfb.FilialCode, bkp.BudgetNum, bkp.TotalSumPRK, SPACE(30) AS State, bkp.State AS StateNum, bkp.WorkType, bkp.SuperBudget, bkp.SuperQty, cfb.IdFilial, bkp.Id
      FROM View_ClientGroupMax_Filial_Bank AS cfb (READPAST)
      INNER JOIN BudgetKP AS bkp (READPAST) ON bkp.IdFilial = cfb.IdFilial
      WHERE bkp.IdShift = ?f_IdShift_2 AND bkp.Worktype = '1' AND bkp.State IN (0, 1) AND cfb.IdClnInk IN (<<f_IdClient>>)
      ORDER BY cfb.FilialCode, bkp.BudgetNum

    ENDTEXT

ENDCASE

IF RunSQL_Load('LstBudget', @ar) <> 1
  DO exit_sql_server IN RunSQL_Load
  =CloseDB()
  RETURN -1
ENDIF

colvo_zap = RECCOUNT()

* ---------------------------------------------------------------- *
DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

SELECT LstBudget

REPLACE ALL State WITH WorkStat[StateNum + 1]

GO TOP
* BROWSE WINDOW brows TITLE 'RunSQL LstBudget'

RETURN 0


********************************************************************************************* FUNCTION CloseDB ******************************************************************

FUNCTION CloseDB

IF USED('LstBudget') = .T.
  SELECT LstBudget
  USE
ENDIF

IF USED('LstBudget1') = .T.
  SELECT LstBudget1
  USE
ENDIF

IF USED('BudgetKP') = .T.
  SELECT BudgetKP
  USE
ENDIF

ENDFUNC


*************************************************************** FUNCTION GetOpenedAliases *******************************************************

FUNCTION GetOpenedAliases
LOCAL i,j,strret

j=select()
strret=''

FOR i=1 TO 100
  IF USED(i)
    strret=strret+ IIF(i==j,'*','#')+STR(i,3)+'   '+PADL(ALIAS(i),20)+'('+TRANSFORM(RECCOUNT(i))+','+TRANSFORM(recno(i))+')'+CHR(13)+CHR(10)
  ENDIF
NEXT

select j

RETURN strret
ENDFUNC


***********************************************************************************************************************************************************************************


