********************************************************************** PROCEDURE start *******************************************************

PROCEDURE start
sel = SELECT()

DIMENSION ar[1,2]

PRIVATE NewFileName, NewFileName_M5

NewFileName = ''
NewFileName_M5 = ''

SELECT Id ;
  FROM Export_Sumki ;
  WHERE Fl = .T. ;
  INTO ARRAY aTmp

* BROWSE WINDOW brows

SelNumber1 = IIF(_TALLY = 0, 0, _TALLY)

IF EMPTY(selNumber1) = .T.
  MESSAGEBOX('Нет выбранных сумок - операция отменена!', 16, 'Операция передачи сумок в CBS')
  RETURN
ENDIF

* ---------------------------------------------------------------------- Получаем путь к месту, где расположена папка обмена с локальной станцией --------------------------------------------------------- *

IF DIRECTORY(ALLTRIM(put_tmp) + (put_ObmenToCBS), 1) = .T.  && Проверяем доступность сетевой папки

* ------------------------------------------------- Новый вариант выгрузки файла для CBS по суперсумкам ---------------------------------------- *

  DesktopPath = ALLTRIM(put_tmp)

  m.rand = INT(99 * RAND(-1) + 1)
  m.rand = IIF(m.rand < 10, '0' + ALLTRIM(STR(m.rand)), ALLTRIM(STR(m.rand)))

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  DO CASE
    CASE ALLTRIM(put_ObmenToCBS) == 'ObmenToCBS\'

      TEXT TO NewFileName NOSHOW TEXTMERGE  PRETEXT 15

        ObmenToCBS\
        Sup_BudgetKp_<<PADL(ALLTRIM(STR(f_IdShift_2)), 7, '0')>>_
        <<m.rand>>_
        <<CHRTRANC(DTOC(DATE()), [.], [])>>_
        <<CHRTRANC(TIME(), [:], [])>>
        .XML

      ENDTEXT

*        IF sergeizhuk_flag = .T.

*          IF DIRECTORY(ALLTRIM(put_tmp) + (put_ObmenToGdasuc), 1) = .T.  && Проверяем доступность сетевой папки

*            TEXT TO NewFileName_M5 NOSHOW TEXTMERGE  PRETEXT 15

*              ObmenToGdasuc\
*              Sup_BudgetKp_<<PADL(ALLTRIM(STR(f_IdShift_2)), 7, '0')>>_
*              <<m.rand>>_
*              <<CHRTRANC(DTOC(DATE()), [.], [])>>_
*              <<CHRTRANC(TIME(), [:], [])>>
*              .XML

*            ENDTEXT

*          ENDIF

*        ENDIF

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

    CASE ALLTRIM(put_ObmenToCBS) == 'ObmenToCBSYAR\'

      TEXT TO NewFileName NOSHOW TEXTMERGE  PRETEXT 15

        ObmenToCBSYAR\
        Sup_BudgetKp_<<PADL(ALLTRIM(STR(f_IdShift_2)), 7, '0')>>_
        <<m.rand>>_
        <<CHRTRANC(DTOC(DATE()), [.], [])>>_
        <<CHRTRANC(TIME(), [:], [])>>
        .XML

      ENDTEXT

  ENDCASE

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  NewFileName = CHRTRANC(NewFileName, [ ], [])
  NewFileName = ALLTRIM(put_tmp) + ALLTRIM(NewFileName)
  NewFileName = ["] + NewFileName + ["]

* WAIT WINDOW 'NewFileName = ' + ALLTRIM(NewFileName) TIMEOUT 10

* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*    NewFileName_M5 = CHRTRANC(NewFileName_M5, [ ], [])
*    NewFileName_M5 = ALLTRIM(put_tmp) + ALLTRIM(NewFileName_M5)
*    NewFileName_M5 = ["] + NewFileName_M5 + ["]

* WAIT WINDOW 'NewFileName_M5 = ' + ALLTRIM(NewFileName_M5) TIMEOUT 10

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

  appdir = SYS(5) + CURDIR()

  IF DIRECTORY(ALLTRIM(put_tmp) + (put_ObmenToCBS), 1) = .T.  && Проверяем доступность сетевой папки

* Получаем путь где расположена папка обмена с локальной станцией.

    FilePath = ALLTRIM(put_tmp) + (put_ObmenToCBS)

    SET DEFAULT TO (FilePath)

    NumFiles = 0

    NumFiles = ADIR(WorkDir, 'Sup_BudgetKP_*.XML')

*    WAIT WINDOW 'NumFiles = ' + ALLTRIM(STR(NumFiles)) + '   WorkDir = ' + ALLTRIM(WorkDir[1, 1]) TIMEOUT 3

    IF NumFiles > 0

      MESSAGEBOX('Присутствуют необработанные файлы для экспорта данных в CBS' + CHR(13) + 'Количество найденных файлов = ' + ALLTRIM(STR(NumFiles)) + CHR(13) + CHR(13) + ;
        'Данные файлы будут перемещены в архив!!!' + CHR(13) + 'Операция создания нового файла будет продолжена', 16, 'Операция передачи сумок в CBS')

      PRIVATE X, FilesCBS_Error

      X = 0

      FOR X = 1 TO NumFiles

        dopmetka = ALLTRIM(SYS(2015))

        FilesCBS_Error = SUBSTR((dopmetka), 2, 9) + '_' + WorkDir[X, 1]

*        WAIT WINDOW 'NumFiles = ' + ALLTRIM(STR(NumFiles)) + '   WorkDir = ' + ALLTRIM(WorkDir[X, 1]) + '   FilesCBS_Error = ' + ALLTRIM(FilesCBS_Error) TIMEOUT 3

        IF FILE(ALLTRIM(FilePath) + WorkDir[X, 1]) = .T.
          COPY FILE ALLTRIM(FilePath) + WorkDir[X, 1] TO ALLTRIM(FilePath) + ('Error\') + ALLTRIM(FilesCBS_Error)
        ELSE
          MESSAGEBOX('Файл XML с данными для пересчета, по пути :' + CHR(13) + ALLTRIM(FilePath) + WorkDir[X, 1] + CHR(13) + 'для перемещения в архив НЕ ДОСТУПЕН', 16, 'Операция передачи сумок в CBS')
        ENDIF

        IF FILE(ALLTRIM(FilePath) + ('Error\') + ALLTRIM(FilesCBS_Error)) = .T.

          DELETE FILE ALLTRIM(FilePath) + WorkDir[X, 1]  && Удаляем ранее выгруженный файл

        ENDIF

      ENDFOR

    ENDIF

    SET DEFAULT TO (appdir)

  ENDIF

* ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* ---------------------------------------------------------------- *
  DO vxod_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

* Начало транзакции

  sqlReq = 'BEGIN TRANSACTION' && Начало транзакции

  =RunSQL_Load('', @ar)

  TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7
    DELETE XmlToCBS
  ENDTEXT

  sqlRet = RunSQL_Load('', @ar)

* WAIT WINDOW 'DELETE XmlToCBS - sqlRet = ' + ALLTRIM(STR(sqlRet)) TIMEOUT 2

  IF sqlRet = 1  && Удаление ранее добавленных записей выполнено УСПЕШНО

* Update суперСумки с проверкой её состаяния

    f_IdControler = IdControler_

* WAIT WINDOW 'f_Lstbudget_Id = ' + ALLTRIM(STR(f_Lstbudget_Id)) TIMEOUT 5

    TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

      UPDATE BudgetKP SET
      State = 1,
      SostCbs = 1,
      DtToCbs = GetDate(),
      UserToCbs = ?f_IdControler
      OUTPUT ?f_IdUser, 'АРМ Контролёра КП', 'Export_to_xml_sup.prg', 'Kasper->CBS', 'Изменение BudgetKP'
      , 'IdUser: ' + Convert(char(6), ISNULL(Deleted.IdUser, 0)) + ', SostCBS: ' + Convert(char(1), ISNULL(Deleted.SostCbs, 0)) + ', ToCBS: ' + ISNULL(Convert(char(20), Deleted.DtToCbs, 120), 'NULL') + ', FromCBS: ' + ISNULL(Convert(char(20), Deleted.DtFromCbs, 120), 'NULL') + ', TotalSumKP: ' + Convert(char(25), Isnull(Deleted.TotalSumKP, 0))
      , 'IdUser: ' + Convert(char(6), ISNULL(Inserted.IdUser, 0)) + ', SostCBS: ' + Convert(char(1), ISNULL(Inserted.SostCbs, 0)) + ', ToCBS: ' + ISNULL(Convert(char(20), Inserted.DtToCbs, 120), 'NULL') + ', FromCBS: ' + ISNULL(Convert(char(20), Inserted.DtFromCbs, 120), 'NULL') + ', TotalSumKP: ' + Convert(char(25), Isnull(Inserted.TotalSumKP, 0))
      , Inserted.Id, GetDate()
      INTO UserLog (IdUser, ARM, Program, TypeRecord, [Text], OldValue, NewValue, IdUpdRec, DtRecord)
      WHERE Id = ?f_Lstbudget_Id AND State IN (0, 1)

    ENDTEXT

    sqlRet = RunSQL_Load('', @ar)

* WAIT WINDOW '1 - UPDATE BudgetKP - sqlRet = ' + ALLTRIM(STR(sqlRet)) TIMEOUT 2

    IF sqlRet = 1

      TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

        INSERT INTO XmlToCBS
        (IdbudgetKP, IdSuperKP) VALUES (?f_Lstbudget_Id, ?f_SuperBudget)

      ENDTEXT

      sqlRet = RunSQL_Load('', @ar)

* WAIT WINDOW '1 - INSERT INTO XmlToCBS - sqlRet = ' + ALLTRIM(STR(sqlRet)) TIMEOUT 2

    ENDIF

* --------------------------------------------------------------- UPDATE отмеченных вложенных сумок с проверкой их состаяния ------------------------------------------------ *

    SELECT Export_Sumki
* BROWSE WINDOW brows

    SCAN FOR Export_Sumki.Fl = .T. AND sqlRet = 1

      f_Lstbudget_Id = Export_Sumki.Id
      f_SuperBudget = Export_Sumki.SuperBudget

* WAIT WINDOW 'f_Lstbudget_Id = ' + ALLTRIM(STR(f_Lstbudget_Id)) TIMEOUT 2

      TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

        UPDATE BudgetKP SET
        State = 1,
        SostCbs = 1,
        DtToCbs = GetDate(),
        UserToCbs = ?f_IdControler
        OUTPUT ?f_IdUser, 'АРМ Контролёра КП', 'Export_to_xml_sup.prg', 'Kasper->CBS', 'Изменение BudgetKP'
        , 'IdUser: ' + Convert(char(6), ISNULL(Deleted.IdUser, 0)) + ', SostCBS: ' + Convert(char(1), ISNULL(Deleted.SostCbs, 0)) + ', ToCBS: ' + ISNULL(Convert(char(20), Deleted.DtToCbs, 120), 'NULL') + ', FromCBS: ' + ISNULL(Convert(char(20), Deleted.DtFromCbs, 120), 'NULL') + ', TotalSumKP: ' + Convert(char(25), Isnull(Deleted.TotalSumKP, 0))
        , 'IdUser: ' + Convert(char(6), ISNULL(Inserted.IdUser, 0)) + ', SostCBS: ' + Convert(char(1), ISNULL(Inserted.SostCbs, 0)) + ', ToCBS: ' + ISNULL(Convert(char(20), Inserted.DtToCbs, 120), 'NULL') + ', FromCBS: ' + ISNULL(Convert(char(20), Inserted.DtFromCbs, 120), 'NULL') + ', TotalSumKP: ' + Convert(char(25), Isnull(Inserted.TotalSumKP, 0))
        , Inserted.Id, GetDate()
        INTO UserLog (IdUser, ARM, Program, TypeRecord, [Text], OldValue, NewValue, IdUpdRec, DtRecord)
        WHERE State IN (0, 1) AND Id = ?f_Lstbudget_Id

      ENDTEXT

      sqlRet = RunSQL_Load('', @ar)

* WAIT WINDOW '2 - UPDATE BudgetKP - sqlRet = ' + ALLTRIM(STR(sqlRet)) TIMEOUT 2

      IF ar[1,2] = 0 && Сумка уже обработана или удалена
        sqlRet = -2
        myMess = 'Сумка №' + ALLTRIM(budgetNum) + ' уже пересчитывается - Операция ОТМЕНЕНА!'
        EXIT
      ENDIF

* Сохранение значений Id во временной таблице

      TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

        INSERT INTO XmlToCBS
        (IdbudgetKP, IdSuperKP) VALUES (?f_Lstbudget_Id, ?f_SuperBudget)

      ENDTEXT

      sqlRet = RunSQL_Load('', @ar)

* WAIT WINDOW '2 - INSERT INTO XmlToCBS - sqlRet = ' + ALLTRIM(STR(sqlRet)) TIMEOUT 2

      SELECT Export_Sumki
    ENDSCAN

* ------------------------------------------------------------------------------------ Выборка для XML -------------------------------------------------------------------- *

    IF sqlRet = 1

* Id - уникальный номер записи
* TotalsumPRK - заявленная сумма (в препроводительной ведомости)
* BudgetNum - номер сумки
* Worktype = 0 - тип обрабатываемой сумки (0-просто сумка, 1-супер сумка, 2-вложенная сумка)
* SuperBudget - Id супер сумки для вложенных сумок
* Superqty - количество вложенных сумок для суперсумок (пакетов для сумок)
* CONVERT(Char(10), bkp.DtIncome, 104) AS DtIncome - новый выгружаемый параметр, дата инкассации, для супер сумок пока не применяется, реализован 21.08.2015
* fil.FilialName - новый выгружаемый параметр, реквизиты ПИНа (Наименование + ПИН + Адресные данные), для супер сумок пока не применяется, реализован 21.08.2015

* SELECT bkp.Id, CAST(bkp.TotalSumPRK AS Decimal(18, 2)) AS TotalSumPRK, bkp.BudgetNum, bkp.WorkType, bkp.SuperBudget, bkp.Superqty, CONVERT(Char(10), bkp.DtIncome, 104) AS DtIncome, fil.FilialName

      TEXT TO sqlReq NOSHOW TEXTMERGE PRETEXT 7

        SELECT bkp.Id, CAST(bkp.TotalSumPRK AS Decimal(18, 2)) AS TotalSumPRK, bkp.BudgetNum, bkp.WorkType, bkp.SuperBudget, bkp.Superqty
        FROM BudgetKP AS bkp (READPAST)
        INNER JOIN Filial AS fil (READPAST) ON bkp.IdFilial = fil.Id
        INNER JOIN [XmlToCBS] AS cbs (READPAST) ON cbs.IdbudgetKP = bkp.Id
        ORDER BY bkp.Id DESC, bkp.WorkType

      ENDTEXT

      sqlRet = RunSQL_Load('BudgetKP', @ar)

* WAIT WINDOW 'Выборка для XML - sqlRet = ' + ALLTRIM(STR(sqlRet)) TIMEOUT 2

      IF USED('BudgetKP') = .T.

        SELECT BudgetKP

        colvo_zap = RECCOUNT()

        GO TOP
* BROWSE WINDOW brows

        IF colvo_zap <> 0

* WAIT WINDOW 'put_tmp = ' + ALLTRIM(put_tmp) TIMEOUT 5
* WAIT WINDOW 'NewFileName = ' + ALLTRIM(NewFileName) TIMEOUT 5
* WAIT WINDOW 'NewFileName_M5 = ' + ALLTRIM(NewFileName_M5) TIMEOUT 5

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

          FOR I =1 TO 10

            =CURSORTOXML('BudgetKP', NewFileName, 1, 512)  && Преобразует Visual FoxPro курсор в XML

***      =CURSORTOXML('BudgetKP', NewFileName, 1, 512 + 16)  && Преобразует Visual FoxPro курсор в XML с явным указанием кодировки вывода. Вывод кодируется по кодовой странице курсора - INTO CURSOR BudgetKP. ***

*** Для гарании аккуратной трансляции символов, кодовая страница Visual FoxPro, используемая по умолчанию должна соответствовать кодовой странице курсора. ***
*** Вы можете обеспечить это поведение установкой для символьных и мемо полей аттрибута NOCPTRAN ( character binary/memo binary) ***

            IF FILE(NewFileName) = .T.
              EXIT
            ELSE
              LOOP
            ENDIF

          ENDFOR

          IF FILE(NewFileName) = .F.

            sqlReq = 'ROLLBACK TRANSACTION'
            sqlRet = RunSQL_Load('', @ar)

            IF sqlRet = 1
              MESSAGEBOX('Необходимый для работы файл по пути:' + CHR(13) + ALLTRIM(NewFileName) + ' - НЕ СОЗДАН!!!' + CHR(13) + 'Операция ОТМЕНЕНА!' + CHR(13) + 'Повторите процедуру формирования файла ...' , 16, 'Создание выходного файла для CBS')
            ENDIF

            RETURN

          ENDIF

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

*            FOR I =1 TO 10

*              =CURSORTOXML('BudgetKP', NewFileName_M5, 1, 512)  && Преобразует Visual FoxPro курсор в XML

*  ***      =CURSORTOXML('BudgetKP', NewFileName_M5, 1, 512 + 16)  && Преобразует Visual FoxPro курсор в XML с явным указанием кодировки вывода. Вывод кодируется по кодовой странице курсора - INTO CURSOR BudgetKP. ***

*  *** Для гарании аккуратной трансляции символов, кодовая страница Visual FoxPro, используемая по умолчанию должна соответствовать кодовой странице курсора. ***
*  *** Вы можете обеспечить это поведение установкой для символьных и мемо полей аттрибута NOCPTRAN ( character binary/memo binary) ***

*              IF FILE(NewFileName_M5) = .T.
*                EXIT
*              ELSE
*                LOOP
*              ENDIF

*            ENDFOR

*            IF FILE(NewFileName_M5) = .F.

*              sqlReq = 'ROLLBACK TRANSACTION'
*              sqlRet = RunSQL_Load('', @ar)

*              IF sqlRet = 1
*                MESSAGEBOX('Необходимый для работы файл по пути:' + CHR(13) + ALLTRIM(NewFileName_M5) + ' - НЕ СОЗДАН!!!' + CHR(13) + 'Операция ОТМЕНЕНА!' + CHR(13) + 'Повторите процедуру формирования файла ...' , 16, 'Создание выходного файла для CBS')
*              ENDIF

*              RETURN

*            ENDIF

* -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- *

* 'BudgetKP' - указывает рабочую таблицу, из которой создается XML строка

* NewFileName - указывает путь и имя файла, в которые будет послан результат преобразования
* NewFileName_M5 - указывает путь и имя файла, в которые будет послан результат преобразования

* 1 - ELEMENTS (По умолчанию) XML на базе элементов, производит неформатированный XML в виде длинной строки

* 512 - Результат преобразования посылается в файл, указанный в параметре cOutput, Если файл еще не существует, то он будет создан. Если файл уже существует, то его содержимое будет переписано новым.

          IF sqlRet = 1
            sqlReq = 'COMMIT TRANSACTION'
          ELSE
            sqlReq = 'ROLLBACK TRANSACTION'
          ENDIF

          sqlRet = RunSQL_Load('', @ar)

          IF sqlRet = 1
            MESSAGEBOX('Операция выполнена успешно!', 48, 'Операция передачи сумок в CBS')
          ELSE
            MESSAGEBOX('Операция ОТМЕНЕНА!', 16, 'Операция передачи сумок в CBS')
          ENDIF

        ELSE
          =err('Внимание! Рабочая таблица Export_Sumki не сформирована RECCOUNT Export_Sumki ... ')
        ENDIF

      ELSE
        =err('Внимание! Рабочая таблица Export_Sumki не сформирована USED Export_Sumki ... ')
      ENDIF

    ELSE
      MESSAGEBOX('Курсор для XML не создан, операция ОТМЕНЕНА!', 16, 'Операция передачи сумок в CBS')
    ENDIF

  ELSE

    sqlReq = 'ROLLBACK TRANSACTION'

    sqlRet = RunSQL_Load('', @ar)

    MESSAGEBOX('Удаление записей в таблице XmlToCBS НЕ ВЫПОЛНЕНО - операция ОТМЕНЕНА!', 16, 'Операция передачи сумок в CBS')

  ENDIF

* ---------------------------------------------------------------- *
  DO exit_sql_server IN RunSQL_Load
* ---------------------------------------------------------------- *

ELSE
  =err('Вам не доступна сетевая папка - ' + ALLTRIM(put_tmp) + (put_ObmenToCBS) + ' для записи отчета ... ')
ENDIF

SELECT Lstbudget

SELECT(sel)
RETURN


************************************************************************************************************************************************************


