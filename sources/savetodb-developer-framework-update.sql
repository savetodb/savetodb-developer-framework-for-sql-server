-- =============================================
-- SaveToDB Developer Framework for Microsoft SQL Server
-- Version 10.8, January 9, 2023
--
-- This script updates SaveToDB Developer Framework to the latest version
--
-- Copyright 2018-2023 Gartle LLC
--
-- License: MIT
-- =============================================

IF 1008 <= COALESCE((SELECT CAST(LEFT(HANDLER_CODE, CHARINDEX('.', HANDLER_CODE) - 1) AS int) * 100 + CAST(RIGHT(HANDLER_CODE, LEN(HANDLER_CODE) - CHARINDEX('.', HANDLER_CODE)) AS float) FROM xls.handlers WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME = 'developer_framework' AND COLUMN_NAME = 'version' AND EVENT_NAME = 'Information'), 0)
    RAISERROR('SaveToDB Developer Framework is up-to-date. Update skipped', 11, 0)
GO

UPDATE xls.handlers SET EVENT_NAME = 'Information', HANDLER_CODE = '10.8' WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME = 'developer_framework' AND COLUMN_NAME = 'version' AND EVENT_NAME = 'Information';
IF @@ROWCOUNT = 0
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES ('xls', 'developer_framework', 'version', 'Information', NULL, NULL, 'ATTRIBUTE', '10.8', NULL, NULL, NULL);
GO

IF (SELECT COUNT(*) FROM xls.handlers WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME = 'usp_translations' AND EVENT_NAME = 'Actions' AND HANDLER_NAME = 'SaveToDB Framework Online Help') = 0
    BEGIN
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_translations', NULL, N'Actions', N'xls', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_translations', NULL, N'Actions', N'xls', N'SaveToDB Framework Online Help', N'HTTP', N'https://www.savetodb.com/help/savetodb-framework-procedures.htm#xls.usp_translations', NULL, 91, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'view_all_translations', NULL, N'Actions', N'xls', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'view_all_translations', NULL, N'Actions', N'xls', N'SaveToDB Framework Online Help', N'HTTP', N'https://www.savetodb.com/help/savetodb-framework-views.htm#xls.view_all_translations', NULL, 91, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'view_foreign_keys', NULL, N'Actions', N'xls', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'view_foreign_keys', NULL, N'Actions', N'xls', N'SaveToDB Framework Online Help', N'HTTP', N'https://www.savetodb.com/help/savetodb-framework-views.htm#xls.view_foreign_keys', NULL, 91, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'view_primary_keys', NULL, N'Actions', N'xls', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'view_primary_keys', NULL, N'Actions', N'xls', N'SaveToDB Framework Online Help', N'HTTP', N'https://www.savetodb.com/help/savetodb-framework-views.htm#xls.view_primary_keys', NULL, 91, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'view_unique_keys', NULL, N'Actions', N'xls', N'MenuSeparator90', N'MENUSEPARATOR', NULL, NULL, 90, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'view_unique_keys', NULL, N'Actions', N'xls', N'SaveToDB Framework Online Help', N'HTTP', N'https://www.savetodb.com/help/savetodb-framework-views.htm#xls.view_unique_keys', NULL, 91, NULL);
    END
GO

IF (SELECT COUNT(*) FROM xls.handlers WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME = 'usp_translations' AND EVENT_NAME = 'License') = 0
    BEGIN
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'usp_translations', NULL, N'License', NULL, NULL, N'ATTRIBUTE', N'Qk2DbSbKRNh+J1ggAJXUdJelhLwjerKlbWkwzFd3vehZKIaatIun1W2G5wFxboaBe7fofD1WaVDls7m2eAsgJ2ukffCwJ4OanMoN2NapGhknRMZOkElAoB1jJCeoxx8qzCCcPzQtx1ChxmOJxIgqlVpAz7Swtm2764TKZHvbysU=', NULL, NULL, NULL);
    INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES (N'xls', N'view_all_translations', NULL, N'License', NULL, NULL, N'ATTRIBUTE', N'DuDdxJu6cbtH8oRAhcpmZE8xW51MoJlJYcnjZ1UzNSObX/0gQ5ayac9ZW7eAeAopOtT35rPqfbZjW6B9lBY6YER3FnoQIShlVBXFvNXpig1UUMuYAiqGg4LNCnZjHHd7LxDrjzI5kVf0DN5GYasFpjJ32VU0fw5FRwgCBFCW0eE=', NULL, NULL, NULL);
    END
GO

UPDATE xls.handlers SET HANDLER_TYPE = N'ATTRIBUTE' WHERE HANDLER_TYPE IS NULL AND EVENT_NAME = 'License';
GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: Generated event handlers for developer actions
--
-- 9.1 MySqlStyle
-- =============================================

ALTER VIEW [xls].[view_developer_handlers]
AS

SELECT
    t.TABLE_SCHEMA
    , t.TABLE_NAME
    , CAST(NULL AS nvarchar(128)) AS COLUMN_NAME
    , 'Actions' AS EVENT_NAME
    , v.handler_schema AS HANDLER_SCHEMA
    , v.handler_name AS HANDLER_NAME
    , v.handler_type AS HANDLER_TYPE
    , CAST(v.handler_code AS nvarchar(max)) AS HANDLER_CODE
    , CAST(NULL AS nvarchar(128)) AS TARGET_WORKSHEET
    , v.menu_order AS MENU_ORDER
    , CAST(1 AS bit) AS EDIT_PARAMETERS
FROM
    INFORMATION_SCHEMA.TABLES t
    CROSS JOIN (VALUES
        (200, 'xls', 'MenuSeparator200', 'MENUSEPARATOR', NULL),
        (201, 'xls', 'Generate Select Procedure', 'CODE', 'EXEC xls.xl_actions_generate_procedures NULL, @TableName, @SelectObjectSchema, @SelectObjectName, 1, 0, 0, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (202, 'xls', 'Generate Select and Edit Procedures', 'CODE', 'EXEC xls.xl_actions_generate_procedures NULL, @TableName, @SelectObjectSchema, @SelectObjectName, 1, 1, 0, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (203, 'xls', 'Generate Select and Change Procedures', 'CODE', 'EXEC xls.xl_actions_generate_procedures NULL, @TableName, @SelectObjectSchema, @SelectObjectName, 1, 0, 1, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (204, 'xls', 'MenuSeparator204', 'MENUSEPARATOR', NULL),
        (205, 'xls', 'Generate View', 'CODE', 'EXEC xls.xl_actions_generate_procedures NULL, @TableName, @SelectObjectSchema, @SelectObjectName, 1, 0, 0, 1, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (206, 'xls', 'Generate View and Edit Procedures', 'CODE', 'EXEC xls.xl_actions_generate_procedures NULL, @TableName, @SelectObjectSchema, @SelectObjectName, 1, 1, 0, 1, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (207, 'xls', 'Generate View and Change Handler', 'CODE', 'EXEC xls.xl_actions_generate_procedures NULL, @TableName, @SelectObjectSchema, @SelectObjectName, 1, 0, 1, 1, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (208, 'xls', 'MenuSeparator208', 'MENUSEPARATOR', NULL),
        (209, 'xls', 'Generate Edit Procedures', 'CODE', 'EXEC xls.xl_actions_generate_procedures NULL, @TableName, @SelectObjectSchema, @SelectObjectName, 0, 1, 0, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (210, 'xls', 'Generate Change Handler', 'CODE', 'EXEC xls.xl_actions_generate_procedures NULL, @TableName, @SelectObjectSchema, @SelectObjectName, 0, 0, 1, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (211, 'xls', 'MenuSeparator211', 'MENUSEPARATOR', NULL),
        (212, 'xls', 'Generate Validation List View', 'CODE', 'EXEC xls.xl_actions_generate_handlers NULL, @TableName, @TargetObjectSchema, @TargetObjectName, 1, 1, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (213, 'xls', 'Generate Validation List Procedure', 'CODE', 'EXEC xls.xl_actions_generate_handlers NULL, @TableName, @TargetObjectSchema, @TargetObjectName, 1, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (214, 'xls', 'Generate Parameter Values', 'CODE', 'EXEC xls.xl_actions_generate_handlers NULL, @TableName, @TargetObjectSchema, @TargetObjectName, 2, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (215, 'xls', 'Generate Parameter Values with Empty', 'CODE', 'EXEC xls.xl_actions_generate_handlers NULL, @TableName, @TargetObjectSchema, @TargetObjectName, 3, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (216, 'xls', 'MenuSeparator216', 'MENUSEPARATOR', NULL),
        (217, 'xls', 'Generate Actions Handler', 'CODE', 'EXEC xls.xl_actions_generate_handlers NULL, @TableName, @TargetObjectSchema, @TargetObjectName, 7, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (218, 'xls', 'Generate ContextMenu Handler', 'CODE', 'EXEC xls.xl_actions_generate_handlers NULL, @TableName, @TargetObjectSchema, @TargetObjectName, 4, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (219, 'xls', 'Generate DoubleClick Handler', 'CODE', 'EXEC xls.xl_actions_generate_handlers NULL, @TableName, @TargetObjectSchema, @TargetObjectName, 5, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (220, 'xls', 'Generate SelectionChange Handler', 'CODE', 'EXEC xls.xl_actions_generate_handlers NULL, @TableName, @TargetObjectSchema, @TargetObjectName, 6, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle')
        ) v(menu_order, handler_schema, handler_name, handler_type, handler_code)
WHERE
    (t.TABLE_TYPE = 'BASE TABLE'
    OR t.TABLE_TYPE = 'VIEW' AND v.menu_order IN (200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 216, 217, 218, 219, 220))
    AND NOT t.TABLE_SCHEMA IN ('xls', 'dbo01', 'dbo02', 'savetodb_dev', 'logs', 'doc')
    AND NOT t.TABLE_NAME LIKE 'sys%'
    AND NOT t.TABLE_NAME LIKE 'xl_%'
    AND NOT t.TABLE_NAME LIKE 'viewQueryList%'
    AND NOT t.TABLE_NAME LIKE 'viewParameterValues%'
    AND NOT t.TABLE_NAME LIKE 'viewValidationList%'
    AND NOT t.TABLE_NAME LIKE 'view_query_list%'
    AND NOT t.TABLE_NAME LIKE 'view_xl_%'
    AND EXISTS(SELECT TOP 1 r.ROUTINE_NAME FROM INFORMATION_SCHEMA.ROUTINES r WHERE
        r.ROUTINE_SCHEMA = 'xls' AND r.ROUTINE_NAME = 'xl_actions_generate_procedures' AND r.ROUTINE_TYPE = 'PROCEDURE')

UNION ALL
SELECT
    t.ROUTINE_SCHEMA AS TABLE_SCHEMA
    , t.ROUTINE_NAME AS TABLE_NAME
    , CAST(NULL AS nvarchar(128)) AS COLUMN_NAME
    , 'Actions' AS EVENT_NAME
    , v.handler_schema AS HANDLER_SCHEMA
    , v.handler_name AS HANDLER_NAME
    , v.handler_type AS HANDLER_TYPE
    , v.handler_code AS HANDLER_CODE
    , CAST(NULL AS nvarchar(128)) AS TARGET_WORKSHEET
    , v.menu_order AS MENU_ORDER
    , CAST(1 AS bit) AS EDIT_PARAMETERS
FROM
    INFORMATION_SCHEMA.ROUTINES t
    CROSS JOIN (VALUES
        (208, 'xls', 'MenuSeparator208', 'MENUSEPARATOR', NULL),
        (209, 'xls', 'Generate Edit_procedures', 'CODE', 'EXEC xls.xl_actions_generate_procedures NULL, @TableName, @SelectObjectSchema, @SelectObjectName, 0, 1, 0, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (210, 'xls', 'Generate Change Handler', 'CODE', 'EXEC xls.xl_actions_generate_procedures NULL, @TableName, @SelectObjectSchema, @SelectObjectName, 0, 0, 1, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (216, 'xls', 'MenuSeparator216', 'MENUSEPARATOR', NULL),
        (217, 'xls', 'Generate Actions Handler', 'CODE', 'EXEC xls.xl_actions_generate_handlers NULL, @TableName, @TargetObjectSchema, @TargetObjectName, 7, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (218, 'xls', 'Generate ContextMenu Handler', 'CODE', 'EXEC xls.xl_actions_generate_handlers NULL, @TableName, @TargetObjectSchema, @TargetObjectName, 4, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (219, 'xls', 'Generate DoubleClick Handler', 'CODE', 'EXEC xls.xl_actions_generate_handlers NULL, @TableName, @TargetObjectSchema, @TargetObjectName, 5, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle'),
        (220, 'xls', 'Generate SelectionChange Handler', 'CODE', 'EXEC xls.xl_actions_generate_handlers NULL, @TableName, @TargetObjectSchema, @TargetObjectName, 6, 0, @ExecuteScript, @RecreateProceduresIfExist, @DataLanguage, NULL, NULL, @MySqlStyle')
        ) v(menu_order, handler_schema, handler_name, handler_type, handler_code)
WHERE
    t.ROUTINE_TYPE = 'PROCEDURE'
    AND NOT t.ROUTINE_SCHEMA IN ('xls', 'dbo01', 'dbo02', 'savetodb_dev', 'logs', 'doc')
    AND NOT t.ROUTINE_NAME LIKE 'sp%'
    AND NOT t.ROUTINE_NAME LIKE 'xl%'
    AND NOT t.ROUTINE_NAME LIKE '%_insert'
    AND NOT t.ROUTINE_NAME LIKE '%_update'
    AND NOT t.ROUTINE_NAME LIKE '%_delete'
    AND NOT t.ROUTINE_NAME LIKE '%_merge'
    AND NOT t.ROUTINE_NAME LIKE '%_change'
    AND NOT t.ROUTINE_NAME LIKE 'uspExcelEvent%'
    AND NOT t.ROUTINE_NAME LIKE 'uspParameterValues%'
    AND NOT t.ROUTINE_NAME LIKE 'uspValidationList%'
    AND NOT t.ROUTINE_NAME LIKE 'uspAdd%'
    AND NOT t.ROUTINE_NAME LIKE 'uspSet%'
    AND NOT t.ROUTINE_NAME LIKE 'uspInsert%'
    AND NOT t.ROUTINE_NAME LIKE 'uspUpdate%'
    AND NOT t.ROUTINE_NAME LIKE 'uspDelete%'
    AND NOT t.ROUTINE_NAME LIKE 'Add%'
    AND NOT t.ROUTINE_NAME LIKE 'Set%'
    AND NOT t.ROUTINE_NAME LIKE 'Insert%'
    AND NOT t.ROUTINE_NAME LIKE 'Update%'
    AND NOT t.ROUTINE_NAME LIKE 'Delete%'
    AND NOT t.ROUTINE_NAME LIKE 'usp_xl_%'
    AND NOT t.ROUTINE_NAME LIKE 'usp_import_%'
    AND NOT t.ROUTINE_NAME LIKE 'usp_export_%'
    AND EXISTS(SELECT TOP 1 r.ROUTINE_NAME FROM INFORMATION_SCHEMA.ROUTINES r WHERE
        r.ROUTINE_SCHEMA = 'xls' AND r.ROUTINE_NAME = 'xl_actions_generate_procedures' AND r.ROUTINE_TYPE = 'PROCEDURE')


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: The procedure generates SaveToDB handler objects
--
-- @HandlerType
-- 1 - ValidationList
-- 2 - ParameterValues
-- 3 - ParameterValues with the null value
-- 4 - ContextMenu
-- 5 - DoubleClick
-- 6 - SelectionChange
-- 7 - Actions
--
-- 9.1 MySQL style
-- =============================================

ALTER PROCEDURE [xls].[xl_actions_generate_handlers]
    @BaseTableSchema nvarchar(128) = NULL
    , @BaseTableName nvarchar(128) = NULL
    , @TargetObjectSchema nvarchar(128) = NULL
    , @TargetObjectName nvarchar(128) = NULL
    , @HandlerType int = NULL
    , @GenerateTargetAsView bit = 0
    , @ExecuteScript bit = 0
    , @RecreateProceduresIfExist bit = 0
    , @DataLanguage varchar(10) = NULL
    , @SelectCommands bit = NULL
    , @PrintCommands bit = NULL
    , @MySqlStyle bit = 0
AS
BEGIN

BEGIN -- CHECKS --

SET NOCOUNT ON

IF @ExecuteScript IS NULL SET @ExecuteScript = 0
IF @RecreateProceduresIfExist IS NULL SET @RecreateProceduresIfExist = 0
IF @DataLanguage IS NULL SET @DataLanguage = 'en'

IF @BaseTableSchema IS NULL AND @BaseTableName IS NOT NULL AND CHARINDEX('.', @BaseTableName) > 1
    BEGIN
    SET @BaseTableSchema = LEFT(@BaseTableName, CHARINDEX('.', @BaseTableName) - 1)
    SET @BaseTableName = SUBSTRING(@BaseTableName, CHARINDEX('.', @BaseTableName) + 1, LEN(@BaseTableName))
    END

DECLARE @message nvarchar(max)

DECLARE @Tab char(4) = REPLICATE(' ', 4)
DECLARE @CrLf char(2) = CHAR(13) + CHAR(10)

IF @BaseTableSchema IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage) + ' @BaseTableSchema'
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF @BaseTableName IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage) + ' @BaseTableName'
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF LEFT(@BaseTableSchema, 1) = '[' AND RIGHT(@BaseTableSchema, 1) = ']'
    SET @BaseTableSchema = REPLACE(SUBSTRING(@BaseTableSchema, 2, LEN(@BaseTableSchema) - 2), ']]', ']')

IF LEFT(@BaseTableName, 1) = '[' AND RIGHT(@BaseTableName, 1) = ']'
    SET @BaseTableName = REPLACE(SUBSTRING(@BaseTableName, 2, LEN(@BaseTableName) - 2), ']]', ']')

IF @TargetObjectSchema IS NULL
    SET @TargetObjectSchema = @BaseTableSchema

DECLARE @BaseTable nvarchar(255)

IF xls.get_escaped_parameter_name(@BaseTableSchema) = @BaseTableSchema
    AND xls.get_escaped_parameter_name(@BaseTableName) = @BaseTableName
    SET @BaseTable = @BaseTableSchema + '.' + @BaseTableName
ELSE
    SET @BaseTable = QUOTENAME(@BaseTableSchema) + '.' + QUOTENAME(@BaseTableName)

IF OBJECT_ID(@BaseTable) IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Table ''%s'' does not exist', @DataLanguage)
    RAISERROR(@message, 11, 0, @BaseTable);
    RETURN
    END

IF SCHEMA_ID(@TargetObjectSchema) IS NULL AND @ExecuteScript = 1
    BEGIN
    SET @message = xls.get_translated_string('Target schema ''%s'' does not exist', @DataLanguage)
    RAISERROR(@message, 11, 0, @TargetObjectSchema);
    RETURN
    END

IF NOT @HandlerType IN (1, 2, 3, 4, 5, 6, 7)
    BEGIN
    SET @message = xls.get_translated_string('Unknown handler type: %i', @DataLanguage) + @CrLf + @CrLf
        + xls.get_translated_string('Specify', @DataLanguage)
        + ' @HandlerType:' + @CrLf + @CrLf
        + '1 - ValidationList' + @CrLf
        + '2 - ParameterValues' + @CrLf
        + '3 - ParameterValues with the null value' + @CrLf
        + '4 - ContextMenu' + @CrLf
        + '5 - DoubleClick' + @CrLf
        + '6 - SelectionChange' + @CrLf
        + '7 - Actions' + @CrLf
    RAISERROR(@message, 11, 0, @HandlerType);
    RETURN
    END

END

BEGIN -- PRIMARY KEY AND UNIQUE COLUMNS --

DECLARE @PrimaryColumnCount int
DECLARE @PrimaryColumn nvarchar(128)
DECLARE @UniqueColumn nvarchar(128)
DECLARE @ExampleColumn nvarchar(128)

IF @HandlerType IN (1, 2, 3)
    BEGIN
    SELECT
        @PrimaryColumnCount = COUNT(ccu.COLUMN_NAME)
        , @PrimaryColumn = MAX(ccu.COLUMN_NAME)
    FROM
        INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE ccu
        INNER JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
            ON tc.TABLE_SCHEMA = ccu.TABLE_SCHEMA AND tc.TABLE_NAME = ccu.TABLE_NAME
                AND tc.CONSTRAINT_SCHEMA = ccu.CONSTRAINT_SCHEMA AND tc.CONSTRAINT_NAME = ccu.CONSTRAINT_NAME
                AND tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
    WHERE
        ccu.TABLE_SCHEMA = @BaseTableSchema
        AND ccu.TABLE_NAME = @BaseTableName

    IF @PrimaryColumnCount = 0
        BEGIN
        SET @message = xls.get_translated_string('''%s.%s'' has no primary key', @DataLanguage)
        RAISERROR(@message, 11, 0, @BaseTableSchema, @BaseTableName);
        RETURN
        END

    IF @PrimaryColumnCount > 1
        BEGIN
        SET @message = xls.get_translated_string('''%s.%s'' has more than one primary key column', @DataLanguage)
        RAISERROR(@message, 11, 0, @BaseTableSchema, @BaseTableName);
        RETURN
        END

    SELECT
        TOP 1
        @UniqueColumn = COLUMN_NAME
    FROM
        (
            SELECT
                MAX(ccu.COLUMN_NAME) AS COLUMN_NAME
                , MAX(CASE
                    WHEN c.CHARACTER_MAXIMUM_LENGTH BETWEEN 10 AND 50 THEN 0
                    WHEN c.CHARACTER_MAXIMUM_LENGTH > 50 THEN 1
                    WHEN c.CHARACTER_MAXIMUM_LENGTH < 10 THEN 2
                    ELSE 3 END) AS LENGTH_RANK
            FROM
                INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE ccu
                INNER JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
                    ON tc.TABLE_SCHEMA = ccu.TABLE_SCHEMA AND tc.TABLE_NAME = ccu.TABLE_NAME
                        AND tc.CONSTRAINT_SCHEMA = ccu.CONSTRAINT_SCHEMA AND tc.CONSTRAINT_NAME = ccu.CONSTRAINT_NAME
                        AND tc.CONSTRAINT_TYPE = 'UNIQUE'
                INNER JOIN INFORMATION_SCHEMA.COLUMNS c
                    ON c.TABLE_SCHEMA = ccu.TABLE_SCHEMA AND c.TABLE_NAME = ccu.TABLE_NAME AND c.COLUMN_NAME = ccu.COLUMN_NAME
            WHERE
                ccu.TABLE_SCHEMA = @BaseTableSchema
                AND ccu.TABLE_NAME = @BaseTableName
            GROUP BY
                ccu.TABLE_SCHEMA
                , ccu.TABLE_NAME
            HAVING
                COUNT(*) = 1
        ) t
    ORDER BY
        t.LENGTH_RANK

    IF @UniqueColumn IS NULL
        BEGIN
        SELECT
            TOP 1
            @UniqueColumn = COLUMN_NAME
        FROM
            (
                SELECT
                    c.COLUMN_NAME
                    , CASE
                        WHEN c.CHARACTER_MAXIMUM_LENGTH BETWEEN 10 AND 50 THEN 0
                        WHEN c.CHARACTER_MAXIMUM_LENGTH > 50 THEN 1
                        WHEN c.CHARACTER_MAXIMUM_LENGTH < 10 THEN 2
                        ELSE 3 END AS LENGTH_RANK
                FROM
                    INFORMATION_SCHEMA.COLUMNS c
                WHERE
                    c.TABLE_SCHEMA = @BaseTableSchema
                    AND c.TABLE_NAME = @BaseTableName
                    AND c.CHARACTER_MAXIMUM_LENGTH IS NOT NULL
            ) t
        ORDER BY
            t.LENGTH_RANK
        END

    SET @ExampleColumn = @UniqueColumn
    END
ELSE IF OBJECT_ID(@BaseTable, 'P') IS NOT NULL
    SET @ExampleColumn = 'id'
ELSE
    BEGIN
    SELECT
        TOP 1
        @ExampleColumn = COLUMN_NAME
    FROM
        INFORMATION_SCHEMA.COLUMNS c
    WHERE
        c.TABLE_SCHEMA = @BaseTableSchema
        AND c.TABLE_NAME = @BaseTableName
    ORDER BY
        c.ORDINAL_POSITION
    END

END

BEGIN -- HANDLER PARAMETERS --

DECLARE @termColumnName nvarchar(128) = '@ColumnName'
DECLARE @termCellValue nvarchar(128) = '@CellValue'
DECLARE @termCellNumberValue nvarchar(128) = '@CellNumberValue'
DECLARE @termCellDateTimeValue nvarchar(128) = '@CellDateTimeValue'
DECLARE @termDataLanguage nvarchar(128) = '@DataLanguage'

IF @MySqlStyle = 1
    BEGIN
    SET @termColumnName = '@column_name'
    SET @termCellValue = '@cell_value'
    SET @termCellNumberValue = '@cell_number_value'
    SET @termCellDateTimeValue = '@cell_datetime_value'
    SET @termDataLanguage = '@data_language'
    END

DECLARE @HasDates bit
DECLARE @HasNumbers bit

DECLARE @HandlerParameters nvarchar(max) = ''

DECLARE @sql nvarchar(max)

IF OBJECT_ID(@BaseTable, 'P') IS NOT NULL
    BEGIN
    SET @HasDates = 1
    SET @HasNumbers = 1

    IF OBJECT_ID('sys.dm_exec_describe_first_result_set') IS NOT NULL
        BEGIN
        SET @sql = 'SELECT @HandlerParameters = STUFF((
    SELECT ''
    , @'' + xls.get_escaped_parameter_name(c.name)
        + '' '' + c.system_type_name
        + '' = NULL''
    FROM
        sys.dm_exec_describe_first_result_set(N''' + REPLACE(@BaseTable, '''', '''''') + ''', NULL, 0) c
    WHERE
        c.is_hidden = 0
        AND NOT c.system_type_name IN (''geography'', ''geometry'', ''image'')
    ORDER BY
        c.column_ordinal
    FOR XML PATH(''''), TYPE).value(''.'', ''nvarchar(MAX)''), 1, 2, '''')'

        EXEC sys.sp_executesql @stmt = @sql, @params = N'@HandlerParameters nvarchar(max) out', @HandlerParameters = @HandlerParameters out
        END
    END
ELSE
    BEGIN
    SELECT
        @HasDates =    MAX(CASE WHEN c.DATETIME_PRECISION IS NULL THEN 0 ELSE 1 END)
        , @HasNumbers = MAX(CASE WHEN c.NUMERIC_PRECISION IS NULL THEN 0 WHEN c.DATA_TYPE IN ('bit') THEN 1 ELSE 1 END)
    FROM
        INFORMATION_SCHEMA.COLUMNS c
    WHERE
        c.TABLE_SCHEMA = @BaseTableSchema
        AND c.TABLE_NAME = @BaseTableName

    SELECT @HandlerParameters = STUFF((
        SELECT
            @CrLf + '    , @' + xls.get_escaped_parameter_name(c.COLUMN_NAME)
            + ' ' + c.DATA_TYPE
            + CASE WHEN c.CHARACTER_MAXIMUM_LENGTH IS NULL THEN '' ELSE '(' + CAST(c.CHARACTER_MAXIMUM_LENGTH AS varchar(5)) + ')' END
            + ' = NULL'
        FROM
            INFORMATION_SCHEMA.COLUMNS c
        WHERE
            c.TABLE_SCHEMA = @BaseTableSchema
            AND c.TABLE_NAME = @BaseTableName
            AND NOT c.DATA_TYPE IN ('geography', 'geometry', 'image')
        ORDER BY
            c.ORDINAL_POSITION
        FOR XML PATH(''), TYPE).value('.', 'nvarchar(MAX)'), 1, 2, '')
    END

DECLARE @LanguageParameters nvarchar(max) =
    + @termDataLanguage + ' varchar(10) = NULL' + @CrLf

DECLARE @ContextParameters nvarchar(max) =
    '    ' + @termColumnName + ' nvarchar(128) = NULL' + @CrLf
    + '    , ' + @termCellValue + ' nvarchar(255) = NULL' + @CrLf
    + CASE WHEN @HasNumbers = 1 THEN '    , ' + @termCellNumberValue + ' float = NULL' + @CrLf ELSE '' END
    + CASE WHEN @HasDates = 1 THEN '    , ' + @termCellDateTimeValue + ' datetime = NULL' + @CrLf ELSE '' END
    + '    , ' + @LanguageParameters

END

BEGIN -- PROCEDURE NAMES --

DECLARE @TargetObject nvarchar(255)

IF @TargetObjectName IS NULL
    BEGIN
    IF @HandlerType = 1
        SET @TargetObjectName = 'xl_validation_list_' + @BaseTableName + '_' + @PrimaryColumn
    ELSE IF @HandlerType = 2
        SET @TargetObjectName = 'xl_parameter_values_' + @BaseTableName + '_' + @PrimaryColumn
    ELSE IF @HandlerType = 3
        SET @TargetObjectName = 'xl_parameter_values_' + @BaseTableName + '_' + @PrimaryColumn + '_or_null'
    ELSE IF @HandlerType = 4
        SET @TargetObjectName = 'xl_context_menu_' + @BaseTableName
    ELSE IF @HandlerType = 5
        SET @TargetObjectName = 'xl_double_click_' + @BaseTableName
    ELSE IF @HandlerType = 6
        SET @TargetObjectName = 'xl_selection_change_' + @BaseTableName
    ELSE IF @HandlerType = 7
        SET @TargetObjectName = 'xl_actions_' + @BaseTableName
    END

IF xls.get_escaped_parameter_name(@TargetObjectSchema) = @TargetObjectSchema
    AND xls.get_escaped_parameter_name(@TargetObjectName) = @TargetObjectName
    SET @TargetObject = @TargetObjectSchema + '.' + @TargetObjectName
ELSE
    SET @TargetObject = QUOTENAME(@TargetObjectSchema) + '.' + QUOTENAME(@TargetObjectName)

END

BEGIN -- HEADERS --

DECLARE @HeaderText nvarchar(max) = '
-- =============================================
-- Author:      ' + xls.get_translated_string('<Author>', 'en') + '
-- Release:     ' + xls.get_translated_string('<Release>', 'en') + ', ' + CONVERT(char(10), GETDATE(), 120)
    + CASE
        WHEN @HandlerType = 1 THEN '
-- Description: Validation list of ' + @BaseTableSchema + '.' + @BaseTableName + '.' + @PrimaryColumn
        WHEN @HandlerType IN (2, 3) THEN '
-- Description: Parameter values of ' + @BaseTableSchema + '.' + @BaseTableName + '.' + @PrimaryColumn
        WHEN @HandlerType = 4 THEN '
-- Description: Context menu handler for ' + @BaseTableSchema + '.' + @BaseTableName
        WHEN @HandlerType = 5 THEN '
-- Description: Double-click handler for ' + @BaseTableSchema + '.' + @BaseTableName
        WHEN @HandlerType = 6 THEN '
-- Description: Selection change handler for ' + @BaseTableSchema + '.' + @BaseTableName
        WHEN @HandlerType = 7 THEN '
-- Description: Actions menu handler for ' + @BaseTableSchema + '.' + @BaseTableName
        ELSE '
-- Description: '
        END + '
-- =============================================

'

END

BEGIN -- PROCEDURE DEFINITIONS --

DECLARE @DeleteSQL nvarchar(max) =
    CASE WHEN @GenerateTargetAsView = 1 THEN
        'IF OBJECT_ID(N''' + REPLACE(@TargetObject, '''', '''''') + ''', ''V'') IS NOT NULL' + @CrLf
        + @Tab + 'DROP VIEW '+ @TargetObject + @CrLf
    ELSE
        'IF OBJECT_ID(N''' + REPLACE(@TargetObject, '''', '''''') + ''', ''P'') IS NOT NULL' + @CrLf
        + @Tab + 'DROP PROCEDURE '+ @TargetObject + @CrLf
    END

DECLARE @Body nvarchar(max)

IF @HandlerType IN (1, 2, 3)
    BEGIN
    SET @Body =
    CASE WHEN @HandlerType = 3 THEN
        'SELECT NULL AS ' + xls.get_friendly_column_name(@PrimaryColumn) + COALESCE(', NULL AS ' + xls.get_friendly_column_name(@UniqueColumn), '') + ' UNION ALL' + @CrLf
        ELSE '' END
    + 'SELECT' + @CrLf
    + @Tab + 't.' + xls.get_friendly_column_name(@PrimaryColumn) + @CrLf
    + COALESCE(@Tab + ', t.' + xls.get_friendly_column_name(@UniqueColumn) + @CrLf, '')
    + 'FROM'+ @CrLf
    + @Tab + @BaseTable + ' t' + @CrLf
    + CASE WHEN @GenerateTargetAsView = 1 THEN '' ELSE
      'ORDER BY' + @CrLf
       + @Tab + CASE WHEN @UniqueColumn IS NOT NULL THEN 't.' + xls.get_friendly_column_name(@UniqueColumn) + @CrLf + @Tab + ', ' ELSE '' END
       + 't.' + xls.get_friendly_column_name(@PrimaryColumn) + @CrLf
       END
    END
ELSE IF @HandlerType IN (4)
    BEGIN
    SET @Body = xls.get_translated_string('-- Place your code for the context menu handler here', @DataLanguage) + @CrLf
        + @CrLf
        + xls.get_translated_string('-- You may execute SELECT, INSERT, UPDATE and DELETE commands', @DataLanguage) + @CrLf
        + @CrLf
        + xls.get_translated_string('-- You may use predefined parameters like @ColumnName or @CellValue', @DataLanguage) + @CrLf
        + xls.get_translated_string('-- and values of the current row using table-specific parameters like ', @DataLanguage)
        + '@' + xls.get_escaped_parameter_name(@ExampleColumn) + @CrLf
    END
ELSE IF @HandlerType IN (5)
    BEGIN
    SET @Body = xls.get_translated_string('-- Place your code for the double-click handler here', @DataLanguage) + @CrLf
        + @CrLf
        + xls.get_translated_string('-- You may execute SELECT commands', @DataLanguage) + @CrLf
        + @CrLf
        + xls.get_translated_string('-- You may use predefined parameters like @ColumnName or @CellValue', @DataLanguage) + @CrLf
        + xls.get_translated_string('-- and values of the current row using table-specific parameters like ', @DataLanguage)
        + '@' + xls.get_escaped_parameter_name(@ExampleColumn) + @CrLf
    END
ELSE IF @HandlerType IN (6)
    BEGIN
    SET @Body = xls.get_translated_string('-- Place your code for the selection change handler here', @DataLanguage) + @CrLf
        + @CrLf
        + xls.get_translated_string('-- You may execute SELECT commands', @DataLanguage) + @CrLf
        + @CrLf
        + xls.get_translated_string('-- You may use predefined parameters like @ColumnName or @CellValue', @DataLanguage) + @CrLf
        + xls.get_translated_string('-- and values of the current row using table-specific parameters like ', @DataLanguage)
        + '@' + xls.get_escaped_parameter_name(@ExampleColumn) + @CrLf
    END
ELSE IF @HandlerType IN (7)
    BEGIN
    SET @Body = xls.get_translated_string('-- Place your code for the Actions menu handler here', @DataLanguage) + @CrLf
        + @CrLf
        + xls.get_translated_string('-- You may execute any command', @DataLanguage) + @CrLf
        + @CrLf
        + xls.get_translated_string('-- Unlike the context menu handlers, these handlers have no row context values', @DataLanguage) + @CrLf
        + xls.get_translated_string('-- when the active cell is outside of the table', @DataLanguage) + @CrLf
    END
ELSE
    SET @Body = xls.get_translated_string('-- Place your code here', @DataLanguage) + @CrLf

DECLARE @ProcedureSQL nvarchar(max) =
    @HeaderText
    + CASE WHEN @GenerateTargetAsView = 1 THEN
        'CREATE' + ' VIEW ' + @TargetObject + @CrLf
        ELSE
        'CREATE' + ' PROCEDURE ' + @TargetObject + @CrLf
        + CASE
            WHEN @HandlerType IN (4, 5, 6, 7) THEN @ContextParameters + @HandlerParameters + @CrLf
            ELSE '' END
        END
    + 'AS' + @CrLf
    + CASE WHEN @GenerateTargetAsView = 1 THEN '' ELSE 'BEGIN' + @CrLf + @CrLf + 'SET NOCOUNT ON' + @CrLf END
    + @CrLf
    + @Body
    + @CrLf
    + CASE WHEN @GenerateTargetAsView = 1 THEN '' ELSE 'END' + @CrLf END
    + @CrLf

END

BEGIN -- HELP --

DECLARE @GoLine nvarchar(10) = 'GO' + @CrLf + @CrLf

DECLARE @Help nvarchar(max) =
    CASE WHEN xls.get_translated_string('<Author>', 'en') = '<Author>' THEN
    '-- You may define the <Author> and <Release> values in the xls.translations table:' + @CrLf
    + @CrLf
    + '-- TABLE_SCHEMA TABLE_NAME COLUMN_NAME LANGUAGE_NAME TRANSLATED_NAME' + @CrLf
    + '-- ------------ ---------- ----------- ------------- ---------------' + @CrLf
    + '-- xls          strings    <Author>    en            <Your value>' + @CrLf
    + '-- xls          strings    <Release>   en            <Your value>' + @CrLf
    + @CrLf
    ELSE '' END

    + '-- Use the following code to attach the handler to the target object:'+ @CrLf
    + @CrLf
    + '-- INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE)' + @CrLf
    + '--     VALUES('
    + CASE @HandlerType
        WHEN 1 THEN 'N''<target object schema>'', N''<target object name>'''
        WHEN 2 THEN 'N''<target object schema>'', N''<target object name>'''
        WHEN 3 THEN 'N''<target object schema>'', N''<target object name>'''
        WHEN 4 THEN 'N''' + REPLACE(@BaseTableSchema, '''', '''''') + ''', N''' + REPLACE(@BaseTableName, '''', '''''') + ''''
        WHEN 5 THEN 'N''' + REPLACE(@BaseTableSchema, '''', '''''') + ''', N''' + REPLACE(@BaseTableName, '''', '''''') + ''''
        WHEN 6 THEN 'N''' + REPLACE(@BaseTableSchema, '''', '''''') + ''', N''' + REPLACE(@BaseTableName, '''', '''''') + ''''
        WHEN 7 THEN 'N''' + REPLACE(@BaseTableSchema, '''', '''''') + ''', N''' + REPLACE(@BaseTableName, '''', '''''') + ''''
        ELSE 'N''<target object schema>'', N''<target object name>'''
        END + ', '
    + CASE @HandlerType
        WHEN 1 THEN 'N''<target column>'''
        WHEN 2 THEN 'N''<target parameter>'''
        WHEN 3 THEN 'N''<target parameter>'''
        WHEN 4 THEN 'N''<target column or NULL>'''
        WHEN 5 THEN 'N''<target column or NULL>'''
        WHEN 6 THEN 'N''<target column or NULL>'''
        WHEN 7 THEN 'NULL'
        ELSE 'N''<target column>'''
        END + ', '
    + '''' + CASE @HandlerType
        WHEN 1 THEN 'ValidationList'
        WHEN 2 THEN 'ParameterValues'
        WHEN 3 THEN 'ParameterValues'
        WHEN 4 THEN 'ContextMenu'
        WHEN 5 THEN 'DoubleClick'
        WHEN 6 THEN 'SelectionChange'
        WHEN 7 THEN 'Actions'
        ELSE 'N''<target event>'''
        END + ''', '
    + 'N''' + REPLACE(@TargetObjectSchema, '''', '''''') + ''', N''' + @TargetObjectName + ''', '''
    + CASE WHEN @GenerateTargetAsView = 1 THEN 'VIEW' ELSE 'PROCEDURE' END + ''')' + @CrLf
    + @CrLf

    + @GoLine

END

BEGIN -- EXECUTE GENERATED CODES --

IF @SelectCommands IS NULL AND @PrintCommands IS NULL SET @SelectCommands = 1

IF @PrintCommands = 1
    BEGIN
    RAISERROR(@Help, 0, 1) WITH NOWAIT
    RAISERROR(@DeleteSQL, 0, 1) WITH NOWAIT
    RAISERROR(@GoLine,  0, 1) WITH NOWAIT
    RAISERROR(@ProcedureSQL, 0, 1) WITH NOWAIT
    RAISERROR(@GoLine,  0, 1) WITH NOWAIT
    END

IF @ExecuteScript = 1
    BEGIN
    IF @RecreateProceduresIfExist = 1
        EXEC (@DeleteSQL)

    IF OBJECT_ID(@TargetObject, 'P') IS NULL
        EXEC (@ProcedureSQL)

    IF @SelectCommands = 1
        SELECT 'Created' AS [message]
    END

ELSE IF @SelectCommands = 1
    BEGIN
    SELECT
        @Help
        + @DeleteSQL
        + 'GO' + @CrLf + @CrLf
        + @ProcedureSQL
        + 'GO' + @CrLf + @CrLf
        AS [message]
    END

END

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: The procedure generates SELECT, INSERT, UPDATE, DELETE, and CHANGE procedures
--
-- 9.1 MySQL style; improved cell change handlers.
-- =============================================

ALTER PROCEDURE [xls].[xl_actions_generate_procedures]
    @BaseTableSchema nvarchar(128) = NULL
    , @BaseTableName nvarchar(128) = NULL
    , @SelectObjectSchema nvarchar(128) = NULL
    , @SelectObjectName nvarchar(128) = NULL
    , @GenerateSelectObject bit = 1
    , @GenerateEditProcedures bit = 1
    , @GenerateChangeHandler bit = 0
    , @GenerateSelectAsView bit = 0
    , @ExecuteScript bit = 0
    , @RecreateProceduresIfExist bit = 0
    , @DataLanguage varchar(10) = NULL
    , @SelectCommands bit = NULL
    , @PrintCommands bit = NULL
    , @MySqlStyle bit = 0
AS
BEGIN

BEGIN -- CHECKS --

SET NOCOUNT ON

IF @ExecuteScript IS NULL SET @ExecuteScript = 0
IF @RecreateProceduresIfExist IS NULL SET @RecreateProceduresIfExist = 0
IF @GenerateSelectObject IS NULL SET @GenerateSelectObject = 1
IF @GenerateEditProcedures IS NULL SET @GenerateEditProcedures = 1
IF @GenerateChangeHandler IS NULL SET @GenerateChangeHandler = 0
IF @DataLanguage IS NULL SET @DataLanguage = 'en'

IF @BaseTableSchema IS NULL AND @BaseTableName IS NOT NULL AND CHARINDEX('.', @BaseTableName) > 1
    BEGIN
    SET @BaseTableSchema = LEFT(@BaseTableName, CHARINDEX('.', @BaseTableName) - 1)
    SET @BaseTableName = SUBSTRING(@BaseTableName, CHARINDEX('.', @BaseTableName) + 1, LEN(@BaseTableName))
    END

DECLARE @message nvarchar(max)

IF @BaseTableSchema IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage) + ' @BaseTableSchema'
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF @BaseTableName IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage) + ' @BaseTableName'
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF LEFT(@BaseTableSchema, 1) = '[' AND RIGHT(@BaseTableSchema, 1) = ']'
    SET @BaseTableSchema = REPLACE(SUBSTRING(@BaseTableSchema, 2, LEN(@BaseTableSchema) - 2), ']]', ']')

IF LEFT(@BaseTableName, 1) = '[' AND RIGHT(@BaseTableName, 1) = ']'
    SET @BaseTableName = REPLACE(SUBSTRING(@BaseTableName, 2, LEN(@BaseTableName) - 2), ']]', ']')

IF @SelectObjectSchema IS NULL
    SET @SelectObjectSchema = @BaseTableSchema

DECLARE @BaseTable nvarchar(255)

IF xls.get_escaped_parameter_name(@BaseTableSchema) = @BaseTableSchema
    AND xls.get_escaped_parameter_name(@BaseTableName) = @BaseTableName
    SET @BaseTable = @BaseTableSchema + '.' + @BaseTableName
ELSE
    SET @BaseTable = QUOTENAME(@BaseTableSchema) + '.' + QUOTENAME(@BaseTableName)

IF OBJECT_ID(@BaseTable) IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Table ''%s'' does not exist', @DataLanguage)
    RAISERROR(@message, 11, 0, @BaseTable);
    RETURN
    END

IF SCHEMA_ID(@SelectObjectSchema) IS NULL AND @ExecuteScript = 1
    BEGIN
    SET @message = xls.get_translated_string('Target schema ''%s'' does not exist', @DataLanguage)
    RAISERROR(@message, 11, 0, @SelectObjectSchema);
    RETURN
    END

DECLARE @UnderlyingTableSchema nvarchar(128) = @BaseTableSchema
DECLARE @UnderlyingTableName nvarchar(128) = @BaseTableName
DECLARE @UnderlyingTable nvarchar(255) = @BaseTable

IF OBJECT_ID(@BaseTable, 'V') IS NOT NULL
    BEGIN
    SET @UnderlyingTable = xls.get_view_underlying_table(@BaseTableSchema, @BaseTableName)
    IF @UnderlyingTable IS NOT NULL
        BEGIN
        SET @UnderlyingTableSchema = OBJECT_SCHEMA_NAME(OBJECT_ID(@UnderlyingTable))
        SET @UnderlyingTableName = OBJECT_NAME(OBJECT_ID(@UnderlyingTable))
        END
    END
ELSE IF OBJECT_ID(@BaseTable, 'P') IS NOT NULL
    BEGIN
    SET @UnderlyingTable = xls.get_procedure_underlying_table(@BaseTableSchema, @BaseTableName)
    IF @UnderlyingTable IS NOT NULL
        BEGIN
        SET @UnderlyingTableSchema = OBJECT_SCHEMA_NAME(OBJECT_ID(@UnderlyingTable))
        SET @UnderlyingTableName = OBJECT_NAME(OBJECT_ID(@UnderlyingTable))
        END
    END

IF (
    SELECT
        COUNT(*)
    FROM
        INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
    WHERE
        tc.TABLE_SCHEMA = @UnderlyingTableSchema
        AND tc.TABLE_NAME = @UnderlyingTableName
        AND tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
    ) = 0
    BEGIN
    SET @message = xls.get_translated_string('''%s.%s'' has no primary key', @DataLanguage)
    RAISERROR(@message, 11, 0, @UnderlyingTableSchema, @UnderlyingTableName);
    RETURN
    END

END

BEGIN -- DECLARATIONS --

DECLARE @SelectFields nvarchar(max) = ''
DECLARE @UpdateFields nvarchar(max) = ''
DECLARE @InsertFields nvarchar(max) = ''
DECLARE @InsertValues nvarchar(max) = ''
DECLARE @InsertParameters nvarchar(max) = ''
DECLARE @UpdateParameters nvarchar(max) = ''
DECLARE @DeleteParameters nvarchar(max) = ''
DECLARE @DeleteWhere nvarchar(max) = ''
DECLARE @SelectJoin nvarchar(max) = ''
DECLARE @JoinNumber int = 0
DECLARE @QuotedColumnName nvarchar(128)
DECLARE @ParameterSQL nvarchar(255)

DECLARE @COLUMN_NAME nvarchar(128)
DECLARE @ORDINAL_POSITION int
DECLARE @DATA_TYPE nvarchar(128)
DECLARE @CHARACTER_MAXIMUM_LENGTH int

DECLARE @COLUMN_EXISTS bit
DECLARE @IS_NULLABLE nvarchar(3)
DECLARE @IS_PRIMARY_KEY bit
DECLARE @IS_IDENTITY bit
DECLARE @IS_ROWGUIDCOL bit
DECLARE @IS_COMPUTED bit

DECLARE @LINKED_TABLE_SCHEMA nvarchar(128)
DECLARE @LINKED_TABLE_NAME nvarchar(128)
DECLARE @LINKED_COLUMN_NAME nvarchar(128)

DECLARE @Tab char(4) = REPLICATE(' ', 4)
DECLARE @CrLf char(2) = CHAR(13) + CHAR(10)

END

BEGIN -- COLUMNS --

DECLARE @t TABLE (ORDINAL_POSITION int PRIMARY KEY, COLUMN_NAME nvarchar(128), DATA_TYPE nvarchar(128), CHARACTER_MAXIMUM_LENGTH int, IS_NULLABLE char(3), SOURCE_COLUMN nvarchar(128))
DECLARE @sql nvarchar(max)

IF OBJECT_ID(@BaseTable, 'P') IS NOT NULL
    BEGIN
    SET @sql = '
        SELECT
            c.column_ordinal
            , c.name
            , c.system_type_name
            , NULL AS CHARACTER_MAXIMUM_LENGTH
            , CASE WHEN c.is_nullable = 1 THEN ''YES'' ELSE ''NO'' END AS IS_NULLABLE
            , CASE WHEN c.source_schema = ''' + REPLACE(@UnderlyingTableSchema, '''', '''''') + ''' AND c.source_table = ''' + REPLACE(@UnderlyingTableName, '''', '''''') + ''' THEN c.source_column ELSE NULL END AS SOURCE_COLUMN
        FROM
            sys.dm_exec_describe_first_result_set(N''EXEC ' + REPLACE(@BaseTable, '''', '''''')  +''', NULL, 1) c
        WHERE
            c.is_hidden = 0'

    INSERT INTO @t (ORDINAL_POSITION, COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE, SOURCE_COLUMN) EXEC (@sql)
    END
ELSE
    BEGIN
    INSERT INTO @t (ORDINAL_POSITION, COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE, SOURCE_COLUMN)
    SELECT
        c.ORDINAL_POSITION
        , c.COLUMN_NAME
        , c.DATA_TYPE
        , c.CHARACTER_MAXIMUM_LENGTH
        , c.IS_NULLABLE
        , c.COLUMN_NAME
    FROM
        INFORMATION_SCHEMA.COLUMNS c
    WHERE
        c.TABLE_SCHEMA = @BaseTableSchema
        AND c.TABLE_NAME = @BaseTableName
    END

END

BEGIN -- FIELD CURSOR --

DECLARE FieldCursor CURSOR FORWARD_ONLY LOCAL READ_ONLY FOR
    SELECT
        c.COLUMN_NAME
        , c.ORDINAL_POSITION
        , c.DATA_TYPE
        , c.CHARACTER_MAXIMUM_LENGTH
        , c.IS_NULLABLE

        , c.COLUMN_EXISTS
        , c.IS_PRIMARY_KEY
        , c.is_identity AS IS_IDENTITY
        , c.is_rowguidcol AS IS_ROWGUIDCOL
        , c.is_computed AS IS_COMPUTED

        , c.LINKED_TABLE_SCHEMA
        , c.LINKED_TABLE_NAME
        , c.LINKED_COLUMN_NAME
    FROM
    (
    SELECT
        c.COLUMN_NAME
        , c.ORDINAL_POSITION
        , c.DATA_TYPE
        , c.CHARACTER_MAXIMUM_LENGTH
        , c.IS_NULLABLE

        , CASE WHEN uc.COLUMN_NAME IS NULL THEN 0 ELSE 1 END COLUMN_EXISTS
        , CASE WHEN ccu.COLUMN_NAME IS NULL THEN 0 ELSE 1 END AS IS_PRIMARY_KEY
        , COALESCE(sc.is_identity, 0) AS is_identity
        , COALESCE(sc.is_rowguidcol, 0) AS is_rowguidcol
        , COALESCE(sc.is_computed, 0) AS is_computed

        , r.LINKED_TABLE_SCHEMA
        , r.LINKED_TABLE_NAME
        , r.LINKED_COLUMN_NAME
    FROM
        @t c

        LEFT OUTER JOIN INFORMATION_SCHEMA.COLUMNS uc ON
            uc.TABLE_SCHEMA = @UnderlyingTableSchema AND uc.TABLE_NAME = @UnderlyingTableName
            AND uc.COLUMN_NAME = c.SOURCE_COLUMN

        LEFT OUTER JOIN sys.columns sc ON sc.[object_id] = OBJECT_ID(@UnderlyingTable) AND sc.name = c.SOURCE_COLUMN

        LEFT OUTER JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc ON
            tc.TABLE_SCHEMA = uc.TABLE_SCHEMA AND tc.TABLE_NAME = uc.TABLE_NAME
            AND tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
        LEFT OUTER JOIN INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE ccu ON
            ccu.TABLE_SCHEMA = uc.TABLE_SCHEMA AND ccu.TABLE_NAME = uc.TABLE_NAME AND ccu.COLUMN_NAME = uc.COLUMN_NAME
            AND tc.CONSTRAINT_SCHEMA = ccu.CONSTRAINT_SCHEMA AND tc.CONSTRAINT_NAME = ccu.CONSTRAINT_NAME

        LEFT OUTER JOIN (
            SELECT
                uc.COLUMN_NAME

                , rcu.TABLE_SCHEMA  AS LINKED_TABLE_SCHEMA
                , rcu.TABLE_NAME    AS LINKED_TABLE_NAME
                , rcu.COLUMN_NAME   AS LINKED_COLUMN_NAME
            FROM
                INFORMATION_SCHEMA.COLUMNS uc

                INNER JOIN INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE ccu ON
                    ccu.TABLE_SCHEMA = uc.TABLE_SCHEMA AND ccu.TABLE_NAME = uc.TABLE_NAME
                    AND ccu.COLUMN_NAME = uc.COLUMN_NAME
                INNER JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc ON
                    tc.TABLE_SCHEMA = uc.TABLE_SCHEMA AND tc.TABLE_NAME = uc.TABLE_NAME
                    AND tc.CONSTRAINT_SCHEMA = ccu.CONSTRAINT_SCHEMA AND tc.CONSTRAINT_NAME = ccu.CONSTRAINT_NAME
                    AND tc.CONSTRAINT_TYPE = 'FOREIGN KEY'

                INNER JOIN INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS rc ON
                    rc.CONSTRAINT_SCHEMA = ccu.CONSTRAINT_SCHEMA AND rc.CONSTRAINT_NAME = ccu.CONSTRAINT_NAME
                INNER JOIN INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE rcu ON
                    rcu.CONSTRAINT_SCHEMA = rc.UNIQUE_CONSTRAINT_SCHEMA AND rcu.CONSTRAINT_NAME = rc.UNIQUE_CONSTRAINT_NAME
            WHERE
                uc.TABLE_SCHEMA = @UnderlyingTableSchema AND uc.TABLE_NAME = @UnderlyingTableName
        ) r ON r.COLUMN_NAME = c.COLUMN_NAME
    ) c
    ORDER BY
        c.ORDINAL_POSITION

END

BEGIN -- OPENING CURSOR --

OPEN FieldCursor
FETCH NEXT FROM FieldCursor
INTO
    @COLUMN_NAME
    , @ORDINAL_POSITION
    , @DATA_TYPE
    , @CHARACTER_MAXIMUM_LENGTH
    , @IS_NULLABLE

    , @COLUMN_EXISTS
    , @IS_PRIMARY_KEY
    , @IS_IDENTITY
    , @IS_ROWGUIDCOL
    , @IS_COMPUTED

    , @LINKED_TABLE_SCHEMA
    , @LINKED_TABLE_NAME
    , @LINKED_COLUMN_NAME

END

BEGIN -- GENERATING FIELDS --

WHILE @@FETCH_STATUS = 0
    BEGIN

    SET @QuotedColumnName = xls.get_friendly_column_name(@COLUMN_NAME)

    IF @DATA_TYPE = 'uniqueidentifier'
        SET @SelectFields = @SelectFields + @Tab + ', CAST(t.' + @QuotedColumnName +' AS varchar(56))'
    ELSE
        SET @SelectFields = @SelectFields + @Tab + ', t.' + @QuotedColumnName

    -- SET @SelectFields = @SelectFields + ' AS ' + @QuotedColumnName

    SET @SelectFields = @SelectFields + @CrLf

    SET @ParameterSQL = @Tab + ', @' + xls.get_escaped_parameter_name(@COLUMN_NAME) + ' ' + @DATA_TYPE

    IF ISNULL(@CHARACTER_MAXIMUM_LENGTH, 0) > 0
        SET @ParameterSQL = @ParameterSQL + '(' + CAST(@CHARACTER_MAXIMUM_LENGTH AS nvarchar(10)) + ')'

    IF @IS_NULLABLE = 'YES'
        SET @ParameterSQL = @ParameterSQL + ' = NULL'

    IF @IS_IDENTITY = 0
        SET @InsertParameters = @InsertParameters + @ParameterSQL + @CrLf

    SET @UpdateParameters = @UpdateParameters + @ParameterSQL + @CrLf

    IF @LINKED_COLUMN_NAME IS NOT NULL
        BEGIN
        SET @JoinNumber = @JoinNumber + 1

        SET @SelectJoin = @SelectJoin
                + @Tab + CASE WHEN @IS_NULLABLE = 'YES' THEN 'LEFT OUTER JOIN ' ELSE 'INNER JOIN ' END
                + CASE WHEN xls.get_escaped_parameter_name(@LINKED_TABLE_SCHEMA) = @LINKED_TABLE_SCHEMA
                    AND xls.get_escaped_parameter_name(@LINKED_TABLE_NAME) = @LINKED_TABLE_NAME
                    THEN @LINKED_TABLE_SCHEMA + '.' + @LINKED_TABLE_NAME
                    ELSE QUOTENAME(@LINKED_TABLE_SCHEMA) + '.' + QUOTENAME(@LINKED_TABLE_NAME)
                    END
                + ' t' + CAST(@JoinNumber AS nvarchar(128))
                + ' ON t' + CAST(@JoinNumber AS nvarchar(128)) + '.' + xls.get_friendly_column_name(@LINKED_COLUMN_NAME)
                + ' = t.' + @QuotedColumnName + @CrLf
        END

    IF NOT (@IS_IDENTITY = 1 OR @IS_COMPUTED = 1) AND @COLUMN_EXISTS = 1
        BEGIN
        SET @InsertFields = @InsertFields + @Tab + ', ' + @QuotedColumnName + @CrLf
        SET @InsertValues = @InsertValues + @Tab + ', @' + xls.get_escaped_parameter_name(@COLUMN_NAME) + @CrLf
        END

    IF @IS_PRIMARY_KEY = 1
        BEGIN
        SET @DeleteParameters = @DeleteParameters + @Tab + ', @' + xls.get_escaped_parameter_name(@COLUMN_NAME) + ' ' + @DATA_TYPE

        IF NOT @CHARACTER_MAXIMUM_LENGTH IS NULL
            SET @DeleteParameters = @DeleteParameters + '(' + CAST(@CHARACTER_MAXIMUM_LENGTH AS varchar(10)) + ')'

        SET @DeleteParameters = @DeleteParameters + @CrLf

        SET @DeleteWhere = @DeleteWhere + @Tab + 'AND ' + @QuotedColumnName + ' = @' + xls.get_escaped_parameter_name(@COLUMN_NAME) + @CrLf
        END
    ELSE IF @COLUMN_EXISTS = 1
        BEGIN
        SET @UpdateFields = @UpdateFields + @Tab + ', ' + @QuotedColumnName + ' = @' + xls.get_escaped_parameter_name(@COLUMN_NAME) + @CrLf
        END

    FETCH NEXT FROM FieldCursor
    INTO
        @COLUMN_NAME
        , @ORDINAL_POSITION
        , @DATA_TYPE
        , @CHARACTER_MAXIMUM_LENGTH
        , @IS_NULLABLE

        , @COLUMN_EXISTS
        , @IS_PRIMARY_KEY
        , @IS_IDENTITY
        , @IS_ROWGUIDCOL
        , @IS_COMPUTED

        , @LINKED_TABLE_SCHEMA
        , @LINKED_TABLE_NAME
        , @LINKED_COLUMN_NAME
    END

CLOSE FieldCursor
DEALLOCATE FieldCursor

END

BEGIN -- FINISH FIELDS --

SET @SelectFields = @Tab + ' ' + SUBSTRING(@SelectFields, 6, LEN(@SelectFields))
SET @InsertFields = @Tab + '(' + SUBSTRING(@InsertFields, 6, LEN(@InsertFields)) + @Tab + ')' + @CrLf
SET @InsertValues = @Tab + '(' + SUBSTRING(@InsertValues, 6, LEN(@InsertValues)) + @Tab + ')' + @CrLf
SET @UpdateFields = @Tab       + SUBSTRING(@UpdateFields, 7, LEN(@UpdateFields))
SET @InsertParameters = @Tab + ' ' + SUBSTRING(@InsertParameters, 6, LEN(@InsertParameters))
SET @UpdateParameters = @Tab + ' ' + SUBSTRING(@UpdateParameters, 6, LEN(@UpdateParameters))
SET @DeleteParameters = @Tab + ' ' + SUBSTRING(@DeleteParameters, 6, LEN(@DeleteParameters))
SET @DeleteWhere = @Tab          + SUBSTRING(@DeleteWhere, 9, LEN(@DeleteWhere))

IF CHARINDEX(@InsertParameters, CHAR(13)) = 0 SET @InsertParameters = SUBSTRING(@InsertParameters, 3, LEN(@InsertParameters))
IF CHARINDEX(@UpdateParameters, CHAR(13)) = 0 SET @UpdateParameters = SUBSTRING(@UpdateParameters, 3, LEN(@UpdateParameters))
IF CHARINDEX(@DeleteParameters, CHAR(13)) = 0 SET @DeleteParameters = SUBSTRING(@DeleteParameters, 3, LEN(@DeleteParameters))

END

BEGIN -- CHANGE HANDLER --

DECLARE @ChangeWhere nvarchar(max) = ''

SET @ChangeWhere = STUFF((
        SELECT
            ' AND ' + xls.get_friendly_column_name(c.SOURCE_COLUMN) + ' = @' + xls.get_escaped_parameter_name(c.COLUMN_NAME)
        FROM
            INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
            INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu ON
                kcu.CONSTRAINT_SCHEMA = tc.CONSTRAINT_SCHEMA AND kcu.CONSTRAINT_NAME = tc.CONSTRAINT_NAME
            INNER JOIN @t c ON c.SOURCE_COLUMN = kcu.COLUMN_NAME
        WHERE
            tc.TABLE_SCHEMA = @UnderlyingTableSchema
            AND tc.TABLE_NAME = @UnderlyingTableName
            AND tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
        ORDER BY
            c.ORDINAL_POSITION
        FOR XML PATH('')
    ), 1, 5, '')

DECLARE @HasDates bit
DECLARE @HasNumbers bit

SELECT
    @HasDates =    MAX(CASE WHEN c.DATETIME_PRECISION IS NULL THEN 0 ELSE 1 END)
    , @HasNumbers = MAX(CASE WHEN c.NUMERIC_PRECISION IS NULL THEN 0 WHEN c.DATA_TYPE IN ('bit') THEN 1 ELSE 1 END)
FROM
    INFORMATION_SCHEMA.COLUMNS c
WHERE
    c.TABLE_SCHEMA = @UnderlyingTableSchema AND c.TABLE_NAME = @UnderlyingTableName
    OR (c.TABLE_SCHEMA = @BaseTableSchema AND c.TABLE_NAME = @BaseTableName)

DECLARE @termColumnName nvarchar(128) = '@ColumnName'
DECLARE @termCellValue nvarchar(128) = '@CellValue'
DECLARE @termCellNumberValue nvarchar(128) = '@CellNumberValue'
DECLARE @termCellDateTimeValue nvarchar(128) = '@CellDateTimeValue'
DECLARE @termChangedCellAction nvarchar(128) = '@ChangedCellAction'
DECLARE @termChangedCellCount nvarchar(128) = '@ChangedCellCount'
DECLARE @termChangedCellIndex nvarchar(128) = '@ChangedCellIndex'
DECLARE @termDataLanguage nvarchar(128) = '@DataLanguage'

IF @MySqlStyle = 1
    BEGIN
    SET @termColumnName = '@column_name'
    SET @termCellValue = '@cell_value'
    SET @termCellNumberValue = '@cell_number_value'
    SET @termCellDateTimeValue = '@cell_datetime_value'
    SET @termChangedCellAction = '@changed_cell_ction'
    SET @termChangedCellCount = '@changed_cell_count'
    SET @termChangedCellIndex = '@changed_cell_index'
    SET @termDataLanguage = '@data_language'
    END

DECLARE @ContextParameters nvarchar(max) =
    '    ' + @termColumnName + ' nvarchar(128) = NULL' + @CrLf
    + '    , ' + @termCellValue + ' nvarchar(255) = NULL' + @CrLf
    + CASE WHEN @HasNumbers = 1 THEN '    , ' + @termCellNumberValue + ' float = NULL' + @CrLf ELSE '' END
    + CASE WHEN @HasDates = 1 THEN '    , ' + @termCellDateTimeValue + ' datetime = NULL' + @CrLf ELSE '' END
    + '    , ' + @termChangedCellAction + ' nvarchar(255) = NULL' + @CrLf
    + '    , ' + @termChangedCellCount + ' int = NULL' + @CrLf
    + '    , ' + @termChangedCellIndex + ' int = NULL' + @CrLf
    + '    , ' + @termDataLanguage + ' varchar(10) = NULL' + @CrLf

DECLARE @ChangeParameters nvarchar(max)

SELECT @ChangeParameters = STUFF((
    SELECT
        @CrLf + '    , @' + xls.get_escaped_parameter_name(c.COLUMN_NAME)
        + ' ' + c.DATA_TYPE
        + CASE WHEN c.CHARACTER_MAXIMUM_LENGTH IS NULL THEN '' ELSE '(' + CAST(c.CHARACTER_MAXIMUM_LENGTH AS varchar(5)) + ')' END
        + ' = NULL'
    FROM
        @t c
    WHERE
        NOT c.DATA_TYPE IN ('geography', 'geometry', 'image')
    ORDER BY
        c.ORDINAL_POSITION
    FOR XML PATH(''), TYPE).value('.', 'nvarchar(max)'), 1, 2, '')

DECLARE @ChangeBody nvarchar(max)

SELECT @ChangeBody = COALESCE(STUFF((
    SELECT
        'ELSE IF ' + @termColumnName + ' = N''' + REPLACE(c.COLUMN_NAME, '''', '''''') + '''' + @CrLf
        + @Tab
        + CASE
            WHEN sc.name IS NULL THEN 'RETURN'
            WHEN sc.is_identity = 1 OR sc.is_computed = 1 THEN 'RETURN'
            WHEN @ChangeWhere IS NULL THEN 'RETURN'
            ELSE
                'UPDATE ' + @UnderlyingTable
                + ' SET '
                + xls.get_friendly_column_name(c.COLUMN_NAME)
                + ' = '
                + CASE
                    WHEN t.name IN ('date', 'time', 'datetime', 'datetime2', 'smalldatetime', 'datetimeoffset') THEN @termCellDateTimeValue
                    WHEN t.name IN ('bit', 'tinyint', 'smallint', 'int', 'bigint', 'real', 'float', 'money', 'smallmoney', 'decimal', 'numeric') THEN @termCellNumberValue
                    ELSE @termCellValue
                    END
                + ' WHERE '
                + @ChangeWhere
            END
        + @CrLf + @CrLf
    FROM
        @t c
        LEFT OUTER JOIN sys.columns sc
            ON sc.[object_id] = OBJECT_ID(@UnderlyingTable) AND sc.name = c.SOURCE_COLUMN
        LEFT OUTER JOIN sys.types t ON t.user_type_id = sc.user_type_id
    WHERE
        NOT c.DATA_TYPE IN ('timestamp', 'geography', 'geometry', 'image')
    ORDER BY
        c.ORDINAL_POSITION
    FOR XML PATH(''), TYPE).value('.', 'nvarchar(max)'), 1, 5, ''), '')

END

BEGIN -- PROCEDURE NAMES --

DECLARE @SelectProcedureName nvarchar(255)
DECLARE @InsertProcedureName nvarchar(255)
DECLARE @UpdateProcedureName nvarchar(255)
DECLARE @DeleteProcedureName nvarchar(255)
DECLARE @ChangeProcedureName nvarchar(255)

IF @SelectObjectName IS NULL
    BEGIN
    IF OBJECT_ID(@BaseTable, 'V') IS NOT NULL
        AND @GenerateSelectObject = 0
        BEGIN
        SET @SelectObjectSchema = @BaseTableSchema
        SET @SelectObjectName = @BaseTableName
        END
    ELSE IF OBJECT_ID(@BaseTable, 'P') IS NOT NULL
        AND @GenerateSelectObject = 0
        BEGIN
        SET @SelectObjectSchema = @BaseTableSchema
        SET @SelectObjectName = @BaseTableName
        END
    ELSE
        BEGIN
        IF LOWER(@BaseTableName) = @BaseTableName
            SET @SelectObjectName = CASE WHEN @GenerateSelectAsView = 1 THEN 'view_' ELSE 'usp_' END + @BaseTableName
        ELSE
            SET @SelectObjectName = CASE WHEN @GenerateSelectAsView = 1 THEN 'view' ELSE 'usp' END + @BaseTableName
        END
    END

IF xls.get_escaped_parameter_name(@SelectObjectSchema) = @SelectObjectSchema
    AND xls.get_escaped_parameter_name(@SelectObjectName) = @SelectObjectName
    BEGIN
    SET @SelectProcedureName = @SelectObjectSchema + '.' + @SelectObjectName
    SET @InsertProcedureName = @SelectObjectSchema + '.' + @SelectObjectName + '_insert'
    SET @UpdateProcedureName = @SelectObjectSchema + '.' + @SelectObjectName + '_update'
    SET @DeleteProcedureName = @SelectObjectSchema + '.' + @SelectObjectName + '_delete'
    SET @ChangeProcedureName = @SelectObjectSchema + '.' + @SelectObjectName + '_change'
    END
ELSE
    BEGIN
    SET @SelectProcedureName = QUOTENAME(@SelectObjectSchema) + '.' + QUOTENAME(@SelectObjectName)
    SET @InsertProcedureName = QUOTENAME(@SelectObjectSchema) + '.' + QUOTENAME(@SelectObjectName + '_insert')
    SET @UpdateProcedureName = QUOTENAME(@SelectObjectSchema) + '.' + QUOTENAME(@SelectObjectName + '_update')
    SET @DeleteProcedureName = QUOTENAME(@SelectObjectSchema) + '.' + QUOTENAME(@SelectObjectName + '_delete')
    SET @ChangeProcedureName = QUOTENAME(@SelectObjectSchema) + '.' + QUOTENAME(@SelectObjectName + '_change')
    END

END

BEGIN -- HEADERS --

DECLARE @SelectHeaderText nvarchar(max) = '
-- =============================================
-- Author:      ' + xls.get_translated_string('<Author>', 'en') + '
-- Release:     ' + xls.get_translated_string('<Release>', 'en') + ', ' + CONVERT(char(10), GETDATE(), 120)

DECLARE @ChangeHeaderText nvarchar(max) = @SelectHeaderText + '
-- Description: The procedure processes cell changes of ' + @SelectObjectSchema + '.' + @SelectObjectName + '
-- =============================================

'

SET @SelectHeaderText = @SelectHeaderText + '
-- Description: The procedure selects data from ' + @BaseTableSchema + '.' + @BaseTableName + '
-- =============================================

'

DECLARE @InsertHeaderText nvarchar(max) = REPLACE(@SelectHeaderText, 'selects data from', 'inserts data into')
DECLARE @UpdateHeaderText nvarchar(max) = REPLACE(@SelectHeaderText, 'selects data from', 'updates data of')
DECLARE @DeleteHeaderText nvarchar(max) = REPLACE(@SelectHeaderText, 'selects data from', 'deleted data from')

IF @GenerateSelectAsView = 1
    SET @SelectHeaderText = REPLACE(@SelectHeaderText, 'The procedure', 'The view')

END

BEGIN -- PROCEDURE DEFINITIONS --

DECLARE @DeleteSelectSQL nvarchar(max) =
    CASE WHEN @GenerateSelectAsView = 1 THEN
    'IF OBJECT_ID(N''' + REPLACE(@SelectProcedureName, '''', '''''') + ''', ''V'') IS NOT NULL' + @CrLf
    + @Tab + 'DROP VIEW '+ @SelectProcedureName + @CrLf
    ELSE
    'IF OBJECT_ID(N''' + REPLACE(@SelectProcedureName, '''', '''''') + ''', ''P'') IS NOT NULL' + @CrLf
    + @Tab + 'DROP PROCEDURE '+ @SelectProcedureName + @CrLf
    END

DECLARE @DeleteInsertSQL nvarchar(max) =
    'IF OBJECT_ID(N''' + REPLACE(@InsertProcedureName, '''', '''''') + ''', ''P'') IS NOT NULL' + @CrLf
    + @Tab + 'DROP PROCEDURE '+ @InsertProcedureName + @CrLf

DECLARE @DeleteUpdateSQL nvarchar(max) =
    'IF OBJECT_ID(N''' + REPLACE(@UpdateProcedureName, '''', '''''') + ''', ''P'') IS NOT NULL' + @CrLf
    + @Tab + 'DROP PROCEDURE '+ @UpdateProcedureName + @CrLf

DECLARE @DeleteDeleteSQL nvarchar(max) =
    'IF OBJECT_ID(N''' + REPLACE(@DeleteProcedureName, '''', '''''') + ''', ''P'') IS NOT NULL' + @CrLf
    + @Tab + 'DROP PROCEDURE '+ @DeleteProcedureName + @CrLf

DECLARE @DeleteChangeSQL nvarchar(max) =
    'IF OBJECT_ID(N''' + REPLACE(@ChangeProcedureName, '''', '''''') + ''', ''P'') IS NOT NULL' + @CrLf
    + @Tab + 'DROP PROCEDURE '+ @ChangeProcedureName + @CrLf

DECLARE @SelectProcedureSQL nvarchar(max) = ''
        + @SelectHeaderText
        + CASE WHEN @GenerateSelectAsView = 1 THEN
          'CREATE' + ' VIEW ' + @SelectProcedureName + @CrLf
          ELSE
          'CREATE' + ' PROCEDURE ' + @SelectProcedureName + @CrLf
          END
        + 'AS' + @CrLf
        + CASE WHEN @GenerateSelectAsView = 1 THEN '' ELSE 'BEGIN' + @CrLf + @CrLf + 'SET NOCOUNT ON' + @CrLf END
        + @CrLf
        + 'SELECT' + @CrLf
        + @SelectFields
        + 'FROM'+ @CrLf
        + @Tab + @BaseTable + ' t' + @CrLf
        + @SelectJoin
        + @CrLf
        + CASE WHEN @GenerateSelectAsView = 1 THEN '' ELSE 'END' + @CrLf END
        + @CrLf

DECLARE @InsertProcedureSQL nvarchar(max) = ''
        + @InsertHeaderText
        + 'CREATE' + ' PROCEDURE ' + @InsertProcedureName + @CrLf
        + @InsertParameters
        + 'AS' + @CrLf
        + 'BEGIN' + @CrLf
        + @CrLf
        + 'INSERT INTO ' + @UnderlyingTable + @CrLf
        + @InsertFields
        + 'VALUES' + @CrLf
        + @InsertValues
        + @CrLf
        + 'END' + @CrLf + @CrLf

DECLARE @UpdateProcedureSQL nvarchar(max) = ''
        + @UpdateHeaderText
        + 'CREATE' + ' PROCEDURE ' + @UpdateProcedureName + @CrLf
        + @UpdateParameters
        + 'AS' + @CrLf
        + 'BEGIN' + @CrLf
        + @CrLf
        + 'UPDATE ' + @UnderlyingTable + @CrLf
        + 'SET' + @CrLf
        + @UpdateFields
        + 'WHERE' + @CrLf
        + @DeleteWhere
        + @CrLf
        + 'END' + @CrLf + @CrLf

DECLARE @DeleteProcedureSQL nvarchar(max) = ''
        + @DeleteHeaderText
        + 'CREATE' + ' PROCEDURE ' + @DeleteProcedureName + @CrLf
        + @DeleteParameters
        + 'AS' + @CrLf
        + 'BEGIN' + @CrLf
        + @CrLf
        + 'DELETE FROM ' + @UnderlyingTable + @CrLf
        + 'WHERE' + @CrLf
        + @DeleteWhere
        + @CrLf
        + 'END' + @CrLf + @CrLf

DECLARE @ChangeProcedureSql nvarchar(max) = ''
        + @ChangeHeaderText
        + 'CREATE' + ' PROCEDURE ' + @ChangeProcedureName + @CrLf
        + @ContextParameters
        + @ChangeParameters + @CrLf
        + 'AS' + @CrLf
        + 'BEGIN' + @CrLf
        + @CrLf
        + 'SET NOCOUNT ON' + @CrLf
        + @CrLf
        + @ChangeBody
        + 'END' + @CrLf
        + @CrLf

END

BEGIN -- HELP --

DECLARE @GoLine nvarchar(10) = 'GO' + @CrLf + @CrLf

DECLARE @Help nvarchar(max) =
    CASE WHEN xls.get_translated_string('<Author>', 'en') = '<Author>' THEN
        '-- You may define the <Author> and <Release> values in the xls.translations table:' + @CrLf
        + @CrLf
        + '-- TABLE_SCHEMA TABLE_NAME COLUMN_NAME LANGUAGE_NAME TRANSLATED_NAME' + @CrLf
        + '-- ------------ ---------- ----------- ------------- ---------------' + @CrLf
        + '-- xls          strings    <Author>    en            <Your value>' + @CrLf
        + '-- xls          strings    <Release>   en            <Your value>' + @CrLf
        + @CrLf
        ELSE '' END

    + CASE WHEN @GenerateEditProcedures = 1 THEN
        + '-- The SaveToDB add-in uses the insert, update, and delete procedures for ' +  @SelectProcedureName + @CrLf
        + '-- automatically due to the same name with the special _insert, _update, and _delete suffixes.'+ @CrLf
        + @CrLf
        + '-- You may use these edit procedures with appropriate objects using the configuration like this:'+ @CrLf
        + @CrLf
        + '-- INSERT INTO xls.objects (TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE, INSERT_OBJECT, UPDATE_OBJECT, DELETE_OBJECT)' + @CrLf
        + '--     VALUES('
        + 'N''' + REPLACE(@SelectObjectSchema, '''', '''''') + ''', N''' + @SelectObjectName + ''', '''
        +  CASE WHEN @GenerateSelectAsView = 1 THEN 'VIEW' ELSE 'PROCEDURE' END + ''', '
        + 'N''' + REPLACE(@InsertProcedureName, '''', '''''') + ''', N''' + @UpdateProcedureName + ''', N''' + @DeleteProcedureName + ''')' + @CrLf
        + @CrLf
        ELSE '' END

    + CASE WHEN @GenerateChangeHandler = 1 THEN
        + '-- The SaveToDB add-in attaches ' + @ChangeProcedureName + ' as a change handler for' + @CrLf
        + '-- ' +  @SelectProcedureName + ' automatically due to the same name with the _change suffix.'+ @CrLf
        + @CrLf
        + '-- You may attach it to any object using the configuration like this:'+ @CrLf
        + @CrLf
        + '-- INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE)' + @CrLf
        + '--     VALUES('
        + 'N''' + REPLACE(@SelectObjectSchema, '''', '''''') + ''', N''' + @SelectObjectName + ''', ''Change'', '
        + 'N''' + REPLACE(@SelectObjectSchema, '''', '''''') + ''', N''' + @BaseTableName + '_change'', ''PROCEDURE'')' + @CrLf
        + @CrLf
        ELSE '' END

    + @GoLine

END

BEGIN -- EXECUTE GENERATED CODES --

IF @SelectCommands IS NULL AND @PrintCommands IS NULL SET @SelectCommands = 1

IF @PrintCommands = 1
    BEGIN

    RAISERROR(@Help, 0, 1) WITH NOWAIT

    IF @GenerateSelectObject = 1
        BEGIN
        RAISERROR(@DeleteSelectSQL, 0, 1) WITH NOWAIT
        RAISERROR(@GoLine, 0, 1) WITH NOWAIT
        END

    IF @GenerateEditProcedures = 1
        BEGIN
        RAISERROR(@DeleteInsertSQL, 0, 1) WITH NOWAIT
        RAISERROR(@GoLine, 0, 1) WITH NOWAIT
        RAISERROR(@DeleteUpdateSQL, 0, 1) WITH NOWAIT
        RAISERROR(@GoLine, 0, 1) WITH NOWAIT
        RAISERROR(@DeleteDeleteSQL, 0, 1) WITH NOWAIT
        RAISERROR(@GoLine, 0, 1) WITH NOWAIT
        END

    IF @GenerateChangeHandler = 1
        BEGIN
        RAISERROR(@DeleteChangeSQL, 0, 1) WITH NOWAIT
        RAISERROR(@GoLine, 0, 1) WITH NOWAIT
        END

    IF @GenerateSelectObject = 1
        BEGIN
        RAISERROR(@SelectProcedureSQL, 0, 1) WITH NOWAIT
        RAISERROR(@GoLine, 0, 1) WITH NOWAIT
        END

    IF @GenerateEditProcedures = 1
        BEGIN
        RAISERROR(@InsertProcedureSQL, 0, 1) WITH NOWAIT
        RAISERROR(@GoLine, 0, 1) WITH NOWAIT
        RAISERROR(@UpdateProcedureSQL, 0, 1) WITH NOWAIT
        RAISERROR(@GoLine, 0, 1) WITH NOWAIT
        RAISERROR(@DeleteProcedureSQL, 0, 1) WITH NOWAIT
        RAISERROR(@GoLine, 0, 1) WITH NOWAIT
        END

    IF @GenerateChangeHandler = 1
        BEGIN
        RAISERROR(@ChangeProcedureSQL, 0, 1) WITH NOWAIT
        RAISERROR(@GoLine, 0, 1) WITH NOWAIT
        END
    END

IF @ExecuteScript = 1
    BEGIN
    IF @RecreateProceduresIfExist = 1
        BEGIN
        IF @GenerateSelectObject = 1
            EXEC (@DeleteSelectSQL)
        IF @GenerateEditProcedures = 1
            BEGIN
            EXEC (@DeleteInsertSQL)
            EXEC (@DeleteUpdateSQL)
            EXEC (@DeleteDeleteSQL)
            END
        IF @GenerateChangeHandler = 1
            EXEC (@DeleteChangeSQL)
        END

    IF @GenerateSelectObject = 1 AND OBJECT_ID(@SelectProcedureName, 'P') IS NULL
        EXEC (@SelectProcedureSQL)

    IF @GenerateEditProcedures = 1 AND OBJECT_ID(@InsertProcedureName, 'P') IS NULL
        EXEC (@InsertProcedureSQL)
    IF @GenerateEditProcedures = 1 AND OBJECT_ID(@UpdateProcedureName, 'P') IS NULL
        EXEC (@UpdateProcedureSQL)
    IF @GenerateEditProcedures = 1 AND OBJECT_ID(@DeleteProcedureName, 'P') IS NULL
        EXEC (@DeleteProcedureSQL)

    IF @GenerateChangeHandler = 1 AND OBJECT_ID(@ChangeProcedureName, 'P') IS NULL
        EXEC (@ChangeProcedureSQL)

    IF @SelectCommands = 1
        SELECT 'Created' AS [message]
    END

ELSE IF @SelectCommands = 1
    BEGIN
    SELECT
        @Help
        + CASE WHEN @GenerateSelectObject = 1 THEN
            @DeleteSelectSQL
            + @GoLine
            ELSE '' END
        + CASE WHEN @GenerateEditProcedures = 1 THEN
            @DeleteInsertSQL
            + @GoLine
            + @DeleteUpdateSQL
            + @GoLine
            + @DeleteDeleteSQL
            + @GoLine
            ELSE '' END
        + CASE WHEN @GenerateChangeHandler = 1 THEN
            @DeleteChangeSQL
            + @GoLine
            ELSE '' END
        + CASE WHEN @GenerateSelectObject = 1 THEN
            @SelectProcedureSQL
            + @GoLine
            ELSE '' END
        + CASE WHEN @GenerateEditProcedures = 1 THEN
            @InsertProcedureSQL
            + @GoLine
            + @UpdateProcedureSQL
            + @GoLine
            + @DeleteProcedureSQL
            + @GoLine
            ELSE '' END
        + CASE WHEN @GenerateChangeHandler = 1 THEN
            @ChangeProcedureSQL
            + @GoLine
            ELSE '' END
        AS [message]
    END

END

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: The procedure creates or drops constraints
-- =============================================

ALTER PROCEDURE [xls].[xl_actions_generate_constraints]
    @Drop bit = 0
    , @Create bit = 0
    , @ConstraintType tinyint = NULL
    , @SCHEMA nvarchar(128) = NULL
    , @TABLE nvarchar(128) = NULL
    , @COLUMN nvarchar(128) = NULL
    , @REFERENTIAL_SCHEMA nvarchar(128) = NULL
    , @REFERENTIAL_NAME nvarchar(128) = NULL
    , @REFERENTIAL_COLUMN nvarchar(128) = NULL
    , @ON_UPDATE nvarchar(128) = NULL
    , @ON_DELETE nvarchar(128) = NULL
    , @CONSTRAINT nvarchar(128) = NULL
    , @ExecuteScript bit = 0
    , @DataLanguage varchar(10) = NULL
    , @SelectCommands bit = NULL
    , @PrintCommands bit = NULL
AS
BEGIN

SET NOCOUNT ON

IF @ExecuteScript IS NULL SET @ExecuteScript = 0
IF @DataLanguage IS NULL SET @DataLanguage = 'en'

DECLARE @message nvarchar(max)

DECLARE @Tab char(4) = REPLICATE(' ', 4)
DECLARE @CrLf char(2) = CHAR(13) + CHAR(10)

IF @Drop IS NULL AND @Create IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage)
        + ' @Create:' + @CrLf + @CrLf
        + '1 - CREATE' + @CrLf
        + '0 - skip' + @CrLf
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF @Drop IS NULL AND @Create = 0
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage)
        + ' @Drop:' + @CrLf + @CrLf
        + '1 - DROP' + @CrLf
        + '0 - skip' + @CrLf
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF @Drop = 0 AND @Create = 0
    RETURN

IF @ConstraintType NOT IN (1, 2, 3, 4)
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage)
        + ' @ConstraintType:' + @CrLf + @CrLf
        + '1 - PRIMARY KEY' + @CrLf
        + '2 - UNIQUE' + @CrLf
        + '3 - INDEX' + @CrLf
        + '4 - FOREIGN KEY' + @CrLf
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF @SCHEMA IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage) + ' @SCHEMA'
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF @TABLE IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage) + ' @TABLE'
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF @COLUMN IS NULL AND @Create = 1
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage) + ' @COLUMN'
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF @ConstraintType = 4 AND @Create = 1 AND @REFERENTIAL_SCHEMA IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage) + ' @REFERENTIAL_SCHEMA'
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF @ConstraintType = 4 AND @Create = 1 AND @REFERENTIAL_NAME IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage) + ' @REFERENTIAL_NAME'
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF @ConstraintType = 4 AND @Create = 1 AND @REFERENTIAL_COLUMN IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage) + ' @REFERENTIAL_COLUMN'
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF @Drop = 1 AND @CONSTRAINT IS NULL
    BEGIN
    SET @message = xls.get_translated_string('Specify', @DataLanguage) + ' @CONSTRAINT'
    RAISERROR(@message, 11, 0);
    RETURN
    END

IF @CONSTRAINT IS NULL
    BEGIN
    IF @ConstraintType = 1
        SET @CONSTRAINT = 'PK_' + @TABLE + '_' + @SCHEMA
    ELSE IF @ConstraintType = 2
        SET @CONSTRAINT = 'IX_' + @TABLE + '_' + @COLUMN + '_' + @SCHEMA
    ELSE IF @ConstraintType = 3
        SET @CONSTRAINT = 'IX_' + @TABLE + '_' + @COLUMN + '_' + @SCHEMA
    ELSE IF @ConstraintType = 4
        SET @CONSTRAINT = 'FK_' + @TABLE + '_' + @REFERENTIAL_NAME + '_' + @COLUMN + '_' + @SCHEMA
    END

DECLARE @sql nvarchar(max)

SET @sql = CASE WHEN @Drop = 0 THEN '' ELSE
    CASE WHEN @ConstraintType = 3 THEN
        'DROP INDEX ' + QUOTENAME(REPLACE(@CONSTRAINT, '''', ''''''))
        + ' ON ' + QUOTENAME(REPLACE(@SCHEMA, '''', ''''''))+ '.' + QUOTENAME(REPLACE(@TABLE, '''', '''''')) + ';' + @CrLf
        ELSE
        'ALTER TABLE ' + QUOTENAME(REPLACE(@SCHEMA, '''', ''''''))+ '.' + QUOTENAME(REPLACE(@TABLE, '''', ''''''))
        + ' DROP CONSTRAINT ' + QUOTENAME(REPLACE(@CONSTRAINT, '''', '''''')) + ';' + @CrLf
        END
    END

    + CASE WHEN @Create = 0 THEN '' ELSE
    CASE WHEN @ConstraintType = 3 THEN
        'CREATE INDEX ' + QUOTENAME(REPLACE(@CONSTRAINT, '''', ''''''))
        + ' ON ' + QUOTENAME(REPLACE(@SCHEMA, '''', ''''''))+ '.' + QUOTENAME(REPLACE(@TABLE, '''', ''''''))
        + ' (' + QUOTENAME(REPLACE(@COLUMN, '''', '''''')) + ');' + @CrLf
        ELSE
        'ALTER TABLE ' + QUOTENAME(REPLACE(@SCHEMA, '''', ''''''))+ '.' + QUOTENAME(REPLACE(@TABLE, '''', ''''''))
        + ' ADD CONSTRAINT ' + QUOTENAME(REPLACE(@CONSTRAINT, '''', ''''''))
        + ' ' + CASE @ConstraintType
                WHEN 1 THEN 'PRIMARY KEY'
                WHEN 2 THEN 'UNIQUE'
                WHEN 4 THEN 'FOREIGN KEY'
                ELSE '?'
                END
        + ' (' + QUOTENAME(REPLACE(@COLUMN, '''', '''''')) + ')'
        + CASE WHEN @ConstraintType = 4 THEN
            + ' REFERENCES ' + QUOTENAME(REPLACE(@REFERENTIAL_SCHEMA, '''', ''''''))+ '.' + QUOTENAME(REPLACE(@REFERENTIAL_NAME, '''', ''''''))
            + ' (' + QUOTENAME(REPLACE(@REFERENTIAL_COLUMN, '''', '''''')) + ')'
            + CASE UPPER(@ON_UPDATE)
                    WHEN 'CASCADE' THEN ' ON UPDATE CASCADE'
                    WHEN 'SET NULL' THEN ' ON UPDATE SET NULL'
                    WHEN 'SET DEFAULT' THEN ' ON UPDATE SET NULL'
                    ELSE ''
                END
            + CASE UPPER(@ON_DELETE)
                    WHEN 'CASCADE' THEN ' ON DELETE CASCADE'
                    WHEN 'SET NULL' THEN ' ON DELETE SET NULL'
                    WHEN 'SET DEFAULT' THEN ' ON DELETE SET NULL'
                    ELSE ''
                END
            ELSE ''
            END
        + ';' + @CrLf
        END
        END

IF @SelectCommands IS NULL AND @PrintCommands IS NULL SET @SelectCommands = 1

IF @PrintCommands = 1
    BEGIN
    PRINT @sql + @CrLf + 'GO' + @CrLf + @CrLf
    END

IF @ExecuteScript = 1
    BEGIN
    EXEC (@sql)

    IF @SelectCommands = 1
        IF @Drop = 1 AND @Create = 1
            SELECT @sql + @CrLf + xls.get_translated_string('Dropped and created', @DataLanguage) AS [message]
        ELSE IF @Drop = 1
            SELECT @sql + @CrLf + xls.get_translated_string('Dropped', @DataLanguage) AS [message]
        ELSE
            SELECT @sql + @CrLf + xls.get_translated_string('Created', @DataLanguage) AS [message]
    END

ELSE IF @SelectCommands = 1
    BEGIN
    SELECT
        @sql
        + 'GO' + @CrLf + @CrLf
        AS [message]
    END

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: Exports SaveToDB Framework data
-- =============================================

ALTER PROCEDURE [xls].[xl_export_settings]
    @part tinyint = NULL
    , @as_exec_import bit = 0
    , @sort_by_names bit = 0
    , @schema nvarchar(128) = NULL
    , @language varchar(10) = NULL
    , @use_go bit = 0
AS
BEGIN

SET NOCOUNT ON;

IF @as_exec_import IS NULL SET @as_exec_import = 0

DECLARE @app_schema_only bit = 0

IF @schema = 'x'
    BEGIN
    SET @schema = NULL
    SET @app_schema_only = 1
    END

DECLARE @GoLine nvarchar(10) = CASE WHEN @use_go = 1 THEN 'GO' + CHAR(13) + CHAR(10) ELSE '' END

DECLARE @cmd0 nvarchar(1024), @cmd1 nvarchar(1024), @cmd2 nvarchar(1024), @cmd3 nvarchar(1024), @cmd4 nvarchar(1024), @cmd5 nvarchar(1024)

IF @as_exec_import = 1
    BEGIN
    SET @cmd0 = ';'
    SET @cmd1 = 'EXEC xls.xl_import_objects '
    SET @cmd2 = 'EXEC xls.xl_import_handlers '
    SET @cmd3 = 'EXEC xls.xl_import_translations '
    SET @cmd4 = 'EXEC xls.xl_import_formats '
    SET @cmd5 = 'EXEC xls.xl_import_workbooks '
    END
ELSE
    BEGIN
    SET @cmd0 = ');'
    SET @cmd1 = 'INSERT INTO xls.objects (TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE, TABLE_CODE, INSERT_OBJECT, UPDATE_OBJECT, DELETE_OBJECT) VALUES ('
    SET @cmd2 = 'INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS) VALUES ('
    SET @cmd3 = 'INSERT INTO xls.translations (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, LANGUAGE_NAME, TRANSLATED_NAME, TRANSLATED_DESC, TRANSLATED_COMMENT) VALUES ('
    SET @cmd4 = 'INSERT INTO xls.formats (TABLE_SCHEMA, TABLE_NAME, TABLE_EXCEL_FORMAT_XML) VALUES ('
    SET @cmd5 = 'INSERT INTO xls.workbooks (NAME, TEMPLATE, DEFINITION, TABLE_SCHEMA) VALUES ('
    END

SELECT
    t.command
FROM
    (
SELECT
    1 AS part
    , CASE WHEN @sort_by_names = 1 THEN ROW_NUMBER() OVER(ORDER BY TABLE_SCHEMA, TABLE_NAME) ELSE ID END AS ID
    , @cmd1
           + CASE WHEN TABLE_SCHEMA              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_SCHEMA, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_NAME                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_TYPE                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_TYPE, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_CODE                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_CODE, '''', '''''') + '''' END
    + ', ' + CASE WHEN INSERT_OBJECT             IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(INSERT_OBJECT, '''', '''''') + '''' END
    + ', ' + CASE WHEN UPDATE_OBJECT             IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(UPDATE_OBJECT, '''', '''''') + '''' END
    + ', ' + CASE WHEN DELETE_OBJECT             IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(DELETE_OBJECT, '''', '''''') + '''' END
    + @cmd0 AS command
FROM
    xls.objects
WHERE
    TABLE_SCHEMA LIKE COALESCE(@schema, TABLE_SCHEMA)
    AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('xls'))

UNION ALL SELECT 2 AS part, -1 AS ID, @GoLine AS command
    WHERE EXISTS(
        SELECT
            ID
        FROM
            xls.objects
        WHERE
            TABLE_SCHEMA LIKE COALESCE(@schema, TABLE_SCHEMA)
            AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('xls'))
    )

UNION ALL
SELECT
    2 AS part
    , CASE WHEN @sort_by_names = 1 THEN ROW_NUMBER() OVER(ORDER BY TABLE_SCHEMA, TABLE_NAME, EVENT_NAME, CASE WHEN COLUMN_NAME IS NULL THEN 0 ELSE 1 END, COLUMN_NAME, HANDLER_SCHEMA, HANDLER_NAME, MENU_ORDER) ELSE ID END AS ID
    , @cmd2
           + CASE WHEN TABLE_SCHEMA              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_SCHEMA, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_NAME                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN COLUMN_NAME               IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(COLUMN_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN EVENT_NAME                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(EVENT_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN HANDLER_SCHEMA            IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(HANDLER_SCHEMA, '''', '''''') + '''' END
    + ', ' + CASE WHEN HANDLER_NAME              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(HANDLER_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN HANDLER_TYPE              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(HANDLER_TYPE, '''', '''''') + '''' END
    + ', ' + CASE WHEN HANDLER_CODE              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(HANDLER_CODE, '''', '''''') + '''' END
    + ', ' + CASE WHEN TARGET_WORKSHEET          IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TARGET_WORKSHEET, '''', '''''') + '''' END
    + ', ' + CASE WHEN MENU_ORDER                IS NULL THEN 'NULL' ELSE CAST(MENU_ORDER AS nvarchar(128))  END
    + ', ' + CASE WHEN EDIT_PARAMETERS           IS NULL THEN 'NULL' ELSE CAST(EDIT_PARAMETERS AS nvarchar(128))  END
    + @cmd0 AS command
FROM
    xls.handlers
WHERE
    TABLE_SCHEMA LIKE COALESCE(@schema, TABLE_SCHEMA)
    AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('xls'))

UNION ALL SELECT 3 AS part, -1 AS ID, @GoLine AS command
    WHERE EXISTS(
        SELECT
            ID
        FROM
            xls.handlers
        WHERE
            TABLE_SCHEMA LIKE COALESCE(@schema, TABLE_SCHEMA)
            AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('xls'))
    )

UNION ALL
SELECT
    3 AS part
    , CASE WHEN @sort_by_names = 1 THEN ROW_NUMBER() OVER(ORDER BY LANGUAGE_NAME, CASE WHEN COLUMN_NAME IS NULL THEN 0 ELSE 1 END, TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME) ELSE ID END AS ID
    , @cmd3
           + CASE WHEN TABLE_SCHEMA              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_SCHEMA, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_NAME                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN COLUMN_NAME               IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(COLUMN_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN LANGUAGE_NAME             IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(LANGUAGE_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN TRANSLATED_NAME           IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TRANSLATED_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN TRANSLATED_DESC           IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TRANSLATED_DESC, '''', '''''') + '''' END
    + ', ' + CASE WHEN TRANSLATED_COMMENT        IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TRANSLATED_COMMENT, '''', '''''') + '''' END
    + @cmd0 AS command
FROM
    xls.translations
WHERE
    TABLE_SCHEMA LIKE COALESCE(@schema, TABLE_SCHEMA)
    AND COALESCE(LANGUAGE_NAME, '') = COALESCE(@language, LANGUAGE_NAME, '')
    AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('xls'))

UNION ALL SELECT 4 AS part, -1 AS ID, @GoLine AS command
    WHERE EXISTS(
        SELECT
            ID
        FROM
            xls.translations
        WHERE
            TABLE_SCHEMA LIKE COALESCE(@schema, TABLE_SCHEMA)
            AND COALESCE(LANGUAGE_NAME, '') = COALESCE(@language, LANGUAGE_NAME, '')
            AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('xls'))
    )

UNION ALL
SELECT
    4 AS part
    , CASE WHEN @sort_by_names = 1 THEN ROW_NUMBER() OVER(ORDER BY TABLE_SCHEMA, TABLE_NAME) ELSE ID END AS ID
    , @cmd4
           + CASE WHEN TABLE_SCHEMA              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_SCHEMA, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_NAME                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_EXCEL_FORMAT_XML    IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(CAST(TABLE_EXCEL_FORMAT_XML AS nvarchar(max)), '''', '''''') + '''' END
    + @cmd0 AS command
FROM
    xls.formats
WHERE
    TABLE_SCHEMA LIKE COALESCE(@schema, TABLE_SCHEMA)
    AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('xls'))

UNION ALL SELECT 5 AS part, -1 AS ID, @GoLine AS command
    WHERE EXISTS(
        SELECT
            ID
        FROM
            xls.formats
        WHERE
            TABLE_SCHEMA LIKE COALESCE(@schema, TABLE_SCHEMA)
            AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('xls'))
    )

UNION ALL
SELECT
    5 AS part
    , CASE WHEN @sort_by_names = 1 THEN ROW_NUMBER() OVER(ORDER BY NAME) ELSE ID END AS ID
    , @cmd5
           + CASE WHEN NAME                      IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(NAME, '''', '''''') + '''' END
    + ', ' + CASE WHEN TEMPLATE                  IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TEMPLATE, '''', '''''') + '''' END
    + ', ' + CASE WHEN DEFINITION                IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(DEFINITION, '''', '''''') + '''' END
    + ', ' + CASE WHEN TABLE_SCHEMA              IS NULL THEN 'NULL' ELSE 'N''' + REPLACE(TABLE_SCHEMA, '''', '''''') + '''' END
    + @cmd0 AS command
FROM
    xls.workbooks
WHERE
    COALESCE(TABLE_SCHEMA, '') LIKE COALESCE(@schema, TABLE_SCHEMA, '')
    AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('xls'))

UNION ALL SELECT 6 AS part, -1 AS ID, @GoLine AS command
    WHERE EXISTS(
        SELECT
            ID
        FROM
            xls.workbooks
        WHERE
            TABLE_SCHEMA LIKE COALESCE(@schema, TABLE_SCHEMA)
            AND (@app_schema_only = 0 OR NOT TABLE_SCHEMA IN ('xls'))
    )
    ) t
WHERE
    t.part = COALESCE(@part, t.part)
    AND (@part IS NULL OR NOT t.command = '')
ORDER BY
    t.part
    , t.ID

END


GO

IF OBJECT_ID('[xls].[view_framework_objects]', 'V') IS NULL
    EXEC ('CREATE VIEW [xls].[view_framework_objects] AS SELECT 1 AS t')

GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: The view selects SaveToDB Framework objects
--
-- This code is used in the SaveToDB Add-In, DBEdit, ODataDB, and DBGate
-- to detect SaveToDB Framework objects
-- =============================================

ALTER VIEW [xls].[view_framework_objects]
AS

SELECT
    MAX(s.name) AS TABLE_SCHEMA
    , MAX(o.name) AS TABLE_NAME
    , 81 AS TABLE_VERSION
FROM
    sys.columns c
    INNER JOIN sys.objects o ON o.[object_id] = c.[object_id]
    INNER JOIN sys.schemas s ON s.[schema_id] = o.[schema_id]
WHERE
       c.name = 'TABLE_SCHEMA'      COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TABLE_NAME'        COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TABLE_TYPE'        COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TABLE_CODE'        COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'INSERT_OBJECT'     COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'UPDATE_OBJECT'     COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'DELETE_OBJECT'     COLLATE SQL_Latin1_General_CP1_CI_AI
GROUP BY
    c.[object_id]
HAVING
    COUNT(*) = 7
    AND MAX(c.column_id) <= 8
UNION ALL
SELECT
    MAX(s.name) AS TABLE_SCHEMA
    , MAX(o.name) AS TABLE_NAME
    , 51 AS TABLE_VERSION
FROM
    sys.columns c
    INNER JOIN sys.objects o ON o.[object_id] = c.[object_id]
    INNER JOIN sys.schemas s ON s.[schema_id] = o.[schema_id]
WHERE
       c.name = 'TABLE_SCHEMA'      COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TABLE_NAME'        COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TABLE_TYPE'        COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TABLE_CODE'        COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'INSERT_PROCEDURE'  COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'UPDATE_PROCEDURE'  COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'DELETE_PROCEDURE'  COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'PROCEDURE_TYPE'    COLLATE SQL_Latin1_General_CP1_CI_AI
GROUP BY
    c.[object_id]
HAVING
    COUNT(*) IN (8)
    AND MAX(c.column_id) <= 9
UNION ALL
SELECT
    MAX(s.name) AS TABLE_SCHEMA
    , MAX(o.name) AS TABLE_NAME
    , 52 AS TABLE_VERSION
FROM
    sys.columns c
    INNER JOIN sys.objects o ON o.[object_id] = c.[object_id]
    INNER JOIN sys.schemas s ON s.[schema_id] = o.[schema_id]
WHERE
       c.name = 'TABLE_SCHEMA'      COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TABLE_NAME'        COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'COLUMN_NAME'       COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'EVENT_NAME'        COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'HANDLER_SCHEMA'    COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'HANDLER_NAME'      COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'HANDLER_TYPE'      COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'HANDLER_CODE'      COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TARGET_WORKSHEET'  COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'MENU_ORDER'        COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'EDIT_PARAMETERS'   COLLATE SQL_Latin1_General_CP1_CI_AI
GROUP BY
    c.[object_id]
HAVING
    COUNT(*) IN (11)
    AND MAX(c.column_id) <= 12
UNION ALL
SELECT
    MAX(s.name) AS TABLE_SCHEMA
    , MAX(o.name) AS TABLE_NAME
    , 53 AS TABLE_VERSION
FROM
    sys.columns c
    INNER JOIN sys.objects o ON o.[object_id] = c.[object_id]
    INNER JOIN sys.schemas s ON s.[schema_id] = o.[schema_id]
WHERE
    c.name = 'TABLE_SCHEMA'         COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TABLE_NAME'        COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'PARAMETER_NAME'    COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'SELECT_SCHEMA'     COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'SELECT_NAME'       COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'SELECT_TYPE'       COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'SELECT_CODE'       COLLATE SQL_Latin1_General_CP1_CI_AI
GROUP BY
    c.[object_id]
HAVING
    COUNT(*) IN (7)
    AND MAX(c.column_id) <= 8
UNION ALL
SELECT
    MAX(s.name) AS TABLE_SCHEMA
    , MAX(o.name) AS TABLE_NAME
    , CASE WHEN COUNT(*) = 7 THEN 85 WHEN MIN(c.name) = 'COLUMN_NAME' THEN 25 ELSE 24 END AS TABLE_VERSION
FROM
    sys.columns c
    INNER JOIN sys.objects o ON o.[object_id] = c.[object_id]
    INNER JOIN sys.schemas s ON s.[schema_id] = o.[schema_id]
WHERE
       c.name = 'TABLE_SCHEMA'          COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TABLE_NAME'            COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'COLUMN_NAME'           COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'LANGUAGE_NAME'         COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TRANSLATED_NAME'       COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TRANSLATED_DESC'       COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TRANSLATED_COMMENT'    COLLATE SQL_Latin1_General_CP1_CI_AI
GROUP BY
    c.[object_id]
HAVING
    COUNT(*) = 7 AND MAX(c.column_id) <= 8
    OR COUNT(*) = 6 AND MAX(c.column_id) <= 7
UNION ALL
SELECT
    MAX(s.name) AS TABLE_SCHEMA
    , MAX(o.name) AS TABLE_NAME
    , 26 AS TABLE_VERSION
FROM
    sys.columns c
    INNER JOIN sys.objects o ON o.[object_id] = c.[object_id]
    INNER JOIN sys.schemas s ON s.[schema_id] = o.[schema_id]
WHERE
       c.name = 'TABLE_SCHEMA'              COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TABLE_NAME'                COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TABLE_EXCEL_FORMAT_XML'    COLLATE SQL_Latin1_General_CP1_CI_AI
GROUP BY
    c.[object_id]
HAVING
    COUNT(*) = 3
    AND MAX(c.column_id) <= 4
UNION ALL
SELECT
    MAX(s.name) AS TABLE_SCHEMA
    , MAX(o.name) AS TABLE_NAME
    , 27 AS TABLE_VERSION
FROM
    sys.parameters c
    INNER JOIN sys.objects o ON o.[object_id] = c.[object_id]
    INNER JOIN sys.schemas s ON s.[schema_id] = o.[schema_id]
WHERE
       c.name = '@SCHEMA'           COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = '@NAME'             COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = '@EXCELFORMATXML'   COLLATE SQL_Latin1_General_CP1_CI_AI
GROUP BY
    c.[object_id]
HAVING
    COUNT(*) = 3
    AND MAX(c.parameter_id) <= 3
UNION ALL
SELECT
    MAX(s.name) AS TABLE_SCHEMA
    , MAX(o.name) AS TABLE_NAME
    , 88 AS TABLE_VERSION
FROM
    sys.columns c
    INNER JOIN sys.objects o ON o.[object_id] = c.[object_id]
    INNER JOIN sys.schemas s ON s.[schema_id] = o.[schema_id]
WHERE
       c.name = 'NAME'          COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'TEMPLATE'      COLLATE SQL_Latin1_General_CP1_CI_AI
    OR c.name = 'DEFINITION'    COLLATE SQL_Latin1_General_CP1_CI_AI
GROUP BY
    c.[object_id]
HAVING
    COUNT(*) = 3
    AND MAX(c.column_id) <= 4

GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: The view selects translations of all objects
-- =============================================

ALTER VIEW [xls].[view_all_translations]
AS

WITH cte AS (SELECT TABLE_SCHEMA, TABLE_NAME FROM xls.view_framework_objects)

SELECT
    1 AS SECTION
    , ROW_NUMBER() OVER (ORDER BY t.TABLE_TYPE, t.TABLE_SCHEMA, t.TABLE_NAME, g.LANGUAGE_NAME) AS SORT_ORDER
    , 'object' AS TRANSLATION_TYPE
    , CASE WHEN t.TABLE_TYPE = 'BASE TABLE' THEN 'table' ELSE LOWER(t.TABLE_TYPE) END AS TABLE_TYPE
    , t.TABLE_SCHEMA
    , t.TABLE_NAME
    , CAST(NULL AS nvarchar(128)) AS COLUMN_NAME
    , g.LANGUAGE_NAME
    , tr.TRANSLATED_NAME
    , tr.TRANSLATED_DESC
FROM
    INFORMATION_SCHEMA.TABLES t
    LEFT OUTER JOIN cte f ON f.TABLE_SCHEMA = t.TABLE_SCHEMA AND f.TABLE_NAME = t.TABLE_NAME

    CROSS JOIN (SELECT LANGUAGE_NAME FROM xls.translations UNION SELECT 'en' AS LANGUAGE_NAME) g

    LEFT OUTER JOIN xls.translations tr
        ON tr.TABLE_SCHEMA = t.TABLE_SCHEMA AND tr.TABLE_NAME = t.TABLE_NAME
            AND tr.COLUMN_NAME IS NULL AND tr.LANGUAGE_NAME = g.LANGUAGE_NAME
WHERE
    NOT t.TABLE_NAME IN ('sysdiagrams')
    AND NOT t.TABLE_SCHEMA IN ('sys', 'dbo01', 'etl01', 'xls01', 'savetodb_dev', 'savetodb_xls', 'savetodb_etl', 'SAVETODB_DEV', 'SAVETODB_XLS', 'SAVETODB_ETL', 'xls', 'doc', 'logs')
    AND f.TABLE_NAME IS NULL

UNION ALL
SELECT
    2 AS SECTION
    , ROW_NUMBER() OVER (ORDER BY p.ROUTINE_TYPE, p.SPECIFIC_SCHEMA, p.SPECIFIC_NAME, g.LANGUAGE_NAME) AS SORT_ORDER
    , 'object' AS TRANSLATION_TYPE
    , LOWER(p.ROUTINE_TYPE) AS TABLE_TYPE
    , p.SPECIFIC_SCHEMA AS TABLE_SCHEMA
    , p.SPECIFIC_NAME AS TABLE_NAME
    , CAST(NULL AS nvarchar(128)) AS COLUMN_NAME
    , g.LANGUAGE_NAME
    , tr.TRANSLATED_NAME
    , tr.TRANSLATED_DESC
FROM
    INFORMATION_SCHEMA.ROUTINES p
    LEFT OUTER JOIN cte f ON f.TABLE_SCHEMA = p.SPECIFIC_SCHEMA AND f.TABLE_NAME = p.SPECIFIC_NAME

    CROSS JOIN (SELECT LANGUAGE_NAME FROM xls.translations UNION SELECT 'en' AS LANGUAGE_NAME) g

    LEFT OUTER JOIN xls.translations tr
        ON tr.TABLE_SCHEMA = p.SPECIFIC_SCHEMA AND tr.TABLE_NAME = p.SPECIFIC_NAME
            AND tr.COLUMN_NAME IS NULL AND tr.LANGUAGE_NAME = g.LANGUAGE_NAME
WHERE
    NOT p.SPECIFIC_NAME LIKE 'sp_%'
    AND NOT p.SPECIFIC_NAME LIKE 'fn_%'
    AND p.ROUTINE_TYPE = 'PROCEDURE'
    AND NOT p.SPECIFIC_SCHEMA IN ('sys', 'dbo01', 'etl01', 'xls01', 'savetodb_dev', 'savetodb_xls', 'savetodb_etl', 'SAVETODB_DEV', 'SAVETODB_XLS', 'SAVETODB_ETL', 'xls', 'doc', 'logs')
    AND NOT p.SPECIFIC_NAME LIKE '%_insert'
    AND NOT p.SPECIFIC_NAME LIKE '%_update'
    AND NOT p.SPECIFIC_NAME LIKE '%_delete'
    AND NOT p.SPECIFIC_NAME LIKE '%_change'
    AND NOT p.SPECIFIC_NAME LIKE '%_merge'
    AND NOT p.SPECIFIC_NAME LIKE 'usp_import_%'
    AND NOT p.SPECIFIC_NAME LIKE 'usp_export_%'
    AND NOT p.SPECIFIC_NAME LIKE 'xl_parameter_values_%'
    AND NOT p.SPECIFIC_NAME LIKE 'xl_validation_list_%'
    AND f.TABLE_NAME IS NULL

UNION ALL
SELECT
    3 AS SECTION
    , ROW_NUMBER() OVER (ORDER BY c.TABLE_SCHEMA, g.LANGUAGE_NAME) AS SORT_ORDER
    , 'schema' AS TRANSLATION_TYPE
    , 'column' AS TABLE_TYPE
    , c.TABLE_SCHEMA
    , CAST(NULL AS nvarchar(128)) AS TABLE_NAME
    , c.COLUMN_NAME
    , g.LANGUAGE_NAME
    , tr.TRANSLATED_NAME
    , tr.TRANSLATED_DESC
FROM
    (
        SELECT
            c.TABLE_SCHEMA
            , c.COLUMN_NAME
        FROM
            INFORMATION_SCHEMA.COLUMNS c
            LEFT OUTER JOIN cte f ON f.TABLE_SCHEMA = c.TABLE_SCHEMA AND f.TABLE_NAME = c.TABLE_NAME
        WHERE
            NOT c.TABLE_NAME IN ('sysdiagrams')
            AND NOT c.TABLE_SCHEMA IN ('sys', 'dbo01', 'etl01', 'xls01', 'savetodb_dev', 'savetodb_xls', 'savetodb_etl', 'SAVETODB_DEV', 'SAVETODB_XLS', 'SAVETODB_ETL', 'xls', 'doc', 'logs')
            AND f.TABLE_NAME IS NULL
        UNION
        SELECT
            p.TABLE_SCHEMA
            , p.COLUMN_NAME
        FROM
            INFORMATION_SCHEMA.ROUTINE_COLUMNS p
            LEFT OUTER JOIN cte f ON f.TABLE_SCHEMA = p.TABLE_SCHEMA AND f.TABLE_NAME = p.TABLE_NAME
        WHERE
            NOT p.TABLE_NAME LIKE 'fn_%'
            AND NOT p.TABLE_SCHEMA IN ('sys', 'dbo01', 'etl01', 'xls01', 'savetodb_dev', 'savetodb_xls', 'savetodb_etl', 'SAVETODB_DEV', 'SAVETODB_XLS', 'SAVETODB_ETL', 'xls', 'doc', 'logs')
            AND f.TABLE_NAME IS NULL
        UNION
        SELECT
            p.SPECIFIC_SCHEMA AS TABLE_SCHEMA
            , SUBSTRING(p.PARAMETER_NAME, 2, 127) AS COLUMN_NAME
        FROM
            INFORMATION_SCHEMA.PARAMETERS p
            INNER JOIN INFORMATION_SCHEMA.ROUTINES r ON r.SPECIFIC_SCHEMA = p.SPECIFIC_SCHEMA AND r.SPECIFIC_NAME = p.SPECIFIC_NAME
            LEFT OUTER JOIN cte f ON f.TABLE_SCHEMA = p.SPECIFIC_SCHEMA AND f.TABLE_NAME = p.SPECIFIC_NAME
        WHERE
            NOT p.SPECIFIC_NAME LIKE 'sp_%'
            AND NOT p.SPECIFIC_NAME LIKE 'fn_%'
            AND r.ROUTINE_TYPE = 'PROCEDURE'
            AND p.ORDINAL_POSITION > 0
            AND NOT p.SPECIFIC_SCHEMA IN ('sys', 'dbo01', 'etl01', 'xls01', 'savetodb_dev', 'savetodb_xls', 'savetodb_etl', 'SAVETODB_DEV', 'SAVETODB_XLS', 'SAVETODB_ETL', 'xls', 'doc', 'logs')
            AND NOT p.SPECIFIC_NAME LIKE '%_insert'
            AND NOT p.SPECIFIC_NAME LIKE '%_update'
            AND NOT p.SPECIFIC_NAME LIKE '%_delete'
            AND NOT p.SPECIFIC_NAME LIKE '%_change'
            AND NOT p.SPECIFIC_NAME LIKE '%_merge'
            AND NOT p.SPECIFIC_NAME LIKE 'usp_import_%'
            AND NOT p.SPECIFIC_NAME LIKE 'usp_export_%'
            AND NOT p.SPECIFIC_NAME LIKE 'xl_parameter_values_%'
            AND NOT p.SPECIFIC_NAME LIKE 'xl_validation_list_%'
            AND NOT p.PARAMETER_NAME IN ('@data_language')
            AND f.TABLE_NAME IS NULL
    ) c
    CROSS JOIN (SELECT LANGUAGE_NAME FROM xls.translations UNION SELECT 'en' AS LANGUAGE_NAME) g

    LEFT OUTER JOIN xls.translations tr
        ON tr.TABLE_SCHEMA = c.TABLE_SCHEMA AND tr.TABLE_NAME IS NULL
            AND tr.COLUMN_NAME = c.COLUMN_NAME AND tr.LANGUAGE_NAME = g.LANGUAGE_NAME

UNION ALL
SELECT
    4 AS SECTION
    , ROW_NUMBER() OVER (ORDER BY c.TABLE_SCHEMA, c.TABLE_NAME, c.ORDINAL_POSITION, g.LANGUAGE_NAME) AS SORT_ORDER
    , 'column' AS TRANSLATION_TYPE
    , CASE WHEN t.TABLE_TYPE = 'BASE TABLE' THEN 'table' ELSE LOWER(t.TABLE_TYPE) END AS TABLE_TYPE
    , c.TABLE_SCHEMA
    , c.TABLE_NAME
    , c.COLUMN_NAME
    , g.LANGUAGE_NAME
    , tr.TRANSLATED_NAME
    , tr.TRANSLATED_DESC
FROM
    INFORMATION_SCHEMA.COLUMNS c
    INNER JOIN INFORMATION_SCHEMA.TABLES t ON t.TABLE_SCHEMA = c.TABLE_SCHEMA AND t.TABLE_NAME = c.TABLE_NAME
    LEFT OUTER JOIN cte f ON f.TABLE_SCHEMA = t.TABLE_SCHEMA AND f.TABLE_NAME = t.TABLE_NAME

    CROSS JOIN (SELECT LANGUAGE_NAME FROM xls.translations UNION SELECT 'en' AS LANGUAGE_NAME) g

    LEFT OUTER JOIN xls.translations tr
        ON tr.TABLE_SCHEMA = c.TABLE_SCHEMA AND tr.TABLE_NAME = c.TABLE_NAME
            AND tr.COLUMN_NAME = c.COLUMN_NAME AND tr.LANGUAGE_NAME = g.LANGUAGE_NAME

WHERE
    NOT c.TABLE_NAME IN ('sysdiagrams')
    AND NOT c.TABLE_SCHEMA IN ('sys', 'dbo01', 'etl01', 'xls01', 'savetodb_dev', 'savetodb_xls', 'savetodb_etl', 'SAVETODB_DEV', 'SAVETODB_XLS', 'SAVETODB_ETL', 'xls', 'doc', 'logs')
    AND f.TABLE_NAME IS NULL

UNION ALL
SELECT
    5 AS SECTION
    , ROW_NUMBER() OVER (ORDER BY p.TABLE_SCHEMA, p.TABLE_NAME, p.ORDINAL_POSITION, g.LANGUAGE_NAME) AS SORT_ORDER
    , 'column' AS TRANSLATION_TYPE
    , 'function' AS TABLE_TYPE
    , p.TABLE_SCHEMA
    , p.TABLE_NAME
    , p.COLUMN_NAME
    , g.LANGUAGE_NAME
    , tr.TRANSLATED_NAME
    , tr.TRANSLATED_DESC
FROM
    INFORMATION_SCHEMA.ROUTINE_COLUMNS p
    LEFT OUTER JOIN cte f ON f.TABLE_SCHEMA = p.TABLE_SCHEMA AND f.TABLE_NAME = p.TABLE_NAME

    CROSS JOIN (SELECT LANGUAGE_NAME FROM xls.translations UNION SELECT 'en' AS LANGUAGE_NAME) g

    LEFT OUTER JOIN xls.translations tr
        ON tr.TABLE_SCHEMA = p.TABLE_SCHEMA AND tr.TABLE_NAME = p.TABLE_NAME AND tr.COLUMN_NAME = p.COLUMN_NAME AND tr.LANGUAGE_NAME = g.LANGUAGE_NAME
WHERE
    NOT p.TABLE_NAME LIKE 'fn_%'
    AND NOT p.TABLE_SCHEMA IN ('sys', 'dbo01', 'etl01', 'xls01', 'savetodb_dev', 'savetodb_xls', 'savetodb_etl', 'SAVETODB_DEV', 'SAVETODB_XLS', 'SAVETODB_ETL', 'xls', 'doc', 'logs')
    AND f.TABLE_NAME IS NULL

UNION ALL
SELECT
    6 AS SECTION
    , ROW_NUMBER() OVER (ORDER BY p.SPECIFIC_SCHEMA, p.SPECIFIC_NAME, p.ORDINAL_POSITION, g.LANGUAGE_NAME) AS SORT_ORDER
    , 'parameter' AS TRANSLATION_TYPE
    , LOWER(r.ROUTINE_TYPE) AS TABLE_TYPE
    , p.SPECIFIC_SCHEMA AS TABLE_SCHEMA
    , p.SPECIFIC_NAME AS TABLE_NAME
    , SUBSTRING(p.PARAMETER_NAME, 2, 127) AS COLUMN_NAME
    , g.LANGUAGE_NAME
    , tr.TRANSLATED_NAME
    , tr.TRANSLATED_DESC
FROM
    INFORMATION_SCHEMA.PARAMETERS p
    INNER JOIN INFORMATION_SCHEMA.ROUTINES r ON r.SPECIFIC_SCHEMA = p.SPECIFIC_SCHEMA AND r.SPECIFIC_NAME = p.SPECIFIC_NAME
    LEFT OUTER JOIN cte f ON f.TABLE_SCHEMA = p.SPECIFIC_SCHEMA AND f.TABLE_NAME = p.SPECIFIC_NAME

    CROSS JOIN (SELECT LANGUAGE_NAME FROM xls.translations UNION SELECT 'en' AS LANGUAGE_NAME) g

    LEFT OUTER JOIN xls.translations tr
        ON tr.TABLE_SCHEMA = p.SPECIFIC_SCHEMA AND tr.TABLE_NAME = p.SPECIFIC_NAME
            AND tr.COLUMN_NAME = p.PARAMETER_NAME AND tr.LANGUAGE_NAME = g.LANGUAGE_NAME
WHERE
    NOT p.SPECIFIC_NAME LIKE 'sp_%'
    AND NOT p.SPECIFIC_NAME LIKE 'fn_%'
    AND r.ROUTINE_TYPE = 'PROCEDURE'
    AND p.ORDINAL_POSITION > 0
    AND NOT p.SPECIFIC_SCHEMA IN ('sys', 'dbo01', 'etl01', 'xls01', 'savetodb_dev', 'savetodb_xls', 'savetodb_etl', 'SAVETODB_DEV', 'SAVETODB_XLS', 'SAVETODB_ETL', 'xls', 'doc', 'logs')
    AND NOT p.SPECIFIC_NAME LIKE '%_insert'
    AND NOT p.SPECIFIC_NAME LIKE '%_update'
    AND NOT p.SPECIFIC_NAME LIKE '%_delete'
    AND NOT p.SPECIFIC_NAME LIKE '%_change'
    AND NOT p.SPECIFIC_NAME LIKE '%_merge'
    AND NOT p.SPECIFIC_NAME LIKE 'usp_import_%'
    AND NOT p.SPECIFIC_NAME LIKE 'usp_export_%'
    AND NOT p.SPECIFIC_NAME LIKE 'xl_parameter_values_%'
    AND NOT p.SPECIFIC_NAME LIKE 'xl_validation_list_%'
    AND NOT p.PARAMETER_NAME IN ('@data_language')
    AND f.TABLE_NAME IS NULL


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.0, 2022-07-05
-- Description: Selects translations
-- =============================================

ALTER PROCEDURE [xls].[usp_translations]
    @field nvarchar(128) = NULL
    , @schema nvarchar(128) = NULL
    , @is_complete bit = NULL
AS
BEGIN

SET NOCOUNT ON;

IF @field IS NULL
    SET @field = 'TRANSLATED_NAME'
ELSE IF @field NOT IN ('TRANSLATED_NAME', 'TRANSLATED_DESC', 'TRANSLATED_COMMENT')
    BEGIN
    DECLARE @message nvarchar(max) = N'Invalid column name: %s' + CHAR(13) + CHAR(10)
         + 'Use TRANSLATED_NAME, TRANSLATED_DESC, or TRANSLATED_COMMENT'
    RAISERROR(@message, 11, 0, @field);
    RETURN
    END

DECLARE @sql nvarchar(max)
DECLARE @languages nvarchar(max)
DECLARE @complete nvarchar(max)

SELECT @languages = STUFF((
    SELECT
        t.name
    FROM
        (
        SELECT
            DISTINCT
            ', [' + t.LANGUAGE_NAME + ']' AS name
            , CASE
                WHEN t.LANGUAGE_NAME = 'en' THEN '1'
                WHEN t.LANGUAGE_NAME = 'fr' THEN '2'
                WHEN t.LANGUAGE_NAME = 'it' THEN '3'
                WHEN t.LANGUAGE_NAME = 'es' THEN '4'
                WHEN t.LANGUAGE_NAME = 'pt' THEN '5'
                WHEN t.LANGUAGE_NAME = 'de' THEN '6'
                WHEN t.LANGUAGE_NAME = 'ru' THEN '7'
                ELSE t.LANGUAGE_NAME
                END AS sort_order
        FROM
            xls.translations t
        ) t
    ORDER BY
        t.sort_order
    FOR XML PATH(''), TYPE).value('.', 'nvarchar(MAX)'), 1, 2, '')

IF @languages IS NULL SET @languages = '[en]'

SELECT @complete = STUFF((
    SELECT
        t.name
    FROM
        (
        SELECT
            DISTINCT
            ' OR [' + t.LANGUAGE_NAME + '] IS NULL' AS name
            , CASE
                WHEN t.LANGUAGE_NAME = 'en' THEN '1'
                WHEN t.LANGUAGE_NAME = 'fr' THEN '2'
                WHEN t.LANGUAGE_NAME = 'it' THEN '3'
                WHEN t.LANGUAGE_NAME = 'es' THEN '4'
                WHEN t.LANGUAGE_NAME = 'pt' THEN '5'
                WHEN t.LANGUAGE_NAME = 'de' THEN '6'
                WHEN t.LANGUAGE_NAME = 'ru' THEN '7'
                ELSE t.LANGUAGE_NAME
                END AS sort_order
        FROM
            xls.translations t
        ) t
    ORDER BY
        t.sort_order
    FOR XML PATH(''), TYPE).value('.', 'nvarchar(MAX)'), 1, 4, '')

IF @complete IS NULL SET @complete = '[en] IS NULL'

SET @sql = 'SELECT
    t.TABLE_SCHEMA
    , t.TABLE_NAME
    , t.COLUMN_NAME AS [COLUMN]
    , ' + @languages + '
    , CASE WHEN ' + @complete + ' THEN 0 ELSE 1 END AS is_complete
FROM
    (
        SELECT
            t.TABLE_SCHEMA
            , t.TABLE_NAME
            , t.COLUMN_NAME
            , t.LANGUAGE_NAME
            , t.' + @field + ' AS name
        FROM
            xls.translations t'
        + CASE WHEN @schema IS NULL THEN '' ELSE '
        WHERE
            ' + CASE WHEN LEFT(@schema, 1) = '(' AND RIGHT(@schema, 1) = ')' THEN 't.TABLE_SCHEMA IN ' + @schema
                ELSE 'COALESCE(t.TABLE_SCHEMA, '''') = COALESCE(' + COALESCE('''' + @schema + '''', 'NULL') + ', t.TABLE_SCHEMA, '''')'
                END
            END + '
    ) s PIVOT (
        MAX(name) FOR LANGUAGE_NAME IN (' + @languages + ')
    ) t'
    + CASE WHEN @is_complete IS NULL THEN '' WHEN @is_complete = 1 THEN '
WHERE
    NOT (' + @complete + ')'
    ELSE '
WHERE
    ' + @complete END + '
ORDER BY
    t.TABLE_SCHEMA
    , t.TABLE_NAME
    , t.COLUMN_NAME'

-- PRINT @sql

EXEC (@sql)

END


GO

print 'SaveToDB Developer Framework updated';
