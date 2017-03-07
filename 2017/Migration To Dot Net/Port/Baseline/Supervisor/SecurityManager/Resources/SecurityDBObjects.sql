/*
ORACLE Script.

-- Create the User-to-Item mapping table.
CREATE TABLE "STEFB"."SUPERVISORUSERITEMS" (
	"USERID" VARCHAR2(10) NOT NULL,
	"ITEMID" NUMBER(4) NOT NULL,
	PRIMARY KEY("USERID", "ITEMID")
) TABLESPACE "STEFB"


-- Grant or deny permissions to an item for the specified user.
CREATE OR REPLACE PROCEDURE "STEFB"."USP_SETUSERITEM"    
(
	pUSERID IN VARCHAR2,
	pITEMID IN INTEGER,
	pALLOW IN INTEGER
)
IS
	iCOUNT INTEGER;
BEGIN
	IF pALLOW = 0 THEN
		DELETE
		FROM	SUPERVISORUSERITEMS 
		WHERE 	USERID = pUSERID AND 
			ITEMID = pITEMID;
	ELSE
		-- Count will let us ascertain whether the permission is already set.
		SELECT	COUNT(*) INTO iCOUNT 
		FROM 	SUPERVISORUSERITEMS 
		WHERE 	USERID = pUSERID AND 
			ITEMID = pITEMID;

		IF iCOUNT < 1 THEN
			INSERT INTO SUPERVISORUSERITEMS (USERID, ITEMID) 
			VALUES (pUSERID, pITEMID);
		END IF;
	END IF;
END;


-- Remove all granted access for the specified user.
CREATE OR REPLACE PROCEDURE "STEFB"."USP_DELETEUSERITEM"
(
	pUSERID IN VARCHAR2
)
IS
BEGIN
	DELETE
	FROM	SUPERVISORUSERITEMS
	WHERE 	USERID = pUSERID;
END;


-- Retrieve ALL assigned items, all those for a given user or simply
-- Set ALLOW to 1 or 0 to indicate if the specific item is allocated to
-- the specific user. 
CREATE OR REPLACE PROCEDURE "STEFB"."USP_GETUSERITEM"
(
	pUSERID IN VARCHAR2 DEFAULT NULL, 
	pITEMID IN INTEGER DEFAULT NULL,
	pALLOW IN OUT INTEGER,
	pREFCURSOR OUT NOCOPY sp_supervisor.Supervisor_IO_CUR)
IS
BEGIN
	SP_SUPERVISOR.GetUserItems(pUserID, pITEMID, pALLOW, pREFCURSOR);
END;


-- Package Header - GetUserItems
CREATE OR REPLACE PACKAGE "STEFB"."SP_SUPERVISOR" IS 
  TYPE Supervisor_IO_CUR IS REF CURSOR;
  PROCEDURE GetUserItems(pUserID VARCHAR2 DEFAULT NULL, pITEMID INTEGER DEFAULT NULL, pALLOW OUT INTEGER, pREFCURSOR OUT NOCOPY Supervisor_IO_CUR);
END;


-- Package Body - GetUserItems
CREATE OR REPLACE PACKAGE BODY "STEFB"."SP_SUPERVISOR"  
AS
	PROCEDURE GetUserItems
	(
		pUSERID VARCHAR2 DEFAULT NULL,
		pITEMID INTEGER DEFAULT NULL,
		pALLOW OUT INTEGER,
		pREFCURSOR OUT NOCOPY Supervisor_IO_CUR
	)
	IS
		iCOUNT INTEGER;
	BEGIN
		-- If USERID is null, get all records.
		IF pUSERID IS NULL THEN
			-- Execute the SQL and return any results.
			OPEN pREFCURSOR FOR SELECT USERID, ITEMID FROM SUPERVISORUSERITEMS;

		-- Otherwise, if OBJECTID is null get all records for the USERID.
		ELSIF pITEMID IS NULL THEN
			-- Execute the SQL and return any results.
			OPEN pREFCURSOR FOR SELECT USERID, ITEMID FROM SUPERVISORUSERITEMS WHERE USERID = pUSERID;

		-- Otherwise, if USERID and ITEMID are specified, use the output parameter.
		ELSE
			-- TODO: Dummy cursor, without it, we get an ORA-01023.
			OPEN pREFCURSOR FOR SELECT 1 FROM SUPERVISORUSERITEMS;
			
			-- Count will let us ascertain whether the permission is set.
			SELECT COUNT(*) INTO iCOUNT FROM SUPERVISORUSERITEMS WHERE USERID = pUSERID AND ITEMID = pITEMID;
	
			IF iCOUNT < 1 THEN
				pALLOW := 0;
			ELSE
				pALLOW := 1;
			END IF;			
		END IF;
	END;
END;
*/


-- Create the User-to-Item mapping table.
CREATE TABLE [SUPERVISORUSERITEMS] (
	[USERID] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[ITEMID] [int] NOT NULL,
	PRIMARY KEY 
	(
		[USERID],
		[ITEMID]
	) ON [PRIMARY]
) ON [PRIMARY]
GO


-- Retrieve ALL assigned items, all those for a given user or simply
-- Set ALLOW to 1 or 0 to indicate if the specific item is allocated to
-- the specific user. 
CREATE PROCEDURE USP_GETUSERITEM
	@USERID VARCHAR(10) 	= NULL,
	@ITEMID INT		= NULL,
	@ALLOW BIT 		= NULL OUTPUT
AS
BEGIN

	-- The SQL we'll execute is conditional.	
	DECLARE @sSQL VARCHAR(2000)

	SET NOCOUNT ON

	-- Build the first part of the SQL.
	SET @sSQL = 'SELECT USERID, ITEMID '
	SET @sSQL = @sSQL + 'FROM SUPERVISORUSERITEMS'
	
	-- If USERID is null, get all records.
	IF @USERID IS NULL BEGIN
		-- Execute the SQL and return any results.
		EXEC (@sSQL)

	-- Otherwise, if ITEMID is null get all records for the USERID.
	END ELSE IF @ITEMID IS NULL BEGIN
		-- Build a WHERE clause.
		SET @sSQL = @sSQL + ' WHERE USERID = ' + CHAR(39) + @USERID + CHAR(39)

		-- Execute the SQL and return any results.
		EXEC (@sSQL)

	-- Otherwise, if USERID and ITEMID are specified, use the output parameter.
	END ELSE BEGIN
		IF EXISTS (SELECT 1 FROM SUPERVISORUSERITEMS WHERE USERID = @USERID AND ITEMID = @ITEMID) BEGIN
			SET @ALLOW = 1
		END ELSE BEGIN
			SET @ALLOW = 0
		END
	END	
END


-- Grant or deny permissions to an item for the specified user.
CREATE PROCEDURE USP_SETUSERITEM
	@USERID VARCHAR(10),
	@ITEMID INT,
	@ALLOW BIT
AS
BEGIN

	SET NOCOUNT ON

	-- If ALLOW is false, delete any matching records.
	IF @ALLOW = 0 BEGIN
		DELETE
		FROM	SUPERVISORUSERITEMS
		WHERE	USERID = @USERID AND
			ITEMID = @ITEMID 

	-- Otherwise, we're adding the item.
	END ELSE BEGIN
		-- If there is no record, then add one (if the flag is set).
		IF NOT EXISTS (SELECT 1 FROM SUPERVISORUSERITEMS WHERE USERID = @USERID AND ITEMID = @ITEMID) BEGIN
			INSERT INTO SUPERVISORUSERITEMS (USERID, ITEMID)
			VALUES (@USERID, @ITEMID)
		END
	END
END


-- Remove all granted access for the specified user.
CREATE PROCEDURE USP_DELETEUSERITEM
	@USERID VARCHAR(10)
AS
BEGIN

	SET NOCOUNT ON

	DELETE
	FROM	SUPERVISORUSERITEMS
	WHERE 	USERID = @USERID

END

