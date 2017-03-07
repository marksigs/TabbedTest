----------------------------------------------------------------------------------------
-- DatabaseObject Name: USP_MQLOMMQGETMESSAGEORNTEXT
-- DatabaseObject Type: Stored procedure
-- Description: 	SYS2705
--
-- Copyright:   Copyright ©2002 Marlborough Stirling
--
-- Date created:	07/05/2002
-- Created by:		Lee Dawson
-- 
-- Called by:		Message Queue
-- Input parameters:	@p_MessageId
-- Output parameters:	Record set containing ProgId and XML
-- Calls/uses:		
--
----------------------------------------------------------------------------------------
-- DatabaseObject History
----------------------------------------------------------------------------------------
-- Developer    Date          Description
-- LD       07/05/2002   	Created

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_MQLOMMQGETMESSAGEORNTEXT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_MQLOMMQGETMESSAGEORNTEXT]
GO
CREATE PROCEDURE USP_MQLOMMQGETMESSAGEORNTEXT(
		@p_MessageId BINARY(16) -- MQLOMMQ.MessageId%TYPE,
)
AS
	DECLARE @nLengthXml INTEGER;
	DECLARE @szProgId [nvarchar] (80);
BEGIN
	SET CONCAT_NULL_YIELDS_NULL OFF
	SET NOCOUNT ON

	SELECT @szProgId = ProgId, @nLengthXml = ISNULL(LEN(Xml), 0) FROM MQLOMMQ WHERE MessageId = @p_MessageId

	IF @nLengthXml = 0
	BEGIN 
	    SELECT @szProgId, Xml FROM MQLOMMQNTEXT WHERE MessageId = @p_MessageId
	    DELETE FROM MQLOMMQNTEXT WHERE MessageId = @p_MessageId
	END
	ELSE
	BEGIN
	    SELECT ProgId, Xml FROM MQLOMMQ WHERE MessageId = @p_MessageId
	END
	DELETE FROM MQLOMMQ WHERE MessageId = @p_MessageId

END
