----------------------------------------------------------------------------------------
-- DatabaseObject Name: USP_MQLOMMQSENDMESSAGENTEXT
-- DatabaseObject Type: Stored procedure
-- Description: 	SYS2705
--
-- Copyright:   Copyright ©2002 Marlborough Stirling
--
-- Date created:	07/05/2002
-- Created by:		Lee Dawson
-- 
-- Called by:		Message Queue
-- Input parameters:	@p_MessageId, @p_QueueName, @p_ProgId, @p_Xml, @p_Priority, @p_dtExecuteAfterDate
-- Output parameters:	
-- Calls/uses:		
--
----------------------------------------------------------------------------------------
-- DatabaseObject History
----------------------------------------------------------------------------------------
-- Developer    Date          Description
-- LD       07/05/2002   	Created


----------------------------------------------------------------------------------------

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_MQLOMMQSENDMESSAGENTEXT]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_MQLOMMQSENDMESSAGENTEXT]
GO
CREATE PROCEDURE USP_MQLOMMQSENDMESSAGENTEXT (
	@p_MessageId BINARY(16),  -- MQLOMMQ.MessageId%TYPE
	@p_QueueName NVARCHAR(39),   -- MQLOMMQ.QueueName%TYPE
	@p_ProgId    NVARCHAR(80),    -- MQLOMMQ.ProgId%TYPE
	@p_Xml       NTEXT, -- MQLOMMQNTEXT.Xml%TYPE
	@p_Priority  NUMERIC(5),     -- MQLOMMQ.Priority%TYPE
	@p_dtExecuteAfterDate DATETIME  -- MQLOMMQ.ExecuteAfterDate%TYPE
	)
AS
BEGIN
	SET CONCAT_NULL_YIELDS_NULL OFF
	SET NOCOUNT ON

	INSERT INTO MQLOMMQ (MessageId, QueueName, ProgId, DateTimePriority, ExecuteAfterDate, Priority) 
	VALUES (@p_MessageId, @p_QueueName, @p_ProgId, GETDATE(), ISNULL(@p_dtExecuteAfterDate, GETDATE()), @p_Priority)

	INSERT INTO MQLOMMQNTEXT (MessageId, Xml) 
	VALUES (@p_MessageId, @p_Xml)

END
