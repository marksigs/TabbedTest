----------------------------------------------------------------------------------------
-- DatabaseObject Name: USP_MQLOMMQSENDMESSAGE
-- DatabaseObject Type: Stored procedure
-- Description: 	SYS2705
--
-- Copyright:   Copyright ©2001 Marlborough Stirling
--
-- Date created:	13/09/2001
-- Created by:		Dale Le Maitre
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
-- DM       13/09/2001   	Created
-- LD	    29/04/2002		Allow null to be passed into SendMessage for execute after date.


----------------------------------------------------------------------------------------

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USP_MQLOMMQSENDMESSAGE]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[USP_MQLOMMQSENDMESSAGE]
GO
CREATE PROCEDURE USP_MQLOMMQSENDMESSAGE (
	@p_MessageId BINARY(16),  -- MQLOMMQ.MessageId%TYPE
	@p_QueueName NVARCHAR(39),   -- MQLOMMQ.QueueName%TYPE
	@p_ProgId    NVARCHAR(80),    -- MQLOMMQ.ProgId%TYPE
	@p_Xml       NVARCHAR(2000), -- MQLOMMQ.Xml%TYPE
	@p_Priority  NUMERIC(5),     -- MQLOMMQ.Priority%TYPE
	@p_dtExecuteAfterDate DATETIME  -- MQLOMMQ.ExecuteAfterDate%TYPE
	)
AS
BEGIN
	SET CONCAT_NULL_YIELDS_NULL OFF
	SET NOCOUNT ON

	INSERT INTO MQLOMMQ (MessageId, QueueName, ProgId, Xml, DateTimePriority, ExecuteAfterDate, Priority) 
	VALUES (@p_MessageId, @p_QueueName, @p_ProgId, @p_Xml, GETDATE(), ISNULL(@p_dtExecuteAfterDate, GETDATE()), @p_Priority)
END
