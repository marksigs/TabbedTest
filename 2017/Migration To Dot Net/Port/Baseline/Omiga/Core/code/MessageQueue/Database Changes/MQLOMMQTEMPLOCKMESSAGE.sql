----------------------------------------------------------------------------------------
-- DatabaseObject Name: MQLOMMQTEMPLOCKMESSAGE					
-- DatabaseObject Type: TEMPORARY TABLE
-- Description: 	
--
-- Copyright:   Copyright ©2002 Marlborough Stirling
--
-- Date created:	26/02/2001
-- Created by:		Lee Dawson
-- 
-- Used in:		Package SP_MQLOMMQ
--
----------------------------------------------------------------------------------------
-- DatabaseObject History
----------------------------------------------------------------------------------------
-- Developer    Date          Description
-- LD	       15/04/2002     SYS4414 First created
----------------------------------------------------------------------------------------

SET Serveroutput ON;
SET TrimSpool    ON;
SET Verify       OFF;
SET PAUSE        OFF;
SET LineSize     200;


Define Schema_Name      = &1;

CREATE GLOBAL TEMPORARY TABLE &&Schema_Name..MQLOMMQTEMPLOCKMESSAGE (
    QueryId RAW(16),
	MessageId RAW(16) NOT NULL,
	ProgId VARCHAR2(80) NOT NULL
) ON COMMIT PRESERVE ROWS;

