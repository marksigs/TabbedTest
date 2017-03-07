----------------------------------------------------------------------------------------
-- DatabaseObject Name: MQLOMMQCLOB					
-- DatabaseObject Type: TABLE
-- Description: 	Table used for handling (larger) messages
--
-- Copyright:   Copyright ©2002 Marlborough Stirling
--
-- Date created:	03/05/2002
-- Created by:		Lee Dawson
-- 
-- Used in:		Package SP_MQLOMMQ
--
----------------------------------------------------------------------------------------
-- DatabaseObject History
----------------------------------------------------------------------------------------
-- Developer    	Date      	Description
-- LD	       	03/05/2002 	First created
----------------------------------------------------------------------------------------
SET Serveroutput ON;
SET TrimSpool    ON;
SET Verify       OFF;
SET PAUSE        OFF;
SET LineSize     200;

Define Schema_Name      = &1;
Define Table_Tablespace = &2;
Define Table_Initial    = &3;
Define Table_Next       = &4;
Define Index_Tablespace = &5;
Define Index_Initial    = &6;
Define Index_Next       = &7;
Define Pctincrease      = &8;

CREATE TABLE &&Schema_Name..MQLOMMQCLOB
(
	MessageId RAW(16) NOT NULL,
	Xml CLOB, -- a store for large messages (>2000 bytes)
	CONSTRAINT PK_MQLOMMQCLOB
      PRIMARY KEY (MessageId)
          USING INDEX TABLESPACE &&Index_Tablespace
               STORAGE (INITIAL &&Index_Initial
                         NEXT &&Index_Next
                         PCTINCREASE &&Pctincrease))
          TABLESPACE &&Table_Tablespace
               STORAGE (INITIAL &&Table_Initial
                         NEXT &&Table_Next
                         PCTINCREASE &&Pctincrease
);



