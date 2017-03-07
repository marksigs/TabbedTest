----------------------------------------------------------------------------------------
-- DatabaseObject Name: sp_MQLOMMQ					
-- DatabaseObject Type: PACKAGE
-- Description: 		Package for processing messages
--
-- Copyright:   	Copyright ©2001 Marlborough Stirling
--
-- Date created:	26/02/2001
-- Created by:	Tim Corcoran
-- 
-- Called by:	Message Queue VB code, eg. omToOMMQ/OmigaToMessageQueue.cls
-- Calls/uses:	Table MQLOMMQ
--
----------------------------------------------------------------------------------------
-- DatabaseObject History
----------------------------------------------------------------------------------------
-- Developer    	Date      	Description
-- TimC	     	26/02/2001 	First created
-- Andy Dowson 	21/03/2001  Extended for additional columns in MQLOMMQ (AGR SYS1998)
-- Mike Miller 	26/03/2001 	Tidied up (comment header); keyword FROM added in Get_Message
--                            Spec of LockMessage changed in header to match body changes
-- Lee Dawson  	11/04/2001  Corrections to SYS1998 (LockMessage)
-- TimC		17/04/2001  Changed code to make use of bind variables and single execution plans
-- Lee Dawson  	31/05/2001	Corrections to the above change
-- Lee Dawson   15/05/2002	SYS4414 Support Oracle's OLEDB provider (use REFCURSORs)/
--							Allow null to be passed into SendMessage for execute after date.
--							Added GetMessageOrCLOB, GetMessageOrLONG, SendMessageCLOB,
--							SendMessageLONG.  Removed GetMessage
----------------------------------------------------------------------------------------

SET Serveroutput ON;
SET TrimSpool    ON;
SET Verify       OFF;
SET PAUSE        OFF;
SET LineSize     200;


Define Schema_Name      = &1;

CREATE OR REPLACE PACKAGE &&Schema_Name..sp_MQLOMMQ AS

    TYPE IO_CUR IS REF CURSOR;

    -- place a message on the queue for processing
	PROCEDURE SendMessage(
		rawMessageId IN MQLOMMQ.MessageId%TYPE,
		szQueueName IN MQLOMMQ.QueueName%TYPE,
        szProgId IN MQLOMMQ.ProgId%TYPE,
		szXml IN MQLOMMQ.Xml%TYPE,
		nPriority IN MQLOMMQ.Priority%TYPE,
		dtExecuteAfterDate IN MQLOMMQ.ExecuteAfterDate%TYPE		
        );

    -- place a message on the queue for processing
	PROCEDURE SendMessageCLOB(
		rawMessageId IN MQLOMMQ.MessageId%TYPE,
		szQueueName IN MQLOMMQ.QueueName%TYPE,
        szProgId IN MQLOMMQ.ProgId%TYPE,
		szXml IN MQLOMMQCLOB.Xml%TYPE,
		nPriority IN MQLOMMQ.Priority%TYPE,
		dtExecuteAfterDate IN MQLOMMQ.ExecuteAfterDate%TYPE		
        );

    -- place a message on the queue for processing
	PROCEDURE SendMessageLONG(
		rawMessageId IN MQLOMMQ.MessageId%TYPE,
		szQueueName IN MQLOMMQ.QueueName%TYPE,
        szProgId IN MQLOMMQ.ProgId%TYPE,
		szXml IN MQLOMMQLONG.Xml%TYPE, -- LONG datatype (approx 32K)
		nPriority IN MQLOMMQ.Priority%TYPE,
		dtExecuteAfterDate IN MQLOMMQ.ExecuteAfterDate%TYPE		
        );

	-- allocates a number of messages, and returns their Ids
	-- -- called in preparation for processing and is in its own transaction
    PROCEDURE LockMessage(
		szQueueName IN MQLOMMQ.QueueName%TYPE,
		szLockedBy IN MQLOMMQ.LockedBy%TYPE,
		nMessagesToAllocate IN NUMBER,
		iocurLockMessage OUT NOCOPY IO_CUR
        );

	-- unlock a specific message 
	-- -- called in a transaction of its own, so that the processing of the message can be
	-- -- tried again at a later date
    PROCEDURE UnlockMessage(
		rawMessageId IN MQLOMMQ.MessageId%TYPE
        );

	-- unlock all messages for this user
	-- -- called when the listening server restarts (i.e. knowing it is processing no messages) in order
	-- -- to clean up after a power failure
    PROCEDURE UnlockMessages(
		szQueueName IN MQLOMMQ.QueueName%TYPE,
		szLockedBy IN MQLOMMQ.LockedBy%TYPE
        );

	-- get the details of the message
	-- -- called in the same transaction as the processed message
    PROCEDURE GetMessageOrCLOB(
		rawMessageId IN MQLOMMQ.MessageId%TYPE,
		iocurGetMessage OUT NOCOPY IO_CUR
        );

	-- get the details of the message
	-- -- called in the same transaction as the processed message
    PROCEDURE GetMessageOrLONG(
		rawMessageId IN MQLOMMQ.MessageId%TYPE,
		iocurGetMessage OUT NOCOPY IO_CUR
        );

    -- moves a message from QueueName to QueueName + 'X'
	PROCEDURE MoveMessage(
		rawMessageId IN MQLOMMQ.MessageId%TYPE
        );

    -- resets the queue (all messages on the queue can be processed)
	PROCEDURE ResetQueue(
		szQueueName IN MQLOMMQ.QueueName%TYPE
        );

END sp_MQLOMMQ;
/
SHOW ERRORS;


CREATE OR REPLACE PACKAGE BODY &&Schema_Name..sp_MQLOMMQ AS

	PROCEDURE SendMessage(
		rawMessageId IN MQLOMMQ.MessageId%TYPE,
		szQueueName IN MQLOMMQ.QueueName%TYPE,
        szProgId IN MQLOMMQ.ProgId%TYPE,
		szXml IN MQLOMMQ.Xml%TYPE,
		nPriority IN MQLOMMQ.Priority%TYPE,
		dtExecuteAfterDate IN MQLOMMQ.ExecuteAfterDate%TYPE		
        ) 
	AS
	BEGIN
		EXECUTE IMMEDIATE 'INSERT INTO MQLOMMQ (MessageId, QueueName, ProgId, Xml, DateTimePriority, ExecuteAfterDate, Priority) VALUES (:1, :2, :3, :4, SYSDATE, NVL(:5, SYSDATE), :6)'
			USING rawMessageId, szQueueName, szProgId, szXml, dtExecuteAfterDate, nPriority;
	END;


	PROCEDURE SendMessageCLOB(
		rawMessageId IN MQLOMMQ.MessageId%TYPE,
		szQueueName IN MQLOMMQ.QueueName%TYPE,
        szProgId IN MQLOMMQ.ProgId%TYPE,
		szXml IN MQLOMMQCLOB.Xml%TYPE,
		nPriority IN MQLOMMQ.Priority%TYPE,
		dtExecuteAfterDate IN MQLOMMQ.ExecuteAfterDate%TYPE		
        ) 
	AS
	BEGIN
		EXECUTE IMMEDIATE 'INSERT INTO MQLOMMQ (MessageId, QueueName, ProgId, DateTimePriority, ExecuteAfterDate, Priority) VALUES (:1, :2, :3, SYSDATE, NVL(:4, SYSDATE), :5)'
			USING rawMessageId, szQueueName, szProgId, dtExecuteAfterDate, nPriority;

		EXECUTE IMMEDIATE 'INSERT INTO MQLOMMQCLOB (MessageId, Xml) VALUES (:1, :2)'
			USING rawMessageId, szXml;

	END;


	PROCEDURE SendMessageLONG(
		rawMessageId IN MQLOMMQ.MessageId%TYPE,
		szQueueName IN MQLOMMQ.QueueName%TYPE,
        szProgId IN MQLOMMQ.ProgId%TYPE,
		szXml IN MQLOMMQLONG.Xml%TYPE, -- LONG datatype (approx 32K)
		nPriority IN MQLOMMQ.Priority%TYPE,
		dtExecuteAfterDate IN MQLOMMQ.ExecuteAfterDate%TYPE		
        ) 
	AS
	BEGIN
		EXECUTE IMMEDIATE 'INSERT INTO MQLOMMQ (MessageId, QueueName, ProgId, DateTimePriority, ExecuteAfterDate, Priority) VALUES (:1, :2, :3, SYSDATE, NVL(:4, SYSDATE), :5)'
			USING rawMessageId, szQueueName, szProgId, dtExecuteAfterDate, nPriority;

		EXECUTE IMMEDIATE 'INSERT INTO MQLOMMQLONG (MessageId, Xml) VALUES (:1, :2)'
			USING rawMessageId, szXml;

	END;


    PROCEDURE LockMessage(
		szQueueName IN MQLOMMQ.QueueName%TYPE,
		szLockedBy IN MQLOMMQ.LockedBy%TYPE,
		nMessagesToAllocate IN NUMBER,
		iocurLockMessage OUT NOCOPY IO_CUR
        ) AS

		nIndex INTEGER;
		rawMessageId MQLOMMQ.MessageId%TYPE;
		szProgId MQLOMMQ.ProgId%TYPE;
		rawQueryId MQLOMMQTEMPLOCKMESSAGE.QueryId%TYPE;

		CURSOR curMessageQueue(v_szQueueName IN MQLOMMQ.QueueName%TYPE) IS 
			SELECT MessageId, ProgId FROM MQLOMMQ
			WHERE QueueName = v_szQueueName AND LockedBy IS NULL AND DateTimePriority IS NOT NULL AND SYSDATE >= ExecuteAfterDate
			ORDER BY Priority ASC, DateTimePriority FOR UPDATE;

	BEGIN
		rawQueryId := SYS_GUID();

		nIndex := 1;
		OPEN curMessageQueue(szQueueName);
		
		FETCH curMessageQueue INTO rawMessageId, szProgId;
		WHILE curMessageQueue%FOUND AND nIndex <= nMessagesToAllocate LOOP
			
			UPDATE MQLOMMQ SET LockedBy = szLockedBy, DateTimePriority = SYSDATE WHERE CURRENT OF curMessageQueue;
			-- insert result into a temporary table
			EXECUTE IMMEDIATE 'INSERT INTO MQLOMMQTEMPLOCKMESSAGE (QueryId, MessageId, ProgId) VALUES (:1, :2, :3)'
				USING rawQueryId, rawMessageId, szProgId;

			nIndex := nIndex + 1;
			FETCH curMessageQueue INTO rawMessageId, szProgId;
		END LOOP;		
		CLOSE curMessageQueue;

		-- return results from temporary table
		OPEN iocurLockMessage FOR
			'SELECT RAWTOHEX(MessageId), ProgId FROM MQLOMMQTEMPLOCKMESSAGE WHERE QueryId = :1'
		USING rawQueryId;

		-- clean up temporary table
		EXECUTE IMMEDIATE 'DELETE MQLOMMQTEMPLOCKMESSAGE WHERE QueryId = :1' 
			USING rawQueryId;
	END;


    PROCEDURE UnlockMessages(
		szQueueName IN MQLOMMQ.QueueName%TYPE,
		szLockedBy IN MQLOMMQ.LockedBy%TYPE
        ) 
	AS
	BEGIN
		EXECUTE IMMEDIATE 'UPDATE MQLOMMQ SET LockedBy = NULL, DateTimePriority = SYSDATE WHERE QueueName = :1 AND LockedBy = :2'
			USING szQueueName, szLockedBy;
	END;


    PROCEDURE UnlockMessage(
		rawMessageId IN MQLOMMQ.MessageId%TYPE
        ) 
	AS
	BEGIN
		EXECUTE IMMEDIATE 'UPDATE MQLOMMQ SET LockedBy = NULL, DateTimePriority = NULL WHERE MessageId = :1'
			USING rawMessageId;
	END;


    PROCEDURE GetMessageOrCLOB(
		rawMessageId IN MQLOMMQ.MessageId%TYPE,
		iocurGetMessage OUT NOCOPY IO_CUR
		) 
	AS
		nLengthXml INTEGER;
		szProgId MQLOMMQ.ProgId%TYPE;
	BEGIN
		EXECUTE IMMEDIATE 'SELECT ProgId, NVL(LENGTH(Xml), 0) FROM MQLOMMQ WHERE MessageId = :1'
			INTO szProgId, nLengthXml USING rawMessageId;
		
		IF nLengthXml = 0 THEN
			OPEN iocurGetMessage FOR
				'SELECT :1, Xml FROM MQLOMMQCLOB WHERE MessageId = :2'
			USING szProgId, rawMessageId;

			EXECUTE IMMEDIATE 'DELETE FROM MQLOMMQCLOB WHERE MessageId = :1'
				USING rawMessageId;
		ELSE
			OPEN iocurGetMessage FOR
				'SELECT ProgId, Xml FROM MQLOMMQ WHERE MessageId = :1'
			USING rawMessageId;
		END IF;

		EXECUTE IMMEDIATE 'DELETE FROM MQLOMMQ WHERE MessageId = :1'
			USING rawMessageId;
	END;

    PROCEDURE GetMessageOrLONG(
		rawMessageId IN MQLOMMQ.MessageId%TYPE,
		iocurGetMessage OUT NOCOPY IO_CUR
		) 
	AS
		nLengthXml INTEGER;
		szProgId MQLOMMQ.ProgId%TYPE;
	BEGIN
		EXECUTE IMMEDIATE 'SELECT ProgId, NVL(LENGTH(Xml), 0) FROM MQLOMMQ WHERE MessageId = :1'
			INTO szProgId, nLengthXml USING rawMessageId;
		
		IF nLengthXml = 0 THEN
			OPEN iocurGetMessage FOR
				'SELECT :1, Xml FROM MQLOMMQLONG WHERE MessageId = :2'
			USING szProgId, rawMessageId;

			EXECUTE IMMEDIATE 'DELETE FROM MQLOMMQLONG WHERE MessageId = :1'
				USING rawMessageId;
		ELSE
			OPEN iocurGetMessage FOR
				'SELECT ProgId, Xml FROM MQLOMMQ WHERE MessageId = :1'
			USING rawMessageId;
		END IF;

		EXECUTE IMMEDIATE 'DELETE FROM MQLOMMQ WHERE MessageId = :1'
			USING rawMessageId;
	END;

	PROCEDURE MoveMessage(
		rawMessageId IN MQLOMMQ.MessageId%TYPE
        ) 
	AS
	BEGIN
		EXECUTE IMMEDIATE 'UPDATE MQLOMMQ SET QueueName = CONCAT(QueueName, ''X'') WHERE MessageId = :1'
			USING rawMessageId;
	END;

	PROCEDURE ResetQueue(
		szQueueName IN MQLOMMQ.QueueName%TYPE
        ) 
	AS
	BEGIN
		EXECUTE IMMEDIATE 'UPDATE MQLOMMQ SET DateTimePriority = SYSDATE WHERE QueueName = :1 AND DateTimePriority IS NULL'
			USING szQueueName;
	END;

END;
/
SHOW ERRORS;



