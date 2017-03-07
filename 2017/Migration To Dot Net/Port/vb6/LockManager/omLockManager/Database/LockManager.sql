SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

----------------------------------------------------------------------------------------  
-- DatabaseObject Name: usp_LockManagerLockApplication  
-- Description: Lock an application and all the customers on the application.
--
-- A record is inserted on the APPLICATIONLOCK table for @applicationNumber.
-- The CUSTOMERROLE table is used to determine which customers need to be locked for this 
-- application; the CUSTOMERROLE records must exist before calling this stored procedure.
-- For each customer, a record is inserted into the CUSTOMERLOCK and 
-- CUSTOMERLOCKAPPLICATIONLOCK tables, if they do not already exist for this application.
-- If the customer records already exist for this application, then no action is taken.
--
-- Where there are existing, conflicting locking records, then the conflicting records
-- are returned as XML to the caller.
--
-- If @createNew = 1, then there should not already be a lock on the
-- APPLICATIONLOCK table - we are locking the application.
--
-- If @createNew = 0, then there should either not already be a lock on the
-- APPLICATIONLOCK table, or if there is a lock, it should be for this user (@userID).
-- This is useful when this user adds new customers to an application that the user already
-- has locked. The CUSTOMERROLE table is used to ensure that all customers on the application
-- are locked.
--
-- Locking hints prevent deadlocking and primary key violations due to concurrency problems,
-- by using update locks that are held until the end of the transaction.
--
----------------------------------------------------------------------------------------  
-- DatabaseObject History  
----------------------------------------------------------------------------------------  
-- Author		Date		Description  
-- AS			13/04/2005	BBGR34
----------------------------------------------------------------------------------------  

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.usp_LockManagerLockApplication') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.usp_LockManagerLockApplication
GO

CREATE PROCEDURE dbo.usp_LockManagerLockApplication(
	@applicationNumber NVARCHAR(12),
	@userID NVARCHAR(10),  
	@unitID NVARCHAR(10),  
	@machineID NVARCHAR(30),
	@createNew BIT = 1)
AS  
BEGIN  
	SET NOCOUNT ON  
	SET LOCK_TIMEOUT 90000

	IF 
	(
		@createNew = 1 AND EXISTS 
		(
			-- If creating a new application lock, then the application should not already
			-- be locked to anyone, including this user.
			SELECT 1 FROM APPLICATIONLOCK WITH (UPDLOCK, HOLDLOCK)  
			WHERE APPLICATIONNUMBER = @applicationNumber
		)
	)
	OR 
	(
		@createNew = 0 AND EXISTS 
		(
			-- If not creating a new application lock, then the application should either not be 
			-- already locked, or it can be locked but only to this user; 
			-- in other words it should not be locked to anyone who is not this user.
			SELECT 1 FROM APPLICATIONLOCK WITH (UPDLOCK, HOLDLOCK)  
			WHERE APPLICATIONNUMBER = @applicationNumber
			AND USERID <> @userID
		)	
	)
	BEGIN
		-- The application is already locked; return existing record to caller.
		SELECT * FROM APPLICATIONLOCK 
		WHERE APPLICATIONNUMBER = @applicationNumber FOR XML AUTO
	END
	ELSE
	BEGIN
		-- The application is not already locked to another user.

		IF EXISTS
		(
			SELECT 1 FROM CUSTOMERLOCK WITH (UPDLOCK, HOLDLOCK)
			WHERE CUSTOMERNUMBER IN 
			(	
				SELECT CUSTOMERNUMBER FROM CUSTOMERROLE WITH (NOLOCK) 
				WHERE APPLICATIONNUMBER = @applicationNumber
			)
			AND	USERID <> @userID
		)
		BEGIN
			-- One or more of the customers on this application are already locked,
			-- and they are not locked to this user.
			-- Return existing records to caller.
			SELECT * FROM CUSTOMERLOCK
			WHERE CUSTOMERNUMBER IN 
			(	
				SELECT CUSTOMERNUMBER FROM CUSTOMERROLE WITH (NOLOCK) 
				WHERE APPLICATIONNUMBER = @applicationNumber
			)
			AND	USERID <> @userID
			FOR XML AUTO
		END
		ELSE
		BEGIN
			-- None of the customers on the application are already locked.

			IF EXISTS
			(
				SELECT 1 FROM CUSTOMERLOCKAPPLICATIONLOCK WITH (UPDLOCK, HOLDLOCK)  
				WHERE APPLICATIONNUMBER <> @applicationNumber 
				AND CUSTOMERNUMBER IN 
				(	
					SELECT CUSTOMERNUMBER FROM CUSTOMERROLE WITH (NOLOCK) 
					WHERE APPLICATIONNUMBER = @applicationNumber
				)
			)			
			BEGIN 
				-- One or more of the customers on the application are locked to another application; return records to caller.
				SELECT * FROM CUSTOMERLOCKAPPLICATIONLOCK 
				WHERE APPLICATIONNUMBER <> @applicationNumber 
				AND CUSTOMERNUMBER IN 
				(	
					SELECT CUSTOMERNUMBER FROM CUSTOMERROLE WITH (NOLOCK) 
					WHERE APPLICATIONNUMBER = @applicationNumber
				)
				FOR XML AUTO
			END
			ELSE
			BEGIN
				-- None of the customers on the application are locked to other applications.

				DECLARE @lockDate DATETIME
				SET @lockDate = GETDATE()
			
				-- Only add an application lock if it does not already exist.
				IF NOT EXISTS 
				(
					SELECT 1 FROM APPLICATIONLOCK WITH (UPDLOCK, HOLDLOCK)  
					WHERE APPLICATIONNUMBER = @applicationNumber
				)
				BEGIN
					INSERT INTO APPLICATIONLOCK WITH (ROWLOCK) (APPLICATIONNUMBER, UNITID, USERID, LOCKDATE, MACHINEID, TYPEOFLOCK)
					VALUES (@applicationNumber, @unitID, @userID, @lockDate, @machineID, 'On')  
				END
				
				-- Add a customer lock for each customer on the application where they don't already have a lock.  
				INSERT INTO CUSTOMERLOCK WITH (ROWLOCK) (CUSTOMERNUMBER, UNITID, USERID, LOCKDATE, MACHINEID, TYPEOFLOCK)  
				SELECT CUSTOMERNUMBER, @unitID, @userID, @lockDate, @machineID, 'On'  
				FROM CUSTOMERROLE WITH (NOLOCK)  
				WHERE APPLICATIONNUMBER = @applicationNumber 
				AND NOT CUSTOMERNUMBER IN 
				(  
					SELECT CL.CUSTOMERNUMBER  
					FROM CUSTOMERLOCK CL WITH (NOLOCK)  
					WHERE CL.CUSTOMERNUMBER = CUSTOMERNUMBER
				)
				
				-- Add necessary customerlock-applicationlock mapping records, if they do not already exist.
				INSERT INTO CUSTOMERLOCKAPPLICATIONLOCK WITH (ROWLOCK) (APPLICATIONNUMBER, CUSTOMERNUMBER)  
				SELECT @applicationNumber, CUSTOMERNUMBER  
				FROM CUSTOMERROLE WITH (NOLOCK)  
				WHERE APPLICATIONNUMBER = @applicationNumber 
				AND NOT CUSTOMERNUMBER IN
				(
					SELECT CUSTOMERNUMBER 
					FROM CUSTOMERLOCKAPPLICATIONLOCK CLAL WITH (NOLOCK)
					WHERE CLAL.APPLICATIONNUMBER = @applicationNumber
					AND CLAL.CUSTOMERNUMBER = CUSTOMERNUMBER
				)
			END
		END
	END
END  
GO

----------------------------------------------------------------------------------------  
-- DatabaseObject Name: usp_LockManagerUnlockApplication  
-- Description: Unlock an application and all the customers on the application.
----------------------------------------------------------------------------------------  
-- DatabaseObject History  
----------------------------------------------------------------------------------------  
-- Author		Date		Description  
-- AS			13/04/2005	BBGR34
----------------------------------------------------------------------------------------  

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.usp_LockManagerUnlockApplication') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.usp_LockManagerUnlockApplication
GO

CREATE PROCEDURE dbo.usp_LockManagerUnlockApplication(@applicationNumber NVARCHAR(12))
AS  
BEGIN  
  
	SET NOCOUNT ON  
	
	
	-- Remove the application lock for this application.
	DELETE FROM APPLICATIONLOCK WITH (ROWLOCK) WHERE APPLICATIONNUMBER = @applicationNumber  
	
	-- Remove all the customer locks for this application.  
	DELETE FROM CUSTOMERLOCK WITH (ROWLOCK) WHERE CUSTOMERNUMBER IN 
	(  
		SELECT CUSTOMERNUMBER FROM CUSTOMERROLE WITH (NOLOCK) 
		WHERE APPLICATIONNUMBER = @applicationNumber
	)  

	-- Remove the customer-application lock mapping records.  
	DELETE FROM CUSTOMERLOCKAPPLICATIONLOCK WITH (ROWLOCK) 
	WHERE APPLICATIONNUMBER = @applicationNumber  
	
END  

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

----------------------------------------------------------------------------------------  
-- DatabaseObject Name: usp_LockManagerLockCustomer  
-- Description: Locks a customer.
----------------------------------------------------------------------------------------  
-- DatabaseObject History  
----------------------------------------------------------------------------------------  
-- Author		Date		Description  
-- AS			13/04/2005	BBGR34
----------------------------------------------------------------------------------------  

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.usp_LockManagerLockCustomer') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.usp_LockManagerLockCustomer
GO

CREATE PROCEDURE dbo.usp_LockManagerLockCustomer(
	@customerNumber NVARCHAR(12),
	@userID NVARCHAR(10),  
	@unitID NVARCHAR(10),  
	@machineID NVARCHAR(30))
AS  
BEGIN  
	SET NOCOUNT ON  
	SET LOCK_TIMEOUT 90000

	DECLARE @lockDate DATETIME
	SET @lockDate = GETDATE()

	-- Locking hints prevent deadlocking and primary key violations due to concurrency problems.
	-- Creates a update lock that is held until the end of the transaction.
	IF NOT EXISTS (SELECT 1 FROM CUSTOMERLOCK WITH (UPDLOCK, HOLDLOCK) WHERE CUSTOMERNUMBER = @customerNumber)
		-- Customer is not already locked.
		INSERT INTO CUSTOMERLOCK (CUSTOMERNUMBER, UNITID, USERID, LOCKDATE, MACHINEID, TYPEOFLOCK)  
		VALUES (@customerNumber, @unitID, @userID, @lockDate, @machineID, 'On')  			
	ELSE
		-- Customer is already locked; return existing lock record to caller.
		SELECT * FROM CUSTOMERLOCK WITH (UPDLOCK, HOLDLOCK) 
		WHERE CUSTOMERNUMBER = @customerNumber FOR XML AUTO
END  
GO

----------------------------------------------------------------------------------------  
-- DatabaseObject Name: usp_LockManagerUnlockCustomer  
-- Description: Unlock a customer.
----------------------------------------------------------------------------------------  
-- DatabaseObject History  
----------------------------------------------------------------------------------------  
-- Author		Date		Description  
-- AS			13/04/2005	BBGR34
----------------------------------------------------------------------------------------  

if exists (select * from dbo.sysobjects where id = object_id(N'dbo.usp_LockManagerUnlockCustomer') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure dbo.usp_LockManagerUnlockCustomer
GO

CREATE PROCEDURE dbo.usp_LockManagerUnlockCustomer(@customerNumber NVARCHAR(12))
AS  
BEGIN  
	SET NOCOUNT ON  
	DELETE FROM CUSTOMERLOCK WHERE CUSTOMERNUMBER = @customerNumber
END  
GO

----------------------------------------------------------------------------------------  
-- UNIT TESTS
----------------------------------------------------------------------------------------  
/*
delete from customerlock
delete from applicationlock
delete from customerlockapplicationlock

select * from customerrole
select * from applicationlock
select * from customerlock
select * from customerlockapplicationlock

select * from customerrole
update applicationlock set userid = 'msguser1' where userid = 'msguser'
update customerlock set userid = 'msguser1' where userid = 'msguser'
update customerlockapplicationlock set applicationnumber = '100000777'

exec usp_LockManagerLockCustomer '100000533', 'msguser1', '10', 'CH007770'
exec usp_LockManagerLockCustomer '100000541', 'msguser1', '10', 'CH007770'
exec usp_LockManagerLockApplication '100000703', 'msguser', '10', 'CH007770', 1
exec usp_LockManagerUnlockApplication '100000703

exec usp_LockManagerLockApplication '100000002', 'msguser', '10', 'CH007770', 0

declare @count int
set @count = 0
while @count >= 0
begin
	begin tran
	exec usp_LockManagerLockCustomer '1234', 'msguser1', '10', 'CH007770'
	commit tran
	begin tran
	exec usp_LockManagerUnlockCustomer '1234'	
	commit tran
	set @count = @count + 1
	print @count
end

declare @count int
set @count = 0
while @count >= 0
begin
	begin tran
	exec usp_LockManagerLockApplication '100000703', 'msguser1', '10', 'CH007770', 0
	commit tran
	begin tran
	exec usp_LockManagerUnlockApplication '100000703'	
	commit tran
	set @count = @count + 1
	print @count
end

*/
