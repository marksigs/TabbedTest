if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MQLOMMQNTEXT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MQLOMMQNTEXT]
GO

CREATE TABLE [dbo].[MQLOMMQNTEXT] (
	[MESSAGEID] [BINARY] (16) NOT NULL ,
	[XML] [NTEXT] NULL 
)
GO

