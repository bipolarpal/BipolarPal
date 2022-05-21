CREATE TABLE [dbo].[Accounts]
(
	[Id] INT NOT NULL PRIMARY KEY, 
    [username] VARCHAR(50) NOT NULL, 
    [email] VARCHAR(255) NOT NULL UNIQUE, 
    [passwordhash] VARCHAR(255) NULL, 
    [resetcode] VARCHAR(8) NULL, 
    [regdate] DATETIME NULL, 
    [codeexpiray] DATETIME NULL, 
    [future] VARCHAR(255) NULL
)
