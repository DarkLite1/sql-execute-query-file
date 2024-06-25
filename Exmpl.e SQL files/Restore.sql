-- enable advanced options
EXEC sp_configure 'show advanced options', 1
GO
RECONFIGURE
GO

-- enable xp_cmdshell
EXEC sp_configure 'xp_cmdshell', 1
GO
RECONFIGURE
GO

-- map network drive
EXEC xp_cmdshell 'NET USE Z: \\RemoteSrv\Path'
-- EXEC xp_cmdshell 'NET USE Z: \\Srv\Path password /USER:Domain\UserName'

-- restore database from remote UNC path
RESTORE DATABASE DataBaseName FROM DISK = 'Z:\Backup.bak' WITH REPLACE
GO

-- remove mapped drive
EXEC xp_cmdshell 'net use H: /delete'