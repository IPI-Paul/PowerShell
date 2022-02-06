CREATE PROCEDURE [dbo].[spEntityCache_Fix]	
AS
begin
	declare @Sql as nvarchar(100)
	set @Sql = 'alter database scoped configuration set IDENTITY_CACHE = off'
	exec (@Sql)
end