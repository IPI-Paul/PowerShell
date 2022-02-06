declare @dbName as nvarchar(100), @tblContact as nvarchar(100)
set @dbName = 'PowerShellModulesDb'
set @tblContact = 'psContacts'
select 
	cast(case when DB_ID(@dbName ) is not null then 1 else 0 end as bit) as [DatabaseExists]
	,cast(case when OBJECT_ID(@dbName + '.dbo.' + @tblContact) is not null then 1 else 0 end as bit) as [ContactsExist]