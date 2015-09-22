-------------------------------------------
-- NAME:   DatabaseMail.sql
-- AUTHOR: Niels van de Coevering
-- DATE:   17 August 2014
--
-- Enable and configure Database Mail for SQL Server instance
-------------------------------------------

EXECUTE sp_configure 'show advanced', 1;
RECONFIGURE;
EXECUTE sp_configure 'Database Mail XPs',1;
RECONFIGURE;
GO

-- Constants
DECLARE
    @mailServer AS varchar(50) = '',
    @operatorName AS varchar(50) = '',
    @operatorMail AS varchar(50) = '',
    @localDomain AS varchar(50) = '',
    @testRecipients AS varchar(50) = '';

-- Variables
DECLARE 
    @serverName AS varchar(50),
    @displayName AS varchar(50),
    @instantName AS varchar(50),
    @serverEmail AS varchar(50),
    @pver AS varchar(10),
    @version AS varchar(2),
    @regPath AS varchar(200);

SELECT @serverName = SUBSTRING(@@SERVERNAME, 1, CHARINDEX('\', @@SERVERNAME, 1)-1)
SELECT @displayName = REPLACE(@@SERVERNAME, '\', '-')
SELECT @instantName = SUBSTRING(@@SERVERNAME, CHARINDEX('\', @@SERVERNAME, 1)+1, 50);
SELECT @serverEmail = @displayName+@localDomain

-- Get SQL Server product version
SELECT @pver = CAST(SERVERPROPERTY('productversion') AS varchar(10));
SELECT @version = SUBSTRING(@pver, 1, CHARINDEX('.', @pver, 1)-1);

-- Delete existing profile and account
IF EXISTS (SELECT * FROM msdb.dbo.sysmail_profile WHERE name = @serverName)
BEGIN
    EXEC msdb.dbo.sysmail_delete_profile_sp @profile_name = @serverName
END
IF EXISTS (SELECT * FROM msdb.dbo.sysmail_account WHERE name = @serverName)
BEGIN
    EXEC msdb.dbo.sysmail_delete_account_sp @account_name = @serverName
END

-- Create a Database Mail account
EXECUTE msdb.dbo.sysmail_add_account_sp
    @account_name = @serverName,
    @description = '',
    @email_address = @serverEmail,
    @replyto_address = '',
    @display_name = @displayName,
    @mailserver_name = @mailServer;
 
-- Create a Database Mail profile
EXECUTE msdb.dbo.sysmail_add_profile_sp
    @profile_name = @serverName,
    @description = '';
 
-- Add the account to the profile
EXECUTE msdb.dbo.sysmail_add_profileaccount_sp
    @profile_name = @serverName,
    @account_name = @serverName,
    @sequence_number = 1;
 
-- Grant access to the profile to all msdb database users
EXECUTE msdb.dbo.sysmail_add_principalprofile_sp
    @profile_name = @serverName,
    @principal_name = 'public',
    @is_default = 1;

-- Delete and create VDR IT Error operator
IF EXISTS (SELECT * FROM MSDB.dbo.sysoperators WHERE name = @operatorName) 
BEGIN
    EXEC msdb.dbo.sp_delete_operator @name=@operatorName
END

EXEC msdb.dbo.sp_add_operator @name=@operatorName, 
		@enabled=1, 
		@weekday_pager_start_time=90000, 
		@weekday_pager_end_time=180000, 
		@saturday_pager_start_time=90000, 
		@saturday_pager_end_time=180000, 
		@sunday_pager_start_time=90000, 
		@sunday_pager_end_time=180000, 
		@pager_days=0, 
		@email_address=@operatorMail, 
		@category_name=N'[Uncategorized]'

EXECUTE sp_configure 'show advanced', 0;
RECONFIGURE;

SELECT 'Don''t forget to enable mail in SQL Server Agent > Alert System > Enable mail profile'

--send a test email
EXECUTE msdb.dbo.sp_send_dbmail
    @subject = 'Test Database Mail Message',
    @recipients = @testRecipients,
    @query = 'SELECT @@SERVERNAME';
GO
