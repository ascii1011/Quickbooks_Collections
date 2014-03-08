CREATE PROCEDURE grab_importance_sp
@listid varchar(100)
AS

select * from qb_importance
where importance_listid = @listid
GO



CREATE procedure insert_importance_sp
@listid varchar(100),
@type integer,
@name varchar(100),
@user varchar(100)
as

insert qb_importance
(importance_listid, importance_type, importance_name, importance_created_date, importance_created_by)
values
(@listid, @type, @name, getdate(), @user)
GO



CREATE procedure update_importance_sp
@listid varchar(100),
@type integer,
@name varchar(100),
@user varchar(100)
as

update qb_importance
set importance_modified_date = getdate(),
importance_modified_by = @user, 
importance_type = @type, 
importance_name = @name
where importance_listid = @listid
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Insert_web_reg_list_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Insert_web_reg_list_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_user_drive_records_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_user_drive_records_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_blank_end_datestamp_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_blank_end_datestamp_sp]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE procedure Insert_web_reg_list_sp
@reg_list_hkey varchar(100),
@reg_list_path varchar(1000),
@reg_list_key	varchar(500),
@reg_list_type varchar(20),
@reg_list_value varchar(1000),
@reg_list_notes varchar(1000),
@reg_list_section varchar(50),
@reg_list_user varchar(50),
@reg_list_computer varchar(100)
as

insert reg_list
(reg_list_hkey, reg_list_path, reg_list_key, reg_list_type, reg_list_value,
reg_list_notes, reg_list_section, reg_list_user, reg_list_computer, reg_list_datestamp, reg_list_active)
values
(@reg_list_hkey, @reg_list_path, @reg_list_key, @reg_list_type, @reg_list_value,
@reg_list_notes, @reg_list_section, @reg_list_user, @reg_list_computer, getdate(), '1')
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE grab_user_drive_records_sp
@user varchar(50)
AS

select * from user_drives
where drive_user = @user and drive_active ='1'
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

CREATE procedure update_blank_end_datestamp_sp
@computer_current_user varchar(100)
as

update internal_computers
set computer_end_datestamp = getdate()
where computer_current_user = @computer_current_user and computer_end_datestamp is null
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE grab_note_sp
@listid varchar(100)
AS

select * from qb_note
where note_listid = @listid
GO



CREATE procedure insert_note_sp
@listid varchar(100),
@msg varchar(2000),
@user varchar(100)
as

insert qb_note
(note_listid, note_msg, note_created_by, note_datestamp, note_msg_backup)
values
(@listid, @msg, @user, getdate(), @msg)
GO



CREATE procedure update_note_sp
@listid varchar(100),
@msg varchar(2000),
@user varchar(100)
as

update qb_note
set note_modified_date = getdate(),
note_modified_by = @user, 
note_msg = @msg
where note_listid = @listid
GO

CREATE PROCEDURE grab_importance_sp
@listid varchar(100)
AS

select * from qb_importance
where importance_listid = @listid
GO



CREATE procedure insert_importance_sp
@listid varchar(100),
@type integer,
@name varchar(100),
@user varchar(100)
as

insert qb_importance
(importance_listid, importance_type, importance_name, importance_created_date, importance_created_by)
values
(@listid, @type, @name, getdate(), @user)
GO



CREATE procedure update_importance_sp
@listid varchar(100),
@type integer,
@name varchar(100),
@user varchar(100)
as

update qb_importance
set importance_modified_date = getdate(),
importance_modified_by = @user, 
importance_type = @type, 
importance_name = @name
where importance_listid = @listid
GO


CREATE procedure grab_note_callback_from_date_sp
@callback_date varchar(100)
AS

select * from qb_note
where note_callback_date = @callback_date
order by note_callback_time asc
go



CREATE procedure update_upfront_sp
@listid varchar(100),
@upfront varchar(20)
as

update qb_importance
set upfront = @upfront
where importance_listid = @listid
GO