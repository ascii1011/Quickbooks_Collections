if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Insert_web_reg_list_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Insert_web_reg_list_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[check_4_active_kill_process_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[check_4_active_kill_process_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[check_4_all_active_process_by_user_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[check_4_all_active_process_by_user_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[check_4_all_active_process_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[check_4_all_active_process_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[create_checkin_record_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[create_checkin_record_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[create_xinv_payment_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[create_xinv_payment_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[get_updatemgr_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[get_updatemgr_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_active_reg_list_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_active_reg_list_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_active_user_reg_list_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_active_user_reg_list_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_active_users_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_active_users_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_all_drive_records_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_all_drive_records_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_all_user_records_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_all_user_records_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_allactive_drive_records_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_allactive_drive_records_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_bad_process_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_bad_process_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_callback_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_callback_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_good_process_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_good_process_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_importance_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_importance_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_note_by_date_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_note_by_date_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_note_callback_from_date_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_note_callback_from_date_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_note_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_note_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_old_drive_records_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_old_drive_records_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_old_records_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_old_records_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_profile_dtls_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_profile_dtls_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_profile_names_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_profile_names_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_profile_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_profile_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_reg_list_by_user_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_reg_list_by_user_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_reg_list_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_reg_list_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_svr_checkin_down_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_svr_checkin_down_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[grab_user_drive_records_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[grab_user_drive_records_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_bad_user_process_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_bad_user_process_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_callback_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_callback_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_computer_info_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_computer_info_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_good_process_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_good_process_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_importance_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_importance_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_note_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_note_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_user_drives_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_user_drives_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_user_process_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_user_process_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_xcust_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_xcust_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_xinv_lineitem_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_xinv_lineitem_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_xinv_payment_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_xinv_payment_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_xinv_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_xinv_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[insert_xrep_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[insert_xrep_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[request_all_server_info_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[request_all_server_info_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_blank_end_datestamp_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_blank_end_datestamp_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_callback_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_callback_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_checkin_record_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_checkin_record_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_computer_checkin_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_computer_checkin_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_decline_unlive_updatemgr_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_decline_unlive_updatemgr_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_importance_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_importance_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_kill_process_result_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_kill_process_result_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_kill_process_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_kill_process_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_live_updatemgr_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_live_updatemgr_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_logout_checkin_record_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_logout_checkin_record_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_note_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_note_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_profile_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_profile_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_success_unlive_updatemgr_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_success_unlive_updatemgr_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_to_kill_process_directly_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_to_kill_process_directly_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_to_kill_process_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_to_kill_process_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_unactivate_reg_list_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_unactivate_reg_list_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_unactivate_srv_checkin_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_unactivate_srv_checkin_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_unactivate_user_drives_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_unactivate_user_drives_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_upfront_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_upfront_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_user_process_not_current_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_user_process_not_current_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_user_process_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_user_process_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_xcust_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_xcust_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_xinv_lineitem_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_xinv_lineitem_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_xinv_payment_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_xinv_payment_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_xinv_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_xinv_sp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[update_xrep_sp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[update_xrep_sp]
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


CREATE procedure check_4_active_kill_process_sp
@up_user varchar(100)
as

select * 
from user_process
where up_active = '1' and up_kill_process = '1' and up_kill_result = '0' and up_kill_datestamp is null and up_user = @up_user

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure check_4_all_active_process_by_user_sp
@up_user varchar(100)
as

select * 
from user_process
where up_active = '1' and up_user = @up_user
order by up_name asc

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE check_4_all_active_process_sp
as

select * from user_process
where up_active = '1'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE create_checkin_record_sp
@name varchar(50),
@ip varchar(50),
@mac varchar(50),
@hdd varchar(2000)
AS

insert into srv_checkin
(srv_name, srv_ip, srv_mac, srv_created, srv_active, srv_harddrive)
values
(@name, @ip, @mac, getdate(), '1', @hdd)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure create_xinv_payment_sp

    @inv_pay_bit varchar(200),
    @inv_pay_txnid varchar(200),
    @inv_pay_txntype varchar(200),
    @inv_pay_txndate datetime,
    @inv_pay_amount varchar(200),
    @inv_pay_refnumber varchar(100),
    @inv_pay_linktype varchar(100),
    @inv_pay_type varchar(100)
as 

insert qbx_inv_payments
 
(inv_pay_bit, inv_pay_txnid, inv_pay_txntype, inv_pay_txndate, inv_pay_amount,
inv_pay_refnumber, inv_pay_linktype, inv_pay_type)
values
(@inv_pay_bit, @inv_pay_txnid, @inv_pay_txntype, @inv_pay_txndate, @inv_pay_amount,
@inv_pay_refnumber, @inv_pay_linktype, @inv_pay_type)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure get_updatemgr_sp
as

select * from qb_updatemgr
where updatemgr_index = '6758'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_active_reg_list_sp
AS

SELECT 
reg_list_index as cindex, 
reg_list_hkey as hkey, 
reg_list_path as path, 
reg_list_key as ckey, 
reg_list_type as type, 
reg_list_value as value,
reg_list_datestamp as date,
reg_list_notes as notes,
reg_list_section as csection,
reg_list_user as cuser,
reg_list_computer as puter

FROM reg_list

where reg_list_active = '1'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_active_user_reg_list_sp
@cuser 	varchar(100)
AS

SELECT 
reg_list_index as cindex, 
reg_list_hkey as hkey, 
reg_list_path as path, 
reg_list_key as ckey, 
reg_list_type as type, 
reg_list_value as value,
reg_list_datestamp as date,
reg_list_notes as notes,
reg_list_section as csection,
reg_list_user as cuser,
reg_list_computer as puter

FROM reg_list

where reg_list_active = '1' and reg_list_user = @cuser

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_active_users_sp
AS

select * from internal_computers
where computer_end_datestamp is null
order by computer_current_user asc

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_all_drive_records_sp
AS

select * from user_drives

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_all_user_records_sp
AS

select * from internal_computers
order by computer_current_user asc

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_allactive_drive_records_sp
AS

select * from user_drives
where drive_active ='1'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_bad_process_sp
AS

SELECT  *
FROM bad_process

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE grab_callback_sp
@listid varchar(100)
AS

select * from qb_callback
where callback_listid = @listid


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_good_process_sp
AS

SELECT  *
FROM good_process

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE grab_importance_sp
@listid varchar(100)
AS

select * from qb_importance
where importance_listid = @listid


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_note_by_date_sp
@listid varchar(100),
@ddate datetime
AS

select * from qb_note
where note_listid = @listid and note_datestamp > @ddate
order by note_datestamp

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_note_callback_from_date_sp
@lookupdate varchar(15)
AS

select * from qb_note
where note_callback_date = @lookupdate 
order by note_callback_time asc

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE grab_note_sp
@listid varchar(100)
AS

select * from qb_note
where note_listid = @listid
order by note_datestamp desc

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_old_drive_records_sp
AS

select * from user_drives
where drive_active ='0'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_old_records_sp
AS

select * from internal_computers
where computer_end_datestamp is not null
order by computer_current_user asc, computer_start_datestamp asc

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure grab_profile_dtls_sp as

select * from qb_attr

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure grab_profile_names_sp as

select * from qb_profiles
where profiles_level > '0'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE procedure grab_profile_sp
@username as varchar(50)
as

select * from qb_profiles a, qb_features b
where a.profiles_username = @username and  a.profiles_index = b.features_index

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE grab_reg_list_by_user_sp
@reg_list_user varchar(50)
AS

SELECT 
reg_list_index, 
reg_list_hkey, 
reg_list_path, 
reg_list_key, 
reg_list_type, 
reg_list_value,
reg_list_datestamp,
reg_list_notes,
reg_list_section,
reg_list_user,
reg_list_computer

FROM reg_list 
where reg_list_user = @reg_list_user and reg_list_active = '1'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE grab_reg_list_sp
AS

SELECT 
reg_list_index, 
reg_list_hkey, 
reg_list_path, 
reg_list_key, 
reg_list_type, 
reg_list_value,
reg_list_datestamp,
reg_list_notes,
reg_list_section,
reg_list_user,
reg_list_computer

FROM reg_list


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE grab_svr_checkin_down_sp
AS

SELECT srv_name,  srv_created, srv_checkin, srv_stopped, srv_mac, srv_ip, srv_harddrive, srv_memory, DATEDIFF(second, srv_checkin, getdate()) AS seconds
FROM srv_checkin
where srv_active = '1'

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


CREATE procedure insert_bad_user_process_sp
@up_pid integer,
@up_name varchar (255),
@up_user varchar(100),
@up_computer varchar(100),
@up_type varchar (20),
@up_executioner varchar(50)
AS

insert user_process
(up_name, up_user, up_computer, 
up_type, up_datestamp, up_execution_datestamp, 
up_kill_datestamp, up_executioner, up_kill_process,
up_active, up_kill_result, up_pid)
values
(@up_name, @up_user, @up_computer, 
@up_type, getdate(), getdate(), 
getdate(), @up_executioner, '1',
'1', '0', @up_pid)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE procedure insert_callback_sp
@listid varchar(100),
@cb_date datetime,
@reason varchar(2000),
@user varchar(100)
as

insert qb_callback
(callback_listid, callback_callbackdate, callback_reason, callback_created_datestamp, callback_created_by, callback_completed)
values
(@listid, @cb_date, @reason, getdate(), @user, '0')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE procedure insert_computer_info_sp
@computer_name varchar (255),
--@computer_cpu varchar (100),
--@computer_memory_size varchar (50),
@computer_current_user varchar (100),
@computer_ip varchar (50),
@computer_mac varchar (50),
@computer_os varchar (100),
@computer_os_build varchar (100)
AS

insert internal_computers
(computer_name, computer_current_user, 
computer_ip, computer_mac, computer_os, computer_os_build, computer_start_datestamp)
values
(@computer_name, @computer_current_user, 
@computer_ip, @computer_mac, @computer_os, @computer_os_build, getdate())

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE insert_good_process_sp
@gp_name varchar(200),
@gp_level varchar(2),
@gp_desc varchar(200)
AS

insert good_process
(gp_name, gp_level, gp_desc)
values
(@gp_name, @gp_level, @gp_desc)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE procedure insert_importance_sp
@listid varchar(100),
@type integer,
@name varchar(100),
@user varchar(100)
as

insert qbx_cust
(cust_listid, importance_type, importance_upfront, importance_created_date, importance_created_by, importance_name)
values
(@listid, @type, 'True', getdate(), @user, @name)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE procedure insert_note_sp
@listid varchar(100),
@msg varchar(2000),
@callbackdate varchar(15),
@callbacktime varchar(15),
@callbackcompany varchar(200),
@user varchar(100),
@amount varchar(10),
@jobstatus varchar(50)
as

insert qb_note
(note_listid, note_msg, note_callback_date, note_callback_time, note_company_name,  note_created_by, note_datestamp, note_msg_backup, note_company_amount, note_company_status)
values
(@listid, @msg, @callbackdate, @callbacktime, @callbackcompany, @user, getdate(), @msg, @amount, @jobstatus)
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE insert_user_drives_sp
@letter varchar(50),
@available  varchar(50),
@type  varchar(50),
@filesystem  varchar(50),
@freespace varchar(50),
@isready  varchar(50),
@path  varchar(2000),
@rootfolder  varchar(2000),
@serialnumber  varchar(255),
@sharename  varchar(255),
@totalsize varchar(50),
@volumename  varchar(50),
@subfolderscount  varchar(50),
@user  varchar(50)
 AS

insert user_drives
(drive_letter, drive_available, drive_type, drive_filesystem,
drive_freespace, drive_isready, drive_path, drive_rootfolder,
drive_serialnumber, drive_sharename, drive_totalsize, drive_volumename,
drive_subfolderscount, drive_user, drive_datestamp, drive_active)
values
(@letter, @available, @type, @filesystem,
@freespace, @isready, @path, @rootfolder,
@serialnumber, @sharename, @totalsize, @volumename,
@subfolderscount, @user, getdate(), '1')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure insert_user_process_sp
@up_pid integer,
@up_name varchar (255),
@up_user varchar(100),
@up_computer varchar(100),
@up_type varchar (20)

AS

insert user_process
(up_pid, up_name, up_user, up_computer, up_type, up_datestamp)
values
(@up_pid, @up_name, @up_user, @up_computer, @up_type, getdate())

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure insert_xcust_sp
    
@cust_bit varchar(100),
@cust_index varchar(100),
@cust_listid varchar(100),
@cust_name varchar(100),
@cust_fullname varchar(100),
@cust_companyname varchar(100),
@cust_contact varchar(100),
@cust_salutation varchar(10),
@cust_firstname varchar(50),
@cust_middlename varchar(50),
@cust_lastname varchar(50),
@cust_billaddress_add1 varchar(200),
@cust_billaddress_add2 varchar(200),
@cust_billaddress_add3 varchar(200),
@cust_billaddress_add4 varchar(200),
@cust_billcity varchar(100),
@cust_billstate varchar(100),
@cust_billpostalcode varchar(12),
@cust_billcountry varchar(100),
@cust_shipaddress_add1 varchar(200),
@cust_shipaddress_add2 varchar(200),
@cust_shipaddress_add3 varchar(200),
@cust_shipaddress_add4 varchar(200),
@cust_shipcity varchar(100),
@cust_shipstate varchar(100),
@cust_shippostalcode varchar(12),
@cust_shipcountry varchar(100),
@cust_phone1 varchar(30),
@cust_phone2 varchar(30),
@cust_fax1 varchar(30),
@cust_fax2 varchar(30),
@cust_email1 varchar(100),
@cust_email2 varchar(100),
@cust_notes varchar(2000),
@cust_isactive varchar(20),
@cust_sublevel varchar(20),
@cust_timecreated datetime,
@cust_timemodified datetime,
@cust_jobstatus varchar(50),
@cust_JobStartDate datetime,
@cust_status varchar(50),
@cust_balance varchar(50),
@cust_totalbalance varchar(50),
@cust_accountnumber varchar(50),
@cust_termsref_listid varchar(50),
@cust_termsref_fullname varchar(100),
@cust_customertyperef_listid varchar(50),
@cust_customertyperef_residential varchar(100),
@cust_salestaxcoderef_listid varchar(50),
@cust_salestaxcoderef_fullname varchar(100),
@cust_itemsalestaxref_listid varchar(50),
@cust_itemsalestaxref_fullname varchar(100),
@cust_salesrepref_listid varchar(50),
@cust_salesrepref_fullname varchar(100),
@cust_altcontact varchar(100),
@importance_type integer
as

insert qbx_cust 


(cust_bit, cust_index, cust_listid, cust_name, cust_fullname, cust_companyname, cust_contact, cust_salutation,
cust_firstname, cust_middlename, cust_lastname, cust_billaddress_add1, cust_billaddress_add2, cust_billaddress_add3, cust_billaddress_add4,
cust_billcity, cust_billstate, cust_billpostalcode, cust_billcountry, cust_shipaddress_add1, cust_shipaddress_add2, cust_shipaddress_add3, cust_shipaddress_add4,
cust_shipcity, cust_shipstate, cust_shippostalcode, cust_shipcountry, cust_phone1, cust_phone2, cust_fax1, cust_fax2,
cust_email1, cust_email2, cust_notes, cust_isactive, cust_sublevel, cust_timecreated, cust_timemodified, cust_jobstatus,
cust_JobStartDate, cust_status, cust_balance, cust_totalbalance, cust_accountnumber, cust_termsref_listid, cust_termsref_fullname,
cust_customertyperef_listid, cust_customertyperef_residential, cust_salestaxcoderef_listid, cust_salestaxcoderef_fullname,
cust_itemsalestaxref_listid, cust_itemsalestaxref_fullname, cust_salesrepref_listid, cust_salesrepref_fullname, cust_altcontact, importance_type) --, 
--cust_balance_money, cust_totalbalance_money, cust_accountnumber_numeric)
values
(@cust_bit, @cust_index, @cust_listid, @cust_name, @cust_fullname, @cust_companyname, @cust_contact, @cust_salutation,
@cust_firstname, @cust_middlename, @cust_lastname, @cust_billaddress_add1, @cust_billaddress_add2, @cust_billaddress_add3, @cust_billaddress_add4,
@cust_billcity, @cust_billstate, @cust_billpostalcode, @cust_billcountry, @cust_shipaddress_add1, @cust_shipaddress_add2, @cust_shipaddress_add3, @cust_shipaddress_add4,
@cust_shipcity, @cust_shipstate, @cust_shippostalcode, @cust_shipcountry, @cust_phone1, @cust_phone2, @cust_fax1, @cust_fax2,
@cust_email1, @cust_email2, @cust_notes, @cust_isactive, @cust_sublevel, @cust_timecreated, @cust_timemodified, @cust_jobstatus,
@cust_JobStartDate, @cust_status, @cust_balance, @cust_totalbalance, @cust_accountnumber, @cust_termsref_listid, @cust_termsref_fullname,
@cust_customertyperef_listid, @cust_customertyperef_residential, @cust_salestaxcoderef_listid, @cust_salestaxcoderef_fullname,
@cust_itemsalestaxref_listid, @cust_itemsalestaxref_fullname, @cust_salesrepref_listid, @cust_salesrepref_fullname, @cust_altcontact, @importance_type) --,
--CONVERT(money, '@cust_balance'), CONVERT(money, '@cust_totalbalance') , CONVERT(int, '@cust_accountnumber')  )
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure insert_xinv_lineitem_sp
    
    @inv_txnid_link varchar(200),
    @inv_line_bit varchar(200),
    @inv_line_txnlineid varchar(20),
    @inv_line_itemref_listid varchar(50),
    @inv_line_itemref_fullname varchar(50),
    @inv_line_desc varchar(500),
    @inv_line_quantity varchar(20),
    @inv_line_rate varchar(20),
    @inv_line_amount varchar(20),
    @inv_line_salestaxcoderef_listid varchar(50),
    @inv_line_salestaxcoderef_fullname varchar(100)
as 

insert qbx_inv_lineitems
 
(inv_txnid_link, inv_line_bit, inv_line_txnlineid, inv_line_itemref_listid, inv_line_itemref_fullname,
inv_line_desc, inv_line_quantity, inv_line_rate, inv_line_amount, 
inv_line_salestaxcoderef_listid, inv_line_salestaxcoderef_fullname)
values
(@inv_txnid_link, @inv_line_bit, @inv_line_txnlineid, @inv_line_itemref_listid, @inv_line_itemref_fullname,
@inv_line_desc, @inv_line_quantity, @inv_line_rate, @inv_line_amount, 
@inv_line_salestaxcoderef_listid, @inv_line_salestaxcoderef_fullname)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure insert_xinv_payment_sp

    @inv_txnid_link varchar(200),
    @inv_pay_bit varchar(200),
    @inv_pay_txnid varchar(200),
    @inv_pay_txntype varchar(200),
    @inv_pay_txndate datetime,
    @inv_pay_amount varchar(200),
    @inv_pay_refnumber varchar(100),
    @inv_pay_linktype varchar(100),
    @inv_pay_type varchar(100)
as 

insert qbx_inv_payments
(inv_txnid_link, inv_pay_bit, inv_pay_txnid, inv_pay_txntype, inv_pay_txndate,
inv_pay_amount, inv_pay_refnumber, inv_pay_linktype, inv_pay_type)
values
(@inv_txnid_link, @inv_pay_bit, @inv_pay_txnid, @inv_pay_txntype, @inv_pay_txndate,
@inv_pay_amount, @inv_pay_refnumber, @inv_pay_linktype, @inv_pay_type)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure insert_xinv_sp

    @inv_bit varchar(200),
    @inv_txnid varchar(200),
    @inv_timecreated datetime,
    @inv_timemodified datetime,
    @inv_txnnumber varchar(200),
    @inv_customerref_listid varchar(200),
    @inv_customerref_fullname varchar(200),
    @inv_classref_listid varchar(200),
    @inv_classref_fullname varchar(200),
    @inv_araccountref_listid varchar(200),
    @inv_araccountref_fullname varchar(200),
    @inv_txndate datetime,
    @inv_refnumber varchar(200),
    @inv_billaddress_add1 varchar(200),
    @inv_billaddress_add2 varchar(200),
    @inv_billaddress_add3 varchar(200),
    @inv_billaddress_add4 varchar(200),
    @inv_billcity varchar(100),
    @inv_billstate varchar(100),
    @inv_billpostalcode varchar(12),
    @inv_billcountry varchar(200),
    @inv_shipaddress_add1 varchar(200),
    @inv_shipaddress_add2 varchar(200),
    @inv_shipaddress_add3 varchar(200),
    @inv_shipaddress_add4 varchar(200),
    @inv_shipcity varchar(100),
    @inv_shipstate varchar(100),
    @inv_shippostalcode varchar(12),
    @inv_shipcountry varchar(200),
    @inv_ispending varchar(10),
    @inv_isfinancecharge varchar(10),
    @inv_termsref_listid varchar(50),
    @inv_termsref_fullname varchar(100),
    @inv_duedate datetime,
    @inv_salesrepref_listid varchar(50),
    @inv_salesrepref_fullname varchar(100),
    @inv_shipdate datetime,
    @inv_subtotal varchar(20),
    @inv_salestaxpercentage varchar(20),
    @inv_salestaxtotal varchar(20),
    @inv_appliedamount varchar(20),
    @inv_balanceremaining varchar(20),
    @inv_customermsgref_listid varchar(50),
    @inv_customermsgref_fullname varchar(100),
    @inv_istobeprinted varchar(10),
    @inv_enabled varchar(1)
as

insert qbx_inv 

(inv_bit, inv_txnid, inv_timecreated, inv_timemodified, inv_txnnumber, inv_customerref_listid,
inv_customerref_fullname, inv_classref_listid, inv_classref_fullname, inv_araccountref_listid,
inv_araccountref_fullname, inv_txndate, inv_refnumber, inv_billaddress_add1, inv_billaddress_add2,
inv_billaddress_add3, inv_billaddress_add4, inv_billcity, inv_billstate, inv_billpostalcode, inv_billcountry,
inv_shipaddress_add1, inv_shipaddress_add2, inv_shipaddress_add3, inv_shipaddress_add4, inv_shipcity,
inv_shipstate, inv_shippostalcode, inv_shipcountry, inv_ispending, inv_isfinancecharge, inv_termsref_listid,
inv_termsref_fullname, inv_duedate, inv_salesrepref_listid, inv_salesrepref_fullname, inv_shipdate,
inv_subtotal, inv_salestaxpercentage, inv_salestaxtotal, inv_appliedamount, inv_balanceremaining,
inv_customermsgref_listid, inv_customermsgref_fullname, inv_istobeprinted, inv_enabled)
values
(@inv_bit, @inv_txnid, @inv_timecreated, @inv_timemodified, @inv_txnnumber, @inv_customerref_listid,
@inv_customerref_fullname, @inv_classref_listid, @inv_classref_fullname, @inv_araccountref_listid,
@inv_araccountref_fullname, @inv_txndate, @inv_refnumber, @inv_billaddress_add1, @inv_billaddress_add2,
@inv_billaddress_add3, @inv_billaddress_add4, @inv_billcity, @inv_billstate, @inv_billpostalcode, @inv_billcountry,
@inv_shipaddress_add1, @inv_shipaddress_add2, @inv_shipaddress_add3, @inv_shipaddress_add4, @inv_shipcity,
@inv_shipstate, @inv_shippostalcode, @inv_shipcountry, @inv_ispending, @inv_isfinancecharge, @inv_termsref_listid,
@inv_termsref_fullname, @inv_duedate, @inv_salesrepref_listid, @inv_salesrepref_fullname, @inv_shipdate,
@inv_subtotal, @inv_salestaxpercentage, @inv_salestaxtotal, @inv_appliedamount, @inv_balanceremaining,
@inv_customermsgref_listid, @inv_customermsgref_fullname, @inv_istobeprinted, @inv_enabled)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure insert_xrep_sp
    @rep_listid varchar(200),
    @rep_initial varchar(50),
    @rep_isactive varchar(50),
    @rep_salesrepentityref_listid varchar(200),
    @rep_salesrepentityref_fullname varchar(100)
as 

insert qbx_reps
 
(rep_listid, rep_initial, rep_isactive, rep_salesrepentityref_listid, rep_salesrepentityref_fullname)
values
(@rep_listid, @rep_initial, @rep_isactive, @rep_salesrepentityref_listid, @rep_salesrepentityref_fullname)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE request_all_server_info_sp
AS

select * from systems
where sys_active = '1'

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

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE procedure update_callback_sp
@index varchar(100),
@followup varchar(2000),
@completed bit,
@user varchar(100)
as

update qb_callback
set callback_completed_date = getdate(),
callback_completed_by = @user, 
callback_followup = @followup,
callback_completed = @completed
where callback_index = @index


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE update_checkin_record_sp
@name varchar(50),
@ip varchar(50),
@mac varchar(50),
@hdd varchar(2000)
AS

update srv_checkin

set srv_checkin = getdate(),
srv_harddrive = @hdd

where srv_stopped is null and 
srv_name = @name and 
srv_ip = @ip and 
srv_mac = @mac and 
srv_active ='1'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_computer_checkin_sp
@computer_current_user varchar (100)
AS

update internal_computers
set computer_last_checkin = getdate()
where computer_current_user = @computer_current_user and computer_end_datestamp is null

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_decline_unlive_updatemgr_sp
as

update qb_updatemgr
set updatemgr_last_timestamp = getdate(),
updatemgr_value = '0',
updatemgr_status = '0'
where updatemgr_index = '6758'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE procedure update_importance_sp
@listid varchar(100),
@type integer,
@name varchar(100),
@user varchar(100),
@upfront varchar(10)
as

update qbx_cust
set importance_modified_date = getdate(),
importance_modified_by = @user, 
importance_type = @type, 
importance_name = @name,
importance_upfront = @upfront
where cust_listid = @listid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_kill_process_result_sp 
@up_pid integer,
@up_user varchar(50),
@up_name varchar(100),
@up_kill_result varchar(5)
as

update user_process
set up_kill_result = @up_kill_result
where up_user = @up_user and
up_name = @up_name and 
up_active = '1' and 
up_kill_process= '1' and 
up_kill_result = '0' and 
up_pid = @up_pid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_kill_process_sp
@up_index integer,
@up_user varchar(100),
@up_kill_result varchar(1)
 as

update user_process
set up_kill_datestamp = getdate(),
up_kill_result = @up_kill_result, up_active = '0'
where up_user = @up_user and up_index = @up_index

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_live_updatemgr_sp
as

update qb_updatemgr
set updatemgr_timestamp = getdate(),
updatemgr_value = '1'
where updatemgr_index = '6758'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE PROCEDURE update_logout_checkin_record_sp
@name varchar(50),
@ip varchar(50),
@mac varchar(50)
AS

update srv_checkin

set srv_stopped = getdate()

where srv_checkin is null and 
srv_stopped is null and 
srv_name = @name and 
srv_ip = @ip and 
srv_mac = @mac and 
srv_active ='1'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
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
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_profile_sp
@username as varchar(50), 
@value	as integer
as
update qb_profiles
set profiles_enabled = @value 
where profiles_username = @username

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_success_unlive_updatemgr_sp
as

update qb_updatemgr
set updatemgr_last_timestamp = getdate(),
updatemgr_last_success_timestamp = getdate(),
updatemgr_value = '0',
updatemgr_status = '1'
where updatemgr_index = '6758'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_to_kill_process_directly_sp
@up_index integer,
@up_executioner varchar(50),
@up_user varchar(100),
@up_execution_datestamp datetime,
@up_kill_process bit
 as

update user_process
set up_execution_datestamp = getdate(),
up_kill_datestamp = getdate(),
up_kill_result = '1',
up_active = '0',
up_kill_process = @up_kill_process, 
up_executioner = @up_executioner

where up_user = @up_user and up_index = @up_index and up_active='1'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_to_kill_process_sp
@up_index integer,
@up_user varchar(100),
@up_execution_datestamp datetime,
@up_kill_process bit
 as

update user_process
set up_execution_datestamp = @up_execution_datestamp,
up_kill_process = @up_kill_process
where up_user = @up_user and up_index = @up_index

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


CREATE procedure update_unactivate_reg_list_sp
@reg_list_user varchar(50),
@reg_list_section varchar(50)
as

update reg_list
set reg_list_active = '0'
where reg_list_user = @reg_list_user and reg_list_section = @reg_list_section

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_unactivate_srv_checkin_sp
@name varchar(50),
@ip varchar(50),
@mac varchar(50)
as

update srv_checkin
set srv_active = '0'
where srv_name = @name and srv_mac = @mac and srv_ip = @ip and srv_active = '1'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_unactivate_user_drives_sp
@user varchar(50)
as

update user_drives
set drive_active = '0'
where drive_user = @user and drive_active = '1'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO





CREATE procedure update_upfront_sp
@listid varchar(100),
@upfront varchar(10)
as

update qbx_cust
set importance_upfront = @upfront
where cust_listid = @listid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


create procedure update_user_process_not_current_sp
@up_user varchar(50)
as

update user_process
set up_active = '0'
where up_user = @up_user

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_user_process_sp
@up_pid integer,
@up_name varchar(100),
@up_user varchar(50),
@up_type varchar(20)
as

update user_process
set up_type  = @up_type, up_active = '0'
where up_pid = @up_pid and
up_name = @up_name and 
up_user = @up_user and 
up_active = '1'

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_xcust_sp
    
@cust_bit varchar(100),
@cust_index varchar(100),
@cust_listid varchar(100),
@cust_name varchar(100),
@cust_fullname varchar(100),
@cust_companyname varchar(100),
@cust_contact varchar(100),
@cust_salutation varchar(10),
@cust_firstname varchar(50),
@cust_middlename varchar(50),
@cust_lastname varchar(50),
@cust_billaddress_add1 varchar(200),
@cust_billaddress_add2 varchar(200),
@cust_billaddress_add3 varchar(200),
@cust_billaddress_add4 varchar(200),
@cust_billcity varchar(100),
@cust_billstate varchar(100),
@cust_billpostalcode varchar(12),
@cust_billcountry varchar(100),
@cust_shipaddress_add1 varchar(200),
@cust_shipaddress_add2 varchar(200),
@cust_shipaddress_add3 varchar(200),
@cust_shipaddress_add4 varchar(200),
@cust_shipcity varchar(100),
@cust_shipstate varchar(100),
@cust_shippostalcode varchar(12),
@cust_shipcountry varchar(100),
@cust_phone1 varchar(30),
@cust_phone2 varchar(30),
@cust_fax1 varchar(30),
@cust_fax2 varchar(30),
@cust_email1 varchar(100),
@cust_email2 varchar(100),
@cust_notes varchar(2000),
@cust_isactive varchar(20),
@cust_sublevel varchar(20),
@cust_timecreated datetime,
@cust_timemodified datetime,
@cust_jobstatus varchar(50),
@cust_JobStartDate datetime,
@cust_status varchar(50),
@cust_balance varchar(50),
@cust_totalbalance varchar(50),
@cust_accountnumber varchar(50),
@cust_termsref_listid varchar(50),
@cust_termsref_fullname varchar(100),
@cust_customertyperef_listid varchar(50),
@cust_customertyperef_residential varchar(100),
@cust_salestaxcoderef_listid varchar(50),
@cust_salestaxcoderef_fullname varchar(100),
@cust_itemsalestaxref_listid varchar(50),
@cust_itemsalestaxref_fullname varchar(100),
@cust_salesrepref_listid varchar(50),
@cust_salesrepref_fullname varchar(100),
@cust_altcontact varchar(100)--,
--@cust_balance_money varchar(50)--,
--@cust_totalbalance_money numeric,
--@cust_accountnumber_numeric numeric
as

update qbx_cust 
set 
cust_bit = @cust_bit,
cust_index = @cust_index,
cust_name = @cust_name,
cust_fullname = @cust_fullname,
cust_companyname = @cust_companyname,
cust_contact = @cust_contact,
cust_salutation = @cust_salutation,
cust_firstname = @cust_firstname,
cust_middlename = @cust_middlename,
cust_lastname = @cust_lastname,
cust_billaddress_add1 = @cust_billaddress_add1,
cust_billaddress_add2 = @cust_billaddress_add2,
cust_billaddress_add3 = @cust_billaddress_add3,
cust_billaddress_add4 = @cust_billaddress_add4,
cust_billcity = @cust_billcity,
cust_billstate = @cust_billstate,
cust_billpostalcode = @cust_billpostalcode,
cust_billcountry = @cust_billcountry,
cust_shipaddress_add1 = @cust_shipaddress_add1,
cust_shipaddress_add2 = @cust_shipaddress_add2,
cust_shipaddress_add3 = @cust_shipaddress_add3,
cust_shipaddress_add4 = @cust_shipaddress_add4,
cust_shipcity = @cust_shipcity,
cust_shipstate = @cust_shipstate,
cust_shippostalcode = @cust_shippostalcode,
cust_shipcountry = @cust_shipcountry,
cust_phone1 = @cust_phone1,
cust_phone2 = @cust_phone2,
cust_fax1 = @cust_fax1,
cust_fax2 = @cust_fax2,
cust_email1 = @cust_email1,
cust_email2 = @cust_email2,
cust_notes = @cust_notes,
cust_isactive = @cust_isactive,
cust_sublevel = @cust_sublevel,
cust_timecreated = @cust_timecreated,
cust_timemodified = @cust_timemodified,
cust_jobstatus = @cust_jobstatus,
cust_JobStartDate = @cust_JobStartDate,
cust_status = @cust_status,
cust_balance = @cust_balance,
cust_totalbalance = @cust_totalbalance,
cust_accountnumber = @cust_accountnumber,
cust_termsref_listid = @cust_termsref_listid,
cust_termsref_fullname = @cust_termsref_fullname,
cust_customertyperef_listid = @cust_customertyperef_listid,
cust_customertyperef_residential = @cust_customertyperef_residential,
cust_salestaxcoderef_listid = @cust_salestaxcoderef_listid,
cust_salestaxcoderef_fullname = @cust_salestaxcoderef_fullname,
cust_itemsalestaxref_listid = @cust_itemsalestaxref_listid,
cust_itemsalestaxref_fullname = @cust_itemsalestaxref_fullname,
cust_salesrepref_listid = @cust_salesrepref_listid,
cust_salesrepref_fullname = @cust_salesrepref_fullname,
cust_altcontact = @cust_altcontact--,
--cust_balance_money = CONVERT(money, '@cust_balance_money')--, 
--cust_totalbalance_money = CONVERT(money, '@cust_totalbalance_money'), 
--cust_accountnumber_numeric = CONVERT(int, '@cust_accountnumber_numeric')

where cust_listid = @cust_listid
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_xinv_lineitem_sp

    @inv_txnid_link varchar(200),
    @inv_line_bit varchar(200),
    @inv_line_txnlineid varchar(20),
    @inv_line_itemref_listid varchar(50),
    @inv_line_itemref_fullname varchar(50),
    @inv_line_desc varchar(500),
    @inv_line_quantity varchar(20),
    @inv_line_rate varchar(20),
    @inv_line_amount varchar(20),
    @inv_line_salestaxcoderef_listid varchar(50),
    @inv_line_salestaxcoderef_fullname varchar(100)
as 

update qbx_inv_lineitems 
set 
inv_txnid_link = @inv_txnid_link,
inv_line_bit = @inv_line_bit,
inv_line_itemref_listid = @inv_line_itemref_listid,
inv_line_itemref_fullname = @inv_line_itemref_fullname,
inv_line_desc = @inv_line_desc,
inv_line_quantity = @inv_line_quantity,
inv_line_rate = @inv_line_rate,
inv_line_amount = @inv_line_amount,
inv_line_salestaxcoderef_listid = @inv_line_salestaxcoderef_listid,
inv_line_salestaxcoderef_fullname = @inv_line_salestaxcoderef_fullname

where inv_line_txnlineid = @inv_line_txnlineid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_xinv_payment_sp

    @inv_txnid_link varchar(200),
    @inv_pay_bit varchar(200),
    @inv_pay_txnid varchar(200),
    @inv_pay_txntype varchar(200),
    @inv_pay_txndate datetime,
    @inv_pay_amount varchar(200),
    @inv_pay_refnumber varchar(100),
    @inv_pay_linktype varchar(100),
    @inv_pay_type varchar(100)
as 

update qbx_inv_payments
set
inv_txnid_link = @inv_txnid_link,
inv_pay_bit = @inv_pay_bit,
inv_pay_txntype = @inv_pay_txntype,
inv_pay_txndate = @inv_pay_txndate,
inv_pay_amount = @inv_pay_amount,
inv_pay_refnumber = @inv_pay_refnumber,
inv_pay_linktype = @inv_pay_linktype,
inv_pay_type = @inv_pay_type

where inv_pay_txnid = @inv_pay_txnid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_xinv_sp

    @inv_bit varchar(200),
    @inv_txnid varchar(200),
    @inv_timecreated datetime,
    @inv_timemodified datetime,
    @inv_txnnumber varchar(200),
    @inv_customerref_listid varchar(200),
    @inv_customerref_fullname varchar(200),
    @inv_classref_listid varchar(200),
    @inv_classref_fullname varchar(200),
    @inv_araccountref_listid varchar(200),
    @inv_araccountref_fullname varchar(200),
    @inv_txndate datetime,
    @inv_refnumber varchar(200),
    @inv_billaddress_add1 varchar(200),
    @inv_billaddress_add2 varchar(200),
    @inv_billaddress_add3 varchar(200),
    @inv_billaddress_add4 varchar(200),
    @inv_billcity varchar(100),
    @inv_billstate varchar(100),
    @inv_billpostalcode varchar(12),
    @inv_billcountry varchar(200),
    @inv_shipaddress_add1 varchar(200),
    @inv_shipaddress_add2 varchar(200),
    @inv_shipaddress_add3 varchar(200),
    @inv_shipaddress_add4 varchar(200),
    @inv_shipcity varchar(100),
    @inv_shipstate varchar(100),
    @inv_shippostalcode varchar(12),
    @inv_shipcountry varchar(200),
    @inv_ispending varchar(10),
    @inv_isfinancecharge varchar(10),
    @inv_termsref_listid varchar(50),
    @inv_termsref_fullname varchar(100),
    @inv_duedate datetime,
    @inv_salesrepref_listid varchar(50),
    @inv_salesrepref_fullname varchar(100),
    @inv_shipdate datetime,
    @inv_subtotal varchar(20),
    @inv_salestaxpercentage varchar(20),
    @inv_salestaxtotal varchar(20),
    @inv_appliedamount varchar(20),
    @inv_balanceremaining varchar(20),
    @inv_customermsgref_listid varchar(50),
    @inv_customermsgref_fullname varchar(100),
    @inv_istobeprinted varchar(10),
    @inv_enabled varchar(1)
as


update qbx_inv 
set 
inv_bit = @inv_bit,
inv_timecreated = @inv_timecreated,
inv_timemodified = @inv_timemodified,
inv_txnnumber = @inv_txnnumber,
inv_customerref_listid = @inv_customerref_listid,
inv_customerref_fullname = @inv_customerref_fullname,
inv_classref_listid = @inv_classref_listid,
inv_classref_fullname = @inv_classref_fullname,
inv_araccountref_listid = @inv_araccountref_listid,
inv_araccountref_fullname = @inv_araccountref_fullname,
inv_txndate = @inv_txndate,
inv_refnumber = @inv_refnumber,
inv_billaddress_add1 = @inv_billaddress_add1,
inv_billaddress_add2 = @inv_billaddress_add2,
inv_billaddress_add3 = @inv_billaddress_add3,
inv_billaddress_add4 = @inv_billaddress_add4,
inv_billcity = @inv_billcity,
inv_billstate = @inv_billstate,
inv_billpostalcode = @inv_billpostalcode,
inv_billcountry = @inv_billcountry,
inv_shipaddress_add1 = @inv_shipaddress_add1,
inv_shipaddress_add2 = @inv_shipaddress_add2,
inv_shipaddress_add3 = @inv_shipaddress_add3,
inv_shipaddress_add4 = @inv_shipaddress_add4,
inv_shipcity = @inv_shipcity,
inv_shipstate = @inv_shipstate,
inv_shippostalcode = @inv_shippostalcode,
inv_shipcountry = @inv_shipcountry,
inv_ispending = @inv_ispending,
inv_isfinancecharge = @inv_isfinancecharge,
inv_termsref_listid = @inv_termsref_listid,
inv_termsref_fullname = @inv_termsref_fullname,
inv_duedate = @inv_duedate,
inv_salesrepref_listid = @inv_salesrepref_listid,
inv_salesrepref_fullname = @inv_salesrepref_fullname,
inv_shipdate = @inv_shipdate,
inv_subtotal = @inv_subtotal,
inv_salestaxpercentage = @inv_salestaxpercentage,
inv_salestaxtotal = @inv_salestaxtotal,
inv_appliedamount = @inv_appliedamount,
inv_balanceremaining = @inv_balanceremaining,
inv_customermsgref_listid = @inv_customermsgref_listid,
inv_customermsgref_fullname = @inv_customermsgref_fullname,
inv_istobeprinted = @inv_istobeprinted,
inv_enabled = @inv_enabled

where inv_txnid = @inv_txnid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO


CREATE procedure update_xrep_sp
    @rep_listid varchar(200),
    @rep_initial varchar(50),
    @rep_isactive varchar(50),
    @rep_salesrepentityref_listid varchar(200),
    @rep_salesrepentityref_fullname varchar(100)
as 

update qbx_reps 
set 
rep_initial = @rep_initial,
rep_isactive = @rep_isactive,
rep_salesrepentityref_listid = @rep_salesrepentityref_listid,
rep_salesrepentityref_fullname = @rep_salesrepentityref_fullname

where rep_listid = @rep_listid

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

