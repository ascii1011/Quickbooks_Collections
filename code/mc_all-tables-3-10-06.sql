if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LDAP_LIST]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LDAP_LIST]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MailLog_items]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[MailLog_items]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[a_notes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[a_notes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[backup_day_objects]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[backup_day_objects]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[backup_days]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[backup_days]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[backup_emailaccounts]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[backup_emailaccounts]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[backup_groups]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[backup_groups]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[backup_properties]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[backup_properties]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[backup_sessions]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[backup_sessions]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[backup_users]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[backup_users]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[bad_process]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[bad_process]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[bad_process_file_location]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[bad_process_file_location]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[bad_process_files]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[bad_process_files]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[bad_process_registry_entries]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[bad_process_registry_entries]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[bad_process_websites]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[bad_process_websites]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[catch_this]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[catch_this]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[email_limit_messages]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[email_limit_messages]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[good_process]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[good_process]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[internal_computers]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[internal_computers]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[kill_user_process]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[kill_user_process]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[maillog_files]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[maillog_files]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_answer]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_answer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_client_cache_firefox]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_client_cache_firefox]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_client_cache_ie]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_client_cache_ie]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_client_events]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_client_events]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_client_keys]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_client_keys]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_client_processes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_client_processes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_client_services]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_client_services]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_client_wmi_hardware]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_client_wmi_hardware]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_client_wmiclass]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_client_wmiclass]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_functions]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_functions]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_messages]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_messages]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_requests]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_requests]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_send2fax_batch]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_send2fax_batch]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_send2fax_mailitems]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_send2fax_mailitems]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_send2fax_retries]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_send2fax_retries]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_send2fax_sent_items]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_send2fax_sent_items]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_user_attributes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_user_attributes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_user_log]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_user_log]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[mc_via]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[mc_via]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qb_attr]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qb_attr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qb_callback_dates]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qb_callback_dates]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qb_features]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qb_features]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qb_importance]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qb_importance]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qb_note]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qb_note]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qb_profiles]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qb_profiles]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qb_updatemgr]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qb_updatemgr]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_alert_history]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_alert_history]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_alert_settings]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_alert_settings]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_alerts]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_alerts]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_cust]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_cust]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_delete_check]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_delete_check]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_gvars]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_gvars]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_importance_levels]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_importance_levels]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_install_previous_locations]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_install_previous_locations]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_inv]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_inv]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_inv_lineitems]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_inv_lineitems]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_inv_payments]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_inv_payments]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_inv_payments_dtls]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_inv_payments_dtls]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_inv_temp_record]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_inv_temp_record]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_letter]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_letter]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_properties]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_properties]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_remarks]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_remarks]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_reps]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_reps]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_update_list]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_update_list]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qbx_update_user_history]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[qbx_update_user_history]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[reg_changes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[reg_changes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[reg_list]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[reg_list]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[srv_checkin]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[srv_checkin]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[startup_registries]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[startup_registries]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[user_drives]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[user_drives]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[user_process]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[user_process]
GO

CREATE TABLE [dbo].[LDAP_LIST] (
	[ldap_index] [int] IDENTITY (1, 1) NOT NULL ,
	[uid] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sAMAccountName] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[givenname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sn] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cn] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[distinguishedname] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[mail] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[legacyexchangedn] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[description] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[physicalDeliveryOfficeName] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[title] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[proxyaddress] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[targetaddress] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[c] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[company] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[department] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[homephone] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[L] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[location] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[postalcode] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[st] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[streetaddress] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[profilepath] [varchar] (400) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[scriptpath] [varchar] (400) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[homedirectory] [varchar] (400) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[objectCategory] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[mDBUseDefaults] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[msExchHomeServerName] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[homeMDB] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[userPrincipalName] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[userAccountControl] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[MailLog_items] (
	[item_index] [int] IDENTITY (1, 1) NOT NULL ,
	[item_datetime] [datetime] NULL ,
	[item_send_datetime] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_received_datetime] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_from] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_display_from] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_to] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_subject] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_body] [varchar] (3000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_attachment_name] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_error_log] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[file_index] [int] NULL ,
	[email_sent_to] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[email_error_log] [varchar] (400) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[a_notes] (
	[n_index] [int] IDENTITY (1, 1) NOT NULL ,
	[n_desc] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_company] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_contact] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_phone] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_fax] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_email] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_status] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_type] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_callback_hour] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_callback_minute] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_callback_area] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_notes] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_start_date] [datetime] NULL ,
	[n_end_date] [datetime] NULL ,
	[n_last_modified_date] [datetime] NULL ,
	[n_callback_date] [datetime] NULL ,
	[n_active] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[n_remove_date] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[backup_day_objects] (
	[day_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[object_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[object_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[object_last_result] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[object_last_datetime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[backup_days] (
	[day_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_action] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_session_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_session_description] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_template_path] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_template_filename] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_template_organization] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_template_group] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_template_cn] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_ex_path] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_ex_filename] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_ex_firstline] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_ex_mergeaction] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_ex_sourceservername] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_ex_datadirectoryname] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_ex_filecontaininglistofmailboxes] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_ex_logfilename] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_compressed] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_last_result] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[day_last_datetime] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[backup_emailaccounts] (
	[account_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[account_isuser] [int] NULL ,
	[account_available] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[backup_groups] (
	[group_index] [int] IDENTITY (1, 1) NOT NULL ,
	[group_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[group_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[backup_properties] (
	[property_auto_hour] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[property_auto_day] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[backup_sessions] (
	[session_id] [int] IDENTITY (1, 1) NOT NULL ,
	[session_name] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_description] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_seed] [float] NULL ,
	[session_start_datetime] [datetime] NULL ,
	[session_end_datetime] [datetime] NULL ,
	[session_results] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_report] [varchar] (5000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_userlist] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_total_mailbox_size] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_total_mail_count] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_stage] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[backup_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_update_msg] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_computer] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_zipfilename] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_zipfilepath] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_zipfilesize] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[session_filestatus] [int] NULL ,
	[session_deleted] [int] NULL ,
	[session_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[backup_users] (
	[user_index] [int] IDENTITY (1, 1) NOT NULL ,
	[user_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[user_datetime] [datetime] NULL ,
	[user_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[user_mailbox_size] [float] NULL ,
	[user_mail_count] [float] NULL ,
	[group_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[bad_process] (
	[bp_index] [int] IDENTITY (1, 1) NOT NULL ,
	[bp_name] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[bp_location] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[bp_active] [int] NULL ,
	[bp_desc] [varchar] (510) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[bp_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[bad_process_file_location] (
	[bp_index] [int] NULL ,
	[bpfl_file] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[bpfl_location] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[bpfl_desc] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[bad_process_files] (
	[bp_index] [int] NULL ,
	[bpf_file] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[bpf_location] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[bpf_desc] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[bad_process_registry_entries] (
	[bp_index] [int] NULL ,
	[bpr_value] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[bpr_key] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[bpr_type] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[bpr_root] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[bpr_location] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[bpr_desc] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[bad_process_websites] (
	[target_index] [int] IDENTITY (1, 1) NOT NULL ,
	[bp_index] [int] NULL ,
	[tw_name] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[tw_address] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[tw_desc] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[tw_datestamp] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[catch_this] (
	[catch_index] [int] IDENTITY (1, 1) NOT NULL ,
	[catch_string] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[catch_date] [datetime] NULL ,
	[catch_string_date] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[email_limit_messages] (
	[email_limits_index] [int] IDENTITY (1, 1) NOT NULL ,
	[email_limits_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[email_limits_mail_body] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[good_process] (
	[gp_index] [int] IDENTITY (1, 1) NOT NULL ,
	[gp_name] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[gp_location] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[gp_level] [int] NULL ,
	[gp_active] [bit] NULL ,
	[gp_desc] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[gp_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[internal_computers] (
	[computer_index] [int] IDENTITY (1, 1) NOT NULL ,
	[computer_name] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[computer_cpu] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[computer_memory_size] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[computer_current_user] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[computer_ip] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[computer_mac] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[computer_start_datestamp] [datetime] NULL ,
	[computer_end_datestamp] [datetime] NULL ,
	[computer_os] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[computer_os_build] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[computer_last_checkin] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[kill_user_process] (
	[kp_index] [int] IDENTITY (1, 1) NOT NULL ,
	[up_index] [int] NULL ,
	[kp_confirm_kill] [int] NULL ,
	[kp_datestamp] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[maillog_files] (
	[file_index] [int] IDENTITY (1, 1) NOT NULL ,
	[file_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[file_path] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[file_size] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[file_modified_datetime] [datetime] NULL ,
	[file_created_datetime] [datetime] NULL ,
	[file_processed_start_datetime] [datetime] NULL ,
	[file_processed_end_datetime] [datetime] NULL ,
	[file_stats] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[file_current_flag] [int] NULL ,
	[file_error_log] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_answer] (
	[answer_index] [int] IDENTITY (1, 1) NOT NULL ,
	[answer_location] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[answer_function] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[answer_output] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[answer_attributes] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[answer_string] [varchar] (3000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[answer_picture] [image] NULL ,
	[answer_by] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[request_by] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[answer_datetime] [datetime] NULL ,
	[answer_active] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_client_cache_firefox] (
	[cache_fox_index] [int] IDENTITY (1, 1) NOT NULL ,
	[cache_fox_entry] [varchar] (4000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cache_fox_type] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cache_fox_profile] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cache_fox_folder] [varchar] (512) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[request_id] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cache_fox_datetime] [datetime] NULL ,
	[cache_fox_user] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cache_fox_request_by] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cache_fox_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_client_cache_ie] (
	[cache_ie_index] [int] IDENTITY (1, 1) NOT NULL ,
	[cache_ie_entry] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cache_ie_type] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cache_ie_datetime] [datetime] NULL ,
	[cache_ie_user] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cache_ie_request_by] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cache_ie_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_client_events] (
	[event_index] [int] IDENTITY (1, 1) NOT NULL ,
	[event_category] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_category_string] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_computer_name] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_data] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_code] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_identifier] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_event_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_insertionstrings] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_logfile] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_message] [varchar] (512) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_record_number] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_source_name] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_time_generated] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_time_written] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_user] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_datetime] [datetime] NULL ,
	[event_system_user] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[event_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_client_keys] (
	[key_index] [int] IDENTITY (1, 1) NOT NULL ,
	[key_string] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[key_user] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[key_datetime_from] [datetime] NULL ,
	[key_datetime_to] [datetime] NULL ,
	[key_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_client_processes] (
	[process_index] [int] IDENTITY (1, 1) NOT NULL ,
	[process_id] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[process_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[process_cntthreads] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[process_cntusage] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[process_dwflags] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[process_dwsize] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[process_pcpriclassbase] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[process_th32defaultheapid] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[process_th32moduleid] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[process_th32parentprocessid] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[process_user] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[process_datetime] [datetime] NULL ,
	[process_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_client_services] (
	[service_index] [int] IDENTITY (1, 1) NOT NULL ,
	[service_name] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[service_display_name] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[service_description] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[service_status] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[service_startup_type] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[service_path] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[service_temp] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[service_datetime] [datetime] NULL ,
	[service_user] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[service_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_client_wmi_hardware] (
	[wmi_hardware_index] [int] IDENTITY (1, 1) NOT NULL ,
	[wmi_hardware_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[wmi_hardware_location] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[wmi_hardware_output] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[wmi_hardware_string] [varchar] (4000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[wmi_hardware_user] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[wmi_hardware_datetime] [datetime] NULL ,
	[wmi_hardware_requested_by] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[wmi_hardware_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_client_wmiclass] (
	[wmiclass_index] [int] IDENTITY (1, 1) NOT NULL ,
	[wmiclass_path] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[wmiclass_user] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[wmiclass_datetime] [datetime] NULL ,
	[wmiclass_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_functions] (
	[function_index] [int] IDENTITY (1, 1) NOT NULL ,
	[function_display_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[function_request_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[function_input] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[function_output] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[function_attributes] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[function_string] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[function_access_level] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[function_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_messages] (
	[msg_index] [int] IDENTITY (1, 1) NOT NULL ,
	[msg_process] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[msg_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[msg_datetime] [datetime] NULL ,
	[msg_msg] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[msg_by] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[msg_status] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[msg_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_requests] (
	[request_index] [int] IDENTITY (1, 1) NOT NULL ,
	[request_to_user] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[request_function] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[request_status] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[request_input] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[request_output] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[request_attributes] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[request_string] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[request_datetime] [datetime] NULL ,
	[request_by] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[request_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_send2fax_batch] (
	[batch_index] [int] IDENTITY (1, 1) NOT NULL ,
	[batch_datetime] [datetime] NULL ,
	[batch_finished] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[batch_item_count] [int] NULL ,
	[batch_status] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[retry_body] [varchar] (5000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[retry_subject] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[retry_from] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_send2fax_mailitems] (
	[item_index] [int] IDENTITY (1, 1) NOT NULL ,
	[item_number] [int] NULL ,
	[item_batch] [int] NULL ,
	[item_to] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_how_many_to] [int] NULL ,
	[item_subject] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_invoice] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_attachment_name] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_datetime] [datetime] NULL ,
	[item_sent] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_retries] [int] NULL ,
	[item_recieved] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[item_selected] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_send2fax_retries] (
	[sent_index] [int] NULL ,
	[batch_number] [int] NULL ,
	[item_number] [int] NULL ,
	[sent_to] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sent_to_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sent_sent] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sent_recieved] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sent_result] [varchar] (600) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sent_datetime] [datetime] NULL ,
	[retry_count] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_send2fax_sent_items] (
	[sent_index] [int] IDENTITY (1, 1) NOT NULL ,
	[batch_number] [int] NULL ,
	[item_number] [int] NULL ,
	[sent_to] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sent_to_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sent_sent] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sent_recieved] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sent_result] [varchar] (600) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sent_datetime] [datetime] NULL ,
	[retry_count] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_user_attributes] (
	[attr_index] [int] IDENTITY (1, 1) NOT NULL ,
	[attr_timer1_end] [int] NULL ,
	[attr_timer1_interval] [int] NULL ,
	[attr_show_tray_icon] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[attr_user] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[attr_datetime] [datetime] NULL ,
	[attr_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_user_log] (
	[log_index] [int] IDENTITY (1, 1) NOT NULL ,
	[log_user] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[log_computer] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[log_ip] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[log_version] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[log_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[log_status] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[log_datetime] [datetime] NULL ,
	[log_out_datetime] [datetime] NULL ,
	[log_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[mc_via] (
	[via_index] [int] IDENTITY (1, 1) NOT NULL ,
	[via_process_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[via_process_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[via_process_location] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[via_process_status] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[via_process_status_datetime] [datetime] NULL ,
	[via_for_who] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[via_for_computer] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[via_for_ip] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[via_from_who] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[via_from_datetime] [datetime] NULL ,
	[via_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qb_attr] (
	[attr_index] [int] IDENTITY (1, 1) NOT NULL ,
	[attr_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[attr_desc] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[attr_enabled] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qb_callback_dates] (
	[callback_index] [int] IDENTITY (1, 1) NOT NULL ,
	[note_index] [int] NULL ,
	[callback_date] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[callback_time] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[callback_created_by] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[callback_active] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qb_features] (
	[features_index] [int] NOT NULL ,
	[1] [int] NULL ,
	[2] [int] NULL ,
	[3] [int] NULL ,
	[4] [int] NULL ,
	[5] [int] NULL ,
	[6] [int] NULL ,
	[7] [int] NULL ,
	[8] [int] NULL ,
	[9] [int] NULL ,
	[10] [int] NULL ,
	[11] [int] NULL ,
	[12] [int] NULL ,
	[13] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[14] [int] NULL ,
	[15] [int] NULL ,
	[16] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[17] [int] NULL ,
	[18] [int] NULL ,
	[19] [int] NULL ,
	[20] [int] NULL ,
	[21] [int] NULL ,
	[22] [int] NULL ,
	[23] [int] NULL ,
	[24] [int] NULL ,
	[25] [int] NULL ,
	[26] [int] NULL ,
	[27] [int] NULL ,
	[28] [int] NULL ,
	[29] [int] NULL ,
	[30] [int] NULL ,
	[31] [int] NULL ,
	[32] [int] NULL ,
	[33] [int] NULL ,
	[34] [int] NULL ,
	[35] [int] NULL ,
	[36] [int] NULL ,
	[37] [int] NULL ,
	[38] [int] NULL ,
	[39] [int] NULL ,
	[40] [int] NULL ,
	[41] [int] NULL ,
	[42] [int] NULL ,
	[43] [int] NULL ,
	[44] [int] NULL ,
	[45] [int] NULL ,
	[46] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qb_importance] (
	[importance_index] [int] IDENTITY (1, 1) NOT NULL ,
	[importance_listid] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[importance_type] [int] NULL ,
	[importance_name] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[importance_created_by] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[importance_created_date] [datetime] NULL ,
	[importance_modified_by] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[importance_modified_date] [datetime] NULL ,
	[upfront] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qb_note] (
	[note_index] [int] IDENTITY (1, 1) NOT NULL ,
	[note_listid] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[note_datestamp] [datetime] NULL ,
	[note_msg] [varchar] (3000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[note_created_by] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[note_modified_date] [datetime] NULL ,
	[note_modified_by] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[note_msg_backup] [varchar] (3000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[note_callback_date] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[note_callback_time] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[note_company_name] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[note_company_amount] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[note_state] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[note_company_status] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[note_callback_today] [int] NULL ,
	[note_callback_today_date] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qb_profiles] (
	[profiles_index] [int] IDENTITY (1, 1) NOT NULL ,
	[profiles_username] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[profiles_full] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[profiles_fullname] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[profiles_email] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[profiles_phone] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[profiles_fax] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[profiles_enabled] [int] NOT NULL ,
	[profiles_level] [int] NULL ,
	[profiles_logins] [bigint] NULL ,
	[profiles_login] [datetime] NULL ,
	[profiles_logout] [datetime] NULL ,
	[profiles_login_progress] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[profiles_update_request] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[profiles_update_request_datestamp] [datetime] NULL ,
	[profiles_update_request_handled_datestamp] [datetime] NULL ,
	[profiles_update_result] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qb_updatemgr] (
	[updatemgr_index] [int] NULL ,
	[updatemgr_name] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[updatemgr_value] [int] NULL ,
	[updatemgr_status] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[updatemgr_timestamp] [datetime] NULL ,
	[updatemgr_last_timestamp] [datetime] NULL ,
	[updatemgr_last_success_timestamp] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_alert_history] (
	[alert_history_index] [int] IDENTITY (1, 1) NOT NULL ,
	[alert_id] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[alert_importance] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[alert_importance_type] [int] NULL ,
	[alert_total_balance] [money] NULL ,
	[alert_total_invoices] [int] NULL ,
	[alert_reason] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[alert_datetime] [datetime] NULL ,
	[cust_isactive] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_fullname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_alert_settings] (
	[alert_setting_max_dollar] [money] NULL ,
	[alert_setting_start_at_level] [int] NULL ,
	[alert_setting_max_invoices] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_alerts] (
	[alert_index] [int] IDENTITY (1, 1) NOT NULL ,
	[alert_id] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[alert_importance] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[alert_importance_type] [int] NULL ,
	[alert_total_balance] [money] NULL ,
	[alert_total_invoices] [int] NULL ,
	[alert_reason] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[alert_datetime] [datetime] NULL ,
	[cust_isactive] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_fullname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_cust] (
	[cust_unique] [int] IDENTITY (1, 1) NOT NULL ,
	[cust_bit] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_index] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_listid] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_name] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_fullname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_companyname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_contact] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_altcontact] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_salutation] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_firstname] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_middlename] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_lastname] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_billaddress_add1] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_billaddress_add2] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_billaddress_add3] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_billaddress_add4] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_billcity] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_billstate] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_billpostalcode] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_billcountry] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_shipaddress_add1] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_shipaddress_add2] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_shipaddress_add3] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_shipaddress_add4] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_shipcity] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_shipstate] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_shippostalcode] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_shipcountry] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_phone1] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_phone2] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_fax1] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_fax2] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_email1] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_email2] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_notes] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_isactive] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_sublevel] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_timecreated] [datetime] NULL ,
	[cust_timemodified] [datetime] NULL ,
	[cust_jobstatus] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_JobStartDate] [datetime] NULL ,
	[cust_webstatus] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_status] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_balance_numeric] [numeric](18, 0) NULL ,
	[cust_balance_money] [money] NULL ,
	[cust_balance] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_totalbalance_money] [money] NULL ,
	[cust_totalbalance] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_accountnumber] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_accountnumber_numeric] [int] NULL ,
	[cust_termsref_listid] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_termsref_fullname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_customertyperef_listid] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_customertyperef_residential] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_salestaxcoderef_listid] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_salestaxcoderef_fullname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_itemsalestaxref_listid] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_itemsalestaxref_fullname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_salesrepref_listid] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_salesrepref_fullname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[importance_type] [int] NULL ,
	[importance_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[importance_created_by] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[importance_created_date] [datetime] NULL ,
	[importance_modified_by] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[importance_modified_date] [datetime] NULL ,
	[importance_upfront] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cust_priority_alert] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_delete_check] (
	[delete_check_index] [int] NULL ,
	[delete_check_id] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[delete_check_good] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_gvars] (
	[gvars_index] [int] IDENTITY (1, 1) NOT NULL ,
	[gvars_name] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[gvars_value] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[gvars_type] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_importance_levels] (
	[import_index] [int] IDENTITY (1, 1) NOT NULL ,
	[import_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[import_id] [int] NULL ,
	[import_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_install_previous_locations] (
	[install_index] [int] IDENTITY (1, 1) NOT NULL ,
	[install_locations] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[install_set] [int] NULL ,
	[install_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_inv] (
	[inv_unique] [int] IDENTITY (1, 1) NOT NULL ,
	[inv_bit] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_txnid] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_timecreated] [datetime] NULL ,
	[inv_TimeModified] [datetime] NULL ,
	[inv_txnnumber] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_customerref_listid] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_customerref_fullname] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_classref_listid] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_classref_fullname] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_araccountref_listid] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_araccountref_fullname] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_txndate] [datetime] NULL ,
	[inv_refnumber] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_billaddress_add1] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_billaddress_add2] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_billaddress_add3] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_billaddress_add4] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_billcity] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_billstate] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_billpostalcode] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_billcountry] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_shipaddress_add1] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_shipaddress_add2] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_shipaddress_add3] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_shipaddress_add4] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_shipcity] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_shipstate] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_shippostalcode] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_shipcountry] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_ispending] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_isfinancecharge] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_termsref_listid] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_termsref_fullname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_duedate] [datetime] NULL ,
	[inv_salesrepref_listid] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_salesrepref_fullname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_shipdate] [datetime] NULL ,
	[inv_subtotal] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_salestaxpercentage] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_salestaxtotal] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_appliedamount] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_balanceremaining] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_customermsgref_listid] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_customermsgref_fullname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_istobeprinted] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_enabled] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_inv_lineitems] (
	[inv_line_unique] [int] IDENTITY (1, 1) NOT NULL ,
	[inv_line_bit] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_line_txnlineid] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_line_itemref_listid] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_line_itemref_fullname] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_line_desc] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_line_quantity] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_line_rate] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_line_amount] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_line_salestaxcoderef_listid] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_line_salestaxcoderef_fullname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_txnid_link] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_inv_payments] (
	[inv_pay_unique] [int] IDENTITY (1, 1) NOT NULL ,
	[inv_pay_bit] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_pay_txnid] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_pay_txntype] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_pay_txndate] [datetime] NULL ,
	[inv_pay_amount] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_pay_refnumber] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_pay_linktype] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_pay_type] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_txnid_link] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_inv_payments_dtls] (
	[inv_pay_dtls_unique] [int] IDENTITY (1, 1) NOT NULL ,
	[inv_pay_dtls_bit] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_pay_dtls_amount] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_pay_dtls_txnlineid] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_pay_dtls_type] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[inv_pay_txnid_link] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_inv_temp_record] (
	[qbx_inv_rs_index] [int] IDENTITY (1, 1) NOT NULL ,
	[qbx_inv_rs_txndate] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[qbx_inv_rs_ref_num] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[qbx_inv_rs_amount] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[qbx_inv_rs_color] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_letter] (
	[letter_index] [int] IDENTITY (1, 1) NOT NULL ,
	[letter_name] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_title] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_description] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_start] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_header] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_date_between] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_date] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_addresse_between] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_addresse] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_dear_between] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_dear] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_body_between] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_body] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_closing_between] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_closing] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_footer_between] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_footer] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_end] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_addon_after] [char] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_part_addon_message] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_timestamp] [timestamp] NULL ,
	[letter_createdby] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_active] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[letter_can_be_modified] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_properties] (
	[property_name] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[property_value] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_remarks] (
	[remark_index] [int] IDENTITY (1, 1) NOT NULL ,
	[remark_msg] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[remark_delete] [int] NULL ,
	[remark_created_by] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[remark_created_date] [datetime] NULL ,
	[remark_modified_by] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[remark_modified_date] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_reps] (
	[rep_unique] [int] IDENTITY (1, 1) NOT NULL ,
	[rep_listid] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[rep_initial] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[rep_isactive] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[rep_salesrepentityref_listid] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[rep_salesrepentityref_fullname] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[rep_salesrep_email] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_update_list] (
	[update_index] [int] IDENTITY (1, 1) NOT NULL ,
	[update_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[update_version] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[update_location] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[update_datetime] [datetime] NULL ,
	[update_active] [int] NULL ,
	[update_supp_speech] [int] NULL ,
	[update_supp_speech_ver] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[update_supp_agent] [int] NULL ,
	[update_supp_agent_ver] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[qbx_update_user_history] (
	[update_index] [int] IDENTITY (1, 1) NOT NULL ,
	[update_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[update_user] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[update_os_version] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[update_computer] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[update_app_version] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[update_lastupdated] [datetime] NULL ,
	[update_active] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[reg_changes] (
	[reg_changes_index] [int] IDENTITY (1, 1) NOT NULL ,
	[reg_changes_hkey] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_changes_path] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_changes_key] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_changes_type] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_changes_value] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_changes_datestamp] [datetime] NULL ,
	[reg_changes_notes] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_changes_section] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_changes_user] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_changes_computer] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_changes_active] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_changes_modifier] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[reg_list] (
	[reg_list_index] [int] IDENTITY (1, 1) NOT NULL ,
	[reg_list_hkey] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_list_path] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_list_key] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_list_type] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_list_value] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_list_datestamp] [datetime] NULL ,
	[reg_list_notes] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_list_section] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_list_user] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_list_computer] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[reg_list_active] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[srv_checkin] (
	[srv_index] [int] IDENTITY (1, 1) NOT NULL ,
	[srv_created] [datetime] NULL ,
	[srv_checkin] [datetime] NULL ,
	[srv_stopped] [datetime] NULL ,
	[srv_name] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[srv_mac] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[srv_ip] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[srv_harddrive] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[srv_memory] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[srv_active] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[startup_registries] (
	[sr_index] [int] IDENTITY (1, 1) NOT NULL ,
	[sr_root] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sr_location] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sr_key] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sr_value] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sr_datestamp] [datetime] NULL ,
	[sr_desc] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sr_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[sr_id] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[user_drives] (
	[drive_index] [int] IDENTITY (1, 1) NOT NULL ,
	[drive_letter] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_available] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_type] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_filesystem] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_freespace] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_isready] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_path] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_rootfolder] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_serialnumber] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_sharename] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_totalsize] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_volumename] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_subfolderscount] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_user] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[drive_datestamp] [datetime] NULL ,
	[drive_active] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[user_process] (
	[up_index] [int] IDENTITY (1, 1) NOT NULL ,
	[up_pid] [int] NULL ,
	[up_name] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[up_type] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[up_location] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[up_user] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[up_computer] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[up_datestamp] [datetime] NULL ,
	[up_executioner] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[up_execution_datestamp] [datetime] NULL ,
	[up_kill_process] [bit] NULL ,
	[up_kill_datestamp] [datetime] NULL ,
	[up_kill_result] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[up_kill_keep_trying] [bit] NULL ,
	[up_active] [bit] NULL ,
	[up_desc] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[up_desc_id] [int] NULL 
) ON [PRIMARY]
GO

