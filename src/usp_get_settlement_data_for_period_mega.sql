USE [postilion_office]
GO

/****** Object:  StoredProcedure [dbo].[usp_get_settlement_data_for_period]    Script Date: 09/27/2016 13:49:27 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO




CREATE PROCEDURE  [dbo].[usp_get_settlement_data_for_period] 
 @Start_Date DATETIME=NULL,    -- yyyymmdd
      @End_Date DATETIME=NULL     -- yyyymmdd
      AS 
    begin
DECLARE @norm_cutover BIT;

EXEC   dbo.check_norm_cutover_status  @is_cutover_successful= @norm_cutover OUTPUT;

 if( @norm_cutover = 1) BEGIN 
 
DECLARE @from_date DATETIME
DECLARE @to_date DATETIME
SET @from_date = ISNULL(@Start_Date,REPLACE(CONVERT(VARCHAR(10),  DATEADD(D, -1,GETDATE()),111),'/', ''))
SET @to_date = ISNULL(@End_Date,REPLACE(CONVERT(VARCHAR(10),  DATEADD(D, -1,GETDATE()),111),'/', ''))
 
 

	IF (OBJECT_ID('postilion_office.dbo.temp_journal_data') is not null )begin
	  DROP TABLE temp_journal_data
	end

	IF (OBJECT_ID('postilion_office.dbo.temp_post_tran_data') is not null )begin
	  DROP TABLE temp_post_tran_data
	end


	SELECT  

	J.adj_id
	,J.entry_id
	,J.config_set_id
	,J.session_id
	,J.post_tran_id
	,J.post_tran_cust_id
	,J.sdi_tran_id
	,J.acc_post_id
	,J.nt_fee_acc_post_id
	,J.coa_id
	,J.coa_se_id
	,J.se_id
	,J.amount
	,J.amount_id
	,J.amount_value_id
	,J.fee
	,J.fee_id
	,J.fee_value_id
	,J.nt_fee
	,J.nt_fee_id
	,J.nt_fee_value_id
	,J.debit_acc_nr_id
	,J.debit_acc_id
	,J.debit_cardholder_acc_id
	,J.debit_cardholder_acc_type
	,J.credit_acc_nr_id
	,J.credit_acc_id
	,J.credit_cardholder_acc_id
	,J.credit_cardholder_acc_type
	,J.business_date
	,J.granularity_element
	,J.tag
	,J.spay_session_id
	,J.spst_session_id
	,DebitAccNr.config_set_id DebitAccNr_config_set_id
	,DebitAccNr.acc_nr_id  DebitAccNr_acc_nr_id
	,DebitAccNr.se_id	DebitAccNr_se_id
	,DebitAccNr.acc_id	DebitAccNr_acc_id
	,DebitAccNr.acc_nr	DebitAccNr_acc_nr
	,DebitAccNr.aggregation_id DebitAccNr_aggregation_id
	,DebitAccNr.state	DebitAccNr_state
	,DebitAccNr.config_state DebitAccNr_config_state
	,CreditAccNr.config_set_id CreditAccNr_config_set_id
	,CreditAccNr.acc_nr_id  CreditAccNr_acc_nr_id
	,CreditAccNr.se_id	CreditAccNr_se_id
	,CreditAccNr.acc_id	CreditAccNr_acc_id
	,CreditAccNr.acc_nr	CreditAccNr_acc_nr
	,CreditAccNr.aggregation_id CreditAccNr_aggregation_id
	,CreditAccNr.state	CreditAccNr_state
	,CreditAccNr.config_state CreditAccNr_config_state
	,Amount.config_set_id	Amount_config_set_id
	,Amount.amount_id	Amount_amount_id
	,Amount.se_id	Amount_se_id
	,Amount.name	Amount_name
	,Amount.description	Amount_description
	,Amount.config_state	Amount_config_state
	,Fee.config_set_id Fee_config_set_id
	,Fee.fee_id	Fee_fee_id
	,Fee.se_id	Fee_se_id
	,fee.name	Fee_name
	,Fee.description Fee_description
	,Fee.type	Fee_type
	,Fee.amount_id Fee_amount_id
	,Fee.config_state Fee_config_state
	,coa.config_set_id coa_config_set_id
	,coa.coa_id	coa_coa_id
	,coa.name	coa_name
	,coa.description	coa_description
	,coa.type	coa_type
	,coa.config_state	coa_config_state
	,FeeValue.description FeeValue_description
	INTO temp_journal_data
	FROM
	dbo.sstl_journal_all AS J (NOLOCK)
	JOIN (SELECT  [DATE] recon_business_date FROM  dbo.[get_dates_in_range](@from_date, @to_date))rdate
	ON (rdate.recon_business_date =  J.business_date)
	LEFT OUTER JOIN dbo.sstl_se_acc_nr_w AS DebitAccNr (NOLOCK)
	ON (J.debit_acc_nr_id = DebitAccNr.acc_nr_id AND J.config_set_id = DebitAccNr.config_set_id)
	LEFT OUTER JOIN dbo.sstl_se_acc_nr_w AS CreditAccNr  (NOLOCK)
	ON (J.credit_acc_nr_id = CreditAccNr.acc_nr_id AND J.config_set_id = CreditAccNr.config_set_id)
	LEFT OUTER JOIN dbo.sstl_se_amount_w AS Amount  (NOLOCK)
	ON (J.amount_id = Amount.amount_id AND J.config_set_id = Amount.config_set_id)
	LEFT OUTER JOIN dbo.sstl_se_fee_w AS Fee (NOLOCK)
	ON (J.fee_id = Fee.fee_id AND J.config_set_id = Fee.config_set_id)
	LEFT OUTER JOIN dbo.sstl_se_fee_value_w AS FeeValue (NOLOCK)
ON (J.fee_id = FeeValue.fee_id AND J.config_set_id = FeeValue.config_set_id)
	LEFT OUTER JOIN dbo.sstl_coa_w AS Coa  (NOLOCK)
	ON (J.coa_id = Coa.coa_id AND J.config_set_id = Coa.config_set_id)
	OPTION (recompile,maxdop 8)


	SELECT
	PT.[post_tran_id]	PT_post_tran_id
	,PT.[post_tran_cust_id]	PT_post_tran_cust_id
	,PT.[settle_entity_id]	PT_settle_entity_id
	,PT.[batch_nr]	PT_batch_nr
	,PT.[prev_post_tran_id]	PT_prev_post_tran_id
	,PT.[next_post_tran_id]	PT_next_post_tran_id
	,PT.[sink_node_name]	PT_sink_node_name
	,PT.[tran_postilion_originated]	PT_tran_postilion_originated
	,PT.[tran_completed]	PT_tran_completed
	,PT.[message_type]	PT_message_type
	,PT.[tran_type]	PT_tran_type
	,PT.[tran_nr]	PT_tran_nr
	,PT.[system_trace_audit_nr]	PT_system_trace_audit_nr
	,PT.[rsp_code_req]	PT_rsp_code_req
	,PT.[rsp_code_rsp]	PT_rsp_code_rsp
	,PT.[abort_rsp_code]	PT_abort_rsp_code
	,PT.[auth_id_rsp]	PT_auth_id_rsp
	,PT.[auth_type]	PT_auth_type
	,PT.[auth_reason]	PT_auth_reason
	,PT.[retention_data]	PT_retention_data
	,PT.[acquiring_inst_id_code]	PT_acquiring_inst_id_code
	,PT.[message_reason_code]	PT_message_reason_code
	,PT.[sponsor_bank]	PT_sponsor_bank
	,PT.[retrieval_reference_nr]	PT_retrieval_reference_nr
	,PT.[datetime_tran_gmt]	PT_datetime_tran_gmt
	,PT.[datetime_tran_local]	PT_datetime_tran_local
	,PT.[datetime_req]	PT_datetime_req
	,PT.[datetime_rsp]	PT_datetime_rsp
	,PT.[realtime_business_date]	PT_realtime_business_date
	,PT.[recon_business_date]	PT_recon_business_date
	,PT.[from_account_type]	PT_from_account_type
	,PT.[to_account_type]	PT_to_account_type
	,PT.[from_account_id]	PT_from_account_id
	,PT.[to_account_id]	PT_to_account_id
	,PT.[tran_amount_req]	PT_tran_amount_req
	,PT.[tran_amount_rsp]	PT_tran_amount_rsp
	,PT.[settle_amount_impact]	PT_settle_amount_impact
	,PT.[tran_cash_req]	PT_tran_cash_req
	,PT.[tran_cash_rsp]	PT_tran_cash_rsp
	,PT.[tran_currency_code]	PT_tran_currency_code
	,PT.[tran_tran_fee_req]	PT_tran_tran_fee_req
	,PT.[tran_tran_fee_rsp]	PT_tran_tran_fee_rsp
	,PT.[tran_tran_fee_currency_code]	PT_tran_tran_fee_currency_code
	,PT.[tran_proc_fee_req]	PT_tran_proc_fee_req
	,PT.[tran_proc_fee_rsp]	PT_tran_proc_fee_rsp
	,PT.[tran_proc_fee_currency_code]	PT_tran_proc_fee_currency_code
	,PT.[settle_amount_req]	PT_settle_amount_req
	,PT.[settle_amount_rsp]	PT_settle_amount_rsp
	,PT.[settle_cash_req]	PT_settle_cash_req
	,PT.[settle_cash_rsp]	PT_settle_cash_rsp
	,PT.[settle_tran_fee_req]	PT_settle_tran_fee_req
	,PT.[settle_tran_fee_rsp]	PT_settle_tran_fee_rsp
	,PT.[settle_proc_fee_req]	PT_settle_proc_fee_req
	,PT.[settle_proc_fee_rsp]	PT_settle_proc_fee_rsp
	,PT.[settle_currency_code]	PT_settle_currency_code
	,PT.[pos_entry_mode]	PT_pos_entry_mode
	,PT.[pos_condition_code]	PT_pos_condition_code
	,PT.[additional_rsp_data]	PT_additional_rsp_data
	,PT.[tran_reversed]	PT_tran_reversed
	,PT.[prev_tran_approved]	PT_prev_tran_approved
	,PT.[issuer_network_id]	PT_issuer_network_id
	,PT.[acquirer_network_id]	PT_acquirer_network_id
	,PT.[extended_tran_type]	PT_extended_tran_type
	,PT.[from_account_type_qualifier]	PT_from_account_type_qualifier
	,PT.[to_account_type_qualifier]	PT_to_account_type_qualifier
	,PT.[bank_details]	PT_bank_details
	,PT.[payee]	PT_payee
	,PT.[card_verification_result]	PT_card_verification_result
	,PT.[online_system_id]	PT_online_system_id
	,PT.[participant_id]	PT_participant_id
	,PT.[opp_participant_id]	PT_opp_participant_id
	,PT.[receiving_inst_id_code]	PT_receiving_inst_id_code
	,PT.[routing_type]	PT_routing_type
	,PT.[pt_pos_operating_environment]	PT_pt_pos_operating_environment
	,PT.[pt_pos_card_input_mode]	PT_pt_pos_card_input_mode
	,PT.[pt_pos_cardholder_auth_method]	PT_pt_pos_cardholder_auth_method
	,PT.[pt_pos_pin_capture_ability]	PT_pt_pos_pin_capture_ability
	,PT.[pt_pos_terminal_operator]	PT_pt_pos_terminal_operator
	,PT.[source_node_key]	PT_source_node_key
	,PT.[proc_online_system_id]	PT_proc_online_system_id
	,PTC.[post_tran_cust_id]	PTC_post_tran_cust_id
	,PTC.[source_node_name]	PTC_source_node_name
	,PTC.[draft_capture]	PTC_draft_capture
	,PTC.[pan]	PTC_pan
	,PTC.[card_seq_nr]	PTC_card_seq_nr
	,PTC.[expiry_date]	PTC_expiry_date
	,PTC.[service_restriction_code]	PTC_service_restriction_code
	,PTC.[terminal_id]	PTC_terminal_id
	,PTC.[terminal_owner]	PTC_terminal_owner
	,PTC.[card_acceptor_id_code]	PTC_card_acceptor_id_code
	,PTC.[mapped_card_acceptor_id_code]	PTC_mapped_card_acceptor_id_code
	,PTC.[merchant_type]	PTC_merchant_type
	,PTC.[card_acceptor_name_loc]	PTC_card_acceptor_name_loc
	,PTC.[address_verification_data]	PTC_address_verification_data
	,PTC.[address_verification_result]	PTC_address_verification_result
	,PTC.[check_data]	PTC_check_data
	,PTC.[totals_group]	PTC_totals_group
	,PTC.[card_product]	PTC_card_product
	,PTC.[pos_card_data_input_ability]	PTC_pos_card_data_input_ability
	,PTC.[pos_cardholder_auth_ability]	PTC_pos_cardholder_auth_ability
	,PTC.[pos_card_capture_ability]	PTC_pos_card_capture_ability
	,PTC.[pos_operating_environment]	PTC_pos_operating_environment
	,PTC.[pos_cardholder_present]	PTC_pos_cardholder_present
	,PTC.[pos_card_present]	PTC_pos_card_present
	,PTC.[pos_card_data_input_mode]	PTC_pos_card_data_input_mode
	,PTC.[pos_cardholder_auth_method]	PTC_pos_cardholder_auth_method
	,PTC.[pos_cardholder_auth_entity]	PTC_pos_cardholder_auth_entity
	,PTC.[pos_card_data_output_ability]	PTC_pos_card_data_output_ability
	,PTC.[pos_terminal_output_ability]	PTC_pos_terminal_output_ability
	,PTC.[pos_pin_capture_ability]	PTC_pos_pin_capture_ability
	,PTC.[pos_terminal_operator]	PTC_pos_terminal_operator
	,PTC.[pos_terminal_type]	PTC_pos_terminal_type
	,PTC.[pan_search]	PTC_pan_search
	,PTC.[pan_encrypted]	PTC_pan_encrypted
	,PTC.[pan_reference]	PTC_pan_reference
	 INTO temp_post_tran_data
	FROM   post_tran AS PT  WITH (NOLOCK, INDEX(ix_post_tran_9))
	JOIN (SELECT  [DATE] recon_business_date FROM  dbo.[get_dates_in_range](@from_date, @to_date))rdate
	ON (rdate.recon_business_date = PT.recon_business_date) AND  PT.tran_postilion_originated = 0
	AND PT.post_tran_id NOT IN (
						SELECT  post_tran_id  FROM tbl_late_reversals ll (NOLOCK) 
			WHERE ll.recon_business_date >= @from_date
			and (datepart(D,ll.rev_datetime_req) - datepart(D, ll.trans_datetime_req ))>1 
						) AND
						rsp_code_rsp IN ('00','11','09')
	 LEFT JOIN  Post_tran_cust AS PTC WITH  (NOLOCK, INDEX(pk_post_tran_cust))
	 on PT.post_tran_cust_id = PTC.post_tran_cust_id
	OPTION (RECOMPILE,maxdop 8)

	CREATE NONCLUSTERED INDEX ix_temp_post_tran_data_1 ON temp_post_tran_data (PT_message_type )
	CREATE NONCLUSTERED INDEX ix_temp_post_tran_data_2 ON temp_post_tran_data (PT_tran_type);
	CREATE NONCLUSTERED INDEX ix_temp_post_tran_data_3 ON temp_post_tran_data (PTC_source_node_name);
	CREATE NONCLUSTERED INDEX ix_temp_post_tran_data_4 ON temp_post_tran_data (PT_sink_node_name);
	CREATE NONCLUSTERED INDEX ix_temp_post_tran_data_5 ON temp_post_tran_data (PTC_terminal_id);
	CREATE NONCLUSTERED INDEX ix_temp_post_tran_data_6 ON temp_post_tran_data (PTC_totals_group);
	CREATE NONCLUSTERED INDEX ix_temp_post_tran_data_7 ON temp_post_tran_data (PTC_pan);
	CREATE NONCLUSTERED INDEX ix_temp_post_tran_data_9 ON temp_post_tran_data (PT_acquiring_inst_id_code);
	CREATE NONCLUSTERED INDEX ix_temp_post_tran_data_10 ON temp_post_tran_data (PT_settle_amount_impact);

	CREATE NONCLUSTERED INDEX ix_temp_post_tran_data_11 ON temp_post_tran_data (

	PT_message_type,
	PT_tran_type,
	PTC_source_node_name,
	PT_sink_node_name,
	PTC_terminal_id,
	PTC_totals_group,
	PTC_pan,
	PT_settle_amount_impact,
	PT_acquiring_inst_id_code
	);

	CREATE INDEX [ix_temp_post_tran_data_12]    ON [dbo].[temp_post_tran_data]    ([PT_tran_postilion_originated])      INCLUDE ([PT_post_tran_cust_id], [PT_tran_nr], [PT_retention_data])      WITH (FILLFACTOR=90)
	CREATE INDEX [ix_temp_post_tran_data_13]    ON [dbo].[temp_post_tran_data]    ([PT_post_tran_cust_id], [PT_tran_postilion_originated], [PT_tran_nr])      INCLUDE ([PT_retention_data])      WITH (FILLFACTOR=90)
	CREATE INDEX [ix_temp_post_tran_data_14]    ON [dbo]    .[temp_post_tran_data]([PT_tran_postilion_originated], [PTC_totals_group])      INCLUDE ([PT_post_tran_id], [PT_post_tran_cust_id], [PT_sink_node_name], [PT_message_type], [PT_tran_type], [PT_tran_nr], [PT_system_trace_audit_nr], [PT_acquiring_inst_id_code], [PT_retrieval_reference_nr], [PT_settle_amount_impact], [PT_settle_amount_rsp], [PT_settle_currency_code], [PT_tran_reversed], [PT_extended_tran_type], [PT_payee], [PTC_source_node_name], [PTC_pan], [PTC_terminal_id], [PTC_card_acceptor_id_code], [PTC_merchant_type], [PTC_card_acceptor_name_loc])      WITH (FILLFACTOR=90)
	CREATE NONCLUSTERED INDEX [ix_temp_post_tran_data_15] ON [dbo].[temp_post_tran_data] ([PT_recon_business_date])

 CREATE INDEX [ix_temp_post_tran_data_16]    ON [dbo]    .[temp_post_tran_data]    ([PT_tran_postilion_originated], [PT_message_type], [PT_rsp_code_rsp], [PT_settle_amount_impact])      INCLUDE ([PT_sink_node_name], [PT_system_trace_audit_nr], [PT_retrieval_reference_nr], [PT_to_account_id], [PT_payee], [PTC_source_node_name], [PTC_pan], [PTC_terminal_id])      WITH (FILLFACTOR=70, ONLINE=ON)
	CREATE NONCLUSTERED INDEX [temp_post_tran_data_recon] ON [dbo].[temp_post_tran_data] 
(
	[PT_tran_postilion_originated] ASC,
	[PT_settle_currency_code] ASC,
	[PT_sink_node_name] ASC,
	[PT_tran_type] ASC,
	[PT_rsp_code_rsp] ASC,
	[PTC_source_node_name] ASC,
	[PTC_card_acceptor_id_code] ASC,
	[PTC_totals_group] ASC
)
	INCLUDE (
	
	 [PT_post_tran_id],
	[PT_post_tran_cust_id],
	[PT_message_type],
	[PT_system_trace_audit_nr],
	[PT_acquiring_inst_id_code],
	[PT_retrieval_reference_nr],
	[PT_settle_amount_impact],
	[PT_settle_amount_rsp],
	[PT_tran_reversed],
	[PT_extended_tran_type],
	[PT_payee],
	[PTC_pan],
	[PTC_terminal_id],
	[PTC_terminal_owner],
	[PTC_merchant_type]) WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
	CREATE NONCLUSTERED INDEX ix_temp_post_tran_data_17 ON temp_post_tran_data (PT_extended_tran_type);
	CREATE NONCLUSTERED INDEX ix_temp_post_tran_data_18 ON temp_post_tran_data (PT_retention_data);

	CREATE NONCLUSTERED INDEX ix_temp_journal_data_1 ON temp_journal_data (DebitAccNr_acc_nr) INCLUDE (CreditAccNr_acc_nr);
	CREATE NONCLUSTERED INDEX ix_temp_journal_data_2 ON temp_journal_data (CreditAccNr_acc_nr)INCLUDE (DebitAccNr_acc_nr);
	CREATE NONCLUSTERED INDEX ix_temp_journal_data_3 ON temp_journal_data (business_date);
	CREATE NONCLUSTERED INDEX ix_temp_journal_data_4 ON temp_journal_data (fee_name);
	CREATE NONCLUSTERED INDEX ix_temp_journal_data_5 ON temp_journal_data (feevalue_description);
    CREATE NONCLUSTERED INDEX ix_temp_journal_data_6 ON temp_journal_data (post_tran_id);
    CREATE NONCLUSTERED INDEX ix_temp_journal_data_7 ON temp_journal_data (post_tran_cust_id);
END

END

GO


