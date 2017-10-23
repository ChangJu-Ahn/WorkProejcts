<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%	
Call LoadBasisGlobalInf()

	Dim lgOpModeCRUD
	Err.Clear 

	Call HideStatusWnd

	lgOpModeCRUD	=	Request("txtMode")

	Select Case lgOpModeCRUD
	        Case CStr(UID_M0002)
	             Call SubBizSave()
	        Case CStr(UID_M0003)                                                         '☜: Delete
	             Call SubBizDelete()
	End Select

'============================================================================================================
' Name : SubBizDelete
'============================================================================================================
Sub SubBizDelete()
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear  
    
    DIM Command 
    DIM I1_b_biz_partner_bp_cd 
    DIM I2_b_biz_partner_bp_cd 
    DIM I3_m_lc_amend_hdr 
    DIM I4_s_wks_user 
    DIM I5_b_pur_grp
    dim I6_m_lc_amend_hdr 
    dim strConvDt
  
    CONST M435_I3_lc_amd_no = 0
    Const M435_I3_lc_no = 1
    Const M435_I3_lc_doc_no = 2
    Const M435_I3_lc_amend_seq = 3
 
    DIM PM4G211
  
    
	If Request("txtLCAmdNo") = "" Then										'⊙: 삭제를 위한 값이 들어왔는지 체크 
			Call DisplayMsgBox("229909", vbOKOnly, "", "", I_MKSCRIPT)           
			Exit Sub 
	End If
	
    ReDim I3_m_lc_amend_hdr(1)
    I3_m_lc_amend_hdr(M435_I3_lc_amd_no)  =    Request("txtLCAmdNo")
    
	   Set PM4G211 = Server.CreateObject("PM4G211.cMMaintLcAmendHdrS")
		 
		Command = "Delete"
		   
		If CheckSYSTEMError(Err,True) = True Then
			Set PM4G211 = Nothing
			Exit Sub
		End If
		
	   CALL PM4G211.M_MAINT_LC_AMEND_HDR_SVR(gStrGlobalCollection , cstr(Command), cstr(I1_b_biz_partner_bp_cd),cstr(I2_b_biz_partner_bp_cd),I3_m_lc_amend_hdr,I4_s_wks_user,I5_b_pur_grp,I6_m_lc_amend_hdr) 
		    
		if CheckSYSTEMError2(Err,True,"","","","","") = true then 		
			Set PM4G211 = Nothing												'☜: ComProxy Unload
			Exit Sub														'☜: 비지니스 로직 처리를 종료함 
 		end if

		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "Parent.DBDeleteOK "           & vbCr
		Response.Write "</Script>"                  & vbCr 
                                    '☜: Clear Error status
                                    
End Sub
'============================================================================================================
' Name : SubBizSave
'============================================================================================================
Sub SubBizSave()
														'☜: 현재 Save 요청을 받음 
    DIM Command 
    DIM I1_b_biz_partner_bp_cd 
    DIM I2_b_biz_partner_bp_cd 
    DIM I3_m_lc_amend_hdr 
    DIM I4_s_wks_user 
    DIM I5_b_pur_grp 
    dim strConvDt
    DIM PM4G211
    
     Const M435_I3_lc_amd_no = 0
     Const M435_I3_lc_no = 1
     Const M435_I3_lc_doc_no = 2
     Const M435_I3_lc_amend_seq = 3
     Const M435_I3_adv_no = 4
     Const M435_I3_pre_adv_ref = 5
     Const M435_I3_open_dt = 6
     Const M435_I3_be_expiry_dt = 7
     Const M435_I3_at_expiry_dt = 8
     Const M435_I3_manufacturer = 9
     Const M435_I3_agent = 10
     Const M435_I3_amend_dt = 11
     Const M435_I3_amend_req_dt = 12
     Const M435_I3_currency = 13
     Const M435_I3_be_doc_amt = 14
     Const M435_I3_at_doc_amt = 15
     Const M435_I3_at_xch_rate = 16
     Const M435_I3_inc_amt = 17
     Const M435_I3_dec_amt = 18
     Const M435_I3_be_loc_amt = 19
     Const M435_I3_at_loc_amt = 20
     Const M435_I3_be_partial_ship = 21
     Const M435_I3_at_partial_ship = 22
     Const M435_I3_be_latest_ship_dt = 23
     Const M435_I3_at_latest_ship_dt = 24
     Const M435_I3_open_bank = 25
     Const M435_I3_be_xch_rate = 26
     Const M435_I3_ext1_amt = 27
     Const M435_I3_ext1_cd = 28
     Const M435_I3_remark = 29
     Const M435_I3_lc_kind = 30
     Const M435_I3_remark2 = 31
     Const M435_I3_be_transhipment = 32
     Const M435_I3_at_transhipment = 33
     Const M435_I3_be_transfer = 34
     Const M435_I3_at_transfer = 35
     Const M435_I3_be_loading_port = 36
     Const M435_I3_at_loading_port = 37
     Const M435_I3_be_dischge_port = 38
     Const M435_I3_at_dischge_port = 39
     Const M435_I3_be_transport = 40
     Const M435_I3_at_transport = 41
     Const M435_I3_biz_area = 42
     Const M435_I3_charge_flg = 43
     Const M435_I3_adv_bank = 44
     Const M435_I3_ext1_qty = 45
     Const M435_I3_ext2_qty = 46
     Const M435_I3_ext3_qty = 47
     Const M435_I3_ext2_amt = 48
     Const M435_I3_ext3_amt = 49
     Const M435_I3_ext2_cd = 50
     Const M435_I3_ext3_cd = 51
     Const M435_I3_ext1_rt = 52
     Const M435_I3_ext2_rt = 53
     Const M435_I3_ext3_rt = 54
     Const M435_I3_ext1_dt = 55
     Const M435_I3_ext2_dt = 56
     Const M435_I3_ext3_dt = 57
     
    Dim  lgIntFlgMode
   
   On Error Resume Next                                                            '☜: Protect system from crashing
	Err.Clear   

    lgIntFlgMode = CInt(Request("txtFlgMode"))	
    
    	Set PM4G211 = Server.CreateObject("PM4G211.cMMaintLcAmendHdrS")

        If CheckSYSTEMError(Err,True) = True Then
			Set PM4G211 = Nothing
			Exit Sub
		End If
		
       redim I3_m_lc_amend_hdr(60)
        
        '-----------------------
		I3_m_lc_amend_hdr(M435_I3_lc_amd_no)  = Request("txtLCAmdNo1")
		I3_m_lc_amend_hdr(M435_I3_lc_doc_no)  = Request("txtLCDocNo")

		I3_m_lc_amend_hdr(M435_I3_lc_no ) = Request("txtLCNo")
		I3_m_lc_amend_hdr(M435_I3_remark ) = Trim(Request("txtRemark"))
		I3_m_lc_amend_hdr(M435_I3_currency ) = UCase(Trim(Request("txtCurrency")))
		I3_m_lc_amend_hdr(M435_I3_at_xch_rate ) = UNIConvNum(Trim(Request("txtXchRt")),0)
		'If Len(Trim(Request("txtBeDocAmt"))) Then
		I3_m_lc_amend_hdr(M435_I3_at_doc_amt ) = UNIConvNum(Request("txtAtDocAmt"),0)
		I3_m_lc_amend_hdr(M435_I3_be_doc_amt ) = UNIConvNum(Request("txtBeDocAmt"),0)
		'End If
		
		If Len(Trim(Request("txtOpenDt"))) Then
			strConvDt = UNIConvDate(Request("txtOpenDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtOpenDt", 1, I_MKSCRIPT)

				Set M32211 = Nothing											'☜: ComProxy UnLoad

				Exit Sub													'☜: Process End
			Else
				I3_m_lc_amend_hdr(M435_I3_open_dt ) = strConvDt
			End If
		End If

		I3_m_lc_amend_hdr(M435_I3_adv_bank ) = UCase(Trim(Request("txtAdvBank")))
		I3_m_lc_amend_hdr(M435_I3_open_bank ) = UCase(Trim(Request("txtOpenBank")))
		I3_m_lc_amend_hdr(M435_I3_adv_no ) = UCase(Trim(Request("txtAdvNo")))
		I1_b_biz_partner_bp_cd = UCase(Trim(Request("txtBeneficiary")))
		I3_m_lc_amend_hdr(M435_I3_lc_amd_no ) = UCase(Trim(Request("txtLCAmdNo1")))
		I3_m_lc_amend_hdr(M435_I3_lc_doc_no ) = UCase(Trim(Request("txtLCDocNo")))
		

		If Len(Trim(Request("txtLCAmendSeq"))) Then
			I3_m_lc_amend_hdr(M435_I3_lc_amend_seq ) = UNIConvNum(Request("txtLCAmendSeq"),0)
		End If

		I2_b_biz_partner_bp_cd = UCase(Trim(Request("txtApplicant")))

		If Len(Trim(Request("txtAmendDt"))) Then
			strConvDt = UNIConvDate(Request("txtAmendDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtAmendDt", 1, I_MKSCRIPT)

				Set M32211 = Nothing											'☜: ComProxy UnLoad

				Exit Sub													'☜: Process End
			Else
				I3_m_lc_amend_hdr(M435_I3_amend_dt ) = strConvDt
			End If
		End If

		If Len(Trim(Request("txtAmendReqDt"))) Then
			strConvDt = UNIConvDate(Request("txtAmendReqDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtAmendReqDt", 1, I_MKSCRIPT)

				Set M32211 = Nothing											'☜: ComProxy UnLoad

				Exit Sub													'☜: Process End
			Else
				I3_m_lc_amend_hdr(M435_I3_amend_req_dt ) = strConvDt
			End If	
		End If

		If Request("rdoAtDocAmt") = "I" Then
			If Len(Trim(Request("txtAmendAmt"))) Then
				I3_m_lc_amend_hdr(M435_I3_inc_amt ) = UNIConvNum(Request("txtAmendAmt"),0)
				I3_m_lc_amend_hdr(M435_I3_at_doc_amt ) = UNIConvNum(Request("txtAtDocAmt"),0)
				'response.write "rdoAtDoc:I    " & M32211.ImportMLcAmendHdrAtDocAmt & "=="
			End If
		ElseIf Request("rdoAtDocAmt") = "D" Then
			If Len(Trim(Request("txtAmendAmt"))) Then
				I3_m_lc_amend_hdr(M435_I3_dec_amt ) = UNIConvNum(Request("txtAmendAmt"),0)
				I3_m_lc_amend_hdr(M435_I3_at_doc_amt ) = UNIConvNum(Request("txtAtDocAmt"),0)
				'response.write "rdoAtDoc:D    " & M32211.ImportMLcAmendHdrAtDocAmt & "=="
			End If
		End If

		If Len(Trim(Request("txtAtExpiryDt"))) Then
			strConvDt = UNIConvDate(Request("txtAtExpiryDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtAtExpiryDt", 1, I_MKSCRIPT)

				Set M32211 = Nothing											'☜: ComProxy UnLoad

				Exit Sub													'☜: Process End
			Else
				I3_m_lc_amend_hdr(M435_I3_at_expiry_dt) = strConvDt
			End If
		End If

		If Len(Trim(Request("txtBeExpiryDt"))) Then
			strConvDt = UNIConvDate(Request("txtBeExpiryDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtBeExpiryDt", 1, I_MKSCRIPT)

				Set M32211 = Nothing											'☜: ComProxy UnLoad

				Exit Sub													'☜: Process End
			Else
				I3_m_lc_amend_hdr(M435_I3_be_expiry_dt) = strConvDt
			End If
		End If

		If Len(Trim(Request("txtAtLatestShipDt"))) Then
			strConvDt = UNIConvDate(Request("txtAtLatestShipDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtAtLatestShipDt", 1, I_MKSCRIPT)

				Set M32211 = Nothing											'☜: ComProxy UnLoad

				Exit Sub													'☜: Process End
			Else
				I3_m_lc_amend_hdr(M435_I3_at_latest_ship_dt ) = strConvDt
			End If
		End If

		If Len(Trim(Request("txtBeLatestShipDt"))) Then
			strConvDt = UNIConvDate(Request("txtBeLatestShipDt"))

			If strConvDt = "" Then
				Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
				Call LoadTab("parent.frm1.txtBeLatestShipDt", 1, I_MKSCRIPT)

				Set M32211 = Nothing											'☜: ComProxy UnLoad

				Exit Sub													'☜: Process End
			Else
				I3_m_lc_amend_hdr(M435_I3_be_latest_ship_dt ) = strConvDt
			End If
		End If
		
		I3_m_lc_amend_hdr(M435_I3_at_transhipment ) 	= Trim(Request("rdoAtTranshipment"))
		I3_m_lc_amend_hdr(M435_I3_be_transhipment ) 	= Trim(Request("txtBeTranshipment"))
		I3_m_lc_amend_hdr(M435_I3_at_partial_ship ) 	= Trim(Request("rdoAtPartialShip"))
		I3_m_lc_amend_hdr(M435_I3_be_partial_ship ) 	= Trim(Request("txtBePartialShip"))
		I3_m_lc_amend_hdr(M435_I3_at_transfer ) 		= Trim(Request("rdoAtTransfer"))
		I3_m_lc_amend_hdr(M435_I3_be_transfer ) 		= Trim(Request("txtBeTransfer"))
		I3_m_lc_amend_hdr(M435_I3_at_transport ) 	=  UCase(Trim(Request("txtAtTransport")))
		I3_m_lc_amend_hdr(M435_I3_be_transport ) 	=  UCase(Trim(Request("txtBeTransport")))
		I3_m_lc_amend_hdr(M435_I3_at_loading_port ) 	= Trim(Request("txtAtLoadingPort"))
		I3_m_lc_amend_hdr(M435_I3_be_loading_port ) 	= Trim(Request("txtBeLoadingPort"))
		I3_m_lc_amend_hdr(M435_I3_at_dischge_port ) 	= Trim(Request("txtAtDischgePort"))
		I3_m_lc_amend_hdr(M435_I3_be_dischge_port ) 	= Trim(Request("txtBeDischgePort"))
		I3_m_lc_amend_hdr(M435_I3_lc_kind  )	= "M"
		
		
		If lgIntFlgMode = OPMD_CMODE Then
			Command = "Create"
		ElseIf lgIntFlgMode = OPMD_UMODE Then
			Command = "Update"
		End If
	
		Dim I6_m_lc_amend_hdr

	    CALL PM4G211.M_MAINT_LC_AMEND_HDR_SVR(gStrGlobalCollection , cstr(Command), cstr(I1_b_biz_partner_bp_cd),cstr(I2_b_biz_partner_bp_cd),I3_m_lc_amend_hdr,I4_s_wks_user,I5_b_pur_grp,I6_m_lc_amend_hdr) 
		

	    if CheckSYSTEMError2(Err,True,"","","","","") = true then 		
			Set PM4G211 = Nothing												'☜: ComProxy Unload
			Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	    end if

	                 
	    Response.Write "<Script Language=vbscript>" & vbCr
	    Response.Write "If """ & ConvSPChars(Trim(I6_m_lc_amend_hdr)) & """ <> """"  Then " & vbCr   
		Response.Write " parent.frm1.txtLCAmdNo.value = """ & ConvSPChars(Trim(I6_m_lc_amend_hdr)) & """" & vbCr
		Response.Write " parent.frm1.txtLCAmdNo1.value = """ & ConvSPChars(Trim(I6_m_lc_amend_hdr)) & """" & vbCr
		Response.Write "End If" & vbCr
		Response.Write " Parent.DBSaveOK "           & vbCr
	    Response.Write "</Script>"                  & vbCr 

End Sub

%>