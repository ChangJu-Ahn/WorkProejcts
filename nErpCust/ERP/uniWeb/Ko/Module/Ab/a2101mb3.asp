<%@ LANGUAGE=VBSCript%>
<%Option Explicit%> 
<!-- #Include file="../../inc/incSvrMain.asp"  -->

<%
Call LoadBasisGlobalInf()

Call HideStatusWnd		

On Error Resume Next

Dim iPD1G041																'☆ : 조회용 ComPlus Dll 사용 변수 

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strCmd
Dim strKey
Dim LngRow
Dim GroupCount
Dim strData
Dim LngMaxRow
Dim NegaValue
Dim PosiValue

Dim Exp_Acct_Gp()
Const E1_gp_cd = 0
Const E1_gp_nm = 1
Const E1_gp_full_nm = 2
Const E1_gp_lvl = 3
Const E1_gp_seq = 4
Const E1_gp_eng_nm = 5
Const E1_bdg_ctrl_fg = 6
Const E1_group_type = 7
const E1_par_gp_cd = 8

Dim Exp_Acct_Cd
Const E2_acct_cd		= 0
Const E2_acct_seq		= 1
Const E2_acct_nm		= 2
Const E2_acct_full_nm	= 3
Const E2_bdg_ctrl_fg	= 4
Const E2_bal_fg			= 5
Const E2_bs_pl_fg		= 6
Const E2_del_fg			= 7
Const E2_fx_eval_fg		= 8
Const E2_temp_acct_fg	= 9
Const E2_acct_type		= 10
Const E2_hq_brch_fg		= 11
Const E2_rel_biz_area_cd = 12
Const E2_rel_biz_area_nm = 13
Const E2_subledger_1	= 14
Const E2_subledger_1_nm = 15
Const E2_subledger_2	= 16
Const E2_subledger_2_nm = 17
Const E2_acct_eng_nm	= 18
Const E2_bdg_ctrl_gp_lvl = 19
Const E2_temp_fg_3		= 20
Const E2_temp_fg_4		= 21
Const E2_temp_fg_5		= 22
Const E2_temp_fg_6		= 23
Const E2_temp_fg_7		= 24
Const E2_gp_cd			= 25
Const E2_mgnt_type      = 26
Const E2_txtMgnt_Cd1	= 27
Const E2_txtMgnt_Cd1_Nm	= 28
Const E2_txtMgnt_Cd2	= 29
Const E2_txtMgnt_Cd2_Nm	= 30
Const E2_txtAcct_type_nm= 31
Const E2_txtBs_pl_fg_nm	= 32
Const E2_txtGp_type_nm	     = 33
Const E2_Subsys_type	     = 34
Const E2_subledger_modigy_fg = 35
Const E2_mgnt_cd_modigy_fg   = 36
Const E2_acct_type_modigy_fg = 37

Dim ExpG_Ctrl_item
Const EG1_E2_ctrl_cd = 0
Const EG1_E2_ctrl_nm = 1
Const EG1_E2_ctrl_item_seq = 2
Const EG1_E2_dr_fg = 3
Const EG1_E2_cr_fg = 4
Const EG1_E2_default_value = 5
Const EG1_E2_default_gl_field = 6
Const EG1_E2_sys_fg = 7
Const EG1_E2_mandatory_fg = 8

strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

If strMode = "" Then
	Response.End 
ElseIf strMode <> CStr(UID_M0001) Then										'☜ : 조회 전용 Biz 이므로 다른값은 그냥 종료함 
	Response.End 
ELSEIF 	Request("strkey") = "" THEN
	Response.End
End If

strKey = Request("strKey")

Set iPD1G041 = Server.CreateObject("PD1G041.cALkUpAcctSvr")

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err, True) = True Then					
	Response.End 
End If    

redim Exp_Acct_Gp(7)
'-----------------------
'Com Action Area
'-----------------------
IF UCase(Trim(Request("strCmd"))) = "LOOKUPAC" THEN
	Call iPD1G041.A_LOOKUP_ACCT_SVR(gStrGlobalCollection,Trim(Request("strCmd")),,Trim(strKey),,Exp_Acct_Cd,ExpG_Ctrl_item)
ELSE
	Call iPD1G041.A_LOOKUP_ACCT_SVR(gStrGlobalCollection,Trim(Request("strCmd")),Trim(strKey),,Exp_Acct_Gp)
END IF	

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err, True) = True Then					
	Set iPD1G041 = Nothing
	Response.End 
End If    


Response.Write " <Script Language=vbscript>	" & vbCr
Response.Write " With parent.frm1           " & vbCr														

If Request("strCmd") = "LOOKUPGP" Then
	Response.Write ".txtGP_CD.value          = """ & ConvSPChars(Exp_Acct_Gp(E1_gp_cd))       & """" & vbCr
	Response.Write ".txtGP_SH_NM.value       = """ & ConvSPChars(Exp_Acct_Gp(E1_gp_nm))       & """" & vbCr
	Response.Write ".txtGP_FULL_NM.value     = """ & ConvSPChars(Exp_Acct_Gp(E1_gp_full_nm))  & """" & vbCr
	Response.Write ".txtGP_ENG_NM.value	     = """ & ConvSPChars(Exp_Acct_Gp(E1_gp_eng_nm))   & """" & vbCr
	Response.Write ".txtGP_LVL.value         = """ & Exp_Acct_Gp(E1_gp_lvl)					  & """" & vbCr
	Response.Write ".txtGP_SEQ.value         = """ & Exp_Acct_Gp(E1_gp_seq)					  & """" & vbCr
	Response.Write ".cboGP_BDG_CTRL_FG.value = """ & ConvSPChars(Exp_Acct_Gp(E1_bdg_ctrl_fg)) & """" & vbCr
	Response.Write ".txtParentGp_Cd.value    = """ & ConvSPChars(Exp_Acct_Gp(E1_par_gp_cd))   & """" & vbCr
Else
	NegaValue = "N"
	PosiValue = "Y"
	
	Response.Write ".txtACCT_CD.value		      = """ & ConvSPChars(Exp_Acct_Cd(E2_acct_cd))					 & """" & vbCr
	Response.Write ".txtACCT_SH_NM.value	      = """ & ConvSPChars(Exp_Acct_Cd(E2_acct_nm))					 & """" & vbCr
	Response.Write ".txtACCT_FULL_NM.value	      = """ & ConvSPChars(Exp_Acct_Cd(E2_acct_full_nm))				 & """" & vbCr
	Response.Write ".txtACCT_SEQ.value		      = """ & Exp_Acct_Cd(E2_acct_seq)								 & """" & vbCr
	Response.Write ".txtACCT_TYPE.value		      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_acct_type)))			 & """" & vbCr
	Response.Write ".txtACCT_TYPE_NM.value	      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_txtAcct_type_nm)))	 & """" & vbCr
	Response.Write ".cboBAL_FG.value		      = """ & ConvSPChars(Exp_Acct_Cd(E2_bal_fg))					 & """" & vbCr		
	Response.Write ".txtBS_PL_FG.value		      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_bs_pl_fg)))			 & """" & vbCr
	Response.Write ".txtBS_PL_FG_NM.value	      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_txtBs_pl_fg_nm)))		 & """" & vbCr
	Response.Write ".cboBDG_CTRL_FG.value	      = """ & ConvSPChars(Exp_Acct_Cd(E2_bdg_ctrl_fg))				 & """" & vbCr
	Response.Write ".txtBDG_CTRL_GP_LVL.value     = """ & Exp_Acct_Cd(E2_bdg_ctrl_gp_lvl)						 & """" & vbCr
	Response.Write ".cboFX_EVAL_FG.value	      = """ & ConvSPChars(Exp_Acct_Cd(E2_fx_eval_fg))				 & """" & vbCr
	Response.Write ".cboTEMP_ACCT_FG.value	      = """ & ConvSPChars(Exp_Acct_Cd(E2_temp_acct_fg))				 & """" & vbCr		
	Response.Write ".txtGP_TYPE.value		      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_temp_fg_3)))			 & """" & vbCr		
	Response.Write ".txtGP_TYPE_NM.value	      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_txtGp_type_nm)))		 & """" & vbCr		
	Response.Write ".cboDEL_FG.value		      = """ & ConvSPChars(Exp_Acct_Cd(E2_del_fg))					 & """" & vbCr		
	If Exp_Acct_Cd(E2_hq_brch_fg) = "" Then
		Response.Write ".cboHQ_BRCH_FG.value      = """ & ConvSPChars(NegaValue)								 & """" & vbCr		                                 
	else
		Response.Write ".cboHQ_BRCH_FG.value      = """ & ConvSPChars(Exp_Acct_Cd(E2_hq_brch_fg))				 & """" & vbCr		
	end if
	Response.Write ".txtREL_BIZ_AREA_CD.value     = """ & ConvSPChars(Exp_Acct_Cd(E2_rel_biz_area_cd))			 & """" & vbCr
	Response.Write ".txtREL_BIZ_AREA_Nm.value     = """ & ConvSPChars(Exp_Acct_Cd(E2_rel_biz_area_nm))			 & """" & vbCr
	Response.Write ".txtSUBLEDGER1.value	      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_subledger_1)))		 & """" & vbCr
	Response.Write ".txtSUBLEDGER1_Nm.value	      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_subledger_1_nm)))		 & """" & vbCr
	Response.Write ".txtSUBLEDGER2.value	      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_subledger_2)))		 & """" & vbCr
	Response.Write ".txtSUBLEDGER2_Nm.value	      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_subledger_2_nm)))		 & """" & vbCr							
	Response.Write ".txtParentGp_Cd.value         = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_gp_cd)))				 & """" & vbCr
	Response.Write ".cboMgntType.value		      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_mgnt_type)))	    	 & """" & vbCr								
	Response.Write ".txtMgntCd1.value		      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_txtMgnt_Cd1)))		 & """" & vbCr							
	Response.Write ".txtMgntCd1_Nm.value	      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_txtMgnt_Cd1_Nm)))		 & """" & vbCr							
	Response.Write ".txtMgntCd2.value		      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_txtMgnt_Cd2)))		 & """" & vbCr							
	Response.Write ".txtMgntCd2_Nm.value	      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_txtMgnt_Cd2_Nm)))		 & """" & vbCr		
	Response.Write ".txtACCT_TYPE_Nm.value	      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_txtAcct_type_nm)))	 & """" & vbCr		
	Response.Write ".txtBS_PL_FG_Nm.value	      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_txtBs_pl_fg_nm)))		 & """" & vbCr		
	Response.Write ".txtGP_TYPE_Nm.value	      = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_txtGp_type_nm)))       & """" & vbCr		
	Response.Write ".cboSubSystemType.value		  = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_Subsys_type)))         & """" & vbCr					
	Response.Write ".txtsubledger_modigy_fg.value = """ & Trim(ConvSPChars(Exp_Acct_Cd(txtsubledger_modigy_fg))) & """" & vbCr					
	Response.Write ".txtmgnt_modigy_fg.value	  = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_mgnt_cd_modigy_fg)))   & """" & vbCr					
	Response.Write ".txtaccttype_modigy_fg.value  = """ & Trim(ConvSPChars(Exp_Acct_Cd(E2_acct_type_modigy_fg))) & """" & vbCr
	If Trim(ConvSPChars(Exp_Acct_Cd(E2_Subsys_type))) = "OD" Or Trim(ConvSPChars(Exp_Acct_Cd(E2_Subsys_type))) = "OC" Then
		Response.Write ".txtOpenAcctFg.value = """ & ConvSPChars(PosiValue)	& """" & vbCr		                                 
	Else
		Response.Write ".txtOpenAcctFg.value = """ & ConvSPChars(NegaValue)	& """" & vbCr		                                 
	End If	
	Response.Write ".vspdData.MaxRows = 0 " & vbCr
  
	If Not isEmpty(ExpG_Ctrl_Item) Then 
		For LngRow = 0 To UBound(ExpG_Ctrl_Item,1)
		    strData = strData & Chr(11) & ConvSPChars(ExpG_Ctrl_Item(LngRow,EG1_E2_ctrl_cd)) 
		    strData = strData & Chr(11) & ""
		    strData = strData & Chr(11) & ConvSPChars(ExpG_Ctrl_Item(LngRow,EG1_E2_ctrl_nm))  
		    strData = strData & Chr(11) & ExpG_Ctrl_Item(LngRow,EG1_E2_ctrl_item_seq)
		    strData = strData & Chr(11) & ConvSPChars(ExpG_Ctrl_Item(LngRow,EG1_E2_dr_fg))                 
		    strData = strData & Chr(11) & ConvSPChars(ExpG_Ctrl_Item(LngRow,EG1_E2_cr_fg))
		    strData = strData & Chr(11) & ConvSPChars(ExpG_Ctrl_Item(LngRow,EG1_E2_default_value))
		    strData = strData & Chr(11) & ConvSPChars(ExpG_Ctrl_Item(LngRow,EG1_E2_default_gl_field))
		    strData = strData & Chr(11) & ""
		    strData = strData & Chr(11) & ConvSPChars(ExpG_Ctrl_Item(LngRow,EG1_E2_sys_fg))
		    strData = strData & Chr(11) & ConvSPChars(ExpG_Ctrl_Item(LngRow,EG1_E2_mandatory_fg))		    
		    strData = strData & Chr(11) & ""
		    strData = strData & Chr(11) & LngRow                                               
		    strData = strData & Chr(11) & Chr(12)	   					
		Next

		Response.Write "parent.ggoSpread.Source = .vspdData " & vbCr
		Response.Write "parent.ggoSpread.SSShowData """ & strData                                     & """" & vbCr			
	End if		
End IF
	
	Response.write "End With	      " & vbCr
	Response.write "parent.DbQueryOk  " & vbCr
	Response.write " </Script>        " & vbCr

Set iPD1G041 = Nothing                                                    '☜: Unload Complus
%>
