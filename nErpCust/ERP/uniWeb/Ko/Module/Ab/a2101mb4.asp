<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->

<%
	Call LoadBasisGlobalInf()
	Call HideStatusWnd

	On Error Resume Next
	Err.Clear 

	Dim iPABG005																	'☆ : 저장용 ComPlus DLL 사용 변수 
	Dim iCommandSent
	Dim IntRows
	Dim IntCols
	Dim vbIntRet
	Dim lEndRow
	Dim boolCheck
	Dim lgIntFlgMode
	Dim LngMaxRow
	Dim strCmd
	Dim iCoommandSent 

	Dim imp_to_cd
	Const I1_acct_cd         = 0
	Const I1_acct_seq        = 1
	Const I1_insrt_user_id   = 2

	Dim imp_from_par_gp_cd

	Dim imp_to_gp
	Const I3_gp_cd			 = 0
	Const I3_gp_lvl			 = 1
	Const I3_gp_seq			 = 2

	Dim imp_to_par_gp_cd

	Dim imp_gp_cd
	Const I5_gp_cd			 = 0
	Const I5_gp_nm			 = 1
	Const I5_gp_full_nm		 = 2
	Const I5_gp_eng_nm		 = 3
	Const I5_gp_lvl			 = 4
	Const I5_gp_seq			 = 5
	Const I5_group_type		 = 6
	Const I5_bdg_ctrl_fg	 = 7

	Dim imp_acct_cd
	Const I6_acct_cd		 = 0
	Const I6_acct_seq		 = 1
	Const I6_acct_nm		 = 2
	Const I6_acct_full_nm	 = 3
	Const I6_acct_eng_nm	 = 4
	Const I6_bdg_ctrl_fg	 = 5
	Const I6_bdg_ctrl_gp_lvl = 6
	Const I6_bal_fg			 = 7
	Const I6_bs_pl_fg		 = 8
	Const I6_del_fg			 = 9
	Const I6_fx_eval_fg		 = 10
	Const I6_temp_acct_fg	 = 11
	Const I6_acct_type		 = 12
	Const I6_hq_brch_fg		 = 13
	Const I6_rel_biz_area_cd = 14
	Const I6_temp_fg_3		 = 15
	Const I6_subledger_1	 = 16
	Const I6_subledger_2	 = 17

	strCmd		 = Request("lgstrCmd")												'☜: 저장시 Create/Update 판별 

	Set iPABG005 = Server.CreateObject("PABG005.cAMngAcctSvr")

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Response.End
	End If    

	'-----------------------
	'Data manipulate area
	'-----------------------
	If strCmd = "GP" Then
		iCommandSent = "MOVEGP"

		Redim imp_gp_cd(7)
		Imp_gp_cd(I5_gp_cd)		  = Trim(Request("txtGP_CD"))
		Imp_gp_cd(I5_gp_nm)		  = Trim(Request("txtGP_SH_NM"))
		Imp_gp_cd(I5_gp_full_nm)  = Trim(Request("txtGP_FULL_NM"))
		Imp_gp_cd(I5_gp_eng_nm)	  = Trim(Request("txtGP_ENG_NM"))
		Imp_gp_cd(I5_gp_lvl)	  = UNIConvNum(Request("txtGP_LVL"),0)
		Imp_gp_cd(I5_gp_seq)	  = UNIConvNum(Request("txtGP_SEQ"),0)
		Imp_gp_cd(I5_bdg_ctrl_fg) = Request("cboGP_BDG_CTRL_FG")

		imp_to_par_gp_cd		  = Trim(Request("txtToParentGP_CD"))

		Redim imp_to_gp(2)
	    imp_to_gp(I3_gp_cd) 	  = Trim(Request("txtToGP_CD"))
		imp_to_gp(I3_gp_lvl)	  = UniconvNum(Request("txtToGP_LVL"),0)
		imp_to_gp(I3_gp_seq)	  = UniconvNum(Request("txtToGP_SEQ"),0)

		imp_from_par_gp_cd		  = Trim(Request("txtParentGP_CD"))
	Else
		iCommandSent = "MOVEACCT"

		Redim imp_to_cd(1)
		imp_to_cd(I1_acct_cd)		= Trim(Request("txtToACCT_CD"))
		imp_to_cd(I1_acct_seq)	    = UniconvNum(Request("txtToACCT_SEQ"),0)

		imp_from_par_gp_cd			= Trim(Request("txtParentGP_CD"))

		Redim imp_acct_cd(17)
		imp_acct_cd(I6_acct_cd)			= Trim(Request("txtACCT_CD"))
		imp_acct_cd(I6_acct_seq)		= UNIConvNum(Request("txtACCT_SEQ"),0)
		imp_acct_cd(I6_acct_nm)			= Trim(Request("txtACCT_SH_NM"))
		imp_acct_cd(I6_acct_full_nm)	= Trim(Request("txtACCT_FULL_NM"))
		imp_acct_cd(I6_acct_eng_nm)		= Trim(Request("txtACCT_FULL_NM"))
		imp_acct_cd(I6_bdg_ctrl_fg)		= Request("cboBDG_CTRL_FG")	
		imp_acct_cd(I6_bdg_ctrl_gp_lvl)	= UNIConvNum(Request("txtBDG_CTRL_GP_LVL"),0)
		imp_acct_cd(I6_bal_fg)			= Request("cboBAL_FG")
		imp_acct_cd(I6_bs_pl_fg)		= Request("cboBS_PL_FG")
		imp_acct_cd(I6_del_fg)			= Request("cboDEL_FG")
		imp_acct_cd(I6_fx_eval_fg)		= Request("cboFX_EVAL_FG")
		imp_acct_cd(I6_temp_acct_fg)	= Request("cboTEMP_ACCT_FG")
		imp_acct_cd(I6_acct_type)		= Request("cboACCT_TYPE")
		imp_acct_cd(I6_temp_fg_3)		= Trim(Request("cboGP_TYPE"))
		
		If Trim(Request("cboHQ_BRCH_FG")) = "" Then
			imp_acct_cd(I6_hq_brch_fg)	= "N"
		Else
			imp_acct_cd(I6_hq_brch_fg)	= Request("cboHQ_BRCH_FG")
		End If
		
		imp_acct_cd(I6_rel_biz_area_cd)	= Trim(Request("txtREL_BIZ_AREA_CD"))
		imp_acct_cd(I6_subledger_1)		= Trim(Request("txtSUBLEDGER1"))
		imp_acct_cd(I6_subledger_2)		= Trim(Request("txtSUBLEDGER2"))

		imp_to_par_gp_cd				= Trim(Request("txtToParentGP_CD"))
	End If

	'-----------------------
	'Com action  area
	'-----------------------
	Call iPABG005.A_MANAGE_ACCT_SVR(gStrGloBalCollection,iCommandSent,imp_to_cd ,imp_to_par_gp_cd,imp_to_gp, _
										imp_from_par_gp_cd,Imp_gp_cd,imp_acct_cd)
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then					
		Set iPABG005 = Nothing
		Response.End
	End If 

	Set iPABG005 = Nothing

	Response.Write "<Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk           " & vbCr
	Response.Write "</Script>				   " & vbCr	
%>