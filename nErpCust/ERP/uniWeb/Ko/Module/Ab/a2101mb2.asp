<%@ LANGUAGE=VBSCript%>
<%Option Explicit%> 
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->


<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*","NOCOOKIE","MB")

Call HideStatusWnd		

On Error Resume Next

Dim iPABG005																	'☆ : 조회용 ComPlus Dll 사용 변수 
Dim iErrorPosition	

Dim IntRows
Dim IntCols
Dim vbIntRet
Dim lEndRow
Dim boolCheck
Dim lgIntFlgMode
Dim LngMaxRow
Dim strCmd
Dim iCommandSent
Dim txtSpread
Dim imp_par_gp_cd

Dim Imp_gp_cd
Const I5_gp_cd = 0
Const I5_gp_nm = 1
Const I5_gp_full_nm = 2
Const I5_gp_eng_nm = 3
Const I5_gp_lvl = 4
Const I5_gp_seq = 5
Const I5_group_type = 6
Const I5_bdg_ctrl_fg = 7

Dim imp_acct_cd
Const I6_acct_cd = 0
Const I6_acct_seq = 1
Const I6_acct_nm = 2
Const I6_acct_full_nm = 3
Const I6_acct_eng_nm = 4
Const I6_bdg_ctrl_fg = 5
Const I6_bdg_ctrl_gp_lvl = 6
Const I6_bal_fg = 7
Const I6_bs_pl_fg = 8
Const I6_del_fg = 9
Const I6_fx_eval_fg = 10
Const I6_temp_acct_fg = 11
Const I6_acct_type = 12
Const I6_hq_brch_fg = 13
Const I6_rel_biz_area_cd = 14
Const I6_temp_fg_3 = 15
Const I6_subledger_1 = 16
Const I6_subledger_2 = 17
Const I6_mgnt_type = 18
Const I6_mgnt_cd1 = 19
Const I6_mgnt_cd2 = 20
Const I6_subsys_type = 21

IF Request("txtlgMode") <> "" Then
	lgIntFlgMode = CInt(Request("txtlgMode"))								'☜: 저장시 Create/Update 판별 
END IF	
strCmd = Request("lgstrCmd")												'☜: 저장시 Create/Update 판별 

Set iPABG005 = Server.CreateObject("PABG005.cAMngAcctSvr")

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If CheckSYSTEMError(Err, True) = True Then				
	Response.Write "<Script Language=vbscript>				" & vbCr
	Response.write " .frm1.uniTree1.MousePointer = 0	" & vbCr
	Response.Write " </Script>								" & vbCr
	Response.End															'☜: 비지니스 로직 처리를 종료함 
End If    
'-----------------------
'Data manipulate area
'-----------------------			
IF Request("lgstrCmd") = "GP" THEN		
	IF Request("txtParentGP_CD") = "" Then
		imp_par_gp_cd = ""		
	Else	
		imp_par_gp_cd = Trim(Request("txtParentGP_CD"))
	end if	
	redim Imp_gp_cd(I5_bdg_ctrl_fg)
	Imp_gp_cd(I5_gp_cd)			= Trim(Request("txtGP_CD"))
	Imp_gp_cd(I5_gp_nm)			= Trim(Request("txtGP_SH_NM"))
	Imp_gp_cd(I5_gp_full_nm)	= Trim(Request("txtGP_FULL_NM"))
	Imp_gp_cd(I5_gp_eng_nm)		= Trim(Request("txtGP_ENG_NM"))
	Imp_gp_cd(I5_gp_lvl)		= UNIConvNum(Request("txtGP_LVL"),0)	
	Imp_gp_cd(I5_gp_seq)		= UNIConvNum(Request("txtGP_SEQ"),0)
	Imp_gp_cd(I5_bdg_ctrl_fg)	= Request("cboGP_BDG_CTRL_FG")

ELSE
	imp_par_gp_cd		= Trim(Request("txtParentGP_CD"))

	Redim imp_acct_cd(I6_subsys_type)
	imp_acct_cd(I6_acct_cd)			= Trim(Request("txtACCT_CD"))
	imp_acct_cd(I6_acct_seq)		= UNIConvNum(Request("txtACCT_SEQ"),0)
	imp_acct_cd(I6_acct_nm)			= Trim(Request("txtACCT_SH_NM"))
	imp_acct_cd(I6_acct_full_nm)	= Trim(Request("txtACCT_FULL_NM"))
	imp_acct_cd(I6_acct_eng_nm)		= Trim(Request("txtACCT_FULL_NM")) 
	imp_acct_cd(I6_bdg_ctrl_fg)		= Request("cboBDG_CTRL_FG")	
	imp_acct_cd(I6_bdg_ctrl_gp_lvl)	= UNIConvNum(Request("txtBDG_CTRL_GP_LVL"),0)
	imp_acct_cd(I6_bal_fg)			= Request("cboBAL_FG")
	imp_acct_cd(I6_bs_pl_fg)		= Request("txtBS_PL_FG")
	imp_acct_cd(I6_del_fg)			= Request("cboDEL_FG")
	imp_acct_cd(I6_fx_eval_fg)		= Request("cboFX_EVAL_FG")
	imp_acct_cd(I6_temp_acct_fg)	= Request("cboTEMP_ACCT_FG")
	imp_acct_cd(I6_acct_type)		= Request("txtACCT_TYPE")	
	imp_acct_cd(I6_temp_fg_3)		= Trim(Request("txtGP_TYPE"))

	if Request("cboHQ_BRCH_FG") = "" then
		imp_acct_cd(I6_hq_brch_fg)	= "N"
	Else
		imp_acct_cd(I6_hq_brch_fg)	= Request("cboHQ_BRCH_FG")
	End if
	imp_acct_cd(I6_rel_biz_area_cd)	= Trim(Request("txtREL_BIZ_AREA_CD")) 
	imp_acct_cd(I6_subledger_1)		= Trim(Request("txtSUBLEDGER1")) 
	imp_acct_cd(I6_subledger_2)		= Trim(Request("txtSUBLEDGER2"))
	imp_acct_cd(I6_mgnt_type)		= Trim(Request("cboMgntType"))	
	imp_acct_cd(I6_mgnt_cd1)		= Trim(Request("txtMgntCd1")) 
	imp_acct_cd(I6_mgnt_cd2)		= Trim(Request("txtMgntCd2")) 
	imp_acct_cd(I6_subsys_type)     = Trim(Request("cboSubSystemType")) 
END IF

If lgIntFlgMode = OPMD_CMODE Then
	IF strcmd = "GP" then
		iCommandSent = "CREATEGP"
	ELSE
		iCommandSent = "CREATE"
	END IF
ElseIf lgIntFlgMode = OPMD_UMODE Then
	IF strcmd = "GP" then
		iCommandSent = "UPDATEGP"
	ELSE
		iCommandSent = "UPDATE"
	END IF
Else
	IF strcmd = "GP" then
		iCommandSent = "DELETEGP"
	ELSE
		iCommandSent = "DELETE"
	END IF
End If

'-----------------------
'Com action  area
'-----------------------
txtSpread = Trim(request("txtSpread")) 
Call iPABG005.A_MANAGE_ACCT_SVR(gStrGloBalCollection,iCommandSent,,,,imp_par_gp_cd,Imp_gp_cd,imp_acct_cd, _
                                      txtSpread,iErrorPosition)

'-----------------------
'Com action result check area(OS,internal)
'-----------------------

'if CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
If CheckSYSTEMError(Err, True) = True Then					
	Set iPABG005 = Nothing												'☜: ComPlus Unload	
	Response.Write "<Script Language=vbscript>				" & vbCr
	Response.write " parent.frm1.uniTree1.MousePointer = 0	" & vbCr
	Response.Write " </Script>								" & vbCr
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

Set iPABG005 = Nothing                                                  '☜: Unload Complus

Response.Write "<Script Language=vbscript> " & vbCr
Response.Write " parent.DbSaveOk           " & vbCr 
'Response.Write " parent.DbQueryOk           " & vbCr 
Response.Write " </Script>                 " & vbCr

Response.Write " <Script Language=vbscript RUNAT=server> " & vbCr

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function ArrowMouse()

	Dim temp
	Dim strHTML

	temp = "D"

	strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
	strHTML = strHTML & "if parent.lgSaveModFg = ""D"" then " & vbCrLf
   	strHTML = strHTML & "parent.lgSaveModFg = """" " & vbCrLf
	strHTML = strHTML & "end if " & vbCrLf
	strHTML = strHTML & "parent.frm1.uniTree1.MousePointer = 0 " & vbCrLf
	strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
	Response.Write strHTML

End Function
Response.Write "</Script> " & vbCr
%>