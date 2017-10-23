<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1115MB3
'*  4. Program Name         : 연/월 품질목표등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG090
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/08/12
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")
Call LoadBasisGlobalInf

On Error Resume Next
Call HideStatusWnd 

Dim PQBG090																	'☆ : 조회용 ComProxy Dll 사용 변수 

Dim iErrorPosition
Dim iCommandSent

Dim I1_q_yearly_target
    'Const Q025_I1_plant_cd = 0
    'Const Q025_I1_insp_class_cd = 1
    'Const Q025_I1_yr = 2
    'Const Q025_I1_target_value = 3
    'Const Q025_I1_defect_unit_cd = 4
ReDim I1_q_yearly_target(4)

Dim IG1_q_monthly_target
    'Const Q025_IG1_row_num = 0
    'Const Q025_IG1_mnth = 1
    'Const Q025_IG1_monthly_target_value = 2
ReDim IG1_q_monthly_target(11, 2)	   	

	Err.Clear                                                               '☜: Protect system from crashing
			
	iCommandSent = "DELETE"
	I1_q_yearly_target(0) = UCase(Trim(Request("txtPlantCd")))
	I1_q_yearly_target(1) = Request("cboInspClassCd")	
	I1_q_yearly_target(2) = Request("txtYr")	


	If I1_q_yearly_target(0) = "" Or I1_q_yearly_target(1) = "" Or I1_q_yearly_target(2) = "" Then
		Call DisplayMsgBox("229909", vbOKOnly, "", "", I_MKSCRIPT) 
		Response.End 
	End If

	Set PQBG090 = Server.CreateObject("PQBG090.cQMaintTargetSvr")
			
	If CheckSYSTEMError(Err,True) Then
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)
		Response.End
	End if

	Call PQBG090.Q_MAINT_TARGET_SVR (gStrGlobalCollection, _
									iCommandSent, _
									I1_q_yearly_target, _
									IG1_q_monthly_target, _
									iErrorPosition)

			
	If CheckSYSTEMError(Err,True) Then
		Set PQBG090 = Nothing  
		Response.End
	End If	           	

	Set PQBG090 = Nothing
%>
<Script Language=vbscript>
	Call parent.DbDeleteOk()
</Script>