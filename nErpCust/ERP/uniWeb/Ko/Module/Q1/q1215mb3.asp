<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")%>
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1215MB3
'*  4. Program Name         : 선별형검사조건 등록 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG190
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next
Call HideStatusWnd 

Dim PQBG190																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim strPlantCd
Dim strItemCd
Dim strInspClassCd
Dim strInspItemCd																'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strRoutNo
Dim strOprNo
Dim iCommandSent
Dim I6_q_inspection_standard_detail3


Err.Clear                                                               '☜: Protect system from crashing
	
strPlantCd		= Request("txtPlantCd")
strItemCd		= Request("txtItemCd")
strInspClassCd	= Request("cboInspClassCd")
strInspItemCd	= Request("txtInspItemCd")
strRoutNo		= Request("txtRoutNo")
strOprNo		= Request("txtOprNo") 
	
If strPlantCd = "" Or strItemCd = "" Or strInspClassCd = "" Or strItemCd = "" Then
	IF strInspClassCd = "P" AND (strRoutNo = "" Or strOprNo = "" ) Then
		Call DisplayMsgBox("229909", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End
	End if 
End If
	
Set PQBG190 = Server.CreateObject("PQBG190.cQMaintInspDtl3Svr")

If CheckSYSTEMError(Err,True) Then
	Response.End 
End If

iCommandSent = "DELETE"
	
Call PQBG190.Q_MAINT_INSP_DTL3_SVR(gStrGlobalCollection, _
								   iCommandSent, _
								   strPlantCd, strItemCd, strInspItemCd, strInspClassCd, _
								   strRoutNo, strOprNo, I6_q_inspection_standard_detail3)
If CheckSYSTEMError(Err,True) Then
	Set PQBG190 = Nothing
	Response.End 
End If

Set PQBG190 = Nothing
%>
<Script Language=vbscript>
Call parent.DbDeleteOk()
</Script>