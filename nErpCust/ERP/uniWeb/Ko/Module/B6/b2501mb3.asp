<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : B2501mb3.asp
'*  4. Program Name         : Entry Plant (Delete)
'*  5. Program Desc         :
'*  6. Component List       : PB6G102.cBMngPlt
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

On Error Resume Next
Err.Clear

Dim pPB6G102

Dim I1_b_plant, iCommandSent

' I1_b_plant
Const P062_I1_plant_cd = 0
Redim I1_b_plant(P062_I1_plant_cd)
If Request("txtPlantCd2") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
End If

'-----------------------
'Data manipulate area
'-----------------------												'⊙: Single 데이타 저장 
I1_b_plant(P062_I1_plant_cd) = Request("txtPlantCd2")
iCommandSent = "DELETE"
	
Set pPB6G102 = Server.CreateObject("PB6G102.cBMngPlt")
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPB6G102.B_MANAGE_PLANT(gStrGlobalCollection, iCommandSent, I1_b_plant)
If CheckSYSTEMError(Err,True) = True Then
	Set pPB6G102 = Nothing                                                   '☜: Unload Component
	Response.End
End If
	
Set pPB6G102 = Nothing                                                   '☜: Unload Component

Response.Write "<Script Language=vbscript>" & vbCrLf
	Response.Write "parent.DbDeleteOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>