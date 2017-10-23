<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%Call LoadBasisGlobalInf%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1301mb3.asp
'*  4. Program Name         : Entry Work Center (Delete)
'*  5. Program Desc         :
'*  6. Component List       : PP1G203.cMngWkCtr
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next														'☜: 
Err.Clear

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pPP1G203																	'☆ : 입력/수정용 Component Dll 사용 변수 
Dim I4_p_work_center, iCommandSent

Const P118_I4_wc_cd = 0

Redim I4_p_work_center(P118_I4_wc_cd)

If Request("txtPlantCd") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)              
	Response.End 
End If
	
If Request("txtDataWcCd") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)         
	Response.End 
End If
	
'-----------------------
'Data manipulate area
'-----------------------												'⊙: Single 데이타 저장 
I4_p_work_center(P118_I4_wc_cd) = Trim(Request("txtDataWcCd"))
iCommandSent = "DELETE"

Set pPP1G203 = Server.CreateObject("PP1G203.cPMngWkCtr")

If CheckSYSTEMError(Err, True) = True Then
	Response.End
End If

Call pPP1G203.P_MANAGE_WORK_CENTER(gStrGlobalCollection, iCommandSent, , , ,I4_p_work_center)

If CheckSYSTEMError(Err, True) = True Then
	Set pPP1G203 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPP1G203 = Nothing																'☜: Unload Component

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.DbDeleteOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>