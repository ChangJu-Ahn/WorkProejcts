<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1209mb3_ko441.asp
'*  4. Program Name         : Entry Standard RoutingEntry (Delete)
'*  5. Program Desc         :
'*  6. Component List       : PP1G507.cPMngStdRtng
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2002/11/21
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim pPP1G507																	'☆ : 입력/수정용 Component Dll 사용 변수 
Dim I1_plant_cd, I2_rout_no, iCommandSent

If Request("txtPlantCd") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)           
	Response.End 
ElseIf Request("txtRoutingNo") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)             
	Response.End 
End If
	
'-----------------------
'Data manipulate area
'-----------------------												'⊙: Single 데이타 저장 
I1_plant_cd = Request("txtPlantCd")
I2_rout_no = Request("txtRoutingNo")
iCommandSent = "DELETE"
	
Set pPP1G507 = Server.CreateObject("PP1G507_KO441.cPMngItemGrpStdRtng")
If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPP1G507.P_MANAGE_STANDARD_ROUTING(gStrGlobalCollection, iCommandSent, , I1_plant_cd, I2_rout_no)
If CheckSYSTEMError(Err,True) = True Then
	Set pPP1G507 = Nothing                                                   '☜: Unload Component
	Response.End
End If

Set pPP1G507 = Nothing                                                   '☜: Unload Component

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.DbDeleteOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>
