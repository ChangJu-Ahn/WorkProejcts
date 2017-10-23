<%@ LANGUAGE = VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1401mb16.asp
'*  4. Program Name         : BOM Delete Multi
'*  5. Program Desc         :
'*  6. Component List       : PP1S407.cPMngBomHdrMulti
'*  7. Modified date(First) : 2001/10/30
'*  8. Modified date(Last)  : 2002/11/19
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Call LoadBasisGlobalInf

Dim pPP1S407																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim iCommandSent, iErrorPosition
Dim I1_select_char, I2_p_bom_header, I3_plant_cd, I4_item_cd

Const P1A2_I2_bom_no	= 0

If Request("txtPlantCd") = "" Then														'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)            
	Response.End 
ElseIf Request("txtItemCd") = "" Then													'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)            
	Response.End 
ElseIf Request("txtBomNo") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700114", vbOKOnly, "", "", I_MKSCRIPT)            
	Response.End 	
End If
	
'-----------------------
'Data manipulate area
'-----------------------												'⊙: Single 데이타 저장 
Redim I2_p_bom_header(P1A2_I2_bom_no)

I2_p_bom_header(P1A2_I2_bom_no)	= UCase(Trim(Request("txtBomNo")))
I3_plant_cd		= UCase(Trim(Request("txtPlantCd")))
I4_item_cd		= UCase(Trim(Request("txtItemCd")))

iCommandSent = "DELETE"
	
'-----------------------
'Com action result check area(OS,internal)
'-----------------------
Set pPP1S407 = Server.CreateObject("PP1S407.cPMngBomHdrMulti")    

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

Call pPP1S407.P_MANAGE_BOM_HEADER_MULTI(gStrGlobalCollection, iCommandSent, "", _
				 I1_select_char, I2_p_bom_header, I3_plant_cd, I4_item_cd, iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "행", "", "", "", "") = True Then
	Set pPP1S407 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPP1S407 = Nothing													'☜: Unload Comproxy
	
Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "parent.DbDeleteOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>