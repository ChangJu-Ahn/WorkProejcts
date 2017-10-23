<%@ LANGUAGE = VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1711mb4.asp	
'*  4. Program Name         : BOM Header Entry
'*  5. Program Desc         :
'*  6. Component List       : PP1S401.cPMngBomHdr
'*  7. Modified date(First) : 2000/05/2
'*  8. Modified date(Last)  : 2002/11/19
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Call LoadBasisGlobalInf

Dim pPP1S401																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim iCommandSent, I1_plant_cd, I2_item_cd, I3_p_bom_header

Const P131_I3_bom_no = 0
Const P131_I3_description = 1
Const P131_I3_valid_from_dt = 2
Const P131_I3_valid_to_dt = 3
Const P131_I3_drawing_path = 4

If Request("txtPlantCd") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
End If

If Request("txtItemCd1") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)
	Response.End 
End If
	
'-----------------------
'Data manipulate area
'-----------------------
Redim I3_p_bom_header(P131_I3_drawing_path)
			
I1_plant_cd	= UCase(Trim(Request("txtPlantCd")))
I2_item_cd	= UCase(Trim(Request("txtItemCd1")))
	
I3_p_bom_header(P131_I3_description) = Request("txtBomDesc")
I3_p_bom_header(P131_I3_drawing_path) = Request("txtDrawPath")
	
If Request("txtHdrMode") = "C" Then
	I3_p_bom_header(P131_I3_bom_no)			= UCase(Trim(Request("txtBomNo1")))
	iCommandSent = "CREATE"
	I3_p_bom_header(P131_I3_valid_from_dt)	= UniConvDate(Request("txtValidFromDt"))
	I3_p_bom_header(P131_I3_valid_to_dt)	= UniConvDate(Request("txtValidToDt"))
ElseIf Request("txtHdrMode") = "U" Then
	I3_p_bom_header(P131_I3_bom_no) 		= UCase(Trim(Request("txtBomNo1")))
	iCommandSent = "UPDATE"
End If


Set pPP1S401 = Server.CreateObject("PP1S401.cPMngBomHdr")
	    
If CheckSYSTEMError(Err,True) = True Then
	Set pPP1S401 = Nothing		
	Response.End
End If
	
Call pPP1S401.P_MANAGE_BOM_HEADER(gStrGlobalCollection, iCommandSent, I1_plant_cd, I2_item_cd, I3_p_bom_header)

If CheckSYSTEMError(Err, True) = True Then
	Set pPP1S401 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPP1S401 = Nothing      

Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf
		Response.Write ".frm1.txtItemCd.value = """ & ConvSPChars(Request("txtItemCd1")) & """" & vbCrLf
		Response.Write ".frm1.txtItemNm.value = """ & ConvSPChars(Request("txtItemNm1")) & """" & vbCrLf
		Response.Write ".frm1.txtBomNo.value = """ & ConvSPChars(Request("txtBomNo1")) & """" & vbCrLf
		Response.Write ".frm1.hBomType.value = """ & ConvSPChars(Request("txtBomNo1")) & """" & vbCrLf
		Response.Write ".DbSaveOk" & vbCrLf
	Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End																				'☜: Process End
%>