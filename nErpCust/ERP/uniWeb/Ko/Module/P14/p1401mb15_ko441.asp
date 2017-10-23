<%@ LANGUAGE = VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1401mb15.asp
'*  4. Program Name         : BOM entry Multi
'*  5. Program Desc         :
'*  6. Component List       : PP1S407.cPMngBomHdrMulti
'*  7. Modified date(First) : 2001/10/30
'*  8. Modified date(Last)  : 2003/03/18
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
Dim iIntFlgMode

Const P1A2_I2_bom_no		= 0
Const P1A2_I2_description	= 1
Const P1A2_I2_valid_from_dt	= 2
Const P1A2_I2_valid_to_dt	= 3
Const P1A2_I2_drawing_path	= 4

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii

If Request("txtPlantCd") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)                          
	Response.End
End If
	
iIntFlgMode = CInt(Request("txtFlgMode"))

If iIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf iIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
End If

Redim I2_p_bom_header(P1A2_I2_drawing_path)

I2_p_bom_header(P1A2_I2_bom_no)			= UCase(Trim(Request("hBomType")))
I2_p_bom_header(P1A2_I2_valid_from_dt)	= UniConvDate(Trim(Request("txtHdrValidFromDt")))
I2_p_bom_header(P1A2_I2_valid_to_dt)	= UniConvDate(Trim(Request("txtHdrValidToDt")))
I2_p_bom_header(P1A2_I2_drawing_path)	= Trim(Request("txtHdrDrawingPath"))
I3_plant_cd		= UCase(Trim(Request("txtPlantCd")))
I4_item_cd		= UCase(Trim(Request("hItemCd")))


itxtSpread = ""
             
iCUCount = Request.Form("txtCUSpread").Count
iDCount  = Request.Form("txtDSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount + iDCount)
             
For ii = 1 To iDCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
Next
For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
Set pPP1S407 = Server.CreateObject("PP1S407.cPMngBomHdrMulti")    

If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

Call pPP1S407.P_MANAGE_BOM_HEADER_MULTI(gStrGlobalCollection, iCommandSent, itxtSpread, _
				 I1_select_char, I2_p_bom_header, I3_plant_cd, I4_item_cd, iErrorPosition)

If CheckSYSTEMError(Err,True) = True Then
    Set pPP1S407 = Nothing
    Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
    Response.End 
End If

Set pPP1S407 = Nothing													'☜: Unload Component
Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf
		Response.Write ".frm1.txtPlantCd.Value = """ & ConvSPChars(Trim(Request("txtPlantCd"))) & """" & vbCrLf
		Response.Write ".frm1.txtItemCd.Value = """ & ConvSPChars(Trim(Request("hItemCd"))) & """" & vbCrLf
		Response.Write ".frm1.txtBomno.Value = """ & ConvSPChars(Trim(Request("hBomType"))) & """" & vbCrLf
		Response.Write ".DbSaveOk" & vbCrLf
	Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>