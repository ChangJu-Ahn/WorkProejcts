<%@ LANGUAGE = VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : 설계BOM관리 
'*  2. Function Name        : 
'*  3. Program ID           : p1713mb13.asp
'*  4. Program Name         : Request BOM entry Multi
'*  5. Program Desc         :
'*  6. Component List       : PY3S113.cYTransBomHdrMulti
'*  7. Modified date(First) : 2005.01.31
'*  8. Modified date(Last)  : 2005.01.31
'*  9. Modifier (First)     : Cho Yong Chill
'* 10. Modifier (Last)      : Cho Yong Chill
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Call LoadBasisGlobalInf

Dim pPY3S113
Dim pPY3S114
																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
Dim iCommandSent, iErrorPosition
Dim I1_select_char, I2_p_trans_bom_header, I3_plant_cd, I4_item_cd
Dim iIntFlgMode

Const Y311_I2_bom_no			= 0 
Const Y311_I2_req_trans_no		= 1 
Const Y311_I2_design_plant_cd	= 2 
Const Y311_I2_req_trans_dt		= 3 
Const Y311_I2_trans_dt			= 4 
Const Y311_I2_status			= 5 
Const Y311_I2_description		= 6 
Const Y311_I2_valid_from_dt		= 7 
Const Y311_I2_valid_to_dt		= 8 
Const Y311_I2_drawing_path		= 9 
            
Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii

If Request("txtDestPlantCd") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)                          
	Response.End
End If

If Request("txtBasePlantCd") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)                          
	Response.End
End If

iIntFlgMode = CInt(Request("txtFlgMode"))

If iIntFlgMode = OPMD_CMODE Then
	iCommandSent = "CREATE"
ElseIf iIntFlgMode = OPMD_UMODE Then
	iCommandSent = "UPDATE"
End If

Redim I2_p_trans_bom_header(Y311_I2_drawing_path)

I2_p_trans_bom_header(Y311_I2_bom_no)			= 		UCase(	Trim(Request("hBomType")))
I2_p_trans_bom_header(Y311_I2_req_trans_no)		= 		UCase(	Trim(Request("hReqTransNo")))
I2_p_trans_bom_header(Y311_I2_design_plant_cd)	= 		UCase(	Trim(Request("hBasePlantCd")))
I2_p_trans_bom_header(Y311_I2_req_trans_dt)		= 	UniConvDate(Trim(Request("hReqTransDt")))
I2_p_trans_bom_header(Y311_I2_status)			= 		UCase(	Trim(Request("hStatus")))
I2_p_trans_bom_header(Y311_I2_description)		= 				Trim(Request("hDescription"))
I2_p_trans_bom_header(Y311_I2_valid_from_dt)	= 	UniConvDate(Trim(Request("hHdrValidFromDt")))
I2_p_trans_bom_header(Y311_I2_valid_to_dt)		= 	UniConvDate(Trim(Request("hHdrValidToDt")))
I2_p_trans_bom_header(Y311_I2_drawing_path)		= 				Trim(Request("hDrawingPath"))
I3_plant_cd										= 		UCase(	Trim(Request("hDestPlantCd")))
I4_item_cd										= 		UCase(	Trim(Request("hItemCd")))

'=================================================
'이관의뢰처리중인경우 
'=================================================
If Request("hRequestingFlg") = "Y" Then
	
	Set pPY3S114 = Server.CreateObject("PY3S114.cYMngTransBomHdr")
	    
	If CheckSYSTEMError(Err,True) = True Then
		Set pPY3S114 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLF
			Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		Response.End
	End If
	
	Call pPY3S114.Y_UPDATE_STATUS_TRANS_BOM_HEADER(gStrGlobalCollection, I3_plant_cd, I4_item_cd, I2_p_trans_bom_header(Y311_I2_req_trans_no), I2_p_trans_bom_header(Y311_I2_status))
	
	If CheckSYSTEMError(Err,True) = True Then
	    Set pPY3S114 = Nothing
	    Response.Write "<Script Language=VBScript>" & vbCrLF
			Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
		Response.Write "</Script>" & vbCrLF
	    Response.End 
	End If
End If

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

Set pPY3S113 = Server.CreateObject("PY3S113.cYTransBomHdrMulti")    

If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

Call pPY3S113.Y_TRANS_BOM_HEADER_MULTI(gStrGlobalCollection, iCommandSent, itxtSpread, _
				 I1_select_char, I2_p_trans_bom_header, I3_plant_cd, I4_item_cd, iErrorPosition)

If CheckSYSTEMError(Err,True) = True Then
    Set pPY3S113 = Nothing
    Set pPY3S114 = Nothing
    Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
    Response.End 
End If

Set pPY3S113 = Nothing													'☜: Unload Component
Set pPY3S114 = Nothing

Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "With parent" & vbCrLf
		Response.Write ".frm1.txtBasePlantCd.Value = """ 	& ConvSPChars(Trim(Request("hBasePlantCd"))) & """" & vbCrLf
		Response.Write ".frm1.txtDestPlantCd.Value = """ 	& ConvSPChars(Trim(Request("hDestPlantCd"))) & """" & vbCrLf
		Response.Write ".frm1.txtItemCd.Value = """ 		& ConvSPChars(Trim(Request("hItemCd"))) & """" & vbCrLf
'		Response.Write ".frm1.txtBomno.Value = """ 			& ConvSPChars(Trim(Request("hBomType"))) & """" & vbCrLf
		Response.Write ".DbSaveOk" & vbCrLf
	Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>