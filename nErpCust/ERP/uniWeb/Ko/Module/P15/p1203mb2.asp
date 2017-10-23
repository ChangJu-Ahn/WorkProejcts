<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1203mb2.asp
'*  4. Program Name         : Entry Routing(Create, Update)
'*  5. Program Desc         :
'*  6. Component List       : PP1S502.cPMngRtng
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2000/11/20
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim pPP1S502																	'☆ : 입력/수정용 Component Dll 사용 변수 
Dim I1_plant_cd, I2_item_cd, I2_ALTRTVALUE, I3_p_routing_header, iCommandSent, iErrorPosition
Dim iIntFlgMode

'rout header
Const P137_I3_rout_no = 0
Const P137_I3_bom_no = 1
Const P137_I3_major_flg = 2
Const P137_I3_description = 3
Const P137_I3_valid_from_dt = 4
Const P137_I3_valid_to_dt = 5
Const P137_I3_cost_center = 6

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount
Dim iDCount

Dim ii

Redim I3_p_routing_header(P137_I3_cost_center)

If Request("txtPlantCd") = "" Then												'⊙: 조회를 위한 값이 들어왔는지 체크 
	Call DisplayMsgBox("189220", vbOKOnly, "", "", I_MKSCRIPT)                          
	Response.End
End If
	
iIntFlgMode = Request("txtFlgMode")

If CInt(iIntFlgMode) = CInt(OPMD_CMODE) Then
	iCommandSent = "CREATE"
ElseIf CInt(iIntFlgMode) = CInt(OPMD_UMODE) Then
	iCommandSent = "UPDATE"
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
'Routing Header
'-----------------------
I1_plant_cd		= UCase(Trim(Request("txtPlantCd")))
I2_item_cd		= UCase(Trim(Request("txtItemCd1")))
I2_ALTRTVALUE	= UNIConvNum(Request("txtALTRTVALUE"), 0)


I3_p_routing_header(P137_I3_rout_no)	= UCase(Trim(Request("txtRoutingNo1")))
I3_p_routing_header(P137_I3_bom_no)		= Trim(Request("txtBomNo"))
I3_p_routing_header(P137_I3_major_flg)	= UCase(Trim(Request("rdoMajorRouting")))
I3_p_routing_header(P137_I3_description)= Request("txtRoutingNm1")
I3_p_routing_header(P137_I3_cost_center) = 	UCase(Trim(Request("txtCostCd")))

If Len(Trim(Request("txtValidFromDt"))) Then
	If UniConvDate(Request("txtValidFromDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidFromDt", 0, I_MKSCRIPT)
		Response.End	
	Else
		I3_p_routing_header(P137_I3_valid_from_dt) = UniConvDate(Request("txtValidFromDt"))
	End If
End If
	
If Len(Trim(Request("txtValidToDt"))) Then
	If UniConvDate(Request("txtValidToDt")) = "" Then	 
		Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		Call LoadTab("parent.frm1.txtValidToDt", 0, I_MKSCRIPT)
		Response.End	
	Else
		I3_p_routing_header(P137_I3_valid_to_dt) = UniConvDate(Request("txtValidToDt"))
	End If
End If  

Set pPP1S502 = Server.CreateObject("PP1S502.cPMngRtng")

If CheckSYSTEMError(Err,True) = True Then
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

Call pPP1S502.P_MANAGE_ROUTING(gStrGlobalCollection, iCommandSent, itxtSpread, _
                               I1_plant_cd, I2_item_cd, I3_p_routing_header,I2_ALTRTVALUE, iErrorPosition)

If CheckSYSTEMError2(Err, True, iErrorPosition & "행", "", "", "", "") = True Then
	Set pPP1S502 = Nothing															'☜: Unload Component
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

Set pPP1S502 = Nothing																'☜: Unload Component

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End
%>
