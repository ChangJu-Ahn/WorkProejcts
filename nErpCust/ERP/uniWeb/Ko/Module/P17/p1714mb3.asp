<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: 설계BOM관리 
'*  2. Function Name		: 
'*  3. Program ID			: p1714mb1.asp
'*  4. Program Name			: 
'*  5. Program Desc			: 
'*  6. Comproxy List		: 
'*  7. Modified date(First) : 2005-02-14
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Yoon, Jeong Woo
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call HideStatusWnd
On Error Resume Next
Dim oPY3S117									'☆ : 입력/수정용 ComProxy Dll 사용 변수 

Dim iErrorPosition

Dim iCommandSent
Dim I0_gubun, I1_plant_cd, I2_item_cd, I3_base_dt
Dim I3_bom_no, I4_req_trans_no	'삭제 

'Dim I1_select_char, I2_p_bom_header, I3_plant_cd, I4_item_cd

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount

Dim ii

Err.Clear										'☜: Protect system from crashing

I0_gubun    = Trim(Request("hgubun"))
I1_plant_cd = Trim(UCase(Request("txtDestPlantCd")))
I2_item_cd  = Trim(UCase(Request("txtItemCd")))
I3_base_dt  = Trim(Request("hStartDate"))

'Response.Write I0_gubun & "<P>"
'Response.Write I1_plant_cd & "<P>"
'Response.Write I2_item_cd & "<P>"
'Response.Write "Save Save Save" & "<P>"
'Response.End

itxtSpread = ""
             
iCUCount = Request.Form("txtCUSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount)

For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")

'Response.Write I0_gubun & "<P>"
'Response.End

Set oPY3S117 = Server.CreateObject("PY3S117.cPMngEBomToPBomHdrMulti2")

If CheckSYSTEMError(Err,True) = True Then
	Set oPY3S117 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF					
	Response.End
End If

Call oPY3S117.P_MANAGE_EBOM_TO_PBOM_HEADER_MULTI2(gStrGlobalCollection, itxtSpread, _
				 I0_gubun, I1_plant_cd, I2_item_cd, I3_base_dt, iErrorPosition)

'If  CheckSYSTEMError2(Err,True,iErrorPosition & "행","","","","") = True Then
If CheckSYSTEMError(Err,True) = True Then
	Set oPY3S117 = Nothing					
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If

Set oPY3S117 = Nothing															'☜: Unload Comproxy

Response.Write "<Script Language=VBScript>" & vbCrLF
Response.Write "parent.DbSaveOk" & vbCrLF
Response.Write "</Script>" & vbCrLF
Response.End
%>
