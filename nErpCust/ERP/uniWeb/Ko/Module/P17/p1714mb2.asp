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
Dim oPY3S116									'☆ : 입력/수정용 ComProxy Dll 사용 변수 

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

Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3, rs4, rs5, rs6
Dim ADF
Dim strRetMsg
Dim arrRowVal
Dim arrColVal
Dim Count

Err.Clear										'☜: Protect system from crashing

I0_gubun    = Trim(Request("hgubun"))
I1_plant_cd = Trim(UCase(Request("txtDestPlantCd")))
I2_item_cd  = Trim(UCase(Request("txtItemCd")))
I3_base_dt  = Trim(Request("hStartDate"))

itxtSpread = ""
             
iCUCount = Request.Form("txtCUSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount)

For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")

Set oPY3S116 = Server.CreateObject("PY3S116.cPMngEBomToPBomHdrMulti")

If CheckSYSTEMError(Err,True) = True Then
	Set oPY3S116 = Nothing
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF					
	Response.End
End If

Call oPY3S116.P_MANAGE_EBOM_TO_PBOM_HEADER_MULTI(gStrGlobalCollection, itxtSpread, _
				 I0_gubun, I1_plant_cd, I2_item_cd, I3_base_dt, iErrorPosition)


If CheckSYSTEMError(Err,True) = True Then
	Set oPY3S116 = Nothing					
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
	Response.End
End If


Set oPY3S116 = Nothing

If itxtSpread <> "" Then
	arrRowVal = Split(itxtSpread, gRowSep)
	
	For Count = 0 To UBound(arrRowVal) - 1
		arrColVal = Split(arrRowVal(Count), gColSep)

		'=======================================================================================================
		'	Error List Check : 에러 유무 확인 
		'=======================================================================================================
		Redim UNISqlId(0)
		Redim UNIValue(0, 3)

		UNISqlId(0) = "P1714MA3"
		UNIValue(0, 0) = FilterVar(Ucase(arrColVal(2)),"''","S")
		UNIValue(0, 1) = FilterVar(Ucase(arrColVal(4)),"''","S")
		UNIValue(0, 2) = FilterVar(Ucase(arrColVal(0)),"''","S")
		UNIValue(0, 3) = FilterVar(Ucase(arrColVal(1)),"''","S")

		UNILock = DISCONNREAD :	UNIFlag = "1"

		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6)
		Set ADF = Nothing

		If (rs0.EOF And rs0.BOF) Then		'// Error 없을시 

			rs0.Close
			Set rs0 = Nothing
			Response.Write "<Script Language=VBScript>" & vbCrLF
			Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
			Response.Write "parent.DbCheckOk" & vbCrLF
			Response.Write "</Script>" & vbCrLF
			Response.End

		Else		'// Error 발생시 

			rs0.Close
			Set rs0 = Nothing
			Response.Write "<Script Language=VBScript>" & vbCrLF
			Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
			Response.Write "parent.DbErrorPrcOk" & vbCrLF
			Response.Write "</Script>" & vbCrLF
			Response.End

		End If

	Next
	
End If

%>
