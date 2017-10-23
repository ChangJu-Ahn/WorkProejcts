<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1104mb1.asp
'*  4. Program Name         : Shift Query
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2002/11/21
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2							'DBAgent Parameter 선언 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey1, iStrPrevKey2
Dim strData
Dim TmpBuffer
Dim iTotalStr

Const C_SHEETMAXROWS_D = 50

strQryMode = Request("lgIntFlgMode")
iLngMaxRows = Request("txtMaxRows")
iStrPrevKey1 = Trim(UCase(Request("lgStrPrevKeyIndex")))
iStrPrevKey2 = Trim(UCase(Request("lgStrPrevKeyIndex1")))

'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
Redim UNISqlId(0)
Redim UNIValue(0, 0)

UNISqlId(0) = "180000saa"

UNIValue(0, 0) =FilterVar(UCase(Request("txtPlantCd")), "''", "S")

UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
Response.Write "</Script>" & vbCrLf

' Plant 명 Display      
If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantCd.Focus()" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
	Response.End
Else
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs0("PLANT_NM")) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	rs0.Close
	Set rs0 = Nothing
End If

' Routing Header Display
If strQryMode = CStr(OPMD_CMODE) Then

	Redim UNISqlId(0)
	Redim UNIValue(0, 1)

	UNISqlId(0) = "p1104mb1h"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = FilterVar(UCase(Request("txtShiftCd")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
			  
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("180400", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtShiftNm1.value = """"" & vbCrLf
			Response.Write "parent.frm1.txtShiftCd1.Focus()" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "With parent" & vbCrLf
			Response.Write ".frm1.txtShiftNm1.value = """ & ConvSPChars(rs1("Description")) & """" & vbCrLf
			Response.Write ".frm1.txtShiftCd2.value = """ & Request("txtShiftCd") & """" & vbCrLf
			Response.Write ".frm1.txtShiftNm2.value = """ & ConvSPChars(rs1("Description")) & """" & vbCrLf
			Response.Write ".frm1.txtValidFromDt.text = """ & UniDateClientFormat(rs1("Valid_From_Dt")) & """" & vbCrLf
			Response.Write ".frm1.txtValidToDt.text = """ & UniDateClientFormat(rs1("Valid_To_Dt")) & """" & vbCrLf
		Response.Write "End With" & vbCrLf
		Response.Write "</Script>" & vbCrLf	
		rs1.Close
		Set rs1 = Nothing
	End If

End If
	
Redim UNISqlId(0)
Redim UNIValue(0, 4)

UNISqlId(0) = "p1104mb1d"
	
UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
UNIValue(0, 1) = FilterVar(UCase(Request("txtShiftCd")), "''", "S")
UNIValue(0, 2) = FilterVar(iStrPrevKey1, "''", "S")
UNIValue(0, 3) = FilterVar(iStrPrevKey1, "''", "S")
UNIValue(0, 4) = FilterVar(iStrPrevKey2, "''", "S")

UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs2)

If (rs2.EOF And rs2.BOF) Then
	Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "parent.frm1.hPlantCd.Value = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCrLf
		Response.Write "parent.frm1.hShiftCd.Value = """ & ConvSPChars(Request("txtShiftCd")) & """" & vbCrLf
		Response.Write "parent.DbQueryOk(0)" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Call DisplayMsgBox("180500", vbOKOnly, "", "", I_MKSCRIPT)
	rs2.Close
	Set rs2 = Nothing
	Set ADF = Nothing
	Response.End
End If

Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf
	
	If Not(rs2.EOF And rs2.BOF) Then
	
		If C_SHEETMAXROWS_D < rs2.RecordCount Then 
			ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
		Else
			ReDim TmpBuffer(rs2.RecordCount - 1)
		End If

		For iIntCnt = 0 To rs2.RecordCount - 1
			If iIntCnt < C_SHEETMAXROWS_D Then 
				strData = ""
				strData = strData & Chr(11) & ""			
				strData = strData & Chr(11) & rs2("START_TIME")
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & rs2("END_TIME")
				strData = strData & Chr(11) & rs2("OVER_RUN_FLG")
				strData = strData & Chr(11) & rs2("MUST_COMPLETE")
				strData = strData & Chr(11) & rs2("START_DAY")
				strData = strData & Chr(11) & rs2("END_DAY")
		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(iIntCnt) = strData
				rs2.MoveNext
			End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		
		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If rs2("START_DAY") = NULL Then
			Response.Write ".lgStrPrevKeyIndex = """ & 0 & """" & vbCrLf
			Response.Write ".lgStrPrevKeyIndex1 = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKeyIndex = """ & rs2("START_DAY") & """" & vbCrLf
			Response.Write ".lgStrPrevKeyIndex1 = """ & ConvSPChars(Trim(rs2("START_TIME"))) & """" & vbCrLf
		End If
	End If

	rs2.Close
	Set rs2 = Nothing

	Response.Write "If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKeyIndex <> 0 and.lgStrPrevKeyIndex1 <> """" Then" & vbCrLf
		Response.Write ".DbQuery" & vbCrLf
	Response.Write "Else" & vbCrLf
		Response.Write ".frm1.hPlantCd.value = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCrLf
		Response.Write ".frm1.hShiftCd.value = """ & ConvSPChars(Request("txtShiftCd")) & """" & vbCrLf
			
		Response.Write ".DbQueryOk(" & iLngMaxRows & " + 1)" & vbCrLf
	Response.Write "End If" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
