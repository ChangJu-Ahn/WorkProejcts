<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1206mb1.asp
'*  4. Program Name         : List Manufacturing Instruction Detail
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2002/03/19
'*  7. Modified date(Last)  : 2002/11/20
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0					'DBAgent Parameter 선언 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData
Dim TmpBuffer
Dim iTotalStr

Dim strWICd
Dim strValidDt
Dim strFlag

Const C_SHEETMAXROWS_D = 50

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = Request("lgStrPrevKey")
iLngMaxRows = Request("txtMaxRows")

Redim UNISqlId(0)
Redim UNIValue(0, 3)

UNISqlId(0) = "P1206MB1"

Select Case strQryMode
	Case CStr(OPMD_CMODE)
		IF Request("txtWICd") = "" Then
			strWICd = "|"
		Else
			StrWICd = FilterVar(Request("txtWICd"), "''", "S")
		End IF	
	Case CStr(OPMD_UMODE) 
		StrWICd = FilterVar(iStrPrevKey, "''", "S")
End Select 
		
IF Request("txtValidDt") = "" Then
	strValidDt = "|"
Else
	strValidDt = " " & FilterVar(UNIConvDate(Request("txtValidDt")), "''", "S") & ""
End IF
	
UNIValue(0, 0) = "^"
UNIValue(0, 1) = strValidDt
UNIValue(0, 2) = strValidDt
UNIValue(0, 3) = strWICd
		
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	rs0.Close
	Set rs0 = Nothing
	Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "parent.frm1.txtWINm.Value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End													'☜: 비지니스 로직 처리를 종료함 
End If
	
Response.Write "<Script Language = VBScript>" & vbCrLf

If Trim(Request("txtWICd")) = Trim(rs0("MFG_INSTRUCTION_DTL_CD")) Then
	Response.Write "parent.frm1.txtWINm.Value = """ & ConvSPChars(rs0("MFG_INSTRUCTION_DTL_DESC")) & """" & vbCrLf
Else
	Response.Write "parent.frm1.txtWINm.Value = """"" & vbCrLf
End If
Response.Write "</Script>" & vbCrLf

Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf
	
	If Not(rs0.EOF And rs0.BOF) Then
		
		If C_SHEETMAXROWS_D < rs0.RecordCount Then 
			ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
		Else
			ReDim TmpBuffer(rs0.RecordCount - 1)
		End If
		
		For iIntCnt = 0 To rs0.RecordCount - 1
			If iIntCnt < C_SHEETMAXROWS_D Then 
				strData = ""
				strData = strData & Chr(11) & ConvSPChars(rs0("MFG_INSTRUCTION_DTL_CD"))			'단위작업 
				strData = strData & Chr(11) & ConvSPChars(rs0("MFG_INSTRUCTION_DTL_DESC"))			'단위작업내역 
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("VALID_START_DT"))			'유효시작일 
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("VALID_END_DT"))				'유효종료일			
				strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)

				rs0.MoveNext
				
				TmpBuffer(iIntCnt) = strData
				
			End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf

		If rs0("MFG_INSTRUCTION_DTL_CD") = Null Then
			Response.Write ".lgStrPrevKey = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey = """ & Trim(rs0("MFG_INSTRUCTION_DTL_CD")) & """" & vbCrLf
		End If
	End If

	rs0.Close
	Set rs0 = Nothing

	Response.Write ".frm1.hWICd.value = """ & ConvSPChars(Request("txtWICd")) & """" & vbCrLf
	Response.Write ".frm1.hValidDt.Value = """ & Request("txtValidDt") & """" & vbCrLf
			
	Response.Write ".DbQueryOk()" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf	

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
