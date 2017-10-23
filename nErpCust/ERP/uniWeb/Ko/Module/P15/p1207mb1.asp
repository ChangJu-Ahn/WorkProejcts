<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%
'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1207mb1.asp
'*  4. Program Name         : List Standard Manufacturing Instruction
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2002/03/19
'*  7. Modified date(Last)  : 2002/11/21
'*  8. Modifier (First)     : Hong Chang Ho
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1								'DBAgent Parameter ���� 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData
Dim TmpBuffer
Dim iTotalStr

Const C_SHEETMAXROWS_D = 50

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = Request("lgStrPrevKey")
iLngMaxRows = Request("txtMaxRows")

Dim strStdInstrCd, strStdDt, strSEQ

'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
' Order Header Display
Redim UNISqlId(1)
Redim UNIValue(1, 3)

UNISqlId(0) = "P1207MB1"
UNISqlId(1) = "P1207MB2"
	
Select Case strQryMode
	Case CStr(OPMD_CMODE)
		IF Request("txtStdInstrCd") = "" Then
			strStdInstrCd = "|"
		Else
			StrStdInstrCd = FilterVar(Request("txtStdInstrCd"), "''", "S")
		End IF	

		IF Request("txtStdDt") = "" Then
			strStdDt = "|"
		Else
			StrStdDt = FilterVar(Request("txtStdDt"), "''", "S")
		End IF
		strSEQ = "|"	
	Case CStr(OPMD_UMODE) 
		StrStdInstrCd = FilterVar(Request("txtStdInstrCd"), "''", "S")
		StrSEQ = FilterVar(Request("lgStrPrevKey"), "''", "S")
		strStdDt = "|"
End Select 

UNIValue(0, 0) = "^"
UNIValue(0, 1) = strStdInstrCd
UNIValue(0, 2) = strStdDt
UNIValue(0, 3) = strStdDt

UNIValue(1, 0) = "^"
UNIValue(1, 1) = strStdInstrCd
UNIValue(1, 2) = strSEQ
UNIValue(1, 3) = "|"
		
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

If (rs0.EOF And rs0.BOF) Then
	Call DisplayMsgBox("181422", vbOKOnly, "", "", I_MKSCRIPT)
	rs0.Close
	rs1.Close
	Set rs0 = Nothing
	Set rs1 = Nothing
	Response.End													'��: �����Ͻ� ���� ó���� ������ 
End If

Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf
	
	If Not(rs0.EOF) And Not(rs0.BOF) Then
		Response.Write ".frm1.txtStdInstrCd.Value = """ & ConvSPChars(rs0("MFG_INSTRUCTION_CD")) & """" & vbCrLf		'ǥ���۾����� 
		Response.Write ".frm1.txtStdInstrNm.Value = """ & ConvSPChars(rs0("MFG_INSTRUCTION_NM")) & """" & vbCrLf		'ǥ���۾����ø� 
		Response.Write ".frm1.txtStdInstrCd1.Value = """ & ConvSPChars(rs0("MFG_INSTRUCTION_CD")) & """" & vbCrLf		'ǥ���۾����� 
		Response.Write ".frm1.txtStdInstrNm1.Value = """ & ConvSPChars(rs0("MFG_INSTRUCTION_NM")) & """" & vbCrLf		'ǥ���۾����ø� 
		Response.Write ".frm1.txtValidFromDt.Text = """ & UNIDateClientFormat(rs0("VALID_FROM_DT")) & """" & vbCrLf		'��ȿ������ 
		Response.Write ".frm1.txtValidToDt.Text = """ & UNIDateClientFormat(rs0("VALID_TO_DT")) & """" & vbCrLf			'��ȿ������ 
	End If

	If Not(rs1.EOF) And Not(rs1.BOF) Then
		
		If C_SHEETMAXROWS_D < rs1.RecordCount Then 
			ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
		Else
			ReDim TmpBuffer(rs1.RecordCount - 1)
		End If
		
		For iIntCnt = 0 To rs1.RecordCount - 1
			If iIntCnt < C_SHEETMAXROWS_D Then 
				strData = ""
				strData = strData & Chr(11) & ConvSPChars(rs1("SEQ"))						'�Ϸù�ȣ 
				strData = strData & Chr(11) & ConvSPChars(rs1("MFG_INSTRUCTION_DTL_CD"))	'�����۾����� 
				strData = strData & Chr(11) & ""											'���� 
				strData = strData & Chr(11) & ConvSPChars(rs1("MFG_INSTRUCTION_DTL_DESC"))	'�����۾��������� 
				strData = strData & Chr(11) & UNIDateClientFormat(rs1("VALID_START_DT"))	'��ȿ������ 
				strData = strData & Chr(11) & UNIDateClientFormat(rs1("VALID_END_DT"))		'��ȿ������			
		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
				rs1.MoveNext
				TmpBuffer(iIntCnt) = strData
			End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		
		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf

		If rs1("SEQ") = Null Then
			Response.Write ".lgStrPrevKey = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey = """ & Trim(rs1("SEQ")) & """" & vbCrLf
		End If

	End If
	rs0.Close
	rs1.Close
	Set rs0 = Nothing
	Set rs1 = Nothing

	Response.Write ".frm1.hStdInstrCd.value	= """ & ConvSPChars(Request("txtStdInstrCd")) & """" & vbCrLf
			
	Response.Write ".DbQueryOk(" & iLngMaxRows & " + 1)" & vbCrLf

Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>