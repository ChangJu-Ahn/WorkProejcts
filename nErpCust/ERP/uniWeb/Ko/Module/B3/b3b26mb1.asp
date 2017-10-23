<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%'======================================================================================================
'*  1. Module Name          : Item Master
'*  2. Function Name        : 
'*  3. Program ID           : b3b26mb1.asp 
'*  4. Program Name         : Called By B3B26MA1 (Characteristic Query)
'*  5. Program Desc         : Lookup Characteristic Information
'*  6. Modified date(First) : 2003/02/07
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Woo Guen
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

On Error Resume Next								'��: ���⼭ ���� ������ �����Ͻ� ������ 

Call HideStatusWnd									'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2				'DBAgent Parameter ���� 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData
Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim TmpBuffer
Dim iTotalStr

Dim strCharCd
Dim strCharNm
Dim strCharValueDigit

Const C_SHEETMAXROWS_D = 50

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = Request("lgStrPrevKey1")
iLngMaxRows = Request("txtMaxRows")
	
'======================================================================================================
'	����׸�� ó�����ִ� �κ� 
'======================================================================================================

	IF Request("txtCharCd") <> "" Then

		Redim UNISqlId(0)
		Redim UNIValue(0, 0)
	
		UNISqlId(0) = "b3b28mb1a"
		
		strCharCd = Request("txtCharCd") 
		UNIValue(0, 0) = FilterVar(strCharCd, "''", "S")

		UNILock = DISCONNREAD :	UNIFlag = "1"

		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
		If rs0.EOF And rs0.BOF Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtCharNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtCharNm.value = """ & ConvSPChars(rs0(0)) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing	

	ELSE
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtCharNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	END IF
	
'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 0)
	
	UNISqlId(0) = "b3b26mb1b"
	
	IF Request("txtCharCd") = "" Then
		strCharValueCd = "|"
	ELSE
		strCharValueCd = Request("txtCharCd") 
	END IF
			
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 0) = FilterVar(strCharCd, "''", "S")
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 0) = FilterVar(iStrPrevKey, "''", "S")
	End Select

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
      
	If rs1.EOF And rs1.BOF Then
		Call DisplayMsgBox("122630", vbInformation, "", "", I_MKSCRIPT)	'��: DB �����ڵ�, �޼���Ÿ��, %ó��, ��ũ��Ʈ���� 
		rs1.Close
		Set rs1 = Nothing					
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If

Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf
	
    If Not(rs1.EOF And rs1.BOF) Then
		
		If C_SHEETMAXROWS_D < rs1.RecordCount Then 

			ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)

		Else
			
			ReDim TmpBuffer(rs1.RecordCount - 1)

		End If

		
		For iIntCnt = 0 To rs1.RecordCount - 1
			 
			If iIntCnt < C_SHEETMAXROWS_D Then
				
				strData = ""
				strData = strData & Chr(11) & ConvSPChars(rs1("CHAR_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs1("CHAR_NM"))
				strData = strData & Chr(11) & rs1("CHAR_VALUE_DIGIT")
				
		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
		
				rs1.MoveNext
				
				TmpBuffer(iIntCnt) = strData
			
			End If
			
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		
		Response.Write ".ggoSpread.Source = .frm1.vspdData1" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If rs1("CHAR_CD") = Null Then
			Response.Write ".lgStrPrevKey1 = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey1 = """ & Trim(rs1("CHAR_CD")) & """" & vbCrLf
		End If
	End If	

	rs1.Close
	Set rs1 = Nothing
	
	Response.Write ".frm1.hCharCd.value = """ & ConvSPChars(Request("txtCharCd")) & """" & vbCrLf

	Response.Write ".DbQueryOk()" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>