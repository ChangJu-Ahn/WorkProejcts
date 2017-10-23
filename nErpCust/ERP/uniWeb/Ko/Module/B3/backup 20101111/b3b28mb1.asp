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
'*  3. Program ID           : b3b28mb1.asp 
'*  4. Program Name         : Called By B3B28MA1 (Characteristic Value Query)
'*  5. Program Desc         : Lookup Characteristic Value Information
'*  6. Modified date(First) : 2003/02/06
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Woo Guen
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================


On Error Resume Next								'��: 

Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "*", "NOCOOKIE", "MB")

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2				'DBAgent Parameter ���� 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData

Dim strCharCd
Dim strCharValueCd

Dim strAvailableItem

Dim TmpBuffer
Dim iTotalStr

Const C_SHEETMAXROWS_D = 100

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = Request("lgStrPrevKey1")
iLngMaxRows = Request("txtMaxRows")

'======================================================================================================
'	����׸�� ó�����ִ� �κ� 
'======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 0)
	
	UNISqlId(0) = "b3b28mb1a"
	
	IF Request("txtCharCd") = "" Then
		UNIValue(0, 0) = "|"
	ELSE
		strCharCd = Request("txtCharCd") 
		UNIValue(0, 0) = FilterVar(strCharCd, "''", "S")
	END IF

	UNILock = DISCONNREAD :	UNIFlag = "1"

	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("122630", vbInformation, "", "", I_MKSCRIPT)	'��: DB �����ڵ�, �޼���Ÿ��, %ó��, ��ũ��Ʈ���� 
		rs0.Close
		Set rs0 = Nothing					
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtCharNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End													'��: �����Ͻ� ���� ó���� ������	
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtCharNm.value = """ & ConvSPChars(rs0(0)) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If
	
	rs0.Close
	Set rs0 = Nothing
	Set ADF = Nothing
	
'======================================================================================================
'	��簪�� ó�����ִ� �κ� 
'======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)
	
	UNISqlId(0) = "b3b28mb1b"	
	
	IF Request("txtCharCd") = "" Then
		UNIValue(0, 0) = "|"
	ELSE
		strCharCd = Request("txtCharCd") 
		UNIValue(0, 0) = FilterVar(strCharCd, "''", "S")
	END IF
	
	IF Request("txtCharValueCd") <> "" Then
		strCharValueCd = Request("txtCharValueCd") 
		UNIValue(0, 1) = FilterVar(strCharValueCd, "''", "S")
			
		UNILock = DISCONNREAD :	UNIFlag = "1"

		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)

		If rs1.EOF And rs1.BOF Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtCharValueNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtCharValueNm.value = """ & ConvSPChars(rs1(0)) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If	
	
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
	ELSE
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtCharValueNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	END IF
			

'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)
	
	UNISqlId(0) = "b3b28mb1c"
	
	IF Request("txtCharCd") = "" Then
		strCharCd = "|"
	ELSE
		strCharCd = Request("txtCharCd") 
	END IF
	
	IF Request("txtCharValueCd") = "" Then
		strCharValueCd = "|"
	ELSE
		strCharValueCd = Request("txtCharValueCd") 
	END IF

	UNIValue(0, 0) = FilterVar(strCharCd, "''", "S")	
			
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 1) = FilterVar(strCharValueCd, "''", "S")
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 1) = FilterVar(iStrPrevKey, "''", "S")
	End Select
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs2)
      
	If rs2.EOF And rs2.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'��: DB �����ڵ�, �޼���Ÿ��, %ó��, ��ũ��Ʈ���� 
		rs2.Close
		Set rs2 = Nothing					
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
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
				strData = strData & Chr(11) & ConvSPChars(rs2("CHAR_VALUE_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs2("CHAR_VALUE_NM"))

		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
			
				rs2.MoveNext
				
				TmpBuffer(iIntCnt) = strData
				
			End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")

		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If rs2("CHAR_VALUE_CD") = Null Then
			Response.Write ".lgStrPrevKey1 = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey1 = """ & Trim(rs2("CHAR_VALUE_CD")) & """" & vbCrLf
		End If
	End If	

	rs2.Close
	Set rs2 = Nothing
	
	Response.Write ".frm1.hCharCd.value = """ & ConvSPChars(Request("txtCharCd")) & """" & vbCrLf
	Response.Write ".frm1.hCharValueCd.value = """ & ConvSPChars(Request("txtCharValueCd")) & """" & vbCrLf

	Response.Write ".DbQueryOk()" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>