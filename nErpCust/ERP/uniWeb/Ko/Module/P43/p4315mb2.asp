<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Template (Save)
'*  3. Program ID           : p4315mb2
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/01/15
'*  7. Modified date(Last)  : 2002/12/17
'*  8. Modifier (First)     : Jung Yu Kyung
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "P", "NOCOOKIE","MB")
On Error Resume Next								'��: 

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0		'DBAgent Parameter ���� 

Const C_SHEETMAXROWS_D = 100

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim strQryMode
Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey2	' ���� �� 
Dim lgStrPrevKey3	' ���� �� 
Dim lgStrPrevKey4	' ���� �� 
Dim lgStrPrevKey5	' ���� �� 
Dim strTemp
Dim i
Dim strFromDt
Dim strToDt
Dim strWcCd

'@Var_Declare

Call HideStatusWnd

On Error Resume Next
	
	strQryMode = Request("lgIntFlgMode")
	
	lgStrPrevKey2 = FilterVar(UniConvDate(Request("lgStrPrevKey2")), "''", "S")
	lgStrPrevKey3 = FilterVar(Request("lgStrPrevKey3"), "''", "S")
	lgStrPrevKey4 = FilterVar(Request("lgStrPrevKey4"), "''", "S")
	lgStrPrevKey5 = FilterVar(Request("lgStrPrevKey5"),"","SNM")
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 6)
	
	strItemCd = FilterVar(Request("txtItemCd"), "''", "S")
	
	If Request("txtFromDt") = "" Then
		strFromDt = "|"
	Else 
		strFromDt = FilterVar(UniConvDate(Request("txtFromDt")), "''", "S")	
	End If
	
	If Request("txtToDt") = "" Then
		strToDt = "|"
	Else 
		strToDt = FilterVar(UniConvDate(Request("txtToDt")), "''", "S")	
	End If
	
	If Request("txtConWcCd") = "" Then
		strWcCd = "|"
	Else 
		strWcCd = FilterVar(Request("txtConWcCd"), "''", "S")	
	End If
	
	strTemp = ""
	
	If lgStrPrevKey5 <> "" Then
		strTemp = "(A.REQ_DT > " & lgStrPrevKey2 & " OR ("  
		strTemp = strTemp & "A.REQ_DT = " & lgStrPrevKey2 & " AND A.PRODT_ORDER_NO > " & lgStrPrevKey3 & ") OR ("
		strTemp = strTemp & "A.REQ_DT = " & lgStrPrevKey2 & " AND A.PRODT_ORDER_NO = " & lgStrPrevKey3 & " AND A.OPR_NO > " & lgStrPrevKey4 & ") OR ("
		strTemp = strTemp & "A.REQ_DT = " & lgStrPrevKey2 & " AND A.PRODT_ORDER_NO = " & lgStrPrevKey3 & " AND A.OPR_NO = " & lgStrPrevKey4 & " AND "
		strTemp = strTemp & "A.SEQ >= "	  & lgStrPrevKey5 & ")) "
	Else
		strTemp = "|"	
	End If
	
	UNISqlId(0) = "P4315MB2"	
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(0, 2) = FilterVar(Request("txtItemCd"), "''", "S")
	UNIValue(0, 3) = strFromDt
	UNIValue(0, 4) = strToDt
	UNIValue(0, 5) = strWcCd
	
	UNIValue(0, 6) = strTemp
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		Set rs0 = Nothing					
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
%>
<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
		
<%  
	If C_SHEETMAXROWS_D < rs0.RecordCount Then 
%>
		ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
<%
	Else
%>			
		ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
	End If

    For i=0 to rs0.RecordCount-1 
		If i < C_SHEETMAXROWS_D Then
%>
			strData = ""
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("REQ_DT"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("RESVD_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("ISSUED_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("NONISSUE_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASE_UNIT"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SEQ"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ISSUE_MTHD"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>			
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
		End If
	Next
%>
	iTotalStr = Join(TmpBuffer, "")
	.ggoSpread.Source = .frm1.vspdData2
	.ggoSpread.SSShowDataByClip iTotalStr
		
	.lgStrPrevKey2 = "<%=UniDateClientFormat(rs0("REQ_DT"))%>"
	.lgStrPrevKey3 = "<%=Trim(ConvSPChars(rs0("PRODT_ORDER_NO")))%>"
	.lgStrPrevKey4 = "<%=Trim(ConvSPChars(rs0("OPR_NO")))%>"
	.lgStrPrevKey5 = "<%=Trim(ConvSPChars(rs0("SEQ")))%>"		
<%			
	rs0.Close
	Set rs0 = Nothing
%>
	.DbDtlQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
