<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4115mb3.asp
'*  4. Program Name         : List Resource
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2002/08/20
'*  8. Modifier (First)     : Park ,Bum Soo 
'*  9. Modifier (Last)      : Chen ,Jae Hyun
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

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter ���� 
Dim strQryMode
Dim i

Const C_SHEETMAXROWS_D = 100

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey4
Dim lgStrPrevKey5

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

Dim strPlantCd
Dim strProdOrdNo
Dim strOprNo
	
	lgStrPrevKey4 = UCase(Trim(Request("lgStrPrevkey4")))
	lgStrPrevKey5 = UCase(Trim(Request("lgStrPrevkey5")))
	
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 5)
	
	UNISqlId(0) = "189201sac"
		
	IF Trim(Request("txtProdOrderNo")) = "" Then
	   strProdOrdNo = "|"
	ELSE
	   strProdOrdNo = FilterVar(UCase(Request("txtProdOrderNo")),"''","S")
	END IF
	
	IF Trim(Request("txtOprNo")) = "" Then
	   strOprNo = "|"
	ELSE
	   strOprNo = FilterVar(UCase(Request("txtOprNo")),"''","S")
	END IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")),"''","S")
	UNIValue(0, 2) = strProdOrdNo
	UNIValue(0, 3) = strOprNo
	
	If Not(lgStrPrevKey4 <> "" Or lgStrPrevKey5 <> "") Then
		UNIValue(0, 4) = "|"
		UNIValue(0, 5) = "|"
	Else
		UNIValue(0, 4) = "a.resource_cd >  " & FilterVar(lgStrPrevKey4, "''", "S") & " or (a.resource_cd =  " & FilterVar(lgStrPrevKey4, "''", "S")
		UNIValue(0, 5) = FilterVar(UniConvDate(lgStrPrevKey5), "''", "S")
	End If
			
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	'-------------------------------------------
	' Display Spread 3
	'-------------------------------------------
	If rs0.EOF And rs0.BOF Then
		rs0.Close
		Set rs0 = Nothing
		Response.End
	End If
	
%>
<Script Language=vbscript>
	Dim LngMaxRow
	Dim strData
	Dim TmpBuffer
	Dim iTotalStr
    	
	With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
		LngMaxRow = .frm1.vspdData3.MaxRows										'Save previous Maxrow
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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_CD"))%>"		
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_NM"))%>"
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("START_DT"))%>"
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("END_DT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("START_FLG"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("END_FLG"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_GROUP_CD"))%>"		
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
				
<%		
				rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData3
		.ggoSpread.SSShowDataByClip iTotalStr
			
		.lgStrPrevKey4 = "<%=ConvSPChars(Trim(rs0("RESOURCE_CD")))%>"		
		.lgStrPrevKey5 = "<%=UniDateClientFormat(rs0("START_DT"))%>"		

	End With	

</Script>	
<%
	rs0.Close
	Set rs0 = Nothing
	
	Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
