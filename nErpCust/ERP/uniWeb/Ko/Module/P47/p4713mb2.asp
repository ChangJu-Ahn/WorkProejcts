<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4713mb2.asp
'*  4. Program Name         : List Bill of Resources
'*  5. Program Desc         : 
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2002/08/22
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : Kang, Seong Moon
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
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

On Error Resume Next

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0										'DBAgent Parameter ���� 
Dim i

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Call HideStatusWnd

Dim strItemCd
Dim strOprNo
Dim strRoutNo
Dim strFlag
Dim strResourceCd

Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 5)

	UNISqlId(0) = "p4713mb2a"
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	End IF

	IF Request("txtOprNo") = "" Then
		strOprNo = "|"
	Else
		strOprNo = FilterVar(UCase(Request("txtOprNo")), "''", "S")
	End IF
	
	IF Request("txtRoutNo") = "" Then
		strRoutNo = "|"
	Else
		strRoutNo = FilterVar(UCase(Request("txtRoutNo")), "''", "S")
	End IF

	IF Request("txtResourceCd") = "" Then
		strResourceCd = "|"
	Else
		strResourceCd = FilterVar(UCase(Request("txtResourceCd")), "''", "S")
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strItemCd 
	UNIValue(0, 3) = strRoutNo
	UNIValue(0, 4) = strOprNo
	UNIValue(0, 5) = strResourceCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set ADF = Nothing
    
	If (rs0.EOF And rs0.BOF) Then
		rs0.Close
		Set rs0 = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)	
<%  
	If Not(rs0.EOF And rs0.BOF) Then
		For i=0 to rs0.RecordCount-1 
%>
			strData = ""
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("RESOURCE_CD"))%>")					'�ڿ� 
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("DESCRIPTION"))%>")					'�ڿ��� 
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("RESOURCE_TYPE"))%>")				'�ڿ�Ÿ�� 
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("RESOURCE_GROUP_CD"))%>")			'�ڿ��׷� 
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("GROUP_NM"))%>")						'�ڿ��׷�� 
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RANK"),ggQty.DecPoint,0)%>"		'���� 
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("BOR_EFFICIENCY"),ggQty.DecPoint,0)%>"	'�ڿ�ȿ�� 
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("VALID_FROM_DT"))%>"				'������ 
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("VALID_TO_DT"))%>"				'������ 

			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.h2ItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.frm1.h2OprNo.value			= "<%=ConvSPChars(Request("txtOprNo"))%>"
	.frm1.h2RoutNo.value		= "<%=ConvSPChars(Request("txtRoutNo"))%>"
	
End With

</Script>	

<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : ConvToTimeFormat
' Description : �ð� �������� ���� 
'==============================================================================
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec
				
	If IVal = 0 Then
		ConvToTimeFormat = "00:00:00"
	Else
		iMin = Fix(IVal Mod 3600)
		iTime = Fix(IVal /3600)
		
		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		 
	End If
End Function
</script>
