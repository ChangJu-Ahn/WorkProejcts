<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4712mb2.asp
'*  4. Program Name         : List resource consumption
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2001-12-05
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : JEON, Jaehyun
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call LoadBasisGlobalInf
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
Dim StrNextKey		' ���� �� 

Call HideStatusWnd

Dim strOprNo
Dim strProdOrderNo
Dim strFlag

Err.Clear																	'��: Protect system from crashing

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================

	' Order Header Display
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	UNISqlId(0) = "189752SAB"
	
	IF Request("txtProdOrderNo") = "" Then
		strProdOrderNo = "|"
	Else
		strProdOrderNo = FilterVar(UCase(Request("txtProdOrderNo")), "''", "S")
	End IF

	IF Request("txtOprNo") = "" Then
		strOprNo = "|"
	Else
		strOprNo = FilterVar(UCase(Request("txtOprNo")), "''", "S")
	End IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = strProdOrderNo 
	UNIValue(0, 3) = strOprNo

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set ADF = Nothing
    
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "	LngMaxRow = 0" & vbcr
		Response.Write "	Parent.DbDtl2QuerynotOk(LngMaxRow)" & vbcr
		Response.Write "</Script>" & vbcr
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)	
<%  
	If Not(rs0.EOF And rs0.BOF) Then
		For i=0 to rs0.RecordCount-1 
%>
			strData = ""
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("RESOURCE_CD"))%>")			'�ڿ� 
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("DESCRIPTION"))%>")			'�ڿ��� 
			strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("CONSUMED_DT"))%>"		'�ڿ��Һ��� 
			strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("CONSUMED_TIME"))%>"			'�ڿ��Һ� 
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("MINOR_NM"))%>")				'�ڿ�����														'�ڿ��׷�� 
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("RESOURCE_GROUP_CD"))%>")	'�ڿ��׷� 
			strData = strData & Chr(11) & Trim("<%=ConvSPChars(rs0("GROUP_NM"))%>")				'�ڿ��׷�� 
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing

%>	
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hProdOrderNo.value	= "<%=ConvSPChars(Request("txtProdOrderNo"))%>"
	.frm1.hOprNo.value			= "<%=ConvSPChars(Request("txtOprNo"))%>"
		
	.DbDtl2QueryOk(LngMaxRow+1)

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
