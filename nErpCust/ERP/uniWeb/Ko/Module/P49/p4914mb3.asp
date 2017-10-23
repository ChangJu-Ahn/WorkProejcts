<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : p4914mb3.asp
'*  4. Program Name         : �۾��Ϻ� ���
'*  5. Program Desc         :
'*  6. Modified date(First) : 2005-01-17
'*  7. Modifier (First)     : Yoon, Jeong Woo
'*  8. Modifier (Last)      :
'*  9. Comment              :
'* 10. Common Coding Guide  : this mark(��) means that "Do not change"
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
On Error Resume Next								'��:

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1			'DBAgent Parameter ���� 

Const C_SHEETMAXROWS_D = 100

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim lgStrPrevKey3
Dim i

'@Var_Declare

Call HideStatusWnd

	lgStrPrevKey3 = FilterVar(Ucase(Trim(Request("lgStrPrevKey3"))),"","SNM")

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 4)

	UNISqlId(0) = "P4913MA4"

	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtWcCd"))),"''","S")
	UNIValue(0, 2) = FilterVar(UNIConvDate(Request("txtprodDt")),"''","S")
	UNIValue(0, 3) = FilterVar(Ucase(Trim(Request("txtProdtOrderNo"))),"''","S")
	UNIValue(0, 4) = FilterVar(Ucase(Trim(Request("txtOprNo"))),"''","S")

'	If lgStrPrevKey3 = "" Then
'		UNIValue(0, 2) = 0
'	Else
'		UNIValue(0, 2) = lgStrPrevKey3
'	End If

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
'	Response.Write strRetMsg & "<P>"

	If rs0.EOF And rs0.BOF Then
		rs0.Close
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
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLANT_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("REPORT_DT"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"
			strData = strData & Chr(11) & "<%=rs0("SEQ_NO")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(rs0("ITEM_CD")))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(rs0("RESOURCE_CD")))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(rs0("RESOURCE_DESC")))%>"
			strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("ST_TIME"))%>"
			strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("END_TIME"))%>"
			strData = strData & Chr(11) & "<%=rs0("LOSS_MAN")%>"
			strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("WK_LOSS_QTY"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(rs0("WK_LOSS_CD")))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(rs0("WK_LOSS_DESC")))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(rs0("WK_LOSS_TYPE")))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(rs0("RT_DEPT_CD")))%>"
			strData = strData & Chr(11) & ""
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(rs0("RT_DEPT_NM")))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(rs0("NOTES")))%>"
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

'	.lgStrPrevKey3 = "<%=Trim(rs0("SEQ"))%>"
<%
	rs0.Close
	Set rs0 = Nothing
%>
End With
</Script>
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : ConvToTimeFormat
' Description : �ð� �������� ���� 
'==============================================================================
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec

	Dim iVal2

	iVal2 = Fix(iVal)

	If iVal2 = 0 Then
		ConvToTimeFormat = "00:00:00"
	ElseIf iVal2 > 0 Then
		iMin = Fix(iVal2 Mod 3600)
		iTime = Fix(iVal2 /3600)

		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)

		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
	Else
		iVal2 = Replace(iVal2, "-", "")
		iMin = Fix(iVal2 Mod 3600)
		iTime = Fix(iVal2 /3600)

		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		ConvToTimeFormat = "-" & ConvToTimeFormat

	End If
End Function
</script>