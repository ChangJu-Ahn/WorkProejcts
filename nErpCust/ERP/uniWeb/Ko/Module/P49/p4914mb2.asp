<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : p4914mb2.asp
'*  4. Program Name         : �۾��Ϻ� ��� 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2005-01-17
'*  7. Modified date(Last)  :
'*  8. Modifier (First)     : Yoon, Jeong Woo
'*  9. Modifier (Last)      :
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
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
Dim lgStrPrevKey2

'@Var_Declare

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
'strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strProdOrdNo
Dim NextFlg
Dim i

'	lgStrPrevKey2 = FilterVar(Ucase(Trim(Request("lgStrPrevKey2"))),"","SNM")

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================

	Redim UNISqlId(0)
	Redim UNIValue(0, 4)

	UNISqlId(0) = "P4913MA3"

	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtWcCd"))),"''","S")
	UNIValue(0, 2) = FilterVar(UNIConvDate(Request("txtprodDt")),"''","S")
	UNIValue(0, 3) = FilterVar(Ucase(Trim(Request("txtProdtOrderNo"))),"''","S")
	UNIValue(0, 4) = FilterVar(Ucase(Trim(Request("txtOprNo"))),"''","S")

'	If lgStrPrevKey2 = "" Then
'		UNIValue(0, 1) = ""
'		NextFlg = "N"
'	Else
'		UNIValue(0, 1) = lgStrPrevKey2
'		NextFlg = "Y"
'	End If

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

'	Response.Write UNIValue(0, 0) & "<P>"
'	Response.Write UNIValue(0, 1) & "<P>"
'	Response.Write UNIValue(0, 2) & "<P>"
'	Response.Write UNIValue(0, 3) & "<P>"
'	Response.Write strRetMsg & "<P>"

	If rs0.EOF And rs0.BOF Then
		rs0.Close
		Set rs0 = Nothing
%>

<Script Language=vbscript>
	parent.DbQuery2Ok
</Script>

<%


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
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PLANT_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("REPORT_DT"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"
			strData = strData & Chr(11) & "<%=rs0("SEQ_NO")%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WORK_TYPE"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WORK_TYPE_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WORK_MAN"))%>"
			strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("WORK_TIME"))%>"
'			msgbox "<%=rs0("WORK_TIME")%>"
'			msgbox "<%=ConvToTimeFormat(rs0("WORK_TIME"))%>"
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

'		.lgStrPrevKey2 = "<%=ConvSPChars(Trim(rs0("OPR_NO")))%>"

'		If "<%=NextFlg%>" = "N" Then
		.DbQuery2Ok
'		End If
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