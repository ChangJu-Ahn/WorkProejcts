<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : �������� 
'*  2. Function Name        : 
'*  3. Program ID           : P6215mb1.asp
'*  4. Program Name         : �������˰�ȹ��� 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2005-01-25
'*  7. Modified date(Last)  :
'*  8. Modifier (First)     : Lee Sang Ho
'*  9. Modifier (Last)      : 
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

Dim ADF										'ActiveX Data Factory ���� �������� 
Dim strRetMsg								'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter ���� 
Dim rs0, rs1, rs2, rs3, rs4, rs5			'DBAgent Parameter ���� 
Dim strQryMode
'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Err.Clear																	'��: Protect system from crashing

'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================
Redim UNISqlId(0)
Redim UNIValue(0, 2)



UNISqlId(0) = "Y6215MB101"
UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtSetPlantCd"))),"''","S")
UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtCarKind"))),"''","S")
UNIValue(0, 2) = FilterVar(Ucase(Trim(Request("txtCastCd"))),"''","S")

UNILock = DISCONNREAD :	UNIFlag = "1"

Set ADF = Server.CreateObject("prjPublic.cCtlTake")

strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

Set ADF = Nothing

If  rs0.EOF And rs0.BOF  Then
	Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
    Response.Write "<Script Language=vbscript>" & vbCr
    Response.Write "</Script>"		& vbCr
    Response.end
End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
Dim lWork_Dt_Temp
Dim iData    	
With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow

<%  
	If Not(rs0.EOF And rs0.BOF) Then
		
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
				strData = strData & Chr(11) & ""																				'��: Check Field
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CAST_CD"))%>"											'��: Production Order No
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CAST_NM"))%>"												'��: Item Code
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SET_PLANT"))%>"												'��: Item Name
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SET_PLANT_NM"))%>"													'��: Spec
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CAR_KIND"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CAR_KIND_NM"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("MAKE_DT"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("INSP_PRID"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("CHECK_END_DT"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("FIN_CUR_ACCNT"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("FIN_AJ_DT"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("CUR_ACCNT"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("C_DATE"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("C_DATE"))%>"
				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=i%>) = strData
				
<%		
			rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
		
		
<%	
	End If

	rs0.Close
	Set rs0 = Nothing
	
%>	

End With

</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
