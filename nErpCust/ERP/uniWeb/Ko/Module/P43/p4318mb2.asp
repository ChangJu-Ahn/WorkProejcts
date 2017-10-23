<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4218mb2.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2003/05/22
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Chen, Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P","NOCOOKIE","MB")

On Error Resume Next								'��: 

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1			'DBAgent Parameter ���� 
Dim strQryMode

Const C_SHEETMAXROWS_D = 50

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim lgStrPrevKey2	
Dim i

'@Var_Declare

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strPlantCd, strItemCd, strWcCd, strTrackingNo

	lgStrPrevKey2 = UCase(Trim(Request("lgStrPrevKey2")))

	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 6)
	
	UNISqlId(0) = "p4318mb2"
		
	IF Trim(Request("txtPlantCd")) = "" Then
	   strPlantCd = "|"
	ELSE
	   strPlantCd = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & ""
	END IF
	
	IF Trim(Request("txtItemCd")) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & ""
	END IF

	IF Trim(Request("txtWcCd")) = "" Then
	   strWcCd = "|"
	ELSE
	   strWcCd = " " & FilterVar(UCase(Request("txtWcCd")), "''", "S") & ""
	END IF
		
	IF Trim(Request("txtTrackingNo")) = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = " " & FilterVar(UCase(Request("txtTrackingNo")), "''", "S") & ""
	END IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strPlantCd
	UNIValue(0, 2) = strItemCd
	If lgStrPrevKey2 = "" Then
		UNIValue(0, 3) = strWcCd
	Else
		UNIValue(0, 3) = " " & FilterVar(UCase(lgStrPrevKey2), "''", "S") & ""
	End If
	
	UNIValue(0, 4) = strTrackingNo
	UNIValue(0, 5) = " " & FilterVar(UCase(UniConvDate(Request("txtReqStartDt"))), "''", "S") & ""
	UNIValue(0, 6) = " " & FilterVar(UCase(UniConvDate(Request("txtReqEndDt"))), "''", "S") & ""

	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

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
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("WC_NM"))%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("REQ_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RESVD_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("ISSUED_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("CONSUMED_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("REMAIN_QTY"),ggQty.DecPoint,0)%>"
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
		
		.lgStrPrevKey2 = "<%=Trim(rs0("WC_CD"))%>"
		
		.DbQuery2Ok
<%			
		rs0.Close
		Set rs0 = Nothing
%>
End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
