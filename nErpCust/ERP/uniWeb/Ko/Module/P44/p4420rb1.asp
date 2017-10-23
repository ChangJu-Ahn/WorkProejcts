<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4420rb1.asp
'*  4. Program Name         : ��ǰ�������(�Ϻ���)
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2000/12/19
'*  7. Modified date(Last)  : 2002/12/12
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Ryu Sung Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=====================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "RB")

Dim ADF												'ActiveX Data Factory ���� �������� 
Dim strRetMsg										'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0		'DBAgent Parameter ���� 
Dim strQryMode
Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim i

Dim strItemCd
Dim strItemAcct
Dim strWcCd
Dim strShiftCd

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next
Err.Clear
	'=======================================================================================================
	'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
	'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 7)
	
	UNISqlId(0) = "189611sab"
	
	IF Request("txtItemCd") = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF
		
	IF Request("cboItemAcct") = "" Then
	   strItemAcct = "|"
	ELSE
	   strItemAcct = " " & FilterVar(Request("cboItemAcct"), "''", "S") & ""
	END IF
	
	IF Request("txtWcCd") = "" Then
	   strWcCd = "|"
	ELSE
	   strWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	END IF
	
	IF Request("txtShliftCd") = "" Then
	   strShiftCd = "|"
	ELSE
	   strShiftCd = FilterVar(UCase(Request("txtShliftCd")), "''", "S")
	END IF	
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = " " & FilterVar(Request("txtWorkDt"), "''", "S") & "" 
	UNIValue(0, 3) = " " & FilterVar(UniDateAdd("M",1,Request("txtWorkDt"),gServerDateFormat), "''", "S") & ""
	UNIValue(0, 4) = strShiftCd
	UNIValue(0, 5) = strItemCd
	UNIValue(0, 6) = strItemAcct
	UNIValue(0, 7) = strWcCd
		
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
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
	LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow

	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)			
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REPORTDT"))%>"	
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PRODQTY"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("GOODQTY"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("BADQTY"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RCPTQTY"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("UNIT"))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
			
		TmpBuffer(<%=i%>) = strData
<%		
		rs0.MoveNext
	Next
%>
	iTotalStr = Join(TmpBuffer, "")
	.ggoSpread.Source = .frm1.vspdData
	.ggoSpread.SSShowDataByClip iTotalStr
		
<%			
	rs0.Close
	Set rs0 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>
