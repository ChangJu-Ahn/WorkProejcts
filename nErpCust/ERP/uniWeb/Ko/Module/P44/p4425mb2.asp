<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4425mb2.asp
'*  4. Program Name         : 오더별실적조회 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2003-02-27
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Chen, Jae Hyun
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
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
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4							'DBAgent Parameter 선언 

Dim strPlantCd
Dim strReportFromDt
Dim strReportToDt
Dim strProdtOrderNo
Dim strShiftCd

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim i

Call HideStatusWnd

On Error Resume Next
Err.Clear

	strMode = Request("txtMode")						'☜ : 현재 상태를 받음 
	strPlantCd = Request("txtPlantCd")
	strReportFromDt = Request("txtReportFromDt")
	strReportToDt = Request("txtReportToDt")
	strProdtOrderNo = Request("txtProdOrderNo")
	strShiftCd = Request("txtShiftCd")
	
	strPlantCd = FilterVar(UCase(strPlantCd), "''", "S")
	strReportFromDt = FilterVar(UniConVDate(strReportFromDt), "''", "S")
	strReportToDt = FilterVar(UniConVDate(strReportToDt), "''", "S")
	strProdtOrderNo = FilterVar(UCase(strProdtOrderNo), "''", "S")
	
	IF Trim(strShiftCd) = "" Then
	   strShiftCd = "|"
	ELSE
	   strShiftCd = FilterVar(UCase(strShiftCd), "''", "S")
	END IF
		
	Redim UNISqlId(0)
	Redim UNIValue(0, 5)

	UNISqlId(0) = "p4425mb1D"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strPlantCd
	UNIValue(0, 2) = strProdtOrderNo
	UNIValue(0, 3) = strReportFromDt
	UNIValue(0, 4) = strReportToDt 	
	UNIValue(0, 5) = strShiftCd
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT) 	
		rs0.Close
		Set rs0 = Nothing					
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow 
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent
																'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData2.MaxRows									'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%  		
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("REPORT_DT"))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PRODT_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("GOOD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("BAD_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RCPT_QTY_IN_ORDER_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("PRODT_ORDER_UNIT"))))%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PRODT_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("GOOD_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("BAD_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RCPT_QTY_IN_BASE_UNIT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("BASE_UNIT"))))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("PRODT_ORDER_NO"))))%>"	
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
	rs0.Close
	Set rs0 = Nothing
%>
	
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
