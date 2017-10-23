<%@LANGUAGE = VBScript%>
<%'==============================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4711pb1.asp
'*  4. Program Name         : 자원소비이력 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2001/12/19
'*  7. Modified date(Last)  : 2002/12/11
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Ryu Sung Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'===============================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "PB")

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0		'DBAgent Parameter 선언 
Dim strQryMode
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim i

Dim strItemCd
Dim strItemAcct
Dim strWcCd
Dim strShiftCd

Call HideStatusWnd

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next
Err.Clear

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)
	
	UNISqlId(0) = "p4711pb1"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")	
	UNIValue(0, 2) = FilterVar(UCase(Request("txtBatchRunNo")), "''", "S")
	UNIValue(0, 3) = " " & FilterVar(Request("txtrdoflag"), "''", "S") & ""

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
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .vspdData.MaxRows										'Save previous Maxrow
	ReDim TmpBuffer(<%=rs0.RecordCount-1%>) 	
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("batch_run_no"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("exec_start_dt"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no_from"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no_to"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd_from"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd_to"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_cd_from"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_cd_to"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("shift_cd_from"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("shift_cd_to"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("report_dt_from"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("report_dt_to"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("status"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("string_from"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("string_to"))%>"		
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("insrt_user_id"))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData
<%		
		rs0.MoveNext
	Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .vspdData
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
