<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4318mb3.asp
'*  4. Program Name         : 
'*  5. Program Desc         :
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2003/05/22
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Chen, Jae Hyun
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
Call LoadInfTB19029B("Q", "P","NOCOOKIE","MB")

On Error Resume Next								'☜: 

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1			'DBAgent Parameter 선언 
Dim strQryMode

Const C_SHEETMAXROWS_D = 100

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim lgStrPrevKey3	
Dim strPlantCd, strItemCd, strWcCd, strTrackingNo
Dim i

'@Var_Declare

Call HideStatusWnd

strMode = Request("txtMode")						'☜ : 현재 상태를 받음 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next

	lgStrPrevKey3 = UCase(Trim(Request("lgStrPrevKey3")))	

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 7)
	
	UNISqlId(0) = "p4318mb3"
			
	IF Trim(Request("txtPlantCd")) = "" Then
	   strPlantCd = "|"
	ELSE
	   strPlantCd = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	END IF
	
	IF Trim(Request("txtItemCd")) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF

	IF Trim(Request("txtWcCd")) = "" Then
	   strWcCd = "|"
	ELSE
	   strWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	END IF
		
	IF Trim(Request("txtTrackingNo")) = "" Then
	   strTrackingNo = "|"
	ELSE
	   strTrackingNo = FilterVar(UCase(Request("txtTrackingNo")), "''", "S")
	END IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strPlantCd
	UNIValue(0, 2) = strItemCd
	UNIValue(0, 3) = strWcCd
	UNIValue(0, 4) = strTrackingNo
	If lgStrPrevKey3 = "" Then
		UNIValue(0, 5) = "|"
	Else
		UNIValue(0, 5) = FilterVar(lgStrPrevKey3, "''", "S")
	End If

	UNIValue(0, 6) = " " & FilterVar(UniConvDate(Request("txtReqStartDt")), "''", "S") & ""
	UNIValue(0, 7) = " " & FilterVar(UniConvDate(Request("txtReqEndDt")), "''", "S") & ""
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
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
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("REQ_DT"))%>"			
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("REQ_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("RESVD_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("ISSUED_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("CONSUMED_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("REMAIN_QTY"),ggQty.DecPoint,0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TRACKING_NO"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ORDER_STATUS"))%>"
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
		
		.lgStrPrevKey3 = "<%=Trim(rs0("PRODT_ORDER_NO"))%>"
<%			
		rs0.Close
		Set rs0 = Nothing
%>

End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
