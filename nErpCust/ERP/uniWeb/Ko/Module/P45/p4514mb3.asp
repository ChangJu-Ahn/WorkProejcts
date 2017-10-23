<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4513mb3.asp
'*  4. Program Name         : 입고대기현황조회 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/11/25
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "*", "NOCOOKIE", "MB")
On Error Resume Next								'☜: 

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1			'DBAgent Parameter 선언 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim lgStrPrevKey3	
Dim i
Dim strPlantCd, strItemCd, strWcCd, strProdtOrderNo, strSlCd, strTrackingNo

'@Var_Declare

Call HideStatusWnd

	lgStrPrevKey3 = UCase(Trim(Request("lgStrPrevKey3")))	

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 6)
	
	UNISqlId(0) = "p4514mb3"
			
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
		
	IF Trim(Request("txtProdtOrderNo")) = "" Then
	   strProdtOrderNo = "|"
	ELSE
	   strProdtOrderNo = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	END IF
		
	IF Trim(Request("txtSlCd")) = "" Then
	   strSlCd = "|"
	ELSE
	   strSlCd = FilterVar(UCase(Request("txtSlCd")), "''", "S")
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
	UNIValue(0, 4) = strProdtOrderNo
	UNIValue(0, 5) = strSlCd
	UNIValue(0, 6) = strTrackingNo
	
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
		
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)	
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRODT_ORDER_NO"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("OPR_NO"))%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("waitqtyinorderunit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_unit"))%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prodqtyinorderunit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("goodqtyinorderunit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("badqtyinorderunit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("rcptqtyinorderunit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("waitqtyinbaseunit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("base_unit"))%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prodqtyinbaseunit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("goodqtyinbaseunit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("badqtyinbaseunit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("rcptqtyinbaseunit"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"						
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sl_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("sl_nm"))%>"			
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData
<%		
		rs0.MoveNext
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
