<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4711mb5.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/12/01
'*  7. Modified date(Last)  : 2001/12/01
'*  8. Modifier (First)     : Park, Bum Soo
'*  9. Modifier (Last)      : Park, Bum Soo 
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
'Call loadInfTB19029("P", "Q")
Call loadInfTB19029B("Q", "P", "NOCOOKIE","MB")
Call HideStatusWnd

On Error Resume Next

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag					'DBAgent Parameter 선언 
Dim	rs0
Dim strConsumedDtFrom, strConsumedDtTo, strItemCd, strWcCd, strProdtOrderNo, strResourceGroupCd

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim lgStrPrevKey1	
Dim i

Err.Clear																'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 9)

	UNISqlId(0) = "p4711mb5"

	IF Trim(Request("txtProdtOrderNo")) = "" Then
	   strProdtOrderNo = "|"
	ELSE
	   strProdtOrderNo = FilterVar(UCase(Request("txtProdtOrderNo")), "''", "S")
	END IF
		
	IF Trim(Request("txtWcCd")) = "" Then
	   strWcCd = "|"
	ELSE
	   strWcCd = FilterVar(UCase(Request("txtWcCd")), "''", "S")
	END IF
	
	IF Trim(Request("txtItemCd")) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF

	IF Trim(Request("txtConsumedDtFrom")) = "" Then
	   strConsumedDtFrom = "|"
	ELSE
	   strConsumedDtFrom = " " & FilterVar(UNIConvDate(Request("txtConsumedDtFrom")), "''", "S") & ""
	END IF
	
	IF Trim(Request("txtConsumedDtTo")) = "" Then
	   strConsumedDtTo = "|"
	ELSE
	   strConsumedDtTo = " " & FilterVar(UNIConvDate(Request("txtConsumedDtTo")), "''", "S") & ""
	END IF

	IF Trim(Request("txtResourceGroupCd")) = "" Then
	   strResourceGroupCd = "|"
	ELSE
	   strResourceGroupCd = FilterVar(UCase(Request("txtResourceGroupCd")), "''", "S")
	END IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtBatchRunNo")), "''", "S")
	UNIValue(0, 3) = strProdtOrderNo
	UNIValue(0, 4) = FilterVar(UCase(Request("txtResourceCd")), "''", "S")
	UNIValue(0, 5) = strConsumedDtFrom
	UNIValue(0, 6) = strConsumedDtTo
	UNIValue(0, 7) = strResourceGroupCd
	UNIValue(0, 8) = strItemCd	
	UNIValue(0, 9) = strWcCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData2.MaxRows										'Save previous Maxrow
	
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)	
<%  
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_no"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("opr_no"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("consumed_dt"))%>"
		strData = strData & Chr(11) & "<%=ConvToTimeFormat(rs0("consumed_time"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
		strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("prod_qty_in_order_unit"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("prodt_order_unit"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("report_type"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("job_nm"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_cd"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("wc_nm"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("rout_no"))%>"
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
		
	.lgStrPrevKey2 = "<%=Trim(rs0("prodt_order_no"))%>"
	.lgStrPrevKey3 = "<%=Trim(rs0("opr_no"))%>"
	.lgStrPrevKey4 = "<%=Trim(rs0("consumed_dt"))%>"
		
<%			
	rs0.Close
	Set rs0 = Nothing
%>
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : ConvToTimeFormat
' Description : 시간 형식으로 변경 
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
