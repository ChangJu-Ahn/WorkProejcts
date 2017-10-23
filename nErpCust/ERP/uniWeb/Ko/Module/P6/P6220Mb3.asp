<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : 금형관리 
'*  2. Function Name        : 
'*  3. Program ID           : P6220mb3.asp
'*  4. Program Name         : 금형점검계획등록 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2005-01-25
'*  7. Modified date(Last)  :
'*  8. Modifier (First)     : Lee Sang Ho
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
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4, rs5			'DBAgent Parameter 선언 
Dim strQryMode
'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim i

Const C_SHEETMAXROWS_D = 100000

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Err.Clear																	'☜: Protect system from crashing

'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
Redim UNISqlId(0)
Redim UNIValue(0, 1)

UNISqlId(0) = "Y6220MB03"
UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtCastCd"))),"''","S")
UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtWorkDt"))),"''","S")

UNILock = DISCONNREAD :	UNIFlag = "1"

Set ADF = Server.CreateObject("prjPublic.cCtlTake")

strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

Set ADF = Nothing

If (rs0.EOF And rs0.BOF) Then
	rs0.Close
	Set rs0 = Nothing
	Response.End
End If


	
%>

<Script Language=vbscript>

Dim LngMaxRow
Dim strData
Dim TmpBuffer
Dim iTotalStr
Dim lWork_Dt_Temp
Dim iData    	
With parent																	'☜: 화면 처리 ASP 를 지칭함 
	
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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SEQ"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("INSP_EMP_GB"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("INSP_EMP_GB_NM"))%>" 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("INSP_EMP_CD"))%>"
				strData = strData & Chr(11)
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("INSP_EMP_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CUST_CD"))%>"
				strData = strData & Chr(11) 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CUST_NM"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("INSP_HOUR"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("INSP_MIN"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PAYROLL"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
				strData = strData & Chr(11) & LngMaxRow + <%=cint(i)%>
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(<%=cint(i)%>) = strData
				
<%		
			rs0.MoveNext
			End If
		Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowDataByClip iTotalStr

<%	
	End If

	rs0.Close
	Set rs0 = Nothing
%>	

End With

</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
Call parent.DbDtlQueryOk2()
%>
