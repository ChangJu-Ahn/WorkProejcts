<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4213mb3.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : ?
'*  7. Modified date(Last)  : 2002/08/21
'*  8. Modifier (First)     : 2003/05/21
'*  9. Modifier (Last)      : Chen, Jae Hyun
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
On Error Resume Next								'☜: 

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1			'DBAgent Parameter 선언 

Const C_SHEETMAXROWS_D = 100

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim lgStrPrevKey3	
Dim i

'@Var_Declare

Call HideStatusWnd

	lgStrPrevKey3 = FilterVar(UCase(Trim(Request("lgStrPrevKey3"))),"''","S")

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)
	
	UNISqlId(0) = "189200sac"
			
	UNIValue(0, 0) = FilterVar(UCase(Trim(Request("txtProdOrderNo"))),"''","S")
	UNIValue(0, 1) = FilterVar(UCase(Trim(Request("txtOprNo"))),"''","S")
		
	If lgStrPrevKey3 = "" Then
		UNIValue(0, 2) = 0
	Else
		UNIValue(0, 2) = lgStrPrevKey3
	End If
	
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
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CHILD_ITEM_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("Spec"))%>"
			strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("REQ_DT"))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("REQ_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(UCase(rs0("BASE_UNIT")))%>"
			strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("ISSUED_QTY"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)%>"		
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
		
	.lgStrPrevKey3 = "<%=Trim(rs0("SEQ"))%>"	
<%			
	rs0.Close
	Set rs0 = Nothing
%>
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
