<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4115mb3.asp
'*  4. Program Name         : List Resource
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2002/08/20
'*  8. Modifier (First)     : Park ,Bum Soo 
'*  9. Modifier (Last)      : Chen ,Jae Hyun
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
Call loadInfTB19029B("Q", "P", "NOCOOKIE","MB")
On Error Resume Next								'☜: 

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter 선언 
Dim strQryMode
Dim i

Const C_SHEETMAXROWS_D = 100

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey4
Dim lgStrPrevKey5

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

Dim strPlantCd
Dim strProdOrdNo
Dim strOprNo
	
	lgStrPrevKey4 = UCase(Trim(Request("lgStrPrevkey4")))
	lgStrPrevKey5 = UCase(Trim(Request("lgStrPrevkey5")))
	
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 5)
	
	UNISqlId(0) = "189201sac"
		
	IF Trim(Request("txtProdOrderNo")) = "" Then
	   strProdOrdNo = "|"
	ELSE
	   strProdOrdNo = FilterVar(UCase(Request("txtProdOrderNo")),"''","S")
	END IF
	
	IF Trim(Request("txtOprNo")) = "" Then
	   strOprNo = "|"
	ELSE
	   strOprNo = FilterVar(UCase(Request("txtOprNo")),"''","S")
	END IF
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")),"''","S")
	UNIValue(0, 2) = strProdOrdNo
	UNIValue(0, 3) = strOprNo
	
	If Not(lgStrPrevKey4 <> "" Or lgStrPrevKey5 <> "") Then
		UNIValue(0, 4) = "|"
		UNIValue(0, 5) = "|"
	Else
		UNIValue(0, 4) = "a.resource_cd >  " & FilterVar(lgStrPrevKey4, "''", "S") & " or (a.resource_cd =  " & FilterVar(lgStrPrevKey4, "''", "S")
		UNIValue(0, 5) = FilterVar(UniConvDate(lgStrPrevKey5), "''", "S")
	End If
			
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	'-------------------------------------------
	' Display Spread 3
	'-------------------------------------------
	If rs0.EOF And rs0.BOF Then
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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_CD"))%>"		
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_NM"))%>"
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("START_DT"))%>"
				strData = strData & Chr(11) & "<%=UniDateClientFormat(rs0("END_DT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("START_FLG"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("END_FLG"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("RESOURCE_GROUP_CD"))%>"		
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
			
		.lgStrPrevKey4 = "<%=ConvSPChars(Trim(rs0("RESOURCE_CD")))%>"		
		.lgStrPrevKey5 = "<%=UniDateClientFormat(rs0("START_DT"))%>"		

	End With	

</Script>	
<%
	rs0.Close
	Set rs0 = Nothing
	
	Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
