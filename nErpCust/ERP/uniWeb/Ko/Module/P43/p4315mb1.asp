<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4315mb1
'*  4. Program Name         : Query Component Reservation
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/01/15
'*  7. Modified date(Last)  : 2002/12/17
'*  8. Modifier (First)     : Jung Yu Kyung
'*  9. Modifier (Last)      : Chen Jae Hyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
On Error Resume Next								'☜: 

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4

Const C_SHEETMAXROWS_D = 100

Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strQryMode

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim strFlag
Dim strFromDt
Dim strToDt
Dim strProdOrderNo
Dim strPlantCd
Dim strItemCd1
Dim strItemCd2
Dim strConWcCd
Dim strItemAcct
Dim i

Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(3)
	Redim UNIValue(3, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sab"
	UNISqlId(3) = "180000sac"
	
	UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(1, 0) = FilterVar(Request("txtItemCd1"), "''", "S")
	UNIValue(2, 0) = FilterVar(Request("txtItemCd2"),"" & FilterVar("zzzzzzzzzzzzzzzzzz", "''", "S") & "","S")
	UNIValue(3, 0) = FilterVar(Request("txtConWcCd"), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)
	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtItemNm.value = ""
		parent.frm1.txtItemNm2.value = ""
		parent.frm1.txtConWcNm.value = ""
	</Script>	
	<%
	' Plant 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		rs1.Close
		Set rs1 = Nothing
		strFlag = "ERROR_PLANT"
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		</Script>	
		<%    	
		rs1.Close
		Set rs1 = Nothing
	End If

	' 품목명 Display
	IF Trim(Request("txtItemCd1")) <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			rs2.Close
			Set rs2 = Nothing
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			</Script>	
			<%
			rs2.Close
			Set rs2 = Nothing
		End If
	End IF

	' 품목명 Display
	IF Trim(Request("txtItemCd2")) <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			rs3.Close
			Set rs3 = Nothing
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtItemNm2.value = "<%=ConvSPChars(rs3("ITEM_NM"))%>"
			</Script>	
			<%
			rs3.Close
			Set rs3 = Nothing
		End If
	End IF
		
	' 작업장명 Display
	IF Trim(Request("txtConWcCd")) <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			rs4.Close
			Set rs4 = Nothing
			If strFlag <> "ERROR_PLANT" Then
				strFlag = "ERROR_WCCD"
			End If	
			
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.txtConWcNm.value = "<%=ConvSPChars(rs4("WC_NM"))%>"
			</Script>	
			<%
			rs4.Close
			Set rs4 = Nothing
		End If
	End IF
	
	If strFlag <> "" Then	
		If strFlag = "ERROR_PLANT" Then
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
				parent.frm1.txtPlantCd.Focus()
			</Script>	
			<%
			Response.End	
		End If
		If strFlag = "ERROR_WCCD" Then
			Call DisplayMsgBox("182100", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
			parent.frm1.txtConWcCd.Focus()
			</Script>	
			<%
			Response.End
		End If
	End IF
	
	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 7)
	
	strPlantCd	= FilterVar(Request("txtPlantCd"), "''", "S")
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			If Request("txtItemCd1") = "" Then
				strItemCd1	= "|"
			Else
				strItemCd1	= FilterVar(Request("txtItemCd1"), "''", "S")
			End If
		Case CStr(OPMD_UMODE) 
			
			strItemCd1	= FilterVar(Request("lgStrPrevKey1"), "''", "S")
			
	End Select
	
	If Request("txtItemCd2") = "" Then
		strItemCd2	= "|"
	Else
		strItemCd2	= FilterVar(Request("txtItemCd2"),"" & FilterVar("zzzzzzzzzzzzzzzzzz", "''", "S") & "","S")
	End If
	
	If Request("txtItemAcct") = "" Then
		strItemAcct	= "|"
	Else
		strItemAcct	= FilterVar(Request("txtItemAcct"), "''", "S")
	End If
		 
    If Request("txtConWcCd") = "" Then
		strConWcCd	= "|"
	Else
		strConWcCd	= FilterVar(Request("txtConWcCd"), "''", "S")
	End If
    
	If Request("txtFromDt") = "" Then
		strFromDt	= "|"
	Else
		strFromDt	= FilterVar(UniConvDate(Request("txtFromDt")), "''", "S")
	End If
	
	If Request("txtToDt") = "" Then
		strToDt	= "|"
	Else
		strToDt	= FilterVar(UniConvDate(Request("txtToDt")), "''", "S")
	End If
	
	UNISqlId(0) = "P4315MB1"
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strPlantCd
	UNIValue(0, 2) = strItemCd1
	UNIValue(0, 3) = strItemCd2
	UNIValue(0, 4) = strFromDt
	UNIValue(0, 5) = strToDt
	UNIValue(0, 6) = strItemAcct
	UNIValue(0, 7) = strConWcCd

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
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
	LngMaxRow = .frm1.vspdData1.MaxRows										'Save previous Maxrow		
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

    For i = 0 to rs0.RecordCount-1
		If i < C_SHEETMAXROWS_D Then
%>
			strData = ""
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_cd"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_nm"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("spec"))%>"
			strData = strData & Chr(11) & LngMaxRow + <%=i%>
			strData = strData & Chr(11) & Chr(12)
			
			TmpBuffer(<%=i%>) = strData
<%		
			rs0.MoveNext
		End If
	Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData1
		.ggoSpread.SSShowDataByClip iTotalStr
		
		
		.lgStrPrevKey1 = "<%=Trim(rs0("item_cd"))%>"
		
		.frm1.hPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hConWcCd.value = "<%=ConvSPChars(Request("txtConWcCd"))%>"
		.frm1.hFromDate.value = "<%=Request("txtFromDt")%>"
		.frm1.hToDate.value = "<%=Request("txtToDt")%>"
		.frm1.hItemCd1.value = "<%=Request("txtItemCd1")%>"
		.frm1.hItemCd2.value = "<%=Request("txtItemCd2")%>"
		.frm1.hItemAcct.value = "<%=Request("txtItemAcct")%>"		
<%			
		rs0.Close
		Set rs0 = Nothing
%>
	.DbQueryOk
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
