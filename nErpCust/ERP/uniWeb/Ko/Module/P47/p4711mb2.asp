<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4711mb2.asp
'*  4. Program Name         : 자원소비실적등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : +P11011ManageLotPeriod
'*  7. Modified date(First) : 2001-12-07
'*  8. Modified date(Last)  : 2001-12-07
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	
On Error Resume Next														'☜: 

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4, rs5
Dim strReturnVal

'-----------------------------------------------------------
' SQL Server, APS DB Server Information Read
'-----------------------------------------------------------
 	Err.Clear																'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(5)
	Redim UNIValue(5, 0)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000sab"
	UNISqlId(2) = "180000sab"
	UNISqlId(3) = "180000sac"
	UNISqlId(4) = "180000sac"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCdFrom")), "''", "S")
	UNIValue(2, 0) = FilterVar(UCase(Request("txtItemCdTo")), "''", "S")
	UNIValue(3, 0) = FilterVar(UCase(Request("txtWCCdFrom")), "''", "S")
	UNIValue(4, 0) = FilterVar(UCase(Request("txtWCCdTo")), "''", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4, rs5)
	Set ADF = Nothing
	%>
	<Script Language=vbscript>
		parent.frm1.txtPlantNm.value = ""
		parent.frm1.txtItemNmFrom.value = ""
		parent.frm1.txtItemNmTo.value = ""
		parent.frm1.txtWCNmFrom.value = ""
		parent.frm1.txtWCNmTo.value = ""
	</Script>	
	<%    	

	' Plant 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "	parent.frm1.txtPlantNm.value = """"" & vbcr
		Response.Write "	parent.frm1.txtPlantCd.Focus()" & vbcr
		Response.Write "</Script>" & vbcr
		Response.End    	
	Else
		Response.Write "<Script Language=vbscript>" & vbcr
		Response.Write "	parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1("PLANT_NM")) & """" & vbcr
		Response.Write "</Script>" & vbcr
	End If
	
	'redim rs1
	rs1.Close
	Set rs1 = Nothing

	' 품목명 Display
	IF Request("txtItemCdFrom") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "	parent.frm1.txtItemNmFrom.value = """"" & vbcr
			Response.Write "</Script>" & vbcr
		Else
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "	parent.frm1.txtItemNmFrom.value = """ & ConvSPChars(rs2("ITEM_NM")) & """" & vbcr
			Response.Write "</Script>	" & vbcr
		End If
	End IF
	rs2.Close
	Set rs2 = Nothing
		
	' 품목명 Display
	IF Request("txtItemCdTo") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "	parent.frm1.txtItemNmTo.value = """"" & vbcr
			Response.Write "</Script>" & vbcr
		Else
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "	parent.frm1.txtItemNmTo.value = """ & ConvSPChars(rs3("ITEM_NM")) & """" & vbcr
			Response.Write "</Script>	" & vbcr
		End If
	End IF
	rs3.Close
	Set rs3 = Nothing
		
	' 작업장명 Display
	IF Request("txtWCCdFrom") <> "" Then
		If (rs4.EOF And rs4.BOF) Then
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "	parent.frm1.txtWCNmFrom.value = """"" & vbcr
			Response.Write "</Script>" & vbcr
		Else
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "	parent.frm1.txtWCNmFrom.value = """ & ConvSPChars(rs4("WC_NM")) & """" & vbcr
			Response.Write "</Script>" & vbcr
		End If
	End IF
	rs4.Close
	Set rs4 = Nothing

	' 작업장명 Display
	IF Request("txtWCCdTo") <> "" Then
		If (rs5.EOF And rs5.BOF) Then
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "	parent.frm1.txtWCNmTo.value = """"" & vbcr
			Response.Write "</Script>"
		Else
			Response.Write "<Script Language=vbscript>" & vbcr
			Response.Write "	parent.frm1.txtWCNmTo.value = """ & ConvSPChars(rs5("WC_NM")) & """" & vbcr
			Response.Write "</Script>" & vbcr
		End If
	End IF
	rs5.Close
	Set rs5 = Nothing
	Set ADF = Nothing	

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0,11)

	UNISqlId(0) = "p4711mb2"
	
	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	UNIValue(0, 1) = FilterVar(UCase(Request("txtProdtOrderNoFrom")), "''", "S")
	UNIValue(0, 2) = FilterVar(UCase(Request("txtProdtOrderNoTo")), "''", "S")
	UNIValue(0, 3) = FilterVar(UCase(Request("txtItemCdFrom")), "''", "S")
	UNIValue(0, 4) = FilterVar(UCase(Request("txtItemCdTo")), "''", "S")
	UNIValue(0, 5) = FilterVar(UCase(Request("txtWcCdFrom")), "''", "S")
	UNIValue(0, 6) = FilterVar(UCase(Request("txtWcCdTo")), "''", "S")
	UNIValue(0, 7) = FilterVar(UCase(Request("cboShiftCdFrom")), "''", "S")
	UNIValue(0, 8) = FilterVar(UCase(Request("cboShiftCdTo")), "''", "S")
	UNIValue(0, 9) = " " & FilterVar(UNIConvDate(Request("txtReportDtFrom")), "''", "S") & ""
	UNIValue(0,10) = " " & FilterVar(UNIConvDate(Request("txtReportDtTo")), "''", "S") & ""
	UNIValue(0,11) = FilterVar(UCase(gUsrID), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	Set ADF = Nothing
	
   'If strRetMsg <> "0;Success" Then
	strReturnVal = split(strRetMsg,gColSep)
	If strReturnVal(0) <> "0" Then
		Call DisplayMsgBox(strRetMsg, vbOKOnly, "", "", I_MKSCRIPT)
	Else
		Call DisplayMsgBox(rs0("error_msg"), vbOKOnly, "", "", I_MKSCRIPT)
	End If
	
%>

<Script Language=vbscript>

parent.frm1.txtBatchRunNo.value	= "<%=ConvSPChars(rs0("batch_run_no"))%>"
parent.frm1.cboStatus.value		= "<%=ConvSPChars(rs0("status"))%>"
parent.frm1.txtSuccessCnt.value	= "<%=ConvSPChars(rs0("success_cnt"))%>"
parent.frm1.txtErrorCnt.value	= "<%=ConvSPChars(rs0("error_cnt"))%>"

<%			
	rs0.Close
	Set rs0 = Nothing
%>
	
</Script>	
