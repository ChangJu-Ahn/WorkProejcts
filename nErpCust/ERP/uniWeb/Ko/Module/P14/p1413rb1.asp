<%@LANGUAGE = VBScript%>
<%'*******************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1413rb1.asp
'*  4. Program Name         : BOM Mass Replacement
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/03/14
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : RYU SUNG WON
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'********************************************************************************************%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "P", "NOCOOKIE", "RB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4					'DBAgent Parameter 선언 
Dim strQryMode

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Call HideStatusWnd

Dim strPlantCd
Dim strItemCd
Dim strBomType

On Error Resume Next
Err.Clear																	'☜: Protect system from crashing

	strQryMode = Request("lgIntFlgMode")
	
	Redim UNISqlId(3)
	Redim UNIValue(3, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000saq"
	UNISqlId(2) = "AMINORNM"
	UNISqlId(3) = "p1412mb1b"	'bom history flg

	UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(1, 0) = FilterVar(Request("txtItemCd"), "''", "S")
	UNIValue(1, 1) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(2, 0) = FilterVar("P1401","''","S")
	UNIValue(2, 1) = FilterVar(Request("txtBomType"),"''","S")
	UNIValue(3, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2, rs3, rs4)

	' Plant 명 Display      
	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		rs1.Close
		Set rs1 = Nothing
	End If

	' 품목명 Display
	IF Request("txtItemCd") <> "" Then
		If (rs2.EOF And rs2.BOF) Then
			Call DisplayMsgBox("122700", vbOKOnly, "", "", I_MKSCRIPT)
			%>
			<Script Language=vbscript>
				parent.frm1.hItemCd.value = ""
				parent.frm1.txtItemNm.value = ""
				parent.frm1.txtAcct.value = ""
				parent.frm1.txtProcurType.value = ""
				parent.frm1.txtSpec.value = ""
				'parent.frm1.txtValidFromDt.Text = ""
				parent.frm1.txtValidToDt.Text = ""
				parent.frm1.txtItemCd.focus
			</Script>
			<%
			rs2.Close
			Set rs2 = Nothing
			Set ADF = Nothing
			Response.End
		Else
			%>
			<Script Language=vbscript>
				parent.frm1.hItemCd.value = "<%=ConvSPChars(rs2("ITEM_CD"))%>"
				parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
				parent.frm1.txtAcct.value = "<%=ConvSPChars(rs2("ITEM_ACCT_NM"))%>"
				parent.frm1.txtProcurType.value = "<%=ConvSPChars(rs2("PROCUR_TYPE_NM"))%>"
				parent.frm1.txtSpec.value = "<%=ConvSPChars(rs2("SPEC"))%>"
				'parent.frm1.txtValidFromDt.Text = "<%=UniDateClientFormat(rs2("PLANT_VALID_FROM_DT"))%>"
				parent.frm1.txtValidToDt.Text = "<%=UniDateClientFormat(rs2("PLANT_VALID_TO_DT"))%>"	'2003-09-13
			</Script>	
			<%
			rs2.Close
			Set rs2 = Nothing
		End If
	End IF
	
	' BOM No Check
	IF Request("txtBomType") <> "" Then
		If (rs3.EOF And rs3.BOF) Then
			Call DisplayMsgBox("182622", vbOKOnly, "", "", I_MKSCRIPT)
			rs3.Close
			Set rs3 = Nothing
			Set ADF = Nothing
			Response.End
		Else
			rs3.Close
			Set rs3 = Nothing
		End If
	End IF

	' BOM HISTORY FLG - P_Plant_Configuration
	If (rs4.EOF And rs4.BOF) Then
		%>
		<Script Language=vbscript>
			parent.frm1.hBomHistoryFlg.value = "N"
		</Script>
		<%
		rs4.Close
		Set rs4 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		%>
		<Script Language=vbscript>
			parent.frm1.hBomHistoryFlg.value = "<%=Trim(rs4("BOM_HISTORY_FLG"))%>"
		</Script>
		<%
		rs4.Close
		Set rs4 = Nothing
	End If
	
	Set ADF = Nothing
	
	Response.Write "<Script Language=vbscript>" & vbCrLf
	Response.Write "parent.DbQueryOk()" & vbCrLf
	Response.Write "</Script>" & vbCrLf

%>
