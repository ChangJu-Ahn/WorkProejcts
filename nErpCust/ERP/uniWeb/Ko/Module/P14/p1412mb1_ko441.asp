<%'======================================================================================================
'*  1. Module Name          : Purchase
'*  2. Function Name        : 
'*  3. Program ID           : P1412MB1.asp
'*  4. Program Name         : 자품목 일괄변경 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2003-03-17
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

Dim ADF												'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg										'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4, rs5					'DBAgent Parameter 선언 
Dim strQryMode

Dim strPlantCd
Dim strItemCd
Dim strBomType
	
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Dim i

Call HideStatusWnd

On Error Resume Next
Err.Clear

	strMode		= Request("txtMode")						'☜ : 현재 상태를 받음 
	strQryMode	= Request("lgIntFlgMode")
	
	strPlantCd	= Request("txtPlantCd")
	strItemCd	= Request("txtItemCd")
	strBomType	= Request("txtBomType")

	IF Trim(strPlantCd) = "" Then
	   Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
	   Response.End
	ELSE
	   strPlantCd = FilterVar(UCase(strPlantCd), "''", "S")
	END IF
	
	IF Trim(strItemCd) = "" Then
	   Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
	   Response.End
	ELSE
	   strItemCd = FilterVar(UCase(strItemCd), "''", "S")
	END IF

	IF Trim(strBomType) = "" Then
	   Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)
	   Response.End
	ELSE
	   strBomType = FilterVar(UCase(strBomType), "''", "S")
	END IF
	
	If Cint(strQryMode) = Cint(OPMD_UMODE) Then	Response.End

	Redim UNISqlId(3)
	Redim UNIValue(3, 1)

	UNISqlId(0) = "180000saa"
	UNISqlId(1) = "180000saq"
	UNISqlId(2) = "AMINORNM"
	UNISqlId(3) = "p1412mb1b"	'bom history flg

	UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
	UNIValue(1, 0) = FilterVar(Request("txtItemCd")	, "''", "S")
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
		%>
		<Script Language=vbscript>
			parent.frm1.txtPlantNm.value = ""
			parent.frm1.txtPlantCd.focus
		</Script>
		<%
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End
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
	If (rs2.EOF And rs2.BOF) Then
		Call DisplayMsgBox("122700", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
			parent.frm1.txtItemNm.value = ""
			parent.frm1.txtAcct.value = ""
			parent.frm1.txtBaseUnit.value = ""
			parent.frm1.txtSpec.value = ""
			parent.frm1.txtValidFromDt.Text = ""
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
			parent.frm1.txtItemNm.value = "<%=ConvSPChars(rs2("ITEM_NM"))%>"
			parent.frm1.txtAcct.value = "<%=ConvSPChars(rs2("ITEM_ACCT_NM"))%>"
			parent.frm1.txtBaseUnit.value = "<%=ConvSPChars(rs2("BASIC_UNIT"))%>"
			parent.frm1.txtSpec.value = "<%=ConvSPChars(rs2("SPEC"))%>"
			parent.frm1.txtValidFromDt.Text = "<%=UNIDateClientFormat(rs2("VALID_FROM_DT"))%>"
			parent.frm1.txtValidToDt.Text = "<%=UNIDateClientFormat(rs2("VALID_TO_DT"))%>"
		</Script>	
		<%
		rs2.Close
		Set rs2 = Nothing
	End If
	
	' BOM No Check
	If (rs3.EOF And rs3.BOF) Then
		Call DisplayMsgBox("182622", vbOKOnly, "", "", I_MKSCRIPT)
		%>
		<Script Language=vbscript>
			parent.frm1.txtBomType.focus
		</Script>
		<%
		rs3.Close
		Set rs3 = Nothing
		Set ADF = Nothing
		Response.End
	Else
		rs3.Close
		Set rs3 = Nothing
	End If
	
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
	
	

	' Display BOM Detail Infomation
	Redim UNISqlId(0)
	Redim UNIValue(0, 3)

	UNISqlId(0) = "p1412mb1a"
	
	IF Request("txtPlantCd") = "" Then
		strPlantCd = "|"
	Else
		strPlantCd = FilterVar(UCase(Trim(Request("txtPlantCd")))	, "''", "S")
	End IF

	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		strItemCd = FilterVar(UCase(Trim(Request("txtItemCd")))	, "''", "S")
	End IF
	
	IF Request("txtBomType") = "" Then
		strBomType = "|"
	Else
		strBomType = FilterVar(UCase(Request("txtBomType")), "''", "S")
	End IF

	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strPlantCd
	UNIValue(0, 2) = strItemCd
	UNIValue(0, 3) = strBomType
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
%>

<Script Language=vbscript>
Dim LngMaxRow 
Dim strData
Dim TmpBuffer
Dim iTotalStr

With parent																'☜: 화면 처리 ASP 를 지칭함 
	LngMaxRow = .frm1.vspdData.MaxRows									'Save previous Maxrow
  			
	ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
<%
    For i=0 to rs0.RecordCount-1 
%>
		strData = ""
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("PRNT_ITEM_CD"))))%>"	
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"	
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("SPEC"))))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("ITEM_ACCT_NM"))))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("PROCUR_TYPE_NM"))))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("CHILD_ITEM_SEQ"))))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("CHILD_ITEM_CD"))))%>"
		strData = strData & Chr(11) & "<%= UniConvNumberDBToCompany(rs0("CHILD_ITEM_QTY"), 6, 3, "", 0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("CHILD_ITEM_UNIT"))))%>"
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "<%= UniConvNumberDBToCompany(rs0("PRNT_ITEM_QTY"), 6, 3, "", 0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("PRNT_ITEM_UNIT"))))%>"
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("SAFETY_LT"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("LOSS_RATE"),ggQty.DecPoint,0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("SUPPLY_TYPE"))))%>"
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("VALID_FROM_DT"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("VALID_TO_DT"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("ECN_NO"))))%>"
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ECN_DESC"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(UCase(Trim(rs0("REASON_CD"))))%>"
		strData = strData & Chr(11) & ""
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REASON_NM"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REMARK"))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=i%>
		strData = strData & Chr(11) & Chr(12)
		
		TmpBuffer(<%=i%>) = strData		
<%		
		rs0.MoveNext
	Next
%>
		iTotalStr = Join(TmpBuffer, "")
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowDataByClip iTotalStr
<%		
		rs0.Close
		Set rs0 = Nothing
%>
	.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
	.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
	.frm1.hBomType.value		= "<%=ConvSPChars(Request("txtBomType"))%>"
	.DbQueryOk	'(LngMaxRow+1)
End With	
</Script>	
<%
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
