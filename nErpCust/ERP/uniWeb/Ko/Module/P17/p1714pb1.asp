<%@LANGUAGE = VBScript%>
<%'*******************************************************************************************
'*  1. Module Name          : 설계BOM관리 
'*  2. Function Name        :
'*  3. Program ID           : p1714pb2.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +P32118ListProdOrderHeader
'*  7. Modified date(First) : 2005-02-18
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : yjw
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
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "PB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2							'DBAgent Parameter 선언 
Dim strQryMode
Dim i

'=======================================================================================================
'	아래 선언되어 있는 변수들은 COOL:Gen 의 Record Return Count 의 제한에 따른 것이다.
'	따라서, ADO를 사용할 경우 그와같은 문제성이 없기 때문에 아래의 변수들은 사용하지 않지만 추후 
'	uniERP2000 에서 한번에 조회되는 Record Count 의 수를 30으로 제한하고 있는 만큼 그에 따른 
'	표준은 동시에 추가될 예정이므로 변수삭제는 하지 않고 그대로 놔둔다.
'=======================================================================================================
Const C_SHEETMAXROWS_D = 50

Call HideStatusWnd

Dim strReqTransNo

strQryMode = Request("lgIntFlgMode")
strReqTransNo = Trim(Request("txtReqTransNo"))

On Error Resume Next
Err.Clear																	'☜: Protect system from crashing

	'=======================================================================================================
	'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
	'=======================================================================================================
'	Redim UNISqlId(2)
'	Redim UNIValue(2, 0)
'
'	UNISqlId(0) = "180000saa"
'	UNISqlId(1) = "180000sab"
'
'	UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
'	UNIValue(1, 0) = FilterVar(UCase(Request("txtItemCd")), "''", "S")
'
'	UNILock = DISCONNREAD :	UNIFlag = "1"
'
'    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
'    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2)

	' Order Header Display
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "p1714pb1"
	UNISqlId(1) = "p1714pb11"

'	UNIValue(0, 0) = "^"
	UNIValue(0, 0) = FilterVar(strReqTransNo,"''","S")
	UNIValue(1, 0) = FilterVar(strReqTransNo,"''","S")

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

'	Response.Write "<Script Language = VBScript> " & vbCrLf
'	Response.Write "With parent " & vbCrLf
'		.txtDestPlantCd.value			= """ & ConvSPChars(rs1(PLANT_CD)) & """" & vbCrLf
'		.txtDestPlantNm.value			= """ & ConvSPChars(rs1(PLANT_NM)) & """" & vbCrLf
'		.txtBasePlantCd.value			= """ & ConvSPChars(rs1(DESIGN_PLANT_CD)) & """" & vbCrLf
'		.txtDestPlantNm.value			= """ & ConvSPChars(rs1(DESIGN_PLANT_NM)) & """" & vbCrLf
'		.txtItemCd.value				= """ & ConvSPChars(rs1(ITEM_CD)) & """" & vbCrLf
'		.txtItemNm.value				= """ & ConvSPChars(rs1(ITEM_NM)) & """" & vbCrLf
'		.txtSpec.value					= """ & ConvSPChars(rs1(SPEC)) & """" & vbCrLf
'		.txtTransDt.value				= """ & ConvSPChars(rs1(TRANS_DT)) & """" & vbCrLf
''		parent.DbQueryOk " & vbCr								'☜: 조회가 성공 
'	Response.Write "End With " & vbCrLf
'	Response.Write "</Script> " & vbCrLf

	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

	If (rs1.EOF And rs1.BOF) Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs1.Close
		Set rs1 = Nothing
		Set ADF = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If

%>
<Script Language=vbscript>
    Dim LngMaxRow
    Dim strData
    Dim TmpBuffer
    Dim iTotalStr

	With parent
		.txtDestPlantCd.value	= "<%=ConvSPChars(rs1("PLANT_CD"))%>"
		.txtDestPlantNm.value	= "<%=ConvSPChars(rs1("PLANT_NM"))%>"
		.txtBasePlantCd.value	= "<%=ConvSPChars(rs1("DESIGN_PLANT_CD"))%>"
		.txtBasePlantNm.value	= "<%=ConvSPChars(rs1("DESIGN_PLANT_NM"))%>"
		.txtItemCd.value		= "<%=ConvSPChars(rs1("ITEM_CD"))%>"
		.txtItemNm.value		= "<%=ConvSPChars(rs1("ITEM_NM"))%>"
		.txtSpec.value			= "<%=ConvSPChars(rs1("SPEC"))%>"
		.txtTransDt.value		= "<%=ConvSPChars(rs1("TRANS_DT"))%>"
	End With


    With parent												'☜: 화면 처리 ASP 를 지칭함 

 	LngMaxRow = .vspdData.MaxRows							'Save previous Maxrow

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
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("LEVEL"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CHILD_ITEM_SEQ"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CHILD_ITEM_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CHILD_ITEM_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CHILD_ITEM_SPEC"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_ACCT_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PROCUR_TYPE_NM"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("CHILD_ITEM_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CHILD_ITEM_UNIT"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("PRNT_ITEM_QTY"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PRNT_ITEM_UNIT"))%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("SAFETY_LT"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=UniNumClientFormat(rs0("LOSS_RATE"),ggQty.DecPoint,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SUPPLY_TYPE_NM"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("VALID_FROM_DT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("VALID_TO_DT"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ECN_NO"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ECN_DESC"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("REASON_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DRAWING_PATH"))%>"

				strData = strData & Chr(11) & LngMaxRow + <%=i%>
				strData = strData & Chr(11) & Chr(12)

				TmpBuffer(<%=i%>) = strData
<%
			rs0.MoveNext
			End If
		Next
%>

	iTotalStr = Join(TmpBuffer, "")

	.ggoSpread.Source = .vspdData
	.ggoSpread.SSShowDataByClip iTotalStr

'	.lgStrPrevKey = "<%=Trim(rs0("Prodt_Order_No"))%>"

<%
	End If

	rs0.Close
	Set rs0 = Nothing

	rs1.Close
	Set rs1 = Nothing

%>

	If .vspdData.MaxRows < .PopupParent.VisibleRowCnt(.vspdData,0) and .lgStrPrevKey <> "" Then	 ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
		.InitData(LngMaxRow)
		.DbQuery
	Else
		.hReqTransNo.value		= "<%=ConvSPChars(Request("txtReqTransNo"))%>"
		.DbQueryOk(LngMaxRow)
	End If

    End With
</Script>
