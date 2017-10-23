<%@ LANGUAGE=VBSCript%>
<%Option Explicit	%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call HideStatusWnd
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

'**********************************************************************************************
'*  1. Module Name		  : Procurement
'*  2. Function Name		:
'*  3. Program ID		   : m2111mb201
'*  4. Program Name		 : 업체지정 
'*  5. Program Desc		 :
'*  6. Comproxy List		:
'
'*  7. Modified date(First) : 2003/01/14
'*  8. Modified date(Last)  : 2003/03/03
'*  9. Modifier (First)	 : Oh Chang Won
'* 10. Modifier (Last)	  : Kim Jin Ha
'* 11. Comment			  :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*							this mark(⊙) Means that "may  change"
'*							this mark(☆) Means that "must change"
'* 13. History			  :
'* 14. Business Logic of m2111ma2(업체지정)
'**********************************************************************************************
	Dim lgOpModeCRUD

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3				'☜ : DBAgent Parameter 선언 
	Dim istrData
	Dim iTotstrData
	Dim iStrPrNo
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim index	 ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
	Dim lgTailList												'☜ : Orderby절에 사용될 field 리스트 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	Dim sRow
	Dim lglngHiddenRows
	DIM MaxRow2
	Dim MaxCount
	Dim intARows
	Dim intTRows
	intARows=0
	intTRows=0

	Dim iStrSpplNm
	Dim iStrPlantNm
	Dim istrPoNo
	Dim strTrackNo

	Const C_SHEETMAXROWS_D  = 100

	On Error Resume Next															 '☜: Protect system from crashing
	Err.Clear																		'☜: Clear Error status

	lgOpModeCRUD  = Request("txtMode")

	Select Case lgOpModeCRUD
		Case CStr(UID_M0001)
			 Call  SubBizQueryMulti()
		Case CStr(UID_M0002)														 '☜: Save,Update
			 Call SubBizSaveMulti()
	End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	On Error Resume Next
	Err.Clear

	lgPageNo	   = UNICInt(Trim(Request("lgPageNo")),0)	'☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist	= "No"
	iLngMaxRow	 = CLng(Request("txtMaxRows"))

	Call FixUNISQLData()
	Call QueryData()

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Dim strVal
	Redim UNISqlId(3)													 '☜: SQL ID 저장을 위한 영역확보 
	Redim UNIValue(3,6)												 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
																		'	parameter의 수에 따라 변경함 
	UNISqlId(0) = "M3111MA601" 											' header
	UNISqlId(1) = "M3111QA102"											  '공급처명 
	UNISqlId(2) = "M2111QA302"											  '공장명 
	UNISqlId(3) = "M3111MA603"											  '발주번호 

	UNIValue(0,0) = "^"

	'공급처 
	If Trim(Request("txtSupplierCd")) <> "" Then
		UNIValue(0,1) = "  " & FilterVar(UCase(Request("txtSupplierCd")), "''", "S") & "  "
		UNIValue(1,0) = "  " & FilterVar(UCase(Request("txtSupplierCd")), "''", "S") & "  "
	Else
		UNIValue(0,1) = "|"
		UNIValue(1,0) = "  " & FilterVar(UCase(Request("txtSupplierCd")), "''", "S") & "  "
	End If

	'발주일 
	If Trim(Request("txtPrFrDt")) <> "" Then
		UNIValue(0,2) = "  " & FilterVar(UNIConvDate(Request("txtPrFrDt")), "''", "S") & " "
	Else
		UNIValue(0,2) = "" & FilterVar("1900-01-01", "''", "S") & ""
	End If

	'발주일 
	If Trim(Request("txtPrToDt")) <> "" Then
		UNIValue(0,3) = "  " & FilterVar(UNIConvDate(Request("txtPrToDt")), "''", "S") & " "
	Else
		UNIValue(0,3) = "" & FilterVar("2999-12-31", "''", "S") & ""
	End If

	'발주번호 
	If Trim(Request("txtPoNo")) <> "" Then
		UNIValue(0,4) = "  " & FilterVar(UCase(Request("txtPoNo")), "''", "S") & "  "
		UNIValue(3,0) = "  " & FilterVar(UCase(Request("txtPoNo")), "''", "S") & "  "
	Else
		UNIValue(0,4) = "|"
		UNIValue(3,0) = "  " & FilterVar(UCase(Request("txtPoNo")), "''", "S") & "  "
	End If

	'공장 
	If Trim(Request("txtPlantCd")) <> "" Then
		UNIValue(0,5) = "  " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & "  "
		UNIValue(2,0) = "  " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & "  "
	Else
		UNIValue(0,5) = "|"
		UNIValue(2,0) = "  " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & "  "
	End If

	If Trim(Request("txtTrackNo")) <> "" Then
		UNIValue(0,6) = "  " & FilterVar(UCase(Request("txtTrackNo")), "''", "S") & "  "
	Else
		UNIValue(0,6) = "|"
	End If


	'UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
	UNILock = DISCONNREAD :	UNIFlag = "1"								 '☜: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	Dim lgstrRetMsg											 '☜ : Record Set Return Message 변수선언 
	Dim lgADF												   '☜ : ActiveX Data Factory 지정 변수선언 
	Dim iStr

	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2, rs3)

	Set lgADF   = Nothing

	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If

	Dim FalsechkFlg
	FalsechkFlg = False

	if SetConditionData = False then Exit sub

	If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		Call DisplayMsgBox("173200", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End
	Else
		Call  MakeSpreadSheetData()
	End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr

	Response.Write "	.frm1.txtSupplierNm.value = """ & Trim(ConvSPChars(iStrSpplNm))			  	& """" & vbCr
	Response.Write "	.frm1.txtPlantNm.value = """ & Trim(ConvSPChars(iStrPlantNm))			  	& """" & vbCr
	Response.Write "	.frm1.txtPoNo.value = """ & Trim(ConvSPChars(Request("txtPoNo")))			  	& """" & vbCr
	Response.Write "	.frm1.txtPoNo.value = """ & Trim(ConvSPChars(Request("txtPoNo")))			  	& """" & vbCr

	Response.Write "	.ggoSpread.Source	   = .frm1.vspdData "			& vbCr
	Response.Write "	.ggoSpread.SSShowData	 """ & iTotstrData	 & """" & vbCr
	Response.Write "	.lgPageNo  = """ & lgPageNo   & """" & vbCr

	Response.Write "	.frm1.hdnSupplier.value = """ & Trim(ConvSPChars(Request("txtSpplCd")))			  	& """" & vbCr
	Response.Write "	.frm1.hdnPlant.value = """ & Trim(ConvSPChars(Request("txtPlantCd")))			  	& """" & vbCr
	Response.Write "	.frm1.hdnPoNo.value = """ & Trim(ConvSPChars(Request("txtPoNo")))			  	& """" & vbCr
	Response.Write "	.frm1.hdnPrFrDt.value = """ & ConvSPChars(Request("txtPrFrDt"))			  	& """" & vbCr
	Response.Write "	.frm1.hdnPrToDt.value = """ & ConvSPChars(Request("txtPrFrDt"))			  	& """" & vbCr

	Response.Write "	.frm1.hdnTrackNo.value = """ & ConvSPChars(Request("txtTrackNo"))			  	& """" & vbCr

	Response.Write "	.DbQueryOk " & intARows & "," & intTRows & vbCr

	Response.Write "End With"		& vbCr
	Response.Write "</Script>"		& vbCr

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()

	SetConditionData = False
	On Error Resume Next
	Err.Clear

	If Not(rs1.EOF Or rs1.BOF) Then
		iStrSpplNm =  rs1(1)
		Set rs1 = Nothing
	Else
		Set rs1 = Nothing
		If Len(Request("txtSupplierCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "공급처" & " : " & Request("txtSupplierCd"), "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			Call SetFoucs("SpplCd")
			Exit function
		End If
	End If

	If Not(rs2.EOF Or rs2.BOF) Then
		iStrPlantNm =  rs2(1)
		Set rs2 = Nothing
	Else
		Set rs2 = Nothing
		If Len(Request("txtPlantCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "공장" & " : " & Request("txtPlantCd"), "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			Call SetFoucs("PlantCd")
			Exit function
		End If
	End If


	If Not(rs3.EOF Or rs3.BOF) Then
	   istrPoNo =  rs3(0)
		Set rs3 = Nothing
	Else
		Set rs3 = Nothing
		If Len(Request("txtPoNo")) Then
			Call DisplayMsgBox("970000", vbInformation, "발주번호" & " : " & Request("txtPoNo"), "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			Call SetFoucs("PoNo")
			Exit function
		End If
	End If

	SetConditionData = True

End Function
'----------------------------------------------------------------------------------------------------------
' Name : SetFoucs
' Desc :
'----------------------------------------------------------------------------------------------------------
Sub SetFoucs(ByVal Opt)
	Response.Write "<Script Language=vbscript>"								& vbCr
	Response.Write "With parent.frm1"										& vbCr
	Response.Write "	If  """ & Opt & """ = ""SpplCd"" Then "			& vbCr
	Response.Write "		.txtSupplierCd.focus() "							& vbCr
	Response.Write "	Elseif  """ & Opt & """ = ""PlantCd"" Then "		& vbCr
	Response.Write "		.txtPlantCd.focus() "							& vbCr
	Response.Write "	Else  "								& vbCr
	Response.Write "		.txtPoNo.focus() "				& vbCr
	Response.Write "	End If"								& vbCr
	Response.Write "End With"								& vbCr
	Response.Write "</Script>"								& vbCr
End Sub
'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Dim iLoopCount
	Dim iRowStr
	Dim ColCnt
	DIM i
	Dim PvArr

	lgDataExist	= "Yes"

	If CLng(lgPageNo) > 0 Then
		rs0.Move	 = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)				  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		intTRows	 = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
	End If

	iLoopCount = 0
	ReDim PvArr(C_SHEETMAXROWS_D - 1)

	Do while Not (rs0.EOF Or rs0.BOF)

		iLoopCount =  iLoopCount + 1
		iRowStr = ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("bp_cd"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("bp_nm"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("plant_cd"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("plant_nm"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("item_cd"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("item_nm"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("spec"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("po_no"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("po_seq_no"))
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs0("po_dt"))
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("po_qty"),ggQty.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("po_unit"))
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("rcpt_qty"),ggQty.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("sl_cd"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("sl_nm"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("tracking_no"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("pur_grp_nm"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("pr_no"))
		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount

		If iLoopCount - 1 < C_SHEETMAXROWS_D Then
		   istrData = istrData & iRowStr & Chr(11) & Chr(12)
		   PvArr(iLoopCount - 1) = istrData
		   istrData = ""
		Else
		   lgPageNo = lgPageNo + 1
		   Exit Do
		End If
		rs0.MoveNext

   Loop
	intARows = iLoopCount
	iTotstrData = Join(PvArr, "")

	If CLng(lgPageNo) > 0 Then
	 	MaxRow2 = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo) + iLoopCount
	Else
		MaxRow2 = CLng(iLoopCount)
	End If

	If iLoopCount < C_SHEETMAXROWS_D Then									  '☜: Check if next data exists
	   lgPageNo = 0
	End If

	MaxCount = iLoopCount
	rs0.Close													   '☜: Close recordset object
	Set rs0 = Nothing												'☜: Release ADF
End Sub
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)

	If Trim(lRow) = "" Then Exit Function
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function

%>
