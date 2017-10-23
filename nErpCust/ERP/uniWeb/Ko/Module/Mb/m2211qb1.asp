<%'======================================================================================================
'*  1. Module Name		  : Basic Architect
'*  2. Function Name		: ADO Template (Save)
'*  3. Program ID		   :
'*  4. Program Name		 :
'*  5. Program Desc		 :
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2003/06/01
'*  8. Modifier (First)	 : KimTaeHyun
'*  9. Modifier (Last)	  : Kim Jin Ha
'* 10. Comment			  :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
Option Explicit
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%																		 '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	On Error Resume Next

	Dim lgADF																  '☜ : ActiveX Data Factory 지정 변수선언 
	Dim lgstrRetMsg															'☜ : Record Set Return Message 변수선언 
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0							  '☜ : DBAgent Parameter 선언 
	Dim rs1, rs2, rs3, rs4, rs5, rs6							'☜ : DBAgent Parameter 선언 
	Dim lgstrData															  '☜ : data for spreadsheet data
	Dim iTotstrData

	Dim lgTailList															 '☜ : Orderby절에 사용될 field 리스트 
	Dim lgSelectList
	Dim lgSelectListDT

	Dim lgDataExist
	Dim lgPageNo

	Dim strPoType															   '⊙ : 발주형태 
	Dim strPoFrDt															   '⊙ : 발주일 
	Dim strPoToDt															   '⊙ :
	Dim strSpplCd															   '⊙ : 공급처 
	Dim strPurGrpCd															   '⊙ : 구매그룹 
	Dim strItemCd															   '⊙ : 품목 
	Dim strTrackNo															   '⊙ : Tracking No
	Dim arrRsVal(3)														   '* : 화면에 조회해온 Name을 담아놓기 위해 만든 Array

	Call HideStatusWnd
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "QB")

	lgPageNo		 = UNICInt(Trim(Request("lgPageNo")),0)			  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")							   '☜ : select 대상목록 
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)			 '☜ : 각 필드의 데이타 타입 
	lgTailList	 = Request("lgTailList")								 '☜ : Orderby value

	Call TrimData()
	Call FixUNISQLData()
	Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100

	Dim iLoopCount
	Dim iRowStr
	Dim ColCnt
	Dim PvArr

	lgDataExist	= "Yes"
	lgstrData	  = ""

	If CLng(lgPageNo) > 0 Then
	   rs0.Move	 = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)				  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	End If

	iLoopCount = -1
	ReDim PvArr(C_SHEETMAXROWS_D - 1)

	Do while Not (rs0.EOF Or rs0.BOF)

		iLoopCount =  iLoopCount + 1
		iRowStr = ""

		For ColCnt = 0 To UBound(lgSelectListDT) - 1
			iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next

		If iLoopCount < C_SHEETMAXROWS_D Then
		   lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
		   PvArr(iLoopCount) = lgstrData
		   lgstrData = ""
		Else
		   lgPageNo = lgPageNo + 1
		   Exit Do
		End If

		rs0.MoveNext
	Loop

	iTotstrData = Join(PvArr, "")

	If iLoopCount < C_SHEETMAXROWS_D Then									  '☜: Check if next data exists
	   lgPageNo = ""
	End If
	rs0.Close													   '☜: Close recordset object
	Set rs0 = Nothing												'☜: Release ADF

End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
 Sub FixUNISQLData()

	Dim strVal
	Redim UNISqlId(5)													 '☜: SQL ID 저장을 위한 영역확보 
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	Redim UNIValue(5,16)

	UNISqlId(0) = "M2211QA101"

	UNISqlId(1) = "M2111QA302"											  '공장명 
	UNISqlId(2) = "M2111QA303"											  '품목명 
	UNISqlId(3) = "M4111QA502"											  '창고명 
	UNISqlId(4) = "M3111QA102"											  '거래처명 

	UNIValue(0,0) = lgSelectList										  '☜: Select list

	If Len(Request("txtPlantCd")) Then
		UNIValue(0,1)	=  " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
		UNIValue(0,2)	=  " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	else
		UNIValue(0,1)	=  "''"
		UNIValue(0,2)	=  "" & FilterVar("zzzzzzzzz", "''", "S") & ""
	End If

	If Len(Request("txtItemCd")) Then
		UNIValue(0,3)	= " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
		UNIValue(0,4)	= " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
	else
		UNIValue(0,3)	=  "''"
		UNIValue(0,4)	=  "" & FilterVar("zzzzzzzzz", "''", "S") & ""
	End If

	If Len(Request("txtDlvyFrDt")) Then
		UNIValue(0,5)	= " " & FilterVar(uniConvDate(Request("txtDlvyFrDt")), "''", "S") & ""
	else
		UNIValue(0,5)	= "" & FilterVar("1900/01/01", "''", "S") & ""
	End If

	If Len(Request("txtDlvyToDt")) Then
		UNIValue(0,6)	= " " & FilterVar(uniConvDate(Request("txtDlvyToDt")), "''", "S") & ""
	else
		UNIValue(0,6)	= "" & FilterVar("2999/12/30", "''", "S") & ""
	End If

	If Len(Request("txtSpplCd")) Then
		UNIValue(0,7)	= " " & FilterVar(UCase(Request("txtSpplCd")), "''", "S") & " "
		UNIValue(0,8)	= " " & FilterVar(UCase(Request("txtSpplCd")), "''", "S") & " "
	else
		UNIValue(0,7)	= "''"
		UNIValue(0,8)	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
	End If

	If Len(Request("txtSlCd")) Then
		UNIValue(0,9)	= " " & FilterVar(UCase(Request("txtSlCd")), "''", "S") & " "
		UNIValue(0,10)	= " " & FilterVar(UCase(Request("txtSlCd")), "''", "S") & " "
	else
		UNIValue(0,9)	=  "''"
		UNIValue(0,10)	=  "" & FilterVar("zzzzzzzzz", "''", "S") & ""
	End If

	If Request("rdoUseflg") = "A"then
		 UNIValue(0,11)	= ""
	elseif Request("rdoUseflg") = "F"then
		 UNIValue(0,11)	=" AND B.SPPL_TYPE = " & FilterVar("F", "''", "S") & " "
	else
		 UNIValue(0,11)	= " AND B.SPPL_TYPE = " & FilterVar("C", "''", "S") & " "
	end if

	If Len(Trim(Request("txtTrackNo"))) Then
		UNIValue(0,12) 	= " " & FilterVar(Trim(Request("txtTrackNo")), "''", "S") & ""
	Else
		UNIValue(0,12)	= " '' "
	End If
	
	//2006.10.9 Modified by KSJ
	If Request("rdoClsflg") = "A"then
		UNIValue(0,13)	= ""
	elseif Request("rdoClsflg") = "N"then
		UNIValue(0,13)	=" AND L.cls_flg = " & FilterVar("N", "''", "S") & "  "
	else
		UNIValue(0,13)	= " AND L.cls_flg = " & FilterVar("Y", "''", "S") & "  "
	end if


	UNIValue(1,0)  = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	UNIValue(2,0)  = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	UNIValue(2,1)  = " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
	UNIValue(3,0)  = " " & FilterVar(UCase(Request("txtSlCd")), "''", "S") & " "
	UNIValue(4,0)  = " " & FilterVar(UCase(Request("txtSpplCd")), "''", "S") & " "

	UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
	UNILock = DISCONNREAD :	UNIFlag = "1"								 '☜: set ADO read mode

End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
	Dim iStr
	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2,rs3,rs4)

	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If

	Dim FalsechkFlg

	FalsechkFlg = False

	If  rs1.EOF And rs1.BOF Then
		rs1.Close
		Set rs1 = Nothing
		If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		   FalsechkFlg = True
		End If
	Else
		arrRsVal(0) = rs1(1)
		rs1.Close
		Set rs1 = Nothing
	End If

	If  rs2.EOF And rs2.BOF Then
		rs2.Close
		Set rs2 = Nothing
		If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("122700", vbInformation, "모품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		   FalsechkFlg = True
		End If
	Else
		arrRsVal(1) = rs2(1)
		rs2.Close
		Set rs2 = Nothing
	End If

	If  rs3.EOF And rs3.BOF Then
		rs3.Close
		Set rs3 = Nothing
		If Len(Request("txtSlCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "출고창고", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		   FalsechkFlg = True
		End If
	Else
		arrRsVal(2) = rs3(1)
		rs3.Close
		Set rs3 = Nothing
	End If

	If  rs4.EOF And rs4.BOF Then
		rs4.Close
		Set rs4 = Nothing
		If Len(Request("txtSpplCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "공급처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		   FalsechkFlg = True
		End If
	Else
		arrRsVal(3) = rs4(1)
		rs4.Close
		Set rs4 = Nothing
	End If

	If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		rs0.Close
		Set rs0 = Nothing
	Else
		Call  MakeSpreadSheetData()
	End If

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
 Sub TrimData()
End Sub


%>

<Script Language=vbscript>
	With parent
		 .ggoSpread.Source	= .frm1.vspdData
		 .ggoSpread.SSShowData "<%=iTotstrData%>"							'☜: Display data
		 .frm1.vspdData.Redraw = False
		 .lgPageNo			=  "<%=lgPageNo%>"			   '☜ : Next next data tag

		 .frm1.hdnPlantCd.value	= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		 .frm1.hdnItemCd.value	 = "<%=ConvSPChars(Request("txtItemCd"))%>"
		 .frm1.hdnSlCd.value	   = "<%=ConvSPChars(Request("txtSlCd"))%>"
		 .frm1.hdnDlvyFrDt.value   = "<%=ConvSPChars(Request("txtDlvyFrDt"))%>"
		 .frm1.hdnDlvyToDt.value   = "<%=ConvSPChars(Request("txtDlvyToDt"))%>"
		 .frm1.hdnSpplCd.value	 = "<%=ConvSPChars(Request("txtSpplCd"))%>"
		 .frm1.hdnrdoUseflg.value  = "<%=ConvSPChars(Request("rdoUseflg"))%>"
		.frm1.hdnTrackNo.value  = "<%=ConvSPChars(Request("txtTrackNo"))%>"
		.frm1.hdnrdoClsflg.value  = "<%=ConvSPChars(Request("rdoClsflg"))%>"

		 .frm1.txtPlantNm.value		=  "<%=ConvSPChars(arrRsVal(0))%>"
  		 .frm1.txtItemNm.value		=  "<%=ConvSPChars(arrRsVal(1))%>"
  		 .frm1.txtSlNm.value		=  "<%=ConvSPChars(arrRsVal(2))%>"
  		 .frm1.txtSpplNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>"

  		 .DbQueryOk
  		 .frm1.vspdData.Redraw = True
	End with
</Script>

<%
	Response.End												'☜: 비지니스 로직 처리를 종료함 
%>

