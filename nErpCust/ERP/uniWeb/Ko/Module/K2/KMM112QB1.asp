<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        :
'*  3. Program ID           : MM112QB1
'*  4. Program Name         : 멀티컴퍼니수발주진행조회 
'*  5. Program Desc         : 멀티컴퍼니수발주진행조회 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/07/02
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Han Kwang Soo
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%

call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")


'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4,rs5
Dim istrData
Dim iStrPoNo
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim iLngMaxRow		' 현재 그리드의 최대Row
Dim iLngRow
Dim GroupCount
Dim lgCurrency
Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
Dim lgDataExist
Dim lgPageNo
Dim SoCompanyNm			'☜ : 수주법인 
Dim lgOpModeCRUD
Dim Inti
Dim intARows
Dim intTRows

intARows=0
intTRows=0
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status


Call HideStatusWnd                                                               '☜: Hide Processing message
lgOpModeCRUD  = Request("txtMode")

Select Case lgOpModeCRUD
    Case CStr(UID_M0001)                                                         '☜: Query
		 Call  SubBizQueryMulti()
End Select

Sub SubBizQueryMulti()


	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist      = "No"
	iLngMaxRow = CLng(Request("txtMaxRows"))
	lgStrPrevKey = Request("lgStrPrevKey")

'	Call DisplayMsgBox(lgStrSQL, vbInformation, "", "", I_MKSCRIPT)


	Call FixUNISQLData()		'☜ : DB-Agent로 보낼 parameter 데이타 set

	Call QueryData()			'☜ : DB-Agent를 통한 ADO query

	'-----------------------
	'Result data display area
	'-----------------------

%>

	<Script Language=vbscript>
		With parent
			.frm1.txtSoCompanyCd.value = "<%=ConvSPChars(Request("txtSoCompanyCd"))%>"
			.frm1.txtSoCompanyNm.Value	= "<%=SoCompanyNm%>"
			.frm1.txtSoCompanyCd.focus

			Set .gActiveElement = .document.activeElement

			If "<%=lgDataExist%>" = "Yes" Then

				'Show multi spreadsheet data from this line

				.ggoSpread.Source    = .frm1.vspdData
				.ggoSpread.SSShowData "<%=istrData%>"                  '☜: Display data

				.lgPageNo			 =  "<%=lgPageNo%>"				    '☜: Next next data tag

				.DbQueryOk <%=intARows%>,<%=intTRows%>

			End If
		End with
	</Script>
<%
End Sub


'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim PvArr
	Const C_SHEETMAXROWS_D  = 100
    Dim iLoopCount
    Dim iRowStr
    Dim ColCnt

		Const		M_MC_PO_SO_STS_SO_COMPANY		=	0
		Const		B_BIZ_PARTNER_BP_FULL_NM		=	1
		Const		M_MC_PO_SO_STS_PO_NO		    =	2
		Const		M_MC_PO_SO_STS_PO_SEQ_NO	    =	3
		Const		M_PUR_ORD_DTL_ITEM_CD	        =	4
		Const		B_ITEM_ITEM_NM					=	5
		Const		B_ITEM__SPEC					=	6
		Const		M_MC_PO_SO_STS_POSTS		    =	7
		Const		M_MC_PO_SO_STS_POSTS_NM	        =	8
		Const		M_MC_PO_SO_STS_SOSTS		    =	9
		Const		M_MC_PO_SO_STS_SOSTS_NM	        =	10
		Const		M_PUR_ORD_DTL_PO_UNIT	        =	11
		Const		M_MC_PO_SO_STS_PO_QTY		    =	12
		Const		M_MC_PO_SO_STS_SO_QTY		    =	13
		Const		M_MC_PO_SO_STS_PO_LC_QTY	    =	14
		Const		M_MC_PO_SO_STS_SO_LC_QTY	    =	15
		Const		M_MC_PO_SO_STS_SO_REQ_QTY	    =	16
		Const		M_MC_PO_SO_STS_SO_ISSUE_QTY     =	17
		Const		M_MC_PO_SO_STS_SO_CC_QTY	    =	18
		Const		M_MC_PO_SO_STS_PO_CC_QTY	    =	19
		Const		M_MC_PO_SO_STS_PO_RCPT_QTY	    =	20
		Const		M_MC_PO_SO_STS_SO_BILL_QTY	    =	21
		Const		M_MC_PO_SO_STS_PO_IV_QTY	    =	22
		Const		M_MC_PO_SO_STS_SO_NO		    =	23
		Const		M_MC_PO_SO_STS_SO_SEQ_NO	    =	24

    lgDataExist    = "Yes"

    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       intTRows		= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If

	'----- 레코드셋 칼럼 순서 ----------
	'A.SO_COMPANY, G.BP_FULL_NM, A.PO_NO, A.PO_SEQ_NO, B.ITEM_CD, D.ITEM_NM, D.SPEC,
	'A.PO_STS, E.MINOR_NM,
	'A.SO_STS, F.MINOR_NM,
	'B.PO_UNIT, A.PO_QTY, A.SO_QTY, A.PO_LC_QTY,
	'A.SO_REQ_QTY, A.SO_ISSUE_QTY, A.SO_CC_QTY, A.PO_CC_QTY, A.PO_RCPT_QTY,
	'A.SO_BILL_QTY, A.PO_IV_QTY
	'-----------------------------------

	iLoopCount = 0
    ReDim PvArr(C_SHEETMAXROWS_D - 1)

	Do while Not (rs0.EOF Or rs0.BOF)

		iLoopCount =  iLoopCount + 1
		iRowStr = ""

		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_PO_SO_STS_SO_COMPANY))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(B_BIZ_PARTNER_BP_FULL_NM))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_PO_SO_STS_PO_NO))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_PO_SO_STS_PO_SEQ_NO))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_PO_SO_STS_SO_NO))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_PO_SO_STS_SO_SEQ_NO))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_ORD_DTL_ITEM_CD))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(B_ITEM_ITEM_NM))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(B_ITEM__SPEC))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_PO_SO_STS_POSTS))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_PO_SO_STS_POSTS_NM))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_PO_SO_STS_SOSTS))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_MC_PO_SO_STS_SOSTS_NM))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(M_PUR_ORD_DTL_PO_UNIT))
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_PO_SO_STS_PO_QTY), ggAmtOfMoney.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_PO_SO_STS_SO_QTY), ggAmtOfMoney.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_PO_SO_STS_PO_LC_QTY), ggAmtOfMoney.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_PO_SO_STS_SO_LC_QTY), ggAmtOfMoney.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_PO_SO_STS_SO_REQ_QTY), ggAmtOfMoney.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_PO_SO_STS_SO_ISSUE_QTY), ggAmtOfMoney.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_PO_SO_STS_SO_CC_QTY), ggAmtOfMoney.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_PO_SO_STS_PO_CC_QTY), ggAmtOfMoney.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_PO_SO_STS_PO_RCPT_QTY), ggAmtOfMoney.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_PO_SO_STS_SO_BILL_QTY), ggAmtOfMoney.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(M_MC_PO_SO_STS_PO_IV_QTY), ggAmtOfMoney.DecPoint,0)

		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount

		If iLoopCount - 1 < C_SHEETMAXROWS_D Then
		   istrData = istrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount-1) = istrData
		   istrData = ""
		Else
		   lgPageNo = lgPageNo + 1
		   Exit Do
		End If

		rs0.MoveNext
	Loop


	istrData = Join(PvArr, "")

	intARows = iLoopCount
	If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
	  lgPageNo = ""
	End If

	rs0.Close                                                       '☜: Close recordset object
	Set rs0 = Nothing	                                            '☜: Release ADF
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next
    SetConditionData = false


	If Not(rs1.EOF Or rs1.BOF) Then
		SoCompanyNm = rs1("BP_NM")
		Set rs1 = Nothing
	Else
		Set rs1 = Nothing
		If Len(Request("txtSoCompanyCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "수주법인", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		    exit function
		End If
	End If


    SetConditionData = True
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(1,4)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 

	strVal = ""
    UNISqlId(0) = "MM112QA101"
    UNISqlId(1) = "MM111MA103"		'수주법인조회 
    
    UNIValue(0,0) = "'^'"
	UNIValue(1,0) = "'zzzzzzzzzz'"

    '발주법인조회 
    'If Trim(Request("txtPoCompanyCd")) <> "" Then
	'	UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("txtPoCompanyCd"))), " " , "SNM") & "' "
	'Else
	'    UNIValue(0,0) = "|"
	'End If

    '수주법인조회 
    If Trim(Request("txtSoCompanyCd")) <> "" Then
	    UNIValue(0,1) = " '"& FilterVar(Trim(UCase(Request("txtSoCompanyCd"))), " " , "SNM") & "' "
	    UNIValue(1,0) = " '"& FilterVar(Trim(UCase(Request("txtSoCompanyCd"))), " " , "SNM") & "' "
	Else
	    UNIValue(0,1) = "|"
	End If


    '발주일 
    If Trim(Request("txtPoFrDt")) <> "" Then
		UNIValue(0,2) =  " '" & Trim(UniConvDate(Request("txtPoFrDt"))) & "' "
    Else
        UNIValue(0,2) = "|"
	End If

    If Trim(Request("txtPoToDt")) <> "" Then
		UNIValue(0,3) =  " '" & Trim(UniConvDate(Request("txtPoToDt"))) & "' "
    Else
        UNIValue(0,3) = "|"
	End If

    'B/L확정처리여부 
    If Trim(Request("rdoImportFlg")) <> "" Then
		UNIValue(0,4) = " '"& FilterVar(Trim(UCase(Request("rdoImportFlg"))), " " , "SNM") & "' "
	Else
	    UNIValue(0,4) = "|"
	End If


     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")


    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	Set lgADF   = Nothing

    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If

	If SetConditionData = False Then Exit Sub


    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else
            Call  MakeSpreadSheetData()

    End If
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
