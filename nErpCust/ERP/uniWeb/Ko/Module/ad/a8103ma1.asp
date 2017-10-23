<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a8103ma1
'*  4. Program Name         : 결의전표미달거래조회 
'*  5. Program Desc         : 결의전표내역을 등록, 수정, 삭제, 조회 
'*  6. Component List       : PAGG005.dll
'*  7. Modified date(First) : 2001/01/16
'*  8. Modified date(Last)  : 2003/10/28
'*  9. Modifier (First)     : Jang Sung Hee
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">			 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">		 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs">	 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs">		 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs">		 </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE=vbscript>

Option Explicit		
'########################################################################################################
'#                       4.  Data Declaration Part
'========================================================================================================
'=                       4.1 External ASP File
Const BIZ_PGM_ID        = "a8103MB1.asp"                         '☆: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "a8103MB2.asp"                         '☆: Biz logic spread sheet for #2
'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
Const C_MaxKey            = 5                                    '☆☆☆☆: Max key value
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim  lgIsOpenPop                                             '☜: Popup status
Dim  lgStrPrevKey_1, lgStrPrevKey_2, lgStrPrevKey_3, lgStrPrevKey_4

Dim lsClickRow
Dim lsClickRow2

<%
Dim lsSvrDate
lsSvrDate = GetSvrDate
%>

'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
Sub  InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed

    lgStrPrevKey_1   = ""
    lgStrPrevKey_2   = ""
    lgStrPrevKey_3   = ""
    lgStrPrevKey_4   = ""
    
    lsClickRow	= ""
	lsClickRow2	= ""
End Sub

'========================================================================================================
Sub  SetDefaultVal()
	Dim strToData

	strToData = UNIDateClientFormat("<%=lsSvrDate%>")

	frm1.txtToDt.Text	= strToData
	frm1.txtFromDt.Text	= UNIDateAdd("m", -1, strToData, parent.gDateFormat )
End Sub

'========================================================================================================
Sub  LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub  InitSpreadSheet(Byval iOpt)
	Select Case iOpt
		Case "A"
			Call SetZAdoSpreadSheet("A8103MA01", "S", "A", "V20021224", parent.C_SORT_DBAGENT, frm1.vspdData,  C_MaxKey, "X", "X")
		Case "B"
			Call SetZAdoSpreadSheet("A8103MA01", "S", "B", "V20021224", parent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey, "X", "X")
		Case "C"
			Call SetZAdoSpreadSheet("A8103MA01", "S", "C", "V20021224", parent.C_SORT_DBAGENT, frm1.vspdData3, C_MaxKey, "X", "X")
		Case "D"
			Call SetZAdoSpreadSheet("A8103MA01", "S", "D", "V20021224", parent.C_SORT_DBAGENT, frm1.vspdData4, C_MaxKey, "X", "X")
	End Select

    Call SetSpreadLock (iOpt)
End Sub

'=========================================================================================================
Sub  SetSpreadLock(Byval iOpt )
	Select Case iOpt
		Case "A"
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLockWithOddEvenRowColor()
		Case "B"		
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLockWithOddEvenRowColor()       
		Case "C"
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SpreadLockWithOddEvenRowColor()       
		Case "D"
			ggoSpread.Source = frm1.vspdData4
			ggoSpread.SpreadLockWithOddEvenRowColor()       
    End Select
End Sub

'========================================================================================================
Sub  Form_Load()
	Call LoadInfTB19029()													'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field

	Call InitVariables()													'⊙: Initializes local global variables
	Call SetDefaultVal()
	Call InitSpreadSheet("A")
	Call InitSpreadSheet("B")
	Call InitSpreadSheet("C")
	Call InitSpreadSheet("D")
	Call SetToolbar("1100000000011111")										'⊙: 버튼 툴바 제어 

	frm1.txtFromAmt.Text = ""
	frm1.txtToAmt.Text = ""
	frm1.txtBizArea.focus 
End Sub

'========================================================================================================
Function  FncQuery() 
	Dim IntRetCD 

    On Error Resume Next
    Err.Clear

    FncQuery = False                                                '⊙: Processing is NG

    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
		Exit Function
    End If

	If CompareDateByFormat(frm1.txtFromDt.text, frm1.txtToDt.text, frm1.txtFromDt.Alt, frm1.txtToDt.Alt, _
                        "970025", frm1.txtFromDt.UserDefinedFormat, parent.gComDateType, True) = False Then         	
		Exit Function
	End If	

	If frm1.txtFromAmt.Text <> "" And frm1.txtToAmt.Text <> "" Then
		If UNICDbl(frm1.txtFromAmt.Text) > UNICDbl(frm1.txtToAmt.Text) Then
			IntRetCD = DisplayMsgBox("114104","X","X","X")			'⊙: "Will you destory previous data"
			Exit Function
		End If	
	End If

    Call ggoOper.ClearField(Document, "2")							'⊙: Clear Contents  Field
   	
   	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ClearSpreadData()
   	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ClearSpreadData()
   	ggoSpread.Source = frm1.vspdData3
	Call ggoSpread.ClearSpreadData()
   	ggoSpread.Source = frm1.vspdData4
	Call ggoSpread.ClearSpreadData()
	
    Call InitVariables()											'⊙: Initializes local global variables

    If DbQuery("A") = False Then
		Exit Function
    End If															'☜: Query db data

    If Err.number = 0 Then
       FncQuery = True												'☜: Processing is OK
    End If

	Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function  FncPrint() 
    Call parent.FncPrint()    
End Function

'========================================================================================================
Function  FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================================
Function  FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)
End Function

'========================================================================================================
Function FncSplitColumn()
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit

    iColumnLimit  = UBound(lgFieldNM_T)

    If parent.gMouseClickStatus = "SPCRP" Then
		ACol = Frm1.vspdData.ActiveCol
		ARow = Frm1.vspdData.ActiveRow
		If ACol > iColumnLimit Then
			iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
			Exit Function
		End If

		Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
		ggoSpread.Source = Frm1.vspdData
		ggoSpread.SSSetSplit(ACol)

		Frm1.vspdData.Col = ACol
		Frm1.vspdData.Row = ARow
		Frm1.vspdData.Action = 0
		Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If

    If parent.gMouseClickStatus = "SP1CRP" Then
		ACol = Frm1.vspdData2.ActiveCol
		ARow = Frm1.vspdData2.ActiveRow
		If ACol > iColumnLimit Then
			iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
			Exit Function
		End If

		Frm1.vspdData2.ScrollBars = parent.SS_SCROLLBAR_NONE
		ggoSpread.Source = Frm1.vspdData2
		ggoSpread.SSSetSplit(ACol)

		Frm1.vspdData2.Col = ACol
		Frm1.vspdData2.Row = ARow
		Frm1.vspdData2.Action = 0
		Frm1.vspdData2.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If
    
    If parent.gMouseClickStatus = "SP2CRP" Then
		ACol = Frm1.vspdData3.ActiveCol
		ARow = Frm1.vspdData3.ActiveRow
		If ACol > iColumnLimit Then
			iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
			Exit Function
		End If

		Frm1.vspdData3.ScrollBars = parent.SS_SCROLLBAR_NONE
		ggoSpread.Source = Frm1.vspdData3
		ggoSpread.SSSetSplit(ACol)

		Frm1.vspdData3.Col = ACol
		Frm1.vspdData3.Row = ARow
		Frm1.vspdData3.Action = 0
		Frm1.vspdData3.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If
    
    If parent.gMouseClickStatus = "SP3CRP" Then
		ACol = Frm1.vspdData4.ActiveCol
		ARow = Frm1.vspdData4.ActiveRow
		If ACol > iColumnLimit Then
			iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
			Exit Function
		End If

		Frm1.vspdData4.ScrollBars = parent.SS_SCROLLBAR_NONE
		ggoSpread.Source = Frm1.vspdData4
		ggoSpread.SSSetSplit(ACol)

		Frm1.vspdData4.Col = ACol
		Frm1.vspdData4.Row = ARow
		Frm1.vspdData4.Action = 0
		Frm1.vspdData4.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If
End Function

'========================================================================================================
Function  FncExit()
	Dim IntRetCD

    On Error Resume Next 
    Err.Clear           

	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    If Err.number = 0 Then
       FncExit = True   
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function  DbQuery(ByVal iOpt) 
	Dim strVal
	Dim strBizCd
	Dim strFromAmt
	Dim strToAmt

	On Error Resume Next
	Err.Clear

    DbQuery = False

    Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
    
	frm1.txtIOpt.value =  iOpt

	If frm1.rdoQuery.checked = True Then
		frm1.txtRdoFg.value = "1"	'미달 
	Else
		frm1.txtRdoFg.value = "2"	'완결 
	End If
		
    Select Case iOpt
		Case "A"
		   strBizCd   = Trim(frm1.txtBizArea.value)
		   strFromAmt = frm1.txtFromAmt.Text
		   strToAmt   = frm1.txtToAmt.Text
			   
	       strVal = BIZ_PGM_ID & "?txtFromDt=" & Trim(frm1.txtFromDt.Text)
	       strVal = strVal & "&txtToDt="	   & Trim(frm1.txtToDt.Text)
	       strVal = strVal & "&txtBizArea="   & EnCoding(strBizCd)
	       strVal = strVal & "&txtFromAmt="   & strFromAmt
	       strVal = strVal & "&txtToAmt="     & strToAmt
	       strVal = strVal & "&txtIOpt="      & EnCoding(iOpt)
	       strVal = strVal & "&txtRdoFg="     & Trim(frm1.txtRdoFg.value)
	       strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey_1                      '☜: Next key tag
	           
	       lsClickRow = 1
		Case "B"
			If frm1.Vspddata.MaxRows < 1 Then 
				frm1.txtErrFg.Value = "Y" 
			Else	
				frm1.txtErrFg.Value = "N" 
			End If

			strBizCd   = Trim(frm1.txtBizArea.value)
			strFromAmt = frm1.txtFromAmt.Text
			strToAmt   = frm1.txtToAmt.Text
				   
			strVal = BIZ_PGM_ID & "?txtFromDt=" & Trim(frm1.txtFromDt.Text)
			IF frm1.txtRdoFg.value="1" then 
			strVal = strVal & "&txtBizArea="   & EnCoding(strBizCd)
			end if
			strVal = strVal & "&txtToDt="	    & Trim(frm1.txtToDt.Text)
			strVal = strVal & "&txtFromAmt="   & strFromAmt
			strVal = strVal & "&txtToAmt="     & strToAmt
			strVal = strVal & "&txtIOpt="      & EnCoding(iOpt)
			strVal = strVal & "&txtRdoFg="     & Trim(frm1.txtRdoFg.value)
			strVal = strVal & "&hTemphq="      & Trim(frm1.hTemphq.value)
			strVal = strVal & "&txtErrFg="     & Trim(frm1.txtErrFg.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey_2                      '☜: Next key tag
			    
			lsClickRow2 = 1
		Case "C"
			strVal = BIZ_PGM_ID1 & "?hTempNo=" & Trim(frm1.hTempNo.value)
			strVal = strVal & "&txtIOpt="      & EnCoding(iOpt)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey_3                      '☜: Next key tag
		Case "D"
			strVal = BIZ_PGM_ID1 & "?hTempNo2=" & Trim(frm1.hTempNo2.value)
			strVal = strVal & "&txtIOpt="      & EnCoding(iOpt)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey_4                      '☜: Next key tag
	End Select

    Select Case iOpt
		Case "A"
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		Case "B"
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
		Case "C"
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("C")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("C")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("C"))
		
		Case "D"
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("D")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("D")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("D"))
    End Select

    Call RunMyBizASP(MyBizASP, strVal)										           '☜: 비지니스 ASP 를 가동 
    
	If Err.number = 0 Then
       DbQuery = True																'☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Sub DbQueryOk(ByVal iOpt)
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	Select Case	iOpt
		Case "A"
			If frm1.vspdData.MaxRows >= 1 Then
				Call SetSpreadColumnValue("A", frm1.vspdData,frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
				frm1.hTempNo.value = GetKeyPosVal("A",3)
				lsClickRow = "1"
			Else
				frm1.hTempNo.value = ""
			End If
		
			If  frm1.rdoQueryHq.checked = True Then		'완결 
				If frm1.vspdData.MaxRows >= 1 Then
					Call SetSpreadColumnValue("A", frm1.vspdData,frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
					frm1.hTemphq.value = GetKeyPosVal("A",5)
	   				ggoSpread.Source = frm1.vspdData2
					Call ggoSpread.ClearSpreadData()
					Call Dbquery("B")					
				Else
					ggoSpread.Source = frm1.vspdData3
					Call ggoSpread.ClearSpreadData()
					ggoSpread.Source = frm1.vspdData4
					Call ggoSpread.ClearSpreadData()
				End If
			Else	'미달 
				ggoSpread.Source = frm1.vspdData2
				Call ggoSpread.ClearSpreadData()
				Call Dbquery("B")
			End If
		Case "B"
			If frm1.vspdData2.MaxRows >= 1 Then		
				Call SetSpreadColumnValue("B", frm1.vspdData2,frm1.vspdData2.ActiveCol,frm1.vspdData2.ActiveRow)
				frm1.hTempNo2.value = GetKeyPosVal("B",3)
				lsClickRow2 = "1"
			Else
				frm1.hTempNo2.value = ""
			End If		
		
			If  frm1.vspdData.MaxRows >= 1 Then
				ggoSpread.Source = frm1.vspdData3
				Call ggoSpread.ClearSpreadData()
				Call Dbquery("C")
			Else
				ggoSpread.Source = frm1.vspdData3
				Call ggoSpread.ClearSpreadData()
				
				If frm1.vspdData2.MaxRows >= 1 Then
					ggoSpread.Source = frm1.vspdData4
					Call ggoSpread.ClearSpreadData()
					Call Dbquery("D")
				Else
					ggoSpread.Source = frm1.vspdData4
					Call ggoSpread.ClearSpreadData()
				End If
			End If
		Case "C"
			If  frm1.vspdData2.MaxRows >= 1 Then
				ggoSpread.Source = frm1.vspdData4
				Call ggoSpread.ClearSpreadData()
				Call Dbquery("D")
			Else
				ggoSpread.Source = frm1.vspdData4
				Call ggoSpread.ClearSpreadData()
			End If
		Case Else
	End Select
End Sub

'========================================================================================================
Function OpenPopUp(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim Field_fg

	If lgIsOpenPop = True Then Exit Function	
	
	lgIsOpenPop = True
	
	Field_fg = 1
	
	arrParam(0) = "사업장팝업"									' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"									' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizArea.value) 							' Code Condition
	arrParam(3) = ""												' Name Cindition
	arrParam(4) = ""
	arrParam(5) = "사업장코드"			
		
	arrField(0) = "BIZ_AREA_CD"	
	arrField(1) = "BIZ_AREA_NM"	
	    
	arrHeader(0) = "사업장코드"		
	arrHeader(1) = "사업장명"	
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetReturnVal(arrRet,Field_fg)
	End If	

	frm1.txtBizArea.focus
End Function

'========================================================================================================
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		Case 1
			frm1.txtBizArea.value = arrRet(0)
			frm1.txtBizAreaNm.value = arrRet(1)				
	End Select
	
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Select Case UCase(Trim(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			Call OpenOrderByPopup("A")
		Case "VSPDDATA2"
			Call OpenOrderByPopup("B")
		Case "VSPDDATA3"
			Call OpenOrderByPopup("C")
		Case "VSPDDATA4"
			Call OpenOrderByPopup("D")
	End Select
End Sub

'========================================================================================================
Function  OpenOrderByPopup(Byval pSpdNo)
	Dim arrRet

	On Error Resume Next

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp", Array(ggoSpread.GetXMLData(pSpdNo),gMethodText), "dialogWidth=" & parent.SORTW_WIDTH & " px; dialogHeight=" & parent.SORTW_HEIGHT & " px; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False

	If arrRet(0) = "X" Then
		Exit Function
	Else
		Call ggoSpread.SaveXMLData(pSpdNo, arrRet(0), arrRet(1))
		Call InitVariables()
		Call InitSpreadSheet(pSpdNo)       
	End If
End Function

'========================================================================================================
Sub  vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
   	End If

	ggoSpread.Source = frm1.vspdData

	If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
	    Exit Sub
	End If

'	Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
'	
'	If lsClickRow <> Row Then
'		ggoSpread.Source = frm1.vspdData3
'		Call ggoSpread.ClearSpreadData()
'
'	    lgStrPrevKey_3   = ""
'		frm1.hTempNo.value = GetKeyPosVal("A",3)
'
'	    If frm1.rdoQueryHq.checked = True Then
'			frm1.hTemphq.value = GetKeyPosVal("A",5)
'			ggoSpread.Source = frm1.vspdData2
'			Call ggoSpread.ClearSpreadData()
'
'			Call DbQuery("B")
'		Else
'			ggoSpread.Source = frm1.vspdData3
'			Call ggoSpread.ClearSpreadData()
'
'			Call DbQuery("C")
'		End If	
'
'		lsClickRow = Row
'	End If
End Sub

'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
		Call SetSpreadColumnValue("A", frm1.vspdData, NewCol, NewRow)
	
		If lsClickRow <> Row Then
			ggoSpread.Source = frm1.vspdData3
			Call ggoSpread.ClearSpreadData()

		    lgStrPrevKey_3   = ""
			frm1.hTempNo.value = GetKeyPosVal("A",3)

		    If frm1.rdoQueryHq.checked = True Then
				frm1.hTemphq.value = GetKeyPosVal("A",5)
				ggoSpread.Source = frm1.vspdData2
				Call ggoSpread.ClearSpreadData()
				Call DbQuery("B")
			Else
				ggoSpread.Source = frm1.vspdData3
				Call ggoSpread.ClearSpreadData()
				Call DbQuery("C")
			End If	
	
			lsClickRow = Row
		End If
    End If
End Sub

'========================================================================================================
Sub vspdData2_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
		Call SetSpreadColumnValue("B", frm1.vspdData2, NewCol, NewRow)

		If lsClickRow2 <> Row Then
			ggoSpread.Source = frm1.vspdData4
			Call ggoSpread.ClearSpreadData()
		    lgStrPrevKey_4    = ""
			frm1.hTempNo2.value = GetKeyPosVal("B",3)
			ggoSpread.Source = frm1.vspdData4
			Call ggoSpread.ClearSpreadData()

			Call DbQuery("D")
			lsClickRow2 = Row
		End If
    End If
End Sub

'========================================================================================================
Sub  vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SP1C"
	Set gActiveSpdSheet = frm1.vspdData2

	If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
   	End If

	ggoSpread.Source = frm1.vspdData2

	If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
	    Exit Sub
	End If
	
'	Call SetSpreadColumnValue("B", frm1.vspdData2, Col, Row)
'
'	If lsClickRow2 <> Row Then
'		ggoSpread.Source = frm1.vspdData4
'		Call ggoSpread.ClearSpreadData()
'	    lgStrPrevKey_4    = ""
'		frm1.hTempNo2.value = GetKeyPosVal("B",3)
'		ggoSpread.Source = frm1.vspdData4
'		Call ggoSpread.ClearSpreadData()
'
'		Call DbQuery("D")
'		lsClickRow2 = Row
'	End If
End Sub

'========================================================================================================
Sub  vspdData3_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData3

	If frm1.vspdData3.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
   	End If

	ggoSpread.Source = frm1.vspdData3

	If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
	    Exit Sub
	End If
End Sub

'=======================================================================================================
Sub  vspdData4_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SP3C"
	Set gActiveSpdSheet = frm1.vspdData4

	If frm1.vspdData4.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

	ggoSpread.Source = frm1.vspdData4

	If Row <= 0 Then
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
	    Exit Sub
	End If
End Sub

'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'========================================================================================================
Sub vspdData3_GotFocus()
    ggoSpread.Source = frm1.vspdData3
End Sub

'========================================================================================================
Sub vspdData4_GotFocus()
    ggoSpread.Source = frm1.vspdData4
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
Sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP1C" Then
		gMouseClickStatus = "SP1CR"
	End If
End Sub

'========================================================================================================
Sub vspdData3_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'========================================================================================================
Sub vspdData4_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP3C" Then
		gMouseClickStatus = "SP3CR"
	End If
End Sub

'========================================================================================================
Sub  vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
		Exit Sub
    End If

	If CheckRunningBizProcess = True Then
		Exit Sub
	End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey_1 <> "" Then                         
			If DbQuery("A") = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
        End If
    End If
End Sub

'========================================================================================================
Sub  vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If CheckRunningBizProcess = True Then
		Exit Sub
	End If
    
    If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	           
    	If lgStrPrevKey_2 <> "" Then                         
			If DbQuery("B") = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
        End If
    End If
End Sub

'========================================================================================================
Sub  vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    If frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3,NewTop) Then	           
    	If lgStrPrevKey_3 <> "" Then                         
			If DbQuery("C") = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
        End If
    End If
End Sub

'========================================================================================================
Sub  vspdData4_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If CheckRunningBizProcess = True Then
		Exit Sub
	End If

    If frm1.vspdData4.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData4,NewTop) Then	           
    	If lgStrPrevKey_4 <> "" Then                         
			If DbQuery("D") = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
        End If
    End If
End Sub

'========================================================================================================
Sub txtFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtToDt.focus	
	End If
'	Call FncQuery
End Sub

'========================================================================================================
Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromDt.focus	
'	Call FncQuery
	End If
End Sub

'========================================================================================================
Sub txtfromamt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

'========================================================================================================
Sub txttoamt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

'=========================================================================================================
Sub  txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFromDt.focus
	End If
End Sub

'========================================================================================================
Sub  txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtToDt.focus	
	End If
End Sub

'========================================================================================================
Sub rdoQuery_click()
    frm1.rdoQuery.checked = True
    frm1.rdoQueryHq.checked = False
End Sub

'========================================================================================================
'   Event Name : rdoQuery_click
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub rdoQueryHq_click()
    frm1.rdoQuery.checked = False
    frm1.rdoQueryHq.checked = True
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL=NO>
<FORM NAME=frm1 TARGET=MyBizASP METHOD=POST>
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS=CLSMTABP>
						<TABLE ID=MyTab CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH=9 HEIGHT=23></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=CENTER CLASS=CLSMTAB><FONT COLOR=WHITE>결의전표미달거래조회</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=RIGHT><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH=10 HEIGHT=23></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>					
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS=Tab11>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS=CLSFLD>
							<TABLE <%=LR_SPACE_TYPE_40%>>
			 					<TR>
									<TD CLASS=TD5 NOWRAP>사업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME=txtBizArea   ALT="사업장"   MAXLENGTH=10 SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnBizArea ALIGN=TOP TYPE=BUTTON ONCLICK="vbscript:Call OpenPopup(frm1.txtBizArea.Value)">&nbsp;
														 <INPUT TYPE=TEXT NAME=txtBizAreaNm ALT="사업장명" MAXLENGTH=20 SIZE=20 STYLE="TEXT-ALIGN: left" tag="14X"></TD>
									<TD CLASS=TD5 NOWRAP>결의일</TD>
									<TD CLASS=TD6 NOWRAP>								
												<script language =javascript src='./js/a8103ma1_fpDateTime1_txtFromDt.js'></script>~
								                <script language =javascript src='./js/a8103ma1_fpDateTime2_txtToDt.js'></script>
								    </TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>금액</TD>
									<TD CLASS=TD6 NOWRAP>
												 <script language =javascript src='./js/a8103ma1_fpDoubleSingle1_txtfromamt.js'></script>
												 &nbsp;~&nbsp;
												 <script language =javascript src='./js/a8103ma1_fpDoubleSingle2_txttoamt.js'></script>
									<TD CLASS=TD5 NOWRAP >조회유형</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=RADIO CLASS=RADIO NAME=rdoQuery    VALUE=ASC  ID=rdoQuery   ONCLICK="vbscript:Call rdoQuery_Click()" CHECKED><LABEL FOR="rdoSortMethod1">미달</LABEL>
										<INPUT TYPE=RADIO CLASS=RADIO NAME=rdoQueryHq  VALUE=DESC ID=rdoQueryHq ONCLICK="vbscript:Call rdoQueryHq_Click()"><LABEL FOR="rdoSortMethod1">완결</LABEL>
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>  
							   <TD  HEIGHT="70%">
							   	<script language =javascript src='./js/a8103ma1_vaSpread1_vspdData.js'></script>
							   </TD>
							   <TD>
							   	<script language =javascript src='./js/a8103ma1_vaSpread2_vspdData2.js'></script>
							   </TD> 
							</TR>
							<TR>
								<TD HEIGHT=2 WIDTH="100%" COLSPAN=2></TD>
							</TR>
							<TR>
								<TD  HEIGHT="30%">
									<script language =javascript src='./js/a8103ma1_vaSpread3_vspdData3.js'></script>
								</TD>
								<TD>
									<script language =javascript src='./js/a8103ma1_vaSpread4_vspdData4.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME=MyBizASP SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no NORESIZE FRAMESPACING=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME=txtMode				tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtUpdtUserId		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtInsrtUserId		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtMaxRows			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtiOpt				tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtRdoFg			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtErrFg			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hTempNo				tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hTempNo2			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hTemphq				tag="24" TABINDEX="-1">
</FORM>
<DIV ID=MousePT NAME=MousePT>
<IFRAME NAME=MouseWindow FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
</HTML>
