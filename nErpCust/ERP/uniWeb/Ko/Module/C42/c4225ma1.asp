<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 배부요소DATA조회 
'*  3. Program ID           : c4225ma1.asp
'*  4. Program Name         : 배부요소DATA조회 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2005-11-24
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : choe0tae 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js">			</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c4225mb1.asp"                               'Biz Logic ASP

Dim iDBSYSDate
Dim iStrFromDt

iDBSYSDate = "<%=GetSvrDate%>"
iStrFromDt = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)	

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgQueryFlag
Dim IsOpenPop          
Dim gSelframeFlg
Dim lgCopyVersion
Dim lgErrRow, lgErrCol

Dim lgSTime	' -- 디버깅 타임체크 

Dim lgRow, lgStrPrevKey2, lgStrPrevKey3

Const TAB1 = 1													'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
		
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'======================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    
    lgStrPrevKey = ""	: lgStrPrevKey2 = "" : lgStrPrevKey3 = ""
    lgRow = 0

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.txtSTART_DT.Text =UniConvDateAToB(iStrFromDt, parent.gServerDateFormat, parent.gDateFormat)	
	Call ggoOper.FormatDate(frm1.txtSTART_DT, parent.gDateFormat, 2)
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call LoadInfTB19029A("Q","C", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(Byval pMaxCols)
	Dim i, ret

	With frm1.vspdData

		.Redraw = False

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021122",,parent.gForbidDragDropSpread

		.MaxRows = 0
		.MaxCols = pMaxCols
		
		.Col = pMaxCols
		.ColHidden = True

		'헤더를 2줄로    
		.ColHeaderRows = 2

		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		
		ggoSpread.SSSetEdit		1,	"배부요소", 10,,,,1	
		ggoSpread.SSSetEdit		2,	"배부요소명", 15
		ggoSpread.SSSetEdit		3,	"생성구분"	, 10,,,,1
		
		 Call AppendNumberPlace("6","18","6")

		For i = 4 To pMaxCols - 1
			ggoSpread.SSSetFloat	i,	""	, 15,		"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		Next		
		ggoSpread.SSSetFloat	pMaxCols,	"ROW_SEQ"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
				
		Call ggoSpread.SSSetColHidden(pMaxCols , pMaxCols, True)

		ggoSpread.SSSetSplit2(3)

		For i = 1 To 3
			ret = .AddCellSpan(i, -1000 , 1, 2)
		Next

		.rowheight(-1000) = 12	' 높이 재지정 
		.rowheight(-999) = 12	' 높이 재지정 
	
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.Redraw = True
	End With
End Sub
		
Sub InitSpreadSheet2(Byval pMaxCols)
	Dim i, ret

	With frm1.vspdData2

		.Redraw = False

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021122",,parent.gForbidDragDropSpread

		.MaxRows = 0
		.MaxCols = pMaxCols
		
		.Col = pMaxCols
		.ColHidden = True

		'헤더를 2줄로    
		.ColHeaderRows = 2

		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		
		ggoSpread.SSSetEdit		1,	"배부요소", 10,,,,1	
		ggoSpread.SSSetEdit		2,	"배부요소명", 15
		ggoSpread.SSSetEdit		3,	"생성구분"	, 10,,,,1
		ggoSpread.SSSetEdit		4,	"C/C"	, 10,,,,1
		ggoSpread.SSSetEdit		5,	"C/C명"	, 15

		 Call AppendNumberPlace("6","18","6")
		For i = 6 To pMaxCols - 1
			ggoSpread.SSSetFloat	i,	""	, 15,		"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		Next		
		ggoSpread.SSSetFloat	pMaxCols,	"ROW_SEQ"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
				
		Call ggoSpread.SSSetColHidden(pMaxCols , pMaxCols, True)

		ggoSpread.SSSetSplit2(5)

		For i = 1 To 5
			ret = .AddCellSpan(i, -1000 , 1, 2)
		Next

		.rowheight(-1000) = 12	' 높이 재지정 
		.rowheight(-999) = 12	' 높이 재지정 

		ggoSpread.SpreadLockWithOddEvenRowColor()
		.Redraw = True
	End With

End Sub

Sub InitSpreadSheet3()
	Dim i, ret

	With frm1.vspdData3

		.Redraw = False

		ggoSpread.Source = frm1.vspdData3
		ggoSpread.Spreadinit "V20021122",,parent.gForbidDragDropSpread

		.MaxRows = 0
		.MaxCols = 10
		
		.Col = .MaxCols
		.ColHidden = True

		.Col = 1: .Row = -1: .ColMerge = 1
		.Col = 2: .Row = -1: .ColMerge = 1
		.Col = 3: .Row = -1: .ColMerge = 1
		.Col = 4: .Row = -1: .ColMerge = 1
		.Col = 5: .Row = -1: .ColMerge = 1
		.Col = 6: .Row = -1: .ColMerge = 1
		.Col = 7: .Row = -1: .ColMerge = 1
		
		ggoSpread.SSSetEdit		1,	"배부요소", 10,,,,1	
		ggoSpread.SSSetEdit		2,	"배부요소명", 15
		ggoSpread.SSSetEdit		3,	"생성구분"	, 10,,,,1
		ggoSpread.SSSetEdit		4,	"C/C"	, 10,,,,1
		ggoSpread.SSSetEdit		5,	"C/C명"	, 15
		ggoSpread.SSSetEdit		6,	"공정"	, 10,,,,1
		ggoSpread.SSSetEdit		7,	"공정명"	, 15
		ggoSpread.SSSetEdit		8,	"오더번호"	, 20,,,,1
		
		 Call AppendNumberPlace("6","18","6")
		
		ggoSpread.SSSetFloat	9,	"배부요소DATA"	, 15,		"6",		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	10,	"ROW_SEQ"	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
				
		Call ggoSpread.SSSetColHidden(10 , 10, True)
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.Redraw = True
	End With

End Sub


Sub SetGridHead(Byval pData)
	Dim arrRows, arrCols, i, j, iColCnt, iCol

	' -- 그리드 1 정의 
	With frm1.vspdData		
		arrRows = Split(pData, Parent.gRowSep)	
		For i = 0 To UBound(arrRows, 1) -1
			arrCols = Split(arrRows(i), Parent.gColSep)
			iColCnt = UBound(arrCols, 1)
			.Row	= CDbl(arrCols(iColCnt))		' -- 마지작 컬럼에 행번호가 들어있다.			
			iCol = 4				
			For j = 0 To iColCnt 
				.Col = iCol
				Select Case j
					Case 0, 1, 2, 3, 4, 5, iColCnt
						.Text = arrCols(j)
						iCol = iCol + 1
					Case Else
						.Text = arrCols(j)
						 iCol = iCol + 1	: .Col = iCol	' -- 금액 
				End SElect
				
			Next
		Next
	End With
End Sub

Sub SetGridHead2(Byval pData)
	Dim arrRows, arrCols, i, j, iColCnt, iCol

	' -- 그리드 1 정의 
	With frm1.vspdData2		
		arrRows = Split(pData, Parent.gRowSep)    
		For i = 0 To UBound(arrRows, 1) -1
			arrCols = Split(arrRows(i), Parent.gColSep)
			iColCnt = UBound(arrCols, 1)
			.Row	= CDbl(arrCols(iColCnt))		' -- 마지작 컬럼에 행번호가 들어있다.
			iCol = 6	
			For j = 0 To iColCnt 
				.Col = iCol
				Select Case j
					Case 0, 1, 2, iColCnt
						.Text = arrCols(j)
						iCol = iCol + 1
					Case Else
						.Text = arrCols(j)
						 iCol = iCol + 1	: .Col = iCol	' -- 금액 
				End SElect				
			Next
		Next
	End With
End Sub


'======================================================================================================
' 기능: Tab Click
' 설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)									'~~~ 첫번째 Tab 
	gSelframeFlg = TAB1

	If lgIntFlgMode = parent.OPMD_UMODE And frm1.vspdData.MaxRows = 0 Then
		Call DBQuery
	End If
End Function

Function ClickTab2()
 
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)									'~~~ 첫번째 Tab 
	gSelframeFlg = TAB2  

	If lgIntFlgMode = parent.OPMD_UMODE And frm1.vspdData2.MaxRows = 0 Then
		Call DBQuery
	End If
End Function

Function ClickTab3()

	If gSelframeFlg = TAB3 Then Exit Function
	Call changeTabs(TAB3)									'~~~ 첫번째 Tab 
	gSelframeFlg = TAB3 

	If lgIntFlgMode = parent.OPMD_UMODE And frm1.vspdData3.MaxRows = 0 Then
		Call DBQuery
	End If

End Function

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
 ' -- 그리드1에서 팝업 클릭시 
Function OpenPopUp(Byval iWhere)
	Dim arrRet, sTmp
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
	
	Select Case iWhere
		Case 0
			arrParam(0) = "배부요소 팝업"
			arrParam(1) = "dbo.b_minor"	
			arrParam(2) = Trim(.txtDSTB_FCTR_CD.value)
			arrParam(3) = ""
			arrParam(4) = " MAJOR_CD=" & FilterVar("C4000", "''", "S")
			arrParam(5) = "배부요소" 

			arrField(0) = "MINOR_CD"	
			arrField(1) = "MINOR_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "배부요소"	
			arrHeader(1) = "배부요소명"
			arrHeader(2) = ""
			
		Case 1
			arrParam(0) = "C/C 팝업"
			arrParam(1) = "dbo.b_cost_center"	
			arrParam(2) = Trim(.txtCOST_CD.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "C/C" 

			arrField(0) = "COST_CD"
			arrField(1) = "COST_NM"		
			
			arrHeader(0) = "C/C"
			arrHeader(1) = "C/C명"

		Case 2
			arrParam(0) = "공정 팝업"
			arrParam(1) = "dbo.p_work_center"	
			arrParam(2) = Trim(.txtWC_CD.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "공정" 

			arrField(0) = "WC_CD"
			arrField(1) = "WC_NM"		
			
			arrHeader(0) = "공정"
			arrHeader(1) = "공정명"

	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

	End With
End Function


Function SetPopUp(Byval arrRet, Byval iWhere)
	Dim sTmp
	
	With frm1

		Select Case iWhere
		
			Case 0
				.txtDSTB_FCTR_CD.value		= arrRet(0)
				.txtDSTB_FCTR_NM.value		= arrRet(1)
				
			Case 1
				.txtCOST_CD.value		= arrRet(0)
				.txtCOST_NM.value		= arrRet(1)

			Case 2
				.txtWC_CD.value		= arrRet(0)
				.txtWC_NM.value		= arrRet(1)
				
		End Select
		lgBlnFlgChgValue = True
	End With
	
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
	
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatDate(frm1.txtSTART_DT, parent.gDateFormat,2)
    'Call InitSpreadSheet
    Call InitVariables
    
    Call AppendNumberPlace("6","18","6")
    
	'Call InitComboBox
    Call SetDefaultVal
    Call SetToolbar("110000000001111")	
    frm1.txtSTART_DT.focus
   	Set gActiveElement = document.activeElement		

    frm1.vspdData.style.display = "none"
    frm1.vspdData2.style.display = "none"
    frm1.vspdData3.style.display = "none"
   	
   	Call ClickTab1	    
    
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
Function GetText4Grid(Byref pGrid, Byval pCol, Byval pRow)
	With pGrid
		If pGrid.MaxRows = 0 Then Exit Function
		If pRow = "" Then pRow = .ActiveRow
		.Col = pCol : .Row = pRow : GetText4Grid = Trim(.Text)
	End With
End Function


'========================================================================================================

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

Sub txtSTART_DT_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtSTART_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtSTART_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtSTART_DT.Focus
    End If
End Sub


Sub txtDSTB_FCTR_CD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtCOST_CD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtWC_CD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub


'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
End Sub
'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData_MouseDown(Button,Shift,x,y)
	gSelframeFlg = TAB1
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

Sub vspdData2_MouseDown(Button,Shift,x,y)
	gSelframeFlg = TAB2	
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub

Sub vspdData3_MouseDown(Button,Shift,x,y)
	gSelframeFlg = TAB3
	If Button <> "1" And gMouseClickStatus = "SP3C" Then
		gMouseClickStatus = "SP3CR"
	End If

End Sub

'========================================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) And lgStrPrevKey2 <> "" Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub


Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3,NewTop) And lgStrPrevKey3 <> "" Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub
'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    

End Sub

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    

End Sub

Sub vspdData3_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
    Dim IntRetCD , sStartDt, sEndDt
    
    FncQuery = False
    
    Err.Clear
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    If ChkKeyField=False Then Exit Function
    
    Call ggoOper.ClearField(Document, "2")

	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

	frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData

	frm1.vspdData3.MaxRows = 0
    ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData

    Call InitVariables 	

    frm1.vspdData.style.display = "none"
    frm1.vspdData2.style.display = "none"
    frm1.vspdData3.style.display = "none"

    IF DbQuery = False Then
		Exit Function
	END IF
       
    FncQuery = True		
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False 
    
    Err.Clear     

    Call ggoOper.ClearField(Document, "1") 
    Call ggoOper.ClearField(Document, "2") 

	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

	frm1.vspdData2.MaxRows = 0
    ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData

	frm1.vspdData3.MaxRows = 0
    ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData

     
    Call ggoOper.LockField(Document, "N") 
    Call InitVariables 
    Call SetDefaultVal
    
    FncNew = True 

End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave() 
    
    FncSave = True      
    
End Function


'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy() 

    if frm1.vspdData.maxrows = 0 then exit function 

	frm1.vspdData.ReDraw = False
	   
	frm1.vspdData.ReDraw = True
End Function


Function FncCancel() 
    Dim lDelRows

	FncCancel = True
End Function


'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD, iSeqNo, iSubSeqNo
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
			Exit Function
		End If	
	End If

	With frm1

	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	End With
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
End Function


Function FncDeleteRow() 
    Dim lDelRows
	
	lgBlnFlgChgValue = True
End Function

Function FncPrint()
    Call parent.FncPrint() 
End Function

Function FncPrev() 
End Function

Function FncNext() 
End Function

Function FncExcel() 
	If lgStrPrevKey <> "" Then
		lgStrPrevKey = lgStrPrevKey & parent.gColSep & "*"	' 현재 키값 & * 를 보낸다 
		Call DBQuery()
		Exit Function
	End If

	Call parent.FncExport(Parent.C_MULTI)

End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
End Function
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    'Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	'Call InitData()
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    FncExit = True
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF

    Err.Clear	    
    Dim sStartDt, sEndDt, sYear, sMon, sDay
    
    With frm1    
		If lgIntFlgMode = Parent.OPMD_CMODE Then	' -- 처음 조회일 경우		
			Call parent.ExtractDateFromSuper(.txtSTART_DT.Text, parent.gDateFormat,sYear,sMon,sDay)		
			sStartDt = sYear & sMon		
		Else
			sStartDt = .hSTART_DT.value 
		End If
				
		strVal = BIZ_PGM_ID & "?txtMode=" & lgIntFlgMode
		strVal = strVal & "&txtStartDt=" & sStartDt

		If lgIntFlgMode = Parent.OPMD_UMODE Then	' -- 재조회(다음 데이타 조회)
			Select Case gSelframeFlg
				Case TAB1
					strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
				Case TAB2
					strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey2
				Case TAB3
					strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey3
			End Select
		End If
		
		If lgIntFlgMode = Parent.OPMD_CMODE Then
			strVal = strVal & "&txtDSTB_FCTR_CD=" & Trim(.txtDSTB_FCTR_CD.value)
			strVal = strVal & "&txtCOST_CD=" & Trim(.txtCOST_CD.value)
			strVal = strVal & "&txtWC_CD=" & Trim(.txtWC_CD.value)
		Else
			strVal = strVal & "&txtDSTB_FCTR_CD=" & Trim(.hDSTB_FCTR_CD.value)
			strVal = strVal & "&txtCOST_CD=" & Trim(.hCOST_CD.value)
			strVal = strVal & "&txtWC_CD=" & Trim(.hWC_CD.value)
		End If

		strVal = strVal & "&TYPE_FLAG=" & CStr(gSelframeFlg)

		lgSTime = Time	' -- 디버깅용 
		Call RunMyBizASP(MyBizASP, strVal)
   End With
    
    DbQuery = True 
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	


	Select Case gSelframeFlg
		Case TAB1
			frm1.vspdData.style.display = ""
			Frm1.vspdData.Focus
		Case TAB2
			frm1.vspdData2.style.display = ""
			Frm1.vspdData2.Focus
		Case TAB3
			frm1.vspdData3.style.display = ""
			Frm1.vspdData3.Focus
	End Select

    Set gActiveElement = document.ActiveElement   
	lgIntFlgMode = Parent.OPMD_UMODE	
	'window.status = "응답시간: " & DateDiff("s", lgSTime, Time) & " 초"
	
	If lgStrPrevKey = "*" Then
		lgStrPrevKey = ""
		Call FncExcel() 
	End If
End Function
'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave() 

    DbSave = True    
    
End Function

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()	
   
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'=================================================================================
'	Name : ChkKeyField()
'	Description : check the valid data
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere , strFrom 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       

	ChkKeyField = true		
'check DSTB_FCTR
	If Trim(frm1.txtDSTB_FCTR_CD.value) <> "" Then
		strWhere = " minor_cd= " & FilterVar(frm1.txtDSTB_FCTR_CD.value, "''", "S") & "  "
		strWhere = strWhere & "		and MAJOR_CD=" & FilterVar("C4000", "''", "S")

		Call CommonQueryRs(" minor_nm ","	 b_minor ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtDSTB_FCTR_CD.alt,"X")
			frm1.txtDSTB_FCTR_CD.focus 
			frm1.txtDSTB_FCTR_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtDSTB_FCTR_NM.value = strDataNm(0)
	Else
		frm1.txtDSTB_FCTR_NM.value=""
	End If
'check CC
	If Trim(frm1.txtCOST_CD.value) <> "" Then
		strFrom = " b_cost_center  "
		strWhere = " COST_CD  = " & FilterVar(frm1.txtCOST_CD.value, "''", "S") & " "		
		
		Call CommonQueryRs(" COST_NM  ", strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtCOST_CD.alt,"X")
			frm1.txtCOST_CD.focus 
			frm1.txtCOST_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtCOST_NM.value = strDataNm(0)
	ELSE
		frm1.txtCOST_NM.value=""
	End If	
	
	'check WC
	If Trim(frm1.txtWC_CD.value) <> "" Then
		strFrom = " p_work_center "	

		strWhere = " WC_CD  = " & FilterVar(frm1.txtWC_CD.value, "''", "S") & " "		
		
		Call CommonQueryRs("WC_NM  ", strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtWC_CD.alt,"X")
			frm1.txtWC_CD.focus 
			frm1.txtWC_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If	
		strDataNm = split(lgF0,chr(11))
		frm1.txtWC_NM.value = strDataNm(0)
	ELSE
		frm1.txtWC_NM.value=""
	End If
End Function 
	
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" oncontextmenu="javascript:return false">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onClick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>C/C</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onClick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공정</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onClick="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>오더</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">작업년월</TD>
									<TD CLASS="TD6" valign=top><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtSTART_DT CLASS=FPDTYYYYMM title=FPDATETIME ALT="시작 작업년월" tag="12" id=txtSTART_DT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>배부요소</TD>
									<TD CLASS="TD6" valign=top><input NAME="txtDSTB_FCTR_CD" TYPE="Text" MAXLENGTH="2" tag="15XXXU" size="10" ALT="배부요소"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(0)"><input NAME="txtDSTB_FCTR_NM" TYPE="TEXT" tag="14XXX" size="20">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5" NOWRAP>C/C</TD>
									<TD CLASS="TD6" valign=top><input NAME="txtCOST_CD" TYPE="Text" MAXLENGTH="10" tag="15XXXU" size="10" ALT="배부요소"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(1)"><input NAME="txtCOST_NM" TYPE="TEXT"  tag="14XXX" size="20">
									<TD CLASS="TD5" NOWRAP>공정</TD>
									<TD CLASS="TD6" valign=top><input NAME="txtWC_CD" TYPE="Text" MAXLENGTH="7" tag="15XXXU" size="10" ALT="배부요소"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(2)"><input NAME="txtWC_NM" TYPE="TEXT" tag="14XXX" size="20">
								</TR>    
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
					<DIV ID=TabDiv STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%;" SCROLL=NO>
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR HEIGHT=100%>
								<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					<DIV ID=TabDiv STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%;" SCROLL=NO>
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR HEIGHT=100%>
								<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData2 NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					<DIV ID=TabDiv STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%;" SCROLL=NO>
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR HEIGHT=100%%>
								<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData3 NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hSTART_DT" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hDSTB_FCTR_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCOST_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hWC_CD" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

