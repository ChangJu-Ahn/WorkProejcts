<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        :공정별원가분석2 
'*  3. Program ID           : c4231ma1.asp
'*  4. Program Name         : 공정별원가분석2
'*  5. Program Desc         : 공정별원가분석2
'*  6. Modified date(First) : 2005-12-12
'*  7. Modified date(Last)  : 2005-12-12
'*  8. Modifier (First)     : HJO
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
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c4231mb1.asp"                               'Biz Logic ASP

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2
Const TAB3 = 3																		'☜: Tab의 위치 



Dim iDBSYSDate
Dim iStrFromDt
Dim iStrToDt

iDBSYSDate = "<%=GetSvrDate%>"
iStrToDt = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)	
iStrFromDt= UNIDateAdd("m", -1,iStrToDt, parent.gServerDateFormat)
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgQueryFlag
Dim IsOpenPop          
Dim lgCurrGrid
Dim lgCopyVersion
Dim lgErrRow, lgErrCol
Dim lgStrPrevKey2
Dim lgSTime		' -- 디버깅 타임체크 
Dim  gSelframeFlg
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'======================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    
    lgStrPrevKey = ""

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	frm1.txtFrom_YYYYMM.Text =UniConvDateAToB(iStrFromDt, parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtTo_YYYYMM.text =UniConvDateAToB(iStrToDt, parent.gServerDateFormat, parent.gDateFormat)
	
	
	Call ggoOper.FormatDate(frm1.txtFrom_YYYYMM, parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtTo_YYYYMM, parent.gDateFormat, 2)
	
	frm1.txtWc_cd.focus 
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
	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "MA") %>
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
Sub InitSpreadSheet(byVal iTab, byVal iMaxCols)
	Dim i, ret
		
    'Call AppendNumberPlace("6","3","0")
    '--------------TAB1
    SELECT CASE ITAB
		CASE TAB1
			With frm1.vspdData
		
			ggoSpread.Source = frm1.vspdData
			'ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			ggoSpread.Spreadinit "V20021106", , ""
				

			.style.display = "none"
			.Redraw = False

			.MaxRows = 0
			.MaxCols = iMaxCols		
			'.ColHidden = True

			'헤더를 2줄로    
			.ColHeaderRows = 1

			.Col = -1: .Row = -1000: .RowMerge = 1
			.Col = 1: .Row = -1: .ColMerge = 1
			.Col = 2: .Row = -1: .ColMerge = 1

			
			ggoSpread.SSSetEdit		1,	"공정"	, 10,,,,1
			ggoSpread.SSSetEdit		2,	"공정명"	, 20		


			For i = 3 To iMaxCols	
					ggoSpread.SSSetFloat	i,		""		, 13,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			Next		


			.Col = iMaxCols		
			.ColHidden = True
			
			.rowheight(-1000) = 12	' 높이 재지정 
			
			
'			Call ggoSpread.SSSetColHidden(1,1,True)
			ggoSpread.SSSetSplit2(2) 
			.ReDraw = True		
			End With
		
			ggoSpread.SpreadLockWithOddEvenRowColor()
    
		CASE TAB2
			With frm1.vspdData2
		
			ggoSpread.Source = frm1.vspdData2
			'ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			ggoSpread.Spreadinit "V20021106", , ""
				

			.style.display = "none"
			.Redraw = False


			.MaxRows = 0
			.MaxCols = iMaxCols
			.ColHeaderRows =2
		
			.Col = -1: .Row = -1000: .RowMerge = 1
			
			.Col = 1: .Row = -1: .ColMerge = 1
			.Col = 2: .Row = -1: .ColMerge = 1
			.Col = 3: .Row = -1: .ColMerge = 1
			.Col = 4: .Row = -1: .ColMerge = 1
			.Col = 5: .Row = -1: .ColMerge = 1
			.Col =6: .Row = -1: .ColMerge = 1
			
		
			ggoSpread.SSSetEdit		1,	"공정"	, 10,,,,1
			ggoSpread.SSSetEdit		2,	"공정명"	, 18
			ggoSpread.SSSetEdit		3,	"C/C"	, 10,,,,1
			ggoSpread.SSSetEdit		4,	"C/C명"	, 18
			ggoSpread.SSSetEdit		5,	"배부요소"	, 10,,,,1
			ggoSpread.SSSetEdit		6,	"배부요소명"	,20
			


			For i = 7 To iMaxCols	STEP 3
					ggoSpread.SSSetFloat	i,		""		, 20,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+1,		""		, 20,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+2,		""		, 20,		Parent.ggExchRateNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			'ggExchRateNo	
			Next		
			
			For i = 1 To 6
				ret = .AddCellSpan(i, -1000 , 1, 2)
			Next

			.rowheight(-1000) = 12	' 높이 재지정 
			.rowheight(-999) = 12	' 높이 재지정 

			
			.Col = iMaxCols		 
			.ColHidden = True
'			Call ggoSpread.SSSetColHidden(1,1,True)
			'ggoSpread.SSSetSplit2(4) 
			.ReDraw = True		
			End With
		
			ggoSpread.SpreadLockWithOddEvenRowColor()
    

		
		CASE TAB3
		
			With frm1.vspdData3
		
			ggoSpread.Source = frm1.vspdData3
			'ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
			ggoSpread.Spreadinit "V20021106", , ""
				

			.style.display = "none"
			.Redraw = False


			.MaxRows = 0
			.MaxCols =iMaxCols
			.ColHeaderRows = 2

			.Col = -1: .Row = -1000: .RowMerge = 1
			.Col = 1: .Row = -1: .ColMerge = 1
			.Col = 2: .Row = -1: .ColMerge = 1
			.Col = 3: .Row = -1: .ColMerge = 1

			
			ggoSpread.SSSetEdit		1,	"공장"	, 10,,,,1
			ggoSpread.SSSetEdit		2,	"공정", 10,,,,1
			ggoSpread.SSSetEdit		3,	"공정명"	, 18



			For i =4 To iMaxCols 	STEP 4
					ggoSpread.SSSetFloat	i,	""	, 15,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+1,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+2,	""	, 15,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
					ggoSpread.SSSetFloat	i+3,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
			next
			For i = 1 To 3
				ret = .AddCellSpan(i, -1000 , 1, 2)
			Next

			.rowheight(-1000) = 12	' 높이 재지정 
			.rowheight(-999) = 17	' 높이 재지정 
			.Col = iMaxCols		
			.ColHidden = True
			'Call ggoSpread.SSSetColHidden(1,1,True)
			'ggoSpread.SSSetSplit2(4) 
			.ReDraw = True		
			End With
		
			ggoSpread.SpreadLockWithOddEvenRowColor()	
	END SELECT 	
End Sub



'========================================================================================
' Function Name : SetGridHead
' Function Desc : set grid head row
'========================================================================================
Sub SetGridHead(Byval pData)
	Dim arrRows, arrCols, i, j, iColCnt, iCol
	
	Select Case 	gSelframeFlg
		CASE TAB1
		' -- 그리드 1 정의 
			With frm1.vspdData			
				arrRows = Split(pData, Parent.gRowSep)
				For i = 0 To UBound(arrRows, 1) -1
					arrCols = Split(arrRows(i), Parent.gColSep)
					iColCnt = UBound(arrCols, 1)
					.Row	= CDbl(arrCols(iColCnt))		' -- 마지막 컬럼에 행번호가 들어있다.
					iCol =3
					For j = 0 To iColCnt 
						.Col = iCol
						Select Case j
							Case 0, 1,  2,iColCnt
								.Text = arrCols(j)
								iCol = iCol + 1
							Case Else
								.Text = arrCols(j)
								 iCol = iCol + 1	: .Col = iCol	' -- 금액 
						End SElect									
					Next
				Next
			End With		
		CASE TAB2	
			With frm1.vspdData2			
				arrRows = Split(pData, Parent.gRowSep)
				For i = 0 To UBound(arrRows, 1) -1
					arrCols = Split(arrRows(i), Parent.gColSep)
					iColCnt = UBound(arrCols, 1)
					.Row	= CDbl(arrCols(iColCnt))		' -- 마지막 컬럼에 행번호가 들어있다.
					iCol =7
					For j = 0 To iColCnt 
						.Col = iCol
						Select Case j
							Case 0, 1,  2,3,4,5,6,iColCnt
								.Text = arrCols(j)
								iCol = iCol + 1
							Case Else
								.Text = arrCols(j)
								 iCol = iCol + 1	: .Col = iCol	' -- 금액 
						End SElect									
					Next
				Next
			End With
	
		CASE TAB3	
			' -- 그리드 1 정의 
			With frm1.vspdData3				
				arrRows = Split(pData, Parent.gRowSep)
				For i = 0 To UBound(arrRows, 1) -1
					arrCols = Split(arrRows(i), Parent.gColSep)
					iColCnt = UBound(arrCols, 1)
					.Row	= CDbl(arrCols(iColCnt))		' -- 마지막 컬럼에 행번호가 들어있다.
					iCol =4
					For j = 0 To iColCnt 
						.Col = iCol
						Select Case j
							Case 0, 1,  2,3,iColCnt
								.Text = arrCols(j)
								iCol = iCol + 1
							Case Else
								.Text = arrCols(j)
								 iCol = iCol + 1	: .Col = iCol	' -- 금액 
						End SElect									
					Next
				Next
			End With
	
	END SELECT
End Sub

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
	If Not chkField(Document, "1") Then
			   Exit Function
	End If
	
	With frm1
		
	Select Case iWhere
			Case 0
			arrParam(0) = "공장 팝업"
			arrParam(1) = " B_PLANT "
			arrParam(2) = Trim(.txtPlant_cd.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "공장" 

			arrField(0) ="ED10" & parent.gColsep &  "PLANT_CD"	
			arrField(1) ="ED20" & parent.gColsep &  "PLANT_NM"    
			arrHeader(0) = "공장"	
			arrHeader(1) = "공장명"
		Case 1
			arrParam(0) = "C/C 팝업"
			arrParam(1) = " B_COST_CENTER "
			arrParam(2) = Trim(.txtCost_cd.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "C/C" 

			arrField(0) ="ED10" & parent.gColsep &  "COST_CD"	
			arrField(1) ="ED30" & parent.gColsep &  "COST_NM"    
			arrHeader(0) = "C/C"	
			arrHeader(1) = "C/C명"
		Case 2
			arrParam(0) = "공정 팝업"
			arrParam(1) = "(select wc_cd, wc_nm from  p_work_center union all "
			arrParam(1) = arrParam(1) &  " select pur_grp as wc_cd,pur_grp_nm as wc_nm from b_pur_grp ) a"
			arrParam(2) = Trim(.txtWC_cd.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "공정" 

			arrField(0) ="ED10" & parent.gColsep &  "WC_CD"	
			arrField(1) ="ED30" & parent.gColsep &  "WC_NM"    
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
			Case 1
			Case 2
				.txtWC_Cd.value		= arrRet(0)
				.txtWC_NM.value		= arrRet(1)				
				.txtWC_Cd.focus
		End Select
		lgBlnFlgChgValue = True
	End With
	
End Function

'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox
    
End Sub

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
	
	Call ggoOper.FormatDate(frm1.txtFrom_YYYYMM,   parent.gDateFormat,2)
	Call ggoOper.FormatDate(frm1.txtTo_YYYYMM, parent.gDateFormat,2)
	
	Call SetDefaultVal
   
   Call ClickTab1()	
	 gIsTab     = "Y" 
	 gTabMaxCnt = 3

   Call SetToolbar("110000000001111")	    
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


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

Sub txtFrom_YYYYMM_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub  txtFrom_YYYYMM_DblClick(Button)
    If Button = 1 Then
        frm1.txtFrom_YYYYMM.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrom_YYYYMM.Focus
    End If
End Sub
Sub txtTo_YYYYMM_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub  txtTo_YYYYMM_DblClick(Button)
    If Button = 1 Then
        frm1.txtFrom_YYYYMM.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtTo_YYYYMM.Focus
    End If
End Sub


Sub txtCost_cd_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub
Sub txtCost_cd_onChange()
	If frm1.txtCost_cd.value ="" then frm1.txtCost_nm.value=""
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    'ggoSpread.Source = frm1.vspdData
    'Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    	
End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1
            .vspdData.Row = NewRow 
            .vspdData.Col = 1
			.hGridKey.value= .vspdData.Text
		
		End With
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1
            .vspdData2.Row = NewRow
            .vspdData2.Col = 1
            .hGridKey.value=.vspdData2.text

        End With

    End If
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData3_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
   If Row <> NewRow And NewRow > 0 Then
       With frm1
            .vspdData3.Row = NewRow
             .vspdData3.Col = 2
			.hGridKey.value=.vspdData3.text

        End With

    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) And lgStrPrevKey <> "" Then
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub

Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3,NewTop) And lgStrPrevKey <> "" Then
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

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
    
    sStartDt= Replace(frm1.txtFrom_YYYYMM.text, parent.gComDateType, "")
    sEndDt= Replace(frm1.txtTo_YYYYMM.text, parent.gComDateType, "")

	If ValidDateCheck(frm1.txtFrom_YYYYMM, frm1.txtTo_YYYYMM) = False Then 
		frm1.txtFrom_YYYYMM.focus 
		Exit Function
	End If
	
    IF ChkKeyField()=False Then Exit Function 

    
    Call ggoOper.ClearField(Document, "2")
    
    Call InitVariables 	
    frm1.hGridKey.value=""

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


End Function


Function FncCancel() 
    Dim lDelRows

	lgBlnFlgChgValue = True
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
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub


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
    Call InitSpreadSheet(gActiveSpdSheet.id)      
    'Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
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
	Dim strVal, strNext
	Dim tmpGrp, iRow, iMaxRows

    DbQuery = False
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF

    Err.Clear	
    
    Dim sStartDt, sEndDt, sYear, sMon, sDay
    
    With frm1
		Call parent.ExtractDateFromSuper(.txtFrom_YYYYMM.Text, parent.gDateFormat,sYear,sMon,sDay)	
		sStartDt= (sYear&sMon)
		Call parent.ExtractDateFromSuper(.txtTo_YYYYMM.Text, parent.gDateFormat,sYear,sMon,sDay)
		sEndDt=sYear&sMon
		
		strNext=lgStrPrevKey		
		tmpGrp=split(.txtTmp.value, Parent.gColSep)
		iMaxRows= ubound(tmpGrp) 
		
		For iRow = 0 to iMaxRows
			select case gSelframeFlg
				Case TAB1
					If .vspdData.Row =tmpGrp(iRow) then 
							.hGridKey.value=""					
					End If				
				CASE TAB2
					If .vspdData2.Row =tmpGrp(iRow) then 
							.hGridKey.value=""					
					End If	
				CASE TAB3
					If  .vspdData3.Row =tmpGrp(iRow) then 
							.hGridKey.value=""					
					End If					
			End Select
		Next

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & strNext
			strVal = strVal & "&txtFrom_YYYYMM=" & Trim(.hYYYYMM.value)
			strVal = strVal & "&txtTo_YYYYMM=" & Trim(.hYYYYMM2.value)
			strVal = strVal & "&txtWc_cd=" & Trim(.hWc_CD.value)
			strVal = strVal & "&txtFrame=" & gSelframeFlg
			strVal = strVal & "&txtGridKey=" & trim(.hGridKey.value)

		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & strNext
			strVal = strVal & "&txtFrom_YYYYMM=" & sStartDt
			strVal = strVal & "&txtTo_YYYYMM=" & sEndDt					
			strVal = strVal & "&txtWc_cd=" & Trim(.txtWc_cd.value)			
			strVal = strVal & "&txtFrame=" & gSelframeFlg
			strVal = strVal & "&txtGridKey=" & trim(.hGridKey.value)
		End If
		Call RunMyBizASP(MyBizASP, strVal)
   
    End With
    
    DbQuery = True

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk(byval currRow)	
	lgIntFlgMode = parent.OPMD_UMODE	

	SELECT CASE gSelframeFlg
	CASE TAB1 
		frm1.vspdData.style.display = ""	'-- 그리드 보이게..	
		frm1.vspdData.Col=1 : frm1.vspdData.Row=currRow '공정  
		frm1.hGridKey.value= frm1.vspdData.Text
		Frm1.vspdData.Focus

	CASE TAB2
		frm1.vspdData2.style.display = ""	'-- 그리드 보이게..	
		frm1.vspdData2.Col=1 : frm1.vspdData2.Row=currRow '공정  
		frm1.hGridKey.value= frm1.vspdData2.Text
		Frm1.vspdData2.Focus
		
	CASE TAB3
		frm1.vspdData3.style.display = ""	'-- 그리드 보이게..
		frm1.vspdData3.Col=2 : frm1.vspdData3.Row=currRow '공정  
		frm1.hGridKey.value= frm1.vspdData3.Text	
		Frm1.vspdData3.Focus
	
	END SELECT 

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================
' Function Name : SetQuerySpreadColor
' Function Desc : 소계 및 총계 색상변경 
'========================================================================================
Sub SetQuerySpreadColor(byVal arrStr)

	Dim arrRow, arrCol, iRow
	Dim iLoopCnt, i
	Dim ret, iCnt, strRowI
	
	Select case  gSelframeFlg
	CASE TAB1 

	CASE TAB2
			
			
	CASE TAB3
		 With frm1.vspdData3		
			.ReDraw = False			
			arrRow = Split(arrStr, Parent.gRowSep)			
			iLoopCnt = UBound(arrRow, 1)
			For i = 0 to iLoopCnt -1
				arrCol = Split(arrRow(i), Parent.gColSep)
			
				.Col = -1
				.Row = CDbl(arrCol(2))	' -- 행 
			
				Select Case arrCol(0)
					Case "%1"
						iRow = .Row	: .Row2=.Row
						.Col = arrCol(1) : .Col2=.MaxCols
						.BlockMode = True
					   'ret = .AddCellSpan(C_PlantCd, 1 ,5, iRow)   '시작컬럼, 시작로, 길이컬럼, 길이행 
					   ret = .AddCellSpan(1, iRow ,3, 1)
						.BackColor = RGB(250,250,210) 
						.ForeColor = vbBlack
						.BlockMode = False
					Case "%2"
						iRow = .Row	: .Row2=.Row
						.Col = arrCol(1) : .Col2=.MaxCols
						.BlockMode = True
					   'ret = .AddCellSpan(C_PlantCd, 1 ,5, iRow)   '시작컬럼, 시작로, 길이컬럼, 길이행 
					   ret = .AddCellSpan(2, iRow ,2, 1)
						.BackColor =RGB(204,255,153)  
						.ForeColor = vbBlack
						.BlockMode = False
					Case "%3"
						iRow = .Row	: .Row2=.Row
						.Col = arrCol(1) : .Col2=.MaxCols
						.BlockMode = True
					   'ret = .AddCellSpan(C_PlantCd, 1 ,5, iRow)   '시작컬럼, 시작로, 길이컬럼, 길이행 
					   ret = .AddCellSpan(4, iRow ,1, 1)
						.BackColor = RGB(204,255,255) 
						.ForeColor = vbBlack
						.BlockMode = False

					Case "%4"
						iRow = .Row	: .Row2=.Row
						.Col = arrCol(1) : .Col2=.MaxCols
						.BlockMode = True
					   'ret = .AddCellSpan(C_PlantCd, 1 ,5, iRow)   '시작컬럼, 시작로, 길이컬럼, 길이행 
					   ret = .AddCellSpan(5, iRow ,2, 1)
						.BackColor = RGB(255,228,181) 
						.ForeColor = vbBlack
						.BlockMode = False
						
			
				End Select
				.Col = 1: .Row = -1: .ColMerge = 1
				.Col = 2: .Row = -1: .ColMerge = 1
				.Col = 3: .Row = -1: .ColMerge = 1
				strRowI = strRowI & CDbl(arrCol(2)) & Parent.gColSep
			Next

			frm1.txtTmp.value=frm1.txtTmp.value & strRowI
			.ReDraw = True
			End With	
	
	End SELECT

End Sub

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
'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
   
	Call changeTabs(TAB1)
	gSelframeFlg = TAB1
 '   Call InitSpreadSheet(gSelframeFlg)
 
	frm1.vspdData.style.display="none"
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables

	IF frm1.hGridKey.value <>"" THEN 	CALL DBQUERY	
		 
End Function

Function ClickTab2()
   
	Call changeTabs(TAB2)	 
	gSelframeFlg = TAB2
	frm1.vspdData2.style.display="none"
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData

	Call InitVariables

	IF frm1.hGridKey.value <>"" THEN 	CALL DBQUERY
    
End Function

Function ClickTab3()
   
	Call changeTabs(TAB3)	 
	gSelframeFlg = TAB3
	frm1.vspdData3.style.display="none"
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData	
	Call InitVariables

	IF frm1.hGridKey.value <>"" THEN 	CALL DBQUERY
End Function

Function ClickTab4()
   
	Call changeTabs(TAB4)	 
	gSelframeFlg = TAB4
	frm1.vspdData4.style.display="none"
	ggoSpread.Source = frm1.vspdData4
	ggoSpread.ClearSpreadData
	
	Call InitVariables
	IF frm1.hGridKey.value <>"" THEN 	CALL DBQUERY

End Function


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

	'check wc CD
	If Trim(frm1.txtwc_cd.value) <> "" Then
		strWhere = " wc_cd = " & FilterVar(frm1.txtwc_cd.value, "''", "S") & "  "
		strFrom ="  (select wc_cd, wc_nm from  p_work_center union all  select pur_grp as wc_cd,pur_grp_nm as wc_nm from b_pur_grp ) a "

		Call CommonQueryRs(" wc_nm ",strFrom, strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtwc_cd.alt,"X")			
			frm1.txtwc_nm.value = ""
			ChkKeyField = False
			frm1.txtwc_cd.focus 
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtwc_nm.value = strDataNm(0)
	Else
		frm1.txtwc_nm.value=""
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
					<TD CLASS="CLSMTABP" >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공정불량현황</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP" >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공정별배부요소DATA</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP" >
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>공정별재공현황</font></td>
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
									<TD CLASS="TD6" valign=top> 
										<TABLE>
											<TR>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFrom_YYYYMM CLASS=FPDTYYYYMM title=FPDATETIME ALT="시작 작업년월" tag="12" id=txtFrom_YYYYMM></OBJECT>');</SCRIPT>
												</TD>
												<TD>~</TD>
												<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtTo_YYYYMM CLASS=FPDTYYYYMM title=FPDATETIME ALT="종료 작업년월" tag="12" id=txtTo_YYYYMM></OBJECT>');</SCRIPT>	
												
												</TD>
											</TR>
										 </TABLE>
									</TD>
									
									<TD CLASS="TD5">공정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtWC_Cd" SIZE=10 MAXLENGTH=7 tag="11xxxU" ALT="공정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup(2)">
									<INPUT TYPE=TEXT NAME="txtWC_NM" SIZE=25 tag="14">
									</TD>
								</TR>  
															
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top >
						<DIV ID="TabDiv" style="display:none;"STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%;" SCROLL=NO>
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
						</DIV>
						<DIV ID="TabDiv"  style="display:none;" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%;" SCROLL=NO>
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData2 NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
						</DIV>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%;" SCROLL=NO style="display:none;">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData3 NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO  noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hYYYYMM2" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCost_cd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hWc_cd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlant_cd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hGridKey" tag="24" TABINDEX= "-1">


<INPUT TYPE=HIDDEN NAME="txtTmp" tag="24" TABINDEX= "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

