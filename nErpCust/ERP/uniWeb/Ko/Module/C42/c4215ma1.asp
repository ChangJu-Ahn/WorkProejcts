<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 요소별실제원가조회 
'*  3. Program ID           : c4215ma1.asp
'*  4. Program Name         : 요소별실제원가조회 
'*  5. Program Desc         : 요소별실제원가조회 
'*  6. Modified date(First) : 2005-08-30
'*  7. Modified date(Last)  : 2005-08-30
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c4215mb1.asp"                               'Biz Logic ASP

Dim iDBSYSDate
Dim iStrFromDt
Dim lgStrPrevKey2
Dim lgRow, lgEOF1, lgEOF2

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
Dim lgCurrGrid
Dim lgCopyVersion
Dim lgErrRow, lgErrCol

Dim lgSTime		' -- 디버깅 타임체크 
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
    
    lgStrPrevKey = ""	
    lgRow = 0
    lgStrPrevKey2 = ""	
    lgEOF1 = False
    lgEOF2 = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.txtSTART_DT.Text = Left(iStrFromDt, 7)

    If parent.gPlant <> "" Then
		frm1.txtPlant_Cd.value = UCase(parent.gPlant)
		frm1.txtPlant_Nm.value = parent.gPlantNm
	End If
	
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

		.style.display = "none"
		.Redraw = False

		.MaxRows = 0
		.MaxCols = pMaxCols

		.ColHeaderRows = 3

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021122",,"" ',,parent.gAllowDragDropSpread 

		ggoSpread.SSSetEdit		1,	"공장"	, 5,,,,1	
		ggoSpread.SSSetEdit		2,	"작업지시C/C"	, 7,,,,1	
		ggoSpread.SSSetEdit		3,	"작업지시C/C명"	, 10
		ggoSpread.SSSetEdit		4,	"품목계정"	, 7,,,,1	
		ggoSpread.SSSetEdit		5,	"품목그룹"	,10 ,,,,1
		ggoSpread.SSSetEdit		6,	"품목그룹명"	,15
		ggoSpread.SSSetEdit		7,	"품목"	, 10,,,,1
		ggoSpread.SSSetEdit		8,	"품목명"	, 12
		ggoSpread.SSSetFloat	9,	"생산수량"	, 12,		Parent.ggQtyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		ggoSpread.SSSetFloat	10,	"제조원가"	, 12,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		
		For i = 11 To pMaxCols	
			ggoSpread.SSSetFloat	i,	""	, 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec
		Next
		
		For i = 1 To 8
			.Col = i: .Row = -1: .ColMerge = 1
			ret = .AddCellSpan(i, -1000, 1, 3)	' -- 기초 
		Next
		
		ret = .AddCellSpan(9, -1000, 1, 3)	' -- 기초 
		ret = .AddCellSpan(10, -1000, 1, 3)	' -- 기초 
		
		.Col = -1: .Row = -1000: .RowMerge = 1

		
		
		ggoSpread.SSSetSplit2(8)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()

		.rowheight(-1000) = 12	' 높이 재지정 
		.rowheight(-999) = 12	' 높이 재지정 
		.rowheight(-998) = 12	' 높이 재지정 
		
		.Redraw = True
	End With
End Sub
		

Sub SetGridHead(Byval pData)
	Dim arrRows, arrCols, i, j, iColCnt, iCol

	' -- 그리드 1 정의 
	With frm1.vspdData
		
		arrRows = Split(pData, Parent.gRowSep)

		'헤더를 ?줄로    
		.ColHeaderRows = UBound(arrRows, 1)
		
		For i = 0 To UBound(arrRows, 1) -1
			arrCols = Split(arrRows(i), Parent.gColSep)
			iColCnt = UBound(arrCols, 1)
			.Row	= CDbl(arrCols(iColCnt)) 		' -- 마지작 컬럼에 행번호가 들어있다.
			iCol = 11
			For j = 0 To iColCnt 
				.Col = iCol
				.Text = arrCols(j)
				iCol = iCol + 1
			Next
		Next
	End With
End Sub

Sub AddSpanSpecialColumn(Byval pRow1, Byval pRow2)
	Dim i, ret
	With frm1.vspdData
		For i = pRow1 To pRow2 Step 2
			ret = .AddCellSpan(9, i, 1, 2)	' -- 기초 
			ret = .AddCellSpan(10, i, 1, 2)	' -- 기초 
		Next
	End With
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

	With frm1
	
	Select Case iWhere
		Case 0
			arrParam(0) = "공장 팝업"
			arrParam(1) = "dbo.B_PLANT"	
			arrParam(2) = Trim(.txtPLANT_CD.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "공장" 

			arrField(0) = "PLANT_CD"	
			arrField(1) = "PLANT_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "공장"	
			arrHeader(1) = "공장명"
			arrHeader(2) = ""
		Case 1
			arrParam(0) = "작업지시C/C 팝업"
			arrParam(1) = "b_cost_center"	
			arrParam(2) = Trim(.txtCOST_CD.value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "작업지시C/C" 

			arrField(0) ="ED10" & Parent.gColSep &  "COST_CD"					' Field명(0)
			arrField(1) = "ED30" & Parent.gColSep & "COST_NM"					' Field명(1)
    
			arrHeader(0) = "작업지시C/C"
			arrHeader(1) = "작업지시C/C명"
			
		Case 2
			arrParam(0) = "품목계정 팝업"
			arrParam(1) = "dbo.B_MINOR"	
			arrParam(2) = Trim(.txtITEM_ACCT.value)
			arrParam(3) = ""
			arrParam(4) = "MAJOR_CD =" & FilterVar("P1001", "''", "S")
			arrParam(5) = "품목계정" 

			arrField(0) ="ED10" & Parent.gColSep & "MINOR_CD"
			arrField(1) ="ED30" & Parent.gColSep & "MINOR_NM"		
			arrField(2) = ""	
    
			arrHeader(0) = "품목계정"
			arrHeader(1) = "품목계정명"
			arrHeader(2) = "C/C Level"	
		Case 3
			arrParam(0) = "품목그룹 팝업"
			arrParam(1) = "dbo.b_item_group"	
			arrParam(2) = Trim(.txtITEM_GROUP_CD.value)
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "품목" 

			arrField(0) = "ED20" & Parent.gColSep &"ITEM_GROUP_CD"	
			arrField(1) = "ED30" & Parent.gColSep &"ITEM_GROUP_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "품목그룹"	
			arrHeader(1) = "품목그룹명"
			arrHeader(2) = ""			
		Case 4
			arrParam(0) = "품목 팝업"
			arrParam(1) = "dbo.B_ITEM"	
			arrParam(2) = Trim(.txtITEM_CD.value)
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "품목" 

			arrField(0) = "ED20" & Parent.gColSep &"ITEM_CD"	
			arrField(1) = "ED30" & Parent.gColSep &"ITEM_NM"
			arrField(2) = ""		
    
			arrHeader(0) = "품목"	
			arrHeader(1) = "품목명"
			arrHeader(2) = ""
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
				.txtPLANT_CD.value		= arrRet(0)
				.txtPLANT_NM.value		= arrRet(1)
				
			Case 1
				.txtCOST_CD.value		= arrRet(0)
				.txtCOST_NM.value		= arrRet(1)
				
			Case 2
				.txtITEM_ACCT.value		= arrRet(0)
				.txtITEM_ACCT_NM.value		= arrRet(1)
				
			Case 3
				.txtITEM_GROUP_CD.value		= arrRet(0)
				.txtITEM_GROUP_NM.value		= arrRet(1)
				
			Case 4
				.txtITEM_CD.value		= arrRet(0)
				.txtITEM_NM.value		= arrRet(1)
				
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
    
'	Call InitComboBox
    Call SetDefaultVal
    Call SetToolbar("110000000001111")	
    frm1.txtSTART_DT.focus
   	Set gActiveElement = document.activeElement			    
    
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

'==========================================================================================
'   Event Desc : 배부규칙 설정확인 버튼 클릭시 
'==========================================================================================
Function BtnPrint(byval strPrintType)
	Dim varCo_Cd, varFISC_YEAR, varREP_TYPE,EBR_RPT_ID,EBR_RPT_ID2
	Dim StrUrl  , i

	Dim intCnt,IntRetCD

	MsgBox "미개발"
	Exit Function

    If Not chkField(Document, "1") Then					'⊙: This function check indispensable field
       Exit Function
    End If
    

    StrUrl = StrUrl & "VER_CD|"			& frm1.txtVER_CD.value 

     ObjName = AskEBDocumentName("C4002MA1", "ebr")
     
     if  strPrintType = "VIEW" then
		Call FncEBRPreview(ObjName, StrUrl)
     else
		Call FncEBRPrint(EBAction,ObjName,StrUrl)
     end if	
     
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

Sub txtEND_DT_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtEND_DT_DblClick(Button)
    If Button = 1 Then
        frm1.txtEND_DT.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtEND_DT.Focus
    End If
End Sub

Sub txtCOST_ELMT_CD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

Sub txtACCT_CD_onKeyPress()
	If window.event.keyCode = 13 Then
		Call FncQuery
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	If lgRow <> Row Then
		With frm1.vspdData 
			.Col = .MaxCols : .Row = Row
			If .Text = "0" Then
				lgStrPrevKey2 = "" : lgEOF2 = False
				'Call DBQuery2
			End If
		End With
	End If
End Sub

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
		With frm1.vspdData 
			.Col = .MaxCols : .Row = NewRow
			If .Text = "0" Then
				lgStrPrevKey2 = "" : lgEOF2 = False
				'Call DBQuery2
			End If
		End With
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" And lgEOF1 = False Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
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
    If ChkKeyField=false then Exit Function 
    
    Call ggoOper.ClearField(Document, "2")

	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables 	

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

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    	IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")   
			If IntRetCD = vbNo Then
				Exit Function
			End If
    End If
    

    Call ggoOper.ClearField(Document, "1") 
    Call ggoOper.ClearField(Document, "2") 
    ggoSpread.Source = frm1.vspdData
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
	frm1.vspdData.ReDraw = False
	
    if frm1.vspdData.maxrows = 0 then exit function 
	   
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow ,frm1.vspdData.ActiveRow
    
    Dim iSeqNo
    
	frm1.vspdData.ReDraw = True
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
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
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
		
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtStartDt=" & sStartDt
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		
		If lgIntFlgMode = Parent.OPMD_CMODE Then	' -- 처음 조회일 경우 
			strVal = strVal & "&txtPLANT_CD=" & Trim(.txtPLANT_CD.value)
			strVal = strVal & "&txtCOST_CD=" & Trim(.txtCOST_CD.value)
			strVal = strVal & "&txtITEM_ACCT=" & Trim(.txtITEM_ACCT.value)
			strVal = strVal & "&txtITEM_GROUP_CD=" & Trim(.txtITEM_GROUP_CD.value)
			strVal = strVal & "&txtITEM_CD=" & Trim(.txtITEM_CD.value)
		Else
			strVal = strVal & "&txtPLANT_CD=" & Trim(.hPLANT_CD.value)
			strVal = strVal & "&txtCOST_CD=" & Trim(.hCOST_CD.value)
			strVal = strVal & "&txtITEM_ACCT=" & Trim(.hITEM_ACCT.value)
			strVal = strVal & "&txtITEM_GROUP_CD=" & Trim(.hITEM_GROUP_CD.value)
			strVal = strVal & "&txtITEM_CD=" & Trim(.hITEM_CD.value)
		End If
		
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
	
	frm1.vspdData.style.display = ""	'-- 그리드 보이게..
	
	Frm1.vspdData.Focus
   	
    Set gActiveElement = document.ActiveElement   

	lgIntFlgMode = Parent.OPMD_UMODE
	
	window.status = "응답시간: " & DateDiff("s", lgSTime, Time) & " 초"
End Function


'========================================================================================
' Function Name : SetQuerySpreadColor
' Function Desc : 소계 및 총계 색상변경 
'========================================================================================
Sub SetQuerySpreadColor(Byval pGrpRow)

	Dim arrRow, arrCol, iRow
	Dim iLoopCnt, i
	Dim ret, iCnt

	With frm1.vspdData	
	.ReDraw = False
	arrRow = Split(pGrpRow, Parent.gRowSep)
	
	iLoopCnt = UBound(arrRow, 1)
	
	For i = 0 to iLoopCnt -1
		arrCol = Split(arrRow(i), Parent.gColSep)
	
		.Col = -1
		.Row = CDbl(arrCol(1))*2 -1	' -- 행 
		
		Select Case arrCol(0)
			Case "1"
			   ret = .AddCellSpan(7, .Row , 2, 2)
				.BackColor = RGB(250,250,210) 
				.ForeColor = vbBlack
			Case "2"
				ret = .AddCellSpan(5, .Row , 4, 2)
				.BackColor = RGB(204,255,153) 
				.ForeColor = vbBlack
			Case "3"
				ret = .AddCellSpan(4, .Row , 5, 2)
				.BackColor = RGB(204,255,255) 
				.ForeColor = vbBlack
			Case "4"  
				ret = .AddCellSpan(2, .Row , 7, 2)
				.BackColor = RGB(255,228,181) 
				.ForeColor = vbBlack
			Case "5" 
				ret = .AddCellSpan(1, .Row , 8, 2)
				.BackColor = RGB(255,240,245) 
				.ForeColor = vbBlack
		End Select
		'.BlockMode = False
	Next

	.ReDraw = True
	End With

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
'check plant
	If Trim(frm1.txtPLANT_CD.value) <> "" Then
		strWhere = " plant_cd= " & FilterVar(frm1.txtPLANT_CD.value, "''", "S") & "  "

		Call CommonQueryRs(" plant_nm ","	 b_plant ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtPLANT_CD.alt,"X")
			frm1.txtPLANT_CD.focus 
			frm1.txtPLANT_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPLANT_NM.value = strDataNm(0)
	Else
		frm1.txtPLANT_NM.value=""
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
'check item_acct
	If Trim(frm1.txtITEM_ACCT.value) <> "" Then
		strWhere = " minor_cd  = " & FilterVar(frm1.txtITEM_ACCT.value, "''", "S") & " "		
		strWhere = strWhere & "		and major_cd=" & filterVar("P1001","","S")
		
		Call CommonQueryRs(" minor_nm  ","	 b_minor ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtITEM_ACCT.alt,"X")
			frm1.txtITEM_ACCT.focus 
			frm1.txtITEM_ACCT_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtITEM_ACCT_NM.value = strDataNm(0)
	ELSE
		frm1.txtITEM_ACCT_NM.value=""
	End If
'check item group 
	If Trim(frm1.txtITEM_GROUP_CD.value) <> "" Then
		strFrom = " B_ITEM_GROUP  "
		strWhere = " item_group_cd  = " & FilterVar(frm1.txtITEM_GROUP_CD.value, "''", "S") & " "		
		
		Call CommonQueryRs(" item_group_nm  ", strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtITEM_GROUP_CD.alt,"X")
			frm1.txtITEM_GROUP_CD.focus 
			frm1.txtITEM_GROUP_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtITEM_GROUP_NM.value = strDataNm(0)
	ELSE
		frm1.txtITEM_GROUP_NM.value=""
	End If


	'check ITEM
		If Trim(frm1.txtITEM_CD.value) <> "" Then
		strFrom = " B_ITEM  "
		strWhere = " item_cd  = " & FilterVar(frm1.txtITEM_CD.value, "''", "S") & " "		
		
		Call CommonQueryRs(" item_nm  ", strFrom , strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtITEM_CD.alt,"X")
			frm1.txtITEM_CD.focus 
			frm1.txtITEM_NM.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtITEM_NM.value = strDataNm(0)
	ELSE
		frm1.txtITEM_NM.value=""
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD6" valign=top><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtSTART_DT CLASS=FPDTYYYYMM title=FPDATETIME ALT="시작 기준년월" tag="12" id=txtSTART_DT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtPLANT_CD" TYPE="Text" MAXLENGTH="4" tag="15XXXU" size="10" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(0)">
									<input NAME="txtPLANT_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">작업지시 C/C</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtCOST_CD" TYPE="Text" MAXLENGTH="10" tag="15XXXU" size="10" ALT="작업지시 C/C"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(1)">
									<input NAME="txtCOST_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_ACCT" TYPE="Text" MAXLENGTH="2" tag="15XXXU" size="10" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(2)">
									<input NAME="txtITEM_ACCT_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
								</TR>    
								<TR>
									<TD CLASS="TD5">품목그룹</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_GROUP_CD" TYPE="Text" MAXLENGTH="10" tag="15XXXU" size="10" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(3)">
									<input NAME="txtITEM_GROUP_NM" TYPE="TEXT"  tag="14XXX" size="20">
									</TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><input NAME="txtITEM_CD" TYPE="Text" MAXLENGTH="18" tag="15XXXU" size="20" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenPopup(4)">
									<input NAME="txtITEM_NM" TYPE="TEXT"  tag="14XXX" size="20">
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData1 NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" style="display: none"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hSTART_DT" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPLANT_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCOST_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hITEM_ACCT" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hITEM_GROUP_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hITEM_CD" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

