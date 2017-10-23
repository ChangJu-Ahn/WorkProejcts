
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : 본지점전표등록 
'*  3. Program ID        : a8101ma
'*  4. Program 이름      : 본지점전표 등록 
'*  5. Program 설명      : 본지점전표 등록 수정 삭제 조회 
'*  6. Comproxy 리스트   :
'*  7. 최초 작성년월일   : 2000/09/22,2000/10/07
'*  8. 최종 수정년월일   : 2001/02/15
'*  9. 최초 작성자       : 안혜진 
'* 10. 최종 작성자       : Hersheys
'* 11. 전체 comment      :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ag/AcctCtrl_ko441_4.vbs">				</SCRIPT>

<SCRIPT LANGUAGE=vbscript>
Option Explicit		

Const BIZ_PGM_ID      = "a8101mb1_KO441.asp"	

<!-- #Include file="../../inc/lgvariables.inc" -->

'⊙: Grid Columns
Dim C_ItemSeq
Dim C_Bizareacd
Dim C_BizareaPopup
Dim C_Bizareanm
Dim C_Deptcd
Dim C_DeptPopup
Dim C_Deptnm
Dim C_AcctCd
Dim C_AcctPopup
Dim C_AcctNm
Dim C_DrCrFg
Dim C_DrCrNm
Dim C_ItemAmt
Dim C_ItemLocAmt
Dim C_ItemDesc
Dim C_ExchRate

'20090922 kbs 
Dim C_ItemSeq2

Dim lgCurrRow
Dim lgStrPrevKeyDtl
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim intItemCnt
Dim lgstartfnc
Dim lgFormLoad
Dim lgQueryOk

Dim IsOpenPop

Const C_MENU_NEW	=	"1110010100001111"
'Const C_MENU_CRT	=	"1110111100111111"
'Const C_MENU_UPD	=	"1111111100111111"
'Const C_MENU_PRT	=	"1110000000011111"	

<%
Dim lsSvrDate
lsSvrDate = GetSvrDate
%>

'========================================================================================================
Sub initSpreadPosVariables()

	C_ItemSeq		= 1								'☆: Spread Sheet 의 Columns 인덱스 
	C_Bizareacd		= 2
	C_BizareaPopup	= 3
	C_Bizareanm		= 4
	C_Deptcd		= 5
	C_DeptPopup		= 6
	C_Deptnm		= 7
	C_AcctCd		= 8
	C_AcctPopup	    = 9
	C_AcctNm		= 10
	C_DrCrFg		= 11
	C_DrCrNm		= 12
	C_ItemAmt		= 13
	C_ItemLocAmt	= 14
	C_ItemDesc		= 15
	C_ExchRate		= 16

	'20090922 kbs 
	C_ItemSeq2		= 17

End Sub

'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE 
    lgBlnFlgChgValue = False  
    lgIntGrpCount = 0  
    lgStrPrevKey = ""   
    lgCurrRow = 0 

End Sub

'=========================================================================================================
Sub SetDefaultVal()

	Call ggoOper.ClearField(Document, "1")				'Condition field clear
    frm1.txtGLDt.text = UNIFormatDate("<%=lsSvrDate%>")
    frm1.txtDocCur.value = parent.gCurrency
	frm1.txtDeptCd.value	= parent.gDepart
	frm1.hOrgChangeId.Value = parent.gChangeOrgId 

End Sub

'========================================================================================================
Sub LoadInfTB19029()

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "A", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>

End Sub

'========================================================================================================
Sub CookiePage(ByVal Kubun)
End Sub

'========================================================================================================
Sub MakeKeyStream(ByVal pOpt)

End Sub

'========================================================================================================
Sub InitComboBox()

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("A1013", "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboTaxPolicy, lgF0, lgF1, Chr(11))

End Sub

'========================================================================================================
Sub InitSpreadComboBox()

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("A1012", "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.SetCombo Replace(lgF0,Chr(11), vbTab), C_DrCrFg
	ggoSpread.SetCombo Replace(lgF1,Chr(11), vbTab), C_DrCrNm

End Sub

'========================================================================================================
Sub InitData()

	Dim intRow
	Dim intIndex

	With frm1.vspdData

		For intRow = 1 to .MaxRows
			.Row = intRow
			.Col = C_DrCrFg
			intIndex = .Value
			.col = C_DrCrNm
			.Value = intindex
		Next

	End With

End Sub

'========================================================================================================
Sub InitSpreadSheet()


	Call initSpreadPosVariables()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20051216",,parent.gAllowDragDropSpread

	With frm1.vspdData

		.ReDraw = False


		'20090922 kbs 
		'.MaxCols	= C_ExchRate + 1
		.MaxCols	= C_ItemSeq2 + 1

		.Col		= .MaxCols				'☜: 공통콘트롤 사용 Hidden Column

		Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

		Call AppendNumberPlace("6","3","0")
		Call GetSpreadColumnPos("A")

        	ggoSpread.SSSetFloat  C_ItemSeq,    "",					4,	"6", ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec
        	ggoSpread.SSSetEdit   C_Bizareacd,  "사업장",		9,	,	,	10
       		ggoSpread.SSSetButton C_BizareaPopup
        	ggoSpread.SSSetEdit   C_Bizareanm,  "사업장명",		15,	,	,	30
        	ggoSpread.SSSetEdit   C_Deptcd,     "부서코드",		10,	,	,	10
        	ggoSpread.SSSetButton C_DeptPopup
        	ggoSpread.SSSetEdit   C_Deptnm,     "부서명",		17,	,	,	30
        	ggoSpread.SSSetEdit   C_AcctCd,     "계정코드",		10,	,	,	18
		ggoSpread.SSSetButton C_AcctPopup
		ggoSpread.SSSetEdit   C_AcctNm,     "계정코드명",	16,	,	,	30
		ggoSpread.SSSetCombo  C_DrCrFg,     "", 8
	    	ggoSpread.SSSetCombo  C_DrCrNm,     "차대구분",		9
		ggoSpread.SSSetFloat  C_ItemAmt,    "금액",			15,	parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec
		ggoSpread.SSSetFloat  C_ItemLocAmt, "금액(자국)",	15,	parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec
		ggoSpread.SSSetEdit   C_ItemDesc,   "비  고",		60,	,	,	128
		ggoSpread.SSSetFloat  C_ExchRate,   "",					15,	parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec


		'20090922 kbs 
        	ggoSpread.SSSetFloat  C_ItemSeq2,    "",				4,	"6", ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec

		call ggoSpread.MakePairsColumn(C_Deptcd,C_DeptPopup)
		call ggoSpread.MakePairsColumn(C_Bizareacd,C_BizareaPopup)
		call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPopup)

		Call ggoSpread.SSSetColHidden(C_ItemSeq,C_ItemSeq,True)
		Call ggoSpread.SSSetColHidden(C_DrCrFg,C_DrCrFg,True)
		Call ggoSpread.SSSetColHidden(C_ExchRate,C_ExchRate,True)

		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

		.ReDraw = True

	End With

    Call SetSpreadLock ("I", 0, -1, -1)

End Sub

'========================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )

    Dim objSpread

    With frm1

		ggoSpread.SSSetRequired C_DrCrNm,  -1, -1		' 차대구분 
		ggoSpread.SSSetRequired C_ItemAmt, -1, -1		' 금액 

		Select Case Index
			Case 0
				ggoSpread.Source = .vspdData
				lRow2 = .vspdData.MaxRows
				.vspdData.Redraw = False

		        ggoSpread.SpreadLock  C_Bizareacd,		lRow,  C_Bizareacd,		lRow2
				ggoSpread.SpreadLock  C_BizareaPopup,	lRow,  C_BizareaPopup,	lRow2
				ggoSpread.SpreadLock  C_Bizareanm,		lRow,  C_Bizareanm,		lRow2
'		        ggoSpread.SpreadLock  C_Deptcd,			lRow,  C_Deptcd,		lRow2
'				ggoSpread.SpreadLock  C_DeptPopup,		lRow,  C_DeptPopup,		lRow2
				ggoSpread.SpreadLock  C_Deptnm,			lRow,  C_Deptnm,		lRow2
		        ggoSpread.SpreadLock  C_AcctCd,			lRow,  C_AcctCd,		lRow2
				ggoSpread.SpreadLock  C_AcctPopup,		lRow,  C_AcctPopup,		lRow2
				ggoSpread.SpreadLock  C_AcctNm,			lRow,  C_AcctNm,		lRow2
		        ggoSpread.SpreadLock  C_DrCrFg,			lRow,  C_DrCrFg,		lRow2
				ggoSpread.SpreadLock  C_DrCrNm,			lRow,  C_DrCrNm,		lRow2
				ggoSpread.SpreadLock  C_ItemAmt,		lRow,  C_ItemAmt,		lRow2
		        ggoSpread.SpreadLock  C_ItemLocAmt,		lRow,  C_ItemLocAmt,	lRow2
'				ggoSpread.SpreadLock  C_ItemDesc,		lRow,  C_ItemDesc,		lRow2
				ggoSpread.SpreadLock  C_ExchRate,		lRow,  C_ExchRate,		lRow2

		End Select
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		.vspdData.Redraw = True

    End With

End Sub


'=======================================================================================================
Sub SetSpread2Lock(Byval stsFg,Byval Index,ByVal lRow  ,ByVal lRow2 )
    
    With frm1
		ggoSpread.Source = .vspdData2			
		If lRow = "" Then
			lRow = 1
		End If	
		If lRow2 = "" Then
			lRow2 = .vspdData2.MaxRows
		End If
			
		.vspdData2.Redraw = False	
		Select Case Index
			Case 0			
			Case 1
				ggoSpread.SpreadLock 1, lRow, .vspdData2.MaxCols, lRow2	
		End Select		
		.vspdData2.Redraw = True
		    
    End With
End Sub

'========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

    With frm1

		.vspdData.ReDraw = False
		' 필수 입력 항목으로 설정 
		ggoSpread.SSSetProtected C_ItemSeq,   pvStartRow, pvEndRow		'
		ggoSpread.SSSetRequired  C_Bizareacd, pvStartRow, pvEndRow		' 사업장코드 
		ggoSpread.SSSetRequired  C_Deptcd,    pvStartRow, pvEndRow		' 부서코드 
		ggoSpread.SSSetRequired  C_AcctCd,    pvStartRow, pvEndRow		' 계정코드 
		ggoSpread.SSSetRequired  C_DrCrNm,    pvStartRow, pvEndRow		' 차대구분 
		ggoSpread.SSSetRequired  C_ItemAmt,   pvStartRow, pvEndRow		' 금액 
		ggoSpread.SSSetProtected C_Bizareanm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Deptnm,    pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_AcctNm,    pvStartRow, pvEndRow

		.vspdData.ReDraw = True

    End With

End Sub
'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to position of error
'               : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)

    Dim iDx
    Dim iRow

    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to
              Exit For
           End If
       Next
    End If

End Sub
'========================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc :
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ItemSeq      = iCurColumnPos(1)
			C_Bizareacd    = iCurColumnPos(2)
			C_BizareaPopup = iCurColumnPos(3)
			C_Bizareanm    = iCurColumnPos(4)
			C_Deptcd       = iCurColumnPos(5)
			C_DeptPopup    = iCurColumnPos(6)
			C_Deptnm       = iCurColumnPos(7)
			C_AcctCd       = iCurColumnPos(8)
			C_AcctPopup    = iCurColumnPos(9)
			C_AcctNm       = iCurColumnPos(10)
			C_DrCrFg       = iCurColumnPos(11)
			C_DrCrNm       = iCurColumnPos(12)
			C_ItemAmt      = iCurColumnPos(13)
			C_ItemLocAmt   = iCurColumnPos(14)
			C_ItemDesc     = iCurColumnPos(15)
			C_ExchRate     = iCurColumnPos(16)

			'20090922 kbs 
			C_ItemSeq2     = iCurColumnPos(17)
    End Select 
 
End Sub

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029										'⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")					'⊙: Lock  Suitable  Field	
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)

    Call InitSpreadSheet()									'⊙: Setup the Spread sheet    

    Call InitCtrlSpread()									'관리항목 그리드 초기화 

    Call InitCtrlHSpread()

    Call InitSpreadComboBox()

    Call SetDefaultVal()

	Call InitVariables()										'⊙: Initializes local global variables

    Call SetToolbar(C_MENU_NEW)						'⊙: 버튼 툴바 제어 
    frm1.txttempglno.focus
    frm1.txtCommandMode.value = "CREATE"
	ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.ClearSpreadData()

    '20090922 kbs
    'frm1.vspdData3.MaxCols = 16
     frm1.vspdData3.MaxCols = 17

End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================================
Function FncQuery()

    Dim IntRetCD
    Dim RetFlag

    On Error Resume Next
    Err.Clear
    lgstartfnc = True

    FncQuery = False												'⊙: Processing is NG

    ggoSpread.Source = frm1.vspdData

    If lgBlnFlgChgValue = True  and ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")	'데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      		Exit Function
     	End If
    End If

	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData3
	Call ggoSpread.ClearSpreadData()
	
	If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If

    Call InitVariables												'⊙: Initializes local global variables
    Call InitSpreadComboBox

    If DbQuery = False Then
		Exit Function
    End If															'☜: Query db data

    If frm1.vspddata.MaxRows = 0 Then
       frm1.txttempglno.value = ""
    End If

    If Err.number = 0 Then
       FncQuery = True												'☜: Processing is OK
    End If

    lgstartfnc = False
	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncNew()
	
	Dim IntRetCD

    On Error Resume Next											'☜: Protect system from crashing
    Err.Clear

    FncNew = False													'⊙: Processing is NG
    lgstartfnc = True

    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")	'데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If

	Call ggoOper.ClearField(Document, "1")						'⊙: Clear Condition Field
	Call ggoOper.ClearField(Document, "2")						'⊙: Clear Condition Field
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData3
	Call ggoSpread.ClearSpreadData()
	
    Call ggoOper.LockField(Document, "N")						'⊙: Lock  Suitable  Field
    Call InitVariables											'⊙: Initializes local global variables
    Call SetDefaultVal
    Call InitSpreadSheet()
    Call InitCtrlSpread()										'관리항목 그리드 초기화 
    Call InitCtrlHSpread()
    Call InitSpreadComboBox
    
    frm1.txttempglno.foucs
    frm1.txtCommandMode.value = "CREATE"

    Call SetToolbar(C_MENU_NEW)							'버튼 툴바 제어 
	Call ggoOper.SetReqAttr(frm1.txtDeptCd, "N")
	Call ggoOper.SetReqAttr(frm1.txtDocCur, "N")
	Call ggoOper.SetReqAttr(frm1.txtGlDt,   "N")

    If Err.number = 0 Then
		FncNew = True													'⊙: Processing is OK
    End If
    lgstartfnc = false
    lgFormLoad = True							' gldt read
    lgQueryOk = False

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncDelete()
	
	Dim IntRetCD

    On Error Resume Next
    Err.Clear															'☜: Protect system from crashing

    FncDelete = False													'⊙: Processing is NG

	intRetCd = DisplayMsgBox("990008", parent.VB_YES_NO, "X", "X")	'☜ 바뀐부분 
	If intRetCd = VBNO Then
		Exit Function
	End If

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900002", parent.VB_YES_NO, "X", "X")	'데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If

    frm1.txtCommandMode.value = "DELETE"

    If DbDelete = False Then
		Exit Function
	End If																'☜: Delete db data

	If Err.number = 0 Then
		FncDelete = True
	End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncSave()
    
    Dim IntRetCD

    On Error Resume Next												'☜: Protect system from crashing
    Err.Clear

    FncSave = False
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False  And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")		'No data changed!!
        Exit Function
    End If

    If CheckSpread3 = False Then
		IntRetCD = DisplayMsgBox("110420", "X", "X", "X")		'필수입력 check!!
        Exit Function
    End If

    If Not chkField(Document, "2") Then
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then							'⊙: Check contents area
       Exit Function
    End If

    If DbSave = False Then												'☜: Save db data
		Exit Function
	End If

	If Err.number  = 0 Then
		FncSave = True
    End If



	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncCopy()

	Dim IntRetCD
	Dim Indx

    On Error Resume Next												'☜: Protect system from crashing
    Err.Clear

	FncCopy = False

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
            .ReDraw = True
		    .Focus
		 End If
	End With
    If frm1.vspdData.MaxRows > 0 Then
		For Indx = frm1.vspdData.ActiveRow to frm1.vspdData.ActiveRow
			Call MaxSpreadVal(Indx)
		Next
	End If
	Call vspdData_Change(C_AcctCd, frm1.vspddata.activerow)

	If Err.number = 0 Then
		FncCopy = True
	End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncCancel()

    On Error Resume Next												'☜: Protect system from crashing
    Err.Clear

	FncCancel = False

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData

        .Row = .ActiveRow
        .Col = 0
        If .Text = ggoSpread.InsertFlag Then
            .Col = C_ItemSeq
            Call DeleteHSheet(.Text)
        End if

    End With

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo

    If frm1.vspdData.MaxRows < 1 Then
		Call SetToolbar(C_MENU_NEW)					'버튼 툴바 제어 
	End If

	If Err.number = 0 Then
		FncCancel = True
	End If

	Set gActiveElement = document.ActiveElement

End Function
'========================================================================================================
Function FncInsertRow(Byval pvRowCnt)

	Dim imRow
	Dim Indx
        Dim imRow2

    On Error Resume Next												'☜: Protect system from crashing
    Err.Clear

	FncInsertRow = False

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt) 
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End If

	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData

        For imRow2 = 1 To imRow  

          ggoSpread.InsertRow ,1
          SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	 .vspdData.ReDraw = True

                                  .vspdData.col = C_ItemDesc
                                  .vspddata.text	= .txtDesc.Value
                   Next
	End With
  	frm1.vspdData.Row = .vspdData.ActiveRow
	frm1.vspdData.Action = 0
	If frm1.vspdData.MaxRows > 0 Then
		For Indx = frm1.vspdData.ActiveRow to frm1.vspdData.ActiveRow + imRow - 1

			Call MaxSpreadVal(Indx)
		Next
	End If

	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ClearSpreadData()
	Call SetToolbar("1110110100111111")					'버튼 툴바 제어 

	If Err.number = 0 Then
		FncInsertRow = True
	End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncDeleteRow()

	Dim lDelRows
	Dim iDelRowCnt
    Dim DelItemSeq

    On Error Resume Next												'☜: Protect system from crashing
    Err.Clear

	FncDeleteRow = False

	With frm1.vspdData 

		ggoSpread.Source = frm1.vspdData
		.Row = .ActiveRow
		.Col = 0

		If frm1.vspdData.MaxRows < 1 Or .Text = ggoSpread.InsertFlag Then Exit Function

		.Col = 1
		DelItemSeq = .Text
		lDelRows = ggoSpread.DeleteRow

    End With

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.ClearSpreadData()
    Call DeleteHsheet(DelItemSeq)

    If Err.number = 0 Then
		FncDeleteRow = True
	End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncPrev() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrev = False                                                               '☜: Processing is NG
    If Err.number = 0 Then
       FncPrev = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function
'========================================================================================================
Function FncNext() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNext = False                                                               '☜: Processing is NG
    If Err.number = 0 Then
       FncNext = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(Parent.C_MULTI)

    If Err.number = 0 Then
       FncExcel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function
'========================================================================================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	Call Parent.FncFind(Parent.C_MULTI, True)

    If Err.number = 0 Then
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function
'========================================================================================================
' Function Name : FncSplitColumn
' Function Desc :
'========================================================================================================
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
	Dim indx

	On Error Resume Next
	Err.Clear 		

	ggoSpread.Source = gActiveSpdSheet
    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA"		
			Call PrevspdDataRestore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()						
			Call InitSpreadSheet()
            Call InitComboBoxGrid      
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData()			
            'Call SetSpreadColor("Q", 0,1, .vspddata.MaxRows)                 
            Call SetSpread2Color() 		                
		Case "VSPDDATA2"
			Call PrevspdData2Restore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()			
			Call InitCtrlSpread()
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData()
			Call SetSpread2Color()
	End Select
	
	If frm1.vspdData2.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If

End Sub

'========================================================================================================
' Function Name : PrevspdDataRestore
' Function Desc :
'========================================================================================================
Sub PrevspdDataRestore(pActiveSheetName)

	Dim indx, indx1

	For indx = 0 to frm1.vspdData.MaxRows
        	frm1.vspdData.Row    = indx
        	frm1.vspdData.Col    = 0
		
		If frm1.vspdData.Text <> "" Then
			Select Case frm1.vspdData.Text			
				Case ggoSpread.InsertFlag					
					frm1.vspdData.Col = C_ItemSeq					
					Call DeleteHsheet(frm1.vspdData.Text)					
				Case ggoSpread.UpdateFlag		
					For indx1 = 0 to frm1.vspdData3.MaxRows					
						frm1.vspdData3.Row = indx1
						frm1.vspdData3.Col = 0
						Select Case frm1.vspdData3.Text 
							Case ggoSpread.UpdateFlag
								frm1.vspdData.Col = C_ItemSeq
								frm1.vspdData3.Col = 1					
								If UCase(Trim(frm1.vspdData.Text)) = UCase(Trim(frm1.vspdData3.Text)) Then
									Call DeleteHsheet(frm1.vspdData.Text)										
									Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.htxtTempGlNo.value)
								End If
						End Select
					Next
					'ggoSpread.Source = frm1.vspdData					
					'ggoSpread.EditUndo
					
				Case ggoSpread.DeleteFlag
					Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.htxtTempGlNo.value)
					'ggoSpread.Source = frm1.vspdData
					'ggoSpread.EditUndo

			End Select
			
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName

End Sub

'========================================================================================================
' Function Name : PrevspdData2Restore
' Function Desc :
'========================================================================================================
Sub PrevspdData2Restore(pActiveSheetName)

	Dim indx, indx1

	For indx = 0 to frm1.vspdData2.MaxRows
        frm1.vspdData2.Row    = indx
        frm1.vspdData2.Col    = 0

		If frm1.vspdData2.Text <> "" Then
			Select Case frm1.vspdData2.Text
				Case ggoSpread.InsertFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData	
					        ggoSpread.EditUndo							
						End If
					Next
				Case ggoSpread.UpdateFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData
							ggoSpread.EditUndo
							Call fncRestoreDbQuery2(indx1, frm1.vspdData.ActiveRow, frm1.htxtTempGlNo.value)
						End If
					Next

				Case ggoSpread.DeleteFlag

			End Select
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName

End Sub

'========================================================================================================
' Name : fncRestoreDbQuery2																				
' Desc : This function is data query and display												
'========================================================================================================
Function fncRestoreDbQuery2(Row, CurrRow, Byval pInvalue1)

	Dim strItemSeq
	Dim strSelect, strFrom, strWhere
	Dim arrTempRow, arrTempCol
	Dim Indx1
	Dim strTableid, strColid, strColNm, strMajorCd
	Dim strNmwhere
	Dim arrVal
	Dim strVal

	on Error Resume Next
	Err.Clear

	fncRestoreDbQuery2 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
	With frm1

		.vspdData.row = Row
	    .vspdData.col = C_ItemSeq
		strItemSeq    = .vspdData.Text
	    If Trim(strItemSeq) = "" Then
	        Exit Function
	    End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM, LTrim(ISNULL(C.CTRL_VAL, '')), '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END, D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID, '')), LTrim(ISNULL(A.DATA_COLM_ID, '')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM, '')), LTrim(ISNULL(A.COLM_DATA_TYPE, '')), LTrim(ISNULL(A.DATA_LEN, '')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END,  " & strItemSeq & ", " 
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "

		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_TEMP_GL_DTL C (NOLOCK), A_TEMP_GL_ITEM D (NOLOCK) "

		strWhere =			  " D.TEMP_GL_NO = " & FilterVar(UCase(pInvalue1), "''", "S")   
		strWhere = strWhere & " AND D.ITEM_SEQ = " & strItemSeq & " "
		strWhere = strWhere & " AND D.TEMP_GL_NO  =  C.TEMP_GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD *= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD *= B.CTRL_CD "
		strWhere = strWhere & " AND C.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "
		

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
			arrTempRow =  Split(lgF2By2, Chr(12))
			For Indx1 = 0 To Ubound(arrTempRow) - 1
				arrTempCol = split(arrTempRow(indx1), Chr(11))
				If Trim(arrTempCol(8)) <> "" Then
					strTableid = arrTempCol(8)
					strColid   = arrTempCol(9)
					strColNm   = arrTempCol(10)
					strMajorCd = arrTempCol(15)
					
					strNmwhere = strColid & " =   " & FilterVar(arrTempCol(C_CtrlVal), "''", "S") & "  " 

					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If

					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						arrVal = Split(lgF0, Chr(11))
						arrTempCol(6) = arrVal(0)
					End If
				End If

				strVal = strVal & Chr(11) & strItemSeq
				strVal = strVal & Chr(11) & arrTempCol(1)
				strVal = strVal & Chr(11) & arrTempCol(2)
				strVal = strVal & Chr(11) & arrTempCol(3)
				strVal = strVal & Chr(11) & arrTempCol(4)
				strVal = strVal & Chr(11) & arrTempCol(5)
				strVal = strVal & Chr(11) & arrTempCol(6)
				strVal = strVal & Chr(11) & arrTempCol(7)
				strVal = strVal & Chr(11) & arrTempCol(8)
				strVal = strVal & Chr(11) & arrTempCol(9)
				strVal = strVal & Chr(11) & arrTempCol(10)
				strVal = strVal & Chr(11) & arrTempCol(11)
				strVal = strVal & Chr(11) & arrTempCol(12)
				strVal = strVal & Chr(11) & arrTempCol(13)
				strVal = strVal & Chr(11) & arrTempCol(15)
				strVal = strVal & Chr(11) & Indx1 + 1
				strVal = strVal & Chr(11) & Chr(12)
			Next
			ggoSpread.Source = .vspdData3
			ggoSpread.SSShowData strVal	
		End If 		

		If Row = CurrRow Then
			Call CopyFromData (strItemSeq)
		End If

		Call LayerShowHide(0)
		Call RestoreToolBar()
	End With

	If Err.number = 0 Then
		fncRestoreDbQuery2 = True
	End If
End Function
'========================================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================================
Function FncExit()

	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	FncExit = False

    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True OR ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End if

    If Err.number = 0 Then
       FncExit = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================
Function FncBtnPreview() 
    Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId
    Dim StrUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim intRetCD
	Dim ObjName
    
    If Not chkField(Document, "1") Then
		Exit Function
    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId)

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|TempGlNoFr|" & VarTempGlNoFr
	StrUrl = StrUrl & "|TempGlNoTo|" & VarTempGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	StrUrl = StrUrl & "|GlPutType|" & "%"

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)
End Function

'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
	Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId
    Dim StrEbrFile
    Dim intRetCd
	Dim ObjName

	If Not chkField(Document, "1") Then
       Exit Function
    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId)

    lngPos = 0

	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|TempGlNoFr|" & VarTempGlNoFr
	StrUrl = StrUrl & "|TempGlNoTo|" & VarTempGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	StrUrl = StrUrl & "|GlPutType|" & "%"

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")	
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
End Function

'=======================================================================================================
Sub SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId)
	Dim intRetCd

	StrEbrFile = "a8101ma1"
	VarDateFr = UniConvDateToYYYYMMDD(frm1.txtGlDt.Text, parent.gDateFormat, parent.gServerDateType)	
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txtGlDt.Text, parent.gDateFormat, parent.gServerDateType)

	' 회계전표의 key는 GL_NO이기 때문에 GL_NO만 넘긴다.	
	VarDeptCd = "%"
	VarBizAreaCd = "%"
	VarTempGlNoFr = Trim(frm1.txttempGlNo.value)
	VarTempGlNoTo = Trim(frm1.txtHqBrchNo.value)
	varOrgChangeId = Trim(frm1.hOrgChangeId.value)
End Sub

'========================================================================================================
Function DbQuery()

	Dim strVal
	Dim RetFlag

	On Error Resume Next
	Err.Clear

    DbQuery = False

    Call DisableToolBar(parent.TBC_QUERY)
    Call LayerShowHide(1)

    With frm1

	    If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txttempglno=" & UCase(Trim(.htxttempglno.value))	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
			strVal = strVal & "&txttempglno=" & UCase(Trim(.txttempglno.value))		'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    End With

	If Err.number = 0 Then
       DbQuery = True																'☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=======================================================================================================
Function DbQuery2(ByVal Row)

	Dim strVal
	Dim lngRows
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim strTableid
	Dim strColid
	Dim strColNm
	Dim strMajorCd
	Dim strNmwhere
	Dim indx,Indx1
	Dim arrVal
	Dim arrTemp

	On Error Resume Next
    Err.Clear

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)

	DbQuery2 = False

	With frm1

		.htxttempglno.Value = frm1.txttempglno.Value

	    .vspdData.Row = Row
	    .vspdData.Col = C_ItemSeq
	    .hItemSeq.Value = .vspdData.Text

	    If Trim(.hItemSeq.Value) = "" Then
			Call LayerShowHide(0)
			Call RestoreToolBar()
	    	Exit Function
	    End If
        If CopyFromData(.hItemSeq.Value) = True Then
			If lgIntFlgMode = parent.OPMD_UMODE Then
				Call SetSpread2Lock("",1,1,"")
			Else
				Call SetSpread2Color()
			End If	
			Call RestoreToolBar()
			Call LayerShowHide(0)

			Exit Function
        End If
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq

		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM, LTrim(ISNULL(C.CTRL_VAL, '')), '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END, D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID, '')), LTrim(ISNULL(A.DATA_COLM_ID, '')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM, '')), LTrim(ISNULL(A.COLM_DATA_TYPE, '')), LTrim(ISNULL(A.DATA_LEN, '')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END,  " & .hItemSeq.Value & ", " 
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "

		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_TEMP_GL_DTL C (NOLOCK), A_TEMP_GL_ITEM D (NOLOCK) "

		strWhere =			  " D.TEMP_GL_NO = " & FilterVar(UCase(.htxtTempGlNo.value), "''", "S")
		strWhere = strWhere & " AND D.ITEM_SEQ = " & .hItemSeq.Value
		strWhere = strWhere & " AND D.TEMP_GL_NO  =  C.TEMP_GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD *= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD *= B.CTRL_CD "
		strWhere = strWhere & " AND C.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "
	
		

		frm1.vspdData2.ReDraw = False
		If CommonQueryRs2by2(strSelect, strFrom, strWhere, lgF2By2) Then
			ggoSpread.Source = frm1.vspdData2
			arrTemp =  Split(lgF2By2,Chr(12))
			For Indx1 = 0 To Ubound(arrTemp) - 1
				arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
			Next
			lgF2By2 = Join(arrTemp,Chr(12))
			ggoSpread.SSShowData lgF2By2

			For lngRows = 1 to frm1.vspdData2.MaxRows
				frm1.vspdData2.Row = lngRows
				frm1.vspdData2.Col = C_Tableid
				If Trim(frm1.vspdData2.Text) <> "" Then
					frm1.vspdData2.Col = C_Tableid
					strTableid = frm1.vspdData2.Text
					frm1.vspdData2.Col = C_Colid
					strColid = frm1.vspdData2.Text
					frm1.vspdData2.Col = C_ColNm
					strColNm = frm1.vspdData2.Text
					frm1.vspdData2.Col = C_MajorCd
					strMajorCd = frm1.vspdData2.Text
					frm1.vspdData2.Col = C_CtrlVal

					strNmwhere = strColid & " =  " & FilterVar(frm1.vspddata2.Text, "''", "S")

					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD = " & FilterVar(strMajorCd, "''", "S")
					End If

					If CommonQueryRs(strColNm, strTableid, strNmwhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
						frm1.vspdData2.Col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))
						frm1.vspdData2.Text = arrVal(0)
					End If
				End If

				strVal = strVal & Chr(11) & .hItemSeq.Value
				
				.vspdData2.Col = C_DtlSeq
				strVal = strVal & Chr(11) & .vspdData2.Text
                
				.vspdData2.Col = C_CtrlCd
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_CtrlNm
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_CtrlVal
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_CtrlPB
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_CtrlValNm
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_Seq
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_Tableid
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_Colid
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_ColNm
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_Datatype
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_DataLen
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_DRFg
				strVal = strVal & Chr(11) & .vspdData2.Text
				    
				.vspdData2.Col = C_MajorCd
				strVal = strVal & Chr(11) & .vspdData2.Text
				
				.vspdData2.Col = C_MajorCd + 1
				.vspdData2.Text = lngRows
				strVal = strVal & Chr(11) & .vspdData2.Text
				
				strVal = strVal & Chr(11) & Chr(12)
				
			Next

			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSShowData strVal

		End If
		intItemCnt = .vspddata.MaxRows
		If lgIntFlgMode = parent.OPMD_UMODE Then
			Call SetSpread2Lock("",1,1,"")
		Else
			Call SetSpread2Color()
		End If	
	End With

	frm1.vspdData2.ReDraw = True

	Call RestoreToolBar()
	Call LayerShowHide(0)

	If Err.number = 0 Then
		DbQuery2 = True
	End If

	Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
Function DbSave()

    Dim pAP010M
    Dim lngRows, itemRows
    Dim lGrpcnt
    DIM strVal
    Dim strDel
    Dim tempItemSeq
    Dim bChkBizArea
    Dim tmpBizArea   
    Dim TempItemSeq2


    On Error Resume Next
    Err.Clear

    DbSave = False
	bChkBizArea = False
	
    Call DisableToolBar(parent.TBC_SAVE)
    Call LayerShowHide(1)

    Call SetSumItem()

	With frm1
		.txtFlgMode.value     = lgIntFlgMode
		.txtUpdtUserId.value  = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID

		If UCase(Trim(frm1.txtCommandMode.value)) = "CREATE"  Then
			.txtMode.value = parent.UID_M0002
		ElseIf  UCase(Trim(frm1.txtCommandMode.value)) = "UPDATE" Then
			.txtMode.value = parent.UID_M0004
		Else
			.txtMode.value = parent.UID_M0003
		End If
		.txtOrgChangeId.value	= parent.gChangeOrgId
		.txtgCurrency.value	= parent.gCurrency
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data 연결 규칙 
    lGrpCnt = 1
    strVal = ""
    strDel = ""

    ggoSpread.Source = frm1.vspdData

    With frm1.vspdData

		For lngRows = 1 to .MaxRows
			.Row = lngRows
			.Col = 0

			If .Text = ggoSpread.InsertFlag Then
				strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep		'C=Create, Sheet가 2개 이므로 구별 
			ElseIf .Text = ggoSpread.UpdateFlag Then
				strVal = strVal & "U" & parent.gColSep & lngRows & parent.gColSep		'U=Update
			ElseIf .Text = ggoSpread.DeleteFlag Then
				strDel = strDel & "D" & parent.gColSep & lngRows & parent.gColSep		'D=Delete
			ELSE
				strVal = strVal & "U" & parent.gColSep & lngRows & parent.gColSep		'U=Update
			End If

			Select Case .Text
			    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag


				'20090922 kbs 
				TempItemSeq2 = lngRows

			        .Col  = C_ItemSeq2
			        .Text = lngRows

			        .Col 	   = C_ItemSeq
				tempItemSEq = frm1.vspdData.Text

				Call RepDtlSeq(tempItemSEq, TempItemSeq2)
			       '.Col = C_ItemSeq	'1
			        .Col = C_ItemSeq2	'1
			        strVal = strVal & Trim(.Text) & parent.gColSep


			        .Col = C_Bizareacd	'2
					' Check If all the bizarea are the same 
			        If lngRows = 1	 Then
				        tmpBizArea = UCase(Trim(.Text))
			        End If
			        
			        If tmpBizArea <> UCase(Trim(.Text)) Then
						bChkBizArea =True
			        End If
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_Deptcd	    '3
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_AcctCd		'4
			        strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_DrCrFG		'5
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_ItemAmt	'6
			        strVal = strVal & Trim(.Text) & parent.gColSep
   			        .Col = C_ItemLocAmt	'7
			        strVal = strVal & Trim(.Text) & parent.gColSep
                                
                                '------------------------------------
                                ' 상단의 비고 Grid 의 대변계정의 비고항목으로 복사
                                '------------------------------------
                            '    strVal = strVal & Trim(frm1.txtDesc.Value) & parent.gColSep
                                .Col = C_ItemDesc	'8
			        strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_ExchRate	'9
			        strVal = strVal & Trim(.Text) & parent.gRowSep

			        lGrpCnt = lGrpCnt + 1
			        
			    Case ggoSpread.DeleteFlag
			        .Col = C_ItemSeq	'1
			        strDel = strDel & Trim(.Text) & parent.gRowSep						'마지막 데이타는 Row 분리기호를 넣는다 
					lGrpcnt = lGrpcnt + 1

				CASE ELSE
				    .Col = C_ItemSeq	'1
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_Bizareacd	'2
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_Deptcd	    '3
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_AcctCd		'4
			        strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_DrCrFG		'5
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_ItemAmt	'6
			        strVal = strVal & Trim(.Text) & parent.gColSep
   			        .Col = C_ItemLocAmt	'7
			        strVal = strVal & Trim(.Text) & parent.gColSep

			        .Col = C_ItemDesc	'8
			        strVal = strVal & Trim(.Text) & parent.gColSep
                                'strVal = strVal & Trim(frm1.txtDesc.Value) & parent.gColSep
					.Col = C_ExchRate	'9
			        strVal = strVal & Trim(.Text) & parent.gRowSep
			        lGrpCnt = lGrpCnt + 1

			End Select

		Next

    End With


    frm1.txtMaxRows.value = lGrpCnt-1													'Spread Sheet의 변경된 최대갯수 
    frm1.txtSpread.value =  strDel & strVal												'Spread Sheet 내용을 저장 

	If bChkBizArea = False And lgIntFlgMode = parent.OPMD_CMODE Then
		IntRetCD = DisplayMsgBox("124203", "X", "X", "X")		'No data changed!!
		Call SetToolbar(C_MENU_NEW)	
	    Call LayerShowHide(0)
		Exit Function
	End If
    lGrpCnt = 1
    strVal = ""
    strDel = ""

    ggoSpread.Source = frm1.vspdData3

    With frm1.vspdData3      ' Dtl 저장 
    
		For itemRows = 1 to frm1.vspdData.MaxRows
 		    frm1.vspdData.Row = itemRows
		    frm1.vspdData.Col = 0

		    if frm1.vspdData.Text = ggoSpread.DeleteFlag Or _
 		       frm1.vspdData.Text = ggoSpread.UpdateFlag Or _
		       frm1.vspdData.Text = ggoSpread.InsertFlag Then

			'20090922  kbs
		       'frm1.vspdData.Col = 1
		        frm1.vspdData.Col = C_ItemSeq
			tempItemSEq	  = frm1.vspdData.Text

			'20090922  kbs
		        frm1.vspdData.Col = C_ItemSeq2
			tempItemSEq2	  = frm1.vspdData.Text


			    For lngRows = 1 to .MaxRows
					.Row = lngRows
					.Col = 1
					If .Text = tempitemseq Then
					    .Col = 0
						Select Case .Text
						    Case ggoSpread.DeleteFlag
								strDel = strDel & "D" & parent.gColSep
								.Col = 1 		
						    '   strDel = strDel & Trim(.Text) & parent.gColSep
						        strDel = strDel & tempItemSEq & parent.gColSep

								.Col =  .Col + 1   			'Dtl SEQ
						        strDel = strDel & Trim(.Text) & parent.gRowSep
						        lGrpCnt = lGrpCnt + 1

					        Case Else
							strVal = strVal & "C" & parent.gColSep

						        .Col = 1 		 			'ItemSEQ
						    '   strVal = strVal & Trim(.Text) & parent.gColSep

							'20090922 kbs
						        'strVal = strVal & tempitemseq  & parent.gColSep
						         strVal = strVal & tempitemseq2 & parent.gColSep

						        .Col = .Col + 1 			'Dtl SEQ
						        strVal = strVal & Trim(.Text) & parent.gColSep
								.Col = .Col + 1		 		'관리항목코드 
						        strVal = strVal & Trim(.Text) & parent.gColSep
						        .Col = .Col + 2				'관리항목 Value
						        strVal = strVal & Trim(.Text) & parent.gRowSep
								lGrpCnt = lGrpCnt + 1

					    End Select
					End If
			    Next
			End If
		Next

    End With

    frm1.txtMaxRows3.value = lGrpCnt-1													'Spread Sheet의 변경된 최대갯수 
    frm1.txtSpread3.value  = strDel & strVal
											'Spread Sheet 내용을 저장 


    Call ExecMyBizASP(frm1, BIZ_PGM_ID)													'저장 비지니스 ASP 를 가동 

	If Err.number = 0 Then
       DbSave = True																	'☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function


'20090922 kbs
Function RepDtlSeq(byval tempitemseq, byval tempitemseq2 )
   Dim lngRows 

    ggoSpread.Source = frm1.vspdData3

    With frm1.vspdData3      ' Dtl 저장 

	    For lngRows = 1 to .MaxRows
		.Row = lngRows
		.Col = 1
		If .Text = tempitemseq Then
		    .Col = 0

		    Select Case .Text
			    Case ggoSpread.DeleteFlag
				.Col  = C_HItemSeq2
				.Text = tempitemseq2

		        Case Else
				.Col  = C_HItemSeq2
				.Text = tempitemseq2

		    End Select
		End If

	    Next

    End With

End Function 

'========================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================================
Function DbDelete()

	Dim strVal

	On Error Resume Next
    Err.Clear

    DbDelete = False														'⊙: Processing is NG

    Call DisableToolBar(parent.TBC_DELETE)
    Call LayerShowHide(1)

	frm1.txtOrgChangeId.value = parent.gChangeOrgId

	strVal = BIZ_PGM_ID & "?txtMode="		 & parent.UID_M0003					'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal		& "&txttempglno="	 & UCase(Trim(frm1.txttempglno.value))	'☜: 삭제 조건 데이타 
    strVal = strVal		& "&txtDeptCd="		 & UCase(Trim(frm1.txtDeptCd.value))
	strVal = strVal		& "&txtOrgChangeId=" & Trim(frm1.hOrgChangeId.value)
    strVal = strVal		& "&txtGlinputType=" & Trim(frm1.txtGlinputType.value)
    strVal = strVal		& "&txtTempGlDt="	 & Trim(frm1.txtGlDt.text)

	Call RunMyBizASP(MyBizASP, strVal)

    If Err.number = 0 Then
       DbDelete = True                                                             '☜: Processing is OK
    End If

	Set gActiveElement = document.ActiveElement

End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'========================================================================================================
Sub DbQueryOk()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	With frm1
		.vspdData.Col = 1
		intItemCnt = .vspddata.MaxRows
		Call SetSpreadLock ("Q", 0, -1, -1)
        lgIntFlgMode = parent.OPMD_UMODE									'Indicates that current mode is Update mode
        Call ggoOper.LockField(Document, "I")								'This function lock the suitable field
	    Call SetToolbar("1111100000011111")									'버튼 툴바 제어 
	    
        .txtCommandMode.value = "UPDATE"
        Call InitData()
        Call SetSumItem()

        If .vspdData.MaxRows > 0 Then
            .vspdData.Row = 1
            .vspdData.Col = 1
            .hItemSeq.Value = .vspdData.Text

 			Call ggoOper.SetReqAttr(frm1.txtDeptCd, "Q")
			Call ggoOper.SetReqAttr(frm1.txtDocCur, "Q")
			Call ggoOper.SetReqAttr(frm1.txtGlDt,   "Q")

            Call DbQuery2(1)
        End If

    End With
    
    lgBlnFlgChgValue = False                    'Indicates that no value changed

'    Call InitVariables()
	Call QueryDeptCd_OnChange()

    Set gActiveElement = document.ActiveElement

End Sub
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Sub DbSaveOk(Byval GlNo)					'☆: 저장 성공후 실행 로직 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	frm1.txttempglno.value = UCase(Trim(GlNo))
    frm1.txtCommandMode.value = "UPDATE"
	Call SetToolbar("1110000000001111")							'버튼 툴바 제어 
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ClearSpreadData()
	ggoSpread.Source = frm1.vspdData3
	Call ggoSpread.ClearSpreadData()
	
    Call InitVariables()														'⊙: Initializes local global variables

	If DbQuery = False Then
       Call RestoreToolBar()
       Exit Sub
    End If

    Set gActiveElement = document.ActiveElement

End Sub
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Sub DbDeleteOk()												'삭제 성공후 실행 로직 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	Call MainNew()

	Set gActiveElement = document.ActiveElement

End Sub

'========================================================================================================
Function OpenRefTempGl()

	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(3)

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("A8101RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A8101RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "0" Then
		Call SetRefTempGl(arrRet)
	End If
	frm1.txtTempGlNo.focus

End Function
'=======================================================================================================
' Function Name : SetRefTempGl
' Function Desc :
'=======================================================================================================
Function SetRefTempGl(ByVal arrRet)

	With frm1
		If Trim(arrRet(0)) <> "" Then
			.txtTempGlNo.Value = UCase(Trim(arrRet(0)))
		End If
    End With

	
End Function
'=======================================================================================================
' Function Name : OpenPopUp
' Function Desc :
'=======================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strBizAreaCd

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.txtOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
		Case 1
			If UCase(frm1.txtDeptCd.className) = "PROTECTED" Then Exit Function
			arrParam(0) = "부서 팝업"				' 팝업 명칭 
			arrParam(1) = "B_ACCT_DEPT"    				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "ORG_CHANGE_ID = " & FilterVar(frm1.txtOrgChangeId.value, "''", "S") 		' Where Condition
			arrParam(5) = "부서코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "DEPT_CD"	     				' Field명(0)
			arrField(1) = "DEPT_NM"			    		' Field명(1)

			arrHeader(0) = "부서코드"				' Header명(0)
			arrHeader(1) = "부서명"				    ' Header명(1)

		Case 2
			If UCase(frm1.txtDocCur.className) = "PROTECTED" Then Exit Function
			arrParam(0) = "통화코드 팝업"			' 팝업 명칭 
			arrParam(1) = "B_Currency"	    			' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "통화코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "Currency"	    			' Field명(0)
			arrField(1) = "Currency_desc"	    		' Field명(1)

			arrHeader(0) = "통화코드"				' Header명(0)
			arrHeader(1) = "통화코드명"				' Header명(1)

		Case 3
			arrParam(0) = "사업장팝업"				' 팝업 명칭 
			arrParam(1) = "b_biz_area B, A_ACCT A"		' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.REL_BIZ_AREA_CD = B.BIZ_AREA_CD" ' Where Condition
			arrParam(5) = "사업장코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "B.BIZ_AREA_CD"				' Field명(0)
			arrField(1) = "B.BIZ_AREA_NM"				' Field명(1)

			arrHeader(0) = "사업장코드"				' Header명(0)
			arrHeader(1) = "사업장명"				' Header명(1)

		Case 4
			' 선택한 사업장에 속한 부서만 PopUp
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = C_Bizareacd
			strBizAreaCd = Trim(frm1.vspdData.text)
			If strBizAreaCd = "" then
				strBizAreaCd = "%"
			End If
		
			frm1.vspdData.Col = C_deptcd

			arrParam(0) = "부서코드팝업"			' 팝업 명칭 
			arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C, A_ACCT D "    				' TABLE 명칭 
			arrParam(2) = frm1.vspdData.text						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.ORG_CHANGE_ID = (select distinct org_change_id" & _
			              " from b_acct_dept where org_change_dt = ( select max(org_change_dt)" & _
			              " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))" & _
			              " and C.biz_area_cd LIKE " & FilterVar(strBizAreaCd, "''", "S")  & _
			              " AND B.cost_cd = A.cost_cd " & _
			              " AND C.biz_area_cd = B.biz_area_cd AND D.REL_BIZ_AREA_CD = C.BIZ_AREA_CD"
			arrParam(5) = "부서코드"				' 조건필드의 라벨 명칭 
			
			arrField(0) = "A.DEPT_CD"	     				' Field명(0)
			arrField(1) = "A.DEPT_NM"			    		' Field명(1)
			arrField(2) = "C.BIZ_AREA_CD"			    		' Field명(2)
			arrField(3) = "C.BIZ_AREA_NM"			    		' Field명(3)
    
			arrHeader(0) = "부서코드"				' Header명(0)
			arrHeader(1) = "부서명"				    ' Header명(1)						
			arrHeader(2) = "사업장코드"				    ' Header명(2)		
			arrHeader(3) = "사업장명"				    ' Header명(3)	

		    ' Header명(1)

		Case 5
			arrParam(0) = "계정코드팝업"			' 팝업 명칭 
			arrParam(1) = "A_Acct" 						' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "계정코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "Acct_CD"						' Field명(0)
			arrField(1) = "Acct_NM"						' Field명(1)

			arrHeader(0) = "계정코드"				' Header명(0)
			arrHeader(1) = "계정코드명"				' Header명(1)
	End Select

    If iWhere = 0 Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPopUp(arrRet, iWhere)
	End If	
	Call FocusAfterPopup (iWhere)

End Function

'========================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	
	With frm1

		Select Case iWhere
			Case 0
				.txttempglno.value = UCase(Trim(arrRet(0)))
				lgBlnFlgChgValue = True

			Case 1
				.txtDeptCd.value = UCase(Trim(arrRet(0)))
				.txtDeptNm.value = arrRet(1)
				lgBlnFlgChgValue = True

			Case 2
				.txtDocCur.value = UCase(Trim(arrRet(0)))
				lgBlnFlgChgValue = True
				Call txtDocCur_OnChange()
			Case 3
				frm1.vspdData.Row = frm1.vspdData.ActiveRow
				.vspdData.Col  = C_Bizareacd
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_Bizareanm
				.vspdData.Text = arrRet(1)

			Case 4
				frm1.vspdData.Row = frm1.vspdData.ActiveRow
				.vspdData.Col  = C_Deptcd
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_Deptnm
				.vspdData.Text = arrRet(1)
				If Trim(.vspdData.Text) = "" Then
					.vspdData.Text = arrRet(2)
					.vspdData.Col  = C_bizareanm
					.vspdData.Text = arrRet(3)
				End IF 
				Call deptCd_underChange(arrRet(0))

			Case 5
				frm1.vspdData.Row = frm1.vspdData.ActiveRow
				.vspdData.Col  = C_AcctCD
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_AcctNm
				.vspdData.Text = arrRet(1)

				Call vspdData_Change(C_AcctCd, frm1.vspddata.activeRow )

		End Select

	End With

End Function
'=======================================================================================================
Function FocusAfterPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txttempglno.focus
			Case 1  
				.txtDeptCd.focus
			Case 2 
				.txtDocCur.focus
			Case 3
				Call SetActiveCell(.vspdData,C_Bizareacd,.vspdData.ActiveRow ,"M","X","X")
			Case 4
				Call SetActiveCell(.vspdData,C_Deptcd,.vspdData.ActiveRow ,"M","X","X")
			Case 5
				Call SetActiveCell(.vspdData,C_AcctCD,.vspdData.ActiveRow ,"M","X","X")

		End Select    
	End With

End Function
'========================================================================================================
Function OpenDept(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strBizAreaCd

	If frm1.txtDeptCd.readOnly = true then
		IsOpenPop = False
		Exit Function
	End If
	If IsOpenPop = True Then Exit Function

	If lgQueryOk <> True then

		iCalledAspName = AskPRAspName("DeptPopupDtA2")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
			IsOpenPop = False
			Exit Function
		End If

		IsOpenPop = True

		arrParam(0) = strCode		            '  Code Condition
		arrParam(1) = frm1.txtGLDt.Text
		arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  

	' T : protected F: 필수 
		arrParam(3) = "F"									' 결의일자 상태 Condition  
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	Else

		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_Bizareacd
		strBizAreaCd = Trim(frm1.vspdData.text)
		If strBizAreaCd = "" then
			strBizAreaCd = "%"
		End If
		
		frm1.vspdData.Col = C_deptcd

		arrParam(0) = "부서코드팝업"			' 팝업 명칭 
		arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C, A_ACCT D "    				' TABLE 명칭 
		arrParam(2) = strCode						' Code Condition
		arrParam(3) = ""							' Name Cindition
		arrParam(4) = "A.ORG_CHANGE_ID = (select distinct org_change_id" & _
		              " from b_acct_dept where org_change_dt = ( select max(org_change_dt)" & _
		              "  from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))" & _
		              " and C.biz_area_cd LIKE " & FilterVar(strBizAreaCd, "''", "S")  & _
		              " AND B.cost_cd = A.cost_cd " & _
		              " AND C.biz_area_cd = B.biz_area_cd AND D.REL_BIZ_AREA_CD = C.BIZ_AREA_CD"
		arrParam(5) = "부서코드"				' 조건필드의 라벨 명칭 
			
		arrField(0) = "A.DEPT_CD"	     				' Field명(0)
		arrField(1) = "A.DEPT_NM"			    		' Field명(1)
		arrField(2) = "A.INTERNAL_CD"			    		' Field명(1)
    
		arrHeader(0) = "부서코드"				' Header명(0)
		arrHeader(1) = "부서명"				    ' Header명(1)						
		arrHeader(2) = "내부부서코드"				    ' Header명(1)	
				
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	
	End If

	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetDept(arrRet, iWhere)
	End If	
	frm1.txtDeptCd.focus
			
End Function
'========================================================================================================
Function SetDept(Byval arrRet, Byval iWhere)
			
	With frm1
		Select Case iWhere
		Case "0"
			.txtDeptCd.value = arrRet(0)
			.txtDeptNm.value = arrRet(1)
			.txtOrgChangeId.value = arrRet(2)
			If lgQueryOk <> True Then
					   .txtGLDt.text = arrRet(3)
			Else 

			End If    
			call txtDeptCd_OnChange()  
							
		Case Else
		End Select
	End With
End Function       

'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtDrAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtCrAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec		
	End With

End Sub

'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_ItemAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec		
		
	End With

End Sub

'========================================================================================================
'	Name : MaxSpreadVal
'	Description :개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'========================================================================================================
Function MaxSpreadVal(byval Row)
  
  Dim iRows
  Dim MaxValue
  Dim tmpVal

	MAxValue = 0
	
	With frm1
		For iRows = 1 to  .vspdData.MaxRows
			.vspddata.Row = iRows
		    .vspddata.Col = C_ItemSeq

			If .vspdData.Text = "" Then
			   tmpVal = 0
			Else
  			   tmpVal = UNICDbl(.vspdData.Text)
			End If

			If tmpval > MaxValue Then
			   MaxValue = UNICDbl(tmpVal)
			End If
		Next

		MaxValue = UNICDbl(MaxValue) + 1
		.vspddata.Row = Row
		.vspddata.Col = C_ItemSeq
		If UNICDbl(.vspdData.Text) <= UNICDbl(MaxValue) Then
			.vspdData.Text = MaxValue
		End If
	End With

End Function
'========================================================================================================
Function SetSumItem()

    Dim DblTotDrAmt
    Dim DblTotLocDrAmt
    Dim DblTotCrAmt
    Dim DblTotLocCrAmt
    Dim lngRows

	ggoSpread.Source = frm1.vspdData
	
    With frm1.vspdData

	If .MaxRows > 0 Then
	        For lngRows = 1 To .MaxRows
				.Row = lngRows
				.Col = 0
                If .text <> ggoSpread.DeleteFlag Then
					.col = C_DrCrFg
		            If .text = "DR" then
			            .Col = C_ItemAmt	'6
			            If .Text = "" Then
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + 0
			            Else
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + UNICDbl(.Text)
			            End If
			            .Col = C_ItemLocAmt	'7
			            If .Text = "" Then
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + 0
			            Else
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + UNICDbl(.Text)
			            End If
		            ElseIf .text = "CR" Then
			            .Col = C_ItemAmt	'6
			            If .Text = "" Then
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + 0
			            Else
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + UNICDbl(.Text)
			            End If
			            .Col = C_ItemLocAmt	'7
			            If .Text = "" Then
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + 0
			            Else
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + UNICDbl(.Text)
			            End If
					End If
				End If
	        Next
		End If

		frm1.txtDrAmt.Text		= UNIConvNumPCToCompanyByCurrency(DblTotDrAmt,		parent.gCurrency, parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
		frm1.txtCrAmt.Text		= UNIConvNumPCToCompanyByCurrency(DblTotCrAmt,		parent.gCurrency, parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
		frm1.txtDrLocAmt.Text	= UNIConvNumPCToCompanyByCurrency(DblTotLocDrAmt,	parent.gCurrency, parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
		frm1.txtCrLocAmt.Text	= UNIConvNumPCToCompanyByCurrency(DblTotLocCrAmt,	parent.gCurrency, parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")

	End With

End Function

'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Dim iFld1
	Dim iFld2
	Dim iTable
	Dim istrCode
	ggoSpread.Source = frm1.vspdData
	
	With frm1.vspdData
		If Row > 0 And Col = C_BizareaPopup Then
			.Col = Col - 1
			.Row = Row
			Call OpenPopUp(.Text, 3 )
    	End If

		If Row > 0 And Col = C_DeptPopup Then
			.Col = Col - 1
			.Row = Row
			Call OpenPopUp(.Text, 4 )
    	End If

    	If Row > 0 And Col = C_AcctPopUp Then
			.Col = Col - 1
			.Row = Row
			Call OpenPopUp(.Text, 5 )
    	End If

	End With

End Sub

'========================================================================================================
Sub  vspdData_Change (ByVal Col, ByVal Row )

	Dim tmpAcctCd

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    frm1.vspdData.Row = Row

    Select Case Col
		Case   C_AcctCd
			frm1.vspdData.Col = 0
			If  frm1.vspdData.Text = ggoSpread.InsertFlag Then
				frm1.vspdData.Col = C_ItemSeq
				frm1.hItemSeq.value = frm1.vspdData.Text
				frm1.vspdData.Col = C_AcctCd
				If Len(frm1.vspdData.Text) > 0 Then
					frm1.vspdData.Row = Row
					frm1.vspdData.Col = C_ItemSeq
					DeleteHsheet frm1.vspdData.Text
					Call Dbquery3(Row)
				Else
					frm1.vspdData.Col = C_AcctNm
					frm1.vspdData.Text = ""
				End If
	    	End If
		Case	C_deptcd
			frm1.vspdData.Col = C_DeptCd

			'Call DeptCd_underChange(frm1.vspdData.text)
			Call DeptCd_underChange2(Col,Row,frm1.vspdData.text)

		case C_DrCrNM

		call vspdData_ComboSelChange( Col, Row)
    End Select

End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Dim tmpDrCrFG



	frm1.vspdData.Row = Row
	frm1.vspdData.Col = 0
	If frm1.vspdData.MaxRows <= 0 Or UCase(Trim(frm1.vspdData.Text)) = UCase(Trim(ggoSpread.InsertFlag)) Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("0000111111")
	End If

    gMouseClickStatus = "SPC"	'Split 상태코드 

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

 	frm1.vspdData.Col = C_AcctCd
	If Len(frm1.vspdData.Text) < 1 Then
		ggoSpread.Source = frm1.vspdData2
		Call ggoSpread.ClearSpreadData()
	End If


End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")

End Sub

'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    If Row <> NewRow And NewRow > 0 Then
        With frm1
            .vspdData.Row = NewRow
            .vspdData.Col = C_ItemSeq
            .hItemSeq.value = .vspdData.Text
            ggoSpread.Source = .vspdData2
            Call ggoSpread.ClearSpreadData()
        End With

        frm1.vspddata.Col = 0
        If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
			Exit Sub
		End if

        lgCurrRow = NewRow
		Call DbQuery2(lgCurrRow)
    End If

End Sub


'========================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex
	Dim tmpDrCrFg
	 '---------- Coding part -------------------------------------------------------------
	 ' 이 Template 화면에서는 없는 로직임, 콤보(Name)가 변경되면 콤보(Code, Hidden)도 변경시켜주는 로직 
	With frm1
		.vspddata.Row = Row
		Select Case Col
			Case C_DrCrNm
       			.vspddata.Col = Col
				intIndex = .vspddata.Value
				.vspddata.Col = C_DrCrFg
				.vspddata.Value = intIndex
				Call SetSpread2Color()

		End Select
	End With

End Sub

'========================================================================================================
Sub txtTempGlNo_OnkeyPress()
	If window.event.keycode = 39 then	'Single quotation mark 입력불가 
		window.event.keycode = 0
	End If
End Sub

'========================================================================================================
Sub txtTempGlNo_OnKeyUp()
	If Instr(1,frm1.txtTempGlNo.value,"'") > 0 then
		frm1.txtTempGlNo.value = Replace(frm1.txtTempGlNo.value, "'", "")
	End if
End Sub

'========================================================================================================
Sub txtTempGlNo_onPaste()

	Dim iStrTempGlNo

	iStrTempGlNo = window.clipboardData.getData("Text")
	iStrTempGlNo = RePlace(iStrTempGlNo, "'", "")
	Call window.clipboardData.setData("text",iStrTempGlNo)

End Sub

'========================================================================================================
Sub txtGLDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtGLDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtGLDt.focus
    End If
End Sub

'==========================================================================================

Sub DeptCd_underChange(Byval strCode)
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
    Dim TempBizArea
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj
	
    If Trim(frm1.txtGLDt.Text) = "" Then    
		Exit sub
    End If
    If Trim(strCode) = "" Then    
		Exit sub
    End If    

    lgBlnFlgChgValue = True

With frm1.vspdData
	.row = .ActiveRow
		.focus
		.action=0
end with

	frm1.vspdData.Col = C_bizareacd			
	frm1.vspdData.Row = frm1.vspdData.ActiveRow	
	TempBizArea = Trim(frm1.vspdData.text)

	If TempBizArea = "" Then	'사업장이 입력된 경우 

		strSelect	=			 " b.biz_area_cd, c.biz_area_nm, a.dept_cd, a.dept_nm, a.org_change_id, a.internal_cd "    		
		strFrom		=			 " b_acct_dept a(NOLOCK), b_cost_center b(NOLOCK), b_biz_area c(nolock) "		
		strWhere	=			 " a.dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S") 
		strWhere	= strWhere & " and a.cost_cd = b.cost_cd and b.biz_area_cd = c.biz_area_cd "
		strWhere	= strWhere & " and a.org_change_id = (select distinct org_change_id "
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = True Then
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
			frm1.vspdData.Col = C_bizareacd		
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.text = Trim(arrVal2(1))
			frm1.vspdData.Col = C_bizareanm		
			frm1.vspdData.text = Trim(arrVal2(2))			
			frm1.vspdData.Col = C_deptcd			
			frm1.vspdData.text = Trim(arrVal2(3))
			frm1.vspdData.Col = C_deptnm		
			frm1.vspdData.text = Trim(arrVal2(4))
			Next	
		Else	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  

			frm1.vspdData.Col = C_deptcd			
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.text = ""
			frm1.vspdData.Col = C_deptnm		
			frm1.vspdData.text = ""
		End If 	
	Else	'사업장이 입력되지 않은 경우 

		strSelect	=			 " b.biz_area_cd, a.dept_cd,a. dept_nm, a.org_change_id, a.internal_cd "    		
		strFrom		=			 " b_acct_dept a(NOLOCK), b_cost_center b(NOLOCK) "		
		strWhere	=			 " a.dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S") 
		strWhere	= strWhere & " and a.cost_cd = b.cost_cd  "
		strWhere	= strWhere & " and b.biz_area_cd = " & FilterVar(LTrim(RTrim(TempBizArea)), "''", "S") 
		strWhere	= strWhere & " and a.org_change_id = (select distinct org_change_id "
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = True Then
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.vspdData.Col = C_deptcd			
				frm1.vspdData.text = Trim(arrVal2(2))
				frm1.vspdData.Col = C_deptnm		
				frm1.vspdData.text = Trim(arrVal2(3))
			Next
		Else
			IntRetCD = DisplayMsgBox("117218","X","X","X")  
			frm1.vspdData.Col = C_deptcd			
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.text = ""
			frm1.vspdData.Col = C_deptnm		
			frm1.vspdData.text = ""
		End If 	
	End If

End Sub


Sub DeptCd_underChange2(col,Row,strcode)
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
    Dim TempBizArea
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj


    If Trim(frm1.txtGLDt.Text) = "" Then    
		Exit sub
    End If
    If Trim(strCode) = "" Then    
		Exit sub
    End If    

    lgBlnFlgChgValue = True

With frm1.vspdData
	.row = Row
		.focus
		.action=0
end with

	frm1.vspdData.Col = C_bizareacd			
	frm1.vspdData.Row = frm1.vspdData.ActiveRow	
	TempBizArea = Trim(frm1.vspdData.text)

	If TempBizArea = "" Then	'사업장이 입력된 경우 

		strSelect	=			 " b.biz_area_cd, c.biz_area_nm, a.dept_cd, a.dept_nm, a.org_change_id, a.internal_cd "    		
		strFrom		=			 " b_acct_dept a(NOLOCK), b_cost_center b(NOLOCK), b_biz_area c(nolock) "		
		strWhere	=			 " a.dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S") 
		strWhere	= strWhere & " and a.cost_cd = b.cost_cd and b.biz_area_cd = c.biz_area_cd "
		strWhere	= strWhere & " and a.org_change_id = (select distinct org_change_id "
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = True Then
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
			frm1.vspdData.Col = C_bizareacd		
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.text = Trim(arrVal2(1))
			frm1.vspdData.Col = C_bizareanm		
			frm1.vspdData.text = Trim(arrVal2(2))			
			frm1.vspdData.Col = C_deptcd			
			frm1.vspdData.text = Trim(arrVal2(3))
			frm1.vspdData.Col = C_deptnm		
			frm1.vspdData.text = Trim(arrVal2(4))
			Next	
		Else	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  

			frm1.vspdData.Col = C_deptcd			
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.text = ""
			frm1.vspdData.Col = C_deptnm		
			frm1.vspdData.text = ""
		End If 	
	Else	'사업장이 입력되지 않은 경우 

		strSelect	=			 " b.biz_area_cd, a.dept_cd,a. dept_nm, a.org_change_id, a.internal_cd "    		
		strFrom		=			 " b_acct_dept a(NOLOCK), b_cost_center b(NOLOCK) "		
		strWhere	=			 " a.dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S") 
		strWhere	= strWhere & " and a.cost_cd = b.cost_cd  "
		strWhere	= strWhere & " and b.biz_area_cd = " & FilterVar(LTrim(RTrim(TempBizArea)), "''", "S") 
		strWhere	= strWhere & " and a.org_change_id = (select distinct org_change_id "
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = True Then
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.vspdData.Col = C_deptcd			
				frm1.vspdData.text = Trim(arrVal2(2))
				frm1.vspdData.Col = C_deptnm		
				frm1.vspdData.text = Trim(arrVal2(3))
			Next
		Else
			IntRetCD = DisplayMsgBox("117218","X","X","X")  
			frm1.vspdData.Col = C_deptcd			
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.text = ""
			frm1.vspdData.Col = C_deptnm		
			frm1.vspdData.text = ""
		End If 	
	End If

End Sub


'==========================================================================================

Sub QueryDeptCd_OnChange()
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj

	If Trim(frm1.txtGLDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.Value)), "''", "S") 
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
	
		
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
'			IntRetCD = DisplayMsgBox("124600","X","X","X")  
		frm1.txtDeptCd.Value = ""
		frm1.txtDeptNm.Value = ""
'			frm1.hOrgChangeId.Value = ""
	Else 
		
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
		jj = Ubound(arrVal1,1)
					
		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))			
			frm1.hOrgChangeId.Value = Trim(arrVal2(2))
		Next	
			
	End If
	

End Sub

'==========================================================================================
Sub txtGLDt_Change()

	If lgstartfnc = False Then
	    If lgFormLoad = True Then
			Dim strSelect
			Dim strFrom
			Dim strWhere 	
			Dim IntRetCD 
			Dim ii
			Dim arrVal1
			Dim arrVal2
			Dim jj


			lgBlnFlgChgValue = True
			With frm1
			
			If LTrim(RTrim(.txtDeptCd.Value)) <> "" and Trim(.txtGLDt.Text <> "") Then
				'----------------------------------------------------------------------------------------
				strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
				strFrom		=			 " b_acct_dept(NOLOCK) "		
				strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(.txtDeptCd.Value)), "''", "S") 
				strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
				strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
				strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
					If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
						IntRetCD = DisplayMsgBox("124600","X","X","X")
						.txtDeptCd.Value = ""
						.txtDeptNm.Value = ""
						'.hOrgChangeId.Value = ""
						If .vspdData.MaxRows <> 0 Then
							For ii = 1 To .vspdData.MaxRows
							.vspdData.Col = C_deptcd			
						    .vspdData.Row = ii
						    .vspdData.text = ""
						    .vspdData.Col = C_deptnm	
						    .vspdData.text = ""
							Next		
						End If
					Else
						arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
						jj = Ubound(arrVal1,1)
								
						For ii = 0 to jj - 1
							arrVal2 = Split(arrVal1(ii), chr(11))			
							frm1.hOrgChangeId.Value = Trim(arrVal2(2))
						Next	
					End If 
				End If
			End With
		End If
	End IF
End Sub


'==========================================================================================

Sub txtDeptCd_OnChange()
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj

	If Trim(frm1.txtGLDt.Text) = "" Then    
		Exit sub
    End If
    
	If Trim(frm1.txtDeptCd.value) = "" Then    
		Exit sub
    End If
    
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
	
		
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
		End If
		frm1.txtDeptCd.value = ""
		frm1.txtDeptNm.value = ""
		frm1.txtOrgChangeId.value = ""
	Else 
		
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
		jj = Ubound(arrVal1,1)
					
		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))			
			frm1.txtOrgChangeId.value = Trim(arrVal2(2))
		Next	
			
	End If
	
End Sub

'==========================================================================================
Sub txtDocCur_OnChange()
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()

	END IF	    

End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN=CENTER CLASS=CLSMTABP><FONT COLOR=WHITE>본지점전표등록(KO441)</FONT></TD>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH="100%"> </TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
					    <FIELDSET CLASS=CLSFLD>
						  <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>결의번호</TD>
								<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME=txttempglno ALT="결의번호" MAXLENGTH=18 SIZE=20 STYLE="TEXT-ALIGN: LEFT" tag ="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnTempGlNo ALIGN=TOP TYPE=BUTTON ONCLICK="vbscript:Call OpenRefTempGl()"></TD>
							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%" ></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP >
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>결의일자</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME=txtGLDt CLASS=FPDTYYYYMMDD TITLE=FPDATETIME tag="22" ALT="회계일자" ID=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME=txtDeptCd ALT="부서코드" MAXLENGTH=10 SIZE=10 STYLE="TEXT-ALIGN: LEFT" tag ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnCostCd ALIGN=TOP TYPE=BUTTON ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)">&nbsp;
													 <INPUT TYPE=TEXT NAME=txtDeptNm ALT="부서명"   MAXLENGTH=20 SIZE=20 STYLE="TEXT-ALIGN: LEFT" tag ="24X"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME=txtDocCur ALT="거래통화" MAXLENGTH=3 SIZE=10 STYLE="TEXT-ALIGN: LEFT" tag ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnCostCd ALIGN=TOP TYPE=BUTTON ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.Value, 2)"></TD>
								<TD CLASS=TD5 NOWRAP>이체번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME=txtHqBrchNo ALT="이체번호" MAXLENGTH=10 SIZE=13 STYLE="TEXT-ALIGN: CENTER" tag ="24X"></TD>
						   </TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT NAME="txtDesc" ALT="비고" MAXLENGTH="128" SIZE="70" tag="22N" ></TD>
							</TR>							   						   						   
							<TR>
								<TD HEIGHT="60%" COLSPAN=4>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH="100%" HEIGHT="100%" tag="23" TITLE=SPREAD> <PARAM NAME=MaxCols VALUE=0><PARAM NAME=MaxRows VALUE=0></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>차대합계(거래)</TD>
								<TD>
								    &nbsp;
								    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtDrAmt STYLE="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" TITLE=FPDOUBLESINGLE tag="24X2" ALT="차변합계(거래)"></OBJECT>');</SCRIPT>
								    &nbsp;
								    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtCrAmt STYLE="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" TITLE=FPDOUBLESINGLE tag="24X2" ALT="대변합계(거래)"></OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS=TD5 NOWRAP>차대합계(자국)</TD>
								<TD>
								    &nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtDrLocAmt STYLE="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" TITLE=FPDOUBLESINGLE ALT="차변합계(자국)" tag="24X2"></OBJECT>');</SCRIPT>
									&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> NAME=txtCrLocAmt STYLE="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" TITLE=FPDOUBLESINGLE ALT="대변합계(자국)" tag="24X2"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="40%" COLSPAN=4>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH="100%" HEIGHT="100%" tag="23" TITLE=SPREAD> <PARAM NAME=MaxCols VALUE=0><PARAM NAME=MaxRows VALUE=0></OBJECT>');</SCRIPT></TD>
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
		<TR HEIGHT=20>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTToN NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTToN>&nbsp;
						<BUTToN NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTToN>&nbsp;
					</TD>															
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
		<IFRAME NAME=MyBizASP WIDTH="100%" tag="2" HEIGHT=<%=BizSize%>  SRC="../../blank.htm" FRAMEBORDER=0 SCROLLING=NO NORESIZE FRAMESPACING=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS=HIDDEN NAME=txtSpread  tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA CLASS=HIDDEN NAME=txtSpread3 tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME=txtMode        tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=htxttempglno   tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtCommandMode tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hOrgChangeId"   tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME=txtUpdtUserId  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtInsrtUserId tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtMaxRows     tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtFlgMode     tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtGlinputType tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hCongFg        tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hItemSeq       tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=hAcctCd        tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtMaxRows3    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtOrgChangeId tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME=txtgCurrency   tag="24" TABINDEX="-1">
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 NAME=vspdData3 WIDTH="100%" TABINDEX="-1" id=OBJECT2><PARAM NAME=MaxCols VALUE=0><PARAM NAME=MaxRows VALUE=0></OBJECT>');</SCRIPT>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<ForM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</Form>
</BODY>
</HTML>
