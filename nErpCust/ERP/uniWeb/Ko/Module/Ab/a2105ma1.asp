<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID		    : A2105MA1
'*  4. Program Name         : 거래유형 등록 
'*  5. Program Desc         : 거래유형 등록 수정 삭제 조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/09/08
'*  8. Modified date(Last)  : 2001/03/05
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : Chang Joo, Kim
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================= -->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'⊙: 비지니스 로직 ASP명 
Const BIZ_PGM_ID = "a2105mb1.asp"			'☆: 비지니스 로직 ASP명 
'==========================================================================================================
'⊙: Grid Columns
Dim C_TransType
Dim C_TransNM
Dim C_TransEngNM
Dim C_BatchFg
Dim C_BatchFgNm	
Dim C_GlFg
Dim C_GlFgNm
Dim C_MoCD
Dim C_MoNM
Dim C_SysFg
Dim C_ReverseFgCd
Dim C_ReverseFgNm
Dim C_AcctSumFgCd
Dim C_AcctArrayalFgCd
Dim C_Inv_Post_Fg


'========================================================================================================
Sub InitSpreadPosVariables()
	 C_TransType		= 1
	 C_TransNM			= 2
	 C_TransEngNM		= 3
	 C_BatchFg			= 4
	 C_BatchFgNm		= 5
	 C_GlFg				= 6
	 C_GlFgNm			= 7
	 C_MoCD				= 8
	 C_MoNM				= 9
	 C_SysFg			= 10
	 C_ReverseFgCd		= 11
	 C_ReverseFgNm		= 12
	 C_AcctSumFgCd		= 13
	 C_AcctArrayalFgCd	= 14
	 C_Inv_Post_Fg      = 15
End Sub

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'========================================================================================================
Dim  IsOpenPop

'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count

    lgSortKey = 1
    lgPageNo   = 0
End Sub

'========================================================================================================= 
Sub SetDefaultVal()

End Sub

'======================================================================================== 
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
End Sub

'========================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021203",,parent.gAllowDragDropSpread

	With frm1.vspdData

		.MaxCols = C_Inv_Post_Fg + 1
		.MaxRows = 0
		.ReDraw = False

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit		C_TransType,		"거래유형코드",			20,	,	,	20,	2
		ggoSpread.SSSetEdit		C_TransNM,			"거래유형명",			25,	,	,	50
		ggoSpread.SSSetEdit		C_TransEngNM,		"거래유형영문명",		30,	,	,	50
		ggoSpread.SSSetCombo	C_BatchFg,			"ON-LINE/BATCH 구분",   1
		ggoSpread.SSSetCombo	C_BatchFgNm,		"ON-LINE/BATCH 구분",   18	
		ggoSpread.SSSetCombo	C_GlFg,				"전표구분",				1
		ggoSpread.SSSetCombo	C_GlFgNm,			"전표구분",				12
		ggoSpread.SSSetCombo	C_MoCD,				"업무구분",				2
		ggoSpread.SSSetCombo	C_MoNM,				"업무구분명",			20
		ggoSpread.SSSetCheck	C_SysFg,			"시스템구분",			12,	-10,"",	True, -1
		ggoSpread.SSSetCombo	C_ReverseFgCd,		"분개구분",				12
		ggoSpread.SSSetCombo	C_ReverseFgNm,		"분개구분",				12
		ggoSpread.SSSetCombo	C_AcctSumFgCd,		"계정합",				12
		ggoSpread.SSSetCombo	C_AcctArrayalFgCd,	"전표생성순서",			12
		ggoSpread.SSSetEdit		C_Inv_Post_Fg,		"",		2

		Call ggoSpread.MakePairsColumn(C_BatchFg,C_BatchFgNm,"1")
		Call ggoSpread.MakePairsColumn(C_GlFg,C_GlFgNm,"1")
		Call ggoSpread.MakePairsColumn(C_MoCD,C_MoNM,"1")
		Call ggoSpread.MakePairsColumn(C_ReverseFgCd,C_ReverseFgNm,"1")

		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_BatchFg,C_BatchFg,True)
		'Call ggoSpread.SSSetColHidden(C_GlFg,C_GlFg,True)
		Call ggoSpread.SSSetColHidden(C_MoCD,C_MoCD,True)
		Call ggoSpread.SSSetColHidden(C_ReverseFgCd,C_ReverseFgCd,True)
		Call ggoSpread.SSSetColHidden(C_ReverseFgNm,C_ReverseFgNm,True)
		Call ggoSpread.SSSetColHidden(C_Inv_Post_Fg,C_Inv_Post_Fg,True)						

		.ReDraw = True

		Call SetSpreadLock                                              '바뀐부분 
		Call InitComboBox
    End With
End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1.vspdData
		.ReDraw = False

		'SpreadLock(ByVal Col1, ByVal Row1, Optional ByVal Col2 = -10, Optional ByVal Row2 = -10)
		ggoSpread.SpreadLock		C_TransType,		-1, C_TransType		' Trans Type 코드를 Lock
		ggoSpread.SSSetRequired		C_TransNM,			-1, C_TransNM
		ggoSpread.SpreadLock		C_MoNM,				-1, C_MoNM				' Trans Type 코드를 Lock
		ggoSpread.SpreadLock		C_SysFg,			-1, C_SysFg				' 시스템구분을 Lock
		ggoSpread.SSSetRequired		C_AcctSumFgCd,		-1
		ggoSpread.SSSetRequired		C_AcctArrayalFgCd,	-1
		ggoSpread.SSSetProtected	.MaxCols,			-1,-1
		.ReDraw = True
    End With
End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False

		' 필수 입력 항목으로 설정 
		' SSSetRequired(ByVal Col, ByVal Row, Optional ByVal Row2 = -10)
		ggoSpread.SSSetRequired C_TransType      ,pvStartRow, pvEndRow		' 거래유형 
		ggoSpread.SSSetRequired C_TransNM        ,pvStartRow, -1			' 거래유형명 
		ggoSpread.SSSetRequired C_BatchFgNM      ,pvStartRow, pvEndRow		' 기표구분명 
		ggoSpread.SSSetRequired C_GlFgNM         ,pvStartRow, pvEndRow		' 전표구분명 
		ggoSpread.SSSetRequired C_MoNM           ,pvStartRow, pvEndRow		' 업무구분명 
		ggoSpread.SpreadLock	C_SysFg          ,pvStartRow, pvEndRow		' 시스템구분을 Lock
		ggoSpread.SpreadUnLock	C_AcctSumFgCd    ,pvStartRow, pvEndRow		' 계정합을 UnLock		
		ggoSpread.SSSetRequired C_AcctSumFgCd    ,pvStartRow, -1			' 계정합 
		ggoSpread.SSSetRequired C_AcctArrayalFgCd,pvStartRow, -1			' 전표생성순서 
		.vspdData.ReDraw = True
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor_AfterQry()
	Dim ii

    With frm1
		.vspddata.ReDraw = False
		For ii = 1 To .vspddata.Maxrows
			.vspddata.Row = ii
			.vspddata.col = C_Inv_Post_Fg
			If Trim(.vspddata.Text) <> "M" Then
				ggoSpread.SSSetProtected C_BatchFgNM ,ii,ii								
			End If
		Next
        .vspdData.ReDraw = True		
    End With
End Sub

'========================================================================================================= 
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx

	'On-Line/Batch구분 
	Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD=" & FilterVar("A1026", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_BatchFg
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_BatchFgNm

	'전표구분 
	Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD=" & FilterVar("A1027", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_GlFg
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_GlFgNm

	'업무구분명 
	Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD=" & FilterVar("B0001", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_MoCD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_MoNM

	'분개구분 
	Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD=" & FilterVar("A1021", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    iNameArr = lgF1

    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_ReverseFgCd
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_ReverseFgNM

	Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD=" & FilterVar("A1020", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	iCodeArr = lgF0
    
	'계정합 
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_AcctSumFgCd
    '전표생성순서 
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_AcctArrayalFgCd
End Sub


'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_TransType		  = iCurColumnPos(1)
			C_TransNM		  = iCurColumnPos(2)
			C_TransEngNM	  = iCurColumnPos(3)
			C_BatchFg		  = iCurColumnPos(4)
			C_BatchFgNm		  = iCurColumnPos(5)
			C_GlFg			  = iCurColumnPos(6)
			C_GlFgNm		  = iCurColumnPos(7)
			C_MoCD			  = iCurColumnPos(8)
			C_MoNM			  = iCurColumnPos(9)
			C_SysFg			  = iCurColumnPos(10)
			C_ReverseFgCd	  = iCurColumnPos(11)
			C_ReverseFgNm	  = iCurColumnPos(12)
			C_AcctSumFgCd	  = iCurColumnPos(13)
			C_AcctArrayalFgCd = iCurColumnPos(14)
			C_Inv_Post_Fg     = iCurColumnPos(15)  
	End Select    
End Sub

'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래유형 팝업"				' 팝업 명칭 
	arrParam(1) = "A_ACCT_TRANS_TYPE" 			' TABLE 명칭 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "거래유형"					' 조건필드의 라벨 명칭 

    arrField(0) = "TRANS_TYPE"					' Field명(0)
    arrField(1) = "TRANS_NM"					' Field명(1)
    arrField(2) = "TRANS_ENG_NM"				' Field명(2)

    arrHeader(0) = "거래유형코드"					' Header명(0)
    arrHeader(1) = "거래유형명"					' Header명(1)
    arrHeader(2) = "거래유형영문명"				' Header명(2)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtTransType.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If
End Function

'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then		' Condition
			.txtTransType.focus
			.txtTransType.value = Trim(arrRet(0))
			.txtTransNM.value = arrRet(1)
		End If
	End With
End Function

'========================================================================================================= 
Sub Form_Load()
    On Error Resume Next
    Err.Clear

	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
    Call InitVariables

    'Call SetDefaultVal    
    Call SetToolbar("1100110100101111")
	frm1.txtTransType.focus
End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
		.Row = Row

        Select Case Col
            Case C_BatchFgNm
                .Col = Col
                intIndex = .Value        '  COMBO의 VALUE값 
				.Col = C_BatchFg      '  CODE값란으로 이동 
				.Value = intIndex        '  CODE란의 값은 COMBO의 VALUE값이된다.
		    Case C_GlFgNm
                .Col = Col
                intIndex = .Value
				.Col = C_GlFg
				.Value = intIndex
		    Case C_MoNM
                .Col = Col
                intIndex = .Value
				.Col = C_MoCD
				.Value = intIndex
			Case C_ReverseFgNm
                .Col = Col
                intIndex = .Value
				.Col = C_ReverseFgCd
				.Value = intIndex
		End Select
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow

			.Col = C_BatchFg		:	intIndex = .value
			.col = C_BatchFgNM		:	.value = intindex

			.Col = C_GlFg			:	intIndex = .value
			.col = C_GlFgNM			:	.value = intindex
            
			.Col = C_ReverseFgCd	:	intIndex = .value
			.col = C_ReverseFgNm	:	.value = intindex

			.Col = C_MoCD			:	intIndex = .value
			.col = C_MoNM			:	.value = intindex
			
		Next
	End With
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    If frm1.vspdData.MaxRows <= 0 Then
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
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
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
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

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
	
	
    Call CheckMinNumSpread(frm1.vspdData, Col, Row) 
	
	
  	Call typecheck()
			
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
End Sub

Sub typecheck()
Dim intRow
With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow

		Frm1.vspdData.Col = C_TransType
		Frm1.vspdData.Row = intRow
		if Frm1.vspdData.value = "AU002" then
			Frm1.vspdData.Col = C_GlFg
			Frm1.vspdData.text = "G"
			Frm1.vspdData.Col = C_GlFgNM
			Frm1.vspdData.text = "회계전표"
			ggoSpread.SSSetProtected	C_GlFgNM,			intRow,intRow
		End if
		
		if Frm1.vspdData.value = "AU001" then
			Frm1.vspdData.Col = C_GlFg
			Frm1.vspdData.text = "T"
			Frm1.vspdData.Col = C_GlFgNM
			Frm1.vspdData.text = "결의전표"
			ggoSpread.SSSetProtected	C_GlFgNM,			intRow,intRow
		End If	
			
		if Frm1.vspdData.value = "FN005" then
			Frm1.vspdData.Col = C_GlFg
			Frm1.vspdData.text = "G"
			Frm1.vspdData.Col = C_GlFgNM
			Frm1.vspdData.text = "회계전표"
			ggoSpread.SSSetProtected	C_GlFgNM,			intRow,intRow
		End if	
		
		Next
	End With

End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell()
'   Event Desc : 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub

'========================================================================================================= 
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End If
End Sub

'========================================================================================
Function FncQuery() 
	Dim IntRetCD 

    FncQuery = False
    Err.Clear

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If

    Call ggoOper.ClearField(Document, "2") 
    Call InitVariables

    If Not chkField(Document, "1") Then	
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery

    FncQuery = True
End Function

'========================================================================================
Function FncNew()
	On Error Resume Next
End Function

'========================================================================================
Function FncDelete()
	On Error Resume Next
End Function

'========================================================================================
Function FncSave()
	Dim IntRetCD 

    FncSave = False 
    Err.Clear
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                          'No data changed!!
        Exit Function
    End If

    '-----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
		Exit Function
    End If

    Call DbSave
    FncSave = True
End Function

'========================================================================================
Function FncCopy() 
	Dim IntRetCD

	frm1.vspdData.ReDraw = False

    If frm1.vspdData.MaxRows < 1 then Exit Function

    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow

    frm1.vspdData.Col  = C_TransType
    frm1.vspdData.Text = ""
    frm1.vspdData.Col  = C_SysFg
    frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True
End Function

'========================================================================================================= 
Function FncCancel() 
    If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
    Call InitData
End Function

'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow

    On Error Resume Next
    Err.Clear

    FncInsertRow = False

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
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
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True
    End If

    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
Function FncDeleteRow() 
	Dim lDelRows
	Dim iDelRowCnt, i

    If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData
		.focus
		ggoSpread.Source = frm1.vspdData
		lDelRows = ggoSpread.DeleteRow
    End With
End Function

'========================================================================================
Function FncPrev()
    On Error Resume Next
End Function

'========================================================================================
Function FncNext()
    On Error Resume Next
End Function

'========================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================================================================
Function FncExcel()
    Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
Function FncFind()
    Call parent.FncFind(Parent.C_MULTI, False)
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")   '☜ 바뀐부분 
		If IntRetCD = vbNo Then
		     Exit Function
		End If
	End If

    FncExit = True
End Function

'========================================================================================
Function DbQuery()
	Dim strVal

    DbQuery = False
    Err.Clear

	Call LayerShowHide(1)

    With frm1
	    If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtTransType=" & Trim(.hTransType.value)	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtTransType=" & Trim(.txtTransType.value)	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		End If

		strVal = strVal & "&lgPageNo="       & lgPageNo
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

		Call RunMyBizASP(MyBizASP, strVal)
    End With

    DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")
	Call SetSpreadLock
    Call InitData
	Call SetToolbar("110011110011111")
	Call SetSpreadColor_AfterQry
	call typecheck()
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
Function DbSave() 
	Dim lRow      
	Dim strVal, strDel
	
    DbSave = False
    'On Error Resume Next

	Call LayerShowHide(1)

	With frm1
		.txtMode.value = Parent.UID_M0002

		'-----------------------
		'Data manipulate area
		'-----------------------
		strVal = ""
		strDel = ""

		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
		    .vspdData.Row = lRow
		    .vspdData.Col = 0

		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag							'☜: 신규 
														strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep
		            .vspdData.Col = C_TransType		:	strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_TransNM		:	strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_TransEngNM	:   strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_BatchFg		:   strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_GlFg			:   strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MoCD			:   strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MoNM			:   strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_SysFg
		            If .vspdData.Text <> "" Then
						If .vspdData.Text = 0 Then
							strVal = strVal & "N" & Parent.gColSep
						Else
							strVal = strVal & "Y" & Parent.gColSep
						End If
					Else
						strVal = strVal & "N" & Parent.gColSep
					End If
					
		            .vspdData.Col = C_ReverseFgCd	:	strVal = strVal & "M" & Parent.gColSep
		            .vspdData.Col = C_AcctSumFgCd	:	strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_AcctArrayalFgCd	:	strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
				Case ggoSpread.UpdateFlag							'☜: 수정 
														strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep
				    .vspdData.Col = C_TransType		:	strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				    .vspdData.Col = C_TransNM		:	strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_TransEngNM	:	strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_BatchFg		:   strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_GlFg			:   strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				    .vspdData.Col = C_MoCD			:	strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				    .vspdData.Col = C_MoNM			:   strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_SysFg			
		            If .vspdData.Text <> "" Then
						If .vspdData.Text = 0 Then
							strVal = strVal & "N" & Parent.gColSep
						Else
							strVal = strVal & "Y" & Parent.gColSep
						End If
					Else
						strVal = strVal & "N" & Parent.gColSep
					End If
					
		            .vspdData.Col = C_ReverseFgCd	:	strVal = strVal & "M" & Parent.gColSep
		            .vspdData.Col = C_AcctSumFgCd	:	strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_AcctArrayalFgCd	:	strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		        Case ggoSpread.DeleteFlag							'☜: 삭제 
														strVal = strVal & "D" & Parent.gColSep & lRow & Parent.gColSep
		            .vspdData.Col = C_TransType		:   strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
		    End Select
		Next

		.txtMode.value			= Parent.UID_M0002
'		.txtUpdtUserId.value	= Parent.gUsrID
'		.txtInsrtUserId.value	= Parent.gUsrID
'		.txtMaxRows.value		= lGrpCnt - 1
		.txtSpread.value		= strDel & strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

    DbSave = True
End Function

'========================================================================================
Function DbSaveOk()
	Call ggoOper.ClearField(Document, "2")
	Call InitVariables
	Call DbQuery
End Function


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>거래유형등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>거래유형</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtTransType" MAXLENGTH="20" SIZE=20 ALT ="거래유형" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtTransType.Value, 0)">&nbsp;
													<INPUT NAME="txtTransNM" MAXLENGTH="50" SIZE=30 ALT  ="거래유형명" tag="14X"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
								<script language =javascript src='./js/a2105ma1_I421677112_vspdData.js'></script>
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
			<IFRAME NAME="MyBizASP" src="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24" tabindex="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="hTransType" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT" >
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
