<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : BA
'*  2. Function Name        : 조직개편진행현황 
'*  3. Program ID           : B2604ba1
'*  4. Program Name         : 조직개편진행현황 
'*  5. Program Desc         : 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2005/10/12
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'=======================================================================================================
'												1. 선 언 부 
'=======================================================================================================
=======================================================================================================
'                                               1.1 Inc 선언   
'	기능: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																			'☜: indicates that All variables must be declared in advance 

'=======================================================================================================
'                                               1.2 Global 변수/상수 선언  
'	.Constant는 반드시 대문자 표기.
'	.변수 표준에 따름. prefix로 g를 사용함.
'	.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=======================================================================================================
'@PGM_ID
Const BIZ_PGM_ID = "b2406bb1.asp"													'조회 비지니스 로직 ASP명 

'@Grid_Column
Dim C_ORG_CHANGE_ID
Dim C_ORGNM
Dim C_WORK_FLAG
Dim C_WORKFLAG_NM
Dim C_WORK_DT
Dim	C_OK_OPEN_FG
Dim	C_CANCEL_OPEN_FG
Dim C_WORK_OK
Dim C_WORK_CANCEL


'@Global_Var
Dim lgBlnFlgChgValue           'Variable is for Dirty flag
Dim lgIntGrpCount              'Group View Size를 조사할 변수 
Dim lgIntFlgMode               'Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows

Dim IsOpenPop
Dim lgSortKey          

'======================================================================================================
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'=======================================================================================================

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
 '   lgIntGrpCount = 0                           'initializes Group View Size
    
'    lgStrPrevKey = ""                           'initializes Previous Key
'    lgLngCurRows = 0                            'initializes Deleted Rows Count
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub SetDefaultVal()

End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call loadInfTB19029A("Q", "*", "NOCOOKIE", "MA")%>
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()

	C_ORG_CHANGE_ID    = 1
	C_ORGNM            = 2
	C_WORK_FLAG        = 3
	C_WORKFLAG_NM      = 4
	C_WORK_DT          = 5
	C_OK_OPEN_FG       = 6
	C_CANCEL_OPEN_FG   = 7
	C_WORK_OK          = 8
	C_WORK_CANCEL      = 9

End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	With frm1.vspdData
		.ReDraw = False
		
		.MaxCols = C_WORK_CANCEL + 1										'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols												'☜: 공통콘트롤 사용 Hidden Column
		.ColHidden = True

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021218", ,parent.gAllowDragDropSpread
		ggoSpread.ClearSpreadData

	   .ReDraw = True

		Call GetSpreadColumnPos()

		ggoSpread.SSSetEdit  C_ORG_CHANGE_ID,   "조직개편아이디"  ,15, 0
		ggoSpread.SSSetEdit  C_ORGNM,           "조직개편명"      ,20, 0
		ggoSpread.SSSetEdit  C_WORK_FLAG,       ""                ,15, 0		
		ggoSpread.SSSetEdit  C_WORKFLAG_NM,     "진행구분"        ,20, 0		
		ggoSpread.SSSetEdit  C_WORK_DT ,        "처리일자"        ,15, 2
		ggoSpread.SSSetEdit  C_OK_OPEN_FG ,     ""                , 2, 2
		ggoSpread.SSSetEdit  C_CANCEL_OPEN_FG , ""                , 2, 2		
		ggoSpread.SSSetCheck C_WORK_OK,         "확정"            , 8, ,"",True
		ggoSpread.SSSetCheck C_WORK_CANCEL,     "취소"            , 8, ,"",True				

		Call ggoSpread.SSSetColHidden(C_WORK_FLAG,C_WORK_FLAG,True)
		Call ggoSpread.SSSetColHidden(C_OK_OPEN_FG,C_OK_OPEN_FG,True)
		Call ggoSpread.SSSetColHidden(C_CANCEL_OPEN_FG,C_CANCEL_OPEN_FG,True)				

		.ReDraw = True

		Call SetSpreadLock 
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_ORG_CHANGE_ID, -1, C_WORK_CANCEL ,-1
		.vspdData.ReDraw = True
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor()
	Dim ii
	
	With frm1
		.vspdData.ReDraw = False

		For ii = 1 To .vspdData.MaxRows
			.vspdData.Col = C_OK_OPEN_FG
			.vspdData.Row = ii

			If Ucase(Trim(.vspdData.text)) = "Y" Then
				ggoSpread.SpreadUnLock C_WORK_OK,ii,C_WORK_OK,ii
			Else
				ggoSpread.SpreadLock   C_WORK_OK,ii,C_WORK_OK,ii	
			End If					

			.vspdData.Col = C_CANCEL_OPEN_FG
			If Ucase(Trim(.vspdData.text)) = "Y" Then
				ggoSpread.SpreadUnLock C_WORK_CANCEL,ii,C_WORK_CANCEL,ii
			Else
				ggoSpread.SpreadLock C_WORK_CANCEL,ii,C_WORK_CANCEL,ii
			End If													
		Next			
		.vspdData.ReDraw = True
	End With
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'============================================================================================================
Sub InitComboBox()

End Sub

 '******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
 '========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 

'======================================================================================================
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'=======================================================================================================

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos()
    Dim iCurColumnPos
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	C_ORG_CHANGE_ID    = iCurColumnPos(1)
	C_ORGNM            = iCurColumnPos(2)	
	C_WORK_FLAG        = iCurColumnPos(3)
	C_WORKFLAG_NM      = iCurColumnPos(4)
	C_WORK_DT          = iCurColumnPos(5)
	C_OK_OPEN_FG	   = iCurColumnPos(6)
	C_CANCEL_OPEN_FG   = iCurColumnPos(7)
	C_WORK_OK          = iCurColumnPos(8)
	C_WORK_CANCEL      = iCurColumnPos(9)

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
'    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
End Sub

'======================================================================================================
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'=======================================================================================================

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field

    Call InitSpreadSheet													'⊙: Setup the Spread sheet
	Call InitVariables
'    Call InitComboBox

    Call SetDefaultVal
    Call SetToolbar("1100000000001111")										'⊙: 버튼 툴바 제어 

    Call FncQuery
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")
	
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    With frm1
		If col = C_WORK_OK  Then
			.vspddata.col = C_WORK_OK
			.vspddata.row = row

			If .vspddata.text="1" Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.EditUndo
			End If	
		End If

		If col = C_WORK_CANCEL  Then
			.vspddata.col = C_WORK_CANCEL
			.vspddata.row = row
			
			If .vspddata.text="1" Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.EditUndo
			End If			
		End	If
    End With		
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
End Sub

'======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'=======================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData 
		If Row >= NewRow Then
			Exit Sub
		End If
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub  

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	IF CheckRunningBizProcess = True Then
		Exit Sub
	End If

	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then		'☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then								'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			DbQuery
		End If
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : 
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)
    If frm1.vspdData.MaxRows = 0 Then								'no data일 경우 vspdData_LeaveCell no 실행 
		Exit Sub													'tab이동시에 잘못된 140318 message 방지 
    End If
    
    'With frm1.vspdData
	'	 If NewCol > 0 Then 
	'		If Col = C_TAGET_WORKING_MNTH Then
	'			.Row = Row
	'			.Col = Col
	'		
	'			If .Text <> "" Then
     '               If CheckDateFormat(.Text, parent.gDateFormatYYYYMM) = False  Then
	'					Call DisplayMsgBox("140318","X","X","X")	'년월을 올바로 입력하세요.
	'					.Text = ""
	'				End If
	'			End If
	'		End If
	'	
	'	End If
    'End With
End Sub

'======================================================================================================
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'=======================================================================================================

'======================================================================================================
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'=======================================================================================================
'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery()
    Dim IntRetCD 
	
	FncQuery = False                                                        '⊙: Processing is NG
	
	Err.Clear																'☜: Protect system from crashing
	
	ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013",Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	'-----------------------
	'Erase contents area
	'-----------------------
	ggoSpread.ClearSpreadData()
	
	Call InitVariables
	
	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False Then
		Exit Function
	End If																	'☜: Query db data
	
	FncQuery = True															'⊙: Processing is OK
End Function

'======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'=======================================================================================================
Function FncNew() 
	On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncDelete() 
	On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : 
'=======================================================================================================
Function FncSave() 
    Dim IntRetCD 
	Dim var1
	Dim ChkCnt

    FncSave = False                                                         
    
    Err.Clear                                                               
    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False And var1 = False Then									'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")									'⊙: Display Message(There is no changed data.)
		Exit Function
    End If

    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
		Exit Function
    End If

	ChkCnt = VerifySelCount

	If ChkCnt > 1 Then
		Call DisplayMsgBox("237000","X","X","X")
		Exit Function
	ElseIf 	ChkCnt < 1 Then
		Call DisplayMsgBox("236021","X","X","X")
	End if

	'-----------------------
	'Save function call area
	'----------------------- 	
	If DbSave = False Then
		Exit Function																	'☜: Save db data
    End If
        
	FncSave = True                                      								'⊙: Processing is OK
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy() 
	On Error Resume Next                                           	       '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel() 
	On Error Resume Next                                           	       '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow() 
	On Error Resume Next                                           	       '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
	On Error Resume Next                                           	       '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)						'☜: 화면 유형 
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
	Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
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

'======================================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'=======================================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016",Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	FncExit = True
End Function

'======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbQuery() 
	Dim strVal

	Err.Clear                                                               			'☜: Protect system from crashing
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
		
	DbQuery = False

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		
	Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동 
		
	DbQuery = True																		'⊙: Processing is NG
End Function

'======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=======================================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE										'⊙: Indicates that current mode is Update mode
	lgBlnFlgChgValue = False
	
	Call SetSpreadColor()
    Call SetToolbar("1100100000001111")										'⊙: 버튼 툴바 제어 	
End Function

'======================================================================================================
' Function Name : VerifySelCount
' Function Desc : 선택한 작업의 갯수 체크 (한개 이상이면 에러처리)
'=======================================================================================================
Function VerifySelCount()
	Dim ii 
	Dim ChkCnt

	ChkCnt = 0

	With frm1
		For ii = 1 To .vspddata.MaxRows
			.vspddata.row = ii 
			.vspddata.col = C_OK_OPEN_FG
			If .vspddata.text = "Y" Then
				.vspddata.col = C_WORK_OK

				If .vspddata.text = "1" Then
					ChkCnt = ChkCnt + 1
				End If
			End If

			.vspddata.col = C_CANCEL_OPEN_FG			
			If .vspddata.text = "Y" Then
				.vspddata.col = C_WORK_CANCEL
				If .vspddata.text = "1" Then
					ChkCnt = ChkCnt + 1
				End If
			End If	
		Next
	End With

	VerifySelCount = ChkCnt
End Function


'======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbSave() 
	Dim lRow
	Dim strVal,tmpVal
	Dim strYear,strMonth,strDay	
	Dim ChkCnt

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	DbSave = False																		'⊙: Processing is NG

	On Error Resume Next																		'☜: Protect system from crashing
	Err.Clear

	With frm1
		.txtMode.value  = parent.UID_M0002

		'-----------------------
		'Data manipulate area
		'-----------------------
   		strVal = ""
		    		
		For lRow = 1 To .vspdData.MaxRows
	    	.vspdData.Row = lRow
			.vspddata.Col = 0
			
			If Trim(.vspddata.Text) = ggoSpread.UpdateFlag Then
				strVal = strVal & BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0002
				.vspdData.Col = C_ORG_CHANGE_ID
				strVal = strVal & "&txtOrgChangeID=" & Trim(.vspdData.Text) 
				.vspdData.Col = C_WORK_FLAG
				strVal = strVal & "&txtWorkType=" & Trim(.vspdData.Text) 
				.vspdData.Col = C_OK_OPEN_FG
				If .vspdData.Text="Y" Then
					.vspdData.Col = C_WORK_OK
					If .vspdData.Text="1" Then
						tmpVal = "Y"
					End If
				End If	
				.vspdData.Col = C_CANCEL_OPEN_FG
				If .vspdData.Text="Y" Then
					.vspdData.Col = C_WORK_CANCEL
					If .vspdData.Text="1" Then
						tmpVal = "N"
					End If
				End If
				strVal = strVal & "&txtYNfg=" & tmpVal
				Exit For
			End If	
		Next

		Call RunMyBizASP(MyBizASP, strVal)														'☜: 비지니스 ASP 를 가동 
	End With

	DbSave = True																			
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'=======================================================================================================
Function DbSaveOk()				            '☆: 저장 성공후 실행 로직 
   	Call InitVariables

	ggoSpread.Source = frm1.vspddata
	ggoSpread.ClearSpreadData()

   	Call MainQuery()
End Function

'======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'=======================================================================================================
Function DbDelete()

End Function



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<!--'======================================================================================================
'       					6. Tag부 
'	기능: Tag부분 설정 
	
'======================================================================================================= -->

<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD>&nbsp;</TD>					
					<TD>&nbsp;</TD>					
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% VALIGN=top COLSPAN=4>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

