<%@ LANGUAGE="VBSCRIPT" %>
<!--
'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : ������������ 
'*  3. Program ID           : c3980ma1
'*  4. Program Name         : ������Ȳ 
'*  5. Program Desc         : 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2004/12/16
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'=======================================================================================================
'												1. �� �� �� 
'=======================================================================================================
=======================================================================================================
'                                               1.1 Inc ����   
'	���: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																			'��: indicates that All variables must be declared in advance 

'=======================================================================================================
'                                               1.2 Global ����/��� ����  
'	.Constant�� �ݵ�� �빮�� ǥ��.
'	.���� ǥ�ؿ� ����. prefix�� g�� �����.
'	.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=======================================================================================================
'@PGM_ID
Const BIZ_PGM_ID = "C3980mb1.asp"													'��ȸ �����Ͻ� ���� ASP�� 

'@Grid_Column
Dim C_CLOSE_GB
Dim C_CLOSE_GB_NM
Dim C_TARGET
Dim C_CLOSE_MNTH
Dim C_WORKING_DT
Dim C_WORKING_ID
Dim C_TAGET_WORKING_MNTH
Dim C_CLOSE
Dim C_CANCEL


'@Global_Var
Dim lgBlnFlgChgValue           'Variable is for Dirty flag
Dim lgIntGrpCount              'Group View Size�� ������ ���� 
Dim lgIntFlgMode               'Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows

Dim IsOpenPop
Dim lgSortKey          

'======================================================================================================
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'=======================================================================================================

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
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
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
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
	C_CLOSE_GB           = 1
	C_CLOSE_GB_NM        = 2
	C_TARGET             = 3
	C_CLOSE_MNTH         = 4
	C_WORKING_DT         = 5
	C_WORKING_ID         = 6
	C_TAGET_WORKING_MNTH = 7
	C_CLOSE              = 8
	C_CANCEL             = 9
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	Dim strMaskYM
	
	strMaskYM = parent.gDateFormatYYYYMM

	strMaskYM = Replace(strMaskYM,"YYYY"      ,"9999")
	strMaskYM = Replace(strMaskYM,"YY"        ,"99")
	strMaskYM = Replace(strMaskYM,"MM"        ,"99")
	strMaskYM = Replace(strMaskYM,parent.gComDateType,"X")

	Call InitSpreadPosVariables()

	With frm1.vspdData
		.ReDraw = False
		
		.MaxCols = C_CANCEL + 1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols												'��: ������Ʈ�� ��� Hidden Column
		.ColHidden = True

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021218", ,parent.gAllowDragDropSpread
		ggoSpread.ClearSpreadData

	   .ReDraw = True

		Call GetSpreadColumnPos()

		ggoSpread.SSSetEdit  C_CLOSE_GB   ,        ""  , 2, 0
		ggoSpread.SSSetEdit  C_CLOSE_GB_NM,        "�۾�����"  ,15, 0		
		ggoSpread.SSSetEdit  C_TARGET	 ,         "���"      ,15, 0		
		ggoSpread.SSSetEdit  C_CLOSE_MNTH ,        "������"    ,12, 2		
		ggoSpread.SSSetEdit  C_WORKING_DT ,        "�۾���"    ,15, 2
		ggoSpread.SSSetEdit  C_WORKING_ID ,		   "������"    ,15, 2	
        ggoSpread.SSSetMask  C_TAGET_WORKING_MNTH, "�۾�����",10, 2, strMaskYM            		
		ggoSpread.SSSetCheck C_CLOSE,              "����"      , 8, ,"",True
		ggoSpread.SSSetCheck C_CANCEL,             "���"      , 8, ,"",True

		Call ggoSpread.SSSetColHidden(C_CLOSE_GB,C_CLOSE_GB,True)
		Call ggoSpread.SSSetColHidden(C_TAGET_WORKING_MNTH,C_TAGET_WORKING_MNTH,True)		

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
		ggoSpread.SpreadLock C_CLOSE_GB, -1, C_CANCEL ,-1
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

		For ii = 1 To .vspdData.Maxrows
			.vspdData.row = ii
			.vspdData.Col = C_CLOSE_GB

			If  Ucase(Trim(.vspdData.text)) = "CC" Then
				.vspdData.Col = C_CLOSE

				If Ucase(Trim(.vspdData.text)) = "1" Then
					ggoSpread.SpreadUnLock C_CLOSE,ii,C_CLOSE,ii
					.vspdData.text = "0"
				Else 
					.vspdData.Col = C_CANCEL
					If Ucase(Trim(.vspdData.text)) = "1" Then
						.vspdData.Col = C_CLOSE_MNTH
						If Ucase(Trim(.vspdData.text)) <> "" Then					
							ggoSpread.SpreadUnLock C_CANCEL,ii,C_CANCEL,ii
							.vspdData.text = "0"	
							.vspdData.Col = C_CLOSE
							.vspdData.text = "0"
						Else	
							.vspdData.Col = C_CANCEL
							ggoSpread.SpreadLock C_CANCEL,ii,C_CANCEL,ii
							.vspdData.text = "0"								
						End If	
					End If
				End If	
			ElseIf Ucase(Trim(.vspdData.text)) = "AC" Or Ucase(Trim(.vspdData.text)) = "MC" Then	
				.vspdData.Col = C_CLOSE
				If Ucase(Trim(.vspdData.text)) = "1" Then
					ggoSpread.SpreadUnLock C_CLOSE,ii,C_CLOSE,ii
					.vspdData.text = "0"
				End If					
				.vspdData.Col = C_CANCEL
				If Ucase(Trim(.vspdData.text)) = "1" Then
					ggoSpread.SpreadUnLock C_CANCEL,ii,C_CANCEL,ii
					.vspdData.text = "0"
				End If													
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

 '******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 
 '========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 

'======================================================================================================
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'=======================================================================================================

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow

    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
		iRow = CInt(iPosArr(0))
		For iDx = 1 To  frm1.vspdData.MaxCols - 1
			frm1.vspdData.Col = iDx
			frm1.vspdData.Row = iRow
			If frm1.vspdData.ColHidden <> True And frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
				frm1.vspdData.Col = iDx
				frm1.vspdData.Row = iRow
				frm1.vspdData.Action = 0 ' go to
				Exit For
			End If
		Next
    End If
End Sub


'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos()
    Dim iCurColumnPos
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	C_CLOSE_GB           = iCurColumnPos(1)
	C_CLOSE_GB_NM        = iCurColumnPos(2)
	C_TARGET             = iCurColumnPos(3)
	C_CLOSE_MNTH         = iCurColumnPos(4)
	C_WORKING_DT         = iCurColumnPos(5)
	C_WORKING_ID         = iCurColumnPos(6)
	C_TAGET_WORKING_MNTH = iCurColumnPos(7)
	C_CLOSE              = iCurColumnPos(8)
	C_CANCEL             = iCurColumnPos(9)

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
'	Call InitData()
End Sub

'======================================================================================================
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'=======================================================================================================

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

    Call InitSpreadSheet													'��: Setup the Spread sheet
	Call InitVariables
    Call InitComboBox

    Call SetDefaultVal
    Call SetToolbar("1100000000001111")										'��: ��ư ���� ���� 

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
'   Event Desc : Spread Split �����ڵ� 
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

	ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
End Sub

'======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
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

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then		'��: ������ üũ 
		If lgStrPrevKey <> "" Then								'���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			DbQuery
		End If
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : 
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)
    If frm1.vspdData.MaxRows = 0 Then								'no data�� ��� vspdData_LeaveCell no ���� 
		Exit Sub													'tab�̵��ÿ� �߸��� 140318 message ���� 
    End If
    
    With frm1.vspdData
		 If NewCol > 0 Then 
			If Col = C_TAGET_WORKING_MNTH Then
				.Row = Row
				.Col = Col
			
				If .Text <> "" Then
                    If CheckDateFormat(.Text, parent.gDateFormatYYYYMM) = False  Then
						Call DisplayMsgBox("140318","X","X","X")	'����� �ùٷ� �Է��ϼ���.
						.Text = ""
					End If
				End If
			End If
		
		End If
    End With
End Sub

'======================================================================================================
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'=======================================================================================================

'======================================================================================================
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'=======================================================================================================
'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery()
    Dim IntRetCD 
	
	FncQuery = False                                                        '��: Processing is NG
	
	Err.Clear																'��: Protect system from crashing
	
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
	End If																	'��: Query db data
	
	FncQuery = True															'��: Processing is OK
End Function

'======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'=======================================================================================================
Function FncNew() 
	On Error Resume Next                                                    '��: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncDelete() 
	On Error Resume Next                                                    '��: Protect system from crashing
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

    If lgBlnFlgChgValue = False And var1 = False Then									'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")									'��: Display Message(There is no changed data.)
		Exit Function
    End If

    If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
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
		Exit Function																	'��: Save db data
    End If
        
	FncSave = True                                      								'��: Processing is OK
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy() 
	On Error Resume Next                                           	       '��: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel() 
	On Error Resume Next                                           	       '��: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow() 
	On Error Resume Next                                           	       '��: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
	On Error Resume Next                                           	       '��: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)						'��: ȭ�� ���� 
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
	Call parent.FncFind(Parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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

	Err.Clear                                                               			'��: Protect system from crashing
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
		
	DbQuery = False

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		
	Call RunMyBizASP(MyBizASP, strVal)												'��: �����Ͻ� ASP �� ���� 
		
	DbQuery = True																	'��: Processing is NG
End Function

'======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'=======================================================================================================
Function DbQueryOk()					'��: ��ȸ ������ ������� 
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE			'��: Indicates that current mode is Update mode
	lgBlnFlgChgValue = False
	
	Call SetSpreadColor()
    Call SetToolbar("1100100000001111")										'��: ��ư ���� ���� 	
End Function

Function VerifySelCount()
	Dim ii 
	Dim ChkCnt

	ChkCnt = 0
	With frm1
		For ii = 1 To .vspddata.maxrows
			.vspddata.row = ii 
			.vspddata.col = C_CLOSE

			If .vspddata.text = "1" Then
				ChkCnt = ChkCnt + 1
			End If
			
			.vspddata.col = C_CANCEL
			If .vspddata.text = "1" Then
				ChkCnt = ChkCnt + 1
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

	DbSave = False																		'��: Processing is NG

	On Error Resume Next																		'��: Protect system from crashing
	Err.Clear

	With frm1
		.txtMode.value  = parent.UID_M0002

		'-----------------------
		'Data manipulate area
		'-----------------------
   		strVal = ""
		    		
		For lRow = 1 To .vspdData.MaxRows
	    	.vspdData.Row = lRow
			.vspddata.col = 0
			
			If Trim(.vspddata.Text) = ggoSpread.UpdateFlag Then
				.vspdData.Col = C_CLOSE_GB
				If UCase(Trim(.vspdData.Text)) = "CV" Then 
					.vspdData.Col = C_TAGET_WORKING_MNTH			
					tmpVal = Trim(.vspdData.Text)
				End If	
			
				If UCase(Trim(.vspdData.Text)) = "MC" Or UCase(Trim(.vspdData.Text)) = "CC" Or UCase(Trim(.vspdData.Text)) ="AC" Then
					.vspdData.Col = C_CLOSE_GB
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_CLOSE_MNTH
					If Trim(.vspdData.Text) = "" Then
						Call ExtractDateFrom(tmpVal,parent.gDateFormatYYYYMM,parent.gComDateType,strYear,strMonth,strDay)
					Else
						Call ExtractDateFrom(.vspdData.Text,parent.gDateFormatYYYYMM,parent.gComDateType,strYear,strMonth,strDay)
					End If	
					strVal = strVal & strYear & strMonth & Parent.gColSep
					.vspdData.Col = C_CLOSE
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_CANCEL
					strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
				End If
			End If	
		Next

		.txtSpread.value =  strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)													'��: �����Ͻ� ASP �� ���� 
	End With
	
	DbSave = True																			
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function DbSaveOk()				            '��: ���� ������ ���� ���� 
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
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>������Ȳ</font></td>
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
						<script language =javascript src='./js/c3980ma1_vaSpread1_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX= "-1"></IFRAME>
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

