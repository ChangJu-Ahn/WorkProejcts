
<%@ LANGUAGE="VBSCRIPT" %>

<!--'========================================================================================================
'*  1. Module Name          : Finance
'*  2. Function Name        : F_Notes
'*  3. Program ID           : f5102ma1
'*  4. Program Name         : ������ǥ��ȣ��� 
'*  5. Program Desc         : ����/��ǥå�� ���/����/����/��ȸ 
'*  6. Modified date(First) : 2000/09/22
'*  7. Modified date(Last)  : 2002/09/07
'*  8. Modifier (First)     : hersheys
'*  9. Modifier (Last)      : Shin Myoung Ha
'* 10. Comment              : 1. (ǥ��) FilterVar()�Լ� ���� - 2002/07/31
'*							  2. �����ÿ� �߽����࿡ Ư������ "'" �� ������ �����޼��� ��¾ȵ� 
'*							  3. FilterVar()�Լ� ����(Com���� ������) - 2002/08/08
'*							  4. ���翵���� ���̴� ���� ����, ��¥,����OCX TEXT�� VALUE �߸��Ȼ�� ���� - 2002/08/09
'*                            5. �ؽ�ƮŰ���� ���� ��ȸ���¿����� ���������� ��ȸ������ �ؽ�ƮŰ���� 
'*                               �������� ��ȸ������ ������ ���� ������ �����͵� ���� ��ȸ��(������) - 2002/09/06
'*							  6. ��Ƽ ��� ��Ұ����ϵ��� ������ - 2002/09/06
'*							  7. ��ü���� �����ÿ� �������ؿ��� �ʵ��� ������ - 2002/09/07
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'=======================================================================================================
'												1. �� �� �� 
'=======================================================================================================

'=======================================================================================================
'                                               1.1 Inc ����   
'	���: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>


<SCRIPT LANGUAGE="VBScript">

Option Explicit							'��: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global ����/��� ����  
'	.Constant�� �ݵ�� �빮�� ǥ��.
'	.���� ǥ�ؿ� ����. prefix�� g�� �����.
'	.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Const BIZ_PGM_ID = "f5102mb1.asp"										'�����Ͻ� ���� ASP�� 
Const JUMP_PGM_ID_NOTE_INF = "f5101ma1"									'����������� 

 
Dim C_NOTE_KIND_NM
Dim C_NOTE_KIND
Dim C_BANK_CD	
Dim C_BANK_PB	
Dim C_BANK_NM	
Dim C_NOTE_NO	
Dim C_ISSUE_DT	
Dim C_STS		
Dim C_STS_NM	
Dim C_COL_END	

Dim IsOpenPop
Dim glDeletedRow

'=======================================================================================================
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'=======================================================================================================

'=======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

sub initSpreadPosVariables()

	C_NOTE_KIND_NM= 1
	C_NOTE_KIND	= 2
	C_BANK_CD		= 3
	C_BANK_PB		= 4
	C_BANK_NM		= 5
	C_NOTE_NO		= 6
	C_ISSUE_DT	= 7
	C_STS			= 8
	C_STS_NM		= 9
	C_COL_END		= 10

end sub

'=======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()
	frm1.txtIssueDt.text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat) 
End Sub


'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("A","*","NOCOOKIE","MA") %>

End Sub


'=======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()

    Call initSpreadPosVariables()
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    
		.ReDraw = False
		.MaxCols = C_COL_END												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	    .Col = .MaxCols														'������Ʈ�� ��� Hidden Column
		.ColHidden = True
	    .MaxRows = 0

        Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCombo	C_NOTE_KIND,	"",					15
	    ggoSpread.SSSetCombo	C_NOTE_KIND_NM,	"������ǥ����",	15
		ggoSpread.SSSetEdit		C_BANK_CD,		"��������",		15, , , 10
	    ggoSpread.SSSetButton	C_BANK_PB
		ggoSpread.SSSetEdit		C_BANK_NM,		"���������",	29
	    ggoSpread.SSSetEdit		C_NOTE_NO,		"������ǥ��ȣ",	25, , , 30
		ggoSpread.SSSetDate		C_ISSUE_DT,		"������",		15,	2, Parent.gDateFormat
	    ggoSpread.SSSetCombo	C_STS,			"",					15
		ggoSpread.SSSetCombo	C_STS_NM,		"����",			15
 
        call ggoSpread.MakePairsColumn(C_BANK_CD,C_BANK_PB)
        Call ggoSpread.SSSetColHidden(C_NOTE_KIND,C_NOTE_KIND,True)
        Call ggoSpread.SSSetColHidden(C_STS,C_STS,True)
		.ReDraw = True

    End With
    
    Call SetSpreadLock
    CALL InitSpreadCombo
    
End Sub

Sub InitSpreadCombo()

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", "MAJOR_CD=" & FilterVar("f1001", "''", "S") & " "  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_NOTE_KIND			'�ڵ� 
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_NOTE_KIND_NM		'�̸� 
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", "MAJOR_CD=" & FilterVar("f1002", "''", "S") & " "  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_STS			'�ڵ� 
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_STS_NM		'�̸� 
        
End Sub


'=======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()

    With frm1
		.vspdData.ReDraw = False
		
		ggoSpread.SpreadLock	        C_NOTE_KIND_NM,	-1, C_NOTE_KIND_NM
		ggoSpread.SpreadLock	        C_BANK_NM,		-1, C_BANK_NM
		ggoSpread.SpreadLock	        C_NOTE_NO,		-1, C_NOTE_NO
		ggoSpread.SSSetRequired		C_BANK_CD,	    -1, -1
		ggoSpread.SSSetRequired		C_ISSUE_DT,	    -1, -1
		ggoSpread.SSSetProtected	    C_STS_NM,	    -1, -1

		.vspdData.ReDraw = True
    End With
    
End Sub


'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
  
    With frm1
		.vspdData.ReDraw = False
    	ggoSpread.SSSetProtected	    C_BANK_NM,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	    C_STS_NM,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_NOTE_KIND_NM,	pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_BANK_CD,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_NOTE_NO,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired		C_ISSUE_DT,		pvStartRow, pvEndRow
		.vspdData.ReDraw = True

    End With
End Sub

'--------------------------------------------------------------
' ComboBox �ʱ�ȭ 
'-------------------------------------------------------------- 
Sub InitComboBox()
	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1001", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboNoteKind ,lgF0  ,lgF1  ,Chr(11))
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1002", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboSts ,lgF0  ,lgF1  ,Chr(11))
    
End Sub

'======================================================================================================
'	Name : OpenPopupBank()
'	Description : Bank Code Popup
'=======================================================================================================
Function OpenPopupBank(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = "�����˾�"				'�˾� ��Ī 
	arrParam(1) = "B_BANK"						'TABLE ��Ī 
	arrParam(2) = strCode						'Code Condition
	arrParam(3) = ""							'Name Cindition
	arrParam(4) = ""							'Where Condition
	arrParam(5) = "����"			
	
	arrField(0) = "BANK_CD"						'Field��(0)
	arrField(1) = "BANK_NM"						'Field��(1)
    
	arrHeader(0) = "�����ڵ�"				'Header��(0)
	arrHeader(1) = "�����"					'Header��(1)
			
	IsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=430px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		With frm1
			Select Case iWhere
				Case "0"
					.txtBankCd.focus
				Case "1"
					Call SetActiveCell(.vspdData,C_BANK_CD,.vspdData.ActiveRow ,"M","X","X")
				Case Else
					Exit Function
			End Select
		End With
		Exit Function
	End If
	
	With frm1
		Select Case iWhere
			Case "0"
				.txtBankCd.value = arrRet(0)
				.txtBankNm.value = arrRet(1)
				.txtBankCd.focus
				
			Case "1"
				.vspdData.Col  = C_BANK_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_BANK_NM
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(.vspdData.Col, .vspdData.Row)
				Call SetActiveCell(.vspdData,C_BANK_CD,.vspdData.ActiveRow ,"M","X","X")
				
			Case Else
				Exit Function
		End Select
	End With
End Function

'======================================================================================================
'	Name : OpenPopupNote()
'	Description : NoteNo PopUp
'=======================================================================================================
Function OpenPopupNote(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	arrParam(0) = "������ȣ�˾�"			'�˾� ��Ī 
	arrParam(1) = "F_NOTE_NO A, B_MINOR B"		'TABLE ��Ī 
	arrParam(2) = strCode						'Code Condition
	arrParam(3) = ""							'Name Cindition
	arrParam(4) = "A.NOTE_KIND = B.MINOR_CD AND B.MAJOR_CD = " & FilterVar("F1001", "''", "S") & "  "							'Where Condition
	arrParam(5) = "������ǥ��ȣ"			
	
	arrField(0) = "A.NOTE_NO"					'Field��(0)
	arrField(1) = "B.MINOR_NM"					'Field��(1)
    
	arrHeader(0) = "������ǥ��ȣ"			'Header��(0)
	arrHeader(1) = "������ǥ����"			'Header��(1)

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtNoteNo.focus
		Exit Function
	End If	

	With frm1
		.txtNoteNo.value = arrRet(0)
		.txtNoteNo.focus
	End With
	
End Function

'======================================================================================================
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'=======================================================================================================

'========================================================================================================
'	Desc : Cookie Setting
'========================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp

    Select Case Kubun		

	       Case "FORM_LOAD"
			     strTemp = ReadCookie("NOTE_NO")
			     Call WriteCookie("NOTE_NO", "")
			
			     If strTemp = "" then Exit Function
        			frm1.txtNoteNoQry.value = strTemp
	            		If Err.number <> 0 Then
				        Err.Clear
				        Call WriteCookie("NOTE_NO", "")
				        Exit Function 
		         End If
				
		         Call MainQuery()

		   Case JUMP_PGM_ID_NOTE_INF	'����������� 
		
				With frm1.vspdData
				.Row = .ActiveRow
				.Col = C_NOTE_NO
				strTemp = .Text
			    End With
                If frm1.vspdData.ActiveRow = 0 then Exit Function
		        Call WriteCookie("NOTE_NO", strTemp)
	
     	   Case Else
			    Exit Function		
	End Select
End Function	

'========================================================================================================
'	Desc : ȭ���̵� 
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD
	
	If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")			    '����Ÿ�� ����Ǿ����ϴ�. ����Ͻðڽ��ϱ�?
			If IntRetCD = vbNo Then
				Exit Function
			End If
		End If
	    
	    Call CookiePage(strPgmId)
	    Call PgmJump(strPgmId)
	End Function

'======================================================================================================
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'=======================================================================================================

'==========================================================
'���/������ ��ư Ŭ�� 
'==========================================================
Function FnButtonExec(strMode)
	Dim IntRetCD
    Dim strVal, strNoteNo
    Dim lGrpCnt
    Dim lRow
	
	'-----------------------
	'Check previous data area
	'----------------------- 
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		Call DisplayMsgBox("141427","X","X","X")	'�����Ͱ� ����Ǿ����ϴ�. �����ϰų� ����� ���� �۾��ϼ���.
      	Exit Function
    End If

	With frm1.vspdData
		If .MaxRows <= 0 Then
			Call DisplayMsgBox("900025","X","X","X")	'���õ� �׸��� �����ϴ�.
			Exit Function
		End If
		
		.Row = .ActiveRow
		.Col = C_NOTE_NO		
		strNoteNo = .Text
	End With
    
	'-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")	'�۾��� �����Ͻðڽ��ϱ�?
    IF IntRetCD = vbNO Then
		Exit Function
	End IF	
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '��: Protect system from crashing
    
    With frm1
		.txtMode.value = Parent.UID_M0002

		'-----------------------
		'Data manipulate area
		'----------------------- 
		lGrpCnt = 1
		strVal = ""
		
		'For lRow = 1 To .vspdData.MaxRows 
			'.vspdData.Row = lRow
			'.vspdData.Col = 0
			
			'if .vspdData.IsCellSelected(0, lRow) Then
				strVal = strVal & strMode & Parent.gColSep & .vspdData.Row & Parent.gColSep 			'��: strMode=���/������, Row��ġ ���� 
				.vspdData.Col = C_BANK_CD
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				.vspdData.Col = C_NOTE_NO        
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				.vspdData.Col = C_NOTE_KIND
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				.vspdData.Col = C_STS                
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				.vspdData.Col = C_ISSUE_DT
				strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep				
				lGrpCnt = lGrpCnt + 1
			'end if		
		'Next				
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value  = strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'��: �����Ͻ� ASP �� ���� 

    End With
    
End Function

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     'Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field                         
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitSpreadSheet                                                    'Setup the Spread sheet
    Call InitVariables                                                      'Initializes local global variables
    Call SetDefaultVal
    Call InitComboBox
	Call SetToolbar("110011010010111")										'��ư ���� ���� 

    frm1.txtBankCd.focus 
    Set gActiveElement = document.activeElement
End Sub


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
            C_NOTE_KIND_NM= iCurColumnPos(1)
            C_NOTE_KIND	= iCurColumnPos(2)
            C_BANK_CD		= iCurColumnPos(3)
            C_BANK_PB		= iCurColumnPos(4)
            C_BANK_NM		= iCurColumnPos(5)
            C_NOTE_NO		= iCurColumnPos(6)
            C_ISSUE_DT	= iCurColumnPos(7)
            C_STS			= iCurColumnPos(8)
            C_STS_NM		= iCurColumnPos(9)
            C_COL_END		= iCurColumnPos(10)
            
    End Select    
End Sub
'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'======================================================================================================
'   Event Name : txtIssueDt_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtIssueDt_DblClick(Button)
	if Button = 1 then
		frm1.txtIssueDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtIssueDt.Focus
	End if
End Sub

Sub txtIssueDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery
	End If   
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    
    Set gActiveSpdSheet = frm1.vspdData
    
    If Row = 0 Then
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

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
     Dim iColumnName
    
    If Row <= 0 Then
      Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
       Exit Sub
    End If     
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_BANK_PB Then
        .Col = Col
        .Row = Row
        
        Call OpenPopupBank(.Text, "1")
    End If
    
    End With
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'=======================================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	 '----------  Coding part  -------------------------------------------------------------   
	 ' �� Template ȭ�鿡���� ���� ������, �޺�(Name)�� ����Ǹ� �޺�(Code, Hidden)�� ��������ִ� ���� 
	With frm1.vspdData
		.Row = Row
    
		Select Case Col
			Case  C_NOTE_KIND_NM
				.Col = Col
				intIndex = .Value
				.Col = C_NOTE_KIND
				.Value = intIndex
			Case  C_STS_NM
				.Col = Col
				intIndex = .Value
				.Col = C_STS
				.Value = intIndex
			Case  C_STS
				.Col = Col
				intIndex = .Value
				.Col = C_STS_NM
				.Value = intIndex
		End Select
	End With
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
        
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ 
    	If lgStrPrevKey <> "" Then                  '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DbQuery
    	End If

    End if
    
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_NOTE_KIND_NM Or NewCol <= C_NOTE_KIND_NM Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


'=======================================================================================================
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'=======================================================================================================


'=======================================================================================================
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'=======================================================================================================


'=======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               'Protect system from crashing

	'-----------------------
	'Check previous data area
	'----------------------- 
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If    	
    End If
    
  '-----------------------
    'Erase contents area
    '----------------------- 
   Call ggoOper.ClearField(Document, "2")
   ggoSpread.Source = frm1.vspdData
   ggospread.ClearSpreadData
   Call InitVariables                                                      'Initializes local global variables
    															
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not ChkField(Document, "1") Then								'This function check indispensable field
       Exit Function
    End If
    
  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery
           
    FncQuery = True															
	Set gActiveElement = document.activeElement    
	
End Function


'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False
    
    Err.Clear                                                               '��: Protect system from crashing
    'On Error Resume Next                                                   '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData										'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then									'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")					'��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData										'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then							'��: Check required field(Multi area)
       Exit Function
    End If
    
	'-----------------------
    'Save function call area
    '----------------------- 
    Call DbSave				                                                '��: Save db data
    
    FncSave = True                                                          
	Set gActiveElement = document.activeElement    
	
End Function


'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy() 
	With frm1.vspdData
	
		If .MaxRows < 1 Then Exit Function
	
		.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		Call SetSpreadColor(.ActiveRow, .ActiveRow)
    
		.ReDraw = True
		.Focus
	End With
	Set gActiveElement = document.activeElement
End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel() 

	If frm1.vspdData.MaxRows < 1 Then 		
		Exit Function
	end if
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo
    
    If frm1.vspdData.MaxRows < 1 Then 				
		Call SetToolbar("110011010010111")										'��ư ���� ����                                  
		Exit Function
	end if
    
	Set gActiveElement = document.activeElement
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow(Byval pvRowcnt) 
	Dim IntRetCD
    Dim imRow
    Dim CurRow
    On Error Resume Next                                                          '��: If process fails
    Err.Clear   

    FncInsertRow = False                                                         '��: Processing is NG

     If IsNumeric(Trim(pvRowcnt)) Then 
    
       imRow  = Cint(pvRowcnt)
       
       else

    imRow = AskSpdSheetAddRowCount()
    If imRow = "" Then
        Exit Function
    End If
    End If                              
	
	ggoSpread.Source = frm1.vspdData
	Call SetToolbar("110011110010111")										'��ư ���� ���� 
	
	With frm1.vspdData

		.ReDraw = False
		
		ggoSpread.InsertRow, imRow

		Call SetSpreadColor(.ActiveRow, .ActiveRow + imRow - 1)
		
		For CurRow = .ActiveRow To .ActiveRow + imRow - 1
			.Col = C_STS		' Default�� '����' Setting
		    .Row = CurRow 
			.Text = "NP"
			Call vspdData_ComboSelChange(C_STS, CurRow)
		Next
		.ReDraw = True
		.Focus
    End With
    Set gActiveElement = document.activeElement
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

	If frm1.vspdData.MaxRows < 1 Then 
		Call SetToolbar("110011010010111")										'��ư ���� ���� 
		Exit Function
	end if
	
	ggoSpread.Source = frm1.vspdData 
	lDelRows = ggoSpread.DeleteRow
	
	If frm1.vspdData.MaxRows < 1 Then 
		Call SetToolbar("110011010010111")										'��ư ���� ���� 
		Exit Function
	end if
	
	frm1.vspdData.focus
    Set gActiveElement = document.activeElement
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   
	Set gActiveElement = document.activeElement
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
    Call parent.FncExport(Parent.C_MULTI)											 '��: ȭ�� ���� 
	Set gActiveElement = document.activeElement
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                      '��:ȭ�� ����, Tab ���� 
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)  
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
    Call InitSpreadCombo()
	Call ggoSpread.ReOrderingSpreadData()
	'Call InitData()
End Sub

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()

    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
    
    Set gActiveElement = document.activeElement
End Function

'======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear																	'��: Protect system from crashing

	Dim strVal
    
    With frm1
    
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode		=" & Parent.UID_M0001						'Hidden�� �˻��������� Query
			strVal = strVal & "&txtBankCd		=" & .hBankCd.value				
			strVal = strVal & "&lgStrPrevKey	=" & lgStrPrevKey
			strVal = strVal & "&cboNoteKind		=" & .hNoteKind.value
			strVal = strVal & "&txtNoteNo		=" & .hNoteNo.value		
			strVal = strVal & "&txtIssueDt		=" & .hIssueDt.value
			strVal = strVal & "&txtSts			=" & .hSts.value
			strVal = strVal & "&txtMaxRows		=" & .vspdData.MaxRows		
		Else
			strVal = BIZ_PGM_ID & "?txtMode		=" & Parent.UID_M0001						'���� �˻��������� Query
			strVal = strVal & "&txtBankCd		=" & .txtBankCd.value				
			strVal = strVal & "&lgStrPrevKey	=" & lgStrPrevKey
			strVal = strVal & "&cboNoteKind		=" & Trim(.cboNoteKind.value)
			strVal = strVal & "&txtNoteNo		=" & .txtNoteNo.value 
			strVal = strVal & "&txtIssueDt		=" & .txtIssueDt.text
			strVal = strVal & "&txtSts			=" & Trim(.cboSts.value)
			strVal = strVal & "&txtMaxRows		=" & .vspdData.MaxRows		
		End If

		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With   
    
    DbQuery = True
    
End Function

'======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'=======================================================================================================
Function DbQueryOk()													'��ȸ ������ ������� 
	frm1.vspdData.Redraw = False
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
	
	If frm1.vspdData.MaxRows < 1 Then 
		Call SetToolbar("110011010010111")										'��ư ���� ���� 
	else
		Call SetToolbar("110011110011111")										'��ư ���� ����		
	end if
	
	Call InitData()
	
	'SetGridFocus
	
	frm1.vspdData.Redraw = True
End Function

Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row = intRow
			
			.Col = C_STS
			intIndex = .value
			.col = C_STS_NM
			.value = intindex
			
			.Col = C_NOTE_KIND
			intIndex = .value
			.col = C_NOTE_KIND_NM
			.value = intindex
		Next	
	End With
End Sub

'======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbSave() 
    Dim lRow
    Dim lGrpCnt
	Dim strVal, strDel
	
    DbSave = False                                                          
    
    Call LayerShowHide(1)
    

	With frm1
		.txtMode.value = Parent.UID_M0002
    
	'-----------------------
	'Data manipulate area
	'----------------------- 
    lGrpCnt = 1
    strVal = ""
    strDel = ""
    glDeletedRow = 0
	'-----------------------
	'Data manipulate area
	'----------------------- 
    ' Data ���� ��Ģ 
    ' 0: Flag , 1: Row��ġ, 2~N: �� ����Ÿ   
    
    For lRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = lRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag									    '��: �ű� 

				strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep 			'��: C=Create, Row��ġ ���� 
                .vspdData.Col = C_BANK_CD
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_NOTE_NO                
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_NOTE_KIND
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_STS                
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_ISSUE_DT                
		        strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & Parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
			Case ggoSpread.UpdateFlag										'��: ���� 
					
				strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep 			'��: U=Update, Row��ġ ���� 
                .vspdData.Col = C_BANK_CD
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_NOTE_NO
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_NOTE_KIND
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_STS                
                strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_ISSUE_DT
		        strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & Parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag										'��: ���� 

				strDel = strDel & "D" & Parent.gColSep & lRow & parent.gColSep			'��: D=Delete, Row��ġ ���� 
                .vspdData.Col = C_BANK_CD
                strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_NOTE_NO
                strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_NOTE_KIND
                strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_STS
                strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_ISSUE_DT
		        strDel = strDel & UNIConvDate(Trim(.vspdData.Text)) & Parent.gRowSep
                
                lGrpCnt = lGrpCnt + 1
                glDeletedRow = glDeletedRow + 1
                
        End Select
        
    Next   

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal	
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True                                                           
    
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function DbSaveOk()													        ' ���� ������ ���� ���� 
	Call InitVariables
	
	If glDeletedRow = frm1.vspdData.MaxRows Then
		frm1.vspdData.MaxRows = 0
	Else
		frm1.vspdData.MaxRows = 0		
		Call MainQuery()
	End If
End Function

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc : 
'=======================================================================================================    
Sub SetGridFocus()	
   
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
		
End Sub


'=======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<!--
'======================================================================================================
'       					6. Tag�� 
'	���: Tag�κ� ���� 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������ǥ��ȣ���</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBankCd" SIZE=10 MAXLENGTH=10  tag="12XXXU" ALT="�����ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopupBank(frm1.txtBankCd.value, '0')">&nbsp;<INPUT TYPE=TEXT NAME="txtBankNm" SIZE=25 tag="14" ALT="�����"></TD>
									<TD CLASS="TD5" NOWRAP>������ǥ����</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboNoteKind" tag="12" STYLE="WIDTH: 100px;" ALT="������ǥ����"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>������ǥ��ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="txtNoteNo" TYPE=TEXT NAME="txtNoteNo" SIZE=20 MAXLENGTH=30 tag="11XXXU" ALT="������ǥ��ȣ" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnNoteNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopupNote frm1.txtNoteNo.value"></TD>
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/f5102ma1_I841659063_txtIssueDt.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboSts" tag="11X1" STYLE="WIDTH: 100px;" ALT="���౸��"><OPTION value=""></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
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
								<TD HEIGHT="100%" NOWRAP>
									<script language =javascript src='./js/f5102ma1_I758642280_vspdData.js'></script>
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
	<TR HEIGHT=20>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<BUTTON NAME="btnExecDu" CLASS="CLSSBTN" OnClick="VBScript:Call FnButtonExec('X')" Flag=1>���</BUTTON>&nbsp;
						<BUTTON NAME="btnExecCn" CLASS="CLSSBTN" OnClick="VBScript:Call FnButtonExec('Z')" Flag=1>������</BUTTON>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>
						<A HREF="VBSCRIPT:PgmJumpChk(JUMP_PGM_ID_NOTE_INF)">�����������</A>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT= <%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="txtMode"   tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="hBankCd"   tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="hNoteKind" tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="hNoteNo"   tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="hIssueDt"  tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="hSts"      tag="24" TABINDEX = "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

