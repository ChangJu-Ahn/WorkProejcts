<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : ������������ 
'*  3. Program ID           : c4101ba1
'*  4. Program Name         : �������� ��� 
'*  5. Program Desc         : �������� ��� 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/11/13
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : HJO
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

Option Explicit											'��: indicates that All variables must be declared in advance 

'=======================================================================================================
'                                               1.2 Global ����/��� ����  
'	.Constant�� �ݵ�� �빮�� ǥ��.
'	.���� ǥ�ؿ� ����. prefix�� g�� �����.
'	.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=======================================================================================================
'@PGM_ID
Const BIZ_PGM_QRY_ID = "c4101bb9.asp"					'��ȸ �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "c4101bb1.asp"


'@Grid_Column
Dim C_ChkFlag 

Dim C_MinorCd 										'Spread Sheet�� Column�� ���� 
Dim C_MinorNM 
Dim C_PrgYn 
Dim C_UsrId
Dim C_WorkDt
Dim C_ErrCnt
Dim C_ErrPop
Dim C_Reference 

'--- Karrman_ADO
'Const Parent.DISCONNUPD  = "1"										'Disconnect + Update Mode
'Const Parent.DISCONNREAD = "2"										'Disconnect + ReadOnly Mode
'---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------


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
    lgIntGrpCount = 0                           'initializes Group View Size
    
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()
	Dim StartDate
	Dim EndDate
	StartDate	= "<%=GetSvrDate%>"
	EndDate		= UNIDateAdd("m", -1, StartDate,Parent.gServerDateFormat)
	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
    Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "BA")%>
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	 C_ChkFlag		= 1
	 
	 C_MinorCd		= 2										'Spread Sheet�� Column�� ��� 
	 C_MinorNM		= 3
	 C_PrgYn		= 4
	 C_UsrId		=5
	C_WorkDt	=6
	C_ErrCnt=7
	C_ErrPop=8
	 'C_Reference	= 9
End Sub


'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
	With frm1.vspdData
		
		.ReDraw = false
		
		.MaxCols = C_ErrPop + 1			'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols								'��: ������Ʈ�� ��� Hidden Column
		.ColHidden = True

		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021218", ,parent.gAllowDragDropSpread
		ggoSpread.ClearSpreadData

	   .ReDraw = false
	
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetCheck C_ChkFlag, "���౸��", 15, ,"",true    		
		ggoSpread.SSSetEdit		C_MinorCd, "�۾��ܰ��ڵ�",15,0,,10,2
		ggoSpread.SSSetEdit		C_MinorNm, "�۾��ܰ��",30,0,,50,2
		'ggoSpread.SSSetCombo C_PrgYn, "���۾�����",15 ,0
		ggoSpread.SSSetEdit		C_PrgYn, "�۾�����",10,0,,10,2
		ggoSpread.SSSetEdit		C_UsrId, "�۾���",15,0,,50,2
		'ggoSpread.SSSetDate	C_WorkDt, "�۾��Ͻ�", 15, 0
		ggoSpread.SSSetEdit	C_WorkDt, "�۾��Ͻ�", 20, 0
		ggoSpread.SSSetFloat    C_ErrCnt,      "ERROR COUNT",   15,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetButton  C_ErrPop
		
		'ggoSpread.SSSetEdit C_Reference, "", 20, 0, -1, 40
		
		Call ggoSpread.MakePairsColumn(C_ErrCnt,C_ErrPop)
		'Call ggoSpread.SSSetColHidden(C_Reference,C_Reference,True)
		
		.ReDraw = true
		
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
		ggoSpread.SpreadLock C_MinorCd, -1, -1
		ggoSpread.SpreadLock C_MinorNm, -1, C_MinorNm
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		.vspdData.ReDraw = True
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

	Dim	lRow
	
	With frm1
		.vspdData.ReDraw = False
		
		ggoSpread.SSSetProtected C_MinorCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_MinorNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_PrgYn, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_UsrId, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_WorkDt, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ErrCnt, pvStartRow, pvEndRow

		' ����ǰ ������̹ݿ�, ����ǰ ������� �ݿ�, �԰����� �ݿ� �� ���� Protect
		For lRow = 1 To .vspdData.MaxRows

			.vspdData.Row = lRow
			.vspdData.Col = C_MinorCd
			
			if .vspdData.value = "13" or .vspdData.value = "14" or .vspdData.value = "15"  then
				ggoSpread.SSSetProtected C_ChkFlag, lRow, lRow
				ggoSpread.SSSetProtected C_PrgYn, lRow, lRow
			End if
		Next

		.vspdData.ReDraw = True
	End With
End Sub

 '==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'============================================================================================================
Sub InitComboBox()

	ggoSpread.source = frm1.vspdData
	ggoSpread.SetCombo "Y" & vbtab & "N" , C_PrgYn
	

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
' Function Name : FncBtnExe
' Function Desc : This function is related to BtnExe
'=======================================================================================================
Function FncBtnExe() 
	Dim IntRetCD 
	
	FncBtnExe = False                                                  		       '��: Processing is NG

	Err.Clear                                                            	 		  '��: Protect system from crashing
	
	On Error Resume Next                                           	       '��: Protect system from crashing

	if SpreadWorkingChk = false then  Exit Function      'spread check box üũ ���� 

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	'-----------------------
	'Check content area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	
	If Not chkField(Document, "1")  Then  '��: Check contents area
		Exit Function
	End If
    	
	'-----------------------
	'Save function call area
	'----------------------- 	
	IF DbSave = False Then
		Exit Function				                                                  '��: Save db data
    END IF
    
	FncBtnExe = True                                      	                    '��: Processing is OK
End Function

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
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ChkFlag				= iCurColumnPos(1)			
			C_MinorCd     			= iCurColumnPos(2)    
			C_MinorNM       		= iCurColumnPos(3)
			C_PrgYn					= iCurColumnPos(4)
			C_UsrId					= iCurColumnPos(5)
			C_WorkDt				= iCurColumnPos(6)
			C_ErrCnt				= iCurColumnPos(7)
			C_ErrPop				= iCurColumnPos(8)
			'C_Reference       		= iCurColumnPos(9)
    End Select    
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
    'Call InitComboBox
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
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	
    Call InitSpreadSheet 
	Call InitVariables                                                     '��: Setup the Spread sheet
    'Call InitComboBox

    Call SetDefaultVal
    Call SetToolbar("11000000000011")										'��: ��ư ���� ���� 
    
		frm1.txtYyyymm.focus
		   
'    FncQuery

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

'	ggoSpread.UpdateRow Row
	 '----------  Coding part  -------------------------------------------------------------
	
	

End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row

	End With
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	 '----------  Coding part  -------------------------------------------------------------   
	
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

	 '----------  Coding part  -------------------------------------------------------------   

	End With
End Sub

'======================================================================================================
'   Event Name : txtYyyymm_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
	End If
End Sub
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    Dim yyyymm
    Dim work_Step
    Dim  strYear,strMonth,strDay,strYYYYMM
  
    frm1.vspdData.Row = Row
	Select Case Col
        Case C_ErrPop
'            frm1.vspdData.Col = C_ERR_CNT
'            If CInt(frm1.vspdData.Text) > 0 Then

				Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

				strYYYYMM = strYear & strMonth	
				
				
                frm1.vspdData.Col = C_MinorCd 
                work_step = frm1.vspdData.Text
                'frm1.vspdData.Col = C_YYYYMM
                Call OpenErr(strYYYYMM,work_step)
'            End If
    End Select
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
	END IF
	
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then		'��: ������ üũ 
		If lgStrPrevKey <> "" Then								'���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			DbQuery
		End If
	End If
	
End Sub


Sub txtYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub
'=======================================================================================================
Function OpenErr(yyyymm, work_step)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "BATCH JOB ERROR"		<%' �˾� ��Ī %>
	arrParam(1) = " c_batch_job_error_s "                 <%' TABLE ��Ī %>
	arrParam(2) = "0"'work_step                             <%' Code Condition%>
	arrParam(3) = "" 		            	<%' Name Cindition%>
	arrParam(4) = " YYYYMM = " & FilterVar(yyyymm, "''", "S") & " AND work_step = " & FilterVar(work_step, "''", "S")                       <%' Where Condition%>
	arrParam(5) = "BATCH JOB ERROR"
	
	arrField(0) = "HH" & parent.gColSep & "SEQ_NO"	     			            <%' Field��(1)%>
    arrField(1) = "ED10" & parent.gColSep & "WORK_STEP"	     			            <%' Field��(1)%>
    arrField(2) = "ED10" & parent.gColSep & "SEQ_NO"	     			            <%' Field��(1)%>
    arrField(3) = "ED500" & parent.gColSep & "MSG_TEXT"					<%' Field��(0)%>

	arrHeader(0) = ""			    	    <%' Header��(0)%>
	arrHeader(1) = "�۾��ܰ��ڵ�"			    	    <%' Header��(0)%>
    arrHeader(2) = "SEQ_NO"			    	    <%' Header��(0)%>
    arrHeader(3) = "MSG_TEXT"				<%' Header��(1)%>

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: Yes; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	'Else
'		Call SetCode(arrRet)
	End If

End Function

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
	
	Err.Clear                                                            		   '��: Protect system from crashing
	
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
	Call ggoOper.ClearField(Document, "2")				'��: Clear Contents  Field
	Call InitVariables
	
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	'-----------------------
	'Query function call area
	'-----------------------
	IF DbQuery = False Then
		Exit Function
	END IF						'��: Query db data
	
	FncQuery = True						'��: Processing is OK
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
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
	On Error Resume Next                                           	       '��: Protect system from crashing
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
	Dim strYyyymm
	Dim	strYear, strMonth, strDay

	Err.Clear                                                               			'��: Protect system from crashing
	
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
		
	DbQuery = False

    Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth	

	With frm1
	
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtYyyymm=" & .hYyyymm.value
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtYyyymm=" & strYyyymm
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
		
		DbQuery = True                                                          '��: Processing is NG
	End With

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
	
	Call SetSpreadColor(-1, -1)
	
	Call ggoOper.LockField(Document, "Q")		'��: This function lock the suitable field
	
End Function

'======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbSave() 
	Dim lRow
	Dim lGrpCnt
	Dim strVal
	Dim strDel
	Dim strYyyymm
	Dim	strYear, strMonth, strDay


	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
		
	DbSave = False                                                          '��: Processing is NG
    
	On Error Resume Next                                                   '��: Protect system from crashing

    Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth
    
	With frm1
		
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
    		strVal = ""
    		
    	'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
	    		.vspdData.Row = lRow
			.vspdData.Col = C_ChkFlag
			
			if .vspdData.value = 1  then
					.vspdData.Col = C_MinorCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					strVal = strVal & strYyyymm & Parent.gColSep

					strVal = strVal & CStr(lRow) & Parent.gRowSep	'13
					lGrpCnt = lGrpCnt + 1
			End if
		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value =  strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)					'��: �����Ͻ� ASP �� ���� 
	End With
	
	DbSave = True                                                           '��: Processing is NG
    
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function DbSaveOk()				            '��: ���� ������ ���� ���� 
   	Call InitVariables
	frm1.vspdData.MaxRows = 0
   	Call DisplayMsgBox("990000","X","X","X")
   	Call MainQuery()
End Function

'======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'=======================================================================================================
Function DbDelete()
End Function

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'======================================================================================================
' Function Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================

Function SpreadWorkingChk()
   Dim iRows
   Dim ichkCnt
   Dim IntRetCD

   SpreadWorkingChk = False
   ichkCnt = 0

   with frm1.vspdData
	For iRows = 1 to .MaxRows
	    .Col =  C_ChkFlag
	    .Row =  iRows
	    
	    if .Value = 1 then 
		.Col = C_PrgYn
		'if .Text = "Y" then
		'  IntRetCD = DisplayMsgBox("236020","X","X","X")  '���۾������� Y �� �۾��� ������ �� �����ϴ�.
		'  Exit Function
		'end if
		ichkCnt = ichkCnt + 1
	    end if

	Next
	if ichkCnt = 0 then 
	   IntRetCD = DisplayMsgBox("236021","X","X","X")  '���õ� �۾��� �����ϴ�.
	   Exit Function
	end if
   End With
   
   SpreadWorkingChk = True

 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�����������</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>�۾����</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYyyymm" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT="�۾����" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TDT" NOWRAP>&nbsp</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
				<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnExe()" Flag=1>�� ��</BUTTON>&nbsp;</TD>
				<TD>&nbsp</TD>
				<TD>&nbsp</TD>				
			</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hInspClassCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

