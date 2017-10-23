<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B81106MA1
'*  4. Program Name         : ���ù�������
'*  5. Program Desc         : ���ù�������
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/01/23
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee wol san
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" --> 

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
Option Explicit											

CONST BIZ_PGM_ID = "B81106MB1.ASP"
Const BIZ_PGM_REG_ID = "B81106MA1"
<!-- #Include file="../../inc/lgvariables.inc" -->	

'@Grid_Column
Dim C_FILE_NO
Dim C_INRT_USER_ID

Dim C_USE_DT
Dim C_REQ_DT
Dim C_TITLE
Dim C_FILE_ABBR

Dim C_CurrPopup
Dim C_CostDt
Dim C_SupplierCd
Dim C_SupplierPopup
Dim C_SupplierNm
Dim C_Cost

'@Global_Var
Dim lgSortKey1
Dim IsOpenPop
Dim lgitem_lvl
Dim EndDate, StartDate
Dim lgAcct_item_cd
Dim lgAcct_kind_cd
   
EndDate = "<%=GetSvrDate%>"

'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------


StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", Parent.gServerDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
EndDate = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False
    lgLngCurRows = 0						'initializes Deleted Rows Count
    lgSortKey1 = 2
	lgPageNo = ""
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()
	'Call SetToolBar("110000010011111")				'��ư ���� ���� 
	
	
	frm1.txtFromInsrtDt.text = StartDate	
	frm1.txtToInsrtDt.text = EndDate
	
	frm1.txtFromUseDt.text = EndDate	
	frm1.txtToUseDt.text = UNIDateAdd("m", 3, EndDate, Parent.gDateFormat)
	
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
		.ReDraw = false

		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20050103",, parent.gAllowDragDropSpread

	   .MaxCols = C_FILE_ABBR + 1
	   .MaxRows = 0
		Call GetSpreadColumnPos("A")
		ggoSpread.SSSetEdit C_FILE_NO, "���Ϲ�ȣ", 10
		ggoSpread.SSSetEdit C_INRT_USER_ID, "�ۼ���", 12
		ggoSpread.SSSetEdit C_USE_DT, "��ȿ��", 12
		ggoSpread.SSSetEdit C_REQ_DT, "�����", 12
		ggoSpread.SSSetEdit C_TITLE, "����", 30,,,100
		ggoSpread.SSSetEdit   C_FILE_ABBR, "��༳��", 30,,,100
		
		Call ggoSpread.SSSetColHidden(C_FILE_NO,C_FILE_NO,True)
		Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols,	True)	
		
		.ReDraw = True
    End With

    Call SetSpreadLock()
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
    With frm1.vspdData
		.ReDraw = False
    
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLock		-1,			-1
   		
		.ReDraw = True
    End With    
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1.vspdData
		.ReDraw = False

  		ggoSpread.Source = frm1.vspdData
  
   		ggoSpread.SpreadUnLock		1, pvStartRow, ,pvEndRow
		ggoSpread.SSSetRequired  C_FILE_ABBR,			pvStartRow,	pvEndRow

	    .ReDraw = True
    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_FILE_NO      =   1
	C_INRT_USER_ID      =   2
	C_USE_DT      =   3
	C_REQ_DT      =   4
	C_TITLE		=	5
	C_FILE_ABBR			=	6	

	
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
			C_FILE_NO          =	iCurColumnPos(1)
			C_INRT_USER_ID          =	iCurColumnPos(2)
			C_USE_DT          =	iCurColumnPos(3)
			C_REQ_DT          =	iCurColumnPos(4)
			C_TITLE			=	iCurColumnPos(5)
			C_FILE_ABBR				=	iCurColumnPos(6)	
			
			
			
	End Select    
End Sub


'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'==========================================================================================================
Sub Form_Load()	'###�׸��� ������ ���Ǻκ�###
	Call LoadInfTB19029                                                         'Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                       'Lock  Suitable  Field
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet  
	Call SetToolbar("11000000000111")
	call dbQuery()

	
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 
'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = Col
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)        '  <------����� ǥ�� ���� 

End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	'###�׸��� ������ ���Ǻκ�###
	dim sFile_no
	Dim strval
	Call SetPopupMenuItemInf("1101111111")         'ȭ�麰 ���� 
 	
 	gMouseClickStatus = "SPC" 
 	  
 	sFile_no=GetSpreadText(frm1.vspdData,1,Row,"X","X")
 	
 	Set gActiveSpdSheet = frm1.vspdData
 	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 			
 		Else
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
 		Exit Sub
 	End If
 	
 	with frm1
 	    strVal = BIZ_PGM_ID & "?txtMode=view"
	    strVal = strVal & "&file_no=" & sFile_no
	    strVal = strVal & "&lgPageNo="	 & lgPageNo						'��: Next key tag 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	end with 
	 
	  MyBizASP1.location.href = "B81106RA1.asp?File_no=" & sFile_no
	  
	 
 	//Call RunMyBizASP(MyBizASP, strVal)	//zerry �߾ȵ�..�����Ұ�.	
 	
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     

    If OldLeft <> NewLeft Then Exit Sub
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '��: ������ üũ 
		If Trim(lgPageNo) = "" Then Exit Sub
		If lgPageNo > 0 Then
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End If
End Sub

'========================================================================================
' Function Name : vspdData_ButtonClicked
' Function Desc : �˾���ư ���ý� 
'========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

  
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================================
'   Event Name : vspdData_EditMode
'   Event Desc : 
'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_Cost
            Call EditModeCheck(frm1.vspdData, Row, C_Curr, C_Cost,    "C" ,"I", Mode, "X", "X")
    End Select
End Sub
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'======================================================================================================
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'=======================================================================================================
Sub txtFromInsrtDt_DblClick(Button)
	If Button = 1 Then 
		frm1.txtFromInsrtDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtFromInsrtDt.Focus
	End If
End Sub

Sub txtToInsrtDt_DblClick(Button)
	If Button = 1 Then 
		frm1.txtToInsrtDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtToInsrtDt.Focus
	End If
End Sub

Sub txtFromInsrtDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

Sub txtToInsrtDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

Sub txtFromUseDt_DblClick(Button)
	If Button = 1 Then 
		frm1.txtFromUseDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtFromUseDt.Focus
	end If
End Sub

Sub txtToUsedt_DblClick(Button)
	If Button = 1 Then 
		frm1.txtToUsedt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtToUsedt.Focus
	End If
End Sub

Sub txtFromUsedt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

Sub txtToUsedt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub


'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
  
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() '###�׸��� ������ ���Ǻκ�###
    Dim IntRetCD     
    FncQuery = False                                                        
	
	If ggoSpread.SSCheckChange = True Then 'lgBlnFlgChgValue = True Or lgBtnClkFlg = True Or 
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")		'��: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData
    Call InitVariables															'Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         
       Exit Function
    End If
    If ValidDateCheck(frm1.txtFromInsrtDt, frm1.txtToInsrtDt)	=	False	Then Exit	Function
    If ValidDateCheck(frm1.txtFromUseDt, frm1.txtToUseDt)	=	False	Then Exit	Function
	If DbQuery = False then	Exit Function
		      
    FncQuery = True	
End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
    Err.Clear                                                               '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData	   
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                         '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                         '��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetDefaultVal
    
    FncNew = True  
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    FncSave = False                                                         
    
    If frm1.vspdData.maxrows < 1 then exit function    

	'-----------------------
    'Precheck area
    '-----------------------
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                           
        Exit Function
    End If
    
    '----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then Exit Function

   	Call SetToolbar("10000000000111")
   	
    If CompareDateByFormat(frm1.txtFromInsrtDt.text,frm1.txtToInsrtDt.text,frm1.txtFromInsrtDt.Alt,frm1.txtToInsrtDt.Alt, _
        	               "970024",frm1.txtFromInsrtDt.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtFromInsrtDt.focus
	   Exit Function
	End If
	
    '-----------------------
    'Save function call area
    '-----------------------
	If DbSave = False then Exit Function
	  
    FncSave = True                                                       
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    With frm1.vspdData
		If .maxrows < 1 then exit function
		.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor .ActiveRow, .ActiveRow
		.Row = .ActiveRow
		.Col = C_FILE_ABBR
		.Text = ""
	
		.ReDraw = True
	End With
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 

    Dim iDx

   On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCancel = False                                                             '��: Processing is NG

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
     Frm1.vspdData.Row = frm1.vspdData.ActiveRow
     
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	
	'On Error Resume Next
	
	FncInsertRow = False
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         
       Exit Function
    End If
   
   
   
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End If 
	
    With frm1.vspdData	
		.ReDraw = False
		.focus
		ggoSpread.Source = frm1.vspdData
		ggoSpread.InsertRow .ActiveRow, imRow
		SetSpreadColor .ActiveRow, .ActiveRow + imRow - 1
		.ReDraw = True
    End With

    Set gActiveElement = document.ActiveElement
    
    If Err.number = 0 Then FncInsertRow = True
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	Dim lDelRows 
	Dim lTempRows 

	If frm1.vspdData.maxrows < 1 then exit function
	
 '----------  Coding part  ------------------------------------------------------------- 
	ggoSpread.Source = frm1.vspdData	
	lDelRows = ggoSpread.DeleteRow
	lgLngCurRows = lDelRows + lgLngCurRows
	lTempRows = frm1.vspdData.MaxRows - lgLngCurRows
End Function
'========================================================================================
' Function Name : FncDelete
' Function Desc : 
'========================================================================================
Function FncDelete()	
	
	if frm1.vspdData.maxRows >= 1 then
		If DisplayMsgBox("210034", parent.VB_YES_NO, "x", "x") = vbYes Then '�����Ͻðڽ��ϱ�?
		 
		End If   
	end if	
	    
    'MyBizASP.location.href = "B81106MA1_frwriteBiz.asp?txtMode=" & UID_M0003  & "&txtKeyNo=" & MyBizAsp.frTitle.intKeyNo
    //MyBizASPForDelete.location.href = "B81106MA1_frwriteBiz.asp?txtMode=" & UID_M0003  & "&txtKeyNo=" & MyBizAsp.frTitle.intKeyNo	    
	    
End Function
	
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
 	Call parent.FncExport(Parent.C_MULTI)		
 End Function
 
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()

	FncExit = False
	
	Dim IntRetCD
	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then	Exit Function
    End If
    
    FncExit = True    
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
	Dim strVal

	DbQuery = False                                                             

	Call LayerShowHide(1)
	
	With frm1
	//MyBizASP1.location.href = "../../blank.htm"
	If lgIntFlgMode = Parent.OPMD_UMODE Then
	 	
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&txtIns_person=" & .hdnPlantCd.value
	    strVal = strVal & "&txtTitle=" & .hdnitemcd.value
	    strVal = strVal & "&lgPageNo="		 & lgPageNo						'��: Next key tag 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else	
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&txtIns_person=" & .txtIns_person.value
	    strVal = strVal & "&txtTitle=" & .txtTitle.value
	    strVal = strVal & "&txtFromInsrtDt=" & .txtFromInsrtDt.text
	    strVal = strVal & "&txtToInsrtDt=" & .txtToInsrtDt.text
	    strVal = strVal & "&txtFromUseDt=" & .txtFromUseDt.text
        strVal = strVal & "&txtToUseDt=" & .txtToUseDt.text
	    strVal = strVal & "&lgPageNo="	 & lgPageNo						'��: Next key tag 
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If 
    
	End With

	Call RunMyBizASP(MyBizASP, strVal)													'��: �����Ͻ� ASP �� ���� 
	

	DbQuery = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryOk()
	Dim ii
	
    lgIntFlgMode = Parent.OPMD_UMODE							'��: Indicates that current mode is Update mode
	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
	Call SetToolBar("110000000001111")				'��ư ���� ���� 

    If frm1.vspdData.MaxRows > 0 Then
		Call SetToolBar("110010110001111")
		frm1.vspddata.focus
		MyBizASP1.location.href = "B81106RA1.asp?File_no=" & ""
	End If
	Set gActiveElement = document.activeElement
	call vspdData_click (1,1)
	frm1.txtIns_person.focus()
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow
	Dim lGrpCnt     
	Dim strVal,strDel
	Dim ColSep, RowSep
	Dim iTmpCUBuffer         '������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount    '������ ���� Position
	Dim iTmpCUBufferMaxCount '������ ���� Chunk Size
	Dim iTmpDBuffer          '������ ���� [����] 
	Dim iTmpDBufferCount     '������ ���� Position
	Dim iTmpDBufferMaxCount  '������ ���� Chunk Size
	Dim ii
	
	ColSep = parent.gColSep               
	RowSep = parent.gRowSep               

    DbSave = False                                                          '��: Processing is NG
	Call LayerShowHide(1)
	
	frm1.txtMode.value = Parent.UID_M0002
	'-----------------------
	'Data manipulate area
	'-----------------------
	lGrpCnt = 0
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����]
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ֱ� ������ ����[����,�ű�]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '�ֱ� ������ ����[����,�ű�]

	iTmpCUBufferMaxCount = -1 
	iTmpDBufferMaxCount = -1 
	
	With frm1
		.txtMode.value = parent.UID_M0002
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	ggoSpread.source = frm1.vspdData

      For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep		'��: C=Create
		        Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep		'��: U=Update
				Case ggoSpread.DeleteFlag
					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep		'��: U=Delete
						
			End Select			
 
		    Select Case .vspdData.Text 
		        Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag		'��: �ű�, ���� 
	
					strVal = strVal & .txtIns_person.value  & ColSep
					strVal = strVal & .txtTitle.value  & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_TITLE,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_FILE_ABBR,lRow,"X","X")) & ColSep
					lGrpCnt = lGrpCnt + 1

		        Case ggoSpread.DeleteFlag							'��: ���� 
					
					strDel = strDel & Trim(GetSpreadText(.vspdData,1,lRow,"X","X")) & ColSep &  ColSep & RowSep
					
  		            lGrpCnt = lGrpCnt + 1
		    End Select
		 
		Next
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)			'��: �����Ͻ� ASP �� ���� 

	DbSave = True                                                      
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	Call MainQuery()
End Function


'=======================================================================================================
' Function Name : FncWrite
' Function Desc : 
'========================================================================================================
Function FncWrite()

		Dim arrRet,parm
		
		If IsOpenPop = True Then Exit Function
         reDim parm(3)
		IsOpenPop = True	
	
		arrRet = window.showModalDialog ("B81106PA1.asp?strMode=" & parent.UID_M0001,Array(window.parent,parm(0),parm(1)), _
       "dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			//"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	

		If arrRet = True Then
			call dbsaveOK()
			MyBizASP1.location.reload						
		End If
				
		IsOpenPop = False

	End Function

'=======================================================================================================
' Function Name : FncModify
' Function Desc : 
'========================================================================================================
	Function FncModify()

		Dim arrRet,sFile_no
		
		If IsOpenPop = True Then Exit Function
		if frm1.vspdData.maxRows = 0 then
			
			 call DisplayMsgBox("900025", "X", "X", "X") 
			 Exit Function
		end if		 
		sFile_no=GetSpreadText(frm1.vspdData,1,frm1.vspdData.ActiveRow,"X","X")
		IsOpenPop = True
	
		
		arrRet = window.showModalDialog ("B81106PA1.asp?strMode=" & parent.UID_M0002 & "&file_no="&sFile_no,Array(window.parent,sFile_no), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	

		If arrRet = True Then
			call dbsaveOK()
			MyBizASP1.location.reload
		End If
		
		IsOpenPop = False

	End Function
	
	

'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(lRow)
	With frm1.vspdData
		.Row = lRow
		.Col = C_FILE_ABBR
		.Action = 0
		Call SetFocusToDocument("M") 
		.focus
	End With
End Function


Function Jump()	

    Dim iRet
    Dim iRet2
    Dim strVal
    Dim iArr
    Dim eisWindow ,strPgmID
    'On Error Resume Next
   
	CookiePage("")
    PgmJump(BIZ_PGM_REG_ID)
  
End Function


		

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<!--
'########################################################################################################
'#						6. TAG 																		#
'######################################################################################################## 
-->
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%> >
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
							<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
							<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strAspMnuMnunm")%></font></td>
							<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						</TR>
					</TABLE>
					</TD>
				
					<TD WIDTH=* Align=right>
					<A onclick="vbscript:FncWrite()">���</A>&nbsp;|&nbsp;<A onclick="vbscript:FncModify()">����</A>
					<A onclick="vbscript:FncDelete()"></A></TD>
					<TD WIDTH=10>&nbsp;</TD>
					
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=50%>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
					<FIELDSET CLASS="CLSFLD" >
					<TABLE <%=LR_SPACE_TYPE_40%>>
					   <TR>
							<TD CLASS="TD5" NOWRAP>�������</TD>
							<TD CLASS="TD6" NOWRAP >
							<script language =javascript src='./js/b81106ma1_OBJECT2_txtFromInsrtDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/b81106ma1_OBJECT3_txtToInsrtDt.js'></script>
							
							<TD CLASS="TD5" NOWRAP>��ȿ����</TD>
							<TD CLASS="TD6" NOWRAP >
							<script language =javascript src='./js/b81106ma1_O_txtFromUseDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/b81106ma1_OBJECT5_txtToUseDt.js'></script>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>�ۼ���</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="�ۼ���" NAME="txtIns_person" SIZE=32 MAXLENGTH=50  tag="11NXXU" >
							<TD CLASS="TD5" NOWRAP>����</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����" NAME="txtTitle" SIZE=32 MAXLENGTH=50 tag="11NXXU"></TD>
						</TR>
						
					</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH="100%" <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
										<script language =javascript src='./js/b81106ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
		<TD <%=HEIGHT_TYPE_02%> ></TD>
	</TR>
		
	<TR>
		<TD WIDTH=100% HEIGHT=100% VALIGN=TOP>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO  framespacing=0></IFRAME>
			<IFRAME NAME="MyBizASP1" SRC="B81106RA1.asp" WIDTH=100% HEIGHT=90% FRAMEBORDER=0 SCROLLING=NO  framespacing=0></IFRAME>
			
			<IFRAME NAME="MyBizASPForDelete" SRC="../../blank.htm" WIDTH=10% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO  framespacing=0></IFRAME>
		</TD>
		


	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hValidDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hCurrency" tag="24">
</FORM>
</BODY>
</HTML>
 
