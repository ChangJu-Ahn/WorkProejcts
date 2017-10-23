<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m4114qa1
'*  4. Program Name         : �������԰�������Ȳ 
'*  5. Program Desc         :  
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/10/20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Sim Hae Young
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
<!-- '#########################################################################################################
'												1. �� �� �� 
'#########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  =======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   =====================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
Const BIZ_PGM_QRY_ID = "m4114qb1.asp"								'��: Head Query �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMP_ID1	= "M4111QA5"
Const BIZ_PGM_JUMP_ID2	= "M5111QA1"
Const BIZ_PGM_JUMP_ID3	= "M5114QA2"

Dim lglngHiddenRows		'Multi���� �������� ���� ����	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.
Dim EndDate, StartDate

Dim C_MV_DT
Dim C_BP_CD
Dim C_BP_NM
Dim C_MVMT_AMT_SUM
Dim C_MVMT_AMT_SUM_POPUP
Dim C_IV_AMT_SUM
Dim C_IV_AMT_SUM_POPUP
Dim C_BALANCE_AMT

   
EndDate = "<%=GetSvrDate%>"


EndDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
EndDate = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gDateFormat)


'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop										'Popup
'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False					'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgPageNo = 0
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub


'==========================================  2.2.1 SetDefaultVal()  ==========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtFromDt.text = EndDate
	frm1.txtToDt.text   = EndDate
	Call SetToolbar("1100000000001111")
End Sub
   
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    
    Call InitSpreadPosVariables()
    
    With frm1
    
	    ggoSpread.Source = .vspdData
	    ggoSpread.Spreadinit "V20051020", , Parent.gAllowDragDropSpread

	 	.vspdData.ReDraw = false
	    .vspdData.MaxCols = C_BALANCE_AMT + 1
	    .vspdData.MaxRows = 0

	    Call GetSpreadColumnPos("A")

	    ggoSpread.SSSetEdit			C_MV_DT, 				"��ȸ���", 10,2
	    ggoSpread.SSSetEdit			C_BP_CD,				"����ó", 15
	    ggoSpread.SSSetEdit			C_BP_NM,				"����ó��", 25
    	SetSpreadFloatLocal			C_MVMT_AMT_SUM, 		"�߻��ݾ�(GR)",15,1,2       '�߰� 
    	ggoSpread.SSSetButton 		C_MVMT_AMT_SUM_POPUP
    	SetSpreadFloatLocal			C_IV_AMT_SUM, 			"�����ݾ�(IR)",15,1,2       '�߰� 
    	ggoSpread.SSSetButton 		C_IV_AMT_SUM_POPUP
    	SetSpreadFloatLocal			C_BALANCE_AMT, 			"�ܾ�",15,1,2        		'�߰� 
    
		Call ggoSpread.MakePairsColumn(C_MVMT_AMT_SUM,C_MVMT_AMT_SUM_POPUP)
		Call ggoSpread.MakePairsColumn(C_IV_AMT_SUM, C_IV_AMT_SUM_POPUP)

	    Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)
		
		.vspdData.ReDraw = true
		
	   Call SetSpreadLock()

    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_MV_DT                 = 1
	C_BP_CD                 = 2
	C_BP_NM                 = 3
	C_MVMT_AMT_SUM          = 4
	C_MVMT_AMT_SUM_POPUP    = 5
	C_IV_AMT_SUM            = 6
	C_IV_AMT_SUM_POPUP      = 7
	C_BALANCE_AMT			= 8
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_MV_DT                 = iCurColumnPos(1)
			C_BP_CD                 = iCurColumnPos(2)
			C_BP_NM                 = iCurColumnPos(3)
			C_MVMT_AMT_SUM          = iCurColumnPos(4)
			C_MVMT_AMT_SUM_POPUP    = iCurColumnPos(5)
			C_IV_AMT_SUM            = iCurColumnPos(6)
			C_IV_AMT_SUM_POPUP      = iCurColumnPos(7)
			C_BALANCE_AMT			= iCurColumnPos(8)
            
 	End Select
End Sub     
            
'==============================================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False
	        
    With ggoSpread
            
		.SpreadLock 	C_MV_DT ,		-1,C_MV_DT , -1
		.SpreadLock 	C_BP_CD , 		-1,C_BP_CD , -1
		.SpreadLock 	C_BP_NM , 		-1,C_BP_NM , -1
		.SpreadLock 	C_MVMT_AMT_SUM, -1,C_MVMT_AMT_SUM , -1
		.SpreadLock 	C_IV_AMT_SUM, 	-1,C_IV_AMT_SUM , -1
		.SpreadLock 	C_BALANCE_AMT, 	-1,C_BALANCE_AMT , -1
    End With
    frm1.vspdData.ReDraw = True
End Sub     

'==============================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
   
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
    
		if Col = C_MVMT_AMT_SUM_POPUP then
			Call OpenMvmtAmtPoup()
		elseif Col = C_IV_AMT_SUM_POPUP then
			Call OpenIvAmtPoup()
		End if
    End With
End Sub

'------------------------------------------  OpenBpCd()  -------------------------------------------------
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"						 
	arrParam(1) = "B_Biz_Partner"					 
	arrParam(2) = Trim(frm1.txtBpCd.Value)		 
'	arrParam(3) = Trim(frm1.txtBpNm.Value)		 
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					 
	arrParam(5) = "����ó"						 
	
    arrField(0) = "BP_CD"							 
    arrField(1) = "BP_NM"						 
    
    arrHeader(0) = "����ó"						 
    arrHeader(1) = "����ó��"					 
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenMvmtAmtPoup()  -------------------------------------------------
'	Name : OpenMvmtAmtPoup()
'	Description : OpenMvmtAmtPoup PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMvmtAmtPoup()
	
	Dim strRet
	Dim arrParam(3)
	Dim iCalledAspName
	Dim IntRetCD
	Dim iCurRow
	
	If IsOpenPop = True Then Exit Function
	
	iCurRow = frm1.vspdData.ActiveRow
	
	IsOpenPop = True

	arrParam(0) = GetSpreadText(frm1.vspdData,C_MV_DT,iCurRow,"X","X")
	arrParam(1) = GetSpreadText(frm1.vspdData,C_BP_CD,iCurRow,"X","X")
	arrParam(2) = GetSpreadText(frm1.vspdData,C_BP_NM,iCurRow,"X","X")
	arrParam(3) = ""

	iCalledAspName = AskPRAspName("m4114pa1")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m4114pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=640px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
				
End Function


'------------------------------------------  OpenIvAmtPoup()  -------------------------------------------------
'	Name : OpenIvAmtPoup()
'	Description : OpenIvAmtPoup PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenIvAmtPoup()
	
	Dim strRet
	Dim arrParam(3)
	Dim iCalledAspName
	Dim IntRetCD
	Dim iCurRow
	
	If IsOpenPop = True Then Exit Function
	
	iCurRow = frm1.vspdData.ActiveRow
	
	IsOpenPop = True

	arrParam(0) = GetSpreadText(frm1.vspdData,C_MV_DT,iCurRow,"X","X")
	arrParam(1) = GetSpreadText(frm1.vspdData,C_BP_CD,iCurRow,"X","X")
	arrParam(2) = GetSpreadText(frm1.vspdData,C_BP_NM,iCurRow,"X","X")
	arrParam(3) = ""
	
	iCalledAspName = AskPRAspName("m4114pa2")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m4114pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=640px; dialogHeight=400px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
				
End Function



'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029                                                         'Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                       'Lock  Suitable  Field

	Call SetDefaultVal
	Call ggoOper.FormatDate(frm1.txtFromDt,Parent.gDateFormat,"2")
	Call ggoOper.FormatDate(frm1.txtToDt,Parent.gDateFormat,"2")

    Call InitSpreadSheet 
	Call InitVariables		'��: Initializes local global variables
	frm1.txtBpCd.focus
	
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")
	
    gMouseClickStatus = "SPC"
    
    Set gActiveSpdSheet = frm1.vspdData
         
    With frm1.vspdData
		If .MaxRows <= 0 Then Exit Sub
		If Row < 1 Then
			ggoSpread.Source = frm1.vspdData
			 
 			If lgSortKey = 1 Then
 				ggoSpread.SSSort Col					'Sort in Ascending
 				lgSortKey = 2
 			Else
 				ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 				lgSortKey = 1
			End If 
		End If
	
    End With
	
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt
    Dim LngLastRow
    Dim LngMaxRow

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    '/* 9�� ������ġ: �ػ󵵿� ������� �������ǵ��� ���� - START */
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '��: ������ üũ 
    '/* 9�� ������ġ: �ػ󵵿� ������� �������ǵ��� ���� - END */
		If lgPageNo <> "" Then			'���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If

			If DbQuery = False Then
				Exit Sub
			End If
		End If

    End If
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

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� MainQuery�Ѵ�.
'=======================================================================================================
Sub txtFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishEndDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� MainQuery�Ѵ�.
'=======================================================================================================
Sub txtToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
   
	If frm1.txtBpCd.value = "" Then
		frm1.txtBpNm.value = "" 
	End If
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then Exit Function										'��: This function check indispensable field

    If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables															'��: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function										'��: Query db data
       
    Set gActiveElement = document.ActiveElement   
    FncQuery = True																'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()                                           '��: Protect system from crashing
    Call parent.FncPrint()
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)												'��: ȭ�� ���� 
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'��: Protect system from crashing
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    Set gActiveElement = document.ActiveElement   
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
   
    Err.Clear							'��: Protect system from crashing

    DbQuery = False                                                         			'��: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function
 
    Dim strVal
    
	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001						'��: 
	strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
	
	If Trim(frm1.txtFromDt.text) <> "" Then
		strVal = strVal & "&txtFromDt=" & Trim(frm1.txtFromDt.Year) & Trim(frm1.txtFromDt.Month)
	Else
		strVal = strVal & "&txtFromDt=" & ""
	End If
	
	If Trim(frm1.txtToDt.text) <> "" Then
		strVal = strVal & "&txtToDt=" & Trim(frm1.txtToDt.Year) & Trim(frm1.txtToDt.Month)
	Else
		strVal = strVal & "&txtToDt=" & ""
	End If

	
	strVal = strVal & "&txtBpCd=" & Trim(frm1.txtBpCd.value)
	strVal = strVal & "&lgPageNo=" & lgPageNo                  		'��: Next key tag
	strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows

    Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 
    
    DbQuery = True                                                          	'��: Processing is NG
    
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryOk(byVal intARow,byVal intTRow)
	DbQueryOk = False

	Dim i
	Dim lRow
	Dim TmpArrPrevKey
	Dim TmpArrHiddenRows

	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
	Call SetToolBar("11000000000111")				'��ư ���� ���� 

	With frm1
		'-----------------------
		'Reset variables area
		'-----------------------
		lRow = .vspdData.MaxRows

		i=0
		If lRow > 0 And intARow > 0 Then
			If intTRow<=0 Then
				ReDim lglngHiddenRows(intARow - 1)			'lRow = .vspdData.MaxRows	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.
			Else
				TmpArrHiddenRows=lglngHiddenRows

				ReDim lglngHiddenRows(intTRow+intARow - 1)			'lRow = .vspdData.MaxRows	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.
				For i = 0 To intTRow-1
					lglngHiddenRows(i) = TmpArrHiddenRows(i)
				Next
			End If

			For i = intTRow To intTRow+intARow-1
				lglngHiddenRows(i) = 0
			Next

		    lgIntFlgMode = Parent.OPMD_UMODE
		End If
	End With
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtBpCd.focus
	End If
	Set gActiveElement = document.activeElement
    DbQueryOk = true
End Function

'=========================  SetSpreadFloatLocal() ==================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	     
   Select Case iFlag
        Case 2                                                              '�ݾ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '�ܰ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              '����������� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0","999"
    End Select
         
End Sub

'--------------------------------------------------------------------
'		Cookie ����Լ� 
'--------------------------------------------------------------------
Function CookiePage(Byval Kubun)

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                           
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
	
	WriteCookie "CookieIoIvFlg" , "Y"
	
	If Trim(frm1.txtFromDt.text) <> "" Then
		WriteCookie "CookieFromDt" , Trim(frm1.txtFromDt.Year) & "-" & Trim(frm1.txtFromDt.Month) & "-" & "01"
	Else
		WriteCookie "CookieFromDt" , ""
	End If
	
	If Trim(frm1.txtToDt.text) <> "" Then
		WriteCookie "CookieToDt" , Trim(frm1.txtToDt.Year) & "-" & Trim(frm1.txtToDt.Month) & "-" & "01"
	Else
		WriteCookie "CookieToDt" , ""
	End If

	If Trim(frm1.txtBpCd.value) <> "" Then
		WriteCookie "CookieBpCd" , Trim(frm1.txtBpCd.value)
	Else
		WriteCookie "CookieBpNm" , Trim(frm1.txtBpCd.value)
	End If

	If Kubun = 1 Then		'�԰�������ȸ 
		Call PgmJump(BIZ_PGM_JUMP_ID1)
	elseIf Kubun = 2 Then	'���Գ���������ȸ 
		Call PgmJump(BIZ_PGM_JUMP_ID2)
	elseIf Kubun = 3 Then	'�̸��԰����ȸ 
		Call PgmJump(BIZ_PGM_JUMP_ID3)
	End IF
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	' �Է� �ʵ��� ��� MaxLength=? �� ��� 
	' CLASS="required" required  : �ش� Element�� Style �� Default Attribute 
		' Normal Field�϶��� ������� ���� 
		' Required Field�϶��� required�� �߰��Ͻʽÿ�.
		' Protected Field�϶��� protected�� �߰��Ͻʽÿ�.
			' Protected Field�ϰ�� ReadOnly �� TabIndex=-1 �� ǥ���� 
	' Select Type�� ��쿡�� className�� ralargeCB�� ���� width="153", rqmiddleCB�� ���� width="90"
	' Text-Transform : uppercase  : ǥ�Ⱑ �빮�ڷ� �� �ؽ�Ʈ 
	' ���� �ʵ��� ��� 3���� Attribute ( DDecPoint DPointer DDataFormat ) �� ��� 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�������԰�������Ȳ</font></td>
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
									<TD CLASS="TD5" NOWRAP>��ȸ�Ⱓ</TD>
									<TD CLASS="TD6">
										<table cellpadding=0 cellspacing=0>
											<tr>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="������" NAME="txtFromDt" CLASS=FPDTYYYYMM CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 tag="11N" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
												<td>~</td>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="������" NAME="txtToDt" CLASS=FPDTYYYYMM CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 tag="11N" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</td>
											</tr>
										</table>
									</TD>
									<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����ó" NAME="txtBpCd"  SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>
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
								<TD HEIGHT=* WIDTH=100%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData ID=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
    <TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD align="Left">&nbsp;</TD>   
					<td WIDTH="*" align="right"><a href="VBSCRIPT:CookiePage(1)">�԰�������ȸ</a> | <a href="VBSCRIPT:CookiePage(2)">���Գ���������ȸ</a> | <a href="VBSCRIPT:CookiePage(3)">�̸����԰����ȸ</a></td>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<Input TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
