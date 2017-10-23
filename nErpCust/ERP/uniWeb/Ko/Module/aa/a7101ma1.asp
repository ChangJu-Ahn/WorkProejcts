<%@ LANGUAGE="VBSCRIPT"%>
<%Response.Expires = -1%>
<!--
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7101ma1
'*  4. Program Name         : �����ڻ� ��泻����� 
'*  5. Program Desc         : �����ڻ꺰 ���������� ���,����,����,��ȸ 
'*  6. Comproxy List        : +As0011ManageSvr
'*                            +As0018ListSvr
'*                            +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/03/13
'*  8. Modified date(Last)  : 2000/09/08
'*  9. Modifier (First)     : ���ͼ� 
'* 10. Modifier (Last)      : hersheys
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

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->


<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->


<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit         '��: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "a7101mb1.asp"   '��: �����Ͻ� ���� ASP�� 

'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================
Dim C_AcctCd
Dim C_AcctCdPopup
Dim C_AcctNm
Dim C_DeprMthd
Dim C_DeprMthdNm
Dim C_DurYrs 
Dim C_AcctFg 
Dim C_AcctFgNm
Dim C_DeprFg 
Dim C_DeprFgNm

Const C_SHEETMAXROWS = 30             ' : �� ȭ�鿡 �������� �ִ밹��*1.5

On Error Resume Next
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
Dim lgRetFlag
Dim IsOpenPop        


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize column position value in spreadsheet
'========================================================================================================
Sub initSpreadPosVariables()
	C_AcctCd      = 1
	C_AcctCdPopup = 2
	C_AcctNm      = 3
	C_DeprMthd    = 4
	C_DeprMthdNm  = 5
	C_DurYrs      = 6
	C_AcctFg      = 7
	C_AcctFgNm    = 8
	C_DeprFg      = 9
	C_DeprFgNm    = 10
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgPageNo     = ""
    
End Sub

'========================================================================================================
' Name : SetDefaultVal()
' Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
 		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
    	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread  

		.MaxCols = C_DeprFgNm +1       '��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols
		.ColHidden = True
		.MaxRows = 0
      

		.ReDraw = false
		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("6","3","0")

		ggoSpread.SSSetEdit  C_AcctCd,     "�����ڵ�",20,,,18,2'1
		ggoSpread.SSSetButton C_AcctCdPopup      '2
		ggoSpread.SSSetEdit  C_AcctNm,     "������",30 '3
		ggoSpread.SSSetCombo C_DeprMthd,   "�󰢹��", 5 '4
		ggoSpread.SSSetCombo C_DeprMthdNm, "�󰢹��", 14 '5
		
		ggoSpread.SSSetFloat C_DurYrs,     "���뿬��", 14, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetCombo C_AcctFg,     "��������", 5
		ggoSpread.SSSetCombo C_AcctFgNm,   "��������", 20   '15
		ggoSpread.SSSetCombo C_DeprFg,     "�󰢴����������", 5
		ggoSpread.SSSetCombo C_DeprFgNm,   "�󰢴����������", 30  '25

		Call ggoSpread.MakePairsColumn(C_DeprMthd,C_DeprMthdNm,"1")
		Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctCdPopup,"1")
		Call ggoSpread.MakePairsColumn(C_AcctFg,C_AcctFgNm,"1")
		Call ggoSpread.MakePairsColumn(C_DeprFg,C_DeprFgNm,"1")


		Call ggoSpread.SSSetColHidden(C_DeprMthd,C_DeprMthd,True)
		Call ggoSpread.SSSetColHidden(C_AcctFg,C_AcctFg,True)
		Call ggoSpread.SSSetColHidden(C_DeprFg,C_DeprFg,True)

		Call InitComboBox
		.ReDraw = true
    End With

	Call SetSpreadLock
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
	With frm1

		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_AcctCd, -1, C_AcctNm
		ggoSpread.SSSetRequired  C_DeprMthdNm, -1
		ggoSpread.SSSetRequired  C_AcctFgNm,   -1
		ggoSpread.SSSetRequired  C_DeprFgNm,   -1
		ggoSpread.SSSetRequired  C_DurYrs,   -1
		ggoSpread.SpreadLock C_DeprFgNm+1, -1, C_DeprFgNm+1

	.vspdData.ReDraw = True

	End With
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired  C_AcctCd,     pvStarRow, pvEndRow
    ggoSpread.SSSetProtected C_AcctNm,     pvStarRow, pvEndRow

'    ggoSpread.SSSetRequired  C_DeprMthd,   pvStarRow, pvEndRow
    ggoSpread.SSSetRequired  C_DeprMthdNm, pvStarRow, pvEndRow

	ggoSpread.SSSetRequired  C_DurYrs,     pvStarRow, pvEndRow

'	ggoSpread.SSSetRequired  C_AcctFg,     pvStarRow, pvEndRow
	ggoSpread.SSSetRequired  C_AcctFgNm,   pvStarRow, pvEndRow

'	ggoSpread.SSSetRequired  C_DeprFg,     pvStarRow, pvEndRow
	ggoSpread.SSSetRequired  C_DeprFgNm,   pvStarRow, pvEndRow
  
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_AcctCd      = iCurColumnPos(1)
			C_AcctCdPopup = iCurColumnPos(2)
			C_AcctNm      = iCurColumnPos(3)
			C_DeprMthd    = iCurColumnPos(4)
			C_DeprMthdNm  = iCurColumnPos(5)
			C_DurYrs      = iCurColumnPos(6)
			C_AcctFg      = iCurColumnPos(7)
			C_AcctFgNm    = iCurColumnPos(8)
			C_DeprFg      = iCurColumnPos(9)
			C_DeprFgNm    = iCurColumnPos(10)
	End Select
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
' Name : InitComboBox()
' Description : Combo Display
'========================================================================================================= 

Sub InitComboBox()

' ------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim IntRetCD1
	Dim IntRetCD2
	Dim IntRetCD3
	  
	On Error Resume Next

	'IntRetCD1 = CommonQueryRs("MINOR_CD,MINOR_NM", "B_MINOR", "(MAJOR_CD = 'A2002')", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	IntRetCD1 = CommonQueryRs("DEPR_MTHD,DEPR_MTHD_NM", "A_ASSET_DEPR_METHOD", "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	  
	If IntRetCD1 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_DeprMthd
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_DeprMthdNm
	End If

	IntRetCD2 = CommonQueryRs("MINOR_CD,MINOR_NM", "B_MINOR", "(MAJOR_CD = " & FilterVar("A2007", "''", "S") & " )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	  
	If IntRetCD2 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0, Chr(11), vbTab), C_AcctFg
		ggoSpread.SetCombo Replace(lgF1, Chr(11), vbTab), C_AcctFgNm
	End If

	IntRetCD3 = CommonQueryRs("MINOR_CD,MINOR_NM", "B_MINOR", "(MAJOR_CD = " & FilterVar("A2008", "''", "S") & " )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	  
	If intRetCD <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(vbTab & lgF0, Chr(11), vbTab), C_DeprFg
		ggoSpread.SetCombo Replace(vbTab & lgF1, Chr(11), vbTab), C_DeprFgNm
	End If
' ------ Developer Coding part (End )   --------------------------------------------------------------


end sub

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
' ���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'       �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 

'========================================== 2.4.2 Open???()  =============================================
' Name : Open???()
' Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'      ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'=========================================================================================================

'------------------------------------------  OpenAcct()  -------------------------------------------------
' Name : OpenAcct()
' Description : Account PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAcct(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����ڵ��˾�"   ' �˾� ��Ī 

	If iWhere = 0 Then
		arrParam(1) = "a_acct a, a_asset_acct b" ' TABLE ��Ī 
		arrParam(2) = frm1.txtAcctCd.value      ' Code Condition
		arrParam(3) = ""       ' Name Cindition
		arrParam(4) = "a.acct_cd = b.acct_cd and a.acct_type = " & FilterVar("K0", "''", "S") & " " ' Where Condition
		arrParam(5) = "�����ڵ�"    ' �����ʵ��� �� ��Ī 
		 
		arrField(0) = "a.acct_cd"     ' Field��(0)
		arrField(1) = "a.acct_nm"     ' Field��(1)
	Else
		arrParam(1) = "a_acct"      ' TABLE ��Ī 
		arrParam(2) = frm1.vspdData.Text      ' Code Condition
		arrParam(3) = ""       ' Name Cindition
		arrParam(4) = "acct_type = " & FilterVar("K0", "''", "S") & " "   ' Where Condition
		arrParam(5) = "�����ڵ�"    ' �����ʵ��� �� ��Ī 
		 
		arrField(0) = "acct_cd"      ' Field��(0)
		arrField(1) = "acct_nm"      ' Field��(1)
	End If
	    
	arrHeader(0) = "�����ڵ�"    ' Header��(0)
	arrHeader(1) = "������"     ' Header��(1)
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
						"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAcct(arrRet, iWhere)
	End If 
 
End Function


'==========================================  2.4.3 Set???()  =============================================
' Name : Set???()
' Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'------------------------------------------  SetAcct()  --------------------------------------------------
' Name : SetAcct()
' Description : Account Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetAcct(Byval arrRet, Byval iWhere)

	With frm1
		If iWhere = 0 Then 'textbox
			.txtAcctCd.value = arrRet(0)
			.txtAcctNm.value = arrRet(1)
		Else 'spread
			.vspdData.Col  = C_AcctCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col  = C_AcctNm
			.vspdData.Text = arrRet(1)

			lgBlnFlgChgValue = True
		End If
	   
	End With
End Function

Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
'  IntRetCD = DisplayMsgBox(frm1.vspdData.Maxcols , parent.VB_YES_NO, "X", "X")
	

	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    Call InitVariables                                                      '��: Initializes local global variables

    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitComboBox
    'Call MASetToolbar("11001101001011")          '��: ��ư ���� ���� 
    Call SetToolbar("11100100000011")   
    frm1.txtAcctCd.focus 
    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
' Document�� TAG���� �߻� �ϴ� Event ó�� 
' Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
' Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag ó��  *********************************************
' Window�� �߻� �ϴ� ��� Even ó�� 
'*********************************************************************************************************
Sub txtAcctCd_OnChange()
	If Trim(frm1.txtAcctCd.value) = "" Then
		frm1.txtAcctNm.value = ""
	End If
End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================

Sub vspdData_Change(ByVal Col, ByVal Row)
    
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)  

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
 Dim strTemp
 Dim intPos1

 With frm1.vspdData 

 If Row > 0 And Col = C_AcctCdPopUp Then
     .Col = C_AcctCd
     .Row = Row
         
     Call OpenAcct(1)
 End If
     
 End With
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("1101111111")
	gMouseClickStatus = "SPC"
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


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) �÷� width ���� �̺�Ʈ �ڵ鷯 
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				'8) �÷� title ���� 
    Dim iColumnName
 	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
   
End Sub



'==========================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub


'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub


Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'==========================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
		.Row = Row
	Select Case Col
		Case C_DeprMthdNm
			.Col = Col
			intIndex = .Value
			.Col = C_DeprMthd
			.Value = intIndex
			Call subVspdSettingChange(Col, Row)
		Case C_AcctFgNm
			.Col = Col
			intIndex = .Value
			.Col = C_AcctFg
			.Value = intIndex
		Case C_DeprFgNm
			.Col = Col
			intIndex = .Value
			.Col = C_DeprFg
			.Value = intIndex
	End Select
	End With

End Sub

Sub subVspdSettingChange(ByVal Col, ByVal Row)
Dim intIndex
Dim varData
	With frm1.vspdData
	
		.Row = Row

		frm1.vspdData.ReDraw = False
		Select Case Col
			Case  C_DeprMthdNm
				.Col = Col
				intIndex = .Value
				.Col = C_DeprMthd
				.Value = intIndex
				varData = .text
				If Trim(varData) <> "" Then 
					IF CommonQueryRs( " DEPR_PROC_FG " , "A_ASSET_DEPR_METHOD" , " DEPR_MTHD =  " & FilterVar(varData , "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						Select Case UCase(lgF0)
							Case "Y" & Chr(11)			' ������ 
								ggoSpread.SSSetRequired  C_DeprFgNm,		Row,	Row
							Case Else
								ggoSpread.SpreadUnLock  C_DeprFgNm,			Row,	C_DeprFgNm,	Row
						End Select
					Else
					End if
				End if
		End Select
	End With

	frm1.vspdData.ReDraw = True	
	lgBlnFlgChgValue = True
	
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
End Sub


'#########################################################################################################
'            4. Common Function�� 
' ���: Common Function
' ����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


'#########################################################################################################
'            5. Interface�� 
' ���: Interface
' ����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'       Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
' << ���뺯�� ���� �κ� >>
'  ���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'    �����ϵ��� �Ѵ�.
'  1. ������Ʈ���� Call�ϴ� ���� 
'        ADF (ADS, ADC, ADF�� �״�� ���)
'        - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
'  2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'      strRetMsg
'#########################################################################################################

'********************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
' ���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False                                                    '��: Processing is NG
    
    Err.Clear                                                           '��: Protect system from crashing

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
  IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")       '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
     If IntRetCD = vbNo Then
       Exit Function
     End If
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then       '��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")        '��: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call InitVariables
                     '��: Initializes local global variables
 If frm1.txtAcctCd.value = "" Then
  frm1.txtAcctNm.value = ""
 End If
 
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery                '��: Query db data
       
    FncQuery = True                '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    'On Error Resume Next                                                    '��: Protect system from crashing
    
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
  IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
  'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. �ű��۾��� �Ͻðڽ��ϱ�?", vbYesNo)
  If IntRetCD = vbNo Then
   Exit Function
  End If
       
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                         '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                         '��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetDefaultVal
    
    Call SetToolbar("11000100000011")          '��: ��ư ���� ���� 
    
    FncNew = True                                                           '��: Processing is OK

End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    'On Error Resume Next                                                    '��: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
  Call DisplayMsgBox("900002", "X", "X", "X")                                  '��:
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")  '�� �ٲ�κ� 
    If IntRetCD = vbNo Then
        Exit Function
    End If

    If DbDelete = False Then                                                '��: Delete db data
       Exit Function                                                        '��:
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                         '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                         '��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    
    FncDelete = True                                                        '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    'On Error Resume Next                                                    '��: Protect system from crashing
    
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
    If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave                                                      '��: Save db data
    
' frm1.vspdData.ReDraw = False
' ggoSpread.SSDeleteFlag 1 , frm1.vspdData.MaxRows
'   Call SetSpreadLock
' frm1.vspdData.ReDraw = True

 FncSave = True                                                          '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 

 If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData
  .ReDraw = False
 
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.CopyRow
	SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	Call subVspdSettingChange(C_DeprMthdNm,frm1.vspdData.ActiveRow)
    
  'Key field clear
  .Col = C_AcctCd
  .Text = ""
  
  .Col = C_AcctNm
  .Text = ""

  .ReDraw = True
    End With
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel()


 Call SetToolbar("11001111001111")          '��: ��ư ���� ���� 

 If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData 
    ggoSpread.EditUndo             
    
    Call InitData 
                                         '��: Protect system from crashing
 lgBlnFlgChgValue = False
     
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(Byval pvRowCnt) 
	Dim imRow
	FncInsertRow = False
'	imRow = AskSpdSheetAddRowCount()
'	If imRow = "" then
'		Exit Function
'	End If

	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

	If ImRow="" then
		Exit Function
	End If
	End If
	
 With frm1
	.vspdData.focus
	ggoSpread.Source = .vspdData
	'.vspdData.EditMode = True
	.vspdData.ReDraw = False
	ggoSpread.InsertRow ,imRow
	SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	.vspdData.ReDraw = True
 End With
 Call SetToolbar("11001111001111")
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement  
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)            '��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
 Dim IntRetCD
 FncExit = False
    ggoSpread.Source = frm1.vspdData 
    If ggoSpread.SSCheckChange = True Then
  IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
  If IntRetCD = vbNo Then
   Exit Function
  End If
    End If
    FncExit = True
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
' ���� : 
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing

 
 Call InitVariables
 Call LayerShowHide(1)

 Dim strVal
    
    With frm1

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001       '��: 

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = strVal & "&txtAcctCd=" & Trim(.hAcctCd.value) '�Ѱ��� ��� hidden�� �ʿ� ���� 
	Else
		strVal = strVal & "&txtAcctCd=" & Trim(.txtAcctCd.value)   '��: ��ȸ ���� ����Ÿ 
	End If    
	strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

	Call RunMyBizASP(MyBizASP, strVal)          '��: �����Ͻ� ASP �� ���� 

	End With

DbQuery = True
    

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()              '��: ��ȸ ������ ������� 
 Dim iRow
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE            '��: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")         '��: This function lock the suitable field
	Call InitData
	With frm1
		.vspdData.Redraw = False
		For iRow = 0 To frm1.vspdData.MaxRows
			Call subVspdSettingChange(C_DeprMthdNm,iRow)
		Next
		.vspdData.Redraw = True
	End With
    Call SetToolbar("11001111001111")          '��: ��ư ���� ���� 
End Function

Sub InitData()
 Dim intRow
 Dim intIndex 
 
 With frm1.vspdData
  For intRow = 1 To .MaxRows
   
   .Row = intRow
   
   .Col = C_DeprMthd
   intIndex = .value
   .col = C_DeprMthdNm
   .value = intindex
    
   .Col = C_AcctFg
   intIndex = .value
   .col = C_AcctFgNm
   .value = intindex
       
   .Col = C_DeprFg
   intIndex = .value
   .col = C_DeprFgNm
   .value = intindex
           
  Next 
 End With
End Sub


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
    Dim aAs0011     'As New AS0011ManageSvr
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
 'Dim ColSep, RowSep
 
    DbSave = False                                                          '��: Processing is NG
    
    On Error Resume Next                                                   '��: Protect system from crashing

	Call LayerShowHide(1)
 
 With frm1
  .txtMode.value = parent.UID_M0002
  
  '-----------------------
  'Data manipulate area
  '-----------------------
  lGrpCnt = 1
  strVal = ""
  strDel = ""
    
  '-----------------------
  'Data manipulate area
  '-----------------------
  For lRow = 1 To .vspdData.MaxRows
    
      .vspdData.Row = lRow
      .vspdData.Col = 0
      
      Select Case .vspdData.Text

          Case ggoSpread.InsertFlag       '��: �ű� 
     
     strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep '��: C=Create, Row��ġ ���� 

              .vspdData.Col = C_AcctCd
              strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
              
              .vspdData.Col = C_DeprMthd
              strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
              
              .vspdData.Col = C_DurYrs
              strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

              .vspdData.Col = C_AcctFg
              strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

              .vspdData.Col = C_DeprFg
              strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                            
              lGrpCnt = lGrpCnt + 1
              
          Case ggoSpread.UpdateFlag       '��: ���� 

     strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep '��: U=Update

              .vspdData.Col = C_AcctCd
              strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

              .vspdData.Col = C_DeprMthd
              strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
              
              .vspdData.Col = C_DurYrs
              strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

              .vspdData.Col = C_AcctFg
              strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

              .vspdData.Col = C_DeprFg
              strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
              
              lGrpCnt = lGrpCnt + 1
              
          Case ggoSpread.DeleteFlag       '��: ���� 

     strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep'��: D=Delete

              .vspdData.Col = C_AcctCd
              strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
              
              lGrpCnt = lGrpCnt + 1
      End Select

  Next
  
  .txtMaxRows.value = lGrpCnt-1
  .txtSpread.value = strDel & strVal
  'msgbox GetUserPath 
  'msgbox BIZ_PGM_ID
  Call ExecMyBizASP(frm1, BIZ_PGM_ID)          '��: �����Ͻ� ASP �� ���� 

 End With
 
    DbSave = True                                                           '��: Processing is NG

End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()             '��: ���� ������ ���� ���� 
 
 Call ggoOper.ClearField(Document, "1")                                         '��: Clear Condition Field
   
 Call InitVariables
    Call ggoOper.ClearField(Document, "2")        '��: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    Call InitComboBox
 'lgBlnFlgChgValue = False
 
 Call DBQuery()
 'Call MainQuery()
 
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>
<!--
'#########################################################################################################
'            6. Tag�� 
'######################################################################################################### 
 -->
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
 <TR>
  <TD  <%=HEIGHT_TYPE_00%>>&nbsp;</TD>
 </TR>
 <TR HEIGHT=23>
  <TD WIDTH=100%>
   <TABLE <%=LR_SPACE_TYPE_10%>>
    <TR>
     <TD WIDTH=10>&nbsp;</TD>
     <TD CLASS="CLSLTABP">
      <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
       <TR>
        <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����ڻ�����������</font></td>
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
         <TD CLASS="TD5" NOWRAP>�����ڵ�</TD>
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenAcct(0)">
         <INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=30 tag="14"></TD>
         <TD CLASS="TD6"></TD>
         <TD CLASS="TD6"></TD>
        </TR>
       </TABLE>
      </FIELDSET>
     </TD>
    </TR>
    <TR>
     <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
    </TR>    
    <TR>
     <TD WIDTH=100% valign=top>
      <TABLE <%=LR_SPACE_TYPE_20%>>
       <TR>
        <TD WIDTH="100%" NOWRAP>
         <script language =javascript src='./js/a7101ma1_I581708055_vspdData.js'></script>
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
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
  </TD>
 </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hAcctCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>



