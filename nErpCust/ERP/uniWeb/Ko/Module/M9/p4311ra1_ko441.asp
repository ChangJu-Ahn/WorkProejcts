
<%@ LANGUAGE="VBSCRIPT" %>
<%'********************************************************************************************************
'*  1. Module Name			: Production																*
'*  2. Function Name		: Reference Popup Component List											*
'*  3. Program ID			: p4311ra1																	*
'*  4. Program Name			: ��ǰ��������																*
'*  5. Program Desc			: Reference Popup															*
'*  7. Modified date(First)	: 2000/04/06																*
'*  8. Modified date(Last)	: 2002/06/25																*
'*  9. Modifier (First)    	: Kim, Gyoung-Don															*
'* 10. Modifier (Last)		: Ryu Sung Won																*
'* 11. Comment 				:																			*
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)         
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin                  *
'********************************************************************************************************%>

<HTML>
<HEAD>
<!--'####################################################################################################
'#						1. �� �� ��																		#
'#####################################################################################################-->
<!--'********************************************  1.1 Inc ����  ****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'============================================  1.1.1 Style Sheet  ===================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'============================================  1.1.2 ���� Include  ==================================
'=====================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY1_ID	= "p4311rb1_ko441.asp"								'��: Head Query �����Ͻ� ���� ASP�� 
Const BIZ_PGM_QRY2_ID	= "p4311rb2_ko441.asp"								'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
' Grid 1(vspdData1) - Operation
Dim arrReturn					<% '--- Return Parameter Group %>
Dim C_CompntCd
Dim C_CompntNm
Dim C_CompntSpec
Dim C_RqrdQty
Dim C_Unit
Dim C_RqrdDt
Dim C_TrackingNo
Dim C_IssuedQty
Dim C_ConsumedQty
Dim C_TTlIssueQty
Dim C_MajorSLCd
Dim C_MajorSLNm
Dim C_OprNo
Dim C_WcCD
Dim C_ReqSeqNo
Dim C_ReqNo
Dim C_ResrvStatus
Dim C_ResrvStatusDesc
Dim C_IssueMeth
Dim C_IssueMethDesc

' Grid 2(vspdData2) - Operation
Dim C_BlockIndicator
Dim C_SLCd
Dim C_SLNm
Dim C_AllTrackingNo
Dim C_LotNo
Dim C_LotSubNo
Dim C_OnHandQty
Dim C_PrevOnHandQty
Dim C_StkOnInspQty
Dim C_StkOnTrnsQty
Dim C_IssueQty

'==========================================  1.2.2 Global ���� ����  ==================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'======================================================================================================
<!-- #Include file="../../inc/lgVariables.inc" -->
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4
Dim lgStrPrevKey5

Dim lgPlantCD
Dim lgProdOrdNo
Dim lgOprNo

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
Dim IsOpenPop			'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim lgOldRow

'*********************************************  1.3 �� �� �� ��  ****************************************
'*	����: Constant�� �ݵ�� �빮�� ǥ��.																*
'********************************************************************************************************
Dim arrParent
Dim arrParam					
		
'------ Set Parameters from Parent ASP ------
arrParent	= window.dialogArguments
Set PopupParent = arrParent(0)
lgPlantCD	= arrParent(1)
lgProdOrdNo = arrParent(2)
lgOprNo		= arrParent(3)

top.document.title = PopupParent.gActivePRAspName

'########################################################################################################
'#						2. Function ��																	#
'#																										#
'#	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� ���					#
'#	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.							#
'#						 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����)			#
'########################################################################################################

'*******************************************  2.1 ���� �ʱ�ȭ �Լ�  *************************************
'*	���: �����ʱ�ȭ																					*
'*	Description : Global���� ó��, �����ʱ�ȭ ���� �۾��� �Ѵ�.											*
'********************************************************************************************************
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(pvSpdNo)
		Case "A"
			C_CompntCd			= 1	
			C_CompntNm			= 2
			C_CompntSpec		= 3
			C_RqrdQty			= 4
			C_Unit				= 5
			C_RqrdDt			= 6
			C_TrackingNo		= 7
			C_IssuedQty			= 8
			C_ConsumedQty		= 9
			C_TTlIssueQty		= 10
			C_MajorSLCd			= 11
			C_MajorSLNm			= 12
			C_OprNo				= 13
			C_WcCD				= 14
			C_ReqSeqNo			= 15
			C_ReqNo				= 16
			C_ResrvStatus		= 17
			C_ResrvStatusDesc	= 18
			C_IssueMeth			= 19
			C_IssueMethDesc		= 20
		Case "B"
			C_BlockIndicator	= 1
			C_SLCd				= 2
			C_SLNm				= 3
			C_AllTrackingNo		= 4
			C_LotNo				= 5
			C_LotSubNo			= 6
			C_OnHandQty			= 7
			C_PrevOnHandQty		= 8
			C_StkOnInspQty		= 9
			C_StkOnTrnsQty		= 10
			C_IssueQty			= 11
	End Select			
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	lgStrPrevKey = ""
	lgStrPrevKey2 = ""
	lgStrPrevKey3 = ""
	lgStrPrevKey4 = ""
	Self.Returnvalue = Array("")
End Function

'*******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "RA")%>
End Sub

'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : ȭ�� �ʱ�ȭ(���� Field�� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)		=
'========================================================================================================

Sub SetDefaultVal()
	txtProdOrdNo.value = lgProdOrdNo
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	Select Case UCase(pvSpdNo)
		Case "A"
			'------------------------------------------
			' Grid 1 - Operation Spread Setting
			'------------------------------------------
			Call InitSpreadPosVariables(pvSpdNo)
			ggoSpread.Source = vspdData1
			ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread

			With vspdData1 
			.ReDraw = false
			.MaxCols = C_IssueMethDesc +1											'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0

			Call GetSpreadColumnPos(pvSpdNo)
	
			ggoSpread.SSSetEdit		C_CompntCd,		"��ǰ", 18
			ggoSpread.SSSetEdit		C_CompntNm,		"��ǰ��", 25
			ggoSpread.SSSetEdit		C_CompntSpec,	"�԰�", 25
			ggoSpread.SSSetFloat	C_RqrdQty, 		"�ʿ����", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_Unit, 		"����", 6
			ggoSpread.SSSetDate 	C_RqrdDt, 		"�ʿ���", 11, 2, PopupParent.gDateFormat
			ggoSpread.SSSetEdit 	C_TrackingNo,	"Tracking No.", 25
			ggoSpread.SSSetFloat	C_IssuedQty,	"������", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ConsumedQty,	"�Һ����", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_TTlIssueQty,	"������", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"	
			ggoSpread.SSSetEdit		C_MajorSLCd,	"���â��", 10
			ggoSpread.SSSetEdit		C_MajorSLNm,	"���â���", 20
			ggoSpread.SSSetEdit		C_OprNo,		"����", 6
			ggoSpread.SSSetEdit		C_WcCD,			"�۾���", 10
			ggoSpread.SSSetEdit		C_ReqSeqNo,		"����", 6
			ggoSpread.SSSetEdit		C_ReqNo,		"����", 6
			ggoSpread.SSSetCombo	C_ResrvStatus,	"������", 10
			ggoSpread.SSSetCombo	C_ResrvStatusDesc, "������", 10
			ggoSpread.SSSetCombo	C_IssueMeth,	"�����", 15
			ggoSpread.SSSetCombo	C_IssueMethDesc,"�����", 15

			Call ggoSpread.SSSetColHidden(.MaxCols,		.MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_ReqSeqNo,	C_ReqSeqNo, True)
			Call ggoSpread.SSSetColHidden(C_ReqNo,		C_ReqNo, True)
			Call ggoSpread.SSSetColHidden(C_TTlIssueQty, C_TTlIssueQty, True)
'			Call ggoSpread.SSSetColHidden(C_IssuedQty,	C_IssuedQty, True)
			Call ggoSpread.SSSetColHidden(C_ResrvStatus, C_ResrvStatus, True)
			Call ggoSpread.SSSetColHidden(C_IssueMeth,	C_IssueMeth, True)
	
			ggoSpread.SSSetSplit2(2)
			.ReDraw = true
			End With

		Case "B"
			'------------------------------------------
			' Grid 2 - Component Spread Setting
			'------------------------------------------
			Call InitSpreadPosVariables("B")
			ggoSpread.Source = vspdData2
			ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread
	
			With vspdData2
			.ReDraw = false		
			.MaxCols = C_IssueQty +1													'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0

			Call GetSpreadColumnPos("B")
	
			ggoSpread.SSSetEdit		C_BlockIndicator,"Block", 6
			ggoSpread.SSSetEdit		C_SLCd,			"â��", 10
			ggoSpread.SSSetEdit		C_SLNm,			"â���", 20
			ggoSpread.SSSetEdit		C_AllTrackingNo,"Tracking No.", 25
			ggoSpread.SSSetEdit		C_LotNo,		"Lot No.", 12
			ggoSpread.SSSetEdit		C_LotSubNo,		"����", 6
			ggoSpread.SSSetFloat	C_OnHandQty,	"��ǰ����", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_PrevOnHandQty,"������ǰ���", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_StkOnInspQty, "�˻��߼�", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_StkOnTrnsQty, "�̵��߼�", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"	
			ggoSpread.SSSetFloat	C_IssueQty,		"������", 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"				
	
			Call ggoSpread.SSSetColHidden(.MaxCols,	.MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_IssueQty,	C_IssueQty, True)
	
			ggoSpread.SSSetSplit2(3)
			.ReDraw = true
			End With
	End Select
	    
    vspdData1.OperationMode = 5 '20080218::hanc

    Call SetSpreadLock()
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    '--------------------------------
    'Grid 1
    '--------------------------------
    ggoSpread.Source = vspdData1
	ggoSpread.SpreadLockWithOddEvenRowColor()
    
    '--------------------------------
    'Grid 2
    '--------------------------------
    ggoSpread.Source = vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
            ggoSpread.Source = vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_CompntCd			= iCurColumnPos(1)	
			C_CompntNm			= iCurColumnPos(2)
			C_CompntSpec		= iCurColumnPos(3)
			C_RqrdQty			= iCurColumnPos(4)
			C_Unit				= iCurColumnPos(5)
			C_RqrdDt			= iCurColumnPos(6)
			C_TrackingNo		= iCurColumnPos(7)
			C_IssuedQty			= iCurColumnPos(8)
			C_ConsumedQty		= iCurColumnPos(9)
			C_TTlIssueQty		= iCurColumnPos(10)
			C_MajorSLCd			= iCurColumnPos(11)
			C_MajorSLNm			= iCurColumnPos(12)
			C_OprNo				= iCurColumnPos(13)
			C_WcCD				= iCurColumnPos(14)
			C_ReqSeqNo			= iCurColumnPos(15)
			C_ReqNo				= iCurColumnPos(16)
			C_ResrvStatus		= iCurColumnPos(17)
			C_ResrvStatusDesc	= iCurColumnPos(18)
			C_IssueMeth			= iCurColumnPos(19)
			C_IssueMethDesc		= iCurColumnPos(20)
		Case "B"
			ggoSpread.Source = vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_BlockIndicator	= iCurColumnPos(1)
			C_SLCd				= iCurColumnPos(2)
			C_SLNm				= iCurColumnPos(3)
			C_AllTrackingNo		= iCurColumnPos(4)
			C_LotNo				= iCurColumnPos(5)
			C_LotSubNo			= iCurColumnPos(6)
			C_OnHandQty			= iCurColumnPos(7)
			C_PrevOnHandQty		= iCurColumnPos(8)
			C_StkOnInspQty		= iCurColumnPos(9)
			C_StkOnTrnsQty		= iCurColumnPos(10)
			C_IssueQty			= iCurColumnPos(11)
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
	Call InitSpreadSheet(gActiveSpdSheet.Id)

    If gActiveSpdSheet.Id = "A" Then
		ggoSpread.Source = vspdData1
		Call InitComboBox()
		Call ggoSpread.ReOrderingSpreadData()
		Call InitData(1,1)
	Else
		ggoSpread.Source = vspdData2
		Call ggoSpread.ReOrderingSpreadData()
	End If
End Sub

'========================== 2.2.6 InitComboBox()  ========================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitComboBox()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1017", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = vspdData1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ResrvStatus
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_ResrvStatusDesc

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1016", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    ggoSpread.Source = vspdData1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_IssueMeth
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_IssueMethDesc
End Sub

'========================== 2.2.7 InitData()  =============================================
'	Name : InitData()
'	Description : Combo Display
'==========================================================================================
Sub InitData(ByVal lngStartRow, ByVal iPos)
	Dim intRow
	Dim intIndex
	Dim intMaxRows

	With vspdData1
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.col = C_ResrvStatus
			intIndex = .value
			.Col = C_ResrvStatusDesc
			.value = intindex
			.Row = intRow
			.col = C_IssueMeth
			intIndex = .value
			.Col = C_IssueMethDesc
			.value = intindex			
		Next	
	End With
End Sub

'=========================================  2.3.2 CancelClick()  ========================================
' Name : CancelClick()
' Description : Return Array to Opener Window for Cancel button click
'========================================================================================================
Function CancelClick()
	Redim arrReturn(1,1)
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = vspdData1
	Call SetPopupMenuItemInf("0000111111")
		
	If vspdData1.MaxRows <= 0 Then Exit Sub

	If Row <= 0 Then
        ggoSpread.Source = vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
	If lgOldRow <> Row Then
		
		vspdData1.Col = 1
		vspdData1.Row = row
		
		lgOldRow = Row
		
		vspdData2.MaxRows = 0
	  	
		lgStrPrevKey3 = ""
		lgStrPrevKey4 = ""
  		lgStrPrevKey5 = ""
  		
		If DbDtlQuery = False Then	
			Exit Sub
		End If	
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SP2C"					'SpreadSheet ������ vspdData�ϰ�� 
	Set gActiveSpdSheet = vspdData2
	Call SetPopupMenuItemInf("0000111111")
	
    If vspdData2.MaxRows <= 0 Then Exit Sub
   	  
	If Row <= 0 Then
        ggoSpread.Source = vspdData2
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
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if vspdData1.MaxRows < NewTop + VisibleRowCnt(vspdData1,NewTop) Then
		If lgStrPrevKey <> "" Or lgStrPrevKey2 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
End Sub

'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if vspdData2.MaxRows < NewTop + VisibleRowCnt(vspdData2,NewTop) Then
		If lgStrPrevKey3 <> "" Or lgStrPrevKey4 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

Sub vspdData1_KeyPress(keyAscii)
	If keyAscii=27 Then
 		Call CancelClick()
		Exit Sub
	End If
End Sub	

Sub vspdData2_KeyPress(keyAscii)
	If keyAscii=27 Then
 		Call CancelClick()
		Exit Sub
	End If
End Sub	

'------------------------------------------  OpenSLCd()  -------------------------------------------------
'	Name : OpenSLCd()
'	Description : Storage Location PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSLCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCode

	If IsOpenPop = True or UCase(txtSlCd.className) = "PROTECTED" Then Exit Function

	strCode = txtSLCd.value

	IsOpenPop = True

	arrParam(0) = "â���˾�"											' �˾� ��Ī 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE ��Ī 
	arrParam(2) = strCode													' Code Condition
	arrParam(3) = ""'Trim(frm1.txtSLNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(lgPlantCD), "''", "S") ' Where Condition
	arrParam(5) = "â��"												' TextBox ��Ī 
   	arrField(0) = "SL_CD"													' Field��(0)
   	arrField(1) = "SL_NM"													' Field��(1)
   	arrHeader(0) = "â��"												' Header��(0)
   	arrHeader(1) = "â���"												' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSLCd(arrRet)
	End If
	
	Call SetFocusToDocument("P")
	txtSLCd.focus
	
End Function

'------------------------------------------  SetSLCd()  --------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetSLCd(byval arrRet)
    txtSLCd.value = arrRet(0)  
	txtSLNm.Value = arrRet(1)
End Function

Sub rdoIssueSlCdFlg1_OnClick()
	txtSLCd.value = ""
	txtSLNm.value = ""
	Call ggoOper.SetReqAttr(txtSlCd,"Q")
End Sub

Sub rdoIssueSlCdFlg2_OnClick()
	Call ggoOper.SetReqAttr(txtSlCd,"D")
End Sub


'########################################################################################################
'#						3. Event ��																		#
'#	���: Event �Լ��� ���� ó��																		#
'#	����: Windowó��, Singleó��, Gridó�� �۾�.														#
'#		  ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.								#
'#		  �� Object������ Grouping�Ѵ�.																	#
'########################################################################################################

'********************************************  3.1 Windowó��  ******************************************
'*	Window�� �߻� �ϴ� ��� Even ó��																	*
'********************************************************************************************************

'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029											'��: Load table , B_numeric_format			
	Call SetDefaultVal
	Call InitVariables											'��: Initializes local global variables
	Call ggoOper.LockField(Document, "N")						'��: This function lock the suitable field
	Call InitSpreadSheet("A")
	Call InitSpreadSheet("B")
    Call InitComboBox()
'20080225::hanc	Call FncQuery()
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

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

'*********************************************  3.2 Tag ó��  *******************************************
'*	Document�� TAG���� �߻� �ϴ� Event ó��																*
'*	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ�							*
'*	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.																	*
'********************************************************************************************************
'==========================================  3.2.1 Search_OnClick =======================================
'========================================================================================================
Function FncQuery
	FncQuery = False
		If vspddata1.MaxRows = 0 Then
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Function
			End If
		Else
			Call SetActiveCell(vspdData1,1,1,"P","X","X")
			Set gActiveElement = document.activeElement
			Call DbPreDtlQuery()
		End If
	FncQuery = False
End Function

'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************




'########################################################################################################
'#					     4. Common Function��															#
'########################################################################################################
'########################################################################################################
'#						5. Interface ��																	#
'########################################################################################################
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()

    Err.Clear												'��: Protect system from crashing
	    
    DbQuery = False											'��: Processing is NG
	    
    Call LayerShowHide(1)
	    
    Dim strVal

    strVal =  BIZ_PGM_QRY1_ID & "?txtMode=" & PopupParent.UID_M0001		'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd=" & lgPlantCD			'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtProdOrdNo=" & txtProdOrdNo.Value 'lgProdOrdNo		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtOprNo=" & lgOprNo				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
	strVal = strVal & "&txtMaxRows=" & vspdData1.MaxRows

    Call RunMyBizASP(MyBizASP, strVal)						'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                          '��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(ByVal LngMaxRow)											'��: ��ȸ ������ ������� 
	Call InitData(LngMaxRow,1)
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    
    Call SetActiveCell(vspdData1,1,1,"P","X","X")
	Set gActiveElement = document.activeElement
    
	vspdData2.MaxRows = 0
	If DbDtlQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If	
	lgOldRow = 1
	vspdData1.Focus
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbPreDtlQuery()											'��: ��ȸ ������ ������� 
	lgStrPrevKey3 = ""
	lgStrPrevKey4 = ""
  	lgStrPrevKey5 = ""
	vspdData2.MaxRows = 0
	
	If DbDtlQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If	
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 

Dim strVal
Dim lngRows
    
	DbDtlQuery = False   
	vspdData1.Row = vspdData1.ActiveRow

	Call LayerShowHide(1)

	strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & PopupParent.UID_M0001
	strVal = strVal & "&txtPlantCd=" & lgPlantCD
	vspdData1.Col = C_CompntCd
	strVal = strVal & "&txtChildItemCd=" & Trim(vspdData1.Text)
	vspdData1.Col = C_TrackingNo
	strVal = strVal & "&txtTrackingNo=" & Trim(vspdData1.Text)
		
	If rdoIssueSlCdFlg1.checked = True Then
		vspdData1.Col = C_MajorSLCd
		strVal = strVal & "&txtMajorSlCd=" & Trim(vspdData1.Text)
	Else
		strVal = strVal & "&txtMajorSlCd=" & Trim(txtSlCd.Value)
	End If
	strVal = strVal & "&txtSlCd=" & Trim(txtSlCd.Value)
	strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
	strVal = strVal & "&lgStrPrevKey4=" & lgStrPrevKey4
	strVal = strVal & "&lgStrPrevKey5=" & lgStrPrevKey5
	strVal = strVal & "&txtMaxRows=" & vspdData2.MaxRows
			
	Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 

    DbDtlQuery = True

End Function


Function DbDtlQueryOk(ByVal LngMaxRow)												'��: ��ȸ ������ ������� 

End Function

Function OKClick()

	Dim intColCnt, intRowCnt, intInsRow, i_RowCnt
	Dim before_supplier, curr_supplier, before_MvmtType, curr_MvmtType, curr_FLAG, before_FLAG

		If vspdData1.SelModeSelCount > 0 Then 

			intInsRow = 0
			i_RowCnt        =   0
			before_supplier =   "" 
			curr_supplier   =   "" 
			before_MvmtType =   "" 
			curr_MvmtType   =   ""
			before_FLAG =   "" 
			curr_FLAG   =   "" 

			Redim arrReturn(vspdData1.SelModeSelCount-1, vspdData1.MaxCols - 2)


			For intRowCnt = 1 To vspdData1.MaxRows
				vspdData1.Row = intRowCnt

				If vspdData1.SelModeSelected Then
				i_RowCnt    =   i_RowCnt  + 1
					For intColCnt = 0 To vspdData1.MaxCols - 2
                        
                    	Select Case intColCnt
                    		Case 0                                        
                                vspdData1.Col = C_CompntCd             
                    		Case 1                                        
                                vspdData1.Col = C_CompntNm
                    		Case 3                           
                                vspdData1.Col = C_RqrdQty
                    		Case 4
                                vspdData1.Col = C_Unit
                    		Case 7
                                vspdData1.Col = C_IssuedQty
                    		Case 14
                                vspdData1.Col = C_ReqSeqNo
                    	End Select
                        
						arrReturn(intInsRow, intColCnt) = vspdData1.Text
					Next
						arrReturn(intInsRow, 2) = txtProdOrdNo.Value
					intInsRow = intInsRow + 1
				End IF								
                before_supplier     =   curr_supplier
                before_MvmtType     =   curr_MvmtType
                before_FLAG         =   curr_FLAG
                
			Next
			
		End if			
		Self.Returnvalue = arrReturn
		Self.Close()
End Function	

'20080218::hanc
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
	'If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'20080225::hanc
Function OpenProdOrderNo()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(8)
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If IsOpenPop = True Or UCase(txtProdOrdNo.className) = "PROTECTED" Then Exit Function

	If lgPlantCD = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		IsOpenPop = False
		Exit Function
	End If
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = lgPlantCD
	arrParam(1) = "" 'frm1.txtProdFromDt.Text
	arrParam(2) = "" 'frm1.txtProdToDt.Text
	arrParam(3) = "" 'frm1.cboOrderStatus.value
	arrParam(4) = "" 'frm1.cboOrderStatus.value
	arrParam(5) = "" 'Trim(frm1.txtProdOrdNo.value) 
	arrParam(6) = "" 'Trim(frm1.txtTrackingNo.value)
	arrParam(7) = "" 'Trim(frm1.txtItemCd.value)
	arrParam(8) = "" 'Trim(frm1.cboOrderType.value)
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.PopupParent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	txtProdOrdNo.Focus
	
End Function
'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'20080225::hanc---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
    txtProdOrdNo.Value    = arrRet(0)		
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. Tag ��																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=10>
			<FIELDSET CLASS="CLSFLD">
				<TABLE WIDTH=100% CELLSPACING=0>
					<TR>
						<TD CLASS=TD5 NOWRAP>���â��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoIssueSlCdFlg" ID="rdoIssueSlCdFlg1" CLASS="RADIO" tag="11" Value="Y" CHECKED><LABEL FOR="rdoIssueSlCdFlg1">��</LABEL>
						     				 <INPUT TYPE="RADIO" NAME="rdoIssueSlCdFlg" ID="rdoIssueSlCdFlg2" CLASS="RADIO" tag="11" Value="N"><LABEL FOR="rdoIssueSlCdFlg2">�ƴϿ�</LABEL></TD>
						<TD CLASS=TD5 NOWRAP>â��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=10 MAXLENGTH=7 tag="14xxxU" ALT="â��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSLCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSLCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=20 tag="14" ALT="â���"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=10>
			<FIELDSET CLASS="CLSFLD">
				<TABLE WIDTH=100% CELLSPACING=0>	
					<TR>
						<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
    					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrdNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
						<TD CLASS=TD5 NOWRAP>ǰ��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="24xxxU" ALT="ǰ��">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR HEIGHT="50%">
		<TD WIDTH="100%" colspan=4>
			<script language =javascript src='./js/p4311ra1_ko441_A_vspdData1.js'></script>
		</TD>
	</TR>
	<TR HEIGHT="50%">
		<TD WIDTH="100%" colspan=4>
			<script language =javascript src='./js/p4311ra1_ko441_B_vspdData2.js'></script>
		</TD>
	</TR>
	<TR><TD HEIGHT=30>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="DbQuery()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG>&nbsp;
						                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
				<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
