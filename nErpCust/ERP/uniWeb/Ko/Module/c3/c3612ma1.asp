<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
'*  1. Module Name          : Cost
'*  2. Function Name        : 
'*  3. Program ID           : C3612MA1
'*  4. Program Name         : ���������γ�����ȸ 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY1_ID	= "c3612mb1.asp"							'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_QRY2_ID	= "c3612mb2.asp"							'��: �����Ͻ� ���� ASP�� 

'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================

' Grid 1(vspdData1) 
Dim C_ChildPlantCd
Dim C_ChildCostCd
Dim C_ChildCostNm
Dim C_ChildItemCd
Dim C_ChildItemNm
Dim C_ChildItemAcct
Dim C_ChildItemAcctNm
Dim C_ChildProcurTypeNm
Dim C_IssueQty
Dim C_IssueAmt

' Grid 2(vspdData2) 
Dim C_PrntPlantCd
Dim C_PrntCostCd
Dim C_PrntCostNm
Dim C_PrntItemCd
Dim C_PrntItemNm
Dim C_PrntItemAcct
Dim C_PrntItemAcctNm
Dim C_DstbFctrNm
Dim C_AllocMthd
Dim C_BasisData
Dim C_DstbAmt

dim	strYYYYMM  

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgBlnFlgChgValue							'Variable is for Dirty flag
Dim lgIntGrpCount								'Group View Size�� ������ ���� 
Dim lgIntFlgMode								'Variable is for Operation Status

Dim lgStrPrevKey1
Dim lgLngCurRows

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgSortKey1
Dim lgSortKey2

Dim strDate
Dim iDBSYSDate
Dim lgCloseFlgMode

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""							'initializes Previous Key 
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
    frm1.hSpid.value = ""
    lgCloseFlgMode	 = "O"			'C : Spid ���� O : �ʱ�ȭ 
    
    frm1.txtSum1.text	= 0
    frm1.txtSum2.text	= 0
    frm1.txtSum3.text	= 0
    frm1.txtSum4.text	= 0.0
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	
	frm1.txtYYYYMM.text	= UniConvDateAToB(LocSvrDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","P","NOCOOKIE","MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)


	Call InitSpreadPosVariables(pvSpdNo)

	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
			
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021224", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_IssueAmt + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("A")


			ggoSpread.SSSetEdit 	C_ChildPlantCd,		"����"	,10
			ggoSpread.SSSetEdit 	C_ChildItemCd,		"ǰ��"	,18
			ggoSpread.SSSetEdit 	C_ChildItemNm,		"ǰ���",30
			ggoSpread.SSSetEdit 	C_ChildCostCd,      "Cost Center"	,15
			ggoSpread.SSSetEdit 	C_ChildCostNm,		"Cost Center��"	,15
			ggoSpread.SSSetEdit 	C_ChildItemAcct,	"ǰ�����",8
			ggoSpread.SSSetEdit 	C_ChildItemAcctNm,	"ǰ�������",10
			ggoSpread.SSSetEdit 	C_ChildProcurTypeNm,"���ޱ���",10
			ggoSpread.SSSetFloat 	C_IssueQty,			"���Լ���"	,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat 	C_IssueAmt,			"���Աݾ�"	,15,parent.ggAmtofMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			
			
			
			Call ggoSpread.MakePairsColumn(C_ChildCostCd, C_ChildCostNm )
			Call ggoSpread.MakePairsColumn(C_ChildItemCd, C_ChildItemNm )
			Call ggoSpread.MakePairsColumn(C_ChildItemAcct, C_ChildItemAcctNm )
			
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(3)
			
			Call SetSpreadLock("A")
			
			.ReDraw = true    
    
		End With
	
    End If
    
    If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData2 
			
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021225", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_DstbAmt + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit 	C_PrntPlantCd,		"����"		,10
			ggoSpread.SSSetEdit		C_PrntItemCd,		"���ǰ��"			,18
			ggoSpread.SSSetEdit 	C_PrntItemNm,		"���ǰ���"		,30
			ggoSpread.SSSetEdit		C_PrntCostCd,		"Cost Center"			,15
			ggoSpread.SSSetEdit		C_PrntCostNm,		"Cost Center��"		,15
			ggoSpread.SSSetEdit		C_PrntItemAcct,     "ǰ�����"		,8
			ggoSpread.SSSetEdit		C_PrntItemAcctNm,	"ǰ�������"		,10
			ggoSpread.SSSetEdit		C_DstbFctrNm,		"��ο��"		,15
			ggoSpread.SSSetEdit 	C_AllocMthd,		"��α���"		,15
			ggoSpread.SSSetFloat	C_BasisData,		"��α��� Data"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_DstbAmt,			"��αݾ�"		,15,parent.ggAmtOfMoneyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"

			
			Call ggoSpread.MakePairsColumn(C_PrntCostCd, C_PrntCostNm )
			Call ggoSpread.MakePairsColumn(C_PrntItemCd, C_PrntItemNm )
			Call ggoSpread.MakePairsColumn(C_PrntItemAcct, C_PrntItemAcctNm )
			
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(3)
			
			Call SetSpreadLock("B")
			
			.ReDraw = true    
    
		End With
    End If
    
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
	With frm1
		If pvSpdNo = "A" Then
		   ggoSpread.Source = frm1.vspdData1
		   ggoSpread.SpreadLockWithOddEvenRowColor()
		End If
		
		If pvSpdNo = "B" Then 
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If	
		   
    End With
End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub



'============================  2.2.7 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)


	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData1) - Order Header
		C_ChildPlantCd				= 1
		C_ChildItemCd				= 2
		C_ChildItemNm				= 3
		C_ChildCostCd				= 4
		C_ChildCostNm				= 5
		C_ChildItemAcct				= 6
		C_ChildItemAcctNm			= 7
		C_ChildProcurTypeNm			= 8
		C_IssueQty					= 9
		C_IssueAmt					= 10
	End If	

	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2) - Result
		C_PrntPlantCd			= 1
		C_PrntItemCd			= 2
		C_PrntItemNm			= 3
		C_PrntCostCd			= 4
		C_PrntCostNm			= 5
		C_PrntItemAcct			= 6
		C_PrntItemAcctNm		= 7
		C_DstbFctrNm			= 8
		C_AllocMthd				= 9
		C_BasisData				= 10
		C_DstbAmt				= 11
	End If	

End Sub

'============================  2.2.8 GetSpreadColumnPos()  ==============================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
      
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
		
 			ggoSpread.Source = frm1.vspdData1
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ChildPlantCd			= iCurColumnPos(1)
			C_ChildItemCd			= iCurColumnPos(2)
			C_ChildItemNm			= iCurColumnPos(3)
			C_ChildCostCd			= iCurColumnPos(4)
			C_ChildCostNm			= iCurColumnPos(5)
			C_ChildItemAcct			= iCurColumnPos(6)
			C_ChildItemAcctNm		= iCurColumnPos(7)
			C_ChildProcurTypeNm		= iCurColumnPos(8)
			C_IssueQty				= iCurColumnPos(9)
			C_IssueAmt				= iCurColumnPos(10)
	
		Case "B"
		
			ggoSpread.Source = frm1.vspdData2
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			C_PrntPlantCd			= iCurColumnPos(1)
			C_PrntItemCd			= iCurColumnPos(2)
			C_PrntItemNm			= iCurColumnPos(3)
			C_PrntCostCd			= iCurColumnPos(4)
			C_PrntCostNm			= iCurColumnPos(5)
			C_PrntItemAcct			= iCurColumnPos(6)
			C_PrntItemAcctNm		= iCurColumnPos(7)
			C_DstbFctrNm			= iCurColumnPos(8)
			C_AllocMthd				= iCurColumnPos(9)
			C_BasisData				= iCurColumnPos(10)
			C_DstbAmt				= iCurColumnPos(11)

    End Select

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
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenPopup(ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strYYYYMM1,strYear,strMonth,strDay
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strYYYYMM1 = strYear & strMonth
	
	select case iWhere
		case 1
			arrParam(0) = "ǰ������˾�"	
			arrParam(1) = "B_MINOR"				
			arrParam(2) = Trim(frm1.txtItemAcctCd.Value)
			arrParam(3) = ""
			arrParam(4) = "major_cd = " & FilterVar("P1001", "''", "S") & " "			
			arrParam(5) = "ǰ�����"			
	
			arrField(0) = "minor_cd"	
			arrField(1) = "minor_nm"	
    
			arrHeader(0) = "ǰ�����"		
			arrHeader(1) = "ǰ�������"		
		case 2
			arrParam(0) = "�����˾�"	
			arrParam(1) = "B_PLANT"				
			arrParam(2) = Trim(frm1.txtPlantCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""			
			arrParam(5) = "����"			
	
			arrField(0) = "plant_cd"	
			arrField(1) = "plant_nm"	
    
			arrHeader(0) = "����"		
			arrHeader(1) = "�����"	
		case 3
			arrParam(0) = "Cost Center�˾�"	
			arrParam(1) = "B_COST_CENTER"				
			arrParam(2) = Trim(frm1.txtCostCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""			
			arrParam(5) = "Cost Center"			
	
			arrField(0) = "cost_cd"	
			arrField(1) = "cost_nm"	
    
			arrHeader(0) = "Cost Center"		
			arrHeader(1) = "Cost Center��"							
	End select 
    
	arrRet = window.showModalDialog("../../comasp/AdoCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	  select case iWhere
		case 1
			frm1.txtItemAcctCd.focus
		case 2
			frm1.txtPlantCd.focus
		case 3
			frm1.txtCostCd.focus
	  End Select	
		Exit Function
	Else
		Call SetPopup(iWhere,arrRet)
	End If	
	
End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPlant()  -------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetPopup(byval iWhere,byval arrRet)
	select case iWhere
		case 1
			frm1.txtItemAcctCd.Value    = arrRet(0)		
			frm1.txtItemAcctNm.Value    = arrRet(1)
			frm1.txtItemAcctCd.focus()
		case 2
			frm1.txtPlantCd.Value    = arrRet(0)		
			frm1.txtPlantNm.Value    = arrRet(1)
			frm1.txtPlantCd.focus()
		case 3
			frm1.txtCostCd.Value    = arrRet(0)		
			frm1.txtCostNm.Value    = arrRet(1)
			frm1.txtCostCd.focus()
			
	End select		
End Function


'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitSpreadSheet("*")                                               '��: Setup the Spread sheet
   
       '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 
 
    frm1.txtYyyymm.focus
	Set gActiveElement = document.activeElement
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    IF frm1.hSpid.value <> "" Then
		lgCloseFlgMode = "C"
		call DBQuery()
	END IF	
End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'=======================================================================================================
'   Event Name : txtYyyymm_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtYyyymm_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtYyyymm.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
	End If 
End Sub


'=======================================================================================================
'   Event Name : txtYyyymm_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtYyyymm_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub



'==========================================================================================
'   Event Name : vspdData1_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then							'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then									'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey1 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey1 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey1
            lgSortKey1 = 1
        End If
   
    End If
    
    lgOldRow = frm1.vspdData1.ActiveRow
			
	frm1.vspdData2.MaxRows = 0 
					
	If DbDtlQuery(Row) = False Then	
		Call RestoreToolBar()
		Exit Sub
	End If
    
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
	
	Set gActiveSpdSheet = frm1.vspdData2

    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey2 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey2 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey2
            lgSortKey2 = 1
        End If
    Else
        
    End If
    
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

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If

End Sub

'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'==========================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'==========================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row )

End Sub

'==========================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'==========================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row )

End Sub


'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	gMouseClickStatus = "SPC"	'Split �����ڵ�    
	
    If Row <> NewRow And NewRow > 0 Then
		lgOldRow = frm1.vspdData1.ActiveRow
				
		frm1.vspdData2.MaxRows = 0
		If DbDtlQuery(NewRow) = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If	    
	    
	End If    
	    

End Sub

'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False														'��: Processing is NG
    Err.Clear																'��: Protect system from crashing


    IF frm1.hSpid.value <> "" Then
		lgCloseFlgMode = "C"
	
		If DbQuery = False Then Exit Function		
    END IF
    
    
	If frm1.txtItemAcctCd.value = "" Then
		frm1.txtItemAcctNm.value = "" 
	End If	
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call InitVariables														'��: Initializes local global variables
	
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then										'��: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------



    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function														'��: Query db data
	End If
	
    FncQuery = True															'��: Processing is OK
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	On Error Resume Next	
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
    On Error Resume Next													'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next													'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)									'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'��: Protect system from crashing
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
	FncExit = True
End Function

'******************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  **************************
'	���� : 
'**************************************************************************************** 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

	Dim strVal
    Dim strYear,strMonth,strDay
    
    DbQuery = False

	Call LayerShowHide(1)
    
    Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strYYYYMM = strYear & strMonth

	With frm1
		If lgIntFlgMode <> parent.OPMD_UMODE and lgCloseFlgMode <> "C" Then
			strVal = BIZ_PGM_QRY1_ID & "?txtMode="	& parent.UID_M0001						'��: 
			strVal = strVal & "&txtYyyymm="			& strYYYYMM			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemAcctCd="		& UCase(Trim(.txtItemAcctCd.value))		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.txtPlantCd.value))		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtCostCd="			& UCase(Trim(.txtCostCd.value))		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtMaxRows="		& .vspdData1.MaxRows
			strVal = strVal & "&txtSpid="			& Trim(.hspid.value)
		Else
			strVal = BIZ_PGM_QRY1_ID & "?txtMode="	& parent.UID_M0003						'��: 
			strVal = strVal & "&txtMaxRows="		& .vspdData1.MaxRows
			strVal = strVal & "&txtSpid="			& Trim(.hspid.value)
		End If
	End With

    Call RunMyBizASP(MyBizASP, strVal)														'��: �����Ͻ� ASP �� ���� 
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
	Call SetToolBar("11000000000111")														'��: ��ư ���� ���� 
	lgIntFlgMode = parent.OPMD_UMODE														'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True
	
	If DbDtlQuery(frm1.vspdData1.ActiveRow) = False Then	
		Call RestoreToolBar()
		Exit Function
	End If
	
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery(ByVal NewRow) 
    
    Dim strVal
       
    DbDtlQuery = False

	Call LayerShowHide(1)

    
	With frm1
		strVal = BIZ_PGM_QRY2_ID & "?txtMode="	& parent.UID_M0001						'��: 

		.vspdData1.Row = NewRow
		.vspdData1.Col = C_ChildPlantCd
		strVal = strVal & "&txtChildPlantCd="	& UCase(Trim(.vspdData1.text))			'��: ��ȸ ���� ����Ÿ 
	
		.vspdData1.Col = C_ChildCostCd	
		strVal = strVal & "&txtChildCostCd="	& UCase(Trim(.vspdData1.text))			'��: ��ȸ ���� ����Ÿ 
		
		.vspdData1.Col = C_ChildItemCd	
		strVal = strVal & "&txtChildItemCd="	& UCase(Trim(.vspdData1.text))			'��: ��ȸ ���� ����Ÿ 

		
		strVal = strVal & "&txtSpid="			& UCase(Trim(.hSpid.value))		

		strVal = strVal & "&txtCondItemAcctCd="	& UCase(Trim(.hItemAcctCd.value))			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtCondPlantCd="	& UCase(Trim(.hPlantCd.value))			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtCondCostCd="	& UCase(Trim(.hCostCd.value))			'��: ��ȸ ���� ����Ÿ 
	End With
    
    Call RunMyBizASP(MyBizASP, strVal)													'��: �����Ͻ� ASP �� ���� 
    
    DbDtlQuery = True
    
End Function


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
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
	Call ggoSpread.ReOrderingSpreadData()
	
End Sub 

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!-- '#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���������γ�����ȸ</font></td>
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
									<TD CLASS=TD5 NOWRAP>�۾����</TD> 
									<TD CLASS="TD6">
										<script language =javascript src='./js/c3612ma1_OBJECT1_txtYyyymm.js'></script>
									</TD>
									<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemAcctCd" SIZE=10 MAXLENGTH=18 tag="11xxxU" ALT="ǰ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(1)">&nbsp;<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=20 tag="14"></TD>

								</TR>
			 					<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=18 tag="11xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(2)">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>Cost Center</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=18 tag="11xxxU" ALT="�ڽ�Ʈ����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup(3)">&nbsp;<INPUT TYPE=TEXT NAME="txtCostNm" SIZE=20 tag="14"></TD>

								</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=40% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="50%">
								<TD WIDTH="100%">
									<script language =javascript src='./js/c3612ma1_A_vspdData1.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS=TDT NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
									<TD CLASS=TD5 NOWRAP>�ѹ�δ��ݾ�</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3612ma1_fpDoubleSingle1_txtSum1.js'></script>
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>  
				<TR>
					<TD WIDTH="100%" HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
									<script language =javascript src='./js/c3612ma1_B_vspdData2.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
								<TD CLASS=TD5>�ѹ�αݾ�</TD>
								<TD CLASS=TD5 NOWRAP>
									<script language =javascript src='./js/c3612ma1_fpDoubleSingle2_txtSum2.js'></script>&nbsp;
    							<TD CLASS=TD5>��αݾ��հ�</TD>
								<TD CLASS=TD5 NOWRAP>
									<script language =javascript src='./js/c3612ma1_fpDoubleSingle2_txtSum3.js'></script>&nbsp;
                               </TD>
    							<TD CLASS=TD5>��α���Data�հ�</TD>
								<TD CLASS=TD5 NOWRAP>
									<script language =javascript src='./js/c3612ma1_fpDoubleSingle2_txtSum4.js'></script>&nbsp;
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
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hSpid" tag="24">
<INPUT TYPE=HIDDEN NAME="hYYYYMM" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemAcctCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hCostCd" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

