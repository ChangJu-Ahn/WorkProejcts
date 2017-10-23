<%@ LANGUAGE="VBSCRIPT" %> 
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b25ma1.asp
'*  4. Program Name         : Copy Item by Plant
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/03/23
'*  8. Modified date(Last)  : 2004/03/19
'*  9. Modifier (First)     : Mr  Kim
'* 10. Modifier (Last)      : Park In Sik
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT> 

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '��: indicates that All variables must be declared in advance
On Error Resume Next
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Const BIZ_PGM_QRY_ID	= "B3B25MB1.asp"												'��: Detail Query �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID	= "B3B25MB2.asp"												'��: Detail Query �����Ͻ� ���� ASP�� 

'==========================================================================================================
'==========================================================================================================

Dim C_Select 
Dim C_Item
Dim C_ItmNm
Dim C_ItmSpec
Dim C_PrcCtrlInd
Dim C_PrcCtrlIndNm
Dim C_UnitPrice
Dim C_IBPValidFromDt
Dim C_IBPValidToDt
Dim C_ClassCd       
Dim C_ClassDesc     
Dim C_CharValue1    
Dim C_CharValueDesc1
Dim C_CharValue2    
Dim C_CharValueDesc2
Dim C_ItmAcc
Dim C_HdnItmAcc
Dim C_Unit
Dim C_Phantom
Dim C_ItmGroupCd
Dim C_ItmGroupNm
Dim C_DefaultFlg
Dim C_ValidFromDt
Dim C_ValidToDt

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgInsrtFlg
Dim lgFlgAllSelected		'When Selected All
Dim lgFlgCancelClicked		'Cancel Button Clicked
Dim lgFlgCopyClicked		'Copy Button Clicked
Dim lgFlgBtnSelectAllClicked 'When btnSelectAll Clicked

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim IsOpenPop					 'Popup
Dim iDBSYSDate
Dim StartDate, EndDate

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	
	C_Select          		= 1
	C_Item            		= 2
	C_ItmNm           		= 3
	C_ItmSpec         		= 4
	C_PrcCtrlInd      		= 5
	C_PrcCtrlIndNm    		= 6
	C_UnitPrice       		= 7
	C_IBPValidFromDt  		= 8	
	C_IBPValidToDt    		= 9
	C_ClassCd         		= 10
	C_ClassDesc       		= 11
	C_CharValue1      		= 12
	C_CharValueDesc1  		= 13
	C_CharValue2      		= 14
	C_CharValueDesc2  		= 15 
	C_ItmAcc          		= 16
	C_HdnItmAcc       		= 17
	C_Unit            		= 18
	C_Phantom				= 19 
	C_ItmGroupCd      		= 20
	C_ItmGroupNm      		= 21
	C_DefaultFlg      		= 22
	C_ValidFromDt     		= 23
	C_ValidToDt       		= 24 

End Sub

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count

	frm1.btnCopy.disabled = True
	frm1.btnSelectAll.disabled = True
	frm1.btnSelectAll.value = "��ü����"
	lgFlgAllSelected = False
	lgFlgCancelClicked = False
	lgFlgCopyClicked = False
	lgFlgBtnSelectAllClicked = False

End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	
    With frm1.vspdData
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030208",,parent.gAllowDragDropSpread   
	
    .MaxCols = C_ValidToDt + 1											'��: �ִ� Columns +1
    .MaxRows = 0
    
	.ReDraw = false

 	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetCheck	C_Select ,		"",					2,,,1
	ggoSpread.SSSetEdit 	C_Item,			"ǰ��",			20,,,18,2
	ggoSpread.SSSetEdit 	C_ItmNm,		"ǰ���",		25,,,40
	ggoSpread.SSSetEdit 	C_ItmSpec,		"�԰�",			25,,,40
	ggoSpread.SSSetCombo 	C_PrcCtrlInd,	"�ܰ�����",		12
	ggoSpread.SSSetCombo 	C_PrcCtrlIndNm, "�ܰ�����",		20
	ggoSpread.SSSetFloat	C_UnitPrice,	"�ܰ�",			15,parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetDate		C_IBPValidFromDt,"��ȿ������",	12, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_IBPValidToDt,	"��ȿ������",	12, 2, parent.gDateFormat
	ggoSpread.SSSetEdit 	C_ClassCd,		"Ŭ����",		20,,,18,2
	ggoSpread.SSSetEdit 	C_ClassDesc,	"Ŭ������",		25,,,40
	ggoSpread.SSSetEdit 	C_CharValue1,	"��簪1",		20,,,18,2
	ggoSpread.SSSetEdit 	C_CharValueDesc1,"��簪��1",	25,,,40
	ggoSpread.SSSetEdit 	C_CharValue2,	"��簪2",		20,,,18,2
	ggoSpread.SSSetEdit 	C_CharValueDesc2,"��簪��2",	25,,,40
	ggoSpread.SSSetCombo 	C_ItmAcc,		"ǰ�����",		12
	ggoSpread.SSSetCombo 	C_HdnItmAcc,	"ǰ�����",		16
	ggoSpread.SSSetEdit 	C_Unit,			"���ش���",		10,,,3,2
	ggoSpread.SSSetEdit 	C_Phantom,		"����",			10,2
	ggoSpread.SSSetEdit 	C_ItmGroupCd,	"ǰ��׷�",		18,,,10,2
	ggoSpread.SSSetEdit 	C_ItmGroupNm,	"ǰ��׷��",	25
	ggoSpread.SSSetEdit 	C_DefaultFlg,	"��ȿ����",		8,2
	ggoSpread.SSSetDate		C_ValidFromDt,	"ǰ�������",	12, 2, parent.gDateFormat
	ggoSpread.SSSetDate		C_ValidToDt,	"ǰ��������",	12, 2, parent.gDateFormat
	
	Call ggoSpread.SSSetColHidden(C_HdnItmAcc,	C_HdnItmAcc,	True)
	Call ggoSpread.SSSetColHidden(C_PrcCtrlInd,	C_PrcCtrlInd,	True)
	
	Call ggoSpread.SSSetColHidden(.MaxCols,		.MaxCols,		True)
	
	ggoSpread.SSSetSplit2(2)										'frozen ����߰� 
	Call SetSpreadLock 
	
	.ReDraw = true
	
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

 	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLock -1, -1
	ggoSpread.spreadUnLock C_Select, -1, C_Select

End Sub

'================================== 2.2.5 SetSpreadLock1() ==================================================
' Function Name : SetSpreadLock1
' Function Desc : This method set color and protect in spread sheet celles When An Specific Row is Selected
'=============================================================================================================
Sub SetSpreadLock1(ByVal Col, ByVal Row)

 	ggoSpread.SpreadLock		C_PrcCtrlIndNm,		Row, C_PrcCtrlIndNm,	Row
	ggoSpread.SpreadLock		C_UnitPrice,		Row, C_UnitPrice,		Row
	ggoSpread.SpreadLock		C_IBPValidFromDt,	Row, C_IBPValidFromDt,	Row
	ggoSpread.SpreadLock		C_IBPValidToDt,		Row, C_IBPValidToDt,	Row
	
	ggoSpread.SSSetProtected	C_PrcCtrlIndNm, 	Row, Row
	ggoSpread.SSSetProtected	C_UnitPrice, 		Row, Row
	ggoSpread.SSSetProtected	C_IBPValidFromDt,	Row, Row
	ggoSpread.SSSetProtected	C_IBPValidToDt,		Row, Row

End Sub

'================================== 2.2.6 SetSpreadUnLock() ==================================================
' Function Name : SetSpreadUnLock
' Function Desc : This method set color and protect in spread sheet celles When A Specific Row is Selected
'=============================================================================================================
Sub SetSpreadUnLock(ByVal Col, ByVal Row)

	ggoSpread.SpreadUnLock		C_PrcCtrlIndNm,		Row, C_PrcCtrlIndNm,	Row
	ggoSpread.SpreadUnLock		C_UnitPrice,		Row, C_UnitPrice,		Row
	ggoSpread.SpreadUnLock		C_IBPValidFromDt,	Row, C_IBPValidFromDt,	Row
	ggoSpread.SpreadUnLock		C_IBPValidToDt,		Row, C_IBPValidToDt,	Row
    
	ggoSpread.SSSetRequired 	C_PrcCtrlIndNm, 	Row, Row
	ggoSpread.SSSetRequired 	C_UnitPrice, 		Row, Row
	ggoSpread.SSSetRequired		C_IBPValidFromDt,	Row, Row
	ggoSpread.SSSetRequired		C_IBPValidToDt,		Row, Row

End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
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
			C_Select         		= iCurColumnPos(1)
			C_Item           		= iCurColumnPos(2)
			C_ItmNm          		= iCurColumnPos(3)
			C_ItmSpec        		= iCurColumnPos(4)
			C_PrcCtrlInd     		= iCurColumnPos(5)
			C_PrcCtrlIndNm   		= iCurColumnPos(6)
			C_UnitPrice      		= iCurColumnPos(7)
			C_IBPValidFromDt 		= iCurColumnPos(8)	
			C_IBPValidToDt   		= iCurColumnPos(9)
			C_ClassCd        		= iCurColumnPos(10)
			C_ClassDesc      		= iCurColumnPos(11)
			C_CharValue1     		= iCurColumnPos(12)
			C_CharValueDesc1 		= iCurColumnPos(13)
			C_CharValue2     		= iCurColumnPos(14)
			C_CharValueDesc2 		= iCurColumnPos(15)
			C_ItmAcc         		= iCurColumnPos(16)
			C_HdnItmAcc      		= iCurColumnPos(17)
			C_Unit           		= iCurColumnPos(18)
			C_Phantom	     		= iCurColumnPos(19)
			C_ItmGroupCd     		= iCurColumnPos(20)
			C_ItmGroupNm     		= iCurColumnPos(21)
			C_DefaultFlg     		= iCurColumnPos(22)
			C_ValidFromDt    		= iCurColumnPos(23)
			C_ValidToDt      		= iCurColumnPos(24)
    End Select    
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    Dim strCboCd 
    Dim strCboNm
	
	'****************************
    ' ǰ����� 
    '****************************     
    strCboCd = ""
    strCboNm = ""

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    	                 
	'Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))
	
    strCboCd = Replace(lgF0,chr(11),vbTab)
    strCboNm = Replace(lgF1,chr(11),vbTab)  
    
	ggoSpread.SetCombo strCboCd,C_HdnItmAcc
    ggoSpread.SetCombo strCboNm,C_ItmAcc
    	
	'****************************
    ' ����,���ձ���,��ȿ���� 
    '****************************    
    strCboCd = ""
    strCboNm = ""
    
    strCboCd = "Y" & vbTab & "N"
		
	ggoSpread.SetCombo strCboCd,C_DefaultFlg	'parent.ggoSpread.SSGetColsIndex()              'Job Code setting 
    	
	'****************************
    'Price Control Ind
    '****************************
	strCboCd = "" 
	strCboNm = ""
	
	ggoSpread.Source = frm1.vspdData

    strCboCd = strCboCd & "S" & vbTab				'Setting Job Cd in Detail Sheet
    strCboNm = strCboNm & "ǥ�شܰ�" & vbTab    'Setting Job Nm in Detail Sheet

    strCboCd = strCboCd & "M"						'& vbTab		'Setting Job Cd in Detail Sheet
    strCboNm = strCboNm & "�̵���մܰ�"		'& vbTab            'Setting Job Nm in Detail Sheet

    ggoSpread.SetCombo strCboCd,C_PrcCtrlInd		'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
	ggoSpread.SetCombo strCboNm,C_PrcCtrlIndNm		'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
End Sub

'==========================================  2.2.6 InitSpreadComboBox()  =======================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display in Spread(s)
'========================================================================================================= 
Sub InitSpreadComboBox()
    Dim strCboCd 
    Dim strCboNm
    
    '****************************
    ' ǰ����� 
    '****************************    
    strCboCd = ""
    strCboNm = ""

	Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  
    	                 
	'Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))
	
    strCboCd = Replace(lgF0,chr(11),vbTab)
    strCboNm = Replace(lgF1,chr(11),vbTab)  
    
	ggoSpread.SetCombo strCboCd,C_HdnItmAcc
    ggoSpread.SetCombo strCboNm,C_ItmAcc
    
	'****************************
    ' ����,���ձ���,��ȿ���� 
    '****************************    
    strCboCd = ""
    strCboNm = ""
    
    strCboCd = "Y" & vbTab & "N"
		
	ggoSpread.SetCombo strCboCd,C_DefaultFlg	'parent.ggoSpread.SSGetColsIndex()              'Job Code setting 
	 	
	'****************************
    'Price Control Ind
    '****************************	
	strCboCd = "" 
	strCboNm = ""
	
	ggoSpread.Source = frm1.vspdData

    strCboCd = strCboCd & UCase("S") & vbTab		'Setting Job Cd in Detail Sheet
    strCboNm = strCboNm & "ǥ�شܰ�" & vbTab    'Setting Job Nm in Detail Sheet

    strCboCd = strCboCd & UCase("M") & vbTab		'Setting Job Cd in Detail Sheet
    strCboNm = strCboNm & "�̵���մܰ�" & vbTab            'Setting Job Nm in Detail Sheet

    ggoSpread.SetCombo strCboCd,C_PrcCtrlInd		'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
	ggoSpread.SetCombo strCboNm,C_PrcCtrlIndNm		'parent.ggoSpread.SSGetColsIndex()              'Job Code setting
End Sub

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
Sub InitData(ByVal lngStartRow, ByVal iPos)
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		'.ReDraw = False
		
		For intRow = lngStartRow To .MaxRows
			If iPos = 1 Then
				.Row = intRow
				.Col = C_HdnItmAcc
				intIndex = .value
				.col = C_ItmAcc
				.value = intindex

				.Row = intRow
				.Col = C_PrcCtrlInd
				intIndex = .value
				.col = C_PrcCtrlIndNm
				.value = intindex
				
			Else
				.Row = intRow
				.Col = C_ItmAcc
				intIndex = .value
				.col = C_HdnItmAcc
				.value = intindex
			
				.Row = intRow
				.Col = C_PrcCtrlInd
				intIndex = .value
				.col = C_PrcCtrlIndNm
				.value = intindex
			End IF							
		Next	
		'.ReDraw = True
	End With
End Sub

Function SetFieldProp(ByVal lRow, ByVal sType)
	ggoSpread.Source = frm1.vspdData
    
	ggoSpread.SSSetRequired	C_PrcCtrlInd,	lRow, lRow
	ggoSpread.SSSetRequired	C_UnitPrice,	lRow, lRow
End Function

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

'------------------------------------------  OpenCondPlant()  ---------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim activateField
	
	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet, 0)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(strCode, iPos)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function
	
	If iPos = 1 Then
		If frm1.txtPlantCd.value = "" Then
			Call DisplayMsgBox("971012","X", "����","x")
			frm1.txtPlantCd.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If		
	End If
	
	IsOpenPop = True
	
	If iPos = 0 Then
		arrParam(0) = strCode						' Item Code
		arrParam(1) = ""							' Item Name
		arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
		arrParam(3) = ""							' Default Value
		
		iCalledAspName = AskPRAspName("B1B01PA2")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA2", "X")
			IsOpenPop = False
			Exit Function
		End If
	ElseIf iPos = 1 Then
		arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
		arrParam(1) = strCode						' Item Code
		arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
		arrParam(3) = ""							' Default Value
		
		iCalledAspName = AskPRAspName("B1B11PA4")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
			IsOpenPop = False
			Exit Function
		End If
	End If

    arrField(0) = 1 								' Field��(0) :"ITEM_CD"
    arrField(1) = 2 								' Field��(1) :"ITEM_NM"
    arrField(2) = 3 								' Field��(2) :"SPEC"
    arrField(3) = 9 								' Field��(2) :"ProcurType"
    arrField(4) = 10 								' Field��(2) :"ProcurType"
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet,iPos)
	End If	

	Call SetFocusToDocument("M")
	If iPos = "0" Then
		frm1.txtItemCd.focus
	Else
		frm1.txtItemCd1.focus
	End If	

End Function

'------------------------------------------  OpenItemGroup()  ---------------------------------------------
'	Name : OpenItemGroup()
'	Description : Condition Item Group PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = frm1.txtHighItemGroupCd.value  
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  "
	arrParam(5) = "ǰ��׷�"
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	
'    arrField(3) = "LEAF_FLG"	
'    arrField(4) = "UPPER_ITEM_GROUP_CD"	
    
    arrHeader(0) = "ǰ��׷�"		
    arrHeader(1) = "ǰ��׷��"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemGroupCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtHighItemGroupCd.focus
	
End Function

'------------------------------------------  OpenClassCd()  -------------------------------------------------
'	Name : OpenClassCd()
'	Description : Class PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenClassCd()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtClasscd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtClassCd.value)	' Class Code
	arrParam(1) = ""							' Class Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field��(0) : "Class_CD"
    arrField(1) = 2 							' Field��(1) : "Class_NM"
	
	iCalledAspName = AskPRAspName("B3B31PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B31PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) <> "" Then
		Call SetClassCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtClassCd.focus
	
End Function

'==========================================================================================
'   Event Name : txtItemCd1_onChange()
'   Event Desc :
'==========================================================================================
Sub txtItemCd1_onChange()
	With frm1
		If .txtItemCd1.value = "" Then
			.txtItemNm1.value = ""
			.txtItemSpec1.value = ""
			.txtItemProcType1.value = ""
	
			.txtItemCd1.focus
			Set gActiveElement = document.activeElement
		Else	
			Call LookUpItemByPlant()
		End If
	End With
End Sub

'-------------------------------------  LookUpItem ByPlant()  -----------------------------------------
'	Name : LookUpItem ByPlant()
'	Description : LookUp Item By Plant
'--------------------------------------------------------------------------------------------------------- 
Function LookUpItemByPlant()
	Dim iStrWhereSQL
	Dim strITEM_CD
	Dim strITEM_NM
	Dim strSPEC
	Dim strPROCUR_TYPE_CD
	Dim strPROCUR_TYPE_NM

	iStrWhereSQL = "A.ITEM_CD = B.ITEM_CD AND A.ITEM_CD = " & FilterVar(frm1.txtItemCd1.value, "''", "S") & " AND B.PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")
	Call CommonQueryRs(" A.ITEM_CD, A.ITEM_NM, A.SPEC, B.PROCUR_TYPE, dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & " , B.PROCUR_TYPE) "," B_ITEM A, B_ITEM_BY_PLANT B ",iStrWhereSQL ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	strITEM_CD = lgF0
	strITEM_NM = lgF1
	strSPEC = lgF2
	strPROCUR_TYPE_CD = lgF3
	strPROCUR_TYPE_NM = lgF4
		
	strITEM_CD			=	replace(strITEM_CD,Chr(11),"")
	strITEM_NM			=	replace(strITEM_NM,Chr(11),"")
	strSPEC				=	replace(strSPEC,Chr(11),"")
	strPROCUR_TYPE_CD	=	replace(strPROCUR_TYPE_CD,Chr(11),"")
	strPROCUR_TYPE_NM	=	replace(strPROCUR_TYPE_NM,Chr(11),"")
	
	frm1.txtItemNm1.value = strITEM_NM
	frm1.txtItemSpec1.value = strSPEC
	frm1.txtItemProcType1.value = strPROCUR_TYPE_NM
	frm1.htxtItemProcType1.value = strPROCUR_TYPE_CD
End Function

'------------------------------------------  SetClassCd()  ------------------------------------------------
'	Name : SetClassCd()
'	Description : Class Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetClassCd(byval arrRet)
	frm1.txtClassCd.Value    = arrRet(0)		
	frm1.txtClassNm.Value    = arrRet(1)
	frm1.txtClassCd.focus
	Set gActiveElement = document.activeElement 		
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemInfo(Byval arrRet,ByVal iPos)
	With frm1
		If iPos = 0 Then
			.txtItemCd.value = arrRet(0)
			.txtItemNm.value = arrRet(1)
		Else
			.txtItemCd1.value	= arrRet(0)
			.txtItemNm1.value	= arrRet(1)
			.txtItemSpec1.value = arrRet(2)
			.txtItemProcType1.value = arrRet(4)
			.htxtItemProcType1.value = arrRet(3)
		End If

	End With
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet, byval iPos)
	With frm1
		If iPos = 0 Then
			.txtPlantCd.Value    = arrRet(0)		
			.txtPlantNm.Value    = arrRet(1)
		Else
			.txtPlantCd1.Value    = arrRet(0)		
		End If
	End With
End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetUnit(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_Unit
	frm1.vspdData.Text = arrRet(0)
End Function

'------------------------------------------  SetItemGroup()  --------------------------------------------------
'	Name : SetItemGroup()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemGroup(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ItmGroupCd
	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col = C_ItmGroupNm
	frm1.vspdData.Text = arrRet(1)		
End Function

'------------------------------------------  SetItemGroupCd()  --------------------------------------------------
'	Name : SetItemGroupCd()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemGroupCd(Byval arrRet)
	frm1.txtHighItemGroupCd.value = arrRet(0)
	frm1.txtHighItemGroupNm.value = arrRet(1)
End Function

'------------------------------------------  SetBaseItem()  --------------------------------------------------
'	Name : SetBaseItem()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBasisItemCd(Byval arrRet)
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_BaseItm
	frm1.vspdData.Text = arrRet(0)
	frm1.vspdData.Col = C_BaseItmNm
	frm1.vspdData.Text = arrRet(1)		

End Function

'------------------------------------------  ChkBtnAll()  --------------------------------------------------
'	Name : ChkBtnAll()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function btnSelectAll_Clicked()
	Dim LngRow
	
	If frm1.vspdData.MaxRows <= 0 Then Exit Function

	lgFlgBtnSelectAllClicked = True
	frm1.btnSelectAll.disabled = True
	
	With frm1.vspdData
		
		.ReDraw = False

		If lgFlgAllSelected = False Then 'select all clicked
				
			For LngRow = 1 To .MaxRows
				Call .SetText(C_Select,LngRow,"1")
				Call SetSpreadUnLock(C_Select, LngRow)
				If lgInsrtFlg <> True Then
					ggoSpread.UpdateRow LngRow
				End If
			Next

			Call InitData(1,1)	
			
			frm1.btnSelectAll.value = "��ü�������"
			lgFlgAllSelected = True
			
		Else 'deselect all clicked
			
			For LngRow = 1 To .MaxRows
				If GetSpreadText(frm1.vspdData,C_Select,LngRow,"X","X") = "1" _
				And GetSpreadText(frm1.vspdData,0,LngRow,"X","X") <> ggoSpread.InsertFlag Then
					Call .SetText(C_Select,LngRow,"0")
					Call ggoSpread.EditUndo(LngRow, LngRow)
					Call SetSpreadLock1(C_Select, LngRow)
				End If
			Next
			
			Call InitData(1,1)

			frm1.btnSelectAll.value = "��ü����"
			lgFlgAllSelected = False
		End If
		.ReDraw = True
	End With
	
	frm1.btnSelectAll.disabled = False
	lgFlgBtnSelectAllClicked = False		

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
	Err.Clear
	
	iDBSYSDate = "<%=GetSvrDate%>"											'��: DB�� ���� ��¥�� �޾ƿͼ� ���۳�¥�� ����Ѵ�.
	StartDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
	
	Call LoadInfTB19029																'��: Load table , B_numeric_format
	Call AppendNumberPlace("6","3","2")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
	
	'----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet															'��: Setup the Spread sheet
	Call InitComboBox
	Call SetDefaultVal	'�Լ� ���ǰ� ���� 
	Call InitVariables	'�Լ� ���ǰ� ����											'��: Initializes local global variables
	
	Call SetToolbar("11000000000011")												'��: ��ư ���� ���� 
	
	If frm1.txtPlantCd.value = "" Then
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
			frm1.txtPlantCd1.value = parent.gPlant
			frm1.txtItemCd.focus
			Set gActiveElement = document.activeElement  
		Else
			frm1.txtPlantCd.focus
			Set gActiveElement = document.activeElement
		End If
	End If
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Dim IntRetCD
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("0000110111")
	Else 	
		If frm1.vspdData.MaxRows = 0 Then 
			Call SetPopupMenuItemInf("0000110111")
		Else
			Call SetPopupMenuItemInf("0001111111") 
		End if			
	End If	
	
	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData  
	
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
       
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
    End If
    
	If Row <= 0 Or Col < 0 Then
		ggoSpread.Source = frm1.vspdData
		Exit Sub
	End If
	
	frm1.vspdData.Row = Row
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
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button = "2" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


'==========================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'==========================================================================================
Sub vspdData_DblClick(ByVal Col , ByVal Row )
    If Row <= 0 Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
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


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		.Col = C_Select
		.Row = Row

		If .Value = "1" Then
			ggoSpread.UpdateRow Row
		End if
		
		If Col = C_PrcCtrlIndNm Then
		   Call vspdData_ComboSelChange (C_PrcCtrlIndNm,Row)
		End If   

	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp 
	Dim intPos1

	'----------  Coding part  -------------------------------------------------------------   

	If frm1.vspdData.Row <= 0 Or lgFlgBtnSelectAllClicked = True Then Exit Sub
	
	ggoSpread.Source = frm1.vspdData
	
	With frm1.vspdData
		If gMouseClickStatus = "SPC" Or lgFlgCancelClicked = True Then
			If Col = C_Select And Not (lgFlgCopyClicked) Then
				If GetSpreadText(frm1.vspdData,C_Select,Row,"X","X") = "0" Then
					.Redraw = false
					Call SetSpreadLock1(C_Select, Row)
					call ggoSpread.EditUndo(Row, Row)
					Call InitData(1,1)
					.Redraw = true
				Else
					.Redraw = false
					Call SetSpreadUnLock(C_Select, Row)	
					.Redraw = true
				End If
			End If
		End If
				
		Select Case Col
			Case C_Select
				If lgInsrtFlg <> True Then
					If Buttondown = 1 Then
						ggoSpread.Source = frm1.vspdData
						ggoSpread.UpdateRow Row
					Else
						ggoSpread.Source = frm1.vspdData
						ggoSpread.SSDeleteFlag Row,Row
					End If
				End If
		End Select
    End With
End Sub

'==========================================================================================
'   Event Name :vspddata_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
			Case  C_ItmAcc
				.Col = Col
				intIndex = .Value
				.Col = C_HdnItmAcc
				.Value = intIndex
			'Case  C_SumItmClass
			'	.Col = Col
			'	intIndex = .Value
			'	.Col = C_HdnSumItmClass
			'	.Value = intIndex
			Case  C_PrcCtrlIndNm
				.Col = Col
				intIndex = .value
				.Col = C_PrcCtrlInd
				.value = intIndex
				If Trim(.Text) = "M" Then
					ggoSpread.SpreadLock		C_UnitPrice,		Row, C_UnitPrice,		Row
					ggoSpread.SSSetProtected 	C_UnitPrice, 		Row, Row
				Else
					ggoSpread.SpreadUnLock		C_UnitPrice,		Row, C_UnitPrice,		Row
					ggoSpread.SSSetRequired 	C_UnitPrice, 		Row, Row
				End If
						
		End Select
    
    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row >= NewRow Then Exit Sub
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    'If CheckRunningBizProcess = True Then
	'	Exit Sub
	'End If	
	
	'----------  Coding part  ------------------------------------------------------------- 				
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'��: ������ üũ	
    	If lgStrPrevKey <> "" Then		                                            '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(Parent.TBC_QUERY)									': Query ��ư�� disable ��Ŵ.
			frm1.vspdData.ReDraw = False
			If DbQuery = False Then
				Call RestoreToolBar()
				frm1.vspdData.ReDraw = True
				Exit Sub
			End if
			frm1.vspdData.ReDraw = True
    	End If
    End If

End Sub
'==========================================================================================
'   Event Name : txtPlantCd_OnChange
'   Event Desc : This function is Setting the txtPlantCd,txtPlantNm
'==========================================================================================
Sub txtPlantCd_OnBlur()
	With frm1
		If Trim(.txtPlantCd.value) = "" Then
			.txtPlantNm.value = ""
			.txtPlantCd1.value = ""
		Else
			.txtPlantCd1.value = UCase(Trim(.txtPlantCd.value))
		End If
	End With		
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
    Dim strPlantCd
    Dim strPlantNm
    Dim strPlantItem
    Dim strPlantItemNm
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

	'-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			'��: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    	
	Call ggoOper.ClearField(Document, "3")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables
	
	 '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
   
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
	
	If frm1.txtItemCd1.value = "" Then
		frm1.txtItemNm1.value = ""
	Else
		strPlantItem = frm1.txtItemCd1.value 
		strPlantItemNm = frm1.txtItemNm1.value 
	End If
	
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	    
    'Call SetDefaultVal
	
	If strPlantItem <> "" Then    
		frm1.txtItemCd1.value = strPlantItem
		frm1.txtItemNm1.value = strPlantItemNm
	End If
   
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

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
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    On Error Resume Next                                                    '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                            '��: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If frm1.txtItemCd1.value = "" Then
		frm1.txtItemNm1.value = ""
	End If
	    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Precheck area
    '-----------------------
   	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     							                                      '��: Save db data
    
    FncSave = True                                                          '��: Processing is OK
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
    On Error Resume Next                                                          '��: If process fails
    Err.Clear
End Function


'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 

	If frm1.vspdData.MaxRows <= 0 Then Exit Function
	
	lgFlgCancelClicked = True
	
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		.ReDraw = False
		Call ggoSpread.EditUndo(.ActiveRow,.ActiveRow)
		Call InitData(1,1)
		Call SetSpreadLock1(C_Select, .ActiveRow)
		.ReDraw = True
	End With
	
	lgFlgCancelClicked = False
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
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()	
	Call InitData(1,1)
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function
'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
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
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================
Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function

'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()												'��: ���� ������ ���� ���� 
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    DbQuery = False                                                         '��: Processing is NG
    
    Call LayerShowHide(1)
    
    Dim strVal
	
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001								'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)					'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtClassCd=" & Trim(frm1.htxtClassCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd1=" & Trim(frm1.txtItemCd1.value)				'��: ��ȸ ���� ����Ÿ 
		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows

    Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtHighItemGroupCd.value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtClassCd=" & Trim(frm1.txtClassCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd1=" & Trim(frm1.txtItemCd1.value)				'��: ��ȸ ���� ����Ÿ 
			
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows

    End If
		
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                                          '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(LngMaxRow)													'��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If
    
    lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
    
    Call InitData(LngMaxRow,1)
	
    Call ggoOper.LockField(Document, "Q")										'��: This function lock the suitable field
    
    Call SetToolbar("11000001000111")
    
    frm1.btnCopy.disabled = True
	frm1.btnSelectAll.disabled = True
	frm1.btnSelectAll.value = "��ü����"
	lgFlgAllSelected = False    
    
    lgBlnFlgChgValue = False
    
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim IntRows 
    Dim lGrpcnt 
    Dim strVal
    Dim GenVal
	Dim IntRetCD
	Dim iColSep
	Dim TmpBuffer
	Dim iTotalStr
	
	DbSave = False														'��: Processing is NG
    Call LayerShowHide(1)

    With frm1
		.txtMode.value = parent.UID_M0002									'��: ���� ���� 
		.txtFlgMode.value = lgIntFlgMode									'��: �ű��Է�/���� ���� 
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    ReDim TmpBuffer(0)
    iColSep = Parent.gColSep 
	GenVal = "10000000"
	
	If UCase(frm1.htxtItemProcType1.value) = "M" Then
		
	End If	
	
	With frm1.vspdData

		For IntRows = 1 To .MaxRows
    
			.Row = IntRows
			.Col = 0
	
			Select Case .Text
				Case ggoSpread.UpdateFlag
					
					strVal = ""
						
					strVal = strVal & "C" & iColSep	& IntRows & iColSep				'��: U=Update
					
					.Col = C_Item								'2
					strVal = strVal & Trim(.Text) & iColSep

					.Col = C_PrcCtrlInd							'3
					strVal = strVal & Trim(.Text) & iColSep		
					
					If (Trim(UCase(frm1.htxtItemProcType1.value)) = "M" And Trim(UCase(.Text)) = "M") Or (Trim(UCase(frm1.htxtItemProcType1.value)) = "O" And Trim(UCase(.Text)) = "M") Then
						IntRetCD = DisplayMsgBox("122726", parent.VB_INFORMATION, "X", "X")	'���ޱ����� �系����ǰ�̸� �ܰ������� ǥ�شܰ��� �����մϴ�.
						Call LayerShowHide(0)
						Exit Function
					End If
					
					.Col = C_UnitPrice							'4
					strVal = strVal & UNIConvNum(Trim(.Text),0) & iColSep		

					.Col = C_Phantom							'5
			        strVal = strVal & Trim(.Text) & iColSep

					.Col = C_IBPValidFromDt						'6
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep						
						    
					.Col = C_IBPValidToDt						'7
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep

					.Col = C_ValidFromDt						'8
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep						
						    
					.Col = C_ValidToDt						'9
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep

					strVal = strVal & GenVal & parent.gRowSep			'10			'��: ������ ����Ÿ�� Row �и���ȣ�� �ִ´�		        
					
					ReDim Preserve TmpBuffer(lGrpcnt-1)
					
					TmpBuffer(lGrpcnt-1) = strVal
							
					lGrpcnt = lGrpcnt + 1             
			End Select
	   Next
	
	End With
	
	iTotalStr = Join(TmpBuffer, "")
	
	frm1.txtMaxRows.value = lGrpCnt-1										'��: Spread Sheet�� ����� �ִ밹�� 
	frm1.txtSpread.value = iTotalStr										'��: Spread Sheet ������ ���� 
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'��: ���� �����Ͻ� ASP �� ���� 

    DbSave = True                                                           '��: Processing is OK
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	Call InitVariables

    ggoSpread.Source = frm1.vspdData
    frm1.vspdData.MaxRows = 0

    Call FncQuery()
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

</HEAD>
<!--'#########################################################################################################
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���庰ǰ������ COPY</font></td>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="24"></TD>
									<TD CLASS=TD5 NOWRAP>Ŭ����</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtClassCd" SIZE=20 MAXLENGTH=16 tag="11XXXU"  ALT="Ŭ����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnClassCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenClassCd()">
										<INPUT TYPE=TEXT NAME="txtClassNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value,0">
										<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14">
									</TD>
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=TEXT NAME="txtHighItemGroupCd" SIZE=20 MAXLENGTH=10 tag="11XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btHighItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()" >
										<INPUT TYPE=TEXT NAME="txtHighItemGroupNm" SIZE=20 tag="14">
									</TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>����ǰ��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=25 MAXLENGTH=18 tag="22XXXU" ALT="����ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd1.value,1"></TD>
								<TD CLASS="TD5" NOWRAP>����ǰ��԰�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemSpec1" SIZE=40 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����ǰ���</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=40 tag="24"></TD>
								<TD CLASS="TD5" NOWRAP>����ǰ�����ޱ���</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemProcType1" SIZE=40 tag="24"></TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="32" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<BUTTON NAME="btnCopy" CLASS="CLSMBTN" Flag=1 ONCLICK="FncSave">COPY</BUTTON>&nbsp;
						<BUTTON NAME="btnSelectAll" CLASS="CLSMBTN" Flag=1 ONCLICK="btnSelectAll_Clicked">��ü����</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TabIndex="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="hItemAccount" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hPhantomFlg" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="hItemClass" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="htxtClassCd" tag="24" TabIndex="-1">
<INPUT TYPE=HIDDEN NAME="txtPlantCd1" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="htxtItemProcType1" tag="24" TabIndex="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TabIndex="-1"></iframe>
</DIV>
</BODY>
</HTML>
