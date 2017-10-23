
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1504ma1.asp
'*  4. Program Name         : Shift Exception
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : RYU SUNG WON
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID  = "p1504mb1.asp"												'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "p1504mb2.asp"											'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const C_SHEETMAXROWS = 30

Dim C_ExceptionCd
Dim C_ExceptionNm
Dim C_ExceptionWorkFlgNm
Dim C_ExceptionStartDt
Dim C_ExceptionStartTime
Dim C_ExceptionEndDt
Dim C_ExceptionEndTime
Dim C_ExceptionType
Dim C_ExceptionWorkFlg


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->
'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop
Dim iDBSYSDate
Dim StartDate, EndDate
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

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
	C_ExceptionCd			= 1
	C_ExceptionNm			= 2
	C_ExceptionWorkFlgNm	= 3
	C_ExceptionStartDt		= 4
	C_ExceptionStartTime	= 5
	C_ExceptionEndDt		= 6
	C_ExceptionEndTime		= 7
	C_ExceptionType			= 8
	C_ExceptionWorkFlg		= 9
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
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

Sub SetReadCookie
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value	= ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value	= ReadCookie("txtPlantNm")
		frm1.txtResourceCd.Value= ReadCookie("txtResourceCd")
		frm1.txtResourceNm.value= ReadCookie("txtResourceNm")
		frm1.txtShiftCd.Value	= ReadCookie("txtShiftCd")
		frm1.txtShiftNm.value	= ReadCookie("txtShiftNm")		
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtResourceCd", ""
	WriteCookie "txtResourceNm", ""
	WriteCookie "txtShiftCd", ""
	WriteCookie "txtShiftNm", ""	
	
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
    ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread    
    .MaxCols = C_ExceptionWorkFlg+1    
    .MaxRows = 0    

	.ReDraw = false
	
	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit 	C_ExceptionCd,			"�����ڵ�",			13,,,10,2
	ggoSpread.SSSetEdit 	C_ExceptionNm,			"���ܳ���",			34,,,40
	ggoSpread.SSSetCombo	C_ExceptionWorkFlgNm,	"�ܾ�/Break ����",	20
	ggoSpread.SSSetDate 	C_ExceptionStartDt,		"������",			12, 2, parent.gDateFormat
	ggoSpread.SSSetTime 	C_ExceptionStartTime,	"���۽ð�",			12, 2, 1 ,1	    
	ggoSpread.SSSetDate 	C_ExceptionEndDt,		"������",			12, 2, parent.gDateFormat
	ggoSpread.SSSetTime 	C_ExceptionEndTime,		"����ð�",			12, 2, 1 ,1	   		
	ggoSpread.SSSetCombo	C_ExceptionType,		"����Ÿ��",			12
	ggoSpread.SSSetCombo	C_ExceptionWorkFlg,		"�ܾ�/Break ����",	15

	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	Call ggoSpread.SSSetColHidden(.MaxCols - 1, .MaxCols - 1, True)
	Call ggoSpread.SSSetColHidden(.MaxCols - 2, .MaxCols - 2, True)

	ggoSpread.SSSetSplit2(1)										'frozen ����߰� 
	.ReDraw = true

	Call SetSpreadLock 

    End With

End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock		C_ExceptionCd, -1, C_ExceptionCd
		ggoSpread.SSSetRequired 	C_ExceptionWorkFlgNm, -1	
		ggoSpread.SSSetRequired		C_ExceptionStartDt, -1
		ggoSpread.SSSetRequired		C_ExceptionEndDt, -1
		ggoSpread.SSSetRequired		C_ExceptionStartTime, -1
		ggoSpread.SSSetRequired		C_ExceptionEndTime, -1
		ggoSpread.SSSetProtected	.vspdData.MaxCols, -1
		.vspdData.ReDraw = True
	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False

		ggoSpread.SSSetRequired 	C_ExceptionCd,			pvStartRow, pvEndRow
		ggoSpread.SSSetRequired 	C_ExceptionWorkFlgNm,	pvStartRow, pvEndRow		
		ggoSpread.SSSetRequired 	C_ExceptionStartDt,		pvStartRow, pvEndRow	
		ggoSpread.SSSetRequired 	C_ExceptionEndDt,		pvStartRow, pvEndRow
		ggoSpread.SSSetRequired 	C_ExceptionStartTime,	pvStartRow, pvEndRow	
		ggoSpread.SSSetRequired 	C_ExceptionEndTime,		pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	.vspdData.MaxCols,		pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
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
			C_ExceptionCd			= iCurColumnPos(1)
			C_ExceptionNm			= iCurColumnPos(2)
			C_ExceptionWorkFlgNm	= iCurColumnPos(3)
			C_ExceptionStartDt		= iCurColumnPos(4)
			C_ExceptionStartTime	= iCurColumnPos(5)
			C_ExceptionEndDt		= iCurColumnPos(6)
			C_ExceptionEndTime		= iCurColumnPos(7)
			C_ExceptionType			= iCurColumnPos(8)
			C_ExceptionWorkFlg		= iCurColumnPos(9)
    End Select    
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Dim strCboCd
	Dim strCboNm
	
	strCboCd = ""
	strCboCd = "Y" & vbTab & "N"
	strCboNm = ""
	strCboNm = "OverTime" & vbTab & "Downtime"

	ggoSpread.SetCombo	strCboCd,	C_ExceptionWorkFlg
	ggoSpread.SetCombo	strCboNm,	C_ExceptionWorkFlgNm
	
	strCboCd = "R" & vbTab 
	    
    ggoSpread.SetCombo	strCboCd,	C_ExceptionType
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
'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function
'------------------------------------------  OpenResource()  -------------------------------------------------
'	Name : OpenResource()
'	Description : Resource PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResource()

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)


	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
			
	IsOpenPop = True
	arrParam(0) = "�ڿ��˾�"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = Trim(frm1.txtResourceCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "			
	arrParam(5) = "�ڿ�"
	
    arrField(0) = "RESOURCE_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "�ڿ�"		
    arrHeader(1) = "�ڿ���"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetResource(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtResourceCd.focus
		
End Function

'------------------------------------------  OpenShift()  -------------------------------------------------
'	Name : OpenShift()
'	Description : ShiftPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenShift()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	elseif frm1.txtResourceCd.value = "" Then
		Call DisplayMsgBox("971012","X", "�ڿ�","X")
		frm1.txtResourceCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True	
	
	arrParam(0) = "Shift�˾�"	
	arrParam(1) = "P_SHIFT_HEADER,P_RESOURCE_ON_SHIFT"				
	arrParam(2) = Trim(frm1.txtShiftCd.Value)
	arrParam(3) = ""
	arrParam(4) = "P_RESOURCE_ON_SHIFT.RSC_PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " " & _
				  " AND P_RESOURCE_ON_SHIFT.SHIFT_PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " " & _
				  " AND P_RESOURCE_ON_SHIFT.RESOURCE_CD =  " & FilterVar(frm1.txtResourceCd.value, "''", "S") & " " & _
				  " AND P_SHIFT_HEADER.PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " " & _
				  " AND P_RESOURCE_ON_SHIFT.SHIFT_CD = P_SHIFT_HEADER.SHIFT_CD " 			  
	
	arrParam(5) = "Shift��"			
	
    arrField(0) = "P_RESOURCE_ON_SHIFT.SHIFT_CD"	
    arrField(1) = "P_SHIFT_HEADER.DESCRIPTION"	
       
    
    arrHeader(0) = "Shift"		
    arrHeader(1) = "Shift��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetShift(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtShiftCd.focus
	
End Function
'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetResource()  --------------------------------------------------
'	Name : SetResource()
'	Description : Resource Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResource(byval arrRet)
	frm1.txtResourceCd.Value    = arrRet(0)		
	frm1.txtResourceNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetShift()  --------------------------------------------------
'	Name : SetShift()
'	Description : Shift Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetShift(Byval arrRet)
	frm1.txtShiftCd.Value    = arrRet(0)		
	frm1.txtShiftNm.Value    = arrRet(1)		
End Function
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 


Sub InitData(ByVal lngStartRow, ByVal iPos)
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		.ReDraw = False
		
		For intRow = lngStartRow To .MaxRows
			.Row = intRow
			.Col = C_ExceptionWorkFlg
			intIndex = .value
			.col = C_ExceptionWorkFlgNm
			.value = intindex
		Next	
		.ReDraw = True
	End With
End Sub

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
	
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                            '��: Lock  Suitable  Field    
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    Call InitVariables                                                      '��: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------
    
    Call InitComboBox
    Call SetToolbar("11001101000011")								'��: ��ư ���� ���� 
    Call SetReadCookie   	
    	
    If Trim(frm1.txtPlantCd.value) = "" Then	
		If parent.gPlant <> "" then	
			frm1.txtPlantCd.value = parent.gPlant	
			frm1.txtPlantNm.value = parent.gPlantNm	
			frm1.txtResourceCd.focus 	
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
Sub vspddata_Click(ByVal Col , ByVal Row )
	Dim IntRetCD
	
	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("1101111111")    
	
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
'   Event Name : vspddata_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspddata_MouseDown(Button,Shift,x,y)
	If Button = "2" And gMouseClickStatus = "SPC" Then
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
End Sub


'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'==========================================================================================
Sub vspddata_KeyPress(index , KeyAscii)
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
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
			Case  C_ExceptionWorkFlgNm
				.Col = Col
				intIndex = .Value
				.Col = C_ExceptionWorkFlg
				.Value = intIndex
		End Select
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
    FncQuery = False															'��: Processing is NG
    Err.Clear																    '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
		
	If frm1.txtResourceCd.value = "" Then
		frm1.txtResourceNm.value = ""
	End If
	
	If frm1.txtShiftCd.value = "" Then
		frm1.txtShiftNm.value = ""
	End If
	
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables
  
    If Not chkField(Document, "1") Then	Exit Function					'��: This function check indispensable field
    If DbQuery = False Then				Exit Function

    FncQuery = True																'��: Processing is OK
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
    
    FncSave = False																'��: Processing is NG
    
    On Error Resume Next														'��: Protect system from crashing
    Err.Clear																	'��: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")								'��: No data changed!!
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then	Exit Function
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		If Not chkField(Document, "1") Then	Exit Function
    End IF

    If DbSave = False Then	Exit Function
    
    FncSave = True																'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD

    On Error Resume Next                                                          '��: If process fails
    Err.Clear
    
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
	FncCopy = False                                                               '��: Processing is NG
	ggoSpread.Source = frm1.vspdData	
	
	With frm1.vspdData
		.ReDraw = False
		If .ActiveRow > 0 Then
			ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
			
			.EditMode = True
			.ReDraw = True
			.Focus
		End If
	End With
	
	If Err.number = 0 Then	
       FncCopy = True                                                            '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo															'��: Protect system from crashing
    Call InitData(1,1)
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
    Dim imRow
	Dim newRow
    
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False
    
    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then Exit Function
    End If
    
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		
		'--------------------------------------------
		' Default Setting
		' �߰��� Row������ŭ �÷��� �ʱ�ȭ ��Ų��.
		'--------------------------------------------
		For newRow = 0 To Cint(imRow) - 1
			.vspdData.Col = C_ExceptionType
			.vspdData.Row = .vspdData.ActiveRow + newRow
			.vspdData.Text = "R"
    
			.vspdData.Col = C_ExceptionWorkFlg
			.vspdData.Row = .vspdData.ActiveRow + newRow
			.vspdData.Text = "Y"
    
			.vspdData.Col = C_ExceptionStartDt
			.vspdData.Row = .vspdData.ActiveRow + newRow 
			.vspdData.Text = startdate
    
			.vspdData.Col = C_ExceptionEndDt
			.vspdData.Row = .vspdData.ActiveRow  + newRow
			.vspdData.Text = startdate
		Next

		.vspdData.EditMode = True
		.vspdData.ReDraw = True
    End With    
	
	Set gActiveElement = document.activeElement 
	    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If 
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
   
    '----------------------
    ' �����Ͱ� ���� ��� 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function
    
    With frm1.vspdData 
		.focus
		Set gActiveElement = document.activeElement 
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
    Call InitComboBox
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
	Dim strVal
	    
    DbQuery = False
    LayerShowHide(1) 
    Err.Clear                                                               '��: Protect system from crashing
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtResourceCd=" & Trim(.hResourceCd.value)
			strVal = strVal & "&txtShiftCd=" & Trim(.hShiftCd.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtResourceCd=" & Trim(.txtResourceCd.value)
			strVal = strVal & "&txtShiftCd=" & Trim(.txtShiftCd.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk(ByVal LngMaxRow)														'��: ��ȸ ������ ������� 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    Call InitData(LngMaxRow,1)
    
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	Call SetToolbar("11001111001111")
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
    
    Dim lRow        
    Dim lGrpCnt     
   	Dim strVal
	Dim strDel
	
    DbSave = False                                                          '��: Processing is NG
    
    LayerShowHide(1) 
		
    On Error Resume Next                                                   '��: Protect system from crashing
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		.txtFlgMode.value = lgIntFlgMode
    
		lGrpCnt = 1
    
		strVal = ""
		strDel = ""
    
		For lRow = 1 To .vspdData.MaxRows
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text
		        Case ggoSpread.InsertFlag												'��: �ű� 
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep					'��: C=Create
					
		            .vspdData.Col = C_ExceptionCd	'1
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionNm	'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_ExceptionWorkFlg	'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionStartDt	'4
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionStartTime	'5
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionEndDt	'6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					
					.vspdData.Col = C_ExceptionEndTime	'7
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_ExceptionType	'8
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
		                            
		            lGrpCnt = lGrpCnt + 1
		        Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep	& lRow & parent.gColSep					'��: U=Update
					
		            .vspdData.Col = C_ExceptionCd	'1
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionNm	'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionWorkFlg	'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionStartDt	'4
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionStartTime	'5
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionEndDt	'6
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_ExceptionEndTime	'7
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep

		            .vspdData.Col = C_ExceptionType	'8
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
		                                                                           
		            lGrpCnt = lGrpCnt + 1
		        Case ggoSpread.DeleteFlag												'��: ���� 
					strDel = strDel & "D" & parent.gColSep	& lRow & parent.gColSep
					
		            .vspdData.Col = C_ExceptionCd	'1
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
		                            
		            lGrpCnt = lGrpCnt + 1
		    End Select
		            
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
	
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'��: �����Ͻ� ASP �� ���� 
	End With
	
    DbSave = True																	'��: Processing is NG
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()																	'��: ���� ������ ���� ���� 
	Call InitVariables
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.MaxRows = 0
	
    Call MainQuery()
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 

End Function
Function DbDeleteOk()

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
<TABLE  <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Shift���ܵ��</font></td>
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
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14" ALT="�����"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�ڿ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd" SIZE=15 MAXLENGTH=10 tag="12XXXU" ALT="�ڿ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResource()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm" SIZE=30 tag="14" ALT="�ڿ���"></TD>								
									<TD CLASS=TD5 NOWRAP>Shift</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtShiftCd" SIZE=8 MAXLENGTH=2 tag="12XXXU" ALT="Shift"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnShiftCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenShift()">&nbsp;<INPUT TYPE=TEXT NAME="txtShiftNm" SIZE=25 tag="14" ALT="Shift��"></TD>
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
							<TD HEIGHT=100%>
								<script language =javascript src='./js/p1504ma1_I212316697_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hResourceCd" tag="24"><INPUT TYPE=HIDDEN NAME="hShiftCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

