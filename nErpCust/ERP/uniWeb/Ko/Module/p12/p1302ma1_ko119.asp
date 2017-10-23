<%@ LANGUAGE="VBSCRIPT" %> 
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1302ma1_ko119.asp
'*  4. Program Name         : �۾����ø� ���� ������� ��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2006/04/06
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
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
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'Const BIZ_PGM_QRY_ID  = "p1302mb1_ko119.asp"												'��: �����Ͻ� ���� ASP�� 
'Const BIZ_PGM_SAVE_ID = "p1302mb1_ko119.asp"
Const BIZ_PGM_ID = "p1302mb1_ko119.asp"	                                 'Biz Logic ASP 					
'Const BIZ_PGM_JUMPSHIFTEXECEPTION_ID  = "p1504ma1"						'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Dim C_Plant_Cd
Dim C_Line_GroupCd
Dim C_Line_Group
Dim C_Work_Line
Dim C_Work_Line_Desc
Dim C_Remark

Dim lgStrPrevKeyPlantCd
Dim lgStrPrevKeyWorkLine
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop          
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
    C_Plant_Cd          = 1
    C_Line_GroupCd		= 2 
	C_Line_Group		= 3
	C_Work_Line			= 4
	C_Work_Line_Desc	= 5
	C_Remark			= 6
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
    
'    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKeyPlantCd = ""
    lgStrPrevKeyWorkLine = ""
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

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row = intRow
		
			.Col = C_LINE_GROUPCD
			intIndex = .Value
			.Col = C_LINE_GROUP
			.Value = intIndex
			
		Next	
	End With
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
		
		.MaxCols = C_Remark+1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.MaxRows = 0
		
		.ReDraw = false
		
		Call ggoSpread.ClearSpreadData()    '��: Clear spreadsheet data 
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit 		C_Plant_Cd,  "�����ڵ�"		,15,,,2,2
		ggoSpread.SSSetCombo		C_Line_GroupCd,	"���α׷�CD"	,16	,0	
		ggoSpread.SSSetCombo		C_Line_Group, "���α׷�"	,16	,0	
		ggoSpread.SSSetEdit 		C_Work_Line,  "Line"		,20,,,10,2
		ggoSpread.SSSetEdit 		C_Work_Line_Desc,  "Line��"	,25,,,30,2
		ggoSpread.SSSetEdit 		C_Remark,  "���"	,40,,,100,2
'		ggoSpread.SSSetButton 		C_ShiftPopup
'		ggoSpread.SSSetEdit 		C_ShiftNm,	"Shift��",	94,,,40
	
'		Call ggoSpread.MakePairsColumn(C_Work_Line,C_Work_Line_Desc)

		Call ggoSpread.SSSetColHidden(C_Plant_Cd ,C_Plant_Cd	,True)
		Call ggoSpread.SSSetColHidden(C_Line_GroupCd ,C_Line_GroupCd	,True)
		Call ggoSpread.SSSetColHidden(C_Work_Line_Desc ,C_Work_Line_Desc	,True)
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)

'		ggoSpread.SSSetSplit2(1)										'frozen ����߰� 
		.ReDraw = true

		Call SetSpreadLock
		Call InitData
'		Call initComboBox_two 
    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadLock()
    With frm1
		ggoSpread.Source = .vspdData
	
		.vspdData.ReDraw = False
'		ggoSpread.SpreadLock	 -1, -1 
		ggoSpread.SSSetRequired C_Line_Group	, -1, C_Line_Group
		ggoSpread.SSSetProtected C_Work_Line		, -1, C_Work_Line
		ggoSpread.SSSetProtected C_Work_Line_Desc	, -1, C_Work_Line_Desc
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
'		ggoSpread.SSSetRequired 	C_ShiftCd,			pvStartRow, pvEndRow
'		ggoSpread.SSSetProtected	C_ShiftNm,			pvStartRow, pvEndRow
        ggoSpread.SSSetRequired		C_Line_Group		,pvStartRow	,pvEndRow
        ggoSpread.SSSetRequired		C_Work_Line			,pvStartRow	,pvEndRow
'		ggoSpread.SSSetRequired		C_Work_Line_Desc	,pvStartRow	,pvEndRow
		
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
            C_Plant_Cd          = iCurColumnPos(1)
            C_Line_GroupCd		= iCurColumnPos(2)
			C_Line_Group		= iCurColumnPos(3)
			C_Work_Line			= iCurColumnPos(4)
			C_Work_Line_Desc	= iCurColumnPos(5)
			C_Remark			= iCurColumnPos(6)
    End Select    
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("PX901", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboLineGrp, lgF0, lgF1, Chr(11)) 
	
'	ggoSpread.SetCombo "1" & vbtab & "2" & vbtab & "3" , C_LINE_GROUPCD
'   ggoSpread.SetCombo "TV" & vbtab & "MN" & vbtab & "NB", C_LINE_GROUP   

    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_LINE_GROUPCD			
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_LINE_GROUP
End Sub

'Sub InitComboBox_two()
'	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("PX901", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
'	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DiFlag			'COLM_DATA_TYPE
'    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DiFlagNm
'End Sub

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

'------------------------------------------  OpenShiftPopup()  -------------------------------------------------
'	Name : OpenShiftPopup()
'	Description : ShiftPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenShiftPopup(ByVal str)
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True	
	
	arrParam(0) = "Shift�˾�"	
	arrParam(1) = "p_shift_header"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " " 
	arrParam(5) = "Shift��"			
	
    arrField(0) = "SHIFT_CD"	
    arrField(1) = "DESCRIPTION"	
       
    
    arrHeader(0) = "Shift"		
    arrHeader(1) = "Shift��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetShift(arrRet)
	End If	
	
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
	With frm1
		.vspdData.Col = C_ShiftCd
		.vspdData.Text = arrRet(0)
		
		.vspdData.Col = C_ShiftNm
		.vspdData.Text = arrRet(1)		
		
		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		' ������ �Ͼ�ٰ� �˷��� 

	End With
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
'-------------------------------------  JumpShiftException()  -----------------------------------------
'	Name : JumpShiftException()
'	Description : Shift ���ܵ������ Jump�Ѵ�.
'--------------------------------------------------------------------------------------------------------- 

Function JumpShiftException()
    Dim IntRetCd, strVal

    '-----------------------
    'Precheck area
    '-----------------------
    
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002","X","X","X")
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900017",parent.VB_YES_NO,"X","X")
        If IntRetCd = vbNo Then
			Exit Function
		End If
	End If
		
    If frm1.vspdData.ActiveRow <= 0 Then 
		Call DisplayMsgBox("181216",parent.VB_YES_NO,"X","X")
		Exit Function
	End If
	
	'-----------------------------
	' Write Cookie
	'-----------------------------	
	WriteCookie "txtPlantCd", FilterVar(UCase(Trim(frm1.txtPlantCd.value)),,"SNM")
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
'	WriteCookie "txtResourceCd", FilterVar(UCase(Trim(frm1.txtResourceCd.value)),,"SNM")
'	WriteCookie "txtResourceNm", frm1.txtResourceNm.value 
	
'	frm1.vspdData.Row = frm1.vspdData.ActiveRow
'	frm1.vspdData.Col = C_ShiftCd
	
'	WriteCookie "txtShiftCd", UCase(Trim(frm1.vspdData.Text))
	
'	frm1.vspdData.Row = frm1.vspdData.ActiveRow
'	frm1.vspdData.Col = C_ShiftNm
	
'	WriteCookie "txtShiftNm", UCase(Trim(frm1.vspdData.Text))
	
'	PgmJump(BIZ_PGM_JUMPSHIFTEXECEPTION_ID)	
	
End Function

'-------------------------------------  CntMaxRows()  -----------------------------------------
'	Name : CntMaxRows()
'	Description : �ڿ��� ��ϵ� �� �ִ� Shift������ �����Ѵ�. iPos : �߰��� Row�� 
'------------------------------------------------------------------------------------------------- 

Function CntMaxRows(iPos)
	Dim TotRowCnt
	Dim iRows
	
	On Error Resume Next
	Err.Clear
	
	CntMaxRows = False
		
	TotRowCnt = frm1.vspdData.MaxRows
	
	ggoSpread.Source = frm1.vspdData
	
	'--------------------------------------------------------------------------
	' ������ ���� ������ �� Row�� ������ ����Ѵ�.
	'--------------------------------------------------------------------------
	For iRows = 1 To TotRowCnt
		frm1.vspdData.Col = 0
		frm1.vspdData.Row = iRows
		
		If frm1.vspdData.Text = ggoSpread.DeleteFlag Then
			TotRowCnt = TotRowCnt - 1
		End If
	Next
	
	'--------------------------------------------------------------------------
	' FncInsertRow�� FncCopy�̸� ���� �߰��ϱ� ���� ���� �ϳ� ���ؼ� ����Ѵ�.
	'--------------------------------------------------------------------------
	TotRowCnt = TotRowCnt + iPos	
	
	If TotRowCnt > 4 Then
		Call DisplayMsgBox("181814","X","X","X")
		Exit Function
	End If

	CntMaxRows = True
	
End Function
'########################################################################################################
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
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetDefaultVal
    
    '----------  Coding part  -------------------------------------------------------------
   
    Call SetToolbar("11001101000011")								'��: ��ư ���� ���� 
    Call InitComboBox
        
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtResourceCd.focus
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus  
		Set gActiveElement = document.activeElement 
	End IF
	
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	'----------  Coding part  -------------------------------------------------------------   
	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_ShiftPopUp Then
        .Col = C_ShiftCd
        .Row = Row
        
        Call OpenShiftPopup(.Text)     
        Call SetActiveCell(frm1.vspdData,C_ShiftCd,Row,"M","X","X")
		Set gActiveElement = document.activeElement    
    End If
    
    End With
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
'   Event Name :vspddata_DblClick
'   Event Desc :
'==========================================================================================
Sub vspddata_DblClick(ByVal Col , ByVal Row )       
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

'==========================================================================================
'   Event Name : vspddata_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspddata_MouseDown(Button,Shift,x,y)
		
	If Button = "2" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

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

	'----------  Coding part  -------------------------------------------------------------   

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
		If lgStrPrevKeyPlantCd <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
		    
			Case  C_LINE_GROUP
				.Col = Col
				intIndex = .Value
				.Col = C_LINE_GROUPCD
				.Value = intIndex
					
				
			Case  C_LINE_GROUPCD
				.Col = Col
				intIndex = .Value
				.Col = C_LINE_GROUP
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
    ggoSpread.Source = frm1.vspdData 
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
		
'	If frm1.txtResourceCd.value = "" Then
'		frm1.txtResourceNm.value = ""
'	End If
	
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables
  
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     														'��: Query db data

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
    
    Err.Clear																	'��: Protect system from crashing
    On Error Resume Next														'��: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData 
    
    If ggoSpread.SSCheckChange = False Then 
       IntRetCD = DisplayMsgBox("900001","X","X","X")                            '��: No data changed!!
       Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    ggoSpread.Source = frm1.vspdData 
	If Not ggoSpread.SSDefaultCheck Then
		Exit Function
	End If
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		If Not chkField(Document, "1") Then
			Exit Function
		End If
    End IF
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     													'��: Save db data
    
    FncSave = True																'��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
	
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
	'Row �ϳ��� �߰��Ҷ� 4���� �ʰ��ϴ��� üũ 
'	If Not CntMaxRows(1) Then Exit Function
	
	frm1.vspdData.ReDraw = False
	
    ggoSpread.Source = frm1.vspdData	
    frm1.vspdData.EditMode = True
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow
    
	frm1.vspdData.ReDraw = True
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
	
	  On Error Resume Next
    
    Dim iDx
	
    FncCancel = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
 
    
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo 
    call InitData 
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
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
   Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
			Exit Function
		End If	
	End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement  
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
	Call InitData()
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
    
    DbQuery = False
    
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF
		
    Err.Clear                                                               '��: Protect system from crashing

	Dim strVal
    
    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboLineGrp=" & Trim(.hcboLineGrp.Value)
		strVal = strVal & "&lgStrPrevKeyPlantCd=" & lgStrPrevKeyPlantCd
		strVal = strVal & "&lgStrPrevKeyWorkLine=" & lgStrPrevKeyWorkLine
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&cboLineGrp=" & Trim(.cboLineGrp.value)
		strVal = strVal & "&lgStrPrevKeyPlantCd=" & lgStrPrevKeyPlantCd
		strVal = strVal & "&lgStrPrevKeyWorkLine=" & lgStrPrevKeyWorkLine
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
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
    
    Call InitData()
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	Call SetToolbar("11001111000111")
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
    
'    If Not CntMaxRows(0) Then Exit Function
    
    LayerShowHide(1) 
		
    'On Error Resume Next                                                   '��: Protect system from crashing
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		.txtFlgMode.value = lgIntFlgMode
    
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

            Case ggoSpread.InsertFlag												'��: �ű� 
				
				strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep					'��: C=Create
				
				
                .vspdData.Col = C_Line_GroupCd	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_Work_Line	'3
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_Work_Line_Desc  '4	
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_Remark	'5
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag

				strVal = strVal & "U" & parent.gColSep	& lRow & parent.gColSep					'��: U=Update
				
                .vspdData.Col = C_Line_GroupCd	'2
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_Work_Line_Desc	'3
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_Remark	'4
                strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_Remark	'5
                strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                                               
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag												'��: ���� 

				strDel = strDel & "D" & parent.gColSep	& lRow & parent.gColSep
				
                .vspdData.Col = C_Line_GroupCd	'2
                strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                
                .vspdData.Col = C_Work_Line	'3
                strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                                
                lGrpCnt = lGrpCnt + 1
        End Select
                
    Next

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'��: �����Ͻ� ASP �� ���� 
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������ε��(S)</font></td>
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
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="�����"></TD>
									<TD CLASS=TD5 NOWRAP>���α׷�</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboLineGrp" ALT="���α׷�" STYLE="Width: 168px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
<!--			
									<TD CLASS=TD5 NOWRAP>�ڿ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd" SIZE=15 MAXLENGTH=10 tag="12XXXU" ALT="�ڿ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResource()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm" SIZE=25 tag="14" ALT="�ڿ���"></TD>
-->									
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
									<script language =javascript src='./js/p1302ma1_ko119_I510278844_vspdData.js'></script>
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
<!--	
	<TR>
		<TD HEIGHT=3></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD>
							<TD WIDTH=* Align=right><A href="vbscript:JumpShiftException()">Shift���ܵ��</A></TD>
							<TD WIDTH=10>&nbsp;</TD>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
-->	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%>  FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hcboLineGrp" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

