<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
*  1. Module Name          : Production
*  2. Function Name        : 
*  3. Program ID           : b1b12ma1.asp
*  4. Program Name         : Lot Control
*  5. Program Desc         :
*  6. Component List       : 
*  7. Modified date(First) : 2000/04/19
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : Mr  Kim Gyoung-Don
* 10. Modifier (Last)      : Lee Hwa Jung
* 11. Comment              :
**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--#########################################################################################################
												1. �� �� �� 
##########################################################################################################-->
<!--
========================================================================================================
=                          1.1.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

						
<!--==========================================  1.1.1 Style Sheet  ======================================
==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--==========================================  1.1.2 ���� Include   ======================================
==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Dim BaseDate
Dim StartDate

'========================================================================================================
'=                       1.2.1 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)

Const BIZ_PGM_QRY_ID = "b1b12mb1.asp"											
Const BIZ_PGM_SAVE_ID = "b1b12mb2.asp"											
Const BIZ_PGM_DEL_ID = "b1b12mb3.asp"											
Const BIZ_PGM_JUMPITEMBYPLANT_ID = "b1b11ma1"

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgNextNo						'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo						' ""
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  --------------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  +++++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop          
Dim  lgRdoOldVal1

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

    lgIntFlgMode = parent.OPMD_CMODE                                               
    lgBlnFlgChgValue = False                                                
    '----------  Coding part  -----------------------------------------------------------------
    IsOpenPop = False														
    
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===============================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
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
	frm1.cboLotType.value = "A"
	frm1.rdoValidPerdFlg2.checked = True 
	lgRdoOldVal1 = 2
	
	Call ggoOper.SetReqAttr(frm1.txtValidPerd,"Q")
	frm1.txtValidPerd.Text = 0
	frm1.txtLotInc.Value = 1
	
	frm1.txtValidFromDt.text  = StartDate
	frm1.txtValidToDt.text	   = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
End Sub

Sub InitComboBox()
	Call SetCombo(frm1.cboLotType,"A","�ڵ�")
	Call SetCombo(frm1.cboLotType,"M","����") 
End Sub

'******************************************  2.3 Operation ó���Լ�  *************************************
'	���: Operation ó���κ� 
'	����: Tabó��, Reference���� ���Ѵ�. 
'*********************************************************************************************************

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
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenPlant()  ------------------------------------------------
'	Name : OpenPlant()
'	Description :  Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

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

'------------------------------------------  OpenItemCd()  -----------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd(ByVal lIndex)

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	If lIndex = 0 Then
		If UCase(frm1.txtItemcd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	Else
		If UCase(frm1.txtItemCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	
	
	If lIndex = 0 Then
		arrParam(1) = Trim(frm1.txtItemCd.value)
	Else
		arrParam(1) = Trim(frm1.txtItemCd1.value)	
	End If
	
	arrParam(2) = ""							
	arrParam(3) = ""	
	arrParam(4) = ""
	If lIndex = 0 Then
		arrParam(5) = " AND A.LOT_FLG = " & FilterVar("Y", "''", "S")			
	End If
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet,lIndex)
	End If
	
	Call SetFocusToDocument("M")
	If lIndex = 0 Then
		frm1.txtItemCd.focus
	Else
		frm1.txtItemCd1.focus
	End If
		
End Function

'------------------------------------------  SetPlantCd()  -----------------------------------------------
'	Name : SetPlantCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemCd()  ------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(byval arrRet, ByVal lIndex)
	If lIndex = 0 Then
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
	Else
		frm1.txtItemCd1.Value    = arrRet(0)		
		frm1.txtItemNm1.Value    = arrRet(1)
		lgBlnFlgChgValue = True   			
	End If		
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetCookieVal()  ---------------------------------------------
'	Name : SetCookieVal()
'	Description : Cookie Setting
'---------------------------------------------------------------------------------------------------------

Sub SetCookieVal()
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value = ReadCookie("txtPlantNm")
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
		frm1.txtItemNm.value = ReadCookie("txtItemNm") 
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm",""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm",""

End Sub

'------------------------------------------  SetItemCd()  ------------------------------------------------
'	Name : JumpItemByPlant()
'	Description : Item by Plant�� Jump�Ѵ�.
'---------------------------------------------------------------------------------------------------------

Function JumpItemByPlant()

	WriteCookie "txtPlantCd", Trim(frm1.txtPlantCd.value)
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value  
	WriteCookie "txtItemCd", Trim(frm1.txtItemCd.value)
	WriteCookie "txtItemNm", frm1.txtItemNm.value 
	WriteCookie "MainFormFlg", "LOT"
	
	PgmJump(BIZ_PGM_JUMPITEMBYPLANT_ID)
	
End Function
 
'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'#########################################################################################################

'******************************************  3.1 Window ó��  ********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ==========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029																'��: Load table , B_numeric_format
	Call AppendNumberPlace("6","3","0")
	Call AppendNumberPlace("7","5","0")
	Call AppendNumberRange("6","0","100")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("11101000000011")
    Call InitComboBox
    Call SetDefaultVal
    Call InitVariables
    Call SetCookieVal
    
    If frm1.txtPlantCd.value <> "" and frm1.txtItemCd.value <> "" Then

        If DbQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
		frm1.cboLotType.focus
		Set gActiveElement = document.activeElement
	ElseIf frm1.txtPlantCd.value = "" Then
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
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


'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'#########################################################################################################

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtValidFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidFromDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtValidToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidToDt.Focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidToDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtLotInc_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtValidperd_Change()
    lgBlnFlgChgValue = True
End Sub

Sub cboLotType_OnChange()
	If frm1.cboLotType.value = "M" Then
		Call ggoOper.SetReqAttr(frm1.txtLotStartChar,"Q")
		Call ggoOper.SetReqAttr(frm1.txtLotInc,"Q")
		frm1.txtLotInc.value = "0"
		frm1.txtLotStartChar.value = "" 
	Else
		Call ggoOper.SetReqAttr(frm1.txtLotStartChar,"D")
		Call ggoOper.SetReqAttr(frm1.txtLotInc,"N")
		frm1.txtLotInc.Value = "1"
	End If 
    lgBlnFlgChgValue = True
End Sub
		    
Sub rdoValidPerdFlg1_OnClick()
	If lgRdoOldVal1 = 1 Then Exit Sub
	
	lgBlnFlgChgValue = True
	lgRdoOldVal1 = 1
	
	Call ggoOper.SetReqAttr(frm1.txtValidPerd,"N")
	
	frm1.txtValidPerd.Text = 1
	  
End Sub

Sub rdoValidPerdFlg2_OnClick()
	If lgRdoOldVal1 = 2 Then Exit Sub
	
	lgBlnFlgChgValue = True
	lgRdoOldVal1 = 2 
	
	Call ggoOper.SetReqAttr(frm1.txtValidPerd,"Q")
	
	frm1.txtValidPerd.Text = 0
    
End Sub

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
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function ********************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                       
    
    Err.Clear                                                              

	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")				
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
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
	
    Call ggoOper.ClearField(Document, "2")										
    Call SetDefaultVal
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

    If DbQuery = False Then   
		Exit Function           
    End If 
           
    FncQuery = True		
    														
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False																
    
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")					
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
    frm1.txtItemCd.value = ""
    frm1.txtItemNm.value = ""

	Call SetToolbar("11101000000011")
	
    Call ggoOper.ClearField(Document, "2")                                      
    Call ggoOper.LockField(Document, "N")                                       
    
    Call SetDefaultVal
    Call InitVariables															
    
    frm1.txtItemCd1.focus
    Set gActiveElement = document.activeElement 

    FncNew = True																

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCD
    
    FncDelete = False														
    
	'-----------------------
    'Precheck area
    '-----------------------

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     
        Call DisplayMsgBox("900002","X","X","X")                                
        Exit Function
    End If
    
	'-----------------------
    'Delete function call area
    '-----------------------

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            
	If IntRetCD = vbNo Then													
		Exit Function	
	End If
	
    If DbDelete = False Then 
		Exit Function           
    End If 
    
    FncDelete = True  

End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               
    
	'-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                         
        Exit Function
    End If
    
	'-----------------------
    'Check content area
    '-----------------------

    If Not chkField(Document, "2") Then                             
       Exit Function
    End If
    
	'-----------------------
    'Save function call area
    '-----------------------

    If DbSave = False Then   
		Exit Function           
    End If 
    
    FncSave = True                                                          
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE												
    Call SetToolbar("11101000000011")
    Call ggoOper.LockField(Document, "N")									
    
    ' ���Ǻ� �ʵ带 �����Ѵ�.
	frm1.txtItemCd.value = ""
	frm1.txtItemNm.value = ""
	frm1.txtItemCd1.value = ""
	frm1.txtItemNm1.value = ""
	frm1.txtValidFromDt.text = StartDate
	frm1.txtValidToDt.text	 = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
    
    Call cboLotType_OnChange	'2003-09-17
    
    frm1.txtItemCd1.focus
    Set gActiveElement = document.activeElement 
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
    Call parent.fncPrint()                                                  
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    Dim strVal
    Dim	IntRetCD
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002","X","X","X")                            
        Exit Function
    End If
    
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")				
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
    '------------------------------------
    'Data Sheet �ʱ�ȭ 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")									
    
    Call SetDefaultVal
    Call InitVariables														
    
    Err.Clear                                                               

    LayerShowHide(1)								
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		
	strVal = strVal & "&txtUpdtUserId=" & parent.gUsrID
	strVal = strVal & "&PrevNextFlg=" & "P"
	    
	Call RunMyBizASP(MyBizASP, strVal)									

End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    Dim strVal
	Dim	IntRetCD
	
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002","X","X","X")                            
        Exit Function
    End If
    
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")				
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
    '------------------------------------
    'Data Sheet �ʱ�ȭ 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")									
    
    Call SetDefaultVal
    Call InitVariables														

    Err.Clear                                                               

	LayerShowHide(1)								
    
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		
	strVal = strVal & "&txtUpdtUserId=" & parent.gUsrID
	strVal = strVal & "&PrevNextFlg=" & "N"
	
	Call RunMyBizASP(MyBizASP, strVal)									

End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)												
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)						'��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  ******************************
'	���� : 
'*********************************************************************************************************

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
    Err.Clear                                                               
    
    DbDelete = False														
    
    LayerShowHide(1)									
    
    Dim strVal
    
    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003					
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		
    
	Call RunMyBizASP(MyBizASP, strVal)	

    DbDelete = True      
                                                   
	
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()	

	Call InitVariables()
	Call FncNew()
	
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    
    Err.Clear                                                             
    
    DbQuery = False                                                       
    
    LayerShowHide(1)								
   
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		
	strVal = strVal & "&txtUpdtUserId=" & parent.gUsrID
	strVal = strVal & "&PrevNextFlg=" & ""
	    
	Call RunMyBizASP(MyBizASP, strVal)									
	
    DbQuery = True                                                      

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												

    Call ggoOper.LockField(Document, "Q")									

    Call SetToolbar("11111000111111")

	If frm1.rdoValidPerdFlg1.checked = True then
		lgRdoOldVal1 = 1
		Call ggoOper.SetReqAttr(frm1.txtValidPerd,"N")
	Else
		lgRdoOldVal1 = 1
		Call rdoValidPerdFlg2_OnClick()
	End If
	
	If frm1.cboLotType.value = "M" Then
		Call ggoOper.SetReqAttr(frm1.txtLotStartChar,"Q")
		Call ggoOper.SetReqAttr(frm1.txtLotInc,"Q")
	End If
	
	frm1.cboLotType.focus 
	Set gActiveElement = document.activeElement 
	
	lgBlnFlgChgValue = False
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
Function DbSave() 

    Err.Clear																

	DbSave = False															
	
	If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function       

	If frm1.rdoValidPerdFlg1.checked = True Then
		If UNIConvNum(frm1.txtValidPerd.Text, 0) < 1 Then
			Call DisplayMsgBox("970022","X", "ǰ����ȿ�Ⱓ","0")
			frm1.txtValidPerd.focus
			Set gActiveElement = document.activeElement  
			Exit Function
		End If
	End If
	
	If frm1.cboLotType.value = "A" Then
		If UNIConvNum(frm1.txtLotInc.value, 0) < 1 Then
			Call DisplayMsgBox("970022","X", "Lot ����","0")
			frm1.txtLotInc.focus
			Set gActiveElement = document.activeElement  
			Exit Function
		End If
	End If
			 
	LayerShowHide(1)								
    
    Dim strVal

	With frm1
		.txtMode.value = parent.UID_M0002										
		.txtFlgMode.value = lgIntFlgMode

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	
	End With
	
    DbSave = True                                                       
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()															

    
    Call InitVariables
    
    frm1.txtItemCd.value = frm1.txtItemCd1.value 
    
    Call MainQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��Ʈ ����</font></td>
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
									<TD CLASS=TD656 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=50 tag="14"></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=25 MAXLENGTH=18 tag="12XXXU"  ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd 0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=50 tag="14"></TD>
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
								<TD CLASS=TD5 NOWRAP>ǰ��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=25 MAXLENGTH=18 tag="23XXXU"  ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd 1">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=50 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Lot �ο����</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboLotType" ALT="Lot �ο����" STYLE="Width: 115px;" tag="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ֽ� Lot No.</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNewLotNo" SIZE=25 MAXLENGTH=25 tag="24X7" ALT="�ֽ� Lot No."></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Lot ���۹���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLotStartChar" SIZE=10 MAXLENGTH=5 tag="21XXXU" ALT="Lot ���۹���"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Lot ����</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS90 name=txtLotInc CLASSID=<%=gCLSIDFPDS%> tag="22X66" ALT="Lot ����" MAXLENGTH="10" SIZE="10"> </OBJECT>');</SCRIPT>												
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ǰ����ȿ�Ⱓ��������</TD>
								<TD CLASS=TD6 NOWRAP>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidPerdFlg" tag="2X" ID="rdoValidPerdFlg1" VALUE="Y"><LABEL FOR="rdoValidPerdFlg1">��</LABEL>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidPerdFlg" tag="2X" CHECKED ID="rdoValidPerdFlg2" VALUE="N"><LABEL FOR="rdoValidPerdFlg2">�ƴϿ�</LABEL></TD>													
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ǰ����ȿ�Ⱓ</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLPADDING=0 CELLSPACING=0>
										<TR>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS90 name=txtValidPerd CLASSID=<%=gCLSIDFPDS%> SIZE="10" MAXLENGTH=5 ALT="ǰ����ȿ�Ⱓ" tag="24X7Z"></OBJECT>');</SCRIPT>
											</TD>
											<TD valign=bottom>
												&nbsp;��
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��ȿ�Ⱓ</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtValidFromDt CLASSID=<%=gCLSIDFPDT%> ALT="��ȿ�Ⱓ������" tag="23X1"> </OBJECT>');</SCRIPT>&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtValidToDt CLASSID=<%=gCLSIDFPDT%> ALT="��ȿ�Ⱓ������" tag="22X1"> </OBJECT>');</SCRIPT>										
								</TD>
							</TR>
							<% Call SubFillRemBodyTD656(11)%>
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
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpItemByPlant">���庰 ǰ����</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
