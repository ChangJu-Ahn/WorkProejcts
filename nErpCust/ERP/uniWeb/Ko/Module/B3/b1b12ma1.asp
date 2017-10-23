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
												1. 선 언 부 
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

<!--==========================================  1.1.2 공통 Include   ======================================
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

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

BaseDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(BaseDate, parent.gServerDateFormat, parent.gDateFormat)

Const BIZ_PGM_QRY_ID = "b1b12mb1.asp"											
Const BIZ_PGM_SAVE_ID = "b1b12mb2.asp"											
Const BIZ_PGM_DEL_ID = "b1b12mb3.asp"											
Const BIZ_PGM_JUMPITEMBYPLANT_ID = "b1b11ma1"

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo						' ""
'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  --------------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  +++++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop          
Dim  lgRdoOldVal1

'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
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

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
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
	Call SetCombo(frm1.cboLotType,"A","자동")
	Call SetCombo(frm1.cboLotType,"M","수동") 
End Sub

'******************************************  2.3 Operation 처리함수  *************************************
'	기능: Operation 처리부분 
'	설명: Tab처리, Reference등을 행한다. 
'*********************************************************************************************************

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'*********************************************************************************************************

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
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

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"
    
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
		Call DisplayMsgBox("971012","X", "공장","X")
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
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    
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
'	Description : Item Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemCd()  ------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup에서 Return되는 값 setting
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

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
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
'	Description : Item by Plant로 Jump한다.
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
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'#########################################################################################################

'******************************************  3.1 Window 처리  ********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ==========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6","3","0")
	Call AppendNumberPlace("7","5","0")
	Call AppendNumberRange("6","0","100")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
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
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'#########################################################################################################

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
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
'   Event Desc : 달력을 호출한다.
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
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function ********************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
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
    
    ' 조건부 필드를 삭제한다.
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
    'Data Sheet 초기화 
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
    'Data Sheet 초기화 
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
    Call parent.FncFind(parent.C_SINGLE, False)						'☜:화면 유형, Tab 유무 
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

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  ******************************
'	설명 : 
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
' Function Desc : DbDelete가 성공적일때 수행 
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
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
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
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 

    Err.Clear																

	DbSave = False															
	
	If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function       

	If frm1.rdoValidPerdFlg1.checked = True Then
		If UNIConvNum(frm1.txtValidPerd.Text, 0) < 1 Then
			Call DisplayMsgBox("970022","X", "품목유효기간","0")
			frm1.txtValidPerd.focus
			Set gActiveElement = document.activeElement  
			Exit Function
		End If
	End If
	
	If frm1.cboLotType.value = "A" Then
		If UNIConvNum(frm1.txtLotInc.value, 0) < 1 Then
			Call DisplayMsgBox("970022","X", "Lot 증분","0")
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
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>로트 관리</font></td>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD656 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=50 tag="14"></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=25 MAXLENGTH=18 tag="12XXXU"  ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd 0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=50 tag="14"></TD>
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
								<TD CLASS=TD5 NOWRAP>품목</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=25 MAXLENGTH=18 tag="23XXXU"  ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd 1">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=50 tag="24"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Lot 부여방법</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboLotType" ALT="Lot 부여방법" STYLE="Width: 115px;" tag="22"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>최신 Lot No.</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtNewLotNo" SIZE=25 MAXLENGTH=25 tag="24X7" ALT="최신 Lot No."></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Lot 시작문자</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLotStartChar" SIZE=10 MAXLENGTH=5 tag="21XXXU" ALT="Lot 시작문자"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Lot 증분</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS90 name=txtLotInc CLASSID=<%=gCLSIDFPDS%> tag="22X66" ALT="Lot 증분" MAXLENGTH="10" SIZE="10"> </OBJECT>');</SCRIPT>												
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>품목유효기간관리여부</TD>
								<TD CLASS=TD6 NOWRAP>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidPerdFlg" tag="2X" ID="rdoValidPerdFlg1" VALUE="Y"><LABEL FOR="rdoValidPerdFlg1">예</LABEL>
											<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoValidPerdFlg" tag="2X" CHECKED ID="rdoValidPerdFlg2" VALUE="N"><LABEL FOR="rdoValidPerdFlg2">아니오</LABEL></TD>													
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>품목유효기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLPADDING=0 CELLSPACING=0>
										<TR>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS90 name=txtValidPerd CLASSID=<%=gCLSIDFPDS%> SIZE="10" MAXLENGTH=5 ALT="품목유효기간" tag="24X7Z"></OBJECT>');</SCRIPT>
											</TD>
											<TD valign=bottom>
												&nbsp;일
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>유효기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtValidFromDt CLASSID=<%=gCLSIDFPDT%> ALT="유효기간시작일" tag="23X1"> </OBJECT>');</SCRIPT>&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtValidToDt CLASSID=<%=gCLSIDFPDT%> ALT="유효기간종료일" tag="22X1"> </OBJECT>');</SCRIPT>										
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
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpItemByPlant">공장별 품목등록</A></TD>
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
