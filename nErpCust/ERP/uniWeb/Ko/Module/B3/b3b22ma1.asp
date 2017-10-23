<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b22ma1.asp
'*  4. Program Name         : Entry Class
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/03
'*  8. Modified date(Last)  :  
'*  9. Modifier (First)     : Lee Woo Guen
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  <!-- '☆: 해당 위치에 따라 달라짐, 상대 경로 -->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "b3b22mb1.asp"
Const BIZ_PGM_SAVE_ID = "b3b22mb2.asp"
Const BIZ_PGM_DEL_ID = "b3b22mb3.asp"
Const BIZ_PGM_LOOKUP_CHAR_ID = "b3b22mb4.asp"

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop
Dim lgRdoOldVal1
Dim lgRdoOldVal2
Dim lgRdoOldVal3					

Dim lgNextNo						'☜: 화면이 Single/SingleMulti 인경우만 해당 
Dim lgPrevNo

Dim blnFlgSetValue1
Dim blnFlgSetValue2
Dim blnFlgIsUsedByItem
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    IsOpenPop = False		
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA")%>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()

End Sub

Sub InitComboBox()
    On Error Resume Next
    Err.Clear

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1010", "''", "S") & "  ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboClassMgr, "" & Chr(11) & lgF0, "" & Chr(11) & lgF1, Chr(11))
End Sub

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
	
    arrField(0) = 1 							' Field명(0) : "Class_CD"
    arrField(1) = 2 							' Field명(1) : "Class_NM"
	
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

'------------------------------------------  OpenCharCd1()  -------------------------------------------------
'	Name : OpenCharCd1()
'	Description : Characteristic PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCharCd1()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtCharCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtCharCd1.value)	' Characteristic Code
	arrParam(1) = ""							' Characteristic Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field명(0) : "Characteristic_CD"
    arrField(1) = 2 							' Field명(1) : "Characteristic_NM"
    arrField(2) = 3 							' Field명(1) : "Char_Value_Digit"
    
	iCalledAspName = AskPRAspName("B3B30PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B30PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=600px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
    
	If arrRet(0) <> "" Then
		Call SetCharCd1(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCharCd1.focus
	
End Function

'------------------------------------------  OpenCharCd2()  -------------------------------------------------
'	Name : OpenCharCd2()
'	Description : Characteristic PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCharCd2()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtCharCd2.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtCharCd2.value)	' Characteristic Code
	arrParam(1) = ""							' Characteristic Name
	arrParam(2) = ""							' ----------
	arrParam(3) = ""							' ----------
	arrParam(4) = ""
	
    arrField(0) = 1 							' Field명(0) : "Characteristic_CD"
    arrField(1) = 2 							' Field명(1) : "Characteristic_NM"
    arrField(2) = 3 							' Field명(1) : "Char_Value_Digit"
    
	iCalledAspName = AskPRAspName("B3B30PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B3B30PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=600px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCharCd2(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCharCd2.focus
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetClassCd()  ------------------------------------------------
'	Name : SetClassCd()
'	Description : Class Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetClassCd(byval arrRet)
	frm1.txtClassCd.Value    = arrRet(0)		
	frm1.txtClassNm.Value    = arrRet(1)
	
	frm1.txtClassCd.focus
	Set gActiveElement = document.activeElement 
End Function

'------------------------------------------  SetCharCd1()  --------------------------------------------------
'	Name : SetCharCd1()
'	Description : Characteristic Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCharCd1(byval arrRet)
	frm1.txtCharCd1.Value			= arrRet(0)	
	frm1.txtCharNm1.Value			= arrRet(1)
	frm1.txtCharValueDigit1.Value   = arrRet(2)
	
	frm1.txtCharCd1.focus
	Set gActiveElement = document.activeElement
	
	blnFlgSetValue1 = True		'onChange, onBlur Event에서 사용 
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetCharCd2()  --------------------------------------------------
'	Name : SetCharCd2()
'	Description : Characteristic Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCharCd2(byval arrRet)
	frm1.txtCharCd2.Value	= arrRet(0)	
	frm1.txtCharNm2.Value   = arrRet(1)
	frm1.txtCharValueDigit2.Value   = arrRet(2)
	
	frm1.txtCharCd2.focus
	Set gActiveElement = document.activeElement
	
	blnFlgSetValue2 = True		'onChange, onBlur Event에서 사용 
	lgBlnFlgChgValue = True
End Function

Sub SetCookieVal()
	If ReadCookie("txtClassCd") <> "" Then
		frm1.txtClassCd.Value = ReadCookie("txtClassCd")
		frm1.txtClassNm.value = ReadCookie("txtClassNm")
	End If	

	WriteCookie "txtClassCd", ""
	WriteCookie "txtClassNm", ""
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'========================================================================================
' Sub Name : Set Class_Digit
' Sub Desc : 
'========================================================================================
Sub SetClassDigit()
	Dim intClassDigit
	Dim intCharValueDigit1
	Dim intCharValueDigit2

	If Trim(frm1.txtCharCd1.value) <> "" Then
		intCharValueDigit1 = frm1.txtCharValueDigit1.value
	Else
		intCharValueDigit1 = 0
	End If
	
	If Trim(frm1.txtCharCd2.value) <> "" Then
		intCharValueDigit2 = frm1.txtCharValueDigit2.value
	Else
		intCharValueDigit2 = 0
	End If
	
	If intCharValueDigit2 <> 0 Then
		intClassDigit = 16 - intCharValueDigit1 - intCharValueDigit2
	Else
		intClassDigit = 17 - intCharValueDigit1
	End If

	If intClassDigit < 0 Then
		intClassDigit = 0
	End If

	frm1.txtClassDigit.value = intClassDigit
End Sub

'========================================================================================
' Sub Name : Set Class_Digit
' Sub Desc : 
'========================================================================================
Sub LookUpChar(ByVal charChoice)
	Dim strVal
	Dim strCharCd

    Err.Clear                                                       
	
	If   LayerShowHide(1) = False Then Exit Sub
	
	If charChoice = "1" Then
		strCharCd = Trim(frm1.txtCharCd1.value)
	Else
		strCharCd = Trim(frm1.txtCharCd2.value)
	End If

	strVal = BIZ_PGM_LOOKUP_CHAR_ID & "?txtMode=" & parent.UID_M0001
	strVal = strVal & "&txtCharCd=" & strCharCd
	strVal = strVal & "&charChoice=" & charChoice

	Call RunMyBizASP(MyBizASP, strVal)
End Sub

'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************

'==========================================  3.1.1 Form_Load()  ==========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029

	Call AppendNumberPlace("7","3","2")
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
                       
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field

    '----------  Coding part  -------------------------------------------------------------
    Call SetCookieVal
    
    Call SetToolbar("11101000000011")
    Call InitComboBox
    Call SetDefaultVal    
	Call InitVariables																'⊙: Initializes local global variables
	frm1.txtClassCd.focus
	Set gActiveElement = document.activeElement
End Sub

'==========================================================================================
'   Event Name : cboClassMgr_onChange()
'   Event Desc :
'==========================================================================================
Sub cboClassMgr_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtCharCd1_onChange()
	If Trim(frm1.txtCharCd1.value) = "" Then
		frm1.txtCharValueDigit1.value = 0
		Call SetClassDigit
	Else
		Call LookUpChar("1")
	End If	
End Sub
Sub txtCharCd2_onChange()
	If Trim(frm1.txtCharCd2.value) = "" Then
		frm1.txtCharValueDigit2.value = 0
		Call SetClassDigit
	Else
		Call LookUpChar("2")
	End If	
End Sub

Sub txtCharCd1_onFocus()
	If blnFlgSetValue1 Then
		If Trim(frm1.txtCharCd1.value) = "" Then
			frm1.txtCharValueDigit1.value = 0
		End If	
		Call SetClassDigit
	End If
	blnFlgSetValue1 = false
End Sub
Sub txtCharCd2_onFocus()
	If blnFlgSetValue2 Then
		If Trim(frm1.txtCharCd2.value) = "" Then
			frm1.txtCharValueDigit2.value = 0
		End If	
		Call SetClassDigit
	End If
	blnFlgSetValue2 = false
End Sub
'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    FncQuery = False                                                        

	'-----------------------
    'Check previous data area
    '----------------------- 

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				
		If IntRetCD = vbNo Then
			Exit Function
		End If
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
    
	'-----------------------
    'Check previous data area
    '-----------------------

    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")	           
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    
    Call ggoOper.ClearField(Document, "A")                                      '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field
    
    Call SetToolbar("11101000000011")
    Call SetDefaultVal    
    Call InitVariables															'⊙: Initializes local global variables
    frm1.txtClassCd1.focus
    Set gActiveElement = document.activeElement
      
    FncNew = True																'⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim intRetCd

    FncDelete = False												
    
	'-----------------------
    'Precheck area
    '-----------------------

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If

	'-----------------------
    'Delete function call area
    '-----------------------%>
    IntRetCd = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		           
	If IntRetCd = vbNo Then
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

	'-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        Call DisplayMsgBox("900001", "X", "X", "X") 
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
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")				
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE
    
    ' 조건부 필드를 삭제한다.
    Call ggoOper.LockField(Document, "N")
    Call SetToolbar("11101000000011")
    
    frm1.txtClassCd1.value = ""
    frm1.txtClassCd1.focus

    Set gActiveElement = document.activeElement  

    lgBlnFlgChgValue = True
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
    Dim strVal
    Dim IntRetCD
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")                            
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")					
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
    strVal = strVal & "&txtClassCd=" & Trim(frm1.txtClassCd.value)
	strVal = strVal & "&PrevNextFlg=" & "P"
	
	Call RunMyBizASP(MyBizASP, strVal)										
	
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    Dim strVal
	Dim IntRetCD
	
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")                            
        Exit Function
    End If
    
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")				
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
    '------------------------------------
    'Data Sheet 초기화 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")									
    
    Call SetDefaultVal
    Err.Clear         
    
    LayerShowHide(1)
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						
    strVal = strVal & "&txtClassCd=" & Trim(frm1.txtClassCd.value)			
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
    Call parent.FncFind(parent.C_SINGLE, False)                                         
End Function
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function
'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
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
    strVal = strVal & "&txtClassCd1=" & Trim(frm1.txtClassCd1.value)		

	Call RunMyBizASP(MyBizASP, strVal)									
	
    DbDelete = True                                                     

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()														
	Call InitVariables
	Call FncNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    DbQuery = False                                                         
    
    LayerShowHide(1)							
    
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						
    strVal = strVal & "&txtClassCd=" & Trim(frm1.txtClassCd.value)			
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
    dim LayerN1
	frm1.hClassCd.value = frm1.txtClassCd.value
    
	Set LayerN1 = window.document.all("MousePT").style
	
    lgIntFlgMode = parent.OPMD_UMODE											
    lgBlnFlgChgValue = false
	frm1.txtClassNm1.focus 
	Set gActiveElement = document.activeElement 
    Call ggoOper.LockField(Document, "Q")

    If blnFlgIsUsedByItem Then
		Call ggoOper.SetReqAttr(frm1.txtCharCd1, "Q")
		Call ggoOper.SetReqAttr(frm1.txtCharCd2, "Q")
		Call SetToolbar("11101000111111")
	Else
		Call ggoOper.SetReqAttr(frm1.txtCharCd1, "N")
		Call ggoOper.SetReqAttr(frm1.txtCharCd2, "D")
		Call SetToolbar("11111000111111")
	End If
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 

    Err.Clear																

	DbSave = False															

	LayerShowHide(1)
		
	With frm1
		.txtMode.value = parent.UID_M0002										
		.txtFlgMode.value = lgIntFlgMode
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
    
    DbSave = True    
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()															

    dim LayerN1
   
	Set LayerN1 = window.document.all("MousePT").style
	
    frm1.txtClassCd.value = frm1.txtClassCd1.value 
    Call InitVariables
    
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>클래스등록</font></td>
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
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD HEIGHT=5 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE CLASS="BasicTB" CELLSPACING=0>
								<TR>
									<TD CLASS=TD5 NOWRAP>클래스</TD>
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtClassCd" SIZE=20 MAXLENGTH=16 tag="12XXXU"  ALT="클래스"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnClassCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenClassCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtClassNm" SIZE=40 tag="14"></TD>
								</TR>	
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=2 WIDTH=100%></TD>
				</TR>				
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE CLASS="TB2" CELLSPACING=0>
							<TR>
								<TD WIDTH=100%  valign=top>
									<FIELDSET>
										<LEGEND>일반정보</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>클래스</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtClassCd1" SIZE=20 MAXLENGTH=16 tag="23XXXU"  ALT="클래스"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>클래스명</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtClassNm1" SIZE=40 MAXLENGTH=40 tag="22" ALT="클래스명"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>클래스자리수</TD>
												<TD CLASS=TD656 NOWRAP>														
													<script language =javascript src='./js/b3b22ma1_I276539456_txtClassDigit.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>사양항목1</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtCharCd1" SIZE=20 MAXLENGTH=18 tag="22XXXU" ALT="사양항목1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCharCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCharCd1()">&nbsp;<INPUT TYPE=TEXT NAME="txtCharNm1" SIZE=40 tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>사양값자리수1</TD>
												<TD CLASS=TD656 NOWRAP>
													<script language =javascript src='./js/b3b22ma1_I287442260_txtCharValueDigit1.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>사양항목2</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtCharCd2" SIZE=20 MAXLENGTH=18 tag="21XXXU" ALT="사양항목2"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCharCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCharCd2()">&nbsp;<INPUT TYPE=TEXT NAME="txtCharNm2" SIZE=40 tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>사양값자리수2</TD>
												<TD CLASS=TD6 NOWRAP>															
													<script language =javascript src='./js/b3b22ma1_I372232627_txtCharValueDigit2.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>클래스담당자</TD>
												<TD CLASS=TD656 NOWRAP><SELECT NAME="cboClassMgr" ALT="클래스담당자" STYLE="Width: 140px;" tag="21"></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
												<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
							</TR>	
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="hClassCd" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
