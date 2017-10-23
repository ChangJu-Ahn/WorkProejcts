
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : WorkCenter
'*  3. Program ID           : p1301ma1.asp
'*  4. Program Name         : 작업장등록 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/03/10
'*  8. Modified date(Last)  : 2002/11/15
'*  9. Modifier (First)     : Mr  Kim Gyoung-Don
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "p1301mb1.asp"
Const BIZ_PGM_SAVE_ID = "p1301mb2.asp"
Const BIZ_PGM_DEL_ID = "p1301mb3.asp"

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop 
Dim lgClnrType      

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)


'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                               
    lgBlnFlgChgValue = False                                                
    '----------  Coding part  -------------------------------------------------------------
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
'============================================================================================================
Sub SetDefaultVal()

	frm1.txtValidFromDt.text  = StartDate
	frm1.txtValidToDt.text	   = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
	
	frm1.cboInsideFlg.value = "Y"
End Sub

Sub InitComboBox()
	Call SetCombo(frm1.cboInsideFlg, "Y", "사내")								'⊙: InitCombo 에서 해야 되는데 임시로 넣은 것임 
    Call SetCombo(frm1.cboInsideFlg, "N", "외주")

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1013", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboWCMgr, lgF0, lgF1, Chr(11))
End Sub

'------------------------------------------  OpenCondPlant()  --------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						
	arrParam(1) = "B_PLANT"								
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			
	arrParam(3) = ""									
	arrParam(4) = ""									
	arrParam(5) = "공장"							
	
    arrField(0) = "PLANT_CD"							
    arrField(1) = "PLANT_NM"							
    arrField(2) = "CAL_TYPE"							
    
    arrHeader(0) = "공장"							
    arrHeader(1) = "공장명"							
    arrHeader(2) = "칼렌다 타입"						
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenConWC()  ------------------------------------------------
'	Name : OpenConWC()
'	Description : Condition Work Center PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConWC()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "작업장팝업"					
	arrParam(1) = "P_WORK_CENTER"					
	arrParam(2) = Trim(frm1.txtConWcCd.Value)		
	arrParam(3) = ""								
	arrParam(4) = "P_WORK_CENTER.PLANT_CD =" & FilterVar(frm1.txtPlantCd.value, "''", "S")
	arrParam(5) = "작업장"						
	
    arrField(0) = "WC_CD"							
    arrField(1) = "WC_NM"							
    arrField(2) = "CASE WHEN INSIDE_FLG=" & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("사내", "''", "S") & " ELSE " & FilterVar("외주", "''", "S") & " END"
    
    arrHeader(0) = "작업장"						
    arrHeader(1) = "작업장명"					
    arrHeader(2) = "작업장구분"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConWC(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtConWcCd.focus
	
End Function

'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name : OpenCalType()
'	Description : Calendar Type Popup
'---------------------------------------------------------------------------------------------------------
Function OpenCalType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "칼렌다 타입 팝업"			
	arrParam(1) = "P_MFG_CALENDAR_TYPE"				
	arrParam(2) = Trim(frm1.txtClnrType.Value)		
	arrParam(3) = ""								
	arrParam(4) = ""								
	arrParam(5) = "칼렌다 타입"					
	
    arrField(0) = "CAL_TYPE"						
    arrField(1) = "CAL_TYPE_NM"					
    
    arrHeader(0) = "칼렌다 타입"				
    arrHeader(1) = "칼렌다 타입명"				

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetCalType(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtClnrType.focus
    
End Function

'------------------------------------------  OpenCostCtr()  ----------------------------------------------
'	Name : OpenCostCtr()
'	Description : Cost Center Popup
'---------------------------------------------------------------------------------------------------------
Function OpenCostCtr()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  
		Exit Function
	End If

	IsOpenPop = True 

	arrParam(0) = "Cost Center 팝업"			' 팝업 명칭 
	arrParam(1) = "B_COST_CENTER"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCostCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "B_COST_CENTER.PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & _
				" AND B_COST_CENTER.COST_TYPE =" & FilterVar("M", "''", "S") & " " & _
				" AND B_COST_CENTER.DI_FG =" & FilterVar("D", "''", "S") & " "			' Where Condition
	arrParam(5) = "Cost Center"					' TextBox 명칭 
	
    arrField(0) = "COST_CD"							' Field명(0)
    arrField(1) = "COST_NM"							' Field명(1)
    
    arrHeader(0) = "Cost Center"				' Header명(0)
    arrHeader(1) = "Cost Center 명"				' Header명(1)

	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCostCtr(arrRet)
	End If	
    
End Function

'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtClnrType.value   = arrRet(2)
	lgClnrType	= arrRet(2)		
End Function

'------------------------------------------  SetConWC()  --------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConWC(byval arrRet)
	frm1.txtConWcCd.Value    = arrRet(0)		
	frm1.txtConWcNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetCalType()  -----------------------------------------------
'	Name : SetCalType()
'	Description : Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCalType(byval arrRet)
	frm1.txtClnrType.value = arrRet(0)
	frm1.txtClnrTypeNm.value = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetCostCtr()  -----------------------------------------------
'	Name : SetCostCtr()
'	Description : Cost Center Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCostCtr(byval arrRet)
	frm1.txtCostCd.value = arrRet(0)
	frm1.txtCostNm.value = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  LookUp()  ---------------------------------------------------
'	Name : LookUp()
'	Description : 현재 공장에 해당하는 칼렌다타입을 가져오는 함수 
'---------------------------------------------------------------------------------------------------------
Function LookUp()
	Dim arrCalType
	Err.Clear                                                         
    
    LookUp = False 
    
    If frm1.txtPlantCd.value = "" Then
		Exit Function
	End If
    
	IF CommonQueryRs(" CAL_TYPE ", " B_PLANT ", " PLANT_CD =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrCalType = Split(lgF0, Chr(11))
		lgClnrType = arrCalType(0)
		Call LookUpOk()
	End If
    LookUp = True      
End Function

'------------------------------------------  LookUpOk()  --------------------------------------------------
'	Name : LookUpOk()
'	Description : LookUp함수를 마치면서 실행하는 함수 
'---------------------------------------------------------------------------------------------------------
Function LookUpOk()
	If lgIntFlgMode = parent.OPMD_CMODE Then
		frm1.txtClnrType.value = lgClnrType
	End If
End Function

'==========================================  3.1.1 Form_Load()  ==========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
    
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("11101000000011")
    Call InitComboBox
    Call SetDefaultVal
	Call InitVariables
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtConWcCd.focus
		Set gActiveElement = document.activeElement  
		Call LookUp
		
		frm1.txtClnrType.value = lgClnrType
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement  
	End If
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

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

Function cboInsideFlg_onchange()
	lgBlnFlgChgValue = True
End Function

Function cboWCMgr_onchange()
	lgBlnFlgChgValue = True
End Function
		 
Function txtPlantCd_OnChange()
	Call LookUP
End Function

Function txtPlantCd_OnKeyPress()
	frm1.txtPlantNm.value = ""
	frm1.txtConWcCd.value = ""
	frm1.txtConWcNm.value = "" 
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 

	If frm1.txtConWcCd.value = "" Then
		frm1.txtConWcNm.value = ""
	End If
		
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	
    Call ggoOper.ClearField(Document, "2")									
    Call Setdefaultval
    Call InitVariables														
   
    frm1.txtClnrType.value = lgClnrType
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
    frm1.txtConWcCd.value = ""
    frm1.txtConWcNm.value = ""
    Call ggoOper.ClearField(Document, "2")                                      
    Call ggoOper.LockField(Document, "N")                                       
    Call SetToolbar("11101000000011")
    Call SetDefaultVal
    Call InitVariables	
	frm1.txtClnrType.value = lgClnrType
	frm1.txtDataWcCd.focus
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
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If
	
	'-----------------------
    'Delete function call area
    '-----------------------

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		          
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
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                       
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
    
    frm1.txtValidFromDt.text  = StartDate
    frm1.txtConWcCd.value = ""
    frm1.txtConWcNm.value = "" 
    frm1.txtDataWcCd.value = ""
    'frm1.txtDataWcNm.value = ""
    frm1.txtDataWcCd.focus
    Set gActiveElement = document.activeElement 
    
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
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			
	strVal = strVal & "&txtConWcCd=" & Trim(frm1.txtConWcCd.value)			
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
    Call InitVariables														
    
    Err.Clear                                                               
    
    LayerShowHide(1)
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			
	strVal = strVal & "&txtConWcCd=" & Trim(frm1.txtConWcCd.value)			
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
	strVal = strVal & "&txtDataWcCd=" & Trim(frm1.txtDataWcCd.value)	
    
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
    
    Err.Clear                                                              
    
    DbQuery = False                                                        
    
    LayerShowHide(1)
		    
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			
	strVal = strVal & "&txtConWcCd=" & Trim(frm1.txtConWcCd.value)			
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
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")									
	Call SetToolbar("11111000111111")
	
	frm1.txtDataWcNm.focus
	Set gActiveElement = document.activeElement  
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave() 

    Err.Clear																

	DbSave = False															
	
	If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function       
	
	LayerShowHide(1)
		
    Dim strVal

	With frm1
		.txtMode.value = parent.UID_M0002											
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.Value = parent.gUsrID
		
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	
	End With
	
    DbSave = True                                                           
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()															

    frm1.txtConWcCd.value = frm1.txtDataWcCd.value 
    
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>작업장등록</font></td>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=50 tag="14"></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>작업장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtConWcCd" SIZE=15 MAXLENGTH=7 tag="12XXXU" ALT="작업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenConWC()"> <INPUT TYPE=TEXT NAME="txtConWcNm" SIZE=50 tag="14"></TD>
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
						<TABLE CLASS="TB2" CELLSPACING=0>
							<TR>
								<TD WIDTH=100% valign=top>
									<FIELDSET>
										<LEGEND>일반정보</LEGEND>
										<TABLE CLASS="TB2" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>작업장</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDataWcCd" SIZE=15 MAXLENGTH=7 tag="23XXXU" ALT="작업장">&nbsp;<INPUT TYPE=TEXT NAME="txtDataWcNm" MAXLENGTH=40 SIZE=50 tag="21" ALT="작업장명"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>작업장 구분</TD>
												<TD CLASS=TD6 NOWRAP><SELECT NAME="cboInsideFlg" ALT="작업장구분" STYLE="Width: 98px;" tag="22"></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>칼렌다 타입</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtClnrType" SIZE=10 MAXLENGTH=2 tag="22XXXU" ALT="칼렌다 타입"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCalType()">&nbsp;<INPUT TYPE=TEXT NAME="txtClnrTypeNm" SIZE=30 tag="24"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>작업장 담당자</TD>
												<TD CLASS=TD6 NOWRAP><SELECT NAME="cboWCMgr" ALT="작업장 담당자" STYLE="Width: 98px;" tag="21"><OPTION VALUE=""></OPTION></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>코스트센터</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCostCd" SIZE=17 MAXLENGTH=10 tag="22XXXU" ALT="Cost Center"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCtr" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCostCtr()">&nbsp;<INPUT NAME="txtCostNm" MAXLENGTH="20" SIZE=30 ALT ="코스트센타명" tag="24"></TD>
											</TR>										
											<TR>
												<TD CLASS=TD5 NOWRAP>유효기간</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1301ma1_I825323069_txtValidFromDt.js'></script> &nbsp;~&nbsp;
													<script language =javascript src='./js/p1301ma1_I517775994_txtValidToDt.js'></script>		
												</TD>
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
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TabIndex="-1"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TabIndex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
