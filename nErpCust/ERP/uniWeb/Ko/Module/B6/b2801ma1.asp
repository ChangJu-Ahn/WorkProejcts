<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : b2801ma1.asp
'*  4. Program Name         : Storage Location
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/04/08
'*  8. Modified date(Last)  : 2003/05/26
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              : VB Conversion
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc 선언   ****************************************** -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--'==========================================  1.1.1 Style Sheet  ====================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--==========================================  1.1.2 공통 Include   ===================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit                                                          

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  =====================================
Const BIZ_PGM_QRY_ID  = "b2801mb1.asp"										
Const BIZ_PGM_SAVE_ID = "b2801mb2.asp"										
Const BIZ_PGM_DEL_ID  = "b2801mb3.asp"										
'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Dim lgNextNo						  
Dim lgPrevNo						  
Dim IsOpenPop          

<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                                     
    lgBlnFlgChgValue = False                                             
    lgIntGrpCount = 0                                                   
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False												
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.cboSLType.value = "I"
	
	cboInvMgrTitle.style.display = ""
	cboExtSLTypeTitle.style.display = "none"
	txtBPTitle.style.display = "none"

	ggoOper.SetReqAttr frm1.cboExtSLType, "Q"
	ggoOper.SetReqAttr frm1.txtBPCd, "Q"
End Sub


Sub InitComboBox()

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("I0003", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboSLGroup,lgF0  ,lgF1  ,Chr(11))

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("I0004", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboInvMgr ,lgF0  ,lgF1  ,Chr(11))

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("B9021", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboTaxClass ,lgF0  ,lgF1  ,Chr(11))
	
	Call SetCombo(frm1.cboSLType, "I", "사내")
	Call SetCombo(frm1.cboSLType, "E", "거래처")
		
	Call SetCombo(frm1.cboExtSLType, "C", "고객")
	Call SetCombo(frm1.cboExtSLType, "S", "외주처")

End Sub

'------------------------------------------  OpenConPlant()  -------------------------------------------------
'	Name : OpenConPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장코드"		
    arrHeader(1) = "공장명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetConPlant(arrRet)
	End If	
	
End Function


'------------------------------------------  OpenConSLCd()  -------------------------------------------------
'	Name : OpenConSLCd()
'	Description : Condition Storage Location PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConSLCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtslcd.className) = "PROTECTED"  Then Exit Function
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("169901","X","X","X")   
		frm1.txtPlantCd.focus 
		Exit Function
	Else
		If Plant_SLCd_Check(0) = False Then Exit Function    
	End If

	IsOpenPop = True

	arrParam(0) = "창고팝업"											
	arrParam(1) = "B_STORAGE_LOCATION"								        
	arrParam(2) = Trim(frm1.txtSLCd.Value)						            
	arrParam(3) = ""														
	arrParam(4) = "PLANT_CD = " & Parent.FilterVar(frm1.txtPlantCd.value, "''", "S")	
	arrParam(5) = "창고"												
	
    arrField(0) = "SL_CD"													
    arrField(1) = "SL_NM"													
    
    arrHeader(0) = "창고코드"										    
    arrHeader(1) = "창고명"											    
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSLCd.focus
		Exit Function
	Else
		Call SetConSLCd(arrRet)
	End If	
End Function


'------------------------------------------  OpenBPCd()  -------------------------------------------------
'	Name : OpenBPCd()
'	Description : Business Partner Center Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenBPCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtBPCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래처팝업"					 
	arrParam(1) = "B_BIZ_PARTNER"					 
	arrParam(2) = Trim(frm1.txtBPCd.Value)	         
	arrParam(3) = ""								 
	arrParam(4) = ""								 
	arrParam(5) = "거래처"						 
	
    arrField(0) = "BP_CD"							 
    arrField(1) = "BP_NM"					         
    arrField(2) = "BP_TYPE"							 
    
    arrHeader(0) = "거래처코드"				
    arrHeader(1) = "거래처명"				
    arrHeader(2) = "종류"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBPCd.focus
		Exit Function
	Else
		Call SetBPCd(arrRet)
	End If	
    
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
			
End Function

'------------------------------------------  SetConWC()  --------------------------------------------------
'	Name : SetConWC()
'	Description : Storage Location Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConSLCd(byval arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)		
	frm1.txtSLCd.focus
	
End Function

'------------------------------------------  SetBPCd()  --------------------------------------------------
'	Name : SetBPCd()
'	Description : Business Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBPCd(byval arrRet)
	frm1.txtBPCd.value = arrRet(0)
	frm1.txtBPNm.value = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
Sub cboSLType_Onchange()
	With frm1
		If .cboSLType.value = "I" Then
			cboInvMgrTitle.style.display = ""
			cboExtSLTypeTitle.style.display = "none"
			txtBPTitle.style.display = "none"
			ggoOper.SetReqAttr .cboExtSLType, "Q"
			ggoOper.SetReqAttr .txtBPCd, "Q"
		ElseIf .cboSLType.value = "E" Then
			ggoOper.SetReqAttr .cboExtSLType, "N"
			ggoOper.SetReqAttr .txtBPCd, "N"
			cboInvMgrTitle.style.display = "none"
			cboExtSLTypeTitle.style.display = ""
			txtBPTitle.style.display = ""
		End If	
	End With
	lgBlnFlgChgValue = True
End Sub
Sub cboSLGroup_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub cboInvMgr_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub cboExtSLType_Onchange()
	lgBlnFlgChgValue = True
End Sub
Sub cboTaxClass_Onchange()
	lgBlnFlgChgValue = True
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    
    Call InitVariables														
    Call LoadInfTB19029														
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")									
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolBar("11101000000011")
    Call InitComboBox
    Call SetDefaultVal
    
    If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Parent.gPlant
		frm1.txtPlantNm.value = Parent.gPlantNm
        
		frm1.txtSLCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
		
End Sub



'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
Sub optMrpUsedFlg1_OnClick()
	lgBlnFlgChgValue = True
	frm1.txtMrpUsedFlag.value = "Y" 
End Sub

Sub optMrpUsedFlg2_OnClick()
	lgBlnFlgChgValue = True
	frm1.txtMrpUsedFlag.value = "N" 
End Sub


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                   
    
    Err.Clear                                                          

    '-----------------------
    'Check condition area
    '----------------------- 

    If Not chkField(Document, "1") Then								
       Exit Function
    End If
    
    '-----------------------
    'Check previous data area
    '----------------------- 
   

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X","X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    If Plant_SLCd_Check(1) = False Then Exit Function                                     

    '-----------------------
    'Erase contents area
    '----------------------- 

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables														
    Call SetDefaultVal
    
    '-----------------------
    'Query function call area
    '----------------------- 

    If DBQuery = False Then
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
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")         
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------

    Call ggoOper.ClearField(Document, "A")                                    
    Call ggoOper.LockField(Document, "N")                                     
    Call InitVariables														
    Call SetToolBar("11101000000011")
    Call SetDefaultVal
    frm1.optMrpUsedFlg1.disabled = false
    frm1.optMrpUsedFlg1.checked  = true
	frm1.optMrpUsedFlg2.disabled = false
	
    If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Parent.gPlant
		frm1.txtPlantNm.value = Parent.gPlantNm
        
		frm1.txtSLCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
    
    Set gActiveElement = document.activeElement
    FncNew = True															

End Function



'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    Dim intRetCD
    
    FncDelete = False														
    
    '-----------------------
    'Precheck area 
    '-----------------------

    If lgIntFlgMode = Parent.OPMD_CMODE Or _         
		UCase(Trim(frm1.txtPlantCd.Value)) <> UCase(Trim(frm1.txthPlantCd.Value)) Or _       
		UCase(Trim(frm1.txtSLCd.Value)) <> UCase(Trim(frm1.txtSLCd1.Value)) Then             
       
        Call DisplayMsgBox("900002", "X","X","X" )     
        Exit Function
    End If
    
		
    '-----------------------
    'Delete function call area
    '-----------------------

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")		       
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
    'Check content area
    '-----------------------

    If Not chkField(Document, "2") Then                            
       Exit Function
    End If
     
    '-----------------------
    'Precheck area
    '-----------------------

    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                      
        Exit Function
    End If
    
	If Trim(frm1.txtPlantCd.value) = "" then
	   Call DisplayMsgBox("169901", "X", "X", "X")
	   frm1.txtPlantCd.focus
	   Set gActiveElement = document.activeElement
	   Exit Function
	End If
   
    If lgIntFlgMode = Parent.OPMD_UMODE Then
			If UCase(Trim(frm1.txtPlantCd.Value)) <> UCase(Trim(frm1.txthPlantCd.Value)) Or _   
				UCase(Trim(frm1.txtSLCd.Value)) <> UCase(Trim(frm1.txtSLCd1.Value)) Then        
				Call DisplayMsgBox("900002", "X","X","X" )     
				Exit Function
			End If
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
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = Parent.OPMD_CMODE											
    
    
    Call ggoOper.ClearField(Document, "1")                        
    Call ggoOper.LockField(Document, "N")						
    
    frm1.txtSLCd1.value = ""
    frm1.txtSLNm1.value = ""
	
	Call cboSlType_OnChange()
    
    frm1.txtSLCd1.focus
     
   
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
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                 
        Call DisplayMsgBox("900002","X", "X", "X")                            
        
        Exit Function
    ElseIf lgPrevNo = "" Then
		Call DisplayMsgBox("900011","X", "X", "X")
		Exit Function
    End If

    strVal = BIZ_PGM_QRY_ID &	"?txtMode="		& Parent.UID_M0001				& _
								"&txtPlantCd="	& Trim(frm1.txtPlantCd.value)	& _
								"&txtSLCd="		& lgPrevNo								
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    Dim strVal

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                   
        Call DisplayMsgBox("900002", "X", "X", "X")                             
        Exit Function
    ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")
		Exit Function
    End If

    strVal = BIZ_PGM_QRY_ID &	"?txtMode="    & Parent.UID_M0001				& _
								"&txtPlantCd=" & Trim(frm1.txtPlantCd.value)	& _
								"&txtSLCd="    & lgNextNo								
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)											
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                    
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")				
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
    
    Dim strVal
    
    Call LayerShowHide(1) 
    
    strVal = BIZ_PGM_DEL_ID &	"?txtMode="    & Parent.UID_M0003				& _
								"&txtPlantCd=" & Trim(frm1.txtPlantCd.value)	& _			
								"&txtSLCd1="   & Trim(frm1.txtSLCd1.value)			
    
	Call RunMyBizASP(MyBizASP, strVal)										
	
    DbDelete = True                                                         

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================

Function DbDeleteOk()														

    Call InitVariables
	Call MainNew()
	
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    
    Err.Clear                                                           
    
    DbQuery = False                                                     
    
    Dim strVal
    
    Call LayerShowHide(1) 
    
    strVal = BIZ_PGM_QRY_ID &	"?txtMode="    & Parent.UID_M0001				& _
								"&txtPlantCd=" & Trim(frm1.txtPlantCd.value)	& _		
								"&txtSLCd="    & Trim(frm1.txtSLCd.value)			
	    
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
    lgIntFlgMode = Parent.OPMD_UMODE											
    
    Call ggoOper.LockField(Document, "Q")									
    
    If frm1.cboSLType.value = "E" Then
		cboInvMgrTitle.style.display = "none"
		cboExtSLTypeTitle.style.display = ""
		txtBPTitle.style.display = ""
	Else
		cboInvMgrTitle.style.display = ""
		cboExtSLTypeTitle.style.display = "none"
		txtBPTitle.style.display = "none"
	End If

	Call SetToolBar("11111000001111")
	frm1.txtSLCd.focus
	Set gActiveElement = document.activeElement
	End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================

Function DbSave() 

    Err.Clear														

	DbSave = False													

    Dim strVal

	Call LayerShowHide(1) 
	
	With frm1
		.txtMode.value       = Parent.UID_M0002							
		.txtFlgMode.value    = lgIntFlgMode
		
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	
	End With
	
    DbSave = True                                                      
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================

Function DbSaveOk()														

    frm1.txtSLCd.value = frm1.txtSLCd1.value 
    Call InitVariables
    Call MainQuery()

End Function

'========================================================================================
' Function Name : Plant_SLCd_Check
' Function Desc : 
'========================================================================================

Function Plant_SLCd_Check(ByVal ChkIndex)
	
	
	'-----------------------
	'Check Plant CODE		
	'-----------------------
	If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & Parent.FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Plant_SLCd_Check = False
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
			
	If ChkIndex	>= 1 Then        

		'-----------------------
		'Check SLCd CODE	
		'-----------------------
		If 	CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " SL_CD = " & Parent.FilterVar(frm1.txtSLCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
			Call DisplayMsgBox("125700","X","X","X")
			frm1.txtSLNm.Value = ""
			frm1.txtSLCd.focus
			Plant_SLCd_Check = False
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtSLNm.Value = lgF0(0)
		
	End If
	
	Plant_SLCd_Check = True
	
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2KCM.inc"  -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
    <TR>
		<TD <%=HEIGHT_TYPE_00%> >
		</TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
			<TABLE <%=LR_SPACE_TYPE_20%> >
				<TR>
					<TD <%=HEIGHT_TYPE_02%> >
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD656 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 tag="14"></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>창고</TD>
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=10 MAXLENGTH=7 tag="12XXXU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSLCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSLCd()"> <INPUT TYPE=TEXT NAME="txtSLNm" SIZE = 40 tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% >
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>창고</TD>
								<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd1" SIZE=10  MAXLENGTH=7 tag="23XXXU" ALT="창고">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm1" SIZE=25 MAXLENGTH = 40 tag="22" ALT="창고명"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>창고타입</TD>
								<TD CLASS=TD656 NOWRAP><SELECT NAME="cboSLType" ALT="창고타입" STYLE="Width: 98px;" tag="23"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>창고그룹</TD>
								<TD CLASS=TD656 NOWRAP><SELECT NAME="cboSLGroup" ALT="창고그룹" STYLE="Width: 98px;" tag="21"><OPTION VALUE = ""></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>MRP사용여부</TD>												
								<TD CLASS=TD656 NOWRAP><SPAN STYLE="width:70;"><INPUT TYPE="RADIO" NAME="optMrpUsedFlg" ID="optMrpUsedFlg1" CLASS="RADIO" tag="2N" Value="Y" CHECKED><LABEL FOR="optMrpUsedFlg1">예</LABEL></SPAN>
							    					   <SPAN STYLE="width:70;"><INPUT TYPE="RADIO" NAME="optMrpUsedFlg" ID="optMrpUsedFlg2" CLASS="RADIO" tag="2N" Value="N"><LABEL FOR="optMrpUsedFlg2">아니오</LABEL></SPAN></TD>												
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>과세유형</TD>
								<TD CLASS=TD656 NOWRAP><SELECT NAME="cboTaxClass" ALT="과세유형" STYLE="Width: 98px;" tag="21"><OPTION VALUE = ""></OPTION></SELECT></TD>
							</TR>
							<TR ID="cboInvMgrTitle" STYLE="DISPLAY: ">
								<TD CLASS=TD5 NOWRAP>재고담당자</TD>
								<TD CLASS=TD656 NOWRAP><SELECT NAME="cboInvMgr" ALT="재고담당자" STYLE="Width: 98px;" tag="21"><OPTION VALUE=""></OPTION></SELECT></TD>
							</TR>
							<TR ID="cboExtSLTypeTitle" STYLE="DISPLAY: none">
							   <TD CLASS=TD5 NOWRAP>거래처타입</TD>
								<TD CLASS=TD656 NOWRAP><SELECT NAME="cboExtSLType" ALT="거래처타입" STYLE="Width: 98px;" tag="23"><OPTION Value = ""></OPTION></SELECT></TD>
							</TR>
							<TR ID="txtBPTitle" STYLE="DISPLAY: none">
								<TD CLASS=TD5 NOWRAP>거래처</TD>
								<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtBPCd" SIZE=10 MAXLENGTH=10 tag="23XXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBPCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBPCd()">
							        &nbsp;<INPUT TYPE=TEXT NAME="txtBPNm" SIZE=30 MAXLENGTH=20 tag="24" ALT="거래처명"></TD>
							</TR>
							<% SubFillRemBodyTD656 (11)%>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
	    <TD <%=HEIGHT_TYPE_01%> >
	    </TD>
	</TR>
	<TR HEIGHT=20 >
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%> >
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMrpUsedFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txthPlantCd" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

