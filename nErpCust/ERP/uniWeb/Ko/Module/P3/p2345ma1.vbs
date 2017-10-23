
'==========================================================================================================
Const BIZ_PGM_QRY_ID			= "p2345mb1.asp"
Const BIZ_PGM_BATCH_ID			= "p2345mb2.asp"
Const BIZ_PGM_APPROVECANCEL_ID  = "p2345mb3.asp"


Dim IsOpenPop   

'==========================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE 
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0 
    
    IsOpenPop = False

End Sub

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_PLANT"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "공장"							' TextBox 명칭 
	
   	arrField(0) = "PLANT_CD"							' Field명(0)
    arrField(1) = "PLANT_NM"							' Field명(1)
    
    arrHeader(0) = "공장"							' Header명(0)
    arrHeader(1) = "공장명"							' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPlant(arrRet)
		If DBQuery = False Then 
           Call RestoreToolBar()
           Exit Function
        End If 
	End If	
End Function

'------------------------------------------  OpenErrorList()  -------------------------------------------------
'	Name : OpenErrorList()
'	Description : Part Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenErrorList()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value  = "" Then
		call DisplayMsgBox("220705", "X","X","X")
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtMRPHisNo.value)
	
	iCalledAspName = AskPRAspName("P2345RA1")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P2345RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName , Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(ByRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement			
End Function

Sub txtPlantCd_OnChange()
    If frm1.txtPlantCd.value <> "" Then
		If DBQuery = False Then 
           Call RestoreToolBar()
           Exit Sub
        End If 
	End If
End Sub

'==========================================================================================================
'        Name : BatchMRP()    
'        Description : MRP 전개 Main Function          
'==========================================================================================================
Function BatchConfirmMRP()
	Dim IntRetCD

    Err.Clear

    BatchConfirmMRP = False	
	
    If Not chkField(Document, "2") Then
       Exit Function
    End If	
	
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Call LayerShowHide(1)
    
    With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
	
		Call ExecMyBizASP(frm1, BIZ_PGM_BATCH_ID)										
    End With	
    
    BatchConfirmMRP = True 

    lgBlnFlgChgValue = False
            
End Function

'==========================================================================================================
'        Name : ApproveCancelMRP()    
'        Description : 
'==========================================================================================================
Function ApproveCancelMRP()
	Dim IntRetCD

    Err.Clear
    ApproveCancelMRP = False
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If	

	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Call LayerShowHide(1)

    With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode

		Call ExecMyBizASP(frm1, BIZ_PGM_APPROVECANCEL_ID)										
    End With	
    
    ApproveCancelMRP = True  

End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
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
	FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

	If gLookUpEnable = False Then Exit Function

    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear

    Call ExecMyBizASP(frm1, BIZ_PGM_QRY_ID)
    
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
	lgBlnFlgChgValue = False
End Function