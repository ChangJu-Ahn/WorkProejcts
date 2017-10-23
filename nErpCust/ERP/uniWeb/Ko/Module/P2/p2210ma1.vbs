
'==========================================================================================================
Const BIZ_PGM_QRY_ID       = "p2210mb1.asp"
Const BIZ_PGM_EXECUTE_ID   = "p2210mb2.asp"


Dim IsOpenPop         

'=========================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
    IsOpenPop = False

End Sub


'---------------------------------------------------------------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						'팝업 명칭 
	arrParam(1) = "B_PLANT"								'TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			'Code Condition
	arrParam(3) = ""									'Name Cindition
	arrParam(4) = ""									'Where Condition
	arrParam(5) = "공장"							'TextBox 명칭 
	
   	arrField(0) = "PLANT_CD"							'Field명(0)
    arrField(1) = "PLANT_NM"							'Field명(1)
    
    arrHeader(0) = "공장"							'Header명(0)
    arrHeader(1) = "공장명"							'Header명(1)
    
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

'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Function

'---------------------------------------------------------------------------------------------------------
'        Name : ExecuteMPS()    
'        Description : MPS 전개 Main Function          
'---------------------------------------------------------------------------------------------------------
Function ExecuteMPS()
    Err.Clear
    ExecuteMPS = False
	
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	Dim IntRetCD
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If		
    
    Call LayerShowHide(1)	
    
    With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode
    End With	

	Call ExecMyBizASP(frm1, BIZ_PGM_EXECUTE_ID)										
    
    ExecuteMPS = True 

    lgBlnFlgChgValue = False
            
End Function


'=======================================================================================================
'   Event Name : txtPlanDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtPlanDt_Change() 
	lgBlnFlgChgValue = True 
End Sub

Sub txtPlanDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtPlanDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtPlanDt.Focus
	End If 
End Sub

Sub txtPlantCd_OnChange()
    If frm1.txtPlantCd.value <> "" Then
    
	    If DBQuery = False Then 
           Call RestoreToolBar()
           Exit Sub 
        End If 
	End If
End Sub

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
    
    Dim strVal
    
    If gLookUpEnable = False Then Exit Function
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear
    
    Call ExecMyBizASP(frm1, BIZ_PGM_QRY_ID)
    
    DbQuery = True
        
End Function
