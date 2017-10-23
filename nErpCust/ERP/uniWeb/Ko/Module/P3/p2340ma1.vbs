
Const BIZ_PGM_EXECUTE_ID	= "p2340mb1.asp"
Const BIZ_PGM_DATACHECK_ID  = "p2340mb2.asp"
Const BIZ_PLANT_ID			= "p2340mb3.asp"
Const BIZ_PGM_BATCH_ID		= "p2340mb4.asp"


Dim lgNextNo
Dim lgPrevNo

Dim IsOpenPop         
Dim lgInvCloseDt
 
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


'=========================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtFixExecFromDt.text = EndDate
End Sub

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

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
	End If	
End Function


'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(ByRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)	
	Call ExecMyBizASP(frm1, BIZ_PLANT_ID)	
	frm1.txtErrorQty.text = 0
	
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Function

Sub txtPlantCd_OnChange()
	If frm1.txtPlantCd.value <> "" Then
	
		If gLookUpEnable = False Then Exit Sub
		
		Call ExecMyBizASP(frm1, BIZ_PLANT_ID)
		frm1.txtErrorQty.text = 0
	End If

End Sub


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
	
	iCalledAspName = AskPRAspName("P2340RA1")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P2340RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName , Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
'        Name : DataCheck()    
'        Description : MRP 전개 Main Function          
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function DataCheck()
    Err.Clear
    DataCheck = False

    If Not chkField(Document, "1") Then
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

	Call ExecMyBizASP(frm1, BIZ_PGM_DATACHECK_ID)										

    End With	
    
    DataCheck = True 
    lgBlnFlgChgValue = False
            
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
'        Name : ExecuteMRP()    
'        Description : MRP 전개 Main Function          
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function ExecuteMRP()
    Err.Clear
    ExecuteMRP = False

    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If    
	
	If ValidDateCheck(frm1.txtFixExecFromDt, frm1.txtFixExecToDt)  = False Then
		frm1.txtFixExecToDt.focus 
		Exit Function
	End If

	If ValidDateCheck(frm1.txtFixExecToDt, frm1.txtPlanExecToDt)  = False Then		
		frm1.txtPlanExecToDt.focus 
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

	Call ExecMyBizASP(frm1, BIZ_PGM_EXECUTE_ID)										

    End With	
    
    ExecuteMRP = True 
    lgBlnFlgChgValue = False
            
End Function

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
'        Name : ExecuteBatch()    
'        Description : MRP 전개 & 승인Main Function          
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function ExecuteBatch()
    Err.Clear
    ExecuteBatch = False

    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If    
	
	If ValidDateCheck(frm1.txtFixExecFromDt, frm1.txtFixExecToDt)  = False Then	
		Call DisplayMsgBox("183117", "X", "X", "X")
		frm1.txtFixExecToDt.focus 
		Exit Function
	End If
	
	If ValidDateCheck(frm1.txtFixExecToDt, frm1.txtPlanExecToDt)  = False Then	
		Call DisplayMsgBox("183118", "X", "X", "X")
		frm1.txtPlanExecToDt.focus 
		Exit Function
	End If   
	
	Dim IntRetCD
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Dim strVal
    
    Call LayerShowHide(1)
    
    With frm1
	.txtMode.value = parent.UID_M0002
	.txtFlgMode.value = lgIntFlgMode

	Call ExecMyBizASP(frm1, BIZ_PGM_BATCH_ID)										

    End With	
    
    ExecuteBatch = True 
    lgBlnFlgChgValue = False
            
End Function


'=======================================================================================================
'   Event Name : txtFixExecFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFixExecFromDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtFixExecFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFixExecFromDt.Focus 
	End If 
End Sub
'=======================================================================================================
'   Event Name : txtFixExecFromDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtFixExecFromDt_Change() 
	lgBlnFlgChgValue = True 
End Sub

'=======================================================================================================
'   Event Name : txtFixExecFromDt_OnBlur()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtFixExecFromDt_OnBlur()
	Dim DtInvCloseDt
	Dim DtExecFromDt

	If frm1.txtFixExecFromDt.text = "" Then Exit Sub
	
	DtInvCloseDt = UniConvDateAToB(lgInvCloseDt, parent.gDateFormat, parent.gServerDateFormat)
	DtExecFromDt = UniConvDateAToB(frm1.txtFixExecFromDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	
	If DtExecFromDt <= DtInvCloseDt Then
		Call DisplayMsgBox("189250", "x", "x", "x")
		frm1.txtFixExecFromDt.text = UNIDateAdd ("D", 1, lgInvCloseDt, parent.gDateFormat)
		frm1.txtFixExecFromDt.focus
		Set gActiveElement = document.activeElement
		Exit Sub
	End If
	
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFixExecToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtFixExecToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFixExecToDt.Focus
	End If 
End Sub
'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtFixExecToDt_Change() 
	lgBlnFlgChgValue = True 
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtPlanExecToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtPlanExecToDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtPlanExecToDt.Focus 
	End If 
End Sub
'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtPlanExecToDt_Change() 
	lgBlnFlgChgValue = True 
End Sub


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
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
