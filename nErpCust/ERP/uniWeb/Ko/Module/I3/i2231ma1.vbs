'******************************************  1.2 Global 변수/상수 선언  ***********************************
' 1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Const BIZ_SIMUL_ID		= "i2231mb3.asp" 
Const BIZ_CONFIRM_ID	= "i2231mb2.asp"
Const BIZ_CANCEL_ID		= "i2231mb4.asp" 


'==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                                     
    lgBlnFlgChgValue = False                                             
    lgIntGrpCount = 0                                                    
    
    IsOpenPop = False         
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	If UCase(Parent.gPlant) <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		Call txtPlantCd_LostFocus
	End If
End Sub

'------------------------------------------ OpenPlant()  --------------------------------------------------
' Name : OpenPlant()
' Description : Plant Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
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
 arrField(2) = "Convert(varchar(40), INV_CLS_DT)" 
 
 arrHeader(0) = "공장"  
 arrHeader(1) = "공장명" 
 arrHeader(2) = "최종마감일" 
    
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
 IsOpenPop = False
 
 If arrRet(0) = "" Then
	frm1.txtPlantCd.focus
	Exit Function
 Else
	Call SetPlant(arrRet)
 End If 

End Function

'==========================================  2.4.3 Set???()  =============================================
' Name : Set???()
' Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '------------------------------------------  SetPlant()  --------------------------------------------------
' Name : SetPlant()
' Description : Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byRef arrRet)
    Dim strYear
    Dim strMonth
    Dim strDay
    
	frm1.txtPlantCd.Value    = arrRet(0)  
	frm1.txtPlantNm.Value    = arrRet(1)
 
	Call ExtractDateFrom(arrRet(2),Parent.gDateFormat, Parent.gComDateType, strYear, strMonth, strDay)

	frm1.txtInvClsDt.Year  =  strYear
	frm1.txtInvClsDt.Month =  strMonth
	lgBlnFlgChgValue = True  
End Function

'=======================================================================================================
'   Event Name : txtInvClsDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtInvClsDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInvClsDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtInvClsDt.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtPlantCd_LostFocus()
'   Event Desc : 공장명과 최종마감년월을 찾는다.
'=======================================================================================================
Sub txtPlantCd_LostFocus()
    Dim strYear
    Dim strMonth
    Dim strDay
	
	If frm1.txtPlantCd.value <> "" Then
		If  CommonQueryRs(" PLANT_NM, CONVERT(CHAR(10), INV_CLS_DT, 21) "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			frm1.txtPlantNm.Value  = ""
			frm1.txtInvClsDt.text  = ""
			Exit Sub
		Else
			lgF0 = Split(lgF0,Chr(11))
			lgF1 = Split(lgF1,Chr(11))
			
			frm1.txtPlantNm.Value = lgF0(0)
			Call ExtractDateFrom(lgF1(0), Parent.gServerDateFormat, Parent.gServerDateType , strYear, strMonth, strDay)

			If CommonQueryRs("CLOSE_DT","I_INV_CLOSING_HISTORY","CLOSE_FLAG = " & FilterVar("Y", "''", "S") & "  AND PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then										
				frm1.btnCancel.disabled = false
			Else				
				frm1.btnCancel.disabled = true
			End if
		End If

		frm1.txtInvClsDt.Year  =  strYear
		frm1.txtInvClsDt.Month =  strMonth
	Else
		frm1.txtPlantNm.Value  = ""
		frm1.txtInvClsDt.text  = ""
	End If
End Sub

'========================================================================================
' Function Name : FncSave1
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave1() 
    Dim IntRetCD 
    
    FncSave1 = False                                                        
    
    Err.Clear                                                               
    
  '-----------------------
    'Check content area
    '-----------------------
    If Not chkFieldByCell(frm1.txtPlantCd, "A",1) Then Exit Function
    If Not chkFieldByCell(frm1.txtInvClsDt, "A",1) Then Exit Function                      
    
  '-----------------------
    'Save function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO,"X","X")
	If IntRetCD = vbNo Then Exit Function

	If DbSave1 = False Then 
		Exit Function
	End If

    FncSave1 = True                                                      
    
End Function

'========================================================================================
' Function Name : FncSave2
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave2() 
    Dim IntRetCD 
    
    FncSave2 = False                                                    
    
    Err.Clear                                                             
    
  '-----------------------
    'Check content area
    '-----------------------
    If Not chkFieldByCell(frm1.txtPlantCd, "A",1) Then Exit Function
    If Not chkFieldByCell(frm1.txtInvClsDt, "A",1) Then Exit Function
    
  '-----------------------
    'Save function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO,"X","X")
	If IntRetCD = vbNo Then Exit Function

	If DbSave2 = False Then
		Exit Function
	End If
    
    FncSave2 = True                                                       
    
End Function



'========================================================================================
' Function Name : FncSave1
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave3() 
    Dim IntRetCD 
    
    FncSave3 = False                                                        
    
    Err.Clear                                                               
    
  '-----------------------
    'Check content area
    '-----------------------
    If Not chkFieldByCell(frm1.txtPlantCd, "A",1) Then Exit Function
    If Not chkFieldByCell(frm1.txtInvClsDt, "A",1) Then Exit Function
    
  '-----------------------
    'Save function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO,"X","X")
	If IntRetCD = vbNo Then Exit Function

	If DbSave3 = False Then
		Exit Function
	End If
    
    FncSave3 = True                                                        
    
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , True)                
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : DBSave1
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave1() 
 
	Call LayerShowHide(1)
	    
	Err.Clear             

	DbSave1 = False             

	frm1.txtInsrtUserId.value	= Parent.gUsrID
	frm1.btnCancel.disabled		= True
	frm1.btnConfirm.disabled	= True
	frm1.btnRun.disabled		= True

	Call ExecMyBizASP(frm1, BIZ_SIMUL_ID)          

	DbSave1 = True                                                     
	
End Function

'========================================================================================
' Function Name : DBSave2
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave2() 
 
	Call LayerShowHide(1)
	    
	Err.Clear              

	DbSave2 = False            

	frm1.txtInsrtUserId.value	= Parent.gUsrID
	frm1.btnCancel.disabled		= True
	frm1.btnConfirm.disabled	= True
	frm1.btnRun.disabled		= True

	Call ExecMyBizASP(frm1, BIZ_CONFIRM_ID)          
	
	DbSave2 = True                                                          
	
End Function

'========================================================================================
' Function Name : DBSave3
' Function Desc : 재고마감취소, 성공적이면 DBSaveOk 호출됨 
'========================================================================================
Function DbSave3() 
 
	Call LayerShowHide(1)
	    
	Err.Clear              

	DbSave3 = False            

	frm1.txtInsrtUserId.value	= Parent.gUsrID
	frm1.btnCancel.disabled		= True
	frm1.btnConfirm.disabled	= True
	frm1.btnRun.disabled		= True

	Call ExecMyBizASP(frm1, BIZ_CANCEL_ID)          
  
	DbSave3 = True                                                         

End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk1()              
	Dim IntRetCD
	IntRetCD = DisplayMsgBox("990000","X", "X", "X")
	frm1.txtPlantCd.focus
	Call txtPlantCd_LostFocus()
End Function

Function DbSaveOk2()              
	Dim IntRetCD
	IntRetCD = DisplayMsgBox("990000","X", "X", "X")
	frm1.txtPlantCd.focus
End Function

Function DbSaveOk3()              
	Dim IntRetCD
	IntRetCD = DisplayMsgBox("990000","X", "X", "X")
	frm1.txtPlantCd.focus
	Call txtPlantCd_LostFocus()
End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function Btnabled()
	If CommonQueryRs("CLOSE_DT","I_INV_CLOSING_HISTORY","CLOSE_FLAG = " & FilterVar("Y", "''", "S") & "  AND PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
	lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then										
		frm1.btnCancel.disabled = False
	Else				
		frm1.btnCancel.disabled = True
	End if
	frm1.btnConfirm.disabled	= False
	frm1.btnRun.disabled		= False
End Function

'------------------------------------------  OpenOnhandDtlRef()  -------------------------------------------------
' Name : OpenOnhandDtlRefCode()
' Description : OnahndStock detail Reference
'--------------------------------------------------------------------------------------------------------- 
Function OpenCancelListRef()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1 
	Dim Param2
	Dim Param3
	 
	If IsOpenPop = True Then Exit Function

	Param1 = Trim(frm1.txtPlantCd.value)
	Param2 = Trim(frm1.txtPlantNm.value)
	Param3 = Trim(frm1.txtInvClsDt.Text)
	  
	if Param1 = "" then
		Call DisplayMsgBox("169901","X", "X", "X")    
		frm1.txtPlantCd.focus
		Exit Function
	End If
	 
	IsOpenPop = True

	iCalledAspName = AskPRAspName("I2231RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2231RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	 
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2,Param3), _
	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")      
	     
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	End If 
End Function
