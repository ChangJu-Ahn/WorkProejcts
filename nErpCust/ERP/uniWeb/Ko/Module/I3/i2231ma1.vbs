'******************************************  1.2 Global ����/��� ����  ***********************************
' 1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Const BIZ_SIMUL_ID		= "i2231mb3.asp" 
Const BIZ_CONFIRM_ID	= "i2231mb2.asp"
Const BIZ_CANCEL_ID		= "i2231mb4.asp" 


'==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                                     
    lgBlnFlgChgValue = False                                             
    lgIntGrpCount = 0                                                    
    
    IsOpenPop = False         
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
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

 arrParam(0) = "�����˾�" 
 arrParam(1) = "B_PLANT"    
 arrParam(2) = Trim(frm1.txtPlantCd.Value)
 arrParam(3) = ""
 arrParam(4) = ""   
 arrParam(5) = "����"   
 
 arrField(0) = "PLANT_CD" 
 arrField(1) = "PLANT_NM"
 arrField(2) = "Convert(varchar(40), INV_CLS_DT)" 
 
 arrHeader(0) = "����"  
 arrHeader(1) = "�����" 
 arrHeader(2) = "����������" 
    
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
' Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '------------------------------------------  SetPlant()  --------------------------------------------------
' Name : SetPlant()
' Description : Plant Popup���� Return�Ǵ� �� setting
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
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : ������ ������������� ã�´�.
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
' Function Desc : ȭ�� �Ӽ�, Tab���� 
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
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
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
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
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
' Function Desc : ��������, �������̸� DBSaveOk ȣ��� 
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
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
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
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
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
