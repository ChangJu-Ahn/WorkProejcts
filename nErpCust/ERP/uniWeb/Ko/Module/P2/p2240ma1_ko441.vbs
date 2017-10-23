
'==========================================================================================================
Const BIZ_PGM_QRY_ID       = "p2240mb1_ko441.asp"
Const BIZ_PGM_EXECUTE_ID   = "p2240mb2_ko441.asp"
Const BIZ_PGM_EXECUTE_ID1  = "p2240mb3_ko441.asp"       '20080304::hanc
Const BIZ_PGM_ID3          = "p2240mb4_ko441.asp"			    '20080303::hanc         '��: Biz Logic ASP Name



Dim IsOpenPop         

'=========================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
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

	arrParam(0) = "�����˾�"						'�˾� ��Ī 
	arrParam(1) = "B_PLANT"								'TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			'Code Condition
	arrParam(3) = ""									'Name Cindition
	arrParam(4) = ""									'Where Condition
	arrParam(5) = "����"							'TextBox ��Ī 
	
   	arrField(0) = "PLANT_CD"							'Field��(0)
    arrField(1) = "PLANT_NM"							'Field��(1)
    
    arrHeader(0) = "����"							'Header��(0)
    arrHeader(1) = "�����"							'Header��(1)
    
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

'20080303::hanc----------------------------------------------------------------------
'�����ȹ�Ⱓ ��������---------------------------------------------------------------
Function DbQueryPeriod()
    DbQueryPeriod = False
    Err.Clear                                                                        '��: Clear err status

    Dim strVal

    With frm1
        strVal = BIZ_PGM_ID3 & "?txtMode="      & parent.UID_M0004
    End With

    Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic

    DbQueryPeriod = True
End Function

'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Function

'---------------------------------------------------------------------------------------------------------
'        Name : ExecuteMPS()    
'        Description : MPS ���� Main Function          
'---------------------------------------------------------------------------------------------------------
Function ExecuteMPS()
    Err.Clear
    ExecuteMPS = False
	
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	Dim IntRetCD
	IntRetCD = DisplayMsgBox("ZZ0015",parent.VB_YES_NO, "X", "X")       '20080312::hanc  900018
	
	If IntRetCD = vbNo Then
		Exit Function
	End If		
    
    Call LayerShowHide(1)	
    
    With frm1
		.txtMode.value = parent.UID_M0004
		.txtFlgMode.value = lgIntFlgMode
    End With	

	Call ExecMyBizASP(frm1, BIZ_PGM_EXECUTE_ID1)										
    
    ExecuteMPS = True 

    lgBlnFlgChgValue = False
            
End Function

'20080304::hanc
Function DbExecOk()
    Call DisplayMsgBox("ZZ0010","X", "X", "X")
 
End Function


'20080304::hanc
Function DbQueryOk()
    Call DbQueryPeriod      '20080303::hanc

'    If lgIntFlgMode <> parent.OPMD_UMODE Then
'		frm1.vspdData1.Col = C_OrdNo
'		frm1.vspdData1.Row = 1
'		frm1.KeyProdOrdNo.value = Trim(frm1.vspdData1.Text)
'			
'		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
'		Set gActiveElement = document.activeElement
'		
'		If DbQuery2 = False Then	
'			Call RestoreToolBar()
'			Exit Function
'		End If	
'		
'		lgOldRow1 = 1
'		
'    End If

       
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
