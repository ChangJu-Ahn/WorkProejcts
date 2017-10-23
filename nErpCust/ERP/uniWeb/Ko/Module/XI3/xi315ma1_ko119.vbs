Option Explicit   

' External ASP File
'========================================
Const BIZ_PGM_ID = "xi315mb1_ko119.asp"											'��: �����Ͻ� ���� ASP��

' Constant variables 
'========================================
Const C_MaxKey          = 1                                           

' User-defind Variables
'========================================
Dim IsOpenPop     
Dim lgIsOpenPop     

'==========================================
Dim lsDnNo         

Dim GridBLNo			
Dim GridBLDocNo
Dim GridApplicantCd
Dim GridApplicantNm
Dim GridCur
Dim GridDocAmt
Dim lgRadio


'=========================================
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_CMODE                          
    lgSortKey        = 1
End Sub


'=========================================
Sub SetDefaultVal()

	lgRadio	= "MU"
	
	frm1.txtSendStartDt.Text = StartDate
	frm1.txtSendEndDt.Text = EndDate
	frm1.txtPlanStartDt.Text = StartDate
	frm1.txtPlanEndDt.Text = EndDate

End Sub

'=========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("xi315ma1_KO119","S","A","V20021106", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetSpreadLock 
End Sub

'=========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'=========================================
Sub Form_Load()

	Call LoadInfTB19029

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec) 
'	Call ggoOper.FormatNumber(frm1.txtMassSumQty, "999999999999", "0", True)
'	Call ggoOper.FormatNumber(frm1.txtSampleSumQty,  "999999999999", "0", True)
'	Call ggoOper.FormatNumber(frm1.txtInventorySumQty,   "999999999999", "0", True)
'	Call ggoOper.FormatNumber(frm1.txtOutSumQty,   "999999999999", "0", True)
'	Call ggoOper.FormatNumber(frm1.txtNowSumQty,   "999999999999", "0", True)
'	Call ggoOper.FormatNumber(frm1.txtHoldSumQty,   "999999999999", "0", True)
'	Call ggoOper.FormatNumber(frm1.txtVOutSumQty,   "999999999999", "0", True)
	
    Call ggoOper.LockField(Document, "Q")                                   '��: Lock  Suitable  Field
         
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
			
    Call SetDefaultVal
    Call InitVariables		'��: Initializes local global variables

	Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 
	
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
'		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If

	frm1.txtMassSumQty.allownull = False
    frm1.txtSampleSumQty.allownull = False
    frm1.txtInventorySumQty.allownull = False
    frm1.txtOutSumQty.allownull = False
    frm1.txtNowSumQty.allownull = False
    frm1.txtHoldSumQty.allownull = False
    frm1.txtVOutSumQty.allownull = False
	
End Sub

'=========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================
Function FncQuery() 

    FncQuery = False                                                        
   
    Err.Clear     
                                                             
	If ValidDateCheck(frm1.txtPlanStartDt, frm1.txtPlanEndDt) = False Then Exit Function
	If ValidDateCheck(frm1.txtSendStartDt, frm1.txtSendEndDt) = False Then Exit Function

	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

    Call ggoOper.ClearField(Document, "2")
    Call InitVariables 														

    If DbQuery = False Then Exit Function

    FncQuery = True															

End Function

'========================================
Function FncPrint()
    FncPrint = False                                                             
    Err.Clear                                                                    
	Call Parent.FncPrint()                                                       
    FncPrint = True                                                              
End Function

'========================================
Function FncExcel() 
    FncExcel = False                                                             
    Err.Clear                                                                    

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              
End Function

'========================================
Function FncFind() 
    FncFind = False                                                              
    Err.Clear                                                                    

	Call Parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               
End Function

'========================================
Function FncExit()
    FncExit = True                                                               
End Function

'========================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               
	Call LayerShowHide(1)
    
    With frm1

	If lgIntFlgMode = parent.OPMD_UMODE Then  
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(.txthPlantCd.Value)
		strVal = strVal & "&txtPlanStartDt=" & Trim(.txthPlanStartDt.value)
		strVal = strVal & "&txtPlanEndDt=" & Trim(.txthPlanEndDt.value)
		strVal = strVal & "&txtSendStartDt=" & Trim(.txthSendStartDt.value)
		strVal = strVal & "&txtSendEndDt=" & Trim(.txthSendEndDt.value)
		strVal = strVal & "&txtSecItemCd=" & Trim(.txthSecItemCd.Value)									
		strVal = strVal & "&txtRadio="			& lgRadio		
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag   	
   	Else  	   	
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode											
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)
		strVal = strVal & "&txtPlanStartDt=" & Trim(.txtPlanStartDt.Text)
		strVal = strVal & "&txtPlanEndDt=" & Trim(.txtPlanEndDt.Text)
		strVal = strVal & "&txtSendStartDt=" & Trim(.txtSendStartDt.Text)
		strVal = strVal & "&txtSendEndDt=" & Trim(.txtSendEndDt.Text)
		strVal = strVal & "&txtSecItemCd=" & Trim(.txtSecItemCd.Value)					
		strVal = strVal & "&txtRadio="			& lgRadio
		strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag
	End If			
        strVal =     strVal & "&lgPageNo="       & lgPageNo                          '��: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

    Call RunMyBizASP(MyBizASP, strVal)										

    End With
    
    DbQuery = True
End Function

'========================================
Function DbQueryOk()														
		
	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE								
	
	Call SetToolbar("11000000000111")	
	
    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    Else
       frm1.txtPlantCd.focus	
    End if
    
    Call Check2()  	

End Function

Function DbQueryNotOk()														
		
	lgBlnFlgChgValue = False
'    lgIntFlgMode     = parent.OPMD_UMODE								
	
	Call SetToolbar("11000000000111")	
	
    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    Else
       frm1.txtPlantCd.focus	
    End if
    
    Call Check2()  	

End Function


Function Check2()
	Dim MassInv, SampleInv,  InvSum    '�԰��
	Dim MassOutSum, SampleOutSum, OutSum   '����
	Dim MassHoldGSum, MassUseGSum, MGoodSum   '�������
	Dim SampleHoldGSum, SampleUseGSum, SampleGSum   'Sample����
	Dim GoodSum  '����

	MassInv = cdbl(frm1.txtMassSumQty.text)  '����԰�
	SampleInv = cdbl(frm1.txtSampleSumQty.text)  'Sample�԰�

	InvSum   = cdbl(MassInv) + cdbl(SampleInv)   '�԰��
	
	MassOutSum = cdbl(frm1.txtMOutSumQty.text)	        '������
	SampleOutSum = cdbl(frm1.txtSampleOutSumQty.text)  'Sample���
	OutSum   =  cdbl(MassOutSum) + cdbl(SampleOutSum)	'����
	
	MassHoldGSum = cdbl(frm1.txtMHoldSumQty.text)     '���Hold���
	MassUseGSum  = cdbl(frm1.txtMUseSumQty.text)      '��갡�����
	MGoodSum     = cdbl(MassHoldGSum) + cdbl(MassUseGSum)   '������
	
	SampleHoldGSum = cdbl(frm1.txtSampleHoldSumQty.text)  'Sample Hold ���
	SampleUseGSum  = cdbl(frm1.txtSampleUseSumQty.text)     'Sample �������
	SampleGSum   = cdbl(SampleHoldGSum) + cdbl(SampleUseGSum)  'Sample ���
	
	GoodSum      = cdbl(MGoodSum) + cdbl(SampleGSum)    '����			

	frm1.txtInventorySumQty.text = InvSum
	frm1.txtOutSumQty.text = OutSum
	frm1.txtMGoodsSumQty.text = MGoodSum
	frm1.txtSampleGoodsSumQty.text = SampleGSum
	frm1.txtGoodsSumQty.text = GoodSum

End Function

'------------------------------------------  OpenPlant()  ------------------------------------------------
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  SetPlant()  -------------------------------------------------
'	Name : SetPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function


'------------------------------------------  OpenSecItem()  -------------------------------------------------
' Name : OpenSecItem()
' Description : OpenItem PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSecItem()

	Dim arrRet
	Dim arrParam(5), arrField(7), arrHeader(7)
	Dim iCalledAspName

	If IsOpenPop = True  Then Exit Function

	IsOpenPop = True

	arrParam(0) = "SECǰ��"
	arrParam(1) = "(SELECT FR1_CD PAR_ITEM_CD, FR2_CD PAR_ITEM_NM, FR3_CD PAR_SPEC, TO1_CD ITEM_CD FROM J_CODE_MAPPING  WHERE MAJOR_CD = 'J0010' AND MINOR_CD = '0000' ) A, B_ITEM B"
	arrParam(2) = FilterVar(Trim(frm1.txtSecItemCd.Value),"","SNM")

	arrParam(4) = " RTRIM(A.ITEM_CD) *= B.ITEM_CD "
	arrParam(4) = arrParam(4) & " AND B.VALID_FROM_DT <=  " & FilterVar(UNIConvDate(EndDate), "''", "S") & " "
	arrParam(4) = arrParam(4) & " AND B.VALID_TO_DT   >=  " & FilterVar(UNIConvDate(EndDate), "''", "S") & " "

	arrParam(5) = "ǰ��"

	arrField(0) = "A.PAR_Item_Cd"
	arrField(1) = "A.PAR_Item_NM"
	arrField(2) = "A.PAR_SPEC"
	arrField(3) = "B.ITEM_CD"
	arrField(4) = "B.ITEM_NM"
	arrField(5) = "B.SPEC"

	arrHeader(0) = "SEC ǰ���ڵ�"
	arrHeader(1) = "SCE ǰ���"
	arrHeader(2) = "SEC SPEC"
	arrHeader(3) = "���� ǰ���ڵ�"
	arrHeader(4) = "���� ǰ���"
	arrHeader(5) = "���� SPEC"

	iCalledAspName = AskPRAspName("J2b02pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "J2b02pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField, arrHeader), "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtSecItemCd.Value   = arrRet(0)
		frm1.txtSecItemNm.Value   = arrRet(1)
		frm1.txtSecItemCd.focus
	End If

End Function

'========================================
Function OpenRequried(ByVal iRequried)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iRequried
	Case 1												
		arrParam(0) = "��������"
		arrParam(1) = "B_MINOR A, I_MOVETYPE_CONFIGURATION B"				
		arrParam(2) = Trim(frm1.txtDn_Type.value)		
		arrParam(3) = ""
		arrParam(4) = "A.MINOR_CD=B.MOV_TYPE AND (B.TRNS_TYPE = " & FilterVar("DI", "''", "S") & " OR (B.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " AND B.STCK_TYPE_FLAG_DEST = " & FilterVar("T", "''", "S") & " )) AND A.MAJOR_CD=" & FilterVar("I0001", "''", "S") & " "	
		arrParam(5) = "��������"

		arrField(0) = "A.MINOR_CD"
		arrField(1) = "A.MINOR_NM"

		arrHeader(0) = "��������"					
		arrHeader(1) = "�������¸�"
		
		frm1.txtDn_Type.focus

	Case 2												
		arrParam(0) = "��ǰó"						
		arrParam(1) = "B_BIZ_PARTNER_FTN PARTNER_FTN,B_BIZ_PARTNER PARTNER"		
		arrParam(2) = Trim(frm1.txtShip_to_party.value)							
		'arrParam(3) = Trim(frm1.txtShip_to_partyNm.value)						
		arrParam(4) = "PARTNER_FTN.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND PARTNER_FTN.PARTNER_FTN=" & FilterVar("SSH", "''", "S") & " AND PARTNER.BP_CD=PARTNER_FTN.BP_CD AND PARTNER.BP_TYPE <= " & FilterVar("CS", "''", "S") & ""
		arrParam(5) = "��ǰó"						
		
	    arrField(0) = "PARTNER_FTN.PARTNER_BP_CD"				
	    arrField(1) = "PARTNER.BP_NM"					
	    arrField(2) = "PARTNER_FTN.BP_CD"		
	    arrField(3) = "PARTNER_FTN.PARTNER_FTN"			
	    arrField(4) = "PARTNER_FTN.USAGE_FLAG"					
	    
	    arrHeader(0) = "��ǰó"						
	    arrHeader(1) = "��ǰó��"					
	    arrHeader(2) = "�ŷ�ó�ڵ�"					
	    arrHeader(3) = "�ŷ�óŸ��"					
	    arrHeader(4) = "��뿩��"	
	    
	    frm1.txtShip_to_party.focus

	Case 3												
		arrParam(0) = "��۹��"					
		arrParam(1) = "B_MINOR"							
		arrParam(2) = Trim(frm1.txtTrans_meth.value)	
		'arrParam(3) = Trim(frm1.txtTrans_meth_nm.value)	
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9009", "''", "S") & ""				
		arrParam(5) = "��۹��"					
		
	    arrField(0) = "MINOR_CD"						
	    arrField(1) = "MINOR_NM"						
	    
	    arrHeader(0) = "��۹��"					
	    arrHeader(1) = "��۹����"
	    
		frm1.txtTrans_meth.focus					

	End Select
    
	Select Case iRequried
	Case 2
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetRequried(arrRet,iRequried)
	End If	
	
End Function

'========================================
Function OpenConSalesGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����׷�"					
	arrParam(1) = "B_SALES_GRP"						
	arrParam(2) = Trim(frm1.txtSalesGrp.value)		
	arrParam(3) = ""
	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "�����׷�"					
		
	arrField(0) = "SALES_GRP"						
	arrField(1) = "SALES_GRP_NM"					
	    
	arrHeader(0) = "�����׷�"					
	arrHeader(1) = "�����׷��"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtSalesGrp.focus
	
	If arrRet(0) <> "" Then
		frm1.txtSalesGrp.value = arrRet(0)
		frm1.txtSalesGrpNm.value = arrRet(1)
	End If	

End Function

'========================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'=========================================
Function SetRequried(Byval arrRet,ByVal iRequried)

	Select Case iRequried
	Case 1
		frm1.txtDn_Type.value = arrRet(0) 
		frm1.txtDn_TypeNm.value = arrRet(1)   
	Case 2
		frm1.txtShip_to_party.value = arrRet(0) 
		frm1.txtShip_to_partyNm.value = arrRet(1)   
	Case 3
		frm1.txtTrans_meth.value = arrRet(0) 
		frm1.txtTrans_meth_nm.value = arrRet(1)   
	End Select

End Function

'=========================================
Function CookiePage(ByVal Kubun)

	On Error Resume Next

	Const CookieSplit = 4877

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

    ' Check Ryu
    If Kubun = 1 Then
		WriteCookie CookieSplit , lsDnNo
	End If

End Function

'=========================================
Sub SetQuerySpreadColor(ByVal lRow)
	Dim GCol
    With frm1

		.vspdData.ReDraw = False
		For GCol = 1  To 12
			ggoSpread.SSSetProtected GCol, lRow, .vspdData.MaxRows
		Next
		.vspdData.ReDraw = True
    End With

End Sub


'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 AND frm1.vspdData.ActiveRow > 0 Then
		If frm1.vspdData.ActiveRow = Row Then
			Call OpenSORef
		End If
	End If
End Function

'=========================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
	End If

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col			'Sort In Assending
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort In Desending
			lgSortKey = 1
		End If
		Exit Sub
	End If

	If Row < 1 Then Exit Sub

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = GetKeyPos("A",1) ' 1
	lsDnNo=frm1.vspdData.Text
  
End Sub

'=======================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'========================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'=======================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ	
    	If lgPageNo <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ����
			If CheckRunningBizProcess Then Exit Sub
			Call DisableToolBar(parent.TBC_QUERY)
			Call DbQuery
    	End If
    End If
    
End Sub

'------------------------------------------  txtProdFromDt_KeyDown ----------------------------------------
'	Name : txtSendStartDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtSendStartDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtProdToDt_KeyDown ------------------------------------------
'	Name : txtSendEndDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtSendEndDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtSendStartDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================

Sub txtSendStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSendStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtSendStartDt.Focus
    End If
End Sub

Sub txtSendEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSendEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtSendEndDt.Focus
    End If
End Sub

'------------------------------------------  txtPlanStartDt_KeyDown ----------------------------------------
'	Name : txtPlanStartDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtPlanStartDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtPlanEndDt_KeyDown ------------------------------------------
'	Name : txtPlanEndDt_KeyDown
'	Description : Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Sub txtPlanEndDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtPlanStartDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================

Sub txtPlanStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlanStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtPlanStartDt.Focus
    End If
End Sub

Sub txtPlanEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlanEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtPlanEndDt.Focus
    End If
End Sub

'��갡�����
Function Radio1_onChange()
	
	IF lgRadio = "MU" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData			
	ggoSpread.ClearSpreadData
	
	Call ggoOper.ClearField(Document, "2")

	call initVariables()
	
	lgRadio = "MU"
	
	lgBlnFlgChgValue = True
End Function

'Sample�������
Function Radio2_onChange()
	
	IF lgRadio = "SU" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData			
	ggoSpread.ClearSpreadData
	
	Call ggoOper.ClearField(Document, "2")		

	call initVariables()
	
	lgRadio = "SU"
	
	lgBlnFlgChgValue = True
End Function

'���Hold���
Function Radio3_onChange()
	
	IF lgRadio = "MH" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData			
	ggoSpread.ClearSpreadData
	
	Call ggoOper.ClearField(Document, "2")		

	call initVariables()
	
	lgRadio = "MH"
	
	lgBlnFlgChgValue = True
End Function

'Sample Hold���
Function Radio4_onChange()
	
	IF lgRadio = "SH" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData			
	ggoSpread.ClearSpreadData
	
	Call ggoOper.ClearField(Document, "2")		

	call initVariables()
	
	lgRadio = "SH"
	
	lgBlnFlgChgValue = True
End Function

'����԰�
Function Radio5_onChange()
	
	IF lgRadio = "MINV" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData			
	ggoSpread.ClearSpreadData
	
	Call ggoOper.ClearField(Document, "2")		

	call initVariables()
	
	lgRadio = "MINV"
	
	lgBlnFlgChgValue = True
End Function

'Sample�԰�
Function Radio6_onChange()
	
	IF lgRadio = "SINV" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData			
	ggoSpread.ClearSpreadData
	
	Call ggoOper.ClearField(Document, "2")		

	call initVariables()
	
	lgRadio = "SINV"
	
	lgBlnFlgChgValue = True
End Function

'������
Function Radio7_onChange()
	
	IF lgRadio = "MOUT" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData			
	ggoSpread.ClearSpreadData
	
	Call ggoOper.ClearField(Document, "2")		

	call initVariables()
	
	lgRadio = "MOUT"
	
	lgBlnFlgChgValue = True
End Function

'Sample���
Function Radio8_onChange()
	
	IF lgRadio = "SOUT" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData			
	ggoSpread.ClearSpreadData
	
	Call ggoOper.ClearField(Document, "2")		

	call initVariables()
	
	lgRadio = "SOUT"
	
	lgBlnFlgChgValue = True
End Function

'�������
Function Radio9_onChange()
	
	IF lgRadio = "VOUT" Then
		Exit Function
	ENd IF
	
	ggoSpread.Source = frm1.vspdData			
	ggoSpread.ClearSpreadData
	
	Call ggoOper.ClearField(Document, "2")		

	call initVariables()
	
	lgRadio = "VOUT"
	
	lgBlnFlgChgValue = True
End Function