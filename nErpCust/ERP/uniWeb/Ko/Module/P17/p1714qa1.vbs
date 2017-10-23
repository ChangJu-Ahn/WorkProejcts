Const BIZ_PGM_QRY_ID = "p1714qb1.asp"                          '��: �����Ͻ� ���� ASP�� 

Dim C_Seq               '���� 
Dim C_ReqTransNo        '�̰��Ƿڹ�ȣ 
Dim C_Status            '�̰����� 
Dim C_ChildItemCd       '��ǰ�� 
Dim C_ChildItemNm       '��ǰ��� 
Dim C_Spec              '�԰� 
Dim C_ReqTransDt        '�̰���û�� 
Dim C_TransDt           '�̰��� 
Dim C_ItemAcctNm        'ǰ����� 
Dim C_ProcTypeNm        '���ޱ��� 
Dim C_ChildItemBaseQty  '��ǰ����ؼ� 
Dim C_ChildBasicUnit    '��ǰ����ش��� 
Dim C_PrntItemBaseQty   '��ǰ����ؼ� 
Dim C_PrntBasicUnit     '��ǰ����ش��� 
Dim C_SafetyLT          '����L/T
Dim C_LossRate          'LOSS RATE
Dim C_SupplyFlgNm       '�����󱸺� 
Dim C_ValidFromDt       '������ 
Dim C_ValidToDt         '������ 
Dim C_ReasonNm          '���躯��ٰ� 
Dim C_ECNNo             '���躯���ȣ 
Dim C_ECNDesc           '���躯�泻�� 
Dim C_DrawingPath       '������ 
Dim C_Row				

Dim isClicked
Dim iCol
Dim iRow
Dim IsOpenPop
Dim lgStrBOMHisFlg
Dim iStrFree


'========================================================================================================
' Name : InitSpreadPosVariables()
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()

	C_Seq               = 1			'���� 
	C_ReqTransNo        = 2			'�̰��Ƿڹ�ȣ 
	C_Status            = 3			'�̰����� 
	C_ChildItemCd       = 4			'��ǰ�� 
	C_ChildItemNm       = 5			'��ǰ��� 
	C_Spec              = 6			'�԰� 
	C_ReqTransDt        = 7			'�̰���û�� 
	C_TransDt           = 8			'�̰��� 
	C_ItemAcctNm        = 9			'ǰ����� 
	C_ProcTypeNm        = 10		'���ޱ��� 
	C_ChildItemBaseQty  = 11		'��ǰ����ؼ� 
	C_ChildBasicUnit    = 12		'��ǰ����ش��� 
	C_PrntItemBaseQty   = 13		'��ǰ����ؼ� 
	C_PrntBasicUnit     = 14		'��ǰ����ش��� 
	C_SafetyLT          = 15		'����L/T
	C_LossRate          = 16		'LOSS RATE
	C_SupplyFlgNm       = 17		'�����󱸺� 
	C_ValidFromDt       = 18		'������ 
	C_ValidToDt         = 19		'������ 
	C_ReasonNm          = 20		'���躯��ٰ� 
	C_ECNNo             = 21		'���躯���ȣ 
	C_ECNDesc           = 22		'���躯�泻�� 
	C_DrawingPath       = 23		'������ 
	C_Row				= 24								

                
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'       Name : InitVariables()
'       Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    '[ Coding part ]======================================================================================
    lgStrPrevKey = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1                               '��: initializes sort direction
    '[ Coding part End ]==================================================================================
    
End Sub

Sub SetDefaultVal()
    '[ Coding part ]======================================================================================
    
    '[ Coding part End ]==================================================================================
End Sub

'=============================================== 2.2.3 InitSpreadSheet() =================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables

    '============================================================================================
    '��: Spreadsheet vspdData
    '============================================================================================
    ggoSpread.Source = frm1.vspdData
        
    With frm1.vspdData
        
		ggoSpread.Spreadinit "V20050204", , Parent.gAllowDragDropSpread
		.ReDraw = False
		.MaxCols = C_Row
		.MaxRows = 0
        
		Call GetSpreadColumnPos()
		
        ggoSpread.SSSetFloat C_Seq, 			 "����", 6, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, 1, False, "Z"
        ggoSpread.SSSetEdit	 C_ReqTransNo, 		 "�̰��Ƿڹ�ȣ", 12
        ggoSpread.SSSetEdit	 C_Status, 	 		 "�̰�����", 10
        ggoSpread.SSSetEdit	 C_ChildItemCd, 	 "��ǰ��", 20, , , 18, 2
		ggoSpread.SSSetEdit	 C_ChildItemNm, 	 "��ǰ���", 30
		ggoSpread.SSSetEdit	 C_Spec, 			 "�԰�", 30
        ggoSpread.SSSetDate	 C_ReqTransDt, 		 "�̰���û��", 11, 2, Parent.gDateFormat
        ggoSpread.SSSetDate	 C_TransDt, 		 "�̰���", 11, 2, Parent.gDateFormat
        ggoSpread.SSSetEdit	 C_ItemAcctNm, 		 "ǰ�����", 10
        ggoSpread.SSSetEdit	 C_ProcTypeNm, 		 "���ޱ���", 12
        ggoSpread.SSSetFloat C_ChildItemBaseQty, "��ǰ����ؼ�", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        ggoSpread.SSSetEdit  C_ChildBasicUnit,   "����", 6, , , 3, 2
        ggoSpread.SSSetFloat C_PrntItemBaseQty,  "��ǰ����ؼ�", 15, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
        ggoSpread.SSSetEdit  C_PrntBasicUnit, 	 "����", 6, , , 3, 2
        ggoSpread.SSSetFloat C_SafetyLT, 		 "����L/T", 10, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, 1, False, "Z"
		ggoSpread.SSSetFloat C_LossRate, 		 "Loss��", 10, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, 1, False, "Z"
		ggoSpread.SSSetEdit C_SupplyFlgNm, 	 	 "�����󱸺�", 10                        
        ggoSpread.SSSetDate  C_ValidFromDt, 	 "������", 11, 2, Parent.gDateFormat
        ggoSpread.SSSetDate  C_ValidToDt, 		 "������", 11, 2, Parent.gDateFormat
        ggoSpread.SSSetEdit  C_ReasonNm, 		 "���躯��ٰŸ�", 14                
        ggoSpread.SSSetEdit  C_ECNNo, 			 "���躯���ȣ", 18, , , 18, 2
        ggoSpread.SSSetEdit  C_ECNDesc, 		 "���躯�泻��", 30, , , 100
        ggoSpread.SSSetEdit  C_DrawingPath, 	 "������", 30, , , 100
        ggoSpread.SSSetEdit  C_Row, 			 "����", 5

        'ggoSpread.SSSetSplit2 (C_ChildItemPopUp)     'frozen ��� �߰�(?)
        
		'====================================================================================
		'���� Column�� �����ش�.
		'====================================================================================
        'Call ggoSpread.MakePairsColumn(C_PrntItemBaseQty, C_PrntBasicUnitPopup)
        '====================================================================================
		
		'====================================================================================
		'Hidden Column ���� 
		'====================================================================================
		Call ggoSpread.SSSetColHidden(C_Row, C_Row, True)
		'====================================================================================
        .ReDraw = True
                       
	End With
    
	Call SetSpreadLock
	
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
        ggoSpread.Source = frm1.vspdData
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected -1, -1
		ggoSpread.SpreadLockWithOddEvenRowColor()
		'UnLock		
'		ggoSpread.SpreadUnLock C_Remark, -1, C_Remark
		
		'�ʼ��Է¼��� 
'		ggoSpread.SSSetRequired C_ChildItemBaseQty, -1, -1
		                
		.vspdData.ReDraw = True
    End With
End Sub

'================================== 2.2.6 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc :
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow, ByVal QueryStatus)

	ggoSpread.Source = frm1.vspdData
	
	frm1.vspdData.ReDraw = False
	
	If QueryStatus = 1 Then         'When Query is OK
'		ggoSpread.SSSetProtected 	-1, pvStartRow, pvEndRow
'		ggoSpread.SpreadUnLock 		C_Remark, pvStartRow, C_Remark, pvEndRow
'		ggoSpread.SSSetRequired 	C_Seq, pvStartRow, pvEndRow
	Else
'		ggoSpread.SSSetProtected 	-1, pvStartRow, pvEndRow
'		ggoSpread.SpreadUnLock 		C_Remark, pvStartRow, C_Remark, pvEndRow
'		ggoSpread.SSSetRequired 	C_Seq, pvStartRow, pvEndRow
	End If
	
	frm1.vspdData.ReDraw = True
        
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   :
'========================================================================================
Sub GetSpreadColumnPos()
	
	Dim iCurColumnPos
   
	ggoSpread.Source = frm1.vspdData
	
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	C_Seq               = iCurColumnPos(1)		'����                   
	C_ReqTransNo        = iCurColumnPos(2)		'�̰��Ƿڹ�ȣ           
	C_Status            = iCurColumnPos(3)		'�̰�����               
	C_ChildItemCd       = iCurColumnPos(4)		'��ǰ��                 
	C_ChildItemNm       = iCurColumnPos(5)		'��ǰ���               
	C_Spec              = iCurColumnPos(6)		'�԰�                   
	C_ReqTransDt        = iCurColumnPos(7)		'�̰���û��             
	C_TransDt           = iCurColumnPos(8)		'�̰���                 
	C_ItemAcctNm        = iCurColumnPos(9)		'ǰ�����               
	C_ProcTypeNm        = iCurColumnPos(10)		'���ޱ���               
	C_ChildItemBaseQty  = iCurColumnPos(11)		'��ǰ����ؼ�           
	C_ChildBasicUnit    = iCurColumnPos(12)		'��ǰ����ش���         
	C_PrntItemBaseQty   = iCurColumnPos(13)		'��ǰ����ؼ�           
	C_PrntBasicUnit     = iCurColumnPos(14)		'��ǰ����ش���         
	C_SafetyLT          = iCurColumnPos(15)		'����L/T                
	C_LossRate          = iCurColumnPos(16)		'LOSS RATE              
	C_SupplyFlgNm       = iCurColumnPos(17)		'�����󱸺�             
	C_ValidFromDt       = iCurColumnPos(18)		'������                 
	C_ValidToDt         = iCurColumnPos(19)		'������                 
	C_ReasonNm          = iCurColumnPos(20)		'���躯��ٰ�           
	C_ECNNo             = iCurColumnPos(21)		'���躯���ȣ           
	C_ECNDesc           = iCurColumnPos(22)		'���躯�泻��           
	C_DrawingPath       = iCurColumnPos(23)		'������               
	C_Row				= iCurColumnPos(24)								

End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   :
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   :
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim iIntCnt
	
	ggoSpread.Source = gActiveSpdSheet
	
	Call ggoSpread.RestoreSpreadInf
	Call InitSpreadSheet
	
	frm1.vspdData.ReDraw = False
		Call InitComboBox
		Call ggoSpread.ReOrderingSpreadData
		Call SetSpreadColor(1, 1, 0, 1)
		
		With frm1
			.vspdData.Col = C_Row
			If .vspdData.Text <> "" Then
				For iIntCnt = 2 To .vspdData.MaxRows
					.vspdData.Col = C_HdrProcType
					.vspdData.Row = iIntCnt
					
					If UCase(Trim(.vspdData.Text)) = "O" Then
						Call SetFieldProp(iIntCnt, "D", "O")
					Else
						Call SetFieldProp(iIntCnt, "D", "P")
					End If
				Next
			End If
		End With

	frm1.vspdData.ReDraw = True

End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'       Name : InitComboBox()
'       Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
End Sub

'==========================================  2.2.6 InitData()  =======================================
'       Name : InitData()
'       Description : Combo Display
'=====================================================================================================
Sub InitData(ByVal lngStartRow)
End Sub

'------------------------------------------  OpenConBasePlant()  -----------------------------------------
'       Name : OpenConBasePlant()
'       Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConBasePlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = "���ذ����˾�"             						' �˾� ��Ī 
	arrParam(1) = "B_PLANT A, P_PLANT_CONFIGURATION B"                  ' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBasePlantCd.Value)						' Code Condition
	arrParam(3) = ""            										' Name Cindition
	arrParam(4) = "A.PLANT_CD = B.PLANT_CD AND B.ENG_BOM_FLAG = 'Y'"    ' Where Condition
	arrParam(5) = "���ذ���"											' TextBox ��Ī 
	
	arrField(0) = "A.PLANT_CD"                      					' Field��(0)
	arrField(1) = "A.PLANT_NM"                      					' Field��(1)
	
	arrHeader(0) = "����"    										' Header��(0)
	arrHeader(1) = "�����"  										' Header��(1)
	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
	        Call SetConBasePlant(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	
	frm1.txtDestPlantCd.focus

End Function


'------------------------------------------  OpenCondDestPlant()  ----------------------------------------
'       Name : OpenCondDestPlant()
'       Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConDestPlant()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = "�������˾�"             	' �˾� ��Ī 
	arrParam(1) = "B_PLANT"                      	' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtDestPlantCd.Value)	' Code Condition
	arrParam(3) = ""            					' Name Cindition
	arrParam(4) = ""            					' Where Condition
	arrParam(5) = "������"						' TextBox ��Ī 
	
	arrField(0) = "PLANT_CD"                      	' Field��(0)
	arrField(1) = "PLANT_NM"                      	' Field��(1)
	
	arrHeader(0) = "����"    					' Header��(0)
	arrHeader(1) = "�����"  					' Header��(1)
	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
	        Call SetConDestPlant(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	
    frm1.txtItemCd.focus
        
End Function


'------------------------------------------  OpenReqTransNo()  -------------------------------------------------
'       Name : OpenReqTransNo()
'       Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenReqTransNo()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strPlantCd
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	strPlantCd = Trim(frm1.txtDestPlantCd.Value)
	
	' �˾� ��Ī 
	arrParam(0) = "�̰��Ƿڹ�ȣ"
	' TABLE ��Ī 
	arrParam(1) = "P_EBOM_TO_PBOM_MASTER A, B_ITEM B, B_PLANT C"
	' Code Condition
	arrParam(2) = Trim(frm1.txtReqTransNo.Value)
	' Name Cindition
	arrParam(3) = ""
	' Where Condition
	arrParam(4) = "A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = C.PLANT_CD AND A.PLANT_CD = " & FilterVar(strPlantCd, "''", "S")
	' TextBox ��Ī 
	arrParam(5) = "�̰��Ƿڹ�ȣ"
	
	arrField(0) = "A.REQ_TRANS_NO"             	' Field��(0)
	arrField(1) = "A.PLANT_CD"                 	' Field��(1)
	arrField(2) = "C.PLANT_NM"                 	' Field��(2)
	arrField(3) = "A.ITEM_CD"                  	' Field��(3)
	arrField(4) = "B.ITEM_NM"                  	' Field��(4)
	arrField(5) = "dbo.ufn_GetCodeName('Y4001', A.STATUS) STATUS"                   	' Field��(5)
	
	arrHeader(0) = "�̰��Ƿڹ�ȣ"          	' Header��(0)
	arrHeader(1) = "������"              	' Header��(1)
	arrHeader(2) = "�������"            	' Header��(2)
	arrHeader(3) = "ǰ��"                  	' Header��(3)
	arrHeader(4) = "ǰ���"                	' Header��(4)
	arrHeader(5) = "�̰�����"              	' Header��(5)
	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetReqTransNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	
	frm1.txtReqTransNo.focus
        
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'       Name : OpenItemCd()
'       Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd(ByVal str, ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(11)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtDestPlantCd.Value = "" Then
		Call DisplayMsgBox("971012", "X", frm1.txtDestPlantCd.Alt, "X")
		frm1.txtDestPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtDestPlantCd.Value)   ' Plant Code
	arrParam(1) = Trim(str) ' Item Code
	
	arrField(0) = 1         'ITEM_CD
	arrField(1) = 2     	'ITEM_NM
	arrField(2) = 5         'ITEM_ACCT
	arrField(3) = 9     	'PROC_TYPE
	arrField(4) = 4     	'BASIC_UNIT
	arrField(5) = 51    	'SINGLE_ROUT_FLG
	arrField(6) = 52    	'Major_Work_Center
	arrField(7) = 13    	'Phantom_flg
	arrField(8) = 18    	'valid_from_dt
	arrField(9) = 19    	'valid_to_dt
	arrField(10) = 3    	'Field��(1) : "SPECIFICATION"
	
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
	IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "B1B11PA4", "X")
	IsOpenPop = False
	Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.Parent, arrParam, arrField), _
										"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet, iPos)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus
        
End Function

'------------------------------------------  SetConBasePlant()  ----------------------------------------------
'       Name : SetConBasePlant()
'       Description : Condition Base Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConBasePlant(ByVal arrRet)
	frm1.txtBasePlantCd.Value = arrRet(0)
	frm1.txtBasePlantNm.Value = arrRet(1)
	
	Call txtBasePlantCd_OnChange
End Function

'------------------------------------------  SetConDestPlant()  ----------------------------------------------
'       Name : SetConDestPlant()
'       Description : Condition Destination Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConDestPlant(ByVal arrRet)
	frm1.txtDestPlantCd.Value = arrRet(0)
	frm1.txtDestPlantNm.Value = arrRet(1)
	
	Call txtDestPlantCd_OnChange
End Function

'------------------------------------------  SetItemCd()  ----------------------------------------------
'       Name : SetItemCd()
'       Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(ByVal arrRet, ByVal iPos)
	frm1.txtItemCd.Value = arrRet(0)
	frm1.txtItemNm.Value = arrRet(1)
End Function

'------------------------------------------  SetReqTransNo()  --------------------------------------------------
'       Name : SetReqTransNo()
'       Description : SetReqTransNo
'---------------------------------------------------------------------------------------------------------
Function SetReqTransNo(ByVal arrRet)
    frm1.txtReqTransNo.Value 	= arrRet(0)
    frm1.txtDestPlantCd.Value 	= arrRet(1)
    frm1.txtDestPlantNm.Value 	= arrRet(2)
    frm1.txtItemCd.Value 		= arrRet(3)
    frm1.txtItemNm.Value 		= arrRet(4)
End Function

'==========================================================================================
'   Function Name :SetFieldProp
'   Function Desc :���� Case�� ���� Field���� �Ӽ��� �����Ѵ�.
'==========================================================================================
Function SetFieldProp(ByVal lRow, ByVal Level, ByVal ProcType)

End Function


'==========================================================================================
'   Function Name :LookUpItemByPlant
'   Function Desc :������ ǰ���� Item Acct�� �д´�.
'==========================================================================================
Sub LookUpItemByPlant(ByVal strItemCd, ByVal iRow)

	Err.Clear                                                                                                                   '��: Protect system from crashing
	
	Dim strSelect
	If strItemCd = "" Then Exit Sub
	
	frm1.vspdData.Col = C_ChildItemCd
	frm1.vspdData.Row = iRow
	
	strSelect = " b.ITEM_NM, a.ITEM_ACCT, dbo.ufn_GetCodeName(" & FilterVar("P1001", "''", "S") & ", a.ITEM_ACCT) ITEM_ACCT_NM, a.PROCUR_TYPE, dbo.ufn_GetCodeName(" & FilterVar("P1003", "''", "S") & ", a.PROCUR_TYPE) PROCUR_TYPE_NM, b.SPEC, b.BASIC_UNIT, dbo.ufn_GetItemAcctGrp(a.ITEM_ACCT) ITEM_ACCT_GRP "
	
	If CommonQueryRs2by2(strSelect, " B_ITEM_BY_PLANT a, B_ITEM b ", " a.ITEM_CD = b.ITEM_CD AND a.PLANT_CD = " & _
								FilterVar(frm1.txtDestPlantCd.Value, "''", "S") & " AND a.ITEM_CD = " & FilterVar(strItemCd, "''", "S"), lgF0) = False Then
		Call DisplayMsgBox("122700", "X", strItemCd, "X")
		Call LookUpItemByPlantNotOk
		Exit Sub
	End If
	
	lgF0 = Split(lgF0, Chr(11))
	
	Call LookUpItemByPlantOk(lgF0(1), lgF0(2), lgF0(3), lgF0(4), lgF0(5), lgF0(6), lgF0(7), iRow, lgF0(8))
End Sub

'==========================================================================================
'   Function Name :LookUpItemByPlantOk
'   Function Desc :������ ǰ���� ���翩�θ� Check�Ը� �д´�.
'==========================================================================================
Function LookUpItemByPlantOk(ByVal strItemNm, ByVal strItemAcct, ByVal strItemAcctNm, ByVal strProcType, ByVal strProcTypeNm, ByVal strSpec, ByVal strBasicUnit, ByVal iRow, ByVal strItemAcctGrp)
End Function

Function LookUpItemByPlantNotOk()
End Function

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
    Call SetPopupMenuItemInf("0000110111")
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If
		 Exit Sub     
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 	
	

End Sub	


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc :
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc :
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos()
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
        If Button = 2 And gMouseClickStatus = "SPC" Then
                gMouseClickStatus = "SPCR"
        End If
End Sub

'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : This event is spread sheet data changed jslee
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData

                If Row >= NewRow Then
                    Exit Sub
                End If
        '----------  Coding part  -------------------------------------------------------------
    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	'----------  Coding part  -------------------------------------------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
	
		If lgStrPrevKeyIndex <> "" Then            '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar
				Exit Sub
			End If
		End If
	End If
End Sub

Sub txtBasePlantCd_OnChange()
	Dim arrVal
	
	If Trim(frm1.txtBasePlantCd.Value) <> "" Then
		If CommonQueryRs("PLANT_NM", "B_PLANT", "PLANT_CD = " & FilterVar(frm1.txtBasePlantCd.Value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			arrVal = Split(lgF0, Chr(11))
			frm1.txtBasePlantNm.Value = Trim(arrVal(0)) 	
		Else
			frm1.txtBasePlantNm.Value = ""
		End If		
	End If
End Sub
 
Sub txtDestPlantCd_OnChange()
	Dim arrVal
	
	If Trim(frm1.txtDestPlantCd.Value) <> "" Then
		If CommonQueryRs("PLANT_NM", "B_PLANT", "PLANT_CD = " & FilterVar(frm1.txtDestPlantCd.Value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			arrVal = Split(lgF0, Chr(11))
			frm1.txtDestPlantNm.Value = Trim(arrVal(0)) 	
		Else
			frm1.txtDestPlantNm.Value = ""
		End If		
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
    Dim IntRetCD
    
    FncQuery = False                                                        '��: Processing is NG
    Err.Clear                                                               '��: Protect system from crashing

    '�������� �ʱ�ȭ 
    Call ggoSpread.ClearSpreadData
    
    Call InitVariables                                                                                                                  '��: Initializes local global variables
                                                                                                                                                        
    '-----------------------
    'Check condition area
    'TAG = '12'�� ������Ʈ�� ���� �ִ��� üũ 
    '-----------------------
    If Not chkField(document, "1") Then                                                                                 '��: This function check indispensable field
       Exit Function
    End If

	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False Then                                                                                             '��: Query db data (����BOM)
		Exit Function
	End If
	
	FncQuery = True                                                                                                                             '��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew()

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete()

End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave()

End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()
    
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()
	Dim strLevel, strChildLevel
	Dim TempChildLevel
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_Level
		strLevel = CLng(Replace(.Text, ".", ""))
		
		Do
			ggoSpread.EditUndo
			
			If .MaxRows = 0 Then Exit Do
			
			.Col = C_Level
			.Row = .ActiveRow
			If Trim(.Text) = "" Then
			    strChildLevel = CLng(0)
			Else
			    strChildLevel = CLng(Replace(Trim(.Text), ".", ""))
			End If
		Loop While (strLevel < strChildLevel)
	End With
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow()
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()
   Call Parent.FncPrint
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'========================================================================================
Function FncExcel()
    Call Parent.FncExport(Parent.C_SINGLEMULTI)                                                 '��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
Function FncFind()
    Call Parent.FncFind(Parent.C_SINGLEMULTI, False)                       '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc :
'========================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit (gActiveSpdSheet.ActiveCol)
    
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================
Function FncExit()
    Dim IntRetCD
    FncExit = False
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
	    IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                  '��: "Will you destory previous data"
	    If IntRetCD = vbNo Then
            Exit Function
	    End If
    End If
    FncExit = True
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display(��ܺ�)
'========================================================================================
Function DbQuery()

    Dim LngLastRow
    Dim LngMaxRow
    Dim LngRow
    Dim strTemp
    Dim StrNextKey
    Dim iStrBasePlantCd, iStrDestPlantCd, iStrItemCd, iStrReqTransNo
    Dim strQueryType

    DbQuery = False

    LayerShowHide (1)
                
    Err.Clear                                                               '��: Protect system from crashing

    Dim strVal

    iStrBasePlantCd	= UCase(Trim(frm1.txtBasePlantCd.Value))
    iStrDestPlantCd	= UCase(Trim(frm1.txtDestPlantCd.Value))
    iStrItemCd 		= UCase(Trim(frm1.txtItemCd.Value))
    iStrReqTransNo 	= UCase(Trim(frm1.txtReqTransNo.Value))
    
    With frm1
        strVal = BIZ_PGM_QRY_ID & "?txtBasePlantCd=" & iStrBasePlantCd      '��: ��ȸ ���� ����Ÿ 
        strVal = strVal & "&txtDestPlantCd=" & iStrDestPlantCd              '��: ��ȸ ���� ����Ÿ 
        strVal = strVal & "&txtItemCd=" & iStrItemCd                        '��: ��ȸ ���� ����Ÿ 
        strVal = strVal & "&txtReqTransNo=" & iStrReqTransNo                '��: ��ȸ ���� ����Ÿ 
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        strVal = strVal & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex         '��: Next key tag
  
		Call RunMyBizASP(MyBizASP, strVal)                                                                      '��: �����Ͻ� ASP �� ���� 

    End With
    
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(LngMaxRow)                                                                           '��: ��ȸ ������ ������� 
        
    Dim lRow
    Dim i
    '-----------------------
    'Reset variables area
    '-----------------------
    If frm1.vspdData.MaxRows > 0 Then
        lgIntFlgMode = Parent.OPMD_UMODE                                                                '��: Indicates that current mode is Update mode
    End If

    frm1.vspdData.ReDraw = False

        With frm1
        
			.vspdData.Col = C_Row
			
			If .vspdData.Text <> "" Then
			
				For i = LngMaxRow To frm1.vspdData.MaxRows
					frm1.vspdData.Col = C_HdrProcType
					frm1.vspdData.Row = i
					
					If UCase(Trim(frm1.vspdData.Text)) = "O" Then
					    Call SetFieldProp(i, "D", "O")
					Else
					    Call SetFieldProp(i, "D", "P")
					End If
				Next
			
			End If
                        
        End With
        
        
        frm1.vspdData.focus
        lgBlnFlgChgValue = False
	
	frm1.vspdData.ReDraw = True

End Function
        
Function DbQueryNotOk()
         
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave()
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()                                                                                                     '��: ���� ������ ���� ���� 

End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()                                                                                           '��: ���� ������ ���� ���� 
    Call InitVariables
    Call FncNew
End Function

'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : ������, �������� ������ HTML ��ü(TEXTAREA)�� Clear���� �ش�.
'========================================================================================
Function RemovedivTextArea()
	
	Dim ii
	        
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild (divTextArea.children(0))
	Next

End Function

'========================================================================================
' Function Name : CheckPlant
' Function Desc : ����Configuration�� ����������� ������ �Ǿ����� Check
'========================================================================================
Function CheckPlant(ByVal sPlantCd)	
														
    Err.Clear																

    CheckPlant = False
    
	Dim arrVal, strWhere, strFrom

	If Trim(sPlantCd) <> "" Then
	
		strFrom = "B_PLANT A, P_PLANT_CONFIGURATION B"
		strWhere = 				" A.PLANT_CD = B.PLANT_CD AND B.ENG_BOM_FLAG = 'Y' AND"
		strWhere = strWhere & 	" A.PLANT_CD = " & FilterVar(sPlantCd, "''", "S")

		If Not CommonQueryRs("A.PLANT_NM", strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
    		Exit Function
		End If
	End If

	CheckPlant = True
	
End Function
