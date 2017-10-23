Const BIZ_PGM_QRY_ID					= "p1711mb1.asp"		'��: ��ü ��ȸ 
Const BIZ_PGM_HDRSAVE_ID				= "p1711mb4.asp"		'��: ��ǰ�� ����(�ֻ��� ���)
Const BIZ_PGM_DTLSAVE_ID				= "p1711mb5.asp"		'��: ��ǰ�� �Է�,����,���� 
Const BIZ_PGM_HDRDEL_ID					= "p1711mb6.asp"		'��: ��ü BOM ���� 

Const C_Sep  = "/"
Const C_PROD  = "PROD"
Const C_MATL  = "MATL"
Const C_PHANTOM ="PHANTOM"
Const C_ASSEMBLY = "ASSEMBLY"
Const C_SUBCON  = "SUBCON"

Const C_IMG_PROD = "../../../CShared/image/product.gif"
Const C_IMG_MATL = "../../../CShared/image/material.gif"
Const C_IMG_PHANTOM = "../../../CShared/image/phantom.gif"
Const C_IMG_ASSEMBLY = "../../../CShared/image/Assembly.gif"
Const C_IMG_SUBCON = "../../../CShared/image/subcon.gif"


Const tvwChild = 4

Const C_MNU_OPEN	= 0
Const C_MNU_ADD		= 1
Const C_MNU_DELETE	= 2
'Const C_MNU_RENAME	= 3

Const C_NEW_FOLDER = "������"
Const C_NEW_FOLDER_KEY = "COMPONENT"	

Dim lgBlnFlgConChg				'��: Condition ���� Flag
Dim lgNextNo					'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo					' ""

Dim IsOpenPop
Dim lgBlnBizLoadMenu
Dim	lgSelNode

Dim lgQueryMode
Dim lgNodeClick

Dim lgClkInsrtRow
Dim lgClkCopy

Dim lgRdoOldVal
Dim lgRdoOldVal2
Dim lgRdoOldVal3

Dim lgBomType
Dim lgBomTypeNm
Dim lgStrBOMHisFlg
Dim lgStrHeaderFlg

'--------------------------------------------------------------------------------------------------------
'						Mode										   FieldProp
'	    Form_Load		: 0(C, )  - �ֻ��� BOM �ű� ���	 
'		FncNew			: 0(C, )  - �ֻ��� BOM �ű� ���	 
'		FncInsertRow	: X
'		DBQueryOK		: 6(U, )  - �ֻ��� ǰ�� ���� ���� 
'		DBQueryNotOK	: 0(C, )  - �ֻ��� ǰ�� �ű� ��� ���� 
'		FncCopy			: 7(M, )  - �ֻ��� ǰ�� 
'						: 8(M,U)  - ����ǰ 
'		LookUpItemOk	: 3( ,C)  - ��ǰ/����ǰ , ��ǰ�� �߰��� 
'						: 3( ,C)  - ������		, ��ǰ�� �߰��� 
'						: x		  - �ֻ��� ǰ�� �ű� ��Ͻ� 
'		LookUpBomOk		: 2(U,C)  - ��ǰ�� �߰��� 
'						: 6(U, )  - �ֻ��� BOM �ű� ��Ͻ� 
'		LookUpBomNotOk	: 1(C,C)  - ��ǰ�� �߰��� 
'						: 0(C, )  - �ֻ��� BOM �ű� ��Ͻ� 
'		LookUpHdrOk		: 6(U, )  - �ֻ��� BOM �ű� ��Ͻ� 
'		LookUpHdrNotOk	:
'		LookUpDtlOk		: 4(U,U)
'						  5( ,U)
'
'		LookUpDtlNotOk	:
'-----------------------------------------------------------------------------------------------------------
	
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'==================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE			'��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False			'��: Indicates that no value changed
    lgIntGrpCount = 0					'��: Initializes Group View Size
    
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
    lgBlnBizLoadMenu = False
   
    lgSelNode = ""
    lgClkInsrtRow = False
    lgClkCopy = False
	lgQueryMode = False
	lgNodeClick = False
End Sub

'========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=================================================================================================== 
Sub SetDefaultVal()
	
	'------------------------
	' Default Date Setting
	'------------------------ 
	With frm1
		.txtValidFromDt.Value	= ""
		.txtValidToDt.Value		= ""	
		
		.txtBaseDt.Text = StartDate	'2003-09-13
		
		If .txtBomNo.value = "" Then
			.txtBomNo.value = "E"
		End If

		If .txtBomNo1.value = "" Then
			.txtBomNo1.value = "E"
		End If

		If .hBomType.value = "" Then
			.hBomType.value = "E"
		End If
						
		.cboItemAcct.value = ""
		
	End With
		
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtBaseDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBaseDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtBaseDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtBaseDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtBaseDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"													' �˾� ��Ī 
	arrParam(1) = "B_PLANT A, P_PLANT_CONFIGURATION B"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)									' Code Condition
	arrParam(3) = ""															' Name Cindition
	arrParam(4) = "A.PLANT_CD = B.PLANT_CD AND B.ENG_BOM_FLAG = 'Y'"			' Where Condition
	arrParam(5) = "����"														' TextBox ��Ī 
	
    arrField(0) = "A.PLANT_CD"													' Field��(0)
    arrField(1) = "A.PLANT_NM"													' Field��(1)
    
    arrHeader(0) = "����"													' Header��(0)
    arrHeader(1) = "�����"													' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenBomNo()  -------------------------------------------------
'	Name : OpenBomNo()
'	Description : Condition BomNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBomNo(ByVal iPos)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6) 
   '---------------------------------------------
	 ' Validation Check Area
	 '--------------------------------------------- 
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	If iPos = 0 Then
		If frm1.txtItemCd.value = "" Then
			Call DisplayMsgBox("971012", "X", "ǰ��", "X")
			frm1.txtItemCd.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If		
	Else
		If frm1.txtItemCd1.value = "" Then
			Call DisplayMsgBox("971012", "X", "��ǰ��", "X")
			frm1.txtItemCd1.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If		
	End If
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If iPos = 1 Then
		If UCase(frm1.txtBomNo1.className) = UCase(parent.UCN_PROTECTED) Then
			Exit Function
		End If
	End If
	   
   '---------------------------------------------
	 ' Parameter Setting
	 '--------------------------------------------- 

	IsOpenPop = True

	arrParam(0) = "BOM�˾�"						' �˾� ��Ī 
	arrParam(1) = "B_MINOR"							' TABLE ��Ī 
	
	If iPos = 0 Then
		arrParam(2) = Trim(frm1.txtBomNo.value)		' Code Condition
	Else				   
		arrParam(2) = Trim(frm1.txtBomNo1.Value)	' Code Condition	
	End If
	
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1401", "''", "S") & " "
	
	arrParam(5) = "BOM Type"						' TextBox ��Ī 
	
    arrField(0) = "MINOR_CD"						' Field��(0)
    arrField(1) = "MINOR_NM"						' Field��(1)
        
    arrHeader(0) = "BOM Type"					' Header��(0)
    arrHeader(1) = "BOM Ư��"					' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBomNo(arrRet,iPos)
	End If	
	
	If iPos = 0 Then
		Call SetFocusToDocument("M")
		frm1.txtBomNo.focus			
	Else
		Call SetFocusToDocument("M")
		frm1.txtBomNo1.focus
	End If
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd(ByVal iVal)

	Dim arrRet
	Dim arrParam(6), arrField(11)
	Dim iCalledAspName, IntRetCD
		
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If

	If iVal = 1 Then
		If UCase(frm1.txtItemCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" or CheckPlant(frm1.txtPlantCd.value) = False Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)								' Plant Code

	If iVal = 0 Then
		arrParam(1) = Trim(frm1.txtItemCd.value)							' Item Code
	ElseIf iVal =1 Then
		arrParam(1) = Trim(frm1.txtItemCd1.value)						' Item Code
		arrParam(5) = " AND B.VALID_FLG = 'Y' "
	End If
		
    arrField(0) = 1							'ITEM_CD
    arrField(1) = 2 						'ITEM_NM											
    arrField(2) = 5							'ITEM_ACCT
    arrField(3) = 9 						'PROC_TYPE
    arrField(4) = 4 						'BASIC_UNIT
    arrField(5) = 51						'SINGLE_ROUT_FLG
    arrField(6) = 52						'Major_Work_Center
    arrField(7) = 13						'Phantom_flg
    arrField(8) = 18						'valid_from_dt
    arrField(9) = 19						'valid_to_dt
    arrField(10) = 3						'Spec
    arrField(11) = 54						'Item_Acct_Grp
    
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet,iVal)
	End If
			
	If iVal = 0 Then
		Call SetFocusToDocument("M")
		frm1.txtItemCd.focus			
	Else
		Call SetFocusToDocument("M")
		frm1.txtItemCd1.focus
	End IF

frm1.txtBomNo1.value = "E"
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : Unit PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenUnit(ByVal iVal)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	
	If iVal = 0 Then
		If UCase(frm1.txtChildItemUnit.className) = UCase(parent.UCN_PROTECTED) Then 
			IsOpenPop = False
			Exit Function
		End If		
		arrParam(2) = Trim(frm1.txtChildItemUnit.Value)
	Else
		If UCase(frm1.txtPrntItemUnit.className) = UCase(parent.UCN_PROTECTED) Then 
			IsOpenPop = False
			Exit Function
		End If		         		
		arrParam(2) = Trim(frm1.txtPrntItemUnit.Value)
	End If
	
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "������"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetUnit(arrRet,iVal)
	End If	
	
	If iVal = 0 Then
		Call SetFocusToDocument("M")
		frm1.txtChildItemUnit.focus			
	Else
		Call SetFocusToDocument("M")
		frm1.txtPrntItemUnit.focus
	End IF
	
End Function

'------------------------------------------  OpenECNNo()  -------------------------------------------------
'	Name : OpenECNNo()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenECNNo(ByVal strECNNo)
	Dim arrRet
	Dim arrParam(4), arrField(10)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	If UCase(frm1.txtECNNo1.className) = UCase(parent.UCN_PROTECTED) Then 
		IsOpenPop = False
		Exit Function
	End If		
	
	arrParam(0) = Trim(strECNNo)   ' ECN No.

	iCalledAspName = AskPRAspName("P1410PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P1410PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetECNNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtECNNo1.focus
	
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenReasonCd()
'	Description : OpenReasonCd
'--------------------------------------------------------------------------------------------------------- 
Function OpenReasonCd(ByVal str)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If UCase(frm1.txtReasonCd1.className) = UCase(parent.UCN_PROTECTED) Then 
		IsOpenPop = False
		Exit Function
	End If		

	arrParam(0) = "����ٰ��˾�"
	arrParam(1) = "B_MINOR"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1402", "''", "S") & ""			
	arrParam(5) = "����ٰ�"			
	
    arrField(0) = "MINOR_CD"	
    arrField(1) = "MINOR_NM"	
   
    
    arrHeader(0) = "����ٰ�"		
    arrHeader(1) = "����ٰŸ�"		
    
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetReason(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtReasonCd1.focus
	
End Function

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(byval arrRet,ByVal iVal)
	If iVal = 0 Then
		frm1.txtItemCd.value = UCase(Trim(arrRet(0)))
		frm1.txtItemNm.value =	Trim(arrRet(1))
	Else
		Call InitFieldData()
		
		frm1.txtItemCd1.value = UCase(Trim(arrRet(0)))
		frm1.txtItemNm1.value =	Trim(arrRet(1))
		
		'--- ǰ����� ---
		
		frm1.cboItemAcct.value = Trim(arrRet(2))
		
		frm1.txtProcType.value = UCase(Trim(arrRet(3)))
		'--- �ܰ��� ���� ---

		frm1.txtPlantItemFromDt.Text = arrRet(8)
		frm1.txtPlantItemToDt.Text = arrRet(9)
		frm1.txtSpec.value = arrRet(10)
		frm1.txtItemAcctGrp.value = arrRet(11)
		
		If lgClkInsrtRow = True Then
			'frm1.txtChildItemQty.value = "=UniNumClientFormat(1,ggQty.DecPoint,0)"
			frm1.txtChildItemQty.Text = "1" & parent.gComNumDec & "0000"
			frm1.txtChildItemUnit.value = UCase(arrRet(4))
		End If
		
		Call LookUpItemByPlantOk()
		
		lgBlnFlgChgValue = True
		
	End If
		
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	
	Call txtPlantCd_OnChange()
	
	frm1.txtPlantCd.focus 
	Set gActiveElement = document.activeElement 		
End Function

'------------------------------------------  SetBomNo()  --------------------------------------------------
'	Name : SetBomNo()
'	Description : Bom No Popup���� return�� �� 
'--------------------------------------------------------------------------------------------------------- 

Function SetBomNo(byval arrRet,ByVal iPos)

	If iPos = 0 Then
		frm1.txtBomNo.Value    = arrRet(0)		
	Else
		frm1.txtBomNo1.Value    = arrRet(0)

		Call LookUpBomNoForChild()
		
		lgBlnFlgChgValue = True
	End If

End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Unit Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetUnit(byval arrRet,ByVal iVal)
	If iVal = 0 Then
		frm1.txtChildItemUnit.Value    = UCase(arrRet(0))		
	Else
		frm1.txtPrntItemUnit.Value    = UCase(arrRet(0))				
	End If
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Unit Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetECNNo(ByVal arrRet)
	frm1.txtECNNo1.Value	= arrRet(0)
	frm1.txtECNDesc1.Value	= arrRet(1)
	frm1.txtReasonCd1.Value	= arrRet(2)
	frm1.txtReasonNm1.Value	= arrRet(3)		

'    If lgStrBOMHisFlg = "Y" Then
'		Call ggoOper.SetReqAttr(frm1.txtECNNo1, "N")
'	Else
'		Call ggoOper.SetReqAttr(frm1.txtECNNo1, "Q")
'	End If

	Call ggoOper.SetReqAttr(frm1.txtECNDesc1, "Q")
	Call ggoOper.SetReqAttr(frm1.txtReasonCd1, "Q")
	
	frm1.txtPlantCd.focus 
	Set gActiveElement = document.activeElement 
	
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetReason()
'	Description : SetReason
'--------------------------------------------------------------------------------------------------------- 
Function SetReason(Byval arrRet)
	frm1.txtReasonCd1.Value    = arrRet(0)		
	frm1.txtReasonNm1.Value    = arrRet(1)	
	
	frm1.txtPlantCd.focus 
	Set gActiveElement = document.activeElement	
End Function

'==========================================================================================
'   Function Name :LookUpBomNoForChild
'   Function Desc :������ ǰ���� BOM�� �����ϴ� �� üũ 
'==========================================================================================
Sub LookUpBomNoForChild()
    
    If gLookUpEnable = False Then Exit Sub
    
    Dim strVal
    
    Err.Clear															'��: Protect system from crashing
	
	'--------------------------------------
	' ��ǰ���� �����簡 �ƴϰ� BOM ���簡 �ƴҶ� 
	'--------------------------------------
	If Trim(frm1.txtBomNo1.value) <> "" And lgClkCopy = False Then
		If (Trim(frm1.hBomType.value) = "1" And Trim(frm1.txtBomNo1.value) <> "1") _ 
			Or (Trim(frm1.hBomType.value) = "E" And Trim(frm1.txtBomNo1.value) <> "E") Then
				Call DisplayMsgBox("182621", "X", "X", "X")		
				frm1.txtBomNo1.focus
				Set gActiveElement = document.activeElement 
				Exit Sub
		End If
	End If
	
	'---------------------------------
	'Query Itey By Plant
	'---------------------------------
	<!-- frm1.txtBomNo1.value = "E"	-->
	
	LayerShowHide(1)
			
    strVal = BIZ_PGM_QRY_ID & "?txtValidToDt1=" & parent.UID_M0001	'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&QueryType=" & "B"									'��: LookUP ���� ����Ÿ 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd1.value)					'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtBomNo=" & Trim(frm1.txtBomNo1.value)
	strVal = strVal & "&CurDate=" & UniConvYYYYMMDDToDate(parent.gDateFormat, "1900","01","01")
											   
    Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

End Sub

'==========================================================================================
'   Function Name :LookUpBomNoForChildOk
'   Function Desc :������ ǰ���� BOM�� �����ϴ� �� üũ 
'==========================================================================================
Sub LookUpChildBomNoOk()

    If lgClkInsrtRow = True Then
		Call SetModChange(2)				'Header:U Detail:C 
		frm1.txtItemSeq.focus
		Set gActiveElement = document.activeElement  
	Else
		If lgClkCopy <> True Then
			Call SetModChange(6)			'Header:U Detail:  		
		End If
		frm1.txtBOMDesc.focus
		Set gActiveElement = document.activeElement  
	End If
	
	Call SetFieldProp(44)					'Header Update 
	
End Sub

'==========================================================================================
'   Function Name :LookUpBomNoForChildNotOk
'   Function Desc :��ǰ���� ��� BOM�� �����ϴ� �� üũ 
'==========================================================================================
Sub LookUpChildBomNoNotOk()
	
	If Trim(frm1.hBomType.value) <> "" Then
		If Trim(frm1.hBomType.value) = "1" or Trim(frm1.hBomType.value) = "E" Then
			Call SetFieldProp(64)					'Header Create 
		Else
			Call SetFieldProp(54)					'Header Create 
		End IF 
	Else
		Call SetFieldProp(54)					'Header Create 
	End If	
	
	If lgClkInsrtRow = True Then
		Call SetModChange(1)				'Header:C Detail:C 
	Else
		If lgClkCopy <> True Then
			Call SetModChange(0)			'Header:C Detail:  
		End If
	End IF
	
	frm1.txtValidFromDt.Value = StartDate
	frm1.txtValidToDt.Value = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
	
End Sub

'==========================================================================================
'   Function Name :LookUpItemByPlant
'   Function Desc :������ ǰ���� Item Acct�� �д´�.
'==========================================================================================
Sub LookUpItemByPlant()
    
    If gLookUpEnable = False Then Exit Sub
    
    Err.Clear															'��: Protect system from crashing
    
    Dim strVal

	LayerShowHide(1)
				
    strVal = BIZ_PGM_QRY_ID & "?txtValidToDt1=" & parent.UID_M0001		'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&QueryType=" & "I"									'��: LookUP ���� ����Ÿ 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd1.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtPrntItemCd=" & Trim(frm1.txtPrntItemCd.value)		'��: ��ȸ ���� ����Ÿ 
    
    If lgClkInsrtRow = True Then
		strVal = strVal & "&CurPos=" & "1"								'��: ��ȸ ���� ����Ÿ 
	Else
		strVal = strVal & "&CurPos=" & "0"								'��: ��ȸ ���� ����Ÿ 
	End If    
	strVal = strVal & "&CurDate=" & StartDate
    
    Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

End Sub

Function LookUpItemByPlantOk()
	
	If lgClkInsrtRow = True Then					'��ǰ���߰��� üũ ���� 
		If UCase(Trim(frm1.txtItemCd1.value)) = UCase(Trim(frm1.txtPrntItemCd.value)) Then
			Call DisplayMsgBox("127421", "X", "��ǰ��", "��ǰ��")
			frm1.txtItemCd1.Value = "" 
			frm1.txtItemCd1.focus
			Set gActiveElement = document.activeElement 
			
			Exit Function
		End If
		
		If Trim(frm1.hBomType.value) <> "" Then
			frm1.txtBomNo1.value = frm1.hBomType.value
		End If
		
		If Trim(frm1.txtItemAcctGrp.value) = "1FINAL" Or Trim(frm1.txtItemAcctGrp.value) = "2SEMI" Or frm1.hBomType.value = "E" Then
			
			Call SetModChange(3)					'Header:       Detail:C 	
			
			If Trim(frm1.hBomType.value) <> "" Then
				If Trim(frm1.hBomType.value) = "1" Then		'��ǰ�� BOM type�� �����ϰ� ��ǰ�� BOM type���� 
					Call SetFieldProp(64)					'Header:Create Detail:Create 
					frm1.txtBomNo1.value = 1				
					Call LookUpBomNoForChild()
				Else
					Call SetFieldProp(64)					'Header:Create Detail:Create 
					frm1.txtBomNo1.value = "E"

					'Call LookUpBomNoForChild()
				End If
			Else
			'If Trim(frm1.hBomType.value) = "" Then
				Call SetFieldProp(54)
				frm1.txtBomNo1.focus
				Set gActiveElement = document.activeElement 
			End If
			 
		ElseIf Trim(frm1.txtItemAcctGrp.value) = "3RAW" Or Trim(frm1.txtItemAcctGrp.value) = "4SUB" Then
			Call SetFieldProp(24)					'Header:	     Detail:Create 
			Call SetModChange(3)					'Header:         Detail:C 
			 
			frm1.txtItemSeq.focus
			Set gActiveElement = document.activeElement  
		Else
			Call DisplayMsgBox("182720", "X", "X", "X")
			
			frm1.txtItemCd1.Focus
			Set gActiveElement = document.activeElement 
			
			Exit Function 
		End If

		If Trim(frm1.hPrntProcType.value) = "O" Then
			Call ggoOper.SetReqAttr(frm1.rdoSupplyFlg1, "N")
			Call ggoOper.SetReqAttr(frm1.rdoSupplyFlg2, "N")
		Else
			Call ggoOper.SetReqAttr(frm1.rdoSupplyFlg1, "Q")
			Call ggoOper.SetReqAttr(frm1.rdoSupplyFlg2, "Q")
			frm1.rdoSupplyFlg1.checked = True
			lgRdoOldVal2 = 1
		End If

		If lgStrBOMHisFlg = "Y" Then
			Call ggoOper.SetReqAttr(frm1.txtECNNo1, "N")
		Else
			Call ggoOper.SetReqAttr(frm1.txtECNNo1, "Q")
		End If
				
	Else											'�űԳ� BOM����� üũ ���� 
		If Trim(frm1.txtProcType.value) = "P" _
		Or Trim(frm1.txtItemAcctGrp.value) = "3RAW" Or Trim(frm1.txtItemAcctGrp.value) = "4SUB"  _
		Or Trim(frm1.txtItemAcctGrp.value) = "5GOOD" Or Trim(frm1.txtItemAcctGrp.value) = "6MRO" Then
			Call DisplayMsgBox("182618", "X", "X", "X")
			frm1.txtItemCd1.focus
			Set gActiveElement = document.activeElement  
			Exit Function 
		End If
	End If
End Function

Function LookUpItemByPlantNotOk()
	Call InitFieldData()	
	frm1.txtItemCd1.focus 
	Set gActiveElement = document.activeElement 	
End Function

Function LookUpBOMHdrExist(ByVal PlantCd, ByVal ItemCd, ByVal BOMNo)
	Dim iStrWhereSQL
	
	iStrWhereSQL = "PLANT_CD = " & FilterVar(PlantCd, "''", "S") & " AND ITEM_CD = " & FilterVar(ItemCd, "''", "S") & " AND BOM_NO = " & FilterVar(BOMNo, "''", "S")
	Call CommonQueryRs("ITEM_CD", "P_BOM_HEADER", iStrWhereSQL, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	If Trim(lgF0) <> "" Then
		LookUpBOMHdrExist = True
	Else
		LookUpBOMHdrExist = False
	End If
End Function

'==========================================================================================
'   Function Name :InitFieldData
'   Function Desc :���� Case�� ���� Field���� �Ӽ��� �����Ѵ�.
'==========================================================================================
Sub InitFieldData()
	
	'frm1.txtItemCd1.value = ""
	frm1.txtItemNm1.value = ""
	frm1.cboItemAcct.value = ""
	frm1.txtSpec.value = ""
	frm1.txtchildItemQty.Text = ""
	frm1.txtChildItemUnit.value = ""
	frm1.txtPlantItemFromDt.Text = ""
	frm1.txtPlantItemToDt.Text = "" 
	
	'frm1.txtBomNo1.value = ""
	frm1.txtBOMDesc.value = ""
	frm1.txtDrawPath.value = ""
	frm1.txtValidFromdt.Value = ""
	frm1.txtValidToDt.Value = ""
	'frm1.rdoDefaultFlg1.checked = True
	'lgRdoOldVal = 1
	
	frm1.txtSafetyLt.Text = "" 
	frm1.txtItemSeq.Text = "" 
	frm1.txtChildItemUnit.value = ""
	frm1.txtChildItemQty.Text = ""
	frm1.txtLossRate.Text = "" 
	frm1.txtRemark.value = ""
	'frm1.txtValidFromdt1.value = ""
	'frm1.txtValidToDt1.value = ""
	If lgClkInsrtRow = True Then
	
'		frm1.txtPrntItemQty = "1.0000"
'		frm1.txtPrntItemUnit.value = frm1.txtBasicUnit.value 
'		frm1.txtValidFromDt1.text = UNIFormatDate("<%=EndDate%>")
'		frm1.txtValidToDt1.text = UNIFormatDate("2999-12-31")
	Else
	    frm1.txtPrntItemQty = ""
		frm1.txtPrntItemUnit.value = ""
		frm1.txtValidFromDt1.text = ""
		frm1.txtValidToDt1.text = ""
	End If
    frm1.rdoSupplyFlg1.checked = True
	lgRdoOldVal2 = 1  
	
End Sub

'==========================================================================================
'   Function Name :SetFieldProp
'   Function Desc :���� Case�� ���� Field���� �Ӽ��� �����Ѵ�.
'==========================================================================================
Function SetFieldProp(ByVal sVal)
	
	Dim iHdrVal
	Dim iDtlVal
	
	iHdrVal = CInt(Left(CStr(sVal),1))
	iDtlVal = CInt(Right(CStr(sVal),1))
	
	With frm1
		
		'------------------------------------------------------------------
		   ' Header: 1 - ��ü Protect ���� 
		   '         2 - Copy�� 
		   '         3 - ��ȸ ���� 
		   '		 4 - LookUpBomNoOk���� 
		   '         5 - ��ü �Է� ���� ���� 
		   '		 6 - ���� �Һ� ���� 
		   ' Detail: 1 - ��ü Protect
		   '		 2 - child item seq �� Protect
		   '		 3 - ��ü �Է� ���� ���� 
		   '         4 - ���� ���� �Һ�	
		   '------------------------------------------------------------------
		   
			If iHdrVal = 1 Or iHdrVal = 2 Then
				If iHdrVal = 1 Then
					Call ggoOper.SetReqAttr(.txtItemCd1, "Q")
				Else				
					Call ggoOper.SetReqAttr(.txtItemCd1, "N")
				End If
				Call ggoOper.SetReqAttr(.txtBomNo1, "Q")
				Call ggoOper.SetReqAttr(.txtBOMDesc, "Q")
				'Call ggoOper.SetReqAttr(.rdoDefaultFlg1,"Q")
				'Call ggoOper.SetReqAttr(.rdoDefaultFlg2,"Q")
'				Call ggoOper.SetReqAttr(.txtValidFromDt,"Q")
'				Call ggoOper.SetReqAttr(.txtValidToDt,"Q")
				Call ggoOper.SetReqAttr(.txtDrawPath, "Q")
			ElseIf iHdrVal = 3 Or iHdrVal = 4 Or iHdrVal = 5 Or iHdrVal=6 Then
				If iHdrVal = 3 then
					Call ggoOper.SetReqAttr(.txtItemCd1, "Q")
					Call ggoOper.SetReqAttr(.txtBomNo1, "Q")
'					Call ggoOper.SetReqAttr(.txtValidFromDt,"N")
				ElseIf iHdrVal = 4 then
					'Call ggoOper.SetReqAttr(.txtItemCd1,"N")
					'Call ggoOper.SetReqAttr(.txtBomNo1,"N")
'					Call ggoOper.SetReqAttr(.txtValidFromDt,"N")
				ElseIf iHdrVal = 5 Then
					Call ggoOper.SetReqAttr(.txtItemCd1, "N")
					Call ggoOper.SetReqAttr(.txtBomNo1, "N")
'					Call ggoOper.SetReqAttr(.txtValidFromDt,"N")
				Else
					Call ggoOper.SetReqAttr(.txtItemCd1, "N")
					Call ggoOper.SetReqAttr(.txtBomNo1, "Q")
'					Call ggoOper.SetReqAttr(.txtValidFromDt,"N")
				End IF				
				Call ggoOper.SetReqAttr(.txtBOMDesc, "D")
				'Call ggoOper.SetReqAttr(.rdoDefaultFlg1,"N")
				'Call ggoOper.SetReqAttr(.rdoDefaultFlg2,"N")
'				Call ggoOper.SetReqAttr(.txtValidToDt,"N")
				Call ggoOper.SetReqAttr(.txtDrawPath, "D")	
				Call ggoOper.SetReqAttr(.txtECNNo1, "Q")
				Call ggoOper.SetReqAttr(.txtECNDesc1, "Q")
				Call ggoOper.SetReqAttr(.txtReasonCd1, "Q")	   
			End If
			
			'Detail Setting
			
			If iDtlVal = 1 Then
				Call ggoOper.SetReqAttr(.txtItemSeq, "Q")
				Call ggoOper.SetReqAttr(.txtChildItemQty, "Q")
				Call ggoOper.SetReqAttr(.txtChildItemUnit, "Q")
				Call ggoOper.SetReqAttr(.txtPrntItemQty, "Q")
				Call ggoOper.SetReqAttr(.txtPrntItemUnit, "Q")
				Call ggoOper.SetReqAttr(.txtSafetyLt, "Q")
				Call ggoOper.SetReqAttr(.txtLossRate, "Q")
				Call ggoOper.SetReqAttr(.rdoSupplyFlg1, "Q")
				Call ggoOper.SetReqAttr(.rdoSupplyFlg2, "Q")
				Call ggoOper.SetReqAttr(.txtRemark, "Q")
				Call ggoOper.SetReqAttr(.txtValidFromDt1, "Q")
				Call ggoOper.SetReqAttr(.txtValidToDt1, "Q")
			ElseIf iDtlVal = 2 Or iDtlVal = 3 Then
				If iDtlVal = 2 Then
					Call ggoOper.SetReqAttr(.txtItemSeq, "Q")
					Call ggoOper.SetReqAttr(.txtValidFromDt1, "Q")
				Else
				    Call ggoOper.SetReqAttr(.txtItemSeq, "N")
				    Call ggoOper.SetReqAttr(.txtValidFromDt1, "N")
				End If
				Call ggoOper.SetReqAttr(.txtChildItemQty, "N")
				Call ggoOper.SetReqAttr(.txtChildItemUnit, "N")
				Call ggoOper.SetReqAttr(.txtPrntItemQty, "N")
				Call ggoOper.SetReqAttr(.txtPrntItemUnit, "N")
				Call ggoOper.SetReqAttr(.txtSafetyLt, "D")
				Call ggoOper.SetReqAttr(.txtLossRate, "D")
				Call ggoOper.SetReqAttr(.rdoSupplyFlg1, "Q")
				Call ggoOper.SetReqAttr(.rdoSupplyFlg2, "Q")
				Call ggoOper.SetReqAttr(.txtRemark, "D")
				Call ggoOper.SetReqAttr(.txtValidToDt1, "N")
				If lgStrBOMHisFlg = "Y" Then
					Call ggoOper.SetReqAttr(.txtECNNo1, "N")
				Else
					Call ggoOper.SetReqAttr(.txtECNNo1, "Q")
				End If
				Call ggoOper.SetReqAttr(.txtECNDesc1, "Q")
				Call ggoOper.SetReqAttr(.txtReasonCd1, "Q")
				
			End If			
			
	End With
	
End Function


'==========================================================================================
'   Function Name :SetModChange()
'   Function Desc :Header�� Detail Sheet�� ���� ���¸� Change(Create,Update)
'==========================================================================================
Function SetModChange(ByVal iVal)
	
	If iVal = 0 Then							'Form_Load �ĳ� FncNew �� �ֻ��� ��ǰ���� BOM Header�Է� ���� 
		frm1.txtHdrMode.value = "C" 
		frm1.txtDtlMode.value = ""			
	ElseIf iVal = 1 Then						'����ǰ ��ǰ�� �Է½� BOM�� ���� ��� Header�� Detail Create
		frm1.txtHdrMode.value = "C" 
		frm1.txtDtlMode.value = "C"			
	ElseIf iVal = 2 Then						'����ǰ ��ǰ�� �Է½� BOM�� �ִ� ��� Header Update, Detail Create
		frm1.txtHdrMode.value = "U" 
		frm1.txtDtlMode.value = "C"			
	ElseIf iVal = 3 Then						'������ ��ǰ�� �Է½� Detail�� Create
		frm1.txtHdrMode.value = "" 
		frm1.txtDtlMode.value = "C"			
	ElseIf iVal = 4 Then						'����ǰ ��ǰ�� �Է½� Header�� Detail Update
		frm1.txtHdrMode.value = "U" 
		frm1.txtDtlMode.value = "U"			
	ElseIf iVal = 5 Then						'������ ��ǰ�� ������ BOM Detail�� Update
		frm1.txtHdrMode.value = "" 
		frm1.txtDtlMode.value = "U"			
	ElseIf iVal = 6 Then						'DBQueryOk�� LookUpHdrOk �� - �ֻ��� ǰ�� ��ȸ �� 
		frm1.txtHdrMode.value = "U" 
		frm1.txtDtlMode.value = ""			
	ElseIf iVal = 7 Then						'�ֻ��� ǰ�� BOM Copy
		frm1.txtHdrMode.value = "M"
		frm1.txtDtlMode.value = ""
	ElseIf iVal = 8 Then						'����ǰ BOM ���� 
		frm1.txtHdrMode.value = "M"
		frm1.txtDtlMode.value = "U"
	End If

End Function

'==========================================================================================
'   Function Name :LookUpHdr()
'   Function Desc :��ȸ������ ǰ�� ���� BOM header���� ��ȸ 
'==========================================================================================
Sub LookUpHdr(ByVal txtItemCd, ByVal txtBomNo)
	Dim strVal
	
	Call ggoOper.ClearField(Document, "2")
	Call SetFieldProp(31)
	
	LayerShowHide(1)
			    
	'------------------------------
	' Server Logic Call
	'------------------------------
	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&QueryType=" & "H"									'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&txtItemCd=" & txtItemCd							'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&txtBomNo=" & txtBomNo							'��: LookUP ���� ����Ÿ    
	
	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
End Sub

Sub LookUpHdrOk()
	Call SetModChange(6)
	Call SetToolbar("11111000000111")
	lgIntFlgMode = parent.OPMD_UMODE
	lgBlnFlgChgValue = False
	lgQueryMode = False

	frm1.hBomType.value = UCase(Trim(frm1.txtBomNo1.value))
	Call ggoOper.SetReqAttr(frm1.txtECNNo1, "Q")
	Call ggoOper.SetReqAttr(frm1.txtECNDesc1, "Q")
	Call ggoOper.SetReqAttr(frm1.txtReasonCd1, "Q")
End Sub


'==========================================================================================
'   Function Name :LookUpDtl()
'   Function Desc :��ǰ����� Header������ Detail ������ȸ 
'==========================================================================================
Sub LookUpDtl(ByVal txtChildItemCd,ByVal txtPrntBomNo,ByVal txtPrntItemCd,ByVal intChildItemSeq,ByVal intLevel,ByVal txtChildBomNo)
	Dim strVal
	
    Call ggoOper.ClearField(Document, "2")                                  '��: Clear Contents  Field

	LayerShowHide(1)
	
	If LookUpBOMHdrExist(Trim(frm1.txtPlantCd.value), Trim(txtChildItemCd), Trim(txtChildBomNo)) = True Then
		lgStrHeaderFlg = 1
	Else
		lgStrHeaderFlg = 0
	End If

	'------------------------------
	' Server Logic Call
	'------------------------------
	strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001				'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&QueryType=" & "D"									'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&txtPrntItemCd=" & txtPrntItemCd					'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&txtPrntBomNo=" & txtPrntBomNo					'��: LookUP ���� ����Ÿ    
	strVal = strVal & "&intChildItemSeq=" & intChildItemSeq				'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&txtChildBomNo=" & txtChildBomNo					'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&txtChildItemCd=" & txtChildItemCd
	strVal = strVal & "&txtBOMHeader=" & lgStrHeaderFlg
		
	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

End Sub

Sub LookUpDtlOk()
	If lgStrHeaderFlg = 1 Then					'frm1.hBomType.value = "E" Or 
		Call SetFieldProp(32)
		Call SetModChange(4)
		'Call SetToolbar("1111100000001")											'��: Insert Row, Delete Row ��ư�� Disable
	Else
		Call SetFieldProp(12)
		Call SetModChange(5)
		'Call SetToolbar("1111101000001")											'��: Insert Row, Delete Row ��ư�� Disable
	End If
	
	Call SetToolbar("11101000000111")
	
	If frm1.hPrntProcType.value = "O" Then
		Call ggoOper.SetReqAttr(frm1.rdoSupplyFlg1, "N")
		Call ggoOper.SetReqAttr(frm1.rdoSupplyFlg2, "N") 
	Else
		Call ggoOper.SetReqAttr(frm1.rdoSupplyFlg1, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoSupplyFlg2, "Q") 
	End If
	
	lgIntFlgMode = parent.OPMD_UMODE
	lgQueryMode = False
	lgBlnFlgChgValue = False

    If lgStrBOMHisFlg = "Y" Then
		Call ggoOper.SetReqAttr(frm1.txtECNNo1, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtECNNo1, "Q")
	End If

	Call ggoOper.SetReqAttr(frm1.txtReasonCd1, "Q")
	Call ggoOper.SetReqAttr(frm1.txtECNDesc1, "Q")
	
End Sub


'==========================================================================================
'   Function Name :SetChildNode(���)
'   Function Desc :
'==========================================================================================
Function SetChildNode()
	Dim NodX
	Dim Node
	Dim PrntNode
	
	SetChildNode = False
		
	Set NodX = frm1.uniTree1.SelectedItem 
	Set PrntNode = NodX.Parent
	
	If NodX is Nothing Then										'���õ� Item�� ���� ��� 
		Exit Function
	End If
	
	If frm1.rdoSrchType1.checked = True Then					'�ܴܰ��̸�		
		If Not(PrntNode is Nothing) Then						'Root�� �ƴϸ� ��� 
			Call DisplayMsgBox("182722", "X", "X", "X")
			Exit Function
		End If
	Else
		If Not (Trim(frm1.txtItemAcctGrp.value) = "1FINAL" Or Trim(frm1.txtItemAcctGrp.value) = "2SEMI" Or frm1.hBomType.value = "E") Then
			Call DisplayMsgBox("182618", "X", "X", "X")
			Exit Function
		End If
		
		If Not(PrntNode is Nothing) Then						'Root�� �ƴϸ� ��� 
			frm1.txtPrntBomNo.value = frm1.txtBomNo1.value  
			frm1.txtPrntItemCd.value = NodX.Text				'�Է��ϰ��� �ϴ� ��ǰ���� ��ǰ���� �ӽ÷� �����Ѵ�.
        End If
	End If
		
	Set Node = frm1.uniTree1.Nodes.Add(NodX.Key, tvwChild, C_NEW_FOLDER_KEY, C_NEW_FOLDER, C_MATL, C_MATL)
	
	NodX.Expanded = True
	
	Set NodX = Nothing
	Set Node = Nothing
	Set PrntNode = Nothing

	SetChildNode = True

End Function

'==========================================================================================
'   Function Name :DdDeleteNode(���)
'   Function Desc :Node�� ������ �� (history ������ ECN�� Child Item Code �ѱ�)
'==========================================================================================
Function DbDeleteNode()

	Dim strVal
	
	LayerShowHide(1)
				    
	'------------------------------
	' Server Logic Call
	'------------------------------
	frm1.txtHdrMode.value = ""
	frm1.txtDtlMode.value = "D"

	Call ExecMyBizASP(frm1, BIZ_PGM_DTLSAVE_ID)

End Function

'==========================================================================================
'   Function Name :InitTreeImage
'   Function Desc :TreeView Image�� �ʱ�ȭ�Ѵ�.
'==========================================================================================
Function InitTreeImage()
	
	Dim NodX, lHwnd
	
	With frm1
	
	.uniTree1.SetAddImageCount = 4
	.uniTree1.Indentation = "200"	' �� ���� 

	.uniTree1.AddImage C_IMG_PROD, C_PROD, 0												'��: TreeView�� ���� �̹��� ���� 
	.uniTree1.AddImage C_IMG_MATL, C_MATL, 0
	.uniTree1.AddImage C_IMG_ASSEMBLY, C_ASSEMBLY, 0												'��: TreeView�� ���� �̹��� ���� 
	.uniTree1.AddImage C_IMG_PHANTOM, C_PHANTOM, 0
	.uniTree1.AddImage C_IMG_SUBCON, C_SUBCON, 0

	.uniTree1.OLEDragMode = 0														'��: Drag & Drop �� �����ϰ� �� ���ΰ� ���� 
	.uniTree1.OLEDropMode = 0
	
	End With

End Function

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
 Sub txtValidFromDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidFromDt1.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidFromDt1.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
 Sub txtValidToDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidToDt1.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidToDt1.Focus
    End If
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

Sub txtChildItemQty_Change()
lgBlnFlgChgValue = True	
End Sub

Sub txtPrntItemQty_Change()
lgBlnFlgChgValue = True
End Sub

Sub txtLossRate_Change()
lgBlnFlgChgValue = True
End Sub

Sub txtSafetyLt_change()
lgBlnFlgChgValue = True
End Sub

Sub cboBomFlg_onChange()
lgBlnFlgChgValue = True
End Sub

Sub txtValidToDt1_change()
lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : txtChildItemCd_OnChange()
'   Event Desc : ��ǰ���ڵ� ����� LookUp ���� 
'==========================================================================================
Sub txtItemCd1_OnKeyPress()
	
	If UCase(frm1.txtItemCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Sub
	'Call InitFieldData()
	lgBlnFlgChgValue = True
End Sub

Sub txtBomNo1_OnKeyPress()
	
	If UCase(frm1.txtBomNo1.className) = UCase(parent.UCN_PROTECTED) Then Exit Sub
	
	'Call SetDefaultVal
	lgBlnFlgChgValue = True
End Sub

Sub txtItemCd1_OnChange()

	If Trim(frm1.txtItemCd1.value) = "" Then Exit Sub
	
	Call LookUpItemByPlant()
	
End Sub


'==========================================================================================
'   Event Name : Bom No OnChange�� LookUp�� �� 
'   Event Desc : 
'==========================================================================================

Sub txtBomNo1_OnChange()
	If frm1.txtBomNo1.value <> "" Then
		Call LookUpBomNoForChild()
	'Else
	'	Call SetModChange(3)			'BOM No�� �����Ǹ� Header������ �Է����� �ʴ� �ɷ� ���� 
	End If
	IsOpenPop = True					'Look Up�� Popup�� ����Ǵ� �� ���� 
End Sub

Sub rdoSupplyFlg1_OnClick()
	If lgRdoOldVal2 = 1 Then Exit Sub
	
	lgRdoOldVal2 = 1
	lgBlnFlgChgValue = True
End Sub

Sub rdoSupplyFlg2_OnClick()
	If lgRdoOldVal2 = 2 Then Exit Sub
	
	lgRdoOldVal2 = 2
	lgBlnFlgChgValue = True
End Sub

Sub txtPlantCd_OnChange()
	If Trim(frm1.txtPlantCd.value) <> "" Then
		Call CommonQueryRs("BOM_HISTORY_FLG", "P_PLANT_CONFIGURATION", "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
		If lgF0 = "" Or Left(lgF0, 1) = "N" Then
			lgStrBOMHisFlg = "N"
		Else
			lgStrBOMHisFlg = "Y"
		End If
	End If
End Sub

Sub txtECNNo1_OnChange()
	Dim iStrColSQL, iStrEcnDesc, iStrEcnNo, iStrReasonCd, iStrReasonNm
	Dim iArrECN(10)
	
	iStrColSQL = "ECN_NO, ECN_DESC, REASON_CD, dbo.ufn_GetCodeName(" & FilterVar("P1402", "''", "S") & ", REASON_CD)"
	Call CommonQueryRs(iStrColSQL, "P_ECN_MASTER", "ECN_NO = " & FilterVar(frm1.txtECNNo1.value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	If Trim(lgF0) <> "" Then
		iArrECN(0) = Split(lgF0, Chr(11))(0)
		'iArrECN(10) = Split(lgF3, Chr(11))(0)
		iArrECN(1) = Split(lgF1, Chr(11))(0)
		iArrECN(2) = Split(lgF2, Chr(11))(0)
		iArrECN(3) = Split(lgF3, Chr(11))(0)
		
		Call SetEcnNo(iArrECN)
	Else		
		frm1.txtECNDesc1.value = ""
		frm1.txtReasonCd1.value = ""
		frm1.txtReasonNm1.value = ""
		Call ggoOper.SetReqAttr(frm1.txtECNDesc1, "N")
		Call ggoOper.SetReqAttr(frm1.txtReasonCd1, "N")
		
		frm1.txtECNDesc1.focus 
		Set gActiveElement = document.activeElement 
	End If
End Sub

'==========================================================================================
'   Event Name : uniTree1_NodeClick
'   Event Desc : Node Click�� Look Up Call
'==========================================================================================
Sub uniTree1_NodeClick(ByVal Node)
    Dim tmpSelNode
    Dim NodX
    Dim tmpNode
    Dim prntNode
        
    Dim intRetCD
	Dim iPos1
	Dim iPos2
	Dim iPos3
	
	Dim txtPrntItemCd
	Dim txtPrntBomNo
	Dim txtChildItemCd
	Dim txtChildBomNo
	Dim intChildItemSeq
	Dim intLevel
		
	Err.Clear																		'��: Protect system from crashing
	
	lgNodeClick = True			'Node ���� Click�ߴ� �� ���� 
		
	With frm1
	
		If lgQueryMode = True Then Exit Sub
		
		Set NodX = .uniTree1.SelectedItem
    
		If Not NodX Is Nothing Then													' ���õ� ������ ������ 
			
			'-------------------------------------
			'If Same Node Clicked, Exit
			'---------------------------------------
			tmpSelNode = lgSelNode
			
			If NodX.Key = lgSelNode Then
				Set NodX = Nothing
				Exit Sub
			Else
				lgSelNode = NodX.Key
			End If
			
			'-------------------------------------
			'If Data Changed, Msg Display
			'---------------------------------------
			
			If lgBlnFlgChgValue = True Then
				IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'��: "Will you destory previous data"
				If IntRetCD = vbNo Then
					lgSelNode = tmpSelNode
					lgNodeClick = False
					Exit Sub
				End If
			End If
		
			'-------------------------------------
			'Insert Row�ϰ� Save���� ���� ���¿��� �ٸ� Node�� Ŭ���� ��� 
			'---------------------------------------

			If lgClkInsrtRow = True Then													'InsertRow�� ���¿��� �ٸ� Node�� Ŭ���� ��� �����Ѵ�.	
				Set tmpNode = .uniTree1.Nodes(C_NEW_FOLDER_KEY)									'Insert Row�� �������� üũ�ϰ� insert row�� Node�� �����Ѵ�.
				If Not tmpNode is nothing Then
					.uniTree1.Nodes.Remove(tmpNode.Index) 
				End If
				lgClkInsrtRow = False 
				
				Set tmpNode = Nothing
			End If	
			
			'-------------------------------------
			'Get Parent Information And LookUp Detail
			'---------------------------------------
			
			Set PrntNode = NodX.Parent
			
			If PrntNode is Nothing Then
				If frm1.txtHdrMode.value  = "C" And frm1.txtDtlMode.value = "" Then
					Exit Sub
				End If	
			    iPos1 = InStr(1, NodX.Key, "|^|^|")										'Parent Bom No
			    txtPrntItemCd  = Trim(Mid(NodX.Key, 1, iPos1 - 1))
				txtPrntBomNo  = Trim(Right(NodX.Key, Len(NodX.Key) - (iPos1 + 4)))
				
				Call LookUpHdr(txtPrntItemCd, txtPrntBomNo) 

			Else
		        txtPrntItemCd = Trim(PrntNode.Text)

		        iPos1 = InStr(1, PrntNode.Key, "|^|^|")									'Parent Info
				iPos2 = InStr(iPos1 + 5, PrntNode.Key, "|^|^|")
				
				If iPos2 <> 0 Then
					txtPrntBomNo = Trim(Mid(PrntNode.Key, iPos1 + 5, iPos2 - (iPos1 + 5)))
				Else 
					txtPrntBomNo = Trim(Right(PrntNode.Key, Len(PrntNode.Key) - (iPos1 + 4)))
				End If
		    		    
				txtChildItemCd  = Trim(NodX.Text)

				iPos1 = InStr(1, NodX.Key, "|^|^|")
				intChildItemSeq = Mid(NodX.Key, 1, iPos1 - 1)
				
				iPos2 = InStr(iPos1 + 5, NodX.Key, "|^|^|")								'Child Item Seq
				txtChildBomNo = Mid(NodX.Key, iPos1 + 5, iPos2 - (iPos1 + 5))
			   
				iPos3 = InStr(iPos2 + 5, NodX.Key, "|^|^|")								'Level
				If iPos3 <> 0 Then
					intLevel = Mid(NodX.Key, iPos2  +5, iPos3 - (iPos2 + 5))
				Else
					intLevel = Mid(NodX.Key, iPos2 + 5, Len(NodX.Key) - (iPos2 +5))
				End If

				Call LookUpDtl(txtChildItemCd, txtPrntBomNo, txtPrntItemCd, intChildItemSeq, intLevel, txtChildBomNo)
			
			End If

			.txtPrntItemCd.value = txtPrntItemCd
			.txtPrntBomNo.value = txtPrntBomNo
			
		End If
    End With

	'-----------------------------
	' ���� ��ȸ�� ���������� ���� 
	'-----------------------------
	lgQueryMode = True	

	'-----------------------------
	' Object Nothing
	'-----------------------------
    Set NodX = Nothing
    Set PrntNode = Nothing
End Sub


'==========================================================================================
'   Event Name : uniTree1_MouseUp
'   Event Desc : Node�� Drag�Ҷ� �̺�Ʈ 
'==========================================================================================

Sub uniTree1_MouseUp(Node, Button , Shift, X, Y )
	Dim NodX
	Dim PrntNode
	Dim NodFlg
	
	With frm1
	
	'--------------------------------------------
	' ���� ��ȸ������ ��� ���� Ŭ���Ǿ��� �� üũ 
	'--------------------------------------------
	
	If lgQueryMode = True or lgNodeClick = False Then Exit Sub
		
	'--------------------------------------------
	' ������ ���콺�� Ŭ���Ǿ��� �� üũ 
	'--------------------------------------------
	
	If Button = 2 Or Button = 3 Then
		'--------------------------------------------
		' �ֻ��� ǰ������ üũ�Ͽ� �޴����� �ٲ۴�.
		'--------------------------------------------
		.uniTree1.OpenTitle = "��ǰ���߰�"
		
		.uniTree1.RenameTitle = ""
		
		'--------------------------------------------
		' BOM ���� ����� �������� ���� 
		' BOM No�� �������� ���Ͽ� üũ�� ����� 
		'--------------------------------------------
		.uniTree1.DeleteTitle = ""			
		
		Set NodX = .uniTree1.SelectedItem
		Set PrntNode = NodX.Parent 
		
		If NodX.Key = C_NEW_FOLDER_KEY Then Exit Sub		
		                           
		If PrntNode is Nothing Then
			.uniTree1.AddTitle = "BOM ����"
			lgStrHeaderFlg = "1"
			'.uniTree1.DeleteTitle = "BOM ����"
			NodFlg = 1
		Else
			.uniTree1.AddTitle = "��ǰ�����"
			'.uniTree1.DeleteTitle = "BOM ����"			
			NodFlg = 2
		End If
		
		Set NodX = Nothing
		Set PrntNode = Nothing
		
		'--------------------------------------------
		' �޴��� Display�Ѵ�.
		'--------------------------------------------
		
		If .txtHdrMode.value = "U" Then
			If lgStrHeaderFlg = 1 Then
				.uniTree1.MenuEnabled C_MNU_OPEN, TRUE
				.uniTree1.MenuEnabled C_MNU_ADD, TRUE

				'------------------------------------------------
				' ** NodFlg�� üũ 
				'    ����� �ֻ��� ǰ���� �ƴ� �߰��ܰ��� ǰ�� 
				'    ���ؼ��� ���縦 ���ϵ��� ���� ������.
				'    ���Ŀ� �߰��ܰ� ǰ���� ���縦 �ϰ��� �ϸ� 
				'    �� �κи� �����ϸ� �� 
				'------------------------------------------------
				
				'If NodFlg = 1 Then
				'	.uniTree1.MenuEnabled C_MNU_DELETE, TRUE
				'Else
				'	.uniTree1.MenuEnabled C_MNU_DELETE, FALSE
				'End If
				'--------------------------------------------------
				
				'.uniTree1.MenuEnabled C_MNU_RENAME, FALSE
			End If
		Else
			.uniTree1.MenuEnabled C_MNU_OPEN, FALSE
			.uniTree1.MenuEnabled C_MNU_ADD, TRUE
			'.uniTree1.MenuEnabled C_MNU_DELETE, FALSE
			'.uniTree1.MenuEnabled C_MNU_RENAME, FALSE
		End If
	
		.uniTree1.PopupMenu
	End If 
	End With	
	
	lgNodeClick = False
	
End Sub


'==========================================================================================
'   Event Name : uniTree1_MenuOpen
'   Event Desc : Node�� Drag�Ҷ� �̺�Ʈ 
'==========================================================================================


Sub uniTree1_MenuOpen(Node)
	Call FncInsertRow()
End Sub


'==========================================================================================
'   Event Name : uniTree1_MenuAdd
'   Event Desc : Node�� Drag�Ҷ� �̺�Ʈ 
'==========================================================================================


Sub uniTree1_MenuAdd(Node)
	Call FncDeleteRow()
End Sub


'==========================================================================================
'   Event Name : uniTree1_MenuDelete
'   Event Desc : BOM Copy/BOM ���� 
'==========================================================================================


Sub uniTree1_MenuDelete(Node)
	'Call FncCopy()
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
      
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

 '-----------------------
    'Check previous data area
    '----------------------- 

    If lgBlnFlgChgValue = True Then
	IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    End If
    
 '-----------------------
    'Erase contents area
    '----------------------- 

	frm1.uniTree1.Nodes.Clear		
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	   											'��: Tree View Content
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")
    'Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables
  
 '-----------------------
    'Check condition area
    '----------------------- 

    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
 '-----------------------
    'Query function call area
    '----------------------- 

    If DbQuery = False Then   
		Exit Function           
    End If     																	'��: Query db data
       
    FncQuery = True																'��: Processing is OK
        
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    Dim sPlantCd
    Dim sPlantNm
    
    FncNew = False                                                          '��: Processing is NG
    
 '-----------------------
    'Check previous data area
    '-----------------------

    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
 '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------

    sPlantCd = frm1.txtPlantCd.value 
    sPlantNm = frm1.txtPlantNm.value
    
    frm1.uniTree1.Nodes.Clear  
	
	Call ggoOper.ClearField(Document, "A")                                      '��: Clear Contents  Field
    
    frm1.txtPlantCd.value = sPlantCd
    frm1.txtPlantNm.value = sPlantNm
    
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables
    Call SetToolbar("11101000000011")
    Call SetModChange(0)
    Call SetFieldProp(51)    
    
    frm1.txtItemCd1.focus
    Set gActiveElement = document.activeElement  
    
    FncNew = True																'��: Processing is OK

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    Dim intRetCd
    
    FncDelete = False														'��: Processing is NG
    
 '-----------------------
    'Precheck area
    '-----------------------

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If
    
 '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		            '��: "Will you destory previous data"	
	If IntRetCD = vbNo Then	
		Exit Function	
	End If
	
	If DbDelete = False Then   
		Exit Function           
    End If     							
       
    FncDelete = True                                                        '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 

    Dim IntRetCD 

    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

	'-----------------------
    'Precheck area
    '-----------------------

    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '��: No data changed!!
        Exit Function
    End If
    
	'-----------------------
    'Check content area
    '-----------------------

    If Not chkField(Document, "2") Then                             '��: Check contents area
       Exit Function
    End If
    
	'-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then
		Exit Function           
    End If     							                                          '��: Save db data

    FncSave = True                                                          '��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
	Dim IntRetCD
    Dim NodX
    Dim PrntNode
    
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	' �ֻ��� ǰ������ �ƴϸ� ����ǰ���� ���� 
	
	Set NodX = frm1.uniTree1.SelectedItem
	
	If NodX is Nothing Then Exit Function
	
	Set PrntNode = NodX.Parent
	
	If PrntNode is Nothing Then						' �ֻ��� ǰ���̸� 
		Call SetFieldProp(51)
		Call SetModChange(7)
	Else											' ����ǰ�̸� 
		Call SetFieldProp(52)
		Call SetModChange(8)
	End If 
    
    '-----------------------------------------------
	'BOM Copy�� ����Ǳ� �� ǰ��� BOM
	'-----------------------------------------------
	frm1.txtBaseItemCd.value = frm1.txtItemCd1.value
	frm1.txtBaseBomNo.value = frm1.txtBomNo1.value
    	
    'Default Value Setting
    Call InitFieldData
	frm1.txtItemCd1.value = ""

	Call SetDefaultVal
	
	frm1.txtItemCd1.focus 
	Set gActiveElement = document.activeElement 

	lgClkCopy = True
	 
	Set NodX = Nothing
	Set PrntNode = Nothing
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
    Dim IntRetCD 
    Dim BlnRetCd 
        
    On Error Resume Next                                                   '��: Protect system from crashing
    
    FncInsertRow = False                                                          '��: Processing is NG
    
	'-----------------------
    'Check previous data area
    '-----------------------

    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '--------------------------------------------------
    'Node Index�� �� Valid Check
    '--------------------------------------------------
	BlnRetCd = SetChildNode()
	
	If BlnRetCd = False Then
		Exit Function
	End If	

 '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
	Call ggoOper.ClearField(Document, "2")                                      '��: Clear Contents  Field
    
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables
    Call SetFieldProp(23)
    
    If Trim(frm1.txtProcType.value) = "O" Then
		Call ggoOper.SetReqAttr(frm1.rdoSupplyFlg1, "N")
		Call ggoOper.SetReqAttr(frm1.rdoSupplyFlg2, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.rdoSupplyFlg1, "Q")
		Call ggoOper.SetReqAttr(frm1.rdoSupplyFlg2, "Q")
		frm1.rdoSupplyFlg1.checked = True
	End If		
	
	
	
    frm1.txtPrntItemQty = "1" & parent.gComNumDec & "0000"
	frm1.txtPrntItemUnit.value = frm1.txtBasicUnit.value 
	frm1.txtValidFromDt1.text = Startdate
	frm1.txtValidToDt1.text = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
	
    lgSelNode = C_NEW_FOLDER_KEY
    
    frm1.txtItemCd1.focus 
    Set gActiveElement = document.activeElement 
    
    lgClkInsrtRow = True														'FncInsertRow�� ����Ǿ����� ���� 
     
    FncInsertRow = True																'��: Processing is OK
     
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
	Dim intRetCd
	Dim NodX
	Dim PrntNode
	
	On Error Resume Next								'��: Protect system from crashing

    If Not chkField(Document, "2") Then                             '��: Check contents area
       Exit Function
    End If
	
	DeleteNode = False
		
	Set NodX = frm1.uniTree1.SelectedItem 
	Set PrntNode = NodX.Parent
	
	If NodX is Nothing Then								'���õ� Item�� ���� ��� 
		Exit Function
	End If
	
	If PrntNode is Nothing Then							'Root�� ��� 
		Call FncDelete()
	Else												'Child Node�� ��� 
		intRetCd = DisplayMsgBox("182721", parent.VB_YES_NO, "X", "X")
		If intRetCd = vbNo Then
			Exit Function
		End If
		
		Call DbDeleteNode()
	End If
	
	Set NodX = Nothing
	Set PrntNode = Nothing

	DeleteNode = True  	
	
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                               '��: 
        Exit Function
    ElseIf lgPrevNo = "" Then
		Call DisplayMsgBox("900011", "X", "X", "X")
		Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: 
    strVal = strVal & "&txtPlantCd=" & lgPrevNo							'��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                               '��: 
        Exit Function
    ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900012", "X", "X", "X")
		Exit Function
    End If
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���°� 
    strVal = strVal & "&txtPlantCd=" & lgNextNo							'��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
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
    Err.Clear                                                           '��: Protect system from crashing
    Dim strVal
    
    DbDelete = False													'��: Processing is NG
    
	LayerShowHide(1)
		   
	'------------------------------
	' Server Logic Call
	'------------------------------
    strVal = BIZ_PGM_HDRDEL_ID & "?txtMode=" & parent.UID_M0001				'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)		'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd1.value)				'��: LookUP ���� ����Ÿ 
	strVal = strVal & "&txtBomNo=" & Trim(frm1.txtBomNo1.value)			'��: LookUP ���� ����Ÿ    
		
	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
    
    DbDelete = True                                                     '��: Processing is NG

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================

Function DbDeleteOk()														'��: ���� ������ ���� ���� 
	Call FncNew()
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
    Dim PrntKey
    Dim strVal
    Dim Node
    
    Err.Clear															'��: Protect system from crashing
    
    DbQuery = False														'��: Processing is NG
    
    Err.Clear															'��: Protect system from crashing
				   
    LayerShowHide(1)
		
	
    '----------------------------------------------
    '- Call Query ASP
    '----------------------------------------------
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&QueryType=" & "A"
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtBomNo=" & Trim(frm1.txtBomNo.value)
    strVal = strVal & "&txtBaseDt=" & Trim(frm1.txtBaseDt.text)
    
    If frm1.rdoSrchType1.checked = True Then
		strVal = strval & "&rdoSrchType=" & frm1.rdoSrchType1.value 
    ElseIf frm1.rdoSrchType2.checked = True Then
		strVal = strval & "&rdoSrchType=" & frm1.rdoSrchType2.value 
    End If       

    Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 

    DbQuery = True														'��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()										'��: ��ȸ ������ ������� 
    Dim NodX
    
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Set NodX = frm1.uniTree1
		NodX.SetFocus
		Set NodX = Nothing
		Set gActiveElement = document.activeElement
    End If
    Call InitVariables
    lgIntFlgMode = parent.OPMD_UMODE								'��: Indicates that current mode is Update mode

    Call SetToolbar("11111000000111")
    Call SetFieldProp(31)									'Header�� �Է°��� ���� 
    Call SetModChange(6)									'��ȸ �� ù ���¸� Set
    
	Call ggoOper.SetReqAttr(frm1.txtECNNo1, "Q")
	Call ggoOper.SetReqAttr(frm1.txtECNDesc1, "Q")
	Call ggoOper.SetReqAttr(frm1.txtReasonCd1, "Q")
	 
End Function

Function DbQueryNotOk()
    lgIntFlgMode = parent.OPMD_UMODE								'��: Indicates that current mode is Update mode
    Call SetFieldProp(51)									'Header�� �Է°��� ���� 
    Call SetModChange(0)									'��ȸ �� ù ���¸� Set
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================

Function DbSave() 
		
    Err.Clear																'��: Protect system from crashing

    DbSave = False															'��: Processing is NG

	'---------------------------------------------------------------------------
	' ����ȯ�濡 ����������� �����Ǵ� ���� ��쿡�� ��ȯ 
	'---------------------------------------------------------------------------
	If Trim(frm1.txtPlantCd.value) <> "" Then
		If CheckPlant(frm1.txtPlantCd.value) = False Then
			Call DisplayMsgBox("125000", "X", "����", "0")					
			frm1.txtPlantCd.focus
			frm1.txtPlantNm.value = ""
			Set gActiveElement = document.activeElement
			Call LayerShowHide(0)
			Exit Function
		End If
	End If
	
	'---------------------------------------------------------------------------
	' ���� Detail�� ��¥ �ʵ带 ����ϰ� Header�� ��¥�ʵ带 ������� ���� ��� 
	' ���� üũ �κ��� �����ϰ� �Ʒ� �ּ��� ���� �� �� 
	'---------------------------------------------------------------------------
	If frm1.txtValidFromDt1.Text <> "" And frm1.txtValidToDt1.Text <> "" Then
		If ValidDateCheck(frm1.txtValidFromDt1, frm1.txtValidToDt1) = False Then Exit Function      

	End If

	'---------------------------------------------------------------------------
	' �߰��ϴ� ��ǰ���� �����簡 �ƴϰ� BOM���簡 �ƴ� ��� üũ 
	'---------------------------------------------------------------------------
	If Trim(frm1.txtBomNo1.value) <> "" And lgClkCopy = False Then
	
		If Trim(frm1.hBomType.value) = "1" And Trim(frm1.txtBomNo1.value) <> "1" Then
			Call DisplayMsgBox("182621", "X", "X", "X")
			frm1.txtBomNo1.focus
			Set gActiveElement = document.activeElement 
			Exit Function
		End If
	End If
	
    LayerShowHide(1)

    Dim strVal
	
    With frm1
		.txtMode.value = parent.UID_M0002											'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value  = parent.gUsrID            
		
		If frm1.txtHdrMode.value = "C" Or frm1.txtHdrMode.value = "M" Then
			frm1.txtValidFromDt.Value  = Startdate
			frm1.txtValidToDt.Value = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31")
		End If
			 
		If UCase(Trim(frm1.txtHdrMode.value)) = "M" Then	
			Call ExecMyBizAsp(frm1,BIZ_PGM_COPY_ID)
		Else
			If frm1.txtHdrMode.value  <> "" And frm1.txtDtlMode.value = "" Then

				Call ExecMyBizASP(frm1, BIZ_PGM_HDRSAVE_ID)						'��: �����Ͻ� ASP �� ���� 
			ElseIf frm1.txtDtlMode.value <> "" Then

				If UNICDbl(frm1.txtChildItemQty.Text) = 0 Then
					Call DisplayMsgBox("970022", "X", "��ǰ����ؼ�", "0")
					frm1.txtChildItemQty.focus
					Set gActiveElement = document.activeElement
					Call LayerShowHide(0)
					Exit Function
				End If

				If UNICDbl(frm1.txtPrntItemQty.Text) = 0 Then
					Call DisplayMsgBox("970022", "X", "��ǰ����ؼ�", "0")	
					frm1.txtPrntItemQty.focus
					Set gActiveElement = document.activeElement 
					Call LayerShowHide(0)
					Exit Function
				End If
				Call ExecMyBizASP(frm1, BIZ_PGM_DTLSAVE_ID)	
			End If
		End If					
	End With

    DbSave = True
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()															'��: ���� ������ ���� ���� 

    Call InitVariables
    
    lgBlnFlgChgValue = False
    Call MainQuery()

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
