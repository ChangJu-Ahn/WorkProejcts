<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1216MA1
'*  4. Program Name         : �������� ���� 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG010,PD6G020
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/05/15
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit							'��: indicates that All variables must be declared in advance

Const BIZ_PGM_QRY_ID = "q1216mb1.asp"					'��: ��ȸ �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "q1216mb2.asp"					'��: ���� �����Ͻ� ���� ASP�� 

Const BIZ_PGM_JUMP1_ID = "q1211ma1"
Const BIZ_PGM_JUMP5_ID = "q1216ma1"

Dim C_CheckItem				'= 1							
Dim C_InspItemCd			'= 2							'��: Spread Sheet�� Column�� ��� 
Dim C_InspItemNm			'= 3
Dim C_InspCharNm			'= 4
Dim C_InspOrder				'= 5
Dim C_InspMthdCd			'= 6
Dim C_InspMthdNm			'= 7
Dim C_InspUnitIndctnNm		'= 8
Dim C_WeightNm				'= 9
Dim C_InspSpec				'= 10
Dim C_LSL					'= 11
Dim C_USL					'= 12
Dim C_MthdOfCLCalNm			'= 13
Dim C_CalculatedQty			'= 14
Dim C_LCL					'= 15
Dim C_UCL					'= 16
Dim C_MeasmtEquipmtCd		'= 17
Dim C_MeasmtEquipmtNm		'= 18
Dim C_MeasmtUnitCd			'= 19
Dim C_InspProcessDesc		'= 20
Dim C_Remark				'= 21

Dim C_ItemCd				'= 1							
Dim C_ItemPopup				'= 2
Dim C_ItemNm				'= 3
Dim C_RoutNo				'= 4							
Dim C_RoutNoPopup			'= 5
Dim C_RoutNoDesc			'= 6
Dim C_OprNo					'= 7							
Dim C_OprNoPopup			'= 8
Dim C_OprNoDesc				'= 9

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgSaveFlag
Dim lgCheckall

Dim lgSortKey1
Dim lgSortKey2

Dim lgChangeFlag
Dim lgChangeFlag2

Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False                    'Indicates that no value changed
	lgIntGrpCount = 0                           'initializes Group View Size
	
	'---- Coding part--------------------------------------------------------------------
	lgStrPrevKey = ""                           'initializes Previous Key
	lgLngCurRows = 0                            'initializes Deleted Rows Count
	lgSaveFlag = False
	lgSortKey    = 1                            '��: initializes sort direction
	lgCheckall = 0
	Call SetToolBar("11000000000011")							'��: ��ư ���� ���� 
	
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()

	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
	
	frm1.cboInspClassCd.value		= "R"
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If	
	
	If ReadCookie("txtItemCd") <> "" Then
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
	End If	
	
	If ReadCookie("txtItemNm") <> "" Then
		frm1.txtItemNm.Value = ReadCookie("txtItemNm")
	End If	
	
	If ReadCookie("txtInspClassCd") <> "" Then
		frm1.cboInspClassCd.Value = ReadCookie("txtInspClassCd")
	End If	
	
	If ReadCookie("txtInspClassCd") = "P" Then
		If ReadCookie("txtRoutNo") <> "" Then
			frm1.txtRoutNo.Value = ReadCookie("txtRoutNo")
		End If
		
		If ReadCookie("txtRoutNoDesc") <> "" Then
			frm1.txtRoutNoDesc.Value = ReadCookie("txtRoutNoDesc")
		End If
		
		If ReadCookie("txtOprNo") <> "" Then
			frm1.txtOprNo.Value = ReadCookie("txtOprNo")
		End If
		
		If ReadCookie("txtOprNoDesc") <> "" Then
			frm1.txtOprNoDesc.Value = ReadCookie("txtOprNoDesc")
		End If
	End If
		
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""
	WriteCookie "txtItemNm", ""
	WriteCookie "txtInspClassCd", ""
	WriteCookie "txtRoutNo", ""
	WriteCookie "txtRoutNoDesc", ""
	WriteCookie "txtOprNo", ""
	WriteCookie "txtOprNoDesc", ""
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	Call initSpreadPosVariables(pvSpdNo)  
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then							'��: ����� vspdData1�϶� 
	
		With frm1.vspdData
	
			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20040518", , Parent.gAllowDragDropSpread
			
			.ReDraw = false
			
			.MaxCols = C_Remark +1			'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			.MaxRows = 0
			
			Call GetSpreadColumnPos("A")
			
			ggoSpread.SSSetCheck C_CheckItem, "", 4,,,1
			ggoSpread.SSSetEdit C_InspItemCd, "�˻��׸��ڵ�", 14, 0, -1, 5, 2		
			ggoSpread.SSSetEdit C_InspItemNm, "�˻��׸�� ",20, 0, -1, 40
			ggoSpread.SSSetEdit C_InspCharNm, "ǥ�üӼ�", 10, 0, -1, 40 		
			ggoSpread.SSSetFloat C_InspOrder, "�˻����", 10, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "P"
			ggoSpread.SSSetEdit C_InspMthdCd, "�˻����ڵ�", 14, 0, -1, 4, 2
			ggoSpread.SSSetEdit C_InspMthdNm, "�˻��ĸ�", 20, 0, -1, 40
			ggoSpread.SSSetEdit C_InspUnitIndctnNm, "�˻���� ǰ��ǥ��", 10, 0, False
			ggoSpread.SSSetEdit C_WeightNm, "�߿䵵", 10, 0, False
			ggoSpread.SSSetEdit C_InspSpec , "�˻�԰�", 20, 2, -1, 40
			ggoSpread.SSSetFloat C_LSL, "���ѱ԰�", 16, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C_USL, "���ѱ԰�", 16, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetEdit C_MthdOfCLCalNm, "�����Ѱ������", 18, 0, False
			ggoSpread.SSSetFloat C_CalculatedQty, "�����Ѱ����", 16, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
			ggoSpread.SSSetFloat C_LCL, "��������", 16, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetFloat C_UCL, "��������", 16, "8", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
			ggoSpread.SSSetEdit C_MeasmtEquipmtCd, "�������ڵ�", 20, 0, -1, 10, 2
			ggoSpread.SSSetEdit C_MeasmtEquipmtNm , "�������", 20, 0, -1, 40
			ggoSpread.SSSetEdit C_MeasmtUnitCd, "��������", 14, 0, -1, 3
			ggoSpread.SSSetEdit C_InspProcessDesc , "�˻���", 60, 0, -1, 400
			ggoSpread.SSSetEdit C_Remark , "���", 40, 0, -1, 200
			
 			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols , True)
			
			.ReDraw = true
			
			Call SetSpreadLock 
	
		End With
	End If
	
	IF	pvSpdNo = "B" Or pvSpdNo = "*" Then					'��: ����� vspdData2�϶� 
	
		With frm1.vspdData2
	
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20040830", , Parent.gAllowDragDropSpread
			
			.ReDraw = false
			
			.MaxCols = C_OprNoDesc + 1			'��: �ִ� Columns�� �׻� 1�� ������Ŵ		
			.MaxRows = 0
			
			Call GetSpreadColumnPos("B")
			ggoSpread.SSSetEdit		C_ItemCd, "ǰ���ڵ�", 18, 0, -1, 18, 2
			ggoSpread.SSSetButton	C_ItemPopup
			ggoSpread.SSSetEdit		C_ItemNm, "ǰ���",35, 0, -1, 40
			ggoSpread.SSSetEdit		C_RoutNo, "�����", 15, 0, -1, 20, 2
			ggoSpread.SSSetButton	C_RoutNoPopup
			ggoSpread.SSSetEdit		C_RoutNoDesc, "����ø�",35, 0, -1, 40
			ggoSpread.SSSetEdit		C_OprNo, "����", 5, 0, -1, 3, 2
			ggoSpread.SSSetButton	C_OprNoPopup
			ggoSpread.SSSetEdit		C_OprNoDesc, "�����۾���",35, 0, -1, 40
			
			Call ggoSpread.MakePairsColumn(C_ItemCd, C_ItemPopup)
			Call ggoSpread.MakePairsColumn(C_RoutNo, C_RoutNoPopup)
			Call ggoSpread.MakePairsColumn(C_OprNo, C_OprNoPopup)
 			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			.ReDraw = true
			
			Call SetSpreadLock2 
	
		End With
	End If
	
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock -1, -1
		ggoSpread.SpreadUnLock C_CheckItem, -1, C_CheckItem
		.vspdData.ReDraw = True
	End With
End Sub

Sub SetSpreadLock2()
	With frm1
		Call ggoSpread.SpreadLock(frm1.vspdData2.MaxCols, -1, frm1.vspdData2.MaxCols)
	End With
End Sub


'==========================================  2.2.6 SetSpreadColor1()  =======================================
'	Name : SetSpreadColor1()
'	Description : Combo Display
'========================================================================================================= 
Sub SetSpreadColor1(ByVal pvStartRow, ByVal pvEndRow)
	With frm1
		.vspdData2.ReDraw = False
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SSSetProtected C_ItemCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemPopup, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemNm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RoutNo, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RoutNoPopup, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RoutNoDesc, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_OprNo, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_OprNoPopup, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_OprNoDesc, pvStartRow, pvEndRow
		.vspdData2.ReDraw = True

		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_CheckItem, -1, C_CheckItem	
		.vspdData.ReDraw = True
	End With
End Sub

'==========================================  2.2.6 SetSpreadColor2()  =======================================
'	Name : SetSpreadColor2()
'	Description : Combo Display
'========================================================================================================= 
Sub SetSpreadColor2(ByVal pvStartRow, ByVal pvEndRow)
	With frm1
		.vspdData2.ReDraw = False
		ggoSpread.SSSetRequired C_ItemCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemNm, pvStartRow, pvEndRow
		
		If lgIntFlgMode = Parent.OPMD_CMODE Then
			If .cboInspClass.value = "P" Then
				ggoSpread.SpreadUnLock C_RoutNo, -1, C_RoutNo, -1
				ggoSpread.SpreadUnLock C_OprNo, -1, C_OprNo, -1
				
				ggoSpread.SSSetRequired C_RoutNo, -1,  -1
				ggoSpread.SSSetRequired C_OprNo, -1,  -1
			Else	
				ggoSpread.SpreadLock C_RoutNo, -1, C_RoutNo, -1
				ggoSpread.SpreadLock C_OprNo, -1, C_OprNo, -1
		
				ggoSpread.SSSetProtected C_RoutNo, -1, -1
				ggoSpread.SSSetProtected C_OprNo, -1, -1
			End If
		Else
			If .hInspClassCd.value = "P" Then
				ggoSpread.SpreadUnLock C_RoutNo, -1, C_RoutNo, -1
				ggoSpread.SpreadUnLock C_OprNo, -1, C_OprNo, -1
					
				ggoSpread.SSSetRequired C_RoutNo, -1,  -1
				ggoSpread.SSSetRequired C_OprNo, -1,  -1
			Else
				ggoSpread.SpreadLock C_RoutNo, -1, C_RoutNo, -1
				ggoSpread.SpreadLock C_OprNo, -1, C_OprNo, -1
		
				ggoSpread.SSSetProtected C_RoutNo, -1, -1
				ggoSpread.SSSetProtected C_OprNo, -1, -1
			End If
		End If
		
		ggoSpread.SSSetProtected C_RoutNoDesc, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_OprNoDesc, pvStartRow, pvEndRow
		.vspdData2.ReDraw = True
	End With
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    Err.Clear
    Call CommonQueryRs(" Minor_Cd, Minor_Nm ","B_Minor", "Major_Cd=" & FilterVar("Q0001", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboInspClassCd ,lgF0  ,lgF1  ,Chr(11))
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

	If pvSpdNo = "A" or pvSpdNo = "*" Then							'��: ����� vspdData�϶� 
		'vspdData
		C_CheckItem = 1							
		C_InspItemCd = 2							'��: Spread Sheet�� Column�� ��� 
		C_InspItemNm = 3
		C_InspCharNm = 4
		C_InspOrder = 5
		C_InspMthdCd = 6
		C_InspMthdNm = 7
		C_InspUnitIndctnNm = 8
		C_WeightNm = 9
		C_InspSpec = 10
		C_LSL = 11
		C_USL = 12
		C_MthdOfCLCalNm = 13
		C_CalculatedQty = 14
		C_LCL = 15
		C_UCL = 16
		C_MeasmtEquipmtCd = 17
		C_MeasmtEquipmtNm = 18
		C_MeasmtUnitCd = 19
		C_InspProcessDesc = 20
		C_Remark = 21
	End If
	
	If pvSpdNo = "B" or pvSpdNo = "*" Then							'��: ����� vspdData2�϶� 
		'vspdData2
		C_ItemCd				= 1							
		C_ItemPopup				= 2
		C_ItemNm				= 3
		C_RoutNo			    = 4		
		C_RoutNoPopup			= 5
		C_RoutNoDesc			= 6
		C_OprNo					= 7
		C_OprNoPopup			= 8
		C_OprNoDesc				= 9
	End If

End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos

 	Select Case UCase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 			
			C_CheckItem = iCurColumnPos(1)							
			C_InspItemCd = iCurColumnPos(2)
			C_InspItemNm = iCurColumnPos(3)
			C_InspCharNm = iCurColumnPos(4)
			C_InspOrder = iCurColumnPos(5)
			C_InspMthdCd = iCurColumnPos(6)
			C_InspMthdNm = iCurColumnPos(7)
			C_InspUnitIndctnNm = iCurColumnPos(8)
			C_WeightNm = iCurColumnPos(9)
			C_InspSpec = iCurColumnPos(10)
			C_LSL = iCurColumnPos(11)
			C_USL = iCurColumnPos(12)
			C_MthdOfCLCalNm = iCurColumnPos(13)
			C_CalculatedQty = iCurColumnPos(14)
			C_LCL = iCurColumnPos(15)
			C_UCL = iCurColumnPos(16)
			C_MeasmtEquipmtCd = iCurColumnPos(17)
			C_MeasmtEquipmtNm = iCurColumnPos(18)
			C_MeasmtUnitCd = iCurColumnPos(19)
			C_InspProcessDesc = iCurColumnPos(20)
			C_Remark = iCurColumnPos(21)
			
		Case "B"
 			ggoSpread.Source = frm1.vspdData2 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 			
 			C_ItemCd = iCurColumnPos(1)							
			C_ItemPopup = iCurColumnPos(2)
			C_ItemNm = iCurColumnPos(3)
			C_RoutNo = iCurColumnPos(4)					
			C_RoutNoPopup = iCurColumnPos(5)
			C_RoutNoDesc = iCurColumnPos(6)
			C_OprNo = iCurColumnPos(7)				
			C_OprNoPopup = iCurColumnPos(8)
			C_OprNoDesc	= iCurColumnPos(9)
 	End Select
 
End Sub

'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	OpenItem = false
	
	Dim arrRet
	Dim arrParam1, arrParam2, arrParam3, arrParam4, arrParam5
	Dim arrField(6)
	Dim iCalledAspName, IntRetCD

	'�����ڵ尡 �ִ� �� üũ 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X") 		'���������� �ʿ��մϴ� 
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam1 = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam2 = Trim(frm1.txtPlantNm.Value)	' Plant Name
	arrParam3 = Trim(frm1.txtItemCd.Value)	' Item Code
	arrParam4 = ""	'Trim(frm1.txtItemNm.Value)	' Item Name
	arrParam5 = Trim(frm1.cboInspClassCd.Value)
	
	iCalledAspName = AskPRAspName("q1211pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		  
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
	End If	

	frm1.txtItemCd.Focus
	Set gActiveElement = document.activeElement
	OpenItem = true
End Function

'====================  OpenRoutNo  ======================================
' Function Name : OpenRoutNo
' Function Desc : OpenRoutNo Reference Popup
'==========================================================================
Function OpenRoutNo()

	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	If frm1.txtItemCd.value= "" Then
		Call DisplayMsgBox("971012","X", "ǰ��","X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If
		
	arrParam(0) = "����� �˾�"					' �˾� ��Ī 
	arrParam(1) = "P_ROUTING_HEADER"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtRoutNo.value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "P_ROUTING_HEADER.PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
				  "And P_ROUTING_HEADER.ITEM_CD = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S") 	
	arrParam(5) = "�����"			
	
    arrField(0) = "ROUT_NO"							' Field��(0)
    arrField(1) = "DESCRIPTION"						' Field��(1)
    arrField(2) = "BOM_NO"							' Field��(1)
    arrField(3) = "MAJOR_FLG"						' Field��(1)
   
    arrHeader(0) = "�����"						' Header��(0)
    arrHeader(1) = "����ø�"					' Header��(1)
    arrHeader(2) = "BOM Type"					' Header��(1)
    arrHeader(3) = "�ֶ����"					' Header��(1)        
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False
    
	If arrRet(0) <> "" Then
		frm1.txtRoutNo.Value		= arrRet(0)		
		frm1.txtRoutNoDesc.Value		= arrRet(1)		
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtRoutNo.focus
	
End Function



'**************************** Function OpenOprNo() ***********************************8
Function OpenOprNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function    

	IsOpenPop = True
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	If frm1.txtItemCd.value= "" Then
		Call DisplayMsgBox("971012","X", "ǰ��","X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If	
	
	If frm1.txtRoutNo.value= "" Then
		Call DisplayMsgBox("971012","X", "�����","X")
		frm1.txtRoutNo.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If	

	arrParam(0) = "�����˾�"	
	arrParam(1) = "P_ROUTING_DETAIL A inner join P_WORK_CENTER B on A.wc_cd = B.wc_cd and A.plant_cd = B.plant_cd " & _
				  " left outer join B_MINOR C on A.job_cd = C.minor_cd and C.major_cd = " & FilterVar("P1006", "''", "S") & ""				
	arrParam(2) = UCase(Trim(frm1.txtOprNo.Value))
	arrParam(3) = ""
	arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
				  " and	A.item_cd = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S") & _
				  " and	A.rout_no = " & FilterVar(UCase(frm1.txtRoutNo.value), "''", "S") & _
				  "	and	A.rout_order in (" & FilterVar("F", "''", "S") & " ," & FilterVar("I", "''", "S") & " ) "	
	arrParam(5) = "����"			
	
	arrField(0) = "A.OPR_NO"	
	arrField(1) = "A.WC_CD"
	arrField(2) = "B.WC_NM"
	arrField(3) = "C.MINOR_NM"
	arrField(4) = "A.INSIDE_FLG"
	arrField(5) = "A.MILESTONE_FLG"
	arrField(6) = "A.INSP_FLG"
	
	arrHeader(0) = "����"		
	arrHeader(1) = "�۾���"	
	arrHeader(2) = "�۾����"
	arrHeader(3) = "�����۾���"
	arrHeader(4) = "�系����"
	arrHeader(5) = "Milestone"
	arrHeader(6) = "�˻翩��"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtOprNo.Value = arrRet(0)
		frm1.txtOprNoDesc.Value	= arrRet(3)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtOprNo.focus
	
End Function

'------------------------------------------  OpenSpreadItem()  -------------------------------------------------
'	Name : OpenSpreadItem()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenSpreadItem(ByVal ItemCd)
	OpenSpreadItem = false
	
	Dim arrRet
	Dim arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrParam6
	Dim arrField(6)
	Dim iCalledAspName, IntRetCD
	
	With frm1

		If IsOpenPop = True Then Exit Function

		IsOpenPop = True

		arrParam1 = Trim(.hPlantCd.value)	' Plant Code
		arrParam2 = Trim(.txtPlantNm.Value)	' Plant Name
		arrParam3 = Trim(ItemCd)
		arrParam4 = ""	'Trim(.txtItemNm.Value)	' Item Name
		arrParam5 = Trim(.cboInspClassCd.Value)
  		arrParam6 = "NO_STANDARD"	'ǰ��˻����� �߰� 
  
		If Trim(.cboInspClassCd.Value) = "P" Then
			iCalledAspName = AskPRAspName("q1211pa4")
			If Trim(iCalledAspName) = "" Then
				IntRetCD = DisplayMsgBox("900040", "X", "q1211pa4", "X")
				IsOpenPop = False
				Exit Function
			End If
		Else
			iCalledAspName = AskPRAspName("q1211pa2")
			If Trim(iCalledAspName) = "" Then
				IntRetCD = DisplayMsgBox("900040", "X", "q1211pa2", "X")
				IsOpenPop = False
				Exit Function
			End If
		End if

	
		arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrParam6, arrField), _
		              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
			  
		IsOpenPop = False

		If arrRet(0) <> "" Then
			.vspdData2.Col  = C_ItemCd
			.vspdData2.Text = arrRet(0)			
			.vspdData2.Col  = C_ItemNm
			.vspdData2.Text = arrRet(1)
		End If	
		
		Call SetActiveCell(.vspdData2,C_ItemCd,.vspdData2.ActiveRow,"M","X","X")
		Set gActiveElement = document.activeElement
	End With

	OpenSpreadItem = true
End Function

'------------------------------------------  OpenSpreadRoutNo()  -------------------------------------------------
'	Name : OpenSpreadRoutNo()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenSpreadRoutNo(Byval pvItemCd, ByVal pvRoutNo)
	OpenSpreadRoutNo = false
	
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	If pvItemCd= "" Then
		Call DisplayMsgBox("971012","X", "ǰ��","X")
		Call SetActiveCell(frm1.vspdData2, C_ItemCd, frm1.vspdData2.ActiveRow, "M","X","X")
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
		
	arrParam(0) = "����� �˾�"					' �˾� ��Ī 
	arrParam(1) = "P_ROUTING_HEADER"				' TABLE ��Ī 
	arrParam(2) =  pvRoutNo							' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "P_ROUTING_HEADER.PLANT_CD = " & FilterVar(UCase(frm1.hPlantCd.value), "''", "S") & _
				  "And P_ROUTING_HEADER.ITEM_CD = " & FilterVar(UCase(pvItemCd), "''", "S") 	
	arrParam(5) = "�����"			
	
    arrField(0) = "ROUT_NO"							' Field��(0)
    arrField(1) = "DESCRIPTION"						' Field��(1)
    arrField(2) = "BOM_NO"							' Field��(1)
    arrField(3) = "MAJOR_FLG"						' Field��(1)
   
    arrHeader(0) = "�����"						' Header��(0)
    arrHeader(1) = "����ø�"					' Header��(1)
    arrHeader(2) = "BOM Type"					' Header��(1)
    arrHeader(3) = "�ֶ����"					' Header��(1)        
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.vspdData2.Col  = C_RoutNo
		frm1.vspdData2.Text = arrRet(0)			
		frm1.vspdData2.Col  = C_RoutNoDesc
		frm1.vspdData2.Text = arrRet(1)
	End If	
		
	Call SetActiveCell(frm1.vspdData2, C_RoutNo,frm1.vspdData2.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement

	OpenSpreadRoutNo = true
End Function

'------------------------------------------  OpenSpreadOprNo()  -------------------------------------------------
'	Name : OpenSpreadOprNo()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenSpreadOprNo(Byval pvItemCd, ByVal pvRoutNo, ByVal pvOprNo)
	OpenSpreadOprNo = false
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function    

	IsOpenPop = True

	If pvItemCd = "" Then
		Call DisplayMsgBox("971012","X", "ǰ��","X")
		Call SetActiveCell(frm1.vspdData2, C_ItemCd, frm1.vspdData2.ActiveRow, "M","X","X")
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If	
	
	If pvRoutNo = "" Then
		Call DisplayMsgBox("971012","X", "�����","X")
		Call SetActiveCell(frm1.vspdData2, C_RoutNo, frm1.vspdData2.ActiveRow, "M","X","X")
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If	

	arrParam(0) = "�����˾�"	
	arrParam(1) = "P_ROUTING_DETAIL A inner join P_WORK_CENTER B on A.wc_cd = B.wc_cd and A.plant_cd = B.plant_cd " & _
				  " left outer join B_MINOR C on A.job_cd = C.minor_cd and C.major_cd = " & FilterVar("P1006", "''", "S") & ""				
	arrParam(2) = UCase(pvOprNo)
	arrParam(3) = ""
	arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.hPlantCd.value), "''", "S") & _
				  " and	A.item_cd = " & FilterVar(UCase(pvItemCd), "''", "S") & _
				  " and	A.rout_no = " & FilterVar(UCase(pvRoutNo), "''", "S") & _
				  "	and	A.rout_order in (" & FilterVar("F", "''", "S") & " ," & FilterVar("I", "''", "S") & " ) "	
	arrParam(5) = "����"			
	
	arrField(0) = "A.OPR_NO"	
	arrField(1) = "A.WC_CD"
	arrField(2) = "B.WC_NM"
	arrField(3) = "C.MINOR_NM"
	arrField(4) = "A.INSIDE_FLG"
	arrField(5) = "A.MILESTONE_FLG"
	arrField(6) = "A.INSP_FLG"
	
	arrHeader(0) = "����"		
	arrHeader(1) = "�۾���"	
	arrHeader(2) = "�۾����"
	arrHeader(3) = "�����۾���"
	arrHeader(4) = "�系����"
	arrHeader(5) = "Milestone"
	arrHeader(6) = "�˻翩��"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.vspdData2.Col  = C_OprNo
		frm1.vspdData2.Text = arrRet(0)			
		frm1.vspdData2.Col  = C_OprNoDesc
		frm1.vspdData2.Text = arrRet(3)
	End If	
		
	Call SetActiveCell(frm1.vspdData2, C_OprNo,frm1.vspdData2.ActiveRow,"M","X","X")
	Set gActiveElement = document.activeElement
	
	OpenSpreadOprNo = true
End Function

 '------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	OpenPlant = false
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
	
	arrHeader(0) = "�����ڵ�"		
	arrHeader(1) = "�����"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
	End If	
	
	frm1.txtPlantCd.Focus
	Set gActiveElement = document.activeElement
	OpenPlant = true
End Function

'------------------------------------------  ()  --------------------------------------------------
'	Name : Checkall()
'	Description : ��ü����Button Click�� ��üCheck Box���� 
'--------------------------------------------------------------------------------------------------------- 
Function Checkall()
	
	Checkall = false

	Dim IRowCount
	Dim lngMaxRows
	Dim IntRetCD
	
	lngMaxRows = frm1.vspdData.Maxrows
	If lngMaxRows < 1 Then
		IntRetCD = DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End If
	 
	ggoSpread.Source = frm1.vspdData
	 
	With frm1.vspdData
		
		IF lgCheckall = 0 Then 

			For IRowCount = 1 to lngMaxRows
				.Row = IRowCount 
				.Col = C_CheckItem	 
				.Text = 1     
			Next    

			lgCheckall = 1	
		
		Else
			   
			For IRowCount = 1 to lngMaxRows
				.Row = IRowCount 
				.Col = C_CheckItem	 
				.Text = 0     
			Next    
			   
			lgCheckall = 0
		  
		End If	    
 
	End With
	 
	lgChangeFlag = True  
	Checkall = True

End Function


'=============================================  2.5.1 LoadInspStand()  ======================================
'=	Event Name : LoadInspStand
'=	Event Desc :
'========================================================================================================
Function LoadInspStand()

	Dim intRetCD
	
	'/* 9�� ������ġ: �������� üũ���� ���� �� Link�� �Ѱ��ִ� ����Ÿ ���� - START */
	If lgSaveFlag = False Then
		ggoSpread.Source = frm1.vspdData2
		
		If (lgChangeFlag = True Or ggoSpread.SSCheckChange = True) Then
			IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then	Exit Function
		End If
		
	End if
	With frm1
		'�����ڵ�/��/ǰ���ڵ�/��/�˻�з��ڵ� 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspClassCd", Trim(.cboInspClassCd.value)
		If Trim(.cboInspClassCd.value) = "P" Then
			WriteCookie "txtRoutNo", Trim(.txtRoutNo.value)
			WriteCookie "txtRoutNoDesc", Trim(.txtRoutNoDesc.value)
			WriteCookie "txtOprNo", Trim(.txtOprNo.value)
			WriteCookie "txtOprNoDesc", Trim(.txtOprNoDesc.value)
		End if
		
		if .vspdData2.MaxRows > 0 then
			.vspdData2.Row = .vspdData2.ActiveRow
			.vspdData2.Col = C_ItemCd
			WriteCookie "txtItemCd", Trim(.vspdData2.Text)
			.vspdData2.Col = C_ItemNm
			WriteCookie "txtItemNm", Trim(.vspdData2.Text)
		Else
			WriteCookie "txtItemCd", Trim(.txtItemCd.value)
			WriteCookie "txtItemNm", Trim(.txtItemNm.value)
		End If
	
	End With
	
	'/* 9�� ������ġ: �������� üũ���� ���� �� Link�� �Ѱ��ִ� ����Ÿ ���� - END */
	PgmJump(BIZ_PGM_JUMP1_ID)
End Function

'==========================================  CopyInspStand()  ======================================
'	Name : CopyInspStand()
'	Description : 
'========================================================================================================= 
Function CopyInspStand()
	CopyInspStand = false
	Call fncSave()
	CopyInspStand = true
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     	'��: Load table , B_numeric_format
	Call ggoOper.LockField(Document, "N")                                   	'��: Lock  Suitable  Field

	Call InitVariables                                                      '��: Initializes local global variables
	Call InitSpreadSheet("*")
	Call InitComboBox
	Call SetDefaultVal
	
	Call EnableField(frm1.cboInspClassCd.value)
	Call ProtectField(frm1.cboInspClassCd.value)
	
    If Trim(frm1.txtPlantCd.Value) = "" Then
    	frm1.txtPlantCd.focus 
    Else
    	frm1.txtItemCd.focus 
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		If Row < 0 Then	Exit Sub

		If Row > 0 And Col = C_CheckItem Then
			.Row = Row
			.Col = Col
			IF .Text = "1" Then
				lgChangeFlag = true
			End If
		End If	
			
			
	End With

End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
	
	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then Exit Sub
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If 
End Sub

'======================================================================================================
'   Event Name : vspdData2_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strItemCd
	Dim strRoutNo
	Dim strOprNo
	With frm1.vspddata2
		If Row > 0 then
			Select Case Col 
			Case C_ItemPopup
				.Row = Row
				.Col = C_ItemCd
				'Call OpenSpreadItem(.text)	
				If OpenSpreadItem(.text) = false then Exit Sub
			Case C_RoutNoPopup
				.Row = Row
				.Col = C_ItemCd
				strItemCd = Trim(.text)
				.Col = C_RoutNo
				strRoutNo = Trim(.text)
				
				If OpenSpreadRoutNo(strItemCd, strRoutNo) = false then Exit Sub
			Case C_OprNoPopup
				.Row = Row
				.Col = C_ItemCd
				strItemCd = Trim(.text)
				.Col = C_RoutNo
				strRoutNo = Trim(.text)
				.Col = C_OprNo
				strOprNo = Trim(.text)
				
				If OpenSpreadOprNo(strItemCd, strRoutNo, strOprNo) = false then Exit Sub
			
			End Select
		End if		
	End With
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	gMouseClickStatus = "SPC"   
    
	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 

 	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
	
 	End If
 	
End Sub

'========================================================================================
' Function Name : vspdData2_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
 	gMouseClickStatus = "SP2C"   
    
	Call SetPopupMenuItemInf("1001011111")         'ȭ�麰 ���� 

 	Set gActiveSpdSheet = frm1.vspdData2
    
 	If frm1.vspdData2.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData2 
 		If lgSortKey2 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey2 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey2		'Sort in Descending
 			lgSortKey2 = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
	
 	End If
 	
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 

Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If
End Sub
 
'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData2_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	If NewCol = C_CheckItem or Col = C_CheckItem Then
		Cancel = True
		Exit Sub
	End If

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : vspdData2_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub
 
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim pvSpdNo
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    pvSpdNo = gActiveSpdSheet.id
    Call InitSpreadSheet(pvSpdNo)  
    
    If pvSpdNo = "A" Then
		ggoSpread.Source = frm1.vspdData
	Else
		ggoSpread.Source = frm1.vspdData2
	End If
	
    Call ggoSpread.ReOrderingSpreadData

	Call ggoOper.LockField(Document, "Q")		'��: This function lock the suitable field
	lgChangeFlag = false
	lgChangeFlag2 = false
End Sub 

'============================================= EnableField()  ======================================
'=	Event Name : EnableField
'=	Event Desc :
'========================================================================================================
Sub EnableField(Byval strInspClass)
	If	strInspClass = "P" Then
		Process.style.display	= ""
		Call ggoOper.SetReqAttr(frm1.txtRoutNo, "N")
		Call ggoOper.SetReqAttr(frm1.txtOprNo, "N")
	Else	
		Process.style.display	= "none"
		Call ggoOper.SetReqAttr(frm1.txtRoutNo, "Q")
		Call ggoOper.SetReqAttr(frm1.txtOprNo, "Q")
	End if
End Sub

'============================================= ProtectField()  ======================================
'=	Event Name : ProtectField
'=	Event Desc :
'========================================================================================================
Sub ProtectField(Byval strInspClass)
	ggoSpread.Source = frm1.vspdData2
	
	frm1.vspdData2.redraw = false
	If	strInspClass = "P" Then
		Call ggoSpread.SSSetColHidden(C_RoutNo, C_RoutNo, False)
		Call ggoSpread.SSSetColHidden(C_RoutNoPopup, C_RoutNoPopup, False)
		Call ggoSpread.SSSetColHidden(C_RoutNoDesc, C_RoutNoDesc, False)
		Call ggoSpread.SSSetColHidden(C_OprNo, C_OprNo, False)
		Call ggoSpread.SSSetColHidden(C_OprNoPopup, C_OprNoPopup, False)
		Call ggoSpread.SSSetColHidden(C_OprNoDesc, C_OprNoDesc, False)
		
		ggoSpread.SpreadUnLock C_RoutNo, -1, C_RoutNo, -1
		ggoSpread.SpreadUnLock C_OprNo, -1, C_OprNo, -1
		
		ggoSpread.SSSetRequired C_RoutNo, -1,  -1
		ggoSpread.SSSetRequired C_OprNo, -1,  -1
		
	Else	
		Call ggoSpread.SSSetColHidden(C_RoutNo, C_RoutNo, True)
		Call ggoSpread.SSSetColHidden(C_RoutNoPopup, C_RoutNoPopup, True)
		Call ggoSpread.SSSetColHidden(C_RoutNoDesc, C_RoutNoDesc, True)
		Call ggoSpread.SSSetColHidden(C_OprNo, C_OprNo, True)
		Call ggoSpread.SSSetColHidden(C_OprNoPopup, C_OprNoPopup, True)
		Call ggoSpread.SSSetColHidden(C_OprNoDesc, C_OprNoDesc, True)
		
		ggoSpread.SpreadLock C_RoutNo, -1, C_RoutNo, -1
		ggoSpread.SpreadLock C_OprNo, -1, C_OprNo, -1
		
		ggoSpread.SSSetProtected C_RoutNo, -1, -1
		ggoSpread.SSSetProtected C_OprNo, -1, -1
	End if
	frm1.vspdData2.redraw = True
End Sub

'============================================= cboInspClassCd_onchange()  ======================================
'=	Event Name : cboInspClassCd_onchange()
'=	Event Desc :
'========================================================================================================
Sub cboInspClassCd_onchange()
	Call EnableField(frm1.cboInspClassCd.value)
	Call ProtectField(frm1.cboInspClassCd.value)
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
	FncQuery = False                                                        '��: Processing is NG
	
	Err.Clear                                                            		   '��: Protect system from crashing
	
	If lgSaveFlag = False Then
		ggoSpread.Source = frm1.vspdData
		
		If (lgChangeFlag = True  or lgChangeFlag2 = True) Then
			IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then	Exit Function
		End If
		frm1.vspdData2.MaxRows = 0
		
	End If
	'-----------------------
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "2")  
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
	Call InitVariables
	
	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then	Exit Function
	
	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False then	Exit Function
	
	FncQuery = True						'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	Dim IntRetCD 
    
	FncNew = False                                                          '��: Processing is NG
	
	Err.Clear                                                               '��: Protect system from crashing
	'On Error Resume Next                                                    '��: Protect system from crashing
	ggoSpread.Source = frm1.vspdData
	'-----------------------
	'Check previous data area
	'-----------------------
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then	Exit Function
	
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	
	Call ggoOper.ClearField(Document, "A")
	Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field
	Call InitVariables                                                      '��: Initializes local global variables
	Call SetDefaultVal
	Call EnableField(frm1.cboInspClassCd.value)
	Call ProtectField(frm1.cboInspClassCd.value)
	
	If Trim(frm1.txtPlantCd.value) =  "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtItemCd.focus 
	End If

	FncNew = True 

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	Dim IntRetCD 
	
	FncDelete = False                                                       '��: Processing is NG
	
	Err.Clear                                                               '��: Protect system from crashing
	'On Error Resume Next                                                    '��: Protect system from crashing
	
	'-----------------------
	'Precheck area
	'-----------------------
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")  
		Exit Function
	End If

	IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then	Exit Function
	
	'-----------------------
	'Delete function call area
	'-----------------------
	If DbDelete = False Then Exit Function
	
	FncDelete = True                                                        '��: Processing is OK

End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
	Dim blnReturn
	Dim i
	
	FncSave = False                                                  		       '��: Processing is NG

	Err.Clear                                                            	 		  '��: Protect system from crashing
	
	On Error Resume Next                                           	       '��: Protect system from crashing
	
	With frm1.vspdData   
		'-----------------------
		'Precheck area
		'-----------------------
		If .MaxRows < 1 Then
			IntRetCD = DisplayMsgBox("900002", "X", "X", "X")
       		Exit Function
		End If
		
		blnReturn = False
		.Col = 1
		For i = 1 To .MaxRows
			.Row = i
			If .Value = 1 Then
				blnReturn = True
				Exit For
			End if
		Next
	End With
	
	If blnReturn = False Then
		IntRetCD = DisplayMsgBox("900025", "X", "X", "X")
       	Exit Function
	End If
	
	'-----------------------
	'Check content area
	'-----------------------
    ggoSpread.Source = frm1.vspdData2
	If ggoSpread.SSCheckChange = False Then 
	    IntRetCD = DisplayMsgBox("900024","X", "X", "X")                     
 		Exit Function
	End If

    If Not ggoSpread.SSDefaultCheck Then Exit Function
       	
    If frm1.vspdData2.Maxrows < 1 Then
       IntRetCD = DisplayMsgBox("229916", "X", "X", "X")
       	Exit Function
    End If
    	
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False then Exit Function                              '��: Save db data
    
	FncSave = True                                      	                    '��: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	FncCopy = false
	With frm1.vspdData2
		If .MaxRows < 1 then Exit function
		
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData2	
		ggoSpread.CopyRow
		SetSpreadColor2 .ActiveRow, .ActiveRow
	    .Row = .ActiveRow
	    .Col = C_ItemCd
	    .Text = ""
	    .Col = C_ItemNm
	    .Text = ""
	    .ReDraw = True                                   	'��: Protect system from crashing
	
		Call SetActiveCell(frm1.vspdData2,C_ItemCd,.ActiveRow,"M","X","X")
		Set gActiveElement = document.ActiveElement	
	End With
	FncCopy= true
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	
	FncCancel = false
	
	With frm1
	
		If .vspdData2.MaxRows < 1 then Exit function
	
		lgChangeFlag2 = false
	
		ggoSpread.Source = .vspdData2	
		ggoSpread.EditUndo 
	End With
	
	FncCancel = true                                                 '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim IntRetCD
	Dim imRow
	
	On Error Resume Next
	 
	FncInsertRow = false
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)

	Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then Exit Function
	End If
	
	With frm1
	
		lgChangeFlag2 = true
		.vspdData2.focus
		ggoSpread.Source = .vspdData2
		'.vspdData.EditMode = True
		.vspdData2.ReDraw = False
		ggoSpread.InsertRow .vspdData2.ActiveRow, imRow
    	SetSpreadColor2 .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow -1
		.vspdData2.ReDraw = True
		
		Call SetActiveCell(.vspdData2,C_ItemCd,.vspdData2.ActiveRow,"M","X","X")
		Set gActiveElement = document.ActiveElement

    End With
  
    If Err.number = 0 Then FncInsertRow = True
    
End Function


'========================================================================================
' Function Name : FncSplitColumn 
' Function Desc : vspdData (Grid1)
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
	FncPrev =false
    	'On Error Resume Next                                                   '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
	FncNext = false
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)					'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : This function is related to Exit 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
	
	If lgSaveFlag = False Then
		ggoSpread.Source = frm1.vspdData2
		If (lgChangeFlag = True  or lgChangeFlag2 = True) Then
			IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then	Exit Function
		End If
	End if
	
	FncExit = True
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	
	Err.Clear                                                               			'��: Protect system from crashing
	
	Call LayerShowHide(1)
	
	DbQuery = False
	
	With frm1	
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
									& "&txtPlantCd=" & .hPlantCd.value _
									& "&txtItemCd=" & .hItemCd.value _
									& "&cboInspClassCd=" & .hInspClassCd.value _
									& "&txtRoutNo=" & .hRoutNo.value _
									& "&txtOprNo=" & .hOprNo.value _
									& "&lgStrPrevKey=" & lgStrPrevKey _
									& "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001 _
									& "&txtPlantCd=" & Trim(.txtPlantCd.Value) _
									& "&txtItemCd=" & Trim(.txtItemCd.value) _
									& "&cboInspClassCd=" & Trim(.cboInspClassCd.Value) _
									& "&txtRoutNo=" & Trim(.txtRoutNo.value) _
									& "&txtOprNo=" & Trim(.txtOprNo.value) _
									& "&lgStrPrevKey=" & lgStrPrevKey _
									& "&txtMaxRows=" & .vspdData.MaxRows
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)					'��: �����Ͻ� ASP �� ���� 
		
		DbQuery = True                                      '��: Processing is NG
	End With

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()									'��: ��ȸ ������ ������� 
	DbQueryOk = false				'��: ��ȸ ������ ������� 
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = Parent.OPMD_UMODE			'��: Indicates that current mode is Update mode
	Call SetToolBar("11100101000111")							'��: ��ư ���� ���� 
	
	Call ggoOper.LockField(Document, "Q")		'��: This function lock the suitable field
	Call EnableField(frm1.cboInspClassCd.value)
	Call ProtectField(frm1.cboInspClassCd.value)
	
	lgChangeFlag = false
	lgChangeFlag2 = false
	
	Call ggoSpread.SpreadLock(C_ItemPopup, -1, C_ItemPopup)
	
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False
	ggoSpread.SpreadUnLock C_CheckItem, -1, C_CheckItem, -1
	frm1.vspdData.ReDraw = True
	
	DbQueryOk =true
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
	Dim lRow        
	Dim lGrpCnt     
	Dim strVal 
	Dim arrVal
	
	Call LayerShowHide(1)
	lgSaveFlag = False
	DbSave = False                                                          '��: Processing is NG
    
	On Error Resume Next                                                   '��: Protect system from crashing

	With frm1
		.txtMode.value = Parent.UID_M0002
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 0
    	Redim arrVal(0)
    	'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
			If GetSpreadText(.vspdData,C_CheckItem,lRow,"X","X") = "1" Then
				Redim Preserve arrVal(lGrpCnt)
				arrVal(lGrpCnt) = GetSpreadText(.vspdData,C_InspItemCd,lRow,"X","X") & Parent.gColSep & Parent.gRowSep
				lGrpCnt = lGrpCnt + 1
			End If	
		Next	
		'Target Item
		.txtMaxRows.value = lGrpCnt - 1
		.txtSpread.value = Join(arrVal, "")
		
		lGrpCnt = 0
    	Redim arrVal(0)
		
		For lRow=1 to .vspdData2.MaxRows
			If GetSpreadText(.vspdData2,0,lRow,"X","X") <> "" Then
				Redim Preserve arrVal(lGrpCnt)
				arrVal(lGrpCnt) = GetSpreadText(.vspdData2,C_ItemCd,lRow,"X","X") & Parent.gColSep & _
								  GetSpreadText(.vspdData2,C_RoutNo,lRow,"X","X") & Parent.gColSep & _
								  GetSpreadText(.vspdData2,C_OprNo,lRow,"X","X") & Parent.gColSep & _
								  CStr(lRow) & Parent.gRowSep
				
				lGrpCnt = lGrpCnt + 1
			End If	
		Next
		
		.txtMaxRows2.value = lGrpCnt-1
		.txtSpread2.value = Join(arrVal, "")
		
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)					'��: �����Ͻ� ASP �� ���� 
	End With
	
	DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()							'��: ���� ������ ���� ���� 
	DbSaveOk = false				            '��: ���� ������ ���� ���� 
   	lgSaveFlag = True

	With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			.txtPlantCd.Value = .hPlantCd.value
			.txtItemCd.Value = .hItemCd.value
			.cboInspClassCd.Value = .hInspClassCd.value
			.txtRoutNo.value = .hRoutNo.value
			.txtOprNo.value	= .hOprNo.value
		End If
	
		Dim lRow

		Call SetSpreadColor1(1, .vspdData2.MaxRows)
	
		For lRow=1 to .vspdData2.MaxRows
			.vspdData2.Col=0
			.vspdData2.Row = lRow
			.vspdData2.Text = ""
		Next
		
	End With
	DbSaveOk = false
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()		
	DbDelete = false
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ�� �˻���� ����</font></td>
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
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% <%=LR_SPACE_TYPE_40%>>
								<TR>
									<td CLASS="TD5" NOWPAP>����</td>
									<td CLASS="TD6" NOWPAP>
										<input TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="����" tag="12XXXU" ><IMG align=top height=20 name=btnPlantCd1 onclick=vbscript:OpenPlant() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtPlantNm" SIZE="20" tag="14" >
									</td>
									<td CLASS="TD5" NOWPAP>�˻�з�</td>
									<td CLASS="TD6" NOWPAP>
										<SELECT Name="cboInspClassCd" ALT="�˻�з�" STYLE="WIDTH: 150px" tag="12"></SELECT>
									</td>
								</TR>
								<TR>
									<td CLASS="TD5" NOWPAP>ǰ��</td>
									<td CLASS="TD6" NOWPAP>
										<input TYPE=TEXT NAME="txtItemCd" SIZE="15" MAXLENGTH="18" ALT="ǰ��" tag="12XXXU" ><IMG align=top height=20 name=btnItemCd1 onclick=vbscript:OpenItem() src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtItemNm" SIZE="20" tag="14" >
									</td>
									<td CLASS="TD5" NOWPAP>&nbsp;</td>
									<td CLASS="TD6" NOWPAP>&nbsp;</td>										
								</TR>
								<TR ID="Process">
					      			<TD CLASS="TD5" NOWPAP>�����</TD>
									<TD CLASS="TD6" NOWPAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=12 MAXLENGTH=20 tag="12XXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRoutNo()">&nbsp;<input TYPE=TEXT NAME="txtRoutNoDesc" SIZE="30" tag="14"></TD>
									<TD CLASS="TD5" NOWPAP>����</TD>
									<TD CLASS="TD6" NOWPAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE=10 MAXLENGTH=3 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprNo()">&nbsp;<input TYPE=TEXT NAME="txtOprNoDesc" SIZE="30" tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH="100%" HEIGHT="100%" <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/q1216ma1_A_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH="100%" HEIGHT="100%" <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/q1216ma1_B_vspdData2.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<TR HEIGHT=20>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<BUTTON NAME="btnSelect" CLASS="CLSMBTN" ONCLICK="vbscript:CheckAll()">��ü ����/���</BUTTON>&nbsp;
	        			<BUTTON NAME="btnCopy" CLASS="CLSMBTN" ONCLICK="vbscript:CopyInspStand()">ǰ������ ����</BUTTON>
	        		</TD>
   		     		<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadInspStand">�˻����</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
    			</TR>
    		</TABLE>
    	</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm"  tabindex=-1 WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex=-1></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" tabindex=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows2" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hInspClassCd" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hOprNo" tag="24" tabindex=-1>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
