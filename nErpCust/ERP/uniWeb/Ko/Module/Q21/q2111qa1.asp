<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2111QA1
'*  4. Program Name         : �Ϻ���ȸ 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2003/08/04
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Ahn Jung Je
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

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit														'��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim strInspClass
Dim IsOpenPop
'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID		= "q2111qb1.asp"                 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMP_ID	= "q2111ma1"                     '��: Cookie���� ����� ��� 

Dim C_InspReqNo
Dim C_InspResultNo
Dim C_ReleaseDt
Dim C_InspDt
Dim C_ItemCd
Dim C_ItemNm
Dim C_BPCd
Dim C_BPNm
Dim C_LotNo
Dim C_LotSubNo
Dim C_MinorCd
Dim C_MinorNm
Dim C_LotSize
Dim C_InspQty
Dim C_DefectQty
Dim C_DefectRatio

'--------------- ������ coding part(��������,End)-------------------------------------------------------------

'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
Dim CompanyYMD
CompanyYMD = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gDateFormat)                                           '��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ -----
'--------------- ������ coding part(�������,End)------------------------------------------------------------- 

'==========================================  InitComboBox()  ======================================
'	Name : InitComboBox()
'	Description : Init ComboBox
'==================================================================================================
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("Q0010", "''", "S") & " ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboDecision , lgF0, lgF1, Chr(11))
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
	lgBlnFlgChgValue = False
	IsOpenPop = False
    '###�˻�з��� ����κ� Start###
    strInspClass = "R"
	'###�˻�з��� ����κ� End###	
End Sub                          

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtDtFr.Text	= CompanyYMD
	frm1.txtDtTo.Text	= CompanyYMD
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q", "NOCOOKIE","QA") %>
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "�����˾�"						' �˾� ��Ī 
	arrParam(1) = "B_Plant"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""									' Name Condition
	arrParam(4) = ""
	arrParam(5) = "����"							' TextBox ��Ī 

    arrField(0) = "Plant_Cd"					' Field��(0)
    arrField(1) = "Plant_NM"					' Field��(1)
        
    arrHeader(0) = "�����ڵ�"						' Header��(0)
    arrHeader(1) = "�����"							' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtPlantCd.Focus
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
	Else
		Exit Function
	End If	
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenInspReqNo()  -------------------------------------------------
'	Name : OpenInspReqNo()
'	Description : InspReqNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspReqNo()        
	OpenInspReqNo = false
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'���������� �ʿ��մϴ� 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If
	
	IsOpenPop = True
	
	Param1 = Trim(frm1.txtPlantCd.value)		
	Param2 = Trim(frm1.txtPlantNm.Value)	
	Param3 = Trim(frm1.txtInspReqNo.Value)	
	'###�˻�з��� ����κ� Start###	
	Param4 = strInspClass 		'�˻�з� 
	'###�˻�з��� ����κ� End###
	Param5 = ""			'���� 
	Param6 = "R"			'�˻�������� 
	
	iCalledAspName = AskPRAspName("Q4111pa1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "Q4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1, Param2, Param3, Param4, Param5, Param6), _
		"dialogWidth=820px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	frm1.txtInspReqNo.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtInspReqNo.Value    = arrRet(0)		
		frm1.txtInspReqNo.Focus		
	End If	
	
	Set gActiveElement = document.activeElement
	OpenInspReqNo = true
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	'�����ڵ尡 �ִ� �� üũ 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705","X","X","X") 		'���������� �ʿ��մϴ� 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function	
	End If

	IsOpenPop = True
	
	arrParam(0) = "ǰ���˾�"							' �˾� ��Ī 
	arrParam(1) = "B_Item_By_Plant a, B_Item b"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtItemCd.Value)				' Code Condition
	arrParam(3) = ""										' Name Condition
	arrParam(4) = "a.Item_Cd = b.Item_Cd And a.Plant_Cd = " & FilterVar(frm1.txtPlantCd.Value, "''", "S")  			' Where Condition
	arrParam(5) = "ǰ��"								' TextBox ��Ī 
	
	arrField(0) = "a.Item_Cd"					' Field��(0)
	arrField(1) = "b.Item_NM"							' Field��(1)
	arrField(2) = "b.SPEC"								' Field��(2)
	
	arrHeader(0) = "ǰ���ڵ�"							' Header��(0)
	arrHeader(1) = "ǰ���"								' Header��(1)
	arrHeader(2) = "�԰�"								' Header��(2)
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtItemCd.Focus
	If Trim(arrRet(0)) <> "" Then
		frm1.txtItemCd.Value = Trim(arrRet(0))
		frm1.txtItemNm.Value = Trim(arrRet(1))
	Else
		Exit Function
	End If
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenBp()  -------------------------------------------------
'	Name : OpenBp()
'	Description : Supplier PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó�˾�"						' �˾� ��Ī 
	arrParam(1) = "B_Biz_Partner"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBpCd.Value)				' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "(BP_TYPE = " & FilterVar("CS", "''", "S") & " Or BP_TYPE = " & FilterVar("S", "''", "S") & " )"	' Where Condition	
	arrParam(5) = "����ó"							' TextBox ��Ī 
	
    arrField(0) = "BP_CD"								' Field��(0)
    arrField(1) = "BP_NM"								' Field��(1)
    
    arrHeader(0) = "����ó�ڵ�"						' Header��(0)
    arrHeader(1) = "����ó��"						' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtBpCd.Focus
	If Trim(arrRet(0)) <> "" Then
		frm1.txtBpCd.Value = Trim(arrRet(0))
		frm1.txtBpNm.Value = Trim(arrRet(1))
	Else
		Exit Function
	End If	
	Set gActiveElement = document.activeElement
End Function

'==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'==================================================================================================== 
Function CookiePage(Byval Kubun)
	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strCookie
	Dim ii,jj,kk
	Dim iSeq
	Dim IntRetCD
    Dim strTemp
    Dim arrVal
         	
    If Kubun = 1 Then								 'Jump�� ȭ���� �̵��� ��� 
		If  lgSaveRow <  1 Then
			IntRetCD = DisplayMsgBox("900002",Parent.VB_YES_NO,"X","X")
			Exit Function
		End If	
		
		Redim  lgMark(UBound(lgFieldNM)) 
		
		strCookie  = ""
		iSeq       = 0
		
		For ii = 0 to Parent.C_MaxSelList - 1 
			For jj = 0 to UBound(lgFieldNM) -1
				If lgPopUpR(ii,0) = lgFieldCD(jj) Then
					iSeq = iSeq + 1
					lgMark(jj) = "X"
					strCookie = strCookie & "" & TRIM(LGFIELDNM(JJ)) & "" & Parent.gRowSep
					frm1.vspdData.Row = lgSaveRow
					frm1.vspdData.Col = iSeq
					strCookie = strCookie & frm1.vspdData.Text & Parent.gRowSep
				
					kk = CInt(lgNextSeq(jj)) 
					If kk > 0 And kk <= UBound(lgFieldNM) Then 
						lgMark(kk - 1) = "X"
						iSeq = iSeq + 1
						
						strCookie = strCookie & "" & TRIM(LGFIELDNM(KK-1)) & "" & Parent.gRowSep
						frm1.vspdData.Row = lgSaveRow
						frm1.vspdData.Col = iSeq
						strCookie = strCookie & frm1.vspdData.Text & Parent.gRowSep
					End If    
					jj =  UBound(lgFieldNM)  + 100
				End If    
			Next
		Next      
		
		WriteCookie CookieSplit , strCookie
		
		'--------------- ������ coding part(�������,Start)--------------------------------------------------

		'--------------- ������ coding part(�������,End)----------------------------------------------------
		
		Call PgmJump(BIZ_PGM_JUMP_ID)
	
	ElseIf Kubun = 0 Then							 'Jump�� ȭ���� �̵��� ������� 
		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		Dim iniSep

		'--------------- ������ coding part(�������,Start)---------------------------------------------------
			
		If ReadCookie("txtPlantCd") <> "" Then
			frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
		End If
		
		If ReadCookie("txtPlantNm") <> "" Then
			frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
		End If	
				
		WriteCookie "txtPlantCd", ""
		WriteCookie "txtPlantNm", ""
		'--------------- ������ coding part(�������,End)---------------------------------------------------

		'If Err.number <> 0 Then
		'	Err.Clear
		'	WriteCookie CookieSplit , ""
		'	Exit Function 
		'End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF
End Function

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    
	Call InitSpreadPosVariables()
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030804", , Parent.gAllowDragDropSpread

	With frm1.vspdData
		.ReDraw = false
		.MaxCols = C_DefectRatio + 1
		.MaxRows = 0
		
 		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("6", "15", "0")
		Call AppendNumberPlace("7", "13", "2")
	
		ggoSpread.SSSetEdit  C_InspReqNo,	"�˻��Ƿڹ�ȣ",	15
		ggoSpread.SSSetEdit  C_InspResultNo,"SEQ",		   5
		ggoSpread.SSSetDate  C_ReleaseDt,	"Release��",  10, 2,Parent.gDateFormat  
		ggoSpread.SSSetDate  C_InspDt,		"�˻���",     10, 2,Parent.gDateFormat  
		ggoSpread.SSSetEdit  C_ItemCd,		"ǰ���ڵ�",   15
		ggoSpread.SSSetEdit  C_ItemNm,		"ǰ���",     20
		ggoSpread.SSSetEdit  C_BPCd,		"����ó�ڵ�", 10
		ggoSpread.SSSetEdit  C_BPNm,		"����ó��",   15
		ggoSpread.SSSetEdit  C_LotNo,		"��Ʈ��ȣ",	  20
		ggoSpread.SSSetFloat C_LotSubNo,	"����",		   5, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetEdit  C_MinorCd,     "����",		   5
		ggoSpread.SSSetEdit  C_MinorNm,		"����",		  10
		ggoSpread.SSSetFloat C_LotSize,		"��Ʈũ��",	  15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_InspQty,     "�˻��",	  15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_DefectQty,	"�ҷ���",	  15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat C_DefectRatio,	"�ҷ���(%)",  15, "7", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"

 		Call ggoSpread.SSSetColHidden(C_InspResultNo, C_InspResultNo, True)
 		Call ggoSpread.SSSetColHidden(C_MinorCd, C_MinorCd, True)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
	    ggoSpread.SSSetSplit2(2)  
		
		.ReDraw = true
		
    End With
End Sub

'==========================================  2.6.1 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()

	C_InspReqNo	= 1
	C_InspResultNo = 2
	C_ReleaseDt	= 3
	C_InspDt = 4
	C_ItemCd = 5
	C_ItemNm = 6
	C_BPCd = 7
	C_BPNm = 8
	C_LotNo = 9
	C_LotSubNo = 10
	C_MinorCd = 11
	C_MinorNm = 12
	C_LotSize = 13
	C_InspQty = 14
	C_DefectQty = 15
	C_DefectRatio = 16

End Sub

'==========================================  2.6.2 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
		C_InspReqNo		= iCurColumnPos(1)
		C_InspResultNo  = iCurColumnPos(2)
		C_ReleaseDt	= iCurColumnPos(3)
		C_InspDt	= iCurColumnPos(4)
		C_ItemCd		= iCurColumnPos(5)									
		C_ItemNm		= iCurColumnPos(6)
		C_BPCd			= iCurColumnPos(7)
		C_BPNm			= iCurColumnPos(8)
		C_LotNo			= iCurColumnPos(9)
		C_LotSubNo		= iCurColumnPos(10)
		C_MinorCd		= iCurColumnPos(11)
		C_MinorNm		= iCurColumnPos(12)
		C_LotSize		= iCurColumnPos(13)
		C_InspQty		= iCurColumnPos(14)
		C_DefectQty		= iCurColumnPos(15)
		C_DefectRatio	= iCurColumnPos(16)
				
 	End Select
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
	
	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitComboBox()
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")										'��: ��ư ���� ����	
'--------------- ������ coding part(�������,Start)----------------------------------------------------
   	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
	   	frm1.txtPlantNm.value = Parent.gPlantNm
	End IF
	frm1.txtPlantCd.focus
'--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode ) 
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col				
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		
 			lgSortKey = 1
 		End If
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

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
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
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
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
Sub PopRestoreSpreadColumnInf()	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)	
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
	 
	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
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

'==========================================================================================
'   Event Name : txtDtFr
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtDtFr_DblClick(Button)
	If Button = 1 Then
		frm1.txtDtFr.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDtFr.Focus 
	End If
End Sub

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtDtTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtDtTo.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDtTo.Focus 
	End If
End Sub

'==========================================================================================
'   Event Name : txtDtFr
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtDtFr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Function

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function txtDtTo_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Function

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtPlantCd_onChange()
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantNm.Value = ""
	End If
End Function

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtItemCd_onChange()
	If Trim(frm1.txtItemCd.Value) = "" Then
		frm1.txtItemNm.Value = ""
	End If
End Function

'==========================================================================================
'   Event Name : txtBpCd
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtBpCd_onChange()
	If Trim(frm1.txtBpCd.Value) = "" Then
		frm1.txtBpNm.Value = ""
	End If
End Function

'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery() 

    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then Exit Function
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then Exit Function								'��: This function check indispensable field
    
    If ValidDateCheck(frm1.txtDtFr, frm1.txtDtTo) = False Then 
   		frm1.txtDtFr.focus 
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData 

	If Name_check("A") = False Then
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
    Call InitVariables 	
    
    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False then Exit Function

    FncQuery = True															'��: Processing is OK
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
	Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
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
    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	Call LayerShowHide(1)
    
    With frm1

		strVal = BIZ_PGM_ID & "?txtPlantCd="   & Trim(.txtPlantCd.value) & _
							  "&txtDtFr="      & Trim(.txtDtFr.Text) & _
							  "&txtDtTo="      & Trim(.txtDtTo.Text) & _
							  "&txtInspReqNo=" & Trim(.txtInspReqNo.value) & _
							  "&txtLotNo="	   & Trim(.txtLotNo.value) & _
							  "&txtItemCd="    & Trim(.txtItemCd.value) & _
							  "&txtBpCd="	   & Trim(.txtBpCd.value) & _
							  "&cboDecision=" & Trim(.cboDecision.value) & _
							  "&txtMaxRows="   & .vspdData.MaxRows & _
							  "&lgStrPrevKey=" & lgStrPrevKey                      '��: Next key tag

		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolbar("11000000000111")
	lgBlnFlgChgValue = False
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : Name_Check
'========================================================================================
Function Name_Check(ByVal Check)

	Name_Check = False
	
	With frm1

		'-----------------------
		'Check Plant_Cd	 
		'-----------------------
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			.txtPlantNm.Value = ""
			Call DisplayMsgBox("125000","X","X","X")
			.txtPlantCd.focus 
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		.txtPlantNm.Value = lgF0(0)

		If Check = "A" Then
			'-----------------------
			'Check Item_Cd	 
			'-----------------------
			If Trim(.txtItemCd.value) <> "" Then
				
				If 	CommonQueryRs(" b.ITEM_NM "," B_ITEM_BY_PLANT a inner join B_ITEM b on a.ITEM_CD = b.ITEM_CD " , _
								" a.ITEM_CD = " & FilterVar(.txtItemCd.Value, "''", "S") & " AND a.PLANT_CD = " & FilterVar(.txtPlantCd.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
							
					lgF0 = Split(lgF0, Chr(11))
					.txtItemNm.Value = lgF0(0)
				Else
				
					If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(.txtItemCd.Value, "''", "S"), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
								
						lgF0 = Split(lgF0, Chr(11))
						.txtItemNm.Value = lgF0(0)
						Call DisplayMsgBox("122700","X","X","X")
						.txtItemCd.focus 
					Else
						.txtItemNm.Value = ""
						Call DisplayMsgBox("122600","X","X","X")
						.txtItemCd.focus 
					End If
					Exit Function
				End If
			End If
			 
			'-----------------------
			'Check BP_Cd	 
			'-----------------------
			If Trim(.txtBPCd.Value) <> "" Then
				If 	CommonQueryRs(" BP_NM, BP_TYPE "," B_BIZ_PARTNER ", " BP_CD = " & FilterVar(.txtBPCd.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
							
					lgF0 = Split(lgF0, Chr(11))
					lgF1 = Split(lgF1, Chr(11))
					.txtBPNm.Value = lgF0(0)
					If Trim(UCase(lgF1(0))) = "C" Then
						Call DisplayMsgBox("179020","X","X","X")
						.txtBPCd.focus 
						Exit Function
					End If
				Else
					.txtBPNm.Value = ""
					Call DisplayMsgBox("229927","X","X","X")
					.txtBPCd.focus 
					Exit Function
				End If
			End If
		End If
	End With
	
	Name_Check = True

End Function

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���԰˻��Ϻ���ȸ</font></td>
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
		<TD  WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
        									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="����" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE="20" MAXLENGTH=40 tag="14" ></TD>								
        									<TD CLASS="TD5" NOWRAP>�Ⱓ</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/q2111qa1_fpDateTime5_txtDtFr.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/q2111qa1_fpDateTime6_txtDtTo.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�˻��Ƿڹ�ȣ</TD>
        									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo" SIZE=20 MAXLENGTH=18 ALT="�˻��Ƿڹ�ȣ" tag="11XXXU"><IMG src="../../../CShared/image/btnPopup.gif" name=btnInspReqNo align=top  TYPE="BUTTON" width=16 height=20 onclick="vbscript:OpenInspReqNo()"></TD>
        									<TD CLASS="TD5" NOWRAP>��Ʈ��ȣ</TD>
							   		<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotNo" SIZE=20 MAXLENGTH=25 ALT="��Ʈ��ȣ" tag="11XXXU">
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="ǰ��" tag="11XXXU"><IMG src="../../../CShared/image/btnPopup.gif" name=btnItemCd align=top  TYPE="BUTTON" width=16 height=20 onclick="vbscript:OpenItem()">
															<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=20 ALT="����ó" tag="11XXXU"><IMG align=top height=20 name=btnBpCd onclick="vbscript:OpenBp()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">
															<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboDecision" ALT="����" STYLE="WIDTH: 150px" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWRAP></TD>
	        						<TD CLASS="TD6" NOWRAP></TD>	
	     							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TR>
							<TD HEIGHT=100% WIDTH=100% Colspan=2>
								<script language =javascript src='./js/q2111qa1_I273334807_vspdData.js'></script>
							</TD>	
						</TR>	
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
      		<TD WIDTH=100%>
      			<TABLE <%=LR_SPACE_TYPE_30%>>
        				<TR>        				
        					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN="RIGHT">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
    </DIV>
</BODY>
</HTML>
