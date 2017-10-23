<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q2113MA1
'*  4. Program Name         : �ҷ�������� 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : 
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

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "Q2113MB1.asp"								'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_QRY2_ID = "Q2113MB3.asp"								'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "Q2113MB2.asp"								'��: �����Ͻ� ���� ASP�� 
'/* ��ü ���� ���� - START */
Const BIZ_PGM_DEL_ID = "Q2113MB4.asp"
'/* ��ü ���� ���� - END */
Const BIZ_PGM_JUMP1_ID = "Q2111MA1"
Const BIZ_PGM_JUMP2_ID = "Q2114MA1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim C_InspItemCd '= 1
Dim C_InspItemNm '= 2
Dim C_InspSeries '= 3
Dim C_DefectQty1 '= 4

Dim C_DefectTypeCd '= 1
Dim C_DefectTypePopup '= 2
Dim C_DefectTypeNm '= 3
Dim C_DefectQty2 '= 4
Dim C_ParentRowNo '= 5
Dim C_Flag '= 6

Dim lgIntFlgModeM                 'Variable is for Operation Status

Dim lgStrPrevKeyM()			'Multi���� �������� ���� ���� 
Dim lglngHiddenRows()		'Multi���� �������� ���� ���� 

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim lgInspClassCd

Dim lgSortKey1
Dim lgSortKey2

Dim IsOpenPop		
Dim strInspClass	

Dim lgSpdHdrClicked	'2003-03-01 Release �߰� 

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntFlgModeM = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
        
    lgStrPrevKey1 = ""                            'initializes Previous Key
    lgStrPrevKey2 = ""                         'initializes Previous Key
    
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey1 = 2
    lgSortKey2 = 2
    
    '###�˻�з��� ����κ� Start###
    strInspClass = "R"
	'###�˻�з��� ����κ� End###
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()
	 If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
	End If
	
	If ReadCookie("txtPlantCd") <> "" Then
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If
		
	If ReadCookie("txtInspReqNo") <> "" Then
		frm1.txtInspReqNo.Value = ReadCookie("txtInspReqNo")
	End If
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtInspReqNo", ""
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()	
	Call InitSpreadPosVariables()
	
	With frm1
		ggoSpread.Source = .vspdData
		ggoSpread.Spreadinit "V20020101", , Parent.gAllowDragDropSpread
		
		.vspdData.Redraw = False
   	 	   	 	
   	 	.vspdData.MaxCols = C_DefectQty1 + 1
		.vspdData.MaxRows = 0

        '.vspdData.OperationMode = 3
		
		Call GetSpreadColumnPos("A")
		
		Call AppendNumberPlace("6", "3","0")
		
		ggoSpread.SSSetEdit C_InspItemCd, "�˻��׸��ڵ�",12, 0, -1, 5, 2				'��: Spread Sheet�� Column�� ��� 
		ggoSpread.SSSetEdit C_InspItemNm, "�˻��׸��", 20, 0, -1, 40
   		ggoSpread.SSSetFloat C_InspSeries, "����", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat C_DefectQty1, "�ҷ���", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec

 		Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)	    

		Call SetSpreadLock

		.vspdData.Redraw = True
	End With
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet2
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet2()
	Call InitSpreadPosVariables2()

	With frm1
		ggoSpread.Source = .vspdData2
		ggoSpread.Spreadinit "V20020101", , Parent.gAllowDragDropSpread

		.vspdData2.Redraw = False
		
		.vspdData2.MaxCols = C_Flag + 1
		.vspdData2.MaxRows = 0

		Call GetSpreadColumnPos("B")
		
		ggoSpread.SSSetEdit C_DefectTypeCd, "�ҷ������ڵ�", 12, 0, -1, 3, 2			'��: Spread Sheet�� Column�� ��� 
	    ggoSpread.SSSetButton C_DefectTypePopup
	    ggoSpread.SSSetEdit C_DefectTypeNm, "�ҷ�������", 20, 0, -1, 40
	    ggoSpread.SSSetFloat C_DefectQty2, "�ҷ���", 15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , , "Z"
	    ggoSpread.SSSetEdit C_ParentRowNo , "C_ParentRowNo", 5
	    ggoSpread.SSSetEdit C_Flag , "C_Flag", 5

		Call ggoSpread.MakePairsColumn(C_DefectTypeCd, C_DefectTypePopup)
		Call ggoSpread.SSSetColHidden(C_ParentRowNo,C_ParentRowNo, True)
 		Call ggoSpread.SSSetColHidden(C_Flag, C_Flag, True)
 		Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols, .vspdData2.MaxCols, True)		
   
		Call SetSpreadLock2
		
		.vspdData2.Redraw = True
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLock 1, -1, frm1.vspdData.MaxCols
End Sub

'======================================================================================================
' Function Name : SetSpreadLock2
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock2()
	Call ggoSpread.SpreadLock(C_ParentRowNo, -1, C_ParentRowNo)
	Call ggoSpread.SpreadLock(C_Flag, -1, C_Flag)
	Call ggoSpread.SpreadLock(frm1.vspdData2.MaxCols, -1, frm1.vspdData2.MaxCols)
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal lRow)
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_InspItemCd = 1
	C_InspItemNm = 2
	C_InspSeries = 3
	C_DefectQty1 = 4
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables2
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables2()
	C_DefectTypeCd = 1
	C_DefectTypePopup = 2
	C_DefectTypeNm = 3
	C_DefectQty2 = 4
	C_ParentRowNo = 5
	C_Flag = 6	
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

			C_InspItemCd = iCurColumnPos(1)
			C_InspItemNm = iCurColumnPos(2)
			C_InspSeries = iCurColumnPos(3)
			C_DefectQty1 = iCurColumnPos(4)
		Case "B"
 			ggoSpread.Source = frm1.vspdData2
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_DefectTypeCd = iCurColumnPos(1)
			C_DefectTypePopup = iCurColumnPos(2)
			C_DefectTypeNm = iCurColumnPos(3)
			C_DefectQty2 = iCurColumnPos(4)
			C_ParentRowNo = iCurColumnPos(5)
			C_Flag = iCurColumnPos(6)
 	End Select
End Sub

'======================================================================================================
' Function Name : SetSpread2ColorRange
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpread2ColorRange(ByVal Row, Byval Row2, Byval Mode)
	Dim i	
	ggoSpread.Source = frm1.vspdData2
	With ggoSpread
		If Mode = "I" Then
			.SpreadUnLock C_DefectTypeCd, Row, C_DefectTypeCd, Row2
			.SSSetRequired C_DefectTypeCd, Row, Row2
		Else			 
			.SSSetProtected C_DefectTypeCd, Row, Row2
			.SSSetProtected C_DefectTypePopup, Row, Row2
		End If
		
		If Mode <> "I" Then
			For i = Row To Row2
				frm1.vspdData2.Col = 0	
				frm1.vspdData2.Row = i
				If frm1.vspdData2.Text = ggoSpread.InsertFlag Then
					ggoSpread.SpreadUnLock C_DefectTypeCd, i, C_DefectTypeCd, i
					ggoSpread.SSSetRequired C_DefectTypeCd, i, i
					
					ggoSpread.SpreadUnLock C_DefectTypePopup, i, C_DefectTypePopup, i
				End If
			Next		
		End If
		
		.SSSetProtected C_DefectTypeNm, Row, Row2
	    ggoSpread.SpreadUnLock C_DefectQty2, Row, C_DefectQty2, Row2
	    .SSSetRequired C_DefectQty2, Row, Row2
	End With
End Sub

'======================================================================================================
'	Name : OpenPlant()
'	Description :Plant PopUp
'======================================================================================================
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
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtPlantCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
		frm1.txtPlantCd.Focus		
	End If	
	
	Set gActiveElement = document.activeElement
	OpenPlant = true
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
	Param6 = ""			'�˻�������� 
	
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

 '------------------------------------------  OpenDefectType()  -------------------------------------------------
'	Name : OpenDefectType()
'	Description : DefectType PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenDefectType(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("220705", "X", "X", "X")		'���������� �ʿ��մϴ� 
		Exit Function	
	End If
	
	IsOpenPop = True
'###�˻�з��� ����κ� Start###
	arrParam(0) = "�ҷ������˾�"
	arrParam(1) = "Q_DEFECT_TYPE"
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "" & _
	                    " AND INSP_CLASS_CD =  " & FilterVar(strInspClass , "''", "S") & ""
	arrParam(5) = "�ҷ�����"
	
    arrField(0) = "DEFECT_TYPE_CD"
    arrField(1) = "DEFECT_TYPE_NM"
    
    arrHeader(0) = "�ҷ������ڵ�"
    arrHeader(1) = "�ҷ�����"
'###�˻�з��� ����κ� End###    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	Call SetActiveCell(frm1.vspdData2,C_DefectTypeCd,frm1.vspdData2.ActiveRow,"M","X","X")
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData2.Col = C_DefectTypeCd
		frm1.vspdData2.Text = Trim(arrRet(0))
		frm1.vspdData2.Col = C_DefectTypeNm
		frm1.vspdData2.Text = Trim(arrRet(1))
		
		Call vspdData2_Change(frm1.vspdData2.Col, frm1.vspdData2.Row)						 ' ������ �о�ٰ� �˷��� 
		frm1.vspdData2.Col = C_DefectTypeCd
		frm1.vspdData2.Action = 0
		Call SetActiveCell(frm1.vspdData2,C_DefectTypeCd,frm1.vspdData2.ActiveRow,"M","X","X")
	End If	
	
	Set gActiveElement = document.activeElement
End Function

'=============================================  2.5.1 LoadInspection()  ======================================
'=	Event Name : LoadInspection
'=	Event Desc :
'========================================================================================================
Function LoadInspection()
	Dim intRetCD
	
	If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With
	
	PgmJump(BIZ_PGM_JUMP1_ID)
End Function

'=============================================  2.5.2 LoadDefectCause()  ======================================
'=	Event Name : LoadDefectCause
'=	Event Desc :
'========================================================================================================
Function LoadDefectCause()
	Dim intRetCD
	If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO,"X","X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
		
	End With
	PgmJump(BIZ_PGM_JUMP2_ID)
End Function

'=============================================  2.5.2 LoadDecision()  ======================================
'=	Event Name : LoadDecision
'=	Event Desc :
'========================================================================================================
Function LoadDecision()
	Dim intRetCD
	
	If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")
		If intRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	With frm1
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtInspReqNo", Trim(.txtInspReqNo.value)
	End With
	
	PgmJump(BIZ_PGM_JUMP2_ID)
End Function

'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2(ByVal Row, Byval NextQueryFlag)
	DbQuery2 = False
	
	Dim strVal
	Dim lngRet
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim iRow

	If Trim(frm1.hInspItemCd.Value) = "" Or Trim(frm1.hInspSeries.Value) = "" Then
		Exit Function
	End If
	
	'/* 9�� ������ġ: ���� ���������� �ణ �̵� �� �̹� ��ȸ�� �ڷᳪ �Էµ� �ڷḦ �о� ���� ������ '' â ���� - START */
	Call LayerShowHide(1)
	
	With frm1
		.vspdData.Row = CInt(Row)
		.vspdData.Col = .vspdData.MaxCols
		iRow = CInt(.vspdData.Text)
		
		If lglngHiddenRows(iRow - 1) <> 0 And NextQueryFlag = False Then
			.vspdData2.ReDraw = False
			lngRet = ShowFromData(iRow, lglngHiddenRows(iRow - 1))
			
			'/* ��ü ���� ���� - START */
			If lgIntFlgMode = Parent.OPMD_UMODE Then
				Call SetToolBar("11111111001111")
			Else
				Call SetToolBar("11101101001111")
			End If
			'/* ��ü ���� ���� - END */
			Call LayerShowHide(0)
						
'			'Spread Color ó��	'@@@
			lngRangeFrom = ShowDataFirstRow
			lngRangeTo = ShowDataLastRow		
						
			Call SetSpread2ColorRange(lngRangeFrom, lngRangeTo, "Q")
			.vspdData2.ReDraw = True
			DbQuery2 = True
			Exit Function
		End If
		
		If lgIntFlgModeM = Parent.OPMD_UMODE Then
			'@Query_Hidden
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtPlantCd=" & .hPlantCd.value			'��: ��ȸ ���� 
			strVal = strVal & "&txtInspReqNo=" & .hInspReqNo.value			'��: ��ȸ ���� 
			strVal = strVal & "&txtInspItemCd=" & .hInspItemCd.value			'��: ��ȸ ���� 
			strVal = strVal & "&txtInspSeries=" & .hInspSeries.value			'��: ��ȸ ���� 
			strVal = strVal & "&lgStrPrevKeyM=" & lgStrPrevKeyM(iRow - 1)
			strVal = strVal & "&lglngHiddenRows=" & lglngHiddenRows(iRow - 1)
			strVal = strVal & "&lRow=" & CStr(iRow)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		Else
			'@Query_Text        
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & Parent.UID_M0001	
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)		'��: ��ȸ ����				
			strVal = strVal & "&txtInspReqNo=" & .txtInspReqNo.value		'��: ��ȸ ����				
			'.vspdData.Row = Row	'
			.vspdData.Col = C_InspItemCd
			strVal = strVal & "&txtInspItemCd=" & .vspdData.Text			'��: ��ȸ ���� 
			.vspdData.Col = C_InspSeries
			strVal = strVal & "&txtInspSeries=" & .vspdData.Text			'��: ��ȸ ���� 
			strVal = strVal & "&lgStrPrevKeyM=" & lgStrPrevKeyM(iRow - 1)
			strVal = strVal & "&lglngHiddenRows=" & lglngHiddenRows(iRow - 1)
			strVal = strVal & "&lRow=" & CStr(iRow)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		End If
		
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)
	DbQuery2 = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryOk2(Byval DataCount)
	DbQueryOk2 = False
	
	Dim lngRangeFrom
	Dim lngRangeTo
	
	'/* ��ü ���� ���� - START */
	lgIntFlgMode = Parent.OPMD_UMODE
	'/* ��ü ���� ���� - END */
	
	With frm1.vspdData2
		lngRangeFrom = .MaxRows - DataCount + 1
		lngRangeTo = .MaxRows
		
		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Col = C_Flag
		
		.Col2 = C_Flag
		.DestCol = 0
		.DestRow = lngRangeFrom
		.Action = 19	'SS_ACTION_COPY_RANGE
		.BlockMode = False
	End With
	
	With frm1
		'Spread Color ó�� 
		Call SetSpread2ColorRange(.vspdData2.MaxRows - DataCount + 1, .vspdData2.MaxRows, "Q")
		.vspdData2.ReDraw = True
		
		lgIntFlgModeM = Parent.OPMD_UMODE
		'/* ��ü ���� ���� - START */
		Call SetToolBar("11111111001111")
		'/* ��ü ���� ���� - END */
		.vspdData.Focus
		
	End With
	
	DbQueryOk2 = True
End Function

'=======================================================================================================
'   Event Name : ParentGetFocusCell
'   Event Desc :
'=======================================================================================================
Sub ParentGetFocusCell(ByVal ParentRow, ByVal strDefectType, Byval strFlag)
	Dim CheckField1		
	Dim CheckField2
	Dim i 
	Dim lngStart
	Dim lngEnd
	
	With frm1.vspdData
		.Row = ParentRow
		.Col = .MaxCols
		.Row = CInt(.Text)
		.Col = 1
		.Action = 0		'Active Cell
	End With
	
	With frm1.vspdData2
		.ReDraw = False
		lngStart = ShowFromData(ParentRow, lglngHiddenRows(ParentRow - 1))
		.ReDraw = True
		lngEnd = lngStart + lglngHiddenRows(ParentRow - 1) - 1
		
		For i = lngStart To lngEnd
			.Row = i
			.Col = C_DefectTypeCd
			CheckField1 = .Text
			.Col = C_Flag
			CheckField2 = .Text
			If CheckField1 = strDefectType And CheckField2 = strFlag Then
				Exit For
			End If
		Next
		
'		Call SetSpread2ColorRange(lngStart, lngEnd, "I")
'		ggoSpread.SpreadUnLock C_DefectTypePopup, lngStart, C_DefectTypePopup, lngEnd

		Call SetSpread2ColorRange(i, i, "I")
		ggoSpread.SpreadUnLock C_DefectTypePopup, i, C_DefectTypePopup, i
				
		.Row = i			
		.Col = C_DefectQty2
		.Action = 0
		.Focus
	End With
	
	Set gActiveElement = document.activeElement
End Sub

'=======================================================================================================
'   Function Name : ShowFromData
'   Function Desc : 
'=======================================================================================================
Function ShowFromData(Byval Row, Byval lngShowingRows)
	ShowFromData = 0	
	
	Dim lngRow
	Dim lngStartRow
	
	With frm1.vspdData2
		
		Call SortSheet()
		
		'------------------------------------
		' Find First Row
		'------------------------------------ 
		lngStartRow = 0
		
		If .MaxRows < 1 Then Exit Function	
		
		For lngRow = 1 To .MaxRows
			.Row = lngRow
			.Col = C_ParentRowNo  
			If Row = Clng(.Text) Then
				lngStartRow = lngRow
				ShowFromData = lngRow
				Exit For
			End If
		Next
		
		'------------------------------------
		' Show Data
		'------------------------------------ 
		If lngStartRow > 0 Then
			.BlockMode = True
			.Row = 1
			.Row2 = .MaxRows
			.Col = C_Flag
			.Col2 = C_Flag
			.DestCol = 0
			.DestRow = 1
			.Action = 19	'SS_ACTION_COPY_RANGE
			.RowHidden = False
			
			.BlockMode = False
			
			'ex) ù��° �׸����� Ư�� Row�� �ش��ϴ� �ι�° �׸����� Row���� 10���϶� ������ �����Ͱ� 3��° ���� 6��°���� 4���̸� ù��° ���� 2��° ������ Row�� �����.
			If lngStartRow > 1 Then
				ggoSpread.SpreadLock 1, 1, .MaxCols, lngStartRow - 1	
				.BlockMode = True
				.Row = 1
				.Row2 = lngStartRow - 1
				.RowHidden = True
				.BlockMode = False
			End If

			'ex) ù��° �׸����� Ư�� Row�� �ش��ϴ� �ι�° �׸����� Row���� 10���϶� ������ �����Ͱ� 3��° ���� 6��°���� 4���̸� 7��° ���� ������ ������ Row�� �����.
			If lngStartRow < .MaxRows Then
				If lngStartRow + lngShowingRows <= .MaxRows Then
					ggoSpread.SpreadLock 1, lngStartRow + lngShowingRows, .MaxCols, .MaxRows	
					.BlockMode = True
					.Row = lngStartRow + lngShowingRows
					.Row2 = .MaxRows
					.RowHidden = True
					.BlockMode = False
				End If
			End If
			
			.BlockMode = False
			
			.Row = lngStartRow	'2003-03-01 Release �߰� 
			.Col = 0			'2003-03-01 Release �߰� 
			.Action = 0			'2003-03-01 Release �߰� 
		End If
	End With   
End Function

'======================================================================================================
' Function Name : SortSheet
' Function Desc : This function set Muti spread Flag
'=======================================================================================================
Function SortSheet()
	SortSheet = false
    
    With frm1.vspdData2
        .BlockMode = True
        .Col = 0
        .Col2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .SortBy = 0 'SS_SORT_BY_ROW
        
        .SortKey(1) = C_ParentRowNo
        .SortKey(2) = C_DefectTypeCd
        
        .SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
        .SortKeyOrder(2) = 1 'SS_SORT_ORDER_ASCENDING
        
        .Col = 1	'C_DefectTypeCd	
        .Col2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .Action = 25 'SS_ACTION_SORT
        
        .BlockMode = False
    End With        
    SortSheet = true
End Function

'=======================================================================================================
' Function Name : DefaultCheck
' Function Desc : 
'=======================================================================================================
Function DefaultCheck()
	DefaultCheck = False
	Dim i
	Dim j
	Dim RequiredColor 
	
	ggoSpread.Source = frm1.vspdData2
	RequiredColor = ggoSpread.RequiredColor
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				.Col = 0
				If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Then
					For j = 1 To .MaxCols
						.Col = j
						If .BackColor = RequiredColor Then
							If Len(Trim(.Text)) < 1 Then
								.Row = 0
								Call DisplayMsgBox("970021","X",.Text,"")
								.Row = i
								.Action = 0
								Exit Function
							End If
						End If			
					Next
				End If
			End If
		Next
	End With
	DefaultCheck = True
End Function

'=======================================================================================================
' Function Name : ChangeCheck
' Function Desc : 
'=======================================================================================================
Function ChangeCheck()
	ChangeCheck = False
	
	Dim i
	Dim strInsertMark
	Dim strDeleteMark
	Dim strUpdateMark
	
	ggoSpread.Source = frm1.vspdData2
	strInsertMark = ggoSpread.InsertFlag
	strDeleteMark = ggoSpread.UpdateFlag
	strUpdateMark = ggoSpread.DeleteFlag
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			.Col = 0
			If .Text = strInsertMark Or .Text = strDeleteMark Or .Text = strUpdateMark Then
				ChangeCheck = True
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : CheckDataExist
' Function Desc : 
'=======================================================================================================
Function CheckDataExist()
	CheckDataExist = False
	Dim i
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				CheckDataExist = True
				Exit Function
			End IF
		Next
	End With
End Function

'=======================================================================================================
' Function Name : ShowDataFirstRow
' Function Desc : 
'=======================================================================================================
Function ShowDataFirstRow()
	ShowDataFirstRow = 0
	Dim i
	
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				ShowDataFirstRow = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : ShowDataLastRow
' Function Desc : 
'=======================================================================================================
Function ShowDataLastRow()
	ShowDataLastRow = 0
	Dim i
	
	With frm1.vspdData2
		For i = .MaxRows To 1 Step -1
			.Row = i
			If .RowHidden = False Then
				ShowDataLastRow = i
				Exit Function
			End If
		Next
	End With
End Function

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
	If y<20 Then			'2003-03-01 Release �߰� 
	    lgSpdHdrClicked = 1 
	End If
	
    If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
    End If
End Sub 

'========================================================================================
' Function Name : vspdData2_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
    End If
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	
 	gMouseClickStatus = "SPC"   

 	Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	
	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
	If DefaultCheck = False Then	'2003-03-01 Release �߰� 
    	Exit Sub
	End If  	
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData
 		
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		ElseIf lgSortKey1 = 2 Then
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
 		lgIntFlgModeM = Parent.OPMD_CMODE
 		
 		lgSpdHdrClicked = 0		'2003-03-01 Release �߰� 
 		Call Sub_vspdData_ScriptLeaveCell(0, 0, Col, frm1.vspdData.ActiveRow, False)
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
 	Dim lngStartRow
 	Dim iActiveRow
 	Dim iActiveRow2
 	gMouseClickStatus = "SP2C"   

	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	
 	Set gActiveSpdSheet = frm1.vspdData2
	iActiveRow2 = frm1.vspdData2.ActiveRow
 	If frm1.vspdData2.IsVisible(C_DefectTypeCd, iActiveRow2, True) = False AND frm1.vspdData2.IsVisible(C_DefectTypePopup, iActiveRow2, True) = False AND frm1.vspdData2.IsVisible(C_DefectTypeNm, iActiveRow2, True) = False AND frm1.vspdData2.IsVisible(C_DefectQty2, iActiveRow2, True) = False Then
		Call SetPopupMenuItemInf("1101011011")         'ȭ�麰 ���� 
 		Exit Sub
 	Else
		Call SetPopupMenuItemInf("1101111111")         'ȭ�麰 ����	 	
 	End If
 	
 	If Row <= 0 AND Col <> 0 Then	'2003-03-01 Release �߰� 
 		ggoSpread.Source = frm1.vspdData2

 		frm1.vspdData.Row = frm1.vspdData.ActiveRow
 		frm1.vspdData.Col = frm1.vspdData.MaxCols

 		iActiveRow = CInt(frm1.vspdData.Text)
 		
 		frm1.vspdData2.Redraw = False
		lngStartRow = CInt(ShowFromData(iActiveRow, CInt(lglngHiddenRows(iActiveRow - 1))))
		frm1.vspdData2.Redraw = True
		
		If lgSortKey2 = 1 Then
 			ggoSpread.SSSort Col, lgSortKey2, lngStartRow, lngStartRow + CInt(lglngHiddenRows(iActiveRow - 1)) - 1	'Sort in Ascending
 			lgSortKey2 = 2
 		ElseIf lgSortKey2 = 2 Then
 			ggoSpread.SSSort Col, lgSortKey2, lngStartRow, lngStartRow + CInt(lglngHiddenRows(iActiveRow - 1)) - 1	'Sort in Descending
 			lgSortKey2 = 1
		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
        
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'========================================================================================
' Function Name : vspdData2_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
        
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData2_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                         'Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)                     
    Call ggoOper.LockField(Document, "N")                                       'Lock  Suitable  Field
	Call InitSpreadSheet                                                        'Setup the Spread sheet1
	Call InitSpreadSheet2
    Call InitVariables                                                          'Initializes local global variables
    Call SetDefaultVal
    '/* ��ü ���� ���� - START */
    Call SetToolBar("11100000000001")
    '/* ��ü ���� ���� - END */
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtInspReqNo.focus 
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
	Dim iActiveRow
	Dim iConvActiveRow
	Dim lngRangeFrom
	Dim lngRangeTo	
	Dim lRow
	Dim i
	Dim strFlag
	Dim strParentRowNo
	
    ggoSpread.Source = gActiveSpdSheet
    If gActiveSpdSheet.Name = "vspdData" Then
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet
		Call ggoSpread.ReOrderingSpreadData
    ElseIf gActiveSpdSheet.Name = "vspdData2" Then
		For i = 1 To frm1.vspdData2.MaxRows
			frm1.vspdData2.Row = i
			frm1.vspdData2.Col = 0
			strFlag = frm1.vspdData2.Text
			If strFlag = ggoSpread.InsertFlag Then
				frm1.vspdData2.Col = C_ParentRowNo
				strParentRowNo = CInt(frm1.vspdData2.Text)
				lglngHiddenRows(strParentRowNo - 1) = CInt(lglngHiddenRows(strParentRowNo - 1)) - 1
			End IF
		Next

		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet2
		frm1.vspdData2.Redraw = False
		
		Call ggoSpread.ReOrderingSpreadData("F")
		
		Call DbQuery2(frm1.vspdData.ActiveRow,False)
		
		lngRangeFrom = Clng(ShowDataFirstRow)
		lngRangeTo = Clng(ShowDataLastRow)
		
		Call SetSpread2ColorRange(lngRangeFrom, lngRangeTo, "Q")
		frm1.vspdData2.Redraw = True
    End If
    
 	'------ Developer Coding part (Start)	
 	'------ Developer Coding part (End) 
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )    
End Sub

'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)	
	If lgSpdHdrClicked = 1 Then	'2003-03-01 Release �߰� 
		Exit Sub
	End If
	
	Call Sub_vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)	
End Sub

'=======================================================================================================
'   Event Name : Sub_vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub Sub_vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)	
	'/* 9�� ������ġ : ������ Ű���� �Է��� ä �ٸ� ��������� �ű��� ���ϵ��� �������� ���� �߰� - START */
	Dim lRow		
	
	'/* 9�� ������ġ : ������ Ű���� �Է��� ä �ٸ� ��������� �ű��� ���ϵ��� �������� ���� �߰� - END */
	
	If Row <> NewRow And NewRow > 0 Then
		With frm1
			'/* 8�� ������ġ : ���� �������忡 �ʼ��Է� �ʵ� üũ - START */
			If DefaultCheck = False Then
				.vspdData.Row = Row
				.vspdData.Col = Col	
				.vspdData.Action = 0		'Active Cell
				.vspdData.focus
    			Exit Sub
			End If
			'/* 8�� ������ġ : ���� �������忡 �ʼ��Է� �ʵ� üũ - END */
			
			'/* 9�� ������ġ: '�ٸ� �۾��� �̷������ ��Ȳ���� �ٸ� �� �̵� �� ��ȸ�� �̷�� ���� �ʵ��� �Ѵ�. - START */
			If CheckRunningBizProcess = True Then
				.vspdData.Row = Row
				.vspdData.Col = Col	
				.vspdData.Action = 0
				Exit Sub
			End If
			'/* 9�� ������ġ: '�ٸ� �۾��� �̷������ ��Ȳ���� �ٸ� �� �̵� �� ��ȸ�� �̷�� ���� �ʵ��� �Ѵ�. - END */
			
			.vspdData.Row = NewRow
			.vspdData.Col = NewCol	
			.vspdData.Action = 0
			
			'Bottom�κп� Data�ѷ��ֱ� 
			.vspdData.Col = C_InspItemCd
			.hInspItemCd.value = .vspdData.Text
			.vspdData.Col = C_InspSeries
			.hInspSeries.value = .vspdData.Text
			
			lgIntFlgModeM = Parent.OPMD_CMODE
		
			With frm1.vspdData2
				.ReDraw = False
				.BlockMode = True
				.Row = 1
				.Row2 = .MaxRows
				.RowHidden = True
				.BlockMode = False
				.ReDraw = True
			End With
			
			'/* ��ü ���� ���� - START */
			If lgIntFlgMode = Parent.OPMD_UMODE Then
				Call SetToolBar("11111101001111")
			Else
				Call SetToolBar("11101101001111")
			End If
				
			Call DbQuery2(NewRow, False)
			'/* ��ü ���� ���� - END */
			
		End With
	End If
End Sub

'/* 9�� ������ġ: �Է��� ��ġ�� ��ȿ���� üũ�� Change Event������ �ƴ϶� EditChange���� �ϵ��� �Ѵ� - START */
'======================================================================================================
'   Event Name :vspddata2_EditChange
'   Event Desc :
'=======================================================================================================
Sub vspdData2_EditChange(ByVal Col, ByVal Row )
	With frm1
		If Col = C_DefectQty2 Then
			.vspdData2.Row = Row
			.vspdData2.Col = C_DefectQty2
			.vspdData.Row = .vspdData.ActiveRow
			.vspdData.Col = C_DefectQty1

			If UNICDbl(.vspdData2.Text) > UNICDbl(.vspdData.Text) Then
				'�ҷ������� �ҷ����� �˻��׸� �ҷ��� ���� Ŭ �� �����ϴ�.
				.vspdData2.Text = UNIFormatNumber(.vspdData2.TypeFloatMin, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
				Call DisplayMsgBox("221620", "X", "X", "X")
				.vspdData2.Row = Row
				.vspdData2.Col = C_DefectQty2
				Call SetFocusToDocument("M")
				.vspdData2.Focus
			End If
		End If
   	End With
End Sub
'/* 9�� ������ġ: �Է��� ��ġ�� ��ȿ���� üũ�� Change Event������ �ƴ϶� EditChange���� �ϵ��� �Ѵ� - END */

'=======================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)
	Dim strMark

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row
	With frm1.vspdData2
		.Row = Row
		.Col = 0
		strMark = .Text
		.Col = C_Flag 
		.Text = strMark
		
		If Col = C_DefectQty2 Then
			.Row = Row
			.Col = Col
			If UNICDbl(.text) < .TypeFloatMin Then
				.Text = UNIFormatNumber(.TypeFloatMin, ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
			End If
		End If
	End With
End Sub	

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '/* 9�� ������ġ: �ػ󵵿� ������� �������ǵ��� ���� - START */
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '��: ������ üũ 
    '/* 9�� ������ġ: �ػ󵵿� ������� �������ǵ��� ���� - END */
    	If lgStrPrevKey1 <> "" And lgStrPrevKey2 <> ""  Then            '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
      	    If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
		
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If      	    
    	End If
    End If
End Sub

'======================================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim lRow
    Dim lConvRow	'2003-03-01 Release �߰� 
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    With frm1
		lRow = .vspdData.ActiveRow
		.vspdData.Row = lRow
		.vspdData.Col = .vspdData.MaxCols
		lConvRow = .vspdData.Text
				
    	If lConvRow = 0 Then
    		lConvRow = 1
    	End If
    	
    	'/* 9�� ������ġ: �ػ󵵿� ������� ������ �ǵ��� ���� - START */
		If ShowDataLastRow < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then	        '��: ������ üũ 
		'/* 9�� ������ġ: �ػ󵵿� ������� ������ �ǵ��� ���� - END */
			If lgStrPrevKeyM(lConvRow - 1) <> "" Then            '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If CheckRunningBizProcess = True Then
					Exit Sub
				End If	
				Call DisableToolBar(Parent.TBC_QUERY)
				
				If DbQuery2(lRow, True) = False Then
					Call RestoreToolBar()
					Exit Sub
				End If				
			End If
		End if
    End With
End Sub

'==========================================================================================
'   Event Name : vspdData2_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	ggoSpread.Source = frm1.vspdData2	
	With frm1.vspdData2 
		If Row > 0 And Col = C_DefectTypePopUp Then
    		.Col = C_DefectTypeCd
			.Row = Row
    		Call OpenDefectType(.Text)
    	End If
    End With
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
    FncQuery = False 
    
    Dim IntRetCD 
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables															'Initializes local global variables

	ggoSpread.Source = frm1.vspdData	
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'This function check indispensable field
       Exit Function
    End If
  
    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False then
		Exit Function
	End If																	'��: Query db data
           
    FncQuery = True
End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function FncNew() 
    FncNew = False
    
    Dim IntRetCD 
    
	'-----------------------
    'Check previous data area
    '----------------------- 
    If ChangeCheck = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")               
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
  '-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
    Call InitVariables                                                      'Initializes local global variables
    
	ggoSpread.Source = frm1.vspdData	
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    
    Call SetDefaultVal
    '/* ��ü ���� ���� - START */
    Call SetToolBar("11100000000001")	
    '/* ��ü ���� ���� - END */
    If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantCd.focus 
	Else
		frm1.txtInspReqNo.focus 
	End If

    FncNew = True
End Function

'=======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncDelete() 
    '/* ��ü ���� ���� - START */
    FncDelete = False 
    Dim IntRetCD
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		            'Will you destory previous data"
	If IntRetCD = vbNo Then	Exit Function
    
    '-----------------------
    'Delete function call area
    '-----------------------
	If DbDelete = False Then Exit Function											'��: Delete db data
    
    FncDelete = True
    '/* ��ü ���� ���� - END */
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncSave() 
    FncSave = False
    
    Dim IntRetCD 
    Dim iRows
    Dim iParentRowNo

    '-----------------------
    'Precheck area
    '-----------------------
    If ChangeCheck = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                           
        Exit Function
    End If
    
    
    '8�� ������ġ: ȭ�鿡 ���̴� ���� �������忡 ���߰� �Ǿ����� Hidden �������忡 �ݿ��� �ȵ� �� üũ START
    If DefaultCheck = False Then
    	Exit Function
    End If
    '8�� ������ġ: ȭ�鿡 ���̴� ���� �������忡 ���߰� �Ǿ����� Hidden �������忡 �ݿ��� �ȵ� �� üũ END
    
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then
    	Exit Function
    End If

    '-----------------------
    'Check DefectQty
    '-----------------------
    
    '/* ���� START */
    Dim strDefectType 
	Dim strFlag
				
	With frm1.vspdData2
		For iRows = 1 To .MaxRows
			.Row = iRows
			.Col = C_DefectQty2
			If UNICDbl(.text) = 0 Then
				Call DisplayMsgBox("225001", "X", "X", "X")
				.Col = C_DefectTypeCd
				strDefectType = .Text
				.Col = C_Flag
				strFlag = .Text
				.Col = C_ParentRowNo
				iParentRowNo = CLng(.Text)
				Call ParentGetFocusCell(iParentRowNo, strDefectType, strFlag)
				Exit Function
			End If
		Next
	End With
	'/* ���� END */
    	
    '-----------------------
    'Save function call area
    '-----------------------
	If DbSave = False then	
		Exit Function
	End If				                                             '��: Save db data
    
    FncSave = True                                                       
End Function

'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy() 
	FncCopy = False		
	
	Dim IntRetCD
	Dim lRow
	Dim lConvRow
	Dim lRow2
	
	With frm1
		'Check Spread2 Data Exists for the keys
		If CheckDataExist = False Then
			Exit function
		End If
    	
		.vspdData2.ReDraw = False
		
		ggoSpread.Source = frm1.vspdData2	
		ggoSpread.CopyRow

		lRow = .vspdData.ActiveRow
		.vspdData.Row = lRow
		.vspdData.Col = .vspdData.MaxCols
		lConvRow = .vspdData.Text

		
		lRow2 = .vspdData2.ActiveRow
		
		.vspdData2.Col = C_DefectTypeCd
		.vspdData2.Text = ""
		
		.vspdData2.Col = C_DefectTypeNm
		.vspdData2.Text = ""
		
		.vspdData2.Col = C_ParentRowNo
		.vspdData2.Text = lConvRow
				
		.vspdData2.Col = C_Flag
		.vspdData2.Text = ggoSpread.InsertFlag
		
		Call SetSpread2ColorRange(lRow2, lRow2, "I")
    		
		'�������� ���� �ش� Ű�� ���� Client�� Data Row���� ������ 
		lglngHiddenRows(lConvRow - 1) = lglngHiddenRows(lConvRow - 1) + 1
		.vspdData2.ReDraw = True
		.vspdData2.focus
	End With

	Call SetActiveCell(frm1.vspdData2,C_DefectTypeCd,frm1.vspdData2.ActiveRow,"M","X","X")
	Set gActiveElement = document.ActiveElement		
	FncCopy = True
End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
	FncCancel = false
	Dim InsertDataCnt
	Dim lRow
	Dim i
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim iActiveRow
	Dim iConvActiveRow

	iActiveRow = frm1.vspdData.ActiveRow
	frm1.vspdData.Row = iActiveRow
	frm1.vspdData.Col = frm1.vspdData.MaxCols
	iConvActiveRow = frm1.vspdData.Text
	
	If frm1.vspdData.MaxRows < 1 then
		Exit function
	End if
	'Check Spread2 Data Exists for the keys
	If CheckDataExist = False Then
    	Exit function
    End If
    	
	ggoSpread.Source = frm1.vspdData2	
	With frm1.vspdData2
		'������ ������ �ʴ� ������ �Ѿ�� ��쿡 ���� ó�� - START	    
	    lngRangeFrom = .SelBlockRow
	    .Row = lngRangeFrom
		If .RowHidden = True Then
			lngRangeFrom = ShowDataFirstRow()
		End If		
		
		lngRangeTo = .SelBlockRow2
		.Row = lngRangeTo
		If .RowHidden = True Then
			lngRangeTo = ShowDataLastRow()
		End If		
		
		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Action = 2			'Select Block	SS_ACTION_SELECT_BLOCK
		.BlockMode = False
		'������ ������ �ʴ� ������ �Ѿ�� ��쿡 ���� ó�� - END
		
		InsertDataCnt = 0
		For i = lngRangeFrom To lngRangeTo
			.Row = i
			.Col = 0 
			If .Text = ggoSpread.InsertFlag Then
				InsertDataCnt = InsertDataCnt + 1
			End If
		Next
	
		.Redraw = False
		ggoSpread.EditUndo
		
		'�ٽ� ����� �κ��� Sequencial�ϰ� �ο�.- START
		lngRangeFrom = ShowDataFirstRow()
		lngRangeTo = ShowDataLastRow()
		
		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Col = C_Flag
		.Col2 = C_Flag
		.DestCol = 0
		.DestRow = lngRangeFrom
		If lngRangeFrom <> 0 Then
			.Action = 19	'SS_ACTION_COPY_RANGE	'@@@���� 
		End If
		.BlockMode = False
		'�ٽ� ����� �κ��� Sequencial�ϰ� �ο�.- END
		.Redraw = True
	End With
	
	lRow = frm1.vspdData.ActiveRow
	lglngHiddenRows(lRow - 1) = lglngHiddenRows(lRow - 1) - InsertDataCnt

	FncCancel = true
End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)		
	FncInsertRow = false	
	
	Dim lRow
	Dim lRow2
	Dim lconvRow
	Dim strMark
	Dim iInsertRow
	Dim imRow
	
	With frm1
		If .vspdData.MaxRows <= 0 Then
			Exit Function
		End If
		
		.vspdData2.ReDraw = False

		If IsNumeric(Trim(pvRowCnt)) Then
			imRow = CInt(pvRowCnt)
		Else
			imRow = AskSpdSheetAddRowCount()
			If imRow = "" Then
				Exit Function
			End If
		End If

		'Insert Row in Spread2
		.vspdData2.focus		
		ggoSpread.Source = .vspdData2
		ggoSpread.InsertRow .vspdData2.ActiveRow, imRow

		.vspdData.Row = .vspdData.ActiveRow
		.vspdData.Col = .vspdData.MaxCols
		lconvRow = CInt(.vspdData.Text)
		
		For iInsertRow = 0 To imRow - 1			
			lRow2 = .vspdData2.ActiveRow + iInsertRow

			.vspdData2.Row = lRow2
			.vspdData2.Col = 0
			strMark = .vspdData2.Text

			.vspdData2.Col = C_Flag 
			.vspdData2.Text = strMark
		
			.vspdData2.Col = C_ParentRowNo
			.vspdData2.Text = lconvRow
		
			'�������� ���� �ش� Ű�� ���� Client�� Data Row���� ������ 
			lglngHiddenRows(lconvRow - 1) = CInt(lglngHiddenRows(lconvRow - 1)) + 1
			Call SetSpread2ColorRange(lRow2, lRow2, "I")
		Next
		
		'/* ���� : ����� �� �ѹ��� ���� �߰� START */
		Dim i 
		Dim lngRangeFrom
		Dim lngRangeTo
		Dim strFlag
		Dim k
		
		lngRangeFrom = ShowDataFirstRow()
		lngRangeTo = ShowDataLastRow()
		k = 0
		
		for i = lngRangeFrom To lngRangeTo
			k = k + 1
			.vspdData2.Row = i
			.vspdData2.Col = 0
			strFlag = .vspdData2.Text
			If strFlag <> ggoSpread.InsertFlag and strFlag <> ggoSpread.UpdateFlag and strFlag <> ggoSpread.DeleteFlag then
				.vspdData2.Text = CStr(k)
			End If
		Next
		
		'/* ���� END */
		.vspdData2.focus
		.vspdData2.ReDraw = True
		
	End With
	
	Call SetActiveCell(frm1.vspdData2,C_DefectTypeCd,frm1.vspdData2.ActiveRow,"M","X","X")
	Set gActiveElement = document.ActiveElement		
	FncInsertRow = true
End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow()		
	FncDeleteRow = false
	
	Dim lDelRows
	Dim iDelRowCnt, i
    Dim lngRangeFrom
    Dim lngRangeTo
    
	If frm1.vspdData.MaxRows < 1 then
		Exit function
	End if
		
	'Check Spread2 Data Exists for the keys
	If CheckDataExist = False Then
		Exit function
	End If
		
	With frm1.vspdData2
		.Redraw = False
		
		.Focus
	
		'������ ������ �ʴ� ������ �Ѿ�� ��쿡 ���� ó�� - START	    
	    lngRangeFrom = .SelBlockRow
	    .Row = lngRangeFrom
		If .RowHidden = True Then
			lngRangeFrom = ShowDataFirstRow()
		End If		
		
		lngRangeTo = .SelBlockRow2
		.Row = lngRangeTo
		If .RowHidden = True Then
			lngRangeTo = ShowDataLastRow()
		End If		
		
		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Action = 2			'Select Block	SS_ACTION_SELECT_BLOCK
		.BlockMode = False
		'������ ������ �ʴ� ������ �Ѿ�� ��쿡 ���� ó�� - END
		
	    ggoSpread.Source = frm1.vspdData2 
	     '----------  Coding part  -------------------------------------------------------------   
		lDelRows = ggoSpread.DeleteRow
		
		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Col = 0
		.Col2 = 0
		.DestCol = C_Flag
		.DestRow = .SelBlockRow
		.Action = 19	'SS_ACTION_COPY_RANGE
		.BlockMode = False
	
		.Redraw = True
	End With
	
	FncDeleteRow = true
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	FncExcel = False
 	Call parent.FncExport(Parent.C_MULTI)		
 	FncExcel = True
 End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint() 
	FncPrint = False
    Call parent.FncPrint()
    FncPrint = True
End Function

'=======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================================
Function FncPrev() 
	FncPrev = False                                           
End Function

'=======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================================
Function FncNext() 
    FncNext = False                                              
End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : ȭ�� �Ӽ�, Tab���� 
'========================================================================================================
Function FncFind()
	FncFind = False 
    Call parent.FncFind(Parent.C_SINGLEMULTI , True)  
	FncFind = True                        
End Function
 
'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub 
 
'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
	FncExit = False
	
	Dim IntRetCD
	
	If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
End Function

'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================================
Function DbDelete() 
	'/* ��ü ���� ���� - START */
	DbDelete = False                                                             
	
	Dim strVal
	
	Call LayerShowHide(1)
	
	with frm1
	
		strVal = BIZ_PGM_DEL_ID & "?txtInspReqNo=" & Trim(.hInspReqNo.value) _
								& "&txtPlantCd=" & Trim(.hPlantCd.value) _
		
	End With

	Call RunMyBizASP(MyBizASP, strVal)													'��: �����Ͻ� ASP �� ���� 

	DbDelete = True
	'/* ��ü ���� ���� - END */
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================================
Function DbDeleteOk()	
	'/* ��ü ���� ���� - START */
	frm1.txtPlantCd.value = frm1.hPlantCd.value
	frm1.txtInspReqNo.value  = frm1.hInspReqNo.value
	
	Call MainQuery()
	'/* ��ü ���� ���� - END */
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
	DbQuery = False
	
	Dim strVal
	                                                             
	Call LayerShowHide(1)
	
	with frm1
	
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001					'��: 
			strVal = strVal & "&txtinspReqNo=" & Trim(.hInspReqNo.value)				'��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001					'��: 
			strVal = strVal & "&txtinspReqNo=" & Trim(.txtInspReqNo.value)				'��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				
			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)										    '��: �����Ͻ� ASP �� ���� 
	
	DbQuery = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryOk()
	DbQueryOk = False 
	Dim i
	Dim lRow
	
	Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
	
	With frm1
		lRow = .vspdData.MaxRows
		If lRow > 0 Then
			ReDim lgStrPrevKeyM(lRow - 1)	
			ReDim lglngHiddenRows(lRow - 1)
			For i = 0 To lRow - 1
				lglngHiddenRows(i) = 0
			Next
			.vspdData.Row = 1
			
			.vspdData.Col = C_InspItemCd
			'.vspdData.Action = 0 'Active Cell
			.hInspItemCd.Value = .vspdData.Text 
			.vspdData.Col = C_InspSeries
			.hInspSeries.Value = .vspdData.Text 
			
			lgIntFlgModeM = Parent.OPMD_CMODE
			
			'/* ��ü ���� ���� - START */
			Call SetToolBar("11101101001111")
			If DbQuery2(1, False) = False then
				Exit Function
			End If					
			'/* ��ü ���� ���� - END */
		End If
		lgBlnFlgChgValue = False
	End With
	DbQueryOk = True
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    DbSave = False
    
    Dim lRow  
    Dim lRow2      
	Dim lGrpCnt     
	Dim lValCnt
	Dim iSpdCount     
	Dim strVal 
	Dim strDel
	Dim iParentNum
	Dim iTargetParentNum
	Dim strHTML

	Dim strInspItemCd
	Dim strInspSeries
	Dim strDefectTypeCd
	Dim strDefectQty2
	
	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		.txtFlgMode.value = lgIntFlgModeM
	End With	    
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	lGrpCnt = 1
	strVal = ""
    strDel = ""
	strHTML = ""
	iSpdCount = 0   
    '-----------------------
	'Data manipulate area
	'-----------------------
	ggoSpread.source = frm1.vspdData2	
	With frm1
		For lRow = 1 To .vspdData2.MaxRows
			.vspdData2.Row = lRow
			.vspdData2.Col = C_ParentRowNo
			
			iParentNum = CInt(.vspdData2.Text)
			For lRow2 = 1 To .vspdData.MaxRows
				.vspdData.Row = lRow2
				.vspdData.Col = .vspdData.MaxCols
				If iParentNum = CInt(.vspdData.Text) Then
					iTargetParentNum = lRow2
					Exit For
				End IF
			Next
			
			'.vspdData.Row = CLng(.vspdData2.Text)
			.vspdData.Row = CInt(lRow2)
			.vspdData2.Col = 0
			
			Select Case .vspdData2.Text
				Case ggoSpread.InsertFlag				'��: �ű� 
					.vspdData.Col = C_InspItemCd
					strInspItemCd = Trim(.vspdData.Text)
					.vspdData.Col = C_InspSeries
					strInspSeries = .vspdData.Text
					.vspdData2.Col = C_DefectTypeCd
					strDefectTypeCd = Trim(.vspdData2.Text)
					.vspdData2.Col = C_DefectQty2
					strDefectQty2 = .vspdData2.Text
					
					
					strVal = strVal & "C" & Parent.gColSep & _
							 strInspItemCd & Parent.gColSep & _
							 strInspSeries & Parent.gColSep & _
							 strDefectTypeCd & Parent.gColSep & _
							 strDefectQty2 & Parent.gColSep & _
						     CStr(lRow) & Parent.gRowSep
					lGrpCnt = lGrpCnt + 1
					lValCnt = lValCnt + 1
				Case ggoSpread.UpdateFlag				'��: ���� 
					.vspdData.Col = C_InspItemCd
					strInspItemCd = Trim(.vspdData.Text)
					.vspdData.Col = C_InspSeries
					strInspSeries = .vspdData.Text
					.vspdData2.Col = C_DefectTypeCd
					strDefectTypeCd = Trim(.vspdData2.Text)
					.vspdData2.Col = C_DefectQty2
					strDefectQty2 = .vspdData2.Text
					
					strVal = strVal & "U" & Parent.gColSep & _
							 strInspItemCd & Parent.gColSep & _
							 strInspSeries & Parent.gColSep & _
							 strDefectTypeCd & Parent.gColSep & _
							 strDefectQty2 & Parent.gColSep & _
						     CStr(lRow) & Parent.gRowSep
					lGrpCnt = lGrpCnt + 1
					lValCnt = lValCnt + 1
				Case ggoSpread.DeleteFlag				'��: ���� 
					strDel = strDel & "D" & Parent.gColSep			'��: D=Delete
					.vspdData.Col = C_InspItemCd
					strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
					.vspdData.Col = C_InspSeries
					strDel = strDel & .vspdData.Text & Parent.gColSep
					.vspdData2.Col = C_DefectTypeCd			'3
					strDel = strDel & Trim(.vspdData2.Text) & Parent.gColSep
					strDel = strDel & CStr(lRow) & Parent.gRowSep	'4
					lGrpCnt = lGrpCnt + 1
			End Select
			
			.vspdData2.Col = 0
			If (lValCnt Mod 1000 = 0) AND (ggoSpread.DeleteFlag <> .vspdData2.Text) Then
				iSpdCount = iSpdCount + 1
				strHTML = strHTML & "<TEXTAREA CLASS=hidden Name=txtSpread" & iSpdCount & " Width=100% tag=""24"" TABINDEX=""-1"">" & strVal & "</TEXTAREA>"
				strVal = ""
			End If
		Next
		iSpdCount = iSpdCount + 1
		strHTML = strHTML & "<TEXTAREA CLASS=hidden Name=txtSpread" & iSpdCount & " Width=100% tag=""24"" TABINDEX=""-1"">" & strVal & "</TEXTAREA>"
	End With
			
	divTextArea.innerHTML = strHTML
	frm1.SpdCount.value = iSpdCount
	frm1.txtMaxRows.value = lGrpCnt-1
	frm1.txtSpread.value = strDel
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)			'��: �����Ͻ� ASP �� ���� 
	
	DbSave = True                                                  
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	Call InitVariables
	frm1.vspdData2.MaxRows = 0
	Call MainQuery()
End Function

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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0">
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>���԰˻� �ҷ�����</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></TD>
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
				        					<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="����" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
												<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14"></TD>								
				        					<TD CLASS="TD5" NOWRAP>�˻��Ƿڹ�ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspReqNo" SIZE=20  MAXLENGTH=18 ALT="�˻��Ƿڹ�ȣ" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspReqNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspReqNo()"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>		
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>ǰ��</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=20 ALT="ǰ��" tag="14">
								<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
								<TD CLASS="TD5" NOWRAP>����ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 ALT="����ó" tag="14">
								<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>��Ʈ��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotNo" SIZE=15 MAXLENGTH=12 ALT="LOT NO" tag="14">
									<INPUT TYPE=TEXT NAME="txtLotSubNo" SIZE=10 MAXLENGTH=5 tag="14" STYLE="Text-Align: Right"></TD>
								<TD CLASS="TD5" NOWRAP>��Ʈũ��</TD>            
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/q2113ma1_fpDoubleSingle1_txtLotSize.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% Colspan=4>
									<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
										<TR HEIGHT="*">
											<TD WIDTH="52%">
												<script language =javascript src='./js/q2113ma1_I213201183_vspdData.js'></script>
											</TD>
											<TD WIDTH=10>&nbsp;</TD>
											<TD WIDTH="48%">
												<script language =javascript src='./js/q2113ma1_I796115529_vspdData2.js'></script>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
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
    				<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadInspection">���԰˻�</A>&nbsp;|&nbsp;<A href="vbscript:LoadDecision">����</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
    			</TR>
    		</TABLE>
    	</TD>
    </TR>
    <TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP"  tabindex=-1 SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noreSIZE framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>	<!--'2003-03-01 Release �߰� -->
<INPUT TYPE=HIDDEN NAME="SpdCount" tag="24" TABINDEX="-1">	<!--'2003-03-01 Release �߰� -->
<TEXTAREA class=hidden name=txtSpread Width=100% tag="24" tabindex=-1 ></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtDate" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hInspReqNo" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hInspItemCd" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hInspSeries" tag="24" tabindex=-1 >
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" tabindex=-1 >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

