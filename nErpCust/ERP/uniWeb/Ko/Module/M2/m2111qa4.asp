<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name		  : Procurement
'*  2. Function Name		:
'*  3. Program ID		   : M2111QA4
'*  4. Program Name		 : ���ſ�û����ȸ 
'*  5. Program Desc		 : ���ſ�û����ȸ 
'*  6. Component List	   :
'*  7. Modified date(First) : 2001/01/08
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)	 : Min, Hak-jun
'* 10. Modifier (Last)	  : KANG SU HWAN
'* 11. Comment			  :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*							this mark(��) Means that "may  change"
'*							this mark(��) Means that "must change"
'* 13. History			  :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
Option Explicit

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'*********************************************************************************************************
Const BIZ_PGM_ID 		= "m2111qb4.asp"
Const BIZ_PGM_ID2 		= "m2111mb1_1.asp"
Const BIZ_PGM_JUMP_ID 	= "m2111ma1"
Const C_MaxKey		  = 21
Const C_SHEETMAXROWS2  = 20
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= %>
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= %>
'----------------  ���� Global ������ ����  ----------------------------------------------------------- %>
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ %>
Dim lgIsOpenPop
Dim lgSaveRow
Dim IscookieSplit

Dim DBQueryCheck

Dim C_SpplCd
Dim C_SpplNm
Dim C_QuotaRate
Dim C_ApportionQty
Dim C_PlanDt
Dim C_GrpCd
Dim C_GrpNm
Dim lgPageNo2

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'			 		 	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����)
'#########################################################################################################
Sub InitVariables()
	lgStrPrevKey	 = ""
	lgBlnFlgChgValue = False
	lgSortKey		= 1
	lgSaveRow		= 0
	lgPageNo		 = ""
	lgPageNo2		 = ""
	DBQueryCheck	 = True
	lgIntFlgMode = Parent.OPMD_CMODE
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim StartDate, EndDate, EndDate1

	StartDate   = uniDateAdd("m", -1, "<%=GetSvrDate%>", parent.gServerDateFormat)
	StartDate   = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
	EndDate	 = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
	EndDate1	= uniDateAdd("m", +1, "<%=GetSvrDate%>", parent.gServerDateFormat)
	EndDate1   = UniConvDateAToB(EndDate1, parent.gServerDateFormat, parent.gDateFormat)

	With frm1
 		.txtPrFrDt.Text	= StartDate
 		.txtPrToDt.Text	= EndDate
 		.txtPdFrDt.Text	= StartDate
 		.txtPdToDt.Text	= EndDate1
		.txtPlantCd.value= parent.gPlant
		.txtPlantNm.value= parent.gPlantNm

		.txtPlantCd.focus
	End With
	Set gActiveElement = document.activeElement

End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
<%'========================================================================================%>
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M2111QA4","S","A","V20041210", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A")

	Call InitSpreadSheet2
End Sub

Sub InitSpreadSheet2()
	Call InitSpreadPosVariables()

	With frm1
		.vspdData2.ReDraw = false

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021103",, parent.gAllowDragDropSpread

	   .vspdData2.MaxCols = C_GrpNm+1
	   .vspdData2.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit   C_SpplCd, "����ó", 15,,,15,2
		ggoSpread.SSSetEdit	  C_SpplNm, "����ó��", 20
		SetSpreadFloatLocal	  C_QuotaRate, "��к���(%)",15,1,5
		SetSpreadFloatLocal   C_ApportionQty, "��η�", 15, 1,3
		ggoSpread.SSSetDate	  C_PlanDt, "���ֿ�����", 15,2,gDateFormat
		ggoSpread.SSSetEdit	  C_GrpCd, "���ű׷�", 10,,,10,2
		ggoSpread.SSSetEdit   C_GrpNm, "���ű׷��", 20

		Call ggoSpread.MakePairsColumn(C_SpplCd,C_SpplNm)
		Call ggoSpread.MakePairsColumn(C_GrpCd,C_GrpNm)
		Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols,	.vspdData2.MaxCols,	True)

		.vspdData2.ReDraw = True
	End With

	Call SetSpreadLock("B")
End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pOpt)
	If pOpt = "A" Then
	  ggoSpread.Source = frm1.vspdData
	  ggoSpread.SpreadLockWithOddEvenRowColor()
	ElseIF pOpt = "B" Then
	  ggoSpread.Source = frm1.vspdData2
	  ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_SpplCd			=	1
	C_SpplNm			=	2
	C_QuotaRate			=	3
	C_ApportionQty		=	4
	C_PlanDt			=	5
	C_GrpCd				=	6
	C_GrpNm				=	7
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   :
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
	   Case "A"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			 	C_SpplCd			=	iCurColumnPos(1)
				C_SpplNm			=	iCurColumnPos(2)
				C_QuotaRate			=	iCurColumnPos(3)
				C_ApportionQty		=	iCurColumnPos(4)
				C_PlanDt			=	iCurColumnPos(5)
				C_GrpCd				=	iCurColumnPos(6)
				C_GrpNm				=	iCurColumnPos(7)

	End Select
End Sub

Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
		Case 2															  '�ݾ� 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign
		Case 3															  '���� 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo	   ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"P"
		Case 4															  '�ܰ� 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
		Case 5															  'ȯ�� 
			ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
	End Select
End Sub

'------------------------------------------  OpenPlantCd()  --------------------------------------------
'	Name : OpenPlantCd()
'	Description : Plant PopUp
'-------------------------------------------------------------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����"
	arrParam(1) = "B_Plant"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)
	arrParam(4) = ""
	arrParam(5) = "����"

	arrField(0) = "Plant_Cd"
	arrField(1) = "Plant_NM"

	arrHeader(0) = "����"
	arrHeader(1) = "�����"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)

	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
	End If
End Function
'------------------------------------------  OpenPrStsCd()  -------------------------------------------------
'	Name : OpenPrStsCd()
'	Description : PrStatus PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPrStsCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "��û�������"
	arrParam(1) = "B_MINOR"
	arrParam(2) = Trim(frm1.txtPrStsCd.Value)
'	arrParam(3) = Trim(frm1.txtPrStsNm.Value)
	arrParam(4) = "MAJOR_CD = " & FilterVar("M2101", "''", "S") & ""
	arrParam(5) = "��û�������"

	arrField(0) = "MINOR_CD"
	arrField(1) = "MINOR_NM"

	arrHeader(0) = "��û�������"
	arrHeader(1) = "��û������¸�"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPrStsCd.focus
		Exit Function
	Else
		frm1.txtPrStsCd.Value = arrRet(0)
		frm1.txtPrStsNm.Value = arrRet(1)
		frm1.txtPrStsCd.focus
		Set gActiveElement = document.activeElement
	End If
End Function
'------------------------------------------  OpenRqDeptCd()  -------------------------------------------
'	Name : OpenRqDeptCd()
'	Description : Req Dept PopUp
'--------------------------------------------------------------------------------------------------------
Function OpenRqDeptCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "��û�μ�"
	arrParam(1) = "B_ACCT_DEPT"
	arrParam(2) = Trim(frm1.txtRqDeptCd.Value)
'	arrParam(3) = Trim(frm1.txtRqDeptNm.Value)
	arrParam(4) = "ORG_CHANGE_ID= " & FilterVar(parent.gChangeOrgId, "''", "S") & " "
	arrParam(5) = "��û�μ�"

	arrField(0) = "DEPT_CD"
	arrField(1) = "DEPT_NM"

	arrHeader(0) = "��û�μ�"
	arrHeader(1) = "��û�μ���"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtRqDeptCd.focus
		Exit Function
	Else
		frm1.txtRqDeptCd.Value = arrRet(0)
		frm1.txtRqDeptNm.Value = arrRet(1)
		frm1.txtRqDeptCd.focus
		Set gActiveElement = document.activeElement
	End If
End Function

'------------------------------------------  OpenPrType()  -------------------------------------------------
'	Name : OpenPrType()
'	Description : PR Type PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPrType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "���ſ�û����"
	arrParam(1) = "B_MINOR"
	arrParam(2) = Trim(frm1.txtPrType.Value)
'	arrParam(3) = Trim(frm1.txtPrTypeNm.Value)
	arrParam(4) = "MAJOR_CD = " & FilterVar("M2102", "''", "S") & " "
	arrParam(5) = "���ſ�û����"

	arrField(0) = "MINOR_CD"
	arrField(1) = "MINOR_NM"

	arrHeader(0) = "���ſ�û����"
	arrHeader(1) = "���ſ�û���и�"

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPrType.focus
		Exit Function
	Else
		frm1.txtPrType.Value = arrRet(0)
		frm1.txtPrTypeNm.Value = arrRet(1)
		frm1.txtPrType.focus
		Set gActiveElement = document.activeElement
	End If
End Function

Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim IntRetCD
	Dim iCalledAspName

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = ""	'�ֹ�ó 
	arrParam(1) = ""	'�����׷� 
	arrParam(2) = Trim(frm1.txtPlantCd.value)	'���� 
	arrParam(3) = ""	'��ǰ�� 
	arrParam(4) = ""	'���ֹ�ȣ 
	arrParam(5) = ""	'�߰� Where�� 

'	arrRet = window.showModalDialog("../s3/s3135pa1.asp", Array(arrParam), _
'			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 	iCalledAspName = AskPRAspName("S3135PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3135PA1", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet = "" Then
		frm1.txtTrackNo.focus
		Exit Function
	Else
		frm1.txtTrackNo.Value = Trim(arrRet)
		frm1.txtTrackNo.focus
	End If
End Function


'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
'	Name : PopZAdoConfigGrid()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	Call OpenSortPopup("A")
End Sub
'========================================================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenSortPopup Reference Popup
'========================================================================================================
Function OpenSortPopup(ByVal pSpdNo)
	Dim arrRet

	On Error Resume Next

	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
	   Call InitVariables
	   Call InitSpreadSheet()
   End If
End Function

'==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'====================================================================================================
Function CookiePage(ByVal Kubun)
	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877

	If Kubun = 1 Then

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		WriteCookie CookieSplit , IscookieSplit

		if frm1.vspdData.ActiveRow > 0 then
			strTemp = ReadCookie(CookieSplit)

			If strTemp = "" then Exit Function
			arrVal = Split(strTemp, parent.gRowSep)
			frm1.vspdData.Row = frm1.vspdData.ActiveRow

			WriteCookie "ReqNo" , arrVal(0)

			WriteCookie CookieSplit , ""

		end if
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		Dim iniSep

		If Len(ReadCookie ("PlantCd")) Then
			frm1.txtPlantCd.Value	=  ReadCookie ("PlantCd")
			WriteCookie "PlantCd",""
		Else
			frm1.txtPlantCd.Value	=  arrVal(0)
		End If

		frm1.txtPlantNm.value	=  arrVal(1)

		If Len(ReadCookie ("ItemCd")) Then
			frm1.txtItemCd.Value	=  ReadCookie ("ItemCd")
			WriteCookie "ItemCd",""
		Else
			frm1.txtItemCd.Value	=  arrVal(2)
		End If

		frm1.txtItemNm.Value	=  arrVal(3)

		If arrVal(5) = "" or arrVal(5) = Null Then
			frm1.txtPrFrDt.Text	=  ReadCookie ("PrFrDt")
			WriteCookie "PrFrDt",""
		Else
			frm1.txtPrFrDt.Text		=  arrVal(5)
		End If

		If arrVal(5) = "" or arrVal(5) = Null Then
			frm1.txtPrToDt.Text	=  ReadCookie ("PrToDt")
			WriteCookie "PrToDt",""
		Else
			frm1.txtPrToDt.Text		=  arrVal(5)
		End If

		If arrVal(6) = "" or arrVal(6) = Null Then
			frm1.txtPdFrDt.Text	=  ReadCookie ("PdFrDt")
			WriteCookie "PdFrDt",""
		Else
			frm1.txtPdFrDt.Text		=  arrVal(6)
		End If

		If arrVal(6) = "" or arrVal(6) = Null Then
			frm1.txtPdToDt.Text	=  ReadCookie ("PdToDt")
			WriteCookie "PdToDt",""
		Else
			frm1.txtPdToDt.Text		=  arrVal(6)
		End If

		If Len(ReadCookie ("PrStsCd")) Then
			frm1.txtPrStsCd.Value	=  ReadCookie ("PrStsCd")
			WriteCookie "PrStsCd",""
		Else
			frm1.txtPrStsCd.Value	=  arrVal(7)
		End If

		frm1.txtPrStsNm.Value	=  arrVal(8)

		If Len(ReadCookie ("RqDeptCd")) Then
			frm1.txtRqDeptCd.Value	=  ReadCookie ("RqDeptCd")
			WriteCookie "RqDeptCd",""
		Else
			frm1.txtRqDeptCd.Value	=  arrVal(9)
		End If

		frm1.txtRqDeptNm.Value	=  arrVal(10)

		If Len(ReadCookie ("PrType")) Then
			frm1.txtPrType.Value	=  ReadCookie ("PrType")
			WriteCookie "PrType",""
		Else
			frm1.txtPrType.Value	=  arrVal(11)
		End If

		frm1.txtPrTypeNm.value	=  arrVal(12)

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			WriteCookie "PlantCd",""
			WriteCookie "ItemCd",""
			WriteCookie "PrFrDt",""
			WriteCookie "PrToDt",""
			WriteCookie "PdFrDt",""
			WriteCookie "PdToDt",""
			WriteCookie "PrStsCd",""
			WriteCookie "RqDeptCd",""
			WriteCookie "PrType",""
			Exit Function
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""
	End IF
End Function


'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'		 ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'		 �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó�� 
'*********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
 Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
	Call SetToolbar("1100000000001111")
	Call CookiePage(0)

	Set gActiveElement = document.activeElement
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


'==========================================================================================
'   Event Name : txtPdFrDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
 Sub txtPdFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPdFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPdFrDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtPdToDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
 Sub txtPdToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPdToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPdToDt.Focus
	End If
End Sub

'==========================================================================================
'   Event Name : txtPrFrDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
 Sub txtPrFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPrFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPrFrDt.Focus
End If
End Sub
'==========================================================================================
'   Event Name : txtPrToDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
 Sub txtPrToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPrToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtPrToDt.Focus
	End If
End Sub
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   :
'========================================================================================
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   :
'========================================================================================
Sub PopRestoreSpreadColumnInf()

	ggoSpread.Source = gActiveSpdSheet

	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet2()
	Call ggoSpread.ReOrderingSpreadData()

End Sub
'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc :
'========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then

	  gMouseClickStatus = "SPCR"
   End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc :
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
Sub vspdData2_GotFocus()
	ggoSpread.Source = frm1.vspdData2
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc :
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
	ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
	Call GetSpreadColumnPos("A")
End Sub
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc :
'==========================================================================================
Sub txtPrFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtPrToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtPdFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtPdToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
 Sub vspdData_GotFocus()
	ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'==========================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		'	Call CookiePage(1)
		End If
	End If
End Function

'==========================================================================================
' Event Name : vspdData_LeaveCell
' Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If Row <> NewRow And NewRow > 0 Then
		Call vspdData_Click(NewCol, NewRow)
		frm1.vspdData2.MaxRows = 0
		Call Dbquery2(NewRow)
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"

	If frm1.vspdData.MaxRows = 0 Then Exit Sub

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If

		Exit Sub
	End If
	Call SetSpreadColumnValue("A",Frm1.vspdData, Col, Row)

	With frm1.vspdData
	.Row = Row
	IscookieSplit = ""
	Dim inti
	For inti=1 To C_MaxKey
		.Col = GetKeyPos("A",inti)
		IscookieSplit = IscookieSplit & Trim(.text) & parent.gRowSep
	Next

	End With

'	frm1.vspdData2.MaxRows = 0
'	Call Dbquery2()
End Sub

Sub vspdData2_Click(ByVal Col, ByVal Row)	'###�׸��� ������ ���Ǻκ�###
 	gMouseClickStatus = "SPC"

 	Set gActiveSpdSheet = frm1.vspdData2

	Call SetPopupMenuItemInf("0000111111")		 'ȭ�麰 ���� 

	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
 Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspddata,NewTop) Then
		If lgPageNo <> "" Then
 			If CheckRunningBizProcess = True Then
				Exit Sub
			End If
			Call DisableToolBar(parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If
End Sub


'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If

	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspddata2,NewTop) Then
		If lgPageNo2 <> "" Then
 			If CheckRunningBizProcess = True Then
				Exit Sub
			End If
			Call DisableToolBar(parent.TBC_QUERY)
			If DBQuery2(frm1.vspdData.ActiveRow) = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If
End Sub

'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'*********************************************************************************************************
Function FncQuery()
	FncQuery = False
	Err.Clear

	With frm1
		If CompareDateByFormat(.txtPrFrDt.text,.txtPrToDt.text,.txtPrFrDt.Alt,.txtPrToDt.Alt, _
				   "970025",.txtPrFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtPrFrDt.text) <> "" And Trim(.txtPrToDt.text) <> "" Then
			Call DisplayMsgBox("17a003", "X","���ſ�û��", "X")
			Exit Function
		End if
	End With

	With frm1
		If CompareDateByFormat(.txtPdFrDt.text,.txtPdToDt.text,.txtPdFrDt.Alt,.txtPdToDt.Alt, _
				   "970025",.txtPdFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtPdFrDt.text) <> "" And Trim(.txtPdToDt.text) <> "" Then
			Call DisplayMsgBox("17a003", "X","�ʿ䳳����", "X")
			Exit Function
		End if
	End With

	'-----------------------
	'Erase contents area
	'-----------------------
'	Call ggoOper.ClearField(Document, "2")
	ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2	'###�׸��� ������ ���Ǻκ�###
	ggoSpread.ClearSpreadData
	Call InitVariables

	'-----------------------
	'Check condition area
	'-----------------------
'	If Not chkField(Document, "1") Then
'	   Exit Function
'	End If
	DBQueryCheck = True
	'-----------------------
	'Query function call area
	'-----------------------

	If DbQuery = False Then Exit Function
	FncQuery = True
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
	Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
 Function FncFind()
	Call parent.FncFind(parent.C_MULTI , False)
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

	Err.Clear
	If LayerShowHide(1) = False Then Exit Function

	With frm1
	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.hdnPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.hdnItemCd.Value)
		strVal = strVal & "&txtPrFrDt=" & Trim(.hdnPrFrDt.Value)
		strVal = strVal & "&txtPrToDt=" & Trim(.hdnPrToDt.Value)
		strVal = strVal & "&txtPdFrDt=" & Trim(.hdnPdFrDt.Value)
		strVal = strVal & "&txtPdToDt=" & Trim(.hdnPdToDt.Value)
		strVal = strVal & "&txtPrStsCd=" & Trim(.hdnPrStsCd.value)
		strVal = strVal & "&txtRqDeptCd=" & Trim(.hdnRqDeptCd.value)
		strVal = strVal & "&txtPrTypeCd=" & Trim(.hdnPrTypeCd.value)
		strVal = strVal & "&txtTrackNo=" & Trim(.hdnTrackNo.Value)

		strVal = strVal & "&txtchangorgid=" & parent.gchangeorgid
	else
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)
		strVal = strVal & "&txtPrFrDt=" & Trim(.txtPrFrDt.Text)
		strVal = strVal & "&txtPrToDt=" & Trim(.txtPrToDt.Text)
		strVal = strVal & "&txtPdFrDt=" & Trim(.txtPdFrDt.Text)
		strVal = strVal & "&txtPdToDt=" & Trim(.txtPdToDt.Text)
		strVal = strVal & "&txtPrStsCd=" & Trim(.txtPrStsCd.value)
		strVal = strVal & "&txtRqDeptCd=" & Trim(.txtRqDeptCd.value)
		strVal = strVal & "&txtPrTypeCd=" & Trim(.txtPrType.value)
		strVal = strVal & "&txtTrackNo=" & Trim(.txtTrackNo.Value)
		strVal = strVal & "&txtchangorgid=" & parent.gchangeorgid
	end if

		strVal = strVal & "&lgPageNo="   & lgPageNo
		strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="	 & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		Call RunMyBizASP(MyBizASP, strVal)
	End With

	DbQuery = True
	Call SetToolbar("1100000000011111")
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
	'-----------------------
	'Reset variables area
	'-----------------------
	lgBlnFlgChgValue = False
	lgSaveRow		= 1
	lgIntFlgMode = parent.OPMD_UMODE

	if DBQueryCheck = True then
		Call DbQuery2(1)
	end if

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement

	DBQueryCheck = False
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
 Function DbQuery2(ByVal pRow)
	Dim strVal

	DbQuery2 = False

	Err.Clear
	If LayerShowHide(1) = False Then Exit Function
	If frm1.vspdData.MaxRows <= 0 Then Exit Function
	With frm1
		.vspdData.Row = pRow
		.vspdData.Col = 1
		strVal = BIZ_PGM_ID2 & "?txtPrno=" & Trim(.vspdData.text)
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		strVal = strVal & "&lgPageNo="		 & lgPageNo2						'��: Next key tag

		Call RunMyBizASP(MyBizASP, strVal)
	End With

	DbQuery2 = True
	Call SetToolbar("1100000000011111")
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk2()
	'-----------------------
	'Reset variables area
	'-----------------------
	lgBlnFlgChgValue = False
	lgSaveRow		= 1
	lgIntFlgMode = parent.OPMD_UMODE
End Function

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<!-- '#########################################################################################################
'	   					6. Tag�� 
'	���: Tag�κ� ���� 
	' �Է� �ʵ��� ��� MaxLength=? �� ��� 
	' CLASS="required" required  : �ش� Element�� Style �� Default Attribute
		' Normal Field�϶��� ������� ���� 
		' Required Field�϶��� required�� �߰��Ͻʽÿ�.
		' Protected Field�϶��� protected�� �߰��Ͻʽÿ�.
			' Protected Field�ϰ�� ReadOnly �� TabIndex=-1 �� ǥ���� 
	' Select Type�� ��쿡�� className�� ralargeCB�� ���� width="153", rqmiddleCB�� ���� width="90"
	' Text-Transform : uppercase  : ǥ�Ⱑ �빮�ڷ� �� �ؽ�Ʈ 
	' ���� �ʵ��� ��� 3���� Attribute ( DDecPoint DPointer DDataFormat ) �� ��� 
'######################################################################################################### -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ſ�û��</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right></td>
					<TD WIDTH=10>&nbsp;</TD>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����"  NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd() ">
														   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>

									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ǰ��"  NAME="txtItemCd" SIZE=18 LANG="ko" MAXLENGTH=18 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd() ">
														   <INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
								</TR>
									<TD CLASS="TD5" NOWRAP>���ſ�û��</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellpadding=0 cellspacing=0>
											<tr>
												<td NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtPrFrDt CLASSID=<%=gCLSIDFPDT%> tag="11X1" ALT="���ſ�û��"></OBJECT>');</SCRIPT>
												</td>
												<td NOWRAP>~</td>
												<td NOWRAP>
												   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtPrToDt CLASSID=<%=gCLSIDFPDT%> ALT="���ſ�û��" tag="11X1"></OBJECT>');</SCRIPT>
												</td>
											</tr>
										</table>
									</TD>
									<TD CLASS="TD5" NOWRAP>�ʿ䳳����</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellpadding=0 cellspacing=0>
											<tr>
												<td NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtPdFrDt CLASSID=<%=gCLSIDFPDT%> tag="11X1" ALT="�ʿ䳳����"></OBJECT>');</SCRIPT>
												</td>
												<td NOWRAP>~</td>
												<td NOWRAP>
												   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtPdToDt CLASSID=<%=gCLSIDFPDT%> ALT="�ʿ䳳����" tag="11X1"></OBJECT>');</SCRIPT>
												</td>
											</tr>
										</table>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��û�������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="��û�������" NAME="txtPrStsCd"  SIZE=10 MAXLENGTH=5 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrStsCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPrStsCd()">
														   <INPUT TYPE=TEXT NAME="txtPrStsNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>��û�μ�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="��û�μ�" NAME="txtRqDeptCd" SIZE=10 MAXLENGTH=10  MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRqDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRqDeptCd()">
														   <INPUT TYPE=TEXT NAME="txtRqDeptNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>���ſ�û����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="���ſ�û����" NAME="txtPrType" SIZE=10 MAXLENGTH=18  MAXLENGTH=5 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPrType()">
														   <INPUT TYPE=TEXT NAME="txtPrTypeNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="Tracking No." NAME="txtTrackNo" SIZE=34 MAXLENGTH=25  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingNo()"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=240 valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT2> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:CookiePage(1)">���ſ�û���</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPrFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPrToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPdFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPdToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPrStsCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRqDeptCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPrTypeCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnTrackNo" tag="24">
</FORM>
	<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>
</BODY>
</HTML>
