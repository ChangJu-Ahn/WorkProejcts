<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales ���� 
'*  2. Function Name        : 
'*  3. Program ID           : s2114ra1.asp
'*  4. Program Name         : �ǸŰ�ȹ�����ڷ� 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS2G141.dll
'*  7. Modified date(First) : 2000/10/10
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Cho song-hyon
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/10/10 : 4th ȭ�� Layout ���� 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<!--TITLE>�ǸŰ�ȹ�����ڷ�</TITLE-->
<TITLE></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script Language="VBS">
	Option Explicit					'��: indicates that All variables must be declared in advance

	Dim C_ItemCode ' 1
	Dim C_ItemName ' 2
	Dim C_PlanUnit ' 3
	Dim C_YearQty ' 4
	Dim C_YearAmt ' 5

	Dim C_01PlanQty ' 6
	Dim C_02PlanQty ' 8
	Dim C_03PlanQty ' 10
	Dim C_04PlanQty ' 12
	Dim C_05PlanQty ' 14
	Dim C_06PlanQty ' 16
	Dim C_07PlanQty ' 18
	Dim C_08PlanQty ' 20
	Dim C_09PlanQty ' 22
	Dim C_10PlanQty ' 24
	Dim C_11PlanQty ' 26
	Dim C_12PlanQty ' 28

	Dim C_01PlanAmt ' 7
	Dim C_02PlanAmt ' 9
	Dim C_03PlanAmt ' 11
	Dim C_04PlanAmt ' 13
	Dim C_05PlanAmt ' 15
	Dim C_06PlanAmt ' 17
	Dim C_07PlanAmt ' 19
	Dim C_08PlanAmt ' 21
	Dim C_09PlanAmt ' 23
	Dim C_10PlanAmt ' 25
	Dim C_11PlanAmt ' 27
	Dim C_12PlanAmt ' 29

	Dim arrParent
	ArrParent = window.dialogArguments
	Set PopupParent  = ArrParent(0)
	top.document.title = PopupParent.gActivePRAspName

	Dim iDBSYSDate
	Dim EndDate, StartDate

	iDBSYSDate = "<%=GetSvrDate%>"
	'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
	EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

	'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
	StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

	Const BIZ_PGM_QRY_ID = "s2114rb1.asp"			<% '��: �����Ͻ� ���� ASP�� %>

	Dim arrReturn						'--- Return Parameter Group
	Dim lgIntGrpCount					'��: Group View Size�� ������ ���� 
	Dim lgIntFlgMode					'��: Variable is for Operation Status

	Dim lgStrPrevKey
	Dim gblnWinEvent					'~~~ ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
										'	PopUp Window�� ��������� ���θ� ��Ÿ���� variable
	Dim lgSortKey						'Sort Key
	Const lsPLANNUM  = "PLANNUM"		'��ȹ���� 

'========================================================================================================
Sub initSpreadPosVariables()  
	C_ItemCode = 1
	C_ItemName = 2
	C_PlanUnit = 3
	C_YearQty = 4
	C_YearAmt = 5

	C_01PlanQty = 6
	C_02PlanQty = 8
	C_03PlanQty = 10
	C_04PlanQty = 12
	C_05PlanQty = 14
	C_06PlanQty = 16
	C_07PlanQty = 18
	C_08PlanQty = 20
	C_09PlanQty = 22
	C_10PlanQty = 24
	C_11PlanQty = 26
	C_12PlanQty = 28

	C_01PlanAmt = 7
	C_02PlanAmt = 9
	C_03PlanAmt = 11
	C_04PlanAmt = 13
	C_05PlanAmt = 15
	C_06PlanAmt = 17
	C_07PlanAmt = 19
	C_08PlanAmt = 21
	C_09PlanAmt = 23
	C_10PlanAmt = 25
	C_11PlanAmt = 27
	C_12PlanAmt = 29
End Sub

'========================================================================================================
	Function InitVariables()
		lgIntGrpCount = 0										<%'��: Initializes Group View Size%>
		lgIntFlgMode = PopupParent.OPMD_CMODE					<%''Indicates that current mode is Create mode %>
		lgStrPrevKey = ""										<%'initializes Previous Key%>
		
		<% '------ Coding part ------ %>
		gblnWinEvent = False

		Redim arrReturn(0,0)
		Self.Returnvalue = arrReturn

	End Function
	
'========================================================================================================
	<% '== ��ȸ,��� == %>
	Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "RA") %>
	End Sub

'========================================================================================================
	Sub SetDefaultVal()

		Dim arrRowSep, arrColValue


		arrRowSep = Split(ArrParent(1), PopupParent.gRowSep)
		arrColValue = Split(arrRowSep(0),PopupParent.gColSep)

		<% '�������� %>
		txtConSalesOrg.value = arrColValue(0)
		txtConSalesOrgNm.value = arrColValue(1)
		<% '�ǸŰ�ȹ�⵵ %>
		txtConSpYear.Text = arrColValue(2)
		<% '��ȹ���� %>
		txtConPlanTypeCd.value = arrColValue(3)
		txtConPlanTypeNm.value = arrColValue(4)
		<% '�ŷ����� %>
		txtConDealTypeCd.value = arrColValue(5)
		txtConDealTypeNm.value = arrColValue(6)
		<% 'ȭ�� %>
		txtConCurr.value = arrColValue(7)

		txtBasicInfo.value = UCase(arrColValue(8))
		txtInfo.value = rdoInfoS.value

		txtSalesTitle.value = UCase(arrColValue(9))

		<% '�����׷�/�������� ���� %>
		Select Case UCase(txtSalesTitle.value)
		Case "GRP"
			lblSalesTitle.innerHTML = "�����׷�"
			txtConSalesOrg.Alt= "�����׷�"
		Case "ORG"
			lblSalesTitle.innerHTML = "��������"
			txtConSalesOrg.Alt= "��������"
		End Select

		If Trim(txtConSpYear.Text) = 0 Then txtConSpYear.Text = Year(UniConvDateToYYYYMMDD(EndDate,PopupParent.gDateFormat,PopupParent.gServerDateType))
        

		txtConSpYearFrom.Text = CInt(txtConSpYear.Text) - 2
		txtConSpYearTo.Text = Cint(txtConSpYear.Text) - 1

	End Sub
	
'========================================================================================================
	Sub InitSpreadSheet()
	
		Call initSpreadPosVariables()    	
		
		ggoSpread.Source = vspdData
		ggoSpread.Spreadinit "V20021120",,PopupParent.gAllowDragDropSpread    				

		vspdData.OperationMode = 5	'Multi Select Mode

		vspdData.ReDraw = False

		vspdData.MaxCols = C_12PlanAmt + 1	'���ΰ�ħ 
		vspdData.MaxRows = 0

	    Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit C_ItemCode, "ǰ��׷캰", 20,,,18,2
	    ggoSpread.SSSetEdit C_ItemName, "ǰ��׷��", 30
	    ggoSpread.SSSetEdit C_PlanUnit, "��ȹ����", 10,,,3,2
		ggoSpread.SSSetFloat C_YearQty,"�� ��ȹ�� �հ�" ,20,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_YearAmt,"�� ��ȹ�ݾ� �հ�",20,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_01PlanQty,"1����ȹ��" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_02PlanQty,"2����ȹ��" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_03PlanQty,"3����ȹ��" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_04PlanQty,"4����ȹ��" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_05PlanQty,"5����ȹ��" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_06PlanQty,"6����ȹ��" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_07PlanQty,"7����ȹ��" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_08PlanQty,"8����ȹ��" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_09PlanQty,"9����ȹ��" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_10PlanQty,"10����ȹ��" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_11PlanQty,"11����ȹ��" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_12PlanQty,"12����ȹ��" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"								
		ggoSpread.SSSetFloat C_01PlanAmt,"1����ȹ�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_02PlanAmt,"2����ȹ�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_03PlanAmt,"3����ȹ�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"		
		ggoSpread.SSSetFloat C_04PlanAmt,"4����ȹ�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_05PlanAmt,"5����ȹ�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_06PlanAmt,"6����ȹ�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_07PlanAmt,"7����ȹ�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_08PlanAmt,"8����ȹ�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"		
		ggoSpread.SSSetFloat C_09PlanAmt,"9����ȹ�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_10PlanAmt,"10����ȹ�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_11PlanAmt,"11����ȹ�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat C_12PlanAmt,"12����ȹ�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"

		ggoSpread.SpreadLockWithOddEvenRowColor()

		Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)

		vspdData.ReDraw = True

	End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCode			= iCurColumnPos(1)
			C_ItemName			= iCurColumnPos(2)
			C_PlanUnit			= iCurColumnPos(3)
			C_YearQty			= iCurColumnPos(4)
			C_YearAmt			= iCurColumnPos(5)
			
			C_01PlanQty			= iCurColumnPos(6)
			C_02PlanQty			= iCurColumnPos(8)
			C_03PlanQty			= iCurColumnPos(10)
			C_04PlanQty			= iCurColumnPos(12)
			C_05PlanQty			= iCurColumnPos(14)
			C_06PlanQty			= iCurColumnPos(16)
			C_07PlanQty			= iCurColumnPos(18)
			C_08PlanQty			= iCurColumnPos(20)
			C_09PlanQty			= iCurColumnPos(22)
			C_10PlanQty			= iCurColumnPos(24)
			C_11PlanQty			= iCurColumnPos(26)
			C_12PlanQty			= iCurColumnPos(28)
			
			C_01PlanAmt			= iCurColumnPos(7)
			C_02PlanAmt			= iCurColumnPos(9)
			C_03PlanAmt			= iCurColumnPos(11)
			C_04PlanAmt			= iCurColumnPos(13)
			C_05PlanAmt			= iCurColumnPos(15)
			C_06PlanAmt			= iCurColumnPos(17)
			C_07PlanAmt			= iCurColumnPos(19)
			C_08PlanAmt			= iCurColumnPos(21)
			C_09PlanAmt			= iCurColumnPos(23)
			C_10PlanAmt			= iCurColumnPos(25)
			C_11PlanAmt			= iCurColumnPos(27)
			C_12PlanAmt			= iCurColumnPos(29)			
		
	End Select

End Sub	

'========================================================================================================
	Sub InitComboBox()

	    Err.Clear                                                               <%'��: Protect system from crashing%>
	    
	    
		If   LayerShowHide(1) = False Then
             Exit Sub
        End If
	    
	    Dim strVal
	    
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & lsPLANNUM						<%'��: �����Ͻ� ó�� ASP�� ���� %>
		strVal = strVal & "&txtConSalesOrg=" & Trim(txtConSalesOrg.value)		<%'��: ��ȸ ���� ����Ÿ %>
		strVal = strVal & "&txtConSpYear=" & Trim(txtConSpYear.Text)
		strVal = strVal & "&txtConPlanTypeCd=" & Trim(txtConPlanTypeCd.value)
		strVal = strVal & "&txtConDealTypeCd=" & Trim(txtConDealTypeCd.value)
		strVal = strVal & "&txtConCurr=" & Trim(txtConCurr.value)
		strVal = strVal & "&txtSelectChr=" & Trim(txtSelectChr.value)
		
		Call RunMyBizASP(MyBizASP, strVal)										<%'��: �����Ͻ� ASP �� ���� %>
		
	End Sub
	
'========================================================================================================
	Function OKClick()
	
		Dim intColCnt, intRowCnt, intInsRow

		If vspdData.SelModeSelCount > 0 Then 

			intInsRow = 0

			Redim arrReturn(vspdData.SelModeSelCount, vspdData.MaxCols)

			For intRowCnt = 1 To vspdData.MaxRows

				vspdData.Row = intRowCnt

				If vspdData.SelModeSelected Then
					For intColCnt = 0 To vspdData.MaxCols - 1
						vspdData.Col = intColCnt + 1
						arrReturn(intInsRow, intColCnt) = vspdData.Text
					Next
					intInsRow = intInsRow + 1
				End IF

			Next

		End if			
		
		Self.Returnvalue = arrReturn
		Self.Close()
	End Function	

'========================================================================================================
	Function CancelClick()
		Self.Close()
	End Function
	
'========================================================================================================
Function OpenPlanNumber()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "��ȹ����"				<%' �˾� ��Ī %>
	arrParam(1) = "B_MINOR"						<%' TABLE ��Ī %>
	arrParam(2) = Trim(cboConPlanNum.Value)		<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD=" & FilterVar("S2001", "''", "S") & ""			<%' Where Condition%>
	arrParam(5) = "��ȹ����"				<%' TextBox ��Ī %>
		 
	arrField(0) = "MINOR_CD"					<%' Field��(0)%>
	arrField(1) = "MINOR_NM"					<%' Field��(1)%>
		    
	arrHeader(0) = "��ȹ����"				<%' Header��(0)%>
	arrHeader(1) = "��ȹ������"				<%' Header��(1)%>

	cboConPlanNum.focus 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then	
		Exit Function	
	Else	
		Call SetPlanNumber(arrRet)	
	End If 
 
End Function

'========================================================================================================
Function SetPlanNumber(Byval arrRet)
		cboConPlanNum.value = arrRet(0) 
		cboConPlanNumNm.value = arrRet(1)
		cboConPlanNum.focus 			
End Function

'========================================================================================================
Function OpenSaleOrg()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	Select Case UCase(Trim(txtSalesTitle.value))
	Case "GRP"
		arrParam(0) = "�����׷�"						<%' �˾� ��Ī %>
		arrParam(1) = "B_SALES_GRP"							<%' TABLE ��Ī %>
		arrParam(2) = Trim(txtConSalesOrg.Value)			<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "						<%' Where Condition%>
		arrParam(5) = "�����׷�"						<%' TextBox ��Ī %>
	
		arrField(0) = "SALES_GRP"							<%' Field��(0)%>
		arrField(1) = "SALES_GRP_NM"						<%' Field��(1)%>
    
		arrHeader(0) = "�����׷�"						<%' Header��(0)%>
		arrHeader(1) = "�����׷��"						<%' Header��(1)%>
		
		txtConSalesOrg.focus 
	Case "ORG"
		arrParam(0) = "��������"						<%' �˾� ��Ī %>
		arrParam(1) = "B_SALES_ORG"							<%' TABLE ��Ī %>
		arrParam(2) = Trim(txtConSalesOrg.Value)			<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = "END_ORG_FLAG=" & FilterVar("Y", "''", "S") & "  AND USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "	<%' Where Condition%>
		arrParam(5) = "��������"						<%' TextBox ��Ī %>
	
		arrField(0) = "SALES_ORG"							<%' Field��(0)%>
		arrField(1) = "SALES_ORG_NM"						<%' Field��(1)%>
    
		arrHeader(0) = "��������"						<%' Header��(0)%>
		arrHeader(1) = "����������"						<%' Header��(1)%>
		
		txtConSalesOrg.focus 		
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		txtConSalesOrg.value = arrRet(0) 
		txtConSalesOrgNm.value = arrRet(1)
		txtConSalesOrg.focus 
	End If	
	
End Function

'========================================================================================================
Function OpenPlanType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "��ȹ����"							<%' �˾� ��Ī %>
	arrParam(1) = "B_MINOR"									<%' TABLE ��Ī %>
	arrParam(2) = Trim(txtConPlanTypeCd.Value)				<%' Code Condition%>
	arrParam(3) = ""										<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD=" & FilterVar("S4089", "''", "S") & ""						<%' Where Condition%>
	arrParam(5) = "��ȹ����"							<%' TextBox ��Ī %>
	
	arrField(0) = "MINOR_CD"								<%' Field��(0)%>
	arrField(1) = "MINOR_NM"								<%' Field��(1)%>
    
	arrHeader(0) = "��ȹ����"							<%' Header��(0)%>
	arrHeader(1) = "��ȹ���и�"							<%' Header��(1)%>
	
	txtConPlanTypeCd.focus 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		txtConPlanTypeCd.value = arrRet(0) 
		txtConPlanTypeNm.value = arrRet(1)   
		txtConPlanTypeCd.focus 
	End If	
	
End Function


'=========================================================================== 
Function OpenDealType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "�ŷ�����"							<%' �˾� ��Ī %>
	arrParam(1) = "B_MINOR"									<%' TABLE ��Ī %>
	arrParam(2) = Trim(txtConDealTypeCd.Value)				<%' Code Condition%>
	arrParam(3) = ""										<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD=" & FilterVar("S4225", "''", "S") & ""						<%' Where Condition%>
	arrParam(5) = "�ŷ�����"							<%' TextBox ��Ī %>
	
	arrField(0) = "MINOR_CD"								<%' Field��(0)%>
	arrField(1) = "MINOR_NM"								<%' Field��(1)%>
    
	arrHeader(0) = "�ŷ�����"							<%' Header��(0)%>
	arrHeader(1) = "�ŷ����и�"							<%' Header��(1)%>

	txtConDealTypeCd.focus
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		txtConDealTypeCd.value = arrRet(0) 
		txtConDealTypeNm.value = arrRet(1)
		txtConDealTypeCd.focus 
	End If	
	
End Function

'=========================================================================== 
Function OpenCurr()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "ȭ��"								<%' �˾� ��Ī %>
	arrParam(1) = "B_CURRENCY"								<%' TABLE ��Ī %>
	arrParam(2) = Trim(txtConCurr.Value)					<%' Code Condition%>
	arrParam(3) = ""										<%' Name Cindition%>
	arrParam(4) = ""										<%' Where Condition%>
	arrParam(5) = "ȭ��"								<%' TextBox ��Ī %>
	
	arrField(0) = "CURRENCY"								<%' Field��(0)%>
	arrField(1) = "CURRENCY_DESC"							<%' Field��(1)%>
    
	arrHeader(0) = "ȭ��"								<%' Header��(0)%>
	arrHeader(1) = "ȭ���"								<%' Header��(1)%>
	txtConCurr.focus 
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		txtConCurr.value = arrRet(0)
	End If	
	
End Function

<%'=========================================================================== 
' Function Desc : This function is Grid Column Header Name of Change		
'=========================================================================== %>
Function GridHeadColName()

	Select Case txtBasicInfo.value
	Case UCase(rdoBasicInfoGrp.value)

		rdoBasicInfoGrp.checked = True

		vspdData.Col = C_ItemCode
		vspdData.Row = 0
		vspdData.Text = "ǰ��׷�"

		vspdData.Col = C_ItemName
		vspdData.Row = 0
		vspdData.Text = "ǰ��׷��"

		Call ggoSpread.SSSetColHidden(C_PlanUnit,C_PlanUnit,True)	
		
	Case UCase(rdoBasicInfoItem.value)

		rdoBasicInfoItem.checked = True
		
		vspdData.Col = C_ItemCode
		vspdData.Row = 0
		vspdData.Text = "ǰ��"

		vspdData.Col = C_ItemName
		vspdData.Row = 0
		vspdData.Text = "ǰ���"


	Case UCase(rdoBasicInfoCus.value)

		rdoBasicInfoCus.checked = True
		
		vspdData.Col = C_ItemCode
		vspdData.Row = 0
		vspdData.Text = "��"

		vspdData.Col = C_ItemName
		vspdData.Row = 0
		vspdData.Text = "����"

		Call ggoSpread.SSSetColHidden(C_PlanUnit,C_PlanUnit,True)				
	End Select

End Function

<%'=========================================================================== 
' Function Desc : �ڷ�������忡 ���� ���Ű�ȹ�⵵ Protected ����			
'=========================================================================== %>
Function SelectProtect()

	Select Case txtInfo.value
	Case rdoInfoS.value		<% '�����ǸŰ�ȹ %>
		Call ggoOper.SetReqAttr(txtConSpYearFrom, "N")
		Call ggoOper.SetReqAttr(txtConSpYearTo, "N")
		cboConPlanNum.value = ""
		cboConPlanNumNm.value = ""
		Call ggoOper.SetReqAttr(cboConPlanNum, "Q")
		window.document.btnConPlanNum.disabled = True
				
	Case rdoInfoP.value		<% '�������� %>
		Call ggoOper.SetReqAttr(txtConSpYearFrom, "Q")
		Call ggoOper.SetReqAttr(txtConSpYearTo, "Q")
		
		Call ggoOper.SetReqAttr(cboConPlanNum, "N")
		window.document.btnConPlanNum.disabled = False
	End Select

End Function


<%
'=======================================================================================================
' Function Desc : �� �Ǹż���/�ݾ��� �� 
'=======================================================================================================
%>
Function MonthTotalSum(GCol,GTotal)
    
	Dim SumTotal, iMonth, lRow, iCnt

    ggoSpread.Source = vspdData	
    
	For lRow = 1 To vspdData.MaxRows 

		SumTotal = 0

':: column 
		For iMonth = 0 To 22
			vspdData.Row = lRow
			vspdData.Col = GCol + iMonth
			If IsNumeric(vspdData.Text) = True Then
				SumTotal = UNICDbl(SumTotal) + UNICDbl(vspdData.Text)
			End If
			iMonth = iMonth + 1
		Next
'::

		vspdData.Row = lRow
		vspdData.Col = GTotal
		vspdData.Text= UNIFormatNumber(SumTotal,ggQty.DecPoint, -2, 0, ggQty.RndPolicy, ggQty.RndUnit)
	Next

End Function

<%
'======================================================================================================
' Function Desc : ���ڸ� �Է¹޴� ���� üũ 
'=======================================================================================================
%>
Function NumericCheck()

	Dim objEl, KeyCode
	
	Set objEl = window.event.srcElement
	KeyCode = window.event.keycode

	Select Case KeyCode
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
	Case Else
		window.event.keycode = 0
	End Select

End Function

'========================================================================================================
	Sub Form_Load()

		Call LoadInfTB19029											'��: Load table , B_numeric_format
		Call ggoOper.LockField(Document, "N")						<% '��: Lock  Suitable  Field %>
		Call AppendNumberPlace("6","4","0")
		Call AppendNumberPlace("7","3","0")
		Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,False,1)
		Call SetDefaultVal
		Call InitSpreadSheet()
		Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
		Call InitVariables

		Call SelectProtect()
		Call GridHeadColName()

	End Sub

'========================================================================================================
	Sub Form_QueryUnload(Cancel, UnloadMode)
	End Sub

'==========================================================================================
	Sub vspdData_Click(ByVal Col , ByVal Row)

		gMouseClickStatus = "SPC"

	    Set gActiveSpdSheet = vspdData

		If vspdData.MaxRows = 0 Then Exit Sub End If

		If Row <= 0 Then
			ggoSpread.Source = vspdData
			If lgSortKey = 1 Then
				ggoSpread.SSSort Col				'Sort in Ascending
				lgSortKey = 2
			Else
				ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
				lgSortKey = 1
			End If

			Exit Sub     

		End If

	End Sub

'========================================================================================================
	Function vspdData_DblClick(ByVal Col, ByVal Row)
		If Row = 0 Or vspdData.MaxRows = 0 Then 
		     Exit Function
		End If

		If vspdData.MaxRows > 0 Then
			If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End Function

'========================================================================================================
	Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
	    ggoSpread.Source = vspdData
	    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

	End Sub

'==========================================================================================
	Sub vspdData_MouseDown(Button , Shift , x , y)

	    If Button = 2 And gMouseClickStatus = "SPC" Then
	       gMouseClickStatus = "SPCR"
	    End If

	End Sub  
	
'========================================================================================================
	Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
		If OldLeft <> NewLeft Then
		    Exit Sub
		End If
    
		<% '----------  Coding part  -------------------------------------------------------------%>   
		if vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop)  Then	'��: ������ üũ 
			If lgStrPrevKey <> "" Then						'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If CheckRunningBizProcess = True Then
					Exit Sub
				End If	
				If DBQuery = False Then	
					Exit Sub
				End If
			End if
		End if	    


	End Sub

'==========================================================================================
    Function vspdData_KeyPress(KeyAscii)
         On Error Resume Next
         If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
            Call OKClick()
         ElseIf KeyAscii = 27 Then
            Call CancelClick()
         End If
    End Function

'==========================================================================================
<% 'ǰ��׷캰 %>
Sub rdoBasicInfoGrp_OnClick()
	txtBasicInfo.value = rdoBasicInfoGrp.value
	Call GridHeadColName()
End Sub

<% 'ǰ�� %>
Sub rdoBasicInfoItem_OnClick()
	txtBasicInfo.value = rdoBasicInfoItem.value
	Call GridHeadColName()
End Sub

<% '�ŷ�ó�� %>
Sub rdoBasicInfoCus_OnClick()
	txtBasicInfo.value = rdoBasicInfoCus.value
	Call GridHeadColName()
End Sub

<% '�����ǸŰ�ȹ %>
Sub rdoInfoS_OnClick()
	txtInfo.value = rdoInfoS.value
	Call SelectProtect()
End Sub

<% '�������� %>
Sub rdoInfoP_OnClick()
	txtInfo.value = rdoInfoP.value
	Call SelectProtect()
End Sub


<%
'=======================================================================================================
' Function Desc : ���ڸ� �Է¹޴� TextBox KeyIn �۾��� 
'=======================================================================================================
%>
Sub cboConPlanNum_onKeyPress()
	Call NumericCheck()
End Sub

<%
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : ��ȸ���Ǻ��� OCX_KeyDown�� EnterKey�� ���� Query
'==========================================================================================
%>
	Sub txtConSpYearFrom_Keypress(KeyAscii)
	    On Error Resume Next
	    If KeyAscii = 27 Then
	       Call CancelClick()
	    Elseif KeyAscii = 13 Then
	       Call FncQuery()
	    End if
	End Sub

	Sub txtConSpYearTo_Keypress(KeyAscii)
	    On Error Resume Next
	    If KeyAscii = 27 Then
	       Call CancelClick()
	    Elseif KeyAscii = 13 Then
	       Call FncQuery()
	    End if
	End Sub

	Sub txtConSpYear_Keypress(KeyAscii)
	    On Error Resume Next
	    If KeyAscii = 27 Then
	       Call CancelClick()
	    Elseif KeyAscii = 13 Then
	       Call FncQuery()
	    End if
	End Sub


'========================================================================================
	Function FncQuery()

		Err.Clear															<%'��: Protect system from crashing%>
		
'kek 0814	
		'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'�� 'pObjFromDt'���� ũ�ų� ���ƾ� �Ҷ� **

		'If ValidDateCheck(txtConSpYearFrom, txtConSpYearTo) = False Then Exit Function
		If Len(Trim(txtConSpYearTo.Text)) Then
		   If Len(Trim(txtConSpYearFrom.Text)) Then
				If txtConSpYearFrom.Text > txtConSpYearTo.Text Then
					Call DisplayMsgBox("970023","X", txtConSpYearTo.Alt, txtConSpYearFrom.Alt)
					txtConSpYearTo.Focus
					Set gActiveElement = document.activeElement                            
					Exit Function
				End If
			End If
		End If



		<% '------ Check condition area ------ %>
		If Not chkField(Document, "1") Then									<% '��: This function check indispensable field %>
			Exit Function
		End If

		FncQuery = False													<%'��: Processing is NG%>

		<% '------ Erase contents area ------ %>
		Call ggoOper.ClearField(Document, "2")								<% '��: Clear Contents  Field %>
		Call InitVariables													<% '��: Initializes local global variables %>

		Call DbQuery()

		FncQuery = True													<%'��: Processing is NG%>

	End Function

'========================================================================================
	Function DbQuery()

		Err.Clear															<%'��: Protect system from crashing%>

		DbQuery = False														<%'��: Processing is NG%>

		
		If   LayerShowHide(1) = False Then
             Exit Function 
        End If

		Dim strVal

		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001						<%'��: �����Ͻ� ó�� ASP�� ���� %>
			strVal = strVal & "&txtConSalesOrg=" & Trim(HConSalesOrg.value)			<%'��: ��ȸ ���� ����Ÿ %>
			strVal = strVal & "&txtConSpYear=" & Trim(HConSpYear.value)
			strVal = strVal & "&txtConPlanTypeCd=" & Trim(HPlanTypeCd.value)
			strVal = strVal & "&txtConDealTypeCd=" & Trim(HConDealTypeCd.value)
			strVal = strVal & "&txtConCurr=" & Trim(HConCurr.value)
			strVal = strVal & "&cboConPlanNum=" & Trim(HConPlanNum.value)
			strVal = strVal & "&txtConSpYearFrom=" & Trim(HConFrmYear.value)
			strVal = strVal & "&txtConSpYearTo=" & Trim(HConToYear.value)
			strVal = strVal & "&txtBasicInfo=" & Trim(HBasicInfo.value)			
			strVal = strVal & "&txtInfo=" & Trim(HInfo.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtSalesTitle=" & Trim(txtSalesTitle.value)

		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001						<%'��: �����Ͻ� ó�� ASP�� ���� %>
			strVal = strVal & "&txtConSalesOrg=" & Trim(txtConSalesOrg.value)		<%'��: ��ȸ ���� ����Ÿ %>
			strVal = strVal & "&txtConSpYear=" & Trim(txtConSpYear.Text)
			strVal = strVal & "&txtConPlanTypeCd=" & Trim(txtConPlanTypeCd.value)
			strVal = strVal & "&txtConDealTypeCd=" & Trim(txtConDealTypeCd.value)
			strVal = strVal & "&txtConCurr=" & Trim(txtConCurr.value)
			strVal = strVal & "&cboConPlanNum=" & Trim(cboConPlanNum.value)
			strVal = strVal & "&txtConSpYearFrom=" & Trim(txtConSpYearFrom.Text)
			strVal = strVal & "&txtConSpYearTo=" & Trim(txtConSpYearTo.Text)
			strVal = strVal & "&txtBasicInfo=" & Trim(txtBasicInfo.value)			
			strVal = strVal & "&txtInfo=" & Trim(txtInfo.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtSalesTitle=" & Trim(txtSalesTitle.value)
		End If

		Call RunMyBizASP(MyBizASP, strVal)									<%'��: �����Ͻ� ASP �� ���� %>

		DbQuery = True														<%'��: Processing is NG%>

	End Function

'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
    lgIntFlgMode = PopupParent.OPMD_UMODE												<%'��: Indicates that current mode is Update mode%>
'	Call SelectProtect()													<%' �ڷ������� ���� ���Ű�ȹ�⵵,��ȹ���� Protectó�� %>

	Call MonthTotalSum(C_01PlanQty,C_YearQty)
	Call MonthTotalSum(C_01PlanAmt,C_YearAmt)


	If vspdData.MaxRows > 0 Then
		vspdData.Focus
		vspdData.Row = 1	:	vspdData.SelModeSelected = True		
	Else
		txtConSalesOrg.focus
	End If

End Function

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()

End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP><SPAN CLASS="normal" ID="lblSalesTitle">&nbsp;</SPAN></TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtConSalesOrg" TYPE="Text" MAXLENGTH=4 SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSaleOrg()">&nbsp;<INPUT NAME="txtConSalesOrgNm" TYPE="Text" MAXLENGTH="20" SIZE=18 tag="14"></TD>
						<TD CLASS="TD5" NOWRAP>��ȹ�����⵵</TD>
						<TD CLASS="TD6" NOWRAP>
							<script language =javascript src='./js/s2114ra1_fpDoubleSingle1_txtConSpYear.js'></script>
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>��ȹ����</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtConPlanTypeCd" ALT="��ȹ����" TYPE="Text" MAXLENGTH=1 SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlanType()">&nbsp;<INPUT NAME="txtConPlanTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=18 tag="14"></TD>
						<TD CLASS="TD5" NOWRAP>�ŷ�����</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtConDealTypeCd" ALT="�ŷ�����" TYPE="Text" MAXLENGTH=1 SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemSalePlan" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDealType()">&nbsp;<INPUT NAME="txtConDealTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=18 tag="14"></TD>							</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>��ȹ����</TD>
						<TD CLASS="TD6"><INPUT NAME="cboConPlanNum" ALT="��ȹ����" TYPE="Text" MAXLENGTH=2 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConPlanNum" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlanNumber()">&nbsp;<INPUT NAME="cboConPlanNumNm" TYPE="Text" MAXLENGTH="20" SIZE=18 tag="14"></TD>
						<TD CLASS="TD5" NOWRAP>ȭ��</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtConCurr" ALT="ȭ��" TYPE="Text" MAXLENGTH=3 SIZE=10 tag="14XXXU"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>���Ű�ȹ�⵵</TD>
						<TD CLASS=TD6 NOWRAP>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<script language =javascript src='./js/s2114ra1_fpDoubleSingle3_txtConSpYearFrom.js'></script>
									</TD>
									<TD>
										&nbsp~&nbsp
									</TD>
									<TD>
										<script language =javascript src='./js/s2114ra1_fpDoubleSingle4_txtConSpYearTo.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>����ǸŰ�ȹ</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=radio CLASS="RADIO" NAME="rdoBasicInfo" id="rdoBasicInfoGrp" VALUE="G" tag = "14">
								<LABEL FOR="rdoBasicInfoGrp">ǰ��׷캰</LABEL>
							<INPUT TYPE=radio CLASS="RADIO" NAME="rdoBasicInfo" id="rdoBasicInfoItem" VALUE="T" tag = "14">
								<LABEL FOR="rdoBasicInfoItem">ǰ��</LABEL>
							<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoBasicInfo" id="rdoBasicInfoCus" VALUE="C" tag = "14">
								<LABEL FOR="rdoBasicInfoCus">�ŷ�ó��</LABEL></TD>
						<TD CLASS=TD5 NOWRAP>�ڷ�������</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=radio CLASS="RADIO" NAME="rdoInfo" id="rdoInfoS" VALUE="S" tag = "11" CHECKED>
								<LABEL FOR="rdoInfoS">�����ǸŰ�ȹ</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;
							<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoInfo" id="rdoInfoP" VALUE="P" tag = "11">
								<LABEL FOR="rdoInfoP">��������</LABEL></TD>
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
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/s2114ra1_vspdData_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO NORESIZE framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">

<INPUT TYPE=HIDDEN NAME="txtBasicInfo" tag="14">
<INPUT TYPE=HIDDEN NAME="txtInfo" tag="14">
<INPUT TYPE=HIDDEN NAME="txtSelectChr" tag="14">
<INPUT TYPE=HIDDEN NAME="txtSalesTitle" tag="14">

<INPUT TYPE=HIDDEN NAME="HConSalesOrg" tag="24">
<INPUT TYPE=HIDDEN NAME="HConSpYear" tag="24">
<INPUT TYPE=HIDDEN NAME="HPlanTypeCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HConDealTypeCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HConCurr" tag="24">
<INPUT TYPE=HIDDEN NAME="HConPlanNum" tag="24">

<INPUT TYPE=HIDDEN NAME="HConFrmYear" tag="24">
<INPUT TYPE=HIDDEN NAME="HConToYear" tag="24">
<INPUT TYPE=HIDDEN NAME="HBasicInfo" tag="24">
<INPUT TYPE=HIDDEN NAME="HInfo" tag="24">

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
