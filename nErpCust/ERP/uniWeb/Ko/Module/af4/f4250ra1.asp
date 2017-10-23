
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4250ra1
'*  4. Program Name         : ���Աݹ�ȣ�˾� 
'*  5. Program Desc         : Popup of Loan No.
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2002.02.19
'*  8. Modified date(Last)  : 2003.11.10
'*  9. Modifier (First)     : Oh, Soo Min
'* 10. Modifier (Last)      : ahn, do hyun ���� 
'* 11. Comment              : 
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">			</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "F4250RB1.asp"												'��: �����Ͻ� ���� ASP�� 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey          = 31                                   '��: key count of SpreadSheet

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================


'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Dim lgIsOpenPop
Dim lgMaxFieldCount
Dim lgCookValue
Dim lgSaveRow
'--------------- ������ coding part(��������,Start)-----------------------------------------------------------

Dim arrReturn
Dim arrParent
Dim arrParam

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

	'------ Set Parameters from Parent ASP ------ 
	arrParent        = window.dialogArguments
	Set PopupParent = arrParent(0)
	arrParam		= arrParent(1)
	
	top.document.title = "���Ա������˾�"
'--------------- ������ coding part(��������,End)-------------------------------------------------------------

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
	'If Trim("<%=Request("PGM")%>") = "F4250MA1" Then
	If Trim("<%=Request("PGM")%>") = "F4250MA1" Or Trim("<%=Request("PGM")%>") = "F4250MA1_KO441" Then
	    Redim arrReturn(0)
	Else
		Redim arrReturn(0,0)
	End If

    lgPageNo			= ""
    lgIntFlgMode		= popupparent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue	= False                    'Indicates that no value changed
	lgSortKey			= 1
	lgSaveRow			= 0
	Self.Returnvalue = arrReturn
	
	' ���Ѱ��� �߰� 
	If UBound(arrParam) > 5 Then
		lgAuthBizAreaCd		= arrParam(5)
		lgInternalCd		= arrParam(6)
		lgSubInternalCd		= arrParam(7)
		lgAuthUsrID			= arrParam(8)
	End If	
End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= 
Sub SetDefaultVal()
	Dim strSvrDate
	DIm strYear, strMonth, strDay
	Dim frDt, toDt
	Dim LastDate, FirstDate
	
	strSvrDate = "<%=GetSvrDate%>"
	LastDate     = UNIGetLastDay (strSvrDate,popupparent.gServerDateFormat)                                  'Last  day of this month
	FirstDate    = UNIGetFirstDay(strSvrDate,popupparent.gServerDateFormat)                                  'First day of this month

	Call ExtractDateFrom(strSvrDate, popupparent.gServerDateFormat, popupparent.gServerDateType, strYear,strMonth,strDay)
		
	frDt = UniConvYYYYMMDDToDate(popupparent.gDateFormat, strYear, strMonth, FirstDate)
	toDt = UniConvYYYYMMDDToDate(popupparent.gDateFormat, strYear, strMonth, LastDate)
	
	frm1.txtLoanFromDt.Text = frDt
	frm1.txtLoanToDt.Text = toDt
	
	frm1.hParentLoanNo.value = Trim(arrParam(0))
	frm1.hParentPayPlanDt.value = Trim(arrParam(1))
End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>                                '��: 
	<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","RA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Function CookiePage(ByVal Kubun)
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	'If Trim("<%=Request("PGM")%>") = "F4250MA1" Then
	If Trim("<%=Request("PGM")%>") = "F4250MA1" OR Trim("<%=Request("PGM")%>") = "F4250MA1_KO441" Then
		frm1.vspdData.OperationMode = 3
	Else
		frm1.vspdData.OperationMode = 5
	End If
	
	Call SetZAdoSpreadSheet("F4250RA101","S","A","V20030510",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock() 
End Sub

'========================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.vspdData.ReDraw = True
    End With
End Sub

'==================================++++==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal lRow)

End Sub

'===========================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : PopZAdoConfigGrid Reference Popup
'===========================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
		frm1.txtLoanFromDt.focus
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
	frm1.txtLoanFromDt.focus

End Function

'===========================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenSortPopup Reference Popup
'===========================================================================
Function OpenSortPopup()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & Popupparent.SORTW_WIDTH & "px; dialogHeight=" & Popupparent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
		frm1.txtLoanFromDt.focus
		Exit Function
	Else
		Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
		Call InitVariables
		Call InitSpreadSheet()       
	End If

	frm1.txtLoanFromDt.focus
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    '---------Developer Coding part (Start)----------------------------------------------------------------
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,popupparent.gDateFormat,popupparent.gComNum1000,popupparent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")												'��: Lock  Suitable  Field

    Call InitVariables																	'��: Initializes local global variables
    Call SetDefaultVal
	Call txtLoanPlcfg_onchange()
    Call InitSpreadSheet()

    frm1.txtLoanFromDt.focus
    
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 

    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    						
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If

	If CompareDateByFormat(frm1.txtLoanFromDt.text,frm1.txtLoanToDt.text,frm1.txtLoanFromDt.Alt,frm1.txtLoanToDt.Alt, _
        	               "970025",frm1.txtLoanFromDt.UserDefinedFormat,popupparent.gComDateType, true) = False Then	   
	   frm1.txtLoanFromDt.focus
	   Set gActiveElement = document.ActiveElement
	   
		Exit Function
	End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

    FncQuery = True													

End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint()
    FncPrint = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call popupparent.FncPrint()                                                       '��: Protect system from crashing
    FncPrint = True                                                              '��: Processing is OK
End Function


'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	Call popupparent.FncExport(popupparent.C_MULTI)

    FncExcel = True  
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    FncFind = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	Call popupparent.FncFind(popupparent.C_MULTI, True)

    FncFind = True                                                               '��: Processing is OK
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
	Dim ACol
	Dim ARow
	Dim iRet
	Dim iColumnLimit
	
	iColumnLimit = frm1.vspdData.MaxCols
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		Frm1.vspdData.Col = iColumnLimit	
		Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
End Function

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	Err.Clear
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncExit = True
    
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim txtLoanfg
	Dim txtLoanPlcfg

	Err.Clear                                                                   '��: Protect system from crashing
	DbQuery = False

	Call LayerShowHide(1)
	
	With frm1
		strVal = BIZ_PGM_ID
    '---------Developer Coding part (Start)----------------------------------------------------------------
    
		If frm1.txtLoanfg1.checked Then
			txtLoanfg = frm1.txtLoanfg1.value
		ElseIf frm1.txtLoanfg2.checked Then
			txtLoanfg = frm1.txtLoanfg2.value
		ElseIf frm1.txtLoanfg3.checked Then
			txtLoanfg = frm1.txtLoanfg3.value
		End if
		If frm1.txtLoanPlcfg1.checked Then
			txtLoanPlcfg = frm1.txtLoanPlcfg1.value
		ElseIf frm1.txtLoanPlcfg2.checked Then
			txtLoanPlcfg = frm1.txtLoanPlcfg2.value
		End if

		If lgIntFlgMode <> popupparent.OPMD_UMODE Then		
		
'			strVal = strVal & "?txtLoanFromDt=" & .txtLoanFromDt.Year & Right("0" & .txtLoanFromDt.Month,2) & Right("0" & .txtLoanFromDt.Day,2)
'			strVal = strVal & "&txtLoanToDt="	& .txtLoanToDt.Year & Right("0" & .txtLoanToDt.Month,2) & Right("0" & .txtLoanToDt.Day,2)
'			strVal = strVal & "&txtDocCur="		& Trim(.txtDocCur.value)
'			strVal = strVal & "&txtLoanfg="		& Trim(txtLoanfg)
'			strVal = strVal & "&txtLoanType="	& Trim(.txtLoanType.value)
'			strVal = strVal & "&txtLoanPlcfg="	& Trim(txtLoanPlcfg)
'			strVal = strVal & "&txtLoanPlcCd="	& Trim(.txtLoanPlcCd.value)
'			strVal = strVal & "&txtLoanNo="		& Trim(.txtLoanNo.value)
'			strVal = strVal & "&hParentLoanNo="	& Trim(.hParentLoanNo.value)
'			strVal = strVal & "&hParentPayPlanDt="	& Trim(.hParentPayPlanDt.value)

			.hLoanFromDt.value	= .txtLoanFromDt.Year & Right("0" & .txtLoanFromDt.Month,2) & Right("0" & .txtLoanFromDt.Day,2)
			.hLoanToDt.value	= .txtLoanToDt.Year & Right("0" & .txtLoanToDt.Month,2) & Right("0" & .txtLoanToDt.Day,2)
			.hDocCur.value		=  Trim(.txtDocCur.value)
			.hLoanfg.value		=  Trim(txtLoanfg)
			.hLoanType.value	=  Trim(.txtLoanType.value)
			.hLoanPlcfg.value	=  Trim(txtLoanPlcfg)
			.hLoanPlcCd.value	=  Trim(.txtLoanPlcCd.value)
			.hLoanNo.value		=  Trim(.txtLoanNo.value)
			.hParentLoanNo.value =  Trim(.hParentLoanNo.value)
			.hParentPayPlanDt.value =  Trim(.hParentPayPlanDt.value)
		Else
		
'			strVal = strVal & "?txtLoanFromDt=" & Trim(.hLoanFromDt.value)
'			strVal = strVal & "&txtLoanToDt="	& Trim(.hLoanToDt.value)
'			strVal = strVal & "&txtDocCur="		& Trim(.hDocCur.value)
'			strVal = strVal & "&txtLoanfg="		& Trim(.hLoanfg.value)
'			strVal = strVal & "&txtLoanType="	& Trim(.hLoanType.value)
'			strVal = strVal & "&txtLoanPlcfg="	& Trim(.hLoanPlcfg.value)
'			strVal = strVal & "&txtLoanPlcCd="	& Trim(.hLoanPlcCd.value)
'			strVal = strVal & "&txtLoanNo="		& Trim(.hLoanNo.value)
'			strVal = strVal & "&hParentLoanNo="	& Trim(.hParentLoanNo.value)
'			strVal = strVal & "&hParentPayPlanDt="	& Trim(.hParentPayPlanDt.value)
		End if
		
	'---------Developer Coding part (End)----------------------------------------------------------------
'		strVal = strVal & "&lgPageNo="			& lgPageNo								'Next key tag
'		strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")					'field type
'		strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
'		strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))               'order by ���� ��������� 

		.lgPageNo.value = lgPageNo								'Next key tag
		.lgSelectListDT.value = GetSQLSelectListDataType("A")					'field type
		.lgTailList.value =  MakeSQLGroupOrderByList("A")
        .lgSelectList.value = EnCoding(GetSQLSelectList("A"))               'order by ���� ��������� 

		' ���Ѱ��� �߰� 
		.lgAuthBizAreaCd.value = lgAuthBizAreaCd			' ����� 
		.lgInternalCd.value = lgInternalCd				' ���κμ� 
		.lgSubInternalCd.value = lgSubInternalCd			' ���κμ�(��������)
		.lgAuthUsrID.value = lgAuthUsrID				' ���� 
		
'		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
               
    End With
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
	lgBlnFlgChgValue	=False
	lgIntFlgMode		= popupparent.OPMD_UMODE												'��: Indicates that current mode is Update mode
	lgSaveRow			= 1
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtLoanFromDt.focus
	End If

'	Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
'	Call SetToolbar("110000000001111")										'��: ��ư ���� ���� 
End Function


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  �� �κп��� �÷� �߰��ϰ� ����Ÿ ������ �Ͼ�� �մϴ�.   							=
'========================================================================================================
Function OKClick()
	Dim ii ,jj ,kk
	
	'If Trim("<%=Request("PGM")%>") = "F4250MA1" Then						'�̱����Աݻ�ȯ�̸� 
	If Trim("<%=Request("PGM")%>") = "F4250MA1" Or Trim("<%=Request("PGM")%>") = "F4250MA1_KO441" Or Trim("<%=Request("PGM")%>") = "F4250MA1_TEST" Then		
		If frm1.vspdData.ActiveRow > 0 Then 				
			Redim arrReturn(C_MaxKey)										'�̱ۼ���Ʈ �̹Ƿ� 1���� �迭 �̿� 
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			For ii = 0 To C_MaxKey - 1
				frm1.vspdData.Col  = GetKeyPos("A",ii + 1)		
				arrReturn(ii) = frm1.vspdData.Text
			Next						
		End If	
	Else
		If frm1.vspdData.SelModeSelCount > 0 Then 							'��Ƽ��ȯ�̸� 
			Redim arrReturn(frm1.vspdData.SelModeSelCount - 1,C_MaxKey)		'��Ƽ����Ʈ �̹Ƿ� 2���� �迭 �̿� 
		
			kk = 0
			For ii = 0 To frm1.vspdData.MaxRows - 1
				frm1.vspdData.Row = ii + 1			
				If frm1.vspdData.SelModeSelected Then
					For jj = 1 To C_MaxKey 
						frm1.vspdData.Col	 = GetKeyPos("A",jj )		
						arrReturn(kk,jj-1) = frm1.vspdData.Text
					Next			
					kk = kk + 1
				End If
			Next	
		End If			
	End If	
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()			
End Function

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim ii
    gMouseClickStatus = "SPC"   
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
	
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
End Sub


'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
'Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
'End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : 
'==========================================================================================
'Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
'End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
	
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	 	 
    	If lgPageNo <> "" Then		
    		dbquery						
'           Call DisableToolBar(Parent.TBC_QUERY)
'           If DbQuery = False Then
'              Call RestoreToolBar()
'              Exit Sub
'			End if
    	End If
    End If
    
End Sub

'========================================================================================================
'   Event Name : vspdData_KeyPress
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================


'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtLoanPlcCd.focus
		Exit Function
	Else
		Call SetPopUP(arrRet, iWhere)
	End If

End Function
 '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

    Select case iWhere
	Case 0
		If frm1.txtLoanPlcfg1.Checked = true Then
			arrParam(0) = "�����˾�"
			arrParam(1) = "B_BANK A"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "�����ڵ�"

			arrField(0) = "A.BANK_CD"
			arrField(1) = "A.BANK_NM"
					    
			arrHeader(0) = "�����ڵ�"
			arrHeader(1) = "�����"
		Else
				Call OpenBp(strCode, iWhere)
				exit function
		End If
    Case 1	
        arrParam(0) = "���Կ뵵�˾�"			' �˾� ��Ī 
		arrParam(1) = "b_minor" 				    ' TABLE ��Ī 
		arrParam(2) = strCode						' Code Condition
		arrParam(3) = ""							' Name Cindition
		arrParam(4) = "major_cd=" & FilterVar("f1000", "''", "S") & " "	        ' Where Condition
		arrParam(5) = "���Կ뵵"				' �����ʵ��� �� ��Ī 

		arrField(0) = "minor_cd"						' Field��(0)
		arrField(1) = "minor_nm"						' Field��(1)
    
		arrHeader(0) = "���Կ뵵�ڵ�"				' Header��(0)
		arrHeader(1) = "���Կ뵵��"				    ' Header��(1)
	Case 2
			arrParam(0) = frm1.txtDocCur.Alt								' �˾� ��Ī 
			arrParam(1) = "B_CURRENCY"	 									' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = frm1.txtDocCur.Alt								' �����ʵ��� �� ��Ī 

		    arrField(0) = "CURRENCY"										' Field��(0)
		    arrField(1) = "CURRENCY_DESC"									' Field��(1)

		    arrHeader(0) = "��ȭ�ڵ�"									' Header��(0)
			arrHeader(1) = "��ȭ�ڵ��"									' Header��(1)
	Case 3
		arrParam(0) = "���Աݹ�ȣ�˾�"
		arrParam(1) = "f_ln_info A"
		arrParam(2) = strCode
		arrParam(3) = ""
		arrParam(4) = "A.CONF_FG IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("E", "''", "S") & " )"

		' ���Ѱ��� �߰� 
		If lgAuthBizAreaCd <> "" Then
			arrParam(4) = arrParam(4) & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
		End If

		If lgInternalCd <> "" Then
			arrParam(4) = arrParam(4) & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
		End If

		If lgSubInternalCd <> "" Then
			arrParam(4) = arrParam(4) & " AND A.INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
		End If

		If lgAuthUsrID <> "" Then
			arrParam(4) = arrParam(4) & " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
		End If


		arrParam(5) = frm1.txtLoanNo.Alt
	
		arrField(0) = "A.Loan_NO"
		arrField(1) = "A.Loan_NM"
				    
		arrHeader(0) = frm1.txtLoanNo.Alt
		arrHeader(1) = "���Ը�"

	Case Else
		Exit Function
	End Select

	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0		' ���� 
				frm1.txtLoanPlcCd.focus
			Case 1		'���Կ뵵 
				frm1.txtLoanType.Focus
			Case 2
				frm1.txtDocCur.Focus
			Case 3
				frm1.txtLoanNo.focus
		End Select
		Exit Function
	Else
		Call SetPopUP(arrRet, iWhere)
	End If

End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval iWhere)
	
	With frm1
		Select Case iWhere
			
			Case 0		' ���� 
				frm1.txtLoanPlcCd.value = arrRet(0)
				frm1.txtLoanPlcNm.value = arrRet(1)
				frm1.txtLoanPlcCd.focus
			Case 1		'���Կ뵵 
				frm1.txtLoanType.value = arrRet(0)
				frm1.txtLoanTypeNm.value = arrRet(1)
				frm1.txtLoanType.Focus
			Case 2
				frm1.txtDocCur.value = arrRet(0)
				frm1.txtDocCur.Focus
			Case 3
				frm1.txtLoanNo.value = arrRet(0)
				frm1.txtLoanNo.focus
		End Select

	End With
	
End Function

'==========================================================================================
'   Event Name : DblClick
'   Event Desc :
'==========================================================================================
Sub txtLoanFromDt_DblClick(Button)
	if Button = 1 then
		frm1.txtLoanFromDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtLoanFromDt.Focus
	End if
End Sub

Sub txtLoanToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtLoanToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtLoanToDt.Focus
	End if
End Sub

Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : KeyPress
'   Event Desc :
'==========================================================================================
Sub txtLoanFromDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtLoanToDt.Focus
		Call FncQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

Sub txtLoanToDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtLoanFromDt.Focus
		Call FncQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

'======================================================================================================
'   Event Name : txtLoanPlcfg_onchange
'   Event Desc : 
'=======================================================================================================
Function txtLoanPlcfg_onchange()
	If frm1.txtLoanPlcfg0.checked = true then
		Call ggoOper.SetReqAttr(frm1.txtLoanPlcCd, "Q")
		frm1.txtLoanPlcCd.value = ""
		frm1.txtLoanPlcNm.value = ""
	Else
		Call ggoOper.SetReqAttr(frm1.txtLoanPlcCd, "D")
	End If
End Function

'Function Radio1_onChange	
'	If frm1.cboIntBaseMthd.value <> "" Then
'		frm1.cboIntBaseMthd.value = ""
'	End If
'	
'	Call IntPayPerd_Change()
'	lgBlnFlgChgValue = True
'End Function

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>��ȯ��������</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtLoanFromDt name=txtLoanFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="13Z" ALT="������������"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtLoanToDt name=txtLoanToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="13Z" ALT="������������"></OBJECT>');</SCRIPT></TD>
						<TD CLASS="TD5" NOWRAP>�ŷ���ȭ</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" SIZE="10" MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.value,2)">
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��ܱⱸ��</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanfg ID=txtLoanfg1 VALUE="SLLL" Checked tag="11xxxU"><LABEL FOR=txtLoanfg1>�ܱ�+���</LABEL>&nbsp;
											 <INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanfg ID=txtLoanfg2 VALUE="SL" ><LABEL FOR=txtLoanfg2>�ܱ�</LABEL>&nbsp;
											 <INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanfg ID=txtLoanfg3 VALUE="LL" ><LABEL FOR=txtLoanfg3>���</LABEL></TD>
						<TD CLASS="TD5" NOWRAP>���Կ뵵</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanType" ALT="���Կ뵵" SIZE="10" MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanType.value,1)">
												<INPUT NAME="txtLoanTypeNm" ALT="���Կ뵵��" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>����ó����</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg0 VALUE="" Checked tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg0>����+�ŷ�ó</LABEL>&nbsp;
												<INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg1 VALUE="BK" tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg1>����</LABEL>&nbsp;
												<INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg2 VALUE="BP" tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg2>�ŷ�ó</LABEL></TD>
						<TD CLASS="TD5" NOWRAP>����ó</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanPlcCd" ALT="����ó" SIZE="10" MAXLENGTH="18"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankLoanCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanPlcCd.value,0)">
												<INPUT NAME="txtLoanPlcNm" ALT="����ó��" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>���Աݹ�ȣ</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanNo" ALT="���Աݹ�ȣ" SIZE="20" MAXLENGTH="18"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankLoanCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanNo.value,3)"></TD>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% id=vspdData tag="2" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"><PARAM NAME="ReDraw" VALUE="0"><PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ONCLICK="FncQuery()"></IMG>
					&nbsp;<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)" ONCLICK="OkClick()"></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hLoanFromDt"		tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hLoanToDt"			tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hDocCur"			tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hLoanfg"			tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hLoanType"			tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hLoanPlcfg"		tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hLoanPlcCd"		tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="hLoanNo"			tag="24" tabindex=-1>
<INPUT TYPE=hidden NAME="hParentLoanNo"		tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="hParentPayPlanDt"	tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgPageNo"			tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgMaxCount"		tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgSelectListDT"	tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgTailList"		tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgSelectList"		tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgAuthBizAreaCd"	tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgInternalCd"		tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgSubInternalCd"	tag="14" tabindex=-1>
<INPUT TYPE=hidden NAME="lgAuthUsrID"		tag="14" tabindex=-1>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

