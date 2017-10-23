
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4255ra2
'*  4. Program Name         : ��ȯ��ȣ�˾� 
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
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "F4255RB2.asp"												'��: �����Ͻ� ���� ASP�� 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey          = 6                                    '��: key count of SpreadSheet

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

'Dim lgSelectList
'Dim lgSelectListDT

'Dim lgSortFieldNm
'Dim lgSortFieldCD

Dim lgMaxFieldCount

'Dim lgPopUpR
'Dim lgKeyPos
'Dim lgKeyPosVal
Dim lgCookValue

Dim lgSaveRow
'--------------- ������ coding part(��������,Start)-----------------------------------------------------------

Dim arrReturn
Dim arrParent
Dim arrParam
Dim txtPayNo
Dim txtBankPayCd
Dim txtPayFromDt
Dim txtPaytoDt

	 '------ Set Parameters from Parent ASP ------ 
	arrParent = window.DialogArguments
	Set PopupParent = arrParent(0)
	arrParam = arrParent(0)
	
	top.document.title = "��ȯ��ȣ�����˾�"
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
    Redim arrReturn(0,1)

'    lgStrPrevKey		= ""
    lgPageNo			= ""
    lgIntFlgMode		= popupparent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue	= False                    'Indicates that no value changed
	lgSortKey			= 1
	lgSaveRow			= 0
	Self.Returnvalue	= arrReturn

End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= 
Sub SetDefaultVal()

'--------------- ������ coding part(�������,Start)--------------------------------------------------
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
	
	frm1.txtPayFromDt.Text	= frDt
	frm1.txtPayToDt.Text	= toDt

'--------------- ������ coding part(�������,End)----------------------------------------------------
	
End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "*", "NOCOOKIE", "RA") %>                                '��: 
	<% Call LoadBNumericFormatA("Q", "*","NOCOOKIE","RA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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

	frm1.vspdData.OperationMode = 3
	
    Call SetZAdoSpreadSheet("F4255RA201","S","A","V20030522",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
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
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
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
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
   
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
    Call InitSpreadSheet()

    frm1.txtPayFromDt.focus
    
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
	
	If CompareDateByFormat(frm1.txtPayFromDt.text,frm1.txtPaytoDt.text,frm1.txtPayFromDt.Alt,frm1.txtPaytoDt.Alt, _
        	               "970025",frm1.txtPayFromDt.UserDefinedFormat,popupparent.gComDateType, true) = False Then	   
	   frm1.txtPayFromDt.focus
	   Set gActiveElement = document.ActiveElement
	   
		Exit Function
	End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

    FncQuery = True													

End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    On Error Resume Next
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow() 
    On Error Resume Next
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
    On Error Resume Next
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
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
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
	
	frm1.vspdData.ScrollBars = popupparent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = popupparent.SS_SCROLLBAR_BOTH
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

	Err.Clear                                                                   '��: Protect system from crashing
	DbQuery = False

	Call LayerShowHide(1)
	
	With frm1
		strVal = BIZ_PGM_ID
    '---------Developer Coding part (Start)----------------------------------------------------------------

		If lgIntFlgMode <> popupparent.OPMD_UMODE Then		
		
			strVal = strVal & "?txtPayFromDt	=" & frm1.txtPayFromDt.Year & Right("0" & frm1.txtPayFromDt.Month,2) & Right("0" & frm1.txtPayFromDt.Day,2)
			strVal = strVal & "&txtPaytoDt		=" & frm1.txtPaytoDt.Year & Right("0" & frm1.txtPaytoDt.Month,2) & Right("0" & frm1.txtPaytoDt.Day,2)
			strVal = strVal & "&txtPayNo		=" & Trim(.txtPayNo.value)
			strVal = strVal & "&txtBankPayCd	=" & Trim(.txtBankPayCd.value)
		Else
		
			strVal = strVal & "?txtPayFromDt	=" & Trim(.hPayFromDt.value)
			strVal = strVal & "&txtPaytoDt		=" & Trim(.hPayToDt.value)
			strVal = strVal & "&txtPayNo		=" & Trim(.hPayNo.value)
			strVal = strVal & "&txtBankPayCd	=" & Trim(.hBankPayCd.value)
		End if
	'---------Developer Coding part (End)----------------------------------------------------------------
		strVal = strVal & "&lgPageNo="			& lgPageNo								'Next key tag
		strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")					'field type
		strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))     			'order by ���� ��������� 
		
		
		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
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
		frm1.txtPayFromDt.focus
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
	if frm1.vspdData.SelModeSelCount > 0 Then 			
		Redim arrReturn(frm1.vspdData.SelModeSelCount - 1,C_MaxKey)
		kk = 0
		For ii = 0 To frm1.vspdData.MaxRows - 1
			frm1.vspdData.Row = ii + 1			
			If frm1.vspdData.SelModeSelected Then
				For jj = 1 To C_MaxKey 
					frm1.vspdData.Col	 = GetKeyPos("A",jj )		
					arrReturn(kk,jj) = frm1.vspdData.Text
				Next			
				kk = kk + 1
			End If
		Next	
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
'           Call DisableToolBar(popupparent.TBC_QUERY)
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

'========================================================================================================
' Function Name : Radio_Loan_Fg
' Function Desc : 
'========================================================================================================
Sub Radio_Loan_Fg()
	If frm1.Radio_Loan_fg1.checked Then
		txtLoanType = frm1.Radio_Loan_fg1.value
	ElseIf frm1.Radio_Loan_fg2.checked Then
		txtLoanType = frm1.Radio_Loan_fg2.value
	ElseIf frm1.Radio_Loan_fg3.checked Then
		txtLoanType = frm1.Radio_Loan_fg3.value
	End if
End Sub



 '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 
Function OpenPopUp(Byval strCode,Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	
	Select Case iWhere
		Case 1
		   	arrParam(0) = "��ȯ��ȣ�˾�"	
			arrParam(1) = "F_LN_REPAY"				
			arrParam(2) = strCode
			arrParam(3) = "" 
			arrParam(4) = "ST_ADV_INT_FG <> " & FilterVar("IA", "''", "S") & "  AND CONF_FG IN (" & FilterVar("E", "''", "S") & " ," & FilterVar("C", "''", "S") & " )"
			arrParam(5) = frm1.txtPayNo.Alt
	
			arrField(0) = "PAY_NO"
			arrField(1) = "PAY_DT"
    
			arrHeader(0) = frm1.txtPayNo.Alt
			arrHeader(1) = frm1.txtPayFromDt.Alt
			arrHeader(2) = frm1.txtPayToDt.Alt
		Case 2
			arrParam(0) = "�����˾�"
			arrParam(1) = "B_BANK A"
			arrParam(2) = frm1.txtBankPayCd.value
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = frm1.txtBankPayCd.Alt
	
			arrField(0) = "A.BANK_CD"
			arrField(1) = "A.BANK_NM"
					    
			arrHeader(0) = frm1.txtBankPayCd.Alt
			arrHeader(1) = frm1.txtBankPayCd.Alt
		Case Else
		Exit Function
    End Select    
    
		lgIsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
			Select Case iWhere
			Case 1
				frm1.txtPayNo.focus
			Case 2
				frm1.txtBankPayCd.focus
			End Select
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SubSetSchoolInf()
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SetPopUp(arrRet,iWhere)
	With Frm1
		Select Case iWhere
		Case 1
			.txtPayNo.value = arrRet(0)
			.txtPayNo.focus
		Case 2
			.txtBankPayCd.value = arrRet(0)
			.txtBankPayNm.value = arrRet(1)
			.txtBankPayCd.focus
		End Select
	End With
End Sub

'==========================================================================================
'   Event Name : DblClick
'   Event Desc :
'==========================================================================================
Sub txtPayFromDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPayFromDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtPayFromDt.Focus
	End if
End Sub

Sub txtPaytoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPaytoDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtPaytoDt.Focus
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
Sub txtPayFromDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtPaytoDt.Focus
		Call FncQuery()
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

Sub txtPaytoDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtPayFromDt.Focus
		Call FncQuery()
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub


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
						<TD CLASS=TD5 NOWRAP>��ȯ����</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/f4255ra2_txtPayFromDt_txtPayFromDt.js'></script>&nbsp;~&nbsp;<script language =javascript src='./js/f4255ra2_txtPaynToDt_txtPayToDt.js'></script></TD>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��ȯ��ȣ</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayNo" ALT="��ȯ��ȣ" SIZE="20" MAXLENGTH="18"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankPayCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtPayNo.value,1)"></TD>
						<TD CLASS=TD5 NOWRAP>��������</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBankPayCd" ALT="��������" SIZE="10" MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankPayCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBankPayCd.value,2)">&nbsp;<INPUT NAME="txtBankPayNm" ALT="���������" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>
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
						<script language =javascript src='./js/f4255ra2_vspdData_vspdData.js'></script>
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
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hPayFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hPayToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hPayNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hBankPayCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</B
