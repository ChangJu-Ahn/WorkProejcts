<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4240ra1
'*  4. Program Name         : ���Աݹ�ȣ�˾� 
'*  5. Program Desc         : Popup of Loan No.
'*  6. Comproxy List        : DB agent
'*  7. Modified date(First) : 2001.02.19
'*  8. Modified date(Last)  : 2003.05.19
'*  9. Modifier (First)     : Hwang Eun Hee
'* 10. Modifier (Last)      : Ahn do hyun
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'############################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">			</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

Const BIZ_PGM_ID 		= "f4240rb1.asp"                              '��: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey          = 4                                           '��: key count of SpreadSheet
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                    
Dim lgMaxFieldCount
Dim lgCookValue 
Dim IsOpenPop  
Dim lgSaveRow 

Dim CPGM_ID
Dim arrReturn
Dim arrParent
Dim arrParam					

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

	 '------ Set Parameters from Parent ASP ------ 
	arrParent = Window.DialogArguments
	Set PopupParent = arrParent(0)
	arrParam = arrParent(1)
	
	Select Case Trim("<%=Request("PGM")%>")
	Case "F4240MA1"
		top.document.title = "�������� ���Աݹ�ȣ�˾�"
	Case "F4241MA1"
		top.document.title = "���������� ���Աݹ�ȣ�˾�"
	End Select

'--------------- ������ coding part(��������,End)-------------------------------------------------------------

 '#########################################################################################################
'												2. Function�� 
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

 '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = popupparent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1
    lgSaveRow        = 0

    Redim arrReturn(0)
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
	Dim strSvrDate, LastDate
	DIm strYear, strMonth, strDay
	Dim frDt, toDt
	
	strSvrDate = "<%=GetSvrDate%>"
	LastDate     = UNIGetLastDay ("<%=GetSvrDate%>",popupparent.gServerDateFormat) 
	
	Call ExtractDateFrom(strSvrDate, popupparent.gServerDateFormat, popupparent.gServerDateType, strYear,strMonth,strDay)
	frDt = UniConvYYYYMMDDToDate(popupparent.gDateFormat, strYear, strMonth, "01")
	
	Call ExtractDateFrom(LastDate,popupparent.gServerDateFormat,popupparent.gServerDateType,strYear,strMonth,strDay)
	toDt= UniConvYYYYMMDDToDate(popupparent.gDateFormat, strYear, strMonth, strDay)
	
	frm1.txtLoanFrDt.Text = frDt
	frm1.txtLoanToDt.Text = toDt   
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

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  �� �κп��� �÷� �߰��ϰ� ����Ÿ ������ �Ͼ�� �մϴ�.   							=
'========================================================================================================
Function OKClick()
	Dim ii

	If frm1.vspdData.ActiveRow > 0 Then 				
		Redim arrReturn(C_MaxKey)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		For ii = 0 To C_MaxKey - 1
			frm1.vspdData.Col  = GetKeyPos("A",ii + 1)		
			arrReturn(ii) = frm1.vspdData.Text
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

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
      
		frm1.vspdData.OperationMode = 3	
		Call SetZAdoSpreadSheet("F4240RA1","S","A","V20030407",popupparent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
		Call SetSpreadLock() 
     
End Sub


'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	 ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True

    End With
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

'========================================================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenSortPopup()

	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("./ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If

End Function


 '++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
 '******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,popupparent.gDateFormat,popupparent.gComNum1000,popupparent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
'    lgMaxFieldCount =  UBound(popupparent.gFieldNM)                      

'    ReDim lgPopUpR(popupparent.C_MaxSelList - 1,1)

'    Call popupparent.MakePopData(popupparent.gDefaultT,popupparent.gFieldNM,popupparent.gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,popupparent.C_MaxSelList)    ' You must not this line

    
    Call InitComboBox
	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal
	Call txtLoanPlcfg_onchange()
	Call InitSpreadSheet()
	    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
'	Call ElementVisible(frm1.txtDummy, 0)	'InVisible
'--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

 '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************

'==========================================================================================
'   Event Name : DblClick
'   Event Desc :
'==========================================================================================
Sub txtLoanFrDt_DblClick(Button)
	if Button = 1 then
		frm1.fpLoanFrDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpLoanFrDt.Focus
	End if
End Sub

Sub txtLoanToDt_DblClick(Button)
	if Button = 1 then
		frm1.fpLoanToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpLoanToDt.Focus
	End if
End Sub
Sub txtDueFrDt_DblClick(Button)
	if Button = 1 then
		frm1.fpDueFrDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpDueFrDt.Focus
	End if
End Sub

Sub txtDueToDt_DblClick(Button)
	if Button = 1 then
		frm1.fpDueToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpDueToDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : KeyPress
'   Event Desc :
'==========================================================================================
Sub txtLoanFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtLoanToDt.focus
		Call FncQuery
	End If
End Sub

Sub txtLoanToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtLoanFrDt.focus
		Call FncQuery
	End If
End Sub

Sub txtDueFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtLoanFrDt.focus
		Call FncQuery
	End If
End Sub

Sub txtDueToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtLoanFrDt.focus
		Call FncQuery
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

   	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ	
    	If lgPageNo <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Exit Sub
			End if
    	End If
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
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
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	
'	For ii = 1 to UBound(lgKeyPos)
'        frm1.vspdData.Col = lgKeyPos(ii)
'        frm1.vspdData.Row = Row
'        lgKeyPosVal(ii) = frm1.vspdData.text
'		lgCookValue = lgCookValue & Trim(lgKeyPosVal(ii)) & popupparent.gRowSep 
'	Next
'--------------- ������ coding part(�������,End)------------------------------------------------------
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
    
End Sub

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboLoanFg ,lgF0  ,lgF1  ,Chr(11))
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 

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
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	'Call InitData()
   	Call SetSpreadLock()

End Sub

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
   
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	
	If Row <> NewRow And NewRow > 0 Then
        Call SetSpreadColumnValue("A",frm1.vspdData,NewCol,NewRow)    
	End If
End Sub		

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	if UCase(frm1.txtLoanPlcCd.className) = "PROTECTED" Then Exit Function

	
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.PopupParent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtLoanFrDt.focus
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
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

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		case 0
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
    
			arrHeader(0) = frm1.txtLoanType.Alt				' Header��(0)
			arrHeader(1) = frm1.txtLoanTypeNm.Alt				    ' Header��(1)
		Case 2
			arrParam(0) = "�ŷ���ȭ�˾�"								' �˾� ��Ī 
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
			arrParam(0) = "���Աݹ�ȣ�˾�"								' �˾� ��Ī 
			arrParam(1) = "F_LN_INFO"	 									' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = " 1=1 "												' Where Condition

			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = arrParam(4) & " AND BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			End If

			If lgInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
			End If

			If lgSubInternalCd <> "" Then
				arrParam(4) = arrParam(4) & " AND INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
			End If

			If lgAuthUsrID <> "" Then
				arrParam(4) = arrParam(4) & " AND INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
			End If


			arrParam(5) = frm1.txtLoanNo.Alt								' �����ʵ��� �� ��Ī 

		    arrField(0) = "Loan_No"										' Field��(0)
		    arrField(1) = "Loan_Nm"									' Field��(1)

		    arrHeader(0) = frm1.txtLoanNo.Alt									' Header��(0)
			arrHeader(1) = frm1.txtLoanNm.Alt									' Header��(1)
		Case Else
			Exit Function
	End Select

	
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtLoanFrDt.focus
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If	

End Function

'------------------------------------------  SetReturnPopUp()  --------------------------------------------------
'	Name : SetReturnPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnPopUp(Byval arrRet, Byval iWhere)
	
	With frm1
		Select Case iWhere
			
			Case 0		' �ŷ�ó 
				frm1.txtLoanPlcCd.value = arrRet(0)
				frm1.txtLoanPlcNm.value = arrRet(1)
				frm1.txtLoanPlcCd.focus
			Case 1		'���Կ뵵 
				frm1.txtLoanType.value = arrRet(0)
				frm1.txtLoanTypeNm.value = arrRet(1)
				frm1.txtLoanType.Focus
			Case 2		'�ŷ���ȭ 
				frm1.txtDocCur.value = arrRet(0)
				frm1.txtDocCur.focus
			Case 3		'���Թ�ȣ 
				frm1.txtLoanNo.value = arrRet(0)
				frm1.txtLoanNm.value = arrRet(1)
				frm1.txtLoanNo.focus
		End Select

	End With
	
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

 '#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery() 
Dim IntRetCD
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
   
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
	
	If frm1.txtLoanFrDt.Text <> "" And frm1.txtLoanToDt.Text <> "" Then
		If CompareDateByFormat(frm1.txtLoanFrDt.Text, frm1.txtLoanToDt.Text, frm1.txtLoanFrDt.Alt, frm1.txtLoanToDt.Alt, _
							"970025", frm1.txtLoanFrDt.UserDefinedFormat, popupparent.gComDateType, true) = False Then
				frm1.txtLoanFrDt.focus											'��: GL Date Compare Common Function
				Exit Function
		End if
	End If
	If frm1.txtDueFrDt.Text <> "" And frm1.txtDueToDt.Text <> "" Then
		If CompareDateByFormat(frm1.txtDueFrDt.Text, frm1.txtDueToDt.Text, frm1.txtDueFrDt.Alt, frm1.txtDueToDt.Alt, _
							"970025", frm1.txtDueFrDt.UserDefinedFormat, popupparent.gComDateType, true) = False Then
				frm1.txtDueFrDt.focus											'��: GL Date Compare Common Function
				Exit Function
		End if
	End If

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
    Call popupparent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call popupparent.FncExport(popupparent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call popupparent.FncFind(popupparent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
	Dim strVal
	Dim txtLoanPlcfg
    DbQuery = False

    Err.Clear     
	Call LayerShowHide(1)
    
	If frm1.txtLoanPlcfg1.checked Then
		txtLoanPlcfg = frm1.txtLoanPlcfg1.value
	ElseIf frm1.txtLoanPlcfg2.checked Then
		txtLoanPlcfg = frm1.txtLoanPlcfg2.value
	End if

    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------
		If lgIntFlgMode <> popupparent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtLoanFrDt=" & Trim(.txtLoanFrDt.Text)
			strVal = strVal & "&txtLoanToDt="		& Trim(.txtLoanToDt.Text) 
			strVal = strVal & "&txtDueFrDt="		& Trim(.txtDueFrDt.Text) 
			strVal = strVal & "&txtDueToDt="		& Trim(.txtDueToDt.Text) 
			strVal = strVal & "&txtDocCur="			& Trim(.txtDocCur.value)
			strVal = strVal & "&txtLoanFg="			& Trim(.cboLoanFg.value)
			strVal = strVal & "&txtLoanType="		& Trim(.txtLoanType.value)
			strVal = strVal & "&txtLoanNo="			& Trim(.txtLoanNo.value)
			strVal = strVal & "&txtLoanPlcFg="		& Trim(txtLoanPlcFg)
			strVal = strVal & "&txtLoanPlcCd="		& Trim(.txtLoanPlcCd.value)
		Else 
			strVal = BIZ_PGM_ID & "?txtLoanFrDt=" & Trim(.hLoanFrDt.value)
			strVal = strVal & "&txtLoanToDt="		& Trim(.hLoanToDt.value)
			strVal = strVal & "&txtDueFrDt="		& Trim(.hDueFrDt.value) 
			strVal = strVal & "&txtDueToDt="		& Trim(.hDueToDt.value) 
			strVal = strVal & "&txtDocCur="			& Trim(.hDocCur.value)
			strVal = strVal & "&txtLoanFg="			& Trim(.hLoanFg.value)
			strVal = strVal & "&txtLoanType="		& Trim(.hLoanType.value)
			strVal = strVal & "&txtLoanNo="			& Trim(.hLoanNo.value)
			strVal = strVal & "&txtLoanPlcFg="		& Trim(.hLoanPlcFg.value)
			strVal = strVal & "&txtLoanPlcCd="		& Trim(.hLoanPlcCd.value)
		End If
'--------------- ������ coding part(�������,End)------------------------------------------------]
		strVal = strVal & "&txtPgmId=" & Trim("<%=Request("PGM")%>")
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

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
    lgBlnFlgChgValue = False                                                 'Indicates that no value changed
	lgIntFlgMode = popupparent.OPMD_UMODE
	lgSaveRow        = 1
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtLoanFrDt.focus
	End If
	
End Function

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


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################

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
						<TD CLASS=TD5 NOWRAP>�����</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpLoanFrDt name=txtLoanFrDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="���۰������"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpLoanToDt name=txtLoanToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="����������"></OBJECT>');</SCRIPT></TD>
						<TD CLASS="TD5" NOWRAP>�ŷ���ȭ</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" SIZE="10" MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.Value, 2)">
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��ȯ��������</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDueFrDt name=txtDueFrDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="���۸�������"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
											 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDueToDt name=txtDueToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="���Ḹ������"></OBJECT>');</SCRIPT></TD>
						<TD CLASS="TD5" NOWRAP>��ܱⱸ��</TD>
						<TD CLASS="TD6" NOWRAP><SELECT NAME="cboLoanFg" ALT="��ܱⱸ��" STYLE="WIDTH: 135px" tag="11"><OPTION VALUE=""></OPTION></SELECT>
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>���Աݹ�ȣ</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanNo" MAXLENGTH="18" SIZE=15  ALT ="���Աݹ�ȣ" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtLoanNo.Value,3)">
											   <INPUT NAME="txtLoanNm" MAXLENGTH="40" SIZE=20  ALT ="���Աݳ���" tag="14"></TD>
						<TD CLASS="TD5" NOWRAP>���Կ뵵</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtLoanType" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="���Կ뵵�ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtLoanType.Value,1)">
											   <INPUT TYPE="Text" NAME="txtLoanTypeNm" SIZE=20 tag="14X" ALT="���Կ뵵��">
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>����ó����</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg0 VALUE="" Checked tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg0>����+�ŷ�ó</LABEL>&nbsp;
												<INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg1 VALUE="BK" tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg1>����</LABEL>&nbsp;
												<INPUT TYPE="RADIO" CLASS="Radio" NAME=txtLoanPlcfg ID=txtLoanPlcfg2 VALUE="BP" tag="11xxxU" onClick=txtLoanPlcfg_onchange()><LABEL FOR=txtLoanPlcfg2>�ŷ�ó</LABEL></TD>
						<TD CLASS="TD5" NOWRAP>����ó</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanPlcCd" ALT="����ó" SIZE="10" MAXLENGTH="18"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanPlcCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanPlcCd.Value, 0)">
												<INPUT NAME="txtLoanPlcNm" ALT="����ó��" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>
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
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hLoanFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hDueFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hDueToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hDocCur" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanFg" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanType" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanPlcFg" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanPlcCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

