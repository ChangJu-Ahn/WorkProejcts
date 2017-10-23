
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Loan
'*  3. Program ID           : f4205ra1
'*  4. Program Name         : ���Աݹ�ȣ�˾� 
'*  5. Program Desc         : Popup of Loan No.
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.02.19
'*  8. Modified date(Last)  : 2001.11.10
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Oh, Soo Min
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
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                                             '��: Popup status                           

Dim lgMark

Dim IsOpenPop                                                  '��: ��ũ                                  
Dim CPGM_ID
'---------------  coding part(�������,Start)-----------------------------------------------------------
'   Call GetAdoFiledInf("F4205RA1","S","A")                        '��: spread sheet �ʵ����� query   -----
                                                                  ' 1. Program id
                                                                  ' 2. G is for Qroup , S is for Sort     
                                                                  ' 3. Spreadsheet no                                                               
                                                                
'--------------- ������ coding part(�������,End)-------------------------------------------------------------


'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "f4205rb1.asp"

Const C_MaxKey = 4

Dim arrReturn
Dim arrParent
Dim arrParam					

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd	' ����� 
Dim lgInternalCd	' ���κμ� 
Dim lgSubInternalCd	' ���κμ�(��������)
Dim lgAuthUsrID		' ���� 

	 '------ Set Parameters from Parent ASP ------ 
	arrParent		= window.dialogArguments
	Set PopupParent = arrParent(0)
	arrParam		= arrParent(1)
	
	If Trim("<%=Request("PGM")%>") = "F4235MA1"  Then
		top.document.title = "���⿬���ȣ�˾�"
	ElseIf Trim("<%=Request("PGM")%>") = "F4205MA1"  Then
		top.document.title = "�ŷ�ó���Թ�ȣ�˾�"
	ElseIf Trim("<%=Request("PGM")%>") = "F4206MA1"  Then
		top.document.title = "�ŷ�ó�������Թ�ȣ�˾�"
	Else
		top.document.title = "���Աݹ�ȣ�˾�"
	End If

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
    Redim arrReturn(0)
    
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgPageNo		= ""                                  'initializes Previous Key
    lgSortKey        = 1
    lgIntFlgMode     = PopupParent.OPMD_CMODE
    
	Self.Returnvalue = arrReturn


	' ���Ѱ��� �߰� 
	If UBound(arrParam) > 5 Then
		lgAuthBizAreaCd		= arrParam(5)
		lgInternalCd		= arrParam(6)
		lgSubInternalCd		= arrParam(7)
		lgAuthUsrID			= arrParam(8)
	End If
	
End Sub
'==========================================  2.1 InitComboBox()  ========================================
'	Name : InitComboBox()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitComboBox()
		
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboLoanFg ,lgF0  ,lgF1  ,Chr(11))    

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
	
	strSvrDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(strSvrDate, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear,strMonth,strDay)
		
	frDt = UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, "01")
	toDt = UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)
	
	frm1.txtLoanFromDt.Text = frDt
	frm1.txtLoanToDt.Text = toDt   
'--------------- ������ coding part(�������,End)----------------------------------------------------
	
End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE", "RA") %>	
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
	frm1.vspddata.OperationMode = 3 
    Call SetZAdoSpreadSheet("F4205RA1","S","A","V20030407",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
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

'**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 

 '-----------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------- 

'========================================================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
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

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
'   ReDim lgPopUpR(Parent.C_MaxSelList - 1,1)
	Call InitVariables														'��: Initializes local global variables
	Call InitComboBox()
	Call SetDefaultVal	
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
Sub txtLoanFromDt_DblClick(Button)
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

Sub txtDueFromDt_DblClick(Button)
	if Button = 1 then
		frm1.fpDuefrDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpDuefrDt.Focus
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

Sub txtDueFromDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtLoanFromDt.Focus
		Call FncQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

Sub txtDueToDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtLoanFromDt.Focus
		Call FncQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
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
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	If Row < 1 Then Exit Sub

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

 '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 
'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
	
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
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.value = arrRet(0)
		frm1.txtBpNm.value = arrRet(1)
		frm1.txtBpCd.focus
	End If	
End Function




Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0
			arrParam(0) = frm1.txtDocCur.Alt								' �˾� ��Ī 
			arrParam(1) = "B_CURRENCY"	 									' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = frm1.txtDocCur.Alt								' �����ʵ��� �� ��Ī 

		    arrField(0) = "CURRENCY"										' Field��(0)
		    arrField(1) = "CURRENCY_DESC"									' Field��(1)
'   
		    arrHeader(0) = "��ȭ�ڵ�"									' Header��(0)
			arrHeader(1) = "��ȭ�ڵ��"									' Header��(1)
			
		Case 3
			arrParam(0) = "���Աݹ�ȣ�˾�"
			arrParam(1) = "f_ln_info A"
			arrParam(2) = strCode
			arrParam(3) = ""
'			arrParam(4) = "A.CONF_FG IN ('C','E')"
			If Trim("<%=Request("PGM")%>") = "F4205MA1"  Then
				arrParam(4) = "A.LOAN_BASIC_FG = " & FilterVar("LN", "''", "S") & " "				
			ElseIf Trim("<%=Request("PGM")%>") = "F4206MA1"  Then
				arrParam(4) = "A.LOAN_BASIC_FG = " & FilterVar("LT", "''", "S") & " "				
			ElseIf Trim("<%=Request("PGM")%>") = "F4235MA1"  Then
				arrParam(4) = "A.LOAN_BASIC_FG = " & FilterVar("LR", "''", "S") & " "				
			Else
			End If
			arrParam(4) = arrParam(4) & "AND A.LOAN_PLC_TYPE = " & FilterVar("BP", "''", "S") & " "

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
		Case 5		'���԰ŷ�ó 
			arrParam(0) = "�ŷ�ó�˾�"
			arrParam(1) = "B_BIZ_PARTNER A"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = frm1.txtBpCd.Alt
	
			arrField(0) = "A.BP_CD"
			arrField(1) = "A.BP_NM"
			    
			arrHeader(0) = frm1.txtBpCd.Alt
			arrHeader(1) = frm1.txtBpNm.Alt

		Case 6		'���Կ뵵 
			arrParam(0) = "���Կ뵵�˾�"
			arrParam(1) = "B_MINOR A"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = "A.MAJOR_CD = " & FilterVar("F1000", "''", "S") & " "
			arrParam(5) = frm1.txtLoanType.Alt
	
			arrField(0) = "A.MINOR_CD"
			arrField(1) = "A.MINOR_NM"
			    
			arrHeader(0) = frm1.txtLoanType.Alt
			arrHeader(1) = frm1.txtLoanTypeNm.Alt
		
		Case Else
			Exit Function
	End Select
	
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtLoanFromDt.focus
		Exit Function
	End If

	Select Case iWhere
		Case 0	'�ŷ���ȭ 
			frm1.txtDocCur.value = arrRet(0)
			frm1.txtDocCur.focus
		Case 3	'���Աݹ�ȣ 
			frm1.txtLoanNo.value = arrRet(0)	
			frm1.txtLoanNo.focus
		Case 5	'�������� 
			frm1.txtBpCd.value = arrRet(0)
			frm1.txtBpNm.value = arrRet(1)
			frm1.txtBpCd.focus
		Case 6	'���Կ뵵 
			frm1.txtLoanType.value = arrRet(0)
			frm1.txtLoanTypeNm.value = arrRet(1)
			frm1.txtLoanType.focus
	End Select
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
	
	If frm1.txtLoanFromDt.Text <> "" And frm1.txtLoanToDt.Text <> "" Then
		If CompareDateByFormat(frm1.txtLoanFromDt.Text, frm1.txtLoanToDt.Text, frm1.txtLoanFromDt.Alt, frm1.txtLoanToDt.Alt, _
							"970025", frm1.txtLoanFromDt.UserDefinedFormat, popupparent.gComDateType, true) = False Then
				frm1.txtLoanFromDt.focus											'��: GL Date Compare Common Function
				Exit Function
		End if
	End If

	If frm1.txtDueFromDt.Text <> "" And frm1.txtDueToDt.Text <> "" Then
		If CompareDateByFormat(frm1.txtDueFromDt.Text, frm1.txtDueToDt.Text, frm1.txtDueFromDt.Alt, frm1.txtDueToDt.Alt, _
					"970025", frm1.txtDueFromDt.UserDefinedFormat, popupparent.gComDateType, true) = False Then
			frm1.txtDueFromDt.focus											'��: GL Date Compare Common Function
			Exit Function
		End if
	End If
	
		
    '-----------------------
    'Query function call area
    '-----------------------
	
    Call DbQuery															'��: Query db data

    FncQuery = True		
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call PopupParent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call PopupParent.FncExport(PopupParent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call PopupParent.FncFind(PopupParent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
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

    DbQuery = False
    
    Err.Clear           
'    frm1.vspdData.MaxRows = 0                                                   '��: Protect system from crashing                                                    '��: Protect system from crashing
	Call LayerShowHide(1)       
	     
    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------
'	If lgIntFlgMode = Parent.OPMD_UMODE Then
'		strVal = BIZ_PGM_ID & "?txtLoanFromDt=" & Trim(.hLoanFromDt.value)
'		strVal = strVal & "&txtLoanToDt=" & Trim(.hLoanToDt.value)
'		strVal = strVal & "&txtDueFromDt=" & Trim(.hDueFromDt.value)
'		strVal = strVal & "&txtDueToDt=" & Trim(.hDueToDt.value)
'		strVal = strVal & "&txtLoanType=" & Trim(.hLoanType.value)
''		strVal = strVal & "&txtLoanType_Alt=" & Trim(.txtLoanType.Alt)
'		strVal = strVal & "&txtBpCd=" & Trim(.hBankLoanCd.value)
'		strVal = strVal & "&txtBpCd_Alt=" & Trim(.txtBpCd.Alt)
'	Else 
		strVal = BIZ_PGM_ID & "?txtLoanFromDt=" & Trim(.txtLoanFromDt.Text)
		strVal = strVal & "&txtLoanToDt=" & Trim(.txtLoanToDt.Text) 
		strVal = strVal & "&txtDocCur=" & Trim(.txtDocCur.Value)
		strVal = strVal & "&txtDueFromDt=" & Trim(.txtDueFromDt.Text)
		strVal = strVal & "&txtDueToDt=" & Trim(.txtDueToDt.Text)
		strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)
		strVal = strVal & "&txtBpCd_Alt=" & Trim(.txtBpCd.Alt)		
		strVal = strVal & "&cboLoanFg=" & Trim(.cboLoanFg.value)		
		strVal = strVal & "&txtLoanType=" & Trim(.txtLoanType.value)
		strVal = strVal & "&txtLoanType_Alt=" & Trim(.txtLoanType.Alt)			
		strVal = strVal & "&txtLoanNo=" & Trim(.txtLoanNo.Value)
'	End If
'--------------- ������ coding part(�������,End)------------------------------------------------
		strVal = strVal & "&txtPgmId=" & Trim("<%=Request("PGM")%>")			 '��: F4101MA1 (reference ���� �߰�)
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
'    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
'	lgIntFlgMode = Parent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtLoanFromDt.focus
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
						<TD CLASS=TD5 NOWRAP>��������</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpLoanFrDt name=txtLoanFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������������"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpLoanToDt name=txtLoanToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������������"></OBJECT>');</SCRIPT></TD>
						<TD CLASS="TD5" NOWRAP>�ŷ���ȭ</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" SIZE = "10" MAXLENGTH="3"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCurCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.value, 0)"></TD>						
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��ȯ��������</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDueFrDt name=txtDueFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="���۸�������"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDueToDt name=txtDueToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="���Ḹ������"></OBJECT>');</SCRIPT></TD>
						<TD CLASS="TD5" NOWRAP>���԰ŷ�ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBpCd" SIZE="10" MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���԰ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankLoanCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value, 5)">
									                       <INPUT TYPE=TEXT NAME="txtBpNm" ALT="���԰ŷ�ó��" SIZE=20 tag="24X"></TD>																		
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>��ܱⱸ��</TD>
						<TD CLASS="TD6" NOWRAP><SELECT NAME="cboLoanFg" ALT="��ܱⱸ��" STYLE="WIDTH: 135px" tag="11X"><OPTION VALUE=""></OPTION></SELECT></TD>
						<TD CLASS=TD5 NOWRAP>���Կ뵵</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoanType" ALT="���Կ뵵" SIZE="10" MAXLENGTH="2"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoanType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanType.value, 6)">&nbsp;<INPUT NAME="txtLoanTypeNm" ALT="���Կ뵵��" STYLE="TEXT-ALIGN: Left" tag="24X"></TD>						
					</TR>
					<TR>
<% IF Request("PGM") = "F4235MA1" THEN %>
						<TD CLASS="TD5" NOWRAP>���⿬���ȣ</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanNo" ALT="���⿬���ȣ" SIZE="20" MAXLENGTH="18"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankLoanCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanNo.value,3)"></TD>
<% ELSE%>
						<TD CLASS="TD5" NOWRAP>���Աݹ�ȣ</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtLoanNo" ALT="���Աݹ�ȣ" SIZE="20" MAXLENGTH="18"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankLoanCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtLoanNo.value,3)"></TD>
<% END IF %>
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
			<IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hBankLoanCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanType" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hLoanToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hDueFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hDueToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPgmId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

