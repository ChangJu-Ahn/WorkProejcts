 <%@ LANGUAGE="VBSCRIPT" %>

<!--'**********************************************************************************************
'*  1. Module Name          : �ڱݰ��� 
'*  2. Function Name        : �������� 
'*  3. Program ID           : F5107RA1.ASP
'*  4. Program Name         : Note Reference Popup
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001/04/17
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : Kim Hee Jung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2001/02/07
'********************************************************************************************** -->
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

Dim lgIsOpenPop                                             '��: Popup status                           

Dim lgPopUpR                                                '��: Orderby default ��                    
Dim lgMark

Dim IsOpenPop                                                  '��: ��ũ                                  
CONST C_MaxKey = 1

<!-- #Include file="../../inc/lgvariables.inc" -->
'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "f5107rb1.asp"
Const C_SHEETMAXROWS_D  = 30                                   '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 

Dim arrReturn
Dim arrParent
Dim arrParam					

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm

	 '------ Set Parameters from Parent ASP ------ 
	arrParent = window.dialogArguments
	Set PopupParent = arrParent(0)
	arrParam = arrParent(1)
	
	top.document.title = "���������˾�"

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
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
    
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

'--------------- ������ coding part(�������,Start)--------------------------------------------------
	Dim strSvrDate
	Dim frDt, toDt
	Redim arrParam(5)
	strSvrDate = "<%=GetSvrDate%>"
	frDt = UNIDateAdd("M", -1, strSvrDate,Popupparent.gServerDateFormat)
		
	frm1.txtFrPrDt.Text = UniConvDateAToB(frDt ,Popupparent.gServerDateFormat,Popupparent.gDateFormat) 
	frm1.txtToPrDt.Text = UniConvDateAToB("<%=GetSvrDate%>" ,Popupparent.gServerDateFormat,Popupparent.gDateFormat) 
	
	If arrParam(0) <> "" Then	'������From
		frm1.txtFrPrDt.Text = UniConvDateAToB(arrParam(0),Popupparent.gServerDateFormat,Popupparent.gDateFormat) 
	End If
	If arrParam(1) <> "" Then	'������To
		frm1.txtToPrDt.Text = UniConvDateAToB(arrParam(1),Popupparent.gServerDateFormat,Popupparent.gDateFormat) 
	End If
	If arrParam(2) <> "" Then	'������ȣFrom
		frm1.txtFrPrNo.value = arrParam(2)
	End If
	If arrParam(3) <> "" Then	'������ȣTo
		frm1.txtToPrNo.value = arrParam(3)
	End If
	If arrParam(4) <> "" Then	'�ŷ�ó�ڵ� 
		frm1.txtBpCd.value = arrParam(4)
	End If
	
'--------------- ������ coding part(�������,End)----------------------------------------------------

End Sub
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
		Call SetPopUp(iWhere,arrRet)
	End If	
End Function



Function OpenPopUp(Byval PopFg,strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	if PopFg = "BP" Then
	
			arrParam(0) = "�ŷ�ó �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_BIZ_PARTNER"    			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "�ŷ�ó�ڵ�"					' �����ʵ��� �� ��Ī 

			arrField(0) = "BP_CD"	     				' Field��(0)
			arrField(1) = "BP_NM"			    		' Field��(1)
    
			arrHeader(0) = "�ŷ�ó�ڵ�"					' Header��(0)
			arrHeader(1) = "�ŷ�ó��"				' Header��(1)
	
	else
			arrParam(0) = "���������˾�"				' �˾� ��Ī 
			arrParam(1) = "F_NOTE"    			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "������ȣ"					' �����ʵ��� �� ��Ī 

			arrField(0) = "NOTE_NO"	     				' Field��(0)
    
			arrHeader(0) = "������ȣ"					' Header��(0)

	end if
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	     "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		With frm1
		select case PopFg
			case "BP"
				.txtBpCd.focus
			case "FrPr"
				.txtFrPrNo.focus
			case "ToPr"	
				.txtToPrNo.focus
			end select 
		End With	
		Exit Function
	Else
		Call SetPopUp(PopFg,arrRet)
	End If	

End Function

Function SetPopUp(Byval PopupFg,Byval arrRet)
	
	With frm1
	select case PopupFg
		case "BP"
			.txtBpCd.value	 = arrRet(0)
			.txtBpNm.value	 = arrRet(1)
			.txtBpCd.focus
		case "FrPr"
			.txtFrPrNo.value = arrRet(0)			
			.txtFrPrNo.focus
		case "ToPr"	
			.txtToPrNo.value = arrRet(0)			
			.txtToPrNo.focus
		end select 
	End With

End Function


'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->	
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE", "RA") %> 
End Sub



'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  �� �κп��� �÷� �߰��ϰ� ����Ÿ ������ �Ͼ�� �մϴ�.   							=
'========================================================================================================
	
Function OKClick()
		
		If frm1.vspdData.ActiveRow > 0 Then 				
		Redim arrReturn(1)
		frm1.vspdData.row	= frm1.vspdData.ActiveRow
		frm1.vspdData.Col	= GetKeyPos("A",1)		
		arrReturn(0)		= frm1.vspdData.Text
	End if			
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

'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================

	Function MousePointer(pstr1)
	      Select case UCase(pstr1)
	            case "PON"
					window.document.search.style.cursor = "wait"
	            case "POFF"
					window.document.search.style.cursor = ""
	      End Select
	End Function


'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
	frm1.vspddata.OperationMode = 3
    Call SetZAdoSpreadSheet("F5107RA1","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock() 
End Sub
'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	
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

'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
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


 '++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'==================================================================================================== 

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
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Popupparent.gDateFormat,Popupparent.gComNum1000,Popupparent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
'--------------- ������ coding part(�������,Start)----------------------------------------------------
   
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
Sub txtFrPrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrPrDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtFrPrDt.Focus
	End if
End Sub

Sub txtToPrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPrDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToPrDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : KeyPress
'   Event Desc :
'==========================================================================================
Sub txtFrPrDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtToPrDt.focus
		Call FncQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

Sub txtToPrDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtFrPrDt.focus
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
    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + PopupParent.VisibleRowCnt(frm1.vspdData,NewTop) Then	'��: ������ üũ'
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			DbQuery
		End If
   End if
    
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
    
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)  
'--------------- ������ coding part(�������,End)------------------------------------------------------
    
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
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
	
	If (frm1.txtFrPrDt.Text <> "") And (frm1.txtToPrDt.Text <> "") Then
		If Popupparent.CompareDateByFormat(frm1.txtFrPrDt.Text, frm1.txtToPrDt.Text, frm1.txtFrPrDt.Alt, frm1.txtToPrDt.Alt, _
					"970025", frm1.txtFrPrDt.UserDefinedFormat, Popupparent.gComDateType, true) = False Then
			frm1.txtFrPrDt.focus											
			Exit Function
		End if	
	End If
	
    '-----------------------
    'Query function call area
    '-----------------------
    frm1.vspdData.MaxRows = 0                                                   '��: Protect system from crashing
    Call DbQuery															'��: Query db data

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
	Call parent.FncExport(Popupparent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Popupparent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
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

	Call LayerShowHide(1)

    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtFrPrDt=" & Trim(.txtFrPrDt.Text)
		strVal = strVal & "&txtToPrDt=" & Trim(.txtToPrDt.Text)
		strVal = strVal & "&txtFrPrNo=" & Trim(.txtFrPrNo.value)
		strVal = strVal & "&txtToPrNo=" & Trim(.txtToPrNo.value)
		strVal = strVal & "&txtBpCd="   & Trim(.txtBpCd.value)
		'strVal = strVal & "&cboNoteFg=" & .cboNoteFg.value		

'--------------- ������ coding part(�������,End)------------------------------------------------
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag
        strVal = strVal & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D)            '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
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
						<TD CLASS=TD5 NOWRAP>������</TD>
						<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtFrPrDt CLASSID=<%=gCLSIDFPDT%> ALT="���۹�����" tag="12"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtToPrDt CLASSID=<%=gCLSIDFPDT%> ALT="���������" tag="12"></OBJECT>');</SCRIPT>
						</TD>
						<TD CLASS=TD5 NOWRAP>������ȣ</TD>				
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="Text" NAME="txtFrPrNo" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="11XXXU" ALT="���۾�����ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('FrPr',frm1.txtFrPrNo.Value)">&nbsp;~
							<INPUT TYPE="Text" NAME="txtToPrNo" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="11XXXU" ALT="���������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('ToPr',frm1.txtToPrNo.Value)">
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>�ŷ�ó�ڵ�</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtBpCd" ALT="�ŷ�ó�ڵ�" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.Value,'BP')">
							<INPUT NAME="txtBpNm" ALT="�ŷ�ó��" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="14X">
						</TD>
						<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
						<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
	<!--				<TD CLASS=TD5>��������</TD>
						<TD CLASS=TD6>
							<SELECT ID="cboNoteFg" NAME="cboNoteFg" ALT="��������" STYLE="WIDTH: 132px" tag="12X"></SELECT>
						</TD>
	-->
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2" id=vspdData> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"><PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

