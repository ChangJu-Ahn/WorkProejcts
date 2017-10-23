<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2108ra1
'*  4. Program Name         : ���������˾� 
'*  5. Program Desc         : Popup of Budget
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.03.31
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : 
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
'*********************************************************************************************************** -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
 '==========================================  1.2.3 Global Variable�� ����  ===============================
 <!-- #Include file="../../inc/lgvariables.inc" -->	
'=========================================================================================================  
Dim lgIsOpenPop                                             '��: Popup status                           
Dim lgKeyPos                                                '��: Key��ġ                               
Dim lgKeyPosVal                                             '��: Key��ġ Value                         

Dim lgPopUpR                                                '��: Orderby default ��                    
Dim lgMark

Dim IsOpenPop                                                  '��: ��ũ                                  
Dim strFrDt
Dim strToDt

Dim lgStrComDateType		'Company Date Type�� ����(��� Mask�� �����.)
'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "f2108rb1.asp"

Const C_SHEETMAXROWS_D  = 100                                  '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 12                                    '�١١١�: Max key value

Dim arrReturn
Dim arrParent
Dim arrParam

	' ���Ѱ��� �߰� 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
	Dim lgAuthUsrID, lgAuthUsrNm					' ���� 


	 '------ Set Parameters from Parent ASP ------ 
	arrParent		= Window.DialogArguments
	Set PopupParent	= arrParent(0)	 
	arrParam		= arrParent(1)
	
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
    lgIntFlgMode     = PopupParent.OPMD_CMODE
    
	Self.Returnvalue = arrReturn

	' ���Ѱ��� �߰� 
	If UBound(arrParam) > 7 Then
		lgAuthBizAreaCd	= arrParam(7)
		lgInternalCd	= arrParam(8)
		lgSubInternalCd	= arrParam(9)
		lgAuthUsrID		= arrParam(10)    
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
	strSvrDate = "<%=GetSvrDate%>"

	frm1.hChgFg.value		= arrParam(0)
	frm1.txtDeptCd.value	= arrParam(3)
	frm1.txtBdgCdFr.value	= arrParam(4)
	frm1.txtBdgCdTo.value	= arrParam(5)
	frm1.hOrgChangeId.value = arrParam(6)

	frm1.txtBdgYymmFr.Text = UniConvDateAToB(strSvrDate ,popupparent.gServerDateFormat,popupparent.gDateFormat) 
	frm1.txtBdgYymmTo.Text = UniConvDateAToB(strSvrDate ,popupparent.gServerDateFormat,popupparent.gDateFormat) 
    Call ggoOper.FormatDate(frm1.txtBdgYymmFr, popupparent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtBdgYymmTo, popupparent.gDateFormat, 2)

	'Call ggoOper.FormatDate(frm1.txtBdgYymmFr, popupparent.gDateFormat, 2)
    'Call ggoOper.FormatDate(frm1.txtBdgYymmTo, popupparent.gDateFormat, 2)
	'frm1.txtBdgYymmFr.Text = UNIMonthClientFormat("<%=GetSvrDate%>")
	'frm1.txtBdgYymmTo.Text = UNIMonthClientFormat("<%=GetSvrDate%>")

End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","RA") %>                                '��: 
End Sub


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  �� �κп��� �÷� �߰��ϰ� ����Ÿ ������ �Ͼ�� �մϴ�.   							=
'========================================================================================================


Function OKClick()
		
	Dim intColCnt, intRowCnt, intInsRow
		
	if frm1.vspdData.ActiveRow > 0 Then 			
		
		intInsRow = 0

		Redim arrReturn(9)
			
		For intRowCnt = 0 To frm1.vspdData.MaxRows - 1
			
			frm1.vspdData.Row = intRowCnt + 1
		
			If frm1.vspdData.SelModeSelected Then
				frm1.vspdData.row	= frm1.vspdData.ActiveRow
				frm1.vspdData.Col	= GetKeyPos("A",1)	'�����ڵ� 
				arrReturn(0)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",2)	'������ 
				arrReturn(1)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",3)	'�μ��ڵ� 
				arrReturn(2)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",4)	'��������ID	
				arrReturn(3)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",5)	'�μ��� 
				arrReturn(4)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",6)	'��������ID
				arrReturn(5)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",9)	'�����Ⱓ���� 
				arrReturn(6)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",10)	'�߰� 
				arrReturn(7)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",11)	'�̿� 
				arrReturn(8)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",12)	'���� 
				arrReturn(9)		= frm1.vspdData.Text
				intInsRow = intInsRow + 1
			End IF
		Next
		
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
'========================================================================================================
' Name : 
' Desc : 
'========================================================================================================
Sub ConditionKeypress()
	If window.event.keyCode = 13 Then
		Call FncQuery()
	Elseif window.event.keyCode = 27 Then		
       Call CancelClick()
	End If
End sub
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
	frm1.vspdData.OperationMode = 5
	Call SetZAdoSpreadSheet("F2108RA1","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock() 
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLock 1 , -1
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub

'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)
	Dim tmpBdgYymmddFr, tmpBdgYymmddTo
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True	

	tmpBdgYymmddFr	=  UniConvDateAToB(frm1.txtBdgYymmFr,popupparent.gDateFormatYYYYMM,popupparent.gServerDateFormat)
	tmpBdgYymmddTo	=  UniConvDateAToB(frm1.txtBdgYymmTo,popupparent.gDateFormatYYYYMM,popupparent.gServerDateFormat)
	tmpBdgYymmddTo	=  UNIDateAdd("M", +1, tmpBdgYymmddTo,popupparent.gServerDateFormat)
	tmpBdgYymmddTo	=  UNIDateAdd("D", -1, tmpBdgYymmddTo,popupparent.gServerDateFormat)	    
	
	'Company Date Type ���� ���� 
	tmpBdgYymmddFr  =  UniConvDateAToB(tmpBdgYymmddFr,popupparent.gServerDateFormat,gDateFormat)
	tmpBdgYymmddTo =  UniConvDateAToB(tmpBdgYymmddTo,popupparent.gServerDateFormat,gDateFormat)

	arrParam(0) = tmpBdgYymmddFr				
   	arrParam(1) = tmpBdgYymmddTo
	arrParam(2) = lgUsrIntCd                           ' �ڷ���� Condition  
	arrParam(3) = frm1.txtDeptCd.value				
	arrParam(4) = "F"										' �������� ���� Condition  
	
	' ���Ѱ��� �߰� 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(popupparent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
End Function

'------------------------------------------  SetDept()  --------------------------------------------------
'	Name : SetDept()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetDept(Byval arrRet)	
		
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.hOrgChangeId.value=arrRet(2)

		frm1.txtDeptCd.focus		
End Function


 '**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	Select Case iWhere

		Case "BdgCdFr", "BdgCdTo"
			arrParam(0) = "�����ڵ� �˾�"								' �˾� ��Ī 
			arrParam(1) = "F_BDG_ACCT A "									' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""
			arrParam(5) = "�����ڵ�"									' �����ʵ��� �� ��Ī 

			arrField(0) = "A.BDG_CD"	     								' Field��(0)
			arrField(1) = "A.GP_ACCT_NM"			    					' Field��(1)
			
			arrHeader(0) = "�����ڵ�"									' Header��(0)
			arrHeader(1) = "�����"										' Header��(1)
			
		Case Else
			Exit Function
	End Select	
	
	lgIsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		    
		    Case "BdgCdFr"
				.txtBdgCdFr.value = arrRet(0)
				.txtBdgNmFr.value = arrRet(1)
				.txtBdgCdFr.focus				
		    Case "BdgCdTo"
				.txtBdgCdTo.value = arrRet(0)
				.txtBdgNmTo.value = arrRet(1)
				.txtBdgCdTo.focus				
			
		End Select
    
	End With

End Function

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
    
	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	'Call ggoOper.FormatDate(frm1.txtBdgYymmFr, PopupParent.gDateFormat, 2)
    'Call ggoOper.FormatDate(frm1.txtBdgYymmTo, PopupParent.gDateFormat, 2)
	Call InitSpreadSheet()
	    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	frm1.txtBdgYymmFr.focus
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
Sub txtBdgYymmFr_DblClick(Button)
	if Button = 1 then
		frm1.txtBdgYymmFr.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBdgYymmFr.Focus
	End if
End Sub

Sub txtBdgYymmTo_DblClick(Button)
	if Button = 1 then
		frm1.txtBdgYymmTo.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBdgYymmTo.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : KeyPress
'   Event Desc :
'==========================================================================================
Sub txtBdgYymmFr_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 Then 	
		frm1.txtBdgYymmTo.focus	
		Call Fncquery()		
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

Sub txtBdgYymmTo_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		frm1.txtBdgYymmFr.focus
		Call Fncquery()
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
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'��: ������ üũ'
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DbQuery
		End If
   End if

End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
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
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	If Row < 1 Then Exit Sub

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
Dim strFrYear, strFrMonth, strFrDay
Dim strToYear, strToMonth, strToDay
	
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
    	
    If CompareDateByFormat(frm1.txtBdgYymmFr.Text, frm1.txtBdgYymmTo.Text, frm1.txtBdgYymmFr.Alt, frm1.txtBdgYymmTo.Alt, _
						"970025", frm1.txtBdgYymmFr.UserDefinedFormat, PopupParent.gComDateType, true) = False Then
			frm1.txtBdgYymmFr.focus														'��: GL Date Compare Common Function
			Exit Function
	End if

    Call ExtractDateFrom(frm1.txtBdgYymmFr.Text,frm1.txtBdgYymmFr.UserDefinedFormat,PopupParent.gComDateType,strFrYear,strFrMonth,strFrDay)
    strFrDt = strFrYear & strFrMonth

    Call ExtractDateFrom(frm1.txtBdgYymmTo.Text,frm1.txtBdgYymmTo.UserDefinedFormat,PopupParent.gComDateType,strToYear,strToMonth,strToDay)
    strToDt = strToYear & strToMonth
	
    frm1.txtBdgCdFr.value = Trim(frm1.txtBdgCdFr.value)
    frm1.txtBdgCdTo.value = Trim(frm1.txtBdgCdTo.value)
    
    If frm1.txtBdgCdFr.value <> "" And frm1.txtBdgCdTo.value <> "" Then
		If frm1.txtBdgCdFr.value > frm1.txtBdgCdTo.value Then
			Call DisplayMsgBox("970025", "X", frm1.txtBdgCdFr.Alt, frm1.txtBdgCdTo.Alt)
			frm1.txtBdgCdFr.focus 
			Exit Function
		End If
    End If
	
	IF NOT CheckOrgChangeId Then
		  IntRetCD = DisplayMsgBox("800600","X",frm1.txtBdgYymmFr.alt,"X")            '��: Display Message(There is no changed data.)
		Exit Function
	End if
	
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
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call parent.FncExport(PopupParent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(PopupParent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
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
	Dim ColYymm1, ii
    Dim TempGetSqlSelectListA
    DbQuery = False
    Err.Clear           

	Call LayerShowHide(1)

	TempGetSqlSelectListA    = Split(EnCoding(GetSQLSelectList("A")),",")
          For ii = LBound(TempGetSqlSelectListA) To UBound(TempGetSqlSelectListA)
              If TempGetSqlSelectListA(ii) = "A.BDG_YYYYMM" Then 
                  ColYymm1 = ii	'������ �÷� 
                  Exit For
               End If
          Next

    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtBdgYymmFr="	& Trim(.hBdgYymmFr.value)
		strVal = strVal & "&txtBdgYymmTo="		& Trim(.hBdgYymmTo.value)
		strVal = strVal & "&txtDeptCd="			& Trim(.hDeptCd.value)
		strVal = strVal & "&txtBdgCdFr="		& Trim(.hBdgCdFr.value)
		strVal = strVal & "&txtBdgCdTo="		& Trim(.hBdgCdTo.value)
		strVal = strVal & "&txtChgFg="			& Trim(.hChgFg.value)
		strVal = strVal & "&txtDeptCd_Alt="		& Trim(.txtDeptCd.Alt)
		strVal = strVal & "&txtBdgCdFr_Alt="	& Trim(.txtBdgCdFr.Alt)
		strVal = strVal & "&txtBdgCdTo_Alt="	& Trim(.txtBdgCdTo.Alt)
		strVal = strVal & "&txtColYymm="		& ColYymm1
		strVal = strVal & "&txtDateType="		& PopupParent.gComDateType
	Else
		strVal = BIZ_PGM_ID & "?txtBdgYymmFr="	& strFrDt
		strVal = strVal & "&txtBdgYymmTo="		& strToDt
		strVal = strVal & "&txtDeptCd="			& Trim(.txtDeptCd.value)
		strVal = strVal & "&txtBdgCdFr="		& Trim(.txtBdgCdFr.value)
		strVal = strVal & "&txtBdgCdTo="		& Trim(.txtBdgCdTo.value)
		strVal = strVal & "&txtChgFg="			& Trim(.hChgFg.value)
		strVal = strVal & "&txtDeptCd_Alt="		& Trim(.txtDeptCd.Alt)
		strVal = strVal & "&txtBdgCdFr_Alt="	& Trim(.txtBdgCdFr.Alt)
		strVal = strVal & "&txtBdgCdTo_Alt="	& Trim(.txtBdgCdTo.Alt)
		strVal = strVal & "&txtColYymm="		& ColYymm1
		strVal = strVal & "&txtDateType="		& PopupParent.gComDateType
	End If
	
'--------------- ������ coding part(�������,End)------------------------------------------------
	strVal = strVal & "&lgStrPrevKey="		& lgStrPrevKey                      '��: Next key tag
	strVal = strVal & "&lgMaxCount="		& CStr(C_SHEETMAXROWS_D)            '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")
	strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
	strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))


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
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	lgIntFlgMode = PopupParent.OPMD_UMODE
	
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
'==========================================================================================
'   Event Name : CheckOrgChangeId
'   Event Desc : 
'==========================================================================================
Function CheckOrgChangeId()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
	Dim tmpBdgYymmddFr, tmpBdgYymmddTo	
 
	tmpBdgYymmddFr = UniConvDateAToB(frm1.txtBdgYymmFr,popupparent.gDateFormatYYYYMM,popupparent.gServerDateFormat)
	tmpBdgYymmddTo =  UniConvDateAToB(frm1.txtBdgYymmTo,popupparent.gDateFormatYYYYMM,popupparent.gServerDateFormat)
	tmpBdgYymmddTo =  UNIDateAdd("M", +1, tmpBdgYymmddTo,popupparent.gServerDateFormat)
	tmpBdgYymmddTo =  UNIDateAdd("D", -1, tmpBdgYymmddTo,popupparent.gServerDateFormat)	
	
	CheckOrgChangeId = True
 
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" Then
			'----------------------------------------------------------------------------------------
			strSelect = "Distinct ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(tmpBdgYymmddFr , "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(tmpBdgYymmddTo , "''", "S") & ") "
			strWhere = strWhere & " AND ORG_CHANGE_ID =  " & FilterVar(.hOrgChangeId.value , "''", "S") & ""
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")

			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)			
								
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					.txtDeptCd.focus
					CheckOrgChangeId = False
			End if
		End If
	End With
		'----------------------------------------------------------------------------------------

End Function

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
						<TD CLASS="TD5" NOWRAP>������</TD>
						<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtBdgYymmFr" CLASS=FPDTYYYYMM tag="12" Title="FPDATETIME" ALT=���ۿ����� id=fpBdgYymmFr></OBJECT>');</SCRIPT>&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtBdgYymmTo" CLASS=FPDTYYYYMM tag="12" Title="FPDATETIME" ALT=���Ό���� id=fpBdgYymmTo></OBJECT>');</SCRIPT></TD>
						<TD CLASS="TD5" NOWRAP>�μ��ڵ�</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDeptCd" MAXLENGTH="10" SIZE=10  ALT ="�μ��ڵ�" tag="11XXXU" onkeypress="ConditionKeypress"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">&nbsp;<INPUT NAME="txtDeptNm" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="�μ���" tag="24X"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>���ۿ����ڵ�</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBdgCdFr" MAXLENGTH="18" SIZE=10  ALT ="���ۿ����ڵ�" tag="11XXXU" onkeypress="ConditionKeypress"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBdgCdFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(txtBdgCdFr.Value, 'BdgCdFr')">&nbsp;<INPUT NAME="txtBdgNmFr" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="�����ڵ��" tag="24X"></TD>
						<TD CLASS="TD5" NOWRAP>���Ό���ڵ�</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBdgCdTo" MAXLENGTH="18" SIZE=10  ALT ="���Ό���ڵ�" tag="11XXXU" onkeypress="ConditionKeypress"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBdgCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(txtBdgCdTo.Value, 'BdgCdTo')">&nbsp;<INPUT NAME="txtBdgNmTo" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="�����ڵ��" tag="24X"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hBdgYymmFr" tag="24">
<INPUT TYPE=HIDDEN NAME="hBdgYymmTo" tag="24">
<INPUT TYPE=HIDDEN NAME="hDeptCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hBdgCdFr" tag="24">
<INPUT TYPE=HIDDEN NAME="hBdgCdTo" tag="24">
<INPUT TYPE=HIDDEN NAME="hChgFg" tag="14">
<INPUT TYPE=hidden NAME="hOrgChangeId"	tag="14" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

