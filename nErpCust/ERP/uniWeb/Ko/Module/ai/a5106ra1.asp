<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1 %>
<!--'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2000/12/09
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Byun Jee Hyun
'* 11. Comment              :
'*                            2000/12/09
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--
========================================================================================================
=                          3.3 Client Side Script
========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs">					</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../ag/AcctCtrl.vbs">							</SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit                                 '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Dim lgBlnFlgChgValue                                        '��: Variable is for Dirty flag            
Dim lgStrPrevKey                                            '��: Next Key tag                          
Dim lgSortKey                                               '��: Sort���� ���庯��                      
Dim lgIsOpenPop                                             '��: Popup status                           
Dim lgPopUpR                                                '��: Orderby default ��                    
Dim lgMark
Dim IsOpenPop                                                  '��: ��ũ                                  


'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "a5106rb1.asp"
Const C_SHEETMAXROWS    = 16                                   '��: Spread sheet���� �������� row
Const C_SHEETMAXROWS_D  = 30                                   '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey			= 1

Dim lsPoNo                 
Dim arrReturn
Dim arrParent
Dim arrParam					

'------ Set Parameters from Parent ASP -----------------------------------------------------------------------

arrParent		= window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)

	
top.document.title = "������ǥ�˾�"

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
'    popupparent.lgAuthorityFlag = arrParam(4)                          '���Ѱ��� �߰� 
    
	Self.Returnvalue = arrReturn
End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= 
Sub SetDefaultVal()
	Dim EndDate

	EndDate = UniConvDateAToB("<%=GetSvrDate%>" ,PopupParent.gServerDateFormat,PopupParent.gDateFormat)

	frm1.txtfrtempgldt.Text	= EndDate
	frm1.txttotempgldt.Text	= EndDate
End Sub

Function OpenPopUp(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrStrRet				'���Ѱ��� �߰�   							  
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrStrRet =  AutorityMakeSql("DEPT",PopupParent.gChangeOrgId, "","","","")	'���Ѱ��� �߰�   							  
	
	arrParam(0) = "�μ� �˾�"				' �˾� ��Ī 
	arrParam(1) = arrstrRet(0)										'���Ѱ��� �߰�   							  				
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = arrstrRet(1)										'���Ѱ��� �߰�   							  
	arrParam(5) = "�μ��ڵ�"				' �����ʵ��� �� ��Ī 

	arrField(0) = "DEPT_CD"	     				' Field��(0)
	arrField(1) = "DEPT_NM"			    		' Field��(1)
    
	arrHeader(0) = "�μ��ڵ�"				' Header��(0)
	arrHeader(1) = "�μ���"					' Header��(1)
			
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet)
	End If	
End Function

Function SetPopUp(Byval arrRet)
	With frm1
		.txtDeptCd.value = arrRet(0)
		.txtDeptNm.value = arrRet(1)
	End With
End Function

'========================================  2.3 LoadInfTB19029()  ==========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","RA") %>
	<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","RA") %>
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

'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
Function MousePointer(pstr1)
    Select Case UCase(pstr1)
        Case "PON"
	  	  window.document.search.style.cursor = "wait"
        Case "POFF"
	  	  window.document.search.style.cursor = ""
    End Select
End Function

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    frm1.vspdData.OperationMode = 3

    Call SetZAdoSpreadSheet("A5106RA1", "S", "A", "V20021108", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
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
' Function Name : OpenSortPopup
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

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
    Call LoadInfTB19029					
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N") 

	Call InitVariables					
	Call SetDefaultVal	
	Call InitSpreadSheet()
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
'   Event Name : txtfrtempgldt
'   Event Desc :
'==========================================================================================
Sub txtfrtempgldt_DblClick(Button)
	If Button = 1 Then
		frm1.txtfrtempgldt.Action = 7
	End if
End Sub

'==========================================================================================
'   Event Name : txttotempgldt
'   Event Desc :
'==========================================================================================
Sub txttotempgldt_DblClick(Button)
	If Button = 1 Then
		frm1.txttotempgldt.Action = 7
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
    
    If frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS_D Then				'��: ������ üũ'
		If lgStrPrevKey <> "" Then											'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Exit Sub
			End If	
		End If
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'======================================================================================================
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
End Sub

'======================================================================================================
'   Event Name : vspdData_KeyPress
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'======================================================================================================
Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

'======================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'======================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.MaxRows > 0 Then
		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

'======================================================================================================
'   Event Name : txtfrtempgldt_Keypress
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'======================================================================================================
Sub txtfrtempgldt_Keypress(KeyAscii)
    On Error Resume Next

    If KeyAscii = 27 Then
		Call CancelClick()
    Elseif KeyAscii = 13 Then
		Call FncQuery()
    End if
End Sub

'======================================================================================================
'   Event Name : txttotempgldt_Keypress
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'======================================================================================================
Sub txttotempgldt_Keypress(KeyAscii)
    On Error Resume Next

    If KeyAscii = 27 Then
		Call CancelClick()
    Elseif KeyAscii = 13 Then
		Call FncQuery()
    End if
End Sub

'==========================================================================================
'   Event Name : txtFrTempGlNo_OnKeyPress
'   Event Desc : 
'==========================================================================================
Sub txtFrTempGlNo_OnKeyPress()	
	If window.event.keycode = 39 then	'Single quotation mark �ԷºҰ� 
		window.event.keycode = 0	
	End If
End Sub

'==========================================================================================
'   Event Name : txtFrTempGlNo_onpaste
'   Event Desc : 
'==========================================================================================
Sub txtFrTempGlNo_onpaste()	
	Dim iStrTempGlNo 	

	iStrTempGlNo = window.clipboardData.getData("Text")
	iStrTempGlNo = RePlace(iStrTempGlNo, "'", "")
	Call window.clipboardData.setData("text",iStrTempGlNo)		
End Sub

'==========================================================================================
'   Event Name : txtToTempGlNo_OnKeyPress
'   Event Desc : 
'==========================================================================================
Sub txtToTempGlNo_OnKeyPress()	
	If window.event.keycode = 39 then	'Single quotation mark �ԷºҰ� 
		window.event.keycode = 0	
	End If
End Sub

'==========================================================================================
'   Event Name : txtToTempGlNo_onpaste
'   Event Desc : 
'==========================================================================================
Sub txtToTempGlNo_onpaste()	
	Dim iStrTempGlNo 	

	iStrTempGlNo = window.clipboardData.getData("Text")
	iStrTempGlNo = RePlace(iStrTempGlNo, "'", "")
	Call window.clipboardData.setData("text",iStrTempGlNo)		
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
    
    If CompareDateByFormat(frm1.txtFrTempGlDt.text,frm1.txtToTempGlDt.text,frm1.txtFrTempGlDt.Alt,frm1.txtToTempGlDt.Alt, _
                        "970025",frm1.txtFrTempGlDt.UserDefinedFormat,PopupParent.gComDateType,True) = False Then				
		Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData()

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery	= False Then														'��: Query db data
		Exit Function
	End If
		
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
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", PopupParent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

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
		strVal = BIZ_PGM_ID & "?txtfrtempgldt=" & Trim(.txtfrtempgldt.Text)
		strVal = strVal & "&txttotempgldt=" & Trim(.txttotempgldt.Text)
		strVal = strVal & "&txtfrtempglno=" & Trim(.txtfrtempglNo.value)
		strVal = strVal & "&txttotempglno=" & Trim(.txttotempglNo.value)
		strVal = strVal & "&txtdeptcd=" & Trim(.txtdeptcd.value)
				
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey							'��: Next key tag
        strVal = strVal & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D)				'��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")         
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")		
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		strVal = strVal & "&lgAuthorityFlag=" & EnCoding(lgAuthorityFlag)            '���Ѱ��� �߰�		
		
        Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()																'��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgBlnFlgChgValue = True															'Indicates that no value changed

	If frm1.vspdData.MaxRows > 0  Then
		frm1.vspdData.focus
	End If
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
						<TD CLASS=TD5 NOWRAP>ȸ������</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/a5106ra1_fpDateTime1_txtfrtempgldt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/a5106ra1_fpDateTime2_txttotempgldt.js'></script>
						</TD>												
						<TD CLASS=TD5 NOWRAP>���ʹ�ȣ</TD>				
						<TD CLASS=TD6 NOWRAP>
						<INPUT TYPE="Text" NAME="txtfrtempglNo" SIZE=12 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1" ALT="���ʹ�ȣ">&nbsp;~&nbsp;
						<INPUT TYPE="Text" NAME="txttotempglNo" SIZE=12 MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="1" ALT="���ʹ�ȣ">
						</TD>
					</TR>
					<TR>				
						<TD CLASS=TD5 NOWRAP>�μ��ڵ�</TD>
						<TD CLASS=TD6 NOWRAP>
						<INPUT NAME="txtDeptCd" ALT="�μ��ڵ�" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtDeptCd.Value)">&nbsp;
						<INPUT NAME="txtDeptNm" ALT="�μ���"   MAXLENGTH="20" SIZE=18 STYLE="TEXT-ALIGN: left" tag="14XXXU"></TD>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP></TD>				
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
						<script language =javascript src='./js/a5106ra1_vspdData_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
									 <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>	
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

