<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : f6101ra1
'*  4. Program Name         : ���ޱݹ�ȣ PopUp
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/04/12
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hersheys
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'*
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
<!-- #Include file="../../inc/IncServer.asp"  -->
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs">					</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "F6102rb1.asp"                              '��: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 3					                          '��: SpreadSheet�� Ű�� ���� 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim  lgIsOpenPop                                          
Dim  lgPopUpR                                              
Dim  lgQueryFlag
Dim  lgCode		

Dim  arrReturn
Dim  arrParent
Dim  arrParam						


Dim  IsOpenPop  
<%
	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	dtToday = GetSvrDate
	Call ExtractDateFrom(dtToday, gServerDateFormat, gServerDateType, strYear, strMonth, strDay)

	EndDate = UNIConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, strDay)
	StartDate = UNIDateAdd("M", -1, EndDate, gDateFormat)
       
%>
' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd	' ����� 
Dim lgInternalCd	' ���κμ� 
Dim lgSubInternalCd	' ���κμ�(��������)
Dim lgAuthUsrID		' ���� 

arrParent        = window.dialogArguments
Set PopupParent = arrParent(0)	 
arrParam		= arrParent(1)
	 '------ Set Parameters from Parent ASP ------ 

	top.document.title = "���ޱ� �˾�"

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()
    Redim arrReturn(0)
    
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
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

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtfrtempgldt.Text	= "<%= StartDate %>"
	frm1.txttotempgldt.Text	= "<%= EndDate %>"
End Sub

'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	Select Case iWhere
		Case 0
			arrParam(0) = "�μ� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_ACCT_DEPT"    				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "ORG_CHANGE_ID = " & FilterVar(gChangeOrgId, "''", "S")	' Where Condition
			arrParam(5) = "�μ��ڵ�"				' �����ʵ��� �� ��Ī 

			arrField(0) = "DEPT_CD"	     				' Field��(0)
			arrField(1) = "DEPT_NM"			    		' Field��(1)
    
			arrHeader(0) = "�μ��ڵ�"				' Header��(0)
			arrHeader(1) = "�μ���"					' Header��(1)
		Case 1
			arrParam(0) = "�ŷ�ó�˾�"						' �˾� ��Ī 
			arrParam(1) = "b_biz_partner"						' TABLE ��Ī 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "�ŷ�ó"			
	
			arrField(0) = "BP_CD"								' Field��(0)
			arrField(1) = "BP_NM"								' Field��(1)
    
    
			arrHeader(0) = "�ŷ�ó"							' Header��(0)
			arrHeader(1) = "�ŷ�ó��"						' Header��(1)
	
	End Select 
	
			
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
			Case 0
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
			Case 1
				.txtDealBpCd.value = arrRet(0)
				.txtDealBpNm.value = arrRet(1)
		End Select
	End With

End Function

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE", "RA") %>                                '��: 
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "RA") %>
End Sub


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'=				  �� �κп��� �÷� �߰��ϰ� ����Ÿ ������ �Ͼ�� �մϴ�.   							=
'========================================================================================================	
Function OKClick()
		
	Dim intColCnt, intRowCnt, intInsRow
		
	If frm1.vspdData.ActiveRow > 0 Then 			
	
		intInsRow = 0

		Redim arrReturn(1)
			
		For intRowCnt = 0 To frm1.vspdData.MaxRows - 1
			
			frm1.vspdData.Row = intRowCnt + 1
			
			If frm1.vspdData.SelModeSelected Then
				frm1.vspdData.Col = 1
				arrReturn(intColCnt) = frm1.vspdData.Text		
				
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
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	frm1.vspdData.OperationMode = 5
    Call SetZAdoSpreadSheet("F6102RA1","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
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
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function OpenOrderBy()

	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	
	On Error Resume Next
	
	ReDim arrParam(C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = "<%=gMethodText%>"    
  
	For ii = 0 to C_MaxSelList * 2 - 1 Step 2
      arrParam(ii + 0 ) = lgPopUpR(ii / 2  , 0)
      arrParam(ii + 1 ) = lgPopUpR(ii / 2  , 1)
    Next  
  
	arrRet = window.showModalDialog("../../ComAsp/ADOGrpSortPopup.asp",Array(lgSortFieldCD,lgSortFieldNm,arrParam,TInf),"dialogWidth=420px; dialogHeight=250px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "0" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	
	   For ii = 0 to C_MaxSelList * 2 - 1 Step 2
           lgPopUpR(ii / 2 ,0) = arrRet(ii + 1)  
           lgPopUpR(ii / 2 ,1) = arrRet(ii + 2)
       Next    
	   
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function


'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'========================================================================================================
Function CookiePage(ByVal Kubun)

End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
			
End Sub
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
	Call LoadInfTB19029()
	
    Call ggoOper.FormatField(Document, "1",PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,PopupParent.ggStrMinPart,PopupParent.ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   

	Call InitVariables()
	Call SetDefaultVal()
	Call InitSpreadSheet()
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

Sub txtfrtempgldt_DblClick(Button)
	if Button = 1 then
		frm1.txtfrtempgldt.Action = 7
	End if
End Sub

Sub txttotempgldt_DblClick(Button)
	if Button = 1 then
		frm1.txttotempgldt.Action = 7
	End if
End Sub

'==========================================================================================
'   Event Name : KeyPress
'   Event Desc :
'==========================================================================================
Sub txtfrtempgldt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		Call FncQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

Sub txttotempgldt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		Call FncQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

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
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If
    
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
			ggoSpread.SSSort Col 
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col,lgSortKey 
			lgSortKey = 1
		End If 
    End If
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
End Sub

Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function


Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Then              		' Title cell�� dblclick�߰ų�....
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows = 0 Then  	'NO Data
		Exit Sub
	End If
	Call OKClick
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

'********************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery() 
	FncQuery = False                                            
    
    Err.Clear                                                   

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables 											
	frm1.vspdData.MaxRows = 0                                                   '��: Protect system from crashing    
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If

	If CompareDateByFormat(frm1.txtFrTempGlDt.text,frm1.txtToTempGlDt.text,frm1.txtFrTempGlDt.Alt,frm1.txtToTempGlDt.Alt, _
        	               "970025",frm1.txtFrTempGlDt.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   frm1.txtFrTempGlDt.focus
	   Exit Function
	End If
	
	If frm1.txtDealBpCd.value = "" then
		frm1.txtDealBpNm.value=""
    End if
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
	Call parent.FncExport(C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO, "X", "X")
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

    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1

        strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> OPMD_UMODE Then   ' This means that it is first search
			strVal = strVal & "?txtfrtempgldt=" & Trim(.txtfrtempgldt.Text)
			strVal = strVal & "&txttotempgldt=" & Trim(.txttotempgldt.Text)
			strVal = strVal & "&txtfrtempglno=" & Trim(.txtfrtempglNo.value)
			strVal = strVal & "&txttotempglno=" & Trim(.txttotempglNo.value)
			strVal = strVal & "&txtdeptcd="		& Trim(.txtdeptcd.value)
			strVal = strVal & "&txtDealBpCd="	& Trim(.txtDealBpCd.value)
        Else
			
			strVal = strVal & "?txtfrtempgldt=" & Trim(.htxtfrtempgldt.value)
			strVal = strVal & "&txttotempgldt=" & Trim(.htxttotempgldt.value)			
			strVal = strVal & "&txtfrtempglno=" & Trim(.htxtfrtempglNo.value)
			strVal = strVal & "&txttotempglno=" & Trim(.htxttotempglNo.value)
			strVal = strVal & "&txtdeptcd="		& Trim(.htxtdeptcd.value)
			strVal = strVal & "&txtDealBpCd="	& Trim(.htxtDealBpCd.value)
        End If   
         
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
        strVal = strVal & "&lgSelectListDT=" & lgSelectListDT
         
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList(lgMaxFieldCount,lgPopUpR,gFieldCD,gNextSeq,gTypeCD(0),C_MaxSelList)
		strVal = strVal & "&lgSelectList="   & EnCoding(lgSelectList)

'--------------- ������ coding part(�������,Start)----------------------------------------------
		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
	
        Call RunMyBizASP(MyBizASP, strVal)							

    End With
    
    DbQuery = True

End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()

	lgBlnFlgChgValue = False
    lgIntFlgMode     = OPMD_UMODE												'��: Indicates that current mode is Update mode
    lgSaveRow        = 1
  
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	End If
	
End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>
<!--
'#########################################################################################################
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
						<TD CLASS=TD5 NOWRAP>�������</TD>
						<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtfrtempgldt CLASSID=<%=gCLSIDFPDT%> ALT="�����������" tag="12"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txttotempgldt CLASSID=<%=gCLSIDFPDT%> ALT="�����������" tag="12"></OBJECT>');</SCRIPT>
						</TD>
						<TD CLASS=TD5 NOWRAP>���ޱݹ�ȣ</TD>				
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="Text" NAME="txtfrtempglNo" SIZE=15 MAXLENGTH=20 tag="1XXXXU" ALT="���ۼ��ޱݹ�ȣ">&nbsp;~
							<INPUT TYPE="Text" NAME="txttotempglNo" SIZE=15 MAXLENGTH=20 tag="1XXXXU" ALT="���ἱ�ޱݹ�ȣ">
						</TD>
					</TR>
					<TR>				
						<TD CLASS=TD5 NOWRAP>�μ��ڵ�</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtDeptCd" ALT="�μ��ڵ�" MAXLENGTH="10" SIZE=10 tag ="11XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtDeptCd.Value, 0)">
							<INPUT NAME="txtDeptNm" ALT="�μ���"   MAXLENGTH="20" SIZE=18 tag ="14XXXU">
						</TD>
						<TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
						<TD CLASS=TD6 NOWRAP>
						<INPUT NAME="txtDealBpCd" ALT="�ŷ�ó�ڵ�" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU"><IMG align=top name=btnCalType onclick="vbscript:CALL OpenPopUp(frm1.txtDealBpCd.Value, 1)" src="../../image/btnPopup.gif"  TYPE="BUTTON">
						<INPUT NAME="txtDealBpNm" ALT="�ŷ�ó��" SIZE="20" tag = "14X" ></TD>
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2" id=vspdData> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"><PARAM NAME="ReDraw" VALUE="0"><PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
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
					<TD WIDTH=70% NOWRAP><IMG SRC="../../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search"  ONCLICK="FncQuery()"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1"  ONCLICK="OkClick()"></IMG>
							                  <IMG SRC="../../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"    ONCLICK="CancelClick()"></IMG></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtfrtempgldt"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxttotempgldt"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtfrtempglNo"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxttotempglNo"        tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtdeptcd"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtDealBpCd"			tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
