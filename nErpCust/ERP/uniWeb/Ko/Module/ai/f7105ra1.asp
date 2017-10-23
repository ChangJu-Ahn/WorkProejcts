
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Template
*  2. Function Name        : 
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         :  Ado query Sample with DBAgent(Sort)
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2001/04/18
*  9. Modifier (First)     :
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
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
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>

Const BIZ_PGM_ID 		= "f7105rb1.asp"                              '��: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 5					                          '��: SpreadSheet�� Ű�� ���� 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          

Dim lgSelectList                                         
Dim lgSelectListDT                                       


Dim lgSortFieldNm                                        
Dim lgSortFieldCD                                         

Dim lgMaxFieldCount

Dim lgPopUpR                                              

Dim lgKeyPos                                              
Dim lgKeyPosVal                                         
Dim lgCookValue 

Dim lgSaveRow 

Dim arrReturn
Dim arrParent
Dim arrParam					


	 '------ Set Parameters from Parent ASP ------ 
arrParent        = window.dialogArguments
Set PopupParent = arrParent(0)	 
arrParam		= arrParent(1)


	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	dtToday = "<%=GetSvrDate%>"
	Call PopupParent.ExtractDateFrom(dtToday, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)

	EndDate = PopupParent.UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)
	StartDate = PopupParent.UNIDateAdd("M", -1, EndDate, PopupParent.gDateFormat)

'--------------- ������ coding part(�������,End)-------------------------------------------------------------
top.document.title = PopupParent.gActivePRAspName
	'top.document.title = "�����������˾�"

'--------------- ������ coding part(��������,End)-------------------------------------------------------------

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
    lgSaveRow        = 0
    
	Self.Returnvalue = arrReturn
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtFrPrDt.Text	= StartDate
	frm1.txtToPrDt.Text	= EndDate

End Sub

Function OpenPopUp(Byval PopFg,strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
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
			arrParam(0) = "������ �˾�"				' �˾� ��Ī 
			arrParam(1) = "F_PRRCPT"    			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "�����ݹ�ȣ"					' �����ʵ��� �� ��Ī 

			arrField(0) =  "PRRCPT_NO"	     				' Field��(0)
			arrField(1) =  "DD" & PopupParent.gColSep & "PRRCPT_DT"	
			arrField(2) =  "F2" & PopupParent.gColSep & "PRRCPT_AMT"	
			arrField(3) =  "F2" & PopupParent.gColSep & "LOC_PRRCPT_AMT"	
			
			arrHeader(0) = "�����ݹ�ȣ"					' Header��(0)
			arrHeader(1) = "�Ա�����"					' Header��(0)
			arrHeader(2) = "�ݾ�"					' Header��(0)
			arrHeader(3) = "�ݾ�(�ڱ�)"					' Header��(0)

	end if
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	     "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
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
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "A","NOCOOKIE", "RA") %>                                '��: 
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "RA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    frm1.vspdData.OperationMode = 3
    
    Call SetZAdoSpreadSheet("A7101RA1","S","A","V20021211",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
	End With
	
    Call SetSpreadLock() 
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	
    With frm1
		.vspdData.ReDraw = False
		 ggoSpread.Source = frm1.vspdData
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
Function OpenOrderByPopup()

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

    Call LoadInfTB19029														
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call ggoOper.FormatField(Document, "1",PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,PopupParent.ggStrMinPart,PopupParent.ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
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
		Frm1.txtFrPrDt.Focus		
	End if
End Sub

Sub txtToPrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPrDt.Action = 7
		Call SetFocusToDocument("P")
		Frm1.txtToPrDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : KeyPress
'   Event Desc :
'==========================================================================================
Sub txtFrPrDt_KeyPress(KeyAscii)
	On Error Resume Next

	If KeyAscii = 13 Then 
		Call FncQuery
	ElseIf KeyAscii = 27 Then
		Call CancelClick
	End If
End Sub

Sub txtToPrDt_KeyPress(KeyAscii)
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
    
	If frm1.vspdData.MaxRows < NewTop + PopupParent.VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           If DbQuery = False Then
              Exit Sub
           End if
    	End If
    End If
    
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
	'lgCookValue = ""
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
'	If Frm1.vspdData.MaxRows > 0 Then
'		If Frm1.vspdData.ActiveRow = Row Or Frm1.vspdData.ActiveRow > 0 Then
'			Call OKClick
'		End If
'	End If
End Sub

 
'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 

    FncQuery = False                                            
    
    Err.Clear                                                   

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables 											
	ggoSpread.Source= frm1.vspdData
	ggoSpread.ClearSpreadData
	
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If
    
    If PopupParent.CompareDateByFormat(frm1.fpDateTime1.text,frm1.fpDateTime2.text,frm1.fpDateTime1.Alt,frm1.fpDateTime2.Alt, _
     	               "970025",frm1.fpDateTime1.UserDefinedFormat,PopupParent.gComDateType, true) = False Then
	   frm1.fpDateTime1.focus
		Exit Function
	End if		

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
'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1

        strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> PopupParent.OPMD_UMODE Then   ' This means that it is first search
			strVal = strVal & "?txtFrPrDt=" & Trim(.txtFrPrDt.Text)
			strVal = strVal & "&txtToPrDt=" & Trim(.txtToPrDt.Text)
			strVal = strVal & "&txtFrPrNo=" & Trim(.txtFrPrNo.value)
			strVal = strVal & "&txtToPrNo=" & Trim(.txtToPrNo.value)
			strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)
			strVal = strVal & "&txtBpcd_Alt=" & Trim(.txtBpCd.alt)
        Else
			strVal = strVal & "?txtFrPrDt=" & Trim(.htxtFrPrDt.Value)
			strVal = strVal & "&txtToPrDt=" & Trim(.htxtToPrDt.Value)
			strVal = strVal & "&txtFrPrNo=" & Trim(.htxtFrPrNo.value)
			strVal = strVal & "&txtToPrNo=" & Trim(.htxtToPrNo.value)
			strVal = strVal & "&txtBpCd=" & Trim(.htxtBpCd.value)
			strVal = strVal & "&txtBpcd_Alt=" & Trim(.txtBpCd.alt)
        End If   
           
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
   		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))


'--------------- ������ coding part(�������,Start)----------------------------------------------
		
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
    lgIntFlgMode     = PopupParent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    lgSaveRow        = 1

End Function

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
						<TD CLASS=TD5 NOWRAP>�Ա�����</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/f7105ra1_fpDateTime1_txtFrPrDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/f7105ra1_fpDateTime2_txtToPrDt.js'></script>
						</TD>												
						<TD CLASS=TD5 NOWRAP>�����ݹ�ȣ</TD>				
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="Text" NAME="txtFrPrNo" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="1" ALT="���ۼ����ݹ�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('FrPr',frm1.txtFrPrNo.Value)">&nbsp;~
							<INPUT TYPE="Text" NAME="txtToPrNo" SIZE=15 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="1" ALT="���ἱ���ݹ�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('ToPr',frm1.txtToPrNo.Value)">
						</TD>
					</TR>
					<TR>				
						<TD CLASS=TD5 NOWRAP>�ŷ�ó�ڵ�</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtBpCd" ALT="�ŷ�ó�ڵ�" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="11"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('BP',frm1.txtBpCd.Value)">
							<INPUT NAME="txtBpNm" ALT="�ŷ�ó��" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="14X">
						</TD>
						<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
						<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
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
						<script language =javascript src='./js/f7105ra1_vspdData_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()"></IMG>&nbsp;
					<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME=Config ONMOUSEOUT="javascript:MM_swapImgRestore()" ONMOUSEOVER="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ONCLICK="OpenOrderByPopup()"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()"></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()"></IMG></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtFrPrDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtToPrDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtFrPrNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtToPrNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBpCd"   tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

