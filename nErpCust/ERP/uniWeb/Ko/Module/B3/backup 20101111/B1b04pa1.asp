<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : ��������																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : B1b04pa1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : HS Code PopUp ASP															*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/04/18																*
'*  8. Modified date(Last)  : 2000/04/18		
'*                            2002/04/28  														*
'*  9. Modifier (First)     : 																			*
'* 10. Modifier (Last)      : Park JIn Uk																			*
'*                            Kim Jae Soon
'* 11. Comment              :																			*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>
<!--<TITLE>HS��ȣ</TITLE> -->
<% '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################%>
<% '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* %>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<%'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<%'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================%>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit																	'��: indicates that All variables must be declared in advance

<%'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************%>
Const BIZ_PGM_ID        = "B1b04pb1.asp"                       ' �����Ͻ� ���� ������ ���� 

'========================================================================================================
'=									1.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '��: Fetch max count at once
Const C_MaxKey          = 5                                           '��: key count of SpreadSheet

Dim C_MaxSelList

'========================================================================================================
'=									1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=									1.4 User-defind Variables
'========================================================================================================
'Dim lgSelectList                                            '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
'Dim lgSelectListDT                                          '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 

'Dim lgSortFieldNm                                           '��: Orderby popup�� ����Ÿ(�ʵ弳��)      
'Dim lgSortFieldCD                                           '��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      

'Dim lgPopUpR                                                '��: Orderby default ��                    

'Dim lgKeyPos                                                '��: Key��ġ                               
'Dim lgKeyPosVal                                             '��: Key��ġ Value                         
'Dim IscookieSplit 

Dim arrParent
Dim lgIsOpenPop  
Dim gblnWinEvent											'��: ShowModal Dialog(PopUp) 
														    'Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
														    'PopUp Window�� ��������� ���θ� ��Ÿ�� 
Dim arrReturn												'��: Return Parameter Group
Dim arrParam

Const C_Hs 		= 1		
Const C_HsNm    = 2

'--------------- ������ coding part(��������,End)-------------------------------------------------------------
<% '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= %>
<% '----------------  ���� Global ������ ����  ----------------------------------------------------------- %>

Dim arrValue(3)                    ' Popup�Ǵ� â���� �ѱ涧 �μ��� �迭�� �ѱ� 
Dim strReturn						<% '--- Return Parameter Group %>

	
	'------ Set Parameters from Parent ASP ------ 
	arrParent = window.dialogArguments
	Set PopupParent = arrParent(0)
	arrParam = arrParent(1)
	top.document.title = PopupParent.gActivePRAspName
	
<% '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ %>

<% '#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### %>
<% '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= %>
Sub InitVariables()
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    gblnWinEvent = False
    arrParent = window.dialogArguments
	Self.Returnvalue = Array("")
End Sub

<% '******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* %>

<%'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
<% '== ��ȸ,��� == %>
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "*", "NOCOOKIE", "PA") %>
End Sub
<%
'==========================================  2.2.3 InitSpreadSheet()  ===================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
%>
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("B1b04pa1","S","A","V20021202",Popupparent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock("A")
	frm1.vspdData.OperationMode = 3
End Sub
<%
'============================================ 2.2.4 SetSpreadLock()  ====================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
%>
Sub SetSpreadLock(ByVal pOpt)
   
    If pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	Else
	
	End If
	
End Sub

<%
'++++++++++++++++++++++++++++++++++++++++++  2.3 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	������ ���� Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>
<%
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
%>	
Function OKClick()
	Dim intColCnt
	
	With frm1.vspdData
		Redim arrReturn(.MaxCols -1)
		If .MaxRows > 0 Then 
			.Row = .ActiveRow
			.Col = GetKeyPos("A",C_Hs)
			arrReturn(0) = .Text
			.Col = GetKeyPos("A",C_HsNm)
			arrReturn(1) = .Text
		End if
	End With			
	Self.Returnvalue = arrReturn
	Self.Close()
			
End Function
<%
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
%>
Function CancelClick()
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()	
End Function
<%
'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
%>
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function
	
<% 
'*******************************************  2.4 POP-UP ó���Լ�  **************************************
'*	���: POP-UP																						*
'*	Description : POP-UP Call�ϴ� �Լ� �� Return Value setting ó��										*
'********************************************************************************************************
%>
<%
'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================
%>
<%
'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================
%>

<% '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ %>

<%
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	���� ���α׷����� �ʿ��� ������ ���� Procedure(Sub, Function, Validation & Calulation ���� �Լ�)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>

<%
'########################################################################################################
'#						3. Event ��																		#
'#	���: Event �Լ��� ���� ó��																		#
'#	����: Windowó��, Singleó��, Gridó�� �۾�.														#
'#		  ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.								#
'#		  �� Object������ Grouping�Ѵ�.																	#
'########################################################################################################
%>
<%
'********************************************  3.1 Windowó��  ******************************************
'*	Window�� �߻� �ϴ� ��� Even ó��																	*
'********************************************************************************************************
%>
<%
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
%>
Sub Form_Load()
   									'��: Load table , B_numeric_format
    ReDim lgPopUpR(C_MaxSelList - 1,1)
    Call LoadInfTB19029				
    'Call GetAdoFieldInf("B1B04PA1","S","A")			              '��: spread sheet �ʵ����� query
                                                                   ' 3. Spreadsheet no     
    'Html���� tag ���ڰ� 1�� 2�� �����ϴ� �κ� ����Format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Popupparent.gDateFormat,Popupparent.gComNum1000,Popupparent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Popupparent.gDateFormat,Popupparent.gComNum1000,Popupparent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    
	Call ggoOper.LockField(Document, "N")                         '��: Lock  Suitable  Field
    'Call MakePopData(gDefaultT,gFieldNM,gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,C_MaxSelList)    ' You must not this line    
    Call InitVariables											  '��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	
	Call FncQuery()
<%'--------------- ������ coding part(�������,Start)----------------------------------------------------%>

<%'--------------- ������ coding part(�������,End)------------------------------------------------------%>
End Sub



Function OpenSortPopup()

	Dim arrRet
	
	On Error Resume Next
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"), gMethodText),"dialogWidth=" & PopupParent.GROUPW_WIDTH & "px; dialogHeight=" & PopupParent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False

	If arrRet(0) = "X" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A", arrRet(0), arrRet(1))
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function


<% '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'                 ���α׷� ID�� �ְ� go��ư�� �����ų� menu tree���� Ŭ���ϴ� ���� �Ѿ��                  
'========================================================================================================= %>
Sub SetDefaultVal()
	
End Sub

<%
'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
%>
	Sub Form_QueryUnload(Cancel, UnloadMode)
	End Sub
<%
'*********************************************  3.2 Tag ó��  *******************************************
'*	Document�� TAG���� �߻� �ϴ� Event ó��																*
'*	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ�							*
'*	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.																	*
'********************************************************************************************************
%>
<%
'==========================================  3.2.1 Search_OnClick =======================================
'========================================================================================================
%>

<%
'======================================  3.2.1 Search_OnClick()  ====================================
'=	Event Name : Search_OnClick																		=
'=	Event Desc :																						=
'========================================================================================================
%>
	Sub Search_OnClick()
		Call fncquery()
	End Sub
<%
'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
%>
	Function vspdData_DblClick(ByVal Col, ByVal Row)
		If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
			Exit Function
		End If
    
		With frm1.vspdData 
			If .MaxRows > 0 Then
				If .ActiveRow = Row Or .ActiveRow > 0 Then
					Call OKClick
				End If
			End If
		End With
	End Function

    Function vspdData_KeyPress(KeyAscii)
         On Error Resume Next
         If KeyAscii = 13 And Frm1.vspdData.ActiveRow > 0 Then    'Frm1������ frm1���� 
            Call OKClick()
         ElseIf KeyAscii = 27 Then
            Call CancelClick()
         End If
    End Function

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    

		If OldLeft <> NewLeft Then
		    Exit Sub
		End If		

		If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ	
			If lgPageNo <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				If DbQuery = False Then
					Exit Sub
				End if
			End If
		End If		  

End Sub

<% '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'######################################################################################################### %>

<% '#########################################################################################################
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
'######################################################################################################### %>
<%
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
%>

Function FncQuery()

    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									
	FncQuery = True		
End Function

<%
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
%>
Function DbQuery() 
	Err.Clear														'��: Protect system from crashing
	DbQuery = False													'��: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ����	
		strVal = strVal & "&txtHsCd=" & Trim(frm1.txtHsCd.value)					<%'��: ��ȸ ���� ����Ÿ %>
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey     
			
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'��: Next key tag 
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D             '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ�  
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
		Call RunMyBizASP(MyBizASP, strVal)		    						'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True    
    
End Function

<%
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
%>
Function DbQueryOk()														<%'��: ��ȸ ������ ������� %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	else
		frm1.vspdData.Focus
	End If

End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG ��																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5>HS��ȣ</TD>
						<TD CLASS=TD6><INPUT NAME="txtHsCd" MAXLENGTH="12" SIZE=20 ALT ="HS��ȣ" tag="11"></TD>
						<TD CLASS=TD6><div style="display:none"><INPUT NAME="none"></div></TD>
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
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
					    <script language =javascript src='./js/b1b04pa1_vaSpread1_vspdData.js'></script>
					</TD>		
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
						                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                         <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
