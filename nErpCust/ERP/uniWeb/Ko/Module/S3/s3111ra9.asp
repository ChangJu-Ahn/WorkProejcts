<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : s3111ra9
'*  4. Program Name         : project popup
'*  5. Program Desc         : project popup
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/11/25
'*  8. Modified date(Last)  : 2005/11/25
'*  9. Modifier (First)     : Nhg
'* 10. Modifier (Last)      : Nhg
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>
<!--
'********************************************  1.1 Inc ����  ********************************************
'*	Description : Inc. Include																			*
'********************************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>
<!--
'============================================  1.1.2 ���� Include  ======================================
'========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit					<% '��: indicates that All variables must be declared in advance %>

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************%>
Const BIZ_PGM_ID        = "s3111rb9.asp"                       ' �����Ͻ� ���� ������ ���� 
'========================================================================================================
'=									1.2 Constant variables 
'========================================================================================================
Const C_MaxKey  = 12                                           '��: key count of SpreadSheet
Const C_No 		= 1		


Dim C_PROJECT
Dim	C_PROJECTNM
Dim	C_VER
Dim	C_PMRESOURCECD
Dim	C_PLAN_START_DT
Dim	C_PLAN_END_DT
Dim	C_REAL_START_DT
Dim C_REAL_END_DT

'========================================================================================================
'=									1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=									1.4 User-defind Variables
'========================================================================================================
Dim gblnWinEvent
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam
Dim IsOpenPop  

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam= arrParent(1)
top.document.title = PopupParent.gActivePRAspName

'--------------- ������ coding part(��������,End)-------------------------------------------------------------
'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= %>
'----------------  ���� Global ������ ����  ----------------------------------------------------------- %>
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ %>
Dim arrValue(3)                    ' Popup�Ǵ� â���� �ѱ涧 �μ��� �迭�� �ѱ� 
'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Sub InitVariables()
	Dim arrParent
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                         'Indicates that current mode is Create mode
    IsOpenPop = False
    gblnWinEvent = False
    
    <% '------ Coding part ------ %>
	Self.Returnvalue = Array("")
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>
End Sub
'******************************************* 2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ====================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)	
'=======================================================================================================
Sub SetDefaultVal()
	
End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("Y7001PA1","S","A","V20021202",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")


    Call SetSpreadLock("A")
    frm1.vspdData.OperationMode = 3 
    
End Sub
'============================================ 2.2.4 SetSpreadLock()  ====================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    IF pOpt = "A" Then
	  ggoSpread.Source = frm1.vspdData
	  ggoSpread.SpreadLockWithOddEvenRowColor()
	Else
	
	End IF
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	Dim intColCnt

	With frm1.vspdData	
		Redim arrReturn(.MaxCols - 1)
		If .MaxRows > 0 Then 
			.Row = .ActiveRow
			.Col = GetKeyPos("A",C_No)	
			arrReturn(0) = .Text
			.Col = GetKeyPos("A",2)	
			arrReturn(1) = .Text
			.Col = GetKeyPos("A",3)	
			arrReturn(2) = .Text
		end if
	End With
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'========================================================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenSortPopup Reference Popup
'========================================================================================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
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

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function OpenProject()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtProjectCd.ClassName)=UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "������Ʈ"	
	arrParam(1) = "PMS_PROJECT"				
	
	arrParam(2) = Trim(frm1.txtProjectCd.value)
'	arrParam(3) = ""
	
	arrParam(4) = ""			
	arrParam(5) = "������Ʈ"			
	
    arrField(0) = "Project_CODE"	
    arrField(1) = "Project_NM"	
    
    arrHeader(0) = "������Ʈ"		
    arrHeader(1) = "������Ʈ��"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtProjectCd.focus
		Exit Function
	Else
		frm1.txtProjectCd.Value= arrRet(0)		
		frm1.txtProjectNm.Value= arrRet(1)		
		frm1.txtProjectCd.focus
	End If	
	
End Function

'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029													'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
'    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                         '��: Lock  Suitable  Field
    Call InitVariables											  '��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
	frm1.txtProjectCd.focus
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    
End Sub

'=========================================  3.3.1 vspdData_DblClick()  ==================================
'=	Event Name : vspdData_DblClick																		=
'=	Event Desc :																						=
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or frm1.vspdData.MaxRows = 0 Then 
        Exit Sub
    End If

	With frm1.vspdData 
		If .MaxRows > 0 Then
			If .ActiveRow = Row Or .ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End With
End Sub

'========================================  3.3.2 vspdData_KeyPress()  ===================================
'=	Event Name : vspdData_KeyPress																		=
'=	Event Desc :																						=
'========================================================================================================
Sub vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Sub

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
'=	Event Name : vspdData_TopLeftChange																	=
'=	Event Desc :																						=
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If frm1.vspdData.MaxRows < NewTop +  + VisibleRowCnt(frm1.vspdData,NewTop) And lgPageNo <> "" Then	    <%'��: ������ üũ %>
		If CheckRunningBizProcess = True Then
			Exit Sub
		End If
		If DBQuery = False Then		
			Exit Sub
		End If
    End If
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
    Call ggoOper.ClearField(Document, "2")	         						'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then								'��: This function check indispensable field
'       Exit Function
'    End If

    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
End Function	
	
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Err.Clear														'��: Protect system from crashing
	DbQuery = False													'��: Processing is NG

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ����	
			strVal = strVal & "&txtProject=" & .txtProjectCd.value
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey     
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtProject=" & .txtProjectCd.value
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		End If				
			
			strVal = strVal & "&lgPageNo="		 & lgPageNo						'��: Next key tag 
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)		    						'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True    
End Function	

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()	    												'��: ��ȸ ������ ������� 
	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	End If
	
	frm1.vspdData.focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--
'########################################################################################################
'#						6. TAG ��																		#
'########################################################################################################
-->
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
						<TD CLASS="TD5" NOWRAP>������Ʈ</TD>
						<TD CLASS="TD656" NOWRAP colspan = "3"><INPUT TYPE=TEXT ALT="������Ʈ" NAME="txtProjectCd" SIZE=18 MAXLENGTH=25 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenProject()">
											   <INPUT TYPE=TEXT NAME="txtProjectNm" SIZE=30 tag="14x"></TD>
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
					<TD HEIGHT="100%">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnProject" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
