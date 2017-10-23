<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : s3112pa20
'*  4. Program Name         : ���ֹ�ȣ�˾� 
'*  5. Program Desc         : ���ֹ�ȣ�˾� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/01/19
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : HJO
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
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '��: indicates that All variables must be declared in advance
                                                                            ' ��������� ������ ���� 
'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_ID 		= "s3112rb21.asp"                              '��: Biz Logic ASP Name
Const C_MaxKey          = 13                                           '��: key count of SpreadSheet

'========================================================================================================
'=									1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=									1.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop  
Dim arrReturn												'��: Return Parameter Group
Dim arrParam
Dim arrParent
					
arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
arrParam = ArrParent(1)

top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate
iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------

'--------------- ������ coding part(��������,End)-------------------------------------------------------------

 '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
                
'	frm1.hdnRetFlg.value	= arrParam(0)
'	frm1.hdnSTOFlg.Value	= arrParam(2)
	frm1.hdnPlant.value = arrParam(0)
	

	Self.Returnvalue = Array("")
End Sub
 
 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'                 ���α׷� ID�� �ְ� go��ư�� �����ų� menu tree���� Ŭ���ϴ� ���� �Ѿ��                  
'========================================================================================================= 
Sub SetDefaultVal()

'--------------- ������ coding part(�������,Start)--------------------------------------------------
	frm1.txtPlant.value=arrParam(0)
	frm1.txtplantnm.value = arrParam(1)
	
	frm1.vspdData.OperationMode = 3	
	frm1.txtFrPoDt.Text = StartDate
	frm1.txtToPoDt.Text = EndDate
'--------------- ������ coding part(�������,End)----------------------------------------------------
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>                                '��: 
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "PA") %>
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
		frm1.vspdData.Col = GetKeyPos("A",1)		'���ֹ�ȣ 
		arrReturn(0) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",2)		'���ּ��� 
		arrReturn(1) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",4)		'ǰ�� 
		arrReturn(2) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",5)		'ǰ��� 
		arrReturn(3) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",6)		'�԰� 
		arrReturn(4) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",8)		'���� 
		arrReturn(5) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",9)		'���ش��� 
		arrReturn(6) = frm1.vspdData.Text
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
		self.Returnvalue = arrReturn
		Self.Close()
End Function

'==========================================================================================
'   Event Name : txtFrPoDt
'   Event Desc :
'==========================================================================================
Sub txtFrPoDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrPoDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtFrPoDt.Focus
	End If
End Sub

'==========================================================================================
'   Event Name : txtToPoDt
'   Event Desc :
'==========================================================================================
Sub txtToPoDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToPoDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtToPoDt.Focus
	End If
End Sub	
 
'--------------------------����ó �˾� 
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"						' �˾� ��Ī 
	arrParam(1) = "B_BIZ_PARTNER"						' TABLE ��Ī 

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	' Code Condition
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)	' Name Cindition
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "							' Where Condition
	arrParam(5) = "����ó"							' TextBox ��Ī 
	
    arrField(0) = "BP_Cd"					' Field��(0)
    arrField(1) = "BP_NM"					' Field��(1)
    
    arrHeader(0) = "����ó"						' Header��(0)
    arrHeader(1) = "����ó��"						' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		frm1.txtSupplierCd.focus
	End If	
End Function

'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenOrderBy()
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

'==========================================  2.2.3 InitSpreadSheet()  ===================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("S3112RA21","S","A","V20030331",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
								
    Call SetSpreadLock 
   
End Sub

'============================================ 2.2.4 SetSpreadLock()  ====================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub	

 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029													'��: Load table , B_numeric_format
                                                                  ' 3. Spreadsheet no     
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
'    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '��: Lock  Suitable  Field    
	Call InitVariables											  '��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	
	Call FncQuery()
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
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

'==========================================================================================
'   Event Name : OCX_Keypress()
'   Event Desc : 
'==========================================================================================
Sub txtFrPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
    ElseIf KeyAscii = 13 Then
		Call FncQuery
	End if
End Sub

Sub txtToPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
    ElseIf KeyAscii = 13 Then
		Call FncQuery		
	End if
End Sub

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
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'======================================================================================================
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
    gMouseClickStatus = "SPC"
	
	If Row < 1 Then Exit Sub
	
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
'    Call ggoOper.ClearField(Document, "2")	         						'��: Clear Contents  Field
	frm1.vspdData.Maxrows = 0
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then								'��: This function check indispensable field
'       Exit Function
'    End If
	
	with frm1
		if (UniConvDateToYYYYMMDD(.txtFrPoDt.text,gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToPoDt.text,gDateFormat,"")) And Trim(.txtFrPoDt.text) <> "" And Trim(.txtToPoDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","������", "X")	
			
			Exit Function
		End if   
	End with

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
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
			strVal = strVal & "&txtPlant=" & Trim(.hdnPlant.value)
		    strVal = strVal & "&txtSupplierCd=" & .hdnSupplier.value
			strVal = strVal & "&txtFrPoDt=" & .hdnFrDt.value
			strVal = strVal & "&txtToPoDt=" & .hdnToDt.value
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey     
		else
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
			strVal = strVal & "&txtPlant=" & Trim(.txtPlant.value)
		    strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)
			strVal = strVal & "&txtFrPoDt=" & Trim(.txtFrPoDt.text)
			strVal = strVal & "&txtToPoDt=" & Trim(.txtToPoDt.text)
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey     
		end if 
	
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
						<TD CLASS="TD5" NOWRAP>����</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="����" NAME="txtPlant" MAXLENGTH=5 SIZE=10 tag="14xxxU">
											   <INPUT TYPE=TEXT AlT="�����" NAME="txtPlantNm" SIZE=20 tag="14X" ></TD>
						<TD CLASS="TD5" NOWRAP>����ó</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="����ó" NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
											   <INPUT TYPE=TEXT AlT="����ó" ID="txtSupplierNm" NAME="arrCond" tag="14X"></TD>
					</TR>	
					<TR>	
						<TD CLASS="TD5" NOWRAP>������</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<script language =javascript src='./js/s3112ra21_fpDateTime1_txtFrPoDt.js'></script>
									</td>
									<td>~</td>
									<td>
										<script language =javascript src='./js/s3112ra21_fpDateTime1_txtToPoDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
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
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						<script language =javascript src='./js/s3112ra21_vaSpread1_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"></IMG>
					<IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderBy()"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> SCROLLING=NO noresize  FRAMEBORDER=0 framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
