<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111pa9
'*  4. Program Name         : ���ֹ�ȣ 
'*  5. Program Desc         : ���ֹ�ȣ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/04	
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Min, HJ	
'* 10. Modifier (Last)      : Kang Su Hwan
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
'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
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
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************%>
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                                             <%'��: Popup status                          %> 

Dim arrReturn					 '--- Return Parameter Group %>
Dim arrParam	
Dim arrParent
					
arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

Const C_PoNo 		= 1								 '��: Spread Sheet �� Columns �ε��� %>
Const BIZ_PGM_ID    = "m3111pb9.asp"
Const C_MaxKey      = 10                                    '�١١١�: Max key value

'==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= %>
Function InitVariables()
	Dim arrParam

	lgIntFlgMode = PopupParent.OPMD_CMODE								'��: Indicates that current mode is Create mode%>
	lgIntGrpCount = 0										'��: Initializes Group View Size%>
	lgPageNo = ""										'initializes Previous Key%>
		
	arrParam = arrParent(1)

	frm1.hdnRcptFlg.value = arrParam(0)
	frm1.hdnIvFlg.value = arrParam(1)
	frm1.hdnSubcontraflg.value = arrParam(2) 	' ���ְ������� �߰� 

	 '------ Coding part ------ %>
	Self.Returnvalue = Array("")
End Function
	
 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.vspdData.OperationMode = 3	
	frm1.txtFrPoDt.Text = StartDate
	frm1.txtToPoDt.Text = EndDate 
End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>                                '��: 
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA")%>
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M3111PA9","S","A","V20021202",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock 
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
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
			.Col = GetKeyPos("A",1)
			arrReturn(0) = .Text
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
	Self.Close()
End Function
	
 '------------------------------------------  OpenPoType()  -------------------------------------------------
'	Name : OpenPoType()
'	Description : OpenPoType PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPotype()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "��������"				' �˾� ��Ī 
	arrParam(1) = "M_CONFIG_PROCESS"			' TABLE ��Ī 
	
	arrParam(2) = Trim(frm1.txtPotypeCd.Value)	' Code Condition
	'arrParam(3) = Trim(frm1.txtPotypeNm.Value)	' Name Cindition
	
	arrParam(4) = "ret_flg=" & FilterVar("N", "''", "S") & " "					' Where Condition
	arrParam(5) = "��������"				' TextBox ��Ī 
	
    arrField(0) = "PO_TYPE_CD"					' Field��(0)
    arrField(1) = "PO_TYPE_NM"					' Field��(1)
    
    arrHeader(0) = "��������"				' Header��(0)
    arrHeader(1) = "�������¸�"				' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPoTypeCd.focus
		Exit Function
	Else
		frm1.txtPoTypeCd.Value    = arrRet(0)		
		frm1.txtPoTypeNm.Value    = arrRet(1)
		frm1.txtPoTypeCd.focus
	End If	
End Function

Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����ó"				' �˾� ��Ī 
	arrParam(1) = "B_BIZ_PARTNER"			' TABLE ��Ī 

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	' Code Condition
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)	' Name Cindition
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "	' Where Condition
	arrParam(5) = "����ó"				' TextBox ��Ī 
	
    arrField(0) = "BP_Cd"					' Field��(0)
    arrField(1) = "BP_NM"					' Field��(1)
    
    arrHeader(0) = "����ó"				' Header��(0)
    arrHeader(1) = "����ó��"			' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		frm1.txtSupplierCd.focus
	End If	
End Function


Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "���ű׷�"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
'	arrParam(3) = Trim(frm1.txtGroupNm.Value)	
	
	arrParam(4) = ""			
	arrParam(5) = "���ű׷�"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "���ű׷�"		
    arrHeader(1) = "���ű׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)		
		frm1.txtGroupNm.Value= arrRet(1)		
		frm1.txtGroupCd.focus
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

 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

Sub txtToPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPoDt.Action = 7
		Call SetFocusToDocument("P")                                    ' 7 : Popup Calendar ocx
		frm1.txtToPoDt.Focus
	End if
End Sub

Sub txtFrPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrPoDt.Action = 7
		Call SetFocusToDocument("P")                                    ' 7 : Popup Calendar ocx
		frm1.txtFrPoDt.Focus
	End if
End Sub


'==========================================================================================
'   Event Name : OCX_Keypress()
'   Event Desc : 
'==========================================================================================
Sub txtFrPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub


'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
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


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'��: ������ üũ 
		If lgPageNo <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			If DBQuery = False Then			
				Exit Sub
			End If
		End If
    End if
End Sub


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    
  '-----------------------
    'Erase contents area
    '----------------------- 
'    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
	frm1.vspdData.Maxrows = 0
    Call InitVariables		
    													'��: Initializes local global variables
    frm1.vspdData.Maxrows = 0
  '-----------------------
    'Check condition area
    '----------------------- 
'    If Not chkField(Document, "1") Then									'��: This function check indispensable field
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
    If Dbquery = False then Exit Function

       
    FncQuery = True																'��: Processing is OK
        
End Function	


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    
    Err.Clear                                                               '��: Protect system from crashing
    
    DbQuery = False                                                         '��: Processing is NG
    
    Dim strVal
    
    with frm1
    If LayerShowHide(1) = False Then Exit Function
	
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
	    strVal = strVal & "&lgPageNo=" & lgPageNo
	    strVal = strVal & "&txtPotype=" & .hdnPotype.value
	    strVal = strVal & "&txtSupplier=" & .hdnSupplier.value
		strVal = strVal & "&txtFrPoDt=" & .hdnFrDt.value
		strVal = strVal & "&txtToPoDt=" & .hdnToDt.value
	    strVal = strVal & "&txtGroup=" & .hdnGroup.value
	    strVal = strVal & "&txtRcptFlg=" & Trim(.hdnRcptFlg.value) '���� �߰� 
	    strVal = strVal & "&txtIvFlg=" & Trim(.hdnIvFlg.value) '���� �߰� 
	    strVal = strVal & "&txtSubcontraFlg=" & Trim(.hdnSubcontraflg.value) '���ְ������� �߰� 
	    strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
  	    strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
	    strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	else
	    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
	    strVal = strVal & "&lgPageNo=" & lgPageNo
	    strVal = strVal & "&txtPotype=" & Trim(.txtPotypeCd.value)
	    strVal = strVal & "&txtSupplier=" & Trim(.txtSupplierCd.value)
		strVal = strVal & "&txtFrPoDt=" & Trim(.txtFrPoDt.text)
		strVal = strVal & "&txtToPoDt=" & Trim(.txtToPoDt.text)
	    strVal = strVal & "&txtGroup=" & Trim(.txtGroupCd.Value)
	    strVal = strVal & "&txtRcptFlg=" & Trim(.hdnRcptFlg.value) '���� �߰� 
		strVal = strVal & "&txtIvFlg=" & Trim(.hdnIvFlg.value) '���� �߰� 
	    strVal = strVal & "&txtSubcontraFlg=" & Trim(.hdnSubcontraflg.value) '���ְ������� �߰� 
	    strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
  	    strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
	    strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
    
	end if 
	end with
	
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbQuery = True                                                          '��: Processing is NG
End Function	

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
    lgIntFlgMode = PopupParent.OPMD_UMODE
End Function

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--
'#########################################################################################################
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	' �Է� �ʵ��� ��� MaxLength=? �� ��� 
	' CLASS="required" required  : �ش� Element�� Style �� Default Attribute 
		' Normal Field�϶��� ������� ���� 
		' Required Field�϶��� required�� �߰��Ͻʽÿ�.
		' Protected Field�϶��� protected�� �߰��Ͻʽÿ�.
			' Protected Field�ϰ�� ReadOnly �� TabIndex=-1 �� ǥ���� 
	' Select Type�� ��쿡�� className�� ralargeCB�� ���� width="153", rqmiddleCB�� ���� width="90"
	' Text-Transform : uppercase  : ǥ�Ⱑ �빮�ڷ� �� �ؽ�Ʈ 
	' ���� �ʵ��� ��� 3���� Attribute ( DDecPoint DPointer DDataFormat ) �� ��� 
'######################################################################################################### 
-->
<BODY TABINDEX="-1" SCROLL="no">
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
						<TD CLASS="TD5" NOWRAP>��������</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="��������" NAME="txtPotypeCd" MAXLENGTH=5 SIZE=10 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPotype()">
											   <INPUT TYPE=TEXT AlT="��������" NAME="txtPotypeNm" SIZE=20 tag="14X" ></TD>
						<TD CLASS="TD5" NOWRAP>����ó</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="����ó" NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
											   <INPUT TYPE=TEXT AlT="����ó" ID="txtSupplierNm" NAME="arrCond" tag="14X"></TD>
					</TR>	
					<TR>	
						<TD CLASS="TD5" NOWRAP>���ֵ����</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<script language =javascript src='./js/m3111pa9_fpDateTime1_txtFrPoDt.js'></script>
									</td>
									<td>~</td>
									<td>
										<script language =javascript src='./js/m3111pa9_fpDateTime1_txtToPoDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
						<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="���ű׷�" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
											   <INPUT TYPE=TEXT AlT="���ű׷�" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
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
						<script language =javascript src='./js/m3111pa9_vspdData_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
					<IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderBy()"></IMG></TD>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnPotype" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRcptFlg" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnIvFlg" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
