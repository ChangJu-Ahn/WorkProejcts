<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M3111PA7
'*  4. Program Name         : ��ǰ���ֹ�ȣ 
'*  5. Program Desc         : ��ǰ���ֹ�ȣ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Jin-hyun Shin
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit                                                             '��: indicates that All variables must be declared in advance
'========================================================================================================
'=									1.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID 		= "m3111pb7.asp"                              '��: Biz Logic ASP Name
Const C_MaxKey          = 11                                           '��: key count of SpreadSheet
Const C_PoNo 		= 1								 '��: Spread Sheet �� Columns �ε��� 

'========================================================================================================
'=									1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

                                                                            ' ��������� ������ ���� 
'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Dim IsOpenPop  
Dim gblnWinEvent											'��: ShowModal Dialog(PopUp) 
														    'Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
														    'PopUp Window�� ��������� ���θ� ��Ÿ�� 
Dim arrReturn												'��: Return Parameter Group
Dim arrParam
Dim arrParent
					
arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
arrParam		= arrParent(1)						' add 20040310 by TJ.Kim
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'========================================== 2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================
Function InitVariables()
	Redim arrReturn(0) 
	
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    
    lgSortKey        = 1   
    lgIntGrpCount	 = 0										'��: Initializes Group View Size
	gblnWinEvent	 = False
    Self.Returnvalue = arrReturn     
End Function

'==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'                 ���α׷� ID�� �ְ� go��ư�� �����ų� menu tree���� Ŭ���ϴ� ���� �Ѿ��                  
'========================================================================================================= 
Sub SetDefaultVal()
	Dim arrTemp
	ON ERROR RESUME NEXT
	with frm1
'------------------------- add 20040309	by JT.Kim
		If "" & Trim(arrParam(0)) <> "" Then
			.txtPotypeCd.value	= arrParam(0)
		End If 	
'--------------------------
		.vspdData.focus
		.vspdData.OperationMode = 3	
		.txtFrPoDt.Text = StartDate
		.txtToPoDt.Text = EndDate
	end with
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtGroupCd, "Q") 
  	frm1.txtGroupCd.value = lgPGCd
	End If
End Sub


'==========================================  2.2.2 LoadInfTB19029() =====================================
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
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "PA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "PA")%>
End Sub


'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M3111PA7","S","A","V20021202",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
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
		
	If frm1.vspdData.ActiveRow > 0 Then	
		Redim arrReturn(frm1.vspdData.MaxCols-2)

		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		For intColCnt = 0 To frm1.vspdData.MaxCols - 2
			frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
			arrReturn(intColCnt) = frm1.vspdData.Text
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
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'------------------------------------------  OpenPoType()  -------------------------------------------------
'	Name : OpenPoType()
'	Description : OpenPoType PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConSItemDC(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True

	Select Case iWhere
		Case 0	'�������� 
			arrParam(1) = "M_CONFIG_PROCESS"					' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtPotypeCd.Value)			' Code Condition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "��������"							' TextBox ��Ī 
	
			arrField(0) = "PO_TYPE_CD"							' Field��(0)
			arrField(1) = "PO_TYPE_NM"							' Field��(1)
    
			arrHeader(0) = "��������"						' Header��(0)
			arrHeader(1) = "�������¸�"						' Header��(1)
    
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		Case 1	'����ó 
			arrParam(1) = "B_BIZ_PARTNER"							' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtSupplierCd.Value)			' Code Condition
			arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "' Where Condition
			arrParam(5) = "����ó"								' TextBox ��Ī 
	
			arrField(0) = "BP_Cd"									' Field��(0)
			arrField(1) = "BP_NM"									' Field��(1)
    
			arrHeader(0) = "����ó"								' Header��(0)
			arrHeader(1) = "����ó��"							' Header��(1)
    
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
		Case 2
			gblnWinEvent = False
			If frm1.txtGroupCd.className = "protected" Then Exit Function
			gblnWinEvent = True

			arrParam(1) = "B_Pur_Grp"				
			arrParam(2) = Trim(frm1.txtGroupCd.Value)
			arrParam(4) = ""			
			arrParam(5) = "���ű׷�"			
			
		    arrField(0) = "PUR_GRP"	
		    arrField(1) = "PUR_GRP_NM"	
		    
		    arrHeader(0) = "���ű׷�"		
		    arrHeader(1) = "���ű׷��"		
		    
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select
	
	arrParam(0) = arrParam(5)												' �˾� ��Ī	

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function

'-------------------------------------------------------------------------------------------------------
'	Name : SetConSItemDC()
'	Description : OpenConSItemDC Popup���� Return�Ǵ� �� setting
'-------------------------------------------------------------------------------------------------------
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtPoTypeCd.Value		= arrRet(0)		
				.txtPoTypeNm.Value		= arrRet(1)   
			Case 1
				.txtSupplierCd.Value	= arrRet(0)		
				.txtSupplierNm.Value	= arrRet(1)	
			Case 2
				.txtGroupCd.Value		= arrRet(0)		
				.txtGroupNm.Value		= arrRet(1)	
		End Select
	End With
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

	On Error Resume Next

	Call LoadInfTB19029													'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
'    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                         '��: Lock  Suitable  Field
	Call InitVariables											  '��: Initializes local global variables
  Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub

'=========================================  3.3.1 vspdData_DblClick()  ==================================
'=	Event Name : vspdData_DblClick																		=
'=	Event Desc :																						=
'========================================================================================================
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

'*********************************************  3.3 Object Tag ó��  ************************************
'*	Object���� �߻� �ϴ� Event ó��																		*
'********************************************************************************************************
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'==========================================================================================
'   Event Name : OCX_KeyDown()
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

'==========================================================================================
'   Event Name : OCX_DbClick()
'   Event Desc : OCX_DbClick() �� Calendar Popup
'==========================================================================================
Sub txtFrPoDt_DblClick(Button)
	If Button = 1 Then
       frm1.txtFrPODt.Action = 7
       Call SetFocusToDocument("P")                                    ' 7 : Popup Calendar ocx
       frm1.txtFrPoDt.Focus
    End If
End Sub

Sub txtToPoDt_DblClick(Button)
	If Button = 1 Then
       frm1.txtToPoDt.Action = 7  
       Call SetFocusToDocument("P")                                  ' 7 : Popup Calendar ocx
       frm1.txtToPoDt.Focus
    End If
End Sub	
'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
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

Function FncQuery() 
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

	'-----------------------
    'Erase contents area
    '-----------------------
'    Call ggoOper.ClearField(Document, "2")	         						'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
	frm1.vspdData.Maxrows = 0
	
	'-----------------------
    'Check condition area
    '----------------------- 
'    If Not chkField(Document, "1") Then									'��: This function check indispensable field
'       Exit Function
'    End If
    
	If ValidDateCheck(frm1.txtFrPoDt, frm1.txtToPoDt) = False Then Exit Function
	
	'-----------------------
    'Query function call area
    '----------------------- 
    If frm1.rdoPostFlg1.checked = True Then
		frm1.hdnRdoFlg.value = ""
	ElseIf frm1.rdoPostFlg2.checked = True Then
		frm1.hdnRdoFlg.value = "Y"
	ElseIf frm1.rdoPostFlg3.checked = True Then
		frm1.hdnRdoFlg.value = "N"
	End If
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
	Dim strVal

	Err.Clear                                                               '��: Protect system from crashing
    DbQuery = False
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	strVal = ""
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode="&PopupParent.UID_M0001
		    strVal = strVal & "&txtPotypeCd=" & Trim(frm1.hdnPotype.value)
		    strVal = strVal & "&txtSupplierCd=" & Trim(frm1.hdnSupplier.value)
			strVal = strVal & "&txtFrPoDt=" & Trim(frm1.hdnFrDt.Value)
			strVal = strVal & "&txtToPoDt=" & Trim(frm1.hdnToDt.Value)
		    strVal = strVal & "&txtGroupCd=" & Trim(frm1.hdnGroup.value)
		    strVal = strVal & "&txtRadio="&Trim(frm1.hdnRadio.value) '13�� �߰�	
		    strVal = strVal & "&hdnRetFlg="&Trim(frm1.hdnRetFlg.value) '��ǰ���� �߰� 
		
		else
		    strVal = BIZ_PGM_ID & "?txtMode="&PopupParent.UID_M0001
		    strVal = strVal & "&txtPotypeCd=" & Trim(.txtPotypeCd.value)
		    strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)
			strVal = strVal & "&txtFrPoDt=" & Trim(.txtFrPoDt.text)
			strVal = strVal & "&txtToPoDt=" & Trim(.txtToPoDt.text)
		    strVal = strVal & "&txtGroupCd=" & Trim(.txtGroupCd.Value)
		    strVal = strVal & "&txtRadio=" & Trim(.hdnRdoFlg.value) '13�� �߰�	
		    strVal = strVal & "&hdnRetFlg=" & Trim(.hdnRetFlg.value) '��ǰ���� �߰� 
		
		end if 

		strVal = strVal & "&lgPageNo="		 & lgPageNo						'��: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  
        
        Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
    lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtPotypeCd.focus
	End If
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
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP>��ǰ����</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="��ǰ����" NAME="txtPotypeCd" MAXLENGTH=5 SIZE=10 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 0">
											   <INPUT TYPE=TEXT AlT="��ǰ����" NAME="txtPotypeNm" SIZE=20 tag="14X" ></TD>
						<TD CLASS="TD5" NOWRAP>����ó</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="����ó" NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 1">
											   <INPUT TYPE=TEXT AlT="����ó" ID="txtSupplierNm" NAME="arrCond" tag="14X"></TD>
					</TR>	
					<TR>	
						<TD CLASS="TD5" NOWRAP>��ǰ�����</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<script language =javascript src='./js/m3111pa7_fpDateTime1_txtFrPoDt.js'></script>
									</td>
									<td>~</td>
									<td>
										<script language =javascript src='./js/m3111pa7_fpDateTime1_txtToPoDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
						<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="���ű׷�" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 2">
											   <INPUT TYPE=TEXT AlT="���ű׷�" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>Ȯ������</TD> 
						<TD CLASS=TD6 colspan=3 NOWRAP>
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostFlg" TAG="11X" VALUE=""  ID="rdoPostFlg1"><LABEL FOR="rdoPostFlg1">��ü</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostFlg" TAG="11X" VALUE="Y" ID="rdoPostFlg2"><LABEL FOR="rdoPostFlg2">Ȯ��</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPostFlg" TAG="11X" VALUE="N" CHECKED ID="rdoPostFlg3"><LABEL FOR="rdoPostFlg3">��Ȯ��</LABEL>			
						</TD>
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
						<script language =javascript src='./js/m3111pa7_vspdData_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnPotype" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRdoFlg" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnRadio" TAG="24">
<INPUT TYPE=HIDDEN NAME="hdnRetFlg" TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
