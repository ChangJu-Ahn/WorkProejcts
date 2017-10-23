<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M4132RA1
'*  4. Program Name         : ���ܹ�ǰ������� 
'*  5. Program Desc         : ���ܹ�ǰ������� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/11/21
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Duk Hyun
'* 10. Modifier (Last)      : 
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

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
Const ivType = "ST"
Const BIZ_PGM_ID        = "M4132RB1_KO441.asp"
Const C_MaxKey          = 26                                    '�١١١�: Max key value
Const gstPaytermsMajor = "B9004"

Dim lgIsOpenPop                                             '��: Popup status                           
Dim arrReturn					 '--- Return Parameter Group 
DIM lblnWinEvent
Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
 
 '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
	Dim arrParam
    
	arrParam = arrParent(1)

	frm1.hdnSupplierCd.value 	= arrParam(0)
	frm1.hdnSubcontra2flg.value	= arrParam(1)
    
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
    lgIntFlgMode = PopupParent.OPMD_CMODE	
	frm1.vspdData.OperationMode = 5
	
	Redim arrReturn(0, 0)
	Self.Returnvalue = arrReturn
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
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>                                '��: 
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M4132RA1","S","A","V20021202",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
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

'------------------------------------------  OpenMvmtNo()  -------------------------------------------------
'	Name : OpenMvmtNo()
'--------------------------------------------------------------------------------------------------------- 
Function OpenMvmtNo()
	
	Dim strRet
	Dim arrParam(3)
	Dim iCalledAspName
	Dim IntRetCD
	
	If lblnWinEvent = True Or UCase(frm1.txtMvmtNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function
	
	arrParam(0)	=	""
	arrParam(1)	=	""
	arrParam(2)	=	""
	arrParam(3)	=	"N" 'Rcpt Flag
	
	lblnWinEvent = True
	
	iCalledAspName = AskPRAspName("M4141PA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "M4141PA2", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False

	If strRet(0) = "" Then
		frm1.txtMvmtNo.focus
		Exit Function
	Else
		frm1.txtMvmtNo.value = strRet(0)
		frm1.txtMvmtNo.focus
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
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call InitVariables														'��: Initializes local global variables
    Call GetValue_ko441()
	frm1.txtFrIvDt.text = StartDate
	frm1.txtToIvDt.text = EndDate
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'==================================== 3.2.23 txtToIvDt_DblClick()  =====================================
'   Event Name : txtToIvDt_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtToIvDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToIvDt.Action = 7
		Call SetFocusToDocument("P")                                    ' 7 : Popup Calendar ocx
		frm1.txtToIvDt.Focus
	End If
End Sub
'==================================== 3.2.23 txtFrIvDt_DblClick()  =====================================
'   Event Name : txtFrIvDt_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtFrIvDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrIvDt.Action = 7
		Call SetFocusToDocument("P")                                    ' 7 : Popup Calendar ocx
		frm1.txtFrIvDt.Focus
	End If
End Sub


'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : ��ȸ���Ǻ��� OCX_KeyDown�� EnterKey�� ���� Query
'==========================================================================================
Sub txtFrIvDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToIvDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub


'=========================================  3.3.1 vspdData_DblClick()  ==================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
	     Exit Function
	End If
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
	

'========================================  3.3.2 vspdData_LeaveCell()  ==================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKey <> "" Then							 '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
				DbQuery
			End If
		End If
	End With
End Sub
	

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
	    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'��: ������ üũ'
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			DbQuery
		End If
   End if
End Sub


'======================================  3.3.4 vspdData_KeyPress()  ================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And Frm1.vspdData.ActiveRow > 0 Then    'Frm1������ frm1���� 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function


'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	Dim intColCnt, intRowCnt, intInsRow

	If frm1.vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		Redim arrReturn(frm1.vspdData.SelModeSelCount-1, frm1.vspdData.MaxCols-2)

		For intRowCnt = 1 To frm1.vspdData.MaxRows

			frm1.vspdData.Row = intRowCnt

			If frm1.vspdData.SelModeSelected Then
				For intColCnt = 0 To frm1.vspdData.MaxCols - 2
					frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
					arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
				Next
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

'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery() 

    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

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

    '-----------------------
    'Query function call area
    '-----------------------
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'�� 'pObjFromDt'���� Ŀ�� �Ҷ� **
	If ValidDateCheck(frm1.txtFrIvDt, frm1.txtToIvDt) = False Then Exit Function

    Call DbQuery															'��: Query db data

    FncQuery = True		
End Function


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing

	If LayerShowHide(1) = False Then
		Exit Function
	End If
    
    With frm1
       If lgIntFlgMode = PopupParent.OPMD_UMODE Then
'--------------- ������ coding part(�������,Start)----------------------------------------------
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		    strVal = strVal & "&txtMvmtNo=" & Trim(.hdnMvmtNo.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtFrIvDt=" & Trim(.txtFrIvDt.text)
		    strVal = strVal & "&txtToIvDt=" & Trim(.txtToIvDt.text)
		    strVal = strVal & "&txtSupplierCd= " & Trim(.hdnSupplierCd.value)
			strVal = strVal & "&txtSubcontra2flg=" & Trim(.hdnSubcontra2flg.value)		' ���ְ������� �߰� 
'--------------- ������ coding part(�������,End)------------------------------------------------
            strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                   		'��: Next key tag
		    strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
            strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		    strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	    Else
'--------------- ������ coding part(�������,Start)----------------------------------------------
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		    strVal = strVal & "&txtMvmtNo=" & Trim(.txtMvmtNo.value)					'��: ��ȸ ���� ����Ÿ 
		    strVal = strVal & "&txtFrIvDt=" & Trim(.txtFrIvDt.text)
		    strVal = strVal & "&txtToIvDt=" & Trim(.txtToIvDt.text)
		    strVal = strVal & "&txtSupplierCd= " & Trim(.hdnSupplierCd.value)
			strVal = strVal & "&txtSubcontra2flg=" & Trim(.hdnSubcontra2flg.value)		' ���ְ������� �߰� 
'--------------- ������ coding part(�������,End)------------------------------------------------
            strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                   		'��: Next key tag
		    strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
            strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		    strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
       End if   
       
        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

       Call RunMyBizASP(MyBizASP, strVal)												'��: �����Ͻ� ASP �� ���� 
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
		frm1.vspdData.focus
	Else
		frm1.txtMvmtNo.focus
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
						<TD CLASS="TD5" NOWRAP>��ǰ����ȣ</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="��ǰ����ȣ" NAME="txtMvmtNo" MAXLENGTH=18 SIZE=32 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMvmt" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMvmtNo()">
											   <div style="Display:none"><input type=text name=none></div></TD>
						<TD CLASS="TD5" NOWRAP>��ǰ�����</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=��ǰ����� NAME="txtFrIvDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
									<td>~</td>
									<td>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=��ǰ����� NAME="txtToIvDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
								<tr>
							</table>
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
					<TD HEIGHT="100%">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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

<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnMvmtNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnCurrency" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSubcontra2flg" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
