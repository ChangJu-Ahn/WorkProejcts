<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���԰��� 
'*  3. Program ID           : m5111pa1
'*  4. Program Name         : ���Թ�ȣ 
'*  5. Program Desc         : ���Գ�������� ���Թ�ȣ 
'*  6. Comproxy List        : M51118ListIvHdrSvr
'*  7. Modified date(First) : 2002/02/16
'*  8. Modified date(Last)  : 2003/06/04
'*  9. Modifier (First)     : Shin Jin Hyen				
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<!--<TITLE>���Թ�ȣ</TITLE> -->
<TITLE></TITLE>
<!--
'******************************************  1.1 Inc ����   **********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 ���� Include   ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

Const BIZ_PGM_ID 		= "m5111pb1.asp"                              '��: Biz Logic ASP Name
Const C_MaxKey          = 12                                          '��: key count of SpreadSheet
Const IV_NO = 1

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IscookieSplit 
Dim IsOpenPop  
Dim gblnWinEvent											'��: ShowModal Dialog(PopUp) 
Dim arrReturn												'��: Return Parameter Group
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

'========================================== 2.1.1 InitVariables()  ======================================
	Function InitVariables()
		lgStrPrevKey     = ""								   'initializes Previous Key
		lgPageNo         = ""
        lgBlnFlgChgValue = False	                           'Indicates that no value changed
        lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
                
        gblnWinEvent = False
        Redim arrReturn(0)        
        Self.Returnvalue = arrReturn
	End Function

'==========================================  2.2.1 SetDefaultVal()  ====================================
Sub SetDefaultVal()
	Dim arrParam
    
	arrParam = arrParent(1)

	frm1.txtFrIvDt.text=StartDate									
	frm1.txtToIvDt.text=EndDate									

    ' ���������� Ȯ���� Default
    ' �������������� arrParent(1)�� 2�� �Ѿ���� ������.
	If UBound(arrparam) = 1 Then
	    If arrparam(1) = "Y" Then
			frm1.hdtxtRadio.value = "Y"
			frm1.rdoPostFlg2.checked = True
		End If		    
	End If
End Sub

'==========================================  2.2.2 LoadInfTB19029() =====================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "PA")%>
End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M5111PA1","S","A","V20030428",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
    Call SetSpreadLock
    frm1.vspdData.OperationMode = 3
End Sub
'============================================ 2.2.4 SetSpreadLock()  ====================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub	

'==========================================  2.3.1 OkClick()  ===========================================
Function OKClick()
	Dim intColCnt
		
	If frm1.vspdData.ActiveRow > 0 Then	
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",IV_NO)
		arrReturn(0) = frm1.vspdData.Text
	End if
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'======================================== OpenConSItemDC() ===============================================
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
	
	Select Case iWhere
	Case 0
		arrHeader(0) = "��������"						<%' Header��(0)%>
	    arrHeader(1) = "�������¸�"						<%' Header��(1)%>
	    
	    arrField(0) = "IV_TYPE_CD"							<%' Field��(0)%>
	    arrField(1) = "IV_TYPE_NM"							<%' Field��(1)%>
	    
		arrParam(0) = "��������"						<%' �˾� ��Ī %>
		arrParam(1) = "M_IV_TYPE"								<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtIvTypeCd.Value)			<%' Code Condition%>
		'arrParam(3) = Trim(frm1.txtIvTypeNm.Value)			<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "��������"						<%' TextBox ��Ī %>
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	Case 1
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
	
	Case 2
		arrParam(0) = "����ó"							<%' �˾� ��Ī %>
		arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtSupplierCd.Value)		<%' Code Condition%>
		arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "							<%' Where Condition%>
		arrParam(5) = "����ó"							<%' TextBox ��Ī %>
		
	    arrField(0) = "BP_Cd"								<%' Field��(0)%>
	    arrField(1) = "BP_NM"								<%' Field��(1)%>
	    
	    arrHeader(0) = "����ó"							<%' Header��(0)%>
	    arrHeader(1) = "����ó��"						<%' Header��(1)%>
	    
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

	arrParam(0) = arrParam(5)												' �˾� ��Ī	

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtIvTypeCd.focus
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function

'------------------------------------  SetConSItemDC()  ------------------------------------------------
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			.txtIvTypeCd.Value 		= arrRet(0)
			.txtIvTypeNm.Value 		= arrRet(1)
			.txtIvTypeCd.focus
		Case 1
			.txtGroupCd.Value		= arrRet(0)		
			.txtGroupNm.Value		= arrRet(1)		
			.txtGroupCd.focus
		Case 2
			.txtSupplierCd.Value    = arrRet(0)		
			.txtSupplierNm.Value    = arrRet(1)		
			.txtSupplierCd.focus
		End Select
	End With
	Set gActiveElement = document.activeElement
End Function

'==========================================  3.1.1 Form_Load()  =========================================
Sub Form_Load()
	Call LoadInfTB19029							
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")       
	Call InitVariables							
	Call SetDefaultVal	
	Call InitSpreadSheet()	
    Call MM_preloadimages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
	
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================  OpenSortPopup()  =============================================
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

'=========================================  3.3.1 vspdData_DblClick()  ==================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or frm1.vspdData.MaxRows = 0 Then 
		Exit Function
	End If
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================  3.3.2 vspdData_KeyPress()  ===================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1������ frm1���� 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If		
		
	If CheckRunningBizProcess = True Then
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

'========================================================================================================
'   Event Name : OCX_DbClick()
'========================================================================================================
Sub txtFrIvDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrIvDt.Action = 7
		Call SetFocusToDocument("P")	
		frm1.txtFrIvDt.focus	
	End If
End Sub

Sub txtToIvDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToIvDt.Action = 7	
		Call SetFocusToDocument("P")	
		frm1.txtToIvDt.focus		
	End If
End Sub

'=======================================================================================================
'   Event Name : OCX_KeyDown()
'=======================================================================================================
Sub txtFrIvDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

Sub txtToIvDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
Function FncQuery() 
 
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables 														'��: Initializes local global variables

	If ValidDateCheck(frm1.txtFrIvDt, frm1.txtToIvDt) = False Then Exit Function


    '13�� �߰� 
    If frm1.rdoPostFlg1.checked = True Then
		frm1.hdtxtRadio.value = ""
	ElseIf frm1.rdoPostFlg2.checked = True Then
		frm1.hdtxtRadio.value = "Y"
	ElseIf frm1.rdoPostFlg3.checked = True Then
		frm1.hdtxtRadio.value = "N"
	End If		    

	If DbQuery = False Then Exit Function									

    FncQuery = True		
    Set gActiveElement = document.activeElement
End Function

'======================================  DbQuery()  ==================================================
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
			strVal = strVal & "&txtSupplier=" & .hdnSupplier.value			'����ó 
			strVal = strVal & "&txtGroup=" & .hdnGroup.value				'���ű׷� 
			strVal = strVal & "&txtIvType=" & .hdnIvType.Value				'�������� 
			strVal = strVal & "&txtFrIvDt=" & .hdnFrDt.Value				'���Ե���� 
			strVal = strVal & "&txtToIvDt=" & .hdnToDt.Value
			strVal = strVal & "&txtRadio=" & Trim(frm1.hdtxtRadio.value) 	'Ȯ������ 
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtSupplier=" & Trim(.txtSupplierCd.value)	'����ó 
			strVal = strVal & "&txtGroup=" & Trim(.txtGroupCd.value)		'���ű׷� 
			strVal = strVal & "&txtIvType=" & Trim(.txtIvTypeCd.Value)		'�������� 
			strVal = strVal & "&txtFrIvDt=" & .txtFrIvDt.text				'���Ե���� 
			strVal = strVal & "&txtToIvDt=" & .txtToIvDt.text
			strVal = strVal & "&txtRadio=" & Trim(frm1.hdtxtRadio.value) 	'Ȯ������ 
		End If				
			
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'��: Next key tag 
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)		    						'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True    

End Function

'========================================  DbQueryOk()  ==================================================
Function DbQueryOk()	    												'��: ��ȸ ������ ������� 

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtIvTypeCd.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc" -->	
</HEAD>
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
						<TD CLASS="TD5" nowrap>��������</TD>
						<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtIvTypeCd" ALT="��������" SIZE=10 MAXLENGTH=5 SIZE=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 0">
											   <INPUT TYPE=TEXT NAME="txtIvTypeNm" ALT="��������" SIZE=20 tag="14X"></TD>
					   	<TD CLASS="TD5" NOWRAP>���Ե����</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr NOWRAP>
									<td NOWRAP>
										<script language =javascript src='./js/m5111pa1_fpDateTime2_txtFrIvDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/m5111pa1_fpDateTime2_txtToIvDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>	
					<TR>			
						<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="���ű׷�" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 1">
										       <INPUT TYPE=TEXT AlT="���ű׷�" NAME="txtGroupNm" SIZE=20 tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>����ó</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="����ó" NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSpplCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 2">
					   			 			   <INPUT TYPE=TEXT AlT="����ó" Name="txtSupplierNm" tag="14X"></TD>
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
						<script language =javascript src='./js/m5111pa1_vspdData_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp; <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>&nbsp;
									  <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT>  <IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>&nbsp;
							          <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>&nbsp;&nbsp;</TD>
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
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdtxtRadio" TAG="14">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 sr                                                                      