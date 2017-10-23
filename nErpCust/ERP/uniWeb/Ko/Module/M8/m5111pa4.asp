<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m5111pa1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open IvNo Popup ASP														*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/21																			*
'*  8. Modified date(Last)  : 2001/07/06																*
'*  9. Modifier (First)     : Shin Jin Hyen																			*
'* 10. Modifier (Last)      : Ma Jin Ha															*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 																*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>���Թ�ȣ</TITLE>
<!--
'********************************************  1.1 Inc ����  ********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>
<!--
'============================================  1.1.2 ���� Include  ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/incImage.js"></SCRIPT>

<Script Language="VBS">
Option Explicit	
	

Dim arrParent
					
arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)


Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

<%
	Const BIZ_PGM_QRY_ID = "m5111pb1.asp"
%>
	Const BIZ_PGM_QRY_ID = "m5111pb1.asp"			<% '��: �����Ͻ� ���� ASP�� %>

	Const C_IvNo 		= 1
	Const C_IvTypeCd 	= 2
	Const C_IvTypeNm 	= 3
	Const C_ApPostFlg 	= 4
	Const C_SpplCd 		= 5
	Const C_SpplNm 		= 6
	Const C_IvAmt 		= 7
	Const C_Cur 		= 8
	Const C_IvDt 		= 9
	Const C_GrpCd 		= 10
	Const C_GrpNm 		= 11
	
	Dim arrReturn						<% '--- Return Parameter Group %>
	Dim lgIntGrpCount					<% '��: Group View Size�� ������ ���� %>
	Dim lgStrPrevKey
	Dim lgBlnFlgChgValue				'��: Variable is for Dirty flag
	Dim lgIntFlgMode					'��: Variable is for Operation Status

	Dim IsOpenPop						' Popup
	Dim ivType							'�����Ϲ� ���� ���ܸ��� ������ ���� 


'==========================================  2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
	frm1.txtFrIvDt.Text = StartDate
	frm1.txtToIvDt.Text = EndDate
End Sub
'==========================================  2.1.1 InitVariables()  =====================================
Function InitVariables()
	lgIntFlgMode = PopupParent.OPMD_CMODE								<%'��: Indicates that current mode is Create mode%>
	lgIntGrpCount = 0										<%'��: Initializes Group View Size%>
	lgStrPrevKey = ""										<%'initializes Previous Key%>
		
	Self.Returnvalue = Array("")
End Function
		
'=============================== 2.1.2 LoadInfTB19029() ========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>
End Sub
'==========================================  2.2.3 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()
	
	ggoSpread.Source = frm1.vspdData
		
	With frm1.vspdData

		.ReDraw = False
		.OperationMode = 3

		.MaxCols = C_GrpNm + 1
		.Col = C_GrpNm + 1
		.ColHidden = True
		.MaxRows = 0

		ggoSpread.SpreadInit
				
		ggoSpread.SSSetEdit		C_IvNo, "���Թ�ȣ", 20
		ggoSpread.SSSetEdit		C_IvTypeCd, "��������", 16
		ggoSpread.SSSetEdit		C_IvTypeNm, "�������¸�", 20
		ggoSpread.SSSetEdit		C_ApPostFlg, "Ȯ������", 20
		ggoSpread.SSSetEdit		C_SpplCd, "����ó", 10
		ggoSpread.SSSetEdit		C_SpplNm, "����ó��", 20 
		SetSpreadFloat			C_IvAmt, "���Աݾ�", 15, 1,2
		ggoSpread.SSSetEdit		C_Cur,"ȭ��", 10,2
		ggoSpread.SSSetDate		C_IvDt, "������", 15, 2, PopupParent.gDateFormat
		ggoSpread.SSSetEdit		C_GrpCd, "���ű׷�", 18
		ggoSpread.SSSetEdit		C_GrpNm, "���ű׷��", 20
			
		Call SetSpreadLock 
			
		.ReDraw = True
	End With
		
End Sub
	
'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	With frm1.vspdData	
		Redim arrReturn(.MaxCols - 1)
		If .MaxRows > 0 Then 
			.Row = .ActiveRow
			.Col = C_IvNo
			arrReturn(0) = .Text
		end if
	End With
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
	
'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Self.Close()
End Function
'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

'===========================================  OpenSupplier()  =================================
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"									<%' �˾� ��Ī %>
	arrParam(1) = "B_BIZ_PARTNER"								<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)				<%' Code Condition%>
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "	<%' Where Condition%>
	arrParam(5) = "����ó"									<%' TextBox ��Ī %>
	
    arrField(0) = "BP_Cd"										<%' Field��(0)%>
    arrField(1) = "BP_NM"										<%' Field��(1)%>
    
    arrHeader(0) = "����ó"									<%' Header��(0)%>
    arrHeader(1) = "����ó��"								<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		lgBlnFlgChgValue = True
		frm1.txtSupplierCd.focus
		Set gActiveElement = document.activeElement
	End If	
End Function

'===========================================  OpenIvType()  =================================
Function OpenIvType()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	
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
	
    If arrRet(0) <> "" Then
		frm1.txtIvTypeCd.Value = arrRet(0)
		frm1.txtIvTypeNm.Value = arrRet(1)
    End If
    frm1.txtIvTypeCd.focus
	Set gActiveElement = document.activeElement
End Function

'===========================================  OpenGroup()  =================================
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

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
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)		
		frm1.txtGroupNm.Value= arrRet(1)
		frm1.txtGroupCd.focus
	End If	
	Set gActiveElement = document.activeElement
End Function 

'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()
	
	Call LoadInfTB19029							
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")       
	Call SetDefaultVal
	Call InitVariables							
	Call InitSpreadSheet()	
	Call MM_preloadimages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
	
End Sub
'========================================== OCX_EVENT  ====================================
Sub txtFrIvDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrIvDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtFrIvDt.focus
	End if
End Sub

Sub txtToIvDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToIvDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtToIvDt.focus
	End if
End Sub
'=======================================  vspdData_DblClick()  ================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	
	With frm1.vspdData 
		If .MaxRows > 0 Then
			If .ActiveRow = Row Or .ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End With
End Function


'=======================================  vspdData_TopLeftChange()  ================================
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
'=======================================  vspdData_KeyPress()  ================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function
'=======================================  txtFrIvDt_Keypress()  ================================
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

'=======================================  FncQuery()  ============================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        <%'��: Processing is NG%>
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData    
    Call InitVariables															<%'��: Initializes local global variables%>
    
	with frm1
        If CompareDateByFormat(.txtFrIvDt.text,.txtToIvDt.text,.txtFrIvDt.Alt,.txtToIvDt.Alt, _
                   "970025",.txtFrIvDt.UserDefinedFormat,PopupParent.gComDateType,False) = False And Trim(.txtFrIvDt.text) <> "" And Trim(.txtToIvDt.text) <> "" Then
           Call DisplayMsgBox("17a003","X","������","X")	
    
           Exit Function
        End if  	
	End with
   
    '13�� �߰� 
    If frm1.rdoPostFlg1.checked = True Then
		frm1.hdtxtRadio.value = ""
	ElseIf frm1.rdoPostFlg2.checked = True Then
		frm1.hdtxtRadio.value = "Y"
	ElseIf frm1.rdoPostFlg3.checked = True Then
		frm1.hdtxtRadio.value = "N"
	End If		    
    
    
    If DbQuery = False Then Exit Function	
       
    FncQuery = True																<%'��: Processing is OK%>
    Set gActiveElement = document.activeElement    
End Function	
	
'=======================================  DbQuery()  ============================================
Function DbQuery() 
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    
    DbQuery = False                                                         <%'��: Processing is NG%>
    
    Dim strVal
    
    with frm1
    	
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
	    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtSupplier=" & .hdnSupplier.value
		strVal = strVal & "&txtGroup=" & .hdnGroup.value
		strVal = strVal & "&txtIvType=" & .hdnIvType.Value
		strVal = strVal & "&txtFrIvDt=" & .hdnFrDt.Value
		strVal = strVal & "&txtToIvDt=" & .hdnToDt.Value
		strVal = strVal & "&txtRadio=" & Trim(frm1.hdtxtRadio.value) '13�� �߰�	
	else
	    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtSupplier=" & Trim(.txtSupplierCd.value)
		strVal = strVal & "&txtGroup=" & Trim(.txtGroupCd.value)
		strVal = strVal & "&txtIvType=" & Trim(.txtIvTypeCd.Value)
		strVal = strVal & "&txtFrIvDt=" & .txtFrIvDt.text
		strVal = strVal & "&txtToIvDt=" & .txtToIvDt.text
		strVal = strVal & "&txtRadio=" & Trim(frm1.hdtxtRadio.value) '13�� �߰�	
	End if

	end with
	
    if LayerShowHide(1) = false then
		exit function
	end if
    	
	Call RunMyBizASP(MyBizASP, strVal)										<%'��: �����Ͻ� ASP �� ���� %>
	
    DbQuery = True                                                          <%'��: Processing is NG%>

End Function

'=======================================  DbQueryOk()  ============================================
Function DbQueryOk()
	lgIntFlgMode = PopupParent.OPMD_UMODE
	Frm1.vspdData.Focus
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
						<TD CLASS="TD6" nowrap><INPUT TYPE=TEXT NAME="txtIvTypeCd" ALT="��������" SIZE=10 MAXLENGTH=5 SIZE=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvType()">
											   <INPUT TYPE=TEXT NAME="txtIvTypeNm" ALT="��������" SIZE=20 tag="14X"></TD>
					   	<TD CLASS="TD5" NOWRAP>������</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr NOWRAP>
									<td NOWRAP>
										<script language =javascript src='./js/m5111pa4_fpDateTime2_txtFrIvDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/m5111pa4_fpDateTime2_txtToIvDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>	
					<TR>			
						<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="���ű׷�" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
										       <INPUT TYPE=TEXT AlT="���ű׷�" NAME="txtGroupNm" SIZE=20 tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>����ó</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="����ó" NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSpplCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
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
						<script language =javascript src='./js/m5111pa4_vspdData_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
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
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdtxtRadio" TAG="14">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>

