<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1413PA1
'*  4. Program Name         : �㺸������ȣ 
'*  5. Program Desc         : �㺸������ȣ 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              : 2000/12/09
'*                            2002/12/11 Include ������� ���ر� 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>�㺸������ȣ</TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" --> 
Public lgIsOpenPop                                             <%'��: Popup status                          %> 
'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s1413pb1.asp"
Const C_MaxKey          = 3                                    '�١١١�: Max key value
Const gstrWarrantTypeMajor = "S0002"
 
Dim arrParent

ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
dim colino

'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1
    lgIntFlgMode     = PopupParent.OPMD_CMODE				'Indicates that current mode is Create mode

End Sub

Sub SetDefaultVal()

	frm1.txtAsingFrDt.text = StartDate
	frm1.txtAsingToDt.text = EndDate
	frm1.vspdData.Focus

End Sub

'===========================================================================================================
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
		<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
		<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "PA") %>
End Sub

'==========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S1413PA1","S","A","V20030320", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	colino = 1
    Call SetSpreadLock 
End Sub


'=========================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
	.vspddata.OperationMode = 3
    .vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Function OpenBizPartner()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
			
	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True
			
	arrParam(0) = "�ŷ�ó"							<%' �˾� ��Ī %>
	arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtBpCd.value)				<%' Code Condition%>
	arrParam(3) = ""									<%' Name Cindition%>
	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				<%' Where Condition%>
	arrParam(5) = "�ŷ�ó"							<%' TextBox ��Ī %>
		
	arrField(0) = "BP_CD"								<%' Field��(0)%>
	arrField(1) = "BP_NM"								<%' Field��(1)%>
		
	arrHeader(0) = "�ŷ�ó"							<%' Header��(0)%>
	arrHeader(1) = "�ŷ�ó��"						<%' Header��(1)%>
			
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
		
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizPartner(arrRet)
	End If
End Function

'=========================================================================================================
Function OpenMinorCd(strMinorCD, strMinorNM, strPopPos, strMajorCd)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = strPopPos								<%' �˾� ��Ī %>
	arrParam(1) = "B_Minor"								<%' TABLE ��Ī %>
	arrParam(2) = Trim(strMinorCD)						<%' Code Condition%>
	arrParam(3) = ""						            <%' Name Cindition%>
	arrParam(4) = "MAJOR_CD= " & FilterVar(strMajorCd, "''", "S") & ""		<%' Where Condition%>
	arrParam(5) = strPopPos								<%' TextBox ��Ī %>

	arrField(0) = "Minor_CD"							<%' Field��(0)%>
	arrField(1) = "Minor_NM"							<%' Field��(1)%>

	arrHeader(0) = strPopPos							<%' Header��(0)%>
	arrHeader(1) = strPopPos & "��"					<%' Header��(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetMinorCd(strMajorCd, arrRet)
	End If
End Function

'=========================================================================================================
Function OpenSalesGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "�����׷�"								<%' �˾� ��Ī %>
	arrParam(1) = "B_SALES_GRP"									<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtSalesGroup.value)						<%' Code Condition%>
	arrParam(3) = ""											<%' Name Cindition%>
	arrParam(4) = ""											<%' Where Condition%>
	arrParam(5) = "�����׷�"								<%' TextBox ��Ī %>

	arrField(0) = "SALES_GRP"									<%' Field��(0)%>
	arrField(1) = "SALES_GRP_NM"										<%' Field��(1)%>

	arrHeader(0) = "�����׷�"								<%' Header��(0)%>
	arrHeader(1) = "�����׷��"								<%' Header��(1)%>
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSalesGroup(arrRet)
	End If
End Function

'=========================================================================================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
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

'=========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'=========================================================================================================
Function SetBizPartner(arrRet)
	frm1.txtBpCd.value = arrRet(0)
	frm1.txtBpNm.value = arrRet(1)
End Function

'=========================================================================================================
Function SetMinorCd(strMajorCd, arrRet)
	frm1.txtWarrentType.Value = arrRet(0)
	frm1.txtWarrentTypeNm.Value = arrRet(1)
End Function

'=========================================================================================================
Function SetSalesGroup(arrRet)
	frm1.txtSalesGroup.Value = arrRet(0)
	frm1.txtSalesGroupNm.Value = arrRet(1)
End Function	


'=========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================================================================================
Sub btnBpCdOnClick()
	frm1.txtBpCd.focus 
	Call OpenBizPartner()
End Sub

'=========================================================================================================
Sub btnSalesGroupOnClick()
	frm1.txtSalesGroup.focus 
	Call OpenSalesGroup()
End Sub

'=========================================================================================================
Sub btnWarrentTypeOnClick()
	frm1.txtWarrentType.focus 
	Call OpenMinorCd(frm1.txtWarrentType.value, frm1.txtWarrentTypeNm.value, "�㺸����", gstrWarrantTypeMajor)
End Sub


'=========================================================================================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End  function


'=========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
	
'=========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKey <> "" Then							<% '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� %>
				DbQuery
			End If
		End If
	End With
End Sub
	
'=========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ	
		If lgStrPrevKey <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DbQuery()
		End If
	End If

End Sub

'=========================================================================================================
Sub txtAsingFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtAsingFrDt.Action = 7
		Call SetFocusToDocument("P")
		Frm1.txtAsingFrDt.Focus
	End If
End Sub

'=========================================================================================================
Sub txtAsingToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtAsingToDt.Action = 7
		Call SetFocusToDocument("P")
		Frm1.txtAsingToDt.Focus
	End If
End Sub

'=========================================================================================================
Sub txtAsingFrDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'=========================================================================================================
Sub txtAsingToDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

'=========================================================================================================
Function OKClick()
		
	dim arrReturn
	If frm1.vspdData.ActiveRow > 0 Then				
		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1) 
		arrReturn = frm1.vspdData.Text
		Self.Returnvalue = arrReturn
	End If

	Self.Close()
End Function

'=========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'=========================================================================================================
Function FncQuery() 

    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", PopupParent.VB_YES_NO, "x", "x")
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
	If ValidDateCheck(frm1.txtAsingFrDt, frm1.txtAsingToDt) = False Then Exit Function

    Call DbQuery															'��: Query db data

    FncQuery = True		
End Function

'=========================================================================================================
Function FncPrint() 
    Call PopupParent.FncPrint()
End Function

'=========================================================================================================
Function FncExcel() 
	Call PopupParent.FncExport(PopupParent.C_MULTI)
End Function

'=========================================================================================================
Function FncFind() 
    Call PopupParent.FncFind(PopupParent.C_MULTI , False)                                     <%'��:ȭ�� ����, Tab ���� %>
End Function

'=========================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", PopupParent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'=========================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If
    
    With frm1
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				<%'��: �����Ͻ� ó�� ASP�� ���� %>
		strVal = strVal & "&txtBpCd=" & Trim(frm1.txtHBpCd.value)	<%'��: ��ȸ ���� ����Ÿ %>
		strVal = strVal & "&txtWarrentType=" & Trim(frm1.txtHWarrentType.value)
		strVal = strVal & "&txtSalesGroup=" & Trim(frm1.txtHSalesGroup.value)
		strVal = strVal & "&txtAsingFrDt=" & Trim(frm1.txtHAsingFrDt.value)
		strVal = strVal & "&txtAsingToDt=" & Trim(frm1.txtHAsingToDt.value)
	Else		
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				<%'��: �����Ͻ� ó�� ASP�� ���� %>
		strVal = strVal & "&txtBpCd=" & Trim(frm1.txtBpCd.value)	<%'��: ��ȸ ���� ����Ÿ %>
		strVal = strVal & "&txtWarrentType=" & Trim(frm1.txtWarrentType.value)
		strVal = strVal & "&txtSalesGroup=" & Trim(frm1.txtSalesGroup.value)
		strVal = strVal & "&txtAsingFrDt=" & Trim(frm1.txtAsingFrDt.text)
		strVal = strVal & "&txtAsingToDt=" & Trim(frm1.txtAsingToDt.text)
	End If
	
        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
       
        Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True


End Function

'=========================================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True
		If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
			lgIntFlgMode = PopupParent.OPMD_UMODE
		End If
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

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
						<TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10  MAXLENGTH=10 TAG="11XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnBpCdOnClick">&nbsp;
							<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 TAG="14">
						</TD>
						<TD CLASS=TD5 NOWRAP>������</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/s1413pa1_fpDateTime2_txtAsingFrDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/s1413pa1_fpDateTime2_txtAsingToDt.js'></script>
						</TD>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>�㺸����</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtWarrentType" TYPE=TEXT SIZE=10  MAXLENGTH="5" TAG="11XXXU" ALT="�㺸����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWarrentType" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnWarrentTypeOnClick">&nbsp;
							<INPUT TYPE=TEXT NAME="txtWarrentTypeNm"  SIZE="20" MAXLENGTH="30" TAG="24"></TD>
						</TD>
						<TD CLASS=TD5 NOWRAP>�����׷�</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10  MAXLENGTH=5 TAG="11XXXU" ALT="�����׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnSalesGroupOnClick">&nbsp;
							<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14">
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
						<script language =javascript src='./js/s1413pa1_vaSpread_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
		<TR HEIGHT=20>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
							                  <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" OnClick="OpenSortPopup()" ></IMG>
											     </TD>
						<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
								                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO NORESIZE framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHBpCd" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHWarrentType" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGroup" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHAsingFrDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHAsingToDt" TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
