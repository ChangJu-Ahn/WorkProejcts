<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : �ڷ� ����ȭ
'*  2. Function Name        : ������ǰ������ �۽���Ȳ(S)
'*  3. Program ID           : XI214MA1_KO441
'*  4. Program Name         : ������ǰ������ �۽���Ȳ(S)
'*  5. Program Desc         : ������ǰ������ �۽���Ȳ(S)
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2006/04/28
'*  8. Modified date(Last)  : 2006/04/28
'*  9. Modifier (First)     : �Ǽ���
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              : 
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit				'��: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim lgIsOpenPop                                             <%'��: Popup status                          %> 
Dim lgTypeCD                                                <%'��: 'G' is for group , 'S' is for Sort    %>
Dim lgFieldCD                                               <%'��: �ʵ� �ڵ尪                           %>
Dim lgFieldNM                                               <%'��: �ʵ� ����                           %>
Dim lgFieldLen                                              <%'��: �ʵ� ��(Spreadsheet����)              %>
Dim lgFieldType                                             <%'��: �ʵ� ����                           %>
Dim lgDefaultT                                              <%'��: �ʵ� �⺻��                           %>
Dim lgNextSeq                                               <%'��: �ʵ� Pair��                           %>
Dim lgKeyTag                                                <%'��: Key ����                              %>
Dim lgNextSeq_T                                             <%'��: �ʵ� Pair��                           %>
Dim lgKeyTag_T                                              <%'��: Key ����                              %>
Dim lgSortTitleNm                                           <%'��: Orderby popup�� ����Ÿ(�ʵ弳��)      %>
Dim lgSortFieldCD1                                          <%'��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      %>
Dim lgMark                                                  <%'��: ��ũ                                  %>
Dim lgKeyPos                                                <%'��: Key��ġ                               %>
Dim lgKeyPosVal                                             <%'��: Key��ġ Value                         %>
Dim IsOpenPop
Dim arrParam

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID = "XI214MB1_KO441.asp"	
Const BIZ_PGM_JUMP_ID = "B3B30MA1_KO119"				      '��: Jump ASP�� 

Const C_MaxKey = 2																'�١١١�: Max key value

<% '----------------  ���� Global ������ ����  ----------------------------------------------------------- %>
Dim lsConcd
Dim lsConNm
Dim StartDate
Dim LastDate

Dim ITEM_CD
Dim ITEM_NM
Dim BP_CODE
Dim BP_NAME
Dim BP_ITCD
Dim BP_ITNM
Dim BP_SPEC
Dim BP_PRSP
Dim CM_FLAG
Dim CM_SEND
Dim CM_UPDT
Dim CM_ERRD
Dim CM_RECV

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
'LastDate =  UNIDateAdd("m", 1, StartDate, Parent.gDateFormat)
LastDate = StartDate

<% '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ %>

'========================================================================================================= 
Sub InitVariables()
   lgBlnFlgChgValue = False                               'Indicates that no value changed
   lgStrPrevKey     = ""                                  'initializes Previous Key
   lgPageNo         = ""
   lgSortKey        = 1
   lgIntFlgMode = parent.OPMD_CMODE	
End Sub

'========================================================================================================
Sub SetDefaultVal()
	frm1.txtFrDt.text = StartDate
	frm1.txtToDt.text = LastDate
	frm1.txtBp_CD.focus
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp, arrVal

	Call vspdData_Click(frm1.vspdData.ActiveCol, frm1.vspdData.ActiveRow)

	If flgs = 1 Then
		WriteCookie "txtItemCd", lsConcd
		WriteCookie "txtBpCd", lsConnm

	ElseIf flgs = 0 Then

		strTemp =  ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function

		If Err.number <> 0 Then
			Err.Clear
			 WriteCookie CookieSplit , ""
			Exit Function 
		End If

		 WriteCookie CookieSplit , ""
		
		Call MainQuery()
	End If

End Function

'========================================================================================================
Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData  
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.Spreadinit "V20021206",,parent.gAllowDragDropSpread

		.ReDraw = false

		.MaxCols = CM_RECV + 1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols														'��: ����� �� Hidden Column
		.ColHidden = True

		.MaxRows = 0
		ggoSpread.ClearSpreadData

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit  ITEM_CD, "ǰ���ڵ�",		15, , ,10
		ggoSpread.SSSetEdit  ITEM_NM, "ǰ �� ��",		20, , ,50
		ggoSpread.SSSetEdit  BP_CODE, "�����ڵ�",		10, , ,10
		ggoSpread.SSSetEdit  BP_NAME, "�����",		20, , ,50
		ggoSpread.SSSetEdit  BP_ITCD, "����ǰ���ڵ�",	15, , ,10
		ggoSpread.SSSetEdit  BP_ITNM, "����ǰ���",	20, , ,50
		ggoSpread.SSSetEdit  BP_SPEC, "����԰�",		20, , ,50
		ggoSpread.SSSetEdit  BP_PRSP, "������±԰�",	20, , ,50
		ggoSpread.SSSetEdit  CM_FLAG, "��������",		 8, , , 1
		ggoSpread.SSSetEdit  CM_SEND, "�����ۼ����Ͻ�", 18, , ,20
		ggoSpread.SSSetEdit  CM_UPDT, "MES���ſ���",		12, , , 1
		ggoSpread.SSSetEdit  CM_ERRD, "��������",		25, , ,50
		ggoSpread.SSSetEdit  CM_RECV, "MES���������Ͻ�",18, , ,20

		.ReDraw = true

		Call SetSpreadLock 
    End With
End Sub

'========================================================================================================
Sub InitSpreadPosVariables()
    ITEM_CD = 1
    ITEM_NM = 2
    BP_CODE = 3
    BP_NAME = 4
    BP_ITCD = 5
    BP_ITNM = 6
    BP_SPEC = 7
    BP_PRSP = 8
    CM_FLAG = 9
    CM_SEND = 10
    CM_UPDT = 11
    CM_ERRD = 12
    CM_RECV = 13

End Sub

'========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			ITEM_CD = iCurColumnPos(1)
			ITEM_NM = iCurColumnPos(2)
			BP_CODE = iCurColumnPos(3)
			BP_NAME = iCurColumnPos(4)
			BP_ITCD = iCurColumnPos(5)
			BP_ITNM = iCurColumnPos(6)
			BP_SPEC = iCurColumnPos(7)
			BP_PRSP = iCurColumnPos(8)
			CM_FLAG = iCurColumnPos(9)
			CM_SEND = iCurColumnPos(10)
			CM_UPDT = iCurColumnPos(11)
			CM_ERRD = iCurColumnPos(12)
			CM_RECV = iCurColumnPos(13)
			
    End Select
End Sub

'========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================================
Sub SetSpreadColor(ByVal lRow)
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected  ITEM_CD, lRow, lRow
		ggoSpread.SSSetProtected  ITEM_NM, lRow, lRow
		ggoSpread.SSSetProtected  BP_CODE, lRow, lRow
		ggoSpread.SSSetProtected  BP_NAME, lRow, lRow
		ggoSpread.SSSetProtected  BP_ITCD, lRow, lRow
		ggoSpread.SSSetProtected  BP_ITNM, lRow, lRow
		ggoSpread.SSSetProtected  BP_SPEC, lRow, lRow
		ggoSpread.SSSetProtected  BP_PRSP, lRow, lRow
		ggoSpread.SSSetProtected  CM_FLAG, lRow, lRow
		ggoSpread.SSSetProtected  CM_SEND, lRow, lRow
		ggoSpread.SSSetProtected  CM_UPDT, lRow, lRow
		ggoSpread.SSSetProtected  CM_ERRD, lRow, lRow
		ggoSpread.SSSetProtected  CM_RECV, lRow, lRow
		
		.vspdData.ReDraw = True
    End With
End Sub

'========================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


	If iWhere = 0 Then
		arrParam(0) = "����"								<%' �˾� ��Ī %>
		arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtBp_cd.value)				<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "����"								<%' TextBox ��Ī %>
		
		arrField(0) = "BP_CD"								<%' Field��(0)%>
		arrField(1) = "BP_NM"								<%' Field��(1)%>
	    
		arrHeader(0) = "����"								<%' Header��(0)%>
		arrHeader(1) = "�����Ī"							<%' Header��(1)%>

		frm1.txtBp_cd.focus 
	Else
		arrParam(0) = "ǰ  ��"								<%' �˾� ��Ī %>
		arrParam(1) = "B_ITEM"								<%' TABLE ��Ī %>
		arrParam(2) = Trim(frm1.txtItemCD.value)			<%' Code Condition%>
		arrParam(3) = ""
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "ǰ  ��"								<%' TextBox ��Ī %>
		
		arrField(0) = "ITEM_CD"								<%' Field��(0)%>
		arrField(1) = "ITEM_NM"								<%' Field��(1)%>
	    
		arrHeader(0) = "ǰ���ڵ�"							<%' Header��(0)%>
		arrHeader(1) = "ǰ �� ��"							<%' Header��(1)%>
															<%' Name Cindition%>
		frm1.txtItemCD.focus 
	End If

    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
									Array(arrParam, arrField, arrHeader), _
									"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBpCode(arrRet, iWhere)
	End If	
End Function

'========================================================================================================= 
Function SetBpCode(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then
			.txtBp_cd.value = arrRet(0) 
			.txtBp_nm.value = arrRet(1)   
		Else
			.txtItemCD.value = arrRet(0) 
			.txtItemNM.value = arrRet(1)   
		End If
	End With
End Function

'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029										'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)	
    Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field
		
    '----------  Coding part  -------------------------------------------------------------
	Call InitVariables														    '��: Initializes local global variables
	Call SetDefaultVal		
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")				'��: ��ư ���� ���� 
	Call CookiePage(0)
	frm1.txtBp_cd.focus
End Sub

'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If

    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 1
    lsConcd=frm1.vspdData.Text
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 3
    lsConnm=frm1.vspdData.Text  

End Sub

'==========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True
End Sub

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
End Sub

'========================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    If OldLeft <> NewLeft Then Exit Sub
    
	If CheckRunningBizProcess = True Then  Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ	
    	If lgStrPrevKey <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFrDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFrDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromReqrdDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToReqrdDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
Function FncQuery() 
	
    FncQuery = False											'��: Processing is NG
    
    If Trim(frm1.txtFrDt.Text) > Trim(frm1.txtToDt.Text) Then
		MsgBox "�������� ������ �����̾�� �մϴ�."
        Exit Function
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'��: Clear Contents  Field
    Call InitVariables 											'��: Initializes local global variables
'    Call SetDefaultVal

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							'��: This function check indispensable field
       Exit Function
    End If
	
    '-----------------------
    'Radio Button Check area
    '-----------------------
	If frm1.rdoUsage_flag1.checked = True Then
		frm1.txtRadioFlag.value  = "" 
	ElseIf frm1.rdoUsage_flag2.checked = True Then
		frm1.txtRadioFlag.value = "Y" 
	ElseIf frm1.rdoUsage_flag3.checked = True Then
		frm1.txtRadioFlag.value = "N" 
	End If

    '-----------------------
    'Query function call area
    '------------------------
    Call DbQuery												'��: Query db data

    FncQuery = True												'��: Processing is OK
End Function

'========================================================================================
Function FncPrint()
    ggoSpread.Source = frm1.vspdData
	Call parent.FncPrint()
End Function

'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_MULTI, False)
End Function

'========================================================================================
Function FncSplitColumn()
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
	iColumnLimit  = frm1.vspdData.MaxCols

	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
       iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
       Exit Function  
    End If   
    
	Frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE

	ggoSpread.Source = frm1.vspdData

	ggoSpread.SSSetSplit(ACol)    

	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0    
	frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
End Function

'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO,"x","x")   'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	FncExit = True
End Function

'========================================================================================
Function DbQuery() 

    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey      

    DbQuery = False

	If LayerShowHide(1) = False Then
		Exit Function 
	End If

	Dim strVal

    With frm1
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 & _
				 "&txtBp_cd=" & Trim(.txtBp_cd.value) & _
				 "&txtFrDT=" & Trim(.txtFrDt.text) & _
				 "&txtToDT=" & Trim(.txtToDt.text) & _
				 "&txtItemCD=" & Trim(.txtItemCD.value) & _
				 "&txtRadioFlag=" & Trim(.txtRadioFlag.value) & _
				 "&txtMaxRows="   & .vspdData.MaxRows & _
				 "&lgStrPrevKey=" & lgStrPrevKey

		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    End With
   
    DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	lgIntFlgMode = parent.OPMD_UMODE										'Indicates that current mode is Update mode
	lgBlnFlgChgValue = False
	'-----------------------
	'Reset variables area
	'-----------------------
	Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	'����/��ȸ/�Է�
	'/����/����/����In
	'/����Out/���/����
	'/����/����/����
	'/�μ�/ã��
	Call SetToolbar("11000000000111")										'��: ��ư ���� ���� 

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					<TD WIDTH=* align=right></TD>
					<TD WIDTH=10>&nbsp;</TD>	
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
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
									<TD CLASS="TD6" NOWRAP>
										<INPUT NAME="txtBp_cd" TYPE="Text" MAXLENGTH="10" TAG="11XXXU" SIZE="12"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="ImgBp_cd" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenPopUp frm1.txtBp_cd.value, 0"> <INPUT NAME="txtBp_nm" TYPE="Text" TAG="14" SIZE="35">
									</TD>
									<TD CLASS="TD5" NOWRAP>�۽űⰣ</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript>
										ExternalWrite('<OBJECT title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtFrDt CLASSID=<%=gCLSIDFPDT%> tag="12" ALT="������"></OBJECT>');
										</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> 
										ExternalWrite('<OBJECT title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtToDt CLASSID=<%=gCLSIDFPDT%> tag="12" ALT="������"></OBJECT>');
										</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>ǰ ��</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT NAME="txtItemCD" TYPE="Text" MAXLENGTH="18" TAG="11XXXU" SIZE="20"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="ImgItemCD" ALIGN="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenPopUp frm1.txtItemCD, 1"> <INPUT NAME="txtItemNM" TYPE="Text" TAG="14" SIZE="27">
									</TD>
									<TD CLASS="TD5" NOWRAP>���ſ���</TD>
									<TD CLASS="TD6" NOWRAP>
									<input type=radio CLASS="RADIO" name="rdoUsage_flag" id="rdoUsage_flag1" value="" tag = "11" checked>
										<label for="rdoUsage_flag1">��ü</label>
									<input type=radio CLASS = "RADIO" name="rdoUsage_flag" id="rdoUsage_flag2" value="Y" tag = "11">
										<label for="rdoUsage_flag2">����</label>
									<input type=radio CLASS="RADIO" name="rdoUsage_flag" id="rdoUsage_flag3" value="N" tag = "11">
										<label for="rdoUsage_flag3">����</label></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23XXX" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">��ǰ�ڵ�����������</a>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../Blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="HBp_cdFrom" tag="24">
<INPUT TYPE=HIDDEN NAME="HBp_cdTo" tag="24">
<INPUT TYPE=HIDDEN NAME="HRadioFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="HRadioType" tag="24">
			
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1" ></iframe>
</DIV>
</BODY>
</HTML>