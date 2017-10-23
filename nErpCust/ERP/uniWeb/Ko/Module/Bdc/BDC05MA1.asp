<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : DBC
'*  2. Function Name        : ��ġ�۾� ����ȸ 
'*  3. Program ID           : BDC05MA1.ASP
'*  4. Program Name         : �۾�����ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005.02.07
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kweon, Soon Tae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
'��: indicates that All variables must be declared in advance 
<!-- #Include file="../../inc/lgvariables.inc" -->
'==========================================================================================================
' ��� �� ���� ���� 
'----------------------------------------------------------------------------------------------------------
Const BIZ_PGM_ID = "BDC05MB1.ASP"										'��: ��ȸ �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMP_ID = "BDC04MA1"

Dim lgRetFlag
Dim IsOpenPop
Dim iColSep, iRowSep
Dim lgOldRow

Dim C_SP1_SEQ
Dim C_SP1_TIM
Dim C_SP1_RES
Dim C_SP1_COM
Dim C_SP1_MTH

Dim strMode

Dim szCurMth
Dim szCurPrm
Dim szCurJon

'==========================================================================================================
' ������ �ε尡 �Ϸ�Ǹ� �ڵ����� ȣ��Ǵ� �Լ�.
' �ʱ�ȭ ��ƾ�� �̰��� ���߽��� �־�� ��.
' ../../inc/incCliMAMain.vbs ���Ͽ� �� �Լ��� ȣ�� �ϵ��� �ϴ� ����� �ֽ� 
'----------------------------------------------------------------------------------------------------------
Sub Form_Load()

	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")
	Call InitSpreadSheet
	Call InitVariables
	Call InitComboBox
	Call InitGridComboBox
	Call SetToolbar("1100000000000111")
	
	If parent.ReadCookie("txtJobId") <> "" Then

		Call SetCookieVal
	End If
	frm1.txtJobID.focus
	
	
End Sub

'=========================================================================================================
Function FncCancel()
    ggoSpread.EditUndo
End Function

'==========================================================================================================
' �ý��ۿ� ������ ȭ�����, ����ڵ�, ������ �������� �ʱ�ȭ �ϴ� �Լ�.
' ../../inc/incCliVariables.vbs �� ../../ComAsp/LoadInfTB19029.asp  ���Ͽ� �������̴�.
'----------------------------------------------------------------------------------------------------------
Sub LoadInfTB19029()
<% Call loadInfTB19029A("I", "*","NOCOOKIE", "MA") %>
End Sub

'==========================================================================================================
' �������� �ʱ�ȭ �Լ� 
' ���α׷��� ���� ����ڵ��� ������ �־�� �ϴ� �κ� 
'----------------------------------------------------------------------------------------------------------
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()

	With frm1.vspdData
        .ReDraw = False
		.RowHeadersShow = True
		.MaxCols = C_SP1_MTH
        .MaxRows = 0

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20041121", , parent.gAllowDragDropSpread
		
		Call GetSpreadColumnPos()
		
		ggoSpread.SSSetEdit   C_SP1_SEQ,  "����", 6, , , 3
		ggoSpread.SSSetEdit   C_SP1_TIM,  "ó���ð�", 20, , , 26
		ggoSpread.SSSetEdit   C_SP1_RES,  "���", 6, , , 6
		ggoSpread.SSSetEdit   C_SP1_COM,  "�������", 20, , , 40
		ggoSpread.SSSetEdit   C_SP1_MTH,  "��������", 40, , , 200
		
		ggoSpread.SSSetSplit2(1)
		.ReDraw = True
	End With
	
	Call SetSpreadLock()
	
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()	
	
	' Grid 1(vspdData) - Operation 
	C_SP1_SEQ = 1
	C_SP1_TIM = 2
	C_SP1_RES = 3
	C_SP1_COM = 4
	C_SP1_MTH = 5

End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos()
 	Dim iCurColumnPos

 	ggoSpread.Source = frm1.vspdData
 			
 	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 	
 	C_SP1_SEQ = iCurColumnPos(1)
	C_SP1_TIM = iCurColumnPos(2)
	C_SP1_RES = iCurColumnPos(3)
	C_SP1_COM = iCurColumnPos(4)
	C_SP1_MTH = iCurColumnPos(5)
	
End Sub				


'==========================================================================================================
' ���� �������� �ʱ�ȭ ��Ų��.
'----------------------------------------------------------------------------------------------------------
Sub InitVariables()
	Dim i, j

    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

'========================================  2.2.1 SetCookieVal()  ======================================
'	Name : SetCookieVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=================================================================================================== 
Sub SetCookieVal()
   	
	frm1.txtJobID.value	= ReadCookie("txtJobId")
    call dbQuery()
	WriteCookie "txtJobId", ""
		
End Sub

'==========================================================================================================
' ���������Ʈ �̿��� �޺��ڽ����� �ʱ�ȭ �Ѵ�.
'----------------------------------------------------------------------------------------------------------
Sub InitComboBox()
End Sub

'==========================================================================================================
' �������� ��Ʈ�� �޺��ڽ��� ���� �ʱ�ȭ �Ѵ�.
'----------------------------------------------------------------------------------------------------------
Sub InitGridComboBox()
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey			'Sort in Descending
 			lgSortKey = 1
 		End If
 		
 		lgOldRow = Row
 		
	Else
 		'------ Developer Coding part (Start)
 		If lgOldRow <> Row Then		
			frm1.vspdData.Row = row
		
			lgOldRow = Row
		
		End If		
	 	'------ Developer Coding part (End)
	
 	End If
	
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub


'=========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)        
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
  
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos()
End Sub 

'==========================================================================================================
' �����ڵ� ���� �˾� â�� ������Ų��.
'----------------------------------------------------------------------------------------------------------
Function OpenPopup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrParam(0) = "�۾��˾�"				' �˾� ��Ī 
	arrParam(1) = "B_BDC_JOBS"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtJobID.Value)
	arrParam(3) = ""
	arrParam(4) = " job_state= " & Filtervar("D", "''", "S")						' Code Condition
	arrParam(5) = "����"

	arrField(0) = "JOB_ID"					' Field��(0)
	arrField(1) = "JOB_TITLE"				' Field��(1)

	arrHeader(0) = "�۾��ڵ�"				' Header��(0)
	arrHeader(1) = "�۾���"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
	                                Array(arrParam, arrField, arrHeader), _
		                            "dialogWidth=420px; dialogHeight=450px; " & _
		                            "center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtJobID.focus
		Exit Function
	Else
		frm1.txtJobID.value = arrRet(0)
		frm1.txtJobNm.value = arrRet(1)
		frm1.txtJobID.focus
	End If
End Function

'==========================================================================================================
' �޴����� ��ȸ ��ư�� �������� ȣ��Ǵ� �޼��� �ڵ鷯�̴�.
' �����μ�:
'----------------------------------------------------------------------------------------------------------
Function FncQuery()
    Dim IntRetCD 
    FncQuery = False
    Err.Clear

    Call ggoSpread.ClearSpreadData()
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then
		Exit Function
    End If

    If DbQuery = False Then
       Exit Function
    End If
       
    FncQuery = True
End Function

'==========================================================================================================
' �޴����� ��ȸ ��ư�� �������� ȣ��Ǵ� �޼��� �ڵ鷯�̴�.
' �����μ�:
'----------------------------------------------------------------------------------------------------------
Function DbQuery() 
    Dim strVal    
    Dim IntRetCD

    DbQuery = False

    Call LayerShowHide(1)
    With frm1
        strVal = BIZ_PGM_ID & _
                "?txtMode=" & Parent.UID_M0001 & _
                "&txtJobId=" & Trim(.txtJobId.value) & _
                "&txtMaxRows=" & .vspdData.MaxRows & _
                "&lgStrPrevKey=" & lgStrPrevKey
        Call RunMyBizASP(MyBizASP, strVal)
    End With

    DbQuery = True
End Function

'==========================================================================================================
' ��ȸ �۾��� �Ϸ� �Ǿ��� �� �ڽ� �����ӿ� ���� ȣ��ȴ�.
' �����μ�:
'----------------------------------------------------------------------------------------------------------
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE

    Call ggoOper.LockField(Document, "Q")
End Function

'==========================================================================================================
' �޴����� ���� ��ư�� �������� ȣ��Ǵ� �޼��� �ڵ鷯�̴�.
' �����μ�:
' ����: ����ڰ� �Է��� ���� �����Ѵ�.
'----------------------------------------------------------------------------------------------------------
Function FncSave() 
End Function

'==========================================================================================================
' ���� �۾� �������� ������ ȣ��Ǵ� �޼��� �ڵ鷯�̴�.
' �����μ�:
' ��    ��: ����� ���� ������ ���� ��� ���� ������������ �����.
'----------------------------------------------------------------------------------------------------------
Function FncExit()
    Dim IntRetCD
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    FncExit = True
End Function



'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    
	Dim LngRow
	 
    ggoSpread.Source = gActiveSpdSheet

    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()  
    
	Call ggoSpread.ReOrderingSpreadData

End Sub 

Function JumpJobRun()

    Dim IntRetCd, strVal
    
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	WriteCookie "txtJobId", UCase(Trim(frm1.txtJobID.value))
	
	PgmJump(BIZ_PGM_JUMP_ID)
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>�۾�����ȸ</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*" align=right>&nbsp;</td>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>�۾��ڵ�</TD>
									<TD CLASS="TD656" NOWRAP>
									    <INPUT NAME="txtJobID" MAXLENGTH="18" SIZE=18 ALT ="�۾��ڵ�" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcID" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup()">
										<INPUT NAME="txtJobNm" MAXLENGTH="80" SIZE=50 ALT ="�۾���" tag="14X"  STYLE="TEXT-ALIGN:left"></TD>
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
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=8>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
                                 </TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>   
        <TD WIDTH=100%>   
            <TABLE <%=LR_SPACE_TYPE_30%>>   
                <TR>   
                    <TD>	
						<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
							  <TD WIDTH=10>&nbsp;</TD>
							  <TD align="left">
							  </TD>
							  <TD WIDTH=* Align=right><A href="vbscript:JumpJobRun">�۾�����</A> </TD>
							  <TD WIDTH=10>&nbsp;</TD>
							</TR>
						</TABLE>
					</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
</HTML>

