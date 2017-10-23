<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : A_VAT
'*  3. Program ID		    : A6104BA
'*  4. Program Name         : ���ڼ��ݰ�꼭�ϰ��ݿ� 
'*  5. Program Desc         : ���ڼ��ݰ�꼭�ϰ��ݿ� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/04/27
'*  8. Modified date(Last)  : 2002/08/28
'*  9. Modifier (First)     : Jong Hwan, Kim
'* 10. Modifier (Last)      : hersheys
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit

'==========================================================================================================

Const BIZ_PGM_ID = "a6104bb1.asp"											 '��: �����Ͻ� ���� ASP�� 
 '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Dim lgBlnFlgConChg				'��: Condition ���� Flag
Dim lgBlnFlgChgValue
Dim lgIntGrpCount
Dim lgIntFlgMode


'========================================================================================================= 
Dim lgMpsFirmDate, lgLlcGivenDt

Dim lgCurName()					'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
'Dim cboOldVal          
Dim IsOpenPop          


'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False
    lgMpsFirmDate=""
    lgLlcGivenDt=""
End Sub

'=============================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "A","NOCOOKIE","MA") %>
End Sub


'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	Dim svrDate

	svrDate =  UNIDateClientFormat("<%=GetSvrDate%>")
	
	frm1.txtFromIssuedDt.text = UNIGetFirstDay(svrDate, Parent.gDateFormat)
	frm1.txtToIssuedDt.text   = svrDate
End Sub


'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: VAT ó�� �Լ� 
'######################################################################################################### 
Function executeSP()
	Dim RetFlag
	Dim intRetCD

    Err.Clear                                                               '��: Protect system from crashing
    
    Dim strVal
    
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
    If UniConvDateToYYYYMMDD(frm1.txtFromIssuedDt.text, Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtToIssuedDt.text, Parent.gDateFormat,"")Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'��: "Will you destory previous data"
		Exit Function
    End If

	RetFlag = DisplayMsgBox("900018", Parent.VB_YES_NO, "X", "X")   '�� �ٲ�κ� 
	'RetFlag = Msgbox("�۾��� ���� �Ͻðڽ��ϱ�?", vbOKOnly + vbInformation, "����")
	If RetFlag = VBNO Then
		Exit Function
	End IF

	Call LayerShowHide(1)
	
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0002							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtFromDt=" & Trim(frm1.txtFromIssuedDt.text)		'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtToDt="   & Trim(frm1.txtToIssuedDt.text)			'��: ��ȸ ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
End Function


'===============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()


    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call SetToolbar("10000000000011")    
    frm1.txtFromIssuedDt.focus 
End Sub


'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'=======================================================================================================
'   Event Name : txtFromIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromIssuedDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromIssuedDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtFromIssuedDt.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtToIssuedDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToIssuedDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToIssuedDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtToIssuedDt.Focus
    End If
End Sub


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Call Parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function


'******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' ���� ���� --></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ڼ��ݰ�꼭�ϰ��ݿ�</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
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
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5">������</TD>
								<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFromIssuedDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="������" id=fpDateTime1></OBJECT>');</SCRIPT> ~ 
											    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToIssuedDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="������" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>
                                    <font color="red">* ����<br>
                                                      &nbsp;&nbsp;�� ȭ���� ��ť�� ���� ��ü�� ��� �����մϴ�.<br>
                                                      &nbsp;&nbsp;��ť�� ������ ��ü�� ������ �Ұ� �մϴ�.<br>
                                                      &nbsp;&nbsp;��ť���� ���� ����û ������ �Ϸ�� �ǿ� ���ؼ��� ����˴ϴ�.
                                    </font>
                                </TD>
							</TR>
							<TR>
								<TD CLASS=TD5>&nbsp;</TD>
								<TD CLASS=TD6>&nbsp;</TD>
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
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnExec" CLASS="CLSMBTN" OnClick="VBScript:Call executeSP()" Flag=1>�� ��</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tabindex="-1" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

