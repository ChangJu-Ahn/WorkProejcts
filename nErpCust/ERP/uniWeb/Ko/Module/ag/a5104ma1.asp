
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5101ma1
'*  4. Program Name         : ȸ����ǥ ��� 
'*  5. Program Desc         : ȸ����ǥ������ ���, ����, ����, ��ȸ 
'*  6. Component List       : PAGG020.dll
'*  7. Modified date(First) : 2003/01/02
'*  8. Modified date(Last)  : 2003/10/14
'*  9. Modifier (First)     : Kim Ho Young
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="Acctctrl_ko441_1.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="a5104ma1.vbs"></SCRIPT>

<SCRIPT LANGUAGE=vbscript>

Option Explicit	

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate

iDBSYSDate = "<%=GetSvrDate%>"

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029
    
    Call FormatDATEField(frm1.txtGLDt)
    Call LockObjectField(frm1.txtGLDt,"R")

    Call LockHTMLField(frm1.cboGlType,"R")
    Call LockHTMLField(frm1.cboGlInputType,"P")            

    Call FormatDoubleSingleField(frm1.txtDrLocAmt)
    Call LockObjectField(frm1.txtDrLocAmt,"P")

    Call FormatDoubleSingleField(frm1.txtCrLocAmt)
    Call LockObjectField(frm1.txtCrLocAmt,"P")    
    
	Call InitSpreadSheet 
    Call InitCtrlSpread()
    Call InitCtrlHSpread()
    Call InitComboBox
    Call InitComboBoxGrid
    Call SetAuthorityFlag                                               '���Ѱ��� �߰�    
    Call SetToolbar(MENU_NEW)
    Call SetDefaultVal
	Call InitVariables

	Call CookiePage("FORM_LOAD")

	' ���Ѱ��� �߰� 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc)
	
	' ����� 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' ���κμ� 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' ���κμ�(��������)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' ���� 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing
		
End Sub

'=========================================================================================================
Sub InitComboBox()

	Dim iData

	'iData = CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("A1013", "''", "S") & "  and minor_cd <> " & FilterVar("04", "''", "S") & "  order by minor_cd ")	
	'Response.Write " Call SetCombo3(frm1.cboGlType,""" & iData & """)" & vbCrLf
	Call CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("A1013", "''", "S") & "  and minor_cd <> " & FilterVar("04", "''", "S") & "  order by minor_cd ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboGlType, lgF0, lgF1, Chr(11))

	'iData = CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("A1001", "''", "S") & "  order by minor_cd ")
	'Response.Write " Call SetCombo3(frm1.cboGlInputType,""" & iData & """)" & vbCrLf
	Call CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("A1001", "''", "S") & "  order by minor_cd ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboGlInputType, lgF0, lgF1, Chr(11))

End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' ���� ���� --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>					
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%> </TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǥ��ȣ</TD>
									<TD CLASS=TD656 NOWRAP><INPUT NAME="txtGlNo" ALT="��ǥ��ȣ" MAXLENGTH="18" SIZE=20 STYLE="TEXT-ALIGN: left" tag  ="12XXXU" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRefGL()"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%" ></TD>
				</TR>
				<TR>		
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP >
						<TABLE <%=LR_SPACE_TYPE_60%>>					
							<TR>
								<TD CLASS=TD5 NOWRAP>��ǥ����</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtGLDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="ȸ������" id=OBJECT7 class=required></OBJECT>');</SCRIPT></TD>								
								<TD CLASS=TD5 NOWRAP>��ǥ����</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlType" tag="23" STYLE="WIDTH:82px:" ALT="��ǥ����" class=required></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�μ�</TD>								
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="�μ��ڵ�" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="22XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)">&nbsp;
													 <INPUT NAME="txtDeptNm" ALT="�μ���"   MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="24X" class=protected readonly=true tabindex="-1"></TD>
													 <INPUT NAME="txtInternalCd" ALT="���κμ��ڵ�" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14"  TABINDEX="-1">
								<TD CLASS=TD5 NOWRAP>��ǥ�Է°��</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlInputType" tag="24" STYLE="WIDTH:82px:" ALT="��ǥ�Է°��"><OPTION VALUE="" selected></OPTION></SELECT></TD>								
			    
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>���</TD>
								<TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT NAME="txtDesc" ALT="���" MAXLENGTH="128" SIZE="70" tag="22" class=required ></TD>
							</TR>							
							<TR> 
								<TD HEIGHT="60%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>								
								</TD>
							</TR>
							<TR>

								<TD CLASS=TD656 WIDTH=* align=right COLSPAN=2><BUTTON NAME="btnCalc" CLASS="CLSSBTNCALC" ONCLICK="vbscript:FncBtnCalc()" Flag=1>�ڱ��ݾװ��</BUTTON>&nbsp;
								<TD CLASS=TD5 NOWRAP>�����հ�(�ڱ�)</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDrLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="�����հ�(�ڱ�)" id=OBJECT3 class=protected readonly=true tabindex="-1"></OBJECT>');</SCRIPT>
									&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCrLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="�뺯�հ�(�ڱ�)" id=OBJECT4 class=protected readonly=true tabindex="-1"></OBJECT>');</SCRIPT></TD>							
							</TR>
							<TR>						                 
								<TD HEIGHT="40%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT5> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>									
			  			  
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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON>&nbsp;
					</TD>					
					<TD WIDTH=* ALIGN=RIGHT>					
						<A HREF="VBSCRIPT:PgmJumpChk(JUMP_PGM_ID_TAX_REP)">��꼭����</a>		
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>

</TABLE>

<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspdData3 width="100%" tag="23" TITLE="SPREAD" id=OBJECT6 TABINDEX="-1"><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>

<TEXTAREA class=hidden name=txtSpread		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3		tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtGlNo"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtCommandMode"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtAuthorityFlag"  tag="24" TABINDEX="-1"><!--���Ѱ����߰� -->
<INPUT TYPE=HIDDEN NAME="txtGlinputType"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP"   METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname"       TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname"      TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename"    TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar"     TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date">	
</Form>
</BODY>
</HTML>
