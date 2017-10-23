<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        :
'*  3. Program ID           : a5101ma1
'*  4. Program Name         : ������ǥ ��� 
'*  5. Program Desc         : ������ǥ������ ���, ����, ����, ��ȸ 
'*  6. Component List       : PAGG005.dll
'*  7. ModIfied date(First) : 2003/01/10
'*  8. ModIfied date(Last)  : 2003/10/31
'*  9. ModIfier (First)     : Kim Ho Young
'* 10. ModIfier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. HisTory              :
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="a5101ma1_ko441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="Acctctrl_ko441_1.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate

iDBSYSDate = "<%=GetSvrDate%>"

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 


Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

Sub Form_Load()
    Call LoadInfTB19029

    Call FormatDATEField(frm1.txttempGLDt)
    Call LockObjectField(frm1.txttempGLDt,"R")

'    Call LockHTMLField(frm1.cboGlType,"R")
    Call LockHTMLField(frm1.cboConfFg,"P")
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
	Call CookiePage("ForM_LOAD")

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
	
'	msgbox lgInternalCd
'	msgbox lgSubInternalCd
'	msgbox lgAuthUsrID
'	msgbox lgAuthBizAreaCd
	
	
End Sub

'============================================================================================================
Sub InitComboBox()
	Dim iData

	'iData = CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("A1013", "''", "S") & "  and minor_cd <> " & FilterVar("04", "''", "S") & "  order by minor_cd ")	
	'Response.Write " Call SetCombo3(frm1.cboGlType,""" & iData & """)" & vbCrLf
    Call CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm) "," B_MINOR "," MAJOR_CD = " & FilterVar("A1013", "''", "S") & "   and minor_cd <> " & FilterVar("04", "''", "S") & "  order by minor_cd ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	Call SetCombo2(frm1.cboGlType, lgF0, lgF1, Chr(11))
	
	'iData = CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("A1001", "''", "S") & "  order by minor_cd ")
	'Response.Write " Call SetCombo3(frm1.cboGlInputType,""" & iData & """)" & vbCrLf
	Call CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("A1001", "''", "S")	& "  order by minor_cd ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboGlInputType, lgF0, lgF1, Chr(11))
	
	
'-- eWare Inf Begin 
	If	Trim(parent.gEware) = "" Then
	
		'iData = CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("A1007", "''", "S") & "  order by minor_cd ")	
		'Response.Write " Call SetCombo3(frm1.cboConfFg,""" & iData & """)" & vbCrLf
		Call CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("A1007", "''", "S") & "  order by minor_cd ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
		Call SetCombo2(frm1.cboConfFg, lgF0, lgF1, Chr(11))

	Else
		'iData = CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("AI001", "''", "S") & "  order by minor_cd ")	
		'Response.Write " Call SetCombo3(frm1.cboConfFg,""" & iData & """)" & vbCrLf	
		Call CommonQueryRs("rTrim(minor_cd), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("AI001", "''", "S") & "  order by minor_cd ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)		
		Call SetCombo2(frm1.cboConfFg, lgF0, lgF1, Chr(11))

	End If	

'-- eWare Inf End 	
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

</HEAD>
<!-- '#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<ForM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf"><IMG src="../../../CShared/image/table/seltab_up_left.gIf" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gIf" width="10" height="23" ></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH="100%"> </TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%"> 
					    <FIELDSET CLASS="CLSFLD">
						  <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>���ǹ�ȣ</TD>
								<TD CLASS=TD656 NOWRAP><INPUT NAME="txtTempGlNo" ALT="���ǹ�ȣ" MAXLENGTH="18" SIZE=20 STYLE="TEXT-ALIGN: left" tag="12XXXU" class=required><IMG SRC="../../../CShared/image/btnPopup.gIf" NAME="btnTempGlNo" align=Top TYPE="BUTToN" ONCLICK="vbscript:Call OpenReftempgl()"></TD>
							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%" ></TD>
				</TR>
				<TR>		
					<TD WIDTH="100%" HEIGHT=* VALIGN=ToP >
						<TABLE <%=LR_SPACE_TYPE_60%>>					
							<TR>								
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txttempGLDt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��������" tag="22" id=fpDateTime1 class=required></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>��ǥ�μ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoginDeptNm" ALT="�μ���"   MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="24X" class=protected readonly=true tabindex="-1"></TD>								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ȸ��μ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="�μ��ڵ�" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag="22XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gIf" NAME="btnCostCd" align=Top TYPE="BUTToN" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value, 0)" tag="22">&nbsp;
													 <INPUT NAME="txtDeptNm" ALT="�μ���"   MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="24X" class=protected readonly=true tabindex="-1"></TD>
													 <INPUT NAME="txtInternalCd" ALT="���κμ��ڵ�" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14"  TABINDEX="-1">
								<TD CLASS=TD5 NOWRAP>��ǥ����</TD>								
								<TD CLASS=TD6 NOWRAP><Select NAME="cboGlType" tag="23" STYLE="WIDTH:82px:" ALT="��ǥ����" class=required></Select></TD> 
						   </TR>
						   <TR>									
								<TD CLASS=TD5 NOWRAP>��ǥ�Է°��</TD>
								<TD CLASS=TD6 NOWRAP><Select NAME="cboGlInputType" tag="24" STYLE="WIDTH:82px:" ALT="��ǥ�Է°��" class=protected readonly=true tabindex="-1"><OPTION VALUE="" Selected></OPTION></Select></TD>								
<!---- eWare Inf Begin -->
<% If Trim(gEware) = "" Then %>
								<TD CLASS=TD5 NOWRAP>���λ���</TD>
								<TD CLASS=TD6 NOWRAP><Select NAME="cboConfFg" tag="24" STYLE="WIDTH:82px:" ALT="���λ���"><OPTION VALUE="" Selected></OPTION></Select></TD>
								<INPUT TYPE=HIDDEN NAME="cboAppFg" tag="24" STYLE="WIDTH:82px:" ALT="�������"TABINDEX="-1">								
<% Else %>
								<TD CLASS=TD5 NOWRAP>�������</TD>
								<TD CLASS=TD6 NOWRAP><Select NAME="cboConfFg" tag="24" STYLE="WIDTH:82px:" ALT="�������"><OPTION VALUE="" Selected></OPTION></Select></TD>
								<INPUT TYPE=HIDDEN NAME="cboConfFg" tag="24" STYLE="WIDTH:82px:" ALT="���λ���" TABINDEX="-1">								
<% End If %>
<!-- --eWare Inf End -->
							</TR>						
							<TR>
								<TD CLASS=TD5 NOWRAP>���</TD>
								<TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT NAME="txtDesc" ALT="���" MAXLENGTH="128" SIZE="70" tag="22XXXU" class=required></TD>
							</TR>	
							<TR>
								<TD HEIGHT="60%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD656 WIDTH=* align=right COLSPAN=2><BUTToN NAME="btnCalc" CLASS="CLSSBTNCALC" ONCLICK="vbscript:FncBtnCalc()" Flag=1>�ڱ��ݾװ��</BUTToN>&nbsp;
								<TD CLASS=TD5 NOWRAP>�����հ�(�ڱ�)</TD>
								<TD CLASS=TD6 NOWRAP>
									&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDrLocAmt style="HEIGHT: 20px; LEFT: 0px; ToP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="�����հ�(�ڱ�)" id=OBJECT3 class=protected readonly=true tabindex="-1"></OBJECT>');</SCRIPT>
									&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCrLocAmt style="HEIGHT: 20px; LEFT: 0px; ToP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="�뺯�հ�(�ڱ�)" tag="24X2" id=OBJECT4 class=protected readonly=true tabindex="-1"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
			                <TR>
								<TD HEIGHT="40%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
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
					<TD><BUTToN NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTToN>&nbsp;
						<BUTToN NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTToN>&nbsp;
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IfRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IfRAME></TD>
	</TR>
</TABLE>
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TABINDEX="-1" CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=0 HEIGHT=0 tag="23" TITLE="SPREAD" id=vaSpread3><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
<TEXTAREA class=hidden name=txtSpread		tag="24" Tabindex="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3		tag="24" Tabindex="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtTempGlNo"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtCommAndMode"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtAuthorityFlag"  tag="24" Tabindex="-1"><!--���Ѱ����߰� -->
<INPUT TYPE=HIDDEN NAME="hCongFg"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hItemSeq"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hAcctCd"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthpjt_no"		tag="24" Tabindex="-1">
</Form>
<DIV ID="MousePT" NAME="MousePT">
<Iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></Iframe>
</DIV>
<ForM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</Form>
</BODY>
</HTML>
