<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : ���� 
'*  2. Function Name        : �ǸŰ�ȹ���� 
'*  3. Program ID           : S2211BA3.asp
'*  4. Program Name         : �ǸŰ�ȹ�Ⱓ�������� 
'*  5. Program Desc         : �ǸŰ�ȹ�Ⱓ�������� 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2002/12/13
'*  8. Modified date(Last)  : 2003/02/13
'*  9. Modifier (First)     : Park Yong Sik
'* 10. Modifier (Last)      : Hwang Seong Bae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<% '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################%>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/lgvariables.inc"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             <% '��: indicates that All variables must be declared in advance %>

<%'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************%>
<%'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================%>

Const BIZ_PGM_ID = "s2211bb3.asp"											<% '��: �����Ͻ� ���� ASP�� %>
Const BIZ_PGM_JUMP_ID = "s2211ma3"

<% '==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= %>

<% '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= %>
Dim IsOpenPop

<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                       	              '��: Indicates that current mode is Create mode
    
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False							'��: ����� ���� �ʱ�ȭ 
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.cboSpType.value = "E"
	Call cboSpType_onChange()
End Sub

'======================================================================================================== 
Sub InitComboBox()	
	' �ǸŰ�ȹ���� 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S0023", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboSpType,lgF0,lgF1,Chr(11))

	'�����⵵ 
	Call CommonQueryRs(" C_YEAR", " S_YEAR ", " USAGE = " & FilterVar("Y", "''", "S") & "  ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboCYear, lgF0,lgF0, Chr(11))

End Sub

'=======================================================================================================
'	Description : �Ⱓ������� Fetch
'========================================================================================================= 
Sub GetMethodofCreatePeriod()
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs
	
	On Error Resume Next
	
	Err.Clear
	
	With frm1
		iStrSelectList	= " MN.MINOR_CD, MN.MINOR_NM "
		iStrFromList	= " dbo.B_MINOR MN INNER JOIN dbo.B_CONFIGURATION CF ON (CF.MAJOR_CD = MN.MAJOR_CD AND CF.MINOR_CD = MN.MINOR_CD) "
		iStrWhereList	= " CF.MAJOR_CD = " & FilterVar("S0018", "''", "S") & " " & _
						  " AND CF.SEQ_NO = (SELECT CAST(REFERENCE AS SMALLINT) " & _
						  " FROM B_CONFIGURATION " & _
						  " WHERE MAJOR_CD = " & FilterVar("S0023", "''", "S") & " " & _
						  " AND	SEQ_NO = 1 " & _	
						  " AND	MINOR_CD =  " & FilterVar(.cboSptype.value , "''", "S") & ")"
	
		If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
			iArrRs = Split(iStrRs, parent.gColSep)
			.txtPeriodMethodCd.value = Trim(iArrRs(1))
			.txtPeriodMethodNm.value = Trim(iArrRs(2))
		Else
			If Err.number = 0 Then
				.txtPeriodMethodCd.value = ""
				.txtPeriodMethodNm.value = ""
			Else
				MsgBox Err.description, vbInformation,Parent.gLogoName
				Err.Clear
				Exit Sub
			End If
		End If
	End With
End Sub

'=======================================================================================================
'	Description : ���������⵵�� Fetch
'========================================================================================================= 
Sub GetLastCrYear()
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs
	
	On Error Resume Next
	
	Err.Clear

	With frm1
		iStrSelectList	= " ISNULL(YEAR(MAX(FROM_DT)), 0) "
		iStrFromList	= " dbo.S_SP_PERIOD_HISTORY "
		iStrWhereList	= " SP_TYPE =  " & FilterVar(.cboSptype.value , "''", "S") & ""
	
		If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
			iArrRs = Split(iStrRs, parent.gColSep)
			If CInt(Trim(iArrRs(1))) > 0 Then
				.txtLastCrYear.value = Trim(iArrRs(1))
			Else
				.txtLastCrYear.value = ""
			End If
		Else
			If Err.number = 0 Then
				.txtLastCrYear.value = ""
			Else
				MsgBox Err.description, vbInformation,Parent.gLogoName
				Err.Clear
				Exit Sub
			End If
		End If
		
		'�����⵵ Default �� ó�� 
		If .txtLastCrYear.value <> "" Then
			.cboCYear.value = Cstr(Cint(.txtLastCrYear.value) + 1 )
		Else
			.cboCYear.value = Mid(Trim("<%=GetSvrDate%>"),1,4)
		End If
	End With
End Sub

<%
'=======================================================================================================
' Function Desc : �Ⱓ�������� jump
'=======================================================================================================
%>
Function LoadSPPeriod()
	On Error Resume Next
	Dim iArrSpPeriod, iArrSpPeriodDesc
	
	Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>

	With frm1
		If .cboSpType.value <> "" Then
			If .cboCYear.value <> "" then
				If CommonQueryRs("TOP 1 SP_PERIOD, SP_PERIOD_DESC ", "S_SP_PERIOD_INFO", "SP_YEAR= " & Trim(.cboCYear.value) & " AND SP_TYPE =  " & FilterVar(.cboSpType.value , "''", "S") & " ORDER BY SP_PERIOD_SEQ ASC", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
					iArrSpPeriod = Split(lgF0, parent.gColSep)
					iArrSpPeriodDesc = Split(lgF1, parent.gColSep)
					
					WriteCookie CookieSplit , .cboSpType.value & Parent.gColSep & Trim(iArrSpPeriod(0)) & Parent.gColSep & Trim(iArrSpPeriodDesc(0))
				Else
					WriteCookie CookieSplit , .cboSpType.value & Parent.gColSep & "" & Parent.gColSep & ""
				End If
			Else
				WriteCookie CookieSplit , .cboSpType.value & Parent.gColSep & "" & Parent.gColSep & ""
			end if
		End If
	End With

	Call PgmJump(BIZ_PGM_JUMP_ID)
End Function

'========================================================================================================= 
Sub Form_Load()
																		'��: Load Common DLL
    Call InitVariables													'��: Initializes local global variables    
    Call ggoOper.LockField(Document, "N")								'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------           
    Call InitComboBox
    Call SetDefaultVal
    Call SetToolbar("1000000000000111")
    
	frm1.cboSpType.focus

End Sub

'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

  
'========================================================================================================= 
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         <%'��:ȭ�� ����, Tab ���� %>
End Function

'========================================================================================================= 
Function FncExit()
    FncExit = True
End Function

'=======================================================================================================
Function ExeReflect() 
	Call BtnDisabled(1)
	Dim strVal
	Dim IntRetCD

	ExeReflect = False                                                          '��: Processing is NG
    
	On Error Resume Next                                                   '��: Protect system from crashing

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0)
		Exit Function
	End If
	
	With frm1

		if .txtPeriodMethodCd.value =  "" then
			' �Ⱓ��������� �����Ǿ� ���� �ʽ��ϴ�.(Major Code : S0018).
			Call DisplayMsgBox("202420", "X", "X", "X")
			Call BtnDisabled(0)
			Exit Function
		end if

		If DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X") = vbNo Then
			Call BtnDisabled(0)
			Exit Function
		End If

		.txtFromDt.value = UniConvYYYYMMDDToDate(gDateFormat,.cboCYear.value,"01","01")
		.txtToDt.value = UniConvYYYYMMDDToDate(gDateFormat,.cboCYear.value,"12","31")
	
		strVal = BIZ_PGM_ID & "?txtSpType="	& .cboSpType.value				'��ȹ���� 
		strVal = strVal  & "&txtFromDt=" & .txtFromDt.value					'�����⵵�� 12�� 31�� 
		strVal = strVal  & "&txtToDt=" & .txtToDt.value						'�����⵵�� 12�� 31�� 
		strVal = strVal  & "&txtMethod=" & .txtPeriodMethodCd.value			'�Ⱓ������� 
		strVal = strVal  & "&txtUserId=" & parent.gUsrId					'User Id

	End With

	If LayerShowHide(1) = False then
		Call BtnDisabled(0)
		Exit Function 
	End if

	Call RunMyBizASP(MyBizASP, strVal)	                                        '��: �����Ͻ� ASP �� ���� 

	ExeReflect = True                                                           '��: Processing is NG
End Function

'=======================================================================================================
Function ExeReflectOk()				            '��: ���� ������ ���� ���� 
	Call DisplayMsgBox("990000","X","X","X")
End Function

'========================================================================================================
'                        Tag Event
'========================================================================================================
Sub cboSpType_onChange()
	Call GetMethodofCreatePeriod
	Call GetLastCrYear
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB4" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ǸŰ�ȹ�Ⱓ��������</font></td>
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
					<TD CLASS="TD5" NOWRAP>�ǸŰ�ȹ����</TD>
					<TD CLASS="TD6"><SELECT Name="cboSpType" ALT="�ǸŰ�ȹ����" tag="13XXXU"></SELECT></TD>
				</TR>
				<TR>		
					<TD CLASS="TD5">���������⵵</TD>
					<TD CLASS="TD6"><INPUT NAME="txtLastCrYear" ALT="���������⵵" TYPE="Text" MAXLENGTH="4" SIZE=13 tag="14XXXU">
					</TD>									
				</TR>
				<TR>
					<TD CLASS="TD5">�����⵵</TD>
					<TD CLASS="TD6"><SELECT NAME="cboCYear" ALT="�����⵵" tag="13XXXU"></SELECT></TD>									
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>�Ⱓ�������</TD>
					<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPeriodMethodNm" ALT="�Ⱓ�������" TYPE="Text" MAXLENGTH="10" SIZE=13 tag="14XXXU">
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
				<TABLE WIDTH=100%>
				      <TR>
				         <TD WIDTH=10>&nbsp;</TD>
				         <TD><BUTTON NAME="btnRun" ONCLICK="vbscript:ExeReflect()" CLASS="CLSMBTN" flag=1>�Ⱓ����</BUTTON>
				         </TD>
						 <TD WIDTH=* ALIGN=RIGHT><A HREF="vbscript:LoadSPPeriod">�Ⱓ��������</A></TD>
						 <TD WIDTH=10>&nbsp;</TD> 
				      </TR>
				</TABLE>
			</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPeriodMethodCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpPeriod" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpPeriodDesc" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
