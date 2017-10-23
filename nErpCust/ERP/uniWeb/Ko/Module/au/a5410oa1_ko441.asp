<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Closing and Financial Statements
'*  3. Program ID           : a5410oa1_ko441
'*  4. Program Name         : ������ü����Ʈ ��� 
'*  5. Program Desc         : Report of G/L
'*  6. Component List       : 
'*  7. Modified date(First) : 2008/07/17
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit																	'��: indicates that All variables must be declared in advance

'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	              ' Variable is for Operation Status 


' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop

'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
End Sub


'========================================================================================

Sub SetDefaultVal()
	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate

	EndDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(EndDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

	StartDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
	EndDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)
	frm1.txtBaseDt.Text = StartDate 
End Sub


'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "A", "NOCOOKIE", "PA") %>
End Sub

'========================================================================================
Sub InitSpreadSheet()
End Sub

'========================================================================================
Sub SetSpreadLock()
End Sub

'========================================================================================
Sub SetSpreadColor(ByVal lRow)
End Sub

'========================================================================================
Sub InitComboBox()

End Sub

'========================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Select Case iWhere
		Case 0, 3
			arrParam(0) = "������ڵ� �˾�"								' �˾� ��Ī %>
			arrParam(1) = "B_BIZ_AREA" 										' TABLE ��Ī %>
			arrParam(2) = strCode											' Code Condition%>
			arrParam(3) = ""												' Name Cindition%>
			
			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "������ڵ�"									' �����ʵ��� �� ��Ī %>

			arrField(0) = "BIZ_AREA_CD"										' Field��(0)%>
			arrField(1) = "BIZ_AREA_NM"										' Field��(1)%>

			arrHeader(0) = "������ڵ�"									' Header��(0)%>
			arrHeader(1) = "������"									' Header��(1)%>

		Case 1, 2
			arrParam(0) = "���� �˾�"									' �˾� ��Ī 
			arrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE ��Ī 
			arrParam(2) = Trim(strCode)									' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD"					' Where Condition
			arrParam(5) = "�����ڵ�"									' �����ʵ��� �� ��Ī 

			arrField(0) = "A_ACCT.Acct_CD"									' Field��(0)
			arrField(1) = "A_ACCT.Acct_NM"									' Field��(1)
    		arrField(2) = "A_ACCT_GP.GP_CD"									' Field��(2)
			arrField(3) = "A_ACCT_GP.GP_NM"									' Field��(3)
			
			arrHeader(0) = "�����ڵ�"									' Header��(0)
			arrHeader(1) = "�����ڵ��"									' Header��(1)
			arrHeader(2) = "�׷��ڵ�"									' Header��(2)
			arrHeader(3) = "�׷��"										' Header��(3)

		Case Else
			Exit Function
	End Select

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0	'������ڵ� 
				frm1.txtBizAreaCd.focus
			Case 1	'���۰����ڵ� 
				frm1.txtAcctCdFr.focus
			Case 2	'��������ڵ� 
				frm1.txtAcctCdTo.focus
			Case 3	'������ڵ� 
				frm1.txtBizAreaCd1.focus
		End select	
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If

End Function


'========================================================================================
Function SetReturnPopUp(ByVal arrRet, ByVal iWhere)
	
	Select Case iWhere
		Case 0	'������ڵ� %>
			frm1.txtBizAreaCd.focus
			frm1.txtBizAreaCd.value = arrRet(0)
			frm1.txtBizAreaNm.value = arrRet(1)
		Case 1	'���۰����ڵ� %>
			frm1.txtAcctCdFr.focus
			frm1.txtAcctCdFr.value = arrRet(0)
			frm1.txtAcctNmFr.value = arrRet(1)
			frm1.txtAcctCdTo.value = arrRet(0)
			frm1.txtAcctNmTo.value = arrRet(1)
		Case 2	'��������ڵ� %>
			frm1.txtAcctCdTo.focus
			frm1.txtAcctCdTo.value = arrRet(0)
			frm1.txtAcctNmTo.value = arrRet(1)
		Case 3	'������ڵ� %>
			frm1.txtBizAreaCd1.focus
			frm1.txtBizAreaCd1.value = arrRet(0)
			frm1.txtBizAreaNm1.value = arrRet(1)
		Case Else
	End select	

End Function


'========================================================================================
Sub Form_Load()

    Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call SetDefaultVal
	Call InitComboBox
    Call SetToolBar("1000000000001111")

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

    
	frm1.txtBizAreaCd.focus 
End Sub

'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'========================================================================================
Sub txtBaseDt_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime1.Action = 7
        Call SetFocusToDocument("M")
        frm1.fpDateTime1.Focus
    End If
End Sub

'========================================================================================================
'   Event Name : txtBizAreaCd_Onchange()
'   Event Desc : ������ڵ带 �����Է��Ұ�쿡 ������ڵ���� �������ش�.
'========================================================================================================
sub txtBizAreaCd_Onchange()
	Dim strCd
	Dim strWhere
	Dim IntRetCD

	strCd = Trim(frm1.txtBizAreaCd.value)
	If strCd = "" Then
		frm1.txtBizAreaNm.value = ""
	Else
		If lgAuthBizAreaCd <> "" AND UCASE(lgAuthBizAreaCd) <> UCASE(strCd) Then
			frm1.txtBizAreaNm.value = ""
			IntRetCD = DisplayMsgBox("124200","x","x","x")
		Else
			strWhere = "BIZ_AREA_CD = " & FilterVar(strCd, "''", "S")
			
			Call CommonQueryRs("BIZ_AREA_NM","B_BIZ_AREA",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			if Trim(Replace(lgF0,Chr(11),"")) = "X" then
				frm1.txtBizAreaNm.value = ""
			else
				frm1.txtBizAreaNm.value = Trim(Replace(lgF0,Chr(11),""))
			end if
		End If
	End If
	
End sub

'========================================================================================================
'   Event Name : txtBizAreaCd1_Onchange()
'   Event Desc : ������ڵ带 �����Է��Ұ�쿡 ������ڵ���� �������ش�.
'========================================================================================================
sub txtBizAreaCd1_Onchange()
	Dim strCd
	Dim strWhere
	Dim IntRetCD

	strCd = Trim(frm1.txtBizAreaCd1.value)
	If strCd = "" Then
		frm1.txtBizAreaNm1.value = ""
	Else
		If lgAuthBizAreaCd <> "" AND UCASE(lgAuthBizAreaCd) <> UCASE(strCd) Then
			frm1.txtBizAreaNm1.value = ""
			IntRetCD = DisplayMsgBox("124200","x","x","x")
		Else
			strWhere = "BIZ_AREA_CD = " & FilterVar(strCd, "''", "S")
			
			Call CommonQueryRs("BIZ_AREA_NM","B_BIZ_AREA",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			if Trim(Replace(lgF0,Chr(11),"")) = "X" then
				frm1.txtBizAreaNm1.value = ""
			else
				frm1.txtBizAreaNm1.value = Trim(Replace(lgF0,Chr(11),""))
			end if
		End If
	End If
 
End sub


'========================================================================================
Function FncBtnPrint() 
    Dim StrUrl, StrEbrFile
    Dim VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt
	Dim IntRetCD

    If Not chkField(Document, "1") Then									'��: This function check indispensable field%>
       Exit Function
    End If

	StrEbrFile = "a5410oa1_ko441"


	ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	StrUrl = StrUrl & "Base_dt|" & frm1.txtBaseDt.Text
	StrUrl = StrUrl & "|BizAreaCd|" & "%"

	Call FncEBRPrint(EBAction,ObjName,StrUrl)

End Function


'========================================================================================
Function FncBtnPreview() 
    Dim StrUrl, StrEbrFile
    Dim VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt
	Dim IntRetCD
    
    If Not chkField(Document, "1") Then									'��: This function check indispensable field%>
       Exit Function
    End If

	StrEbrFile = "a5410oa1_ko441"

	ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	StrUrl = StrUrl & "Base_dt|" & frm1.txtBaseDt.Text
	StrUrl = StrUrl & "|BizAreaCd|" & "%"

	Call FncEBRPreview(ObjName,StrUrl)

End Function

'========================================================================================
Function GetFiscDate( ByVal strFromDate)

	Dim strFiscYYYY, strFiscMM, strFiscDD
	Dim strFromYYYY, strFromMM, strFromDD
	
	GetFiscDate				="19000101"	
	
	Call ExtractDateFrom(Parent.gFiscStart,	Parent.gServerDateFormat,	Parent.gServerDateType,	strFiscYYYY,	strFiscMM,	strFiscDD)
	Call ExtractDateFrom(strFromDate,	Parent.gDateFormat,		Parent.gComDateType,		strFromYYYY,	strFromMM,	strFromDD)
	
	strFiscYYYY =  strFromYYYY
	
	If isnumeric(strFromYYYY) And isnumeric(strFromMM) And isnumeric(strFiscMM) Then
	
		If Cint(strFiscMM) > Cint(strFromMM)  then
		   GetFiscDate	= Cstr(Cint(strFromYYYY) - 1) & strFiscMM & strFiscDD
		Else
		   GetFiscDate	= strFromYYYY & strFiscMM & strFiscDD
		End If
	
	End If

End Function

'========================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'========================================================================================
Function FncExcel() 
End Function

'========================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function


'========================================================================================
Function FncExit()
    FncExit = True
End Function


'========================================================================================
Function DbQuery()
End Function

'========================================================================================
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->

</HEAD>

<!--
'#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>������(������)</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtBaseDt" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=������ id=fpDateTime1></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP></TD>
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
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1">
</FORM>
</BODY>
</HTML>

