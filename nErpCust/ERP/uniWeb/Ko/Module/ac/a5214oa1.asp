<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Closing and Financial Statements
'*  3. Program ID           : A5214OA1
'*  4. Program Name         : �����ⳳ�� ��� 
'*  5. Program Desc         : Report of Cash Flow
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/06/01
'*  8. Modified date(Last)  : 2004/01/12
'*  9. Modifier (First)     : Cho, Ig Sung
'* 10. Modifier (Last)      : Kim Chang Jin
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance
'========================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	              ' Variable is for Operation Status 

Dim IsOpenPop

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'========================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0 
End Sub

'========================================================================================
Sub SetDefaultVal()

	Dim FiscYear, FiscMonth, FiscDay, strYear, strMonth, strDay, FiscDate, EndDate, StartDate

	EndDate = "<%=GetSvrDate%>"

	Call ExtractDateFrom(Parent.gFiscStart, Parent.gServerDateFormat, Parent.gServerDateType, FiscYear, FiscMonth, FiscDay)
	Call ExtractDateFrom(EndDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

	if strMonth < FiscMonth then
		FiscYear	= cstr(cint(strYear) - 1)
	else
		FiscYear	= strYear
	end if

	FiscDate	= UniConvYYYYMMDDToDate(Parent.gDateFormat, FiscYear, FiscMonth, FiscDay)
	StartDate	= UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
	EndDate		= UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)


	frm1.txtDateFr.Text			= StartDate
	frm1.txtDateTo.Text			= EndDate 
	frm1.hFiscStartDt.value		= FiscDate 	'company format

End Sub

'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "OA") %>
<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","OA") %>
End Sub


'========================================================================================
Sub txtDateFr_Change()
	dim startyyyy, startmm, startdd, fiscyyyy, fiscmm, fiscdd
	with frm1
		Call ExtractDateFrom(.txtDateFr.text,.txtDateFr.UserDefinedFormat,Parent.gComDateType,startyyyy,startmm,startdd)
		Call ExtractDateFrom(.hFiscStartDt.value,.txtDateFr.UserDefinedFormat,Parent.gComDateType,fiscyyyy,fiscmm,fiscdd)
		If startyyyy = "" Or fiscyyyy = "" Then
			Exit Sub
		End If

		'��ȸ ���ۿ��� �����ۿ����� �����̸� �����۳⵵�� ��ȸ ���۳⵵���� 1�� ���ش�.
		if startmm < fiscmm then
			fiscyyyy	= cstr(cint(startyyyy) - 1)
		else
			fiscyyyy	= startyyyy
		end if
		.hFiscStartDt.value	= UniConvYYYYMMDDToDate(Parent.gDateFormat,fiscyyyy,fiscmm,fiscdd)
	end with
End Sub

'========================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 0, 2
			arrParam(0) = "������ڵ� �˾�"								' �˾� ��Ī 
			arrParam(1) = "B_BIZ_AREA" 										' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			
			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "������ڵ�"									' �����ʵ��� �� ��Ī 

			arrField(0) = "BIZ_AREA_CD"										' Field��(0)
			arrField(1) = "BIZ_AREA_NM"										' Field��(1)
    
			arrHeader(0) = "������ڵ�"									' Header��(0)
			arrHeader(1) = "������"									' Header��(1)

		Case 1
			arrParam(0) = "�ŷ���ȭ�˾�"					' �˾� ��Ī 
			arrParam(1) = "b_currency"							' TABLE ��Ī 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = ""									' Where Condition
			arrParam(5) = "�ŷ���ȭ"

			arrField(0) = "CURRENCY"							' Field��(0)
			arrField(1) = "CURRENCY_DESC"						' Field��(1)

			arrHeader(0) = "�ŷ���ȭ"						' Header��(0)
			arrHeader(1) = "�ŷ���ȭ��"						' Header��(1)

		Case Else
			Exit Function
	End Select

	IsOpenPop = True

    Select Case iWhere

	Case 0, 1, 2
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0	'������ڵ� 
				frm1.txtBizAreaCd.focus
			Case 1
				frm1.txtDocCur.focus
			Case 2	'������ڵ� 
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
		Case 0	'������ڵ� 
			frm1.txtBizAreaCd.focus
			frm1.txtBizAreaCd.value = arrRet(0)
			frm1.txtBizAreaNm.value = arrRet(1)

		Case 1
			frm1.txtDocCur.focus
			frm1.txtDocCur.value = arrRet(0)
			frm1.txtDocCurNm.value = arrRet(1)
			
		Case 2	'������ڵ� 
			frm1.txtBizAreaCd1.focus
			frm1.txtBizAreaCd1.value = arrRet(0)
			frm1.txtBizAreaNm1.value = arrRet(1)


		Case Else
			Exit Function
	End select	

End Function
'========================================================================================
Sub Form_Load()
    Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")

    Call InitVariables 
    Call SetDefaultVal

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
	frm1.PrintOpt1.checked = True

	DocCur.innerHTML = ""
	frm1.txtDocCur.value	= Parent.gCurrency
	Call ElementVisible(frm1.txtDocCur, 0)
	Call ElementVisible(frm1.txtDocCurNm, 0)
	Call ElementVisible(frm1.btnDocCur, 0)
End Sub

'========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
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
Function PrintOpt1_OnClick() 
	if frm1.PrintOpt1.checked = True then
		DocCur.innerHTML = ""
		Call ElementVisible(frm1.txtDocCur, 0)
		Call ElementVisible(frm1.txtDocCurNm, 0)
		Call ElementVisible(frm1.btnDocCur, 0)

		frm1.txtDocCur.value	= Parent.gCurrency
	end if
End Function

'========================================================================================
Function PrintOpt2_OnClick() 
	if frm1.PrintOpt2.checked = True then
		DocCur.innerHTML = "�ŷ���ȭ"
		Call ElementVisible(frm1.txtDocCur, 1)
		Call ElementVisible(frm1.txtDocCurNm, 1)
		Call ElementVisible(frm1.btnDocCur, 1)

		frm1.txtDocCur.value	= ""
		frm1.txtDocCurNm.value	= ""
	end if
End Function


'========================================================================================
Function SetPrintCond(StrEbrFile, VarFiscDt, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarDocCur)
	Dim IntRetCD
	
	if frm1.PrintOpt1.checked = True then
		StrEbrFile = "a5214ma1"
		VarDocCur	= Parent.gCurrency
	else
		StrEbrFile = "a5214oa2"
		VarDocCur	= UCase(frm1.txtDocCur.value)
	end if
	
	VarFiscDt	= UniConvDateToYYYYMMDD(frm1.hFiscStartDt.value,Parent.gDateFormat,"")
	VarDateFr	= UniConvDateToYYYYMMDD(frm1.txtDateFr.Text,Parent.gDateFormat,"")
	VarDateTo	= UniConvDateToYYYYMMDD(frm1.txtDateTo.Text,Parent.gDateFormat,"")

	'���Ѱ���	
	If frm1.txtBizAreaCd.value = "" then 
		frm1.txtBizAreaNm.value = ""
		If lgAuthBizAreaCd <> "" Then			
			VarBizAreaCd  = lgAuthBizAreaCd
		Else
			VarBizAreaCd = "0"
		End If			
	Else 
		If lgAuthBizAreaCd <> "" Then
			VarBizAreaCd = Trim(FilterVar(frm1.txtBizAreaCD.value,"","SNM"))
			If UCASE(lgAuthBizAreaCd) <> UCASE(VarBizAreaCd) Then
				IntRetCD = DisplayMsgBox("124200","x","x","x")
				frm1.txtBizAreaCD.focus()
				SetPrintCond =  False
				Exit Function
			End If
		Else
			VarBizAreaCd = FilterVar(frm1.txtBizAreaCD.value,"","SNM")
		End If
	End if

	If frm1.txtBizAreaCd1.value = "" then
		frm1.txtBizAreaNm1.value = ""
		If lgAuthBizAreaCd <> "" Then			
			VarBizAreaCd1 = lgAuthBizAreaCd
		Else
			VarBizAreaCd1 = "ZZZZZZZZZZ"
		End If			
	Else 
		If lgAuthBizAreaCd <> "" Then
			VarBizAreaCd1 = Trim(FilterVar(frm1.txtBizAreaCD1.value,"","SNM"))
			If UCASE(lgAuthBizAreaCd) <> UCASE(VarBizAreaCd1) Then
				IntRetCD = DisplayMsgBox("124200","x","x","x")
				frm1.txtBizAreaCD1.focus()
				SetPrintCond =  False
				Exit Function
			End If
		Else
			VarBizAreaCd1 = FilterVar(frm1.txtBizAreaCD1.value,"","SNM")
		End If
	End if

	SetPrintCond =  True
	
End Function

'========================================================================================
Function FncBtnPrint()
    Dim StrEbrFile, VarFiscDt, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarDocCur
    Dim StrUrl
    Dim IntRetCD

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	If UniConvDateToYYYYMMDD(frm1.txtDateFr.Text, Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, Parent.gDateFormat,"") Then
		Call DisplayMsgBox("970025", "X", frm1.txtDateFr.Alt, frm1.txtDateTo.Alt)
		frm1.txtDateFr.focus
		Exit Function
	End If

	IntRetCD = SetPrintCond(StrEbrFile, VarFiscDt, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarDocCur)
	If IntRetCD = False Then
	    Exit Function
 	End If

	ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	StrUrl = StrUrl & "DocCur|" & VarDocCur
	StrUrl = StrUrl & "|DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|" & VarBizAreaCd1

	if frm1.PrintOpt1.checked = True then
		StrUrl = StrUrl & "|FiscDt|" & VarFiscDt
	end if

	Call FncEBRPrint(EBAction,ObjName,StrUrl)	

End Function


'========================================================================================
Function FncBtnPreview() 

    Dim StrEbrFile, VarFiscDt, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarDocCur
    Dim StrUrl
    Dim IntRetCD

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	If UniConvDateToYYYYMMDD(frm1.txtDateFr.Text, Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, Parent.gDateFormat,"") Then
		Call DisplayMsgBox("970025", "X", frm1.txtDateFr.Alt, frm1.txtDateTo.Alt)
		frm1.txtDateFr.focus
		Exit Function
	End If

	IntRetCD = SetPrintCond(StrEbrFile, VarFiscDt, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarDocCur)
	If IntRetCD = False Then
	    Exit Function
 	End If

	ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	StrUrl = StrUrl & "DocCur|" & VarDocCur
	StrUrl = StrUrl & "|DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|" & VarBizAreaCd1

	if frm1.PrintOpt1.checked = True then
		StrUrl = StrUrl & "|FiscDt|" & VarFiscDt
	end if

	Call FncEBRPreview(ObjName,StrUrl)
		
End Function

'========================================================================================
Function FncPrint()
    Call Parent.FncPrint()
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
Sub txtDateFr_DblClick(Button)
    If Button = 1 Then
        frm1.txtDateFr.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtDateFr.Focus
    End If
End Sub

'========================================================================================
Sub txtDateTo_DblClick(Button)
    If Button = 1 Then
        frm1.txtDateTo.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtDateTo.Focus
    End If
End Sub


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

</HEAD>

<!--
'#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' ���� ���� --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100% colspan="2">
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
	<TR>
		<TD WIDTH=100% CLASS="Tab11" HEIGHT=* colspan="2">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>��±���</TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" CHECKED ID="PrintOpt1" VALUE="Y" tag="25"><LABEL FOR="PrintOpt1">�ڱ���ȭ</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt2" VALUE="N" tag="25"><LABEL FOR="PrintOpt2">�ŷ���ȭ</LABEL></SPAN></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,0)"> <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="������"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,2)"> <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="������"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>ȸ������</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateFr" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=����ȸ������ id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
													   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateTo" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=����ȸ������ id=fpDateTime2></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" ID="DocCur" NOWRAP>�ŷ���ȭ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDocCur" SIZE=10 MAXLENGTH=3 tag="12XXXU" ALT="�ŷ���ȭ" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.Value,1)"> <INPUT TYPE="Text" NAME="txtDocCurNm" SIZE=25 tag="14X" ALT="�ŷ���ȭ��"></TD>
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
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hFiscStartDt" tag="24">
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

