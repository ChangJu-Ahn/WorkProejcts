<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Closing and Financial Statements
'*  3. Program ID           : A5215OA1
'*  4. Program Name         : 총계정원장 출력 
'*  5. Program Desc         : Report of G/L
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/12/18
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
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

Option Explicit																	'☜: indicates that All variables must be declared in advance

'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	              ' Variable is for Operation Status 


' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
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
	frm1.txtDateFr.Text = StartDate 
	frm1.txtDateTo.Text = EndDate 
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
			arrParam(0) = "사업장코드 팝업"								' 팝업 명칭 %>
			arrParam(1) = "B_BIZ_AREA" 										' TABLE 명칭 %>
			arrParam(2) = strCode											' Code Condition%>
			arrParam(3) = ""												' Name Cindition%>
			
			' 권한관리 추가 
			If lgAuthBizAreaCd <>  "" Then
				arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "사업장코드"									' 조건필드의 라벨 명칭 %>

			arrField(0) = "BIZ_AREA_CD"										' Field명(0)%>
			arrField(1) = "BIZ_AREA_NM"										' Field명(1)%>

			arrHeader(0) = "사업장코드"									' Header명(0)%>
			arrHeader(1) = "사업장명"									' Header명(1)%>

		Case 1, 2
			arrParam(0) = "계정 팝업"									' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE 명칭 
			arrParam(2) = Trim(strCode)									' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD"					' Where Condition
			arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "A_ACCT.Acct_CD"									' Field명(0)
			arrField(1) = "A_ACCT.Acct_NM"									' Field명(1)
    		arrField(2) = "A_ACCT_GP.GP_CD"									' Field명(2)
			arrField(3) = "A_ACCT_GP.GP_NM"									' Field명(3)
			
			arrHeader(0) = "계정코드"									' Header명(0)
			arrHeader(1) = "계정코드명"									' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"										' Header명(3)

		Case Else
			Exit Function
	End Select

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0	'사업장코드 
				frm1.txtBizAreaCd.focus
			Case 1	'시작계정코드 
				frm1.txtAcctCdFr.focus
			Case 2	'종료계정코드 
				frm1.txtAcctCdTo.focus
			Case 3	'사업장코드 
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
		Case 0	'사업장코드 %>
			frm1.txtBizAreaCd.focus
			frm1.txtBizAreaCd.value = arrRet(0)
			frm1.txtBizAreaNm.value = arrRet(1)
		Case 1	'시작계정코드 %>
			frm1.txtAcctCdFr.focus
			frm1.txtAcctCdFr.value = arrRet(0)
			frm1.txtAcctNmFr.value = arrRet(1)
			frm1.txtAcctCdTo.value = arrRet(0)
			frm1.txtAcctNmTo.value = arrRet(1)
		Case 2	'종료계정코드 %>
			frm1.txtAcctCdTo.focus
			frm1.txtAcctCdTo.value = arrRet(0)
			frm1.txtAcctNmTo.value = arrRet(1)
		Case 3	'사업장코드 %>
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

	' 권한관리 추가 
	Dim xmlDoc

	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc)

	' 사업장		
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서		
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text

	' 내부부서(하위포함)		
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text

	' 개인						
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text

	Set xmlDoc = Nothing

    
	frm1.txtBizAreaCd.focus 
End Sub

'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'========================================================================================
Sub txtDateFr_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime1.Action = 7
        Call SetFocusToDocument("M")
        frm1.fpDateTime1.Focus
    End If
End Sub

'========================================================================================
Sub txtDateTo_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime2.Action = 7
        Call SetFocusToDocument("M")
        frm1.fpDateTime2.Focus
    End If
End Sub

'========================================================================================================
'   Event Name : txtBizAreaCd_Onchange()
'   Event Desc : 사업장코드를 직접입력할경우에 사업장코드명을 설정해준다.
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
'   Event Desc : 사업장코드를 직접입력할경우에 사업장코드명을 설정해준다.
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
Function SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt)
	Dim IntRetCD
	
	StrEbrFile = "a5215ma1"
	
	VarDateFr	= UniConvDateToYYYYMMDD(frm1.txtDateFr.Text,Parent.gDateFormat,"")
	VarDateTo	= UniConvDateToYYYYMMDD(frm1.txtDateTo.Text,Parent.gDateFormat,"")
	VarFiscDt	= GetFiscDate(frm1.txtDateFr.Text)

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
'msgbox 	VarBizAreaCd & "**" & VarBizAreaCd1

	VarAcctCdFr = "0"
	VarAcctCdTo = "ZZZZZZZZZZZZZZZZZZZZ"

	If Len(frm1.txtAcctCdFr.value) > 0 Then VarAcctCdFr = frm1.txtAcctCdFr.value
	If Len(frm1.txtAcctCdTo.value) > 0 Then VarAcctCdTo = frm1.txtAcctCdTo.value

	SetPrintCond =  True
	
End Function

'========================================================================================
Function FncBtnPrint() 
    Dim StrUrl, StrEbrFile
    Dim VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt
	Dim IntRetCD

    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field%>
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtDateFr.Text,frm1.txtDateTo.Text,frm1.txtDateFr.Alt,frm1.txtDateTo.Alt, _
	 "970025", frm1.txtDateFr.UserDefinedFormat,Parent.gComDateType, true)=False then
		frm1.txtDateTo.Focus
		Exit Function
	End If

	frm1.txtAcctCdFr.value = Trim(frm1.txtAcctCdFr.value)
	frm1.txtAcctCdTo.value = Trim(frm1.txtAcctCdTo.value)

	If frm1.txtAcctCdFr.value > frm1.txtAcctCdTo.value Then
		Call DisplayMsgBox("970025", "X", frm1.txtAcctCdFr.Alt, frm1.txtAcctCdTo.Alt)
		frm1.txtAcctCdFr.focus
		Exit Function
	End If

	IntRetCD = SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt)
	If IntRetCD = False Then
	    Exit Function
 	End If

	ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|FiscDt|" & VarFiscDt
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|" & VarBizAreaCd1
	StrUrl = StrUrl & "|AcctCdFr|" & VarAcctCdFr
	StrUrl = StrUrl & "|AcctCdTo|" & VarAcctCdTo

	Call FncEBRPrint(EBAction,ObjName,StrUrl)

End Function


'========================================================================================
Function FncBtnPreview() 
    Dim StrUrl, StrEbrFile
    Dim VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt
	Dim IntRetCD
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field%>
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtDateFr.Text,frm1.txtDateTo.Text,frm1.txtDateFr.Alt,frm1.txtDateTo.Alt, _
	 "970025", frm1.txtDateFr.UserDefinedFormat,Parent.gComDateType, true)=False then
		frm1.txtDateTo.Focus
		Exit Function
	End If

	frm1.txtAcctCdFr.value = Trim(frm1.txtAcctCdFr.value)
	frm1.txtAcctCdTo.value = Trim(frm1.txtAcctCdTo.value)

	If frm1.txtAcctCdFr.value > frm1.txtAcctCdTo.value Then
		Call DisplayMsgBox("970025", "X", frm1.txtAcctCdFr.Alt, frm1.txtAcctCdTo.Alt)
		frm1.txtAcctCdFr.focus
		Exit Function
	End If
	
	IntRetCD = SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt)
	If IntRetCD = False Then
	    Exit Function
 	End If
 	
	ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|FiscDt|" & VarFiscDt
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|" & VarBizAreaCd1
	StrUrl = StrUrl & "|AcctCdFr|" & VarAcctCdFr
	StrUrl = StrUrl & "|AcctCdTo|" & VarAcctCdTo

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
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' 상위 여백 --></TD>
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
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,0)"> <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="사업장명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,3)"> <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="사업장명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>회계일자</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateFr" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=시작회계일자 id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
													   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateTo" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=종료회계일자 id=fpDateTime2></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>시작계정코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctCdFr" SIZE=10 MAXLENGTH=20 tag="12XXXU" ALT="시작계정코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFr" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtAcctCDFr.Value,1)"> <INPUT TYPE="Text" NAME="txtAcctNmFr" SIZE=25 tag="14X" ALT="시작계정명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>종료계정코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctCdTo" SIZE=10 MAXLENGTH=20 tag="12XXXU" ALT="종료계정코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtAcctCDTo.Value,2)"> <INPUT TYPE="Text" NAME="txtAcctNmTo" SIZE=25 tag="14X" ALT="종료계정명"></TD>
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
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
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

