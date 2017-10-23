<%@ LANGUAGE="VBSCRIPT" %>
<!-- '======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Closing and Financial Statements
'*  3. Program ID           : A5212OA1
'*  4. Program Name         : 보조부항목 check list
'*  5. Program Desc         : Report of Subledger Detail
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/09/09
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : kim ho young
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">				</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit																	'☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop


'========================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
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
Function OpenPopUp(Byval param, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Select Case iWhere
		Case 0, 7
			arrParam(0) = "사업장코드 팝업"								' 팝업 명칭 
			arrParam(1) = "B_BIZ_AREA" 										' TABLE 명칭 
			arrParam(2) = param												' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "사업장코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "BIZ_AREA_CD"										' Field명(0)
			arrField(1) = "BIZ_AREA_NM"										' Field명(1)

			arrHeader(0) = "사업장코드"									' Header명(0)
			arrHeader(1) = "사업장명"									' Header명(1)

		Case 1, 2

			arrParam(0) = "계정코드 팝업"								' 팝업 명칭 
			arrParam(1) = " A_ACCT A, A_ACCT_GP B  "										' TABLE 명칭 
			arrParam(2) = Trim(param)										' Code Condition
			arrParam(3) = ""												' Name Cindition

			arrParam(4) = "ISNULL(A.SUBLEDGER_1,'') <> '' AND A.GP_CD=B.GP_CD"	' Where Condition


			arrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "A.ACCT_CD"										' Field명(0)
			arrField(1) = "A.ACCT_NM"										' Field명(1)
     		arrField(2) = "B.GP_CD"											' Field명(2)
			arrField(3) = "B.GP_NM"											' Field명(3)

			arrHeader(0) = "계정코드"									' Header명(0)
			arrHeader(1) = "계정명"										' Header명(1)
			arrHeader(2) = "그룹코드"									' Header명(2)
			arrHeader(3) = "그룹명"

		Case 3
			arrParam(0) = "보조부항목 팝업"								' 팝업 명칭 
			arrParam(1) = "A_CTRL_ITEM A"									' TABLE 명칭 
			arrParam(2) = Trim(param)										' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "A.CTRL_CD in (select subledger_1 from A_Acct) "	' Where Condition
			arrParam(5) = "관리항목코드"									' 조건필드의 라벨 명칭 

			arrField(0) = "A.CTRL_CD"						' Field명(0)
			arrField(1) = "A.CTRL_NM"						' Field명(1)

			arrHeader(0) = "관리항목코드"					' Header명(0)
			arrHeader(1) = "관리항목명"						' Header명(1)

		Case 4
			arrParam(0) = Trim(frm1.txtCtrlNm.value)							' 팝업 명칭 
			arrParam(1) = Trim(frm1.hTblId.value) 
			arrParam(2) = ""												' Code Condition
			arrParam(3) = ""												' Name Cindition
			
			arrParam(4) = Trim(frm1.hDataColmID.value) & _
					" in (select distinct CTRL_Val1 from A_SUBLEDGER_SUM where convert(datetime,fisc_yr+fisc_mnth+(case when fisc_dt in (" & FilterVar("00", "''", "S") & " ," & FilterVar("99", "''", "S") & " ) then " & FilterVar("01", "''", "S") & "  else fisc_dt end),112) between '" & _
					 UniConvDateToYYYYMMDD(frm1.txtDateFr.Text,Parent.gDateFormat,"") & "' and '" & _
					 UniConvDateToYYYYMMDD(frm1.txtDateTo.Text,Parent.gDateFormat,"") & "'	 )"	' Where Condition

			arrParam(5) = Trim(frm1.txtCtrlNm.value)									' 조건필드의 라벨 명칭 

			arrField(0) = Trim(frm1.hDataColmID.value)			' Field명(0)
			arrField(1) = Trim(frm1.hDataColmNm.value)						' Field명(1)

			arrHeader(0) = Trim(frm1.hDataColmID.value)					' Header명(0)
			arrHeader(1) = Trim(frm1.hDataColmNm.value)						' Header명(1)

		Case 5
			arrParam(0) = Trim(frm1.txtCtrlNm.value)							' 팝업 명칭 
			arrParam(1) = "A_ACCT A,A_SUBLEDGER_SUM B"
			arrParam(2) = ""												' Code Condition
			arrParam(3) = ""												' Name Cindition

			arrParam(4) = " A.SUBLEDGER_1 = " & FilterVar(frm1.txtCtrlCd.value, "''", "S")  & " and " & _
						" a.acct_cd = b.acct_cd and convert(datetime,b.fisc_yr+b.fisc_mnth+(case when b.fisc_dt in (" & FilterVar("00", "''", "S") & " ," & FilterVar("99", "''", "S") & " ) then " & FilterVar("01", "''", "S") & "  else b.fisc_dt end),112) between '" & _
					 UniConvDateToYYYYMMDD(frm1.txtDateFr.Text,Parent.gDateFormat,"") & "' and '" & _
					 UniConvDateToYYYYMMDD(frm1.txtDateTo.Text,Parent.gDateFormat,"") & "'	 "	' Where Condition

			arrParam(5) = Trim(frm1.txtCtrlNm.value)									' 조건필드의 라벨 명칭 

			arrField(0) = "b.ctrl_val1"			' Field명(0)
			arrField(1) = ""

			arrHeader(0) = Trim(frm1.txtCtrlNm.value)					' Header명(0)
			arrHeader(1) = ""
		Case Else
			Exit Function
	End Select

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0	'사업장코드 
				frm1.txtBizAreaCd.focus
			Case 3
				frm1.txtCtrlCd.focus
			Case 4
				frm1.txtCtrlVal.focus
			Case 5
				frm1.txtCtrlVal.focus
			Case 7	'사업장코드 
				frm1.txtBizAreaCd1.focus
		End select
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If	

End Function


'========================================================================================================= 
Function SetReturnPopUp(ByVal arrRet, ByVal iWhere)
	
	Select Case iWhere
		Case 0	'사업장코드 
			frm1.txtBizAreaCd.focus
			frm1.txtBizAreaCd.value = arrRet(0)
			frm1.txtBizAreaNm.value = arrRet(1)
		Case 3
			frm1.txtCtrlCd.focus
			frm1.txtCtrlCd.value = arrRet(0)
			frm1.txtCtrlNm.value = arrRet(1)

			CtrlVal.innerHTML = frm1.txtCtrlNm.value 
			frm1.txtCtrlVal.value	= ""
			frm1.txtCtrlValNm.value	= ""
			Call ElementVisible(frm1.txtCtrlVal, 1)
			Call ElementVisible(frm1.txtCtrlValNm, 1)
			Call ElementVisible(frm1.btnCtrlVal, 1)
		Case 4
			frm1.txtCtrlVal.focus
			frm1.txtCtrlVal.value = arrRet(0)
			frm1.txtCtrlValNm.value = arrRet(1)	
		Case 5
			frm1.txtCtrlVal.focus
			frm1.txtCtrlVal.value = arrRet(0)
		Case 7	'사업장코드 
			frm1.txtBizAreaCd1.focus
			frm1.txtBizAreaCd1.value = arrRet(0)
			frm1.txtBizAreaNm1.value = arrRet(1)
		Case Else
			Exit Function
	End select

End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function QueryCtrlVal()

    Dim ArrRet

    IF frm1.txtCtrlCd.value = "" Then
		Call DisplayMsgBox("205152", "X", "보조부항목","X")
		frm1.txtCtrlCd.focus
	END IF

    Call CommonQueryRs( "TBL_ID,DATA_COLM_ID,DATA_COLM_NM" , _ 
				"A_CTRL_ITEM" , _
				 "CTRL_CD = " & FilterVar(frm1.txtCtrlCd.value, "''", "S"), _ 
				 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)


	ArrRet 	= Split(lgF0,Chr(11))

	IF Trim(ArrRet(0)) <> "" then
		frm1.hTblId.value  = ArrRet(0)
		
		ArrRet 	= Split(lgF1,Chr(11))
		frm1.hDataColmID.value  = ArrRet(0)
		ArrRet 	= Split(lgF2,Chr(11))
		frm1.hDataColmNm.value = ArrRet(0)

		Call OpenPopUp(0,4)
	ELSE
		Call OpenPopUp(0,5)
	END IF

End Function


'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitVariables
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")
	frm1.txtCtrlCd.focus
	Call ElementVisible(frm1.txtCtrlVal, 0)
	Call ElementVisible(frm1.txtCtrlValNm, 0)
	Call ElementVisible(frm1.btnCtrlVal, 0)
End Sub

'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'======================================================================================================
'   Event Name : txtDateFr_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
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

'========================================================================================
Sub txtCtrlCd_OnBlur()
	
	On error Resume next
	Dim ArrRet
	Dim ArrParam(2)
	
  
    Call CommonQueryRs( "CTRL_CD,CTRL_NM" , _ 
				"A_CTRL_ITEM", _
				 "CTRL_CD = " & FilterVar(frm1.txtCtrlCd.value, "''", "S"), _
				 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
	
	ArrRet 	= Split(lgF0,Chr(11))

	IF ArrRet(0) = "" Then
		frm1.txtCtrlNm.value = ""
		
		CtrlVal.innerHTML = ""
		Call ElementVisible(frm1.txtCtrlVal, 0)
		Call ElementVisible(frm1.txtCtrlValNm, 0)
		Call ElementVisible(frm1.btnCtrlVal, 0)

		Exit Sub	
	END IF
	ArrParam(0) = ArrRet(0)
	ArrRet 	= Split(lgF1,Chr(11))
	ArrParam(1) = ArrRet(0)

	frm1.txtCtrlCd.value = ArrParam(0)
	frm1.txtCtrlNm.value = ArrParam(1)
	
	CtrlVal.innerHTML = frm1.txtCtrlNm.value 
	frm1.txtCtrlVal.value	= ""
	frm1.txtCtrlValNm.value	= ""

	Call ElementVisible(frm1.txtCtrlVal, 1)
	Call ElementVisible(frm1.txtCtrlValNm, 1)
	Call ElementVisible(frm1.btnCtrlVal, 1)

End Sub


'========================================================================================
Sub SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt, VarCtrlCd, VarCtrlVal)
	
	VarDateFr = UniConvDateToYYYYMMDD(frm1.txtDateFr.Text,Parent.gDateFormat,"")
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txtDateTo.Text,Parent.gDateFormat,"")

	VarCtrlCd	= FilterVar(Trim(frm1.txtCtrlCd.value),"","SNM")
	
	VarCtrlVal	= "%"

	VarAcctCdFr = "0"
	VarAcctCdTo = "ZZZZZZZZZZ"
		
	'VarBizAreaCd = "%"

	If Len(frm1.txtCtrlVal.value) > 0 Then 
		VarCtrlVal = FilterVar(Trim(frm1.txtCtrlVal.value),"","SNM")
	Else 
		frm1.txtCtrlValNm.value = ""
	End If

	If frm1.txtBizAreaCd.value = "" then 
		frm1.txtBizAreaNm.value = ""
		VarBizAreaCd = "0"
	else 
		VarBizAreaCd = FilterVar(frm1.txtBizAreaCD.value,"","SNM")
	end if
	
	If frm1.txtBizAreaCd1.value = "" then
		frm1.txtBizAreaNm1.value = ""
		VarBizAreaCd1 = "ZZZZZZZZZZ"
	else 
		VarBizAreaCd1 = FilterVar(frm1.txtBizAreaCD1.value,"","SNM")
	end if

	StrEbrFile = "a5211oa3"
'	if frm1.PrintOpt2.checked = True and frm1.txtBizAreaCd.value <> "" then
'		StrEbrFile = "a5211oa2"
'	elseif  frm1.PrintOpt2.checked = True and frm1.txtBizAreaCd.value = "" then
'		StrEbrFile = "a5211oa3"
'	end if
	

End Sub

'========================================================================================
Function FncBtnPrint() 
    Dim StrUrl
    Dim StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt, VarCtrlCd, VarCtrlVal
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
	
	'회계일자 조회기간 Check
	If UniConvDateToYYYYMMDD(frm1.txtDateFr.Text, Parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, Parent.gDateFormat, "") Then
		Call DisplayMsgBox("970025", "X", frm1.txtDateFr.Alt, frm1.txtDateTo.Alt)
		frm1.txtDateFr.focus
		Exit Function
	End If
	
	
	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt, VarCtrlCd, VarCtrlVal)
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")	

	StrUrl = StrUrl & "DateFr|" & VarDateFr					' '|'-> ebr파일을 부를때 사용되는 구분자.(url에서 ?뒤에파라메터로 붙여주는 것이라고 보면 됨.
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|" & VarBizAreaCd1
'	StrUrl = StrUrl & "|FiscDt|" & VarFiscDt
	StrUrl = StrUrl & "|Currency|" & Parent.gCurrency

	StrUrl = StrUrl & "|CtrlCd|" & VarCtrlCd
	StrUrl = StrUrl & "|CtrlVal|" & VarCtrlVal

	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
		
End Function


'========================================================================================
Function FncBtnPreview() 

    Dim StrUrl
    Dim StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt, VarCtrlCd, VarCtrlVal
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
	
	'회계일자 조회기간 Check
	If UniConvDateToYYYYMMDD(frm1.txtDateFr.Text, Parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, Parent.gDateFormat, "") Then
		Call DisplayMsgBox("970025", "X", frm1.txtDateFr.Alt, frm1.txtDateTo.Alt)
		frm1.txtDateFr.focus
		Exit Function
	End If
	
	
	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarBizAreaCd, VarBizAreaCd1, VarAcctCdFr, VarAcctCdTo, VarFiscDt, VarCtrlCd, VarCtrlVal)
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|BizAreaCd1|" & VarBizAreaCd1
	StrUrl = StrUrl & "|Currency|" & Parent.gCurrency

	StrUrl = StrUrl & "|CtrlCd|" & VarCtrlCd
	StrUrl = StrUrl & "|CtrlVal|" & VarCtrlVal

	Call FncEBRPreview(ObjName,StrUrl)
		
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
	Call parent.FncFind(Parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
Function FncExit()
    FncExit = True
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    //alert(value)
    //alert(escape(value))
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
 -->
</SCRIPT>

</HEAD>


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><% ' 상위 여백 %></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>보조부항목CheckList</font></td>
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
								<TD CLASS="TD5" NOWRAP>회계일자</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/a5212oa1_fpDateTime1_txtDateFr.js'></script>&nbsp;~&nbsp;
													   <script language =javascript src='./js/a5212oa1_fpDateTime2_txtDateTo.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" ID="CtrlCd" NOWRAP>보조부항목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtCtrlCd" SIZE=10 MAXLENGTH=20 tag="12XXXU" ALT="보조부항목" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCtrlCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCtrlCd.Value,3)"> <INPUT TYPE="Text" NAME="txtCtrlNm" SIZE=25 tag="14X" ALT="보조부항목명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" ID="CtrlVal" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtCtrlVal" SIZE=10 MAXLENGTH=20 tag="11XXXU" ALT="" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCtrlVal" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call QueryCtrlVal()"> <INPUT TYPE="Text" NAME="txtCtrlValNm" SIZE=25 tag="14X" ALT=""></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,0)"> <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="14X" ALT="사업장명"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="사업장코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,7)"> <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="14X" ALT="사업장명"></TD>
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
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hTblId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hDataColmID" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hDataColmNm" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
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

