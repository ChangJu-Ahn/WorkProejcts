
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : Template
'*  2. Function Name        : Template
'*  3. Program ID           : Template
'*  4. Program Name         : Template
'*  5. Program Desc         : Template
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2004/04/
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 


Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    
End Sub

Sub SetDefaultVal()


	Dim strFrDate, strToDate
	strFrDate	= Parent.gFiscStart
	strToDate	= "<%=GetSvrDate%>"

	frm1.txtFrYyyymm.text	= UniConvDateAToB(strFrDate,Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtToYyyymm.text	= UniConvDateAToB(strToDate,Parent.gServerDateFormat,Parent.gDateFormat)

	'Call ggoOper.FormatDate(frm1.txtFrYyyymm, Parent.gDateFormat, 2)
	'Call ggoOper.FormatDate(frm1.txtToYyyymm, Parent.gDateFormat, 2)
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

' 모듈 코드로 입력 예:원가 "C"
<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "OA") %>
End Sub

Sub InitComboBox()

	'전표입력경로 Combo
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F3012", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboDpstType, lgF0  ,lgF1  ,Chr(11))

End Sub


Function OpenPopup(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	if iWhere = 1 then
		arrParam(0) = "은행코드팝업"
		arrParam(1) = "b_bank"
		arrParam(2) = strCode
		arrParam(3) = ""
		arrParam(4) = ""

		arrParam(5) = "은행코드"			
	
		arrField(0) = "BANK_CD"
		arrField(1) = "BANK_NM"
		 
		arrHeader(0) = "은행코드"
		arrHeader(1) = "은행코드명"
	
	elseif iWhere = 2 then
	
		arrParam(0) = "통화코드팝업"
		arrParam(1) = "b_currency"
		arrParam(2) = strCode
		arrParam(3) = ""
		arrParam(4) = ""

		arrParam(5) = "통화코드"			
	
		arrField(0) = "CURRENCY"
		arrField(1) = "CURRENCY_DESC"
		 
		arrHeader(0) = "통화코드"
		arrHeader(1) = "통화코드명"		
	else
		Exit Function
	end if
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 1 Then
			frm1.txtBankCd.focus		
		elseif iWhere = 2 Then
			frm1.txtCurCd.focus 
		End If		
		Exit Function			
	Else
		Call SetPopup(arrRet, iWhere)
	End If	

End Function

Function SetPopup(Byval arrRet, Byval iWhere)
	
	With frm1
	
    	If iWhere = 1 Then
    		.txtBankCd.value  = arrRet(0)
    		.txtBankNm.value  = arrRet(1)    		

    	ElseIf iWhere = 2 Then
    		.txtCurCd.focus 
    		.txtCurCd.value = arrRet(0)
    		.txtCurNm.value = arrRet(1)

    	Else
			Exit Function
    	End If
	End With
	
End Function


Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitVariables
    Call SetDefaultVal

	Call InitComboBox
    Call SetToolbar("10000000000011")
    
    frm1.txtFrYyyymm.focus 
    Set gActiveElement = document.activeElement	


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
    
End Sub

Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

Sub txtFrYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrYyyymm.focus
	End If
End Sub

Sub txtToYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtToYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToYyyymm.focus
	End If
End Sub

Sub txtFrYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

Sub txtToYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

Function FncQuery()
	' 엔터키 입력시 미리보기 실행 
    FncBtnPreview()
End Function

Function SetPrintCond(StrEbrFile, strUrl)

	Dim	strFrYear,     strFrMonth,  strFrDay
	Dim	strToYear,     strToMonth,  strToDay
	Dim strBankCd,	   strCurCd,  strDpstType,strDpstTypeNm, strBanknm
	Dim strFrYyyymm,   strToYyyymm
	Dim	strAuthCond
    
    SetPrintCond = False
    
	''StrEbrFile = "F3108OA1"

	' 시작월 종료월 비교 
    If CompareDateByFormat(frm1.txtFrYyyymm.Text,frm1.txtToYyyymm.Text,frm1.txtFrYyyymm.Alt,frm1.txtToYyyymm.Alt, _
	 "970024", frm1.txtFrYyyymm.UserDefinedFormat,Parent.gComDateType, true)=False then
		frm1.txtToYyyymm.Focus
		Exit Function
	End If

	Call ExtractDateFrom(frm1.txtFrYyyymm.Text,frm1.txtFrYyyymm.UserDefinedFormat,Parent.gComDateType,strFrYear,strFrMonth,strFrDay)
	Call ExtractDateFrom(frm1.txtToYyyymm.Text,frm1.txtToYyyymm.UserDefinedFormat,Parent.gComDateType,strToYear,strToMonth,strToDay)

	strFrYyyymm		= strFrYear & strFrMonth & strFrDay
	strToYyyymm		= strToYear & strToMonth & strToDay
	
	strBankCd		= Trim(UCase(frm1.txtBankCd.value))
	strCurCd		= Trim(UCase(frm1.txtCurCd.value))
	strDpstType		= Trim(UCase(frm1.cboDpstType.value))
	
	if strBankCd = "" then
		strBankCd = "%"
		strBankNm = "%"
	else		
		Call CommonQueryRs("BANK_NM "," B_BANK "," BANK_CD =  " & FilterVar(strBankCd , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		strBankNm = Trim(Replace(lgF0,Chr(11),"")) 
		lgF0 = ""
	End if
	
	If strCurCd = "" then
		strCurCd = "%"
		StrEbrFile = "F3108OA2"   
	Else
		StrEbrFile = "F3108OA1"	
	End if
	
	if Len(strDpstType) = 0 then	
		strDpstType = "%"
		strDpstTypeNm = "%"
	else		
		Call CommonQueryRs("MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F3012", "''", "S") & "  AND MINOR_CD =  " & FilterVar(strDpstType , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		strDpstTypeNm = Trim(Replace(lgF0,Chr(11),""))
		lgF0 = ""
	End if

	' 권한관리 추가 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_DPST_ITEM.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND F_DPST_ITEM.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_DPST_ITEM.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND F_DPST_ITEM.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	strUrl	= strUrl & "trans_dt_fr|"	& strFrYyyymm
	strUrl	= strUrl & "|trans_dt_to|"	& strToYyyymm
	strUrl	= strUrl & "|bank_cd|"		& strBankCd
	strUrl	= strUrl & "|doc_cur|"		& strCurCd
	strUrl	= strUrl & "|dpst_Type|"	& strDpstType
	strUrl	= strUrl & "|dpst_Type_Nm|"	& strDpstTypeNm
	strUrl	= strUrl & "|bank_nm|"		& strBanknm

	StrUrl	= StrUrl & "|strAuthCond|"	& strAuthCond

	SetPrintCond = True
			
End Function



Function FncBtnPreview() 
    
    Dim StrEbrFile, strUrl

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If

	ObjName = AskEBDocumentName(StrEbrFile,"ebc")
	call FncEBCPreview(ObjName , strUrl)
	
End Function

Function FncPrint() 
    Call parent.FncPrint()
End Function

Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLE,False)
End Function

Function FncExit()
    FncExit = True
End Function

Function txtBankCd_onChange()
	if Trim(frm1.txtBankCd.value) = "" then
		frm1.txtBankNm.value = ""
	end if 
End Function

Function txtCurCd_onChange()
	if Trim(frm1.txtCurCd.value) = "" then
		frm1.txtCurNm.value = ""
	end if 
End Function

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
								<TD CLASS="TD5" NOWRAP>입출일자</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtFrYyyymm CLASS=FPDTYYYYMM title=FPDATETIME tag="12X1" ALT="시작년월"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtToYyyymm CLASS=FPDTYYYYMM title=FPDATETIME tag="12X1" ALT="종료년월"></OBJECT>');</SCRIPT></TD>
							</TR>		
							<TR>	
								<TD CLASS="TD5" NOWRAP>은행</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBankCd" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT="은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFrItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup frm1.txtBankCd.value, 1">&nbsp;<INPUT TYPE=TEXT NAME="txtBankNm" SIZE=30 tag="14"></TD>
							</TR>											
							<TR>	
								<TD CLASS="TD5" NOWRAP>통화</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCurCd" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT="통화"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFrItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup frm1.txtCurCd.value, 2">&nbsp;<INPUT TYPE=TEXT NAME="txtCurNm" SIZE=30 tag="14"></TD>
							</TR>		
							<TR>
								<TD CLASS="TD5" NOWRAP>예적금유형</TD>
								<TD CLASS="TD6" NOWRAP><SELECT ID="cboDpstType" NAME="cboDpstType" ALT="예적금유형" STYLE="WIDTH: 120px" tag="11X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
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
					<TD>
						<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="strUrl" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1" >	
</FORM>
</BODY>
</HTML>

