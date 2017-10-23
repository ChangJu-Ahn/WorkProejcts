<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5964MA1
'*  4. Program Name         : 월차 결산 자료발행 
'*  5. Program Desc         : 월차 결산 자료발행 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/01/17
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : song sang min
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance



'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim IsOpenPop          
Dim lsInternal_cd 


'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
         
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

dim StartDate
StartDate	= "<%=GetSvrDate%>"                                               'Get Server DB Date 

Call ggoOper.FormatDate(frm1.txtBas_dt, Parent.gDateFormat,2)

    frm1.txtBas_dt.focus 
	frm1.txtBas_dt.text	= UNIMonthClientFormat(StartDate)
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr 
    Dim strYear,strMonth,strDay
  
    Call ExtractDateFrom(frm1.txtBas_dt.Text,frm1.txtBas_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	Call CommonQueryRs(" A.PAY_TYPE,B.MINOR_NM "," A_BONUS_BASE A, B_MINOR B "," A.PAY_TYPE = B.MINOR_CD AND B.MAJOR_CD = " & FilterVar("H0040", "''", "S") & "  and A.YYYY = " & FilterVar(strYear, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	iCodeArr = lgF0    
    iNameArr = lgF1

	If iCodeArr <> "" Then

			Call SetCombo2(frm1.cboOcpt_type,iCodeArr, iNameArr,Chr(11))
	End If
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitVariables
    Call SetDefaultVal
    'Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    Call InitComboBox
   	Call ggoOper.SetReqAttr(frm1.cboOcpt_type,"Q")
   	Call SetToolbar("1000000000000111")
        
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub


'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncQuery = True                                                              '☜: Processing is OK

End Function


'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False

	FncExit = True
End Function

'======================================================================================================
'	Name : OpenCode()
'	Description : 사업장 
'======================================================================================================
Function OpenCode()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장 팝업"		            ' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA "	                ' TABLE 명칭 
	arrParam(2) = frm1.txtCurrencyCode.value            ' Code Condition
	arrParam(3) = ""   		                    ' Name Cindition
	arrParam(4) = ""        ' Where Condition
	arrParam(5) = "사업장"

   	arrField(0) = "BIZ_AREA_CD"	     		' Field명(1)
    arrField(1) = "BIZ_AREA_NM"					    ' Field명(0)


    arrHeader(0) = "사업장"			    ' Header명(0)
    arrHeader(1) = "사업장명"			' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCurrencyCode.focus
		Exit Function
	Else
		Call SetCode(arrRet)
	End If

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : 사업장코드 Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet)
	With frm1
		.txtCurrencyCode.focus
		.txtCurrencyCode.value = arrRet(0)
		.txtCurrency.value = arrRet(1)
	End With
End Function
'======================================================================================================
'	Name : OpenCodeCon()
'	Description : 월차 구분 코드 
'=======================================================================================================
Function OpenCodeCon()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "월차구분"
	arrParam(1) = "B_MINOR A, B_MAJOR B, A_MONTHLY_BASE C "
	arrParam(2) = frm1.txtReg.value
	arrParam(3) = ""
	arrParam(4) = "A.MINOR_TYPE = " & FilterVar("S", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("A1029", "''", "S") & "  AND a.minor_cd = c.reg_cd AND c.USE_YN = " & FilterVar("Y", "''", "S") & "  "        <%' Where Condition%>
	arrParam(5) = "월차구분"

   	arrField(0) = "A.MINOR_CD"
    arrField(1) = "A.MINOR_NM"


    arrHeader(0) = "월차구분"
    arrHeader(1) = "월차구분명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtReg.focus
		Exit Function
	Else
		Call SetReg(arrRet)
	End If

End Function
'======================================================================================================
'	Name : SetCode()
'	Description : 월차구분코드 Popup에서 Return되는 값 setting
'======================================================================================================
Function SetReg(Byval arrRet)
	With frm1
		.txtReg.focus
		.txtReg.value = arrRet(0)
		.txtRegnm.value = arrRet(1)
	
		If .txtReg.value = "09" Then
	       Call ggoOper.SetReqAttr(frm1.cboOcpt_type,"N")
	    Else
	       .cboOcpt_type.value = ""
			Call ggoOper.SetReqAttr(frm1.cboOcpt_type,"Q")
		End IF
    End With
End Function
'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================

Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
	Dim org_code
    Dim strMin
    Dim strMax     
    Dim bonus

    'Call FuncGetTermDept(lgUsrIntCd,"",strMin,strMax)

    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
     
	'--출력조건을 지정하는 부분 수정 
	select case CStr(frm1.txtReg.value)	
	  Case "01"
		StrEbrFile = "a5964oa1_01"
	  Case "02"
		StrEbrFile = "a5964oa1_02"
	  Case "03"
		StrEbrFile = "a5964oa1_03"
	  Case "04"
		StrEbrFile = "a5964oa1_04"
	  Case "05"
		StrEbrFile = "a5964oa1_05"
	  Case "06"
		StrEbrFile = "a5964oa1_06"
	  Case "07"
		StrEbrFile = "a5964oa1_07"
	  Case "08"
		StrEbrFile = "a5964oa1_08"
	  Case "09"
		StrEbrFile = "a5964oa1_09"			
	  Case "10"
		StrEbrFile = "a5964oa1_10"				
	  Case Else
		Call DisplayMsgBox("970029", "X",frm1.txtReg.alt,"X")
		Exit Function	 		
	End select
		
	Call ExtractDateFrom(frm1.txtBas_dt.Text,frm1.txtBas_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM =   strYear & strMonth
    
	org_code = FilterVar(frm1.txtCurrencyCode.value,"","SNM")
	
	If frm1.cboOcpt_type.value <>"" Then
       bonus = Trim(frm1.cboOcpt_type.value)
    End If   
    
	if org_code = "" then
		org_code = "%"
	End if		
	
		
	'--출력조건을 지정하는 부분 수정 - 끝 
	
'    On Error Resume Next                                                    '☜: Protect system from crashing
    
    '--출력조건을 지정하는 부분 수정 
	
	strUrl = strUrl & "YYYYMM|" & strYYYYMM
	strUrl = strUrl & "|ORG_CODE|" & org_code
	If CStr(frm1.txtReg.value)	="09" Then
		strUrl = strUrl & "|BONUS|" & bonus
		strUrl = strUrl & "|YYYY|" & strYear
	End If
	ObjName = AskEBDocumentName(StrEbrFile, "ebr")
	call FncEBRPrint(EBAction , objName, strUrl)
	'--출력조건을 지정하는 부분 수정 - 끝 
   	
End Function


'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================


Function FncBtnPreview()
'On Error Resume Next                                                    '☜: Protect system from crashing
    Dim strUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
	Dim org_code
	Dim strMin
    Dim strMax     
    Dim bonus
    Dim strYYYYMM1
    'Call FuncGetTermDept(lgUsrIntCd,"",strMin,strMax)

    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
    
   
	select case CStr(frm1.txtReg.value)	
	  Case "01"
		StrEbrFile = "a5964oa1_01"
	  Case "02"
		StrEbrFile = "a5964oa1_02"
	  Case "03"
		StrEbrFile = "a5964oa1_03"
	  Case "04"
		StrEbrFile = "a5964oa1_04"
	  Case "05"
		StrEbrFile = "a5964oa1_05"
	  Case "06"
		StrEbrFile = "a5964oa1_06"
	  Case "07"
		StrEbrFile = "a5964oa1_07"
	  Case "08"
		StrEbrFile = "a5964oa1_08"
	  Case "09"
		StrEbrFile = "a5964oa1_09"			
	  Case "10"
		StrEbrFile = "a5964oa1_10"			
	  Case Else
		
		Call DisplayMsgBox("970029", "X",frm1.txtReg.alt,"X")
		Exit Function	 			
	End select
		
	Call ExtractDateFrom(frm1.txtBas_dt.Text,frm1.txtBas_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
    strYYYYMM =   strYear & strMonth
    
    If frm1.cboOcpt_type.value <>"" Then
       bonus = Trim(frm1.cboOcpt_type.value)
    End If   
    
	org_code = FilterVar(frm1.txtCurrencyCode.value,"","SNM")
	
	if org_code = "" then
		org_code = "%"
	End if		
	
	strUrl = strUrl & "YYYYMM|" & strYYYYMM
	strUrl = strUrl & "|ORG_CODE|" & org_code
	If CStr(frm1.txtReg.value)	="09" Then
	strUrl = strUrl & "|BONUS|" & bonus
	strUrl = strUrl & "|YYYY|" & strYear
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile, "ebr")
	call FncEBRPreview(ObjName , strUrl)

End Function

'========================================================================================================
'   Event Name : txtBas_dt_Onchange()
'   Event Desc : 년도를 직접입력할경우에 상여종류을 설정해준다.
'========================================================================================================
sub txtBas_dt_change()
    Dim reg_code
    Dim iDx    
    If frm1.txtReg.value = "09" Then
		Call ggoOper.SetReqAttr(frm1.cboOcpt_type,"N")

		For iDx = 1 To frm1.cboOcpt_type.length
			frm1.cboOcpt_type.remove(1)
		Next
	
		Call InitComboBox	       
	Else
		frm1.cboOcpt_type.value = ""
		Call ggoOper.SetReqAttr(frm1.cboOcpt_type,"Q")
			       
	End IF
	
	
End sub

'========================================================================================================
'   Event Name : txtReg_Onchange()
'   Event Desc : 월차코드를 직접입력할경우에 월차코드명을 설정해준다.
'========================================================================================================
sub txtReg_Onchange()
    Dim reg_code
    
    
    If frm1.txtReg.value = "09" Then
	      Call ggoOper.SetReqAttr(frm1.cboOcpt_type,"N")
	       
	Else
	       frm1.cboOcpt_type.value = ""
	       Call ggoOper.SetReqAttr(frm1.cboOcpt_type,"Q")
	       
	End IF
	
	reg_code = frm1.txtReg.value
	
	Call CommonQueryRs("A.MINOR_NM","B_MINOR A, B_MAJOR B","A.MINOR_TYPE = " & FilterVar("S", "''", "S") & "  AND A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("A1029", "''", "S") & "  and a.minor_cd = " & FilterVar(reg_code, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	if Trim(Replace(lgF0,Chr(11),"")) = "X" then
	   frm1.txtRegnm.value = ""
	else
	  frm1.txtRegnm.value = Trim(Replace(lgF0,Chr(11),""))
	end if
 
End sub
'========================================================================================================
' Name : txtFr_dept_cd_Onchange()
' Desc : 사업장코드를 직접입력시에 사업장 명을 설정해준다.
'========================================================================================================
	
sub txtFr_dept_cd_Onchange()
    Dim org_code 
    
    org_code = frm1.txtCurrencyCode.value
	Call CommonQueryRs(" BIZ_AREA_NM ","B_BIZ_AREA","BIZ_AREA_CD = " & FilterVar(org_code, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if Trim(Replace(lgF0,Chr(11),"")) = "X" then
	   frm1.txtCurrency.value = ""
	else
	  frm1.txtCurrency.value = Trim(Replace(lgF0,Chr(11),""))
	end if
 
End sub
'========================================================================================================
' Name : txtBas_dt_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtBas_dt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBas_dt.Action = 7
 		Call SetFocusToDocument("M")
		Frm1.txtBas_dt.Focus
	End If
End Sub
'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

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

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>월차결산자료발행</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
				   <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>   
							
							<TR>
							    <TD CLASS=TD5 NOWRAP>작업 년월</TD>
							    <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5964ma1_fpDateTime1_txtBas_dt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtFr_dept_cd" NAME="txtCurrencyCode" SIZE=10 MAXLENGTH=10  tag="12XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode()">
								                       <INPUT TYPE="Text" NAME="txtCurrency" SIZE=20 MAXLENGTH=40  tag="14XXXU" ALT="사업장명">
							</TR>	
							<TR>
								<TD CLASS="TD5" NOWRAP>월차 구분</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtReg" NAME="txtReg" SIZE=10 MAXLENGTH=10  tag="12XXXU" ALT="월차구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCodeCon">
								                       <INPUT TYPE="Text" NAME="txtRegnm" SIZE=20 MAXLENGTH=40  tag="14XXXU" ALT="월차구분명">
							</TR>	
							<TR>
							    <TD CLASS=TD5 NOWRAP>상여 종류</TD>
							    <TD CLASS=TD6 NOWRAP><SELECT NAME="cboOcpt_type" ALT="상여 종류" STYLE="WIDTH: 120px" TAG=13><OPTION VALUE=""></OPTION></SELECT></TD>	
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
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
                         <BUTTON NAME="btnRun"   CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
                         <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>
		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=100 <%=BizSize%> FRAMEBORDER=0 SCROLLING=yse noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1">
</FORM>
</BODY>
</HTML>

