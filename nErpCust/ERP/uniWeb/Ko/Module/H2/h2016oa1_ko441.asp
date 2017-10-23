<%@ LANGUAGE="VBSCRIPT" %> 
<!--
'======================================================================================================
'*  1. Module Name          : 인사/급여관리 
'*  2. Function Name        : 인사기본자료관리 
'*  3. Program ID           : h2016oa1
'*  4. Program Name         : 재직증명서출력 
'*  5. Program Desc         : 재직증명서출력 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/05/27
'*  8. Modified date(Last)  : 2003/06/10
'*  9. Modifier (First)     : Song Myeongsik
'* 10. Modifier (Last)      : Lee SiNa
'* 11. Comment              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBsCRIPT"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit 

<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgOldRow

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
        
End Sub
'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtProv_dt.focus

	frm1.txtProv_dt.Year = strYear 		 '년월일default value setting
	frm1.txtProv_dt.Month = strMonth 
	frm1.txtProv_dt.Day = strDay 
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call InitVariables
   
    Call ggoOper.FormatDate(frm1.txtProv_dt, gDateFormat, 1)  
    
    Call FuncGetAuth(gStrRequestMenuID , Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
                                          
    Call SetDefaultVal
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
    
    If  txtEmp_no_Onchange() then
        Exit Function
    End If
    
    
    FncQuery = True                                                              '☜: Processing is OK

End Function

'======================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                        '☜: Protect system from crashing
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


'========================================================================================================
' Name : OpenEmptName()
' Desc : developer describe this line(grid외에서 사용) 
'========================================================================================================
Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(3)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value		' Name Cindition
	End If

    arrParam(2) = lgUsrIntcd
    
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
        frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmpName(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmpName()
'	Description : Item Popup에서 Return되는 값 setting(grid외에서 사용)
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		.txtEmp_no.focus
		Set gActiveElement = document.ActiveElement

		lgBlnFlgChgValue = False
	End With
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================%>

Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
	dim ObjName    
	
    If Not chkField(Document, "1") Then			'⊙: This function check indispensable field%>
       Exit Function
    End If

	Dim emp_no, use, title, prov_dt, emp_name ,print_emp_no, etc
	
	If frm1.txtHprt_type(0).checked Then 
		title = "1"
		StrEbrFile = "h2016oa1_ko441"
	Elseif frm1.txtHprt_type(1).checked Then 
		title = "2"
		StrEbrFile = "h2016oa2_ko441"		
    Else 
		title = "3"
		StrEbrFile = "h2016oa1_1"
	End if

	emp_name = frm1.txtname.value
	prov_dt=UniConvDateAToB(frm1.txtProv_dt.Text,gDateFormat,Parent.gServerDateFormat)
	
	emp_no = frm1.txtEmp_no.value
	use = frm1.txtUse.value
	
	etc = frm1.txtEtc.value
	
	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if	
	
	if Trim(etc) = "" then
		etc = "#space#"
	
	End if	
	
	if len(etc) - len(replace(etc,vbCR,"")) > 6 then
		Call DisplayMsgBox("800500","X","X","X")	' 비고는 7줄까지 입력가능합니다 
		exit function
	end if
	If  txtEmp_no_Onchange() then
        Exit Function
    End If
    
    if valid_emp(emp_no, prov_dt) = false then
		Call DisplayMsgBox("800499","X","X","X")	' 증명자격에 해당하지 않는 사원입니다 
		exit function
	end if

	strUrl = "title|" & title
	strUrl = strUrl & "|emp_name|" & emp_name
	strUrl = strUrl & "|prov_dt|" & prov_dt
	strUrl = strUrl & "|emp_no|" & emp_no 
	strUrl = strUrl & "|print_emp_no|" & Parent.gUsrID
	strUrl = strUrl & "|use|" & use
	strUrl = strUrl & "|Etc|" & etc

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	call FncEBRPrint(EBAction,ObjName, strUrl)	

End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview() 
	Dim strUrl
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	dim ObjName		
	Dim emp_no, use, title, prov_dt, emp_name ,print_emp_no ,etc
    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>

       Exit Function
    End If	
	
	If frm1.txtHprt_type(0).checked Then 
		title = "1"
		StrEbrFile = "h2016oa1_ko441"
	Elseif frm1.txtHprt_type(1).checked Then 
		title = "2"
		StrEbrFile = "h2016oa2_ko441"		
    Else 
		title = "3"
		StrEbrFile = "h2016oa1_1"
	End if
	
	emp_name = frm1.txtname.value
	prov_dt=UniConvDateAToB(frm1.txtProv_dt.Text,gDateFormat,Parent.gServerDateFormat)
	emp_no = frm1.txtEmp_no.value
	use = frm1.txtUse.value
	etc = frm1.txtEtc.value
	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if
	
	
	if Trim(etc) = "" then
		etc = "#space#"   ' 수정하지 말것 
	End if	
	
	if len(etc) - len(replace(etc,vbCR,"")) > 6 then
		Call DisplayMsgBox("800500","X","X","X")	' 비고는 7줄까지 입력가능합니다 
		exit function
	end if

	If  txtEmp_no_Onchange() then
        Exit Function
    End If
    
    if valid_emp(emp_no, prov_dt) = false then
		Call DisplayMsgBox("800499","X","X","X")	' 증명자격에 해당하지 않는 사원입니다 
		exit function
	end if
	
	
	strUrl = "title|" & title
	strUrl = strUrl & "|emp_name|" & emp_name
	strUrl = strUrl & "|prov_dt|" & prov_dt
	strUrl = strUrl & "|emp_no|" & emp_no 
	strUrl = strUrl & "|print_emp_no|" & Parent.gUsrID
	strUrl = strUrl & "|use|" & use
	strUrl = strUrl & "|Etc|" & etc
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	call FncEBRPreview(ObjName , strUrl)

End Function

'========================================================================================================
'   Event Name : txtEmp_no_Onchange           
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
			frm1.txtName.value = ""
            Frm1.txtEmp_no.focus 
            Set gActiveElement = document.ActiveElement
			txtEmp_no_Onchange = true
        Else
			frm1.txtName.value = strName
        End if 
    End if  
End Function

'=======================================================================================================
'   Event Name : txt________Keypress
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtEmp_no_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub     


'========================================================================================================
' Name : txtProv_dt_DblClick
' Desc : 달력 Popup을 호출 
'========================================================================================================
Sub txtProv_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtProv_dt.Action = 7
		frm1.txtProv_dt.focus
	End If
End Sub



function valid_emp(emp_no, prov_dt )
	Dim strWhere
	Dim retCD
	valid_emp = false
	
	if frm1.Rb_jaejik.checked = true then
		strWhere = " (Retire_dt is null OR Retire_dt > " & FilterVar(prov_dt , "''", "S")  & ")"
		strWhere = strWhere & " AND entr_dt <= " & FilterVar(prov_dt , "''", "S") 
		strWhere = strWhere & " AND emp_no = " & FilterVar(emp_no, "''", "S")
		retCD = CommonQueryRs(" emp_no "," HAA010T ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		if ((retCd = true) and Trim(replace(lgF0,chr(11),"")) = Trim(emp_no))  then
			valid_emp = true
		end if	
	else
		strWhere = " (Retire_dt is Not null AND Retire_dt <= " & FilterVar(prov_dt , "''", "S")  & ")"
		strWhere = strWhere & " AND entr_dt <= " & FilterVar(prov_dt , "''", "S") 
		strWhere = strWhere & " AND emp_no = " & FilterVar(emp_no, "''", "S")
		retCD = CommonQueryRs(" emp_no "," HAA010T ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		if ((retCd = true) and Trim(replace(lgF0,chr(11),"")) = Trim(emp_no))  then
			valid_emp = true
		end if	
	
	end if
	
end function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재직증명서출력</font></td>
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
					<TD HEIGHT=100>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						    <TR>
								<TD CLASS="TD5" NOWRAP>기준일</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h2016oa1_txtProv_dt_txtProv_dt.js'></script></TD> 
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>용도</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID="txtUse" NAME="txtUse" SIZE=30 MAXLENGTH=100 tag="12XXXU"  ALT="용도"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>제출처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID="txtEtc" NAME="txtEtc" SIZE=30 MAXLENGTH=100 tag="12XXXU"  ALT="제출처"></TD>
							</TR>
							<!--<TR>
								<TD CLASS="TD5" NOWRAP>비고</TD>
								<TD CLASS="TD6" NOWRAP><TEXTAREA rows=6 cols=80  ID="txtEtc" NAME="txtEtc" maxrows=5 size=200 tag="2XXXXU"  ALT="비고"></TEXTAREA></TD>
							</TR>-->
							<TR>			
							    <TD CLASS="TD5" NOWRAP>대상자</TD>
							    <TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID="txtEmp_no" NAME="txtEmp_no" SIZE=13 MAXLENGTH=13  tag="12XXXU" ALT="대상자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmptName (1)">
								                       <INPUT TYPE="Text" ID="txtName" NAME="txtName" SIZE=20 MAXLENGTH=30  tag="14XXXU" ALT="대상자코드명"></TD>
	                        </TR>
	                        <TR>
								<TD CLASS="TD5" NOWRAP>출력구분</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=txtHprt_type VALUE = "1" ID=Rb_jaejik TAG="11" Checked><LABEL FOR=Rb_jaejik>재직</LABEL>&nbsp;
													   <INPUT TYPE="RADIO" CLASS="Radio" NAME=txtHprt_type VALUE = "2" ID=Rb_career TAG="11"><LABEL FOR=Rb_career>경력</LABEL>&nbsp;
													   <INPUT TYPE="RADIO" CLASS="Radio" NAME=txtHprt_type VALUE = "3" ID=Rb_retire TAG="11"><LABEL FOR=Rb_retire>퇴직</LABEL></TD>
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
		                <BUTTON NAME="btnPreview" CLASS="CLSSBTN" onclick="VBScript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
		                <BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()" Flag=1>인쇄</BUTTON>

		            </TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=0><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM>
</BODY>
</HTML>


