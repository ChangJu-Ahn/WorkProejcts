<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : 인사/급여관리 
'*  2. Function Name        : 퇴직정산관리/퇴직금 
'*  3. Program ID           : ha107oa1
'*  4. Program Name         : 퇴직금지급명세서출력 
'*  5. Program Desc         : 퇴직금지급명세서출력 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/06/11
'*  8. Modified date(Last)  : 2003/05/27
'*  9. Modifier (First)     : mok young bin
'* 10. Modifier (Last)      : Lee Si Na
'* 11. Comment              :
=======================================================================================================-->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBsCRIPT"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit 
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	Dim strYear,strMonth,strDay
	
	frm1.txtStand_yy.focus 
	Call ExtractDateFrom("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gServerDateType,strYear,strMonth,strDay)	
    Call ggoOper.FormatDate(frm1.txtStand_yy, parent.gDateFormat, 3)
    frm1.txtStand_yy.Year	= strYear
    frm1.txtStand_yy.Month	= strMonth
    frm1.txtStand_yy.Day	= strDay
        
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "OA") %>
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call InitVariables
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

    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    FncQuery = True                                                              '☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
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
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, False)
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
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	Else 'spread

	End If
	arrParam(2) = ""

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SubSetCondEmp(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SubSetCondEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		End If
	End With
End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================%>
Function FncBtnPrint() 

	Dim condvar
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim strFrDt
    Dim strToDt
	Dim ObjName    	
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If

    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if    
    strFrDt = frm1.txtFr_retire_dt.Text
    strToDt = frm1.txtTo_retire_dt.Text
    
    If Trim(strFrDt) = "" then
        strFrDt = UniConvYYYYMMDDToDate(parent.gDateFormat,"1950","01","01")
    End if
    If Trim(strToDt) = "" then
        strToDt = UniConvYYYYMMDDToDate(parent.gDateFormat,"2100","12","31")
    End If
	If CompareDateByFormat(Trim(strFrDt),Trim(strToDt),frm1.txtFr_retire_dt.Alt,frm1.txtTo_retire_dt.Alt,"800002",parent.gDateFormat,parent.gComDateType,True) = False Then
        frm1.txtFr_retire_dt.focus
        Set gActiveElement = document.activeElement
        Exit Function
    End if 
        
	dim  to_retire_dt, fr_dept_cd,  stand_yy, fr_retire_dt, to_dept_cd, emp_no
	
	StrEbrFile = "ha107oa1"

	to_retire_dt = UniConvDateToYYYYMMDD(strToDt,parent.gDateFormat,parent.gServerDateType)
	fr_dept_cd = "0"
	stand_yy = frm1.txtStand_yy.Year
	fr_retire_dt = UniConvDateToYYYYMMDD(strFrDt,parent.gDateFormat,parent.gServerDateType)
	to_dept_cd = "ZZZZZZZZZZZZ"
	emp_no = frm1.txtEmp_no.value
	
	if emp_no = "" then
		emp_no = "%"
	End if 

	condvar = "end_dt|" & to_retire_dt
	condvar = condvar & "|fr_dept_cd|" & fr_dept_cd 
	condvar = condvar & "|retire_yy|" & stand_yy
	condvar = condvar & "|start_dt|" & fr_retire_dt
	condvar = condvar & "|to_dept_cd|" & to_dept_cd
	condvar = condvar & "|emp_no|" & emp_no

  	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 	call FncEBRPrint(EBAction , ObjName , condvar)

End Function

'========================================================================================
' Function Name : FncBtnPreview()
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview()
    Dim strFrDt
    Dim strToDt
	Dim ObjName    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If

    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if    
    strFrDt = frm1.txtFr_retire_dt.Text
    strToDt = frm1.txtTo_retire_dt.Text

    If Trim(strFrDt) = "" then
        strFrDt = UniConvYYYYMMDDToDate(parent.gDateFormat,"1950","01","01")
    End if
    If Trim(strToDt) = "" then
        strToDt = UniConvYYYYMMDDToDate(parent.gDateFormat,"2100","12","31")
    End If
	If CompareDateByFormat(Trim(strFrDt),Trim(strToDt),frm1.txtFr_retire_dt.Alt,frm1.txtTo_retire_dt.Alt,"800002",parent.gDateFormat,parent.gComDateType,True) = False Then
        frm1.txtFr_retire_dt.focus
        Set gActiveElement = document.activeElement
        Exit Function
    End if 
	
	dim condvar
	dim arrParam, arrField, arrHeader
    Dim StrEbrFile
		
	dim  to_retire_dt, fr_dept_cd,  stand_yy, fr_retire_dt, to_dept_cd, emp_no 
	
	StrEbrFile = "ha107oa1"

	to_retire_dt = UniConvDateToYYYYMMDD(strToDt,parent.gDateFormat,parent.gServerDateType)
	fr_dept_cd = "0"
	stand_yy = frm1.txtStand_yy.Year
	fr_retire_dt = UniConvDateToYYYYMMDD(strFrDt,parent.gDateFormat,parent.gServerDateType)
	to_dept_cd = "ZZZZZZZZZZZ"
	emp_no = frm1.txtEmp_no.value
	
	if emp_no = "" then
		emp_no = "%"
	End if 

	condvar = "end_dt|" & to_retire_dt
	condvar = condvar & "|fr_dept_cd|" & fr_dept_cd 
	condvar = condvar & "|retire_yy|" & stand_yy
	condvar = condvar & "|start_dt|" & fr_retire_dt
	condvar = condvar & "|to_dept_cd|" & to_dept_cd
	condvar = condvar & "|emp_no|" & emp_no
	
  	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	call FncEBRPreview(ObjName , condvar)

End Function

Function FncPrint() 
    Call parent.FncPrint()
End Function

Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE,False)
End Function

'========================================================================================================
'   Event Name : txtEmp_no_change             '<==인사마스터에 있는 사원인지 확인 
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    
    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	                
        IntRetCd = CommonQueryRs(" NAME "," HAA010T "," EMP_NO =  " & FilterVar(frm1.txtEmp_no.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If  IntRetCd = false then
            Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
            Exit Function      
        Else
            frm1.txtName.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if
    
End Function 

'======================================================================================================
'   Event Name : txtYyyymm_DblClick
'   Event Desc : 달력 Popup을 호출 
'======================================================================================================
Sub txtStand_yy_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtStand_yy.Action = 7
		frm1.txtStand_yy.focus
	End If
End Sub
'-------------------------------------------
Sub txtFr_retire_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtFr_retire_dt.Action = 7
		frm1.txtFr_retire_dt.focus
	End If
End Sub
'-------------------------------------------
Sub txtTo_retire_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M") 
		frm1.txtTo_retire_dt.Action = 7
		frm1.txtTo_retire_dt.focus
	End If
End Sub

'=======================================================================================================
'   Event Name : txtTo_retire_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtStand_yy_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtFr_retire_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtTo_retire_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>퇴직지급명세서출력</font></td>
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
								<TD CLASS="TD5" NOWRAP>기준년도</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha107oa1_txtStand_yy_txtStand_yy.js'></script></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>퇴직일</TD>
							    <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha107oa1_txtFr_retire_dt_txtFr_retire_dt.js'></script>&nbsp;~&nbsp;
							                           <script language =javascript src='./js/ha107oa1_txtTo_retire_dt_txtTo_retire_dt.js'></script></TD></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>대상자</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" ID = "txtEmp_no" NAME="txtEmp_no" SIZE=13 MAXLENGTH=13 tag="11XXXU" ALT="사번"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmptName (0)">
								                       <INPUT TYPE="Text" NAME="txtName" SIZE=20 MAXLENGTH=30 tag="14XXXU" ALT="성명"></TD>
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
		<TD HEIGHT=20><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
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


