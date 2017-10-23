<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/incServer.asp"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/CommStyleSheet.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

<%
EndDate		= GetSvrDate
%>

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    if  pOpt = "Q" then
        lgKeyStream    = Trim(parent.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
    else
        lgKeyStream    = Trim(frm1.txtEmp_no.Value) & gColSep
    end if
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029(gCurrency, "Q", "H") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
    parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Call LayerShowHide(0)
    Call LoadInfTB19029 
    Call SetToolBar("00000")
    Call LockField(Document)

    frm1.txtBas_dt.value = UniConvDateAToB("<%=EndDate%>",gServerDateFormat,gDateFormat)
    
    frm1.txtBasYymm.value = uniConvDateAtoB(frm1.txtBas_dt.value,gDateFormat, gDateFormatYYYYMM)
    
    frm1.txtemp_no1.value = parent.txtemp_no.value
    frm1.txtname1.value = parent.txtname.value
    frm1.txtBas_dt.focus() 
End Sub
'========================================================================================
' Function Name : Window_onUnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_unLoad()
End Sub

Function DbQuery()

    DbQuery = False                                                              '☜: Processing is NG
    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview() 
'On Error Resume Next                                                    '☜: Protect system from crashing
    
	Dim strUrl
	Dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim strDate
	Dim emp_no, presentdt, basyymm, presentnm, officenm, cntsum, reason
	Dim strYear,strMonth,strDay
	Dim strDate1
	dim strGetDate, strGetType
	Dim strCnt
	Dim i

	StrEbrFile = "h9115oa1.ebr"
	
	if  Date_chk(frm1.txtBas_dt.value, strDate) = True then
        frm1.txtBas_dt.value = strDate
    else
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtBas_dt.focus()
        exit function
    end if

	strDate1 = uniGetFirstDay(frm1.txtBasYymm.value,gDateFormatYYYYMM)	
	if instr(1, gDateFormatYYYYMM, gComDateType) <> instr(1, Trim(frm1.txtBasYYmm.value),gComDateType) then 
	    Call DisplayMsgBox("800094","X","X","X")
        frm1.txtBasYymm.focus()
        exit function
    end if
	if	Date_chk(strDate1, strDate) = False then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtBasYymm.focus()
        exit function
    end if

    if frm1.txtBasYymm.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtBasYymm.focus()
        exit function
    end if

    If Not chkFieldLength(Document, "1") Then									         '☜: This function check required field
		Exit Function
	end if    
    
	With frm1
	    emp_no    = .txtEmp_no1.value
	    presentdt = uniConvDateToYYYYMMDD(.bas_dt.value,gDateFormat,"")

        basyymm = uniGetLastDay(.txtBasYymm.value,gDateFormatYYYYMM)
        basyymm = UNIConvDateToYYYYMMDD(basyymm, gDateFormat, "")
        strYear = Mid(basyymm,1,4)
		strMonth = Mid(basyymm,5,2)

	    basyymm   = strYear & strMonth

	    presentnm = "%"
	    officenm  = "%"
	    cntsum    = 1
	    reason    = "%"
    End With
	
	strUrl = "Emp_no|" & emp_no
	strUrl = strUrl & "|PresentDt|" & presentdt
	strUrl = strUrl & "|Basyymm|" & basyymm
	strUrl = strUrl & "|PresentNm|" & presentnm
	strUrl = strUrl & "|OfficeNm|" & officenm
	strUrl = strUrl & "|CntSum|" & cntsum
	strUrl = strUrl & "|Reason|" & reason

	call FncEBRPreview(StrEbrFile , strUrl)

End Function
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
FncBtnPreview()
End Function


Function DbQueryFail()
    Err.Clear
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

Sub Query_OnClick()
    Call DbQuery()
End Sub

Sub Print_onClick()
    Call SubPrint(MyBizASP)
End Sub


Sub GRID_PAGE_OnChange()
End Sub

Sub DELETE_OnClick()
    Call Grid1.DeleteClick()
End Sub

Sub CANCEL_OnClick()
    Call Grid1.CancelClick()
End Sub

Sub FncPrintPrev()

    Call DbQuery()

End Sub

Function Document_onKeyDown()
Dim CuEvObj,KeyCode
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode
	Select Case KeyCode
		Case 13 'enter key
			Call FncBtnPreview()
	End Select		
	Document_onKeyDown	= True	
End Function


</SCRIPT>
<!-- #Include file="../../inc/uniSimsClassID.inc" --> 
<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 width=770 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 border=0 bgcolor=#ffffff width=743>
                    <TR height=15 valign=middle>
                        <TD class=base1><INPUT type=hidden NAME="txtEmp_no" MAXLENGTH=13 SiZE=12  tag=14></TD>
                        <TD class=base1><INPUT type=hidden NAME="txtName" MAXLENGTH=20 SiZE=10  tag=14></TD>
                        <TD class=base1><INPUT type=hidden NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10  tag=14></TD>
                        <TD class=base1><INPUT type=hidden NAME="txtDept_nm" MAXLENGTH=25 SiZE=15  tag=14></TD>
                    </TR>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#ffffff>
                                <TR>
	                        		<TD CLASS=TDFAMILY_TITLE NOWRAP>발급일</TD>
	                        		<TD CLASS=TDFAMILY>
	                        		    <INPUT CLASS="SINPUTTEST_STYLE" id=bas_dt NAME="txtBas_dt" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="22D" ondblclick="VBScript:Call OpenCalendar('txtBas_dt',3)">
	                        		</TD>
                                </TR>
                                <TR>
	                        		<TD CLASS=TDFAMILY_TITLE NOWRAP>기준년월</TD>
	                        		<TD CLASS=TDFAMILY>
	                        		    <INPUT CLASS="SINPUTTEST_STYLE" id=BasYymm NAME="txtBasYymm" TYPE="Text" MAXLENGTH=10 SiZE=7 tag="22M" ondblclick="VBScript:Call OpenCalendar('txtBasYymm',2)">
	                        		</TD>
                                </TR>
                                <TR>
	                        		<TD CLASS=TDFAMILY_TITLE NOWRAP>대상자</TD>
	                        		<TD CLASS=TDFAMILY align=left>
	                        		    <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtEmp_no1" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="24">
                                        <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtName1" TYPE="Text" MAXLENGTH=15 SiZE=15 tag="24">
	                        		</TD>
                                </TR>
                                <TR valign=middle height=50>
                                    <TD colspan=2 align=center>
	                        			<INPUT style="WIDTH: 150px; HEIGHT: 20px" TYPE=button NAME=printprev VALUE="미리보기/출력" OnClick="Vbscript: call FncBtnPreview()">
                                    </TD>
                                </TR>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
            <TD width=14></TD>
        </TR>
    </TABLE>
    <TABLE cellSpacing=0 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=HIDDEN NAME="txtMode">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext">
    <INPUT TYPE=HIDDEN NAME="txtres_no">
    <INPUT TYPE=HIDDEN NAME="txtdomi">
    <INPUT TYPE=HIDDEN NAME="txtaddr">
    <INPUT TYPE=HIDDEN NAME="txtentr_dt">
    <INPUT TYPE=HIDDEN NAME="txtretire_dt">
    <INPUT TYPE=HIDDEN NAME="txtrepre_nm">
    <INPUT TYPE=HIDDEN NAME="txtco_full_nm">
    <INPUT TYPE=HIDDEN NAME="txtissueno">
</FORM>	
<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM>

</BODY>
</HTML>
