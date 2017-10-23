<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>

<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!-- #Include file="../ESSinc/incServer.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<Script Language="VBScript">
Option Explicit  

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

<% EndDate		= GetSvrDate %>

'========================================================================================================
' Name : LoadInfTB19029()	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(gCurrency, "Q", "H") %>
End Sub

'========================================================================================================
' Name : Form_Load
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
' Function Name : Form_unLoad
'========================================================================================
Sub Form_unLoad()
End Sub

'========================================================================================
' Function Name : FncBtnPreview
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

</SCRIPT>
<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 border=0 width=732>
        <TR height=15 valign=middle>
            <TD><INPUT type=hidden NAME="txtEmp_no" MAXLENGTH=13 SiZE=12  tag=14></TD>
            <TD><INPUT type=hidden NAME="txtName" MAXLENGTH=20 SiZE=10  tag=14></TD>
            <TD><INPUT type=hidden NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10  tag=14></TD>
            <TD><INPUT type=hidden NAME="txtDept_nm" MAXLENGTH=25 SiZE=15  tag=14></TD>
        </TR>
        <TR>
            <TD valign="top">
                <TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
                    <TR>
                        <TD>
                            <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#DDDDDD>
						    <TR>
								<TD CLASS=ctrow01 NOWRAP>발급일</TD>
								<TD CLASS=ctrow06>
								    <INPUT CLASS="form01" id=bas_dt NAME="txtBas_dt" TYPE="Text" MAXLENGTH=10 SiZE=10 ondblclick="VBScript:Call OpenCalendar('txtBas_dt',3)">
								</TD>
						    </TR>
						    <TR>
								<TD CLASS=ctrow01 NOWRAP>기준년월</TD>
								<TD CLASS=ctrow06>
								    <INPUT CLASS="form01" id=BasYymm NAME="txtBasYymm" TYPE="Text" MAXLENGTH=10 SiZE=7 ondblclick="VBScript:Call OpenCalendar('txtBasYymm',2)">
								</TD>
						    </TR>
						    <TR>
								<TD CLASS=ctrow01 NOWRAP>대상자</TD>
								<TD CLASS=ctrow06 align=left>
								    <INPUT CLASS="form02" NAME="txtEmp_no1" TYPE="Text" MAXLENGTH=13 SiZE=13 readonly>
						            <INPUT CLASS="form02" NAME="txtName1" TYPE="Text" MAXLENGTH=15 SiZE=15 readonly>
								</TD>
						    </TR>
							</TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
        </TR>
        <TR valign=middle height=50>
            <TD colspan=2 align=center>
	    		<IMG SRC="../ESSimage/button_04.gif" NAME=printprev VALUE="미리보기/출력" OnClick="Vbscript: call FncBtnPreview()" onMouseOver="javascript:this.src='../ESSimage/button_r_04.gif';" onMouseOut="javascript:this.src='../ESSimage/button_04.gif';">
            </TD>
        </TR>
    </TABLE>
    <TABLE cellSpacing=0 cellPadding=0 border=0>
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
