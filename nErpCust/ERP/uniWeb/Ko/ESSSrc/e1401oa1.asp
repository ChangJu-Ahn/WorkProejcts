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

Const BIZ_PGM_ID      = "e1401ob1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "e1401ma2.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

Dim StartDate

StartDate = "<%=GetSvrDate%>"

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
    if  pOpt = "Q" then
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
    end if
End Sub        

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

	Call LoadInfTB19029()
    Call LayerShowHide(0)

    Call SetToolBar("00000")

    frm1.txtProv_dt.value = UniConvDateAToB(StartDate,gServerDateFormat,gDateFormat)

    frm1.txtemp_no1.value = parent.txtemp_no.value
    frm1.txtname1.value = parent.txtname.value
    Call LockField(Document)
End Sub

'========================================================================================
' Function Name : Form_unLoad
'========================================================================================
Sub Form_unLoad()
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery()

    Dim strVal

    Err.Clear                                                                    '☜: Clear err status


    if frm1.txtUse.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtUse.focus()
        exit function
    end if

    DbQuery = False                                                              '☜: Processing is NG
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : FncBtnPreview
'========================================================================================
Function FncBtnPreview() 
    
	Dim strUrl
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	Dim emp_no, use, title, prov_dt, emp_name ,print_emp_no ,etc,remark
	dim strDate

	title = "1"
	StrEbrFile = "h2016oa1_ko441.ebr"
	emp_name = frm1.txtname.value
	prov_dt = uniConvDateToYYYYMMDD(frm1.txtProv_dt.value,gDateFormat,"")
	emp_no = frm1.txtEmp_no.value
	use = frm1.txtUse.value
	etc = frm1.remark.value
	
	if emp_no = "" then
		emp_no = "%"
		frm1.txtName.value = ""
	End if
 	
    if  Date_chk(frm1.txtProv_dt.value, strDate) = True then
        frm1.txtProv_dt.value = strDate
    else
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtProv_dt.focus()
        exit function
    end if    

	if len(frm1.remark.value) - len(replace(frm1.remark.value,vbCR,"")) > 4 then
		Call DisplayMsgbox("800601","X","5","X")	' 비고는 5줄까지 입력가능합니다	
		exit function
	else 'less than 7 line 
		 if gf_LenAtDb(frm1.remark.value)>300  then
			Call DisplayMsgBox("900028","X","비고","X") 
	        exit function  			
		 end if
	end if 
	if frm1.txtUse.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtUse.focus()
        exit function
    end if	    

    If Not chkFieldLength(Document, "1") Then									         '☜: This function check required field
		Exit Function
	end if    
	
	strUrl = "title|" & title
	strUrl = strUrl & "|emp_name|" & emp_name
	strUrl = strUrl & "|prov_dt|" & prov_dt
	strUrl = strUrl & "|emp_no|" & emp_no 
	strUrl = strUrl & "|print_emp_no|" & emp_no
	strUrl = strUrl & "|use|" & use
	strUrl = strUrl & "|remark|" & remark
	strUrl = strUrl & "|etc|" & etc

	call FncEBRPreview(StrEbrFile , strUrl)
End Function

'========================================================================================================
' Name : DbQueryOk
'========================================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
FncBtnPreview()
End Function

'========================================================================================================
' Name : DbQueryFail
'========================================================================================================
Function DbQueryFail()
    Err.Clear
End Function

'========================================================================================================
' Function Name : gf_LenAtDb
'========================================================================================================
Function gf_LenAtDb(szAllText) 
        Dim nLen 
        Dim nCnt 
        Dim szEach 

        nLen = 0 
        szAllText = Trim(szAllText) 
        For nCnt = 1 To Len(szAllText) 

                szEach = Mid(szAllText,nCnt,1) 
                If 0 <= Asc(szEach) And Asc(szEach) <= 255 Then 
                        nLen = nLen + 1             '한글이 아닌 경우 
                Else 
                        nLen = nLen + 2             '한글인 경우 
                End If 
        Next 

        gf_LenAtDb = nLen 
End Function

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Sub Query_OnClick()
    Call DbQuery()
End Sub

Sub FncPrintPrev()
    Call DbQuery()
End Sub

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
	                        		<TD CLASS=ctrow01 NOWRAP>기준일</TD>
	                        		<TD CLASS=ctrow06>
	                        		    <INPUT CLASS="form01" ID="txtProv_dt" NAME="txtProv_dt" TYPE="Text" MAXLENGTH=10 SiZE=10 ondblclick="VBScript:Call OpenCalendar('txtProv_dt',3)">
										
	                        		</TD>
                                </TR>
                                <TR>
	                        		<TD CLASS=ctrow01 NOWRAP>용도</TD>
	                        		<TD CLASS=ctrow06 align=left>
	                        		    <INPUT CLASS="form01" NAME="txtUse" TYPE="Text" MAXLENGTH=50 SiZE=50>
	                        		</TD>
                                </TR>
                               <TR>
	                        		<TD CLASS=ctrow01 NOWRAP>제출처</TD>
	                        		<TD CLASS=ctrow06 align=left>
									<!--<TEXTAREA rows=5 cols=80  ID="remark" NAME="remark" size=200 wrap=HARD ></TEXTAREA>-->
	                        		    <INPUT CLASS="form01" NAME="remark" TYPE="remark" MAXLENGTH=50 SiZE=50>
	                        		</TD>
                                </TR>                                
                                <TR>
	                        		<TD CLASS=ctrow01 NOWRAP>대상자</TD>
	                        		<TD CLASS=ctrow06 align=left>
	                        		    <INPUT CLASS="form02" NAME="txtEmp_no1" TYPE="Text" MAXLENGTH=13 SiZE=13 readonly>
                                        <INPUT CLASS="form02" NAME="txtName1" TYPE="Text" MAXLENGTH=15 SiZE=15 readonly>
	                        		</TD>
                                </TR>
                                <TR>
	                        		<TD CLASS=ctrow01 NOWRAP>&nbsp;</TD>
	                        		<TD CLASS=ctrow06 align=left><font face=times size=2 color=blue>※ 인사그룹 인장 날인 필요</font>
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
	    		<IMG SRC="../ESSimage/button_04.gif" NAME=printprev VALUE="미리보기/출력" OnClick="Vbscript: call FncPrintPrev()" onMouseOver="javascript:this.src='../ESSimage/button_r_04.gif';" onMouseOut="javascript:this.src='../ESSimage/button_04.gif';">
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
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   

