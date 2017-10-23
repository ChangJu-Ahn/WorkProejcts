<%@ LANGUAGE="VBSCRIPT" %>

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

'========================================================================================================
' Name : LoadInfTB19029()	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(gCurrency, "Q", "H") %>
End Sub

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
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    Dim lgYear,i,stYear
    Dim strWhere
	  
	Call CommonQueryRs(" close_type,year(close_dt) close_year "," hda270t "," org_cd=1 and pay_gubun=" & FilterVar("Z", "''", "S") & " and pay_type=" & FilterVar("*", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	 
	
	if lgF0="1" then
		lgYear = cint(Replace(lgF1, Chr(11), ""))-1
	else
		lgYear = cint(Replace(lgF1, Chr(11), ""))
	end if 
	
	if Trim(parent.txtemp_no.value)="unierp" then
		stYear=lgYear-1
	else
		Call CommonQueryRs("entr_dt "," haa010t ","emp_no =  " & FilterVar(parent.txtemp_no.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if lgF0="" then
			stYear = "1990"
		else
			stYear =Year(lgF0)
		end if
	end if		
	
    For i = lgYear To cint(stYear) step -1
   		Call SetCombo(frm1.txtYear, i, i)
   	Next

    frm1.txtYear.value = CStr(lgYear)
    '신고 사업장    
    
    strWhere = " YEAR_AREA_CD = (SELECT YEAR_AREA_CD FROM HAA010T WHERE emp_no =  " & FilterVar(parent.txtemp_no.value, "''", "S") & ")"
    Call CommonQueryRs(" YEAR_AREA_CD, YEAR_AREA_NM "," HFA100T ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = lgF0
    iNameArr = lgF1
    
    Call SetCombo2(frm1.txtcust_cd,iCodeArr,iNameArr,Chr(11))     

End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
    Call InitComboBox()

    parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Call LayerShowHide(0)
	Call LoadInfTB19029()
    Call SetToolBar("00000")
    Call LockField(Document)

    frm1.txtemp_no1.value = parent.txtemp_no.value
    frm1.txtname1.value = parent.txtname.value
    frm1.txtYear.focus() 
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
	Dim strDate, Fr_Dept_cd, To_Dept_cd
	Dim emp_no,bas_dt,bas_yy, biz_area_cd,ocpt_type,prov_dt,tax_nm,prt_check_flag

	with frm1
	    bas_dt =.txtYear.value & "1231"
	    bas_yy = .txtYear.value
	    biz_area_cd = .txtcust_cd.value 
	    emp_no = .txtEmp_no1.value 
	    ocpt_type = "%" 
	    prov_dt = .txtYear.value & "1231"
	    prt_check_flag = "2"
	end with

    if frm1.txtYear.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtYear.focus()
        exit function
    end if
	if frm1.txtcust_cd.value = "" then
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtcust_cd.focus()
        exit function
    end if

    If Not chkFieldLength(Document, "1") Then									         '☜: This function check required field
		Exit Function
	end if    
	
	strUrl = "bas_dt|" & bas_dt
	strUrl = strUrl & "|bas_yy|" & bas_yy 
	strUrl = strUrl & "|biz_area_cd|" & biz_area_cd
	strUrl = strUrl & "|emp_no|" & emp_no
	strUrl = strUrl & "|fr_dept_cd|0"
	strUrl = strUrl & "|ocpt_type|" & ocpt_type
	strUrl = strUrl & "|prov_dt|" & prov_dt
	strUrl = strUrl & "|prt_check_flag|" & prt_check_flag	
	strUrl = strUrl & "|to_dept_cd|ZZZZ"
	
	StrEbrFile = "h9114oa1_12006.ebr"

	call FncEBRPreview(StrEbrFile , strUrl)
	
	StrEbrFile = "h9114oa1_12006p2.ebr"

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
                <TABLE width="100%" cellSpacing=1 cellPadding=0 border=0 bgcolor=#DDDDDD>
                    <TR>
	                	<TD CLASS=ctrow01 NOWRAP>기준년</TD>
	                	<TD CLASS=ctrow06>
					        <SELECT class="form01" Name="txtYear" tabindex=-1 STYLE="WIDTH: 100px">
					        </SELECT>
	                	</TD>
                    </TR>
                    <TR>
	                	<TD CLASS=ctrow01 NOWRAP>신고사업장</TD>
	                	<TD CLASS=ctrow06 align=left>
                            <Select class="form02" NAME="txtcust_cd" ALT="신고사업장" STYLE=" WIDTH: 200px" disabled></SELECT>
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
        <TR valign=middle height=50>
            <TD align=center>
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
</BODY>
</HTML>
