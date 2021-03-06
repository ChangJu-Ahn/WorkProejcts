<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Buffer = True
Response.Expires = -1%>

<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!-- #Include file="../ESSinc/IncServer.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<Script Language="VBScript">
Option Explicit 

Const BIZ_PGM_ID      = "e1109mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXCOLS  = 15
Const C_SHEETMAXROWS  = 10

<!-- #Include file="../ESSinc/lgvariables.inc" --> 
<!-- #Include file="../ESSinc/incGrid.inc" -->

Dim Grid1

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################
Sub LoadInfTB19029()
	<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029(gCurrency, "Q", "H") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    if  pOpt = "Q" then
        if  Trim(parent.txtEmp_no2.Value) = "" then
            lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
        else
            lgKeyStream = Trim(parent.txtEmp_no2.Value) & gColSep
        end if
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtdept_nm.value) & gColSep        
        lgKeyStream = lgKeyStream & "" & gColSep
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtdept_nm.value) & gColSep        
    end if
End Sub        

'========================================================================================================
' Function Name : InitSpreadSheet
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = C_SHEETMAXCOLS
    Grid1.SheetMaxrows = C_SHEETMAXROWS
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()
	on Error Resume Next
    Err.Clear                                                                       '☜: Clear err status

    if  parent.txtDEPT_AUTH.value = "Y" then
        parent.document.All("nextprev").style.VISIBILITY = "visible"
        Call SetToolBar("10000")    
    else
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
        Call SetToolBar("00000")    
    end if

    Call LayerShowHide(0)

	Call LoadInfTB19029()
    
    Call InitGrid()
    Call LockField(Document)

    Call DbQuery(1)

End Sub

'========================================================================================
' Function Name : Form_UnLoad
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
 	Set gActiveElement = Nothing
    Set Grid1 = Nothing
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery(ppage)

    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
     Call Grid1.ShowData(frm1,frm1.grid_page.value)
End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
	Call Grid1.Clear(frm1,frm1.grid_page.value)
End Function

'========================================================================================================
' Name : DbSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	Call LayerShowHide(1)

	With Frm1
		.txtMode.value        = "UID_M0002"                                        '☜: Save
'		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
    DbSave  = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
'========================================================================================================
Function DbSaveOk()
    Call DbQuery(1)
End Function

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

Sub Query_OnClick()
    Call DbQuery(1)
End Sub

Sub GRID_PAGE_OnChange()
End Sub

Function txtEmp_no2_Onchange()
        Call DbQuery(1)	
End Function
</SCRIPT>

<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 border=0>
        <TR>
            <TD valign="top">
                <TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
                    <TR>
                       <td height="5"></td>
                    </TR>
                    <TR>
                        <td><table border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD" width=733>
                            <tr> 
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">사번</td>
								<td width="85" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">성명</td>
								<td width="86" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtName" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">직위</td>
								<td width="100" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">부서</td>
								<td width="153" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtDept_nm" MAXLENGTH=25 SiZE=22  readonly></td>
                            </tr>
                            </table>
                        </td>
                    </TR>
                    <TR>
                       <td height="5"></td>
                    </TR>
                    <TR>
                        <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
								<TR> 
								    <TD class=TDFAMILY_TITLE1></TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>교육명</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1 colspan=2>교육기간</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>기관</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>국가</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>교육코드</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>구분</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>점수</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>의무교육기간</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>교육비</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>정산</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>고용보험환급</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>레포터</TD>
	                        		<TD CLASS=TDFAMILY_TITLE1>고과</TD>
                                </TR>
							    <% 
                                For i=1 To 10
                                    Response.Write "<TR bgcolor=#F8F8F8 height=24 onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
                                    Response.Write "<TD><INPUT name='" & i & "'  class=listrow01 flag='SPREADCELL' style='WIDTH:  30px; text-align: center;'  readonly></TD>"
                                    Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH: 120px; text-align: left;'  readonly></TD>"
                                    Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  75px; text-align: center;' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  75px; text-align: center;' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH: 110px; text-align: left;' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  70px; text-align: left;' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH: 100px; text-align: left;' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  50px; text-align: center;' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  50px; text-align: right;' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  100px; text-align: right;' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  70px; text-align: right;' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  50px; text-align: right;' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  100px; text-align: right;' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  60px; text-align: center;' readonly></TD>"
                                	Response.Write "<TD><INPUT name='SPREADCELL' class=listrow01 flag='SPREADCELL' style='WIDTH:  50px; text-align: right;' readonly></TD>"
                                    Response.Write "</TR>"
                                Next
								%>
                           </table>
                        </td>
                    </TR>
                </TABLE>
            </TD>
        </TR>
        <TR>
            <TD height=5></TD>
        </TR>
        <TR height=13>
            <TD VALIGN=center ALIGN=center>
                <A onclick="VBSCRIPT:CALL GRID1.PREPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="이전페이지" SRC=../ESSimage/button_07.gif border=0 ></A>&nbsp;
                <A onclick="VBSCRIPT: CALL GRID1.NEXTPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="다음페이지" SRC=../ESSimage/button_08.gif border=0 ></A>&nbsp;&nbsp;
            </TD>
        </TR>
    </TABLE>
    <TABLE cellSpacing=0 cellPadding=0 border=0>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=auto noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=hidden NAME="txtMode">
    <INPUT TYPE=hidden NAME="txtKeyStream">
    <INPUT TYPE=hidden NAME="txtUpdtUserId">
    <INPUT TYPE=hidden NAME="txtInsrtUserId">
    <INPUT TYPE=hidden NAME="txtFlgMode">
    <INPUT TYPE=hidden NAME="txtPrevNext">
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES>
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1>
 </FORM>	
</BODY>
</HTML>
