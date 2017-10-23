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
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<% EndDate   = GetSvrDate 

	Dim dilig_cd,day,year

    day		= Request("day")	
    dilig_cd= Request("dilig_cd")
    year    = Request("year")

%>
<Script Language="VBScript">
Option Explicit

Const C_SHEETMAXCOLS = 5
Const C_SHEETMAXROWS = 9
Const BIZ_PGM_ID      = "e1608mb2.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 
<!-- #Include file="../ESSinc/incGrid.inc" -->

Dim Grid1
dim fDiligAuth,fAuthCheck
Dim dilig_cd,day

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
    else     
            lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep   
    end if 	
        lgKeyStream = lgKeyStream & Trim(parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(day) & gColSep
        lgKeyStream = lgKeyStream & Trim(dilig_cd) & gColSep
        lgKeyStream = lgKeyStream & Trim(fDiligAuth) & gColSep        
        lgKeyStream = lgKeyStream & Trim(fAuthCheck) & gColSep     
        
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
    On Error Resume Next
    Err.Clear                                                                       '☜: Clear err status

    call FncGetDiligAuth(fDiligAuth,fAuthCheck)
    parent.document.All("nextprev").style.VISIBILITY = "hidden"
    
    Call LayerShowHide(0)
    Call InitGrid()

    Call SetToolBar("00000")
    Call LockField(Document)

    Call DbQuery(1)
End Sub

'========================================================================================
' Function Name : Form_unLoad
'========================================================================================
Sub Form_unLoad()
	On Error Resume Next
    Set Grid1 = Nothing
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery(ppage)

    Dim strDate
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

	day= "<%=day%>"
	dilig_cd = <%=dilig_cd%>
	
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

    Call Grid1.ShowData(frm1,frm1.GRID_PAGE.VALUE)

End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status

End Function

'========================================================================================
' Function Name : MouseRow
'========================================================================================
Sub MouseRow(pRow)
	If frm1.grid_totpages.value = "" Then Exit Sub
    Dim objList   

	Grid1.ActiveRow = pRow	
	Set objList = window.event.srcElement	
	
	If  UCase(objList.getAttribute("flag")) = "SPREADCELL" then
        if objList.value = "" then            
             objList.style.cursor = "auto"
        else
             objList.style.cursor = "hand"
        end if
    End If        

End Sub

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

'========================================================================================================
' Name : FncGetDiligAuth()
'========================================================================================================
Function FncGetDiligAuth(fDiligAuth,fAuthCheck)
    fDiligAuth = ""
    fAuthCheck = ""
    Call CommonQueryRs(" internal_cd,internal_auth "," e11090t "," emp_no =  " & FilterVar(parent.txtEmp_no.Value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    fDiligAuth = replace(lgF0,chr(11),chr(12))
    fDiligAuth = replace(fDiligAuth," ","")    
    fAuthCheck = replace(lgF1,chr(11),chr(12))
    fAuthCheck = replace(fAuthCheck," ","")      
End Function

</SCRIPT>
<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

</HEAD>
<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="post">
    <TABLE cellSpacing=0 cellPadding=0 border=0>
        <TR>
            <TD valign="top">
                <TABLE width="100%" cellSpacing=0 cellPadding=0 border=0>
                    <TR>
                       <td height="10"></td>
                    </TR>
                    <TR>
                        <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
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
		                        	<TD class=TDFAMILY_TITLE1>근태명</TD>
		                        	<TD class=TDFAMILY_TITLE1>근태일자</TD>
		                        	<TD class=TDFAMILY_TITLE1>입력자</TD>
									<TD class=TDFAMILY_TITLE1>사번</TD>		                        	
                                </TR>
								<%
						        For i=1 To 9
						            Response.Write "<TR bgcolor=#E9EDF9 height=20 onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
						            Response.Write "<TD><INPUT class=listrow01 name='" & i & "'  flag='SPREADCELL' style='WIDTH:  50px; TEXT-ALIGN: center' readonly></TD>"
						            Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL' flag='SPREADCELL' style='WIDTH: 226px; TEXT-ALIGN: left' readonly></TD>"
						        	Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL' flag='SPREADCELL' style='WIDTH: 130px; TEXT-ALIGN: center' readonly></TD>"
						        	Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL' flag='SPREADCELL' style='WIDTH: 160px; TEXT-ALIGN: center' readonly></TD>"
						        	Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL' flag='SPREADCELL' style='WIDTH: 160px; TEXT-ALIGN: center' readonly></TD>"        	
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
        <TR height=20>
            <TD VALIGN=center ALIGN=center>
                <A onclick="VBSCRIPT:CALL GRID1.PREPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="이전페이지" src=../ESSimage/button_07.gif border=0 ></A>&nbsp;
                <A onclick="VBSCRIPT: CALL GRID1.NEXTPAGES()" onMouseOver="javascript: this.style.cursor='hand'"><IMG alt="다음페이지" src=../ESSimage/button_08.gif border=0 ></A>&nbsp;&nbsp;
            </TD>
        </TR>
    </TABLE>
    <TABLE border=0  align = center>
        <TR align = center>
			<td>
    			<INPUT TYPE=image NAME=printprev SRC="../ESSimage/button_15.gif" VALUE="돌아가기" OnClick="javascript:history.back()" alt='돌아가기' onMouseOver="javascript:this.src='../ESSimage/button_r_15.gif';" onMouseOut="javascript:this.src='../ESSimage/button_15.gif';">
    		</td>
    	</tr>
    </table>
            
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
    
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES STYLE="WIDTH: 50px; HEIGHT: 20px">
    <INPUT TYPE=hidden NAME=GRID_PAGE  value=1 >
 </FORM>	
</BODY>
</HTML>
