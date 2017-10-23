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
<Script Language="VBScript">
Option Explicit    


Const BIZ_PGM_ID      = "e1806mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1      = "e1805ma1.asp"

Const C_SHEETMAXCOLS = 7

<!-- #Include file="../ESSinc/lgvariables.inc" --> 
<!-- #Include file="../ESSinc/incGrid.inc" -->

Dim Grid1

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
        lgKeyStream       = ""
        If frm1.txtuse_yn1.checked = true then
            lgKeyStream = lgKeyStream & "emp_no" & gColSep
        else
            lgKeyStream = lgKeyStream & "dept_cd" & gColSep
        end if
End Sub 
       
'========================================================================================================
' Function Name : InitSpreadSheet
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = C_SHEETMAXCOLS
    Grid1.SheetMaxrows = 10
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

    parent.document.All("nextprev").style.VISIBILITY = "hidden"

    Call LayerShowHide(0)
    Call InitGrid()

    Call SetToolBar("10000")
    Call LockField(Document)
    frm1.txtuse_yn1.click()
    Call DbQuery(1)
End Sub

'========================================================================================
' Function Name : Form_UnLoad
'========================================================================================
Sub Form_UnLoad()
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

    frm1.GRID_PAGE.VALUE = ppage

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
' Function Name : GetRow
'========================================================================================
Function GetRow(pRow)
	GetRow=False
    Grid1.ActiveRow = pRow
    If Mid(document.activeElement.getAttribute("tag"),3,1) = "1" Then
	    arrRet = window.showModalDialog("../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	GetRow=True
End Function

'========================================================================================
' Function Name : DoubleGetRow
'========================================================================================
Function DoubleGetRow(pRow)
    Dim objList
    Dim elmCnt

    Dim emp_no,dept_cd
    Dim strVal

	DoubleGetRow = False
	Grid1.ActiveRow = pRow

    emp_no = ""
    with frm1
    	For elmCnt = 0 to .length - 1
    		Set objList = .elements(elmCnt)
    		If objList.name = "SPREADCELL_emp_no" & pRow then
               emp_no  = objList.value
            ElseIf objList.name = "SPREADCELL_deptcd" & pRow then
               dept_cd = objList.value
    		End if
    	Next
    End With
    If  emp_no <> "" then
        strVal = BIZ_PGM_ID1 & "?emp_no=" & emp_no & "&dept_cd=" & dept_cd
        Call CommonQueryRs(" MENU_NAME "," E11000T "," Menu_id = " & FilterVar("E1805MA1", "''", "S") & " AND LANG_CD =  " & FilterVar(gLang , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		parent.txtTitle.value = replace(lgF0,chr(11),"")
        document.location = strVal

    end if

	DoubleGetRow = True
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
    Call DbQuery()
End Sub

Sub GRID_PAGE_OnChange()
End Sub

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
								<TD width="100" height="30" bgcolor="D4E5E8" class=base1>정렬순서
								</TD>
								<TD width="636" height="30" bgcolor="FFFFFF" class="base2" align=left valign=center colspan=3>&nbsp;&nbsp;
    								<INPUT TYPE="RADIO" NAME="txtuse_yn" CLASS="radio_title" ID="txtuse_yn1" VALUE="emp_no" ><LABEL FOR="txtuse_yn1">사번별</LABEL>
									<INPUT TYPE="RADIO" NAME="txtuse_yn" CLASS="radio_title" ID="txtuse_yn2" VALUE="dept_cd" ><LABEL FOR="txtuse_yn2">부서별</LABEL>
								</TD>
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
		                    	<TD class=TDFAMILY_TITLE1>사번</TD>
		                    	<TD class=TDFAMILY_TITLE1>성명</TD>
		                        <TD class=TDFAMILY_TITLE1>담당부서</TD>
		                        <TD class=TDFAMILY_TITLE1>부서명</TD>
		                    	<TD class=TDFAMILY_TITLE1>등록일</TD>		                            
		                    	<TD class=TDFAMILY_TITLE1>하위부서권한</TD>
                            </TR>
							<% 
							For i=1 To 10
							     Response.Write "<TR bgcolor=#F8F8F8 height=24 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
							     Response.Write "<TD><INPUT class=listrow01 name='" & i & "'  flag='SPREADCELL' style='WIDTH: 30px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_emp_no" & i & "' flag='SPREADCELL' style='WIDTH: 100px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_NAME' flag='SPREADCELL' style='WIDTH: 120px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_deptcd" & i & "' flag='SPREADCELL' style='WIDTH:  110px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_deptnm" & i & "' flag='SPREADCELL' style='WIDTH:  190px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_date" & i & "' flag='SPREADCELL' style='WIDTH: 80px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_internal_auth' flag='SPREADCELL' style='WIDTH: 95px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
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
    <TABLE cellSpacing=0 cellPadding=0 border=0>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=auto noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=hidden NAME="txtMode">
    <INPUT TYPE=hidden NAME="txtKeyStream">
    <INPUT TYPE=hidden NAME="txtUpdtUserId">
    <INPUT TYPE=hidden NAME="txtInsrtUserId">
    <INPUT TYPE=hidden NAME="txtFlgMode">
    <INPUT TYPE=hidden NAME="txtPrevNext">
    <INPUT TYPE=hidden NAME="txtMaxRows"> 
    <TEXTAREA style="display: none" name=txtSpread></TEXTAREA>
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES>
    <INPUT TYPE=hidden NAME=GRID_PAGE value=1>
 </FORM>	
</BODY>
</HTML>
