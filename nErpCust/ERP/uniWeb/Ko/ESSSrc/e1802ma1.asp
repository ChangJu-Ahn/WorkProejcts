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

Const BIZ_PGM_ID      = "e1802mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1      = "e1801ma1.asp"
Const C_SHEETMAXCOLS = 9

<!-- #Include file="../ESSinc/lgvariables.inc" --> 
<!-- #Include file="../ESSinc/incGrid.inc" -->

Dim Grid1
Dim IsOpenPop
'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)

    if  pOpt = "Q" then
        lgKeyStream       = Trim(parent.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
        if frm1.txtuse_yn0.checked = true then
            lgKeyStream = lgKeyStream & "" & gColSep
        elseif frm1.txtuse_yn1.checked = true then
            lgKeyStream = lgKeyStream & "Y" & gColSep
        else
            lgKeyStream = lgKeyStream & "N" & gColSep
        end if
        
        lgKeyStream = lgKeyStream & frm1.txtEmp_no1.Value & gColSep        
    else
        lgKeyStream       = Trim(frm1.txtEmp_no.Value) & gColSep
    end if

End Sub        

'========================================================================================================
' Function Name : InitSpreadSheet
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = C_SHEETMAXCOLS
    Grid1.SheetMaxrows = 8
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

    parent.document.All("nextprev").style.VISIBILITY = "hidden"

    Call LayerShowHide(0)
    Call InitGrid()    
    Call SetToolBar("10000")
    Call LockField(Document)

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
    
    If txtEmp_no1_Onchange() Then  
        Exit Function
    End if

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
	Call Grid1.Clear(frm1,frm1.grid_page.value)    
    Call ClearField(Document,2)                                                                  '☜: Clear err status
End Function


'========================================================================================
' Function Name : DoubleGetRow
'========================================================================================
Function DoubleGetRow(pRow)
    Dim objList
    Dim elmCnt

    Dim emp_no
    Dim strVal

	DoubleGetRow = False
	Grid1.ActiveRow = pRow

    emp_no = ""
    with frm1
    	For elmCnt = 0 to .length - 1
    		Set objList = .elements(elmCnt)
    		If objList.name = "SPREADCELL_emp_no" & pRow then
               emp_no = objList.value
    		End if
    	Next
    End With

    If  emp_no <> "" then
        strVal = BIZ_PGM_ID1 & "?emp_no=" & emp_no
		strVal = strVal& "&updateok=ok"
		Call CommonQueryRs(" MENU_NAME "," E11000T "," Menu_id = " & FilterVar("E1801MA1", "''", "S") & " AND LANG_CD =  " & FilterVar(gLang , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
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
' Name : OpenEmp()
'========================================================================================================
Function OpenEmp(pEmpNo)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True   or lgIntFlgMode = OPMD_UMODE  Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtEmp_no1.value)			' Code Condition
	arrParam(1) = ""'frm1.txtname1.value			' Name Cindition
    arrParam(2) = Trim(parent.txtinternal_cd.Value)'lgUsrIntCd
	
	arrRet = window.showModalDialog("E1EmpPopa4.asp", Array(arrParam), _
		"dialogWidth=546px; dialogHeight=387px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
	    frm1.txtEmp_no1.value = Trim(arrRet(0))
	    frm1.txtname1.value = Trim(arrRet(1))
	End If	
End Function

'========================================================================================================
' Name : txtEmp_no1_Onchange()
'========================================================================================================
Function txtEmp_no1_Onchange()
    On Error Resume Next
    Err.Clear
    
    Dim iDx
    Dim IntRetCd
    Dim strEmp_no

    IF Trim(frm1.txtEmp_no1.value) = "" THEN
        frm1.txtname1.value = ""
    ELSE
		strEmp_no = Trim(frm1.txtEmp_no1.value)
        IntRetCd = CommonQueryRs(" NAME "," HAA010T "," EMP_NO =  " & FilterVar(strEmp_no , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
        If IntRetCd = false then
            frm1.txtname1.value = ""
            txtEmp_no1_Onchange = true            
            frm1.txtEmp_no1.focus
        ELSE    
            frm1.txtname1.value = Trim(Replace(ConvSPChars(lgF0),Chr(11),""))   '사번에 해당하는 이름 

        END IF
    END IF 
End Function

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
								<TD width="100" height="20" bgcolor="D4E5E8" class=base1>사용여부
								</TD>
								<TD width="626" height="20" bgcolor="FFFFFF" class="base2" align=left valign=center colspan=3>&nbsp;&nbsp;
									<INPUT TYPE="RADIO" NAME="txtuse_yn" CLASS="radio_title" CHECKED ID="txtuse_yn0" VALUE=A ><LABEL FOR="txtuse_yn0">전체</LABEL>
    								<INPUT TYPE="RADIO" NAME="txtuse_yn" CLASS="radio_title" ID="txtuse_yn1" VALUE=Y ><LABEL FOR="txtuse_yn1">사용</LABEL>
									<INPUT TYPE="RADIO" NAME="txtuse_yn" CLASS="radio_title" ID="txtuse_yn2" VALUE=N ><LABEL FOR="txtuse_yn2">미사용</LABEL>
								</TD>
                            </tr>
                            <tr>
								<TD width="100" height="30" bgcolor="D4E5E8" class=base1>사번</TD>
								<TD width="626" height="30" bgcolor="FFFFFF" class="base2" align=left valign=center colspan=3>&nbsp;&nbsp;
								<INPUT CLASS="form01" NAME="txtEmp_no1" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="12">&nbsp;<IMG SRC="../ESSimage/button_13.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenEmp(frm1.txtEmp_no1.value)">
								<INPUT CLASS="form02" NAME="txtname1" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="14">								
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
		                        <TD class=TDFAMILY_TITLE1>사용자ID</TD>
		                        <TD class=TDFAMILY_TITLE1>성명</TD>
		                        <TD class=TDFAMILY_TITLE1>사번</TD>
		                        <TD class=TDFAMILY_TITLE1>부서명</TD>
		                        <TD class=TDFAMILY_TITLE1>레벨</TD>
		                        <TD class=TDFAMILY_TITLE1>자료권한</TD>
		                        <TD class=TDFAMILY_TITLE1>사용</TD>
		                        <TD class=TDFAMILY_TITLE1>등록일</TD>
		                        <TD class=TDFAMILY_TITLE1></TD>
                            </TR>
							<% 
							For i=1 To 8
							     Response.Write "<TR bgcolor=#F8F8F8 height=24 onclick='vbscript: Call DoubleGetRow(" & i & ")' onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
							     Response.Write "<TD><INPUT class=listrow01 name='" & i & "'  flag='SPREADCELL' style='WIDTH: 30px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_ID" & i & "' flag='SPREADCELL' style='WIDTH: 100px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_NAME' flag='SPREADCELL' style='WIDTH: 90px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_emp_no" & i & "' flag='SPREADCELL' style='WIDTH:  100px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_deptnm" & i & "' flag='SPREADCELL' style='WIDTH:  128px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_pro_auth" & i & "' flag='SPREADCELL' style='WIDTH: 80px; text-align: left;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_dept_auth' flag='SPREADCELL' style='WIDTH: 65px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL' flag='SPREADCELL' style='WIDTH: 50px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
							     Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL' flag='SPREADCELL' style='WIDTH: 78px; text-align: center;' onMouseOver='vbscript: Call MouseRow(" & i & ")' readonly></TD>"
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
