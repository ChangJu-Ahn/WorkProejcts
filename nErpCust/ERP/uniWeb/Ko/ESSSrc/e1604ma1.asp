<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>

<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!-- #Include file="../ESSinc/incServer.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/Common.css">

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

Const BIZ_PGM_ID      = "e1604mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXCOLS = 8

<!-- #Include file="../ESSinc/lgvariables.inc" --> 
<!-- #Include file="../ESSinc/incGrid.inc" -->

Dim Grid1
Dim StartDate,EndDate

<%StartDate	= GetSvrDate
EndDate		= UNIDateAdd("M",-1,GetSvrDate,gServerDateFormat)%>

EndDate = "<%=StartDate%>"
StartDate =  "<%=EndDate%>"

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    if  pOpt = "Q" then
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & "" & gColSep
        'lgKeyStream = lgKeyStream & Trim(frm1.txtfrom.Value) & gColSep
        'lgKeyStream = lgKeyStream & Trim(frm1.txtto.Value) & gColSep
        lgKeyStream = lgKeyStream & UniConvDateAToB(frm1.txtfrom.Value,gDateFormat, gServerDateFormat) & gColSep
        lgKeyStream = lgKeyStream & UniConvDateAToB(frm1.txtto.Value,gDateFormat, gServerDateFormat) & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtdilig_cd.Value) & gColSep
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
    end if
End Sub  
      
'========================================================================================================
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx

	Call CommonQueryRs(" dilig_cd, dilig_nm "," hca010t ", " dilig_cd not in (" & FilterVar("98", "''", "S") & "," & FilterVar("99", "''", "S") & ") " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtdilig_cd, iCodeArr, iNameArr,Chr(11))
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = C_SHEETMAXCOLS
    Grid1.SheetMaxrows = 9
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

    parent.document.All("nextprev").style.VISIBILITY = "hidden"

    Call InitComboBox()

    Call LayerShowHide(0)

    Call InitGrid()

    Call SetToolBar("10000")
    Call LockField(Document)

    frm1.txtfrom.value = UniConvDateAToB(StartDate,gServerDateFormat,gDateFormat)
    frm1.txtto.value = UniConvDateAToB(EndDate,gServerDateFormat,gDateFormat)

    Call DbQuery(1)
End Sub

'========================================================================================
' Function Name : Form_unLoad
'========================================================================================
Sub Form_unLoad()
	On Error Resume Next
 	Set gActiveElement = Nothing
    Set Grid1 = Nothing
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery(ppage)

    Dim strDate
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

	With Frm1

        if  Date_chk(.txtfrom.value, strDate) = True then
            .txtfrom.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txtfrom.focus()
            exit function
        end if

        if  Date_chk(.txtto.value, strDate) = True then
            .txtto.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txtto.focus()
            exit function
        end if
  		If CompareDateByFormat(.txtfrom.value,.txtto.value,"근태시작일","근태종료일","970025", gDateFormat, gComDateType,TRUE) = False Then
		    .txtfrom.focus()
		    exit function
		end if
    End With
    If Not chkFieldLength(Document, "1") Then									         '☜: This function check required field
		Exit Function
	end if
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

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	Call LayerShowHide(1)

	With Frm1
		.txtMode.value        = "UID_M0002"                                        '☜: Save
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
' Function Name : GetRow
'========================================================================================================
Function GetRow(pRow)
	GetRow=False
    Grid1.ActiveRow = pRow
    If Mid(document.activeElement.getAttribute("tag"),3,1) = "1" Then
	    arrRet = window.showModalDialog("../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	GetRow=True
End Function

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Sub Query_OnClick()
    Call DbQuery(1)
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
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">사번</td>
								<td width="85" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">성명</td>
								<td width="86" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtName" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">직위</td>
								<td width="100" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">부서</td>
								<td width="153" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtDept_nm" MAXLENGTH=25 SiZE=22  readonly></td>
                            </tr>
							<tr height=25 valign=top>
							    <TD width="60" height="30" bgcolor="D4E5E8" class=base1 valign=center>기간</TD>
							    <TD bgcolor="FFFFFF" align=left valign=center colspan=3>&nbsp;&nbsp;
							        <INPUT ID="txtfrom" NAME="txtfrom" MAXLENGTH=16 SiZE=12 tag="12" ondblclick="VBScript:Call OpenCalendar('txtfrom',3)" style='font-family: "돋움"; font-size: 9pt; color: 002232; padding-left: 12px;'>&nbsp;~
							        <INPUT ID="txtto" NAME="txtto" MAXLENGTH=16 SiZE=12 tag="12" ondblclick="VBScript:Call OpenCalendar('txtto',3)" style='font-family: "돋움"; font-size: 9pt; color: 002232; padding-left: 12px; valign:center;'>
							    </TD>
							    <TD width="60" height="30" bgcolor="D4E5E8" class=base1 valign=center>근태</TD>
							    <TD bgcolor="FFFFFF" align=left valign=center colspan=3>&nbsp;&nbsp;
									<SELECT NAME="txtDilig_cd" tabindex=-1 ALT="근태" class=form01><OPTION VALUE=""></OPTION></SELECT>
							    </TD>
							</tr>
                            </table>
                        </td>
                    </TR>
                    <TR>
                       <td height="10"></td>
                    </TR>
                    <TR>
                        <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
								<TR> 
								    <TD class=TDFAMILY_TITLE1></TD>
		                        	<TD class=TDFAMILY_TITLE1>근태일자</TD>
		                        	<TD class=TDFAMILY_TITLE1>근태</TD>
		                        	<TD class=TDFAMILY_TITLE1>회수</TD>
		                        	<TD class=TDFAMILY_TITLE1>시간</TD>
		                        	<TD class=TDFAMILY_TITLE1>분</TD>
								    <TD class=TDFAMILY_TITLE1></TD>
                                </TR>
							    <% 
                                For i=1 To 9
                                    Response.Write "<TR bgcolor=#F8F8F8 height=24 onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
                                    Response.Write "<TD><INPUT class=listrow01 name='" & i & "'  flag='SPREADCELL' style='WIDTH:  30px; text-align: center;'  readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_DILIG_DT" & i & "' flag='SPREADCELL' style='WIDTH: 100px; text-align: center;' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL_DILIG_CD" & i & "' type=hidden flag='SPREADCELL' style='WIDTH:   0px; text-align: center;'>"
                                	Response.Write "    <INPUT class=listrow01 name='SPREADCELL" & i & "' flag='SPREADCELL' style='WIDTH: 230px; text-align: left;' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL" & i & "' flag='SPREADCELL' style='WIDTH: 85px; text-align: left;' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL" & i & "' flag='SPREADCELL' style='WIDTH: 85px; text-align: center;' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL" & i & "' flag='SPREADCELL' style='WIDTH: 85px; text-align: center;' readonly></TD>"
                                    Response.Write "<TD><INPUT class=listrow01 name='SPREADCELL" & i & "' flag='SPREADCELL' style='WIDTH:110px; text-align: center;' readonly></TD>"
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
