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

Const C_SHEETMAXCOLS = 8

Const BIZ_PGM_ID      = "e1608mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "e1608ma2.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 
<!-- #Include file="../ESSinc/incGrid.inc" -->

Dim Grid1
dim fDiligAuth,fAuthCheck

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
    lgKeyStream = lgKeyStream & Trim(frm1.txtYear.Value) & gColSep
    lgKeyStream = lgKeyStream & Trim(fDiligAuth) & gColSep        
    lgKeyStream = lgKeyStream & Trim(fAuthCheck) & gColSep     
End Sub   
     
'========================================================================================================
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    Dim lgYear,i,stYear

	lgYear = Year(date)

	if Trim(parent.txtemp_no.value)="unierp" then
		stYear=lgyear-1
	else
		Call CommonQueryRs("entr_dt "," haa010t ","emp_no =  " & FilterVar(parent.txtemp_no.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		if lgF0="" then
			stYear = "1990"
		else
			stYear =Year(lgF0)
		end if
	end if	
    For i=lgYear To cint(stYear) step -1
    	Call SetCombo(frm1.txtYear, i, i)
    Next

    frm1.txtYear.value = CStr(lgYear)
    
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
'========================================================================================================
Sub Form_Load()
   
    Err.Clear                                                                       '☜: Clear err status

    call FncGetDiligAuth(fDiligAuth,fAuthCheck)
    If Replace(fDiligAuth,Chr(11),"") = "" Then
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
    Else
        parent.document.All("nextprev").style.VISIBILITY = "visible"
    End If
    
    Call InitComboBox()    
    Call LayerShowHide(0)
    Call InitGrid()
    Call SetToolBar("10000")
	if parent.txtName2.value = "" then
		parent.txtEmp_no2.Value = parent.txtemp_no.value 
	end if

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

'========================================================================================
' Function Name : Detail_Dilig
'========================================================================================
Function Detail_Dilig(pRow,pCol)
	Dim strVal

    If pRow <> "" and pCol <> "" then

        strVal = BIZ_PGM_ID1 & "?emp_no=" & frm1.txtEmp_no.value
        strVal = strVal & "&day=" & pRow & "&dilig_cd=" & pCol
        strVal = strVal & "&year=" & Trim(frm1.txtYear.Value)

        document.location = strVal

    end if
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
								<td width="80" height="27" bgcolor="D4E5E8" class="base1">사번</td>
								<td width="85" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">성명</td>
								<td width="80" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtName" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">직위</td>
								<td width="80" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">부서</td>
								<td width="153" bgcolor="FFFFFF"><INPUT class=base2 NAME="txtDept_nm" MAXLENGTH=25 SiZE=22  readonly></td>
                            </tr>
                            <tr> 
								<td width="80" height="30" bgcolor="D4E5E8" class=base1 valign=middle>정산년도
								</td>
								<td width="85" bgcolor="FFFFF" align=center>
								    <SELECT Name="txtYear" tabindex=-1 class=base2>
								    </SELECT>
								</td>
								<td bgcolor="FFFFFF" class="base1">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base2">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base1">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base2">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base1">&nbsp;</td>
								<td bgcolor="FFFFFF" class="base2">&nbsp;</td>
                            </tr>
                            </table>
                        </td>
                    </TR>
                    <TR>
                       <td height="5"></td>
                    </TR>
                    <TR>
                        <td><table width="100%" border="0" cellspacing="1" cellpadding="0" bgcolor="DDDDDD">
							<TR height=27> 
		                    	<TD class=TDFAMILY_TITLE1 style='WIDTH: 75px'>월</TD>
		                    	<TD class=TDFAMILY_TITLE1><INPUT class=TDFAMILY_TITLE1 name='TITLE1' style='WIDTH: 90px' readonly></TD>
		                    	<TD class=TDFAMILY_TITLE1><INPUT class=TDFAMILY_TITLE1 name='TITLE2' style='WIDTH: 90px' readonly></TD>		                        	
		                    	<TD class=TDFAMILY_TITLE1><INPUT class=TDFAMILY_TITLE1 name='TITLE3' style='WIDTH: 90px' readonly></TD>
		                    	<TD class=TDFAMILY_TITLE1><INPUT class=TDFAMILY_TITLE1 name='TITLE4' style='WIDTH: 90px' readonly></TD>
		                    	<TD class=TDFAMILY_TITLE1><INPUT class=TDFAMILY_TITLE1 name='TITLE5' style='WIDTH: 90px' readonly></TD>
		                    	<TD class=TDFAMILY_TITLE1><INPUT class=TDFAMILY_TITLE1 name='TITLE6' style='WIDTH: 90px' readonly></TD>
		                    	<TD class=TDFAMILY_TITLE1><INPUT class=TDFAMILY_TITLE1 name='TITLE7' style='WIDTH: 90px' readonly></TD>		                        	
                            </TR>
<%            
        For i=1 To 12
            Response.Write "<TR bgcolor=#F8F8F8 height=24>"
            Response.Write "<TD class=TDFAMILY_TITLE1 >"
            Response.Write  "<INPUT class=TDFAMILY_TITLE1 name='MONTH" & i & "'  style='WIDTH:  75px; '  readonly>"
			Response.Write  "<INPUT class=TDFAMILY_TITLE1 name='temp"	& i & "' type='hidden' flag='SPREADCELL'  style='WIDTH:  80px;' readonly></TD>"            
            
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
            Response.Write "<INPUT class=listrow01 name='SPREADCELL1_" & i & "' flag='SPREADCELL' tag='2' style='WIDTH: 90px; TEXT-ALIGN: center; ' onMouseOver='vbscript: Call MouseRow(" & i &")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,1)' readonly></TD>"
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
            Response.Write "<INPUT class=listrow01 name='SPREADCELL2_" & i & "' flag='SPREADCELL' tag='2' style='WIDTH: 90px; TEXT-ALIGN: center; '  onMouseOver='vbscript: Call MouseRow(" & i & ")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,2)' readonly></TD>"
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
            Response.Write "<INPUT class=listrow01 name='SPREADCELL3_" & i & "' flag='SPREADCELL' tag='2' style='WIDTH: 90px; TEXT-ALIGN: center; '	onMouseOver='vbscript: Call MouseRow(" & i & ")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,3)' readonly></TD>"
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
            Response.Write "<INPUT class=listrow01 name='SPREADCELL4_" & i & "' flag='SPREADCELL' tag='2' style='WIDTH: 90px; TEXT-ALIGN: center; '	onMouseOver='vbscript: Call MouseRow(" & i & ")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,4)' readonly></TD>"
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
        	Response.Write "<INPUT class=listrow01 name='SPREADCELL5_" & i & "' flag='SPREADCELL' tag='2' style='WIDTH: 90px; TEXT-ALIGN: center; '	onMouseOver='vbscript: Call MouseRow(" & i & ")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,5)' readonly></TD>"
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
        	Response.Write "<INPUT class=listrow01 name='SPREADCELL6_" & i & "' flag='SPREADCELL' tag='2' style='WIDTH: 90px; TEXT-ALIGN: center; '	onMouseOver='vbscript: Call MouseRow(" & i & ")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,6)' readonly></TD>"
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='E1EEF1'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
        	Response.Write "<INPUT class=listrow01 name='SPREADCELL7_" & i & "' flag='SPREADCELL' tag='2' style='WIDTH: 90px; TEXT-ALIGN: center; '	onMouseOver='vbscript: Call MouseRow(" & i & ")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,7)' readonly></TD>"
        	Response.Write "</TR>"
        Next
%>
            <TR bgcolor=#F8F8F8 >
            <TD class=TDFAMILY_TITLE1 style='WIDTH: 80px'>합계</TD>
            <TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM1' class=TDFAMILY_TITLE2 style='WIDTH: 30px;' tag='2' readonly></TD>
            <TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM2' class=TDFAMILY_TITLE2 style='WIDTH: 30px;' tag='2' readonly></TD>
            <TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM3' class=TDFAMILY_TITLE2 style='WIDTH: 30px;' tag='2' readonly></TD>
        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM4' class=TDFAMILY_TITLE2 style='WIDTH: 30px;' tag='2' readonly></TD>
        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM5' class=TDFAMILY_TITLE2 style='WIDTH: 30px;' tag='2' readonly></TD>
        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM6' class=TDFAMILY_TITLE2 style='WIDTH: 30px;' tag='2' readonly></TD>
        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM7' class=TDFAMILY_TITLE2 style='WIDTH: 30px;' tag='2' readonly></TD>
            </TR>
                           </table>
                        </td>
                    </TR>
                </TABLE>
            </TD>
        </TR>
        <TR>
            <TD height=5></TD>
        </TR>
    </TABLE>
    <TABLE cellSpacing=0 cellPadding=0 border=0>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100" HEIGHT=100 FRAMEBORDER=0 SCROLLING=auto noresize framespacing=0></IFRAME></TD></TR>
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
