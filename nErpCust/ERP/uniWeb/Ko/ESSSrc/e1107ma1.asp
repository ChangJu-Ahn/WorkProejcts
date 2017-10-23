<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>

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
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<Script Language="VBScript">
Option Explicit 

Const BIZ_PGM_ID      = "e1107mb1.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

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
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0019", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtMil_type, iCodeArr, iNameArr, Chr(11))    

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0020", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtMil_kind, iCodeArr, iNameArr, Chr(11))    

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0021", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtMil_grade, iCodeArr, iNameArr, Chr(11))    

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0022", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtMil_branch, iCodeArr, iNameArr, Chr(11))    
End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = 4+1
    Grid1.SheetMaxrows = 3
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
    if  parent.txtDEPT_AUTH.value = "Y" then
        parent.document.All("nextprev").style.VISIBILITY = "visible"
        Call SetToolBar("10010")    
    else
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
        Call SetToolBar("00010")    
    end if

    Call InitComboBox()

    Call LayerShowHide(0)
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

'========================================================================================================
' Name : FncNew
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
	frm1.txtEmp_no.value = ""
	frm1.txtName.value = ""
	frm1.txtroll_pstn.value = ""
	frm1.txtDept_nm.value = ""
	
    Call ClearField(document,2)
    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
' Name : DbQuery
'========================================================================================================
Function DbQuery(ppage)

    Dim strVal,IntRetCD
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    
    Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG

End Function

'========================================================================================================
' Name : DbQueryOk
'========================================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

    if  parent.txtDEPT_AUTH.value = "Y" then
        parent.document.All("nextprev").style.VISIBILITY = "visible"
        if  frm1.txtEmp_no.value = parent.txtEmp_no.Value then
            Call SetToolBar("10010")
        else
            Call SetToolBar("10000")
        end if
    else
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
        Call SetToolBar("00010")
    end if

End Function

'========================================================================================================
' Name : DbQueryFail
'========================================================================================================
Function DbQueryFail()
    Err.Clear
    if  parent.txtDEPT_AUTH.value = "Y" then
        parent.document.All("nextprev").style.VISIBILITY = "visible"
        if  frm1.txtEmp_no.value = parent.txtEmp_no.Value then
            Call SetToolBar("10010")
        else
            Call SetToolBar("10000")
        end if
    else
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
        Call SetToolBar("00010")
    end if

    Call ClearField(Document,2)                                                                    '☜: Clear err status

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
	Dim strDate
    Err.Clear                                                                    '☜: Clear err status

    if frm1.txtMil_start.value = "" then
    elseif Date_chk(frm1.txtMil_start.value, strDate) = True then
        frm1.txtMil_start.value = strDate
    else
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtMil_start.focus()
        exit function
    end if

    if frm1.txtMil_end.value = "" then
    elseif Date_chk(frm1.txtMil_end.value, strDate) = True then
        frm1.txtMil_end.value = strDate
    else
        Call DisplayMsgBox("800094","X","X","X")
        frm1.txtMil_end.focus()
        exit function
    end if

	With Frm1	
	    if chkField(Document,"2") then
	    	exit function
	    end if
	    if  uniConvDateToYYYYMMDD(frm1.txtMil_start.value,gDateFormat,"-") >= uniConvDateToYYYYMMDD(frm1.txtMil_end.value,gDateFormat,"-") then
            Call DisplayMsgBox("800002","X","X","X")
           .txtMil_start.focus()
            exit function
        end if
	End With
    
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
<FORM NAME=frm1 target=MyBizASP METHOD="POST">
    <TABLE width=733 cellSpacing=0 cellPadding=0 border=0>
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
								<td width="85" bgcolor="FFFFFF"><INPUT class="base2" NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">성명</td>
								<td width="86" bgcolor="FFFFFF"><INPUT class="base2" NAME="txtName" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">직위</td>
								<td width="100" bgcolor="FFFFFF"><INPUT class="base2" NAME="txtroll_pstn" MAXLENGTH=20 SiZE=16  readonly></td>
								<td width="60" height="27" bgcolor="D4E5E8" class="base1">부서</td>
								<td width="153" bgcolor="FFFFFF"><INPUT class="base2" NAME="txtDept_nm" MAXLENGTH=25 SiZE=22  readonly></td>
                            </tr>
                            </table>
                        </td>
                    </TR>
                    <TR>
                       <td height="10"></td>
                    </TR>
                    <TR>
                        <TD>
                            <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#DDDDDD>
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>병역구분</TD>
		                            <TD CLASS="ctrow02">
		                                <SELECT NAME="txtMil_type" CLASS=ctrow02 STYLE="WIDTH: 120px"><OPTION VALUE=""></OPTION></SELECT>
		                            </TD>
		                            <TD CLASS="ctrow01" NOWRAP>병역군별</TD>
		                            <TD CLASS="ctrow02">
		                                <SELECT NAME="txtMil_kind" CLASS=ctrow02 STYLE="WIDTH: 120px"><OPTION VALUE=""></OPTION></SELECT>
	                                </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="ctrow01" NOWRAP>복무기간</TD>
		                            <TD CLASS="ctrow02">
		                                <INPUT CLASS="form02" NAME="txtMil_start" ALT="복무기간1" TYPE="Text" MAXLENGTH=10 SiZE=12 ondblclick="VBScript:Call OpenCalendar('txtMil_start',3)">&nbsp;~&nbsp;
		                                <INPUT CLASS="form02" NAME="txtMil_End" ALT="복무기간2" TYPE="Text" MAXLENGTH=10 SiZE=12 ondblclick="VBScript:Call OpenCalendar('txtMil_End',3)">
                                    </TD>      
		                            <TD CLASS="ctrow01" NOWRAP>병역등급</TD>
		                            <TD CLASS="ctrow02">
		                                <SELECT NAME="txtMil_grade" CLASS=ctrow02 STYLE="WIDTH: 120px"><OPTION VALUE=""></OPTION></SELECT>
		                            </TD>
		                        </TR>
		                        <TR>
		                            <TD CLASS="ctrow01" NOWRAP>병역병과</TD>
		                            <TD CLASS="ctrow02">
		                                <SELECT NAME="txtMil_branch" CLASS=ctrow02 STYLE="WIDTH: 120px"><OPTION VALUE=""></OPTION></SELECT>
		                            </TD>
		                            <TD CLASS="ctrow01" NOWRAP>군번</TD>
		                            <TD CLASS="ctrow02"><INPUT CLASS="form02" NAME="txtMil_no" ALT="군번" TYPE="Text" MAXLENGTH=10 SiZE=14></TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
                                </TR>
		                        <TR>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
		                            <TD CLASS="ctrow01"></TD>
		                            <TD CLASS="ctrow02"></TD>
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
