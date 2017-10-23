<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!-- #Include file="../../inc/incServer.asp"  -->
<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/CommStyleSheet.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXCOLS = 8

Const BIZ_PGM_ID      = "e1608mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "e1608ma2.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
<!-- #Include file="../../inc/incGrid.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim Grid1
dim fDiligAuth,fAuthCheck

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
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
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    Dim lgYear,i,stYear

'	lgYear = "<%=request("year")%>"	    
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
' Function Desc : This method initializes spread sheet column property
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
    On Error Resume Next
    
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
' Function Name : Window_onUnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_unLoad()
	On Error Resume Next
    Set Grid1 = Nothing
End Sub

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

Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
    Call Grid1.ShowData(frm1,frm1.GRID_PAGE.VALUE)
End Function

Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status
End Function

Sub SubPrint(objFrame)
    Set objActiveEl = document.activeElement
    objFrame.focus()
    objFrame.print()
    objActiveEl.focus
    Set objActiveEl = nothing
End Sub

Function GetRow(pRow)
	GetRow=False
    Grid1.ActiveRow = pRow
    If Mid(document.activeElement.getAttribute("tag"),3,1) = "1" Then
	    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	GetRow=True
End Function

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
Function Detail_Dilig(pRow,pCol)
	Dim strVal

    If pRow <> "" and pCol <> "" then

        strVal = BIZ_PGM_ID1 & "?emp_no=" & frm1.txtEmp_no.value
        strVal = strVal & "&day=" & pRow & "&dilig_cd=" & pCol

        document.location = strVal

    end if
End Function

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

Sub Query_OnClick()
    Call DbQuery(1)
End Sub

Sub Print_onClick()
    Call SubPrint(MyBizASP)
End Sub


Sub GRID_PAGE_OnChange()
End Sub

Sub DELETE_OnClick()
    Call Grid1.DeleteClick()
End Sub

Sub CANCEL_OnClick()
    Call Grid1.CancelClick()
End Sub
Function txtEmp_no2_Onchange()
        Call DbQuery(1)	
End Function
'========================================================================================================
' Name : FncGetDiligAuth()
' Desc : developer describe this line 
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
<!-- #Include file="../../inc/uniSimsClassID.inc" --> 

</HEAD>
<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<TABLE CELLPADDING=0 CELLSPACING=0 BORDER=0>
<TR>
<FORM NAME="frm1" TARGET="MyBizASP" METHOD="post">
<TD>
    <TABLE cellSpacing=0 cellPadding=0 width=749 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=722 border=0 bgcolor=#ffffff>
                    <TR height=30 valign=center>
                        <TD class=base1>사번:<INPUT class=base1 NAME="txtEmp_no" MAXLENGTH=13 SiZE=12 TAG=14></TD>
                        <TD class=base1>성명:<INPUT class=base1 NAME="txtName" MAXLENGTH=20 SiZE=10 TAG=14></TD>
                        <TD class=base1>직위:<INPUT class=base1 NAME="txtroll_pstn" MAXLENGTH=20 SiZE=10 TAG=14></TD>
                        <TD class=base1>부서:<INPUT class=base1 NAME="txtDept_nm" MAXLENGTH=25 SiZE=20 TAG=14></TD>
                    </TR>
                    <TR height=25 valign=top>
                        <TD class=base1 valign=middle>정산년도:
						    <SELECT Name="txtYear" tabindex=-1 STYLE="WIDTH: 60px">
						    </SELECT>
                        </TD>
                    </TR>
                   <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width="100%" border=0 bgcolor=#ffffff>
                                <TR bgcolor=#d0d6e4 height=20>
		                        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 80px' colspan=1>월</TD>
		                        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px' colspan=1><INPUT name='TITLE1' tag='25X'  style='WIDTH: 30px; TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>
		                        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px' colspan=1><INPUT name='TITLE2' tag='25X'  style='WIDTH: 30px; TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>		                        	
		                        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px' colspan=1><INPUT name='TITLE3' tag='25X'  style='WIDTH: 30px; TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>
		                        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px' colspan=1><INPUT name='TITLE4' tag='25X'  style='WIDTH: 30px; TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>
		                        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px' colspan=1><INPUT name='TITLE5' tag='25X'  style='WIDTH: 30px; TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>
		                        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px' colspan=1><INPUT name='TITLE6' tag='25X'  style='WIDTH: 30px; TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>
		                        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px' colspan=1><INPUT name='TITLE7' tag='25X'  style='WIDTH: 30px; TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>		                        	
                                </TR>
<%            
        For i=1 To 12
            Response.Write "<TR bgcolor=#E9EDF9 height=20>"
            Response.Write "<TD class=TDFAMILY_TITLE1 style='WIDTH: 80px'>"
            Response.Write  "<INPUT name='MONTH"	  & i & "' tag='25X'  style='WIDTH:  80px; TEXT-ALIGN: center;FONT-WEIGHT: bolder'  readonly>"
			Response.Write  "<INPUT name='temp"	  & i & "' type='hidden' tag='25X' flag='SPREADCELL'  style='WIDTH:  80px; TEXT-ALIGN: center;FONT-WEIGHT: bolder'  readonly></TD>"            
            
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
            Response.Write "<INPUT name='SPREADCELL1_"	  & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 90px; TEXT-ALIGN: center'' onMouseOver='vbscript: Call MouseRow(" & i &")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,1)' readonly></TD>"
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
            Response.Write "<INPUT name='SPREADCELL2_"	  & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 90px; TEXT-ALIGN: center'  onMouseOver='vbscript: Call MouseRow(" & i & ")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,2)' readonly></TD>"
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
            Response.Write "<INPUT name='SPREADCELL3_"		  & i & "' tag='25x' flag='SPREADCELL' style='WIDTH: 90px; TEXT-ALIGN: center'	onMouseOver='vbscript: Call MouseRow(" & i & ")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,3)' readonly></TD>"
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
            Response.Write "<INPUT name='SPREADCELL4_" & i & "' tag='25x' flag='SPREADCELL' style='WIDTH: 90px; TEXT-ALIGN: center'	onMouseOver='vbscript: Call MouseRow(" & i & ")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,4)' readonly></TD>"
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
        	Response.Write "<INPUT name='SPREADCELL5_"  & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 90px; TEXT-ALIGN: center'	onMouseOver='vbscript: Call MouseRow(" & i & ")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,5)' readonly></TD>"
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
        	Response.Write "<INPUT name='SPREADCELL6_"		  & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 90px; TEXT-ALIGN: center'	onMouseOver='vbscript: Call MouseRow(" & i & ")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,6)' readonly></TD>"
            Response.Write "<TD onMouseOver=" & chr(34) & "javascript: this.style.backgroundColor='FEE2E3'" & chr(34) & " onMouseOut=" & chr(34) & "javascript: this.style.backgroundColor=''" & chr(34) & ">"
        	Response.Write "<INPUT name='SPREADCELL7_"	 & i & "' tag='25X' flag='SPREADCELL' style='WIDTH: 90px; TEXT-ALIGN: center'	onMouseOver='vbscript: Call MouseRow(" & i & ")' onclick ='vbscript: Call Detail_Dilig(frm1.MONTH" & i &".value,7)' readonly></TD>"
        	Response.Write "</TR>"
        Next
%>
            <TR bgcolor=#E9EDF9 >
            <TD class=TDFAMILY_TITLE1 style='WIDTH: 80px'>합계</TD>
            <TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM1' 	tag='25X' style='WIDTH: 30px;TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>
            <TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM2'	tag='25X' style='WIDTH: 30px;TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>
            <TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM3'	tag='25x' style='WIDTH: 30px;TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>
        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM4'	tag='25X' style='WIDTH: 30px;TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>
        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM5'	tag='25X' style='WIDTH: 30px;TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>
        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM6'	tag='25X' style='WIDTH: 30px;TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>
        	<TD class=TDFAMILY_TITLE1 style='WIDTH: 90px'><INPUT name='SUM7'	tag='25X' style='WIDTH: 30px;TEXT-ALIGN: center;FONT-WEIGHT: bolder' readonly></TD>
            </TR>
                            </TABLE>
                        </TD>
                    </TR>                                                    
                </TABLE>
            </TD>
            <TD width=14></TD>
        </TR>
    </TABLE>
    <TABLE cellSpacing=2 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=hidden NAME="txtMode">
    <INPUT TYPE=hidden NAME="txtKeyStream">
    <INPUT TYPE=hidden NAME="txtUpdtUserId">
    <INPUT TYPE=hidden NAME="txtInsrtUserId">
    <INPUT TYPE=hidden NAME="txtFlgMode">
    <INPUT TYPE=hidden NAME="txtPrevNext">
    
    <INPUT TYPE=hidden NAME=GRID_TOTPAGES STYLE="WIDTH: 50px; HEIGHT: 20px">
    <INPUT TYPE=hidden NAME=GRID_PAGE  value=1 >
</TD>
</FORM>	
</TR>
</TABLE>
</BODY>
</HTML>
