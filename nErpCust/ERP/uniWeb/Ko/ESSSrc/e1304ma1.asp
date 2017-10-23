<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../ESSinc/IncServer.asp"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
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
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "e1304mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "e1304ma2.asp"						           '☆: Biz Logic ASP Name
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../ESSinc/lgvariables.inc" --> 
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim StartDate
<%
StartDate	= GetSvrDate
%>

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 


End Sub
'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    if  pOpt = "Q" then
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & "" & gColSep
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
    end if
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status

    parent.document.All("nextprev").style.VISIBILITY = "hidden"
    
    Call InitComboBox()

    Call LayerShowHide(0)

    Call SetToolBar("00000")

    frm1.txtBas_dt.value = UniConvDateAToB("<%=StartDate%>",gServerDateFormat,gDateFormat)
	frm1.txtRetire_dt.value = frm1.txtBas_dt.value
    Call LockField(Document)
    Call DbQuery(1)

End Sub
'========================================================================================
' Function Name : Form_UnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
 	Set gActiveElement = Nothing
End Sub

Function DbQuery(ppage)
    Dim strVal

    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    'If Grid1.ChkChange() Then Exit Function
    'Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")

    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True      
                                                             '☜: Processing is NG
    If Not chkFieldLength(Document, "1") Then									         '☜: This function check required field
		Exit Function
	end if
                                                             
End Function


Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status
	lgIntFlgMode      = OPMD_UMODE                                              '⊙: Indicates that current mode is Create mode
    'Call Grid1.ShowData(frm1,1)
'    Call ElementVisible(window.parent.document.all("RunQuery"), 1)


End Function

Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status

End Function


Function ExeReflectOk()

    Dim strVal
    Err.Clear                                                                    '☜: Clear err status
    'Call Grid1.ShowData(frm1,1)
'    Call ElementVisible(window.parent.document.all("RunQuery"), 1)
    strVal = BIZ_PGM_ID1 & "?txtEmp_no=" & frm1.txtEmp_no.value
    strVal = strVal & "&txtBas_dt=" & frm1.txtBas_dt.value
 '   strVal = strVal & "&txtRetire_dt=" & frm1.txtRetire_dt.value    
    strVal = strVal & "&txtBackPgmId=" & self.location
    
    document.location = strVal
    
End Function

Function ExeReflectNo()
    Err.Clear
    'Call ClearField(Document,2)                                                                    '☜: Clear err status
'    Call ElementVisible(window.parent.document.all("RunQuery"), 0)


'    Call DisplayMsgBox("187742","X","X","X")

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

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	With Frm1
		.txtMode.value        = "UID_M0002"                                        '☜: Save
'		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

    'Call InitVariables

    strVal = BIZ_PGM_ID1 & "?txtEmp_no=" & frm1.txtEmp_no.value
    strVal = strVal & "&txtBas_dt=" & frm1.txtBas_dt.value
    
    document.location = strVal
	
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
  '  Call DbQuery(1)
End Function

Sub SubPrint(objFrame)
    Set objActiveEl = document.activeElement
    objFrame.focus()
    objFrame.print()
    objActiveEl.focus
    Set objActiveEl = nothing
End Sub

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
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

Function FncCalc()

    Dim strDate, strBas_dt, strEntr_dt
	Dim strVal
	Dim intRetCD
    Err.Clear                                                                    '☜: Clear err status

	With Frm1

        if  Date_chk(.txtBas_dt.value , strDate) = True then
            .txtBas_dt.value = strDate
        else
            Call DisplayMsgBox("800094","X","X","X")
            .txtBas_dt.focus()
            exit function
        end if
        intRetCD = CommonQueryRs(" CONVERT(VARCHAR(8),ISNULL(group_entr_dt,entr_dt),112) "," haa010t "," emp_no =  " & FilterVar(frm1.txtEmp_no.Value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)   

        If intRetCD = false Then
        Else
           strEntr_dt = Replace(lgF0, Chr(11), "")
           strBas_dt = UNIConvDateToYYYYMMDD(.txtBas_dt.value,gDateFormat,"")

           If strEntr_dt >= strBas_dt Then
               Call DisplayMsgBox("800443","X","계산기준일","입사일")  
               .txtBas_dt.focus()
               Exit function     
           End If
        End If
    End With

	FncCalc = False														         '☜: Processing is NG
		
	Call LayerShowHide(1)
    Call MakeKeyStream("Q")
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	With Frm1
		.txtMode.value        = "UID_M0002"                                        '☜: Save
'		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    FncCalc  = True                                                               '☜: Processing is NG                                '☜:  Run biz logic

End Function

</SCRIPT>

<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->

</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 border=0 width=732>
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
	                        		<TD CLASS=ctrow01 >예상퇴사일</TD>
	                        		<TD CLASS=ctrow02 align=left>
	                        		    <INPUT CLASS="form01" ID="txtRetire_dt" NAME="txtRetire_dt" TYPE="Text" MAXLENGTH=10 SiZE=12 style='text-align: center;' ondblclick="VBScript:Call OpenCalendar('txtRetire_dt',3)">
	                        		</TD>
                                </TR>                                 
                                <TR>
	                        		<TD CLASS=ctrow01 >퇴직금산정기준일</TD>
	                        		<TD CLASS=ctrow02 align=left>
	                        		    <INPUT CLASS="form01" ID="txtBas_dt" NAME="txtBas_dt" TYPE="Text" MAXLENGTH=10 SiZE=12 style='text-align: center;' ondblclick="VBScript:Call OpenCalendar('txtBas_dt',3)">
	                        		</TD>
                                </TR>                                
                                <TR>
	                        		<TD CLASS=ctrow01>*유의사항</TD>
	                        		<TD align=left CLASS=ctrow04 width=300>계산된 결과는 실제와 다를 수 있습니다.
	                        		</TD>
                                </TR>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
        </TR>
        <TR valign=middle height=50>
            <TD colspan=2 align=center background=#000000>
				<INPUT type="image" SRC="../ESSimage/button_14.gif" border="0" OnClick="vbscript: FncCalc()" name="printprev" alt='계산' onMouseOver="javascript:this.src='../ESSimage/button_r_14.gif';" onMouseOut="javascript:this.src='../ESSimage/button_14.gif';">
        </TR>
    </TABLE>
    <TABLE cellSpacing=0 cellPadding=0 border=0 bgcolor=#ffffff>
        <TR><TD WIDTH="100%" HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>

    <INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtres_no"    TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtdomi"    TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtaddr"    TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtentr_dt"    TAG="24">
</FORM>	
</BODY>
</HTML>
