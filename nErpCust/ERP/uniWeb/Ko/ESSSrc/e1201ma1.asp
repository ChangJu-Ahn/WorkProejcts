<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>

<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!-- #Include file="../ESSinc/incServer.asp" -->

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

Const BIZ_PGM_ID      = "e1201mb1.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
    if  pOpt = "Q" then
'        if  Trim(parent.txtEmp_no2.Value) = "" then
            lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep
'        else
'            lgKeyStream = Trim(parent.txtEmp_no2.Value) & gColSep
'        end if
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
        lgKeyStream = lgKeyStream & "" & gColSep
    else
        lgKeyStream = Trim(frm1.txtEmp_no.Value) & gColSep
        lgKeyStream = lgKeyStream & Trim(parent.txtinternal_cd.Value) & gColSep
    end if
End Sub        

'========================================================================================================
' Function Name : InitGrid
'========================================================================================================
Sub InitGrid()
    Set Grid1 = New Grid
    Grid1.MaxCols = 4+1
    Grid1.SheetMaxrows = 3
    Set Grid1.Source = document.frm1
End Sub

'========================================================================================================
' Name : LoadInfTB19029
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Private Sub LoadInfTB19029()

<!--#Include file="../ComAsp/LoadInfTB19029.asp"-->

<%Call loadInfTB19029(gCurrency,"Q","H")%>

End Sub
'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear           
'    if  parent.txtDEPT_AUTH.value = "Y" then
'        parent.document.All("nextprev").style.VISIBILITY = "visible"
'        Call SetToolBar("10000")    
'    else
        parent.document.All("nextprev").style.VISIBILITY = "hidden"
        Call SetToolBar("00000")    
'    end if

    Call LayerShowHide(0)
    Call LockField(Document)
	Call LoadInfTB19029()
    Call DbQuery(1)
End Sub

'========================================================================================
' Function Name : Form_UnLoad
'========================================================================================
Sub Form_UnLoad()
	On Error Resume Next
    Set Grid1 = Nothing
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
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
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear
    Call ClearField(Document,2)                                                                    '☜: Clear err status

End Function

'========================================================================================================
' Name : GetRow
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

'==========================================================================================
'   Event Name : Radio OnClick()
'==========================================================================================
Sub rdoUnionFlag1_OnClick()
	lgBlnFlgChgValue = True	
End Sub

Sub rdoUnionFlag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoPressFlag1_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoPressFlag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoOverseaFlag1_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoOverseaFlag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoResFlag1_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub rdoResFlag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

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
                       <td height="10"></td>
                    </TR>
                    <TR>
                        <TD>
                            <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#DDDDDD>
                                <TR>
				                	<TD CLASS="ctrow01">급여구분</TD>
				                	<TD CLASS="ctrow02"><INPUT class="form02" TYPE="Text" size=20 name=txtpay_cd readonly></TD>
				                	<TD CLASS="ctrow01">연봉(연봉직)</TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="form02" TYPE="Text" size=20 name=txtAnnualSal style="text-align:right" readonly></TD>
				                </TR>
				                <TR>
				                	<TD CLASS="ctrow01">기본급(연봉)</TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="form02" TYPE="Text" size=20 name=txtsalary style="TEXT-ALIGN: right" readonly></TD>
				                	<TD CLASS="ctrow01">상여기준금(연봉)</TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="form02" TYPE="Text" size=20 name=txtBonusSalary style="TEXT-ALIGN: right" readonly></TD>
				                </TR>
				                <TR>
				                	<TD CLASS="ctrow01">연장비과세적용구분</TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="form02" TYPE="Text" ALT="세액구분" size=20 name=txttax_cd readonly></TD>
				                	<TD CLASS="ctrow01">은행/계좌번호</TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="form02" TYPE="Text" size=13 NAME="txtBankNm" readonly>
				                						<INPUT CLASS="form02" TYPE="Text" size=20 NAME="txtAccntNo" readonly></TD>
				                </TR>
				                <TR>
				                	<TD CLASS="ctrow01">거주구분</TD>
				                	<TD CLASS="ctrow02">
				                		<INPUT TYPE="RADIO" CLASS="ftgray" NAME="rdoResFlag" VALUE="Y" ID="rdoResFlag1" disabled><LABEL FOR="rdoResFlag1">거주자</LABEL>&nbsp;&nbsp;&nbsp;
								        <INPUT TYPE="RADIO" CLASS="ftgray" NAME="rdoResFlag" VALUE="N" ID="rdoResFlag2" disabled><LABEL FOR="rdoResFlag2">비거주자</LABEL>			
                                    </TD>
				                	<TD CLASS="ctrow01">기자구분</TD>
				                	<TD CLASS="ctrow02">
				                		<INPUT TYPE="RADIO" CLASS="ftgray" NAME="rdoPressFlag" VALUE="Y" ID="rdoPressFlag1" disabled><LABEL FOR="rdoPressFlag1">기자</LABEL>&nbsp;&nbsp;&nbsp;
								        <INPUT TYPE="RADIO" CLASS="ftgray" NAME="rdoPressFlag" VALUE="N" ID="rdoPressFlag2" disabled><LABEL FOR="rdoPressFlag2">비기자</LABEL>							                	
                                    </TD>
				                </TR>
				                <TR>
				                	<TD CLASS="ctrow01">국외근로자구분</TD>
				                	<TD CLASS="ctrow02">
				                		<INPUT TYPE="RADIO" CLASS="ftgray" NAME="rdoOverseaFlag" VALUE="Y" ID="rdoOverseaFlag1" disabled><LABEL FOR="rWWSdoOverseaFlag1">국외근로자</LABEL>&nbsp;&nbsp;&nbsp;
								        <INPUT TYPE="RADIO" CLASS="ftgray" NAME="rdoOverseaFlag" VALUE="N" ID="rdoOverseaFlag2" disabled><LABEL FOR="rdoOverseaFlag2">국내근로자</LABEL>
                                    </TD>
				                	<TD CLASS="ctrow01">노조구분</TD>
				                	<TD CLASS="ctrow02">
				                		<INPUT TYPE="RADIO" CLASS="ftgray" NAME="rdoUnionFlag" VALUE="Y" ID="rdoUnionFlag1" disabled><LABEL FOR="rdoUnionFlag1" >노조원</LABEL>&nbsp;&nbsp;&nbsp;
								        <INPUT TYPE="RADIO" CLASS="ftgray" NAME="rdoUnionFlag" VALUE="N" ID="rdoUnionFlag2" disabled><LABEL FOR="rdoUnionFlag2" >비노조원</LABEL>
				                	</TD>
				                </TR>
				                <TR>
				                	<TD CLASS="ctrow01"></TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="ftgray" TYPE=CHECKBOX NAME="chkPayFlg" readonly disabled>
				                	                    임금지급대상여부</TD>
				                	<TD CLASS="ctrow01"></TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="ftgray" TYPE=CHECKBOX NAME="chkEmpInsurFlg" readonly disabled>
				                	                    고용보험여부</TD>
				                </TR>
				                <TR>
				                	<TD CLASS="ctrow01"></TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="ftgray" TYPE=CHECKBOX NAME="chkYearFlg" readonly disabled>
				                	                     연월차지급대상</TD>
				                	<TD CLASS="ctrow01"></TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="ftgray" TYPE=CHECKBOX NAME="chkRetireFlg" readonly disabled>
				                	                    퇴직금지급대상</TD>
				                </TR>
				                <TR>
				                	<TD CLASS="ctrow01"></TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="ftgray" TYPE=CHECKBOX NAME="chkTaxFlg" readonly disabled>
				                	                    세액계산대상</TD>
				                	<TD CLASS="ctrow01"></TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="ftgray" TYPE=CHECKBOX NAME="chkYearTaxFlg" readonly disabled>
				                	                        연말정산신고대상</TD>
				                </TR>
				                <TR>
				                	<TD CLASS="ctrow01">부양자(노)</TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="form02" TYPE="Text" size=20 name=txtOld style="TEXT-ALIGN: right" readonly></TD>
				                	<TD CLASS="ctrow01">부양자(소)</TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="form02" TYPE="Text" size=20 name=txtYoung style="TEXT-ALIGN: right" readonly></TD>
				                </TR>
				                <TR>
				                	<TD CLASS="ctrow01">장애자</TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="form02" TYPE="Text" size=20 name=txtParia style="TEXT-ALIGN: right" readonly></TD>
				                	<TD CLASS="ctrow01">경로자</TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="form02" TYPE="Text" size=20 name=txtOldCnt style="TEXT-ALIGN: right" readonly></TD>
				                </TR>
				                <TR>
				                	<TD CLASS="ctrow01">자녀양육수</TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="form02" TYPE="Text" size=20 name=txtChild style="TEXT-ALIGN: right" readonly></TD>
				                	<TD CLASS="ctrow01"></TD>
				                	<TD CLASS="ctrow02"><INPUT CLASS="ftgray" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" NAME="chkSpouseFlg" disabled>배우자
				                	                    <INPUT CLASS="ftgray" TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid; BACKGROUND-COLOR: #E9EDF9" NAME="chkLadyFlg" disabled>부녀자</TD>
				                </TR>
                           </table>
                        </td>
                    </TR>
                </TABLE>
            </TD>
        </TR>
        <TR>
            <TD height=10></TD>
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
