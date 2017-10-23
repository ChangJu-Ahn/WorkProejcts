
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : User Management
*  3. Program ID           : za004pa1
*  4. Program Name         : Message History Popup
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000.03.13
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : LeeJaeJoon
* 10. Modifier (Last)      : 
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>


<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">        

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/IncCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/IncCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/IncCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/IncCliVariables.vbs"></SCRIPT>
<Script Language="JavaScript" SRC="../../inc/incImage.js"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit    
'=========================================================================================================

Dim arrParent
Dim arrParam                                               '--- Parameter Group
Dim PopupParent
Dim arrReturn                                              '--- Return Parameter Group
Dim gintDataCnt                                            '--- Data Counts to Query

Dim IsOpenPop        

<!-- #Include file="../../inc/lgvariables.inc" -->    

'=========================================================================================================
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam = arrParent(1)
top.document.title = PopupParent.gActivePRAspName

'=========================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE","QA") %>
End Sub

'=========================================================================================================
Function InitVariables()

    txtStartTime.DateTimeFormat = 5
    txtStartTime.UserDefinedFormat = "hh:nn:ss"
    txtStartTime.TimeStyle = 2
    
    txtEndTime.DateTimeFormat = 5
    txtEndTime.UserDefinedFormat = "hh:nn:ss"
    txtEndTime.TimeStyle = 2
        
    rdoStart(arrParam(0)).checked = True
    txtStartDate.text = arrParam(1)
    txtStartTime.text = arrParam(2)
    
    rdoEnd(arrParam(3)).checked = True
    txtEndDate.text = arrParam(4)
    txtEndTime.text = arrParam(5)
    
    chkSeverity1.checked = arrParam(6)
    chkSeverity2.checked = arrParam(7)
    chkSeverity3.checked = arrParam(8)
    chkSeverity4.checked = arrParam(9)
    
    chkMsgType1.checked = arrParam(10)
    chkMsgType2.checked = arrParam(11)
    chkMsgType3.checked = arrParam(12)
    
    txtUser.value = arrParam(13)
    txtUserNm.value = arrParam(14)
    txtMsg.value = arrParam(15)
    txtPgm.value = arrParam(16)
    txtPgmNm.value = arrParam(17)
    txtClient.value = arrParam(18)
    
    Call rdoClick()
    Call rdoClick2()    
    
End Function

'=========================================================================================================
Sub SetDefaultVal()
End Sub

'=========================================================================================================
Function FncQuery()
    Call OKClick()
End Function

'=========================================================================================================
Function OKClick()
    Dim IntRetCD
    Redim arrReturn(18)
            
    If rdoStart(0).checked = True Then
        arrReturn(0) = 0
    Else
        arrReturn(0) = 1
        
        If (UNIConvDate(txtStartDate.text) & " " & txtStartTime.text) > (UNIConvDate(txtEndDate.text) & " " & txtEndTime.text) Then
           IntRetCD = DisplayMsgBox("210404", VBOKONLY,"X","X")
           Exit Function
        End If        
        
    End If
    arrReturn(1) = txtStartDate.text
    arrReturn(2) = txtStartTime.text
    
    If rdoEnd(0).checked = True Then
        arrReturn(3) = 0
    Else
        arrReturn(3) = 1
    End If

    arrReturn(4) = txtEndDate.text
    arrReturn(5) = txtEndTime.text
    
    arrReturn(6) = chkSeverity1.checked
    arrReturn(7) = chkSeverity2.checked
    arrReturn(8) = chkSeverity3.checked
    arrReturn(9) = chkSeverity4.checked

    arrReturn(10) = chkMsgType1.checked
    arrReturn(11) = chkMsgType2.checked
    arrReturn(12) = chkMsgType3.checked
    
    arrReturn(13) = txtUser.value
    arrReturn(14) = txtUserNm.value
    arrReturn(15) = txtMsg.value
    arrReturn(16) = txtPgm.value
    arrReturn(17) = txtPgmNm.value
    arrReturn(18) = txtClient.value
                
    Self.Returnvalue = arrReturn
    Self.Close()
End Function
'=========================================================================================================
Function CancelClick()
    Self.Close()
End Function

'=========================================================================================================
Sub txtUser_onchange()
    txtUserNm.value = ""
End Sub

'=========================================================================================================
Sub txtPgm_onchange()
    txtPgmNm.value = ""
End Sub
'=========================================================================================================
Sub rdoClick()

    If rdoStart(0).checked Then
        txtStartDate.Enabled = false
        txtStartTime.Enabled = false
        'rdoEnd(0).click
    Else
        txtStartDate.Enabled = true
        txtStartTime.Enabled = true
        'rdoEnd(1).click
    End If    
    
End Sub
'=========================================================================================================
Sub rdoClick2()

    If rdoEnd(0).checked Then
        txtEndDate.Enabled = false
        txtEndTime.Enabled = false
        'rdoStart(0).click
    Else
        txtEndDate.Enabled = true
        txtEndTime.Enabled = true
        'rdoStart(1).click
    End If
    
End Sub

'=========================================================================================================
Sub Form_Load()
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)                
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call SetDefaultVal()

    redim arrReturn(0)
    Self.Returnvalue = arrReturn
End Sub

'=========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'=========================================================================================================
'   Event Name : txtStartDate_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=========================================================================================================
Sub txtStartDate_DblClick(Button)
    If Button = 1 Then
        txtStartDate.Action = 7        
    End If
End Sub
'=========================================================================================================
'   Event Name : txtStartDate_Change()
'   Event Desc : 달력을 호출한다.
'=========================================================================================================
Sub txtStartDate_Change()    
    lgBlnFlgChgValue = True
End Sub


'=========================================================================================================
'   Event Name : txtEndDate_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=========================================================================================================
Sub txtEndDate_DblClick(Button)
    If Button = 1 Then
        txtEndDate.Action = 7        
    End If
End Sub

'=========================================================================================================
'   Event Name : txtEndDate_Change()
'   Event Desc : 달력을 호출한다.
'=========================================================================================================
Sub txtEndDate_Change()    
    lgBlnFlgChgValue = True
End Sub

'=========================================================================================================
'    Name : OpenUser()
'    Description : User PopUp
'=========================================================================================================
Function OpenUser()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "사용자 팝업"                        ' 팝업 명칭 
    arrParam(1) = "z_usr_mast_rec "    ' TABLE 명칭 
    arrParam(2) = Trim(txtUser.value)                    ' Code Condition
    arrParam(3) = ""                                    ' Name Cindition
    arrParam(4) = ""                ' Where Condition
    arrParam(5) = "사용자"                        ' 조건필드의 라벨 명칭 
    
    arrField(0) = "usr_id"                        ' Field명(0)
    arrField(1) = "usr_nm"                        ' Field명(1)
    
    arrHeader(0) = "사용자 ID"                    ' Header명(0)
    arrHeader(1) = "사용자 명"                    ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        txtUser.value = arrRet(0)
        txtUserNm.value = arrRet(1)
    End If    
    txtUser.focus 
End Function

'=========================================================================================================
'    Name : OpenMsg()
'    Description : Msg PopUp
'=========================================================================================================
Function OpenMsg()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)
    Dim severityWhereCnt
    Dim MsgTypeWhereCnt
    
    severityWhereCnt = 0
    MsgTypeWhereCnt = 0
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "메세지 팝업"            ' 팝업 명칭 
    arrParam(1) = "b_message"                    ' TABLE 명칭 
    arrParam(2) = Trim(txtMsg.value)            ' Code Condition
    arrParam(3) = ""                            ' Name Cindition
    
    arrParam(4) = ""    ' Where Condition

    If chkSeverity1.checked Then
        arrParam(4) = "(Severity=" & FilterVar("1", "''", "S") & "  "
        severityWhereCnt = severityWhereCnt + 1
    End If

    If chkSeverity2.checked Then
        If severityWhereCnt = 0 Then
            arrParam(4) = arrParam(4) & "("
        Else
            arrParam(4) = arrParam(4) & "Or "
        End If
        arrParam(4) = arrParam(4) & "Severity=" & FilterVar("2", "''", "S") & " "
        severityWhereCnt = severityWhereCnt + 1
    End If
    
    If chkSeverity3.checked Then
        If severityWhereCnt = 0 Then
            arrParam(4) = arrParam(4) & "("
        Else
            arrParam(4) = arrParam(4) & "Or "
        End If
        arrParam(4) = arrParam(4) & "Severity=" & FilterVar("3", "''", "S") & " "
        severityWhereCnt = severityWhereCnt + 1
    End If
    
    If chkSeverity4.checked Then
        If severityWhereCnt = 0 Then
            arrParam(4) = arrParam(4) & "("
        Else
            arrParam(4) = arrParam(4) & "Or "
        End If
        arrParam(4) = arrParam(4) & "Severity=" & FilterVar("4", "''", "S") & " "
        severityWhereCnt = severityWhereCnt + 1
    End If
    
    If severityWhereCnt > 0 Then
        arrParam(4) = arrParam(4) & ") "
    End If

    If chkMsgType1.checked Then
        If severityWhereCnt > 0 Then
            arrParam(4) = arrParam(4) & "And "
        End If
        arrParam(4) = arrParam(4) & "(msg_type=" & FilterVar("A", "''", "S") & "  "
        msgTypeWhereCnt = msgTypeWhereCnt + 1
    End If
    
    If chkMsgType2.checked Then
        If msgTypeWhereCnt = 0 Then
            If severityWhereCnt > 0 Then
                arrParam(4) = arrParam(4) & "And "
            End If
            arrParam(4) = arrParam(4) & "("
        Else
            arrParam(4) = arrParam(4) & "Or "
        End If
        arrParam(4) = arrParam(4) & "msg_type=" & FilterVar("D", "''", "S") & "  "
        msgTypeWhereCnt = msgTypeWhereCnt + 1
    End If
    
    If chkMsgType3.checked Then
        If msgTypeWhereCnt = 0 Then
            If severityWhereCnt > 0 Then
                arrParam(4) = arrParam(4) & "And "
            End If
            arrParam(4) = arrParam(4) & "("
        Else
            arrParam(4) = arrParam(4) & "Or "
        End If
        arrParam(4) = arrParam(4) & "msg_type=" & FilterVar("S", "''", "S") & "  "
        msgTypeWhereCnt = msgTypeWhereCnt + 1
    End If
    
    If msgTypeWhereCnt > 0 Then
        arrParam(4) = arrParam(4) & ") and lang_cd =  " & FilterVar(PopupParent.gLang , "''", "S") & ""
    End If
    
    If severityWhereCnt = 0 And msgTypeWhereCnt = 0 Then
        arrParam(4) = "severityWhereCnt=" & FilterVar("x", "''", "S") & " And msgTypeWhereCnt=" & FilterVar("x", "''", "S") & ""
    End If

    arrParam(5) = "메세지"                        ' 조건필드의 라벨 명칭 
    
    arrField(0) = "msg_cd"                        ' Field명(0)
    arrField(1) = "msg_text"                    ' Field명(1)

    arrHeader(0) = "메세지 코드"                ' Header명(0)
    arrHeader(1) = "메세지"                        ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        txtMsg.value = arrRet(0)
    End If    
    txtMsg.focus 
End Function

'=========================================================================================================
'    Name : OpenPgm()
'    Description : Pgm PopUp
'=========================================================================================================
Function OpenPgm()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "프로그램 팝업"            ' 팝업 명칭 
    arrParam(1) = "z_co_mast_mnu a, z_lang_co_mast_mnu b"                ' TABLE 명칭 
    arrParam(2) = Trim(txtPgm.value)            ' Code Condition
    arrParam(3) = ""                            ' Name Cindition
    arrParam(4) = "a.mnu_id = b.mnu_id and a.mnu_type = " & FilterVar("P", "''", "S") & "  and b.lang_cd =  " & FilterVar(PopupParent.gLang , "''", "S") & "" ' Where Condition
    arrParam(5) = "프로그램"                    ' 조건필드의 라벨 명칭 
    
    arrField(0) = "a.mnu_id"                    ' Field명(0)
    arrField(1) = "b.mnu_nm"                    ' Field명(1)
    
    arrHeader(0) = "프로그램 ID"                ' Header명(0)
    arrHeader(1) = "프로그램 명"                ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        txtPgm.value = arrRet(0)
        txtPgmNm.value = arrRet(1)
    End If
    txtPgm.focus 
End Function

'=========================================================================================================
'    Name : OpenClient()
'    Description : Client PopUp
'=========================================================================================================
Function OpenClient()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)
    Dim whereCnt
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "클라이언트 팝업"                ' 팝업 명칭 
    arrParam(1) = "( select client_ip, client_nm from z_msg_logging group by client_ip, client_nm ) a"                    ' TABLE 명칭 
    arrParam(2) = Trim(txtClient.value)                ' Code Condition
    arrParam(3) = ""                                ' Name Cindition
    arrParam(4) = ""                                ' Where Condition

    arrParam(5) = "클라이언트"                    ' 조건필드의 라벨 명칭 
    
    arrField(0) = "a.client_nm"                        ' Field명(0)
    arrField(1) = "a.client_ip"                        ' Field명(1)
    
    arrHeader(0) = "클라이언트명"                ' Header명(0)
    arrHeader(1) = "클라이언트 IP"                ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        txtClient.value = arrRet(0)
    End If    
    txtClient.focus 
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->    
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
    <TR>
        <TD WIDTH=50%>
            <FIELDSET>
                <LEGEND>시작</LEGEND>
                <TABLE CLASS="basicTB" CELLSPACING=0 CELLPADDING=0>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="radio" name=rdoStart ID="rdoStart1" CLASS="Radio" tag="11" onclick="rdoClick"></TD>
                        <TD CLASS="TD6"><LABEL FOR="rdoStart1">첫 메세지</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="radio" name=rdoStart ID="rdoStart2" CLASS="Radio" tag="11" onclick="rdoClick"></TD>
                        <TD CLASS="TD6"><LABEL FOR="rdoStart2">시간</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5">&nbsp;</TD>
                        <TD CLASS="TD6"><script language =javascript src='./js/za004pa1_I238862198_txtStartDate.js'></script></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5">&nbsp;</TD>
                        <TD CLASS="TD6"><script language =javascript src='./js/za004pa1_OBJECT1_txtStartTime.js'></script></TD>
                    </TR>
                </TABLE>
            </FIELDSET>
        </TD>
        <TD WIDTH=50%>
            <FIELDSET>
                <LEGEND>마지막</LEGEND>
                <TABLE CLASS="basicTB" CELLSPACING=0 CELLPADDING=0>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="radio" name=rdoEnd ID="rdoEnd1" CLASS="Radio" tag="11" onclick="rdoClick2"></TD>
                        <TD CLASS="TD6"><LABEL FOR="rdoEnd1">마지막 메세지</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="radio" name=rdoEnd ID="rdoEnd2" CLASS="Radio" tag="11" onclick="rdoClick2"></TD>
                        <TD CLASS="TD6"><LABEL FOR="rdoEnd2">시간</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5">&nbsp;</TD>
                        <TD CLASS="TD6"><script language =javascript src='./js/za004pa1_I365152990_txtEndDate.js'></script></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5">&nbsp;</TD>
                        <TD CLASS="TD6"><script language =javascript src='./js/za004pa1_OBJECT2_txtEndTime.js'></script></TD>                        
                    </TR>
                </TABLE>
            </FIELDSET>
        </TD>
    </TR>
    <TR>
        <TD COLSPAN=2>
            <FIELDSET>
                <LEGEND>Severity</LEGEND>
                <TABLE CLASS="basicTB" CELLSPACING=0 CELLPADDING=0>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="checkbox" name=chkSeverity1 ID="chkSeverity1" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6"><LABEL FOR="chkSeverity1">Information</LABEL></TD>
                        <TD CLASS="TD5"><INPUT type="checkbox" name=chkSeverity2 ID="chkSeverity2" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6"><LABEL FOR="chkSeverity2">Warning</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="checkbox" name=chkSeverity3 ID="chkSeverity3" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6"><LABEL FOR="chkSeverity3">Error</LABEL></TD>
                        <TD CLASS="TD5"><INPUT type="checkbox" name=chkSeverity4 ID="chkSeverity4" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6"><LABEL FOR="chkSeverity4">Fatal Error</LABEL></TD>
                    </TR>
                </TABLE>
            </FIELDSET>
        </TD>
    </TR>
    <TR>
        <TD COLSPAN=2>
            <FIELDSET>
                <LEGEND>메세지 유형</LEGEND>
                <TABLE CLASS="basicTB" CELLSPACING=0 CELLPADDING=0>
                    <TR>
                        <TD CLASS="TD5" STYLE="WIDTH:8%"><INPUT type="checkbox" name=chkMsgType1 ID="chkMsgType1" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6" STYLE="WIDTH:25%"><LABEL FOR="chkMsgType1">Application</LABEL></TD>
                        <TD CLASS="TD5" STYLE="WIDTH:8%"><INPUT type="checkbox" name=chkMsgType2 ID="chkMsgType2" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6" STYLE="WIDTH:25%"><LABEL FOR="chkMsgType2">DBMS</LABEL></TD>
                        <TD CLASS="TD5" STYLE="WIDTH:8%"><INPUT type="checkbox" name=chkMsgType3 ID="chkMsgType3" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6" STYLE="WIDTH:25%"><LABEL FOR="chkMsgType3">System(OS)</LABEL></TD>
                    </TR>
                </TABLE>
            </FIELDSET>
        </TD>
    </TR>
    <TR>
        <TD>&nbsp;</TD>
    </TR>
    <TR>
        <TD COLSPAN=2>
            <TABLE CLASS="basicTB" CELLSPACING=0 CELLPADDING=0>
                <TR>
                    <TD CLASS="TD5">사용자 ID</TD>
                    <TD CLASS="TD6">
                        <INPUT TYPE="Text" NAME=txtUser SIZE=15 MAXLENGTH=13 tag="11" ALT="사용자 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUser" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenUser">
                        <INPUT TYPE="Text" NAME=txtUserNm SIZE=20 tag="14X">
                    </TD>
                </TR>
                <TR>
                    <TD CLASS="TD5">메세지 코드</TD>
                    <TD CLASS="TD6"><INPUT TYPE="Text" NAME=txtMsg SIZE=15 MAXLENGTH=10 tag="11" ALT="메세지 코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMsg" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMsg"></TD>
                </TR>
                <TR>
                    <TD CLASS="TD5">프로그램 ID</TD>
                    <TD CLASS="TD6">
                        <INPUT TYPE="Text" NAME=txtPgm SIZE=15 MAXLENGTH=15 tag="11" ALT="프로그램 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPgm" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPgm">
                        <INPUT TYPE="Text" NAME=txtPgmNm SIZE=20 tag="14X">
                    </TD>
                </TR>
                <TR>
                    <TD CLASS="TD5">클라이언트명</TD>
                    <TD CLASS="TD6"><INPUT TYPE="Text" NAME=txtClient SIZE=25 MAXLENGTH=20 tag="11" ALT="클라이언트명"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnClient" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenClient"></TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD HEIGHT=30 COLSPAN=2>
            <TABLE CLASS="basicTB">
                <TR>
                    <TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
                    <TD WIDTH=30% ALIGN=RIGHT>
                        <IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
                        <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
</TABLE>
</BODY>
</HTML>

