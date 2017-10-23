<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Basis Architect                                                            *
*  2. Function Name        : User Management                                                            *
*  3. Program ID           : ZA003PA1                                                                    *
*  4. Program Name         : Login History Popup 1                                                        *
*  5. Program Desc         :                                                                            *
*  7. Modified date(First) :                                                                             *
*  8. Modified date(Last)  :                                                                            *
*  9. Modifier (First)     : Park Sang Hoon                                                                *
* 10. Modifier (Last)      : Park Sang Hoon                                                                *
* 11. Comment              :                                                                            *
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE>로그인 내역 상세조회 팝업</TITLE>


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
Dim lgBlnFlgChgValue
Dim arrParent
Dim arrParam                    <% '--- Parameter Group %>
Dim PopupParent
DIm arrReturn                <% '--- Return Parameter Group %>
Dim gintDataCnt                <% '--- Data Counts to Query %>

Dim IsOpenPop        

'=========================================================================================================
arrParent   = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam    = arrParent(1)
top.document.title = PopupParent.gActivePRAspName

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
    
    chkStatus1.checked = arrParam(6)
    chkStatus2.checked = arrParam(7)
    chkStatus3.checked = arrParam(8)
    chkStatus4.checked = arrParam(9)
    chkStatus5.checked = arrParam(10)    
    
    txtUser.value = arrParam(11)
    txtUserNm.value = arrParam(12)
    txtClient.value = arrParam(13)
    
    Call rdoClick()
    Call rdoClick2()
    
End Function

'=========================================================================================================
Sub SetDefaultVal()
'    Self.Returnvalue = Array("")
End Sub

'=========================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE","QA") %>
End Sub

Function FncQuery()
    Call OKClick()
End Function

'=========================================================================================================
Function OKClick()
    Dim IntRetCD
    Redim arrReturn(13)
            
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
    
    arrReturn(6) = chkStatus1.checked
    arrReturn(7) = chkStatus2.checked
    arrReturn(8) = chkStatus3.checked
    arrReturn(9) = chkStatus4.checked
    arrReturn(10) = chkStatus5.checked    
    
    arrReturn(11) = txtUser.value
    arrReturn(12) = txtUserNm.value
    arrReturn(13) = txtClient.value
                
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
Sub rdoClick()

    If rdoStart(0).checked Then
        txtStartDate.Enabled = false
        txtStartTime.Enabled = false        
    Else
        txtStartDate.Enabled = true
        txtStartTime.Enabled = true        
    End If    
    
End Sub

'=========================================================================================================
Sub rdoClick2()

    If rdoEnd(0).checked Then
        txtEndDate.Enabled = false
        txtEndTime.Enabled = false        
    Else
        txtEndDate.Enabled = true
        txtEndTime.Enabled = true        
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
Sub Form_QueryUnload(Cancel , UnloadMode )
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
    arrParam(4) = ""                    ' Where Condition
    arrParam(5) = "사용자"                            ' 조건필드의 라벨 명칭 

    arrField(0) = "usr_id"                            ' Field명(0)
    arrField(1) = "usr_nm"                            ' Field명(1)
    
    arrHeader(0) = "사용자 ID"                        ' Header명(0)
    arrHeader(1) = "사용자 명"                        ' Header명(1)

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
    arrParam(1) = "(select client_id, client_ip from z_log_in_hstry group by client_id, client_ip) a"                    ' TABLE 명칭 
    arrParam(2) = Trim(txtClient.value)                ' Code Condition
    arrParam(3) = ""                                ' Name Cindition
    arrParam(4) = ""                                ' Where Condition

    arrParam(5) = "클라이언트"                    ' 조건필드의 라벨 명칭 
    
    arrField(0) = "a.client_id"                        ' Field명(0)
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

<BODY SCROLL=no TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
    <TR>
        <TD WIDTH="50%">
            <FIELDSET>
                <LEGEND>시작</LEGEND>
                <TABLE CLASS="basicTB" CELLSPACING=0 CELLPADDING=0>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="radio" name=rdoStart ID="rdoStart1" CLASS="Radio" tag="11" onclick="rdoClick"></TD>
                        <TD CLASS="TD6"><LABEL FOR="rdoStart1">첫 로그온 정보</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="radio" name=rdoStart ID="rdoStart2" CLASS="Radio" tag="11" onclick="rdoClick"></TD>
                        <TD CLASS="TD6"><LABEL FOR="rdoStart2">시간</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5">&nbsp;</TD>                        
                        <TD CLASS="TD6"><script language =javascript src='./js/za003pa1_I536928039_txtStartDate.js'></script></TD>                        
                    </TR>
                    <TR>
                        <TD CLASS="TD5">&nbsp;</TD>
                        <TD CLASS="TD6"><script language =javascript src='./js/za003pa1_OBJECT1_txtStartTime.js'></script></TD>                                                
                        <!--TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtStartTime" SIZE=10 MAXLENGTH=8 STYLE="TEXT-ALIGN: center" tag="11"></TD-->
                    </TR>
                </TABLE>
            </FIELDSET>
        </TD>
        <TD WIDTH=50%>
            <FIELDSET>
                <LEGEND>마지막</LEGEND>
                <TABLE CLASS="basicTB" CELLSPACING=0 CELLPADDING=0>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="radio" name=rdoEnd ID="rdoEnd1" CLASS="Radio" onclick="rdoClick2" tag="11"></TD>
                        <TD CLASS="TD6"><LABEL FOR="rdoEnd1">마지막 로그온 정보</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="radio" name=rdoEnd ID="rdoEnd2" CLASS="Radio" onclick="rdoClick2" tag="11"></TD>
                        <TD CLASS="TD6"><LABEL FOR="rdoEnd2">시간</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5">&nbsp;</TD>
                        <TD CLASS="TD6"><script language =javascript src='./js/za003pa1_I172212791_txtEndDate.js'></script></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5">&nbsp;</TD>
                        <TD CLASS="TD6"><script language =javascript src='./js/za003pa1_OBJECT2_txtEndTime.js'></script></TD>                        
                    </TR>
                </TABLE>
            </FIELDSET>
        </TD>
    </TR>
    <TR>
        <TD COLSPAN=2>
            <FIELDSET>
                <LEGEND>상태</LEGEND>
                <TABLE CLASS="basicTB" CELLSPACING=0 CELLPADDING=0>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="checkbox" name=chkStatus1 ID="chkStatus1" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6"><LABEL FOR="chkStatus1">정상 종료</LABEL></TD>
                        <TD CLASS="TD5"><INPUT type="checkbox" name=chkStatus2 ID="chkStatus2" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6"><LABEL FOR="chkStatus2">비정상 종료</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="checkbox" name=chkStatus3 ID="chkStatus3" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6"><LABEL FOR="chkStatus3">접속 중</LABEL></TD>
                        <TD CLASS="TD5"><INPUT type="checkbox" name=chkStatus4 ID="chkStatus4" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6"><LABEL FOR="chkStatus4">Login Locking</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="checkbox" name=chkStatus5 ID="chkStatus5" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6"><LABEL FOR="chkStatus5">Locking 해제</LABEL></TD>
                        <TD CLASS="TD5"></TD>
                        <TD CLASS="TD6"></TD>
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
                    <TD CLASS="TD5">사용자</TD>
                    <TD CLASS="TD6">
                        <INPUT NAME=txtUser SIZE=13 MAXLENGTH="13" tag="11" ALT="사용자"><IMG onclick=vbscript:OpenUser src="../../../CShared/image/btnPopup.gif" align=top name=btnUser  TYPE="BUTTON">
                        <INPUT NAME=txtUserNm tag="14X" >
                    </TD>
                </TR>
                <TR>
                    <TD CLASS="TD5">클라이언트명</TD>
                    <TD CLASS="TD6"><INPUT NAME=txtClient SIZE=25 MAXLENGTH="20" tag="11" ALT="클라이언트명"><IMG onclick=vbscript:OpenClient src="../../../CShared/image/btnPopup.gif" align=top name=btnClient  TYPE="BUTTON"></TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD HEIGHT=30 COLSPAN=2>
            <TABLE CLASS="basicTB">
                <TR>
                    <TD WIDTH="70%" NOWRAP>&nbsp;&nbsp;
                    <TD WIDTH="30%" ALIGN=right>
                        <IMG onmouseover="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)" style="CURSOR: hand" onclick=OkClick() onmouseout=javascript:MM_swapImgRestore() alt=OK src="../../../CShared/image/ok_d.gif" name=pop1></IMG>&nbsp;&nbsp;
                        <IMG onmouseover="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" style="CURSOR: hand" onclick=CancelClick() onmouseout=javascript:MM_swapImgRestore() alt=CANCEL src="../../../CShared/image/cancel_d.gif" name=pop2></IMG>&nbsp;&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
</TABLE>
</BODY>
</HTML>

