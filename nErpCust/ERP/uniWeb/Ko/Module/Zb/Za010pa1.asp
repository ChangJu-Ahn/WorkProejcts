
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za010pa1
*  4. Program Name         : Audit Info Detail Query Popup
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000.03.13
*  8. Modified date(Last)  : 2002.06.24
*  9. Modifier (First)     : LeeJaeJoon
* 10. Modifier (Last)      : LeeJaeWan
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
<!-- #Include file="../../inc/lgvariables.inc" -->    

DIm arrParent
Dim PopupParent
DIm arrParam
DIm arrReturn
DIm gintDataCnt

DIm IsOpenPop        

'=========================================================================================================
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam = arrParent(1)
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
    
    chkTrans1.checked = arrParam(6)
    chkTrans2.checked = arrParam(7)
    chkTrans3.checked = arrParam(8)
    
    txtUser.value = arrParam(9)
    txtUserNm.value = arrParam(10)
    txtTable.value = arrParam(11)
    txtTableNm.value = arrParam(12)
    
    Call rdoClick()
    Call rdoClick2()
End Function

'=========================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE","QA") %>
End Sub

'=========================================================================================================
Sub SetDefaultVal()
    Self.Returnvalue = Array("")
End Sub
'=========================================================================================================
Function OKClick()
    Dim IntRetCD

    If Not chkfield(Document, "1") Then                                    
       Exit Function
    End If

    Redim arrReturn(12)
            
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
    
    arrReturn(6) = chkTrans1.checked
    arrReturn(7) = chkTrans2.checked
    arrReturn(8) = chkTrans3.checked
    
    arrReturn(9) = txtUser.value
    arrReturn(10) = txtUserNm.value
    arrReturn(11) = txtTable.value
    arrReturn(12) = txtTableNm.value
                
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
Sub txtTable_onchange()
    txtTableNm.value = ""
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
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    Set ggoBase = Nothing                                    '☜: Unload Base   Common DLL
    Set ggoOper = Nothing                                    '☜: Unload Scr    Common DLL
    Set ggoSpread = Nothing                                    '☜: Unload Spread Common DLL
End Sub
'=========================================================================================================
Function FncQuery()
    Call OkClick()
End Function

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

    arrParam(0) = "사용자 팝업"                             ' 팝업 명칭 
    arrParam(1) = "z_usr_mast_rec"         ' TABLE 명칭 
    arrParam(2) = Trim(txtUser.value)                        ' Code Condition
    arrParam(3) = ""                                         ' Name Cindition
    arrParam(4) = ""                        ' Where Condition
    arrParam(5) = "사용자"                               ' 조건필드의 라벨 명칭 
     
    arrField(0) = "usr_id"                                 ' Field명(0)
    arrField(1) = "usr_nm"                                 ' Field명(1)
    
    arrHeader(0) = "사용자 ID"                           ' Header명(0)
    arrHeader(1) = "사용자명"                            ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        txtUser.value    = Trim(arrRet(0))
        txtUserNm.value = Trim(arrRet(1))
    End If    
    txtUser.focus
End Function

'=========================================================================================================
'    Name : OpenTable()
'    Description : Table PopUp
'=========================================================================================================
Function OpenTable()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "테이블 팝업"                        ' 팝업 명칭 
    arrParam(1) = "(select distinct table_id from z_audit_policy) a, z_table_info b"        ' TABLE 명칭 
    arrParam(2) = txtTable.value                        ' Code Condition
    arrParam(3) = ""                                    ' Name Cindition

    arrParam(4) = "a.table_id=b.table_id And b.lang_cd= " & FilterVar(PopupParent.gLang, "''", "S") & ""    ' Where Condition
    arrParam(5) = "테이블"                            ' 조건필드의 라벨 명칭 
    
    arrField(0) = "ED24" & PopupParent.gColSep & "a.table_id"                        ' Field명(0)
    arrField(1) = "b.table_nm"                        ' Field명(1)
    
    arrHeader(0) = "테이블 ID"                        ' Header명(0)
    arrHeader(1) = "테이블 명"                        ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=520px; dialogHeight=455px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        txtTable.value = arrRet(0)
        txtTableNm.value = arrRet(1)
    End If    
    txtTable.focus
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
                <TABLE CLASS="basicTB" CELLSPACING=0>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="radio" name=rdoStart ID="rdoStart1" CLASS="RADIO" tag="11" onclick="rdoClick"></TD>
                        <TD CLASS="TD6"><LABEL FOR="rdoStart1">첫 Auditing 정보</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="radio" name=rdoStart ID="rdoStart2" CLASS="RADIO" tag="11" onclick="rdoClick"></TD>
                        <TD CLASS="TD6"><LABEL FOR="rdoStart2">시간</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5">&nbsp;</TD>
                        <TD CLASS="TD6"><script language =javascript src='./js/za010pa1_I574451929_txtStartDate.js'></script></TD>                        
                    </TR>
                    <TR>
                        <TD CLASS="TD5">&nbsp;</TD>
                        <TD CLASS="TD6"><script language =javascript src='./js/za010pa1_OBJECT1_txtStartTime.js'></script></TD>
                    </TR>
                </TABLE>
            </FIELDSET>
        </TD>
        <TD WIDTH=50%>
            <FIELDSET>
                <LEGEND>마지막</LEGEND>
                <TABLE CLASS="basicTB" CELLSPACING=0>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="radio" name=rdoEnd ID="rdoEnd1" CLASS="RADIO" tag="11" onclick="rdoClick2"></TD>
                        <TD CLASS="TD6"><LABEL FOR="rdoEnd1">마지막 Auditing 정보</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5"><INPUT type="radio" name=rdoEnd ID="rdoEnd2" CLASS="RADIO" tag="11" onclick="rdoClick2"></TD>
                        <TD CLASS="TD6"><LABEL FOR="rdoEnd2">시간</LABEL></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5">&nbsp;</TD>
                        <TD CLASS="TD6"><script language =javascript src='./js/za010pa1_I578635263_txtEndDate.js'></script></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5">&nbsp;</TD>
                        <TD CLASS="TD6"><script language =javascript src='./js/za010pa1_OBJECT2_txtEndTime.js'></script></TD>
                    </TR>
                </TABLE>
            </FIELDSET>
        </TD>
    </TR>
    <TR>
        <TD COLSPAN=2>
            <FIELDSET>
                <LEGEND>발생 Transaction</LEGEND>
                <TABLE CLASS="basicTB" CELLSPACING=0>
                    <TR>
                        <TD CLASS="TD5" STYLE="WIDTH:8%"><INPUT type="checkbox" name=chkTrans1 ID="chkTrans1" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6" STYLE="WIDTH:25%"><LABEL FOR="chkTrans1">입력</LABEL></TD>
                        <TD CLASS="TD5" STYLE="WIDTH:8%"><INPUT type="checkbox" name=chkTrans2 ID="chkTrans2" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6" STYLE="WIDTH:25%"><LABEL FOR="chkTrans2">수정</LABEL></TD>
                        <TD CLASS="TD5" STYLE="WIDTH:8%"><INPUT type="checkbox" name=chkTrans3 ID="chkTrans3" CLASS="Check" tag="11"></TD>
                        <TD CLASS="TD6" STYLE="WIDTH:25%"><LABEL FOR="chkTrans3">삭제</LABEL></TD>
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
            <TABLE CLASS="basicTB" CELLSPACING=0>
                <TR>
                    <TD CLASS="TD5">사용자 ID</TD>
                    <TD CLASS="TD6">
                        <INPUT TYPE="Text" NAME=txtUser SIZE=15 MAXLENGTH=13 tag="11" ALT="사용자 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUser" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenUser">
                        <INPUT TYPE="Text" NAME=txtUserNm SIZE=20 tag="14X">
                    </TD>
                </TR>
                <TR>
                    <TD CLASS="TD5">테이블 ID</TD>
                    <TD CLASS="TD6">
                        <INPUT TYPE="Text" NAME=txtTable SIZE=15 MAXLENGTH=30 tag="12" ALT="테이블 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTable" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenTable">
                        <INPUT TYPE="Text" NAME=txtTableNm SIZE=20 tag="14X">
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD HEIGHT=30 COLSPAN=2>
            <TABLE CLASS="basicTB" CELLSPACING=0>
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
