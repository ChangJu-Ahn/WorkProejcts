<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Generated Authority Result Query
*  2. Function Name        : 
*  3. Program ID           :
*  4. Program Name         :
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 
*  8. Modified date(Last)  : 2001/07/15
*  9. Modifier (First)     : Park Sang Hoon
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<Script Language="JavaScript" SRC="../../inc/incImage.js"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                                    

'=========================================================================================================
Const BIZ_PGM_ID = "Za014rb2.asp"                                                
'=========================================================================================================
Dim C_MnuId
Dim C_MnuNm
Dim C_MnuType
Dim C_ActnId
Dim C_UpperMnu

Dim IsOpenPop
Dim arrParent
Dim PopupParent

<!-- #Include file="../../inc/lgvariables.inc" -->    


arrParent   = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

Sub InitSpreadPosVariables()
    C_MnuId = 1
    C_MnuNm = 2
    C_MnuType = 3
    C_ActnId = 4
    C_UpperMnu = 5
End Sub    

'=========================================================================================================
 Sub InitVariables()
    lgIntFlgMode = PopupParent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
    lgLngCurRows = 0                            
    lgStrPrevKey = ""                               
End Sub

'=========================================================================================================
sub SetDefaultVal()
End Sub

'=========================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE","QA") %>
End Sub

'=========================================================================================================
Sub InitSpreadSheet()

    Call InitSpreadPosVariables()
    
    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        Call ggoSpread.Spreadinit("V20021124",,PopupParent.gAllowDragDropSpread)
    
        .ReDraw = false
        .MaxCols = C_UpperMnu +1                                                
        .MaxRows = 0

         Call GetSpreadColumnPos("A")        
       
        ggoSpread.SSSetEdit C_MnuId, "메뉴ID", 25
        ggoSpread.SSSetEdit C_MnuNm, "메뉴명", 30
        ggoSpread.SSSetEdit C_MnuType, "메뉴유형", 18, 2
        ggoSpread.SSSetEdit C_ActnId, "Action ID", 18, 2
        ggoSpread.SSSetEdit C_UpperMnu, "상위메뉴 ID", 25
        .ReDraw = true
        
        Call SetSpreadLock
        
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_MnuId, -1
    ggoSpread.SpreadLock C_MnuNm, -1
    ggoSpread.SpreadLock C_MnuType, -1
    ggoSpread.SpreadLock C_ActnId, -1
    ggoSpread.SpreadLock C_UpperMnu, -1
    .vspdData.ReDraw = True

    End With
End Sub

'=========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
            C_MnuId    =  iCurColumnPos(1)
            C_MnuNm      =  iCurColumnPos(2)
            C_MnuType  =  iCurColumnPos(3)
            C_ActnId      =  iCurColumnPos(4)
            C_UpperMnu      =  iCurColumnPos(5)
            
    End Select
End Sub

'=========================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> UC_PROTECTED Then
              Frm1.vspdData.Action = 0 
              Exit For
           End If
           
       Next
          
    End If   
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

    arrParam(0) = "사용자 팝업"                    ' 팝업 명칭 
    arrParam(1) = "z_usr_mast_rec"                    ' TABLE 명칭 
    arrParam(2) = Trim(frm1.txtUsr.value)            ' Code Condition
    arrParam(3) = ""                                ' Name Cindition
    arrParam(4) = ""                                ' Where Condition
    arrParam(5) = "사용자"                        ' 조건필드의 라벨 명칭 
    
    arrField(0) = "usr_id"                            ' Field명(0)
    arrField(1) = "usr_nm"                            ' Field명(1)
    
    arrHeader(0) = "사용자 ID"                    ' Header명(0)
    arrHeader(1) = "사용자 명"                    ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetUser(arrRet)
    End If    
	frm1.txtUsr.focus     
End Function

'=========================================================================================================
'    Name : OpenMnuInfo()
'    Description : Menu PopUp
'=========================================================================================================
Function OpenMnuInfo()

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "메뉴 팝업"                    ' 팝업 명칭 
    arrParam(1) = "Z_LANG_CO_MAST_MNU"                ' TABLE 명칭 
    arrParam(2) = Trim(frm1.txtMnuID.value)
    arrParam(3) = ""                                ' Name Cindition
    arrParam(4) = "LANG_CD =  " & FilterVar(PopupParent.gLang, "''", "S") & ""        ' Where Condition
    arrParam(5) = "메뉴ID"

    arrField(0) = "MNU_ID"                          ' Field명(0)
    arrField(1) = "MNU_NM"                          ' Field명(1)
    
    arrHeader(0) = "메뉴ID"                     ' Header명(0)
    arrHeader(1) = "메뉴명"                        ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetMnuInfo(arrRet)
    End If    
	frm1.txtMnuID.focus
End Function


'=========================================================================================================
'    Name : SetLangCD()
'    Description : 
'=========================================================================================================
Function SetUser(Byval arrRet)
    frm1.txtUsr.value = arrRet(0)
    frm1.txtUsrNm.value = arrRet(1)
End Function

'=========================================================================================================
'    Name : SetMnuInfo()
'    Description : 
'=========================================================================================================
Function SetMnuInfo(Byval arrRet)
    frm1.txtMnuID.Value = arrRet(0)
    frm1.txtMnuNm.Value = arrRet(1)
End Function

'=========================================================================================================
sub Form_Load()
    Dim strUsrId

    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")    
    Call ggoOper.LockField(Document, "N")                                   
    
    Call InitSpreadSheet                                                    
    Call InitVariables                                                      
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call CookiePage(0)

    strUsrId = arrParent(1)

    If Trim(strUsrId) = "" Then
        frm1.txtUsr.focus 
    Else
       frm1.txtUsr.value = strUsrId
       Call FncQuery()
    End if       

End Sub
'=========================================================================================================
sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'=========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
    
    gMouseClickStatus = "SPC"   
    
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    
       Exit Sub
       End If
           
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'=========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)                
    If Row <= 0 Then
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
    
End Sub

'=========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)        
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'=========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub
'=========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'=========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If CheckRunningBizProcess = True Then
       Exit Sub
    End If

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)    And Not(lgStrPrevKey = "") Then
       If DbQuery = False Then                                                       
          Exit Sub
        End If
    End if
    
End Sub

'=========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'=========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
    Call InitData()
End Sub

'=========================================================================================================
Function FncQuery()

    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               


    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900013", PopupParent.VB_YES_NO,"x","x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    

    frm1.txtUsrNm.value = ""
    Call ggoOper.ClearField(Document, "2")                                        
    Call ggoSpread.ClearSpreadData        
    Call InitVariables
                                                                
    

    If Not chkField(Document, "1") Then                                    
       Exit Function
    End If
    

    If DBQuery = False Then                                                        
        Exit Function 
    End If     
           
    FncQuery = True                                                                
    
End Function

'=========================================================================================================
Function FncExcel() 
    Call parent.FncExport(PopupParent.C_MULTI)                                                
End Function

'=========================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   
End Function

'=========================================================================================================
Function FncFind() 
    Call parent.FncFind(PopupParent.C_MULTI, False)                                         
End Function

'=========================================================================================================
Function FncExit()
    Dim IntRetCD
    FncExit = False
    If lgBlnFlgChgValue = True Then
        IntRetCD = PopupDisplayMsgBox("900016", PopupParent.VB_YES_NO,"x","x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    FncExit = True
End Function


'=========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'=========================================================================================================
Function DbQuery() 
    Dim strVal
    Dim strType

    DbQuery = False
    Call LayerShowHide(1)
    
    
    With frm1
        If .rdoConf(0).checked then strType = "M"
        If .rdoConf(1).checked then strType = "P"
        If .rdoConf(2).checked then    strType = ""

        If lgIntFlgMode = PopupParent.OPMD_UMODE Then                                                                    
            strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001                            
            strVal = strVal & "&txtUsr=" & Trim(.htxtUsr.value)                
            strVal = strVal & "&txtMnuID=" & Trim(.htxtMnuID.value)
            strVal = strVal & "&txtMnuType=" & strType
            strVal = strVal & "&txtLangCD=" & PopupParent.gLang
            strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        Else
            strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001                            
            strVal = strVal & "&txtUsr=" & Trim(.txtUsr.value)                
            strVal = strVal & "&txtMnuID=" & Trim(.txtMnuID.value)
            strVal = strVal & "&txtMnuType=" & strType
            strVal = strVal & "&txtLangCD=" & PopupParent.gLang        
            strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey            
        End If

        Call RunMyBizASP(MyBizASP, strVal)                                        

    End With
    
    DbQuery = True
    

End Function
'=========================================================================================================
Function DbQueryOk()                                                        
    

    lgIntFlgMode = PopupParent.OPMD_UMODE                                                
    
    Call ggoOper.LockField(Document, "Q")                                    
    Frm1.vspdData.Focus
End Function

'=========================================================================================================
Function CancelClick()
    Self.Close()
End Function

'=========================================================================================================
Function CookiePage(Byval Kubun)

    Dim strTemp

    If Kubun = 0 Then

        strTemp = ReadCookie("UsrAuth")
        
        If strTemp = "" then Exit Function

        frm1.txtUsr.value =  strTemp
        
        if Trim(frm1.txtUsr.value) <> "" then
            Call dbquery()
        end if
        
        WriteCookie "UsrAuth" , ""
    End If          
        
End Function



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->    
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
    <TR>
        <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
    </TR>
    <TR>
        <TD HEIGHT=20 WIDTH=100%>
            <FIELDSET CLASS="CLSFLD">        
                <TABLE <%=LR_SPACE_TYPE_40%>>
                    <TR>
                        <TD CLASS="TD5">사용자</TD>
                        <TD CLASS="TD656">
                            <INPUT CLASS="clstxt" TYPE=TEXT NAME="txtUsr" SIZE=13 MAXLENGTH=13 tag="12" ALT="사용자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUsr" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenUser()">
                            <INPUT TYPE=TEXT NAME="txtUsrNm" SIZE=20 tag="14">
                        </TD>
                    </TR>
                    <TR>    
                        <TD CLASS="TD5">메뉴 ID</TD>
                        <TD CLASS="TD656">
                            <INPUT CLASS="clstxt" TYPE=TEXT NAME="txtMnuID" SIZE=15 MAXLENGTH=15 tag="11XXX" ALT="메뉴 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMnuID" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMnuInfo">
                            <INPUT TYPE=TEXT NAME="txtMnuNm" SIZE=40 tag="14">
                        </TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5">메뉴유형</TD>
                        <TD CLASS="TD656">
                            <SPAN STYLE="width:120;"><INPUT TYPE="RADIO" NAME="rdoConf" ID="rdoConf1" CLASS="RADIO" tag="11"><LABEL FOR="rdoConf1">메뉴(M)</LABEL></SPAN>
                             <SPAN STYLE="width:120;"><INPUT TYPE="RADIO" NAME="rdoConf" ID="rdoConf2" CLASS="RADIO" tag="11"><LABEL FOR="rdoConf2">프로그램(P)</LABEL></SPAN>
                             <SPAN STYLE="width:120;"><INPUT TYPE="RADIO" NAME="rdoConf" ID="rdoConf3" CLASS="RADIO" checked tag="11"><LABEL FOR="rdoConf3">전체</LABEL></SPAN>
                         </TD>                
                    </TR>
                </TABLE>
            </FIELDSET>
        </TD>
    </TR>
    <TR>
        <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=* valign=top>
            <TABLE <%=LR_SPACE_TYPE_20%>>
                <TR>
                    <TD HEIGHT="100%">
                        <script language =javascript src='./js/za014ra2_I166482949_vspdData.js'></script>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR> 
    <TR HEIGHT="20">
        <TD WIDTH="100%">
            <TABLE <%=LR_SPACE_TYPE_30%>>
                <TR>
                    <TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
                        <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
                    <TD WIDTH=30% ALIGN=RIGHT>
                        <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>    
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="Zb009mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtUsr" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtMnuID" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
