<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : User Management
*  3. Program ID           : za008ra1
*  4. Program Name         : Audit Management Reference Popup
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
<TITLE>사용자 추가 팝업</TITLE>


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
 Const BIZ_PGM_ID = "Za012RB1.asp"                                                
'=========================================================================================================
 Dim C_UserID
 Dim C_UserNm
 Dim C_UserEngNm
 Dim C_UserValidDt

'=========================================================================================================
Dim arrReturn
Dim arrParent
Dim PopupParent

Dim lgStrPrevDt
Dim lgOrgType
Dim lgOrgCd
Dim IsOpenPop

<!-- #Include file="../../inc/lgvariables.inc" -->    

arrParent   = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName
'=========================================================================================================
Sub InitSpreadPosVariables()
    C_UserID        = 1                                                        
    C_UserNm        = 2                                                        
    C_UserEngNm     = 3                                                        
    C_UserValidDt   = 4                                                        
End Sub    

'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = PopupParent.OPMD_CMODE                   
    lgStrPrevKey = ""                           
    lgLngCurRows = 0                            
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
        .MaxCols = C_UserValidDt + 1                        
        .MaxRows = 0

        Call GetSpreadColumnPos("A")        
       
        ggoSpread.SSSetEdit        C_UserID,    "사용자 ID",    12,,,13
        ggoSpread.SSSetEdit        C_UserNm,    "사용자명",        20,,,30
        ggoSpread.SSSetEdit        C_UserEngNm,    "사용자 영문명", 20,,,30
        ggoSpread.SSSetDate        C_UserValidDt, "사용자 유효기간", 15, 2, PopupParent.gDateFormat   
        .ReDraw = true
        
        Call SetSpreadLock

        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_UserID   , -1, C_UserID
        ggoSpread.SpreadLock C_UserNm   , -1, C_UserNm
        ggoSpread.SpreadLock C_UserEngNm   , -1, C_UserEngNm
        ggoSpread.SpreadLock C_UserValidDt   , -1, C_UserValidDt
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
            
            C_UserID    =  iCurColumnPos(1)
            C_UserNm      =  iCurColumnPos(2)
            C_UserEngNm  =  iCurColumnPos(3)
            C_UserValidDt      =  iCurColumnPos(4)

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
Sub Form_Load()

    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

    Call ggoOper.LockField(Document, "N")                                   
    
    Call InitSpreadSheet                                                    
    Call InitVariables                                                      

    lgOrgType = arrParent(1)(0)
    lgOrgCd = arrParent(1)(1)

    '----------  Coding part  -------------------------------------------------------------   
    Redim arrReturn(0, 0)
    Self.Returnvalue = arrReturn

    frm1.txtUsrID.focus
    Set gActiveElement = document.activeElement
    
    Call Fncquery()
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
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
Function vspdData_DblClick(ByVal Col, ByVal Row)
    If frm1.vspdData.MaxRows > 0 Then
        If frm1.vspdData.ActiveRow = Row and frm1.vspdData.ActiveRow > 0 and Row > 0 Then
            Call OKClick()
        End If
    End If
End Function

'=========================================================================================================
Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

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
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)    And Not(lgStrPrevKey = "" And lgStrPrevDt = "") Then
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
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'=========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               



    Call ggoOper.ClearField(Document, "2")                                        
    Call ggoSpread.ClearSpreadData()    
    Call InitVariables

    If DbQuery = False Then
       Exit Function
    End If
       
    FncQuery = True                                                                
    
End Function

'=========================================================================================================
Function DbQuery() 

    Dim strVal

    DbQuery = False
    
    Call LayerShowHide(1)

    With frm1

    If lgIntFlgMode = PopupParent.OPMD_UMODE Then                                                    
        strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001                            
        strVal = strVal & "&txtUsrID=" & Trim(.hUsrID.value)                
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        strVal = strVal & "&lgOrgType=" & lgOrgType
        strVal = strVal & "&lgOrgCd=" & lgOrgCd
    Else
        strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001                            
        strVal = strVal & "&txtUsrID=" & Trim(.txtUsrID.value)                
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        strVal = strVal & "&lgOrgType=" & lgOrgType
        strVal = strVal & "&lgOrgCd=" & lgOrgCd    
    End If
    

    Call RunMyBizASP(MyBizASP, strVal)                                        
        
    End With
    
    DbQuery = True
    
End Function

'=========================================================================================================
Function DbQueryOk()                                                        
    lgIntFlgMode = PopupParent.OPMD_UMODE                                                
    frm1.vspdData.Focus
End Function


'=========================================================================================================
Function OKClick()
    Dim intColCnt, intRowCnt, intInsRow
    Dim iRowCnt
    
    iRowCnt = frm1.vspdData.SelBlockRow2 - frm1.vspdData.SelBlockRow
    
    if frm1.vspdData.SelBlockRow2 > 0 Then 
        
        intInsRow = 0
        
        Redim arrReturn(iRowCnt, frm1.vspdData.MaxCols - 1)
        
        For intRowCnt = frm1.vspdData.SelBlockRow To frm1.vspdData.SelBlockRow2
            For intColCnt = 0 To frm1.vspdData.MaxCols - 2
                Select case intColCnt
                    case 0
                    	arrReturn(intInsRow, intColCnt) = GetSpreadText(frm1.vspdData, C_UserID, intRowCnt, "X", "X")
                    case 1
                    	arrReturn(intInsRow, intColCnt) = GetSpreadText(frm1.vspdData, C_UserNm, intRowCnt, "X", "X")
                    case 2
                    	arrReturn(intInsRow, intColCnt) = GetSpreadText(frm1.vspdData, C_UserEngNm, intRowCnt, "X", "X")
                    case 3
                    	arrReturn(intInsRow, intColCnt) = GetSpreadText(frm1.vspdData, C_UserValidDt, intRowCnt, "X", "X")
                End Select
            Next
                    
            intInsRow = intInsRow + 1
                    
        Next
    End if            
        
    Self.Returnvalue = arrReturn
    Self.Close()
End Function

'=========================================================================================================
Function CancelClick()
    Self.Close()
End Function
'=========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
                window.document.search.style.cursor = "wait"
            case "POFF"
                window.document.search.style.cursor = ""
      End Select
End Function


'=========================================================================================================
'    Name : OpenUserID()
'    Description : TableID PopUp
'=========================================================================================================
Function OpenUserID(Byval strCode)

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function
    
    IsOpenPop = True

    arrParam(0) = "사용자 ID Popup"                  ' 팝업 명칭 
    arrParam(1) = "z_Usr_mast_rec"                      ' TABLE 명칭 
    arrParam(2) = strCode                              ' Code Condition
    arrParam(3) = ""                                  ' Name Cindition
    arrParam(4) = ""                                  '"LANG_CD='" & PopupParent.gLang & "' and USE_YN='Y'"     Where Condition
    arrParam(5) = "사용자 ID"            
    
    arrField(0) = "Usr_id"                            ' Field명(0)
    arrField(1) = "Usr_nm"                            ' Field명(1)
    
    arrHeader(0) = "사용자 ID"                    ' Header명(0)
    arrHeader(1) = "사용자명"                     ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetUserID(arrRet)
    End If    
	frm1.txtUsrID.focus
End Function

'=========================================================================================================
'    Name : SetUserID()
'    Description : Plant Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetUserID(byval arrRet)
    frm1.txtUsrID.Value = arrRet(0)
    frm1.txtUsrNm.Value = arrRet(1)
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->    
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
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
                    <TD CLASS="TD5">사용자 ID</TD>
                    <TD CLASS="TD656" COLSPAN=3><INPUT TYPE=TEXT NAME="txtUsrID" SIZE=15 MAXLENGTH=13 tag="11"  ALT="사용자 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTableID" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUserID frm1.txtUsrID.value">&nbsp;<INPUT TYPE=TEXT NAME="txtUsrNm" SIZE=25 tag="14N"></TD>
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
                        <script language =javascript src='./js/za012ra1_vspdData_vspdData.js'></script>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <tr>
        <TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <TR HEIGHT="20">
        <TD WIDTH="100%">
            <TABLE <%=LR_SPACE_TYPE_30%>>
                <TR>
                    <TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
                        <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="Fncquery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
                    <TD WIDTH=30% ALIGN=RIGHT>
                        <IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
                        <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="Za008RB1.asp" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hUsrId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

