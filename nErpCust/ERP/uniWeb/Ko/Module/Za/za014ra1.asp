
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : User Management
*  3. Program ID           : za014ra1
*  4. Program Name         : User-Organization Map
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2002.12.03
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : ParkSangHoon
* 10. Modifier (Last)      : 
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE>사용자별 조직변경 내역 팝업</TITLE>


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
Const BIZ_PGM_ID = "Za014rb1.asp"            
'=========================================================================================================
Dim C_OccurDt
Dim C_UseYn
Dim C_OrgTypeNm
Dim C_OrgCd
Dim C_OrgNm

<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim arrReturn
Dim arrParent
Dim PopupParent

Dim lgStrPrevOccurDt
Dim lgStrPrevOrgType
Dim lgStrPrevOrgCd
'=========================================================================================================
Dim IsOpenPop    

arrParent   = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName
'=========================================================================================================
Sub InitSpreadPosVariables()
    C_OccurDt   = 1
    C_UseYn     = 2
    C_OrgTypeNm = 3
    C_OrgCd     = 4
    C_OrgNm     = 5
End Sub    
    
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = PopupParent.OPMD_CMODE                   
    lgStrPrevOccurDt = ""                           
    lgStrPrevOrgCd   = ""                           
    lgStrPrevOrgType = ""
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
'        .OperationMode = 5            
        Call ggoSpread.Spreadinit("V20021124",,PopupParent.gAllowDragDropSpread)
    
        .ReDraw = false
        .MaxCols = C_OrgNm + 1                        
        .MaxRows = 0

        Call GetSpreadColumnPos("A")        
       
        ggoSpread.SSSetDate C_OccurDt, "조직변경일", 12, 2, PopupParent.gDateFormat      
        ggoSpread.SSSetCheck C_UseYn, "현재조직", 10, 2, "", True    
        ggoSpread.SSSetEdit C_OrgTypeNm, "조직형태", 25
        ggoSpread.SSSetEdit C_OrgCd, "조직코드", 15
        ggoSpread.SSSetEdit C_OrgNm, "조직명", 25
        .ReDraw = true
        
        Call SetSpreadLock

        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_OccurDt   , -1, C_OccurDt
        ggoSpread.SpreadLock C_UseYn   , -1, C_UseYn
        ggoSpread.SpreadLock C_OrgTypeNm   , -1, C_OrgTypeNm
        ggoSpread.SpreadLock C_OrgCd   , -1, C_OrgCd
        ggoSpread.SpreadLock C_OrgNm   , -1, C_OrgNm
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
        
            C_OccurDt    =  iCurColumnPos(1)
            C_UseYn      =  iCurColumnPos(2)
            C_OrgTypeNm  =  iCurColumnPos(3)
            C_OrgCd      =  iCurColumnPos(4)
            C_OrgNm      =  iCurColumnPos(5)
            
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
    Dim strUsrId

    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

    Call ggoOper.LockField(Document, "N")                                   
    
    Call InitSpreadSheet                                                    
    Call InitVariables                                                      

    strUsrId = arrParent(1)

    If Trim(strUsrId) = "" Then
       frm1.txtUsrId.focus 
    Else
       frm1.txtUsrId.value = strUsrId
       Call FncQuery()
    End if

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
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
        If lgStrPrevOccurDt <> "" Then                                                        
            If DbQuery = False Then                                                       
              Exit Sub
            End If
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
    Call ggoSpread.ClearSpreadData        
    Call InitVariables

    If DbQuery = False Then
       Exit Function
    End If
       
    FncQuery = True                                                                
    
End Function

'=========================================================================================================
Function DbQuery() 
    Dim strVal
    
    On Error Resume next
    Err.Clear                                                               
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    With frm1
        If lgIntFlgMode = PopupParent.OPMD_UMODE Then                                                    
            strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001                            
            strVal = strVal & "&txtUsrId=" & Trim(.htxtUsrId.value)                
            strVal = strVal & "&lgStrPrevOccurDt=" & lgStrPrevOccurDt
            strVal = strVal & "&lgStrPrevOrgCd="   & lgStrPrevOrgCd
            strVal = strVal & "&lgStrPrevOrgType=" & lgStrPrevOrgType                        
            strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        Else
            strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001                            
            strVal = strVal & "&txtUsrId=" & Trim(.txtUsrId.value)                
            strVal = strVal & "&lgStrPrevOccurDt=" & lgStrPrevOccurDt
            strVal = strVal & "&lgStrPrevOrgCd="   & lgStrPrevOrgCd
            strVal = strVal & "&lgStrPrevOrgType=" & lgStrPrevOrgType                                    
            strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows        
        End If

    Call RunMyBizASP(MyBizASP, strVal)                                        

    End With
    
    DbQuery = True
    
End Function

'=========================================================================================================
Function DbQueryOk()                                                        

    lgIntFlgMode = PopupParent.OPMD_UMODE                                                
    
    Call ggoOper.LockField(Document, "Q")                                    
    frm1.vspdData.Focus
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
'    Name : OpenUsrId()
'    Description : UserId PopUp
'=========================================================================================================
Function OpenUsrId(Byval strCode, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "사용자정보 팝업"                                ' 팝업 명칭 
    arrParam(1) = "z_usr_mast_rec"                                     ' TABLE 명칭 
    arrParam(2) = strCode                                              ' Code Condition
    arrParam(3) = ""                                                   ' Name Cindition
    arrParam(4) = ""                                                   ' Where Condition
    arrParam(5) = "사용자 ID"            
    
    arrField(0) = "Usr_id"                                             ' Field명(0)
    arrField(1) = "Usr_nm"                                             ' Field명(1)
    
    arrHeader(0) = "사용자"                                        ' Header명(0)
    arrHeader(1) = "사용자명"                                      ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",  Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetUsrId(arrRet, iWhere)        'return value setting
    End If    
	frm1.txtUsrId.focus             
End Function


'=========================================================================================================
'    Name : SetUsrId()
'    Description : UsrId Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetUsrId(Byval arrRet, Byval iWhere)
    With frm1
        If iWhere = 0 Then
            .txtUsrId.value = arrRet(0)
            .txtUsrNm.value = arrRet(1)
        End If
    End With
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
                        <TD CLASS="TD5">사용자</TD>
                        <TD CLASS="TD656">
                            <INPUT TYPE=TEXT NAME="txtUsrId" SIZE=13 MAXLENGTH=13 tag="12"
                              ALT="사용자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUsrId" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUsrId frm1.txtUsrId.value,0">&nbsp;
                            <INPUT TYPE=TEXT ID="txtUsrNm" NAME="arrCond" SIZE=20 tag="14X">
                        </TD>
                    </TR>
                    <TR>

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
                        <script language =javascript src='./js/za014ra1_vaSpread1_vspdData.js'></script>
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
                            <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
                        <TD WIDTH=30% ALIGN=RIGHT>
                            <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="<%=BIZ_PGM_ID%>" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="htxtUsrId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
