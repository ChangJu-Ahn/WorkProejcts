<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za008ra1
*  4. Program Name         : Audit Management Reference Popup
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000.03.13
*  8. Modified date(Last)  : 2002.06.13
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
 Const BIZ_PGM_ID = "Za008RB1.asp"                                                
'=========================================================================================================
Dim C_TableID
Dim C_TableNm
Dim C_TableTypeCd
Dim C_TableType
Dim C_Insert
Dim C_Update
Dim C_Delete
 

Dim arrReturn
Dim arrParent
Dim PopupParent

<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim IsOpenPop     

arrParent   = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

'=========================================================================================================
Sub InitSpreadPosVariables()
    C_TableID        = 1
    C_TableNm        = 2
    C_TableTypeCd    = 3
    C_TableType        = 4
    C_Insert        = 5
    C_Update        = 6
    C_Delete        = 7
End Sub    
   
'=========================================================================================================
Sub InitVariables()
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
        .MaxCols = C_Delete + 1                        
        .MaxRows = 0

        Call GetSpreadColumnPos("A")        
       
        ggoSpread.SSSetEdit        C_TableID,    "테이블 ID",    25,,,30
        ggoSpread.SSSetEdit        C_TableNm,    "테이블명",        30,,,40
        ggoSpread.SSSetEdit        C_TableTypeCd, "",                15
        ggoSpread.SSSetEdit        C_TableType,"테이블속성",    15,,,16
        ggoSpread.SSSetCheck    C_Insert,    "입력",            6,,,1
        ggoSpread.SSSetCheck    C_Update,    "수정",            6,,,1
        ggoSpread.SSSetCheck    C_Delete,    "삭제",            6,,,1
        .ReDraw = true
        
        Call SetSpreadLock

        Call ggoSpread.SSSetColHidden(C_TableTypeCd, C_TableTypeCd, True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_TableID   , -1, C_Delete
        ggoSpread.SpreadLock C_TableNm   , -1, C_TableNm
        ggoSpread.SpreadLock C_TableTypeCd   , -1, C_TableTypeCd
        ggoSpread.SpreadLock C_TableType   , -1, C_TableType
        ggoSpread.SpreadLock C_Insert   , -1, C_Insert
        ggoSpread.SpreadLock C_Update   , -1, C_Update
        ggoSpread.SpreadLock C_Delete   , -1, C_Delete
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
            
            C_TableID      =  iCurColumnPos(1)
            C_TableNm      =  iCurColumnPos(2)
            C_TableTypeCd  =  iCurColumnPos(3)
            C_TableType    =  iCurColumnPos(4)
            C_Insert       =  iCurColumnPos(5)
            C_Update       =  iCurColumnPos(6)
            C_Delete       =  iCurColumnPos(7)

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

    '----------  Coding part  -------------------------------------------------------------
   
    Redim arrReturn(0, 0)
    Self.Returnvalue = arrReturn

    frm1.txtTableID.focus
    
    Call FncQuery()
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
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
Sub vspdData_Change(ByVal Col, ByVal Row)
    Dim iDx
       
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
End Sub

'=========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'=========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If CheckRunningBizProcess = True Then
       Exit Sub
    End If

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)    And lgStrPrevKey <> "" Then
        DbQuery
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
    
    Err.Clear                                                               
    DbQuery = False
    
    Call LayerShowHide(1)
    
    With frm1
    
        strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001                          
        strVal = strVal & "&txtTableID="   & Trim(.txtTableID.value)                
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtMaxRows="   & .vspdData.MaxRows

        Call RunMyBizASP(MyBizASP, strVal)                                        
        
    End With
    
    DbQuery = True
    
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
            arrReturn(intInsRow, 0) = GetSpreadText(frm1.vspdData, C_TableID, intRowCnt, "X", "X")
            arrReturn(intInsRow, 1) = GetSpreadText(frm1.vspdData, C_TableNm, intRowCnt, "X", "X")
            arrReturn(intInsRow, 2) = GetSpreadText(frm1.vspdData, C_TableTypeCd, intRowCnt, "X", "X")
            arrReturn(intInsRow, 3) = GetSpreadText(frm1.vspdData, C_TableType, intRowCnt, "X", "X")
            arrReturn(intInsRow, 4) = GetSpreadText(frm1.vspdData, C_Insert, intRowCnt, "X", "X")
            arrReturn(intInsRow, 5) = GetSpreadText(frm1.vspdData, C_Update, intRowCnt, "X", "X")
            arrReturn(intInsRow, 6) = GetSpreadText(frm1.vspdData, C_Delete, intRowCnt, "X", "X")
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
'    Name : OpenTableID()
'    Description : TableID PopUp
'=========================================================================================================
Function OpenTableID(Byval strCode)

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "테이블 ID Popup"
    arrParam(1) = "z_pkg_audit_policy p, z_table_info t"
    arrParam(2) = strCode                                                 ' Code Condition
    arrParam(3) = ""                                                      ' Name Cindition
    arrParam(4) = "p.LANG_CD= " & FilterVar(PopupParent.gLang, "''", "S") & " " & _
                  "and p.lang_cd = t.lang_cd " & _
                  "and p.table_id = t.table_id "
'                  "and t.USE_YN='1'"
    arrParam(5) = "테이블 ID"            
    
    arrField(0) = "ED24" & PopupParent.gColSep & "p.Table_ID"
    arrField(1) = "t.Table_NM"                     
    
    arrHeader(0) = "테이블 ID"                                        ' Header명(0)
    arrHeader(1) = "테이블명"                                         ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=520px; dialogHeight=455px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetTableID(arrRet)
    End If    
	frm1.txtTableID.focus 
End Function

'=========================================================================================================
'    Name : SetTableID()
'    Description : Plant Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetTableID(byval arrRet)
    With frm1
        .txtTableID.Value = arrRet(0)
        .txtTableNm.Value = arrRet(1)
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
                        <TD CLASS="TD5">테이블 ID</TD>
                        <TD CLASS="TD656" COLSPAN=3><INPUT TYPE=TEXT NAME="txtTableID" SIZE=30 MAXLENGTH=30 tag="11"  ALT="테이블 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTableID" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTableID frm1.txtTableID.value">&nbsp;<INPUT TYPE=TEXT NAME="txtTableNm" SIZE=40 tag="14N"></TD>
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
                        <script language =javascript src='./js/za008ra1_vspdData_vspdData.js'></script>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <tr>
        <TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <TR>
        <TD HEIGHT=30>
            <TABLE <%=LR_SPACE_TYPE_30%>>
                <TR>
                    <TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
                        <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
                    <TD WIDTH=30% ALIGN=RIGHT>
                        <IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
                        <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="Za008RB1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
