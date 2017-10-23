<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za005ra1
*  4. Program Name         : Table-Program Mapping Popup
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2002.07.15
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : LeeJaeWan
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
Const BIZ_PGM_ID = "Za005rb1.asp"            
'=========================================================================================================
Dim C_MnuId
Dim C_MnuNm


Dim arrReturn
Dim arrParent
Dim PopupParent

<!-- #Include file="../../inc/lgvariables.inc" -->    

'=========================================================================================================
Dim IsOpenPop

arrParent   = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

'=========================================================================================================
Sub InitSpreadPosVariables()
    C_MnuId = 1
    C_MnuNm = 2
End Sub    
        
'=========================================================================================================
Sub InitVariables()
    lgStrPrevKey = ""                           
    lgLngCurRows = 0                            
    lgIntFlgMode = PopupParent.OPMD_CMODE
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
        .MaxCols = C_MnuNm + 1                        
        .MaxRows = 0

         Call GetSpreadColumnPos("A")        
       
        ggoSpread.SSSetEdit C_MnuId, "프로그램 ID", 30, , ,20
        ggoSpread.SSSetEdit C_MnuNm, "프로그램명" , 50, , ,40
        .ReDraw = true
        
        Call SetSpreadLock

        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_MnuId   , -1, C_MnuId
        ggoSpread.SpreadLock C_MnuNm, -1, C_MnuNm
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
            C_MnuNm    =  iCurColumnPos(2)
            
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
    Dim arrParam
    
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

    Call ggoOper.LockField(Document, "N")                                   
    
    Call InitSpreadSheet                                                    
    Call InitVariables                                                      

    '----------  Coding part  -------------------------------------------------------------
    arrParam   = arrParent(1)
    
    frm1.txtTable.value = arrParam
    frm1.txtTable.focus 
    
    Call FncQuery()
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
        Call DbQuery
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
    
    Err.Clear                                                               

    With frm1
        If lgIntFlgMode = PopupParent.OPMD_UMODE Then
            strVal = BIZ_PGM_ID & "?txtMode="    & PopupParent.UID_M0001                            
            strVal = strVal     & "&txtMaxRows=" & .vspdData.MaxRows                    
            strVal = strVal     & "&txtCode="    & Trim(.txtTable.value)                 
            strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey        
        Else
            strVal = BIZ_PGM_ID & "?txtMode="    & PopupParent.UID_M0001                            
            strVal = strVal     & "&txtMaxRows=" & .vspdData.MaxRows                    
            strVal = strVal     & "&txtCode="    & Trim(.txtTable.value)                
            strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey        
        End If
    
        Call LayerShowHide(1)
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
'    Name : OpenTable()
'    Description : 
'=========================================================================================================
Function OpenTable()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

	'---------------------------------------------------------	
	Dim StrMod
    Dim StrLen
    Dim StrSql
    Dim StrMd
    Dim i 
        
    StrMod = Trim(PopupParent.gSetupMod)	 
    StrLen = Len(StrMod)
    
    If StrLen > 0 Then
        For i = 1 To StrLen
            StrMd = Mid(StrMod, i, 1)
            If i =1 Then 
				StrSql = StrSql + "and ("
            ElseIf i > 1 Then
                StrSql = StrSql + "or"
            End If
            StrSql = StrSql + " module_id = " & FilterVar(StrMd, "''", "S") & " "
        Next
        StrSql = StrSql + ")"
    End If
	'---------------------------------------------------------
	
    arrParam(0) = "테이블 팝업"                      ' 팝업 명칭 
    arrParam(1) = "z_table_info"                      ' TABLE 명칭 
    arrParam(2) = frm1.txtTable.value                  ' Code Condition
    arrParam(3) = ""                                  ' Name Cindition
    arrParam(4) = "lang_cd =  " & FilterVar(PopupParent.gLang , "''", "S") & " " & StrSql           ' Where Condition
    arrParam(5) = "테이블"                          ' 조건필드의 라벨 명칭 
    
    arrField(0) = "ED24" & PopupParent.gColSep & "table_id"                          ' Field명(0)
    arrField(1) = "table_nm"                          ' Field명(1)
    
    arrHeader(0) = "테이블 ID"                      ' Header명(0)
    arrHeader(1) = "테이블 명"                      ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=520px; dialogHeight=455px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        frm1.txtTable.value = arrRet(0)
        frm1.txtTableNm.value = arrRet(1)
    End If    
	frm1.txtTable.focus     
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc" -->    
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
                        <TD CLASS="TD5" STYLE="Width:15%">테이블 ID</TD>
                        <TD CLASS="TD6" STYLE="Width:85%">
                            <INPUT TYPE=TEXT NAME="txtTable" SIZE=30 MAXLENGTH=30 tag="12"  ALT="테이블 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btntablepop" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenTable" >
                            <INPUT TYPE=TEXT NAME="txtTableNm" size=40 MAXLENGTH=40 tag="14X">
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
                        <script language =javascript src='./js/za005ra1_vaSpread1_vspdData.js'></script>
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
                            &nbsp;&nbsp;
                            <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="<%=BIZ_PGM_ID%>" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtTable" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
