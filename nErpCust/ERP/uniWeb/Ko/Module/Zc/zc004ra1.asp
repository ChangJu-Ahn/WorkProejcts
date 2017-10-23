<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : User Role Reference                                                        *
*  2. Function Name        : ZC004ra1.asp                                                                *
*  3. Program ID           :                                                                             *
*  4. Program Name         :                                                                            *
*  5. Program Desc         : Reference Popup                                                           *
*  6. Comproxy List        :                                                                            *
*  7. Modified date(First) : 2000/03/29                                                                *
*  8. Modified date(Last)  : 2001/02/28                                                                *
*  9. Modifier (First)     : Kang Tae Bum                                                                *
* 10. Modifier (Last)      : Kang Doo Sig                                                                *
* 11. Comment              :                                                                            *
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
Const BIZ_PGM_ID = "ZC004RB1.asp"                                
'=========================================================================================================
Dim C_RoleId
Dim C_Action

dim lgQueryFlag
dim lgCode                      
dim arrReturn
dim arrParent
Dim PopupParent
dim arrParam

<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim IsOpenPop            
        
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam = arrParent(1)
top.document.title = PopupParent.gActivePRAspName

'=========================================================================================================
Sub InitSpreadPosVariables()
    C_RoleId = 1
    C_Action = 2
End Sub    

'=========================================================================================================
Function InitVariables()
    Redim arrReturn(0, 0)
    Self.Returnvalue = arrReturn
    lgSortKey = 1
End Function
'=========================================================================================================
Sub SetDefaultVal()
    lblTitle.innerHTML = "Menu ID"
    txtCd.value = arrParam(0)
    txtNm.value = arrParam(1)
End Sub
    
'=========================================================================================================
Sub InitSpreadSheet()
    Call InitSpreadPosVariables()

    with vspddata
        
        .ReDraw = False
            
        ggoSpread.Source = vspdData
        Call ggoSpread.Spreadinit("V20021124",,PopupParent.gAllowDragDropSpread)

        .MaxCols = C_Action + 1
        .MaxRows = 0
        
         Call GetSpreadColumnPos("A")    
            
        ggoSpread.SSSetEdit C_RoleId, "Role ID",    30            
        ggoSpread.SSSetEdit C_Action, "Action",    30            
        .ReDraw = True
        
        Call SetSpreadLock

        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    End With
End Sub

'=========================================================================================================
Sub SetSpreadLock()
        vspdData.ReDraw = False
        ggoSpread.SpreadLock C_RoleId, -1, C_RoleId
        ggoSpread.SpreadLock C_Action, -1, C_Action
        vspdData.ReDraw = True
End Sub

'=========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
        
            C_RoleId    =  iCurColumnPos(1)
            C_Action    =  iCurColumnPos(2)
            
    End Select
End Sub

'=========================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  vspdData.MaxCols - 1
           vspdData.Col = iDx
           vspdData.Row = iRow
           If vspdData.ColHidden <> True And vspdData.BackColor <> UC_PROTECTED Then
              vspdData.Action = 0 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

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
Sub InitComboBox()
End Sub

'=========================================================================================================
Sub InitSpreadComboBox()
End Sub

'=========================================================================================================
Sub Form_Load()
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

    Call InitVariables
    Call SetDefaultVal()
    Call InitSpreadSheet()
        
    If DBQuery = False Then 
       Exit Sub 
    End If 
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    
End Sub

'=========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
    
    gMouseClickStatus = "SPC"   
    
    Set gActiveSpdSheet = vspdData

    If vspdData.MaxRows <= 0 Then                                                    
       Exit Sub
       End If
           
    If Row <= 0 Then
        ggoSpread.Source = vspdData
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
    ggoSpread.Source = vspdData
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
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'=========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
	Dim iDx
	vspdData.Row = Row
	vspdData.Col = Col
End Sub

'=========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = vspdData
End Sub

'=========================================================================================================
Function FncQuery()
    vspdData.MaxRows = 0
    lgQueryFlag = "1"    
    lgCode = ""
        
    txtNm.value = ""
        
    If DBQuery = False Then                                                        
        Exit Function 
    End If          
End Function

'=========================================================================================================
Function document_onkeypress()
    If window.event.keyCode = 27 Then
        Call CancelClick()
    End If
End Function
'=========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
    If vspdData.MaxRows > 0 Then
        If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
        End If
    End If
End Function
'=========================================================================================================
Function vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function
'=========================================================================================================
sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If CheckRunningBizProcess = True Then                    
        Exit Sub
    End If
            
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
        
    
    'msgbox "vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) " & vbcrlf &  vspdData.MaxRows & vbcrlf & NewTop & vbcrlf &  VisibleRowCnt(vspdData,NewTop)  & vbcrlf & lgQueryFlag
    If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) And lgQueryFlag <> "1" Then
        If lgCode <> "" Then
            If DBQuery = False Then 
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
    Call InitSpreadComboBox
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
Function DbQuery()
    'Err.Clear                                                               
        
    DbQuery = False                                                         
        
    Dim strVal

    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?"                                                
    strVal = strVal & "txtCd=" & Trim(txtCd.value)                    
    strVal = strVal & "&NextCd=" & lgCode
    strVal = strVal & "&txtMaxRows=" & vspdData.MaxRows                    
                
    Call RunMyBizASP(MyBizASP, strVal)                                        
        
    DbQuery = True                                                          
End Function

'=========================================================================================================
Function OpenMnuInfo(Byval strCode, Byval iWhere)

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "메뉴 팝업"
    arrParam(1) = "Z_LANG_CO_MAST_MNU"    
    arrParam(2) = strCode
    arrParam(3) = ""        
    arrParam(4) = "LANG_CD =  " & FilterVar(PopupParent.gLang, "''", "S") & ""
    
                
    arrParam(5) = "메뉴ID"
    
    arrField(0) = "MNU_ID"
    arrField(1) = "MNU_NM"
    
    arrHeader(0) = "메뉴ID"
    arrHeader(1) = "메뉴명"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetMnuInfo(arrRet, iWhere)
    End If    

    txtCD.focus
    
End Function
'=========================================================================================================
Function SetMnuInfo(Byval arrRet, Byval iWhere)

        Select Case iWhere
            Case  1
                txtCD.Value    = arrRet(0)
                txtNm.Value    = arrRet(1)            
        End Select

End Function
    
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->    
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
    <TR><TD HEIGHT=40>
            <FIELDSET CLASS="CLSFLD"><TABLE WIDTH=*>
                <TR>
                    <TD CLASS="TD5"><SPAN CLASS="normal" ID="lblTitle">&nbsp;</SPAN></TD>
                    <TD CLASS="TD656"><INPUT TYPE="Text" ID="txtCd" NAME="txtCd" SIZE=18 MAXLENGTH=18 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif"   NAME="btnMnuID" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMnuInfo txtCD.value,1 ">&nbsp;<INPUT TYPE="Text" NAME="txtNm" SIZE=40 MAXLENGTH=40 tag="12"></TD>
                </TR>        
            </TABLE>
            </FIELDSET>
        </TD>
    </TR>
    <TR><TD HEIGHT=*>
            <script language =javascript src='./js/zc004ra1_vspdData_vspdData.js'></script>
        </TD>
    </TR>
    <TR HEIGHT=20>
        <TD WIDTH=100%>
            <TABLE <%=LR_SPACE_TYPE_30%>>
                <TR>
                    <TD WIDTH=10>&nbsp;</TD>
                    <TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
                    <TD WIDTH=30% ALIGN=RIGHT>
                                              <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>                    </TD>
                    <TD WIDTH=10>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR HEIGHT=20>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE> 
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
