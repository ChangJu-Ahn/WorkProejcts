<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Role ID 상세                                                                  *
*  2. Function Name        :                                                                              *
*  3. Program ID           : Zb002pa1.asp                                                                *
*  4. Program Name         :                                                                            *
*  5. Program Desc         : Role ID 상세 
*  6. Comproxy List        : Role ID 상세 
*  7. Modified date(First) : 2000/04/03                                                                *
*  8. Modified date(Last)  : 2001/02/28 
*  9. Modifier (First)     : Kang Doo Sig                                                                *
* 10. Modifier (Last)      :                                                                           *
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
Const BIZ_PGM_ID = "Zb002pb1.asp"                              
        

Const POPUP_TITLE    = 0                                     
Const TABLE_NAME    = 1                                     
Const CODE_CON        = 2                                   
Const NAME_CON        = 3                                   
Const WHERE_CON        = 4                                  
Const TEXT_NAME        = 5    

Dim C_MenuId
Dim C_MenuNm
Dim C_MenuType
Dim C_MenuAction
        
Dim SET_CD
Dim SET_NM
'=========================================================================================================
Dim lgQueryFlag                
Dim lgCode                     
Dim lgName                     

<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim arrParent
Dim PopupParent
Dim arrParam                 
Dim arrTblField              
Dim arrGridHdr               
Dim arrReturn                
Dim gintDataCnt              
Dim arrSetVal

        
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

arrParam = arrParent(1)
arrTblField = arrParent(2)
arrGridHdr = arrParent(3)
arrSetVal = arrParent(4)
        

Sub InitSpreadPosVariables()            
    C_MenuId     = 1
    C_MenuNm     = 2
    C_MenuType   = 3
    C_MenuAction = 4
End Sub    

'=========================================================================================================
Function InitVariables()
    Dim intLoopCnt
    Dim CheckedType
    gintDataCnt = 0
            
    For intLoopCnt = 0 To 6
        If arrTblField(intLoopCnt) <> "" Then
            gintDataCnt = gintDataCnt + 1    
        Else
            Exit For
        End If
    Next
        
    lgCode = ""
        
    Call ggoOper.LockField(Document, "N")                                  
        
    lgSortKey = 1
        
    If rdoConf(0).checked = True Then
        CheckedType =  "M"
    else
        if rdoConf(1).checked = True Then
            CheckedType =  "P"
        else 
            CheckedType = ""
        End if
    End if    
        
        
    txtMenuType.value= CheckedType
        
End Function
'=========================================================================================================
Sub SetDefaultVal()
    lblTitle.innerHTML = arrParam(TEXT_NAME)
    txtCd.value = arrSetVal(SET_CD)
    txtNm.value = arrSetVal(SET_NM)
    Self.Returnvalue = Array("")
End Sub
'=========================================================================================================
Sub InitSpreadSheet()

    Call InitSpreadPosVariables()
    
    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        Call ggoSpread.Spreadinit("V20021124",,PopupParent.gAllowDragDropSpread)
    
        .ReDraw = false
        .MaxCols = C_MenuAction + 1
        .MaxRows = 0

        Call GetSpreadColumnPos("A")        
       
        ggoSpread.SSSetEdit C_MenuId, arrGridHdr(0), 20    ' 메뉴 ID 
        ggoSpread.SSSetEdit C_MenuNm, arrGridHdr(1), 35    ' 메뉴명 
        ggoSpread.SSSetEdit C_MenuType, arrGridHdr(2), 15    ' 메뉴 Type 
        ggoSpread.SSSetEdit C_MenuAction, arrGridHdr(3), 17    ' Action     
        .ReDraw = true
        
        Call SetSpreadLock

        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_MenuId   , -1, C_MenuId
        ggoSpread.SpreadLock C_MenuNm, -1, C_MenuNm
        ggoSpread.SpreadLock C_MenuType   , -1, C_MenuType
        ggoSpread.SpreadLock C_MenuAction, -1, C_MenuAction
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

            C_MenuId     = iCurColumnPos(1)
            C_MenuNm     = iCurColumnPos(2)
            C_MenuType   = iCurColumnPos(3)
            C_MenuAction = iCurColumnPos(4)
        
    End Select
End Sub

'=========================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       With frm1.vspdData
            For iDx = 1 To  .MaxCols - 1
                .Col = iDx
                .Row = iRow
                If .ColHidden <> True And .BackColor <> UC_PROTECTED Then
                   .Action = 0 
                   Exit For
                End If
            Next
       End With
    End If   
End Sub

'=========================================================================================================
Sub Form_Load()
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

    Call InitVariables
    Call SetDefaultVal()
    Call InitSpreadSheet()
    Call OpenDetailRole()
        
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
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
Function vspdData_KeyPress(KeyAscii)
    ON Error Resume Next
    If KeyAscii = 27 Then
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
    
        
                
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)  Then
        'Call DisableToolBar(PopupParent.TBC_QUERY)                                            
            
        If lgCode <> ""  Then
            If DBQuery = False Then 
               Call RestoreToolBar()
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
Function FncQuery()
        
    Dim CheckedType
    frm1.vspdData.MaxRows = 0
                        
    Call InitVariables    
            
    lgQueryFlag = "1"

    If rdoConf(0).checked = True Then
        CheckedType =  "M"
    else
        if rdoConf(1).checked = True Then
            CheckedType =  "P"
        else 
            CheckedType = ""
        End if
    End if    
        
        
    txtMenuType.value= CheckedType
        
    Call DbQuery()

End Function
'=========================================================================================================
Function document_onkeypress()
    If window.event.keyCode = 27 Then
        Call CancelClick()
    End If
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
    Dim CheckedType
        
    DbQuery = False                                                         

    Call LayerShowHide(1)
        
    txtNm.value =""

    strVal = BIZ_PGM_ID & "?"
    strVal = strVal & "txtCode=" & lgCode        
    strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows                        
    strVal = strVal & "&txtroleid=" & Trim(txtCd.value)
    strVal = strVal & "&txtmenuid=" & Trim(txtMNUID.value)
    strVal = strVal & "&txtmenutype=" & txtMenuType.value         
    strVal = strVal & "&Flag=" & lgQueryFlag

    Call RunMyBizASP(MyBizASP, strVal)                                        
                
        
    DbQuery = True                                                          
End Function
'=========================================================================================================
Function DbQueryOK()
    Dim IntRetCD

    If frm1.vspdData.MaxRows = 0 Then
        Call DisplayMsgBox("900014","x","x","x") 
        If Trim(txtCd.value) > "" Then
            txtCd.Select 
            txtCd.Focus
        Else   
            txtNm.Select 
            txtNm.Focus
        End If
    End If          
End Function

'=========================================================================================================
'    Name : OpenRole()
'    Description : Composite Role PopUp
'=========================================================================================================
Function OpenRole(Byval strCode, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    'If IsOpenPop = True Then Exit Function

    'IsOpenPop = True

    arrParam(0) = "Role 팝업"                    ' 팝업 명칭 
    arrParam(1) = "Z_USR_ROLE"                            ' TABLE 명칭 
    arrParam(2) = strCode                                    ' Code Condition
    arrParam(3) = ""                                        ' Name Cindition
    arrParam(4) = "COMPST_ROLE_TYPE =" & FilterVar("0", "''", "S") & " "                                        ' Where Condition
    arrParam(5) = "Role ID"            
    
    arrField(0) = "USR_ROLE_ID"                            ' Field명(0)
    arrField(1) = "USR_ROLE_NM"                            ' Field명(1)
    
    arrHeader(0) = "Role ID"                    ' Header명(0)
    arrHeader(1) = "Role 명"                    ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
             "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    'IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetComposite(arrRet)
    End If    
    
    txtCd.focus 
    
End Function
'=========================================================================================================
'    Name : SetComposite()
'    Description : Plant Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetComposite(byval arrRet)
    txtCd.Value = arrRet(0)
    txtNm.Value  = arrRet(1)
End Function
'=========================================================================================================
Function OKClik()
    Dim intColCnt
    Dim nActiveRow
        
    If frm1.vspdData.ActiveRow > 0 Then    
        ReDim arrReturn(frm1.vspdData.MaxCols - 1)
        nActiveRow = frm1.vspdData.ActiveRow
        For intColCnt = 0 To frm1.vspdData.MaxCols - 1
            arrReturn(intColCnt) = GetSpreadText(.vspdData, intColCnt + 1, nActiveRow, "X", "X")
        Next
        Self.Returnvalue = arrReturn
    End If
        
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
Sub OpenDetailRole()

    frm1.vspdData.MaxRows = 0

    lgQueryFlag = "1"
        
    Call DbQuery()

End Sub    
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->    
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB">
    <TR>
        <TD HEIGHT=40>            
            <FIELDSET CLASS="CLSFLD">
            <TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
                <TR>
                    <TD CLASS="TD5" WIDTH=30%><SPAN CLASS="normal" ID="lblTitle">&nbsp;</SPAN></TD>
                    <TD CLASS="TD656" WIDTH=70%><INPUT TYPE="Text" Name="txtCd" SIZE=16 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRole(txtcd.value,0)">&nbsp;<INPUT TYPE="Text" NAME="txtNm" SIZE=40 MAXLENGTH=40 tag="14N" onkeypress="ConditionKeypress"></TD>
                </TR>        
                
                <TR>
                    <TD CLASS="TD5" WIDTH=30%>Menu ID</TD>
                    <TD CLASS="TD656" WIDTH=70%><INPUT TYPE="Text" Name="txtMNUID" SIZE=18 MAXLENGTH=18 tag="11"></TD>
                    
                </TR>
                
                <TR><TD CLASS="TD5">Menu Type</TD> 
                    <TD HEIGHT=9>
                        <SPAN STYLE="width:120;"><INPUT TYPE="RADIO" NAME="rdoConf" ID="rdoConf1" CLASS="RADIO" tag="11"><LABEL FOR="rdoConf1">Menu(M)</LABEL></SPAN>
                         <SPAN STYLE="width:120;"><INPUT TYPE="RADIO" NAME="rdoConf" ID="rdoConf2" CLASS="RADIO" tag="11"><LABEL FOR="rdoConf2">Program(P)</LABEL></SPAN>
                           <SPAN STYLE="width:120;"><INPUT TYPE="RADIO" NAME="rdoConf" ID="rdoConf3" CLASS="RADIO" checked tag="11"><LABEL FOR="rdoConf3">All(A)</LABEL></SPAN>
                    </TD>
                </TR>
            </TABLE>
            </FIELDSET>
        </TD>
    </TR>
    <TR><TD HEIGHT=*>
        <FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
            <script language =javascript src='./js/zb002pa1_vspdData_vspdData.js'></script>
        </FORM>
        </TD>
    </TR>
    <TR HEIGHT=30>
        <TD WIDTH=100%>
            <TABLE <%=LR_SPACE_TYPE_30%>>
                <TR>
                    <TD WIDTH=10>&nbsp;</TD>
                    <TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"></IMG></TD>
                    <TD WIDTH=30% ALIGN=RIGHT>
                    <%'<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>%>
                                               <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
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

<INPUT TYPE=HIDDEN NAME="txtMenuType" tag="24">


    
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>




