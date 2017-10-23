
<!--
======================================================================================================
*  1. Module Name          : Menu ID Reference                                                            *
*  2. Function Name        : Zb003ra1.asp                                                                *
*  3. Program ID           :                                                                             *
*  4. Program Name         :                                                                            *
*  5. Program Desc         : Menu ID Reference Popup                                                   *
*  6. Comproxy List        :                                                                             *
*  7. Modified date(First) : 2000/04/17                                                                *
*  8. Modified date(Last)  : 2001/12/06                                                                *
*  9. Modifier (First)     : Kang Tae Bum                                                                *
* 10. Modifier (Last)      : LEE SEOK GON                                                                *
* 11. Comment              :                                                                            *                                                                                           *
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
Const BIZ_PGM_ID = "ZB003RB1.asp"                              
'=========================================================================================================
Dim lgQueryFlag
Dim lgMnuType
Dim lgActType
Dim lgCode                     
Dim lgSortKey
        
Dim C_MenuId
Dim C_MenuNm
Dim C_MenuTypeCd
Dim C_MenuType
Dim C_MenuActionCd
Dim C_MenuAction
        
Dim arrReturn
Dim arrParent
Dim PopupParent
Dim arrParam                    
        
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam = arrParent(1)
top.document.title = PopupParent.gActivePRAspName

'=========================================================================================================
Sub InitSpreadPosVariables()            
    C_MenuId        = 1
    C_MenuNm        = 2
    C_MenuTypeCd    = 3
    C_MenuType      = 4
    C_MenuActionCd  = 5
    C_MenuAction    = 6
End Sub    
'=========================================================================================================
Function InitVariables()
    ReDim arrReturn(0, 0)
    Self.Returnvalue = arrReturn
End Function
'=========================================================================================================

Sub SetDefaultVal()
    lblTitle.innerHTML = "Menu ID"
    txtRoleID.value = arrParam(0)
    txtCd.value = arrParam(1)
                
    If arrParam(2) = "M" Then 
        rdoConf1.checked = true
        lgMnuType = "M"
    ElseIf arrParam(2) = "P" Then 
        rdoConf2.checked = true
        lgMnuType = "P"
    Else
        rdoConf3.checked = true
        lgMnuType = ""
    End If        
        
    If arrParam(3) = "Q" Then             
        rdoAct1.checked=TRUE
        lgActType = "Q"
    ElseIf arrParam(3) = "E" Then             
        rdoAct2.checked=TRUE
        lgActType = "E"
    ElseIf arrParam(3) = "N"Then             
        rdoAct3.checked=TRUE
        lgActType = "N"
    Else'A            
        rdoAct4.checked=TRUE
        lgActType = "A"
    End If
    
    txtMenuType.value  = lgMnuType 
    txtActionID.value  = lgActType
        
End Sub
'=========================================================================================================
Sub InitSpreadSheet()

    Call InitSpreadPosVariables()
    
    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData    
    	.OperationMode = 5			'khy200307(멀티선택)
        Call ggoSpread.Spreadinit("V20021124",,PopupParent.gForbidDragDropSpread)
        
        .ReDraw = false
        .MaxCols = C_MenuAction + 1
        .MaxRows = 0

        Call GetSpreadColumnPos("A")        

        ggoSpread.SSSetEdit C_MenuId, "Menu ID",    20            ' 메뉴 ID 
        ggoSpread.SSSetEdit C_MenuNm, "Menu명",    35            ' 메뉴명 
        ggoSpread.SSSetEdit C_MenuTypeCd, "",                16
        ggoSpread.SSSetEdit C_MenuType, "Menu Type", 16            ' 메뉴 Type 
        ggoSpread.SSSetEdit C_MenuActionCd, "",                17
        ggoSpread.SSSetEdit C_MenuAction, "Action",    17            ' Action 
        .ReDraw = true
        
        Call SetSpreadLock
        
        Call ggoSpread.SSSetColHidden(C_MenuTypeCd, C_MenuTypeCd, True)
        Call ggoSpread.SSSetColHidden(C_MenuActionCd, C_MenuActionCd, True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_MenuId   , -1, C_MenuId
        ggoSpread.SpreadLock C_MenuNm, -1, C_MenuNm
        ggoSpread.SpreadLock C_MenuTypeCd, -1, C_MenuTypeCd
        ggoSpread.SpreadLock C_MenuType   , -1, C_MenuType
        ggoSpread.SpreadLock C_MenuActionCd, -1, C_MenuActionCd
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

            C_MenuId       = iCurColumnPos(1)
            C_MenuNm       = iCurColumnPos(2)
            C_MenuTypeCd   = iCurColumnPos(3)                        
            C_MenuType     = iCurColumnPos(4)
            C_MenuActionCd = iCurColumnPos(5)
            C_MenuAction   = iCurColumnPos(6)                                
            
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
			       .Col = iDx
			       .Row = iRow
			       .Action = 0 ' go to 
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
        
        
    If DBQuery = False Then 
       Exit Sub 
    End If 
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    
End Sub

'=========================================================================================================
Function OKClick()'khy200307
	Dim intColCnt, intRowCnt, intInsRow
		
	if frm1.vspdData.SelModeSelCount > 0 Then 
		
		intInsRow = 0
		
		Redim arrReturn(frm1.vspdData.SelModeSelCount - 1, frm1.vspdData.MaxCols - 1)
		
		For intRowCnt = 0 To frm1.vspdData.MaxRows - 1
			
			frm1.vspdData.Row = intRowCnt + 1
			
			If frm1.vspdData.SelModeSelected Then
				For intColCnt = 0 To frm1.vspdData.MaxCols - 1
					frm1.vspdData.Col = intColCnt + 1
					arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
				Next
					
				intInsRow = intInsRow + 1
					
			End IF
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
sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If CheckRunningBizProcess = True Then                    
        Exit Sub
    End If
            
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    With frm1.vspdData        
        If .MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgQueryFlag <> "1" Then
            If lgCode <> "" Then
                If DBQuery = False Then 
                   Exit Sub 
                End If 
            End If
        End if
    End With
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
    frm1.vspdData.MaxRows = 0
    lgQueryFlag = "1"
        
    If rdoConf(0).checked then 
        lgMnuType = "M"
    ElseIf rdoConf(1).checked then 
        lgMnuType = "P"
    ElseIf rdoConf(2).checked then    
        lgMnuType = ""
    End If

    If rdoAct1.checked Then             
        lgActType = "Q"
    ElseIf rdoAct2.checked Then             
        lgActType = "E"
    ElseIf rdoAct3.checked Then             
        lgActType = "N"
    Else'A            
        lgActType = "A"
    End If
        
    lgCode = ""

    txtMenuType.value  = lgMnuType 
    txtActionID.value  = lgActType
        
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
    Dim strVal
        
    DbQuery = False  
            
    Call LayerShowHide(1)
    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
    strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows                            
    strVal = strVal & "&txtRoleID="   & Trim(txtRoleID.value)                    
    strVal = strVal & "&txtCd="       & Trim(txtCd.value)                    
    strVal = strVal & "&txtNm="       & Trim(txtNm.value)                                        
    strVal = strVal & "&MnuType="     & txtMenuType.value 
    strVal = strVal & "&ActType="     & txtActionID.value 
    strVal = strVal & "&NextCd="      & lgCode
        
    Call RunMyBizASP(MyBizASP, strVal)                                       
        
    DbQuery = True                                                             
End Function
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtCd.focus
	End If

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
                <TD CLASS="TD656"><INPUT TYPE="Text" ID="txtCd" NAME="txtCd" SIZE=18 MAXLENGTH=18  tag="14XXXU">&nbsp;<INPUT TYPE="Text" NAME="txtNm" SIZE=40 MAXLENGTH=40 tag="12"></TD>                
            </TR>        
            <TR><TD CLASS="TD5">Menu Type</TD> <TD HEIGHT=9>
                <INPUT TYPE="RADIO" NAME="rdoConf" ID="rdoConf1" CLASS="RADIO" tag="11" value="M">            <LABEL FOR="rdoConf1">Menu(M)</LABEL></SPAN>
                 <INPUT TYPE="RADIO" NAME="rdoConf" ID="rdoConf2" CLASS="RADIO" tag="11" value="P">            <LABEL FOR="rdoConf2">Program(P)</LABEL></SPAN>
                   <INPUT TYPE="RADIO" NAME="rdoConf" ID="rdoConf3" CLASS="RADIO" tag="11" value="" Checked>    <LABEL FOR="rdoConf3">All(A)</LABEL></SPAN>
            </TD></TR>    
            <TR><TD CLASS="TD5">Action</TD> <TD HEIGHT=9>
                <INPUT TYPE="RADIO" NAME="rdoAct" ID="rdoAct1" CLASS="RADIO" tag="12" value="Q">            <LABEL FOR="rdoAct1">QUERY</LABEL></SPAN>
                 <INPUT TYPE="RADIO" NAME="rdoAct" ID="rdoAct2" CLASS="RADIO" tag="12" value="E">            <LABEL FOR="rdoAct2">Excel/Print</LABEL></SPAN>
                 <INPUT TYPE="RADIO" NAME="rdoAct" ID="rdoAct3" CLASS="RADIO" tag="12" value="N">            <LABEL FOR="rdoAct3">NONE</LABEL></SPAN>
                   <INPUT TYPE="RADIO" NAME="rdoAct" ID="rdoAct4" CLASS="RADIO" tag="12" value="A" Checked>    <LABEL FOR="rdoAct4">ALL</LABEL></SPAN>
            </TD></TR>                        
        </TABLE></FIELDSET>
    </TD></TR>
    <TR><TD HEIGHT=*>
        <FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
            <script language =javascript src='./js/zb003ra1_vspdData_vspdData.js'></script>
        </FORM>
    </TD></TR>
    <TR HEIGHT=20>
        <TD WIDTH=100%>
            <TABLE <%=LR_SPACE_TYPE_30%>>
                <TR>
                    <TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
                        <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
                    <TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
                                              <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>                    </TD>
                    <TD WIDTH=10>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>    
    <TR HEIGHT=20>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtRoleID" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMenuType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtActionID" tag="24">

    
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
