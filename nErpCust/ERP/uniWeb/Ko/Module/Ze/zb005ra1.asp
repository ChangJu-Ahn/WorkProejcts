<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : User Role Reference                                                        *
*  2. Function Name        : Zb005ra1.asp                                                                *
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
Const BIZ_PGM_ID = "ZB005RB1.asp"                                
'=========================================================================================================
Dim lgQueryFlag
Dim lgCode                      
Dim arrReturn
Dim arrParent
Dim PopupParent
Dim arrParam

Dim C_RoleId
Dim C_RoleNm

<!-- #Include file="../../inc/lgvariables.inc" -->    
        
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam = arrParent(1)
top.document.title = PopupParent.gActivePRAspName

'=========================================================================================================
Sub InitSpreadPosVariables()            
    C_RoleId = 1
    C_RoleNm = 2
End Sub    

'=========================================================================================================
Function InitVariables()
    ReDim arrReturn(0, 0)
    Self.Returnvalue = arrReturn
    lgSortKey = 1
End Function
'=========================================================================================================
Sub SetDefaultVal()
    lblTitle.innerHTML = "Role ID"
    'txtCd.value = arrParam(0)
    'txtNm.value = arrParam(1)
    txtCompstRoleId.value = arrParam(2)
End Sub
    
'=========================================================================================================
Sub InitSpreadSheet()

    Call InitSpreadPosVariables()
    
    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        .OperationMode = 5'khy200307(멀티)
        Call ggoSpread.Spreadinit("V20021124",,PopupParent.gAllowDragDropSpread)
    
        .ReDraw = false
        .MaxCols = C_RoleNm + 1
        .MaxRows = 0

         Call GetSpreadColumnPos("A")        

        ggoSpread.SSSetEdit C_RoleId, "Role ID",    31            ' 코드명 
        ggoSpread.SSSetEdit C_RoleNm, "Role명",    33            ' 코드     
        .ReDraw = true
        
        Call SetSpreadLock

        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1.vspdData
        .ReDraw = False
        ggoSpread.SpreadLock C_RoleId   , -1, C_RoleId
        ggoSpread.SpreadLock C_RoleNm, -1, C_RoleNm
        .ReDraw = True
    End With
End Sub

'=========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_RoleId       = iCurColumnPos(1)
            C_RoleNm       = iCurColumnPos(2)
            
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
Function document_onkeypress()
    If window.event.keyCode = 27 Then
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
sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If CheckRunningBizProcess = True Then                    
        Exit Sub
    End If
            
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
        
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgQueryFlag <> "1" Then
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
    frm1.vspdData.MaxRows = 0
    lgQueryFlag = "1"    
    lgCode = ""
        
    If DBQuery = False Then                                                        
        Exit Function 
    End If          
End Function
    
'=========================================================================================================
Function DbQuery()
    'Err.Clear                                                               
        
    DbQuery = False                                                         
        
    Dim strVal

    Call LayerShowHide(1)
        
    txtNm.value=""
    strVal = BIZ_PGM_ID & "?"                                             
    strVal = strVal & "txtCd=" & Trim(txtCd.value)                        
    strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows                                
    strVal = strVal & "&txtCompstCd=" & Trim(txtCompstRoleId.value)                            
    strVal = strVal & "&NextCd=" & lgCode
    
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
                    <TD CLASS="TD656"><INPUT TYPE="Text" ID="txtCd" NAME="txtCd" SIZE=18 MAXLENGTH=18 tag="12XXXU">&nbsp;<INPUT TYPE="Text" NAME="txtNm" SIZE=40 MAXLENGTH=40 tag="12"></TD>
                </TR>        
            </TABLE>
            </FIELDSET>
        </TD>
    </TR>
    <TR><TD HEIGHT=*>
        <FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
            <script language =javascript src='./js/zb005ra1_vspdData_vspdData.js'></script>
        </FORM>
        </TD>
    </TR>
    <TR HEIGHT=20>
        <TD WIDTH=100%>
            <TABLE <%=LR_SPACE_TYPE_30%>>
                <TR>
                    <TD WIDTH=10>&nbsp;</TD>
                    <TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
                    <TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
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
<INPUT TYPE=HIDDEN NAME="txtCompstRoleId" tag="24"> 
    
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
