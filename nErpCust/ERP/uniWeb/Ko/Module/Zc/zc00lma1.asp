
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : 
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 1999/09/10
*  8. Modified date(Last)  : 1999/09/10
*  9. Modifier (First)     : Lee JaeHoo
* 10. Modifier (Last)      : Lee JaeHoo
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">        

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                                 

Const BIZ_PGM_ID = "zc00lmb1.asp"

Dim C_UpperMnu
Dim C_MnuID
Dim C_MnuID_ORG
Dim C_MnuNm
Dim C_MnuType
Dim C_SysLvl
Dim C_MnuSeq

<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim IsOpenPop

'=========================================================================================================
Sub InitSpreadPosVariables()
    C_UpperMnu  = 1
    C_MnuID     = 2
    C_MnuID_ORG = 3
    C_MnuNm     = 4
    C_MnuType   = 5
    C_SysLvl    = 6
    C_MnuSeq    = 7
End Sub

       
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgIntGrpCount = 0
    
    lgStrPrevKey = ""                           
    lgLngCurRows = 0
    lgSortKey = 1
    
End Sub
'=========================================================================================================
Sub SetDefaultVal()
End Sub
'=========================================================================================================
Sub LoadInfTB19029()
End Sub

'=========================================================================================================
Sub InitSpreadSheet()

    Call InitSpreadPosVariables()
    
    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        Call ggoSpread.Spreadinit("V20030725",,Parent.gAllowDragDropSpread)

        .ReDraw = false                   
        .MaxCols = C_MnuSeq +1
        .MaxRows = 0

        Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetEdit    C_UpperMnu, "상위메뉴", 17,,,15, 2
        ggoSpread.SSSetEdit    C_MnuID, "메뉴ID", 17,,,15, 2
        ggoSpread.SSSetEdit    C_MnuID_ORG, "", 0
        ggoSpread.SSSetEdit	C_MnuNm, "메뉴명", 40,,,40
        ggoSpread.SSSetCombo C_MnuType, "메뉴타입", 15, 2, false
        ggoSpread.SSSetEdit C_SysLvl, "메뉴레벨", 15    
        ggoSpread.SSSetEdit C_MnuSeq, "메뉴순서", 20,,,10,2
        .ReDraw = true

        Call SetSpreadLock
        
        Call ggoSpread.MakePairsColumn(C_MnuID,C_MnuNm,"1")
        Call ggoSpread.MakePairsColumn(C_MnuID_ORG,C_MnuID,"1")
   
		Call ggoSpread.SSSetColHidden(C_MnuID_ORG, C_MnuID_ORG, True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
    
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_UpperMnu, -1, C_UpperMnu
        ggoSpread.SpreadLock C_MnuID, -1, C_MnuID
        ggoSpread.SpreadLock C_MnuNm, -1, C_MnuNm
        ggoSpread.SpreadLock C_MnuType, -1, C_MnuType
        ggoSpread.SpreadLock C_SysLvl, -1, C_SysLvl                                
        ggoSpread.SpreadUnLock C_MnuSeq, -1, C_MnuSeq
        .vspdData.ReDraw = True

    End With
End Sub
'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1
    
        .vspdData.ReDraw = False
        ggoSpread.SSSetProtected    C_UpperMnu, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected    C_MnuID, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected    C_MnuNm, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected    C_MnuType, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected    C_SysLvl, pvStartRow, pvEndRow    
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
    
            C_UpperMnu   =  iCurColumnPos(1)
            C_MnuID      =  iCurColumnPos(2)
            C_MnuID_ORG  =  iCurColumnPos(3)
            C_MnuNm      =  iCurColumnPos(4)
            C_MnuType    =  iCurColumnPos(5)
            C_SysLvl     =  iCurColumnPos(6)
            C_MnuSeq     =  iCurColumnPos(7)

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
Sub InitSpreadComboBox()
    Dim strCboData

    strCboData = "M" & vbTab & "P"
    ggoSpread.SetCombo strCboData, C_MnuType
End Sub
'=========================================================================================================
Sub Form_Load()

    Call ggoOper.LockField(Document, "N")
    
    Call InitSpreadSheet
    Call InitVariables

    Call InitSpreadComboBox
    Call SetDefaultVal
    Call SetToolbar("11000000000111")
    
    frm1.txtMnuID.focus
    Set gActiveElement = document.activeElement
    
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub
'=========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
End Sub

'=========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
    Call SetPopupMenuItemInf("0001111111")    
    
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
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim index
    
    With frm1.vspdData
        If Col = C_TypeNm And Row > 0 Then
            .Row = Row
            .Col = Col
    		index = .TypeComboBoxCurSel
            .Col = C_Type
            .TypeComboBoxCurSel = index

            If index = 0 Then
                .Text = "M"
            ElseIf index = 1 Then
                .Text = "S"
            ElseIf index = 2 Then
                .Text = "T"
            ElseIf index = 3 Then
                .Text = "X"
            End if

        End If
    End With
End Sub

'=========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'=========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop)
    
    If CheckRunningBizProcess = True Then
       Exit Sub
    End If

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
        If lgStrPrevKey <> "" Then                            
            Call DisableToolBar(Parent.TBC_QUERY)
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
    Call InitSpreadComboBox
    Call ggoSpread.ReOrderingSpreadData()
    Call InitData()
End Sub
'=========================================================================================================
Function FncQuery()

    Dim IntRetCD 
    
    FncQuery = False

    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
          Exit Function
        End If
    End If

    Call ggoOper.ClearField(Document, "2")
    Call ggoSpread.ClearSpreadData()            
    Call InitVariables

    If DbQuery = False Then
       Exit Function
    End If

    FncQuery = True
    
End Function
'=========================================================================================================
Function FncNew() 
End Function
'=========================================================================================================
Function FncDelete() 
End Function
'=========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False
    
    Err.Clear
    On Error Resume Next
    
    'validate check----------------------------------------------------------------------    
    Dim strvalidate
    
    If CheckNumeric(Trim(GetSpreadText(frm1.vspdData, C_MnuSeq, frm1.vspdData.ActiveRow, "X", "X")))=1 then          
        IntRetCD = DisplayMsgBox("211430", "x","x","x")                'Menu sequence only decimal
        Exit Function        
    End If

    '---------------------------------------------------------------------------------------
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If
    
    If DbSave = False Then
       Exit Function
    End If
    
    FncSave = True
    
End Function
'=========================================================================================================
Function FncCopy() 
End Function
'=========================================================================================================
Function FncCancel() 
    ggoSpread.EditUndo
End Function

'=========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    
    On Error Resume Next                                                              
    Err.Clear                                                                     
    
    FncInsertRow = False                                                             

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If
    
    With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 

    '------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncInsertRow = True                                                              
    End If   
    
    Set gActiveElement = document.ActiveElement   
End Function
'=========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
        .focus
        ggoSpread.Source = frm1.vspdData 
        lDelRows = ggoSpread.DeleteRow
    End With
End Function
'=========================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'=========================================================================================================
Function FncPrev() 
End Function
'=========================================================================================================
Function FncNext() 
End Function
'=========================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function
'=========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)
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
Function FncExit()
	Dim IntRetCD
    
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    FncExit = True
End Function
'=========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Call LayerShowHide(1)    
    
    Err.Clear

    Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
        strVal = strVal & "&txtMnuID=" & Trim(.hMnuID.value)
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
        strVal = strVal & "&txtLangCd=" & Parent.gLang
        strVal = strVal & "&txtMnuID=" & Trim(.txtMnuID.value)
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
    End If
    
    Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True
    
    Call SetToolbar("11001001000111")
    
End Function
'=========================================================================================================
Function DbQueryOk()
    
    lgIntFlgMode = Parent.OPMD_UMODE
    
    Call ggoOper.LockField(Document, "Q")

End Function
'=========================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
    Dim iColSep, iRowSep
    Dim Cntlen
    Dim ConvSeq

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep
        
    DbSave = False
    
    Call LayerShowHide(1)    
    
    On Error Resume Next

    With frm1
        .txtMode.value = Parent.UID_M0002
        .txtUpdtUserId.value = Parent.gUsrID
        .txtInsrtUserId.value = Parent.gUsrID
        
        lGrpCnt = 1
    
        strVal = ""
        strDel = ""
    
        For lRow = 1 To .vspdData.MaxRows
            Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")

                Case ggoSpread.InsertFlag
                    
                    strVal = strVal & "C" & iColSep & lRow & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuID, lRow, "X", "X")) & iColSep            
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuType, lRow, "X", "X")) & iColSep
                    Cntlen =len(Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X")))					
                        Select Case Cntlen
                        case 0 
							ConvSeq = "000"
                        Case 1
                            ConvSeq = "00" & Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X"))
                        Case 2
                            ConvSeq = "0" & Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X"))                        
                        Case else
                            ConvSeq = Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X"))
                        End Select
                    strVal = strVal & ConvSeqs & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_UseYN, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_UpperMnuID, lRow, "X", "X")) & iRowSep
                    
                    lGrpCnt = lGrpCnt + 1

                Case ggoSpread.UpdateFlag

                    strVal = strVal & "T" & iColSep & lRow & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuID_ORG, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuNm, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuType, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_SysLvl, lRow, "X", "X")) & iColSep
					Cntlen =len(Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X")))					
                        Select Case Cntlen
                        case 0 
							ConvSeq = "000"
                        Case 1
                            ConvSeq = "00" & Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X"))
                        Case 2
                            ConvSeq = "0" & Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X"))                        
                        Case else
                            ConvSeq = Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X"))
                        End Select
                    strVal = strVal & ConvSeq & iRowSep
                    
                    lGrpCnt = lGrpCnt + 1
                    
                Case ggoSpread.DeleteFlag

                    strDel = strDel & "D" & iColSep & lRow & iColSep
                    strDel = strDel & Trim(GetSpreadText(.vspdData, C_MnuID, lRow, "X", "X")) & iColSep
                    strDel = strDel & Trim(GetSpreadText(.vspdData, C_MnuType, lRow, "X", "X")) & iColSep                  
                    strDel = strDel & Trim(GetSpreadText(.vspdData, C_UpperMnuID, lRow, "X", "X")) & iRowSep
                    
                    lGrpCnt = lGrpCnt + 1
                    
            End Select
                    
        Next
    
        .txtMaxRows.value = lGrpCnt-1
        .txtSpread.value = strDel & strVal
    
        Call ExecMyBizASP(frm1, BIZ_PGM_ID)
    
    End With
    
    DbSave = True
    
End Function
'=========================================================================================================
Function DbSaveOk()
   
    Call InitVariables
    frm1.vspdData.MaxRows = 0    

    Call MainQuery()

End Function

'=========================================================================================================
Function DbDelete() 
End Function
'=========================================================================================================

Function CheckNumeric(ByVal strNum) 
  Dim Ret
  Dim intlen, intCnt, intAsc

  intlen = len(strNum)

  for intCnt = 1 to intlen

      intAsc = asc(mid(strNum, intCnt, 1))

      if intAsc < 48 or intAsc > 57  then
         CheckNumeric = 1
         Exit function
      end if
  next

End Function
'=========================================================================================================
Function OpenMnuInfo(Byval strCode)

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True
    
    arrParam(0) = "메뉴 팝업"
    arrParam(1) = "Z_USR_MNU"
    arrParam(2) = strCode
    arrParam(3) = ""
    arrParam(4) = "LANG_CD =  " & FilterVar(Parent.gLang , "''", "S") & " AND USR_ID = " & FilterVar(Parent.gUsrID, "''", "S")    
        
    arrParam(5) = "메뉴ID"
    
    arrField(0) = "SUBSTRING(MNU_ID,1,CHARINDEX(" & FilterVar("^", "''", "S") & ",MNU_ID)-1)"
    arrField(1) = "MNU_NM"
    
    arrHeader(0) = "메뉴ID"
    arrHeader(1) = "메뉴명"
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetMnuInfo(arrRet)
    End If    

    frm1.txtMnuID.focus 
    Set gActiveElement = document.activeElement
    
End Function

'=========================================================================================================
Function SetMnuInfo(Byval arrRet)

    frm1.txtMnuID.Value    = arrRet(0)
    frm1.txtMnuNm.Value    = arrRet(1)
    frm1.txtMnuID.focus    

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->    
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BasicTB" CELLSPACING=0>
    <TR>
        <TD HEIGHT=5>&nbsp;</TD>
    </TR>
    <TR HEIGHT=23>
        <TD WIDTH=100%>
            <TABLE CLASS="BasicTB" CELLSPACING=0>
                <TR>
                    <TD WIDTH=10>&nbsp;</TD>
                    <TD CLASS="CLSMTABP">
                        <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
                            <TR>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>사용자 메뉴 관리</font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=*>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR HEIGHT=*>
        <TD WIDTH=100% CLASS="Tab11">
            <TABLE CLASS="BasicTB" CELLSPACING=0>
                <TR>
                    <TD HEIGHT=5 WIDTH=100%></TD>
                </TR>
                <TR>
                    <TD HEIGHT=20 WIDTH=100%>
                    <FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                    <TR>
                        <TD CLASS="TD5">메뉴 ID</TD>
                        <TD CLASS="TD6"><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtMnuID" SIZE=15 MAXLENGTH=15 tag="11XXXU"  ALT="메뉴 ID"><IMG SRC="../../../CShared/image/btnPopup.gif"   NAME="btnMnuID" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMnuInfo frm1.txtMnuID.value ">&nbsp;<INPUT TYPE=TEXT NAME="txtMnuNm" SIZE=40 tag="14"></TD>
                        <TD CLASS="TDT"></TD>
                        <TD CLASS="TD6"></TD>
                    </TR>
                </TABLE></FIELDSET></TD>
            </TR>
            <TR>
                <TD WIDTH=100% HEIGHT=* valign=top><TABLE WIDTH="100%" HEIGHT="100%">
                    <TR>
                        <TD HEIGHT="100%">
                        <script language =javascript src='./js/zc00lma1_I180927181_vspdData.js'></script></TD>
                    </TR></TABLE>
                </TD>
            </TR>
        </TABLE></TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hLangCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hMnuID" tag="24">
<INPUT TYPE=HIDDEN NAME="hMnuType" tag="24">
<INPUT TYPE=HIDDEN NAME="hUseYN" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
