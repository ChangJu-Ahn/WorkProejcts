
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit     


Const BIZ_PGM_ID = "ZC050MB1.asp"
Const JUMP_PGM_ID = "zc055ma1"

Dim C_MNU_ID
Dim C_MNU_NM
Dim C_ALLOW_YN        
Dim C_BIZ_AREA_YN     
Dim C_INTERNAL_YN     
Dim C_SUB_INTERNAL_YN 
Dim C_PERSONAL_YN       
Dim C_PLANT_YN        
Dim C_PUR_ORG_YN      
Dim C_PUR_GRP_YN      
Dim C_SALES_ORG_YN    
Dim C_SALES_GRP_YN    
Dim C_SL_YN           
Dim C_WC_YN           

Dim C_DUMMY

<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim IsOpenPop
'=========================================================================================================
Sub InitSpreadPosVariables()
    C_MNU_ID           = 1
    C_MNU_NM           = 2
    C_ALLOW_YN         = 3
    C_BIZ_AREA_YN      = 4
    C_INTERNAL_YN      = 5
    C_SUB_INTERNAL_YN  = 6
    C_PERSONAL_YN      = 7
    C_PLANT_YN         = 8
    C_PUR_ORG_YN       = 9
    C_PUR_GRP_YN       = 10
    C_SALES_ORG_YN     = 11
    C_SALES_GRP_YN     = 12
    C_SL_YN            = 13
    C_WC_YN            = 14
    C_DUMMY            = 15
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
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE","QA") %>
End Sub

'=========================================================================================================
Sub InitSpreadSheet()

    On Error Resume Next

    Call InitSpreadPosVariables()
    
    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)

        .ReDraw = false                   
        .MaxCols = C_DUMMY
        .MaxRows = 0

        Call GetSpreadColumnPos("A")

        frm1.vspdData.RowHeight(0) = 20
       
        ggoSpread.SSSetEdit   C_MNU_ID           , "메뉴ID"             ,12
        ggoSpread.SSSetEdit   C_MNU_NM           , "메뉴명"             ,18
        ggoSpread.SSSetCheck  C_ALLOW_YN        , "메뉴권한여부"       ,12,,,True
        ggoSpread.SSSetCheck  C_BIZ_AREA_YN     , "사업장"             ,12,,,True
        ggoSpread.SSSetCheck  C_INTERNAL_YN     , "내부부서"           ,10,,,True
        ggoSpread.SSSetCheck  C_SUB_INTERNAL_YN , "내부부서" & vbCrLf & "(하위포함)" ,12,,,True
        ggoSpread.SSSetCheck  C_PERSONAL_YN     , "개인"               , 8,,,True
        ggoSpread.SSSetCheck  C_PLANT_YN        , "공장"               , 6,,,True
        ggoSpread.SSSetCheck  C_PUR_ORG_YN      , "구매조직"           ,17,,,True
        ggoSpread.SSSetCheck  C_PUR_GRP_YN      , "구매그룹"           ,13,,,True
        ggoSpread.SSSetCheck  C_SALES_ORG_YN    , "영업조직"           , 8,,,True
        ggoSpread.SSSetCheck  C_SALES_GRP_YN    , "영업그룹"           ,10,,,True
        ggoSpread.SSSetCheck  C_SL_YN           , "창고"               ,13,,,True
        ggoSpread.SSSetCheck  C_WC_YN           , "작업장"             ,10,,,True

        .ReDraw = true

        Call SetSpreadLock

        Call ggoSpread.SSSetColHidden(C_DUMMY, C_DUMMY, True)

    End With
    
    
End Sub
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
    
        .vspdData.ReDraw = False
        
        ggoSpread.SpreadLock    C_MNU_ID , -1, C_MNU_ID
        ggoSpread.SpreadLock    C_MNU_NM , -1, C_MNU_NM

        .vspdData.ReDraw = True    

    End With
End Sub
'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired        C_MNU_ID, pvStartRow, pvEndRow        
        ggoSpread.SSSetRequired        C_MNU_NM, pvStartRow, pvEndRow
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
    
            C_MNU_ID            = iCurColumnPos(1)
            C_MNU_NM            = iCurColumnPos(2)
            C_ALLOW_YN         = iCurColumnPos(3)
            C_BIZ_AREA_YN      = iCurColumnPos(4)
            C_INTERNAL_YN      = iCurColumnPos(5)
            C_SUB_INTERNAL_YN  = iCurColumnPos(6)
            C_PERSONAL_YN      = iCurColumnPos(7)
            C_PLANT_YN         = iCurColumnPos(8)
            C_PUR_ORG_YN       = iCurColumnPos(9)
            C_PUR_GRP_YN       = iCurColumnPos(10)
            C_SALES_ORG_YN     = iCurColumnPos(11)
            C_SALES_GRP_YN     = iCurColumnPos(12)
            C_SL_YN            = iCurColumnPos(13)
            C_WC_YN            = iCurColumnPos(14)
            C_DUMMY            = iCurColumnPos(15)

    End Select
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    
    If IsNumeric(iPosArr) Then
       iRow = CInt(iPosArr)
       
       If iRow <=0 Then
          Exit Sub
       End if
       
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If           
       Next          
    End If   
End Sub

'=========================================================================================================
Sub InitComboBox()

End Sub

'=========================================================================================================
Sub InitSpreadComboBox()

    Dim strCboData
    Dim IntRetCD

    ggoSpread.Source = frm1.vspdData

End Sub
'=========================================================================================================
Sub Form_Load()

    Dim IntRetCD
    
    On Error Resume Next
    
    Call ggoOper.LockField(Document, "N")

    Call InitSpreadSheet
    
    Call InitVariables

    Call InitComboBox
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
    Call SetPopupMenuItemInf("1101111111")    
    
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
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

    Dim strTemp
    Dim intPos1
   
End Sub

'=========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    
    If CheckRunningBizProcess = True Then
       Exit Sub
    End If
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) _
    And Not(lgStrPrevKey = "") Then    
        Call DisableToolBar(Parent.TBC_QUERY)
        If DBQuery = False Then
            Call RestoreToolBar()
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
                                                                
    If Not chkField(Document, "1") Then
       Exit Function
    End If

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
	Dim nActiveRow
    With frm1.vspdData
        If .ActiveRow > 0 Then
            .focus
            .ReDraw = False
            
            ggoSpread.Source = frm1.vspdData 
            ggoSpread.CopyRow
            nActiveRow = frm1.vspdData.ActiveRow
            SetSpreadColor nActiveRow, nActiveRow
    
    		frm1.vspdData.SetText C_MNU_ID, nActiveRow, ""
            .ReDraw = True
        End If
    End With
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

    Dim IntRetCD

    DbQuery = False
    
    Call LayerShowHide(1)    
    
    Err.Clear

    Dim strVal    
    
    With frm1
    
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then    
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
        
        strVal = strVal & "&txtLangCd=" & parent.gLang
        strVal = strVal & "&txtMnuID=" & Trim(.hMnuID.value)
        strVal = strVal & "&cboMnuType=P"
        strVal = strVal & "&cboUseYN=1"   ' 1 : 사용  , 0 : 미사용 
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    Else    
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
        
        strVal = strVal & "&txtLangCd=" & parent.gLang
        strVal = strVal & "&txtMnuID=" & Trim(.txtMnuID.value)
        strVal = strVal & "&cboMnuType=P"
        strVal = strVal & "&cboUseYN=1"
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    End If   
    
    
    Call RunMyBizASP(MyBizASP, strVal)
    
    End With
       
    
    DbQuery = True
    
End Function
'=========================================================================================================
Function DbQueryOk()
    
    lgIntFlgMode = Parent.OPMD_UMODE
    
    Call ggoOper.LockField(Document, "Q")
    Call SetToolbar("11001000000111")

    Call AutoHWidth(frm1.vspdData)
    
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
    Dim IntRetCD
    
    Dim iALLOW_YN       
    Dim iBIZ_AREA_YN    
    Dim iINTERNAL_YN    
    Dim iSUB_INTERNAL_YN
    Dim iPERSONAL_YN    
    Dim iPLANT_YN       
    Dim iPUR_ORG_YN     
    Dim iPUR_GRP_YN     
    Dim iSALES_ORG_YN   
    Dim iSALES_GRP_YN   
    Dim iSL_YN          
    Dim iWC_YN  
    
    Dim strValT        
    
    iColSep = parent.gColSep
    iRowSep = parent.gRowSep
        
    DbSave = False
    
    Call LayerShowHide(1)    
    
    On Error Resume Next

    With frm1
        .txtMode.value        = Parent.UID_M0002
        .txtUpdtUserId.value  = Parent.gUsrID
        .txtInsrtUserId.value = Parent.gUsrID
        
        lGrpCnt = 1
    
        strVal = ""
        strDel = ""
        
        For lRow = 1 To .vspdData.MaxRows
            Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")

                Case ggoSpread.UpdateFlag

                    strVal = strVal & "U"                                                               & iColSep   '0
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MNU_ID          , lRow, "X", "X")) & iColSep   '1
                    strVal = strVal & "P"                                                               & iColSep   '2
                    
                    
                    iALLOW_YN        = Trim(GetSpreadText(.vspdData, C_ALLOW_YN       , lRow, "X", "X"))
                    iBIZ_AREA_YN     = Trim(GetSpreadText(.vspdData, C_BIZ_AREA_YN    , lRow, "X", "X"))
                    iINTERNAL_YN     = Trim(GetSpreadText(.vspdData, C_INTERNAL_YN    , lRow, "X", "X"))
                    iSUB_INTERNAL_YN = Trim(GetSpreadText(.vspdData, C_SUB_INTERNAL_YN, lRow, "X", "X"))
                    iPERSONAL_YN     = Trim(GetSpreadText(.vspdData, C_PERSONAL_YN    , lRow, "X", "X"))
                    iPLANT_YN        = Trim(GetSpreadText(.vspdData, C_PLANT_YN       , lRow, "X", "X"))
                    iPUR_ORG_YN      = Trim(GetSpreadText(.vspdData, C_PUR_ORG_YN     , lRow, "X", "X"))
                    iPUR_GRP_YN      = Trim(GetSpreadText(.vspdData, C_PUR_GRP_YN     , lRow, "X", "X"))
                    iSALES_ORG_YN    = Trim(GetSpreadText(.vspdData, C_SALES_ORG_YN   , lRow, "X", "X"))
                    iSALES_GRP_YN    = Trim(GetSpreadText(.vspdData, C_SALES_GRP_YN   , lRow, "X", "X"))
                    iSL_YN           = Trim(GetSpreadText(.vspdData, C_SL_YN          , lRow, "X", "X"))
                    iWC_YN           = Trim(GetSpreadText(.vspdData, C_WC_YN          , lRow, "X", "X"))

                    
                    if iALLOW_YN = "1" then

                       strValT =           iBIZ_AREA_YN     
                       strValT = strValT & iINTERNAL_YN     
                       strValT = strValT & iSUB_INTERNAL_YN 
                       strValT = strValT & iPERSONAL_YN     
                       strValT = strValT & iPLANT_YN        
                       strValT = strValT & iPUR_ORG_YN      
                       strValT = strValT & iPUR_GRP_YN      
                       strValT = strValT & iSALES_ORG_YN    
                       strValT = strValT & iSALES_GRP_YN    
                       strValT = strValT & iSL_YN           
                       strValT = strValT & iWC_YN 
                       
                       if instr(strValT,"1") = 0 then
                          Call DisplayMsgBox("214018", "X", "x", "x")
                          Call LayerShowHide(0)
                          Call SubSetErrPos(lRow)
                          exit function
                       end if


                       
                       strVal = strVal & iALLOW_YN        & iColSep   '3
                       strVal = strVal & iBIZ_AREA_YN     & iColSep   '5
                       strVal = strVal & iINTERNAL_YN     & iColSep   '6
                       strVal = strVal & iSUB_INTERNAL_YN & iColSep   '7
                       strVal = strVal & iPERSONAL_YN     & iColSep   '8
                       strVal = strVal & iPLANT_YN        & iColSep   '9
                       strVal = strVal & iPUR_ORG_YN      & iColSep   '10
                       strVal = strVal & iPUR_GRP_YN      & iColSep   '11
                       strVal = strVal & iSALES_ORG_YN    & iColSep   '12
                       strVal = strVal & iSALES_GRP_YN    & iColSep   '13
                       strVal = strVal & iSL_YN           & iColSep   '14
                       strVal = strVal & iWC_YN           & iColSep   '15

                    else
                                 
                       strVal = strVal & "0" & iColSep   '3
                       strVal = strVal & "0" & iColSep   '5
                       strVal = strVal & "0" & iColSep   '6
                       strVal = strVal & "0" & iColSep   '7
                       strVal = strVal & "0" & iColSep   '8
                       strVal = strVal & "0" & iColSep   '9
                       strVal = strVal & "0" & iColSep   '10
                       strVal = strVal & "0" & iColSep   '11
                       strVal = strVal & "0" & iColSep   '12
                       strVal = strVal & "0" & iColSep   '13
                       strVal = strVal & "0" & iColSep   '14
                       strVal = strVal & "0" & iColSep   '15
                       
                    end if
               
                    strVal = strVal & lRow                                                              & iRowSep   '16

                    '---------------------------------------------------------------------------------------
                    
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
Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZC004RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZC004RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
    
    With frm1
        If .vspdData.ActiveRow > 0 then 
            arrParam(0) = GetSpreadText(.vspdData, 3, .vspdData.ActiveRow, "X", "X")
            arrParam(1) = GetSpreadText(.vspdData, 4, .vspdData.ActiveRow, "X", "X")
        End If            
    End With    

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=550px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function



Function SetLangInfo(Byval arrRet)
	Dim nActiveRow

    With frm1.vspdData
    	nActiveRow = .ActiveRow
    	.SetText C_LangCD, nActiveRow, arrRet(0)
        Call vspdData_Change(C_LangCD, nActiveRow)
    End With

End Function
'==============================================================================================================
Function OpenMnuInfo(Byval strCode, Byval iWhere)

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    Dim IntRetCD    
    
    IntRetCD = CommonQueryRs("Lang_Nm","B_LANGUAGE","Lang_Cd =  " & FilterVar(parent.gLang, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    lgF0 = Replace(lgF0, Chr(11), "") 'unusual case
    'lgF0 = Replace(lgF0, " ","")    

    If lgF0 = "" then 
        Call DisplayMsgBox("211432", "x", "x", "x")
        Exit Function
    End if     

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "메뉴 팝업"
    arrParam(1) = "Z_LANG_CO_MAST_MNU"    
    arrParam(2) = strCode
    arrParam(3) = ""    
    
    Select Case iWhere
            Case  1
                arrParam(4) = "LANG_CD = " & FilterVar(Parent.gLang, "''", "S") & ""
            Case  2
                arrParam(4) = "LANG_CD = " & FilterVar(Parent.gLang, "''", "S") & ""
            Case  3
                arrParam(4) = "LANG_CD = " & FilterVar(Parent.gLang, "''", "S") & " AND MNU_TYPE = " & FilterVar("M", "''", "S") & " "
    End Select
                
    arrParam(5) = "메뉴ID"
    
    arrField(0) = "MNU_ID"
    arrField(1) = "MNU_NM"
    
    arrHeader(0) = "메뉴ID"
    arrHeader(1) = "메뉴명"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    	If iWhere = 1 Then
		    frm1.txtMnuID.focus
	        Set gActiveElement = document.activeElement
	    End If
        Exit Function
    Else
        Call SetMnuInfo(arrRet, iWhere)
    End If    

    
End Function

'=========================================================================================================
Function SetMnuInfo(Byval arrRet, Byval iWhere)
	Dim nActiveRow
    Select Case iWhere
        Case  1
            frm1.txtMnuID.Value    = arrRet(0)
            frm1.txtMnuNm.Value    = arrRet(1)
            frm1.txtMnuID.focus
            Set gActiveElement = document.activeElement
        Case  2
            With frm1.vspdData
            	nActiveRow = .ActiveRow
            	.SetText C_MNU_ID, nActiveRow, arrRet(0)
            	.SetText C_MNU_NM, nActiveRow, arrRet(1)
                Call vspdData_Change(C_MNU_NM, nActiveRow)
            End With
        Case  3
            With frm1.vspdData
            	nActiveRow = .ActiveRow
            	.SetText C_UpperMnuID, nActiveRow, arrRet(0)
                Call vspdData_Change(C_UpperMnuID, nActiveRow)
            End With
    End Select

End Function
'=========================================================================================================
Function ProgramJump

    Dim IntRetCD
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
          Exit Function
        End If
    End If
    Call PgmJump(JUMP_PGM_ID)
    
End Function

Sub AutoHWidth(pSpread)
    Dim iLoop
    
    For iLoop = 1 To pSpread.MaxCols
        pSpread.ColWidth(iLoop) = pSpread.MaxTextColWidth(iLoop) + 1
    Next

End Sub



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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* align=right>&nbsp;</TD>
                    <TD WIDTH=10>&nbsp;</TD>
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
                        <TD CLASS="TD656" colspan=3><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtMnuID" SIZE=15 MAXLENGTH=15 tag="11XXXU"  ALT="메뉴 ID"><IMG SRC="../../../CShared/image/btnPopup.gif"   NAME="btnMnuID" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMnuInfo frm1.txtMnuID.value,1 ">&nbsp;<INPUT TYPE=TEXT NAME="txtMnuNm" SIZE=40 tag="14"></TD>                        
                    </TR>
                </TABLE></FIELDSET></TD>
            </TR>
            <TR>
                <TD WIDTH=100% HEIGHT=* valign=top><TABLE WIDTH="100%" HEIGHT="100%">
                    <TR>
                        <TD HEIGHT="100%">
                        <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
                    </TR></TABLE>
                </TD>
            </TR>
        </TABLE></TD>
    </TR>
    <TR>
        <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>  
    <TR HEIGHT=20>   
        <TD WIDTH=100%>   
            <TABLE <%=LR_SPACE_TYPE_30%>>   
                <TR>   
                <TD WIDTH=50%>&nbsp;</TD>   
                <TD WIDTH=50%>   
                    <TABLE WIDTH=100%>                           
                        <TD WIDTH=* Align=right><A href="Vbscript:ProgramJump()">프로그램별자료권한(값)</A>&nbsp;</TD>                                                                                     
                        <TD WIDTH=10>&nbsp;</TD>                           
                    </TABLE>   
                </TD>   
                </TR>   
            </TABLE>   
        </TD>   
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
