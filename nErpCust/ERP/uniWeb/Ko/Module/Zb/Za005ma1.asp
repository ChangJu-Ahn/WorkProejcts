<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za005ma1
*  4. Program Name         : Table Information Management
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000.03.13
*  8. Modified date(Last)  : 2002.06.03
*  9. Modifier (First)     : LeeJaeJoon
* 10. Modifier (Last)      : LeeJaeWan
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
'=========================================================================================================
Const BIZ_PGM_ID = "Za005mb1.asp"            
'=========================================================================================================
Dim C_TableId
Dim C_TablePopup
Dim C_TableNm
Dim C_Mod
Dim C_Mod_Nm
Dim C_Type
Dim C_TypeNm
Dim C_UseYn



'=========================================================================================================
Dim IsOpenPop        

<!-- #Include file="../../inc/lgvariables.inc" -->    


'=========================================================================================================
Sub InitSpreadPosVariables()
    C_TableId    = 1
    C_TablePopup = 2
    C_TableNm    = 3
    C_Mod        = 4
    C_Mod_Nm     = 5
    C_Type       = 6
    C_TypeNm     = 7
    C_UseYn         = 8
End Sub

       
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   

    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           
    lgLngCurRows = 0                            
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

    Call InitSpreadPosVariables()
    
    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)

        .ReDraw = false                   
        .MaxCols = C_UseYn + 1                            
        .MaxRows = 0

        Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetEdit C_TableId    , "테이블 ID"        , 35, , ,30
        ggoSpread.SSSetButton  C_TablePopup    ,-1
        ggoSpread.SSSetEdit C_TableNm    , "테이블 명"        , 35, , ,40
        ggoSpread.SSSetCombo C_Mod        ,"사용모듈"            , 15
        ggoSpread.SSSetCombo C_Mod_Nm    , "사용모듈명"        , 20
        ggoSpread.SSSetEdit  C_Type        , "테이블 속성코드"    , 5
        ggoSpread.SSSetCombo C_TypeNm    , "테이블 속성"        , 20 '4
        ggoSpread.SSSetCheck C_UseYn    , "사용 유무"        , 10, 2, "", True '4
        'ggoSpread.SSSetSplit2(2)
        .ReDraw = true

        Call SetSpreadLock

        Call ggoSpread.MakePairsColumn(C_TableId,C_TableNm,"1")
        Call ggoSpread.MakePairsColumn(C_Mod,C_Mod_Nm,"1")
        Call ggoSpread.MakePairsColumn(C_Type,C_TypeNm,"1")

        Call ggoSpread.SSSetColHidden(C_Mod,C_Mod,True)
        Call ggoSpread.SSSetColHidden(C_Type,C_Type,True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
                           
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_TableId   , -1, C_TableId
        ggoSpread.SpreadLock C_TablePopup, -1, C_TablePopup
        
        ggoSpread.SSSetRequired    C_TableNm, -1, -1
        ggoSpread.SSSetRequired    C_Mod     , -1, -1
        ggoSpread.SSSetRequired    C_Mod_Nm , -1, -1
        ggoSpread.SSSetRequired    C_TypeNm , -1, -1
        ggoSpread.SSSetRequired    C_UseYn     , -1, -1
        .vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired C_TableId, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_TableNm, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_Mod     , pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_TypeNm , pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_UseYn  , pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_Mod_Nm , pvStartRow, pvEndRow
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

            C_TableId      =  iCurColumnPos(1)
            C_TablePopup   =  iCurColumnPos(2)
            C_TableNm      =  iCurColumnPos(3)
            C_Mod          =  iCurColumnPos(4)
            C_Mod_Nm       =  iCurColumnPos(5)
            C_Type         =  iCurColumnPos(6)
            C_TypeNm       =  iCurColumnPos(7)
            C_UseYn        =  iCurColumnPos(8)

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

    Dim IntRetCD
    Dim iPos0, iPos1
    
    On Error Resume Next
    
    '------ Developer Coding part (Start ) --------------------------------------------------------------     
    '---------------------------------------------------------
	Dim StrMod
    Dim StrLen
    Dim StrSql
    Dim StrMd
    Dim i 
        
    StrMod = Trim(Parent.gSetupMod)	      
    StrLen = Len(StrMod)
    
    If StrLen > 0 Then
        For i = 1 To StrLen
            StrMd = Mid(StrMod, i, 1)
            If i =1 Then 
				StrSql = StrSql + "and ( MINOR_CD=" & FilterVar("*", "''", "S") & "  or"
            ElseIf i > 1 Then
                StrSql = StrSql + "or"
            End If
            
            If StrMd="A" Then 
				StrSql = StrSql + " MINOR_CD = " & FilterVar("F", "''", "S") & "  or "
            End If
            StrSql = StrSql + " MINOR_CD = " & FilterVar(StrMd, "''", "S") & " "
        Next
        StrSql = StrSql + ")"
    End If
	'---------------------------------------------------------
    ggoSpread.Source = frm1.vspdData
    
    IntRetCD = CommonQueryRs("minor_cd, minor_nm", "B_Minor", "major_cd = " & FilterVar("B0001", "''", "S") & "" & StrSql, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = Replace(lgF0, Chr(11), vbTab)
    lgF1 = Replace(lgF1, Chr(11), vbTab)
    iPos0 = InStr(lgF0,vbTab) + 1
    iPos1 = InStr(lgF1,vbTab) + 1
    lgF0 = Mid(lgF0, iPos0)
    lgF1 = Mid(lgF1, iPos1)
    ggoSpread.SetCombo lgF0, C_Mod
    ggoSpread.SetCombo lgF1, C_Mod_Nm
    
    IntRetCD = CommonQueryRs("minor_nm","B_Minor","major_cd = " & FilterVar("Z0009", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = Replace(lgF0, Chr(11), vbTab)
    ggoSpread.SetCombo lgF0, C_TypeNm
    
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'=========================================================================================================
Sub Form_Load()

    Call ggoOper.LockField(Document, "N")                                   
    
    Call InitSpreadSheet                                                    
    Call InitVariables                                                      

    '----------  Coding part  -------------------------------------------------------------
    Call InitSpreadComboBox
    Call SetDefaultVal
    Call SetToolBar("11001101001011")                                        

    frm1.txtTable.focus 
    Set gActiveElement = document.activeElement
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub
'=========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)
    With frm1.vspdData 
        ggoSpread.Source = frm1.vspdData
        .Row = Row
        If Row > 0 Then
            Select Case Col
            Case C_TablePopup
                Call OpenTableCode(GetSpreadText(frm1.vspdData, Col - 1, Row, "X", "X"),Row)
            End Select
        End If
    
    End With
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
Sub vspdData_Change(ByVal Col, ByVal Row)
    Dim iDx
       
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

    Select Case Col
	Case  C_Mod_Nm
		iDx = Frm1.vspdData.value
		Frm1.vspdData.Col = C_Mod
		Frm1.vspdData.value = iDx
	Case Else
    End Select    

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
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
Sub vspdData_KeyPress(index , KeyAscii )
     lgBlnFlgChgValue = True                                                 
End Sub

'=========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If CheckRunningBizProcess = True Then
       Exit Sub
    End If

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
       Call DisableToolBar(Parent.TBC_QUERY)                                                       

       If DbQuery = False Then                                                       
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
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")                   '데이타가 변경되었습니다. 조회하시겠습니까?
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
    With frm1
        If .vspdData.ActiveRow > 0 Then
            .vspdData.focus
            .vspdData.ReDraw = False
        
            ggoSpread.Source = frm1.vspdData    
            ggoSpread.CopyRow
            nActiveRow = .vspdData.ActiveRow
            SetSpreadColor nActiveRow, nActiveRow
            .vspdData.SetText C_TableId, nActiveRow, ""
            
            .vspdData.ReDraw = True
        End If
    End With
End Function

'=========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData    
    ggoSpread.EditUndo                                                  
End Function

'=========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim nActiveRow
    
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
        nActiveRow = .vspdData.ActiveRow
        SetSpreadColor nActiveRow, nActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    frm1.vspdData.SetText C_USEYN, nActiveRow, "Y"
    '------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncInsertRow = True                                                              
    End If   
    
    Set gActiveElement = document.ActiveElement   
End Function


'=========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    frm1.vspdData.focus
    ggoSpread.Source = frm1.vspdData 

    lDelRows = ggoSpread.DeleteRow
End Function

'=========================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'=========================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    
End Function

'=========================================================================================================
Function FncNext() 
    On Error Resume Next                                                    
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
    Dim Strmode      'khy200308  

    DbQuery = False
    
    Err.Clear                
    
    Strmode = Trim(Parent.gSetupMod)     'khy200308                                                         

    With frm1
        If lgIntFlgMode = Parent.OPMD_UMODE Then
                strVal = BIZ_PGM_ID & "?txtMode="    & Parent.UID_M0001                            
                strVal = strVal     & "&txtMaxRows=" & .vspdData.MaxRows                    
                strVal = strVal     & "&txtCode="    & .htxtTable.value                     
         
                strVal = strVal & "&txtChk1="         & .hchk1.value
                strVal = strVal & "&txtChk2="        & .hchk2.value 
                strVal = strVal & "&txtChk3="        & .hchk3.value 
                strVal = strVal & "&txtChk4="        & .hchk4.value 
                strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey       
                strVal = strVal & "&Strmode="   & Strmode      'khy200308           
        Else
                strVal = BIZ_PGM_ID & "?txtMode="    & Parent.UID_M0001                     
                strVal = strVal     & "&txtMaxRows=" & .vspdData.MaxRows                    
                strVal = strVal     & "&txtCode="    & Trim(.txtTable.value)                
         
                If .chk1.Checked Then
                    strVal = strVal & "&txtChk1=S"
                End If
                
                If .chk2.Checked Then
                    strVal = strVal & "&txtChk2=M"
                End If
                If .chk3.Checked Then
                    strVal = strVal & "&txtChk3=X"
                End If
                If .chk4.Checked Then
                    strVal = strVal & "&txtChk4=T"
                End If
                strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
                strVal = strVal & "&Strmode="   & Strmode      'khy200308        
        End If
    
        Call LayerShowHide(1)
        Call RunMyBizASP(MyBizASP, strVal)                                        
    End With
    
    DbQuery = True
End Function
'=========================================================================================================
Function DbQueryOk()                                                        
    
    lgIntFlgMode = Parent.OPMD_UMODE                                                
    
    Call ggoOper.LockField(Document, "Q")                                    
    Call SetToolBar("11001111001111")
    frm1.vspdData.Focus
End Function

'=========================================================================================================
Function DbSave()                                             
    Dim lRow        
    Dim lGrpCnt     
    Dim strVal, strDel
    Dim a, b
    Dim iColSep, iRowSep
    iColSep = parent.gColSep
    iRowSep = parent.gRowSep
        
    On Error Resume Next                                                   
    DbSave = False                                                          

    Call LayerShowHide(1)
    
    With frm1
        .txtMode.value = Parent.UID_M0002
    

        lGrpCnt = 1
        strVal = ""
        strDel = ""
    

        For lRow = 1 To .vspdData.MaxRows
            Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")
                      
                  Case ggoSpread.InsertFlag                                      
                                                    strVal = strVal & "C"                       & iColSep
                                                    strVal = strVal & lRow                      & iColSep
								                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_TableId, lRow, "X", "X"))      & iColSep
								                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_TableNm, lRow, "X", "X"))      & iColSep
								                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_Mod, lRow, "X", "X"))      & iColSep
								                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_Type, lRow, "X", "X"))      & iColSep
								                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_UseYn, lRow, "X", "X"))      & iRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      
                                                    strVal = strVal & "U"                       & iColSep
                                                    strVal = strVal & lRow                      & iColSep
                                                    strVal = strVal & Trim(.txtProfileCd.value) & iColSep
								                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_TableId, lRow, "X", "X"))      & iColSep
								                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_TableNm, lRow, "X", "X"))      & iColSep
								                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_Mod, lRow, "X", "X"))      & iColSep
								                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_Type, lRow, "X", "X"))      & iColSep
								                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_UseYn, lRow, "X", "X"))      & iRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.DeleteFlag                                      
                                                    strDel = strDel & "D"                       & iColSep
                                                    strDel = strDel & lRow                      & iColSep
                    								strDel = strDel & Trim(GetSpreadText(.vspdData, C_TableId, lRow, "X", "X"))      & iRowSep
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
    DBDelete = False

    DBDelete = True
End Function


'=========================================================================================================
'    Name : OpenTable()
'    Description : Table PopUp
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
        
    StrMod = Trim(Parent.gSetupMod)	      
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
	
    arrParam(0) = "테이블 팝업"						' 팝업 명칭 
    arrParam(1) = "z_table_info"                        ' TABLE 명칭 
    arrParam(2) = frm1.txtTable.value					' Code Condition
    arrParam(3) = ""									' Name Cindition
    arrParam(4) = "lang_cd =  " & FilterVar(gLang , "''", "S") & "" & strsql	' Where Condition
    arrParam(5) = "테이블"							' 조건필드의 라벨 명칭 
    
    arrField(0) = "ED24" & Parent.gColSep & "table_id"  ' Field명(0)
    arrField(1) = "table_nm"							' Field명(1)
    
    arrHeader(0) = "테이블 ID"						' Header명(0)
    arrHeader(1) = "테이블 명"						' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=520px; dialogHeight=455px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetTable(arrRet)
    End If    
    frm1.txtTable.focus
    Set gActiveElement = document.activeElement
End Function

'========================================================================================================
' Name : OpenTableCode()
' Desc : developer describe this line 
'========================================================================================================
Function OpenTableCode(TableCode,Row)

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "테이블 팝업"
    arrParam(1) = "sysobjects"
    arrParam(2) = TableCode       
    arrParam(3) = ""              
    arrParam(4) = "xtype = " & FilterVar("U", "''", "S") & " " & vbCrLf & _
                  "and name not in(select table_id" & vbCrLf & _
                  "                from z_table_info" & vbCrLf & _
                  "                where lang_cd =  " & FilterVar(gLang , "''", "S") & ")"
    arrParam(5) = "테이블"
    
    arrField(0) = "name"          
    arrField(1) = "name"          
    
    arrHeader(0) = "테이블 ID"
    arrHeader(1) = "테이블 명"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SubSetTableCode(arrRet,Row)
    End If    
    
End Function
'=========================================================================================================
Sub SubSetTableCode(arrRet,Row)
    With frm1.vspdData 
    	.SetText C_TableId, Row, arrRet(0)
    	.SetText C_TableNm, Row, arrRet(1)
    End With
End Sub

'=========================================================================================================
Function OpenTableProgram()
    Dim arrRet, IntRetCD
    Dim arrParam, arrField, arrHeader
    Dim iCalledAspName
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZA005RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZA005RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
    
    
    If frm1.vspdData.MaxRows = 0 Then
        IntRetCD = DisplayMsgBox("900002", "x", "x", "x")
        frm1.txtTable.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False
        Exit Function
    End If
    
    
    With frm1
        arrParam = GetSpreadText(.vspdData, C_TableId, .vspdData.ActiveRow, "X", "X")
    End With

    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=730px; dialogHeight=465px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False
End Function
'=========================================================================================================
Function OpenTableLayout()
    Dim arrRet, IntRetCD
    Dim arrParam, arrField, arrHeader
    Dim iCalledAspName
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZA005RA2")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZA005RA2", "x")
        IsOpenPop = False
        Exit Function
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        IntRetCD = DisplayMsgBox("900002", "x", "x", "x")
        frm1.txtTable.focus
        Set gActiveElement = document.activeElement
        IsOpenPop = False        
        Exit Function
    End If

    With frm1
        arrParam = GetSpreadText(.vspdData, C_TableId, .vspdData.ActiveRow, "X", "X")
    End With

    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=770px; dialogHeight=465px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False
End Function


'=========================================================================================================
'    Name : SetLogonGp()
'    Description : Country Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetTable(Byval arrRet)
    With frm1
        .txtTable.value = arrRet(0)
        .txtTableNm.value = arrRet(1)
    End With
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc" -->    
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
    <TR>
        <TD <%=HEIGHT_TYPE_00%>></TD>
    </TR>
    <TR HEIGHT=23>
        <TD WIDTH=100%>
            <TABLE <%=LR_SPACE_TYPE_10%>>
                <TR>
                    <TD WIDTH=10>&nbsp;</TD>
                    <TD CLASS="CLSMTABP">
                        <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
                            <TR>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>테이블 정보 관리</font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* align=right>
                        <a href="vbscript:OpenTableProgram">테이블별 프로그램 매핑현황</A>
                        &nbsp;|&nbsp;
                        <a href="vbscript:OpenTableLayout">테이블 레이아웃</A>
                    </TD>
                    <TD WIDTH=10>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR HEIGHT=*>
        <TD WIDTH=100% CLASS="Tab11">
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
                                        <INPUT TYPE=TEXT NAME="txtTable" SIZE=30 MAXLENGTH=30 tag="11" ALT="테이블 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btntablepop" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenTable">&nbsp;&nbsp;
                                        <INPUT TYPE=TEXT NAME="txtTableNm" size=40 MAXLENGTH=40 tag="14X">
                                    </TD>
                                </TR>
                                <TR>
                                    <TD CLASS="TD5">테이블 속성</TD>
                                    <TD CLASS="TD6">
                                        <INPUT type="checkbox" NAME="chk1" ID="chk1" tag="11" checked class=check><LABEL FOR="chk1">System 테이블</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;
                                        <INPUT type="checkbox" NAME="chk2" ID="chk2" tag="11" checked class=check><LABEL FOR="chk2">Master 테이블</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;
                                        <INPUT type="checkbox" NAME="chk3" ID="chk3" tag="11" checked class=check><LABEL FOR="chk3">Transaction 테이블</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;
                                        <INPUT type="checkbox" NAME="chk4" ID="chk4" tag="11" checked class=check><LABEL FOR="chk4">Temp 테이블</LABEL>
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
                    <TD WIDTH=100% HEIGHT=100% valign=top>
                        <TABLE <%=LR_SPACE_TYPE_20%>>
                            <TR>
                                <TD HEIGHT="100%">
                                    <script language =javascript src='./js/za005ma1_vaSpread1_vspdData.js'></script>
                                </TD>
                            </TR>
                        </TABLE>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="Za005mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtTable" tag="24">
<INPUT TYPE=HIDDEN NAME="hchk1" tag="24">
<INPUT TYPE=HIDDEN NAME="hchk2" tag="24">
<INPUT TYPE=HIDDEN NAME="hchk3" tag="24">
<INPUT TYPE=HIDDEN NAME="hchk4" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

