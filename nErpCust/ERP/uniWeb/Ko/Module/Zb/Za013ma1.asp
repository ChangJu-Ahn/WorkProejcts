<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Basis Architect
*  2. Function Name        : System Management
*  3. Program ID           : za013ma1.asp
*  4. Program Name         : Object Management
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000/05/09
*  8. Modified date(Last)  : 2002/06/06
*  9. Modifier (First)     : ParkSangHoon
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
Const BIZ_PGM_ID = "za013mb1.asp"                                                
'=========================================================================================================

Dim C_hObjType
Dim C_ObjType
Dim C_ObjID
Dim C_ObjNm
Dim C_hModuleID
Dim C_ModuleID
Dim C_hRegType
Dim C_RegType
Dim C_hObjUser
Dim C_ObjUser
Dim C_UseYN
Dim C_ObjPath

<!-- #Include file="../../inc/lgvariables.inc" -->    
'=========================================================================================================
Dim hObjType
Dim hModuleID
Dim hRegType
Dim hObjUser

'=========================================================================================================
Dim IsOpenPop          

'=========================================================================================================
Sub InitSpreadPosVariables()
    C_hObjType       = 1
    C_ObjType        = 2
    C_ObjID          = 3
    C_ObjNm          = 4
    C_hModuleID      = 5
    C_ModuleID       = 6
    C_hRegType       = 7
    C_RegType        = 8
    C_hObjUser       = 9
    C_ObjUser        = 10
    C_UseYN          = 11
    C_ObjPath        = 12
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
        .MaxCols = C_ObjPath+1                                                    
        .MaxRows = 0

         Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetCombo    C_hObjType, "유형코드", 15
        ggoSpread.SSSetCombo    C_ObjType,    "오브젝트 유형", 15
        
        ggoSpread.SSSetEdit     C_ObjID,    "오브젝트 ID",    25, , ,40
        ggoSpread.SSSetEdit     C_ObjNm,    "오브젝트명",    30, , ,50

        ggoSpread.SSSetCombo    C_hModuleID,"모듈 코드", 15
        ggoSpread.SSSetCombo    C_ModuleID, "사용 모듈", 20

        ggoSpread.SSSetCombo    C_hRegType, "등록유형코드", 15
        ggoSpread.SSSetCombo    C_RegType,    "등록 유형", 15

        ggoSpread.SSSetCombo    C_hObjUser, "사용주체 코드", 15
        ggoSpread.SSSetCombo    C_ObjUser,    "오브젝트 사용주체", 16

        ggoSpread.SSSetCheck    C_UseYn,    "사용유무", 10, 2, "", True
        ggoSpread.SSSetEdit     C_ObjPath,    "오브젝트 경로", 30, , ,150
        'ggoSpread.SSSetSplit2(3)
        
        'ToolTip display on
        .TextTip = 1 'Fixed
        frm1.vspdData.SetTextTipAppearance "MS Sans Serif", 12, 0, 0, &HC0FFFF, &H0 'Set font, color
        .ReDraw = true

        Call SetSpreadLock    
        
        Call ggoSpread.MakePairsColumn(C_hObjType,C_ObjType,"1")
        Call ggoSpread.MakePairsColumn(C_ObjID,C_ObjNm,"1")
        Call ggoSpread.MakePairsColumn(C_hModuleID,C_ModuleID,"1")
        Call ggoSpread.MakePairsColumn(C_hRegType,C_RegType,"1")
        Call ggoSpread.MakePairsColumn(C_hObjUser,C_ObjUser,"1")

        Call ggoSpread.SSSetColHidden(C_hObjType,C_hObjType,True)
        Call ggoSpread.SSSetColHidden(C_hModuleID,C_hModuleID,True)
        Call ggoSpread.SSSetColHidden(C_hRegType,C_hRegType,True)
        Call ggoSpread.SSSetColHidden(C_hObjUser,C_hObjUser,True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()        
    With frm1
        .vspdData.ReDraw = False

        ggoSpread.SpreadLock    C_ObjType    , -1, -1
        ggoSpread.SpreadLock    C_ObjID        , -1, -1
        ggoSpread.SSSetRequired    C_ModuleID    , -1, -1
        ggoSpread.SSSetRequired    C_RegType    , -1, -1
        ggoSpread.SSSetRequired    C_ObjUser    , -1, -1
        ggoSpread.SSSetRequired    C_UseYn        , -1, -1
        ggoSpread.SSSetRequired    C_ObjPath    , -1, -1

        .vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired     C_ObjType,    pvStartRow, pvEndRow
        ggoSpread.SSSetRequired     C_ObjID,    pvStartRow, pvEndRow
        ggoSpread.SSSetRequired     C_ModuleID, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired     C_RegType,    pvStartRow, pvEndRow
        ggoSpread.SSSetRequired     C_ObjUser,    pvStartRow, pvEndRow
        ggoSpread.SSSetRequired     C_UseYN,    pvStartRow, pvEndRow
        ggoSpread.SSSetRequired     C_ObjPath,    pvStartRow, pvEndRow
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
    
            C_hObjType   =  iCurColumnPos(1)
            C_ObjType    =  iCurColumnPos(2)
            C_ObjID      =  iCurColumnPos(3)
            C_ObjNm      =  iCurColumnPos(4)
            C_hModuleID  =  iCurColumnPos(5)            
            C_ModuleID   =  iCurColumnPos(6)            
            C_hRegType   =  iCurColumnPos(7)            
            C_RegType    =  iCurColumnPos(8)            
            C_hObjUser   =  iCurColumnPos(9)            
            C_ObjUser    =  iCurColumnPos(10)            
            C_UseYN      =  iCurColumnPos(11)            
            C_ObjPath    =  iCurColumnPos(12)            
                                                                                    
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
Sub InitComboBox()        

    Dim IntRetCD

    On Error Resume next
    
    ggoSpread.Source = frm1.vspdData
    '------ Developer Coding part (Start ) --------------------------------------------------------------     
    IntRetCD = CommonQueryRs("minor_cd, minor_nm","B_Minor","major_cd = " & FilterVar("Z0004", "''", "S") & " order by minor_nm",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboObjType ,lgF0  ,lgF1  ,Chr(11))
    
    IntRetCD = CommonQueryRs("minor_cd, minor_nm","B_Minor","major_cd = " & FilterVar("B0001", "''", "S") & " order by minor_nm",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboModuleID ,lgF0  ,lgF1  ,Chr(11))

    '------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'=========================================================================================================
Sub InitSpreadComboBox()

    Dim IntRetCD
    Dim iPos0, iPos1
    
    On Error Resume next
    
    ggoSpread.Source = frm1.vspdData
    '------ Developer Coding part (Start ) --------------------------------------------------------------     
    IntRetCD = CommonQueryRs("minor_cd, minor_nm","B_Minor","major_cd = " & FilterVar("Z0004", "''", "S") & " order by minor_nm",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = Replace(lgF0, Chr(11), vbTab)
    lgF1 = Replace(lgF1, Chr(11), vbTab)
    ggoSpread.SetCombo lgF0, C_hObjType
    ggoSpread.SetCombo lgF1, C_ObjType                       
    
    IntRetCD = CommonQueryRs("minor_cd, minor_nm","B_Minor","major_cd = " & FilterVar("B0001", "''", "S") & " order by minor_nm",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = Replace(lgF0, Chr(11), vbTab)
    lgF1 = Replace(lgF1, Chr(11), vbTab)
    iPos0 = InStr(lgF0,vbTab) + 1
    iPos1 = InStr(lgF1,vbTab) + 1
    lgF0 = Mid(lgF0, iPos0)
    lgF1 = Mid(lgF1, iPos1)
    ggoSpread.SetCombo lgF0, C_hModuleID
    ggoSpread.SetCombo lgF1, C_ModuleID
                       
    IntRetCD = CommonQueryRs("minor_cd, minor_nm","B_Minor","major_cd = " & FilterVar("Z0011", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = Replace(lgF0, Chr(11), vbTab)
    lgF1 = Replace(lgF1, Chr(11), vbTab)
    ggoSpread.SetCombo lgF0, C_hRegType
    ggoSpread.SetCombo lgF1, C_RegType
    
    IntRetCD = CommonQueryRs("minor_cd, minor_nm","B_Minor","major_cd = " & FilterVar("Z0003", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = Replace(lgF0, Chr(11), vbTab)
    lgF1 = Replace(lgF1, Chr(11), vbTab)
    ggoSpread.SetCombo lgF0, C_hObjUser
    ggoSpread.SetCombo lgF1, C_ObjUser

    '------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'=========================================================================================================
Sub cboObjType_Change()
    frm1.txtObjID.value = ""
    frm1.txtObjNm.value = ""
End Sub

'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                         
    Call ggoOper.LockField(Document, "N")                                   
    
    Call InitSpreadSheet                                                    
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitComboBox
    Call InitSpreadComboBox
    Call InitVariables
    Call SetToolbar("11001101001011")                                        
    
    frm1.txtObjID.focus
    frm1.cboObjType.value  = "A"
    frm1.cboModuleID.value = "*"
    Set gActiveElement = document.activeElement  
    
End Sub

'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    
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
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Select Case Col
         Case  C_ObjType
                iDx = Frm1.vspdData.value
                   Frm1.vspdData.Col = C_hObjType
                Frm1.vspdData.value = iDx
         Case  C_ModuleID
                iDx = Frm1.vspdData.value
                   Frm1.vspdData.Col = C_hModuleID
                Frm1.vspdData.value = iDx
         Case  C_RegType
                iDx = Frm1.vspdData.value
                   Frm1.vspdData.Col = C_hRegType
                Frm1.vspdData.value = iDx
         Case  C_ObjUser
                iDx = Frm1.vspdData.value
                   Frm1.vspdData.Col = C_hObjUser
                Frm1.vspdData.value = iDx
         Case Else
    End Select    
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'=========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'=========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

    End With

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
        If lgStrPrevKey <> "" Then                                                        
            Call DisableToolBar(Parent.TBC_QUERY)                                                

            If DbQuery = False Then                                                       
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
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")                    '⊙: "Will you destory previous data"
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If


    If frm1.txtObjID.value = "" Then
        frm1.txtObjNm.value = ""
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
    On Error Resume Next                                                        
End Function

'=========================================================================================================
Function FncDelete() 
    On Error Resume Next                                                        
End Function

'=========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                                
    
    Err.Clear                                                                    
    

    ggoSpread.Source = frm1.vspdData                              
    If ggoSpread.SSCheckChange = False Then                   
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '⊙: Display Message(There is no changed data.)
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
        If .maxrows < 1 Then Exit Function
        .focus
        Set gActiveElement = document.activeElement 
        ggoSpread.Source = frm1.vspdData

        .EditMode = True    
        .ReDraw = False
        ggoSpread.CopyRow

        frm1.vspdData.Row = frm1.vspdData.ActiveRow
        .Text = ""
        
        .ReDraw = True
        nActiveRow = frm1.vspdData.ActiveRow
        SetSpreadColor nActiveRow, nActiveRow

        'Key field clear
        .SetText C_hObjType, nActiveRow, ""
        .SetText C_ObjType, nActiveRow, ""
        .SetText C_ObjID, nActiveRow, ""
        
    End with
    
End Function

'=========================================================================================================
Function FncCancel() 
    If frm1.vspdData.maxrows < 1 Then Exit Function
    
    ggoSpread.Source = frm1.vspdData    
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
    Dim iDelRowCnt, i


    With frm1.vspdData     
        If .maxrows < 1 Then Exit Function
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
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")        
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
    DbQuery = False

    Call LayerShowHide(1)

    Err.Clear                                                               

    Dim strVal
   
    With frm1
        If lgIntFlgMode = Parent.OPMD_UMODE Then                                                
            strVal = BIZ_PGM_ID & "?txtMode="        & Parent.UID_M0001                            
            strVal = strVal & "&cboObjType="        & Trim(.htxtObjType.value)            
            strVal = strVal & "&txtModuleID="        & Trim(.htxtModuleID.value)            
            strVal = strVal & "&txtObjID="            & Trim(.htxtObjID.value)            
            strVal = strVal & "&txtMaxRows="        & .vspdData.MaxRows
            strVal = strVal & "&lgStrPrevKey="        & lgStrPrevKey
        Else            
            strVal = BIZ_PGM_ID & "?txtMode="        & Parent.UID_M0001                            
            strVal = strVal & "&cboObjType="        & Trim(.cboObjType.value)            
            strVal = strVal & "&txtModuleID="        & Trim(.cboModuleID.value)            
            strVal = strVal & "&txtObjID="            & Trim(.txtObjID.value)                
            strVal = strVal & "&txtMaxRows="        & .vspdData.MaxRows
            strVal = strVal & "&lgStrPrevKey="        & lgStrPrevKey
        End If

        Call RunMyBizASP(MyBizASP, strVal)                                                
    End With
    
    DbQuery = True
End Function
'=========================================================================================================
Function DbQueryOk()                                                        

    lgIntFlgMode = Parent.OPMD_UMODE                                                
    
    Call ggoOper.LockField(Document, "Q")                                    
    Call SetToolbar("11001111001111")
    frm1.vspdData.Focus
End Function
'=========================================================================================================
Function DbSave()        
    Dim lRow        
    Dim lGrpCnt     
    Dim strVal, strDel
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
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_hObjType, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_ObjID, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_ObjNm, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_hModuleID, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_hRegType, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_ObjPath, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_hObjUser, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_UseYN, lRow, "X", "X"))      & iRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      
                                                    strVal = strVal & "U"                       & iColSep
                                                    strVal = strVal & lRow                      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_hObjType, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_ObjID, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_ObjNm, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_hModuleID, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_hRegType, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_ObjPath, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_hObjUser, lRow, "X", "X"))      & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_UseYN, lRow, "X", "X"))      & iRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.DeleteFlag                                      
                                                    strDel = strDel & "D"                       & iColSep
                                                    strDel = strDel & lRow                      & iColSep
                    								strDel = strDel & Trim(GetSpreadText(.vspdData, C_hObjType, lRow, "X", "X"))      & iColSep
                    								strDel = strDel & Trim(GetSpreadText(.vspdData, C_ObjID, lRow, "X", "X"))      & iRowSep
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
    ggoSpread.Source = frm1.vspdData
    frm1.vspdData.MaxRows = 0
    Call MainQuery()

End Function

'=========================================================================================================
'    Name : OpenObjID()
'    Description : Calendar Type Popup
'=========================================================================================================
Function OpenObjID()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)
    Dim IntRetCD
    Dim iModuleId                                    

    If Trim(frm1.cboObjType.Value) = "" Then
        IntRetCD = DisplayMsgBox("210024", VBOKONLY,"X","X")
        Exit Function
    End If

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True
    
    iModuleId = Replace(Trim(frm1.cboModuleID.Value),"*","%")                                                    
    If Len(iModuleId) = 0 Then
        iModuleId = "%"
    End If
    
    arrParam(0) = "오브젝트 ID 팝업"
    arrParam(1) = "Z_OBJECT_INFO z, b_minor b, b_minor c"
    arrParam(2) = Trim(frm1.txtObjID.Value)        ' Code Condition
    arrParam(3) = ""                            ' Name Condition
    arrParam(4) = "z.obj_type =  " & FilterVar(frm1.cboObjType.Value, "''", "S") & " " & _
                  "and z.module_id LIKE  " & FilterVar(iModuleId, "''", "S") & "  " & _
                  "and b.major_cd = " & FilterVar("z0004", "''", "S") & " " & _
                  "and b.minor_cd = z.obj_type " & _
                  "and c.major_cd = " & FilterVar("b0001", "''", "S") & " " & _
                  "and c.minor_cd = z.module_id " & _
                  "and z.lang_cd =  " & FilterVar(Parent.gLang , "''", "S") & "" 
    arrParam(5) = "오브젝트 ID"
    
    arrField(0) = "ED12" & Parent.gColSep & "z.OBJ_ID"                        ' Field(0)
    arrField(1) = "ED25" & Parent.gColSep & "z.OBJ_NM"                        ' Field(1)
    arrField(2) = "ED15" & Parent.gColSep & "b.minor_nm"                    ' Field(2)
    arrField(3) = "ED15" & Parent.gColSep & "c.minor_nm"                    ' Field(3)
    
    arrHeader(0) = "오브젝트 ID"                ' Header(0)
    arrHeader(1) = "오브젝트명"                    ' Header(1)
    arrHeader(2) = "오브젝트 유형"                ' Header(2)
    arrHeader(3) = "사용모듈"                    ' Header(3)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=625px; dialogHeight=455px; center: Yes; help: No; resizable: No; status: No;")                ' End
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetObjID(arrRet)
    End If    
	frm1.txtObjID.focus
	Set gActiveElement = document.activeElement
    
End Function

'=========================================================================================================
'    Name : SetObjID()
'    Description : 오브젝트 Id Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetObjID(byval arrRet)
    frm1.txtObjID.value = arrRet(0)
    frm1.txtObjNm.value = arrRet(1)
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->    
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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>오브젝트 관리</font></td>
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
            <TABLE <%=LR_SPACE_TYPE_20%>>
                <TR>
                    <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
                </TR>
                <TR>
                    <TD HEIGHT=20 WIDTH=100%>
                        <FIELDSET CLASS="CLSFLD">
                            <TABLE <%=LR_SPACE_TYPE_40%>>
                                <TR>
                                    <TD CLASS=TD5 NOWRAP>오브젝트 유형</TD>
                                    <TD CLASS=TD6 NOWRAP><SELECT NAME="cboObjType" ALT="오브젝트 유형" STYLE="WIDTH: 150px" tag="12" onChange="cboObjType_Change()"><OPTION VALUE="" selected></OPTION></SELECT></TD>    
                                    
                                    <TD CLASS=TD5 NOWRAP>사용 모듈</TD>
                                    <TD CLASS=TD6 NOWRAP><SELECT NAME="cboModuleID" ALT="사용 모듈" STYLE="WIDTH: 150px" tag="12"><OPTION VALUE="" selected></OPTION></SELECT></TD>                        
                                </TR>
                                <TR>
                                    <TD CLASS=TD5 NOWRAP>오브젝트 ID</TD>
                                    <TD CLASS=TD656 NOWRAP colspan=3><INPUT TYPE=TEXT NAME="txtObjID" SIZE=20 MAXLENGTH=30 tag="11" ALT="오브젝트 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenObjID()">&nbsp;<INPUT TYPE=TEXT NAME="txtObjNm" SIZE=40 tag="14" ALT="오브젝트명"></TD>

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
                            <TD HEIGHT="100%" colspan=4>
                                <script language =javascript src='./js/za013ma1_I612612360_vspdData.js'></script>
                                </TD>
                            </TR>
                        </TABLE>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="htxtObjType" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtObjID" tag="24"><INPUT TYPE=HIDDEN NAME="htxtModuleID" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
