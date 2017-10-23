
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za008ma1
*  4. Program Name         : Audit Management
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
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">        

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             

'=========================================================================================================
 Const BIZ_PGM_ID = "Za008MB1.asp"                                                
 Const JUMP_PGM_ID1 = "Za009ma1"
 Const JUMP_PGM_ID2 = "Za010ma1"
'=========================================================================================================
 Dim C_TableID
 Dim C_TableNm
 Dim C_TableTypeCd
 Dim C_TableType
 Dim C_Insert
 Dim C_Update
 Dim C_Delete
 

<!-- #Include file="../../inc/lgvariables.inc" -->    
'=========================================================================================================
Dim IsOpenPop          

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

    lgIntFlgMode = Parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
    '---- Coding part--------------------------------------------------------------------
    
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
        Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)

        .ReDraw = false                   
        .MaxCols = C_Delete + 1                                                    
        .MaxRows = 0

        Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetEdit        C_TableID,    "테이블 ID",    30,,,30
        ggoSpread.SSSetEdit        C_TableNm,    "테이블명",        40,,,40
        ggoSpread.SSSetEdit        C_TableTypeCd, "",                16
        ggoSpread.SSSetEdit        C_TableType,"테이블속성",    16,,,16
        ggoSpread.SSSetCheck    C_Insert,    "입력",            10,,,1
        ggoSpread.SSSetCheck    C_Update,    "수정",            10,,,1
        ggoSpread.SSSetCheck    C_Delete,    "삭제",            10,,,1
        .ReDraw = true

        Call SetSpreadLock
        
        Call ggoSpread.MakePairsColumn(C_TableId,C_TableNm,"1")
        Call ggoSpread.MakePairsColumn(C_TableTypeCd,C_TableType,"1")

        Call ggoSpread.SSSetColHidden(C_TableTypeCd,C_TableTypeCd,True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()

    With frm1
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock    C_TableID,        -1,    C_TableID
    ggoSpread.spreadLock    C_TableNm,        -1, C_TableNm
    ggoSpread.spreadLock    C_TableType,    -1, C_TableType
    .vspdData.ReDraw = True
    End With

End Sub

'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
End Sub

'=========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_TableID      =  iCurColumnPos(1)
            C_TableNm   =  iCurColumnPos(2)
            C_TableTypeCd      =  iCurColumnPos(3)
            C_TableType          =  iCurColumnPos(4)
            C_Insert       =  iCurColumnPos(5)
            C_Update         =  iCurColumnPos(6)
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
    Call ggoOper.LockField(Document, "N")                                   
    
    Call InitSpreadSheet                                                    
    Call InitVariables                                                      
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("11001001000011")
    frm1.txtTableID.focus 
    Set gActiveElement = document.activeElement
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
    Call SetPopupMenuItemInf("0101111111")    

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
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
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
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")                '데이타가 변경되었습니다. 조회하시겠습니까?
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
    frm1.vspdData.focus
    frm1.vspdData.ReDraw = False
    
    ggoSpread.Source = frm1.vspdData    
    frm1.vspdData.ReDraw = False
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    frm1.vspdData.ReDraw = True
End Function

'=========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData    
    ggoSpread.EditUndo                                                  
End Function

'=========================================================================================================
Function FncInsertRow() 
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
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")                '데이타가 변경되었습니다. 종료 하시겠습니까?
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
    Err.Clear                                                               

    DbQuery = False
    
    Call LayerShowHide(1)
    
    With frm1
        If lgIntFlgMode = Parent.OPMD_UMODE Then
            strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                            
            strVal = strVal & "&txtTableID="   & .hTableID.value                         
            strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
            strVal = strVal & "&txtMaxRows="   & .vspdData.MaxRows
        Else
            strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                            
            strVal = strVal & "&txtTableID="   & Trim(.txtTableID.value)                
            strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
            strVal = strVal & "&txtMaxRows="   & .vspdData.MaxRows
        End If    
   
        Call RunMyBizASP(MyBizASP, strVal)                                            
    End With
    
    DbQuery = True
    
End Function
'=========================================================================================================
Function DbQueryOk()                                                        
    
    lgIntFlgMode = Parent.OPMD_UMODE                                                
    
    Call ggoOper.LockField(Document, "Q")                                    
    Call SetToolbar("11001011000111")            
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
                    								 strVal = strVal & Trim(GetSpreadText(.vspdData, C_TableID, lRow, "X", "X"))      & iColSep
                    If Trim(GetSpreadText(.vspdData, C_Insert, lRow, "X", "X")) <> "" Then
                                                     strVal = strVal & Trim(GetSpreadText(.vspdData, C_Insert, lRow, "X", "X"))      & iColSep
                    Else
                                                     strVal = strVal & "2"                         & iColSep
                    End If
                    If Trim(GetSpreadText(.vspdData, C_Update, lRow, "X", "X")) <> "" Then
                                                     strVal = strVal & Trim(GetSpreadText(.vspdData, C_Update, lRow, "X", "X"))      & iColSep
                    Else
                                                     strVal = strVal & "2"                         & iColSep
                    End If
                    If Trim(GetSpreadText(.vspdData, C_Delete, lRow, "X", "X")) <> "" Then
                                                     strVal = strVal & Trim(GetSpreadText(.vspdData, C_Delete, lRow, "X", "X"))      & iRowSep
                    Else
                                                     strVal = strVal & "2"                         & iRowSep
                    End If
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      
                                                     strVal = strVal & "U"                       & iColSep
                                                     strVal = strVal & lRow                      & iColSep        '
                    								 strVal = strVal & Trim(GetSpreadText(.vspdData, C_TableID, lRow, "X", "X"))      & iColSep
                    If Trim(GetSpreadText(.vspdData, C_Insert, lRow, "X", "X")) <> "" Then
                                                     strVal = strVal & Trim(GetSpreadText(.vspdData, C_Insert, lRow, "X", "X"))      & iColSep
                    Else
                                                     strVal = strVal & "2"                         & iColSep
                    End If
                    If Trim(GetSpreadText(.vspdData, C_Update, lRow, "X", "X")) <> "" Then
                                                     strVal = strVal & Trim(GetSpreadText(.vspdData, C_Update, lRow, "X", "X"))      & iColSep
                    Else
                                                     strVal = strVal & "2"                         & iColSep
                    End If
                    If Trim(GetSpreadText(.vspdData, C_Delete, lRow, "X", "X")) <> "" Then
                                                     strVal = strVal & Trim(GetSpreadText(.vspdData, C_Delete, lRow, "X", "X"))      & iRowSep
                    Else
                                                     strVal = strVal & "2"                         & iRowSep
                    End If
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.DeleteFlag                                      
                                                     strDel = strDel & "D"                       & iColSep        
                                                     strDel = strDel & lRow                      & iColSep
                    								 strDel = strDel & Trim(GetSpreadText(.vspdData, C_TableID, lRow, "X", "X"))      & iRowSep
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
'    Name : OpenTableID()
'    Description : TableID PopUp
'=========================================================================================================
Function OpenTableID(Byval strCode)

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "테이블 ID 팝업"
    arrParam(1) = "z_table_info t, z_audit_policy a"
    arrParam(2) = strCode                            ' Code Condition
    arrParam(3) = ""                                ' Name Cindition
    arrParam(4) = "a.lang_cd =  " & FilterVar(Parent.gLang , "''", "S") & " " & _
                  "and t.lang_cd = a.lang_cd " & _
                  "and t.table_id = a.table_id " & _
                  "and t.use_yn = " & FilterVar("1", "''", "S") & " "
    arrParam(5) = "테이블 ID"            
    
    arrField(0) = "ED24" & Parent.gColSep & "a.table_id"
    arrField(1) = "t.table_nm"                    
    
    arrHeader(0) = "테이블 ID"                
    arrHeader(1) = "테이블명"                
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=520px; dialogHeight=455px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetTableID(arrRet)
    End If    
    frm1.txtTableID.focus
    Set gActiveElement = document.activeElement

End Function
'=========================================================================================================
Function OpenRefAudit()
    Dim arrRet, IntRetCD
    Dim arrParam, arrField, arrHeader
    Dim iCalledAspName
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZA008RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZA008RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
        
    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=800px; dialogHeight=466px; center: Yes; help: No; resizable: No; status: No;")
                        '600
    IsOpenPop = False
    
    If arrRet(0, 0) = "" Then
        Exit Function
    Else
        Call SetRefAudit(arrRet)
    End If
    
End Function

'=========================================================================================================
'    Name : SetTableID()
'    Description : Plant Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetTableID(byval arrRet)
    frm1.txtTableID.Value = arrRet(0)
    frm1.txtTableNm.Value = arrRet(1)
End Function

'=========================================================================================================
'    Name : SetRefTableID()
'    Description : Table Reference에서 선택된 Table ID Return
'=========================================================================================================
Function SetRefAudit(Byval arrRet)
    
    Dim intRtnCnt, strData
    Dim TempRow, I
    
    With frm1
    
        .vspdData.focus
        lgBlnFlgChgValue = True
        ggoSpread.Source = .vspdData
        .vspdData.ReDraw = False    
    
        TempRow = .vspdData.MaxRows                                                
        .vspdData.MaxRows = .vspdData.MaxRows + (Ubound(arrRet, 1) + 1)            
 
        For I = TempRow to .vspdData.MaxRows - 1
        	.vspdData.SetText 0, I+1, ggoSpread.InsertFlag
        	.vspdData.SetText C_TableID, I+1, arrRet(I - TempRow, 0)
        	.vspdData.SetText C_TableNm, I+1, arrRet(I - TempRow, 1)
        	.vspdData.SetText C_TableTypeCd, I+1, arrRet(I - TempRow, 2)
        	.vspdData.SetText C_TableType, I+1, arrRet(I - TempRow, 3)
        	.vspdData.SetText C_Insert, I+1, arrRet(I - TempRow, 4)
        	If GetSpreadText(.vspdData, C_Insert, I+1, "X", "X") = "" Then
        		.vspdData.Col = C_Insert
        		.vspdData.Row = I+1
                .vspdData.CellType = 1
                .vspdData.Protect = True
                .vspdData.Lock = True
            End If
			.vspdData.SetText C_Update, I+1, arrRet(I - TempRow, 5)
        	If GetSpreadText(.vspdData, C_Update, I+1, "X", "X") = "" Then
        		.vspdData.Col = C_Update
        		.vspdData.Row = I+1
                .vspdData.CellType = 1
                .vspdData.Protect = True
                .vspdData.Lock = True
            End If

			.vspdData.SetText C_Delete, I+1, arrRet(I - TempRow, 5)
        	If GetSpreadText(.vspdData, C_Delete, I+1, "X", "X") = "" Then
        		.vspdData.Col = C_Delete
        		.vspdData.Row = I+1
                .vspdData.CellType = 1
                .vspdData.Protect = True
                .vspdData.Lock = True
            End If
        Next    
        
        .vspdData.ReDraw = True
    End With
End Function
'=========================================================================================================
Function JumpProgram()
    On Error Resume Next    
    PgmJump(JUMP_PGM_ID1)    
End Function
'=========================================================================================================
Function JumpProgram1()
    Dim strTable
    strTable = ""

    On Error Resume Next
    With frm1.vspddata
        If .MaxRows <> 0 then
            strTable = GetSpreadText(frm1.vspddata, C_TableID, .activeRow, "X", "X")
        End If        
    End With
    
    WriteCookie "Za009ma1_TableID", Trim(strTable)

    PgmJump(JUMP_PGM_ID2)
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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>감사 정책 관리</font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* align=right>
                                            <a href="vbscript:OpenRefAudit">감사 정책 추가</A>  
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
                                    <TD CLASS="TD5">테이블 ID</TD>
                                    <TD CLASS="TD656" COLSPAN=3><INPUT TYPE=TEXT NAME="txtTableID" SIZE=30 MAXLENGTH=30 tag="11" ALT="테이블 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTableID" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTableID frm1.txtTableID.value">&nbsp;<INPUT TYPE=TEXT NAME="txtTableNm" SIZE=40 tag="14N"></TD>
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
                                    <script language =javascript src='./js/za008ma1_vspdData_vspdData.js'></script>
                                </TD>
                            </TR>
                        </TABLE>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR HEIGHT=20>
        <TD>
            <TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
                <TR>
                    <TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpProgram">감사 정보 개요 조회</A>&nbsp;|&nbsp;
                                            <A href="vbscript:JumpProgram1">감사 정보 상세 조회</A>
                    </TD>
                    <TD WIDTH=10>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="Za008mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hTableID" tag="24"
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
