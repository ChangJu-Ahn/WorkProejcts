<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Role별 Menu Assign
*  2. Function Name        : zb003ma1.asp
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         : 
*  6. Comproxy List        : 
*                              zb0051
*                              zb0058
*                              
*  7. Modified date(First) :
*  8. Modified date(Last)  : 2001/11/15 
*  9. Modifier (First)     : Kang Doo Sig
* 10. Modifier (Last)      : Lee Seok Gon
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
Const BIZ_PGM_ID = "ZB003MB1_KO441.asp"                                                
Const JUMP_PGM_ID = "ZB005MA1"
'=========================================================================================================
Dim C_MenuID
Dim C_MenuNm
Dim C_MenuTypeCD
Dim C_MenuType
Dim C_ActionCd
Dim C_Action

'=========================================================================================================

<!-- #Include file="../../inc/lgvariables.inc" -->    

dim lgStrPrevActId
dim    lgStrQueryFlag
dim IsOpenPop

'=========================================================================================================
Sub InitSpreadPosVariables()
    C_MenuID        =    1
    C_MenuNm        =    2
    C_MenuTypeCD    =    3
    C_MenuType        =    4
    C_ActionCd        =    5
    C_Action        =    6
End Sub

'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
    
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           
    lgStrPrevActId = ""
    lgLngCurRows = 0                            
    lgSortKey = 1
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
        .MaxCols = C_Action + 1                                                    
        .MaxRows = 0

        Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetEdit C_MenuID,    "Menu ID",        26
        ggoSpread.SSSetEdit C_MenuNm,    "Menu Name",    46
        ggoSpread.SSSetEdit C_MenuTypeCd, "",                22
        ggoSpread.SSSetEdit C_MenuType, "Menu Type",    22
        ggoSpread.SSSetEdit C_ActionCd,    "",                    22
        ggoSpread.SSSetEdit C_Action,    "Action",        22
        .ReDraw = true

        Call SetSpreadLock
        
        Call ggoSpread.MakePairsColumn(C_MenuID,C_MenuNm,"1")
        Call ggoSpread.MakePairsColumn(C_MenuTypeCd,C_MenuType,"1")
        Call ggoSpread.MakePairsColumn(C_ActionCd,C_Action,"1")

        Call ggoSpread.SSSetColHidden(C_MenuTypeCd,C_MenuTypeCd,True)
        Call ggoSpread.SSSetColHidden(C_ActionCd,C_ActionCd,True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()

    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock C_MenuID,        -1, -1
    ggoSpread.spreadLock C_MenuNm,        -1, -1
    ggoSpread.spreadLock C_MenuType,    -1, -1
    ggoSpread.spreadLock C_Action,        -1, -1
    
    .vspdData.ReDraw = True

    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetProtected    C_MenuID,    pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_MenuNm,    pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_MenuType, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_Action,    pvStartRow, pvEndRow
    
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
        
            C_MenuID        =  iCurColumnPos(1)
            C_MenuNm        =  iCurColumnPos(2)
            C_MenuTypeCD    =  iCurColumnPos(3)
            C_MenuType      =  iCurColumnPos(4)
            C_ActionCd      =  iCurColumnPos(5)
            C_Action        =  iCurColumnPos(6)

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
    
End Sub

'=========================================================================================================
sub Form_Load()
    Call LoadInfTB19029                                                         
    Call ggoOper.LockField(Document, "N")                                   
    
    Call InitSpreadSheet                                                    
    Call InitVariables                                                      
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("1100100100101111")                                            
    Call CookiePage(0)
    frm1.txtRoleID.focus
    Set gActiveElement = document.activeElement
End Sub
'=========================================================================================================
sub Form_QueryUnload(Cancel , UnloadMode )
   
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
sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True

End Sub

'=========================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim index
    
    With frm1.vspdData
        If Col = C_TypeNm And Row > 0 Then
            .Row = Row
            .Col = Col
     		.TypeComboBoxCurSel
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
sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop)
    
    
    If CheckRunningBizProcess = True Then                            
        Exit Sub
    End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    lgStrQueryFlag = "Q"

     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then    
        If lgStrPrevKey <> "" And lgStrPrevActId <> "" Then                            
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
    Call InitComboBox
    Call ggoSpread.ReOrderingSpreadData()
    Call InitData()
End Sub

'=========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               
    
    lgStrQueryFlag = "Q"

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")                '데이타가 변경되었습니다. 조회하시겠습니까?
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
    
    If DBQuery = False Then 
       Exit Function 
    End If 
     
    FncQuery = True                                                            
    
End Function

'=========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
    Err.Clear                                                               
    On Error Resume next                                                    
    
    

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
       
    End If
    

    Call ggoOper.ClearField(Document, "A")
    Call ggoSpread.ClearSpreadData    
    Call ggoOper.LockField(Document, "N")                                          
    Call InitVariables                                                              
    
    FncNew = True                                                                  

End Function

'=========================================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       
    
    Err.Clear                                                               
    On Error Resume next                                                    
    
    

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
        Call DisplayMsgbox("900002","x","x","x")                                  
        Exit Function
    End If
    

    If DbDelete = False Then                                                
       Exit Function                                                        
    End If
    

    Call ggoOper.ClearField(Document, "A")
    Call ggoSpread.ClearSpreadData    
        
    FncDelete = True                                                        
    
End Function

'=========================================================================================================
Function FncSave() 
    
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               
    On Error Resume next                                                    
    

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001","x","x","x")                          
        Exit Function
    End If
    

    ggoSpread.Source = frm1.vspdData
    If Not chkField(Document, "2") OR ggoSpread.SSDefaultCheck = False Then                                  
       Exit Function
    End If
    

    Call DbSave                                                                  
    
    FncSave = True                                                          
    
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
    
    .focus
    ggoSpread.Source = frm1.vspdData 
    
    .Col = 8'ggoSpread.SSGetColsIndex(8)
    If .Text = "A" Then
        Call DisplayMsgBox("183104","x","x","x")              '☆: you must release this line if you change msg into code
        Exit Function
    End If
    
    lDelRows = ggoSpread.DeleteRow

    lgBlnFlgChgValue = True
    
    End With
End Function

'=========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData    
    ggoSpread.EditUndo                                                  
End Function

'=========================================================================================================
Function FncPrev() 
    Dim IntRetCD 
    
    Err.Clear
    
    FncPrev = False                                                            
    
    lgStrQueryFlag = "P"
                                                              

    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013",Parent.VB_YES_NO,"x","x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    
    Call InitVariables                                                        

    If Not chkField(Document, "1") Then                                
        Exit Function
    End If
    

    If DBQuery = False Then                                                        
        Exit Function 
    End If                                                             
           
    FncPrev = True                                                                
    
End Function

'=========================================================================================================
Function FncNext() 
    Dim IntRetCD 
    
    Err.Clear
    
    FncNext = False                                                          

    lgStrQueryFlag = "N"
    

    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    
    Call InitVariables

    If Not chkField(Document, "1") Then                                    
       Exit Function
    End If
    

    If DBQuery = False Then                                                        
        Exit Function 
    End If                                                             
           
    FncNext = True                                                                
    
End Function

'=========================================================================================================
Function FncCopy() 
    With frm1
        If .vspdData.ActiveRow > 0 Then
            frm1.vspdData.focus
            frm1.vspdData.ReDraw = False
                
            ggoSpread.Source = frm1.vspdData    
            ggoSpread.CopyRow
            SetSpreadColor frm1.vspdData.ActiveRow
                
            frm1.vspdData.ReDraw = True
        End If
    End With    
End Function

'=========================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)                
End Function

'=========================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   
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
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")                '데이타가 변경되었습니다. 종료 하시겠습니까?
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    FncExit = True
End Function

'=========================================================================================================
Function Clear()
    Call ggoOper.ClearField(Document, "2")
    Call ggoSpread.ClearSpreadData        
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

    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    Dim mnuType
    Dim IntRetCD

    DbQuery = False
    
    Call LayerShowHide(1)  
    
    Err.Clear                                                               

    Dim strVal    
    Dim ArrVar
    
    With frm1
    

    IntRetCD = CommonQueryRs("usr_role_nm","Z_Usr_Role","usr_role_id =  " & FilterVar(.txtRoleID.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = Replace(lgF0, Chr(11), vbTab)
    lgF0 = Replace(lgF0, " ","")
    
        
    .txtRoleNm.value = lgF0
        
 
    If lgIntFlgMode = Parent.OPMD_UMODE Then        
        if .rdoConf(0).checked then mnuType = "M"
        if .rdoConf(1).checked then mnuType = "P"
        if .rdoConf(2).checked then    mnuType = ""
            
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                            
        strVal = strVal & "&txtRoleID=" & Trim(.hRoleID.value)                    
        strVal = strVal & "&txtRoleNm=" & Trim(.txtRoleNm.value)
        strVal = strVal & "&txtMenuType=" & mnuType                                      
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&lgStrPrevActId=" & lgStrPrevActId
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        strVal = strVal & "&txtPrvNext=" &     lgStrQueryFlag 
    Else
        if .rdoConf(0).checked then mnuType = "M"
        if .rdoConf(1).checked then mnuType = "P"
        if .rdoConf(2).checked then    mnuType = ""
    
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                            
        strVal = strVal & "&txtRoleID=" & Trim(.txtRoleID.value)                
        strVal = strVal & "&txtRoleNm=" & Trim(.txtRoleNm.value)
        strVal = strVal & "&txtMenuType=" & mnuType                                     
        strVal = strVal & "&lgStrPrevKey=" & Trim(.txtMenuID.value)
        strVal = strVal & "&lgStrPrevActId=" & lgStrPrevActId
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        strVal = strVal & "&txtPrvNext=" &     lgStrQueryFlag 
    End If
        
        
        Call RunMyBizASP(MyBizASP, strVal)                                        
        
    End With
    
    DbQuery = True

End Function
'=========================================================================================================
Function DbQueryOk()                                                        
    

    lgIntFlgMode = Parent.OPMD_UMODE                                                
    
    'Call ggoOper.LockField(Document, "Q")                                    
    
    Call SetToolbar("1100101111011111")                                            
    Frm1.vspdData.Focus
End Function

'=========================================================================================================
function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
    Dim strVal
    Dim ActFlag
    Dim TmpType
    Dim IntRetCD
    Dim iColSep, iRowSep

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep
    
    DbSave = False                                                          
    
    Call LayerShowHide(1)  
    
    On Error Resume next                                                   
    TmpType=0
    
    With frm1
        .txtMode.value = Parent.UID_M0002
        .txtUpdtUserId.value = Parent.gUsrID
        .txtInsrtUserId.value = Parent.gUsrID
    

        lGrpCnt = 1
    
        strVal = ""
        

        For lRow = 1 To .vspdData.MaxRows
            If GetSpreadText(.vspdData, 0, lRow, "X", "X") = ggoSpread.InsertFlag Or GetSpreadText(.vspdData, 0, lRow, "X", "X") = ggoSpread.UpdateFlag Or GetSpreadText(.vspdData, 0, lRow, "X", "X") = ggoSpread.DeleteFlag Then
                Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")
                    Case ggoSpread.InsertFlag                                            
                        ActFlag = "C"
                    Case ggoSpread.UpdateFlag                                            
                        ActFlag = "U"
                    Case ggoSpread.DeleteFlag                                            
                        ActFlag = "D"            
                End Select                    

                strVal = strVal & ActFlag & iColSep & lRow & iColSep         
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_MenuID, lRow, "X", "X")) & iColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_MenuTypeCD, lRow, "X", "X")) & iColSep
                If Trim(GetSpreadText(.vspdData, C_MenuTypeCD, lRow, "X", "X")) ="M" and ActFlag= "D" Then 
                    TmpType = 1
                End If 
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_ActionCD, lRow, "X", "X")) & iColSep & Trim(.txtRoleID.Value) & iRowSep
                lGrpCnt = lGrpCnt + 1
            End if                
         Next

        If TmpType = 1 then 
            IntRetCD = DisplayMsgBox("211438", Parent.VB_YES_NO,"x","x")                '메뉴를 삭제하시면 하위 프로그램도 삭제됩니다. 계속하시겠습니까?
            If IntRetCD = vbNo Then
                Call LayerShowHide(0) 
                Call SetToolbar("1100101111011111")                                             
                Exit Function
            End If
        end if 
                
        .txtMaxRows.value = lGrpCnt-1
        .txtSpread.value = strVal
        
        
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
function DbDelete() 
End Function

'=========================================================================================================
'    Name : OpenRoleID()
'    Description : Role ID PopUp
'=========================================================================================================
Function OpenRoleID(Byval strCode)

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "Role ID Popup"                    ' 팝업 명칭 
    arrParam(1) = "Z_USR_ROLE"                        ' TABLE 명칭 
    arrParam(2) = strCode                            ' Code Condition
    arrParam(3) = ""                                ' Name Cindition
    arrParam(4) = "Compst_Role_Type = 0"                                ' Where Condition
    arrParam(5) = "Role ID"            
    
    arrField(0) = "USR_ROLE_ID"                        ' Field명(0)
    arrField(1) = "USR_ROLE_NM"                        ' Field명(1)
    
    arrHeader(0) = "Role ID"                    ' Header명(0)
    arrHeader(1) = "Role명"                        ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
             "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
    If arrRet(0) = "" Then
    Else
        Call SetRoleID(arrRet)
    End If    
    
    
    frm1.txtRoleID.focus 
    Set gActiveElement = document.activeElement
End Function
'=========================================================================================================
Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(3)
    Dim mnuType
    Dim ii
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZB003RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZB003RA1", "x")
        IsOpenPop = False
        Exit Function
    End If

    If frm1.txtRoleID.Value = "" Then
        call DisplayMsgBox ("900002","X","X","X")
        IsOpenPop = False
        Exit Function
    End If
    
    if frm1.rdoConf(0).checked then mnuType = "M"
    if frm1.rdoConf(1).checked then mnuType = "P"
    if frm1.rdoConf(2).checked then    mnuType = ""                
        
    arrParam(0) = frm1.txtCd.Value                ' RoleID를 넘긴다 
    arrParam(1) = frm1.txtMenuID.Value
    arrParam(2) = mnuType    
    arrParam(3) = ""    
            
    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False
    
    If arrRet(0, 0) = "" Then
        Exit Function
    Else
        Call SetRefMenu(arrRet)
    End If
    
End Function

'=========================================================================================================
Function CookiePage(byval nflag)

    Dim strTemp
    
    If nflag=1 Then
    
    Else
        strTemp = ReadCookie("ZB_05_Cmp_Role_ID")
        WriteCookie "ZB_05_Cmp_Role_ID" ,""
        If Trim(strTemp) <>"" Then 
            frm1.txtRoleID.value  = strTemp        
            Call DbQuery
        End If 
    End If
    
    
End Function

'=========================================================================================================
'    Name : SetRoleID()
'    Description : Role ID Popup에서 선택된 Role ID Insert
'=========================================================================================================
Function SetRoleID(byval arrRet)

    frm1.txtRoleID.Value = arrRet(0)
    frm1.txtRoleNm.Value  = arrRet(1)

End Function

'=========================================================================================================
'    Name : SetRefComposite()
'    Description : Plant Popup에서 Return되는 값 setting
'=========================================================================================================

Function SetRefMenu(Byval arrRet)
    
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
            .vspdData.SetText C_MenuID, I+1, arrRet(I - TempRow, 0)
            .vspdData.SetText C_MenuNm, I+1, arrRet(I - TempRow, 1)
            .vspdData.SetText C_MenuTypeCd, I+1, arrRet(I - TempRow, 2)
            .vspdData.SetText C_MenuType, I+1, arrRet(I - TempRow, 3)
            .vspdData.SetText C_ActionCd, I+1, arrRet(I - TempRow, 4)
            .vspdData.SetText C_Action, I+1, arrRet(I - TempRow, 5)
        Next    
        
        ggoSpread.SpreadUnlock C_MenuID, TempRow + 1, 2, .vspdData.MaxRows                
        ggoSpread.ssSetProtected C_MenuID, TempRow + 1, .vspdData.MaxRows
        ggoSpread.ssSetProtected C_MenuNm, TempRow + 1, .vspdData.MaxRows
        ggoSpread.ssSetProtected C_MenuType, TempRow + 1, .vspdData.MaxRows
        ggoSpread.ssSetProtected C_Action, TempRow + 1, .vspdData.MaxRows                
        
        .vspdData.ReDraw = True
    End With
    
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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>메뉴 Assignment to Role</font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* align=right>
                                            <a href="vbscript:OpenRefMenu">Menu ID 참조</A>  
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
                            <TABLE <%=LR_SPACE_TYPE_40%>><!--</TD>-->
                                <TR>
                                    <TD CLASS="TD5">Role ID</TD>
                                    <TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtRoleID" SIZE=20  MAXLENGTH=20 tag="12XXXU" ALT="Role Id"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRoleID(frm1.txtRoleID.value)">&nbsp;<INPUT TYPE=TEXT NAME="txtRoleNm" SIZE=30 tag="14N"></TD>
                                </TR>
                                <TR>
                                    <TD CLASS="TD5">Menu Type</TD>
                                    <TD CLASS="TD656">
                                    <INPUT TYPE="RADIO" NAME="rdoConf" ID="rdoConf1" CLASS="RADIO" tag="11" value="M">            <LABEL FOR="rdoConf1">Menu(M)</LABEL></SPAN>
                                     <INPUT TYPE="RADIO" NAME="rdoConf" ID="rdoConf2" CLASS="RADIO" tag="11" value="P">            <LABEL FOR="rdoConf2">Program(P)</LABEL></SPAN>
                                       <INPUT TYPE="RADIO" NAME="rdoConf" ID="rdoConf3" CLASS="RADIO" tag="11" value="" Checked>    <LABEL FOR="rdoConf3">All(A)</LABEL></SPAN>
                                    </TD>
                                </TR>
                                <TR>
                                    <TD CLASS="TD5">Menu ID</TD>
                                    <TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtMenuID" SIZE=20 MAXLENGTH=8 tag="11XXXU" ALT="Menu Id"></TD>
                                </TR>                        
                            </TABLE>
                        </FIELDSET>
                    </TD>
                </TR>
                <TR>
                    <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
                </TR>                
                
                <TR>
                    <TD WIDTH=100% HEIGHT=* VALIGN=TOP>
                        <TABLE <%=LR_SPACE_TYPE_60%>>
                            <TR>
                                <TD CLASS="TD5">Role ID</TD>
                                <TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtCd" SIZE=20 MAXLENGTH=10 tag="14N">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtNm" SIZE=30 tag="14N"></TD>
                            </TR>
                            <TR>
                                <TD HEIGHT="100%" COLSPAN=2>
                                <script language =javascript src='./js/zb003ma1_vspdData_vspdData.js'></script></TD>
                            </TR>
                        </TABLE>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR >
        <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
    

    
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hRoleID" tag="24">
</FORM>
    
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
