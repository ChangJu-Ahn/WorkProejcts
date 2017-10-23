<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Role 등록 
*  2. Function Name        : zb002ma1.asp
*  3. Program ID           :
*  4. Program Name         :
*  5. Program Desc         :
*  6. Comproxy List        :
*  7. Modified date(First) : 2000/03/27
*  8. Modified date(Last)  : 2001/02/27
*  9. Modifier (First)     : Kang Doo Sig
* 10. Modifier (Last)      : 
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

Const BIZ_PGM_ID = "zb002mb1.asp"    
Const BIZ_PGM_JUMP_ID = "ZB005MA1"
Const JUMP_PGM_ID = "ZB003MA1"

Dim C_RoleID
Dim C_RoleNm
Dim C_CompstRoleType
Dim C_CopyRole
Dim C_CopyRolePopUp

<!-- #Include file="../../inc/lgvariables.inc" -->    
dim IsOpenPop          

'=========================================================================================================
Sub InitSpreadPosVariables()
    C_RoleID         = 1                                                            
    C_RoleNm         = 2                                                              
    C_CompstRoleType = 3
    C_CopyRole       = 4
    C_CopyRolePopUp  = 5
End Sub

'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           
    lgLngCurRows = 0                            
    lgSortKey = 1
    
End Sub

'=========================================================================================================
sub SetdefaultVal()

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
        .MaxCols = C_CopyRolePopUp + 1                                                    
        .MaxRows = 0

        Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetEdit C_RoleID,        "Role ID",    30,,,20,2
        ggoSpread.SSSetEdit C_RoleNm,        "Role명",    40,,,20    
        ggoSpread.SSSetCombo C_CompstRoleType,        "Composite Role Type",25,2,False
        ggoSpread.SSSetEdit C_CopyRole,        "Copy From",20,2,False ' copy role
        ggoSpread.SSSetButton C_CopyRolePopUp
        .ReDraw = true

        Call SetSpreadLock

        Call ggoSpread.MakePairsColumn(C_RoleID,C_RoleNm,"1")
        Call ggoSpread.MakePairsColumn(C_CopyRole,C_CopyRolePopUp,"1")

        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock    C_RoleID, -1,C_RoleID
        ggoSpread.SpreadLock    C_CopyRole, -1,C_CopyRolePopUp
        .vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired            C_RoleID, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired            C_RoleNm, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired            C_CompstRoleType, pvStartRow, pvEndRow
        .vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Sub InitComboBox()
    Dim strCboData
        
    ggoSpread.Source = frm1.vspdData
    
    strCboData = "Menu Role" & vbTab & "Composite Role"
    ggoSpread.SetCombo strCboData, C_CompstRoleType
End Sub

'=========================================================================================================
Function CookiePage(byval strJumpFlg )

    Dim strTemp
    Dim IntRetCD
    Dim arrSetVal(1)
    Dim ArrVar
    Dim nActiveRow
    
    With frm1
        If .vspdData.MaxRows > 0 then 
        	nActiveRow = .vspdData.ActiveRow
            arrSetVal(0) = GetSpreadText(.vspdData, 1, nActiveRow, "X", "X")
            arrSetVal(1) = GetSpreadText(.vspdData, 2, nActiveRow, "X", "X")
        End If 
    End With
            
    If strJumpFlg =0 Then 
        IntRetCD = CommonQueryRs("COMPST_ROLE_TYPE","Z_USR_ROLE","USR_ROLE_ID =  " & FilterVar(arrSetVal(0), "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        lgF0 = Replace(lgF0, Chr(11), vbTab)
        lgF0 = Replace(lgF0, " ","")        
        
        If lgF0="" Then 
            arrSetVal(0)=""
            arrSetVal(1)=""
            Call PgmJump(BIZ_PGM_JUMP_ID)
        Else         
            If Trim(lgF0) = 1 Then     
                WriteCookie "ZB_02_Role_ID" , Trim(arrSetVal(0))
                WriteCookie "ZB_02_Role_NM" , Trim(arrSetVal(1))            
                Call PgmJump(BIZ_PGM_JUMP_ID)
            Else
                WriteCookie "ZB_02_Role_ID" , ""
                WriteCookie "ZB_02_Role_NM" , ""

                IntRetCD = DisplayMsgBox("211425", "x",arrSetVal(0),"x")                '데이타가 변경되었습니다. 조회하시겠습니까?
                If IntRetCD = vbNo Then
                      Exit Function
                End If                
            End If
        End If
    Else
        IntRetCD = CommonQueryRs("COMPST_ROLE_TYPE","Z_USR_ROLE","USR_ROLE_ID =  " & FilterVar(arrSetVal(0), "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        lgF0 = Replace(lgF0, Chr(11), vbTab)
        lgF0 = Replace(lgF0, " ","")
                
        If lgF0="" Then 
            arrSetVal(0)=""
            arrSetVal(1)=""
            Call PgmJump(JUMP_PGM_ID)
        Else         
            If Trim(lgF0) = 0 Then     
                WriteCookie "ZB_05_Cmp_Role_ID" , Trim(arrSetVal(0))
                Call PgmJump(JUMP_PGM_ID)
            Else
                WriteCookie "ZB_05_Cmp_Role_ID" , ""
                
                IntRetCD = DisplayMsgBox("211426", "x",arrSetVal(0),"x")                '데이타가 변경되었습니다. 조회하시겠습니까?
                If IntRetCD = vbNo Then
                      Exit Function
                End If    
            End If
        End IF
    End If
    
End Function

'=========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_RoleID         =  iCurColumnPos(1)
            C_RoleNm         =  iCurColumnPos(2)
            C_CompstRoleType =  iCurColumnPos(3)
            C_CopyRole       =  iCurColumnPos(4)
            C_CopyRolePopUp  =  iCurColumnPos(5)

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
sub Form_Load()
    
    Call LoadInfTB19029                                                         
    Call ggoOper.LockField(Document, "N")                                   
    Call InitSpreadSheet                                                    
    Call InitVariables  
    
    Call InitComboBox
                                                        
    Call SetToolbar("1100110100101111")                                        
    frm1.txtRoleID.focus 
    
End Sub

'=========================================================================================================
sub Form_QueryUnload(Cancel , UnloadMode )
  
End Sub


'=========================================================================================================
sub vspdData_Change( Col ,  Row )
    
    If Col = C_CompstRoleType Then 
        With frm1.vspdData
            .SetText C_CopyRole, Row, ""
        End With
    End If
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    lgBlnFlgChgValue = True
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
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

    Dim strTemp
    Dim intPos1
   
    With frm1.vspdData 
		.Row = Row
	    ggoSpread.Source = frm1.vspdData
   
        If Row > 0 Then
            Select Case Col    
                Case C_CopyRolePopUp
                    Call OpenRoleInfo(GetSpreadText(frm1.vspdData, C_CopyRole, Row, "X", "X"), 3)
            End Select
        End If
    End With
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
Sub vspdData_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    
    If CheckRunningBizProcess = True Then                    
        Exit Sub
    End If
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then    
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
    Call InitComboBox
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
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")                '데이타가 변경되었습니다. 조회하시겠습니까?
        If IntRetCD = vbNo Then
          Exit Function
        End If
    End If
    

    Call ggoOper.ClearField(Document, "2")                                        
    Call ggoSpread.ClearSpreadData()            
    Call InitVariables
    
    Call DbQuery                                                                
       
    FncQuery = True                                                                
    
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

    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               

    Dim strVal
    
    With frm1
   
    If lgIntFlgMode = Parent.OPMD_UMODE Then
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                            
        strVal = strVal & "&txtRoleID=" & Trim(.hRoleID.value)                        
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                            
        strVal = strVal & "&txtRoleID=" & Trim(.txtRoleID.value)                        
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If    

    Call RunMyBizASP(MyBizASP, strVal)                                        
    

    Call SetToolbar("1100111100111111")
    
    End With
    
    DbQuery = True
    
End Function

'=========================================================================================================
Function DbQueryOk()                                                        
    
 
    lgIntFlgMode = Parent.OPMD_UMODE                                                
    
    Call ggoOper.LockField(Document, "Q")                                    
    Call SetToolbar("1100111100111111")
    
    frm1.vspddata.Focus  
End Function

'=========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
    Err.Clear                                                               
    On Error Resume next                                                   
    

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")                 'Will you destory previous data
        If IntRetCD = vbNo Then
            Exit Function
        End If
       
    End If
    
    Call ggoOper.ClearField(Document, "A")
    Call ggoSpread.ClearSpreadData        
    Call ggoOper.LockField(Document, "N")                                          
    Call InitVariables                                                      
    Call SetDefaultVal
    
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
    
    .Col = 8
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

    FncCancel = False 
    
    If frm1.vspdData.MaxRows < 1 Then
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData    
    ggoSpread.EditUndo                                        
    FncCancel = True
    
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
Function FncCopy() 
	Dim nActiveRow
    With frm1
        If .vspdData.ActiveRow > 0 Then
            .vspdData.focus
            .vspdData.ReDraw = False
                
            ggoSpread.Source = frm1.vspdData    
            ggoSpread.CopyRow
            nActiveRow = frm1.vspdData.ActiveRow
            SetSpreadColor nActiveRow, nActiveRow
            
            .vspdData.SetText C_RoleID, nActiveRow, ""
            
            .vspdData.ReDraw = True
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
    Dim iColSep, iRowSep

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep
        
    DbSave = False                                                          
    
    Call LayerShowHide(1)
    
    On Error Resume next                                                   

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
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_RoleID, lRow, "X", "X")) & iColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_RoleNm, lRow, "X", "X")) & iColSep
                If Trim(GetSpreadText(.vspdData, C_CompstRoleType, lRow, "X", "X")) = "Composite Role" Then 
                    strVal = strVal & "1" & iColSep
                Else 
                    strVal = strVal & "0" & iColSep
                End if 
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_CopyRole, lRow, "X", "X")) & iRowSep
                
                lGrpCnt = lGrpCnt + 1
            End if                
         Next
         
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
    ggoSpread.SSDeleteFlag 0
    Call SetSpreadLock()

End Function

'=========================================================================================================
function DbDelete() 
End Function

'=========================================================================================================
'    Name : OpenRole()
'    Description : Role PopUp
'=========================================================================================================
Function OpenRole( strCode,  iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "Role Popup"                    ' 팝업 명칭 
    arrParam(1) = "Z_USR_ROLE"                        ' TABLE 명칭 
    arrParam(2) = strCode                            ' Code Condition
    arrParam(3) = ""                                ' Name Cindition
    arrParam(4) = ""                                ' Where Condition
    arrParam(5) = "Role ID"            
    
    arrField(0) = "USR_ROLE_ID"                        ' Field명(0)
    arrField(1) = "USR_ROLE_NM"                        ' Field명(1)
    
    arrHeader(0) = "Role ID"                    ' Header명(0)
    arrHeader(1) = "Role Name"                    ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    IsOpenPop = False
    
    frm1.txtRoleID.focus 
    
    
    If arrRet(0) = "" Then
    Else
        Call SetRole(arrRet)
    End If    
    frm1.txtRoleID.focus
    Set gActiveElement = document.activeElement

End Function

'=========================================================================================================
'    Name : OpenRoleInfo()
'    Description : Role PopUp
'=========================================================================================================
Function OpenRoleInfo(Byval strCode, Byval iWhere)

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)
    Dim RoleType

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    
    arrParam(0) = "Role Popup"
    arrParam(1) = "Z_USR_ROLE"    
    arrParam(2) = strCode
    arrParam(3) = ""    
    
    
    frm1.vspdData.Col = C_CompstRoleType
    
    If frm1.vspdData.Text = "Composite Role" Then 
        RoleType = "1"
    Else
        RoleType= "0"
    End If
                    
    arrParam(4) = "COMPST_ROLE_TYPE= " & FilterVar(RoleType, "''", "S") & ""                                ' Where Condition
                    
    arrParam(5) = "Role ID"
    
    arrField(0) = "USR_ROLE_ID"
    arrField(1) = "USR_ROLE_NM"
    
    arrHeader(0) = "Role ID"
    arrHeader(1) = "Role Name"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetRoleInfo(arrRet, iWhere)
    End If    

    
End Function
'=========================================================================================================
Function SetRoleInfo(Byval arrRet, Byval iWhere)
    Select Case iWhere
    Case  1
    Case  2
        With frm1.vspdData
            .Col = C_CopyRole
            .Text = arrRet(0)
            Call vspdData_Change(.Col, .Row)
        End With
    Case  3
        With frm1.vspdData
            .Col = C_CopyRole
            .Text = arrRet(0)
            Call vspdData_Change(.Col, .Row)
        End With
    End Select
End Function
'=========================================================================================================
'    Name : OpenDetailRole()
'    Description : Role ID Detail Information
'=========================================================================================================
Function OpenDetailRole()

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6), arrSetVal(1)
    Dim IntRetCD
    Dim ArrVar
    Dim iCalledAspName
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZB002PA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZB002PA1", "x")
        IsOpenPop = False
        Exit Function
    End If
        
    With frm1
        If .vspdData.MaxRows > 0 then 
            arrSetVal(0) = GetSpreadText(.vspdData, 1, .vspdData.ActiveRow, "X", "X")
            arrSetVal(1) = GetSpreadText(.vspdData, 2, .vspdData.ActiveRow, "X", "X")
        Else
            Call DisplayMsgBox ("900002","x","x","x")
            IsOpenPop = False
            Exit Function
        End If            
    End With        
    
    If Trim(arrsetval(0)) ="" Then 		
		IsOpenPop = False
		exit function
    End If 
    
    IntRetCD = CommonQueryRs("compst_role_type","z_usr_role","Usr_Role_Id =  " & FilterVar(arrSetVal(0), "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    lgF0 = Replace(lgF0, Chr(11), vbTab)
    lgF0 = Replace(lgF0, " ","")
    
    If lgF0 = 1 then 
        IntRetCD = DisplayMsgBox("211424", "x","x","x")                '데이타가 변경되었습니다. 조회하시겠습니까?
        If IntRetCD = vbNo Then
            IsOpenPop = False
              Exit Function
        End If
        
        IsOpenPop = False
        Exit Function
    End If 

    arrParam(0) = "Role ID 상세"                                                                        ' 팝업 명칭 
    arrParam(1) = "z_usr_role_mnu_authztn_asso a, z_lang_co_mast_mnu b, b_minor c, b_minor d"                ' TABLE 명칭 
    arrParam(2) = ""                                                                                        ' Code Condition
    arrParam(3) = ""                                                                                        ' Name Cindition
    arrParam(4) = "b.lang_cd= " & FilterVar(Parent.gLang, "''", "S") & " and a.usr_role_id =  " & FilterVar(arrSetVal(0), "''", "S") & " and a.mnu_id = b.mnu_id" _
                    & " And a.mnu_type = c.minor_cd And c.major_cd=" & FilterVar("z0006", "''", "S") & " " _
                    & " And a.action_id = d.minor_cd And d.major_cd=" & FilterVar("z0013", "''", "S") & " "                                ' Where Condition
    
    arrParam(5) = "Role ID"            
    
    arrField(0) = "a.mnu_id"                                                            ' Field명(0)
    arrField(1) = "b.mnu_nm"                                                            ' Field명(1)
    arrField(2) = "c.minor_nm"                                                            ' Field명(2)
    arrField(3) = "d.minor_nm"                                                            ' Field명(3)
    
    arrHeader(0) = "Menu ID"                                                        ' Header명(0)
    arrHeader(1) = "Menu명"                                                            ' Header명(1)
    arrHeader(2) = "Menu Type"                                                        ' Header명(2)
    arrHeader(3) = "Action"            

    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam, arrField, arrHeader, arrSetVal), _
        "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetRole(arrRet)
    End If    

End Function
'=========================================================================================================
Function OpenUsedRole()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZB002RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZB002RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
    
    With frm1
        If .vspdData.ActiveRow > 0 then 
            arrParam(0) = GetSpreadText(.vspdData, 1, .vspdData.ActiveRow, "X", "X")
            arrParam(1) = GetSpreadText(.vspdData, 2, .vspdData.ActiveRow, "X", "X")
        Else
            Call DisplayMsgBox ("900002","x","x","x")
            IsOpenPop = False
            Exit Function
        End If            
    End With    
    
    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0, 0) = "" Then
        Exit Function
    Else
        Call SetRefComposite(arrRet)
    End If
    
End Function
'=========================================================================================================
'    Name : SetRole()
'    Description : Role Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetRole( arrRet)
    frm1.txtRoleID.Value = arrRet(0)
    frm1.txtRoleNm.Value = arrRet(1)
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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Role 등록</font></td>    
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>   
                            </TR>   
                        </TABLE>   
                    </TD>   
                    <TD WIDTH=* align=right>
                        <a href="vbscript:OpenDetailRole">Role ID 상세</A>&nbsp;|&nbsp<a href="vbscript:OpenUsedRole">사용자별 Role Assign 현황</A></TD> 
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
                                    <TD CLASS="TD5">Role ID</TD>    
                                    <TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtRoleID" SIZE=20 MAXLENGTH=20 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoleID" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRole(frm1.txtRoleID.value,0)">&nbsp;<INPUT TYPE=TEXT NAME="txtRoleNm" SIZE=30 tag="14N"></TD>                                       
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
                                <script language =javascript src='./js/zb002ma1_vspdData_vspdData.js'></script></TD>   
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
    <TR HEIGHT=20>   
        <TD WIDTH=100%>   
            <TABLE <%=LR_SPACE_TYPE_30%>>   
                <TR>   
                <TD WIDTH=50%>&nbsp;</TD>   
                <TD WIDTH=50%>   
                    <TABLE WIDTH=100%>                           
                        <TD WIDTH=* Align=right><A href="vbscript:CookiePage(1)">메뉴 Assignment to Role</A>&nbsp;|&nbsp;<A href="vbscript:CookiePage(0)">Role Assignment to Composite Role</A></TD>                                                                                     
                        <TD WIDTH=10>&nbsp;</TD>                           
                    </TABLE>   
                </TD>   
                </TR>   
            </TABLE>   
        </TD>   
    </TR>       
    <TR>   
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 Tabindex=-1></IFRAME>   
        </TD>   
    </TR>   
</TABLE>   
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" rows="1" cols="20"></TEXTAREA>   
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">    
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">    
<INPUT TYPE=HIDDEN NAME="hRoleID" tag="24">    

</FORM>   
    <DIV ID="MousePT" NAME="MousePT">   
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>   
    </DIV>   
</BODY>   
</HTML>
