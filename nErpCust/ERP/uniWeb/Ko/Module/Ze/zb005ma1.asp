<%@ LANGUAGE="VBSCRIPT"%>
<!--
======================================================================================================
*  1. Module Name          : ZB005MA1    COMPOSITE ROLE별 ROLE ASSIGN
*  2. Function Name        : COMPOSITE ROLE별 ROLE ASSIGN
*  3. Program ID           : ZB005MA1
*  4. Program Name         : 
*  5. Program Desc         : COMPOSITE ROLE별 ROLE ASSIGN 화면처리 ASP
*  6. Comproxy List        : +ZB0081_CTRL_ROLE_COMPST_ROLE_ASSO
*                              +ZB0028_LIST_USR_ROLE
*                              +ZB0088_LIST_USR_ROLE_COMPST_ROLE_ASSO
*                              +ZB0039_LOOKUP_COMPST_ROLE
*  7. Modified date(First) : 2000/04/04
*  8. Modified date(Last)  : 2001/12/05 
*  9. Modifier (First)     : Kang Tae Bum
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
Const BIZ_PGM_ID = "ZB005MB1.asp"                                                
Const JUMP_PGM_ID = "ZB003MA1"
'=========================================================================================================
Dim C_RoleID
Dim C_RoleNm
'=========================================================================================================

<!-- #Include file="../../inc/lgvariables.inc" -->    
dim    lgStrQueryFlag

'=========================================================================================================
dim IsOpenPop

'=========================================================================================================
Sub InitSpreadPosVariables()
    C_RoleID = 1
    C_RoleNm = 2
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
        .MaxCols = C_RoleNm + 1                                                
        .MaxRows = 0

        Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetEdit C_RoleID, "Role ID", 40
        ggoSpread.SSSetEdit C_RoleNm, "Role Name", 78
        .ReDraw = true

        Call SetSpreadLock
        
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_RoleID, -1, -1
    ggoSpread.spreadLock C_RoleNm, -1, -1
    .vspdData.ReDraw = True

    End With
End Sub

'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetProtected    C_RoleID, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected    C_RoleNm, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

'=========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case Ucase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_RoleID        =  iCurColumnPos(1)
            C_RoleNm         =  iCurColumnPos(2)

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
    Call SetToolbar("1100101100101111")                                            
    Call CookiePage(0)
    
    frm1.txtCompositeID.focus
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
sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 
        If Row >= NewRow Then
            Exit Sub
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
    
    lgStrQueryFlag = ""
    
     '----------  Coding part  -------------------------------------------------------------   
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
        IntRetCD = DisplayMsgbox("900015",Parent.VB_YES_NO,"x","x")                  ' 데이타가 변경되었습니다. 신규작업을 하시겠습니까?
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
        Call DisplayMsgbox("900002","x","x","x")                            ' 조회를 먼저 하십시오.
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
    

    If lgBlnFlgChgValue = False Then
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
        Call DisplayMsgBox("183104","x","x","x")           '☆: you must release this line if you change msg into code
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
    
    FncPrev = False                                                        
    
    Err.Clear                                                               
    
    lgStrQueryFlag = "P"
                                                               

    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013",Parent.VB_YES_NO,"x","x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    

    If Not chkField(Document, "1") Then                                    
       Exit Function
    End If
    

    If DBQuery = False Then                                                        
        Exit Function 
    End If            
    FncPrev = True                                                                
    
End Function
'=========================================================================================================
Sub EraseContents()
    Call ggoOper.ClearField(Document, "2")                                        
    Call InitVariables                                                                
End Sub

'=========================================================================================================
Function FncNext() 
    Dim IntRetCD 
    
    FncNext = False                                                        
    
    Err.Clear                                                               

    lgStrQueryFlag = "N"


    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If


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
            .vspdData.focus
            .vspdData.ReDraw = False
                
            ggoSpread.Source = frm1.vspdData    
            ggoSpread.CopyRow
            SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
                
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
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
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
    Dim IntRetCD

    Call LayerShowHide(1)

    DbQuery = False
                    
    Err.Clear                                                               

    
    Dim strval
    
    With frm1
      
    
    IntRetCD = CommonQueryRs("usr_role_nm","Z_Usr_Role","usr_role_id = '" & FilterVar(Trim(.txtCompositeID.value),"","SNM") & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = Replace(lgF0, Chr(11), vbTab)
    lgF0 = Replace(lgF0, " ","")
    
    .txtCompositeNm.value = lgF0
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                          
        strVal = strVal & "&txtCompositeID=" & Trim(.hCompositeID.value)                                
        strVal = strVal & "&txtCompositeNm=" & Trim(.hCompositeNm.value)                                
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        strVal = strVal & "&txtPrvNext=" & lgStrQueryFlag 
    Else
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                          
        strVal = strVal & "&txtCompositeID=" & Trim(.txtCompositeID.value)                
        strVal = strVal & "&txtCompositeNm=" & Trim(.txtCompositeNm.value)                
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        strVal = strVal & "&txtPrvNext=" & lgStrQueryFlag
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
    Dim iColSep, iRowSep

    iColSep = parent.gColSep
    iRowSep = parent.gRowSep    
    
    DbSave = False      
        
    If frm1.txtCompositeID.value = "" Then
        Call DisplayMsgBox("970021",,"Composite Role ID","x")
        Exit Function
    End If
    
    Call LayerShowHide(1)                                                    
       
    
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
                
                If .txtCd.Value <> "" Then
                    strVal = strVal & Trim(.txtCd.Value) & iRowSep
                Else
                    strVal = strVal & Trim(.txtCompositeID.Value) & iRowSep
                End If
               
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
    Call ggoOper.ClearField(Document, "2")                                    
    Call ggoSpread.ClearSpreadData        
    Call MainQuery()
End Function

'=========================================================================================================
function DbDelete() 
End Function

'=========================================================================================================
'    Name : OpenComposite()
'    Description : Composite Role PopUp
'=========================================================================================================
Function OpenComposite(Byval strCode, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "Composite Role 팝업"                    ' 팝업 명칭 
    arrParam(1) = "Z_USR_ROLE"                            ' TABLE 명칭 
    arrParam(2) = strCode                                    ' Code Condition
    arrParam(3) = ""                                        ' Name Cindition
    arrParam(4) = "COMPST_ROLE_TYPE ='1'"                                        ' Where Condition
    arrParam(5) = "Composite Role ID"            
    
    arrField(0) = "USR_ROLE_ID"                            ' Field명(0)
    arrField(1) = "USR_ROLE_NM"                            ' Field명(1)
    
    arrHeader(0) = "Composite Role ID"                    ' Header명(0)
    arrHeader(1) = "Composite Role명"                    ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
             "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetComposite(arrRet)
    End If    
    
    frm1.txtCompositeID.focus 
    Set gActiveElement = document.activeElement
    
End Function
'=========================================================================================================
Function OpenDetailRole()

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6), arrSetVal(1)
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZB002PA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZA003PA2", "x")
        IsOpenPop = False
        Exit Function
    End If
        
    With frm1
        If .vspdData.MaxRows > 0 then 
            arrSetVal(0) = GetSpreadText(.vspdData, 1, .vspdData.ActiveRow, "X", "X")
            arrSetVal(1) = GetSpreadText(.vspdData, 2, .vspdData.ActiveRow, "X", "X")
        Else
			If frm1.txtCd.value ="" Then 
				Call DisplayMsgBox ("900002","x","x","x")                    ' 조회를 먼저 하십시오.
				IsOpenPop = False
				Exit Function    
			End If
        End If            
    End With        

    arrParam(0) = "Role ID 상세"                                                                        ' 팝업 명칭 
    arrParam(1) = "z_usr_role_mnu_authztn_asso a, z_lang_co_mast_mnu b, b_minor c, b_minor d"                ' TABLE 명칭 
    arrParam(2) = ""                                                                                        ' Code Condition
    arrParam(3) = ""                                                                                        ' Name Cindition
    arrParam(4) = "b.lang_cd='" & Parent.gLang & "' and a.usr_role_id = '" & arrSetVal(0) & "' and a.mnu_id = b.mnu_id" _
                    & " And a.mnu_type = c.minor_cd And c.major_cd='z0006' " _
                    & " And a.action_id = d.minor_cd And d.major_cd='z0013' "                                ' Where Condition

    arrParam(5) = "Role ID"            
    
    arrField(0) = "a.MNU_ID"                                                                                ' Field명(0)
    arrField(1) = "B.MNU_NM"                                                                                ' Field명(1)
    arrField(2) = "c.minor_nm"                                                                                ' Field명(2)    
    arrField(3) = "d.minor_nm"                                                                                ' Field명(3)
    
    arrHeader(0) = "Menu ID"                                                                            ' Header명(0)
    arrHeader(1) = "Menu명"                                                                                ' Header명(1)
    arrHeader(2) = "Menu Type"                                                                            ' Header명(2)
    arrHeader(3) = "Action"                                                                                ' Header명(3)
    
    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader, arrSetVal), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetRole(arrRet)
    End If    
    
End Function
'=========================================================================================================
Function OpenRefRole()
    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZB005RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZB005RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
    
    If frm1.txtCompositeID.Value = "" Then
        call DisplayMsgBox ("900002","X","X","X")
        IsOpenPop = False
        Exit Function
    End If
    
    With frm1
        If .vspdData.ActiveRow > 0 then 
            arrParam(0) = GetSpreadText(.vspdData, 1, .vspdData.ActiveRow, "X", "X")
            arrParam(1) = GetSpreadText(.vspdData, 2, .vspdData.ActiveRow, "X", "X")
        
        End If            
    End With    

    arrParam(2) = frm1.txtCompositeID.value
    
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
Function CookiePage(byval nflag)

    Dim strTemp
    Dim arrParam(1)
    
    With frm1
        If .vspdData.ActiveRow > 0 then 
            arrParam(0) = GetSpreadText(.vspdData, 1, .vspdData.ActiveRow, "X", "X")
            arrParam(1) = GetSpreadText(.vspdData, 2, .vspdData.ActiveRow, "X", "X")
        End If            
    End With    
    
    If nflag=1 Then
        WriteCookie "ZB_05_Cmp_Role_ID" , Trim(arrParam(0))
        Call PgmJump(JUMP_PGM_ID)
    Else
        strTemp = ReadCookie("ZB_02_Role_ID")
        WriteCookie "ZB_02_Role_ID" ,""
        If Trim(strTemp)<>"" Then 
            frm1.txtCompositeID.value = strTemp
            Call FncQuery
        End If 
    End If
    
    
End Function

'=========================================================================================================
'    Name : SetComposite()
'    Description : Plant Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetComposite(byval arrRet)
    frm1.txtCompositeID.Value = arrRet(0)
    frm1.txtCompositeNm.Value  = arrRet(1)
End Function

'=========================================================================================================
'    Name : SetRefComposite()
'    Description : Plant Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetRefComposite(Byval arrRet)
    
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
            .vspdData.SetText C_RoleID, I+1, arrRet(I - TempRow, 0)
            .vspdData.SetText C_RoleNm, I+1, arrRet(I - TempRow, 1)
        Next    
        
        ggoSpread.SpreadUnlock 1, TempRow + 1, 2, .vspdData.MaxRows 
        ggoSpread.ssSetProtected 1, TempRow + 1, .vspdData.MaxRows
        ggoSpread.ssSetProtected 2, TempRow + 1, .vspdData.MaxRows  
        
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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Role Assign.to Composite Role</font></td> 
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* align=right>
                                <a href="vbscript:OpenDetailRole">Role ID 상세</A>&nbsp;|&nbsp;<a href="vbscript:OpenRefRole">Role ID 참조</A>   
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
                            <TABLE <%=HEIGHT_TYPE_02%>>
                                <TR>
                                    <TD CLASS="TD5">Composite Role ID</TD> 
                                    <TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtCompositeID" SIZE=20  MAXLENGTH=20 tag="12XXXU"  ALT="Composite Role ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenComposite(frm1.txtCompositeID.value,0)">&nbsp;<INPUT TYPE=TEXT NAME="txtCompositeNm" SIZE=30 tag="14N"></TD>
                                </TR>
                                </TABLE>
                        </FIELDSET>
                    </TD>
                </TR>
                <TR>
                    <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
                </TR>
                <TR>
                    <TD HEIGHT=20 WIDTH=100%>
                        <TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                        <TR>
                            <TD CLASS="TD5">Composite Role ID</TD> 
                            <TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtCd" SIZE=20 MAXLENGTH=20 tag="14N">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtNm" SIZE=30 tag="14N"></TD>
                        </TR>
                        </TABLE>
                    </TD>
                </TR>
            
            <TR>
                <TD WIDTH=100% HEIGHT=* valign=top>
                    <TABLE <%=LR_SPACE_TYPE_20%>>
                        <TR>
                            <TD HEIGHT="100%">
                            <script language =javascript src='./js/zb005ma1_vspdData_vspdData.js'></script></TD>
                        </TR>
                    </TABLE>
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
                        <TD WIDTH=* Align=right><A href="vbscript:CookiePage(1)">메뉴 Assignment to Role</A></TD> 
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
<INPUT TYPE=HIDDEN NAME="hCompositeID" tag="24"> 
<INPUT TYPE=HIDDEN NAME="hCompositeNm" tag="24"> 
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>

