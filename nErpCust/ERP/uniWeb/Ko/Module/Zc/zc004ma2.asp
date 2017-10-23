
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


Const BIZ_PGM_ID = "zc004mb2.asp"
Const JUMP_PGM_ID = "ZB003MA1"

Dim C_LangCD
Dim C_LangPopup'khy200307
Dim C_MnuID
Dim C_MnuNm
Dim C_MnuType
Dim C_SysLvl
Dim C_CallFrmID
Dim C_MnuSeq
Dim C_UseYN
Dim C_UpperMnuID
Dim C_UpperMnuPopUp

<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim IsOpenPop
'=========================================================================================================
Sub InitSpreadPosVariables()
    C_LangCD        = 1
    C_LangPopup     = 2
    C_MnuID         = 3
    C_MnuNm         = 4
    C_MnuType       = 5
    C_SysLvl        = 6
    C_CallFrmID     = 7
    C_MnuSeq        = 8
    C_UseYN         = 9
    C_UpperMnuID    = 10
    C_UpperMnuPopUp = 11
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

    Call InitSpreadPosVariables()
    
    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)

        .ReDraw = false                   
        .MaxCols = C_UpperMnuPopUp +1
        .MaxRows = 0

        Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetEdit C_LangCD, "언어코드", 11,,,15, 2
        ggoSpread.SSSetButton C_LangPopUp'khy200307
        ggoSpread.SSSetEdit    C_MnuID, "메뉴ID", 14,,,15, 2
        ggoSpread.SSSetEdit C_MnuNm, "메뉴명", 30,,,40
        ggoSpread.SSSetCombo C_MnuType, "메뉴타입", 10, 2, false
        ggoSpread.SSSetEdit C_SysLvl, "메뉴레벨", 10,,,1,2
        ggoSpread.SSSetEdit C_CallFrmID, "호출ID", 14,,,15,2
        ggoSpread.SSSetEdit C_MnuSeq, "메뉴순서", 11,,,10,2        
        ggoSpread.SSSetCheck C_UseYN, "사용여부", 12,,,True
        ggoSpread.SSSetEdit C_UpperMnuID, "상위메뉴ID", 15,,,15,2
        ggoSpread.SSSetButton C_UpperMnuPopUp
        'ggoSpread.SSSetSplit2(4)        
        .ReDraw = true

        Call SetSpreadLock
        
        Call ggoSpread.MakePairsColumn(C_LangCD,C_LangPopup,"1")'khy200307
        Call ggoSpread.MakePairsColumn(C_MnuID,C_MnuNm,"1")
        Call ggoSpread.MakePairsColumn(C_UpperMnuID,C_UpperMnuPopUp,"1")

        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
    
        .vspdData.ReDraw = False
        
        ggoSpread.SpreadLock C_LangCD, -1, C_LangPopup'khy200307
        ggoSpread.SpreadLock C_MnuID, -1, C_MnuID
        ggoSpread.SSSetRequired C_MnuNm, -1
        ggoSpread.SpreadLock C_MnuType,-1,C_MnuType
        ggoSpread.SpreadLock C_SysLvl,-1,C_SysLvl        
        ggoSpread.spreadlock C_UpperMnuID, -1,C_UpperMnuPopUp

        .vspdData.ReDraw = True    

    End With
End Sub
'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired        C_LangCD, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired        C_MnuID, pvStartRow, pvEndRow        
        ggoSpread.SSSetRequired        C_MnuNm, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired        C_MnuType, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired        C_SysLvl, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired        C_CallFrmID, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired        C_UpperMnuID, pvStartRow, pvEndRow        
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
    
            C_LangCD        =  iCurColumnPos(1)
            C_LangPopup     =  iCurColumnPos(2)
            C_MnuID         =  iCurColumnPos(3)
            C_MnuNm         =  iCurColumnPos(4)
            C_MnuType       =  iCurColumnPos(5)
            C_SysLvl        =  iCurColumnPos(6)
            C_CallFrmID     =  iCurColumnPos(7)
            C_MnuSeq        =  iCurColumnPos(8)
            C_UseYN         =  iCurColumnPos(9)
            C_UpperMnuID    =  iCurColumnPos(10)
            C_UpperMnuPopUp =  iCurColumnPos(11)

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

    Call SetCombo(frm1.cboMnuType, "M", "메뉴")
    Call SetCombo(frm1.cboMnuType, "P", "프로그램")
    
    Call SetCombo(frm1.cboUseYN, "1", "사용")
    Call SetCombo(frm1.cboUseYN, "0", "미사용")    
    
End Sub

'=========================================================================================================
Sub InitSpreadComboBox()

    Dim strCboData
    Dim IntRetCD

    ggoSpread.Source = frm1.vspdData
 
    'IntRetCD = CommonQueryRs("lang_cd","b_language","LANG_CD >= ''",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    'lgF0 = Replace(lgF0, Chr(11), vbTab)    
    'lgF0 = Replace(lgF0, " ","")
    'ggoSpread.SetCombo lgF0, C_LangCD
    
    strCboData = "M" & vbTab & "P"
    ggoSpread.SetCombo strCboData, C_MnuType

End Sub
'=========================================================================================================
Sub Form_Load()

    Dim IntRetCD
    'Call LoadInfTB19029  
    Call ggoOper.LockField(Document, "N")

    Call InitSpreadSheet
    
    Call InitVariables

    Call InitComboBox
    Call InitSpreadComboBox
    Call SetDefaultVal
    Call SetToolbar("11001101001111")
    
    frm1.txtMnuID.focus
        
    frm1.txtLangCd.focus
    frm1.txtLangCd.Value = Parent.gLang
	Set gActiveElement = document.activeElement
    
    IntRetCD = CommonQueryRs("Lang_Nm","B_LANGUAGE","Lang_Cd =  " & FilterVar(Parent.gLang , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)        
    lgF0 = Replace(lgF0, Chr(11), "")    'unusual case    
    'lgF0 = Replace(lgF0," ","")            
    frm1.txtLangNm.value = Trim(lgF0)
    
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
   
    With frm1.vspdData 
	    ggoSpread.Source = frm1.vspdData
	   	.Row = Row
	    If Row > 0 Then
	        Select Case Col
				Case C_LangPopup'khy200307
					Call OpenLangInfo(GetSpreadText(frm1.vspdData, C_LangCD, Row, "X", "X"))
	            Case C_UpperMnuPopUp
	                Call OpenMnuInfo(GetSpreadText(frm1.vspdData, C_UpperMnuID, Row, "X", "X"), 3)
	        End Select
	        
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
    
    		frm1.vspdData.SetText C_MnuID, nActiveRow, ""
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
    
    frm1.txtLangNm.value = ""
    IntRetCD = CommonQueryRs("Lang_Nm","B_LANGUAGE","Lang_Cd =  " & FilterVar(frm1.txtLangCd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    lgF0 = Replace(lgF0, Chr(11), "")'unusual case
    'lgF0 = Replace(lgF0, " ","")    
    
    If lgF0 = "" then 
        IntRetCD = DisplayMsgBox("211432", "x", "x", "x")        
        Exit Function
    End if     
    frm1.txtLangNm.value = Trim(lgF0)
        
    DbQuery = False
    
    Call LayerShowHide(1)    
    
    Err.Clear

    Dim strVal    
    
    With frm1
    
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then    
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
        
        strVal = strVal & "&txtLangCd=" & Trim(.hLangCd.value)
        strVal = strVal & "&txtMnuID=" & Trim(.hMnuID.value)
        strVal = strVal & "&cboMnuType=" & Trim(.hMnuType.value)                
        strVal = strVal & "&cboUseYN=" & Trim(.hUseYN.value)
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    Else    
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
        
        strVal = strVal & "&txtLangCd=" & Trim(.txtLangCd.value)        
        strVal = strVal & "&txtMnuID=" & Trim(.txtMnuID.value)
        strVal = strVal & "&cboMnuType=" & Trim(.cboMnuType.value)        
        strVal = strVal & "&cboUseYN=" & Trim(.cboUseYN.value)
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
    Call SetToolbar("11001111001111")

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
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_LangCD, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuID, lRow, "X", "X")) & iColSep          
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuNm, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuType, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_SysLvl, lRow, "X", "X")) & iColSep  
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_CallFrmID, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_UseYN, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_UpperMnuID, lRow, "X", "X")) & iRowSep

                    'validate check----------------------------------------------------------------------
                    If CheckNumeric(Trim(GetSpreadText(.vspdData, C_SysLvl, lRow, "X", "X")))=1 then          
                        IntRetCD = DisplayMsgBox("211433", "x","x","x")                'Menu sequence only decimal
                        Call LayerShowHide(0)                                                
                        Exit Function        
                    End If 
                    If Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X")) > "" then
                        If CheckNumeric(Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X")))=1 then          
                            IntRetCD = DisplayMsgBox("211430", "x","x","x")                'Menu sequence only decimal
                                                    Call LayerShowHide(0)
                            Exit Function        
                        End If 
                    End If
                    '---------------------------------------------------------------------------------------
                    
                    lGrpCnt = lGrpCnt + 1
                    
                Case ggoSpread.UpdateFlag

                    strVal = strVal & "U" & iColSep & lRow & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_LangCD, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuID, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuNm, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuType, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_SysLvl, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_CallFrmID, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_UseYN, lRow, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_UpperMnuID, lRow, "X", "X")) & iRowSep
                    'validate check----------------------------------------------------------------------
                    If CheckNumeric(Trim(GetSpreadText(.vspdData, C_SysLvl, lRow, "X", "X")))=1 then          
                        IntRetCD = DisplayMsgBox("211430", "x","x","x")                'Menu sequence only decimal
                        Call LayerShowHide(0)                        
                        Exit Function        
                    End If 
                    If CheckNumeric(Trim(GetSpreadText(.vspdData, C_MnuSeq, lRow, "X", "X")))=1 then          
                        IntRetCD = DisplayMsgBox("211430", "x","x","x")                'Menu sequence only decimal
                        Call LayerShowHide(0)                                                
                        Exit Function        
                    End If 
                    '---------------------------------------------------------------------------------------
                    
                    lGrpCnt = lGrpCnt + 1
                    
                Case ggoSpread.DeleteFlag
					If Trim(GetSpreadText(.vspdData, C_MnuType, lRow, "X", "X")) = "M" Then 											
						IntRetCD = DisplayMsgBox("211438", Parent.VB_YES_NO, "x", "x") 						
						If IntRetCD = vbNo Then							
							Call LayerShowHide(0)                                                
							Exit Function
						End If
					End If
					
                    strDel = strDel & "D" & iColSep & lRow & iColSep
                    strDel = strDel & Trim(GetSpreadText(.vspdData, C_LangCD, lRow, "X", "X")) & iColSep         
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

'=========================================================================================================
Function OpenLangCD()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "언어코드 팝업"
    arrParam(1) = "B_LANGUAGE"
    arrParam(2) = Trim(frm1.txtLangCd.Value)
    arrParam(3) = ""
    arrParam(4) = ""
    arrParam(5) = "언어 코드"
    
    arrField(0) = "LANG_CD"
    arrField(1) = "LANG_NM"
    
    arrHeader(0) = "언어코드"
    arrHeader(1) = "언어명"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
    If arrRet(0) = "" Then
    Else
        Call SetLangCD(arrRet)
    End If    
    frm1.txtLangCD.focus
    Set gActiveElement = document.activeElement
End Function
'=========================================================================================================
Function OpenLangInfo(Byval strCode)'khy200307

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "언어코드 팝업"
    arrParam(1) = "B_LANGUAGE"
    arrParam(2) = Trim(strCode)
    arrParam(3) = ""
    arrParam(4) = ""
    arrParam(5) = "언어 코드"
    
    arrField(0) = "LANG_CD"
    arrField(1) = "LANG_NM"
    
    arrHeader(0) = "언어코드"
    arrHeader(1) = "언어명"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
    
    If arrRet(0) = "" Then    
        Exit Function
    Else
        Call SetLangInfo(arrRet)
    End If    

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
    
    IntRetCD = CommonQueryRs("Lang_Nm","B_LANGUAGE","Lang_Cd =  " & FilterVar(frm1.txtLangCd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    lgF0 = Replace(lgF0, Chr(11), "") 'unusual case
    'lgF0 = Replace(lgF0, " ","")    

    If lgF0 = "" then 
        Call DisplayMsgBox("211432", "x", "x", "x")
        frm1.txtLangNm.value = ""        
        frm1.txtLangCd.select
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
Function SetLangCD(Byval arrRet)
    frm1.txtLangCD.Value    = Trim(arrRet(0))
    frm1.txtLangNm.value    = Trim(arrRet(1))
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
            	.SetText C_MnuID, nActiveRow, arrRet(0)
            	.SetText C_MnuNm, nActiveRow, arrRet(1)
                Call vspdData_Change(C_MnuNm, nActiveRow)
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
    Call PgmJump(JUMP_PGM_ID)
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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>컴퍼니 메뉴 관리</font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* align=right>
                                            <a href="vbscript:OpenRefMenu">메뉴별 Role Assign 현황</A>  
                    </TD>
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
                        <TD CLASS="TD5">언어코드</TD>
                        <TD CLASS="TD6"><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtLangCd" SIZE=10 MAXLENGTH=2 tag="12XXXU"  ALT="언어 코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLangCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenLangCD()">&nbsp;<INPUT TYPE=TEXT NAME="txtLangNm" SIZE=20 tag="14"></TD>                    
                        <TD CLASS="TD5">메뉴 ID</TD>
                        <TD CLASS="TD6"><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtMnuID" SIZE=15 MAXLENGTH=15 tag="11XXXU"  ALT="메뉴 ID"><IMG SRC="../../../CShared/image/btnPopup.gif"   NAME="btnMnuID" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMnuInfo frm1.txtMnuID.value,1 ">&nbsp;<INPUT TYPE=TEXT NAME="txtMnuNm" SIZE=40 tag="14"></TD>                        
                    </TR>
                    <TR>
                        <TD CLASS="TD5">메뉴타입</TD>
                        <TD CLASS="TD6"><SELECT NAME=cboMnuType tag="11X" STYLE="WIDTH: 82px;"><OPTION value=""></OPTION></SELECT></TD>
                        <TD CLASS="TD5">사용여부</TD>
                        <TD CLASS="TD6"><SELECT NAME=cboUseYN tag="11X" STYLE="WIDTH: 82px;"><OPTION value=""></OPTION></SELECT></TD>
                    </TR>                    
                </TABLE></FIELDSET></TD>
            </TR>
            <TR>
                <TD WIDTH=100% HEIGHT=* valign=top><TABLE WIDTH="100%" HEIGHT="100%">
                    <TR>
                        <TD HEIGHT="100%">
                        <script language =javascript src='./js/zc004ma2_I540804743_vspdData.js'></script></TD>
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
                        <TD WIDTH=* Align=right><A href="Vbscript:ProgramJump()">메뉴 Assignment to Role</A>&nbsp;</TD>                                                                                     
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
