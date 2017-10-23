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

Const BIZ_PGM_ID = "zc00kmb1.asp"

Dim C_PrcsSeq
Dim C_PrcsSubSeq
Dim C_MnuID
Dim C_MnuPopUp
Dim C_MnuNm
Dim C_OptnFlag
Dim C_Remark

<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim lgStrPrevKey2
dim    lgStrQueryFlag

Dim IsOpenPop


'=========================================================================================================
Sub InitSpreadPosVariables()
    C_PrcsSeq = 1
    C_PrcsSubSeq = 2
    C_MnuID = 3
    C_MnuPopUp = 4
    C_MnuNm = 5
    C_OptnFlag = 6
    C_Remark = 7
End Sub
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
    lgStrPrevKey = ""
    lgStrPrevKey2 = ""

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
        .MaxCols = C_Remark + 1
        .MaxRows = 0

         Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetEdit    C_PrcsSeq, "순서",7,,,3,2        
        ggoSpread.SSSetEdit    C_PrcsSubSeq, "세부순서",10,,,2,2        
        ggoSpread.SSSetEdit C_MnuID, "메뉴ID", 15,,,15,2
        ggoSpread.SSSetButton C_MnuPopUp
        ggoSpread.SSSetEdit C_MnuNm, "메뉴명", 30,,,40
        ggoSpread.SSSetCheck C_OptnFlag, "옵션여부", 9,,,True
        ggoSpread.SSSetEdit    C_Remark, "비고", 50,,,200
        .ReDraw = true

        Call SetSpreadLock
        
        Call ggoSpread.MakePairsColumn(C_MnuID,C_MnuPopUp,"1")

        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_PrcsSeq, -1, C_PrcsSeq
        ggoSpread.SpreadLock C_PrcsSubSeq, -1, C_PrcsSubSeq
        ggoSpread.SpreadLock C_MnuID, -1, C_MnuID
        ggoSpread.SpreadLock C_MnuPopUp, -1, C_MnuPopUp
        ggoSpread.SpreadLock C_MnuNm, -1, C_MnuNm                                
        .vspdData.ReDraw = True
    End With
End Sub
'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired    C_PrcsSeq, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_PrcsSubSeq, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired    C_MnuID, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_MnuNm, pvStartRow, pvEndRow
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

            C_PrcsSeq    = iCurColumnPos(1)
            C_PrcsSubSeq = iCurColumnPos(2)
            C_MnuID      = iCurColumnPos(3)
            C_MnuPopUp   = iCurColumnPos(4)
            C_MnuNm      = iCurColumnPos(5)
            C_OptnFlag   = iCurColumnPos(6)
            C_Remark     = iCurColumnPos(7)

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
    Call ggoOper.LockField(Document, "Q")'khy200303
    
    Call InitSpreadSheet

    Call InitVariables
    Call SetDefaultVal

    'Call SetToolbar("1110000000111111")
    Call SetToolbar("1110110100111111")
    frm1.txtPrcsCD.focus
    Set gActiveElement = document.activeElement
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub
'=========================================================================================================
Sub cbxSysFlag_OnClick()
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
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
   
    With frm1.vspdData 
	    ggoSpread.Source = frm1.vspdData
	   	.Row = Row
	    If Row > 0 And Col = C_MnuPopUp Then
	        Call OpenMnuInfo(GetSpreadText(frm1.vspdData, C_MnuID, Row, "X", "X"))
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
    
    lgStrQueryFlag = ""
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) _
    And Not(lgStrPrevKey = "" And lgStrPrevKey2 = "") Then
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
    Call ggoSpread.ReOrderingSpreadData()
    Call InitData()
End Sub
'=========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False                                                        
    Call SetToolbar("1110000000111111")    
    Err.Clear

    lgStrQueryFlag = "Q"
    
    frm1.txtPrcsNm.value=""
    frm1.txtUppMnuNm.value="" 'khy khy2003.03
    
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
              Exit Function
        End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
    Call ggoSpread.ClearSpreadData()            
    Call InitVariables

    If Not chkField(Document, "1") Then
       Call SetToolbar("1110000000111111")    'khy200304
       Call ggoOper.LockField(Document, "Q")    'khy200304        
       Exit Function
    End If
    
    If DbQuery = False Then
       Exit Function
    End If
       
    FncQuery = True                                                            

End Function
'=========================================================================================================
Function FncNew()

    Dim IntRetCD

    FncNew = False

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    
    
    frm1.vspdData.MaxRows = 0
    Call ggoOper.ClearField(Document, "A")
    Call ggoSpread.ClearSpreadData        
    Call ggoOper.LockField(Document, "N")        
    Call InitVariables
    Call SetDefaultVal
    frm1.txtPrcsCD2.focus 'khy200303
    Call SetToolbar("1110110100111111")
            
    FncNew = True                                                           
    
End Function
'=========================================================================================================
Function FncDelete()

    Dim IntRetCD

    FncDelete = False

    IntRetCD = DisplayMsgBox("210034", Parent.VB_YES_NO, "x", "x")
    If IntRetCD = vbNo Then
        Exit Function
    End If

    If DbDelete = False Then 
       Exit Function
    End If
    
    FncDelete = True

End Function
'=========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False

    Err.Clear                                                               
    On Error Resume Next
    
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData    
    if Not ggoSpread.SSDefaultCheck then
        Exit Function
    End if
    
    If frm1.vspdData.MaxRows < 1 Then
        IntRetCD = DisplayMsgBox("211431", "x", "x", "x")
        Exit Function
    Else
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
            nActiveRow = .ActiveRow
            SetSpreadColor nActiveRow, nActiveRow
    		.SetText C_PrcsSeq, nActiveRow, ""
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
    Dim IntRetCD 

 
    FncPrev = False                                                        
    
    Err.Clear                                                               
    
    lgStrQueryFlag = "P"
                                                               

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
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

    FncNext = False                                                        
    
    Err.Clear                                                               

    lgStrQueryFlag = "N"


    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
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
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLEMULTI)
End Function
'=========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLEMULTI , True)                               
End Function
'=========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = True Then
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
Function DbDelete() 
    
    Err.Clear

    DbDelete = False

    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003
    strVal = strVal & "&txtPrcsCd2=" & Trim(frm1.txtPrcsCd2.value)
    strVal = strVal & "&txtPrcsNm2=" & Trim(frm1.txtPrcsNm2.value)
    strVal = strVal & "&txtUppMnu=" & Trim(frm1.txtUppMnu.value) 'khy khy2003.03
    
    Call RunMyBizASP(MyBizASP, strVal)
    
    DbDelete = True
        
    Call SetToolbar("1110000000111111")'khy200303
        
End Function
'=========================================================================================================
Function DbDeleteOk()
    Call FncNew()
End Function

'=========================================================================================================
Function Clear()
    Call ggoOper.ClearField(Document, "2")    
    Call ggoSpread.ClearSpreadData        
End Function
'=========================================================================================================
Function DbQuery() 
    
    DbQuery = False
    
    Call LayerShowHide(1)
        
    Dim strVal
    
    With frm1
    
    Call SetToolbar("1110000000111111")    'khy200304
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
    
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
        strVal = strVal & "&txtPrcsCd=" & Trim(.txtPrcsCd.value)
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows    
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2 
        strVal = strVal & "&txtPrvNext=" & lgStrQueryFlag 
    Else    
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
        strVal = strVal & "&txtPrcsCd=" & Trim(.txtPrcsCd.value)
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows    
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2 
        strVal = strVal & "&txtPrvNext=" & lgStrQueryFlag        

    End If 
    
    End With
    
    Call RunMyBizASP(MyBizASP, strVal)
    
    DbQuery = True     
    
    
End Function
'=========================================================================================================
Function DbQueryOk()
    
    lgIntFlgMode = Parent.OPMD_UMODE
    lgBlnFlgChgValue = False    

    Call ggoOper.LockField(Document, "Q")
    
    Call SetToolbar("1111111111111111")'1111111100111111
    

End Function
'=========================================================================================================
Function DbSave()

    Dim IntRows 
    Dim IntCols 
    
    Dim lGrpcnt 
    Dim strVal
    Dim strDel
    Dim StrOpt'khy20030111    
    Dim iColSep, iRowSep
    
    iColSep = parent.gColSep
    iRowSep = parent.gRowSep
    
    DbSave = False

    Call LayerShowHide(1)
    
    On Error Resume Next
    
    With frm1
        if lgIntFlgMode = Parent.OPMD_CMODE then
            .txtMode.value = Parent.UID_M0002
        else
            .txtMode.value = Parent.UID_M0005
        end if
        .txtFlgMode.value = lgIntFlgMode
        .txtUpdtUserId.value = Parent.gUsrID
        .txtInsrtUserId.value  = Parent.gUsrID
    End With
    
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
        If frm1.cbxSysFlag.checked Then
            frm1.hSysFlag.value = "1"
        Else
            frm1.hSysFlag.value = "0"
        End If
        
    With frm1.vspdData
        
        For IntRows = 1 To .MaxRows
        
            .Row = IntRows
            .Col = 0
        	
        	Select Case GetSpreadText(frm1.vspdData, 0, IntRows, "X", "X")
            Case ggoSpread.InsertFlag
                strVal = strVal & "C" & iColSep & IntRows & iColSep
            Case ggoSpread.UpdateFlag
                strVal = strVal & "U" & iColSep & IntRows & iColSep
            Case ggoSpread.DeleteFlag
                strDel = strDel & "D" & iColSep & IntRows & iColSep
            End Select
        
            Select Case GetSpreadText(frm1.vspdData, 0, IntRows, "X", "X")
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
                    strVal = strVal & Trim(GetSpreadText(frm1.vspdData, C_PrcsSeq, IntRows, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(frm1.vspdData, C_PrcsSubSeq, IntRows, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(frm1.vspdData, C_MnuID, IntRows, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(frm1.vspdData, C_OptnFlag, IntRows, "X", "X")) & iColSep
                    strVal = strVal & Trim(GetSpreadText(frm1.vspdData, C_Remark, IntRows, "X", "X")) & iRowSep

                    'validate check----------------------------------------------------------------------
        
                    If CheckNumeric(Trim(GetSpreadText(frm1.vspdData, C_PrcsSeq, IntRows, "X", "X")))=1 then              
                        IntRetCD = DisplayMsgBox("211434", "x","x","x")                'Menu sequence only decimal
                        Call LayerShowHide(0)                        
                        Exit Function        
                    End If
    
                    If CheckNumeric(Trim(GetSpreadText(frm1.vspdData, C_PrcsSubSeq, IntRows, "X", "X")))=1 then              
                        IntRetCD = DisplayMsgBox("211435", "x","x","x")                'Menu sequence only decimal
                        Call LayerShowHide(0)                        
                        Exit Function        
                    End If  
                    '---------------------------------------------------------------------------------------

                    lGrpCnt = lGrpCnt + 1
    
                Case ggoSpread.DeleteFlag
                    strDel = strDel & Trim(GetSpreadText(frm1.vspdData, C_PrcsSeq, IntRows, "X", "X")) & iColSep
                    strDel = strDel & Trim(GetSpreadText(frm1.vspdData, C_PrcsSubSeq, IntRows, "X", "X")) & iRowSep

                    lGrpcnt = lGrpcnt + 1

            End Select
        Next
    End With
    
    frm1.txtMaxRows.value = lGrpCnt-1
    frm1.txtSpread.value = strDel & strVal

    Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                           
    
End Function

'=========================================================================================================
Function DbSaveOk()

    Call InitVariables
    frm1.vspdData.MaxRows = 0
    
    lgIntFlgMode = Parent.OPMD_UMODE
    
    Call ggoOper.LockField(Document, "Q")
    
    lgBlnFlgChgValue = False
    
    Call MainQuery()

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
Function OpenPrcsCd()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "프로세스코드 팝업"
    arrParam(1) = "Z_PRCS_MAST"
    arrParam(2) = Trim(frm1.txtPrcsCd.Value)
    arrParam(3) = ""
    arrParam(4) = ""
    arrParam(5) = "프로세스 코드"
    
    arrField(0) = "PRCS_CD"    
    arrField(1) = "PRCS_NM"    

    arrHeader(0) = "프로세스 코드"
    arrHeader(1) = "프로세스 명"

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetPrcsCD(arrRet)
    End If    
    
    frm1.txtPrcsCd.focus     
    Set gActiveElement = document.activeElement
    
End Function
'=========================================================================================================
'khy2003.03
Function OpenUppMnu()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    'khy200304
    If IsOpenPop = True Or UCase(frm1.txtUppMnu.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "상위메뉴 팝업"
    arrParam(1) = "Z_LANG_CO_MAST_MNU"
    arrParam(2) = Trim(frm1.txtUppMnu.Value)
    arrParam(3) = ""
    arrParam(4) = "LANG_CD = " & FilterVar(gLang, "''", "S") & " AND MNU_TYPE = " & FilterVar("M", "''", "S") & "  AND MNU_ID LIKE " & FilterVar("R%", "''", "S") & ""    
    arrParam(5) = "상위메뉴 ID"
    
    arrField(0) = "MNU_ID"    
    arrField(1) = "MNU_NM"    

    arrHeader(0) = "상위메뉴 ID"
    arrHeader(1) = "상위메뉴 명"

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetUppMnu(arrRet)
    End If    
    
    frm1.txtUppMnu.focus     
    Set gActiveElement = document.activeElement
End Function
'=========================================================================================================
Function OpenMnuInfo(Byval strCode)

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "메뉴 팝업"
    arrParam(1) = "Z_LANG_CO_MAST_MNU"
    arrParam(2) = Trim(strCode)
    arrParam(3) = ""
    arrParam(4) = "LANG_CD = " & FilterVar(Parent.gLang, "''", "S") & " AND MNU_TYPE = " & FilterVar("P", "''", "S") & " "
    arrParam(5) = "메뉴ID"
    
    arrField(0) = "MNU_ID"
    arrField(1) = "MNU_NM"
    
    arrHeader(0) = "메뉴ID"
    arrHeader(1) = "메뉴명"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetMnuInfo(arrRet)
    End If    
    

End Function
'=========================================================================================================
Function SetPrcsCD(byval arrRet)
    frm1.txtPrcsCd.Value = arrRet(0)
    frm1.txtPrcsNm.Value = arrRet(1)
End Function
'=========================================================================================================
'khy2003.03
Function SetUppMnu(byval arrRet)
    frm1.txtUppMnu.Value = arrRet(0)    
    frm1.txtUppMnuNM.Value = arrRet(1)
End Function

'=========================================================================================================
Function SetMnuInfo(Byval arrRet)
	Dim nActiveRow
    With frm1
		nActiveRow = .vspdData.ActiveRow
		.vspdData.SetText C_MnuID, nActiveRow, arrRet(0)
		.vspdData.SetText C_MnuNm, nActiveRow, arrRet(1)
	    Call vspdData_Change(C_MnuNm, nActiveRow)
    End With

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BasicTB" CELLSPACING=0>
    <TR>
        <TD <%=HEIGHT_TYPE_00%>></TD>
    </TR>
    <TR HEIGHT=23>
        <TD WIDTH=100%>
            <TABLE CLASS="BasicTB" CELLSPACING=0>
                <TR>
                    <TD WIDTH=10>&nbsp;</TD>
                    <TD CLASS="CLSMTABP">
                        <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
                            <TR>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>프로세스별 메뉴관리</font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=500>&nbsp;</TD>
                    <TD WIDTH=*>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR HEIGHT=*>
        <TD WIDTH=100% CLASS="Tab11">
                <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
                    <TR>
                        <TD HEIGHT=5 WIDTH=100%></TD>
                    </TR>
                    <TR>
                        <TD HEIGHT=20 WIDTH=100%>
                            <FIELDSET CLASS="CLSFLD">
                                <TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                                    <TR>
                                        <TD CLASS="TD5" NOWRAP>프로세스 코드</TD>
                                        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPrcsCd" SIZE=10 MAXLENGTH=10 tag="12XXXU"  ALT="프로세스 코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrcsCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPrcsCD()">&nbsp;<INPUT TYPE=TEXT NAME="txtPrcsNm" SIZE=40 tag="14"></TD>
                                        <TD CLASS="TDT"></TD>
                                        <TD CLASS="TD6"></TD>
                                    </TR>
                                </TABLE>
                            </FIELDSET>
                        </TD>
                    </TR>
                    <TR>
                        <TD HEIGHT=5 WIDTH=100%></TD>
                    </TR>
                    <TR>
                        <TD HEIGHT=20 WIDTH=100%  valign=top>
                            <TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                                <TR>
                                    <TD CLASS="TD5" NOWRAP>프로세스 코드</TD>
                                        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPrcsCD2" SIZE=10 MAXLENGTH=10 tag="22X"  ALT="프로세스 코드">&nbsp;</TD>
                                    <TD CLASS="TD5" NOWRAP>시스템 프로세스 여부</TD>
                                    <TD CLASS="TD6" NOWRAP><INPUT TYPE=CHECKBOX NAME="cbxSysFlag" CLASS="STYLE CHECK" tag="22" ALT="시스템 프로세스 여부"></TD>                                        
                                </TR>
                                <TR><!--khy2003.03-->
                                    <TD CLASS="TD5" NOWRAP>코드명</TD>
                                        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPrcsNm2" MAXLENGTH=40 SIZE=40 tag="22X" ALT="코드명"></TD>                                
                                    <TD CLASS="TD5" NOWRAP>상위 메뉴</TD>                                    
                                        <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtUppMnu" SIZE=10 MAXLENGTH=10 tag="22X"  ALT="상위 메뉴" ><IMG SRC="../../image/btnPopup.gif" NAME="btnUppMnu" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUppMnu()" >&nbsp;<INPUT TYPE=TEXT NAME="txtUppMnuNm" SIZE=25 tag="14"></TD>
                                </TR>
                            </TABLE>
                        </TD>
                    </TR>
                    <TR>
                        <TD WIDTH=100% HEIGHT=* valign=top><TABLE WIDTH="100%" HEIGHT="100%">
                            <TR>
                                <TD HEIGHT="100%">
                                    <script language =javascript src='./js/zc00kma1_I847206836_vspdData.js'></script>
                                </TD>
                            </TR></TABLE>
                        </TD>
                    </TR>                    
                </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../BLANK.HTM" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPrcsCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hSysFlag" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
