<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : Dynamic Call
*  3. Program ID           : ZC006MA1
*  4. Program Name         : 
*  5. Program Desc         : Help Dynamic Call
*  6. Comproxy List        : 
*  7. Modified date(First) : 2005/03/04
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : Lee Sang Ho
* 10. Modifier (Last)      : 
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">        

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                    '

'=========================================================================================================
Const BIZ_PGM_ID = "zc006mb1.asp"
'=========================================================================================================
Dim C_LangCD    
Dim C_LangPopup'khy200307
Dim C_MnuID     
Dim C_MnuPopUp  
Dim C_MnudocNm     
Dim C_MnudocCallNm 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
'=========================================================================================================
Dim IsOpenPop
'=========================================================================================================
Sub initSpreadPosVariables()

 C_LangCD          = 1
 C_LangPopup       = 2
 C_MnuID           = 2
 C_MnuPopUp        = 4
 C_MnudocNm         = 5
 C_MnudocCallNm     = 6

End Sub

'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE

    lgIntGrpCount = 0
    
    lgStrPrevKey = ""
 
    lgLngCurRows = 0
    lgSortKey = 1    
    
End Sub
'=========================================================================================================
    
Sub SetDefaultVal()
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'=========================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 

    <% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>

    '------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'=========================================================================================================
Sub CookiePage(ByVal Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'=========================================================================================================
Sub MakeKeyStream(ByVal pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub
'=========================================================================================================
Sub InitSpreadComboBox()

    Dim IntRetCD
    
    
    'IntRetCD = CommonQueryRs("lang_cd","b_language","LANG_CD >= ''",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    'ggoSpread.Source = frm1.vspdData
    'lgF0 = Replace(lgF0, Chr(11), vbTab)    
    'lgF0 = Replace(lgF0, " ","")
    'ggoSpread.SetCombo lgF0, C_LangCD
    
    'IntRetCD = CommonQueryRs("MINOR_NM","B_MINOR"," MAJOR_CD = " & FilterVar("z0014", "''", "S") & " ORDER BY MINOR_CD",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    'ggoSpread.Source = frm1.vspddata
    'lgF0 = Replace(lgF0, Chr(11), vbTab)    
    'lgF0 = Replace(lgF0, " ","")
    'ggoSpread.SetCombo lgF0, C_MnudocType
End Sub  

'=========================================================================================================
Sub InitSpreadSheet()

    Call InitSpreadPosVariables()
    
    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)

        .ReDraw = false                   
        .MaxCols = C_MnudocCallNm +1

        ggoSpread.source=frm1.vspdData
        Call ggoSpread.ClearSpreadData()

        Call GetSpreadColumnPos("A")

        ggoSpread.SSSetEdit   C_LangCD          , "언어코드", 11,,,15, 2
        ggoSpread.SSSetButton C_LangPopUp'khy200307
        ggoSpread.SSSetEdit   C_MnuID           , "메뉴ID", 14,,,15, 2
        ggoSpread.SSSetButton C_MnuPopUp     
        ggoSpread.SSSetEdit   C_MnudocNm         , "문서명"    , 30,,,40
        ggoSpread.SSSetEdit   C_MnudocCallNm     , "호출문서ID", 21,,,21,2
        
        Call ggoSpread.MakePairsColumn(C_LangCD,C_LangPopup,"1")'khy200307     
        Call ggoSpread.MakePairsColumn(C_MnuID,C_MnuPopUp,"1")
        
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
        
        .ReDraw = true
  
        Call SetSpreadLock 
 
    End With
    
End Sub
'=========================================================================================================
Sub SetSpreadLock()                                                
    With frm1
    
        .vspdData.ReDraw = False

        ggoSpread.SpreadLock     C_LangCD, -1, C_LangPopup'khy200307
        ggoSpread.SpreadLock     C_MnuID, -1, C_MnuID
        ggoSpread.SpreadLock     C_MnuPopUp, -1, C_MnuPopUp
        ggoSpread.SSSetRequired     C_MnudocNm, -1, C_MnudocCallNm
        ggoSpread.SSSetProtected .vspdData.Maxcols, -1,-1
        
        .vspdData.ReDraw = True    

    End With
End Sub      
'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1
        .vspdData.ReDraw = False
        
        ggoSpread.SSSetRequired        C_LangCD, pvStartRow,pvEndRow
        ggoSpread.SSSetRequired        C_MnuID, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired        C_MnudocNm, pvStartRow, pvEndRow  
        ggoSpread.SSSetRequired        C_MnudocCallNm, pvStartRow, pvEndRow              
        
        .vspdData.ReDraw = True

    End With
  
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
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Action = 0 
              Exit For
           End If
           
       Next
          
    End If   
End Sub
'=========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    Select Case UCase(pvSpdNo)
    Case "A"
        ggoSpread.Source = frm1.vspdData
        Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_LangCD          = iCurColumnPos(1)
            C_LangPopup       = iCurColumnPos(2)
            C_MnuID           = iCurColumnPos(3)
            C_MnuPopUp        = iCurColumnPos(4)
            C_MnudocNm         = iCurColumnPos(5)
            C_MnudocCallNm     = iCurColumnPos(6)
            
    End Select
End Sub
'=========================================================================================================
Sub Form_Load()                                        

    Dim IntRetCD

    Call AppendNumberPlace("6","4","0")
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet
    Call InitVariables
    Call SetDefaultVal
    Call SetToolbar("11001101001111")

    
    frm1.txtLangCd.focus
    frm1.txtLangCd.Value = parent.gLang
    Set gActiveElement = document.activeElement
    
    IntRetCD = CommonQueryRs("Lang_Nm","B_LANGUAGE","Lang_Cd =  " & FilterVar(parent.gLang , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)        
    lgF0 = Replace(lgF0, Chr(11), "")    'unusual case    
    'lgF0 = Replace(lgF0," ","")            
    frm1.txtLangNm.value = Trim(lgF0)
    Call InitSpreadComboBox

    Call CookiePage (0) 

    Call FncQuery() 
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub
'=========================================================================================================
Function FncQuery()

    Dim IntRetCD 
   
    FncQuery = False
    
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
   
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
        IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")                 '☜: "Will you destory previous data"        
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If    
    
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.source=frm1.vspdData
    Call ggoSpread.ClearSpreadData()

    Call InitVariables
                                                                
    If Not chkField(Document, "1") Then
       Exit Function
    End If
  
    If DbQuery = False Then
       Exit Function
    End If
   
    If Err.number = 0 Then    
       FncQuery = True                                                                
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'=========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    On Error Resume Next                                                              
    Err.Clear                                                                     

    FncNew = False                                                                      
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    'In Multi, You need not to implement this area
    
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then    
       FncNew = True                                                                  
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'=========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    On Error Resume Next                                                              
    Err.Clear                                                                     

    FncDelete = False                                                                 
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    'In Multi, You need not to implement this area
    
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then    
       FncDelete = True                                                               
    End If

    Set gActiveElement = document.ActiveElement   

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
    
    If Err.number = 0 Then    
       FncNew = True                                                                  
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'=========================================================================================================
Function FncCopy()
    Dim IntRetCD, nActiveRow

    On Error Resume Next                                                              
    Err.Clear                                                                     

    FncCopy = False                                                                   

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
    With Frm1.VspdData
         .ReDraw = False
         If .ActiveRow > 0 Then
            ggoSpread.CopyRow
            nActiveRow = .ActiveRow
            SetSpreadColor nActiveRow, nActiveRow
            .SetText C_MnudocId, nActiveRow, ""
            .ReDraw = True
            .Focus
         End If
    End With
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    ' Clear key field
    '---------------------------------------------------------------------------------------------------- 
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then    
       FncCopy = True                                                                
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'=========================================================================================================
Function FncCancel() 
    Dim iDx

    On Error Resume Next                                                              
    Err.Clear                                                                     

    FncCancel = False                                                                 

    ggoSpread.Source = Frm1.vspdData    
    ggoSpread.EditUndo  
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then    
       FncCancel = True                                                                
    End If

    Set gActiveElement = document.ActiveElement   

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

    On Error Resume Next                                                              
    Err.Clear                                                                     

    FncDeleteRow = False                                                              

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
    End if    
    
    With Frm1.vspdData 
        .focus
        ggoSpread.Source = frm1.vspdData 
        lDelRows = ggoSpread.DeleteRow
    End With
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then    
       FncDeleteRow = True                                                                
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'=========================================================================================================
Function FncPrint()

    On Error Resume Next                                                              
    Err.Clear                                                                     

    FncPrint = False                                                                  
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call Parent.FncPrint()                                                        

    If Err.number = 0 Then     
       FncPrint = True                                                                
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'=========================================================================================================
Function FncPrev() 

    On Error Resume Next                                                              
    Err.Clear                                                                     

    FncPrev = False                                                                   
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then     
       FncPrev = True                                                                 
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'=========================================================================================================
Function FncNext() 

    On Error Resume Next                                                              
    Err.Clear                                                                     

    FncNext = False                                                                   
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then     
       FncNext = True                                                                 
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'=========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                              
    Err.Clear                                                                     

    FncExcel = False                                                                  

    Call Parent.FncExport(Parent.C_MULTI)

    If Err.number = 0 Then     
       FncExcel = True                                                                
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'=========================================================================================================
Function FncFind() 

    On Error Resume Next                                                              
    Err.Clear                                                                     

    FncFind = False                                                                   

    Call Parent.FncFind(Parent.C_MULTI, True)

    If Err.number = 0 Then     
       FncFind = True                                                                 
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'=========================================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
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
    Call InitSpreadSheet      
    Call InitSpreadComboBox
    Call ggoSpread.ReOrderingSpreadData()
End Sub
'=========================================================================================================
Function FncExit()
    Dim IntRetCD

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncExit = False                                                                   
    
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")                      '⊙: Data is changed.  Do you want to exit? 
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    If Err.number = 0 Then     
       FncExit = True                                                                 
    End If

    Set gActiveElement = document.ActiveElement   

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
    'Call InitSpreadComboBox
    Err.Clear

    Dim strVal
    
    
    
    With frm1

    
    If lgIntFlgMode = parent.OPMD_UMODE Then  

        strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
        strVal = strVal & "&txtLangCd=" & Trim(.hLangCd.value)
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows    
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

    Else    
        strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
        strVal = strVal & "&txtLangCd=" & Trim(.txtLangCd.value)
        strVal = strVal & "&txtMnuId=" & Trim(.txtMnuId.value)
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

    End If       

    Call RunMyBizASP(MyBizASP, strVal)
    
    End With
    
    DbQuery = True

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
    Dim iMenuDocCalNm 
    Dim IMenuCheck
    Dim iCount
    Dim iMenuId
    Dim lgF0
    Dim iPos
    Dim iExt
    iColSep = parent.gColSep
    iRowSep = parent.gRowSep
        
    DbSave = False
    
    Call LayerShowHide(1)    

    On Error Resume Next

    With frm1

        .txtMode.value = parent.UID_M0002
        .txtUpdtUserId.value = parent.gUsrID
        .txtInsrtUserId.value = parent.gUsrID
        
        lGrpCnt = 1
    
        strVal = ""
        strDel = ""
       
        For lRow = 1 To .vspdData.MaxRows
			iMenuId = Trim(GetSpreadText(.vspdData, C_MnuID, lRow, "X", "X"))

			Call CommonQueryRs("MNU_ID "," Z_LANG_CO_MAST_MNU "," LANG_CD = 'KO' AND MNU_TYPE <> 'M' And MNU_ID = " & FilterVar(iMenuId, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			If Trim(Replace(lgF0,Chr(11),"")) = "" Then
				IntRetCD = DisplayMsgBox("800480", "x", "x", "x")
				Call LayerShowHide(0)
				Exit Function
			End If
			
            iMenuDocCalNm = Trim(GetSpreadText(.vspdData, C_MnudocCallNm     , lRow, "X", "X"))
			iCount = Len(iMenuDocCalNm)
			If iCount > 21 Then
				IntRetCD = DisplayMsgBox("215516", "x", "x", "x")
				Call LayerShowHide(0)
				Exit Function
			End If

			iPos = instr(1, iMenuDocCalNm, ".", 1)
			If iPos > 18 Then
				IntRetCD = DisplayMsgBox("215514", "x", "x", "x")
				Call LayerShowHide(0)
				Exit Function
			End If
			
			iExt = mid(iMenuDocCalNm, iPos + 1, 5)
			If Len(Trim(iExt)) <> 3 Then
				IntRetCD = DisplayMsgBox("215515", "x", "x", "x")
				Call LayerShowHide(0)
				Exit Function
			End If
			
			iPos = instr(1, iMenuDocCalNm, "_", 1)
			If iPos = 0 Then
				IntRetCD = DisplayMsgBox("215516", "x", "x", "x")
				Call LayerShowHide(0)
				Exit Function
			End If
			
            Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")

                Case ggoSpread.InsertFlag
				                   
                                                    strVal = strVal & "C"                        & iColSep
                                                    strVal = strVal & lRow                    & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_LangCD          , lRow, "X", "X")) & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuID           , lRow, "X", "X")) & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnudocNm         , lRow, "X", "X")) & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnudocCallNm     , lRow, "X", "X")) & iRowSep
                    lGrpCnt = lGrpCnt + 1

                Case ggoSpread.UpdateFlag

                                                    strVal = strVal & "U"                        & iColSep
                                                    strVal = strVal& lRow                        & iColSep         
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_LangCD          , lRow, "X", "X")) & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuID           , lRow, "X", "X")) & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnudocNm         , lRow, "X", "X")) & iColSep
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnudocCallNm     , lRow, "X", "X")) & iRowSep
                    lGrpCnt = lGrpCnt + 1

                Case ggoSpread.DeleteFlag

                                                    strDel = strDel & "D"                        & iColSep
                                                    strDel = strDel& lRow                        & iColSep
								                    strDel = strDel & Trim(GetSpreadText(.vspdData, C_LangCD          , lRow, "X", "X")) & iColSep
								                    strDel = strDel & Trim(GetSpreadText(.vspdData, C_MnuID           , lRow, "X", "X")) & iColSep               
								                    strDel = strDel & Trim(GetSpreadText(.vspdData, C_MnudocNm         , lRow, "X", "X")) & iColSep
								                    strDel = strDel & Trim(GetSpreadText(.vspdData, C_MnudocCallNm     , lRow, "X", "X")) & iRowSep
                                       
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
Function DbDelete()

    On Error Resume Next                                                              
    Err.Clear                                                                     

    DbDelete = False                                                                  
    '------ Developer Coding part (Start)  -------------------------------------------------------------- 
    'In Multi, You need not to implement this area

    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then     
       DbDelete = True                                                                 
    End If

    Set gActiveElement = document.ActiveElement   

End Function
'=========================================================================================================
Function DbQueryOk()
    
    lgIntFlgMode = parent.OPMD_UMODE
    
    Call ggoOper.LockField(Document, "Q")
    Call SetToolbar("11001111001111")

End Function
'=========================================================================================================
Function DbSaveOk()
   
    Call InitVariables
    Call ggoOper.ClearField(Document, "2")                                             
    ggoSpread.source=frm1.vspdData
    Call ggoSpread.ClearSpreadData()
    
    Call MainQuery()

End Function
'=========================================================================================================
Function DbDelete() 
End Function


'=========================================================================================================
Sub DbDeleteOk()

    On Error Resume Next                                                              
    Err.Clear                                                                     

    '------ Developer Coding part (Start)  -------------------------------------------------------------- 

    '------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
End Sub

' Name : OpenLangCD
'========================================================================================================
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
    
    frm1.txtLangCd.focus
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

'========================================================================================================
' Name : OpenMnuInfo
'========================================================================================================
Function OpenMnuInfo(Byval strCode, Byval iWhere)         

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)
    Dim strVal
    
    If IsOpenPop = True Then Exit Function
    
    IsOpenPop = True

    arrParam(0) = "메뉴 팝업"
    arrParam(1) = "Z_LANG_CO_MAST_MNU"    
    arrParam(2) = strCode
    arrParam(3) = ""
        
    arrParam(4) = "LANG_CD = " & FilterVar(Parent.gLang, "''", "S") & " AND MNU_TYPE <> " & FilterVar("M", "''", "S") & " "
                
    arrParam(5) = "메뉴ID"
    
    arrField(0) = "MNU_ID"
    arrField(1) = "MNU_NM"
    
    arrHeader(0) = "메뉴ID"
    arrHeader(1) = "메뉴명"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetMnuInfo(arrRet, iWhere)
    End If    
	If iWhere = 1 Then
	    frm1.txtMnuId.focus
	    Set gActiveElement = document.activeElement
	End If
    
End Function
'========================================================================================================
' Name : SetLangCD
'========================================================================================================
Function SetLangCD(Byval arrRet)                       
    frm1.txtLangCD.Value    = Trim(arrRet(0))
    frm1.txtLangNm.value    = Trim(arrRet(1))
End Function

'========================================================================================================
' Name : SetMnuInfo
'========================================================================================================
Function SetMnuInfo(Byval arrRet, Byval iWhere)            
	Dim nActiveRow
    Select Case iWhere
    Case  1
        frm1.txtMnuId.Value  = arrRet(0)
        frm1.txtMnuNm.Value    = arrRet(1)
    Case  2
        With frm1.vspdData
        	nActiveRow = .ActiveRow
        	.SetText C_MnuID, nActiveRow, arrRet(0)
        	.SetText C_MnudocNm, nActiveRow, arrRet(1)
            Call vspdData_Change(C_MnudocNm, nActiveRow)
        End With
    End Select

End Function

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
				Case C_MnuPopUp
					Call OpenMnuInfo(GetSpreadText(frm1.vspdData, C_MnuID, Row, "X", "X"), 2)   
			End Select
        End If
    End With
End Sub
'=========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
End Sub
'=========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")     
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    
       Exit Sub
       End If

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col,lgSortKey
            lgSortKey = 1
        End If    
 
        Exit Sub
    End If
End Sub

'=========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)        
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'=========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)                
    If Row <= 0 Then
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
    
End Sub
'=========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'=========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
    End If

End Sub

'=========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
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
        Call DisableToolBar(parent.TBC_QUERY)
        If DBQuery = False Then
            Call RestoreToolBar()
            Exit Sub
        End If 
    End if
    
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc" -->    
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
    <TR>
        <TD <%=HEIGHT_TYPE_00%>></TD>
    </TR>
    <TR HEIGHT=23>
        <TD WIDTH=100%>
            <TABLE <%=LR_SPACE_TYPE_10%> BORDER=0>
                <TR>
                    <TD WIDTH=10>&nbsp;</TD>
                    <TD CLASS="CLSMTABP">
                        <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
                            <TR>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>On-Line Help 관리</font></td>
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
                        <TD CLASS="TD5">언어코드</TD>
                        <TD CLASS="TD6"><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtLangCd" SIZE=10 MAXLENGTH=2 tag="12XXXU"  ALT="언어 코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLangCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenLangCD()">&nbsp;<INPUT TYPE=TEXT NAME="txtLangNm" SIZE=20 tag="14"></TD>                    
                        <TD CLASS="TD5">메뉴ID</TD>
                        <TD CLASS="TD6"><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtMnuId" SIZE=15 MAXLENGTH=15 tag="11XXXU"  ALT="메뉴 ID"><IMG SRC="../../../CShared/image/btnPopup.gif"   NAME="btnMnuID" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMnuInfo frm1.txtMnuId.value,1 ">&nbsp;<INPUT TYPE=TEXT NAME="txtMnuNm" SIZE=40 tag="14"></TD>                        
                    </TR>
                    
                </TABLE></FIELDSET></TD>
            </TR>
            <TR>
                <TD WIDTH=100% HEIGHT=* valign=top><TABLE WIDTH="100%" HEIGHT="100%">
                    <TR>
                        <TD HEIGHT="100%">
                        <script language =javascript src='./js/zc006ma1_I324619233_vspdData.js'></script></TD>
                    </TR></TABLE>
                </TD>
            </TR>
        </TABLE></TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMdocORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
    <INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
    <INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
    <INPUT TYPE=HIDDEN NAME="hLangCd" tag="24">
    <INPUT TYPE=HIDDEN NAME="hMnuId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMdocORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
