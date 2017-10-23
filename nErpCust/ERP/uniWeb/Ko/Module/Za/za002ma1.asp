<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BASIS ARCHITECT
*  2. Function Name        : 
*  3. Program ID           : ZA002MA1.ASP
*  4. Program Name         : LOGON GROUP ENTRY
*  5. Program Desc         : LOGON GROUP MANAGEMENT
*  6. Comproxy List        : PZAG006.DLL, PZAG007.DLL
*  7. Modified date(First) : 2001/09/01
*  8. Modified date(Last)  : 2001/12/01
*  9. Modifier (First)     : PARK, SANGHOON
* 10. Modifier (Last)      : PARK, SANGHOON
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

Const BIZ_PGM_ID = "Za002mb1.asp"            

Dim C_LogonGp
Dim C_LogonGpNm
Dim C_AppServerId
Dim C_Port
Dim C_DB_Server_Id
Dim C_DB_Nm
Dim C_DsnNo
Dim C_SuplUsrCnt

<!-- #Include file="../../inc/lgvariables.inc" -->    


Dim IsOpenPop          

'=========================================================================================================
Sub InitSpreadPosVariables()
    C_LogonGp = 1
    C_LogonGpNm = 2
    C_AppServerId = 3
    C_Port = 4
    C_DB_Server_Id = 5
    C_DB_Nm = 6
    C_DsnNo = 7
    C_SuplUsrCnt = 8
End Sub


'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                               
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
        .MaxCols = C_SuplUsrCnt + 1                        
        .MaxRows = 0

        Call GetSpreadColumnPos("A")        
        
        'Column position 재지정후 할당해야 함 
        .SetText C_Port, 0, "0"
        .SetText C_DsnNo, 0, "1"

        ggoSpread.SSSetEdit C_LogonGp, "로그온그룹 코드", 18, , , 30, 2               '1
        ggoSpread.SSSetEdit C_LogonGpNm, "로그온 그룹", 25, , , 30                    '2
        ggoSpread.SSSetEdit C_AppServerId, "App. 서버", 15, , , 20                         '3
        ggoSpread.SSSetEdit C_Port, "Port 번호", 10, 2, , 4                           '4
        ggoSpread.SSSetEdit C_DB_Server_Id, "DB 서버", 15, , , 15                   '5
        ggoSpread.SSSetEdit C_DB_Nm, "DB명", 15, , , 25                               '6
        ggoSpread.SSSetEdit C_DsnNo, "DSN 번호", 15, 2, , 1                           '7        
        ggoSpread.SSSetEdit C_SuplUsrCnt, "사용자 수", 20, 1                         '8
        .ReDraw = true
    
        Call SetSpreadLock 
        
        Call ggoSpread.MakePairsColumn(C_LogonGp,C_LogonGpNm,"1")

        Call ggoSpread.SSSetColHidden(C_Port,C_Port,True)
        Call ggoSpread.SSSetColHidden(C_DsnNo,C_DsnNo,True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

    
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_LogonGp, -1, C_LogonGp
        ggoSpread.SSSetRequired    C_LogonGpNm, -1, C_LogonGpNm
        ggoSpread.SSSetRequired    C_AppServerId, -1, C_AppServerId
        ggoSpread.SSSetRequired    C_DB_Server_Id, -1, C_DB_Server_Id
        ggoSpread.SSSetRequired    C_DB_Nm, -1, C_DB_Nm
        ggoSpread.SpreadLock C_SuplUsrCnt, -1, C_SuplUsrCnt
        .vspdData.ReDraw = True
    End With
End Sub


'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired C_LogonGp, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_LogonGpNm, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_AppServerId, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_DB_Server_Id, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_DB_Nm, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_SuplUsrCnt, pvStartRow, pvEndRow
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

            C_LogonGp      =  iCurColumnPos(1)
            C_LogonGpNm    =  iCurColumnPos(2)
            C_AppServerId  =  iCurColumnPos(3)
            C_Port         =  iCurColumnPos(4)
            C_DB_Server_Id =  iCurColumnPos(5)
            C_DB_Nm        =  iCurColumnPos(6)
            C_DsnNo        =  iCurColumnPos(7)
            C_SuplUsrCnt   =  iCurColumnPos(8)
            
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
Sub Form_Load()

    Call LoadInfTB19029                                                              
    Call ggoOper.LockField(Document, "N")                                             

    Call InitSpreadSheet  

    Call InitVariables                                                      
    Call SetDefaultVal   

    frm1.txtLogonGp.focus         
    Set gActiveElement = document.activeElement
    Call SetToolbar("11001101001011")                                        

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
Sub vspdData_Change(ByVal Col, ByVal Row)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
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

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)  And lgStrPrevKey <> "" Then
           Call DisableToolBar(parent.TBC_QUERY)                                                       
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
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x", "x")                
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
    With frm1.vspdData    	
        If .ActiveRow > 0 Then
	    	
            .focus
            .ReDraw = False
            
            ggoSpread.Source = frm1.vspdData 
            ggoSpread.CopyRow
            nActiveRow = .ActiveRow
            SetSpreadColor nActiveRow, nActiveRow
            
            .SetText C_LogonGp, nActiveRow, ""
			.SetText C_SuplUsrCnt, nActiveRow, ""
						
            .ReDraw = True
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
    'On Error Resum Next                                                    
End Function

'=========================================================================================================
Function FncNext() 
    'On Error Resum Next                                                    
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
        
    'On Error Resum Next    
    Err.Clear                                                               

    DbQuery = False

    Call DisableToolBar(parent.TBC_QUERY)                                                
    Call LayerShowHide(1)                                                         
    
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
        strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001                            
        strVal = strVal & "&txtLogonGp=" & .hLogonGp.value             
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    Else
        strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001                            
        strVal = strVal & "&txtLogonGp=" & Trim(.txtLogonGp.value)            
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    End If    

    Call LayerShowHide(1)
    Call RunMyBizASP(MyBizASP, strVal)                                        
        
    End With
    
    DbQuery = True
    
End Function

'=========================================================================================================
Function DbQueryOk()                                                        
    
    lgIntFlgMode = parent.OPMD_UMODE                                                
    
    Call ggoOper.LockField(Document, "Q")                                    
    Call SetToolbar("11001111001111")

    frm1.vspdData.Focus
End Function

'=========================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim strVal, strDel
    Dim intRet
    Dim iColSep, iRowSep
    iColSep = Parent.gColSep
    iRowSep = Parent.gRowSep
    
    'On Error Resum Next
    Err.Clear

    DbSave = False

    Call DisableToolBar(parent.TBC_SAVE)
    Call LayerShowHide(1)

    With frm1
        .txtMode.value = parent.UID_M0002
    

        lGrpCnt = 1
        strVal = ""
        strDel = ""
    

        For lRow = 1 To .vspdData.MaxRows
            Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")
                Case ggoSpread.InsertFlag
                    strVal = strVal & "C" & iColSep                    
                Case ggoSpread.UpdateFlag
                    strVal = strVal & "U" & iColSep                    
                Case ggoSpread.DeleteFlag                            
                    strDel = strDel & "D" & iColSep                    
            End Select            

            Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag                            
					strVal = strVal & lRow                      & iColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_LogonGp, lRow, "X", "X"))      & iColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_LogonGpNm, lRow, "X", "X"))      & iColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_AppServerId , lRow, "X", "X"))      & iColSep   
					strVal = strVal & "0"      & iColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_DB_Server_Id, lRow, "X", "X"))      & iColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_DB_Nm, lRow, "X", "X"))      & iColSep   
					strVal = strVal & "1"      & iRowSep

                    If CheckNumeric(Trim(GetSpreadText(.vspdData, C_Port, lRow, "X", "X"))) = 1 Then
                         IntRet = DisplayMsgBox("210213", "x", "x", "x")
                       'Port번호에 숫자 이외의 데이터는 입력할 수 없습니다.
                       Call LayerShowHide(0)
                       Exit Function
                    End if

                    If CheckNumeric(Trim(GetSpreadText(.vspdData, C_DsnNo, lRow, "X", "X"))) = 1 Then
                         IntRet = DisplayMsgBox("210214", "x", "x", "x")
                       'DSN번호에 숫자 이외의 데이터는 입력할 수 없습니다.
                       Call LayerShowHide(0)
                       Exit Function
                    End if

                    lGrpCnt = lGrpCnt + 1

                Case ggoSpread.DeleteFlag                            '☜: 삭제 
					strDel = strDel & lRow                      & iColSep                
					strDel = strDel & Trim(GetSpreadText(.vspdData, C_LogonGp, lRow, "X", "X"))      & iRowSep
  
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
    
    If DbQuery = False Then
       Call RestoreToolBar()
    End If
    
End Function

'=========================================================================================================
Function DbDelete() 
End Function


'=========================================================================================================
Function OpenLogonGp()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "로그온그룹 팝업"                ' 팝업 명칭 
    arrParam(1) = "z_logon_gp"                        ' TABLE 명칭 
    arrParam(2) = frm1.txtLogonGp.value                ' Code Condition
    arrParam(3) = ""                                ' Name Cindition
    arrParam(4) = ""                                ' Where Condition
    arrParam(5) = "로그온그룹"                    ' 조건필드의 라벨 명칭 
    
    arrField(0) = "logon_gp"                        ' Field명(0)
    arrField(1) = "logon_gp_nm"                        ' Field명(1)
    
    arrHeader(0) = "로그온그룹 코드"            ' Header명(0)
    arrHeader(1) = "로그온그룹"                    ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetLogonGp(arrRet)
    End If    

	frm1.txtLogonGp.focus
	Set gActiveElement = document.activeElement
End Function

'=========================================================================================================
Function SetLogonGp(Byval arrRet)
    With frm1
        .txtLogonGp.value = Trim(arrRet(0))
        .txtLogonGpNm.value = arrRet(1)
    End With
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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>로그온 그룹 관리</font></td>
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
                                    <TD CLASS="TD5">로그온 그룹</TD>
                                    <TD CLASS="TD656">
                                        <INPUT TYPE=TEXT NAME="txtLogonGp" SIZE=30 MAXLENGTH=30 tag="11" ALT="로그온 그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLogonGp" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenLogonGp">
                                        <INPUT TYPE=TEXT NAME="txtLogonGpNm" SIZE=30 tag="14">
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
                                    <script language =javascript src='./js/za002ma1_vaSpread1_vspdData.js'></script>
                                </TD>
                            </TR>
                        </TABLE>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="Za002mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hLogonGp" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
