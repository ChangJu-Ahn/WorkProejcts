
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Login History Management, UI
'*  2. Function Name        : 
'*  3. Program ID           : ZA003ma1.asp
'*  4. Program Name       	: 
'*  5. Program Desc         : Lists login history information in details and manages locking status.
'*  6. Comproxy List        : ADO query program.
'*  7. Modified date(First) : 2002/05/20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park Sang Hoon
'* 10. Modifier (Last)      : Park Sang Hoon
'* 11. Comment              :
'**********************************************************************************************-->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                                            


Const BIZ_PGM_ID = "Za003mb1.asp"            

Dim C_LoginDate
Dim C_LoginTime
Dim C_LogoutDate
Dim C_LogoutTime
Dim C_UserId
Dim C_UserNm
Dim C_Status
Dim C_ClientId
Dim C_ClientIp
Dim C_hLoginDateTo

<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim lgArrCondition(13)
Dim IsOpenPop          

Sub InitSpreadPosVariables()
    C_LoginDate = 1                                                            
    C_LoginTime = 2
    C_LogoutDate = 3
    C_LogoutTime = 4
    C_UserId = 5
    C_UserNm = 6
    C_Status = 7
    C_ClientId = 8
    C_ClientIp = 9
    C_hLoginDateTo = 10
End Sub

'=========================================================================================================
Sub InitVariables()
    
    lgIntFlgMode      = parent.OPMD_CMODE                   
    lgBlnFlgChgValue  = False                        
    lgIntGrpCount     = 0                             
    lgStrPrevKey = "" 
    lgLngCurRows      = 0                            
    
End Sub

'=========================================================================================================
Sub GetStartAndCurrentDate(StartDate,CurrentDate)
   Dim strYear,strMonth,strDay   
   Call ExtractDateFrom("<%=GetSvrDate()%>",Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
   StartDate = UniConvYYYYMMDDToDate(Parent.gDateFormat,strYear,strMonth,"01")   
   CurrentDate = UniConvYYYYMMDDToDate(Parent.gDateFormat,strYear,strMonth,strDay)
End Sub

'=========================================================================================================
Sub InitCondition()
    Dim thisTime
    Dim StartDate
    Dim CurrentDate
    
    Call GetStartAndCurrentDate(StartDate,CurrentDate)    
       lgArrCondition(0) = 0
    lgArrCondition(1) = StartDate
    lgArrCondition(2) = "00:00:00"
    
    thisTime = Time()

    lgArrCondition(3) = 0
    lgArrCondition(4) = CurrentDate
    lgArrCondition(5) = Right("0" & Hour(thisTime), 2) & ":" _
                        & Right("0" & Minute(thisTime), 2) & ":" _
                        & Right("0" & Second(thisTime), 2) 
    
    lgArrCondition(6) = True
    lgArrCondition(7) = True
    lgArrCondition(8) = True
    lgArrCondition(9) = True
    lgArrCondition(10) = True
    
    If ReadCookie("Za014ma1_UsrId") <> "" Then
        lgArrCondition(11) = ReadCookie("Za014ma1_UsrId")
        WriteCookie "Za014ma1_UsrId", "" 'Delete            
    Else
        lgArrCondition(11) = ""
    End If

    lgArrCondition(12) = ""
    lgArrCondition(13) = ""
End Sub

'=========================================================================================================
Sub SetDefaultVal()
End Sub

'=========================================================================================================
Sub InitSpreadSheet()

    Call InitSpreadPosVariables()
    
    With  frm1.vspdData

        ggoSpread.Source = frm1.vspdData
'        .OperationMode = 5                               'Sets spreadsheet to operate like an extended-selection list box                    
        Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)

        .ReDraw = false
        .MaxCols = C_hLoginDateTo + 1                                                
        .MaxRows = 0

        Call GetSpreadColumnPos("A")        
           
        ggoSpread.SSSetDate C_LoginDate, "Login 날짜", 12, 2, parent.gDateFormat '1
        ggoSpread.SSSetEdit C_LoginTime, "Login 시간", 12, 2 '2
        ggoSpread.SSSetDate C_LogoutDate, "Logout 날짜", 12, 2, parent.gDateFormat '3
        ggoSpread.SSSetEdit C_LogoutTime, "Logout 시간", 12, 2 '4
        ggoSpread.SSSetEdit C_UserId, "사용자 ID", 12 '5
        ggoSpread.SSSetEdit C_UserNm, "사용자명", 19 '6
        ggoSpread.SSSetEdit C_Status, "상태", 12 '7
        ggoSpread.SSSetEdit C_ClientId, "클라이언트명", 12 '8
        ggoSpread.SSSetEdit C_ClientIp, "클라이언트 IP", 12 '9
        ggoSpread.SSSetEdit C_hLoginDateTo, "", 12 '9        
        'ggoSpread.SSSetSplit2(2)
        .ReDraw = true

        Call SetSpreadLock 
        
        Call ggoSpread.MakePairsColumn(C_LoginDate,C_LoginTime,"1")
        Call ggoSpread.MakePairsColumn(C_LogoutDate,C_LogoutTime,"1")

        Call ggoSpread.SSSetColHidden(C_hLoginDateTo,C_hLoginDateTo,True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock -1, -1
        .vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetProtected C_LoginDate, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_LoginTime, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_LogoutDate, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_LogoutTime, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_UserId, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_UserNm, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_Status, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_ClientId, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_ClientIp, pvStartRow, pvEndRow
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

            C_LoginDate      =  iCurColumnPos(1)
            C_LoginTime      =  iCurColumnPos(2)
            C_LogoutDate     =  iCurColumnPos(3)
            C_LogoutTime     =  iCurColumnPos(4)
            C_UserId         =  iCurColumnPos(5)
            C_UserNm         =  iCurColumnPos(6)
            C_Status         =  iCurColumnPos(7)
            C_ClientId       =  iCurColumnPos(8)
            C_ClientIp       =  iCurColumnPos(9)
            C_hLoginDateTo   =  iCurColumnPos(10)
            
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
    Call ggoOper.LockField(Document, "N")                                          
    '----------  Coding part  -------------------------------------------------------------
    Call InitSpreadSheet
    Call InitVariables                                                      
    Call SetDefaultVal
    Call SetToolbar("11000000000111")                                        
    Call InitCondition
    Call FncQuery()
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
    Call SetPopupMenuItemInf("0000111111")    
    
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
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
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
Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                                  


    Call ggoOper.ClearField(Document, "2")                                    
    Call ggoSpread.ClearSpreadData()    
    Call InitVariables                                                         
                                                                         
    If DbQuery() = False Then                                 
        Exit Function
    End if
       
    FncQuery = True                                                                
    
End Function

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
Function FncPrint() 
    Call parent.FncPrint()
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
        IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x","x")                
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
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey

    Dim strFromDate      
    Dim strToDate
    Dim strStatus(5)
    
    Dim i
    
    If lgArrCondition(0) = 0 Then
        strFromDate = "1900-01-01 00:00:00"
    Else
        strFromDate = UNIConvDate(lgArrCondition(1)) & " " & lgArrCondition(2)
    End If
    
    If lgArrCondition(3) = 0 Then
        strToDate = "2999-12-31 23:59:59"
    Else
        strToDate = UNIConvDate(lgArrCondition(4)) & " " & lgArrCondition(5)
    End If
    
    For i = 1 To 5
        If lgArrCondition(i+5) = True Then
            strStatus(i) = CStr(i)
        Else
            strStatus(i) = "0"
        End If
    Next

    DbQuery = False
    
    Err.Clear                                                               

    Dim strVal
    
    With frm1
        If lgIntFlgMode = parent.OPMD_UMODE Then                                                
            strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001                            
            strVal = strVal & "&txtFromDt="                    & strFromDate
            strVal = strVal & "&txtToDt="                   & Trim(.htxtToDt.value)            
            strVal = strVal & "&txtUser="                   & Trim(lgArrCondition(11))
            strVal = strVal & "&txtClient="                 & Trim(lgArrCondition(13))
            strVal = strVal & "&txtS1="                     & strStatus(1)
            strVal = strVal & "&txtS2="                     & strStatus(2)
            strVal = strVal & "&txtS3="                     & strStatus(3)
            strVal = strVal & "&txtS4="                     & strStatus(4)
            strVal = strVal & "&txtS5="                     & strStatus(5)
            strVal = strVal & "&txtKeyStream="        & lgKeyStream
            strVal = strVal & "&txtMaxRows="          & .vspdData.MaxRows
            strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        Else                    
            strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001                            
            strVal = strVal & "&txtFromDt="                    & strFromDate
            strVal = strVal & "&txtToDt="                  & strToDate    
            strVal = strVal & "&txtUser="                  & Trim(lgArrCondition(11))
            strVal = strVal & "&txtClient="                 & Trim(lgArrCondition(13))
            strVal = strVal & "&txtS1="                     & strStatus(1)
            strVal = strVal & "&txtS2="                     & strStatus(2)
            strVal = strVal & "&txtS3="                     & strStatus(3)
            strVal = strVal & "&txtS4="                     & strStatus(4)
            strVal = strVal & "&txtS5="                     & strStatus(5)
            strVal = strVal & "&txtKeyStream="        & lgKeyStream
            strVal = strVal & "&txtMaxRows="          & .vspdData.MaxRows
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
    frm1.vspdData.Focus
End Function

'----------------------------------------  Open Popup Functions  ------------------------------------------
'    Name : Open???()
'    Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenFilter()
    Dim i
    Dim arrRet
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZA003PA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZA003PA2", "x")
        IsOpenPop = False
        Exit Function
    End If

    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,lgArrCondition, arrField, arrHeader), _
        "dialogWidth=600px; dialogHeight=350px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        For i = 0 To 13
            lgArrCondition(i) = arrRet(i)
        Next
        Call FncQuery
    End If
End Function

Function OpenLock()
    Dim i
    Dim arrRet
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD

    If IsOpenPop = True Then Exit Function
    
    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZA003PA2")

    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZA003PA2", "x")
        IsOpenPop = False
        Exit Function
    End If

    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,lgArrCondition, arrField, arrHeader), _
        "dialogWidth=600px; dialogHeight=350px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False
    
    If arrRet(0, 0) = "" Then
        Exit Function
        Call FncQuery
    End If

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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>로그인 내역 관리</font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* align=right><A href="VBScript:OpenLock">Locking 해제</A>
                    &nbsp;|&nbsp;<A href="VBScript:OpenFilter">상세 검색</A></TD>
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
                    <TD WIDTH=100% HEIGHT=100% valign=top>
                        <TABLE <%=LR_SPACE_TYPE_40%>>
                            <TR>
                                <TD HEIGHT="100%">
                                    <script language =javascript src='./js/za003ma1_vaSpread1_vspdData.js'></script>
                                </TD>
                            </TR>
                        </TABLE>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtToDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
