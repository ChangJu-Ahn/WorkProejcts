<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : User Management
*  3. Program ID           : za004ma1
*  4. Program Name         : Message History Management
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2002/05/20
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : Park Sang Hoon
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                                            
                                                                

'=========================================================================================================
Const BIZ_PGM_ID = "Za004mb1.asp"            
'=========================================================================================================
Dim C_OccurDate
Dim C_OccurTime
Dim C_MsgCd
Dim C_Msg
Dim C_MsgType
Dim C_User
Dim C_UserNm
Dim C_Severity
Dim C_PgmId
Dim C_Client
Dim C_ClientIp
Dim C_hOccurDate
Dim lgStrPrevDt
Dim lgArrCondition(18)
Dim IsOpenPop        

<!-- #Include file="../../inc/lgvariables.inc" -->    

Sub InitSpreadPosVariables()
    C_OccurDate = 1
    C_OccurTime = 2
    C_MsgCd = 3
    C_Msg = 4
    C_MsgType = 5
    C_User = 6
    C_UserNm = 7
    C_Severity = 8
    C_PgmId = 9
    C_Client = 10
    C_ClientIp = 11
    C_hOccurDate = 12
End Sub

'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           
    lgStrPrevDt = ""                           
    lgLngCurRows = 0                            
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
    lgArrCondition(11) = True
    lgArrCondition(12) = True

    If ReadCookie("Za014ma1_UsrId") <> "" Then
        lgArrCondition(13) = ReadCookie("Za014ma1_UsrId")
        WriteCookie "Za014ma1_UsrId", "" 'Delete        
    Else
        lgArrCondition(13) = ""
    End If        
    
    lgArrCondition(14) = ""
    lgArrCondition(15) = ""
    lgArrCondition(16) = ""
    lgArrCondition(17) = ""
    lgArrCondition(18) = ""
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

        .ReDraw = False                
        .MaxCols = C_hOccurDate + 1                        
        .MaxRows = 0
        Call GetSpreadColumnPos("A")        

        ggoSpread.SSSetDate C_OccurDate, "발생 날짜", 15, 2, parent.gDateFormat '1
        ggoSpread.SSSetEdit C_OccurTime, "발생 시간", 12, 2 '2
        ggoSpread.SSSetEdit C_MsgCd, "메세지코드", 13 '3
        ggoSpread.SSSetEdit C_Msg, "메세지", 60, , ,256 '4
        ggoSpread.SSSetEdit C_MsgType, "메세지유형", 20 '5
        ggoSpread.SSSetEdit C_User, "사용자 ID", 12 '6
        ggoSpread.SSSetEdit C_UserNm, "사용자명", 20 '7
        ggoSpread.SSSetEdit C_Severity, "Severity", 12 '8
        ggoSpread.SSSetEdit C_PgmId, "프로그램ID", 12 '9
        ggoSpread.SSSetEdit C_Client, "클라이언트명", 14 '10
        ggoSpread.SSSetEdit C_ClientIp, "클라이언트 IP", 16 '11
        ggoSpread.SSSetEdit C_hOccurDate, "", 16 '12
        'ggoSpread.SSSetSplit2(3)
        'ToolTip display on
        .TextTip = 1 'Fixed
        frm1.vspdData.SetTextTipAppearance "MS Sans Serif", 12, 0, 0, &HC0FFFF, &H0 'Set font, color
        .ReDraw = True

        Call SetSpreadLock 
        
        Call ggoSpread.MakePairsColumn(C_OccurDate,C_OccurTime,"1")

        Call ggoSpread.SSSetColHidden(C_hOccurDate, C_hOccurDate, True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock 1, -1
        .vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetProtected C_OccurDate, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_OccurTime, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_MsgCd, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_Msg, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_MsgType, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_User, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_UserNm, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_Severity, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_PgmId, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_Client, pvStartRow, pvEndRow
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
    
            C_OccurDate      =  iCurColumnPos(1)
            C_OccurTime      =  iCurColumnPos(2)
            C_MsgCd     =  iCurColumnPos(3)
            C_Msg     =  iCurColumnPos(4)
            C_MsgType         =  iCurColumnPos(5)
            C_User         =  iCurColumnPos(6)
            C_UserNm         =  iCurColumnPos(7)
            C_Severity       =  iCurColumnPos(8)
            C_PgmId       =  iCurColumnPos(9)
            C_Client   =  iCurColumnPos(10)
            C_ClientIp   =  iCurColumnPos(11)
            C_hOccurDate   =  iCurColumnPos(12)            
            
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
    Call LoadInfTB19029                                                         
    Call ggoOper.LockField(Document, "N")                                   
    '----------  Coding part  -------------------------------------------------------------
    Call InitSpreadSheet                                                    
    Call InitVariables                                                          
    Call SetToolbar("11000000000111")                                        
    Call InitCondition
    Call MainQuery
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
End Function

'=========================================================================================================
Function FncCopy()
End Function

'=========================================================================================================
Function FncCancel() 
End Function

'=========================================================================================================
Function FncInsertRow() 
End Function

'=========================================================================================================
Function FncDeleteRow() 
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
    Call parent.FncExport(parent.C_MULTI)                                                
End Function

'=========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         
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

    Dim strFromDate      
    Dim strToDate
    Dim strS(4)
    Dim strM(3)
    
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
    
    For i = 1 To 4
        If lgArrCondition(i+5) = True Then
            strS(i) = CStr(i)
        Else
            strS(i) = "0"
        End If
    Next
    
    If lgArrCondition(10) = True Then
        strM(1) = "A"
    Else
        strM(1) = "0"
    End If
    
    If lgArrCondition(11) = True Then
        strM(2) = "D"
    Else
        strM(2) = "0"
    End If
    
    If lgArrCondition(12) = True Then
        strM(3) = "S"
    Else
        strM(3) = "0"
    End If
    

    DbQuery = False
    
    Err.Clear                                                               

    Dim strVal
    
    With frm1
        If lgIntFlgMode = parent.OPMD_UMODE Then                                                                
            strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001                            
            strVal = strVal & "&txtFromDt=" & strFromDate
            strVal = strVal & "&txtToDt=" & Trim(.hOccurDt.value)            
            strVal = strVal & "&txtUser=" & Trim(lgArrCondition(13))
            strVal = strVal & "&txtMsg=" & Trim(lgArrCondition(15))
            strVal = strVal & "&txtPgm=" & Trim(lgArrCondition(16))
            strVal = strVal & "&txtClient=" & Trim(lgArrCondition(18))
            strVal = strVal & "&txtS1=" & strS(1)
            strVal = strVal & "&txtS2=" & strS(2)
            strVal = strVal & "&txtS3=" & strS(3)
            strVal = strVal & "&txtS4=" & strS(4)
            strVal = strVal & "&txtM1=" & strM(1)
            strVal = strVal & "&txtM2=" & strM(2)
            strVal = strVal & "&txtM3=" & strM(3)
            strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
            strVal = strVal & "&lgStrPrevDt=" & lgStrPrevDt
        Else            
            strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001                            
            strVal = strVal & "&txtFromDt=" & strFromDate
            strVal = strVal & "&txtToDt=" & strToDate
            strVal = strVal & "&txtUser=" & Trim(lgArrCondition(13))
            strVal = strVal & "&txtMsg=" & Trim(lgArrCondition(15))
            strVal = strVal & "&txtPgm=" & Trim(lgArrCondition(16))
            strVal = strVal & "&txtClient=" & Trim(lgArrCondition(18))
            strVal = strVal & "&txtS1=" & strS(1)
            strVal = strVal & "&txtS2=" & strS(2)
            strVal = strVal & "&txtS3=" & strS(3)
            strVal = strVal & "&txtS4=" & strS(4)
            strVal = strVal & "&txtM1=" & strM(1)
            strVal = strVal & "&txtM2=" & strM(2)
            strVal = strVal & "&txtM3=" & strM(3)
            strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
            strVal = strVal & "&lgStrPrevDt=" & lgStrPrevDt        
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

'=========================================================================================================
'    Name : OpenFilter()
'    Description : Filter PopUp
'=========================================================================================================
Function OpenFilter()
    Dim i
    Dim arrRet
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZA004PA1")

    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZA004PA1", "x")
        IsOpenPop = False
        Exit Function
    End If
        
    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,lgArrCondition, arrField, arrHeader), _
        "dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        For i = 0 To 18
            lgArrCondition(i) = arrRet(i)
        Next
        
        Call MainQuery
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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>메세지 내역 관리</font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* align=right><A href="VBScript:OpenFilter">상세 검색</A></TD>
                    <TD WIDTH=10>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR HEIGHT=*>
        <TD WIDTH=100% CLASS="Tab11">
            <TABLE <%=LR_SPACE_TYPE_20%>>
                <TR>
                    <TD WIDTH=100% HEIGHT=100% valign=top>
                        <TABLE <%=LR_SPACE_TYPE_40%>>
                            <TR>
                                <TD HEIGHT="100%">
                                    <script language =javascript src='./js/za004ma1_vaSpread1_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="hOccurDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
