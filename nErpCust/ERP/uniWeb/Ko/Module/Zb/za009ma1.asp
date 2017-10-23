<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za009ma1
*  4. Program Name         : Audit Info Overview Query
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000.03.13
*  8. Modified date(Last)  : 2002.06.21
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
Const BIZ_PGM_ID = "Za009mb1.asp"            
Const JUMP_PGM_ID = "Za010ma1"                
Const BIZ_DETAIL = "Za009mb2.asp"            
'=========================================================================================================
Dim C_OccurDate
Dim C_OccurTime
Dim C_TransCd
Dim C_Trans
Dim C_User
Dim C_UserNm
Dim C_TableId

'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim lgArrCondition(12)

'=========================================================================================================
Dim IsOpenPop
Dim lgCurrRow

Dim FirstDate
Dim LastDate
'=========================================================================================================
Sub InitSpreadPosVariables()
    C_OccurDate = 1
    C_OccurTime = 2
    C_TransCd = 3
    C_Trans = 4
    C_User = 5
    C_UserNm = 6
    C_TableId = 7
End Sub

'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           
    lgLngCurRows = 0                            
    
    lgCurrRow = 0
End Sub
'=========================================================================================================
Sub InitCondition()
    Dim thisTime
    
    Call GetFirstAndLastDate(FirstDate,LastDate)
    
       lgArrCondition(0) = 0
    lgArrCondition(1) = FirstDate
    lgArrCondition(2) = "00:00:00"
    
    thisTime = Time()

    lgArrCondition(3) = 0
    lgArrCondition(4) = LastDate
    lgArrCondition(5) = Right("0" & Hour(thisTime), 2) & ":" _
                & Right("0" & Minute(thisTime), 2) & ":" _
                & Right("0" & Second(thisTime), 2) 
    
    lgArrCondition(6) = True
    lgArrCondition(7) = True
    lgArrCondition(8) = True

    lgArrCondition(9) = ""
    lgArrCondition(10) = ""
    lgArrCondition(11) = ""
    lgArrCondition(12) = ""
End Sub

'=========================================================================================================
Sub SetDefaultVal()
End Sub

'=========================================================================================================
'Sub LoadInfTB19029()
'    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
'    <% Call loadInfTB19029A("I", "*", "NOCOOKIE","QA") %>
'End Sub

'=========================================================================================================
Sub InitSpreadSheet()

    Call InitSpreadPosVariables()
    
    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        '.OperationMode = 5        
        Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)

        .ReDraw = false                   
        .MaxCols = C_TableId + 1                        
        .MaxRows = 0

        Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetDate C_OccurDate, "발생 날짜", 12, 2, Parent.gDateFormat '1
        ggoSpread.SSSetEdit C_OccurTime, "발생 시간", 12, 2 '2
        ggoSpread.SSSetEdit C_TransCd, "", 10 '3
        ggoSpread.SSSetEdit C_Trans, "트랜잭션", 10 '4
        ggoSpread.SSSetEdit C_User, "사용자 ID", 13 '5
        ggoSpread.SSSetEdit C_UserNm, "사용자명", 15 '6
        ggoSpread.SSSetEdit C_TableId, "테이블ID", 20 '7
        'ggoSpread.SSSetSplit2(2)
        
        'ToolTip display on
        .TextTip = 1 'Fixed
        frm1.vspdData.SetTextTipAppearance "MS Sans Serif", 12, 0, 0, &HC0FFFF, &H0 'Set font, color
        
        .ReDraw = true

        Call SetSpreadLock

        Call ggoSpread.SSSetColHidden(C_TransCd, C_TransCd, True)
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
Sub SetSpreadColor(ByVal lRow)
End Sub

'=========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_OccurDate   =  iCurColumnPos(1)
            C_OccurTime   =  iCurColumnPos(2)
            C_TransCd     =  iCurColumnPos(3)
            C_Trans       =  iCurColumnPos(4)
            C_User        =  iCurColumnPos(5)
            C_UserNm      =  iCurColumnPos(6)
            C_TableId     =  iCurColumnPos(7)

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

    iCalledAspName = AskPRAspName("ZA009PA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZA009PA1", "x")
        IsOpenPop = False
        Exit Function
    End If
    
    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,lgArrCondition, arrField, arrHeader), _
        "dialogWidth=600px; dialogHeight=350px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        For i = 0 To 12
            lgArrCondition(i) = arrRet(i)
        Next
        
        Call MainQuery
    End If
End Function

'=========================================================================================================
Sub GetDetail()
    Dim strVal, nActiveRow
    With frm1.vspdData
    	nActiveRow = .ActiveRow
        If nActiveRow > 0 And lgCurrRow <> nActiveRow Then
            lgCurrRow = nActiveRow
            strVal = BIZ_DETAIL & "?txtTable=" & Trim(GetSpreadText(frm1.vspdData, C_TableId, nActiveRow, "X", "X"))
            strVal = strVal & "&txtOccurDt=" & UNIConvDate(Trim(GetSpreadText(frm1.vspdData, C_OccurDate, nActiveRow, "X", "X")))
            strVal = strVal & "&txtOccurTm=" & Trim(GetSpreadText(frm1.vspdData, C_OccurTime, nActiveRow, "X", "X"))
            strVal = strVal & "&txtTransCd=" & Trim(GetSpreadText(frm1.vspdData, C_TransCd, nActiveRow, "X", "X"))
            strVal = strVal & "&txtTrans=" & Trim(GetSpreadText(frm1.vspdData, C_Trans, nActiveRow, "X", "X"))
            strVal = strVal & "&txtUser=" & Trim(GetSpreadText(frm1.vspdData, C_User, nActiveRow, "X", "X"))
            strVal = strVal & "&strASPMnuMnuNm=" & "<%=Request("strASPMnuMnuNm")%>"
            strval = strval & "&strRequestMenuID=" & "<%=Request("strRequestMenuID")%>"'''''''''''''''''''''khy200307
            strval = strval & "&strRequestUpperMenuID=" & "<%=Request("strRequestUpperMenuID")%>" '''''''''''''''''''''khy200307
            Call LayerShowHide(1)
            Call RunMyBizASP(Detail, strVal)                                                            
            
        End If
    End With
End Sub
'=========================================================================================================
Sub SetDetailInCaseOfNoData()    '조회된 데이터가 없을때, 우측 Table을 초기화한다.                
    Dim strVal
    
    strVal = BIZ_DETAIL & "?txtTable="
    strVal = strVal & "&txtOccurDt="
    strVal = strVal & "&txtOccurTm="
    strVal = strVal & "&txtTransCd="
    strVal = strVal & "&txtTrans="
    strVal = strVal & "&txtUser="

    strVal = strVal & "&strASPMnuMnuNm=" & "<%=Request("strASPMnuMnuNm")%>"
    strval = strval & "&strRequestMenuID=" & "<%=Request("strRequestMenuID")%>"'''''''''''''''''''''khy200307
    strval = strval & "&strRequestUpperMenuID=" & "<%=Request("strRequestUpperMenuID")%>" '''''''''''''''''''''khy200307

        
    Call LayerShowHide(1)
    Call RunMyBizASP(Detail, strVal)                                                

End Sub

'=========================================================================================================
Function JumpProgram()
    Dim strTable
    strTable = ""

    On Error Resume Next    
    strTable = Detail.document.frm1.txtTable.value
    On Error Goto 0

    WriteCookie "Za009ma1_TableID", Trim(strTable)

    PgmJump(JUMP_PGM_ID)    
End Function
'=========================================================================================================
Sub GetFirstAndLastDate(FirstDate,LastDate)
   Dim strYear,strMonth,strDay   
   Call ExtractDateFrom("<%=GetSvrDate()%>",Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
   FirstDate = UniConvYYYYMMDDToDate(Parent.gDateFormat,strYear,strMonth,"01")
   LastDate = UniConvDateAToB("<%=GetSvrDate()%>",Parent.gServerDateFormat,Parent.gDateFormat)           
End Sub


'=========================================================================================================
Sub Form_Load()
	
	
    Call ggoOper.LockField(Document, "N")                                   
    
    Call InitSpreadSheet                                                    
    Call InitVariables                                                      

    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call SetToolbar("11000000000011")                                        
    
    Call InitCondition
    Call MainQuery
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
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
    
    if (frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)) Then
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
Sub vspdData_Click(ByVal Col, ByVal Row)
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
    Call GetDetail()
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
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

    End With

End Sub
'=========================================================================================================
Sub vspdData_KeyPress(KeyAscii)
    If KeyAscii = 13 Then
        Call GetDetail()
    End If
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
    
    Err.Clear                                                               

    FncQuery = False                                                        
    
    ggoSpread.Source = frm1.vspdData
    
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
    Call parent.FncExport(Parent.C_MULTI)                                                
End Function

'=========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         
End Function

'=========================================================================================================
Function FncExit()
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
    Dim strFromDate      
    Dim strToDate
    Dim IntRetCD

    Err.Clear                                                               

    DbQuery = False
    
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
    
    If strFromDate > strToDate Then
        IntRetCD = DisplayMsgBox("210404", VBOKONLY,"X","X")
        Call SetDetailInCaseOfNoData()
        Call OpenFilter()
        Exit Function
    End If
    
    With frm1
        If lgIntFlgMode = Parent.OPMD_UMODE Then                                                
            strVal = BIZ_PGM_ID & "?txtMode="    & Parent.UID_M0001                               
            strVal = strVal & "&txtFromDt="        & strFromDate
            strVal = strVal & "&txtToDt="        & lgStrPrevKey
            strVal = strVal & "&txtUser="        & Trim(lgArrCondition(9))
            strVal = strVal & "&txtTable="        & Trim(lgArrCondition(11))            
            If lgArrCondition(6) = True Then
                strVal = strVal & "&txtT1=I"
            End If
            If lgArrCondition(7) = True Then
                strVal = strVal & "&txtT2=U"
            End If
            If lgArrCondition(8) = True Then
                strVal = strVal & "&txtT3=D"
            End If
            strVal = strVal & "&txtMaxRows="    & .vspdData.MaxRows
            strVal = strVal & "&lgStrPrevKey="    & lgStrPrevKey    
        Else
            strVal = BIZ_PGM_ID & "?txtMode="    & Parent.UID_M0001                                
            strVal = strVal & "&txtFromDt="        & strFromDate
            strVal = strVal & "&txtToDt="        & strToDate
            strVal = strVal & "&txtUser="        & Trim(lgArrCondition(9))
            strVal = strVal & "&txtTable="        & Trim(lgArrCondition(11))
            If lgArrCondition(6) = True Then
                strVal = strVal & "&txtT1=I"
            End If
            If lgArrCondition(7) = True Then
                strVal = strVal & "&txtT2=U"
            End If
            If lgArrCondition(8) = True Then
                strVal = strVal & "&txtT3=D"
            End If
            strVal = strVal & "&txtMaxRows="    & .vspdData.MaxRows
            strVal = strVal & "&lgStrPrevKey="    & lgStrPrevKey    
        End If
                        
        Call LayerShowHide(1)
        Call RunMyBizASP(MyBizASP, strVal)
    End With
    
    DbQuery = True
End Function

'=========================================================================================================
Function DbQueryOk()                                                        

    lgIntFlgMode = Parent.OPMD_UMODE                                                
    
    Call ggoOper.LockField(Document, "Q")                                    
        
    If frm1.vspdData.MaxRows > 0 Then
        Call SetToolbar("11000000000111")
        If lgIntFlgMode = Parent.OPMD_UMODE Then                                    
            Call GetDetail()
        End If
    End If
    
    frm1.vspdData.Focus
End Function

'=========================================================================================================
Function DbSave() 
    DBSave = False
    DBSave = True
End Function

'=========================================================================================================
Function DbSaveOk()                                                    
End Function

'=========================================================================================================
Function DbDelete() 
    DBDelete = False

    DBDelete = True
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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>감사 정보 개요 조회</font></td>
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
                    <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
                </TR>    
                <TR>
                    <TD WIDTH=100% HEIGHT=100% valign=top>
                        <TABLE <%=LR_SPACE_TYPE_40%>>
                            <TR>
                                <TD HEIGHT="100%" WIDTH=*>
                                    <script language =javascript src='./js/za009ma1_vaSpread1_vspdData.js'></script>
                                </TD>
                                <TD WIDTH=300px>
                                    <IFRAME NAME="Detail" SRC="../../blank.htm" WIDTH=100% HEIGHT=100% FRAMEBORDER=1 SCROLLING=NO noresize></IFRAME>
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
                    <TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpProgram">감사 정보 상세 조회</A></TD>
                    <TD WIDTH=10>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = -1></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
