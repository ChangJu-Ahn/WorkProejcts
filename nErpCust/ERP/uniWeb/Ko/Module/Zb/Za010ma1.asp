
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : System Management
*  3. Program ID           : za010ma1
*  4. Program Name         : Audit Info Detail Query
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2000.03.13
*  8. Modified date(Last)  : 2002.06.24
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

Const BIZ_PGM_ID = "Za010mb1.asp"            
'=========================================================================================================
Dim C_OccurDate
Dim C_OccurTime
Dim C_Trans
Dim C_User

<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim lgArrCondition(12)

Dim gblnWinEvent
Dim strReturn
Dim IsOpenPop        

Dim FirstDate
Dim LastDate

'=========================================================================================================
Sub InitSpreadPosVariables()
    C_OccurDate = 1                                                            
    C_OccurTime = 2
    C_Trans     = 3
    C_User      = 4
End Sub


'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
    lgStrPrevKey = ""
    lgLngCurRows = 0                            
    '---- Coding part--------------------------------------------------------------------
    gblnWinEvent = False    

    lgSortKey = 1
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
    lgArrCondition(11) = frm1.txtTable.Value
    lgArrCondition(12) = ""
End Sub

'=========================================================================================================
Sub SetDefaultVal()
    frm1.txtTable.Value= ReadCookie("Za009ma1_TableID")
    WriteCookie "Za009ma1_TableID", "" 'Delete
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
        .OperationMode = 5        
        Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)

        .ReDraw = false                   
        .MaxCols = C_User + 1                        
        .Col = .MaxCols                                    
        .ColHidden = True
    
        .MaxRows = 0

         Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetDate C_OccurDate, "발생 날짜" , 12, 2, Parent.gDateFormat    '1
        ggoSpread.SSSetEdit C_OccurTime, "발생 시간" , 12, 2                '2
        ggoSpread.SSSetEdit C_Trans,     "트랜잭션"  , 10                    '3
        ggoSpread.SSSetEdit C_User,         "사용자 ID" , 12                    '4
    
        'ToolTip display on
        .TextTip = 1 'Fixed
        frm1.vspdData.SetTextTipAppearance "MS Sans Serif", 12, 0, 0, &HC0FFFF, &H0 
        
        .ReDraw = true

        Call SetSpreadLock
    
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
            C_Trans       =  iCurColumnPos(3)
            C_User        =  iCurColumnPos(4)

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

    Call InitCondition
    
    Call frm1.txtTable.focus()
    Set gActiveElement = document.activeElement

    If frm1.txtTable.value <> "" Then
       Call MainQuery()
    Else
       Call SetToolbar("11000000000011")                                        
    End if
    
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub
'=========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    If Row = 0 Then

        ggoSpread.Source = frm1.vspdData
        
        If lgSortKey = 1 Then
             ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If
    
End Sub
'=========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'=========================================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
        If lgStrPrevKey <> "" Then
            If DbQuery = False Then
                Exit Sub
            End if
        End if
    End if
    
    
End Sub
'=========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Err.Clear                                                               

    FncQuery = False                                                        
    

    Call ggoOper.ClearField(Document, "2")                                        
    Call ggoSpread.ClearSpreadData()        
    Call InitVariables
                                                                
    If Not chkfield(Document, "1") Then                                    
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
Function FncPrev() 
    On Error Resume Next                                                    
End Function

'=========================================================================================================
Function FncNext() 
    On Error Resume Next                                                    
End Function

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
        strToDate = UNIConvDate(lgArrCondition(4))     & " " & lgArrCondition(5)
    End If
    
    If strFromDate > strToDate Then
        IntRetCD = DisplayMsgBox("210404", VBOKONLY,"X","X")
        Call OpenFilter()
        Exit Function
    End If
    
    strVal = BIZ_PGM_ID & "?txtFromDt=" & strFromDate
    strVal = strVal        & "&txtToDt="    & strToDate
    strVal = strVal        & "&txtUser="    & Trim(lgArrCondition(9))
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then    
        strVal = strVal & "&txtTable="    & Trim(frm1.htxtTable.value)
    Else
        strVal = strVal & "&txtTable="    & Trim(frm1.txtTable.value)
    End If

    If lgArrCondition(6) = True Then
        strVal = strVal & "&txtT1=I"
    Else
        strVal = strVal & "&txtT1=X"
    End If
    If lgArrCondition(7) = True Then
        strVal = strVal & "&txtT2=U"
    Else
        strVal = strVal & "&txtT2=X"        
    End If
    If lgArrCondition(8) = True Then
        strVal = strVal & "&txtT3=D"
    Else
        strVal = strVal & "&txtT3=X"        
    End If

    strVal = strVal        & "&txtMaxRows="    & frm1.vspdData.MaxRows
    strVal = strVal        & "&lgStrPrevKey=" & lgStrPrevKey
    strVal = strVal     & "&txtMode="        & Parent.UID_M0001

    Call LayerShowHide(1)
    Call RunMyBizASP(MyBizASP, strVal)                                        

    DbQuery = True
End Function
'=========================================================================================================
Function DbQueryOk()                                                        
    

    lgIntFlgMode = Parent.OPMD_UMODE                                                
    
    Call ggoOper.LockField(Document, "Q")                                    
    Call SetToolbar("11000000000111")            
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
'=========================================================================================================
Function OpenTable()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "테이블 팝업"                        ' 팝업 명칭 
    arrParam(1) = "(select distinct table_id from z_audit_policy) a, z_table_info b"        ' TABLE 명칭 
    arrParam(2) = frm1.txtTable.value                        ' Code Condition
    arrParam(3) = ""                                    ' Name Cindition

    arrParam(4) = "a.table_id=b.table_id And b.lang_cd= " & FilterVar(Parent.gLang, "''", "S") & ""    ' Where Condition
    arrParam(5) = "테이블"                            ' 조건필드의 라벨 명칭 
    
    arrField(0) = "ED24" & Parent.gColSep & "a.table_id"                            ' Field명(0)
    arrField(1) = "b.table_nm"                            ' Field명(1)
    
    arrHeader(0) = "테이블 ID"                        ' Header명(0)
    arrHeader(1) = "테이블 명"                        ' Header명(1)

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=520px; dialogHeight=455px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        frm1.txtTable.value = arrRet(0)
        frm1.txtTableNm.value = arrRet(1)
    End If    
    Set gActiveElement = document.activeElement
End Function

'=========================================================================================================
'    Name : OpenFilter()
'    Description : Filter PopUp
'=========================================================================================================
Function OpenFilter()
    Dim lngCount
    Dim arrRet
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZA010PA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZA010PA1", "x")
        IsOpenPop = False
        Exit Function
    End If

    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,lgArrCondition, arrField, arrHeader), _
        "dialogWidth=600px; dialogHeight=350px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False

    If arrRet(0) = "" Then
        Exit Function
    Else
        For lngCount = 0 To 12
            lgArrCondition(lngCount) = arrRet(lngCount)
        Next
        
            With frm1
                .txtTable.value = lgArrCondition(11)
                .txtTableNm.value = lgArrCondition(12)
            End With        
        
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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>감사 정보 상세 조회</font></td>
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
                    <TD HEIGHT=20 WIDTH=100%>
                        <FIELDSET CLASS="CLSFLD">
                            <TABLE <%=LR_SPACE_TYPE_40%>>
                                <TR>
                                    <TD CLASS="TD5" STYLE="Width:15%">테이블 ID</TD>
                                    <TD CLASS="TD6" STYLE="Width:85%">
                                        <INPUT TYPE=TEXT NAME="txtTable" SIZE=30  tag="12X" ALT="테이블 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTableId" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTable()">
                                        <INPUT TYPE=TEXT NAME="txtTableNm" size=40 tag="14X">
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
                                <TD HEIGHT="100%" WIDTH=*>
                                    <script language =javascript src='./js/za010ma1_vaSpread1_vspdData.js'></script>
                                </TD>
                            </TR>
                        </TABLE>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtTable" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
