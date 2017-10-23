
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Basis Architect                                                            *
'*  2. Function Name        : User Management, UI                                                    *
'*  3. Program ID             : ZA003PA2                                                                *
'*  4. Program Name        : Login History Popup 2                                                *
'*  5. Program Desc          : Lists and updates locking status.                                *
'*  7. Modified date(First)  : 2002/05/21                                                            *
'*  8. Modified date(Last)  : 2002/05/21                                                            *
'*  9. Modifier (First)         :    PARK, SANGHOON                                                    *
'* 10. Modifier (Last)        : PARK, SANGHOON                                                    *
'* 11. Comment              :                                                                                *
'********************************************************************************************************-->
<HTML>
<HEAD>


<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">        

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/IncCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/IncCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/IncCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/IncCliVariables.vbs"></SCRIPT>
<Script Language="JavaScript" SRC="../../inc/incImage.js"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                                    

Const BIZ_PGM_ID = "za003pb2.asp"            
'==========================================================================================================%>
Dim arrParent
Dim PopupParent

Dim C_LoginDate
Dim C_LoginTime
Dim C_LogoutDate
Dim C_LogoutTime
Dim C_UserId
Dim C_UserNm
Dim C_Status
Dim C_ClientId
Dim C_ClientIp

<!-- #Include file="../../inc/lgvariables.inc" -->    

'==========================================================================================================%>
Dim gblnWinEvent
Dim arrReturn
Dim IsOpenPop          

'=========================================================================================================
arrParent   = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

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
End Sub    

'=========================================================================================================
Sub InitVariables()
    
    lgIntFlgMode      = PopupParent.OPMD_CMODE
    lgBlnFlgChgValue  = False
    lgIntGrpCount     = 0
    lgStrPrevKeyIndex = ""
    lgLngCurRows      = 0
    
    '---- Coding part--------------------------------------------------------------------    
    gblnWinEvent = False
    'Self.Returnvalue = Array("")    
    
End Sub

'=========================================================================================================
Sub SetDefaultVal()
    'Self.Returnvalue = Array("")    
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
        Call ggoSpread.Spreadinit("V20021124",,PopupParent.gAllowDragDropSpread)
                
        .ReDraw = false        
        .MaxCols = C_ClientIp + 1                        
        .MaxRows = 0

        Call GetSpreadColumnPos("A")        
    
        ggoSpread.SSSetDate C_LoginDate, "Login 날짜", 12, 2, PopupParent.gDateFormat '1
        ggoSpread.SSSetEdit C_LoginTime, "Login 시간", 12, 2 '2
        ggoSpread.SSSetDate C_LogoutDate, "Logout 날짜", 12, 2, PopupParent.gDateFormat '3
        ggoSpread.SSSetEdit C_LogoutTime, "Logout 시간", 12, 2 '4
        ggoSpread.SSSetEdit C_UserId, "사용자 ID", 12 '5
        ggoSpread.SSSetEdit C_UserNm, "사용자명", 19 '6
        ggoSpread.SSSetEdit C_Status, "상태", 12 '7
        ggoSpread.SSSetEdit C_ClientId, "클라이언트명", 12 '8
        ggoSpread.SSSetEdit C_ClientIp, "클라이언트 IP", 12 '9
        .ReDraw = true
        
        Call SetSpreadLock

        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
    
    End With
    
End Sub



'=========================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock 1, -1, 9
        .vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Sub SetSpreadColor(ByVal lRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetProtected C_LoginDate, lRow, lRow
        ggoSpread.SSSetProtected C_LoginTime, lRow, lRow
        ggoSpread.SSSetProtected C_LogoutDate, lRow, lRow
        ggoSpread.SSSetProtected C_LogoutTime, lRow, lRow
        ggoSpread.SSSetProtected C_UserId, lRow, lRow
        ggoSpread.SSSetProtected C_UserNm, lRow, lRow
        ggoSpread.SSSetProtected C_Status, lRow, lRow
        ggoSpread.SSSetProtected C_ClientId, lRow, lRow
        ggoSpread.SSSetProtected C_ClientIp, lRow, lRow
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
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")    

    '----------  Coding part  -------------------------------------------------------------
    Call InitSpreadSheet
    Call InitVariables                                                      
    '----------  Coding part  -------------------------------------------------------------    
    Redim arrReturn(0, 0)
    arrReturn(0, 0) = ""
    Self.Returnvalue = arrReturn
    
    If DbQuery = False Then
        Exit Sub
    End if
    
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
Function vspdData_DblClick(ByVal Col, ByVal Row)
    If frm1.vspdData.MaxRows > 0 Then
        If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then                
            Call DBSave()
        End If
    End If
End Function

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

        If NewRow = .MaxRows Then
            If lgStrPrevKeyIndex <> "" Then
                If DbQuery = False Then
                    Exit Sub
                End if
            End if
        End if
    End With
End Sub    

'=========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
        If lgStrPrevKeyIndex <> "" Then
            If DbQuery = False Then
                Exit Sub
            End if
        End if
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
Function FncPrint() 
    Call parent.FncPrint()
End Function

'=========================================================================================================
Function FncExcel() 
    On Error Resume next                                                    
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

    Call LayerShowHide(1)
    DbQuery = False
    
    'Err.Clear                                                               
    Dim strFromDate, strToDate
    Dim strVal

    strFromDate = "1900-01-01 00:00:00"
    strToDate = "2999-12-31 23:59:59"

    With frm1    
        strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001                            '☜: 
        
        If Trim(.txtUsrId.value) = "" Then
           strVal = strVal & "&txtUsrId=" & ""
        Else
           strVal = strVal & "&txtUsrId=" & Trim(.txtUsrId.value)
        End if
        
        strVal = strVal     & "&txtFromDt=" & strFromDate
        strVal = strVal     & "&txtToDt=" & strToDate
        strVal = strVal     & "&txtKeyStream="      & lgKeyStream
        strVal = strVal     & "&txtMaxRows="        & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex

        Call RunMyBizASP(MyBizASP, strVal)                                        

    End With       
    DbQuery = True

End Function

'=========================================================================================================
Function DbQueryOk()                                                            
    Call ggoOper.LockField(Document, "Q")                                    
    frm1.vspdData.Focus
End Function

'=========================================================================================================
Function DBSave()

    Dim intColCnt, intRowCnt, intInsRow, strVal
    Dim iColSep, iRowSep

	iColSep = PopupParent.gColSep
	iRowSep = PopupParent.gRowSep
    intInsRow = frm1.vspdData.SelBlockRow2 - frm1.vspdData.SelBlockRow + 1

    if frm1.vspdData.SelBlockRow2 > 0 Then 

        Redim arrReturn(intInsRow - 1, frm1.vspdData.MaxCols - 1)
    
		For intRowCnt = frm1.vspdData.SelBlockRow To frm1.vspdData.SelBlockRow2
		    strVal = strVal & Trim(GetSpreadText(frm1.vspdData, C_UserId, intRowCnt, "X", "X")) & iColSep
		    strVal = strVal & Trim(GetSpreadText(frm1.vspdData, C_LoginDate, intRowCnt, "X", "X")) & " "
		    strVal = strVal & Trim(GetSpreadText(frm1.vspdData, C_LoginTime, intRowCnt, "X", "X")) & iRowSep
		Next
		
		frm1.txtSpread.value = strVal
		frm1.txtMode.value = PopupParent.UID_M0002
		frm1.txtMaxRows.value = intInsRow

        Call LayerShowHide(1)

        Call ExecMyBizASP(frm1, BIZ_PGM_ID)        

    End if            

    DBSave = true
                
End Function

'=========================================================================================================
Function DbSaveOk()                                                    
   
    Call InitVariables
    frm1.vspdData.MaxRows = 0
    Call FncQuery()

End Function

'=========================================================================================================
Function CancelClick()        
    Self.Close()
End Function

'=========================================================================================================
'----------------------------------------  OpenLogonGp()  ------------------------------------------
'    Name : OpenLogonGp()
'    Description : Country PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenUsrId(Byval strCode, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "사용자정보 팝업"                                           ' 팝업 명칭 
    arrParam(1) = "z_usr_mast_rec"                                                ' TABLE 명칭 
    arrParam(2) = strCode                                                         ' Code Condition
    arrParam(3) = ""                                                              ' Name Cindition
    arrParam(4) = ""                                                              ' Where Condition
    arrParam(5) = "사용자 ID"            
    
    arrField(0) = "Usr_id"                                                        ' Field명(0)
    arrField(1) = "Usr_nm"                                                        ' Field명(1)
    
    arrHeader(0) = "사용자"                                                   ' Header명(0)
    arrHeader(1) = "사용자명"                                                 ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",  Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetUsrId(arrRet, iWhere)        'return value setting
    End If    
	frm1.txtUsrId.focus
End Function

'=========================================================================================================
Function SetUsrId(Byval arrRet, Byval iWhere)
    With frm1
        If iWhere = 0 Then
            .txtUsrId.value = arrRet(0)
            .txtUsrNm.value = arrRet(1)
        End If
    End With
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->    
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
    <TR>
        <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
    </TR>
    <TR>
        <TD HEIGHT=20 WIDTH=100%>
            <FIELDSET CLASS="CLSFLD">
                <TABLE <%=LR_SPACE_TYPE_40%>>
                    <TR>
                        <TD CLASS="TD5">사 용 자</TD>
                        <TD CLASS="TD656">
                            <INPUT TYPE=TEXT NAME="txtUsrId" SIZE=13 MAXLENGTH=13 tag="11N" ALT="사용자"  LANGUAGE=javascript ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUsrId" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUsrId frm1.txtUsrId.value,0">&nbsp;
                            <INPUT TYPE=TEXT ID="txtUsrNm" NAME="arrCond" SIZE=20 tag="14"></TD>
                        </TD>
                    </TR>
                    <TR>
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
                        <script language =javascript src='./js/za003pa2_vaSpread1_vspdData.js'></script>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <tr>
        <TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <TR HEIGHT="20">
        <TD WIDTH="100%">
            <TABLE <%=LR_SPACE_TYPE_30%>>
                <TR>
                        <TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
                            <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
                        <TD WIDTH=30% ALIGN=RIGHT>
                            <IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="DBSave()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
                            <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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
