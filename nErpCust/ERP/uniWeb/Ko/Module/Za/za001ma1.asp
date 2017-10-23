<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BASIS ARCHITECT
*  2. Function Name        : 
*  3. Program ID           : ZA001MA1.ASP
*  4. Program Name         : USER ENTRY
*  5. Program Desc         : USER MANAGEMENT
*  6. Comproxy List        : PZAG004.DLL
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                        '

'=========================================================================================================
Const BIZ_PGM_ID = "za001mb1.asp"                                                
Const JUMP_PGM_ID = "za014ma1"            
'=========================================================================================================
Dim    C_Usr_id 
Dim    C_Usr_nm 
Dim    C_Usr_eng_nm 
Dim    C_Pwd 
Dim    C_Co_cd 
Dim    C_Co_cd_nm 
Dim    C_Logon_gp 
Dim    C_Logon_gp_nm 
Dim    C_Interface_id 
Dim    C_Pwd_valid_dt 
Dim    C_Usr_valid_dt 

<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim IsOpenPop          


'=========================================================================================================
Sub InitSpreadPosVariables()
    C_Usr_id = 1                                                            
    C_Usr_nm = 2                                                              
    C_Usr_eng_nm = 3
    C_Pwd = 4
    C_Co_cd = 5
    C_Co_cd_nm = 6
    C_Logon_gp = 7
    C_Logon_gp_nm = 8
    C_Usr_valid_dt = 9
    C_Interface_id = 10
    C_Pwd_valid_dt = 11
End Sub

'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE            
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           
    lgLngCurRows = 0                            
    
End Sub

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
        .MaxCols = C_Pwd_valid_dt + 1                                   
        .MaxRows = 0

         Call GetSpreadColumnPos("A")
       
        ggoSpread.SSSetEdit C_Usr_id, "사용자ID", 12, , ,13
        ggoSpread.SSSetEdit C_Usr_nm, "사용자명", 16, , ,30
        ggoSpread.SSSetEdit C_Usr_eng_nm, "사용자명(영문)", 18, , ,50
        ggoSpread.SSSetEdit C_Pwd, "비밀번호", 10    
        ggoSpread.SSSetEdit C_Co_cd, "회사코드", 12, , , 10           
        ggoSpread.SSSetEdit C_Co_cd_nm, "회사명", 20    
        ggoSpread.SSSetEdit C_Logon_gp, "로그온그룹", 17, , ,30
        ggoSpread.SSSetEdit C_Logon_gp_nm, "로그온 그룹명", 16            '12-15 
        ggoSpread.SSSetDate C_Usr_valid_dt, "사용자유효일", 21, 2, parent.gDateFormat    
        ggoSpread.SSSetEdit C_Interface_id, "인터페이스 ID", 14, , ,13    
        ggoSpread.SSSetDate C_Pwd_valid_dt, "암호유효기간", 30, 2, parent.gDateFormat    
        'ggoSpread.SSSetSplit2(2)
        .ReDraw = true

        Call SetSpreadLock
        
        Call ggoSpread.MakePairsColumn(C_Usr_id,C_Usr_nm,"1")
        Call ggoSpread.MakePairsColumn(C_Logon_gp,C_Logon_gp_nm,"1")
        Call ggoSpread.MakePairsColumn(C_Co_cd,C_Co_cd_nm,"1")

        Call ggoSpread.SSSetColHidden(C_Pwd,C_Pwd,True)
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
    ggoSpread.SSSetProtected        C_Usr_id, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected        C_Usr_nm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected        C_Usr_eng_nm, pvStartRow, pvEndRow    
    ggoSpread.SSSetProtected        C_Pwd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected        C_Co_cd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected        C_Co_cd_nm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected        C_Logon_gp, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected        C_Logon_gp_nm, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected        C_Usr_valid_dt, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected        C_Interface_id, pvStartRow, pvEndRow        
    ggoSpread.SSSetProtected        C_Pwd_valid_dt, pvStartRow, pvEndRow    
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
            
            C_Usr_id        =  iCurColumnPos(1)
            C_Usr_nm        =  iCurColumnPos(2)
            C_Usr_eng_nm    =  iCurColumnPos(3)
            C_Pwd            =  iCurColumnPos(4)
            C_Co_cd            =  iCurColumnPos(5)
            C_Co_cd_nm        =  iCurColumnPos(6)
            C_Logon_gp        =  iCurColumnPos(7)
            C_Logon_gp_nm    =  iCurColumnPos(8)
            C_Usr_valid_dt    =  iCurColumnPos(9)
            C_Interface_id    =  iCurColumnPos(10)
            C_Pwd_valid_dt    =  iCurColumnPos(11)

    End Select
End Sub

'=========================================================================================================
Sub InitComboBox()
    
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

    Call InitSpreadSheet  
                                                      
    Call InitVariables                                                      
    Call SetDefaultVal   

    frm1.txtUsrId.focus
    Set gActiveElement = document.activeElement
    Call SetToolbar("11000000000011")                                        

End Sub

'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub
'=========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

   frm1.vspdData.Row = Row
   frm1.vspdData.Col = Col

   ggoSpread.UpdateRow Row   

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
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If CheckRunningBizProcess = True Then
        Exit Sub
    End If

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		Call DisableToolBar(parent.TBC_QUERY)                                                       
		If DbQuery = False Then                                                       
			Call RestoreToolBar()
			Exit Sub
		End If
    End if
    
End Sub
'=========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Err.Clear   
    
    FncQuery = False                                                        

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
Function FncPrint() 
    Call Parent.FncPrint()                                                   
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
    Call Parent.FncExport(Parent.C_MULTI)                                                   
End Function
'=========================================================================================================
Function FncFind() 
    Call Parent.FncFind(Parent.C_MULTI, False)                                         
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

    On Error Resume Next
    Err.Clear                                                               
    
    DbQuery = False
    
    Call DisableToolBar(parent.TBC_QUERY)                                                
    Call LayerShowHide(1)                                                         
        
    With frm1

    If lgIntFlgMode = Parent.OPMD_UMODE Then
        strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001                            
        strVal = strVal & "&txtUsrId=" & .hUsrId.value                 
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
        strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001                            
        strVal = strVal & "&txtUsrId=" & Trim(.txtUsrId.value)                
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If    

    Call RunMyBizASP(MyBizASP, strVal)                                        
    
    End With
    
    If Err.number = 0 Then     
        DbQuery = True                                                                 
    End If
    
End Function
'=========================================================================================================
Function DbQueryOk()                                                        
    lgIntFlgMode = Parent.OPMD_UMODE                                                
    
    Call ggoOper.LockField(Document, "Q")                                    
    Call SetToolbar("11000000000111")                                        
    frm1.vspdData.Focus
End Function
'=========================================================================================================
'    Name : OpenUsrId()
'    Description : User PopUp
'=========================================================================================================
Function OpenUsrId(Byval strCode, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "사용자정보 팝업"                                     ' 팝업 명칭 
    arrParam(1) = "z_usr_mast_rec"                                          ' TABLE 명칭 
    arrParam(2) = strCode                                                   ' Code Condition
    arrParam(3) = ""                                                        ' Name Cindition
    arrParam(4) = ""                                                        ' Where Condition
    arrParam(5) = "사용자 ID"            
    
    arrField(0) = "Usr_id"                                                  ' Field명(0)
    arrField(1) = "Usr_nm"                                                  ' Field명(1)
    
    arrHeader(0) = "사용자"                                                ' Header명(0)
    arrHeader(1) = "사용자명"                                           ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",  Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetUsrId(arrRet, iWhere)        'return value setting
    End If    
	frm1.txtUsrId.focus
	Set gActiveElement = document.activeElement

End Function
'=========================================================================================================
'    Name : SetUsrId()
'    Description : User Master Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetUsrId(Byval arrRet, Byval iWhere)
    With frm1
        If iWhere = 0 Then
            .txtUsrId.value = arrRet(0)
            .txtUsrNm.value = arrRet(1)
        End If
    End With
End Function
'=========================================================================================================
Function JumpProgram()
    Dim strUsrId
    strUsrId = ""

    On Error Resume Next
    
    strUsrId = Trim(GetSpreadText(frm1.vspdData, C_Usr_id, frm1.vspdData.ActiveRow, "X", "X"))

    WriteCookie "Za001ma1_UsrId", Trim(strUsrId)
    PgmJump(JUMP_PGM_ID)    
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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>사용자 정보 조회</font></td>
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
                                </TR>
                                    <TD CLASS="TD5">사 용 자</TD>
                                    <TD CLASS="TD656">
                                        <INPUT TYPE=TEXT NAME="txtUsrId" SIZE=13 MAXLENGTH=3000 tag="11N" ALT="사용자" LANGUAGE=javascript ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUsrId" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUsrId frm1.txtUsrId.value,0">&nbsp;
                                        <INPUT TYPE=TEXT ID="txtUsrNm" NAME="arrCond" SIZE=20 tag="14"></TD>
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
                                <script language =javascript src='./js/za001ma1_I132440030_vspdData.js'></script></TD>
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
                    <TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpProgram">사용자 정보 관리</TD>
                    <TD WIDTH=10>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>    
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex=-1></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hUsrId" tag="24">
<DIV ID=ScriptDiv NAME=ScriptDiv></Div>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

