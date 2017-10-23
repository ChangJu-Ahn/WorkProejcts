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


<!-- #Include file="../../inc/IncServer.asp" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">        

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Event.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Operation.vbs"></SCRIPT>


<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/adoQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/lgVariables.inc"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit    


Const BIZ_PGM_ID = "zc005mb1.asp"

Const C_LangCD = 1
Const C_FrmID = 2
Const C_FrmNm = 3
Const C_UseYN = 4
Const C_MnuID = 5
Const C_MnuPopUp = 6
Const C_MnuNm = 7

Dim lgIntGrpCount
Dim lgIntFlgMode

Dim lgStrPrevKey
Dim lgStrPrevKey2

Dim IsOpenPop
Dim lgSortKey          
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = OPMD_CMODE
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
End Sub
'=========================================================================================================
Sub InitSpreadSheet()

    With frm1.vspdData
    
        .MaxCols = C_MnuNm +1
        .Col = .MaxCols
        .ColHidden = True
    
        .MaxRows = 0
        ggoSpread.Source = frm1.vspdData

        .ReDraw = false
    
        ggoSpread.Spreadinit

        ggoSpread.SSSetCombo C_LangCD, "언어코드", 10, 2, False
        ggoSpread.SSSetEdit    C_FrmID, "폼ID", 15,,,15,2
        ggoSpread.SSSetEdit C_FrmNm, "폼명", 35,,,40
        ggoSpread.SSSetCheck C_UseYN, "사용여부", 10,,,true
        ggoSpread.SSSetEdit C_MnuID, "메뉴ID", 15,,,15,2
        ggoSpread.SSSetButton C_MnuPopUp
        ggoSpread.SSSetEdit C_MnuNm, "메뉴명", 30,,,40
        
        ggoSpread.SSSetSplit(3)        

        .ReDraw = true

        Call SetSpreadLock 
    
    End With
    
End Sub
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_LangCD, -1, C_FrmID
    ggoSpread.SSSetRequired C_FrmNm, -1, C_FrmNm
    ggoSpread.SpreadunLock C_UseYN, -1, C_UseYN
    ggoSpread.SpreadLock C_MnuID, -1, C_MnuNm
    .vspdData.ReDraw = True

    End With
End Sub

'=========================================================================================================
Sub SetSpreadColor(ByVal lRow)
    With frm1
    
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired        C_LangCD, lRow, lRow
        ggoSpread.SSSetRequired        C_FrmID, lRow, lRow
        ggoSpread.SSSetRequired        C_FrmNm, lRow, lRow
        ggoSpread.SSSetRequired        C_MnuID, lRow, lRow
        ggoSpread.SSSetProtected    C_MnuNm, lRow, lRow
        .vspdData.ReDraw = True
    
    End With
End Sub
'=========================================================================================================
Sub InitComboBox()

    Dim strCboData
    Dim IntRetCD
    Dim lgF0
    
    Call SetCombo(frm1.cboUseYN, "1", "사용")
    Call SetCombo(frm1.cboUseYN, "0", "미사용")    
     
    ggoSpread.Source = frm1.vspdData
    
    IntRetCD = CommonQueryRs("lang_cd","b_language","LANG_CD >= ''",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = Replace(lgF0, Chr(11), vbTab)    
    lgF0 = Replace(lgF0, " ","")
    ggoSpread.SetCombo lgF0, C_LangCD
    
    ggoSpread.SetCombo lgF0, C_LangCd

End Sub
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
    
    frm1.txtLangCd.focus 
	Set gActiveElement = document.activeElement
End Function
'=========================================================================================================
Function OpenFrmInfo()

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "폼 팝업"
    arrParam(1) = "Z_LANG_CO_REF_FRM"
    arrParam(2) = Trim(frm1.txtFrmID.Value)
    arrParam(3) = ""
    arrParam(4) = "LANG_CD =  " & FilterVar(gLang, "''", "S") & ""
    arrParam(5) = "폼ID"
    
    arrField(0) = "FRM_ID"
    arrField(1) = "FRM_NM"
    
    arrHeader(0) = "폼ID"
    arrHeader(1) = "폼명"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    frm1.txtFrmID.focus 
    
    If arrRet(0) = "" Then
    Else
        Call SetFrmInfo(arrRet)
    End If    
	frm1.txtFrmID.focus
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
    arrParam(4) = "LANG_CD = " & FilterVar(gLang, "''", "S") & " AND MNU_TYPE = " & FilterVar("P", "''", "S") & " "
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
Function SetLangCD(Byval arrRet)
    frm1.txtLangCD.Value    = arrRet(0)
End Function
'=========================================================================================================
Function SetFrmInfo(Byval arrRet)
    frm1.txtFrmID.Value    = arrRet(0)
    frm1.txtFrmNm.Value    = arrRet(1)
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
'=========================================================================================================
Sub Form_Load()

    Call ggoOper.LockField(Document, "N")
    
    Call InitSpreadSheet
    Call InitVariables

    Call InitComboBox
    Call SetDefaultVal
    Call SetToolbar("11001111001111")

    frm1.txtLangCd.Value = gLang    
    frm1.txtLangCd.focus 
    
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
Sub vspdData_Click(ByVal Col, ByVal Row)

    gMouseClickStatus = "SPC"
    
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
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
    End If

End Sub
'=========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
   
    With frm1.vspdData 
		.Row = Row
	    ggoSpread.Source = frm1.vspdData
	   
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
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) _
    And Not(lgStrPrevKey = "" And lgStrPrevKey2 = "") Then
        Call DisableToolBar(TBC_QUERY)
        If DBQuery = False Then
            Call RestoreToolBar()
            Exit Sub
        End If 
    End if
    
End Sub

'=========================================================================================================
Function FncQuery()

    Dim IntRetCD 
    
    FncQuery = False
    
    Err.Clear


    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
          Exit Function
        End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
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
            nActiveRow = .ActiveRow
            SetSpreadColor ActiveRow
            .SetText C_FrmID, ActiveRow, ""

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
Function FncInsertRow() 
	Dim nActiveRow
    With frm1
	    .vspdData.focus
	    ggoSpread.Source = .vspdData
	    .vspdData.ReDraw = False
	    ggoSpread.InsertRow
	    nActiveRow = .vspdData.ActiveRow
	    .vspdData.SetText C_LangCD, nActiveRow, gLang
	    .vspdData.SetText C_UseYN, nActiveRow, "1"
	    .vspdData.ReDraw = True
	    SetSpreadColor nActiveRow
    End With
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
    Call parent.FncExport(C_MULTI)
End Function

'=========================================================================================================
Function FncFind() 
    Call parent.FncFind(C_MULTI, False)
End Function
'=========================================================================================================
Function FncSplitColumn()
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = 5
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
       iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
       Exit Function  
    End If   

    Frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE

    ggoSpread.Source = Frm1.vspdData

    ggoSpread.SSSetSplit(ACol)    

    Call SetActiveCell(Frm1.vspdData,ACol,ARow,"M","X","X")


    Frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
End Function
'=========================================================================================================
Function FncExit()
	Dim IntRetCD
    
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    FncExit = True
End Function

'=========================================================================================================
Function DbQuery() 

    DbQuery = False

    Call LayerShowHide(1)        
    
    Err.Clear

    Dim strVal
    
    With frm1

    If lgIntFlgMode = OPMD_UMODE Then
        strVal = BIZ_PGM_ID & "?txtMode=" & UID_M0001
        strVal = strVal & "&txtLangCd=" & Trim(.hLangCd.value)
        strVal = strVal & "&txtFrmID=" & Trim(.hFrmID.value)
        strVal = strVal & "&cboUseYN=" & Trim(.hUseYN.value)
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2 
    Else
        strVal = BIZ_PGM_ID & "?txtMode=" & UID_M0001
        strVal = strVal & "&txtLangCd=" & Trim(.txtLangCd.value)
        strVal = strVal & "&txtFrmID=" & Trim(.txtFrmID.value)
        strVal = strVal & "&cboUseYN=" & Trim(.cboUseYN.value)
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
        strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2        

    End If    
 
    Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True
    

End Function
'=========================================================================================================
Function DbQueryOk()
    
    lgIntFlgMode = OPMD_UMODE
    
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

    
    DbSave = False    
    
    Call LayerShowHide(1)
    
    On Error Resume Next

    With frm1
        .txtMode.value = UID_M0002
        .txtUpdtUserId.value = gUsrID
        .txtInsrtUserId.value = gUsrID
        
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    

    For lRow = 1 To .vspdData.MaxRows
        Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")

            Case ggoSpread.InsertFlag
                
                strVal = strVal & "C" & gColSep & lRow & gColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_LangCD, lRow, "X", "X")) & gColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_FrmID, lRow, "X", "X")) & gColSep                
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_FrmNm, lRow, "X", "X")) & gColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_UseYN, lRow, "X", "X")) & gColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuID, lRow, "X", "X")) & gRowSep                

                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag

                strVal = strVal & "U" & gColSep & lRow & gColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_LangCD, lRow, "X", "X")) & gColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_FrmID, lRow, "X", "X")) & gColSep                
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_FrmNm, lRow, "X", "X")) & gColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_UseYN, lRow, "X", "X")) & gColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData, C_MnuID, lRow, "X", "X")) & gRowSep                
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag

                strDel = strDel & "D" & gColSep & lRow & gColSep
                strDel = strDel & Trim(GetSpreadText(.vspdData, C_LangCD, lRow, "X", "X")) & gColSep
                strDel = strDel & Trim(GetSpreadText(.vspdData, C_FrmID, lRow, "X", "X")) & gColSep                
                strDel = strDel & Trim(GetSpreadText(.vspdData, C_FrmNm, lRow, "X", "X")) & gColSep
                strDel = strDel & Trim(GetSpreadText(.vspdData, C_UseYN, lRow, "X", "X")) & gColSep
                strDel = strDel & Trim(GetSpreadText(.vspdData, C_MnuID, lRow, "X", "X")) & gRowSep
                
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

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->    
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
                                <td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
                                <td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>컴퍼니 폼 관리</font></td>
                                <td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
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
                        <TD CLASS="TD6"><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtLangCd" SIZE=10 MAXLENGTH=2 tag="12XXXU" ALT="언어 코드"><IMG SRC="../../image/btnPopup.gif" NAME="btnLangCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenLangCD()"></TD>
                        <TD CLASS="TD5">폼 ID</TD>
                        <TD CLASS="TD6"><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtFrmID" SIZE=15 MAXLENGTH=15 tag="11XXXU" ALT="메뉴 ID"><IMG SRC="../../image/btnPopup.gif"   NAME="btnFrmID" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenFrmInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtFrmNm" SIZE=40 tag="14"></TD>
                    </TR>
                    <TR>
                        <TD CLASS="TD5">사용여부</TD>
                        <TD CLASS="TD6"><SELECT NAME=cboUseYN tag="11X" STYLE="WIDTH: 82px;"><OPTION value=""></OPTION></SELECT></TD>
                        <TD CLASS="TD5"></TD>
                        <TD CLASS="TD6"></TD>                        
                    </TR>                    
                </TABLE></FIELDSET></TD>
            </TR>
            <TR>
                <TD WIDTH=100% HEIGHT=* valign=top><TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
                    <TR>
                        <TD HEIGHT="100%">
                        <script language =javascript src='./js/zc005ma1_I255229896_vspdData.js'></script></TD>
                    </TR></TABLE>
                </TD>
            </TR>
        </TABLE></TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="p2111mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hLangCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hFrmID" tag="24">
<INPUT TYPE=HIDDEN NAME="hUseYN" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
