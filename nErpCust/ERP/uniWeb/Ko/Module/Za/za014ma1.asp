<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : User Management
*  3. Program ID           : za014ma1
*  4. Program Name         : User Master Record
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 2002.12.03
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : ParkSangHoon
* 10. Modifier (Last)      : 
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<%
   Dim xxModule
   Dim gUseSCMYN
   
   Set xxModule = Server.CreateObject("xModule.xCA0001")
   gUseSCMYN = Split(xxModule.ReadD(GetGlobalData("gCompany")),";")(0)

%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">        

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                            
<% 
   Dim StartDate
   StartDate = GetSvrDate()                      
%>   



Const BIZ_PGM_ID = "za014mb1.asp"                                                    
Const JUMP_PGM_ID_1 = "za003ma1"            
Const JUMP_PGM_ID_2 = "za004ma1"            

Dim lgStrPwdUpdate             ' Variable indicating whether password is updated or not.
Dim lgBlnUsrCopy               ' Variable is for Function FncCopy Mode
Dim lgBlnUsrNew                ' Variable is for Function FncNew Mode

Dim lgStrPrevUsrRoleId
Dim lgStrPrevOrgType
Dim lgStrPrevOccurDt
Dim lgStrPrevOrgCd

Dim lgStrPrvNext
Dim IsOpenPop                        
Dim ValidDate

Dim C_RoleId
Dim C_RoleIdPopup
Dim C_RoleNm
Dim C_CompstRoleType

Dim C_UseYn
Dim C_OrgType
Dim C_OrgTypePopup
Dim C_OrgTypeNm
Dim C_OrgCd
Dim C_OrgCdPopup
Dim C_OrgNm
Dim C_OccurDt
Dim C_OccurTime
Dim C_hOccurDt

<!-- #Include file="../../inc/lgvariables.inc" -->    

'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
    If pvSpdNo = "A" Then
        ' Grid 1(vspdData) - Operation     
        C_RoleId = 1                                                            
        C_RoleIdPopup = 2    
        C_RoleNm = 3
        C_CompstRoleType = 4
        
    ElseIf pvSpdNo = "B" Then
        ' Grid 1(vspdData1) - Operation         
        C_UseYn = 1
        C_OrgType = 2                                                                    
        C_OrgTypePopup = 3    
        C_OrgTypeNm = 4
        C_OrgCd = 5                                                            
        C_OrgCdPopup = 6    
        C_OrgNm = 7
        C_OccurDt = 8
        C_OccurTime = 9
        C_hOccurDt = 10
        
    Else
        C_RoleId = 1                                                            
        C_RoleIdPopup = 2    
        C_RoleNm = 3
        C_CompstRoleType = 4
        C_UseYn = 1
        C_OrgType = 2                                                                    
        C_OrgTypePopup = 3    
        C_OrgTypeNm = 4
        C_OrgCd = 5                                                            
        C_OrgCdPopup = 6    
        C_OrgNm = 7
        C_OccurDt = 8
        C_OccurTime = 9
        C_hOccurDt = 10    
    End If
End Sub
'=========================================================================================================
Sub GetValidDate(ValidDate)
   Dim strYear,strMonth,strDay   
   Call ExtractDateFrom("<%=StartDate%>",Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
   strYear =  CInt(strYear) + 1 
   ValidDate = UniConvYYYYMMDDToDate(Parent.gDateFormat,strYear,strMonth,strDay)
End Sub
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgBlnUsrCopy = False

    lgIntGrpCount = 0                           

    lgStrPrevUsrRoleId = ""                           
    lgStrPrevOrgType = ""                           
    lgStrPrevOrgCd = ""                                   
    '---- Coding part--------------------------------------------------------------------
    lgCurrentSpd   = "*"
    
End Sub
'=========================================================================================================
Sub SetDefaultVal()

    With frm1
        .txtUsrId1.focus 
        Set gActiveElement = document.activeElement
        .txtUsrId1.Value= ReadCookie("Za001ma1_UsrId")
        Call WriteCookie("Za001ma1_UsrId", "") 'Delete        
        Call GetValidDate(ValidDate)        
        .txtUsrValidDt.text = ValidDate        
        .txtCoCd.value = Parent.gCompany
        .cboUserKind.value = "U"
    End With
End Sub

'=========================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE","QA") %>
End Sub


'=========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
    
    With frm1
        If pvSpdNo = "A" Then
        
            ' Grid 1 - Operation Spread Setting        
            Call InitSpreadPosVariables(pvSpdNo)        
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)        
             
            .vspdData.ReDraw = False
            .vspdData.MaxCols = C_CompstRoleType + 1
            .vspdData.MaxRows = 0            

             Call GetSpreadColumnPos("A")
                    
            ggoSpread.SSSetEdit C_RoleId,    "Role ID",    40,,,20
            ggoSpread.SSSetButton C_RoleIdPopup        
            ggoSpread.SSSetEdit C_RoleNm,    "Role명",    40,,,30
            ggoSpread.SSSetEdit C_CompstRoleType,    "Role Type",    36,,,30                
            .vspdData.ReDraw = True
        
            Call ggoSpread.MakePairsColumn(C_RoleId,C_RoleIdPopup,"1")

            Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)
            
        ElseIf pvSpdNo = "B" Then
        
            ' Grid 2 - Operation Spread Setting                
            Call InitSpreadPosVariables(pvSpdNo)                
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)        
                 
            .vspdData1.ReDraw = False            
            .vspdData1.MaxCols =  C_hOccurDt  + 1                                                
            .vspdData1.MaxRows = 0            

            Call GetSpreadColumnPos("B")        
       
            ggoSpread.SSSetCheck C_UseYn, "적용여부", 10, 2, "", True    
            ggoSpread.SSSetEdit C_OrgType, "조직형태", 13
            ggoSpread.SSSetButton C_OrgTypePopup
            ggoSpread.SSSetEdit C_OrgTypeNm, "조직형태명", 25
            ggoSpread.SSSetEdit C_OrgCd, "조직코드", 22
            ggoSpread.SSSetButton C_OrgCdPopup
            ggoSpread.SSSetEdit C_OrgNm, "조직명", 27
            ggoSpread.SSSetDate C_OccurDt, "조직변경일", 15, 2, Parent.gDateFormat                      
            ggoSpread.SSSetEdit C_OccurTime, "시간", 15, 2
            ggoSpread.SSSetEdit C_hOccurDt,   "", 14, 2            
            'ggoSpread.SSSetSplit2(4)                    
            .vspdData1.ReDraw = True
        
            Call ggoSpread.MakePairsColumn(C_OrgType,C_OrgTypePopup,"1")
            Call ggoSpread.MakePairsColumn(C_OrgCd,C_OrgCdPopup,"1")
            Call ggoSpread.MakePairsColumn(C_OccurDt,C_OccurTime,"1")

            Call ggoSpread.SSSetColHidden(C_hOccurDt,C_hOccurDt,True)
            Call ggoSpread.SSSetColHidden(.vspdData1.MaxCols, .vspdData1.MaxCols, True)
        Else
            ' Grid 1 - Operation Spread Setting        
            Call InitSpreadPosVariables(pvSpdNo)        
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)        
             
            .vspdData.ReDraw = False
            .vspdData.MaxCols = C_CompstRoleType + 1
            .vspdData.MaxRows = 0            

             Call GetSpreadColumnPos("A")
                    
            ggoSpread.SSSetEdit C_RoleId,    "Role ID",    40,,,20
            ggoSpread.SSSetButton C_RoleIdPopup        
            ggoSpread.SSSetEdit C_RoleNm,    "Role명",    40,,,30
            ggoSpread.SSSetEdit C_CompstRoleType,    "Role Type",    36,,,30                
            .vspdData.ReDraw = True
        
            Call ggoSpread.MakePairsColumn(C_RoleId,C_RoleIdPopup,"1")

            Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)

            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)        
                 
            .vspdData1.ReDraw = False            
            .vspdData1.MaxCols =  C_hOccurDt  + 1                                                
            .vspdData1.MaxRows = 0            

            Call GetSpreadColumnPos("B")        
       
            ggoSpread.SSSetCheck C_UseYn, "적용여부", 10, 2, "", True    
            ggoSpread.SSSetEdit C_OrgType, "조직형태", 13
            ggoSpread.SSSetButton C_OrgTypePopup
            ggoSpread.SSSetEdit C_OrgTypeNm, "조직형태명", 25
            ggoSpread.SSSetEdit C_OrgCd, "조직코드", 22
            ggoSpread.SSSetButton C_OrgCdPopup
            ggoSpread.SSSetEdit C_OrgNm, "조직명", 27
            ggoSpread.SSSetDate C_OccurDt, "조직변경일", 15, 2, Parent.gDateFormat                      
            ggoSpread.SSSetEdit C_OccurTime, "시간", 15, 2
            ggoSpread.SSSetEdit C_hOccurDt,   "", 14, 2
            'ggoSpread.SSSetSplit2(4)                    
            .vspdData1.ReDraw = True
        
            Call ggoSpread.MakePairsColumn(C_OrgType,C_OrgTypePopup,"1")
            Call ggoSpread.MakePairsColumn(C_OrgCd,C_OrgCdPopup,"1")
            Call ggoSpread.MakePairsColumn(C_OccurDt,C_OccurTime,"1")

            Call ggoSpread.SSSetColHidden(C_hOccurDt,C_hOccurDt,True)
            Call ggoSpread.SSSetColHidden(.vspdData1.MaxCols, .vspdData1.MaxCols, True)
        
        End If
        
        Call SetSpreadLock                                 
    End With
End Sub

'=========================================================================================================
Sub SetSpreadLock()
    With frm1 
        ggoSpread.Source = .vspdData        
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_RoleId, -1, C_RoleIdPopup
        ggoSpread.spreadLock C_RoleNm,     -1, C_RoleNm
        ggoSpread.spreadLock C_CompstRoleType,     -1, C_CompstRoleType                        
        .vspdData.ReDraw = True
    
        ggoSpread.Source = .vspdData1            
        .vspdData1.ReDraw = False
        ggoSpread.SpreadLock C_OrgType, -1, C_OrgType
        ggoSpread.SpreadLock C_OrgTypePopup, -1, C_OrgTypePopup
        ggoSpread.SpreadLock C_OrgTypeNm, -1, C_OrgTypeNm
        ggoSpread.SpreadLock C_OrgNm, -1, C_OrgNm
        ggoSpread.SpreadLock C_OccurDt, -1, C_OccurDt
        ggoSpread.SpreadLock C_OccurTime, -1, C_OccurTime        
        .vspdData1.ReDraw = True
    End With
End Sub

'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        ggoSpread.Source = frm1.vspdData    
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired     C_RoleId, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected    C_RoleNm, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected    C_CompstRoleType,pvStartRow, pvEndRow
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
            
            C_RoleId          =  iCurColumnPos(1)
            C_RoleIdPopup     =  iCurColumnPos(2)
            C_RoleNm          =  iCurColumnPos(3)
            C_CompstRoleType  =  iCurColumnPos(4)
            
       Case "B"   
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
                
            C_UseYn           =  iCurColumnPos(1)
            C_OrgType         =  iCurColumnPos(2)
            C_OrgTypePopup    =  iCurColumnPos(3)
            C_OrgTypeNm       =  iCurColumnPos(4)
            C_OrgCd           =  iCurColumnPos(5)
            C_OrgCdPopup      =  iCurColumnPos(6)
            C_OrgNm           =  iCurColumnPos(7)
            C_OccurDt         =  iCurColumnPos(8)
            C_OccurTime       =  iCurColumnPos(9)
            C_hOccurDt        =  iCurColumnPos(10)

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

Sub InitComboBox()

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("Z0051", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    
    Call  SetCombo2(frm1.cboUserKind,lgF0, lgF1,Chr(11)) 
	
End Sub


'=========================================================================================================
Sub txtUsrValidDt_DblClick(Button)
    if Button = 1 then
        frm1.txtUsrValidDt.Action = 7
    End if
End Sub
'=========================================================================================================
Sub txtUsrId2_Change()
        lgBlnFlgChgValue = true    
End Sub
'=========================================================================================================
Sub txtUsrNm2_Change()
        lgBlnFlgChgValue = true    
End Sub
'=========================================================================================================
Sub txtUsrEngNm_Change()
    lgBlnFlgChgValue = true    
End Sub
'=========================================================================================================
Sub txtUsrValidDt_Change()            
    lgBlnFlgChgValue = true    
End Sub
'=========================================================================================================
Sub txtInterfaceId_Change()
    lgBlnFlgChgValue = true    
End Sub
'=========================================================================================================
Sub hPassword_Change()
    lgBlnFlgChgValue = true    
End Sub
'=========================================================================================================
Sub txtCoCd_Change()
    lgBlnFlgChgValue = true    
End Sub
'=========================================================================================================
Sub txtLogOnGrp_Change()
    lgBlnFlgChgValue = true    
End Sub
'=========================================================================================================
Sub txtUseYn_Change()
    lgBlnFlgChgValue = true    
End Sub
'=========================================================================================================
Sub txtUsrId1_OnClick()
    'Call ControlToolbarButton
End Sub
'=========================================================================================================
Sub txtUsrId2_OnClick()
    'Call ControlToolbarButton
End Sub
'=========================================================================================================
Sub txtUsrNm1_OnClick()
    'Call ControlToolbarButton
End Sub
'=========================================================================================================
Sub txtUsrNm2_OnClick()
    'Call ControlToolbarButton
End Sub
'=========================================================================================================
Sub txtUsrEngNm_OnClick()
    'Call ControlToolbarButton
End Sub
'=========================================================================================================
Sub txtUsrValidDt_OnClick()
    'Call ControlToolbarButton
End Sub
'=========================================================================================================
Sub txtInterfaceId_OnClick()
    'Call ControlToolbarButton
End Sub
'=========================================================================================================
Sub txtPassword_OnClick()
    'Call ControlToolbarButton
End Sub
'=========================================================================================================
Sub txtCoCd_OnClick()
    'Call ControlToolbarButton
End Sub
'=========================================================================================================
Sub txtLogOnGrp_OnClick()
    'Call ControlToolbarButton
End Sub
'=========================================================================================================
Sub txtUseYn_OnClick()
    'Call ControlToolbarButton
End Sub
'=========================================================================================================
Sub ControlToolbarButton()
    If lgIntFlgMode = Parent.OPMD_UMODE Then
        Call SetToolbar("1111110111111111")                                                
    Else
        If lgBlnUsrNew = True Then        
            Call SetToolbar("1110110100001111")                                                    
        ElseIf lgBlnUsrCopy = True Then    
            Call SetToolbar("1110111100001111")                                                    
        Else
            Call SetToolbar("1110000000001111")                                            
        End If
    End If
End Sub 

'==========================================================================================
'   Event Name : Radio OnClick()
'   Event Desc : Radio Button Click시 lgBlnFlgChgValue 처리 / Value
'==========================================================================================
Sub rdoUseYn1_onClick()
    lgBlnFlgChgValue = True        
End Sub
'=========================================================================================================
Sub rdoUseYn2_onClick()
    lgBlnFlgChgValue = True        
End Sub
'=========================================================================================================
Function JumpToLoginHistory()

    On Error Resume Next
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                  
        Call DisplayMsgBox("900002", "X", "X", "X")                    
        'Call MsgBox("조회를 먼저 하십시오.", vbInformation)         
        Exit Function
    End If

    WriteCookie "Za014ma1_UsrId", Trim(frm1.htxtUsrId.value)    
    PgmJump(JUMP_PGM_ID_1)
    
End Function
'=========================================================================================================
Function JumpToMessageHistory()

    On Error Resume Next
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                  
        Call DisplayMsgBox("900002", "X", "X", "X")                    
        'Call MsgBox("조회를 먼저 하십시오.", vbInformation)         
        Exit Function
    End If

    WriteCookie "Za014ma1_UsrId", Trim(frm1.htxtUsrId.value)    
    PgmJump(JUMP_PGM_ID_2)
    
End Function

'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                         
    Call InitComboBox      
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   

    Call InitSpreadSheet("*")    

    Call InitVariables  
    Call SetDefaultVal

    lgStrPwdUpdate = False
    lgBlnFlgChgValue = False

    '----------  Coding part  -------------------------------------------------------------
    If Trim(frm1.txtUsrId1.value) <> "" Then
    
       Call MainQuery()
    Else

       Call FncNew()
       'Call SetToolbar("1110000000001111")                                        
       frm1.cboUserKind.value = "U"
       frm1.txtUsrId1.focus 
    End if

End Sub

'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)


    If lgIntFlgMode <> Parent.OPMD_UMODE Then
        Call SetPopupMenuItemInf("1111011011")
    Else 
        Call SetPopupMenuItemInf("1101111111")
    End If
    
    'Call SetPopupMenuItemInf("1101111111")
    
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
    lgCurrentSpd = "A"
    If lgIntFlgMode = Parent.OPMD_UMODE Then
        Call SetToolbar("1111111111111111")                                            
    Else
        Call SetToolbar("1110111100001111")                                            
    End If        
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'=========================================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )

    Call SetPopupMenuItemInf("0001111111")
    
    gMouseClickStatus = "SP2C"   
    Set gActiveSpdSheet = frm1.vspdData1
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
    
    lgCurrentSpd = "B"
    If lgIntFlgMode = Parent.OPMD_UMODE Then
        Call SetToolbar("1111100111111111")                                                
    Else
        Call SetToolbar("1110100100001111")                                                
    End If        

End Sub

'=========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
'=========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
'=========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
    End If
End Sub
'=========================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP2C" Then
        gMouseClickStatus = "SP2CR"
    End If
End Sub

'=========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'=========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'=========================================================================================================
Sub vspdData_LostFocus()
End Sub

'=========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
   frm1.vspdData.Row = Row
   frm1.vspdData.Col = Col

   ggoSpread.Source = frm1.vspdData
   ggoSpread.UpdateRow Row
End Sub
'=========================================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row )

   frm1.vspdData1.Row = Row
   frm1.vspdData1.Col = Col
   
   ggoSpread.Source = frm1.vspdData1
   ggoSpread.UpdateRow Row
   
   'lgBlnFlgChgValue    =    true
End Sub

'=========================================================================================================
Sub vspdData_DblClick(ByVal Col , ByVal Row)
    ggoSpread.Source = frm1.vspdData
End Sub

'=========================================================================================================
Sub vspdData1_DblClick(ByVal Col , ByVal Row)
    ggoSpread.Source = frm1.vspdData1
End Sub
'=============================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
    With frm1
        ggoSpread.Source = frm1.vspdData    
        .vspdData.Row = Row
        If Row > 0 And Col = C_RoleIdPopup Then
            Call OpenUsrRoleId(Trim(GetSpreadText(.vspdData, C_RoleId, Row, "X", "X")),0)
        End If               
    End With

End Sub
'=========================================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
    Dim strTemp
    
    With frm1
        ggoSpread.Source = frm1.vspdData1    
        
        If Row > 0 And Col = C_OrgTypePopUp Then
            Call OpenOrgType(Trim(GetSpreadText(.vspdData1, C_OrgType, Row, "X", "X")), 1)                    
        ElseIf Row > 0 And Col = C_OrgCdPopUp Then            
            strTemp = GetSpreadText(.vspdData1, C_OrgType, Row, "X", "X")
            Call OpenOrgCd(GetSpreadText(.vspdData1, C_OrgCd, Row, "X", "X"), strTemp, 1)            
        End If    
    End With

End Sub
'=========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'=========================================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub
'=========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    lgCurrentSpd = "A"

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then    
        If lgStrPrevUsrRoleId <> "" Then    
            If CheckRunningBizProcess = True Then
                Exit Sub
            End If    
                
            Call DisableToolBar(Parent.TBC_QUERY)
            If DBQuery = False Then
                Call RestoreToolBar()
                Exit Sub
            End If
        End if
    End if        
    
End Sub
'=========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    lgCurrentSpd = "B"
    
    If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then    
        If lgStrPrevOrgType <> "" And lgStrPrevOrgCd <> "" Then        
            If CheckRunningBizProcess = True Then
                Exit Sub
            End If    
                
            Call DisableToolBar(Parent.TBC_QUERY)
            If DBQuery = False Then
                Call RestoreToolBar()
                Exit Sub
            End If
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
    Dim pvSpdNo
        
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    pvSpdNo = gActiveSpdSheet.id
    Call InitSpreadSheet(pvSpdNo)  

    If pvSpdNo = "A" Then
        ggoSpread.Source = frm1.vspdData
    Else
        ggoSpread.Source = frm1.vspdData1
    End If

    Call ggoSpread.ReOrderingSpreadData()                


End Sub
'=========================================================================================================
Sub ClearSpreadData()
    ggoSpread.Source = frm1.vspdData    
    Call ggoSpread.ClearSpreadData()
    ggoSpread.Source = frm1.vspdData1    
    Call ggoSpread.ClearSpreadData()
End Sub
'=========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim var1, var2
    
    FncQuery = False                                                            
    
    Err.Clear                                                               

    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData1
    var2 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Then        
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x") '☜ 바뀐부분 
        'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    
    
    If Not chkField(Document, "1") Then                                    
       Exit Function
    End If
    

    Call ggoOper.ClearField(Document, "2")                                        

    ggoSpread.Source = frm1.vspdData    
    Call ggoSpread.ClearSpreadData()

    Call InitVariables                                                            
    

    Call ggoOper.LockField(Document, "N")                                    

    If DbQuery = False Then
       Exit Function
    End If
       
    FncQuery = True                                                                    
           
End Function


'=========================================================================================================
Function FncNew() 
    On Error Resume next

    Dim IntRetCD 

    FncNew = False                                                                  
    lgBlnUsrNew = True
   
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x") 
        'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    


    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")                                           
    
    With frm1
        .txtUsrId1.focus 
        Set gActiveElement = document.activeElement
        .txtUsrValidDt.text = ValidDate        
        .txtCoCd.value = Parent.gCompany
    End With

    Call InitVariables                                                            
    Call SetToolbar("1110110100001111")                                        


    Call FncOrgInfQuery()
    FncNew = True                                                                    
End Function
'=========================================================================================================
Function FncOrgInfQuery()
    On Error Resume next
    Dim IntRetCD1
    Dim iLngRow
    Dim strData
    Dim lgF0
    Dim lgF1
    Dim lgF2
    Dim lgF3
    Dim lgF4
    Dim lgF5
    Dim lgF6
    Dim ArrTmpF0
    Dim ArrTmpF1
    Dim ArrTmpF2
    Dim ArrTmpF3
    Dim ArrTmpF4
    Dim ArrTmpF5
    Dim ArrTmpF6                                            
    Dim iColSep

    iColSep = parent.gColSep
    
    IntRetCD1= CommonQueryRs("Case ISNULL(USE_YN," & FilterVar("N", "''", "S") & " ) When " & FilterVar("Y", "''", "S") & "  Then 1 When " & FilterVar("N", "''", "S") & "  Then 0 End, B.MINOR_CD, B.MINOR_NM, ISNULL(A.ORG_CD,''), ISNULL(ORG_NM,''), OCCUR_DT,  CONVERT(CHAR(23),OCCUR_DT,21) ","Z_USR_ORG_MAST A, B_MINOR B "," B.MINOR_CD *= A.ORG_TYPE AND B.MAJOR_CD = " & FilterVar("Z0001", "''", "S") & " AND   A.USR_ID = '' AND A.USE_YN = " & FilterVar("Y", "''", "S") & "  ORDER BY B.MINOR_CD ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    

    ArrTmpF0 = split(lgF0,iColSep)    
    ArrTmpF1 = split(lgF1,iColSep)    
    ArrTmpF2 = split(lgF2,iColSep)    
    ArrTmpF3 = split(lgF3,iColSep)    
    ArrTmpF4 = split(lgF4,iColSep)    
    ArrTmpF5 = split(lgF5,iColSep)    
    ArrTmpF6 = split(lgF6,iColSep)    

    
    strData = ""

    For iLngRow = 0 To 10                    'frm1.vspdData1.MaxRows
        strData = strData & Chr(11) & UCase(ConvSPChars(ArrTmpF0(iLngRow)))
        strData = strData & Chr(11) & UCase(ConvSPChars(ArrTmpF1(iLngRow)))
        strData = strData & Chr(11) & " "      'PopUp
        strData = strData & Chr(11) & UCase(ConvSPChars(ArrTmpF2(iLngRow)))
        strData = strData & Chr(11) & UCase(ConvSPChars(ArrTmpF3(iLngRow)))
        strData = strData & Chr(11) & " "      'PopUp                
        strData = strData & Chr(11) & UCase(ConvSPChars(ArrTmpF4(iLngRow)))
        strData = strData & Chr(11) & UNIDateClientFormat(ArrTmpF5(iLngRow))
        strData = strData & Chr(11) & ArrTmpF5(iLngRow)
        strData = strData & Chr(11) & UCase(ConvSPChars(ArrTmpF6(iLngRow)))
        strData = strData & Chr(11) & iLngRow + 1
        strData = strData & Chr(11) & Chr(12)
    Next

    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SSShowData strData
        
End Function
'=========================================================================================================
Function SplitTime(Byval dtDateTime)

    If IsNull(dtDateTime)  Then
        SplitTime = ""
        Exit Function
    End If

    SplitTime = Right("0" & Hour(dtDateTime), 2) & ":" _
            & Right("0" & Minute(dtDateTime), 2) & ":" _
            & Right("0" & Second(dtDateTime), 2)
            
End Function


'=========================================================================================================
Function FncDelete() 

    Dim IntRetCD
    
    FncDelete = False                                                            

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                  
        Call DisplayMsgBox("900002", "X", "X", "X")                    
        'Call MsgBox("조회를 먼저 하십시오.", vbInformation)         
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"x","x")
    If IntRetCD = vbNo Then
        Exit Function
    End If
    

    If DbDelete = False Then                                                    
        Call RestoreToolBar()
        Exit Function
    End If

    FncDelete = True                                                            
    
End Function


'=========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim var1, var2
    
    FncSave = False                                                                    
    Err.Clear                                                                    


    If Not chkField(Document, "2") Then                                    
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData                              
    var1 = ggoSpread.SSCheckChange
    
    ggoSpread.Source = frm1.vspdData1                              
    var2 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False And var1 = False And var2 = False  Then        
        IntRetCD = DisplayMsgBox("900001","X","X","X")                        '⊙: Display Message(There is no changed data.)
        Exit Function        
    End If 


    If lgCurrentSpd = "A" Then    
        ggoSpread.Source = frm1.vspdData                              
        If Not ggoSpread.SSDefaultCheck Then                          
           Exit Function
        End If            
    ElseIf lgCurrentSpd = "B" Then        
        ggoSpread.Source = frm1.vspdData1                              
        If Not ggoSpread.SSDefaultCheck Then                          
           Exit Function
        End If            
    Else    
        ggoSpread.Source = frm1.vspdData                              
        If Not ggoSpread.SSDefaultCheck Then                          
           Exit Function
        End If        
        
        ggoSpread.Source = frm1.vspdData1                              
        If Not ggoSpread.SSDefaultCheck Then                          
           Exit Function
        End If            
    End If
            
    If DbSave = False Then
       Exit Function
    End If
    
    FncSave = True                                                                
    
End Function



'=========================================================================================================
Function FncCopy() 
   
    Dim IntRetCD
    
    Err.Clear                                                                        

    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO,"x","x")
        '데이타가 변경되었습니다. 계속 하시겠습니까?
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    FncCopy = False                                                                  
    lgBlnFlgChgValue  = True
    lgBlnUsrCopy      = True
        
    lgIntFlgMode      = Parent.OPMD_CMODE                                                    

    Call ggoOper.ClearField(Document, "1")                                  
    Call ggoOper.LockField(Document, "N")                                    

    frm1.txtUsrId2.value = ""
    frm1.txtUsrId1.focus
    
    frm1.txtUsrValidDt.text = ValidDate        

    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    
    FncCopy = True                                                                   
    Call SetToolbar("1110111100001111")                                        
	
	frm1.txtPassword.value = ""

End Function


'=========================================================================================================
Function FncCancel() 
    If lgCurrentSpd = "A" Or  lgCurrentSpd = "*" Then
        ggoSpread.Source = Frm1.vspdData    
        ggoSpread.EditUndo  
    Else
        ggoSpread.Source = Frm1.vspdData1
        ggoSpread.EditUndo  
    End If    
End Function

'=========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                              
    Err.Clear                                                                     
    
    FncInsertRow = False                                                                 

    If IsNumeric(Trim(pvRowCnt)) then
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
        ggoSpread.InsertRow .vspdData.ActiveRow,imRow
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
    On Error Resume Next                                                    
    Dim lDelRows

    frm1.vspdData.focus
    ggoSpread.Source = frm1.vspdData
    lDelRows = ggoSpread.DeleteRow

End Function


'=========================================================================================================
Function FncPrint() 
    ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
End Function

'=========================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    
    Dim strVal
    Dim IntRetCD
    
    FncPrev = False
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                 
        Call DisplayMsgBox("900002", "X", "X", "X")                    
        Exit Function
    End If

    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x") 
        'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    If Not chkField(Document, "1") Then                                    
       Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                                    
    strVal = strVal & "&txtUsrId=" & Trim(frm1.txtUsrId1.value)    
    strVal = strVal & "&lgStrPrvNext=" & "P"
    strVal = strVal & "&lgCurrentSpd=" & lgCurrentSpd        
            
    Call RunMyBizASP(MyBizASP, strVal)
    
    frm1.txtUsrId1.focus    
    Set gActiveElement = document.activeElement
    FncPrev = True
            
End Function
'=========================================================================================================
Sub EraseContents()
    Call ggoOper.ClearField(Document, "2")                                        
    Call InitVariables                                                                
End Sub

'=========================================================================================================
Function FncNext() 
    On Error Resume Next                                                    
    Dim strVal
    Dim IntRetCD
    
    FncNext = False
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                  
        Call DisplayMsgBox("900002", "X", "X", "X")                     
        Exit Function
    End If

    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x") 
        'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If

    If Not chkField(Document, "1") Then                                    
       Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                                    
    strVal = strVal & "&txtUsrId=" & Trim(frm1.txtUsrId1.value)    
    strVal = strVal & "&lgStrPrvNext=" & "N"
    strVal = strVal & "&lgCurrentSpd=" & lgCurrentSpd        
    
    Call RunMyBizASP(MyBizASP, strVal)
    
    frm1.txtUsrId1.focus    
    Set gActiveElement = document.activeElement
    FncNext = True
    
End Function


'=========================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLEMULTI)
End Function


'=========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLEMULTI, TRUE)
End Function

Function FncExit()
    Dim IntRetCD
    FncExit = False
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")   '☜ 바뀐부분 
        'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vb
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

    If LayerShowHide(1) = False Then
        Exit Function 
    End If

    Dim strVal
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003                                
    strVal = strVal & "&txtUsrId=" & Trim(frm1.htxtUsrId.value)

    Call RunMyBizASP(MyBizASP, strVal)                                        
    
    DbDelete = True                                                             

End Function

'=========================================================================================================
Function DbDeleteOk()                                                            
    Call MainNew()
End Function

'=========================================================================================================
Function DbQuery() 
    
    Err.Clear                                                               

    DbQuery = False                                                             

    If   LayerShowHide(1) = False Then
             Exit Function 
    End If

    Dim strVal

        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001                                    
        strVal = strVal & "&lgCurrentSpd=" & lgCurrentSpd
        
        If lgIntFlgMode = Parent.OPMD_UMODE Then
            strVal = strVal & "&txtUsrId=" & Trim(frm1.htxtUsrId.value)
        Else
            strVal = strVal & "&txtUsrId=" & Trim(frm1.txtUsrId1.value)        
        End If
                                                    
        If lgCurrentSpd = "A" Or lgCurrentSpd = "*" Then            
            If lgIntFlgMode = Parent.OPMD_UMODE Then
                strVal = strVal & "&lgStrPrevUsrRoleId=" & lgStrPrevUsrRoleId
                strVal = strVal & "&lgStrPrvNext=" & lgStrPrvNext                                                                
                strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
            Else
                strVal = strVal & "&lgStrPrevUsrRoleId=" & lgStrPrevUsrRoleId
                strVal = strVal & "&lgStrPrvNext=" & lgStrPrvNext                                                                                        
                strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
            End If    
        End If
        
        If lgCurrentSpd = "B" Or lgCurrentSpd = "*" Then            
            If lgIntFlgMode = Parent.OPMD_UMODE Then
                strVal = strVal & "&txtOrgType=" & Trim(frm1.hOrgType.value)
                strVal = strVal & "&txtOccurDt=" & Trim(frm1.hOccurDt.value)                                                                
                strVal = strVal & "&txtOrgCd=" & Trim(frm1.hOrgCd.value)
                strVal = strVal & "&lgStrPrevOrgType=" & lgStrPrevOrgType
                strVal = strVal & "&lgStrPrevOccurDt=" & lgStrPrevOccurDt
                strVal = strVal & "&lgStrPrevOrgCd=" & lgStrPrevOrgCd
                strVal = strVal & "&txtMaxRows1=" & frm1.vspdData1.MaxRows                
            Else
                strVal = strVal & "&lgStrPrevOrgType=" & lgStrPrevOrgType
                strVal = strVal & "&lgStrPrevOccurDt=" & lgStrPrevOccurDt
                strVal = strVal & "&lgStrPrevOrgCd=" & lgStrPrevOrgCd        
                strVal = strVal & "&txtMaxRows1=" & frm1.vspdData1.MaxRows
            End If        
        End If

        Call RunMyBizASP(MyBizASP, strVal)                                                                            
    DbQuery = True       
                                                             

End Function

'=========================================================================================================
Function DbQueryOk()                                                        
    

    lgIntFlgMode = Parent.OPMD_UMODE                                                
    lgBlnFlgChgValue = False
    lgBlnUsrCopy = False    
    Call ggoOper.LockField(Document, "Q")                                    
    Call SetToolbar("1111110111111111")                                        

    frm1.txtUsrId1.focus    
    Set gActiveElement = document.activeElement
    'Call SetSpreadColumnLock()
    
End Function
'=========================================================================================================
Sub SetSpreadColumnLock()
    Dim lRow
    
    With frm1 
        ggoSpread.Source = .vspdData1        
        .vspdData1.ReDraw = False        

        For lRow = 1 To .vspdData1.MaxRows
            If Trim(GetSpreadText(.vspdData1, C_OrgCd, lRow, "X", "X")) <> "" Then
                ggoSpread.SpreadLock C_OrgCd, lRow, C_OrgCdPopup, lRow
            End If
        Next
    End With
    
End Sub


'=========================================================================================================
Function DbSave()        
    Dim lRow        
    Dim lGrpCnt     
    Dim strVal, strDel    
    Dim IntRetCD 
    Dim iColSep, iRowSep
    iColSep = parent.gColSep
    iRowSep = parent.gRowSep
    
    On Error Resume next                                                   
    DbSave = False                                                          

    Call LayerShowHide(1)
    
    If lgBlnFlgChgValue = True Then
        '사용자정보관리 
        
        If CheckUsrId(frm1.txtUsrId2.value) = 1 Then
             IntRetCD = DisplayMsgBox("210104", "x", "x", "x")
           '사용자아이디에 숫자나 문자 이외의 데이터는 입력할 수 없습니다.
           Call LayerShowHide(0)
           Exit Function
        End if
        
        If frm1.rdoUseYn1.checked Then
            frm1.txtUseYn.value = "Y"
        Else
            frm1.txtUseYn.value = "N"
        End If            

        If lgStrPwdUpdate = True Then          '비밀번호가 update되었는지의 flag
            frm1.txtPwdUpdateOrNot.value = "Y"
        Else
            frm1.txtPwdUpdateOrNot.value = "N"
        End if

        frm1.txtBlnFlgChgValue.value = "True"
        frm1.txtFlgMode.value = lgIntFlgMode    
            
        If lgBlnUsrCopy = True Then
            frm1.txtBlnUsrCopy.value = "True"
        End If

    End If
    

    '사용자별 Role Assign & 사용자별 조직관리    
    With frm1
        '사용자별 Role Assign
        ggoSpread.Source = frm1.vspdData    
        
        .txtMode.value = Parent.UID_M0002
                    
        lGrpCnt = 1
        strVal = ""
        strDel = ""        


        For lRow = 1 To .vspdData.MaxRows            
            If lgBlnUsrCopy = True Then            
                    If GetSpreadText(.vspdData, 0, lRow, "X", "X") <> ggoSpread.DeleteFlag Then
                                                    strVal = strVal & "C"                       & iColSep
                                                    strVal = strVal & lRow                      & iColSep
                                                    strVal = strVal & Trim(.txtUsrId2.value)    & iColSep                                                     
                    								strVal = strVal & Trim(GetSpreadText(.vspdData, C_RoleId, lRow, "X", "X"))      & iRowSep
                    lGrpCnt = lGrpCnt + 1                        
                    End If
            Else            
                Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")
                      Case ggoSpread.InsertFlag                                      
                                                    strVal = strVal & "C"                       & iColSep
                                                    strVal = strVal & lRow                      & iColSep
                                                    strVal = strVal & Trim(.txtUsrId2.value)    & iColSep                                                     
                        							strVal = strVal & Trim(GetSpreadText(.vspdData, C_RoleId, lRow, "X", "X"))      & iRowSep
                        lGrpCnt = lGrpCnt + 1                            
                   Case ggoSpread.DeleteFlag        
                                                    strDel = strDel & "D"                       & iColSep
                                                    strDel = strDel & lRow                      & iColSep
                                                    strDel = strDel & Trim(.txtUsrId2.value)    & iColSep                                                     
                        							strDel = strDel & Trim(GetSpreadText(.vspdData, C_RoleId, lRow, "X", "X"))      & iRowSep
                        lGrpCnt = lGrpCnt + 1                   
                End Select
            End If
            
        Next
    
        .txtMaxRows.value = lGrpCnt-1
        .txtSpread.value = strDel & strVal

        'msgbox "사용자별 Role 할당: " & vbcrlf & "입력: " & vbTab & strVal & vbcrlf & "삭제: " & strDel

        
        '사용자별 조직관리        
        ggoSpread.Source = frm1.vspdData1    

        .txtMode.value = Parent.UID_M0002
                    
        lGrpCnt = 1
        strVal = ""
        strDel = ""        

        For lRow = 1 To .vspdData1.MaxRows

            If lgBlnUsrCopy = True Then
                    .vspdData1.Col = C_UseYn                    
                    If CInt(Trim(GetSpreadText(.vspdData1, 1, lRow, "X", "X"))) = 1 Then        '현재 소속되어 있는 조직이면 
						If CInt(Trim(GetSpreadText(.vspdData1, C_UseYn, lRow, "X", "X"))) = 1 Then	'khy					
							strVal = strVal & "C"                        & iColSep
							strVal = strVal & lRow                       & iColSep
							strVal = strVal & Trim(.txtUsrId2.value)     & iColSep
							If CInt(Trim(GetSpreadText(.vspdData1, C_UseYn, lRow, "X", "X"))) = 1 Then
							    strVal = strVal & "Y"                     & iColSep
							ElseIf CInt(Trim(GetSpreadText(.vspdData1, C_UseYn, lRow, "X", "X"))) = 0 Then 
							    strVal = strVal & "N"                     & iColSep
							End If
						
							strVal = strVal & Trim(GetSpreadText(.vspdData1, C_OrgType, lRow, "X", "X"))      & iColSep
							strVal = strVal & Trim(GetSpreadText(.vspdData1, C_OrgCd, lRow, "X", "X"))      & iColSep
							strVal = strVal & Trim(GetSpreadText(.vspdData1, C_OrgNm, lRow, "X", "X"))      & iRowSep
                       
							If Trim(GetSpreadText(.vspdData1, C_OrgCd, lRow, "X", "X")) = "" Then
							    IntRetCD = DisplayMsgBox("211124", "x", "x", "x")
							    Call LayerShowHide(0)
							    Exit Function
							End if
                        
						End If 
							lGrpCnt = lGrpCnt + 1						
                    End If
            Else   
				Select Case GetSpreadText(.vspdData1, 0, lRow, "X", "X")
				      Case ggoSpread.InsertFlag                                      
						If CInt(Trim(GetSpreadText(.vspdData1, C_UseYn, lRow, "X", "X"))) = 1 Then	'khy					
							strVal = strVal & "C"                        & iColSep
							strVal = strVal & lRow                       & iColSep
							strVal = strVal & Trim(.txtUsrId2.value)     & iColSep
							If CInt(Trim(GetSpreadText(.vspdData1, C_UseYn, lRow, "X", "X"))) = 1 Then
							    strVal = strVal & "Y"                     & iColSep
							ElseIf CInt(Trim(GetSpreadText(.vspdData1, C_UseYn, lRow, "X", "X"))) = 0 Then 
							    strVal = strVal & "N"                     & iColSep
							End If
								     
							strVal = strVal & Trim(GetSpreadText(.vspdData1, C_OrgType, lRow, "X", "X"))      & iColSep
							strVal = strVal & Trim(GetSpreadText(.vspdData1, C_OrgCd, lRow, "X", "X"))      & iColSep
							strVal = strVal & Trim(GetSpreadText(.vspdData1, C_OrgNm, lRow, "X", "X"))      & iRowSep

							If Trim(GetSpreadText(.vspdData1, C_OrgCd, lRow, "X", "X")) = "" Then
							    IntRetCD = DisplayMsgBox("211124", "x", "x", "x")
							    Call LayerShowHide(0)
							    Exit Function
							End if
					      End If
				        lGrpCnt = lGrpCnt + 1

					            
				   Case ggoSpread.UpdateFlag 
						If Trim(GetSpreadText(.vspdData1, C_OrgCd, lRow, "X", "X")) <>"" and Trim(GetSpreadText(.vspdData1, C_OrgNm, lRow, "X", "X")) <>"" Then	'khy					
							strVal = strVal & "U"                        & iColSep
							strVal = strVal & lRow                       & iColSep
							strVal = strVal & Trim(.txtUsrId2.value)     & iColSep
							If CInt(Trim(GetSpreadText(.vspdData1, C_UseYn, lRow, "X", "X"))) = 1 Then
							    strVal = strVal & "Y"                     & iColSep
							ElseIf CInt(Trim(GetSpreadText(.vspdData1, C_UseYn, lRow, "X", "X"))) = 0 Then 
							    strVal = strVal & "N"                     & iColSep
							End If
							strVal = strVal & Trim(GetSpreadText(.vspdData1, C_OrgType, lRow, "X", "X"))      & iColSep
							strVal = strVal & Trim(GetSpreadText(.vspdData1, C_OrgCd, lRow, "X", "X"))      & iColSep
							strVal = strVal & Trim(GetSpreadText(.vspdData1, C_OrgNm, lRow, "X", "X"))      & iColSep
							strVal = strVal & Trim(GetSpreadText(.vspdData1, C_hOccurDt, lRow, "X", "X"))      & iRowSep

							If Trim(GetSpreadText(.vspdData1, C_OrgCd, lRow, "X", "X")) = "" Then
							     IntRetCD = DisplayMsgBox("211124", "x", "x", "x")
							     Call LayerShowHide(0)
							     Exit Function
							End if
						  End If
				        lGrpCnt = lGrpCnt + 1                                                
				End Select
			
            End If

        Next
    
        .txtMaxRows1.value = lGrpCnt-1
        .txtSpread1.value = strDel & strVal        

        Call ExecMyBizASP(frm1, BIZ_PGM_ID)                                                

    End With
    
    DbSave = True   
    
End Function


'=========================================================================================================
Function DbSaveOk()                                                                    

    ggoSpread.Source = frm1.vspdData
    ggoSpread.SSDeleteFlag 1
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SSDeleteFlag 1
    
    frm1.txtUsrId1.value = frm1.txtUsrId2.value
    frm1.txtUsrNm1.value = frm1.txtUsrNm2.value    

    Call InitVariables
    Call MainQuery()

End Function

'========================================================================================
' Function Name : CheckUsrId
' Function Desc : Check if the data is the numeric or character or not
'========================================================================================
Function CheckUsrId(ByVal strNum) 
  Dim Ret
  Dim intlen, intCnt, intAsc

  intlen = len(strNum)

  For intCnt = 1 To intlen

      intAsc = asc(mid(strNum, intCnt, 1))

      If intAsc < 48 Or (intAsc > 57 And intAsc < 65) Or (intAsc > 90 And intAsc < 97) Or (intAsc > 122) then
         CheckUsrId = 1
         Exit function
      End if
  next
  
  CheckUsrId = 0
  
End Function

'=========================================================================================================
'    Name : OpenPWD()
'    Description : 
'=========================================================================================================
Sub OpenPWD(usrId)
    Dim arrRet

    arrRet = window.showModalDialog("../../ComAsp/PWD.asp?txtFlag=P&DisCancel=0&skipSave=1&txtUsr=" & usrID , Array(window.parent), _
        "dialogWidth=290px; dialogHeight=160px; center: Yes; help: No; resizable: No; status: No;")
    If arrRet(0) = "" Then
	Else
		Call SetPWD(arrRet) 
	End If
	'frm1.txtPassword.focus
	'Set gActiveElement = document.activeElement
End Sub

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
	frm1.txtUsrId1.focus
	Set gActiveElement = document.activeElement

End Function
'=========================================================================================================
Function OpenCompany(Byval strCode, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True
   
    arrParam(0) = "회사 정보 팝업"                                ' 팝업 명칭 
    arrParam(1) = "b_company"                                         ' TABLE 명칭 
    arrParam(2) = strCode                                             ' Code Condition
    arrParam(3) = ""                                                  ' Name Cindition
    arrParam(4) = ""                                                  ' Where Condition
    arrParam(5) = "회    사"            
    
    arrField(0) = "co_cd"                                             ' Field명(0)
    arrField(1) = "co_nm"                                              ' Field명(1)
    
    arrHeader(0) = "회사코드"                                     ' Header명(0)
    arrHeader(1) = "회사명"                                          ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetCompany(arrRet, iWhere)        'return value setting
    End If    
	frm1.txtCoCd.focus
	Set gActiveElement = document.activeElement

End Function
'=========================================================================================================
Function OpenLogonGrp(Byval strCode, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "로그온 그룹 정보 팝업"                 ' 팝업 명칭 
    arrParam(1) = "z_logon_gp"                                ' TABLE 명칭 
    arrParam(2) = strCode                                     ' Code Condition
    arrParam(3) = ""                                          ' Name Cindition
    arrParam(4) = ""                                          ' Where Condition
    arrParam(5) = "로그온 그룹"            

    arrField(0) = "logon_gp"                                  ' Field명(0)
    arrField(1) = "logon_gp_nm"                                  ' Field명(1)
    
    arrHeader(0) = "로그온 그룹"                          ' Header명(0)
    arrHeader(1) = "로그온 그룹명"                        ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        Call SetLogonGrp(arrRet, iWhere)        'return value setting
    End If    
	frm1.txtLogOnGrp.focus
	Set gActiveElement = document.activeElement

End Function
'=========================================================================================================
Function OpenUsrRoleId(Byval strCode, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "사용자 Role 정보 팝업"                 ' 팝업 명칭 
    arrParam(1) = "z_usr_role a"                                ' TABLE 명칭 
    arrParam(2) = strCode                                     ' Code Condition
    arrParam(3) = ""                                          ' Name Cindition
    arrParam(4) = "not exists ( select 1 from Z_USR_MAST_REC_USR_ROLE_ASSO b "
    arrParam(4) = arrParam(4) & " where b.usr_role_id = a.usr_role_id and b.usr_id =  " & FilterVar(frm1.htxtUsrId.value, "''", "S") & " ) "                                          ' Where Condition
    arrParam(5) = "Role ID"            

    arrField(0) = "a.usr_role_id"                               ' Field명(0)
    arrField(1) = "a.usr_role_nm"                                  ' Field명(1)
    arrField(2) = "case compst_role_type when " & FilterVar("1", "''", "S") & "  then " & FilterVar("Composite Role", "''", "S") & " else " & FilterVar("Menu Role", "''", "S") & " end"                          ' Field명(2)
    
    arrHeader(0) = "Role ID"                              ' Header명(0)
    arrHeader(1) = "Role명"                               ' Header명(1)
    arrHeader(2) = "Role Type"                            ' Header명(2)
        
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=670px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetRoleId(arrRet, iWhere)        'return value setting
    End If    

End Function
'=========================================================================================================
Function OpenUsrOrgHistory()

    Dim arrRet, strUsr
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                  
        Call DisplayMsgBox("900002", "X", "X", "X")                    
        'Call MsgBox("조회를 먼저 하십시오.", vbInformation)         
        Exit Function
    Else
        strUsr = frm1.htxtUsrId.value     
    End If
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZA014RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZA014RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
                
    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,strUsr, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False
    
    If arrRet = "" Then
        Exit Function
    End If
    
End Function
'=========================================================================================================
Function OpenAuthGenList()

    Dim arrRet, strUsr
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                  
        Call DisplayMsgBox("900002", "X", "X", "X")                    
        'Call MsgBox("조회를 먼저 하십시오.", vbInformation)         
        Exit Function
    Else
        strUsr = frm1.htxtUsrId.value     
    End If

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZA014RA2")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZA014RA2", "x")
        IsOpenPop = False
        Exit Function
    End If
    
    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,strUsr, arrField, arrHeader), _
             "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    IsOpenPop = False
    
    If arrRet = "" Then
        Exit Function
    End If
    
End Function
'=========================================================================================================
Function SetRoleId(Byval arrRet, Byval iWhere)
	Dim nActiveRow
    With frm1
    	nActiveRow = .vspdData.ActiveRow
        If iWhere = 0 Then
        	.vspdData.SetText C_RoleId, nActiveRow, arrRet(0)
        	.vspdData.SetText C_RoleNm, nActiveRow, arrRet(1)
        	.vspdData.SetText C_CompstRoleType, nActiveRow, arrRet(2)

            ggoSpread.Source = .vspdData
            ggoSpread.UpdateRow nActiveRow
        End If
    End With
End Function


'=========================================================================================================
'    Name : SetUsrId()
'    Description : User Master Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetUsrId(Byval arrRet, Byval iWhere)
    With frm1
        If iWhere = 0 Then
            .txtUsrId1.value = arrRet(0)
            .txtUsrNm1.value = arrRet(1)
        End If
    End With
End Function
'=========================================================================================================
Function SetCompany(Byval arrRet, Byval iWhere)
    With frm1
        lgBlnFlgChgValue = true            
        .txtCoCd.value   = arrRet(0)
        .txtCoCdNm.value = arrRet(1)
    End With
End Function

'=========================================================================================================
'    Name : SetLogon_gp()
'    Description : Logon_gp Popup에서 Return되는 값 setting
'=========================================================================================================

Function SetLogonGrp(Byval arrRet, Byval iWhere)
    With frm1
        lgBlnFlgChgValue = true        
        .txtLogOnGrp.value   = Trim(arrRet(0))
        .txtLogOnGrpNm.value = arrRet(1)
    End With
End Function
'=========================================================================================================
Function SetPWD(Byval arrRet)
    With frm1
        If arrRet(0) = "OK" Then 'spread    
            .hPassword.value   = arrRet(1)
            .txtPassword.value = arrRet(1)
            lgStrPwdUpdate = True    
            lgBlnFlgChgValue = True
        End If
    End With
End Function


'=========================================================================================================
'사용자별 조직관리 
Function OpenOrgCd(Byval strCode, ByVal strCond, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "조직 팝업"                    ' 팝업 명칭 
    arrParam(1) = "z_org_mast"                    ' TABLE 명칭 
    arrParam(2) = strCode                            ' Code Condition%>
    arrParam(3) = ""                                ' Name Cindition%>
    arrParam(4) = "org_type =  " & FilterVar(strCond , "''", "S") & ""    ' Where Condition%>
    arrParam(5) = "조직 코드"            
    
    arrField(0) = "org_cd"                            ' Field명(0)%>
    arrField(1) = "org_nm"                            ' Field명(1)%>
    
    arrHeader(0) = "조직코드"                    ' Header명(0)%>
    arrHeader(1) = "조직명"                        ' Header명(1)%>

    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",  Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetOrgCd(arrRet, iWhere)        'return value setting
    End If    

End Function

'=========================================================================================================
Function OpenOrgType(Byval strCode, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "조직 형태 팝업"                    ' 팝업 명칭 
    arrParam(1) = "b_minor"                         ' TABLE 명칭 
    arrParam(2) = strCode                            ' Code Condition%>
    arrParam(3) = ""                                ' Name Cindition%>
    arrParam(4) = "major_cd = " & FilterVar("z0001", "''", "S") & ""    ' Where Condition%>
    arrParam(5) = "조직 형태"            
    
    arrField(0) = "minor_cd"                            ' Field명(0)%>
    arrField(1) = "minor_nm"                            ' Field명(1)%>
    
    arrHeader(0) = "조직형태"                    ' Header명(0)%>
    arrHeader(1) = "조직형태명"                        ' Header명(1)%>
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",  Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        Call SetOrgType(arrRet, iWhere)        'return value setting
    End If    

End Function

'=========================================================================================================
'    Name : SetUsrId()
'    Description : UsrId Popup에서 Return되는 값 setting
'=========================================================================================================
Function SetOrgCd(Byval arrRet, Byval iWhere)
	Dim nActiveRow
    With frm1
    	nActiveRow = .vspdData1.ActiveRow
        If iWhere = 1 Then
        	.vspdData1.SetText C_OrgCd, nActiveRow, arrRet(0)
        	.vspdData1.SetText C_OrgNm, nActiveRow, arrRet(1)

            ggoSpread.Source = .vspdData1
            ggoSpread.UpdateRow nActiveRow

			.vspdData1.SetText C_UseYn, nActiveRow, 1
        End If
    End With
End Function
'=========================================================================================================
Function SetOrgType(Byval arrRet, Byval iWhere)
	Dim nActiveRow
    With frm1
    	nActiveRow = .vspdData1.ActiveRow
        If iWhere = 1 Then
        	.vspdData1.SetText C_OrgType, nActiveRow, arrRet(0)
        	.vspdData1.SetText C_OrgTypeNm, nActiveRow, arrRet(1)

            ggoSpread.Source = .vspdData1
            ggoSpread.UpdateRow nActiveRow
        End If
    End With
End Function
'======================================================================================================
' Function Name : ConvSPChars
' Function Desc : 문자열안의 "를 ""로 바꾼다.
'======================================================================================================
Function ConvSPChars(ByVal strVal)
    ConvSPChars = Replace(strVal, """", """""")
End Function 

Function txtPassword_onfocus()

	call OpenPWD (frm1.txtUsrId2.value)

	frm1.txtCoCd.focus 
	Set gActiveElement = document.ActiveElement   
	lgBlnFlgChgValue = true
	
End Function

Sub cboUserKind_OnChange()
	lgBlnFlgChgValue = true
End Sub

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
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>사용자 정보 관리</font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* align=right><A href="VBScript:OpenUsrOrgHistory()">사용자별 조직변경 내역</A>&nbsp;|&nbsp<A href="VBScript:OpenAuthGenList()">사용자별 권한생성 결과조회</A>&nbsp;</TD>                    
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
                        <FIELDSET CLASS="CLSFLD"><TABLE <%=LR_SPACE_TYPE_40%>>
                            <TR>
                                <TD CLASS="TD5" NOWRAP>사 용 자</TD>
                                <TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtUsrId1" SIZE=13 MAXLENGTH=13 tag="12XXXU" ALT="사용자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUsrId" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUsrId frm1.txtUsrId1.value,0">&nbsp;<INPUT TYPE=TEXT ID="txtUsrNm1" NAME="txtUsrNm1" size=30 tag="14"></TD>
                                <TD CLASS="TDT"></TD>
                                <TD CLASS="TD6"></TD>                                                                    
                            </TR>
                        </TABLE></FIELDSET>
                    </TD>
                </TR>
                <TR>
                    <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
                </TR>
                <TR>
                    <TD WIDTH=100% HEIGHT=* VALIGN=TOP>
                        <TABLE <%=LR_SPACE_TYPE_60%>>
                            <TR>
                                <TD CLASS=TD5 NOWRAP>사 용 자</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtUsrId2" SIZE=13 MAXLENGTH=13 tag="23XXX" ALT="사용자">&nbsp;<INPUT TYPE=TEXT ID="txtUsrNm2" NAME="txtUsrNm2" MAXLENGTH="30" SIZE=20 tag="22" ALT=사용자명>
                                <TD CLASS=TD5 NOWRAP>사용자명(영문)</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtUsrEngNm" ALT="사용자명(영문)" TYPE="Text" MAXLENGTH="50" SIZE="30"  tag="21"></TD>
                            </TR>                        
                            <TR>
                                <TD CLASS=TD5 NOWRAP>사용자유효일</TD>
                                <TD CLASS=TD6 NOWRAP>
                                    <SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtUsrValidDt" CLASS=FPDTYYYYMMDD tag="22" Title="FPDATETIME" ALT=사용자유효일></OBJECT>                                    ');</SCRIPT>
                                <TD CLASS=TD5 NOWRAP>인터페이스 ID</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtInterfaceId" ALT="InterfaceID" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="21XXX"></TD>
                                </TD>
                            </TR>
                            <TR>
                                <TD CLASS=TD5 NOWRAP>비밀번호</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT Type = Password NAME="txtPassword" ALT="비밀번호" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXX"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPassword" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPWD frm1.txtUsrId2.value"></TD>
                                <TD CLASS=TD5 NOWRAP>사용유무</TD>
                                <TD CLASS=TD6 NOWRAP>
                                    <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoUseYn" TAG="21X" VALUE="Y" CHECKED ID="rdoUseYn1"><LABEL FOR="rdoUseYn1">사용함</LABEL>&nbsp;&nbsp;&nbsp;
                                    <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoUseYn" TAG="21X" VALUE="N" ID="rdoUseYn2"><LABEL FOR="rdoUseYn2">사용안함</LABEL>            
                                </TD>                                        
                            </TR>
                            <TR>
                                <TD CLASS=TD5 NOWRAP>회사코드</TD>                                
                                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtCoCd" ALT="회사코드" TYPE="Text" MAXLENGTH="10" SIZE=10  tag="22XXX"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCoCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCompany frm1.txtCoCd.value,0">&nbsp;<INPUT NAME="txtCoCdNm" TYPE="Text" MAXLENGTH="50" SIZE=20 tag="24"></TD>
                                <TD CLASS=TD5 NOWRAP>로그온그룹</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtLogOnGrp" ALT="로그온그룹" TYPE="Text" MAXLENGTH="30" SIZE=10  tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLogOnGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenLogonGrp frm1.txtLogOnGrp.value, 0">&nbsp;<INPUT NAME="txtLogOnGrpNm" TYPE="Text" MAXLENGTH="50" SIZE=20 tag="24"></TD>
                            </TR>
                            <TR>
                                <TD CLASS=TD5 NOWRAP>Kind</TD>                                
                                <TD CLASS=TD6 NOWRAP><SELECT NAME="cboUserKind" ALT="Kind" tag="2<%=Replace(Replace(gUseSCMYN,"N",4),"Y",2)%>X"></SELECT>
                                </TD>
                                <TD CLASS=TD5 NOWRAP>E-Mail</TD>
                                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmail" ALT="Email" TYPE="Text" MAXLENGTH="50" SIZE="30"  tag="21"></TD>
                            </TR>
                            
                            <TR>
                                <TD HEIGHT="35%" WIDTH="100%" COLSPAN=4>
                                    <SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData ID = A WIDTH=100% HEIGHT=100% TAG="23" id=vaSpread TITLE="SPREAD"><PARAM NAME="MaxRows" Value=0><PARAM NAME="MaxCols" Value=0><PARAM NAME="ReDraw" VALUE=0></OBJECT>');</SCRIPT>
                                </TD>
                            </TR>
                            <TR>
                                <TD HEIGHT="65%" WIDTH="100%" COLSPAN=10>
                                    <SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 ID = B WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>                                    
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
                    <TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpToLoginHistory">로그인 내역 관리&nbsp;|&nbsp<A href="vbscript:JumpToMessageHistory">메세지 내역 관리&nbsp;</TD>
                    <TD WIDTH=10>&nbsp;</TD>                    
                </TR>
            </TABLE>
        </TD>
    </TR>    
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex=-1></IFRAME>
        </TD>
    </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows1" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpread1" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPwdUpdateOrNot" tag="24">    
<INPUT TYPE=HIDDEN NAME="txtBlnFlgChgValue" tag="24">   
<INPUT TYPE=HIDDEN NAME="txtBlnUsrCopy" tag="24">   
<INPUT TYPE=HIDDEN NAME="txtUseYn" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtUsrId" tag="24">
<INPUT TYPE=HIDDEN NAME="hPassword" tag="24">
<INPUT TYPE=HIDDEN NAME="hUsrIdValidDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hOrgType" tag="24">
<INPUT TYPE=HIDDEN NAME="hOccurDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hOrgCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

