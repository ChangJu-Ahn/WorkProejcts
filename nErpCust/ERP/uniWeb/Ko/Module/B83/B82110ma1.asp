<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         :  
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>


<Script Language="VBScript">

Option Explicit                                                '☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim IsOpenPop

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------

Const BIZ_PGM_QRY_ID  = "B82110mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_SAVE_ID = "B82110mb2.asp"						           '☆: Biz Logic ASP Name

'Const BIZ_PGM_JUMP_ID1 = ""                                '☆: Cookie에서 사용할 상수 
'Const BIZ_PGM_JUMP_ID2 = ""

Dim C_Trans         '선택 
Dim C_ReqNo         '의뢰번호 
Dim C_ReqId         '의뢰자 
Dim C_ReqIdNm       '의뢰자 
Dim C_ReqDt         '의뢰일자    
Dim C_ReqGbn        '의뢰구분 
Dim C_ReqGbnNm      '의뢰구분명 
Dim C_ItemCd        '품목코드 
Dim C_ItemNm        '품목명 
Dim C_Spec          '규격 
Dim C_EndDt         '완료일자 
Dim C_TransDt       '이관일자 
Dim C_ReqReson      '이관사유 
Dim C_Remark        '비고 

'--------------------------------------------------------------------------------------------------------

Dim strCHECK         '전체선택시 가져가는 변수 
 

Dim StartDate, EndDate

StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", Parent.gServerDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)
EndDate   = UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gDateFormat)
                 
'==========================================  InitComboBox()  ======================================
'    Name : InitComboBox()
'    Description : Init ComboBox
'==================================================================================================
Sub InitComboBox()
    Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1001' ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)    
    Call SetCombo2(frm1.cboItemAcct , lgF0, lgF1, Chr(11))
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'    Name : InitVariables()
'    Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False
    IsOpenPop = False
    strCHECK = 1
    frm1.btnRun.innerHTML = "전체선택"
End Sub 

'==========================================  2.2.1 SetDefaultVal()  ========================================
'    Name : SetDefaultVal()
'    Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
    frm1.txtDtFr.Text    = StartDate
    frm1.txtDtTo.Text    = EndDate
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'==========================================   CookiePage()  ======================================
'    Name : CookiePage()
'    Description : JUMP시 Load화면으로 조건부로 Value
'=================================================================================================
Function CookiePage(Byval Kubun)

    Const CookieSplit = 4877                        'Cookie Split String : CookiePage Function Use
    Dim strCookie
    Dim ii,jj,kk
    Dim iSeq
    Dim IntRetCD
    Dim strTemp
    Dim arrVal
             
    If Kubun = 1 Then                                'Jump로 화면을 이동할 경우 
        If  lgSaveRow <  1 Then
            IntRetCD = DisplayMsgBox("900002",Parent.VB_YES_NO,"X","X")
            Exit Function
        End If    
        
        Redim  lgMark(UBound(lgFieldNM)) 
        
        strCookie  = ""
        iSeq       = 0
        
        For ii = 0 to Parent.C_MaxSelList - 1 
            For jj = 0 to UBound(lgFieldNM) -1
                If lgPopUpR(ii,0) = lgFieldCD(jj) Then
                    iSeq = iSeq + 1
                    lgMark(jj) = "X"
                    strCookie = strCookie & "" & TRIM(LGFIELDNM(JJ)) & "" & Parent.gRowSep
                    frm1.vspdData.Row = lgSaveRow
                    frm1.vspdData.Col = iSeq
                    strCookie = strCookie & frm1.vspdData.Text & Parent.gRowSep
                
                    kk = CInt(lgNextSeq(jj)) 
                    If kk > 0 And kk <= UBound(lgFieldNM) Then 
                        lgMark(kk - 1) = "X"
                        iSeq = iSeq + 1
                        
                        strCookie = strCookie & "" & TRIM(LGFIELDNM(KK-1)) & "" & Parent.gRowSep
                        frm1.vspdData.Row = lgSaveRow
                        frm1.vspdData.Col = iSeq
                        strCookie = strCookie & frm1.vspdData.Text & Parent.gRowSep
                    End If    
                    jj =  UBound(lgFieldNM)  + 100
                End If    
            Next
        Next      
        
        WriteCookie CookieSplit , strCookie
        
        Call PgmJump(BIZ_PGM_JUMP_ID)
    
    ElseIf Kubun = 0 Then                            'Jump로 화면이 이동해 왔을경우 
        
        Call MainQuery()

        WriteCookie CookieSplit , ""

    End IF
End Function

          
'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()

    Call InitSpreadPosVariables()
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20030804", , Parent.gAllowDragDropSpread
  
    With frm1.vspdData
        .ReDraw = false
        .MaxCols = C_Remark + 1
        .MaxRows = 0

        Call ggoSpread.ClearSpreadData()
		Call GetSpreadColumnPos("A")
		Call AppendNumberPlace("6","16","0")
        
        ggoSpread.SSSetCheck C_Trans,       "선택",        8,     ,     ,    True
        ggoSpread.SSSetEdit  C_ReqNo,       "의뢰번호",   15
        ggoSpread.SSSetEdit  C_ReqId,       "의뢰자",     10
        ggoSpread.SSSetEdit  C_ReqIdNm,     "의뢰자",     10
        ggoSpread.SSSetDate  C_ReqDt,       "의뢰일자",   10, 2, Parent.gDateFormat 
        ggoSpread.SSSetEdit  C_ReqGbn,      "의뢰구분",   10 
        ggoSpread.SSSetEdit  C_ReqGbnNm,    "의뢰구분",   10 
        ggoSpread.SSSetEdit  C_ItemCd,      "품목코드",   15
        ggoSpread.SSSetEdit  C_ItemNm,      "품목명",     20
        ggoSpread.SSSetEdit  C_Spec,        "규격",       20
        ggoSpread.SSSetDate  C_EndDt,       "완료일자",   10, 2, Parent.gDateFormat 
        ggoSpread.SSSetDate  C_TransDt,     "이관일자",   10, 2, Parent.gDateFormat          
        ggoSpread.SSSetEdit  C_ReqReson,    "이관사유",   50
        ggoSpread.SSSetEdit  C_Remark,      "비고",       50
        
        Call ggoSpread.SSSetColHidden(C_ReqId, C_ReqId, True) 
        Call ggoSpread.SSSetColHidden(C_ReqGbn, C_ReqGbn, True) 
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
        
        'ggoSpread.SpreadLockWithOddEvenRowColor()
        ggoSpread.SSSetSplit2(2)  
        
        Call SetSpreadLock
        
        .ReDraw = true
        
    End With
End Sub

'==========================================  2.6.1 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()

    C_Trans    = 1      '선택 
    C_ReqNo    = 2      '의뢰번호 
    C_ReqId    = 3      '의뢰자 
    C_ReqIdNm  = 4      '의뢰자 
    C_ReqDt    = 5      '의뢰일자  
    C_ReqGbn   = 6      '의뢰구분 
    C_ReqGbnNm = 7      '의뢰구분명 
    C_ItemCd   = 8      '품목코드 
    C_ItemNm   = 9      '품목명 
    C_Spec     = 10      '규격 
    C_EndDt    = 11     '완료일자 
    C_TransDt  = 12     '이관일자 
    C_ReqReson = 13     '이관사유 
    C_Remark   = 14     '비고 
    
End Sub

'==========================================  2.6.2 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
     Dim iCurColumnPos
     
     Select Case Ucase(pvSpdNo)
     Case "A"
         ggoSpread.Source = frm1.vspdData 
         Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
         
        C_Trans     = iCurColumnPos(1)
        C_ReqNo     = iCurColumnPos(2)
        C_ReqId     = iCurColumnPos(3)
        C_ReqIdNm   = iCurColumnPos(4)
        C_ReqDt     = iCurColumnPos(5)
        C_ReqGbn    = iCurColumnPos(6)
        C_ReqGbnNm  = iCurColumnPos(7)
        C_ItemCd    = iCurColumnPos(8)
        C_ItemNm    = iCurColumnPos(9)
        C_Spec      = iCurColumnPos(10)
        C_EndDt     = iCurColumnPos(11)
        C_TransDt   = iCurColumnPos(12)
        C_ReqReson  = iCurColumnPos(13)
        C_Remark    = iCurColumnPos(14)
        
     End Select
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : 
'======================================================================================================
Sub SetSpreadLock()
    With frm1    
        ggoSpread.SSSetProtected    C_ReqNo, -1 
        ggoSpread.SSSetProtected    C_ReqIdNm,-1
        ggoSpread.SSSetProtected    C_ReqDt, -1 
        ggoSpread.SSSetProtected    C_ReqGbn,-1
        ggoSpread.SSSetProtected    C_ReqGbnNm,-1
        ggoSpread.SSSetProtected    C_ItemCd, -1    
        ggoSpread.SSSetProtected    C_ItemNm, -1 
        ggoSpread.SSSetProtected    C_Spec, -1    
        ggoSpread.SSSetProtected    C_EndDt,-1
        ggoSpread.SSSetProtected    C_TransDt,-1
        ggoSpread.SSSetProtected    C_ReqReson,-1
        ggoSpread.SSSetProtected    C_Remark, -1 
   
        ggoSpread.SSSetProtected    .vspdData.MaxCols, -1
   End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
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
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If           
       Next          
    End If   
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'    Name : Form_Load()
'    Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                                                        '⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    
    Call InitVariables                                                        '⊙: Initializes local global variables
    Call SetDefaultVal    
    Call InitComboBox()
    Call InitSpreadSheet()
    Call SetToolbar("11001000000111")                                        '⊙: 버튼 툴바 제어    
	frm1.txtreq_user.focus()
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode ) 
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then Exit Sub

     If Row <= 0 Then
         ggoSpread.Source = frm1.vspdData 
         If lgSortKey = 1 Then
             ggoSpread.SSSort Col                
             lgSortKey = 2
         Else
             ggoSpread.SSSort Col, lgSortKey        
             lgSortKey = 1
         End If
     End If
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
    End If
End Sub 

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()    '###그리드 컨버전 주의부분###
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
     
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
        If lgStrPrevKey <> "" Then
            If CheckRunningBizProcess = True Then Exit Sub
            
            Call DisableToolBar(Parent.TBC_QUERY)
            If DBQuery = False Then
                Call RestoreToolBar()
                Exit Sub
            End If
        End If
    End If    
End Sub

'==========================================================================================
'   Event Name : txtDtFr
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtDtFr_DblClick(Button)
    If Button = 1 Then
        frm1.txtDtFr.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtDtFr.Focus 
    End If
End Sub

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtDtTo_DblClick(Button)
    If Button = 1 Then
        frm1.txtDtTo.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtDtTo.Focus 
    End If
End Sub

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtDtFr_KeyPress(KeyAscii)
    If KeyAscii = 13 Then
        Call MainQuery()
    End If
End Function

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function txtDtTo_KeyPress(KeyAscii)
    If KeyAscii = 13 Then
        Call MainQuery()
    End If
End Function

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'    설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
        If IntRetCD = vbNo Then
           Exit Function
        End If   
    End If
    
     If Not chkField(Document, "1") Then                                    '⊙: This function check indispensable field
        Exit Function                       
     End If   
    
    If ValidDateCheck(frm1.txtDtFr, frm1.txtDtto)	=	False	Then Exit	Function
    
    ggoSpread.source = frm1.vspddata
    ggoSpread.ClearSpreadData 

    Call InitVariables                                                         '⊙: Initializes local global variables
    
    If DbQuery = False then    
       Exit Function
    End If   

    FncQuery = True                                                            '⊙: Processing is OK
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNew = False																  '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncNew = True                                                              '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function       

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    Dim iDx

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
     Frm1.vspdData.Row = frm1.vspdData.ActiveRow
     
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()    

End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD ,strSpCc2
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncSave = False                                                               '☜: Processing is NG
    
    ggoSpread.Source = frm1.vspdData
    strSpCc2 = ggoSpread.SSCheckChange
    
    If ggoSpread.SSCheckChange = False AND strSpCc2 = False Then                  '☜:match pointer
        IntRetCD = DisplayMsgBox("900001","x","x","x")                            '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                          '☜: Check contents area
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                        '☜: Query db data
       Exit Function
    End If

    If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint() 
   
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False
    
    Call parent.FncPrint()
    
    If Err.number = 0 Then	 
       FncPrint = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(Parent.C_MULTI)

    If Err.number = 0 Then	 
       FncExcel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	Call Parent.FncFind(Parent.C_MULTI, True)

    If Err.number = 0 Then	 
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		              '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If Err.number = 0 Then	 
       FncExit = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub


'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call InitComboBox      
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim strVal
    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
    Call LayerShowHide(1)
    
    With frm1
        If .rdoTrans1.Checked = True Then
           .htxtTransA.Value = "0"
        ElseIf .rdoTrans2.Checked = True Then
           .htxtTransA.Value = "1" 
        ElseIf .rdoTrans3.Checked = True Then
           .htxtTransA.Value = "2"
        ElseIf .rdoTrans4.Checked = True Then
           .htxtTransA.Value = "3"
        End If
        
        If .rdoTransT.Checked = True Then
           .htxtTransB.Value = "T"
        ElseIf .rdoTransC.Checked = True Then
           .htxtTransB.Value = "C" 
        End If
        strVal = ""
      
        strVal = BIZ_PGM_QRY_ID & "?txtMode="    & Parent.UID_M0001 &_
                               "&txtDtFr="       & Trim(.txtDtFr.Text) & _
                               "&txtDtTo="       & Trim(.txtDtTo.Text) & _
                               "&txtreq_user="      & Trim(.txtreq_user.value) & _
                               "&cboItemAcct="   & Trim(.cboItemAcct.value) & _
                               "&txtItem_Kind="   & Trim(.txtItem_Kind.value) & _
                               "&htxtTransA="    & Trim(.htxtTransA.value) & _
                               "&htxtTransB="    & Trim(.htxtTransB.value) & _
                               "&txtMaxRows="    & .vspdData.MaxRows & _
                               "&lgStrPrevKey="  & lgStrPrevKey                      '☜: Next key tag
                               
        Call RunMyBizASP(MyBizASP, strVal)
       
    End With
    
    DbQuery = True

End Function

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
		
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbSave = False                                                                '☜: Processing is NG

    Call DisableToolBar(Parent.TBC_SAVE)                                          '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message
		
    frm1.txtMode.value  = Parent.UID_M0002                                        '☜: Delete
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
    ggoSpread.Source = frm1.vspdData

    strVal  = ""
    lGrpCnt = 0

	With Frm1    
       
	   For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
          
           Select Case .vspdData.Text
 
               Case ggoSpread.UpdateFlag                                      '☜: Update
                    
                    .vspddata.col = C_Trans
                    
					IF .vspddata.value = "1" OR .vspddata.value = "Y" then
					                              strVal = strVal & lRow                    & Parent.gColSep
                                                  strVal = strVal & Trim(.htxtTransB.value) & Parent.gColSep
				       .vspdData.Col = C_ReqNo  : strVal = strVal & Trim(.vspdData.value)   & Parent.gColSep
					   .vspdData.Col = C_ReqGbn : strVal = strVal & Trim(.vspdData.value)   & Parent.gRowSep
					   					
					   lGrpCnt = lGrpCnt + 1
					   
					End If   
					
           End Select

       Next
	   
	   .txtMaxRows.value     = lGrpCnt
	   .txtSpread.value      = strVal

	End With
	   
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then	 
       DbSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   


End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
    Call SetToolbar("11001000000111")                           '⊙: 버튼 툴바 제어     
    Call InitVariables()
    Set gActiveElement = document.activeElement
End Function

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    Call InitData()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field
    
    If DbQuery() = False Then
       Call RestoreToolBar()
       Exit Sub
    End if
    
    Call InitVariables()
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement   
End Sub
	
	
'========================================================================================================
' Name : _Change
' Desc : developer describe this line
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)
       
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
          
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)	
    	
   If Col = C_Trans And frm1.vspdData.value = "0" Then
      ggoSpread.Source = frm1.vspdData
      ggoSpread.EditUndo frm1.vspdData.Row
   Else        
	  ggoSpread.Source = frm1.vspdData
      ggoSpread.UpdateRow Row
   End If
      
End Sub



'======================================================================================================
'        Name : OpenPopup()
'        Description : 
'=======================================================================================================
Function OpenPopup(Byval arPopUp)

        Dim arrRet
        Dim arrParam(7), arrField(8), arrHeader(8)
        Dim sItemAcct , sItemKind, sItemLvl1, sItemLvl2, sItemLvl3

        If IsOpenPop = True  Then  
           Exit Function
        End If   

        IsOpenPop = True
        
        Select Case arPopUp
               Case 1 '의뢰자 
                    
                    If UCase(frm1.txtreq_user.className) = UCase(parent.UCN_PROTECTED) Then 
                       IsOpenPop = False
                       Exit Function
                    End If
                                        
                    arrParam(0) = frm1.txtreq_user.Alt
                    arrParam(1) = "B_MINOR"
                    arrParam(2) = Trim(frm1.txtreq_user.Value)
                    arrParam(4) = "MAJOR_CD = 'Y1006' "
                    arrParam(5) = frm1.txtreq_user.Alt
       
                    arrField(0) = "MINOR_CD"
                    arrField(1) = "MINOR_NM"       
    
                    arrHeader(0) = frm1.txtreq_user.Alt
                    arrHeader(1) = frm1.txtreq_user_Nm.Alt
                    frm1.txtreq_user.focus()
                    
               Case 2 '품목구분 
               
                    If UCase(frm1.txtItem_Kind.className) = UCase(parent.UCN_PROTECTED) Then 
                       IsOpenPop = False
                       Exit Function
                    End If
                    
                    arrParam(0) = frm1.txtItem_Kind.Alt
                    arrParam(1) = "B_MINOR A, B_CIS_CONFIG B"
                    arrParam(2) = Trim(frm1.txtItem_Kind.value)
                    
                    sItemAcct = Trim(frm1.cboItemAcct.value)
                    
                    If sItemAcct = "" Then
                       arrParam(4) = "A.MAJOR_CD = 'Y1001' AND A.MINOR_CD = B.ITEM_KIND " 
                    Else
                       arrParam(1) = "B_MINOR A, B_CIS_CONFIG B"
                       arrParam(4) = "A.MAJOR_CD = 'Y1001' AND A.MINOR_CD = B.ITEM_KIND AND B.ITEM_ACCT = '" & sItemAcct & "'"
                    End If                   
                    
                    arrParam(5) = frm1.txtItem_Kind.Alt

                    arrField(0) = "A.MINOR_CD"
                    arrField(1) = "A.MINOR_NM"
    
                    arrHeader(0) = frm1.txtItem_Kind.Alt
                    arrHeader(1) = frm1.txtItem_Kind_Nm.Alt
                    frm1.txtItem_Kind.focus()
                    
               Case Else
                    IsOpenPop = False
                    Exit Function
      End Select
        
      arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
                "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

      IsOpenPop = False
                
      If arrRet(0) = "" Then
         Exit Function
      Else
         Call SubSetPopup(arrRet,arPopUp)
      End If        
        
End Function

'======================================================================================================
'        Name : SubSetPopup()
'        Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetPopup(Byval arrRet, Byval arPopUp)
    
    With Frm1
        Select Case arPopUp 
               Case 1 '의뢰자 
                    .txtreq_user.value   = arrRet(0)
                    .txtreq_user_Nm.value = arrRet(1) 
                    
               Case 2 '품목구분 
                    .txtItem_Kind.value   = arrRet(0)
                    .txtItem_Kind_Nm.value = arrRet(1)
                                  
               Case Else
                    Exit Sub
                    
              End Select              
              
        End With
        
End Sub

'========================================================================================
' Function Name : txtreq_user_OnChange
' Function Desc : 
'========================================================================================
Function txtreq_user_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtreq_user.value = "" Then
        frm1.txtreq_user_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" minor_nm "," b_minor "," major_cd='Y1006' and minor_cd="&filterVar(frm1.txtreq_user.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtreq_user_nm.value=""
        Else
            frm1.txtreq_user_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function

'========================================================================================
' Function Name : txtItem_kind_OnChange
' Function Desc : 
'========================================================================================
Function txtItem_Kind_OnChange()
    Dim iDx
    Dim IntRetCd
	
    If frm1.txtItem_Kind.value = "" Then
        frm1.txtItem_Kind_nm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" minor_nm "," b_minor "," major_cd='Y1001' and minor_cd="&filterVar(frm1.txtItem_Kind.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 frm1.txtItem_Kind_nm.value=""
        Else
            frm1.txtItem_Kind_nm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function


'======================================================================================================
'        Name : Check()
'        Description : 전체선택/취소 
'=======================================================================================================
Function Check()

    Dim i
        
    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Call DisplayMsgBox("900002", "x","x","x")
       Exit Function
    End If

    frm1.vspdData.Row = i
    
    With frm1.vspdData
    
        If strCHECK = 1 Then
        
           For i = 1 To .MaxRows
              .Row = i
              .col = C_Trans
              .value = "1"        
              ggoSpread.Source = frm1.vspdData
              ggoSpread.UpdateRow frm1.vspdData.Row
            Next
            
            frm1.btnRun.innerHTML = "선택취소"
            strCHECK = 2
             
        ElseIf strCHECK = 2 Then
        
            For i = 1 To .MaxRows            
               .Row = i             
               .col = C_Trans
               .value = "0"
               ggoSpread.Source = frm1.vspdData
               ggoSpread.EditUndo frm1.vspdData.Row
            Next
            
            frm1.btnRun.innerHTML = "전체선택"            
            strCHECK = 1
            
        End If
        
    End With
    
End Function

'=======================================================================================================


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
                                <TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
                                <TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>승인코드ERP전송</font></TD>
                                <TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
                               </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=*>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR HEIGHT=*>
        <TD  WIDTH=100% CLASS="Tab11">
            <TABLE <%=LR_SPACE_TYPE_20%>>
                <TR>
                    <TD <%=HEIGHT_TYPE_02%>></TD>
                </TR>
                <TR>
                    <TD HEIGHT=20 WIDTH=100%>
                        <FIELDSET CLASS="CLSFLD">
                            <TABLE <%=LR_SPACE_TYPE_40%>>
                                
                                <TR>
                                    <TD CLASS="TD5" NOWRAP>완료일자</TD>
                                    <TD CLASS="TD6">
                                        <script language =javascript src='./js/b82110ma1_fpDateTime5_txtDtFr.js'></script>&nbsp;~&nbsp;
                                        <script language =javascript src='./js/b82110ma1_fpDateTime6_txtDtTo.js'></script></TD>                
                                    <TD CLASS="TD5" NOWRAP>의뢰자</TD>
                                    <TD CLASS="TD6"><INPUT NAME="txtreq_user" ALT="의뢰자" TYPE="Text" SiZE=10 MAXLENGTH=13   tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('1')">
                                                    <INPUT NAME="txtreq_user_Nm" ALT="의뢰자명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>    
                                </TR>
                                <TR>
                                   <TD CLASS=TD5 NOWRAP>품목계정</TD>
                                   <TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct"  CLASS=cboNormal TAG="11" ALT="품목계정"><OPTION VALUE=""></OPTION></SELECT></TD>
                                   <TD CLASS=TD5 NOWRAP>품목구분</TD>
                                   <TD CLASS=TD6 NOWRAP><INPUT NAME="txtItem_Kind" ALT="품목구분" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('2')">
                                                        <INPUT NAME="txtItem_Kind_Nm" ALT="품목구분명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>
                                </TR>
                                <TR>                    
                                    <TD CLASS="TD5" NOWRAP>전송선택</TD>
                                    <TD CLASS="TD6"><INPUT TYPE="RADIO" NAME="rdoTransA" ID="rdoTrans1" Value="0" CLASS="RADIO" tag="1X"CHECKED><LABEL FOR="rdoTrans1">전체</LABEL>
                                                    <INPUT TYPE="RADIO" NAME="rdoTransA" ID="rdoTrans2" Value="1" CLASS="RADIO" tag="1X"><LABEL FOR="rdoTrans2">신규의뢰</LABEL>
                                                    <INPUT TYPE="RADIO" NAME="rdoTransA" ID="rdoTrans3" Value="2" CLASS="RADIO" tag="1X"><LABEL FOR="rdoTrans3">품목변경</LABEL>
                                                    <INPUT TYPE="RADIO" NAME="rdoTransA" ID="rdoTrans4" Value="3" CLASS="RADIO" tag="1X"><LABEL FOR="rdoTrans4">품명/규격변경</LABEL>                                                    
                                                    <INPUT TYPE= HIDDEN NAME="htxtTransA"  SIZE= 10 MAXLENGTH=10  TAG="14" ALT="전송선택"></TD>
                                    <TD CLASS="TD5" NOWRAP>전송구분</TD>
                                    <TD CLASS="TD6"><INPUT TYPE="RADIO" NAME="rdoTransB" ID="rdoTransT" Value="T" CLASS="RADIO" tag="1X"CHECKED><LABEL FOR="rdoTransT">전송</LABEL>
                                                    <INPUT TYPE="RADIO" NAME="rdoTransB" ID="rdoTransC" Value="C" CLASS="RADIO" tag="1X"><LABEL FOR="rdoTransC">전송취소</LABEL>
                                                    <INPUT TYPE= HIDDEN NAME="htxtTransB"  SIZE= 10 MAXLENGTH=10  TAG="14" ALT="전송구분"></TD>
                                </TR>                        
                            </TABLE>
                        </FIELDSET>
                    </TD>
                </TR>
                <TR>
                    <TD HEIGHT=*  WIDTH=100% VALIGN=TOP>                        
                        <TR>
                            <TD HEIGHT=100% WIDTH=100% Colspan=2>
                                <script language =javascript src='./js/b82110ma1_I987592315_vspdData.js'></script>
                            </TD>    
                        </TR>    
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
    <TR HEIGHT=12>
        <TD <%=HEIGHT_TYPE_03%> WIDTH=100%>
            <TABLE <%=LR_SPACE_TYPE_20%>>
                    <TR>
                        <TD>
                            <BUTTON NAME="btnRun" CLASS="CLSMBTN" ONCLICK="vbscript:Check()" Flag=1>전체선택</BUTTON>
                        </TD>                        
                    </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
        </TD>
    </TR>
</TABLE>

<TEXTAREA CLASS="HIDDEN" NAME="txtSpread" tag="24" tabindex=-1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex=-1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex=-1>

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
    </DIV>
</BODY>
</HTML>
