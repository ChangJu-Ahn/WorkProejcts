<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name           : Human Resources
*  2. Function Name         : 
*  3. Program ID            : H6001ma1
*  4. Program Name          : H6001ma1
*  5. Program Desc          : 기본급 테이블 등록 
*  6. Comproxy List         :
*  7. Modified date(First)  : 2001/05/31
*  8. Modified date(Last)   : 2003/06/13
*  9. Modifier (First)      : Kang Doo Sig
* 10. Modifier (Last)       : Lee SiNa
* 11. Comment               :
=======================================================================================================-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">  

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "h6001mb1.asp"                                      'bizlogic ASP name
Const BIZ_PGM_JUMP_ID  = "h6001ba1"

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          
Dim lsAllowCd
Dim AllowCd(2)
AllowCd(0) = ""
AllowCd(1) = ""
AllowCd(2) = ""

Dim C_GradeCd ' 급호코드 
Dim C_GradePopup ' 급호Popup
Dim C_Grade   ' 급호명 
Dim C_Hobong  ' 호봉 
Dim C_AllowCd1 ' 기본수당 시작 
Dim C_AllowCd2 ' 기본수당 시작 
Dim C_AllowCd3 ' 기본수당 시작 
Dim C_Total    ' 합계 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

    C_GradeCd   = 1 ' 급호코드 
    C_GradePopup  = 2 ' 급호Popup
    C_Grade   = 3 ' 급호명 
    C_Hobong   = 4 ' 호봉 
    C_AllowCd1  = 5 ' 기본수당 시작 
    C_AllowCd2  = 6 ' 기본수당 시작 
    C_AllowCd3  = 7 ' 기본수당 시작 
    C_Total   = 8 ' 합계 

End Sub

'========================================================================================================
' Name : InitVariables() 
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
 lgIntFlgMode      = parent.OPMD_CMODE          '⊙: Indicates that current mode is Create mode
 lgBlnFlgChgValue  = False           '⊙: Indicates that no value changed
 lgIntGrpCount     = 0            '⊙: Initializes Group View Size
 lgStrPrevKey      = ""                                           '⊙: initializes Previous Key
 lgSortKey         = 1                                             '⊙: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal() 
' Desc : Set default value
'======================================================================================================== 
Sub SetDefaultVal()

	frm1.txtStandardDt.text = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtStandardDt.focus()
End Sub

'========================================================================================================
' Name : LoadInfTB19029() 
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
    On Error Resume Next
    Const CookieSplit = 4877 
    
    If flgs = 1 Then   
       WriteCookie "APPLY_DT" , frm1.txtStandardDt.Text
    End If

End Function

Function PgmJumpCheck()         
 
    If  Trim(frm1.txtStandardDt.Text) = "" Then
        Call DisplayMsgBox("970021","X",frm1.txtStandardDt.alt,"X")                      '기준년월은 필수 입력사항입니다.
        frm1.txtPay_yymm.focus    ' go to 
        Exit Function
    End If
    
    PgmJump(BIZ_PGM_JUMP_ID)
     
End Function   

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
     lgKeyStream   = Frm1.txtStandardDt.Text & parent.gColSep   '--->start apply day 
     lgKeyStream   = lgKeyStream & Trim(Frm1.txtPay_grd.Value) & parent.gColSep              '--->grade
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Dim cd, nm
    Dim x, y, iCnt

	Call initSpreadPosVariables()   'sbk 

    With frm1.vspdData

         ggoSpread.Source = frm1.vspdData

         ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

        .ReDraw = false

        .MaxCols = C_Total + 1             '☜: 최대 Columns의 항상 1개 증가시킴 
        .Col = .MaxCols                 '공통콘트롤 사용 Hidden Column
        .ColHidden = True
        
        .MaxRows = 0
         ggoSpread.ClearSpreadData

         Call GetSpreadColumnPos("A") 'sbk
         
         ggoSpread.SsSetEdit  C_GradeCd         , "급호코드" ,10         
         ggoSpread.SSSetButton  C_GradePopup 
         ggoSpread.SSSetEdit  C_Grade           , "급호명", 16,,, 20, 2
         ggoSpread.SSSetEdit  C_Hobong          , "호봉", 10,,,  3, 2     
         
         Call CommonQueryRs("count(*)"," hda010t "," code_type = " & FilterVar("1", "''", "S") & "  and allow_kind = " & FilterVar("1", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
         iCnt = Replace(lgF0,Chr(11),"")
         
         Call CommonQueryRs("allow_cd, allow_nm"," hda010t "," code_type = " & FilterVar("1", "''", "S") & "  and allow_kind = " & FilterVar("1", "''", "S") & "  order by allow_cd",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
         cd = Split(lgF0,Chr(11))
         nm = Split(lgF1,Chr(11))

         lsAllowCd = Ubound(cd)
         y=0

         For x =  1 To iCnt
            SELECT Case x
                Case 1
                    ggoSpread.SSSetFloat  C_AllowCd1,   nm(0), 19, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
                Case 2
                    ggoSpread.SSSetFloat  C_AllowCd2,   nm(1), 19, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
                Case 3
                    ggoSpread.SSSetFloat  C_AllowCd3,   nm(2), 19, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
            End SELECT
            
            AllowCd(y) = cd(y)
            
            y=y+1
         Next

         If iCnt < 3 Then
            For x =  iCnt+1 To 3
               SELECT Case x
                Case 1
                    ggoSpread.SSSetFloat  C_AllowCd1,   " ", 19, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
                Case 2
                    ggoSpread.SSSetFloat  C_AllowCd2,   " ", 19, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
                Case 3
                    ggoSpread.SSSetFloat  C_AllowCd3,   " ", 19, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
               End SELECT
            Next
         End If 

         ggoSpread.SSSetFloat  C_Total            , "합계", 20, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
 
         Call ggoSpread.MakePairsColumn(C_GradeCd,C_GradePopup)    'sbk

        .ReDraw = True

         Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
        .vspdData.ReDraw = False

        ggoSpread.SpreadLock    C_GradeCd, -1, C_GradeCd, -1     'sbk
        ggoSpread.SpreadLock    C_GradePopup, -1, C_GradePopup, -1
        ggoSpread.SpreadLock    C_Grade, -1, C_Grade, -1
        ggoSpread.SpreadLock    C_Hobong, -1, C_Hobong, -1
        ggoSpread.SpreadLock    C_Total, -1, C_Total, -1
        ggoSpread.SSSetProtected   .vspdData.MaxCols   , -1, -1
   
        .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1    
       .vspdData.ReDraw = False
       ggoSpread.SSSetRequired  C_GradeCd, pvStartRow, pvEndRow
       ggoSpread.SSSetProtected C_Grade  , pvStartRow, pvEndRow
       ggoSpread.SSSetRequired  C_Hobong , pvStartRow, pvEndRow
       ggoSpread.SSSetProtected C_Total  , pvStartRow, pvEndRow
       .vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_GradeCd      = iCurColumnPos(1)
            C_GradePopup   = iCurColumnPos(2)
            C_Grade        = iCurColumnPos(3)
            C_Hobong       = iCurColumnPos(4)
            C_AllowCd1     = iCurColumnPos(5)
            C_AllowCd2     = iCurColumnPos(6)
            C_AllowCd3     = iCurColumnPos(7)
            C_Total        = iCurColumnPos(8)
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    
    Err.Clear                                                                        '☜: Clear err status
    Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")           '⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                               'Initializes local global variables
    Call SetDefaultVal
    
    Call SetToolbar("1100110100101111")            '⊙: Set ToolBar
    Call CookiePage (0)                                                             '☜: Check Cookie
       
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False                '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
       IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")      '☜: Data is changed.  Do you want to display it? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If   

    ggoSpread.ClearSpreadData
                   
    If Not chkField(Document, "1") Then                  '☜: This function check required field
       Exit Function
    End If

    if  txtPay_grd_Onchange() then
		Exit Function
	end if

    Call InitVariables                                                        '⊙: Initializes local global variables
    Call MakeKeyStream("X")

    Call DisableToolBar(parent.TBC_QUERY)
    
    IF DBQUERY =  False Then
       Call RestoreToolBar()
       Exit Function
    End If
       
    FncQuery = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim lRow
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then                  '☜: This function check required field
       Exit Function
    End If
    
    If frm1.txtStandardDt.Text="" Then
        Call DisplayMsgBox("970021","X",frm1.txtStandardDt.alt,"X")                      '기준년월은 필수 입력사항입니다.
        frm1.txtStandardDt.focus    ' go to 
        Exit Function
    End If                    

    With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
          
            Select Case .vspdData.Text
           
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
               
                    .vspdData.Col = C_Grade
                    If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
                        Call DisplayMsgBox("970010","X","급호코드","X")
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if
                    
                    .vspdData.Col = C_Hobong
                    if (Trim(frm1.vspdData.text) > "99" or Trim(frm1.vspdData.text) < "00") then
                        Call DisplayMsgBox("700119","X","호봉","X")             '☜ : 등록되지 않은 코드입니다.
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        exit Function
                    end if
            End Select
        Next
    End With
    
    Call DisableToolBar(parent.TBC_SAVE)
    Call MakeKeyStream("X")

    IF DBSAVE =  False Then
       Call RestoreToolBar()
       Exit Function
    End If
    
    FncSave = True                                                              '☜: Processing is OK
    
End Function
'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
    With Frm1.VspdData
        .ReDraw = False
        If .ActiveRow > 0 Then
           ggoSpread.CopyRow
           SetSpreadColor .ActiveRow, .ActiveRow
           .Col = C_Grade
           .Text = ""
           .Col = C_GradeCd
           .Text = ""
           .Col = C_Hobong
           .Text = ""
           .ReDraw = True
           .Col = C_Grade
           .Focus
           .Action = 0 ' go to 
        End If
    End With
 
    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData 
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)

    Dim imRow

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
 
    FncInsertRow = False                                                         '☜: Processing is NG

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
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
       .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
        
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    
    If Frm1.vspdData.MaxRows < 1 then
       Exit function
    End if 
    
    With Frm1.vspdData 
       .focus
       ggoSpread.Source = frm1.vspdData 
       lDelRows = ggoSpread.DeleteRow
    End With
    
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
 
    FncExit = False
 
    ggoSpread.Source = frm1.vspdData 
    If ggoSpread.SSCheckChange = True Then
       IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")   '⊙: Data is changed.  Do you want to exit? 
       If IntRetCD = vbNo Then
           Exit Function
       End If
    End If
    
    FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 
    Dim strVal

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

    If   LayerShowHide(1) = False Then
         Exit Function
    End If
 
    With Frm1
        strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001               
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With

    Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave = False                                                          
    
    If   LayerShowHide(1) = False Then
         Exit Function
    End If

    strVal = ""
    strDel = ""
    lGrpCnt = 1

    With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
           
           Select Case .vspdData.Text
               Case ggoSpread.InsertFlag                                      '☜: Insert
                                                              strVal = strVal & "C" & parent.gColSep
                                                              strVal = strVal & lRow & parent.gColSep
                                                              strVal = strVal & Trim(Frm1.txtStandardDt.Text) & parent.gColSep
                                                              strVal = strVal & Trim(AllowCd(0)) & parent.gColSep
                                                              strVal = strVal & Trim(AllowCd(1)) & parent.gColSep
                                                              strVal = strVal & Trim(AllowCd(2)) & parent.gColSep
                    .vspdData.Col = C_GradeCd          : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                 
                    .vspdData.Col = C_Hobong          : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_AllowCd1       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_AllowCd2       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_AllowCd3       : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    '.vspdData.Col = C_AllowCd4       : strVal = strVal & Trim(.vspdData.Value) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                              strVal = strVal & "U" & parent.gColSep
                                                              strVal = strVal & lRow & parent.gColSep
                                                              strVal = strVal & Trim(Frm1.txtStandardDt.Text) & parent.gColSep
                                                              strVal = strVal & Trim(AllowCd(0)) & parent.gColSep
                                                              strVal = strVal & Trim(AllowCd(1)) & parent.gColSep
                                                              strVal = strVal & Trim(AllowCd(2)) & parent.gColSep
                    .vspdData.Col = C_GradeCd          : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Hobong          : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_AllowCd1       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_AllowCd2       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_AllowCd3       : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    '.vspdData.Col = C_AllowCd4       : strVal = strVal & Trim(.vspdData.Value) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                              strDel = strDel & "D" & parent.gColSep
                                                              strDel = strDel & lRow & parent.gColSep
                                                              strDel = strDel & Trim(Frm1.txtStandardDt.Text) & parent.gColSep
                    .vspdData.Col = C_GradeCd          : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Hobong          : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

    .txtMode.value        = parent.UID_M0002
    .txtMaxRows.value     = lGrpCnt-1 
    .txtSpread.value      = strDel & strVal

    End With
 Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")              '⊙: "Will you destory previous data"
    If IntRetCD = vbNo Then             '------ Delete function call area ------ 
        Exit Function 
    End If    
    
    Call DisableToolBar(parent.TBC_DELETE)

    IF DbDelete =  False Then
       Call RestoreToolBar()
       Exit Function
    End If
    
    FncDelete = True                                                        '⊙: Processing is OK

End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()                  
    lgIntFlgMode = parent.OPMD_UMODE   
     
    Call ggoOper.LockField(Document, "Q")          '⊙: Lock field
    Call InitData()
    Call SetToolbar("1100111100111111")            '⊙: Set ToolBar
    frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

    Call InitVariables               '⊙: Initializes local global variables

    Call DisableToolBar(parent.TBC_QUERY)

    IF DBQUERY =  False Then
       Call RestoreToolBar()
       Exit Function
    End If

End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function
'========================================================================================================
' Name : FncOpenPopup
' Desc : developer describe this line 
'========================================================================================================
Function FncOpenPopup(Byval iWhere)
 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True  Then  
    Exit Function
 End If   

 IsOpenPop = True
 
 Select Case iWhere
     Case "1"

         arrParam(0) = "급호조회 팝업"   ' 팝업 명칭 
         arrParam(1) = "B_MINOR"       ' TABLE 명칭 
         arrParam(2) = frm1.txtPay_grd.value         ' Code Condition
         arrParam(3) = ""'frm1.txtPay_grd_nm.value  ' Name Cindition
         arrParam(4) = " MAJOR_CD=" & FilterVar("H0001", "''", "S") & ""      ' Where Condition
         arrParam(5) = "급호코드"       ' TextBox 명칭 
 
         arrField(0) = "MINOR_CD"     ' Field명(0)
         arrField(1) = "MINOR_NM"        ' Field명(1)
    
         arrHeader(0) = "급호코드"    ' Header명(0)
         arrHeader(1) = "급호명"           ' Header명(1)
            
       Case "2"   
         arrParam(0) = "기준일 팝업"       ' 팝업 명칭 
         arrParam(1) = "Hdf010T"      ' TABLE 명칭 
         arrParam(2) = ""                            ' Code Condition
         arrParam(3) = ""                       ' Name Cindition
         arrParam(4) = "APPLY_STRT_DT <> '' GROUP BY APPLY_STRT_DT "   ' Where Condition
         arrParam(5) = "기준일"           ' TextBox 명칭 
 
         arrField(0) = "APPLY_STRT_DT"       ' Field명(0)
         arrField(1) = "PAY_GRD2"        ' Field명(1)
    
         arrHeader(0) = "기준일"        ' Header명(0)
         arrHeader(1) = "기준일명"          ' Header명(1)
            
 End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
  
 
	If arrRet(0) = "" Then
			Select Case iWhere
			    Case "1"
			        frm1.txtPay_grd.focus
			    Case "2"
			        frm1.txtStandardDt.focus
			End Select        
				 
			Exit Function
	Else
		  Call SubSetOpenPop(arrRet,iWhere)
	End If 
 
End Function

'======================================================================================================
' Name : SubSetOpenPop()
' Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetOpenPop(Byval arrRet, Byval iWhere)
 With Frm1
  Select Case iWhere
      Case "1"
          .txtPay_grd.value = arrRet(0)
          .txtPay_grd_nm.value = arrRet(1)  
          .txtPay_grd.focus
      Case "2"
          .txtStandardDt.value = arrRet(0)
          .txtStandardDt.focus
  End Select        
 End With
End Sub
'======================================================================================================
' Name : OpenCode()
' Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

 IsOpenPop = True

 Select Case iWhere
     Case C_GradePopup
        arrParam(0) = "급호코드 팝업"           ' 팝업 명칭 
        arrParam(1) = "B_minor"           ' TABLE 명칭 
        arrParam(2) = ""                              ' Code Condition
        arrParam(3) = strCode        ' Name Cindition
        arrParam(4) = "major_cd=" & FilterVar("H0001", "''", "S") & ""        ' Where Condition
        arrParam(5) = "급호코드"                ' TextBox 명칭 
 
        arrField(0) = "minor_cd"       ' Field명(0)
        arrField(1) = "minor_nm"          ' Field명(1)
    
        arrHeader(0) = "급호코드"               ' Header명(0)
        arrHeader(1) = "급호코드명"               ' Header명(1)
 End Select
    
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
 IsOpenPop = False
 
 If arrRet(0) = "" Then
		frm1.vspdData.Col = C_GradeCd
		frm1.vspdData.action = 0
 
  Exit Function
 Else
  Call SetCode(arrRet, iWhere)
        ggoSpread.Source = frm1.vspdData
        ggoSpread.UpdateRow Row
 End If 

End Function

'======================================================================================================
' Name : SetCode()
' Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)

 With frm1

  Select Case iWhere
      Case C_GradePopup
          .vspdData.Col = C_Grade
          .vspdData.text = arrRet(1)   
			.vspdData.Col = C_GradeCd
          .vspdData.text = arrRet(0) 
			.vspdData.action = 0
  End Select

 End With

End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
 
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
                Case C_GradePopup
				    .Col = C_GRADE
				    .Row = Row
                     Call OpenCode(frm1.vspdData.Text, C_GradePopup, Row)
			End Select
		End If
	End With
            
End Sub

'========================================================================================================
'   Event Name : vspdData_Change 
'   Event Desc :
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row)
    Dim IntRetCD
    Dim dblAllow1,dblAllow2,dblAllow3

    frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col

    Select Case Col
     Case C_GradeCd
            IntRetCD = CommonQueryRs(" minor_nm "," b_minor "," major_cd=" & FilterVar("H0001", "''", "S") & " And minor_cd =  " & FilterVar(frm1.vspdData.Text, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

            If IntRetCD=False And Trim(frm1.vspdData.Text)<>"" Then
                frm1.vspdData.Col = C_Grade
                frm1.vspdData.Text= ""
                Call DisplayMsgBox("970000","X","급호코드","X")             '☜ : 등록되지 않은 코드입니다.
                vspdData_Change = true
                Exit Function
            Else
                frm1.vspdData.Col = C_Grade
                frm1.vspdData.Text=Trim(Replace(lgF0,Chr(11),""))
            End If
     
     Case C_Hobong
          frm1.vspdData.Col = C_Hobong
          
          if (Trim(frm1.vspdData.text) > "99" or Trim(frm1.vspdData.text) < "00") then
             Call DisplayMsgBox("700119","X","호봉","X")             '☜ : 등록되지 않은 코드입니다.
             vspdData_Change = true
          end if
       
     Case C_AllowCd1
          frm1.vspdData.Col = C_AllowCd1
          dblAllow1=frm1.vspdData.Text
          frm1.vspdData.Col = C_AllowCd2
          dblAllow2=frm1.vspdData.Text
          frm1.vspdData.Col = C_AllowCd3
          dblAllow3=frm1.vspdData.Text
             
          frm1.vspdData.Col = C_Total
          frm1.vspdData.Text= UNIFormatNumber(UNICDbl(dblAllow1)+UNICDbl(dblAllow2)+UNICDbl(dblAllow3),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
     Case C_AllowCd2
          frm1.vspdData.Col = C_AllowCd1
          dblAllow1=frm1.vspdData.Text
          frm1.vspdData.Col = C_AllowCd2
          dblAllow2=frm1.vspdData.Text
          frm1.vspdData.Col = C_AllowCd3
          dblAllow3=frm1.vspdData.Text
  
          frm1.vspdData.Col = C_Total
          frm1.vspdData.Text= UNIFormatNumber(UNICDbl(dblAllow1)+UNICDbl(dblAllow2)+UNICDbl(dblAllow3),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
     Case C_AllowCd3
          frm1.vspdData.Col = C_AllowCd1
          dblAllow1=frm1.vspdData.Text
          frm1.vspdData.Col = C_AllowCd2
          dblAllow2=frm1.vspdData.Text
          frm1.vspdData.Col = C_AllowCd3
          dblAllow3=frm1.vspdData.Text
  
          frm1.vspdData.Col = C_Total
          frm1.vspdData.Text= UNIFormatNumber(UNICDbl(dblAllow1)+UNICDbl(dblAllow2)+UNICDbl(dblAllow3),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
    End Select   

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
       If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
          Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
       End If
    End If
 
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
         
End Function


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       Exit Sub
    End If

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If

End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub

'======================================================================================================
'   Event Name : txtPay_grd_OnChange
'   Event Desc : 직급코드가 변경될 경우 
'=======================================================================================================
Function txtPay_grd_OnChange()
    Dim IntRetCd

    If Trim(frm1.txtPay_grd.value) = "" Then
        frm1.txtPay_grd_nm.Value = ""
    Else
        IntRetCD = CommonQueryRs(" minor_nm "," b_minor "," major_cd=" & FilterVar("H0001", "''", "S") & " And minor_cd =  " & FilterVar(frm1.txtPay_grd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtPay_grd.Value)<>""  Then
            frm1.txtPay_grd_nm.Value=""
            Call DisplayMsgBox("970000","X","급호코드","X")             '☜ : 등록되지 않은 코드입니다.
			txtPay_grd_OnChange = true
        Else
            frm1.txtPay_grd_nm.Value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function
'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtStandardDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtStandardDt.Action = 7
        frm1.txtStandardDt.focus
    End If
    lgBlnFlgChgValue = True
End Sub
'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : 날짜 스크롤 버튼으로 데이타 변경시 
'=======================================================================================================
Sub txtStandardDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtStandardDt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtStandardDt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

'========================================================================================================
' Name : OpenStandardDt
' Desc : 기준일 POPUP
'========================================================================================================
Function OpenStandardDt(iWhere)
	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "기준일팝업"		
	arrParam(1) = "기준일"
	arrParam(2) = "Hdf010T"
	arrParam(3) = "APPLY_STRT_DT"
	arrParam(4) = frm1.txtStandardDt.text	

	arrRet = window.showModalDialog(HRAskPRAspName("StandardDtPopup"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	if arrRet(0) <> ""	then
		frm1.txtStandardDt.text = arrRet(0)
	end if
	frm1.txtStandardDt.focus
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
    <TD WIDTH="100%">
        <TABLE <%=LR_SPACE_TYPE_10%>>
            <TR>
                <TD WIDTH=10>&nbsp;</TD>
                <TD CLASS="CLSMTABP">
                    <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
                        <TR>
                            <td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
                            <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기본급테이블</font></td>
                            <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
                        </TR>
                    </TABLE>
                </TD>
                <TD WIDTH=*>&nbsp;</TD>
            </TR>
        </TABLE>
    </TD>
 </TR>
  
    <TR HEIGHT=*>
  <TD width=100% CLASS="Tab11">
   <TABLE <%=LR_SPACE_TYPE_20%>>
       <TR>
           <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
       </TR>
    <TR>
     <TD HEIGHT=20 WIDTH=100%>
      <FIELDSET CLASS="CLSFLD">
      <TABLE <%=LR_SPACE_TYPE_40%>>
       <TR>
        <TD CLASS="TD5">기준일</TD>
        <TD CLASS="TD6"><script language =javascript src='./js/h6001ma1_txtStandardDt_txtStandardDt.js'></script>
        <IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPromoteDt" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenStandardDt(0)"></TD>
        <TD CLASS="TD5">급호</TD>
        <TD CLASS=TD6 NOWRAP>
            <INPUT NAME="txtPay_grd"     SIZE=10  MAXLENGTH=2  ALT ="직급" TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: FncOpenPopup(1)">
            <INPUT NAME="txtPay_grd_nm"  SIZE=20  MAXLENGTH=50 TAG="14XXXU">
                                </TD>
       </TR>
      </TABLE>
      </FIELDSET>
     </TD>
    </TR> 
    <TR>
        <TD <%=HEIGHT_TYPE_03%>></TD>
    </TR>
    <TR>
     <TD WIDTH=100% HEIGHT=* VALIGN=TOP>
      <TABLE <%=LR_SPACE_TYPE_20%> >
       <TR>
        <TD HEIGHT=100% WIDTH=100% >
         <script language =javascript src='./js/h6001ma1_vaSpread_vspdData.js'></script>
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
         <TABLE <%=LR_SPACE_TYPE_30%>>
       <TR>
    <TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJumpCheck()" ONCLICK="VBSCRIPT:CookiePage 1">일괄생성</a>
    </TD>
    <TD WIDTH=10>&nbsp;</TD>
    </TR>
   </TABLE>
     </TD>   
 
 </TR>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
  </TD>
 </TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
