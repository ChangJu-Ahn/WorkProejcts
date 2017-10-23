<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 만근수당기준등록 
*  3. Program ID           : H1010ma1
*  4. Program Name         : H1010ma1
*  5. Program Desc         : 기준정보관리/만근수당기준등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/12
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : YBI
* 10. Modifier (Last)      : Lee Si Na
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "H1010mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 15	                                      '☜: Visble row

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim IsOpenPop

Dim C_DILIG 
Dim C_DILIG_POP
Dim C_DILIG_NM
Dim C_DILIG_CNT

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()
	 C_DILIG		= 1 
	 C_DILIG_POP	= 2
	 C_DILIG_NM		= 3
	 C_DILIG_CNT	= 4
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    lgKeyStream = Frm1.txtAllow_cd.Value & parent.gColSep       'You Must append one character( parent.gColSep)
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

	Call initSpreadPosVariables()	
    With frm1.vspdData
 
	    ggoSpread.Source = frm1.vspdData	
        ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_DILIG_CNT + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0
        ggoSpread.ClearSpreadData
	
       Call  AppendNumberPlace("6","3","0")
       Call  GetSpreadColumnPos("A")

        ggoSpread.SSSetEdit   C_DILIG,         "근태코드",     20,,,2,2
        ggoSpread.SSSetButton C_DILIG_POP
        ggoSpread.SSSetEdit   C_DILIG_NM,      "근태코드명",   70,,,15
        ggoSpread.SSSetFloat  C_DILIG_CNT,     "횟수",         24,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
        
        Call ggoSpread.MakePairsColumn(C_DILIG,  C_DILIG_POP)

	   .ReDraw = true
	
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
       ggoSpread.SpreadLock      C_DILIG,	-1, C_DILIG
       ggoSpread.SpreadLock      C_DILIG_NM, -1, C_DILIG_NM
       ggoSpread.SpreadLock      C_DILIG_POP,-1, C_DILIG_POP
       ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
       ggoSpread.SSSetRequired     C_DILIG		, pvStartRow, pvEndRow
       ggoSpread.SSSetprotected    C_DILIG_NM	, pvStartRow, pvEndRow
       ggoSpread.SSSetProtected	.vspdData.MaxCols, pvStartRow, pvEndRow
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
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)                
            
            C_DILIG		= iCurColumnPos(1) 
			C_DILIG_POP	= iCurColumnPos(2)
			C_DILIG_NM	= iCurColumnPos(3)
			C_DILIG_CNT	= iCurColumnPos(4)            
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call  AppendNumberPlace("7","2","0")
    Call  ggoOper.FormatField(Document, "2", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
	Call SetToolbar("1100110100101111")												'⊙: Set ToolBar
	Call CookiePage (0)                                                             '☜: Check Cookie

    ' 수당코드에 값을 
    Call  CommonQueryRs(" MAX(allow_cd) "," hda090t ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    frm1.txtAllow_cd.value = Trim(Replace(lgF0,Chr(11),""))

    Call  CommonQueryRs(" allow_nm "," HDA010T "," allow_cd =  " & FilterVar(frm1.txtAllow_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    frm1.txtAllow_nm.value = Trim(Replace(lgF0,Chr(11),""))

    frm1.txtAllow_cd.Focus    
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
    
    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

     ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			     '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.ClearSpreadData
    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
    
    If txtAllow_cd_Onchange() Then                                                '☜: enter key 로 조회시 수당코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    
    Call MakeKeyStream("X")
    
    Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
              
    FncQuery = True																'☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")		'☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call  ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call  ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
	Call SetToolbar("1110111100111111")							                 '⊙: Set ToolBar
    Call InitVariables                                                           '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd
    
    FncDelete = False                                                             '☜: Processing is NG
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                            'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first. 
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete? 
	If IntRetCD = vbNo Then											        
		Exit Function	
	End If
    
    Call  DisableToolBar( parent.TBC_DELETE)
	If DbDelete = False Then
		Call  RestoreToolBar()
        Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
            
    FncDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim lRow
    Dim strAllow_cd
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
     ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If    

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
		
	ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

    strAllow_cd = ""    
    
    IntRetCd =  CommonQueryRs(" distinct ALLOW_CD "," HDA090T "," ALLOW_CD <>  " & FilterVar(frm1.txtAllow_cd.value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strAllow_cd = Trim(Replace(lgF0,Chr(11),""))

    If IsNull(strAllow_cd) OR strAllow_cd = "" then
    Else
        Call  DisplayMsgBox("800492","x",UCase(strAllow_cd),"만근수당")                        
        Exit Function          
    End if 

	With Frm1
       For lRow = 1 To .vspdData.MaxRows
           .vspdData.Row = lRow
           .vspdData.Col = 0
           if   .vspdData.Text =  ggoSpread.InsertFlag OR .vspdData.Text =  ggoSpread.UpdateFlag then
				.vspdData.Col = C_DILIG_NM
				 if .vspdData.Text = "" then
					Call  DisplayMsgBox("970000","X","근태코드","X")
					.vspdData.focus
       	            exit function
				 end if 
            end if
        next
    End with    
    
    if  frm1.txtAllow.text <= 0 then
        call  DisplayMsgBox("800410", "x","x","x")
        frm1.txtAllow.focus
        exit function
    end if

    Call MakeKeyStream("X")
    
	Call  DisableToolBar( parent.TBC_SAVE)
	If DbSave = False Then
		Call  RestoreToolBar()
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
    
            .ReDraw = True
		    .Focus
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
    Call InitData()
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
        	lDelRows =  ggoSpread.DeleteRow
        End With
        Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport( parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	 ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
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
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    If LayerShowHide(1) = False Then
			Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
	Dim strVal, strDel
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	If LayerShowHide(1) = False Then
			Exit Function
	End If
		
	With frm1
		.txtMode.value        =  parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	    With Frm1
    
           For lRow = 1 To .vspdData.MaxRows
    
               .vspdData.Row = lRow
               .vspdData.Col = 0

               Select Case .vspdData.Text
                   Case  ggoSpread.InsertFlag                                      '☜: Insert
                                                       strVal = strVal & "C" & parent.gColSep
                                                       strVal = strVal & lRow & parent.gColSep
                                                       strVal = strVal & Trim(.txtAllow_cd.Value) & parent.gColSep
                        .vspdData.Col = C_DILIG      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_DILIG_CNT  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep   
                        lGrpCnt = lGrpCnt + 1
                   Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                       strVal = strVal & "U" & parent.gColSep
                                                       strVal = strVal & lRow & parent.gColSep
                                                       strVal = strVal & Trim(.txtAllow_cd.Value) & parent.gColSep
                        .vspdData.Col = C_DILIG      : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                        .vspdData.Col = C_DILIG_CNT  : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep   
                        lGrpCnt = lGrpCnt + 1
                   Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                       strDel = strDel & "D" & parent.gColSep
                                                       strDel = strDel & lRow & parent.gColSep
                                                       strDel = strDel & Trim(.txtAllow_cd.Value) & parent.gColSep
                        .vspdData.Col = C_DILIG 	 : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                        lGrpCnt = lGrpCnt + 1
               End Select
           Next
	
	    End With

    Frm1.txtMaxRows.value = lGrpCnt-1	
	Frm1.txtSpread.value = strDel & strVal
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	If LayerShowHide(1) = False Then
			Exit Function
	End If
		
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003                                '☜: Delete
	strVal = strVal & "&txtAllow_cd=" & Trim(frm1.txtAllow_cd.value)             '☜: 

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False
    Frm1.txtAllow_cd.focus 

	Call SetToolbar("1101111100111111")												'⊙: Set ToolBar
    Call InitData()
    Call  ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    lgBlnFlgChgValue = False
	Call InitVariables
     ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    ggoSpread.ClearSpreadData
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call FncNew()	
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
' Name : SubOpenCollateralNoPop()
' Desc : developer describe this line Call Master L/C No PopUp
'========================================================================================================
Sub SubOpenCollateralNoPop()
	Dim strRet
	If gblnWinEvent = True Then Exit Sub
	gblnWinEvent = True
		
	strRet = window.showModalDialog("s1413pa1.asp", "", _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet = "" Then
       Exit Sub
	Else
       Call SetCollateralNo(strRet)
	End If	
End Sub

'========================================================================================================
' Name : OpenAllowCd()
' Desc : developer describe this line 
'========================================================================================================
Function OpenAllowCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True

	arrParam(0) = "수당코드 팝업"		' 팝업 명칭 
	arrParam(1) = "HDA010T"				 	' TABLE 명칭 
	arrParam(2) = frm1.txtAllow_cd.value	' Code Condition
	arrParam(3) = ""	                    ' Name Cindition
	arrParam(4) = " pay_cd=" & FilterVar("*", "''", "S") & "  AND code_type=" & FilterVar("1", "''", "S") & " "' Where Condition
	arrParam(5) = "수당코드"			
	
    arrField(0) = "allow_cd"				' Field명(0)
    arrField(1) = "allow_nm"				' Field명(1)
    
    arrHeader(0) = "수당코드"			' Header명(0)
    arrHeader(1) = "수당코드명"			' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		Frm1.txtAllow_cd.focus	
		Exit Function
	Else
		Call SubSetAllow(arrRet)
	End If	
	
End Function

'======================================================================================================
'	Name : SetAllow()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetAllow(arrRet)
	With Frm1
		.txtAllow_cd.value = arrRet(0)
		.txtAllow_nm.value = arrRet(1)
		.txtAllow_cd.focus		
	End With
End Sub

'======================================================================================================
' Name : OpenDiligPopup
' Desc : OpenDiligPopup Reference Popup
'======================================================================================================
Function OpenDiligPopup(Byval strCode, Byval Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

     ggoSpread.Source = frm1.vspdData

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
        arrParam(0) = "근태코드"
		arrParam(1) = "HCA010T"         ' TABLE 명칭 
		arrParam(2) = strCode								           ' Code Condition
		arrParam(3) = ""									           ' Name Cindition
		arrParam(4) = ""		                                       ' Where Condition
		arrParam(5) = "근태코드"	         			           ' TextBox 명칭 
	
		arrField(0) = "dilig_cd"			    			           ' Field명(0)
		arrField(1) = "dilig_nm"		    				           ' Field명(1)
    
		arrHeader(0) = "근태코드"						           ' Header명(0)
		arrHeader(1) = "근태코드명"						           ' Header명(1)

	arrParam(0) = arrParam(5)							               ' 팝업 명칭 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.vspdData.action = 0 
		Exit Function
	Else
		Call SetDiligPopUp(arrRet)
	End If	
	
End Function

'========================================================================================================
' Name : SetCollateralNo()
' Desc : developer describe this line 
'========================================================================================================
Function SetCollateralNo(arrRet)
	frm1.txtGlNo.Value = arrRet
End Function

'========================================================================================================
' Name : SetBizPartner()
' Desc : developer describe this line 
'========================================================================================================
Function SetBizPartner(arrRet)
	frm1.txtCustomer.Value = arrRet(0)
	frm1.txtCustomerNm.Value = arrRet(1)
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Name : SetDiligPopUp()
' Desc : OpenSalesPlanPopup에서 Return되는 값 setting
'========================================================================================================
Function SetDiligPopUp(Byval arrRet)

	With frm1
		.vspdData.Col = C_DILIG_NM
		.vspdData.Text = arrRet(1)
		.vspdData.Col = C_DILIG
		.vspdData.Text = arrRet(0)
		.vspdData.action = 0 
	End With

	lgBlnFlgChgValue = True
	
End Function

'========================================================================================================
' Name : SetCurrency
' Desc : developer describe this line 
'========================================================================================================
Function SetCurrency(arrRet)
	frm1.txtCurrency.Value = arrRet(0)
	lgBlnFlgChgValue = True
End Function


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData.Row = Row
     
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData.MaxRows = 0 then
		exit sub
	end if
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
		 ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
			Case C_DILIG_POP
				.Col = Col - 1
				.Row = Row
				Call OpenDiligPopup(.Text, Row)
			End Select
		End If
    
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col

        Case C_DILIG

            iDx =  CommonQueryRs("dilig_nm ","HCA010T","dilig_cd= " & FilterVar(Frm1.vspdData.Value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            if  iDx = true then
                Frm1.vspdData.Col = C_DILIG_NM
                Frm1.vspdData.Value = Replace(lgF0, Chr(11), "")
            else
			    Call  DisplayMsgBox("970027","X","근태코드","X")
  	            
                Frm1.vspdData.Col = C_DILIG_NM
                Frm1.vspdData.value = ""                
            end if

    End Select    
             
   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Private Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
   	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row

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

Sub txtAllow_Change()
    lgBlnFlgChgValue = True
End Sub

function txtAllow_cd_OnChange()
	Dim IntRetCd
    
    If frm1.txtAllow_cd.value = "" Then
		frm1.txtAllow_nm.value = ""
    Else
        IntRetCd =  CommonQueryRs(" ALLOW_NM "," HDA010T "," PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "   AND ALLOW_CD =  " & FilterVar(frm1.txtAllow_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false then
			Call  DisplayMsgBox("800145","X","X","X")  '수당정보에 등록되지 않은 코드입니다.
			frm1.txtAllow_nm.value = ""
            frm1.txtAllow_cd.focus
		    Call  ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field            
			txtAllow_cd_Onchange = true
            Exit Function          
        Else
			frm1.txtAllow_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
End function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>만근수당기준등록</font></td>
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
								<TR>
									<TD CLASS=TD5 NOWRAP>수당코드</TD>
                                    <TD CLASS=TD6 NOWRAP>
                                    
										<INPUT TYPE=TEXT NAME="txtAllow_cd" MAXLENGTH=3 SIZE=10 MAXLENGTH=8 tag=12XXXU ALT="수당코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWarrentNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenAllowCd()">
										<INPUT TYPE=TEXT NAME="txtAllow_nm" tag="14X"></TD>

									<TD CLASS=TDT NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%>>
						    <TR>
						    	<TD HEIGHT="100%" WIDTH="100%">
						    		<script language =javascript src='./js/h1010ma1_vspdData_vspdData.js'></script>
						    	</TD>
						    </TR>
						</TABLE>
                    </TD>
                </TR>
                <TR> 
					<TD WIDTH=100% HEIGHT=30 VALIGN=TOP>
                        <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>만근수당액</LEGEND>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						    <TR>
						    	<TD CLASS=TD5 NOWRAP>정액</TD>
						    	<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h1010ma1_txtAllow_txtAllow.js'></script></TD>
						    	<TD CLASS=TDT NOWRAP></TD>
						    	<TD CLASS=TD6 NOWRAP></TD>
						    </TR>
                        </TABLE>
                        </FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR >
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

