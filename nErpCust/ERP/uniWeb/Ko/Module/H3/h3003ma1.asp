<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h3003ma1
*  4. Program Name         : 인사발령사항 조회 
*  5. Program Desc         : 인사발령사항 조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/12/28
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : Myesongsik Song
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h3003mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row

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
Dim lgOldRow
Dim lsInternal_cd

Dim C_HBA010T_GAZET_DT 									   'Column Dimant for Spread Sheet 
Dim C_HBA010T_GAZET_CD 
Dim C_HAA010T_NAME 
Dim C_HBA010T_EMP_NO 
Dim C_HBA010T_DEPT_NM
Dim C_HBA010T_ROLL_PSTN_NM 
Dim C_HBA010T_PAY_GRD1_NM 
Dim C_HBA010T_PAY_GRD2 
Dim C_HAA010T_ENTR_DT 
Dim C_HAA010T_RETIRE_DT 
Dim C_HAA010T_SCH_SHIP 
Dim C_HAA030T_SCHOOL_NM 
Dim C_HAA030T_MAJOR_NM 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
sub InitSpreadPosVariables()
	C_HBA010T_GAZET_DT = 1
	C_HBA010T_GAZET_CD = 2
	C_HAA010T_NAME = 3
	C_HBA010T_EMP_NO = 4
	C_HBA010T_DEPT_NM = 5
	C_HBA010T_ROLL_PSTN_NM = 6
	C_HBA010T_PAY_GRD1_NM = 7
	C_HBA010T_PAY_GRD2 = 8
	C_HAA010T_ENTR_DT = 9
	C_HAA010T_RETIRE_DT = 10
	C_HAA010T_SCH_SHIP = 11
	C_HAA030T_SCHOOL_NM = 12
	C_HAA030T_MAJOR_NM = 13

end sub

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
	lgOldRow = 0

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
   
    lgKeyStream       = Trim(Frm1.txtEmp_no.Value) & parent.gColSep       'You Must append one character( parent.gColSep)
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtgazet_start_dt.text) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtgazet_end_dt.text) & parent.gColSep    
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtgazet_cd.Value) & parent.gColSep    
    If  lsInternal_cd = "" then
        lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep
    Else
        lgKeyStream = lgKeyStream & lsInternal_cd & parent.gColSep
    End If

End Sub        
	
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
    If Frm1.vspdData.MaxRows > 0 Then
        Call vspdData_Click(1 , 1)
		Frm1.vspdData.focus
        Set gActiveElement = document.ActiveElement
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
            
			C_HBA010T_GAZET_DT = iCurColumnPos(1)
			C_HBA010T_GAZET_CD = iCurColumnPos(2)
			C_HAA010T_NAME = iCurColumnPos(3)
			C_HBA010T_EMP_NO = iCurColumnPos(4)
			C_HBA010T_DEPT_NM = iCurColumnPos(5)
			C_HBA010T_ROLL_PSTN_NM = iCurColumnPos(6)
			C_HBA010T_PAY_GRD1_NM = iCurColumnPos(7)
			C_HBA010T_PAY_GRD2 = iCurColumnPos(8)
			C_HAA010T_ENTR_DT = iCurColumnPos(9)
			C_HAA010T_RETIRE_DT = iCurColumnPos(10)
			C_HAA010T_SCH_SHIP = iCurColumnPos(11)
			C_HAA030T_SCHOOL_NM = iCurColumnPos(12)
			C_HAA030T_MAJOR_NM = iCurColumnPos(13)
    End Select    
End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  

	With frm1.vspdData
	
        ggoSpread.Source = frm1.vspdData
   		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	   .ReDraw = false
       .MaxCols   = C_HAA030T_MAJOR_NM + 1                                          ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:
       .MaxRows = 0
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData 
		Call GetSpreadColumnPos("A")  

         ggoSpread.SSSetDate     C_HBA010T_GAZET_DT,         "발령일", 15,2,  parent.gDateFormat
         ggoSpread.SSSetEdit     C_HBA010T_GAZET_CD,         "발령코드", 15
         ggoSpread.SSSetEdit     C_HAA010T_NAME,             "성명", 10
         ggoSpread.SSSetEdit     C_HBA010T_EMP_NO,           "사번", 11
         ggoSpread.SSSetEdit     C_HBA010T_DEPT_NM,          "발령부서", 15
         ggoSpread.SSSetEdit     C_HBA010T_ROLL_PSTN_NM,     "직위", 10
         ggoSpread.SSSetEdit     C_HBA010T_PAY_GRD1_NM,      "급호", 10
         ggoSpread.SSSetEdit     C_HBA010T_PAY_GRD2,         "호봉", 10
         ggoSpread.SSSetDate     C_HAA010T_ENTR_DT,          "입사일", 10,2,  parent.gDateFormat
         ggoSpread.SSSetDate     C_HAA010T_RETIRE_DT,        "퇴사일", 10,2,  parent.gDateFormat
         ggoSpread.SSSetEdit     C_HAA010T_SCH_SHIP,         "학력", 15
         ggoSpread.SSSetEdit     C_HAA030T_SCHOOL_NM,        "학교", 20
         ggoSpread.SSSetEdit     C_HAA030T_MAJOR_NM,         "전공", 20

        Call ggoSpread.SSSetColHidden(C_HAA010T_SCH_SHIP,C_HAA010T_SCH_SHIP,True)	
        Call ggoSpread.SSSetColHidden(C_HAA030T_SCHOOL_NM,C_HAA030T_SCHOOL_NM,True)	
        Call ggoSpread.SSSetColHidden(C_HAA030T_MAJOR_NM,C_HAA030T_MAJOR_NM,True)	        

       Call SetSpreadLock 
	   .ReDraw = true
    
    End With
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
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

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
    
    Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call  FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
   
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
    
     frm1.txtEmp_no.focus 
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
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    
    FncQuery = False                                                            '☜: Processing is NG
    Err.Clear                                                                   '☜: Protect system from crashing
     ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     
    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    If  txtEmp_no_Onchange() then
        Exit Function
    End If

    If txtGazet_cd_OnChange() Then
        Exit Function
    End If

    If Not( ValidDateCheck(frm1.txtGazet_start_dt, frm1.txtGazet_end_dt)) Then
        Exit Function
    End If

    Call MakeKeyStream("X")
    Call  DisableToolBar( parent.TBC_QUERY)
	If DBQuery()=False Then
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
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call  ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call  ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
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

    Call  DisableToolBar( parent.TBC_QUERY)
	If DbDelete=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If
    
    FncDelete=  True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
     ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	 ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If
    
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
    
	With Frm1
	    
		If .vspdData.ActiveRow > 0 Then
			.vspdData.ReDraw = False
		
			 ggoSpread.Source = frm1.vspdData	
			 ggoSpread.CopyRow
			SetSpreadColor frm1.vspdData.ActiveRow, .ActiveRow
    
			.vspdData.ReDraw = True
			.vspdData.focus
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
    Dim IntRetCD
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
         ggoSpread.InsertRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
       .vspdData.ReDraw = True
    End With
    Set gActiveElement = document.ActiveElement   
End Function
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
' Name : FncPrev
' Desc : developer describe this line Called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     
    Call InitVariables														 '⊙: Initializes local global variables

    if LayerShowHide(1) =false then
		Exit Function
    end if

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "P"	                         '☆: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 
	
    FncPrev = True                                                               '☜: Processing is OK

End Function
'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     
    Call InitVariables														 '⊙: Initializes local global variables

    If LayerShowHide(1)=false then
		Exit Function
    end if

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "N"	                         '☆: Direction
    
	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 
	
    FncNext = True                                                               '☜: Processing is OK
	
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    If LayerShowHide(1)=false then
		Exit Function
    end if

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
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
    If LayerShowHide(1)=false then
		Exit Function
    end if

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
		
    If LayerShowHide(1)=false then
		Exit Function
    end if
		
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003                                '☜: Delete
	strVal = strVal & "&txtGlNo=" & Trim(frm1.txtLcNo.value)             '☜: 
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
	
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    Frm1.txtName.focus 

	Call SetToolbar("1100000000011111")												'⊙: Set ToolBar
    Call InitData()
    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
	Frm1.vspdData.focus
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Frm1.txtGlNo.value =  Frm1.txtLcNo.value  

	Call InitVariables
     ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()	
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

	        arrParam(0) = "발령코드 팝업"			' 팝업 명칭 
	        arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtGazet_cd.value		' Code Condition
	        arrParam(3) = ""'frm1.txtGazet_nm.value		' Name Cindition
	        arrParam(4) = "MAJOR_CD = " & FilterVar("H0029", "''", "S") & ""							' Where Condition
	        arrParam(5) = "발령코드"			    ' TextBox 명칭 
	
            arrField(0) = "minor_cd"					' Field명(0)
            arrField(1) = "minor_nm"				    ' Field명(1)
    
            arrHeader(0) = "코드"				' Header명(0)
            arrHeader(1) = "코드명"			    ' Header명(1)
	    Case "2"
	        arrParam(0) = "사업장코드 팝업"			        ' 팝업 명칭 
	    	arrParam(1) = "b_biz_area"						    ' TABLE 명칭 
	    	arrParam(2) = frm1.txtSect_cd.value        			' Code Condition
	    	arrParam(3) = ""'frm1.txtSect_nm.value				' Name Cindition
	    	arrParam(4) = ""                      		    	' Where Condition
	    	arrParam(5) = "사업장코드" 			            ' TextBox 명칭 
	
	    	arrField(0) = "biz_area_cd"						    	' Field명(0)
	    	arrField(1) = "biz_area_nm"    					    	' Field명(1)
    
	    	arrHeader(0) = "사업장코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "사업장명"	    		            ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case "1"
		        frm1.txtGazet_cd.focus
		    Case "2"
		        frm1.txtSect_cd.focus
        End Select
	
		Exit Function
	Else
		Call SubSetOpenPop(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SubSetOpenPop()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetOpenPop(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtGazet_cd.value = arrRet(0)
		        .txtGazet_nm.value = arrRet(1)
		        .txtGazet_cd.focus
		    Case "2"
		        .txtSect_cd.value = arrRet(0)
		        .txtSect_nm.value = arrRet(1)		
		        .txtSect_cd.focus
        End Select
	End With
End Sub
'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmp()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	arrParam(2) = lgUsrIntCd        			' Internal_cd
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus
		Exit Function
	Else
		Call SetEmp(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmp(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		.txtDept_nm.value = arrRet(2)
		.txtEmp_no.focus
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData 		
		Set gActiveElement = document.ActiveElement

		lgBlnFlgChgValue = False
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
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
   	If lgOldRow <> Row Then
		
		frm1.vspdData.Col = 1
		frm1.vspdData.Row = row
	
		lgOldRow = Row
		  		
		With frm1
		.vspdData.Row = .vspdData.ActiveRow 	

		.vspdData.Col = C_HBA010T_DEPT_NM
		.txtDept_nm.value = .vspdData.Text
		
		.vspdData.Col = C_HBA010T_ROLL_PSTN_NM
		.txtRoll_pstn.value = .vspdData.Text
		
		.vspdData.Col = C_HBA010T_PAY_GRD1_NM
		.txtPay_grd1.value = .vspdData.Text

		.vspdData.Col = C_HBA010T_PAY_GRD2
		.txtPay_grd2.value = .vspdData.Text

		.vspdData.Col = C_HAA010T_SCH_SHIP
		.txtSch_ship.value = .vspdData.Text

		.vspdData.Col = C_HAA030T_SCHOOL_NM
		.txtSchool_nm.value = .vspdData.Text

		.vspdData.Col = C_HAA030T_MAJOR_NM
		.txtMajor_nm.value = .vspdData.Text
		End With   
	End If
       frm1.vspdData.Row = Row
End Sub
'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And  gMouseClickStatus = "SPC" Then
           gMouseClickStatus = "SPCR"
        End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
	If NewRow <= 0 Or NewCol < 0 Then
		Exit Sub
	End If
	
		frm1.vspdData.Col = 1
		frm1.vspdData.Row = NewRow
	
		With frm1

		.vspdData.Col = C_HBA010T_DEPT_NM
		.txtDept_nm.value = .vspdData.Text
		
		.vspdData.Col = C_HBA010T_ROLL_PSTN_NM
		.txtRoll_pstn.value = .vspdData.Text
		
		.vspdData.Col = C_HBA010T_PAY_GRD1_NM
		.txtPay_grd1.value = .vspdData.Text

		.vspdData.Col = C_HBA010T_PAY_GRD2
		.txtPay_grd2.value = .vspdData.Text

		.vspdData.Col = C_HAA010T_SCH_SHIP
		.txtSch_ship.value = .vspdData.Text

		.vspdData.Col = C_HAA030T_SCHOOL_NM
		.txtSchool_nm.value = .vspdData.Text

		.vspdData.Col = C_HAA030T_MAJOR_NM
		.txtMajor_nm.value = .vspdData.Text
		End With   
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

'========================================================================================================
'   Event Name : txtEmp_no_Onchange             
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal

	frm1.txtName.value = ""

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtEmp_no.value = ""
    Else
	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			ggoSpread.Source = Frm1.vspdData    
			ggoSpread.ClearSpreadData             
            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
			txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
        End if 
    End if
End Function
'======================================================================================================
'   Event Name : txtGazet_cd_OnChange
'   Event Desc : 발령코드가 변경될 경우 
'=======================================================================================================
Function txtGazet_cd_OnChange()
    Dim IntRetCd
    dim count
    On Error Resume Next                                                          '☜: If process fails
        IntRetCD =  CommonQueryRs(" minor_cd,minor_nm "," b_minor "," major_cd=" & FilterVar("H0029", "''", "S") & " And minor_cd =  " & FilterVar(frm1.txtGazet_cd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
        If IntRetCD=False And Trim(frm1.txtGazet_cd.Value)<>""  Then
            Call  DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
            frm1.txtGazet_nm.Value=""
		    txtGazet_cd_OnChange = True            
        Else
			count= parent.CountStrings(lgF0, Chr(11))
			If  count > 1 Then                         ' 같은명일 경우 pop up
				Call FncOpenPopup(1)
				txtGazet_cd_OnChange = True            				
			Else
				frm1.txtGazet_nm.Value=Trim(Replace(lgF1,Chr(11),""))
			End If
        End If
End Function

'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtGazet_start_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtGazet_start_dt.Action = 7
        frm1.txtGazet_start_dt.focus
    End If
End Sub

Sub txtGazet_end_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtGazet_end_dt.Action = 7
        frm1.txtGazet_end_dt.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtGazet_start_dt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtGazet_start_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub
'=======================================================================================================
'   Event Name : txtGazet_End_dt_Keypress(Key)
'   Event Desc : 3rd party control에서 Enter 키를 누르면 조회 실행 
'=======================================================================================================
Sub txtGazet_End_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub
'======================================================================================================
'   Event Name : txtSect_cd_OnChange
'   Event Desc : 사업장코드가 변경될 경우 
'=======================================================================================================
Function txtSect_cd_OnChange()
    Dim IntRetCd
    Dim strWhere

        strWhere = " biz_area_cd= " & FilterVar(frm1.txtSect_cd.Value, "''", "S") & ""
        IntRetCD =  CommonQueryRs(" biz_area_cd,biz_area_nm "," b_biz_area ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtSect_cd.Value)<>""  Then
            Call  DisplayMsgBox("800054","X","X","X")                         '☜ : 등록되지 않은 코드입니다.
            frm1.txtSect_cd.Value=""
            frm1.txtSect_nm.Value=""
            txtSect_cd_OnChange = False
        ElseIf  CountStrings(lgF0, Chr(11) ) > 1 Then                         ' 같은명일 경우 pop up
            Call FncOpenPopup(2)
            txtSect_cd_OnChange = True
        Else
            frm1.txtSect_nm.Value=Trim(Replace(lgF1,Chr(11),""))
            txtSect_cd_OnChange = True
        End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00 %> ></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>인사발령사항조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
                           <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
			    	    		<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    		<TD CLASS="TD6" NOWRAP><INPUT NAME="txtEmp_no"  SIZE=13 MAXLENGTH=13 ALT="사번" TYPE="Text"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp()">
                			    	    		      <INPUT NAME="txtName"  SIZE=20  MAXLENGTH=30 ALT="성명" TYPE="Text"  tag="14XXXU"></TD>
								<TD CLASS="TD5" NOWRAP>발령코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT Name="txtGazet_cd" ALT="발령코드" MAXLENGTH="10"  SIZE=10   tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: FncOpenPopup(1)">
								                      <INPUT Name="txtGazet_nm" ALT="발령명칭" MAXLENGTH="50" SIZE=20   tag="14XXXU"></TD>
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>조회기간</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h3003ma1_fpDateTime1_txtGazet_start_dt.js'></script>
													&nbsp;~&nbsp;<script language =javascript src='./js/h3003ma1_fpDateTime2_txtGazet_end_dt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP></TD>
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
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h3003ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
	                        <TR>
                            	<TD>
                            		<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
	                        				<TR>
	                        					<TD CLASS="TD5" NOWRAP>부서</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <INPUT TYPE=TEXT Name="txtDept_nm"  Size="20" MAXLENGTH="80" ALT="현재부서" Tag="24">	                        					</TD>
	                        					<TD CLASS="TD5" NOWRAP>급호</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <INPUT TYPE=TEXT Name="txtPay_grd1"  Size="20" MAXLENGTH="80" ALT="급호" Tag="24">
	                        					</TD>
	                        					<TD CLASS="TD5" NOWRAP>학력</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <INPUT TYPE=TEXT Name="txtSch_ship"  Size="20" MAXLENGTH="80" ALT="학력" Tag="24">
	                        					</TD>
	                        				</TR>
	                        				<TR>
	                        					<TD CLASS="TD5" NOWRAP>직위</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <INPUT TYPE=TEXT Name="txtRoll_pstn" Size="20" MAXLENGTH="80" ALT="직위" Tag="24">
	                        					</TD>
	                        					<TD CLASS="TD5" NOWRAP>호봉</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <INPUT TYPE=TEXT Name="txtPay_grd2"  Size="20" MAXLENGTH="80" ALT="호봉" Tag="24">
	                        					</TD>
	                        					<TD CLASS="TD5" NOWRAP>학교</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <INPUT TYPE=TEXT Name="txtSchool_nm" Size="20" MAXLENGTH="80" ALT="학교" Tag="24">
	                        					</TD>
	                        				</TR>
	                        				<TR>
	                        					<TD CLASS="TD5" NOWRAP></TD>
	                        					<TD CLASS="TD6" NOWRAP></TD>
	                        					<TD CLASS="TD5" NOWRAP></TD>
	                        					<TD CLASS="TD6" NOWRAP></TD>
	                        					<TD CLASS="TD5" NOWRAP>전공</TD>
	                        					<TD CLASS="TD6" NOWRAP>
	                        					    <INPUT TYPE=TEXT Name="txtMajor_nm"  Size="20" MAXLENGTH="80" ALT="전공" Tag="24">
	                        					</TD>
	                        				</TR>
                            		</TABLE>
                            	</TD>
                            </TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
      <TD  HEIGHT=3></TD>
    </TR>    
	<TR>
		<TD WIDTH=100% HEIGHT= <%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

