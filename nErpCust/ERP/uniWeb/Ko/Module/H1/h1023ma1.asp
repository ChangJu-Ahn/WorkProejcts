<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: 사업장별 대표부서 등록 
*  3. Program ID           	: H1023ma1
*  4. Program Name         	: H1023ma1
*  5. Program Desc         	: 기준정보관리 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/05/28
*  8. Modified date(Last)  	: 2003/06/10
*  9. Modifier (First)     	: YBI
* 10. Modifier (Last)     	: Lee SiNa
* 11. Comment              	:
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID = "H1023mb1.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 15	                                      '☜: Visble row

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                 
Dim gblnWinEvent                                                 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim IsOpenPop
Dim lsInternal_cd

Dim C_BIZ_AREA		
Dim C_BIZ_AREAPopup 
Dim C_BIZ_AREA_NM	
Dim C_DEPT_CD		
Dim C_DEPT_NM_POP	
Dim C_DEPT_NM		

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()	 
		 C_BIZ_AREA			= 1
		 C_BIZ_AREAPopup	= 2
		 C_BIZ_AREA_NM		= 3
		 C_DEPT_CD			= 4
		 C_DEPT_NM_POP		= 5
		 C_DEPT_NM			= 6  
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
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
        .MaxCols = C_DEPT_NM + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0	
        ggoSpread.ClearSpreadData
       
       Call  GetSpreadColumnPos("A")
		
		 ggoSpread.SSSetEdit	 C_BIZ_AREA,          "사업장코드", 14,,,20,2 
		 ggoSpread.SSSetButton   C_BIZ_AREAPopup
		 ggoSpread.SSSetEdit	 C_BIZ_AREA_NM,       "사업장명", 28,,,50,2 
         ggoSpread.SSSetEdit     C_DEPT_CD,           "대표부서", 14,,,20,2
         ggoSpread.SSSetButton   C_DEPT_NM_POP
         ggoSpread.SSSetEdit     C_DEPT_NM,           "대표부서명", 28,,,40
         
         Call ggoSpread.MakePairsColumn(C_BIZ_AREA,  C_BIZ_AREAPopup)
         Call ggoSpread.MakePairsColumn(C_DEPT_CD,	 C_DEPT_NM_POP)
				
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
         ggoSpread.Source = frm1.vspdData
        .vspdData.ReDraw = False
         ggoSpread.SSSetProtected    C_BIZ_AREA		, -1	, C_BIZ_AREA
         ggoSpread.SSSetProtected    C_BIZ_AREA_NM	, -1	, C_BIZ_AREA_NM
         ggoSpread.SSSetProtected    C_BIZ_AREAPopup, -1	, C_BIZ_AREAPopup        
         ggoSpread.SSSetRequired     C_DEPT_CD		, -1	, C_DEPT_CD
         ggoSpread.SSSetProtected    C_DEPT_NM		, -1	, C_DEPT_NM
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
         ggoSpread.Source = frm1.vspdData
        .vspdData.ReDraw = False
         ggoSpread.SSSetRequired	C_BIZ_AREA		, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_BIZ_AREA_NM	, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired	C_DEPT_CD		, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected	C_DEPT_NM		, pvStartRow, pvEndRow
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
            
            C_BIZ_AREA			= iCurColumnPos(1)
			C_BIZ_AREAPopup		= iCurColumnPos(2)
			C_BIZ_AREA_NM		= iCurColumnPos(3)
			C_DEPT_CD			= iCurColumnPos(4)
			C_DEPT_NM_POP		= iCurColumnPos(5)
			C_DEPT_NM			= iCurColumnPos(6)            
            
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
    
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call  FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetToolbar("1100110100010111")												'⊙: Set ToolBar
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
    
    FncQuery = False                                                            '☜: Processing is NG
    Err.Clear                                                                   '☜: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    If   ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    

    ggoSpread.ClearSpreadData
    														'⊙: Initializes local global variables
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    Call InitVariables	
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
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
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
	Dim CFlag
	Dim strStrtDt
	Dim strEndDt
	Dim strStrtDt1
	Dim strEndDt1
	Dim strStrtDt2
	Dim strEndDt2
	Dim lRow
	Dim strStrtDtType
	Dim strEndDtType

	FncSave = False                                                              '☜: Processing is NG    
	Err.Clear                                                                    '☜: Clear err status    
	 ggoSpread.Source = frm1.vspdData
	
	If  ggoSpread.SSCheckChange = False Then
	IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
		Exit Function
	End If    
	  ggoSpread.Source = frm1.vspdData
	
	  ggoSpread.Source = frm1.vspdData
	 If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	       Exit Function
	 End If
	
	 ggoSpread.Source = frm1.vspdData
	With Frm1
       For lRow = 1 To .vspdData.MaxRows
           .vspdData.Row = lRow
           .vspdData.Col = 0
           if   .vspdData.Text =  ggoSpread.InsertFlag OR .vspdData.Text =  ggoSpread.UpdateFlag then
				.vspdData.Col = C_BIZ_AREA_NM
				 if .vspdData.Text = "" then
					Call  DisplayMsgBox("970000","X","사업장코드","X")
					.vspdData.focus
					
       	            exit function
				 end if 
				
				.vspdData.Col = C_DEPT_NM
				 if .vspdData.Text = "" then
					Call  DisplayMsgBox("970000","X","부서코드","X")
					.vspdData.focus
					
       	            exit function
				 end if 
				 
            end if
        next

    end with
	    
	Call MakeKeyStream("X")
	Call  DisableToolBar( parent.TBC_SAVE)
	IF DBSAVE =  False Then
	      Call  RestoreToolBar()
	      Exit Function
	End If
	
	FncSave = True               
    
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
    Call Initdata()
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

    If   LayerShowHide(1) = False Then
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
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
    Dim strCloseDate
    Dim strTemp
	
    Err.Clear                                                                    '☜: Clear err status
		
    DbSave = False														         '☜: Processing is NG

    If   LayerShowHide(1) = False Then
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
        ggoSpread.Source = frm1.vspdData
        
       For lRow = 1 To .vspdData.MaxRows
           
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case  ggoSpread.InsertFlag                                      '☜: Update
                                                           strVal = strVal & "C" & parent.gColSep
                                                           strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_BIZ_AREA           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD            : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                           strVal = strVal & "U" & parent.gColSep
                                                           strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_BIZ_AREA           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD            : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case  ggoSpread.DeleteFlag                                      '☜: Delete
                                                           strDel = strDel & "D" & parent.gColSep
                                                           strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_BIZ_AREA           : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                   
           End Select
       Next
	
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

   End With

   Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
    If IntRetCD = vbNo Then													'------ Delete function call area ------ 
	Exit Function	
    End If    
    
     Call  DisableToolBar( parent.TBC_DELETE)					'Query 버튼을 disable시킴 
     If DBDelete = False Then
  	Call  RestoreToolBar()
	Exit Function
    End If    
    
    FncDelete = True                                                        '⊙: Processing is OK

End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim IRow

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	
    IF  Frm1.vspdData.MaxRows > 0 then
        Call SetToolbar("1100111100111111")												'⊙: Set ToolBar
    else
        Call SetToolbar("1100110100111111")
    end if

    Call InitData()
    Call  ggoOper.LockField(Document, "Q")

    Set gActiveElement = document.ActiveElement   
    lgBlnFlgChgValue = False
    frm1.vspdData.focus

End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
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
	Call MainNew()	
End Function

'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(Byval iWhere, ByVal iRow)

	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtDept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
	Else 'spread
	    With frm1.vspdData
	        .Row = iRow
	 	    .Col = C_DEPT_CD
		    arrParam(0) = .Text			    ' Grid에서 누른 경우 Code Condition
		End With
	End If

   	arrParam(1) = ""
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  	
		
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
        frm1.vspdData.Col = C_DEPT_CD                         'spread
		frm1.vspdData.action =0
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
       	 ggoSpread.Source = frm1.vspdData
         ggoSpread.UpdateRow iRow
	End If	
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
               .txtDept_cd.value = arrRet(0)
               .txtDept_nm.value = arrRet(1)
               
             Case "1"  
             Case Else
               .vspdData.Col = C_DEPT_NM                         'spread
               .vspdData.Text = arrRet(1)
               .vspdData.Col = C_DEPT_CD                         'spread
               .vspdData.Text = arrRet(0)
			   .vspdData.action =0
        End Select
	End With
End Function       		


Function OpenPopUp(Byval IRow, Byval Part)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim intRetCd
	Dim ArgMt
	
	
    With Frm1
    	    
	    If IsOpenPop = True Then Exit Function
	    IsOpenPop = True		
	    
	    Select Case Part		                         
            Case "BIZ_AREA"            	 
                arrParam(0) = "사업장 팝업"			        ' 팝업 명칭 
		        arrParam(1) = "B_BIZ_AREA"					    ' TABLE 명칭 
		    
		        If IRow > 0 Then
		        	.vspdData.Row = IRow
		            .vspdData.Col = C_BIZ_AREA     
		        
		            If Trim(.vspdData.Text) = "" Then    		
		                arrParam(2) = "" 'Trim(Replace(lgF0,Chr(11),""))
		                arrParam(3) = ""
		            Else
		              	arrParam(2) = Trim(.vspdData.Text)
		    	        arrParam(3) = ""
		            End If
		    
		            arrParam(4) = " "
		            arrParam(5) = "사업장코드"  			                        ' TextBox 명칭 
		    
		            ArgMt = ""
		        
		            arrField(0) = "BIZ_AREA_CD"
	                arrField(1) = "BIZ_AREA_NM"    
	                        
	                arrHeader(0) = "사업장코드"
	                arrHeader(1) = "사업장명"	        
		        End If 	
        End Select

        arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	   	                  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
        	   
	    IsOpenPop = False
	
	    If arrRet(0) = "" Then
			Frm1.vspdData.Col = C_BIZ_AREA     			                ' Code Condition        	    	        
			Frm1.vspdData.Action = 0 ' go to 
	    	Exit Function
	    Else
	    	Call SetCode(arrRet, IRow, Part)
	    	
	    	If IRow > 0 Then	            
	    	   	 ggoSpread.Source = frm1.vspdData
           		 ggoSpread.UpdateRow IRow
      		End If
	    End If	
     End With
End Function

'======================================================================================================
'	Name : SetCode
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetCode(arrRet, IRow, Part)

	With Frm1	
		If IRow > 0 Then
			SELECT CASE Part
			  Case  "BIZ_AREA"
			    .vspdData.Row = IRow
				.vspdData.Col = C_BIZ_AREA     			                ' Code Condition        
				.vspdData.Value = arrRet(0)
				
				.vspdData.Col = C_BIZ_AREA_NM     			                ' Code Condition        
				.vspdData.Value = arrRet(1)								' Name Cindition      
				
				.vspdData.Row = IRow
				.vspdData.Col = C_BIZ_AREA     			                ' Code Condition        	    	        
				.vspdData.Action = 0 ' go to 
			  End SELECT      
		End If
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)   
    
    frm1.vspdData.Row = Row   	
    frm1.vspdData.Col = Col
    
    If Row > 0 Then
    
	Select Case Col	
		Case C_BIZ_AREAPopup
			Call OpenPopUp(Row, "BIZ_AREA")

	    Case C_DEPT_NM_POP
            Call OpenDept(2, Row)			'spread
	End Select    
   End If         

End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)

    Dim iDx
    Dim IntRetCd
    Dim ArgMt
    Dim iCodeArr 
    Dim iNameArr
    Dim IncurRow
    Dim strBas, strDept_nm
    
     ggoSpread.Source = frm1.vspdData
    
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    With frm1.vspdData
    		Select Case Col
    			Case C_BIZ_AREA
    				ArgMt = ""
    			 	.Col = C_BIZ_AREA
    				ArgMt = Trim(.Text)
    				
    				IF Trim(ArgMt) = "" Then
				    	.Col =  C_BIZ_AREA_NM
						.Value = ""
						.Action = 0
                    Else
    					Call  CommonQueryRs(" BIZ_AREA_NM"," B_BIZ_AREA ","BIZ_AREA_CD =  " & FilterVar(ArgMt , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    					
    					IF Len(Trim(lgF0)) = 0 or Trim(lgF0) = "" Then
							Call  DisplayMsgBox("970000","X","사업장코드","X")
    						'.Value = ""
						    .Col =  C_BIZ_AREA_NM
						    .Value = ""
						    .Action = 0
					    Else
    						.Col =  C_BIZ_AREA_NM
    						.Value = Trim(Replace(lgF0,Chr(11),""))  				
					    End IF
				    End If
    				
    			Case C_DEPT_CD    			
    				ArgMt = ""
    			 	.Col = C_DEPT_CD
    				ArgMt = Trim(.Text)

    				IF Trim(ArgMt) = "" Then
				    	.Col = C_DEPT_NM
						.Value = ""
				    	.Action = 0 ' go to
                    Else
                        strBas =  UniConvDateAToB("<%=GetSvrDate%>",  parent.gServerDateFormat,  parent.gDateFormat)
                        IntRetCd =  FuncDeptName(Trim(ArgMt), UNIConvDate(strBas),lgUsrIntCd,strDept_nm,lsInternal_cd)

                        If  IntRetCd < 0 then
                            If  IntRetCd = -1 then
								Call  DisplayMsgBox("970000","X","부서코드","X")
								.Col = C_DEPT_NM
								.Value = ""

                            Else
                                Call  DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
                            End if
                            lsInternal_cd = ""
                        Else           
    						.Col =  C_DEPT_NM
    						.Value = strDept_nm			
                        End if
				    End If
    		End Select 
    End With

     ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

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
'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
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

Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)
End Sub

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
			<TABLE <%=LR_SPACE_TYPE_40%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>사업장별대표부서등록</font></td>
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
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
		         	<TD WIDTH=100% HEIGHT=* valign=top>
		                <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=100% WIDTH=100%> 
					                <script language =javascript src='./js/h1023ma1_vaSpread_vspdData.js'></script>
					          	</TD>
					        </TR>
				        </TABLE>
		            </TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
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

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

