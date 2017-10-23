<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: 인사/급여관리 
*  2. Function Name        	: 메뉴 사용권한 관리 
*  3. Program ID           	: H41020ma1
*  4. Program Name         	: 메뉴 사용권한 관리 
*  5. Program Desc         	: Multi Sample
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/04/18
*  8. Modified date(Last)  	: 2003/06/10
*  9. Modifier (First)     	: Yoon Suck Kyu
* 10. Modifier (Last)      	: Lee SiNa
* 11. Comment              	:
=======================================================================================================-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "h1021mb1.asp"												'비지니스 로직 ASP명 
Const C_SHEETMAXROWS    = 24	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop
Dim FncCancelFlag
Dim vspdDataChangeFlag

Dim C_MnuID
Dim C_MnuIDPopup
Dim C_Mnunm
Dim C_HMnuID
Dim C_Pkg_Auth_YN
Dim C_Auth_YN
Dim C_Pkg_YN

FncCancelFlag = False
vspdDataChangeFlag = False

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_MnuID		= 1
	 C_MnuIDPopup	= 2
	 C_Mnunm		= 3
	 C_HMnuID		= 4
	 C_Pkg_Auth_YN	= 5
	 C_Auth_YN		= 6
	 C_Pkg_YN		= 7
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
Sub MakeKeyStream(pRow)   
    lgKeyStream = Frm1.txtMnuID.Value & parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtAuthYN.Value & parent.gColSep    
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iNameArr    
 
    Call  CommonQueryRs("MINOR_NM","B_MINOR","MAJOR_CD = " & FilterVar("A1020", "''", "S") & " ORDER BY MINOR_NM DESC",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr = lgF0    
    Call  SetCombo2(frm1.txtAuthYN,iNameArr, iNameArr,Chr(11))            ''''''''DB에서 불러 condition에서    
 
End Sub

Sub InitComboBox2()
    Dim iNameArr    
 
    Call  CommonQueryRs("MINOR_NM","B_MINOR","MAJOR_CD = " & FilterVar("A1020", "''", "S") & " ORDER BY MINOR_NM DESC",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr = lgF0        
    
     ggoSpread. SetCombo Replace(iNameArr,Chr(11),vbTab), C_Pkg_Auth_YN        ''''''''DB에서 불러 gread에서 
     ggoSpread. SetCombo Replace(iNameArr,Chr(11),vbTab), C_Auth_YN    
 
End Sub

'========================================================================================================
' Name : txtAuthYN_type_OnChange()
' Desc : CBO가 변화할때 
'========================================================================================================
Sub txtAuthYN_type_OnChange()
     lgBlnFlgChgValue = True
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub
'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
   	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
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
        .MaxCols = C_Pkg_YN + 1													<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0	
        ggoSpread.ClearSpreadData
       
		Call  GetSpreadColumnPos("A")

		 ggoSpread.SSSetEdit	 C_MnuID,        "메뉴ID", 23,,,15,2    
		 ggoSpread.SSSetButton   C_MnuIDPopup
		 ggoSpread.SSSetEdit	 C_Mnunm,        "메뉴명", 30,,,40,2
		 ggoSpread.SSSetEdit	 C_HMnuID,       "메뉴ID", 10,,,15,2        
		 ggoSpread.SSSetCombo    C_Pkg_Auth_YN,  "자료권한(표준)", 20,2
		 ggoSpread.SSSetCombo 	 C_Auth_YN,      "자료권한(사용자)", 20,2    
		 ggoSpread.SSSetEdit     C_Pkg_YN,		 "추가개발여부", 20,2
		 
		 Call ggoSpread.MakePairsColumn(C_MnuID,  C_MnuIDPopup)
		 Call ggoSpread.SSSetColHidden(C_HMnuID,  C_HMnuID, True)
		    
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
         
     ggoSpread.SpreadLock	  C_MnuID,        -1, C_MnuID
     ggoSpread.SpreadLock	  C_MnuIDPopup,   -1, C_MnuIDPopup     
     ggoSpread.SpreadLock	  C_Mnunm,        -1, C_Mnunm
     ggoSpread.SpreadLock	  C_Pkg_Auth_YN,  -1, C_Pkg_Auth_YN
     ggoSpread.SpreadLock     C_Auth_YN,      -1, C_Auth_YN
     ggoSpread.SpreadLock     C_Pkg_YN,       -1, C_Pkg_YN
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
     ggoSpread.SSSetRequired     C_MnuID		, pvStartRow, pvEndRow
     ggoSpread.SSSetProtected    C_Mnunm		, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired     C_Pkg_Auth_YN	, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired     C_Auth_YN		, pvStartRow, pvEndRow
     ggoSpread.SSSetRequired     C_Pkg_YN		, pvStartRow, pvEndRow
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
			
			C_MnuID			= iCurColumnPos(1)
			C_MnuIDPopup	= iCurColumnPos(2)
			C_Mnunm			= iCurColumnPos(3)
			C_HMnuID		= iCurColumnPos(4)
			C_Pkg_Auth_YN	= iCurColumnPos(5)
			C_Auth_YN		= iCurColumnPos(6)
			C_Pkg_YN		= iCurColumnPos(7)         
            
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

    Call InitVariables                                                              'Initializes local global variables
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitComboBox
    Call InitComboBox2
    Call SetToolbar("1100110100001111")										        '버튼 툴바 제어 
    frm1.txtMnuID.focus               
    
    Call CookiePage(0)
   
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
  
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

     ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    ggoSpread.ClearSpreadData
														
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables
                                             '⊙: Initializes local global variables
    Call MakeKeyStream("X")

    Call SetSpreadLock	  '=====>v표시 
    
    Call  DisableToolBar( parent.TBC_QUERY)					'Query 버튼을 disable시킴 
	If DBQuery = False Then
		Call  RestoreToolBar()

		Exit Function
	End If
   
    FncQuery = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD,i
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear  
    
    FncDelete = True                                                            '☜: Processing is OK
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
    If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If  

    If Not chkField(Document, "2") Then									         '☜: This function check required field
       Exit Function
    End If

    For lRow = 1 To Frm1.vspdData.MaxRows
        Frm1.vspdData.Row = lRow
        Frm1.vspdData.Col = 0
        
		'2003.07.18 by sbk 추가,수정된 건만 메뉴ID Check하도록 함.
		If frm1.vspdData.Text = ggoSpread.InsertFlag OR frm1.vspdData.Text = ggoSpread.UpdateFlag Then
			If MnuIDvspdData_Change(lRow) Then
	 		    Exit Function
	 		End If        
	 	End If        
    Next    

    If Chk_DeleteFlag = False Then
		Exit Function
    End If
        
    Call MakeKeyStream("X")
    
    Call  DisableToolBar( parent.TBC_SAVE)
	IF DBSAVE =  False Then
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
	
	With Frm1.VspdData
           
           	.Col = C_Pkg_Auth_YN
		    .Text = "N"
		    .Col = C_Auth_YN
		    .Text = "N"
		    Call SetSpreadLockCheck(.ActiveRow,C_Pkg_YN)
		    .Col = C_Pkg_YN
		    .Text = "Y"          
           
    End With

    Set gActiveElement = document.ActiveElement   

End Function
'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    FncCancelFlag = True
     ggoSpread.Source = Frm1.vspdData	
     ggoSpread.EditUndo
    Call  initData()
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim imRow
	Dim iRow

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
        For iRow = .vspdData.ActiveRow To .vspdData.ActiveRow + imRow - 1  
			.vspdData.Row = iRow       
			.VspdData.Col = C_Pkg_Auth_YN
			.VspdData.Text = "N"
			.VspdData.Col = C_Auth_YN
			.VspdData.Text = "N"		
			.VspdData.Col = C_Pkg_YN
			.VspdData.Text = "Y"         
        Next
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
    Call parent.FncExport( parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind( parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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
    Call InitComboBox2
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
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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

    DbQuery = False    
    Err.Clear                                                                        '☜: Clear err status


	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	Dim strVal
    
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

    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
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
               Case  ggoSpread.InsertFlag
                                                            strVal = strVal & "C" & parent.gColSep
                                                            : strVal = strVal & lRow & parent.gColSep
                                                            : strVal = strVal & .txtMnuID.value & parent.gColSep
                                                            : strVal = strVal & .txtAuthYN.value & parent.gColSep                                                            
                    .vspdData.Col = C_MnuID 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Pkg_Auth_YN 	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Auth_YN 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Pkg_YN 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep                    
                    lGrpCnt = lGrpCnt + 1               
               Case  ggoSpread.UpdateFlag                                    '☜: Update
                                                            strVal = strVal & "U" & parent.gColSep
                                                            : strVal = strVal & lRow & parent.gColSep
                                                            : strVal = strVal & .txtMnuID.value & parent.gColSep
                                                            : strVal = strVal & .txtAuthYN.value & parent.gColSep
                    .vspdData.Col = C_HMnuID 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Pkg_Auth_YN           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Auth_YN 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_Pkg_YN                : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete    
                                                            strDel = strDel & "D" & parent.gColSep
                                                            : strDel = strDel & lRow & parent.gColSep
                                                            : strDel = strDel & .txtMnuID.value & parent.gColSep
                                                            : strDel = strDel & .txtAuthYN.value & parent.gColSep                                                                               
                    .vspdData.Col = C_HMnuID 	            : strDel = strDel & Trim(.vspdData.Text)  & parent.gRowSep                    
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        =  parent.UID_M0002
       .txtUpdtUserId.value  =  parent.gUsrID
       .txtInsrtUserId.value =  parent.gUsrID
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
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    Call  DisableToolBar( parent.TBC_DELETE)
	IF DBDelete =  False Then
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
    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("1100111100111111")					
	Call SetSpreadLockCheckAll()	
	frm1.vspdData.focus	
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = frm1.vspdData									'⊙: Clear Contents  Field    
    ggoSpread.ClearSpreadData
    Call InitVariables															'⊙: Initializes local global variables
	Call  DisableToolBar( parent.TBC_QUERY)
	IF DBQUERY =  False Then
		Call  RestoreToolBar()
		Exit Function
	End If
	Call SetSpreadLockCheckAll()	
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()    
End Function

'========================================================================================================
' Name : OpenCondAreaPopup()        
' Desc : developer describe this line 
'========================================================================================================
Function OpenPopUp(Byval IRow)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim intRetCd

    With Frm1
	    If IsOpenPop = True Then Exit Function
	    IsOpenPop = True		
	    arrParam(0) = "메뉴 팝업"			                        ' 팝업 명칭 
	    arrParam(1) = "Z_LANG_CO_MAST_MNU"							    ' TABLE 명칭 
	    If IRow > 0 Then	            
	        .vspdData.Row = IRow
	        .vspdData.Col = C_MnuID     
	        If Trim(.vspdData.Text) = "" Then
                Call  CommonQueryRs(" Mnu_id,Mnu_nm "," Z_LANG_CO_MAST_MNU "," Mnu_id = " & FilterVar("H", "''", "S") & "  and lang_cd= " & FilterVar(gLang, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	            arrParam(2) = Trim(Replace(lgF0,Chr(11),""))	        	        
	            arrParam(3) = ""'Trim(Replace(lgF1,Chr(11),""))	        	        
	        Else			                ' Code Condition        
	            arrParam(2) = .vspdData.Text
	            .vspdData.Col = C_Mnunm     			                ' Code Condition        
	            arrParam(3) = ""'.vspdData.Text								' Name Cindition
	        End If
	    Else
	        If Trim(.txtMnuID.value) = "" Then
	            Call  CommonQueryRs(" Mnu_id,Mnu_nm "," Z_LANG_CO_MAST_MNU "," Mnu_id = " & FilterVar("H", "''", "S") & "  and lang_cd= " & FilterVar(gLang, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	            arrParam(2) = Trim(Replace(lgF0,Chr(11),""))	        	        
	            arrParam(3) = ""'Trim(Replace(lgF1,Chr(11),""))	        	        
	        Else
	            arrParam(2) = Trim(.txtMnuID.value)      			                ' Code Condition
	            arrParam(3) = ""'.txtMnunm.value      			                ' Code Condition
	        End If	        	    								' Name Cindition	    	    
	    End If           
'	    arrParam(3) = ""								' Name Cindition	    	     	    	
	    arrParam(4) = "LANG_CD =  " & FilterVar(parent.gLang , "''", "S") & ""
	    arrParam(5) = "메뉴 ID"  			                        ' TextBox 명칭 
	
	    arrField(0) = "MNU_ID"
        arrField(1) = "MNU_NM"
    
        arrHeader(0) = "메뉴ID"
        arrHeader(1) = "메뉴명"
    
	    arrHeader(2) = ""	    							            ' Header명(1)	       
    
            arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    	    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
        	
	    IsOpenPop = False
	
	    If arrRet(0) = "" Then
			If IRow > 0 Then	            
			    Frm1.vspdData.Col = C_MnuID     			                ' Code Condition        
			    Frm1.vspdData.Action = 0 ' go to 
			Else
			    Frm1.txtMnuID.focus
			End If
	    
	    	Exit Function
	    Else
	    	Call SetCode(arrRet, IRow)
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
Sub SetCode(arrRet, IRow)

	With Frm1	
		If IRow > 0 Then	            
	        .vspdData.Row = IRow
	        .vspdData.Col = C_MnuID     			                ' Code Condition        
	        .vspdData.Value = arrRet(0)
	        .vspdData.Col = C_Mnunm     			                ' Code Condition        
	        .vspdData.Value = arrRet(1)								' Name Cindition      
	        .vspdData.Row = IRow
	        .vspdData.Col = C_MnuID     			                ' Code Condition        	    	        
	        .vspdData.Action = 0 ' go to 
        Else
		    .txtMnuID.value = Trim(arrRet(0))
		    .txtMnuNm.value = Trim(arrRet(1))
		    .txtMnuID.focus
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
	        Case C_MnuIDPopup	
	            Call OpenPopUp(Row)
		End Select    
	End If         

End Sub
'========================================================================================================
'   Event Name : vspdData_Change 
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCD1,IntRetCD2

       	Frm1.vspdData.Row = Row
       	Frm1.vspdData.Col = Col

      With frm1.VspdData
            Select Case .Col
                Case C_MnuID 
                        Call MnuIDvspdData_Change(Row)

                Case C_Auth_YN
					.Row = Row
					If Frm1.vspdData.text <> "" Then
                        IntRetCD1 = MnuCombo_Change(Row)
                        If IntRetCD1 Then
							 Call  DisplayMsgBox("970001", "x","사용자자료권한정보에 자료","x")
							 IntRetCD2 =  DisplayMsgBox("900018",  parent.VB_YES_NO,"x","x")
							If IntRetCD2 = VBNO Then
								Call FncCancel() 
								Exit Sub
							End If
                        End If
                    End If
            End Select            
        End With                

    If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
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

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And  gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
End Sub       
'==========================================================================================
'   Event Name : txtUsrID_KeyDown(KeyCode, Shift)
'   Event Desc : 조회조건부의 txtpay_yymm_dt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtUsrID_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==========================================================================================
'   Event Name : txtMnuID_KeyDown(KeyCode, Shift)
'   Event Desc : 조회조건부의 txtpay_yymm_dt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtMnuID_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==========================================================================================
'   Event Name : txtAuthYN_KeyDown(KeyCode, Shift)
'   Event Desc : 조회조건부의 txtpay_yymm_dt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtAuthYN_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'========================================================================================================
'   Event Name : MnuIDvspdData_Change(IRow)
'   Event Desc :
'========================================================================================================
Function MnuIDvspdData_Change(IRow)
    Dim IntRetCd
    Dim strBasDt 
    Dim strBasDtAdd
    Dim MnuID

    If IRow > 0 Then
        frm1.VspdData.Col = C_MnuID
        frm1.VspdData.Row = IRow
        MnuID = frm1.VspdData.Text               
        
        If MnuID <> "" Then	           
            IntRetCd =  CommonQueryRs(" Mnu_nm "," Z_LANG_CO_MAST_MNU "," Mnu_id =  " & FilterVar(MnuID, "''", "S") & " AND LANG_CD =  " & FilterVar(parent.gLang , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                
            If IntRetCd = false then
                Call  DisplayMsgBox("800480","X","X","X")	'존재하지 않는 메뉴ID 입니다 
                frm1.VspdData.Col = C_Mnunm
                frm1.VspdData.Row = IRow
                frm1.VspdData.Text = ""
                Frm1.vspdData.Action = 0 ' go to 
                MnuIDvspdData_Change = true
            Else
                frm1.VspdData.Col = C_Mnunm
                frm1.VspdData.Row = IRow
                frm1.VspdData.Text = Trim(Replace(lgF0,Chr(11),""))
            End if  
        Else         
	    	frm1.VspdData.Col = C_Mnunm
            frm1.VspdData.Row = IRow
            frm1.VspdData.Text = ""
        End If
    End If
End Function
'========================================================================================================
'   Event Name : txtMnuID_OnChange()
'   Event Desc :
'========================================================================================================
Function txtMnuID_OnChange()
    Dim IntRetCd
    Dim strBasDt 
    Dim strBasDtAdd    
    		
    If frm1.txtMnuID.value = "" Then
		frm1.txtMnunm.value = ""		
    Else    
        IntRetCd =  CommonQueryRs(" Mnu_nm "," Z_LANG_CO_MAST_MNU "," Mnu_id =  " & FilterVar(frm1.txtMnuID.value, "''", "S") & " AND LANG_CD =  " & FilterVar(parent.gLang , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)            
        If IntRetCd = false then
	        frm1.txtMnunm.value = ""
        Else
    		frm1.txtMnunm.value = Trim(Replace(lgF0,Chr(11),""))
        End if
    End If
        
End Function
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
'   Event Name : SetSpreadLockCheck
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub SetSpreadLockCheck(Byval IRow, Byval ICol)
    With frm1    
    .vspdData.ReDraw = False                
     ggoSpread.SSSetProtected	    ICol,IRow,IRow
    .vspdData.ReDraw = True
    
    End With
End Sub
'========================================================================================================
'   Event Name : SetSpreadLockCheckAll
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub SetSpreadLockCheckAll()
Dim i,AFlag,PFlag

    With frm1    
    .vspdData.ReDraw = False                
	For i = 1 to .vspdData.MaxRows
				
		AFlag = False
		PFlag = False
		
		.vspdData.Row = i
		.vspdData.Col = C_Pkg_Auth_YN
		If .vspdData.Text = "Y" Then
			AFlag = True
		End If		
		
		
		If AFlag Then			
			 ggoSpread.SSSetProtected 	    C_Pkg_Auth_YN,i,i
			 ggoSpread.SpreadUnLock			C_Auth_YN,i,C_Auth_YN
			 ggoSpread.SSSetRequired 	    C_Auth_YN,i,i
			 ggoSpread.SSSetProtected 	    C_Pkg_YN,i,i
		Else
			 ggoSpread.SSSetProtected 	    C_Pkg_Auth_YN,i,i
			 ggoSpread.SSSetProtected	    C_Auth_YN,i,i
			 ggoSpread.SSSetProtected 	    C_Pkg_YN,i,i
		End If
		
	Next 	
	.vspdData.ReDraw = True    
    End With
End Sub
'========================================================================================================
'   Event Name : SetSpreadLockCheckAll
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Function Chk_DeleteFlag()
	Dim RFlag,i
	RFlag = False

	With Frm1.VspdData	

	    For i = 1 To .MaxRows                                                                   '☜: Clear err status		
			.Col = 0		
			.Row = i
			
			If .Text =  ggoSpread.DeleteFlag Then		
				.Col = C_Pkg_YN
	        
				If .Text = "N" Then			          
						.Action = 0 ' go to 
						Call  DisplayMsgBox("800469","X","표준 메뉴는","X")	'삭제할수 없습니다.
						Call FncCancel() 
						RFlag = False
						Chk_DeleteFlag = RFlag
						Exit Function
				Else
						RFlag = True 
				End If
	        Else
				RFlag = True
			End If
		Next
	End With
	Chk_DeleteFlag = RFlag
	
End Function
'========================================================================================================
'   Event Name : SetSpreadLockCheckAll
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Function MnuCombo_Change(IRow)
Dim IntRetCd
	With frm1.VspdData
	.Row = IRow
	.Col = C_Auth_YN
		If .Text = "N" Then
			If IRow > 0 Then
				
				.Row = IRow
				.Col = C_MnuID		
				
				IntRetCd =  CommonQueryRs(" COUNT(MNU_ID) "," HZA010T "," Mnu_id =  " & FilterVar(.Text, "''", "S") & " AND AUTH_YN = " & FilterVar("Y", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
				If IntRetCd = false then			
					MnuCombo_Change = False			
				Else
					If Trim(Replace(lgF0,Chr(11),"")) > 0 Then
						MnuCombo_Change = True
				    Else
						MnuCombo_Change = False
				    End If
				End If
			End If
		Else
			MnuCombo_Change = False
		End If
	End With
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>메뉴사용권한관리</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
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
					    	<TD CLASS="TD5" NOWRAP>메뉴 ID</TD>
					    	<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtMnuID" SIZE=15 MAXLENGTH=15 tag="11XXXU" ALT="메뉴 ID"><IMG SRC="../../../CShared/image/btnPopup.gif"   NAME="btnMnuID" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(0)">
					    	<INPUT TYPE=TEXT NAME="txtMnuNm" SIZE=20  MAXLENGTH=40 tag="14"></TD>						
					    	<TD CLASS="TD5" NOWRAP>자료권한(표준)</TD>
					    	<TD CLASS="TD6" NOWRAP><SELECT NAME=txtAuthYN tag="11X" ALT="자료권한(표준)" CLASS=cbonormal><OPTION value=""></OPTION></SELECT></TD>
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
									<script language =javascript src='./js/h1021ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
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
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

