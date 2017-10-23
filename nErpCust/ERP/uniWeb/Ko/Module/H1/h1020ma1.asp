<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: 인사/급여관리 
*  2. Function Name        	: 사용자자료권한정보등록 
*  3. Program ID           	: H41020ma1
*  4. Program Name         	: 사용자자료권한정보등록 
*  5. Program Desc         	: multi Sample
*  6. Comproxy List       	:
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
Const BIZ_PGM_ID = "h1020mb1.asp"												'비지니스 로직 ASP명 
Const C_SHEETMAXROWS    = 22	                                      '한 화면에 보여지는 최대갯수*1.5%>

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
Dim lsGetsvrDate

Dim C_MnuID
Dim C_MnuIDPopup
Dim C_Mnunm
Dim C_HMnuID
Dim C_UsrID
Dim C_UsrIDPopup
Dim C_Usrnm
Dim C_UsrIntcd
Dim C_UsrIntcdPopup
Dim C_UsrIntcdnm
Dim C_AuthYN

FncCancelFlag = False
vspdDataChangeFlag = False

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize the position
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_MnuID		 = 1
	 C_MnuIDPopup	 = 2
	 C_Mnunm		 = 3
	 C_HMnuID		 = 4
	 C_UsrID		 = 5
	 C_UsrIDPopup	 = 6
	 C_Usrnm		 = 7
	 C_UsrIntcd		 = 8
	 C_UsrIntcdPopup = 9
	 C_UsrIntcdnm	 = 10
	 C_AuthYN		 = 11 
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
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
  lsGetsvrDate = "<%=GetsvrDate%>"
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
    lgKeyStream = lgKeyStream & Frm1.txtUsrID.Value & parent.gColSep
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
        .MaxCols = C_AuthYN + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0
        ggoSpread.ClearSpreadData
       
       Call  GetSpreadColumnPos("A")
    
     ggoSpread.SSSetEdit    C_MnuID,        "메뉴ID", 13,,,15,2    
     ggoSpread.SSSetButton  C_MnuIDPopup
     ggoSpread.SSSetEdit    C_Mnunm,        "메뉴명", 20,,,40,2
     ggoSpread.SSSetEdit    C_HMnuID,        "메뉴ID", 10,,,15,2        
     ggoSpread.SSSetEdit    C_UsrID,        "사용자ID", 13,,,13,2    
     ggoSpread.SSSetButton  C_UsrIDPopup
     ggoSpread.SSSetEdit    C_Usrnm,        "사용자명", 13,,,30,2    
     ggoSpread.SSSetEdit    C_UsrIntcd,     "권한내부부서코드", 13,,,30,2
     ggoSpread.SSSetButton  C_UsrIntcdPopup
     ggoSpread.SSSetEdit    C_UsrIntcdnm,   "권한부서명", 23,,,40,2    
     ggoSpread.SSSetEdit    C_AuthYN,       "권한부여여부"  ,14,2,,1,2 
     
     Call ggoSpread.MakePairsColumn(C_MnuID,	C_MnuIDPopup)
     Call ggoSpread.MakePairsColumn(C_UsrID,	C_UsrIDPopup)
     Call ggoSpread.MakePairsColumn(C_UsrIntcd, C_UsrIntcdPopup)
     Call ggoSpread.SSSetColHidden(C_HMnuID,	C_HMnuID,		True)    
    
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
     ggoSpread.SSSetRequired     C_MnuID,        -1, C_MnuID            
     ggoSpread.SpreadLock        C_Mnunm,        -1, C_Mnunm
     ggoSpread.SpreadLock		 C_UsrID,        -1, C_UsrID
     ggoSpread.SSSetProtected    C_UsrIDPopup,   -1, C_UsrIDPopup
     ggoSpread.SpreadLock	     C_Usrnm,        -1, C_Usrnm            
     ggoSpread.SSSetRequired     C_UsrIntcd,     -1, C_UsrIntcd     
     ggoSpread.SpreadLock        C_UsrIntcdnm,   -1, C_UsrIntcdnm       
     ggoSpread.SpreadLock        C_AuthYN,       -1, C_AuthYN
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
     ggoSpread.SSSetRequired     C_MnuID,        pvStartRow, pvEndRow
     ggoSpread.SSSetProtected    C_Mnunm,        pvStartRow, pvEndRow
     ggoSpread.SSSetRequired     C_UsrID,        pvStartRow, pvEndRow
     ggoSpread.SSSetProtected    C_Usrnm,        pvStartRow, pvEndRow
     ggoSpread.SSSetRequired     C_Usrintcd,     pvStartRow, pvEndRow
     ggoSpread.SSSetProtected    C_Usrintcdnm,   pvStartRow, pvEndRow
     ggoSpread.SSSetProtected    C_AuthYN,       pvStartRow, pvEndRow
     ggoSpread.SSSetProtected	.vspdData.MaxCols,pvStartRow, pvEndRow
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
				C_UsrID			= iCurColumnPos(5)
				C_UsrIDPopup	= iCurColumnPos(6)
				C_Usrnm			= iCurColumnPos(7)
				C_UsrIntcd		= iCurColumnPos(8)
				C_UsrIntcdPopup = iCurColumnPos(9)
				C_UsrIntcdnm	= iCurColumnPos(10)
				C_AuthYN		= iCurColumnPos(11)            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    
    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

	Call  ggoOper.FormatField(Document, "2", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field    

    Call InitVariables                                                              'Initializes local global variables
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitComboBox
    Call SetDefaultVal    
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

    If txtMnuID_OnChange() Then
        Exit Function
    End If
    If txtUsrID_OnChange() Then
        Exit Function
    End If
    
    Call InitVariables                                                        '⊙: Initializes local global variables
    Call SetDefaultVal
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
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
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
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
           
                Case  ggoSpread.InsertFlag,  ggoSpread.UpdateFlag
   	                .vspdData.Col = C_Mnunm
                    
                    If .vspdData.Text = "" then
                        Call  DisplayMsgBox("970029","X","메뉴ID","X")     '
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    Else
                    End if 

                    .vspdData.Col = C_Usrnm
                    
                    If .vspdData.Text = "" then
                         Call  DisplayMsgBox("210100","X","X","X")	'사용자 정보 관리에 해당하는 자료가 존재하지 않습니다.    '
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    
                    End if 
            
					.vspdData.Col = C_UsrIntcdnm
                    
                    If .vspdData.Text = "" then
                        Call  DisplayMsgBox("800012","X","X","X")	' 등록되지 않은 부서코드입니다.   '
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    
                    End if 
            End Select
        Next
	End With

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

    With frm1.VspdData
      .Col  = C_MnuID
      .Row  = .ActiveRow
      .Text = ""
      .Col  = C_Mnunm
      .Row  = .ActiveRow
      .Text = ""
      .Col  = C_HMnuID
      .Row  = .ActiveRow
      .Text = ""
      .Col  = C_UsrID
      .Row  = .ActiveRow
      .Text = ""
      .Col  = C_Usrnm
      .Row  = .ActiveRow  
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
                                                            : strVal = strVal & .txtUsrID.value & parent.gColSep
                    .vspdData.Col = C_MnuID 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_UsrID 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_UsrIntcd 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_AuthYN                : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1               
               Case  ggoSpread.UpdateFlag                                    '☜: Update
                                                            strVal = strVal & "U" & parent.gColSep
                                                            : strVal = strVal & lRow & parent.gColSep
                                                            : strVal = strVal & .txtMnuID.value & parent.gColSep
                                                            : strVal = strVal & .txtAuthYN.value & parent.gColSep
                                                            : strVal = strVal & .txtUsrID.value & parent.gColSep
                    .vspdData.Col = C_MnuID 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
                    .vspdData.Col = C_HMnuID 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
                    .vspdData.Col = C_UsrID 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                    
                    .vspdData.Col = C_UsrIntcd 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_AuthYN                : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete    
                                                            strDel = strDel & "D" & parent.gColSep
                                                            : strDel = strDel & lRow & parent.gColSep
                                                            : strDel = strDel & .txtMnuID.value & parent.gColSep
                                                            : strDel = strDel & .txtAuthYN.value & parent.gColSep
                                                            : strDel = strDel & .txtUsrID.value & parent.gColSep                    
                    .vspdData.Col = C_HMnuID 	            : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep                                                                                
                    .vspdData.Col = C_UsrID 	            : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
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
    
    
    'Call DbDelete
    Call  DisableToolBar( parent.TBC_DELETE)
	IF DbDelete =  False Then
		Call  RestoreToolBar()
		Exit Function
	End If															'☜: Delete db data
    
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
	frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables															'⊙: Initializes local global variables
	Call  DisableToolBar( parent.TBC_QUERY)
	IF DBQUERY =  False Then
		Call  RestoreToolBar()
		Exit Function
	End If

	If frm1.vspdData.MaxRows > 0 Then
        Call FixSplitColum()
	End If
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
    If frm1.vspdData.MaxRows > 0 Then
        Call FixSplitColum()
	End If
End Function

'========================================================================================================
' Name : OpenCondAreaPopup()        
' Desc : developer describe this line 
'========================================================================================================
Function OpenPopUp(Byval IRow, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim intRetCd

    With Frm1
	    If IsOpenPop = True Then Exit Function
	    IsOpenPop = True		
	    Select Case iWhere
	        Case "MNUID"	    	        
	    		arrParam(0) = "메뉴 팝업"			                        ' 팝업 명칭 
	        	arrParam(1) = "Z_LANG_CO_MAST_MNU"							    ' TABLE 명칭 
	        	If IRow > 0 Then	            
	        	    .vspdData.Row = IRow
	        	    .vspdData.Col = C_MnuID     
	        	    If Trim(.vspdData.Text) = "" Then
               	        Call  CommonQueryRs(" Mnu_id,Mnu_nm "," Z_LANG_CO_MAST_MNU ","Lang_cd =  " & FilterVar(parent.gLang , "''", "S") & " AND Mnu_id = " & FilterVar("H", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	          	        arrParam(2) = Trim(Replace(lgF0,Chr(11),""))	        	        
	        	        arrParam(3) = ""'Trim(Replace(lgF1,Chr(11),""))	        	        
	        	    Else			                ' Code Condition        
	        	        arrParam(2) = Trim(.vspdData.Text)
	        	        .vspdData.Col = C_Mnunm     			                ' Code Condition        
	        	        arrParam(3) = ""'.vspdData.Text								' Name Cindition
	        	    End If
	        	Else
	        	    If Trim(.txtMnuID.value) = "" Then
	        	        Call  CommonQueryRs(" Mnu_id,Mnu_nm "," Z_LANG_CO_MAST_MNU ","Lang_cd =  " & FilterVar(parent.gLang , "''", "S") & " AND Mnu_id = " & FilterVar("H", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	        	        arrParam(2) = Trim(Replace(lgF0,Chr(11),""))	        	        
	        	        arrParam(3) = ""'Trim(Replace(lgF1,Chr(11),""))	        	        
	        	    Else
	        	        arrParam(2) = Trim(.txtMnuID.value)      			                ' Code Condition
	        	        arrParam(3) = ""'.txtMnunm.value      			                ' Code Condition
	        	    End If	        	    								' Name Cindition	    	    
					.txtMnuID.focus	        	    
	        	End If           
	        	arrParam(4) = "LANG_CD =  " & FilterVar(parent.gLang , "''", "S") & ""
	        	arrParam(5) = "메뉴 ID"  			                        ' TextBox 명칭 
	
	        	arrField(0) = "MNU_ID"
                arrField(1) = "MNU_NM"
    
                arrHeader(0) = "메뉴ID"
                arrHeader(1) = "메뉴명"
    
	        	arrHeader(2) = ""	    							            ' Header명(1)
	       Case "USRID"
	    		arrParam(0) = "사용자정보 팝업"					            ' 팝업 명칭 
	            arrParam(1) = "z_usr_mast_rec"						            ' TABLE 명칭 
                If IRow > 0 Then	            
	        	    .vspdData.Row = IRow
	        	    .vspdData.Col = C_UsrID     			                ' Code Condition        
	        	    arrParam(2) = .vspdData.Text
	        	    .vspdData.Col = C_Usrnm     			                ' Code Condition        
	        	    arrParam(3) = ""'.vspdData.Text								' Name Cindition      
	            Else
	                arrParam(2) = frm1.txtUsrID.value 			                    ' Code Condition
	                arrParam(3) = ""'frm1.txtUsrnm.value 			                    ' Name Cindition
	                frm1.txtUsrID.focus
	        	End If
	            arrParam(4) = ""							                    ' Where Condition
	            arrParam(5) = "사용자 ID"			
	
                arrField(0) = "Usr_id"					                        ' Field명(0)
                arrField(1) = "Usr_nm"					                        ' Field명(1)
    
                arrHeader(0) = "사용자"						                ' Header명(0)
                arrHeader(1) = "사용자명"						            ' Header명(1)
                arrHeader(2) = ""	    							            ' Header명(1)	    
	    End Select	    
    
            arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    	    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
        	
	    IsOpenPop = False
	
	    If arrRet(0) = "" Then
			Select Case iWhere
			    Case "MNUID"
					If IRow > 0 Then	            			    
				        Frm1.vspdData.Col = C_MnuID     			                ' Code Condition        
				        Frm1.vspdData.Action = 0 ' go to 
				    else
						frm1.txtMnuID.focus
					end if
			    Case "USRID"
					If IRow > 0 Then	            			    
				        Frm1.vspdData.Col = C_UsrID     			                ' Code Condition        
				        Frm1.vspdData.Action = 0 ' go to 
				    else
				        frm1.txtUsrID.focus
				    end if
			End Select
	    	Exit Function
	    Else
	    	Call SetCode(arrRet, IRow, iWhere)	    	
	    	If IRow > 0 Then	            
	    	     ggoSpread.Source = frm1.vspdData
                 ggoSpread.UpdateRow IRow
            End If
	    End If	
	End With

End Function
'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(IRow)

	Dim arrRet
	Dim arrParam(2)
    Dim strBasDt 
    Dim strBasDtAdd
    Dim iWhere,intRetCd
    Dim lgUsrIntcd
    
	With Frm1
	    If IsOpenPop = True Then Exit Function

	    IsOpenPop = True

        .vspdData.Row = IRow
        .vspdData.Col = C_UsrIntcd     			                ' Code Condition        
    
        IntRetCd  = ""        
	    arrParam(0) = IntRetCd			    ' Grid에서 누른 경우 Code Condition		
        arrParam(1) =  UNIConvDateAToB(lsGetsvrDate ,parent.gServerDateFormat, parent.gDateFormat)
        iWhere = " ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate()) "
        Call  CommonQueryRs(" TOP 1 INTERNAL_CD "," B_ACCT_DEPT ",  iWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        lgUsrIntcd = Trim(Replace(lgF0,Chr(11),""))                  
        arrParam(2) = lgUsrIntcd
	
	    arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
	    	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	    	
	    IsOpenPop = False
	
	    If arrRet(0) = "" Then
			frm1.vspdData.col = C_UsrIntcd
   	        frm1.vspdData.Action = 0 ' go to 
	    	Exit Function
	    Else
	    	Call SetDept(arrRet, IRow)
	    	 ggoSpread.Source = frm1.vspdData
             ggoSpread.UpdateRow IRow
	    End If	
	End With
			
End Function

'======================================================================================================
'	Name : SetCode
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetCode(arrRet, IRow, iWhere)

	With Frm1	
		Select Case iWhere
		    Case "MNUID"
		        If IRow > 0 Then	            		        
	    	        .vspdData.Row = IRow
	    	        .vspdData.Col = C_MnuID     			                ' Code Condition        
	    	        .vspdData.Value = Trim(arrRet(0))
	    	        .vspdData.Col = C_Mnunm    			                ' Code Condition        	    	        
	    	        .vspdData.Value = Trim(arrRet(1))								' Name Cindition      
	    	        Call MnuIDvspdData_Change(IRow)  			                ' Code Condition        	    	        	    	        
	    	        .vspdData.Row = IRow	    	        
	    	        .vspdData.Col = C_MnuID   	    	        
	    	        .vspdData.Action = 0 ' go to 
                Else
		            .txtMnuID.value = Trim(arrRet(0))
		            .txtMnuNm.value = arrRet(1)		        
		        End If
		    Case "USRID"
		        If IRow > 0 Then	            
	    	        .vspdData.Row = IRow
	    	        .vspdData.Col = C_UsrID     			                ' Code Condition        
	    	        .vspdData.Value = arrRet(0)
	    	        .vspdData.Col = C_Usrnm     			                ' Code Condition        
	    	        .vspdData.Value = arrRet(1)								' Name Cindition      
	    	        .vspdData.Row = IRow
	    	        .vspdData.Col = C_UsrID     			                ' Code Condition        
	    	        .vspdData.Action = 0 ' go to 
                Else		    
		            .txtUsrID.value = arrRet(0)
		            .txtUsrnm.value = arrRet(1)		        
		        End If
        End Select
	End With
End Sub
'======================================================================================================
'	Name : SetDept
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetDept(arrRet, IRow)
    With Frm1
        .vspdData.Row = IRow
	    .vspdData.Col = C_UsrIntcd     			                ' Code Condition        
        .vspdData.Value = arrRet(2)
        .vspdData.Col = C_UsrIntcdnm     			                ' Code Condition        
        .vspdData.Value = arrRet(1)								' Name Cindition        
        .vspdData.Row = IRow
	    .vspdData.Col = C_UsrIntcd     			                ' Code Condition                   
	    .vspdData.Action = 0 ' go to 
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
	            Call OpenPopUp(Row,"MNUID")
	        Case C_UsrIDPopup
	            Call OpenPopUp(Row,"USRID")
	        Case C_UsrIntcdPopup
	            Call OpenDept(Row)
		End Select    
	End If         

End Sub
'========================================================================================================
'   Event Name : vspdData_Change 
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCD

       	Frm1.vspdData.Row = Row
       	Frm1.vspdData.Col = Col

      With frm1.VspdData
            Select Case .Col
                Case C_MnuID 
                   .Row = Row
                    If Frm1.vspdData.text <> "" Then
                        Call MnuIDvspdData_Change(Row)
                    Else                    
						Frm1.vspdData.Row = Row	        	
						Frm1.vspdData.Col = C_Mnunm                
						Frm1.vspdData.Text = ""
						Frm1.vspdData.Action = 0 						
                    End If                    
                Case C_UsrID
                    .Row  = Row
                    If Frm1.vspdData.text <> "" Then
                        Call UsrIDvspdData_Change(Row)
                    Else  
						Frm1.vspdData.Row = Row	        	
						Frm1.vspdData.Col = C_Usrnm                
						Frm1.vspdData.Text = ""
						Frm1.vspdData.Action = 0  
                    End If
                Case C_UsrIntcd
                    .Row  = Row
                    If Frm1.vspdData.text <> "" Then
                        Call UsrintcdvspdData_Change(Row)
                    Else
						Frm1.vspdData.Row = Row	        	
						Frm1.vspdData.Col = C_UsrIntcdnm                
						Frm1.vspdData.Text = ""
						Frm1.vspdData.Action = 0    
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

    Dim IntRetCd1,IntRetCd2
    Dim RValue1
    Dim strBasDt 
    Dim strBasDtAdd
    Dim MnuID

	IntRetCd1 = True
	IntRetCd2 = True

    If IRow > 0 Then
        frm1.VspdData.Row = IRow
        frm1.VspdData.Col = C_MnuID        
        MnuID = Trim(frm1.VspdData.Text)               		
        If MnuID <> "" Then
        
            IntRetCd1 =  CommonQueryRs(" Mnu_nm "," Z_LANG_CO_MAST_MNU ","Lang_cd =  " & FilterVar(parent.gLang , "''", "S") & " And Mnu_id =  " & FilterVar(MnuID, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If IntRetCd1 <> False Then
				RValue1 = lgF0
                IntRetCd2 = ChkMnu_Auth_YN(MnuID)            

                If IntRetCd2 = False then
			    	IntRetCd1 = False
                End If
			Else
				RValue1 = ""
				IntRetCd2 = True			
			End If

            If IntRetCd1 = False Then
				If IntRetCd2 = False Then
					Call  DisplayMsgBox("990016","X","X","X")	'해당 메뉴에 대한 권한이 없습니다.!
				Else
					Call  DisplayMsgBox("213700","X","X","X")	'엔터프라이즈 메뉴 정보에 해당하는 자료가 존재하지 않습니다.
			   End If

				frm1.VspdData.Row = IRow	        	
                frm1.VspdData.Col = C_Mnunm                
                frm1.VspdData.Text = ""
                Frm1.vspdData.Action = 0 ' go to 
                MnuIDvspdData_Change = true
            Else
                frm1.VspdData.Row = IRow
                frm1.VspdData.Col = C_Mnunm                
                frm1.VspdData.Text = Trim(Replace(RValue1,Chr(11),""))
                
                frm1.VspdData.Row = IRow
                frm1.VspdData.Col = C_AuthYN                
                frm1.VspdData.Text = IntRetCd2                
            End if  
        Else
  
	    	frm1.VspdData.Col = C_Mnunm
            frm1.VspdData.Row = IRow
            frm1.VspdData.Text = ""
        End If
        
    End If
    

End Function
'========================================================================================================
'   Event Name : UsrIDvspdData_Change(IRow)
'   Event Desc :
'========================================================================================================
Function UsrIDvspdData_Change(IRow)
    Dim IntRetCd
    Dim strBasDt 
    Dim strBasDtAdd
    Dim UsrID

    If IRow > 0 Then
        frm1.VspdData.Col = C_UsrID
        frm1.VspdData.Row = IRow
        UsrID = frm1.VspdData.Text               
        If UsrID <> "" Then	    	
	        IntRetCd =  CommonQueryRs(" Usr_nm "," z_usr_mast_rec "," Usr_id =  " & FilterVar(UsrID, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            
            If IntRetCd = false then
                Call  DisplayMsgBox("210100","X","X","X")	'사용자 정보 관리에 해당하는 자료가 존재하지 않습니다.	        	
	        	
                frm1.VspdData.Row = IRow
                frm1.VspdData.Col = C_Usrnm
                frm1.VspdData.Text = ""
                Frm1.vspdData.Action = 0 ' go to 
                UsrIDvspdData_Change = true
            Else
          
	            frm1.VspdData.Col = C_Usrnm
                frm1.VspdData.Row = IRow
                frm1.VspdData.Text = Trim(Replace(lgF0,Chr(11),""))
            End if  	    
	    Else
            frm1.VspdData.Col = C_Usrnm
            frm1.VspdData.Row = IRow
            frm1.VspdData.Text = ""
        End If        
    End If

End Function
'========================================================================================================
'   Event Name : UsrintcdvspdData_Change(Row)
'   Event Desc :
'========================================================================================================
Function UsrintcdvspdData_Change(IRow)
    Dim IntRetCd
    Dim strBasDt 
    Dim strBasDtAdd
    Dim Usrintcd,iWhere    

    If IRow > 0 Then
        frm1.VspdData.Col = C_Usrintcd
        frm1.VspdData.Row = IRow
        Usrintcd = frm1.VspdData.Text               
        
        If Usrintcd <> "" Then	    	 
            iWhere = " INTERNAL_CD =  " & FilterVar(Usrintcd, "''", "S") & " AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT WHERE ORG_CHANGE_DT < getdate())"                        
            IntRetCd =  CommonQueryRs(" DEPT_NM "," B_ACCT_DEPT ",  iWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                
            If IntRetCd = false then
                Call  DisplayMsgBox("800012","X","X","X")	' 등록되지 않은 부서코드입니다.
	        	frm1.VspdData.Col = C_Usrintcd
                frm1.VspdData.Row = IRow
                frm1.VspdData.Col = C_Usrintcdnm
                frm1.VspdData.Text = ""
                Frm1.vspdData.Action = 0 ' go to 
                UsrintcdvspdData_Change = true
            Else
                frm1.VspdData.Col = C_Usrintcdnm
                frm1.VspdData.Row = IRow
                frm1.VspdData.Text = Trim(Replace(lgF0,Chr(11),""))
            End if  
        Else    
	    	frm1.VspdData.Col = C_Usrintcdnm
            frm1.VspdData.Row = IRow
            frm1.VspdData.Text = ""
        End If
    End If
End Function

'========================================================================================================
'   Event Name : txtUsrID_OnChange()
'   Event Desc :
'========================================================================================================
Function txtUsrID_OnChange()
    Dim IntRetCd
    Dim strBasDt 
    Dim strBasDtAdd
    
    If frm1.txtUsrID.value = "" Then
	    frm1.txtUsrnm.value = ""
    Else    
        IntRetCd =  CommonQueryRs(" Usr_nm "," z_usr_mast_rec "," Usr_id =  " & FilterVar(frm1.txtUsrID.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)            
        If IntRetCd = false then
            Call  DisplayMsgBox("210100","X","X","X")	'사용자 정보 관리에 해당하는 자료가 존재하지 않습니다.	    	
	        frm1.txtUsrnm.value = ""
            frm1.txtUsrID.focus
            Set gActiveElement = document.ActiveElement                         
            txtUsrID_OnChange = true
        Else            
	    	frm1.txtUsrnm.value = Trim(Replace(lgF0,Chr(11),""))            
        End if  
    End If
    
End Function
'========================================================================================================
'   Event Name : txtUsrID_OnChange()
'   Event Desc :
'========================================================================================================
Function txtMnuID_OnChange()
    Dim IntRetCd
    Dim strBasDt 
    Dim strBasDtAdd    
    		
    If frm1.txtMnuID.value = "" Then
		frm1.txtMnunm.value = ""		
    Else    
        IntRetCd =  CommonQueryRs(" Mnu_nm "," Z_LANG_CO_MAST_MNU ","Lang_cd =  " & FilterVar(parent.gLang , "''", "S") & " And Mnu_id =  " & FilterVar(frm1.txtMnuID.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)            
        If IntRetCd = false then
            Call  DisplayMsgBox("213700","X","X","X")	'엔터프라이즈 메뉴 정보에 해당하는 자료가 존재하지 않습니다.
	        frm1.txtMnunm.value = ""
            frm1.txtMnuID.focus
            Set gActiveElement = document.ActiveElement             
            txtMnuID_OnChange = true
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
'   Event Name : ChkMnu_Auth_YN
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Function ChkMnu_Auth_YN(Byval IMnu_ID)
Dim IntRetCd
	IntRetCd =  CommonQueryRs(" Auth_YN "," HZA020t "," Mnu_id =  " & FilterVar(IMnu_ID, "''", "S") & " AND AUTH_YN = " & FilterVar("Y", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	If IntRetCd = False Then
		ChkMnu_Auth_YN = False
		Exit Function
	Else
		ChkMnu_Auth_YN = Trim(Replace(lgF0,Chr(11),""))		
	End If	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>사용자자료권한정보</font></td>
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
					    	<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtMnuID" SIZE=15 MAXLENGTH=15 tag="11XXXU" ALT="메뉴 ID"><IMG SRC="../../../CShared/image/btnPopup.gif"   NAME="btnMnuID" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(0,'MNUID')">
					    	<INPUT TYPE=TEXT NAME="txtMnuNm" SIZE=20  MAXLENGTH=40 tag="14"></TD>						
					    	<TD CLASS="TD5" NOWRAP>권한부여여부</TD>
					    	<TD CLASS="TD6" NOWRAP><SELECT NAME=txtAuthYN ALT="권한부여여부" CLASS ="cbonormal" tag="11X"><OPTION value=""></OPTION></SELECT></TD>
					    </TR>					
					    <TR>
					    	<TD CLASS="TD5" NOWRAP>사용자 ID</TD>
					    	<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtUsrID" SIZE=15 MAXLENGTH=13 tag="11XXXU" ALT="사용자 ID"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLangCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(0,'USRID')">
					    	<INPUT TYPE=TEXT NAME="txtUsrNm" SIZE=20  MAXLENGTH=40 tag="14"></TD>						
					    	<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
					    	<TD CLASS="TD6" NOWRAP></TD>
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
									<script language =javascript src='./js/h1020ma1_vaSpread_vspdData.js'></script>
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
