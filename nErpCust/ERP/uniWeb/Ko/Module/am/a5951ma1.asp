<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5951MA1
'*  4. Program Name         : 월차 기준등록 
'*  5. Program Desc         : 회계관리 / 월차 / 월차 기준등록 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/09
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : Kim Kyoung-Ho
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/IncSvrMain.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>



<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'========================================================================================================
Const BIZ_PGM_ID = "a5951mb1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Dim C_REGCD   
Dim C_REGNM   
Dim C_USEYN   
Dim C_RATE    
Dim C_ACCT    
Dim C_BTN     
Dim C_ACCTNM
Dim C_TRANSTYPE
Dim C_TRANSNM  


Const COOKIE_SPLIT      = 4877	                                      '☆: Cookie Split String
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
Dim lgIsOpenPop          
Dim IsOpenPop



'========================================================================================================
Sub InitSpreadPosVariables()

	 C_REGCD   = 1                                                 'Column constant for Spread Sheet 
	 C_REGNM   = 2															
	 C_USEYN   = 3
	 C_RATE    = 4
	 C_ACCT    = 5
	 C_BTN     = 6
	 C_ACCTNM  = 7
	 C_TRANSTYPE = 8
	 C_TRANSNM  = 9
End Sub

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode		= Parent.OPMD_CMODE
	lgBlnFlgChgValue	= False
	lgIntGrpCount		= 0
    lgStrPrevKey		= ""
    lgPageNo			= ""
    lgSortKey			= 1
		
End Sub

'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub CookiePage(Kubun)
End Sub

'========================================================================================================
Sub MakeKeyStream(pRow)
End Sub        


'========================================================================================================
Sub InitComboBox()

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    ggoSpread.Source = frm1.vspdData    
    ggoSpread.SetCombo  "Y" & vbtab & "N" , C_USEYN

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub InitData()

End Sub

'========================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	ggoSpread.Source = frm1.vspdData
       ggoSpread.Spreadinit "V20030318", ,parent.gAllowDragDropSpread
	With frm1.vspdData
	
       .MaxCols = C_TRANSNM + 1                                                      ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:
    

	   '.Col = C_TypeCd                                                              '
       '.ColHidden = True                                                            '

       ggoSpread.Source= frm1.vspdData
		ggoSpread.ClearSpreadData

	   .ReDraw = false
	
       Call AppendNumberPlace("6","2","2")
       Call GetSpreadColumnPos("A")

       ggoSpread.SSSetEdit  C_REGCd , "월차Code"         ,16,   ,, 5,2
       ggoSpread.SSSetEdit  C_REGNm , "월차손익명"       ,24,   ,, 60    
       ggoSpread.SSSetCombo C_USEYN , "사용여부"		 ,10       
       ggoSpread.SSSetFloat C_RATE  , "대손상각율"		 ,20, 6,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"P"
	   ggoSpread.SSSetEdit  C_ACCT  , "계정코드"         ,16,,, 20,2
       ggoSpread.SSSetButton C_BTN
       ggoSpread.SSSetEdit  C_ACCTNM, "계정코드명"       ,24,,, 20,2
       ggoSpread.SSSetEdit	C_TRANSTYPE,	"거래유형코드",			15,		,		,	20,		2
       ggoSpread.SSSetEdit	C_TRANSNM,		"거래유형명",			20,		,		,	50
       
       
'       Call ggoSpread.MakePairsColumn(C_REGCd,C_REGNm)
       Call ggoSpread.MakePairsColumn(C_ACCT,C_BTN)
	   .ReDraw = true
	
       Call SetSpreadLock        
    
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.SSSetProtected   C_REGCd, -1, -1
      ggoSpread.SSSetProtected    C_REGNm, -1, -1
      ggoSpread.SSSetRequired	 C_USEYN, -1, -1
	  ggoSpread.SSSetProtected   C_RATE , -1, -1	  
	  ggoSpread.SSSetProtected   C_ACCT , -1, -1
	  ggoSpread.SSSetProtected   C_BTN , -1, -1
	  ggoSpread.SSSetProtected   C_ACCTNM , -1, -1 
	  ggoSpread.SSSetProtected   C_TRANSTYPE , -1, -1 
	  ggoSpread.SSSetProtected   C_TRANSNM , -1, -1 
	  ggoSpread.SpreadLock	.vspdData.MaxCols, -1,.vspdData.MaxCols                                                                                                      
    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.SSSetProtected     C_REGCd	, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected     C_REGNm	, pvStartRow, pvEndRow
      ggoSpread.SSSetRequired      C_USEYN	, pvStartRow, pvEndRow            
      ggoSpread.SSSetProtected     C_RATE	, pvStartRow, pvEndRow
      ggoSpread.SSSetProtected     C_ACCT	, pvStartRow, pvEndRow      
      ggoSpread.SSSetRequired      C_ACCTNM , pvStartRow, pvEndRow 
      ggoSpread.SSSetProtected     C_TRANSTYPE  , pvStartRow, pvEndRow      
      ggoSpread.SSSetProtected     C_TRANSNM	, pvStartRow, pvEndRow           
            
    .vspdData.ReDraw = True
    
    End With
End Sub



'======================================================================================================
' Name : SetSpreadColor1                           SELECT 후 수정가능한 컬럼만.조회 
' Desc : This method set color and protect in spread sheet celles
'======================================================================================================


Sub SetSpreadColor1()    
    With frm1
    
    .vspdData.ReDraw = False
	  ggoSpread.SpreadUnLock	 C_RATE, 8 , C_RATE, 8	  
      ggoSpread.SSSetRequired	 C_RATE , 8, 8
      
      ggoSpread.SpreadUnLock	 C_ACCT, 1,C_ACCT, 2      
      ggoSpread.SSSetRequired    C_ACCT , 1, 2  
      
      ggoSpread.SpreadUnLock	 C_BTN, 1,C_BTN, 2                 
    .vspdData.ReDraw = True
    
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

			C_REGCD			= iCurColumnPos(1)
			C_REGNM			= iCurColumnPos(2)
			C_USEYN			= iCurColumnPos(3)    
			C_RATE			= iCurColumnPos(4)
			C_ACCT			= iCurColumnPos(5)
			C_BTN			= iCurColumnPos(6)
			C_ACCTNM		= iCurColumnPos(7)
			C_TRANSTYPE		= iCurColumnPos(8)
			C_TRANSNM		= iCurColumnPos(9)
			
    End Select    
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	call InitData()
	Call SetSpreadColor1()
End Sub


'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
    
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
            
    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal
	
	
	Call SetToolbar("1100000000001111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitComboBox
'	Call CookiePage (0)                                                              '☜: Check Cookie
			
End Sub
	
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                           '⊙: Initializes local global variables
    'Call SetDefaultVal
    'Call MakeKeyStream("X")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If DbQuery = False Then                                                      '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                              '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False
    Err.Clear
    Set gActiveElement = document.ActiveElement
    FncNew = True
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False
    Err.Clear
    Set gActiveElement = document.ActiveElement
    FncDelete = True
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then                                      '☜:match pointer
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' Clear key field
	'---------------------------------------------------------------------------------------------------- 

	With Frm1.VspdData
           .Col  = C_MAJORCD
           .Row  = .ActiveRow
           .Text = ""
    End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    
    On Error Resume Next
    
    FncInsertRow = False														 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	imRow = AskSpdSheetAddRowcount()
	
	If imRow = "" Then
		Exit function
	End If
		
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1
       .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    Set gActiveElement = document.ActiveElement
    
    IF Err.number = 0 Then
	    FncInsertRow = True                                                          '☜: Processing is OK
	End If
	
	
End Function

'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function


'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(Parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function

'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		             '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function DbQuery()
	Dim strVal

    Err.Clear                                                                    '☜: Clear err status
    DbQuery = False                                                              '☜: Processing is NG

      if LayerShowHide(1) = false then                                                        '☜: Show Processing Message
		exit function
	end if
    
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream               '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgPageNo=" & lgPageNo
    End With
		
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	
    Err.Clear                                                                    '☜: Clear err status
    DbSave = False                                                               '☜: Processing is NG
      if LayerShowHide(1) = false then                                                        '☜: Show Processing Message
		exit function
	end if

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update
                                                  strVal = strVal & "C" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                    .vspdData.Col = C_MajorCd	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_MajorNm	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_MinorLen	: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                    .vspdData.Col = C_TypeCd    : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep                    
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & Parent.gColSep
                                                  strVal = strVal & lRow & Parent.gColSep
                   .vspdData.Col = C_REGCD		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_USEYN		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_RATE       : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_ACCT       : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1

               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & Parent.gColSep
                                                  strDel = strDel & lRow & Parent.gColSep
                   .vspdData.Col = C_MajorCd    : strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep									
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = Parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    DbDelete = True                                                              '☜: Processing is OK
End Function


'========================================================================================================
Sub DbQueryOk()
	
    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	Call SetToolbar("1100100100011111")                                              '☆: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitData()
	Call ggoOper.LockField(Document, "Q")
	
'	Call SetSpreadLock1
	Call SetSpreadColor1
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Sub
	

'========================================================================================================
Sub DbSaveOk()
    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field    
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    DBQuery()
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'===========================================================================
' Function Name : OpenCost
' Function Desc : OpenCode Reference Popup
'===========================================================================
Function OpenCost(strCode, iWhere, Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	        arrParam(0) = "계정팝업"		    	    <%' 팝업 명칭 %>
	    	arrParam(1) = "A_ACCT "					' TABLE 명칭 
			frm1.vspdData.Col = C_ACCT
	    	arrParam(2) =  frm1.vspdData.Text	                        ' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = ""							 		' Where Condition
	    	arrParam(5) = "계정코드"		   				    ' TextBox 명칭 

	    	arrField(0) = "ACCT_CD "		                ' Field명(0)
	    	arrField(1) = "ACCT_NM"    						' Field명(1)

	    	arrHeader(0) = "계정코드"		        		' Header명(0)
	    	arrHeader(1) = "계정코드명"	      				' Header명(1)



    arrParam(3) = ""
	arrParam(0) = arrParam(5)								    ' 팝업 명칭 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_ACCT,frm1.vspdData.ActiveRow ,"M","X","X")
		Exit Function
	Else
		Call SetCost(arrRet, iWhere, Row)
	End If

End Function

'------------------------------------------  SetSItemDC()  --------------------------------------------------
'	Name : SetCost()
'	Description : OpenSItemDC Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetCost(arrRet, iWhere, Row)


	With frm1
        .vspdData.Row = Row
		Select Case iWhere
		    Case 1
		        .vspdData.Col = C_ACCT
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_ACCTNM
		    	.vspdData.text = arrRet(1)
		End Select

		lgBlnFlgChgValue = True

		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row
		Call SetActiveCell(.vspdData,C_ACCT,.vspdData.ActiveRow ,"M","X","X")
	End With

End Function




'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
		.Row = Row
	End With
End Sub


'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
    Dim IntRetCD,Input_ACCT,  EFlag
	
	EFlag = False
	
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) --------------------------------------------------------------	



	Frm1.vspdData.Col = C_ACCT	
	Input_ACCT = Frm1.vspdData.Text
IF (Input_ACCT = "" OR Input_ACCT= NULL) THEN
    Frm1.vspdData.Col = C_ACCTNM
	Frm1.vspdData.Text = ""

Else
	IntRetCD = CommonQueryRs( " acct_nm ", " a_acct " , " acct_cd =  " & FilterVar(Input_ACCT , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)		
	If IntRetCD = False Then		
		Call DisplayMsgBox("110100","X","X","X")
		Frm1.vspdData.Col = C_ACCT		
		Frm1.vspdData.Text = ""
		Frm1.vspdData.Col = C_ACCTNM
		Frm1.vspdData.Text = ""
		frm1.vspdData.Col = Col
		Frm1.vspdData.Action = 0
		Set gActiveElement = document.activeElement  
		EFlag = True
	Else
		Frm1.vspdData.Col = C_ACCTNM
		Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
	End If
End IF
	
	'------ Developer Coding part (End   ) --------------------------------------------------------------

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)


	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = 0
        
    If EFlag And Frm1.vspdData.Text <> ggoSpread.InsertFlag Then
		Call FncCancel()				
	End If
End Sub


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)
   
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("0000111111")
	Else
		Call SetPopupMenuItemInf("0001111111")
	End If
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData


    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
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
    
    	frm1.vspdData.Row = Row
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub



'======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Select Case Col				'추가부분을 위해..select로..
	    Case C_BTN        'Cost center
	        frm1.vspdData.Col = C_ACCT
	        Call OpenCost(frm1.vspdData.Text, 1, Row)
	End Select

End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
'3.1 SpreadSheet의 이벤트 [ScriptDragDropBlock]을 추가 
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>


<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>월차기준등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;</TD>
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
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/a5951ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPayCd"     tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

