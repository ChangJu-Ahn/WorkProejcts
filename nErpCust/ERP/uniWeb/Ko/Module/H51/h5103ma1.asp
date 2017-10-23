<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Multi Sample
*  3. Program ID           : H5103ma1
*  4. Program Name         : H5103ma1
*  5. Program Desc         : 월공제내역조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2003/06/11
*  9. Modifier (First)     : TGS 최용철 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID      = "h5103mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "h5103mb2.asp"						           '☆: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 21		                                      '☜: Visble row

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgOldRow

Dim C_NAME
Dim C_EMP_NO
Dim C_SUB_TYPE
Dim C_SUB_CD
Dim C_SUB_AMT
Dim C_CALCU_TYPE

Dim C_NAME2
Dim C_EMP_NO2
Dim C_SUB_TYPE2
Dim C_SUB_CD2
Dim C_SUB_AMT2
Dim C_CALCU_TYPE2

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
         C_NAME			= 1										   
		 C_EMP_NO		= 2
		 C_SUB_TYPE		= 3
		 C_SUB_CD		= 4
		 C_SUB_AMT		= 5
		 C_CALCU_TYPE	= 6
    ElseIf pvSpdNo = "B" Then
         C_NAME2		= 1										   
		 C_EMP_NO2		= 2
		 C_SUB_TYPE2	= 3
		 C_SUB_CD2		= 4
		 C_SUB_AMT2		= 5
		 C_CALCU_TYPE2	= 6
    End If
    
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
	lgOldRow = 0
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call  ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat ,  parent.gServerDateType ,strYear,strMonth,strDay)

	frm1.txtsub_yymm_dt.Year=strYear
	frm1.txtsub_yymm_dt.Month=strMonth
	
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
   Dim strYear
   Dim strMonth
   Dim strBounsDt

    strYear = frm1.txtsub_yymm_dt.year
    strMonth = frm1.txtsub_yymm_dt.month
    
    If len(strMonth) = 1 then
		strMonth = "0" & strMonth
	End if

	strBounsDt = strYear & strMonth
    	   
    lgKeyStream       = strBounsDt & parent.gColSep       'You Must append one character( parent.gColSep)
	lgKeyStream       = lgKeyStream & Trim(frm1.txtname.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtemp_no.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtsub_type.value) & parent.gColSep    
    lgKeyStream       = lgKeyStream & Trim(Frm1.txtsub_cd.Value) & parent.gColSep
  
    lgKeyStream       = lgKeyStream & lgUsrIntcd & parent.gColSep  
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
    Dim dblSum
    	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData2
        frm1.vspdData2.Row = 1
        frm1.vspdData2.Col = 0
         frm1.vspdData2.Text = "합계"
        
        frm1.vspdData2.Col = C_SUB_AMT2
        frm1.vspdData2.value =  FncSumSheet(frm1.vspdData,C_SUB_AMT, 1, .MaxRows , FALSE , -1, -1, "V")
    End With
    
    With frm1
		 ggoSpread.Source = frm1.vspdData2
		.vspdData2.ReDraw = False
		 ggoSpread.SpreadLock      -1 , -1
		.vspdData2.ReDraw = True

    End With
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call  AppendNumberPlace("6","2","0")

	If pvSpdNo = "" OR pvSpdNo = "A" Then

    	Call initSpreadPosVariables("A")   

	    With frm1.vspdData
	
            ggoSpread.Source = frm1.vspdData
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    

	       .ReDraw = false

           .MaxCols   = C_CALCU_TYPE + 1                                                 ' ☜:☜: Add 1 to Maxcols
	       .Col       = .MaxCols                                                         ' ☜:☜: Hide maxcols
           .ColHidden = True                                                             ' ☜:☜:
           
           .MaxRows = 0
           ggoSpread.ClearSpreadData

         Call GetSpreadColumnPos("A")
          
         ggoSpread.SSSetEdit     C_NAME,     "성명", 18,,, 30,2        ' Lock
         ggoSpread.SSSetEdit     C_EMP_NO,   "사번", 17,,, 13,2  'Lock
         ggoSpread.SSSetEdit     C_SUB_TYPE, "공제구분", 20 ,,,50,2'구분HIDDEN
         ggoSpread.SSSetEdit     C_SUB_CD,   "공제코드", 20 ,,,50,2'구분HIDDEN
         ggoSpread.SSSetFloat    C_SUB_AMT,  "공제금액",20, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
         ggoSpread.SSSetEdit     C_CALCU_TYPE,"계산구분", 20       'Lock
       
		.ReDraw = true
	
         Call SetSpreadLock 
    
        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then
	
    	Call initSpreadPosVariables("B")

 	    With frm1.vspdData2
	
            ggoSpread.Source = frm1.vspdData2
            ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread

            .ReDraw = false

	    	.MaxCols = C_CALCU_TYPE2 + 1
	    	.Col = .MaxCols							'☜: 공통콘트롤 사용 Hidden Column
	    	.ColHidden = True
	    	
	    	.MaxRows = 0
	    	ggoSpread.ClearSpreadData

            Call GetSpreadColumnPos("B") 'sbk

			ggoSpread.SSSetEdit     C_NAME2 ,     "", 18          ' Lock
			ggoSpread.SSSetEdit     C_EMP_NO2 ,   "", 17           'Lock
			ggoSpread.SSSetEdit     C_SUB_TYPE2 , "", 20         'Lock
			ggoSpread.SSSetEdit     C_SUB_CD2 ,   "", 20         'Lock
			ggoSpread.SSSetFloat    C_SUB_AMT2 ,  "공제금액",20, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
			ggoSpread.SSSetEdit     C_CALCU_TYPE2 ,"", 20       'Lock
          
	    	.ReDraw = True
            
            Call SetSpreadLock1 
        End with
    End If
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

Sub SetSpreadLock1()
      ggoSpread.Source = frm1.vspdData2
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(lRow)
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
			
		    C_NAME			= iCurColumnPos(1)										   
			C_EMP_NO		= iCurColumnPos(2)
			C_SUB_TYPE		= iCurColumnPos(3)
			C_SUB_CD		= iCurColumnPos(4)
			C_SUB_AMT		= iCurColumnPos(5)
			C_CALCU_TYPE	= iCurColumnPos(6)
    
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_NAME2			= iCurColumnPos(1)										   
			C_EMP_NO2		= iCurColumnPos(2)
			C_SUB_TYPE2		= iCurColumnPos(3)
			C_SUB_CD2		= iCurColumnPos(4)
			C_SUB_AMT2		= iCurColumnPos(5)
			C_CALCU_TYPE2	= iCurColumnPos(6)            
            
    End Select    
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
          
    Call InitSpreadSheet("")                                                            'Setup the Spread sheet
  
    Call InitVariables                                                              'Initializes local global variables
    
    Call  ggoOper.FormatDate(frm1.txtsub_yymm_dt,  parent.gDateFormat, 2) 
    
    Call  FuncGetAuth(gStrRequestMenuID , parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
    
    frm1.txtsub_yymm_dt.focus 

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
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.ClearSpreadData
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    If  txtsub_cd_Onchange()  then
        Exit Function
    End If

    If  txtEmp_no_Onchange() then
        Exit Function
    End If

    If  txtsub_type_Onchange()  then
       Exit Function
    End If
       
    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    Call  DisableToolBar( parent.TBC_QUERY)
	If DBQuery=False Then
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
    
    FncDelete = True                                                             '☜: Processing is OK
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
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
     ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
    Call  DisableToolBar( parent.TBC_SAVE)
	If DBSave=False Then
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

    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "M"
    End If
      
    Select Case UCase(Trim(lgActiveSpd))
        Case  "M"
            If Frm1.vspdData.MaxRows < 1 Then
                Exit Function
            End If
    
	        With Frm1.vspdData
	    
	        	If .ActiveRow > 0 Then
	        		.ReDraw = False
	        	
	        		ggoSpread.Source = frm1.vspdData	
	        		ggoSpread.CopyRow
                    SetSpreadColor .ActiveRow, .ActiveRow
	        		.ReDraw = True
	        		.focus
	        	End If
	        End With
        Case  Else

            If Frm1.vspdData2.MaxRows < 1 Then
                Exit Function
            End If
    
	        With Frm1.vspdData2
	    
	        	If .ActiveRow > 0 Then
	        		.ReDraw = False
	        	
	        		ggoSpread.Source = frm1.vspdData2
	        		ggoSpread.CopyRow
                    SetSpreadColor1 .ActiveRow, .ActiveRow
	        		.ReDraw = True
	        		.focus
	        	End If
	        End With

    End Select 

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

    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "M"
    End If
      
    Select Case UCase(Trim(lgActiveSpd))
        Case  "M"
                  With Frm1
                         .vspdData.ReDraw = False
                         .vspdData.Focus
                          ggoSpread.Source = .vspdData
                          ggoSpread.InsertRow .vspdData.ActiveRow, imRow
                          SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
                         .vspdData.ReDraw = True
                  End With
        Case  Else
                  With Frm1
                         .vspdData2.ReDraw = False
                         .vspdData2.Focus
                          ggoSpread.Source = .vspdData2
                          ggoSpread.InsertRow .vspdData2.ActiveRow, imRow
                          SetSpreadColor1 .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1
                         .vspdData2.ReDraw = True
                  End With
    End Select 

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
	Call Parent.FncExport( parent.C_MULTI)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( parent.C_MULTI, False)
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
	ggoSpread.Source = frm1.vspdData2 
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
    
    If isEmpty(TypeName(gActiveSpdSheet)) Then
		Exit Sub
	Elseif	UCase(gActiveSpdSheet.id) = "VASPREAD" Then
		ggoSpread.Source = frm1.vspdData2 
		Call ggoSpread.SaveSpreadColumnInf()
	End if

End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("A")      
    ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ReOrderingSpreadData()

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("B")      
    ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ReOrderingSpreadData()
	
	'Call InitData()
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
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

	If LayerShowHide(1)=False then
		Exit Function
	End If
	
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
    If lgIntFlgMode =  parent.OPMD_UMODE Then
    Else
    End If
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
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
	
    DbSave = False                                                          
    
	If LayerShowHide(1)=False then
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
 
        Case  ggoSpread.InsertFlag                                      '☜: Update
                                           strVal = strVal & "C" & parent.gColSep
                                           strVal = strVal & lRow & parent.gColSep
                                         
             .vspdData.Col = C_NAME	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
             .vspdData.Col = C_EMP_NO	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
             .vspdData.Col = C_SUB_TYPE	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
             .vspdData.Col = C_SUB_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
             .vspdData.Col = C_SUB_AMT     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
             .vspdData.Col = C_CALCU_TYPE   : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
             lGrpCnt = lGrpCnt + 1
                    
        Case  ggoSpread.UpdateFlag                                      '☜: Update
                                           strVal = strVal & "U" & parent.gColSep
                                           strVal = strVal & lRow & parent.gColSep
                                        
             .vspdData.Col = C_NAME	     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
             .vspdData.Col = C_EMP_NO	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
             .vspdData.Col = C_SUB_TYPE	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
             .vspdData.Col = C_SUB_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
             .vspdData.Col = C_SUB_AMT     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
             .vspdData.Col = C_CALCU_TYPE   : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep   
             lGrpCnt = lGrpCnt + 1
        Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                           strDel = strDel & "D" & parent.gColSep
                                           strDel = strDel & lRow & parent.gColSep
             .vspdData.Col = C_NAME	     : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
             .vspdData.Col = C_EMP_NO	 : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep								
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
	If DBDelete=False Then
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
    Dim strVal

    lgIntFlgMode =  parent.OPMD_UMODE    
     ggoSpread.Source       = Frm1.vspdData2
    Frm1.vspdData2.MaxRows = 0

    Call MakeKeyStream("X")
	If LayerShowHide(1)=False then
		Exit Function
	End If


    strVal = BIZ_PGM_ID1 & "?txtMode="           & parent.UID_M0001                    '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & 1                             '☜: Max fetched data
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
	Call SetToolbar("1100000000011111")	 
	frm1.vspdData.focus
End Function
'========================================================================================================
' Function Name : DbQueryOk1
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk1()

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	
    Frm1.vspdData2.Col = 0
    Frm1.vspdData2.Text = "합계"
    Frm1.txtsub_yymm_dt.focus 
	Call SetToolbar("1100000000011111")												'⊙: Set ToolBar

	Frm1.vspdData.focus
    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	
    ggoSpread.Source = Frm1.vspdData
	ggoSpread.ClearSpreadData
    Call InitVariables															'⊙: Initializes local global variables
    Call  DisableToolBar( parent.TBC_QUERY)
	If DBQuery=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If
End Function
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function

'-----------------------------------------------------------------------------------------------
'	Name : openEmptName()                                                        
'	Description : Employee PopUp
'------------------------------------------------------------------------------------------------
Function openEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_No.value			' Code Condition
		arrParam(1) = ""'frm1.txtName.value            ' Name Cindition
	Else 'spread
		frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	End If
	arrParam(2) = lgUsrIntcd
	
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 
			frm1.txtEmp_no.focus
		Else 
			frm1.vspdData.Col = C_Emp_No
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		Call SetEmp(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetEmp()  ------------------------------------------------
'	Name : SetEmp()
'	Description : Employee Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------
Function SetEmp(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then 
			.txtName.value = arrRet(1)
			.txtEmp_no.value = arrRet(0)
			.txtEmp_no.focus
		Else 
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_Emp_No
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If
	End With
End Function

'========================================================================================================
' Name : OpenCondAreaPopup()      
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	  
        Case "3"
            arrParam(0) = "공제구분 팝업"			' 팝업 명칭 
	        arrParam(1) = "B_MINOR"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtsub_type.value		' Code Condition
	        arrParam(3) = ""'frm1.txtsub_type_nm.value		' Name Cindition
	        arrParam(4) = " MAJOR_CD = " & FilterVar("H0040", "''", "S") & " "        ' Where Condition
	        arrParam(5) = "공제구분"			    ' TextBox 명칭 
	
            arrField(0) = "minor_cd"					' Field명(0)
            arrField(1) = "minor_nm"				    ' Field명(1)
    
            arrHeader(0) = "공제구분코드"				' Header명(0)
            arrHeader(1) = "공제구분명"
	    
	    Case "4"
	        arrParam(0) = "공제코드 팝업"			' 팝업 명칭 
	        arrParam(1) = "HDA010T"				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtsub_cd.value		' Code Condition
	        arrParam(3) = ""'frm1.txtsub_cd_nm.value		' Name Cindition
	        arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("2", "''", "S") & " "  ' Where Condition
	        arrParam(5) = "공제코드"			    ' TextBox 명칭 
	
            arrField(0) = "ALLOW_CD"					' Field명(0)
            arrField(1) = "ALLOW_NM"				    ' Field명(1)
    
            arrHeader(0) = "공제코드"				' Header명(0)
            arrHeader(1) = "공제코드명"
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		Select Case iWhere
		    Case "3"
		        frm1.txtsub_type.focus
		    Case "4"
		        frm1.txtsub_cd.focus
        End Select
	
		Exit Function
	Else
		Call SubSetCondArea(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SetCondArea()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondArea(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		    Case "3"
		        .txtsub_type.value = arrRet(0)    
		        .txtsub_type_nm.value = arrRet(1) 
		        .txtsub_type.focus
		    Case "4"
		        .txtsub_cd.value = arrRet(0)
		        .txtsub_cd_nm.value = arrRet(1)
		        .txtsub_cd.focus
        End Select
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_OnFocus()
End Sub

'========================================================================================================
'   Event Name : vspdData2_OnFocus
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_OnFocus()
End Sub
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
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

    If  Trim(frm1.txtEmp_no.value) = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
			frm1.txtName.value = ""
            Frm1.txtEmp_no.focus 
            Set gActiveElement = document.ActiveElement
			txtEmp_no_Onchange = true
        Else
			frm1.txtName.value = strName
        End if 
    End if  
End Function

'========================================================================================================
'   Event Name : txtsub_type_OnChange             
'   Event Desc :
'========================================================================================================
Function txtsub_type_OnChange()
    Dim iDx
    Dim IntRetCd

    IF Trim(frm1.txtsub_type.value) = "" THEN
        frm1.txtsub_type.value = ""
        frm1.txtsub_type_nm.value = ""
        frm1.txtsub_type.focus       
    ELSE
        IntRetCd =  CommonQueryRs(" minor_nm "," b_minor "," major_cd = " & FilterVar("H0040", "''", "S") & " and minor_cd =  " & FilterVar(frm1.txtsub_type.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        IF IntRetCd = false Then
            Call  DisplayMsgBox("800142","X","X","X")
            frm1.txtsub_type_nm.value = ""
            frm1.txtsub_type.focus   
            Set gActiveElement = document.ActiveElement
            txtsub_type_Onchange = True
        ELSE
            frm1.txtsub_type_nm.value = Trim(Replace(lgF0,Chr(11),""))
        END IF
    END IF  
    
End Function 
'========================================================================================================
'   Event Name : txtsub_cd_Onchange            
'   Event Desc :
'========================================================================================================
Function txtsub_cd_Onchange()
    Dim iDx
    Dim IntRetCd

    IF Trim(frm1.txtsub_cd.value) = "" THEN
        frm1.txtsub_cd.value = ""
        frm1.txtsub_cd_nm.value = ""
        frm1.txtsub_cd.focus       
    ELSE
        IntRetCd =  CommonQueryRs(" allow_nm "," HDA010T "," CODE_TYPE=" & FilterVar("2", "''", "S") & " AND allow_cd =  " & FilterVar(frm1.txtsub_cd.value , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

        IF IntRetCd = false Then
            Call  DisplayMsgBox("800142","X","X","X")
            frm1.txtsub_cd_nm.value = ""
            frm1.txtsub_cd.focus   
            Set gActiveElement = document.ActiveElement
            txtsub_cd_Onchange = True
        ELSE
            frm1.txtsub_cd_nm.value = Trim(Replace(lgF0,Chr(11),""))
        END IF
    END IF 
End Function 

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("0000101111")

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

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)


    Call SetPopupMenuItemInf("0000000000")

    gMouseClickStatus = "SP1C" 

    Set gActiveSpdSheet = frm1.vspdData2
   
End Sub
'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
    End If
End Sub    

Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP1C" Then
       gMouseClickStatus = "SP1CR"
    End If
End Sub  

'=======================================================================================================
'   Event Name : txtsub_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtsub_yymm_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtsub_yymm_dt.Action = 7
        frm1.txtsub_yymm_dt.focus
    End If
End Sub
'==========================================================================================
'   Event Name : txtsub_yymm_dt_KeyDown()
'   Event Desc : 조회조건부의 txtpay_yymm_dt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtsub_yymm_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData.Col = pvCol1
    frm1.vspdData2.ColWidth(pvCol1) = frm1.vspdData.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData2.Col = pvCol1
    frm1.vspdData.ColWidth(pvCol1) = frm1.vspdData2.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
	
    call vspdData2s_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================

Sub vspdData2s_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col , ByVal Row, ByVal newCol , ByVal newRow ,Cancel )
    frm1.vspdData2.Col = newCol
    frm1.vspdData2.Action = 0
End Sub


'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
        frm1.vspdData2.LeftCol=NewLeft   	
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
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
        frm1.vspdData.LeftCol=NewLeft   	
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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows = 0 Then
        Exit Sub
    End If
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
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>월공제내역조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* >&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD width=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>    
			        <TD <%=HEIGHT_TYPE_02%>></TD>
			    </TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>width=100%>
						    <TR>
							    <TD CLASS="TD5" NOWRAP>공제년월</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/h5103ma1_txtsub_yymm_dt_txtsub_yymm_dt.js'></script></TD>		
								<TD CLASS="TD5" NOWRAP>사원</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtEmp_No" MAXLENGTH="13" SIZE="13" ALT ="사번" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: openEmptName(0)">
								                    <INPUT NAME="txtName" MAXLENGTH="30" SIZE="20" ALT ="성명" tag="14XXXU"></TD>
								
							    
							</TR>   
							<TR>	
								<TD CLASS="TD5" NOWRAP>공제구분</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtsub_type"  MAXLENGTH="1" SIZE="10" ALT ="공제구분" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('3')">
								                     <INPUT NAME="txtsub_type_nm"  MAXLENGTH="20" SIZE="20" ALT ="공제구분" tag="14XXXU"></TD>
								<TD CLASS="TD5" NOWRAP>공제코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtsub_cd"  MAXLENGTH="3" SIZE="10" ALT ="공제코드" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('4')">
								                     <INPUT NAME="txtsub_cd_nm"  MAXLENGTH="20" SIZE="20" ALT ="공제코드" tag="14XXXU"></TD>
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
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h5103ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=64 VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%" width=100%>
									<script language =javascript src='./js/h5103ma1_vaSpread2_vspdData2.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
	</TR>
	
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<INPUT TYPE=HIDDEN NAME="txtCheck"       tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


