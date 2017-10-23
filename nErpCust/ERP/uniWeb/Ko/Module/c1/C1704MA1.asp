
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 계정별 배부규칙등록 
'*  3. Program ID           : c1704ma1.asp
'*  4. Program Name         : 계정별 배부규칙등록 
'*  5. Program Desc         : 계정별 배부규칙등록 
'*  6. Modified date(First) : 2000/08/23
'*  7. Modified date(Last)  : 2000/08/23
'*  8. Modifier (First)     : 강창구 
'*  9. Modifier (Last)      : Cho Ig Sung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->


<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE=vbscript>
Option Explicit	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c1704mb1.asp"
Const BIZ_PGM_QRY_ID2 = "c1704mb9.asp"
Const BIZ_PGM_QRY_ID3 = "cb009mb1.asp"

Dim C_ACCT_Cd  
Dim C_ACCT_PB  
Dim C_ACCT_Nm  
Dim C_DstbFctr_Cd  
Dim C_DstbFctr_PB  
Dim C_DstbFctr_Nm  

Dim C_CheckBox  
Dim C_RecvCost_Cd  
Dim C_RecvCost_Nm  
Dim C_Flag  


Dim lgBlnFlgChgValue
Dim lgIntGrpCount 
Dim lgIntFlgMode 

Dim lgLngCurRows
Dim lgCurrRow
Dim lgSortKey

Dim intItemCnt					

Dim IsOpenPop
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6



'========================================================================================================
Sub initSpreadPosVariables()  
	C_ACCT_Cd		= 1
	C_ACCT_PB		= 2
	C_ACCT_Nm		= 3
	C_DstbFctr_Cd	= 4
	C_DstbFctr_PB	= 5
	C_DstbFctr_Nm	= 6
End Sub


'========================================================================================================
Sub initSpreadPosVariables1() 
	C_CheckBox		= 1
	C_RecvCost_Cd	= 2
	C_RecvCost_Nm	= 3
	C_Flag			= 4
End Sub



'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE 
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0

    lgLngCurRows = 0

End Sub

Sub SetDefaultVal()
    Call ggoOper.ClearField(Document, "1") 
End Sub


'========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "COOKIE", "MA") %>
End Sub


'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

Select Case UCase(pvSpdNo)
		Case "A" 

	Call initSpreadPosVariables()  
	        
	With frm1.vspdData
	
		.MaxCols = C_DstbFctr_NM + 1
		.Col = .MaxCols	
		.ColHidden = True

   		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread  

		ggoSpread.ClearSpreadData
		
		.ReDraw = False

		Call GetSpreadColumnPos("A")

	    ggoSpread.SSSetEdit C_ACCT_Cd, "계정코드", 14,,,20,2
		ggoSpread.SSSetButton C_ACCT_PB
 		ggoSpread.SSSetEdit C_ACCT_Nm, "계정명", 25
		ggoSpread.SSSetEdit C_DstbFctr_CD, "배부요소코드", 10,,,2,2
		ggoSpread.SSSetButton C_DstbFctr_PB
	    ggoSpread.SSSetEdit C_DstbFctr_NM, "배부요소명", 20

	call ggoSpread.MakePairsColumn(C_ACCT_Cd,C_ACCT_PB)
	call ggoSpread.MakePairsColumn(C_DstbFctr_CD,C_DstbFctr_PB)
  
		
'	    ggoSpread.SSSetSplit(C_ACCT_Nm)		
		.ReDraw = True
                
         end with
         
   Case "B" 
       
		Call initSpreadPosVariables1()             

     With frm1
                    
		    .vspdData2.MaxCols = C_RecvCost_Nm	+1
            .vspdData2.Col = .vspdData2.MaxCols						
            .vspdData2.ColHidden = True 
                 
            ggoSpread.Source = .vspdData2
            ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread   

			ggoSpread.ClearSpreadData
			
			.vspdData2.ReDraw = false   
        
			Call GetSpreadColumnPos("B")

   	        ggoSpread.SSSetCheck     C_CheckBox ,"대상여부",10 , ,"",true          
            ggoSpread.SSSetEdit		C_RecvCost_Cd, "배부대상", 16,,,10
	        ggoSpread.SSSetEdit		C_RecvCost_Nm, "코스트센타명", 27
            .vspdData2.ReDraw = True
     
	End With
	
 End Select
 
        SetSpreadLock "I", 0, -1, -1
        
End Sub


'======================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
    Dim objSpread
    
    With frm1
    

    Select Case Index
		Case 0
			ggoSpread.Source = .vspdData
			Set objSpread = .vspdData
			lRow2 = objSpread.MaxRows
			objSpread.Redraw = False

			ggoSpread.SpreadLock C_ACCT_Cd, lRow, C_ACCT_Cd, lRow2
			ggoSpread.SpreadLock C_ACCT_PB, lRow, C_ACCT_PB, lRow2	
            ggoSpread.SpreadLock C_ACCT_Nm, lRow, C_ACCT_Cd, lRow2   
	End Select
    
	ggoSpread.SpreadLock C_DstbFctr_Nm, lRow, C_DstbFctr_Nm, lRow2
	ggoSpread.SSSetRequired C_DstbFctr_Cd, -1, C_DstbFctr_Cd	' 배부요소 
    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1      
     
    objSpread.Redraw = True
    Set objSpread = Nothing
    
    End With
    
End Sub



'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1

		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired	C_ACCT_Cd		,pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_ACCT_Nm		,pvStartRow	,pvEndRow
		ggoSpread.SSSetRequired	C_DstbFctr_Cd	,pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_DstbFctr_Nm	,pvStartRow	,pvEndRow

		.vspdData.ReDraw = True
		
    End With

End Sub


'======================================================================================================
Sub SetSpread2Color()
Dim strStartRow, strEndRow

    With frm1
    
		strStartRow = 1
		strEndRow	= .vspdData2.MaxRows

		ggoSpread.Source	= .vspdData2
		.vspdData2.ReDraw	= False
		
		ggoSpread.SSSetProtected   C_RecvCost_Cd	,strStartRow, strEndRow
        ggoSpread.SSSetProtected   C_RecvCost_Nm	,strStartRow, strEndRow
 		
		.vspdData2.ReDraw = True

    End With

End Sub


Function CookiePage(ByVal Kubun)
	
	On Error Resume Next

	Const CookieSplit = 4877						

	Dim strTemp, arrVal

	If Kubun = 1 Then									

		if frm1.vspddata.maxrows = 0 then  exit function
		    

		frm1.vspddata.Row = frm1.vspddata.ActiveRow
		frm1.vspddata.Col = C_Cost_Cd

		WriteCookie CookieSplit , frm1.txtVerCd.value  & Parent.gRowSep & frm1.vspddata.value

	ElseIf Kubun = 0 Then								

		strTemp = ReadCookie(CookieSplit)
			
		If strTemp = "" then Exit Function
			
		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		frm1.txtVerCd.value =  arrVal(0)
		frm1.txtcostcd.value = arrVal(1)

		if Err.number <> 0 then
			Err.Clear
			WriteCookie CookieSplit , ""
			exit function
		end if
		
		Call FncQuery()		
			
		WriteCookie CookieSplit , ""

	End If
	
End Function

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
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ACCT_Cd				= iCurColumnPos(1)
			C_ACCT_PB				= iCurColumnPos(2)
			C_ACCT_Nm				= iCurColumnPos(3)    
			C_DstbFctr_Cd		    = iCurColumnPos(4)
			C_DstbFctr_PB			= iCurColumnPos(5)
			C_DstbFctr_Nm			= iCurColumnPos(6)
	   Case "B"	
			ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_CheckBox				= iCurColumnPos(1)
			C_RecvCost_Cd			= iCurColumnPos(2)
			C_RecvCost_Nm			= iCurColumnPos(3)
			C_Flag					= iCurColumnPos(4)
			
    End Select    
End Sub



Sub InitComboBoxGrid()

	ggoSpread.source = frm1.vspdData2
	ggoSpread.SetCombo "Y" & vbtab & "N" , C_CheckBox

End Sub



Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case iWhere
		Case 0
			arrParam(0) = "버전팝업"	
			arrParam(1) = "C_Dstb_Rule_by_CC"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "버전"

			arrField(0) = "ver_cd"
    
			arrHeader(0) = "버전"	

		Case 1
			arrParam(0) = "코스트센타팝업"
			arrParam(1) = "C_DSTB_RULE_BY_CC , B_COST_CENTER"
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = "Cost_Cd = Give_Cost_Cd and ver_cd = " & FilterVar(frm1.txtvercd.value, "''", "S") & ""	
			arrParam(5) = "코스트센타"

			arrField(0) = "COST_CD"	
			arrField(1) = "COST_NM"	
    
			arrHeader(0) = "코스트센타코드"
			arrHeader(1) = "코스트센타명"
			
		Case 2
			arrParam(0) = "배부요소팝업"	
			arrParam(1) = "C_DSTB_FCTR"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""	
			arrParam(5) = "배부요소"	

			arrField(0) = "DSTB_FCTR_CD"
			arrField(1) = "DSTB_FCTR_NM"
    
			arrHeader(0) = "배부요소코드"		
			arrHeader(1) = "배부요소명"	
		Case 3
			arrParam(0) = "계정팝업"
			arrParam(1) = "A_ACCT" 
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = "temp_fg_3 in (" & FilterVar("M2", "''", "S") & " ," & FilterVar("M3", "''", "S") & " )"	
			arrParam(5) = "계정"	

			arrField(0) = "ACCT_CD"	
			arrField(1) = "ACCT_NM"	
    
			arrHeader(0) = "계정코드"
			arrHeader(1) = "계정명"

	End Select
    
    	If iWhere = 0 Then
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=360px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	  If iWhere = 0 Then
	    frm1.txtVerCd.focus
	  Else 
	    frm1.txtCostCd.focus
	  End If	    
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				 frm1.txtVerCd.focus
				.txtVerCd.value = arrRet(0)
			Case 1
	             frm1.txtCostCd.focus					
				.txtCostCd.value = arrRet(0)
				.txtCostNm.value = arrRet(1)
			Case 2
				.vspdData.Row = .vspdData.ActiveRow	
				.vspdData.Col = C_DstbFctr_Cd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_DstbFctr_Nm
				.vspdData.Text = arrRet(1)

				call vspdData_Change(C_DstbFctr_Cd, frm1.vspddata.activerow )                                
			Case 3
				.vspdData.Row = .vspdData.ActiveRow	
				.vspdData.Col = C_ACCT_Cd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_ACCT_Nm
				.vspdData.Text = arrRet(1)

				call vspdData_Change(C_ACCT_Cd, frm1.vspddata.activerow )

		End Select

	End With

End Function

Function DbQuery2(ByVal Row)
Dim strVal
Dim boolExist
Dim lngRows
	
	boolExist = False
	
	With frm1
	    .vspdData.row = Row
	    .vspdData.col = C_ACCT_Cd
	    .hItemCd.Value = .vspdData.Text

	    If Trim(.hItemCd.Value) = "" Then
	        Exit Function
	    End If
	    
		If CopyFromData(.hItemCd.Value) = True Then
		    Exit Function
		End If
		    	
		
		Call LayerShowHide(1)

		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ACCT_Cd
	
	    If lgIntFlgMode = Parent.OPMD_UMODE Then 
		strVal = BIZ_PGM_QRY_ID2 & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtVerCd=" & Trim(.htxtVerCd.value)	
     		strVal = strVal & "&txtCostCd=" & .htxtCostCd.Value    		
     		strVal = strVal & "&txtAcctCd=" & .hItemCd.Value    		
    		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
	    Else
		strVal = BIZ_PGM_QRY_ID2 & "?txtMode=" & Parent.UID_M0001	
		strVal = strVal & "&txtVerCd=" & Trim(.txtVerCd.value)	
     		strVal = strVal & "&txtCostCd=" & .txtCostCd.Value    		
    		strVal = strVal & "&txtAcctCd=" & .vspdData.text    		
    		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
	    End If
	End With

	Call RunMyBizASP(MyBizASP, strVal)
	
	DbQuery2 = True
	
End Function

Function DbQueryOk2()
	With frm1
		.vspdData.Col = C_ACCT_Cd:    intItemCnt = .vspddata.MaxRows
        	ggoSpread.Source = .vspdData2
				
		SetSpread2Color 
    End With
End Function

Function DbQuery3(ByVal Row)
Dim strVal
Dim lngRows
Dim i 

	DbQuery3 = False
	

	With frm1
	    .vspdData.row = Row
	    .vspdData.col = C_ACCT_Cd
	    
		If CopyFromData(Trim(.vspdData.Text)) = True Then
		    Exit Function
		End If
 
        IF LayerShowHide(1) = False Then
			Exit Function
		END IF
	
		.vspdData.Row = Row
		.vspdData.Col = C_ACCT_Cd
	    
		strVal = BIZ_PGM_QRY_ID3 & "?txtMode=" & Parent.UID_M0001	
 		strVal = strVal & "&txtItemCd=" & .vspdData.Text
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows

	End With

	Call RunMyBizASP(MyBizASP, strVal)	

	DbQuery3 = True
	
End Function

Function DbQueryOk3()
Dim i
	With frm1
		
		ggoSpread.Source = .vspdData2
				
		SetSpread2Color 
		
    End With
End Function

Function CheckSpread3()
	Dim i
	Dim tmpDrCrFG

	CheckSpread3 = False

	With frm1
	 	for i = 1 to .vspdData3.MaxRows
		    .vspdData3.Row = i
		    .vspdData3.Col = C_DrFg + 1
		    if (.vspddata2.text = tmpDrCrFG and .vspddata2.text <> "") _
                            or .vspddata2.text = "Y" or .vspddata2.text = "DC" then

  			  .vspdData3.Col = C_CtrlVal + 1
		
			  if Trim(.vspdData3.text) = "" then
				Exit Function
		  	  end if
		    end if
		Next	
		
        End With
	CheckSpread3 = True
End Function

Function FindNumber(ByVal objSpread, ByVal intCol)
Dim lngRows
Dim lngPrevNum
Dim lngNextNum

    FindNumber = 0

    lngPrevNum = 0
    lngNextNum = 0
    
    With frm1
        
        If objSpread.MaxRows = 0 Then
            Exit Function
        End If
        
        For lngRows = 1 To objSpread.MaxRows
            objSpread.Row = lngRows
            objSpread.Col = intCol
            lngNextNum = UniClng(objSpread.Text,0)
            
            If lngNextNum > lngPrevNum Then
                lngPrevNum = lngNextNum
            End If
            
        Next
       
    End With        
    
    FindNumber = lngPrevNum
    
End Function

Function FindData()
Dim strApNo
Dim strItemSeq
Dim strDtlSeq
Dim lRows

    FindData = 0

    With frm1
        
        For lRows = 1 To .vspdData3.MaxRows
        
            .vspdData3.Row = lRows
            .vspdData3.Col = 1
            strItemSeq = .vspdData3.Text
            .vspdData3.Col = 3
            strDtlSeq = .vspdData3.Text
            
            .vspdData.Row = frm1.vspdData.ActiveRow
            .vspdData2.Row = frm1.vspdData2.ActiveRow
            
            .vspdData.Col = C_ACCT_Cd
            If strItemSeq = .vspdData.Text Then
                
                .vspdData2.Col = C_RecvCost_Cd
                If strDtlSeq = .vspdData2.Text Then
                    
                    FindData = lRows
                    Exit Function
                    
                End If
                
            End If    
        Next
        
    End With        
    
End Function

Function CopyFromData(ByVal strItemSeq)
Dim lngRows , i
Dim boolExist
Dim iCols

    boolExist = False

    frm1.vspdData2.maxrows = 0
    CopyFromData = boolExist
    
     
    With frm1
	
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = 1                

            If Trim(strItemSeq) = Trim(.vspdData3.Text) Then
                boolExist = True
                Exit For
            End If    
        Next
        
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            .vspdData2.Redraw = False
	                

            While lngRows <= .vspdData3.MaxRows

                .vspdData3.Row = lngRows
                
                .vspdData3.Col = 1
                
                If Trim(strItemSeq) <> Trim(.vspdData3.Text) Then
                    lngRows = .vspdData3.MaxRows + 1
                Else

                    .vspdData2.MaxRows = .vspdData2.MaxRows + 1
                    .vspdData2.Row = .vspdData2.MaxRows
                    .vspdData2.Col = 0
                    .vspdData3.Col = 0
                    .vspdData2.Text = .vspdData3.Text

		            .vspdData2.Col = C_CheckBox
                    .vspdData3.Col = 2
                    .vspdData2.Text = .vspdData3.Text
                  
                    .vspdData2.Col = C_RecvCost_Cd
                    .vspdData3.Col = 3
                    .vspdData2.Text = .vspdData3.Text

                    .vspdData2.Col = C_RecvCost_Nm
                    .vspdData3.Col = 4
                    .vspdData2.Text = .vspdData3.Text

                  
                    'For iCols = 1 To .vspdData3.MaxCols
                    '    .vspdData2.Col = iCols
                    '    .vspdData3.Col = iCols + 1
                    '    .vspdData2.Text = .vspdData3.Text

                    'Next
                        
                End If   
                
                lngRows = lngRows + 1
                
            Wend
            
            ggoSpread.Source = frm1.vspdData2

            SetSpread2Color	

            frm1.vspdData.Row = lgCurrRow
            frm1.vspdData.Col = frm1.vspdData.MaxCols
            ggoSpread.Source = frm1.vspdData
            
            frm1.vspdData2.Redraw = True
            
        End If
            
    End With        
    
    CopyFromData = boolExist
    
End Function

Sub CopyToHSheet(ByVal Row)
Dim lRow
Dim iCols

	With frm1 
        
	    lRow = FindData
	    If lRow > 0 Then
            .vspdData3.Row = lRow
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
        
 
			.vspdData2.Col = C_CheckBox
			.vspdData3.Col = 2
            .vspdData3.Text = .vspdData2.value

       

			.vspdData2.Col = C_RecvCost_Cd
			.vspdData3.Col = 3
            .vspdData3.Text = .vspdData2.value

			
			.vspdData2.Col = C_RecvCost_Nm
			.vspdData3.Col = 4
            .vspdData3.Text = .vspdData2.value
            
            'For iCols = 1 To .vspdData2.MaxCols
            '    .vspdData2.Col = iCols
            '    .vspdData3.Col = iCols + 1
            '    .vspdData3.Text = .vspdData2.Text
		
            'Next
            
        End If

	End With
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = 0
	if frm1.vspdData.Text <> ggoSpread.InsertFlag and frm1.vspdData.Text <> ggoSpread.DeleteFlag then
   	    frm1.vspdData.Text = ggoSpread.UpdateFlag
	End if
	
End Sub

Function DeleteHSheet(ByVal strItemSeq)
Dim boolExist
Dim lngRows
 
    DeleteHSheet = False
    boolExist = False

    frm1.vspdData2.MaxRows = 0

    With frm1

        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = 1                

            If strItemSeq = .vspdData3.Text Then
                boolExist = True
                Exit For
            End If    
        Next

         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            
            While lngRows <= .vspdData3.MaxRows

                .vspdData3.Row = lngRows
                .vspdData3.Col = 1
                
                If strItemSeq <> .vspdData3.Text Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
                    .vspdData3.Action = 5
                    .vspdData3.MaxRows = .vspdData3.MaxRows - 1
                End If   

            Wend
            
            ggoSpread.Source = frm1.vspdData2
            
            frm1.vspdData.Row = lgCurrRow
            frm1.vspdData.Col = frm1.vspdData.MaxCols
            ggoSpread.Source = frm1.vspdData
            
            frm1.vspdData2.Redraw = True
            
        End If
            
    End With
        
    DeleteHSheet = True
End Function

Function CancelHSheet(ByVal strItemSeq)
Dim lngRows
 
    CancelHSheet = False
    lngRows = 1
    With frm1
		For lngRows = 1 To .vspdData3.MaxRows
		    .vspdData3.Row = lngRows
		    .vspdData3.Col = 1                

		    If strItemSeq = .vspdData3.Text Then
		        Exit For
            End If    
        Next
        
        
        While lngRows <= .vspdData3.MaxRows
			.vspdData3.Row = lngRows
            .vspdData3.Col = 1
            
            If Trim(strItemSeq) = Trim(.vspdData3.Text) Then
                .vspdData3.Col = 0
                
                IF .vspdData3.Text = ggoSpread.InsertFlag or  .vspdData3.Text = ggoSpread.DeleteFlag Then
					
					.vspdData3.Text = ""						
					.vspdData3.Col = 2
					
					IF .vspdData3.text = "0" Then
						.vspdData3.text = "1"
					ELSE
						.vspdData3.text = "0"
					END IF	
				END IF
				lngRows = lngRows + 1   
			ELSE
				lngRows = .vspdData3.MaxRows + 1
			End If
			
			
        Wend
    End With
    CancelHSheet = True
End Function    

Function SortHSheet()
    
    With frm1
        .vspdData3.BlockMode = True
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.SortBy = 0
        
        .vspdData3.SortKey(1) = 1
        .vspdData3.SortKey(2) = 2
        
        .vspdData3.SortKeyOrder(1) = 1
        .vspdData3.SortKeyOrder(2) = 1
        
        .vspdData3.Col = 1
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 0
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.Action = 25
        .vspdData3.BlockMode = False
    End With        
    
End Function

Sub ShowHidden()
Dim strHidden
Dim lngRows
Dim lngCols
    
    With frm1.vspdData3
        For lngRows = 1 To .MaxRows
            .Row = lngRows
            
            .Col = 1  
            strHidden = strHidden & Parent.gRowSep & .Text
            .Col = 2
            strHidden = strHidden & Parent.gRowSep & .Text
            .Col = 3
            strHidden = strHidden & Parent.gRowSep & .Text
            .Col = 4
            strHidden = strHidden & Parent.gRowSep & .Text
		
            .Col = 5  
            strHidden = strHidden & Parent.gRowSep & .Text		
   		
            strHidden = strHidden & Parent.gRowSep
        Next
    End With        

    
End Sub

Sub SetSpreadFG( pobjSpread , ByVal pMaxRows )
    Dim lngRows 
    
    For lngRows = 1 To pMaxRows
        pobjSpread.Col = 0
        pobjSpread.Row = lngRows
        pobjSpread.Text = ""
    Next
    
End Sub

Sub Form_Load()

	Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet("A")
    Call InitSpreadSheet("B")
    Call InitVariables

    Call InitComboBoxGrid
    Call SetDefaultVal
    Call SetToolbar("110011010010111")

    frm1.txtvercd.focus
    Call CookiePage(0)	
    frm1.txtCommandMode.value = "CREATE"

    frm1.vspdData3.MaxRows = 0
    frm1.vspdData3.MaxCols = 5 
   	Set gActiveElement = document.activeElement			       

End Sub

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

Sub vspdData_onfocus()
    
	If lgIntFlgMode <> Parent.OPMD_UMODE Then    
                                             
'	Call SetToolbar("1100111100111111")		

    Else  

        'Call SetToolbar("11000100001111") 
		Call SetToolbar("1100111100111111")       


    End If  

     
End Sub

Sub vspdData2_onfocus()

 
    If lgIntFlgMode <> Parent.OPMD_UMODE Then 
'        Call SetToolbar("1100100000000011")

    Else      
'        Call SetToolbar("1100100000000011") 

    End If
    
    
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Dim i
    
   	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	

    gMouseClickStatus = "SPC"	'Split 상태코드 

    Set gActiveSpdSheet = frm1.vspdData
    
	if frm1.vspdData.maxrows = 0 then exit sub
	
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If

	IF Row = 0 Then
		Exit Sub
	ENd IF

    ggoSpread.Source = frm1.vspdData
	frm1.vspddata.row = Row
	
	frm1.vspdData2.maxrows = 0

  	frm1.vspdData.Col = C_ACCT_Cd
	
        If Len(Trim(frm1.vspdData.Text)) > 0 Then
           frm1.vspddata.Col = 0
    

	   If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
               Exit Sub
           End if
 
      	   If frm1.vspddata.Text = ggoSpread.InsertFlag Then           
				IF DbQuery3(Row) = False Then
					Exit Sub
				ENd IF
		   Else           
				IF DbQuery2(Row) = False Then
					Exit sub
				END IF	
	       End if	
   
	end if	
        
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0001111111")
	Else 
		Call SetPopupMenuItemInf("0001111111")
	End If	

    gMouseClickStatus = "SP2C"	'Split 상태코드 

    Set gActiveSpdSheet = frm1.vspdData2


        
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
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    	
End Sub


'========================================================================================================
'   Event Name : vspdData2_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
	
End Sub


'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================

Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_ACCT_Nm Or NewCol <= C_ACCT_Nm Then
'        Cancel = True
'        Exit Sub
'    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    'If Col <= C_MinorNm Or NewCol <= C_MinorNm Then
    '    Cancel = True
    '    Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub


Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	ggoSpread.Source = frm1.vspdData


	With frm1
		If Row > 0 then
          select case Col
           Case  C_ACCT_PB 

			.vspdData.Col = C_ACCT_Cd
			.vspdData.Row = Row
									
		   Call OpenPopUp(.vspdData.Text, 3 )
	 	   Case C_DstbFctr_PB
			.vspdData.Col = C_DstbFctr_cd
			.vspdData.Row = Row
									
			Call OpenPopUp(.vspdData.Text, 2 )

          End select	
		End If
		Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
	End With
	
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
End Sub

Sub vspdData_Change(ByVal Col, ByVal Row )
Dim tmpAcctCd
Dim intRetCd

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    
    select case Col
    Case   C_ACCT_Cd
	    

	    frm1.vspdData.Col = 0
		If  frm1.vspdData.Text = ggoSpread.InsertFlag Then
	
		        frm1.vspdData.Col = C_ACCT_Cd
			frm1.hItemCd.value = frm1.vspdData.Text
		        			
		        If Len(frm1.vspdData.Text) > 0 Then

			    frm1.vspdData.Row = Row
			    frm1.vspdData.Col = C_ACCT_Cd	
			    DeleteHsheet frm1.vspdData.Text
		    
		            Call DbQuery3 (Row)
		        End If  
    		End If
   end select

End Sub

Sub vspdData2_Change(ByVal Col, ByVal Row)

	
   	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = 0
	    
	select case Col
	Case   C_CheckBox
			If  frm1.vspdData2.Text <> ggoSpread.InsertFlag and frm1.vspdData2.Text <> ggoSpread.DeleteFlag Then
		
			        frm1.vspdData2.Col = C_CheckBox
			        			
			        If frm1.vspdData2.value = "0" Then
	
				    frm1.vspdData2.Row = Row
				    frm1.vspdData2.Col = 0	
				    frm1.vspdData2.text = ggoSpread.DeleteFlag
			    	else
				    frm1.vspdData2.Row = Row
				    frm1.vspdData2.Col = 0	
				    frm1.vspdData2.text = ggoSpread.InsertFlag 
			        End If  
			elseif frm1.vspdData2.Text = ggoSpread.DeleteFlag Then
				frm1.vspdData2.Col = C_CheckBox
			        			
			    If frm1.vspdData2.value = "1" Then
	
				    frm1.vspdData2.Row = Row
				    frm1.vspdData2.Col = 0	
			        frm1.vspdData2.text = frm1.vspdData2.Row
			    End If
			elseif frm1.vspdData2.Text = ggoSpread.InsertFlag Then
				frm1.vspdData2.Col = C_CheckBox
			        			
			    If frm1.vspdData2.value = "0" Then
				    frm1.vspdData2.Row = Row
				    frm1.vspdData2.Col = 0	
				    frm1.vspdData2.text = frm1.vspdData2.Row
			    End If  

	    	End If  
	
	end select


	CopyToHSheet Row
	
End Sub	

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

		If Row = NewRow Then
		    Exit Sub
		End If

		'call vspdData_Click( NewCol, NewRow)  

    End With

    
End Sub

Sub txtDocCur_Change()
    
    lgBlnFlgChgValue = True

End Sub

Sub txtDeptCd_Change()
    
    lgBlnFlgChgValue = True

End Sub

Sub txttempGLDt_Change()
    
    lgBlnFlgChgValue = True

End Sub

Sub cboGLType_Change()
    
    lgBlnFlgChgValue = True

End Sub


Function FncQuery() 
    Dim IntRetCD 
    Dim RetFlag
    
    FncQuery = False
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")	
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If

    Call ggoOper.ClearField(Document, "2")
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData

   
    Call InitVariables	

    Call InitComboBoxGrid

   if frm1.txtCostCd.value = "" then
		frm1.txtCostNM.value = ""
    end if
    
    If Not chkField(Document, "1") Then	
       Exit Function
    End If

    IF DbQuery = False then
		Exit function
	END IF	
    
    FncQuery = True	
    
End Function

Function FncNew() 
	Dim IntRetCD 
    
    FncNew = False 
    Err.Clear
    On Error Resume Next
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    Call ggoOper.ClearField(Document, "1") 
    Call ggoOper.ClearField(Document, "2")
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData


    Call ggoOper.LockField(Document, "N") 
    Call InitVariables 

    Call SetDefaultVal
    Call InitSpreadSheet
    
    Call InitComboBoxGrid


    frm1.txtCommandMode.value = "CREATE"
    Call SetToolbar("110001000000001")	
    frm1.vspdData.MaxRows = 0
    frm1.vspdData2.MaxRows = 0

    frm1.vspdData3.MaxRows = 0 

    FncNew = True
    
End Function

Function FncDelete() 
	Dim IntRetCD 
    
    FncDelete = False
    Err.Clear
    On Error Resume Next 

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900002", Parent.VB_YES_NO,"x","x")
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If

    IF DbDelete = False Then
		Exit function
	END IF	     
	
    FncDelete = True 
   
End Function

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear     

    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False and ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If

    If Not chkField(Document, "1") Then 
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then 
       Exit Function
    End If

    IF DbSave = False Then
		Exit Function
	END IF
    
    FncSave = True                                                          
    
End Function


'========================================================================================================
Function FncCopy() 

    Dim  IntRetCD
	 
    frm1.vspdData.ReDraw = False
	
    if frm1.vspdData.maxrows < 1 then exit function 

	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    frm1.vspdData.Col = C_Acct_CD
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_Acct_Nm
    frm1.vspdData.Text = ""

    frm1.vspdData.ReDraw = True
    
    frm1.vspdData2.MaxRows =0
End Function


Function FncCancel() 
    Dim iCostCd

    if frm1.vspdData.MaxRows < 1 then Exit Function

    With frm1.vspdData
        .Row = .ActiveRow
        .Col = 0

	if .row = 0 then 
	   Exit Function
	end if

        If .Text = ggoSpread.InsertFlag Then
            .Col = C_ACCT_Cd
            DeleteHSheet(.Text)
        ElseIF .Text = ggoSpread.UpdateFlag Then
		    .Col = C_ACCT_Cd
		     CanCelHSheet(.Text)
		End if
		
        .Col = C_ACCT_Cd
	iCostCd = .Text

        ggoSpread.Source = frm1.vspdData	
        ggoSpread.EditUndo
	if .activerow = 0 then 
	   Exit Function
	end if

	If .Text = ggoSpread.InsertFlag Then
   		    frm1.htxtVerCd.value = Trim(frm1.txtVerCd.value)
			frm1.htxtCostCd.value = Trim(frm1.txtCostCd.value)
            .Col = C_ACCT_Cd
            IF .text <> "" Then
				frm1.hItemCd.value = .Text
				frm1.vspdData2.MaxRows = 0
				Call DbQuery3(.ActiveRow)
			END If
        Else
		    frm1.htxtVerCd.value = Trim(frm1.txtVerCd.value)
			frm1.htxtCostCd.value = Trim(frm1.txtCostCd.value)
            .Col = C_ACCT_Cd
            frm1.hItemCd.value = .Text
            frm1.vspdData2.MaxRows = 0
            Call DbQuery2(.ActiveRow)
        End if

    End With
     
End Function


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
        ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
End Function

Function FncDeleteRow() 
	Dim lDelRows
	Dim iDelRowCnt, i
        Dim DelCostCd
    
	if frm1.vspdData.maxrows < 1 then exit function 

    With frm1.vspdData 
        .Row = .ActiveRow
	.Col = C_ACCT_Cd 
        DelCostCd = .Text
    
    	ggoSpread.Source = frm1.vspdData 

    	lDelRows = ggoSpread.DeleteRow
    
    End With

    DeleteHsheet DelCostCd

End Function

Function FncPrint() 
    Call parent.FncPrint()                                              
End Function

Function FncPrev() 
    On Error Resume Next
End Function

Function FncNext() 
    On Error Resume Next
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
End Function


'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub



'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub


'========================================================================================
Sub PopRestoreSpreadColumnInf()

	Dim indx
	Dim lngActRow1, lngActRow2
	Dim lngActCol1, lngActCol2

	lngActRow1 = frm1.vspdData.ActiveRow
	lngActCol1 = frm1.vspdData.ActiveCol
	lngActRow2 = frm1.vspdData2.ActiveRow
	lngActCol2 = frm1.vspdData2.ActiveCol

	If gActiveSpdSheet.Name <> "" Then
		For indx = 0 To frm1.vspdData.MaxRows
			frm1.vspdData.Row = indx
			frm1.vspdData.Col = 0
'			If frm1.vspdData.Text = ggoSpread.DeleteFlag Or _
'			   frm1.vspdData.Text = ggoSpread.UpdateFlag Then
'				Call FncUndoData(indx)
'			End If
			
			Select Case Trim(UCase(gActiveSpdSheet.Name))
				Case "VSPDDATA"
					frm1.vspdData.Row = lngActRow1 
					frm1.vspdData.Col = lngActCol1
					frm1.vspdData.Action = 0
					
			   		
				Case "VSPDDATA2"
					frm1.vspdData2.Row = lngActRow2
					frm1.vspdData2.Col = lngActCol2
					frm1.vspdData2.Action = 0
			End Select
		Next
	End If

	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
 
    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData()
			SetSpreadLock "I", 0, -1, -1
			
		Case "VSPDDATA2"
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")			' 그리드2 초기화 
			Call InitComboBoxGrid()
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread2Color()
	End Select

	If frm1.vspdData2.MaxRows <= 0 Then	
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If

End Sub


Function FncExit()
	Dim IntRetCD
	FncExit = False
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If

    FncExit = True
End Function

Function DbQuery() 
	Dim strVal
	Dim RetFlag
    DbQuery = False
    
    IF LayerShowHide(1) = False Then
		Exit function
	END IF
	
    frm1.vspdData.MaxRows = 0
    frm1.vspdData2.MaxRows = 0
    frm1.vspdData3.MaxRows = 0 

    Err.Clear
    
    With frm1

    If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtVerCd=" & Trim(.htxtVerCd.value)
			strVal = strVal & "&txtCostCd=" & Trim(.htxtCostCd.value)
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001 
			strVal = strVal & "&txtVerCd=" & Trim(.txtVerCd.value)
			strVal = strVal & "&txtCostCd=" & Trim(.txtCostCd.value)	
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)
		    
    End With
    
    DbQuery = True

End Function

Function DbQueryOk()
	
	With frm1

		SetSpreadLock "Q", 0, 1, ""
    
        lgIntFlgMode = Parent.OPMD_UMODE	
        
        Call ggoOper.LockField(Document, "I")		


        Call SetToolbar("110011110011111")		
        
        If .vspdData.MaxRows > 0 Then
            .vspdData.Row = 1
            .vspdData.Col = C_ACCT_Cd

			frm1.vspddata2.maxrows = 0
			
            Call DbQuery2(1)
	    
        End If
    
    End With


End Function

Function DbSave() 
    Dim pAP010M 
    Dim lngRows , itemRows
    Dim lGrpcnt
    DIM strVal 
    Dim strDel
    Dim tempItemSeq
    Dim iColSep 
    Dim iRowSep   
    

    DbSave = False                                                          
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF
   
	With frm1
		.txtFlgMode.value = lgIntFlgMode									
		.txtMode.value = Parent.UID_M0002
		

	End With

    lGrpCnt = 1
    strVal = ""
    strDel = ""
    
    iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	
    
    ggoSpread.Source = frm1.vspdData
    With frm1.vspdData
	    
    For lngRows = 1 To .MaxRows
    		
		.Row = lngRows
		.Col = 0

		If .Text = ggoSpread.InsertFlag Then
			strVal = strVal & "CREATE" & iColSep & lngRows & iColSep	
		ElseIf .Text = ggoSpread.UpdateFlag Then
			strVal = strVal & "UPDATE" & iColSep & lngRows & iColSep	
		ElseIf .Text = ggoSpread.DeleteFlag Then
			strDel = strDel & "DELETE" & iColSep & lngRows & iColSep
		End If
	
		Select Case .Text
		    
		    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

		        .Col = C_ACCT_Cd	'1
		        strVal = strVal & Trim(.Text) & iColSep
		            
		        .Col = C_DstbFctr_Cd		'2
		        strVal = strVal & Trim(.Text) & iRowSep

		        
		        lGrpCnt = lGrpCnt + 1
		        
		    Case ggoSpread.DeleteFlag
				
		        .Col = C_ACCT_Cd	'1
		        strDel = strDel & Trim(.Text) & iRowSep


			lGrpcnt = lGrpcnt + 1             
		End Select

    Next

    End With
	
    frm1.txtMaxRows.value = lGrpCnt-1	
    frm1.txtSpread.value =  strDel & strVal	

    lGrpCnt = 1
    strVal = ""
    strDel = ""
    
    ggoSpread.Source = frm1.vspdData3

    With frm1.vspdData3 

    	For itemRows = 1 To .MaxRows
		
	  	.Row = itemRows
		.Col = 0 
		Select Case .Text
			  
		Case ggoSpread.DeleteFlag
			strDel = strDel & "D" & iColSep & itemRows & iColSep
			.Col = 1 
			strDel = strDel & Trim(.Text) & iColSep
			.Col = 3 
			strDel = strDel & Trim(.Text) & iRowSep
				        
			lGrpCnt = lGrpCnt + 1
	    Case ggoSpread.InsertFlag
			.Col = 2  
			if .text = "1" then
				strVal = strVal & "C" & iColSep & itemRows & iColSep
				.Col = 1 
				strVal = strVal & Trim(.Text) & iColSep
				.Col =  3
				strVal = strVal & Trim(.Text) & iRowSep
						
				lGrpCnt = lGrpCnt + 1
			end if		
	        End Select
	 Next
	    
    End With

    frm1.txtMaxRows3.value = lGrpCnt-1	
    frm1.txtSpread3.value =   strDel & strVal
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
        
    DbSave = True                                                           
    
End Function

Function DbSaveOk()	

	 Call InitVariables
       frm1.vspdData.maxrows = 0
       frm1.vspdData2.maxrows = 0
	FncQuery

End Function

Function DbDelete()
	Dim strVal
	
    Err.Clear

    IF LayerShowHide(1) = False Then
		Exit Function
	END IF    

	DbDelete = False	
	
	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003
    	strVal = strVal & "&txtTempGlNo=" & Trim(frm1.txtTempGlNo.value)
    	strVal = strVal & "&txtDeptCd=" & Trim(frm1.txtDeptCd.value)

	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True 

End Function

Function DbDeleteOk()	
	Call FncNew()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<HTML>
<HEAD>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>계정별배부규칙등록</font></td>
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
								<TD CLASS="TD5" NOWRAP>버전</TD>
								<TD CLASS="TD6" COLSPAN=3><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtVerCd" SIZE=10 MAXLENGTH=3 tag="12XXXU" ALT="버전"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVerCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(Trim(frm1.txtVerCd.Value), 0)">
									 
								</TD>
								<TD CLASS="TD5" NOWRAP>코스트센타</TD>
								<TD CLASS="TD6" COLSPAN=3><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=10 tag="12XXXU" ALT="코스트센타"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCostCd.Value, 1)">
									 <INPUT NAME="txtCostNM" MAXLENGTH="25" SIZE=25 STYLE="TEXT-ALIGN:left" ALT ="코스트센타명" tag="14X">
								</TD>
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR >
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD WIDTH="65%" HEIGHT="100%">
									<script language =javascript src='./js/c1704ma1_vaSpread1_vspdData.js'></script>
								</TD>

								<TD WIDTH="35%" HEIGHT="100%">
									<script language =javascript src='./js/c1704ma1_vaSpread1_vspdData2.js'></script>
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
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC="" WIDTH="100%" HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3 tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=hidden NAME="htxtVerCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=hidden NAME="htxtCostCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=hidden NAME="hItemCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=hidden NAME="hFocusFlag" tag="24" TABINDEX= "-1">
<INPUT TYPE=hidden NAME="txtCommandMode" tag="24" TABINDEX= "-1">

<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX= "-1">

<INPUT TYPE=hidden NAME="txtMaxRows3" tag="24" TABINDEX= "-1">

<script language =javascript src='./js/c1704ma1_I155961126_vspdData3.js'></script>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

