Dim gClientX
Dim gClientY
Dim gLookUpEnable
Dim gMouseClickStatus
'========================================================================================
' Function Name : Document_onKeyDown
' Function Desc : hand all event of key down
'========================================================================================
Function Document_onKeyDown()
	Dim objEl, KeyCode, iLoc
	Dim boolMinus, boolDot
	
	On Error Resume Next
	
	Document_onKeyDown = True	
	Set objEl = window.event.srcElement
	KeyCode   = window.event.keycode
    
    If window.event.ctrlKey = True Then
       Select Case KeyCode
          Case 71        ' "G"
                   Call HandleTabs()
                   Document_onKeyDown = False                   
                   Exit Function       
          Case 75        ' "K"
                   If Mid(gToolBarBit,3,1) = "1" Then
                      Call MainNew()
                   End If   
                   Document_onKeyDown = False                   
                   Exit Function       
          Case 77        ' "M"
                   If Mid(gToolBarBit,7,1) = "1" Then
                      Call MainDeleteRow()
                   End If
                   Document_onKeyDown = False                   
                   Exit Function       
          Case 81        ' "Q"
                   If Mid(gToolBarBit,6,1) = "1" Then
                      Call MainInsertRow()
                   End If   
                   Document_onKeyDown = False                   
                   Exit Function       
          Case 85       ' "U"
                   If Mid(gToolBarBit,5,1) = "1" Then
                      Call MainSave()
                   End If   
                   Document_onKeyDown = False                   
                   Exit Function
          Case 86       ' "V"
          Case 88       ' "X"
          Case Else
                   Exit Function
        End Select 
    End If
    
	Select Case KeyCode	
		Case 8		' In case of BS key 
			Select Case UCase(objEl.tagName)
                     Case "SELECT","OBJECT"
                           Document_onKeyDown = False
                           Exit Function
                     Case "TEXTAREA"
                     Case "INPUT"
                           If UCase(objEl.TYPE) <> "BUTTON" Or UCase(objEl.TYPE) <> "RADIO" Then
                              If Left(objEl.getAttribute("tag"),1)     = "2" Then
                                 lgBlnFlgChgValue = True	
                              ElseIf Left(objEl.getAttribute("tag"),1) = "3" Then
                                     lgBlnFlgChgValue = True	
                              End If
                           End If	
                     Case Else
                          If parent.name = "frToolbar" Then
                             Document_onKeyDown = False
                             Exit Function
                          Else
                             Document_onKeyDown = False
                             Exit Function
                         End If
                         
            End Select
			If UCase(objEl.className) = "PROTECTED" Then
               Document_onKeyDown = False
			End If
		Case 9    'Tab 
            Exit Function
		Case 13		' Enter Key: Used as Query in Condition
			   If parent.name = "frToolbar" Then
			      If Left(objEl.getAttribute("tag"),1) = "1" Then
				     If UCase(Mid(gStrRequestMenuID,6,1)) = "O" Or  Mid(gToolBarBit,2,1) = "1" Then
				        If UCase(objEl.tagName) = "OBJECT" And CheckOCXQuery = False Then
				        Else
                           Call MainQuery()
                           Document_onKeyDown = False
                        End If   
  				     End If
					 Exit Function
				  End If					
			   End If
		Case 16   'Shift
            Exit Function
		Case 37   'Left
            Exit Function
		Case 38   'Up
            Exit Function
		Case 39   'Right
            Exit Function
		Case 40   'Down
            Exit Function            
		Case 116	' Process F5(Refresh) Event Differently depending on current focus status
			If parent.name = "frToolbar" Then
				window.event.keycode = 9
				Document_onKeyDown = False
				Call window.location.reload()
			End If
			Exit Function
		Case 118  'F7
            Call document.all(window.event.srcElement.sourceIndex+1).onclick
		Case 123  'F12
			Parent.Focus
            Call Parent.MakeF12KeyPressed			
			Document_onKeyDown = False
			Exit Function	
		Case Else
                If UCase(UCN_PROTECTED) <> "" Then
                   If UCase(UCN_PROTECTED) = UCase(objEl.getAttribute("className")) Then
                      Document_onKeyDown = False
                      Exit Function
                   End If   
				End If				
				
				Select Case UCase(Left(objEl.getAttribute("tag"),1))
				   Case "1"
				   
				   Case "2"
                             If UCase(UCN_PROTECTED) <> UCase(objEl.getAttribute("className")) Then
                                If UCase(objEl.tagName) = "OBJECT" Then
                                Else
                                   lgBlnFlgChgValue = True
                                End If   
                             End If   
				   Case "3"
                             If UCase(UCN_PROTECTED) <> UCase(objEl.getAttribute("className")) Then
                                If UCase(objEl.tagName) = "OBJECT" Then
                                Else
                                   lgBlnFlgChgValue = True
                                End If   
                             End If   
				   Case "X"
                             If UCase(UCN_PROTECTED) <> UCase(objEl.getAttribute("className")) Then
                             '   lgBlnFlgChgValue = True
                             End If   
				   Case Else
                             Document_onKeyDown = False
               End Select           
  				   
	End Select
	
End Function

Function Document_OnSelectStart()
    
    If UCase(window.event.srcElement.tagName) = "TEXTAREA" Then
       Exit Function
    End If
    
    If UCase(window.event.srcElement.tagName) = "INPUT" Then
       If UCase(window.event.srcElement.TYPE) = "TEXT"  Then
          Exit Function
       End If
    End If
    
    Document_onselectstart  = False
    
End Function

'========================================================================================
' Function Name : document_onmouseover
' Function Desc : display full value of object in window status bar 
'========================================================================================
Function Document_onMouseOver()
	On Error Resume Next

	Select Case UCASE(window.event.srcElement.tagName)
	   Case "INPUT"
		 window.status = window.event.srcElement.value
	   Case "SELECT"
		 window.status = window.event.srcElement.options(window.event.srcElement.selectedIndex).text
	   Case "OBJECT"
	      If UCase(window.event.srcElement.getAttribute("title")) = "FPDATETIME" Or _
	         UCase(window.event.srcElement.getAttribute("title")) = "FPDOUBLESINGLE" Then
             window.status = window.event.srcElement.text
          End If   
    End Select 

    Err.Clear

End Function

'========================================================================================
' Function Name : document_onmousedown
' Function Desc : show pressed button when you press calendar button
'========================================================================================
Sub Document_onMouseDown()
    Dim leftFrameWidth

	On Error Resume Next

	If Left(Top.frMain.cols,1) <> 0 Then
       leftFrameWidth = Top.Frames(0).frm2.scrollWidth + 2
	Else
       leftFrameWidth = 1
	End If	

	gClientX = window.event.clientX + leftFrameWidth
	gClientY = window.event.clientY + 78 

	If gMouseClickStatus = "SPCR" Then
	   gMouseClickStatus = "SPCRP"
	   Call ShowSpreadRPopup	
	End If 	
	If gMouseClickStatus = "SP1CR" Then
	   gMouseClickStatus = "SP1CRP"
	   Call ShowSpreadRPopup	
	End If 
	If gMouseClickStatus = "SP2CR" Then
	   gMouseClickStatus = "SP2CRP"
	   Call ShowSpreadRPopup	
	End If 
	If gMouseClickStatus = "SP3CR" Then
	   gMouseClickStatus = "SP3CRP"
	   Call ShowSpreadRPopup	
	End If 
	If gMouseClickStatus = "SP4CR" Then
	   gMouseClickStatus = "SP4CRP"
	   Call ShowSpreadRPopup	
	End If 
	If gMouseClickStatus = "SP5CR" Then
	   gMouseClickStatus = "SP5CRP"
	   Call ShowSpreadRPopup	
	End If 
	If gMouseClickStatus = "SP6CR" Then
	   gMouseClickStatus = "SP6CRP"
	   Call ShowSpreadRPopup	
	End If 
	If gMouseClickStatus = "SP7CR" Then
	   gMouseClickStatus = "SP7CRP"
	   Call ShowSpreadRPopup	
	End If 
	
    If UCase(window.event.srcElement.tagName) = "IMG" Then 	
       If InStr(window.event.srcElement.src,"btnPopup") > 0 Then
          window.event.srcElement.src = GetImgPath(window.event.srcElement.src) & "btnPopup_dn.gif"
       End If   
	End If
End Sub

'========================================================================================
' Function Name : document_onmouseup
' Function Desc : show un-pressed button when you press up calendar button
'========================================================================================
Sub Document_onMouseUp()

	On Error Resume Next

    If UCase(window.event.srcElement.tagName) = "IMG" Then 		
       If InStr(window.event.srcElement.src,"btnPopup") > 0 Then
          window.event.srcElement.src = GetImgPath(window.event.srcElement.src) & "btnPopup.gif"
       End If   
    End If

End Sub

'========================================================================================
' Function Name : document_onmouseout
' Function Desc : same as mouse up enevt
'========================================================================================
Sub Document_onMouseOut()
	On Error Resume Next
	
    If UCase(window.event.srcElement.tagName) = "IMG" Then 		
       If InStr(window.event.srcElement.src,"btnPopup") > 0 Then
          window.event.srcElement.src = GetImgPath(window.event.srcElement.src) & "btnPopup.gif"
       End If    
    End If
	
	window.status = ""

End Sub

'========================================================================================
' Function Name : PopUpMouseOver
' Function Desc : this sub procedure handle  lookup event
'               : gLookUpEnable = False  means that program can not call lookup procedure
'========================================================================================
Sub PopUpMouseOver()
    gLookUpEnable = False    
End Sub

'========================================================================================
' Function Name : PopUpMouseOver
' Function Desc : this sub procedure handle  lookup event
'               : gLookUpEnable = True   means that program can     call lookup procedure
'========================================================================================
Sub PopUpMouseOut()
    gLookUpEnable= True
End Sub

'========================================================================================
Sub vspdData0_KeyDown(KeyCode, shift)
    On Error Resume Next
    Call HandleSpreadSheetKeyEvent(KeyCode, shift)    
End Sub

'========================================================================================
Sub vspdData_KeyDown(KeyCode, shift)
    On Error Resume Next
    Call HandleSpreadSheetKeyEvent(KeyCode, shift)
End Sub

'========================================================================================
Sub vspdData1_KeyDown(KeyCode, shift)
    On Error Resume Next    
    Call HandleSpreadSheetKeyEvent(KeyCode, shift)
End Sub

'========================================================================================
Sub vspdData2_KeyDown(KeyCode, shift)
    On Error Resume Next
    Call HandleSpreadSheetKeyEvent(KeyCode, shift)
End Sub

'========================================================================================
Sub vspdData3_KeyDown(KeyCode, shift)
    On Error Resume Next    
    Call HandleSpreadSheetKeyEvent(KeyCode, shift)
End Sub

'========================================================================================
Sub vspdData4_KeyDown(KeyCode, shift)
    On Error Resume Next    
    Call HandleSpreadSheetKeyEvent(KeyCode, shift)
End Sub

'========================================================================================
Sub vspdData5_KeyDown(KeyCode, shift)
    On Error Resume Next    
    Call HandleSpreadSheetKeyEvent(KeyCode, shift)
End Sub

'========================================================================================
Sub vspdData6_KeyDown(KeyCode, shift)
    On Error Resume Next    
    Call HandleSpreadSheetKeyEvent(KeyCode, shift)
End Sub

'========================================================================================
Sub vspdData7_KeyDown(KeyCode, shift)
    On Error Resume Next    
    Call HandleSpreadSheetKeyEvent(KeyCode, shift)

End Sub

'========================================================================================
Sub vspdData8_KeyDown(KeyCode, shift)
    On Error Resume Next
    Call HandleSpreadSheetKeyEvent(KeyCode, shift)

End Sub

'========================================================================================
Sub vspdData9_KeyDown(KeyCode, shift)
    On Error Resume Next    
    Call HandleSpreadSheetKeyEvent(KeyCode, shift)
End Sub

'========================================================================================
Sub HandleSpreadSheetKeyEvent(KeyCode, shift)

    Select Case KeyCode
        Case 71        ' "G"
               If shift = 2 Then
                  Call HandleTabs()
               End If
        Case 75        ' "K"
               If shift = 2 Then
                  If Mid(gToolBarBit,3,1) = "1" Then
                     Call MainNew()
                  End If   
               End If
        Case 77        ' "M"
               If shift = 2 Then
                  If Mid(gToolBarBit,7,1) = "1" Then
                     Call MainDeleteRow()
                  End If   
               End If
        Case 81        ' "Q"
               If shift = 2 Then
                  If Mid(gToolBarBit,6,1) = "1" Then
                     Call MainInsertRow()
                  End If   
               End If
        Case 85       ' "U"
               If shift = 2 Then
                  If Mid(gToolBarBit,5,1) = "1" Then
                     Call MainSave()
                  End If   
               End If
'        Case 86        ' "V"
'               If Shift = 2 Then
'                   ggoSpread.Source = gActiveSpdSheet   
'                   ggoSpread.CPasteRepeatedSpreadData 
'               End If
        Case 123        'F12
              Parent.Focus
              Call Parent.MakeF12KeyPressed			
        Case 113        'F2 
              Call GoToCondition(Document)
	End Select

End Sub

'======================================================================================================
' Name : uni2KMenu_Click
' Desc : This sub call FncSplitColumn function in business module
'======================================================================================================
Sub uni2KMenu_Click(ItemNumber)
	Dim strMenuID

    strMenuID = UCase(uni2KMenu.ItemKey(ItemNumber))
    
    Select Case strMenuID
         Case "MNUAPPENDROW"
              FncInsertRow
         Case "MNUDELETEROW"
              FncDeleteRow
         Case "MNUCANCELROW"
              FncCancel
         Case "MNUFIXCOL"
              FncSplitColumn
'         Case "MNUIMPORTFROMEXCEL"
'              FncImportFormExcel    ' 2002-11-11 컬럼이동관련 추가 (김인태)
'         Case "MNUSORT"
'               PopSortPopup              
'         Case "MNUSAVE"
'               PopSaveSpreadColumnInf
'         Case "MNURESET"
'               PopRestoreSpreadColumnInf     
    End Select 
    
    
     
    If Instr(1,strMenuID,"MNUSHOWHIDECOLUMN") > 0 Then
       strMenuID = Replace(strMenuID,"MNUSHOWHIDECOLUMN","")
       Call PopMakeHiddenColumn(strMenuID,Not uni2KMenu.checked(ItemNumber))
    End If
    
End Sub

'======================================================================================================
' Name : ShowSpreadRPopup
' Desc : Show popup menu of spreadsheet 
'======================================================================================================
Sub ShowSpreadRPopup()
    Dim iRet
    Dim upperIDX
    Dim ii
    Dim iHiddenType						' 2002-11-11 컬럼이동관련 추가 (김인태)
    Dim iTemp							' 2002-11-11 컬럼이동관련 추가 (김인태)
    
    On Error Resume Next
      
   	uni2KMenu.Clear
	uni2KMenu.CSystemUserColor        = "SYSTEM"
	uni2KMenu.CMenuBgColor            = &HB36801
	uni2KMenu.CMenuTextColor          = &HFFFFFF
	uni2KMenu.CMenuHighLightTextColor = RGB(20,40,135)

    If Mid(gToolBarBit,6,1) = "1" Then   	
   	   uni2KMenu.AddItem "행추가", , , , , , , "mnuAppendRow"
   	End If   
    If Mid(gToolBarBit,7,1) = "1" Then   	
   	   uni2KMenu.AddItem "행삭제", , , , , , , "mnuDeleteRow"
   	End If   
    If Mid(gToolBarBit,8,1) = "1" Then   	
   	   uni2KMenu.AddItem "취소"  , , , , , , , "mnuCancelRow"
   	End If   
   	
	uni2KMenu.AddItem "열고정"    , , , , , , , "mnuFixCol"
	'uni2KMenu.AddItem "데이타입력", , , , , , , "mnuImportFromExcel"
	'                 1                2 3 4 5 6 7 8
	uni2KMenu.Store    "User"
		
    uni2kMenu.Restore "User"
    iRet = uni2kMenu.ShowPopupMenu(gClientX ,gClientY)
	
End Sub

Sub HandleTabs()
    On Error Resume Next

    If gIsTab = "Y" Then
       gPageNo = gPageNo + 1
       If gPageNo > gTabMaxCnt Then
          gPageNo = 1
       End If
       Select Case gPageNo
             Case 1  : Call ClickTab1()
             Case 2  : Call ClickTab2()
             Case 3  : Call ClickTab3()
             Case 4  : Call ClickTab4()
             Case 5  : Call ClickTab5()
             Case 6  : Call ClickTab6()
             Case 7  : Call ClickTab7()
             Case 8  : Call ClickTab8()
       End Select   
    End If   
End Sub
