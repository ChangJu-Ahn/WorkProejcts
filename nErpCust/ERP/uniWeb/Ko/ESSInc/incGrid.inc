'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' <<<<<<<<Grid Class>>>>>>>>
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Dim UpdateFlag
Dim InsertFlag
Dim DeleteFlag

InsertFlag       = "입력"
UpdateFlag       = "수정"
DeleteFlag       = "삭제"

Class Grid
    Public gArrData     '쿼리 결과를 저장하는 배열 
    Public MaxRows      '쿼리 결과의 Row수(실제는 +1)
    Public MaxPages     '쿼리 결과의 Page수 
    Public SheetMaxrows '한 page내의 Row수 
    Public MaxCols      '한 page내의 Col수(실제는 +1)
    Public SheetData    '한 page내의 데이터(EDIT 취소시 비교값을 임시로 저장)
    Public Source       'Oppr대상 frm을 지정함(대상프레임부터form까지)
    Public ActiveRow    '현재 포인트된 Row
    Public MaxQueryRows '한번에 쿼리할 Record의 수 
    
    Dim elmCnt
    
    Sub SSSetData(RetData)                      '// MB에서 데이터를 넘겨받는다.
        Dim Temp, offset
        Err.Clear
        Set gArrData = Nothing
        Dim i,j,arrDataRow,arrDataCol

        If RetData="" Then Exit Sub
        arrDataRow = Split(RetData,Chr(12))

        MaxRows = Ubound(arrDataRow,1)-1

        Temp = (MaxRows+1) Mod SheetMaxrows
        If Temp=0 Then
	        offset = 0
	    Else
	        offset = 1
	    End If
        MaxPages = (MaxRows+1)\SheetMaxrows + offset
        Redim gArrData(MaxRows,MaxCols)

        For i=0 To MaxRows
            arrDataCol = Split(arrDataRow(i),Chr(11))
            For j=0 To MaxCols

                If j=0 Or j=MaxCols Then
                     gArrData(i,j)=i+1
                Else
                    if arrDataCol(j) = "" then
                    else
                        gArrData(i,j)=arrDataCol(j)
                    end if
                End If
            Next
        Next
		Source.grid_totpages.Value = MaxPages
    End Sub
    Function ShowData(objFrameToForm,pPage)     '//gArrData에서 받은 
        Err.Clear
        ShowData = False
        Dim iRow,iCol,objList
   			iRow = SkipPage(pPage - 1)

    	    iCol = 0
    	    SheetData = ""

            with objFrameToForm
            	For elmCnt = 0 to .length - 1
            		Set objList = .elements(elmCnt)
            		If Ucase(objList.getAttribute("flag")) = "SPREADCELL" then
                		If iRow = MaxRows + 1 Then
            			    objList.value = ""
                            objList.readonly = true
                    		objList.tabindex = "-1"
                    		
                    		If 	UCase(objList.tagName) = "INPUT" Then
                                If 	UCase(objList.Type) = "CHECKBOX" Then
                    		        objList.style.display="none"
                    		    End If
                   		    End If
            			    SheetData = SheetData & Chr(11) & ""
            			Else
            			    objList.value = gArrData(iRow,iCol)
            			    SheetData = SheetData & Chr(11) & objList.value
                		    iCol = iCol + 1
                		    If iCol = MaxCols Then 
                		        iRow= iRow +1
                		        iCol=0
                		    End If
                    		If 	UCase(objList.tagName) = "INPUT" Then
                          		If 	UCase(objList.Type) = "CHECKBOX" Then
			             		    objList.style.display=""
				         		End If
			         		End If
                		    Call LockElement(objList)
                		End If
            		End if
                	Set objList = Nothing
            	Next
            End With
        ShowData = True
    End Function
    Function Clear(objFrameToForm,pPage)     '//gArrData에서 받은 
        Err.Clear
        Clear = False
        Dim iRow,iCol,objList
   			iRow = SkipPage(pPage - 1)
    	    iCol = 0
    	    SheetData = ""

            with objFrameToForm
            	For elmCnt = 0 to .length - 1
            		Set objList = .elements(elmCnt)
            		If Ucase(objList.getAttribute("flag")) = "SPREADCELL" then
            			    objList.value = ""
                            objList.readonly = true
                    		objList.tabindex = "-1"
                    		If 	UCase(objList.tagName) = "INPUT" Then
	                    		If 	objList.Type = "checkbox" Then
		                		    objList.style.display="none"
			            		End If
		            		End If
            			    SheetData = SheetData & Chr(11) & ""
            		End if
                	Set objList = Nothing
            	Next
            End With
        Clear = True
    End Function


    Function SkipPage(pPage)                '//페이지의 첫Row인덱스를 찾는다.
        Err.Clear
        Dim iDx
        SkipPage = 0
        SkipPage = pPage * SheetMaxrows   
    End Function

    Sub SetPageCounter( pCombo,TotPages)     '//Setting(Resetting) Page Selection Combo & Total pages
        Err.Clear
        Dim ComboLength, iDx                 '//Reset시는 TotPages=0으로 Call하면됨 
        If TotPages=0 Then
            ComboLength= pCombo.length-1
            For iDx = 0 To ComboLength
                pCombo.Remove(0)
            Next
        End If
        For iDx = 1 To TotPages
            Call SetCombo(pCombo,iDx-1,iDx)
        Next
        Source.grid_totpages.Value = MaxPages
    End Sub

    Function PrePages()                      '//왼쪽 조회버튼을 눌렀을때 
        Err.Clear
        If ChkChange() Then Exit Function
        If Source.grid_page.Value = "" Or Source.grid_totpages.Value = "" Then Exit Function
        If Source.grid_page.Value = "1" Then
            Call DisplayMsgBox("900011","X","X","X")
            Exit Function
        Else
            Call ShowData(Source,Source.grid_page.Value-1)
            Source.grid_page.Value = CDbl(Source.grid_page.Value)-1
            Call DbQueryOk()
        End If
    End Function

    Function NextPages()                    '//오른쪽 조회버튼을 눌렀을때 
        Err.Clear
        If ChkChange() Then Exit Function
        If Source.grid_page.Value = "" OR Source.grid_totpages.Value = "" Then Exit Function
        If CDbl(Source.grid_page.Value) = CDbl(Source.grid_totpages.Value) Then
            Call DisplayMsgBox("900012","X","X","X")
            Exit Function
        Else
            Call ShowData(Source,Source.grid_page.Value+1)
            source.grid_page.value = CDbl(Source.grid_page.Value)+1
            Call DbQueryOk()
        End If
    End Function

    Function Chkchange()                    '//변경된 데이터가 있을때 
        Err.Clear
        Dim elmCnt,objList,Reti,VsData

        Chkchange=False
        With Source
        	For elmCnt = 0 to .length - 1
        		Set objList = .elements(elmCnt)
        		If Ucase(objList.getAttribute("flag")) = "SPREADCELL" then
        			VsData= VsData & Chr(11) & objList.value
        		end if
            	Set objList = Nothing
        	Next
        End With
        If VsData<> SheetData And SheetData<>"" Then
		    Reti = DisplayMsgBox("800442", VB_YES_NO,"X","X")
            If Reti = vbYes Then
                Call DbSave()
                Exit Function
            ElseIf  Reti = vbNo Then
                Chkchange=False
                Exit Function
            Else
                Chkchange=True
                Exit Function
            End If
        End If
    End Function
    
    Sub PageSelect()
        Err.Clear
        If ChkChange() Then Exit Sub
        Call ShowData(Source,Source.grid_page.Value)
    End Sub

    Sub DeleteClick()
        Err.Clear
        If ActiveRow <> 0 Then
            If document.all(CStr(ActiveRow)).value="" Then
               document.all(CStr(ActiveRow)).value=DeleteFlag
               ActiveRow = 0
            End If
        End If
    End Sub

    Sub SetUpdateFlag(pRow)
        Err.Clear
        ActiveRow = pRow
        document.all(CStr(ActiveRow)).value=UpdateFlag
    End Sub

    Sub CancelClick()
        Err.Clear
        Dim arrData
        If SheetData<>"" Then 
            arrData = Split(Mid(frm1.txtCurrentData.value,2),Chr(11))
            with document.frm1
            	For elmCnt = 0 to .length - 1
            		Set objList = .elements(elmCnt)
            		If  Ucase(objList.name) = "SPREADCELL" then
            			If TargetCnt=Ubound(arrData)+1 Then
            			    objList.value = ""
            			Else
            			    objList.value = arrData(TargetCnt)
                		    TargetCnt = TargetCnt + 1
                		End If
                    
            		End if
                	Set objList = Nothing
            	Next
            End With
        End If
    End Sub
End Class
