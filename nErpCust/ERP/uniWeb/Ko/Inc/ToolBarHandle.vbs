    Dim gURLPath 
	Dim i_initToolBarVal			'ȭ�� �ʱ�ȭ�� Setting �� ToolBar ���� 
	Dim i_initToolVal	            'ȭ�� �ʱ�ȭ�� Setting �� ToolBar ���� 
	Dim i_compToolBarVal			'SetToolBar ������ ToolBar ���� 
	Dim i_layerToolBarVal			'ȭ�� Type�� ���� ToolBar �ʱ�ȭ ���� 
    Dim ggToolBarBit
' 1999. 3.11 �ֿ��� ���� 
' ���� ������ ���¸� �ٲٴ� ��� ���ʿ��ϰ� ������ ������ ���µ� �ٽ� �����ϹǷμ� �����ս��� ������ 
' �� Class�� üũ�� �׸����� ������ �ʿ���� ���ٵ��� �������� �ʴ°����� ��ü�ϹǷμ� �ð��� ����Ǿ��� 
'======================================================================================================
'	Function Name	: GenerateToolBar
'	Description	: Input String ������ ToolBar Button setting
'	Parameters	: 11�ڸ� String (Query, New, Delete, InsertR, DeleteR, Save, Prev, Next, Copy, Excel, Print) 
'	History		: 99.2.8 Created by Choe Young-tae
'======================================================================================================
	Function GenerateToolBar(pstrVal) 
		Dim i, tmpValue
		With frm1
		ggToolBarBit = pstrVal
		for i=1 to Len(pstrVal) 
			tmpValue = Mid(pstrVal,i,1)
			Select case i
				case 1
					if tmpValue = "0" then
						If .tbExplorer.className <> "disableIMG" Then 
                           .tbExplorer.style.cursor = ""
					       .tbExplorer.className = "disableIMG"
 							Call ChgDisImg("tbExplorer")
						End If
					else
						If .tbExplorer.className <> "enableIMG" Then
                           .tbExplorer.style.cursor = "hand"
						   .tbExplorer.className = "enableIMG"
							Call ChgGryImg("tbExplorer")
						End If
					end if
				case 2
					if tmpValue = "0" then
						If .tbQuery.className <> "disableIMG" Then 
                           .tbQuery.style.cursor = ""
						   .tbQuery.className = "disableIMG"
							Call ChgDisImg("tbQuery")
						End If
					else
						If .tbQuery.className <> "enableIMG" Then
                           .tbQuery.style.cursor = "hand"
						   .tbQuery.className = "enableIMG"
							Call ChgGryImg("tbQuery")
						End If
					end if
				case 3
					if tmpValue = "0" then
						If .tbNew.className <> "disableIMG" Then
                           .tbNew.style.cursor = ""
					       .tbNew.className = "disableIMG"
							Call ChgDisImg("tbNew")
						End If
					else
						If .tbNew.className <> "enableIMG" Then
                           .tbNew.style.cursor = "hand"
						   .tbNew.className = "enableIMG"
							Call ChgGryImg("tbNew")
						End If
					end if
				case 4
					if tmpValue = "0" then
						If .tbDelete.className <> "disableIMG" Then
                           .tbDelete.style.cursor = ""
					       .tbDelete.className = "disableIMG"
							Call ChgDisImg("tbDelete")
						End If
					else
						If .tbDelete.className <> "enableIMG" Then
                           .tbDelete.style.cursor = "hand"
                           .tbDelete.className = "enableIMG"
							Call ChgGryImg("tbDelete")
						End If
					end if
				case 5
					if tmpValue = "0" then
						If .tbSave.className <> "disableIMG" Then
                           .tbSave.style.cursor = ""
					       .tbSave.className = "disableIMG"
							Call ChgDisImg("tbSave")
						End If
					else
						If .tbSave.className <> "enableIMG" Then
                           .tbSave.style.cursor = "hand"
                           .tbSave.className = "enableIMG"
							Call ChgGryImg("tbSave")

						End If
					end if
				case 6
					if tmpValue = "0" then
						If .tbInsertRow.className <> "disableIMG" Then
                           .tbInsertRow.style.cursor = ""
					       .tbInsertRow.className = "disableIMG"
							Call ChgDisImg("tbInsertRow")
						End If
					else
						If .tbInsertRow.className <> "enableIMG" Then
                           .tbInsertRow.style.cursor = "hand"
                           .tbInsertRow.className = "enableIMG"
							Call ChgGryImg("tbInsertRow")
						End If
					end if
				case 7
					if tmpValue = "0" then
						If .tbDeleteRow.className <> "disableIMG" Then
                           .tbDeleteRow.style.cursor = ""
				           .tbDeleteRow.className = "disableIMG"
							Call ChgDisImg("tbDeleteRow")
						End If
					else
						If .tbDeleteRow.className <> "enableIMG" Then
                           .tbDeleteRow.style.cursor = "hand"
						   .tbDeleteRow.className = "enableIMG"
							Call ChgGryImg("tbDeleteRow")

						End if
					end if
				case 8
					if tmpValue = "0" then
						If .tbCancel.className <> "disableIMG" Then
                           .tbCancel.style.cursor = ""
					       .tbCancel.className = "disableIMG"
							Call ChgDisImg("tbCancel")
						End If
					else
						If .tbCancel.className <> "enableIMG" Then
                           .tbCancel.style.cursor = "hand"
                           .tbCancel.className = "enableIMG"
							Call ChgGryImg("tbCancel")

						End If
					end if
				case 9
					if tmpValue = "0" then
						If .tbPrev.className <> "disableIMG" Then
                           .tbPrev.style.cursor = ""
						   .tbPrev.className = "disableIMG"
							Call ChgDisImg("tbPrev")
						End If
					else
						If .tbPrev.className <> "enableIMG" Then
                           .tbPrev.style.cursor = "hand"
						   .tbPrev.className = "enableIMG"
							Call ChgGryImg("tbPrev")

						End If
					end if
				case 10
					if tmpValue = "0" then
						If .tbNext.className <> "disableIMG" Then
                           .tbNext.style.cursor = ""
				           .tbNext.className = "disableIMG"
							Call ChgDisImg("tbNext")
						End If
					else
	  					If .tbNext.className <> "enableIMG" Then
   	                       .tbNext.className = "enableIMG"
                           .tbNext.style.cursor = "hand"
							Call ChgGryImg("tbNext")

						End If
					end if
				case 11
					if tmpValue = "0" then
						If .tbCopy.className <> "disableIMG" Then
                           .tbCopy.style.cursor = ""
						   .tbCopy.className = "disableIMG"
							Call ChgDisImg("tbCopy")
						End If
					else
	  					If .tbCopy.className <> "enableIMG" Then
                           .tbCopy.className = "enableIMG"
                           .tbCopy.style.cursor = "hand"
						   Call ChgGryImg("tbCopy")
						End If   
					end if
				case 12
					if tmpValue = "0" then
						If .tbExcel.className <> "disableIMG" Then
                           .tbExcel.style.cursor = ""
						   .tbExcel.className = "disableIMG"
							Call ChgDisImg("tbExcel")
						End If
					else
						If .tbExcel.className <> "enableIMG" Then
                           .tbExcel.style.cursor = "hand"
                           .tbExcel.className = "enableIMG"
							Call ChgGryImg("tbExcel")
						End If
					end if
				case 13
					if tmpValue = "0" then
						If .tbPrint.className <> "disableIMG" Then
                           .tbPrint.style.cursor = ""
					       .tbPrint.className = "disableIMG"
							Call ChgDisImg("tbPrint")
						End If
					else
						If .tbPrint.className <> "enableIMG" Then
                           .tbPrint.style.cursor = "hand"
                           .tbPrint.className = "enableIMG"
							Call ChgGryImg("tbPrint")
						End If
					end if
				case 14
					if tmpValue = "0" then
						If .tbFind.className <> "disableIMG" Then
                           .tbFind.style.cursor = ""
						   .tbFind.className = "disableIMG"
							Call ChgDisImg("tbFind")
						End If
					else
						If .tbFind.className <> "enableIMG" Then
                           .tbFind.style.cursor = "hand"
						   .tbFind.className = "enableIMG"
						   Call ChgGryImg("tbFind")
						End If   
					end if
				case 15
					if tmpValue = "0" then
						If .tbHelp.className <> "disableIMG" Then
                           .tbHelp.style.cursor = ""
						   .tbHelp.className = "disableIMG"
							Call ChgDisImg("tbHelp")
						End If
					else
						If .tbHelp.className <> "enableIMG" Then
                           .tbHelp.style.cursor = "hand"
	                       .tbHelp.className = "enableIMG"
	 					   Call ChgGryImg("tbHelp")
	 					End If   
					end if
			End Select				
		Next
		
		End With
	End Function

'======================================================================================================
'	Function Name	: openToolBar
'	Description	: ������� ���Ѱ� ȭ���� Type�� ���� onload�� ToolBar Setting
'	Parameters	: pAuthLevel - ������� ���� ���� 
'			  pLayType - ȭ�� Ÿ�� 
'	History		: 99.2.8 Created by Kim Yongtae
'======================================================================================================
	Function openToolBar(pAuthLevel, pLayType)
		on error resume next
		dim i, tmpVal, tmpCompVal, retVal

		i_layerToolBarVal = pLayType
		
		Select case pAuthLevel			'�� ���� ���ؿ� ���� ToolBar ���� 
			case "A"
				i_initToolVal = "000000000000"
			case "B"
				i_initToolVal = "100000000000"
			case "C"
				i_initToolVal = "100000000110"
			case "D"
				i_initToolVal = "111111111111"
			case "E"
				i_initToolVal = "111111111111"
		End Select

		for i=1 to Len(i_initToolVal) 
			tmpVal = CDbl(Mid(i_initToolVal,i,1))
			tmpCompVal = CDbl(Mid(i_layerToolBarVal,i,1))

			If tmpVal * tmpCompVal = 0 Then
				retVal = retVal & Cstr(0)	'������ ���� ��� Disable	
			else
				retVal = retVal & Cstr(1)	'������ �ִ� ��� ȭ�� Type�� ���� �ʱ�ġ�� Setting
			End If
		next
		
		i_initToolBarVal = retVal
		i_compToolBarVal = retVal			'���� ���յ� ���ڸ� i_compToolBarVal�� �Ҵ� 
		Call GenerateToolBar(retVal)

	End Function

'======================================================================================================
'	Function Name	: initToolBar
'	Description	: ������� ���Ѱ� ȭ���� Type�� ���� ȭ�� �ʱ�ȭ�� ������ ToolBar�� ���·� Setting
'	Parameters	: i_initToolBarVal - ���ѿ� ���� ����� �� �ִ� Button�� ǥ�� (�ʱ�ȭ�� Open�� Setting ��)
'			  i_layerToolBarVal - ȭ�� Ÿ�Կ� ���� ToolBar�� �ʱ� Setting
'	History		: 99.2.8 Created by Kim Yongtae
'======================================================================================================
	Function initToolBar(pVal)
		Call GenerateToolBar(pVal)
		i_compToolBarVal = pVal
	End Function


'======================================================================================================
'	Function Name	: SetToolBar
'	Description	: Operation�� ���� �����ڰ� ToolBar�� �����ϱ� ���� �Լ� 
'	Parameters	: 11�ڸ� String 
'			  1 - Enable, 2 - Disable, 0 - �������			  	
'	History		: 99.2.8 Created by Kim Yongtae
'======================================================================================================
	Sub SetToolBar(pstrSetVal) 
		dim i, tmpVal, tmpCompVal, retVal
		on error resume next
		for i=1 to Len(pstrSetVal) 
			tmpVal = CDbl(Mid(pstrSetVal,i,1))
			'tmpCompVal = CDbl(Mid(i_compToolBarVal,i,1))

			'If tmpVal = 2 Then
			'	retVal = retVal & Cstr(tmpCompVal)	' ������ 2�ΰ��� ���������� Setting	
			'Else
				retVal = retVal & Cstr(tmpVal)
			'End If
		next
		i_compToolBarVal = retVal				'���� ���յ� ���ڸ� i_compToolBarVal�� �Ҵ� 
		'Call CompToolBar(i_compToolBarVal) 
		Call GenerateToolBar(retVal)		' ���߿� ���������� ���� 
	End Sub


'======================================================================================================
'	Function Name	: CompToolBar
'	Description	: Operation�� ���� �����ڰ� Setting�� ToolBar�� ���Ѱ� �� 
'			  ���������� ToolBar Setting
'	Parameters	: 11�ڸ� String (Query, New, Delete, InsertR, DeleteR, Save, Prev, Next, Copy, Excel, Print) 
'	History		: 99.2.8 Created by Kim Yongtae
'======================================================================================================
	Sub CompToolBar(pstrSetVal) 
		dim i, tmpVal, tmpCompVal, retVal
		on error resume next
		for i=1 to Len(pstrSetVal) 
			tmpVal = CDbl(Mid(pstrSetVal,i,1))
			'tmpCompVal = CDbl(Mid(i_initToolVal,i,1))

			'If tmpVal * tmpCompVal = 0 Then
			'	retVal = retVal & Cstr(0)		'������ ���� ��� 0(Disable)���� Setting
			'else
				retVal = retVal & Cstr(1)		'������ �ִ� ��� i_compToolBarVal�� ������ Setting
			'End If
		next
		i_compToolBarVal = retVal				'���� ���յ� ���ڸ� i_compToolBarVal�� �Ҵ� 
		Call GenerateToolBar(retVal) 
	End Sub

	Function document_onkeydown()
		If window.event.keycode = 116 Then
			If MsgBox("�ʱ� ȭ������ �ǵ��� ���ðڽ��ϱ�?", vbYesNo) = vbNo Then
				window.event.keycode = 9
				document_onkeydown = False
			End If
		End If
	End Function
	
'========================================================================================
' Function Name : RunMyBizASP
' Function Desc : �����Ͻ� ���� ASP�� Get ������� �����Ų��.
'========================================================================================
Sub RunMyBizASP(objIFrame, strURL)
	objIFrame.location.href = GetUserPath & strURL
End Sub

'========================================================================================
' Function Name : GetUserPath
' Function Desc : ���� ���丮 �н� �˾ƿ��� 
'========================================================================================
Function GetUserPath()
	If gURLPath = "" or isEmpty(gURLPath) Then
		Dim strLoc, iPos , iLoc, strPath
		strLoc = window.location.href
		iLoc = 1: iPos = 0
		Do Until iLoc <= 0						
			iLoc = inStr(iPos+1, strLoc, "/")
			If iLoc <> 0 Then iPos = iLoc
		Loop	
		gURLPath = Left(strLoc, iPos)
	End If
	GetUserPath = gURLPath
End Function
