'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' <<<<<<<<Event ���� �Լ�>>>>>>>>
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'========================================================================================
' Function Name : Document_onKeyDown
' Function Desc : hand all event of key down
'========================================================================================
Function Document_onKeyDown()

	Set objEl = window.event.srcElement
	KeyCode = window.event.keycode
	Select Case KeyCode	
		Case 13		' Enter Key: Used as Query in Condition
			If Left(objEl.getAttribute("tag"),1) = "1" Then
				Call DbQuery(1)
			end if
		Case 33
			if IsObject(Grid1) then
				Call Grid1.PrePages()
			end if
		Case 34
			if IsObject(Grid1) then
				Call Grid1.NextPages()
			end if
		
	End Select
	Set objEl = nothing
End Function 

'========================================================================================
' Function Name : Window_onLoad
' Function Desc : ȭ�� ó�� ASP�� Ŭ���̾�Ʈ�� Load�� �� �����ؾ� �� ���� ó�� 
'========================================================================================
Sub Window_onLoad()
    Dim iDx
    Call Form_Load()
    window.status      = ""
    Set gActiveElement = document.activeElement

End Sub

'========================================================================================
' Function Name : Window_onUnLoad
' Function Desc : ������ ��ȯ�̳� ȭ���� ���� ��� �����ؾ� �� ���� ó�� 
'========================================================================================
Sub Window_onUnLoad()
	Call Form_UnLoad()
 	Set gActiveElement = Nothing
End Sub

