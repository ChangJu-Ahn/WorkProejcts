<Script Language=VBScript Runat=Server>
'==============================================================================
'  ȭ   ��  �� : File UpLoad Common Constant Module
'  ��       �� : File UpLoad�� �������� ���Ǵ� �Լ� ��� 
'  ���� �ۼ��� : ���¹� 
'  ���� �ۼ��� : ���¹� 
'  ���� �ۼ��� : 2000�� 04�� 24�� 
'  ���� �ۼ��� : 2001-03-22
'  ���� �̷�   : Company Code�� Upload Directory�� ��� 
'  ���� ������ : ������ 
'  ��     ��   : ������ ���α׷� ����ڰ� �� ����� ������ �� ���� 
'==============================================================================

'==============================================================================
' ���� Web Site�� ���� ��θ� ��´�.
'==============================================================================
Function GetPhysicalUpLoadPath()

	strPathInfo = Request.ServerVariables("PATH_INFO")
	MyPos = Instr(2, strPathInfo, "/", 1)   
	strPhysicalPath = Server.MapPath(Left(strPathInfo, MyPos))
	if Right(strPhysicalPath, 1) <> "\" Then
'		strPhysicalPath = strPhysicalPath & "\UpLoad\" & "COMPANYCODE001" & "\"	     '��: ���� ������ (UpLoad �����ؿ� ȸ���ڵ� ����)
		strPhysicalPath = strPhysicalPath & "\UpLoad\" & gDBServer & "_" & gDatabase & "\" & gCompany & "\"				 '��: Global���� ���� Setting���� ���� ����� 
	else
'		strPhysicalPath = strPhysicalPath & "UpLoad\" & "COMPANYCODE001" & "\"
		strPhysicalPath = strPhysicalPath & "UpLoad\" & gDBServer & "_" & gDatabase & "\" & gCompany & "\"
	End if		
	
	GetPhysicalUpLoadPath = strPhysicalPath
	
End Function

'==============================================================================
' File UpLoad�� ��� ��θ� ��´�.
'==============================================================================
Function GetLogicalUpLoadPath()

	strPathInfo = Request.ServerVariables("PATH_INFO")
	MyPos = Instr(2, strPathInfo, "/", 1)   
	strPathInfo = Left(strPathInfo, MyPos)
	
	if Right(strPathInfo, 1) <> "/" Then
'		strPathInfo = strPathInfo & "/UpLoad/" & "COMPANYCODE001" & "/"	     '��: ���� ������ (UpLoad �����ؿ� ȸ���ڵ� ����)
		strPathInfo = strPathInfo & "/UpLoad/" & gDBServer & "_" & gDatabase & "/" & gCompany & "/"				 '��: Global���� ���� Setting���� ���� ����� 
	else
'		strPathInfo = strPathInfo & "UpLoad/" & "COMPANYCODE001" & "/"
		strPathInfo = strPathInfo & "UpLoad/" & gDBServer & "_" & gDatabase & "/" & gCompany & "/"
	End if		

	GetLogicalUpLoadPath = strPathInfo
	
End Function

'==============================================================================
' ���� ȭ���� Ȯ����� ��´�.
'==============================================================================
Function GetFileExt(ByVal FileName) 
	Dim DotPos, SepPos

	DotPos = InStrRev(FileName, ".", -1, 1)
	SepPos = InStrRev(FileName, "\", -1, 1)
		
    If DotPos > 0 And DotPos > SepPos Then
       GetFileExt = Mid(FileName, DotPos)
    Else
    	GetFileExt = ""
    End If
    
End Function

'==============================================================================
' ComProxy Error�� ���� UpLoad�� File ���� �� ���� �ʱ�ȭ 
'==============================================================================
Sub DeleteFile(ByVal FileName)
	strRet = objRequest.DeleteFile(FileName) 
    Set objRequest = Nothing											'��: File Upload Component Unload
End Sub

'==============================================================================
' �����̸����� ���� �� ���� Ư�� ���ڸ� "-"�� �����Ѵ�.
'==============================================================================

Function ConvFileName(lsFileName)
	Dim iPos
	Dim sChar
	
	ConvFileName = ""		
	
	'-------------------------------------
	' \,|,/,:,*,?,",<,>  ==> "-"
	'-------------------------------------
	For iPos = 1 To Len(lsFileName)
		sChar = Mid(lsFileName,iPos,1)
		If sChar = "\" Or sChar = "|" Or sChar = "/" Or sChar = ":" _
			Or sChar = "*" Or sChar = "?" Or sChar = "<" Or sChar = ">" Then
				
			Call Replace(lsFileName,sChar,"-")
		End If
	Next
	
	ConvFileName = lsFileName
End Function
		
</Script>
