<Script Language=VBScript Runat=Server>
'==============================================================================
'  화   일  명 : File UpLoad Common Constant Module
'  설       명 : File UpLoad시 공통으로 사용되는 함수 기술 
'  최초 작성자 : 강태범 
'  최종 작성자 : 강태범 
'  최초 작성일 : 2000년 04월 24일 
'  최종 작성일 : 2001-03-22
'  변경 이력   : Company Code를 Upload Directory로 사용 
'  변경 개발자 : 임현수 
'  주     석   : 업무별 프로그램 담당자가 이 모듈을 수정할 수 없음 
'==============================================================================

'==============================================================================
' 현재 Web Site의 절대 경로를 얻는다.
'==============================================================================
Function GetPhysicalUpLoadPath()

	strPathInfo = Request.ServerVariables("PATH_INFO")
	MyPos = Instr(2, strPathInfo, "/", 1)   
	strPhysicalPath = Server.MapPath(Left(strPathInfo, MyPos))
	if Right(strPhysicalPath, 1) <> "\" Then
'		strPhysicalPath = strPhysicalPath & "\UpLoad\" & "COMPANYCODE001" & "\"	     '☜: 실제 저장경로 (UpLoad 폴더밑에 회사코드 폴더)
		strPhysicalPath = strPhysicalPath & "\UpLoad\" & gDBServer & "_" & gDatabase & "\" & gCompany & "\"				 '☜: Global값이 아직 Setting되지 않은 관계로 
	else
'		strPhysicalPath = strPhysicalPath & "UpLoad\" & "COMPANYCODE001" & "\"
		strPhysicalPath = strPhysicalPath & "UpLoad\" & gDBServer & "_" & gDatabase & "\" & gCompany & "\"
	End if		
	
	GetPhysicalUpLoadPath = strPhysicalPath
	
End Function

'==============================================================================
' File UpLoad의 상대 경로를 얻는다.
'==============================================================================
Function GetLogicalUpLoadPath()

	strPathInfo = Request.ServerVariables("PATH_INFO")
	MyPos = Instr(2, strPathInfo, "/", 1)   
	strPathInfo = Left(strPathInfo, MyPos)
	
	if Right(strPathInfo, 1) <> "/" Then
'		strPathInfo = strPathInfo & "/UpLoad/" & "COMPANYCODE001" & "/"	     '☜: 실제 저장경로 (UpLoad 폴더밑에 회사코드 폴더)
		strPathInfo = strPathInfo & "/UpLoad/" & gDBServer & "_" & gDatabase & "/" & gCompany & "/"				 '☜: Global값이 아직 Setting되지 않은 관계로 
	else
'		strPathInfo = strPathInfo & "UpLoad/" & "COMPANYCODE001" & "/"
		strPathInfo = strPathInfo & "UpLoad/" & gDBServer & "_" & gDatabase & "/" & gCompany & "/"
	End if		

	GetLogicalUpLoadPath = strPathInfo
	
End Function

'==============================================================================
' 현재 화일의 확장명을 얻는다.
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
' ComProxy Error시 현재 UpLoad된 File 삭제 및 변수 초기화 
'==============================================================================
Sub DeleteFile(ByVal FileName)
	strRet = objRequest.DeleteFile(FileName) 
    Set objRequest = Nothing											'☜: File Upload Component Unload
End Sub

'==============================================================================
' 파일이름으로 사용될 수 없는 특수 문자를 "-"로 변경한다.
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
