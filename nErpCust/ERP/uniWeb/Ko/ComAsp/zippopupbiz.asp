<!-- #Include file="../inc/IncSvrMain.asp" -->
<!-- #Include file="../inc/lgSvrVariables.inc" -->
<!-- #Include file="../inc/adovbs.inc" -->
<!-- #Include file="../inc/incServeradodb.asp" -->
<%																			'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim LngRow
Dim GroupCount
Dim StrNextKey		' 다음 값 

On Error Resume Next

Call LoadBasisGlobalInf()

Call HideStatusWnd

strMode      = Request("txtMode")												'☜ : 현재 상태를 받음 
lgMaxCount   = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData

If Request("txtMode") <> "" Then
'***********************************************************************************************************
	Dim PB2G151		
    Dim iStrData
    Dim exportData
    Dim exportData1
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrNo
    Dim importArray
    Dim iIntLoopCount
	Dim iCountry, iZipCd, iAddress
	
	Dim arrTemp
	
	Const C_MaxFetchRc = 0
    Const C_NextKey    = 1
    Const C_nextKey1   = 2
	Const C_Country    = 3
	Const C_ZipCd      = 4
	Const C_Address    = 5 	 	
        
	iStrPrevKey		= Trim(Request("lgStrPrevKey"))         '☜: Next Key Value	

	iStrNo          = Trim(Request("txtSerNo"))              '☜: Next Key Value	 
	iCountry        = Trim(Request("txtCountry"))
	iZipCd          = Trim(Request("txtCode"))
	iAddress        = Trim(Request("txtName"))
	         
    ReDim importArray(5)
     
    importArray(C_MaxFetchRc)	= lgMaxCount        
	importArray(C_NextKey)		= iStrPrevKey
	
	importArray(C_NextKey1)		= iStrNo
    importArray(C_Country)		= iCountry
	importArray(C_ZipCd)        = iZipCd
	importArray(C_Address)      = iAddress 
	  
	ReDim exportData(1)
	
	On Error Resume Next 
    Set PB2G151 = Server.CreateObject("PB2G151.cBListZipCode")	
	If CheckSYSTEMError(Err,True) = True Then
        set PB2G151 = nothing
        Response.End  
    End If	
	on error goto 0    
	 
    On Error Resume Next 
    Call PB2G151.B_LIST_ZIP_CODE(gStrGlobalCollection,importArray, exportData, exportData1)		
	If CheckSYSTEMError(Err,True) = True Then
        set PB2G151 = nothing
        Response.End  
    End If	
	on error goto 0   
	    
    Set PB2G151 = nothing    
	
    iStrData = ""
    iIntLoopCount = 0	

	For iLngRow = 0 To UBound(exportData1, 1) 		
	
		iIntLoopCount = iIntLoopCount + 1
		
   	    If  iIntLoopCount < (lgMaxCount + 1) Then
             iStrData = iStrData & Chr(11) & Trim(exportData1(iLngRow, 0))          'zip
             iStrData = iStrData & Chr(11) & Trim(exportData1(iLngRow, 2))          'addr
             iStrData = iStrData & Chr(11) & "" & Trim(exportData1(iLngRow, 1))     'serno
             iStrData = iStrData & Chr(11) & Trim(exportData1(iLngRow, 3))          'addr1
             iStrData = iStrData & Chr(11) & Trim(exportData1(iLngRow, 4))          'addr2
             iStrData = iStrData & Chr(11) & iIntLoopCount
			 iStrData = iStrData & Chr(11) & Chr(12)
	    Else
			iStrPrevKey = exportData1(UBound(exportData1, 1), 0)
			iStrNo      = exportData1(UBound(exportData1, 1), 1)
			Exit For
		
		End If

	Next

	If  iIntLoopCount < (lgMaxCount + 1) Then
		iStrPrevKey = ""
	End If

%>		    
<Script Language="vbscript">  
	With parent		
		.lgSerNo = "<%=Trim(iStrNo)%>"
		.lgCode = "<%=Trim(iZipCd)%>"
		.lgName = "<%=Trim(iAddress)%>"
		.lgStrPrevKey = "<%=Trim(iStrPrevKey)%>"
		.lgIntFlgMode = .PopupParent.OPMD_UMODE
	
	    .ggoSpread.Source = parent.vspdData
		.ggoSpread.SSShowData "<%=Trim(iStrData)%>"

		.vspdData.focus

		If .vspdData.MaxRows = 0 Then
			parent.UNIMsgBox "There's no data.", 48, parent.top.document.title
		End If

	End With

</Script>
<%
End If
%>