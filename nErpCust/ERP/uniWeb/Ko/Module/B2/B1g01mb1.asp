<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Zip code)
'*  3. Program ID           : B1g01mb1.asp
'*  4. Program Name         : B1g01mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        :
'                             +B1g011ControlZipCode
'                             +B1g018ListZipCode
'                             +B16019LookupCountry
'*  7. Modified date(First) : 2000/09/14
'*  8. Modified date(Last)  : 2002/12/13
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************

%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd	
Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim lgstrdata
Dim iErrPosition


Call LoadBasisGlobalInf()

strMode      = Request("txtMode")												'☜ : 현재 상태를 받음 
lgLngMaxRow  = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
lgMaxCount   = CInt(C_SHEETMAXROWS_D)                                  '☜: Fetch count at a time for VspdData
Select Case strMode
Case CStr(UID_M0001)
	
	On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status


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
	iStrNo          = Trim(Request("lgStrNo"))              '☜: Next Key Value	 
	iCountry        = Trim(Request("txtCountry"))
	iZipCd          = Trim(Request("txtZipCd"))
	iAddress        = Trim(Request("txtAddress"))
	         
    ReDim importArray(5)
     
    importArray(C_MaxFetchRc)	= lgMaxCount        
	importArray(C_NextKey)		= iStrPrevKey
	importArray(C_NextKey1)		= iStrNo
    importArray(C_Country)		= iCountry
	importArray(C_ZipCd)        = iZipCd
	importArray(C_Address)      = iAddress 
	  
%>
<Script Language=vbscript>    
	With parent			
        .DbLookUp
	End With
</Script>	
<%      
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
			For iLngCol = 0 To UBound(exportData1, 2)
			   IF iLngCol = 1  Then				   
					iStrData = iStrData & Chr(11) & "" & Trim(exportData1(iLngRow, iLngCol))
			    ELSE
					iStrData = iStrData & Chr(11) & Trim(exportData1(iLngRow, iLngCol))					
				END IF
			Next
			 iStrData = iStrData & Chr(11) & iLngRow + 1
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

	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData					 " & vbCr 			 
    Response.Write " .ggoSpread.SSShowData """ & ConvSPChars(iStrData)			& """" & vbCr    
    Response.Write " .frm1.hCountryCd.value = """ & iCountry    & """" & vbCr
    Response.Write " .frm1.hZipCd.value = """ & iZipCd  & """" & vbCr
    Response.Write " .frm1.hAddress.value = """ & iAddress  & """" & vbCr
    Response.Write " .lgStrPrevKey        = """ & iStrPrevKey			& """" & vbCr
    Response.Write " .lgStrNo             = """ & iStrNo    			& """" & vbCr
    Response.Write " .DbQueryOk " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr


Case CStr(UID_M0002)
																
    on error resume next
    Set PB2G151 = Server.CreateObject("PB2G151.cBControlZipCode")    
    If CheckSYSTEMError(Err,True) = True Then
        set PB2G151 = nothing
        Response.End  
    End If	
	on error goto 0
    
    on error resume next
    Call PB2G151.B_CONTROL_ZIP_CODE(gStrGlobalCollection,Request("txtSpread")) 
    If CheckSYSTEMError(Err,True) = True Then
        set PB2G151 = nothing
        Response.End  
    End If	   
 	on error goto 0

    Set PB2G151 = Nothing                                                   '☜: Unload Comproxy
    
%>
<Script Language=vbscript>
	With parent																		'☜: 화면 처리 ASP 를 지칭함 
		'window.status = "저장 성공"
		.DbSaveOk
	End With
</Script>
<%					
End Select
%>




