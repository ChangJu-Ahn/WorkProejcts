<!--'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Template (Save)
'*  3. Program ID           : i2214mb1
'*  4. Program Name         : 결품조회 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2000/11/01
'*  8. Modifier (First)     : KimNamHoon
'*  9. Modifier (Last)      : Lee Seung Wook
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","MB")   'ggQty.DecPoint Setting...
Call HideStatusWnd 

On Error Resume Next
Err.Clear 

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim strData                                                                '☜ : data for spreadsheet data
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strPlantCd	                                                           '⊙ : 공장코드 
Dim strBaseDate	                                                           '⊙ : 기준일자 
Dim strItemAccnt                                                           '⊙ : 품목계정 
Dim strItemGroup                                                           '⊙ : 품목그룹 

Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 

'Header의 Name부분에 대한 변수 
Dim strPlantNm
Dim strItemAccntNm
Dim strItemGroupNm

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
    Call TrimData()
    Call HeaderData()
    Call FixUNISQLData()
    Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()


    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    strPlantCd    = Trim(Request("txtPlantCd"))              '공장 
    strBaseDate   = UniConvDate(Request("txtBaseDate"))      '기준일 
    strItemAccnt  = Trim(Request("txtItemAccnt"))            '품목계정     
    strItemGroup  = Trim(Request("txtItemGroup"))            '품목계정     
	
    lgStrPrevKey  = Trim(Request("lgStrPrevKey"))            '☜ : Next key flag
	lgMaxCount	  = 100	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub HeaderData()
	Dim iStr

	Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
	Redim UNIValue(0,0)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
	
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

	'공장명 
	UNISqlId(0) = "160901saa"
	UNIValue(0,0)  = FilterVar(strPlantCd, "''", "S")		'---공장 
	
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
  	iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
    		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
    		Call DisplayMsgBox("125000",vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
    		rs0.Close
    		Set rs0 = Nothing
    		Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else    
    		strPlantNm=rs0(0)
    		rs0.Close
    		Set rs0 = Nothing
    End If
	
	'품목계정명 
	UNISqlId(0) = "160904saa"
	UNIValue(0,0)  = FilterVar(strItemAccnt, "''", "S")		'---품목계정 
		
	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	iStr = Split(lgstrRetMsg,gColSep)
	    
	If iStr(0) <> "0" Then
			Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If    
	        
	If  rs0.EOF And rs0.BOF Then
			Call DisplayMsgBox("169952",vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
			rs0.Close
			Set rs0 = Nothing
			Response.End													'☜: 비지니스 로직 처리를 종료함 
	Else    
			strItemAcctNm=rs0(0)
			rs0.Close
			Set rs0 = Nothing
	End If
End Sub
    
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(0,5)

    UNISqlId(0) = "I2214ma1"
    
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    UNIValue(0,0) = FilterVar(strPlantCd, "","S")  	            '---공장 
    UNIValue(0,1) = FilterVar(strItemAccnt, "","S")             '---품목계정 
    UNIValue(0,2) = FilterVar(strBaseDate,"" & FilterVar("1900-01-01", "''", "S") & "","S")   '---기준일자 
    If  Trim(strItemGroup) = ""  Then
    UNIValue(0,3) = "''"  						                '---품목그룹 
    UNIValue(0,4) = "" & FilterVar("zzzzzzzzzz", "''", "S") & ""			                    '---품목그룹    
    Else
    UNIValue(0,3) = FilterVar(strItemGroup, "''", "S")              '---품목그룹 
    UNIValue(0,4) = FilterVar(strItemGroup, "''", "S")              '---품목그룹 
    End if
	UNIValue(0,5) = FilterVar(lgStrPrevKey, "''", "S")				'next Key (C_Item_CD)

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim iRCnt
    Dim iCnt
    Dim PvArr
	
	iCnt			= 0
	strData			= ""
	
	If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If
    
    Redim PvArr(0)		

	For iRCnt = 1 to iCnt  *  lgMaxCount                                   '☜ : Discard previous data
			rs0.MoveNext
	Next
	
	iRCnt = -1	
	
    'Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
        
    iStr = Split(lgstrRetMsg, gColSep)
    
    If iStr(0) <> "0" Then    	
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014",vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    End if        
	'=====================================================================================================================================
	lgStrPrevKey = ConvSPChars(rs0(0))

		Do while Not (rs0.EOF Or rs0.BOF)       
			iRCnt =  iRCnt + 1
			
			ReDim Preserve PvArr(iRCnt)
			
			strData =	Chr(11) & ConvSPChars(rs0(1)) & _											
						Chr(11) & ConvSPChars(rs0(2)) & _
						Chr(11) & ConvSPChars(rs0(3)) & _
						Chr(11) & ConvSPChars(rs0(4)) & _
						Chr(11) & UNIDateClientFormat(rs0(5)) & _
						Chr(11) & UniConvNumberDBToCompany(rs0(6), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
						Chr(11) & CStr((iCnt * lgMaxCount) + iRCnt) & Chr(11) & Chr(12)
			
			PvArr(iRCnt) = strData
			
			If  iRCnt >= lgMaxCount Then
				iCnt = iCnt + 1
				lgStrPrevKey = CStr(iCnt)
				Exit Do
			End If
			rs0.MoveNext	
		Loop
        '=====================================================================================================================================		
		strData = Join(PvArr, "")
		
		If  iRCnt < lgMaxCount Then                                            '☜: Check if next data exists
			lgStrPrevKey = ""                                                  '☜: 다음 데이타 없다.
		End If
		
		If strData = "" Then
			Call DisplayMsgBox("900014",vbOKOnly, "", "", I_MKSCRIPT)   'No Data Found!!
  		End If
		  
		rs0.Close
		Set rs0		= Nothing 
		Set lgADF	= Nothing                                                    '☜: ActiveX Data Factory Object Nothing

End Sub
	Response.Write "<Script Language=vbscript> "					& vbCr
	Response.Write " With Parent "									& vbCr
	
	Response.Write "	.ggoSpread.Source	= .frm1.vspdData "			& vbCr
	Response.Write "	.ggoSpread.SSShowData  """ & strData  & """"	& vbCr
	Response.Write "	.lgStrPrevKey   = """ & ConvSPChars(lgStrPrevKey) & """" & vbCr  
	
	Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then " & vbCr
	Response.Write "		.DbQuery																		"	& vbCr
	Response.Write "	Else																				"	& vbCr
	Response.Write "		.DbQueryOK																		"	& vbCr
	Response.Write "	End If																				"	& vbCr
	
	Response.Write "	.frm1.vspdData.focus																"	& vbCr
	
	Response.Write " End With																				"	& vbCr
	Response.Write "</Script>																				"	& vbCr
	
	Response.End	
	
%>

