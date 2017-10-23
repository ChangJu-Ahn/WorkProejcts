<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                          '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
    Call loadInfTB19029B("Q", "S","NOCOOKIE","QB")
    Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")
    Call LoadBasisGlobalInf()

    'On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3,rs4	'☜ : DBAgent Parameter 선언 
    Dim lgstrData													'☜ : data for spreadsheet data
    Dim lgTailList													'☜ : Orderby절에 사용될 field 리스트 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgStrColorFlag
    
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
   
    Dim lgYMFromDt			'조회기간(년,월)시작 
    Dim lgYMToDt			'조회기간(년,월)끝 

    Dim lgBizAreaCd			'사업장 
    Dim lgSalesGrpCd		'영업그룹 
    Dim lgSalesTypeCd		'판매유형    
    Dim lgBillConfFlag
    Dim lgExceptFlag
    
    lgYMFromDt		= Trim(Request("txtHConYMFromDt"))
    lgYMToDt		= Trim(Request("txtHConYMToDt"))
    lgBizAreaCd		= Trim(Request("txtHConBizAreaCd"))
    lgSalesGrpCd	= Trim(Request("txtHConSalesGrpCd"))
    lgSalesTypeCd	= Trim(Request("txtHConSalesTypeCd"))      
    lgBillConfFlag	= Trim(Request("txtHConBillConfFlag"))
    lgExceptFlag	= Trim(Request("txtHExceptFlag"))
        
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd

    lgSelectList   = Request("txtHlgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("txtHlgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("txtHlgTailList")                                 '☜ : Orderby value

    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    Dim  iRsCnt
    Dim  istrLastRow
    Const C_SHEETMAXROWS_D = 20     

    lgstrData      = ""

    iLoopCount = 0
    lgStrColorFlag = ""    
    
    iRsCnt = rs0.recordcount
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""

        If iRsCnt=iLoopCount Then 
			For ColCnt = 0 To UBound(lgSelectListDT) - 1
			    If ColCnt=3 Then
				    istrLastRow = istrLastRow & Chr(11) & ""
			    Else
			        istrLastRow = istrLastRow & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			    End If 
			Next

			Exit Do
        Else
			For ColCnt = 0 To UBound(lgSelectListDT) - 1
			    If rs0(ColCnt) = "0" and ColCnt=3 Then
				    iRowStr = iRowStr & Chr(11) & "소계"
			    Else
			        iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			    End If 
			Next
		End If
		'COLOR
 		If rs0(0) > 0 Then
			lgStrColorFlag = lgStrColorFlag & CStr(iLoopCount) & gColSep & rs0(0) & gRowSep
		End If
 
        lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)

        rs0.MoveNext
	Loop

	rs0.Close
    Set rs0 = Nothing 
    
    Do while Not (rs4.EOF Or rs4.BOF)
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs4(ColCnt))
		Next
		'COLOR
 		If rs4(0) > 0 Then
			lgStrColorFlag = lgStrColorFlag & CStr(iLoopCount) & gColSep & rs4(0) & gRowSep
		End If
 
        lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)

        iLoopCount =  iLoopCount + 1
        rs4.MoveNext
	Loop
	rs4.Close
    Set rs4 = Nothing 
    
    lgstrData = lgstrData & istrLastRow & Chr(11) & Chr(12)
	lgStrColorFlag = lgStrColorFlag & CStr(iLoopCount) & gColSep & "2" & gRowSep
    
End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
   
    Redim UNISqlId(4)                                       '☜: SQL ID 저장을 위한 영역확보    

    Redim UNIValue(4,3)                                     '⊙: DB-Agent로 전송될 parameter를 위한 변수 
               
	'사업장명=============================================================================================    	
	If Len(lgBizAreaCd) Then
		UNISqlId(1)		= "s0000qa013"	
		UNIValue(1,0)	= FilterVar(lgBizAreaCd , "''", "S") 
	
		iStrVal = iStrVal & " AND BH.BIZ_AREA =  " & FilterVar(lgBizAreaCd , "''", "S") & ""				
	End If

	'영업그룹명===========================================================================================	
    If Len(lgSalesGrpCd) Then
		UNISqlId(2)		= "s0000qa005"	
		UNIValue(2,0)	= FilterVar(lgSalesGrpCd , "''", "S") 

		iStrVal = iStrVal & " AND BH.SALES_GRP =  " & FilterVar(lgSalesGrpCd , "''", "S") & ""				
	End If

	'판매유형명===========================================================================================	
    If Len(lgSalesTypeCd) Then		
		UNISqlId(3)		= "s0000qa000"	
		UNIValue(3,0)	= FilterVar("S0001" , "''", "S") 
		UNIValue(3,1)	= FilterVar(lgSalesTypeCd , "''", "S") 

		iStrVal = iStrVal & " AND BH.DEAL_TYPE =  " & FilterVar(lgSalesTypeCd , "''", "S") & ""				
	End If

	'확정여부=============================================================================================    	
	If Len(lgBillConfFlag) Then
		iStrVal = iStrVal & " AND BH.POST_FLAG =  " & FilterVar(lgBillConfFlag , "''", "S") & ""				
	End If
	
	'예외여부=============================================================================================    	
	If Len(lgExceptFlag) Then
		iStrVal = iStrVal & " AND BH.EXCEPT_FLAG =  " & FilterVar(lgExceptFlag , "''", "S") & ""				
	End If

	UNISqlId(0) = "SD511QA201"					
    UNIValue(0,0) = lgSelectList
	If Len(lgYMFromDt) Then
		UNIValue(0,1) = " " & FilterVar(lgYMFromDt, "''", "S") & ""          
	Else	                 
		UNIValue(0,1) = "" & FilterVar("1900-01-01", "''", "S") & ""          
	End If		
	If Len(lgYMToDt) Then
		UNIValue(0,2) = " " & FilterVar(lgYMToDt, "''", "S") & ""                           	
	Else	                 
		UNIValue(0,2) = "" & FilterVar("2999-12-31", "''", "S") & ""                           	
	End If
	UNIValue(0,3) = iStrVal

	UNISqlId(4) = "SD511QA202"					
	UNIValue(4,0) = " " & FilterVar(lgYMFromDt, "''", "S") & ""          
	UNIValue(4,1) = " " & FilterVar(lgYMToDt, "''", "S") & ""                           	
	UNIValue(4,2) = iStrVal
        
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '☜: set ADO read mode
 
End Sub


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    'on error resume next
    Dim lgstrRetMsg                                                     '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                           '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
    Set lgADF = Nothing													'☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'☜:Write the Script Tag "<Script language=vbscript>"

	'사업장 존재여부 
	If lgBizAreaCd <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call ConNotFound("txtConBizAreaCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBizAreaNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConBizAreaNm", "")		
	End If
	
	'영업그룹 존재여부 
	If lgSalesGrpCd <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call ConNotFound("txtConSalesGrpCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesGrpNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesGrpNm", "")
	End If

	'판매유형 존재여부 
	If lgSalesTypeCd <> "" Then
		If rs3.EOF And rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing			
			Call ConNotFound("txtConSalesTypeCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesTypeNm", rs3(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesTypeNm", "")
	End If

    If  rs0.EOF And rs0.BOF Then	
        rs0.Close
        Set rs0 = Nothing
        Call DataNotFound("txtConYMFromDt")	
        Exit Sub
    Else    
        Call MakeSpreadSheetData()
        Call WriteResult()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Write the Result
'----------------------------------------------------------------------------------------------------------
Sub BeginScriptTag()
	Response.Write "<Script language=VBScript> " & VbCr
End Sub

Sub EndScriptTag()
	Response.Write "</Script> " & VbCr
End Sub

' 데이터가 존재하지 않는 경우 처리 Script 작성(조회조건 포함)
Sub ConNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1." & pvStrField & ".alt, ""X"") " & VbCr
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' 조회조건에 해당하는 명을 Display하는 Script 작성 
Sub WriteConDesc(ByVal pvStrField, Byval pvStrFieldDesc)
	Response.Write " Parent.frm1." & pvStrField & ".value = """ & ConvSPChars(pvStrFieldDesc) & """" &VbCr
End Sub

' 데이터가 존재하지 않는 경우 처리 Script 작성 
Sub DataNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""900014"", ""X"", ""X"", ""X"") " & VbCr
	Response.Write " Call parent.SetFocusToDocument(""M"") " & vbCr	
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' 조회 결과를 Display하는 Script 작성 
Sub WriteResult()
	Response.Write " Parent.ggoSpread.Source  = Parent.frm1.vspdData " & vbCr
	Response.Write " Parent.frm1.vspdData.Redraw = False " & vbCr      	
	Response.Write " Parent.ggoSpread.SSShowData  """ & lgstrData & """ ,""F""" & vbCr
	Response.Write " Parent.lgStrColorFlag = """ & lgStrColorFlag & """" & VbCr
	Response.Write " Parent.DbQueryOk " & vbCr		
 	Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
	Call EndScriptTag()
End Sub

%>


