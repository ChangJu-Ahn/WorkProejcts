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

    On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8, rs9  '☜ : DBAgent Parameter 선언 
    Dim lgstrData															'☜ : data for spreadsheet data
    Dim lgTailList                                                          '☜ : Orderby절에 사용될 field 리스트 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgStrColorFlag
    
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
 
    Dim lgFromDt		'조회기간시작 
    Dim lgToDt			'조회기간끝    
    Dim lgConfFlag		'확정여부 
    Dim lgSoldToParty		'사업장 
    Dim lgDnType		'매출채권형태 
    
    lgFromDt	= Trim(Request("txtHConFromDt"))    
    lgToDt		= Trim(Request("txtHConToDt"))   
    lgConfFlag	= Trim(Request("txtHConRdoConfFlag"))
    lgSoldToParty	= Trim(Request("txtHConSoldToParty"))
    lgDnType	= Trim(Request("txtHConDnType"))
    
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd

    lgPageNo       = UNICInt(Trim(Request("txtHlgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
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
    
    lgstrData      = ""
    
    iLoopCount = 0
    lgStrColorFlag = 0
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If rs0(0) > 0 Then	'집계Row 여부 체크 
			lgStrColorFlag = lgStrColorFlag & CStr(iLoopCount) & gColSep & rs0(0) & gRowSep
		End If
		
        lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
       
        rs0.MoveNext
	Loop
   	
	rs0.Close
    Set rs0 = Nothing 
    
End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Dim iStrVal    
    
    Redim UNISqlId(9)                                       '☜: SQL ID 저장을 위한 영역확보    

    Redim UNIValue(9,2)                                     '⊙: DB-Agent로 전송될 parameter를 위한 변수 
               
    iStrVal = ""					
	
	'조회기간시작=========================================================================================
	If Len(lgFromDt) Then
		If lgConfFlag = "Y" Then
			iStrVal = iStrVal & " DH.ACTUAL_GI_DT >=  " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""		
		Else
			iStrVal = iStrVal & " DH.PROMISE_DT >=  " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""		
		End If
	End If		
	
	'조회기간끝===========================================================================================
	If Len(lgToDt) Then
		If lgConfFlag = "Y" Then
			iStrVal = iStrVal & " AND DH.ACTUAL_GI_DT <=  " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""		
		Else 
			iStrVal = iStrVal & " AND DH.PROMISE_DT <=  " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""		
		End If
	End If
	
	'출고여부=====================================================================================	
	If lgConfFlag = "Y" Then
		iStrVal = iStrVal & " AND DH.POST_FLAG =" & FilterVar("Y", "''", "S") & " "
	ElseIf lgConfFlag = "N" Then
		iStrVal = iStrVal & " AND DH.POST_FLAG =" & FilterVar("N", "''", "S") & " "
	End If
	
	'납품처명=============================================================================================    	
	If Len(lgSoldToParty) Then
		UNISqlId(1)		= "s0000qa002"	
		UNIValue(1,0)	= FilterVar(lgSoldToParty, "''", "S")
		iStrVal = iStrVal & " AND DH.SHIP_TO_PARTY =  " & FilterVar(lgSoldToParty , "''", "S") & ""				
	End If
	
	'출하형태명===========================================================================================
    If Len(lgDnType) Then		
		UNISqlId(2)		= "s0000qa000"	
		UNIValue(2,0)	= FilterVar("I0001", "''", "S")
		UNIValue(2,1)	= FilterVar(lgDnType, "''", "S")   
		iStrVal = iStrVal & " AND DH.MOV_TYPE =  " & FilterVar(lgDnType , "''", "S") & ""				
	End If		
   
    
	UNISqlId(0)	= "S4115QA401"					
    UNIValue(0,0) = lgSelectList                                      
	UNIValue(0,1) = iStrVal	   
	UNIValue(0,2) = iStrVal      
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '☜: set ADO read mode
 
End Sub


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg                                                     '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                           '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8, rs9)
    
    Set lgADF = Nothing													'☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'☜:Write the Script Tag "<Script language=vbscript>"
	
	'납품처 존재여부 
	If lgSoldToParty <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call ConNotFound("txtConSoldToParty")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSoldToPartyNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConSoldToPartyNm", "")		
	End If
	
	'출하형태 존재여부 
	If lgDnType <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call ConNotFound("txtConDnType")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConDnTypeNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtConDnTypeNm", "")
	End If
	
	 
    If  rs0.EOF And rs0.BOF Then	
        rs0.Close
        Set rs0 = Nothing
        Call DataNotFound("txtConFromDt")	
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
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' 조회 결과를 Display하는 Script 작성 
Sub WriteResult()
	Response.Write " Parent.ggoSpread.Source  = Parent.frm1.vspdData " & vbCr
	Response.Write " Parent.frm1.vspdData.Redraw = False " & vbCr      	
	Response.Write " Parent.ggoSpread.SSShowData  """ & lgstrData & """ ,""F""" & vbCr
	Response.Write " parent.lgStrColorFlag = """ & lgStrColorFlag & """" & vbCr	
	Response.Write " Parent.DbQueryOk " & vbCr		
 	Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
	Call EndScriptTag()
End Sub

%>


