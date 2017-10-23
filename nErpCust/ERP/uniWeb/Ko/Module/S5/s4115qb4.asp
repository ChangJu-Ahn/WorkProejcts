<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                          '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
    Call loadInfTB19029B("Q", "S","NOCOOKIE","QB")
    Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")
    Call LoadBasisGlobalInf()

    On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8, rs9  '�� : DBAgent Parameter ���� 
    Dim lgstrData															'�� : data for spreadsheet data
    Dim lgTailList                                                          '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgStrColorFlag
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
 
    Dim lgFromDt		'��ȸ�Ⱓ���� 
    Dim lgToDt			'��ȸ�Ⱓ��    
    Dim lgConfFlag		'Ȯ������ 
    Dim lgSoldToParty		'����� 
    Dim lgDnType		'����ä������ 
    
    lgFromDt	= Trim(Request("txtHConFromDt"))    
    lgToDt		= Trim(Request("txtHConToDt"))   
    lgConfFlag	= Trim(Request("txtHConRdoConfFlag"))
    lgSoldToParty	= Trim(Request("txtHConSoldToParty"))
    lgDnType	= Trim(Request("txtHConDnType"))
    
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd

    lgPageNo       = UNICInt(Trim(Request("txtHlgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("txtHlgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("txtHlgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("txtHlgTailList")                                 '�� : Orderby value

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
 
        If rs0(0) > 0 Then	'����Row ���� üũ 
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

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Dim iStrVal    
    
    Redim UNISqlId(9)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(9,2)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 
               
    iStrVal = ""					
	
	'��ȸ�Ⱓ����=========================================================================================
	If Len(lgFromDt) Then
		If lgConfFlag = "Y" Then
			iStrVal = iStrVal & " DH.ACTUAL_GI_DT >=  " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""		
		Else
			iStrVal = iStrVal & " DH.PROMISE_DT >=  " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""		
		End If
	End If		
	
	'��ȸ�Ⱓ��===========================================================================================
	If Len(lgToDt) Then
		If lgConfFlag = "Y" Then
			iStrVal = iStrVal & " AND DH.ACTUAL_GI_DT <=  " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""		
		Else 
			iStrVal = iStrVal & " AND DH.PROMISE_DT <=  " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""		
		End If
	End If
	
	'�����=====================================================================================	
	If lgConfFlag = "Y" Then
		iStrVal = iStrVal & " AND DH.POST_FLAG =" & FilterVar("Y", "''", "S") & " "
	ElseIf lgConfFlag = "N" Then
		iStrVal = iStrVal & " AND DH.POST_FLAG =" & FilterVar("N", "''", "S") & " "
	End If
	
	'��ǰó��=============================================================================================    	
	If Len(lgSoldToParty) Then
		UNISqlId(1)		= "s0000qa002"	
		UNIValue(1,0)	= FilterVar(lgSoldToParty, "''", "S")
		iStrVal = iStrVal & " AND DH.SHIP_TO_PARTY =  " & FilterVar(lgSoldToParty , "''", "S") & ""				
	End If
	
	'�������¸�===========================================================================================
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
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '��: set ADO read mode
 
End Sub


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg                                                     '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                           '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8, rs9)
    
    Set lgADF = Nothing													'��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'��:Write the Script Tag "<Script language=vbscript>"
	
	'��ǰó ���翩�� 
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
	
	'�������� ���翩�� 
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

' �����Ͱ� �������� �ʴ� ��� ó�� Script �ۼ�(��ȸ���� ����)
Sub ConNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1." & pvStrField & ".alt, ""X"") " & VbCr
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' ��ȸ���ǿ� �ش��ϴ� ���� Display�ϴ� Script �ۼ� 
Sub WriteConDesc(ByVal pvStrField, Byval pvStrFieldDesc)
	Response.Write " Parent.frm1." & pvStrField & ".value = """ & ConvSPChars(pvStrFieldDesc) & """" &VbCr
End Sub

' �����Ͱ� �������� �ʴ� ��� ó�� Script �ۼ� 
Sub DataNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""900014"", ""X"", ""X"", ""X"") " & VbCr
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' ��ȸ ����� Display�ϴ� Script �ۼ� 
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


