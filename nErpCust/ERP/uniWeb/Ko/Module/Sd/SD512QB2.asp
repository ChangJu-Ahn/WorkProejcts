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

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3			'�� : DBAgent Parameter ���� 
    Dim lgstrData															'�� : data for spreadsheet data
    Dim lgTailList                                                          '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgStrColorFlag
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

    Dim iPrevEndRow
    
    Dim lgYMFromDt			'��ȸ�Ⱓ(��,��)���� 
    Dim lgYMToDt			'��ȸ�Ⱓ(��,��)�� 
    Dim lgSalesGrpCd		'�����׷� 
    Dim lgSoldToParty		'�ֹ�ó 
    Dim lgPostFlag		'����ä��Ȯ������ 
    
    lgYMFromDt		= Trim(Request("txtHConYMFromDt"))
    lgYMToDt		= Trim(Request("txtHConYMToDt"))
    lgSalesGrpCd	= Trim(Request("txtHConSalesGrpCd"))
    lgSoldToParty	= Trim(Request("txtHConSoldToPartyCd"))
    lgPostFlag		= Trim(Request("rdoHConPostFlag"))
    lgExceptFlag	= Trim(Request("rdoHConExceptFlag"))
    lgMonthDiff		= CInt(Trim(Request("txtHMonthDiff")))

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd

    lgPageNo       = UNICInt(Trim(Request("txtHlgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("txtHlgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("txtHlgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("txtHlgTailList")                                 '�� : Orderby value
    iPrevEndRow	   = 0

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

    lgstrData   = ""
    iLoopCount	= 0
    lgStrColorFlag = ""    
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
		
		If rs0(0) > 0 Then
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
    
    Redim UNISqlId(3)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(3,7)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 
               
	
	'��ȸ�Ⱓ(��,��)����==================================================================================
	If Len(lgYMFromDt) Then
		UNIValue(0,1) = " " & FilterVar(lgYMFromDt, "''", "S") & ""                           
	End If		
	
	'��ȸ�Ⱓ(��,��)��====================================================================================
	If Len(lgYMToDt) Then
		UNIValue(0,2) = " " & FilterVar(lgYMToDt, "''", "S") & ""                           	
	End If

	
	'�����׷��===========================================================================================	
    If Len(lgSalesGrpCd) Then
		UNISqlId(1)		= "s0000qa005"	
		UNIValue(1,0)	= FilterVar(lgSalesGrpCd, "''", "S")
		UNIValue(0,3)	= " " & FilterVar(lgSalesGrpCd, "''", "S") & ""
	End If		
    
    '�ֹ�ó��=============================================================================================
    If Len(lgSoldToParty) Then		
		UNISqlId(2)		= "s0000qa002"	
		UNIValue(2,0)	= FilterVar(lgSoldToParty, "''", "S")
		UNIValue(0,5)	= " " & FilterVar(lgSoldToParty, "''", "S") & ""
    Else
		UNIValue(0,5)	= "NULL"
	End If	
	
	'����ä��Ȯ������=====================================================================================	
	If Len(lgPostFlag) Then
		UNIValue(0,4)	=  " " & FilterVar(lgPostFlag, "''", "S") & ""
    Else
		UNIValue(0,4)	= "NULL"
	End If	
	
	'���ܿ���=====================================================================================	
	If Len(lgExceptFlag) Then
		UNIValue(0,6)	=  " " & FilterVar(lgExceptFlag, "''", "S") & ""
    Else
		UNIValue(0,6)	= "NULL"
	End If	
	
	UNISqlId(0) = "SD512QA201"					
    UNIValue(0,0) = Replace(lgSelectList,"?",lgMonthDiff)

    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = Replace(UCase(Trim(lgTailList)),"?",lgMonthDiff)
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF = Nothing													'��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'��:Write the Script Tag "<Script language=vbscript>"
	
	'�����׷� ���翩�� 
	If lgSalesGrpCd <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call ConNotFound("txtConSalesGrpCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesGrpNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesGrpNm", "")
	End If
	
	'�ֹ�ó ���翩�� 
	If lgSoldToParty <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call ConNotFound("txtConSoldToPartyCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSoldToPartyNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtConSoldToPartyNm", "")
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
	Response.Write " Call parent.SetFocusToDocument(""M"") " & vbCr	
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' ��ȸ ����� Display�ϴ� Script �ۼ� 
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


