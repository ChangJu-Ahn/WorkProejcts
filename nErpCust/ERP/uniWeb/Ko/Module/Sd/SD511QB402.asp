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
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
	Dim lgConDt
	Dim lgBpCd

	lgConDt = Left(Trim(Request("ConDt")),7)
	lgBpCd = Trim(Request("BpCd"))
    
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd

    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

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
    
    Const C_SHEETMAXROWS_D = 20     

    lgstrData      = ""
    
    iLoopCount = 0
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
		
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
    
    Redim UNISqlId(0)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(0,2)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 
               
    iStrVal = ""	
	
	UNIValue(0,1) = " " & FilterVar(lgConDt, "''", "S") & ""
	UNIValue(0,2) = " " & FilterVar(lgBpCd, "''", "S") & ""
	
	UNISqlId(0) = "SD511QA402"					
    UNIValue(0,0) = lgSelectList                                      
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
'    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing													'��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'��:Write the Script Tag "<Script language=vbscript>"
	
    If  rs0.EOF And rs0.BOF Then	
        rs0.Close
        Set rs0 = Nothing
        Call DataNotFound("txtConDt")	
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
	Response.Write " Parent.ggoSpread.Source  = Parent.frm1.vspdData1 " & vbCr
	Response.Write " Parent.frm1.vspdData1.Redraw = False " & vbCr      	
	Response.Write " Parent.ggoSpread.SSShowData  """ & lgstrData & """ ,""F""" & vbCr
'	Response.Write " Parent.DbQueryOk2 " & vbCr		
 	Response.Write " Parent.frm1.vspdData1.Redraw = True " & vbCr      
	Call EndScriptTag()
End Sub

%>


