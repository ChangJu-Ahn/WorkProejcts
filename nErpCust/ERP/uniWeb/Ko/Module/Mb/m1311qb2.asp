<%'======================================================================================================
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m1311qb2
'*  4. Program Name         : ����PL��ȸ 
'*  5. Program Desc         :
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 2003-06-02
'*  8. Modifier (First)     : MHJ
'*  9. Modifier (Last)      : Kim Jin Ha
'* 10. Comment              :
'* 11. Common Coding Guide  :     
'=======================================================================================================
Option Explicit
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                          '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

call LoadBasisGlobalInf()
call LoadInfTB19029B("Q", "M","NOCOOKIE","QB") 

On Error Resume Next

Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'�� : DBAgent Parameter ���� 
Dim rs1, rs2, rs3, rs4, rs5, rs6							'�� : DBAgent Parameter ���� 
Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim iTotstrData
Dim lgStrPrevKey                                            '�� : ���� �� 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT

Dim lgDataExist
Dim lgPageNo
	
Dim ICount  		                                        '   Count for column index


     Call HideStatusWnd 
     lgPageNo         = UNICInt(Trim(Request("lgPageNo1")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
     lgSelectList     = Request("lgSelectList")
     lgTailList       = Request("lgTailList")
     lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 

     Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
     call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
    
    Dim strPLNoFrom
	Dim strPLNo
                                                                      '    parameter�� ���� ���� ������ 
	UNISqlId(0) = "M1311QA102"
	UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 
	     
	If Len(Trim(Request("txtPLNo"))) Then
	   strPLNo	= " " & FilterVar(UCase(Request("txtPLNo")), "''", "S") & " "
	   strPLNoFrom = strPLNo
	Else
	   strPLNo	= "''"
	   strPLNoFrom = "" & FilterVar("%%", "''", "S") & ""
	End If
	    
	UNIValue(0,1)  = UCase(Trim(strPLNoFrom))			'---���ֹ�ȣ        
	     
	UNIValue(0,UBound(UNIValue,2) ) = " " & Trim(lgTailList)	'---Order By ���� 

	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
     
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)			
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    Dim FalsechkFlg
    
    FalsechkFlg = False 
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		If Request("Query_Msg_Flg") = "T" then
			'Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
         	Call HTMLFocus("Parent.Frm1.vspdData",I_MKSCRIPT)
		End if
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100            

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
    Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
			if ColCnt = 3 or ColCnt = 5 then
				iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(ColCnt),4,0)	
			else
				iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			end if
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount) = lgstrData	
           lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")

    If iLoopCount < C_SHEETMAXROWS_D Then										'��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close																	'��: Close recordset object
    Set rs0 = Nothing															'��: Release ADF
End Sub
%>

<Script Language=vbscript>
    With Parent
         .ggoSpread.Source  = .frm1.vspdData2
         .frm1.vspdData2.Redraw = False
         .ggoSpread.SSShowData "<%=iTotstrData%>"									'�� : Display data
         .lgPageNo1			=  "<%=lgPageNo%>"									'�� : Next next data tag
  		 .DbQueryOk(2)
         .frm1.vspdData2.Redraw = True
	End with
</Script>	

<%
    Response.End																'��: �����Ͻ� ���� ó���� ������ 
%>
