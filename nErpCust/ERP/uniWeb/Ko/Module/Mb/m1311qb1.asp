<%'======================================================================================================
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m1311qb1
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
Dim strSpplCdFrom
DIm strPlantCdFrom
Dim strItemCdFrom
Dim strSpplCd
DIm strPlantCd
Dim strItemCd
Dim arrRsVal(2)											'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array

     Call HideStatusWnd 
     lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
     lgSelectList     = Request("lgSelectList")
     lgTailList       = Request("lgTailList")
     lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 

     Call  TrimData()                                                     '�� : Parent�� ������ ����Ÿ ���� 
     Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
     call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query


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
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
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
	
    If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
 Sub FixUNISQLData()

    Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(3,4)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
     UNISqlId(0) = "M1311QA101"
     
     UNISqlId(1) = "M3111QA102"								              '����ó��     
     UNISqlId(2) = "M2111QA302"								              '�����     
	 UNISqlId(3) = "M2111QA303"											  'ǰ��� 
  																		  'Reusage is Recommended
     UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 
     UNIValue(0,1)  = UCase(Trim(strSpplCdFrom))			'---����ó    
	 UNIValue(0,2)  = UCase(Trim(strPlantCdFrom))		    '---���� 
     UNIValue(0,3)  = UCase(Trim(strItemCdFrom))			'---ǰ�� 
          
     UNIValue(1,0)  = UCase(Trim(strSpplCd))
     UNIValue(2,0)  = UCase(Trim(strPlantCd))
     UNIValue(3,0)  = UCase(Trim(strPlantCd))
     UNIValue(3,1)  = UCase(Trim(strItemCd))
     
     UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By ���� 

     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)			
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    Dim FalsechkFlg
    
    FalsechkFlg = False 
    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
        If Len(Request("txtSpplCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(0) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If

    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(1) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
    
    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
        If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("122700", vbInformation, "��ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       FalsechkFlg = True
	       rs0.Close
	       	Set rs0 = Nothing
			Exit Sub		
		End If
    Else    
		arrRsVal(2) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub
    
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
     '---����ó 
    If Len(Trim(Request("txtSpplCd"))) Then
    	strSpplCd	= " " & FilterVar(UCase(Request("txtSpplCd")), "''", "S") & " "
    	strSpplCdFrom = strSpplCd
    Else
    	strSpplCd	= "''"
    	strSpplCdFrom = "" & FilterVar("%%", "''", "S") & ""    	
    End If

    '---���� 
    If Len(Trim(Request("txtPlantCd"))) Then
    	strPlantCd	= " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
    	strPlantCdFrom = strPlantCd
    Else
    	strPlantCd	= "''"
    	strPlantCdFrom = "" & FilterVar("%%", "''", "S") & ""    	
    End If
    
     '---ǰ�� 
    If Len(Trim(Request("txtItemCd"))) Then
    	strItemCd	= " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
    	strItemCdFrom = strItemCd
    Else
    	strItemCd	= "''"
    	strItemCdFrom = "" & FilterVar("%%", "''", "S") & ""    	
    End If

End Sub

%>

<Script Language=vbscript>
    
    With Parent
         .ggoSpread.Source  = .frm1.vspdData
         .frm1.vspdData.Redraw = False
         .ggoSpread.SSShowData "<%=iTotstrData%>"                  '�� : Display data
         .lgPageNo			=  "<%=lgPageNo%>"               '�� : Next next data tag
  		 
  		 .frm1.hdnSpplCd.value    = "<%=ConvSPChars(Request("txtSpplCd"))%>"
         .frm1.hdnPlantCd.value   = "<%=ConvSPChars(Request("txtPlantCd"))%>"
         .frm1.hdnItemCd.value    = "<%=ConvSPChars(Request("txtItemCd"))%>"
         
         .frm1.txtSpplNm.value			=  "<%=ConvSPChars(arrRsVal(0))%>" 	
  		 .frm1.txtPlantNm.value			=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  		 .frm1.txtItemNm.value			=  "<%=ConvSPChars(arrRsVal(2))%>"
         .DbQueryOk(1)
         .frm1.vspdData.Redraw = True
	End with
</Script>	

<%
    Response.End												'��: �����Ͻ� ���� ó���� ������ 
%>
