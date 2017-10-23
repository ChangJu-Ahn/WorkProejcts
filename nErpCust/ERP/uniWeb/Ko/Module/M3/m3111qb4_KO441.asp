<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111pb4
'*  4. Program Name         : ���ֹ�ȣ��������ȸ 
'*  5. Program Desc         : ���ֹ�ȣ��������ȸ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/06/29
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                          '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'�� : DBAgent Parameter ���� 
Dim rs1, rs2, rs3, rs4, rs5, rs6							'�� : DBAgent Parameter ���� 
Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgStrPrevKey                                            '�� : ���� �� 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT

'--------------- ������ coding part(��������,Start)----------------------------------------------------
Dim ICount  		                                        '   Count for column index
Dim strPurGrpCd												'	���ű׷� 
Dim strPurGrpCdFrom 										
Dim strBpCd                                               '   ����ó 
Dim strBpCdFrom
Dim strPoFrDt                                               '   ������ 
Dim strPoToDt
Dim strItemCd                                               '   ǰ�� 
Dim strItemCdFrom
Dim strPoNo                                              '   Po No
Dim strPoNoFrom
Dim strPoType                                               '   �������� 
Dim strPoTypeFrom	
Dim strCfmFlg								                '   �ܰ����� 
Dim strCfmFlgFrom	
Dim arrRsVal(11)											'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array
Dim lgPageNo
Dim lgDataExist
Dim iFrPoint
iFrPoint=0
'--------------- ������ coding part(��������,End)------------------------------------------------------

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "QB")
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
       iFrPoint     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
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

	lgstrData  = Join(PvArr, "")

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

    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(1,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
     UNISqlId(0) = "m3111qa302"
								
'--------------- ������ coding part(�������,End)------------------------------------------------------

     UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

'--------------- ������ coding part(�������,Start)----------------------------------------------------

     UNIValue(0,1)  = UCase(Trim(strPoNoFrom))			'---���ֹ�ȣ    
     
'--------------- ������ coding part(�������,End)----------------------------------------------------
     
     UNIValue(0,UBound(UNIValue,2)) = Trim(lgTailList)	'---Order By ���� 

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
    
    '============================= �߰��� �κ� =====================================================================
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		if Request("Query_Msg_Flg") = "T" then
			'Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
         	Call HTMLFocus("Parent.Frm1.vspdData",I_MKSCRIPT)

		end if
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

'--------------- ������ coding part(�������,Start)----------------------------------------------------
     
    '---Po No
    If Len(Trim(Request("txtPoNo"))) Then
    	strPoNo	= " " & FilterVar(Trim(UCase(Request("txtPoNo"))), " " , "S") & " "
    	strPoNoFrom = strPoNo
    Else
    	strPoNo	= "''"
    	strPoNoFrom = "" & FilterVar("%%", "''", "S") & ""
    End If

'--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>
    
    With Parent
         .ggoSpread.Source  = .frm1.vspdData2
         Parent.frm1.vspdData2.Redraw = True
         .ggoSpread.SSShowData "<%=lgstrData%>", "F"                  '�� : Display data
                  
         Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData2,"<%=iFrPoint+1%>",parent.frm1.vspdData2.maxrows,Parent.GetKeyPos("B",17), Parent.GetKeyPos("B",8),"C", "Q" ,"X","X")	'���ֱݾ� 
         Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData2,"<%=iFrPoint+1%>",parent.frm1.vspdData2.maxrows,Parent.GetKeyPos("B",17), Parent.GetKeyPos("B",16),"A", "Q" ,"X","X")	'���ֱݾ� 
         Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData2,"<%=iFrPoint+1%>",parent.frm1.vspdData2.maxrows,Parent.GetKeyPos("B",17), Parent.GetKeyPos("B",18),"D", "Q" ,"X","X")	'ȯ�� 
         Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData2,"<%=iFrPoint+1%>",parent.frm1.vspdData2.maxrows, Parent.Parent.gCurrency , Parent.GetKeyPos("B",19),"A", "Q" ,"X","X")	'�����ڱ��ݾ�				'�����ڱ��ݾ� 
         
         .lgPageNo2			=  "<%=lgPageNo%>"               '�� : Next next data tag
  		 
         .DbQueryOk(2)
         Parent.frm1.vspdData2.Redraw = True
	End with
</Script>	

<%
    Response.End												'��: �����Ͻ� ���� ó���� ������ 
%>
