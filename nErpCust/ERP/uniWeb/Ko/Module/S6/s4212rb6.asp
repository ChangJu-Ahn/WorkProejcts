<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4212rb6.asp																*
'*  4. Program Name         : 통관내역정보(통관현황조회에서)											*
'*  5. Program Desc         : 통관내역정보(통관현황조회에서)											*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2002/04/25                                                                *
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : RYU KYUNG RAE																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : -2000/03/21 : 화면 design                                                 *
'*                            -2002/04/22 : ADO변환                                                     *
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3           
	Dim lgStrData                                     
	Dim lgTailList
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	   
	Dim strShiptoPartyNm
	Dim strPlantNm
	Dim strSlNm
	
	On Error Resume Next
	Err.Clear
	
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "RB")   
	Call LoadBNumericFormatB("Q","S","NOCOOKIE","RB")
	Call HideStatusWnd
	Const C_SHEETMAXROWS_D  = 100   
	                                       
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         
    lgDataExist      = "No"

	Call FixUNISQLData()
	Call QueryData()
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < C_SHEETMAXROWS_D Then                                      
       lgPageNo = ""
    End If
    rs0.Close                                                       
    Set rs0 = Nothing	                                            

End Sub
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()
    On Error Resume Next
    
    If Not(rs1.EOF Or rs1.BOF) Then
       strShiptoPartyNm =  rs1(1)
    End If   

    Set rs1 = Nothing 

    If Not(rs2.EOF Or rs2.BOF) Then
       strPlantNm =  rs2(1)
    End If   

    Set rs2 = Nothing 

    If Not(rs3.EOF Or rs3.BOF) Then
       strSlNm =  rs3(1)
    End If   

    Set rs3 = Nothing 

End Sub
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Dim arrVal(2)
    Redim UNISqlId(3)                                                    '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(3,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                         '    parameter의 수에 따라 변경함 
     UNISqlId(0) = "S4212RA601" 
          

     UNIValue(0,0) = Trim(lgSelectList)		                             '☜: Select 절에서 Summary    필드 

	strVal = " "

	If Len(Request("txtCcNo")) Then
		strVal = " AND C.CC_NO =  " & FilterVar(Request("txtCCNo"), "''", "S") & " "
	Else
		strVal = ""
	End If

    UNIValue(0,1) = strVal   '---발주일 
    UNIValue(1,0) = arrVal(0)  
    UNIValue(2,0) = arrVal(1)  
    UNIValue(3,0) = arrVal(2)  
   
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             
    Dim lgADF                                                   
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
 
    If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
       rs0.Close
       Set rs0 = Nothing
       Exit Sub
    Else    
        Call  MakeSpreadSheetData()
        Call  SetConditionData()
    End If
    
End Sub

%>

<Script Language=vbscript>

	With Parent	
		
		call .CurFormatNumericOCX
		If "<%=lgDataExist%>" = "Yes" Then
       
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.HCCNO.value		= "<%=ConvSPChars(Request("txtCCNO"))%>"			
			End If
			       
			.ggoSpread.Source  = .frm1.vspdData
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'☜ : Display data
			Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,-1,-1,"<%=Trim(Request("txtCurrency"))%>",.GetKeyPos("A",1),"C","Q","X","X")
			Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,-1,-1,"<%=Trim(Request("txtCurrency"))%>",.GetKeyPos("A",2),"A","Q","X","X")
			.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
			.DbQueryOk
			.frm1.vspdData.Redraw = True       
	
		End If
	
	End With
    
</Script>

<%
	Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
