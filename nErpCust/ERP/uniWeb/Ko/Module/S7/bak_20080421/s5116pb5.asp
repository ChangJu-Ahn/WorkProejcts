<%'======================================================
'********************************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5116pa5
'*  4. Program Name         : 매출채권상세 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2002/05/03
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Kwak Eunkyoung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

On Error Resume Next

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1			'☜ : DBAgent Parameter 선언 
   Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
   Dim lgMaxCount                                              '☜ : Spread sheet 의 visible row 수 
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "PB")

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

    Dim lgFromDt			'조회기간시작 
    Dim lgToDt				'조회기간끝 
    Dim lgBizArea			'사업장 
	Dim lgBillTypeCd		'매출채권형태 
	Dim lgBpCd				'거래처 
    Dim lgRdoFlag			'매출채권확정여부 
   
    lgFromDt		= Trim(Request("txtHConFromDt"))    
    lgToDt			= Trim(Request("txtHConToDt"))
    lgBizArea		= Trim(Request("txtHConBizArea"))
    lgBillTypeCd	= Trim(Request("txtHConBillType"))
    lgBpCd			= Trim(Request("txtHConBpCd"))
    lgRdoFlag		= Trim(Request("txtHConRdoFlag"))
            
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

    Call HideStatusWnd 

    lgPageNo       = UNICInt(Trim(Request("txtHlgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("txtHlgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("txtHlgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("txtHlgTailList")                                 '☜ : Orderby value

    lgMaxCount       = 50							                       '☜ : 한번에 가져올수 있는 데이타 건수 

    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
    call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
 
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim iStrVal
	Dim arrVal(0)
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 

    Redim UNIValue(1,6)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 

    iStrVal = "WHERE"    				
	iStrVal = iStrVal & " BILL_DT >=  " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""		
	If Len(lgToDt) Then
		iStrVal = iStrVal & " AND BILL_DT <=  " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""			
	End If				

	'사업장=========================================================================================
	If Len(lgBizArea) Then
		UNIValue(0,2)	=  " " & FilterVar(lgBizArea, "''", "S") & ""
	Else
		UNIValue(0,2)	= "NULL"
	End If		

	'매출채권행태명=============================================================================================    	
	If Len(lgBillTypeCd) Then
		UNIValue(0,3)	=  " " & FilterVar(lgBillTypeCd, "''", "S") & ""
	Else
		UNIValue(0,3)	= "NULL"
	End If

	'거래처=========================================================================================
	If Len(lgBpCd) Then
		UNIValue(0,4)	=  " " & FilterVar(lgBpCd, "''", "S") & ""
	Else
		UNIValue(0,4)	= "NULL"
	End If		

	'확정여부===========================================================================================	
	If lgRdoFlag <> "%" Then
		UNIValue(0,5)	= " " & FilterVar(lgRdoFlag, "''", "S") & ""
	Else
		UNIValue(0,5)	= "NULL"
	End If

	UNISqlId(0) = "S5116PA501"
	UNISqlId(1) = "S5116PA501"											
    UNIValue(0,0) = Trim(lgSelectList)                                      
	UNIValue(0,1) = iStrVal	         

	UNIValue(1,0) = " SUM(ISNULL(BH.BILL_AMT_LOC,0) + ISNULL(BH.VAT_AMT_LOC,0)) AS TOTAL_AMT, SUM(ISNULL(BH.BILL_AMT_LOC,0)) AS BILL_AMT, SUM(ISNULL(BH.VAT_AMT_LOC,0)) AS VAT_AMT, SUM(ISNULL(BH.COLLECT_AMT_LOC,0)) AS COLLECT_AMT, SUM(ISNULL(BH.DEPOSIT_AMT_LOC,0)) AS DEPOSIT_AMT "
	Dim iLoop
	For iLoop = 1 To 5
		UNIValue(1,iLoop) = UNIValue(0,iLoop)	
	Next

'================================================   
   
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

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
    End If
    
End Sub

%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then
       'Show multi spreadsheet data from this line
		With parent       
			.ggoSpread.Source  = .frm1.vspdData
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'☜ : Display data
			.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag

			.frm1.txtTotalAmt.text		= "<%=rs1(0)%>"
			.frm1.txtBillAmt.text		= "<%=rs1(1)%>"
			.frm1.txtVatAmt.text		= "<%=rs1(2)%>"
			.frm1.txtCollectAmt.text	= "<%=rs1(3)%>"
			.frm1.txtDepositAmt.text	= "<%=rs1(4)%>"

 			.DbQueryOk
			.frm1.vspdData.Redraw = True        
		End with
    End If   

</Script>	
