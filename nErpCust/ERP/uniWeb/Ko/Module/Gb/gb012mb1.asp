<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->


<%

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear 
    
    Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("Q","*", "NOCOOKIE", "QB")       

    Dim lgtotSalesAmt    '������հ� 
    Dim lgtotCostAmt     '������� �Ѱ� 
    Dim lgtotPorfitAmt   '�������� �Ѱ� 
    Dim lgtotTotCostAmt  '�ѿ��� �Ѱ� 
    Dim lgtotSalesProfitAmt '���������հ� 
    Dim lgtotCurProfitAmt   '������� �Ѱ� 
    Dim lgtotNetProfitAmt   '�������� �Ѱ� 
    
    Dim txtBizUnitnm
	Dim txtCostnm
	Dim txtSalesOrgnm
	Dim txtSalesGrpnm
	Dim txtItemGroupnm
	Dim txtBpnm
'	Dim lgLngMaxRow


	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1
	
	Dim lgDataExist
	Dim lgCodeCond
	Dim lgCodeCond1
	Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	Dim lgPageNo
	Dim lgStrData 

    
    Call HideStatusWnd                                                               '��: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Filtervar(Request("txtKeyStream"),"","SNM"),gColSep)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    
    
	 lgPageNo       = Cint(Trim(Request("lgPageNo")))    
'    lgMaxCount     = Cint(Trim(Request("lgMaxCount")))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
   
       
 	Const C_SHEETMAXROWS_D  = 1000 
    
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '��: Max fetched data at a time

    
    Call FixUNISQLData()
    Call QueryData()


'============================================================================================================
Sub MakeSpreadSheetData()

    Dim  ColCnt
    Dim  iLoopCount

    On Error Resume Next                                                                 '��: Protect system from crashing
    Err.Clear                                                                            '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    lgDataExist    = "Yes"
    lgStrData      = ""


    iLoopCount = 0
       
    If lgPageNo > 0 Then
       rs0.Move     = lgMaxCount * lgPageNo                  'lgMaxCount:Max Fetched Count at once , lgPageNo : Previous PageNo
    End If  
  
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1

       
          lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(0))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(1))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(2))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(3))

          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(4),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(5),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(6),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(7),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(8),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(9),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(10),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(11),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(12),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(13),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(14),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(15),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(16),ggAmtOfMoney.DecPoint, 0)
         
   				
		lgstrData = lgstrData & Chr(11) & iLoopCount 
		lgstrData = lgstrData & Chr(11) & Chr(12)		
			
        
        If  iLoopCount >= lgMaxCount Then
            lgPageNo = lgPageNo + 1
            Exit Do
        End If        
 
        rs0.MoveNext
	Loop

	
    If  iLoopCount < lgMaxCount Then                                            '��: Check if next data exists
        lgPageNo = 0													'��: ���� ����Ÿ ����.
    End If

  	
	rs0.Close
    Set rs0 = Nothing 
	
End Sub    



'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	
   IF Trim(lgKeyStream(2)) <> "" Then                 '����� 
		Call CommonQueryRs("BIZ_UNIT_NM","B_BIZ_UNIT","BIZ_UNIT_CD = " & FilterVar(lgKeyStream(2), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If Trim(Replace(lgF0,Chr(11),"")) = "X" then
			'Call DisplayMsgBox("127900", vbInformation, "", "", I_MKSCRIPT)	
				'����� ����Ÿ�� �������� �ʽ��ϴ�.
			txtBizUnitnm = ""	
			'Exit Sub
		Else
			txtBizUnitnm = Trim(Replace(lgF0,Chr(11),""))
		End if
	END IF
	
	
	IF Trim(lgKeyStream(3)) <> "" Then                 '�ŷ�ó 
  		Call CommonQueryRs("BP_NM","B_BIZ_PARTNER","BP_CD = " & FilterVar(lgKeyStream(3), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		If Trim(Replace(lgF0,Chr(11),"")) = "X" then
			'Call DisplayMsgBox("126100", vbInformation, "", "", I_MKSCRIPT)	
				'�ŷ�ó������ �����ϴ�.
			txtBpnm = ""
			'Exit Sub	
		Else
			txtBpnm = Replace(Trim(Replace(lgF0,Chr(11),"")),"""","")
		End if
	END IF

	

    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(1,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 


	UNISqlId(0) = "Gb012MA01"      'ǰ�� ���ͺ� 
	UNISqlId(1) = "Gb012MA02"      '�Ѱ��� 
	
	
	lgCodeCond	   = ""
	lgCodeCond1	   = ""
	
    IF Trim(lgKeyStream(2)) = "" Then                   ' ������ڵ� 
		lgCodeCond = lgCodeCond & "" 
	ELSE
		lgCodeCond = lgCodeCond & " and a.biz_unit_cd = " & FilterVar(lgKeyStream(2), "''", "S")
	END IF
	

	
	IF Trim(lgKeyStream(3)) = "" Then					 ' �ŷ�ó 
		lgCodeCond =  lgCodeCond & "" 
	ELSE
		lgCodeCond =  lgCodeCond & " and a.bp_cd = " & FilterVar(lgKeyStream(3), "''", "S")
	END IF
	
	lgCodecond1 = lgCodeCond
	
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------

    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	UNIValue(0,0)  = FilterVar(lgKeyStream(0), "''", "S")
	UNIValue(0,1)  = FilterVar(lgKeyStream(1), "''", "S")
	
      
    UNIValue(0,2)  = lgCodeCond 
    
    UNIValue(1,0)  = FilterVar(lgKeyStream(0), "''", "S")
	UNIValue(1,1)  = FilterVar(lgKeyStream(1), "''", "S")
	
	UniValue(1,2)  = lgCodecond1 
	
         
    '--------------- ������ coding part(�������,End)------------------------------------------------------

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF               
	On Error Resume Next
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    
    

    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

       
    IF NOT (rs1.EOF or rs1.BOF) then
		lgtotSalesAmt = rs1(0)
		lgtotCostAmt = rs1(1)
		lgtotPorfitAmt = rs1(2)
		lgtotSalesProfitAmt = rs1(3)
		lgtotCurProfitAmt = rs1(4)
		lgtotTotCostAmt = rs1(5)
		lgtotNetProfitAmt = rs1(6)
	ELSE
		
	End if
    rs1.Close
    Set rs1 = Nothing 
  
		    
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else
        Call  MakeSpreadSheetData()
    End If


    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
		
End Sub
    

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '��: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

%>

<Script Language="VBScript">

    With Parent.Frm1
             .txtBizUnitNm.Value				= "<%=ConvSPChars(txtBizUnitNm)%>"            
             .txtBpNm.Value						= "<%=ConvSPChars(txtBpnm)%>"

  
			.totSalesAmt.text     = "<%=UNINumClientFormat(lgtotSalesAmt,ggAmtofMoney.DecPoint, 0)%>"      '����� �Ѱ� 
            .totCostAmt.text		= "<%=UNINumClientFormat(lgtotCostAmt,ggAmtofMoney.DecPoint, 0)%>"      '������� �Ѱ� 
            .totPorfitAmt.text    = "<%=UNINumClientFormat(lgtotPorfitAmt,ggAmtofMoney.DecPoint, 0)%>"      '�������� �Ѱ� 
            .totSalesProfitAmt.text		= "<%=UNINumClientFormat(lgtotSalesProfitAmt,ggAmtofMoney.DecPoint, 0)%>"			'�������� �Ѱ� 
            .totCurProfitAmt.text = "<%=UNINumClientFormat(lgtotCurProfitAmt,ggAmtofMoney.DecPoint, 0)%>"      '������� �Ѱ� 
            .totNetProfitAmt.text = "<%=UNINumClientFormat(lgtotNetProfitAmt,ggAmtofMoney.DecPoint, 0)%>"      '�������� �Ѱ� 
            .totTotCostAmt.text	= "<%=UNINumClientFormat(lgtotTotCostAmt,ggAmtofMoney.DecPoint, 0)%>"      '�ѿ��� �Ѱ� 
    End With
    
    With Parent      

  		If "<%=lgDataExist%>" = "Yes" Then

		   .ggoSpread.Source  = Parent.frm1.vspdData
		   .ggoSpread.SSShowData	"<%=lgstrData%>"                  '�� : Display data
		   .lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
		   .DbQueryOk
		End If

    End With      
    
       
</Script>	
