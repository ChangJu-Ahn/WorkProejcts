<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4425mb1.asp
'*  4. Program Name         : ������������ȸ 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2003-02-19
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag			'DBAgent Parameter ���� 
Dim rs0, rs1, rs2

Dim	txtMode
Dim	txtChildPlantCd
Dim	txtChildCostCd
Dim	txtChildItemCd
Dim txtCondCostCd
Dim txtCondItemAcctCd
Dim txtCondPlantCd
Dim	lgSpid
Dim lgSum1
Dim lgSum2
Dim lgSum3
Dim lgDataExist
Dim lgCodeCond,lgCodeCond1




									'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================

Call HideStatusWnd

On Error Resume Next
Err.Clear

    
    
	txtMode				= Trim(Request("txtMode"))						'�� : ���� ���¸� ���� 
	txtChildPlantCd		= Trim(Request("txtChildPlantCd"))
	txtChildCostCd		= Trim(Request("txtChildCostCd"))
	txtChildItemCd		= Trim(Request("txtChildItemCd"))
	lgSpid				= Trim(Request("txtSpid"))
	
	txtCondCostCd		= Trim(Request("txtCondCostCd"))
	txtCondItemAcctCd	= Trim(Request("txtCondItemAcctCd"))
	txtCondPlantCd		= Trim(Request("txtCondPlantCd"))	

	lgCodeCond1 = ""
	IF Trim(txtCondItemAcctCd) <> "" Then
		lgCodeCond1 =  lgCodeCond1 & " and a.child_item_acct = " & FilterVar(txtCondItemAcctCd, "''", "S")
	END IF
	
	IF Trim(txtCondPlantCd) <> "" Then
		lgCodeCond1 =  lgCodeCond1 & " and a.child_plant_cd = " & FilterVar(txtCondPlantCd, "''", "S")
	END IF

	IF Trim(txtCondCostCd) <> "" Then
		lgCodeCond1 =  lgCodeCond1 & " and a.cost_cd = " & FilterVar(txtCondCostCd, "''", "S")
	END IF

	
	lgCodeCond = ""

	lgCodeCond = lgCodeCond & " and a.child_plant_cd = " & FilterVar(txtChildPlantCd, "''", "S")
	lgCodeCond = lgCodeCond & " and a.cost_cd = " & FilterVar(txtChildCostCd, "''", "S")
	lgCodeCond = lgCodeCond & " and a.child_item_cd = " & FilterVar(txtChildItemCd, "''", "S")
	
	



	
    select case txtMode
		case CStr(UID_M0001)
			Call FixUNISQLData()
			Call QueryData()
	end select

	
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
   
    
    lgDataExist    = "Yes"
    lgStrData      = ""


    iLoopCount = 0

  
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        		 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(0))		'���� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(1))		'ǰ�� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(2))		'ǰ��� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(3))		'C/C
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(4))		'C/C�� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(5))		'ǰ����� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(6))		'ǰ������� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(7))		'��ο�� 
        lgstrData = lgstrData & Chr(11) & ConvSPChars(rs0(8))		'��α��� 
 		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(9), ggQty.DecPoint, 0) '��αٰ�Data
		lgstrData = lgstrData & Chr(11) & UNINumClientFormat(rs0(10), ggAmtofMoney.DecPoint, 0) '��αݾ� 

        lgstrData = lgstrData & Chr(11) & iLoopCount 
        lgstrData = lgstrData & Chr(11) & Chr(12)
 
        rs0.MoveNext
	Loop

  	
	rs0.Close
    Set rs0 = Nothing 
    
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(2,1)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "C3612MA03"
    UNISqlId(1) = "C3612MA04"		'���� 
    UNISqlId(2) = "C3612MA05"					
    
    
    UNIValue(0,0) = FilterVar(lgSpid, "''", "S") 
    UNIValue(0,1) = lgCodeCond
	
    UNIValue(1,0) = FilterVar(lgSpid, "''", "S") 
	UNIValue(1,1) = lgCodeCond1
	
    UNIValue(2,0) = FilterVar(lgSpid, "''", "S") 
    UNIValue(2,1) = lgCodeCond

	
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    IF NOT (rs1.EOF or rs1.BOF) then
		lgSum1 = rs1(0)				
	ELSE
		lgSum1 = 0
	End if
    rs1.Close
    Set rs1 = Nothing 
    

    IF NOT (rs2.EOF or rs2.BOF) then
		lgSum2 = rs2(0)				
		lgSum3 = rs2(1)				
	ELSE
		lgSum2 = 0
		lgSum3 = 0 
	End if
    rs2.Close
    Set rs2 = Nothing 

 	    
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else
        Call  MakeSpreadSheetData()
    End If


		
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
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '��: Protect system from crashing
    Err.Clear                                                                         '��: Clear Error status
    If CheckSYSTEMError(pErr,True) = True Then
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If

End Sub


%>

<Script Language=vbscript>
 
	With Parent
		select case "<%=txtMode%>"	
			case "<%=UID_M0001%>" 
				   .frm1.txtSum2.text	= "<%=UniNumClientFormat(lgSum1,ggAmtOfMoney.Decpoint,0)%>"
				   .frm1.txtSum3.text	= "<%=UniNumClientFormat(lgSum2,ggAmtOfMoney.Decpoint,0)%>" 
				   .frm1.txtSum4.text	= "<%=UniNumClientFormat(lgSum3,ggQty.Decpoint,0)%>" 
				If "<%=lgDataExist%>" = "Yes" AND "<%=lgErrorStatus%>" <> "YES" Then
				   .ggoSpread.Source  = Parent.frm1.vspdData2
				   .ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
					
				End If
		end select			    
    End With

</Script>	
	

<%
Set lgADF = Nothing												'��: ActiveX Data Factory Object Nothing
%>