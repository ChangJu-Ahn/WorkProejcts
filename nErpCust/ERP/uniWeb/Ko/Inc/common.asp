<%

    Session.Timeout = 60          ' minute 
    Server.ScriptTimeOut = 3600   ' NumSeconds

    Const gTitleWidth = 500
    Dim MasterConnString
    Dim MetaConnString
    Dim SourceConnString
    Dim TargetConnString

    'View=============================
    Public Function GetViewList(ByVal pConn)

        Dim iTemp 
        Dim adoRS 
        
        GetViewList = ""
    
        Set adoRS = Server.CreateObject("ADODB.Recordset")

        adoRS.Open "select name from dbo.sysobjects where xtype ='V' and name not in ('sysconstraints','syssegments')  order by name", pConn, 3  '3 means adOpenStatic
    
        If Err.Number = 0 Then
           If Not (adoRS.EOF Or adoRS.BOF) Then
              iTemp = adoRS.GetString(, , , ",")
              GetViewList = Mid(iTemp, 1, Len(iTemp) - 1)
           End If
        End If
    
        adoRS.Close

        Set adoRS = Nothing

    End Function	
	
    Public Function GetViewCount(ByVal pConn)

        Dim iTemp 
        Dim adoRS 
    
        Set adoRS = Server.CreateObject("ADODB.Recordset")

        adoRS.Open "select count(name) from dbo.sysobjects where xtype ='V' and name not in ('sysconstraints','syssegments')  ", pConn, 3  '3 means adOpenStatic

        If Err.Number = 0 Then
           GetViewCount =  adoRS(0)
        End If
    
        adoRS.Close

        Set adoRS = Nothing

    End Function	
    'View=============================
    
    'SP=============================
    Public Function GetSPList(ByVal pConn)

        Dim iTemp 
        Dim adoRS 
        
        GetSPList = ""
    
        Set adoRS = Server.CreateObject("ADODB.Recordset")

        adoRS.Open "select name from dbo.sysobjects where xtype ='P'  order by name", pConn, 3  '3 means adOpenStatic
    
        If Err.Number = 0 Then
           If Not (adoRS.EOF Or adoRS.BOF) Then
              iTemp = adoRS.GetString(, , , ",")
              GetSPList = Mid(iTemp, 1, Len(iTemp) - 1)
           End If
        End If
    
        adoRS.Close

        Set adoRS = Nothing

    End Function	
	
    Public Function GetSPCount(ByVal pConn)

        Dim iTemp 
        Dim adoRS 
    
        Set adoRS = Server.CreateObject("ADODB.Recordset")

        adoRS.Open "select count(name) from dbo.sysobjects where xtype ='P' and category != '2'  ", pConn, 3  '3 means adOpenStatic

        If Err.Number = 0 Then
           GetSPCount =  adoRS(0)
        End If
    
        adoRS.Close

        Set adoRS = Nothing

    End Function	
    'SP=============================    
    
    'UDF=============================
    Public Function GetUDFList(ByVal pConn)

        Dim iTemp 
        Dim adoRS 
        
        GetUDFList = ""
    
        Set adoRS = Server.CreateObject("ADODB.Recordset")

        adoRS.Open "select name from dbo.sysobjects where xtype IN ('TF','FN','IF') order by name", pConn, 3  '3 means adOpenStatic
    
        If Err.Number = 0 Then
           If Not (adoRS.EOF Or adoRS.BOF) Then
              iTemp = adoRS.GetString(, , , ",")
              GetUDFList = Mid(iTemp, 1, Len(iTemp) - 1)
           End If
        End If
    
        adoRS.Close

        Set adoRS = Nothing

    End Function	
	
    Public Function GetUDFCount(ByVal pConn)

        Dim iTemp 
        Dim adoRS 
    
        Set adoRS = Server.CreateObject("ADODB.Recordset")

        adoRS.Open "select count(name) from dbo.sysobjects where xtype IN ('TF','FN','IF')  ", pConn, 3  '3 means adOpenStatic

        If Err.Number = 0 Then
           GetUDFCount =  adoRS(0)
        End If
    
        adoRS.Close

        Set adoRS = Nothing

    End Function	


Function GetMetaTableList(ByVal pConn,ByVal pM,ByVal pT)

        Dim iTemp 
        Dim adoRS 
        
        GetMetaTableList = ""
    
        Set adoRS = Server.CreateObject("ADODB.Recordset")
        
        If pM = "" And pT = "" Then
           iTemp = " where kind= 'S' or xbit > 0 "
        ElseIf pM = "M" And pT = "T" Then
           iTemp = " "
        ElseIf pM = "M" Then
           iTemp = " where kind in ('S','M') or xbit > 0 "
        ElseIf pT = "T" Then
           iTemp = " where kind in ('S','T') or xbit > 0 "
        End If      

        iTemp = "select table_id from dbo.Z_TABLE_LIST " & iTemp & " order by table_id "

        adoRS.Open iTemp , pConn, 3  '3 means adOpenStatic
    
        If Err.Number = 0 Then
           If Not (adoRS.EOF Or adoRS.BOF) Then
              iTemp = adoRS.GetString(, , , ",")
              GetMetaTableList = Mid(iTemp, 1, Len(iTemp) - 1)
           End If
        End If
    
        adoRS.Close

        Set adoRS = Nothing

End Function



Function GetMetaTableCount(ByVal pConn,ByVal pM,ByVal pT)

        Dim iTemp 
        Dim adoRS 
        
        GetMetaTableCount = ""
        
        pM = Trim(pM)
        pT = Trim(pT)
    
        Set adoRS = Server.CreateObject("ADODB.Recordset")
        
        If pM = "" And pT = "" Then
           iTemp = " where kind= 'S' or xbit > 0 "
        ElseIf pM = "M" And pT = "T" Then
           iTemp = " "
        ElseIf pM = "M" Then
           iTemp = " where kind in ('S','M') or xbit > 0 "
        ElseIf pT = "T" Then
           iTemp = " where kind in ('S','T') or xbit > 0 "
        End If      
		
        adoRS.Open "select count(table_id)  from dbo.Z_TABLE_LIST  " & iTemp , pConn, 3  '3 means adOpenStatic
    
        If Err.Number = 0 Then
           GetMetaTableCount =  adoRS(0)
        End If
    
        adoRS.Close

        Set adoRS = Nothing

End Function	
	

Function GetTableList(ByVal pConn)

        Dim iTemp 
        Dim adoRS 
        
        GetTableList = ""
    
        Set adoRS = Server.CreateObject("ADODB.Recordset")

        adoRS.Open "select name from dbo.sysobjects where xtype ='U'  order by name", pConn, 3  '3 means adOpenStatic
    
        If Err.Number = 0 Then
           If Not (adoRS.EOF Or adoRS.BOF) Then
              iTemp = adoRS.GetString(, , , ",")
              GetTableList = Mid(iTemp, 1, Len(iTemp) - 1)
           End If
        End If
    
        adoRS.Close

        Set adoRS = Nothing

End Function	
	
Function GetTableCount(ByVal pConn)

        Dim iTemp 
        Dim adoRS 
    
        Set adoRS = Server.CreateObject("ADODB.Recordset")

        adoRS.Open "select count(name) from dbo.sysobjects where xtype ='U' ", pConn, 3  '3 means adOpenStatic

        If Err.Number = 0 Then
           GetTableCount =  adoRS(0)
        End If
    
        adoRS.Close

        Set adoRS = Nothing

End Function	

Function DBList(ByVal pConn)
     Dim pRec
     Dim iTemp

     On Error Resume Next

     Set pRec  = Server.CreateObject("ADODB.RecordSet")
     pRec.Open "SELECT name from sysdatabases where name not in ('master','model','msdb','northwind','pubs','tempdb')  order by name"   , pConn
     
     If Err.number = 0 Then
        If pRec.EOF Or pRec.BOF Then
        Else
          iTemp = iTemp & "<OPTION           value= """">"
          Do While  Not (pRec.EOF Or pRec.BOF)

             If UCase(Decode(Session("SDB"))) = UCase(Trim(pRec(0))) Then
                iTemp = iTemp & "<OPTION  SELECTED value=" & pRec(0) & " >" & pRec(0)
             Else
                iTemp = iTemp & "<OPTION           value=" & pRec(0) & " >" & pRec(0)
             End If   
             pRec.MoveNext
          Loop
        End If  
     End If
     
     DBList = iTemp
     pRec.Close
     Set pRec  = Nothing
  
End Function

Function DBRadioList(ByVal pConn,ByVal pName)
     Dim pRec
     Dim iTemp
     Dim iLoop

     On Error Resume Next


     iTemp = " <table border=0 cellspacing=1 cellpadding=1 bgcolor=#cccccc width=100% >"

     Set pRec  = Server.CreateObject("ADODB.RecordSet")
     pRec.Open "SELECT name from sysdatabases where name not in ('master','model','msdb','northwind','pubs','tempdb') order by name"   , pConn
     iLoop = 0 
     If Err.number = 0 Then
        If pRec.EOF Or pRec.BOF Then
        Else
          Do While  Not (pRec.EOF Or pRec.BOF)
             iLoop = iLoop + 1

             If UCase(Decode(Session("SDB"))) = UCase(Trim(pRec(0))) Then
                iTemp = iTemp & "<TR><TD bgcolor=white><INPUT TYPE=RADIO NAME=" & pName & " ID=" & pName & iLoop & "  VALUE = " & pRec(0) & " CLASS=RADIO>" & pRec(0) & "</TD></TR>"
             Else
                iTemp = iTemp & "<TR><TD bgcolor=white><INPUT TYPE=RADIO NAME=" & pName & " ID=" & pName & iLoop & "  VALUE = " & pRec(0) & " CLASS=RADIO>" & pRec(0) & "</TD></TR>"
             End If   
             iTemp = iTemp & vbCrLf
             pRec.MoveNext
          Loop
        End If  
     End If

    DBRadioList = iTemp & "</TABLE>"

     pRec.Close
     Set pRec  = Nothing
  
End Function


Sub MakeButton(pButtonName,pButtonFtnName,pWIDTH)
   Dim ii

   Response.Write "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=0 WIDTH=" & pWIDTH & " >"  & VBCrLf
   Response.Write "<TR ALIGN=RIGHT>"     & VBCrLf
   For ii = 0 To UBound(pButtonName) - 1
       Response.Write "<TD width=100 CLASS=TDButtonOut OnClick=""" &  pButtonFtnName(ii) &  """ onmousedown=""vbscript:MBMouseDown()""  onmouseup=""vbscript:MBMouseUp()""   onmouseout=""vbscript:MBMouseUp()""  >" & pButtonName(ii) & "</TD>" & vbcr 
   Next    
   
   Response.Write "</TR >"     & VBCrLf
   Response.Write "</TABLE>"  & VBCrLf

End Sub

Sub MakeButton2(pButtonName,pButtonWidth,pButtonFtnName,pWIDTH,pAlign,pUpperHeigt,pLowerHeigt)
   Dim ii

   Call MakeBlankLine(pUpperHeigt)

   Response.Write "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=0 WIDTH=" & pWIDTH & " >"  & VBCrLf
   Response.Write "<TR ALIGN=RIGHT>"     & VBCrLf

   If pAlign = "R" Then
      Response.Write "<TD>&nbsp;</TD>" & vbcr 
   End If
   If pAlign = "C" Then
      Response.Write "<TD>&nbsp;</TD>" & vbcr 
   End If

   For ii = 0 To UBound(pButtonName) - 1
       Response.Write "<TD CLASS=TDButtonOut width = " & pButtonWidth(ii)  & " OnClick=""" &  pButtonFtnName(ii) &  """ onmousedown=""vbscript:MBMouseDown()""  onmouseup=""vbscript:MBMouseUp()""   onmouseout=""vbscript:MBMouseUp()""  align=center><center>" & pButtonName(ii) & "</center></TD>" & vbcr 
   Next    
   
   If pAlign = "L" Then
      Response.Write "<TD>&nbsp;</TD>" & vbcr 
   End If
   If pAlign = "C" Then
      Response.Write "<TD>&nbsp;</TD>" & vbcr 
   End If

   Call MakeBlankLine(pLowerHeigt)

   Response.Write "</TR >"     & VBCrLf
   Response.Write "</TABLE>"  & VBCrLf

End Sub

Function MakeConnString(ByVal pIP,ByVal pID,ByVal pPWD,ByVal pCatalog)

  MakeConnString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID= " & pID & ";password= " & pPWD & ";Initial Catalog=" & pCatalog & " ;Data Source= " & pIP
  
End Function  


Function Decode(ByVal strDecode)
 
    Dim i
    Dim lens
    Dim temp
    Dim conv
 
    temp = strDecode
    lens = Len(strDecode)
    temp = ""

    For i = lens To 1 Step -1
        conv = i Mod 3
        temp = temp + Chr((Asc(Mid(strDecode, i, 1)) - conv))
    Next

    Decode = temp
 
End Function



Function Encode(ByVal strEncode)
 
 Dim i
 Dim lens
 Dim temp
 Dim conv
 
 temp = strEncode
 lens = Len(strEncode)
 temp = ""

 For i = 1 To lens Step 1
     conv = i Mod 3
     temp = temp + Chr((Asc(Mid(strEncode, lens - i + 1, 1)) + conv))
 Next

 Encode = temp

End Function



Sub MakeBlankLine(ByVal pHeight)
   Response.Write "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=0 WIDTH=100% >"  & VBCrLf
   Response.Write "<TR HEIGHT=" & pHeight & ">"     & VBCrLf
   Response.Write "<TD WIDTH=100% ></TD>"  & VBCrLf
   Response.Write "</TR>"     & VBCrLf
   Response.Write "</TABLE>"  & VBCrLf
End Sub


Sub HideStatusWnd(ByVal pValue)
	Response.Write "<Script LANGUAGE=VBScript>" & vbCrLf
	Response.Write "Sub Document_onReadyStateChange()" & vbCrLf
	Response.Write "Call LayerShowHide(" & pValue & ")"      & vbCrLf
	Response.Write "End Sub" & vbCrLf
 	Response.Write "</Script>" & vbCrLf
    If Response.Buffer Then Response.Flush
End Sub


Sub WriteTitle(ByVal pTitle)

    Call MakeBlankLine(3)

	Response.Write "<table border=0 cellspacing=1 cellpadding=1 CLASS=TAB13 width=300>" & vbCrLf
	Response.Write "  <tr>" & vbCrLf
	Response.Write "    <TD>" & vbCrLf
	Response.Write "     <table border=0 cellspacing=1 cellpadding=1 bgcolor=#CCCCCC width=100% >" & vbCrLf
	Response.Write "         <tr><TD bgcolor=#E7F1D9 WIDTH=30% ><CENTER>" & pTitle & "</CENTER></TD></tr>" & vbCrLf
	Response.Write "     </table>" & vbCrLf
	Response.Write "    </TD>" & vbCrLf
	Response.Write "  </tr>" & vbCrLf
	Response.Write "</table>" & vbCrLf

    Call MakeBlankLine(3)

End Sub

Function CalElaspeTime(ByVal A, ByVal B)
    
    Dim iTemp
    Dim iHour
    Dim iMinute
    Dim iSecond
    
    iTemp = B - A
    
    If iTemp > 3600 Then
       iHour = CInt((iTemp - (iTemp Mod 3600)) / 3600)
       iTemp = iTemp Mod 3600
    Else
       iHour = 0
    End If
    
    
    If iTemp > 60 Then
       iMinute = CInt((iTemp - (iTemp Mod 60)) / 60)
    Else
       iMinute = 0
    End If
    
    iSecond = CInt(iTemp Mod 60)
    
    CalElaspeTime = Right("0" & iHour, 2) & ":" & Right("0" & iMinute, 2) & ":" & Right("0" & iSecond, 2)
    
End Function

Function GetForeignTable(ByVal pConn, ByVal pTable)
     
     Dim iTemp 
     Dim adoRec 
     Dim adoConn
        
'     iSTRSQL =           " Select so2.Name as ToTable "
 '    iSTRSQL = iSTRSQL & " From  sysforeignkeys fk (nolock) "
  '   iSTRSQL = iSTRSQL & " JOIN sysobjects  so (nolock) on so.[id] = fk.fkeyid "
   '  iSTRSQL = iSTRSQL & " JOIN sysobjects  so1 (nolock) on fk.constid = so1.id "
   ''  iSTRSQL = iSTRSQL & " join syscolumns  sc (nolock) on fk.fkeyid = sc.id and fk.fkey = sc.colid "
   '  iSTRSQL = iSTRSQL & " JOIN sysobjects  so2 (nolock) on fk.rkeyid = so2.id "
   '  iSTRSQL = iSTRSQL & " join syscolumns  sc1 (nolock) on fk.rkeyid = sc1.id and fk.rkey = sc1.colid "
   '  iSTRSQL = iSTRSQL & " where so.name = '" & pTable & "' "

     Set adoRec   = Server.CreateObject("ADODB.Recordset")
     Set adoConn = Server.CreateObject("ADODB.Connection")
     
     adoConn.open pConn
     
     adoConn.execute " usp_tableDependencies '" & pTable & "'"   
     
     GetForeignTable = ""
     
     If Err.Number = 0 Then

        iTemp = "select parentTable from dbo.frtables "

        adoRec.Open iTemp , pConn, 3  '3 means adOpenStatic
    
        If Err.Number = 0 Then
           If Not (adoRec.EOF Or adoRec.BOF) Then
              iTemp = adoRec.GetString(, , , ",")
              GetForeignTable = Mid(iTemp, 1, Len(iTemp) - 1)
           End If
        End If
    
        adoRec.Close

        Set adoRec = Nothing
        
      End If   

End Function


Function execSQL(ByVal pConn,ByVal pSQL)

        Dim iTemp 
        Dim adoConn 
        
        On Error Resume Next
        
        Set adoConn = Server.CreateObject("ADODB.Connection")
        
        execSQL = ""
           
        adoConn.Open pConn
        adoConn.Execute pSQL
        
        If Err.number <> 0 Then
           execSQL = Err.Description 
        End If
    
        adoConn.Close

        Set adoConn = Nothing

End Function	
	
'=======================================================================================================================
' Desc : This function Create DataBase[pTDB] based on the pSDB
' Arg1 : pSDB : Source DataBase
' Arg2 : pTDB : Target DataBase
'=======================================================================================================================
Function CopyDBToDB(ByVal pConn, ByVal pSDB, ByVal pTDB)
    
    Dim iTemp
    
    On Error Resume Next
    
    If pSDB = "" Then
       CopyDBToDB = "Source 데이터베이스명이 지정되지 않았습니다."
       Exit Function
    End If
    
    If pTDB = "" Then
       CopyDBToDB = "생성될 데이터베이스명이 지정되지 않았습니다."
       Exit Function
    End If
    
  
    iTemp = execSQL(pConn, "dbo.dmoCopyDB '" & pSDB & "','" & pTDB & "'")

    If iTemp <> "" Then
        CopyDBToDB = iTemp
    Else
        CopyDBToDB = "데이터베이스가 정상적으로 생성 되었습니다."
    End If
    
End Function


Sub Crete_usp_tableDependencies(ByVal pConn) 
     Dim strSQL

     strSQL = ""     
     strSQL = strSQL  & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[frtables]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)"
     strSQL = strSQL  & "drop table [dbo].[frtables]     "

     Call execSQL(pConn,strSQL)
     
     strSQL = ""     
     strSQL = strSQL  & "create table frtables ("
     strSQL = strSQL  & "  processed int,"
     strSQL = strSQL  & " tableLevel int,"
     strSQL = strSQL  & " childTable sysname,"
     strSQL = strSQL  & "  parentTable sysname,"
     strSQL = strSQL  & ")"
    
     Call execSQL(pConn,strSQL)
    

     strSQL = ""
     strSQL = strSQL  & "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_tableDependencies]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)"
     strSQL = strSQL  & "drop procedure [dbo].[usp_tableDependencies] "

     Call execSQL(pConn,strSQL)
     
     strSQL = ""
     strSQL = strSQL  & "create procedure usp_tableDependencies(" & vbCrLf
     strSQL = strSQL  & " @tableName sysname" & vbCrLf
     strSQL = strSQL  & " ) " & vbCrLf
     strSQL = strSQL  & " as" & vbCrLf
     strSQL = strSQL  & "  begin" & vbCrLf
     strSQL = strSQL  & "    declare @rowsProcessed int" & vbCrLf
     strSQL = strSQL  & "    " & vbCrLf
     strSQL = strSQL  & "    truncate table frtables" & vbCrLf
     strSQL = strSQL  & "    " & vbCrLf
     strSQL = strSQL  & "    insert into frtables " & vbCrLf
     strSQL = strSQL  & "      select 0, 1, childs.name childTable, parents.name parentTable" & vbCrLf
     strSQL = strSQL  & "      from sysobjects childs inner join sysforeignkeys fkeys " & vbCrLf
     strSQL = strSQL  & "        on childs.id = fkeys.fkeyid inner join  sysobjects parents " & vbCrLf
     strSQL = strSQL  & "        on fkeys.rkeyid = parents.id" & vbCrLf
     strSQL = strSQL  & "      where (childs.name = @tableName)" & vbCrLf
     strSQL = strSQL  & "    set @rowsProcessed = @@rowcount" & vbCrLf
     strSQL = strSQL  & "    " & vbCrLf
     strSQL = strSQL  & "    while (@rowsProcessed > 0)" & vbCrLf
     strSQL = strSQL  & "      begin" & vbCrLf
     strSQL = strSQL  & "        update frtables " & vbCrLf
     strSQL = strSQL  & "          set processed = 1 " & vbCrLf
     strSQL = strSQL  & "        where processed = 0 " & vbCrLf
     strSQL = strSQL  & "    " & vbCrLf
     strSQL = strSQL  & "        insert into frtables " & vbCrLf
     strSQL = strSQL  & "          select distinct 0, 1, childs.name childTable, parents.name parentTable" & vbCrLf
     strSQL = strSQL  & "          from sysobjects childs inner join sysforeignkeys fkeys " & vbCrLf
     strSQL = strSQL  & "            on childs.id = fkeys.fkeyid inner join  sysobjects parents " & vbCrLf
     strSQL = strSQL  & "            on fkeys.rkeyid = parents.id inner join frtables" & vbCrLf
     strSQL = strSQL  & "            on childs.name = frtables.parentTable" & vbCrLf
     strSQL = strSQL  & "          where (frtables.processed = 1) " & vbCrLf
     strSQL = strSQL  & "            and (childs.name <> parents.name)" & vbCrLf
     strSQL = strSQL  & "          order by childs.name" & vbCrLf
     strSQL = strSQL  & "        set @rowsProcessed = @@rowcount" & vbCrLf
     strSQL = strSQL  & "        " & vbCrLf
     strSQL = strSQL  & "        update frtables " & vbCrLf
     strSQL = strSQL  & "          set processed = 2 " & vbCrLf
     strSQL = strSQL  & "        where processed = 1" & vbCrLf
     strSQL = strSQL  & "    " & vbCrLf
     strSQL = strSQL  & "      end" & vbCrLf
     strSQL = strSQL  & "    " & vbCrLf
 '   strSQL = strSQL  & "    select childTable, parentTable from frtables" & vbCrLf
 '   strSQL = strSQL  & "    drop table frtables " & vbCrLf
     strSQL = strSQL  & "  end" & vbCrLf
     strSQL = strSQL  & "" & vbCrLf


     Call execSQL(pConn,strSQL)
     
End Sub



Sub DeleteAllTables(ByVal pConn) 
     Dim strSQL
    
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " declare @owner sysname    " & vbCrLF
     strSQL = strSQL &  " declare @tn sysname    " & vbCrLF
     strSQL = strSQL &  " declare @sql nvarchar(4000)    " & vbCrLF
     strSQL = strSQL &  " declare @usertable smallint    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " create table #tbltemp (oType smallint, oObjName sysname, oOwner sysname, oSequence smallint)    " & vbCrLF
     strSQL = strSQL &  " create table #DepList (objid int null, objtype smallint null)    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " /* just user tables*/    " & vbCrLF
     strSQL = strSQL &  " set @usertable = 3    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " insert #DepList select id, @usertable from sysobjects where objectproperty(id, 'IsTable') = 1 and objectproperty(id, 'IsMSShipped') = 0    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " /* flags = 131080 comes from */    " & vbCrLF
     strSQL = strSQL &  " /* 3  - user table + */    " & vbCrLF
     strSQL = strSQL &  " /*  0x20000 (131072) = descending return order*/    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " /* if you want to include transaction, add the */    " & vbCrLF
     strSQL = strSQL &  " /* corresponding begin, commit or rollback transaction*/    " & vbCrLF
     strSQL = strSQL &  " /* and change @intrans to 1 */    " & vbCrLF
     strSQL = strSQL &  " insert #tbltemp EXEC sp_MSdependencies    @objname = null,     " & vbCrLF
     strSQL = strSQL &  "                   @objtype = null,     " & vbCrLF
     strSQL = strSQL &  "                   @flags = 131080,     " & vbCrLF
     strSQL = strSQL &  "                   @objlist = '#DepList',    " & vbCrLF
     strSQL = strSQL &  "                   @intrans = 0    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " drop table #DepList    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " declare mycur cursor    " & vbCrLF
     strSQL = strSQL &  " forward_only    " & vbCrLF
     strSQL = strSQL &  " for    " & vbCrLF
     strSQL = strSQL &  " select oOwner, oObjName from #tbltemp    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " open mycur    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " fetch next from mycur into @owner, @tn    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " while @@fetch_status = 0    " & vbCrLF
     strSQL = strSQL &  "   begin    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  "   set @owner = rtrim(cast(@owner as char))    " & vbCrLF
     strSQL = strSQL &  "         set @tn = @owner + '.' + @tn    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  "   if exists (select * from sysreferences where rkeyid = object_id(@tn))    " & vbCrLF
     strSQL = strSQL &  "                 /*table referenced by a FOREIGN KEY constraint */      " & vbCrLF
     'strSQL = strSQL &  "       set @sql = 'delete ' + @tn    " & vbCrLF
     strSQL = strSQL &  "       set @sql = 'alter table ' + @tn + ' disable trigger all    delete ' + @tn  + '   alter table ' + @tn + ' enable trigger all '    " & vbCrLF
     strSQL = strSQL &  "   else    " & vbCrLF
     strSQL = strSQL &  "       set @sql = 'alter table ' + @tn + ' disable trigger all     truncate table ' + @tn   + '   alter table ' + @tn + ' enable trigger all '  " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  "   execute sp_executesql @sql    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  "   print @sql      " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  "   fetch next from mycur into @owner, @tn    " & vbCrLF
     strSQL = strSQL &  "   end    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " close mycur    " & vbCrLF
     strSQL = strSQL &  " deallocate mycur    " & vbCrLF
     strSQL = strSQL &  "     " & vbCrLF
     strSQL = strSQL &  " drop table #tbltemp    " & vbCrLF

	'Response.Write strSQL & "<br>"
	'Response.End 
     Call execSQL(pConn,strSQL)
     
End Sub

%>

