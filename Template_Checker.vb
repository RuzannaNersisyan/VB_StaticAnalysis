'USEUNIT Library_Common

' Checking template data with database 
public function CheckTemplate(name, caption, ecaption, ttype, docConnected,updatable, _
                              filePath, changeDate) 
  Log.Message("started....")
  Set dbConnection= ADO.CreateConnection
                                                                               
  dbConnection.ConnectionString = cConnectionString
  dbConnection.Open
                           
  Set dbCommand = ADO.CreateCommand
  dbCommand.ActiveConnection = dbConnection
  dbCommand.CommandType = adCmdText
  dbCommand.CommandText = " select fCAPTION, fECAPTION, fIMAGE, fCHANGEDATE, fROWID, fFILE, fUPDATABLE, fDOCCONNECTED, fCHANGEDATE"& _
                            " from TEMPLATES where fNAME=? and fTYPE=?"
  
  Set pName = dbCommand.CreateParameter("pName", DB.adChar, DB.adParamInput, 20)
  Set pType = dbCommand.CreateParameter("pType", DB.adBoolean, DB.adParamInput)  
                  
  dbCommand.Parameters.Append pName
  dbCommand.Parameters.Append pType
  pName.Value = name
  pType.Value = ttype
  
  Set commandResult = dbCommand.Execute
  
  if commandResult.RecordCount = 0 Then
    errorText = "Record with Name=" & name & " Type=" & ttype & " params could not be found."
    Log.Error(errorText)
  else
        Log.Message("Record was found!!!") 
        if commandResult.RecordCount = 1 Then
          if Trim(commandResult("fCAPTION").Value) <> Trim(caption) Then
              Log.error("fCAPTION: Actual = " & Trim(commandResult("fCAPTION").Value)& _
                        ", Expected = " & Trim(caption))
          end if 
          
          if Trim(commandResult("fECAPTION").Value) <> Trim(ecaption) Then
              Log.error("fECAPTION: Actual = " & Trim(commandResult("fECAPTION").Value)& _
                        ", Expected = " & Trim(ecaption))
          end if
          
          if commandResult("fUPDATABLE").Value <> updatable Then
              Log.error("fUPDATABLE: Actual = " & commandResult("fUPDATABLE").Value& _
                        ", Expected = " & updatable)
          end if
          
          if commandResult("fDOCCONNECTED").Value <> docConnected Then
              Log.error("fDOCCONNECTED: Actual = " & commandResult("fDOCCONNECTED").Value& _
                        ", Expected = " & docConnected)
          end if
          
          if Trim(commandResult("fFILE").Value) <> Trim(filePath) Then
              Log.error("fFILE: Actual = " & Trim(commandResult("fFILE").Value)& _
                        ", Expected = " & Trim(filePath))
          end if
          
          dbdate = left(commandResult("fCHANGEDATE").Value, 8)
          if  dbDate <> changeDate Then
              Log.error("fCHANGEDATE: Actual = " & dbDate & _
                        ", Expected = " & changeDate)
          end if
          
          CheckTemplate = commandResult("fROWID").Value
        Else
            Log.Error("More then one record found.")  
        end if     
  end If     
  
  dbConnection.Close                 
end Function


' Checking whether Template with specified name and type exists 
public Function CheckDeleteTemplate(name, ttype)
     Log.Message("Checking delete template started....")
     rowID = FindTemplate(name, ttype)     
       
     if rowid = -1 Then
        Log.Message("Delete template was successful.")
        CheckDeleteTemplate = true 
     Else
        errorText = "Record with Name=" & name & " Type=" & ttype & " params was found."        
        Log.Error(errorText)
        CheckDeleteTemplate = false 
     end if                                                 
end function

' Finds Template with specified name and type 
' returns rowid if exists and -1 otherwise
public function FindTemplate(name, ttype) 
                                         
  Set dbConnection= ADO.CreateConnection
                                                                               
  dbConnection.ConnectionString = cConnectionString
  dbConnection.Open
                           
  Set dbCommand = ADO.CreateCommand
  dbCommand.ActiveConnection = dbConnection
  dbCommand.CommandType = adCmdText
  dbCommand.CommandText = " select fROWID from TEMPLATES where fNAME=? and fTYPE=?"
  
  Set pName = dbCommand.CreateParameter("pName", DB.adChar, DB.adParamInput, 20)
  Set pType = dbCommand.CreateParameter("pType", DB.adBoolean, DB.adParamInput)  
                  
  dbCommand.Parameters.Append pName
  dbCommand.Parameters.Append pType
  pName.Value = name
  pType.Value = ttype
  
  Set commandResult = dbCommand.Execute
  
  if commandResult.RecordCount = 0 Then        
    FindTemplate = -1 
  else    
    FindTemplate = commandResult("fROWID").Value
  end if 
  
  dbConnection.Close
end Function


' Chcking whether template with specified name and type contain expacted data for imoart file
public function CheckImportFile(name, ttype, filePath)
  
  Log.Message("Checking import file for template started....")
  Set dbConnection= ADO.CreateConnection
                                                                               
  dbConnection.ConnectionString = cConnectionString
  dbConnection.Open
                           
  Set dbCommand = ADO.CreateCommand
  dbCommand.ActiveConnection = dbConnection
  dbCommand.CommandType = adCmdText
  dbCommand.CommandText = " select fROWID, fFILE, fIMAGE from TEMPLATES where fNAME=? and fTYPE=?"
  
  Set pName = dbCommand.CreateParameter("pName", DB.adChar, DB.adParamInput, 20)
  Set pType = dbCommand.CreateParameter("pType", DB.adBoolean, DB.adParamInput)  
                  
  dbCommand.Parameters.Append pName
  dbCommand.Parameters.Append pType
  pName.Value = name
  pType.Value = ttype
  
  Set commandResult = dbCommand.Execute
  
 if commandResult.RecordCount = 0 Then
    errorText = "Record with Name=" & name & " Type=" & ttype & " params could not be found."
    Log.Error(errorText)
    CheckImportFile = false 
  else
        Log.Message("Record was found!!!") 
        if commandResult.RecordCount = 1 Then
            if Trim(commandResult("fFILE").Value) <> Trim(filePath) Then
              Log.error("fFILE: Actual = " & Trim(commandResult("fFILE").Value)& _
                        ", Expected = " & Trim(filePath))
            end if
            
            ' checking for file content is needed 
        Else
           Log.Error("More then one record found.")
        end If
        
        CheckImportFile = True        
  end if
  dbConnection.Close
end Function

' check whether specified template conatin mapping for specified document
public Function CheckTemplateMapping(name, ttype, docType, access)

  Log.Message("Checking for template and document mapping started....")
  rowid = FindTemplate(name, ttype)

  if rowid <> -1 then
    Set dbConnection= ADO.CreateConnection                                                                                      
    dbConnection.ConnectionString = cConnectionString
    dbConnection.Open
    
    Set mappingCommand = ADO.CreateCommand
    mappingCommand.ActiveConnection = dbConnection
    mappingCommand.CommandType = adCmdText
    mappingCommand.CommandText = " select * from TEMPLATESMAPPING where fROWID=? and fDOCTYPE=?"
  
    Set pRowID = mappingCommand.CreateParameter("pRowID", DB.adInteger, DB.adParamInput)
    Set pDocType = mappingCommand.CreateParameter("pDocType", DB.adChar, DB.adParamInput, 8)  
                  
    mappingCommand.Parameters.Append pRowID
    mappingCommand.Parameters.Append pDocType
    pRowID.Value = rowid
    pDocType.Value = docType
  
    Set mappingcommandResult = mappingCommand.Execute
    
    if mappingcommandResult.RecordCount = 0 Then
        errorText = "Record with rowid=" & rowid & " docType=" & docType & " params could not be found."
        Log.Error(errorText)
        CheckTemplateMapping = false 
    else
        Log.Message("Record was found!!!") 
        if mappingcommandResult.RecordCount = 1 Then
          if Trim(mappingcommandResult("fACCESS").Value)<> Trim(access) then
            Log.error("fACCESS: Actual = " & Trim(mappingcommandResult("fACCESS").Value)& _
                        ", Expected = " & Trim(access))
          end if
          CheckTemplateMapping = True           
        Else
          Log.Error("More then one row found.")
          CheckTemplateMapping = false 
        end if       
    end if     
        
    dbConnection.Close
  Else
    errorText = "Record with Name=" & name & " Type=" & ttype & " params could not be found."
    Log.Error(errorText)
    CheckTemplateMapping = false  
  end if               
end Function

' Checking whether Template mapping with specified name and type exists
public Function CheckDeleteTemplateMapping(name, ttype, docType)
  
  Log.Message("Checking for delete template and document mapping started....")
  rowid = FindTemplate(name, ttype)

  if rowid <> -1 then
    Set dbConnection= ADO.CreateConnection                                                                                      
    dbConnection.ConnectionString = cConnectionString
    dbConnection.Open
    
    Set mappingCommand = ADO.CreateCommand
    mappingCommand.ActiveConnection = dbConnection
    mappingCommand.CommandType = adCmdText
    mappingCommand.CommandText = " select * from TEMPLATESMAPPING where fROWID=? and fDOCTYPE=?"
  
    Set pRowID = mappingCommand.CreateParameter("pRowID", DB.adInteger, DB.adParamInput)
    Set pDocType = mappingCommand.CreateParameter("pDocType", DB.adChar, DB.adParamInput, 8)  
                  
    mappingCommand.Parameters.Append pRowID
    mappingCommand.Parameters.Append pDocType
    pRowID.Value = rowid
    pDocType.Value = docType
  
    Set mappingcommandResult = mappingCommand.Execute
    
    if mappingcommandResult.RecordCount = 0 Then
        Log.Message("Template mapping was not found.")
        CheckDeleteTemplateMapping = true 
    else
        errorText = "Record with rowid=" & rowid & " docType=" & docType & " exists."
        Log.Error(errorText)
        CheckDeleteTemplateMapping = false    
    end if     
        
    dbConnection.Close
  Else
    errorText = "Record with Name=" & name & " Type=" & ttype & " params could not be found."
    Log.Error(errorText)
    CheckDeleteTemplateMapping = false
  end if  
End Function






