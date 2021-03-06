Option Explicit
'USEUNIT Payment_Order_ConfirmPhases_Library
'USEUNIT Loan_Agreements_Library
'USEUNIT Subsystems_SQL_Library
'USEUNIT Akreditiv_Library
'USEUNIT Library_Common
'USEUNIT Constants

Sub Check_DB_Test()

    Dim startDATE , fDATE ,fBASE
    Dim queryString,sql_Value,colNum,sql_isEqual,result

    startDATE = "20060101"
    fDATE = "20220101"    
     
    'Test StartUp start
    Call Initialize_AsBank("bank", startDATE, fDATE)
    
    Call Create_Connection()
  
    Call ChangeWorkspace(c_Admin40)
    Call wTreeView.DblClickItem("|²¹ÙÇÝÇëïñ³ïáñÇ ²Þî 4.0|Ð³Ù³Ï³ñ·³ÛÇÝ ³ßË³ï³ÝùÝ»ñ|îíÛ³ÉÝ»ñÇ ëïáõ·áõÙ|î´ Ï³éáõóí³ÍùÇ ëïáõ·áõÙ")
    BuiltIn.Delay(2000)
    'Ստուգում է որ ներմուծվածլինի ճիշտ քանակությամբ տող 
    If Not Sys.Process("Asbank").VBObject("MainForm").Window("MDIClient", "", 1).VBObject("frmPttel").VBObject("tdbgView").ApproxCount =38  then
          Log.Error("Count is not 38")
    End If 

  
'''-------------------------------------------------------------------------------------------------------------------------------------------------------
'''''''''''' Ստուգում է փոխված պարամետրերը 
      result = " PARAMETER {NAME = CREDITCODE;  CAPTION = ""ì³ñÏ³ÛÇÝ Ïá¹"";  TYPE = ""C(20)"";  UI = 1;  };" & vbCrLf &_
               " PARAMETER {NAME = NBOUTSTATE;  CAPTION = ""Ð»ïÑ³ßí»ÏßéÇó ¹áõñë·ñÙ³Ý íÇ×³Ï"";  TYPE = ""C(10)"";  UI = 1;  };" & vbCrLf &_
               " PARAMETER {NAME = NEWLRCODE;  CAPTION = ""ìè Ïá¹(Ýáñ)"";  TYPE = ""C(21)"";  };"& vbCrLf &_
               " };" & vbCrLf & ""


       'Կատարում ենք SQL ստուգում
       queryString = " select Substring(fDESCR, 47759, 100000) from USERREPORTS where fNAME = 'ASTOTL'"
       sql_Value = result
       colNum = 0 
       sql_isEqual =  CheckDB_Value(queryString, sql_Value, colNum)
       If Not sql_isEqual Then
          Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
       Else
          Log.Message("Check was succesfull")
       End If
'''------------------------------------------------------------------------------------------------------------------------------------------------------      
       
      
'''------------------------------------------------------------------------------------------------------------------------------------------
'''''''''' Ստուգում է փոխված պարամետրերը   
     result = " PARAMETER {NAME = PARAM36;  CAPTION = ""CondClRekvs"";  TYPE = ""C(2000)"";  UI = 1;  };" & vbCrLf & _
              " PARAMETER {NAME = PARAM37;  CAPTION = ""Reserve"";  TYPE = ""C(1500)"";  UI = 1;  };" & vbCrLf &_
              " PARAMETER {NAME = PARAM38;  CAPTION = ""ShowNotFullClosed"";  TYPE = ""BOOLEAN"";  UI = 1;  PARAMETER = ""0"";  };" & vbCrLf &_
              " PARAMETER {NAME = PARAM39;  CAPTION = ""ShowONLYClosed"";  TYPE = ""BOOLEAN"";  UI = 1;  PARAMETER = ""0"";  };" & vbCrLf &_
              " PARAMETER {NAME = PARAM40;  CAPTION = ""CloseDateStart"";  TYPE = ""DATE"";  UI = 1;  PARAMETER = ""CURRENT_DATE"";  };" & vbCrLf &_
              " PARAMETER {NAME = PARAM41;  CAPTION = ""CloseDateEnd"";  TYPE = ""DATE"";  UI = 1;  PARAMETER = ""CURRENT_DATE"";  };" & vbCrLf &_
              " PARAMETER {NAME = PARAM42;  CAPTION = ""OrderForClicode"";  TYPE = ""BOOLEAN"";  UI = 1;  PARAMETER = ""0"";  };" & vbCrLf &_
              " PARAMETER {NAME = PARAM43;  CAPTION = ""òáõÛó ï³É Ñ³ßÇíÝ»ñÁ"";  TYPE = ""BOOLEAN"";  UI = 1;  PARAMETER = ""0"";  };" & vbCrLf &_
              " PARAMETER {NAME = PARAM44;  CAPTION = ""ShowSSRekv"";  TYPE = ""BOOLEAN"";  UI = 1;  PARAMETER = ""0"";  };" & vbCrLf &_
              " PARAMETER {NAME = PARAM45;  CAPTION = ""ÎÝùÙ³Ý ³Ùë³ÃÇí ëÏÇ½µ"";  TYPE = ""DATE"";  PARAMETER = ""CURRENT_DATE"";  };" & vbCrLf &_
              " PARAMETER {NAME = PARAM46;  CAPTION = ""ÎÝùÙ³Ý ³Ùë³ÃÇí í»ñç"";  TYPE = ""DATE"";  PARAMETER = ""CURRENT_DATE"";  };" & vbCrLf &_
              " PARAMETER {NAME = PARAM47;  CAPTION = ""ìè Ïá¹(Ýáñ)"";  TYPE = ""C(21)"";  };" & vbCrLf &_
              " PARAMETER {NAME = PARAM48;  CAPTION = ""ShowClosedSafAgrs"";  TYPE = ""BOOLEAN"";  };" & vbCrLf &_
              " };" & vbCrLf & ""

    
     'Կատարում ենք SQL ստուգում
     queryString = " select Substring(fDESCR, 28126, 100000) from USERREPORTS where fNAME = 'MASTOTL'"
     sql_Value = result
     colNum = 0 
     sql_isEqual =  CheckDB_Value(queryString, sql_Value, colNum)
     If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
     Else
        Log.Message("Check was succesfull")
     End If
       
'''-----------------------------------------------------------------------------------------------------------

        
'''------------------------------------------------------------------------------------------------------------------------------------------
'''''''''' Ստուգում է փոխված պարամետրերը   
     result = " PARAMETER {NAME = PARAM23;  CAPTION = ""¶ñ³ë»ÝÛ³Ï"";  TYPE = ""C(10)"";  UI = 1;  };" & vbCrLf & _
               " PARAMETER {NAME = PARAM24;  CAPTION = ""´³ÅÇÝ"";  TYPE = ""C(10)"";  UI = 1;  };" & vbCrLf & _
               " PARAMETER {NAME = PARAM25;  CAPTION = ""Ð³ë³Ý-Ý ïÇå"";  TYPE = ""C(12)"";  UI = 1;  };" & vbCrLf & _
               " };" & vbCrLf & ""

    
     'Կատարում ենք SQL ստուգում
     queryString = " select Substring(fDESCR, 30987, 100000) from USERREPORTS where fNAME = 'ASAGRACC'"
     sql_Value = result
     colNum = 0 
     sql_isEqual =  CheckDB_Value(queryString, sql_Value, colNum)
     If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
     Else
        Log.Message("Check was succesfull")
     End If
       
'''-----------------------------------------------------------------------------------------------------------


    
'''------------------------------------------------------------------------------------------------------------------------------------------
'''''''''' Ստուգում է փոխված պարամետրերը   
     result = " PARAMETER {NAME = ACSBRANCH;  CAPTION = ""¶ñ³ë»ÝÛ³Ï"";  TYPE = ""C(10)"";  UI = 1;  };" & vbCrLf & _
               " PARAMETER {NAME = ACSDEPART;  CAPTION = ""´³ÅÇÝ"";  TYPE = ""C(10)"";  UI = 1;  };" & vbCrLf & _
               " PARAMETER {NAME = ACSTYPE;  CAPTION = ""Ð³ë³Ý-Ý ïÇå"";  TYPE = ""C(12)"";  UI = 1;  };" & vbCrLf & _
               " PARAMETER {NAME = SHOWCLIREKV;  CAPTION = ""òáõÛó ï³É Ñ³×³Ëáñ¹Ý»ñÇ Ñ³ïÏ³ÝÇß»ñÁ"";  TYPE = ""BOOLEAN"";  UI = 1;  PARAMETER = ""0"";  };" & vbCrLf & _
               " PARAMETER {NAME = OUTERCODE;  CAPTION = ""²ñï³ùÇÝ N"";  TYPE = ""C(80)"";  UI = 1;  };" & vbCrLf & _
               " PARAMETER {NAME = GRPEDISN;  CAPTION = ""ÊÙµ³ÛÇÝ ËÙµ³·ñÙ³Ý ISN"";  TYPE = ""NP(10,0)"";  UI = 1;  };" & vbCrLf & _
               " PARAMETER {NAME = SHOWLIMITS;  CAPTION = ""òáõÛó ï³É ÙÝ³óáñ¹Ç ë³ÑÙ³ÝÝ»ñÁ"";  TYPE = ""BOOLEAN"";  UI = 1;  PARAMETER = ""0"";  };" & vbCrLf & _
               " PARAMETER {NAME = SHOWOPENDATE;  CAPTION = ""òáõÛó ï³É µ³óÙ³Ý ³Ùë³ÃÇíÁ"";  TYPE = ""BOOLEAN"";  UI = 1;  PARAMETER = ""0"";  };" & vbCrLf & _ 
               " PARAMETER {NAME = SHOWOLDACC;  CAPTION = ""òáõÛó ï³É ÑÇÝ Ñ³ßÇíÁ"";  TYPE = ""BOOLEAN"";  UI = 1;  PARAMETER = ""0"";  };" & vbCrLf & _
               " PARAMETER {NAME = SHOWNEWACC;  CAPTION = ""òáõÛó ï³É Ýáñ Ñ³ßÇíÁ"";  TYPE = ""BOOLEAN"";  UI = 1;  PARAMETER = ""0"";  };" & vbCrLf & _
               " };" & vbCrLf & ""

    
     'Կատարում ենք SQL ստուգում
     queryString = " select Substring(fDESCR, 6278, 100000) from USERREPORTS where fNAME = 'NBASACCS'"
     sql_Value = result
     colNum = 0 
     sql_isEqual =  CheckDB_Value(queryString, sql_Value, colNum)
     If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
     Else
        Log.Message("Check was succesfull")
     End If
       
'''-----------------------------------------------------------------------------------------------------------


   
'''------------------------------------------------------------------------------------------------------------------------------------------
'''''''''' Ստուգում է փոխված պարամետրերը   
     result = " PARAMETER {NAME = MANAGER;  CAPTION = ""îÝûñ»Ý"";  TYPE = ""C(60)"";  UI = 1;  };" & vbCrLf &_
              " PARAMETER {NAME = MNGCLICODE;  CAPTION = ""îÝûñ»Ý(Ñ³×³Ëáñ¹Ç Ïá¹)"";  TYPE = ""C(12)"";  UI = 1;  };"& vbCrLf & _
              " };" & vbCrLf & ""

    
     'Կատարում ենք SQL ստուգում
     queryString = " select Substring(fDESCR,7453, 100000) from USERREPORTS where fNAME = 'ASOWNERS'"
     sql_Value = result
     colNum = 0 
     sql_isEqual =  CheckDB_Value(queryString, sql_Value, colNum)
     If Not sql_isEqual Then
        Log.Error("Querystring = " & queryString & ":  Expected result = " & sql_Value)
     Else
        Log.Message("Check was succesfull")
     End If
       
'''-----------------------------------------------------------------------------------------------------------

    Call Close_AsBank()

End Sub