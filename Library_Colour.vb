
    'For MessageColor  
    Dim  MessageColor
    Set  MessageColor = Log.CreateNewAttributes()
    MessageColor.BackColor = BuiltIn.clMoneyGreen
    MessageColor.FontColor = BuiltIn.clWindowText
    MessageColor.Bold=True 

    'For Error Color
    Dim  ErrorColor
    Set  ErrorColor = Log.CreateNewAttributes()
    ErrorColor.BackColor = BuiltIn.clRed
    ErrorColor.FontColor = BuiltIn.clWindowText
    ErrorColor.Bold=True 
  
    'For Divide Color
    Dim  DivideColor
    Set  DivideColor = Log.CreateNewAttributes()
    DivideColor.BackColor = BuiltIn.clPurple
    DivideColor.FontColor = BuiltIn.clWhite
    DivideColor.Bold=True 
    
    'For Divide Color
    Dim  DivideColor2
    Set  DivideColor2 = Log.CreateNewAttributes()
    DivideColor2.BackColor = BuiltIn.clGray
    DivideColor2.FontColor = BuiltIn.clWhite
    DivideColor2.Bold=True
    
    'For SQL checks
    Dim SqlDivideColor
    Set SqlDivideColor = Log.CreateNewAttributes()
    SqlDivideColor.BackColor = BuiltIn.clSkyBlue
    SqlDivideColor.FontColor = BuiltIn.clWindowText
    SqlDivideColor.Bold=True 
