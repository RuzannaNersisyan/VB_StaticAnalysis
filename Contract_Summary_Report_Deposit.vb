Option Explicit
'USEUNIT Library_Common
'USEUNIT Contract_Summary_Report_Library

'Test Case N 165043

Sub Contract_Summary_Report_Deposit_Check_Rows_Test()

  Dim startDATE, fDATE, Date, cont_date
                                            
  Date = "201211"                
  cont_date = "111111"
  Utilities.ShortDateFormat = "yyyymmdd"
  startDATE = "20031210"
  fDATE = "20251221"
    
  'Test StartUp 
  Call Initialize_AsBank("bank", startDATE, fDATE)
  Call ChangeWorkspace("Ավանդներ (ներգրավված)")
  Call wTreeView.DblClickItem("|²í³Ý¹Ý»ñ (Ý»ñ·ñ³íí³Í)|ä³ÛÙ³Ý³·ñ»ñÇ ³Ù÷á÷áõÙ")
  
  Call Contract_Sammary_Report_Fill(Date, Null, Null, Null, Null, Null, Null, Null, _
                                      Null, Null, Null, Null, Null, Null, Null, _
                                      Null, Null, Null, Null, Null, Null, Null, False, False, _
                                      Null, False, False, False, _
                                      True, True, True, True, True, _
                                      True, True, True, True, True, True, _
                                      True, True, True, False, True, False, False,3)
    
  BuiltIn.Delay(70000)
    
  '¶áõÙ³ñ ëÛ³Ý ëïáõ·áõÙ
  Call Compare_ColumnFooterVlaue("frmPttel", "FAGRSUM", "170,039,106.90")
  'Ä³ÙÏ»ï³Ýó ·áõÙ³ñ ëÛ³Ý ëïáõ·áõÙ
  Call Compare_ColumnFooterVlaue("frmPttel", "FAGRSUMJ", "139,305,606.90")
  'îáÏáë ëÛ³Ý ëïáõ·áõÙ
  Call Compare_ColumnFooterVlaue("frmPttel", "FPERSUM", "11,680,397.16")
  '²ñ¹ÛáõÝ³í»ï îáÏáë ëÛ³Ý ëïáõ·áõÙ
  Call Compare_ColumnFooterVlaue("frmPttel", "FEFFINC", "1,185.90")
  'Ä³Ï»ï³Ýó ïáÏáë ëÛ³Ý ëïáõ·áõÙ
  Call Compare_ColumnFooterVlaue("frmPttel", "FPERSUMJ", "2,687,865.39")
  'Ä³ÙÏ»ï³Ýó ·áõÙ³ñÇ ïáõÛÅ ëÛ³Ý ëïáõ·áõÙ
  Call Compare_ColumnFooterVlaue("frmPttel", "FPENSUM", "84,466.67")
  'Ä³ÙÏ»ï³Ýó ïáÏáëÇ ïáõÛÅ ëÛ³Ý ëïáõ·áõÙ
  Call Compare_ColumnFooterVlaue("frmPttel", "FPENSUM2", "301.45")
  '¶ñ³íÇ ³ñÅ»ùÁ ëÛ³Ý ëïáõ·áõÙ
  Call Compare_ColumnFooterVlaue("frmPttel", "FMORTGAGESUM", "34,375.00")
  'ºñ³ßË³íáñáõÃÛ³Ý ³ñÅ»ùÁ ëÛ³Ý ëïáõ·áõÙ
  Call Compare_ColumnFooterVlaue("frmPttel", "FGUARSUM", "0.00")
  'ä³ÛÙ³Ý³·ñÇ ·áõÙ³ñ ëÛ³Ý ëïáõ·áõÙ
  Call Compare_ColumnFooterVlaue("frmPttel", "FSUMMA", "173,536,606.05")
'  'îñÙ³Ý ûñ»ñÇ ù³Ý³Ï ëÛ³Ý ëïáõ·áõÙ
'  Call Compare_ColumnFooterVlaue("frmPttel", "FDAYQUAN", "289248")
'  'Ø³ñÙ³ÝÁ ÙÝ³ó³Í ûñ»ñÇ ù³Ý³Ï ëÛ³Ý ëïáõ·áõÙ
'  Call Compare_ColumnFooterVlaue("frmPttel", "FDAYBEFMAR", "-1389192")
'  'Ø³ñÙ³ÝÁ ÙÝ³ó³Í ûñ»ñÇ ù.³.Ù.Å. ëÛ³Ý ëïáõ·áõÙ             
'  Call Compare_ColumnFooterVlaue("frmPttel", "FDAYBEFFRMAR", "-1416985")
'  'ºñÏ³ñ³Ó·í³Í íÇ×³ÏáõÙ ·ïÝíáÕ ûñ»ñÇ ù³Ý³ÏÁ ëÛ³Ý ëïáõ·áõÙ
'  Call Compare_ColumnFooterVlaue("frmPttel", "FDAYPROL", "28246")
'  'ºñÏ³ñ³Ó·í³Í ûñ»ñÇ ù³Ý³Ï ëÛ³Ý ëïáõ·áõÙ
'  Call Compare_ColumnFooterVlaue("frmPttel", "FDAYPROLALL", "28167")
'  'îáÏáëÝ»ñÇ Ù³ñÙ³ÝÁ ÙÝ³ó³Í ûñ»ñÇ ù³Ý³Ï ëÛ³Ý ëïáõ·áõÙ
'  Call Compare_ColumnFooterVlaue("frmPttel", "FDAYBEFPRMAR", "-1394208")
'  'Ä³ÙÏ»ï³Ýó ûñ»ñÇ ù³Ý³Ï ëÛ³Ý ëïáõ·áõÙ
'  Call Compare_ColumnFooterVlaue("frmPttel", "FDAYAGRJ", "1219781")
'  'îáÏáëÝ»ñÇ Å³ÙÏ»ï³Ýó ûñ»ñÇ ù³Ý³Ï ëÛ³Ý ëïáõ·áõÙ
'  Call Compare_ColumnFooterVlaue("frmPttel", "FDAYPERJ", "687982")
  'Ü»ñÏ³ ³ñÅ»ù ëÛ³Ý ëïáõ·áõÙ
  Call Compare_ColumnFooterVlaue("frmPttel", "FPRESVALUE", "181,720,689.96")

  Call wMainForm.MainMenu.Click("Դիտում |DepositFilter")
  BuiltIn.Delay(1000) 
        
  'Test CleanUp 
  Call Close_AsBank()
End Sub
