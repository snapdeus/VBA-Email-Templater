Sub ShippingConfirmation()
 
 Dim objMail As Outlook.MailItem
 Set objMail = Application.CreateItem(olMailItem)
 

    'Dim DataObj As MSForms.DataObject
    'Set DataObj = New MSForms.DataObject
    'DataObj.GetFromClipboard
    'strPaste = DataObj.GetText(1)
 
 Dim Pth As String
    Dim mth As Long
    Dim imth As String
    Dim yr As Long
    
    
    yr = Year(Date)
    mth = Month(Date)
    imth = MonthName(mth)
        c00 = "D:\Saved Packlist PDF\" & yr & "\" & mth & " " & imth & "\"
        Pth = Split(CreateObject("wscript.shell").exec("cmd /c dir """ & c00 & "*.pdf"" /b/s/a-d/o-d").stdout.readall, vbCrLf)(0)
    
    Dim MainMessage As String
    Dim MainMessagep As String
    Dim SignOff As String
    Dim SignOff2 As String
    Dim PartialOrderTable As String
    Dim PartialOrderTable2 As String
    Dim ExtraRows As String
    Dim EndTable As String
    
MainMessage = "<html><head><style>p{font-family:'Calibri', sans-serif}; span{font-family:'Calibri', sans-serif}</style> <title>Shipping Confirmation</title></head><body> <div class=Preamble> <p style='line-height:14.65pt;font-family:`Arial`, sans-serif;background:white'><span>Hello namevar,</span></p> <p> <span>Thank you for ordering from HAVE, Inc! </span></p> <p>" & _
 "<span>Attached please find a copy of your packing list showing the items on your order that have shipped. Please review and save for your records. </span></p> <p> <span>Your invoice will be sent in a separate email from our accounts department.</span></p>" & _
 "<table border=1 cellspacing=3 cellpadding=0 width=0 style='width:512.6pt;border:outset #F2F2F2 1.0pt'> <tr style='height:24.75pt'> <td width=347 colspan=2 valign=bottom style='width:257.6pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:24.75pt;overflow:hidden'> <p align=center style='text-align:center;background-image:initial;'> <b><span>SHIPPING INFORMATION</span></b></p>" & _
 "</td> <td width=336 colspan=2 valign=bottom style='width:249.0pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:24.75pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>ORDER INFORMATION </span></b></p> </td> </tr> <tr style='height:24.75pt'> <td width=137 valign=bottom style='width:100.1pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.75pt;overflow:hidden'>" & _
 "<p align=center style='text-align:center'><b><span>Ship Date: </span></b></p> </td> <td width=210 valign=bottom style='width:155.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.75pt;overflow:hidden'> <p><span>datevar </span></p> </td> <td width=138 valign=bottom style='width:101.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.75pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Order Number: </span></b></p></td>" & _
 "<td width=198 valign=bottom style='width:145.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.75pt;overflow:hidden'> <p><span>ordervar </span></p> </td> </tr> <tr style='height:33.0pt'> <td width=137 valign=bottom style='width:100.1pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.0pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Ship Via: </span></b></p> </td>" & _
 "<td width=210 valign=bottom style='width:155.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.0pt;overflow:hidden'> <p><span>shipviavar</span></p> </td> <td width=138 valign=bottom style='width:101.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.0pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>WEB Order ID: </span></b></p> </td>" & _
 "<td width=198 valign=bottom style='width:145.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.0pt;overflow:hidden'> <p><span>webordervar </span></p> </td> </tr> <tr style='height:29.25pt'> <td width=137 valign=bottom style='width:100.1pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:29.25pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Tracking Number: </span></b></p> </td>" & _
 "<td width=210 valign=bottom style='width:155.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:29.25pt;overflow:hidden'> <p><span>trackingnumbervar</span></u></p> </td> <td width=138 valign=bottom style='width:101.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:29.25pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>PO Reference: </span></b></p> </td>" & _
 "<td width=198 valign=bottom style='width:145.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:29.25pt;overflow:hidden'> <p><span>POvar</span></p> </td> </tr> <tr style='height:32.25pt'> <td width=137 valign=bottom style='width:100.1pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:32.25pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Transit Time: </span></b></p> </td><td width=210 valign=bottom style='width:155.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:32.25pt;overflow:hidden'> <p><span>transittimevar business days</span> </p> </td> " & _
 "<td width=138 valign=bottom style='width:101.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:32.25pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Pack List ID: </span></b></p> </td> <td width=198 valign=bottom style='width:145.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:32.25pt;overflow:hidden'> <p><span>paclistvar </span></p> </td> </tr> <tr style='height:11.6pt'> <td width=683 colspan=4 valign=bottom style='width:508.6pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:11.6pt;overflow:hidden'> </td> </tr> <tr style='height:15.75pt'> <td width=137 valign=bottom style='width:100.1pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:15.75pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Sales Rep: </span></b></p> </td>" & _
 "<td width=210 valign=bottom style='width:155.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:15.75pt;overflow:hidden'> <p><span>salesrepvar </span></p> </td> <td width=138 valign=bottom style='width:101.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:15.75pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Sales Rep Phone: </span></b></p> </td> <td width=198 valign=bottom style='width:145.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:15.75pt;overflow:hidden'> <p><span>1-800-999-4283 extvar </span></p> </td> </tr> <tr style='height:15.75pt'> <td width=137 valign=bottom style='width:100.1pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:15.75pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Sales Rep Email: </span></b></p> </td>" & _
 "<td width=546 colspan=3 valign=bottom style='width:406.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:15.75pt;overflow:hidden'> <p><span><a href='mailto:jearl@haveinc.com' target='_blank'><span style='color:#1155CC'>emailvar</span></a> </span></p> </td> </tr> </table> <p style='line-height:14.65pt;background:white'></p> <h4 style='line-height:14.65pt;background:white'><span>THE ITEM(S) WERE SHIPPED TO: </span></h4> <p style='margin:0in;margin-bottom:.0001pt;line-height:14.65pt;'> <span>CustomerAddress</span></p>"
    
    
    MainMessagep = "<html><head><style>p{font-family:'Calibri', sans-serif}; span{font-family:'Calibri', sans-serif}</style> <title>Shipping Confirmation</title></head><body> <div class=Preamble> <p style='line-height:14.65pt;font-family:`Arial`, sans-serif;background:white'><span>Hello namevar,</span></p> <p> <span>Thank you for ordering from HAVE, Inc! </span></p> <p>" & _
 "<span>Attached please find a copy of your packing list showing the items on your order that have shipped. Please review and save for your records. Please note this is a partial order.</span></p> <p> <span>Your invoice will be sent in a separate email from our accounts department.</span></p>" & _
 "<table border=1 cellspacing=3 cellpadding=0 width=0 style='width:512.6pt;border:outset #F2F2F2 1.0pt'> <tr style='height:24.75pt'> <td width=347 colspan=2 valign=bottom style='width:257.6pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:24.75pt;overflow:hidden'> <p align=center style='text-align:center;background-image:initial;'> <b><span>SHIPPING INFORMATION</span></b></p>" & _
 "</td> <td width=336 colspan=2 valign=bottom style='width:249.0pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:24.75pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>ORDER INFORMATION </span></b></p> </td> </tr> <tr style='height:24.75pt'> <td width=137 valign=bottom style='width:100.1pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.75pt;overflow:hidden'>" & _
 "<p align=center style='text-align:center'><b><span>Ship Date: </span></b></p> </td> <td width=210 valign=bottom style='width:155.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.75pt;overflow:hidden'> <p><span>datevar </span></p> </td> <td width=138 valign=bottom style='width:101.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.75pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Order Number: </span></b></p></td>" & _
 "<td width=198 valign=bottom style='width:145.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.75pt;overflow:hidden'> <p><span>ordervar </span></p> </td> </tr> <tr style='height:33.0pt'> <td width=137 valign=bottom style='width:100.1pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.0pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Ship Via: </span></b></p> </td>" & _
 "<td width=210 valign=bottom style='width:155.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.0pt;overflow:hidden'> <p><span>shipviavar</span></p> </td> <td width=138 valign=bottom style='width:101.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.0pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>WEB Order ID: </span></b></p> </td>" & _
 "<td width=198 valign=bottom style='width:145.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:33.0pt;overflow:hidden'> <p><span>webordervar </span></p> </td> </tr> <tr style='height:29.25pt'> <td width=137 valign=bottom style='width:100.1pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:29.25pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Tracking Number: </span></b></p> </td>" & _
 "<td width=210 valign=bottom style='width:155.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:29.25pt;overflow:hidden'> <p><span>trackingnumbervar</span></u></p> </td> <td width=138 valign=bottom style='width:101.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:29.25pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>PO Reference: </span></b></p> </td>" & _
 "<td width=198 valign=bottom style='width:145.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:29.25pt;overflow:hidden'> <p><span>POvar</span></p> </td> </tr> <tr style='height:32.25pt'> <td width=137 valign=bottom style='width:100.1pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:32.25pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Transit Time: </span></b></p> </td><td width=210 valign=bottom style='width:155.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:32.25pt;overflow:hidden'> <p><span>transittimevar business days</span> </p> </td> " & _
 "<td width=138 valign=bottom style='width:101.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:32.25pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Pack List ID: </span></b></p> </td> <td width=198 valign=bottom style='width:145.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:32.25pt;overflow:hidden'> <p><span>paclistvar </span></p> </td> </tr> <tr style='height:11.6pt'> <td width=683 colspan=4 valign=bottom style='width:508.6pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:11.6pt;overflow:hidden'> </td> </tr> <tr style='height:15.75pt'> <td width=137 valign=bottom style='width:100.1pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:15.75pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Sales Rep: </span></b></p> </td>" & _
 "<td width=210 valign=bottom style='width:155.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:15.75pt;overflow:hidden'> <p><span>salesrepvar </span></p> </td> <td width=138 valign=bottom style='width:101.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:15.75pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Sales Rep Phone: </span></b></p> </td> <td width=198 valign=bottom style='width:145.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:15.75pt;overflow:hidden'> <p><span>1-800-999-4283 extvar </span></p> </td> </tr> <tr style='height:15.75pt'> <td width=137 valign=bottom style='width:100.1pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:15.75pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Sales Rep Email: </span></b></p> </td>" & _
 "<td width=546 colspan=3 valign=bottom style='width:406.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:15.75pt;overflow:hidden'> <p><span><a href='mailto:jearl@haveinc.com' target='_blank'><span style='color:#1155CC'>emailvar</span></a> </span></p> </td> </tr> </table> <p style='line-height:14.65pt;background:white'></p> <h4 style='line-height:14.65pt;background:white'><span>THE ITEM(S) WERE SHIPPED TO: </span></h4> <p style='margin:0in;margin-bottom:.0001pt;line-height:14.65pt;'> <span>CustomerAddress</span></p>"
    
    
    
SignOff = "<p style='margin:0in;margin-bottom:.0001pt;line-height:14.65pt;'><p style='line-height:14.65pt;background:white'><span> </span></p> <p><span>Once again, thank you for ordering from HAVE Inc. We do appreciate your business!</span> </p> </div></body></html>" & _
 "<div> <p><span >Regards,</span></p> <p class=MsoNormal style='line-height:14.65pt;background:white'><span></span></p><p><span >Stephen Simpson<br></span><span >Customer Service |</span><span >&nbsp;</span><a href='http://www.haveinc.com/' target='_blank'><span style='font-size:9.0pt;color:blue'>HAVE, INC</span></a><span ><br></span><i><span style='font-size:8.0pt;color:#222222'>Pro Audio, Video, Data, &amp; Fiber Cables</span></i><span style=><br></span><i><span style='font-size:8.0pt;color:#444444' >A/V Equipment &amp; Interconnect Specialists</span></i><span style=><br></span><span style='font-size:8.0pt;color:#666666'>1-800-999-4283 x 254&nbsp;</span>" & _
 "<span style=><br></span><span style='font-size:7.5pt;color:#999999'>Certified NYS&nbsp;WBE | Est. 1977</span></p> <p class=MsoNormal><b><span style='font-size:7.5pt;'>Shipping Policy:</span></b><span style='font-size:7.5pt;'>&nbsp;All orders are shipped F.O.B. primarily from our headquarters location in Hudson, NY 12534. Orders are shipped by Standard Ground service or Best Way, unless other means are specified or selected. HAVE reserves the right to choose shipping method.&nbsp;&nbsp;For our complete Policy Page –<span style='color:black'>&nbsp;</span></span><a href='http://store.haveinc.com/t-policymain.aspx' target='_blank'><span style='font-size:7.0pt;color:blue'>http://store.haveinc.com/t-policymain.aspx</span></a><span style='font-size:7.0pt;color:black'></span><span style='font-size:7.0pt;'> </span></p> </div> </div></body></html></body></html>"
    
SignOff2 = "<p style='margin:0in;margin-bottom:.0001pt;line-height:14.65pt;'><p style='line-height:14.65pt;background:white'><span> </span></p> <p><span>Once again, thank you for ordering from HAVE Inc. We do appreciate your business!</span> </p> </div></body></html>" & _
    "<div> <p><span >Regards,</span></p> <p class=MsoNormal style='line-height:14.65pt;background:white'><span></span></p><div><font color='#274e13' size='4' face='garamond, serif'>Darcie A. Unson</font><br><font style='font-size:13px' color='#351c75' face='garamond, serif'>Customer Service Specialist</font><font face='garamond, serif'><font color='#351c75'><br>" & _
    "<a style='font-size:13px' href='http://store.haveinc.com/default.aspx' target='_blank'>HAVE, Inc</a> (Hudson Audio Video Enterprises, Inc.) </font></font></div><div><font face='garamond, serif'><font color='#351c75'>Certified NYS WBE | Est. 1977</font><font color='#000000'>  </font><font color='#351c75'><br>P: 800-999-4283 ext 251 ~ F: 518-828-2008 <br>E: <a style='font-size:13px' href='mailto:dunson@haveinc.com' target='_blank'>dunson@haveinc.com</a></font><br></font></div>" & _
    "<div><font face='garamond, serif'><b><div style='display:inline'><div style='display:inline'><em style='font-size:x-small;font-weight:normal;color:rgb(0,0,0)'>   </em></div></div></b></font><b><span style='font-size:7.5pt;'>Shipping Policy:</span></b><span style='font-size:7.5pt;'>&nbsp;All orders are shipped F.O.B. primarily from our headquarters location in Hudson, NY 12534. Orders are shipped by Standard Ground service or Best Way, unless other means are specified or selected. HAVE reserves the right to choose shipping method.&nbsp;&nbsp;For our complete Policy Page –<span style='color:black'>&nbsp;</span></span><a href='http://store.haveinc.com/t-policymain.aspx' target='_blank'><span style='font-size:7.0pt;color:blue'>http://store.haveinc.com/t-policymain.aspx</span></a><span style='font-size:7.0pt;color:black'></span><span style='font-size:7.0pt;'> </span></p> </div> </div></body></html></body></html>"

PartialOrderTable = "<br><p style='margin:0in;margin-bottom:.0001pt;line-height:14.65pt;'> <b><span>THE FOLLOWING ITEM(S) ARE BACK ORDERED:</span></b><span> </span></p> <table id='backordertable' border=1 cellspacing=3 cellpadding=0 width=0 style='width:512.6pt;border:outset #F2F2F2 1.0pt'> <tr style='height:20.65pt'> <td width=115 valign=bottom style='width:83.15pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:20.65pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Part ID</span></b> </p> </td> <td width=317 valign=bottom style='width:235.45pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:20.65pt;overflow:hidden'> <p align=center style='text-align:center'>" & _
    "<b><span>Description </span></b></p> </td> <td width=78 valign=bottom style='width:56.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:20.65pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Qty</span></b></p> </td> <td width=174 valign=bottom style='width:127.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:20.65pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Estimated Ship Date </span></b></p> </td> </tr> <tr style='height:21.5pt'> <td width=115 valign=bottom style='width:83.15pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:21.5pt'> <p><span>P1 </span></p> </td> <td width=317 valign=bottom style='width:235.45pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:21.5pt;overflow:hidden'> <p><span>P2</span></p> </td>" & _
    "<td width=78 valign=bottom style='width:56.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:21.5pt;overflow:hidden'> <p><span>P3 </span></p> </td> <td width=174 valign=bottom style='width:127.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:21.5pt;overflow:hidden'> <p><span>P4 </span></p> </td> </tr> </table> <span> &nbsp; </p> <p style='line-height:14.65pt;background:white'><span> </span></p>"
    
PartialOrderTable2 = "<br><p style='margin:0in;margin-bottom:.0001pt;line-height:14.65pt;'> <b><span>THE FOLLOWING ITEM(S) ARE BACK ORDERED:</span></b><span> </span></p> <table id='backordertable' border=1 cellspacing=3 cellpadding=0 width=0 style='width:512.6pt;border:outset #F2F2F2 1.0pt'> <tr style='height:20.65pt'> <td width=115 valign=bottom style='width:83.15pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:20.65pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Part ID</span></b> </p> </td> <td width=317 valign=bottom style='width:235.45pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:20.65pt;overflow:hidden'> <p align=center style='text-align:center'>" & _
    "<b><span>Description </span></b></p> </td> <td width=78 valign=bottom style='width:56.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:20.65pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Qty</span></b></p> </td> <td width=174 valign=bottom style='width:127.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:20.65pt;overflow:hidden'> <p align=center style='text-align:center'><b><span>Estimated Ship Date </span></b></p> </td> </tr> <tr style='height:21.5pt'> <td width=115 valign=bottom style='width:83.15pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:21.5pt'> <p><span>P1 </span></p> </td> <td width=317 valign=bottom style='width:235.45pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:21.5pt;overflow:hidden'> <p><span>P2</span></p> </td>" & _
    "<td width=78 valign=bottom style='width:56.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:21.5pt;overflow:hidden'> <p><span>P3 </span></p> </td> <td width=174 valign=bottom style='width:127.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:21.5pt;overflow:hidden'> <p><span>P4 </span></p> "
    
ExtraRows = "<tr style='height:21.5pt'> <td width=115 valign=bottom style='width:83.15pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:21.5pt'> <p><span>P1 </span></p> </td> <td width=317 valign=bottom style='width:235.45pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:21.5pt;overflow:hidden'> <p><span>P2</span></p> </td>" & _
    "<td width=78 valign=bottom style='width:56.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:21.5pt;overflow:hidden'> <p><span>P3 </span></p> </td> <td width=174 valign=bottom style='width:127.5pt;border:inset #F2F2F2 1.0pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;height:21.5pt;overflow:hidden'> <p><span>P4 </span></p> </td> </tr>"
    
EndTable = "</table> <span> &nbsp; </p> <p style='line-height:14.65pt;background:white'><span> </span></p> </td> </tr>"
    
With objMail
    .BodyFormat = olFormatHTML
    .HTMLBody = MainMessage & SignOff
    .Display
    .To = "custemailvar"
    .Attachments.Add Pth
    .subject = "Order ordervar Shipping Confirmation, from HAVE, INC"
    .CC = "scasimpson@gmail.com"
End With
 
    

If Time > TimeValue("4:00 PM") Then
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "datevar", Date)
    Else
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "datevar ", DateAdd("d", -1, Date))
End If
    
    
sReturnEmail = InputBox("Please enter the customer's email address:")
        If sReturnEmail = vbNullString Then
            Exit Sub
            Else
        End If
    emailvar = objMail.To
    emailvar = Replace(emailvar, "custemailvar", sReturnEmail)
    objMail.To = emailvar
    
    
    
    Dim subject As String
    Dim sReturn1 As String
    Dim i As Long
    Dim j As Integer
    Dim k As Integer
    Dim Answer3 As String
    Dim MyNote3 As String
    
    MyNote = "Is this a Partial Order?"
    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "Partial?")
        If Answer = vbYes Then
            With objMail
             .subject = "Partial Order ordervar Shipping Confirmation, from HAVE, INC"
            End With
    
Dim sReturnOR As String
Dim TheNum As Integer
Dim valid As Boolean: valid = True
        Do
Start:
    sReturnOR = InputBox("How many line items are back ordered? Please enter an integer - 10 or less")
        If IsNumeric(sReturnOR) Then
            TheNum = CInt(sReturnOR)
            valid = True
                Else
                MsgBox "Invalid Input, please enter a number"
                valid = False
                    GoTo Start
        End If
            If sReturnOR > 10 Then
                MsgBox "That's too many line items for this program!"
                valid = False
            End If
        
        Loop Until valid = True

     
     
    
    
    If sReturnOR > 1 Then
    ReDim fullPath(i To sReturnOR) As Variant
    For i = 2 To sReturnOR
    fullPath(i) = ExtraRows
        fullPath(i) = Replace(fullPath(i), "P1", "F" & i + 4)
        fullPath(i) = Replace(fullPath(i), "P2", "D" & i + 4)
        fullPath(i) = Replace(fullPath(i), "P3", "Q" & i + 4)
        fullPath(i) = Replace(fullPath(i), "P4", "E" & i + 4)
    Next i
    
    objMail.HTMLBody = MainMessagep & PartialOrderTable2 & Join(fullPath) & EndTable & SignOff
    Else: objMail.HTMLBody = MainMessagep & PartialOrderTable & SignOff

    
    sReturn9 = InputBox("Please enter Back Ordered Item #1 Product ID")
    If sReturn9 = vbNullString Then
        Exit Sub
        Else
        End If
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P1", sReturn9)
    sReturn10 = InputBox("Please enter Description")
     If sReturn10 = vbNullString Then
        Exit Sub
        Else
        End If
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P2", sReturn10)
    sReturn11 = InputBox("Please enter Quantity")
     If sReturn11 = vbNullString Then
        Exit Sub
        Else
        End If
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P3", sReturn11)
    sReturn12 = InputBox("Please enter Estimated Ship Date")
     If sReturn12 = vbNullString Then
        Exit Sub
        Else
        End If
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P4", sReturn12)
    
    End If
    
   If sReturnOR = 2 Then
  
   sReturn9 = InputBox("Please enter Back Ordered Item #1 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P1", sReturn9)
    sReturn10 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P2", sReturn10)
    sReturn11 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P3", sReturn11)
    sReturn12 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P4", sReturn12)
    
    sReturn13 = InputBox("Please enter Back Ordered Item #2 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F6", sReturn13)
    sReturn14 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D6", sReturn14)
    sReturn15 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q6", sReturn15)
    sReturn16 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E6", sReturn16)
    
    ElseIf sReturnOR = 3 Then
    sReturn9 = InputBox("Please enter Back Ordered Item #1 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P1", sReturn9)
    sReturn10 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P2", sReturn10)
    sReturn11 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P3", sReturn11)
    sReturn12 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P4", sReturn12)
    
    sReturn13 = InputBox("Please enter Back Ordered Item #2 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F6", sReturn13)
    sReturn14 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D6", sReturn14)
    sReturn15 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q6", sReturn15)
    sReturn16 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E6", sReturn16)
    
    sReturn17 = InputBox("Please enter Back Ordered Item #3 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F7", sReturn17)
    sReturn18 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D7", sReturn18)
    sReturn19 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q7", sReturn19)
    sReturn20 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E7", sReturn20)
    
    ElseIf sReturnOR = 4 Then
      sReturn9 = InputBox("Please enter Back Ordered Item #1 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P1", sReturn9)
    sReturn10 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P2", sReturn10)
    sReturn11 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P3", sReturn11)
    sReturn12 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P4", sReturn12)
    
    sReturn13 = InputBox("Please enter Back Ordered Item #2 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F6", sReturn13)
    sReturn14 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D6", sReturn14)
    sReturn15 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q6", sReturn15)
    sReturn16 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E6", sReturn16)
    
    sReturn17 = InputBox("Please enter Back Ordered Item #3 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F7", sReturn17)
    sReturn18 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D7", sReturn18)
    sReturn19 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q7", sReturn19)
    sReturn20 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E7", sReturn20)
    
    sReturn21 = InputBox("Please enter Back Ordered Item #4 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F8", sReturn21)
    sReturn22 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D8", sReturn22)
    sReturn23 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q8", sReturn23)
    sReturn24 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E8", sReturn24)
    
    
    ElseIf sReturnOR = 5 Then
      sReturn9 = InputBox("Please enter Back Ordered Item #1 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P1", sReturn9)
    sReturn10 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P2", sReturn10)
    sReturn11 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P3", sReturn11)
    sReturn12 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P4", sReturn12)
    
    sReturn13 = InputBox("Please enter Back Ordered Item #2 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F6", sReturn13)
    sReturn14 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D6", sReturn14)
    sReturn15 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q6", sReturn15)
    sReturn16 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E6", sReturn16)
    
    sReturn17 = InputBox("Please enter Back Ordered Item #3 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F7", sReturn17)
    sReturn18 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D7", sReturn18)
    sReturn19 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q7", sReturn19)
    sReturn20 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E7", sReturn20)
    
    sReturn21 = InputBox("Please enter Back Ordered Item #4 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F8", sReturn21)
    sReturn22 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D8", sReturn22)
    sReturn23 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q8", sReturn23)
    sReturn24 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E8", sReturn24)
    
    sReturn25 = InputBox("Please enter Back Ordered Item #5 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F9", sReturn25)
    sReturn26 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D9", sReturn26)
    sReturn27 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q9", sReturn27)
    sReturn28 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E9", sReturn28)
    
    ElseIf sReturnOR = 6 Then
    
    sReturn9 = InputBox("Please enter Back Ordered Item #1 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P1", sReturn9)
    sReturn10 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P2", sReturn10)
    sReturn11 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P3", sReturn11)
    sReturn12 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P4", sReturn12)
    
    sReturn13 = InputBox("Please enter Back Ordered Item #2 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F6", sReturn13)
    sReturn14 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D6", sReturn14)
    sReturn15 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q6", sReturn15)
    sReturn16 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E6", sReturn16)
    
    sReturn17 = InputBox("Please enter Back Ordered Item #3 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F7", sReturn17)
    sReturn18 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D7", sReturn18)
    sReturn19 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q7", sReturn19)
    sReturn20 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E7", sReturn20)
    
    sReturn21 = InputBox("Please enter Back Ordered Item #4 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F8", sReturn21)
    sReturn22 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D8", sReturn22)
    sReturn23 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q8", sReturn23)
    sReturn24 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E8", sReturn24)
    
    sReturn25 = InputBox("Please enter Back Ordered Item #5 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F9", sReturn25)
    sReturn26 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D9", sReturn26)
    sReturn27 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q9", sReturn27)
    sReturn28 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E9", sReturn28)
    
    sReturn29 = InputBox("Please enter Back Ordered Item #6 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F10", sReturn29)
    sReturn30 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D10", sReturn30)
    sReturn31 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q10", sReturn31)
    sReturn32 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E10", sReturn32)
    
    
    ElseIf sReturnOR = 7 Then
    
    sReturn9 = InputBox("Please enter Back Ordered Item #1 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P1", sReturn9)
    sReturn10 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P2", sReturn10)
    sReturn11 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P3", sReturn11)
    sReturn12 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P4", sReturn12)
    
    sReturn13 = InputBox("Please enter Back Ordered Item #2 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F6", sReturn13)
    sReturn14 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D6", sReturn14)
    sReturn15 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q6", sReturn15)
    sReturn16 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E6", sReturn16)
    
    sReturn17 = InputBox("Please enter Back Ordered Item #3 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F7", sReturn17)
    sReturn18 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D7", sReturn18)
    sReturn19 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q7", sReturn19)
    sReturn20 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E7", sReturn20)
    
    sReturn21 = InputBox("Please enter Back Ordered Item #4 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F8", sReturn21)
    sReturn22 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D8", sReturn22)
    sReturn23 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q8", sReturn23)
    sReturn24 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E8", sReturn24)
    
    sReturn25 = InputBox("Please enter Back Ordered Item #5 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F9", sReturn25)
    sReturn26 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D9", sReturn26)
    sReturn27 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q9", sReturn27)
    sReturn28 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E9", sReturn28)
    
    sReturn29 = InputBox("Please enter Back Ordered Item #6 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F10", sReturn29)
    sReturn30 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D10", sReturn30)
    sReturn31 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q10", sReturn31)
    sReturn32 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E10", sReturn32)
    
    sReturn33 = InputBox("Please enter Back Ordered Item #7 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F11", sReturn33)
    sReturn34 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D11", sReturn34)
    sReturn35 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q11", sReturn35)
    sReturn36 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E11", sReturn36)
    
    ElseIf sReturnOR = 8 Then
    
    sReturn9 = InputBox("Please enter Back Ordered Item #1 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P1", sReturn9)
    sReturn10 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P2", sReturn10)
    sReturn11 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P3", sReturn11)
    sReturn12 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P4", sReturn12)
    
    sReturn13 = InputBox("Please enter Back Ordered Item #2 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F6", sReturn13)
    sReturn14 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D6", sReturn14)
    sReturn15 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q6", sReturn15)
    sReturn16 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E6", sReturn16)
    
    sReturn17 = InputBox("Please enter Back Ordered Item #3 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F7", sReturn17)
    sReturn18 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D7", sReturn18)
    sReturn19 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q7", sReturn19)
    sReturn20 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E7", sReturn20)
    
    sReturn21 = InputBox("Please enter Back Ordered Item #4 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F8", sReturn21)
    sReturn22 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D8", sReturn22)
    sReturn23 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q8", sReturn23)
    sReturn24 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E8", sReturn24)
    
    sReturn25 = InputBox("Please enter Back Ordered Item #5 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F9", sReturn25)
    sReturn26 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D9", sReturn26)
    sReturn27 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q9", sReturn27)
    sReturn28 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E9", sReturn28)
    
    sReturn29 = InputBox("Please enter Back Ordered Item #6 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F10", sReturn29)
    sReturn30 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D10", sReturn30)
    sReturn31 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q10", sReturn31)
    sReturn32 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E10", sReturn32)
    
    sReturn33 = InputBox("Please enter Back Ordered Item #7 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F11", sReturn33)
    sReturn34 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D11", sReturn34)
    sReturn35 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q11", sReturn35)
    sReturn36 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E11", sReturn36)
    
     sReturn37 = InputBox("Please enter Back Ordered Item #8 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F12", sReturn37)
    sReturn38 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D12", sReturn38)
    sReturn39 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q12", sReturn39)
    sReturn40 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E12", sReturn40)
    
    ElseIf sReturnOR = 9 Then
    sReturn9 = InputBox("Please enter Back Ordered Item #1 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P1", sReturn9)
    sReturn10 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P2", sReturn10)
    sReturn11 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P3", sReturn11)
    sReturn12 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P4", sReturn12)
    
    sReturn13 = InputBox("Please enter Back Ordered Item #2 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F6", sReturn13)
    sReturn14 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D6", sReturn14)
    sReturn15 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q6", sReturn15)
    sReturn16 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E6", sReturn16)
    
    sReturn17 = InputBox("Please enter Back Ordered Item #3 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F7", sReturn17)
    sReturn18 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D7", sReturn18)
    sReturn19 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q7", sReturn19)
    sReturn20 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E7", sReturn20)
    
    sReturn21 = InputBox("Please enter Back Ordered Item #4 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F8", sReturn21)
    sReturn22 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D8", sReturn22)
    sReturn23 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q8", sReturn23)
    sReturn24 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E8", sReturn24)
    
    sReturn25 = InputBox("Please enter Back Ordered Item #5 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F9", sReturn25)
    sReturn26 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D9", sReturn26)
    sReturn27 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q9", sReturn27)
    sReturn28 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E9", sReturn28)
    
    sReturn29 = InputBox("Please enter Back Ordered Item #6 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F10", sReturn29)
    sReturn30 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D10", sReturn30)
    sReturn31 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q10", sReturn31)
    sReturn32 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E10", sReturn32)
    
    sReturn33 = InputBox("Please enter Back Ordered Item #7 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F11", sReturn33)
    sReturn34 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D11", sReturn34)
    sReturn35 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q11", sReturn35)
    sReturn36 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E11", sReturn36)
    
     sReturn37 = InputBox("Please enter Back Ordered Item #8 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F12", sReturn37)
    sReturn38 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D12", sReturn38)
    sReturn39 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q12", sReturn39)
    sReturn40 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E12", sReturn40)
    
    sReturn41 = InputBox("Please enter Back Ordered Item #9 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F13", sReturn41)
    sReturn42 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D13", sReturn42)
    sReturn43 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q13", sReturn43)
    sReturn44 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E13", sReturn44)
   
   ElseIf sReturnOR = 10 Then
    sReturn9 = InputBox("Please enter Back Ordered Item #1 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P1", sReturn9)
    sReturn10 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P2", sReturn10)
    sReturn11 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P3", sReturn11)
    sReturn12 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "P4", sReturn12)
    
    sReturn13 = InputBox("Please enter Back Ordered Item #2 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F6", sReturn13)
    sReturn14 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D6", sReturn14)
    sReturn15 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q6", sReturn15)
    sReturn16 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E6", sReturn16)
    
    sReturn17 = InputBox("Please enter Back Ordered Item #3 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F7", sReturn17)
    sReturn18 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D7", sReturn18)
    sReturn19 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q7", sReturn19)
    sReturn20 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E7", sReturn20)
    
    sReturn21 = InputBox("Please enter Back Ordered Item #4 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F8", sReturn21)
    sReturn22 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D8", sReturn22)
    sReturn23 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q8", sReturn23)
    sReturn24 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E8", sReturn24)
    
    sReturn25 = InputBox("Please enter Back Ordered Item #5 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F9", sReturn25)
    sReturn26 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D9", sReturn26)
    sReturn27 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q9", sReturn27)
    sReturn28 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E9", sReturn28)
    
    sReturn29 = InputBox("Please enter Back Ordered Item #6 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F10", sReturn29)
    sReturn30 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D10", sReturn30)
    sReturn31 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q10", sReturn31)
    sReturn32 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E10", sReturn32)
    
    sReturn33 = InputBox("Please enter Back Ordered Item #7 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F11", sReturn33)
    sReturn34 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D11", sReturn34)
    sReturn35 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q11", sReturn35)
    sReturn36 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E11", sReturn36)
    
     sReturn37 = InputBox("Please enter Back Ordered Item #8 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F12", sReturn37)
    sReturn38 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D12", sReturn38)
    sReturn39 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q12", sReturn39)
    sReturn40 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E12", sReturn40)
    
    sReturn41 = InputBox("Please enter Back Ordered Item #9 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F13", sReturn41)
    sReturn42 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D13", sReturn42)
    sReturn43 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q13", sReturn43)
    sReturn44 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E13", sReturn44)
    
    sReturn45 = InputBox("Please enter Back Ordered Item #10 Product ID")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "F14", sReturn45)
    sReturn46 = InputBox("Please enter Description")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "D14", sReturn46)
    sReturn47 = InputBox("Please enter Quantity")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "Q14", sReturn47)
    sReturn48 = InputBox("Please enter Estimated Ship Date")
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "E14", sReturn48)
    
    
    End If
    End If
 
    If Time > TimeValue("4:00 PM") Then
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "datevar", Date)
    Else
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "datevar ", DateAdd("d", -1, Date))
End If

    
sReturn45 = InputBox("Who is the SALES REP? Enter John, Lowell, Gary, or Have!")
If sReturn45 = vbNullString Then
Exit Sub
ElseIf sReturn45 = "John" Or sReturn45 = "john" Or sReturn45 = "JOHN" Then
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "salesrepvar", "John Earl")
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "extvar ", "x 235")
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "emailvar", "<a href='mailto:jearl@haveinc.com'>" & "jearl@haveinc.com" & "</a>")
ElseIf sReturn45 = "Gary" Or sReturn45 = "gary" Or sReturn45 = "GARY" Then
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "salesrepvar", "Gary Purnhagen")
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "extvar ", "x 231")
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "emailvar", "<a href='mailto:gpurnhagen@haveinc.com'>" & "gpurnhagen@haveinc.com" & "</a>")
ElseIf sReturn45 = "Lowell" Or sReturn45 = "lowell" Or sReturn45 = "LOWELL" Then
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "salesrepvar", "Lowell Stringer")
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "extvar ", "x 247")
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "emailvar", "<a href='mailto:lstringer@haveinc.com'>" & "lstringer@haveinc.com" & "</a>")
ElseIf sReturn45 = "Have" Or sReturn45 = "have" Or sReturn45 = "HAVE" Then
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "salesrepvar", "HAVE, INC")
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "extvar ", "")
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "emailvar", "<a href='mailto:pro_sales@haveinc.com'>" & "pro_sales@haveinc.com" & "</a>")
End If

UserForm2.Show
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "namevar", UserForm2.TextBox1.Value)
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "webordervar", UserForm2.TextBox2.Value)
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "POvar", UserForm2.TextBox3.Value)
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "paclistvar", UserForm2.TextBox4.Value)
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "ordervar", UserForm2.TextBox5.Value)
subject = objMail.subject
subject = Replace(subject, "ordervar", UserForm2.TextBox5.Value)
objMail.subject = subject

Unload UserForm2
Set UserForm2 = Nothing

Dim Answer1 As String
Dim MyNote1 As String
    MyNote = "Is the FOB = Hudson? AND Is it standard ground?"
    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "FOB and Ship Method")
    If Answer = vbNo Then
        sReturn9 = InputBox("Enter FedEx Service or Account : ")
        If sReturn9 = vbNullString Then
        Exit Sub
        Else
        End If
        Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "shipviavar", sReturn9)

    Else
    End If

Dim Answer2 As String
Dim MyNote2 As String
Dim tTime As Integer
Dim ShipDate As Date
Dim sReturn6 As String
Dim TheDate As Date
Dim valid2 As Boolean: valid2 = True
    
    MyNote = "Is this a FedEx Tracking Number?"
    
    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "Ship Via")
    
    If Answer = vbNo Then
    
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "shipviavar", "USPS Priority")
    sReturn7 = InputBox("Please enter the USPS Tracking Number : ")
    If sReturn7 = vbNullString Then
    Exit Sub
    Else
    End If
        Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "trackingnumbervar", "<a href='https://tools.usps.com/go/TrackConfirmAction?tLabels=" & sReturn7 & "'>" & sReturn7 & "</a>")
    
    Do
    sReturn6 = InputBox("Please enter the estimated arrival date:")
    If IsDate(sReturn6) Then
    TheDate = DateValue(sReturn6)
    valid2 = True
    Else
    MsgBox "Invalid Date"
    valid2 = False
    End If
    Loop Until valid2 = True
    If Time > TimeValue("4:00 PM") Then
    ShipDate = Date
    Else
    ShipDate = DateAdd("d", -1, Date)
    End If
    tTime = NetWorkdays2(ShipDate, TheDate, 1)
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "transittimevar business", tTime)
        
        
        
    Else
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "shipviavar", "FedEx Ground")
    sReturn8 = InputBox("Please enter the FedEx Tracking Number : ")
     If sReturn8 = vbNullString Then
    Exit Sub
    Else
    End If
        Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "trackingnumbervar", "<a href='https://www.fedex.com/apps/fedextrack/?tracknumbers=" & sReturn8 & "'>" & sReturn8 & "</a>")
    Do
    sReturn6 = InputBox("Please enter the estimated arrival date:")
    If IsDate(sReturn6) Then
    TheDate = DateValue(sReturn6)
    valid2 = True
    Else
    MsgBox "Invalid Date"
    valid2 = False
    End If
    Loop Until valid2 = True
    If Time > TimeValue("4:00 PM") Then
    ShipDate = Date
    Else
    ShipDate = DateAdd("d", -1, Date)
    End If
    tTime = NetWorkdays2(ShipDate, TheDate, 65)
    Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "transittimevar", tTime)
    
    
End If





Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "1 business days", "1 business day")


UserForm1.Show

UserForm1.TextBox1.Value = Replace(UserForm1.TextBox1.Text, vbCr & vbLf, "<br>")
Application.ActiveInspector.CurrentItem.HTMLBody = Replace(Application.ActiveInspector.CurrentItem.HTMLBody, "CustomerAddress", UserForm1.TextBox1.Value)

Unload UserForm1
Set UserForm1 = Nothing

    
    
End Sub



