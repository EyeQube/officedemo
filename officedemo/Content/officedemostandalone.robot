➜filesearch
timeout.reset value 99999999
file.exists filename ‴♥appdata\Bitoreq AB\OfficeDemo\salesreport.xml‴ timeout 99999999 errorjump ➜filesearch
delay milliseconds 100

♥starttime = (time)System.DateTime.Now

-************************************************************
file.copy path ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\process.txt‴ destinationpath ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\process_start.txt‴

file.copy path ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\process.txt‴ destinationpath ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\one.txt‴

delay seconds 2
ie.attach phrase ‴Office Bot 3000‴
ie.refresh
ie.detach
delay seconds 1
-************************************************************

text.read filename ‴♥appdata\Bitoreq AB\OfficeDemo\salesreport.xml‴ result xmlString
xml text ♥xmlString search ‴Name‴ result reportrow1
♥reportrow2 = ⊂(DateTime.Now.ToShortDateString())⊃
xml text ♥xmlString search ‴Sale‴ result reportrow3
xml text ♥xmlString search ‴Costs‴ result reportrow4
xml text ♥xmlString search ‴Order‴ result reportrow5
xml text ♥xmlString search ‴Month‴ result reportrow6
♥reportrow6 = ⊂♥reportrow6.ToLower()⊃
xml text ♥xmlString search ‴E-mail‴ result reportrow7

♥receiver = ♥reportrow7
file.delete ‴♥appdata\Bitoreq AB\OfficeDemo\salesreport.xml‴

excel.open path ‴♥appdata\Bitoreq AB\OfficeDemo\återförsäljarrapportmall.xlsx‴
window title ‴✱excel✱‴ style maximize
excel.setvalue row 4 col 2 value ♥reportrow1
excel.setvalue row 4 col 3 value ♥reportrow2
excel.setvalue row 4 col 4 value ♥reportrow3
excel.setvalue row 4 col 5 value ♥reportrow4
excel.setvalue row 4 col 6 value ♥reportrow5
excel.setvalue row 4 col 7 value ♥reportrow6
excel.getrow row 4 result reportrow
♥filnamn = ⊂(♥reportrow⟦b⟧ + "_" + ♥reportrow⟦g⟧).Replace(" ", "")⊃

-Check if the reseller already has reported for this month
file.exists filename ‴♥appdata\Bitoreq AB\OfficeDemo\done\♥filnamn.xlsx‴ timeout 10000 errorjump ➜reportdoesnotexist
♥resellerreportexists = true
file.delete filename ‴♥appdata\Bitoreq AB\OfficeDemo\done\♥filnamn.xlsx‴
jump ➜filesearchdone
➜reportdoesnotexist
♥resellerreportexists = false
➜filesearchdone

jump ➜januari if ⊂♥reportrow⟦g⟧ == "januari"⊃
jump ➜februari if ⊂♥reportrow⟦g⟧ == "februari"⊃
jump ➜march if ⊂♥reportrow⟦g⟧ == "mars"⊃
jump ➜april if ⊂♥reportrow⟦g⟧ == "april"⊃
jump ➜may if ⊂♥reportrow⟦g⟧ == "maj"⊃
jump ➜june if ⊂♥reportrow⟦g⟧ == "juni"⊃
jump ➜july if ⊂♥reportrow⟦g⟧ == "juli"⊃
jump ➜august if ⊂♥reportrow⟦g⟧ == "augusti"⊃
jump ➜september if ⊂♥reportrow⟦g⟧ == "september"⊃
jump ➜october if ⊂♥reportrow⟦g⟧ == "oktober"⊃
jump ➜november if ⊂♥reportrow⟦g⟧ == "november"⊃
jump ➜december if ⊂♥reportrow⟦g⟧ == "december"⊃
➜january
♥fictivedate = ‴2017-01-25‴
jump ➜dateset
➜february
♥fictivedate = ‴2017-02-25‴
jump ➜dateset
➜march
♥fictivedate = ‴2017-03-25‴
jump ➜dateset
➜april
♥fictivedate = ‴2017-04-25‴
jump ➜dateset
➜may
♥fictivedate = ‴2017-05-25‴
jump ➜dateset
➜june
♥fictivedate = ‴2017-06-25‴
jump ➜dateset
➜july
♥fictivedate = ‴2017-07-25‴
jump ➜dateset
➜august
♥fictivedate = ‴2017-08-25‴
jump ➜dateset
➜september
♥fictivedate = ‴2017-09-25‴
jump ➜dateset
➜october
♥fictivedate = ‴2017-10-25‴
jump ➜dateset
➜november
♥fictivedate = ‴2017-11-25‴
jump ➜dateset
➜december
♥fictivedate = ‴2017-12-25‴
➜dateset

excel.save path ‴♥appdata\Bitoreq AB\OfficeDemo\done\♥filnamn.xlsx‴
excel.close

-***********************************************************************
file.delete ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\one.txt‴

file.copy path ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\process.txt‴ destinationpath ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\two.txt‴

delay seconds 2
ie.attach phrase ‴Office Bot 3000‴
ie.refresh
ie.detach
delay seconds 1
-***********************************************************************

excel.open path ‴♥appdata\Bitoreq AB\OfficeDemo\konsolidering.xlsx‴
window title ‴✱excel✱‴ style maximize
excel.activatesheet name ‴Datainmatning‴

-The lines until ➜endloop check if the reseller already has reported for the month. If yes, then the previous report is overwritten
♥rownumber = 6

jump ➜nextrowtocheck if ⊂♥resellerreportexists⊃
jump ➤insertnewrow
jump ➜endloop

➜nextrowtocheck
excel.getvalue row ♥rownumber col C result checkreseller
jump ➜endloop if ⊂string.IsNullOrEmpty(♥checkreseller)⊃
♥rownumber = ♥rownumber + 1
jump ➜nextrowtocheck if ⊂♥checkreseller != ♥reportrow⟦b⟧⊃
♥rownumber = ♥rownumber - 1
excel.getvalue row ♥rownumber col H result checkmonth
♥rownumber = ♥rownumber + 1

jump ➜nextrowtocheck if ⊂♥checkmonth != ♥reportrow⟦g⟧⊃
♥rownumber = ♥rownumber - 1
excel.setvalue row ♥rownumber col B value ♥fictivedate
excel.setvalue row ♥rownumber col D value ♥reportrow⟦d⟧ 
excel.setvalue row ♥rownumber col E value ♥reportrow⟦f⟧
excel.setvalue row ♥rownumber col F value ♥reportrow⟦e⟧
➜endloop

keyboard ⋘CTRL+ALT+F5⋙
excel.activatesheet ‴Sammanfattning‴

-excel.getrow row 6 result quarter1results
-excel.getrow row 7 result quarter2results
-excel.getrow row 8 result quarter3results
-excel.getrow row 9 result quarter4results
excel.getvalue row 8 col C result quarter3sales
excel.getvalue row 8 col D result quarter3profit
excel.getvalue row 8 col E result quarter3cost
excel.getvalue row 8 col F result quarter3order
excel.getvalue row 9 col C result quarter4sales
excel.getvalue row 9 col D result quarter4profit
excel.getvalue row 9 col E result quarter4cost
excel.getvalue row 9 col F result quarter4order
excel.save
delay milliseconds 300
excel.close

-********************************************************************
delay seconds 1
file.delete ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\two.txt‴

file.copy path ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\process.txt‴ destinationpath ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\three.txt‴

delay seconds 2
ie.attach phrase ‴Office Bot 3000‴
ie.refresh
ie.detach
delay seconds 1
-********************************************************************

delay milliseconds 300
word.open path ‴♥appdata\Bitoreq AB\OfficeDemo\rapportmall.docx‴
window ‴✱word✱‴ style maximize
delay milliseconds 500
keyboard text ⋘CTRL+B⋙
delay milliseconds 100
keyboard text ‴diagram 1‴
delay milliseconds 200
keyboard ⋘ENTER⋙
delay milliseconds 200
keyboard ⋘ALT⋙⋘Ö⋙⋘G⋙
delay milliseconds 500
keyboard ⋘UP⋙
delay milliseconds 100
keyboard ⋘SHIFT+RIGHT⋙
delay milliseconds 200
keyboard ⋘SHIFT+F10⋙
delay milliseconds 100
keyboard ⋘R⋙
delay milliseconds 200
keyboard ⋘RIGHT⋙⋘DOWN⋙
delay milliseconds 200
keyboard ⋘D⋙
delay milliseconds 500

window title ‴✱i Microsoft Word - Excel‴ style maximize
delay milliseconds 300

keyboard ⋘F5⋙
delay milliseconds 100

keyboard ⋘B⋙⋘5⋙⋘ENTER⋙
delay milliseconds 200

keyboard text ♥quarter4sales
keyboard text ⋘RIGHT⋙

keyboard text ♥quarter4profit
keyboard text ⋘RIGHT⋙

keyboard text ♥quarter4cost
keyboard text ⋘RIGHT⋙

keyboard text ♥quarter4order

keyboard text ⋘ENTER⋙
keyboard text ⋘ALT+F4⋙

-**************************************************************
delay seconds 1
file.delete ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\three.txt‴

file.copy path ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\process.txt‴ destinationpath ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\four.txt‴

delay seconds 2
ie.attach phrase ‴Office Bot 3000‴
ie.refresh
ie.detach
delay seconds 1
-**************************************************************

delay milliseconds 300
excel.open path ‴♥appdata\Bitoreq AB\OfficeDemo\konsolidering.xlsx‴
window title ‴✱excel✱‴ style maximize
excel.activatesheet ‴Försäljningsrapport helår‴

-♥resellername = ‴‴❚
♥salesnumber = ‴‴❚‴‴❚‴‴❚‴‴❚‴‴❚‴‴❚‴‴❚‴‴
♥totalsales = 0
♥rowoffset = 5
♥index = 1
♥row = ♥rowoffset + 1
♥numberofresellers = 0
♥highestsalesindex = 1
♥highestsalesint = 0
➜nextreseller2

excel.getvalue row ♥row col D result sales
♥salesnumber⟦♥index⟧ = ♥sales
-♥salesnumber = ♥salesnumber❚‴‴

-♥sales = ⊂♥sales.Remove(♥sales.Length - 3).Replace(" ", "")⊃
♥sales = ⊂♥sales.Replace(" ", "")⊃

♥salesint = ⊂Convert.ToInt32(♥sales)⊃
jump ➜nothighestsales if ⊂♥highestsalesint>♥salesint⊃ 
♥highestsalesindex = ♥index
♥highestsalesint = ♥salesint
➜nothighestsales

♥index = ♥index + 1
♥row = ♥row + 1
timeout.reset 100000
jump ➜nextreseller2 if ⊂♥index<9⊃
➜nextresellerend2
♥highestsalesrow = ♥highestsalesindex + ♥rowoffset
excel.getvalue row ♥row col D result totalsales
excel.getvalue row ♥highestsalesrow col C result highestreseller

-♥totalsales = ⊂♥totalsales.Remove(♥totalsales.Length - 3).Replace(" ", "")⊃
♥totalsales = ⊂♥totalsales.Replace(" ", "")⊃

♥totalsalesint = ⊂Convert.ToInt32(♥totalsales)⊃
excel.save path ‴♥appdata\Bitoreq AB\OfficeDemo\done\konsolidering.xlsx‴
excel.close
delay milliseconds 300

-**************************************************************
delay seconds 1
file.delete ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\four.txt‴

file.copy path ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\process.txt‴ destinationpath ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\five.txt‴
delay seconds 2
ie.attach phrase ‴Office Bot 3000‴
ie.refresh
ie.detach
delay seconds 1
-**************************************************************

file.delete filename ‴♥appdata\Bitoreq AB\OfficeDemo\konsolidering.xlsx‴
delay milliseconds 300
file.copy path ‴♥appdata\Bitoreq AB\OfficeDemo\done\konsolidering.xlsx‴ destinationpath ‴♥appdata\Bitoreq AB\OfficeDemo\konsolidering.xlsx‴
delay milliseconds 300

window ‴✱rapport✱‴ style maximize
delay milliseconds 500
keyboard text ⋘CTRL+B⋙
delay milliseconds 100
keyboard text ‴Diagram 2‴
delay milliseconds 100
keyboard ⋘ENTER⋙
delay milliseconds 200
keyboard ⋘ALT⋙⋘Ö⋙⋘G⋙
delay milliseconds 500
keyboard ⋘UP⋙
delay milliseconds 100
keyboard ⋘SHIFT+RIGHT⋙
delay milliseconds 200
keyboard ⋘SHIFT+F10⋙
delay milliseconds 200
keyboard ⋘R⋙
delay milliseconds 200
keyboard ⋘RIGHT⋙⋘DOWN⋙
delay milliseconds 200
keyboard ⋘D⋙
delay seconds 1

-****************************************************************
delay seconds 1
file.delete ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\five.txt‴

file.copy path ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\process.txt‴ destinationpath ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\six.txt‴
delay seconds 2
ie.attach phrase ‴Office Bot 3000‴
ie.refresh
ie.detach
delay seconds 1
-****************************************************************

window title ‴✱i Microsoft Word - Excel‴ style maximize
delay milliseconds 300

keyboard ⋘F5⋙
delay milliseconds 200
keyboard ⋘B⋙⋘2⋙⋘ENTER⋙
delay milliseconds 200

♥salesnumberint = 1❚1❚1❚1❚1❚1❚1❚1
♥index = 1
➜nextrow2
-jump ➜nextrowend2 if ⊂string.IsNullOrEmpty(♥salesnumber⟦♥index⟧)⊃
♥salesnumber⟦♥index⟧ = ⊂♥salesnumber⟦♥index⟧.Replace(" ", "")⊃
-♥salesnumber⟦♥index⟧ = ⊂♥salesnumber⟦♥index⟧.Remove(♥salesnumber⟦♥index⟧.Length - 3).Replace(" ", "")⊃
♥salesnumberint⟦♥index⟧ = ⊂Convert.ToInt32(♥salesnumber⟦♥index⟧)*100/♥totalsalesint⊃
keyboard text ♥salesnumberint⟦♥index⟧
keyboard text ⋘DOWN⋙

♥index = ♥index + 1
timeout.reset 100000
jump ➜nextrow2 if ⊂♥index<9⊃
-➜nextrowend2
word.save path ‴♥appdata\Bitoreq AB\OfficeDemo\rapportmall.docx‴

keyboard text ⋘ENTER⋙
keyboard text ⋘ALT+F4⋙
window title ‴✱rapportmall - Word‴

-******************************************************************
delay seconds 1
file.delete ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\six.txt‴
file.copy path ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\process.txt‴ destinationpath ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\seven.txt‴
delay seconds 2
ie.attach phrase ‴Office Bot 3000‴
ie.refresh
ie.detach
delay seconds 1
-******************************************************************

keyboard text ⋘CTRL+B⋙
delay milliseconds 100
keyboard text ‴kvartalsöversikt‴
delay milliseconds 100
keyboard text ⋘ENTER⋙
delay milliseconds 100
keyboard text ⋘ENTER⋙
delay seconds 1

jump ➜increasedquarterlysales if ⊂Convert.ToInt32(♥quarter4sales)>Convert.ToInt32(♥quarter3sales)⊃
jump ➜nextstatement1
➜increasedquarterlysales
♥salesthisquarter = ⊂Convert.ToInt32(♥quarter4sales)⊃
♥salespreviousquarter = ⊂Convert.ToInt32(♥quarter3sales)⊃

♥salesincreaseint = ⊂((♥salesthisquarter-♥salespreviousquarter)*100/♥salespreviousquarter)⊃
word.replace from ‴[FÖRSÄLJNINGSUTVECKLING]‴ to ‴⊂"Försäljningen har utvecklats positivt under det fjärde kvartalet. Försäljningstillväxten är " + ♥salesincreaseint + "%+ jämfört med föregående kvartal."⊃‴
♥searchitem = ⊂♥salesincreaseint.ToString()+"%+"⊃
jump ➤addcolourtonumber searchitem ♥searchitem colour ‴blue‴
➜nextstatement1

jump ➜decreasedquarterlysales if ⊂Convert.ToInt32(♥quarter4sales)<Convert.ToInt32(♥quarter3sales)⊃
jump ➜nextstatement2
➜decreasedquarterlysales
♥salesthisquarter = ⊂Convert.ToInt32(♥quarter4sales)⊃
♥salespreviousquarter = ⊂Convert.ToInt32(♥quarter3sales)⊃

♥salesdecreaseint = ⊂((♥salespreviousquarter-♥salesthisquarter)*100/♥salespreviousquarter)⊃
word.replace from ‴[FÖRSÄLJNINGSUTVECKLING]‴ to ‴⊂"Försäljningen har utvecklats negativt under det fjärde kvartalet. Försäljningsnedgången är " + ♥salesdecreaseint + "%- jämfört med föregående kvartal."⊃‴
♥searchitem = ⊂♥salesdecreaseint.ToString()+"%-"⊃
jump ➤addcolourtonumber searchitem ♥searchitem colour ‴red‴
➜nextstatement2

jump ➜increasedquarterlyorder if ⊂Convert.ToInt32(♥quarter4order)>Convert.ToInt32(♥quarter3order)⊃
jump ➜nextstatement3
➜increasedquarterlyorder
♥orderthisquarter = ⊂Convert.ToInt32(♥quarter4order)⊃
♥orderpreviousquarter = ⊂Convert.ToInt32(♥quarter3order)⊃

♥orderincreaseint = ⊂((♥orderthisquarter-♥orderpreviousquarter)*100/♥orderpreviousquarter)⊃
word.replace from ‴[ORDERUTVECKLING]‴ to ‴⊂"Orderläget är gott för det fjärde kvartalet. Ordervolymen är upp " + ♥orderincreaseint + "%+ jämfört med föregående kvartal, vilket borgar för en stark utveckling av försäljningen under första delen av 2018."⊃‴
jump ➤addcolourtonumber searchitem ♥searchitem colour ‴blue‴
➜nextstatement3

jump ➜decreasedquarterlyorder if ⊂Convert.ToInt32(♥quarter4order)<Convert.ToInt32(♥quarter3order)⊃
jump ➜nextstatement4
➜decreasedquarterlyorder
♥orderthisquarter = ⊂Convert.ToInt32(♥quarter4order)⊃
♥orderpreviousquarter = ⊂Convert.ToInt32(♥quarter3order)⊃

♥orderdecreaseint = ⊂((♥orderpreviousquarter-♥orderthisquarter)*100/♥orderpreviousquarter)⊃
word.replace from ‴[ORDERUTVECKLING]‴ to ‴⊂"Orderläget är svagt för det fjärde kvartalet. Ordervolymen är ner " + ♥orderdecreaseint + "%- jämfört med föregående kvartal, vilket medför att vi kommer att se en lägre försäljning under första delen av 2018."⊃‴
♥searchitem = ⊂♥orderdecreaseint.ToString()+"%-"⊃
jump ➤addcolourtonumber searchitem ♥searchitem colour ‴red‴
➜nextstatement4

word.replace from ‴[STÖRSTA ÅTERFÖRSÄLJARE]‴ to ‴⊂♥highestreseller + " är den största återförsäljaren räknat i försäljning från årets början. Totalt har " + ♥highestreseller + " uppnått en försäljning på " + ♥highestsalesint + " kkr under året."⊃‴
-delay here only for demo purposes
delay milliseconds 200
♥endtime = (time)System.DateTime.Now
♥elapsedtime = ♥endtime - ♥starttime

keyboard text ⋘CTRL+B⋙
keyboard text ‴[datum]‴
delay milliseconds 200
keyboard ⋘ENTER⋙
delay milliseconds 200
keyboard ⋘ALT⋙⋘Ö⋙⋘G⋙
delay milliseconds 200
keyboard text ⊂DateTime.Now.ToShortDateString()⊃
delay milliseconds 100
keyboard ⋘ESC⋙
delay milliseconds 300

♥filnamn = ⊂("rapport_kvartal4_" + System.DateTime.Now.ToShortDateString() + System.DateTime.Now.ToShortTimeString()).Replace(":", "").Replace("-", "")⊃

keyboard text ⋘CTRL+B⋙
delay milliseconds 200
keyboard text ‴filnamn‴
delay milliseconds 200
keyboard ⋘ENTER⋙
delay milliseconds 200
keyboard ⋘ALT⋙⋘Ö⋙⋘G⋙
delay milliseconds 200
keyboard text ‴⊂♥filnamn+".docx"⊃‴

word.save path ‴♥appdata\Bitoreq AB\OfficeDemo\done\♥filnamn.docx‴
word.close

jump ➜skipoutlook if ⊂string.IsNullOrWhiteSpace(♥receiver)⊃
♥subject = ‴Smartare Rapportering‴
♥body = ‴⊂"Nu är rapporten för det fjärde kvartalet uppdaterad. Rapporten är den MS Word-fil som bifogats detta mail tillsammans med en MS Excel-fil som innehåller det konsoliderade underlaget från alla återförsäljare." + "\n" + "Din digitala medarbetare har uppdaterat rapporten på " + ♥elapsedtime.Minute + " minuter och " + ♥elapsedtime.Second + " sekunder. " + "\n" + "Med vänliga hälsningar, Robin"⊃‴
outlook.open
window title ‴✱Outlook✱‴ style maximize
outlook.newmessage to ♥receiver subject ♥subject body ♥body attachments ‴♥appdata\Bitoreq AB\OfficeDemo\done\♥filnamn.docx‴❚‴♥appdata\Bitoreq AB\OfficeDemo\done\konsolidering.xlsx‴
window title ‴✱Meddelande✱‴
outlook.send
delay seconds 10
outlook.close
➜skipoutlook

-****************************************************************************************
file.delete ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\seven.txt‴

file.copy path ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\process.txt‴ destinationpath ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\eight.txt‴

delay seconds 2
ie.attach phrase ‴Office Bot 3000‴
ie.refresh
ie.detach
delay seconds 1

file.delete ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\eight.txt‴
file.delete ‴C:\Users\bitoreq_sattar\source\repos\officedemo\officedemo\Content\process_start.txt‴

delay seconds 1
ie.attach phrase ‴Office Bot 3000‴
ie.refresh
ie.detach
delay seconds 2
-****************************************************************************************

jump ➜filesearch

procedure ➤insertnewrow
-lägg in ny rad
excel.insertrow row 5 where ‴below‴

delay milliseconds 100
keyboard ⋘F5⋙
delay milliseconds 100
keyboard ⋘A⋙⋘7⋙⋘ENTER⋙
delay milliseconds 100
keyboard ⋘SHIFT+SPACE⋙
delay milliseconds 100
keyboard text ⋘ALT⋙
delay milliseconds 100
keyboard text ⋘W⋙
delay milliseconds 100
keyboard text ⋘F⋙
delay milliseconds 100
keyboard text ⋘P⋙
delay milliseconds 100
keyboard ⋘UP⋙
delay milliseconds 100

excel.setvalue row 6 col B value ♥fictivedate
excel.setvalue row 6 col C value ♥reportrow⟦b⟧
excel.setvalue row 6 col D value ♥reportrow⟦d⟧ 
excel.setvalue row 6 col E value ♥reportrow⟦f⟧
excel.setvalue row 6 col F value ♥reportrow⟦e⟧
end

procedure ➤addcolourtonumber searchitem ‴‴ colour ‴‴
delay milliseconds 200
keyboard text ⋘CTRL+B⋙
delay milliseconds 200
keyboard text ‴♥searchitem‴
delay milliseconds 200
keyboard ⋘ALT⋙
delay milliseconds 200
keyboard ‴W‴
delay milliseconds 200
keyboard ‴F‴
delay milliseconds 200
keyboard ‴E‴
delay milliseconds 200
jump ➜blue if ⊂♥colour == "blue"⊃
jump ➜red if ⊂♥colour == "red"⊃
➜blue
keyboard ⋘DOWN⋙⋘DOWN⋙⋘DOWN⋙⋘DOWN⋙⋘DOWN⋙⋘DOWN⋙⋘DOWN⋙
delay milliseconds 100
keyboard ⋘RIGHT⋙⋘RIGHT⋙
delay milliseconds 500
keyboard ⋘ENTER⋙
delay seconds 1
word.replace from ‴%+‴ to ‴%‴
delay milliseconds 200
jump ➜colourdone

➜red
keyboard ⋘DOWN⋙⋘DOWN⋙⋘DOWN⋙⋘DOWN⋙⋘DOWN⋙⋘DOWN⋙⋘DOWN⋙
delay milliseconds 100
keyboard ⋘RIGHT⋙⋘RIGHT⋙⋘RIGHT⋙⋘RIGHT⋙⋘RIGHT⋙
delay milliseconds 500
keyboard ⋘ENTER⋙
delay seconds 1
word.replace from ‴%-‴ to ‴%‴
delay milliseconds 200

➜colourdone
end
