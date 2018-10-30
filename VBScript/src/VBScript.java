1) Option Explicit 'validating for spelling
Dim x
x = Msgbox("Is it your birthday?",36,"Hello")
If x=6  Then
	Msgbox "Happy Birthday"
	If x=7 Then
		Msgbox "Oops"
	End If
End If

2) Dim msg
msg= "Hello Everyone"
Msgbox msg 

3) Option Explicit

Dim age
age =40
print age

4) Option Explicit

Dim a,b,c
a =1
b=2
c=a & b

If a<b and b<c Then
	msgbox "you are correct"
	Else 
	msgbox "you are wrong"
End If

5) Option Explicit

Dim a,b,c
a = Inputbox("Enter First Name")
b = Inputbox("enter Last Name")
c = a*b

msgbox c

6) Option Explicit

Dim a,b,c
a = Inputbox("Enter First Name")
b = Inputbox("enter Last Name")
c = a*b
msgbox c

Function calculation()
	Dim d,e
	d =40
	e = d+a
	msgbox e
End Function
Call calculation()

7) Dim a,b(3)

b(0) = "tanbin"
b(1) = 10
b(2) = 40
b(3) = 50
'b(4) = 20

a = b(1) +  b(3)
msgbox a

8) Dim b
msgbox IsArray(b)

b = Array("Tanbin",10,"VBScript",20,30)
msgbox IsArray(b)
msgbox b(0)
msgbox UBound(b)

9) Dim a,b,operation
a = Inputbox("Enter a Value")
b = Inputbox("Enter b Value")
operation = Inputbox("Enter Value")

Select Case operation
	Case "add"
	msgbox "Addition of a,b is: "& a + b
	Case "sub"
	msgbox "Subtraction of a,b is: "& a - b
	Case "mul"
	msgbox "Multiplication of a,b is: "& a * b
	Case "div"
	msgbox "Division of a,b is: "& a / b
	Case else 
	msgbox "Invalid operation"
End Select

10) Dim obj
Set obj = CreateObject("Scripting.FileSystemObject")
obj.CreateFolder "C:\Users\Tanbin\Desktop\Qtp"
Set obj = Nothing

11) Dim a(3)

a(0) = "Ashraful Hoque Tanbin"
a(1) = "1554 Ocean Avenue"
a(2) = "6316407792"
a(3) = "tanbinnsu101@gmail.com"

For n = 0 To 3 Step 1
	print a(n)
Next

12) Dim a(3),n

a(0) = Inputbox("Enter a name")
a(1) = Inputbox("Enter a address")
a(2) = Inputbox("Enter phone number")
a(3) = Inputbox("Enter Email Address")

13) Dim myName(4) 'variable = place holder
myName(0) = "Rose"
myName(1)= "Qa Engineer"
myName(2)= "Test Automation"
myName(3)= "Hello World"

For k = 0 To 3
    
systemutil.Run "chrome.exe", "www.bing.com"

Browser("title:=.*").Page("title:=.*").WebEdit("name:=q").Set myName(k)
Browser("title:=.*").Page("title:=.*").WebButton("name:=Submit Query").Click
Browser("title:=.*").CloseAllTabs
wait(5)

Next

14) n = inputbox("Enter a value")

For j = 1 To n
	For n = 0 To j
	print a(n)
Next
  	print "..........."
Next

15) myLocation = ucase(inputbox("Please enter Brooklyn"))

If myLocation = "Queen" Then
	print "I eat like to to eat food at yesterday" & " " &mylocation
	ElseIf myLocation = "Brooklyn" Then
	print "I eat like to to eat food at today" & " " &mylocation
	else
	print "I eat like to to eat food at tomorrow" & " " &mylocation
End If

16) Dim sum
sum = 0
 For k = 1 To 100 step 1
 
    If k mod 2 <> 0 Then
        print k
    	sum = sum + k
    End If  

 Next
	print sum


17) Brooklyn = array("tanbin","Kamal","Islam")
Queen = array("Enamul","Zidni","Mustafi")
Bronz = array("Selina","Sunny","Anne")
Canada = array("Masud","Jahangir","Ashraful")

a = inputbox("Enter a City")

If a = "Brooklyn" Then
	For i = 0 To ubound(Brooklyn)
		print Brooklyn(i)
	next

elseif a = "Queen" Then
For j = 0 To ubound(Queen)
		print Queen(j)
    next

elseIf a = "Bronz" Then
For k = 0 To ubound(Bronz)
		print Bronz(k)
    next


elseIf a = "Canada" Then
For l = 0 To ubound(Canada)
		print Canada(l)
    next    
else
   	print "Invalid output"
End If

18) Dim myApp
myApp = "https://www.mlcalc.com/"
myBrowser = "chrome.exe"

systemUtil.Run myBrowser, myApp

Browser("title:=.*").page("title:=.*").WebEdit("name:=ma","index:=0").Set 525000
Browser("title:=.*").page("title:=.*").WebEdit("name:=dp").Set 22
Browser("title:=.*").page("title:=.*").WebEdit("name:=mt").Set 30
Browser("title:=.*").page("title:=.*").WebEdit("name:=ir","index:=0").Set 99
Browser("title:=.*").page("title:=.*").WebEdit("name:=pt").Set 4500
Browser("title:=.*").page("title:=.*").WebEdit("name:=pi").Set 56
Browser("title:=.*").page("title:=.*").WebEdit("name:=zipCode","index:=0").Set 11229

Browser("title:=.*").page("title:=.*").WebList("name:=sm","index:=0").select "Jul"

Browser("title:=.*").page("title:=.*").WebButton("type:=submit","index:=0").Click
wait(4)
Browser("title:=.*").Close

19) operation = Inputbox("Enter a Browser")

Select Case operation
	Case "Chrome"
	systemUtil.run "chrome.exe"
	Case "Internet Explorer"
	systemUtil.run "iexplore.exe"
	Case else 
	msgbox "Invalid operation"
End Select

20) Dim myApp
myApp = "https://www.mlcalc.com/"
myBrowser= "chrome.exe"

	systemUtil.Run myApp,myBrowser

    Browser("title:=.*").Page("title:=.*").WebEdit("name:=ma").Set DataTable("pPrice", dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=dp").Set DataTable("downPayment", dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=mt").Set DataTable("mortgageTerm", dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=ir","index:=0").Set DataTable("interestRate", dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=pt").Set DataTable("propertytax", dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=pi").Set DataTable("propertyInsurance", dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=mi").Set DataTable("PMI", dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=zipCode","index:=0").Set DataTable("zipCode", dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebList("name:=sm","index:=0").Select DataTable("paymentMonth", dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebList("name:=sy","index:=0").Select DataTable("paymentYear", dtGlobalSheet)
 wait(2)
    Browser("title:=.*").Page("title:=.*").WebButton("type:=submit","index:=0").Click
 wait(3)
Browser("title:=.*").Close

21) 
Dim myApp
myApp = "https://www.mlcalc.com/"
UserBrowser = "chrome.exe"

a = array("52500","525001","525002")
b = array("22","23","24")
c = array("30","31","32")
d = array("97","98","99")
e = array("4500","4501","4502")
f = array("56","57","58")
g = array("11227","11228","11229")

	
For It = 0 To 2 
    systemUtil.Run UserBrowser,myApp
    Browser("title:=.*").page("title:=.*").WebEdit("name:=ma","index:=0").Set a(It)  
    Browser("title:=.*").page("title:=.*").WebEdit("name:=dp").Set b(It)   
    Browser("title:=.*").page("title:=.*").WebEdit("name:=mt").Set c(It)
    Browser("title:=.*").page("title:=.*").WebEdit("name:=ir","index:=0").Set d(It) 
    Browser("title:=.*").page("title:=.*").WebEdit("name:=pt").Set e(It)
    Browser("title:=.*").page("title:=.*").WebEdit("name:=pi").Set f(It)
    Browser("title:=.*").page("title:=.*").WebEdit("name:=zipCode","index:=0").Set g(It)
    Browser("title:=.*").Close
    wait(2)
Next    

22) Dim FileLocation 
Dim myBrowser
myBrowser = "iexplore.exe"

myApp = "https://www.mlcalc.com/"'for file location variable
FileLocation = "C:\Users\Tanbin\Desktop\testSheet.xlsx"

DataTable.AddSheet "TestDataInUFT" 'add data sheet after local and global
DataTable.ImportSheet FileLocation,"TestDataForMC","TestDataInUFT"

TotalRows = DataTable.GetSheet("TestDataInUFT").GetRowCount()

For i = 1 To TotalRows
	DataTable.GetSheet("TestDataInUFT").SetCurrentRow(i)
	systemutil.Run myBrowser, myApp
	wait(3)
	Browser("title:=.*").Page("title:=.*").WebEdit("name:=ma").set DataTable.Value("pPrice","TestDataInUFT")
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=dp").Set DataTable.Value("downPayment","TestDataInUFT")
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=mt").set DataTable.Value("Mortgage","TestDataInUFT")
    Browser("title:=.*").page("title:=.*").WebEdit("name:=ir", "index:=0").Set DataTable.Value("Interest_Rate","TestDataInUFT")
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=pt").Set DataTable.Value("Property_TAX","TestDataInUFT")
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=pi").Set DataTable.Value("Property_Insurance","TestDataInUFT")
    Browser("title:=.*").Page("title:=.*").WebEdit("name:=mi").Set DataTable.Value("PMI","TestDataInUFT")
    Browser("title:=.*").page("title:=.*").WebEdit("name:=zipCode","index:=0").Set DataTable.Value("ZipCode","TestDataInUFT")
    Browser("title:=.*").page("title:=.*").WebList("name:=sm", "index:=0").Select DataTable.Value("PaymentMonth","TestDataInUFT")
    Browser("title:=.*").page("title:=.*").WebList("name:=sy","index:=0").Select DataTable.Value("PaymentYear","TestDataInUFT")
    wait(2)
    Browser("title:=.*").page("title:=.*").WebButton("type:=submit","index:=0").Click
    wait(3)
    
    Monthly_Payment = Browser("title:=.*").page("title:=.*").WebTable("name:=WebTable", "index:=3").GetCellData(1,1)
    
    datatable.Value("MonthlyPayment","TestDataInUFT") = Monthly_Payment
    
    
    Browser("title:=.*").Close
	
    Next

    DataTable.ExportSheet "C:\Users\Tanbin\Desktop\testResult.xlsx","TestDataInUFT"

23) Dim myBrowser
myBrowser = "chrome.exe"

myApp = "https://www.mortgagecalculator.org/"

    systemutil.Run myBrowser, myApp 'opening ie browser with app url param[downpayment_type]
    
    Browser("title:=.*").Page("title:=.*").WebNumber("type:=number","index:=0").set DataTable("Home_Value", dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebEdit("html id:=downpayment").Set DataTable("downPayment", dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebRadioGroup("name:=param[downpayment_type]").set DataTable("RadioGroup1", dtLocalSheet)
    Browser("title:=.*").page("title:=.*").WebRadioGroup("name:=param[downpayment_type]").Set Datatable("RadioGroup2",dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebNumber("html id:=loanamt").Set datatable("Loan_Amount",dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebNumber("name:=param[interest_rate]").Set datatable("Interest_Rate",dtGlobalSheet)
    Browser("title:=.*").Page("title:=.*").WebElement("innerhtml:=Get Today's Best Mortgage Rates").Set datatable("Link",dtGlobalSheet)
    Browser("title:=.*").page("title:=.*").WebNumber("html id:=loanterm").Set datatable("Long_Term",dtGlobalSheet)
    Browser("title:=.*").page("title:=.*").WebList("name:=param[start_month]").Select datatable("Start_Month",dtGlobalSheet)
    Browser("title:=.*").page("title:=.*").WebNumber("name:=param[start_year]").Set datatable("Start_Year",dtGlobalSheet)
    Browser("title:=.*").page("title:=.*").WebNumber("name:=param[property_tax]").Set datatable("Property_Tax",dtGlobalSheet)
    Browser("title:=.*").page("title:=.*").WebNumber("name:=param[pmi]").Set datatable("PMI",dtGlobalSheet)
    Browser("title:=.*").page("title:=.*").WebNumber("name:=param[hoi]").Set datatable("Home_Ins",dtGlobalSheet)
    Browser("title:=.*").page("title:=.*").WebNumber("name:=param[hoa]").Set datatable("Monthly_HOA",dtGlobalSheet)
    Browser("title:=.*").page("title:=.*").WebButton("name:=Calculate").Click
    
    Monthly_Payment = Browser("title:=.*").page("title:=.*").WebElement("name:=WebTable", "index:=3").GetCellData(1,1)
    
    datatable.Value("Payment",dtGlobalSheet) = Monthly_Payment
    
    Browser("title:=.*").Close     
    
   