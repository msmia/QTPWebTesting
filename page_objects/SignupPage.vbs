'************************************
' Signup Page							'
'************************************

Public Function signupPageInstance()
	Set signupPageInstance = New SignupPage
End Function


Class SignupPage
	'Declare the variables
	Private myp
	Private objUserName
	Private objPassword
	Private objLogin
	Private objEmailOrPhoneOuterText
	Private objYear
	
	
	
	'Constructor Here
	Private Sub Class_Initialize() 
		
	   ' Set Browser	
	   set myp = Browser("creationtime:=0").Page("title:=.*")
	   	   myp.Sync

	   '============================
	   'Declare all the objects here
	   '============================
			
	   'Email
	    Set objUserName = myp.WebEdit("html id:=email")
	   'Password
	    Set objPassword = myp.WebEdit("html id:=pass")
	   'Login button
	    Set objLogin = myp.WebButton("name:=Log In")
	   'Email or Phone web element
	    Set objEmailOrPhoneOuterText = myp.WebElement("outertext:=Email or Phone","index:=0")
	    'Year WebList field
	    Set objYear = myp.WebList("html id:=year")
		
	End Sub
	
	
	'Wait until home page appears
	Public Function waitForSignupPageToLoad()
		passMsg = "Home page loaded successfully."
		failMsg = "Home page is taking too long to load."
		passfail = waitForPageToLoad(objEmailOrPhoneOuterText, passMsg, failMsg)
	End Function
	
	'Set username
	Public Function setUserName(value)
		passMsg = "Found username field on the Home page."
		failMsg = "Unable to find the username field on the Home page."
		passfail = enterWebEdit(objUserName, value, passMsg, failMsg)
	End Function
	
	'Set password
	Public Function setPassword(value)
		passMsg = "Found password field on the Home page."
		failMsg = "Unable to find the password field on the Home page."
		passfail = enterWebEdit(objPassword, value, passMsg, failMsg)		
	End Function	
	
	'Select year
	Public Function selectYear(value)
		passMsg = "Year was selected successfully from the web list."
		failMsg = "Unable to select the year from the web list."
		passfail = selectFromWebList(objYear, value, passMsg, failMsg)
	End Function
	
	'Click Login	
	Public Function clickLogin()
		clickWebButton(objLogin)
	End Function
	
	
End Class
