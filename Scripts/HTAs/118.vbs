<html>
<title>My HTA</title>
<HTA:APPLICATION 
     ID="objMyHTA" 
     APPLICATIONNAME="MyHTA"
     SCROLL="yes"
     SINGLEINSTANCE="yes"
     WINDOWSTATE="normal"
>

<!-- ************************* -->

<head>
<script language="vbscript">

Sub Window_OnLoad
	self.ResizeTo 300,300
	self.MoveTo 10,10
	section1.InnerHTML = "Enter your credentials"
End Sub

Sub btnOK_OnClick
	Dim oDialog
	Set oDialog = window.Open("about:blank","PleaseWait","height=15,width=250,left=300,top=300,status=no,titlebar=no,toolbar=no,menubar=no,location=no,scrollbars=no")
	oDialog.Focus()
	oDialog.document.body.innerHTML = "Checking..."
	for y = 0 to 500 step 100
		for x = 0 to 800
			oDialog.MoveTo x,y
		next
	next
	oDialog.Close
	section1.InnerHTML = "<font color='red'>Not authorized.</font>"
End Sub

</script>
</head>

<!-- ************************* -->

<body>
<form>
	<p>Name: <input id="txtName" type="text" name="txtName" size="20"></p>
	<p>Password: <input id="txtPassword" type="text" name="txtPassword" size="20"></p>
	<p><input id="btnOK" type="button" value="Submit" name="btnOK"></p>
</form>
<div id="section1" name="section1"></div>
</body>