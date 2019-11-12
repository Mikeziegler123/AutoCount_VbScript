' File Name: AutoCount.vbs
' Author: Michael Ziegler
' Purpose: To Automate the bi-daily process of propagating the counts into
'  the POS IntraWeb system - after policy mandated verification has been completed, signed, and filed, and all case count
'  errors have been reconciled.  
'
' Notes:  This script version relies on a few necessary changes
'	  in order to run correctly:
'		1) On the Internet Explorer Favorites tab, The POS must be on the favorites bar with a shortcut key('F1').
'		2) On the Internet Explorer Favorites tab, The Counts page within our POS must be saved to the Favorites Bar, and be associated with the shortcut 'F2'.The URL can be copied after logging in->Reconcile->Case Counts, and saved on the Favorites Bar under the name: "Counts".
'		3) IMPORTANT: POS credentials/Count Reconciliation credentials/Asset Tag credentials at the moment are: 
'			(1)ONLY my own,
'			(2)hardcoded into the script,
'		 	 and
' 			(3) Expectant that the credentials aresaved in the browser's cache.  
'
'//Declare counter variables for {TAB} / +{TAB} iteration loops, and initialize.
Dim e, counter1, counter2, counter3, counter4, counter5, loopCount

counter1 = 21
counter2 = 187
counter3 = 29
counter4 = 33
loopCount = 25
e = "{ENTER}"
t = "{TAB}"
Set WShell = WScript.CreateObject("WScript.Shell")
WScript.Sleep 500

'//Shortcut expected to open the POS page within Internet Explorer
WShell.SendKeys "{F2}"

WScript.Sleep 1000
WShell.SendKeys "{F11}"
WScript.Sleep 100
WShell.SendKeys "M"
WScript.Sleep 100
WShell.SendKeys "K"
WScript.Sleep 100
WShell.SendKeys "Z"

WScript.Sleep 50

WShell.SendKeys "{DOWN}"
WScript.Sleep 500
WShell.SendKeys e
WScript.Sleep 750
WShell.SendKeys e
WScript.Sleep 2000

'//Shortcut expected to open the Counts(within POS) page within Internet Explorer
WShell.SendKeys "{F2}"
WScript.Sleep 2000

'//Counter Loop (1)
Dim x
For x = 1 To counter1
	WScript.Sleep 75
	WShell.SendKeys t
Next

'//Propagate Asset Tags (Tags must be physically checked prior to running this program)
WShell.SendKeys "4"
WScript.Sleep 100
WShell.SendKeys "6"
WScript.Sleep 100
WShell.SendKeys "2"
WScript.Sleep 100
WShell.SendKeys "5"
WScript.Sleep 100
WShell.SendKeys "-"
WScript.Sleep 100
WShell.SendKeys "4"
WScript.Sleep 100
WShell.SendKeys "6"
WScript.Sleep 100
WShell.SendKeys "2"
WScript.Sleep 100
WShell.SendKeys "5"
WScript.Sleep 300

WShell.SendKeys t
WScript.Sleep 500
WShell.SendKeys e
WScript.Sleep 500

WShell.SendKeys "4"
WScript.Sleep 100
WShell.SendKeys "6"
WScript.Sleep 100
WShell.SendKeys "2"
WScript.Sleep 100
WShell.SendKeys "6"
WScript.Sleep 100
WShell.SendKeys "-"
WScript.Sleep 100
WShell.SendKeys "4"
WScript.Sleep 100
WShell.SendKeys "6"
WScript.Sleep 100
WShell.SendKeys "2"
WScript.Sleep 100
WShell.SendKeys "6"

WShell.SendKeys t
WScript.Sleep 100
WShell.SendKeys t
WScript.Sleep 500

WShell.SendKeys "m"
WScript.Sleep 200
WShell.SendKeys "{DOWN}"
WScript.Sleep 200

WShell.SendKeys e
WScript.Sleep 500

WShell.SendKeys e
WScript.Sleep 500

'//Counter Loop (2)
Dim y
For y = 1 To counter2
	WScript.Sleep 10
	WShell.SendKeys t
Next

WShell.SendKeys e
WScript.Sleep 200
WShell.SendKeys e
WScript.Sleep 500
Dim z
For z = 1 To 4
	WScript.Sleep 10
	WShell.SendKeys "+{TAB}"
Next

WScript.Sleep 100
WShell.SendKeys "^c"
WScript.Sleep 100
WShell.SendKeys "+{TAB}"
WScript.Sleep 100
WShell.SendKeys "^v"

'//LoopCount (propegation) loop (3)
Dim i
For i = 1 To loopCount
	Dim j
	For j = 1 To 6
		WScript.Sleep 10
		WShell.SendKeys "+{TAB}"
	Next
	WShell.SendKeys "^c"
	WScript.Sleep 100
	WShell.SendKeys "+{TAB}"
	WScript.Sleep 100
	WShell.SendKeys "^v"
Next

'//Reconcile loop (4)
WScript.Sleep 500
Dim l
For l = 1 To 5
	WScript.Sleep 10
	WShell.SendKeys "+{TAB}"
Next
WScript.Sleep 500
WShell.SendKeys e
WScript.Sleep 500

'//Final Loop(From Reconcile to Complete Reconciliation (5)
Dim n
For n = 1 To 26
	WScript.Sleep 10
	WShell.SendKeys "+{TAB}"
Next

WScript.Sleep 1000
WShell.SendKeys e
WScript.Sleep 1000
WShell.SendKeys e