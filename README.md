# AutoCount.vbs

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
