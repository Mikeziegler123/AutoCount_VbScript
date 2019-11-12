# AutoCount.Vbs
Purpose: To Automate the bi-daily process 
of propagating the counts into the POS IntraWeb
system - after physical policy mandated verification 
has been completed, signed, and filed, and all case
counterrors have been reconciled.
  
Notes: This script version relies on a few
 necessary changes in order to run correctly: 

(1) On the Internet Explorer Favorites tab,
    the POS must be on the favorites bar with a
    shortcut key('F1'). 

(2) On the Internet Explorer
    favorites tab, The Counts page within our POS
    must be saved to the Favorites Bar, and be
    associated with the shortcut 'F2'.The URL can
    be copied after logging in->Reconcile->Case
    counts, and saved on the Favorites Bar under 
    the name: "Counts".

(3)  POS credentials
     Count Reconciliation credentials
     AssetTag credentials are: 
           (1) Not public. 
           (2) Need to be written into the script. 
           (3) Expectant that the credentials are 
               saved in the browser's cache.  
