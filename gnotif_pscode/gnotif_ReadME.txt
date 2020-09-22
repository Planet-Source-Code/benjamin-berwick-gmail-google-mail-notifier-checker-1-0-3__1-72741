
##########################################
	To PSCode: If you find any mistakes, or make any corrections
	please comment them on the download page or email me. With minor changes
	this application could easily be customized. Enjoy!
--------------------------------
	Email comments or suggestions to finalcry@gmail.com
	with the subject "pscode"
	Do not include any links, files, or suggested advertisement.
	English only please.
--------------------------------
##########################################


  @@-------------------------------------------------------
  @@ GMail Notifier 1.0.3 -- OUTERTOUCH Productions ---- - \
  @@----------------------------------------------------\ - \
                                               _________/-@@-\____________
	[READ ME]                               Coded by Benjamin.Berwick 
                                               ===========================
((Tray Icons))
	~~  Green G with a white tip, no new mails.
	``  

	~~  Green G with a blue tip, 1 or more new emails.
	``

	~~  Red G with a white tip, an error has occured while attempting to check email(s).
	``	Most likely due to an invalid password.	

	~~  Reg G with a blue tip, an error has occured while attempting to check email(s)
	``	on one or more accounts, but others still have new emails.

	~~  Rotated green G with a tan tip, checking for new mail.
	``

	~~  Rotated White G with a green tip, detecting internet connection.
	``

	~~  Within the tooltip an email shown with {E} represents the last attempt
	``  on checking the displayed email has failed.


((Modifying Account Information))
	~~  The password shown in the password text field is an encrypted string
	``  for your protection. You only need to modify the password field if you
	    wish to change your password.

	~~  Alias to your accounts is what will be displayed within popup notices
	``  of new emails.


((Launching and General Information))
	~~  To launch your web browser with a chosen account simply double click
	``  the account you wish to view within either the "Settings" or "Mail Notices"
	    list. Or right click the tray icon and select "Open Browser".

	~~  While no new mail is present, double clicking the tray icon will defaultly
	``  open the settings window.



::::::::::::::::::::::::::
::::::: CHANGE LOG :::::::
::::::::::::::::::::::::::


Future Plans...
	-Would like to add an easier way to view multiple mail notices, perhaps eliminating the need to select the account and from/subject.


version 1.0.3
	-Fixed a crash caused by a rare situation involving the adding, modification, or removal of an account.
	-Added popup displays upon the arrival of new emails. Search "'//remove to erase popup notices" in source to disable.
		-- You may want to consider custom coding them to not show during check on startup.
	-New email information no longer disappears while adding, modifying, or removing an account.
	-Changed the image on the application, no particular reason.
	-Fixed a minor display error that occured while resetting alerts.
	-Made several alterations to previous code to help optimize the application, and to allow the program to be more flexible.

Version 1.0.2
	-Changed icons...yet again.
	-Fixed an incorrect status display under certain circumstances.
	-Removed the return not needed in the tray icon status display.
	-Added a filter system to eliminate character codes such as, &#39; translated to '  -- all editable within filters.txt
	-Added a few ease of use features:
		-- Application now goes to tray when opening an account in browser.
		-- First email of first account is automatically displayed when selecting "Mail Notices".
		-- First email is now selected when clicking an account name under "Mail Notices".
	-Added a 'Reset' option which will ignore current emails as if they don't exist, only prompting new emails after the reset.


Version 1.0.1
	-Fixed an issue where tray icon would display no new mails after untraying the application while having unread emails.
	-Added a sound that would play if an account obtained new emails since last check. Search "'###SOUND###" in source to comment out/disable.
	-Changed the tray icons and animated them a bit more.
	-Made minor changes to inet execution to hopefully ensure all data was recieved properly before attempting to parse.
	-Commented out Inet_Statechange routine since it was not in use, but kept it for a reference.
	-Upon application startup and failed data retrieval, the application will check for a connection every 5 seconds and display so instead of returning a vague 'error checking' report.


Version 1.0.0
	-Released