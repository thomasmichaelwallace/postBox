postBox
=======

Business Time-Savers for Outlook

Currently only Outlook 2007-2013 are supported, however it is likely to be at least partially compatible with older/newer versions.

Installation
------------

Unfortunately Microsoft prevents Outlook Macros from easily being shared between systems, you'll therefore need to follow this fairly convoluted route to get it working:

### Installing the postBox code:
*Outlook requires you to manually import the postBox code, as a module, into the Outlook application.*
* With Outlook as the active window, press <kbd>Alt</kbd>+<kbd>F11</kbd> to open the Visual Basic editor.
* In the Visual Basic for Applications editor menu click 'File' > 'Import File'.
* In the Import File dialog, browse to the 'postBox.bas' file included in the postBox download.
* Select the 'postBox.bas' file and press 'Open'.
* In the Project Properties browser (typically top-left) a new folder called 'Modules' will appear containing the 'postBox' module.
* Press <kbd>Ctrl</kbd>+<kbd>S</kbd> to save postBox to your Outlook.

### Getting Outlook to trust postBox:
*Outlook will not run unsigned code, even if you imported it yourself.*
* Open the Windows Run dialog by pressing <kbd>Windows</kbd>+<kbd>R</kbd>.
* Type the path to the SelfCert.exe correct for your version:
    * Office 2013: "%ProgramFiles%\Microsoft Office\Office15\SELFCERT.EXE"
    * Office 2010: "%ProgramFiles%\Microsoft Office\Office14\SELFCERT.EXE"
    * Office 2007L "%ProgramFiles%\Microsoft Office\Office13\SELFCERT.EXE"
    * Search Google if you can't find it; sometimes it feels like Microsoft are just making things difficult...
* Press 'OK' to launch the Create Digital Certificate dialog.
* Enter the certificate name as 'postBox' and press 'OK', then press 'OK' again on the successful creation message box.
* Return to the Outlook Visual Basic editor menu and click 'Tools' > 'Digital Signature'.
* On the Digital Signature dialog press the 'Choose...' button, select 'postBox' and press 'OK'.
* The top of the Digital Signature dialog will now state that the VBA project has been signed, press 'OK' to close.
* Press <kbd>Ctrl</kbd>+<kbd>S</kbd> to save the signed postBox to your Outlook.
* Press <kbd>Alt</kbd>+<kbd>F4</kbd> to close the Visual Basic editor.

### Making postBox secure:
*Outlook needs to be told to trust code, even though you've just signed it yourself.*
* Back on the main Outlook window open the Outlook Options dialog 'File'/'Office Button' > 'Options'.
* In the side-bar menu of the Outlook Options dialog select Trust Centre, and press the 'Trust Centre Settings...' button.
* In the side-bar menu of the Trust Centre dialog select 'Macro Settings' and ensure "Notifications for digitally signed macros, all other macros disabled" is selected.
* Press 'OK'
* Close Outlook.
* If prompted, opt to save changes to the VBA project by pressing 'Yes' on the dialog.
* Re-open Outlook.
* The first time you run a postBox macro (see below) you will be promped by a security notice.
* To ensure that you are not forever nagged, opt to 'Trust all documents from this publisher'.

### Adding the postBox menu:
*Outlook doesn't provide a reliable way to automatically create ribbon links, so you'll need to add these yourself.*
* From the main Outlook Explorer window click 'New E-mail'
* From the New Message Window right click on the ribbon, select 'Customize the Ribbon' (In 2007 you can only customise the Quick-Access toolbar.)
* On the customisation dialog, use the left-hand (Choose commands from:) drop-down to select 'Macros'.
* You can then add 'Project1.QuickLink' and 'Project1.ScanAttach' (see below for descriptions) using the 'Add >>' button.
* If prompted you'll need to make a new group for these buttons. I reccomend keeping the 'New Mail Message' menu selected on the right-hand plane and pressing 'New Group', pressing 'Rename', and calling it 'postBox'.
* Use the 'Rename' button after the 'Add >>' button to choose better names and icons for the postBox commands; the decision of which, I'll leave to your good selves...

Usage
=====

Quick Link
----------

Quick link automatically creates a link from the highlighted text- useful for providing links to files on shared networks, which Outlook does not do automatically. Just highlight the filepath (e.g. S:\shared_file.jpg) and press the 'Quick Link' button.

Scan Attach
-----------

Scan Attach moves and renames a file from a scanner to a new e-mail- useful for professional scanning attachments.
To use it you must have selected the e-mail from the scanner (with the pdf attachment) in the Outlook Explorer (main) window. In the New Message composer clicking scan attach will prompt you for a filename. The file will then be copied from the scanner e-mail, renamed as instructed, and attached to the new message you're composing.

Author and Licence
==================

postBox is primiarly written by Thomas Michael Wallace (www.thomasmichaelwallace.co.uk), and released under the GPL v3 licence.
