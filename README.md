postBox
=======

Business Time-Savers for Outlook

Currently only Outlook 2010 is supported, however it is likely to be at least partially compatible with older/newer versions.


Installation
------------

Unfortunately Microsoft prevents Outlook Macros from easily being shared between systems, you'll therefore need to follow this fairly convoluted route to get it working:
* From the main Outlook Explorer window click 'New E-mail'
* From the New Message Window right click on the ribbon, select 'Customize the Ribbon'
* Hit the 'Import/Export' button on the bottom-right 'Customizations' area, which will open the Outlook Options window
* Select 'Import customization file' from drop down
* In the File Open dialog Browse to 'postBox.exportedUI' and click 'Open'
* Opt to replace all customizations by clicking 'Yes' when prompted
* Press 'Ok' to close the Outlook Options window
* You should now have to new buttons on the ribbon in the 'postBox' group

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
