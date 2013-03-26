VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub LinkToWorkItem()

    'postBox - Quick Link
    'Copyright (C) 2010-2013 Thomas Michael Wallace <http://www.thomasmichaelwallace.co.uk>

    ' This file is part of postBox.

    '    postBox is free software: you can redistribute it and/or modify
    '    it under the terms of the GNU General Public License as published by
    '    the Free Software Foundation, either version 3 of the License, or
    '    (at your option) any later version.

    '    postBox is distributed in the hope that it will be useful,
    '    but WITHOUT ANY WARRANTY; without even the implied warranty of
    '    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    '    GNU General Public License for more details.

    '    You should have received a copy of the GNU General Public License
    '    along with postBox .  If not, see <http://www.gnu.org/licenses/>.

    'Quick convert selected text into hyperlink
    
    Dim editor As Inspector.WordEditor
    
    'view current e-mail as ms-word
    Set editor = ActiveInspector.WordEditor
    
    'develop link from selection
    editor.Hyperlinks.Add _
        Anchor:=editor.Application.Selection.Range, _
        Address:=editor.Application.Selection.Text
    
    Set editor = Nothing
        
End Sub

Sub ScanAttach()

    'postBox - Scan Attach
    'Copyright (C) 2010-2013 Thomas Michael Wallace <http://www.thomasmichaelwallace.co.uk>

    ' This file is part of postBox.

    '    postBox is free software: you can redistribute it and/or modify
    '    it under the terms of the GNU General Public License as published by
    '    the Free Software Foundation, either version 3 of the License, or
    '    (at your option) any later version.

    '    postBox is distributed in the hope that it will be useful,
    '    but WITHOUT ANY WARRANTY; without even the implied warranty of
    '    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    '    GNU General Public License for more details.

    '    You should have received a copy of the GNU General Public License
    '    along with postBox .  If not, see <http://www.gnu.org/licenses/>.

    'Quickly move scan copy to current e-mail attachments.

    'outlook automation objects
    Dim outlook_explorer As Outlook.Explorer
    Dim outlook_inspector As Outlook.Inspector
    Dim selected_email As Outlook.Selection
    
    'email selections
    Dim from_email As MailItem
    Dim to_email As MailItem
    Dim from_attachments As Attachments
    
    'temporary file paths
    Dim file_name As String
    Dim file_path As String
    
    'connect to outlook explorer window and get selected e-mail
    Set outlook_explorer = Application.ActiveExplorer
    Set outlook_inspector = Application.ActiveInspector
    Set selected_email = outlook_explorer.Selection
    
    'check that explorer has one selected e-mail only
    If selected_email.Count <> 1 Or selected_email.Item(1).Class <> olMail Then
        MsgBox "Select scanner attachment e-mail in Outlook Explorer first...", _
            vbOKOnly + vbExclamation, "Scan Attacher"
        Exit Sub
    End If
    
    'select e-mail and fetch attachments
    Set from_email = selected_email.Item(1)
    Set from_attachments = from_email.Attachments
    
    'check e-mail is from scanner (i.e. with one attachment
    If from_attachments.Count <> 1 Then
        MsgBox "Select scanner attachment e-mail in Outlook Explorer first...", _
            vbOKOnly + vbExclamation, "Scan Attacher"
        Exit Sub
    End If
    
    'halt if inspector is for anything else by an e-mail
    If outlook_inspector.CurrentItem.Class <> olMail Then
        MsgBox "Focus on new e-mail as active window.", _
            vbOKOnly + vbExclamation, "Scan Attacher"
        Exit Sub
    End If
    
    Set to_email = outlook_inspector.CurrentItem
    
    'ask user for prefered filename
    file_name = InputBox("Attach as filename", "Scan Attacher", "attachment")
    file_path = "C:\Temp\" & file_name & ".pdf"
    
    'save attachement to temp
    from_attachments.Item(1).SaveAsFile file_path
    
    'attach to e-mail
    to_email.Attachments.Add file_path
    
    'delete files
    Kill file_path

End Sub
