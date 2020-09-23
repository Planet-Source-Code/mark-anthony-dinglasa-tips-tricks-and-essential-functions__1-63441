VERSION 5.00
Begin VB.Form frmTips 
   Caption         =   "TIPS...TRICKS....FUNCTIONS"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "TIPS,TRICKS INSIDE !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
End
Attribute VB_Name = "frmTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Clears text for the controls that supports text
'property like Textbox and Combo Box And Also in the
'same procedure, You can enable or Disable all
'controls. Just replace ".text" to "enabled" and
'set to "true or false"

'Method1
    Function ClearsTextMethod1(Frm As Form)
        On Error Resume Next
            For i = 0 To Frm.Controls.Count - 1
                Frm.Controls(i).Text = ""
            Next
    End Function

'Method2
    Function ClearsTextMethod2(Frm As Form)
        Dim ctrl As Control
            On Error Resume Next
                For Each ctrl In Frm.Controls
                    ctrl.Text = ""
                Next
    End Function

'Get All Fonts
'supported by your System
'Add a listbox named lstFonts
    Function GetSysFonts()
        Dim i As Integer
            For i = 0 To Screen.FontCount - 1
                lstFonts.AddItem Screen.Fonts(i)
            Next
    End Function
    
'This is very reusable codes
'Especially those dealing with databases
'You can control the closing of your
'program. Just select one or more of them
'to be controlled and put "Cancel=True" on it
    'For Example:
        'This will not allow the program to be
        'close in FormCode and by the Task-Manager
        ' Put this code in Unload or QueryUnload event
            'Select Case UnloadMode
                'Case vbFormCode, vbAppTaskManager
                    'Cancel = True
            'End Select
    Private Sub Form_QueryUnload(Cancel As Integer, _
        UnloadMode As Integer)
            Select Case UnloadMode
                Case vbFormControlMenu ' = 0
                    ' Form is being closed by user.
                Case vbFormCode ' = 1
                    ' Form is being closed by code.
                Case vbAppWindows ' = 2
                    ' The current Windows session is ending.
                Case vbAppTaskManager ' = 3
                    ' Task Manager is closing this application.
                Case vbFormMDIForm ' = 4
                    ' MDI parent is closing this form.
                Case vbFormOwner ' = 5
                    ' The owner form is closing.
            End Select
    End Sub

'This is just a Trick
'You can make your forms uncloseable
'in 3 ways

'First Trick
'In Unload Event put Cancel=true
'Like This:
    Private Sub Form_Unload(Cancel As Integer)
        Cancel = True
    End Sub

'Second Trick
'Just let it close by the user
'and create another instance of that Form
'Here it is:
    Private Sub Form_Unload(Cancel As Integer)
        Dim Frm As Object 'as Object or as Form1 will Do
            Set Frm = New Form1 'Anyform you want
                Frm.Visible = True
            Set Frm = Nothing
    End Sub

'Third Trick
'This trick can be done for Compiled Program
'which is .exe extension
'Here it is:
'Put this code in Unload or QueryUnload Event
    Shell App.Path & "\" & App.EXEName

'If you have any tricks in there, Let me know!

'-------------------------- TIPS ---------------------------

'First Tip:

    'Avoid using "End" command to close your program
    'for sometimes Form_Unload and Form_QueryUnload will
    'not fired up Use "End" command if you have no cleanup
    'codes in your Form_Unload or Form_QueryUnload event.
    'Always use "Unload" command to ternimate your program.
    'Always put your cleanup codes in "Form_QueryUnload " event"
    'co'z "Form_QueryUnload" event fires first Before
    '"Form_Unload" events does when closing the program.
    'Lastly use "Form_Inialize" event instead of "Form_Load"
    'Co'z Form_Initialize fires first before Form_Load does

'Second Tip:
'The Windowless Controls Library

    'Visual Basic 6 comes with a new library of windowless
    'controls that exactly duplicate the appearance and the
    'features of most Visual Basic intrinsic controls. This
    'library isn't mentioned in the main language
    'documentation, and it must be installed manually from
    'the Common\Tools\VB\Winless directory. This folder
    'contains the Mswless.ocx ActiveX control and the
    'Ltwtct98.chm file with its documentation. To install
    'the library, you first copy this directory on your hard
    'disk. Before you can use the control, you must register
    'it using the Regsvr32.exe utility or from within Visual
    'Basic, and then double-click on the Mswless.reg file,
    'which creates the Registry keys that make the ActiveX
    'control available to the Visual Basic environment.

    'Once you have completed the registration step, you can
    'load the library into the IDE by pressing the Ctrl+T key
    'and selecting the Microsoft Windowless Controls 6 item
    'from the list of available ActiveX controls. After you
    'do this, you'll find that a number of new controls
    'have been added to the Toolbox. The library contains a
    'replacement for the TextBox, Frame, CommandButton,
    'CheckBox, OptionButton, ComboBox, ListBox, and the two
    'ScrollBar controls. It doesn't include Label, Timer, or
    'Image controls because the Visual Basic versions are
    'already windowless. Nor does it contain PictureBox and
    'OLE controls, which are containers and can't therefore
    'be rendered as windowless controls.

    'The controls in the Windowless Controls Library don't
    'support the hWnd property. As you might remember from
    'last submission, this property is the handle of the window on
    'which the control is based. Since these controls are
    'windowless, there's no such window and therefore the
    'hWnd property doesn't make any sense. Other properties
    'are missing, namely those that have to do with DDE
    'communications. (DDE is, however, an outdated technology
    'and isn't covered in this Tip.) Another difference is
    'that the WLOption control (the windowless counterpart of
    'the OptionButton intrinsic control) supports the new
    'Group property, which serves to create groups of mutually
    'exclusive radio buttons. (You can't create a group of
    'radio buttons by placing them in a WLFrame control
    'because this control doesn't work as a container.)

    'Apart from the hWnd property and the Group property, the
    'controls in the library are perfectly compatible with
    'Visual Basic's intrinsic controls in the sense that
    'they expose the same properties, methods, and events
    'as their Visual Basic counterparts. Interestingly, the
    'library's controls offer a number of property pages that
    'let the programmer set the properties in a logical
    'manner.

    'The real advantage of using the controls in the
    'Windowless library is that at run time they aren't
    'subject to many of the limitations that the intrinsic
    'controls are. In fact, all their properties can be
    'modified during execution, including the MultiLine and
    'ScrollBars properties of the WLText control, the
    'Sorted and Style properties of the WLList and WLCombo
    'controls, and the Alignment property of the WLCheck and
    'WLOption controls.

    'The ability to modify any property at run time makes the
    'Windowless library a precious tool when you're
    'dynamically creating new controls at run time using the
    'Controls.Add method. When you add a control,
    'it's created with all properties set to their default
    'values. This situation means that you can't use the
    'Controls.Add method to create multiline intrinsic
    'TextBox controls or sorted ListBox or ComboBox controls.
    'The only solution is to use the Windowless Controls
    'Library:
        Dim WithEvents TxtEditor As MSWLess.WLText
            Private Sub Form_Load()
                Set TxtEditor = Controls.Add("MSWLess.WLText", "txtEditor")
                    TxtEditor.MultiLine = True
                    TxtEditor.ScrollBars = vbBoth
                    TxtEditor.Move 0, 0, ScaleWidth, ScaleHeight
                    TxtEditor.Visible = True
            End Sub

'Third Tip:
'Unreferenced Controls
    
    'So far, i 've described what you have to do to add
    'controls that are referenced at design time in the
    'Toolbox. But you can do more with the dynamic control
    'creation feature than I've shown you so far; its
    'greater power lies in letting you create ActiveX
    'controls that aren't referenced in the Toolbox. You
    'can provide support for versions of ActiveX controls
    'that don't exist yet at compile time, for example by
    'storing the control's name in an INI file that you edit
    'when delivering a new version of the control. This adds
    'tremendous flexibility to your applications and lets
    'you transform your forms into generic ActiveX control
    'containers.

    'The first issue you must resolve when working with
    'controls not referenced in the Toolbox is design-time
    'licensing. Even if you're not actually using the control
    'at design time, to dynamically load it at run time you
    'must prove that you're legally allowed to do so. If there
    'weren't any restrictions to dynamically creating ActiveX
    'controls at run time, any programmer could "borrow"
    'ActiveX controls from other commercial software and use
    'them in his or her applications without actually
    'purchasing the license for the controls. This is an
    'issue only for ActiveX controls that aren't referenced
    'in the Toolbox at design time; if you can load a control
    'in the Toolbox, you surely own a design-time license for
    'the control.

    'To dynamically create an ActiveX control not referenced
    'in the Toolbox at compile time, you must exhibit your
    'design-time license at run time. In this context, a
    'license is a string of characters or digits that comes
    'with the control and is stored in the system Registry
    'when you install the control on your machine. Visual
    'Basic doesn't force you to search for this string in the
    'Registry because you can find it by means of the Add
    'method of the Licenses collection:

    ' This statement works only if the MSWLess library is
    ' *NOT* currently referenced in the Toolbox.
        Dim licenseKey As String
            licenseKey = Licenses.Add("MSWLess.WLText")

    'After you have the license string, you must devise a way
    'to make it available to the application at run time. The
    'easier method is storing it in a file:

        Open "MSWLess.lic" For Output As #1
            Print #1, licenseKey
        Close #1

    'The preceding code must be executed just once during the
    'design process, and after you've generated the LIC file
    'you can throw the code away. The application reads this
    'file back into the Licenses collection, again using the
    'Add method but this time with a different syntax:

        Open "MSWLess.lic" For Input As #1
            Line Input #1, licenseKey
        Close #1
            Licenses.Add "MSWLess.WLText", licenseKey

    'The Licenses collection also supports the Remove method,
    'but you will rarely need to invoke it.

    'Late-bound properties, methods, and events
    
    'Once you resolve the licensing issue, you're ready to
    'face another problem that comes up when you're working
    'with ActiveX controls not referenced in the Toolbox at
    'compile time. As you might imagine, if you don't know
    'what control you'll load at run time, you can't assign
    'the return value of the Controls.Add method to an object
    'variable of a specific type. This means that you have no
    'simple way to access properties, methods, or events of
    'your freshly added control.

    'The solution offered by Visual Basic 6 is a special type
    'of object variable named VBControlExtender. This
    'represents a generic ActiveX control inside the Visual
    'Basic IDE:

        Dim WithEvents TxtEditor As VBControlExtender

            Private Sub Form_Load()
                ' Add the license key to the Licenses collection.
                Set TxtEditor = Controls.Add("MSWLess.WLText", "TxtEditor")
                    TxtEditor.Move 0, 0, ScaleWidth, ScaleHeight
                    TxtEditor.Visible = True
                    TxtEditor.Text = "My Text Editor"
            End Sub

    'Trapping events from an ActiveX control not referenced
    'in the Toolbox is a bit more complex than accessing
    'properties and methods. In fact, the VBControlExtender
    'object can't expose the events of the control it will
    'host at run time. Instead, it supports only a single
    'event, named ObjectEvent, which is invoked for all the
    'events raised by the original ActiveX control. The
    'ObjectEvent event receives one argument, an EventInfo
    'object that in turn contains a collection of
    'EventParameter objects. This collection enables the
    'programmer to learn what arguments were passed to the
    'event.

    'Inside the ObjectEvent event procedure, you usually
    'test the EventInfo.Name property to discern which event
    'was fired, and then you read, and sometimes modify, the
    'value of each of its parameters:

        Private Sub TxtEditor_ObjectEvent(Info As EventInfo)
            Select Case Info.Name
                Case "KeyPress"
                    ' The Escape key clears the editor.
                    If Info.EventParameters("KeyAscii") = 27 Then
                        TxtEditor.object.Text = ""
                    End If
                Case "DblClick"
                    ' Just to prove that we can trap any event
                        MsgBox "Why have you double-clicked me?"
            End Select
        End Sub

    'Events trapped in this way are called late-bound events.
    'There's a group of extender events that you don't trap
    'inside the ObjectEvent event. These extender events
    '(one of which is shown in the following code snippet)
    'are available as regular events of the VBControlExtender
    'object. This group of events includes GotFocus,
    'LostFocus, Validate, DragDrop, and DragOver.

        Private Sub TxtEditor_GotFocus()
            ' Highlight textbox's contents on entry.
            TxtEditor.object.SelStart = 0
            TxtEditor.object.SelLength = 9999
        End Sub

    'Thats all folks! See you next time for another Submission.

'Note:

    'Hey all programmers out there especially those advance ones
    'I encouraged you to share your knowledge that you have in VB
    'Tutorials would be a great start in learning VB 's especially
    'for those beginners out there.

    'I encouraged all to submit Tips and Tutorials in Vb
    'Thanks all !
    
'Name: Mark Anthony Dinglasa
'Email: mark_anthony_dinglasa@yahoo.com
'Site: www.geocities.com/mark_anthony_dinglasa/2003

'Date Created: November 30, 2005
'Date Finish : December 1, 2005
