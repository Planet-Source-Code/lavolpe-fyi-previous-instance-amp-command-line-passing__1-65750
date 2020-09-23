VERSION 5.00
Begin VB.Form frmInstance 
   Caption         =   "Form1"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOtherInstance 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3570
      TabIndex        =   4
      Text            =   "???????"
      Top             =   2310
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtFYI 
      Height          =   510
      Left            =   210
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1710
      Width           =   4245
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   195
      TabIndex        =   0
      Top             =   570
      Width           =   4245
   End
   Begin VB.Label Label2 
      Caption         =   "This is 1-way comm from 2nd Instance to 1st Instance >>"
      Height          =   405
      Left            =   1260
      TabIndex        =   5
      Top             =   2250
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "Command Line used to start 2nd instance (if any)"
      Height          =   315
      Index           =   1
      Left            =   225
      TabIndex        =   3
      Top             =   1365
      Width           =   3795
   End
   Begin VB.Label Label1 
      Caption         =   "Command Line used to start this instance (if any)"
      Height          =   315
      Index           =   0
      Left            =   195
      TabIndex        =   2
      Top             =   255
      Width           =   3795
   End
End
Attribute VB_Name = "frmInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This form and the accompanying module is one, non-subclassing, example of a flexible
' "Previous Instance" checker and command line parameter relayer.

' Its only limiting factor is that a top level form must exist and stay resident
' (visible or not), during the lifetime of the app. As almost all apps have
' some main form that does this, it should not be problem for most. For others,
' simply using SetProp (sample in bas module) on the top level form should suffice

' A ton of comments, but the actual code needed is only 8 lines + the module.
' See the copy & paste version at bottom of this page (Ctrl+End)


Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

' for those who are unsure of what Command$ is.
' That is an application level function that exposes any command line parameters used
' when the application started. Command line parameters are typically entered immediately
' after the EXE's name at a DOS prompt or Run window. They can also be entered in
' a shortcut's properties window.  Last but not least, you can even enter them in
' your project's properties window so you can test your command line parsing routines.

' see following for a decent example:
' http://www.devx.com/getHelpOn/10MinuteSolution/20366
' and should the above link no longer point to the example, here's the zip
' http://www.devx.com/assets/download/9335.zip

Private Sub Form_Load()

' One more note before we get started: When you want to prevent previous instances,
' ensure you check for it immediately, not towards the end of this routine. Why?
' Should your form take a few seconds to load, another instance could have been
' started and it would not be aware of the 1st instance. Besides, why process a
' bunch of lines of code if you might be bugging out anyway?
' Waste of time, make the check in the first lines of executable code.

' Use something truly unique for each project.
' Suggest using/creating a GUID for this purpose
Dim myUniqueID As String
    
' If you prefer to use a generic GUID, you could make it somewhat unique
' for each project by appending the project title
myUniqueID = "10928347lsdflijsf07124" & App.Title

' if for some reason you want to allow multiple instances only when they are
' different versions, you could append the version to the unique ID...

' myUniqueID = "10928347lsdflijsf07124" & App.Title & "." & App.Major & "." & App.Minor
'^^ This way user could run v1 with no 2nd instances and also run v2 with no 2nd instances
    
    Dim hPrevInst As Long, lPropValue As Long
    
    ' to allow the 2nd Instance to pass any command line parameters to the previous
    ' instance without use of subclassing or registry entries, we will use a textbox
    ' control for that purpose. That textbox control should be hidden and locked/disabled
    
    ' See the remarks in the module for how the function is used. It is designed
    ' to be used by both the 1st and 2nd instances, therefore, depending on the
    ' instance, the variables have different meanings.
    
    
    'cache because it is passed ByRef, not ByVal; therefore, the Function can change it
    lPropValue = txtOtherInstance.hWnd
    'determine if we are first instance.
    hPrevInst = IsPrevInstance(Me.hWnd, myUniqueID, lPropValue, True)
    '^^ Note for 2nd instances: when last parameter is True which means we will pass
    '   any command line parameters to 1st instance. Should you need to pass
    '   something else with or besides the command line, then set that parameter
    '   to False and use SendMessage to pass the whatever string you need to.
    '   The returned lPropValue will be the property value that the 1st instance
    '   assigned (should have been a textbox control's hWnd for this purpose).
    '   Sample of SendMessage is in the module. However DO PASS SOMETHING.
    '   Otherwise, your 1st instance won't know to show itself, kinda defeating
    
    If hPrevInst = 0 Then ' we are first instance
        Me.Caption = "First instance {hWnd: " & Me.hWnd & "}"
        If Command$ = "" Then
            Text1.Text = "No command line used to start this initial instance"
        Else
            Text1.Text = Command$
        End If
    Else
        ' not first instance. The lPropValue is now the property value
        ' set by the 1st instance. Use it how you need to.
        
        ' For this project, that propperty value is the hWnd of the textbox
        ' we will send the command line used for this instance. We don't need
        ' to pass the command line here though, because when we called
        ' IsPrevInstance we passed True as the last parameter
        
        Unload Me   ' don't show this other instance
        
        ' for fun & to test 2nd instance command line passing....
        ' 1. compile exe
        ' 2. copy exe to root of c
        ' 3. Open the 1st instance from one of the compiled EXEs
        ' 4. Click on taskbar's Start and then Run (or keybd: FlyingWindow+R)
        ' 5. Type and execute: C:\Project1.exe sample cmdline params /a /1
        ' The 1st instance will show the 2nd instance's command line parameters
    End If
    
End Sub

Private Sub txtOtherInstance_Change()
' this is only activated in the Previous Instance. Copy & paste to your app as needed.
' Should you be afraid someone can modify this remotely, simply add some type of
' validation text to the passed string and test the txtOtherInstance.Text
' for that validation before accepting it.

    If Not Len(txtOtherInstance.Text) = 0 Then
    
        Dim sCmdLine2ndInst As String
        
        ' do what is needed...
        sCmdLine2ndInst = Trim$(txtOtherInstance.Text)
        If Len(sCmdLine2ndInst) = 0 Then
            txtFYI.Text = "None: Another instance was just aborted"
        Else
            txtFYI.Text = sCmdLine2ndInst
            '^^ of course you would want to parse out the command line to do what is needed
        End If
    
        ' do whatever is needed to display your this instance which is the 1st instance
        If Me.Enabled = False Then Me.Enabled = True
        Me.Visible = True
        SetForegroundWindow Me.hWnd
        If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
    
        ' additionally, you may want to use more advanced routines to bring your
        ' window to the forefront if SetForegroundWindow doesn't always
        ' work (flashes windows taskbar instead)
        ' i.e. http://www.thescarms.com/vbasic/alttab.asp uses AttachThreadInput
    
        txtOtherInstance.Text = vbNullString
        ' ^^ clear. Otherwise if another instance is again started with same command line
        '    parameters, this event may not fire
        
    End If
End Sub

Private Sub CopyAndPasteVersion()

' copy & paste to the top of your app's Form_Load event
' Also ensure you use the module's code or add the module to your project

    Dim myUniqueID As String, hPrevInst As Long, lPropValue As Long
    myUniqueID = "10928347lsdflijsf07124" & App.Title ' << use something truly unique for each project.
    lPropValue = txtOtherInstance.hWnd ' << change name as needed
    hPrevInst = IsPrevInstance(Me.hWnd, myUniqueID, lPropValue, True)
    If Not hPrevInst = 0 Then
        Unload Me ' don't show this other instance
        Exit Sub
    End If
    
    ' Change last parameter of IsPrevInstance() to FALSE
    ' if you you want to control what is being passed. But pass something,
    ' otherwise your 1st instance won't be aware it needs to show itself
End Sub
