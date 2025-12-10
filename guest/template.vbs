' SPDX-FileCopyrightText: 2025 PaulCombal
' SPDX-License-Identifier: GPL-3.0-or-later
'
' This script is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This script is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this script. If not, see <https://www.gnu.org/licenses/>.

' VBScript to launch a Microsoft app, wait for its closure, and conditionally log off.
' This version uses polling instead of WMI events to avoid requiring Administrator privileges.

' --- Configuration ---
Const APP_EXE_NAME = "$APP_EXE_NAME"
Const APP_DIR = "C:\Program Files\Microsoft Office\root\Office16\"
Const WAIT_TIME_SECONDS = 20

Dim ACTIVE_APPS
Dim APP_PATH
APP_PATH = APP_DIR & APP_EXE_NAME
ACTIVE_APPS = Array("EXCEL.EXE", "MSACCESS.EXE", "MSPUB.EXE", "ONENOTE.EXE", "OUTLOOK.EXE", "POWERPNT.EXE", "SETLANG.EXE", "WINWORD.EXE")

' --- Initialize Objects ---
Set objShell = CreateObject("WScript.Shell")
' Use WMI to manage processes, which works fine for querying and creating without admin rights.
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

' --- Step 1: Launch application and capture the process ID (PID) ---

' Use the Create method of Win32_Process to launch the application
Dim intProcessID
Dim objProcess
Dim objStartup, intReturnValue
Dim strCommandLineArgs
Dim strFullArguments

strCommandLineArgs = ""
If WScript.Arguments.Count > 0 Then
    For Each arg In WScript.Arguments
        strCommandLineArgs = strCommandLineArgs & Chr(34) & arg & Chr(34) & " "
    Next
    strCommandLineArgs = RTrim(strCommandLineArgs)
End If

strFullArguments = Chr(34) & APP_PATH & Chr(34) & " " & strCommandLineArgs
strFullArguments = RTrim(strFullArguments)

Set objStartup = objWMIService.Get("Win32_ProcessStartup").SpawnInstance_
intReturnValue = objWMIService.Get("Win32_Process").Create(strFullArguments, Null, objStartup, intProcessID)

If intReturnValue <> 0 Then
    WScript.Echo "Error: Failed to launch " & APP_EXE_NAME & ". Return value: " & intReturnValue
    WScript.Quit 1
End If

' --- Step 2: Wait for the specific launched instance to close (using polling) ---

Do
    ' Wait 1 second before checking again
    WScript.Sleep 1000

    ' Check if the specific process still exists
    Set colProcess = objWMIService.ExecQuery("SELECT ProcessID FROM Win32_Process WHERE ProcessID = " & intProcessID)

    If colProcess.Count = 0 Then
        ' Process has closed, exit the loop
        Exit Do
    End If
Loop

' --- Step 3: Wait for the configured time ---

' WScript.Echo "Pausing for " & WAIT_TIME_SECONDS & " seconds..."
WScript.Sleep (WAIT_TIME_SECONDS * 1000)

' ------------------------------------------------------------------------------------------
' --- Conditional Logoff based on ACTIVE_APPS ---
' ------------------------------------------------------------------------------------------

Dim bPreventLogoff
bPreventLogoff = False
Dim strApp

' Loop through each application defined in the ACTIVE_APPS array
For Each strApp In ACTIVE_APPS
    ' Query WMI for running instances of the current application name
    Set colProcesses = objWMIService.ExecQuery("SELECT Name FROM Win32_Process WHERE Name = '" & strApp & "'")

    If colProcesses.Count > 0 Then
        ' Found at least one running instance of an 'active' app.
        ' Set flag to TRUE and stop checking.
        bPreventLogoff = True
        Exit For
    End If
Next

' --- Step 4: Conditional exit or logoff ---

If bPreventLogoff Then
    ' Case A: An application from ACTIVE_APPS is running, exit the script gracefully.
    WScript.Quit 0
Else
    ' Case B: NO application from ACTIVE_APPS is running, execute logoff.
    ' Execute logoff command silently (0) and asynchronously (False)
    objShell.Run "logoff", 0, False

    WScript.Quit 0
End If