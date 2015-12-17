Attribute VB_Name = "a_ApplicationStart"
Option Explicit
'Copyright 2010-2015
'------------------------------
'This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
'------------------------------
Global Const myversion = "4.0"
'Revisions
'4.0  20151116
'     added BAS file export capability in order to share the code on github
'     revised chart title generation to reflect that employment numbers are not in 1,000s
'     some bug fixes, notably the annoying tendency of the same panes to be frozen incorrectly
'3.1  20111113 added capability of reading
'     IMPLAN Version 3.0 (2009 with .impdb extension) in addtion to IMPLAN Pro 2.0 [.iap extension]
'3.07 20111113 some bug fixes
'3.06 20110118 some bug fixes
'3.0  2010 preparing the application to support a publication:
'     Rodriguez, Abelardo, Willem Braak, and Philip Watson.
'     “Getting to Know the Economy in Your Community: Automated Social Accounting.”
'     Journal of Extension, August 2011.
'     Available at http://www.joe.org/joe/2011august/iw3.php.
'v1-3  versions that were increasingly interactive and automated in functionality
'------------------------------
'Unless stated otherwise the code is written by Willem Braak
' originally while a graduate student in Bioregional Planning and Community Design
' at the University of Idaho to facilitate IMPLAN analysis for an Economic Methods class
' taught by professor Phil Watson. Later changes were in response to user requests
'================================
Global receipts, dollars
Global js, je, jc, jw, ji2, ji3, ji4 As Integer
    'js=end of lines;
    'jc=end of sectors;
    'je = end of endogoneous;
    'jw, ji2, ji3, ji4 describe factor lines (wages & income, taxes)
Global myrange As Range
Global myaddress As String
Global Const authors = "Braak, W., Watson, P. and Rodriguez, A."

Public Function init()
'lastModified 20100725
    Sheets("tools").Select
    Sheets("tools").Unprotect
    Cells(1, 6) = "release: " & myversion
    Cells(12, 9) = authors & " 2010-" & Year(Date) & ". Automated Social Account Matrix release " & myversion & "."
    Cells(13, 9) = "Available at http://www.ecsInsights.org/ASAM" 'formerly http://www.webpages.uidaho.edu/commecondev/asam.html"
    Cells(1, 9) = "GNU GENERAL PUBLIC LICENSE v3.0"
    Cells(4, 9) = "Copyright 2010-" & Year(Date) & " " & authors
    ActiveWorkbook.Names("mergehouseholds").RefersToRange = 1
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    ActiveSheet.EnableSelection = xlUnlockedCells
    Sheets("MacroHelp").Visible = False
    Sheets("structure").Visible = False
    Sheets("inputEMPL").Visible = False
    Application.ErrorCheckingOptions.BackgroundChecking = False
    ActiveWorkbook.Save
    Sheets("tools").Range("a1:f1").Select
End Function

Public Function getVersion()
    getVersion = myversion
End Function


