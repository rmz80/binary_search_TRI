Attribute VB_Name = "Module1"
'**********************************************************************************
'
'    BinTri - A fast binary search program which exploits additional return values.
'
'**********************************************************************************

' UIUC license
'
' Copyright (c) 2021 Richard Marsden.  All rights reserved.
'
' Developed by: marsden.richard.john@gmail.com
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of
' this software and associated documentation files (the "Software"), to deal with
' the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
' of the Software, and to permit persons to whom the Software is furnished to
' do so, subject to the following conditions:
'
' * Redistributions of source code must retain the above copyright notice,
'   this list of conditions and the following disclaimers.
' * Redistributions in binary form must reproduce the above copyright notice,
'   this list of conditions and the following disclaimers in the documentation
'   and/or other materials provided with the distribution.
' * Neither the names of Richard Marsden, nor the names of its contributors
'   may be used to endorse or promote products derived from this Software
'   without specific prior written permission.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
' CONTRIBUTORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS WITH THE
' SOFTWARE.

'*************************************************************************
'
'   'NOTE: Testing is by UnREMing
'
'******************************************************************************
' In VBA these are module declarations i.e. Applies to everything on this page.
'******************************************************************************
Option Explicit           'All variables must be declared
Option Base 0             'arrays start at element 0
Option Compare Binary     'All comparisions in this module are case sensitive. IMPORTANT !!!
'Option Compare Text       '........................................insensitive.
'*************************************

Public G_arr(999999) As String '****** A global array (0..999999)  = 1 million elements ****
'Public G_arr(8999999) As Long '****** A global array (0..8999999) = 9 million elements ****

Public G_count As Long       '****** A global which is used for test purposes only ***
 
 
'***********************************************************************************
'NOTE #1: If translating this code to another programming language (and version number!)
'      Check how yours handles integer/integer. The \ sign in VBA is integer division
'      operator and means 3 \ 2 = 1........ whereas 3 / 2 = 2 (rounds up!)
'***********************************************************************************

'***********************************************************************************
'
'NOTE: MonoBlock is derived from "monobound_binary_search" published in
'      https://github.com/scandum/binary_search/blob/master/binary_search.c
'
'***********************************************************************************

'*******************************************************************************************
'                 MonoBlock - A fast Binary Search with just one compare in loop
'*******************************************************************************************

 Function MonoBlock(target As String, ByRef trow As Long, ByRef tarr As Long) As Boolean

  
  Dim low As Long     'VBA: Long are signed 32 bit integers
  Dim high As Long
  
  Dim mid As Long
  Dim tval As Long
  'Dim ret As Long   'ALTERNATE TEST see #2

  
  '*********************************************
  'use only on sorted text arrays. No duplicates
  '*********************************************
  trow = -1
  
  low = 0
  
  high = tarr
                               
  MonoBlock = False 'initialise
  '*******************************************
       
  Do While high > 1
    
    mid = high \ 2           'see #1
    
    tval = low + mid
    
    
    G_count = G_count + 1  'FOR TESTING: total up number of compares
    
    '***********************************************
    
    If target >= G_arr(tval) Then   'This is where processor time is eaten
      low = tval
    End If
    
'    '************** ALTERNATE TEST see #2 ********************
'    ret = StrComp(target, G_arr(tval))
'
'    If ret = 0 Or ret = 1 Then   'means >= if StrComp had been used
'      low = tval                 'instead of
'    End If                       'target >= G_arr(tval)
'
'    '**************************************************

    high = high - mid
  
  Loop
  '*****************************************
     
  G_count = G_count + 1          'equality compare
  
  If G_arr(low) = target Then    'fast because avoids an equality
    MonoBlock = True             'compare in loop but has fixed
    trow = low                   'loop count.
  End If
End Function
  
'  '************** ALTERNATE TEST see #2 ********************
'  ret = StrComp(G_arr(low), target)
'  If ret = 0 Then
'    MonoBlock = True
'    trow = low
'  End If
'  '**************************************************'
'End Function

' NOTE #2: ALTERNATE TEST
'
' For strings BinTri uses the function STRCOMP. To ensure its use doesn't give it an unfair speed
' advantage MonoBlock has been timed with and without its use. The speed matches in each case.
' CONCLUSION
' 1) Its the logic that makes BinTri faster NOT its use of STRCOMP.
' 2) The VBA compiler (which uses Microsoft P-Code) will be using the same underlying
'    code whether >= or STRCOMP is used.
'
 

'*******************************************************************************************
'                        BinFind - Standard Binary Search
'*******************************************************************************************
Function BinFind(target As String, ByRef trow As Long, ByRef tarr As Long) As Boolean

  
  Dim low As Long        'VBA: Long are signed 32 bit integers
  Dim high As Long
 
  Dim mid As Long
  
 
  trow = -1
  
  low = 0
 
  high = tarr
  
  BinFind = False
  '*******************************************
       
  Do While low <= high
    mid = (low + high) \ 2  'VBA: an integer div eg 3 \ 2  = 1
  
    
    '*****************************************
    G_count = G_count + 2       'FOR TESTING: total up number of compares
    
    If target = G_arr(mid) Then           'compare 1
      
      BinFind = True
      trow = mid
      
      G_count = G_count - 1               'or too many counts on leaving
      
      Exit Do
    '*****************************************
    ElseIf target < G_arr(mid) Then       'compare 2
    
      
      high = mid - 1
    '*****************************************
    Else
      low = mid + 1
    End If
    '*****************************************
   Loop
 
End Function


'*******************************************************************************************
'                                     BinTri - A Fast Binary Search
'*******************************************************************************************


Function BinTri(target As String, ByRef trow As Long, ByRef tarr As Long) As Boolean
  
  Dim low As Long             'VBA: Long are signed 32 bit integers
  Dim high As Long
  Dim ret As Long
 
  Dim mid As Long
  
 
  trow = -1
  
  low = 0
 
  high = tarr
  
  BinTri = False
  '*******************************************
       
  Do While low <= high
    mid = (low + high) \ 2       'VBA: an integer div eg 3 \ 2  = 1
  
    
    '*****************************************
    G_count = G_count + 1         'FOR TESTING: total up number of compares
       
    
    ret = StrComp(target, G_arr(mid))  ' SEE README (FOR STRINGS)
     
    'ret = target - G_arr(mid) ' SEE README  (FOR POSITIVE INTEGERS)
    
    If ret = 0 Then                'A MATCH!
      
      BinTri = True
      trow = mid
      
      Exit Do
    '*****************************************
    ElseIf ret < 0 Then   'target less than G_arr(mid)
    
      high = mid - 1
    '*****************************************
    Else                  'target more than G_arr(mid)
      low = mid + 1
    End If
    '*****************************************
   Loop
 
End Function

'*******************************************************************************************

 
                       

'************************************************************************
'
' gen_arr - create a test array of random but incrementing and non dupicating
'           strings or 32 bit positive signed integers.
'           Strings are between 58 to 60 char length. last char being a
'           unicode char (UTF16)
'
'************************************************************************
Sub gen_arr()
  
 
  Dim a As Long
  Dim trand As Long
  
  Dim tnum As Long
  Dim tbase As Long


  tbase = 0
  For a = 0 To UBound(G_arr)
  
    Randomize

    trand = Int(100 * Rnd)    ' Generate random value between 0 and 99
   
    tnum = tbase + trand
   
   
    G_arr(a) = String(47, "*") & Format(tnum, String(10, "0")) & String(Int(3 * Rnd), "*") & ChrW(&HC9) 'Unicode capital E with acute  'STRINGS
    
    'G_arr(a) = tnum   'INTEGERS
    
    'Debug.Print G_arr (a)
   
    tbase = tbase + 113 'up count high enough to avoid collision with the added random numbers
                        'Using a prime number for this to reduce a repeating pattern developing
   
  
  Next a
  
  Debug.Print "hello"

  
End Sub

'************************************************************************


'********
'
' MAIN() AAAA_test123 : Different tests have been REMed out.
'
'********

Sub AAAA_test123()

  Dim tlook As String
  'Dim tlook As Long
  
  Dim trow As Long
  Dim ret, a
  Dim bitnum As Long
  Dim tarr As Long
  Dim trand As Long
  Dim tupper As Long
  
  Dim startTime As Variant
  Dim tsecs As Double
  
  Dim buf
  
  'Dim starve_memory(10000000) As Long  'TEST


  gen_arr
  
 
  tarr = UBound(G_arr)
  
  
  G_count = 0
  
  
  startTime = Timer
   
  G_count = 0   'TEST: used for counting compares. Can be removed later.
 
  tupper = 10000000      '*** number of searches ****
  
  
  For a = 1 To tupper     'all tests under a minute to run
  
    Randomize    ' Initialize random-number generator.

    trand = Int(tarr * Rnd)    ' grab a random element position number
    
    
    tlook = G_arr(trand) 'grab it
    
    
    'If a Mod 10 = 0 Then tlook = Replace(tlook, ChrW(&HC9), ChrW(&HE9)) 'STRING TEST: Uppercase Unicode E with accent to lowercase e with accent.
    'If a Mod 10 = 0 Then tlook = Replace(tlook, ChrW(&HC9), "Z")        'STRING TEST: set so 10% of searches are NO finds
    
    'tlook = tlook & "0" 'TEST: for no find
    
    
    '***** *********************************************
    'ret = MonoBlock(tlook, trow, tarr)                     'fast binary (low fixed number of compares)
    '**************************************************
    
    '**************************************************
    'ret = BinFind(tlook, trow, tarr)                       'standard binary
    '**************************************************
    
     '**************************************************
     ret = BinTri(tlook, trow, tarr)                         'trinary (half the compares of standard binary)
    '**************************************************
    
    
    If ret = True Then  'REMed out lines were tests to check functions working ok
      'Debug.Print a; " "; tlook; " "; trow; G_arr(trow); " "; G_count
    Else
      'MsgBox "ERROR a NO FIND on a search value:" & tlook
      Debug.Print a; " "; tlook; " "; "N0 FIND"; " "; G_count  'leave UnREMed just in case
    End If

  Next
  
  
  
 
  tsecs = Timer - startTime
  
  If tsecs = 0 Then tsecs = 0.00000001 'stop 0 error
  
  Debug.Print Round(tsecs, 2) & " secs"
  
  Debug.Print Format(tupper / tsecs, "###,####,### searches per sec")
  
  Debug.Print Round(G_count / tupper, 2) & " average compares per search"
  
  

End Sub



