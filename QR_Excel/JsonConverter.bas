' Attribute VB_Name = "JsonConverter"
''
' VBA-JSON v2.3.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
'
' JSON Converter for VBA
'
' Errors:
' 10001 - JSON parse error
'
' @class JsonConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Based originally on vba-json (with extensive changes)
' BSD license included below
'
' JSONLib, http://code.google.com/p/vba-json/
'
' Copyright (c) 2013, Ryo Yokoyama
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' === VBA-UTC Headers
#If Mac Then

#If VBA7 Then

' 64-bit Mac (2016)
Private Declare PtrSafe Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr
Private Declare PtrSafe Function utc_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" _
    (ByVal utc_File As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_fread Lib "/usr/lib/libc.dylib" Alias "fread" _
    (ByVal utc_buffer As String, ByVal utc_size As LongPtr, ByVal utc_number As LongPtr, ByVal utc_file As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_feof Lib "/usr/lib/libc.dylib" Alias "feof" _
    (ByVal utc_file As LongPtr) As LongPtr

#Else

' 32-bit Mac
Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" _
    (ByVal utc_command As String, ByVal utc_mode As String) As Long
Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" _
    (ByVal utc_file As Long) As Long
Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" _
    (ByVal utc_buffer As String, ByVal utc_size As Long, ByVal utc_number As Long, ByVal utc_file As Long) As Long
Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" _
    (ByVal utc_file As Long) As Long

#End If

#ElseIf VBA7 Then

' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#Else

Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#End If

#If Mac Then

#If VBA7 Then
Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As LongPtr
End Type

#Else

Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As Long
End Type

#End If

#Else

Private Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

#End If
' === End VBA-UTC

Private Type json_Options
    ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
    ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
    ' See: http://support.microsoft.com/kb/269370
    '
    ' By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
    ' to override set `JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
    UseDoubleForLargeNumbers As Boolean

    ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
    AllowUnquotedKeys As Boolean

    ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in ConvertToJson
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert JSON string to object (Dictionary/Collection)
'
' @method ParseJson
' @param {String} json_String
' @return {Object} (Dictionary or Collection)
' @throws 10001 - JSON parse error
''
Public Function ParseJson(ByVal JsonString As String) As Object
    Dim json_index As Long
    json_index = 1

    ' Remove vbCr, vbLf, and vbTab from json_String
    JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")

    json_SkipSpaces JsonString, json_index
    Select Case VBA.Mid$(JsonString, json_index, 1)
    Case "{"
        Set ParseJson = json_ParseObject(JsonString, json_index)
    Case "["
        Set ParseJson = json_ParseArray(JsonString, json_index)
    Case Else
        ' Error: Invalid JSON string
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_index, "Expecting '{' or '['")
    End Select
End Function

''
' Convert object (Dictionary/Collection/Array) to JSON
'
' @method ConvertToJson
' @param {Variant} JsonValue (Dictionary, Collection, or Array)
' @param {Integer|String} Whitespace "Pretty" print json with given number of spaces per indentation (Integer) or given string
' @return {String}
''
Public Function ConvertToJson(ByVal JsonValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal json_CurrentIndentation As Long = 0) As String
    Dim json_buffer As String
    Dim json_buffer_position As Long
    Dim json_buffer_length As Long
    Dim json_index As Long
    Dim json_LBound As Long
    Dim json_UBound As Long
    Dim json_IsFirstItem As Boolean
    Dim json_index2d As Long
    Dim json_LBound2D As Long
    Dim json_UBound2D As Long
    Dim json_IsFirstItem2D As Boolean
    Dim json_key As Variant
    Dim json_value As Variant
    Dim json_date_str As String
    Dim json_converted As String
    Dim json_skip_item As Boolean
    Dim json_pretty_print As Boolean
    Dim json_indentation As String
    Dim json_inner_indentation As String

    json_LBound = -1
    json_UBound = -1
    json_IsFirstItem = True
    json_LBound2D = -1
    json_UBound2D = -1
    json_IsFirstItem2D = True
    json_pretty_print = Not IsMissing(Whitespace)

    Select Case VBA.VarType(JsonValue)
    Case VBA.vbNull
        ConvertToJson = "null"
    Case VBA.vbDate
        ' Date
        json_date_str = ConvertToIso(VBA.CDate(JsonValue))

        ConvertToJson = """" & json_date_str & """"
    Case VBA.vbString
        ' String (or large number encoded as string)
        If Not JsonOptions.UseDoubleForLargeNumbers And json_StringIsLargeNumber(JsonValue) Then
            ConvertToJson = JsonValue
        Else
            ConvertToJson = """" & json_Encode(JsonValue) & """"
        End If
    Case VBA.vbBoolean
        If JsonValue Then
            ConvertToJson = "true"
        Else
            ConvertToJson = "false"
        End If
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        If json_pretty_print Then
            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
                json_inner_indentation = VBA.String$(json_CurrentIndentation + 2, Whitespace)
            Else
                json_indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
                json_inner_indentation = VBA.Space$((json_CurrentIndentation + 2) * Whitespace)
            End If
        End If

        ' Array
        json_buffer_append json_buffer, "[", json_buffer_position, json_buffer_length

        On Error Resume Next

        json_LBound = LBound(JsonValue, 1)
        json_UBound = UBound(JsonValue, 1)
        json_LBound2D = LBound(JsonValue, 2)
        json_UBound2D = UBound(JsonValue, 2)

        If json_LBound >= 0 And json_UBound >= 0 Then
            For json_index = json_LBound To json_UBound
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    ' Append comma to previous line
                    json_buffer_append json_buffer, ",", json_buffer_position, json_buffer_length
                End If

                If json_LBound2D >= 0 And json_UBound2D >= 0 Then
                    ' 2D Array
                    If json_pretty_print Then
                        json_buffer_append json_buffer, vbNewLine, json_buffer_position, json_buffer_length
                    End If
                    json_buffer_append json_buffer, json_indentation & "[", json_buffer_position, json_buffer_length

                    For json_index2d = json_LBound2D To json_UBound2D
                        If json_IsFirstItem2D Then
                            json_IsFirstItem2D = False
                        Else
                            json_buffer_append json_buffer, ",", json_buffer_position, json_buffer_length
                        End If

                        json_converted = ConvertToJson(JsonValue(json_index, json_index2d), Whitespace, json_CurrentIndentation + 2)

                        ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                        If json_converted = "" Then
                            ' (nest to only check if converted = "")
                            If json_IsUndefined(JsonValue(json_index, json_index2d)) Then
                                json_converted = "null"
                            End If
                        End If

                        If json_pretty_print Then
                            json_converted = vbNewLine & json_inner_indentation & json_converted
                        End If

                        json_buffer_append json_buffer, json_converted, json_buffer_position, json_buffer_length
                    Next json_index2d

                    If json_pretty_print Then
                        json_buffer_append json_buffer, vbNewLine, json_buffer_position, json_buffer_length
                    End If

                    json_buffer_append json_buffer, json_indentation & "]", json_buffer_position, json_buffer_length
                    json_IsFirstItem2D = True
                Else
                    ' 1D Array
                    json_converted = ConvertToJson(JsonValue(json_index), Whitespace, json_CurrentIndentation + 1)

                    ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                    If json_converted = "" Then
                        ' (nest to only check if converted = "")
                        If json_IsUndefined(JsonValue(json_index)) Then
                            json_converted = "null"
                        End If
                    End If

                    If json_pretty_print Then
                        json_converted = vbNewLine & json_indentation & json_converted
                    End If

                    json_buffer_append json_buffer, json_converted, json_buffer_position, json_buffer_length
                End If
            Next json_index
        End If

        On Error GoTo 0

        If json_pretty_print Then
            json_buffer_append json_buffer, vbNewLine, json_buffer_position, json_buffer_length

            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_indentation = VBA.String$(json_CurrentIndentation, Whitespace)
            Else
                json_indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
            End If
        End If

        json_buffer_append json_buffer, json_indentation & "]", json_buffer_position, json_buffer_length

        ConvertToJson = json_buffer_to_string(json_buffer, json_buffer_position)

    ' Dictionary or Collection
    Case VBA.vbObject
        If json_pretty_print Then
            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
            Else
                json_indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
            End If
        End If

        ' Dictionary
        If VBA.TypeName(JsonValue) = "Dictionary" Then
            json_buffer_append json_buffer, "{", json_buffer_position, json_buffer_length
            For Each json_key In JsonValue.Keys
                ' For Objects, undefined (Empty/Nothing) is not added to object
                json_converted = ConvertToJson(JsonValue(json_key), Whitespace, json_CurrentIndentation + 1)
                If json_converted = "" Then
                    json_skip_item = json_IsUndefined(JsonValue(json_key))
                Else
                    json_skip_item = False
                End If

                If Not json_skip_item Then
                    If json_IsFirstItem Then
                        json_IsFirstItem = False
                    Else
                        json_buffer_append json_buffer, ",", json_buffer_position, json_buffer_length
                    End If

                    If json_pretty_print Then
                        json_converted = vbNewLine & json_indentation & """" & json_key & """: " & json_converted
                    Else
                        json_converted = """" & json_key & """:" & json_converted
                    End If

                    json_buffer_append json_buffer, json_converted, json_buffer_position, json_buffer_length
                End If
            Next json_key

            If json_pretty_print Then
                json_buffer_append json_buffer, vbNewLine, json_buffer_position, json_buffer_length

                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                Else
                    json_indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                End If
            End If

            json_buffer_append json_buffer, json_indentation & "}", json_buffer_position, json_buffer_length

        ' Collection
        ElseIf VBA.TypeName(JsonValue) = "Collection" Then
            json_buffer_append json_buffer, "[", json_buffer_position, json_buffer_length
            For Each json_value In JsonValue
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    json_buffer_append json_buffer, ",", json_buffer_position, json_buffer_length
                End If

                json_converted = ConvertToJson(json_value, Whitespace, json_CurrentIndentation + 1)

                ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                If json_converted = "" Then
                    ' (nest to only check if converted = "")
                    If json_IsUndefined(json_value) Then
                        json_converted = "null"
                    End If
                End If

                If json_pretty_print Then
                    json_converted = vbNewLine & json_indentation & json_converted
                End If

                json_buffer_append json_buffer, json_converted, json_buffer_position, json_buffer_length
            Next json_value

            If json_pretty_print Then
                json_buffer_append json_buffer, vbNewLine, json_buffer_position, json_buffer_length

                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                Else
                    json_indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                End If
            End If

            json_buffer_append json_buffer, json_indentation & "]", json_buffer_position, json_buffer_length
        End If

        ConvertToJson = json_buffer_to_string(json_buffer, json_buffer_position)
    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
        ' Number (use decimals for numbers)
        ConvertToJson = VBA.Replace(JsonValue, ",", ".")
    Case Else
        ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
        ' Use VBA's built-in to-string
        On Error Resume Next
        ConvertToJson = JsonValue
        On Error GoTo 0
    End Select
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function json_ParseObject(json_string As String, ByRef json_index As Long) As Dictionary
    Dim json_key As String
    Dim json_next_char As String

    Set json_ParseObject = New Dictionary
    json_SkipSpaces json_string, json_index
    If VBA.Mid$(json_string, json_index, 1) <> "{" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_string, json_index, "Expecting '{'")
    Else
        json_index = json_index + 1

        Do
            json_SkipSpaces json_string, json_index
            If VBA.Mid$(json_string, json_index, 1) = "}" Then
                json_index = json_index + 1
                Exit Function
            ElseIf VBA.Mid$(json_string, json_index, 1) = "," Then
                json_index = json_index + 1
                json_SkipSpaces json_string, json_index
            End If

            json_key = json_ParseKey(json_string, json_index)
            json_next_char = json_Peek(json_string, json_index)
            If json_next_char = "[" Or json_next_char = "{" Then
                Set json_ParseObject.Item(json_key) = json_ParseValue(json_string, json_index)
            Else
                json_ParseObject.Item(json_key) = json_ParseValue(json_string, json_index)
            End If
        Loop
    End If
End Function

Private Function json_ParseArray(json_string As String, ByRef json_index As Long) As Collection
    Set json_ParseArray = New Collection

    json_SkipSpaces json_string, json_index
    If VBA.Mid$(json_string, json_index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_string, json_index, "Expecting '['")
    Else
        json_index = json_index + 1

        Do
            json_SkipSpaces json_string, json_index
            If VBA.Mid$(json_string, json_index, 1) = "]" Then
                json_index = json_index + 1
                Exit Function
            ElseIf VBA.Mid$(json_string, json_index, 1) = "," Then
                json_index = json_index + 1
                json_SkipSpaces json_string, json_index
            End If

            json_ParseArray.Add json_ParseValue(json_string, json_index)
        Loop
    End If
End Function

Private Function json_ParseValue(json_string As String, ByRef json_index As Long) As Variant
    json_SkipSpaces json_string, json_index
    Select Case VBA.Mid$(json_string, json_index, 1)
    Case "{"
        Set json_ParseValue = json_ParseObject(json_string, json_index)
    Case "["
        Set json_ParseValue = json_ParseArray(json_string, json_index)
    Case """", "'"
        json_ParseValue = json_ParseString(json_string, json_index)
    Case Else
        If VBA.Mid$(json_string, json_index, 4) = "true" Then
            json_ParseValue = True
            json_index = json_index + 4
        ElseIf VBA.Mid$(json_string, json_index, 5) = "false" Then
            json_ParseValue = False
            json_index = json_index + 5
        ElseIf VBA.Mid$(json_string, json_index, 4) = "null" Then
            json_ParseValue = Null
            json_index = json_index + 4
        ElseIf VBA.InStr("+-0123456789", VBA.Mid$(json_string, json_index, 1)) Then
            json_ParseValue = json_ParseNumber(json_string, json_index)
        Else
            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_string, json_index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function

Private Function json_ParseString(json_string As String, ByRef json_index As Long) As String
    Dim json_quote As String
    Dim json_char As String
    Dim json_code As String
    Dim json_buffer As String
    Dim json_buffer_position As Long
    Dim json_buffer_length As Long

    json_SkipSpaces json_string, json_index

    ' Store opening quote to look for matching closing quote
    json_quote = VBA.Mid$(json_string, json_index, 1)
    json_index = json_index + 1

    Do While json_index > 0 And json_index <= Len(json_string)
        json_char = VBA.Mid$(json_string, json_index, 1)

        Select Case json_char
        Case "\"
            ' Escaped string, \\, or \/
            json_index = json_index + 1
            json_char = VBA.Mid$(json_string, json_index, 1)

            Select Case json_char
            Case """", "\", "/", "'"
                json_buffer_append json_buffer, json_char, json_buffer_position, json_buffer_length
                json_index = json_index + 1
            Case "b"
                json_buffer_append json_buffer, vbBack, json_buffer_position, json_buffer_length
                json_index = json_index + 1
            Case "f"
                json_buffer_append json_buffer, vbFormFeed, json_buffer_position, json_buffer_length
                json_index = json_index + 1
            Case "n"
                json_buffer_append json_buffer, vbCrLf, json_buffer_position, json_buffer_length
                json_index = json_index + 1
            Case "r"
                json_buffer_append json_buffer, vbCr, json_buffer_position, json_buffer_length
                json_index = json_index + 1
            Case "t"
                json_buffer_append json_buffer, vbTab, json_buffer_position, json_buffer_length
                json_index = json_index + 1
            Case "u"
                ' Unicode character escape (e.g. \u00a9 = Copyright)
                json_index = json_index + 1
                json_code = VBA.Mid$(json_string, json_index, 4)
                json_buffer_append json_buffer, VBA.ChrW(VBA.Val("&h" + json_code)), json_buffer_position, json_buffer_length
                json_index = json_index + 4
            End Select
        Case json_quote
            json_ParseString = json_buffer_to_string(json_buffer, json_buffer_position)
            json_index = json_index + 1
            Exit Function
        Case Else
            json_buffer_append json_buffer, json_char, json_buffer_position, json_buffer_length
            json_index = json_index + 1
        End Select
    Loop
End Function

Private Function json_ParseNumber(json_string As String, ByRef json_index As Long) As Variant
    Dim json_char As String
    Dim json_value As String
    Dim json_is_large_number As Boolean

    json_SkipSpaces json_string, json_index

    Do While json_index > 0 And json_index <= Len(json_string)
        json_char = VBA.Mid$(json_string, json_index, 1)

        If VBA.InStr("+-0123456789.eE", json_char) Then
            ' Unlikely to have massive number, so use simple append rather than buffer here
            json_value = json_value & json_char
            json_index = json_index + 1
        Else
            ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
            ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
            ' See: http://support.microsoft.com/kb/269370
            '
            ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
            ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
            json_is_large_number = IIf(InStr(json_value, "."), Len(json_value) >= 17, Len(json_value) >= 16)
            If Not JsonOptions.UseDoubleForLargeNumbers And json_is_large_number Then
                json_ParseNumber = json_value
            Else
                ' VBA.Val does not use regional settings, so guard for comma is not needed
                json_ParseNumber = VBA.Val(json_value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function json_ParseKey(json_string As String, ByRef json_index As Long) As String
    ' Parse key with single or double quotes
    If VBA.Mid$(json_string, json_index, 1) = """" Or VBA.Mid$(json_string, json_index, 1) = "'" Then
        json_ParseKey = json_ParseString(json_string, json_index)
    ElseIf JsonOptions.AllowUnquotedKeys Then
        Dim json_char As String
        Do While json_index > 0 And json_index <= Len(json_string)
            json_char = VBA.Mid$(json_string, json_index, 1)
            If (json_char <> " ") And (json_char <> ":") Then
                json_ParseKey = json_ParseKey & json_char
                json_index = json_index + 1
            Else
                Exit Do
            End If
        Loop
    Else
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_string, json_index, "Expecting '""' or '''")
    End If

    ' Check for colon and skip if present or throw if not present
    json_SkipSpaces json_string, json_index
    If VBA.Mid$(json_string, json_index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_string, json_index, "Expecting ':'")
    Else
        json_index = json_index + 1
    End If
End Function

Private Function json_IsUndefined(ByVal json_value As Variant) As Boolean
    ' Empty / Nothing -> undefined
    Select Case VBA.VarType(json_value)
    Case VBA.vbEmpty
        json_IsUndefined = True
    Case VBA.vbObject
        Select Case VBA.TypeName(json_value)
        Case "Empty", "Nothing"
            json_IsUndefined = True
        End Select
    End Select
End Function

Private Function json_Encode(ByVal json_text As Variant) As String
    ' Reference: http://www.ietf.org/rfc/rfc4627.txt
    ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
    Dim json_index As Long
    Dim json_char As String
    Dim json_asc_code As Long
    Dim json_buffer As String
    Dim json_buffer_position As Long
    Dim json_buffer_length As Long

    For json_index = 1 To VBA.Len(json_text)
        json_char = VBA.Mid$(json_text, json_index, 1)
        json_asc_code = VBA.AscW(json_char)

        ' When AscW returns a negative number, it returns the twos complement form of that number.
        ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
        ' https://support.microsoft.com/en-us/kb/272138
        If json_asc_code < 0 Then
            json_asc_code = json_asc_code + 65536
        End If

        ' From spec, ", \, and control characters must be escaped (solidus is optional)

        Select Case json_asc_code
        Case 34
            ' " -> 34 -> \"
            json_char = "\"""
        Case 92
            ' \ -> 92 -> \\
            json_char = "\\"
        Case 47
            ' / -> 47 -> \/ (optional)
            If JsonOptions.EscapeSolidus Then
                json_char = "\/"
            End If
        Case 8
            ' backspace -> 8 -> \b
            json_char = "\b"
        Case 12
            ' form feed -> 12 -> \f
            json_char = "\f"
        Case 10
            ' line feed -> 10 -> \n
            json_char = "\n"
        Case 13
            ' carriage return -> 13 -> \r
            json_char = "\r"
        Case 9
            ' tab -> 9 -> \t
            json_char = "\t"
        Case 0 To 31
            ' Control characters -> convert to 4-digit hex
            json_char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_asc_code), 4)
        Case 127 To 65535
            ' Non-ascii: 한글 및 다국어 문자는 그대로 유지, 나머지만 이스케이프
            If json_asc_code < 128 Or _
                ' 한글 완성형
               (json_asc_code >= 44032 And json_asc_code <= 55203) Or _
               ' 한글 자모
               (json_asc_code >= 12593 And json_asc_code <= 12643) Or _
               ' CJK 한자
               (json_asc_code >= 19968 And json_asc_code <= 40959) Or _
               ' 괄호 문자 (㈜ 등)
               (json_asc_code >= 12800 And json_asc_code <= 13055) Then
                ' 유니코드 문자 그대로 출력
            Else
                json_char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_asc_code), 4)
            End If
        End Select

        json_buffer_append json_buffer, json_char, json_buffer_position, json_buffer_length
    Next json_index

    json_Encode = json_buffer_to_string(json_buffer, json_buffer_position)
End Function

Private Function json_Peek(json_string As String, ByVal json_index As Long, Optional json_number_of_characters As Long = 1) As String
    ' "Peek" at the next number of characters without incrementing json_index (ByVal instead of ByRef)
    json_SkipSpaces json_string, json_index
    json_Peek = VBA.Mid$(json_string, json_index, json_number_of_characters)
End Function

Private Sub json_SkipSpaces(json_string As String, ByRef json_index As Long)
    ' Increment index to skip over spaces
    Do While json_index > 0 And json_index <= VBA.Len(json_string) And VBA.Mid$(json_string, json_index, 1) = " "
        json_index = json_index + 1
    Loop
End Sub

Private Function json_StringIsLargeNumber(json_string As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See json_ParseNumber)

    Dim json_length As Long
    Dim json_char_index As Long
    json_length = VBA.Len(json_string)

    ' Length with be at least 16 characters and assume will be less than 100 characters
    If json_length >= 16 And json_length <= 100 Then
        Dim json_char_code As String

        json_StringIsLargeNumber = True

        For json_char_index = 1 To json_length
            json_char_code = VBA.Asc(VBA.Mid$(json_string, json_char_index, 1))
            Select Case json_char_code
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                json_StringIsLargeNumber = False
                Exit Function
            End Select
        Next json_char_index
    End If
End Function

Private Function json_ParseErrorMessage(json_string As String, ByRef json_index As Long, ErrorMessage As String)
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing JSON:
    ' {"abcde":True}
    '          ^
    ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['

    Dim json_start_index As Long
    Dim json_stop_index As Long

    ' Include 10 characters before and after error (if possible)
    json_start_index = json_index - 10
    json_stop_index = json_index + 10
    If json_start_index <= 0 Then
        json_start_index = 1
    End If
    If json_stop_index > VBA.Len(json_string) Then
        json_stop_index = VBA.Len(json_string)
    End If

    json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(json_string, json_start_index, json_stop_index - json_start_index + 1) & VBA.vbNewLine & _
                             VBA.Space$(json_index - json_start_index) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Sub json_buffer_append(ByRef json_buffer As String, _
                              ByRef json_append As Variant, _
                              ByRef json_buffer_position As Long, _
                              ByRef json_buffer_length As Long)
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Put "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp
    '
    ' and clsStringAppend from Philip Swannell
    ' https://github.com/VBA-tools/VBA-JSON/pull/82

    Dim json_append_length As Long
    Dim json_length_plus_position As Long

    json_append_length = VBA.Len(json_append)
    json_length_plus_position = json_append_length + json_buffer_position

    If json_length_plus_position > json_buffer_length Then
        ' Appending would overflow buffer, add chunk
        ' (double buffer length or append length, whichever is bigger)
        Dim json_added_length As Long
        json_added_length = IIf(json_append_length > json_buffer_length, json_append_length, json_buffer_length)

        json_buffer = json_buffer & VBA.Space$(json_added_length)
        json_buffer_length = json_buffer_length + json_added_length
    End If

    ' Note: Namespacing with VBA.Mid$ doesn't work properly here, throwing compile error:
    ' Function call on left-hand side of assignment must return Variant or Object
    Mid$(json_buffer, json_buffer_position + 1, json_append_length) = CStr(json_append)
    json_buffer_position = json_buffer_position + json_append_length
End Sub

Private Function json_buffer_to_string(ByRef json_buffer As String, ByVal json_buffer_position As Long) As String
    If json_buffer_position > 0 Then
        json_buffer_to_string = VBA.Left$(json_buffer, json_buffer_position)
    End If
End Function

''
' VBA-UTC v1.0.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
'
' UTC/ISO 8601 Converter for VBA
'
' Errors:
' 10011 - UTC parsing error
' 10012 - UTC conversion error
' 10013 - ISO 8601 parsing error
' 10014 - ISO 8601 conversion error
'
' @module UtcConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Declarations moved to top)

' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @method ParseUtc
' @param {Date} UtcDate
' @return {Date} Local date
' @throws 10011 - UTC parsing error
''
Public Function ParseUtc(utc_UtcDate As Date) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ParseUtc = utc_ConvertDate(utc_UtcDate)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_LocalDate As utc_SYSTEMTIME

    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate

    ParseUtc = utc_SystemTimeToDate(utc_LocalDate)
#End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to UTC date
'
' @method ConvertToUrc
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' @throws 10012 - UTC conversion error
''
Public Function ConvertToUtc(utc_LocalDate As Date) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_UtcDate As utc_SYSTEMTIME

    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate

    ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
#End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.Number & " - " & Err.Description
End Function

''
' Parse ISO 8601 date string to local date
'
' @method ParseIso
' @param {Date} utc_IsoString
' @return {Date} Local date
' @throws 10013 - ISO 8601 parsing error
''
Public Function ParseIso(utc_IsoString As String) As Date
    On Error GoTo utc_ErrorHandling

    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String
    Dim utc_OffsetIndex As Long
    Dim utc_HasOffset As Boolean
    Dim utc_NegativeOffset As Boolean
    Dim utc_OffsetParts() As String
    Dim utc_Offset As Date

    utc_Parts = VBA.Split(utc_IsoString, "T")
    utc_DateParts = VBA.Split(utc_Parts(0), "-")
    ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))

    If UBound(utc_Parts) > 0 Then
        If VBA.InStr(utc_Parts(1), "Z") Then
            utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", ""), ":")
        Else
            utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
            If utc_OffsetIndex = 0 Then
                utc_NegativeOffset = True
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
            End If

            If utc_OffsetIndex > 0 Then
                utc_HasOffset = True
                utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
                utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), Len(utc_Parts(1)) - utc_OffsetIndex), ":")

                Select Case UBound(utc_OffsetParts)
                Case 0
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
                Case 1
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
                Case 2
                    ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), Int(VBA.Val(utc_OffsetParts(2))))
                End Select

                If utc_NegativeOffset Then: utc_Offset = -utc_Offset
            Else
                utc_TimeParts = VBA.Split(utc_Parts(1), ":")
            End If
        End If

        Select Case UBound(utc_TimeParts)
        Case 0
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
        Case 1
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
        Case 2
            ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), Int(VBA.Val(utc_TimeParts(2))))
        End Select

        ParseIso = ParseUtc(ParseIso)

        If utc_HasOffset Then
            ParseIso = ParseIso - utc_Offset
        End If
    End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & utc_IsoString & ": " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
''
Public Function ConvertToIso(utc_LocalDate As Date) As String
    On Error GoTo utc_ErrorHandling

    ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")

    Exit Function

utc_ErrorHandling:
    Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description
End Function

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Then

Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date
    Dim utc_ShellCommand As String
    Dim utc_Result As utc_ShellResult
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String

    If utc_ConvertToUtc Then
        utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
            " +'%s'` +'%Y-%m-%d %H:%M:%S'"
    Else
        utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _
            "+'%Y-%m-%d %H:%M:%S'"
    End If

    utc_Result = utc_ExecuteInShell(utc_ShellCommand)

    If utc_Result.utc_output = "" Then
        Err.Raise 10015, "UtcConverter.utc_ConvertDate", "'date' command failed"
    Else
        utc_Parts = Split(utc_Result.utc_output, " ")
        utc_DateParts = Split(utc_Parts(0), "-")
        utc_TimeParts = Split(utc_Parts(1), ":")

        utc_ConvertDate = DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _
            TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))
    End If
End Function

Private Function utc_ExecuteInShell(utc_ShellCommand As String) As utc_ShellResult
#If VBA7 Then
    Dim utc_file As LongPtr
    Dim utc_read As LongPtr
#Else
    Dim utc_file As Long
    Dim utc_read As Long
#End If

    Dim utc_chunk As String

    On Error GoTo utc_ErrorHandling
    utc_file = utc_popen(utc_ShellCommand, "r")

    If utc_file = 0 Then: Exit Function

    Do While utc_feof(utc_file) = 0
        utc_chunk = VBA.Space$(50)
        utc_read = CLng(utc_fread(utc_chunk, 1, Len(utc_chunk) - 1, utc_file))
        If utc_read > 0 Then
            utc_chunk = VBA.Left$(utc_chunk, CLng(utc_read))
            utc_ExecuteInShell.utc_output = utc_ExecuteInShell.utc_output & utc_chunk
        End If
    Loop

utc_ErrorHandling:
    utc_ExecuteInShell.utc_exit_code = CLng(utc_pclose(utc_file))
End Function

#Else

Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
    utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value)
    utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value)
    utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)
    utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
    utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
    utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
    utc_DateToSystemTime.utc_wMilliseconds = 0
End Function

Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
    utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
        TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
End Function

#End If
