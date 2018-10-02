<%
' Essential functions, grouped by functionality.
' Redundancy in their code is intentional, to avoid increase of the call stack.


' Check existence
    ' Checks if a function does exist in global scope.
    ' @param {string} function_name
    ' @return {bool}
    Function function_exists( function_name )
        Dim procedure

        call Err.Clear()
        On Error Resume Next
            Set procedure = GetRef(function_name)
        On Error Goto 0

        function_exists = (Err.Number = 0)
    End Function
    ' Checks if the object has a method.
    ' @param {object} object_
    ' @param {string} method_
    ' @return {bool}
    Function method_exists( object_, method_)
        Dim procedure

        call Err.Clear()
        On Error Resume Next
            Execute "Set procedure =  GetRef(object_. " & method_ & ") "
        On Error Goto 0

        method_exists = (Err.Number = 0)
    End Function
    ' Checks if the object has a property.
    ' @param {object} object_
    ' @param {string} property_
    ' @return {bool}
    Function property_exists( object_, property_)
        Dim value

        call Err.Clear()
        On Error Resume Next
            Execute "set_ value, object_." & property_
        On Error Goto 0

        property_exists = (Err.Number = 0)
    End Function


' Check type
    ' Checks if the object is fo give class.
    ' @param {object} object_
    ' @param {string} class_
    ' @return {bool}
    Function is_a(object_, class_)
        is_a = (typename(object_) = class_)
    End Function



' Set value
    ' Stores a value in a reference, being it scalar or object. (meant for function return)
    ' @param {function} procedure
    ' @param {mixed} value
    Sub return(byref procedure, value)
        if IsObject(value) then
            Set procedure = value
        else
            procedure = value
        end if
    End Sub
    ' Stores a value in a reference, being it scalar or object.
    ' Does not work with Dictionary keys.
    ' @param {function} procedure
    ' @param {mixed} value
    Sub set_(byref reference, value)
        if IsObject(value) then
            Set reference = value
        else
            reference = value
        end if
    End Sub
' Seach collection
    ' Checks if $haystack collection has $needle value.
    ' @param {mixed} needle
    ' @param {array|Dictionary|IRequestDictionary|ISessionObject|IApplicationObject} haystack
    ' @return {bool}
    Function in_array(needle, haystack)
        Select Case TypeName(haystack)
            Case "Variant()" ' Array
                Dim Size : Size = UBound(haystack)
                Select Case Size
                    Case -1
                        in_array = false
                    Case 0
                        in_array = (haystack(0) = needle)
                    Case 1
                        Dim BeginIndex
                        Dim EndIndex
                        BeginIndex = 0
                        EndIndex = UBound(haystack)

                        Do Until EndIndex < BeginIndex
                            in_array = (haystack(EndIndex) = needle or haystack(BeginIndex) = needle)
                            if (in_array or EndIndex <= BeginIndex) then
                                Exit Do
                            else
                                BeginIndex = BeginIndex + 1
                                EndIndex = EndIndex - 1
                            end if
                        Loop
                End Select
            Case "Dictionary"
                in_array = haystack.Exists(needle)
            Case "IRequestDictionary", "ISessionObject", "IApplicationObject"
            ' Request.[Cookies, Form, Querystring, ServerVariables], Response.Cookies, Session, Application
                in_array = IsEmpty(haystack(needle))
        End Select
    End Function
%>