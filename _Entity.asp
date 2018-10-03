<%
' Reference to the Reflection_Class backing this entity.
'
' @var {Reflection_Class}
Private Class_


' Gets the value of a field on the Entity, acting like a string-keyed indexer.
'
' @param {string} Field_
' @return {mixed}
Public Default Property Get Field( Field_ )
    Dim Key : Key = "Me." & Field_

    Execute "if IsObject(" & Key & ") then" &_
        " set Field = " & Key &_
    " else" &_
        " Field = " & Key &_
    " end if"
End Property
' Sets the value of a field on the Entity, acting like a string-keyed indexer.
'
' @param {string} Field_
' @param {mixed} Value
Public Property Let Field( Field_, Value )
    Dim Key : Key = "Me." & Field_

    Execute "if IsObject(Value) then" &_
        " set " & Key & " = Value" &_
    " else" &_
        " " & Key & " = Value" &_
    " end if"
End Property



' Getter of the Reflection_Class backing this entity.
'
' @return {Reflection_Class}
Public Property Get Self( )
    Set Self = Class_
End Property

' Indicates that this Entity supports reflection.
'
' @return {bool(true)}
Public Property Get SupportsReflection( )
    SupportsReflection = true
End Property

' Alias for *Me* property on the Entity.
'
' @return {bool(true)}
Private Property Get This( )
    Set This = Me
End Property



' Initializer
    ' Initialize current Entity, also initializing the class if needed.
    Sub Class_Initialize()
        Dim ClassName : ClassName = TypeName(Me)

        set Class_ = Class_Loader(ClassName)
        if not Class_.IsInitialized then
            if method_exists(Me, "Static_Initialize") then
                call Static_Initialize()
            end if

            Class_.Name = ClassName
            Class_.IsInitialized = true
        end if

        if method_exists(Me, "Instance_Initialize") then
            call Instance_Initialize()
        end if
    End Sub

' Destructor
    ' Destroys current Entity.
    Sub Class_Terminate()
        Set Class_ = Nothing
        Class_ = Empty
    End Sub



' Import
    ' Creates/feeds Entities with data present on given Source.
    '
    ' @param {Scripting.Dictionary} Source
    ' @return {Object}
    Public Function FromDictionary(Source)
        Call Class_Loader.FromDictionary(Me, Source)

        Set FromDictionary = Me
    End Function
    ' Creates/feeds Entities with data present on given Source.
    '
    ' @param {JSONobject|JSONarray|string} Source
    ' @return {Object|Object[]}
    Public Function FromJSON(Source)
        set_ FromJSON, Class_Loader.FromJSON(Me, Source)
    End Function
    ' Creates/feeds Entities with data present on given request Method.
    ' Uses giver Prefix to identify fields names.
    '
    ' @param {string} Method [Form|Post|Querystring|Get]
    ' @return {Object}
    Public Function FromRequest(Method)
        Call Class_Loader.FromRequest(Me, Method, "")

        Set FromRequest = Me
    End Function
    ' Creates/feeds Entities with a JSON string present on session Key.
    '
    ' @param {string} Key
    ' @return {Object|Object[]}
    Public Function FromSession(Key)
        Call Class_Loader.FromSession(Me, Key)

        Set FromSession = Me
    End Function
    ' Creates/feeds Entities with data present on given Source.
    '
    ' @param {string} Source
    ' @return {Object|Object[]}
    Public Function FromString(Source)
        set_ FromString, Class_Loader.FromString(Me, Source)
    End Function



' Export
    ' Exports this Entity to a Dictionary.
    '
    ' @return {Scripting.Dictionary}
    Public Property Get AsDictionary()
        Set AsDictionary = Class_Loader.ToDictionary(Me)
    End Property
    ' Exports this Entity to a JSONobject.
    '
    ' @return {JSONobject}
    Public Property Get AsJSON()
        Set AsJSON = Class_Loader.ToJSON(Me)
    End Property
    ' Exports this Entity to a JSON String.
    '
    ' @return {String}
    Public Property Get AsString()
        AsString = Class_Loader.ToString(Me)
    End Property
%>