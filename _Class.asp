<!-- #include file="_Functions.asp" -->
<%
' Encapsulates the metadata of a class to provide some reflection capabilities.
' Also provides hability to set static fields on the class.
Class Reflection_Class
    ' @var {string}
    Private Class_Name
    ' @var {bool}
    Private Initialized
    ' @var {Scripting.Dictionary}
    Private Instance_Fields
    ' @var {Object}
    Private Reference_Instance
    ' @var {Scripting.Dictionary}
    Private Static_Fields



    ' Access a static field.
    '
    ' @param {string} Field_
    ' @return {mixed}
    Public Default Property Get Field( Field_ )
        ' Avoiding to search twice in the dictionary to check if is object
        return Field, Static_Fields(Field_)
    End Property
    ' Update a static field.
    '
    ' @param {string} Field_
    ' @param {mixed} Value
    Public Property Let Field( Field_, Value )
        if IsObject(Value) then
            Set Static_Fields(Field_) = Value
        else
            Static_Fields(Field_) = Value
        end if
    End Property



    ' Checks if this class is already initialized.
    '
    ' @return {bool}
    Public Property Get IsInitialized( )
        IsInitialized = Initialized
    End Property
    ' Marks this class as initialized, if it's not already.
    '
    ' @param {bool} Value
    ' @return {bool}
    Public Property Let IsInitialized( Value )
        if not Initialized and Value then
            Initialized = true

            Execute "Set Reference_Instance = new " & Class_Name
        end if
    End Property



    ' Gets the type of a instance member.
    '
    ' @param {string} Member_
    ' @return {string}
    Public Property Get Member( Member_ )
        ' Never will be an object
        Member = Instance_Fields(Member_)
    End Property
    ' Sets the type of a instance member.
    '
    ' @param {string} Member_
    ' @param {mixed} Type_
    ' @return {string}
    Public Property Let Member( Member_, Type_ )
        Instance_Fields(Member_) = Type_
    End Property



    ' Gets the class name.
    '
    ' @return {string|Empty}
    Public Property Get Name( )
        ' Never will be an object too
        Name = Class_Name
    End Property
    ' Sets the class name.
    '
    ' @param {mixed} Value
    ' @return {string|Empty}
    Public Property Let Name( Value )
        if not Initialized then
            if TypeName(Value) = "String" then
                Class_Name = Value
            else
                Err.Raise 13, "Invalid class name" & VbCrLf &_
                    "How did you get here?"
            end if
        else
            Err.Raise 13, "Cannot change class name in Runtime"
        end if
    End Property



    ' Initializer
        ' Simple initializer to create the dictionaries.
        Sub Class_Initialize()
            set Static_Fields  = CreateObject("Scripting.Dictionary")
            set Instance_Fields = CreateObject("Scripting.Dictionary")

            Initialized = false
        End Sub



    ' Destructor
        ' Simple destructor to erase the dictionaries.
        Sub Class_Terminate()
            Static_Fields.RemoveAll()
            Set Static_Fields = Nothing
            Static_Fields = Empty

            Instance_Fields.RemoveAll()
            Set Instance_Fields = Nothing
            Instance_Fields = Empty
        End Sub



    ' Reflection functions
        ' Builds an Entity of this class with the default values.
        '
        ' @return {Object}
        Public Function GetDefault()
            Dim Result
            Execute "Set Result = new " & Class_Name

            For Each Key in Instance_Fields
                Result.Add Key, Reference_Instance(Key)
            Next

            Set GetDefault = Result
        End Function
        ' Builds a new Entity of this class.
        '
        ' @return {Object}
        Public Function GetInstance()
            Execute "Set GetInstance = new " & Class_Name
        End Function
        ' Builds a Dictionary containing all instance fields names (keys) and
        ' types (values).
        '
        ' @return {Scripting.Dictionary}
        Public Function GetMembers()
            Dim Result : Set Result = CreateObject("Scripting.Dictionary")
            Dim Key

            For Each Key in Instance_Fields
                Result.Add Key, Instance_Fields(Key)
            Next

            Set GetMembers = Result
        End Function
End Class



' Implementation of loader for Entities backed by Reflection_Class.
' Provides 2 methods of loading classes:
' * Lazy loading
'   class is loaded only after the first object is created from it (default)
' * Active loading
'   class is loaded immediatily
Class Reflection_Class_Loader

    ' Loaded classes.
    '
    ' @var {Scripting.Dictionary}
    Private Classes
    ' Cache for member of a class.
    ' Simple performance optimization to deal with entity arrays.
    '
    ' @var {Scripting.Dictionary}
    Private Cache_ClassMember
    ' Name of the last class with chached members.
    '
    ' @var {string}
    Private Cache_LastClass



    ' Gets the Reflection_Class object associated with given Class_ name.
    ' If there's no object associated yet, performs Lazy Class Loading, wich is
    ' only activated on the initialization of the first object of the class.
    '
    ' @param {string} Class_
    ' @return {Reflection_Class}
    Public Default Property Get LoadClass( ByVal Class_ )
        if not Classes.Exists(Class_) then
            Set Classes(Class_) = new Reflection_Class
        end if

        Set LoadClass = Classes(Class_)
    End Property



    ' Initializer
        ' Simple initializer to create the Classes dictionary.
        Sub Class_Initialize()
            set Classes = CreateObject("Scripting.Dictionary")
        End Sub



    ' Destructor
        ' Simple destructor to erase the Classes dictionary.
        Sub Class_Terminate()
            Classes.RemoveAll()
            Set Classes = Nothing
            Classes = Empty
        End Sub



    ' Import
        ' Creates/feeds Entities with data present on given Source.
        '
        ' @param {string|Reflection_Class|Object} Entity
        ' @param {Scripting.Dictionary} Source
        ' @return {Object}
        Public Function FromDictionary( Entity, ByVal Source )
            Dim Result

            Dim Class_Object

            Select Case TypeName(Entity)
                Case "Reflection_Class"
                    Set Class_Object = Entity
                    Set Result = Class_Object.GetInstance()
                Case "String"
                    Set Class_Object = Load(Entity)
                    Set Result = Class_Object.GetInstance()
                Case Else
                    if IsObject(Entity) then
                        if property_exists( Entity, "SupportsReflection") then
                            set Result = Entity
                            set Class_Object = Result.Self
                        else
                            Err.Raise 13, "Invalid Entity"
                        end if
                    end if
            End Select

            if IsEmpty(Source) then
                Set FromDictionary = Nothing
            else
                Dim Key
                For Each Key in Members(Class_Object)
                    if Source.Exists(Key) then
                        Result(Key) = Source(Key)
                    end if
                Next

                Set FromDictionary = Result
            end if
        End Function
        ' Creates/feeds Entities with data present on given Source.
        '
        ' @param {string|Reflection_Class|Object} Entity
        ' @param {JSONobject|JSONarray|string} Source
        ' @return {Object|Object[]}
        Public Function FromJSON( Entity, ByVal Source )
            Dim Result
            Dim Class_Object

            Select Case TypeName(Entity)
                Case "Reflection_Class"
                    Set Class_Object = Entity
                    Set Result = Class_Object.GetInstance()
                Case "String"
                    Set Class_Object = Load(Entity)
                    Set Result = Class_Object.GetInstance()
                Case Else
                    if IsObject(Entity) then
                        if property_exists( Entity, "SupportsReflection") then
                            set Result = Entity
                            set Class_Object = Result.Self
                        else
                            Err.Raise 13, "Invalid Entity"
                        end if
                    end if
            End Select

            if IsEmpty(Result) then
                Set FromJSON = Nothing
            else
                Dim Index
                Dim Key
                Dim JSON
                Select Case TypeName(Source)
                    ' Avoiding unecessary increase of the call stack
                    Case "JSONarray"
                        Result = Array(Result)
                        Index = Source.length - 1
                        Redim Preserve Result(Index)

                        ' Avoiding new object (and inloop verification)
                        Set JSON = Source(0)
                        For Each Key in Members(Class_Object)
                            ' There's no way to know if the property exists in ASPJson
                            Result(0)(Key) = JSON(Key)
                        Next

                        For Index = Index To 1 Step -1
                            Set JSON = Source(Index)
                            Set Result(Index) = Class_Object.GetInstance()
                            For Each Key in Members(Class_Object)
                                Result(Index)(Key) = JSON(Key)
                            Next
                        Next
                        FromJSON = Result
                    Case "JSONobject"
                        For Each Key in Members(Class_Object)
                            ' There's no way to know if the property exists in ASPJson
                            Result(Key) = Source(Key)
                        Next
                        Set FromJSON = Result
                    Case "String"
                        ' Avoiding excessive code duplication
                        Set JSON = (new JSONobject).parse(Source)
                        set_ FromJSON, FromJSON( Entity, JSON )
                End Select
            end if
        End Function
        ' Creates/feeds Entities with data present on given request Method.
        ' Uses giver Prefix to identify fields names.
        '
        ' @param {string|Reflection_Class|Object} Entity
        ' @param {string} Method [Form|Post|Querystring|Get]
        ' @param {string} Prefix
        ' @return {Object|Object[]}
        Public Function FromRequest( Entity, ByVal Method, ByVal Prefix )
            Dim Result

            Dim Class_Object
            Dim Source

            Select Case TypeName(Entity)
                Case "Reflection_Class"
                    Set Class_Object = Entity
                    Set Result = Class_Object.GetInstance()
                Case "String"
                    Set Class_Object = Load(Entity)
                    Set Result = Class_Object.GetInstance()
                Case Else
                    if IsObject(Entity) then
                        if property_exists( Entity, "SupportsReflection") then
                            set Result = Entity
                            set Class_Object = Result.Self
                        else
                            Err.Raise 13, "Invalid Entity"
                        end if
                    end if
            End Select

            Select Case UCase(Method)
                Case "POST", "FORM":
                    Set Source = Request.Form
                Case Else
                    Set Source = Request.QueryString
            End Select

            if TypeName(Prefix) <> "String" then
                Prefix = ""
            end if

            Dim Key
            ' Getting only first element to check whole list type
            For Each Key in Members(Class_Object)
                Key = Prefix & Key
                Exit For
            Next

            Select Case Source(Key).Count
                Case 0
                    Set FromRequest = Nothing
                Case 1
                    For Each Key in Members(Class_Object)
                        Key = Prefix & Key
                        if not IsEmpty(Source(Key)) then
                            Result(Key) = Source(Key)
                        end if
                    Next
                    Set FromRequest = Result
                Case Else
                    Result = Array(Result)
                    ' Avoiding new object (and inloop verification)
                    For Each Key in Members(Class_Object)
                        Key = Prefix & Key
                        Result(0)(Key) = Source(1)(Key)
                    Next

                    For Index = Source(Key).Count - 1 To 1 Step -1
                        Set Result(Index) = Class_Object.GetInstance()
                        For Each Key in Members(Class_Object)
                            Key = Prefix & Key
                            Result(Index)(Key) = Source(Index + 1)(Key)
                        Next
                    Next
                    FromRequest = Result
            End Select
        End Function
        ' Creates/feeds Entities with a JSON string present on session Key.
        '
        ' @param {string|Reflection_Class|Object} Entity
        ' @param {string} Key
        ' @return {Object}
        Public Function FromSession( Entity, Key )
            if IsEmpty(Session(Key)) then
                Set FromSession = Nothing
            else
                Set FromSession = FromJSON(Entity, Session(Key))
            end if
        End Function
        ' Creates/feeds Entities with data present on given Source.
        '
        ' @param {string|Reflection_Class|Object} Entity
        ' @param {string} Source
        ' @return {Object}
        Public Function Fromstring( Entity, ByVal Source )
            Dim JSON : set JSON = (new JSONobject).parse(Source)

            set_ Fromstring, FromJSON(Entity, JSON)
        End Function
    ' Export
        ' Exports an Entity to a Dictionary.
        '
        ' @param {Object} Entity
        ' @return {Scripting.Dictionary}
        Public Function ToDictionary( ByRef Entity )
            Dim Result
            if IsObject(Entity) then
                if property_exists( Entity, "SupportsReflection") and TypeName(Entity) <> "Reflection_Class" then
                    Dim Key

                    Set Result = CreateObject("Scripting.Dictionary")
                    For Each Key in Members(Entity.Self)
                        Result.Add Key, Entity(Key)
                    Next
                else
                    Err.Raise 13, "Invalid Entity"
                end if
            end if

            if IsObject(Result) then
                Set ToDictionary = Result
            else
                Set ToDictionary = Nothing
            end if
        End Function
        ' Exports Entities to a JSONobject or a JSONarray.
        '
        ' @param {Object|Object[]} Entity
        ' @return {JSONarray|JSONobject}
        Public Function ToJSON( ByRef Entity )
            Dim Result
            Dim Key

            if IsObject(Entity) then
                if property_exists(Entity, "SupportsReflection") and TypeName(Entity) <> "Reflection_Class" then

                    Set Result = new JSONobject
                    For Each Key in Members(Entity.Self)
                        Result.Add Key, Entity(Key)
                    Next
                else
                    Err.Raise 13, "Invalid Entity or Class"
                end if
            elseif IsArray(Entity) then
                Set Result = new JSONarray

                if UBound(Entity) > -1 then
                    Dim Current
                    Dim JSON
                    For Each Current in Entity
                        if not property_exists(Current, "SupportsReflection") or TypeName(Current) = "Reflection_Class" then
                            Err.Raise 13, "Invalid Entity or Class"
                        end if

                        Set JSON = new JSONobject
                        For Each Key in Members(Current.Self)
                            JSON.Add Key, Current(Key)
                        Next
                        Result.push JSON
                    Next
                end if
            end if

            if IsObject(Result) then
                Set ToJSON = Result
            else
                Set ToJSON = Nothing
            end if
        End Function
        ' Exports an Entities to a JSON String.
        '
        ' @param {Object|Object[]} Entity
        ' @return {String}
        Public Function ToString( ByRef Entity )
            Dim Result : Set Result = ToJSON(Entity)

            if Result Is Nothing then
                Tostring = Null
            else
                Tostring = Result.Serialize()
            end if
        End Function



    ' Utilitary
        ' Gets the instance members of Reflection_Class from the cache.
        ' If there's no cached members object yet, loads it from the class,
        '
        ' @param {Reflection_Class} Class_
        ' @return {Scripting.Dictionary}
        Public Function Members( ByRef Class_ )
            if Cache_LastClass <> Class_.Name then
                Set Cache_ClassMember = Class_.GetMembers()
            end if

            Set Members = Cache_ClassMember
        End Function
        ' Gets the Reflection_Class object associated with given Class_ name.
        ' If there's no object associated yet, performs Active Class Loading,
        ' doing the initialization of the first object immediatily.
        '
        ' @param {string} Class_
        ' @return {Reflection_Class}
        Public Function Load( ByVal Class_ )
            Dim Result : Set Result = LoadClass(Class_)
            Dim Instance

            if not Result.IsInitialized then
                ' Forcing the loading to happen right now
                Execute "Set Instance = new " & Class_
            end if

            Set Load = Result
        End Function

End Class

' Default instance of Reflection_Class_Loader
Dim Class_Loader
Set Class_Loader = new Reflection_Class_Loader
%>
