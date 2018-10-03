# VBScript-Reflect
This library provides a simple implementation of reflection using ASP Classes, allowing you to use classes more in an Object Oriented than the strcutural style of ASP.

You may, optionaly, use **ASPJson** in your project to enable the JSON exportation capabilities of this library.



# Project structure

The file-structure of this library is quite simple, there-'s only 3 code files - and you need to care about only 2 of them.

* **_Class.asp**  
Provides the Class_Loader and the ReflectionClass class, wich is essential to use our reflection capabilites
* **_Entity.asp**  
Provides a standart extension for normal ASP Classes to use our reflection library
* *_Functions.asp*  
Utilitary functions used internaly by the library and avaliable for users



# Using in your project

You must include *_Class.asp* on your page/application before using your classes. 
The library is made to provide lazy loading of class metadata - so you may load just the specs of classes you make use, avoiding overhead.

Along with that, you must include *_Entity.asp* on every ASP Class that you want to have reflection capabilities.



## Workflow

When *_Class.asp* is included, the default Class_Loader is initialized and ready to work.

Later, when the first object of a class marked to have reflection is initialized (created with the *new* operator) it's class is loaded by the Class_Loader - a lazy loading of class metadata.  
All objects implementing *_Entity* include are called Entities.

In that step, a *Reflection_Class* object containg the metadata of the class is created and stored in memory - being accessible by entities with *Self* field and by Class_Loader("{class-name}") call.



## Reflection_Class object

This object encapsulates the following data:

* **Class Name**  
    *Name property*, automatically got with a TypeName() call on the firt entity initializer
* **If this Reflection_Class is initialized**  
    *IsInitialized property*, wich stays false until the first entity is initialized.  
    Manually setting this property (no matter what value is given) will force this Reflection_Class to be initialized.
* **Dicitionary listing instance fields**  
    Names of fields must be stored on Dictionary keys and field types on Dictionary values - if you want to specify types.  
    All fields listed here will be visible by reflection, any field not listed wont.
* **Static fields**  
    Dicitionary that allows registering arbitrary data directly on the Reflection_Class - acting exactly as static fields on really Object Oriented languages.
* **Reference instance**  
    Standard entity of the class with default values on each property marked to reflection.


The followign properties and methods are accessible to anyone (public):

* *mixed* **Field**  
    Gets/sets static fields
* *bool* **IsInitialized**  
    If the class is already initialized
* *mixed* **Member**  
    Gets/sets instance member types
* *string* **Name**  
    Class name
* *Entity* **GetDefault**()  
    Builds an Entity of this class with the default values
* *Entity* **GetInstance**()  
    Builds a new Entity of this class
* *Scripting.Dictionary* **GetMembers**()  
    Builds a Dictionary containing all instance fields names (keys) and types (values)



## Reflection_Class_Loader Object

Provides the class loading functionality, having a default instance on *Class_Loader* variable.

Provides the following properties and functions:

* *Reflection_Class* **LoadClass**(*string* Class_)  
    Gets the Reflection_Class object associated with given Class_ name, doing Active Class Loading if the class is not initialized
* *Reflection_Class* **Load**(*string* Class_)  
    Gets the Reflection_Class object associated with given Class_ name, doing Lazy Class Loading if the class is not initialized
* *Scripting.Dictionary* **Members**(*Reflection_Class* Class_)  
    Gets the instance members of Reflection_Class from the member cache, loading them from the class if they are not cached
* *Import*
    * *Entity* **FromDictionary**(*string|Reflection_Class|Entity* Entity, *Scripting.Dictionary* Source)  
        Creates/feeds Entities with data present on given Source
    * *Entity|Entity[]* **FromJSON**(*string|Reflection_Class|Entity* Entity, *JSONobject|JSONarray|string* Source)  
        Creates/feeds Entities with data present on given Source  
        ***REQUIRES ASPJSON***
    * *Entity* **FromRequest**(*string|Reflection_Class|Entity* Entity, *string* Method, *string* Prefix)  
        Creates/feeds Entities with data present on given request Method, using given Prefix to identify fields names
    * *Entity|Entity[]* **FromSession**(*string|Reflection_Class|Entity* Entity, *string* Key)  
        Creates/feeds Entities with a JSON string present on session Key  
        ***REQUIRES ASPJSON***
    * *Entity|Entity[]* **Fromstring**(*string|Reflection_Class|Entity* Entity, *string* Source)  
        Creates/feeds Entities with data present on given Source  
        ***REQUIRES ASPJSON***
* *Export*
    * *Scripting.Dictionary* **FromDictionary**(*Entity* Entity)  
        Exports an Entity to a Dictionary
    * *JSONarray|JSONobject* **ToJSON**(*Entity|Entity[]* Entity)  
        Exports Entities to a JSONobject or a JSONarray
        ***REQUIRES ASPJSON***
    * *String* **ToString**(*Entity|Entity[]* Entity)  
        Exports Entities to a JSON String
        ***REQUIRES ASPJSON***



## _Entity include

Provides the reflection capability to ASP Classes, adding to it the following public properties and functions:

* *mixed* **Field**(*string* Field_)  
    Gets/sets the value of a field on the Entity, acting like a string-keyed indexer.  
    Can access any field on the class.
* *Reflection_Class* **Self**  
    Gets the Reflection_Class associated with this Entity.
* *bool(true)* **SupportsReflection**  
    Indicates that this Entity supports reflection.
* *Import*
    * *Entity* **FromDictionary**(*Scripting.Dictionary* Source)  
        Creates/feeds Entities with data present on given Source
    * *Entity|Entity[]* **FromJSON**(*JSONobject|JSONarray|string* Source)  
        Creates/feeds Entities with data present on given Source  
        ***REQUIRES ASPJSON***
    * *Entity* **FromRequest**(*string* Method)  
        Creates/feeds Entities with data present on given request Method, using given Prefix to identify fields names
    * *Entity|Entity[]* **FromSession**(*string* Key)  
        Creates/feeds Entities with a JSON string present on session Key  
        ***REQUIRES ASPJSON***
    * *Entity|Entity[]* **Fromstring**(*string* Source)  
        Creates/feeds Entities with data present on given Source  
        ***REQUIRES ASPJSON***
* *Export*
    * *Scripting.Dictionary* **AsDictionary**  
        Exports this Entity to a Dictionary
    * *JSONarray|JSONobject* **AsJSON**  
        Exports this Entity to a JSONobject
        ***REQUIRES ASPJSON***
    * *String* **AsString**  
        Exports this Entity to a JSON String
        ***REQUIRES ASPJSON***


