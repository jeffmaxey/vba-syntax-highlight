Attribute VB_Name = "constVB"
'@lang VBA

Public Sub HighlightVB()
    
    LANGUAGE_ = "VB"
    COMMENT_LINE_ = "'"
    COMMENT_MULTILINE_START_ = ""
    COMMENT_MULTILINE_END_ = ""
    STRING_MULTILINE_START_ = ""
    STRING_MULTILINE_END_ = ""
    RESERVED_ = Array( _
        "AddHandler", "AddressOf", "Alias", "And", "AndAlso", "As", "ByRef", "ByVal", "Call", "Case", _
        "Catch", "Class", "Const", "Continue", "Declare", "Default", "Delegate", "Dim", "DirectCast", "Do", _
        "Each", "Else", "ElseIf", "End", "EndIf", "Enum", "Erase", "Error", "Event", "Exit", _
        "Finally", "For", "Friend", "Function", "Get", "GetType", "GoSub", "GoTo", "Handles", "If", _
        "Implements", "Imports", "In", "Inherits", "Interface", "Is", "IsNot", "Let", "Lib", "Like", _
        "Loop", "Me", "Mod", "Module", "MustInherit", "MustOverride", "MyBase", "MyClass", "Namespace", "Narrowing", _
        "New", "Next", "Not", "NotInheritable", "NotOverridable", "On", "Operator", "Option", "Optional", "Or", _
        "OrElse", "Overloads", "Overridable", "Overrides", "ParamArray", "Partial", "Private", "Property", "Protected", "Public", _
        "RaiseEvent", "ReadOnly", "ReDim", "RemoveHandler", "Resume", "Return", "Select", "Set", "Shadows", "Shared", _
        "Static", "Step", "Stop", "Structure", "Sub", "SyncLock", "Then", "Throw", "To", "Try", _
        "TryCast", "Using", "When", "While", "Widening", "With", "WithEvents", "WriteOnly", "Xor" _
    )
    OPERATORS_ = Array( _
        "=", "<>", ">", "<", ">=", "<=", "+", "-", "*", "/", "\", "^", "&" _
    )
    TYPES_ = Array( _
        "Boolean", "Byte", "Char", "Date", "Decimal", "Double", "Integer", "Long", "Object", "SByte", _
        "Short", "Single", "String", "UInteger", "ULong", "UShort", "Variant" _
    )
    BUILTINS_ = Array( _
        "CBool", "CByte", "CChar", "CDate", "CDbl", "CDec", "CInt", "CLng", "CObj", "CStr", _
        "CType", "CUInt", "CULng", "CShort", "CSng", "CUShort", "Format", "InStr", "InputBox", "Len", _
        "LCase", "Left", "Mid", "MsgBox", "Replace", "Right", "Split", "Trim", "UCase" _
    )
    LITERALS_ = Array( _
        "False", "Nothing", "True" _
    )
    
End Sub
