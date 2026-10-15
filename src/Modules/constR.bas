Attribute VB_Name = "constR"
'@lang VBA

Public Sub HighlightR()
    
    LANGUAGE_ = "Code"
    COMMENT_LINE_ = "#"
    COMMENT_MULTILINE_START_ = ""
    COMMENT_MULTILINE_END_ = ""
    STRING_MULTILINE_START_ = ""
    STRING_MULTILINE_END_ = ""
    RESERVED_ = Array( _
        "break", "else", "for", "function", "if", "in", "next", "repeat", "return", "while" _
    )
    OPERATORS_ = Array( _
        "<<-", "->>", "<-", "->", ":::", "::", "==", "!=", "<=", ">=", "&&", "||", "%in%", "%>%", _
        "+", "-", "*", "/", "^", ":", "<", ">", "=", "!", "&", "|", "$", "~", "?" _
    )
    TYPES_ = Array()
    BUILTINS_ = Array( _
        "abs", "append", "as.character", "as.data.frame", "as.integer", "as.list", "as.logical", "as.numeric", "c", "cbind", _
        "class", "data.frame", "dim", "factor", "head", "length", "levels", "list", "mean", "median", _
        "names", "nrow", "plot", "print", "rbind", "read.csv", "rep", "seq", "sort", "summary", _
        "tail", "table", "unique", "which", "write.csv" _
    )
    LITERALS_ = Array( _
        "FALSE", "Inf", "NA", "NA_character_", "NA_complex_", "NA_integer_", "NA_real_", "NaN", "NULL", "TRUE" _
    )
    
End Sub
