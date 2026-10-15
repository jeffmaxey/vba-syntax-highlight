Attribute VB_Name = "constSQL"
'@lang VBA

Public Sub HighlightSQL()
    
    LANGUAGE_ = "SQL"
    COMMENT_LINE_ = "--"
    COMMENT_MULTILINE_START_ = "/*"
    COMMENT_MULTILINE_END_ = "*/"
    STRING_MULTILINE_START_ = ""
    STRING_MULTILINE_END_ = ""
    RESERVED_ = Array( _
        "ADD", "ALL", "ALTER", "AND", "ANY", "AS", "ASC", "BACKUP", "BEGIN", "BETWEEN", _
        "BY", "CASE", "CAST", "CHECK", "COLUMN", "COMMIT", "CONSTRAINT", "CREATE", "CROSS", "CURRENT_DATE", _
        "CURRENT_TIME", "CURRENT_TIMESTAMP", "DATABASE", "DEFAULT", "DELETE", "DESC", "DISTINCT", "DROP", "ELSE", "END", _
        "EXCEPT", "EXISTS", "FOREIGN", "FROM", "FULL", "FUNCTION", "GRANT", "GROUP", "HAVING", "IN", _
        "INDEX", "INNER", "INSERT", "INTERSECT", "INTO", "IS", "JOIN", "KEY", "LEFT", "LIKE", _
        "LIMIT", "NOT", "OFFSET", "ON", "OR", "ORDER", "OUTER", "PRIMARY", "PROCEDURE", _
        "REFERENCES", "RIGHT", "ROLLBACK", "ROW", "SELECT", "SET", "TABLE", "THEN", "TOP", "TRANSACTION", _
        "TRIGGER", "UNION", "UNIQUE", "UPDATE", "USE", "USING", "VALUES", "VIEW", "WHEN", "WHERE", _
        "WITH" _
    )
    OPERATORS_ = Array( _
        "<>", "!=", ">=", "<=", "||", "::", "=", ">", "<", "+", "-", "*", "/", "%" _
    )
    TYPES_ = Array( _
        "BIGINT", "BINARY", "BIT", "BLOB", "BOOLEAN", "CHAR", "DATE", "DATETIME", "DECIMAL", "DOUBLE", _
        "FLOAT", "INTEGER", "INT", "MONEY", "NCHAR", "NUMERIC", "NVARCHAR", "REAL", "SMALLINT", "TEXT", _
        "TIME", "TIMESTAMP", "UUID", "VARCHAR" _
    )
    BUILTINS_ = Array( _
        "AVG", "COALESCE", "COUNT", "FIRST", "IFNULL", "LCASE", "LENGTH", "LOWER", "MAX", "MIN", _
        "NULLIF", "ROUND", "SUM", "UCASE", "UPPER" _
    )
    LITERALS_ = Array( _
        "FALSE", "NULL", "TRUE" _
    )
    
End Sub
