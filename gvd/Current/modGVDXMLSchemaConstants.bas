Attribute VB_Name = "modGVDXMLSchemaConstants"
Option Explicit

' holds constants for our object store xml schema.

'todo: convert every single one to lowercase for final... will make it easier to cut down on capitalization errors and it
' provides a consistant rule to follow

' node names
Public Const XML_NODE_AUTHOR = "author"
Public Const XML_NODE_VERSION = "version"
Public Const XML_NODE_OPTIONS = "Options"

Public Const XML_NODE_LOCATION = "location"

Public Const XML_NODE_OBJECT = "object"
Public Const XML_NODE_IMAGE = "image"
Public Const XML_NODE_PROPERTY = "property"
Public Const XML_NODE_PROPERTYCOUNT = "propertycount"
Public Const XML_NODE_OPTION_MODIFER_TABLE_COUNT = "option_count"
Public Const XML_NODE_STATS_TABLECOUNT = "statstablecount"
Public Const XML_NODE_TABLE = "table"
Public Const XML_NODE_FORMULA = "formulas"
Public Const XML_NODE_MAXCHILDREN = "maxchildren"
Public Const XML_NODE_CHILD = "child"
Public Const XML_NODE_CHILDCOUNT = "childcount"
Public Const XML_NODE_NAME = "name"
Public Const XML_NODE_DESCRIPTION = "Description"
Public Const XML_NODE_BODY = "Body"
Public Const XML_NODE_DEFPATH = "defpath"
Public Const XML_NODE_GUID = "guid"
Public Const XML_NODE_CLASSNAME = "classname"


'note datatypes
Public Const XML_NODETYPE_DATETIME = "dateTime"
Public Const XML_NODETYPE_F_TABLE = "sng_table"
Public Const XML_NODETYPE_F_ARRAY_2D = "f_table_2D"
Public Const XML_NODETYPE_OBJ_ARRAY = "o_array"

Public Const XML_NODETYPE_V_ARRAY = "v_array"
Public Const XML_NODETYPE_S_ARRAY = "s_array"
Public Const XML_NODETYPE_F_ARRAY = "f_array"
Public Const XML_NODETYPE_L_ARRAY = "l_array"
Public Const XML_NODETYPE_I_ARRAY = "i_array"

Public Const XML_NODETYPE_OBJECTREF = "objectRef"
Public Const XML_NODETYPE_STRUCT = "struct"
Public Const XML_NODETYPE_BOOL = "boolean"
Public Const XML_NODETYPE_FLOAT = "float"
Public Const XML_NODETYPE_STRING = "string"
Public Const XML_NODETYPE_INTEGER = "int"

'attribute names
Public Const XML_ATTRIB_HANDLE = "handle"
Public Const XML_ATTRIB_LOWERBOUND = "lowerBound"
Public Const XML_ATTRIB_UPPERBOUND = "upperBound"
Public Const XML_ATTRIB_NAME = "name"
Public Const XML_ATTRIB_XMLSPACE = "xml:space"
Public Const XML_ATTRIB_PRESERVE = "preserve"

Public Const XML_ATTRIB_ROWLOWERBOUND = "lRow"
Public Const XML_ATTRIB_ROWUPPERBOUND = "uRow"
Public Const XML_ATTRIB_COLUMNLOWERBOUND = "lCol"
Public Const XML_ATTRIB_COLUMNUPPERBOUND = "uCol"


Public Enum GVD_XML_TYPE
    DEF = 0
    cmp = 1
End Enum
