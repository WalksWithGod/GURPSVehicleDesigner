Option Strict Off
Option Explicit On
Module modGVDXMLSchemaConstants
	
	' holds constants for our object store xml schema.
	
	'todo: convert every single one to lowercase for final... will make it easier to cut down on capitalization errors and it
	' provides a consistant rule to follow
	
	' node names
	Public Const XML_NODE_AUTHOR As String = "author"
	Public Const XML_NODE_VERSION As String = "version"
	Public Const XML_NODE_OPTIONS As String = "Options"
	
	Public Const XML_NODE_LOCATION As String = "location"
	
	Public Const XML_NODE_OBJECT As String = "object"
	Public Const XML_NODE_IMAGE As String = "image"
	Public Const XML_NODE_PROPERTY As String = "property"
	Public Const XML_NODE_PROPERTYCOUNT As String = "propertycount"
	Public Const XML_NODE_OPTION_MODIFER_TABLE_COUNT As String = "option_count"
	Public Const XML_NODE_STATS_TABLECOUNT As String = "statstablecount"
	Public Const XML_NODE_TABLE As String = "table"
	Public Const XML_NODE_FORMULA As String = "formulas"
	Public Const XML_NODE_MAXCHILDREN As String = "maxchildren"
	Public Const XML_NODE_CHILD As String = "child"
	Public Const XML_NODE_CHILDCOUNT As String = "childcount"
	Public Const XML_NODE_NAME As String = "name"
	Public Const XML_NODE_DESCRIPTION As String = "Description"
	Public Const XML_NODE_BODY As String = "Body"
	Public Const XML_NODE_DEFPATH As String = "defpath"
	Public Const XML_NODE_GUID As String = "guid"
	Public Const XML_NODE_CLASSNAME As String = "classname"
	
	
	'note datatypes
	Public Const XML_NODETYPE_DATETIME As String = "dateTime"
	Public Const XML_NODETYPE_F_TABLE As String = "sng_table"
	Public Const XML_NODETYPE_F_ARRAY_2D As String = "f_table_2D"
	Public Const XML_NODETYPE_OBJ_ARRAY As String = "o_array"
	
	Public Const XML_NODETYPE_V_ARRAY As String = "v_array"
	Public Const XML_NODETYPE_S_ARRAY As String = "s_array"
	Public Const XML_NODETYPE_F_ARRAY As String = "f_array"
	Public Const XML_NODETYPE_L_ARRAY As String = "l_array"
	Public Const XML_NODETYPE_I_ARRAY As String = "i_array"
	
	Public Const XML_NODETYPE_OBJECTREF As String = "objectRef"
	Public Const XML_NODETYPE_STRUCT As String = "struct"
	Public Const XML_NODETYPE_BOOL As String = "boolean"
	Public Const XML_NODETYPE_FLOAT As String = "float"
	Public Const XML_NODETYPE_STRING As String = "string"
	Public Const XML_NODETYPE_INTEGER As String = "int"
	
	'attribute names
	Public Const XML_ATTRIB_HANDLE As String = "handle"
	Public Const XML_ATTRIB_LOWERBOUND As String = "lowerBound"
	Public Const XML_ATTRIB_UPPERBOUND As String = "upperBound"
	Public Const XML_ATTRIB_NAME As String = "name"
	Public Const XML_ATTRIB_XMLSPACE As String = "xml:space"
	Public Const XML_ATTRIB_PRESERVE As String = "preserve"
	
	Public Const XML_ATTRIB_ROWLOWERBOUND As String = "lRow"
	Public Const XML_ATTRIB_ROWUPPERBOUND As String = "uRow"
	Public Const XML_ATTRIB_COLUMNLOWERBOUND As String = "lCol"
	Public Const XML_ATTRIB_COLUMNUPPERBOUND As String = "uCol"
	
	
	Public Enum GVD_XML_TYPE
		DEF = 0
		cmp = 1
	End Enum
End Module