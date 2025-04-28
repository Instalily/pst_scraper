from enum import Enum
    
class RecipientType(Enum):
    NONE = 0
    TO = 1
    CC = 2
    BCC = 3

class BodyType(Enum):
    TEXT  = 0
    HTML = 1
    RTF = 2

class Sensitivity(Enum):
    NORMAL = 0
    PERSONAL = 1
    PRIVATE = 2
    CONFIDENTIAL = 3