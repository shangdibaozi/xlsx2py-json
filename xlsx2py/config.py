# -*- coding: gb2312 -*-
"""
�����ļ�
"""


#PY_MODULE_PATH = r'res\entities\datum'
######�������ݵı�ͷ����##########
EXPORT_DATA_CODING = "utf-8"


EXPORT_DATA_HEAD = "# -*- coding: %s -*-\n\n"%(EXPORT_DATA_CODING,)



###############################
#����
###############################
#����sheetǰ׺
EXPORT_PREFIX_CHAR = '@'
EXPORT_DEFINE_ROW = 1

EXPORT_KEY_NUMS = 1

MAP_DEFINE_ROW = 1
MAP_DATA_ROW = 3

#���Ա������
EXPORT_MAP_SHEET = '���Ա�='

#�ļ����룬�����ļ��ı�����ΪUTF-8
FILE_CODE = "GB2312"

############################
#�������				   #
############################
EXPORT_SIGN_DOT = '.'
EXPORT_SIGN_DOLLAR = '$'
EXPORT_SIGN_GTH = '!'

CHECK_FUN = None

#format: index:checkfunc
EXPORT_SIGN= {
	EXPORT_SIGN_DOT:	CHECK_FUN,
	EXPORT_SIGN_DOLLAR : 	CHECK_FUN,
	EXPORT_SIGN_GTH:	CHECK_FUN,
}

EXPORT_ALL_SIGNS = [e for e in EXPORT_SIGN.keys()]

####################error�ֵ�##########################
EXPORT_ERROR_NOSHEET = 1
EXPORT_ERROR_NOMAP = 2
EXPORT_ERROR_HEADER = 3
EXPORT_ERROR_NOTNULL = 4
EXPORT_ERROR_REPEAT = 5
EXPORT_ERROR_REPKEY = 6
EXPORT_ERROR_NUMKEY = 7
EXPORT_ERROR_NOKEY = 8
EXPORT_ERROR_NOFUNC = 9
#���ݼ�����
EXPORT_ERROR_DATAINV  = 20
EXPORT_ERROR_NOSIGN = 21
EXPORT_ERROR_NOTMAP = 22
EXPORT_ERROR_FUNC	= 23
#�ļ�IO����
EXPORT_ERROR_CPATH = 30
EXPORT_ERROR_FILEOPENED = 31
EXPORT_ERROR_NOEXISTFILE = 32
EXPORT_ERROR_OTHER = 101
EXPORT_ERROR_FILEOPEN = 102
EXPORT_ERROR_IOOP = 103

EXPORT_ERROR = {
	EXPORT_ERROR_NOSHEET:'�ޱ�ɵ�',
	EXPORT_ERROR_NOMAP:'�޴��Ա�',
	EXPORT_ERROR_HEADER:'�ļ�ͷ����',
	EXPORT_ERROR_NOTNULL:'����Ϊ��',
	EXPORT_ERROR_REPEAT:'�����ظ�',
	EXPORT_ERROR_DATAINV:'�����붨�岻����',
	EXPORT_ERROR_OTHER:'���Ǵ���',
	EXPORT_ERROR_NUMKEY:'��Ҫ��key̫��',
	EXPORT_ERROR_NOSIGN:'�����ڵķ���',
	EXPORT_ERROR_REPKEY:'��Ϊ�ؼ��ֵ������ظ���keyֵ',
	EXPORT_ERROR_NOTMAP:'��Ҫ���ԣ���û�д��Թ�ϵ',
	EXPORT_ERROR_NOKEY:'û����key',
	EXPORT_ERROR_CPATH:'Ŀ¼����ʧ��',
	EXPORT_ERROR_FILEOPENED:"�ļ��Ѵ���رպ�������",
	EXPORT_ERROR_NOFUNC:"�����ڵ�ת������",
	EXPORT_ERROR_NOEXISTFILE:'excel�ļ�������',
	EXPORT_ERROR_FILEOPEN:'�ļ���ʧ��',
	EXPORT_ERROR_IOOP:'�ļ���д����',
	EXPORT_ERROR_FUNC:'��������',
}

EXPORT_INFO_NULL = 0
EXPORT_INFO_OK = 1
EXPORT_INFO_ING = 2
EXPORT_INFO_CDIR = 3
EXPORT_INFO_YN = 4
EXPORT_INFO_RTEXCEL = 5

EXPORT_INFO = {
	EXPORT_INFO_NULL:"\b",
	EXPORT_INFO_YN:"�Ƿ����Y or N",
	EXPORT_INFO_OK:"�ļ�������ȷ���Ƿ�Ҫ����(Y or N)",
	EXPORT_INFO_ING:"���ڵ���",
	EXPORT_INFO_CDIR:"�ļ��Ѵ�",
	EXPORT_INFO_RTEXCEL:'�ر��ļ������ԣ� ���������O�ó������ر�',
}
