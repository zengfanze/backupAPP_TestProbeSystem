Attribute VB_Name = "value"
Option Explicit
'shell ����
Global shellobject As New Shell
'�ļ�Ŀ¼��ַ����
Global openfolder As Folder
Global savefolder As Folder
'Ŀ������Ŀ¼
Global searchPath As String
'Ŀ�걣��Ŀ¼
Global savePath As String
'��һ���ļ��м������飬��̽ͷ�ͺ�����
Global first() As String
'�ڶ����ļ��м������飬��̽ͷ��Ӧ�����к�����
Global second() As String
'�����ȡ���ļ�����
Global filedata() As String
'�Ƚ�����
Global checkdate As String
'�����ƶ�
Global rebackFolders As Boolean

