Attribute VB_Name = "Tests"
Option Explicit

Sub TestScrape()
    Dim v As Variant
    v = Scrape(ActiveSheet)
    Stop
End Sub
Function Scrape(st As Worksheet) As Variant
    '�A�N�e�B�u�V�[�g�̒��g���������炢�ɂ���
    '��������Z������v�f�Ƃ��ăJ�E���g�����
    Dim c As Scraper
    Set c = New Scraper
    
    Dim v As Variant
    c.Init st
    v = c.Scrape(v)
    Scrape = v
End Function

Sub TestArrayUtilCompress()
    Dim v As Variant
    v = Scrape(ActiveSheet)
    
    Dim c As ArrayUtil
    Set c = New ArrayUtil
    Dim ret
    ret = c.Compress(v)
Stop
End Sub

Sub TestDirLooper()
    Dim p
    p = "M:\������\�s1�t���ρE����\1. ���Ϗ�\"
    Dim c As DirLooper
    Set c = New DirLooper
    c.Init p
    'BookData�ɂ�Item=�t���p�X�AKey=���Ϗ��ԍ��̃R���N�V�������Ԃ��Ă���
    'Key���Q�Ƃ���ɂ�For each���[�v���g���B
    Dim BookData As Collection
    Set BookData = New Collection
    Set BookData = c.Indexing
End Sub

