Attribute VB_Name = "Module1"
Option Explicit
Sub �e�X�g()
    Dim ����, ��
    '�����F�N��, �ڎ��, ���N�`����, ��, �O��ڎ��, �O��N��E��
    ���� = �ڎ픻��(5, #2/1/2022#, "�t�@�C�U�[�i�T����P�P�Ηp�j", 2, #1/31/2022#, "")
    �� = ����(0) & vbCrLf & ����(1)
    MsgBox ��
End Sub
Function �ڎ픻��(�N�� As Variant, �ڎ�� As Date, ���N�`���� As String, �� As Long, �O��ڎ�� As Variant, �O��N��E�� As String) As Variant()
    Select Case ���N�`����
    '���s���N�`��
        Case "�R�~�i�e�B�i�w�a�a�D�P�D�T�j": �ڎ픻�� = XBB�t�@�C�U�[����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "�R�~�i�e�B�T����P�P�Ηp�w�a�a�D�P�D�T": �ڎ픻�� = ����XBB�t�@�C�U�[����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "�R�~�i�e�B�U��������S�Ηp�w�a�a�D�P�D�T": �ڎ픻�� = ���c��XBB�t�@�C�U�[����(�N��, �ڎ��, ��, �O��ڎ��, �O��N��E��)
        Case "�X�p�C�N�o�b�N�X�i�w�a�a�D�P�D�T�j": �ڎ픻�� = XBB���f���i����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "�X�p�C�N�o�b�N�X�U�`�P�P�΂w�a�a�D�P�D�T": �ڎ픻�� = ����XBB���f���i����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "�m�o�o�b�N�X": �ڎ픻�� = �m�o�o�b�N�X����(�N��, �ڎ��, ��, �O��ڎ��)
    '�I�����N�`��
        Case "�t�@�C�U�[": �ڎ픻�� = �t�@�C�U�[����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "���c�^���f���i": �ڎ픻�� = ���f���i����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "�A�X�g���[�l�J": �ڎ픻�� = �A�X�g���[�l�J����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "�t�@�C�U�[�i�T����P�P�Ηp�j": �ڎ픻�� = �����t�@�C�U�[����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "�R�~�i�e�B�i�T����P�P�Ηp�a�`�D�S�^�T�j": �ڎ픻�� = ����BA5�t�@�C�U�[����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "�R�~�i�e�B�i�Q���F�a�`�D�P�j": �ڎ픻�� = BA1�t�@�C�U�[����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "�R�~�i�e�B�i�Q���F�a�`�D�S�^�T�j": �ڎ픻�� = BA5�t�@�C�U�[����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "�X�p�C�N�o�b�N�X�i�Q���F�a�`�D�P�j": �ڎ픻�� = BA1���f���i����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "�X�p�C�N�o�b�N�X�i�Q���F�a�`�D�S�^�T�j": �ڎ픻�� = BA5���f���i����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "���f���i�i�U����P�P�Ηp�a�`�D�S�^�T�j": �ڎ픻�� = ����BA5���f���i����(�N��, �ڎ��, ��, �O��ڎ��)
        Case "�R�~�i�e�B�i�U��������S�Ηp�j": �ڎ픻�� = ���c���t�@�C�U�[����(�N��, �ڎ��, ��, �O��ڎ��, �O��N��E��)
    End Select
End Function
Function �N���(�N�� As Variant, ���� As Long, ��� As Long) As String
    If �N�� = "" Then
        �N��� = "�N��s��"
        Exit Function
    End If
    If �N�� < ���� Then �N��� = ���� & "�Ζ���"
    If ��� <> 0 Then
        If �N�� > ��� Then �N��� = ��� + 1 & "�Έȏ�"
    End If
End Function
Function �Ԋu����(�ڎ�� As Date, �O��ڎ�� As Variant, �ݒ�l As Long, �P�� As String) As String
    If �O��ڎ�� = "" Then
        �Ԋu���� = "�O��s��"
        Exit Function
    End If
    Select Case �P��
        Case "��"
            If �ڎ�� - �O��ڎ�� < �ݒ�l Then
                �Ԋu���� = �ݒ�l & �P�� & "����"
            End If
        Case "��"
            If DateAdd("m", -�ݒ�l, �ڎ��) < �O��ڎ�� Then
                �Ԋu���� = �ݒ�l & �P�� & "����"
            End If
    End Select
End Function
Function XBB�t�@�C�U�[����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '�ڎ�J�n
            Select Case ��
                Case 1: XBB�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), "")
                Case 2: XBB�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case 3, 4, 5, 6, 7: XBB�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: XBB�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Else: XBB�t�@�C�U�[���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function ����XBB�t�@�C�U�[����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '�ڎ�J�n
            Select Case ��
                Case 1: ����XBB�t�@�C�U�[���� = Array(�N���(�N��, 5, 11), "")
                Case 2: ����XBB�t�@�C�U�[���� = Array(�N���(�N��, 5, 11), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case 3, 4, 5, 6: ����XBB�t�@�C�U�[���� = Array(�N���(�N��, 5, 11), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: ����XBB�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Else: ����XBB�t�@�C�U�[���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function ���c��XBB�t�@�C�U�[����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant, �O��N��E�� As String) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '�ڎ�J�n
            Select Case ��
                Case 1: ���c��XBB�t�@�C�U�[���� = Array(�N���(�N��, 0, 4), "")
                Case 2
                    If �O��N��E�� = "" Then
                        ���c��XBB�t�@�C�U�[���� = Array("", �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                        Else: ���c��XBB�t�@�C�U�[���� = Array(�N���(�N��, 0, 4), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                    End If
                Case 3
                    If �O��N��E�� = "" Then
                        ���c��XBB�t�@�C�U�[���� = Array("", �Ԋu����(�ڎ��, �O��ڎ��, 56, "��"))
                        Else: ���c��XBB�t�@�C�U�[���� = Array(�N���(�N��, 0, 4), �Ԋu����(�ڎ��, �O��ڎ��, 56, "��"))
                    End If
                Case 4: ���c��XBB�t�@�C�U�[���� = Array(�N���(�N��, 0, 4), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: ���c��XBB�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Else: ���c��XBB�t�@�C�U�[���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function XBB���f���i����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/25/2023# '�ڎ�J�n
            Select Case ��
                Case 3, 4, 5, 6, 7: XBB���f���i���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: XBB���f���i���� = Array("", "�@��O��")
            End Select
        Case Else: XBB���f���i���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function ����XBB���f���i����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/25/2023# '�ڎ�J�n
            Select Case ��
                Case 3, 4, 5, 6: ����XBB���f���i���� = Array(�N���(�N��, 6, 11), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: ����XBB���f���i���� = Array("", "�@��O��")
            End Select
        Case Else: ����XBB���f���i���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function �m�o�o�b�N�X����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '7��ڐڎ�J�n
            Select Case ��
                Case 1: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), "")
                Case 2: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 21, "��"))
                Case 3, 4, 5, 6, 7: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 6, "��"))
                Case Else: �m�o�o�b�N�X���� = Array("", "�@��O��")
            End Select
        Case Is >= #5/8/2023# '6��ڐڎ�J�n
            Select Case ��
                Case 1: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), "")
                Case 2: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 21, "��"))
                Case 3, 4, 5, 6: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 6, "��"))
                Case Else: �m�o�o�b�N�X���� = Array("", "�@��O��")
            End Select
        Case Is >= #3/8/2023# '3�`5��ڔN���������
            Select Case ��
                Case 1: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), "")
                Case 2: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 21, "��"))
                Case 3, 4, 5: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 6, "��"))
                Case Else: �m�o�o�b�N�X���� = Array("", "�@��O��")
            End Select
        Case Is >= #11/8/2022# '4�E5��ڐڎ�J�n�iR4�H�J�n�ڎ�g�Ɉڍs�j
            Select Case ��
                Case 1: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), "")
                Case 2: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 21, "��"))
                Case 3, 4, 5: �m�o�o�b�N�X���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 6, "��"))
                Case Else: �m�o�o�b�N�X���� = Array("", "�@��O��")
            End Select
        Case Is >= #7/22/2022# '����̂ݔN���������
            Select Case ��
                Case 1: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), "")
                Case 2: �m�o�o�b�N�X���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 21, "��"))
                Case 3: �m�o�o�b�N�X���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 6, "��"))
                Case Else: �m�o�o�b�N�X���� = Array("", "�@��O��")
            End Select
        Case Is >= #5/25/2022# '1�`3��ڐڎ�J�n
            Select Case ��
                Case 1: �m�o�o�b�N�X���� = Array(�N���(�N��, 18, 0), "")
                Case 2: �m�o�o�b�N�X���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 21, "��"))
                Case 3: �m�o�o�b�N�X���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 6, "��"))
                Case Else: �m�o�o�b�N�X���� = Array("", "�@��O��")
            End Select
        Case Else: �m�o�o�b�N�X���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function BA5�t�@�C�U�[����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '�ڎ�I��
            BA5�t�@�C�U�[���� = Array("", "�@��O���N�`��")
        Case Is >= #8/7/2023# '1�E2��ڐڎ�J�n
            Select Case ��
                Case 1: BA5�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), "")
                Case 2: BA5�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case 3, 4, 5, 6: BA5�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: BA5�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #5/8/2023# '6��ڐڎ�J�n
            Select Case ��
                Case 3, 4, 5, 6: BA5�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: BA5�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #10/21/2022# '�Ԋu�Z�k
            Select Case ��
                Case 3, 4, 5: BA5�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: BA5�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #10/13/2022# '3�`5��ڐڎ�J�n
            Select Case ��
                Case 3, 4, 5: BA5�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 5, "��"))
                Case Else: BA5�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Else: BA5�t�@�C�U�[���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function BA1�t�@�C�U�[����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '�ڎ�I��
            BA1�t�@�C�U�[���� = Array("", "�@��O���N�`��")
        Case Is >= #8/7/2023# '1�E2��ڐڎ�J�n
            Select Case ��
                Case 1: BA1�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), "")
                Case 2: BA1�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case 3, 4, 5, 6: BA1�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: BA1�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #5/8/2023# '6��ڐڎ�J�n
            Select Case ��
                Case 3, 4, 5, 6: BA1�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: BA1�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #10/21/2022# '�Ԋu�Z�k
            Select Case ��
                Case 3, 4, 5: BA1�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: BA1�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #9/20/2022# '3�`5��ڐڎ�J�n
            Select Case ��
                Case 3, 4, 5: BA1�t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 5, "��"))
                Case Else: BA1�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Else: BA1�t�@�C�U�[���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function �t�@�C�U�[����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '�ڎ�I��
            �t�@�C�U�[���� = Array("", "�@��O���N�`��")
        Case Is >= #4/1/2023# '3�E4��ڎg�p�I��
            Select Case ��
                Case 1: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), "")
                Case 2: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case Else: �t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #10/21/2022# '3�E4��ڊԊu�Z�k
            Select Case ��
                Case 1: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), "")
                Case 2: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case 3: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case 4: �t�@�C�U�[���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: �t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #5/25/2022# '4��ڐڎ�J�n���Ԋu�Z�k
            Select Case ��
                Case 1: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), "")
                Case 2: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case 3: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 5, "��"))
                Case 4: �t�@�C�U�[���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 5, "��"))
                Case Else: �t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #3/25/2022# '3��ڔN���������
            Select Case ��
                Case 1: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), "")
                Case 2: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case 3: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 6, "��"))
                Case Else: �t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #12/17/2021# '3��ڊԊu�Z�k
            Select Case ��
                Case 1: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), "")
                Case 2: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case 3: �t�@�C�U�[���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 6, "��"))
                Case Else: �t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #12/1/2021# '3��ڐڎ�J�n
            Select Case ��
                Case 1: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), "")
                Case 2: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case 3: �t�@�C�U�[���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 8, "��"))
                Case Else: �t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #6/1/2021# '�N���������
            Select Case ��
                Case 1: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), "")
                Case 2: �t�@�C�U�[���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case Else: �t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #3/1/2021# '����
            Select Case ��
                Case 1: �t�@�C�U�[���� = Array(�N���(�N��, 16, 0), "")
                Case 2: �t�@�C�U�[���� = Array(�N���(�N��, 16, 0), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case Else: �t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Else: �t�@�C�U�[���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function ����BA5�t�@�C�U�[����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '�ڎ�I��
            ����BA5�t�@�C�U�[���� = Array("", "�@��O���N�`��")
        Case Is >= #8/7/2023# '1�E2��ڐڎ�J�n
            Select Case ��
                Case 1: ����BA5�t�@�C�U�[���� = Array(�N���(�N��, 5, 11), "")
                Case 2: ����BA5�t�@�C�U�[���� = Array(�N���(�N��, 5, 11), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case 3, 4, 5: ����BA5�t�@�C�U�[���� = Array(�N���(�N��, 5, 11), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: ����BA5�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #5/8/2023# '5��ڐڎ�J�n
            Select Case ��
                Case 3, 4, 5: ����BA5�t�@�C�U�[���� = Array(�N���(�N��, 5, 11), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: ����BA5�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #3/8/2023# '3�E4��ڐڎ�J�n
            Select Case ��
                Case 3, 4: ����BA5�t�@�C�U�[���� = Array(�N���(�N��, 5, 11), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: ����BA5�t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Else: ����BA5�t�@�C�U�[���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function �����t�@�C�U�[����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '�ڎ�I��
            �����t�@�C�U�[���� = Array("", "�@��O���N�`��")
        Case Is >= #4/1/2023# '3��ڎg�p�I��
            Select Case ��
                Case 1: �����t�@�C�U�[���� = Array(�N���(�N��, 5, 11), "")
                Case 2: �����t�@�C�U�[���� = Array(�N���(�N��, 5, 11), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case Else: �����t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #3/8/2023# '3��ڊԊu�Z�k
            Select Case ��
                Case 1: �����t�@�C�U�[���� = Array(�N���(�N��, 5, 11), "")
                Case 2: �����t�@�C�U�[���� = Array(�N���(�N��, 5, 11), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case 3: �����t�@�C�U�[���� = Array(�N���(�N��, 5, 11), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: �����t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #9/6/2022# '3��ڐڎ�J�n
            Select Case ��
                Case 1: �����t�@�C�U�[���� = Array(�N���(�N��, 5, 11), "")
                Case 2: �����t�@�C�U�[���� = Array(�N���(�N��, 5, 11), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case 3: �����t�@�C�U�[���� = Array(�N���(�N��, 5, 11), �Ԋu����(�ڎ��, �O��ڎ��, 5, "��"))
                Case Else: �����t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Is >= #2/1/2022# '�ڎ�J�n
            Select Case ��
                Case 1: �����t�@�C�U�[���� = Array(�N���(�N��, 5, 11), "")
                Case 2: �����t�@�C�U�[���� = Array(�N���(�N��, 5, 11), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                Case Else: �����t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Else: �����t�@�C�U�[���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function ���c���t�@�C�U�[����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant, �O��N��E�� As String) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '�ڎ�I��
            ���c���t�@�C�U�[���� = Array("", "�@��O���N�`��")
        Case Is >= #10/24/2022# '�ڎ�J�n
            Select Case ��
                Case 1: ���c���t�@�C�U�[���� = Array(�N���(�N��, 0, 4), "")
                Case 2
                    If �O��N��E�� = "" Then
                        ���c���t�@�C�U�[���� = Array("", �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                        Else: ���c���t�@�C�U�[���� = Array(�N���(�N��, 0, 4), �Ԋu����(�ڎ��, �O��ڎ��, 19, "��"))
                    End If
                Case 3
                    If �O��N��E�� = "" Then
                        ���c���t�@�C�U�[���� = Array("", �Ԋu����(�ڎ��, �O��ڎ��, 56, "��"))
                        Else: ���c���t�@�C�U�[���� = Array(�N���(�N��, 0, 4), �Ԋu����(�ڎ��, �O��ڎ��, 56, "��"))
                    End If
                Case Else: ���c���t�@�C�U�[���� = Array("", "�@��O��")
            End Select
        Case Else: ���c���t�@�C�U�[���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function BA5���f���i����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '�ڎ�I��
            BA5���f���i���� = Array("", "�@��O���N�`��")
        Case Is >= #5/8/2023# '6��ڐڎ�J�n
            Select Case ��
                Case 3, 4, 5, 6: BA5���f���i���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: BA5���f���i���� = Array("", "�@��O��")
            End Select
        Case Is >= #12/14/2022# '�N���������
            Select Case ��
                Case 3, 4, 5: BA5���f���i���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: BA5���f���i���� = Array("", "�@��O��")
            End Select
        Case Is >= #11/28/2022# '3�`5��ڐڎ�J�n
            Select Case ��
                Case 3, 4, 5: BA5���f���i���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: BA5���f���i���� = Array("", "�@��O��")
            End Select
        Case Else: BA5���f���i���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function BA1���f���i����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '�ڎ�I��
            BA1���f���i���� = Array("", "�@��O���N�`��")
        Case Is >= #5/8/2023# '6��ڐڎ�J�n
            Select Case ��
                Case 3, 4, 5, 6: BA1���f���i���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: BA1���f���i���� = Array("", "�@��O��")
            End Select
        Case Is >= #12/14/2022# '�N���������
            Select Case ��
                Case 3, 4, 5: BA1���f���i���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: BA1���f���i���� = Array("", "�@��O��")
            End Select
        Case Is >= #10/21/2022# '�Ԋu�Z�k
            Select Case ��
                Case 3, 4, 5: BA1���f���i���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: BA1���f���i���� = Array("", "�@��O��")
            End Select
        Case Is >= #9/20/2022# '3�`5��ڐڎ�J�n
            Select Case ��
                Case 3, 4, 5: BA1���f���i���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 5, "��"))
                Case Else: BA1���f���i���� = Array("", "�@��O��")
            End Select
        Case Else: BA1���f���i���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function ����BA5���f���i����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #9/20/2023# '�ڎ�I��
            ����BA5���f���i���� = Array("", "�@��O���N�`��")
        Case Is >= #8/7/2023# '3�`5��ڐڎ�J�n
            Select Case ��
                Case 3, 4, 5: ����BA5���f���i���� = Array(�N���(�N��, 6, 11), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: ����BA5���f���i���� = Array("", "�@��O��")
            End Select
        Case Else: ����BA5���f���i���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function ���f���i����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #2/12/2023# '�ڎ�I��
            ���f���i���� = Array("", "�@��O���N�`��")
        Case Is >= #12/14/2022# '3��ڔN���������
            Select Case ��
                Case 1: ���f���i���� = Array(�N���(�N��, 12, 0), "")
                Case 2: ���f���i���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 21, "��"))
                Case 3: ���f���i���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case 4: ���f���i���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: ���f���i���� = Array("", "�@��O��")
            End Select
        Case Is >= #10/21/2022# '3�E4��ڊԊu�Z�k
            Select Case ��
                Case 1: ���f���i���� = Array(�N���(�N��, 12, 0), "")
                Case 2: ���f���i���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 21, "��"))
                Case 3, 4: ���f���i���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 3, "��"))
                Case Else: ���f���i���� = Array("", "�@��O��")
            End Select
        Case Is >= #5/25/2022# '4��ڐڎ�J�n���Ԋu�Z�k
            Select Case ��
                Case 1: ���f���i���� = Array(�N���(�N��, 12, 0), "")
                Case 2: ���f���i���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 21, "��"))
                Case 3, 4: ���f���i���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 5, "��"))
                Case Else: ���f���i���� = Array("", "�@��O��")
            End Select
        Case Is >= #12/17/2021# '3��ڐڎ�J�n
            Select Case ��
                Case 1: ���f���i���� = Array(�N���(�N��, 12, 0), "")
                Case 2: ���f���i���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 21, "��"))
                Case 3: ���f���i���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 6, "��"))
                Case Else: ���f���i���� = Array("", "�@��O��")
            End Select
        Case Is >= #8/2/2021# '�N���������
            Select Case ��
                Case 1: ���f���i���� = Array(�N���(�N��, 12, 0), "")
                Case 2: ���f���i���� = Array(�N���(�N��, 12, 0), �Ԋu����(�ڎ��, �O��ڎ��, 21, "��"))
                Case Else: ���f���i���� = Array("", "�@��O��")
            End Select
        Case Is >= #5/22/2021#
            Select Case ��
                Case 1: ���f���i���� = Array(�N���(�N��, 18, 0), "")
                Case 2: ���f���i���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 21, "��"))
                Case Else: ���f���i���� = Array("", "�@��O��")
            End Select
        Case Else: ���f���i���� = Array("", "�@��O���N�`��")
    End Select
End Function
Function �A�X�g���[�l�J����(�N�� As Variant, �ڎ�� As Date, �� As Long, �O��ڎ�� As Variant) As Variant()
    Select Case �ڎ��
        Case Is >= #10/13/2022# '�ڎ�I��
            �A�X�g���[�l�J���� = Array("", "�@��O���N�`��")
        Case Is >= #8/2/2021# '�ڎ�J�n
            Select Case ��
                Case 1: �A�X�g���[�l�J���� = Array(�N���(�N��, 18, 0), "")
                Case 2: �A�X�g���[�l�J���� = Array(�N���(�N��, 18, 0), �Ԋu����(�ڎ��, �O��ڎ��, 28, "��"))
                Case Else: �A�X�g���[�l�J���� = Array("", "�@��O��")
            End Select
        Case Else: �A�X�g���[�l�J���� = Array("", "�@��O���N�`��")
    End Select
End Function
