import os

from Crypto.Cipher import XOR
import base64


def encrypt(key, plain_text):
    """
    암호화 키와 평문을 받아서 암호화된 문자열을 반환하는 함수
    :param key: 암호화 키 (32자리 문자열)
    :param plain_text: 암호화 하고자 하는 평문
    :return: 암호화된 byte 타입 데이터
    """
    cipher = XOR.new(key)
    return base64.b64encode(cipher.encrypt(plain_text))


def decrypt(key, encrypted_text):
    """
    암호화된 byte 타입 데이터를 받아서 원문을 반환하는 함수
    :param key: 암호화 키 (32자리 문자열)
    :param encrypted_text: 암호화된 byte 타입 데이터
    :return: 복호화된 원문
    """
    if type(encrypted_text) is str:
        encrypted_text = encrypted_text[2:][:-1].encode('utf-8')

    cipher = XOR.new(key)
    return cipher.decrypt(base64.b64decode(encrypted_text)).decode('utf-8')


def save_block_info(block, file_name):
    count = 1
    while os.path.exists(file_name):
        file_name += str(count)
        count += 1

    with open(file_name, 'wt') as f:
        f.write(block)
