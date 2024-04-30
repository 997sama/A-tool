#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @PyCharm   :2023.3.4.PY-233.14475.56
# @Python    :3.12
# @FileName  :Baidu_Api_ocr.py
# @Time      :2024/4/16 下午2:06
# @Author    :997
# @E-mail    :A997sama@outlook.com
# @Description: 图片转文字API
# --------------------------------

import base64
import requests

API_KEY = '__________________________'  # 自行获取
SECRET_KEY = '__________________________'  # 自行获取
OCR_URL = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"  # OCR接口
TOKEN_URL = 'https://aip.baidubce.com/oauth/2.0/token'  # TOKEN获取接口
RETURN_WORD = ""

def fetch_token():
    params = {'grant_type': 'client_credentials',
              'client_id': API_KEY,
              'client_secret': SECRET_KEY}
    try:
        f = requests.post(TOKEN_URL, params, timeout=5)
        if f.status_code == 200:
            result = f.json()
            if 'access_token' in result.keys() and 'scope' in result.keys():
                if not 'brain_all_scope' in result['scope'].split(' '):
                    return None, 'please ensure has check the  ability'
                return result['access_token'], ''
            else:
                return None, '请输入正确的 API_KEY 和 SECRET_KEY'
        else:
            return None, '请求token失败: code {}'.format(f.status_code)
    except BaseException as err:
        return None, '请求token失败: {}'.format(err)


# 获取token

def read_file(image_path):
    f = None
    try:
        f = open(image_path, 'rb')  # 二进制读取图片信息
        return f.read(), ''
    except BaseException as e:
        return None, '文件({0})读取失败: {1}'.format(image_path, e)
    finally:
        if f:
            f.close()


# 读取图片

def pic2text(img_path): # 图片文件转文字
    def request_orc(img_base, token):
        """
        调用百度OCR接口，图片识别文字
        :param img_base: 图片的base64转码后的字符
        :param token: fetch_token返回的token
        :return: 返回一个识别后的文本字典
        """
        try:
            req = requests.post(
                OCR_URL + "?access_token=" + token,
                data={'image': img_base},
                headers={'Content-Type': 'application/x-www-form-urlencoded'}
            )
            if req.status_code == 200:
                result = req.json()
                if 'words_result' in result.keys():
                    return req.json()["words_result"], ''
                elif 'error_msg' in result.keys():
                    return None, '图片识别失败: {}'.format(req.json()["error_msg"])

            else:
                return None, '图片识别失败: code {}'.format(req.status_code)
        except BaseException as err:
            return None, '图片识别失败: {}'.format(err)

    file_content, file_error = read_file(img_path)
    if file_content:
        token, token_err = fetch_token()
        if token:
            results, result_err = request_orc(base64.b64encode(file_content), token)
            if result_err:  # 打印失败信息
                print(result_err)
            else:
                if results is not None:
                    RETURN_WORD = " ".join([result["words"] for result in results])

            # for result in results:  # 打印处理结果
            #     RETURN_WORD+=result
            print(RETURN_WORD)

def pic3text(img_base64): # base64转文字
    def request_orc(img_base, token):
        """
        调用百度OCR接口，图片识别文字
        :param img_base: 图片的base64转码后的字符
        :param token: fetch_token返回的token
        :return: 返回一个识别后的文本字典
        """
        try:
            req = requests.post(
                OCR_URL + "?access_token=" + token,
                data={'image': img_base},
                headers={'Content-Type': 'application/x-www-form-urlencoded'}
            )
            if req.status_code == 200:
                result = req.json()
                if 'words_result' in result.keys():
                    return req.json()["words_result"], ''
                elif 'error_msg' in result.keys():
                    return None, '图片识别失败: {}'.format(req.json()["error_msg"])
            else:
                return None, '图片识别失败: code {}'.format(req.status_code)
        except BaseException as err:
            return None, '图片识别失败: {}'.format(err)

    if img_base64:
        token, token_err = fetch_token()
        if token:
            results, result_err = request_orc(img_base64, token)
            if result_err:  # 打印失败信息
                print(result_err)
            else:
                if results is not None:
                    RETURN_WORD = " ".join([result["words"] for result in results])
                    return (RETURN_WORD)
        else:
            print("获取token失败")
    else:
        print("未提供图像的Base64编码")


def main():
    pic2text(img_path="../photos/222.png")
    return 0


if __name__ == "__main__":
    main()
else:
    print(f"Import {__name__} Successful")
