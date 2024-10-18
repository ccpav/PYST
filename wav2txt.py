import base64,hashlib,hmac,json,os,time,requests,urllib

lfasr_host = 'https://raasr.xfyun.cn/v2/api'
api_upload = '/upload'
api_get_result = '/getResult'

class RequestApi(object):
    def __init__(self, appid, secret_key, upload_file_path):
        self.appid = appid
        self.secret_key = secret_key
        self.upload_file_path = upload_file_path
        self.ts = str(int(time.time()))
        self.signa = self.get_signa()

    def get_signa(self):
        appid = self.appid
        secret_key = self.secret_key
        m2 = hashlib.md5()
        m2.update((appid + self.ts).encode('utf-8'))
        md5 = m2.hexdigest()
        md5 = bytes(md5, encoding='utf-8')
        signa = hmac.new(secret_key.encode('utf-8'), md5, hashlib.sha1).digest()
        signa = base64.b64encode(signa)
        signa = str(signa, 'utf-8')
        return signa

    def upload(self):
        print("上传部分：")
        upload_file_path = self.upload_file_path
        file_len = os.path.getsize(upload_file_path)
        file_name = os.path.basename(upload_file_path)

        param_dict = {
            'appId': self.appid,
            'signa': self.signa,
            'ts': self.ts,
            "fileSize": file_len,
            "fileName": file_name,
            "duration": "200"
        }
        print("upload参数：", param_dict)
        data = open(upload_file_path, 'rb').read(file_len)

        response = requests.post(url=lfasr_host + api_upload + "?" + urllib.parse.urlencode(param_dict),
                                 headers={"Content-type": "application/json"}, data=data)
        print("upload_url:", response.request.url)
        result = json.loads(response.text)
        print("upload resp:", result)
        return result

    def get_result(self):
        uploadresp = self.upload()
        orderId = uploadresp['content']['orderId']
        param_dict = {
            'appId': self.appid,
            'signa': self.signa,
            'ts': self.ts,
            'orderId': orderId,
            'resultType': "transfer,predict"
        }
        print("")
        print("查询部分：")
        print("get result参数：", param_dict)
        status = 3
        while status == 3:
            response = requests.post(url=lfasr_host + api_get_result + "?" + urllib.parse.urlencode(param_dict),
                                     headers={"Content-type": "application/json"})
            result = json.loads(response.text)
            print(result)
            status = result['content']['orderInfo']['status']
            print("status=", status)
            if status == 4:
                break
            time.sleep(5)
        print("get_result resp:", result)
        return result


def extract_text_from_order_result(order_result):
    order_result = json.loads(order_result)  # 解析 JSON 字符串
    text = []
    for item in order_result.get('lattice', []):
        json_1best = json.loads(item['json_1best'])  # 解析 json_1best 字符串
        words = json_1best['st']['rt'][0]['ws']  # 获取词语列表
        for word_info in words:
            for word in word_info['cw']:
                text.append(word['w'])  # 提取每个词的内容
    return ''.join(text)  # 合并成一个字符串


def list_wav_files():
    return [f for f in os.listdir() if f.endswith('.wav')]


def convert_all_files(wav_files):
    for file in wav_files:
        api = RequestApi(appid="fe56125a",
                         secret_key="b49bcf17eb25b082157b88c6a091b2aa",
                         upload_file_path=file)
        result = api.get_result()
        order_result = result['content']['orderResult']
        extracted_text = extract_text_from_order_result(order_result)
        with open(file.replace('.wav', '.txt'), 'w', encoding='utf-8') as f:
            f.write(extracted_text)


def main():
    print("很高兴你能用pony编制的wav音频转换工具，更新会在https://github.com/ccpav/PYST发布")
    print("1. 转换全部WAV音频文件")
    print("2. 手动选择指定的同目录WAV文件")
    choice = input("您要选哪种形式(1 or 2): ")

    if choice == '1':
        wav_files = list_wav_files()
        if not wav_files:
            print("同目录下没找到wav文件啊，看看文件放的对不对位置！！")
            return
        convert_all_files(wav_files)
        print("恭喜，完成了全部的转换任务")
    elif choice == '2':
        wav_files = list_wav_files()
        print("Available WAV files:")
        for i, file in enumerate(wav_files):
            print(f"{i + 1}. {file}")

        file_choice = int(input("告诉我你要处理那个文件编号: "))
        selected_file = wav_files[file_choice - 1]
        print(f"Selected file: {selected_file}")

        api = RequestApi(appid="#你自己在讯飞上的appid",
                         secret_key="#你自己在讯飞上的secret_key",
                         upload_file_path=selected_file)

        result = api.get_result()
        order_result = result['content']['orderResult']
        extracted_text = extract_text_from_order_result(order_result)

        # 保存提取的文本到同名的 TXT 文件
        with open(selected_file.replace('.wav', '.txt'), 'w', encoding='utf-8') as f:
            f.write(extracted_text)
    else:
        print("Invalid choice. 请选择 1 or 2.")


if __name__ == '__main__':
    main()