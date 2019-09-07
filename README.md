# COSCUP 議程組寄徵稿結果通知信 Script

## 工具

- Python 3+
- Google Spreadsheet -> [範本文件](https://docs.google.com/spreadsheets/d/1yjkHrdNcdF5ghOoiDzEuQf0HBhBn5mro_MqHADHU8ow/edit?usp=sharing)
- AWS

## 安裝相關套件

```sh
$ pip3 install -r requirements.txt
```

## 怎麼使用

1. 撰寫 `accept.txt`、`deny.txt`，錄取通知與落選通知

2. 向行政組請求：`aws_access_key_id`、`aws_secret_access_key`

3. 下載操作 Spreadsheet 所需要的憑證，[詳見](https://pygsheets.readthedocs.io/en/latest/authorization.html)

3. 建立[範本文件](https://docs.google.com/spreadsheets/d/1yjkHrdNcdF5ghOoiDzEuQf0HBhBn5mro_MqHADHU8ow/edit?usp=sharing)的副本

4. 將資料寫入範本文件

5. 編輯 `main.py`，填上對應資料

6. Spreadsheet 的 key 是指這個：

![](https://i.imgur.com/9sjKor5.png)

7. Run

```sh
$ python3 main.py
```


