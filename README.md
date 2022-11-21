# FTencentDoc


一个腾讯文档自动签到的简单实现

## Requirements


使用项目内的requirements.txt安装依赖

```bash
pip install -r requirements.txt
```


## Usage


只需项目目录中运行即可
```bash
python main.py
```
for those who use Windows, you need to use `./`
```bash
python ./main.py
```

And here is a more general way is to run the script
```bash
<path of python.exe> <path to main.py>
```

### Register
运行项目内的register.py即可
```bash
python register.py
```


## Info



`main.py` 为主程序

`localConfig.json` 为配置文件

`trivial.ipynb` 为测试文件, Jupyter Notebook 格式

`logs` 为日志文件夹，每天运行会生成一个日志文件，该文件夹必须存在于项目目录下，否则程序会报错

`requirements.txt` 为依赖文件

其中，`localConfig.json` 的配置如下

```json
{
  "whereami": "<your location>",
  "reference_pos": [
    "<row number of the reference position>",
    "<column number of the reference position>"
  ],
  "doc_url": "<doc url>",
  "user": {
    "<user name in utf-8>": "<distance to the reference position>",
    "\u5510\u6d77\u6210": 1,
    "\u674e\u6587\u950b": 2,
    "\u9ad8\u6cfd\u68ee": -2,
    "\u5362\u56fd\u6d69": 13
  },
  "cookies": ["<your cookies>"]
}
```

具体函数的作用请参考代码，所有函数列表如下
![图片](https://user-images.githubusercontent.com/76607677/203082228-333bd0b0-e5c0-482b-a334-37121229201e.png)



## Others



![图片](https://user-images.githubusercontent.com/76607677/200879574-4797d354-86f4-4f20-b28c-0b27fcfd7c1d.png)

具备Log
![图片](https://user-images.githubusercontent.com/76607677/200880097-f66dc123-438c-4fa4-9829-85a7c766c26a.png)
