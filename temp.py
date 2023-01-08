import requests
# from urllib import request

class calculator:
    def getExcelResponse(url, savepath):
        response = requests.get(url)
        open(savepath).write(response.content)
