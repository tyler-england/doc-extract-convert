import requests

ip='10.8.17.25:8000'

def index(request):
    url='http://'+ip+'/file'
    return requests.post(url)