from django.http import HttpResponse
import datetime

def index(request):
    now=datetime.datetime.now()
    html="<html><title>Misc Error</title><body><h2>Sorry, an error occurred.</h2><br>(sent at %s)</body></html>" % now
    return HttpResponse(html)