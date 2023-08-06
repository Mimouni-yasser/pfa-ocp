from django.shortcuts import render
from django.http import HttpResponse
from .models import IP_field
from django.template.loader import render_to_string
from django.http import JsonResponse
from django.core import serializers



def index(request):
    ctx = {}
    url_par={}
    url_par["ip"] = request.GET.get("ip")
    url_par["mac"] = request.GET.get("mac")
    url_par["comment"] = request.GET.get("comment")
    
    if url_par:
        IPs = IP_field.objects.filter(IP__icontains=url_par['ip'], MAC__icontains=url_par['mac'], comment__icontains=url_par['comment'])
    else:
        IPs = IP_field.objects.all()

    
    is_ajax_request = request.headers.get("x-requested-with") == "XMLHttpRequest"
    if is_ajax_request:
        data = serializers.serialize("json", IPs, fields=('IP', 'MAC', 'comment', 'device_type', 'DateTime'))
        return JsonResponse(data, safe=False)
    
    data = serializers.serialize("json", IPs, fields=('IP', 'MAC', 'comment', 'device_type', 'DateTime'))
    ctx["IPs"] = data
    return render(request, "manager/index.html", context=ctx)