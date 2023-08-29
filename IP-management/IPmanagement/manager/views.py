from django.shortcuts import render
from django.http import HttpResponse
from .models import IP_field
from django.template.loader import render_to_string
from django.http import JsonResponse
from django.core import serializers


def delete(request):
    
     if request.method == 'POST':
        ip = request.POST.get('ip') or None
        pk = request.POST.get('pk') or None
        if ip is None or pk is None or not IP_field.objects.filter(pk=pk).exists():
            return HttpResponse('element  non trouvé')
        else:
            IP_field.objects.filter(pk=pk).delete()
            return HttpResponse('ok')
        
def add(request):
    if request.method == 'POST':
        ip = request.POST.get('ip') or None
        mac = request.POST.get('mac') or None
        comment = request.POST.get('comment') or None
        type = request.POST.get('type') or None
        if ip is None or mac is None or comment is None or type is None:
            return HttpResponse('champs manquants')
        if IP_field.objects.filter(IP=ip).exists():
            return HttpResponse('l\'adresse IP' + ip + ' existe déjà')
        else:
            obj = IP_field(IP=ip)
            obj.MAC = mac
            obj.comment = comment
            obj.device_type = type
            obj.save()
        
            return HttpResponse('ok')

def update(request):
    pass

def index(request):
    
    if request.method == 'GET':
        ctx = {}
        url_par={}
        url_par["ip"] = request.GET.get("ip") or ''
        url_par["mac"] = request.GET.get("mac") or ''
        url_par["comment"] = request.GET.get("comment") or ''
        url_par["type"] = request.GET.get("type") or ''
        
        
        
        if url_par:
            IPs = IP_field.objects.filter(IP__icontains=url_par['ip'], MAC__icontains=url_par['mac'], comment__icontains=url_par['comment'], device_type__icontains=url_par['type'])
        else:
            IPs = IP_field.objects.all()

        
        is_ajax_request = request.headers.get("x-requested-with") == "XMLHttpRequest"
        if is_ajax_request:
            data = serializers.serialize("json", IPs, fields=('IP', 'MAC', 'comment', 'device_type', 'DateTime'))
            return JsonResponse(data, safe=False)
        
        data = serializers.serialize("json", IPs, fields=('IP', 'MAC', 'comment', 'device_type', 'DateTime'))
        return render(request, "manager/index.html")
    elif request.method == 'POST':
        
        ip = request.POST.get('ip')
        mac = request.POST.get('mac')
        comment = request.POST.get('comment')
        type = request.POST.get('type')
        
        obj = IP_field.objects.get(IP=ip)
        obj.MAC = mac
        obj.comment = comment
        obj.device_type = type
        obj.save()
        
        return HttpResponse('ok')