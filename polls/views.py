from django.shortcuts import render

# Create your views here.
# HTTPResponseクラスをインポート
# from django.http import HttpResponse

# View関数を任意に定義
def index(request):
    return render(request, 'polls/index.html')
