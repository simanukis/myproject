from django.shortcuts import render

# Create your views here.
# HTTPResponseクラスをインポート
from django.http import HttpResponse

# View関数を任意に定義
def index(request):
    # 変数設定
    
    # 出力
    return render(request, 'polls/index.html')
