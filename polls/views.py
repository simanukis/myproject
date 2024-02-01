from django.shortcuts import render
from django.views.generic import TemplateView #テンプレートタグ
from .forms import AccountForm, AddAccountForm #ユーザーアカウントフォーム

# ログイン・ログアウト処理に利用
from django.contrib.auth import authenticate, login, logout
from django.http import HttpResponseRedirect, HttpResponse, JsonResponse
from django.urls import reverse
from django.contrib.auth.decorators import login_required

# 給与集計処理に使用


# HTTPResponseクラスをインポート
from django.http import HttpResponse

# application/totalling.pyをインポートする
# from .application import totalling
# Create your views here.

# ログイン
def Login(request):
    # POST
    if request.method == 'POST':
        # フォーム入力のユーザーID・パスワード取得
        ID = request.POST.get('login_id')
        Pass = request.POST.get('password')
        
        # Djangoの認証機能
        user = authenticate(username=ID, password=Pass)
        
        # ユーザー認証
        if user:
            #ユーザーアクティベート判定
            if user.is_active:
                # ログイン
                login(request,user)
                # ホームページ遷移
                return HttpResponseRedirect(reverse('home'))
            else:
                # アカウント利用不可
                return HttpResponse("アカウントが有効ではありません")
        # ユーザー認証失敗
        else:
            return HttpResponse("ログインIDまたはパスワードが間違っています")
    # GET
    else:
        return render(request, 'polls/login.html')    

#ログアウト
@login_required
def Logout(request):
    logout(request)
    # ログイン画面遷移
    return HttpResponseRedirect(reverse('login'))

#ホーム
@login_required
def home(request):
    params = {"login_ID":request.user,}
    return render(request, "polls/index.html",context=params)

#新規登録
class  AccountRegistration(TemplateView):
    def __init__(self):
        self.params = {
            "AccountCreate":False,
            "account_form": AccountForm(),
            "add_account_form": AddAccountForm(),
        }
    
    # Get処理
    def get(self,request):
        self.params["account_form"] = AccountForm()
        self.params["add_account_form"] = AddAccountForm()
        self.params["AccountCreate"] = False
        return render(request,"polls/register.html",context=self.params)

    # Post処理
    def post(self,request):
        self.params["account_form"] = AccountForm(data=request.POST)
        self.params["add_account_form"] = AddAccountForm(data=request.POST)
        
        # フォーム入力の有効検証
        if self.params["account_form"].is_valid() and self.params["add_account_form"].is_valid():
            # アカウント情報をDB保存
            account = self.params["account_form"].save()
            # パスワードをハッシュ化
            account.set_password(account.password)
            # ハッシュ化パスワード更新
            account.save()
            
            # 下記追加情報
            # 下記操作のため、コミットなし
            add_account = self.params["add_account_form"].save(commit=False)
            
            # AccountForm & AddAccountForm 1vs1 紐付け
            add_account.user = account
            
            # モデル保存
            add_account.save()
            
            # アカウント作成情報更新
            self.params["AccountCreate"] = True
        else:
            # フォームが有効でない場合
            print(self.params["account_form"].errors)
            
        return render(request,"polls/register.html",context=self.params)

# 給与集計
@login_required
def totalling(request):
    params = {"login_ID":request.user,}
    return render(request, "polls/totalling.html",context=params)

# ファイルアップロード
def ajax_file_send(request):
    print("OK")
    d = {}
    return JsonResponse(d)

# View関数を任意に定義
# def index(request):
    # 変数設定
    
    # 出力
#    return render(request, 'polls/index.html')

# ajaxでurl指定したメソッド
# def totalling(req):
#    if req.method == 'GET':
        # totalling.pyのwrite_csv()メソッドを呼び出す。
        # ajaxで送信したデータのうち'input_data'を指定して取得する。
#        totalling.write_csv(req.GET.get("input_data"))
#        return HttpResponse()
