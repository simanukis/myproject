from django import forms
from django.contrib.auth.models import User
from .models import Account

# フォームクラス作成
class AccountForm(forms.ModelForm):
    # パスワード入力：非表示対応
    password = forms.CharField(widget=forms.PasswordInput(),label="パスワード")

    class Meta():
        # ユーザー認証
        model = User
        # フィールド指定
        fields = ('login_id','email','password')
        # フィールド名指定
        labels = {'login_id':"ログインID",'email':"メールアドレス"}

class AddAccountForm(forms.ModelForm):
    class Meta():
        # モデルクラスを指定
        model = Account
        fields = ('syain_name','syozoku',)
        labels = {'syain_name':"社員名",'syozoku':"所属",}