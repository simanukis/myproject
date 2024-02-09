from django import forms
# from bootstrap4.widgets import RadioSelectButtonGroup

from django.contrib.auth.models import User
from .models import Account

# フォームクラス作成
# アカウント
class AccountForm(forms.ModelForm):
    # パスワード入力：非表示対応
    password = forms.CharField(widget=forms.PasswordInput(),label="パスワード")

    class Meta():
        # ユーザー認証
        model = User
        # フィールド指定
        fields = ('username','email','password')
        # フィールド名指定
        labels = {'username':"ログインID",'email':"メールアドレス"}

class AddAccountForm(forms.ModelForm):
    class Meta():
        # モデルクラスを指定
        model = Account
        fields = ('syain_name','syozoku',)
        labels = {'syain_name':"社員名",'syozoku':"所属",}

# ファイルアップロード
class FileUploadForm(forms.Form):
    testfile = forms.FileField()
    
        
# ファイルアップロード
# class FileUploadForm(forms.Form):
#    file = forms.FileField(label='ファイル')
#    FileType = forms.ChoiceField(
#        help_text="Select file type",
#        choices=(('xls', 'xls'), ('xlsx', 'xlsx')),
#        initial='xls',
#        required=True,
#        widget=RadioSelectButtonGroup,
#    )