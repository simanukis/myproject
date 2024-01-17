from django.db import models
# ユーザー認証
from django.contrib.auth.models import User

# Create your models here.
# ユーザーアカウントのモデルクラス
class Account(models.Model):
    # ユーザー認証のインスタンス（１：１関係）
    user = models.OneToOneField(User, on_delete=models.CASCADE)
    
    # 追加フィールド
    syain_no = models.IntegerField(
                                    blank=True, 
                                    null=True,
                                    verbose_name="社員番号"
                                )
    syain_name = models.CharField(
                                    max_length=20,
                                    blank=True,
                                    null=True,
                                    verbose_name="社員名"
                                )
    syozoku = models.CharField(
                                max_length=10,
                                blank=True,
                                null=True,
                                verbose_name="所属"
                            )
    # login_id = models.CharField(
    #                            max_length=20,
    #                            blank=True,
    #                            null=True,
    #                            verbose_name="ログインID"
    #                        )
    # password = models.CharField(
    #                             max_length=48,
    #                             blank=True,
    #                             null=True,
    #                             verbose_name="パスワード"
    #                         )
    # email = models.CharField(
    #                             max_length=50,
    #                             blank=True,
    #                             null=True,
    #                             verbose_name="メールアドレス"
    #                        )
    authority = models.CharField(
                                max_length=50,
                                blank=True,
                                null=True,
                                verbose_name="権限"
                            )
    created_datetime = models.DateTimeField(
                                            auto_now_add=True,
                                            blank=True,
                                            null=True,
                                            verbose_name="登録日時"
                                        )
    regist_name = models.CharField(
                                    max_length=20,
                                    blank=True,
                                    null=True, 
                                    verbose_name="登録者名"
                                )
    updated_datetime = models.DateTimeField(
                                            auto_now_add=True,
                                            blank=True,
                                            null=True,
                                            verbose_name="登録日時"
                                        )
    update_name = models.CharField(
                                    max_length=20,
                                    blank=True,
                                    null=True,
                                    verbose_name="更新者名"
                                )
    
    def __str__(self):
        return self.user.username
    