"""GAS と同じ形で受ける窓口。

アプリは `GAS_URL + '?action=getNews'` と呼んでいる。
**同じ形で受けることで、切り替えは GAS_URL を1行変えるだけになる。**
戻すときも URL を戻すだけ。数十秒で元に戻せる。

道を細かく分けない（`/api/news` のようにしない）のは、そのため。
見た目より、**お客様を危険にさらさないこと**を優先している。
"""

from django.urls import path

from . import views

urlpatterns = [
    # GAS の doGet と同じ入口。?action=... で振り分ける。
    path("", views.窓口, name="gas-api"),
]
