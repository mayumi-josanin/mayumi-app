# 段階B の下ごしらえ — 正をサーバーへ移すには何が要るか

段階A（写しをサーバーへ）は 2026-08-18 に完了した。1,751件が入っている。
**ただし正はまだスプレッドシートで、アプリは GAS を見ている。**

段階Bは「どちらを正とするか」を切り替える段階。
**お客様の画面に影響が出る**ので、まず何が足りないかを数えた（2026-08-21）。

---

## いまの落差

| | 数 |
|---|---|
| GAS が受けている種類 | **99** |
| 管理アプリが投げる種類 | 37 |
| お客様アプリが投げる種類 | 12 |
| **サーバーが出せるもの** | **3**（health / entry-info / measurements） |

**サーバー側はまだ 3 つしか無い。**データは全部あるが、
アプリが話しかける口がほとんど無い。

> 99すべてを作る必要は無い。使われていないもの・管理アプリ専用のものがある。
> だが**お客様アプリの12種＋読み取りは、切り替えるなら必ず要る。**

## お客様アプリが使う12種（切り替えに必須）

    cancel                    注文の取り消し
    confirmReceipt            受け取り確認
    drawRewardGacha           特典のガチャ
    order                     注文
    recoverAccount            **復元**
    removeUserDeviceSession   端末の解除
    resetForgottenPasscode    **パスコードの再設定**
    syncUserDeviceSession     端末の記録合わせ（自動）
    syncUserRewardStatus      特典状態の同期（自動）
    unsubscribePush           通知の解除
    updateUser                会員情報の編集
    uploadImage               画像の保存

このうち **recoverAccount と resetForgottenPasscode は入り口そのもの。**
ここを間違えると、お客様が入れなくなる。

**syncUserDeviceSession と syncUserRewardStatus は自動処理で、
93日で8,579回**呼ばれている。切り替えたら、その負荷がサーバーに来る。

---

## 進め方の案（まだ決定ではない）

**一度に全部は切り替えない。**読むだけのものから、少しずつ移す。

| 段 | 何を | なぜ先か |
|---|---|---|
| B-1 | **読むだけ**（お知らせ・カレンダー・メニュー・商品・FAQ） | 壊れても表示が古くなるだけ。お客様は入れる |
| B-2 | 会員の**読み取り**（getInitialData など） | まだ書き込まないので、戻すのが楽 |
| B-3 | **書き込み**（updateUser・注文・スタンプ） | ここから先は戻すのが難しい |
| B-4 | **入り口**（recoverAccount・resetForgottenPasscode） | **最後。**間違えるとお客様が入れなくなる |

各段で必ず決めること:

- **戻し方**。何分で元に戻せるか。戻す判断は誰がするか
- **実施する時間帯**。お客様が使っていない時間
- **確かめ方**。切り替えた直後、何を見て「大丈夫」と判断するか

## 先に片付けておくとよいこと

### ① 写しを最新にする

いまの写しは 8/17〜18 に取り込んだもの。実際のシートとは十数件ずれている
（`gas/写しの突き合わせ.js` で数えられる）。切り替えの直前に一度そろえる。

**操作履歴だけは、そのまま取り込み直してはいけない。**
古い記録の自動削除で先頭が消え、行番号がずれている。
取り込み直すなら先に `records_auditlog` を空にすること。

### ② 「使っていない口」を数える

99のうち、実際に呼ばれているのはいくつか。操作履歴（`records_auditlog`）に
種別が残っているので数えられる。**作らなくてよいものが分かれば、仕事が減る。**

### ③ 札（トークン）の突き合わせ

[移行設計](移行設計.md#入り直しをゼロにするための約束)のとおり、
**GASが作った札を、サーバー側が同じと判定するか**を実データで確かめる。
これが合わない限り切り替えない。`ADMIN_TOKEN_SECRET` は変えない。

---

## いま言えること

**データは揃っている。足りないのはアプリとサーバーをつなぐ口。**
段階Bは「移行」というより「作り直し」に近い分量がある。

急ぐ理由は無い。サーバーは写しのまま置いておいても、
控えが毎日取れて、見張りが動いている。**止まってもお客様には何も起きない。**

---

## ① 使っていない口を数えた（2026-08-21）

**操作履歴（`records_auditlog` 794件・2026-05-17〜08-17）から数えた。**
GAS は5種類だけ記録から外している（`AUDIT_LOG_SKIP_TYPES`）ので、
**それ以外は呼ばれれば必ず残る。**つまり「一度も出てこない＝93日間使われていない」。

| | 数 |
|---|---|
| GAS が受ける種類（重複を除く） | **96** |
| 記録から外れている | 5 |
| 記録に出うる | 91 |
| **実際に出てきた** | **28** |
| **93日間で一度も呼ばれていない** | **63** |

### 実際に使われた28種（多い順）

    updateUser 155 / recoverAccount 140 / addBlog 67 / saveMenuRevenueRecord 61
    updateAdminRewardStatus 59 / resetForgottenPasscode 49 / deleteUser 42
    addCalendar 40 / deleteOrders 30 / updateBlog 29 / drawRewardGacha 27
    saveProductRevenueRecord 16 / updateCalendar 13 / mergeUsers 13 / deleteRow 10
    updateMenu 8 / saveRewardGachaConfig 8 / grantSurveyStamp 7
    runManualBackup 3 / restoreDeletedRecord 3 / issueTransferCode 2 / addMenu 2
    deleteProductRevenueRecord 2 / deleteMenuRevenueRecord 2 / deleteRows 1
    deleteCategory 1 / addCategory 1 / updateProduct 1
    パスコード再設定 2（GASの一覧に無い日本語の種別。別経路で書かれている）

### 読み取りは、この数え方では分からない

**`getNews` などの読み取りが「一度も出てこない」のは、使われていない意味ではない。**
記録されるのは `doPost` で、読み取りの多くは別経路（`doGet` / `action=`）を通る。
実際、お客様アプリは起動のたびに `getInitialData` を呼んでいるはずだが、
記録には1件も無い。**読み取り側は、別の方法で数える必要がある。**

### それでも分かったこと

**書き込み側は28種で足りる。**63種のうち多くは管理アプリの読み取りで、
段階Bでは後回しにできる。**99すべてを作る必要は無い。**

とくに**お客様アプリの書き込みは、実質7種**しか使われていない。

    updateUser / recoverAccount / resetForgottenPasscode / drawRewardGacha
    order / cancel / confirmReceipt

（`syncUserDeviceSession` と `syncUserRewardStatus` は記録外だが、
93日で8,579回。**最も呼ばれる2種**なので当然要る。）

### 次に数えること

**読み取りが実際にどれだけ呼ばれているか。**
操作履歴では分からないので、GAS 側に一時的に数える仕掛けを入れるか、
アプリのコードから「起動時に何を呼ぶか」を追う。


---

## ② 読み取りを数えた（2026-08-21）

操作履歴では分からないので、**アプリのコードから直接拾った。**
読み取りは `getFromGAS(action)` / `fetchFromGAS(action)` を通る。

### お客様アプリが読むもの — **12種**

| action | どこで |
|---|---|
| `getAppRuntimeConfig` | 起動時 |
| `getProducts` | 商品一覧 |
| `getNews` | お知らせ |
| `getPushNotices` | 通知の一覧 |
| `getRewardGachaConfig` | 特典ガチャ |
| `getUserRewardStatus` | 特典状態（会員ごと） |
| `getCalendar` | カレンダー |
| `getCustomerOrders` | 注文履歴（会員ごと） |
| `getSupportFaq` | 使い方チャット |
| `getUserDevices` | 端末の一覧（会員ごと） |
| `getRecoveryCandidates` | **復元の候補** |
| `getMenus` | メニュー一覧 |

### 管理アプリが読むもの — **25種**

    getAdminMenus / getAdminProducts / getAdminUsers / getAdminOrders
    getAdminCalendar / getAdminBlogs / getCategories / getAdminUserOrders
    getPushNotices / getProductRevenueRecords / getMenuRevenueRecords
    getAnalytics / getAdminSupportFaq / getSupportChatAnalytics
    getRewardGachaConfig / getPushUsers / getBackupStatus / getAppRuntimeConfig
    getAdminTrashItems / getAdminTemplates / getAdminSecurityConfig
    getAdminDashboardData / getAdminAuditLogs
    getFirebasePresenceToken / getFirebasePresenceAdminConfig

### 分かったこと

**`getInitialData` はどちらのアプリからも呼ばれていない。**
GAS は受け付けるが、実際には使われていない古い口。**作らなくてよい。**

（①で「記録に0件だから使われているはず」と考えたのは早合点だった。
　**本当に使われていなかった。**コードを見るまで断定しなくてよかった。）

---

## ①②を合わせた結論：作るのは **44種**

| | 数 |
|---|---|
| GAS が受ける種類 | 96 |
| **お客様アプリが使う（書き9＋読み12）** | **21** |
| **管理アプリが使う（書き…＋読み25）** | — |
| **合わせて実際に使われている** | **44**（重複を除く） |
| **どこからも呼ばれていない** | **52** |

**半分以上は作らなくてよい。**

### 切り替えの最小構成 — お客様アプリの21種

**まずここだけ作れば、お客様の側は切り替えられる。**管理アプリは
スプレッドシートを見続けてよい（院内でしか使わないので、多少遅くても困らない）。

    書き込み9  updateUser / recoverAccount / resetForgottenPasscode
              drawRewardGacha / order / cancel / confirmReceipt
              syncUserDeviceSession / syncUserRewardStatus
    読み取り12 getAppRuntimeConfig / getProducts / getNews / getPushNotices
              getRewardGachaConfig / getUserRewardStatus / getCalendar
              getCustomerOrders / getSupportFaq / getUserDevices
              getRecoveryCandidates / getMenus

**このうち読み取り12種は、すでにサーバーにデータが揃っている。**
お知らせ95・カレンダー143・メニュー13・商品9・FAQ46・通知96 は移行済み。
**作るのは「出す口」だけで、中身を用意する必要は無い。**

書き込み9種のうち、注文まわり（order / cancel / confirmReceipt）は
**注文管理シートが0件**なので、いま使われていない。実質6種。

> **次にやること: 札（トークン）の突き合わせ。**
> これが合わない限り、何を作っても切り替えられない。


---

## ③ 札（トークン）の突き合わせ（2026-08-21）

**絶対ルール「お客様の入り直しを発生させない」の要。**
発行済みの札はお客様の端末に入っている。切り替えたあとも同じ札が通らないと、
**全員がログイン画面に戻される。**

### 札の作り（`gas/管理者・お客様.js` から読み取った）

```
札  = base64url( JSON )
JSON = { t: 'member', u: 会員ID, exp: 期限(ミリ秒), sig: 署名 }
署名 = base64url( HMAC-SHA256( 'member|会員ID|期限', ADMIN_TOKEN_SECRET ) )
```

管理者の札も同じ仕組みで、署名のもとが `admin|期限` になるだけ。

| | |
|---|---|
| 作る | `makeMemberToken_()` / `makeAdminToken_()` |
| 検証 | `verifyMemberToken_()` |
| 署名 | `adminSign_()` |
| 有効期間（管理者） | 12時間（`ADMIN_TOKEN_TTL_MS`） |

### 確かめたこと

`gas/札を確かめる.js` の `札を確かめる()` で、**存在しない会員ID**
（`MYM-0000-TEST`）と固定の期限で見本の札を作らせた。
実データは変えず、この札では誰もログインできない。

    署名のもと : member|MYM-0000-TEST|1893456000000
    札        : eyJ0IjoibWVtYmVy...（base64url）
    自分で検証  : 通った

**Python 側で札を読めることは確かめた。**

- `base64url` を復号して JSON に戻せる（GAS は `=` を残すので補う必要がある）
- `t / u / exp / sig` が取り出せる
- **署名のもとの文字列を、同じ形に組み立てられる**

### **まだ確かめていないこと（ここが本番）**

**署名そのものが一致するかは、まだ確かめていない。**

理由: **サーバーに `ADMIN_TOKEN_SECRET` が渡っていない。**
いまサーバーが持っている秘密は `DJANGO_SECRET_KEY` だけ。

    docker compose exec web python -c "..."
    → 環境変数の名前: ['DJANGO_SECRET_KEY']

### 次にやること

1. **`ADMIN_TOKEN_SECRET` をサーバーへ渡す**（`.env` に足す）
   - 値は院長が用意する。**私は見ない。**
   - `.env` は `.gitignore` 済み。**コードに書かない。**
   - 過去に画面へ秘密を出してしまった失敗がある（2026-08-16）。
     **必ず `scp` ＋ `-File` で渡し、画面に出さない。**
2. サーバー側に `verify_member_token()` を書く
3. **上の見本の札を実際に検証させて、`MYM-0000-TEST` が返ることを確かめる**
4. さらに **本物の札**（お客様の端末にあるもの）でも1つ確かめる

**3と4が通るまで、段階Bには進まない。**

> `ADMIN_TOKEN_SECRET` は変えない。変えると発行済みの札が全部無効になり、
> 全員が入り口に戻される（CLAUDE.md 絶対ルール1）。


### 結果：**一致した**（2026-08-21 追記 → 2026-08-22 確認）

`ADMIN_TOKEN_SECRET` をサーバーへ渡し、サーバー側で HMAC-SHA256 を計算して
GAS の署名と比べた。

| 署名のもと | GAS | サーバー | |
|---|---|---|---|
| `admin\|1893456000000` | `0bWBVhv3WvGquKq34othsepm5Rdn_nP-eSUifEDFUp4=` | 同じ | **一致** |
| `member\|MYM-0000-TEST\|1893456000000` | `e_LOXY7kgSYmgxnL7CqQv15fzbpijFvTw95O6gMRrvc=` | 同じ | **一致** |

**サーバーは、GASが作った札を同じと判定できる。**
`ADMIN_TOKEN_SECRET` を変えない限り、切り替えてもお客様は入り直しにならない。

#### 途中で1回、間違った結論を出しかけた

最初、会員の札が「署名不一致」と出た。末尾1文字だけ違っていた（`vc=` と `vg=`）。

原因は**私が実行ログの画面から札の文字列を写し取ったときの読み違い**。
`c` と `g` を取り違えていた。秘密も署名の作り方も正しかった。

**管理者の札が完全一致したことで、秘密と作り方が正しいと分かった。**
片方だけで判断していたら「秘密が違う」と誤って結論づけていた。
**2つ以上で確かめる。**

> **画面の文字を人（や私）が写すと読み違える。**
> `l` と `I`、`0` と `O`、`c` と `g`。
> 突き合わせは、写さずに機械同士で比べるか、**別々の2つで確かめる。**

#### サーバーへ渡すには docker-compose.yml も直す必要があった

`.env` に書くだけでは足りない。**`docker-compose.yml` の `environment:` に
列挙されたものだけがコンテナへ渡る。**

```yaml
  # GAS が札（トークン）の署名に使っている秘密。**変えないこと。**
  ADMIN_TOKEN_SECRET: ${ADMIN_TOKEN_SECRET:-}
```

最初これを忘れて `.env` に書いただけで再起動し、
「入っているか: False」となって一度つまずいた。

---

## ③まで終わった。次は？

| | 状態 |
|---|---|
| ① 使っていない口を数える | **済**。96種のうち作るのは44種 |
| ② 読み取りを数える | **済**。お客様アプリは21種で足りる |
| ③ 札の突き合わせ | **済**。**一致した** |
| ④ 読み取り12種の口を作る | これから |
| ⑤ 書き込み6種の口を作る | その後 |
| ⑥ 切り替え | 最後。**戻し方を決めてから** |

**④からは「作る」段階。**ここまでは読むだけだった。


---

## ④ 読み取りの口を作る（2026-08-22 着手）

### 決めたこと：**GASと同じ形にする**

アプリは `GAS_URL + '?action=getNews'` と呼んでいる。
サーバーも**同じ形で受ける**ことにした（院長の判断）。

    切り替え … アプリの GAS_URL を1行変えるだけ
    戻す   … URL を元に戻すだけ（数十秒）

見た目は今風ではないが、**お客様を危険にさらさないことを優先する。**
`/api/news` のような形にすると、アプリ側を12か所以上書き換えることになり、
戻すときも同じだけ大がかりになる。

### 着手して分かったこと：**同じ名前の関数が5つ、二重に定義されている**

**GASは、同じ名前の関数を後から読んだほうで上書きする。**
つまり**前のほうは一度も動いていない。**

| 関数 | 1つ目 | 2つ目（**こちらが動く**） |
|---|---|---|
| `getBlogNews` | 3279行 | **8668行** |
| `getCalendarEvents` | 3344行 | **8292行** |
| `getAdminBlogs` | 4938行 | **8737行** |
| `getAdminCalendar` | 6558行 | **8348行** |
| `handleUpdateOrder` | 3201行 | **4829行** |

**移すときに1つ目を写すと、動きが変わる。**
中身は似ているが同じではない。たとえば `getBlogNews` は、
2つ目だけが `ensureSortOrderColumn_` と `ensureNoticeDeletedAtColumn_` を見ていて、
返す項目にも `rowIdx` `linkUrl` `linkButtonText` が増えている。

> **移す前に必ず「どちらが動いているか」を確かめる。**
> `grep -n "^function 名前"` で行番号を出し、**大きいほうを読む。**

### `getNews` が返す形（実物・8668行）

```json
{ "status": "ok",
  "news": [ { "rowIdx": 2, "date": "…", "title": "…", "category": "…",
              "type": "お知らせ|ブログ", "icon": "📢", "body": "…",
              "image": "…", "imageUrl": "…", "imageUrls": ["…"],
              "updatedAt": "…", "linkUrl": "…", "linkButtonText": "…" } ] }
```

**`image` と `imageUrl` は同じ値を2つ入れている。**古い呼び名を残したまま
新しい名前を足したため。**片方だけにするとアプリが崩れる恐れがあるので、
そのまま両方返す。**


### 1つ目 `getNews` を作った結果（2026-08-22）

`server/apps/gasapi/` を作り、`?action=getNews` で受けられるようにした。
`config/urls.py` では **measurements の後ろに置く**こと。先に置くと
`health` / `measurements` まで拾ってしまう。

突き合わせた結果:

| | GAS | サーバー | |
|---|---|---|---|
| status | ok | ok | 一致 |
| **項目の数** | **16** | **16** | **一致** |
| 項目の名前 | rowIdx / date / title / category / type / icon / body / image / imageUrl / imageUrls / updatedAt / linkUrl / linkButtonText / publishAt / noticeStatus / sortOrder | 同じ | **一致** |
| categories | 7件 | 7件 | 一致 |
| **news の件数** | **88** | **86** | **2件ちがう** |

#### 途中で自分の間違いを1つ直した

最初、除外条件に `published`（公開設定）を入れていた。**GAS は見ていない。**
GAS が見るのは「お知らせ一覧公開」のほうだけ。
入れたままだと、GASでは出ているものがサーバーでは消える。

    GAS（8668行）が外している4つ:
      notice_delisted_at が入っている      （お知らせ一覧削除日時）
      deleted / deleted_at が入っている     （削除状態・削除日時）
      notice_listed が「非公開」            （お知らせ一覧公開）
      publish_at がまだ来ていない           （公開開始日時）

**条件はコードから写す。似ているから同じだろう、で書かない。**

#### 残る2件の差

**まだ合っていない。**サーバーの写しが 8/17 のもので、
シートはその後 95→97行に増えている（8/21 の突き合わせで確認済み）。
**写しが古いことが原因の可能性が高いが、確かめていない。**

次にやること: 写しを最新にしてから、もう一度数える。
それでも合わなければ、**1件ずつ突き合わせて原因を探す。**

> **件数が合うまで、この口は使わない。**
> 2件足りないまま切り替えると、お客様には「記事が2つ消えた」ように見える。

