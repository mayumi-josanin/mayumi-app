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

