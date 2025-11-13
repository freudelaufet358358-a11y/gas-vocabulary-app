# ⚡️ High-Performance GAS Vocabulary App

Googleスプレッドシートをバックエンドに使用した、高速かつオフライン対応の単語学習Webアプリです。
GAS特有のレスポンス遅延を解消するため、ローカルキャッシュ戦略とSPA（シングルページアプリケーション）構成を採用しています。

![Animation](https://github.com/user-attachments/assets/be23fa1f-4cec-46d0-8cd5-0d712eadc880)

## 🚀 特徴 (Key Features)

* **ゼロ・レイテンシー:** 初回起動時にデータをキャッシュし、回答時のサーバー通信（`google.script.run`）を排除。ネイティブアプリ並みの即答性を実現。
* **オフライン完全対応:** 通信が切れても学習を継続可能。結果はローカルに保存され、オンライン復帰時に自動同期されます。
* **Spreadsheet Backend:** データベース不要。Googleスプレッドシートで単語リストを管理するだけで、即座にアプリに反映。
* **SPA Architecture:** 画面遷移にページリロードを挟まないモダンなUI設計。

## 🛠 技術スタック (Tech Stack)

* **Backend:** Google Apps Script (GAS)
* **Database:** Google Sheets
* **Frontend:** HTML5, CSS3, Vanilla JavaScript
* **Dev Tools:** VS Code, clasp, Git

## 📦 セットアップ (Setup)

1.  Googleスプレッドシートを新規作成
2.  スクリプトエディタを開き、このリポジトリのコードを反映（`clasp push` 推奨）
3.  「デプロイ」からWebアプリとして公開
4.  スプレッドシートに「単語リスト」シートを作成（自動生成機能あり）

## 📝 開発者向けノート

高校3年生の受験勉強の合間に開発しました。
「GASは遅い」という常識を覆すべく、Cache APIとLocalStorageを駆使して高速化に挑戦しています。
詳しい技術解説はZennの記事をご覧ください： [記事のURLをここに貼る]

## 📄 License

MIT
