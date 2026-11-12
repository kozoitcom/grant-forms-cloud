# Google Drive方式・設定手順（開発ブランチ）

このブランチはSupabaseを使わず、Google Drive上の専用フォルダにフォーム、回答、添付書類を保存する試作版です。**本番公開前にセキュリティレビューが必要です。**

1. Google Driveに専用の空フォルダを作り、そのURLの `folders/` の後のIDを控える。フォルダは外部共有しない。
2. Google Cloud ConsoleのOAuthクライアント（ウェブアプリケーション）に承認済みリダイレクトURI `https://grant-forms-cloud.onrender.com/admin/google/callback` を追加。承認済みJavaScript生成元はこのサーバー方式では不要。
3. Google CloudのOAuth同意画面でテストユーザーを登録。Drive APIを有効化し、`https://www.googleapis.com/auth/drive` スコープを追加（機密性の高いスコープであり、外部公開にはGoogle審査が必要となる場合があります）。
4. RenderのEnvironmentに `ADMIN_PASSWORD`（16文字以上）、`OPENAI_API_KEY`、`GOOGLE_CLIENT_ID`、`GOOGLE_CLIENT_SECRET`、`GOOGLE_DRIVE_FOLDER_ID`、`GOOGLE_REDIRECT_URI` を設定し、Save, deploy。
5. `/login` からログインし、`/admin/google/connect` にアクセス。Googleで許可した後に表示されるリフレッシュトークンを **Renderの `GOOGLE_REFRESH_TOKEN`** にのみ登録し、Save, deploy。トークンをGitHub、チャット、スクリーンショットに貼らない。
6. `/admin` にアクセスしてPDFからフォーム下書きを生成、編集・公開する。

**注意**: Google Drive全体へのアクセス権を要求します。データは指定フォルダ内に保存する実装ですが、OAuth権限自体はDrive全体に及びます。アクセスを限定したい場合は別の認証設計が必要です。OAuthテストモードのリフレッシュトークンは7日程度で失効する場合があります。無料プランの制約やAPIクォータも確認してください。

**未対応**: CSRF防御、ログイン試行制限、マルウェア検査、保存期限・削除処理、監査ログ、バックアップ、複数管理者、同時更新対策。個人情報・雇用書類を実際の企業から収集する前に対策してください。
