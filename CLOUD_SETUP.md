# クラウド導入手順

1. Supabaseで専用プロジェクトを新規作成。SQL Editorで schema.sql を実行。
2. RenderでこのGitHubリポジトリを接続し、Web Serviceを作成（render.yamlを参照）。
3. RenderのEnvironmentに SUPABASE_URL、SUPABASE_SERVICE_ROLE_KEY、ADMIN_PASSWORD（16文字以上）、OPENAI_API_KEY を設定。秘密鍵はGitHubに保存しない。
4. デプロイ後 /login にアクセスし管理者ログイン。募集要項PDFを解析し、フォーム項目を確認して公開。
5. 企業には /f/<フォームID> のURLを案内。

**重要**: 本版は試作版です。外部企業の個人情報・雇用関連書類を収集する前に、CSRF対策、ログイン試行制限、ファイルのウイルス検査、保存期間・削除手順、監査ログ、バックアップ、プライバシー告知、リージョン確認を実施してください。Render等の有料プラン費用も確認してください。
