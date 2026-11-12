# Google Apps Script版（開発用）

**目的**: Render・Supabaseを使わず、GitHubでコード管理し、Google Apps ScriptでWebフォームを公開、Google Sheetsに回答を保存、Google Driveに添付書類を保存します。

## 構成
- `gas/Code.gs`: API・管理機能・保存
- `gas/Index.html`: 管理者画面と企業向け入力画面
- `gas/appsscript.json`: 権限と実行設定
- GitHub: ソースの変更履歴（Google Apps Scriptへの反映は未自動化）

## 初回設定（Googleアカウントの操作が必要）
1. https://script.google.com/home/projects/create で独立したApps Scriptプロジェクトを作る。
2. `Code.gs` と `Index.html` を追加し、`appsscript.json` を設定する（マニフェストを表示するにはプロジェクト設定で「appsscript.json をエディタで表示」をオン）。
3. プロジェクトの **設定 → スクリプト プロパティ** で次を登録する。
   - `DRIVE_FOLDER_ID` = `1XJovIo8J46rq_VQl-u_pVdjzQ2cAjsm_`
   - `ADMIN_PASSWORD` = 新しく作った16文字以上の秘密パスワード
   - `OPENAI_API_KEY` = **今は不要**。PDF解析機能は未実装。
4. **デプロイ → 新しいデプロイ → ウェブアプリ** を選択し、実行ユーザーを **自分** にする。アクセス範囲は **自分のみ** にして、まず管理画面を検証する。
5. 管理画面でフォームを作成・編集。回答フォームは `?form=<フォームID>`。
6. 実際の外部企業に配布する前に、公開範囲、アクセス制御、プライバシー告知、添付ファイル検査、レート制限、監査ログ、バックアップを整える。

## 現時点でできること
- 管理者パスワードによるログイン（1時間の一時セッション）
- フォーム項目のJSON編集、公開・非公開
- 入力フォームの表示、回答受付
- 添付ファイルをGoogle Driveへ保存（1ファイル5MB以内）
- Google Sheetsにフォーム・回答・添付メタデータを保存
- 管理画面で回答詳細を閲覧

## 未実装・本番利用前に必須
- 募集要項PDFの本文抽出とAIによる自動フォーム生成
- 公式Excel・Wordへの転記（第2段階）
- CSRF対策、ボット・連投対策、ログイン試行制限
- アップロードのウイルス検査と内容検証、企業ごとの閲覧権限
- データ削除・保存期間・バックアップ・監査ログ
- GitHubからApps Scriptへの自動デプロイ（clasp + GitHub Actions）
- Apps Scriptの容量・実行時間制限に対する検証

**注意**: この版は動作未検証の試作です。実在企業の申請データ・個人情報は投入しないでください。Googleの無料枠と利用規約は変更される可能性があります。
