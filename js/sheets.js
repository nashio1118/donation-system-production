// Google Sheets連携クラス
class GoogleSheetsManager {
    constructor() {
        this.settings = null;
        this.cacheTtlMs = 300000; // 5分キャッシュ
        // 管理者用クイック操作（URLで保存/クリア）
        try {
            const adminParams = new URLSearchParams(window.location.search || '');
            const clearApi = adminParams.get('clearApi');
            const setApi = adminParams.get('setApi');
            const clearSettings = adminParams.get('clearSettings');
            if (clearApi === '1') {
                try { localStorage.removeItem('scriptUrl'); console.log('scriptUrlをクリアしました'); } catch (_) {}
            }
            if (setApi && setApi.startsWith('http')) {
                try { localStorage.setItem('scriptUrl', setApi); console.log('scriptUrlを保存しました'); } catch (_) {}
            }
            if (clearSettings === '1') {
                try {
                    localStorage.removeItem('settings_cache');
                    localStorage.removeItem('form_settings_cache');
                    console.log('設定キャッシュをクリアしました');
                } catch (_) {}
            }
        } catch (_) {}
        // GAS URLの解決（簡単運用: 固定値を基本としつつ、必要に応じて上書き可）
        // 優先順位: 1) URLパラメータ ?api=... 2) localStorage.scriptUrl 3) 既存の固定値
        const __resolveScriptUrl = () => {
            try {
                const params = new URLSearchParams(window.location.search || '');
                const urlFromQuery = params.get('api');
                if (urlFromQuery && typeof urlFromQuery === 'string' && urlFromQuery.startsWith('http')) {
                    return urlFromQuery;
                }
            } catch (_) {}
            try {
                const stored = localStorage.getItem('scriptUrl');
                if (stored && typeof stored === 'string' && stored.startsWith('http')) {
                    return stored;
                }
            } catch (_) {}
            return 'https://script.google.com/macros/s/AKfycbww215RU1fmTXow-k6vCdv1cMo4cNf_0aBVjKwjIddxN1CpVE-Jx6DYC-h8p8mAG9XN4Q/exec'; // 新しいGAS URLをここに設定してください
        };
        this.scriptUrl = __resolveScriptUrl();
        // デモモード判定（URLの ?demo=1 または sessionStorage のフラグ）
        this.isDemo = this.isDemoMode();
        this.loadSettings();
    }

    // デモモードかどうかを判定
    isDemoMode() {
        try {
            const params = new URLSearchParams(window.location.search || '');
            const urlFlag = params.get('demo') === '1';
            if (urlFlag) {
                try { sessionStorage.setItem('isDemo', '1'); } catch (_) {}
            }
            const sessionFlag = (() => { try { return sessionStorage.getItem('isDemo') === '1'; } catch (_) { return false; } })();
            return urlFlag || sessionFlag;
        } catch (_) {
            return false;
        }
    }

    // ローカルキャッシュから読み込み（存在すれば即時反映）
    loadSettingsFromCache() {
        try {
            const raw = localStorage.getItem('settings_cache');
            if (!raw) return false;
            const cached = JSON.parse(raw);
            if (!cached || !cached.timestamp || !cached.data) return false;
            const fresh = (Date.now() - cached.timestamp) < this.cacheTtlMs;
            if (fresh) {
                this.settings = cached.data;
                try { this.updateFormSettings(); } catch (_) {}
                return true;
            }
        } catch (_) {}
        return false;
    }

    saveSettingsToCache(data) {
        try {
            localStorage.setItem('settings_cache', JSON.stringify({ timestamp: Date.now(), data }));
        } catch (_) {}
    }

    // 設定を読み込み（SWR: 即時キャッシュ→裏で最新取得）
    async loadSettings() {
        try {
            // デモモードではGASからの取得を行わず、内蔵デフォルト設定のみを使用
            if (this.isDemo) {
                console.log('デモモードのため、設定読み込みをスキップしてデフォルト設定を適用します');
                this.loadDefaultSettings();
                return;
            }
            // URLが設定されていない場合はデフォルト設定を使用
            if (!this.scriptUrl || this.scriptUrl === '') {
                console.warn('Google Apps ScriptのURLが設定されていません。デフォルト設定を使用します。');
                this.loadDefaultSettings();
                return;
            }
            // まずキャッシュを適用（あれば即時反映）
            const hadCache = this.loadSettingsFromCache();

            // 裏で最新取得（SWR）。キャッシュが無ければawaitしてから反映
            const fetchLatest = async () => {
                console.log('設定を読み込み中...');
                const response = await fetch(`${this.scriptUrl}?action=getSettings`);
                if (!response.ok) {
                    console.warn(`設定読み込みでHTTPエラー: ${response.status} ${response.statusText}`);
                    throw new Error(`HTTP エラー: ${response.status} ${response.statusText}`);
                }
                const result = await response.json();
                console.log('設定読み込みレスポンス:', result);
                if (result && result.success && (result.giftRules || result.amountOptions || result.burdenSettings)) {
                    const latest = {
                        giftRules: result.giftRules || {},
                        amountOptions: result.amountOptions || {},
                        burdenSettings: result.burdenSettings || {}
                    };
                    this.settings = latest;
                    this.saveSettingsToCache(latest);
                    try { this.updateFormSettings(); } catch (_) {}
                } else if (result && result.settings) {
                    this.settings = result.settings;
                    this.saveSettingsToCache(this.settings);
                    try { this.updateFormSettings(); } catch (_) {}
                } else {
                    throw new Error(result && result.error ? result.error : '設定の読み込みに失敗しました');
                }
            };

            if (hadCache) {
                // 非同期で更新（失敗してもUIはキャッシュで表示済み）
                fetchLatest().catch(err => console.warn('設定最新化エラー:', err));
                return true;
            } else {
                // キャッシュが無い場合は取得を待つ
                await fetchLatest();
                return true;
            }
            
        } catch (error) {
            console.error('設定の読み込みエラー:', error);
            console.warn('デフォルト設定を使用します');
            this.loadDefaultSettings();
        }
    }

    // デフォルト設定を読み込み
    loadDefaultSettings() {
        this.settings = {
            giftRules: {
                towel: { minAmount: 10000, name: 'タオル', description: '高級タオル' },
                sweets_large: { minAmount: 10000, name: 'お菓子大', description: '高級お菓子セット' },
                keychain: { minAmount: 5000, name: 'キーホルダー', description: 'オリジナルキーホルダー' },
                sweets_small: { minAmount: 5000, name: 'お菓子小', description: 'お菓子セット' },
                clearfile: { minAmount: 1000, name: 'クリアファイル', description: 'オリジナルクリアファイル' }
            },
            amountOptions: {
                amount_10000: { amount: 10000, displayName: '¥10,000' },
                amount_5000: { amount: 5000, displayName: '¥5,000' }
            },
            burdenSettings: {
                burden_fee: { name: '負担金', amount: 40000, description: '1人分の負担金' },
                encouragement_fee: { name: '奨励会負担金', amount: 1000, description: '奨励会負担金' }
            }
        };
        this.updateFormSettings();
    }

    // フォーム設定を更新
    updateFormSettings() {
        if (!this.settings) return;

        // 1ページ目の負担金設定を更新
        this.updateBurdenFeeSettings();
        
        // 2ページ目の金額オプションを更新
        this.updateAmountOptions();
        
        // 返礼品オプションを更新
        this.updateGiftOptions();
    }

    // 負担金設定を更新
    updateBurdenFeeSettings() {
        if (!this.settings.burdenSettings) return;

        const burdenFee = this.settings.burdenSettings.burden_fee;
        const encouragementFee = this.settings.burdenSettings.encouragement_fee;

        if (burdenFee) {
            // 1ページ目の負担金表示を更新
            const burdenFeeLabel = document.querySelector('label[for="burdenFee"]');
            if (burdenFeeLabel) {
                burdenFeeLabel.textContent = `負担金${burdenFee.amount.toLocaleString()}円`;
            }
        }

        if (encouragementFee) {
            // 激励会協力金の表示を更新
            const encouragementText = `激励会協力金（${encouragementFee.amount.toLocaleString()}円）`;
            const encouragementLabelById = document.getElementById('encouragementFeeLabel');
            if (encouragementLabelById) encouragementLabelById.textContent = encouragementText;
            const encouragementLabelByFor = document.querySelector('label[for="encouragementFee"]');
            if (encouragementLabelByFor) encouragementLabelByFor.textContent = encouragementText;
        }
    }

    // 金額オプションを更新
    updateAmountOptions() {
        if (!this.settings.amountOptions) return;

        // 2ページ目の金額選択肢を更新
        const amountSelects = document.querySelectorAll('.donor-amount-type');
        amountSelects.forEach(select => {
            // 既存のオプションをクリア（「その他」は保持）
            const otherOption = select.querySelector('option[value="other"]');
            const header = '<option value="">金額を選択</option>';
            // 生成候補を先に文字列で組み立て
            let html = header;
            
            // 設定から金額オプションを追加
            Object.entries(this.settings.amountOptions).forEach(([id, option]) => {
                html += `<option value="${option.amount}">${option.displayName}</option>`;
            });
            
            // 「その他」オプションを追加
            if (otherOption) html += otherOption.outerHTML;

            // 同一であればスキップ（不要なreflow抑制）
            if (select.__lastOptionsHtml !== html) {
                select.innerHTML = html;
                select.__lastOptionsHtml = html;
            }
        });
    }

    // 返礼品オプションを更新
    updateGiftOptions() {
        if (!this.settings.giftRules) return;

        // 返礼品オプションは条件分岐なしで全件表示
        window.donationCalculator = {
            generateGiftOptions: (_amount) => {
                let options = '<option value="">返礼品なし</option>';
                Object.entries(this.settings.giftRules).forEach(([id, rule]) => {
                    // 値はIDではなく「返礼品名」を使用
                    options += `<option value="${rule.name}">${rule.name}</option>`;
                });
                return options;
            },
            generateClearfileSheetOptions: () => {
                let options = '<option value="">枚数を選択</option>';
                for (let i = 1; i <= 50; i++) {
                    options += `<option value="${i}">${i}枚</option>`;
                }
                return options;
            }
        };
    }

    // フォームデータを送信
    async submitFormData(formData) {
        try {
            console.log('submitFormData関数開始');
            console.log('this.scriptUrl値:', this.scriptUrl);
            console.log('this.scriptUrl型:', typeof this.scriptUrl);
            console.log('this.scriptUrl長さ:', this.scriptUrl ? this.scriptUrl.length : 'null/undefined');
            
            // デモモード判定（URL または sessionStorage または事前判定）
            const params = new URLSearchParams(window.location.search || '');
            const isDemoMode = this.isDemo || params.get('demo') === '1' || (function(){ try { return sessionStorage.getItem('isDemo') === '1'; } catch(_) { return false; } })();
            if (isDemoMode) {
                console.log('デモモードのため実送信をスキップします');
                await new Promise(resolve => setTimeout(resolve, 500));
                return { success: true, message: 'デモモード: 送信をスキップしました。' };
            }

            // URLが設定されていない場合はエラーを返す
            if (!this.scriptUrl || this.scriptUrl === '') {
                console.error('URLチェック失敗: scriptUrl =', this.scriptUrl);
                throw new Error('Google Apps ScriptのURLが設定されていません');
            }

            console.log('データ送信開始:', formData);
            
            // CORS問題のため、no-cors + keepalive で送信し、待たずに成功を返す（UX高速化）
            console.log('fetch開始 - URL:', this.scriptUrl);
            console.log('送信データ(概要): 基本情報/寄付者データ');
            try {
                fetch(this.scriptUrl, {
                    method: 'POST',
                    mode: 'no-cors',
                    keepalive: true,
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ action: 'submitFormData', formData })
                }).catch(err => console.warn('送信中エラー(非同期):', err));
            } catch (err) {
                console.warn('送信開始エラー:', err);
            }
            // すぐに成功を返す（サーバー処理はバックグラウンド継続）
            return { success: true, message: '送信を開始しました。スプレッドシートと個票を後ほどご確認ください。' };
            
        } catch (error) {
            console.error('送信エラー:', error);
            
            // ネットワークエラーやCORSエラーの場合の詳細メッセージ
            let errorMessage = error.message;
            if (error.name === 'TypeError' && error.message.includes('fetch')) {
                errorMessage = 'Google Apps Scriptとの通信に失敗しました。URLを確認してください。';
            } else if (error.message.includes('CORS')) {
                errorMessage = 'CORS設定に問題があります。Google Apps Scriptの設定を確認してください。';
            }
            
            return { success: false, error: errorMessage };
        }
    }

    // フォーム設定キャッシュ
    loadFormSettingsFromCache() {
        try {
            const raw = localStorage.getItem('form_settings_cache');
            if (!raw) return null;
            const cached = JSON.parse(raw);
            if (!cached || !cached.timestamp || !cached.data) return null;
            const fresh = (Date.now() - cached.timestamp) < this.cacheTtlMs;
            return fresh ? cached.data : null;
        } catch (_) { return null; }
    }
    saveFormSettingsToCache(data) {
        try { localStorage.setItem('form_settings_cache', JSON.stringify({ timestamp: Date.now(), data })); } catch (_) {}
    }

    // フォーム設定を取得（SWR: 即返し→裏で最新）
    async getFormSettings() {
        try {
            // デモモードではフォーム設定もローカルデフォルトを返す
            if (this.isDemo) {
                console.log('デモモードのため、フォーム設定はデフォルトを返します');
                return {
                    formTitle: '奉加帳提出報告フォーム',
                    formSubtitle: '〇〇大会振込報告フォーム'
                };
            }
            // URLが設定されていない場合はデフォルト設定を返す
            if (!this.scriptUrl || this.scriptUrl === '') {
                console.warn('Google Apps ScriptのURLが設定されていません。デフォルト設定を使用します。');
                return {
                    formTitle: '奉加帳提出報告フォーム',
                    formSubtitle: '〇〇大会振込報告フォーム'
                };
            }
            // まずキャッシュを返す（あれば）
            const cached = this.loadFormSettingsFromCache();
            if (cached) {
                // 裏で最新取得
                fetch(`${this.scriptUrl}?action=getFormSettings`).then(async (res) => {
                    if (!res.ok) return;
                    const result = await res.json();
                    if (result.success && result.formSettings) {
                        this.saveFormSettingsToCache(result.formSettings);
                        // タイトル等はページ側で都度取得しているため即時UI反映は任意
                    }
                }).catch(() => {});
                return cached;
            }

            console.log('フォーム設定を読み込み中...');
            const response = await fetch(`${this.scriptUrl}?action=getFormSettings`);
            if (!response.ok) {
                console.warn(`フォーム設定読み込みでHTTPエラー: ${response.status} ${response.statusText}`);
                throw new Error(`HTTP エラー: ${response.status} ${response.statusText}`);
            }
            const result = await response.json();
            console.log('フォーム設定読み込みレスポンス:', result);
            if (result.success && result.formSettings) {
                this.saveFormSettingsToCache(result.formSettings);
                return result.formSettings;
            } else {
                throw new Error(result.error || 'フォーム設定の読み込みに失敗しました');
            }
            
        } catch (error) {
            console.error('フォーム設定の読み込みエラー:', error);
            console.warn('デフォルト設定を使用します');
            return {
                formTitle: '奉加帳提出報告フォーム',
                formSubtitle: '〇〇大会振込報告フォーム'
            };
        }
    }

    // 設定を保存
    async saveSettings(settings) {
        try {
            // URLが設定されていない場合はエラーを返す
            if (!this.scriptUrl || this.scriptUrl === '') {
                throw new Error('Google Apps ScriptのURLが設定されていません');
            }

            console.log('設定保存開始:', settings);
            
            // 通常のfetchを使用（no-corsモードを削除）
            const response = await fetch(this.scriptUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    action: 'saveSettings',
                    settings: settings
                })
            });

            // レスポンスを確認
            if (!response.ok) {
                throw new Error(`HTTP エラー: ${response.status} ${response.statusText}`);
            }

            const result = await response.json();
            console.log('設定保存結果:', result);
            
            if (result.success) {
                return { 
                    success: true, 
                    message: result.message || '設定を正常に保存しました。' 
                };
            } else {
                throw new Error(result.error || '設定の保存に失敗しました');
            }
            
        } catch (error) {
            console.error('設定保存エラー:', error);
            
            // ネットワークエラーやCORSエラーの場合の詳細メッセージ
            let errorMessage = error.message;
            if (error.name === 'TypeError' && error.message.includes('fetch')) {
                errorMessage = 'Google Apps Scriptとの通信に失敗しました。URLを確認してください。';
            } else if (error.message.includes('CORS')) {
                errorMessage = 'CORS設定に問題があります。Google Apps Scriptの設定を確認してください。';
            }
            
            return { success: false, error: errorMessage };
        }
    }


}

// グローバルインスタンスを作成
window.googleSheetsManager = new GoogleSheetsManager(); 