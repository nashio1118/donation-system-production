// デモ用の設定と機能
class DemoManager {
    constructor() {
        this.isDemo = true;
        this.demoData = [];
        this.initializeDemo();
    }

    // デモ用の初期化
    initializeDemo() {
        console.log('デモモードで動作しています');
        this.setupDemoFeatures();
        
        // デモ版でも負担金の初期表示を確実にする
        if (window.location.pathname.includes('page1.html')) {
            // 複数のタイミングで実行して確実性を向上
            setTimeout(() => {
                console.log('デモ版：負担金初期計算を実行（1回目）');
                this.initializeBurdenFee();
            }, 200);
            
            setTimeout(() => {
                console.log('デモ版：負担金初期計算を実行（2回目）');
                this.initializeBurdenFee();
            }, 500);
            
            // DOMContentLoadedでも実行
            if (document.readyState === 'loading') {
                document.addEventListener('DOMContentLoaded', () => {
                    console.log('デモ版：負担金初期計算を実行（DOMContentLoaded）');
                    this.initializeBurdenFee();
                    this.ensureBurdenFeeSet();
                });
            } else {
                // すでにロード済みでも確実に反映
                this.ensureBurdenFeeSet();
            }
        }
    }

    // デモ用の機能設定
    setupDemoFeatures() {
        // 送信ボタンの動作を変更
        const submitButtons = document.querySelectorAll('button[type="submit"]');
        submitButtons.forEach(button => {
            button.addEventListener('click', (e) => {
                e.preventDefault();
                this.handleDemoSubmit();
            });
        });

        // フォームの送信を防止
        const forms = document.querySelectorAll('form');
        forms.forEach(form => {
            form.addEventListener('submit', (e) => {
                e.preventDefault();
                this.handleDemoSubmit();
            });
        });
    }

    // デモ用の送信処理
    handleDemoSubmit() {
        // 現在のページがpage2.htmlの場合は送信控えページに遷移
        if (window.location.pathname.includes('page2.html')) {
            this.handlePage2DemoSubmit();
        } else {
            // その他の場合はデモ用の成功メッセージを表示
            this.showDemoMessage();
            
            // データをローカルに保存（デモ用）
            this.saveDemoData();
        }
    }

    // 2ページ目のデモ送信処理（最新UI/データ構造に追従）
    handlePage2DemoSubmit() {
        // 基本情報をlocalStorageから取得
        const basicInfo = localStorage.getItem('basicInfo');
        
        if (basicInfo) {
            // デモ用の寄付者データを作成
            const demoSubmissionData = {
                basicInfo: JSON.parse(basicInfo),
                submissionDate: new Date().toISOString(),
                donors: this.createDemoDonors()
            };
            
            // 送信控え用のデータを保存
            localStorage.setItem('submissionData', JSON.stringify(demoSubmissionData));
            
            // 送信控えページに遷移
            window.location.href = 'receipt.html';
        } else {
            // 基本情報がない場合は通常のデモメッセージを表示
            this.showDemoMessage();
        }
    }

    // デモ用の寄付者データを作成（接頭辞一致で全行対応）
    createDemoDonors() {
        const donors = [];
        
        // フォームから寄付者データを収集（インデックスに依存しない）
        const donorRows = document.querySelectorAll('.donor-row');
        donorRows.forEach((row) => {
            const nameInput = row.querySelector('input[name^="donor"]');
            const amountInput = row.querySelector('.donor-amount');
            const giftSelect = row.querySelector('.donor-gift');
            const sheetsSelect = row.querySelector('.donor-sheets');
            const noteInput = row.querySelector('input[name^="note"]');
            
            const name = nameInput?.value?.trim();
            if (!name) return; // 名前がない場合はスキップ
            
            // 金額を取得
            const amount = parseInt(amountInput?.value) || 0;
            
            // 返礼品IDを名前に変換
            const giftIdToName = {
                'towel': 'タオル',
                'sweets_large': 'お菓子大',
                'keychain': 'キーホルダー',
                'sweets_small': 'お菓子小',
                'clearfile': 'クリアファイル',
                'sticker': 'ステッカー',
                't-shirt': 'Tシャツ',
                'hoodie': 'フーディー'
            };
            
            let giftName = '';
            if (giftSelect?.value) {
                giftName = giftIdToName[giftSelect.value] || giftSelect.value;
            }
            
            donors.push({
                name: nameInput?.value?.trim(),
                amount: amount,
                gift: giftName,
                sheets: sheetsSelect?.value || '',
                note: noteInput?.value?.trim() || ''
            });
        });
        
        return donors;
    }

    // デモ版用の負担金初期化
    initializeBurdenFee() {
        console.log('initializeBurdenFee() 開始');
        
        const burdenFeeDisplay = document.getElementById('burdenFeeDisplay');
        const burdenFeeInput = document.getElementById('burdenFeeAmount');
        
        console.log('要素チェック:', {
            burdenFeeDisplay: !!burdenFeeDisplay,
            burdenFeeInput: !!burdenFeeInput,
            displayText: burdenFeeDisplay?.textContent,
            inputValue: burdenFeeInput?.value
        });
        
        if (burdenFeeDisplay) {
            // 現在の値に関係なく1人分の負担金を設定
            const defaultAmount = 40000;
            burdenFeeDisplay.textContent = `¥${defaultAmount.toLocaleString()}`;
            console.log('デモ版：負担金表示を更新:', burdenFeeDisplay.textContent);
        } else {
            console.error('burdenFeeDisplay要素が見つかりません');
        }
        
        if (burdenFeeInput) {
            burdenFeeInput.value = 40000;
            console.log('デモ版：負担金隠しフィールドを更新:', burdenFeeInput.value);
        } else {
            console.error('burdenFeeAmount要素が見つかりません');
        }
        
        // 内訳合計も更新
        this.updateDemoBreakdown();
    }

    // 負担金が0のままにならないよう、監視して強制的に1人分に補正
    ensureBurdenFeeSet() {
        const applyOnce = () => {
            try {
                const input = document.getElementById('burdenFeeAmount');
                const display = document.getElementById('burdenFeeDisplay');
                if (!input || !display) return false;
                const current = parseInt(input.value || '0') || 0;
                if (current === 0) {
                    input.value = 40000;
                    display.textContent = `¥${(40000).toLocaleString()}`;
                    this.updateDemoBreakdown();
                    return true;
                }
                return true;
            } catch (_) { return false; }
        };

        // 即時適用を数回リトライ
        const intervals = [0, 150, 300, 600, 1000];
        intervals.forEach(delay => setTimeout(applyOnce, delay));

        // それでも外れる場合に備えて短時間監視
        try {
            const target = document.getElementById('basicInfoForm') || document.body;
            const observer = new MutationObserver(() => { applyOnce(); });
            observer.observe(target, { childList: true, subtree: true, characterData: true });
            // 数秒で監視解除
            setTimeout(() => { try { observer.disconnect(); } catch (_) {} }, 3000);
        } catch (_) {}
    }
    
    // デモ版用の内訳合計更新
    updateDemoBreakdown() {
        const donationAmount = parseInt(document.getElementById('donationAmount')?.value) || 0;
        const burdenFeeAmount = parseInt(document.getElementById('burdenFeeAmount')?.value) || 0;
        const tshirtAmount = parseInt(document.getElementById('tshirtAmount')?.value) || 0;
        const otherFee = parseInt(document.getElementById('otherFee')?.value) || 0;
        const hasEncouragementFee = document.querySelector('input[name="hasEncouragementFee"]')?.checked || false;
        
        const encouragementAmount = hasEncouragementFee ? 1000 : 0;
        const total = donationAmount + burdenFeeAmount + tshirtAmount + encouragementAmount + otherFee;
        
        const totalElement = document.getElementById('totalBreakdown');
        if (totalElement) {
            totalElement.textContent = `¥${total.toLocaleString()}`;
            console.log('デモ版：内訳合計を更新:', totalElement.textContent);
        }
    }

    // デモ用のメッセージ表示
    showDemoMessage() {
        const message = `
            <div class="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
                <div class="bg-white rounded-lg p-6 max-w-md mx-4">
                    <div class="text-center">
                        <div class="text-4xl mb-4">🎉</div>
                        <h3 class="text-lg font-semibold text-gray-800 mb-2">デモ版 - 送信完了</h3>
                        <p class="text-gray-600 mb-4">
                            実際のシステムでは、データがGoogleスプレッドシートに保存されます。
                        </p>
                        <div class="bg-blue-50 border border-blue-200 rounded-lg p-3 mb-4">
                            <h4 class="font-semibold text-blue-800 mb-2">デモ版の特徴</h4>
                            <ul class="text-sm text-blue-700 space-y-1">
                                <li>• リアルタイム計算機能</li>
                                <li>• 複数人まとめて入力</li>
                                <li>• 自動金額照合</li>
                                <li>• 返礼品自動提案</li>
                                <li>• スマホ最適化</li>
                            </ul>
                        </div>
                        <button onclick="this.parentElement.parentElement.parentElement.remove()" 
                                class="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-md">
                            閉じる
                        </button>
                    </div>
                </div>
            </div>
        `;
        document.body.insertAdjacentHTML('beforeend', message);
    }

    // デモ用のデータ保存
    saveDemoData() {
        const formData = this.collectFormData();
        this.demoData.push({
            timestamp: new Date().toISOString(),
            data: formData
        });
        
        // ローカルストレージに保存（デモ用）
        localStorage.setItem('demoData', JSON.stringify(this.demoData));
        console.log('デモデータを保存しました:', formData);
    }

    // フォームデータの収集
    collectFormData() {
        const form = document.querySelector('form');
        if (!form) return {};

        const formData = new FormData(form);
        const data = {};
        
        for (let [key, value] of formData.entries()) {
            data[key] = value;
        }
        
        return data;
    }

    // デモデータの表示
    showDemoData() {
        const data = localStorage.getItem('demoData');
        if (data) {
            const demoData = JSON.parse(data);
            console.log('保存されたデモデータ:', demoData);
        }
    }
}

// デモ用の計算機クラス
class DemoCalculator {
    constructor() {
        this.giftRules = {
            sticker: { minAmount: 1000, name: 'ステッカー', description: 'チームロゴ入りステッカー' },
            towel: { minAmount: 3000, name: 'タオル', description: 'チームカラータオル' },
            't-shirt': { minAmount: 5000, name: 'Tシャツ', description: 'チームロゴTシャツ' },
            hoodie: { minAmount: 10000, name: 'フーディー', description: 'チームロゴフーディー' },
            clearfile: { minAmount: 0, name: 'クリアファイル', description: 'オリジナルクリアファイル' }
        };
    }

    // 返礼品の提案
    suggestGift(amount) {
        const suggestions = [];
        Object.entries(this.giftRules).forEach(([key, rule]) => {
            if (amount >= rule.minAmount) {
                suggestions.push({
                    key: key,
                    name: rule.name,
                    description: rule.description,
                    minAmount: rule.minAmount
                });
            }
        });
        return suggestions.sort((a, b) => b.minAmount - a.minAmount);
    }

    // 合計金額の計算
    calculateTotal(amounts) {
        return amounts.reduce((sum, amount) => sum + (parseInt(amount) || 0), 0);
    }
}

// デモ用のフォーム管理クラス
class DemoFormManager {
    constructor() {
        this.calculator = new DemoCalculator();
        this.setupFormHandlers();
    }

    // フォームハンドラーの設定
    setupFormHandlers() {
        // 金額入力時のリアルタイム計算
        const amountInputs = document.querySelectorAll('input[type="number"]');
        amountInputs.forEach(input => {
            input.addEventListener('input', () => {
                this.updateCalculations();
            });
        });

        // 寄付者追加ボタン
        const addDonorBtn = document.getElementById('addDonorBtn');
        if (addDonorBtn) {
            addDonorBtn.addEventListener('click', () => {
                this.addDonorRow();
            });
        }
    }

    // 計算の更新
    updateCalculations() {
        // 合計金額の更新
        this.updateTotalAmount();
        
        // 返礼品の更新
        this.updateGiftSuggestions();
    }

    // 合計金額の更新
    updateTotalAmount() {
        const amountInputs = document.querySelectorAll('.donor-amount');
        const amounts = Array.from(amountInputs).map(input => input.value);
        const total = this.calculator.calculateTotal(amounts);
        
        const totalElement = document.getElementById('totalAmount');
        if (totalElement) {
            totalElement.textContent = `¥${total.toLocaleString()}`;
        }
    }

    // 返礼品提案の更新
    updateGiftSuggestions() {
        const amountInputs = document.querySelectorAll('.donor-amount');
        amountInputs.forEach((input, index) => {
            const amount = parseInt(input.value) || 0;
            const suggestions = this.calculator.suggestGift(amount);
            
            const giftSelect = input.closest('.donor-row').querySelector('.donor-gift');
            if (giftSelect) {
                this.updateGiftOptions(giftSelect, suggestions);
            }
        });
    }

    // 返礼品オプションの更新
    updateGiftOptions(select, suggestions) {
        // 既存のオプションをクリア（最初の「返礼品なし」は保持）
        const noGiftOption = select.querySelector('option[value=""]');
        select.innerHTML = '';
        if (noGiftOption) {
            select.appendChild(noGiftOption);
        }

        // 新しい返礼品オプションを追加
        suggestions.forEach(suggestion => {
            const option = document.createElement('option');
            option.value = suggestion.key;
            option.textContent = `${suggestion.name}（¥${suggestion.minAmount.toLocaleString()}以上）`;
            select.appendChild(option);
        });
    }

    // 寄付者行の追加
    addDonorRow() {
        const container = document.getElementById('donorsContainer');
        if (!container) return;

        const donorCount = container.children.length + 1;
        const newRow = this.createDonorRow(donorCount);
        container.appendChild(newRow);
        
        // 計算を更新
        this.updateCalculations();
    }

    // 寄付者行の作成
    createDonorRow(index) {
        const row = document.createElement('div');
        row.className = 'donor-row grid grid-cols-1 md:grid-cols-2 gap-4 mb-4 p-4 border border-gray-200 rounded-lg';
        row.innerHTML = `
            <div class="flex justify-between items-center mb-3 col-span-2">
                <h3 class="text-lg font-medium text-gray-800">寄付者 ${index}</h3>
                <button type="button" onclick="this.closest('.donor-row').remove(); updateCalculations();"
                        class="bg-red-500 hover:bg-red-600 text-white px-3 py-2 rounded-md text-sm">
                    削除
                </button>
            </div>
            <div>
                <label class="block text-sm font-medium text-gray-700 mb-1">御芳名 *</label>
                <input type="text" name="donor${index}" required
                       class="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                       placeholder="寄付者名">
            </div>
            <div>
                <label class="block text-sm font-medium text-gray-700 mb-1">金額 *</label>
                <input type="number" name="amount${index}" min="100" step="100" required
                       class="donor-amount w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                       placeholder="0">
            </div>
            <div>
                <label class="block text-sm font-medium text-gray-700 mb-1">返礼品</label>
                <select name="gift${index}" class="donor-gift w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500">
                    <option value="">返礼品なし</option>
                </select>
            </div>
            <div class="donor-sheets-container hidden">
                <label class="block text-sm font-medium text-gray-700 mb-1">クリアファイル枚数</label>
                <select name="sheets${index}" class="donor-sheets w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500">
                    <option value="">枚数を選択</option>
                </select>
            </div>
            <div class="col-span-2">
                <label class="block text-sm font-medium text-gray-700 mb-1">その他特筆事項</label>
                <input type="text" name="note${index}"
                       class="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                       placeholder="特になし">
            </div>
        `;

        // イベントリスナーを設定
        this.setupDonorRowEvents(row, index);
        return row;
    }

    // 寄付者行のイベントリスナーを設定
    setupDonorRowEvents(row, index) {
        const amountInput = row.querySelector('.donor-amount');
        const giftSelect = row.querySelector('.donor-gift');
        const sheetsContainer = row.querySelector('.donor-sheets-container');
        const sheetsSelect = row.querySelector('.donor-sheets');

        // 金額入力時の処理
        amountInput.addEventListener('input', () => {
            this.updateGiftSuggestions(amountInput, giftSelect);
        });

        // 返礼品変更時の処理
        giftSelect.addEventListener('change', () => {
            this.handleGiftChange(giftSelect, sheetsContainer, sheetsSelect);
        });

        // クリアファイル枚数オプションを生成
        this.generateClearfileOptions(sheetsSelect);
    }

    // 返礼品提案の更新
    updateGiftSuggestions(amountInput, giftSelect) {
        const amount = parseInt(amountInput.value) || 0;
        const suggestions = this.calculator.suggestGift(amount);
        this.updateGiftOptions(giftSelect, suggestions);
    }

    // 返礼品変更時の処理
    handleGiftChange(giftSelect, sheetsContainer, sheetsSelect) {
        if (giftSelect.value === 'clearfile') {
            sheetsContainer.classList.remove('hidden');
        } else {
            sheetsContainer.classList.add('hidden');
        }
    }

    // クリアファイル枚数オプションを生成
    generateClearfileOptions(select) {
        for (let i = 1; i <= 50; i++) {
            const option = document.createElement('option');
            option.value = i;
            option.textContent = `${i}枚`;
            select.appendChild(option);
        }
    }
}

// デモ用の初期化（スクリプトがDOMContentLoaded後に読み込まれても必ず実行）
(function initDemoScripts() {
    const init = () => {
        window.demoManager = new DemoManager();
        window.demoFormManager = new DemoFormManager();

        // グローバル関数として公開
        window.updateCalculations = function() {
            if (window.demoFormManager) {
                window.demoFormManager.updateCalculations();
            }
        };

        // 安全策: page1で負担金が0のままなら再初期化（遅延チェック）
        if (window.location.pathname.includes('page1.html')) {
            setTimeout(() => {
                try {
                    const feeInput = document.getElementById('burdenFeeAmount');
                    if (!feeInput || parseInt(feeInput.value || '0') === 0) {
                        if (window.demoManager && typeof window.demoManager.initializeBurdenFee === 'function') {
                            window.demoManager.initializeBurdenFee();
                        }
                    }
                } catch (_) {}
            }, 800);
        }
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();