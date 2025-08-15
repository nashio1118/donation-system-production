// 1ページ目フォーム管理クラス
class BasicInfoForm {
    constructor() {
        this.form = document.getElementById('basicInfoForm');
        // デモフラグの引き継ぎ（page1.html?demo=1 で来た場合、次ページ以降でも維持）
        this.ensureDemoFlag();
        this.initializeEventListeners();
        this.setDefaultDate();
        this.loadFormSettings(); // フォーム設定を読み込み
        this.loadSavedData(); // 保存されたデータを復元
        
        // 初期表示で負担金を計算（非同期処理の完了を待たずに最低限の表示）
        setTimeout(() => {
            console.log('初期負担金計算を実行');
            this.updateBurdenFee();
            this.updateBreakdown();
        }, 100);
    }

    // URLパラメータに demo=1 があれば sessionStorage に保存
    ensureDemoFlag() {
        try {
            const params = new URLSearchParams(window.location.search || '');
            if (params.get('demo') === '1') {
                sessionStorage.setItem('isDemo', '1');
            }
        } catch (e) {
            // 何もしない（セッションストレージ未対応環境の安全策）
        }
    }

    initializeEventListeners() {
        this.form.addEventListener('submit', (e) => this.handleSubmit(e));
        
        // 兄弟機能のイベントリスナー
        const hasSiblingCheckbox = document.getElementById('hasSibling');
        const siblingCountSelect = document.getElementById('siblingCount');
        
        if (hasSiblingCheckbox) {
            hasSiblingCheckbox.addEventListener('change', () => this.toggleSiblingSection());
        }
        
        if (siblingCountSelect) {
            siblingCountSelect.addEventListener('change', () => this.updateSiblingNames());
        }
        
        // 内訳計算のためのイベントリスナー
        const inputs = ['donationAmount', 'tshirtAmount', 'hasEncouragementFee', 'otherFee', 'transferAmount', 'directTransferAmount'];
        inputs.forEach(inputName => {
            const elements = this.form.querySelectorAll(`[name="${inputName}"]`);
            elements.forEach(element => {
                element.addEventListener('change', () => this.updateBreakdown());
                element.addEventListener('input', () => this.updateBreakdown());
            });
        });
    }

    // デフォルト日付を設定（今日）
    setDefaultDate() {
        const today = new Date().toISOString().split('T')[0];
        const dateInput = document.getElementById('transferDate');
        if (dateInput && !dateInput.value) { // 既に値がある場合は上書きしない
            dateInput.value = today;
        }
    }

    // フォーム設定を読み込み
    async loadFormSettings() {
        try {
            if (window.googleSheetsManager) {
                // フォーム設定を取得
                const formSettings = await window.googleSheetsManager.getFormSettings();
                if (formSettings) {
                    this.updateFormTitles(formSettings);
                }
                
                // 負担金設定を取得
                await window.googleSheetsManager.loadSettings();
                if (window.googleSheetsManager.settings) {
                    this.updateBurdenFeeLabel();
                    this.toggleEncouragementVisibility();
                    // 設定読込直後に負担金を計算（兄弟がいなくても1人分）
                    this.updateBurdenFee();
                    this.updateBreakdown();
                }
                // 設定が未取得/取得失敗でも、少なくとも1人分で初期計算を実行
                if (!window.googleSheetsManager.settings) {
                    this.updateBurdenFee();
                    this.updateBreakdown();
                }
            }
        } catch (error) {
            console.error('フォーム設定の読み込みエラー:', error);
            // エラーの場合はデフォルト値のまま
        }
    }

    // 負担金ラベルを更新
    updateBurdenFeeLabel() {
        // 設定から負担金額を取得
        let burdenFeeAmount = 40000; // デフォルト値
        try {
            const s = window.googleSheetsManager && window.googleSheetsManager.settings;
            const fee = s && s.burdenSettings && s.burdenSettings.burden_fee;
            if (fee && typeof fee.amount === 'number') {
                burdenFeeAmount = fee.amount;
            }
        } catch (_) {}

        // ラベルを更新（激励会協力金の呼称も統一）
        const labelElement = document.getElementById('burdenFeeLabel');
        if (labelElement) {
            labelElement.textContent = `負担金（1人あたり${burdenFeeAmount.toLocaleString()}円）`;
        }
        
        // 負担金計算で使用する単価を保存
        this.burdenFeePerPerson = burdenFeeAmount;
    }

    // 激励会協力金の表示/非表示を切り替え（設定にencouragement_feeが無ければ非表示）
    toggleEncouragementVisibility() {
        try {
            const settings = window.googleSheetsManager && window.googleSheetsManager.settings;
            const encouragement = settings && settings.burdenSettings && settings.burdenSettings.encouragement_fee;
            const label = document.getElementById('encouragementFeeLabel');
            if (!label) return;
            const section = label.closest('div');
            if (!encouragement) {
                // 非表示にして入力不可にする
                if (section) section.style.display = 'none';
                const checkbox = document.querySelector('input[name="hasEncouragementFee"]');
                if (checkbox) { checkbox.checked = false; checkbox.disabled = true; }
            } else {
                if (section) section.style.display = '';
                const checkbox = document.querySelector('input[name="hasEncouragementFee"]');
                if (checkbox) { checkbox.disabled = false; }
            }
        } catch (e) {
            console.error('激励会協力金表示制御エラー:', e);
        }
    }

    // フォームタイトルを更新
    updateFormTitles(settings) {
        const titleElement = document.getElementById('formTitle');
        const subtitleElement = document.getElementById('formSubtitle');
        
        if (titleElement && settings.formTitle) {
            titleElement.textContent = settings.formTitle;
            // ブラウザタイトルも更新
            try {
                document.title = settings.formTitle;
            } catch (_) {}
        }
        
        if (subtitleElement && settings.formSubtitle) {
            subtitleElement.textContent = settings.formSubtitle;
        }
    }

    // 兄弟セクションの表示/非表示を切り替え
    toggleSiblingSection() {
        const hasSibling = document.getElementById('hasSibling').checked;
        const siblingCountContainer = document.getElementById('siblingCountContainer');
        const siblingNamesContainer = document.getElementById('siblingNamesContainer');
        
        if (hasSibling) {
            siblingCountContainer.classList.remove('hidden');
        } else {
            siblingCountContainer.classList.add('hidden');
            siblingNamesContainer.classList.add('hidden');
            
            // 兄弟関連のフィールドをリセット
            document.getElementById('siblingCount').value = '';
            document.getElementById('siblingNamesFields').innerHTML = '';
        }
        
        // 負担金を再計算
        this.updateBurdenFee();
        this.updateBreakdown();
    }

    // 兄弟の名前入力欄を更新
    updateSiblingNames() {
        const siblingCount = parseInt(document.getElementById('siblingCount').value) || 0;
        const siblingNamesContainer = document.getElementById('siblingNamesContainer');
        const siblingNamesFields = document.getElementById('siblingNamesFields');
        
        // フィールドをクリア
        siblingNamesFields.innerHTML = '';
        
        if (siblingCount > 0) {
            siblingNamesContainer.classList.remove('hidden');
            
            // 兄弟の数だけ入力欄を作成
            for (let i = 1; i <= siblingCount; i++) {
                const fieldDiv = document.createElement('div');
                fieldDiv.innerHTML = `
                    <label for="sibling${i}Name" class="block text-sm font-medium text-gray-700 mb-1">兄弟${i}の名前</label>
                    <input type="text" id="sibling${i}Name" name="sibling${i}Name"
                           class="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                           placeholder="例：田中花子">
                `;
                siblingNamesFields.appendChild(fieldDiv);
            }
        } else {
            siblingNamesContainer.classList.add('hidden');
        }
        
        // 負担金を再計算
        this.updateBurdenFee();
        this.updateBreakdown();
    }

    // 負担金を更新
    updateBurdenFee() {
        const hasSibling = document.getElementById('hasSibling')?.checked || false;
        const siblingCount = parseInt(document.getElementById('siblingCount')?.value) || 0;
        
        // 本人を含めた総人数を計算（兄弟がいなくても最低1人分）
        const totalPersons = hasSibling ? (siblingCount + 1) : 1;
        
        // 負担金単価を取得（設定から、またはデフォルト値）
        const burdenFeePerPerson = this.burdenFeePerPerson || 40000;
        const burdenFeeAmount = totalPersons * burdenFeePerPerson;
        
        console.log('負担金計算:', { hasSibling, siblingCount, totalPersons, burdenFeePerPerson, burdenFeeAmount });
        
        // 表示を更新
        const burdenFeeDisplay = document.getElementById('burdenFeeDisplay');
        const burdenFeeInput = document.getElementById('burdenFeeAmount');
        
        if (burdenFeeDisplay) {
            burdenFeeDisplay.textContent = `¥${burdenFeeAmount.toLocaleString()}`;
            console.log('負担金表示を更新:', burdenFeeDisplay.textContent);
        } else {
            console.error('burdenFeeDisplay要素が見つかりません');
        }
        
        if (burdenFeeInput) {
            burdenFeeInput.value = burdenFeeAmount;
            console.log('負担金隠しフィールドを更新:', burdenFeeInput.value);
        } else {
            console.error('burdenFeeAmount要素が見つかりません');
        }
    }

    // 保存されたデータを復元
    loadSavedData() {
        const savedData = localStorage.getItem('basicInfo');
        if (savedData) {
            try {
                const data = JSON.parse(savedData);
                this.restoreFormData(data);
            } catch (error) {
                console.error('保存データの復元に失敗しました:', error);
            }
        }
    }

    // フォームデータを復元
    restoreFormData(data) {
        // 基本情報の復元
        if (data.basicInfo) {
            const { grade, name, hasSibling, siblingCount, transferDate } = data.basicInfo;
            
            if (grade) document.getElementById('grade').value = grade;
            if (name) document.getElementById('name').value = name;
            if (transferDate) document.getElementById('transferDate').value = transferDate;
            
            // 兄弟チェックボックス
            const siblingCheckbox = document.getElementById('hasSibling');
            if (siblingCheckbox) {
                siblingCheckbox.checked = hasSibling || false;
                this.toggleSiblingSection();
                
                // 兄弟の人数を復元
                if (hasSibling && siblingCount) {
                    document.getElementById('siblingCount').value = siblingCount;
                    this.updateSiblingNames();
                    
                    // 兄弟の名前を復元
                    for (let i = 1; i <= parseInt(siblingCount); i++) {
                        const siblingNameField = document.getElementById(`sibling${i}Name`);
                        if (siblingNameField && data.basicInfo[`sibling${i}Name`]) {
                            siblingNameField.value = data.basicInfo[`sibling${i}Name`];
                        }
                    }
                }
            }
        }

        // 振込金額の復元
        if (data.breakdownDetails) {
            const {
                transferAmount,
                directTransferAmount,
                donationAmount,
                burdenFee,
                hasEncouragementFee,
                tshirtAmount,
                otherFee
            } = data.breakdownDetails;

            // 振込金額（totalTransferAmountから逆算）
            if (transferAmount !== undefined) {
                const totalTransfer = parseInt(transferAmount) || 0;
                const directTransfer = parseInt(directTransferAmount) || 0;
                const normalTransfer = totalTransfer - directTransfer;
                
                document.getElementById('transferAmount').value = normalTransfer > 0 ? normalTransfer : '';
            }
            
            if (directTransferAmount !== undefined) {
                const directValue = parseInt(directTransferAmount) || 0;
                document.getElementById('directTransferAmount').value = directValue > 0 ? directValue : '';
            }

            // 内訳の復元
            if (donationAmount !== undefined) {
                document.getElementById('donationAmount').value = donationAmount || '';
            }
            
            if (tshirtAmount !== undefined) {
                document.getElementById('tshirtAmount').value = tshirtAmount || '';
            }
            
            if (otherFee !== undefined) {
                document.getElementById('otherFee').value = otherFee || '';
            }

            // 負担金は兄弟機能で自動計算されるため、復元時は計算のみ実行
            this.updateBurdenFee();

            // 奨励会負担金チェックボックス
            const encouragementCheckbox = document.querySelector('input[name="hasEncouragementFee"]');
            if (encouragementCheckbox) {
                encouragementCheckbox.checked = hasEncouragementFee || false;
            }
        }

        // データ復元後に内訳を更新
        this.updateBreakdown();
    }

    // 内訳合計を更新
    updateBreakdown() {
        const donationAmount = parseInt(document.getElementById('donationAmount').value) || 0;
        const burdenFeeAmount = parseInt(document.getElementById('burdenFeeAmount').value) || 0;
        const tshirtAmount = parseInt(document.getElementById('tshirtAmount').value) || 0;
        const otherFee = parseInt(document.getElementById('otherFee').value) || 0;
        const hasEncouragementFee = document.querySelector('input[name="hasEncouragementFee"]').checked;

        // 各項目の金額を計算
        const encouragementAmount = hasEncouragementFee ? 1000 : 0; // 激励会協力金（設定で金額管理）

        const total = donationAmount + burdenFeeAmount + tshirtAmount + encouragementAmount + otherFee;

        // 合計を表示
        document.getElementById('totalBreakdown').textContent = `¥${total.toLocaleString()}`;

        // 振込金額と照合
        this.validateTransferAmount(total);
    }

    // 振込金額との照合
    validateTransferAmount(breakdownTotal) {
        const transferAmount = parseInt(document.getElementById('transferAmount').value) || 0;
        const directTransferAmount = parseInt(document.getElementById('directTransferAmount').value) || 0;
        const totalTransferAmount = transferAmount + directTransferAmount;
        
        if (totalTransferAmount > 0 && breakdownTotal !== totalTransferAmount) {
            // 警告を表示
            const warningDiv = document.getElementById('transfer-warning') || this.createTransferWarningDiv();
            warningDiv.innerHTML = `
                <div class="bg-yellow-50 border-l-4 border-yellow-400 p-4 mb-4">
                    <div class="flex">
                        <div class="flex-shrink-0">
                            <svg class="h-5 w-5 text-yellow-400" viewBox="0 0 20 20" fill="currentColor">
                                <path fill-rule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clip-rule="evenodd" />
                            </svg>
                        </div>
                        <div class="ml-3">
                            <p class="text-sm text-yellow-700">
                                <strong>金額の不一致:</strong> 内訳合計（¥${breakdownTotal.toLocaleString()}）と振込金額合計（¥${totalTransferAmount.toLocaleString()}）が一致しません。
                            </p>
                            <p class="text-xs text-yellow-600 mt-1">
                                振込金額: ¥${transferAmount.toLocaleString()} + 直接振込金額: ¥${directTransferAmount.toLocaleString()} = ¥${totalTransferAmount.toLocaleString()}
                            </p>
                        </div>
                    </div>
                </div>
            `;
            warningDiv.style.display = 'block';
        } else {
            const warningDiv = document.getElementById('transfer-warning');
            if (warningDiv) {
                warningDiv.style.display = 'none';
            }
        }
    }

    // 振込金額警告表示用のdivを作成
    createTransferWarningDiv() {
        const warningDiv = document.createElement('div');
        warningDiv.id = 'transfer-warning';
        const form = document.getElementById('basicInfoForm');
        form.insertBefore(warningDiv, form.firstChild);
        return warningDiv;
    }

    // フォーム送信処理
    async handleSubmit(e) {
        e.preventDefault();
        
        if (!this.validateForm()) {
            alert('必須項目を入力してください。');
            return;
        }

        if (!this.validateAmounts()) {
            alert('金額の照合に失敗しました。内訳合計と振込金額合計（振込金額＋直接振込金額）を確認してください。');
            return;
        }

        const formData = this.collectFormData();
        
        try {
            this.showLoading(true);
            
            // データをローカルストレージに保存
            localStorage.setItem('basicInfo', JSON.stringify(formData));
            
            // 2ページ目に遷移
            window.location.href = 'page2.html';
            
        } catch (error) {
            console.error('処理エラー:', error);
            alert('処理中にエラーが発生しました。もう一度お試しください。');
        } finally {
            this.showLoading(false);
        }
    }

    // フォームバリデーション
    validateForm() {
        const requiredFields = this.form.querySelectorAll('[required]');
        for (let field of requiredFields) {
            if (!field.value.trim()) {
                return false;
            }
        }
        return true;
    }

    // 金額照合の検証
    validateAmounts() {
        const transferAmount = parseInt(document.getElementById('transferAmount').value) || 0;
        const directTransferAmount = parseInt(document.getElementById('directTransferAmount').value) || 0;
        const totalTransferAmount = transferAmount + directTransferAmount;
        
        const donationAmount = parseInt(document.getElementById('donationAmount').value) || 0;
        const burdenFeeAmount = parseInt(document.getElementById('burdenFeeAmount').value) || 0;
        const tshirtAmount = parseInt(document.getElementById('tshirtAmount').value) || 0;
        const otherFee = parseInt(document.getElementById('otherFee').value) || 0;
        const hasEncouragementFee = document.querySelector('input[name="hasEncouragementFee"]').checked;

        const encouragementAmount = hasEncouragementFee ? 1000 : 0;

        const breakdownTotal = donationAmount + burdenFeeAmount + tshirtAmount + encouragementAmount + otherFee;

        // 振込金額合計が0の場合は照合しない
        if (totalTransferAmount === 0) {
            return true;
        }

        return breakdownTotal === totalTransferAmount;
    }

    // フォームデータを収集
    collectFormData() {
        const formData = new FormData(this.form);
        
        // 振込金額の取得
        const transferAmount = parseInt(formData.get('transferAmount')) || 0; // 振込金額（直接振込を除く）
        const directTransferAmount = parseInt(formData.get('directTransferAmount')) || 0; // 直接振込金額
        const totalTransferAmount = transferAmount + directTransferAmount; // 合計（照合用）
        
        // 負担金の金額を取得
        const burdenFeeAmount = parseInt(formData.get('burdenFee')) || 0;
        
        // 奨励会負担金のチェック状態を取得
        const hasEncouragementFee = formData.get('hasEncouragementFee') === 'on';
        
        // 兄弟情報を収集
        const hasSibling = formData.get('hasSibling') === 'on';
        const siblingCount = parseInt(formData.get('siblingCount')) || 0;
        
        // GASが期待する構造に合わせてデータを構築
        const data = {
            basicInfo: {
                grade: formData.get('grade'),
                name: formData.get('name'),
                hasSibling: hasSibling,
                siblingCount: siblingCount,
                transferDate: formData.get('transferDate')
            },
            breakdownDetails: {
                transferAmount: transferAmount.toString(), // 振込金額（直接振込を除く）
                directTransferAmount: directTransferAmount.toString(), // 直接振込金額
                donationAmount: parseInt(formData.get('donationAmount')) || 0,
                burdenFee: burdenFeeAmount, // 計算された負担金額
                hasEncouragementFee: hasEncouragementFee, // boolean値として送信
                tshirtAmount: parseInt(formData.get('tshirtAmount')) || 0,
                otherFee: parseInt(formData.get('otherFee')) || 0
            },
            submissionDate: new Date().toISOString()
        };

        // 兄弟の名前を追加
        for (let i = 1; i <= siblingCount; i++) {
            const siblingName = formData.get(`sibling${i}Name`);
            if (siblingName) {
                data.basicInfo[`sibling${i}Name`] = siblingName;
            }
        }

        // 内訳の詳細計算（表示用）
        const encouragementAmount = hasEncouragementFee ? 1000 : 0;

        data.breakdownDetails.burdenFeeAmount = burdenFeeAmount;
        data.breakdownDetails.encouragementAmount = encouragementAmount;
        data.breakdownDetails.total = data.breakdownDetails.donationAmount + burdenFeeAmount + data.breakdownDetails.tshirtAmount + encouragementAmount + data.breakdownDetails.otherFee;

        return data;
    }

    // ローディング表示
    showLoading(show) {
        const loading = document.getElementById('loading');
        loading.classList.toggle('hidden', !show);
    }
}

// フォームを初期化
document.addEventListener('DOMContentLoaded', () => {
    new BasicInfoForm();
}); 