// フォーム管理クラス
class DonationForm {
    constructor() {
        this.donorCount = 0;
        this.donorsContainer = document.getElementById('donorsContainer');
        this.addDonorBtn = document.getElementById('addDonorBtn');
        this.form = document.getElementById('donationForm');
        this.backToPage1Btn = document.getElementById('backToPage1Btn');
        this.confirmModal = document.getElementById('confirmModal');
        this.confirmCancelBtn = document.getElementById('confirmCancelBtn');
        this.confirmSubmitBtn = document.getElementById('confirmSubmitBtn');
        this.basicInfo = null;
        // デモフラグを維持（page1のdemo=1から遷移時にURLから消えてもデモ扱いを続行）
        this.ensureDemoFlag();
        
        this.initializeEventListeners();
        this.loadBasicInfo();
        this.loadPage2Data(); // 2ページ目のデータを復元
        this.addInitialDonors(); // デフォルトで5名追加（復元データがない場合のみ）
        this.loadFormSettings(); // フォーム設定を読み込み
    }

    ensureDemoFlag() {
        try {
            const params = new URLSearchParams(window.location.search || '');
            if (params.get('demo') === '1') {
                sessionStorage.setItem('isDemo', '1');
            }
        } catch (e) {}
    }

    initializeEventListeners() {
        this.addDonorBtn.addEventListener('click', () => this.addDonorRow());
        this.form.addEventListener('submit', (e) => this.handleSubmit(e));
        
        // 1ページ目に戻るボタン
        if (this.backToPage1Btn) {
            this.backToPage1Btn.addEventListener('click', () => this.backToPage1());
        }
        
        // 確認モーダルのボタン
        if (this.confirmCancelBtn) {
            this.confirmCancelBtn.addEventListener('click', () => this.hideConfirmModal());
        }
        if (this.confirmSubmitBtn) {
            this.confirmSubmitBtn.addEventListener('click', () => this.confirmAndSubmit());
        }
    }

    // 1ページ目のデータを読み込み
    loadBasicInfo() {
        const basicInfoData = localStorage.getItem('basicInfo');
        if (basicInfoData) {
            this.basicInfo = JSON.parse(basicInfoData);
            this.displayBasicInfo();
        }
    }

    // 基本情報を表示
    displayBasicInfo() {
        if (!this.basicInfo) return;

        // ヘッダーに基本情報を追加
        const header = document.querySelector('header');
        const infoDiv = document.createElement('div');
        infoDiv.className = 'mt-4 p-4 bg-blue-50 rounded-lg';
        infoDiv.innerHTML = `
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4 text-sm">
                <div><strong>名前:</strong> ${this.basicInfo.basicInfo.name}</div>
                <div><strong>学年:</strong> ${this.basicInfo.basicInfo.grade}</div>
                <div><strong>振込日:</strong> ${this.basicInfo.basicInfo.transferDate}</div>
            </div>
            <div class="mt-2 p-2 bg-white rounded border">
                <strong>1ページ目の寄付金:</strong> ¥${this.basicInfo.breakdownDetails.donationAmount.toLocaleString()}
            </div>
        `;
        header.appendChild(infoDiv);
    }

    // 寄付者行を追加
	addDonorRow() {
        this.donorCount++;
        const donorRow = this.createDonorRow(this.donorCount);
        this.donorsContainer.appendChild(donorRow);
		// 初期描画のバッチ中は合計計算を遅延
		if (!this.isBatchAppending) {
			this.updateTotals();
		}
    }

    // 寄付者行を削除
    removeDonorRow(button) {
        const donorRow = button.closest('.donor-row');
        donorRow.remove();
        // 削除後に連番へ再採番
        this.renumberDonorRows();
        this.updateTotals();
    }

    // 寄付者行を1..Nで再採番（見出しとname属性を更新）
    renumberDonorRows() {
        const rows = this.donorsContainer.querySelectorAll('.donor-row');
        let index = 1;
        rows.forEach((row) => {
            // 見出しの更新
            const header = row.querySelector('h3');
            if (header) header.textContent = `寄付者 ${index}`;

            // name属性の更新ユーティリティ
            const rename = (selector, baseName) => {
                const el = row.querySelector(selector);
                if (el) el.setAttribute('name', `${baseName}_${index}`);
            };

            rename('input[name^="donorName_"]', 'donorName');
            rename('select[name^="donorAmountType_"]', 'donorAmountType');
            rename('input[name^="donorAmountCustom_"]', 'donorAmountCustom');
            rename('select[name^="donorGift_"]', 'donorGift');
            rename('select[name^="donorSheets_"]', 'donorSheets');
            rename('input[name^="donorNote_"]', 'donorNote');
            rename('input[name^="donorDirectTransfer_"]', 'donorDirectTransfer');
            rename('input[name^="donorDirectDate_"]', 'donorDirectDate');

            index++;
        });

        // 内部カウンタも同期
        this.donorCount = rows.length;
    }

	// デフォルトで5名追加（復元データがない場合のみ）
    addInitialDonors() {
        // 既に寄付者行がある場合（復元された場合）はスキップ
        if (this.donorsContainer.children.length > 0) {
            return;
        }
		// 初期追加中の連続再計算を避ける
		this.isBatchAppending = true;
		for (let i = 0; i < 5; i++) {
			this.addDonorRow();
		}
		this.isBatchAppending = false;
		this.updateTotals();
    }

    // 寄付者行のHTMLを生成
    createDonorRow(index) {
        const row = document.createElement('div');
        row.className = 'donor-row bg-gray-50 rounded-lg p-4 mb-4 border';
        row.innerHTML = `
            <div class="flex justify-between items-center mb-3">
                <h3 class="text-lg font-medium text-gray-800">寄付者 ${index}</h3>
                <button type="button" class="remove-donor-btn text-red-500 hover:text-red-700 text-sm font-medium">
                    削除
                </button>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">御芳名 *</label>
                    <input type="text" name="donorName_${index}" required
                           class="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500">
                    <div class="mt-2">
                        <label class="inline-flex items-center">
                            <input type="checkbox" name="donorDirectTransfer_${index}" class="donor-direct-transfer mr-2">
                            <span class="text-sm font-medium text-gray-700">直接振込の場合はチェック</span>
                        </label>
                    </div>
                    <div class="donor-direct-date hidden mt-2">
                        <label class="block text-sm font-medium text-gray-700 mb-1">直接振込日</label>
                        <input type="date" name="donorDirectDate_${index}"
                               class="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500">
                    </div>
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">金額 *</label>
                    <select name="donorAmountType_${index}" class="donor-amount-type w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500">
                        <option value="">金額を選択</option>
                        <option value="10000">¥10,000</option>
                        <option value="5000">¥5,000</option>
                        <option value="other">その他</option>
                    </select>
                </div>
                <div class="donor-amount-input hidden">
                    <label class="block text-sm font-medium text-gray-700 mb-1">金額入力</label>
                    <input type="number" name="donorAmountCustom_${index}" min="0" step="100"
                           class="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                           placeholder="金額を入力">
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">返礼品</label>
                    <select name="donorGift_${index}" class="donor-gift w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500">
                        <option value="">返礼品なし</option>
                    </select>
                </div>
                <div class="donor-sheets-container hidden">
                    <label class="block text-sm font-medium text-gray-700 mb-1">クリアファイル枚数</label>
                    <select name="donorSheets_${index}" class="donor-sheets w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500">
                        <option value="">枚数を選択</option>
                    </select>
                </div>
                <div>
                    <label class="block text-sm font-medium text-gray-700 mb-1">その他特筆事項</label>
                    <input type="text" name="donorNote_${index}"
                           class="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
                           placeholder="特になし">
                </div>
            </div>
        `;

        // イベントリスナーを設定
        this.setupDonorRowEvents(row, index);
        return row;
    }

    // 寄付者行のイベントリスナーを設定
    setupDonorRowEvents(row, index) {
        const amountTypeSelect = row.querySelector('.donor-amount-type');
        const amountInput = row.querySelector('.donor-amount-input');
        const amountCustomInput = row.querySelector('.donor-amount-input input');
        const giftSelect = row.querySelector('.donor-gift');
        const sheetsContainer = row.querySelector('.donor-sheets-container');
        const sheetsSelect = row.querySelector('.donor-sheets');
        const directTransferCheckbox = row.querySelector('.donor-direct-transfer');
        const directDateContainer = row.querySelector('.donor-direct-date');

        // 金額タイプ変更時の処理
        amountTypeSelect.addEventListener('change', () => {
            this.handleAmountTypeChange(amountTypeSelect, amountInput, giftSelect, sheetsContainer, sheetsSelect);
        });

        // 金額入力変更時の処理
        if (amountCustomInput) {
            amountCustomInput.addEventListener('input', () => {
                this.handleAmountInputChange(amountCustomInput, giftSelect, sheetsContainer, sheetsSelect);
            });
        }

        // 返礼品変更時の処理
        giftSelect.addEventListener('change', () => {
            this.handleGiftChange(giftSelect, sheetsContainer, sheetsSelect);
        });

        // 直接振込チェックボックスの処理
        if (directTransferCheckbox && directDateContainer) {
            directTransferCheckbox.addEventListener('change', () => {
                if (directTransferCheckbox.checked) {
                    directDateContainer.classList.remove('hidden');
                } else {
                    directDateContainer.classList.add('hidden');
                    // チェックを外した時は日付もクリア
                    const dateInput = directDateContainer.querySelector('input[type="date"]');
                    if (dateInput) {
                        dateInput.value = '';
                    }
                }
            });
        }

        // 削除ボタンの処理
        const removeBtn = row.querySelector('.remove-donor-btn');
        removeBtn.addEventListener('click', () => this.removeDonorRow(removeBtn));
    }

    // 金額タイプ変更時の処理
    handleAmountTypeChange(amountTypeSelect, amountInput, giftSelect, sheetsContainer, sheetsSelect) {
        const selectedValue = amountTypeSelect.value;
        
        if (selectedValue === 'other') {
            amountInput.classList.remove('hidden');
            giftSelect.innerHTML = '<option value="">返礼品なし</option>';
            sheetsContainer.classList.add('hidden');
        } else {
            amountInput.classList.add('hidden');
            
            // 金額を取得（シート側追加にも対応）
            const amount = parseInt(selectedValue) || 0;
            
            // 返礼品オプションを更新
            this.updateGiftOptions(giftSelect, amount);
            sheetsContainer.classList.add('hidden'); // 金額が決まったらクリアファイルコンテナを非表示
        }
        
        this.updateTotals();
    }

    // 金額入力変更時の処理
    handleAmountInputChange(amountCustomInput, giftSelect, sheetsContainer, sheetsSelect) {
        const customAmount = parseInt(amountCustomInput.value) || 0;
        
        // 返礼品オプションを更新
        this.updateGiftOptions(giftSelect, customAmount);
        sheetsContainer.classList.add('hidden'); // 金額が決まったらクリアファイルコンテナを非表示
        this.updateTotals();
    }

    // 返礼品オプションを更新（動的生成）
    updateGiftOptions(giftSelect, amount) {
		const calculator = window.donationCalculator;
		if (calculator && calculator.generateGiftOptions) {
			// できるだけ再生成を避ける（同一HTMLならスキップ）
			const next = calculator.generateGiftOptions(amount);
			if (giftSelect.__lastOptionsHtml !== next) {
				giftSelect.innerHTML = next;
				giftSelect.__lastOptionsHtml = next;
			}
		} else {
            // フォールバック: 従来の条件分岐
			const next = '<option value="">返礼品なし</option>' + `
                <option value="タオル">タオル</option>
                <option value="お菓子大">お菓子大</option>
                <option value="キーホルダー">キーホルダー</option>
                <option value="お菓子小">お菓子小</option>
                <option value="クリアファイル">クリアファイル</option>
			`;
			if (giftSelect.__lastOptionsHtml !== next) {
				giftSelect.innerHTML = next;
				giftSelect.__lastOptionsHtml = next;
			}
        }
    }

    // 返礼品変更時の処理
    handleGiftChange(giftSelect, sheetsContainer, sheetsSelect) {
        const selectedGift = giftSelect.value;
        
        // クリアファイルが選択された場合（値は返礼品名を使用）
        if (selectedGift === 'クリアファイル') {
            sheetsContainer.classList.remove('hidden');
            
            // 枚数オプションを更新
            const calculator = window.donationCalculator;
            if (calculator && calculator.generateClearfileSheetOptions) {
                sheetsSelect.innerHTML = calculator.generateClearfileSheetOptions();
            } else {
                // フォールバック
                sheetsSelect.innerHTML = '<option value="">枚数を選択</option>';
                for (let i = 1; i <= 50; i++) {
                    sheetsSelect.innerHTML += `<option value="${i}">${i}枚</option>`;
                }
            }
        } else {
            sheetsContainer.classList.add('hidden');
            sheetsSelect.value = '';
        }
        
        this.updateTotals();
    }

    // 合計を更新
    updateTotals() {
        const calculator = new DonationCalculator();
        const totals = calculator.calculateTotals();
        
        document.getElementById('totalDonors').textContent = totals.donorCount;
        document.getElementById('totalAmount').textContent = `¥${totals.totalAmount.toLocaleString()}`;
        document.getElementById('totalGifts').textContent = totals.giftCount;

        // 1ページ目の寄付金と照合
        this.validateDonationAmount(totals.totalAmount);
    }

    // 1ページ目の寄付金と照合
    validateDonationAmount(page2Total) {
        if (!this.basicInfo) return;

        const page1Donation = this.basicInfo.breakdownDetails.donationAmount;
        
        if (page1Donation > 0 && page2Total !== page1Donation) {
            // 警告を表示
            const warningDiv = document.getElementById('amount-warning') || this.createWarningDiv();
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
                                <strong>金額の不一致:</strong> 1ページ目の寄付金（¥${page1Donation.toLocaleString()}）と2ページ目の合計金額（¥${page2Total.toLocaleString()}）が一致しません。
                            </p>
                        </div>
                    </div>
                </div>
            `;
            warningDiv.style.display = 'block';
        } else {
            const warningDiv = document.getElementById('amount-warning');
            if (warningDiv) {
                warningDiv.style.display = 'none';
            }
        }
    }

    // 警告表示用のdivを作成
    createWarningDiv() {
        const warningDiv = document.createElement('div');
        warningDiv.id = 'amount-warning';
        const form = document.getElementById('donationForm');
        form.insertBefore(warningDiv, form.firstChild);
        return warningDiv;
    }

    // 2ページ目のデータを読み込み（復元）
    loadPage2Data() {
        try {
            const page2Data = localStorage.getItem('page2Data');
            if (page2Data) {
                const data = JSON.parse(page2Data);
                this.restorePage2Data(data);
                console.log('2ページ目のデータを復元しました:', data);
            }
        } catch (error) {
            console.error('2ページ目データ復元エラー:', error);
        }
    }

    // 2ページ目のデータを復元
    restorePage2Data(data) {
        // 既存の寄付者行をクリア
        this.donorsContainer.innerHTML = '';
        this.donorCount = 0;

        // 寄付者データを復元
        if (data.donors && data.donors.length > 0) {
            data.donors.forEach(donor => {
                this.donorCount++;
                const donorRow = this.createDonorRow(this.donorCount);
                this.donorsContainer.appendChild(donorRow);

                // データを設定
                this.setDonorRowData(donorRow, this.donorCount, donor);
            });
        }

        this.updateTotals();
    }

    // 寄付者行にデータを設定
    setDonorRowData(row, index, donor) {
        // インデックスに依存せず、接頭辞一致で要素を取得
        const nameInput = row.querySelector('input[name^="donorName_"]');
        const amountSelect = row.querySelector('select[name^="donorAmountType_"]');
        const amountCustom = row.querySelector('input[name^="donorAmountCustom_"]');
        const giftSelect = row.querySelector('select[name^="donorGift_"]');
        const sheetsSelect = row.querySelector('select[name^="donorSheets_"]');
        const noteInput = row.querySelector('input[name^="donorNote_"]');

        if (nameInput) nameInput.value = donor.name || '';
        if (noteInput) noteInput.value = donor.note || '';

        // 金額タイプを設定し、変更イベントをトリガー
        if (amountSelect) {
            amountSelect.value = donor.amountType || '';
            // 金額タイプ変更時の処理を実行
            const amountInput = row.querySelector('.donor-amount-input');
            const sheetsContainer = row.querySelector('.donor-sheets-container');
            this.handleAmountTypeChange(amountSelect, amountInput, giftSelect, sheetsContainer, sheetsSelect);
        }

        // カスタム金額を設定
        if (amountCustom) {
            amountCustom.value = donor.amountCustom || '';
            // カスタム金額入力時の処理を実行（返礼品選択肢を更新）
            if (donor.amountCustom && amountSelect && amountSelect.value === 'other') {
                const sheetsContainer = row.querySelector('.donor-sheets-container');
                this.handleAmountInputChange(amountCustom, giftSelect, sheetsContainer, sheetsSelect);
            }
        }

        // 返礼品を設定し、変更イベントをトリガー
        if (giftSelect) {
            giftSelect.value = donor.gift || '';
            // 返礼品変更時の処理を実行（クリアファイル枚数の表示制御）
            const sheetsContainer = row.querySelector('.donor-sheets-container');
            this.handleGiftChange(giftSelect, sheetsContainer, sheetsSelect);
        }

        // クリアファイル枚数を設定
        if (sheetsSelect) {
            sheetsSelect.value = donor.sheets || '';
        }

        // 直接振込情報を設定
        const directTransferCheckbox = row.querySelector('input[name^="donorDirectTransfer_"]');
        const directDateInput = row.querySelector('input[name^="donorDirectDate_"]');
        const directDateContainer = row.querySelector('.donor-direct-date');
        
        if (directTransferCheckbox && donor.isDirectTransfer) {
            directTransferCheckbox.checked = true;
            if (directDateContainer) {
                directDateContainer.classList.remove('hidden');
            }
        }
        
        if (directDateInput && donor.directTransferDate) {
            directDateInput.value = donor.directTransferDate;
        }
    }

    // 2ページ目のデータを保存
    savePage2Data() {
        try {
            const data = {
                donors: this.collectPage2Donors()
            };
            localStorage.setItem('page2Data', JSON.stringify(data));
            console.log('2ページ目のデータを保存しました:', data);
        } catch (error) {
            console.error('2ページ目データ保存エラー:', error);
        }
    }

    // 2ページ目の寄付者データを収集
    collectPage2Donors() {
        const donors = [];
        const donorRows = document.querySelectorAll('.donor-row');
        
        donorRows.forEach((row) => {
            const nameInput = row.querySelector('input[name^="donorName_"]');
            const amountSelect = row.querySelector('select[name^="donorAmountType_"]');
            const amountCustom = row.querySelector('input[name^="donorAmountCustom_"]');
            const giftSelect = row.querySelector('select[name^="donorGift_"]');
            const sheetsSelect = row.querySelector('select[name^="donorSheets_"]');
            const noteInput = row.querySelector('input[name^="donorNote_"]');
            const directTransferCheckbox = row.querySelector('input[name^="donorDirectTransfer_"]');
            const directDateInput = row.querySelector('input[name^="donorDirectDate_"]');

            donors.push({
                name: nameInput?.value || '',
                amountType: amountSelect?.value || '',
                amountCustom: amountCustom?.value || '',
                gift: giftSelect?.value || '',
                sheets: sheetsSelect?.value || '',
                note: noteInput?.value || '',
                isDirectTransfer: directTransferCheckbox?.checked || false,
                directTransferDate: directDateInput?.value || ''
            });
        });

        return donors;
    }

    // 1ページ目に戻る
    backToPage1() {
        // 2ページ目のデータを保存
        this.savePage2Data();
        
        // 1ページ目に遷移
        window.location.href = 'page1.html';
    }

    // 確認モーダルを表示
    showConfirmModal() {
        if (!this.confirmModal) return;

        // 確認モーダルにデータを表示
        this.displayConfirmData();
        
        // モーダルを表示
        this.confirmModal.classList.remove('hidden');
    }

    // 確認モーダルを非表示
    hideConfirmModal() {
        if (this.confirmModal) {
            this.confirmModal.classList.add('hidden');
        }
    }

    // 確認モーダルにデータを表示
    displayConfirmData() {
        try {
            // 基本情報を表示
            this.displayConfirmBasicInfo();
            
            // 振込金額を表示
            this.displayConfirmTransferAmounts();
            
            // 内訳を表示
            this.displayConfirmBreakdown();
            
            // 寄付者情報を表示
            this.displayConfirmDonors();
            
            // 合計情報を表示
            this.displayConfirmTotals();
            
        } catch (error) {
            console.error('確認データ表示エラー:', error);
        }
    }

    // 確認モーダル：基本情報を表示
    displayConfirmBasicInfo() {
        if (!this.basicInfo || !this.basicInfo.basicInfo) return;
        
        const basicInfo = this.basicInfo.basicInfo;
        
        const gradeEl = document.getElementById('confirm-grade');
        const nameEl = document.getElementById('confirm-name');
        const transferDateEl = document.getElementById('confirm-transferDate');
        
        if (gradeEl) gradeEl.textContent = basicInfo.grade || '-';
        if (nameEl) nameEl.textContent = basicInfo.name || '-';
        if (transferDateEl) transferDateEl.textContent = basicInfo.transferDate || '-';
        
        // 兄弟情報を表示
        this.displayConfirmSiblingInfo(basicInfo);
    }

    // 確認モーダル：兄弟情報を表示
    displayConfirmSiblingInfo(basicInfo) {
        const siblingInfoEl = document.getElementById('confirm-sibling-info');
        const siblingCountEl = document.getElementById('confirm-siblingCount');
        const siblingNamesEl = document.getElementById('confirm-siblingNames');
        
        if (!siblingInfoEl || !siblingCountEl || !siblingNamesEl) return;
        
        // 兄弟がいる場合のみ表示
        if (basicInfo.hasSibling && basicInfo.siblingCount > 0) {
            siblingInfoEl.classList.remove('hidden');
            siblingCountEl.textContent = `${basicInfo.siblingCount}人`;
            
            // 兄弟の名前を表示
            siblingNamesEl.innerHTML = '';
            for (let i = 1; i <= basicInfo.siblingCount; i++) {
                const siblingName = basicInfo[`sibling${i}Name`];
                if (siblingName) {
                    const nameDiv = document.createElement('div');
                    nameDiv.className = 'flex justify-between';
                    nameDiv.innerHTML = `
                        <span class="text-gray-600">兄弟${i}:</span>
                        <span class="font-medium">${siblingName}</span>
                    `;
                    siblingNamesEl.appendChild(nameDiv);
                }
            }
        } else {
            siblingInfoEl.classList.add('hidden');
        }
    }

    // 確認モーダル：振込金額を表示
    displayConfirmTransferAmounts() {
        if (!this.basicInfo || !this.basicInfo.breakdownDetails) return;
        
        const breakdown = this.basicInfo.breakdownDetails;
        const transferAmount = parseInt(breakdown.transferAmount) || 0;
        const directTransferAmount = parseInt(breakdown.directTransferAmount) || 0;
        
        const transferAmountEl = document.getElementById('confirm-transferAmount');
        const directTransferAmountEl = document.getElementById('confirm-directTransferAmount');
        
        if (transferAmountEl) transferAmountEl.textContent = `¥${transferAmount.toLocaleString()}`;
        if (directTransferAmountEl) directTransferAmountEl.textContent = `¥${directTransferAmount.toLocaleString()}`;
    }

    // 確認モーダル：内訳を表示
    displayConfirmBreakdown() {
        if (!this.basicInfo || !this.basicInfo.breakdownDetails) return;
        
        const breakdown = this.basicInfo.breakdownDetails;
        
        // 各項目の金額を計算
        const donationAmount = breakdown.donationAmount || 0;
        
        // 負担金は新しい形式では直接金額が送信される
        let burdenAmount = 0;
        if (typeof breakdown.burdenFee === 'number') {
            burdenAmount = breakdown.burdenFee;
        } else if (typeof breakdown.burdenFee === 'string') {
            burdenAmount = parseInt(breakdown.burdenFee) || 0;
        }
        
        const tshirtAmount = breakdown.tshirtAmount || 0;
        const otherFee = breakdown.otherFee || 0;
        const encouragementAmount = breakdown.hasEncouragementFee ? 1000 : 0; // 激励会協力金
        
        // 合計を計算
        const breakdownTotal = donationAmount + burdenAmount + tshirtAmount + otherFee + encouragementAmount;
        
        // 表示
        const donationAmountEl = document.getElementById('confirm-donationAmount');
        const burdenAmountEl = document.getElementById('confirm-burdenAmount');
        const tshirtAmountEl = document.getElementById('confirm-tshirtAmount');
        const otherFeeEl = document.getElementById('confirm-otherFee');
        const encouragementAmountEl = document.getElementById('confirm-encouragementAmount');
        const breakdownTotalEl = document.getElementById('confirm-breakdownTotal');
        
        if (donationAmountEl) donationAmountEl.textContent = `¥${donationAmount.toLocaleString()}`;
        if (burdenAmountEl) burdenAmountEl.textContent = `¥${burdenAmount.toLocaleString()}`;
        if (tshirtAmountEl) tshirtAmountEl.textContent = `¥${tshirtAmount.toLocaleString()}`;
        if (otherFeeEl) otherFeeEl.textContent = `¥${otherFee.toLocaleString()}`;
        if (encouragementAmountEl) encouragementAmountEl.textContent = `¥${encouragementAmount.toLocaleString()}`;
        if (breakdownTotalEl) breakdownTotalEl.textContent = `¥${breakdownTotal.toLocaleString()}`;
    }

    // 確認モーダル：寄付者情報を表示
    displayConfirmDonors() {
        const donorsContainer = document.getElementById('confirm-donors');
        if (!donorsContainer) return;
        
        // コンテナをクリア
        donorsContainer.innerHTML = '';
        
        const formData = this.collectFormData();
        const donors = formData.donors;
        
        // 各寄付者の情報を表示
        donors.forEach((donor, index) => {
            if (!donor.name || !donor.amount) return; // 空の寄付者はスキップ
            
            const donorDiv = document.createElement('div');
            donorDiv.className = 'bg-white rounded p-2 text-xs border';
            
            // 返礼品の表示
            let giftDisplay = donor.gift || 'なし';
            if (donor.gift === 'クリアファイル' && donor.sheets) {
                giftDisplay = `クリアファイル(${donor.sheets}枚)`;
            }
            
            // 直接振込日の表示
            const directTransferDateDisplay = donor.isDirectTransfer && donor.directTransferDate 
                ? `<div>直接振込日: ${donor.directTransferDate}</div>` 
                : '';
            
            donorDiv.innerHTML = `
                <div class="flex justify-between items-start">
                    <div class="flex-1">
                        <div class="font-medium text-gray-800">${donor.name}</div>
                        <div class="text-gray-600 mt-1">
                            <div>金額: ¥${donor.amount.toLocaleString()}</div>
                            <div>返礼品: ${giftDisplay}</div>
                            ${directTransferDateDisplay}
                            ${donor.note ? `<div>その他特筆事項: ${donor.note}</div>` : ''}
                        </div>
                    </div>
                </div>
            `;
            
            donorsContainer.appendChild(donorDiv);
        });
        
        // 寄付者がいない場合
        if (donorsContainer.children.length === 0) {
            donorsContainer.innerHTML = '<p class="text-xs text-gray-500 text-center py-2">寄付者情報がありません</p>';
        }
    }

    // 確認モーダル：合計情報を表示
    displayConfirmTotals() {
        const formData = this.collectFormData();
        const donors = formData.donors;
        
        // 集計を計算
        let totalDonors = 0;
        let totalAmount = 0;
        let totalGifts = 0;
        
        donors.forEach(donor => {
            if (donor.name && donor.amount) {
                totalDonors++;
                totalAmount += donor.amount;
                
                if (donor.gift && donor.gift !== '') {
                    totalGifts++;
                }
            }
        });
        
        // 表示
        const totalDonorsEl = document.getElementById('confirm-totalDonors');
        const totalAmountEl = document.getElementById('confirm-totalAmount');
        const totalGiftsEl = document.getElementById('confirm-totalGifts');
        
        if (totalDonorsEl) totalDonorsEl.textContent = totalDonors.toString();
        if (totalAmountEl) totalAmountEl.textContent = `¥${totalAmount.toLocaleString()}`;
        if (totalGiftsEl) totalGiftsEl.textContent = totalGifts.toString();
    }

    // 確認後の送信処理
    async confirmAndSubmit() {
        // モーダルを非表示
        this.hideConfirmModal();
        
        // 実際の送信処理を実行
        await this.performSubmit();
    }

    // 実際の送信処理
    async performSubmit() {
        const formData = this.collectFormData();
        
        try {
            this.showLoading(true);
            const result = await this.submitToGoogleSheets(formData);
            this.showSuccess(result);
            // 送信成功時はresetFormを呼ばない（送信控えページに遷移するため）
        } catch (error) {
            console.error('送信エラー:', error);
            alert('送信中にエラーが発生しました。もう一度お試しください。');
        } finally {
            this.showLoading(false);
        }
    }

    // フォーム送信処理（確認モーダル表示に変更）
    async handleSubmit(e) {
        e.preventDefault();
        
        if (!this.validateForm()) {
            alert('必須項目を入力してください。');
            return;
        }

        if (!this.validateAmounts()) {
            alert('金額の照合に失敗しました。1ページ目の寄付金と2ページ目の合計金額を確認してください。');
            return;
        }

        // 確認モーダルを表示
        this.showConfirmModal();
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
        if (!this.basicInfo) return true;

        const calculator = new DonationCalculator();
        const totals = calculator.calculateTotals();
        const page1Donation = this.basicInfo.breakdownDetails.donationAmount;

        // 1ページ目の寄付金が0の場合は照合しない
        if (page1Donation === 0) {
            return true;
        }

        return totals.totalAmount === page1Donation;
    }

    // フォームデータを収集
    collectFormData() {
        const data = {
            basicInfo: this.basicInfo,
            submissionDate: new Date().toISOString(),
            donors: []
        };

        // 寄付者情報を行ごとに収集（インデックス非依存）
        const donorRows = document.querySelectorAll('.donor-row');
        donorRows.forEach((row) => {
            const nameInput = row.querySelector('input[name^="donorName_"]');
            const amountTypeSelect = row.querySelector('select[name^="donorAmountType_"]');
            const amountCustomInput = row.querySelector('input[name^="donorAmountCustom_"]');
            const giftIdSelect = row.querySelector('select[name^="donorGift_"]');
            const sheetsSelect = row.querySelector('select[name^="donorSheets_"]');
            const noteInput = row.querySelector('input[name^="donorNote_"]');
            const directTransferCheckbox = row.querySelector('input[name^="donorDirectTransfer_"]');
            const directDateInput = row.querySelector('input[name^="donorDirectDate_"]');

            const name = nameInput?.value || '';
            const amountType = amountTypeSelect?.value || '';
            const amountCustom = amountCustomInput?.value || '';
            const giftId = giftIdSelect?.value || '';
            const sheets = sheetsSelect?.value || '';
            const note = noteInput?.value || '';
            const isDirectTransfer = !!(directTransferCheckbox?.checked);
            const directTransferDate = directDateInput?.value || '';

            // 金額を計算
            let amount = 0;
            if (amountType === 'other' && amountCustom) {
                amount = parseInt(amountCustom) || 0;
            } else {
                amount = parseInt(amountType) || 0; // シート側の任意金額に対応
            }

            // 返礼品は名前をそのまま採用
            let giftName = '';
            if (giftId) {
                giftName = giftId;
            }

            data.donors.push({
                name,
                amount,
                gift: giftName,
                sheets,
                note,
                isDirectTransfer,
                directTransferDate
            });
        });

        return data;
    }

    // Google Sheetsに送信
    async submitToGoogleSheets(data) {
        if (window.googleSheetsManager) {
            const result = await window.googleSheetsManager.submitFormData(data);
            if (!result.success) {
                throw new Error(result.error);
            }
            return result;
        } else {
            throw new Error('Google Sheets連携が初期化されていません');
        }
    }

    // ローディング表示
    showLoading(show) {
        const loading = document.getElementById('loading');
        loading.classList.toggle('hidden', !show);
    }

    // 成功メッセージ
    showSuccess(result = null) {
        // 送信控え用のデータをlocalStorageに保存
        const submissionData = this.collectFormData();
        localStorage.setItem('submissionData', JSON.stringify(submissionData));
        
        // 2ページ目のデータをクリア（送信完了のため）
        localStorage.removeItem('page2Data');
        
        // 送信控えページに遷移
        window.location.href = 'receipt.html';
    }

    // フォームをリセット
    resetForm() {
        this.form.reset();
        this.donorsContainer.innerHTML = '';
        this.donorCount = 0;
        this.addInitialDonors();
        this.updateTotals();
    }

    // フォーム設定を読み込み
    async loadFormSettings() {
        try {
            if (window.googleSheetsManager) {
                const settings = await window.googleSheetsManager.getFormSettings();
                if (settings) {
                    this.updateFormSubtitle(settings);
                }
            }
        } catch (error) {
            console.error('フォーム設定の読み込みエラー:', error);
            // エラーの場合はデフォルト値のまま
        }
    }

    // フォームサブタイトルを更新
    updateFormSubtitle(settings) {
        const subtitleElement = document.getElementById('formSubtitle');
        
        if (subtitleElement && settings.formSubtitle) {
            subtitleElement.textContent = settings.formSubtitle;
        }
    }
}

// フォームを初期化
document.addEventListener('DOMContentLoaded', () => {
        // GoogleSheetsManagerの初期化を待つ
        const initForm = () => {
            if (window.googleSheetsManager) {
                const form = new DonationForm();
                // フォーム設定を取得してタイトルとページタイトルを更新
                window.googleSheetsManager.getFormSettings().then((settings) => {
                    try {
                        const headerTitle = document.querySelector('header h1');
                        if (headerTitle && settings.formTitle) headerTitle.textContent = settings.formTitle;
                        if (settings.formTitle) document.title = settings.formTitle;
                        const subtitle = document.getElementById('formSubtitle');
                        if (subtitle && settings.formSubtitle) subtitle.textContent = settings.formSubtitle;
                    } catch (e) { console.error(e); }
                }).catch(() => {});
            } else {
                setTimeout(initForm, 100); // 100ms後に再試行
            }
        };
        
        initForm();
}); 