// 寄付計算クラス
class DonationCalculator {
    constructor() {
        // 返礼品ルールを外部から設定可能にする
        this.giftRules = this.loadGiftRules();
    }

    // 返礼品ルールを読み込み（外部設定ファイルから読み込むことも可能）
    loadGiftRules() {
        // デフォルトの返礼品ルール
        return {
            towel: { minAmount: 10000, name: 'タオル', category: 'premium' },
            sweets_large: { minAmount: 10000, name: 'お菓子大', category: 'standard' },
            keychain: { minAmount: 5000, name: 'キーホルダー', category: 'standard' },
            sweets_small: { minAmount: 5000, name: 'お菓子小', category: 'standard' },
            clearfile: { minAmount: 0, name: '③ クリアファイル', category: 'basic', hasSheets: true }
        };
    }

    // 返礼品オプションを生成（条件分岐なしで全件）
    generateGiftOptions(_amount) {
        const options = ['<option value="">返礼品なし</option>'];
        Object.entries(this.giftRules).forEach(([key, rule]) => {
            // 値はID、表示は名前（IDで統一）
            options.push(`<option value="${key}">${rule.name}</option>`);
        });
        return options.join('');
    }

    // クリアファイル枚数オプションを生成
    generateClearfileSheetOptions() {
        const options = ['<option value="">枚数を選択</option>'];
        
        for (let i = 1; i <= 50; i++) {
            options.push(`<option value="${i}">${i}枚</option>`);
        }
        
        return options.join('');
    }

    // 返礼品ルールを更新（管理画面用）
    updateGiftRules(newRules) {
        this.giftRules = { ...this.giftRules, ...newRules };
        // 必要に応じてローカルストレージに保存
        localStorage.setItem('giftRules', JSON.stringify(this.giftRules));
    }

    // 返礼品ルールをリセット（デフォルトに戻す）
    resetGiftRules() {
        this.giftRules = this.loadGiftRules();
        localStorage.removeItem('giftRules');
    }

    // スプレッドシートから読み込んだ設定を適用
    updateGiftRulesFromSheet(sheetGiftRules) {
        if (sheetGiftRules && typeof sheetGiftRules === 'object') {
            this.giftRules = sheetGiftRules;
            console.log('スプレッドシートから返礼品設定を適用しました');
        }
    }

    // スプレッドシートから設定を読み込み（初期化時用）
    async loadFromSheet() {
        if (window.sheetsManager) {
            const success = await window.sheetsManager.loadGiftRulesFromSheet();
            if (success) {
                this.updateGiftRulesFromSheet(window.sheetsManager.giftRules);
            }
            return success;
        }
        return false;
    }

    // 合計を計算
    calculateTotals() {
        const donorRows = document.querySelectorAll('.donor-row');
        let totalAmount = 0;
        let giftCount = 0;
        let donorCount = donorRows.length;

        donorRows.forEach(row => {
            const amountTypeSelect = row.querySelector('.donor-amount-type');
            const amountCustomInput = row.querySelector('.donor-amount-input input');
            const giftSelect = row.querySelector('.donor-gift');
            
            if (amountTypeSelect && amountTypeSelect.value) {
                let amount = 0;
                if (amountTypeSelect.value === 'other' && amountCustomInput) {
                    amount = parseInt(amountCustomInput.value) || 0;
                } else {
                    amount = parseInt(amountTypeSelect.value) || 0; // シート側の任意金額に対応
                }
                totalAmount += amount;
                
                // 返礼品が選択されている場合はカウント
                if (giftSelect && giftSelect.value) {
                    giftCount++;
                }
            }
        });

        return {
            donorCount,
            totalAmount,
            giftCount
        };
    }

    // 金額に基づいて返礼品を自動選択
    autoSelectGift(amountInput, giftSelect) {
        const amount = parseInt(amountInput.value) || 0;
        
        // 金額に応じて適切な返礼品を提案
        let suggestedGift = '';
        for (const [giftKey, rule] of Object.entries(this.giftRules)) {
            if (amount >= rule.minAmount) {
                suggestedGift = giftKey;
            }
        }

        // 現在選択されている返礼品がない場合のみ自動選択
        if (!giftSelect.value && suggestedGift) {
            giftSelect.value = suggestedGift;
            this.showGiftSuggestion(amountInput, suggestedGift);
        }
    }

    // 返礼品提案を表示
    showGiftSuggestion(amountInput, giftKey) {
        const rule = this.giftRules[giftKey];
        if (rule) {
            console.log(`${rule.name}が選択されました（¥${rule.minAmount}以上）`);
        }
    }

    // 返礼品ルールを取得
    getGiftRules() {
        return this.giftRules;
    }

    // 金額から返礼品を取得
    getGiftForAmount(amount) {
        for (const [giftKey, rule] of Object.entries(this.giftRules)) {
            if (amount >= rule.minAmount) {
                return { key: giftKey, ...rule };
            }
        }
        return null;
    }

    // 返礼品の説明を取得
    getGiftDescription(giftKey) {
        const rule = this.giftRules[giftKey];
        return rule ? rule.name : '';
    }

    // 金額の妥当性をチェック
    validateAmount(amount) {
        if (amount < 0) {
            return { valid: false, message: '金額は0以上で入力してください' };
        }
        if (amount > 1000000) {
            return { valid: false, message: '金額は1,000,000円以下で入力してください' };
        }
        return { valid: true };
    }

    // 合計金額の表示形式を整形
    formatAmount(amount) {
        return `¥${amount.toLocaleString()}`;
    }

    // 寄付者ごとの詳細計算
    calculateDonorDetails(donorData) {
        const amount = donorData.amount || 0;
        const gift = donorData.gift || '';
        
        return {
            amount: amount,
            gift: gift,
            giftDescription: this.getGiftDescription(gift),
            hasGift: !!gift
        };
    }

    // 統計情報を計算
    calculateStatistics(donors) {
        if (!donors || donors.length === 0) {
            return {
                totalAmount: 0,
                averageAmount: 0,
                maxAmount: 0,
                minAmount: 0,
                giftDistribution: {},
                totalDonors: 0
            };
        }

        const amounts = donors.map(d => d.amount || 0).filter(a => a > 0);
        const giftCounts = {};
        
        donors.forEach(donor => {
            if (donor.gift) {
                giftCounts[donor.gift] = (giftCounts[donor.gift] || 0) + 1;
            }
        });

        return {
            totalAmount: amounts.reduce((sum, amount) => sum + amount, 0),
            averageAmount: amounts.length > 0 ? Math.round(amounts.reduce((sum, amount) => sum + amount, 0) / amounts.length) : 0,
            maxAmount: amounts.length > 0 ? Math.max(...amounts) : 0,
            minAmount: amounts.length > 0 ? Math.min(...amounts) : 0,
            giftDistribution: giftCounts,
            totalDonors: donors.length
        };
    }

    // 金額選択の検証
    validateAmountSelection(amountType, customAmount) {
        if (!amountType) {
            return { valid: false, message: '金額を選択してください' };
        }
        
        if (amountType === 'other' && (!customAmount || customAmount <= 0)) {
            return { valid: false, message: 'その他の場合は金額を入力してください' };
        }
        
        return { valid: true };
    }
}

// グローバルに計算機インスタンスを作成
window.donationCalculator = new DonationCalculator(); 