// 送信控えページ管理クラス
class ReceiptManager {
    constructor() {
        this.basicInfo = null;
        this.submissionData = null;
        this.initializePage();
    }

    // ページを初期化
    initializePage() {
        // localStorageからデータを読み込み
        this.loadDataFromStorage();
        
        // データが存在しない場合はpage1にリダイレクト
        if (!this.basicInfo || !this.submissionData) {
            console.warn('送信控えデータが見つかりません。page1にリダイレクトします。');
            window.location.href = 'page1.html';
            return;
        }
        
        // フォーム設定のタイトル/サブタイトルを反映（sheets.jsのgetFormSettingsを利用）
        if (window.googleSheetsManager && typeof window.googleSheetsManager.getFormSettings === 'function') {
            window.googleSheetsManager.getFormSettings().then((settings) => {
                try {
                    if (settings && settings.formTitle) {
                        document.title = `送信控え - ${settings.formTitle}`;
                    }
                    if (settings && settings.formSubtitle) {
                        const subtitle = document.getElementById('formSubtitle');
                        if (subtitle) subtitle.textContent = settings.formSubtitle;
                    }
                } catch (e) { console.error(e); }
            }).catch((e) => console.error('フォーム設定取得エラー:', e));
        }
        
        // 送信控えを表示
        this.displayReceipt();
        
        // データをクリア（セキュリティのため）
        this.clearStorageData();
    }

    // localStorageからデータを読み込み
    loadDataFromStorage() {
        try {
            const basicInfoData = localStorage.getItem('basicInfo');
            const submissionDataStr = localStorage.getItem('submissionData');
            
            if (basicInfoData) {
                this.basicInfo = JSON.parse(basicInfoData);
            }
            
            if (submissionDataStr) {
                this.submissionData = JSON.parse(submissionDataStr);
            }
            
            console.log('送信控えデータを読み込みました:', {
                basicInfo: this.basicInfo,
                submissionData: this.submissionData
            });
            
        } catch (error) {
            console.error('データ読み込みエラー:', error);
        }
    }

    // 送信控えを表示
    displayReceipt() {
        try {
            // 基本情報を表示
            this.displayBasicInfo();
            
            // 振込金額を表示
            this.displayTransferAmounts();
            
            // 内訳を表示
            this.displayBreakdown();
            
            // 寄付者情報を表示
            this.displayDonors();
            
            // 合計情報を表示
            this.displayTotals();
            
            // 送信日時を表示
            this.displaySubmissionDate();
            
        } catch (error) {
            console.error('送信控え表示エラー:', error);
        }
    }

    // 基本情報を表示
    displayBasicInfo() {
        if (!this.basicInfo || !this.basicInfo.basicInfo) return;
        
        const basicInfo = this.basicInfo.basicInfo;
        
        document.getElementById('receipt-grade').textContent = basicInfo.grade || '-';
        document.getElementById('receipt-name').textContent = basicInfo.name || '-';
        document.getElementById('receipt-transferDate').textContent = basicInfo.transferDate || '-';
        
        // 兄弟情報を表示
        this.displaySiblingInfo(basicInfo);
    }

    // 兄弟情報を表示
    displaySiblingInfo(basicInfo) {
        const siblingInfoEl = document.getElementById('receipt-sibling-info');
        const siblingCountEl = document.getElementById('receipt-siblingCount');
        const siblingNamesEl = document.getElementById('receipt-siblingNames');
        
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

    // 振込金額を表示
    displayTransferAmounts() {
        if (!this.basicInfo || !this.basicInfo.breakdownDetails) return;
        
        const breakdown = this.basicInfo.breakdownDetails;
        const transferAmount = parseInt(breakdown.transferAmount) || 0;
        const directTransferAmount = parseInt(breakdown.directTransferAmount) || 0;
        
        document.getElementById('receipt-transferAmount').textContent = `¥${transferAmount.toLocaleString()}`;
        document.getElementById('receipt-directTransferAmount').textContent = `¥${directTransferAmount.toLocaleString()}`;
    }

    // 内訳を表示
    displayBreakdown() {
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
        const encouragementAmount = breakdown.hasEncouragementFee ? 1000 : 0;
        
        // 合計を計算
        const breakdownTotal = donationAmount + burdenAmount + tshirtAmount + otherFee + encouragementAmount;
        
        // 表示
        document.getElementById('receipt-donationAmount').textContent = `¥${donationAmount.toLocaleString()}`;
        document.getElementById('receipt-burdenAmount').textContent = `¥${burdenAmount.toLocaleString()}`;
        document.getElementById('receipt-tshirtAmount').textContent = `¥${tshirtAmount.toLocaleString()}`;
        
        // その他の費用の表示（要素が存在する場合）
        const otherFeeEl = document.getElementById('receipt-otherFee');
        if (otherFeeEl) {
            otherFeeEl.textContent = `¥${otherFee.toLocaleString()}`;
        }
        
        document.getElementById('receipt-encouragementAmount').textContent = `¥${encouragementAmount.toLocaleString()}`;
        document.getElementById('receipt-breakdownTotal').textContent = `¥${breakdownTotal.toLocaleString()}`;
    }

    // 寄付者情報を表示
    displayDonors() {
        if (!this.submissionData || !this.submissionData.donors) return;
        
        const donors = this.submissionData.donors;
        const donorsContainer = document.getElementById('receipt-donors');
        
        // コンテナをクリア
        donorsContainer.innerHTML = '';
        
        // 各寄付者の情報を表示
        donors.forEach((donor, index) => {
            if (!donor.name || !donor.amount) return; // 空の寄付者はスキップ
            
            const donorDiv = document.createElement('div');
            donorDiv.className = 'bg-gray-50 rounded p-2 text-xs';
            
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

    // 合計情報を表示
    displayTotals() {
        if (!this.submissionData || !this.submissionData.donors) return;
        
        const donors = this.submissionData.donors;
        
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
        document.getElementById('receipt-totalDonors').textContent = totalDonors.toString();
        document.getElementById('receipt-totalAmount').textContent = `¥${totalAmount.toLocaleString()}`;
        document.getElementById('receipt-totalGifts').textContent = totalGifts.toString();
    }

    // 送信日時を表示
    displaySubmissionDate() {
        const submissionDate = this.basicInfo?.submissionDate || this.submissionData?.submissionDate;
        
        if (submissionDate) {
            try {
                const date = new Date(submissionDate);
                const formattedDate = date.toLocaleString('ja-JP', {
                    year: 'numeric',
                    month: '2-digit',
                    day: '2-digit',
                    hour: '2-digit',
                    minute: '2-digit'
                });
                document.getElementById('receipt-submissionDate').textContent = formattedDate;
            } catch (error) {
                console.error('日時フォーマットエラー:', error);
                document.getElementById('receipt-submissionDate').textContent = '不明';
            }
        } else {
            document.getElementById('receipt-submissionDate').textContent = '不明';
        }
    }

    // localStorageのデータをクリア
    clearStorageData() {
        try {
            localStorage.removeItem('basicInfo');
            localStorage.removeItem('submissionData');
            console.log('送信控えデータをクリアしました');
        } catch (error) {
            console.error('データクリアエラー:', error);
        }
    }
}

// ページ読み込み時に初期化
document.addEventListener('DOMContentLoaded', () => {
    new ReceiptManager();
});

// ブラウザの戻るボタン対策
window.addEventListener('beforeunload', (e) => {
    // 特に何もしない（データは既にクリア済み）
});

// ページを離れる際の確認（オプション）
window.addEventListener('pagehide', () => {
    console.log('送信控えページを離れます');
});