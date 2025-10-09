const { createApp } = Vue;

createApp({
    data() {
        return {
            workbook: null,
            errorFields: [],
            selectedError: '',
            errorCount: 0,
            dataTable: [],
            // 定義預設欄位組
            fieldSets: [
                {
                    name: '彙寄繳費通知',
                    value:'BILL_OFFICE_CODE|BILL_EQUIP_NO|BILL_YEAR_MONTH|BILL_CYCLE|BILL_TYPE|RECEIPT_NO'
                },
                {
                    name: '彙寄地址參照檔',
                    value: 'CYCLE|BILL-MONTH|BILL-OFFICE-CODE|BILL-EQUIP-NO|DELIVERY-NO'
                },
                {
                    name: '大客戶電子檔',
                    value: 'DISK-DELIVERY-NO|DISK-ACCOUNT|DISK-IDCARD-NO|DISK-BILL-NO|DISK-OPERID|DISK-OFFICE-CODE|DISK-CTRL-CODE|DISK-TEL-OPERID|DISK-TEL-OFFICE-CODE|DISK-PBXNO|DISK-TELNO|DISK-ITEM-OVERFLAG|DISK-SEQ-NO|DISK-INDEX-CODE|DISK-INDEX-KEY|DISK-MONTH|DISK-CYCLE'
                },
            ],
            // 選擇的欄位組索引
            selectedFieldSetIndex: 0,
            // 使用字串定義動態欄位，更簡潔
            dynamicFieldsString: 'BILL_OFFICE_CODE|BILL_EQUIP_NO|BILL_YEAR_MONTH|BILL_CYCLE|BILL_TYPE|RECEIPT_NO',
            // 新增：用於編輯的字串副本
            editableDynamicFieldsString: '',
            // 新增：控制編輯器顯示/隱藏
            showFieldEditor: false,
            // 新增：備份原始字串
            originalDynamicFieldsString: '',
            // 修復：將 filters 改回 data property
            filters: {
                row: '',
                mainValue: '',
                // 新增：相似度過濾欄位
                similarity: '',
                derivedValue: '',
                sheetName: ''
            },
            loading: false,
            currentPage: 1,
            itemsPerPage: 10,
            modalHeaders: [],
            modalRows: []
        };
    },
    computed: {
        // 從字串動態產生欄位配置，支援逗號和直線符號分隔
        dynamicFields() {
            // 支援逗號(,)和直線符號(|)作為分隔符號
            const separator = this.dynamicFieldsString.includes('|') ? '|' : ',';
            return this.dynamicFieldsString.split(separator).map(field => {
                const trimmedField = field.trim();
                return {
                    key: trimmedField,
                    label: trimmedField,
                    placeholder: `過濾 ${trimmedField}`
                };
            }).filter(field => field.key); // 過濾空欄位
        },
        filteredData() {
            return this.dataTable.filter(row => {
                return Object.entries(this.filters).every(([key, value]) => {
                    if (value === '' || value === null || value === undefined) {
                        return true;
                    }
                    if (key === 'similarity') {
                        return this.applyNumericFilter(row.similarity, value);
                    }
                    return String(row[key] ?? '').toUpperCase().includes(String(value).toUpperCase());
                });
            });
        },
        totalPages() {
            return Math.ceil(this.filteredData.length / this.itemsPerPage);
        },
        paginatedData() {
            const start = (this.currentPage - 1) * this.itemsPerPage;
            const end = start + this.itemsPerPage;
            return this.filteredData.slice(start, end);
        },
        visiblePages() {
            const pages = [];
            const startPage = Math.max(1, this.currentPage - 5);
            const endPage = Math.min(this.totalPages, this.currentPage + 5);

            for (let i = startPage; i <= endPage; i++) {
                pages.push(i);
            }

            return pages;
        },
        // 新增：預覽欄位陣列
        previewFields() {
            if (!this.editableDynamicFieldsString.trim()) {
                return [];
            }
            const separator = this.editableDynamicFieldsString.includes('|') ? '|' : ',';
            return this.editableDynamicFieldsString.split(separator).map(field => field.trim()).filter(field => field);
        }
    },
    methods: {
        // 新增：計算 Levenshtein 相似度 (回傳百分比字串)
        calculateSimilarity(a, b) {
            a = (a ?? '').toString();
            b = (b ?? '').toString();
            if (!a && !b) return 100;
            const lenA = a.length;
            const lenB = b.length;
            const dp = Array(lenA + 1).fill(null).map(() => Array(lenB + 1).fill(0));
            for (let i = 0; i <= lenA; i++) dp[i][0] = i;
            for (let j = 0; j <= lenB; j++) dp[0][j] = j;
            for (let i = 1; i <= lenA; i++) {
                for (let j = 1; j <= lenB; j++) {
                    const cost = a[i - 1] === b[j - 1] ? 0 : 1;
                    dp[i][j] = Math.min(
                        dp[i - 1][j] + 1,
                        dp[i][j - 1] + 1,
                        dp[i - 1][j - 1] + cost
                    );
                }
            }
            const distance = dp[lenA][lenB];
            const maxLen = Math.max(lenA, lenB);
            const similarity = maxLen === 0 ? 1 : (maxLen - distance) / maxLen;
            return parseFloat((similarity * 100).toFixed(2));
        },
        handleFile(e) {
            const file = e.target.files[0];
            const reader = new FileReader();
            this.loading = true;

            reader.onload = (event) => {
                const data = new Uint8Array(event.target.result);
                this.workbook = XLSX.read(data, { type: 'array' });

                const sheetName = 'Summary'; // 根據你的實際sheet名稱
                const worksheet = this.workbook.Sheets[sheetName];

                // Find the last non-empty row in column A starting from A5
                let lastRow = 5;
                while (worksheet[`A${lastRow}`] !== undefined) {
                    lastRow++;
                }

                const errorFields = [];
                for (let i = 5; i < lastRow; i++) {
                    const errorField = worksheet[`A${i}`] ? worksheet[`A${i}`].v : null;
                    const errorCount = worksheet[`B${i}`] ? worksheet[`B${i}`].v : null;
                    if (errorField !== null && errorCount !== null) {
                        errorFields.push({ field: errorField, count: errorCount });
                    }
                }

                this.errorFields = errorFields;
                this.selectedError = errorFields.length > 0 ? errorFields[0].field : '';
                this.errorCount = errorFields.length > 0 ? errorFields[0].count : 0;
                this.loading = false;
            };

            reader.readAsArrayBuffer(file);
        },
        showData(selectedError) {
            this.loading = true;
            if (!selectedError) {
                selectedError = this.selectedError;
            }
            const dataTable = [];
            let sheetIndex = 0;

            while (true) {
                const sheetName = `Data_${String(sheetIndex).padStart(9, '0')}`;
                const worksheet = this.workbook.Sheets[sheetName];
                if (!worksheet) break;

                sheetIndex++;
                const matchingColumns = [];
                const headerRow = 1;
                let maxRow = parseInt(worksheet['!ref'].split(':')[1].replace(/[A-Z]/g, ''));

                for (let row = headerRow; row <= maxRow; row++) {
                    for (let col = 0; ; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: row - 1, c: col });
                        const cell = worksheet[cellAddress];
                        if (cell && cell.v === selectedError) {
                            matchingColumns.push([row, col]);
                        }
                        if (!cell) break;
                    }
                }

                matchingColumns.forEach(location => {
                    const row = location[0];
                    const col = location[1];
                    const logIdName = worksheet[XLSX.utils.encode_cell({ r: row - 1, c: 0 })]?.v || '';
                    if ("Log Id" === logIdName) {
                        const mainValue = worksheet[XLSX.utils.encode_cell({ r: row, c: col })]?.v || '';
                        const derivedValue = worksheet[XLSX.utils.encode_cell({ r: row + 1, c: col })]?.v || '';
                        const keyValue = worksheet[XLSX.utils.encode_cell({ r: row, c: 0 })]?.v || '';

                        // 拆分鍵值
                        const keyParts = keyValue.split('_');

                        // 動態建立資料物件
                        const rowData = {
                            row,
                            mainValue,
                            similarity: this.calculateSimilarity(mainValue, derivedValue),
                            derivedValue,
                            sheetName
                        };

                        // 動態設定欄位值
                        this.dynamicFields.forEach((field, index) => {
                            rowData[field.key] = keyParts[index] || '';
                        });

                        dataTable.push(rowData);
                    }
                });
            }

            this.dataTable = dataTable;
            this.loading = false;
        },
        showDetails(row, worksheetName) {
            const worksheet = this.workbook.Sheets[worksheetName];
            if (!worksheet) {
                console.error(`Worksheet ${worksheetName} not found`);
                return;
            }

            // 清空之前的資料
            this.modalHeaders = [];
            this.modalRows = [];

            // 假設表頭在第一行
            for (let col = 0; ; col++) {
                const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                const cell = worksheet[cellAddress];
                if (!cell) break;
                this.modalHeaders.push(cell.v);
            }

            for (let i = 0; i < 3; i++) { // 顯示 row, row+1, row+2 的內容
                const rowIndex = row.row + i;
                const rowCells = [];
                for (let col = 0; ; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex - 1, c: col });
                    const cell = worksheet[cellAddress];
                    if (!cell) break;
                    rowCells.push(cell.v);
                }
                this.modalRows.push({ index: rowIndex, cells: rowCells });
            }

            const modalElement = new bootstrap.Modal(document.getElementById('detailModal'));
            modalElement.show();
        },
        changePage(page) {
            if (page > 0 && page <= this.totalPages) {
                this.currentPage = page;
            }
        },
        // 新增方法：切換欄位編輯器顯示
        toggleFieldEditor() {
            this.showFieldEditor = !this.showFieldEditor;
            if (this.showFieldEditor) {
                // 顯示時同步當前的欄位字串到編輯器
                this.editableDynamicFieldsString = this.dynamicFieldsString;
            }
        },
        // 優化提示方法
        showToast(message, type = 'success') {
            // 建立 Toast 元素
            const toastHtml = `
                <div class="toast align-items-center text-white bg-${type === 'success' ? 'success' : 'danger'} border-0" role="alert">
                    <div class="d-flex">
                        <div class="toast-body">
                            <i class="bi bi-${type === 'success' ? 'check-circle' : 'exclamation-triangle'}"></i>
                            ${message}
                        </div>
                        <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast"></button>
                    </div>
                </div>
            `;
            
            // 加入到 Toast 容器
            const container = document.querySelector('.toast-container');
            container.insertAdjacentHTML('beforeend', toastHtml);
            
            // 顯示 Toast
            const toastElement = container.lastElementChild;
            const toast = new bootstrap.Toast(toastElement, { delay: 3000 });
            toast.show();
            
            // 自動移除
            toastElement.addEventListener('hidden.bs.toast', () => {
                toastElement.remove();
            });
        },
        // 新增方法：套用動態欄位變更
        applyDynamicFields() {
            if (!this.editableDynamicFieldsString.trim()) {
                this.showToast('欄位字串不能為空！', 'error');
                return;
            }
            
            // 驗證欄位格式，支援兩種分隔符號
            const separator = this.editableDynamicFieldsString.includes('|') ? '|' : ',';
            const fields = this.editableDynamicFieldsString.split(separator).map(field => field.trim()).filter(field => field);
            if (fields.length === 0) {
                this.showToast('至少需要一個有效的欄位名稱！', 'error');
                return;
            }
            
            // 套用變更
            this.dynamicFieldsString = this.editableDynamicFieldsString;
            
            // 更新過濾器
            this.updateFilters();
            
            // 如果有資料，重新處理
            if (this.dataTable.length > 0 && this.selectedError) {
                this.showData(this.selectedError);
            }
            
            this.showToast('欄位配置已更新！');
        },
        // 新增方法：重置欄位字串
        resetDynamicFields() {
            if (confirm('確定要重置為原始設定嗎？')) {
                this.editableDynamicFieldsString = this.originalDynamicFieldsString;
                this.dynamicFieldsString = this.originalDynamicFieldsString;
                
                // 更新過濾器
                this.updateFilters();
                
                // 如果有資料，重新處理
                if (this.dataTable.length > 0 && this.selectedError) {
                    this.showData(this.selectedError);
                }
                
                this.showToast('欄位配置已重置！');
            }
        },
        // 新增方法：更新過濾器物件以包含動態欄位
        updateFilters() {
            // 保留現有的過濾值
            const currentFilters = { ...this.filters };
            
            // 重建過濾器物件
            this.filters = {
                row: currentFilters.row || '',
                mainValue: currentFilters.mainValue || '',
                // 新增：保留相似度過濾
                similarity: currentFilters.similarity || '',
                derivedValue: currentFilters.derivedValue || '',
                sheetName: currentFilters.sheetName || ''
            };
            
            // 動態加入每個欄位的過濾器
            this.dynamicFields.forEach(field => {
                this.filters[field.key] = currentFilters[field.key] || '';
            });
        },
        // 切換欄位組
        changeFieldSet() {
            // 套用選擇的欄位組
            this.dynamicFieldsString = this.fieldSets[this.selectedFieldSetIndex].value;
            
            // 更新編輯區內容
            this.editableDynamicFieldsString = this.dynamicFieldsString;
            
            // 更新過濾器
            this.updateFilters();
            
            // 如果有資料，重新處理
            if (this.dataTable.length > 0 && this.selectedError) {
                this.showData(this.selectedError);
            }
            
            this.showToast(`已切換至「${this.fieldSets[this.selectedFieldSetIndex].name}」欄位組！`);
        },
        // 新增自訂欄位組
        addCustomFieldSet() {
            const name = prompt('請輸入欄位組名稱:');
            if (!name) return;
            
            const value = prompt('請輸入欄位定義 (使用 | 分隔):');
            if (!value) return;
            
            // 驗證格式
            const separator = value.includes('|') ? '|' : ',';
            const fields = value.split(separator).map(field => field.trim()).filter(field => field);
            if (fields.length === 0) {
                this.showToast('至少需要一個有效的欄位名稱！', 'error');
                return;
            }
            
            // 新增到欄位組清單
            this.fieldSets.push({ name, value });
            
            // 選擇新增的欄位組
            this.selectedFieldSetIndex = this.fieldSets.length - 1;
            
            // 套用新欄位組
            this.changeFieldSet();
            
            this.showToast(`已新增「${name}」欄位組！`);
        },
        // 新增：數字比對工具
        applyNumericFilter(target, expression) {
            const value = String(expression).trim();
            if (!value) return true;
            if (target === null || target === undefined || isNaN(target)) return false;
            const match = value.match(/^(<=|>=|<>|<|>|=)?\s*(\d+(?:\.\d+)?)\s*%?$/);
            if (!match) return false;
            const operator = match[1] || '=';
            const threshold = parseFloat(match[2]);
            switch (operator) {
                case '<': return target < threshold;
                case '>': return target > threshold;
                case '<=': return target <= threshold;
                case '>=': return target >= threshold;
                case '<>': return target !== threshold;
                case '=': return target === threshold;
                default: return false;
            }
        },
        // 新增：相似度格式化
        formatSimilarity(value) {
            if (value === null || value === undefined || isNaN(value)) return '';
            return `${Number(value).toFixed(2)}%`;
        },
        openUsageModal() {
            const modalElement = document.getElementById('usageModal');
            if (!modalElement) {
                return;
            }
            const modalInstance = bootstrap.Modal.getOrCreateInstance(modalElement);
            modalInstance.show();
        }
    },
    // 新增：初始化時備份原始字串
    mounted() {
        this.originalDynamicFieldsString = this.dynamicFieldsString;
        this.editableDynamicFieldsString = this.dynamicFieldsString;
        // 初始化過濾器
        this.updateFilters();
        // 初始化 Bootstrap tooltips
        this.$nextTick(() => {
            const tooltipTriggerList = Array.prototype.slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
            tooltipTriggerList.forEach(element => {
                bootstrap.Tooltip.getOrCreateInstance(element);
            });
        });
    },
    watch: {
        selectedError(newError) {
            this.currentPage = 1;
            this.showData(newError);
        },
        filters: {
            handler() {
                this.currentPage = 1;
            },
            deep: true
        }
    }
}).mount('#app');