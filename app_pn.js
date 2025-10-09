const { createApp } = Vue;

createApp({
    data() {
        return {
            workbook: null,
            errorFields: [],
            selectedError: '',
            errorCount: 0,
            ctrlCodeCounts: { '00': 0, '01': 0, '02': 0, '03': 0 },
            dataTable: [],
            filters: {
                row: '',
                mainValue: '',
                derivedValue: '',

                BILL_YEAR_MONTH  : '',
                BILL_CYCLE       : '',
                BILL_OFFICE_CODE : '',
                BILL_EQUIP_NO    : '',
                BILL_TYPE        : '',
                RECEIPT_NO       : '',
                sheetName: ''
            },
            loading: false,
            currentPage: 1,
            itemsPerPage: 100,
            modalHeaders: [],
            modalRows: []
        };
    },
    computed: {
        filteredData() {
            return this.dataTable.filter(row => {
                return Object.keys(this.filters).every(key => {
                    return String(row[key]).toUpperCase().includes(this.filters[key].toUpperCase());
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
        }
    },
    methods: {
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
        calculateCtrlCodeCounts(selectedError) {
            let sheetIndex = 0;
            let ctrlCodeCounts = { '00': 0, '01': 0, '02': 0, '03': 0 };

            while (true) {
                const sheetName = `Data_${String(sheetIndex).padStart(9, '0')}`;
                const worksheet = this.workbook.Sheets[sheetName];
                if (!worksheet) break;

                sheetIndex++;
                const matchingColumns = [];
                const headerRow = 1; // 假設表頭在第一行
                let maxRow = parseInt(worksheet['!ref'].split(':')[1].replace(/[A-Z]/g, '')); // 找到最大行數

                for (let row = headerRow; row <= maxRow; row++) {
                    for (let col = 0; ; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: row - 1, c: col });
                        const cell = worksheet[cellAddress];
                        if (cell && cell.v === selectedError) {
                            matchingColumns.push([row, col]);
                        }
                        if (!cell) break; // 到達最後一欄
                    }
                }

                matchingColumns.forEach(location => {
                    const row = location[0];
                    const col = location[1];
                    const keyValue = worksheet[XLSX.utils.encode_cell({ r: row, c: 0 })]?.v || '';
                    // const keyParts = keyValue.split('_');
                    // const CTRL_CODE = keyParts[11] || '';

                    // if (['00', '01', '02', '03'].includes(CTRL_CODE)) {
                    //     ctrlCodeCounts[CTRL_CODE]++;
                    // }
                });
            }

            // this.ctrlCodeCounts = ctrlCodeCounts;
        },
        showData(selectedError) {
            this.loading = true;
            if (!selectedError) {
                selectedError = this.selectedError;
            }
            const dataTable = [];
            let ctrlCodeCounts = { '00': 0, '01': 0, '02': 0, '03': 0 };
            let sheetIndex = 0;

            while (true) {
                const sheetName = `Data_${String(sheetIndex).padStart(9, '0')}`;
                const worksheet = this.workbook.Sheets[sheetName];
                if (!worksheet) break;

                sheetIndex++;
                const matchingColumns = [];
                const headerRow = 1; // 假設表頭在第一行
                let maxRow = parseInt(worksheet['!ref'].split(':')[1].replace(/[A-Z]/g, '')); // 找到最大行數

                for (let row = headerRow; row <= maxRow; row++) {
                    for (let col = 0; ; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: row - 1, c: col });
                        const cell = worksheet[cellAddress];
                        if (cell && cell.v === selectedError) {
                            matchingColumns.push([row, col]);
                        }
                        if (!cell) break; // 到達最後一欄
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

                        const [
                            BILL_YEAR_MONTH,BILL_CYCLE,BILL_OFFICE_CODE,BILL_EQUIP_NO,BILL_TYPE,RECEIPT_NO
                        ] = keyParts;

                        // if (['00', '01', '02', '03'].includes(CTRL_CODE)) {
                        //     ctrlCodeCounts[CTRL_CODE]++;
                        // }

                        dataTable.push({
                            row,
                            mainValue,
                            derivedValue,

                            BILL_YEAR_MONTH ,
                            BILL_CYCLE      ,
                            BILL_OFFICE_CODE,
                            BILL_EQUIP_NO   ,
                            BILL_TYPE       ,
                            RECEIPT_NO      ,
                            sheetName // 新增 sheetName
                        });
                    }
                });
            }

            this.ctrlCodeCounts = ctrlCodeCounts;
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
        }
    },
    watch: {
        selectedError(newError) {
            this.showData(newError);
        }
    }
}).mount('#app');