/**
 * router.js
 * 
 * Responsibilities:
 * 1. Identify the source of the Excel file (Card company, Bank, etc.)
 * 2. Return the appropriate definition and data subset (sheet/rows) for processing.
 */

// Source Definitions (Signatures & Mappings)
const SOURCE_DEFINITIONS = [
    {
        type: 'kb_card',
        name: 'KB국민카드',
        signatures: ['이용일', '이용하신곳', '이용금액'],
        mapping: {
            'date': '이용일',
            'time': '이용시간',
            'raw_description': '이용하신곳',
            'amount': '이용금액(원)',
            'amount_alt': '이용금액'
        }
    },
    {
        type: 'bc_card',
        name: '비씨카드',
        signatures: ['승인일시', '가맹점명', '승인금액'],
        mapping: {
            'date': '승인일시',
            'raw_description': '가맹점명',
            'amount': '승인금액',
            'amount_alt': '거래금액'
        }
    },
    {
        type: 'shinhan_card',
        name: '신한카드',
        signatures: ['거래일', '가맹점명', '금액'],
        mapping: {
            'date': '거래일',
            'raw_description': '가맹점명',
            'amount': '금액'
        }
    },
    {
        type: 'samsung_card',
        name: '삼성카드',
        signatures: ['카드번호', '승인일자', '승인시각', '승인금액(원)'],
        mapping: {
            'date': '승인일자',
            'time': '승인시각',
            'raw_description': '가맹점명',
            'amount': '승인금액(원)'
        }
    },
    {
        type: 'citi_account',
        name: '씨티계좌',
        signatures: ['거래일시', '적요', '찾으신금액', '맡기신금액'],
        mapping: {
            'date': '거래일시',
            'time': '거래시간',
            'raw_description': '적요',
            'amount': '찾으신금액', // Withdrawal
            'amount_alt': '맡기신금액' // Deposit
        }
    },
    {
        type: 'citi_card',
        name: '씨티카드',
        signatures: ['이용일시', '이용카드', '가맹점명', '거래금액'],
        mapping: {
            'date': '이용일시',
            'raw_description': '가맹점명',
            'amount': '거래금액'
        }
    }
];

const Router = {
    definitions: SOURCE_DEFINITIONS,

    /**
     * Identifies the source definition by scanning all sheets in the workbook.
     * @param {Object} workbook - XLSX workbook object
     * @returns {Object} result - { def, headerRow, headerIndex, sheetName, jsonData } or error info
     */
    identifySource(filename, workbook) {
        // Scan each sheet
        for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            // Use header:1 to get array of arrays
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            if (!jsonData || jsonData.length === 0) continue;

            // Search first N rows for a matching header signature
            const MAX_SEARCH_ROWS = 500; // Deep search

            for (let i = 0; i < Math.min(MAX_SEARCH_ROWS, jsonData.length); i++) {
                const row = jsonData[i].map(c => c ? String(c).trim() : '');

                // Prepare clean row for matching
                const cleanRow = row.map(r => r.replace(/\s+/g, ''));

                for (const def of this.definitions) {
                    const signatures = def.signatures;

                    // Check if *all* signature keywords exist in this row
                    const isMatch = signatures.every(sig => {
                        const cleanSig = sig.replace(/\s+/g, '');
                        // Fuzzy match: check if cleanSig is substring of any cell in cleanRow
                        return cleanRow.some(cell => cell.includes(cleanSig));
                    });

                    if (isMatch) {
                        if (row.length < signatures.length) {
                            continue; // Skip malformed rows
                        }

                        return {
                            def,
                            headerRow: row,
                            headerIndex: i,
                            sheetName,
                            jsonData
                        };
                    }
                }
            }
        }

        // Return debug info if no match (from first sheet)
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const firstJson = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        const firstRow = (firstJson && firstJson[0]) ? firstJson[0].slice(0, 10).join(',') : 'EMPTY';
        return { def: null, debugHeader: firstRow };
    }
};

// Expose to window
window.Router = Router;
