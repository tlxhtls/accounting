/**
 * processor.js
 * 
 * Handles file reading, source identification, data normalization, and classification.
 * Separates logic from the UI (index.html) and Configuration (rules.js).
 */

const Processor = {

    /**
     * Main entry point to process a single file.
     * @param {File} file - The file object from input
     * @returns {Promise<Array>} - Array of processed row objects
     */
    async processFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    let data = new Uint8Array(e.target.result);

                    // Pre-process HTML-based .xls files (fix malformed <td> tags)
                    data = this.preprocessHtmlIfNeeded(data, file.name);

                    const workbook = XLSX.read(data, { type: 'array' });

                    // 1. Identify Source (using Router)
                    const sourceInfo = window.Router.identifySource(file.name, workbook);

                    if (!sourceInfo.def) {
                        console.warn(`[${file.name}] Source identification failed. Header probe: ${sourceInfo.debugHeader}`);
                        resolve([{ error: true, filename: file.name, msg: `식별 실패. 헤더: [${sourceInfo.debugHeader}]` }]);
                        return;
                    }

                    console.log(`[${file.name}] Identified as: ${sourceInfo.def.name} in sheet: ${sourceInfo.sheetName}`);

                    // 2. Extract & Normalize
                    let normalizedData;

                    // Special handling for salary summary files
                    if (sourceInfo.def.type === 'salary_summary') {
                        normalizedData = this.normalizeSalarySummary(file.name, sourceInfo.jsonData, sourceInfo);
                    } else {
                        normalizedData = this.normalizeData(file.name, sourceInfo.jsonData, sourceInfo);
                    }

                    // 3. Classify
                    const classifiedData = this.classifyData(normalizedData);

                    resolve(classifiedData);

                } catch (err) {
                    reject(err);
                }
            };

            reader.onerror = (err) => reject(err);
            reader.readAsArrayBuffer(file);
        });
    },

    /**
     * Pre-processes HTML-based Excel files to fix malformed <td> tags.
     * Some banks export .xls files that are actually HTML with missing </td> tags.
     */
    preprocessHtmlIfNeeded(data, filename) {
        const textDecoder = new TextDecoder('utf-8');
        const text = textDecoder.decode(data.slice(0, 500));

        if (!text.includes('<html') && !text.includes('<table')) {
            return data; // Not HTML, return as-is
        }

        // Decode full file
        let fullText = textDecoder.decode(data);

        // Multi-pass fix for unclosed <td> tags
        // Pattern: Find <td...> followed by another <td (without </td> in between)
        let prevLength = 0;
        while (fullText.length !== prevLength) {
            prevLength = fullText.length;
            // This regex finds: <td attrs>content<td and replaces with <td attrs>content</td><td
            fullText = fullText.replace(/(<td[^>]*>)([^<]*?)(<td)/gi, '$1$2</td>$3');
        }

        // Fix last td in each row: <td attrs>content</tr> -> <td attrs>content</td></tr>
        fullText = fullText.replace(/(<td[^>]*>)([^<]*?)(<\/tr>)/gi, '$1$2</td>$3');

        // Encode back to Uint8Array
        const encoder = new TextEncoder();
        return encoder.encode(fullText);
    },

    /**
     * Normalizes raw Excel data into the standard internal schema.
     */
    normalizeData(filename, jsonData, sourceInfo) {
        const { def, headerRow, headerIndex } = sourceInfo;
        const map = def.mapping;
        const results = [];

        // Helper to find column index
        const getColIdx = (fieldName) => {
            const targetCols = fieldName ? (Array.isArray(fieldName) ? fieldName : [fieldName]) : [];
            if (!targetCols.length) return -1;

            for (const target of targetCols) {
                const cleanTarget = target.replace(/\s+/g, '');
                // 1. Exact match
                let idx = headerRow.indexOf(target);
                if (idx !== -1) return idx;

                // 2. Fuzzy match (ignore spaces)
                idx = headerRow.findIndex(h => h.replace(/\s+/g, '') === cleanTarget);
                if (idx !== -1) return idx;

                // 3. Substring match (careful)
                idx = headerRow.findIndex(h => h.includes(target));
                if (idx !== -1) return idx;
            }
            return -1;
        };

        const idxDate = getColIdx(map.date);
        const idxTime = getColIdx(map.time);
        const idxDesc = getColIdx(map.raw_description);
        const idxAmtMain = getColIdx(map.amount);
        const idxAmtAlt = getColIdx(map.amount_alt);

        for (let i = headerIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length === 0) continue;

            const valDate = row[idxDate];
            if (!valDate) continue; // Skip empty rows

            // Amount logic
            let valAmount = 0;
            if (idxAmtMain !== -1 && row[idxAmtMain]) valAmount = row[idxAmtMain];
            else if (idxAmtAlt !== -1 && row[idxAmtAlt]) valAmount = row[idxAmtAlt];

            // Extract KRW amount (handles foreign currency patterns)
            valAmount = this.extractKRW(valAmount);

            // Raw Description
            const valDesc = (idxDesc !== -1 && row[idxDesc]) ? String(row[idxDesc]).trim() : '';

            // Date Normalization
            const stdDate = this.formatDate(valDate);

            results.push({
                id: crypto.randomUUID(),
                date: stdDate,
                time: (idxTime !== -1 && row[idxTime]) ? row[idxTime] : '',
                amount: valAmount,
                raw_description: valDesc,     // 지출처 - always preserved
                item: '',                     // 항목 - filled by classification rules
                category_detail: '',          // (C) - 세부분류
                category_main: '',            // (L) - 중분류
                category_mso: '',             // (M) - MSO
                raw_source: def.name,
                raw_filename: filename
            });
        }
        return results;
    },

    /**
     * Normalizes salary summary data (급여총괄표).
     * Extracts totals from specific columns and creates accounting entries.
     */
    normalizeSalarySummary(filename, jsonData, sourceInfo) {
        const { def, headerRow, headerIndex } = sourceInfo;
        const map = def.mapping;
        const results = [];

        // Helper to find column index
        const getColIdx = (fieldName) => {
            if (!fieldName) return -1;
            const cleanTarget = fieldName.replace(/\s+/g, '');

            // 1. Exact match
            let idx = headerRow.indexOf(fieldName);
            if (idx !== -1) return idx;

            // 2. Fuzzy match (ignore spaces)
            idx = headerRow.findIndex(h => h.replace(/\s+/g, '') === cleanTarget);
            if (idx !== -1) return idx;

            // 3. Substring match
            idx = headerRow.findIndex(h => h.includes(fieldName));
            if (idx !== -1) return idx;

            return -1;
        };

        // Find columns for net payment and total deduction
        const idxNetPayment = getColIdx(map.net_payment);      // 차인지급액
        const idxTotalDeduction = getColIdx(map.total_deduction); // 공제합계
        const idxEmployeeName = getColIdx(map.employee_name);  // 성명

        console.log(`[Salary Summary] Column indices - Net Payment: ${idxNetPayment}, Total Deduction: ${idxTotalDeduction}, Name: ${idxEmployeeName}`);


        // Find the '합계' (total) row
        let totalRow = null;
        let totalRowIndex = -1;

        console.log(`[Salary Summary] Searching for '합계' row in data with ${jsonData.length} total rows`);

        for (let i = headerIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length === 0) continue;

            // Debug: log first 40 rows to help identify the structure (합계 is around row 30)
            if (i - headerIndex <= 40) {
                console.log(`[Salary Summary] Row ${i}:`, row.slice(0, 10)); // Show first 10 columns
            }

            // More flexible matching: check if '합계' appears in ANY column of this row
            const rowContainsTotal = row.some(cell => {
                if (!cell) return false;
                const cellValue = String(cell).trim();
                const cleanValue = cellValue.replace(/\s+/g, ''); // Remove all spaces
                return cleanValue.includes('합계') || cleanValue === '합계';
            });

            if (rowContainsTotal) {
                totalRow = row;
                totalRowIndex = i;
                console.log(`[Salary Summary] Found total row at index ${i}:`, row);
                break;
            }
        }


        if (!totalRow) {
            console.warn(`[${filename}] Could not find '합계' row in salary summary`);
            return results;
        }

        // Extract values
        const netPaymentValue = this.extractKRW(totalRow[idxNetPayment] || 0);
        const totalDeductionValue = this.extractKRW(totalRow[idxTotalDeduction] || 0);

        console.log(`[Salary Summary] Net Payment: ${netPaymentValue}, Total Deduction: ${totalDeductionValue}`);

        // Get current date for the entries
        const currentDate = new Date();
        const stdDate = this.formatDate(currentDate);

        // Create Entry 1: 차인지급액 합계 → 직원인건비
        results.push({
            id: crypto.randomUUID(),
            date: stdDate,
            time: '',
            amount: netPaymentValue,
            raw_description: '직원 인건비',
            item: '직원인건비',
            category_detail: '',
            category_main: '직원인건비',
            category_mso: 'MSO인정경비',
            raw_source: def.name,
            raw_filename: filename
        });

        // Create Entry 2: 공제합계 → 직원소득세
        results.push({
            id: crypto.randomUUID(),
            date: stdDate,
            time: '',
            amount: totalDeductionValue,
            raw_description: '직원소득세',
            item: '직원소득세',
            category_detail: '',
            category_main: '직원인건비',
            category_mso: 'MSO인정경비',
            raw_source: def.name,
            raw_filename: filename
        });

        // Create Entry 3: (차인지급액 + 공제합계) × 1/12 → 직원 퇴직연금
        const retirementPension = (netPaymentValue + totalDeductionValue) / 12;
        results.push({
            id: crypto.randomUUID(),
            date: stdDate,
            time: '',
            amount: retirementPension,
            raw_description: '직원 퇴직연금',
            item: '직원 퇴직연금',
            category_detail: '',
            category_main: '직원인건비',
            category_mso: 'MSO인정경비',
            raw_source: def.name,
            raw_filename: filename
        });

        console.log(`[Salary Summary] Created ${results.length} entries`);
        return results;
    },

    /**
     * Applies classification rules to the normalized data.
     * Also pre-calculates derived columns for UI/Export consistency.
     * 
     * Schema:
     * - raw_description: 지출처 (preserved from source file, never overwritten)
     * - item: 항목 (set by rules, describes what was purchased)
     * - category_main: 중분류 (category)
     */
    classifyData(data) {
        // Apply Classification Rules
        if (window.CLASSIFICATION_RULES) {
            for (const entry of data) {
                for (const rule of window.CLASSIFICATION_RULES) {
                    const keywords = rule.keywords;
                    const desc = entry.raw_description.toLowerCase();
                    const isMatch = keywords.some(k => desc.includes(k.toLowerCase()));

                    if (isMatch) {
                        if (rule.updates) {
                            // item: 항목 - what was purchased (e.g., "주차료", "약품계수용 앱 구독료")
                            if (rule.updates.item !== undefined) {
                                entry.item = rule.updates.item;
                            }
                            // For backward compatibility with description_out
                            if (rule.updates.description_out !== undefined && !entry.item) {
                                entry.item = rule.updates.description_out;
                            }
                            if (rule.updates.category_detail !== undefined) {
                                entry.category_detail = rule.updates.category_detail;
                            }
                            if (rule.updates.category_main !== undefined) {
                                entry.category_main = rule.updates.category_main;
                            }
                            if (rule.updates.category_mso !== undefined) {
                                entry.category_mso = rule.updates.category_mso;
                            }
                        }
                        break;
                    }
                }
            }
        }

        // Then Calculate Derived Columns (Card, Cash, Transfer etc)
        for (const entry of data) {
            const isCard = entry.raw_source.includes('카드');
            const isAccount = entry.raw_source.includes('계좌');

            entry.display_date = entry.date; // User requested Date only (YYYY-MM-DD)

            entry.col_transfer = '';
            entry.col_account = '';
            entry.col_cash = '';
            entry.col_card = '';
            entry.col_card_detail = '';

            if (isCard) {
                entry.col_card = entry.amount;
                entry.col_card_detail = entry.raw_source;
            } else if (isAccount) {
                entry.col_account = entry.raw_source;
                entry.col_transfer = entry.amount;
            } else {
                entry.col_cash = entry.amount;
            }
        }

        return data;
    },

    /**
     * Extracts Korean Won (KRW) amount from a value.
     * Handles foreign currency patterns like "￦36,411<br/>[USD]25.72"
     * @param {any} value - The raw amount value
     * @returns {number} - The extracted KRW amount
     */
    extractKRW(value) {
        // Already a number
        if (typeof value === 'number') return value;
        if (!value) return 0;

        const str = String(value);

        // 1. Look for ￦ symbol followed by amount
        const wonMatch = str.match(/￦\s*([\d,]+)/);
        if (wonMatch) {
            return parseFloat(wonMatch[1].replace(/,/g, '')) || 0;
        }

        // 2. Take content before <br/> or <br> tag (removes foreign currency part)
        const beforeBr = str.split(/<br\s*\/?>/i)[0];

        // 3. Remove currency symbols and extract number
        const cleaned = beforeBr.replace(/[^0-9.-]/g, '');
        return parseFloat(cleaned) || 0;
    },

    /**
     * Formats various date inputs into YYYY-MM-DD
     */
    formatDate(d) {
        if (!d) return '';

        // JavaScript Date object
        if (d instanceof Date) {
            const year = d.getFullYear();
            const month = String(d.getMonth() + 1).padStart(2, '0');
            const day = String(d.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        }

        // Excel Serial Date
        if (typeof d === 'number') {
            // Adjust for Excel epoch (approx)
            const dt = new Date(Math.round((d - 25569) * 86400 * 1000));
            // Return YYYY-MM-DD
            const year = dt.getFullYear();
            const month = String(dt.getMonth() + 1).padStart(2, '0');
            const day = String(dt.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        }

        // String Parsing
        let str = String(d).trim();

        // 1. Split by space or T to remove time part (e.g. "2025-10-21 15:30" -> "2025-10-21")
        str = str.split(/[\sT]+/)[0];

        // 2. Replace . / with -
        str = str.replace(/[\.\/]/g, '-');

        // 3. Handle YYYYMMDD (8 digits)
        if (/^\d{8}$/.test(str)) {
            return `${str.substring(0, 4)}-${str.substring(4, 6)}-${str.substring(6, 8)}`;
        }

        // 4. If it matches YYYY-MM-DD format, return it
        if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
            return str;
        }

        // Fallback: Use standard Date parsing if possible, else return cleaned string
        const parsed = Date.parse(str);
        if (!isNaN(parsed)) {
            const dt = new Date(parsed);
            const year = dt.getFullYear();
            const month = String(dt.getMonth() + 1).padStart(2, '0');
            const day = String(dt.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
        }

        return str;
    },

    /**
     * Exports data to Excel with the specific schema requested.
     * Columns: 발생일시 | 항목 | 지출처 | 총액 | 계좌이체 | 지출계좌 | 현금(인출) | 신용카드 | 신용카드상세 | 중분류 | 대분류
     */
    exportToExcel(data, filename = '지출내역_완료.xlsx') {
        if (!data || data.length === 0) return;

        const wsData = [
            ['발생일시', '항목', '지출처', '총액', '계좌이체', '지출계좌', '현금(인출)', '신용카드', '신용카드상세', '중분류', '대분류']
        ];

        data.forEach(row => {
            wsData.push([
                row.display_date,       // 발생일시
                row.item,               // 항목 (classified item name)
                row.raw_description,    // 지출처 (original source - preserved)
                row.amount,             // 총액
                row.col_transfer,       // 계좌이체
                row.col_account,        // 지출계좌
                row.col_cash,           // 현금(인출)
                row.col_card,           // 신용카드
                row.col_card_detail,    // 신용카드상세
                row.category_main,      // 중분류
                row.category_mso        // 대분류
            ]);
        });

        const ws = XLSX.utils.aoa_to_sheet(wsData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "지출내역_통합");

        // Adjust widths (11 cols)
        ws['!cols'] = [
            { wch: 18 }, { wch: 15 }, { wch: 25 }, { wch: 12 }, // Date, Item, Place, Total
            { wch: 12 }, { wch: 15 }, { wch: 12 },           // Transfer, Acct, Cash
            { wch: 12 }, { wch: 15 },                        // Card, CardDetail
            { wch: 15 }, { wch: 15 }                         // Mid, Big
        ];

        XLSX.writeFile(wb, filename);
    }
};

// Expose to window
window.Processor = Processor;
