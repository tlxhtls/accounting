const assert = require('node:assert/strict');
const os = require('node:os');
const fs = require('node:fs');
const path = require('node:path');
const XLSX = require('xlsx');

global.XLSX = XLSX;
global.window = {};
global.crypto = global.crypto || require('node:crypto').webcrypto;

const { Router } = require('./router');
window.Router = Router;
window.CLASSIFICATION_RULES = [];

const { Processor } = require('./processor');

function fixtureExists(filename) {
    return fs.existsSync(path.join(__dirname, filename));
}

function readSource(filename) {
    let data = new Uint8Array(fs.readFileSync(path.join(__dirname, filename)));
    data = Processor.preprocessHtmlIfNeeded(data);
    const workbook = XLSX.read(data, { type: 'array' });
    const sourceInfo = Router.identifySource(workbook);
    assert.ok(sourceInfo.def, `${filename} should be identified`);
    return sourceInfo;
}

function normalizeFixture(filename) {
    const sourceInfo = readSource(filename);
    return Processor.normalizeData(filename, sourceInfo.jsonData, sourceInfo);
}

{
    if (fixtureExists('롯데카드지출.xls')) {
        const sourceInfo = readSource('롯데카드지출.xls');
        assert.equal(sourceInfo.def.type, 'lotte_card');
        assert.equal(sourceInfo.sheetName, 'Sheet2');

        const rows = Processor.classifyData(Processor.normalizeData('롯데카드지출.xls', sourceInfo.jsonData, sourceInfo));
        assert.equal(rows.length, 16);
        assert.equal(rows[0].date, '2026-04-29');
        assert.equal(rows[0].time, '18:47');
        assert.equal(rows[0].raw_description, '의약품결제_(주)앤플러스팜');
        assert.equal(rows[0].amount, 9700000);
        assert.equal(rows[0].col_card, 9700000);
        assert.equal(rows[0].col_card_detail, '롯데카드');
        assert.equal(rows.some(row => row.raw_description === '네이버페이' && row.amount === 160800), false);
        assert.equal(rows.some(row => row.raw_description === '네이버페이' && row.amount === 128000), false);
    }
}

{
    if (fixtureExists('씨티계좌지출.xls')) {
        const rows = normalizeFixture('씨티계좌지출.xls');
        assert.equal(rows[0].date, '2026-04-30');
        assert.equal(rows[0].time, '18:50:26');
        assert.match(rows[0].date, /^\d{4}-\d{2}-\d{2}$/);
        assert.equal(rows.some(row => row.raw_description === '신한카드캐시백'), false);
        assert.equal(rows.some(row => row.raw_description === '신한카드환불'), false);
        assert.equal(rows.some(row => row.amount <= 0), false);
    }
}

{
    if (fixtureExists('씨티카드지출.xls')) {
        const rows = normalizeFixture('씨티카드지출.xls');
        assert.equal(rows[0].date, '2026-04-30');
        assert.equal(rows[0].raw_description, '쿠팡(로켓와우클럽');
        assert.equal(rows[0].amount, 7890);
        assert.match(rows[0].date, /^\d{4}-\d{2}-\d{2}$/);
        assert.equal(rows.some(row => /[월화수목금토일]요일?/.test(row.date)), false);
    }
}

{
    assert.ok(fixtureExists('신한카드지출.xls'), '신한카드지출.xls fixture is required');
    const rows = normalizeFixture('신한카드지출.xls');

    assert.equal(
        rows.some(row => row.amount < 0),
        false,
        '신한 취소/부분취소 음수 행 should not reduce exported expense totals'
    );
    assert.equal(rows.some(row => row.purchase_status === '부분취소'), false);

    let data = new Uint8Array(fs.readFileSync(path.join(__dirname, '신한카드지출.xls')));
    data = Processor.preprocessHtmlIfNeeded(data);
    const workbook = XLSX.read(data, { type: 'array' });
    const sourceInfo = Router.identifySource(workbook);
    assert.equal(sourceInfo.def.type, 'shinhan_card');
    const misnamedRows = Processor.normalizeData('wrong-upload-name.xls', sourceInfo.jsonData, sourceInfo);
    assert.equal(misnamedRows[0].raw_filename, 'wrong-upload-name.xls');
    assert.equal(misnamedRows.some(row => row.raw_source !== '신한카드'), false);
}

{
    const sourceInfo = {
        def: {
            type: 'test_card',
            name: '테스트카드',
            mapping: {
                date: '일자',
                raw_description: '가맹점',
                amount: '금액',
                cancel_status: '상태',
                cancel_amount: '취소금액',
                purchase_status: '매입구분'
            }
        },
        headerRow: ['일자', '가맹점', '금액', '상태', '취소금액', '매입구분'],
        headerIndex: 0,
        jsonData: [
            ['일자', '가맹점', '금액', '상태', '취소금액', '매입구분'],
            ['2026-04-01', '정상가맹점', '10,000', '정상', '0', '결제확정'],
            ['2026-04-02', '취소가맹점', '10,000', '취소', '0', '승인취소'],
            ['2026-04-03', '환불가맹점', '-5,000', '취소', '0', '승인취소'],
            ['2026-04-04', '전액취소금액가맹점', '10,000', '정상', '-10,000', '결제확정'],
            ['2026-04-05', '부분취소금액가맹점', '10,000', '정상', '-3,000', '결제확정'],
            ['2026-04-06', '부분취소행가맹점', '-2,000', '취소', '0', '부분취소']
        ]
    };

    const rows = Processor.normalizeData('synthetic-card.xls', sourceInfo.jsonData, sourceInfo);
    assert.deepEqual(rows.map(row => row.raw_description), ['정상가맹점', '부분취소금액가맹점']);
    assert.equal(rows[1].amount, 10000);
    assert.equal(rows[1].transaction_type, 'partial_refund');
    assert.equal(rows[1].refund_amount, 3000);
    assert.equal(rows[1].needs_review, true);
}

{
    assert.equal(Processor.formatDate(46142.999), '2026-04-30');
}

{
    const tmpFile = path.join(os.tmpdir(), `accounting-export-${Date.now()}.xlsx`);
    Processor.exportToExcel([{
        display_date: '2026-04-30',
        item: '',
        raw_description: '씨티카드 날짜 서식 검증',
        amount: 7890,
        col_transfer: '',
        col_account: '',
        col_cash: '',
        col_card: 7890,
        col_card_detail: '씨티카드',
        category_main: '',
        category_mso: ''
    }], tmpFile);

    const workbook = XLSX.readFile(tmpFile, { cellStyles: true });
    const cell = workbook.Sheets['지출내역_통합'].A2;
    assert.equal(cell.t, 's');
    assert.equal(cell.v, '2026-04-30');
    assert.equal(cell.z, '@');
    fs.unlinkSync(tmpFile);
}

console.log('parser tests passed');
