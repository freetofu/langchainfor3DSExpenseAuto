var dateFormat = '1';
var expenseItems = new Array();
var lastCellAddress = 'A1';
var expenseMapping = null;
var startButton = null;
var previewEmpty = null;
var previewTable = null;
var previewBody = null;

var HOSP_REGEX = /hosp\./i;

function getVersion() {
    return chrome.runtime.getManifest().version;
}

function normalizeString(value) {
    return (value === undefined || value === null) ? '' : String(value).trim();
}

function parseNumber(value) {
    if (value === undefined || value === null || value === '')
        return null;
    if (typeof value === 'number')
        return value;
    var cleaned = String(value).replace(/,/g, '');
    var numberValue = Number(cleaned);
    return Number.isFinite(numberValue) ? numberValue : null;
}

function parseTimeToMinutes(value) {
    var text = normalizeString(value);
    if (!text)
        return null;
    var parts = text.split(':');
    if (parts.length < 2)
        return null;
    var hours = parseInt(parts[0], 10);
    var minutes = parseInt(parts[1], 10);
    if (Number.isNaN(hours) || Number.isNaN(minutes))
        return null;
    return hours * 60 + minutes;
}

function parseTimeToSeconds(value) {
    var text = normalizeString(value);
    if (!text)
        return null;
    var parts = text.split(':');
    if (parts.length < 2)
        return null;
    var hours = parseInt(parts[0], 10);
    var minutes = parseInt(parts[1], 10);
    var seconds = parts.length >= 3 ? parseInt(parts[2], 10) : 0;
    if (Number.isNaN(hours) || Number.isNaN(minutes) || Number.isNaN(seconds))
        return null;
    return hours * 3600 + minutes * 60 + seconds;
}

function formatDateValue(value) {
    if (value instanceof Date) {
        var year = value.getFullYear();
        var month = String(value.getMonth() + 1).padStart(2, '0');
        var day = String(value.getDate()).padStart(2, '0');
        return year + '-' + month + '-' + day;
    }

    var text = normalizeString(value);
    if (!text)
        return '';

    if (text.indexOf('.') !== -1)
        return text.replace(/\./g, '-');

    return text;
}

function isHospExpenseType(expenseType) {
    return HOSP_REGEX.test(normalizeString(expenseType));
}

function isTaxiText(value) {
    var text = normalizeString(value).toLowerCase();
    return text.indexOf('택시') !== -1 || text.indexOf('taxi') !== -1;
}

function matchesAny(value, candidates) {
    if (!candidates || candidates.length === 0)
        return true;
    var normalizedValue = normalizeString(value).toLowerCase();
    return candidates.some(function (candidate) {
        return normalizedValue.indexOf(normalizeString(candidate).toLowerCase()) !== -1;
    });
}

function buildRowData(row) {
    var dateValue = formatDateValue(row['이용일']);
    return {
        merchantName: normalizeString(row['가맹점명']),
        usagePlace: normalizeString(row['사용처']),
        categoryName: normalizeString(row['가맹점업종명']),
        categoryCode: normalizeString(row['가맹점업종코드']),
        userName: normalizeString(row['이용자명']),
        approvalTimeMinutes: parseTimeToMinutes(row['승인시간']),
        approvalTimeSeconds: parseTimeToSeconds(row['승인시간']),
        amount: parseNumber(row['이용금액']),
        dateKey: dateValue
    };
}

function isTaxiRow(rowData) {
    return isTaxiText(rowData.merchantName) || isTaxiText(rowData.usagePlace) || isTaxiText(rowData.categoryName);
}

function consolidateTaxiRows(rawRows) {
    if (!rawRows || rawRows.length === 0)
        return rawRows;

    var consolidated = [];
    rawRows.forEach(function (entry) {
        var rowData = entry.rowData;
        var lastEntry = consolidated[consolidated.length - 1];
        if (!lastEntry) {
            consolidated.push(entry);
            return;
        }

        var lastData = lastEntry.rowData;
        if (!isTaxiRow(rowData) || !isTaxiRow(lastData)) {
            consolidated.push(entry);
            return;
        }

        if (rowData.dateKey !== lastData.dateKey) {
            consolidated.push(entry);
            return;
        }

        if (rowData.userName !== lastData.userName) {
            consolidated.push(entry);
            return;
        }

        if (rowData.approvalTimeSeconds === null || lastData.approvalTimeSeconds === null) {
            consolidated.push(entry);
            return;
        }

        if (Math.abs(rowData.approvalTimeSeconds - lastData.approvalTimeSeconds) > 10) {
            consolidated.push(entry);
            return;
        }

        var mergedRow = Object.assign({}, lastEntry.row);
        var mergedAmount = (lastData.amount || 0) + (rowData.amount || 0);
        mergedRow['이용금액'] = mergedAmount;
        lastEntry.row = mergedRow;
        lastEntry.rowData.amount = mergedAmount;
    });

    return consolidated;
}

function matchesRule(rowData, rule) {
    var match = rule.match || {};
    if (!matchesAny(rowData.merchantName, match.merchantName))
        return false;
    if (!matchesAny(rowData.usagePlace, match.usagePlace))
        return false;
    if (!matchesAny(rowData.categoryName, match.categoryName))
        return false;
    if (!matchesAny(rowData.categoryCode, match.categoryCode))
        return false;
    if (!matchesAny(rowData.userName, match.userName))
        return false;

    if (rule.amountRange) {
        var amount = rowData.amount;
        if (amount === null)
            return false;
        if (rule.amountRange.min !== undefined && amount < rule.amountRange.min)
            return false;
        if (rule.amountRange.max !== undefined && amount > rule.amountRange.max)
            return false;
    }

    if (rule.timeRange) {
        var timeMinutes = rowData.approvalTimeMinutes;
        if (timeMinutes === null)
            return false;
        var startMinutes = parseTimeToMinutes(rule.timeRange.start);
        var endMinutes = parseTimeToMinutes(rule.timeRange.end);
        if (startMinutes !== null && timeMinutes < startMinutes)
            return false;
        if (endMinutes !== null && timeMinutes > endMinutes)
            return false;
    }

    return true;
}

function resolveExpenseType(rowData, mapping) {
    if (!mapping || !mapping.rules)
        return '';
    var matched = mapping.rules.find(function (rule) {
        return matchesRule(rowData, rule);
    });
    if (matched)
        return normalizeString(matched.expenseType);
    return normalizeString(mapping.defaults && mapping.defaults.expenseTypeFallback);
}

function buildExpenseItemFromRaw(row, rowData) {
    var defaults = (expenseMapping && expenseMapping.defaults) ? expenseMapping.defaults : {};
    var expenseType = resolveExpenseType(rowData, expenseMapping);
    var description = rowData.usagePlace || rowData.merchantName;

    return {
        expenseDate: formatDateValue(row['이용일']),
        expenseType: expenseType,
        description: description,
        amount: rowData.amount === null ? '' : rowData.amount,
        currency: normalizeString(row['통화코드']) || normalizeString(defaults.currency),
        paymentType: normalizeString(defaults.paymentType),
        billingType: normalizeString(defaults.billingType),
        attendees: new Array(),
        requiresAttendees: isHospExpenseType(expenseType)
    };
}

function renderPreview(items) {
    previewBody.innerHTML = '';
    if (!items || items.length === 0) {
        previewEmpty.style.display = 'block';
        previewTable.style.display = 'none';
        return;
    }

    previewEmpty.style.display = 'none';
    previewTable.style.display = 'table';

    items.slice(0, 50).forEach(function (item) {
        var row = document.createElement('tr');
        if (item.requiresAttendees)
            row.classList.add('hosp-required');

        var dateCell = document.createElement('td');
        dateCell.textContent = item.expenseDate || '';
        row.appendChild(dateCell);

        var usageCell = document.createElement('td');
        usageCell.textContent = item.description || '';
        row.appendChild(usageCell);

        var amountCell = document.createElement('td');
        amountCell.textContent = item.amount || '';
        row.appendChild(amountCell);

        var typeCell = document.createElement('td');
        typeCell.textContent = item.expenseType || '';
        row.appendChild(typeCell);

        previewBody.appendChild(row);
    });
}

function enableStartButton(enabled) {
    if (!startButton)
        return;
    startButton.disabled = !enabled;
}

function parseTemplateSheet(ws) {
    var rowIndex = 0;
    var parsedItems = new Array();

    if (ws['A1'].v != 'Expense Template v' + getVersion())
        throw 'Wrong expense template file version';

    while (true) {
        var row = 4 + rowIndex;
        var expenseDate = ws['B' + row];
        if (expenseDate === '' || expenseDate === undefined || expenseDate === null)
            break;

        var expenseItem = new Object();
        expenseItem.expenseDate = expenseDate.w;
        expenseItem.expenseType = ws[lastCellAddress = 'D' + row].v;
        if (ws[lastCellAddress = 'E' + row])
            expenseItem.description = ws['E' + row].v;
        else
            expenseItem.description = '';
        if (ws[lastCellAddress = 'F' + row])
            expenseItem.amount = ws['F' + row].v;
        else
            expenseItem.amount = '';
        expenseItem.currency = ws[lastCellAddress = 'G' + row].v;
        expenseItem.paymentType = ws[lastCellAddress = 'H' + row].v;
        expenseItem.billingType = ws[lastCellAddress = 'I' + row].v;

        cell = ws[lastCellAddress = 'J' + row];
        if (cell != undefined && cell.v != '')
            expenseItem.mileage = cell.v;

        cell = ws[lastCellAddress = 'K' + row];
        if (cell != undefined && cell.v != '')
            expenseItem.numberOfNight = cell.v;

        var attendees = new Array();
        while (true) {
            var cellAddress = {
                c: 11 + attendees.length * 3,
                r: row - 1
            };

            if (!ws[XLSX.utils.encode_cell(cellAddress)])
                break;

            var attendee = new Object();
            attendee.name = ws[lastCellAddress = XLSX.utils.encode_cell(cellAddress)].v;

            cellAddress.c += 1;
            attendee.company = ws[lastCellAddress = XLSX.utils.encode_cell(cellAddress)].v;

            cellAddress.c += 1;
            cell = ws[lastCellAddress = XLSX.utils.encode_cell(cellAddress)];
            if (cell != undefined)
                attendee.title = ws[lastCellAddress = XLSX.utils.encode_cell(cellAddress)].v;
            else
                attendee.title = '';

            attendees.push(attendee);
        }
        expenseItem.attendees = attendees;
        expenseItem.requiresAttendees = isHospExpenseType(expenseItem.expenseType);

        parsedItems.push(expenseItem);
        rowIndex++;
    }

    return parsedItems;
}

function parseRawSheet(ws) {
    var rows = XLSX.utils.sheet_to_json(ws, {
        defval: ''
    });
    var rawRows = rows.map(function (row) {
        return {
            row: row,
            rowData: buildRowData(row)
        };
    });
    var consolidated = consolidateTaxiRows(rawRows);

    return consolidated.map(function (entry) {
        return buildExpenseItemFromRaw(entry.row, entry.rowData);
    });
}

function loadExpenseMapping() {
    if (expenseMapping)
        return Promise.resolve(expenseMapping);

    return fetch(chrome.runtime.getURL('expenseMapping.json'))
        .then(function (response) {
            return response.json();
        })
        .then(function (mapping) {
            expenseMapping = mapping;
            return mapping;
        })
        .catch(function (error) {
            console.log('Failed to load mapping:', error);
            expenseMapping = {
                defaults: {
                    paymentType: '',
                    billingType: '',
                    currency: '',
                    expenseTypeFallback: ''
                },
                rules: []
            };
            return expenseMapping;
        });
}

function handleFile(e) {
    dateFormat = document.getElementById('date-format').value;

    var files = e.target.files;
    if (files.length != 1)
        return;

    var f = files[0];
    var reader = new FileReader();

    reader.onload = async function (e) {
        try {
            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });

            expenseItems = new Array();
            await loadExpenseMapping();

            var ws = workbook.Sheets['Template'] || workbook.Sheets[workbook.SheetNames[0]];

            if (!ws)
                throw 'No worksheet found';

            if (ws['A1'] && ws['A1'].v == 'Expense Template v' + getVersion())
                expenseItems = parseTemplateSheet(ws);
            else
                expenseItems = parseRawSheet(ws);

            console.log(expenseItems);
            renderPreview(expenseItems);
            enableStartButton(expenseItems.length > 0);
        } catch (error) {
            console.log('Error:' + error);
            console.log('LastCellAddress: ' + lastCellAddress);
            alert(
                'Excel file parsing error\n\n' +
                'Cell : ' + lastCellAddress + '\n\n' +
                error
            );
            renderPreview([]);
            enableStartButton(false);
        }
    };

    reader.readAsBinaryString(f);
}

function start() {
    if (!expenseItems || expenseItems.length === 0)
        return;

    chrome.tabs.query({
        active: true,
        currentWindow: true
    }, function (tabs) {
        chrome.tabs.sendMessage(tabs[0].id, {
            action: 'popup_start_input',
            dateFormat: dateFormat,
            data: expenseItems
        });

        window.close();
    });
}

// CORS 프록시를 사용하여 위키 내용을 가져오는 함수
function displayReleaseNotes() {

        $('#eu-release-notes').html(`
            <div style="
                padding: 20px;
                background-color: #ffe6e6;
                border: 1px solid #ff9999;
                border-radius: 5px;
                color: #cc0000;
                text-align: center;
            ">
                <h3></h3>
                <p>날짜 타입 선택 후 엑셀 템플릿 또는 원본 매출내역 파일을 지정해주세요.</p>
				<p>Hosp. 유형은 참석자 입력이 필요하니 미리보기를 확인해주세요.</p>
				<p>엑셀 선택 후 자동 입력 버튼이 활성화됩니다.</p>
            </div>
        `);
}

// 마크다운을 간단한 HTML로 변환하는 함수
function convertMarkdownToHtml(markdown) {
    if (!markdown) return '';
    
    return markdown
        // 제목 변환 (# -> h1, ## -> h2, etc.)
        .replace(/^### (.*$)/gim, '<h3>$1</h3>')
        .replace(/^## (.*$)/gim, '<h2>$1</h2>')
        .replace(/^# (.*$)/gim, '<h1>$1</h1>')
        // 볼드 텍스트 변환
        .replace(/\*\*(.*)\*\*/gim, '<strong>$1</strong>')
        // 링크 변환
        .replace(/\[([^\]]+)\]\(([^\)]+)\)/gim, '<a href="$2" target="_blank">$1</a>')
        // 줄바꿈 변환
        .replace(/\n/gim, '<br>');
}

document.addEventListener('DOMContentLoaded', function () {
    document.getElementById('eu-title').innerHTML = 'Expense Util v' + getVersion();
    document.getElementById('excel-file-selection').addEventListener('change', handleFile, false);
    startButton = document.getElementById('start-input');
    previewEmpty = document.getElementById('preview-empty');
    previewTable = document.getElementById('preview-table');
    previewBody = document.getElementById('preview-body');

    startButton.addEventListener('click', start, false);
    enableStartButton(false);
    displayReleaseNotes()
});
