var dateFormat = '1';
var expenseItems = new Array();
var lastCellAddress = 'A1';

function getVersion() {
    return chrome.runtime.getManifest().version;
}

function handleFile(e) {
    dateFormat = document.getElementById('date-format').value;

    var files = e.target.files;
    if (files.length != 1)
        return;

    var f = files[0];
    var reader = new FileReader();

    reader.onload = function (e) {
        try {
            var data = e.target.result;
            var workbook = XLSX.read(data, {
                type: 'binary'
            });

            var ws = workbook.Sheets['Template'];

            var rowIndex = 0;
            expenseItems = new Array();

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

                expenseItems.push(expenseItem);
                rowIndex++;
            }

            console.log(expenseItems);
            start();
        } catch (error) {
            console.log('Error:' + error);
            console.log('LastCellAddress: ' + lastCellAddress);
            alert(
                'Excel file parsing error\n\n' +
                'Cell : ' + lastCellAddress + '\n\n' +
                error
            );
        }
    };

    reader.readAsBinaryString(f);
}

function start() {
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
                <p>날짜타입 선택하고 로컬에 저장된 엑셀템플릿을 지정해주세요</p>
				<p>적상한 엑셀템플릿을 지정했지만 반응이 없는경우, 브라우저를 관리자로 실행 후 재실부탁드립니다.</p>
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
    displayReleaseNotes()
});
