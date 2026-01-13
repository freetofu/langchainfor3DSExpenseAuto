
// ===== Expense Util - robust PeopleSoft action helper (injected) =====
(function () {
  async function sleep(ms){ return new Promise(r=>setTimeout(r,ms)); }
  function log(){ try { console.debug('[ExpenseUtil]', ...arguments); } catch {} }

  async function getTargetFrame(timeoutMs = 10000) {
    const t0 = Date.now();
    while (Date.now() - t0 < timeoutMs) {
      const cand =
        document.querySelector('iframe#ptifrmtgtframe') ||
        document.querySelector('iframe#TargetContent') ||
        document.querySelector('iframe[name="TargetContent"]') ||
        [...document.querySelectorAll('iframe')].find(i => {
          try { void i.contentWindow.document; return /\/ps[cp]\//.test(i.src||''); } catch { return false; }
        });
      if (cand?.contentWindow?.document?.readyState === 'complete') return cand.contentWindow;
      await sleep(100);
    }
    throw new Error('target iframe not ready');
  }

  function findAddButton(doc){
    return doc.querySelector("a[id*='EXPENSE_LINES$new'],button[id*='EXPENSE_LINES$new'],a[name*='EXPENSE_LINES$new']") ||
           doc.querySelector("a[id*='$newm'],button[id*='$newm'],a[name*='$newm']") ||
           doc.querySelector("a[id*='$add'], button[id*='$add'], a[name*='$add']");
  }

  function parseHrefForAction(el){
    const href = el?.getAttribute?.('href') || '';
    const m = href.match(/submitAction_win(\d+)\s*\(\s*(?:document\.)?win\1\s*,\s*'([^']+)'/i);
    if (m) return { winKey: `win${m[1]}`, token: m[2], submitName: `submitAction_win${m[1]}` };
    return null;
  }

  function detectSubmit(w, winKey){
    const directName = `submitAction_${winKey}`;
    const fn = (typeof w[directName] === 'function') ? w[directName]
            : (typeof w.submitAction   === 'function') ? w.submitAction
            : null;
    const winObj = w[winKey];
    return (fn && winObj) ? { fn: fn.bind(w), winObj } : null;
  }

  async function addExpenseLineInternal(tokenGuess){
    const w = await getTargetFrame();
    const addEl = findAddButton(w.document);
    if (!addEl) throw new Error('Add-line element not found');

    const parsed = parseHrefForAction(addEl);
    const winKey = parsed?.winKey || (Object.keys(w).find(k => /^win\d+$/.test(k)) || null);
    const token  = parsed?.token  || tokenGuess || 'EXPENSE_LINES$newm0$$0';

    if (winKey) {
      const sub = detectSubmit(w, winKey);
      if (sub) {
        try { sub.fn(sub.winObj, token); return true; } catch(e){}
      }
    }
    addEl.click(); // fallback
    return true;
  }

  window.ExpenseUtil = Object.assign(window.ExpenseUtil || {}, {
    addExpenseLine: (tokenGuess) => addExpenseLineInternal(tokenGuess),
    clickAddOnly: async () => {
      const w = await getTargetFrame();
      const addEl = findAddButton(w.document);
      if (!addEl) throw new Error('Add-line element not found');
      addEl.click(); return true;
    }
  });
})();
// ===== end injected helper =====

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function waitProcessing(window) {
    var count = 0;
    while (count < 180) {
        if (typeof window.isLoaderInProcess != 'function')
            return;
        if (!window.isLoaderInProcess())
            return;
        await sleep(1000);
        count++;
    }

    throw 'Time over!!!\nCheck internet connection.';
}

async function getElement(window, id, noWaiting) {
    if (noWaiting)
        return window.document.getElementById(id);
    else {
        var count = 0;
        while (count < 180) {
            try {
                element = window.document.getElementById(id);
                if (element)
                    return element;
            } catch {}
            await sleep(1000);
            count++;
        }

        throw 'Cannot find element: ' + id + '\nTime over!!!\nCheck internet connection.';
    }
}

async function setElementValue(window, element, value, isCombo) {
    if (isCombo) {
        for (var i = 0; i < element.options.length; i++) {
            if (element.options[i].text == value) {
                element.options[i].selected = true;
                break;
            }
        }
    } else
        element.value = value;
    element.onchange();
    await waitProcessing(window);
}

async function waitForIframeLoad(iframeWindow) {
    return new Promise((resolve, reject) => {
        if (iframeWindow && iframeWindow.document.readyState === 'complete') {
            resolve();
        } else {
            iframeWindow.addEventListener('load', resolve);
            setTimeout(() => reject('Iframe load timeout'), 5000); // 5초 대기
        }
    });
}

async function startExpenseInput(dateFormat, expenseItems) {
    try {
        var iframeWindow = document.getElementById('ptifrmtgtframe').contentWindow;

	
        if (expenseItems.length > 50) {
            alert('입력 항목은 50개까지만 자동 입력됩니다.\n 50개 초과 항목은 수동입력하시기 바랍니다.');
        }

        iframeWindow.ICAddCount.value = expenseItems.length - 1 > 50 ? 50 : expenseItems.length - 1;
        
		var dynamicExpenseLine = 'EXPENSE_LINES$newm$0$$0';
		console.log('Dynamic expense line:', dynamicExpenseLine);
		await window.ExpenseUtil.addExpenseLine(dynamicExpenseLine);
        await waitProcessing(iframeWindow);

        var rowIndex = 0;
        for (rowIndex = 0; rowIndex < expenseItems.length; ++rowIndex) {
            if (rowIndex == 50)
                break;

            expenseItem = expenseItems[rowIndex];

            element = await getElement(iframeWindow, 'TRANS_DATE$' + rowIndex);
            var expenseDate = new Date(expenseItem.expenseDate);
            if (dateFormat == 1)
                await setElementValue(iframeWindow, element, expenseDate.getFullYear() + "-" + (expenseDate.getMonth() + 1) + "-" + expenseDate.getDate());
            else if (dateFormat == 2)
                await setElementValue(iframeWindow, element, (expenseDate.getMonth() + 1) + "/" + expenseDate.getDate() + "/" + expenseDate.getFullYear());
            else if (dateFormat == 3)
                await setElementValue(iframeWindow, element, expenseDate.getDate() + "/" + (expenseDate.getMonth() + 1) + "/" + expenseDate.getFullYear());
            else
                await setElementValue(iframeWindow, element, expenseDate.getFullYear() + "/" + (expenseDate.getMonth() + 1) + "/" + expenseDate.getDate());

            if (rowIndex == 0)
                await sleep(3000);

            element = await getElement(iframeWindow, 'DSX_EX_EE_WRK_DESCR100A$' + rowIndex, true);
            if (element != null) {
                element.click();
                await sleep(5000);
                await waitProcessing(iframeWindow);

                var popupWindow = document.getElementById('ptModFrame_' + window.modId).contentWindow;
                await sleep(500);

                var elements = popupWindow.document.getElementsByTagName("a");
                var found;
                for (var i = 0; i < elements.length; i++) {
                    if (elements[i].textContent == expenseItem.expenseType) {
                        found = elements[i];
                        break;
                    }
                }

                found.click();

                await sleep(500);
                await waitProcessing(popupWindow);
                await waitProcessing(iframeWindow);
            }

            element = await getElement(iframeWindow, 'DESCR$' + rowIndex);
            await setElementValue(iframeWindow, element, expenseItem.description);

            element = await getElement(iframeWindow, 'EX_SHEET_LINE_TXN_CURRENCY_CD$' + rowIndex);
            await setElementValue(iframeWindow, element, expenseItem.currency);

            element = await getElement(iframeWindow, 'PAYMENT_TYPE$' + rowIndex);
            await setElementValue(iframeWindow, element, expenseItem.paymentType, true);

            element = await getElement(iframeWindow, 'EX_SHEET_LINE_BILL_CODE_EX$' + rowIndex);
            await setElementValue(iframeWindow, element, expenseItem.billingType, true);

            element = await getElement(iframeWindow, 'TRANS_AMT1$' + rowIndex);
            await setElementValue(iframeWindow, element, expenseItem.amount);

            element = await getElement(iframeWindow, 'EX_LOCATION_VW6_DESCR254$' + rowIndex, true);
            if (element != null && expenseItem.location != null)
                await setElementValue(iframeWindow, element, expenseItem.location);

            element = await getElement(iframeWindow, 'KILOMETERS$' + rowIndex, true);
            if (element != null)
                await setElementValue(iframeWindow, element, expenseItem.mileage);

            element = await getElement(iframeWindow, 'NBR_NIGHTS$' + rowIndex, true);
            if (element != null)
                await setElementValue(iframeWindow, element, expenseItem.numberOfNight);

            element = await getElement(iframeWindow, 'EX_LINE_WRK_PB_ATTENDEES$' + rowIndex, true);
            if (element != null && expenseItem.attendees.length > 0) {
                element.click();
                await sleep(500);
                await waitProcessing(iframeWindow);

                var popupWindow = document.getElementById('ptModFrame_' + window.modId).contentWindow;

                for (attendeeIndex = 0; attendeeIndex < expenseItem.attendees.length; ++attendeeIndex) {
                    element = await getElement(popupWindow, 'EX_SHEET_ATT$new$' + attendeeIndex + '$$0');
                    element.click();
                    await sleep(500);

                    element = await getElement(popupWindow, 'EX_SHEET_ATT_NAME$' + (attendeeIndex + 1));
                    await setElementValue(popupWindow, element, expenseItem.attendees[attendeeIndex].name);

                    element = await getElement(popupWindow, 'EX_SHEET_ATT_ATTENDEE_COMPANY$' + (attendeeIndex + 1));
                    await setElementValue(popupWindow, element, expenseItem.attendees[attendeeIndex].company);

                    element = await getElement(popupWindow, 'EX_SHEET_ATT_TITLE$' + (attendeeIndex + 1));
                    await setElementValue(popupWindow, element, expenseItem.attendees[attendeeIndex].title);
                }

                element = document.getElementById('ptModCloseLnk_' + window.modId);
                element.click();
                await sleep(500);
                await waitProcessing(popupWindow);
                await waitProcessing(iframeWindow);
            }
        }

        await sleep(1000);
        await waitProcessing(iframeWindow);
        alert('Done');
    } catch (error) {
        console.log(error);
        alert('Error !!! \n' + error)
    }

    window.postMessage({
        action: 'page_done'
    }, '*');
}

window.addEventListener('content_start_input', function (event) {
    console.log(event.detail.dateFormat);
    console.log(event.detail.data);
    startExpenseInput(event.detail.dateFormat, event.detail.data);
}, false);